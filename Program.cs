using System.Configuration;
using System.Data.SqlClient;
using System.Globalization;
using System.Xml.Linq;

using databaseAPI;
using GNAgeneraltools;
using GNAsurveytools;
using GNAspreadsheettools;
using OfficeOpenXml;
using Twilio.Rest.Api.V2010.Account;
using System.Diagnostics.Metrics;
using EASendMail;
using System.Data.SqlTypes;
//using Microsoft.Extensions.Configuration;

//
// 20240509 First version created
//

namespace dataFlowAlarm
{
    internal class Program
    {
        static void Main(string[] args)
        {
#pragma warning disable CS0164
#pragma warning disable CS8600
#pragma warning disable CS8602
#pragma warning disable CS8604
#pragma warning disable IDE0059

            gnaTools gnaT = new();
            dbAPI gnaDBAPI = new();
            spreadsheetAPI gnaSpreadsheetAPI = new();
            GNAsurveycalcs gnaSurvey = new();

            string strProgramStart = DateTime.Now.ToString("yyyyMMdd_HHmmss");

            //==== Console settings
            Console.OutputEncoding = System.Text.Encoding.Unicode;
            CultureInfo culture;

            //==== Set the EPPlus license
            ExcelPackage.LicenseContext = LicenseContext.Commercial;

            //==== System config variables

            string strDBconnection = ConfigurationManager.ConnectionStrings["DBconnectionString"].ConnectionString;
            string strProjectTitle = ConfigurationManager.AppSettings["ProjectTitle"];
            string strContractTitle = ConfigurationManager.AppSettings["ContractTitle"];
            string strExcelPath = ConfigurationManager.AppSettings["ExcelPath"];
            string strExcelFile = ConfigurationManager.AppSettings["ExcelFile"];
            string strReferenceWorksheet = ConfigurationManager.AppSettings["ReferenceWorksheet"];
            string strSurveyWorksheet = ConfigurationManager.AppSettings["SurveyWorksheet"];
            string strSMSTitle = ConfigurationManager.AppSettings["SMSTitle"];

            string strFreezeScreen = ConfigurationManager.AppSettings["freezeScreen"];

            string strRootFolder = ConfigurationManager.AppSettings["SystemStatusFolder"];
            string strAlarmsFolder = ConfigurationManager.AppSettings["SystemAlarmsFolder"];

            string strFirstDataRow = ConfigurationManager.AppSettings["FirstDataRow"];

            double dblAlarmWindowHrs = Convert.ToDouble(ConfigurationManager.AppSettings["AlarmWindowHrs"]);
            int iNoOfSuccessfulReadings = Convert.ToInt16(ConfigurationManager.AppSettings["NoOfSuccessfulReadings"]);

            // allocate the sms mobile numbers
            string[] smsMobile = new string[10];
            smsMobile[1] = ConfigurationManager.AppSettings["RecipientPhone1"];
            smsMobile[2] = ConfigurationManager.AppSettings["RecipientPhone2"];
            smsMobile[3] = ConfigurationManager.AppSettings["RecipientPhone3"];
            smsMobile[4] = ConfigurationManager.AppSettings["RecipientPhone4"];
            smsMobile[5] = ConfigurationManager.AppSettings["RecipientPhone5"];
            smsMobile[6] = ConfigurationManager.AppSettings["RecipientPhone6"];
            smsMobile[7] = ConfigurationManager.AppSettings["RecipientPhone7"];
            smsMobile[8] = ConfigurationManager.AppSettings["RecipientPhone8"];
            smsMobile[9] = ConfigurationManager.AppSettings["RecipientPhone9"];

            string strSendEmails = ConfigurationManager.AppSettings["SendEmails"];
            string strEmailLogin = ConfigurationManager.AppSettings["EmailLogin"];
            string strEmailPassword = ConfigurationManager.AppSettings["EmailPassword"];
            string strEmailFrom = ConfigurationManager.AppSettings["EmailFrom"];
            string strEmailRecipients = ConfigurationManager.AppSettings["EmailRecipients"];
            string smsAlarmState = "No alarms";
            string txtMessage = "No SMS";
            string strMobileList = "";
            string strEmailMessage = "";
            string strNow = "";

            //==== [Main program]===========================================================================================

            gnaT.WelcomeMessage("dataFlowAlarm 20240509");
            string strSoftwareLicenseTag = "DATALM";
            _ = gnaT.checkLicenseValidity(strSoftwareLicenseTag, strProjectTitle, strEmailLogin, strEmailPassword, strSendEmails);

            string strMasterWorkbookFullPath = strExcelPath + strExcelFile;

            Console.WriteLine("");
            Console.WriteLine("1. Check system environment");
            Console.WriteLine("     Project: " + strProjectTitle);
            Console.WriteLine("     Master workbook: " + strMasterWorkbookFullPath);

            gnaDBAPI.testDBconnection(strDBconnection);

            if (strFreezeScreen == "Yes")
            {
                gnaSpreadsheetAPI.checkWorksheetExists(strMasterWorkbookFullPath, strSurveyWorksheet);
            }
            else
            {
                Console.WriteLine("     Existance of workbook & worksheets is not checked");
            }

            


            for (int j = 1; j <= 9; j++)
            {
                if (smsMobile[j] != "None")
                {
                    strMobileList = strMobileList + smsMobile[j] + ",";
                }
            }
            strMobileList = strMobileList.Substring(0, strMobileList.Length - 1);
            Console.WriteLine("     Mobile recipients: "+ strMobileList);




            Console.WriteLine("2. Extract point names");
            string[] strPointNames = gnaSpreadsheetAPI.readPointNames(strMasterWorkbookFullPath, strSurveyWorksheet, strFirstDataRow);

            Console.WriteLine("3. Extract SensorID");
            string[,] strSensors = gnaDBAPI.getSensorIDfromDB(strDBconnection, strPointNames, strProjectTitle);

            //Console.WriteLine("4. Write SensorID to workbook");
            //gnaSpreadsheetAPI.writeSensorID(strMasterWorkbookFullPath, strSurveyWorksheet, strSensorID, strFirstDataRow);


            Console.WriteLine("4. Determine current alarm state");

            string strCurrentAlarmState = gnaDBAPI.getAlarmStatus(strDBconnection, strProjectTitle, strSensors, dblAlarmWindowHrs, iNoOfSuccessfulReadings);



            Console.WriteLine("5. Send notifications if needed");

            // Check whether the alarm state has changed & update the Alarm log
            string strSMSaction = gnaT.updateNoDataAlarmFile(strAlarmsFolder, strCurrentAlarmState);

            Console.WriteLine("     Alarm status: " + strCurrentAlarmState);
            Console.WriteLine("     Action: " + strSMSaction);



            string strFullSMSmessage = "No SMS message";
            //strSMSaction = "SendSMS" / "DoNotSendSMS";

            if ((strCurrentAlarmState == "Alarm") && (strSMSaction == "SendSMS"))
            {
                smsAlarmState = "New alarms";
                txtMessage = "Send SMS";
                strFullSMSmessage = "Alarm:\nNo data";
                strEmailMessage = "No data received in monitoringDB in past "+Convert.ToString(dblAlarmWindowHrs)+"hrs. \nEither ATS failure or T4D server has stopped processing";
            }


            if ((strCurrentAlarmState == "Alarm") && (strSMSaction == "DoNotSendSMS"))
            {
                smsAlarmState = "Existing alarms"; 
                txtMessage = "No SMS";
                // No action needed

            }


            if ((strCurrentAlarmState == "No Alarm") && (strSMSaction == "SendSMS"))
            {
                smsAlarmState = "Alarms cancelled";
                txtMessage = "Send SMS";
                strFullSMSmessage = "No data alarm cancelled";
                strEmailMessage = "Data being received by monitoringDB. \nAlarm state reset to OK";
            }

            if ((strCurrentAlarmState == "No Alarm") && (strSMSaction == "DoNotSendSMS"))
            {
                smsAlarmState = "No alarms";
                txtMessage = "No SMS";
                // No action needed
            }

            Console.WriteLine("     Alarm state: " + strCurrentAlarmState);
            Console.WriteLine("       " + txtMessage);
            Console.WriteLine("       " + smsAlarmState);
            Console.WriteLine("       " + strFullSMSmessage);


            // Send SMS
            if (strSMSaction == "SendSMS")
            {
                // Send the SMS
                Console.WriteLine("     Send SMS");
                //gnaT.sendSMSArray(strFullSMSmessage, smsMobile);

                string strMessage = strNow + " : Dataflow failure alarm: " + strSMSTitle + " " + strEmailMessage + " (" + strMobileList + ")";
                gnaT.updateSystemLogFile(strRootFolder, strMessage);
                
                Console.WriteLine("     Send email");
                string strAlarmHeading = "ALARM STATUS:" + strSMSTitle + " (" + strNow + ")";
                strMessage = gnaT.addCopyright(strEmailMessage);

                SmtpMail oMail = new("ES-E1582190613-00131-72B1E1BD67B73FVA-C5TC1DDC612457A3")
                {
                    From = strEmailFrom,
                    To = new AddressCollection(strEmailRecipients),
                    Subject = strAlarmHeading,
                    TextBody = strFullSMSmessage
                };

                // SMTP server address
                SmtpServer oServer = new("smtp.gmail.com")
                {
                    User = strEmailLogin,
                    Password = strEmailPassword,
                    ConnectType = SmtpConnectType.ConnectTryTLS,
                    Port = 587
                };

                //Set sender email address, please change it to yours


                SmtpClient oSmtp = new();
                oSmtp.SendMail(oServer, oMail);

                strMessage = strAlarmHeading + " (emailed " + strEmailRecipients + ")";

                gnaT.updateSystemLogFile(strRootFolder, strMessage);
            }
            else
            {

                Console.WriteLine("     No alarm SMS or emails sent");
            }

ThatsAllFolks:

            Console.WriteLine("\ndataFlowAlarm checking completed...");
            gnaT.freezeScreen(strFreezeScreen);

            Environment.Exit(0);
        }
    }
}

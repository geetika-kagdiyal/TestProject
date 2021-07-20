using Microsoft.Office.Server.ActivityFeed;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;
using System.IO;
using System.Reflection;
using System.Xml.Linq;
using DocumentFormat.OpenXml.Packaging;
using System.Xml;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using MariGold.OpenXHTML;

namespace ApprovalNoteAzure
{
    class common
    {

        public const string SiteURL = "https://mypiramal.sharepoint.com/sites/piramaldev/creditopstest";

        public static dynamic parsed;

        public string emailTemplate;

        public List<WFTaskListItems> WFTaskListItemList;

        /// <summary>
        /// Current Environment from where the request is generated
        /// </summary>
        public static string CurrentEnvironment;

        /// <summary>
        /// Current SiteUrl from where the request is generated
        /// </summary>
        public static string CurrentSiteUrl;

        /// <summary>
        /// Constant for Username
        /// </summary>
        //public const string UserName = "SVC-Automation.UAT1@piramal.com";
        public static string UserName;

        /// <summary>
        /// Constant for Password
        /// </summary>
        //public const string Password = "@utoMTU1";
        public static string Password;

        public static void GetSiteUrl()
        {
            switch (common.CurrentEnvironment)
            {
                case "DEV":
                    CurrentSiteUrl = Constants.DEVSiteURL;
                    UserName = Constants.DEVUserName;
                    Password = Constants.DEVPassword;
                    break;
                case "UAT":
                    CurrentSiteUrl = Constants.UATSiteURL;
                    UserName = Constants.UATUserName;
                    Password = Constants.UATPassword;
                    break;
                case "PROD":
                    CurrentSiteUrl = Constants.PRODSiteURL;
                    UserName = Constants.PRODUserName;
                    Password = Constants.PRODPassword;
                    break;
            }
        }
        public static string GetDateddmmyyyy(string inputDate)
        {
            string outputDate = "";
            try
            {
                if (!string.IsNullOrEmpty(inputDate))
                {
                    DateTime dDate = DateTime.Parse(inputDate, CultureInfo.InvariantCulture);
                    outputDate = dDate.ToString("dd/MM/yyyy");
                }
            }
            catch (Exception ex)
            {

            }

            return outputDate;
        }

        public static string GetDateddmmyyyyhhmm(string inputDate)
        {
            string outputDate = "";
            try
            {
                if (!string.IsNullOrEmpty(inputDate))
                {
                    DateTime dDate = DateTime.Parse(inputDate, CultureInfo.InvariantCulture);
                    //outputDate = dDate.ToString("dd/MM/yyyy hh:mm tt");
                    outputDate = dDate.ToString("dd-MMM-yyyy hh:mm tt");
                }
            }
            catch (Exception ex)
            {

            }

            return outputDate;
        }

        public static string GetDateFieldInIST(string inputDate)
        {
            string ISTDate = "";
            try
            {
                if (!string.IsNullOrEmpty(inputDate))
                {
                    DateTime timeUtc = DateTime.Parse(inputDate, null, System.Globalization.DateTimeStyles.AdjustToUniversal);
                    TimeZoneInfo IndZone = TimeZoneInfo.FindSystemTimeZoneById(Constants.IndTimeZone);
                    DateTime IndTime = TimeZoneInfo.ConvertTimeFromUtc(timeUtc, IndZone);
                    ISTDate = GetDateddmmyyyy(IndTime.ToString());
                }
            }
            catch (Exception ex)
            {
                //Common.logError("GetDateFieldInIST", ex.Message, "88");
            }
            return ISTDate;
        }

        public static string GetRefreshedOnDate(string inputDate)
        {
            string RefreshedOnDate = "";
            try
            {
                if (!string.IsNullOrEmpty(inputDate))
                {
                    DateTime timeUtc = DateTime.Parse(inputDate, null, System.Globalization.DateTimeStyles.AdjustToUniversal);
                    TimeZoneInfo IndZone = TimeZoneInfo.FindSystemTimeZoneById(Constants.IndTimeZone);
                    DateTime IndTime = TimeZoneInfo.ConvertTimeFromUtc(timeUtc, IndZone);
                    RefreshedOnDate = GetDateddmmyyyyhhmm(IndTime.ToString());
                }
            }
            catch (Exception ex)
            {
                //common.objDataAccess.logError("GetDateFieldInIST", ex.Message, "88",  common.parsed["DeviationID"].Value);
            }
            return RefreshedOnDate;
        }

        public static string GetCurrentFinancialYear(DateTime dtDate)
        {
            int CurrentYear = dtDate.Year;
            int PreviousYear = dtDate.Year - 1;
            int NextYear = dtDate.Year + 1;
            string PreYear = PreviousYear.ToString();
            string NexYear = NextYear.ToString();
            string CurYear = CurrentYear.ToString();
            string FinYear = null;

            if (dtDate.Month > 3)
                FinYear = CurYear + "-" + NexYear.Substring(2);
            else
                FinYear = PreYear + "-" + CurYear.Substring(2);
            return FinYear.Trim();
        }

        private string SaveHtmlFile(DeviationDetailsProps objDeviationDetails, string htmlContent)
        {
            string filePath = string.Empty;
            try
            {
                string fileName = objDeviationDetails.DeviationGUID + "_" + DateTime.Now.ToString("yyyy'-'MM'-'dd'T'HH':'mm':'ss").Replace("-", "").Replace(":", "") + ".html";
                filePath = @"D:\home\site\wwwroot\bin\" + fileName;
                objDeviationDetails.InsideSaveHtmlFile = objDeviationDetails.InsideSaveHtmlFile + " Before WriteAllText";
                File.WriteAllText(filePath, htmlContent);
                objDeviationDetails.InsideSaveHtmlFile = objDeviationDetails.InsideSaveHtmlFile + " After WriteAllText";
            }
            catch(Exception ex)
            {
                DataAccess dataAccess = new DataAccess();
                dataAccess.logError("SaveHtmlFile", ex.Message, "108", objDeviationDetails.DeviationItemId);
            }
            return filePath;
        }

        public void CreateCollaborationWordFile(DeviationDetailsProps objDeviationDetails, StringBuilder htmlContent)
        {
            string destLoc = string.Empty;
            string docFileName = objDeviationDetails.DeviationGUID + "_" + DateTime.Now.ToString("yyyy'-'MM'-'dd'T'HH':'mm':'ss").Replace("-", "").Replace(":", "") + ".docx";
            destLoc = @"D:\home\site\wwwroot\bin\" + docFileName;
            try
            {
                WordDocument doc = new WordDocument(destLoc);
                doc.Process(new HtmlParser(htmlContent.ToString()));
                doc.Save();
                HTMLtoPDF hTMLtoPDF = new HTMLtoPDF();
                hTMLtoPDF.SaveDocInList(destLoc, objDeviationDetails);
            }
            catch(Exception ex)
            {
                DataAccess dataAccess = new DataAccess();
                dataAccess.logError("CreateCollaborationWordFile", ex.Message, "108", objDeviationDetails.DeviationItemId);
            }
            finally
            {
                if (File.Exists(destLoc))
                {
                    File.Delete(destLoc);
                }
            }
        }
    }
}

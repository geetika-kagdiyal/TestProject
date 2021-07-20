using System;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
//using Microsoft.SharePoint.Client;
using Newtonsoft.Json;

namespace ApprovalNoteAzure
{
    public static class ApprovalNote
    {
        #region SharePoint Context
        //private static string WebSPOUrl = "https://mypiramal.sharepoint.com/sites/piramaldevops/creditopsdev";
        private static string WebUrl = "https://mypiramal.sharepoint.com";
        private static string WebSPOUrl = "https://mypiramal.sharepoint.com/sites/piramaldev/creditopstest";

        private static string userName = "SVC-Automation.UAT1@piramal.com";
        private static string userPassword = "@utoMTU1";

        //private static string userName = "SVC-Automation.support1@piramal.com";
        //private static string userPassword = "@utoMTS1";

        #endregion


        [FunctionName("ApprovalNote")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)]HttpRequestMessage req, TraceWriter log)
        {
            log.Info("C# HTTP trigger function processed a request.");

            
            common common = new common();
            // This reads your post request body into variable "data"
            string data = await req.Content.ReadAsStringAsync();
            // Here you can process json into an object
            common.parsed = JsonConvert.DeserializeObject(data);

            string jsonObj = string.Empty;

            var res = new HttpResponseMessage(HttpStatusCode.OK);
            common.CurrentEnvironment = common.parsed["CurrentEnvironment"].Value;

            jsonObj = InitiateExecution(common.parsed, log);

            res.Content = new StringContent(jsonObj, Encoding.UTF8, "text/json");
            return res;
        }

        private static string InitiateExecution(dynamic parsed, TraceWriter log)
        {
            string jsonData = "{'PDFUrl':'' ,'EmailBody':''}";
            string emailBody = string.Empty;
            DeviationDetailsProps objDeviationDetails = new DeviationDetailsProps();
            DataAccess objDataAccess = new DataAccess();
            CreateHtml createHtml = new CreateHtml();
            string strDeviationID = parsed["DeviationID"].Value;
            try
            {
                objDeviationDetails = (DeviationDetailsProps)objDataAccess.GetDNData(strDeviationID, parsed);
                createHtml.CreateHtmlFile(objDeviationDetails, parsed);
                if (parsed["UpdatePDF"] == null)
                {
                    objDataAccess.GetStatusPageUrl();
                    createHtml.GenerateApprovalEmailBody(objDeviationDetails, parsed);
                    createHtml.GenerateFinalEmailBody(objDeviationDetails, parsed);
                    createHtml.GenerateRejectionEmailBody(objDeviationDetails, parsed);
                    createHtml.GenerateDMSEmailBody(objDeviationDetails, parsed);

                    jsonData = "{PDFUrl:" + objDeviationDetails.PDFUrl + ", EmailBody:" + objDeviationDetails.ApprovalMailBody +
                        ", FinalEmailBody:" + objDeviationDetails.FinalMailBody + ", PDFItemID:" + objDeviationDetails.PDFFileItemID +
                        ", LenderNonExistenceMailBody:" + objDeviationDetails.LenderNonExistenceMailBody + ", FolderNonExistenceMailBody:" + 
                        objDeviationDetails.FolderNonExistenceMailBody + ", RejectedMailBody:" + objDeviationDetails.RejectedMailBody + "}";
                }
                else
                {
                    jsonData = "{PDFUrl:" + objDeviationDetails.PDFUrl + "}";
                }

                if (parsed["CurrentUserEmail"] != null)
                {
                    objDataAccess.SendPDFInEmail(objDeviationDetails, parsed);
                }
            }
            catch (Exception ex)
            {
                objDataAccess.logError("InitiateExecution", ex.Message, "72", strDeviationID);
            }
            return jsonData;
        }
    }
}


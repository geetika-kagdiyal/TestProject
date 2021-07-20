using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace ApprovalNoteAzure
{
    public class HTMLtoPDF
    {
        #region SharePoint Context
        //private static string WebSPOUrl = "https://mypiramal.sharepoint.com/sites/piramaldevops/creditopsdev";
        private static string WebUrl = "https://mypiramal.sharepoint.com";
        private static string WebSPOUrl = "https://mypiramal.sharepoint.com/sites/piramaldev/creditopstest";

        private static string userName = "SVC-Automation.UAT1@piramal.com";
        private static string userPassword = "@utoMTU1";
        DataAccess objDataAccess = new DataAccess();

        //private static string userName = "SVC-Automation.support1@piramal.com";
        //private static string userPassword = "@utoMTS1";
        static ClientContext oSPContext = null;
        //TraceWriter log;

        public static ClientContext SPContext
        {
            get
            {
                if (oSPContext is null)
                {
                    common.GetSiteUrl();
                    oSPContext = new ClientContext(common.CurrentSiteUrl);
                    oSPContext.Credentials = new SharePointOnlineCredentials(common.UserName, GeneralData.ToSecureString(common.Password.ToCharArray()));
                }
                return oSPContext;
            }
        }


        #endregion

        public string HTML_To_PDF(string strSiteURL, string strFolderPath, string strPDFName, string strHtmlContent, string strPDFHeader, string strPDFFooter,string strDeviation)
        {
            string strFileFolderName = "";
            string strURL = "";
            StringBuilder sbFooter = new StringBuilder();
            try
            {
                if (strFolderPath != "")
                {
                    strFileFolderName = strFolderPath + @"\" + strPDFName;
                }
                else
                {
                    strFileFolderName = strPDFName;
                }

                NReco.PdfGenerator.HtmlToPdfConverter htmlToPdf = new NReco.PdfGenerator.HtmlToPdfConverter();
                htmlToPdf.PdfToolPath = @"D:\home\site\wwwroot\bin\";
                sbFooter.AppendLine("<div style = 'text-align:center; font-size:14px;'> Page <span class='page'></span> of &nbsp;<span class='topage'></span><div>");
                htmlToPdf.PageHeaderHtml = strPDFHeader;
                htmlToPdf.PageFooterHtml = sbFooter.ToString();
                htmlToPdf.Margins.Bottom = 10;

                byte[] pdfBytes = htmlToPdf.GeneratePdf(strHtmlContent);

               // objDataAccess.logError("HTML_To_PDF", "", "1", strDeviation);
                if (pdfBytes != null)
                {
                    strURL = SavePDFInList(pdfBytes, strFileFolderName, false, strPDFName);
                }

            }
            catch (Exception ex)
            {
                objDataAccess.logError("HTML_To_PDF", ex.Message.ToString(), "0", strDeviation);
            }
            return strURL;
        }

        private static string GetScriptPath()
    => Path.Combine(GetEnvironmentVariable("HOME"), @"site\wwwroot");

        private static string GetEnvironmentVariable(string name)
            => System.Environment.GetEnvironmentVariable(name, EnvironmentVariableTarget.Process);

        public string SavePDFInList(byte[] pdfBytes1, string strFileFolder, bool isFinalStage,string strPDFName)
        {
            string strReturnURL = "";
            string fileUrl = "";
            //string siteUrl = "https://mypiramal.sharepoint.com";
            StringBuilder sbLineNo = new StringBuilder();
            //sbLineNo.Appendline(",44");

            strPDFName=strPDFName.Replace("/", "_");
            /*if(facilityId.Contains("/"))
            {
                facilityId = facilityId.Replace("/", "%5C");
                facilityId = "";
            }*/

            try
            {
                //string sFile = @"D:\home\site\wwwroot\bin\Disbursement_Checklist_" + facilityId.Replace("/", "%5C") + "_" + dcDataItemId + ".pdf"; //Path
                // string sFile = @"D:\home\site\wwwroot\bin\Disbursement_Checklist_" + dCMainList.TransactionIqId + "_" + dCMainList.ItemId + ".pdf";

                string sFile = @"D:\home\site\wwwroot\bin\"+ strPDFName;
                System.IO.File.WriteAllBytes(sFile, pdfBytes1);
                //////sbLineNo.Appendline(",47");
                var formLib = SPContext.Web.Lists.GetByTitle(Constants.PDFFile_DocumentLibrary);
                //var formLib = SPContext.Web.Lists.GetByTitle("Documents");
                SPContext.Load(formLib.RootFolder);
                SPContext.ExecuteQuery();
                //string fileName = @"D:\home\site\wwwroot\bin\Disbursement_Checklist_" + facilityId.Replace("/", "%5C") + "_" + dcDataItemId + ".pdf";
                //string fileName = @"D:\home\site\wwwroot\bin\Disbursement_Checklist_" + dCMainList.TransactionIqId + "_" + dCMainList.ItemId + ".pdf";
                string fileName = @"D:\home\site\wwwroot\bin\"+ strPDFName;
                
                ////sbLineNo.Appendline(",52");
                using (var fs = new FileStream(fileName, FileMode.Open))
                {
                    var fi = new FileInfo(fileName); //file Title  
                    fileUrl = String.Format("{0}/{1}", formLib.RootFolder.ServerRelativeUrl, fi.Name);
                    Microsoft.SharePoint.Client.File.SaveBinaryDirect(SPContext, fileUrl, fs, true);
                    SPContext.ExecuteQuery();
                  ////sbLineNo.Appendline(",59");
                    fs.Close();
                   ////sbLineNo.Appendline(",61");
                    //if (isFinalStage)
                    //{
                    //    //UpdatePDFFileSize(fi.Name, transactionIqId);
                    //    UpdatePDFFileSize(fi.Name, dCMainList);
                    //}
                    //sbLineNo.Appendline(",63");
                }
                /*if (System.IO.File.Exists(@"D:\home\site\wwwroot\bin\Disbursement_Checklist_" + facilityId + "_" + dcDataItemId + ".pdf"))
                {
                    System.IO.File.Delete(@"D:\home\site\wwwroot\bin\Disbursement_Checklist_" + facilityId + "_" + dcDataItemId + ".pdf");
                }*/
                //if (System.IO.File.Exists(@"D:\home\site\wwwroot\bin\Disbursement_Checklist_" + dCMainList.TransactionIqId + "_" + dCMainList.ItemId + ".pdf"))
                //{
                //    System.IO.File.Delete(@"D:\home\site\wwwroot\bin\Disbursement_Checklist_" + dCMainList.TransactionIqId + "_" + dCMainList.ItemId + ".pdf");
                //}

                if (System.IO.File.Exists(@"D:\home\site\wwwroot\bin\"+ strPDFName))
                {
                    System.IO.File.Delete(@"D:\home\site\wwwroot\bin\" + strPDFName);
                }
                //sbLineNo.Appendline(",64");
                //strReturnURL = siteUrl + fileUrl;
                strReturnURL = fileUrl;
            }
            catch (Exception ex)
            {
                //sbLineNo.Appendline(",77");
                string lineno = sbLineNo.ToString();
                strReturnURL = fileUrl;
                //logError("GetListItemFolder : " + lineno, ex.Message, "78", dCMainList.TransactionIqId);
                //fileUrl = GetListItemFolderException(pdfBytes1, facilityId, dcDataItemId, isFinalStage, transactionIqId);
                //fileUrl = GetListItemFolderException(pdfBytes1, dCMainList, isFinalStage);
                objDataAccess.logError("SavePDFInList", ex.Message.ToString(), "45", strFileFolder);
            }
            return strReturnURL;
        }

        public void SaveDocInList(string FilePath, DeviationDetailsProps objDeviationDetails)
        {
            try
            {
                #region ConnectToSharePoint
                var securePassword = new SecureString();
                foreach (char c in common.Password)
                { securePassword.AppendChar(c); }
                var onlineCredentials = new SharePointOnlineCredentials(common.UserName, securePassword);
                #endregion
                #region Insert the data
                using (ClientContext CContext = new ClientContext(common.CurrentSiteUrl))
                {
                    CContext.Credentials = onlineCredentials;
                    Web web = CContext.Web;
                    FileCreationInformation newFile = new  FileCreationInformation();
                    objDeviationDetails.InsideSaveDocInList = objDeviationDetails.InsideSaveDocInList + " Before ReadAllBytes";
                    byte[] FileContent = System.IO.File.ReadAllBytes(FilePath);
                    objDeviationDetails.InsideSaveDocInList = objDeviationDetails.InsideSaveDocInList + " After ReadAllBytes";
                    objDeviationDetails.InsideSaveDocInList = objDeviationDetails.InsideSaveDocInList + " Before MemoryStream";
                    newFile.ContentStream = new MemoryStream(FileContent);
                    objDeviationDetails.InsideSaveDocInList = objDeviationDetails.InsideSaveDocInList + " After MemoryStream";
                    objDeviationDetails.InsideSaveDocInList = objDeviationDetails.InsideSaveDocInList + " Before GetFileName";
                    newFile.Url = Path.GetFileName(FilePath);
                    objDeviationDetails.InsideSaveDocInList = objDeviationDetails.InsideSaveDocInList + " After GetFileName";
                    objDeviationDetails.InsideSaveDocInList = objDeviationDetails.InsideSaveDocInList + " Before GetByTitle";
                    List DocumentLibrary = web.Lists.GetByTitle(Constants.PDFFile_DocumentLibrary);
                    //Folder folder = DocumentLibrary.RootFolder.Folders.GetByUrl(ClientSubFolder);
                    //Folder Clientfolder = DocumentLibrary.RootFolder.Folders.Add(ClientSubFolder);
                    //Clientfolder.Update();

                    objDeviationDetails.InsideSaveDocInList = objDeviationDetails.InsideSaveDocInList + " After GetByTitle";
                    objDeviationDetails.InsideSaveDocInList = objDeviationDetails.InsideSaveDocInList + " Before RootFolder";
                    Microsoft.SharePoint.Client.File uploadFile = DocumentLibrary.RootFolder.Files.Add(newFile);

                    objDeviationDetails.InsideSaveDocInList = objDeviationDetails.InsideSaveDocInList + " After RootFolder";
                    objDeviationDetails.InsideSaveDocInList = objDeviationDetails.InsideSaveDocInList + " Before Load";
                    //CContext.Load(DocumentLibrary);
                    CContext.Load(uploadFile);

                    objDeviationDetails.InsideSaveDocInList = objDeviationDetails.InsideSaveDocInList + " After Load";
                    objDeviationDetails.InsideSaveDocInList = objDeviationDetails.InsideSaveDocInList + " Before ExecuteQuery";
                    CContext.ExecuteQuery();

                    objDeviationDetails.InsideSaveDocInList = objDeviationDetails.InsideSaveDocInList + " After ExecuteQuery";
                    //objDeviationDetails.DNCollaborationServerRelativeUrl = 
                    GetCollboratingDocUrl(objDeviationDetails, uploadFile.ServerRelativeUrl);
                    objDataAccess.UpdateRefNo(objDeviationDetails);
                }
                #endregion
            }
            catch (Exception ex)
            {
                objDataAccess.logError("SaveDocInList", ex.Message.ToString(), "45", objDeviationDetails.DeviationGUID);
            }
        }

        private void GetCollboratingDocUrl(DeviationDetailsProps objDeviationDetails, string FileRelativeURL)
        {
            try
            {
                Microsoft.SharePoint.Client.File file = SPContext.Web.GetFileByServerRelativeUrl(FileRelativeURL);
                SPContext.Load(file, f => f.ListItemAllFields);
                SPContext.ExecuteQuery();
                ListItem item = file.ListItemAllFields;
                if (item != null)
                {
                    FieldUrlValue urlValue = (FieldUrlValue)item.FieldValues["_dlc_DocIdUrl"];
                    objDeviationDetails.DNCollaborationDocUrl = urlValue.Url.ToString();
                }
            }
            catch(Exception ex)
            {
                objDataAccess.logError("GetCollboratingDocUrl", ex.Message.ToString(), "45", objDeviationDetails.DeviationGUID);
            }
        }
    }

    public static class GeneralData
    {
        public static SecureString ToSecureString(this char[] _self)
        {
            SecureString knox = new SecureString();
            foreach (char c in _self)
            {
                knox.AppendChar(c);
            }
            return knox;
        }

    }

    
}

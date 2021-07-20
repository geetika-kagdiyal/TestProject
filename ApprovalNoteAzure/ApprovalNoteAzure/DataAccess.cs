using Microsoft.SharePoint;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.UserProfiles;
using Microsoft.SharePoint.Client.Utilities;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace ApprovalNoteAzure
{
    public class DataAccess
    {
        static ClientContext oSPContext = null;
        DeviationDetailsProps objDeviationDetails = new DeviationDetailsProps();

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

        public void logError(string functionName, string errorMessage, string lineno, string inputId)
        {
            try
            {
                List oList = SPContext.Web.Lists.GetByTitle("ErrorList");
                ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                ListItem oListItem = oList.AddItem(itemCreateInfo);

                oListItem["Title"] = "DeviationNote";
                oListItem["Function"] = functionName;
                oListItem["error"] = errorMessage;
                oListItem["lineno"] = lineno;
                oListItem["inputId"] = inputId;
                oListItem.Update();
                SPContext.Load(oListItem);
                SPContext.ExecuteQuery();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public DeviationDetailsProps GetDNData(string strDeviationID, dynamic parsed)
        {
            try
            {
                GetDeviationNote(strDeviationID, parsed);
                GetDeviationFacilityMap(objDeviationDetails.DeviationItemId);
                string strDeal = Convert.ToString(objDeviationDetails.DealItemId);
                GetDealDetails(strDeal);
                GetDeviationsNoteDetails(objDeviationDetails.DeviationItemId);
                if (parsed["ApprovedByEmailId"] != null)
                {
                    string DesignationDepartment = GetUserDataFromSharePointOnline(parsed, parsed["ApprovedByEmailId"].Value);
                    UpdateWFTaskList(parsed, Convert.ToInt32(parsed["WorkFlowItemId"].Value), DesignationDepartment);
                }
                if (parsed["WorkflowIDEmail"] != null)
                {
                    var collection = JsonConvert.DeserializeObject(parsed["WorkflowIDEmail"].Value);
                    foreach(var item in collection)
                    {
                        string DesignationDepartment = GetUserDataFromSharePointOnline(parsed, Convert.ToString(item["Email"]));
                        UpdateWFTaskList(parsed, Convert.ToInt32(item["WorkflowID"]), DesignationDepartment);
                    }
                    
                }
                GetWorkFlowStatus(parsed);

                List<FileLink> lstFile = new List<FileLink>();
                lstFile = GetDeviationNoteFileLink(objDeviationDetails, parsed);
                objDeviationDetails.DeviationNoteFileList = lstFile;
            }
            catch (Exception ex)
            {
                logError("GetDNData", ex.Message, "68", objDeviationDetails.DeviationItemId);
            }
            return objDeviationDetails;
        }

        private void GetDeviationFacilityMap(string deviationID)
        {
            try
            {
                var list = SPContext.Web.Lists.GetByTitle("DeviationFacilityMap");
                CamlQuery camlQuery = new CamlQuery()
                {
                    ViewXml = "<View><Query><Where>" +
                    "<Eq><FieldRef Name='DeviationRefNo' LookupId='TRUE'/><Value Type='Lookup'>" + deviationID + "</Value></Eq>" +
                    "</Where></Query></View>"
                };
                ListItemCollection collListItem = list.GetItems(camlQuery);
                SPContext.Load(collListItem);
                SPContext.ExecuteQuery();
                foreach (ListItem oListItem in collListItem)
                {
                    FieldLookupValue DealLookup = oListItem["Deal_x003a_Deal_x0020_Name"] as FieldLookupValue;
                    objDeviationDetails.DealItemId = Convert.ToString(DealLookup.LookupId);
                    objDeviationDetails.DealName = DealLookup.LookupValue;
                }
                objDeviationDetails.PDFName = objDeviationDetails.DealName + "_" + objDeviationDetails.DeviationItemId;
            }
            catch (Exception ex)
            {
                logError("GetDeviationFacilityMap", ex.Message, "108", objDeviationDetails.DeviationItemId);
            }
        }

        private void GetDealDetails(string strDealId)
        {
            try
            {
                var list = SPContext.Web.Lists.GetByTitle("Deals");
                CamlQuery camlQuery = new CamlQuery()
                {
                    ViewXml = "<View><Query><Where><Eq><FieldRef Name='ID' /><Value Type='Text'>" + strDealId + "</Value></Eq></Where></Query></View>"
                };
                ListItemCollection collListItem = list.GetItems(camlQuery);
                SPContext.Load(collListItem);
                SPContext.ExecuteQuery();
                foreach (ListItem oListItem in collListItem)
                {
                    objDeviationDetails.DealItemId = Convert.ToString(oListItem["ID"]);
                    objDeviationDetails.DealID = Convert.ToString(oListItem["DealID"]);
                    objDeviationDetails.DealName = Convert.ToString(oListItem["Title"]);
                    objDeviationDetails.BusinessSegment = Convert.ToString(oListItem["BusinessSegment"]);

                    FieldLookupValue RegionLookup = oListItem["Region_x003a_Region_x0020_Short_"] as FieldLookupValue;
                    objDeviationDetails.RegionShortCode = RegionLookup.LookupValue;
                }
            }
            catch (Exception ex)
            {
                logError("GetDealDetails", ex.Message, "108", objDeviationDetails.DeviationItemId);
            }
        }

        private void GetDeviationNote(string strDeviationID, dynamic parsed)
        {
            try
            {
                var list = SPContext.Web.Lists.GetByTitle("DeviationNote");
                CamlQuery camlQuery = new CamlQuery()
                {
                    ViewXml = "<View><Query><Where><Eq><FieldRef Name='GUID' /><Value Type='Text'>" + strDeviationID + "</Value></Eq></Where></Query></View>"
                };
                ListItemCollection collListItem = list.GetItems(camlQuery);
                SPContext.Load(collListItem);
                SPContext.ExecuteQuery();
                foreach (ListItem oListItem in collListItem)
                {
                    objDeviationDetails.DeviationItemId = Convert.ToString(oListItem["ID"]);
                    objDeviationDetails.DeviationTitle = Convert.ToString(oListItem["Title"]);
                    objDeviationDetails.NatureOfDeviation = Convert.ToString(oListItem["NatureOfDeviation"]);
                    objDeviationDetails.DeviationGUID = Convert.ToString(oListItem["GUID"]);
                    objDeviationDetails.Tenure = Convert.ToString(oListItem["UpdatedTenure"]);
                    objDeviationDetails.Purpose = Convert.ToString(oListItem["UpdatedPurposeofDeal"]);
                    objDeviationDetails.SecurityDetails_Additional = Convert.ToString(oListItem["UpdatedSecurityAddtional"]);
                    objDeviationDetails.SecurityDetails_Primary = Convert.ToString(oListItem["UpdatedSecurityPrimary"]);
                    objDeviationDetails.SecurityDetails_Collateral = Convert.ToString(oListItem["UpdatedSecurityCollaterals"]);
                    objDeviationDetails.Cover = Convert.ToString(oListItem["UpdatedCashCover"]);
                    objDeviationDetails.PrincipalServicing = Convert.ToString(oListItem["UpdatedRepaymentSchedule"]);
                    objDeviationDetails.IRR_CoupounRate = Convert.ToString(oListItem["UpdatedIRR"]);
                    objDeviationDetails.CouponRate = Convert.ToString(oListItem["UpdatedCouponRate"]);
                    objDeviationDetails.ProcessingFee = Convert.ToString(oListItem["UpdatedProcessingFee"]);
                    objDeviationDetails.SecurityCover = Convert.ToString(oListItem["UpdatedSecurityCover"]);
                    objDeviationDetails.ConductOfAccount = Convert.ToString(oListItem["ConductOfAccount"]);
                    objDeviationDetails.Borrowergroup = Convert.ToString(oListItem["BorrowerGroupName"]);
                    objDeviationDetails.SanctionedAmount = Convert.ToString(oListItem["UpdatedDealSanctionAmount"]);
                    objDeviationDetails.ExposureType = Convert.ToString(oListItem["FacilityInstType"]);

                    objDeviationDetails.UCIC = Convert.ToString(oListItem["UCIC"]);
                    objDeviationDetails.Borrower = Convert.ToString(oListItem["BorrowerName"]);
                    objDeviationDetails.DateOfSanction = Convert.ToString(oListItem["UpdatedSanctionDate"]);
                    objDeviationDetails.Annexure = Convert.ToString(oListItem["Annexure"]);
                    objDeviationDetails.WorkflowStatus = Convert.ToString(oListItem["WorkflowStatus"]);

                    if (parsed["Note"] == "Deviation")
                    {
                        objDeviationDetails.EditUrl = common.CurrentSiteUrl + "/SitePages/DeviationNoteWF.aspx?Deviation=" + objDeviationDetails.DeviationGUID + "&mode=Edit&Note=Deviation";
                    }
                    else
                    {
                        objDeviationDetails.EditUrl = common.CurrentSiteUrl + "/SitePages/DeviationNoteWF.aspx?Approval=" + objDeviationDetails.DeviationGUID + "&mode=Edit&Note=Approval";
                    }   
                    FieldLookupValue LenderLookup = oListItem["Lender"] as FieldLookupValue;
                    objDeviationDetails.Lender = LenderLookup.LookupValue;
                    objDeviationDetails.LenderShortCode = LenderLookup.LookupValue;
                    objDeviationDetails.PrepaymentTillDate = Convert.ToString(oListItem["PrepaymentTillDate1"]);
                    objDeviationDetails.DNNo = Convert.ToString(oListItem["ID"]);

                    objDeviationDetails.DashBoardUrl = Convert.ToString(oListItem["DashBoardUrl"]);

                    if (oListItem["CollaborationDocUrl"] != null)
                    {
                        FieldUrlValue urlValue = (FieldUrlValue)oListItem["CollaborationDocUrl"];
                        objDeviationDetails.DNCollaborationDocUrl = urlValue.Url.ToString();
                    }

                    if (!string.IsNullOrEmpty(Convert.ToString(oListItem["Modified"])))
                    {
                        string strSanctionDate = Convert.ToString(oListItem["Modified"]);
                        DateTime sourceDate = DateTime.Parse(strSanctionDate, null, System.Globalization.DateTimeStyles.AdjustToUniversal);
                        TimeZoneInfo IndZoneAuditDate = TimeZoneInfo.FindSystemTimeZoneById("India Standard Time");
                        sourceDate = TimeZoneInfo.ConvertTimeFromUtc(sourceDate, IndZoneAuditDate);
                        objDeviationDetails.LastModified = sourceDate.ToString("dd-MMM-yyyy");
                    }
                    else
                    {
                        objDeviationDetails.LastModified = "";
                    }
                }
            }
            catch (Exception ex)
            {
                logError("GetDeviationNote", ex.Message, "142", objDeviationDetails.DeviationItemId);
            }
        }

        private void GetFacilityDetails(string strDeviationID)
        {
            DataTable odtFacility = new DataTable();
            odtFacility.Columns.Add("FacilityID");
            try
            {
                odtFacility.Clear();
                var list = SPContext.Web.Lists.GetByTitle("DeviationFacilityMap");
                CamlQuery camlQuery = new CamlQuery()
                {
                    ViewXml = "<View><Query><Where>" +
                    "<Eq><FieldRef Name='DeviationRefNo' LookupId='TRUE'/><Value Type='Lookup'>" + strDeviationID + "</Value></Eq>" +
                    "</Where></Query></View>"
                };
                ListItemCollection collListItem = list.GetItems(camlQuery);
                SPContext.Load(collListItem);
                SPContext.ExecuteQuery();
                foreach (ListItem oListItem in collListItem)
                {
                    FieldLookupValue lookupFacilityID = oListItem["Facility"] as FieldLookupValue;
                    if (!string.IsNullOrEmpty(Convert.ToString(lookupFacilityID.LookupId)))
                    {
                        DataRow dtRow = odtFacility.NewRow();
                        dtRow["FacilityID"] = Convert.ToString(lookupFacilityID.LookupId);
                        odtFacility.Rows.Add(dtRow);
                    }
                }

                if (odtFacility.Rows.Count > 0)
                {
                    GetFacilityDetails(odtFacility);
                }
            }
            catch (Exception ex)
            {
                logError("GetFacilityDetails", ex.Message, "189", objDeviationDetails.DeviationItemId);
            }
        }

        private void GetDeviationsNoteDetails(string strDeviationID)
        {
            List<RTBDetails> objRTBDataList = new List<RTBDetails>();                       
            try
            {
                var list = SPContext.Web.Lists.GetByTitle("DeviationNoteDetails");
                CamlQuery camlQuery = new CamlQuery()
                {
                    ViewXml = "<View><Query><Where><And>" +
                    "<Eq><FieldRef Name='DNRefNo' LookupId='TRUE'/><Value Type='Lookup'>" + strDeviationID + "</Value></Eq>" +
                    "<Eq><FieldRef Name='IsDeleted' /><Value Type='Boolean'>0</Value></Eq>" +
                    "</And></Where></Query></View>"
                };
                ListItemCollection collListItem = list.GetItems(camlQuery);
                SPContext.Load(collListItem);
                SPContext.ExecuteQuery();
                foreach (ListItem oListItem in collListItem)
                {
                    RTBDetails objRTB = new RTBDetails();
                    objRTB.Originalterms = Convert.ToString(oListItem["OriginalTermsHTML"]);
                    objRTB.ProposedDeviation = Convert.ToString(oListItem["ProposedTermsHTML"]);
                    objRTB.Impact = Convert.ToString(oListItem["ImpactHTML"]);
                    objRTB.Justification = Convert.ToString(oListItem["JustificationHTML"]);

                    objRTBDataList.Add(objRTB);
                }

                objDeviationDetails.RTBData = objRTBDataList;
            }
            catch (Exception ex)
            {
                logError("GetDeviationsNoteDetails", ex.Message, "225", objDeviationDetails.DeviationItemId);
            }
        }


        private void GetFacilityDetails(DataTable oDT)
        {
            string strUCIC = "";
            string strBorrower = "";
            string strSanctionDate = "";
            try
            {
                string strCondition = "";
                foreach (DataRow oDrow in oDT.Rows)
                {
                    if (strCondition == "")
                    {
                        strCondition = "<Value Type='Integer'>" + Convert.ToString(oDrow["FacilityID"]) + "</Value>";
                    }
                    else
                    {
                        strCondition = strCondition + "<Value Type='Integer'>" + Convert.ToString(oDrow["FacilityID"]) + "</Value>";
                    }
                }

                var list = SPContext.Web.Lists.GetByTitle("Facility");

                CamlQuery camlQuery = new CamlQuery()
                {
                    ViewXml = "<View><Query><Where><In><FieldRef Name='ID' /><Values>"+strCondition+"</Values></In></Where></Query></View>"
                };
                
                ListItemCollection collListItem = list.GetItems(camlQuery);
                SPContext.Load(collListItem);
                SPContext.ExecuteQuery();
                foreach (ListItem oListItem in collListItem)
                {
                    string strSanctionDateFacility = Convert.ToString(oListItem["SanctionDate"]);
                    DateTime sourceDate = DateTime.Parse(strSanctionDateFacility, null, System.Globalization.DateTimeStyles.AdjustToUniversal);
                    TimeZoneInfo IndZoneAuditDate = TimeZoneInfo.FindSystemTimeZoneById("India Standard Time");
                    sourceDate = TimeZoneInfo.ConvertTimeFromUtc(sourceDate, IndZoneAuditDate);

                    if (strSanctionDate == "")
                    {
                        strSanctionDate = sourceDate.ToString("dd-MMM-yyyy");
                    }
                    else
                    {
                        if (strSanctionDate.Contains(sourceDate.ToString("dd-MMM-yyyy")))
                        { }
                        else
                        {
                            strSanctionDate = strSanctionDate + "," + sourceDate.ToString("dd-MMM-yyyy");
                        }
                    }

                    FieldLookupValue lookupBorrower = oListItem["Borrower"] as FieldLookupValue;//Changed Deal_x003a_Deal_x0020_Name To Facility_x003a_Facility_x0020_In
                    if (strBorrower == "")
                    {
                        strBorrower = Convert.ToString(lookupBorrower.LookupValue);
                    }
                    else
                    {
                        if (strBorrower.Contains(Convert.ToString(lookupBorrower.LookupValue)))
                        { }
                        else
                        {
                            strBorrower = strBorrower + "," + Convert.ToString(lookupBorrower.LookupValue);
                        }
                    }

                    FieldLookupValue lookupUCIC = oListItem["Borrower_x003a_UCIC"] as FieldLookupValue;
                    if (strUCIC == "")
                    {
                        strUCIC = Convert.ToString(lookupUCIC.LookupValue);
                    }
                    else
                    {
                        if (strUCIC.Contains(Convert.ToString(lookupUCIC.LookupValue)))
                        { }
                        else
                        {
                            strUCIC = strUCIC + "," + Convert.ToString(lookupUCIC.LookupValue);
                        }
                    }
                }

                objDeviationDetails.UCIC = strUCIC;
                objDeviationDetails.Borrower = strBorrower;
                objDeviationDetails.DateOfSanction = strSanctionDate;
                //}

            }
            catch (Exception ex)
            {
                logError("GetFacilityDetails", ex.Message, "225", objDeviationDetails.DeviationItemId);
            }
        }

        public void SendPDFInEmail(DeviationDetailsProps objDeviationDetails, dynamic parsed)
        {
            try
            {
                using (ClientContext _context = new ClientContext(common.CurrentSiteUrl))
                {
                    SecureString passWord = new SecureString();
                    List<FileLink> lstFile = objDeviationDetails.DeviationNoteFileList;
                    foreach (char c in common.Password.ToCharArray())
                    {
                        passWord.AppendChar(c);
                    }
                    // SharePoint Online Credentials  
                    _context.Credentials = new Microsoft.SharePoint.Client.SharePointOnlineCredentials(common.UserName, passWord);
                    // Get the SharePoint web  
                    Web web = _context.Web;
                    // Load the Web properties  
                    _context.Load(web);
                    // Execute the query to the server.  
                    _context.ExecuteQuery();

                    var emailprpoperties = new EmailProperties();
                    //emailprpoperties.BCC = new List<string> { "support.sp3@piramal.com" };
                    emailprpoperties.To = new List<string> { parsed["CurrentUserEmail"].Value };
                    //emailprpoperties.CC = new List<string> { "support.sp1@piramal.com" };
                    if(objDeviationDetails.WorkflowStatus != "Draft")
                    {
                        emailprpoperties.CC = GetCreditOpsUser(objDeviationDetails);
                    }
                    emailprpoperties.From = common.UserName;
                    StringBuilder sbText = new StringBuilder();

                    sbText.AppendLine("<b><span style='font-size:16px;font-family: Segoe UI;'>Dear User,</span></b></br></br>");
                    sbText.AppendLine("<span style='font-size:16px;font-family: Segoe UI;'>Below is the PDF Generated: </span></br></br>");
                    sbText.AppendLine("<span style='font-size:16px;font-family: Segoe UI;'><a href='" + objDeviationDetails.PDFUrl + "'>View PDF</a></span></br></br>");

                    sbText.AppendLine("<b><span style='font-size:16px;font-family: Segoe UI;'>Attachments :</span></b>");
                    if (lstFile != null)
                    {
                        if (lstFile.Count > 0)
                        {
                            sbText.AppendLine("</br><br>");
                            for (int i = 0; i < lstFile.Count; i++)
                            {
                                sbText.AppendLine("<span style='font-size:16px;font-family: Segoe UI;'>Attachment " + (i + 1) + " : </span><a style='font-size:16px;font-family: Segoe UI;' href=" + Convert.ToString(lstFile[i].Link) + ">" + Convert.ToString(lstFile[i].Name) + "</a></br><br>");
                            }
                        }
                        else
                        {
                            sbText.AppendLine("<b><span style='font-size:16px;font-family: Segoe UI;'>NA</span></b></br><br>");
                        }
                    }
                    else
                    {
                        sbText.AppendLine("<b><span style='font-size:16px;font-family: Segoe UI;'>NA</span></b></br><br>");
                    }
                    sbText.AppendLine("<span style='font-size:16px;font-family: Segoe UI;'>To update details please <a href='" + objDeviationDetails.EditUrl + "'>Click Here</a></span></br></br>");
                    sbText.AppendLine("<span style='font-size:16px;font-family: Segoe UI;'>This is a system generated email, kindly do not respond to this mail. For queries please drop an email to <a href='#'>pchfl.itws@piramal.com</a></span></br></br>");
                    sbText.AppendLine("<span style='font-size:16px;font-family: Segoe UI;'>Regards,</span></br>");
                    sbText.AppendLine("<span style='font-size:16px;font-family: Segoe UI;'>IT SUPPORT</span></br>");

                    emailprpoperties.Body = sbText.ToString();

                    if(parsed["Note"] == "Deviation")
                    {
                        emailprpoperties.Subject = "Deviation Note for Deal - " + objDeviationDetails.DealName + ". Deviation Reference No: - " + objDeviationDetails.DeviationTitle;
                    }
                    else
                    {
                        emailprpoperties.Subject = "Approval Note for Deal - " + objDeviationDetails.DealName + ". Approval Reference No: - " + objDeviationDetails.DeviationTitle;
                    }

                    Utility.SendEmail(_context, emailprpoperties);
                    _context.ExecuteQuery();
                }
            }
            catch (Exception ex)
            {
                logError("SendPDFInEmail", ex.Message, "1883", parsed["DeviationID"].Value);
            }
        }

        private List<string> GetCreditOpsUser(DeviationDetailsProps objDeviationDetails)
        {
            List<string> cOpsUser = new List<string>();
            try
            {
                GetDNStakeHolders();
                if (!string.IsNullOrEmpty(objDeviationDetails.CreditOpsUser))
                {
                    string[] creditOpsUser = objDeviationDetails.CreditOpsUser.Split(';');
                    for (int i = 0; i < creditOpsUser.Length; i += 1)
                    {
                        if (creditOpsUser[i] != "")
                        {
                            cOpsUser.Add(creditOpsUser[i]);
                        }
                    }
                }
            }
            catch(Exception ex)
            {
                logError("GetCreditOpsUser", ex.Message, "1883", objDeviationDetails.DeviationGUID);
            }
            return cOpsUser;
        }

        public void CreateLogonFromSP(int logoImagekey, string transactionIqId)
        {
            int logoImageItemId = 2;
            try
            {
                var list = SPContext.Web.Lists.GetByTitle("PiramalLogoImage");
                CamlQuery camlQuery = new CamlQuery
                {
                    ViewXml = "<View><Query><Where><Eq><FieldRef Name='keys' /><Value Type='Number'>" + logoImagekey + "</Value></Eq></Where></Query></View>"
                };
                ListItemCollection collListItem = list.GetItems(camlQuery);
                SPContext.Load(collListItem);
                SPContext.ExecuteQuery();
                foreach (ListItem oListItem in collListItem)
                {
                    logoImageItemId = Convert.ToInt32(oListItem["ID"]);
                }
            }
            catch (Exception ex)
            {
                logError("CreateLogonFromSP", ex.Message, "126", transactionIqId);
            }
            try
            {
                var list = SPContext.Web.Lists.GetByTitle("PiramalLogoImage");
                var listItem = list.GetItemById(logoImageItemId);
                SPContext.Load(list);
                SPContext.Load(listItem, i => i.File);
                SPContext.ExecuteQuery();

                var fileRef = listItem.File.ServerRelativeUrl;
                var fileInfo = Microsoft.SharePoint.Client.File.OpenBinaryDirect(SPContext, fileRef);
                var fileName = Path.Combine(@"D:\home\site\wwwroot\bin\", (string)listItem.File.Name);

                using (var fileStream = System.IO.File.Create(fileName))
                {
                    fileInfo.Stream.CopyTo(fileStream);
                    fileStream.Close();
                }
            }
            catch (Exception ex)
            {
                logError("CreateLogonFromSP", ex.Message, "148", transactionIqId);
            }

        }

        public string GetLogoInBase64(int logoImagekey, string transactionIqId)
        {
            string base64URl = string.Empty;
            int logoImageItemId = 2;
            try
            {
                var list = SPContext.Web.Lists.GetByTitle("PiramalLogoImage");
                CamlQuery camlQuery = new CamlQuery
                {
                    ViewXml = "<View><Query><Where><Eq><FieldRef Name='keys' /><Value Type='Number'>" + logoImagekey + "</Value></Eq></Where></Query></View>"
                };
                ListItemCollection collListItem = list.GetItems(camlQuery);
                SPContext.Load(collListItem);
                SPContext.ExecuteQuery();
                foreach (ListItem oListItem in collListItem)
                {
                    logoImageItemId = Convert.ToInt32(oListItem["ID"]);
                    base64URl = Convert.ToString(oListItem.FieldValues["Base64URl"]);
                }
            }
            catch (Exception ex)
            {
                logError("GetLogoInBase64", ex.Message, "126", transactionIqId);
            }
            return base64URl;
        }

        public string GetUserDataFromSharePointOnline(dynamic parsed, string user)
        {
            string targetUser = "i:0#.f|membership|" + user;
            string DesignationDepartment = "";
            try
            {
                using (ClientContext _context = new ClientContext(common.CurrentSiteUrl))
                {
                    //CommonClass.PrepareSPOnlineCredentials(parsed, _context);
                    SecureString passWord = new SecureString();
                    foreach (char c in common.Password.ToCharArray())
                    {
                        passWord.AppendChar(c);
                    }
                    // SharePoint Online Credentials  
                    _context.Credentials = new SharePointOnlineCredentials(common.UserName, passWord);
                    // Get the PeopleManager object. 
                    PeopleManager peopleManager = new PeopleManager(_context);
                    // Retrieve specific properties by using the GetUserProfilePropertiesFor method.  
                    string[] profilePropertyNames = new string[] { "AccountName", "Department", "Title" };
                    string[] DesignationDepart = new string[3];
                    int i = 0;

                    UserProfilePropertiesForUser profilePropertiesForUser = new UserProfilePropertiesForUser(
                        _context, targetUser, profilePropertyNames);

                    IEnumerable<string> profilePropertyValues = peopleManager.GetUserProfilePropertiesFor(profilePropertiesForUser);

                    // Load the request for the set of properties. 
                    _context.Load(profilePropertiesForUser);
                    _context.ExecuteQuery();

                    // Returned collection contains only property values 
                    foreach (var value in profilePropertyValues)
                    {
                        DesignationDepart[i++] = value;
                        //Console.WriteLine(value);
                    }

                    string dash = "";
                    if (string.IsNullOrEmpty(DesignationDepart[1]) || string.IsNullOrEmpty(DesignationDepart[2]))
                        dash = "";
                    else
                        dash = " - ";
                    DesignationDepartment = DesignationDepart[1] + dash + DesignationDepart[2];
                }
            }
            catch (Exception ex)
            {
                logError("GetUserDataFromSharePointOnline", ex.Message, "316", objDeviationDetails.DeviationItemId);
            }
            return DesignationDepartment;
        }

        public void UpdateWFTaskList(dynamic parsed, int Id, string desigDepart)
        {
            try
            {
                using (ClientContext _context = new ClientContext(common.CurrentSiteUrl))
                {
                    SecureString passWord = new SecureString();
                    foreach (char c in common.Password.ToCharArray())
                    {
                        passWord.AppendChar(c);
                    }
                    // SharePoint Online Credentials  
                    _context.Credentials = new SharePointOnlineCredentials(common.UserName, passWord);

                    List oList = _context.Web.Lists.GetByTitle(Constants.DN_WorkFlowTaskList);
                    ListItem oListItem = oList.GetItemById(Id);
                    oListItem["DesignationDepartment"] = desigDepart;
                    oListItem.Update();
                    _context.ExecuteQuery();
                }
            }
            catch (Exception ex)
            {
                logError("UpdateWFTaskList", ex.Message, "335", objDeviationDetails.DeviationItemId);
            }
        }

        private void GetWorkFlowStatus(dynamic parsed)
        {
            string userid = string.Empty;
            string username = string.Empty;
            try
            {
                var list = SPContext.Web.Lists.GetByTitle(Constants.DN_WorkFlowTaskList);
                CamlQuery camlQuery = new CamlQuery()
                {
                    ViewXml = "<View><Query><Where><Eq><FieldRef Name='DeviationId' /><Value Type='Text'>" + parsed["DeviationID"].Value + "</Value></Eq></Where><OrderBy><FieldRef Name='ID' Ascending='True' /></OrderBy></Query></View>"
                };
                ListItemCollection collListItem = list.GetItems(camlQuery);
                SPContext.Load(collListItem);
                SPContext.ExecuteQuery();
                List<Recommenders> RecommenderList = new List<Recommenders>();
                foreach (ListItem oListItem in collListItem)
                {
                    Recommenders recommenders = new Recommenders();
                    recommenders.Status = Convert.ToString(oListItem["DC_Status"]);
                    recommenders.Role = Convert.ToString(oListItem["StakeholderRole"]);
                    recommenders.ApprovedDate = Convert.ToString(oListItem["WFEndDate"]);
                    recommenders.AssignedDate = Convert.ToString(oListItem["WFStartDate"]);
                    recommenders.Email = Convert.ToString(oListItem["StakeholderEmailId"]);
                    recommenders.DesignationDepartment = Convert.ToString(oListItem["DesignationDepartment"]);
                    userid = GetUserDetails(SPContext, recommenders.Email, out username, objDeviationDetails.DeviationItemId);
                    recommenders.Name = username;
                    recommenders.ItemId = Convert.ToInt32(oListItem["ID"]);
                    recommenders.Comments = Convert.ToString(oListItem["DC_Comments"]);
                    if (recommenders.Role == "CREDIT_OPS" && recommenders.Status == "Approve")
                    {
                        objDeviationDetails.CreditOpsStatus = recommenders.Status;
                        objDeviationDetails.CreditOpsApprovedBy = recommenders.Name;
                        objDeviationDetails.CreditOpsApprovedDate = recommenders.ApprovedDate;
                        objDeviationDetails.CreditOpsEmailId = recommenders.Email;
                        objDeviationDetails.CreditOpsComments = recommenders.Comments;
                        objDeviationDetails.CreditOpsAssignedDate = recommenders.AssignedDate;
                    }
                    if (recommenders.Role == "CFO" && recommenders.Status == "Approve")
                    {
                        objDeviationDetails.CFOEndDate = recommenders.ApprovedDate;
                        objDeviationDetails.CFOApprovedBy = recommenders.Name;
                    }
                    if (recommenders.Role != "CREDIT_OPS" && recommenders.Role != "CFO" && recommenders.Role != "BusinessOps" && recommenders.Status == "Approve")
                    {
                        objDeviationDetails.recommendedByCount += 1;
                    }
                    RecommenderList.Add(recommenders);
                    if (recommenders.Status == "Reject")
                    {
                        objDeviationDetails.recommendedByCount = 0;
                        objDeviationDetails.CreditOpsStatus = "";
                        objDeviationDetails.CreditOpsApprovedBy = "";
                        objDeviationDetails.CreditOpsApprovedDate = "";
                        objDeviationDetails.CFOEndDate = "";
                        objDeviationDetails.CFOApprovedBy = "";
                    }
                }
                objDeviationDetails.RecommendersList = RecommenderList;
                GetFilteredRecommendedBy(objDeviationDetails);
            }
            catch (Exception ex)
            {
                logError("GetWorkflowStatus", ex.Message, "225", objDeviationDetails.DeviationItemId);
            }
        }

        private string GetUserDetails(ClientContext ctx, string emailId, out string username, string deviationId)
        {
            string userId = ""; username = "";
            try
            {
                if (emailId != "")
                {
                    var users = ctx.LoadQuery(ctx.Web.SiteUsers.Where(u => u.PrincipalType == PrincipalType.User && u.UserId.NameIdIssuer == "urn:federation:microsoftonline"));
                    ctx.ExecuteQuery();
                    string[] emails = emailId.Split(';');
                    int cntUserName = 0;
                    if (emails.Length == 1)
                    {
                        foreach (User usr in users)
                        {
                            if (usr.Email.ToLower() == emailId.ToLower())
                            {
                                userId = Convert.ToString(usr.Id);
                                if (usr.Title != "")
                                {
                                    string[] name = usr.Title.Split('/');
                                    if (cntUserName == 0)
                                    {
                                        username = name[0];
                                        cntUserName++;
                                    }
                                    else
                                    {
                                        username = username + ", " + name[0];
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        foreach (User usr in users)
                        {
                            for (int ecnt = 0; ecnt < emails.Length; ecnt++)
                            {
                                if (usr.Email.ToLower() == emails[ecnt].ToLower())
                                {
                                    userId = Convert.ToString(usr.Id);
                                    if (usr.Title != "")
                                    {
                                        string[] name = usr.Title.Split('/');
                                        if (cntUserName == 0)
                                        {
                                            username = name[0];
                                            cntUserName++;
                                        }
                                        else
                                        {
                                            username = username + ", " + name[0];
                                        }
                                    }
                                }
                            }
                        }
                    }

                }
            }
            catch (Exception ex)
            {
                logError("GetUserDetails", ex.Message, "1618", deviationId);
            }

            return userId;
        }

        public List<FileLink> GetDeviationNoteFileLink(DeviationDetailsProps objDeviationDetails, dynamic parsed)
        {
            List<FileLink> fileLinks = new List<FileLink>();
            try
            {
                var list = SPContext.Web.Lists.GetByTitle("DC_DeviationNoteFiles");
                CamlQuery camlQuery = new CamlQuery
                {
                    ViewXml = "<View><Query><Where><Eq><FieldRef Name='TransactionIqId' /><Value Type='Text'>" + parsed["DeviationID"].Value + "</Value></Eq></Where></Query></View>"
                };
                ListItemCollection collListItem = list.GetItems(camlQuery);
                SPContext.Load(collListItem);
                SPContext.ExecuteQuery();
                foreach (ListItem oListItem in collListItem)
                {
                    FileLink fileLink = new FileLink();
                    string file = Convert.ToString(oListItem["FileName"]);
                    fileLink.AttachmentType = Convert.ToString(oListItem["AttachmentType"]);
                    if (!string.IsNullOrEmpty(file) && file != "")
                    {
                        if (file.Contains('.'))
                        {
                            fileLink.Name = file.Substring(0, file.LastIndexOf('.'));
                        }
                        else
                        {
                            fileLink.Name = file;
                        }
                    }

                    FieldUrlValue urlValue = (FieldUrlValue)oListItem["_dlc_DocIdUrl"];
                    fileLink.Link = urlValue.Url.ToString();

                    fileLinks.Add(fileLink);
                }
            }
            catch (Exception ex)
            {
                logError("GetDeviationNoteFileLink", ex.Message, "1851", parsed["DeviationID"].Value);
            }
            return fileLinks;
        }

        public string GetPDFItemID(string FileRelativeURL)
        {
            string strFileItemId = "";
            Microsoft.SharePoint.Client.File file = SPContext.Web.GetFileByServerRelativeUrl(FileRelativeURL);
            SPContext.Load(file, f => f.ListItemAllFields);
            SPContext.ExecuteQuery();
            ListItem item = file.ListItemAllFields;

            if (item != null)
            {
                strFileItemId = Convert.ToString(item.Id);
            }
            return strFileItemId;
        }

        public void UpdateRefNo1(DeviationDetailsProps objDeviationDetails)
        {
            try
            {
                List oList = SPContext.Web.Lists.GetByTitle(Constants.DeviationNote);

                CamlQuery camlQuery = new CamlQuery()
                {
                    ViewXml = "<View><Query><Where><Eq><FieldRef Name='ID' /><Value Type='Counter'>" + objDeviationDetails.DeviationItemId + "</Value></Eq></Where></Query></View>"
                };

                ListItemCollection collListItem = oList.GetItems(camlQuery);
                SPContext.Load(collListItem);
                SPContext.ExecuteQuery();
                foreach (ListItem item in collListItem)
                {
                    item["Title"] = objDeviationDetails.DNRefNo;
                    FieldUrlValue fuv = new FieldUrlValue();
                    fuv.Url = objDeviationDetails.DNCollaborationDocUrl;
                    item["CollaborationDocUrl"] = fuv;
                    item["DeviationOkGivenOn"] = objDeviationDetails.CreditOpsApprovedDate;
                    item["TATValue"] = objDeviationDetails.TATValue;
                    //item["TATBurstReason"] = "TAT Burst";
                    string username = string.Empty;
                    string userid = GetUserDetails(SPContext, objDeviationDetails.CreditOpsEmailId, out username, objDeviationDetails.DeviationGUID);
                    if (!string.IsNullOrEmpty(userid) && userid != "")
                    {
                        item["DNDoneBy"] = userid;
                    }
                    if (!string.IsNullOrEmpty(username) && username != "")
                    {
                        item["DNDoneByFNameLName"] = username;
                    }
                    if (!string.IsNullOrEmpty(objDeviationDetails.CreditOpsComments) && objDeviationDetails.CreditOpsComments != "")
                    {
                        item["DNCreditOpsRemarks"] = objDeviationDetails.CreditOpsComments;
                    }
                    if (!string.IsNullOrEmpty(objDeviationDetails.FinancialYear) && objDeviationDetails.FinancialYear != "")
                    {
                        item["DNFinancialYear"] = objDeviationDetails.FinancialYear;
                    }
                    if (!string.IsNullOrEmpty(objDeviationDetails.CreditOpsAssignedDate) && objDeviationDetails.CreditOpsAssignedDate != "")
                    {
                        item["DNRequestReceivedDate"] = objDeviationDetails.CreditOpsAssignedDate;
                    }
                    item.Update();
                    SPContext.Load(item);
                    SPContext.ExecuteQuery();
                }
            }
            catch (Exception ex)
            {
                logError("UpdateRefNo", ex.Message, "1388", objDeviationDetails.DeviationGUID);
            }
        }

        public void UpdateRefNo(DeviationDetailsProps objDeviationDetails)
        {
            try
            {
                {
                    List oList = SPContext.Web.Lists.GetByTitle(Constants.DeviationNote);
                    ListItem item = oList.GetItemById(objDeviationDetails.DeviationItemId);
                    item["Title"] = objDeviationDetails.DNRefNo;
                    item["DeviationOkGivenOn"] = objDeviationDetails.CreditOpsApprovedDate;
                    item["TATValue"] = objDeviationDetails.TATValue;
                    //item["TATBurstReason"] = "TAT Burst Reason";
                    if (!string.IsNullOrEmpty(objDeviationDetails.CreditOpsAssignedDate) && objDeviationDetails.CreditOpsAssignedDate != "")
                    {
                        item["DNRequestReceivedDate"] = objDeviationDetails.CreditOpsAssignedDate;
                    }
                    if (!string.IsNullOrEmpty(objDeviationDetails.CreditOpsComments) && objDeviationDetails.CreditOpsComments != "")
                    {
                        item["DNCreditOpsRemarks"] = objDeviationDetails.CreditOpsComments;
                    }
                    if (!string.IsNullOrEmpty(objDeviationDetails.FinancialYear) && objDeviationDetails.FinancialYear != "")
                    {
                        item["DNFinancialYear"] = objDeviationDetails.FinancialYear;
                    }
                    FieldUrlValue fuv = new FieldUrlValue();
                    fuv.Url = objDeviationDetails.DNCollaborationDocUrl;
                    item["CollaborationDocUrl"] = fuv;
                    item.Update();
                    SPContext.ExecuteQuery();
                }

                {
                    List oList = SPContext.Web.Lists.GetByTitle(Constants.DeviationNote);
                    ListItem item = oList.GetItemById(objDeviationDetails.DeviationItemId);
                    
                    string username = string.Empty;
                    string userid = GetUserDetails(SPContext, objDeviationDetails.CreditOpsEmailId, out username, objDeviationDetails.DeviationGUID);
                    if (!string.IsNullOrEmpty(userid) && userid != "")
                    {
                        item["DNDoneBy"] = userid;
                    }
                    if (!string.IsNullOrEmpty(username) && username != "")
                    {
                        item["DNDoneByFNameLName"] = username;
                    }
                    item.Update();
                    SPContext.ExecuteQuery();
                }
            }
            catch (Exception ex)
            {
                logError("UpdateRefNo", ex.Message, "1388", objDeviationDetails.DeviationGUID);
            }
        }

        private void GetFilteredRecommendedBy(DeviationDetailsProps objDeviationDetails)
        {
            try
            {
                List<Recommenders> RecommendedByList = new List<Recommenders>();
                IEnumerable<Recommenders> itemCollection;
                if (objDeviationDetails.RecommendersList.Where(x => x.Status == "Reject").Count() > 0)
                {
                    int lastRejectionId = objDeviationDetails.RecommendersList.Where(x => x.Status == "Reject").Last().ItemId;
                    if (objDeviationDetails.RecommendersList.Where(x => x.Role == "IM" && x.ItemId > lastRejectionId).Count() > 0)
                    {
                        int IMId = objDeviationDetails.RecommendersList.Where(x => x.Role == "IM" && x.ItemId > lastRejectionId).Last().ItemId;
                        itemCollection = objDeviationDetails.RecommendersList.Where(x => x.ItemId >= IMId
                                && x.Role != "CFO" && x.Status == "Approve");
                    }
                    else
                    {
                        itemCollection = null;
                    }
                }
                else
                {
                    itemCollection = objDeviationDetails.RecommendersList.Where(x => x.Role != "CREDIT_OPS" && x.Role != "CFO" && x.Status == "Approve");
                }
                if (itemCollection != null)
                {
                    foreach (var item in itemCollection)
                    {
                        RecommendedByList.Add(item);
                    }
                }
                objDeviationDetails.RecommendersList = RecommendedByList;
            }
            catch (Exception ex)
            {
                logError("GetFilteredRecommendedBy", ex.Message, "1240", objDeviationDetails.DeviationGUID);
            }
        }

        public void GetStatusPageUrl()
        {
            string dbCtId = string.Empty;
            try
            {
                //string ctid = "0x01200200000DFEF8A15C0D4DA60B6805BEB3E260";
                dbCtId = common.CurrentEnvironment == "UAT" ? Constants.UATDBCTID : Constants.PRODDBCTID;
                objDeviationDetails.StatusPageUrl = common.CurrentSiteUrl+ "/Lists/DN_DiscussionBoard/Flat.aspx?RootFolder=" + objDeviationDetails.DashBoardUrl + "&FolderCTID=" + dbCtId + "&devid='" + objDeviationDetails.DeviationGUID + "'";
            }
            catch (Exception ex)
            {
                logError("GetStatusPageUrl", ex.Message, "99", objDeviationDetails.DeviationGUID);
            }
        }

        private void GetDNStakeHolders()
        {
            try
            {
                var list = SPContext.Web.Lists.GetByTitle(Constants.DN_StakeHolders);
                CamlQuery camlQuery = new CamlQuery()
                {
                    ViewXml = "<View><Query><Where><And><Eq><FieldRef Name='DeviationId' /><Value Type='Text'>" + objDeviationDetails.DeviationGUID + "</Value></Eq><Eq><FieldRef Name='StakeholderRole' /><Value Type='Text'>CREDIT_OPS</Value></Eq></And></Where></Query></View>"
                };
                ListItemCollection collListItem = list.GetItems(camlQuery);
                SPContext.Load(collListItem);
                SPContext.ExecuteQuery();
                foreach (ListItem oListItem in collListItem)
                {
                    objDeviationDetails.CreditOpsUser = Convert.ToString(oListItem["StakeholderEmailId"]);
                }   
            }
            catch (Exception ex)
            {
                logError("GetDNStakeHolders", ex.Message, "225", objDeviationDetails.DeviationItemId);
            }
        }

        public void GetTATValue(DeviationDetailsProps objDeviationDetails)
        {
            try
            {
                string[] requestDate = objDeviationDetails.CreditOpsAssignedDate.Split(' ')[0].Split('/');
                DateTime rDate = new DateTime(Convert.ToInt32(requestDate[2]), Convert.ToInt32(requestDate[0]), Convert.ToInt32(requestDate[1]));

                string[] approveDate = objDeviationDetails.CreditOpsApprovedDate.Split(' ')[0].Split('/');
                DateTime aDate = new DateTime(Convert.ToInt32(approveDate[2]), Convert.ToInt32(approveDate[0]), Convert.ToInt32(approveDate[1]));

                TimeSpan difference = aDate - rDate;
                int dayDiff = difference.Days;

                var list = SPContext.Web.Lists.GetByTitle("HolidayMaster");
                CamlQuery camlQuery = new CamlQuery
                {
                    ViewXml = "<View><Query><Where><And><Geq><FieldRef Name='Date' /><Value IncludeTimeValue='TRUE' Type='DateTime'>" + String.Format("{0:u}", rDate).Replace(' ', 'T') + "</Value></Geq><Leq><FieldRef Name='Date' /><Value IncludeTimeValue='TRUE' Type='DateTime'>" + String.Format("{0:u}", aDate).Replace(' ', 'T') + "</Value></Leq></And></Where></Query></View>"
                };
                ListItemCollection collListItem = list.GetItems(camlQuery);
                SPContext.Load(collListItem);
                SPContext.ExecuteQuery();
                List<HolidayMaster> HolidayMasterList = new List<HolidayMaster>();
                foreach (ListItem oListItem in collListItem)
                {
                    HolidayMaster hMast = new HolidayMaster();
                    hMast.Holidays = Convert.ToString(oListItem["Title"]);
                    hMast.HolidayDate = common.GetDateFieldInIST(Convert.ToString(oListItem["Date"]));
                    HolidayMasterList.Add(hMast);
                }
                objDeviationDetails.HolidayMasterList = HolidayMasterList;

                int matchingDays = 0;
                int startDay = Convert.ToInt32(requestDate[1]);
                for (int i = 0; i < dayDiff; i += 1)
                {
                    DateTime incrementalDate = rDate.AddDays(i);
                    if (incrementalDate.DayOfWeek.ToString() == "Saturday" || incrementalDate.DayOfWeek.ToString() == "Sunday")
                    {
                        matchingDays += 1;
                    }
                    else
                    {
                        for (int j = 0; j < HolidayMasterList.Count; j += 1)
                        {
                            if (common.GetDateFieldInIST(incrementalDate.ToString()) == HolidayMasterList[j].HolidayDate)
                            {
                                matchingDays += 1;
                                break;
                            }
                        }
                    }
                }
                objDeviationDetails.TATValue = dayDiff - matchingDays;
                if (objDeviationDetails.TATValue < 0)
                {
                    objDeviationDetails.TATValue = 0;
                }
            }
            catch (Exception ex)
            {
                logError("GetTATValue", ex.Message, "1518", objDeviationDetails.DeviationGUID);
            }
        }
    }
}

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ApprovalNoteAzure
{
    public class CreateHtml
    {
        public void CreateHtmlFile(DeviationDetailsProps objDeviationDetails, dynamic parsed)
        {
            common common = new common();
            DataAccess objDataAccess = new DataAccess();
            StringBuilder sbHtml = new StringBuilder();
            StringBuilder sbHtmlCollDoc = new StringBuilder();
            string lineNo = string.Empty;
            int imageid = 1;
            try
            {
                sbHtml.AppendLine("<!DOCTYPE html>");

                sbHtml.AppendLine("<html>");

                if (parsed["Note"] == "Deviation")
                {
                    sbHtml.AppendLine("<head><meta http-equiv='Content-Type'content='text/html; charset=utf-8'/><title>Changes / Modification in Proposal</title></head>");
                }
                else
                {
                    sbHtml.AppendLine("<head><meta http-equiv='Content-Type'content='text/html; charset=utf-8'/><title>Approval Note</title></head>");
                }
                sbHtml.AppendLine("<body style='font-family:calibri;font-size:20px;color:#000;width:100%;margin:0 auto' >");

                //sbHtml.AppendLine("<body style='font-family:calibri;font-size:20px;color:#000;width:1000px;margin:0 auto' >");

                try
                {
                    //log.Info("Logo Check at Azure D drive Started");
                    //if (File.Exists(@"D:\home\site\wwwroot\bin\piramal-capital-logo.png"))
                    //{
                    //    File.Delete(@"D:\home\site\wwwroot\bin\piramal-capital-logo.png");
                    //}
                    //log.Info("Logo Check at Azure D drive Completed");
                }
                catch (Exception ex)
                {
                    // log.Error("Failed to Delete File trans: " + name + "line " + lineno + " :" + ex.Message);
                }

                if (objDeviationDetails.LenderShortCode == "PHLFIN")
                {
                    imageid = 4;
                    objDeviationDetails.PiramalLogoName = "piramal-fininvest-logo";
                }
                else if (objDeviationDetails.LenderShortCode == "PSL")
                {
                    imageid = 3;
                    objDeviationDetails.PiramalLogoName = "piramal-securities-logo";
                }
                else if (objDeviationDetails.LenderShortCode == "PCHFL" || objDeviationDetails.LenderShortCode == "PFL")
                {
                    imageid = 2;
                    objDeviationDetails.PiramalLogoName = "piramal-capital-logo";
                }
                else
                {
                    imageid = 1;
                    objDeviationDetails.PiramalLogoName = "piramal-general-logo";
                }


                //if (System.IO.File.Exists(@"D:\home\site\wwwroot\bin\" + objDeviationDetails.PiramalLogoName + ".png"))
                //{
                //    //log.Error("The file exists.");
                //}
                //else
                //{

                //    objDataAccess.CreateLogonFromSP(imageid, objDeviationDetails.DealItemId);
                //}

                String origPath = objDataAccess.GetLogoInBase64(imageid, objDeviationDetails.DealItemId);

                #region Create File Content

                //CREATE HTML FILE CONTENT
                //sbHtml.AppendLine("<!DOCTYPE html>");

                sbHtml.AppendLine("<table cellpadding='0' cellpadding='0' border='0' width='100%' align='center' style='font-family: calibri; font-size: 20px; color: #000; '>");
                sbHtml.AppendLine("<tbody><tr>");
                sbHtml.AppendLine("<td style='width:33%'></td>");
                sbHtml.AppendLine("<td style='width:33%; text-align: center;'>" + Convert.ToString(objDeviationDetails.Borrowergroup) + "</td>");

                //sbHtml.AppendLine("<td style='width:33%;text-align: right;'><img src='D:\\home\\site\\wwwroot\\bin\\" + objDeviationDetails.PiramalLogoName + ".png'/></td>");
                sbHtml.AppendLine("<td style='width:33%;text-align: right;'><img src='" + origPath + "'/></td>");
                sbHtml.AppendLine("</tr>");
                sbHtml.AppendLine("<tr><td align='center' colspan='3' style='height: 2px; background: #622423;'></td></tr>");
                sbHtml.AppendLine("<tr><td align='center' colspan='3' style='height: 2px; background: #fff;'></td></tr>");
                sbHtml.AppendLine("<tr><td align='center' colspan='3' style='height: 7px; background: #622423;'></td></tr>");

                sbHtml.AppendLine("</tbody></table>");
                sbHtml.AppendLine("<br>");

                sbHtml.AppendLine("<table cellpadding='0' cellspacing='0' width='100%' align='center' style='page-break-inside: avoid !important;'>");

                //sbHtmlCollDoc.AppendLine(sbHtml.ToString());

                sbHtml.AppendLine("<tr style='page-break-inside: avoid !important;'>");
                sbHtml.AppendLine("<td align='left' valign='top' style='font-size:18px;'>");

                //objDeviationDetails.LenderShortCode = "PEL";
                lineNo = "163";
                if (objDeviationDetails.LenderShortCode == "PHLFIN" || objDeviationDetails.LenderShortCode == "PEL")
                {
                    if (!string.IsNullOrEmpty(objDeviationDetails.CFOEndDate))
                    {
                        sbHtml.AppendLine("Approved by: " + objDeviationDetails.CFOApprovedBy + "<br>");
                        sbHtml.AppendLine("Date: " + common.GetDateddmmyyyyhhmm(objDeviationDetails.CFOEndDate) + "<br>");
                    }
                }
                else
                {
                    if (!string.IsNullOrEmpty(objDeviationDetails.CFOEndDate))
                    {
                        sbHtml.AppendLine("<strong><span style='text-decoration: underline;'>Chief Financial Officer</span></strong><br>");
                        sbHtml.AppendLine("Approved by: " + objDeviationDetails.CFOApprovedBy + "<br>");
                        sbHtml.AppendLine("Date: " + common.GetDateddmmyyyyhhmm(objDeviationDetails.CFOEndDate) + "<br>");
                    }
                }
                sbHtml.AppendLine("</td>");
                sbHtml.AppendLine("<td align='right' valign='top'>");
                sbHtml.AppendLine( objDeviationDetails.LastModified);
                //sbHtml.AppendLine(DateTime.Now.ToString());
                sbHtml.AppendLine("</td>");
                sbHtml.AppendLine("</tr>");

                //sbHtmlCollDoc.AppendLine(sbHtml.ToString());

                sbHtml.AppendLine("<tr style='page-break-inside: avoid !important;'>");
                sbHtml.AppendLine("<td colspan='2' align='right' style='line-height: 25px'>");
                sbHtml.AppendLine("<strong>Ref No/UCIC</strong> : " + objDeviationDetails.UCIC + "<br>");
                //sbHtml.AppendLine("<strong>Facility ID</strong> : " + "dCMainList.FacilityID" + "<br>");
                if(parsed["Note"] == "Deviation")
                {
                    sbHtml.AppendLine("<strong>DN No.</strong> : ");
                }
                else
                {
                    sbHtml.AppendLine("<strong>AN No.</strong> : ");
                }
                
                sbHtml.AppendLine(objDeviationDetails.DeviationTitle + "</td></tr>");

                sbHtml.AppendLine("<tr style='page-break-inside: avoid !important;'>");
                if(parsed["Note"] == "Deviation")
                {
                    sbHtml.AppendLine("<td align='center' colspan='2' style='font-weight:bold'>Changes / Modification in Proposal</td>");
                }
                else
                {
                    sbHtml.AppendLine("<td align='center' colspan='2' style='font-weight:bold'>Approval Note</td>");
                }
                sbHtml.AppendLine("</tr>");
                sbHtml.AppendLine("<tr style='page-break-inside: avoid !important;'>");
                sbHtml.AppendLine("<td align='center' colspan='2' style='padding: 0 0 5px;'>"+objDeviationDetails.NatureOfDeviation+"</td>");
                sbHtml.AppendLine("</tr>");
                sbHtml.AppendLine("<tr style='page-break-inside: avoid !important;'>");
                sbHtml.AppendLine("<td align='left' valign='top' style='font-size:18px;'>");
                //sbHtml.AppendLine("<br>");
                // sbHtml.AppendLine("Transaction Summary<br>");
                //sbHtml.AppendLine("<strong><span style='text-decoration: underline;'>Transaction Summary</span></strong><br>");
                sbHtml.AppendLine("</td>");
                sbHtml.AppendLine("</tr>");

                sbHtml.AppendLine("<tr style='page-break-inside: avoid !important;'>");
                sbHtml.AppendLine("<td style='text-decoration: underline; padding:5px 0 0 0; font-size: 20px' width='30%'><strong>Transaction Summary:</strong></td>");
                sbHtml.AppendLine("<td width='70%'></td>");
                sbHtml.AppendLine("</tr>");
                sbHtml.AppendLine("</table>");


                //sbHtml.AppendLine("<table cellpadding='0' cellspacing='0' width='100%' align='center' style='font-size: 18px; margin-bottom: 5%;'>");

                //sbHtml.AppendLine("<tr style='page-break-inside: avoid !important;'>");
                //sbHtml.AppendLine("<td colspan='2'>");
                //sbHtml.AppendLine("<table cellpadding='0' cellspacing='0' width='100%' align='center' style='font-size: 18px; margin-bottom: 5%; border-collapse: collapse;'>");
                sbHtml.AppendLine("<table cellpadding='0' cellspacing='0' width='100%' align='center' style='font-size: 18px; margin-bottom: 3%; border-collapse: collapse; page-break-after:always;'>");

                sbHtml.AppendLine("<tr style='page-break-inside: avoid !important;'>");
                sbHtml.AppendLine("<td style='background: #efe9e9; padding: 10px; border:1px solid #333;' width='30%'>Borrower</td>");
                sbHtml.AppendLine("<td style='padding: 10px; border:1px solid #333;' width='70%'>" + objDeviationDetails.Borrower + "</td>");
                sbHtml.AppendLine("</tr>");

                sbHtml.AppendLine("<tr style='page-break-inside: avoid !important;'>");
                sbHtml.AppendLine("<td style='background: #efe9e9; padding: 10px; border:1px solid #333;' width='30%'>Instrument</td>");
                sbHtml.AppendLine("<td style='padding: 10px; border:1px solid #333;' width='70%'>" + objDeviationDetails.ExposureType + "</td>");
                sbHtml.AppendLine("</tr>");

                sbHtml.AppendLine("<tr style='page-break-inside: avoid !important;'>");
                sbHtml.AppendLine("<td style='background: #efe9e9; padding: 10px; border:1px solid #333;' width='30%'>Lender</td>");
                sbHtml.AppendLine("<td style='padding: 10px; border:1px solid #333;' width='70%'>" + objDeviationDetails.Lender + "</td>");
                sbHtml.AppendLine("</tr>");

                sbHtml.AppendLine("<tr style='page-break-inside: avoid !important;'>");
                sbHtml.AppendLine("<td style='background: #efe9e9; padding: 10px; border:1px solid #333;' width='30%'>Sanctioned Amount</td>");
                //sbHtml.AppendLine("<td style='padding: 10px; border:1px solid #333;' width='70%'>INR " + objDeviationDetails.SanctionedAmount + " Crs.</td>");
                sbHtml.AppendLine("<td style='padding: 10px; border:1px solid #333;' width='70%'>" + objDeviationDetails.SanctionedAmount + "</td>");
                sbHtml.AppendLine("</tr>");

                sbHtml.AppendLine("<tr style='page-break-inside: avoid !important;'>");
                sbHtml.AppendLine("<td style='background: #efe9e9; padding: 10px; border:1px solid #333;' width='30%'>Date of Sanction</td>");
                sbHtml.AppendLine("<td style='padding: 10px; border:1px solid #333;' width='70%'>" + objDeviationDetails.DateOfSanction + "</td>");
                sbHtml.AppendLine("</tr>");

                sbHtml.AppendLine("<tr style='page-break-inside: avoid !important;'>");
                sbHtml.AppendLine("<td style='background: #efe9e9; padding: 10px; border:1px solid #333;' width='30%'>Purpose</td>");
                sbHtml.AppendLine("<td style='padding: 10px; border:1px solid #333;' width='70%'>" + objDeviationDetails.Purpose + "</td>");
                sbHtml.AppendLine("</tr>");

                sbHtml.AppendLine("<tr style='page-break-inside: avoid !important;'>");
                sbHtml.AppendLine("<td style='background: #efe9e9; padding: 10px; border:1px solid #333;' width='30%'>Tenure</td>");
                sbHtml.AppendLine("<td style='padding: 10px; border:1px solid #333;' width='70%'>" + objDeviationDetails.Tenure + "</td>");
                sbHtml.AppendLine("</tr>");

                sbHtml.AppendLine("<tr style='page-break-inside: avoid !important;'>");
                sbHtml.AppendLine("<td style='background: #efe9e9; padding: 10px; border:1px solid #333;' width='30%'>Coupon Rate</td>");
                sbHtml.AppendLine("<td style='padding: 10px; border:1px solid #333;' width='70%'>" + objDeviationDetails.CouponRate + "</td>");
                sbHtml.AppendLine("</tr>");

                sbHtml.AppendLine("<tr style='page-break-inside: avoid !important;'>");
                sbHtml.AppendLine("<td style='background: #efe9e9; padding: 10px; border:1px solid #333;' width='30%'>IRR</td>");
                sbHtml.AppendLine("<td style='padding: 10px; border:1px solid #333;' width='70%'>" + objDeviationDetails.IRR_CoupounRate + "</td>");
                sbHtml.AppendLine("</tr>");

                sbHtml.AppendLine("<tr style='page-break-inside: avoid !important;'>");
                sbHtml.AppendLine("<td style='background: #efe9e9; padding: 10px; border:1px solid #333;'  width='30%'>Security Details</td>");
                sbHtml.AppendLine("<td style='border:1px solid #333;'  width='70%'>");

                sbHtml.AppendLine("<table cellpadding='0' cellspacing='0' width='100%' align='center' style='font-size: 18px;  border-collapse: collapse;'>");


                sbHtml.AppendLine("<tr style='page-break-inside: avoid !important;'>");
                //sbHtml.AppendLine("<td style='padding: 10px; ' width='30%'><b>Primary Security :</b></td>");
                //sbHtml.AppendLine("<td style='padding: 10px; ' width='70%'>" + objDeviationDetails.SecurityDetails_Primary + "</td>");
                sbHtml.AppendLine("<td style='padding: 10px' width='100%'><b>Primary Security :</b></br>" + objDeviationDetails.SecurityDetails_Primary + "</td>");
                sbHtml.AppendLine("</tr>");

                sbHtml.AppendLine("<tr style='page-break-inside: avoid !important;'>");
                //sbHtml.AppendLine("<td style='padding: 10px; ' width='30%'><b>Additional Security :</b></td>");
                //sbHtml.AppendLine("<td style='padding: 10px; ' width='70%'>" + objDeviationDetails.SecurityDetails_Additional + "</td>");
                sbHtml.AppendLine("<td style='padding: 0 10px 0 10px;' width='100%'><b>Additional Security :</b></br>" + objDeviationDetails.SecurityDetails_Additional + "</td>");
                sbHtml.AppendLine("</tr>");

                sbHtml.AppendLine("<tr style='page-break-inside: avoid !important;'>");
                //sbHtml.AppendLine("<td style='padding: 10px; ' width='30%'><b>Collateral Security :</b></td>");
                //sbHtml.AppendLine("<td style='padding: 10px; ' width='70%'>" + objDeviationDetails.SecurityDetails_Collateral + "</td>");
                sbHtml.AppendLine("<td style='padding: 10px' width='100%'><b>Collateral Security :</b></br>" + objDeviationDetails.SecurityDetails_Collateral + "</td>");
                sbHtml.AppendLine("</tr>");

                sbHtml.AppendLine("</table>");

                sbHtml.AppendLine("</td>");
                sbHtml.AppendLine("</tr>");

                sbHtml.AppendLine("<tr style='page-break-inside: avoid !important;'>");
                sbHtml.AppendLine("<td style='background: #efe9e9; padding: 10px; border:1px solid #333;' width='30%'>Cash Cover</td>");
                sbHtml.AppendLine("<td style='padding: 10px; border:1px solid #333;' width='70%'>" + objDeviationDetails.Cover + "</td>");
                sbHtml.AppendLine("</tr>");

                sbHtml.AppendLine("<tr style='page-break-inside: avoid !important;'>");
                sbHtml.AppendLine("<td style='background: #efe9e9; padding: 10px; border:1px solid #333;' width='30%'>Security Cover</td>");
                sbHtml.AppendLine("<td style='padding: 10px; border:1px solid #333;' width='70%'>" + objDeviationDetails.SecurityCover + "</td>");
                sbHtml.AppendLine("</tr>");

                sbHtml.AppendLine("<tr style='page-break-inside: avoid !important;'>");
                sbHtml.AppendLine("<td style='background: #efe9e9; padding: 10px; border:1px solid #333;' width='30%'>Processing Fee</td>");
                sbHtml.AppendLine("<td style='padding: 10px; border:1px solid #333;' width='70%'>" + objDeviationDetails.ProcessingFee + "</td>");
                sbHtml.AppendLine("</tr>");

                sbHtml.AppendLine("<tr style='page-break-inside: avoid !important;'>");
                sbHtml.AppendLine("<td style='background: #efe9e9; padding: 10px; border:1px solid #333;' width='30%'>Repayment Schedule</td>");
                sbHtml.AppendLine("<td style='padding: 10px; border:1px solid #333;' width='70%'>" + objDeviationDetails.PrincipalServicing + "</td>");
                sbHtml.AppendLine("</tr>");

                sbHtml.AppendLine("<tr style='page-break-inside: avoid !important;'>");
                sbHtml.AppendLine("<td style='background: #efe9e9; padding: 10px; border:1px solid #333;' width='30%'>Prepayment Till Date</td>");
                sbHtml.AppendLine("<td style='padding: 10px; border:1px solid #333;' width='70%'>" + objDeviationDetails.PrepaymentTillDate + "</td>");
                sbHtml.AppendLine("</tr>");

                sbHtml.AppendLine("<tr style='page-break-inside: avoid !important;'>");
                sbHtml.AppendLine("<td style='background: #efe9e9; padding: 10px; border:1px solid #333;' width='30%'>Conduct Of Account</td>");
                sbHtml.AppendLine("<td style='padding: 10px; border:1px solid #333;' width='70%'>" + objDeviationDetails.ConductOfAccount + "</td>");
                sbHtml.AppendLine("</tr>");

                sbHtml.AppendLine("</br>");

                lineNo = "312";

                sbHtml.AppendLine("</table>");

                sbHtml.AppendLine("<table cellpadding='0' cellspacing='0' width='99.92%' align='center' style='font-size: 18px; margin-bottom: 3%; border-collapse: collapse; border-style: none;'>");
                sbHtml.AppendLine("<tr style='page-break-inside: avoid !important;'>");
                if(parsed["Note"] == "Deviation")
                {
                    sbHtml.AppendLine("<td style='text-decoration: underline; padding: 15px 0; font-size: 20px' colspan='4' width='100%'><strong>Modifications / Deviations Proposed in Terms & Conditions:</strong></td>");
                }
                else
                {
                    sbHtml.AppendLine("<td style='text-decoration: underline; padding: 15px 0; font-size: 20px' colspan='4' width='100%'><strong>Approval sought for:</strong></td>");
                }

                sbHtml.AppendLine("</tr>");
                sbHtml.AppendLine("<tr style='page-break-inside: avoid !important;'>");
                sbHtml.AppendLine("<td style='background: #efe9e9; padding: 10px; border:1px solid #333;' width='25%'><strong>Original term(s) sanctioned by the IC</strong></td>");
                if(parsed["Note"] == "Deviation")
                {
                    sbHtml.AppendLine("<td style='background: #efe9e9; padding: 10px; border:1px solid #333;' width='25%'><strong>Modification(s) / deviation(s) proposed</strong></td>");
                }
                else
                {
                    sbHtml.AppendLine("<td style='background: #efe9e9; padding: 10px; border:1px solid #333;' width='25%'><strong>Approval sought for</strong></td>");
                }
                
                sbHtml.AppendLine("<td style='background: #efe9e9; padding: 10px; border:1px solid #333;' width='25%'><strong>Impact</strong></td>");
                sbHtml.AppendLine("<td style='background: #efe9e9; padding: 10px; border:1px solid #333;' width='25%'><strong>Justification</strong></td>");
                sbHtml.AppendLine("</tr>");


                if (objDeviationDetails.RTBData != null)
                {
                    for (int i = 0; i < objDeviationDetails.RTBData.Count; i++)
                    {
                        sbHtml.AppendLine("<tr style='page-break-inside: avoid !important;'>");
                        sbHtml.AppendLine("<td style='padding: 10px; border:1px solid #333; text-align:left; vertical-align:top' width='25%'>" + objDeviationDetails.RTBData[i].Originalterms + "</td>");
                        sbHtml.AppendLine("<td style='padding: 10px; border:1px solid #333; text-align:left; vertical-align:top' width='25%'>" + objDeviationDetails.RTBData[i].ProposedDeviation + "</td>");
                        sbHtml.AppendLine("<td style='padding: 10px; border:1px solid #333; text-align:left; vertical-align:top' width='25%'>" + objDeviationDetails.RTBData[i].Impact + "</td>");
                        sbHtml.AppendLine("<td style='padding: 10px; border:1px solid #333; text-align:left; vertical-align:top' width='25%'>" + objDeviationDetails.RTBData[i].Justification + "</td>");
                        sbHtml.AppendLine("</tr>");
                    }
                }

                lineNo = "343";
                sbHtml.AppendLine("</table>");

                sbHtmlCollDoc.AppendLine(sbHtml.ToString());
                lineNo = "344";

                sbHtml.AppendLine("</br>");

                string status = "";
                if (objDeviationDetails.CreditOpsStatus == "Approve")
                {
                    status = "OK";
                }
                else
                {
                    status = "";
                }

                sbHtml.AppendLine("<p><strong>Reviewer/Credit Ops (" + objDeviationDetails.CreditOpsApprovedBy + ") Remarks: " + status + "</strong></p>");

                int workflowcount = objDeviationDetails.RecommendersList.Count;

                sbHtml.AppendLine("<table cellspacing='0' cellpadding='0' width='100%' style='page-break-inside: avoid !important;'>");
                sbHtml.AppendLine("<tr style='page-break-inside: avoid !important;'>");
                sbHtml.AppendLine("<td style='padding:15px 0; font-size:20px; text-decoration:underline;'><strong>Recommendation from:</strong></td>");
                sbHtml.AppendLine("</tr>");
                //common common = new common();
                int index = 0, tableCount = 0, startIndex = 0;
                if (workflowcount < 4)
                {
                    tableCount = workflowcount;
                }
                else
                {
                    tableCount = 4;
                }
                double rowCount = (double)objDeviationDetails.recommendedByCount / 4;
                while (index < Math.Ceiling(rowCount))
                {
                    string action = "recommendation";
                    for (int i = 0; i < 5; i += 1)
                    {
                        switch (action)
                        {
                            case "recommendation":
                                sbHtml.AppendLine("<tr>");
                                for (int j = startIndex; j < tableCount; j += 1)
                                {
                                    //if (objDeviationDetails.RecommendersList[j].Status == "Approve")
                                    {
                                        sbHtml.AppendLine("<td width='25%' style='page-break-inside: avoid !important;padding:10px 10px 15px 10px; border:1px solid #333;background: #efe9e9;'>Recommended By</td>");
                                    }
                                    //else
                                    //{
                                    //    if (tableCount < workflowcount)
                                    //    {
                                    //        tableCount += 1;
                                    //        //break;
                                    //    }
                                    //}
                                }
                                sbHtml.AppendLine("</tr>");
                                action = "name";
                                break;
                            case "name":
                                sbHtml.AppendLine("<tr>");
                                for (int j = startIndex; j < tableCount; j += 1)
                                {
                                    //if (objDeviationDetails.RecommendersList[j].Status == "Approve")
                                    {
                                        sbHtml.AppendLine("<td width='25%' style='page-break-inside: avoid !important;padding:10px 10px 15px 10px; border:1px solid #333;'>" + objDeviationDetails.RecommendersList[j].Name + "</td>");
                                    }
                                }
                                sbHtml.AppendLine("</tr>");
                                action = "designationDepartment";
                                break;
                            case "designationDepartment":
                                sbHtml.AppendLine("<tr>");
                                for (int j = startIndex; j < tableCount; j += 1)
                                {
                                    //if (objDeviationDetails.RecommendersList[j].Status == "Approve")
                                    {
                                        sbHtml.AppendLine("<td width='25%' style='page-break-inside: avoid !important;padding:10px 10px 15px 10px; border:1px solid #333;'>" + objDeviationDetails.RecommendersList[j].DesignationDepartment + "</td>");
                                    }
                                }
                                sbHtml.AppendLine("</tr>");
                                action = "approvedDate";    
                                break;
                            case "approvedDate":
                                sbHtml.AppendLine("<tr>");
                                for (int j = startIndex; j < tableCount; j += 1)
                                {
                                    //if (objDeviationDetails.RecommendersList[j].Status == "Approve")
                                    {
                                        sbHtml.AppendLine("<td width='25%' style='page-break-inside: avoid !important;padding:10px 10px 15px 10px; border:1px solid #333;'>" + common.GetDateddmmyyyyhhmm(objDeviationDetails.RecommendersList[j].ApprovedDate) + "</td>");
                                    }
                                }
                                sbHtml.AppendLine("</tr>");
                                action = "role";
                                break;
                                case "role":
                                    sbHtml.AppendLine("<tr>");
                                    for (int j = startIndex; j < tableCount; j += 1)
                                    {
                                        //if (objDeviationDetails.RecommendersList[j].Status == "Approve")
                                        {
                                            sbHtml.AppendLine("<td width='25%' style='page-break-inside: avoid !important;padding:10px 10px 15px 10px; border:1px solid #333; text-align:left; vertical-align:top'>" + objDeviationDetails.RecommendersList[j].Comments + "</td>");
                                        }
                                    }
                                    sbHtml.AppendLine("</tr>");
                                    break;
                        }
                    }
                    index += 1;
                    startIndex = tableCount;
                    if (workflowcount < (4 + startIndex))
                    {
                        tableCount = workflowcount;
                    }
                    else
                    {
                        tableCount = 4 + startIndex;
                    }
                    if (index < Math.Ceiling(rowCount))
                    {
                        sbHtml.AppendLine("<tr><td colspan='4'>.</td></tr>");
                    }
                }


                //sbHtml.AppendLine("<div style = 'text-align:center; font-size:14px;'> Page <span class='page'></span> of &nbsp;<span class='topage'></span><div>");
                sbHtml.AppendLine("</table>");

                sbHtml.AppendLine("</br>");
                if (objDeviationDetails.Annexure != null && objDeviationDetails.Annexure != "")
                {
                    sbHtml.AppendLine("<table cellspacing='0' cellpadding='0' width='100%' style='font-size: 18px; page-break-inside: avoid !important; border-collapse: collapse;'>");
                    sbHtml.AppendLine("<tr>");
                    sbHtml.AppendLine("<td width='25%' style='page-break-inside: avoid !important;padding:10px 10px 15px 10px; border:1px solid #333;background: #efe9e9;'>Annexure</td>");
                    sbHtml.AppendLine("</tr>");
                    sbHtml.AppendLine("<tr style='page-break-inside: avoid !important;'>");
                    sbHtml.AppendLine("<td width='25%' style='page-break-inside: avoid !important;padding:10px 10px 15px 10px; border:1px solid #333;'>" + objDeviationDetails.Annexure + "</td>");
                    sbHtml.AppendLine("</tr>");
                    sbHtml.AppendLine("</table>");

                    sbHtmlCollDoc.AppendLine("</br><table cellspacing='0' cellpadding='0' width='100%' style='font-size: 18px; page-break-inside: avoid !important; border-collapse: collapse;'>");
                    sbHtmlCollDoc.AppendLine("<tr>");
                    sbHtmlCollDoc.AppendLine("<td width='25%' style='page-break-inside: avoid !important;padding:10px 10px 15px 10px; border:1px solid #333;background: #efe9e9;'>Annexure</td>");
                    sbHtmlCollDoc.AppendLine("</tr>");
                    sbHtmlCollDoc.AppendLine("<tr style='page-break-inside: avoid !important;'>");
                    sbHtmlCollDoc.AppendLine("<td width='25%' style='page-break-inside: avoid !important;padding:10px 10px 15px 10px; border:1px solid #333;'>" + objDeviationDetails.Annexure + "</td>");
                    sbHtmlCollDoc.AppendLine("</tr>");
                    sbHtmlCollDoc.AppendLine("</table>");
                    sbHtmlCollDoc.AppendLine("</body></html>");

                }
                sbHtml.AppendLine("</body>");
                sbHtml.AppendLine("</html>");

                if (!objDeviationDetails.DeviationTitle.Contains('/') && !string.IsNullOrEmpty(objDeviationDetails.CreditOpsApprovedDate))
                {
                    objDataAccess.GetTATValue(objDeviationDetails);
                    objDeviationDetails.FinancialYear = common.GetCurrentFinancialYear(DateTime.Parse(objDeviationDetails.CreditOpsApprovedDate));
                    //string FinYear = "";
                    string Note = parsed["Note"] == "Deviation" ? "DN" : "AN";
                    objDeviationDetails.DNRefNo = objDeviationDetails.LenderShortCode + "/" + objDeviationDetails.BusinessSegment + "/" + Note + "/" +
                        objDeviationDetails.RegionShortCode + "/" + objDeviationDetails.DealID + "/" + objDeviationDetails.FinancialYear + "/" + objDeviationDetails.DeviationItemId;
                    common.CreateCollaborationWordFile(objDeviationDetails, sbHtmlCollDoc);
                }

                #endregion
                lineNo = "402";
                HTMLtoPDF generateHtmlToPDF = new HTMLtoPDF();

                //CALL HTML TO PDF FUNCTION
                //objDeviationDetails.PDFUrl = generateHtmlToPDF.HTML_To_PDF(common.CurrentSiteUrl, "DeviationNote_Created_PDF", objDeviationDetails.PDFName + ".pdf", sbHtml.ToString(), "", "", parsed["DeviationID"].Value);

                string strFileRelativeURL = generateHtmlToPDF.HTML_To_PDF(common.CurrentSiteUrl, Constants.PDFFile_DocumentLibrary, objDeviationDetails.PDFName + ".pdf", sbHtml.ToString(), "", "", parsed["DeviationID"].Value);

                if (parsed["UpdatePDF"] == null)
                {
                    objDeviationDetails.PDFFileItemID = objDataAccess.GetPDFItemID(strFileRelativeURL);
                }
                objDeviationDetails.PDFUrl = Constants.BaseSiteURL + strFileRelativeURL;
                lineNo = "407";
            }
            catch (Exception ex)
            {
                objDataAccess.logError("CreateHtmlFile", ex.Message, lineNo, objDeviationDetails.DeviationGUID);
                //objDataAccess.logError("CreateHtmlFile", objDeviationDetails.InsideCreateWordFile, objDeviationDetails.InsideSaveHtmlFile, objDeviationDetails.InsideSaveDocInList);
            }
        }

        public void GenerateApprovalEmailBody(DeviationDetailsProps objDeviationDetails, dynamic parsed)
        {
            StringBuilder sbText = new StringBuilder();
            try
            {
                List<FileLink> lstFile = objDeviationDetails.DeviationNoteFileList;

                sbText.AppendLine("#### Dear User,");
                if(parsed["Note"] == "Deviation")
                {
                    sbText.AppendLine("\nApproval request for the below mentioned Deviation Note Request has been rejected");
                }
                else
                {
                    sbText.AppendLine("\nApproval request for the below mentioned Approval Note Request has been rejected");
                }
                sbText.AppendLine("\nDeal Name : **" + objDeviationDetails.DealName + "**");
                sbText.AppendLine("\nRef No : **" + objDeviationDetails.DeviationTitle + "**");
                if (lstFile != null)
                {
                    if (lstFile.Count > 0)
                    {
                        sbText.AppendLine("\n**Attachments :**\n");
                        for (int i = 0; i < lstFile.Count; i++)
                        {
                            sbText.AppendLine("\nAttachment " + (i + 1) + " : [" + Convert.ToString(lstFile[i].Name) + "](" + Convert.ToString(lstFile[i].Link) + ")");
                        }
                    }
                    else
                    {
                        sbText.AppendLine("\n**Attachments :** NA\n");
                    }
                }
                else
                {
                    sbText.AppendLine("\n**Attachments :** NA\n");
                }
                sbText.AppendLine("\nRejection Comments from Previous Stage : (RejectionComments)");
                sbText.AppendLine("\nKindly review the comments and reinitiate the approval for disbursement");
                if (!string.IsNullOrEmpty(objDeviationDetails.CreditOpsApprovedDate))
                {
                    if(parsed["Note"] == "Deviation")
                    {
                        sbText.AppendLine("\n\n[View Deviation Note](" + Uri.EscapeUriString(objDeviationDetails.PDFUrl) + ") | [View Status](" + Uri.EscapeUriString(objDeviationDetails.StatusPageUrl) + ") | [Queries](" + Uri.EscapeUriString(objDeviationDetails.DNCollaborationDocUrl) + ")");
                    }
                    else
                    {
                        sbText.AppendLine("\n\n[View Approval Note](" + Uri.EscapeUriString(objDeviationDetails.PDFUrl) + ") | [View Status](" + Uri.EscapeUriString(objDeviationDetails.StatusPageUrl) + ") | [Queries](" + Uri.EscapeUriString(objDeviationDetails.DNCollaborationDocUrl) + ")");
                    }
                }
                else
                {
                    if (parsed["Note"] == "Deviation")
                    {
                        sbText.AppendLine("\n\n[View Deviation Note](" + Uri.EscapeUriString(objDeviationDetails.PDFUrl) + ") | [View Status](" + Uri.EscapeUriString(objDeviationDetails.StatusPageUrl) + ")");
                    }
                    else
                    {
                        sbText.AppendLine("\n\n[View Approval Note](" + Uri.EscapeUriString(objDeviationDetails.PDFUrl) + ") | [View Status](" + Uri.EscapeUriString(objDeviationDetails.StatusPageUrl) + ")");
                    }
                }
                sbText.AppendLine("\nThis is a system generated email, kindly do not respond to this email.For queries please drop an email to [pchfl.itws@piramal.com]().");
                sbText.AppendLine("\nThanks & Regards,");
                sbText.AppendLine("\nIT SUPPORT");
                objDeviationDetails.RejectedMailBody = sbText.ToString();
                
                sbText.Clear();
                sbText.AppendLine("#### Dear User,");
                if(parsed["Note"] == "Deviation")
                {
                    sbText.AppendLine("\nApproval request for the below Deviation Note Request has been deleted / cancelled.");
                }
                else
                {
                    sbText.AppendLine("\nApproval request for the below Approval Note Request has been deleted / cancelled.");
                }
                sbText.AppendLine("\nDeal Name : **" + objDeviationDetails.DealName + "**");
                sbText.AppendLine("\nRef No : **" + objDeviationDetails.DeviationTitle + "**");
                if (lstFile != null)
                {
                    if (lstFile.Count > 0)
                    {
                        sbText.AppendLine("\n**Attachments :**\n");
                        for (int i = 0; i < lstFile.Count; i++)
                        {
                            sbText.AppendLine("\nAttachment " + (i + 1) + " : [" + Convert.ToString(lstFile[i].Name) + "](" + Convert.ToString(lstFile[i].Link) + ")");
                        }
                    }
                    else
                    {
                        sbText.AppendLine("\n**Attachments :** NA\n");
                    }
                }
                else
                {
                    sbText.AppendLine("\n**Attachments :** NA\n");
                }
                sbText.AppendLine("\nThe approval workflow has been terminated");
                //sbText.AppendLine("\n\n[View Deviation Note](" + Uri.EscapeUriString(objDeviationDetails.PDFUrl) + ") | [View Status](" + Uri.EscapeUriString(objDeviationDetails.StatusPageUrl) + ")");
                if (!string.IsNullOrEmpty(objDeviationDetails.CreditOpsApprovedDate))
                {
                    if(parsed["Note"] == "Deviation")
                    {
                        sbText.AppendLine("\n\n[View Deviation Note](" + Uri.EscapeUriString(objDeviationDetails.PDFUrl) + ") | [View Status](" + Uri.EscapeUriString(objDeviationDetails.StatusPageUrl) + ") | [Queries](" + Uri.EscapeUriString(objDeviationDetails.DNCollaborationDocUrl) + ")");
                    }
                    else
                    {
                        sbText.AppendLine("\n\n[View Approval Note](" + Uri.EscapeUriString(objDeviationDetails.PDFUrl) + ") | [View Status](" + Uri.EscapeUriString(objDeviationDetails.StatusPageUrl) + ") | [Queries](" + Uri.EscapeUriString(objDeviationDetails.DNCollaborationDocUrl) + ")");
                    }
                }
                else
                {
                    if (parsed["Note"] == "Deviation")
                    {
                        sbText.AppendLine("\n\n[View Deviation Note](" + Uri.EscapeUriString(objDeviationDetails.PDFUrl) + ") | [View Status](" + Uri.EscapeUriString(objDeviationDetails.StatusPageUrl) + ")");
                    }
                    else
                    {
                        sbText.AppendLine("\n\n[View Approval Note](" + Uri.EscapeUriString(objDeviationDetails.PDFUrl) + ") | [View Status](" + Uri.EscapeUriString(objDeviationDetails.StatusPageUrl) + ")");
                    }
                }
                sbText.AppendLine("\nThis is a system generated email, kindly do not respond to this email.For queries please drop an email to [pchfl.itws@piramal.com]().");
                sbText.AppendLine("\nThanks & Regards,");
                sbText.AppendLine("\nIT SUPPORT");
                objDeviationDetails.DeletedMailBody = sbText.ToString();
                
                sbText.Clear();
                sbText.AppendLine("#### Dear User,");
                if (parsed["Note"] == "Deviation")
                {
                    sbText.AppendLine("\nKindly review and approve the deviation note request for");
                }
                else
                {
                    sbText.AppendLine("\nKindly review and approve the approval note request for");
                }
                sbText.AppendLine("\nDeal Name : **" + objDeviationDetails.DealName + "**");
                sbText.AppendLine("\nRef No : **" + objDeviationDetails.DeviationTitle + "**");
                if (lstFile != null)
                {
                    if (lstFile.Count > 0)
                    {
                        sbText.AppendLine("\n**Attachments :**\n");
                        for (int i = 0; i < lstFile.Count; i++)
                        {
                            sbText.AppendLine("\nAttachment " + (i + 1) + " : [" + Convert.ToString(lstFile[i].Name) + "](" + Convert.ToString(lstFile[i].Link) + ")");
                        }
                    }
                    else
                    {
                        sbText.AppendLine("\n**Attachments :** NA\n");
                    }
                }
                else
                {
                    sbText.AppendLine("\n**Attachments :** NA\n");
                }
                //if (parsed["Role"].Count != 0 && parsed["Role"][0].Value == "CFO")
                //{
                //    sbText.AppendLine("\nInvestment Manager's Comments : " + dCMainList.APComments);
                //    sbText.AppendLine("\nPrincipal's Comments : " + dCMainList.PrincipalComments);
                //}
                //sbText.AppendLine("\n\n[View Deviation Note](" + Uri.EscapeUriString(objDeviationDetails.PDFUrl) + ") | [View Status](" + Uri.EscapeUriString(objDeviationDetails.StatusPageUrl) + ")");
                if (!string.IsNullOrEmpty(objDeviationDetails.CreditOpsApprovedDate))
                {
                    if (parsed["Note"] == "Deviation")
                    {
                        sbText.AppendLine("\n\n[View Deviation Note](" + Uri.EscapeUriString(objDeviationDetails.PDFUrl) + ") | [View Status](" + Uri.EscapeUriString(objDeviationDetails.StatusPageUrl) + ") | [Queries](" + Uri.EscapeUriString(objDeviationDetails.DNCollaborationDocUrl) + ")");
                    }
                    else
                    {
                        sbText.AppendLine("\n\n[View Approval Note](" + Uri.EscapeUriString(objDeviationDetails.PDFUrl) + ") | [View Status](" + Uri.EscapeUriString(objDeviationDetails.StatusPageUrl) + ") | [Queries](" + Uri.EscapeUriString(objDeviationDetails.DNCollaborationDocUrl) + ")");
                    }
                }
                else
                {
                    if (parsed["WFStage"] != null && parsed["WFStage"].Value == "0")
                    {
                        if(parsed["Note"] == "Deviation")
                        {
                            sbText.AppendLine("\n\n[View Deviation Note](" + Uri.EscapeUriString(objDeviationDetails.PDFUrl) + ") | [View Status](" + Uri.EscapeUriString(objDeviationDetails.StatusPageUrl) + ") | [Edit Deviation](" + Uri.EscapeUriString(objDeviationDetails.EditUrl) + ")");
                        }
                        else
                        {
                            sbText.AppendLine("\n\n[View Approval Note](" + Uri.EscapeUriString(objDeviationDetails.PDFUrl) + ") | [View Status](" + Uri.EscapeUriString(objDeviationDetails.StatusPageUrl) + ") | [Edit Approval](" + Uri.EscapeUriString(objDeviationDetails.EditUrl) + ")");
                        }
                    }
                    else
                    {
                        if (parsed["Note"] == "Deviation")
                        {
                            sbText.AppendLine("\n\n[View Deviation Note](" + Uri.EscapeUriString(objDeviationDetails.PDFUrl) + ") | [View Status](" + Uri.EscapeUriString(objDeviationDetails.StatusPageUrl) + ")");
                        }
                        else
                        {
                            sbText.AppendLine("\n\n[View Approval Note](" + Uri.EscapeUriString(objDeviationDetails.PDFUrl) + ") | [View Status](" + Uri.EscapeUriString(objDeviationDetails.StatusPageUrl) + ")");
                        }
                    }
                }
                sbText.AppendLine("\nThis is a system generated email, kindly do not respond to this email.For queries please drop an email to [pchfl.itws@piramal.com]().");
                sbText.AppendLine("\nThanks & Regards,");
                sbText.AppendLine("\nIT SUPPORT");
                objDeviationDetails.ApprovalMailBody = sbText.ToString();
            }
            catch (Exception ex)
            {
                //logError("GetDeviationNote", ex.Message, "142", objDeviationDetails.DeviationItemId);
            }

            //objDeviationDetails.EmailTemplate = sbText.ToString();
        }

        public void GenerateFinalEmailBody(DeviationDetailsProps objDeviationDetails, dynamic parsed)
        {
            StringBuilder sbText = new StringBuilder();
            try
            {
                List<FileLink> lstFile = objDeviationDetails.DeviationNoteFileList;

                sbText.AppendLine("<b><span style='font-size:16px;font-family: Segoe UI;'>Dear User,</span></b></br></br>");
                if(parsed["Note"] == "Deviation")
                {
                    sbText.AppendLine("<span style='font-size:16px;font-family: Segoe UI;'>Deviation Note for <b>" + objDeviationDetails.DealName + "</b> Ref No <b>" + objDeviationDetails.DeviationTitle + "</b> has been approved</span></br><br>");
                }
                else
                {
                    sbText.AppendLine("<span style='font-size:16px;font-family: Segoe UI;'>Approval Note for <b>" + objDeviationDetails.DealName + "</b> Ref No <b>" + objDeviationDetails.DeviationTitle + "</b> has been approved</span></br><br>");
                }
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
                if(parsed["Note"] == "Deviation")
                {
                    sbText.AppendLine("<a style='font-size:16px;font-family: Segoe UI;' href=" + Uri.EscapeUriString(objDeviationDetails.PDFUrl) + ">View Deviation Note</a> | <a style='font-size:16px;font-family: Segoe UI;' href=" + Uri.EscapeUriString(objDeviationDetails.StatusPageUrl) + ">View Status</a> </br></br>");
                }
                else
                {
                    sbText.AppendLine("<a style='font-size:16px;font-family: Segoe UI;' href=" + Uri.EscapeUriString(objDeviationDetails.PDFUrl) + ">View Approval Note</a> | <a style='font-size:16px;font-family: Segoe UI;' href=" + Uri.EscapeUriString(objDeviationDetails.StatusPageUrl) + ">View Status</a> </br></br>");
                }
                sbText.AppendLine("<span style='font-size:16px;font-family: Segoe UI;'>This is a system generated email, kindly do not respond to this email. For queries please drop an email to pchfl.itws@piramal.com </span></br><br>");
                sbText.AppendLine("<span style='font-size:16px;font-family: Segoe UI;'>Thanks & Regards, </span></br><br>");
                sbText.AppendLine("<span style='font-size:16px;font-family: Segoe UI;'>IT Support</span>");
                objDeviationDetails.FinalMailBody = sbText.ToString();
            }
            catch(Exception ex)
            {

            }
        }

        public void GenerateRejectionEmailBody(DeviationDetailsProps objDeviationDetails, dynamic parsed)
        {
            StringBuilder sbText = new StringBuilder();
            try
            {
                List<FileLink> lstFile = objDeviationDetails.DeviationNoteFileList;

                sbText.AppendLine("<b><span style='font-size:16px;font-family: Segoe UI;'>Dear User,</span></b></br></br>");
                if(parsed["Note"] == "Deviation")
                {
                    sbText.AppendLine("<span style='font-size:16px;font-family: Segoe UI;'>Deviation Note for <b>" + objDeviationDetails.DealName + "</b> Ref No <b>" + objDeviationDetails.DeviationTitle + "</b> has been rejected.</span></br><br>");
                }
                else
                {
                    sbText.AppendLine("<span style='font-size:16px;font-family: Segoe UI;'>Approval Note for <b>" + objDeviationDetails.DealName + "</b> Ref No <b>" + objDeviationDetails.DeviationTitle + "</b> has been rejected.</span></br><br>");
                }
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
                sbText.AppendLine("\n<b><span style='font-size:16px;font-family: Segoe UI;'>Rejection Status :</b></span></br><br>");
                sbText.AppendLine("\n<span style='font-size:16px;font-family: Segoe UI;'>Rejected by : (RejectedBy)</span></br><br>");
                sbText.AppendLine("\n<span style='font-size:16px;font-family: Segoe UI;'>Rejection Comments : (RejectionComments)</span></br><br>");
                //sbText.AppendLine("<a style='font-size:16px;font-family: Segoe UI;' href=" + Uri.EscapeUriString(objDeviationDetails.PDFUrl) + ">View Deviation Note</a> | <a style='font-size:16px;font-family: Segoe UI;' href=" + Uri.EscapeUriString(objDeviationDetails.StatusPageUrl) + ">View Status</a> </br></br>");
                if (!string.IsNullOrEmpty(objDeviationDetails.CreditOpsApprovedDate))
                {
                    if(parsed["Note"] == "Deviation")
                    {
                        sbText.AppendLine("<a style='font-size:16px;font-family: Segoe UI;' href=" + Uri.EscapeUriString(objDeviationDetails.PDFUrl) + ">View Deviation Note</a> | <a style='font-size:16px;font-family: Segoe UI;' href=" + Uri.EscapeUriString(objDeviationDetails.StatusPageUrl) + ">View Status</a> | <a style='font-size:16px;font-family: Segoe UI;' href=" + Uri.EscapeUriString(objDeviationDetails.DNCollaborationDocUrl) + ">Queries</a> </br></br>");
                    }
                    else
                    {
                        sbText.AppendLine("<a style='font-size:16px;font-family: Segoe UI;' href=" + Uri.EscapeUriString(objDeviationDetails.PDFUrl) + ">View Approval Note</a> | <a style='font-size:16px;font-family: Segoe UI;' href=" + Uri.EscapeUriString(objDeviationDetails.StatusPageUrl) + ">View Status</a> | <a style='font-size:16px;font-family: Segoe UI;' href=" + Uri.EscapeUriString(objDeviationDetails.DNCollaborationDocUrl) + ">Queries</a> </br></br>");
                    }
                }
                else
                {
                    if (parsed["Note"] == "Deviation")
                    {
                        sbText.AppendLine("<a style='font-size:16px;font-family: Segoe UI;' href=" + Uri.EscapeUriString(objDeviationDetails.PDFUrl) + ">View Deviation Note</a> | <a style='font-size:16px;font-family: Segoe UI;' href=" + Uri.EscapeUriString(objDeviationDetails.StatusPageUrl) + ">View Status</a> </br></br>");
                    }
                    else
                    {
                        sbText.AppendLine("<a style='font-size:16px;font-family: Segoe UI;' href=" + Uri.EscapeUriString(objDeviationDetails.PDFUrl) + ">View Approval Note</a> | <a style='font-size:16px;font-family: Segoe UI;' href=" + Uri.EscapeUriString(objDeviationDetails.StatusPageUrl) + ">View Status</a> </br></br>");
                    }
                }
                sbText.AppendLine("<span style='font-size:16px;font-family: Segoe UI;'>This is a system generated email, kindly do not respond to this email. For queries please drop an email to pchfl.itws@piramal.com </span></br><br>");
                sbText.AppendLine("<span style='font-size:16px;font-family: Segoe UI;'>Thanks & Regards, </span></br><br>");
                sbText.AppendLine("<span style='font-size:16px;font-family: Segoe UI;'>IT Support</span>");
                objDeviationDetails.RejectedMailBody = sbText.ToString();
            }
            catch (Exception ex)
            {

            }
        }

        public void GenerateDMSEmailBody(DeviationDetailsProps objDeviationDetails, dynamic parsed)
        {
            StringBuilder sbText = new StringBuilder();
            try
            {
                List<FileLink> lstFile = objDeviationDetails.DeviationNoteFileList;

                sbText.AppendLine("<b><span style='font-size:16px;font-family: Segoe UI;'>Dear User,</span></b></br></br>");
                if(parsed["Note"] == "Deviation")
                {
                    sbText.AppendLine("<span style='font-size:16px;font-family: Segoe UI;'>Deviation Note <b>" + objDeviationDetails.DeviationTitle + "</b> and its attachments for <b>" + objDeviationDetails.DealName + "</b> could not be placed into DMS as required folder does not exists.</span></br><br>");
                }
                else
                {
                    sbText.AppendLine("<span style='font-size:16px;font-family: Segoe UI;'>Approval Note <b>" + objDeviationDetails.DeviationTitle + "</b> and its attachments for <b>" + objDeviationDetails.DealName + "</b> could not be placed into DMS as required folder does not exists.</span></br><br>");
                }
                sbText.AppendLine("<span style='font-size:16px;font-family: Segoe UI;'>Kindly place the required documents in DMS of the captioned deal.</span></br><br>");
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
                if(parsed["Note"] == "Deviation")
                {
                    sbText.AppendLine("<a style='font-size:16px;font-family: Segoe UI;' href=" + Uri.EscapeUriString(objDeviationDetails.PDFUrl) + ">View Deviation Note</a> | <a style='font-size:16px;font-family: Segoe UI;' href=" + Uri.EscapeUriString(objDeviationDetails.StatusPageUrl) + ">View Status</a> </br></br>");
                }
                else
                {
                    sbText.AppendLine("<a style='font-size:16px;font-family: Segoe UI;' href=" + Uri.EscapeUriString(objDeviationDetails.PDFUrl) + ">View Approval Note</a> | <a style='font-size:16px;font-family: Segoe UI;' href=" + Uri.EscapeUriString(objDeviationDetails.StatusPageUrl) + ">View Status</a> </br></br>");
                }
                sbText.AppendLine("<span style='font-size:16px;font-family: Segoe UI;'>This is a system generated email, kindly do not respond to this email. For queries please drop an email to pchfl.itws@piramal.com </span></br><br>");
                sbText.AppendLine("<span style='font-size:16px;font-family: Segoe UI;'>Thanks & Regards, </span></br><br>");
                sbText.AppendLine("<span style='font-size:16px;font-family: Segoe UI;'>IT Support</span>");
                objDeviationDetails.LenderNonExistenceMailBody = sbText.ToString();

                sbText.Clear();
                sbText.AppendLine("<b><span style='font-size:16px;font-family: Segoe UI;'>Dear User,</span></b></br></br>");
                if(parsed["Note"] == "Deviation")
                {
                    sbText.AppendLine("<span style='font-size:16px;font-family: Segoe UI;'>Deviation Note <b>" + objDeviationDetails.DeviationTitle + "</b> and its attachments for <b>" + objDeviationDetails.DealName + "</b> could not be placed into DMS as required folder under given Lender does not exists.</span></br><br>");
                }
                else
                {
                    sbText.AppendLine("<span style='font-size:16px;font-family: Segoe UI;'>Approval Note <b>" + objDeviationDetails.DeviationTitle + "</b> and its attachments for <b>" + objDeviationDetails.DealName + "</b> could not be placed into DMS as required folder under given Lender does not exists.</span></br><br>");
                }
                sbText.AppendLine("<span style='font-size:16px;font-family: Segoe UI;'>Kindly place the required documents in DMS of the captioned deal.</span></br><br>");
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
                if(parsed["Note"] == "Deviation")
                {
                    sbText.AppendLine("<a style='font-size:16px;font-family: Segoe UI;' href=" + Uri.EscapeUriString(objDeviationDetails.PDFUrl) + ">View Deviation Note</a> | <a style='font-size:16px;font-family: Segoe UI;' href=" + Uri.EscapeUriString(objDeviationDetails.StatusPageUrl) + ">View Status</a> </br></br>");
                }
                else
                {
                    sbText.AppendLine("<a style='font-size:16px;font-family: Segoe UI;' href=" + Uri.EscapeUriString(objDeviationDetails.PDFUrl) + ">View Approval Note</a> | <a style='font-size:16px;font-family: Segoe UI;' href=" + Uri.EscapeUriString(objDeviationDetails.StatusPageUrl) + ">View Status</a> </br></br>");
                }
                sbText.AppendLine("<span style='font-size:16px;font-family: Segoe UI;'>This is a system generated email, kindly do not respond to this email. For queries please drop an email to pchfl.itws@piramal.com </span></br><br>");
                sbText.AppendLine("<span style='font-size:16px;font-family: Segoe UI;'>Thanks & Regards, </span></br><br>");
                sbText.AppendLine("<span style='font-size:16px;font-family: Segoe UI;'>IT Support</span>");
                objDeviationDetails.FolderNonExistenceMailBody = sbText.ToString();
            }
            catch (Exception ex)
            {

            }
        }
    }
}

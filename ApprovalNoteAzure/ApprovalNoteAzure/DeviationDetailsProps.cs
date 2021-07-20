using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ApprovalNoteAzure
{
    public class DeviationDetailsProps
    {
        public string ItemId { get; set; }
        public string DealItemId { get; set; }
        public string DealName { get; set; }
        public string DeviationItemId { get; set; }
        public string DeviationTitle { get; set; }
        public string Date { get; set; }
        public string UCIC { get; set; }
        public string DNNo { get; set; }
        public string NatureOfDeviation { get; set; }
        public string Borrower { get; set; }
        public string Instrument { get; set; }
        public string Lender { get; set; }
        public string SanctionedAmount { get; set; }
        public string DateOfSanction { get; set; }
        public string Purpose { get; set; }
        public string Tenure { get; set; }
        public string IRR_CoupounRate { get; set; }
        public string SecurityDetails_Primary { get; set; }
        public string SecurityDetails_Additional { get; set; }
        public string SecurityDetails_Collateral { get; set; }
        public string Cover { get; set; }
        public string PrincipalServicing { get; set; }
        public string PrepaymentTillDate { get; set; }
        public string ConductOfAccount { get; set; }
        
        public string Remarks_Risk_Legal { get; set; }
        public string Reviewer_Creditops { get; set; }
        public string RecommendationFrom { get; set; }

        public string ExposureType { get; set; }

        // public string[][] RTBDetails {get;set;}
        //public DataTable RTBTable = new DataTable();

        public List<RTBDetails> RTBData { get; set; }

        public string LenderShortCode { get; set; }
        public string PiramalLogoName { get; set; }

        public string PDFUrl { get; set; }


        public string LastModified { get; set; }

        public string BusinessSegment { get; set; }

        public string RegionShortCode { get; set; }

        public string DealID { get; set; }

        public string Borrowergroup { get; set; }

        public string PDFName { get; set; }
        public string EditUrl { get; set; }

        public string DeviationGUID { get; set; }

        public string CreditOpsStatus { get; set; }
        public string CreditOpsUser { get; set; }

        public string CreditOpsApprovedBy { get; set; }
        public string CreditOpsApprovedDate { get; set; }

        public int recommendedByCount { get; set; }

        public string CFOEndDate { get; set; }

        public string CFOApprovedBy { get; set; }
        public string ApprovalMailBody { get; set; }
        public string RejectedMailBody { get; set; }
        public string DeletedMailBody { get; set; }
        public string FinalMailBody { get; set; }
        public string LenderNonExistenceMailBody { get; set; }
        public string FolderNonExistenceMailBody { get; set; }
        public string PDFFileItemID { get; set; }
        public string DashBoardUrl { get; set; }
        public string StatusPageUrl { get; set; }
        public string DNCollaborationDocUrl { get; set; }
        public string DNRefNo { get; set; }
        public List<Recommenders> RecommendersList { get; set; }

        public List<FileLink> DeviationNoteFileList { get; set; }

        public string InsideCreateWordFile { get; set; }
        public string InsideSaveHtmlFile { get; set; }
        public string InsideSaveDocInList { get; set; }
        public string SecurityCover { get; set; }
        public string ProcessingFee { get; set; }
        public string CouponRate { get; set; }
        public bool CreditOpsToBeApproved { get; set; }
        public string CreditOpsAssignedDate { get; set; }
        public string CreditOpsEmailId { get; set; }
        public string CreditOpsComments { get; set; }
        public string Annexure { get; set; }
        public List<HolidayMaster> HolidayMasterList { get; set; }
        public int TATValue { get; set; }
        public string FinancialYear { get; set; }
        public string WorkflowStatus { get; set; }
    }

    public class RTBDetails
    {
        public string Originalterms { get; set; }
        public string ProposedDeviation { get; set; }
        public string Impact { get; set; }
        public string Justification{ get; set; }

};
}

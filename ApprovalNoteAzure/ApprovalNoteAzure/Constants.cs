using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ApprovalNoteAzure
{
    public class Constants
    {
        /// <summary>
        /// Constant for SiteURL for Dev
        /// </summary>
        public const string DEVSiteURL = "https://mypiramal.sharepoint.com/sites/piramaldevops/creditopsdev";

        /// <summary>
        /// Constant for SiteURL for UAT
        /// </summary>
        public const string UATSiteURL = "https://mypiramal.sharepoint.com/sites/piramaldev/creditopstest";

        /// <summary>
        /// Constant for SiteURL for Prod
        /// </summary>
        public const string PRODSiteURL = "https://mypiramal.sharepoint.com/sites/finance/SAMVAD";

        /// <summary>
        /// Constant for Time Zone
        /// </summary>
        public const string IndTimeZone = "India Standard Time";

        /// <summary>
        /// Constant for Tasklist
        /// </summary>
        public const string Deviation_WorkFlowTaskList = "DN_WorkFlow_TaskList";

        /// <summary>
        /// Constant for Dev Username
        /// </summary>
        public const string DEVUserName = "SVC-Automation.support1@piramal.com";

        /// <summary>
        /// Constant for Dev Password
        /// </summary>
        public const string DEVPassword = "@utoMTS1";

        /// <summary>
        /// Constant for UAT Username
        /// </summary>
        public const string UATUserName = "SVC-Automation.UAT1@piramal.com";

        /// <summary>
        /// Constant for UAT Password
        /// </summary>
        public const string UATPassword = "@utoMTU1";

        /// <summary>
        /// Constant for PROD Username
        /// </summary>
        public const string PRODUserName = "SVC-Automation.support1@piramal.com";

        /// <summary>
        /// Constant for PROD Password
        /// </summary>
        public const string PRODPassword = "@utoMTS1";

        /// <summary>
        /// Constant for Uat Discussion Board Content Type
        /// </summary>
        public const string UATDBCTID = "0x01200200000DFEF8A15C0D4DA60B6805BEB3E260";

        /// <summary>
        /// Constant for PROD Discussion Board Content Type
        /// </summary>
        public const string PRODDBCTID = "0x01200200ED4F63F478E48D4885675D9861B1D98C";

        /// <summary>
        /// Constant for WorkFlow Task List
        /// </summary>
        public const string DN_WorkFlowTaskList = "DN_WorkFlow_TaskList";

        /// <summary>
        /// Constant for Deviation Note List
        /// </summary>
        public const string DeviationNote = "DeviationNote";

        /// <summary>
        /// Constant for DN_StakeHolders List
        /// </summary>
        public const string DN_StakeHolders = "DN_StakeHolders";

        /// <summary>
        /// Constant for Base Site URL
        /// </summary>
        public const string BaseSiteURL = "https://mypiramal.sharepoint.com";

        /// <summary>
        /// Constant for PDF File Document Library
        /// </summary>
        public const string PDFFile_DocumentLibrary = "DeviationNote_Created_PDF";
    }
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ApprovalNoteAzure
{
    public class Recommenders
    {
        public string Designation { get; set; }//StakeHolder
        public string Department { get; set; }//StakeHolder
        public string DesignationDepartment { get; set; }//StakeHolder
        public string Email { get; set; }//StakeHolder
        public string Name { get; set; }//WorkFlow
        public string ApprovedDate { get; set; }//WorkFlow
        public string Role { get; set; }//StakeHolder & WorkFlow
        public string Status { get; set; }//WorkFlow
        public int ItemId { get; set; }
        public string WFProcess { get; set; }
        public string Comments { get; set; }
        public string AssignedDate { get; set; }//WorkFlow
    }
}

using System;

namespace ContactImport.Models
{
    public class ContractModel
    {
        public string? Client { get; set; }
        public string? SingleMaster { get; set; }
        public string? JointVenture { get; set; }
        public string? Name { get; set; }
        public string? ShortName { get; set; }
        public string? ContactNumber { get; set; }
        public string? StartDate { get; set; }
        public string? EndDate { get; set; }
        public string? ContractManager { get; set; }
        public string? TimesheetVersionType { get; set; }
    }
    public class InputModel
    {

       
        public string data { get; set; }



    }
}

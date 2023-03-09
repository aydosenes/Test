using System;
namespace ImportExportAPI.Model.DataModel
{
    public class ProjectRow
    {
        public ProjectRow()
        {
            
        }
       

        public String Id { get; set; }

        public String Code { get; set; }

        public String SapCode { get; set; }

        public String Name { get; set; }

        public String ShortName { get; set; }

        public String ContractorName { get; set; }

        public String Type { get; set; }

        public String Directorate { get; set; }

        public String HeadShip { get; set; }

        public String Status { get; set; }

    }
}

using System;
using System.Collections.Generic;

namespace ImportExportAPI.Model.DataModel
{
    public class PersonelRow
    {
        public PersonelRow()
        {
            
        }
       

        public String OrderNo { get; set; }

        public String RegistrationNo { get; set; }

        public String FullName { get; set; }

        public String Job { get; set; }

        public String Role { get; set; }

        public Boolean IsGeneralWorker { get; set; }

        public int ProjectCount { get; set; }

        public List<String> Projects { get; set; }

    }
}

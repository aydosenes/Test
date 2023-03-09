using System;
using System.Collections.Generic;

namespace ImportExportAPI.Model.DataModel
{
    public class PersonelCard
    {
        public PersonelCard()
        {

            
        }
        public List<PersonelRow> PersonelList { get; set; } = new List<PersonelRow>();

    }
}

using System;
using System.Collections.Generic;

namespace ImportExportAPI.Model.DataModel
{
    public class ProjectCard
    {
        public ProjectCard()
        {

            
        }
        public List<ProjectRow> ProjectList { get; set; } = new List<ProjectRow>();

        public List<ProjectRow> CurrentProjectList { get; set; } = new List<ProjectRow>();

        public List<ProjectRow> DoneProjectList { get; set; } = new List<ProjectRow>();

    }
}

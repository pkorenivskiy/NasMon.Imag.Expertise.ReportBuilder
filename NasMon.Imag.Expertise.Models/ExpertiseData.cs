using System;
using System.Collections.Generic;
using System.Text;

namespace NasMon.Imag.Expertise.Models
{
    public class ExpertiseData
    {
        public string ProjectLead { get; set; }
        public string ProjectCode { get; set; }        
        public string ProjectTitle { get; set; }
        public List<ExpertPoint> ExpertPoints { get; set; } = new List<ExpertPoint>();
        public decimal AvgPoints { get; set; }
    }
}

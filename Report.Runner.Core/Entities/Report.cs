using System;
using System.Collections.Generic;
using System.Text;
using System.Xml.Serialization;

namespace Report.Runner.Core.Entities
{
    [XmlRoot(ElementName = "Report")]
    public class Report
    {
        public Report()
        {
            ReportVals = new List<ReportVal>();
        }

        [XmlElement()]
        public string Name { get; set; }

        [XmlElement(ElementName = "ReportVal")]
        public List<ReportVal> ReportVals { get; set; }
    }
}

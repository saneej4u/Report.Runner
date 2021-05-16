using Report.Runner.Core.Reports;
using System;
using System.Collections.Generic;
using System.Text;

namespace Report.Runner.Core.Repository
{
    public class ReportRepository : IReportRepository
    {
        private const string F2004Report  = "F2004Report.xlsx";
        private const string SampleReport = "SampleReport.xlsx";
        public byte[] Process(string templateName)
        {
            //TODO: Make it Generic class
            if(templateName == "F2004Report")
            {
                return new F2004Report().ProcessReport(F2004Report).ToArray();
            }
            else
            {
                return new SampleReport().ProcessReport(SampleReport).ToArray();
            }
        }
    }
}

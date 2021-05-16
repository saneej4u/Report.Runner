using Aspose.Cells;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml;
using System.Xml.Serialization;
using System.Linq;
using Report.Runner.Core.Entities;

namespace Report.Runner.Core
{
    public class F2004Report : IReport
    {
        /// <summary>
        /// Process Report
        /// </summary>
        /// <param name="templateName"></param>
        /// <returns>XLSX Memory Stream</returns>
        public MemoryStream ProcessReport(string templateName)
        {
            if (string.IsNullOrEmpty(templateName))
                throw new ArgumentNullException("templateName");

            try
            {
                string templatePath = Path.Combine(Environment.CurrentDirectory, "Templates", templateName);
                string feedPath     = Path.Combine(Environment.CurrentDirectory, "Feeds", "FeedReport.xml");

                XmlDocument feedReportDocument = new XmlDocument();
                feedReportDocument.Load(feedPath);

                string feedReportcontents = feedReportDocument.InnerXml;

                XmlSerializer reportSerializer  = new XmlSerializer(typeof(List<Report.Runner.Core.Entities.Report>), new XmlRootAttribute("Reports"));
                StringReader reportStringReader = new StringReader(feedReportcontents);

                var reports = (List<Report.Runner.Core.Entities.Report>)reportSerializer.Deserialize(reportStringReader);

 
                //Create a workbook object
                Workbook book = new Workbook(templatePath);
                //Access first worksheet - F2004
                Worksheet workSheet = book.Worksheets[0];

                foreach (var report in reports)
                {
                    foreach (var reportVal in report.ReportVals)
                    {
                        var cellIndex   = GetCellIndex(reportVal);
                        Cell reportCell = workSheet.Cells[cellIndex];

                        reportCell.PutValue(reportVal.Val);
                    }
                }

                //Save the workbook
                MemoryStream workBookStream = new MemoryStream();

                book.Save(workBookStream, SaveFormat.Xlsx);

                return workBookStream;
            }
            catch(Exception ex)
            {
                throw new Exception(ex.Message);
            }
        }


        /// <summary>
        /// Get Cell index based on Feed (XML) values
        /// </summary>
        /// <param name="reportVal"></param>
        /// <returns>Cell Index</returns>
        private static string GetCellIndex(ReportVal reportVal)
        {
            if (Convert.ToInt16(reportVal.ReportRow) == 10 && Convert.ToInt16(reportVal.ReportCol) == 10)
            {
                return "E11";
            }
            else if (Convert.ToInt16(reportVal.ReportRow) == 10 && Convert.ToInt16(reportVal.ReportCol) == 11)
            {
                return "F11";
            }
            else if (Convert.ToInt16(reportVal.ReportRow) == 10 && Convert.ToInt16(reportVal.ReportCol) == 12)
            {
                return "G11";
            }
            else if (Convert.ToInt16(reportVal.ReportRow) == 20 && Convert.ToInt16(reportVal.ReportCol) == 10)
            {
                return "E12";
            }
            else if (Convert.ToInt16(reportVal.ReportRow) == 20 && Convert.ToInt16(reportVal.ReportCol) == 11)
            {
                return "F12";
            }
            else if (Convert.ToInt16(reportVal.ReportRow) == 20 && Convert.ToInt16(reportVal.ReportCol) == 12)
            {
                return "G12";
            }


            return string.Empty;
        }
    }
}

using Aspose.Cells;
using Report.Runner.Core.Entities;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml;
using System.Xml.Serialization;

namespace Report.Runner.Core.Reports
{
    public class SampleReport : IReport
    {
        // TODO: Update the logic based on requirement, at the moment it is using sample template with same logic as F2004
        public MemoryStream ProcessReport(string templateName)
        {
            string feedName = "FeedReport.xml";

            string templatePath = Path.Combine(Environment.CurrentDirectory, "Templates", templateName);
            string feedPath = Path.Combine(Environment.CurrentDirectory, "Feeds", feedName);

            XmlDocument doc = new XmlDocument();
            doc.Load(feedPath);
            string xmlcontents = doc.InnerXml;


            XmlSerializer serializer = new XmlSerializer(typeof(List<Report.Runner.Core.Entities.Report>), new XmlRootAttribute("Reports"));

            StringReader stringReader = new StringReader(xmlcontents);

            var reports = (List<Report.Runner.Core.Entities.Report>)serializer.Deserialize(stringReader);

            //open a template excel file.
            Workbook book = new Workbook(templatePath);

            //Create a workbook object
            Workbook wb = new Workbook(templatePath);
            //Access first worksheet
            Worksheet ws = wb.Worksheets[0];

            foreach (var report in reports)
            {
                foreach (var reportVal in report.ReportVals)
                {
                    var cellIndex = GetCellIndex(reportVal);
                    //Access cell A1
                    Cell reportCell = ws.Cells[cellIndex];
                    reportCell.PutValue(reportVal.Val);
                }
            }

            //Save the workbook
            MemoryStream ms2 = new MemoryStream();

            wb.Save(ms2, SaveFormat.Xlsx);

            return ms2;
        }
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

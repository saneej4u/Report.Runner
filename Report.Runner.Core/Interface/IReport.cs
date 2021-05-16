using System.IO;

namespace Report.Runner.Core
{
    public interface IReport
    {
        MemoryStream ProcessReport(string templateName);
    }
}
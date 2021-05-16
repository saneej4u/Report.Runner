namespace Report.Runner.Core.Repository
{
    public interface IReportRepository
    {
        byte[] Process(string templateName);
    }
}
using Aspose.Cells;
using NUnit.Framework;
using Report.Runner.Core;
using System;

namespace Report.Runner.UnitTest
{
    public class WhenF2004ReportIsInvoked
    {
        [SetUp]
        public void Setup()
        {
        }

        [Test]
        public void F2004Report_Process()
        {
            var f2004Report = new F2004Report();
            var process     = f2004Report.ProcessReport("F2004Report.xlsx");

            Assert.IsNotNull(process);

        }

        [Test]
        public void F2004Report_Process_Throws_Error()
        {
            var f2004Report = new F2004Report();
            Assert.Throws<ArgumentNullException>(() => f2004Report.ProcessReport(string.Empty));
        }

    }
}
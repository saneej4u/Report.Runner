using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Report.Runner.Core;
using Report.Runner.Core.Repository;

namespace Report.Runner.Web.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class ReportRunnerController : ControllerBase
    {
        private readonly ILogger<ReportRunnerController> _logger;
        private readonly IReportRepository               _reportRepository;

        public ReportRunnerController(ILogger<ReportRunnerController> logger, IReportRepository reportRepository)
        {
            _logger           = logger;
            _reportRepository = reportRepository;
        }

        [HttpGet]
        public ActionResult<IEnumerable<byte>> Get(string templateName)
        {
            if(string.IsNullOrEmpty(templateName))
            {
                return BadRequest();
            }

            return Ok( _reportRepository.Process(templateName));
        }

        [HttpGet()]
        [HttpHead]
        [Route("types")]
        public ActionResult<IEnumerable<TemplateType>> GetTemplateTypes()
        {
            return Ok(GetTypes());
        }

        private List<TemplateType> GetTypes() => new List<TemplateType>
        {
            new TemplateType {Id = 1, Name = "F2004Report"},
            new TemplateType {Id = 2, Name = "SampleReport"}
        };
    }

    public class TemplateType
    {
        public int Id { get; set; }
        public string Name { get; set; }
    }
}

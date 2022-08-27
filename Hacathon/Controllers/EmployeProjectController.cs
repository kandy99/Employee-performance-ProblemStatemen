using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Threading.Tasks;
using System.IO;
using DevExpress.DataProcessing;
using DevExpress.DataProcessingApi.Model;
using Hacathon.Helper;
using DevExpress.Data.Filtering;
using DevExpress.DataProcessing.InMemoryDataProcessor;

namespace Hacathon.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class EmployeeProjectController : ControllerBase
    {

        private readonly ILogger<EmployeeProjectController> _logger;

        public EmployeeProjectController(ILogger<EmployeeProjectController> logger)
        {
            _logger = logger;
        }




        [HttpGet]

        [Route("GetMeanEffortPerTeam")]

        public IActionResult GetMeanEffortPerTeam(bool considerbillable = false)
        {


            if (considerbillable)
            {

                var getData = DataFlow.FromExcel(Path.Combine("Data", "HackathonTimesheet.xlsx"), "All Projects").Filter("Billing Status", (string value) => value != "Non Billable").Filter("Owner", (string value) => value != "0").ToDataTable().Execute();
                var data = DataFlow.FromObject(getData).Aggregate(e => e.GroupBy("Team", "Project Name").Summary("Hours", AggregationType.Average, "Efforts")).ToDataTable().Execute();

                var meanEffort = DataFlow.FromObject(data).ToJsonString().Execute();




                return Ok(meanEffort);
            }
            else
            {
                var getData = DataFlow.FromExcel(Path.Combine("Data", "HackathonTimesheet.xlsx"), "All Projects").Filter("Owner", (string value) => value != "0").ToDataTable().Execute();
                var data = DataFlow.FromObject(getData).Aggregate(e => e.GroupBy("Team", "Project Name").Summary("Hours", AggregationType.Average, "Efforts")).ToDataTable().Execute();

                var meanEffort = DataFlow.FromObject(data).ToJsonString().Execute();




                return Ok(meanEffort);
            }


           

        }


     


        [HttpGet]
        [Route("LeastEfficientGroup")]

        public IActionResult LeastEfficientGroup(bool considerbillable = false)
        {

            if (considerbillable)
            {

                var GetData = DataFlow.FromExcel(Path.Combine("Data", "HackathonTimesheet.xlsx"), "All Projects").Filter("Billing Status", (string value) => value != "Non Billable" ).Filter("Owner", (string value) => value != "0")
                    .ToDataTable().Execute();

                System.Data.DataTable result = DataFlow
       .FromObject(GetData).Aggregate(e => e.GroupBy("Owner").Summary("Hours", AggregationType.Average, "Efforts")).Sort(e =>
       {
           e.SortColumns.Add("Efforts", SortOrder.Ascending);
       })
      .ToDataTable()
       .Execute();



                var leastEfficientGroup = DataFlow.FromObject(result).Bottom(5, "Efforts").ToJsonString().Execute();

                return Ok(leastEfficientGroup);
            }


            else
            {
                var GetData = DataFlow.FromExcel(Path.Combine("Data", "HackathonTimesheet.xlsx"), "All Projects").Filter("Owner", (string value) => value != "0")
                    .ToDataTable().Execute();

                System.Data.DataTable result = DataFlow
       .FromObject(GetData).Aggregate(e => e.GroupBy("Owner").Summary("Hours", AggregationType.Average, "Efforts")).Sort(e =>
       {
           e.SortColumns.Add("Efforts", SortOrder.Ascending);
       })
      .ToDataTable()
       .Execute();



                var leastEfficientGroup = DataFlow.FromObject(result).Bottom(5, "Efforts").ToJsonString().Execute();

                return Ok(leastEfficientGroup);
            }

        }


    }
}

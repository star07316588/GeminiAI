using Microsoft.AspNetCore.Mvc;
using MES.Models.Dto.MergeSplitQuery;
using MES.Services;
using System;
using System.Threading.Tasks;

namespace MES.Controllers.API
{
    [Route("api/[controller]")]
    [ApiController]
    public class MergeSplitQueryController : ControllerBase
    {
        private readonly MergeSplitQueryService _service;

        public MergeSplitQueryController(MergeSplitQueryService service)
        {
            _service = service;
        }

        [HttpPost("Query")]
        public async Task<IActionResult> Query([FromBody] MergeSplitQueryRequest req)
        {
            // 防呆檢核：至少需要一個條件
            if (string.IsNullOrWhiteSpace(req.ParentLotId) && string.IsNullOrWhiteSpace(req.Ipn) && 
                string.IsNullOrWhiteSpace(req.ChildLotId) && string.IsNullOrWhiteSpace(req.StepName) &&
                string.IsNullOrWhiteSpace(req.TimeStampFrom) && string.IsNullOrWhiteSpace(req.TimeStampTo))
            {
                return BadRequest(new { success = false, message = "請輸入至少一個條件" });
            }

            try
            {
                var result = await _service.ExecuteQueryAsync(req);
                return Ok(new { success = true, data = result });
            }
            catch (Exception ex)
            {
                // 可以加入 Logger
                return StatusCode(500, new { success = false, message = ex.Message });
            }
        }
    }
}

using Microsoft.AspNetCore.Mvc;
using System;
using System.Threading.Tasks;
using MES.Models.Dto.MonitorPCQuery;
using MES.Services;

namespace MES.Controllers.API
{
    [Route("api/[controller]")]
    [ApiController]
    public class MonitorPCQueryController : ControllerBase
    {
        private readonly MonitorPCQueryService _service;

        public MonitorPCQueryController(MonitorPCQueryService service)
        {
            _service = service;
        }

        [HttpGet("InitData")]
        public async Task<IActionResult> GetInitData()
        {
            try
            {
                var result = await _service.GetInitDataAsync();
                return Ok(new { success = true, data = result });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { success = false, message = ex.Message });
            }
        }

        [HttpPost("Query")]
        public async Task<IActionResult> Query([FromBody] MonitorPCQueryRequest req)
        {
            try
            {
                var result = await _service.ExecuteQueryAsync(req ?? new MonitorPCQueryRequest());
                return Ok(new { success = true, data = result });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { success = false, message = ex.Message });
            }
        }

        [HttpPost("Detail")]
        public async Task<IActionResult> GetDetail([FromBody] MonitorPCDetailRequest req)
        {
            if (req == null || string.IsNullOrWhiteSpace(req.IP) || string.IsNullOrWhiteSpace(req.SystemName))
            {
                return BadRequest(new { success = false, message = "IP 與 SystemName 不可為空" });
            }

            try
            {
                var result = await _service.GetDetailAsync(req);
                return Ok(new { success = true, data = result });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { success = false, message = ex.Message });
            }
        }
    }
}

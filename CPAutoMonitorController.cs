using Microsoft.AspNetCore.Mvc;
using System;
using System.Threading.Tasks;
using MES.Services;

namespace MES.Controllers.API
{
    [Route("api/[controller]")]
    [ApiController]
    public class CPAutoMonitorController : ControllerBase
    {
        private readonly CPAutoMonitorService _service;

        public CPAutoMonitorController(CPAutoMonitorService service)
        {
            _service = service;
        }

        [HttpGet("Query")]
        public async Task<IActionResult> Query()
        {
            try
            {
                var result = await _service.ExecuteMonitorQueryAsync();
                return Ok(new { success = true, data = result });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { success = false, message = ex.Message });
            }
        }
    }
}

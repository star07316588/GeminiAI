using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using MES.Models.Dto.YellowCardReSend;
using MES.Services;

namespace MES.Controllers.API
{
    [Route("api/[controller]")]
    [ApiController]
    public class YellowCardReSendController : ControllerBase
    {
        private readonly YellowCardReSendService _service;

        public YellowCardReSendController(YellowCardReSendService service)
        {
            _service = service;
        }

        [HttpGet("LotInfo")]
        public async Task<IActionResult> GetLotInfo([FromQuery] string lotId)
        {
            try
            {
                var result = await _service.GetLotInfoAsync(lotId);
                return Ok(new { success = true, data = result });
            }
            catch (Exception ex) when (ex is ArgumentException || ex is KeyNotFoundException || ex is InvalidOperationException)
            {
                return Ok(new { success = false, message = ex.Message });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { success = false, message = ex.Message });
            }
        }

        [HttpPost("Send")]
        public async Task<IActionResult> Send([FromBody] YellowCardSendRequest req)
        {
            try
            {
                int count = await _service.SendYellowCardAsync(req);
                return Ok(new { success = true, message = $"成功送出 {count} 筆資料。" });
            }
            catch (Exception ex) when (ex is ArgumentException || ex is KeyNotFoundException)
            {
                return Ok(new { success = false, message = ex.Message });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { success = false, message = ex.Message });
            }
        }
    }
}

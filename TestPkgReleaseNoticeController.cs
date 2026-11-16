using Microsoft.AspNetCore.Mvc;
using MES.Models.Dto.TestPkgReleaseNotice;
using MES.Repositories;

namespace MES.Controllers.API
{
    [Route("api/[controller]")]
    [ApiController]
    public class TestPkgReleaseNoticeController : ControllerBase
    {
        private readonly TestPkgReleaseNoticeRepository _repository;

        public TestPkgReleaseNoticeController(TestPkgReleaseNoticeRepository repository)
        {
            _repository = repository;
        }

        [HttpGet("InitData")]
        public async Task<IActionResult> GetInitData()
        {
            try
            {
                var result = await _repository.GetInitDataAsync();
                return Ok(new { success = true, data = result });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { success = false, message = ex.Message });
            }
        }

        [HttpPost("Query")]
        public async Task<IActionResult> QueryLotInfo([FromBody] ReleaseNoticeQueryRequest req)
        {
            // 防呆：確認至少輸入一個條件 (對應原 VB6 防呆)
            if (string.IsNullOrWhiteSpace(req.TdsProdCode) && string.IsNullOrWhiteSpace(req.ProdGroup) &&
                string.IsNullOrWhiteSpace(req.PkgType) && string.IsNullOrWhiteSpace(req.StepName) &&
                string.IsNullOrWhiteSpace(req.EqType2) && string.IsNullOrWhiteSpace(req.SubSystem))
            {
                return BadRequest(new { success = false, message = "Pls input Query Criteria !! 請輸入查詢條件 !!" });
            }

            try
            {
                var result = await _repository.QueryLotInfoAsync(req);
                if (!result.Any())
                {
                    return Ok(new { success = true, data = result, message = "無符合條件之資料 !!" });
                }
                
                return Ok(new { success = true, data = result });
            }
            catch (Exception ex)
            {
                return StatusCode(500, new { success = false, message = ex.Message });
            }
        }
    }
}

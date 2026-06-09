using MES.Net.Application.Services.Query;
using MES.Net.Infrastructure.Logging;
using MES.Net.Web.Filters;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Web.Http;
using static MES.Net.Shared.Constants.GlobalConstants;
using static MES.Net.Shared.DTOs.WsMergeQuery;

namespace MES.Net.Web.Controllers.Query
{
    [RoutePrefix("api/query/ws-merge-query")]
    public class WsMergeQueryController : ApiController
    {
        [HttpGet, HttpPost, Route("init")]
        public IHttpActionResult Init()
        {
            (string[] Steps, string[] Locations) data = WipQueryService.Instance.GetComboBoxOptions();
            var options = new
            {
                success = true,
                Stages = new string[] { StageAY, StageFT, StageWS },
                Steps = data.Steps,
                Locations = data.Locations
            };
            return Ok(options);
        }
        private readonly WsMergeQueryService _wsMergeQueryService;

        public WsMergeQueryController(WsMergeQueryService lotService)
        {
            _wsMergeQueryService = lotService;
        }


        [HttpPost, Route("query"), AuthorizeToken]
        public async Task<IHttpActionResult> ExecQuery([FromBody] QueryLotRequest request)
        {
            // Condition Checking (對應原本的請輸入查詢條件)
            if (string.IsNullOrWhiteSpace(request.LotId))
            {
                return BadRequest("Pls input Query Criteria !! 請輸入查詢條件 !!" );
            }
            try
            {
                var result = await _wsMergeQueryService.GetLotMergeInfoAsync(request.LotId, request.IsHistoryTab);

                if (result == null)
                {
                    return BadRequest($"Lot '{request.LotId}' is no data found !! 查無資料 !!");
                }

                return Ok(new { Success = true, Message = "", Data = result });
            }
            catch (Exception ex)
            {
                AppLogger.Error(this, ex.Message, ex);
                AppLogger.Info(this, JsonConvert.SerializeObject(request));
                if (ex is ArgumentException)
                {
                    return Ok(new { Success = false, Message = ex.Message });
                }
                throw;
            }
        }

        [HttpPost, Route("sorter-control"), AuthorizeToken]
        public async Task<IHttpActionResult> ExecQuerySorterControl([FromBody] QuerySorterControlRequest request)
        {
            // Condition Checking
            if (string.IsNullOrWhiteSpace(request.LotId) || string.IsNullOrWhiteSpace(request.Action))
            {
                return BadRequest("Pls input Query Criterial !! 請輸入查詢條件 !!" );
            }

            try
            {
                var result = await _wsMergeQueryService.GetSorterControlInfoAsync(request.LotId, request.Action);
                return Ok(result);
            }
            catch (Exception ex)
            {
                AppLogger.Error(this, ex.Message, ex);
                AppLogger.Info(this, JsonConvert.SerializeObject(request));
                if (ex is ArgumentException)
                {
                    return Ok(new { Success = false, Message = ex.Message });
                }
                throw;
            }
        }

        [HttpPost, Route("resend-sap")]
        [AuthorizeToken]
        public async Task<IHttpActionResult> ResendSap([FromBody] ResendSapRequest request)
        {
            // Condition Checking (檢查是否有輸入 LotID，以及前端是否有傳入陣列資料)
            if (string.IsNullOrWhiteSpace(request.LotId) || request.Items == null)
            {
                return BadRequest($"Pls input Query Criterial !! / Lot '{request.LotId}' is no data found !! 請輸入查詢條件或查無資料 !!");
            }

            try
            {
                // 執行更新動作
                await _wsMergeQueryService.ResendSapAsync(request);

                // 對應成功的 MsgBox
                return Ok(new
                {
                    Message = $"Lot '{request.LotId}' data is updated !! 資料已更新, 稍後將重傳給SAP !!"
                });
            }
            catch (Exception ex)
            {
                AppLogger.Error(this, ex.Message, ex);
                AppLogger.Info(this, JsonConvert.SerializeObject(request));
                if (ex is ArgumentException)
                {
                    return Ok(new { Success = false, Message = ex.Message });
                }
                throw;
            }
        }

        [HttpPost, Route("testflow")]
        [AuthorizeToken]
        public async Task<IHttpActionResult> TestFlow([FromBody] Dictionary<string, string> data)
        {
            try
            {
                var result = await WipQueryService.Instance.QueryTestFlowAsync(data["lotId"], data["step"],
                    data["ipn"], data["stage"], data["holdType"], data["location"]);
                return Ok(new { Success = true, Message = "", Data = result });
            }
            catch (Exception ex)
            {
                AppLogger.Error(this, ex.Message, ex);
                AppLogger.Info(this, JsonConvert.SerializeObject(data));
                if (ex is ArgumentException)
                {
                    return Ok(new { Success = false, Message = ex.Message });
                }
                throw;
            }
        }

        [HttpPost, Route("detail")]
        public async Task<IHttpActionResult> Detail ([FromBody] Dictionary<string, string> data)
        {
            try
            {
                var result = await WipQueryService.Instance.QueryDetail(data["lotId"]);
                return Ok(new { Success = true, Message = "", Data = result });
            }
            catch (Exception ex)
            {
                AppLogger.Error(this, ex.Message, ex);
                AppLogger.Info(this, JsonConvert.SerializeObject(data));
                if (ex is ArgumentException)
                {
                    return Ok(new { Success = false, Message = ex.Message });
                }
                throw;
            }
        }
    }
}

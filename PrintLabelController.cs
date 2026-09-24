using MES.Net.Application.Services.Print;
using MES.Net.Infrastructure.Logging;
using MES.Net.Web.Filters;
using MES.Net.Shared.DTOs.Print;
using Newtonsoft.Json;
using System;
using System.Threading.Tasks;
using System.Web.Http;

namespace MES.Net.Web.Controllers.Print
{
    [RoutePrefix("api/print/print-label")]
    public class PrintLabelController : ApiController
    {
        private readonly IPrintLabelService _printLabelService;

        public PrintLabelController(IPrintLabelService printLabelService)
        {
            _printLabelService = printLabelService;
        }

        // 1. 畫面初始化 
        // 💡 修正 404 主因: 補上 [FromBody] dynamic req，讓 Web API 能正確接收帶有 {} payload 的 POST 請求
        [HttpPost, Route("init"), AuthorizeToken]
        public async Task<IHttpActionResult> Init([FromBody] dynamic req = null)
        {
            try
            {
                var result = await _printLabelService.GetInitDataAsync();
                return Ok(new { Success = true, Message = "", Data = result });
            }
            catch (Exception ex)
            {
                AppLogger.Error(this, ex.Message, ex);
                return Ok(new { Success = false, Message = "初始化選單資料失敗" });
            }
        }

        // 2. 依據 Stage 取得 Label Formats (💡 補齊遺漏的 API)
        [HttpPost, Route("label-formats"), AuthorizeToken]
        public async Task<IHttpActionResult> GetLabelFormats([FromBody] MappedServerRequest request)
        {
            if (request == null || string.IsNullOrWhiteSpace(request.Stage)) return BadRequest("請先選擇 Stage");
            try
            {
                var result = await _printLabelService.GetLabelFormatsByStageAsync(request.Stage);
                return Ok(new { Success = true, Message = "", Data = result });
            }
            catch (Exception ex)
            {
                AppLogger.Error(this, ex.Message, ex);
                return Ok(new { Success = false, Message = "取得 Label Formats 失敗" });
            }
        }

        // 3. 依據 Stage + Format 取得 印表機伺服器
        [HttpPost, Route("mapped-servers"), AuthorizeToken]
        public async Task<IHttpActionResult> GetMappedPrinterServers([FromBody] MappedServerRequest request)
        {
            try
            {
                var result = await _printLabelService.GetMappedPrinterServersAsync(request.Stage, request.LabelFormat);
                return Ok(new { Success = true, Message = "", Data = result });
            }
            catch (Exception ex)
            {
                AppLogger.Error(this, ex.Message, ex);
                return Ok(new { Success = false, Message = "取得 Printer Server 失敗" });
            }
        }

        // =========================================================================
        // LPI (Label Pack Info) 六層連動選單 API 
        // (💡 修正 Route 名稱，加上 lpi- 前綴以完美對齊前端)
        // =========================================================================

        // LPI-1: 取得 Label Spec
        [HttpPost, Route("lpi-label-specs"), AuthorizeToken]
        public async Task<IHttpActionResult> GetLpiLabelSpecs([FromBody] dynamic req = null)
        {
            try
            {
                var result = await _printLabelService.GetLpiLabelSpecsAsync();
                return Ok(new { Success = true, Message = "", Data = result });
            }
            catch (Exception ex)
            {
                AppLogger.Error(this, ex.Message, ex);
                return Ok(new { Success = false, Message = "取得 LPI Label Spec 失敗" });
            }
        }

        // LPI-2: 取得 CarrierType
        [HttpPost, Route("lpi-carrier-types"), AuthorizeToken]
        public async Task<IHttpActionResult> GetLpiCarrierTypes([FromBody] DropdownRequest request)
        {
            if (request == null || string.IsNullOrWhiteSpace(request.LabelSpecNo)) return BadRequest("參數缺失");
            try
            {
                var result = await _printLabelService.GetLpiCarrierTypesAsync(request.LabelSpecNo);
                return Ok(new { Success = true, Message = "", Data = result });
            }
            catch (Exception ex)
            {
                AppLogger.Error(this, ex.Message, ex);
                return Ok(new { Success = false, Message = "取得 Carrier Type 失敗" });
            }
        }

        // LPI-3: 取得 BoxingSpecNo
        [HttpPost, Route("lpi-boxing-specs"), AuthorizeToken]
        public async Task<IHttpActionResult> GetLpiBoxingSpecs([FromBody] DropdownRequest request)
        {
            if (request == null || string.IsNullOrWhiteSpace(request.CarrierType)) return BadRequest("參數缺失");
            try
            {
                var result = await _printLabelService.GetLpiBoxingSpecsAsync(request);
                return Ok(new { Success = true, Message = "", Data = result });
            }
            catch (Exception ex)
            {
                AppLogger.Error(this, ex.Message, ex);
                return Ok(new { Success = false, Message = ex.Message });
            }
        }

        // LPI-4: 取得 Brand
        [HttpPost, Route("lpi-brands"), AuthorizeToken]
        public async Task<IHttpActionResult> GetLpiBrands([FromBody] DropdownRequest request)
        {
            if (request == null || string.IsNullOrWhiteSpace(request.BoxingSpecNo)) return BadRequest("參數缺失");
            try
            {
                var result = await _printLabelService.GetLpiBrandsAsync(request);
                return Ok(new { Success = true, Message = "", Data = result });
            }
            catch (Exception ex)
            {
                AppLogger.Error(this, ex.Message, ex);
                return Ok(new { Success = false, Message = ex.Message });
            }
        }

        // LPI-5: 取得 PinCount
        [HttpPost, Route("lpi-pin-counts"), AuthorizeToken]
        public async Task<IHttpActionResult> GetLpiPinCounts([FromBody] DropdownRequest request)
        {
            if (request == null || string.IsNullOrWhiteSpace(request.Brand)) return BadRequest("參數缺失");
            try
            {
                var result = await _printLabelService.GetLpiPinCountsAsync(request);
                return Ok(new { Success = true, Message = "", Data = result });
            }
            catch (Exception ex)
            {
                AppLogger.Error(this, ex.Message, ex);
                return Ok(new { Success = false, Message = ex.Message });
            }
        }

        // LPI-6: 取得 PackageCode
        [HttpPost, Route("package-codes"), AuthorizeToken]
        public async Task<IHttpActionResult> GetPackageCodes([FromBody] DropdownRequest request)
        {
            if (request == null || string.IsNullOrWhiteSpace(request.PinCount)) return BadRequest("參數缺失");
            try
            {
                var result = await _printLabelService.GetPackageCodesAsync(request);
                return Ok(new { Success = true, Message = "", Data = result });
            }
            catch (Exception ex)
            {
                AppLogger.Error(this, ex.Message, ex);
                return Ok(new { Success = false, Message = ex.Message });
            }
        }

        // =========================================================================
        // 執行列印
        // =========================================================================
        [HttpPost, Route("print"), AuthorizeToken]
        public async Task<IHttpActionResult> Print([FromBody] PrintLabelRequest request)
        {
            // 💡 移除舊版對於 Printer 的檢核，改為純檢核 PrinterServer
            if (request == null || string.IsNullOrWhiteSpace(request.LabelFormat) || string.IsNullOrWhiteSpace(request.PrinterServer))
            {
                return BadRequest("Pls input Query Criteria !! 請確認伺服器與標籤格式皆已選擇 !!");
            }
            try
            {
                await _printLabelService.ExecutePrintAsync(request);
                return Ok(new { Success = true, Message = "Label print request submitted successfully !! 標籤列印指令已成功送出 !!" });
            }
            catch (Exception ex)
            {
                AppLogger.Error(this, ex.Message, ex);
                AppLogger.Info(this, JsonConvert.SerializeObject(request));
                
                // 攔截特定商業邏輯例外 (例如: 要求原始批號、Hold 防呆)
                if (ex is InvalidOperationException || ex is ArgumentException)
                {
                    return Ok(new { Success = false, Message = ex.Message });
                }
                return Ok(new { Success = false, Message = "列印發生預期外的系統錯誤" });
            }
        }
    }
}

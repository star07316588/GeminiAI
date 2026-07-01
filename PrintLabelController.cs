
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

        // 1. 畫面初始化 (一次撈回不需連動的基礎下拉選單)
        [HttpPost, Route("init"), AuthorizeToken]
        public async Task<IHttpActionResult> Init()
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

        // 2. 依據 Printer Server 取得 Printer 清單
        [HttpPost, Route("printers"), AuthorizeToken]
        public async Task<IHttpActionResult> GetPrinters([FromBody] PrinterRequest request)
        {
            if (request == null || string.IsNullOrWhiteSpace(request.PrinterServer))
            {
                return BadRequest("請先選擇 Printer Server");
            }
            try
            {
                var result = await _printLabelService.GetPrintersAsync(request.PrinterServer);
                return Ok(new { Success = true, Message = "", Data = result });
            }
            catch (Exception ex)
            {
                AppLogger.Error(this, ex.Message, ex);
                return Ok(new { Success = false, Message = ex.Message });
            }
        }

        // 1. 初始化 / 取得 CarrierType 列表
        [HttpPost, Route("carrier-types"), AuthorizeToken]
        public async Task<IHttpActionResult> GetCarrierTypes()
        {
            try
            {
                var result = await _printLabelService.GetCarrierTypesAsync();
                return Ok(new { Success = true, Message = "", Data = result });
            }
            catch (Exception ex)
            {
                AppLogger.Error(this, ex.Message, ex);
                return Ok(new { Success = false, Message = "取得 Carrier Type 失敗" });
            }
        }

        // 2. 依據 CarrierType 取得 BoxingSpecNo
        [HttpPost, Route("boxing-specs"), AuthorizeToken]
        public async Task<IHttpActionResult> GetBoxingSpecs([FromBody] DropdownRequest request)
        {
            if (request == null || string.IsNullOrWhiteSpace(request.CarrierType))
            {
                return BadRequest("請先選擇 Carrier Type");
            }
            try
            {
                var result = await _printLabelService.GetBoxingSpecsAsync(request.CarrierType);
                return Ok(new { Success = true, Message = "", Data = result });
            }
            catch (Exception ex)
            {
                AppLogger.Error(this, ex.Message, ex);
                return Ok(new { Success = false, Message = ex.Message });
            }
        }

        // 3. 依據 CarrierType + BoxingSpecNo 取得 Brand
        [HttpPost, Route("brands"), AuthorizeToken]
        public async Task<IHttpActionResult> GetBrands([FromBody] DropdownRequest request)
        {
            if (request == null || string.IsNullOrWhiteSpace(request.CarrierType) || string.IsNullOrWhiteSpace(request.BoxingSpecNo))
            {
                return BadRequest("參數缺失，無法取得 Brand");
            }
            try
            {
                var result = await _printLabelService.GetBrandsAsync(request);
                return Ok(new { Success = true, Message = "", Data = result });
            }
            catch (Exception ex)
            {
                AppLogger.Error(this, ex.Message, ex);
                return Ok(new { Success = false, Message = ex.Message });
            }
        }

        // 4. 依據前三項條件取得 PinCount
        [HttpPost, Route("pin-counts"), AuthorizeToken]
        public async Task<IHttpActionResult> GetPinCounts([FromBody] DropdownRequest request)
        {
            if (request == null || string.IsNullOrWhiteSpace(request.CarrierType) || 
                string.IsNullOrWhiteSpace(request.BoxingSpecNo) || string.IsNullOrWhiteSpace(request.Brand))
            {
                return BadRequest("參數缺失，無法取得 Pin Count");
            }
            try
            {
                var result = await _printLabelService.GetPinCountsAsync(request);
                return Ok(new { Success = true, Message = "", Data = result });
            }
            catch (Exception ex)
            {
                AppLogger.Error(this, ex.Message, ex);
                return Ok(new { Success = false, Message = ex.Message });
            }
        }

        [HttpPost, Route("print"), AuthorizeToken]
        public async Task<IHttpActionResult> Print([FromBody] PrintLabelRequest request)
        {
            if (request == null || string.IsNullOrWhiteSpace(request.LabelFormat) || string.IsNullOrWhiteSpace(request.Printer))
            {
                return BadRequest("Pls input Query Criteria !! 請確認印表機與標籤格式皆已選擇 !!");
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
                // 攔截特定商業邏輯例外 (例如: 印表機離線、防呆檢核失敗)
                if (ex is InvalidOperationException || ex is ArgumentException)
                {
                    return Ok(new { Success = false, Message = ex.Message });
                }
                return Ok(new { Success = false, Message = "列印發生預期外的系統錯誤" });
            }
        }
    }
}

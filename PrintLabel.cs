namespace MES.Net.Shared.DTOs.PrintLabel
{
    // 1. 取得下拉選單的請求參數 (可共用，視查詢層級帶入不同參數)
    public class DropdownRequest
    {
        public string CarrierType { get; set; }
        public string BoxingSpecNo { get; set; }
        public string Brand { get; set; }
    }

    // 2. 執行列印(或查詢標籤資料)的 Request Payload
    public class PrintLabelRequest
    {
        public string LotId { get; set; } // 假設有批號
        public string CarrierType { get; set; }
        public string BoxingSpecNo { get; set; }
        public string Brand { get; set; }
        public string PinCount { get; set; }
        public string UserId { get; set; } 
        // ... 其他畫面上需要送出的參數
    }
}

using System.Web.Http;
using System.Threading.Tasks;
using MES.Net.Shared.DTOs.PrintLabel;

namespace MES.Net.Web.Controllers
{
    [RoutePrefix("api/print-label")]
    public class PrintLabelController : ApiController
    {
        private readonly IPrintLabelService _printLabelService;

        public PrintLabelController(IPrintLabelService printLabelService)
        {
            _printLabelService = printLabelService;
        }

        // 1. 取得 CarrierType 列表 (通常在畫面初始化時呼叫)
        [HttpGet, Route("carrier-types")]
        public async Task<IHttpActionResult> GetCarrierTypes()
        {
            var result = await _printLabelService.GetCarrierTypesAsync();
            return Ok(new { Success = true, Data = result });
        }

        // 2. 依據 CarrierType 取得 BoxingSpecNo 列表
        [HttpGet, Route("boxing-specs")]
        public async Task<IHttpActionResult> GetBoxingSpecs([FromUri] string carrierType)
        {
            var result = await _printLabelService.GetBoxingSpecsAsync(carrierType);
            return Ok(new { Success = true, Data = result });
        }

        // 3. 依據 CarrierType 與 BoxingSpecNo 取得 Brand 列表
        [HttpGet, Route("brands")]
        public async Task<IHttpActionResult> GetBrands([FromUri] DropdownRequest request)
        {
            var result = await _printLabelService.GetBrandsAsync(request.CarrierType, request.BoxingSpecNo);
            return Ok(new { Success = true, Data = result });
        }

        // 4. 依據前三個條件取得 PinCount 列表
        [HttpGet, Route("pin-counts")]
        public async Task<IHttpActionResult> GetPinCounts([FromUri] DropdownRequest request)
        {
            var result = await _printLabelService.GetPinCountsAsync(request.CarrierType, request.BoxingSpecNo, request.Brand);
            return Ok(new { Success = true, Data = result });
        }

        // 5. 執行列印動作 (可能包含更新資料庫狀態，或是回傳 ZPL 條碼指令)
        [HttpPost, Route("print")]
        public async Task<IHttpActionResult> PrintLabel([FromBody] PrintLabelRequest request)
        {
            if (request == null) return BadRequest("無效的請求資料");

            var result = await _printLabelService.PrintLabelAsync(request);
            return Ok(new { Success = true, Message = "標籤列印成功", Data = result });
        }
    }
}

using System.Collections.Generic;
using System.Threading.Tasks;

namespace MES.Net.Application.Services
{
    public interface IPrintLabelService
    {
        Task<IEnumerable<string>> GetCarrierTypesAsync();
        Task<IEnumerable<string>> GetBoxingSpecsAsync(string carrierType);
        Task<IEnumerable<string>> GetBrandsAsync(string carrierType, string boxingSpecNo);
        Task<IEnumerable<string>> GetPinCountsAsync(string carrierType, string boxingSpecNo, string brand);
        Task<object> PrintLabelAsync(PrintLabelRequest request); // 依實際列印回傳格式決定回傳型別
    }

    public class PrintLabelService : IPrintLabelService
    {
        private readonly IPrintLabelRepository _repo;

        public PrintLabelService(IPrintLabelRepository repo)
        {
            _repo = repo;
        }

        public async Task<IEnumerable<string>> GetBoxingSpecsAsync(string carrierType)
        {
            if (string.IsNullOrWhiteSpace(carrierType)) return new List<string>();
            return await _repo.GetBoxingSpecsAsync(carrierType);
        }

        // ... 實作其他連動選單方法 ...

        public async Task<object> PrintLabelAsync(PrintLabelRequest request)
        {
            // 這裡實作 VB6 中按下 "Print" 按鈕後的邏輯
            // 例如：檢核資料、寫入 Log、產出給印表機的 ZPL/PDF 檔案格式等
            
            return new { /* 列印結果或指令 */ };
        }
    }
}

using Dapper;
using System.Collections.Generic;
using System.Data;
using System.Threading.Tasks;

namespace MES.Net.Infrastructure.Repository
{
    public interface IPrintLabelRepository
    {
        Task<IEnumerable<string>> GetCarrierTypesAsync();
        Task<IEnumerable<string>> GetBoxingSpecsAsync(string carrierType);
        Task<IEnumerable<string>> GetBrandsAsync(string carrierType, string boxingSpecNo);
        Task<IEnumerable<string>> GetPinCountsAsync(string carrierType, string boxingSpecNo, string brand);
    }

    public class PrintLabelRepository : IPrintLabelRepository
    {
        private readonly IDbConnection _dbConnection;

        public PrintLabelRepository(IDbConnection dbConnection)
        {
            _dbConnection = dbConnection;
        }

        public async Task<IEnumerable<string>> GetBoxingSpecsAsync(string carrierType)
        {
            // 假設資料存在 TBL_LABEL_SPEC (請替換為實際的 Table 與欄位名稱)
            string sql = @"
                SELECT DISTINCT BOXING_SPEC_NO 
                FROM TBL_LABEL_SPEC 
                WHERE CARRIER_TYPE = :CarrierType 
                  AND DELETE_FLAG = 'N'
                ORDER BY BOXING_SPEC_NO";

            return await _dbConnection.QueryAsync<string>(sql, new { CarrierType = carrierType });
        }

        public async Task<IEnumerable<string>> GetBrandsAsync(string carrierType, string boxingSpecNo)
        {
            string sql = @"
                SELECT DISTINCT BRAND 
                FROM TBL_LABEL_SPEC 
                WHERE CARRIER_TYPE = :CarrierType 
                  AND BOXING_SPEC_NO = :BoxingSpecNo
                ORDER BY BRAND";

            return await _dbConnection.QueryAsync<string>(sql, new 
            { 
                CarrierType = carrierType, 
                BoxingSpecNo = boxingSpecNo 
            });
        }
        
        // ... 實作 GetPinCountsAsync 等 ...
    }
}

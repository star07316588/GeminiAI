using System.Collections.Generic;

namespace MES.Net.Shared.DTOs.Print
{
    // 連動下拉選單共用的請求物件
    public class DropdownRequest
    {
        public string CarrierType { get; set; }
        public string BoxingSpecNo { get; set; }
        public string Brand { get; set; }
    }

    // 點擊 Print 按鈕時的請求物件
    public class PrintLabelRequest
    {
        public string LotId { get; set; }
        public string CarrierType { get; set; }
        public string BoxingSpecNo { get; set; }
        public string Brand { get; set; }
        public string PinCount { get; set; }
        public string UserId { get; set; }
    }
}
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

        // 5. 執行標籤列印動作
        [HttpPost, Route("print"), AuthorizeToken]
        public async Task<IHttpActionResult> Print([FromBody] PrintLabelRequest request)
        {
            if (request == null || string.IsNullOrWhiteSpace(request.CarrierType) || string.IsNullOrWhiteSpace(request.BoxingSpecNo))
            {
                return BadRequest("Pls input Query Criteria !! 請輸入完整的列印條件 !!");
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
                return Ok(new { Success = false, Message = ex.Message });
            }
        }
    }
}

using MES.Net.Infrastructure.Repository.Print;
using MES.Net.Shared.DTOs.Print;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace MES.Net.Application.Services.Print
{
    public interface IPrintLabelService
    {
        Task<IEnumerable<string>> GetCarrierTypesAsync();
        Task<IEnumerable<string>> GetBoxingSpecsAsync(string carrierType);
        Task<IEnumerable<string>> GetBrandsAsync(DropdownRequest request);
        Task<IEnumerable<string>> GetPinCountsAsync(DropdownRequest request);
        Task ExecutePrintAsync(PrintLabelRequest request);
    }

    public class PrintLabelService : IPrintLabelService
    {
        private readonly IPrintLabelRepository _printLabelRepo;

        public PrintLabelService(IPrintLabelRepository printLabelRepo)
        {
            _printLabelRepo = printLabelRepo;
        }

        public async Task<IEnumerable<string>> GetCarrierTypesAsync()
        {
            return await _printLabelRepo.GetCarrierTypesAsync();
        }

        public async Task<IEnumerable<string>> GetBoxingSpecsAsync(string carrierType)
        {
            return await _printLabelRepo.GetBoxingSpecsAsync(carrierType);
        }

        public async Task<IEnumerable<string>> GetBrandsAsync(DropdownRequest request)
        {
            return await _printLabelRepo.GetBrandsAsync(request.CarrierType, request.BoxingSpecNo);
        }

        public async Task<IEnumerable<string>> GetPinCountsAsync(DropdownRequest request)
        {
            return await _printLabelRepo.GetPinCountsAsync(request.CarrierType, request.BoxingSpecNo, request.Brand);
        }

        public async Task ExecutePrintAsync(PrintLabelRequest request)
        {
            // 可在此加入廠內商業邏輯驗證（例如確認帳號列印權限等）
            bool printSuccess = await _printLabelRepo.SpPrintLabelActionAsync(request);
            
            if (!printSuccess)
            {
                throw new InvalidOperationException("Failed to generate label file or printer offline !! 無法產出標籤檔案或印表機連線中斷 !!");
            }
        }
    }
}

using Dapper;
using MES.Net.Shared.DTOs.Print;
using System.Collections.Generic;
using System.Data;
using System.Threading.Tasks;

namespace MES.Net.Infrastructure.Repository.Print
{
    public interface IPrintLabelRepository
    {
        Task<IEnumerable<string>> GetCarrierTypesAsync();
        Task<IEnumerable<string>> GetBoxingSpecsAsync(string carrierType);
        Task<IEnumerable<string>> GetBrandsAsync(string carrierType, string boxingSpecNo);
        Task<IEnumerable<string>> GetPinCountsAsync(string carrierType, string boxingSpecNo, string brand);
        Task<bool> SpPrintLabelActionAsync(PrintLabelRequest request);
    }

    public class PrintLabelRepository : IPrintLabelRepository
    {
        private readonly IDbConnection _dbConnection;

        public PrintLabelRepository(IDbConnection dbConnection)
        {
            _dbConnection = dbConnection;
        }

        public async Task<IEnumerable<string>> GetCarrierTypesAsync()
        {
            string sql = @"SELECT DISTINCT CARRIER_TYPE FROM TBL_PACK_SPEC_MSTR WHERE DELETE_FLAG = 'N' ORDER BY CARRIER_TYPE";
            return await _dbConnection.QueryAsync<string>(sql);
        }

        public async Task<IEnumerable<string>> GetBoxingSpecsAsync(string carrierType)
        {
            string sql = @"
                SELECT DISTINCT BOXING_SPEC_NO 
                FROM TBL_PACK_SPEC_MSTR 
                WHERE CARRIER_TYPE = :p_CarrierType 
                  AND DELETE_FLAG = 'N' 
                ORDER BY BOXING_SPEC_NO";
            return await _dbConnection.QueryAsync<string>(sql, new { p_CarrierType = carrierType });
        }

        public async Task<IEnumerable<string>> GetBrandsAsync(string carrierType, string boxingSpecNo)
        {
            string sql = @"
                SELECT DISTINCT BRAND 
                FROM TBL_PACK_SPEC_MSTR 
                WHERE CARRIER_TYPE = :p_CarrierType 
                  AND BOXING_SPEC_NO = :p_BoxingSpecNo 
                  AND DELETE_FLAG = 'N' 
                ORDER BY BRAND";
            return await _dbConnection.QueryAsync<string>(sql, new { p_CarrierType = carrierType, p_BoxingSpecNo = boxingSpecNo });
        }

        public async Task<IEnumerable<string>> GetPinCountsAsync(string carrierType, string boxingSpecNo, string brand)
        {
            string sql = @"
                SELECT DISTINCT PIN_COUNT 
                FROM TBL_PACK_SPEC_MSTR 
                WHERE CARRIER_TYPE = :p_CarrierType 
                  AND BOXING_SPEC_NO = :p_BoxingSpecNo 
                  AND BRAND = :p_Brand 
                  AND DELETE_FLAG = 'N' 
                ORDER BY PIN_COUNT";
            return await _dbConnection.QueryAsync<string>(sql, new { p_CarrierType = carrierType, p_BoxingSpecNo = boxingSpecNo, p_Brand = brand });
        }

        public async Task<bool> SpPrintLabelActionAsync(PrintLabelRequest request)
        {
            var p = new DynamicParameters();
            p.Add("p_LotId", request.LotId);
            p.Add("p_CarrierType", request.CarrierType);
            p.Add("p_BoxingSpecNo", request.BoxingSpecNo);
            p.Add("p_Brand", request.Brand);
            p.Add("p_PinCount", request.PinCount);
            p.Add("p_UserId", request.UserId);
            p.Add("p_Result", dbType: DbType.Int32, direction: ParameterDirection.Output);

            await _dbConnection.ExecuteAsync("SP_EXEC_PRINT_LABEL", p, commandType: CommandType.StoredProcedure);

            int result = p.Get<int>("p_Result");
            return result == 0;
        }
    }
}

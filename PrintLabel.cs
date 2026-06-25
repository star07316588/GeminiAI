using System.Collections.Generic;

namespace MES.Net.Shared.DTOs.Print
{
    // 畫面初始化回傳的基礎資料
    public class PrintLabelInitResponse
    {
        public IEnumerable<string> Stages { get; set; }
        public IEnumerable<string> LabelTypes { get; set; }
        public IEnumerable<string> LabelFormats { get; set; }
        public IEnumerable<string> PrinterServers { get; set; }
        public IEnumerable<string> CarrierTypes { get; set; }
    }

    // 連動下拉選單共用的請求物件
    public class DropdownRequest
    {
        public string CarrierType { get; set; }
        public string BoxingSpecNo { get; set; }
        public string Brand { get; set; }
    }

    // 印表機伺服器連動請求物件
    public class PrinterRequest
    {
        public string PrinterServer { get; set; }
    }

    // 點擊 Print 按鈕時的完整請求物件 (對應前端的 payload)
    public class PrintLabelRequest
    {
        public string Stage { get; set; }
        public string LabelType { get; set; }
        public string LabelFormat { get; set; }
        public string LotId { get; set; }           // 對應前端的 labelData
        public string PrintMode { get; set; }
        public string PrinterServer { get; set; }
        public string Printer { get; set; }
        public int BoxQty { get; set; }
        public int PrintQty { get; set; }
        
        public string ReprintType { get; set; }
        public string SearchData { get; set; }

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

using MES.Net.Infrastructure.Repository.Print;
using MES.Net.Shared.DTOs.Print;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace MES.Net.Application.Services.Print
{
    public interface IPrintLabelService
    {
        Task<PrintLabelInitResponse> GetInitDataAsync();
        Task<IEnumerable<string>> GetPrintersAsync(string server);
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

        public async Task<PrintLabelInitResponse> GetInitDataAsync()
        {
            return await _printLabelRepo.GetInitDataAsync();
        }

        public async Task<IEnumerable<string>> GetPrintersAsync(string server)
        {
            return await _printLabelRepo.GetPrintersAsync(server);
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
            // 💡 商業邏輯檢核範例
            if (request.PrintMode == "Reprint" && request.ReprintType == "SearchData" && string.IsNullOrWhiteSpace(request.SearchData))
            {
                throw new ArgumentException("Please enter Search Data for Reprint !! 補印模式請輸入搜尋條件 !!");
            }

            // 呼叫 DB 執行 SP 產生列印資料或 ZPL 檔
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
        Task<PrintLabelInitResponse> GetInitDataAsync();
        Task<IEnumerable<string>> GetPrintersAsync(string server);
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

        public async Task<PrintLabelInitResponse> GetInitDataAsync()
        {
            // 實務上可使用 Dapper 的 QueryMultiple 或是分開查詢系統代碼表 (TBL_SYS_CODE)
            var response = new PrintLabelInitResponse();

            response.Stages = await _dbConnection.QueryAsync<string>("SELECT DISTINCT STAGE FROM TBL_STAGE_DEF ORDER BY STAGE");
            response.LabelTypes = await _dbConnection.QueryAsync<string>("SELECT CODE_DESC FROM TBL_SYS_CODE WHERE CODE_TYPE = 'LABEL_TYPE' ORDER BY CODE_ID");
            response.LabelFormats = await _dbConnection.QueryAsync<string>("SELECT DISTINCT FORMAT_NAME FROM TBL_LABEL_FORMAT ORDER BY FORMAT_NAME");
            response.PrinterServers = await _dbConnection.QueryAsync<string>("SELECT DISTINCT SERVER_NAME FROM TBL_PRINTER_MSTR ORDER BY SERVER_NAME");
            
            // 順便把第一層的 Carrier Type 也查回來，前端就不用多打一次 API
            response.CarrierTypes = await _dbConnection.QueryAsync<string>("SELECT DISTINCT CARRIER_TYPE FROM TBL_PACK_SPEC_MSTR WHERE DELETE_FLAG = 'N' ORDER BY CARRIER_TYPE");

            return response;
        }

        public async Task<IEnumerable<string>> GetPrintersAsync(string server)
        {
            string sql = @"
                SELECT PRINTER_NAME 
                FROM TBL_PRINTER_MSTR 
                WHERE SERVER_NAME = :p_ServerName 
                  AND ENABLE_FLAG = 'Y' 
                ORDER BY PRINTER_NAME";
            
            return await _dbConnection.QueryAsync<string>(sql, new { p_ServerName = server });
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
            
            // 寫入完整參數，供 DB 預存程序或列印伺服器判斷
            p.Add("p_Stage", request.Stage);
            p.Add("p_LabelType", request.LabelType);
            p.Add("p_LabelFormat", request.LabelFormat);
            p.Add("p_LotId", request.LotId);
            p.Add("p_PrintMode", request.PrintMode);
            p.Add("p_PrinterServer", request.PrinterServer);
            p.Add("p_Printer", request.Printer);
            p.Add("p_BoxQty", request.BoxQty);
            p.Add("p_PrintQty", request.PrintQty);
            
            p.Add("p_ReprintType", request.ReprintType);
            p.Add("p_SearchData", request.SearchData);

            p.Add("p_CarrierType", request.CarrierType);
            p.Add("p_BoxingSpecNo", request.BoxingSpecNo);
            p.Add("p_Brand", request.Brand);
            p.Add("p_PinCount", request.PinCount);
            p.Add("p_UserId", request.UserId);
            
            p.Add("p_Result", dbType: DbType.Int32, direction: ParameterDirection.Output);

            await _dbConnection.ExecuteAsync("SP_EXEC_PRINT_LABEL", p, commandType: CommandType.StoredProcedure);

            int result = p.Get<int>("p_Result");
            return result == 0; // 假設 0 代表成功
        }
    }
}

namespace MES.Net.Infrastructure.Printing
{
    /// <summary>
    /// 專屬存放所有 ZPL 標籤模板的統一集中區
    /// </summary>
    public static class ZplTemplates
    {
        // 1. FT_To_SubTR_Normal 標籤模板
        // 注意：使用 C# 的 @ 逐字字串符號，可以讓 ZPL 保持原本的換行排版，極大提升可讀性
        public static readonly string FT_To_SubTR_Normal = @"
^XA^LH00,00^LL550
^FO10,05^GB780,475,3^FS
^FO435,327^GB355,70,3^FS
^FO435,397^GB0,80,3^FS
^FO580,30^GB200,75,3^FS
^FO640,55^A0N,35,50^CI0^FDT&R^FS
{PartialBlock}
^FO40,58^A0N,33,28^FD(1T)LOT NO: {LotNo}^FS
^FO40,158^A0N,33,28^FD(1P)PRODUCT NO: {ProdNo}^FS
^FO40,258^A0N,33,28^FD(Q)QUANTITY: {Qty}^FS
^FO40,90^BY2,2,50^B3,,,N^FD1T{LotNo}^FS
^FO40,190^BY2,2,50^B3,,,N^FD1P{ProdNo}^FS
^FO40,290^BY2,2,50^B3,,,N^FDQ{Qty}^FS
^FO445,350^A0N,36,30^FDPACKER: {Packer}^FS
^FO445,420^A0N,36,30^FDQA:^FS
^PQ1
^XZ1";

        // 這是 Partial 時要額外插入的區塊
        public static readonly string PartialBlock = @"^FO580,102^GB200,75,3^FS
^FO590,130^A0N,35,50^CI0^FDPARTIAL^FS";

        // 💡 新增：FT_To_SubFT_Normal 標籤模板
        public static readonly string FT_To_SubFT_Normal = @"
^XA^LH00,00^LL550
^FO10,05^GB780,475,3^FS
^FO435,327^GB355,70,3^FS
^FO435,397^GB0,80,3^FS
^FO580,30^GB200,75,3^FS
^FO635,55^A0N,35,50^CI0^FDTEST^FS
{PartialBlock}
^FO40,58^A0N,33,28^FD(1T)LOT NO: {LotNo}^FS
^FO40,158^A0N,33,28^FD(1P)PRODUCT NO: {ProdNo}^FS
^FO40,258^A0N,33,28^FD(Q)QUANTITY: {Qty}^FS
^FO40,90^BY2,2,50^B3,,,N^FD1T{LotNo}^FS
^FO40,190^BY2,2,50^B3,,,N^FD1P{ProdNo}^FS
^FO40,290^BY2,2,50^B3,,,N^FDQ{Qty}^FS
^FO445,350^A0N,36,30^FDPACKER: {Packer}^FS
^FO445,420^A0N,36,30^FDQA:^FS
^PQ1
^XZ1";

        // 1. WS_To_SFG 標籤模板
        public static readonly string WS_WS_To_SFG = @"
^XA^LH10,00^LL500
^FO00,05^GB780,475,3^FS
^FO30,20^A0,27,25^FDLOT NO: {LotNo}^FS
^FO30,130^A0,27,25^FDPRODUCT NO: {ProdNo}^FS
^FO30,235^A0,27,25^FDWAFER PCS:  {WQty}PCS^FS
^FO30,340^A0,27,25^FDQUANTITY:  {Qty}EA^FS
^FO30,445^A0,27,25^FDPART NO: {PartNo}^FS
^FO30,55^BY2,2,55^B3,,,N^FD{LotNo}^FS
^FO30,160^BY2,2,55^B3,,,N^FD{ProdNo}^FS
^FO30,265^BY2,2,55^B3,,,N^FD{WQty}^FS
^FO30,370^BY2,2,55^B3,,,N^FD{Qty}^FS
^FO425,385^GB355,50,3^FS
^FO425,425^GB0,50,3^FS
^FO440,400^A0,27,25^FDDATE CODE:^FS
^FO440,440^A0,27,25^FDQA:^FS
^PQ1
^XZ";

        // 2. WS_SUMMARY 標籤模板 (包含動態的 Component ID 列表)
        public static readonly string WS_SUMMARY = @"
^XA
^LH0,0^FS
^LL440
^MD0
^MNY
^FO368,36^A0N,40.32^CI0^FR^FDLOT : {LotNo}^FS
^FO48,10^A0B,30,30^CI0^FR^FDLOT : {LotNo}^FS
^BY2,2.0^FO80,50^B3B,N,60,N,Y^FR^FD{LotNo}^FS
^FO775,36^A0N,40,32^CI0^FR^FDPROD : {ProdNo}^FS
^FO160,10^A0B,30,32^CI0^FR^FDPROD : {ProdNo}^FS
^BY1, 3.0^FO192,50^B3B,N,60,N,Y^FR^FD{ProdNo}^FS
{CompIdBlocks}
^FO250,318^A0N,30,32^CI0^FR^FDTTL WQTY: {WQty} PCS ^FS
^BY2,2.0^FO540,318^B3N,N,40,N,Y^FR^FD{WQty}^FS
^FO250,370^A0N,30,32^CI0^FR^FDTTL CQTY: {CQty}EA ^FS
^BY2,2.0^FO540,370^B3N,N,40,N,Y^FR^FD{CQty}^FS
^FO690,318^A0N,30,32^CI0^FR^FDID No: {UserName}^FS
^FO980,318^A0N,30,32^CI0^FR^FDTQA:        ACC^FS
^FO790,370^A0N,30,32^CI0^FR^FDDATE: {TimeStamp}^FS
^FO1070,284^GB190,140,4^FS
^XZ";

        // 3. FT_SMALL_LABEL 標籤模板群組
        // 3-1. Location == FT 且 有 OriLotID
        public static readonly string FT_SMALL_LABEL_FT_WithOriLot = @"
^XA^PRC^LH0,0^FS^LL280^MD0^LH0,0^FS
^FO30,5^A0N,35,27^CI13^FR^FDCarGradeFlag:{CargradeFlag}^FS
^FO315,5^A0N,35,27^CI13^FR^FDFuSa:{FuSa}^FS
^FO30,45^A0N,35,27^CI13^FR^FDLot:{LotNo}^FS
^FO315,45^A0N,35,27^CI13^FR^FD{ProdLine}^FS
^BY2,3.0^FO30,85^BCN,40,N,Y,N^FR^FD>:{LotNo}^FS
^FO30,138^A0N,35,27^CI13^FR^FD{ProdNo}^FS
^FO400,85^A0N,35,27^CI13^FR^FDGreen:{Green}^FS
^FO30,178^A0N,35,27^CI13^FR^FD{CQty} EA^FS
^FO265,178^A0N,35,27^CI13^FR^FDRoute:{RouteId}^FS
^FO30,215^A0N,35,27^CI13^FR^FDOrg_Lotid:{OriLotID}^FS
^PQ1,1,1,Y^XZ";

        // 3-2. Location == FT 且 無 OriLotID
        public static readonly string FT_SMALL_LABEL_FT_NoOriLot = @"
^XA^PRC^LH0,0^FS^LL280^MD0^LH0,0^FS
^FO30,11^A0N,37,33^CI13^FR^FDCarGradeFlag:{CargradeFlag}^FS
^FO315,11^A0N,37,33^CI13^FR^FDFuSa:{FuSa}^FS
^FO30,55^A0N,37,33^CI13^FR^FDLot:{LotNo}^FS
^FO315,55^A0N,37,33^CI13^FR^FD{ProdLine}^FS
^BY2,3.0^FO30,95^BCN,50,N,Y,N^FR^FD>:{LotNo}^FS
^FO30,161^A0N,37,33^CI13^FR^FD{ProdNo}^FS
^FO400,95^A0N,37,33^CI13^FR^FDGreen:{Green}^FS
^FO30,206^A0N,37,33^CI13^FR^FD{CQty} EA^FS
^FO265,206^A0N,37,33^CI13^FR^FDRoute:{RouteId}^FS
^PQ1,1,1,Y^XZ";

        // 3-3. Location != FT 且 有 OriLotID
        public static readonly string FT_SMALL_LABEL_NonFT_WithOriLot = @"
^XA^PRC^LH0,0^FS^LL280^MD0^LH0,0^FS
^FO30,5^A0N,35,27^CI13^FR^FDCarGradeFlag:{CargradeFlag}^FS
^FO315,5^A0N,35,27^CI13^FR^FD{ProdLine}^FS
^FO30,45^A0N,35,27^CI13^FR^FDLot:{LotNo}^FS
^FO315,45^A0N,35,27^CI13^FR^FDOwner:{Owner}^FS
^BY2,3.0^FO30,85^BCN,40,N,Y,N^FR^FD>:{LotNo}^FS
^FO350,91^A0N,35,27^CI13^FR^FDCsum:{Csum}^FS
^FO30,138^A0N,35,27^CI13^FR^FD{ProdNo}^FS
^FO400,138^A0N,35,27^CI13^FR^FDGreen:{Green}^FS
^FO30,178^A0N,35,27^CI13^FR^FD{CQty} EA^FS
^FO265,178^A0N,35,27^CI13^FR^FDRoute:{RouteId}^FS
^FO30,215^A0N,35,27^CI13^FR^FDOrg_Lotid:{OriLotID}^FS
^PQ1,1,1,Y^XZ";

        // 3-4. Location != FT 且 無 OriLotID
        public static readonly string FT_SMALL_LABEL_NonFT_NoOriLot = @"
^XA^PRC^LH0,0^FS^LL280^MD0^LH0,0^FS
^FO30,11^A0N,37,33^CI13^FR^FDCarGradeFlag:{CargradeFlag}^FS
^FO315,11^A0N,37,33^CI13^FR^FD{ProdLine}^FS
^FO30,55^A0N,37,33^CI13^FR^FDLot:{LotNo}^FS
^FO315,55^A0N,37,33^CI13^FR^FDOwner:{Owner}^FS
^BY2,3.0^FO30,95^BCN,50,N,Y,N^FR^FD>:{LotNo}^FS
^FO350,106^A0N,37,33^CI13^FR^FDCsum:{Csum}^FS
^FO30,161^A0N,37,33^CI13^FR^FD{ProdNo}^FS
^FO400,161^A0N,37,33^CI13^FR^FDGreen:{Green}^FS
^FO30,206^A0N,37,33^CI13^FR^FD{CQty} EA^FS
^FO265,206^A0N,37,33^CI13^FR^FDRoute:{RouteId}^FS
^PQ1,1,1,Y^XZ";
}
using System.Threading.Tasks;

namespace MES.Net.Infrastructure.Messaging
{
    public interface IMessageQueueService
    {
        /// <summary>
        /// 對應原本的 moCwMbx.Send
        /// </summary>
        Task<bool> SendMessageAsync(string queueName, string message);
    }
    
    // 實作細節會依據您廠內現在是用 RabbitMQ, MSMQ 還是 Socket 直連來寫
    public class MessageBoxService : IMessageQueueService
    {
        public async Task<bool> SendMessageAsync(string queueName, string message)
        {
            // TODO: 實作您廠內的佇列或 Socket 傳輸機制
            return true; 
        }
    }
}

using MES.Net.Infrastructure.Printing;
using MES.Net.Infrastructure.Messaging;
using System;
using System.Threading.Tasks;

namespace MES.Net.Application.Services.Print
{
    public class PrintLabelService : IPrintLabelService
    {
        private readonly IMessageQueueService _mqService;

        public PrintLabelService(IMessageQueueService mqService)
        {
            _mqService = mqService;
        }

        /// <summary>
        /// 完美翻寫 VB6 的 Prt_FT_To_SubTR_Normal
        /// </summary>
        public async Task Prt_FT_To_SubTR_Normal_Async(string lotNo, string prodNo, string qty, string packer, bool isPartial, string printerServer)
        {
            try
            {
                // 1. 取得 ZPL 模板
                string zpl = ZplTemplates.FT_To_SubTR_Normal;

                // 2. 處理條件判斷 (取代 {PartialBlock})
                string partialZpl = isPartial ? ZplTemplates.PartialBlock : "";
                zpl = zpl.Replace("{PartialBlock}", partialZpl);

                // 3. 變數綁定與取代
                zpl = zpl.Replace("{LotNo}", lotNo)
                         .Replace("{ProdNo}", prodNo)
                         .Replace("{Qty}", qty)
                         .Replace("{Packer}", packer);

                // 4. 處理 Printer Server 字串過濾 (對應 VB6 的 InStr 判斷)
                // If InStr(1, sPrinterServer, "@") <> 0 Then ...
                string parsedServer = printerServer;
                if (parsedServer.Contains("@"))
                {
                    parsedServer = parsedServer.Substring(0, parsedServer.IndexOf("@"));
                }

                // 5. 組合 Queue 名稱並發送 (對應 moCwMbx.Send)
                string queueName = $"MBX_{parsedServer}";
                
                bool isSent = await _mqService.SendMessageAsync(queueName, zpl);

                if (!isSent)
                {
                    throw new Exception("發送至印表機佇列失敗！");
                }
            }
            catch (Exception ex)
            {
                // 錯誤捕捉對應原本的 ErrorHandler
                throw new InvalidOperationException($"Fail to execute application, please call IT support!! 程式執行失敗: {ex.Message}", ex);
            }
        }
        /// <summary>
        /// 完美翻寫 VB6 的 Prt_FT_To_SubFT_Normal
        /// </summary>
        public async Task Prt_FT_To_SubFT_Normal_Async(string lotNo, string prodNo, string qty, string packer, bool isPartial, string printerServer)
        {
            try
            {
                // 1. 取得 ZPL 模板 (這次取用 FT_To_SubFT_Normal)
                string zpl = ZplTemplates.FT_To_SubFT_Normal;

                // 2. 處理條件判斷 (與 TR 標籤完美共用同一個 PartialBlock 字串)
                string partialZpl = isPartial ? ZplTemplates.PartialBlock : "";
                zpl = zpl.Replace("{PartialBlock}", partialZpl);

                // 3. 變數綁定與取代
                zpl = zpl.Replace("{LotNo}", lotNo)
                         .Replace("{ProdNo}", prodNo)
                         .Replace("{Qty}", qty)
                         .Replace("{Packer}", packer);

                // 4. 處理 Printer Server 字串過濾 (對應 VB6 的 InStr 判斷)
                string parsedServer = printerServer;
                if (parsedServer.Contains("@"))
                {
                    parsedServer = parsedServer.Substring(0, parsedServer.IndexOf("@"));
                }

                // 5. 組合 Queue 名稱並發送
                string queueName = $"MBX_{parsedServer}";
                
                bool isSent = await _mqService.SendMessageAsync(queueName, zpl);

                if (!isSent)
                {
                    throw new Exception("發送至印表機佇列失敗！");
                }
            }
            catch (Exception ex)
            {
                // 統一格式的錯誤捕捉
                throw new InvalidOperationException($"Fail to execute application, please call IT support!! 程式執行失敗: {ex.Message}", ex);
            }
        }
        // --- 共用的發送方法 ---
        private async Task SendToPrinterAsync(string printerServer, string zplCommand)
        {
            string parsedServer = printerServer.Contains("@") 
                ? printerServer.Substring(0, printerServer.IndexOf("@")) 
                : printerServer;

            string queueName = $"MBX_{parsedServer}";
            bool isSent = await _mqService.SendMessageAsync(queueName, zplCommand);
            if (!isSent) throw new Exception("發送至印表機佇列失敗！");
        }


        /// <summary>
        /// 1. 翻寫 Prt_WS_WS_To_SFG
        /// </summary>
        public async Task Prt_WS_WS_To_SFG_Async(string lotNo, string prodNo, string wQty, string qty, string partNo, string printerServer)
        {
            string zpl = ZplTemplates.WS_WS_To_SFG
                .Replace("{LotNo}", lotNo)
                .Replace("{ProdNo}", prodNo)
                .Replace("{WQty}", wQty)
                .Replace("{Qty}", qty)
                .Replace("{PartNo}", partNo);

            await SendToPrinterAsync(printerServer, zpl);
        }


        /// <summary>
        /// 2. 翻寫 Prt_WS_WS_SUMMARY (含動態 Component ID 座標運算)
        /// </summary>
        public async Task Prt_WS_WS_SUMMARY_Async(string lotNo, string prodNo, string wQty, string cQty, Dictionary<string, string> compIds, string userName, string timeStamp, string printerServer)
        {
            // 動態產生 25 個 Component ID 的 ZPL 座標字串
            var sbComps = new StringBuilder();
            
            // 定義 VB6 中寫死的 X 座標陣列 (依據第1, 2, 3, 4, 5列)
            int[] xCoords = { 304, 304, 304, 304, 304 }; // 這裡簡化對應，實際 VB6 是每 5 個換一次 X 座標
            // (註: VB6 中 X = 304, 504, 703, 903, 1103)
            int[] xBases = { 304, 504, 703, 903, 1103 };
            int[] yBases = { 108, 144, 180, 216, 252 };

            for (int i = 1; i <= 25; i++)
            {
                string key = i.ToString("D2"); // "01", "02"...
                string val = compIds.ContainsKey(key) ? compIds[key] : "";
                
                int col = (i - 1) / 5; // 決定 X (第幾欄)
                int row = (i - 1) % 5; // 決定 Y (第幾列)

                sbComps.AppendLine($"^FO{xBases[col]},{yBases[row]}^A0N,25,32^CI0^FR^FD {i,2}:   {val}^FS");
            }

            string zpl = ZplTemplates.WS_SUMMARY
                .Replace("{CompIdBlocks}", sbComps.ToString())
                .Replace("{LotNo}", lotNo)
                .Replace("{ProdNo}", prodNo)
                .Replace("{WQty}", wQty)
                .Replace("{CQty}", cQty)
                .Replace("{UserName}", userName)
                .Replace("{TimeStamp}", timeStamp);

            await SendToPrinterAsync(printerServer, zpl);
        }


        /// <summary>
        /// 3. 翻寫 Prt_FT_FT_SMALL_LABEL (複雜邏輯排版選擇)
        /// </summary>
        public async Task Prt_FT_FT_SMALL_LABEL_Async(
            string lotNo, string prodNo, string cQty, string owner, string routeId, string customer, 
            string cSum, string green, string printerServer, 
            string carGradeFlag = "", string prodLine = "", string oriLotID = "", string fuSa = "", string location = "")
        {
            string zpl = "";

            // 判斷採用哪一種 ZPL 模板
            if (location == "FT")
            {
                zpl = !string.IsNullOrEmpty(oriLotID) 
                    ? ZplTemplates.FT_SMALL_LABEL_FT_WithOriLot 
                    : ZplTemplates.FT_SMALL_LABEL_FT_NoOriLot;
            }
            else
            {
                zpl = !string.IsNullOrEmpty(oriLotID) 
                    ? ZplTemplates.FT_SMALL_LABEL_NonFT_WithOriLot 
                    : ZplTemplates.FT_SMALL_LABEL_NonFT_NoOriLot;
            }

            // 執行變數替換 (因為模板裡的變數名稱一致，我們直接一口氣 Replace 全部可能出現的變數)
            zpl = zpl.Replace("{LotNo}", lotNo)
                     .Replace("{ProdNo}", prodNo)
                     .Replace("{CQty}", cQty)
                     .Replace("{Owner}", owner)
                     .Replace("{RouteId}", routeId)
                     .Replace("{Csum}", cSum)
                     .Replace("{Green}", green)
                     .Replace("{CargradeFlag}", carGradeFlag)
                     .Replace("{ProdLine}", prodLine)
                     .Replace("{OriLotID}", oriLotID)
                     .Replace("{FuSa}", fuSa);

            await SendToPrinterAsync(printerServer, zpl);
        }
    }
}

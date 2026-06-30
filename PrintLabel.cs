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
    // 定義回傳 WS_SMALL_LABEL 所需額外資料的 DTO
    public class WsSmallLabelDbData
    {
        public string HotLotFlag { get; set; }
        public string ErunTicNo { get; set; }
        public string SapRwNo { get; set; }
        public string MfgTicNo { get; set; }
        public string WaferId { get; set; }
        public string FabLotId { get; set; }
    }

    public class InklessMergeItem
    {
        public string ChildLotId { get; set; }
        public string WaferId { get; set; }
        public string SlotNo { get; set; }
    }

    public class EngLocMasterData
    {
        public string WQty { get; set; }
        public string OwnerId { get; set; }
        public string OwnerDep { get; set; }
        public string LocId { get; set; }
    }

    public class LabelSpecInfoData
    {
        public string LabelSpecNoVer { get; set; }
        public string SpecialLabelSize { get; set; }
        public string Serial { get; set; }
        public string MergeLabelSpec { get; set; }
        public string CopiesStdLabel { get; set; }
        public string CopiesSpecialLabel { get; set; }
        public string CopiesMergeLabel { get; set; }
        public string SpecialStickReel { get; set; }
        public string SpecialStickAlbag { get; set; }
        public string SpecialPositionBox_1 { get; set; }
        public string SpecialPositionBox_2 { get; set; }
    }

    public class LabelBoxingInfoData
    {
        public string BoxingSpecNoVer { get; set; }
        public string Vacuum { get; set; }
        public string Hic { get; set; }
        public string Desiccant { get; set; }
        public string Albag { get; set; }
    }

    public class IpnWaferLevelData
    {
        public string WaferLevel { get; set; }
        public string MaskOption { get; set; }
        public string BeOption { get; set; }
        public string BomLevel { get; set; }
    }

    public class VirtualMergeLotData
    {
        public string LotId { get; set; }
        public string Pcs { get; set; }
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
        public partial class PrintLabelRepository : IPrintLabelRepository
    {
        // 1. 翻寫 GetLabelPrintQty
        public async Task<int> GetLabelPrintQtyAsync(string stage, string labelSpecNo)
        {
            string sql = @"
                SELECT PRINT_QTY 
                FROM TBL_LABEL_SPEC 
                WHERE STAGE = :p_Stage AND LABEL_SPECNO = :p_LabelSpecNo";

            int qty = await _dbConnection.QueryFirstOrDefaultAsync<int>(sql, new { p_Stage = stage, p_LabelSpecNo = labelSpecNo });
            
            return qty > 0 ? qty : 1; // 查無資料預設回傳 1
        }

        // 2. 獲取 WS_SMALL_LABEL 缺少的屬性資料
        public async Task<WsSmallLabelDbData> GetWsSmallLabelDataAsync(string lotNo)
        {
            // 將原本 VB6 裡面的 gsCAT_TBL_LOT_INFO 等常數換成實體的 Table 與 Column
            string sql = @"
                SELECT 
                    A.HOTLOTFLAG AS HotLotFlag, 
                    B.ERUNTICNO AS ErunTicNo, 
                    B.SAPRWNO AS SapRwNo, 
                    B.MFGTICNO AS MfgTicNo, 
                    B.WAFERID AS WaferId, 
                    B.FABLOTID AS FabLotId
                FROM TBL_LOT_ATTRIBUTE A
                INNER JOIN TBL_LOT_INFO B ON A.LOT_ID = B.LOT_ID
                WHERE A.LOT_ID = :p_LotNo";

            return await _dbConnection.QueryFirstOrDefaultAsync<WsSmallLabelDbData>(sql, new { p_LotNo = lotNo });
        }
        public async Task<IEnumerable<InklessMergeItem>> GetInklessMergeListAsync(string lotNo)
        {
            string sql = @"
                SELECT childlotid AS ChildLotId, waferid AS WaferId, slotno AS SlotNo 
                FROM tbl_inkless_merge_list 
                WHERE parentlotid = :p_LotNo AND deleteflag = 'N' 
                ORDER BY slotno";

            return await _dbConnection.QueryAsync<InklessMergeItem>(sql, new { p_LotNo = lotNo });
        }
        public async Task<EngLocMasterData> GetEngLocMasterDataAsync(string lotNo)
        {
            string sql = @"
                SELECT WQTY AS WQty, OWNERID AS OwnerId, OWNERDEP AS OwnerDep, LOCID AS LocId
                FROM TBL_WS_ENG_LOC_MASTER
                WHERE LOT_ID = :p_LotNo";
            
            return await _dbConnection.QueryFirstOrDefaultAsync<EngLocMasterData>(sql, new { p_LotNo = lotNo });
        }

        // 2. 翻寫 RefreshPrinterServer (抽離 UI，轉化為純資料陣列)
        public async Task<IEnumerable<string>> GetMappedPrinterServersAsync(string stage, string specNo)
        {
            // Step 1: 先找 LABEL_SPEC 是否有綁定特定的 Server (可能以逗號分隔)
            string specSql = @"
                SELECT SERVER_NAME 
                FROM TBL_LABEL_SPEC 
                WHERE STAGE = :p_Stage AND LABEL_SPECNO = :p_SpecNo";
            
            string mappedServers = await _dbConnection.QueryFirstOrDefaultAsync<string>(specSql, new { p_Stage = stage, p_SpecNo = specNo });

            // Step 2: 查詢 PRINTER_SERVER，並組合出 "SERVER_NAME@DESCRIPTION"
            string basePrinterSql = @"
                SELECT SERVER_NAME || '@' || DESCRIPTION 
                FROM TBL_PRINTER_SERVER 
                WHERE DELETE_FLAG = 'N'";

            if (!string.IsNullOrWhiteSpace(mappedServers))
            {
                // 如果有綁定，則只篩選出綁定的 Server
                var serverList = mappedServers.Split(',').Select(s => s.Trim()).ToList();
                string inClauseSql = basePrinterSql + " AND SERVER_NAME IN :p_Servers";
                
                return await _dbConnection.QueryAsync<string>(inClauseSql, new { p_Servers = serverList });
            }
            else
            {
                // 如果沒綁定，就把所有伺服器都查出來讓 User 選
                return await _dbConnection.QueryAsync<string>(basePrinterSql);
            }
        }
        // 1. 查詢 Label Spec Info
        public async Task<LabelSpecInfoData> GetLabelSpecInfoAsync(string labelSpecNo, string carrierType)
        {
            string sql = @"
                SELECT 
                    LABELSPECNOVER AS LabelSpecNoVer, SPECIALLABELSIZE AS SpecialLabelSize, SERIAL AS Serial, 
                    COPIES_STD_LABEL AS CopiesStdLabel, COPIES_SPECIAL_LABEL AS CopiesSpecialLabel, COPIES_MERGE_LABEL AS CopiesMergeLabel, 
                    MERGE_LABEL_SPEC AS MergeLabelSpec, SPECIAL_STICK_REEL AS SpecialStickReel, SPECIAL_STICK_ALBAG AS SpecialStickAlbag, 
                    SPECIAL_POSITION_BOX_1 AS SpecialPositionBox_1, SPECIAL_POSITION_BOX_2 AS SpecialPositionBox_2
                FROM TBL_FT_LABEL_SPEC_INFO 
                WHERE LABELSPECNO = :p_SpecNo AND CARRIERTYPE = :p_CarrierType AND DELETEFLAG = 'N'";
            
            return await _dbConnection.QueryFirstOrDefaultAsync<LabelSpecInfoData>(sql, new { p_SpecNo = labelSpecNo, p_CarrierType = carrierType });
        }

        // 2. 查詢 IPN_MASTER 取得 Grade 與 BizType
        public async Task<(string Grade, string BizType)> GetIpnMasterForLabelSpecAsync(string ipn)
        {
            string sql = @"
                SELECT GRADE AS Grade, BIZTYPE AS BizType 
                FROM TBL_IPN_MASTER 
                WHERE IPN = :p_IPN AND DELETE_FLAG = 'N'";
            
            return await _dbConnection.QueryFirstOrDefaultAsync<(string, string)>(sql, new { p_IPN = ipn });
        }

        // 3. 查詢 Boxing Info
        public async Task<LabelBoxingInfoData> GetLabelBoxingInfoAsync(string carrierType, string boxingSpecNo, string brand, string pinCount, string packageCode)
        {
            string sql = @"
                SELECT 
                    BOXINGSPECNOVER AS BoxingSpecNoVer, VACUUM AS Vacuum, HIC AS Hic, 
                    DESICCANT AS Desiccant, ALBAG AS Albag
                FROM TBL_FT_LABEL_BOXING_INFO 
                WHERE CARRIERTYPE = :p_CarrierType AND BOXINGSPECNO = :p_BoxingSpecNo 
                  AND BRAND = :p_Brand AND PINCOUNT = :p_PinCount 
                  AND PACKAGECODE = :p_PackageCode AND DELETEFLAG = 'N'";

            return await _dbConnection.QueryFirstOrDefaultAsync<LabelBoxingInfoData>(sql, new { 
                p_CarrierType = carrierType, p_BoxingSpecNo = boxingSpecNo, p_Brand = brand, p_PinCount = pinCount, p_PackageCode = packageCode
            });
        }
        // 1. 取得 ProdType (呼叫 Oracle Function)
        public async Task<string> GetProdTypeAsync(string ipn, string criteriaString)
        {
            string sql = "SELECT GetProdType(:p_IPN, :p_Criteria) AS PRODTYPE FROM DUAL";
            return await _dbConnection.QueryFirstOrDefaultAsync<string>(sql, new { p_IPN = ipn, p_Criteria = criteriaString });
        }

        // 2. 取得 WaferLevel 等 IPN 屬性
        public async Task<IpnWaferLevelData> GetIpnWaferLevelDataAsync(string ipn)
        {
            string sql = @"
                SELECT fun_get_waferlevel(IPN) AS WaferLevel, MASK_OPTION AS MaskOption, BE_OPTION AS BeOption, BOM_LEVEL AS BomLevel 
                FROM TBL_IPN_MASTER 
                WHERE IPN = :p_IPN AND DELETE_FLAG = 'N'";
            return await _dbConnection.QueryFirstOrDefaultAsync<IpnWaferLevelData>(sql, new { p_IPN = ipn });
        }

        // 3. 取得 Rework Print Flag
        public async Task<string> GetReworkPrintFlagAsync(string sapRwNo)
        {
            string sql = "SELECT PRINTFLAG FROM TBL_REWORK_REQ WHERE SAPRWNO = :p_SapRwNo AND DELETEFLAG = 'N'";
            return await _dbConnection.QueryFirstOrDefaultAsync<string>(sql, new { p_SapRwNo = sapRwNo });
        }

        // 4. 取得 FuSa
        public async Task<string> GetFusaAsync(string ipn)
        {
            string sql = "SELECT FUSA FROM TBL_IPN_MASTER WHERE IPN = :p_IPN";
            return await _dbConnection.QueryFirstOrDefaultAsync<string>(sql, new { p_IPN = ipn });
        }

        // 5. Check CpHot (TRACKLOT 查詢)
        public async Task<bool> CheckCpHotAsync(string lotNo)
        {
            string sql = "SELECT 1 FROM TBL_WS_TRACKLOT WHERE LOTID = :p_LotNo AND WSOUT IS NOT NULL";
            return await _dbConnection.QueryFirstOrDefaultAsync<int?>(sql, new { p_LotNo = lotNo }) != null;
        }

        // 6. Check MainProd
        public async Task<bool> CheckMainProdAsync(string prodCode, string ipn, string waferLevel)
        {
            // 完全重現 VB6 裡面的 REPLACE 萬用字元轉換邏輯
            string sql = @"
                SELECT 1 FROM TBL_PC_MAIN_PROD
                WHERE :p_ProdCode LIKE REPLACE(REPLACE(PRODCODE, '%', '_'), '*', '%')
                  AND :p_IPN LIKE REPLACE(REPLACE(IPN, '%', '_'), '*', '%')
                  AND :p_WaferLevel LIKE REPLACE(WAFERLEVEL, '*', '%')
                  AND DELETEFLAG = 'N'";
            return await _dbConnection.QueryFirstOrDefaultAsync<int?>(sql, new { p_ProdCode = prodCode, p_IPN = ipn, p_WaferLevel = waferLevel }) != null;
        }

        // 7. 取得 LabelSpec (GetLabelSpecByStage)
        public async Task<IEnumerable<string>> GetLabelSpecByStageAsync(string stage)
        {
            string sql = "SELECT DISTINCT LABEL_SPECNO FROM TBL_LABEL_SPEC WHERE STAGE = :p_Stage AND DELETE_FLAG = 'N' ORDER BY LABEL_SPECNO";
            return await _dbConnection.QueryAsync<string>(sql, new { p_Stage = stage });
        }

        // 8. 取得 Virtual Merge 清單
        public async Task<IEnumerable<VirtualMergeLotData>> GetVirtualMergeListAsync(string virtualLotId)
        {
            string sql = @"
                SELECT DISTINCT a.LOTID AS LotId, fun_splitlen(tli.WAFERID, ';') AS Pcs  
                FROM TBL_VIRTUAL_MERGE a 
                INNER JOIN TBL_LOT_INFO tli ON a.LOTID = tli.LOT_ID 
                WHERE a.DELETEFLAG = 'N' AND a.VIRTUALLOTID = :p_LotId 
                ORDER BY a.LOTID";
            return await _dbConnection.QueryAsync<VirtualMergeLotData>(sql, new { p_LotId = virtualLotId });
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

        // 💡 1. 擴充：FT_To_SubMK_Normal (REMARK 標籤)
        public static readonly string FT_To_SubMK_Normal = @"
^XA^LH00,00^LL550
^FO10,05^GB780,475,3^FS
^FO435,327^GB355,70,3^FS
^FO435,397^GB0,80,3^FS
^FO580,30^GB200,75,3^FS
^FO590,55^A0N,35,50^CI0^FDREMARK^FS
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

        // 💡 2. 擴充：WS_WS_SMALL_LABEL
        public static readonly string WS_WS_SMALL_LABEL = @"
^XA^PRC^LH0,0^FS^LL280^MD0^LH0,0^FS
^FO15,15^AFN,10,10^CI13^FDLotNo:^FS
^FO130,15^A0N,37,33^CI13^FD{LotNo}^FS
^FO400,15^AFN,10,10^CI13^FDHotFlag:^FS
^FO550,15^CF0N,37,33^CI13^FD{HotLotFlag}^FS
^FO15,50^BY2,3.0^BCN,28,N,Y,N^FR^FD>:{LotNo}^FS
^FO15,90^AFN,10,10^CI13^FDFabLotId:^FS
^FO160,90^A0N,37,33^CI13^FD{PrintFabLotId}^FS
^FO15,140^AFN,10,10^CI13^FDProdNo:^FS
^FO145,140^A0N,37,33^CI13^FD{ProdNo}^FS
^FO15,190^AFN,10,10^CI13^FDQTY:^FS
^FO90,190^A0N,37,33^CI13^FR^FD{WQty} PCS^FS
^FO300,190^BY2,3.0^BCN,28,N,Y,N^FR^FD{WQty}^FS
^FO90,230^A0N,37,33^CI13^FR^FD{CQty} EA^FS
^FO300,230^BY2,3.0^BCN,28,N,Y,N^FR^FD{CQty}^FS
^FO15,270^AFN,10,10^CI13^FDErunTicNo:^FS
^FO170,270^A0N,37,33^CI13^FD{ErunTicNo}^FS
^FO400,270^AFN,10,10^CI13^FDProdType:^FS
^FO600,270^A0N,37,33^CI13^FR^FD{BizType}^FS
^FO15,320^AFN,10,10^CI13^FDSapRwNo:^FS
^FO170,320^A0N,37,33^CI13^FD{SapRwNo}^FS
^FO400,320^AFN,10,10^CI13^FDMfgTicNo:^FS
^FO555,320^A0N,37,33^CI13^FD{MfgTicNo}^FS
^FO15,370^AFN,10,10^CI13^FDID#:^FS
^FO150,370^A0N,37,33^CI13^FD{WaferNo1}^FS
^FO150,420^A0N,37,33^CI13^FD{WaferNo2}^FS
^PQ1,1,1,Y^XZ";

        // 💡 1. 擴充：FT_To_SubFT_Eng_Sample
        // 注意：忠實保留了原本程式碼中的 "123456789012345" 寫死字串
        public static readonly string FT_To_SubFT_Eng_Sample = @"
^XA^LH00,00^LL550
^FO10,05^GB780,475,3^FS
^FO435,327^GB355,70,3^FS
^FO435,397^GB0,80,3^FS
^FO580,30^GB200,75,3^FS
^FO590,55^A0N,35,50^CI0^FDREMARK^FS
{PartialBlock}
^FO580,250^GB200,75,3^FS
^FO640,275^A0N,35,50^CI0^FDENG^FS
^FO40,58^A0N,33,28^FD(1T)LOT NO: 123456789012345^FS
^FO40,158^A0N,33,28^FD(1P)PRODUCT NO: {ProdNo}^FS
^FO40,258^A0N,33,28^FD(Q)QUANTITY: {Qty}^FS
^FO40,90^BY2,2,50^B3,,,N^FD1T{LotNo}^FS
^FO40,190^BY2,2,50^B3,,,N^FD1P{ProdNo}^FS
^FO40,290^BY2,2,50^B3,,,N^FDQ{Qty}^FS
^FO445,350^A0N,36,30^FDPACKER: {Packer}^FS
^FO445,420^A0N,36,30^FDQA:^FS
^PQ1
^XZ1";

        // 💡 2. 擴充：FT_To_SubMK_Eng_Sample
        public static readonly string FT_To_SubMK_Eng_Sample = @"
^XA^LH00,00^LL550
^FO10,05^GB780,475,3^FS
^FO435,327^GB355,70,3^FS
^FO435,397^GB0,80,3^FS
^FO580,30^GB200,75,3^FS
^FO590,55^A0N,35,50^CI0^FDREMARK^FS
{PartialBlock}
^FO580,250^GB200,75,3^FS
^FO640,275^A0N,35,50^CI0^FDENG^FS
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

        // 💡 3. 擴充：FT_To_SubTR_Eng_Sample
        public static readonly string FT_To_SubTR_Eng_Sample = @"
^XA^LH00,00^LL550
^FO10,05^GB780,475,3^FS
^FO435,327^GB355,70,3^FS
^FO435,397^GB0,80,3^FS
^FO580,30^GB200,75,3^FS
^FO640,55^A0N,35,50^CI0^FDT&R^FS
{PartialBlock}
^FO580,250^GB200,75,3^FS
^FO640,275^A0N,35,50^CI0^FDENG^FS
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

        // 💡 4. 擴充：WSMCD_TO_SFG
        public static readonly string WSMCD_TO_SFG = @"
^XA
^LH0,0^FS
^LL1441
^MD0
^MNY
^FO30,50^A0N,30,30^CI0^FR^FDLOT : {ParLotNo}^FS
^BY2, 2.0^FO25,85^B3N,N,60,N,Y^FR^FD{ParLotNo}^FS
^FO30,160^A0N,30,32^CI0^FR^FDPROD : {ParIPN}^FS
^BY1, 3.0^FO25,195^B3N,N,60,N,Y^FR^FD{ParIPN}^FS
^FO360,300^A0R,40,32^CI0^FR^FDLOT^FS
^FO360,500^A0R,40,32^CI0^FR^FDPROD^FS
^FO360,900^A0R,40,32^CI0^FR^FDQTY^FS
^FO360,1100^A0R,40,32^CI0^FR^FDCODE^FS
{DynamicListBlock}
^FO105,300^A0R,30,32^CI0^FR^FDTTL QTY: {ParWqty} PCS,   {ParCQTY}EA^FS
^FO105,720^A0R,30,32^CI0^FR^FDID No: {UserName}^FS
^FO105,950^A0R,30,32^CI0^FR^FDDG: 0^FS
^FO50,300^A0R,30,32^CI0^FR^FDDATE: {TimeStamp}^FS
^FO50,650^A0R,30,32^CI0^FR^FDSerial No: 0206000115^FS
^FO50,1000^A0R,30,32^CI0^FR^FDTQA:^FS
^XZ";
// 💡 1. FT_ETEST_MERGE 標籤 (底圖)
        public static readonly string FT_ETEST_MERGE_Base = @"
^XA^PRC^LH0,0^FS^LL280^MD0^LH0,0^FS
{DynamicMergeBlock}
^PQ1,1,1,Y^XZ";

        // 💡 2. WS_WS_DGRADE_SUMMARY 標籤
        public static readonly string WS_WS_DGRADE_SUMMARY = @"
^XA
^LH0,0^FS
^LL440
^MD0
^MNY
^FO368,36^A0N,40.32^CI0^FR^FDLOT : {LotNo}^FS
^FO48,10^A0B,30,30^CI0^FR^FDLOT : {LotNo}^FS
^BY2, 2.0^FO80,50^B3B,N,60,N,Y^FR^FD{LotNo}^FS
^FO775,36^A0N,40,32^CI0^FR^FDPROD : {ProdNo}^FS
^FO160,10^A0B,30,32^CI0^FR^FDPROD : {ProdNo}^FS
^BY1, 3.0^FO192,50^B3B,N,60,N,Y^FR^FD{ProdNo}^FS
{DynamicWaferBlock}
^FO250,318^A0N,30,32^CI0^FR^FDTTL WQTY: {WQty} PCS ^FS
^BY2,2.0^FO540,318^B3N,N,40,N,Y^FR^FD{WQty}^FS
^FO250,370^A0N,30,32^CI0^FR^FDTTL CQTY: {CQty}EA ^FS
^BY2,2.0^FO540,370^B3N,N,40,N,Y^FR^FD{CQty}^FS
^FO690,318^A0N,30,32^CI0^FR^FDID No: {UserName}^FS
^FO980,318^A0N,30,32^CI0^FR^FDTQA:^FS
^FO790,370^A0N,30,32^CI0^FR^FDDATE: {TimeStamp}^FS
^FO1070,284^GB190,140,4^FS
^XZ";

        // 💡 3. WSMULTILOT_TO_SFG 標籤
        public static readonly string WSMULTILOT_TO_SFG = @"
^XA
^LH0,0^FS
^LL440
^MD0
^MNY
^FO116,36^A0N,40,32^CI0^FR^FDLOT^FS
^FO300,36^A0N,40,32^CI0^FR^FDPROD^FS
^FO655,36^A0N,40,32^CI0^FR^FDQTY^FS
^FO783,36^A0N,40,32^CI0^FR^FDCODE^FS
{DynamicMultiLotBlock}
^FO96,304^A0N,30,32^CI0^FR^FDTTL QTY:  {WQty} PCS,    {CQty} EA^FS
^FO520,304^A0N,30,32^CI0^FR^FDID No:  {UserName} ^FS
^FO747,304^A0N,30,32^CI0^FR^FDDG: 0^FS
^FO96,356^A0N,30,32^CI0^FR^FDDATE: {TimeStamp} ^FS
^FO480,356^A0N,30,32^CI0^FR^FDSerial No: 0206000115^FS
^FO850,356^A0N,30,32^CI0^FR^FDTQA:^FS
^FO930,284^GB190,140,4^FS
^XZ";

        // 💡 4. WS_RENESAS_SHIPPING 標籤 (Renesas 客戶專用出貨標籤)
        public static readonly string WS_RENESAS_SHIPPING = @"
^XA^PRC^LH0,0^FS^LL1441^MD0^MNY^LH0,0^FS
^FO38,370^A0B,30,0^CI0^FR^FDLOT:^FS
^FO38,200^A0B,30,0^CI0^FR^FD{LotNo}^FS
^BY2,2.0^FO70,38^BAB,40,N,,^FR^FD{LotNo}^FS
^FO119,350^A0B,30,0^CI0^FR^FDPART:^FS
^FO119,24^A0B,30,0^CI0^FR^FD{PartNo}^FS
^BY2,2,40^FO154,18^BAB,,N,,^FR^FD{PartNo}^FS
^FO202,350^A0B,30,0^CI0^FR^FDPROD:^FS
^FO202,80^A0B,30,0^CI0^FR^FD{ProdNo}^FS
{DynamicWaferBlock}
^FO943,200^A0N,35,0^CI0^FR^FDOn line^FS
^FO252,384^A0N,35,0^CI0^FR^FDTTL QTY:^FS
^FO390,384^A0N,35,0^CI0^FR^FD{WQty}PCS,^FS
^FO506,384^A0N,35,0^CI0^FR^FD{Qty}EA,^FS
^FO943,244^A0N,35,0^CI0^FR^FDID No:^FS
^FO1039,244^A0N,35,0^CI0^FR^FD{UserID}^FS
^FO694,384^A0N,35,0^CI0^FR^FDDATE:^FS
^FO794,384^A0N,35,0^CI0^FR^FD{PrintDate}^FS
^FO943,288^A0N,35,0^CI0^FR^FDTQA:^FS
^FO1030,284^GB190,140,4^FS
^PQ1,0,0,N^XZ";

        // 💡 1-1. BIN_CARD_LABEL (Pass 變體)
        public static readonly string FT_BIN_CARD_PASS = @"
^XA^PRC ^LH0,0^FS^LL751^MD0^MNM^LH144,0^FS
^FO109,50^A0N,48,71^CI13^FR^FDMacronix ^FS
^FO16,258^A0N,27,28^CI13^FR^FDLOT_NO   :{LotNo}^FS
^BY2,3.0^FO150,300^BCN,50,N,Y,N^FR^FD>:{LotNo}^FS
^FO16,386^A0N,27,28^CI13^FR^FDTest Mode :{StepID} {TestMode}^FS
^FO16,470^A0N,31,29^CI13^FR^FDTA-ID      :^FS
^FO296,467^A0N,29,25^CI13^FR^FDDate   :^FS
^FO16,533^A0N,26,37^CI13^FR^FDBIN ^FS
^FO16,569^A0N,26,37^CI13^FR^FDMODE   :^FS
^FO16,631^A0N,27,40^CI13^FR^FDQ'TY    :^FS
^FO16,709^A0N,27,40^CI13^FR^FDNOTE   :^FS
^FO88,109^A0N,48,60^CI13^FR^FDPASS BIN CARD^FS
^FO16,184^A0N,27,32^CI13^FR^FDProd_ID  :{ProdNo}^FS
^FO5,29^GB525,719,4^FS
^FO11,234^GB512,0,4^FS
^FO11,357^GB512,0,4^FS
^FO11,437^GB512,0,4^FS
^FO11,520^GB512,0,4^FS
^FO13,599^GB512,0,4^FS
^FO8,677^GB512,0,4^FS
^PQ1,0,0,N^XZ
^FX End of job
^XA^IDR:ID*.*^XZ";

        // 💡 1-2. BIN_CARD_LABEL (Fail 變體，多了填寫格)
        public static readonly string FT_BIN_CARD_FAIL = @"
^XA^PRC ^LH0,0^FS^LL751^MD0^MNM^LH128,0^FS
^FO123,42^A0N,48,71^CI13^FR^FDMacronix ^FS
^FO21,517^A0N,27,40^CI13^FR^FDBIN2:______^FS
^FO21,562^A0N,27,40^CI13^FR^FDBIN4:______^FS
^FO19,258^A0N,27,28^CI13^FR^FDLOT_NO    :{LotNo}^FS
^BY2,3.0^FO150,300^BCN,50,N,Y,N^FR^FD>:{LotNo}^FS
^FO19,368^A0N,27,28^CI13^FR^FDTest Mode :{StepID} {TestMode}^FS
^FO19,445^A0N,27,31^CI13^FR^FDTA-ID    :^FS
^FO264,446^A0N,31,28^CI13^FR^FDDate   :^FS
^FO21,607^A0N,27,40^CI13^FR^FDBIN6:______^FS
^FO21,647^A0N,27,40^CI13^FR^FDBIN_:______^FS
^FO21,693^A0N,27,40^CI13^FR^FDNOTE   :^FS
^FO88,109^A0N,48,60^CI13^FR^FDFAIL BIN CARD^FS
^FO19,181^A0N,27,32^CI13^FR^FDProd_ID  :{ProdNo}^FS
^FO282,647^A0N,27,40^CI13^FR^FDREJ :______^FS
^FO282,607^A0N,27,40^CI13^FR^FDBIN7:______^FS
^FO282,562^A0N,27,40^CI13^FR^FDBIN5:______^FS
^FO282,517^A0N,27,40^CI13^FR^FDBIN3:______^FS
^FO11,11^GB525,719,4^FS
^FO13,237^GB520,0,4^FS
^FO13,352^GB520,0,4^FS
^FO13,415^GB520,0,4^FS
^FO11,493^GB520,0,4^FS
^PQ1,0,0,N^XZ
^FX End of job
^XA^IDR:ID*.*^XZ";

        // 💡 2. ENG_LOC_LABEL (工程倉儲標籤)
        public static readonly string WS_ENG_LOC_LABEL = @"
^XA^PRC ^LH0,0^FS^LL480^MD0^MNY^LH0,0^FS
^CWI,V10SPR_A.FNT^FS
^FO29,37^AIN,45,0^CI0^FR^FDLOTID :{LotNo}^FS
^CWJ,U10SPR_A.FNT^FS
^FO29,176^AJN,45,0^CI0^FR^FDWQTY :{WQty}^FS
^CWK,U1WSPR_A.FNT^FS
^FO27,253^AKN,45,0^CI0^FR^FDOwnerID :{OwnerId}^FS
^CWJ,U10SPR_A.FNT^FS
^FO29,330^AJN,45,0^CI0^FR^FDOwnerDep :{OwnerDep}^FS
^BY2,3.0^FO36,107^B3N,N,40,N,Y^FR^FD{LotNo}^FS
^FO13,19^GB757,434,4^FS
^PQ1,0,0,N^XZ
^FX End of job
^XA^IDR:ID*.*^XZ";

        // 💡 3. BOX_COUNTING_LABEL (外箱計數標籤)
        public static readonly string FT_BOX_COUNTING_LABEL = @"
^XA^PRC ^LH0,0^FS^LL400^MD0^MNM^LH128,0^FS
^FO168,30^A0N,43,37^CI13^FR^FDBox Counting^FS
^FO24,129^A0N,31,28^CI13^FR^FDProd_ID:{IPN}^FS
^FO27,217^A0N,31,28^CI13^FR^FDLOT_NO:{LotId}^FS
^BY2,3.0^FO30,290^BCN,50,N,Y,N^FR^FD>:{LotId}^FS
^FO290,196^A0N,69,61^CI13^FR^FDBOX:{PrintCount}^FS
^FO11,8^GB528,360,4^FS
^FO13,83^GB512,0,4^FS
^PQ1,0,0,N^XZ
^FX End of job
^XA^IDR:ID*.*^XZ";

// 💡 擴充：FT_Label_PACK_INFO 標籤 (非常龐大的多欄位表格標籤)
        public static readonly string FT_Label_PACK_INFO = @"
^XA^MCY^XZ
^XA^FWN^CFD,24^LH0,0^CI0^PR2^MNY^MTT^MMT^MD0^PON^PMN^LRN^XZ
^XA^DFR:TEMP_FMT.ZPL^LRN^XZ
^XA^XFR:TEMP_FMT.ZPL
^FO15,17^GB1167,112,2,B,0^FS
^FO15,71^GB1165,2,2,B,0^FS
^FO191,18^GB2,110,2,B,0^FS
^FO247,16^GB2,110,2,B,0^FS
^FO799,18^GB2,110,2,B,0^FS
^FO709,18^GB2,110,2,B,0^FS
^FO988,15^GB2,110,2,B,0^FS
^A0N,23,24^FO40,33^FDLabel Spec.^FS
^A0N,23,24^FO425,33^FDLabel Size^FS
^A0N,23,24^FO199,33^FDVer.^FS
^A0N,23,24^FO723,33^FDSerial^FS
^A0N,23,24^FO858,33^FDSource^FS
^A0N,23,24^FO1047,33^FDLotNo.^FS
^A0N,28,28^FO255,89^FD{SpecialLabelSize}^FS
^A0N,28,28^FO199,89^FD{LabelSpecNoVer}^FS
^A0N,28,28^FO20,89^FD{LabelSpecNo}^FS
^A0N,28,28^FO998,89^FD{PrintLotNO}^FS
^A0N,28,28^FO859,89^FD{PrintMode}^FS
^A0N,28,28^FO741,89^FD{Serial}^FS
^FO13,315^GB1167,112,2,B,0^FS
^FO15,368^GB1165,2,2,B,0^FS
^FO519,315^GB2,110,2,B,0^FS
^FO799,314^GB2,110,2,B,0^FS
^FO398,314^GB2,110,2,B,0^FS
^FO247,314^GB2,110,2,B,0^FS
^FO191,316^GB2,110,2,B,0^FS
^FO988,312^GB2,110,2,B,0^FS
^A0N,23,24^FO36,333^FDBoxing Spec.^FS
^A0N,23,24^FO198,333^FDVer.^FS
^A0N,23,24^FO294,333^FDBrand^FS
^A0N,23,24^FO419,333^FDPackage^FS
^A0N,23,24^FO594,333^FDCarrier Type^FS
^A0N,23,24^FO808,333^FDVacuum^FS
^A0N,23,24^FO921,333^FDHIC^FS
^A0N,23,24^FO992,333^FDDesiccant^FS
^A0N,23,24^FO1093,333^FDAl. bag^FS
^FO895,314^GB2,110,2,B,0^FS
^FO1087,314^GB2,110,2,B,0^FS
^A0N,28,28^FO19,387^FD{BoxingSpecNo}^FS
^A0N,28,28^FO199,388^FD{BoxingSpecNoVer}^FS
^A0N,28,28^FO296,388^FD{Brand}^FS
^A0N,28,28^FO411,388^FD{PackageCodePinCount}^FS
^A0N,28,28^FO527,388^FD{CarrierType}^FS
^A0N,28,28^FO832,388^FD{Vacuum}^FS
^A0N,28,28^FO934,388^FD{Hic}^FS
^A0N,28,28^FO1032,388^FD{Desiccant}^FS
^A0N,28,28^FO1108,387^FD{Albag}^FS
^FO13,136^GB1167,168,2,B,0^FS
^FO13,190^GB1165,2,2,B,0^FS
^FO14,248^GB1165,2,2,B,0^FS
^FO190,190^GB2,110,2,B,0^FS
^FO400,190^GB2,110,2,B,0^FS
^FO520,140^GB2,162,2,B,0^FS
^FO802,140^GB2,162,2,B,0^FS
^FO988,137^GB2,162,2,B,0^FS
^A0N,23,24^FO214,153^FDLabel  Copies^FS
^A0N,23,24^FO69,209^FDSTD^FS
^A0N,23,24^FO251,209^FDSpecial^FS
^A0N,23,24^FO422,209^FDMerge^FS
^A0N,23,24^FO578,153^FDMerge Label^FS
^A0N,23,24^FO621,209^FDSpec.^FS
^A0N,23,24^FO825,153^FDSpecial Stick^FS
^A0N,23,24^FO1018,153^FDSpecial Pos.^FS
^A0N,23,24^FO822,209^FDReel^FS
^FO895,192^GB2,110,2,B,0^FS
^FO1088,193^GB2,110,2,B,0^FS
^A0N,23,24^FO903,209^FDAl. Bag^FS
^A0N,23,24^FO1007,209^FDBox_1^FS
^A0N,23,24^FO1098,208^FDBox_2^FS
^A0N,28,28^FO78,265^FD{CopiesStdLabel}^FS
^A0N,28,28^FO270,265^FD{CopiesSpecialLabel}^FS
^A0N,28,28^FO442,265^FD{CopiesMergeLabel}^FS
^A0N,28,28^FO535,265^FD{MergeLabelSpec}^FS
^A0N,28,28^FO837,265^FD{SpecialStickReel}^FS
^A0N,28,28^FO926,265^FD{SpecialStickAlbag}^FS
^A0N,28,28^FO1000,265^FD{SpecialPositionBox_1}^FS
^A0N,28,28^FO1094,265^FD{SpecialPositionBox_2}^FS
^PQ1,0,1,Y^XZ
^XA^IDR:TEMP_FMT.ZPL^XZ";

// 💡 1. CP_SMALL_LABEL (一般版 / 虛擬併批單張)
        public static readonly string WS_CP_SMALL_LABEL = @"
^XA^PRC^LH0,0^FS^LL280^MD0^LH0,0^FS
^FO15,15^AFN,10,10^CI13^FDLotNo:^FS
^FO130,15^A0N,37,33^CI13^FD{LotNo}^FS
^FO15,50^BY2,3.0^BCN,28,N,Y,N^FR^FD>:{LotNo}^FS
^FO15,105^AFN,10,10^CI13^FDFabLotId:^FS
^FO170,105^A0N,37,33^CI13^FD{PrintFabLotId}^FS
^FO15,150^AFN,10,10^CI13^FDProdBody:^FS
^FO170,150^A0N,37,33^CI13^FD{ProdBody}^FS
^FO290,150^AFN,10,10^CI13^FDHot:^FS
^FO360,150^CF0N,37,33^CI13^FD{HotLotFlag}^FS
^FO290,195^AFN,10,10^CI13^FDQTY:^FS
^FO360,195^A0N,37,33^CI13^FR^FD{WQty} PCS^FS
^FO15,195^AFN,10,10^CI13^FDProdType:^FS
^FO170,195^A0N,37,33^CI13^FR^FD{BizType}^FS
^FO15,240^AFN,10,10^CI13^FDErunTicNo:^FS
^FO170,240^A0N,37,33^CI13^FD{ErunTicNo}^FS
^FO355,240^AFN,10,10^CI13^FDFuSa:^FS
^FO435,240^A0N,37,33^CI13^FD{FuSa}^FS
^FO15,285^AFN,10,10^CI13^FDSapRwNo:^FS
^FO170,285^A0N,37,33^CI13^FD{SapRwNo}^FS
^FO15,330^AFN,10,10^CI13^FDID#:^FS
^FO80,330^A0N,37,33^CI13^FD{WaferNo1}^FS
^FO80,380^A0N,37,33^CI13^FD{WaferNo2}^FS
^FO80,430^A0N,37,33^CI13^FD{WaferNo3}^FS
^PQ1,1,1,Y^XZ";

        // 💡 2. CP_VIRTUAL_MERGE_LIST (虛擬併批清單版，印出長條表單)
        public static readonly string WS_CP_VIRTUAL_MERGE_LIST = @"
^XA^LH0,0^FS^LL1980^MD0^MNY
^FO30,25^A0N,48.36^CI0^FR^FDVitualLotId : {LotNo}^FS
^BY2,2.0^FO540,30^B3N,N,60,N,N^FD{LotNo}^FS
^FO30,102^A0N,48.36^CI0^FR^FDProdBody : {ProdBody}^FS
^FO405,102^A0N,48.36^CI0^FR^FDProdLevel : {WaferLevel}^FS
^FO30,174^A0N,48.36^CI0^FR^FDWaferQty : {WQty}^FS
^BY2,2.0^FO405,174^B3N,N,60,N,N^FD{WQty}^FS
^FO30,246^A0N,48.36^CI0^FR^FDChipQty : {CQty}^FS
^BY2,2.0^FO405,246^B3N,N,60,N,N^FD{CQty}^FS
^FO30,318^A0N,48.36^CI0^FR^FDID# :^FS
{DynamicLotListBlock}
^XZ";
        
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
        private readonly IPrintLabelRepository _repo;

        public PrintLabelService(IMessageQueueService mqService)
        {
            _mqService = mqService;
            _repo = repo;
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
        /// <summary>
        /// 1. 翻寫 Prt_FT_To_SubMK_Normal (REMARK 標籤)
        /// </summary>
        public async Task Prt_FT_To_SubMK_Normal_Async(string lotNo, string prodNo, string qty, string packer, bool isPartial, string printerServer)
        {
            string zpl = ZplTemplates.FT_To_SubMK_Normal;

            string partialZpl = isPartial ? ZplTemplates.PartialBlock : "";
            zpl = zpl.Replace("{PartialBlock}", partialZpl)
                     .Replace("{LotNo}", lotNo)
                     .Replace("{ProdNo}", prodNo)
                     .Replace("{Qty}", qty)
                     .Replace("{Packer}", packer);

            await SendToPrinterAsync(printerServer, zpl);
        }

        /// <summary>
        /// 2. 翻寫 GetLabelPrintQty (直接呼叫 Repo)
        /// </summary>
        public async Task<int> GetLabelPrintQtyAsync(string stage, string labelSpecNo)
        {
            return await _repo.GetLabelPrintQtyAsync(stage, labelSpecNo);
        }

        /// <summary>
        /// 3. 翻寫 Prt_WS_WS_SMALL_LABEL (包含資料庫查詢與複雜字串切割)
        /// </summary>
        public async Task Prt_WS_WS_SMALL_LABEL_Async(
            string lotNo, string prodNo, string wQty, string cQty, string owner, string bizType, 
            string printerServer, string inputWaferId = "", string inputFabLotId = "")
        {
            // 1. 去資料庫撈取其他資訊
            var dbData = await _repo.GetWsSmallLabelDataAsync(lotNo);
            if (dbData == null) dbData = new WsSmallLabelDbData(); // 防呆

            string printFabLotId = !string.IsNullOrEmpty(inputFabLotId) ? inputFabLotId : dbData.FabLotId;
            string rawWaferStr = !string.IsNullOrEmpty(inputWaferId) ? inputWaferId : dbData.WaferId;

            // 2. 處理 Wafer 字串 (取代 VB6 落落長的迴圈與 Split)
            // 將傳入的 "1;3;5" 轉為 int 陣列 [1, 3, 5]
            var parsedWafers = (rawWaferStr ?? "")
                .Split(new[] { ';', ',' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(w => int.TryParse(w, out int val) ? val : 0)
                .ToList();

            var waferList = new List<string>();
            
            // 產生 25 片的格式：若存在填入 "01", 不存在填入 "__"
            for (int i = 1; i <= 25; i++)
            {
                waferList.Add(parsedWafers.Contains(i) ? i.ToString("D2") : "__");
            }

            // 用逗號串接，結果會長得像： "01,__,03,__,05..."
            string fullWaferString = string.Join(",", waferList);

            // 依照 VB6 邏輯：前 38 字元為 WaferNo1，第 40 字元以後為 WaferNo2 (剛好跳過第 39 個的逗號)
            string waferNo1 = fullWaferString.Length > 38 ? fullWaferString.Substring(0, 38) : fullWaferString;
            string waferNo2 = fullWaferString.Length > 39 ? fullWaferString.Substring(39) : "";

            // 3. 模板變數替換
            string zpl = ZplTemplates.WS_WS_SMALL_LABEL
                .Replace("{LotNo}", lotNo)
                .Replace("{HotLotFlag}", dbData.HotLotFlag ?? "")
                .Replace("{PrintFabLotId}", printFabLotId ?? "")
                .Replace("{ProdNo}", prodNo)
                .Replace("{WQty}", wQty)
                .Replace("{CQty}", cQty)
                .Replace("{ErunTicNo}", dbData.ErunTicNo ?? "")
                .Replace("{BizType}", bizType)
                .Replace("{SapRwNo}", dbData.SapRwNo ?? "")
                .Replace("{MfgTicNo}", dbData.MfgTicNo ?? "")
                .Replace("{WaferNo1}", waferNo1)
                .Replace("{WaferNo2}", waferNo2);

            // 4. 發送至印表機
            await SendToPrinterAsync(printerServer, zpl);
        }
        /// <summary>
        /// 共用邏輯：處理三個 Eng_Sample 標籤的發送
        /// </summary>
        private async Task SendEngSampleAsync(string template, string lotNo, string prodNo, string qty, string packer, bool isPartial, string printerServer)
        {
            string partialZpl = isPartial ? ZplTemplates.PartialBlock : "";
            
            string zpl = template
                     .Replace("{PartialBlock}", partialZpl)
                     .Replace("{LotNo}", lotNo)
                     .Replace("{ProdNo}", prodNo)
                     .Replace("{Qty}", qty)
                     .Replace("{Packer}", packer);

            await SendToPrinterAsync(printerServer, zpl);
        }

        // 1. 翻寫 Prt_FT_To_SubFT_Eng_Sample
        public async Task Prt_FT_To_SubFT_Eng_Sample_Async(string lotNo, string prodNo, string qty, string packer, bool isPartial, string printerServer)
        {
            await SendEngSampleAsync(ZplTemplates.FT_To_SubFT_Eng_Sample, lotNo, prodNo, qty, packer, isPartial, printerServer);
        }

        // 2. 翻寫 Prt_FT_To_SubMK_Eng_Sample
        public async Task Prt_FT_To_SubMK_Eng_Sample_Async(string lotNo, string prodNo, string qty, string packer, bool isPartial, string printerServer)
        {
            await SendEngSampleAsync(ZplTemplates.FT_To_SubMK_Eng_Sample, lotNo, prodNo, qty, packer, isPartial, printerServer);
        }

        // 3. 翻寫 Prt_FT_To_SubTR_Eng_Sample
        public async Task Prt_FT_To_SubTR_Eng_Sample_Async(string lotNo, string prodNo, string qty, string packer, bool isPartial, string printerServer)
        {
            await SendEngSampleAsync(ZplTemplates.FT_To_SubTR_Eng_Sample, lotNo, prodNo, qty, packer, isPartial, printerServer);
        }

        /// <summary>
        /// 4. 翻寫 Prt_WSMCD_TO_SFG
        /// 註: VB6 的 oWSMCDInfo (Collection) 這裡對應為 List<string> (格式 "Lot;IPN;Code;Qty")
        /// </summary>
        public async Task Prt_WSMCD_TO_SFG_Async(
            string parLotNo, string parIpn, string parWQty, string parCQty, 
            string userName, string timeStamp, string printerServer, List<string> wsmcdInfo = null)
        {
            // 初始化五組空資料
            var lots = new string[5] { "", "", "", "", "" };
            var ipns = new string[5] { "", "", "", "", "" };
            var codes = new string[5] { "", "", "", "", "" };
            var cQtys = new string[5] { "", "", "", "", "" };

            // 條件1：沒有傳入集合
            if (wsmcdInfo == null || wsmcdInfo.Count == 0)
            {
                lots[0] = parLotNo;
                ipns[0] = parIpn;
                cQtys[0] = parCQty;
            }
            else
            {
                // 條件2：有傳入集合 (VB6 邏輯是「從後面往前取」，且最多取 5 筆)
                // VB6語法：oWSMCDInfo.Item(oWSMCDInfo.Count - iIdx + 1)
                for (int i = 0; i < wsmcdInfo.Count && i < 5; i++)
                {
                    // C# 反向索引
                    string item = wsmcdInfo[wsmcdInfo.Count - 1 - i]; 
                    var parts = item.Split(';');
                    
                    if (parts.Length > 0) lots[i] = parts[0];
                    if (parts.Length > 1) ipns[i] = parts[1];
                    if (parts.Length > 2) codes[i] = parts[2];
                    if (parts.Length > 3) cQtys[i] = parts[3];
                }
            }

            // 動態組裝 5 列資料的 ZPL，初始 X 座標為 304，每一次遞減 36
            var sbDynamicItems = new StringBuilder();
            int currentX = 304;

            for (int i = 0; i < 5; i++)
            {
                // 若陣列中有值，或者沒值(印出空字串，ZPL渲染時自然為空)
                sbDynamicItems.AppendLine($"^FO{currentX},300^A0R,25,32^CI0^FR^FD{lots[i]}^FS");
                sbDynamicItems.AppendLine($"^FO{currentX},500^A0R,25,32^CI0^FR^FD{ipns[i]}^FS");
                sbDynamicItems.AppendLine($"^FO{currentX},900^A0R,25,32^CI0^FR^FD{cQtys[i]}^FS");
                sbDynamicItems.AppendLine($"^FO{currentX},1100^A0R,25,32^CI0^FR^FD{codes[i]}^FS");

                currentX -= 36; // 產生 304, 268, 232, 196, 160 序列
            }

            // 替換主模板變數
            string zpl = ZplTemplates.WSMCD_TO_SFG
                .Replace("{DynamicListBlock}", sbDynamicItems.ToString())
                .Replace("{ParLotNo}", parLotNo)
                .Replace("{ParIPN}", parIpn)
                .Replace("{ParWqty}", parWQty)
                .Replace("{ParCQTY}", parCQty)
                .Replace("{UserName}", userName)
                .Replace("{TimeStamp}", timeStamp);

            await SendToPrinterAsync(printerServer, zpl);
        }
        /// <summary>
        /// 1. 翻寫 Prt_FT_ETEST_MERGE (每 6 個 Lot/Qty 組合印一張標籤)
        /// </summary>
        public async Task Prt_FT_ETEST_MERGE_Async(string sLotsInfo, string printerServer)
        {
            var arrLotInfo = sLotsInfo.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
            
            // VB6 Y 座標：15, 60, 105, 150, 195, 235
            int[] yCoords = { 15, 60, 105, 150, 195, 235 };
            
            // 每次處理 12 個元素 (6 對 Lot 和 Qty)
            for (int i = 0; i < arrLotInfo.Length; i += 12)
            {
                var sbBlock = new StringBuilder();
                int currentYIndex = 0;

                // 處理這一批次內的 12 個項目
                for (int j = 0; j < 12 && (i + j) < arrLotInfo.Length; j++)
                {
                    string item = arrLotInfo[i + j];
                    
                    if (j % 2 == 0) // 偶數是 LotNo (X = 15)
                    {
                        sbBlock.AppendLine($"^FO15,{yCoords[currentYIndex]}^A0N,40,30^CI13^FR^FDLot No: {item}^FS");
                    }
                    else // 奇數是 QTY (X = 300)
                    {
                        sbBlock.AppendLine($"^FO300,{yCoords[currentYIndex]}^A0N,40,30^CI13^FR^FDQTY: {item}^FS");
                        currentYIndex++; // 寫完一對，Y座標往下移動一列
                    }
                }

                // 套入模板並發送
                string zpl = ZplTemplates.FT_ETEST_MERGE_Base.Replace("{DynamicMergeBlock}", sbBlock.ToString());
                await SendToPrinterAsync(printerServer, zpl);
            }
        }


        /// <summary>
        /// 2. 翻寫 Prt_WS_WS_DGRADE_SUMMARY 
        /// (⚠️ 注意：這裡展示了如何解決 VB6 InputBox 的問題)
        /// </summary>
        public async Task Prt_WS_WS_DGRADE_SUMMARY_Async(string lotNo, string prodNo, string wQty, string cQty, string userName, string timeStamp, string printerServer)
        {
            var mergeList = (await _repo.GetInklessMergeListAsync(lotNo)).ToList();

            // 若沒有撈到資料，且 LotNo 長度為 10 且後兩碼是數字
            if (!mergeList.Any())
            {
                if (lotNo.Length == 10 && char.IsDigit(lotNo[8]) && char.IsDigit(lotNo[9]))
                {
                    // 🚨【解決 InputBox 問題】：
                    // 在 Web API 中，我們不能暫停程式等待輸入。因此拋出一個特殊的例外或回傳代碼。
                    // 前端收到這個特定的 Error Message 後，可以彈出 MessageBox.prompt 請 User 輸入原始批號，
                    // 然後前端再重新呼叫一次這支 API，把 User 輸入的值當作 lotNo 傳進來。
                    throw new ArgumentException("REQUIRE_ORIGINAL_LOT: 無母批merge資訊, 請輸入原始批號");
                }
                // 若連防呆都不符合，直接結束 (與 VB6 的 GoTo ExitHandler 行為一致)
                return; 
            }

            // 初始化 25 個槽位
            var items = Enumerable.Repeat("______-__", 25).ToArray();

            bool hasSlotNo = !string.IsNullOrEmpty(mergeList.First().SlotNo);

            if (!hasSlotNo)
            {
                // 邏輯一：沒有 SlotNo，把分號隔開的 waferId 展開依序填入
                int index = 0;
                foreach (var row in mergeList)
                {
                    var wafers = row.WaferId.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
                    foreach (var w in wafers)
                    {
                        if (index >= 25) break;
                        string prefix = row.ChildLotId.Length >= 6 ? row.ChildLotId.Substring(0, 6) : row.ChildLotId;
                        items[index] = $"{prefix}-{w}";
                        index++;
                    }
                }
            }
            else
            {
                // 邏輯二：有 SlotNo，直接依照位置放置
                foreach (var row in mergeList)
                {
                    if (int.TryParse(row.SlotNo, out int sIndex) && sIndex >= 1 && sIndex <= 25)
                    {
                        string prefix = row.ChildLotId.Length >= 6 ? row.ChildLotId.Substring(0, 6) : row.ChildLotId;
                        items[sIndex - 1] = $"{prefix}-{row.WaferId}";
                    }
                }
            }

            // 產生 25 個座標方塊
            int[] xCoords = { 300, 480, 680, 880, 1080 };
            int[] yCoords = { 108, 144, 180, 216, 252 };
            var sbWafers = new StringBuilder();

            for (int i = 0; i < 25; i++)
            {
                int col = i / 5; // X
                int row = i % 5; // Y
                sbWafers.AppendLine($"^FO{xCoords[col]},{yCoords[row]}^A0N,25,32^CI0^FR^FD {items[i]}^FS");
            }

            string zpl = ZplTemplates.WS_WS_DGRADE_SUMMARY
                .Replace("{DynamicWaferBlock}", sbWafers.ToString())
                .Replace("{LotNo}", lotNo)
                .Replace("{ProdNo}", prodNo)
                .Replace("{WQty}", wQty)
                .Replace("{CQty}", cQty)
                .Replace("{UserName}", userName)
                .Replace("{TimeStamp}", timeStamp);

            await SendToPrinterAsync(printerServer, zpl);
        }


        /// <summary>
        /// 3. 翻寫 Prt_WSMULTILOT_TO_SFG
        /// </summary>
        public async Task Prt_WSMULTILOT_TO_SFG_Async(string userName, string timeStamp, string wQty, string cQty, string printerServer, List<string> colLotInfo)
        {
            var sbLots = new StringBuilder();

            // 依序處理每一筆
            for (int i = 0; i < colLotInfo.Count; i++)
            {
                var parts = colLotInfo[i].Split(';'); // childLotId;IPN;Code;CQty
                string childLot = parts.Length > 0 ? parts[0] : "";
                string ipn = parts.Length > 1 ? parts[1] : "";
                string code = parts.Length > 2 ? parts[2] : "";
                string qty = parts.Length > 3 ? parts[3] : "";

                int currentY = 108 + 36 * i;

                sbLots.AppendLine($"^FO56, {currentY}^A0N,25,32^CI0^FR^FD{childLot}^FS");
                sbLots.AppendLine($"^FO296, {currentY}^A0N,25,32^CI0^FR^FD{ipn}^FS");
                sbLots.AppendLine($"^FO655, {currentY}^A0N,25,32^CI0^FR^FD{qty}^FS");
                sbLots.AppendLine($"^FO783, {currentY}^A0N,25,32^CI0^FR^FD{code}^FS");
                sbLots.AppendLine($"^BY2, 2.0^FO875, {currentY}^B3N,N,30,N,Y^FR^FD{childLot}^FS");
            }

            string zpl = ZplTemplates.WSMULTILOT_TO_SFG
                .Replace("{DynamicMultiLotBlock}", sbLots.ToString())
                .Replace("{WQty}", wQty)
                .Replace("{CQty}", cQty)
                .Replace("{UserName}", userName)
                .Replace("{TimeStamp}", timeStamp);

            await SendToPrinterAsync(printerServer, zpl);
        }


        /// <summary>
        /// 4. 翻寫 Prt_WS_RENESAS_SHIPPING
        /// </summary>
        public async Task Prt_WS_RENESAS_SHIPPING_Async(
            string lotNo, string prodNo, string wQty, string qty, string partNo, 
            string[] arrWaferId, string[] arrWaferIdChipQty, string printDate, string userId, string printerServer)
        {
            // Renesas 表單中，因為列印紙張是預先印好的表格，Y 座標間距是不規則的 (配合實體紙張格子)
            int[] yCoords = { 16, 51, 87, 121, 157, 192, 227, 262, 298, 332 };
            
            // X 座標分三欄：1~10、11~20、21~25
            var colsX = new[] { 
                new { idX = 284, qtyX = 465 }, 
                new { idX = 631, qtyX = 813 }, 
                new { idX = 983, qtyX = 1164 } 
            };

            var sbWafers = new StringBuilder();

            for (int i = 1; i <= 25; i++) // 陣列是 1-based 傳入
            {
                string wId = (arrWaferId != null && i < arrWaferId.Length) ? arrWaferId[i] : "";
                string cQtyStr = (arrWaferIdChipQty != null && i < arrWaferIdChipQty.Length) ? arrWaferIdChipQty[i] : "";

                int colIndex = (i - 1) / 10;          // 第幾欄 (0, 1, 2)
                int rowIndex = (i - 1) % 10;          // 第幾列 (0~9)
                
                int labelX = colsX[colIndex].idX - 28; // 標籤號碼 "1:", "2:" 的 X 座標，比內容前面一點點
                int currentY = yCoords[rowIndex];

                sbWafers.AppendLine($"^FO{labelX},{currentY}^A0N,30,0^CI0^FR^FD{i}:^FS");
                sbWafers.AppendLine($"^FO{colsX[colIndex].idX},{currentY}^A0N,30,0^CI0^FR^FD{wId}^FS");
                sbWafers.AppendLine($"^FO{colsX[colIndex].qtyX},{currentY}^A0N,30,0^CI0^FR^FD{cQtyStr}^FS");
            }

            string zpl = ZplTemplates.WS_RENESAS_SHIPPING
                .Replace("{DynamicWaferBlock}", sbWafers.ToString())
                .Replace("{LotNo}", lotNo)
                .Replace("{PartNo}", partNo)
                .Replace("{ProdNo}", prodNo)
                .Replace("{WQty}", wQty)
                .Replace("{Qty}", qty)
                .Replace("{UserID}", userId)
                .Replace("{PrintDate}", printDate);

            await SendToPrinterAsync(printerServer, zpl);
        }
        /// <summary>
        /// 1. 翻寫 Prt_FT_FT_BIN_CARD_LABEL (Pass / Fail 分流)
        /// </summary>
        public async Task Prt_FT_FT_BIN_CARD_LABEL_Async(
            string lotNo, string prodNo, string testMode, string stepId, 
            string printerServer, bool isPassBin, string carGradeFlag = "")
        {
            // 依據是否為 Pass Bin 選擇對應底圖
            string zpl = isPassBin ? ZplTemplates.FT_BIN_CARD_PASS : ZplTemplates.FT_BIN_CARD_FAIL;

            zpl = zpl.Replace("{LotNo}", lotNo)
                     .Replace("{ProdNo}", prodNo)
                     .Replace("{StepID}", stepId)
                     .Replace("{TestMode}", testMode);

            await SendToPrinterAsync(printerServer, zpl);
        }

        /// <summary>
        /// 2. 翻寫 Prt_WS_WS_ENG_LOC_LABEL (包含資料庫查詢)
        /// </summary>
        public async Task Prt_WS_WS_ENG_LOC_LABEL_Async(string lotNo, string printerServer)
        {
            // 查詢資料庫屬性
            var dbData = await _repo.GetEngLocMasterDataAsync(lotNo);
            
            // 若查無資料，給予預設空值以防報錯
            if (dbData == null) dbData = new EngLocMasterData();

            string zpl = ZplTemplates.WS_ENG_LOC_LABEL
                .Replace("{LotNo}", lotNo)
                .Replace("{WQty}", dbData.WQty ?? "")
                .Replace("{OwnerId}", dbData.OwnerId ?? "")
                .Replace("{OwnerDep}", dbData.OwnerDep ?? "");

            await SendToPrinterAsync(printerServer, zpl);
        }

        /// <summary>
        /// 3. 翻寫 Prt_FT_BOX_COUNTING_LABEL
        /// 註: VB6 中的 oLot (FwLot) 被轉換為獨立的欄位傳入，符合 Web API DTO 精神
        /// </summary>
        public async Task Prt_FT_BOX_COUNTING_LABEL_Async(string lotId, string ipn, string printCount, string printerServer)
        {
            string zpl = ZplTemplates.FT_BOX_COUNTING_LABEL
                .Replace("{LotId}", lotId)
                .Replace("{IPN}", ipn)
                .Replace("{PrintCount}", printCount);

            await SendToPrinterAsync(printerServer, zpl);
        }

        /// <summary>
        /// 4. 翻寫 RefreshPrinterServer
        /// 在 Controller 層中，您可以直接把這個方法包裝成 API 回傳給 Vue 的 DropdownList
        /// </summary>
        public async Task<IEnumerable<string>> GetMappedPrinterServersAsync(string stage, string specNo)
        {
            return await _repo.GetMappedPrinterServersAsync(stage, specNo);
        }

        /// <summary>
        /// 翻寫 Prt_FT_Label_PACK_INFO (包含複雜的 Auto Fallback 查詢機制)
        /// </summary>
        public async Task Prt_FT_Label_PACK_INFO_Async(
            string lotId, string labelSpecNo, string carrierType, string boxingSpecNo, 
            string brand, string pinCount, string packageCode, bool isAutoPrint, 
            string printerServer, string ipn)
        {
            // 1. 初始化列印變數
            string printLotNo = isAutoPrint ? lotId : "NA";
            string printMode = isAutoPrint ? "Auto" : "Manual";
            string packageCodePinCount = pinCount + packageCode;

            // -------------------------------------------------------------
            // [Check-Item-1] : 查詢 Label Spec Info
            // -------------------------------------------------------------
            var specInfo = await _repo.GetLabelSpecInfoAsync(labelSpecNo, carrierType);

            // 如果第一輪沒查到
            if (specInfo == null)
            {
                if (!isAutoPrint) 
                {
                    throw new Exception("FT_Label_Info 資訊異常, 請通知主任確認CAT資訊正確性.");
                }

                // Auto Mode 的退回機制：查 IPN_MASTER 重配 labelSpecNo
                var ipnData = await _repo.GetIpnMasterForLabelSpecAsync(ipn);
                if (ipnData.Grade != null) // C# Tuples fallback check
                {
                    string grade = ipnData.Grade;
                    string bizType = ipnData.BizType;
                    bool isHOrW = packageCode == "H" || packageCode == "W";

                    if (brand == "KH" && grade == "Y" && !isHOrW) 
                        labelSpecNo = "6130K-0807.1";
                    else if (brand == "KH" && bizType == "D" && isHOrW) 
                        labelSpecNo = "6130K-0807.1";
                    else if (brand == "KH" && !isHOrW) 
                        labelSpecNo = "6130K-0807";
                    else 
                        labelSpecNo = "6130-0807";
                }

                // 重新查第二次
                specInfo = await _repo.GetLabelSpecInfoAsync(labelSpecNo, carrierType);
                if (specInfo == null)
                {
                    throw new Exception("FT_Label_Info(Default) 資訊異常, 請通知主任確認CAT資訊正確性.");
                }
            }

            // -------------------------------------------------------------
            // [Check-Item-2] : 查詢 Label Boxing Info
            // -------------------------------------------------------------
            var boxingInfo = await _repo.GetLabelBoxingInfoAsync(carrierType, boxingSpecNo, brand, pinCount, packageCode);

            // 如果第一輪沒查到
            if (boxingInfo == null)
            {
                if (!isAutoPrint) 
                {
                    throw new Exception("FT_Boxing_Info 資訊異常, 請通知主任確認CAT資訊正確性.");
                }

                // Auto Mode 的退回機制：強制設定 PinCount = ALL, PackageCode = ALL 重查
                boxingInfo = await _repo.GetLabelBoxingInfoAsync(carrierType, boxingSpecNo, brand, "ALL", "ALL");
                if (boxingInfo == null)
                {
                    throw new Exception("FT_Boxing_Info 資訊異常, 請通知主任確認CAT資訊正確性.");
                }

                // 如果重查成功，設定值要變成 ALL
                packageCodePinCount = "ALL";
            }
            else
            {
                // 如果第一次就查到，且兩者原本就是 ALL，也要設定為 ALL
                if (pinCount == "ALL" && packageCode == "ALL")
                {
                    packageCodePinCount = "ALL";
                }
            }

            // -------------------------------------------------------------
            // 組裝 ZPL 模板
            // -------------------------------------------------------------
            string zpl = ZplTemplates.FT_Label_PACK_INFO
                .Replace("{SpecialLabelSize}", specInfo.SpecialLabelSize ?? "")
                .Replace("{LabelSpecNoVer}", specInfo.LabelSpecNoVer ?? "")
                .Replace("{LabelSpecNo}", labelSpecNo ?? "")
                .Replace("{PrintLotNO}", printLotNo)
                .Replace("{PrintMode}", printMode)
                .Replace("{Serial}", specInfo.Serial ?? "")
                .Replace("{BoxingSpecNo}", boxingSpecNo ?? "")
                .Replace("{BoxingSpecNoVer}", boxingInfo.BoxingSpecNoVer ?? "")
                .Replace("{Brand}", brand ?? "")
                .Replace("{PackageCodePinCount}", packageCodePinCount)
                .Replace("{CarrierType}", carrierType ?? "")
                .Replace("{Vacuum}", boxingInfo.Vacuum ?? "")
                .Replace("{Hic}", boxingInfo.Hic ?? "")
                .Replace("{Desiccant}", boxingInfo.Desiccant ?? "")
                .Replace("{Albag}", boxingInfo.Albag ?? "")
                .Replace("{CopiesStdLabel}", specInfo.CopiesStdLabel ?? "")
                .Replace("{CopiesSpecialLabel}", specInfo.CopiesSpecialLabel ?? "")
                .Replace("{CopiesMergeLabel}", specInfo.CopiesMergeLabel ?? "")
                .Replace("{MergeLabelSpec}", specInfo.MergeLabelSpec ?? "")
                .Replace("{SpecialStickReel}", specInfo.SpecialStickReel ?? "")
                .Replace("{SpecialStickAlbag}", specInfo.SpecialStickAlbag ?? "")
                .Replace("{SpecialPositionBox_1}", specInfo.SpecialPositionBox_1 ?? "")
                .Replace("{SpecialPositionBox_2}", specInfo.SpecialPositionBox_2 ?? "");

            // 發送至印表機
            await SendToPrinterAsync(printerServer, zpl);
        }
        // 1. 翻寫 GetLabelSpecByStage
        public async Task<IEnumerable<string>> GetLabelSpecByStageAsync(string stage)
        {
            return await _repo.GetLabelSpecByStageAsync(stage);
        }

        // 2. 翻寫 GetProdType (解析逗號回傳 tuple)
        public async Task<(string ProdType, string CriteriaData)> GetProdTypeAsync(string ipn, string criteriaString)
        {
            string rawData = await _repo.GetProdTypeAsync(ipn, criteriaString);
            string prodType = "";
            string criteriaData = "Y"; // 一切從嚴，沒有值預設為 Y

            if (!string.IsNullOrEmpty(rawData) && rawData.Contains(","))
            {
                var parts = rawData.Split(',');
                prodType = parts[0].Trim();
                criteriaData = parts[1].Trim();
            }
            return (prodType, criteriaData);
        }

        // 3. 翻寫 GetCpHot / GetMainProd
        public async Task<string> GetCpHotAsync(string lotId, string ipn)
        {
            bool isHot = await _repo.CheckCpHotAsync(lotId);
            if (isHot) return "Y";

            // 如果 TRACKLOT 找不到，去查 MAIN_PROD
            var ipnData = await _repo.GetIpnWaferLevelDataAsync(ipn);
            if (ipnData != null)
            {
                string prodCode = ipn.Substring(0, 4) + ipnData.MaskOption;
                bool isMainProd = await _repo.CheckMainProdAsync(prodCode, ipn, ipnData.WaferLevel);
                if (isMainProd) return "Y";
            }
            return "";
        }

        // 4. 翻寫 Prt_WS_CP_SMALL_LABEL
        public async Task Prt_WS_CP_SMALL_LABEL_Async(
            string lotNo, string prodNo, string wQty, string cQty, string owner, string bizType, 
            string printerServer, string inputWaferId = "", string inputFabLotId = "", bool bPrintLotList = false)
        {
            // (1) 取得 BizType (若前端未傳，從後端補)
            if (string.IsNullOrEmpty(bizType))
            {
                var typeData = await GetProdTypeAsync(prodNo, "");
                bizType = typeData.ProdType;
            }

            // (2) 取得 IPN 基本資料
            var ipnData = await _repo.GetIpnWaferLevelDataAsync(prodNo) ?? new IpnWaferLevelData();
            string prodCode = prodNo.Length >= 4 ? prodNo.Substring(0, 4) : prodNo;
            prodCode += ipnData.MaskOption;
            string prodBody = prodNo.Length >= 4 ? prodNo.Substring(0, 4) : prodNo; // Left(sProdCode, 4)

            // (3) 取得 Hot 標記
            string hotLotFlag = await GetCpHotAsync(lotNo, prodNo);

            // (4) 取得 Lot 詳細屬性
            var lotAttrs = await _repo.GetWsSmallLabelDataAsync(lotNo) ?? new WsSmallLabelDbData();
            
            string printFlag = await _repo.GetReworkPrintFlagAsync(lotAttrs.SapRwNo);
            string sapRwNo = printFlag == "Y" ? lotAttrs.SapRwNo : "";

            string printFabLotId = !string.IsNullOrEmpty(inputFabLotId) ? inputFabLotId : lotAttrs.FabLotId;
            string rawWaferStr = !string.IsNullOrEmpty(inputWaferId) ? inputWaferId : lotAttrs.WaferId;

            // (5) Wafer 切片與分組邏輯 (重構 VB6 的 Mid 取代邏輯)
            var parsedWafers = (rawWaferStr ?? "")
                .Split(new[] { ';', ',' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(w => int.TryParse(w, out int val) ? val : 0)
                .ToList();

            var waferList = new List<string>();
            for (int i = 1; i <= 25; i++)
            {
                waferList.Add(parsedWafers.Contains(i) ? i.ToString("D2") : "__");
            }

            // 將陣列依照 VB6 的數量切成三組 (第1~9, 10~18, 19~25 片)
            string waferNo1 = string.Join(",", waferList.Take(9));
            string waferNo2 = string.Join(",", waferList.Skip(9).Take(9));
            string waferNo3 = string.Join(",", waferList.Skip(18).Take(7));

            // (6) 解析 FuSa IPN
            string fusaIpn = prodNo;
            if (ipnData.BomLevel == "DB" && (ipnData.WaferLevel == "KGD-AEB" || ipnData.WaferLevel == "MXO1") && !string.IsNullOrEmpty(lotAttrs.FgIpn))
            {
                fusaIpn = lotAttrs.FgIpn;
            }
            string fusa = await _repo.GetFusaAsync(fusaIpn);

            // (7) 根據 bPrintLotList 決定排版
            string zpl;
            if (bPrintLotList)
            {
                var mergeList = await _repo.GetVirtualMergeListAsync(lotNo);
                var sbLotList = new StringBuilder();
                int currentY = 480;

                int listIndex = 1;
                foreach (var item in mergeList)
                {
                    sbLotList.AppendLine($"^FO30,{currentY}^A0N,48.36^CI0^FR^FDLotId{listIndex:D2} : {item.LotId}   {item.Pcs} pc^FS");
                    sbLotList.AppendLine($"^BY2,2.0^FO624,{currentY}^B3N,N,48,N,N^FD{item.LotId}^FS");
                    currentY += 60;
                    listIndex++;
                }

                zpl = ZplTemplates.WS_CP_VIRTUAL_MERGE_LIST
                    .Replace("{DynamicLotListBlock}", sbLotList.ToString());
            }
            else
            {
                zpl = ZplTemplates.WS_CP_SMALL_LABEL;
            }

            // (8) 變數替換與發送
            zpl = zpl.Replace("{LotNo}", lotNo)
                     .Replace("{PrintFabLotId}", printFabLotId ?? "")
                     .Replace("{ProdBody}", prodBody)
                     .Replace("{WaferLevel}", ipnData.WaferLevel ?? "")
                     .Replace("{HotLotFlag}", hotLotFlag)
                     .Replace("{WQty}", wQty)
                     .Replace("{CQty}", cQty)
                     .Replace("{BizType}", bizType)
                     .Replace("{ErunTicNo}", lotAttrs.ErunTicNo ?? "")
                     .Replace("{FuSa}", fusa ?? "")
                     .Replace("{SapRwNo}", sapRwNo ?? "")
                     .Replace("{MfgTicNo}", lotAttrs.MfgTicNo ?? "")
                     .Replace("{WaferNo1}", waferNo1)
                     .Replace("{WaferNo2}", waferNo2)
                     .Replace("{WaferNo3}", waferNo3);

            await SendToPrinterAsync(printerServer, zpl);
        }
    }
}

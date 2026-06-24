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
    }
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
    }
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
    }
}

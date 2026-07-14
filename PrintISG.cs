using System.Collections.Generic;

namespace MES.Net.Shared.DTOs.Print
{
    public class PrintISGRequest
    {
        public string ShipNo { get; set; }
    }

    public class PrintISGResponse
    {
        public string ShipNo { get; set; }
        public string ShipType { get; set; } // "FT" (11TU開頭) 或 "WS"
        public string LotId { get; set; }
        public string IPN { get; set; }
        public int WQty { get; set; }
        public int CQty { get; set; }
        public int PassQty { get; set; }
        public int FailQty { get; set; }

        // PO Info
        public string PoNo { get; set; }
        public string PoItem { get; set; }
        public string CustLotNo { get; set; }
        public string OrPoNo { get; set; }
        public string OrPoItem { get; set; }

        public string WaferNoList { get; set; } // WS 專用：串接好的WaferNo (用分號隔開)

        public List<MergeLotInfo> MergeLots { get; set; } = new List<MergeLotInfo>();
    }

    public class MergeLotInfo
    {
        public string MergeLotId { get; set; }
        public int MergeCQty { get; set; }
        public string CustLotNo { get; set; }
        public string CustPoNo { get; set; }
        public string CustPoItem { get; set; }
    }
}


using Dapper;
using System.Data;
using System.Collections.Generic;
using System.Threading.Tasks;
using MES.Net.Shared.DTOs.Print;
using Oracle.ManagedDataAccess.Client;
using System.Configuration;

namespace MES.Net.Infrastructure.Repository.Print
{
    public interface IPrintISGRepository
    {
        Task<PrintISGResponse> GetShipFGAsync(string shipNo);
        Task<dynamic> GetLotInfoAsync(string lotId);
        Task<IEnumerable<MergeLotInfo>> GetMergeLotsAsync(string lotId);
        Task<IEnumerable<string>> GetWaferNumbersAsync(string lotId);
    }

    public class PrintISGRepository : IPrintISGRepository
    {
        private readonly IDbConnection _dbConnection;
        private static readonly string _connString = ConfigurationManager.ConnectionStrings["MESDB"]?.ConnectionString;
        public PrintISGRepository()
        {
            _dbConnection = new OracleConnection(_connString);
        }

        public async Task<PrintISGResponse> GetShipFGAsync(string shipNo)
        {
            string sql = @"
                SELECT LOTID, IPN, SHIPQTY, WAFERQTY as WQty, CHIPQTY as CQty, PASSQTY, FAILQTY 
                FROM SHIP_FG 
                WHERE TICKETNO = :ShipNo AND DELETE_FLAG = 'N' 
                ORDER BY CREATE_TIME DESC";

            return await _dbConnection.QueryFirstOrDefaultAsync<PrintISGResponse>(sql, new { ShipNo = shipNo });
        }

        public async Task<dynamic> GetLotInfoAsync(string lotId)
        {
            string sql = @"
                SELECT CUR_PO_NO as PoNo, CUR_PO_ITEM as PoItem, 
                       CUSTLOTID as CustLotNo, CUSTPOITEM as OrPoItem, CUSTPONO as OrPoNo
                FROM LOT_INFO 
                WHERE LOT_ID = :LotId";

            return await _dbConnection.QueryFirstOrDefaultAsync(sql, new { LotId = lotId });
        }

        public async Task<IEnumerable<MergeLotInfo>> GetMergeLotsAsync(string lotId)
        {
            string sql = @"
                SELECT b.valdata as MergeLotId, 
                       CAST(d.valdata AS INT) as MergeCQty,
                       e.CUSTLOTID as CustLotNo, 
                       e.CUSTPOITEM as CustPoItem, 
                       e.CUSTPONO as CustPoNo
                FROM fwmerge a
                INNER JOIN fwmerge_pn2m b ON a.sysid = b.fromid AND b.linkname = 'childLotCollection'
                INNER JOIN fwlot c ON b.valdata = c.appid
                INNER JOIN fwlot_pn2m d ON c.sysid = d.fromid AND d.keydata = 'ChipQty'
                LEFT JOIN LOT_INFO e ON b.valdata = e.lotid
                WHERE a.wipid = :LotId 
                  AND SUBSTR(a.mergestepid, 1, 5) = '70710'";

            return await _dbConnection.QueryAsync<MergeLotInfo>(sql, new { LotId = lotId });
        }

        public async Task<IEnumerable<string>> GetWaferNumbersAsync(string lotId)
        {
            string sql = @"
                SELECT c.componentid 
                FROM fwlot a
                INNER JOIN fwlot_n2m b ON a.SysId = b.fromid
                INNER JOIN fwcomponent c ON b.toid = c.sysid
                INNER JOIN fwcomponent_pn2m d ON c.sysid = d.fromid
                WHERE a.appid = :LotId 
                ORDER BY c.componentid";

            return await _dbConnection.QueryAsync<string>(sql, new { LotId = lotId });
        }
    }
}

using System;
using System.Linq;
using System.Threading.Tasks;
using MES.Net.Shared.DTOs.Print;
using MES.Net.Infrastructure.Repository.Print;
using MES.Net.Infrastructure.Messaging;

namespace MES.Net.Application.Services.Print
{
    public interface IPrintISGService
    {
        Task<PrintISGResponse> GetISGFormDataAsync(string shipNo);
    }

    public class PrintISGService : IPrintISGService
    {
        private readonly IPrintISGRepository _printISGRepo;
        private readonly IMessageQueueService _mqService;

        private readonly PrintISGRepository _repository;


        public PrintISGService()
        {
            _printISGRepo = new PrintISGRepository();
        }
        public PrintISGService(IPrintISGRepository printISGRepo)
        {
            _printISGRepo = printISGRepo;
        }

        public async Task<PrintISGResponse> GetISGFormDataAsync(string shipNo)
        {
            shipNo = shipNo?.ToUpper()?.Trim();
            if (string.IsNullOrEmpty(shipNo))
            {
                throw new ArgumentException("請輸入出貨單號! (Please key in ShipNo!)");
            }

            // 1. 取得主檔資料
            var response = await _repository.GetShipFGAsync(shipNo);
            if (response == null)
            {
                throw new Exception($"出貨單號({shipNo})不存在!");
            }

            response.ShipNo = shipNo;
            response.ShipType = shipNo.StartsWith("11TU") ? "FT" : "WS";

            // 2. 取得 Lot Info (PO相關資訊)
            var lotInfo = await _repository.GetLotInfoAsync(response.LotId);
            if (lotInfo != null)
            {
                response.PoNo = lotInfo.PoNo;
                response.PoItem = lotInfo.PoItem;
                response.CustLotNo = lotInfo.CustLotNo;
                response.OrPoNo = lotInfo.OrPoNo;
                response.OrPoItem = lotInfo.OrPoItem;
            }

            // 3. 依據 ShipType 處理細節
            if (response.ShipType == "FT")
            {
                var mergeLots = await _repository.GetMergeLotsAsync(response.LotId);
                response.MergeLots = mergeLots.ToList();

                // VB 邏輯：計算剩餘 CQty (總 CQty 減去所有 Child Qty 的總和)
                int childQtyTotal = response.MergeLots.Sum(x => x.MergeCQty);
                int remainingQty = response.CQty - childQtyTotal;

                // 若需依原有邏輯將剩餘數量推入明細，可在此擴充
            }
            else // "WS" 邏輯
            {
                var waferNos = await _repository.GetWaferNumbersAsync(response.LotId);
                response.WaferNoList = string.Join(";", waferNos);
            }

            return response;
        }
    }
}

using MES.Net.Application.Services.Print;
using MES.Net.Web.Filters;
using System;
using System.Threading.Tasks;
using System.Web.Http;

namespace MES.Net.Web.Controllers.Print
{
    [RoutePrefix("api/print/print-ISG")]
    public class PrintISGController : ApiController
    {
        private readonly IPrintISGService _service;

        public PrintISGController()
        {
        }
        public PrintISGController(IPrintISGService service)
        {
            _service = service;
        }

        /// <summary>
        /// 取得 ISG 出貨單報表資料
        /// </summary>
        [HttpPost, Route("shipNo"), AuthorizeToken]
        // 🌟 修正1：改用 [FromBody] PrintISGRequest 來接收前端傳來的 JSON
        public async Task<IHttpActionResult> GetISGData([FromBody] PrintISGRequest request)
        {
            try
            {
                // 防呆：確保前端有傳值過來
                if (request == null || string.IsNullOrWhiteSpace(request.ShipNo))
                {
                    return Ok(new { Success = false, Message = "請輸入出貨單號! (Please key in ShipNo!)" });
                }

                var data = await _service.GetISGFormDataAsync(request.ShipNo);
                return Ok(new { Success = true, Data = data });
            }
            catch (Exception ex)
            {
                // 🌟 修正2：不論是 ArgumentException 還是查無資料的 Exception，
                // 都將 ex.Message 放在 Message 屬性中，並回傳 Ok (HTTP 200)。
                // 這樣前端的 mesfetch 就不會進到 catch(error) 區塊，能正確顯示提示。
                return Ok(new { Success = false, Message = ex.Message });
            }
        }
    }
}

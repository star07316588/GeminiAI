

namespace YourApp.Models.ISGForm
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
using YourApp.Models.ISGForm;

namespace YourApp.Repositories
{
    public interface IISGFormRepository
    {
        Task<PrintISGResponse> GetShipFGAsync(string shipNo);
        Task<dynamic> GetLotInfoAsync(string lotId);
        Task<IEnumerable<MergeLotInfo>> GetMergeLotsAsync(string lotId);
        Task<IEnumerable<string>> GetWaferNumbersAsync(string lotId);
    }

    public class ISGFormRepository : IISGFormRepository
    {
        private readonly IDbConnection _dbConnection;

        public ISGFormRepository(IDbConnection dbConnection)
        {
            _dbConnection = dbConnection;
        }

        public async Task<PrintISGResponse> GetShipFGAsync(string shipNo)
        {
            string sql = @"
                SELECT LOTID, IPN, SHIPQTY, WAFERQTY as WQty, CHIPQTY as CQty, PASSQTY, FAILQTY 
                FROM SHIP_FG 
                WHERE TICKETNO = @ShipNo AND DELETE_FLAG = 'N' 
                ORDER BY CREATE_TIME DESC";

            return await _dbConnection.QueryFirstOrDefaultAsync<PrintISGResponse>(sql, new { ShipNo = shipNo });
        }

        public async Task<dynamic> GetLotInfoAsync(string lotId)
        {
            string sql = @"
                SELECT CUR_PO_NO as PoNo, CUR_PO_ITEM as PoItem, 
                       CUSTLOTID as CustLotNo, CUSTPOITEM as OrPoItem, CUSTPONO as OrPoNo
                FROM LOT_INFO 
                WHERE LOT_ID = @LotId";

            return await _dbConnection.QueryFirstOrDefaultAsync(sql, new { LotId = lotId });
        }

        public async Task<IEnumerable<MergeLotInfo>> GetMergeLotsAsync(string lotId)
        {
            // 轉換原本 VB 的 Oracle 查詢 (加入JOIN取代舊語法)
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
                WHERE a.wipid = @LotId 
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
                WHERE a.appid = @LotId 
                ORDER BY c.componentid";

            return await _dbConnection.QueryAsync<string>(sql, new { LotId = lotId });
        }
    }
}

using System;
using System.Linq;
using System.Threading.Tasks;
using YourApp.Models.ISGForm;
using YourApp.Repositories;

namespace YourApp.Services
{
    public interface IISGFormService
    {
        Task<PrintISGResponse> GetISGFormDataAsync(string shipNo);
        // Task<byte[]> GenerateISGExcelAsync(string shipNo); // 擴充：若需由後端產出Excel
    }

    public class ISGFormService : IISGFormService
    {
        private readonly IISGFormRepository _repository;

        public ISGFormService(IISGFormRepository repository)
        {
            _repository = repository;
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

                // 若有餘額，通常會將母批也當作一筆子明細呈現 (對齊 VB 邏輯中的 .SetText miSpdPrintCol_FT_MergeCQty...)
                if (remainingQty > 0)
                {
                     // 您可以視前端需求決定是否在這裡推入一筆剩餘數量的明細
                }
            }
            else // "WS" 邏輯
            {
                var waferNos = await _repository.GetWaferNumbersAsync(response.LotId);
                response.WaferNoList = string.Join(";", waferNos); // 串接 WaferNo
            }

            return response;
        }
    }
}

namespace MesApi.Models
{
    // API 接收的參數
    public class QueryLotRequest
    {
        public string LotId { get; set; }
        public bool IsHistoryTab { get; set; } // 對應 ssTabQuery.Tab = 1
    }

    // 查詢回傳的資料結構
    public class LotMergeInfo
    {
        public string ParentLotId { get; set; }
        public string SlotNo { get; set; }
        public string ChildLotId { get; set; }
        public string WaferId { get; set; }
        public string WQty { get; set; }
        public string CQty { get; set; }
        public string Source { get; set; }
        public string Ipn { get; set; }
        public string Owner { get; set; }
        public string MiddleLotId { get; set; }
        public string WaferChipQty { get; set; }
        public string ErrMsg { get; set; }
        public string FabCreateTime { get; set; }
        public string RiskLevel { get; set; }
        public string MergeStepId { get; set; }
        public string CreateUserId { get; set; }
        public string CreateTime { get; set; }
        public string DeleteTime { get; set; }
        public string RowId { get; set; }
        public string UpdateTable { get; set; }
    }
}

using Microsoft.AspNetCore.Mvc;
using MesApi.Models;
using MesApi.Services;

namespace MesApi.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class LotController : ControllerBase
    {
        private readonly ILotService _lotService;

        public LotController(ILotService lotService)
        {
            _lotService = lotService;
        }

        [HttpGet("query")]
        public async Task<IActionResult> ExecQuery([FromQuery] QueryLotRequest request)
        {
            // Condition Checking (對應原本的請輸入查詢條件)
            if (string.IsNullOrWhiteSpace(request.LotId))
            {
                return BadRequest(new { Message = "Pls input Query Criteria !! 請輸入查詢條件 !!" });
            }

            try
            {
                var result = await _lotService.GetLotMergeInfoAsync(request.LotId, request.IsHistoryTab);
                
                if (result == null || !result.Any())
                {
                    return NotFound(new { Message = $"Lot '{request.LotId}' is no data found !! 查無資料 !!" });
                }

                return Ok(result);
            }
            catch (ArgumentException ex)
            {
                // 對應 "Lot is not exist in MES" 的例外處理
                return BadRequest(new { Message = ex.Message });
            }
            catch (Exception ex)
            {
                // 對應原本的 ErrorHandler
                // 實務上這裡會把 ex 寫入 NLog 或 Serilog
                return StatusCode(500, new { Message = "Fail to execute application, please call IT support!! 程式執行失敗, 請洽IT人員處理" });
            }
        }
    }
}

using MesApi.Models;
using MesApi.Repositories;

namespace MesApi.Services
{
    public interface ILotService
    {
        Task<IEnumerable<LotMergeInfo>> GetLotMergeInfoAsync(string lotId, bool isHistoryTab);
    }

    public class LotService : ILotService
    {
        private readonly ILotRepository _lotRepo;

        public LotService(ILotRepository lotRepo)
        {
            _lotRepo = lotRepo;
        }

        public async Task<IEnumerable<LotMergeInfo>> GetLotMergeInfoAsync(string lotId, bool isHistoryTab)
        {
            // Step-1 : 先檢查是否為 Virtual Lot ?
            bool isVirtualLot = await _lotRepo.CheckVirtualLotExistsAsync(lotId);

            if (isVirtualLot)
            {
                // 查 Virtual Merge
                return await _lotRepo.GetVirtualMergeInfoAsync(lotId, isHistoryTab);
            }
            else
            {
                // 查一般 Lot
                bool isAllLot = await _lotRepo.CheckAllLotExistsAsync(lotId);
                
                if (!isAllLot)
                {
                    throw new ArgumentException($"Lot '{lotId}' is not exist in MES !! 不存在於 MES !!");
                }

                // 查測後併批 tables (Inkless Merge)
                return await _lotRepo.GetInklessMergeInfoAsync(lotId, isHistoryTab);
            }
        }
    }
}

using MesApi.Models;
using Dapper; // 建議使用 Dapper 執行 SQL
using System.Data;

namespace MesApi.Repositories
{
    public interface ILotRepository
    {
        Task<bool> CheckVirtualLotExistsAsync(string lotId);
        Task<bool> CheckAllLotExistsAsync(string lotId);
        Task<IEnumerable<LotMergeInfo>> GetVirtualMergeInfoAsync(string lotId, bool isHistoryTab);
        Task<IEnumerable<LotMergeInfo>> GetInklessMergeInfoAsync(string lotId, bool isHistoryTab);
    }

    public class LotRepository : ILotRepository
    {
        private readonly IDbConnection _dbConnection; // 由 DI 注入連線 (例如 OracleConnection 或 SqlConnection)

        public LotRepository(IDbConnection dbConnection)
        {
            _dbConnection = dbConnection;
        }

        public async Task<bool> CheckVirtualLotExistsAsync(string lotId)
        {
            string sql = "SELECT 1 FROM TBL_VIRTUAL_LOT WHERE LOTID = @LotId";
            var result = await _dbConnection.QueryFirstOrDefaultAsync<int?>(sql, new { LotId = lotId });
            return result != null;
        }

        public async Task<bool> CheckAllLotExistsAsync(string lotId)
        {
            string sql = "SELECT 1 FROM TBL_ALL_LOT WHERE LOTID = @LotId";
            var result = await _dbConnection.QueryFirstOrDefaultAsync<int?>(sql, new { LotId = lotId });
            return result != null;
        }

        public async Task<IEnumerable<LotMergeInfo>> GetVirtualMergeInfoAsync(string lotId, bool isHistoryTab)
        {
            string sql = @"
                SELECT 
                    VIRTUALLOTID as ParentLotId, SLOTNO as SlotNo, 
                    LOTID as ChildLotId, WAFERID as WaferId, 
                    '1' as WQty, '' as CQty, '' as Source, '' as Ipn, '' as Owner, '' as MiddleLotId, '' as WaferChipQty, 
                    '' as ErrMsg, '' as FabCreateTime, '' as RiskLevel, 
                    MERGESTEPID as MergeStepId, CREATEUSERID as CreateUserId, 
                    CREATETIME as CreateTime, DELETETIME as DeleteTime, 
                    ROWID as RowId, 'TBL_VIRTUAL_MERGE' as UpdateTable 
                FROM TBL_VIRTUAL_MERGE 
                WHERE VIRTUALLOTID = @LotId ";

            if (!isHistoryTab)
            {
                sql += " AND NVL(DELETEFLAG, 'N') <> 'Y'";
            }

            return await _dbConnection.QueryAsync<LotMergeInfo>(sql, new { LotId = lotId });
        }

        public async Task<IEnumerable<LotMergeInfo>> GetInklessMergeInfoAsync(string lotId, bool isHistoryTab)
        {
            string condition = isHistoryTab ? "" : "AND NVL(DELETEFLAG, 'N') <> 'Y'";

            // 使用 UNION 合併兩張表，與原始 VB6 邏輯一致
            string sql = $@"
                SELECT 
                    PARENTLOTID as ParentLotId, SLOTNO as SlotNo, 
                    CHILDLOTID as ChildLotId, WAFERID as WaferId, 
                    WQTY as WQty, CQTY as CQty, SOURCE as Source, IPN as Ipn, 
                    OWNER as Owner, MIDDLELOTID as MiddleLotId, 
                    WAFERCHIPQTY as WaferChipQty, ERRMSG as ErrMsg, 
                    FABCREATETIME as FabCreateTime, RISKLEVEL as RiskLevel, 
                    '' as MergeStepId, CREATEUSERID as CreateUserId, 
                    CREATETIME as CreateTime, DELETETIME as DeleteTime, 
                    ROWID as RowId, 'TBL_INKLESS_MERGE_LIST' as UpdateTable 
                FROM TBL_INKLESS_MERGE_LIST 
                WHERE PARENTLOTID = @LotId {condition}
                
                UNION 
                
                SELECT 
                    PARENTLOTID as ParentLotId, SLOTNO as SlotNo, 
                    CHILDLOTID as ChildLotId, WAFERID as WaferId, 
                    WQTY as WQty, CQTY as CQty, SOURCE as Source, IPN as Ipn, 
                    OWNER as Owner, MIDDLELOTID as MiddleLotId, 
                    WAFERCHIPQTY as WaferChipQty, ERRMSG as ErrMsg, 
                    FABCREATETIME as FabCreateTime, RISKLEVEL as RiskLevel, 
                    '' as MergeStepId, CREATEUSERID as CreateUserId, 
                    CREATETIME as CreateTime, DELETETIME as DeleteTime, 
                    ROWID as RowId, 'TBL_INKLESS_MERGE_LIST' as UpdateTable 
                FROM TBL_INKLESS_MERGE_LIST_HIS 
                WHERE PARENTLOTID = @LotId {condition}";

            return await _dbConnection.QueryAsync<LotMergeInfo>(sql, new { LotId = lotId });
        }
    }
}

namespace MesApi.Models
{
    // API 接收的 POST Payload
    public class ResendSapRequest
    {
        public string LotId { get; set; }
        public string UserId { get; set; } // 對應 oUser.UserName
        
        // 對應 spdLotInfo 裡的所有列資料
        public List<ResendSapItem> Items { get; set; } 
    }

    // 每一列需要更新的關鍵資料
    public class ResendSapItem
    {
        public string UpdateTable { get; set; } // 對應 miSpdLotInfo_Table (例如: TBL_VIRTUAL_MERGE)
        public string RowId { get; set; }       // 對應 miSpdLotInfo_RowId
    }
}

using Microsoft.AspNetCore.Mvc;
using MesApi.Models;
using MesApi.Services;

namespace MesApi.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class LotController : ControllerBase
    {
        private readonly ILotService _lotService;

        public LotController(ILotService lotService)
        {
            _lotService = lotService;
        }

        // ... 前面的 ExecQuery 與 ExecQuerySorterControl ...

        [HttpPost("resend-sap")]
        public async Task<IActionResult> ResendSap([FromBody] ResendSapRequest request)
        {
            // Condition Checking (檢查是否有輸入 LotID，以及前端是否有傳入陣列資料)
            if (string.IsNullOrWhiteSpace(request.LotId) || request.Items == null || !request.Items.Any())
            {
                return BadRequest(new { Message = $"Pls input Query Criterial !! / Lot '{request.LotId}' is no data found !! 請輸入查詢條件或查無資料 !!" });
            }

            try
            {
                // 執行更新動作
                await _lotService.ResendSapAsync(request);

                // 對應成功的 MsgBox
                return Ok(new { 
                    Message = $"Lot '{request.LotId}' data is updated !! 資料已更新, 稍後將重傳給SAP !!" 
                });
            }
            catch (Exception ex)
            {
                // 系統發生預期外的錯誤，對應原本的 ErrorHandler
                return StatusCode(500, new { Message = "Fail to execute application, please call IT support!! 程式執行失敗, 請洽IT人員處理", Details = ex.Message });
            }
        }
    }
}

using MesApi.Models;
using MesApi.Repositories;

namespace MesApi.Services
{
    public interface ILotService
    {
        // ... 前面的方法 ...
        Task ResendSapAsync(ResendSapRequest request);
    }

    public class LotService : ILotService
    {
        private readonly ILotRepository _lotRepo;

        public LotService(ILotRepository lotRepo)
        {
            _lotRepo = lotRepo;
        }

        // ... 前面的方法 ...

        public async Task ResendSapAsync(ResendSapRequest request)
        {
            // 將傳入的陣列交給資料層執行 Update
            await _lotRepo.UpdateSendSapFlagAsync(request.Items, request.UserId);
        }
    }
}

using MesApi.Models;
using Dapper;
using System.Data;

namespace MesApi.Repositories
{
    public interface ILotRepository
    {
        // ... 前面的方法 ...
        Task UpdateSendSapFlagAsync(IEnumerable<ResendSapItem> items, string userId);
    }

    public class LotRepository : ILotRepository
    {
        private readonly IDbConnection _dbConnection;

        public LotRepository(IDbConnection dbConnection)
        {
            _dbConnection = dbConnection;
        }

        // ... 前面的方法 ...

        public async Task UpdateSendSapFlagAsync(IEnumerable<ResendSapItem> items, string userId)
        {
            // 產生更新時間 (對應原本的 to_char(sysdate, 'yyyymmddhh24miss') || '00')
            string updateTime = DateTime.Now.ToString("yyyyMMddHHmmss") + "00";

            if (_dbConnection.State != ConnectionState.Open)
            {
                _dbConnection.Open();
            }

            // 開啟資料庫交易
            using var transaction = _dbConnection.BeginTransaction();

            try
            {
                foreach (var item in items)
                {
                    // ⚠️ 注意: 資料表名稱 (item.UpdateTable) 無法使用 @參數化，必須用字串插值。
                    // 由於這些 TableName 是從我們先前的 API 查詢結果回傳的，相對安全，
                    // 若要嚴謹，實務上可在此加入 TableName 的白名單檢查防範 SQL Injection。
                    
                    string sql = $@"
                        UPDATE {item.UpdateTable}
                        SET SENDSAPFLAG = 'N',
                            SENDSAPTIME = '',
                            UPDATEUSERID = @UserId,
                            UPDATETIME = @UpdateTime
                        WHERE ROWID = @RowId 
                          AND NVL(SENDSAPFLAG, 'N') <> 'N'";

                    // 執行 SQL 並綁定參數與交易
                    await _dbConnection.ExecuteAsync(sql, new 
                    { 
                        UserId = userId, 
                        UpdateTime = updateTime, 
                        RowId = item.RowId 
                    }, transaction);
                }

                // 若所有迴圈內的更新都沒發生例外，則提交交易
                transaction.Commit();
            }
            catch
            {
                // 若發生任何錯誤，復原所有已更新的資料
                transaction.Rollback();
                throw; // 將例外往上拋給 Controller 處理
            }
        }
    }
}

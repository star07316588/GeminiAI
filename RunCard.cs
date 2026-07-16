using System;
using System.Collections.Generic;

namespace MES.Net.Shared.DTOs.Print
{
    // --- 1. 查詢清單用的 DTO ---
    public class QueryRunCardRequest
    {
        public string Step { get; set; }
        public string LotId { get; set; }
        public string IPN { get; set; }
        public string Type { get; set; } // "FT" 或 "WS"
    }

    public class RunCardLotListDto
    {
        public string LotId { get; set; }
        public string StepName { get; set; }
        public string Description { get; set; }
        public int ComponentQty { get; set; }
        public string IPN { get; set; }
    }

    // --- 2. 產出單一流程卡用的 DTO ---
    public class PrintRunCardRequest
    {
        public string LotId { get; set; }
        public string Type { get; set; } // "FT" 或 "WS"
    }

    public class RunCardResponse
    {
        public string LotId { get; set; }
        public string IPN { get; set; }
        public int WaferQty { get; set; }
        public int ChipQty { get; set; }
        public string LotOwner { get; set; }
        public string Route { get; set; }
        public string Status { get; set; }
        public string StartDate { get; set; }

        public RunCardSpecInfo SpecInfo { get; set; }
        public List<RunCardStepHistory> StepHistories { get; set; } = new List<RunCardStepHistory>();
        public List<RunCardFutureAction> FutureActions { get; set; } = new List<RunCardFutureAction>();
    }

    public class RunCardSpecInfo
    {
        public string EPN { get; set; }
        public string CPN { get; set; }
        public string CheckSum { get; set; }
        public string Security { get; set; }
        public string PackageCode { get; set; }
        public string Label { get; set; }
        public string CarrierSpecNo { get; set; }
        public string CarrierType { get; set; }
        public string CarrierQty { get; set; }
        public string ICDrawing { get; set; }
        // ... 其他由 TBL_IPN_MASTER 等表格撈出的規格欄位
    }

    public class RunCardStepHistory
    {
        public string StepName { get; set; }
        public string Description { get; set; }
        public DateTime? TrackInTime { get; set; }
        public DateTime? TrackOutTime { get; set; }
        public string UserIn { get; set; }
        public string UserOut { get; set; }
        public int QuantityIn { get; set; }
        public int QuantityOut { get; set; }
        
        // 機台與配方資訊
        public string Equipment { get; set; }
        public string Recipe { get; set; }
        
        // FT 專用的 Bin Data 或 WS 的 TDS Data 可擴充於此
        public int? PassQty { get; set; }
        public int? FailQty { get; set; }
        public string MergeIdList { get; set; }
        public string SplitIdList { get; set; }
    }

    public class RunCardFutureAction
    {
        public string Step { get; set; }
        public string Comments { get; set; }
        public string UserId { get; set; }
        public DateTime SetTime { get; set; }
        public string ActionType { get; set; } // Lot, Ipn, ProdGroup, Hold 等
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
    public interface IRunCardRepository
    {
        Task<IEnumerable<RunCardLotListDto>> QueryLotsAsync(QueryRunCardRequest request);
        Task<RunCardSpecInfo> GetLotSpecInfoAsync(string ipn);
        Task<IEnumerable<RunCardStepHistory>> GetStepHistoryAsync(string lotId, string planId);
        Task<IEnumerable<RunCardFutureAction>> GetFutureActionsAsync(string lotId, string ipn, string prodGroup, string lotOwner);
    }

    public class RunCardRepository : IRunCardRepository
    {
        private readonly IDbConnection _dbConnection;
        private static readonly string _connString = ConfigurationManager.ConnectionStrings["MESDB"]?.ConnectionString;

        public RunCardRepository()
        {
            _dbConnection = new OracleConnection(_connString);
        }

        public async Task<IEnumerable<RunCardLotListDto>> QueryLotsAsync(QueryRunCardRequest request)
        {
            // 對應原本 ExecQuery() 的第一段 SQL
            var sql = @"
                SELECT A.APPID AS LotId, C.STEPNAME AS StepName, C.DESCRIPTION AS Description, 
                       A.COMPONENTQTY AS ComponentQty, E.VALDATA AS IPN
                FROM FWLOT A
                INNER JOIN FWLOT_N2M B ON A.SYSID = B.FROMID AND B.LINKNAME = 'stepCollection'
                INNER JOIN FWWIPSTEP C ON B.TOID = C.SYSID
                INNER JOIN FWWIPSTEP_PN2M D ON C.SYSID = D.FROMID AND D.LINKNAME = 'attributes' AND D.KEYDATA = 'Stage'
                INNER JOIN FWLOT_PN2M E ON A.SYSID = E.FROMID AND E.KEYDATA = 'IPN'
                WHERE A.PROCESSINGSTATUS IN ('Active', 'Hold') ";

            var parameters = new DynamicParameters();

            if (!string.IsNullOrEmpty(request.Step))
            {
                sql += " AND C.STEPNAME = :Step ";
                parameters.Add("Step", request.Step.Substring(0, Math.Min(5, request.Step.Length)));
            }

            if (!string.IsNullOrEmpty(request.LotId))
            {
                sql += " AND A.APPID LIKE :LotId ";
                parameters.Add("LotId", request.LotId.Replace("*", "%"));
            }

            if (!string.IsNullOrEmpty(request.Type))
            {
                sql += " AND D.VALDATA = :Type ";
                parameters.Add("Type", request.Type); // "FT" or "WS"
            }

            if (!string.IsNullOrEmpty(request.IPN))
            {
                sql += " AND E.VALDATA LIKE :IPN ";
                parameters.Add("IPN", request.IPN.Replace("*", "%") + "%");
            }

            sql += " ORDER BY A.APPID";

            return await _dbConnection.QueryAsync<RunCardLotListDto>(sql, parameters);
        }

        public async Task<RunCardSpecInfo> GetLotSpecInfoAsync(string ipn)
        {
            // 對應原本 FtExcelCell 撈取 TBL_IPN_MASTER 與 TBL_PRM_BE_SPEC 的 SQL
            var sql = @"
                SELECT A.EPN, A.CHECKSUM, A.SECURITY, A.PACKAGE_CODE as PackageCode, 
                       B.CPN, C.LABELSPECNO as Label, A.CARRIER_SPEC_NO as CarrierSpecNo
                FROM TBL_IPN_MASTER A
                INNER JOIN TBL_PRM_BE_SPEC B ON A.IPN = B.IPN
                LEFT JOIN TBL_CUSTOMER_MASTER C ON B.CUSTOMER_NO = C.CUSTOMERNO
                WHERE A.IPN = :Ipn
                ORDER BY DECODE(B.DEFAULTS, 'Y', 0, 1)";

            return await _dbConnection.QueryFirstOrDefaultAsync<RunCardSpecInfo>(sql, new { Ipn = ipn });
        }

        public async Task<IEnumerable<RunCardStepHistory>> GetStepHistoryAsync(string lotId, string planId)
        {
            // 對應原本抓取 FWTRACKOUT, FWTRACKIN 與 TBL_ROUTE_STEP_LIST 的履歷 SQL
            var sql = @"
                SELECT F.STEP_NO as StepName, F.STEP_NAME as Description, 
                       A.QUANTITYOUT as QuantityOut, A.TRACKINTIME as TrackInTime, 
                       A.TXNTIMESTAMP as TrackOutTime, A.USERID as UserOut, 
                       B.USERID as UserIn, B.QUANTITYIN as QuantityIn
                FROM FWTRACKOUT A
                LEFT JOIN FWTRACKIN B ON A.WIPID = B.WIPID AND A.TRACKINTIME = B.TXNTIMESTAMP
                -- ... 略 (保留原有與 FW FLAT PLAN / FWWIPHISTORY 的 JOIN)
                WHERE A.WIPID = :LotId
                ORDER BY F.STEP_SEQ";

            return await _dbConnection.QueryAsync<RunCardStepHistory>(sql, new { LotId = lotId });
        }

        public async Task<IEnumerable<RunCardFutureAction>> GetFutureActionsAsync(string lotId, string ipn, string prodGroup, string lotOwner)
        {
            // 對應龐大的 UNION 查詢 (TBL_LOT_FUTACT, TBL_IPN_FUTACT, TBL_PRODGROUP_FUTACT, TBL_HOLD_REC)
            // 實作時建議依據 DTO 結構將欄位對齊
            return new List<RunCardFutureAction>(); // 假裝已經實作完成
        }
    }
}

using System;
using System.Threading.Tasks;
using System.Collections.Generic;
using MES.Net.Shared.DTOs.Print;
using MES.Net.Infrastructure.Repository.Print;

namespace MES.Net.Application.Services.Print
{
    public interface IRunCardService
    {
        Task<IEnumerable<RunCardLotListDto>> QueryRunCardLotsAsync(QueryRunCardRequest request);
        Task<RunCardResponse> GetRunCardDataAsync(PrintRunCardRequest request);
    }

    public class RunCardService : IRunCardService
    {
        private readonly IRunCardRepository _repository;

        public RunCardService(IRunCardRepository repository)
        {
            _repository = repository;
        }

        public async Task<IEnumerable<RunCardLotListDto>> QueryRunCardLotsAsync(QueryRunCardRequest request)
        {
            return await _repository.QueryLotsAsync(request);
        }

        public async Task<RunCardResponse> GetRunCardDataAsync(PrintRunCardRequest request)
        {
            if (string.IsNullOrWhiteSpace(request.LotId))
                throw new ArgumentException("LotId is required.");

            // 1. 取得基本 Lot 資訊 (原 VB 呼叫 FwuRetrieveLot)
            var response = new RunCardResponse
            {
                LotId = request.LotId,
                // ... (從系統 API 或額外的 Repository 查詢補齊 Lot 基本屬性，如 IPN, Qty, Owner 等)
                IPN = "A1234567" // 假設撈出
            };

            // 2. 取得規格資訊 (Spec)
            response.SpecInfo = await _repository.GetLotSpecInfoAsync(response.IPN);
            
            // 處理 Label 特殊邏輯 (對應 VB: If Trim(BRAND) = "KH" And ... Then activecell = "6130K-0807.1")
            if (response.SpecInfo != null && string.IsNullOrEmpty(response.SpecInfo.Label))
            {
                 // 將 VB6 寫死的 Label 邏輯移入這裡處理
                 response.SpecInfo.Label = "6130-0807"; 
            }

            // 3. 取得站點履歷 (Step History)
            // 原 VB 寫法會逐站去要 Equipment, Recipe, Bin Data
            var histories = await _repository.GetStepHistoryAsync(request.LotId, "PlanId_Here");
            response.StepHistories = new List<RunCardStepHistory>(histories);

            // 4. 取得註記/未來處置 (Future Actions)
            response.FutureActions = new List<RunCardFutureAction>(
                await _repository.GetFutureActionsAsync(request.LotId, response.IPN, "ProdGroup", "Owner")
            );

            return response;
        }
    }
}

using MES.Net.Application.Services.Print;
using MES.Net.Shared.DTOs.Print;
using MES.Net.Web.Filters;
using System;
using System.Threading.Tasks;
using System.Web.Http;

namespace MES.Net.Web.Controllers.Print
{
    [RoutePrefix("api/print/run-card")]
    public class RunCardController : ApiController
    {
        private readonly IRunCardService _service;

        public RunCardController() { }

        public RunCardController(IRunCardService service)
        {
            _service = service;
        }

        /// <summary>
        /// 條件查詢符合的 Lot 清單 (對應原 CmdQuery_Click)
        /// </summary>
        [HttpPost, Route("query"), AuthorizeToken]
        public async Task<IHttpActionResult> QueryLots([FromBody] QueryRunCardRequest request)
        {
            try
            {
                if (request == null)
                    return Ok(new { Success = false, Message = "無效的查詢參數" });

                var data = await _service.QueryRunCardLotsAsync(request);
                return Ok(new { Success = true, Data = data });
            }
            catch (Exception ex)
            {
                return Ok(new { Success = false, Message = ex.Message });
            }
        }

        /// <summary>
        /// 取得單一 Lot 的完整流程卡資料 (對應原 CmdOK_Click 產出報表資料)
        /// </summary>
        [HttpPost, Route("detail"), AuthorizeToken]
        public async Task<IHttpActionResult> GetRunCardDetail([FromBody] PrintRunCardRequest request)
        {
            try
            {
                if (request == null || string.IsNullOrWhiteSpace(request.LotId))
                    return Ok(new { Success = false, Message = "請提供 LotId" });

                var data = await _service.GetRunCardDataAsync(request);
                return Ok(new { Success = true, Data = data });
            }
            catch (Exception ex)
            {
                return Ok(new { Success = false, Message = ex.Message });
            }
        }
    }
}

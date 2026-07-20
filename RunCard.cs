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
        public string Route { get; set; }
        public int WaferQty { get; set; }
        public int ChipQty { get; set; }
        public string LotOwner { get; set; }
        public string Route { get; set; }
        public string Status { get; set; }
        public string StartDate { get; set; }

        // --- 新增：業務邏輯所需的中介與分類欄位 ---
        public string PlanId { get; set; }
        public string CurrentStepSeq { get; set; }
        public string LotType { get; set; } // Normal, Eng, Rework
        public string RunCardType { get; set; } // 標記此報表為 "FT" 或 "WS"

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
        
        // 新增 FT 專用測試數據與處置
        public int? PassQty { get; set; }
        public int? FailQty { get; set; }
        public string Bin1 { get; set; }
        public string Bin2 { get; set; }
        public string Bin3 { get; set; }
        public string Bin4 { get; set; }
        public string Bin5 { get; set; }
        public string Bin6 { get; set; }
        public double? Yield { get; set; } // 良率 (PassQty / InQty)
        
        public string ScrapComment { get; set; }
        public string MergeIdList { get; set; }
        public string SplitIdList { get; set; }
    }

    public class RunCardFutureAction
    {
        public string Step { get; set; }
        public string Comments { get; set; }
        public string UserId { get; set; }
        public DateTime SetTime { get; set; }
        public string ActionType { get; set; } 
        
        // 用於業務邏輯過濾用，不一定會顯示在畫面上
        public string IncludeEngLot { get; set; }
        public string IncludeReworkLot { get; set; }
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
        /// <summary>
        /// 實作那段巨大的 UNION 查詢 (Future Actions)
        /// </summary>
        public async Task<IEnumerable<RunCardFutureAction>> GetFutureActionsAsync(
            string lotId, string ipn, string prodGroup, string lotOwner)
        {
            // 將 VB6 中的龐大 Union 轉換為標準 SQL
            string sql = @"
                SELECT * FROM (
                    -- 1. Lot FutAct
                    SELECT STEP_NO AS Step, COMMENTS AS Comments, USER_ID AS UserId, SET_TIME AS SetTime, 
                           '' AS IncludeEngLot, '' AS IncludeReworkLot, 'Lot' AS ActionType
                    FROM TBL_LOT_FUTACT 
                    WHERE LOT_ID = :LotId AND COMMENTS IS NOT NULL AND DELETE_FLAG <> 'Y'
                    
                    UNION 
                    
                    -- 2. IPN FutAct
                    SELECT STEP_NO AS Step, COMMENTS AS Comments, USER_ID AS UserId, SET_TIME AS SetTime, 
                           INCLUDEENGLOT AS IncludeEngLot, INCLUDEREWORKLOT AS IncludeReworkLot, 'Ipn' AS ActionType
                    FROM TBL_IPN_FUTACT
                    WHERE IPN = :Ipn AND COMMENTS IS NOT NULL AND DELETE_FLAG <> 'Y'
                      AND (NVL(INCLUDELOTOWNER, 'All')='All' OR INCLUDELOTOWNER = :LotOwner 
                           OR INCLUDELOTOWNER LIKE :LotOwner || ',%' OR INCLUDELOTOWNER LIKE '%,' || :LotOwner 
                           OR INCLUDELOTOWNER LIKE '%,' || :LotOwner || ',%')

                    UNION 

                    -- 3. ProdGroup FutAct
                    SELECT STEP_NO AS Step, COMMENTS AS Comments, USER_ID AS UserId, SET_TIME AS SetTime, 
                           INCLUDEENGLOT AS IncludeEngLot, INCLUDEREWORKLOT AS IncludeReworkLot, 'ProdGroup' AS ActionType
                    FROM TBL_PRODGROUP_FUTACT
                    WHERE PROD_GROUP = :ProdGroup AND COMMENTS IS NOT NULL AND DELETE_FLAG <> 'Y'
                      AND (NVL(INCLUDELOTOWNER, 'All')='All' OR INCLUDELOTOWNER = :LotOwner 
                           OR INCLUDELOTOWNER LIKE :LotOwner || ',%' OR INCLUDELOTOWNER LIKE '%,' || :LotOwner 
                           OR INCLUDELOTOWNER LIKE '%,' || :LotOwner || ',%')

                    UNION 

                    -- 4. Hold Record
                    SELECT STEP_ID AS Step, COMMENTS AS Comments, CREATE_USER_ID AS UserId, CREATE_TIME AS SetTime, 
                           '' AS IncludeEngLot, '' AS IncludeReworkLot, 'Hold' AS ActionType
                    FROM TBL_HOLD_REC
                    WHERE LOT_ID = :LotId AND COMMENTS IS NOT NULL AND DELETE_FLAG <> 'Y'

                    UNION 

                    -- 5. Manual Test Qty
                    SELECT STEPNO AS Step, COMMENTS AS Comments, CREATE_USER_ID AS UserId, CREATE_TIME AS SetTime, 
                           '' AS IncludeEngLot, '' AS IncludeReworkLot, 'ManualTest' AS ActionType
                    FROM TBL_MANUAL_TESTQTY
                    WHERE LOT_ID = :LotId AND COMMENTS IS NOT NULL AND DELETE_FLAG <> 'Y'
                ) ORDER BY SetTime";

            return await _dbConnection.QueryAsync<RunCardFutureAction>(sql, new { 
                LotId = lotId, Ipn = ipn, ProdGroup = prodGroup ?? "", LotOwner = lotOwner ?? "" 
            });
        }

        // --- 針對 FT 站點提供擴充的細部查詢 ---
        
        public async Task<dynamic> GetFtBinDataAsync(string lotId, string stepName, string testerId)
        {
            string sql = "SELECT PASSQTY, FAILQTY, BIN1, BIN2, BIN3, BIN4, BIN5, BIN6 FROM TBL_FT_LOTCOMPLETE WHERE LOTID = :LotId AND STEPNO = :StepName AND TESTERID = :TesterId AND DELETEFLAG <> 'Y'";
            return await _dbConnection.QueryFirstOrDefaultAsync(sql, new { LotId = lotId, StepName = stepName, TesterId = testerId });
        }

        public async Task<string> GetScrapCommentAsync(string lotId, DateTime trackIn, DateTime trackOut)
        {
            string sql = "SELECT B.BRIEFDESCRIPTION FROM FWSCRAPLOT A INNER JOIN FWCOMMENT B ON A.TXNCOMMENT = B.SYSID WHERE A.WIPID = :LotId AND A.TXNTIMESTAMP >= :TrackIn AND A.TXNTIMESTAMP <= :TrackOut";
            return await _dbConnection.QueryFirstOrDefaultAsync<string>(sql, new { LotId = lotId, TrackIn = trackIn, TrackOut = trackOut });
        }
    }
}

using System;
using System.Linq;
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
            if (string.IsNullOrWhiteSpace(request.Type))
                throw new ArgumentException("RunCard Type (FT/WS) is required.");

            // 1. 初始化與撈取共用 Lot 基本資訊
            var response = new RunCardResponse
            {
                LotId = request.LotId,
                RunCardType = request.Type.ToUpper(),
                IPN = "A1234567",         // (請替換為實際查詢)
                PlanId = "YOUR_PLAN_ID",  // (請替換為實際查詢)
                CurrentStepSeq = "1234"   // (請替換為實際查詢)
            };

            string prodGroup = "PG_A";    // (請替換為實際查詢)
            string lotOwner = "OWNER_A";  // (請替換為實際查詢)

            // 2. 判斷 LotType 與 Route (共用邏輯)
            response.LotType = await _repository.GetLotTypeAsync(request.LotId);
            response.Route = await _repository.GetStepPathAsync(response.PlanId, response.CurrentStepSeq);

            // 3. ⭐️ 核心分流：依據 FT 或 WS 處理差異化的 Spec 與 History
            if (response.RunCardType == "FT")
            {
                await ProcessFtRunCardAsync(request.LotId, response);
            }
            else if (response.RunCardType == "WS")
            {
                await ProcessWsRunCardAsync(request.LotId, response);
            }
            else
            {
                throw new ArgumentException("Invalid RunCard Type. Must be 'FT' or 'WS'.");
            }

            // 4. 撈取 Future Actions (共用邏輯，傳入對齊後的 5 個參數)
            response.FutureActions = new List<RunCardFutureAction>(
                await _repository.GetFutureActionsAsync(
                    request.LotId, 
                    response.IPN, 
                    prodGroup, 
                    lotOwner, 
                    response.LotType)
            );

            return response;
        }

        // ====================================================================
        // 私有方法：專門處理 FT 的邏輯 (對應原 VB6 的 FtExcelCell)
        // ====================================================================
        private async Task ProcessFtRunCardAsync(string lotId, RunCardResponse response)
        {
            // FT 專屬的 Spec 查詢邏輯
            response.SpecInfo = await _repository.GetLotSpecInfoAsync(response.IPN);
            
            // FT 專屬的特殊 Label 邏輯 (移植自 VB6)
            if (response.SpecInfo != null && string.IsNullOrEmpty(response.SpecInfo.Label))
            {
                 response.SpecInfo.Label = "6130-0807"; 
            }

            // FT 的生產履歷，可能需要額外去撈取 FT Bin Data (Pass/Fail Qty)
            var histories = await _repository.GetStepHistoryAsync(lotId);
            foreach (var history in histories)
            {
                // ... 執行與先前相同的還原機台(EqpAttr)與配方(LotAttr)邏輯 ...
                
                // (擴充) 如果是 FT 測試站，去抓 Bin Data
                if (history.StepName.Contains("TEST"))
                {
                    // history.PassQty = await _repository.GetFtBinDataAsync(...);
                }
            }
            response.StepHistories = new List<RunCardStepHistory>(histories);
        }

        // ====================================================================
        // 私有方法：專門處理 WS 的邏輯 (對應原 VB6 的 WsExcelCell)
        // ====================================================================
        private async Task ProcessWsRunCardAsync(string lotId, RunCardResponse response)
        {
            // WS 的 Spec 來源表或欄位可能與 FT 不同，若不同可另建 Repository 方法
            response.SpecInfo = await _repository.GetLotSpecInfoAsync(response.IPN);

            // WS 的生產履歷，可能需要額外去撈取 Wafer 的 TDS (Test Data Sheet) 等級
            var histories = await _repository.GetStepHistoryAsync(lotId);
            foreach (var history in histories)
            {
                // ... 執行與先前相同的還原機台(EqpAttr)與配方(LotAttr)邏輯 ...

                // (擴充) 如果是 WS 測試站，去抓 Wafer 等級或良率
            }
            response.StepHistories = new List<RunCardStepHistory>(histories);
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

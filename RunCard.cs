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

    public class LotBasicInfo
    {
        public string IPN { get; set; }
        public string PlanId { get; set; }
        public string CurrentStepSeq { get; set; }
        public string ProdGroup { get; set; }
        public string LotOwner { get; set; }
        public int? Qty { get; set; }
        public DateTime? StartDate { get; set; }
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

        // 新增 WS 專用屬性
        public int WaferQty { get; set; }
        public string WaferNoList { get; set; } // 例如: "01, 02, 05, 07~15, 25"
        public string GrossDie { get; set; }
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
        // 取得該批號目前包含的所有晶圓號碼 (Wafer ID)
        Task<IEnumerable<string>> GetWaferListAsync(string lotId);
        
        // 取得 WS 測試站點的良率與測試數據
        Task<dynamic> GetWsTestDataAsync(string lotId, string stepName);
        Task<LotBasicInfo> GetLotBasicInfoAsync(string lotId);
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

        public async Task<IEnumerable<RunCardStepHistory>> GetStepHistoryAsync(string lotId, string planId, string planVersion)
        {
            // 將舊版 Oracle 的 (+) 語法轉換為標準的 LEFT JOIN 與 INNER JOIN
            // 變數 (如 gsCAT_TRSL_STEP_NO) 這裡先替換為預設的常見欄位名，請依實際欄位確認
            var sql = @"
                SELECT 
                    F.STEP_NO AS StepName, 
                    F.STEP_NAME AS Description, 
                    A.QUANTITYOUT AS QuantityOut, 
                    A.TRACKINTIME AS TrackInTime, 
                    A.TXNTIMESTAMP AS TrackOutTime, 
                    A.USERID AS UserOut, 
                    A.LASTTRACKOUT AS LastTrackOut, 
                    B.USERID AS UserIn, 
                    B.QUANTITYIN AS QuantityIn
                FROM FWTRACKOUT A
                -- 1. 關聯 FWTRACKIN (對應舊版 A.WIPID = B.WIPID(+) AND A.TRACKINTIME = B.TXNTIMESTAMP(+))
                LEFT JOIN FWTRACKIN B 
                    ON A.WIPID = B.WIPID AND A.TRACKINTIME = B.TXNTIMESTAMP
                -- 2. 串接核心歷史關聯表
                INNER JOIN FWWIPHISTORY C 
                    ON A.SYSID = C.WIPTXN
                INNER JOIN FWWIPHISTORY_N2M D 
                    ON C.SYSID = D.FROMID
                INNER JOIN FWWIPSTEPHISTORY E 
                    ON D.TOID = E.SYSID
                -- 3. 子查詢 G：串接 FlatPlan 取得節點 Handle
                INNER JOIN (
                    SELECT n.handle, n.StepSeq
                    FROM FwFlatNode n
                    INNER JOIN FwFlatPlan p ON p.SysId = n.PlanSysId
                    INNER JOIN FwStepVersion s ON s.SysId = n.NodeRef
                    WHERE p.revstate = 'Frozen' 
                      AND p.PlanName = :PlanId 
                      AND p.PlanVersion = :PlanVersion 
                      AND n.NodeType = 'S'
                ) G ON E.HANDLE = G.StepSeq
                -- 4. 關聯客製站點表 TBL_ROUTE_STEP_LIST (F)
                INNER JOIN TBL_ROUTE_STEP_LIST F 
                    ON G.handle = F.STEP_HANDLE AND F.ROUTE_ID = :PlanId
                -- 💡 注意：舊 SQL 似乎漏了過濾單一 LotId，這裡補上確保效能與正確性
                WHERE A.WIPID = :LotId
                ORDER BY F.STEP_SEQ ASC";
        
            return await _dbConnection.QueryAsync<RunCardStepHistory>(
                sql, 
                new 
                { 
                    LotId = lotId, 
                    PlanId = planId, 
                    PlanVersion = planVersion 
                }
            );
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

        public async Task<IEnumerable<string>> GetWaferListAsync(string lotId)
        {
            // 查詢批號下的晶圓號碼 (通常取後兩碼作為 Slot No)
            string sql = @"
                SELECT WAFER_ID 
                FROM TBL_LOT_WAFER 
                WHERE LOT_ID = :LotId AND DELETE_FLAG = 'N'
                ORDER BY WAFER_ID";
                
            return await _dbConnection.QueryAsync<string>(sql, new { LotId = lotId });
        }

        public async Task<dynamic> GetWsTestDataAsync(string lotId, string stepName)
        {
            // 撈取 CP (Circuit Probe) 測試的結果，例如 Good Die, Fail Die 等
            string sql = @"
                SELECT GOOD_DIE AS PassQty, FAIL_DIE AS FailQty, GROSS_DIE AS GrossDie, TDS_VERSION AS TdsVersion
                FROM TBL_WS_TEST_RESULT
                WHERE LOT_ID = :LotId AND STEP_NO = :StepName AND DELETE_FLAG = 'N'";
                
            return await _dbConnection.QueryFirstOrDefaultAsync(sql, new { LotId = lotId, StepName = stepName });
        }

        /// <summary>
        /// 取得批號的基本資訊 (IPN, PlanId, ProdGroup, Owner 等)
        /// </summary>
        public async Task<LotBasicInfo> GetLotBasicInfoAsync(string lotId)
        {
            // 💡 請依據貴公司實際的 Oracle 資料表與欄位名稱進行替換
            // 這裡以常見的 MES 主檔表 (例如 FWLOT 或 TBL_LOT) 為例示範
            string sql = @"
                SELECT 
                    IPN AS IPN,
                    PROCESS_PLAN_ID AS PlanId,       -- 途程代碼
                    CURRENT_STEP_SEQ AS CurrentStepSeq, -- 目前站點序號
                    PROD_GROUP AS ProdGroup,         -- 產品群組
                    LOT_OWNER AS LotOwner,           -- 批號擁有者
                    CURRENT_QTY AS Qty,              -- 目前數量 (Chip Qty)
                    CREATE_TIME AS StartDate         -- 批號建立或開始時間
                FROM 
                    FWLOT -- 替換為實際的 Table Name
                WHERE 
                    WIP_ID = :LotId"; -- 替換為實際的 LotId 欄位名稱

            // 使用 Dapper 的 QueryFirstOrDefaultAsync，查無資料會回傳 null
            return await _dbConnection.QueryFirstOrDefaultAsync<LotBasicInfo>(
                sql, 
                new { LotId = lotId }
            );
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
        
            // 1. 撈取共用 Lot 基本資訊 (取代原本寫死的假資料)
            var lotInfo = await _runCardRepository.GetLotBasicInfoAsync(request.LotId);
            if (lotInfo == null)
            {
                throw new Exception($"LotId [{request.LotId}] not found or basic info is missing.");
            }
        
            // 初始化 Response
            var response = new RunCardResponse
            {
                LotId = request.LotId,
                RunCardType = request.Type.ToUpper(),
                
                // 帶入實際查詢到的資料
                IPN = lotInfo.IPN,
                PlanId = lotInfo.PlanId,
                CurrentStepSeq = lotInfo.CurrentStepSeq,
                LotOwner = lotInfo.LotOwner,
                StartDate = lotInfo.StartDate?.ToString("yyyy/MM/dd HH:mm:ss"), // 視您的 DTO 型別調整
                ChipQty = lotInfo.Qty // 假設您的基本表有數量欄位
            };
        
            // 2. 判斷 LotType 與 Route (共用邏輯)
            response.LotType = await _getLotTypeRepository.GetLotTypeAsync(request.LotId);
            response.Route = await _getStepPathRepository.GetStepPathAsync(response.PlanId, response.CurrentStepSeq);
        
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
        
            // 4. 處理 Future Actions 
            // 💡 修正：將 "ProdGroup" 與 "Owner" 替換為實際變數 lotInfo.ProdGroup 與 response.LotOwner
            var rawActions = await _runCardRepository.GetFutureActionsAsync(
                request.LotId, 
                response.IPN, 
                lotInfo.ProdGroup, 
                response.LotOwner);
                
            var filteredActions = new List<RunCardFutureAction>();
        
            // 依據 LotType 過濾
            foreach (var act in rawActions)
            {
                if (act.ActionType == "Ipn" || act.ActionType == "ProdGroup")
                {
                    if (response.LotType == "Normal") filteredActions.Add(act);
                    else if (response.LotType == "Eng" && act.IncludeEngLot == "Y") filteredActions.Add(act);
                    else if (response.LotType == "Rework" && act.IncludeReworkLot == "Y") filteredActions.Add(act);
                }
                else
                {
                    filteredActions.Add(act);
                }
            }
            response.FutureActions = filteredActions;
        
            return response;
        }

        private async Task ProcessFtRunCardAsync(string lotId, RunCardResponse response)
        {
            // ... (取得 Spec 邏輯) ...
            var histories = await _repository.GetStepHistoryAsync(lotId);
            
            foreach (var history in histories)
            {
                if (history.TrackOutTime.HasValue && history.TrackInTime.HasValue)
                {
                    string curEqp = "TESTER_A"; // 假定由 GetLotAttr 取得
                    
                    if (history.Description.StartsWith("FT") || history.Description.StartsWith("TQAE"))
                    {
                        var binData = await _repository.GetFtBinDataAsync(lotId, history.StepName, curEqp);
                        if (binData != null)
                        {
                            history.PassQty = (int?)binData.PASSQTY;
                            history.FailQty = (int?)binData.FAILQTY;
                            history.Bin1 = binData.BIN1?.ToString();
                            // ... Bin2 ~ Bin6 ...
                            
                            if (history.QuantityIn > 0 && history.QuantityOut > 0)
                                history.Yield = Math.Round((double)history.QuantityOut / history.QuantityIn, 2);
                        }
                    }

                    if (history.QuantityIn > history.QuantityOut)
                    {
                        history.ScrapComment = await _repository.GetScrapCommentAsync(lotId, history.TrackInTime.Value, history.TrackOutTime.Value);
                    }
                }
            }
            response.StepHistories = new List<RunCardStepHistory>(histories);
        }

        /// <summary>
        /// 核心：將組裝好的資料產出為 Excel 二進制陣列
        /// </summary>
        public byte[] GenerateExcelReport(RunCardResponse data)
        {
            using (var workbook = new XLWorkbook())
            {
                var ws = workbook.Worksheets.Add(data.RunCardType + "_RunCard");
                
                // --- Title 區塊 (參照 CSV 位置) ---
                ws.Cell(1, 11).Value = "Macronix Final Test Run Card";
                ws.Cell(1, 17).Value = "Date: " + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
                
                // --- Basic Information ---
                ws.Cell(2, 1).Value = "Basic information :";
                
                ws.Cell(3, 4).Value = "Ipn :";         ws.Cell(3, 5).Value = data.IPN;
                ws.Cell(3, 16).Value = "Route :";      ws.Cell(3, 17).Value = data.Route;
                
                ws.Cell(4, 4).Value = "Lot id :";      ws.Cell(4, 5).Value = data.LotId;
                ws.Cell(5, 4).Value = "Qty :";         ws.Cell(5, 5).Value = data.ChipQty;
                
                // ... 繼續對應座標寫入 Spec ...
                if (data.SpecInfo != null)
                {
                    ws.Cell(7, 4).Value = "EPN :";          ws.Cell(7, 5).Value = data.SpecInfo.EPN;
                    ws.Cell(11, 4).Value = "Label spec :";  ws.Cell(11, 5).Value = data.SpecInfo.Label;
                }

                // --- Process Record (Step History) ---
                int currentRow = 20; // 假設歷史紀錄從第 20 行開始
                ws.Cell(18, 1).Value = "Process Record :";
                
                foreach (var hist in data.StepHistories)
                {
                    ws.Cell(currentRow, 4).Value = "Step name :";
                    ws.Cell(currentRow, 5).Value = hist.Description;
                    ws.Cell(currentRow, 16).Value = "Step id :";
                    ws.Cell(currentRow, 17).Value = hist.StepName;
                    
                    ws.Cell(currentRow + 1, 4).Value = "Track In :";
                    ws.Cell(currentRow + 1, 5).Value = hist.TrackInTime?.ToString("yyyy/MM/dd HH:mm:ss");
                    ws.Cell(currentRow + 1, 16).Value = "Track Out :";
                    ws.Cell(currentRow + 1, 17).Value = hist.TrackOutTime?.ToString("yyyy/MM/dd HH:mm:ss");

                    if (hist.PassQty.HasValue)
                    {
                        ws.Cell(currentRow + 3, 16).Value = "Pass Qty :";
                        ws.Cell(currentRow + 3, 17).Value = hist.PassQty;
                        ws.Cell(currentRow + 4, 16).Value = "Fail Qty :";
                        ws.Cell(currentRow + 4, 17).Value = hist.FailQty;
                    }
                    
                    currentRow += 13; // 依據原本 VB 每個 Step 佔據 13 行空間
                }

                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    return stream.ToArray();
                }
            }
        }

        // ====================================================================
        // 私有方法：專門處理 WS 的邏輯 (對應原 VB6 的 WsExcelCell)
        // ====================================================================
        private async Task ProcessWsRunCardAsync(string lotId, RunCardResponse response)
        {
            // 1. 取得 Spec 資訊 (與 FT 相同或取不同的欄位)
            response.SpecInfo = await _repository.GetLotSpecInfoAsync(response.IPN);

            // 2. 🌟 WS 專屬：取得晶圓清單與數量
            var wafers = await _repository.GetWaferListAsync(lotId);
            if (wafers != null && wafers.Any())
            {
                response.WaferQty = wafers.Count();
                // 將 Wafer ID 陣列轉為逗號分隔字串，或撰寫縮寫邏輯 (如: 01~15)
                response.WaferNoList = string.Join(", ", wafers.Select(w => w.Substring(w.Length - 2))); 
            }

            // 3. 取得生產履歷，並還原 WS 測試屬性
            var histories = await _repository.GetStepHistoryAsync(lotId);
            
            foreach (var history in histories)
            {
                if (history.TrackOutTime.HasValue)
                {
                    // 還原機台與 Recipe (與 FT 共用邏輯)
                    string curEqp = await _repository.GetLotAttrAsync(lotId, "CurEqpId", history.TrackOutTime.Value);
                    string pgName = await _repository.GetLotAttrAsync(lotId, "PgName", history.TrackOutTime.Value);
                    history.Recipe = pgName;
                    history.Equipment = curEqp;

                    // 🌟 WS 專屬：若站點名稱包含 CP (Circuit Probe) 或 WS，則撈取測試良率
                    if (history.Description.Contains("CP") || history.Description.Contains("WS"))
                    {
                        var wsData = await _repository.GetWsTestDataAsync(lotId, history.StepName);
                        if (wsData != null)
                        {
                            history.PassQty = (int?)wsData.PASSQTY;
                            history.FailQty = (int?)wsData.FAILQTY;
                            
                            // 記錄 TDS 版本於註解欄位 (或其他您定義的 DTO 欄位)
                            history.ScrapComment = $"TDS: {wsData.TDSVERSION}"; 
                        }
                    }
                }
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
        /// <summary>
        /// 下載 Excel 實體檔案
        /// </summary>
        [HttpPost, Route("download-excel"), AuthorizeToken]
        public async Task<IHttpActionResult> DownloadExcel([FromBody] PrintRunCardRequest request)
        {
            try
            {
                // 1. 取得完整組裝資料
                var data = await _service.GetRunCardDataAsync(request);
                
                // 2. 產出 Excel 二進制流
                byte[] excelBytes = ((RunCardService)_service).GenerateExcelReport(data);
                
                // 3. 設定檔名並回傳實體檔案
                string fileName = $"{request.Type}_{data.LotId}_{DateTime.Now:yyyyMMddHHmm}.xlsx";
                
                var result = new System.Net.Http.HttpResponseMessage(System.Net.HttpStatusCode.OK)
                {
                    Content = new System.Net.Http.ByteArrayContent(excelBytes)
                };
                result.Content.Headers.ContentType = new System.Net.Http.Headers.MediaTypeHeaderValue("application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
                result.Content.Headers.ContentDisposition = new System.Net.Http.Headers.ContentDispositionHeaderValue("attachment")
                {
                    FileName = fileName
                };
                
                return ResponseMessage(result);
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }
    }
}

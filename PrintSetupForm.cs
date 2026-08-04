using System.Collections.Generic;

namespace MES.Net.Shared.DTOs.Print
{
    public class PrintSetupFormQueryRequest
    {
        public string LotId { get; set; }
    }

    public class PrintSetupFormQueryResponse
    {
        public string LotId { get; set; }
        public string Stage { get; set; }
        public string LotStatus { get; set; }
        public string IPN { get; set; }
        public string LotOwner { get; set; }

        // 下拉選單資料來源
        public List<SelectItem> StepList { get; set; } = new List<SelectItem>();
        public string SelectedStep { get; set; } // 預設選取的 Step

        public List<SelectItem> TesterIdList { get; set; } = new List<SelectItem>();
        public string SelectedTesterId { get; set; } // 預設選取的 TesterId

        public List<SelectItem> SubSystemList { get; set; } = new List<SelectItem>();
    }

    public class SelectItem
    {
        public string Text { get; set; }
        public string Value { get; set; }
    }

    // 內部傳遞用的 Lot 基本資訊 Model
    public class SetupLotBasicInfo
    {
        public string LotId { get; set; }
        public string Status { get; set; }
        public string Ipn { get; set; }
        public string Owner { get; set; }
        public string CurrentStepId { get; set; }
        public string CurrentStepName { get; set; }
        public string CurEqpId { get; set; }
        public string SplitAssignEqId { get; set; }
    }

    public class PrintSetupSubsystemRequest
    {
        public string LotId { get; set; }
        public string Stage { get; set; }
        public string TesterId { get; set; }
        public string StepNo { get; set; }
        public string StepName { get; set; }
        public string SubSystemValue { get; set; } // 例如: "SYS1,SITE1" 或 "SYS1"
    }

    public class PrintSetupSubsystemResponse
    {
        public string PgName { get; set; }
        public string PgId { get; set; }
        public string PgMode { get; set; }
        public string Temp { get; set; }
        public string WsDeviceFile { get; set; }
        public string AccType { get; set; }
        public string SpecifyEq { get; set; }
        public string SpecifyEqId { get; set; }
        public string StopTicNo { get; set; }
    }

    // 用來承接三種資料表共同欄位的內部 Model
    public class RecipeSpecData
    {
        public string PgName { get; set; }
        public string PgId { get; set; }
        public string PgMode { get; set; }
        public string Temperature { get; set; }
        public string WsDeviceFile { get; set; }
        public string ProbeCardType { get; set; }
        public string LoadboardType { get; set; }
        public string ContactboardType { get; set; }
        public string BurnInBoard { get; set; }
        public string SpecifyEq { get; set; }
        public string EqId { get; set; }
        public string StopTicNo { get; set; }
        
        // 供 GetPGM 替換用的屬性
        public string RefStepName01 { get; set; }
        public string RefPgName01 { get; set; }
        public string ReplacePgName01 { get; set; }
        // ... (視需求擴充 02, 03)
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
    public interface IPrintSetupFormRepository
    {
        Task<dynamic> GetLotAttributeAsync(string lotId);
        Task<dynamic> GetLotInfoAsync(string lotId);
        Task<string> GetFollowProductAsync(string lotId, string erunTicNo, string stage);
        Task<string> GetIpnBodySizeAsync(string ipn);
        Task<IEnumerable<SelectItem>> GetFtRouteStepsAsync(string path);
        Task<string> GetEqType2Async(string eqId);
        Task<IEnumerable<string>> GetSubSystemsAsync(string sql, object param);
        Task<RecipeSpecData> GetErunRecipeAsync(string erunTicNo, string stepNo, string eqType2, string subSystem);
        Task<RecipeSpecData> GetLotStepEqSpecAsync(string tecnLotId, string stepNo, string eqType2, string subSystem, string path, string maxSite);
        Task<RecipeSpecData> GetProdStepEqSpecAsync(string prodGroup, string stepNo, string eqType2, string subSystem, string path, string maxSite);
    }

    public class PrintSetupFormRepository : IPrintSetupFormRepository
    {
        private readonly IDbConnection _dbConnection;
        private static readonly string _connString = ConfigurationManager.ConnectionStrings["MESDB"]?.ConnectionString;

        public PrintSetupFormRepository()
        {
            _dbConnection = new OracleConnection(_connString);
        }

        public async Task<dynamic> GetLotAttributeAsync(string lotId)
        {
            string sql = "SELECT STAGE, ROUTE as PATH FROM fwadmin.tbl_lot_attribute WHERE lotid = :LotId";
            return await _dbConnection.QueryFirstOrDefaultAsync(sql, new { LotId = lotId });
        }

        public async Task<dynamic> GetLotInfoAsync(string lotId)
        {
            string sql = @"
                SELECT TECN_LOT_ID as TecnLotId, ERUNTICNO as ErunTicNo, 
                       ASSIGNPROBECARD as AssignProbeCard, ASSIGNLOADBOARD as AssignLoadBoard
                FROM TBL_LOT_INFO WHERE LOT_ID = :LotId";
            return await _dbConnection.QueryFirstOrDefaultAsync(sql, new { LotId = lotId });
        }

        public async Task<string> GetFollowProductAsync(string lotId, string erunTicNo, string stage)
        {
            string sql = @"
                SELECT FOLLOW_PRODUCT 
                FROM TBL_ERUN_REQ 
                WHERE LOT_ID = :LotId AND TICKET_NO = :TicNo AND STAGE = :Stage AND DELETE_FLAG = 'N'";
            return await _dbConnection.QueryFirstOrDefaultAsync<string>(sql, new { LotId = lotId, TicNo = erunTicNo, Stage = stage });
        }

        public async Task<string> GetIpnBodySizeAsync(string ipn)
        {
            string sql = "SELECT BODY_SIZE FROM TBL_IPN_MASTER WHERE IPN = :Ipn";
            return await _dbConnection.QueryFirstOrDefaultAsync<string>(sql, new { Ipn = ipn });
        }

        public async Task<IEnumerable<SelectItem>> GetFtRouteStepsAsync(string path)
        {
            string sql = @"
                SELECT DISTINCT STEP_NO as Value, (STEP_NO || ' ' || STEP_NAME) as Text 
                FROM TBL_ROUTE_STEP_LIST 
                WHERE PATH = :Path 
                  AND (STEP_NAME LIKE 'FT_%' OR STEP_NAME LIKE 'FT__%' OR STEP_NAME LIKE 'TQAE%')";
            return await _dbConnection.QueryAsync<SelectItem>(sql, new { Path = path });
        }

        public async Task<string> GetEqType2Async(string eqId)
        {
            string sql = "SELECT DISTINCT eqtype2 FROM view_b2b_fweqarea WHERE eqid = :EqId";
            return await _dbConnection.QueryFirstOrDefaultAsync<string>(sql, new { EqId = eqId });
        }

        public async Task<IEnumerable<string>> GetSubSystemsAsync(string sql, object param)
        {
            return await _dbConnection.QueryAsync<string>(sql, param);
        }

        public async Task<RecipeSpecData> GetErunRecipeAsync(string erunTicNo, string stepNo, string eqType2, string subSystem)
        {
            // 對應 VB 中對 TBL_ERUN_RECIPE 的查詢
            string sql = @"
                SELECT PGNAME, PGID, PGMODE, TEMPERATURE, PROBECARDTYPE, LOADBOARDTYPE, 
                       CONTACTBOARDTYPE, WSDEVICEFILE
                FROM TBL_ERUN_RECIPE
                WHERE DOCNO = :DocNo AND STEPNO = :StepNo AND EQTYPE2 = :EqType2 
                  AND DELETEFLAG = 'N' " + (string.IsNullOrEmpty(subSystem) ? "AND SUBSYSTEM IS NULL" : "AND SUBSYSTEM = :SubSystem");

            return await _dbConnection.QueryFirstOrDefaultAsync<RecipeSpecData>(sql, new { DocNo = erunTicNo, StepNo = stepNo, EqType2 = eqType2, SubSystem = subSystem });
        }

        public async Task<RecipeSpecData> GetLotStepEqSpecAsync(string tecnLotId, string stepNo, string eqType2, string subSystem, string path, string maxSite)
        {
            // 對應 VB 中對 TBL_LOT_STEP_EQ_SPEC 的查詢 (TECN)
            string sql = @"
                SELECT PGNAME, PGID, PGMODE, TEMPERATURE, PROBECARDTYPE, LOADBOARDTYPE, 
                       CONTACTBOARDTYPE, WSDEVICEFILE, SPECIFYEQ, EQID, BURNINBOARD,
                       REFSTEPNAME01, REFPGNAME01, REPLACEPGNAME01
                FROM TBL_LOT_STEP_EQ_SPEC
                WHERE TECNLOTID LIKE :TecnLotId AND STEPNO = :StepNo AND EQTYPE2 = :EqType2 
                  AND SUBSYSTEM = :SubSystem AND PATH = :Path AND NVL(MAXSITE, ' ') = NVL(:MaxSite, ' ')
                  AND DELETEFLAG = 'N'";
            return await _dbConnection.QueryFirstOrDefaultAsync<RecipeSpecData>(sql, new { TecnLotId = tecnLotId, StepNo = stepNo, EqType2 = eqType2, SubSystem = subSystem, Path = path, MaxSite = maxSite });
        }

        public async Task<RecipeSpecData> GetProdStepEqSpecAsync(string prodGroup, string stepNo, string eqType2, string subSystem, string path, string maxSite)
        {
            // 對應 VB 中對 TBL_PROD_STEP_EQ_SPEC 的查詢 (Normal)
            string sql = @"
                SELECT PG_NAME as PgName, PG_ID as PgId, PG_MODE as PgMode, TEMPERATURE, 
                       PROBECARD_TYPE as ProbeCardType, LOADBOARD_TYPE as LoadboardType, 
                       CONTACTBOARD_TYPE as ContactboardType, WS_DEVICE_FILE as WsDeviceFile, 
                       SPECIFYEQ, EQID, STOPTICNO, BURN_IN_BOARD as BurnInBoard,
                       REF_STEP_NAME_01 as RefStepName01, REF_PG_NAME_01 as RefPgName01, REPLACE_PG_NAME_01 as ReplacePgName01
                FROM TBL_PROD_STEP_EQ_SPEC
                WHERE PRODGROUP = :ProdGroup AND STEPNO = :StepNo AND EQTYPE2 = :EqType2 
                  AND SUBSYSTEM = :SubSystem AND PATH = :Path AND NVL(MAX_SITE, ' ') = NVL(:MaxSite, ' ')
                  AND DOCSTATUS = 'Active'";
            return await _dbConnection.QueryFirstOrDefaultAsync<RecipeSpecData>(sql, new { ProdGroup = prodGroup, StepNo = stepNo, EqType2 = eqType2, SubSystem = subSystem, Path = path, MaxSite = maxSite });
        }
    }
}

using System;
using System.Linq;
using System.Threading.Tasks;
using MES.Net.Shared.DTOs.Print;
using MES.Net.Infrastructure.Repository.Print;
using MES.Net.Infrastructure.ExternalServices; // 引入 WipServiceWrapper
using MES.Net.Shared.Constants.FwAttributes;   // 假設貴司常數定義於此

namespace MES.Net.Application.Services.Print
{
    public interface IPrintSetupFormService
    {
        Task<PrintSetupFormQueryResponse> QuerySetupFormDataAsync(string lotId);
        Task<PrintSetupSubsystemResponse> ProcessSubsystemChangeAsync(PrintSetupSubsystemRequest request);
    }

    public class PrintSetupFormService : IPrintSetupFormService
    {
        private readonly IPrintSetupFormRepository _repo;

        public PrintSetupFormService(IPrintSetupFormRepository repo)
        {
            _repo = repo;
        }

        public async Task<PrintSetupFormQueryResponse> QuerySetupFormDataAsync(string lotId)
        {
            lotId = lotId?.ToUpper()?.Trim();
            if (string.IsNullOrEmpty(lotId)) throw new ArgumentException("LotId cannot be empty.");

            // 1. 使用貴司共用模組取得 FwLot 物件
            var olot = WipServiceWrapper.Instance.LotById(lotId);
            if (olot == null) 
            {
                throw new Exception($"此 Lot ({lotId}) 不存在 !! (This Lot was not exist.)");
            }

            // 取得當前站點資訊
            string stepId = olot.CurrentStep.Steps[0].Id;
            string stepDesc = olot.CurrentStep.Steps[0].Description ?? "";

            // 檢查是否在測試站別 (SORT, FT, TQAE, PROGRAMMER)
            var stepDescUpper = stepDesc.ToUpper();
            if (!(stepDescUpper.StartsWith("SORT") || stepDescUpper.StartsWith("FT") || 
                  stepDescUpper.StartsWith("TQAE") || stepDescUpper.StartsWith("PROGRAMMER")))
            {
                throw new Exception("此Lot未在測試站別, 無法列印 !! (This Lot was not in Testing Step.)");
            }

            // 取得機台號碼 (優先取 CurEqpId，若無則取 SplitAssignEqId)
            // (註: 這裡的 "CurEqpId" 等字串請依貴司常數定義自行替換)
            string curEqpId = olot.CustomAttributes("CurEqpId"); 
            string splitAssignEqId = olot.CustomAttributes("SplitAssignEqId");
            string targetEqId = !string.IsNullOrEmpty(curEqpId) ? curEqpId : splitAssignEqId;

            if (string.IsNullOrEmpty(targetEqId))
            {
                throw new Exception("此Lot之測試機台不存在 !! (The Equipment of this Lot was not exited.)");
            }

            // 2. 初始化 Response 並塞入 Lot 基本資料
            var response = new PrintSetupFormQueryResponse
            {
                LotId = lotId,
                LotStatus = olot.CustomAttributes("Status"), // 或對應貴司 Status 常數
                IPN = olot.CustomAttributes(LotCustomAttributes.Ipn),
                LotOwner = olot.CustomAttributes(LotCustomAttributes.LotOwner)
            };

            // 3. 取得 Lot Attribute (Stage, Path) 
            // (註: 依據 VB 寫法，這兩欄來自 tbl_lot_attribute 表)
            var attr = await _repo.GetLotAttributeAsync(lotId);
            response.Stage = attr?.STAGE ?? "";
            string path = attr?.PATH ?? "";

            // 4. 設定 TesterId 選單
            response.TesterIdList.Add(new SelectItem { Text = targetEqId, Value = targetEqId });
            response.SelectedTesterId = targetEqId;
            
            string eqType2 = await _repo.GetEqType2Async(targetEqId);

            // 5. 設定 Step 選單
            if (response.Stage != "FT")
            {
                response.StepList.Add(new SelectItem { Value = stepId, Text = $"{stepId} {stepDesc}" });
                response.SelectedStep = stepId;
            }
            else
            {
                var ftSteps = await _repo.GetFtRouteStepsAsync(path);
                response.StepList = ftSteps.ToList();
                response.SelectedStep = stepId; // 預設選取目前站別
            }

            // 6. 取得 Lot Info 相關資訊 (Tbl_Lot_Info / Tbl_Erun_Req)
            var lotInfo = await _repo.GetLotInfoAsync(lotId);
            string erunTicNo = lotInfo?.ERUNTICNO ?? "";
            string tecnLotId = lotInfo?.TECNLOTID ?? "";
            string assignProbeCard = lotInfo?.ASSIGNPROBECARD ?? "";
            string assignLoadBoard = lotInfo?.ASSIGNLOADBOARD ?? "";

            string followProd = "";
            if (!string.IsNullOrEmpty(erunTicNo))
            {
                followProd = await _repo.GetFollowProductAsync(lotId, erunTicNo, response.Stage) ?? "";
            }

            string bodySize = await _repo.GetIpnBodySizeAsync(response.IPN) ?? "";

            // 7. 組合查詢 SubSystem 的 SQL
            string subSystemSql = BuildSubSystemSql(response.Stage, erunTicNo, followProd, tecnLotId, path, assignProbeCard, assignLoadBoard);
            
            var subSystems = await _repo.GetSubSystemsAsync(subSystemSql, new
            {
                ErunTicNo = erunTicNo,
                StepNo = stepId,
                EqType2 = eqType2,
                Path = path,
                TecnLotId = tecnLotId,
                BodySize = bodySize
                // ProdGroup 等參數視實際 SQL 組裝帶入
            });

            // 整理 SubSystem 格式
            foreach (var sys in subSystems)
            {
                response.SubSystemList.Add(new SelectItem { Text = sys, Value = sys });
            }

            return response;
        }

        private string BuildSubSystemSql(string stage, string erunTicNo, string followProd, string tecnLotId, string path, string assignProbeCard, string assignLoadBoard)
        {
            // (維持前一版的 SQL 組裝邏輯)
            string sql = "";
            if (!string.IsNullOrEmpty(erunTicNo) && followProd == "N")
            {
                sql = "SELECT (SUBSYSTEM || ',' || '') FROM TBL_ERUN_RECIPE WHERE DOCNO = :ErunTicNo AND STEPNO = :StepNo AND EQTYPE2 = :EqType2 AND DELETEFLAG = 'N' AND PATH = :Path";
            }
            else if (!string.IsNullOrEmpty(tecnLotId))
            {
                sql = "SELECT (SUBSYSTEM || ',' || MAXSITE) FROM TBL_LOT_STEP_EQ_SPEC WHERE :TecnLotId LIKE TECNLOTID AND STEPNO = :StepNo AND EQTYPE2 = :EqType2 AND DELETEFLAG = 'N' AND PATH = :Path";
            }
            else
            {
                sql = "SELECT (SUBSYSTEM || ',' || MAXSITE) FROM TBL_PROD_STEP_EQ_SPEC WHERE STEPNO = :StepNo AND EQTYPE2 = :EqType2 AND DOCSTATUS = 'Active' AND PATH = :Path";
            }
            return sql;
        }

        public async Task<PrintSetupSubsystemResponse> ProcessSubsystemChangeAsync(PrintSetupSubsystemRequest request)
        {
            var response = new PrintSetupSubsystemResponse();

            if (string.IsNullOrEmpty(request.SubSystemValue)) return response;

            // 1. 解析 SubSystem 與 MaxSite
            string subSystemType = request.SubSystemValue;
            string maxSite = string.Empty;
            if (request.SubSystemValue.Contains(","))
            {
                var parts = request.SubSystemValue.Split(',');
                subSystemType = parts[0];
                maxSite = parts[1];
            }

            // 2. 取得 Lot 資訊與基本屬性
            var olot = WipServiceWrapper.Instance.LotById(request.LotId);
            string ipn = olot.CustomAttributes(LotCustomAttributes.Ipn);
            string prodGroup = olot.CustomAttributes(LotCustomAttributes.ProdGroup);
            
            var attr = await _repo.GetLotAttributeAsync(request.LotId);
            string path = attr?.PATH ?? "";

            string eqType2 = await _repo.GetEqType2Async(request.TesterId);

            // 3. 取得 Lot Info
            var lotInfo = await _repo.GetLotInfoAsync(request.LotId);
            string erunTicNo = lotInfo?.ERUNTICNO ?? "";
            string tecnLotId = lotInfo?.TECNLOTID ?? "";
            string followProd = string.IsNullOrEmpty(erunTicNo) ? "" : await _repo.GetFollowProductAsync(request.LotId, erunTicNo, request.Stage);

            RecipeSpecData specData = null;

            // 4. 判斷情境並撈取對應的 Spec 資料
            if (!string.IsNullOrEmpty(erunTicNo) && followProd == "N")
            {
                // 情境 1: Erun 重工且不 follow product
                specData = await _repo.GetErunRecipeAsync(erunTicNo, request.StepNo, eqType2, subSystemType);
            }
            else if (!string.IsNullOrEmpty(tecnLotId)) // 或 bTECNLotSpec = True
            {
                // 情境 2: TECN Lot
                specData = await _repo.GetLotStepEqSpecAsync(tecnLotId, request.StepNo, eqType2, subSystemType, path, maxSite);
                
                // 這裡會呼叫貴司既有的 GetPGM, GetFTAccByPgm 等共用函數，這裡以註解示意
                // specData.PgName = CustomMesHelper.GetPGM(olot, specData.PgName, ...);
            }
            else
            {
                // 情境 3: 一般產品
                specData = await _repo.GetProdStepEqSpecAsync(prodGroup, request.StepNo, eqType2, subSystemType, path, maxSite);
                
                // 這裡會呼叫貴司既有的 GetSwapPGName 等共用函數
                // string swapPgName = CustomMesHelper.GetSwapPGName(request.LotId, prodGroup, ...);
                // if (!string.IsNullOrEmpty(swapPgName) && swapPgName != "X") specData.PgName = swapPgName;
            }

            // 5. 將結果填入 Response
            if (specData != null)
            {
                response.PgName = specData.PgName;
                response.PgId = specData.PgId;
                response.PgMode = specData.PgMode;
                response.Temp = specData.Temperature;
                response.WsDeviceFile = specData.WsDeviceFile;
                response.SpecifyEq = specData.SpecifyEq;
                response.SpecifyEqId = specData.EqId;
                response.StopTicNo = specData.StopTicNo;

                // 判斷配件 AccType (優先順序: ProbeCard -> LoadBoard -> ContactBoard)
                if (!string.IsNullOrEmpty(specData.ProbeCardType))
                    response.AccType = specData.ProbeCardType;
                else if (!string.IsNullOrEmpty(specData.LoadboardType))
                    response.AccType = specData.LoadboardType;
                else if (!string.IsNullOrEmpty(specData.ContactboardType))
                    response.AccType = specData.ContactboardType;
                // 若都為空，VB中會去呼叫 GetFTAccByPgm 取得，可在此整合
            }

            // (可選) 呼叫 modTecn.GetTecnPgmRecipeAttr 進行最後 TECN 屬性的覆蓋
            // CustomMesHelper.GetTecnPgmRecipeAttr(..., out getPgm, out getPgId, out getTemp);
            // if (!string.IsNullOrEmpty(getPgm)) response.PgName = getPgm;
            // ...

            return response;
        }
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
    [RoutePrefix("api/print/print-setup-form")]
    public class PrintSetupFormController : ApiController
    {
        private readonly IPrintSetupFormService _service;

        public PrintSetupFormController()
        {
        }

        public PrintSetupFormController(IPrintSetupFormService service)
        {
            _service = service;
        }

        /// <summary>
        /// 取得 E-Test (PrintSetupForm) 查詢後的初始綁定資料
        /// 對應 VB 端的 ExecQuery() 邏輯
        /// </summary>
        [HttpPost, Route("query"), AuthorizeToken]
        public async Task<IHttpActionResult> QuerySetupFormData([FromBody] PrintSetupFormQueryRequest request)
        {
            try
            {
                if (request == null || string.IsNullOrWhiteSpace(request.LotId))
                {
                    return Ok(new { Success = false, Message = "請輸入LotID! (Please key in LotID!)" });
                }

                var data = await _service.QuerySetupFormDataAsync(request.LotId);
                return Ok(new { Success = true, Data = data });
            }
            catch (Exception ex)
            {
                // 將 Exception Message (如：Lot不存在) 原封不動傳給前端 VMessageBox 顯示
                return Ok(new { Success = false, Message = ex.Message });
            }
        }
        /// <summary>
        /// 處理 SubSystem 下拉選單變更，回傳配方明細資料
        /// 對應 VB 端的 cboSubSystem_Click()
        /// </summary>
        [HttpPost, Route("subsystem-change"), AuthorizeToken]
        public async Task<IHttpActionResult> SubsystemChange([FromBody] PrintSetupSubsystemRequest request)
        {
            try
            {
                if (request == null || string.IsNullOrWhiteSpace(request.LotId))
                {
                    return Ok(new { Success = false, Message = "無效的請求參數!" });
                }

                var data = await _service.ProcessSubsystemChangeAsync(request);
                return Ok(new { Success = true, Data = data });
            }
            catch (Exception ex)
            {
                return Ok(new { Success = false, Message = ex.Message });
            }
        }
    }
}

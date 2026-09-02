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
        // === GetTecnPgmRecipeAttr 所需的查詢 ===
        Task<dynamic> GetIpnMasterForTecnAsync(string ipn);
        Task<IEnumerable<dynamic>> GetTecnPgmRecordsAsync(string eqType2, string testMode, int level, string description, string pgName, string tecnNo);
        Task<string> GetTecnControlActionAsync(string tecnNo, string lotId);

        // === GetSwapPGName 所需的查詢 ===
        Task<IEnumerable<string>> GetProdStepSpecStepsAsync(string prodGroup);
        Task<IEnumerable<dynamic>> GetLotPgmRecSequenceAsync(string lotId);
        Task<IEnumerable<dynamic>> GetWsPgmReplaceAsync(string prodGroup);

        Task<string> GetLotStepEqSpecContinueTestAsync(string tecnLotId, string stepNo, string path, string eqType2, string subSystem);
        Task<EqAccessoriesDto> GetEqInfoAccessoriesAsync(string eqId);
        Task<dynamic> GetErunReqContinueTestAsync(string lotId, string erunTicNo, string stage);
        Task<IEnumerable<ProductStopTestRecordDto>> GetProductStopTestRecordsAsync(string eqType2, string prodGroup, string ipn, string pgId, string pgName, string stepName, string concatAcc, string lotId, string curAccNo, string wsDeviceFile);
        Task<IEnumerable<string>> GetNonProdStopTestRecordsAsync(string eqType2, string curEqId, string curAccName, string packageType, string pinCount, string carrierType, string bodySize, int temperature, string curAccNo, string formFactorName, string moduleOption);

        Task<string> GetShiftCodeAsync(string empNo);
        Task<string> GetSoakTimeAsync(string eqType2, string pinCount, string packageCode);
        Task<string> GetSthandFileAsync(string loadBoard);
        Task<dynamic> GetEqInfoAccessoriesFullAsync(string eqId); 
        Task<dynamic> GetTesterFviAccAsync(string pinCount, string pkgCode);
        Task<string> GetSpecCommentsAsync(bool isTecn, string id, string stepNo, string path);
        Task<string> GetAccMatchPlateAsync(string ipn, string kit, string loadBoard, string eqType2);
        Task<IEnumerable<RecipeItemDto>> GetRecipeSpecNamesAsync(string eqId, string lotId, string ipn, string packageName, string pinCount, string bodySize);

        Task<TecnPgmRecipeAttrDto> GetTecnPgmRecipeAttrAsync(string lotId, string stepNo, string eqType2, string subSystem);
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
                SELECT TECN_LOT_ID as TecnLotId, 
                       ERUNTICNO as ErunTicNo, 
                       SAPRWNO as SapRwNo, 
                       ASSIGNPROBECARD as AssignProbeCard, 
                       ASSIGNLOADBOARD as AssignLoadBoard
                FROM TBL_LOT_INFO 
                WHERE LOT_ID = :LotId";
            
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

        public async Task<IEnumerable<SubSystemResultDto>> GetSubSystemsAsync(string sql, object param)
        {
            // 直接讓 Dapper 幫我們把欄位對應到 DTO 的屬性上
            return await _dbConnection.QueryAsync<SubSystemResultDto>(sql, param);
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

{
            // TBL_LOT_STEP_EQ_SPEC (TECN 完整版)
            string sql = @"
                SELECT PGNAME as PgName, PGID as PgId, PGMODE as PgMode, TEMPERATURE as Temperature, 
                       PROBECARDTYPE as ProbeCardType, LOADBOARDTYPE as LoadboardType, 
                       CONTACTBOARDTYPE as ContactboardType, WSDEVICEFILE as WsDeviceFile, 
                       SPECIFYEQ as SpecifyEq, EQID as EqId, BURNINBOARD as BurnInBoard,
                       REFSTEPNAME01 as RefStepName01, REFPGNAME01 as RefPgName01, REPLACEPGNAME01 as ReplacePgName01,
                       TECNNO as TecnNo, STEPCOMMENT as StepComments,
                       CABLETYPE as CableType, KITTYPE as KitType, NEEDJUMPER as NeedJumper, JUMPERPINNO as JumperPinNo
                FROM TBL_LOT_STEP_EQ_SPEC
                WHERE TECNLOTID LIKE :TecnLotId AND STEPNO = :StepNo AND EQTYPE2 = :EqType2 
                  AND NVL(TRIM(SUBSYSTEM), ' ') = NVL(TRIM(:SubSystem), ' ') AND PATH = :Path AND NVL(MAXSITE, ' ') = NVL(:MaxSite, ' ')
                  AND DELETEFLAG = 'N'";
                  
            return await _dbConnection.QueryFirstOrDefaultAsync<RecipeSpecData>(sql, new { TecnLotId = tecnLotId, StepNo = stepNo, EqType2 = eqType2, SubSystem = subSystem, Path = path, MaxSite = maxSite });
        }

        public async Task<RecipeSpecData> GetProdStepEqSpecAsync(string prodGroup, string stepNo, string eqType2, string subSystem, string path, string maxSite)
        {
            // TBL_PROD_STEP_EQ_SPEC (Normal 完整版)
            string sql = @"
                SELECT PG_NAME as PgName, PG_ID as PgId, PG_MODE as PgMode, TEMPERATURE as Temperature, 
                       PROBECARD_TYPE as ProbeCardType, LOADBOARD_TYPE as LoadboardType, 
                       CONTACTBOARD_TYPE as ContactboardType, WS_DEVICE_FILE as WsDeviceFile, 
                       SPECIFYEQ as SpecifyEq, EQID as EqId, STOPTICNO as StopTicNo, BURN_IN_BOARD as BurnInBoard,
                       REF_STEP_NAME_01 as RefStepName01, REF_PG_NAME_01 as RefPgName01, REPLACE_PG_NAME_01 as ReplacePgName01,
                       STEPCOMMENT as StepComments, DOC_TYPE as DocType, DOC_NO as DocNo,
                       CABLE_TYPE as CableType, KIT_TYPE as KitType, NEEDJUMPER as NeedJumper, JUMPERPINNO as JumperPinNo
                FROM TBL_PROD_STEP_EQ_SPEC
                WHERE PRODGROUP = :ProdGroup AND STEPNO = :StepNo AND EQTYPE2 = :EqType2 
                  AND NVL(TRIM(SUBSYSTEM), ' ') = NVL(TRIM(:SubSystem), ' ') AND PATH = :Path AND NVL(MAX_SITE, ' ') = NVL(:MaxSite, ' ')
                  AND DOCSTATUS = 'Active'";
                  
            return await _dbConnection.QueryFirstOrDefaultAsync<RecipeSpecData>(sql, new { ProdGroup = prodGroup, StepNo = stepNo, EqType2 = eqType2, SubSystem = subSystem, Path = path, MaxSite = maxSite });
        }
        /// <summary>
        /// 翻譯自: GetFTAccByPgm
        /// </summary>
        public async Task<FtAccDto> GetFTAccByPgmAsync(string stepName, string bodySize, string pgId, string pgName)
        {
            // 統一參考 DB FUNCTION FUN_GET_FT_ACC_BYPGM[cite: 6]
            string sql = @"
                SELECT fun_get_ft_acc_bypgm(:StepName, :BodySize, :PgId, :PgName, 'loadboard') as Loadboard,
                       fun_get_ft_acc_bypgm(:StepName, :BodySize, :PgId, :PgName, 'contactboard') as Contactboard,
                       fun_get_ft_acc_bypgm(:StepName, :BodySize, :PgId, :PgName, 'burninboard') as Burninboard
                FROM DUAL";

            return await _dbConnection.QueryFirstOrDefaultAsync<FtAccDto>(sql, 
                new { StepName = stepName, BodySize = bodySize, PgId = pgId, PgName = pgName });
        }

        /// <summary>
        /// 翻譯自: GetPgIdAndNameFromFutAct
        /// </summary>
        public async Task<FutActPgmDto> GetPgIdAndNameFromFutActAsync(string lotId, string stepNo, string eqType2, string subSysType, string path)
        {
            // Get PgId and PgName from table "Tbl_Lot_FutAct"[cite: 7]
            // 並限制 HoldPosition = 'STEP_IN'[cite: 7]
            string sql = @"
                SELECT PGID as PgId, PGNAME as PgName, DOC_NO as DocNo, PGMODE as PgMode
                FROM TBL_LOT_FUTACT
                WHERE LOT_ID = :LotId AND STEP_NO = :StepNo AND DELETE_FLAG = 'N'
                  AND HOLD_POSITION = 'STEP_IN' ";

            if (!string.IsNullOrEmpty(eqType2))
            {
                sql += " AND NVL(EQTYPE2, :EqType2) = :EqType2 ";
            }
            if (!string.IsNullOrEmpty(path))
            {
                sql += " AND PATH = :Path ";
            }
            sql += " ORDER BY CREATE_TIME DESC ";

            // 註：SubSystem 的動態組裝條件可在此依傳入的 subSysType 切割後加上
            return await _dbConnection.QueryFirstOrDefaultAsync<FutActPgmDto>(sql, 
                new { LotId = lotId, StepNo = stepNo, EqType2 = eqType2, Path = path });
        }

        /* ---------------------------------------------------------
           GetTecnPgmRecipeAttr 相關 SQL
        --------------------------------------------------------- */
        // Interface 宣告也要記得改
        Task<dynamic> GetIpnMasterForTecnAsync(string lotId);

        // Implementation 實作
        public async Task<dynamic> GetIpnMasterForTecnAsync(string lotId)
        {
            string sql = @"
                SELECT a.IPN, a.PRODGROUPKEY, a.MASK_OPTION, a.BE_OPTION 
                FROM TBL_LOT_ATTRIBUTE tlatt
                INNER JOIN TBL_IPN_MASTER a ON tlatt.IPN = a.IPN
                WHERE tlatt.LOTID = :LotId";
            
            return await _dbConnection.QueryFirstOrDefaultAsync(sql, new { LotId = lotId });
        }

        public async Task<string> GetTecnControlActionAsync(string tecnNo, string lotId)
        {
            // 查詢 TBL_LOT_TECN_CONTROL_LIST 的 Action[cite: 4]
            string sql = @"
                SELECT ACTION FROM TBL_LOT_TECN_CONTROL_LIST
                WHERE TECNNO = :TecnNo AND LOTID = :LotId AND DELETEFLAG = 'N'
                ORDER BY CREATETIME DESC, ACTIONTIME DESC";
            return await _dbConnection.QueryFirstOrDefaultAsync<string>(sql, new { TecnNo = tecnNo, LotId = lotId });
        }

        /* ---------------------------------------------------------
           GetSwapPGName 相關 SQL
        --------------------------------------------------------- */
        public async Task<IEnumerable<string>> GetProdStepSpecStepsAsync(string prodGroup)
        {
            string sql = @"SELECT DISTINCT STEP_NAME FROM TBL_PROD_STEP_SPEC 
                           WHERE PROD_GROUP = :ProdGroup AND DOC_STATUS = 'Active' AND STEP_NAME LIKE 'SORT%' 
                           ORDER BY STEP_NO";
            return await _dbConnection.QueryAsync<string>(sql, new { ProdGroup = prodGroup });
        }

        public async Task<IEnumerable<dynamic>> GetLotPgmRecSequenceAsync(string lotId)
        {
            string sql = @"SELECT PRODGROUP, STEP, PGM FROM VIEW_WS_LOT_PGM_REC_SEQUENCE WHERE CURLOTID = :LotId";
            return await _dbConnection.QueryAsync(sql, new { LotId = lotId });
        }

        public async Task<IEnumerable<dynamic>> GetWsPgmReplaceAsync(string prodGroup)
        {
            string sql = @"SELECT STEP1, STEP2, STEP3, STEP4, STEP5, STEP6, STEP7, STEP8, STEP9, STEP10, STEP11, STEP12, STEP13, STEP14, STEP15,
                                  PGM1, PGM2, PGM3, PGM4, PGM5, PGM6, PGM7, PGM8, PGM9, PGM10, PGM11, PGM12, PGM13, PGM14, PGM15,
                                  SWAPFLAG1, SWAPFLAG2, SWAPFLAG3, SWAPFLAG4, SWAPFLAG5, SWAPFLAG6, SWAPFLAG7, SWAPFLAG8, SWAPFLAG9, SWAPFLAG10, SWAPFLAG11, SWAPFLAG12, SWAPFLAG13, SWAPFLAG14, SWAPFLAG15
                           FROM TBL_WS_PGM_REPLACE WHERE PRODGROUP = :ProdGroup";
            return await _dbConnection.QueryAsync(sql, new { ProdGroup = prodGroup });
        }

        // Interface 宣告也要記得改
        Task<dynamic> GetIpnMasterForTecnAsync(string lotId);

        // Implementation 實作
        public async Task<dynamic> GetIpnMasterForTecnAsync(string lotId)
        {
            string sql = @"
                SELECT a.IPN, a.PRODGROUPKEY, a.MASK_OPTION, a.BE_OPTION 
                FROM TBL_LOT_ATTRIBUTE tlatt
                INNER JOIN TBL_IPN_MASTER a ON tlatt.IPN = a.IPN
                WHERE tlatt.LOTID = :LotId";
            
            return await _dbConnection.QueryFirstOrDefaultAsync(sql, new { LotId = lotId });
        }

        public async Task<string> GetTecnControlActionAsync(string tecnNo, string lotId)
        {
            // 查詢 TBL_LOT_TECN_CONTROL_LIST 的 Action[cite: 4]
            string sql = @"
                SELECT ACTION FROM TBL_LOT_TECN_CONTROL_LIST
                WHERE TECNNO = :TecnNo AND LOTID = :LotId AND DELETEFLAG = 'N'
                ORDER BY CREATETIME DESC, ACTIONTIME DESC";
            return await _dbConnection.QueryFirstOrDefaultAsync<string>(sql, new { TecnNo = tecnNo, LotId = lotId });
        }

        /* ---------------------------------------------------------
           GetSwapPGName 相關 SQL
        --------------------------------------------------------- */
        public async Task<IEnumerable<string>> GetProdStepSpecStepsAsync(string prodGroup)
        {
            string sql = @"SELECT DISTINCT STEP_NAME FROM TBL_PROD_STEP_SPEC 
                           WHERE PROD_GROUP = :ProdGroup AND DOC_STATUS = 'Active' AND STEP_NAME LIKE 'SORT%' 
                           ORDER BY STEP_NO";
            return await _dbConnection.QueryAsync<string>(sql, new { ProdGroup = prodGroup });
        }

        public async Task<IEnumerable<dynamic>> GetLotPgmRecSequenceAsync(string lotId)
        {
            string sql = @"SELECT PRODGROUP, STEP, PGM FROM VIEW_WS_LOT_PGM_REC_SEQUENCE WHERE CURLOTID = :LotId";
            return await _dbConnection.QueryAsync(sql, new { LotId = lotId });
        }

        public async Task<IEnumerable<dynamic>> GetWsPgmReplaceAsync(string prodGroup)
        {
            string sql = @"SELECT STEP1, STEP2, STEP3, STEP4, STEP5, STEP6, STEP7, STEP8, STEP9, STEP10, STEP11, STEP12, STEP13, STEP14, STEP15,
                                  PGM1, PGM2, PGM3, PGM4, PGM5, PGM6, PGM7, PGM8, PGM9, PGM10, PGM11, PGM12, PGM13, PGM14, PGM15,
                                  SWAPFLAG1, SWAPFLAG2, SWAPFLAG3, SWAPFLAG4, SWAPFLAG5, SWAPFLAG6, SWAPFLAG7, SWAPFLAG8, SWAPFLAG9, SWAPFLAG10, SWAPFLAG11, SWAPFLAG12, SWAPFLAG13, SWAPFLAG14, SWAPFLAG15
                           FROM TBL_WS_PGM_REPLACE WHERE PRODGROUP = :ProdGroup";
            return await _dbConnection.QueryAsync(sql, new { ProdGroup = prodGroup });
        }

        /// <summary>
        /// 取得下一個架機單號 (如: ST20260805001)
        /// 翻譯自 InsNewForm 中產生 FormNo 的 SQL
        /// </summary>
        public async Task<string> GetNextSetupFormNoAsync()
        {
            string sql = @"
                SELECT 'ST' || TO_CHAR(SYSDATE, 'YYYYMMDD') || 
                       DECODE(LENGTH(MAX(SUBSTR(FORMNO, 11, 3)) + 1), 
                              1, '00' || (MAX(SUBSTR(FORMNO, 11, 3)) + 1), 
                              2, '0' || (MAX(SUBSTR(FORMNO, 11, 3)) + 1), 
                              NULL, '001', 
                              MAX(SUBSTR(FORMNO, 11, 3)) + 1)
                FROM TBL_WS_EQFORM_BASIC
                WHERE FORMTYPE = 'SETUP' 
                  AND FORMNO LIKE 'ST' || TO_CHAR(SYSDATE, 'YYYYMMDD') || '%'";

            return await _dbConnection.QueryFirstOrDefaultAsync<string>(sql);
        }

        public async Task<int> InsertWsEqFormBasicAsync(object param)
        {
            // 對應 InsNewForm 的 Insert SQL
            string sql = @"
                INSERT INTO TBL_WS_EQFORM_BASIC (
                    FORMNO, FORMTYPE, IPN, LOTID, LOTOWNER, ERUNTICNO, WAFERID, 
                    TESTERID, PROBERID, TESTMODE, PGM, PGID, PGMMATCH, CODE, CHECKSUM, 
                    SPEED, TEMPERATURE, DEVICEFILE, TECNNO, PROBECARD, EQCOMMENT, STEPCOMMENT, 
                    CREATEUSERID, CREATETIME, SETUP_REASON, 
                    SETUPWAFERID_1, SETUPWAFERID_2, SETUPWAFERID_3, STEPID, INSTEPTIME, STOPINFORMATION
                ) VALUES (
                    :FormNo, 'SETUP', :Ipn, :LotId, :LotOwner, :ErunTicNo, :WaferId,
                    :TesterId, :ProberId, :TestMode, :PgName, :PgId, :PgmMatch, :Code, :CheckSum,
                    :Speed, :Temp, :WsDeviceFile, :TecnNo, :ProbeCard, :EqComments, :StepComments,
                    :CreateUserId, TO_CHAR(SYSDATE, 'YYYYMMDD HH24MISS')||'000', :SetupReason,
                    :SetupWaferId1, :SetupWaferId2, :SetupWaferId3, :StepId, :InStepTime, :StopInfo
                )";

            return await _dbConnection.ExecuteAsync(sql, param);
        }

        public async Task<int> UpdateEqInfoFormNoAsync(string eqId, string formNo, string userId)
        {
            // 更新 TBL_EQ_INFO 的單號與時間
            string sql = @"
                UPDATE TBL_EQ_INFO 
                SET FORM_NO = :FormNo, 
                    UPDATE_USER_ID = :UserId, 
                    UPDATE_TIME = TO_CHAR(SYSDATE, 'YYYYMMDD HH24MISS')||'000'
                WHERE EQ_ID = :EqId";

            return await _dbConnection.ExecuteAsync(sql, new { FormNo = formNo, UserId = userId, EqId = eqId });
        }

        public async Task<string> GetAssignWaferIdAsync(string lotId)
        {
            // 呼叫 DB Function
            string sql = "SELECT Fun_GetAssignWaferID(:LotId) FROM DUAL";
            return await _dbConnection.QueryFirstOrDefaultAsync<string>(sql, new { LotId = lotId });
        }
        
        // 在 PrintSetupFormRepository 實作
        public async Task<IEnumerable<string>> GetSetupReasonsAsync()
        {
            // 對應 VB6 查詢 Tbl_ReasonCode 的語法[cite: 3]
            string sql = @"
                SELECT REASONCODE 
                FROM TBL_REASONCODE 
                WHERE CATEGORY = 'SetupFomReason' 
                  AND DELETEFLAG = 'N'";

            return await _dbConnection.QueryAsync<string>(sql);
        }

        //----

        public async Task<string> GetLotStepEqSpecContinueTestAsync(string tecnLotId, string stepNo, string path, string eqType2, string subSystem)
        {
            string sql = @"SELECT CONTINUETEST FROM TBL_LOT_STEP_EQ_SPEC 
                           WHERE :TecnLotId LIKE TECNLOTID AND STEPNO = :StepNo AND PATH = :Path AND EQTYPE2 = :EqType2 
                             AND NVL(TRIM(SUBSYSTEM), ' ') = NVL(TRIM(:SubSystem), ' ') AND DELETEFLAG = 'N'";
            return await _dbConnection.QueryFirstOrDefaultAsync<string>(sql, new { TecnLotId = tecnLotId, StepNo = stepNo, Path = path, EqType2 = eqType2, SubSystem = subSystem });
        }

        public async Task<EqAccessoriesDto> GetEqInfoAccessoriesAsync(string eqId)
        {
            string sql = @"SELECT PROBECARD_ID as ProbeCardId, LOADBOARD_ID as LoadBoardId, CONTACTBOARD_ID as ContactBoardId, WSDEVICE_FILE as WsDeviceFile 
                           FROM TBL_EQ_INFO WHERE EQ_ID = :EqId";
            return await _dbConnection.QueryFirstOrDefaultAsync<EqAccessoriesDto>(sql, new { EqId = eqId });
        }

        public async Task<dynamic> GetErunReqContinueTestAsync(string lotId, string erunTicNo, string stage)
        {
            string sql = @"SELECT FOLLOW_PRODUCT as FollowProd, CONTINUETEST as ContinueTest 
                           FROM TBL_ERUN_REQ WHERE LOT_ID = :LotId AND TICKET_NO = :TicNo AND STAGE = :Stage";
            return await _dbConnection.QueryFirstOrDefaultAsync(sql, new { LotId = lotId, TicNo = erunTicNo, Stage = stage });
        }

        public async Task<IEnumerable<ProductStopTestRecordDto>> GetProductStopTestRecordsAsync(
            string eqType2, string prodGroup, string ipn, string pgId, string pgName, string stepName, string concatAcc, string lotId, string curAccNo, string wsDeviceFile)
        {
            // 翻譯自 GetProductStopTest 複雜的 JOIN SQL
            string sql = @"
                SELECT d.STOPTICNO as StopTicNo, d.SPECIFYEQ as SpecifyEq, d.EQID as EqId, d.ACCNAME as AccName, 
                       d.ACCNO as AccNo, d.DEVICEFILE as DeviceFile, d.BEPE_SET as BePeSet
                FROM (
                    SELECT aa.productname as prodgroup, dd.IPN,
                           SUBSTR(NVL(cc.PROD_GROUP_KEY, dd.PRODGROUPKEY), 1, 5) as prodcode,
                           SUBSTR(NVL(cc.PROD_GROUP_KEY, dd.PRODGROUPKEY), 1, 4) as prodbody,
                           NVL(cc.PROD_GROUP_KEY, dd.PRODGROUPKEY) as prodgroupkey
                    FROM fwproductversion aa
                    INNER JOIN fwproductversion_n2m bb ON bb.fromid = aa.sysid AND bb.linkname = 'attributes' AND bb.keydata = 'TdsProd'
                    LEFT JOIN TBL_PROD_SPEC cc ON aa.productname = cc.prodgroup AND cc.DOC_STATUS = 'Active' AND cc.DELETE_FLAG = 'N'
                    LEFT JOIN TBL_IPN_MASTER dd ON aa.productname = dd.PROD_GROUP
                    WHERE aa.revstate = 'Active'
                ) c
                INNER JOIN TBL_STOP_TEST d ON c.prodgroup = :ProdGroup AND d.EQTYPE2 = :EqType2 AND d.DELETEFLAG = 'N' AND d.BEPE_SET = 'A'
                WHERE NVL(c.IPN, ' ') LIKE NVL(REPLACE(REPLACE(d.IPN, '%', '_'), '*', '%'), NVL(c.IPN, ' '))
                  AND NVL(:PgId, ' ') = NVL(d.PGID, NVL(:PgId, ' '))
                  AND NVL(:PgName, ' ') = NVL(d.PGNAME, NVL(:PgName, ' '))
                  AND NVL(c.prodcode, ' ') = NVL(d.PRODCODE, NVL(c.prodcode, ' '))
                  AND :StepName = NVL(d.STEPNAME, :StepName)
                  AND (NVL(:ConcatAcc, ' ') = NVL(d.ACCNAME, NVL(:ConcatAcc, ' ')) OR INSTR(d.ACCNAME, :ConcatAcc) > 0)
                  AND NVL(:LotId, ' ') LIKE NVL(REPLACE(REPLACE(d.LOTID, '%', '_'), '*', '%'), NVL(:LotId, ' '))
                  AND NVL(c.prodgroupkey, ' ') = NVL(d.PRODGROUPKEY, NVL(c.prodgroupkey, ' '))
                  AND NVL(c.prodbody, ' ') = NVL(d.PRODBODY, NVL(c.prodbody, ' '))
                  AND (NVL(:CurAccNo, ' ') = NVL(d.ACCNO, NVL(:CurAccNo, ' ')) OR INSTR(:CurAccNo || ',', d.ACCNO || ',') > 0)
                  AND NVL(:WsDeviceFile, ' ') = NVL(d.DEVICEFILE, NVL(:WsDeviceFile, ' '))
                  AND NVL(:Ipn, ' ') LIKE NVL(REPLACE(REPLACE(d.IPN, '%', '_'), '*', '%'), NVL(:Ipn, ' '))
                ORDER BY d.SPECIFYEQ DESC, DECODE(d.EQID, NULL, 'A', 'B') || d.STOPTICNO, d.ACCNAME DESC";

            return await _dbConnection.QueryAsync<ProductStopTestRecordDto>(sql, new 
            { 
                ProdGroup = prodGroup, EqType2 = eqType2, PgId = pgId, PgName = pgName, 
                StepName = stepName, ConcatAcc = concatAcc, LotId = lotId, CurAccNo = curAccNo, 
                WsDeviceFile = wsDeviceFile, Ipn = ipn 
            });
        }

        public async Task<IEnumerable<string>> GetNonProdStopTestRecordsAsync(
            string eqType2, string curEqId, string curAccName, string packageType, string pinCount, string carrierType, string bodySize, int temperature, string curAccNo, string formFactorName, string moduleOption)
        {
            // 翻譯自 GetNonProdStopTest
            string sql = @"
                SELECT DISTINCT a.STOPTICNO
                FROM TBL_STOP_TEST_NP a
                WHERE a.DELETEFLAG = 'N'
                  AND NVL(:EqType2, ' ') = NVL(a.EQTYPE2, NVL(:EqType2, ' '))
                  AND NVL(:CurEqId, ' ') = NVL(a.TESTERID, NVL(:CurEqId, ' '))
                  AND NVL(:CurAccName, ' ') = NVL(a.ACCNAME, NVL(:CurAccName, ' '))
                  AND NVL(:PackageType, ' ') = NVL(a.PACKAGETYPE, NVL(:PackageType, ' '))
                  AND NVL(:PinCount, '0') = NVL(TO_CHAR(a.PINCOUNT), NVL(:PinCount, '0'))
                  AND NVL(:CarrierType, ' ') = NVL(a.CARRIERTYPE, NVL(:CarrierType, ' '))
                  AND NVL(:BodySize, ' ') = NVL(a.BODYSIZE, NVL(:BodySize, ' '))
                  AND NVL(:Temp, 0) >= NVL(a.TEMP_MIN, NVL(:Temp, 0))
                  AND NVL(:Temp, 0) <= NVL(a.TEMP_MAX, NVL(:Temp, 0))
                  AND NVL(:FormFactorName, ' ') = NVL(a.FORM_FACTOR_NAME, NVL(:FormFactorName, ' '))
                  AND NVL(:ModuleOption, ' ') = NVL(a.MODULE_OPTION, NVL(:ModuleOption, ' ')) ";

            if (!string.IsNullOrEmpty(curAccNo))
                sql += " AND (FUN_ARRAYPOS(:CurAccNo, a.ACCNO) > 0 OR a.ACCNO IS NULL) ";
            else
                sql += " AND a.ACCNO IS NULL ";

            return await _dbConnection.QueryAsync<string>(sql, new 
            { 
                EqType2 = eqType2, CurEqId = curEqId, CurAccName = curAccName, PackageType = packageType, 
                PinCount = pinCount, CarrierType = carrierType, BodySize = bodySize, Temp = temperature, 
                CurAccNo = curAccNo, FormFactorName = formFactorName, ModuleOption = moduleOption 
            });
        }
        // === 1. 取得機台的 SubSystem ===
        public async Task<dynamic> GetEqpSubSysAsync(string eqId)
        {
            string sql = "SELECT SUBSYS1, SUBSYS2 FROM TBL_EQ_INFO WHERE EQ_ID = :EqId";
            return await _dbConnection.QueryFirstOrDefaultAsync(sql, new { EqId = eqId });
        }

        // === 2. 取得 IPN Master 的 Code, CheckSum, Speed ===
        public async Task<dynamic> GetIpnMasterDetailsAsync(string ipn)
        {
            string sql = "SELECT SPEED, CODE, CHECKSUM FROM TBL_IPN_MASTER WHERE IPN = :Ipn";
            return await _dbConnection.QueryFirstOrDefaultAsync(sql, new { Ipn = ipn });
        }

        // === 3. 檢查是否為 MCD (多Code) 或 MCP (多晶片) ===
        public async Task<bool> CheckIsMcdOrMcpAsync(string ipn)
        {
            // Check MCD
            string mcdSql = "SELECT COUNT(*) FROM TBL_MULTICODE_IPN WHERE IPN = :Ipn AND DELETEFLAG = 'N'";
            int mcdCount = await _dbConnection.QueryFirstOrDefaultAsync<int>(mcdSql, new { Ipn = ipn });
            if (mcdCount > 0) return true;

            // Check MCP
            string mcpSql = @"
                SELECT b.PACKAGECOMPONENT 
                FROM TBL_IPN_BOM a
                INNER JOIN TBL_PROD_BODY b ON b.PROD_BODY = SUBSTR(a.IPN, 1, 4) AND b.DELETE_FLAG = 'N'
                WHERE a.PARENT = :Ipn AND a.BOM_LEVEL = 'FG' AND a.DELETE_FLAG = 'N'";
            string component = await _dbConnection.QueryFirstOrDefaultAsync<string>(mcpSql, new { Ipn = ipn });
            
            if (!string.IsNullOrEmpty(component) && component.StartsWith("MCP")) return true;

            return false;
        }

        // === 4. 取得工程品 (Erun) 的覆寫資料 ===
        public async Task<dynamic> GetErunOverrideDataAsync(string lotId, string erunTicNo, string stage, string stepNo, string eqType2, string subSystem)
        {
            // TBL_ERUN_REQ / TBL_ERUN_RECIPE (工程品 完整版)
            string sql = @"
                SELECT a.SPEED, a.CODE, a.CHECKSUM, 
                       b.PGNAME, b.TEMPERATURE, b.WSDEVICEFILE, b.PROBECARDTYPE, b.LOADBOARDTYPE, b.PGID, b.JUMPERPINNO
                FROM TBL_ERUN_REQ a
                LEFT JOIN TBL_ERUN_RECIPE b ON b.DOCNO = a.TICKET_NO AND b.STEPNO = :StepNo AND b.EQTYPE2 = :EqType2 AND b.DELETEFLAG = 'N' 
                     AND NVL(b.SUBSYSTEM, ' ') = NVL(:SubSystem, ' ')
                WHERE a.LOT_ID = :LotId AND a.TICKET_NO = :TicNo AND a.STAGE = :Stage";
                
            return await _dbConnection.QueryFirstOrDefaultAsync(sql, new { LotId = lotId, TicNo = erunTicNo, Stage = stage, StepNo = stepNo, EqType2 = eqType2, SubSystem = string.IsNullOrEmpty(subSystem) ? " " : subSystem });
        }
        
public async Task<RecipeSpecData> GetLotStepEqSpecAsync(string tecnLotId, string stepNo, string eqType2, string subSystem, string path, string maxSite)
        {
            // TBL_LOT_STEP_EQ_SPEC (TECN 完整版)
            string sql = @"
                SELECT PGNAME as PgName, PGID as PgId, PGMODE as PgMode, TEMPERATURE as Temperature, 
                       PROBECARDTYPE as ProbeCardType, LOADBOARDTYPE as LoadboardType, 
                       CONTACTBOARDTYPE as ContactboardType, WSDEVICEFILE as WsDeviceFile, 
                       SPECIFYEQ as SpecifyEq, EQID as EqId, BURNINBOARD as BurnInBoard,
                       REFSTEPNAME01 as RefStepName01, REFPGNAME01 as RefPgName01, REPLACEPGNAME01 as ReplacePgName01,
                       TECNNO as TecnNo, STEPCOMMENT as StepComments,
                       CABLETYPE as CableType, KITTYPE as KitType, NEEDJUMPER as NeedJumper, JUMPERPINNO as JumperPinNo
                FROM TBL_LOT_STEP_EQ_SPEC
                WHERE TECNLOTID LIKE :TecnLotId AND STEPNO = :StepNo AND EQTYPE2 = :EqType2 
                  AND NVL(TRIM(SUBSYSTEM), ' ') = NVL(TRIM(:SubSystem), ' ') AND PATH = :Path AND NVL(MAXSITE, ' ') = NVL(:MaxSite, ' ')
                  AND DELETEFLAG = 'N'";
                  
            return await _dbConnection.QueryFirstOrDefaultAsync<RecipeSpecData>(sql, new { TecnLotId = tecnLotId, StepNo = stepNo, EqType2 = eqType2, SubSystem = subSystem, Path = path, MaxSite = maxSite });
        }

        public async Task<RecipeSpecData> GetProdStepEqSpecAsync(string prodGroup, string stepNo, string eqType2, string subSystem, string path, string maxSite)
        {
            // TBL_PROD_STEP_EQ_SPEC (Normal 完整版)
            string sql = @"
                SELECT PG_NAME as PgName, PG_ID as PgId, PG_MODE as PgMode, TEMPERATURE as Temperature, 
                       PROBECARD_TYPE as ProbeCardType, LOADBOARD_TYPE as LoadboardType, 
                       CONTACTBOARD_TYPE as ContactboardType, WS_DEVICE_FILE as WsDeviceFile, 
                       SPECIFYEQ as SpecifyEq, EQID as EqId, STOPTICNO as StopTicNo, BURN_IN_BOARD as BurnInBoard,
                       REF_STEP_NAME_01 as RefStepName01, REF_PG_NAME_01 as RefPgName01, REPLACE_PG_NAME_01 as ReplacePgName01,
                       STEPCOMMENT as StepComments, DOC_TYPE as DocType, DOC_NO as DocNo,
                       CABLE_TYPE as CableType, KIT_TYPE as KitType, NEEDJUMPER as NeedJumper, JUMPERPINNO as JumperPinNo
                FROM TBL_PROD_STEP_EQ_SPEC
                WHERE PRODGROUP = :ProdGroup AND STEPNO = :StepNo AND EQTYPE2 = :EqType2 
                  AND NVL(TRIM(SUBSYSTEM), ' ') = NVL(TRIM(:SubSystem), ' ') AND PATH = :Path AND NVL(MAX_SITE, ' ') = NVL(:MaxSite, ' ')
                  AND DOCSTATUS = 'Active'";
                  
            return await _dbConnection.QueryFirstOrDefaultAsync<RecipeSpecData>(sql, new { ProdGroup = prodGroup, StepNo = stepNo, EqType2 = eqType2, SubSystem = subSystem, Path = path, MaxSite = maxSite });
        }
        public async Task<string> GetShiftCodeAsync(string empNo)
        {
            string sql = "SELECT SHIFTCODE FROM TBL_MXIC_EMP WHERE EMPNO = :EmpNo";
            return await _dbConnection.QueryFirstOrDefaultAsync<string>(sql, new { EmpNo = empNo });
        }

        public async Task<string> GetSoakTimeAsync(string eqType2, string pinCount, string packageCode)
        {
            string sql = "SELECT SOAKTIME FROM TBL_SUBSYS_CONFI WHERE EQTYPE2 = :EqType2 AND PINCOUNT = :PinCount AND PACKAGECODE = :PkgCode";
            return await _dbConnection.QueryFirstOrDefaultAsync<string>(sql, new { EqType2 = eqType2, PinCount = pinCount, PkgCode = packageCode });
        }

        public async Task<string> GetSthandFileAsync(string loadBoard)
        {
            string sql = "SELECT STHAND_FILE FROM TBL_ACC_GROUP WHERE LOADBOARD = :LoadBoard";
            return await _dbConnection.QueryFirstOrDefaultAsync<string>(sql, new { LoadBoard = loadBoard });
        }

        public async Task<dynamic> GetEqInfoAccessoriesFullAsync(string eqId)
        {
            string sql = "SELECT LOADBOARDID, CABLEID, CONTACTBOARDID, KITID, MATCHPLATE, SUBSYS1, SUBSYS2 FROM TBL_EQ_INFO WHERE EQ_ID = :EqId";
            return await _dbConnection.QueryFirstOrDefaultAsync(sql, new { EqId = eqId });
        }

        public async Task<dynamic> GetTesterFviAccAsync(string pinCount, string pkgCode)
        {
            string sql = "SELECT PITCH, VACUUMCUP FROM TBL_TESTER_FVI_ACC WHERE PINCOUNT = :PinCount AND PKGCODE = :PkgCode AND DELETEFLAG = 'N'";
            return await _dbConnection.QueryFirstOrDefaultAsync(sql, new { PinCount = pinCount, PkgCode = pkgCode });
        }

        public async Task<string> GetSpecCommentsAsync(bool isTecn, string id, string stepNo, string path)
        {
            string sql = isTecn 
                ? "SELECT COMMENTS FROM TBL_LOT_STEP_SPEC WHERE :Id LIKE TECNLOTID AND STEPNO = :StepNo AND DELETEFLAG = 'N' AND PATH = :Path"
                : "SELECT COMMENTS FROM TBL_PROD_STEP_SPEC WHERE PRODGROUP = :Id AND STEPNO = :StepNo AND DOCSTATUS = 'Active' AND PATH = :Path";
            return await _dbConnection.QueryFirstOrDefaultAsync<string>(sql, new { Id = id, StepNo = stepNo, Path = path });
        }

// ==========================================
        // 2. 實作 GetAccMatchPlateAsync (翻譯自 GetAccMatchPlate)
        // ==========================================
        public async Task<string> GetAccMatchPlateAsync(string ipn, string kit, string loadBoard, string eqType2)
        {
            string kitType = !string.IsNullOrEmpty(kit) && kit.Contains("-") ? kit.Split('-')[0] : kit;
            string loadBoardType = !string.IsNullOrEmpty(loadBoard) && loadBoard.Contains("-") ? loadBoard.Split('-')[0] : loadBoard;

            // Step 1: 取得 IPN 基本資訊與 AccSite (將兩個查詢合併以增進效能)
            string infoSql = @"
                SELECT a.PACKAGE_CODE as PkgCode, a.PIN_COUNT as PinCount, a.BODY_SIZE as BodySize,
                       (SELECT ACCSITE FROM TBL_ACC_BASIC WHERE ACC_NAME = :LoadBoardType AND DELETE_FLAG = 'N') as AccSite
                FROM TBL_IPN_MASTER a WHERE a.IPN = :Ipn";
            
            var info = await _dbConnection.QueryFirstOrDefaultAsync(infoSql, new { LoadBoardType = loadBoardType, Ipn = ipn });
            if (info == null) return "";

            // Step 2: 取得 MatchPlate
            string matchSql = @"
                SELECT MATCHPLATE FROM TBL_ACC_MATCHPLATE 
                WHERE DELETEFLAG = 'N' AND KIT = :KitType AND ACCSITE = :AccSite AND PINCOUNT = :PinCount 
                  AND PACKAGECODE = :PkgCode AND BODYSIZE = :BodySize AND LOADBOARDTYPE = :LoadBoardType AND TESTERTYPE = :EqType2";

            return await _dbConnection.QueryFirstOrDefaultAsync<string>(matchSql, new 
            { 
                KitType = kitType, AccSite = info.ACCSITE, PinCount = info.PINCOUNT, PkgCode = info.PKGCODE, 
                BodySize = info.BODYSIZE, LoadBoardType = loadBoardType, EqType2 = eqType2 
            }) ?? "";
        }

        // ==========================================
        // 3. 實作 GetRecipeSpecNamesAsync (翻譯自 getRecipeSpecName & getRecipeSpecName1)
        // ==========================================
        public async Task<IEnumerable<RecipeItemDto>> GetRecipeSpecNamesAsync(string eqId, string lotId, string ipn, string packageName, string pinCount, string bodySize)
        {
            // Check 是否為 ICOS (SCANNER) 機台
            string checkIcosSql = "SELECT COUNT(1) FROM view_b2b_fweqarea WHERE eqtype = 'SCANNER' AND EQID = :EqId";
            int isIcos = await _dbConnection.QueryFirstOrDefaultAsync<int>(checkIcosSql, new { EqId = eqId });

            string recipeSql = "";

            if (isIcos > 0)
            {
                // SCANNER 邏輯：呼叫 getRecipeSpecName1 的 FUN_GET_SCANPGM
                string scanPgmSql = "SELECT FUN_GET_SCANPGM(:LotId, 'Y', :EqId) FROM DUAL";
                string pgm = await _dbConnection.QueryFirstOrDefaultAsync<string>(scanPgmSql, new { LotId = lotId, EqId = eqId });

                recipeSql = @"
                    SELECT ROW_NUMBER() OVER(ORDER BY SEQUENCE) as Id, SPECNAME as SpecName 
                    FROM TBL_SETUP_RECIPE 
                    WHERE EQID = :EqId AND PGM = :Pgm
                    ORDER BY SEQUENCE";
                return await _dbConnection.QueryAsync<RecipeItemDto>(recipeSql, new { EqId = eqId, Pgm = pgm });
            }
            else
            {
                // 一般機台邏輯
                recipeSql = @"
                    SELECT ROW_NUMBER() OVER(ORDER BY SEQUENCE) as Id, SPECNAME as SpecName 
                    FROM TBL_SETUP_RECIPE 
                    WHERE EQID = :EqId AND PACKAGENAME = :PackageName AND PINCOUNT = :PinCount AND BODYSIZE = :BodySize 
                    ORDER BY SEQUENCE";
                return await _dbConnection.QueryAsync<RecipeItemDto>(recipeSql, new { EqId = eqId, PackageName = packageName, PinCount = pinCount, BodySize = bodySize });
            }
        }
        public async Task<TecnPgmRecipeAttrDto> GetTecnPgmRecipeAttrAsync(string lotId, string stepNo, string eqType2, string subSystem)
        {
            // TODO: 請依據 VB6 modTecn.GetTecnPgmRecipeAttr 的實際 SQL 替換以下語法
            // 概念：透過 LotId 取得 TECN_LOT_ID，再去查對應的 Recipe/Temp 覆蓋設定
            string sql = @"
                SELECT REF_PGM_TECN_NO as RefPgmTecnNo, 
                       REF_PGM as RefPgm, 
                       REF_PGID as RefPgId,
                       REF_TEMP_TECN_NO as RefTempTecnNo, 
                       REF_TEMP as RefTemp
                FROM TBL_TECN_PGM_RECIPE  /* 這裡請替換為實際的 TECN 設定表 */
                WHERE LOT_ID = :LotId 
                  AND STEP_NO = :StepNo 
                  AND EQ_TYPE = :EqType2 
                  AND NVL(SUBSYSTEM, ' ') = NVL(:SubSystem, ' ')";

            return await _dbConnection.QueryFirstOrDefaultAsync<TecnPgmRecipeAttrDto>(sql, new 
            { 
                LotId = lotId, 
                StepNo = stepNo, 
                EqType2 = eqType2,
                SubSystem = string.IsNullOrEmpty(subSystem) ? " " : subSystem
            });
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
        Task<PrintSetupFormSubmitResponse> SubmitSetupFormAsync(PrintSetupFormSubmitRequest request);
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

            // 1. 組裝 SQL (將 AssignLoadBoard 也傳入以利 SQL 判斷)
            string subSystemSql = BuildSubSystemSql(response.Stage, erunTicNo, followProd, tecnLotId, assignLoadBoard);
            
            // 2. 呼叫 Repository 執行查詢 (這裡注意要使用 dynamic 來接回傳的三個欄位)
            var subSystemsResult = await _repo.GetSubSystemsDynamicAsync(subSystemSql, new
            {
                ErunTicNo = erunTicNo,
                StepNo = stepId,
                EqType2 = eqType2,
                Path = path,
                TecnLotId = tecnLotId,
                AssignProbeCard = assignProbeCard,
                AssignLoadBoard = assignLoadBoard,
                BodySize = bodySize,
                ProdGroup = prodGroup
            });
            
            // 3. 整理並過濾 SubSystem 選單格式
            foreach (var item in subSystemsResult)
            {
                // 🌟 這裡直接取用 DTO 的屬性，超級乾淨！
                string sys = item.SubSystem ?? "";
                string maxSite = item.MaxSite ?? "";
                string loadboard = item.LoadBoard ?? "";
                
                string displayText = string.IsNullOrEmpty(maxSite) ? sys : $"{sys},{maxSite}";
        
                if (response.Stage == "FT")
                {
                    if (!string.IsNullOrEmpty(assignLoadBoard))
                    {
                        if (assignLoadBoard == loadboard)
                        {
                            response.SubSystemList.Add(new SelectItem { Text = displayText, Value = displayText });
                        }
                    }
                    else
                    {
                        response.SubSystemList.Add(new SelectItem { Text = displayText, Value = displayText });
                    }
                }
                else
                {
                    response.SubSystemList.Add(new SelectItem { Text = displayText, Value = displayText });
                }
            }

            return response;
        }

        private string BuildSubSystemSql(string stage, string erunTicNo, string followProd, string tecnLotId, string assignLoadBoard)
        {
            string sql = "";
        
            // 抽出共用的 Decode 語法 (對應舊 VB 邏輯)
            string decodeSql = @"
                decode(trim(LOADBOARDTYPE || CONTACTBOARDTYPE || BURNINBOARD), null, 
                fun_get_ft_acc_bypgm(STEPNAME, :BodySize, PGID, PGNAME, 'loadboard'), LOADBOARDTYPE) as loadboard";
        
            if (stage != "FT")
            {
                if (!string.IsNullOrEmpty(erunTicNo) && followProd == "N")
                {
                    sql = @"SELECT SUBSYSTEM, '' as MAXSITE, LOADBOARDTYPE as loadboard 
                            FROM TBL_ERUN_RECIPE 
                            WHERE DOCNO = :ErunTicNo AND STEPNO = :StepNo AND EQTYPE2 = :EqType2 AND DELETEFLAG = 'N' AND PATH = :Path";
                    
                    // WS 站別專屬的 ProbeCard 條件
                    if (stage == "WS")
                        sql += " AND NVL(PROBECARDTYPE, ' ') = NVL(:AssignProbeCard, ' ')";
                }
                else if (!string.IsNullOrEmpty(tecnLotId))
                {
                    sql = $@"SELECT SUBSYSTEM, MAXSITE, {decodeSql} 
                             FROM TBL_LOT_STEP_EQ_SPEC 
                             WHERE :TecnLotId LIKE TECNLOTID AND STEPNO = :StepNo AND EQTYPE2 = :EqType2 AND DELETEFLAG = 'N' AND PATH = :Path";
                    
                    if (stage == "WS")
                        sql += " AND NVL(PROBECARDTYPE, ' ') = NVL(:AssignProbeCard, ' ')";
                }
                else
                {
                    // 正規產品規格查詢
                    sql = $@"SELECT SUBSYSTEM, MAXSITE, {decodeSql} 
                             FROM TBL_PROD_STEP_EQ_SPEC 
                             WHERE PRODGROUP = :ProdGroup AND STEPNO = :StepNo AND EQTYPE2 = :EqType2 AND DOCSTATUS = 'Active' AND PATH = :Path";
                    
                    if (stage == "WS")
                        sql += " AND NVL(PROBECARDTYPE, ' ') = NVL(:AssignProbeCard, ' ')";
                }
            }
            else // Stage == "FT"
            {
                if (!string.IsNullOrEmpty(erunTicNo) && followProd == "N")
                {
                    sql = @"SELECT SUBSYSTEM, '' as MAXSITE, LOADBOARDTYPE as loadboard 
                            FROM TBL_ERUN_RECIPE 
                            WHERE DOCNO = :ErunTicNo AND STEPNO = :StepNo AND EQTYPE2 = :EqType2 AND DELETEFLAG = 'N' AND PATH = :Path";
                    
                    // FT 站別且有 AssignLoadboard 的條件
                    if (!string.IsNullOrEmpty(assignLoadBoard))
                        sql += " AND NVL(LOADBOARDTYPE, ' ') = NVL(:AssignLoadBoard, ' ')";
                }
                else if (!string.IsNullOrEmpty(tecnLotId))
                {
                    sql = $@"SELECT SUBSYSTEM, MAXSITE, {decodeSql} 
                             FROM TBL_LOT_STEP_EQ_SPEC 
                             WHERE :TecnLotId LIKE TECNLOTID AND STEPNO = :StepNo AND EQTYPE2 = :EqType2 AND DELETEFLAG = 'N' AND PATH = :Path";
                }
                else
                {
                    sql = $@"SELECT SUBSYSTEM, MAXSITE, {decodeSql} 
                             FROM TBL_PROD_STEP_EQ_SPEC 
                             WHERE PRODGROUP = :ProdGroup AND STEPNO = :StepNo AND EQTYPE2 = :EqType2 AND DOCSTATUS = 'Active' AND PATH = :Path";
                }
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

            // ====================================================================
            // 🌟 準備 TECN 規則引擎所需的參數
            // ====================================================================
            var olot = WipServiceWrapper.Instance.LotById(request.LotId);
            string ipn = olot.CustomAttributes(LotCustomAttributes.Ipn) ?? "";
            string path = olot.CustomAttributes(LotCustomAttributes.Route) ?? "";
            string stepName = olot.CurrentStep.Steps[0].Description ?? "";
            string grade = olot.CustomAttributes("Grade") ?? "C"; // 預設為 C，請依貴司實際擴充屬性調整

            // 取得 IPN 相關的 ProdCode 與 ProdGroupKey
            var ipnDetails = await _repo.GetIpnMasterForTecnAsync(ipn);
            string prodGroupKey = ipnDetails?.PRODGROUPKEY?.ToString() ?? "";
            string prodCode = ""; 
            if (ipnDetails != null)
            {
                // 對應 VB6 中 WS/FT 組裝 ProdCode 的邏輯
                string maskOption = ipnDetails.MASK_OPTION?.ToString() ?? "";
                string beOption = ipnDetails.BE_OPTION?.ToString() ?? "";
                prodCode = request.Stage == "WS" ? (ipn.Substring(0, 4) + maskOption) : (ipn.Substring(0, 4) + beOption);
            }

            // 取得 IpnTecnNo (通常由 TBL_LOT_INFO 取得)
            var lotInfo = await _repo.GetLotInfoAsync(request.LotId);
            string ipnTecnNo = lotInfo?.TECNLOTID?.ToString() ?? ""; 

            // ====================================================================
            // 🌟 呼叫 TECN 屬性覆蓋邏輯 (Service 層的規則引擎)
            // ====================================================================
            var tecnAttr = await GetTecnPgmRecipeAttrAsync(
                ipnTecnNo: ipnTecnNo, 
                lotId: request.LotId, 
                ipn: ipn, 
                path: path, 
                stepId: request.StepNo, 
                stepName: stepName, 
                eqType2: eqType2, 
                prodGroupKey: prodGroupKey, 
                prodCode: prodCode, 
                grade: grade, 
                runRule: "N" // 預設傳 N，除非是由 RunRule 觸發
            );

            if (tecnAttr != null)
            {
                // 若 TECN 有指定 PGM，則覆蓋標準的 PgName
                if (!string.IsNullOrEmpty(tecnAttr.RefPgm))
                {
                    response.PgName = tecnAttr.RefPgm;
                }

                // 若 TECN 有指定 PGID，則覆蓋標準的 PgId
                if (!string.IsNullOrEmpty(tecnAttr.RefPgId))
                {
                    response.PgId = tecnAttr.RefPgId;
                }

                // 若 TECN 有指定溫度，則覆蓋標準的 Temp
                if (!string.IsNullOrEmpty(tecnAttr.RefTemp))
                {
                    response.Temp = tecnAttr.RefTemp;
                }
            }

            return response;
        }

        public partial class PrintSetupFormService
    {
        /// <summary>
        /// 翻譯自: GetPGM (原位於 frmPrintSetupForm.frm 內)
        /// 包含對 GetPgIdAndNameFromFutAct 的隱藏呼叫
        /// </summary>
        private async Task<string> GetPGMAsync(string lotId, string stepNo, string path, string eqType2, string subSysType, string originalPgName, RecipeSpecData specData)
        {
            // 1. 先搜尋 Future Action[cite: 7]
            var futAct = await _repo.GetPgIdAndNameFromFutActAsync(lotId, stepNo, eqType2, subSysType, path);
            
            if (futAct != null && !string.IsNullOrEmpty(futAct.PgName))
            {
                return futAct.PgName; // 若 Future Action 有設定，以此為準
            }

            // 2. 若沒有，則進行屬性替換比對邏輯 (RefStepName, RefPgName -> ReplacePgName)
            string resultPgmName = originalPgName;

            // 假設貴司已封裝取得 Lot_PN2M 擴充屬性 (Tested1stPgName 等) 的 Wrapper
            // var attPg1 = WipServiceWrapper.Instance.LotCustomAttribute(lotId, "Tested1stPgName");
            string attPg1 = ""; // 示意

            if (!string.IsNullOrEmpty(specData.RefStepName01))
            {
                if (specData.RefStepName01 == "1" && attPg1 == specData.RefPgName01)
                {
                    resultPgmName = specData.ReplacePgName01;
                }
                // ... 依序實作 02, 03 的比對邏輯
            }

            return resultPgmName;
        }

        /// <summary>
        /// 翻譯自: GetTecnPgmRecipeAttr
        /// 處理複雜的 TECN 層級覆蓋邏輯 (Level 1~6)
        /// </summary>
        private async Task<TecnPgmRecipeAttrDto> GetTecnPgmRecipeAttrAsync(
            string lotId, string ipn, string prodGroupKey, string prodCode, string stage, 
            string stepId, string stepName, string eqType2, string pgId, string pgName)
        {
            var result = new TecnPgmRecipeAttrDto();

            // 判斷測試模式
            string testMode = stepName;
            if (stepName.StartsWith("SORT"))
            {
                testMode = "S" + stepName.Replace("SORT", "");[cite: 4]
            }
            else if (stepName.StartsWith("TQAE"))
            {
                // 若為 TQAE 站，需呼叫 Get_TQAE_Mapping_Act_PGmode 來決定 testMode[cite: 4]
                // testMode = ... 
            }

            // 執行 Level 1 到 6 的迴圈尋找 PGM (略過 Level 5)[cite: 4]
            for (int level = 1; level <= 6; level++)
            {
                if (level == 5) continue;[cite: 4]

                string description = "";
                
                switch (level)
                {
                    case 1:
                        description = lotId;[cite: 4]
                        break;
                    case 2:
                    case 3:
                        description = ipn;[cite: 4]
                        break;
                    case 4:
                        description = prodGroupKey;[cite: 4]
                        break;
                    case 6:
                        description = prodCode;[cite: 4]
                        break;
                    default:
                        description = "";
                        break;
                }

                // TODO: 呼叫 Repository 執行對 TBL_TECN_PGM 的查詢 (傳入 Level, EqType2, Description)
                // var tecnRecord = await _repo.GetTecnPgmRecordAsync(level, description, testMode, ...);
                
                // 若找到資料且 ComparePgmAttr == "PASS"[cite: 4]
                // 且 Level 1 時，確認 TBL_LOT_TECN_CONTROL_LIST 的 Action 不為 Delete[cite: 4]
                // {
                //      若符合，則填入 result.RefPgm, result.RefTemp 等資訊並 break;
                // }
            }

            // TQAE站點如果查無資料，在原代碼中還會進行第二次轉換嘗試 (將 TF 轉為 TQ 等字串處理)[cite: 4]
            if (stepName.StartsWith("TQAE") && result.RefPgm != null)
            {
                if (result.RefPgm.StartsWith("TF"))
                    result.RefPgm = "TQ" + result.RefPgm.Substring(2);[cite: 4]
                else if (result.RefPgm.StartsWith("F"))
                    result.RefPgm = "Q" + result.RefPgm.Substring(1);[cite: 4]
            }

            return result;
        }

        /// <summary>
        /// 翻譯自: GetSwapPGName
        /// 判斷產品變更與取代規則
        /// </summary>
        private async Task<string> GetSwapPGNameAsync(string lotId, string prodGroup, string stepName, string eqType2, string originalPgName)
        {
            // 查詢 TBL_PROD_STEP_SPEC 取出該群組下的所有 SORT 站點，建立順序 (iCurStepIndex)[cite: 5]
            
            // 查詢 VIEW_WS_LOT_PGM_REC_SEQUENCE 取得該 Lot 的測試紀錄順序[cite: 5]
            
            // 查詢 TBL_WS_PGM_REPLACE 取出設定檔，逐一比對 Sequence 是否符合設定檔的前置站點組合[cite: 5]
            
            // 若找到符合的設定且 SwapFlag == "Y"[cite: 5]
            // 會回傳對應的取代程式名稱，若有衝突則回傳 "X" (代表需分批)[cite: 5]
            
            return originalPgName; // 實作細節需依賴 Dapper 多重查詢與迴圈封裝
        }
        /// <summary>
        /// 翻譯自: GetTecnPgmRecipeAttr
        /// </summary>
        public async Task<TecnPgmRecipeAttrDto> GetTecnPgmRecipeAttrAsync(
            string ipnTecnNo, string lotId, string ipn, string stepId, string stepName, 
            string eqType2, string prodGroupKey, string prodCode, string stage, string pgId, string pgName)
        {
            var result = new TecnPgmRecipeAttrDto();

            // 若參數不足，透過 LotId 補齊 IPN/ProdGroupKey/ProdCode
            if (!string.IsNullOrEmpty(lotId) && (string.IsNullOrEmpty(ipn) || string.IsNullOrEmpty(prodGroupKey) || string.IsNullOrEmpty(prodCode)))
            {
                // 🌟 這裡參數改傳 lotId
                var ipnData = await _repo.GetIpnMasterForTecnAsync(lotId); 
                
                if (ipnData != null)
                {
                    if (string.IsNullOrEmpty(ipn)) ipn = ipnData.IPN;
                    if (string.IsNullOrEmpty(prodGroupKey)) prodGroupKey = ipnData.PRODGROUPKEY;
                    if (string.IsNullOrEmpty(prodCode))
                    {
                        prodCode = stage == "WS" ? (ipn.Substring(0, 4) + ipnData.MASK_OPTION) : (ipn.Substring(0, 4) + ipnData.BE_OPTION);
                    }
                }
            }

            // 判斷 TestMode[cite: 4]
            string testMode = stepName;
            if (stepName.StartsWith("SORT")) testMode = "S" + stepName.Replace("SORT", "");
            // TQAE 的 mapping 邏輯可依據您的 Get_TQAE_Mapping_Act_PGmode 實作補充

            // 開始 1~6 階層掃描[cite: 4]
            for (int level = 1; level <= 6; level++)
            {
                if (level == 5) continue;[cite: 4]

                string description = level switch
                {
                    1 => lotId,
                    2 or 3 => ipn,
                    4 => prodGroupKey,
                    6 => prodCode,
                    _ => ""
                };[cite: 4]

                // 查詢資料庫[cite: 4]
                var records = await _repo.GetTecnPgmRecordsAsync(eqType2, testMode, level, description, "", ipnTecnNo);

                // 若為 TQAE 且無資料，原邏輯會嘗試再次用不同的條件查詢 (此處簡化示意)
                
                foreach (var record in records)
                {
                    // 若有 ComparePgmAttr 檢核邏輯，可以在此處呼叫
                    // if (ComparePgmAttr(...) == "PASS")
                    
                    result.RefPgmTecnNo = record.TECNNO;

                    // 檢查 Lot Tecn Control List 的 Delete 狀態 (僅 Level 1)[cite: 4]
                    if (level == 1)
                    {
                        string action = await _repo.GetTecnControlActionAsync(result.RefPgmTecnNo, lotId);
                        if (action == "Delete")
                        {
                            result.RefPgmTecnNo = "";
                            continue; // 跳過此筆，相當於 VB 的 GoTo PGMContinue[cite: 4]
                        }
                    }

                    result.RefPgId = record.PGID;
                    result.RefPgmSource = record.SOURCE;
                    result.RefPgm = record.PGNAME;
                    result.RefTemp = record.TEMP;

                    if (!string.IsNullOrEmpty(result.RefTemp))
                    {
                        result.RefTempTecnNo = record.TECNNO;
                        result.RefTempSource = record.SOURCE;
                    }

                    result.RefOverTime = record.OVERTIME;
                    result.RefLevel = record.TECNLEVEL?.ToString();
                    
                    break; // 找到就跳出迴圈
                }

                if (!string.IsNullOrEmpty(result.RefPgm)) break; // 找到 PGM 就可以結束 1~6 階層的掃描
            }

            // (選填) PGM 無變更 TEMP, 單獨取 TEMP 資訊的邏輯也可比照上述迴圈補上[cite: 4]

            return result;
        }

        /// <summary>
        /// 翻譯自: GetSwapPGName
        /// 利用 C# List/LINQ 處理複雜陣列比對
        /// </summary>
        public async Task<string> GetSwapPGNameAsync(string lotId, string prodGroup, string stepName, string eqType2, string originalPgName, bool bSpecQuery = false)
        {
            // 1. 取得該 ProdGroup 的 SORT 站點順序
            var specSteps = (await _repo.GetProdStepSpecStepsAsync(prodGroup)).ToList();
            int curStepIndex = specSteps.IndexOf(stepName) + 1; // 1-based
            if (curStepIndex <= 0) return ""; // 未實測此 Prodgroup 的站點

            // 2. 取得 LOT 實際測過的 Sequence[cite: 5]
            var lotSeqs = (await _repo.GetLotPgmRecSequenceAsync(lotId)).ToList();
            if (!lotSeqs.Any()) return "";

            // 3. 取得 Replace 設定檔[cite: 5]
            var replaceRules = (await _repo.GetWsPgmReplaceAsync(prodGroup)).ToList();

            var matchedPgmList = new List<string>();

            // 4. 開始比對 Sequence (對應 VB6 中極度複雜的字串 Split 與迴圈)[cite: 5]
            foreach (var lotSeq in lotSeqs)
            {
                if (bSpecQuery && lotSeqs.IndexOf(lotSeq) > 0) break;

                string[] vSteps = lotSeq.STEP?.ToString().Split(',') ?? new string[0];
                string[] vPgms = lotSeq.PGM?.ToString().Split(',') ?? new string[0];

                foreach (var rule in replaceRules)
                {
                    int stepMatchedCount = 0;
                    bool bMatch = true;

                    // 比對設定檔的 STEP1~15 是否與實際 LOT 測過的 vSteps 完全吻合[cite: 5]
                    for (int i = 0; i < vSteps.Length; i++)
                    {
                        stepMatchedCount++;
                        if (string.IsNullOrEmpty(vSteps[i]) || vSteps[i] == stepName) break;

                        // 取出 rule 中對應的 Step 和 Pgm (例如 STEP1, PGM1)
                        var dictRule = rule as IDictionary<string, object>;
                        string ruleStep = dictRule[$"STEP{i + 1}"]?.ToString();
                        string rulePgm = dictRule[$"PGM{i + 1}"]?.ToString();

                        if (vSteps[i] != ruleStep || vPgms[i] != rulePgm)
                        {
                            bMatch = false;
                            break;
                        }
                    }

                    if (bMatch && stepMatchedCount > 0)
                    {
                        var dictRule = rule as IDictionary<string, object>;
                        string tempSwapFlag = dictRule[$"SWAPFLAG{stepMatchedCount}"]?.ToString();
                        string tempPgName = dictRule[$"PGM{stepMatchedCount}"]?.ToString();

                        if (tempSwapFlag == "Y")
                        {
                            // 若 CheckSwapPGNameAttr 檢查過關，則加入結果[cite: 5]
                            // if (CheckSwapPGNameAttr(eqType2, tempPgName, originalPgName))
                            matchedPgmList.Add(tempPgName);
                        }
                    }
                }
            }

            // 5. 判斷是否有衝突[cite: 5]
            if (matchedPgmList.Distinct().Count() > 1)
            {
                return "X"; // 有衝突需分批[cite: 5]
            }

            if (matchedPgmList.Any())
            {
                return matchedPgmList.First();
            }

            return "";
        }
        public partial class PrintSetupFormService : IPrintSetupFormService
    {
public async Task<PrintSetupFormSubmitResponse> SubmitSetupFormAsync(PrintSetupFormSubmitRequest request)
        {
            var response = new PrintSetupFormSubmitResponse();
            bool isNewForm = request.Stage?.ToUpper() == "WS"; // 依據 CheckForm 邏輯，Stage=WS 為 mbNewForm = True

            // 1. 防呆驗證
            if (isNewForm)
            {
                if (string.IsNullOrEmpty(request.WaferId)) throw new Exception("WaferID 不可為空白。(WaferID can't Empty.)");
                if (string.IsNullOrEmpty(request.PgName)) throw new Exception("Program Name 不可為空白。(Program Name can't Empty.)");
                if (request.Stage != "FT" && string.IsNullOrEmpty(request.SetupReason)) throw new Exception("Setup Reason 不可為空白。");
            }
            if (request.Stage == "FT" && string.IsNullOrEmpty(request.TesterId))
            {
                throw new Exception("TesterId 不可沒選");
            }

            // 取得核心 Lot 物件
            var olot = WipServiceWrapper.Instance.LotById(request.LotId);
            if (olot == null) throw new Exception("此Lot不存在 !!");

            string ipn = olot.CustomAttributes(LotCustomAttributes.Ipn) ?? "";
            string prodGroup = olot.CustomAttributes(LotCustomAttributes.ProdGroup) ?? "";
            string stepName = olot.CurrentStep.Steps[0].Description ?? "";

            // 處理 SubSystem (若包含逗號，取前半部)
            string subSystem = request.SubSystemValue ?? "";
            if (subSystem.Contains(","))
            {
                subSystem = subSystem.Split(',')[0];[cite: 1]
            }

            // 取 EqType2 與 ErunTicNo[cite: 1]
            string eqType2 = await _repo.GetEqType2Async(request.TesterId) ?? "";
            var lotInfo = await _repo.GetLotInfoAsync(request.LotId);
            
            // 🌟 這裡使用大寫或小寫取決於您的 Dapper 對應，建議用 coalesce 寫法防呆
            string erunTicNo = lotInfo?.ERUNTICNO ?? lotInfo?.ErunTicNo ?? "";
            string sapRwNo = lotInfo?.SAPRWNO ?? lotInfo?.SapRwNo ?? "";

            // ====================================================================
            // 2. 停測檢核 (Stop Test Check)
            // ====================================================================
            // 呼叫剛剛寫好的 CheckNotStopTestAsync，bRunRule 帶入 false (同 VB6 邏輯)[cite: 1]
            var stopTestResult = await CheckNotStopTestAsync(
                lotId: request.LotId,
                ipn: ipn,
                prodGroup: prodGroup,
                stopScope: "ALL", // VB 原始碼使用 gsCAT_STOP_TEST_ALL[cite: 1]
                bRunRule: false,  // VB 原始碼傳入 False[cite: 1]
                stepName: stepName,
                curEqId: request.TesterId,
                eqType2: eqType2,
                curAccName: request.AccType,
                stopTicNo: request.StopTicNo,
                specSpecifyEq: request.SpecifyEq,
                specEqId: request.SpecifyEqId,
                erunTicNo: erunTicNo,
                pgId: request.PgId,
                pgName: request.PgName,
                pgMode: request.PgMode,
                temperature: request.Temp,
                subSystem: subSystem,
                checkWsDeviceFile: request.WsDeviceFile,
                specifyAcc: "" 
            );

            // 若 IsPass 為 False，代表有阻擋型的停測，直接拋出例外中斷流程[cite: 1]
            if (!stopTestResult.IsPass)
            {
                throw new Exception($"停測中，停測資訊:\n{stopTestResult.StopMessage}");
            }

            // 處理純提示型的停測字串 (對應 VB 的 lblStopInfo.Caption 處理)[cite: 1]
            string stopInfoMsg = stopTestResult.StopMessage;
            if (!string.IsNullOrEmpty(stopInfoMsg))
            {
                if (stopInfoMsg.StartsWith(",")) stopInfoMsg = stopInfoMsg.Substring(1);[cite: 1]
                response.StopInfoMsg = stopInfoMsg; // 傳給前端，讓前端 alert "請確認停測資訊:"[cite: 1]
            }
            else
            {
                response.StopInfoMsg = "NA";[cite: 1]
            }

            // ====================================================================
            // 3. 寫入資料庫 (對應 InsNewForm)
            // ====================================================================
            if (isNewForm)
            {
                // 1. 取得機台 ProberID
                var eqSys = await _repo.GetEqpSubSysAsync(request.TesterId);
                string proberId = "";
                if (eqSys != null)
                {
                    string sys1 = eqSys.SUBSYS1?.ToString() ?? "";
                    string sys2 = eqSys.SUBSYS2?.ToString() ?? "";
                    proberId = string.IsNullOrEmpty(sys2) ? sys1 : $"{sys1} {sys2}";
                }

                // 2. 取得預設的 Code, CheckSum, Speed
                string code = "", checkSum = "", speed = "";
                var ipnDetails = await _repo.GetIpnMasterDetailsAsync(ipn);
                if (ipnDetails != null)
                {
                    string fullCode = ipnDetails.CODE?.ToString() ?? "";
                    code = fullCode.Length > 2 ? fullCode.Substring(2, Math.Min(6, fullCode.Length - 2)) : fullCode; // VB: Mid(Code, 3, 6)
                    checkSum = ipnDetails.CHECKSUM?.ToString() ?? "";
                    speed = ipnDetails.SPEED?.ToString() ?? "";
                }

                // 檢查 MCD/MCP 覆寫
                if (await _repo.CheckIsMcdOrMcpAsync(ipn))
                {
                    code = "Check Code Server";
                    checkSum = "Check Code Server";
                }

                // ==========================================
                // 處理 INSTEPTIME (進站時間)
                // ==========================================
                string inStepTime = request.InStepTime;
                if (string.IsNullOrEmpty(inStepTime))
                {
                    // 若前端沒傳，從 WipServiceWrapper 取得目前的進站時間，或者押上當下時間
                    // 這裡先使用當下時間 (格式: YYYYMMDD HHMMSS)，您可依貴司底層設計調整
                    inStepTime = DateTime.Now.ToString("yyyyMMdd HHmmss"); 
                }

                // ==========================================
                // 取得 AssignProbeCard (判斷 ProbeCard 優先順序用)
                // ==========================================
                string assignProbeCard = lotInfo?.ASSIGNPROBECARD ?? "";

                // ==========================================
                // 3. 取得備註與配方參數 (TecnNo, ProbeCard, EqComments, StepComments)
                // ==========================================
                string eqComments = "", stepComments = "", pgmMatch = "", tecnNo = "", probeCard = "";

                if (!string.IsNullOrEmpty(tecnLotId))
                {
                    var tecnSpec = await _repo.GetLotStepEqSpecAsync(tecnLotId, request.StepNo, eqType2, subSystem, path);
                    if (tecnSpec != null)
                    {
                        eqComments = tecnSpec.Temperature; 
                        pgmMatch = tecnSpec.ReplacePgName01; 
                        
                        // ✅ 正確賦值 TecnNo
                        tecnNo = tecnSpec.TecnNo; 
                        
                        // ✅ 正確賦值 StepComments
                        stepComments = tecnSpec.StepComments; 
                        
                        // ✅ 決定 ProbeCard：若 Lot 有強制指定 (AssignProbeCard)，則優先使用，否則使用 SPEC 設定
                        probeCard = string.IsNullOrEmpty(assignProbeCard) ? tecnSpec.ProbeCardType : assignProbeCard; 
                    }
                }
                else
                {
                    var prodSpec = await _repo.GetProdStepEqSpecAsync(prodGroup, request.StepNo, eqType2, subSystem, path);
                    if (prodSpec != null)
                    {
                        stepComments = prodSpec.StepComments;
                        probeCard = string.IsNullOrEmpty(assignProbeCard) ? prodSpec.ProbeCardType : assignProbeCard; 
                        
                        // ✅ 新增：若 DocType 不為空，則將 TecnNo 賦值為 DocNo
                        if (!string.IsNullOrEmpty(prodSpec.DocType))
                        {
                            tecnNo = prodSpec.DocNo;
                        }
                    }
                }

                // 4. FollowProduct = 'N' 工程品強制覆寫邏輯 (維持原樣)
                string followProdCheck = string.IsNullOrEmpty(erunTicNo) ? "" : await _repo.GetFollowProductAsync(request.LotId, erunTicNo, request.Stage);
                if (!string.IsNullOrEmpty(erunTicNo) && followProdCheck == "N")
                {
                    pgmMatch = ""; tecnNo = ""; eqComments = ""; stepComments = ""; probeCard = "";
                    
                    var erunOverride = await _repo.GetErunOverrideDataAsync(request.LotId, erunTicNo, request.Stage, request.StepNo, eqType2, subSystem);
                    if (erunOverride != null)
                    {
                        code = erunOverride.CODE?.ToString() ?? "";
                        checkSum = erunOverride.CHECKSUM?.ToString() ?? "";
                        speed = erunOverride.SPEED?.ToString() ?? "";
                        // 工程品直接覆寫 ProbeCard
                        probeCard = erunOverride.PROBECARDTYPE?.ToString() ?? "";
                    }
                }

                // 5. 進行 Insert (包含 Retry 機制)
                string formNo = "";
                bool insertSuccess = false;

                for (int i = 0; i < 20; i++)
                {
                    formNo = await _repo.GetNextSetupFormNoAsync();
                    if (string.IsNullOrEmpty(formNo)) formNo = $"ST{DateTime.Now:yyyyMMdd}001";

                    try
                    {
                        var insertParam = new
                        {
                            FormNo = formNo,
                            Ipn = ipn,
                            LotId = request.LotId,
                            LotOwner = olot.CustomAttributes(LotCustomAttributes.LotOwner),
                            ErunTicNo = erunTicNo,
                            WaferId = request.WaferId,
                            TesterId = request.TesterId,
                            ProberId = proberId,       // ✅ 已補齊
                            TestMode = stepName,
                            PgName = request.PgName,
                            PgId = request.PgId,
                            PgmMatch = pgmMatch,       // ✅ 已補齊
                            Code = code,               // ✅ 已補齊
                            CheckSum = checkSum,       // ✅ 已補齊
                            Speed = speed,             // ✅ 已補齊
                            Temp = request.Temp,
                            WsDeviceFile = request.WsDeviceFile,
                            TecnNo = tecnNo,           // ✅ 已補齊
                            ProbeCard = probeCard,     // ✅ 已補齊
                            EqComments = eqComments,   // ✅ 已補齊
                            StepComments = stepComments,// ✅ 已補齊
                            CreateUserId = request.UserId,
                            SetupReason = request.SetupReason,
                            SetupWaferId1 = request.SetupWaferId1,
                            SetupWaferId2 = request.SetupWaferId2,
                            SetupWaferId3 = request.SetupWaferId3,
                            StepId = request.StepNo,
                            InStepTime = request.InStepTime,
                            StopInfo = response.StopInfoMsg
                        };

                        await _repo.InsertWsEqFormBasicAsync(insertParam);
                        insertSuccess = true;
                        break; 
                    }
                    catch (Exception)
                    {
                        await Task.Delay(1000);
                    }
                }

                if (!insertSuccess)
                {
                    throw new Exception($"SETUP FORM 單號寫入失敗,請洽IT人員協助.");
                }

                await _repo.UpdateEqInfoFormNoAsync(request.TesterId, formNo, request.UserId);
                response.FormNo = formNo;
            }
            else if (request.Stage == "FT")
            {
                // 1. 蒐集 Excel 所需的所有資料
                var excelDto = await GatherFtSetupExcelDataAsync(request, response.StopInfoMsg, eqType2, erunTicNo, sapRwNo, tecnLotId, prodGroup, path);
                
                // 2. 判斷使用哪個 Sheet
                if (eqType2 == "AT3-300AL" || eqType2 == "FT-940")
                {
                    excelDto.SheetName = eqType2;
                }
                else
                {
                    excelDto.SheetName = "FT";
                }

                // 3. 取得模板實體路徑 (可依據貴司 Web API 設定調整路徑取得方式)
                string templatePath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Templates", "SetupForm.xlsx");

                // 4. 呼叫我們剛寫好的方法產出 byte[]
                byte[] fileBytes = GenerateFtSetupExcelBytes(excelDto, templatePath);

                // 5. (選項A) 存成實體檔案並回傳下載 URL
                // string fileName = $"SetupForm_{request.LotId}_{DateTime.Now:yyyyMMddHHmmss}.xlsx";
                // string savePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Downloads", fileName);
                // File.WriteAllBytes(savePath, fileBytes);
                // response.FileDownloadUrl = $"/Downloads/{fileName}";

                // 5. (選項B) 直接將 Base64 字串塞入 Response，讓前端 Vue 解析下載 (推薦此做法，不佔 Server 空間)
                // response.ExcelBase64 = Convert.ToBase64String(fileBytes);
                // response.FileName = $"SetupForm_{request.LotId}.xlsx";
            }

            return response;
        }
        //------
        
        // ==============================================================================
        // 翻譯自 NotStopTest
        // ==============================================================================
        public async Task<(bool IsPass, string StopMessage)> CheckNotStopTestAsync(
            string lotId, string ipn, string prodGroup, string stopScope, bool bRunRule, 
            string stepName, string curEqId, string eqType2, string curAccName, string stopTicNo, 
            string specSpecifyEq, string specEqId, string erunTicNo, string pgId, string pgName, 
            string pgMode, string temperature, string subSystem, string checkWsDeviceFile, string specifyAcc)
        {
            string sWsDeviceFile = checkWsDeviceFile;
            string sCurAccNo = "";
            string sLotContinueTest = "";
            string sFollowProd = "";
            string sContinueTest = "";

            var olot = WipServiceWrapper.Instance.LotById(lotId);
            string stepId = olot.CurrentStep.Steps[0].Id;
            string stage = olot.CustomAttributes(LotCustomAttributes.Stage);
            string path = olot.CustomAttributes(LotCustomAttributes.Route);

            // 取 TECN LotInfo 與 ContinueTest[cite: 10]
            var lotInfo = await _repo.GetLotInfoAsync(lotId);
            string tecnLotId = lotInfo?.TECNLOTID ?? "";
            
            if (!string.IsNullOrEmpty(tecnLotId))
            {
                sLotContinueTest = await _repo.GetLotStepEqSpecContinueTestAsync(tecnLotId, stepId, path, eqType2, subSystem) ?? "";
            }

            // 處理 bRunRule 與機台配件[cite: 10]
            if (bRunRule)
            {
                var eqAcc = await _repo.GetEqInfoAccessoriesAsync(curEqId);
                if (eqAcc != null)
                {
                    sWsDeviceFile = eqAcc.WsDeviceFile?.Trim() ?? "";
                    
                    var accNos = new List<string>();
                    if (!string.IsNullOrEmpty(eqAcc.ProbeCardId)) { curAccName = eqAcc.ProbeCardId.Split('-')[0]; accNos.Add(eqAcc.ProbeCardId); }
                    if (!string.IsNullOrEmpty(eqAcc.LoadBoardId)) { curAccName = eqAcc.LoadBoardId.Split('-')[0]; accNos.Add(eqAcc.LoadBoardId); }
                    if (!string.IsNullOrEmpty(eqAcc.ContactBoardId)) { curAccName = eqAcc.ContactBoardId.Split('-')[0]; accNos.Add(eqAcc.ContactBoardId); }
                    sCurAccNo = string.Join(",", accNos);
                }
            }

            if (!string.IsNullOrEmpty(specifyAcc))
            {
                sCurAccNo = specifyAcc;
                curAccName = specifyAcc.Contains("-") ? specifyAcc.Split('-')[0] : specifyAcc;
            }

            if (!string.IsNullOrEmpty(erunTicNo))
            {
                var erunData = await _repo.GetErunReqContinueTestAsync(lotId, erunTicNo, stage);
                if (erunData != null)
                {
                    sFollowProd = erunData.FollowProd ?? "";
                    sContinueTest = erunData.ContinueTest ?? "";
                }
            }

            var stopTicNoList = new HashSet<string>();
            if (!string.IsNullOrEmpty(stopTicNo)) stopTicNoList.Add(stopTicNo);

            // 根據 StopScope 分流處理[cite: 10]
            if (stopScope == "ALL" || stopScope == "P")
            {
                if (sLotContinueTest != "Y" && sContinueTest != "Y")
                {
                    if (sFollowProd != "N") 
                    {
                        // 檢查 ChkProdSpecEqID[cite: 10]
                        if (!ChkProdSpecEqID(specSpecifyEq, specEqId, curEqId))
                        {
                            stopTicNoList.Add(stopTicNo);
                        }

                        // 呼叫 GetProductStopTestAsync[cite: 10]
                        var prodStopResult = await GetProductStopTestAsync(prodGroup, eqType2, pgId, pgName, stepName, curAccName, lotId, sCurAccNo, sWsDeviceFile, curEqId, ipn);
                        foreach (var tic in prodStopResult.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries))
                        {
                            stopTicNoList.Add(tic);
                        }
                    }
                }
            }

            if (stopScope == "ALL" || stopScope == "NP")
            {
                var nonProdStopResult = await GetNonProdStopTestAsync(ipn, eqType2, curEqId, curAccName, sCurAccNo, temperature);
                foreach (var tic in nonProdStopResult.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries))
                {
                    stopTicNoList.Add(tic);
                }
            }

            // 整理最終的停測字串[cite: 10]
            string finalStopTicNos = string.Join(",", stopTicNoList.OrderBy(x => x));
            return (finalStopTicNos == "", finalStopTicNos);
        }

        // ==============================================================================
        // 翻譯自 ChkProdSpecEqID[cite: 10]
        // ==============================================================================
        private bool ChkProdSpecEqID(string specifyEq, string prodSpecEqId, string curEqId)
        {
            if (string.IsNullOrEmpty(specifyEq)) return true;

            var eqArray = (prodSpecEqId ?? "").Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
            bool bExist = eqArray.Any(eq => eq.Trim() == curEqId);

            if (specifyEq == "Y" && bExist) return true;
            if (specifyEq == "N" && !bExist) return true;

            return false;
        }

        // ==============================================================================
        // 翻譯自 GetProductStopTest[cite: 10]
        // 使用 LINQ Dictionary 取代複雜的陣列操作
        // ==============================================================================
        private async Task<string> GetProductStopTestAsync(
            string prodGroup, string eqType2, string pgId, string pgName, string stepName, 
            string concatAcc, string lotId, string curAccNo, string wsDeviceFile, string curEqId, string ipn)
        {
            // 處理 lotId 虛擬批號[cite: 10]
            if (lotId.Contains(".")) lotId = lotId.Substring(0, lotId.IndexOf('.'));

            var records = (await _repo.GetProductStopTestRecordsAsync(eqType2, prodGroup, ipn, pgId, pgName, stepName, concatAcc, lotId, curAccNo, wsDeviceFile)).ToList();
            if (!records.Any()) return "";

            var stopTicNos = new HashSet<string>();
            var dicTicNoVsEQ = new Dictionary<string, HashSet<string>>();

            // 建立 Dictionary 來記錄每個單號(區分 SpecifyEq) 下包含的機台清單[cite: 10]
            if (!string.IsNullOrEmpty(curEqId))
            {
                foreach (var rec in records)
                {
                    string key = $"{rec.StopTicNo}/{rec.SpecifyEq}";
                    if (!dicTicNoVsEQ.ContainsKey(key)) dicTicNoVsEQ[key] = new HashSet<string>();
                    
                    if (!string.IsNullOrEmpty(rec.EqId)) dicTicNoVsEQ[key].Add(rec.EqId);
                }
            }

            foreach (var rec in records)
            {
                stopTicNos.Add(rec.StopTicNo);
            }

            // 同一停測單, SpecifyEq有Y也有N邏輯的移除判斷[cite: 10]
            if (!string.IsNullOrEmpty(curEqId))
            {
                var ticsToRemove = new List<string>();
                foreach (var ticNo in stopTicNos)
                {
                    bool bRemoveY = false;
                    bool bRemoveN = false;

                    string keyY = $"{ticNo}/Y";
                    if (dicTicNoVsEQ.ContainsKey(keyY) && dicTicNoVsEQ[keyY].Any())
                    {
                        if (dicTicNoVsEQ[keyY].Contains(curEqId)) bRemoveY = true;
                    }

                    string keyN = $"{ticNo}/N";
                    if (dicTicNoVsEQ.ContainsKey(keyN) && dicTicNoVsEQ[keyN].Any())
                    {
                        if (!dicTicNoVsEQ[keyN].Contains(curEqId)) bRemoveN = true;
                    }

                    if (dicTicNoVsEQ.ContainsKey(keyY) && dicTicNoVsEQ.ContainsKey(keyN))
                    {
                        if (bRemoveY && bRemoveN) ticsToRemove.Add(ticNo);
                    }
                    else
                    {
                        if (bRemoveY || bRemoveN) ticsToRemove.Add(ticNo);
                    }
                }

                foreach (var tic in ticsToRemove)
                {
                    stopTicNos.Remove(tic);
                }
            }

            return string.Join(",", stopTicNos.OrderBy(x => x)); // 對應 SortString[cite: 10]
        }

        // ==============================================================================
        // 翻譯自 GetNonProdStopTest[cite: 10]
        // ==============================================================================
        private async Task<string> GetNonProdStopTestAsync(
            string ipn, string eqType2, string curEqId, string curAccName, string curAccNo, string temperature)
        {
            int tempValue = int.TryParse(temperature, out int parsed) ? parsed : 25; // 假設常溫 25[cite: 10]

            var ipnData = await _repo.GetIpnMasterForTecnAsync(ipn); // 重用 IpnMaster 查詢
            string packageType = ipnData?.PACKAGE_CODE ?? "";
            string bodySize = ipnData?.BODY_SIZE ?? "";
            string pinCount = ipnData?.PIN_COUNT ?? "0";
            string carrierType = ipnData?.CARRIER_TYPE ?? "";
            string formFactorName = ipnData?.FORMFACTORNAME ?? "";
            string moduleOption = ipnData?.MODULEOPTION ?? "";

            var records = await _repo.GetNonProdStopTestRecordsAsync(
                eqType2, curEqId, curAccName, packageType, pinCount, carrierType, bodySize, tempValue, curAccNo, formFactorName, moduleOption);

            // HashSet 保證不重複並使用 OrderBy 排序[cite: 10]
            var stopTicNos = new HashSet<string>(records);
            return string.Join(",", stopTicNos.OrderBy(x => x));
        }

        private async Task<FtSetupExcelDto> GatherFtSetupExcelDataAsync(
            PrintSetupFormSubmitRequest request, string stopInfoMsg, string eqType2, 
            string erunTicNo, string sapRwNo, string tecnLotId, string prodGroup, string path)
        {
            var dto = new FtSetupExcelDto();
            var olot = WipServiceWrapper.Instance.LotById(request.LotId);

            // ==========================================
            // 1. 基本資訊與標頭 (Header)
            // ==========================================
            dto.OperatorId = request.UserId;
            dto.ShiftCode = await _repo.GetShiftCodeAsync(request.UserId) ?? "";
            dto.PrintDate = DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss");
            dto.StopInfo = stopInfoMsg == "NA" ? "" : stopInfoMsg;
            dto.LotId = request.LotId;
            dto.Ipn = olot.CustomAttributes(LotCustomAttributes.Ipn) ?? "";
            dto.LotOwner = olot.CustomAttributes(LotCustomAttributes.LotOwner) ?? "";
            dto.StepName = olot.CurrentStep.Steps[0].Description ?? "";
            dto.TesterId = request.TesterId;

            // ==========================================
            // 🌟 完美還原 VB6 的 ERUNTICNO 與 SAPRWNO 拼接邏輯
            // ==========================================
            if (string.IsNullOrEmpty(erunTicNo))
            {
                dto.ErunTicNo = sapRwNo;
            }
            else if (string.IsNullOrEmpty(sapRwNo))
            {
                dto.ErunTicNo = erunTicNo;
            }
            else
            {
                dto.ErunTicNo = $"{erunTicNo}, {sapRwNo}";
            }

            // ==========================================
            // 2. 產品/封裝資訊 (IPN Master)
            // ==========================================
            var ipnDetails = await _repo.GetIpnMasterForTecnAsync(request.LotId); // 重用我們先前修復的含 Lot 關聯的查詢
            if (ipnDetails != null)
            {
                dto.PackageName = ipnDetails.PACKAGENAME?.ToString() ?? "";
                string fullCode = ipnDetails.CODE?.ToString() ?? "";
                dto.Code = fullCode.Length > 2 ? fullCode.Substring(2) : fullCode;
                dto.CheckSum = ipnDetails.CHECKSUM?.ToString() ?? "";
                dto.Speed = ipnDetails.SPEED?.ToString() ?? "";
                dto.PinCount = ipnDetails.PINCOUNT?.ToString() ?? "";
                dto.PackageCode = ipnDetails.PACKAGECODE?.ToString() ?? "";
                dto.PbFree = ipnDetails.GPTYPE?.ToString() == "XX" ? "N" : "Y";
                
                // 保留 BodySize 給後續查詢用
                string bodySize = ipnDetails.BODYSIZE?.ToString() ?? "";
            }

            // ==========================================
            // 3. 配方與機台設定 (Spec)
            // ==========================================
            string subSystemType = request.SubSystemValue?.Split(',')[0] ?? "";
            string maxSite = request.SubSystemValue?.Contains(",") == true ? request.SubSystemValue.Split(',')[1] : "";

            RecipeSpecData spec = null;
            if (!string.IsNullOrEmpty(tecnLotId))
            {
                spec = await _repo.GetLotStepEqSpecAsync(tecnLotId, request.StepNo, eqType2, subSystemType, path, maxSite);
            }
            else
            {
                spec = await _repo.GetProdStepEqSpecAsync(prodGroup, request.StepNo, eqType2, subSystemType, path, maxSite);
            }

            if (spec != null)
            {
                dto.Temperature = !string.IsNullOrEmpty(request.Temp) ? request.Temp : spec.Temperature;
                dto.LoadBoard = spec.LoadboardType;
                dto.ContactBoard = spec.ContactboardType;
                
                // 若為空，呼叫 GetFTAccByPgmAsync 替補
                if (string.IsNullOrEmpty(dto.LoadBoard) && string.IsNullOrEmpty(dto.ContactBoard))
                {
                    var ftAcc = await _repo.GetFTAccByPgmAsync(dto.StepName, ipnDetails?.BODYSIZE?.ToString() ?? "", spec.PgId, spec.PgName);
                    if (ftAcc != null)
                    {
                        dto.LoadBoard = ftAcc.Loadboard;
                        dto.ContactBoard = ftAcc.Contactboard;
                    }
                }

                // 處理 TECN PGM NO 與 PGM 名稱轉換
                if (!string.IsNullOrEmpty(spec.DocType))
                {
                    // 這裡可以套用 msPGMTECNNo + DocNo 的邏輯
                    string msPGMTECNNo = ""; // 若前方有取得
                    dto.ErunTicNo = string.IsNullOrEmpty(msPGMTECNNo) ? spec.DocNo : $"{msPGMTECNNo}+{spec.DocNo}";
                }

                // 呼叫 GetPGMAsync 取得最終 PgmName
                dto.PgmName = await GetPGMAsync(request.LotId, request.StepNo, path, eqType2, subSystemType, spec.PgName, spec);
                
                // 其他配件
                dto.Cable = spec.CableType;  // 需擴充 DTO/Repository
                dto.KitType = spec.KitType;  // 需擴充 DTO/Repository
                dto.NeedJumper = spec.NeedJumper; // 需擴充 DTO/Repository
                dto.JumperPinNo = spec.JumperPinNo; // 需擴充 DTO/Repository
                
                // 呼叫 GetAccMatchPlate
                dto.MatchPlate = await _repo.GetAccMatchPlateAsync(dto.Ipn, dto.KitType, dto.LoadBoard, eqType2);
            }

            // ==========================================
            // 4. 機台 Subsystem & 額外設定
            // ==========================================
            var eqFullInfo = await _repo.GetEqInfoAccessoriesFullAsync(request.TesterId);
            if (eqFullInfo != null)
            {
                dto.SubSystem1 = eqFullInfo.SUBSYS1?.ToString() ?? "";
                dto.SubSystem2 = eqFullInfo.SUBSYS2?.ToString() ?? "";
                
                // 🌟 AT3 專屬配件擷取 (處理逗號分隔，只取前半段)
                string contactBoardId = eqFullInfo.CONTACTBOARDID?.ToString() ?? "";
                dto.EqContactBoardId = contactBoardId.Contains(",") ? contactBoardId.Split(',')[0] : contactBoardId;

                string vacuumCupZ1 = eqFullInfo.VACUUMCUPZ1?.ToString() ?? "";
                dto.EqVacuumCupZ1 = vacuumCupZ1.Contains(",") ? vacuumCupZ1.Split(',')[0] : vacuumCupZ1;
            }

            // Soak Time
            if (dto.Temperature == "25" || dto.Temperature == "30" || dto.Temperature.ToUpper().Contains("ROOM"))
            {
                dto.SoakTime = "0";
            }
            else
            {
                dto.SoakTime = await _repo.GetSoakTimeAsync(eqType2, dto.PinCount, dto.PackageCode);
            }

            // Sthand File
            dto.WsDeviceFile = await _repo.GetSthandFileAsync(dto.LoadBoard);

            // Comments
            dto.Comments = await _repo.GetSpecCommentsAsync(!string.IsNullOrEmpty(tecnLotId), !string.IsNullOrEmpty(tecnLotId) ? tecnLotId : prodGroup, request.StepNo, path);

            // Pitch & VacuumCup
            var fviAcc = await _repo.GetTesterFviAccAsync(dto.PinCount, dto.PackageCode);
            if (fviAcc != null)
            {
                dto.Pitch = fviAcc.PITCH?.ToString() ?? "";
                dto.VacuumCup = fviAcc.VACUUMCUP?.ToString() ?? "";
            }

            // ==========================================
            // 5. FollowProduct = 'N' 的工程品覆蓋邏輯
            // ==========================================
            string followProdCheck = string.IsNullOrEmpty(erunTicNo) ? "" : await _repo.GetFollowProductAsync(request.LotId, erunTicNo, request.Stage);
            if (!string.IsNullOrEmpty(erunTicNo) && followProdCheck == "N")
            {
                var erunOverride = await _repo.GetErunOverrideDataAsync(request.LotId, erunTicNo, request.Stage, request.StepNo, eqType2, subSystemType);
                if (erunOverride != null)
                {
                    dto.Speed = erunOverride.SPEED?.ToString() ?? dto.Speed;
                    dto.Code = erunOverride.CODE?.ToString() ?? dto.Code;
                    dto.CheckSum = erunOverride.CHECKSUM?.ToString() ?? dto.CheckSum;
                    dto.PgmName = erunOverride.PGNAME?.ToString() ?? dto.PgmName;
                    dto.Temperature = erunOverride.TEMPERATURE?.ToString() ?? dto.Temperature;
                    dto.LoadBoard = erunOverride.LOADBOARDTYPE?.ToString() ?? dto.LoadBoard;
                    dto.ContactBoard = erunOverride.CONTACTBOARDTYPE?.ToString() ?? dto.ContactBoard;
                    dto.JumperPinNo = erunOverride.JUMPERPINNO?.ToString() ?? dto.JumperPinNo; // 需擴充 Repository
                    dto.GoodBin = erunOverride.GOODBIN?.ToString() ?? dto.GoodBin;
                }
            }

            // ==========================================
            // 6. Recipe Spec List (HW Setup Quality check)
            // ==========================================
            var recipeSpecs = await _repo.GetRecipeSpecNamesAsync(request.TesterId, request.LotId, dto.Ipn, dto.PackageName, dto.PinCount, ipnDetails?.BODYSIZE?.ToString() ?? "");
            dto.RecipeList = recipeSpecs.ToList();

            return dto;
        }
        // ==============================================================================
        // 專責處理 FT 站架機單 Excel 匯出 (使用 ClosedXML)
        // ==============================================================================
        public byte[] GenerateFtSetupExcelBytes(FtSetupExcelDto dto, string templatePath)
        {
            if (!File.Exists(templatePath))
            {
                throw new FileNotFoundException($"找不到 Excel 模板檔案: {templatePath}");
            }

            using (var wb = new XLWorkbook(templatePath))
            {
                // 1. 刪除不需要的 Sheet，只保留目標 Sheet (例如 "FT" 或 "AT3-300AL")
                foreach (var sheet in wb.Worksheets.ToList())
                {
                    if (sheet.Name != dto.SheetName)
                    {
                        sheet.Delete();
                    }
                }

                // 2. 取得目標 Sheet
                var ws = wb.Worksheet(dto.SheetName);

                // ==========================================
                // 3. 填寫基本標頭資訊 (依據我們先前盤點的 Cell 位置)
                // ==========================================
                if (dto.SheetName == "AT3-300AL" || dto.SheetName == "FT-940")
                {
                    // === AT3 專屬排版 (根據 VB6 換算) ===
                    ws.Cell("C2").Value = dto.OperatorId;
                    ws.Cell("H2").Value = dto.ShiftCode;
                    ws.Cell("R2").Value = dto.PrintDate;
                    ws.Cell("G4").Value = dto.StopInfo;
                    
                    ws.Cell("E5").Value = dto.LotId;
                    ws.Cell("R5").Value = dto.LotOwner;
                    ws.Cell("E6").Value = dto.Ipn;
                    ws.Cell("R6").Value = dto.ErunTicNo;
                    ws.Cell("E7").Value = dto.StepName;
                    ws.Cell("R7").Value = dto.TesterId;

                    ws.Cell("E8").Value = dto.PackageName;
                    ws.Cell("R8").Value = dto.Temperature;
                    ws.Cell("E9").Value = dto.PgmName;
                    ws.Cell("R9").Value = dto.Code;
                    ws.Cell("R10").Value = dto.CheckSum;
                    ws.Cell("R11").Value = dto.GoodBin; // Erun 專屬
                    
                    ws.Cell("E12").Value = dto.ErunTicNo.Contains("+") ? dto.ErunTicNo : ""; // 簡化 TECN NO 邏輯
                    ws.Cell("E15").Value = dto.Comments; // iPosTitle(0) + 13 + iPosDown(2) = 15

                    // AT3 配件區 (刻意忽略 LoadBoard, Jumper, Subsystem)
                    ws.Cell("E20").Value = dto.EqContactBoardId;
                    ws.Cell("E21").Value = dto.EqVacuumCupZ1;
                    ws.Cell("E27").Value = dto.ContactBoard; // iPosTitle(0) + 25 + iPosDown(2) = 27
                    ws.Cell("E28").Value = dto.VacuumCup;

                    // Recipe Spec List (AT3 從 38 列開始)
                    int startRowAt3 = 38;
                    for (int i = 0; i < dto.RecipeList.Count; i++)
                    {
                        int currentRow = startRowAt3 + (i / 2);
                        if (i % 2 == 0)
                        {
                            ws.Cell($"A{currentRow}").Value = dto.RecipeList[i].Id;
                            ws.Cell($"B{currentRow}").Value = dto.RecipeList[i].SpecName;
                        }
                        else
                        {
                            ws.Cell($"P{currentRow}").Value = dto.RecipeList[i].Id;
                            ws.Cell($"Q{currentRow}").Value = dto.RecipeList[i].SpecName;
                        }
                    }
                }
                else
                {
                    ws.Cell("C3").Value = dto.OperatorId;
                    ws.Cell("H3").Value = dto.ShiftCode;
                    ws.Cell("R3").Value = dto.PrintDate;
                    
                    // 停測提示 (通常是合併儲存格，直接對左上角賦值即可)
                    ws.Cell("F5").Value = dto.StopInfo; 
    
                    // ==========================================
                    // 4. 填寫 Lot 與產品資訊
                    // ==========================================
                    ws.Cell("E6").Value = dto.LotId;
                    ws.Cell("R6").Value = dto.LotOwner;
                    ws.Cell("E7").Value = dto.Ipn;
                    ws.Cell("R7").Value = dto.ErunTicNo;
                    ws.Cell("E8").Value = dto.StepName;
                    ws.Cell("R8").Value = dto.TesterId;
    
                    ws.Cell("E9").Value = dto.PackageName;
                    ws.Cell("E10").Value = dto.Speed;
                    ws.Cell("R10").Value = dto.Code;
                    ws.Cell("E11").Value = dto.PgmName;
                    ws.Cell("R11").Value = dto.CheckSum;
                    ws.Cell("E12").Value = dto.WsDeviceFile;
                    ws.Cell("R12").Value = dto.Temperature;
                    ws.Cell("R13").Value = dto.SoakTime;
    
                    // ==========================================
                    // 5. 填寫 Jumper 與備註 (需依實際版面微調)
                    // ==========================================
                    ws.Cell("E16").Value = dto.NeedJumper;
                    ws.Cell("E17").Value = dto.JumperPinNo;
                    // 若 Comments 要放在旁邊，可指定例如 H16 或與 Jumper 共用
                    ws.Cell("R16").Value = dto.Comments; 
    
                    // ==========================================
                    // 6. 填寫機台配件 (Accessories)
                    // ==========================================
                    ws.Cell("E20").Value = dto.SubSystem1;
                    ws.Cell("R20").Value = dto.SubSystem2;
    
                    ws.Cell("E28").Value = dto.LoadBoard;
                    ws.Cell("R28").Value = dto.LoadBoard; // 依 VB 邏輯可能左右都有
                    ws.Cell("E29").Value = dto.Cable;
                    ws.Cell("R29").Value = dto.Cable;
                    ws.Cell("E30").Value = dto.ContactBoard;
                    ws.Cell("R30").Value = dto.ContactBoard;
                    ws.Cell("E31").Value = dto.KitType;
                    ws.Cell("R31").Value = dto.KitType;
                    ws.Cell("E33").Value = dto.MatchPlate;
                    ws.Cell("R33").Value = dto.MatchPlate;
                    
                    ws.Cell("E35").Value = dto.Pitch;
                    ws.Cell("E36").Value = dto.VacuumCup;
    
                    // PbFree 打勾邏輯 (假設 E36/G36 旁有對應的框)
                    if (dto.PbFree == "Y")
                        ws.Cell("H36").Value = "V"; // 視模板勾選框位置調整
                    else if (dto.PbFree == "N")
                        ws.Cell("J36").Value = "V";
    
                    // ==========================================
                    // 7. 填寫 Recipe Spec List (動態產生列)
                    // ==========================================
                    // 假設 Recipe 清單從第 49 列開始，分左右兩欄 (A/B 與 P/Q)
                    int startRow = 49;
                    for (int i = 0; i < dto.RecipeList.Count; i++)
                    {
                        int currentRow = startRow + (i / 2); // 每兩筆換一列
                        
                        if (i % 2 == 0)
                        {
                            // 填在左半邊 (A, B)
                            ws.Cell($"A{currentRow}").Value = dto.RecipeList[i].Id;
                            ws.Cell($"B{currentRow}").Value = dto.RecipeList[i].SpecName;
                        }
                        else
                        {
                            // 填在右半邊 (P, Q)
                            ws.Cell($"P{currentRow}").Value = dto.RecipeList[i].Id;
                            ws.Cell($"Q{currentRow}").Value = dto.RecipeList[i].SpecName;
                        }
                    }
                }
                // ==========================================
                // 8. 儲存為 MemoryStream 並回傳 Byte Array
                // ==========================================
                using (var ms = new MemoryStream())
                {
                    wb.SaveAs(ms);
                    return ms.ToArray();
                }
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

        public partial class PrintSetupFormController : ApiController
    {
        /// <summary>
        /// 處理 Click OK 提交動作 (包含驗證、停測檢查與資料庫單號寫入)
        /// </summary>
        [HttpPost, Route("submit"), AuthorizeToken]
        public async Task<IHttpActionResult> SubmitSetupForm([FromBody] PrintSetupFormSubmitRequest request)
        {
            try
            {
                if (request == null || string.IsNullOrWhiteSpace(request.LotId))
                {
                    return Ok(new { Success = false, Message = "無效的請求參數!" });
                }

                var data = await _service.SubmitSetupFormAsync(request);
                return Ok(new { Success = true, Data = data });
            }
            catch (Exception ex)
            {
                // 回傳錯誤訊息讓前端 VMessageBox 顯示
                return Ok(new { Success = false, Message = ex.Message });
            }
        }
    }
    }
}

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
            string sql = @"
                SELECT a.SPEED, a.CODE, a.CHECKSUM, 
                       b.PGNAME, b.TEMPERATURE, b.WSDEVICEFILE, b.PROBECARDTYPE, b.LOADBOARDTYPE, b.PGID 
                FROM TBL_ERUN_REQ a
                LEFT JOIN TBL_ERUN_RECIPE b ON b.DOCNO = a.TICKET_NO AND b.STEPNO = :StepNo AND b.EQTYPE2 = :EqType2 AND b.DELETEFLAG = 'N' 
                     AND NVL(b.SUBSYSTEM, ' ') = NVL(:SubSystem, ' ')
                WHERE a.LOT_ID = :LotId AND a.TICKET_NO = :TicNo AND a.STAGE = :Stage";
            return await _dbConnection.QueryFirstOrDefaultAsync(sql, new { LotId = lotId, TicNo = erunTicNo, Stage = stage, StepNo = stepNo, EqType2 = eqType2, SubSystem = string.IsNullOrEmpty(subSystem) ? " " : subSystem });
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

            // (可選) 呼叫 modTecn.GetTecnPgmRecipeAttr 進行最後 TECN 屬性的覆蓋
            // CustomMesHelper.GetTecnPgmRecipeAttr(..., out getPgm, out getPgId, out getTemp);
            // if (!string.IsNullOrEmpty(getPgm)) response.PgName = getPgm;
            // ...

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
            string erunTicNo = lotInfo?.ERUNTICNO ?? "";

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
            // 3. 寫入資料庫 (對應 InsNewForm)[cite: 1]
            // ====================================================================
            if (isNewForm)
            {
                string formNo = "";
                bool insertSuccess = false;

                // Retry 邏輯: 最多 20 次[cite: 1]
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
                            ProberId = "", // 可依據 TesterId 查詢 EQ_INFO 帶入
                            TestMode = stepName,
                            PgName = request.PgName,
                            PgId = request.PgId,
                            PgmMatch = "", 
                            Code = "", // 可補上查 TBL_IPN_MASTER 的邏輯
                            CheckSum = "",
                            Speed = "",
                            Temp = request.Temp,
                            WsDeviceFile = request.WsDeviceFile,
                            TecnNo = "", // 若有取得 TECN NO 可寫入
                            ProbeCard = "",
                            EqComments = "",
                            StepComments = "",
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
                        // 若發生 Unique Constraint 衝突，等待 1 秒後重試[cite: 1]
                        await Task.Delay(1000);
                    }
                }

                if (!insertSuccess)
                {
                    throw new Exception($"SETUP FORM 單號寫入失敗,請洽IT人員協助.");
                }

                // 寫入成功後更新機台狀態[cite: 1]
                await _repo.UpdateEqInfoFormNoAsync(request.TesterId, formNo, request.UserId);
                
                response.FormNo = formNo;
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

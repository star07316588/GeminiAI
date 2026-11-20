using Dapper;
using Oracle.ManagedDataAccess.Client;
using System;
using System.Collections.Generic;
using System.Data;
using System.Threading.Tasks;
using MES.Models.Dto.CPAutoMonitor;

namespace MES.Repositories
{
    public class CPAutoMonitorRepository
    {
        private readonly string _connectionString;

        public CPAutoMonitorRepository(string connectionString)
        {
            _connectionString = connectionString;
        }

        // 取得 Oracle 資料庫目前時間 (取代原本迴圈中反覆查 DUAL)
        public async Task<DateTime> GetDbSysDateAsync()
        {
            using (IDbConnection conn = new OracleConnection(_connectionString))
            {
                return await conn.ExecuteScalarAsync<DateTime>("SELECT SYSDATE FROM DUAL");
            }
        }

        // 第一部分：查詢 TBL_CP_AUTO_COMPLETE + WIP 現況
        public async Task<IEnumerable<CpAutoRawDto>> GetCpAutoActiveRecordsAsync(string[] statusList)
        {
            using (IDbConnection conn = new OracleConnection(_connectionString))
            {
                var sql = @"
                    SELECT TCAC.LOTID               AS LotID,
                           TLATT.CUREQPID           AS EqID,
                           C.STEPNAME               AS StepNo,
                           TCAC.STEPNAME            AS StepName,
                           TLATT.LOTSTATUS1         AS LotStatus,
                           E.RULENAME               AS CurrentRule,
                           C.TIMEHERESINCE          AS TimeHereSince,
                           TCAC.LOTENDSTATUS        AS LotEndStatus,
                           TCAC.LOTCOMPLETESTATUS   AS LotCompleteStatus,
                           TCAC.CREATETIME          AS CreateTime,
                           TCAC.UPDATETIME          AS UpdateTime,
                           TCAC.COMMENTS            AS Comments
                    FROM TBL_CP_AUTO_COMPLETE TCAC
                    INNER JOIN TBL_LOT_ATTRIBUTE TLATT
                        ON TCAC.LOTID = TLATT.LOTID
                       AND TCAC.STEPNAME = TLATT.STEPNAME
                    INNER JOIN FWLOT A ON TCAC.LOTID = A.APPID
                    INNER JOIN FWLOT_N2M B ON A.SYSID = B.FROMID
                    INNER JOIN FWWIPSTEP C ON B.TOID = C.SYSID
                    INNER JOIN FWWIPSTEP_N2M D ON C.SYSID = D.FROMID AND C.CURRENTRULEINDEX = D.SEQUENCE
                    INNER JOIN FWWIPSTEPRULE E ON D.TOID = E.SYSID
                    WHERE (TCAC.LOTID, TCAC.STEPNAME, TCAC.CREATETIME) IN (
                              SELECT LOTID, STEPNAME, MAX(CREATETIME)
                              FROM TBL_CP_AUTO_COMPLETE
                              WHERE DELETEFLAG = 'N'
                              GROUP BY LOTID, STEPNAME
                          )
                      AND (TCAC.LOTENDSTATUS IN :StatusList OR TCAC.LOTCOMPLETESTATUS IN :StatusList)
                      AND TCAC.DELETEFLAG = 'N'
                      AND TLATT.LOTSTATUS1 IN ('RUN', 'ERUN', 'BREAK')
                    ORDER BY TCAC.CREATETIME";

                return await conn.QueryAsync<CpAutoRawDto>(sql, new { StatusList = statusList });
            }
        }

        // 第二部分：MoveOut 超時檢查候選清單
        public async Task<IEnumerable<CpAutoRawDto>> GetMoveOutCandidatesAsync(string[] okStatusList)
        {
            using (IDbConnection conn = new OracleConnection(_connectionString))
            {
                var sql = @"
                    SELECT T.LOTID                  AS LotID,
                           T.CUREQPID               AS EqID,
                           C.STEPNAME               AS StepNo,
                           T.STEPNAME               AS StepName,
                           T.LOTSTATUS1             AS LotStatus,
                           E.RULENAME               AS CurrentRule,
                           CP.LOTENDSTATUS          AS LotEndStatus,
                           CP.LOTCOMPLETESTATUS     AS LotCompleteStatus,
                           CP.UPDATETIME            AS UpdateTime,
                           CP.COMMENTS              AS Comments
                    FROM TBL_LOT_ATTRIBUTE T
                    INNER JOIN FWLOT A ON T.LOTID = A.APPID
                    INNER JOIN FWLOT_N2M B ON A.SYSID = B.FROMID
                    INNER JOIN FWWIPSTEP C ON B.TOID = C.SYSID
                    INNER JOIN FWWIPSTEP_N2M D ON C.SYSID = D.FROMID AND C.CURRENTRULEINDEX = D.SEQUENCE
                    INNER JOIN FWWIPSTEPRULE E ON D.TOID = E.SYSID
                    INNER JOIN TBL_CP_AUTO_COMPLETE CP
                        ON T.LOTID = CP.LOTID
                       AND T.STEPNAME = CP.STEPNAME
                    WHERE T.STAGE = 'WS'
                      AND T.LOTSTATUS1 IN ('WAITD', 'BREAK')
                      AND T.RESOURCETYPE = 'WTESTER'
                      AND (CP.LOTID, CP.STEPNAME, CP.CREATETIME) IN (
                              SELECT SUBCP.LOTID, SUBCP.STEPNAME, MAX(SUBCP.CREATETIME)
                              FROM TBL_CP_AUTO_COMPLETE SUBCP
                              WHERE SUBCP.DELETEFLAG = 'N'
                              GROUP BY SUBCP.LOTID, SUBCP.STEPNAME
                          )
                      AND CP.LOTCOMPLETESTATUS IN :OkStatusList
                      AND E.RULENAME = 'MoveOut'";

                return await conn.QueryAsync<CpAutoRawDto>(sql, new { OkStatusList = okStatusList });
            }
        }

        // 第三部分：SortRun 自動化紀錄初始化失敗查詢
        public async Task<IEnumerable<CpAutoRawDto>> GetSortRunLotsAsync()
        {
            using (IDbConnection conn = new OracleConnection(_connectionString))
            {
                var sql = @"
                    SELECT T.LOTID          AS LotID,
                           T.CUREQPID       AS EqID,
                           C.STEPNAME       AS StepNo,
                           T.STEPNAME       AS StepName,
                           T.LOTSTATUS1     AS LotStatus,
                           E.RULENAME       AS CurrentRule
                    FROM TBL_LOT_ATTRIBUTE T
                    INNER JOIN FWLOT A ON T.LOTID = A.APPID
                    INNER JOIN FWLOT_N2M B ON A.SYSID = B.FROMID
                    INNER JOIN FWWIPSTEP C ON B.TOID = C.SYSID
                    INNER JOIN FWWIPSTEP_N2M D ON C.SYSID = D.FROMID AND C.CURRENTRULEINDEX = D.SEQUENCE
                    INNER JOIN FWWIPSTEPRULE E ON D.TOID = E.SYSID
                    WHERE T.STAGE = 'WS'
                      AND T.LOTSTATUS1 IN ('RUN', 'ERUN')
                      AND T.RESOURCETYPE = 'WTESTER'
                      AND E.RULENAME = 'SortRun'";

                return await conn.QueryAsync<CpAutoRawDto>(sql);
            }
        }

        // 檢查 SortRun 後是否有對應的 CPAuto 紀錄
        public async Task<bool> HasSortRunCpAutoRecordAsync(string lotId)
        {
            using (IDbConnection conn = new OracleConnection(_connectionString))
            {
                var sql = @"
                    SELECT COUNT(1)
                    FROM (
                        SELECT A.WIPID AS LOTID, MAX(A.TXNTIMESTAMP) AS SORTRUNTIME
                        FROM FWWIPHISTORY B
                        INNER JOIN FWWIPTRANSACTION A ON A.SYSID = B.WIPTXN
                        WHERE B.GROUPHISTKEY LIKE 'SortRun-%'
                          AND A.WIPID = :LotId
                        GROUP BY A.WIPID
                    ) SRTIME
                    INNER JOIN TBL_CP_AUTO_COMPLETE CP
                        ON SRTIME.LOTID = CP.LOTID
                       AND SRTIME.SORTRUNTIME <= CP.CREATETIME";

                int cnt = await conn.ExecuteScalarAsync<int>(sql, new { LotId = lotId });
                return cnt > 0;
            }
        }

        // 取得該批號未報廢的 Wafer ID 清單 (對應原 VB6 oLot.Components 檢查 Status <> SCRAPPED)
        public async Task<IEnumerable<string>> GetActiveWaferIdsAsync(string lotId)
        {
            using (IDbConnection conn = new OracleConnection(_connectionString))
            {
                var sql = @"
                    SELECT COMP.COMPONENTID
                    FROM FWLOT A
                    INNER JOIN FWLOT_N2M B ON A.SYSID = B.FROMID AND B.LINKNAME = 'components'
                    INNER JOIN FWCOMPONENT COMP ON B.TOID = COMP.SYSID
                    WHERE A.APPID = :LotId
                      AND NVL(COMP.STATUS, ' ') <> 'Scrapped'";

                return await conn.QueryAsync<string>(sql, new { LotId = lotId });
            }
        }

        // 查詢 TBL_WS_TDS_SUM (對應 CheckTdsTestSum)
        public async Task<IEnumerable<TdsSumRawDto>> GetTdsTestSumAsync(string parentLot, string stepNo, string timeHereSince)
        {
            using (IDbConnection conn = new OracleConnection(_connectionString))
            {
                var sql = @"
                    SELECT WAFER_IDN        AS WaferId,
                           TDS_GOOD_DIEN    AS TdsGoodDie,
                           REVISE_GOOD_DIEN AS ReviseGoodDie,
                           REASON_CODE      AS ReasonCode,
                           CREATE_TIME      AS CreateTime,
                           UPDATE_TIME      AS UpdateTime
                    FROM TBL_WS_TDS_SUM
                    WHERE (TEST_MODE, PARENT_LOT_ID, WAFER_IDN, STEPNO, INSTEPTIME, NVL(TIME_STAMP, 'NULL')) IN (
                        SELECT TEST_MODE, PARENT_LOT_ID, WAFER_IDN, STEPNO, INSTEPTIME, NVL(MAX(TIME_STAMP), 'NULL')
                        FROM TBL_WS_TDS_SUM
                        WHERE INSTEPTIME = :TimeHereSince
                          AND STEPNO = :StepNo
                          AND PARENT_LOT_ID = :ParentLot
                          AND DELETE_FLAG = 'N'
                        GROUP BY TEST_MODE, PARENT_LOT_ID, WAFER_IDN, STEPNO, INSTEPTIME
                    )
                    ORDER BY NVL(TIME_STAMP, ' ') DESC";

                return await conn.QueryAsync<TdsSumRawDto>(sql, new
                {
                    ParentLot = parentLot,
                    StepNo = stepNo,
                    TimeHereSince = timeHereSince
                });
            }
        }
    }
}

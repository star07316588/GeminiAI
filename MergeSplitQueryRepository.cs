using Dapper;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Oracle.ManagedDataAccess.Client;
using MES.Models.Dto.MergeSplitQuery;

namespace MES.Repositories
{
    public class MergeSplitQueryRepository
    {
        private readonly string _connectionString;

        public MergeSplitQueryRepository(string connectionString)
        {
            _connectionString = connectionString;
        }

        // 查詢 Merge 紀錄
        public async Task<IEnumerable<MergeSplitQueryResultDto>> GetMergeRecordsAsync(MergeSplitQueryRequest req)
        {
            using (IDbConnection conn = new OracleConnection(_connectionString))
            {
                var sql = new StringBuilder();
                var p = new DynamicParameters();

                // 使用 ANSI JOIN 替換舊的 Oracle (+) 語法
                sql.Append(@"
                    SELECT A.WIPID AS ParentLotId,
                           A.ACTIVITY AS Activity,
                           A.TXNTIMESTAMP AS TimeStamp,
                           B.VALDATA AS ChildLotId,
                           A.MERGESTEPID AS StepId,
                           C.DESCRIPTION AS StepName,
                           D.VALDATA AS ParIPN,
                           A.SYSID AS SysId
                    FROM FWMERGE A
                    INNER JOIN FWMERGE_PN2M B ON A.SYSID = B.FROMID AND B.LINKNAME = 'childLotCollection'
                    LEFT JOIN FWSTEPVERSION C ON SUBSTR(A.MERGESTEPID, 1, 5) = C.STEPNAME AND C.REVSTATE = 'Active'
                    LEFT JOIN FWLOT_PN2M D ON A.LOTOBJECT = D.FROMID AND D.KEYDATA = 'IPN'
                    WHERE 1=1 ");

                AppendConditions(sql, p, "A.WIPID", req.ParentLotId);
                AppendConditions(sql, p, "D.VALDATA", req.Ipn);
                AppendConditions(sql, p, "C.DESCRIPTION", req.StepName);

                if (!string.IsNullOrWhiteSpace(req.TimeStampFrom))
                {
                    sql.Append(" AND A.TXNTIMESTAMP > :TimeStampFrom ");
                    p.Add("TimeStampFrom", req.TimeStampFrom);
                }
                if (!string.IsNullOrWhiteSpace(req.TimeStampTo))
                {
                    sql.Append(" AND SUBSTR(A.TXNTIMESTAMP,1,8) <= :TimeStampTo ");
                    p.Add("TimeStampTo", req.TimeStampTo);
                }

                // ChildLotId 篩選邏輯 (Subquery)
                if (!string.IsNullOrWhiteSpace(req.ChildLotId))
                {
                    sql.Append(@" AND A.SYSID IN (
                        SELECT PARENTTXN FROM FWMERGE 
                        WHERE WIPID = :ChildLotId AND ACTIVITY = 'Merge' AND PARENTTXN IS NOT NULL
                    ) AND B.VALDATA = :ChildLotId ");
                    p.Add("ChildLotId", req.ChildLotId);
                }

                return await conn.QueryAsync<MergeSplitQueryResultDto>(sql.ToString(), p);
            }
        }

        // 查詢 Split 紀錄
        public async Task<IEnumerable<MergeSplitQueryResultDto>> GetSplitRecordsAsync(MergeSplitQueryRequest req)
        {
            using (IDbConnection conn = new OracleConnection(_connectionString))
            {
                var sql = new StringBuilder();
                var p = new DynamicParameters();

                sql.Append(@"
                    SELECT FSL.WIPID AS ParentLotId,
                           FSL.ACTIVITY AS Activity,
                           FSL.TXNTIMESTAMP AS TimeStamp,
                           FS.CHILDLOTID AS ChildLotId,
                           FSLM.VALDATA AS StepId,
                           C.DESCRIPTION AS StepName,
                           D.VALDATA AS ParIPN,
                           FSL.SYSID AS SysId
                    FROM FWSPLITLOT FSL
                    INNER JOIN FWSPLITLOT_N2M FSLN ON FSL.SYSID = FSLN.FROMID
                    INNER JOIN FWSPLIT FS ON FSLN.TOID = FS.SYSID
                    INNER JOIN FWSPLITLOT_PN2M FSLM ON FSL.SYSID = FSLM.FROMID AND FSLM.KEYDATA = 'stepId'
                    LEFT JOIN FWSTEPVERSION C ON SUBSTR(FSLM.VALDATA, 1, 5) = C.STEPNAME AND C.REVSTATE = 'Active'
                    LEFT JOIN FWLOT_PN2M D ON FSL.LOTOBJECT = D.FROMID AND D.KEYDATA = 'IPN'
                    WHERE 1=1 ");

                AppendConditions(sql, p, "FSL.WIPID", req.ParentLotId);
                AppendConditions(sql, p, "D.VALDATA", req.Ipn);
                AppendConditions(sql, p, "C.DESCRIPTION", req.StepName);

                if (!string.IsNullOrWhiteSpace(req.TimeStampFrom))
                {
                    sql.Append(" AND FSL.TXNTIMESTAMP > :TimeStampFrom ");
                    p.Add("TimeStampFrom", req.TimeStampFrom);
                }
                if (!string.IsNullOrWhiteSpace(req.TimeStampTo))
                {
                    sql.Append(" AND SUBSTR(FSL.TXNTIMESTAMP,1,8) <= :TimeStampTo ");
                    p.Add("TimeStampTo", req.TimeStampTo);
                }

                // ChildLotId 篩選邏輯 (Subquery)
                if (!string.IsNullOrWhiteSpace(req.ChildLotId))
                {
                    sql.Append(@" AND FSL.SYSID IN (
                        SELECT PARENTTXN FROM FWSPLITLOT 
                        WHERE WIPID = :ChildLotId AND ACTIVITY = 'Split' AND PARENTTXN IS NOT NULL
                    ) AND FS.CHILDLOTID = :ChildLotId ");
                    p.Add("ChildLotId", req.ChildLotId);
                }

                return await conn.QueryAsync<MergeSplitQueryResultDto>(sql.ToString(), p);
            }
        }

        // 查詢父子批的 WIP 異動紀錄，用來算出 Qty
        public async Task<IEnumerable<WipTransactionDto>> GetWipTransactionsAsync(string parentLotId, string childLotId, string sysId)
        {
            using (IDbConnection conn = new OracleConnection(_connectionString))
            {
                var sql = @"
                    SELECT A.WIPID AS WipId, A.ATTRIBUTE AS Attribute, A.ORIGINALVALUE AS OriginalValue, A.NEWVALUE AS NewValue
                    FROM FWWIPTRANSACTION A
                    INNER JOIN FWWIPHISTORY B ON A.SYSID = B.WIPTXN
                    WHERE A.WIPID IN (:ParentLotId, :ChildLotId)
                      AND A.ACTIVITY = 'ModifyAttribute'
                      AND A.ATTRIBUTE IN ('ChipQty', 'WaferQty')
                      AND B.GROUPHISTKEY = (
                          SELECT GROUPHISTKEY FROM FWWIPHISTORY WHERE WIPTXN = :SysId
                      )
                    ORDER BY A.TXNTIMESTAMP DESC, A.TXNSEQUENCE DESC";

                return await conn.QueryAsync<WipTransactionDto>(sql, new { ParentLotId = parentLotId, ChildLotId = childLotId, SysId = sysId });
            }
        }

        // 查詢 Child Lot 補償用的歷史 WIP 紀錄
        public async Task<IEnumerable<WipTransactionDto>> GetChildLotHistoryQtyAsync(string childLotId, string timeStamp)
        {
            using (IDbConnection conn = new OracleConnection(_connectionString))
            {
                var sql = @"
                    SELECT A.ATTRIBUTE AS Attribute, A.NEWVALUE AS NewValue
                    FROM FWWIPTRANSACTION A
                    WHERE A.WIPID = :ChildLotId
                      AND A.ACTIVITY = 'ModifyAttribute'
                      AND A.ATTRIBUTE IN ('ChipQty', 'WaferQty')
                      AND A.TXNTIMESTAMP < :TimeStamp
                      AND A.TXNSEQUENCE IN (
                          SELECT MAX(B.TXNSEQUENCE) FROM FWWIPTRANSACTION B
                          WHERE B.WIPID = :ChildLotId AND B.ACTIVITY = 'ModifyAttribute'
                            AND B.ATTRIBUTE = 'ChipQty' AND B.TXNTIMESTAMP < :TimeStamp
                          UNION ALL
                          SELECT MAX(B.TXNSEQUENCE) FROM FWWIPTRANSACTION B
                          WHERE B.WIPID = :ChildLotId AND B.ACTIVITY = 'ModifyAttribute'
                            AND B.ATTRIBUTE = 'WaferQty' AND B.TXNTIMESTAMP < :TimeStamp
                      )";

                return await conn.QueryAsync<WipTransactionDto>(sql, new { ChildLotId = childLotId, TimeStamp = timeStamp });
            }
        }

        // 輔助函式：組裝 WHERE 條件 (處理 *)
        private void AppendConditions(StringBuilder sql, DynamicParameters p, string column, string value)
        {
            if (!string.IsNullOrWhiteSpace(value))
            {
                string paramName = column.Replace(".", "");
                if (value.Contains("*"))
                {
                    sql.Append($" AND {column} LIKE :{paramName} ");
                    p.Add(paramName, value.Replace("*", "%"));
                }
                else
                {
                    sql.Append($" AND {column} = :{paramName} ");
                    p.Add(paramName, value);
                }
            }
        }
    }
}

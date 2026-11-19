using Dapper;
using Oracle.ManagedDataAccess.Client;
using System.Collections.Generic;
using System.Data;
using System.Text;
using System.Threading.Tasks;
using MES.Models.Dto.MonitorPCQuery;

namespace MES.Repositories
{
    public class MonitorPCQueryRepository
    {
        private readonly string _connectionString;

        public MonitorPCQueryRepository(string connectionString)
        {
            _connectionString = connectionString;
        }

        // 取得 Area 與 SystemName 下拉選單 (對應原 VB6 Init)
        public async Task<MonitorPCInitDataDto> GetInitDataAsync()
        {
            using (IDbConnection conn = new OracleConnection(_connectionString))
            {
                var areas = await conn.QueryAsync<string>(@"
                    SELECT DISTINCT AREA
                    FROM TBL_MONITORPC_SYSTEMVERSION_IP
                    WHERE DELETEFLAG = 'N'
                    ORDER BY AREA");

                var systemNames = await conn.QueryAsync<string>(@"
                    SELECT DISTINCT MONITORSYSTEM
                    FROM TBL_MONITORPC_SYSTEMVERSION_IP
                    WHERE DELETEFLAG = 'N'
                    ORDER BY MONITORSYSTEM");

                return new MonitorPCInitDataDto
                {
                    Areas = areas,
                    SystemNames = systemNames
                };
            }
        }

        // 查詢主清單 (對應原 VB6 ExecQuery 主查詢)
        public async Task<IEnumerable<MonitorPCQueryResultDto>> GetMonitorPCListAsync(MonitorPCQueryRequest req)
        {
            using (IDbConnection conn = new OracleConnection(_connectionString))
            {
                var sql = new StringBuilder(@"
                    SELECT a.IP AS IP,
                           a.AREA AS Area,
                           a.MONITORSYSTEM AS SystemName,
                           a.COMMENTS AS Comment
                    FROM TBL_MONITORPC_SYSTEMVERSION_IP a
                    WHERE a.DELETEFLAG = 'N' ");

                var p = new DynamicParameters();

                if (!string.IsNullOrWhiteSpace(req.Area))
                {
                    sql.Append(" AND a.AREA = :Area ");
                    p.Add("Area", req.Area.Trim());
                }

                if (!string.IsNullOrWhiteSpace(req.SystemName))
                {
                    sql.Append(" AND a.MONITORSYSTEM = :SystemName ");
                    p.Add("SystemName", req.SystemName.Trim());
                }

                sql.Append(" ORDER BY a.IP ");

                return await conn.QueryAsync<MonitorPCQueryResultDto>(sql.ToString(), p);
            }
        }

        // 查詢指定 SystemName 的各檔案最新版本 (對應 IP = "ALL" 及比對基準)
        public async Task<IEnumerable<MonitorPCDetailDto>> GetLatestVersionBySystemAsync(string systemName)
        {
            using (IDbConnection conn = new OracleConnection(_connectionString))
            {
                var sql = @"
                    SELECT a.FILENAME AS FileName,
                           MAX(a.FILEUPDATETIME) AS FileUpdateTime
                    FROM TBL_MONITORPC_SYSTEMVERSION a
                    INNER JOIN TBL_MONITORPC_SYSTEMVERSION_IP b
                        ON a.IP = b.IP
                       AND a.MONITORSYSTEM = b.MONITORSYSTEM
                    WHERE a.MONITORSYSTEM = :SystemName
                      AND a.DELETEFLAG = 'N'
                      AND b.DELETEFLAG = 'N'
                    GROUP BY a.FILENAME
                    ORDER BY a.FILENAME";

                return await conn.QueryAsync<MonitorPCDetailDto>(sql, new { SystemName = systemName });
            }
        }

        // 查詢單一 IP + SystemName 的軟體版本明細 (對應 Detail 子視窗與版本比對)
        public async Task<IEnumerable<MonitorPCDetailDto>> GetVersionDetailByIpAsync(string ip, string systemName, bool onlyActive = false)
        {
            using (IDbConnection conn = new OracleConnection(_connectionString))
            {
                var sql = new StringBuilder(@"
                    SELECT FILENAME AS FileName,
                           FILEUPDATETIME AS FileUpdateTime
                    FROM TBL_MONITORPC_SYSTEMVERSION
                    WHERE IP = :IP
                      AND MONITORSYSTEM = :SystemName ");

                if (onlyActive)
                {
                    sql.Append(" AND DELETEFLAG = 'N' ");
                }

                sql.Append(" ORDER BY FILENAME ");

                return await conn.QueryAsync<MonitorPCDetailDto>(sql.ToString(), new { IP = ip, SystemName = systemName });
            }
        }
    }
}

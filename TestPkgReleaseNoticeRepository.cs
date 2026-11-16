using Dapper;
using Oracle.ManagedDataAccess.Client;
using System.Data;
using System.Text;
using MES.Models.Dto.TestPkgReleaseNotice;

namespace MES.Repositories
{
    public class TestPkgReleaseNoticeRepository
    {
        private readonly string _connectionString;

        public TestPkgReleaseNoticeRepository(string connectionString)
        {
            _connectionString = connectionString;
        }

        // 取得下拉選單初始資料 (對應原 VB6 的 InitComboBox)
        public async Task<InitDataDto> GetInitDataAsync()
        {
            using (IDbConnection conn = new OracleConnection(_connectionString))
            {
                var pkgTypes = await conn.QueryAsync<string>("SELECT DISTINCT packagecode FROM tbl_prod_spec WHERE packagecode IS NOT NULL");
                var stepNames = await conn.QueryAsync<string>("SELECT DISTINCT stepname FROM tbl_prod_step_eq_spec WHERE SUBSTR(stepname,1,4) = 'SORT' OR SUBSTR(stepname,1,2) = 'FT'");
                var eqType2s = await conn.QueryAsync<string>("SELECT DISTINCT eqtype2 FROM tbl_prod_step_eq_spec WHERE eqtype2 IS NOT NULL");
                var subSystems = await conn.QueryAsync<string>("SELECT DISTINCT subsystem FROM tbl_prod_step_eq_spec WHERE subsystem IS NOT NULL");

                return new InitDataDto
                {
                    PkgTypes = pkgTypes,
                    StepNames = stepNames,
                    EqType2s = eqType2s,
                    SubSystems = subSystems
                };
            }
        }

        // 執行主查詢 (對應原 VB6 的 ExecQuery)
        public async Task<IEnumerable<LotInfoDto>> QueryLotInfoAsync(ReleaseNoticeQueryRequest req)
        {
            using (IDbConnection conn = new OracleConnection(_connectionString))
            {
                var sql = new StringBuilder();
                var parameters = new DynamicParameters();

                // 使用 ANSI JOIN 取代舊版 Oracle (+) 語法，並直接在 SQL 中串接 StepNo,StepName
                sql.Append(@"
                    SELECT 
                        g.magnitude AS TdsProdCode,
                        a.prodgroup AS ProdGroup,
                        b.packagecode AS PkgType,
                        b.pincount AS PinCount,
                        a.stepno || ',' || a.stepname AS StepName,
                        a.eqtype2 AS EqType2,
                        a.subsystem AS SubSystem,
                        a.uph AS Uph
                    FROM (
                        SELECT prodgroup, stepno, stepname, eqtype2, subsystem, uph 
                        FROM tbl_prod_step_eq_spec 
                        WHERE docstatus = 'Active'
                    ) a
                    INNER JOIN (
                        SELECT prodgroup, packagecode, pincount 
                        FROM tbl_prod_spec 
                        WHERE docstatus = 'Active'
                    ) b ON a.prodgroup = b.prodgroup
                    LEFT JOIN (
                        SELECT c.productname, e.magnitude 
                        FROM fwproductversion c
                        INNER JOIN fwproductversion_n2m d ON c.sysid = d.fromid AND d.linkname = 'attributes'
                        INNER JOIN fwprpattributeinstance e ON d.toid = e.sysid
                        INNER JOIN fwprpattribute f ON e.attrclass = f.sysid AND f.name = 'TdsProd'
                        WHERE c.revstate = 'Active'
                    ) g ON a.prodgroup = g.productname 
                    WHERE 1 = 1 ");

                // 動態組裝條件與萬用字元處理
                AppendCondition(sql, parameters, "a.prodgroup", "ProdGroup", req.ProdGroup);
                AppendCondition(sql, parameters, "a.stepname", "StepName", req.StepName);
                AppendCondition(sql, parameters, "a.eqtype2", "EqType2", req.EqType2);
                AppendCondition(sql, parameters, "a.subsystem", "SubSystem", req.SubSystem);
                AppendCondition(sql, parameters, "g.magnitude", "TdsProdCode", req.TdsProdCode);
                AppendCondition(sql, parameters, "b.packagecode", "PkgType", req.PkgType);

                // VB6 中有查詢筆數限制，為保護資料庫，這裡可加入 ROWNUM 限制 (或由前端分頁)
                sql.Append(" AND ROWNUM <= 1000");

                return await conn.QueryAsync<LotInfoDto>(sql.ToString(), parameters);
            }
        }

        // 共用的條件組裝邏輯，自動判斷等號或 LIKE
        private void AppendCondition(StringBuilder sql, DynamicParameters parameters, string column, string paramName, string value)
        {
            if (!string.IsNullOrWhiteSpace(value))
            {
                string formattedValue = value.Replace("*", "%");
                if (formattedValue.Contains("%"))
                {
                    sql.Append($" AND {column} LIKE :{paramName} ");
                }
                else
                {
                    sql.Append($" AND {column} = :{paramName} ");
                }
                parameters.Add(paramName, formattedValue);
            }
        }
    }
}

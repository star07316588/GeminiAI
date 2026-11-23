using Dapper;
using Oracle.ManagedDataAccess.Client;
using System.Data;
using System.Threading.Tasks;
using MES.Models.Dto.YellowCardReSend;

namespace MES.Repositories
{
    public class YellowCardReSendRepository
    {
        private readonly string _connectionString;

        public YellowCardReSendRepository(string connectionString)
        {
            _connectionString = connectionString;
        }

        // 建立連線 (供 Service 開啟 Transaction 使用)
        public IDbConnection CreateConnection()
        {
            return new OracleConnection(_connectionString);
        }

        // 取得 Lot 基本資訊 (取代 FwuRetrieveLot + CustomAttributes)
        public async Task<YellowCardLotInfoDto> GetLotBaseInfoAsync(string lotId)
        {
            using (IDbConnection conn = new OracleConnection(_connectionString))
            {
                var sql = @"
                    SELECT A.APPID                              AS LotId,
                           A.COMPONENTQTY                       AS Qty,
                           TLATT.IPN                            AS IPN,
                           TLATT.DATECODE                       AS Datecode,
                           TLATT.LOTSTATUS1                     AS LotStatus,
                           TLATT.STAGE                          AS Stage,
                           FUN_GET_CUR_LOCATION(A.APPID)        AS Location
                    FROM FWLOT A
                    LEFT JOIN TBL_LOT_ATTRIBUTE TLATT ON A.APPID = TLATT.LOTID
                    WHERE A.APPID = :LotId";

                return await conn.QueryFirstOrDefaultAsync<YellowCardLotInfoDto>(sql, new { LotId = lotId });
            }
        }

        // 取得 IPN 封裝資訊 (對應 TBL_IPN_MASTER)
        public async Task<IpnMasterDto> GetIpnMasterAsync(string ipn)
        {
            using (IDbConnection conn = new OracleConnection(_connectionString))
            {
                var sql = @"
                    SELECT PACKAGE_NAME AS PackageName,
                           BODY_SIZE    AS BodySize,
                           PACKAGE_CODE AS PackageCode
                    FROM TBL_IPN_MASTER
                    WHERE DELETE_FLAG = 'N'
                      AND IPN = :Ipn";

                return await conn.QueryFirstOrDefaultAsync<IpnMasterDto>(sql, new { Ipn = ipn });
            }
        }

        // 檢查 Lot 是否存在
        public async Task<bool> LotExistsAsync(IDbConnection conn, IDbTransaction trans, string lotId)
        {
            var sql = "SELECT COUNT(1) FROM FWLOT WHERE APPID = :LotId";
            int cnt = await conn.ExecuteScalarAsync<int>(sql, new { LotId = lotId }, trans);
            return cnt > 0;
        }

        // 依 Location 查詢 VendorNo (對應 TBL_VENDOR_MASTER)
        public async Task<string> GetVendorNoByLocationAsync(IDbConnection conn, IDbTransaction trans, string location)
        {
            var sql = @"
                SELECT VENDOR_NO
                FROM TBL_VENDOR_MASTER
                WHERE DELETE_FLAG = 'N'
                  AND VENDOR_NAME = :Location";

            return await conn.QueryFirstOrDefaultAsync<string>(sql, new { Location = location }, trans);
        }

        // 寫入 TBL_YELLOW_CARD
        public async Task InsertYellowCardAsync(
            IDbConnection conn,
            IDbTransaction trans,
            YellowCardLotInfoDto item,
            string wQty,
            string vendorNo,
            string userId)
        {
            var sql = @"
                INSERT INTO TBL_YELLOW_CARD (
                    TKTYPE,
                    NEWLOTNO,
                    IPN,
                    QTY,
                    WQTY,
                    TOVENDOR,
                    DATECODE,
                    PACKAGENAME,
                    BODYSIZE,
                    DATETIME,
                    DATA_SOURCE,
                    CREATEUSERID,
                    CREATETIME
                ) VALUES (
                    'Manual',
                    :LotId,
                    :IPN,
                    :Qty,
                    :WQty,
                    :ToVendor,
                    :Datecode,
                    :PackageName,
                    :Bodysize,
                    TO_CHAR(SYSDATE, 'yyyymmdd'),
                    'Manual',
                    :UserId,
                    TO_CHAR(SYSDATE, 'yyyymmdd hh24miss') || '000'
                )";

            await conn.ExecuteAsync(sql, new
            {
                item.LotId,
                item.IPN,
                item.Qty,
                WQty = wQty,
                ToVendor = vendorNo,
                item.Datecode,
                item.PackageName,
                item.Bodysize,
                UserId = userId
            }, trans);
        }
    }
}

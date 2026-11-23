using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using MES.Models.Dto.YellowCardReSend;
using MES.Repositories;

namespace MES.Services
{
    public class YellowCardReSendService
    {
        private readonly YellowCardReSendRepository _repository;
        private const string DEFAULT_VENDOR_NO = "0000101008";

        public YellowCardReSendService(YellowCardReSendRepository repository)
        {
            _repository = repository;
        }

        // 查詢單筆 Lot 資訊 (對應 GetLotInfo + spdList_ButtonClicked 檢核)
        public async Task<YellowCardLotInfoDto> GetLotInfoAsync(string lotId)
        {
            if (string.IsNullOrWhiteSpace(lotId))
                throw new ArgumentException("請輸入 LotId。");

            var lot = await _repository.GetLotBaseInfoAsync(lotId.Trim());
            if (lot == null)
                throw new KeyNotFoundException($"Lotid : {lotId} 不存在.");

            // 檢核狀態與 Stage (對應原 VB6 勾選時的檢查)
            if (string.Equals(lot.LotStatus, "Terminate", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(lot.LotStatus, "Ship", StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException($"Lotid : {lotId} Status is {lot.LotStatus}.");
            }

            if (!string.Equals(lot.Stage, "FT", StringComparison.OrdinalIgnoreCase))
            {
                throw new InvalidOperationException($"Lotid : {lotId} Stage is not FT.");
            }

            // 帶出 IPN Master 封裝資訊
            if (!string.IsNullOrWhiteSpace(lot.IPN))
            {
                var ipnInfo = await _repository.GetIpnMasterAsync(lot.IPN);
                if (ipnInfo != null)
                {
                    lot.PackageName = ipnInfo.PackageName;
                    lot.Bodysize = ipnInfo.BodySize;
                    lot.PackageCode = ipnInfo.PackageCode;
                }
            }

            return lot;
        }

        // 批次重送黃卡資料 (對應 cmdOk_Click)
        public async Task<int> SendYellowCardAsync(YellowCardSendRequest req)
        {
            if (req?.Items == null || !req.Items.Any())
                throw new ArgumentException("請至少勾選一筆資料。");

            string userId = string.IsNullOrWhiteSpace(req.UserId) ? "SYSTEM" : req.UserId;

            using (var conn = _repository.CreateConnection())
            {
                conn.Open();
                using (var trans = conn.BeginTransaction())
                {
                    try
                    {
                        // 1. 檢核所有勾選的 Lot 是否存在
                        foreach (var item in req.Items)
                        {
                            bool exists = await _repository.LotExistsAsync(conn, trans, item.LotId);
                            if (!exists)
                                throw new KeyNotFoundException($"LotId:{item.LotId} 不存在.");
                        }

                        // 2. 逐筆計算 WQty、轉換 VendorNo 並寫入 TBL_YELLOW_CARD
                        foreach (var item in req.Items)
                        {
                            bool isWaferOrHub = string.Equals(item.PackageCode, "W", StringComparison.OrdinalIgnoreCase) ||
                                                string.Equals(item.PackageCode, "H", StringComparison.OrdinalIgnoreCase);
                            string wQty = isWaferOrHub ? (item.Qty ?? "0") : "0";

                            string vendorNo = await _repository.GetVendorNoByLocationAsync(conn, trans, item.Location);
                            if (string.IsNullOrWhiteSpace(vendorNo))
                            {
                                vendorNo = DEFAULT_VENDOR_NO;
                            }

                            await _repository.InsertYellowCardAsync(conn, trans, item, wQty, vendorNo, userId);
                        }

                        trans.Commit();
                        return req.Items.Count;
                    }
                    catch
                    {
                        trans.Rollback();
                        throw;
                    }
                }
            }
        }
    }
}

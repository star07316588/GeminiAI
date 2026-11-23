using System.Collections.Generic;

namespace MES.Models.Dto.YellowCardReSend
{
    // 查詢單筆 Lot 資訊回傳 (對應 GetLotInfo 與 spdList 欄位)
    public class YellowCardLotInfoDto
    {
        public string LotId { get; set; }
        public string IPN { get; set; }
        public string Qty { get; set; }
        public string Location { get; set; }
        public string Datecode { get; set; }
        public string PackageName { get; set; }
        public string Bodysize { get; set; }
        public string PackageCode { get; set; } // 隱藏欄位，計算 WQty 用

        // 供前端/後端勾選檢核用
        public string LotStatus { get; set; }
        public string Stage { get; set; }
    }

    // 送出請求 (對應 cmdOk_Click)
    public class YellowCardSendRequest
    {
        public string UserId { get; set; }
        public List<YellowCardLotInfoDto> Items { get; set; }
    }

    // 內部使用：承接 IPN Master 查詢結果
    public class IpnMasterDto
    {
        public string PackageName { get; set; }
        public string BodySize { get; set; }
        public string PackageCode { get; set; }
    }
}

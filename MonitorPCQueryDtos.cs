using System.Collections.Generic;

namespace MES.Models.Dto.MonitorPCQuery
{
    // 下拉選單初始資料
    public class MonitorPCInitDataDto
    {
        public IEnumerable<string> Areas { get; set; }
        public IEnumerable<string> SystemNames { get; set; }
    }

    // 主查詢請求
    public class MonitorPCQueryRequest
    {
        public string Area { get; set; }
        public string SystemName { get; set; }
    }

    // 主查詢結果 (對應 spdLotList)
    public class MonitorPCQueryResultDto
    {
        public string IP { get; set; }
        public string Area { get; set; }
        public string SystemName { get; set; }
        public string LastVersionFlag { get; set; }
        public string Comment { get; set; }
    }

    // 子視窗明細請求 (點選 Detail 傳該列 IP，點選「最新版本」傳 "ALL")
    public class MonitorPCDetailRequest
    {
        public string IP { get; set; }
        public string SystemName { get; set; }
    }

    // 子視窗明細結果 (對應 spdEngCost)
    public class MonitorPCDetailDto
    {
        public string FileName { get; set; }
        public string FileUpdateTime { get; set; }
    }
}

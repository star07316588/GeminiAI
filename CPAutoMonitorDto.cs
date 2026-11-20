namespace MES.Models.Dto.CPAutoMonitor
{
    // 回傳前端 Grid 的監控清單項目
    public class CPAutoMonitorItemDto
    {
        public int Priority { get; set; }
        public bool IsAlert { get; set; } // true 時前端 Action 欄位顯示紅字
        public string EqID { get; set; }
        public string LotID { get; set; }
        public string StepName { get; set; }
        public string Action { get; set; }
        public string StepNo { get; set; }
        public string LotStatus { get; set; }
        public string CurrentRule { get; set; }
        public string TdsTestSumTime { get; set; }
        public string CPAutoStatusTime { get; set; }
        public string LotEndStatus { get; set; }
        public string LotCompleteStatus { get; set; }
        public string Comments { get; set; }
    }

    // 內部使用：承接 Part 1 / Part 2 / Part 3 SQL 查詢結果
    public class CpAutoRawDto
    {
        public string LotID { get; set; }
        public string EqID { get; set; }
        public string StepNo { get; set; }
        public string StepName { get; set; }
        public string LotStatus { get; set; }
        public string CurrentRule { get; set; }
        public string TimeHereSince { get; set; }
        public string LotEndStatus { get; set; }
        public string LotCompleteStatus { get; set; }
        public string CreateTime { get; set; }
        public string UpdateTime { get; set; }
        public string Comments { get; set; }
    }

    // 內部使用：承接 TBL_WS_TDS_SUM 查詢結果
    public class TdsSumRawDto
    {
        public string WaferId { get; set; }
        public string TdsGoodDie { get; set; }
        public string ReviseGoodDie { get; set; }
        public string ReasonCode { get; set; }
        public string CreateTime { get; set; }
        public string UpdateTime { get; set; }
    }
}

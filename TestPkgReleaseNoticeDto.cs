namespace MES.Models.Dto.TestPkgReleaseNotice
{
    // 前端查詢條件 Request
    public class ReleaseNoticeQueryRequest
    {
        public string TdsProdCode { get; set; }
        public string ProdGroup { get; set; }
        public string PkgType { get; set; }
        public string StepName { get; set; }
        public string EqType2 { get; set; }
        public string SubSystem { get; set; }
    }

    // Grid 回傳資料
    public class LotInfoDto
    {
        public string TdsProdCode { get; set; }
        public string ProdGroup { get; set; }
        public string PkgType { get; set; }
        public string PinCount { get; set; }
        public string StepName { get; set; } // 將會是 StepNo,StepName
        public string EqType2 { get; set; }
        public string SubSystem { get; set; }
        public string Uph { get; set; }
    }

    // 下拉選單初始資料
    public class InitDataDto
    {
        public IEnumerable<string> PkgTypes { get; set; }
        public IEnumerable<string> StepNames { get; set; }
        public IEnumerable<string> EqType2s { get; set; }
        public IEnumerable<string> SubSystems { get; set; }
    }
}

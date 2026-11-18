using System.Collections.Generic;

namespace MES.Models.Dto.MergeSplitQuery
{
    public class MergeSplitQueryRequest
    {
        public string StepName { get; set; }
        public string Ipn { get; set; }
        public string ParentLotId { get; set; }
        public string ChildLotId { get; set; }
        public string Activity { get; set; } // "All", "Merge", "Split"
        public string TimeStampFrom { get; set; }
        public string TimeStampTo { get; set; }
    }

    public class MergeSplitQueryResultDto
    {
        public string Activity { get; set; }
        public string ParentLotId { get; set; }
        public string ParIPN { get; set; }
        public string StepId { get; set; }
        public string StepName { get; set; }
        public string OriParWQty { get; set; }
        public string OriParCQty { get; set; }
        public string AftParWQty { get; set; }
        public string AftParCQty { get; set; }
        public string ChildLotId { get; set; }
        public string ChildWQty { get; set; }
        public string ChildCQty { get; set; }
        public string TimeStamp { get; set; }
        
        // 內部使用，不需傳到前端
        public string SysId { get; set; } 
    }

    // 用於接收 Repository WIP Transaction 查詢結果的內部 DTO
    public class WipTransactionDto
    {
        public string WipId { get; set; }
        public string Attribute { get; set; }
        public string OriginalValue { get; set; }
        public string NewValue { get; set; }
    }
}

public class DropdownRequest
{
    public string LabelSpecNo { get; set; }
    public string CarrierType { get; set; }
    public string BoxingSpecNo { get; set; }
    public string Brand { get; set; }
    public string PinCount { get; set; }
}
// 💡 新增一個小 DTO 來接參數 (可放在 Controller 上方或 DTO 目錄下)
public class LotInfoRequest
    {
        public string LotId { get; set; }
        // 💡 必須傳入 LabelFormat，後端才知道要用哪套邏輯查資料
        public string LabelFormat { get; set; } 
    }

    public class LotDetailData
    {
        public string ProductNo { get; set; }
        public string WaferQty { get; set; }
        public string ChipQty { get; set; }
        public string LotOwner { get; set; }
        public string Speed { get; set; }
        public string Route { get; set; }
        public string Step { get; set; }
    }

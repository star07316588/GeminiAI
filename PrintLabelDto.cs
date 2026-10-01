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
}

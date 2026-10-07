namespace MES.Net.Shared.DTOs.Print
{
    // 畫面初始化回傳的基礎資料
    public class PrintLabelInitResponse
    {
        public IEnumerable<string> Stages { get; set; }
        public IEnumerable<string> LabelTypes { get; set; }
        public IEnumerable<string> LabelFormats { get; set; }
        public IEnumerable<string> PrinterServers { get; set; }
        public IEnumerable<string> CarrierTypes { get; set; }
    }

    // 連動下拉選單共用的請求物件
    public class DropdownRequest
    {
        public string LabelSpecNo { get; set; }
        public string CarrierType { get; set; }
        public string BoxingSpecNo { get; set; }
        public string Brand { get; set; }
        public string PinCount { get; set; }
    }

    // 印表機伺服器連動請求物件
    public class PrinterRequest
    {
        public string PrinterServer { get; set; }
    }

    /// <summary>
    /// 點擊 Print 按鈕時的完整請求物件 (對應前端的 payload 以及後端 Master Dispatcher)
    /// </summary>

    // 💡 新增一個小 DTO 來接參數 (可放在 Controller 上方或 DTO 目錄下)
    public class LotInfoRequest
    {
        public string LotId { get; set; }
        public string LabelFormat { get; set; } 
    }

    public class LotAttributeData
    {
        public string Ipn { get; set; }
        public string StepName { get; set; }
        public string StepId { get; set; }
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

    public class EtestMergeResponse
    {
        public string TotalQty { get; set; }
        public List<EtestChildLot> ChildLots { get; set; } = new List<EtestChildLot>();
    }

    public class EtestChildLot
    {
        public string ChildLotId { get; set; }
        public string Qty { get; set; }
    }

    public class TrLabelInfoRequest
    {
        public string LotId { get; set; }
        // 若前端有手動更改 Qty/Reel，可傳入此參數重新計算
        public int? OverrideQtyReel { get; set; } 
    }

    public class TrLabelInfoResponse
    {
        public string Ipn { get; set; }
        public int QtyReel { get; set; }
        public int ChipQty { get; set; }
        public List<string> ReelIds { get; set; } = new List<string>();
    }
    
    public class PrintLabelRequest
    {
        // ==========================================
        // 1. 基礎列印條件 (Main Settings)
        // ==========================================
        public string Stage { get; set; }
        public string LabelType { get; set; }
        public string LabelFormat { get; set; }
        public string LotId { get; set; }           // 對應畫面的 Label Data 或主批號
        public string IPN { get; set; }             // 產品料號 (通常在前端輸入 LotId 後帶出，或後端補查)
        public string PrintMode { get; set; }       // Normal, Reprint, Modify

        // ==========================================
        // 2. 印表機與數量設定 (Printer & Qty)
        // ==========================================
        public string PrinterServer { get; set; }
        public string Printer { get; set; }
        public int BoxQty { get; set; }             // 每箱數量 (取代 VB6 的 frmBoxQty 彈窗)
        public int PrintQty { get; set; }           // 列印份數 (對應某些固定印 1 張或自訂張數的標籤)

        // ==========================================
        // 3. 數量與站點資訊 (Lot Info)
        // ==========================================
        public string WQty { get; set; }            // Wafer 數量 (Wafer Pcs)
        public string CQty { get; set; }            // Chip 數量 (Quantity EA)
        public string LotOwner { get; set; }        // 批號擁有者
        public string RouteId { get; set; }         // 流程代碼

        // ==========================================
        // 4. Reprint 補印模式專用條件
        // ==========================================
        public string ReprintType { get; set; }     // LastLabel 或 SearchData
        public string SearchData { get; set; }

        // ==========================================
        // 5. Label Pack Info (包裝規格連動選單)
        // ==========================================
        public string CarrierType { get; set; }
        public string BoxingSpecNo { get; set; }
        public string Brand { get; set; }
        public string PinCount { get; set; }
        public string PackageCode { get; set; }     // 封裝代碼

        // ==========================================
        // 6. 特殊標籤外部傳入參數 (Optional)
        // ==========================================
        public string WaferId { get; set; }         // 外部指定 WaferID (例如 CP_SMALL_LABEL)
        public string FabLotId { get; set; }        // 外部指定 FabLotId
        public string OriLotID { get; set; }        // 原始母批批號 (用於併批/降級標籤)
        public List<string> WSMCDInfoList { get; set; } // WSMCD 併批清單 (格式: "ChildLot;IPN;Code;Qty")

        // ==========================================
        // 7. 系統/使用者資訊
        // ==========================================
        public string UserId { get; set; }          // 登入者工號
        public bool? IsPass { get; set; }
    }
    // 定義回傳 WS_SMALL_LABEL 所需額外資料的 DTO
    public class WsSmallLabelDbData
    {
        public string HotLotFlag { get; set; }
        public string ErunTicNo { get; set; }
        public string SapRwNo { get; set; }
        public string MfgTicNo { get; set; }
        public string WaferId { get; set; }
        public string FabLotId { get; set; }
        public string FgIpn { get; set; } 
    }

    public class InklessMergeItem
    {
        public string ChildLotId { get; set; }
        public string WaferId { get; set; }
        public string SlotNo { get; set; }
    }

    public class EngLocMasterData
    {
        public string WQty { get; set; }
        public string OwnerId { get; set; }
        public string OwnerDep { get; set; }
        public string LocId { get; set; }
    }

    public class LabelSpecInfoData
    {
        public string LabelSpecNoVer { get; set; }
        public string SpecialLabelSize { get; set; }
        public string Serial { get; set; }
        public string MergeLabelSpec { get; set; }
        public string CopiesStdLabel { get; set; }
        public string CopiesSpecialLabel { get; set; }
        public string CopiesMergeLabel { get; set; }
        public string SpecialStickReel { get; set; }
        public string SpecialStickAlbag { get; set; }
        public string SpecialPositionBox_1 { get; set; }
        public string SpecialPositionBox_2 { get; set; }
    }

    public class LabelBoxingInfoData
    {
        public string BoxingSpecNoVer { get; set; }
        public string Vacuum { get; set; }
        public string Hic { get; set; }
        public string Desiccant { get; set; }
        public string Albag { get; set; }
    }

    public class IpnWaferLevelData
    {
        public string WaferLevel { get; set; }
        public string MaskOption { get; set; }
        public string BeOption { get; set; }
        public string BomLevel { get; set; }
    }

    public class VirtualMergeLotData
    {
        public string LotId { get; set; }
        public string Pcs { get; set; }
    }

    // 彙總後的 IPN 與數量模型
    public class VirtualLotAggregateData
    {
        public string Ipn { get; set; }
        public string WQty { get; set; }
        public string CQty { get; set; }
    }

    // 虛擬批 Slot 槽位模型
    public class VirtualMergeSlotData
    {
        public string SlotNo { get; set; }
        public string LotId { get; set; }
        public string Pcs { get; set; }
    }
    // 定義 FT_LOT_INFO 需要的 DTO
    public class FtLotInfoDbData
    {
        public string VendorCode { get; set; }
        public string LotIpn { get; set; }
        public string CQty { get; set; }
        public string RouteId { get; set; }
        public string LotOwner { get; set; }
        public string DateCode { get; set; }
        public string IpnCarGradeFlag { get; set; }
        public string IpnProdLine { get; set; }
        public string CarrierType { get; set; }
        public string ExtraRomFlag { get; set; }
        public string Brand { get; set; }
        public string IcDrawing { get; set; }
        public string MarkingSpecNo { get; set; }
        public string IpnGreen { get; set; }
        public string IpnCsum { get; set; }
        public string NFlag { get; set; }
    }

    // 定義 FT_SMALL_LABEL 需要的 DTO
    public class FtSmallLabelDbData
    {
        public string Customer { get; set; }
        public string Green { get; set; }
        public string Location { get; set; }
        public string Csum { get; set; }
        public string GPType { get; set; }
        public string CarGradeFlag { get; set; }
        public string FuSa { get; set; }
        public string ProdBody { get; set; }
        public string PinCount { get; set; }
        public string PackageCode { get; set; }
        public string BodySize { get; set; }
        public string ProdLine { get; set; }
    }
    /// <summary>
    /// 連動下拉選單共用的請求物件 (擴充 PinCount 以支援第 5 層查詢)
    /// </summary>
    public class DropdownRequest
    {
        public string CarrierType { get; set; }
        public string BoxingSpecNo { get; set; }
        public string Brand { get; set; }
        public string PinCount { get; set; } // 💡 新增：為了查 PackageCode 必須傳入
    }

    /// <summary>
    /// 依據 Stage 與 LabelFormat 動態取得可用印表機的請求物件
    /// </summary>
    public class MappedServerRequest
    {
        public string Stage { get; set; }
        public string LabelFormat { get; set; }
    }

    // 查詢母批與 MCP 旗標的 DTO
    public class McpInfoData
    {
        public string McdParent { get; set; }
        public string McpFlag { get; set; }
        public string ParentIpn { get; set; }
    }

    // 子批號明細資料 DTO
    public class WsmcdChildLotData
    {
        public string LotId { get; set; }
        public string Ipn { get; set; }
        public string BinOrRank { get; set; }
        public int CQty { get; set; }
        public int WQty { get; set; }
    }
}

using System.Collections.Generic;

namespace MES.Net.Shared.DTOs.Print
{
    public class PrintSetupFormQueryRequest
    {
        public string LotId { get; set; }
    }

    public class PrintSetupFormQueryResponse
    {
        public string LotId { get; set; }
        public string Stage { get; set; }
        public string LotStatus { get; set; }
        public string IPN { get; set; }
        public string LotOwner { get; set; }

        // 下拉選單資料來源
        public List<SelectItem> StepList { get; set; } = new List<SelectItem>();
        public string SelectedStep { get; set; } // 預設選取的 Step

        public List<SelectItem> TesterIdList { get; set; } = new List<SelectItem>();
        public string SelectedTesterId { get; set; } // 預設選取的 TesterId

        public List<SelectItem> SubSystemList { get; set; } = new List<SelectItem>();

        // 🌟 新增：存放 Setup Reason 的下拉選單資料
        public List<SelectItem> SetupReasonList { get; set; } = new List<SelectItem>();
    }

    public class SelectItem
    {
        public string Text { get; set; }
        public string Value { get; set; }
    }

    // 內部傳遞用的 Lot 基本資訊 Model
    public class SetupLotBasicInfo
    {
        public string LotId { get; set; }
        public string Status { get; set; }
        public string Ipn { get; set; }
        public string Owner { get; set; }
        public string CurrentStepId { get; set; }
        public string CurrentStepName { get; set; }
        public string CurEqpId { get; set; }
        public string SplitAssignEqId { get; set; }
    }

    public class PrintSetupSubsystemRequest
    {
        public string LotId { get; set; }
        public string Stage { get; set; }
        public string TesterId { get; set; }
        public string StepNo { get; set; }
        public string StepName { get; set; }
        public string SubSystemValue { get; set; } // 例如: "SYS1,SITE1" 或 "SYS1"
    }

    public class PrintSetupSubsystemResponse
    {
        public string PgName { get; set; }
        public string PgId { get; set; }
        public string PgMode { get; set; }
        public string Temp { get; set; }
        public string WsDeviceFile { get; set; }
        public string AccType { get; set; }
        public string SpecifyEq { get; set; }
        public string SpecifyEqId { get; set; }
        public string StopTicNo { get; set; }
    }

    // 用來承接三種資料表共同欄位的內部 Model
    public class RecipeSpecData
    {
        public string PgName { get; set; }
        public string PgId { get; set; }
        public string PgMode { get; set; }
        public string Temperature { get; set; }
        public string WsDeviceFile { get; set; }
        public string ProbeCardType { get; set; }
        public string LoadboardType { get; set; }
        public string ContactboardType { get; set; }
        public string BurnInBoard { get; set; }
        public string SpecifyEq { get; set; }
        public string EqId { get; set; }
        public string StopTicNo { get; set; }
        
        // 供 GetPGM 替換用的屬性
        public string RefStepName01 { get; set; }
        public string RefPgName01 { get; set; }
        public string ReplacePgName01 { get; set; }
        // ... (視需求擴充 02, 03)
    }
    // 對應 GetFTAccByPgm 的回傳值
    public class FtAccDto
    {
        public string Loadboard { get; set; }
        public string Contactboard { get; set; }
        public string Burninboard { get; set; }
    }

    // 對應 GetPgIdAndNameFromFutAct 的回傳值
    public class FutActPgmDto
    {
        public string PgId { get; set; }
        public string PgName { get; set; }
        public string DocNo { get; set; }
        public string PgMode { get; set; }
    }

    // 對應 GetTecnPgmRecipeAttr 的回傳值
    public class TecnPgmRecipeAttrDto
    {
        public string RefPgmTecnNo { get; set; }
        public string RefPgm { get; set; }
        public string RefPgId { get; set; }
        public string RefPgmSource { get; set; }
        public string RefTempTecnNo { get; set; }
        public string RefTemp { get; set; }
        public string RefTempSource { get; set; }
        public string RefOverTime { get; set; }
        public string RefLevel { get; set; }
    }

    public class PrintSetupFormSubmitRequest
    {
        public string LotId { get; set; }
        public string Stage { get; set; }
        public string TesterId { get; set; }
        public string StepNo { get; set; }
        public string SubSystemValue { get; set; }
        public string WaferId { get; set; }
        public string PgName { get; set; }
        public string SetupReason { get; set; }
        
        // 畫面上的 Setup Wafer IDs
        public string SetupWaferId1 { get; set; }
        public string SetupWaferId2 { get; set; }
        public string SetupWaferId3 { get; set; }

        public string InStepTime { get; set; }
        
        // 從前端接回來的隱藏或顯示欄位 (對應 VB 的 txtTemp, txtWsDeviceFile 等)
        public string AccType { get; set; }
        public string SpecifyEq { get; set; }
        public string SpecifyEqId { get; set; }
        public string StopTicNo { get; set; }
        public string PgId { get; set; }
        public string PgMode { get; set; }
        public string Temp { get; set; }
        public string WsDeviceFile { get; set; }
        
        public string UserId { get; set; } // 目前操作的使用者
    }

    public class PrintSetupFormSubmitResponse
    {
        public string FormNo { get; set; }
        public string StopInfoMsg { get; set; }
        // 未來若需要回傳整包列印報表資料 (如 FT 的 Excel 資料)，可擴充於此
    }

    // 新增這個 DTO 來接 SubSystem 的查詢結果
    public class SubSystemResultDto
    {
        public string SubSystem { get; set; }
        public string MaxSite { get; set; }
        public string LoadBoard { get; set; }
    }

    public class EqAccessoriesDto
    {
        public string ProbeCardId { get; set; }
        public string LoadBoardId { get; set; }
        public string ContactBoardId { get; set; }
        public string WsDeviceFile { get; set; }
    }

    public class ProductStopTestRecordDto
    {
        public string StopTicNo { get; set; }
        public string SpecifyEq { get; set; }
        public string EqId { get; set; }
        public string AccName { get; set; }
        public string AccNo { get; set; }
        public string DeviceFile { get; set; }
        public string BePeSet { get; set; }
    }
}

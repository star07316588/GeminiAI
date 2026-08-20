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
        public string TecnNo { get; set; } // ✅ 新增
        public string StepComments { get; set; } // ✅ 新增
        public string DocType { get; set; } // ✅ 新增：文件類別
        public string DocNo { get; set; }   // ✅ 新增：文件單號

        public string CableType { get; set; }     // ✅ 擴充
        public string KitType { get; set; }       // ✅ 擴充
        public string NeedJumper { get; set; }    // ✅ 擴充
        public string JumperPinNo { get; set; }   // ✅ 擴充
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
    // 專門用來裝載要塞進 Excel 裡的資料
    public class FtSetupExcelDto
    {
        public string SheetName { get; set; } // 要顯示的 Sheet: "FT" 或 "AT3-300AL" 或 "FT-940"

        // === 基本標頭資訊 ===
        public string OperatorId { get; set; }      // C3
        public string ShiftCode { get; set; }       // H3
        public string PrintDate { get; set; }       // R3
        public string StopInfo { get; set; }        // F5
        public string LotId { get; set; }           // E6
        public string Ipn { get; set; }             // E7
        public string LotOwner { get; set; }        // R6
        public string StepName { get; set; }        // E8
        public string TesterId { get; set; }        // R8
        public string ErunTicNo { get; set; }       // R7
        
        // === 產品/封裝資訊 (IPN Master) ===
        public string PackageName { get; set; }     // E9
        public string Speed { get; set; }           // E10
        public string Code { get; set; }            // R10
        public string CheckSum { get; set; }        // R11
        public string PinCount { get; set; }
        public string PackageCode { get; set; }
        public string PbFree { get; set; }          // 判斷 E36/G36 的打勾

        // === 配方與機台設定 (Spec & EqInfo) ===
        public string PgmName { get; set; }         // E11
        public string Temperature { get; set; }     // R12
        public string WsDeviceFile { get; set; }    // (Sthand File) E12
        public string SubSystem1 { get; set; }      // E20
        public string SubSystem2 { get; set; }      // R20
        public string LoadBoard { get; set; }       // E28 / R28
        public string Cable { get; set; }           // E29 / R29
        public string ContactBoard { get; set; }    // E30 / R30
        public string KitType { get; set; }         // E31 / R31
        public string MatchPlate { get; set; }      // E33 / R33
        
        public string SoakTime { get; set; }        // R13
        public string NeedJumper { get; set; }      // E16
        public string JumperPinNo { get; set; }     // E17
        public string Comments { get; set; }        // E16 / E17 (對應位置依下移量而定)
        
        public string Pitch { get; set; }           // E35
        public string VacuumCup { get; set; }       // E36

        // === 設備架機確認清單 (Recipe List) ===
        // 裝載 1~40 筆的 Recipe Spec (對應 A49/B49 或 P49/Q49)
        public List<RecipeItemDto> RecipeList { get; set; } = new List<RecipeItemDto>();
    }

    public class RecipeItemDto
    {
        public string Id { get; set; }
        public string SpecName { get; set; }
    }
}

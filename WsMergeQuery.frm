Private Sub ExecQuery()
On Error GoTo ExitHandler:
    Dim sProcID             As String
    Dim typErrInfo          As tErrInfo

    '----
' Init
'----
    sProcID = "ExecQuery"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog)

'----
' Condition Checking
'----
    
     'Modified by Jack on 2021/03/03 for UAT bug fix (兩頁查詢使用不同欄位LotID)
    If ssTabQuery.Tab = 0 Then
        If Trim(txtLotID.Text) = "" Then
                UtShowMsgBox "Pls input Query Criterial !!" & vbNewLine & "請輸入查詢條件 !!"
                GoTo ExitHandler
        End If
    Else
        If Trim(txtLotID_History.Text) = "" Then
                UtShowMsgBox "Pls input Query Criterial !!" & vbNewLine & "請輸入查詢條件 !!"
                GoTo ExitHandler
        End If
    End If
    

'----
' Action
'----
    Dim sSQL As String
    Dim colRS As Collection
    
    Dim lIndex As Long
    
    Screen.MousePointer = vbHourglass

    Me.spdLotInfo.MaxRows = 0
    
    'Added by Jack on 2021/01/05 for CP測前併批專案 <Start>.
    '注意 : 由於共用Spread, 所以查詢 Tbl_Virtual_Lot 或 Tbl_Inkless_Merge_List(_His) 欄位都要同步, 別名要一樣 !!
    'Step-1 : 先檢查是否為 Virtual Lot ?
     'Modified by Jack on 2021/03/03 for UAT bug fix (兩頁查詢使用不同欄位LotID)
    If ssTabQuery.Tab = 0 Then
        sSQL = "select  a.* " & _
                " from " & gsCAT_TBL_VIRTUAL_LOT & " a " & _
               " WHERE a." & gsCAT_TVL_LOTID & " = '" & Me.txtLotID.Text & "' "
    Else
        sSQL = "select  a.* " & _
                " from " & gsCAT_TBL_VIRTUAL_LOT & " a " & _
               " WHERE a." & gsCAT_TVL_LOTID & " = '" & Me.txtLotID_History.Text & "' "
    End If
    
    Set colRS = moProRawSql.QueryDatabase(sSQL)
    If colRS.Count > 0 Then
        sSQL = "select  v." & gsCAT_TVMG_VIRTUALLOTID & " as PARENTLOTID, v." & gsCAT_TVMG_SLOTNO & ", " & _
                      " v." & gsCAT_TVMG_LOTID & " as CHILDLOTID, v." & gsCAT_TVMG_WAFERID & ", " & _
                      " '1' as WQTY, '' as CQTY, '' as SOURCE, '' as ipn, '' as OWNER, '' as MIDDLELOTID, '' as WAFERCHIPQTY, " & _
                      " '' as ERRMSG, '' as FABCREATETIME, '' as RISKLEVEL, " & _
                      " v." & gsCAT_TVMG_MERGESTEPID & " , v." & gsCAT_TVMG_CREATEUSERID & ", " & _
                      " v." & gsCAT_TVMG_CREATETIME & " , v." & gsCAT_TVMG_DELETETIME & ", " & _
                      " v.rowid, '" & gsCAT_TBL_VIRTUAL_MERGE & "' as updatetable "
        sSQL = sSQL & " from " & gsCAT_TBL_VIRTUAL_MERGE & " v "
        
        'Modified by Jack on 2021/03/03 for UAT bug fix (兩頁查詢使用不同欄位LotID)
        If ssTabQuery.Tab = 0 Then
            sSQL = sSQL & " WHERE v." & gsCAT_TVMG_VIRTUALLOTID & " = '" & Me.txtLotID.Text & "' "
        Else
            sSQL = sSQL & " WHERE v." & gsCAT_TVMG_VIRTUALLOTID & " = '" & Me.txtLotID_History.Text & "' "
        End If
        
        'Current 只查 DeleteFlag <> 'Y' 的資料.
        If ssTabQuery.Tab = 0 Then
            sSQL = sSQL & " and nvl(v." & gsCAT_TVMG_DELETEFLAG & ", 'N') <> 'Y'  "
        End If
    Else
        'Modified by Jack on 2021/03/03 for UAT bug fix (兩頁查詢使用不同欄位LotID)
        If ssTabQuery.Tab = 0 Then
            sSQL = "select * from tbl_all_lot a where a.lotid = '" & Me.txtLotID.Text & "' "
        Else
            sSQL = "select * from tbl_all_lot a where a.lotid = '" & Me.txtLotID_History.Text & "' "
        End If
        
        Set colRS = moProRawSql.QueryDatabase(sSQL)
        If Not (colRS.Count > 0) Then
            'Modified by Jack on 2021/03/03 for UAT bug fix (兩頁查詢使用不同欄位LotID)
            If ssTabQuery.Tab = 0 Then
                UtShowMsgBox "Lot '" & Me.txtLotID.Text & "' is not exist in MES !!" & vbNewLine & _
                             "Lot '" & Me.txtLotID.Text & "' 不存在於 MES !!"
            Else
                UtShowMsgBox "Lot '" & Me.txtLotID_History.Text & "' is not exist in MES !!" & vbNewLine & _
                             "Lot '" & Me.txtLotID_History.Text & "' 不存在於 MES !!"
            End If
            GoTo ExitHandler
        End If
        '測後併批要查二個tables : Tbl_Inkless_Merge_List & Tbl_Inkless_Merge_List_His
        '更新SendSapFlag : 由於 Tbl_Inkless_Merge_List_His 不需要重送 SAP, 所以 Livan : 只更新 Tbl_Inkless_Merge_List 即可.
        sSQL = "select  a." & gsCAT_TIML_PARENTLOTID & ", a." & gsCAT_TIML_SLOTNO & ", " & _
                      " a." & gsCAT_TIML_CHILDLOTID & ", a." & gsCAT_TIML_WAFERID & ", " & _
                      " a." & gsCAT_TIML_WQTY & ", a." & gsCAT_TIML_CQTY & ", " & _
                      " a." & gsCAT_TIML_SOURCE & ", a." & gsCAT_TIML_IPN & ", " & _
                      " a." & gsCAT_TIML_OWNER & ", a." & gsCAT_TIML_MIDDLELOTID & ", " & _
                      " a." & gsCAT_TIML_WAFERCHIPQTY & ",a." & gsCAT_TIML_ERRMSG & ", " & _
                      " a." & gsCAT_TIML_FABCREATETIME & ",a." & gsCAT_TIML_RISKLEVEL & ", " & _
                      " '' as mergestepid, " & _
                      " a." & gsCAT_TIML_CREATEUSERID & ", a." & gsCAT_TIML_CREATETIME & ", " & _
                      " a." & gsCAT_TIML_DELETETIME & ", " & _
                      " a.rowid, '" & gsCAT_TBL_INKLESS_MERGE_LIST & "' as updatetable "
        
        sSQL = sSQL & " from " & gsCAT_TBL_INKLESS_MERGE_LIST & " a "
        'Modified by Jack on 2021/03/03 for UAT bug fix (兩頁查詢使用不同欄位LotID)
        If ssTabQuery.Tab = 0 Then
            sSQL = sSQL & " where a." & gsCAT_TIML_PARENTLOTID & " = '" & Me.txtLotID.Text & "' "
        Else
            sSQL = sSQL & " where a." & gsCAT_TIML_PARENTLOTID & " = '" & Me.txtLotID_History.Text & "' "
        End If
        
        'Current 只查 DeleteFlag <> 'Y' 的資料.
        If ssTabQuery.Tab = 0 Then
            sSQL = sSQL & " and nvl(a." & gsCAT_TIML_DELETEFLAG & ", 'N') <> 'Y'"
        End If
        
        sSQL = sSQL & " union "
        
        sSQL = sSQL & _
               "select  a." & gsCAT_TIML_PARENTLOTID & ", a." & gsCAT_TIML_SLOTNO & ", " & _
                      " a." & gsCAT_TIML_CHILDLOTID & ", a." & gsCAT_TIML_WAFERID & ", " & _
                      " a." & gsCAT_TIML_WQTY & ", a." & gsCAT_TIML_CQTY & ", " & _
                      " a." & gsCAT_TIML_SOURCE & ", a." & gsCAT_TIML_IPN & ", " & _
                      " a." & gsCAT_TIML_OWNER & ", a." & gsCAT_TIML_MIDDLELOTID & ", " & _
                      " a." & gsCAT_TIML_WAFERCHIPQTY & ",a." & gsCAT_TIML_ERRMSG & ", " & _
                      " a." & gsCAT_TIML_FABCREATETIME & ",a." & gsCAT_TIML_RISKLEVEL & ", " & _
                      " '' as mergestepid, " & _
                      " a." & gsCAT_TIML_CREATEUSERID & ", a." & gsCAT_TIML_CREATETIME & ", " & _
                      " a." & gsCAT_TIML_DELETETIME & ", " & _
                      " a.rowid, '" & gsCAT_TBL_INKLESS_MERGE_LIST & "' as updatetable "
        
        sSQL = sSQL & " from " & gsCAT_TBL_INKLESS_MERGE_LIST_HIS & " a "
        'Modified by Jack on 2021/03/03 for UAT bug fix (兩頁查詢使用不同欄位LotID)
        If ssTabQuery.Tab = 0 Then
            sSQL = sSQL & " where a." & gsCAT_TIML_PARENTLOTID & " = '" & Me.txtLotID.Text & "' "
        Else
            sSQL = sSQL & " where a." & gsCAT_TIML_PARENTLOTID & " = '" & Me.txtLotID_History.Text & "' "
        End If
        
        'Current 只查 DeleteFlag <> 'Y' 的資料.
        If ssTabQuery.TabIndex = 0 Then
            sSQL = sSQL & " and nvl(a." & gsCAT_TIML_DELETEFLAG & ", 'N') <> 'Y'"
        End If
        
    End If
    'Added by Jack on 2021/01/05 for CP測前併批專案 <End>.

    'Marked by Jack on 2021/01/05 for CP測前併批專案 <Start>.
    'sSQL = "select  a." & gsCAT_TIML_PARENTLOTID & " " & _
    '"  ,a." & gsCAT_TIML_CHILDLOTID & " " & _
    '"  ,a." & gsCAT_TIML_WAFERID & " " & _
    '"  ,a." & gsCAT_TIML_WQTY & " " & _
    '"  ,a." & gsCAT_TIML_CQTY & " " & _
    '"  ,a." & gsCAT_TIML_CREATEUSERID & " " & _
    '"  ,a." & gsCAT_TIML_CREATETIME & " " & _
    '"  ,a." & gsCAT_TIML_IPN & " " & _
    '"  ,a." & gsCAT_TIML_OWNER & " " & _
    '"  ,a." & gsCAT_TIML_SOURCE & " " & _
    '"  ,a." & gsCAT_TIML_ERRMSG & " "
    
    'sSQL = sSQL & " from " & gsCAT_TBL_INKLESS_MERGE_LIST & " a "
        
    'sSQL = sSQL & " WHERE a." & gsCAT_TIML_DELETEFLAG & " = 'N' "

    'If Trim(Me.txtLotID.Text) <> "" Then
    '    sSQL = sSQL & " and a." & gsCAT_TIML_PARENTLOTID & " = '" & Me.txtLotID.Text & "' "
    'End If
    'Marked by Jack on 2021/01/05 for CP測前併批專案 <End>.
    
    Set colRS = moProRawSql.QueryDatabase(sSQL)
    
    'Added by Jack on 2021/01/05 for CP測前併批專案 <Start>.
    If colRS Is Nothing Or colRS.Count = 0 Then
        'Modified by Jack on 2021/03/03 for UAT bug fix (兩頁查詢使用不同欄位LotID)
        If ssTabQuery.Tab = 0 Then
            UtShowMsgBox "Lot '" & Me.txtLotID.Text & "' is no data found !!" & vbNewLine & _
                         "Lot '" & Me.txtLotID.Text & "' 查無資料 !!"
        Else
            UtShowMsgBox "Lot '" & Me.txtLotID_History.Text & "' is no data found !!" & vbNewLine & _
                         "Lot '" & Me.txtLotID_History.Text & "' 查無資料 !!"
        End If
        GoTo ExitHandler
    End If
    'Added by Jack on 2021/01/05 for CP測前併批專案 <End>.
    
    With Me.spdLotInfo
        For lIndex = 1 To colRS.Count
            .MaxRows = .MaxRows + 1
            .SetText miSpdLotInfo_ParentLotID, .MaxRows, colRS.Item(lIndex).Item(gsCAT_TIML_PARENTLOTID)
            .SetText miSpdLotInfo_ChildLotID, .MaxRows, colRS.Item(lIndex).Item(gsCAT_TIML_CHILDLOTID)
            .SetText miSpdLotInfo_WaferID, .MaxRows, colRS.Item(lIndex).Item(gsCAT_TIML_WAFERID)
            .SetText miSpdLotInfo_WQty, .MaxRows, colRS.Item(lIndex).Item(gsCAT_TIML_WQTY)
            .SetText miSpdLotInfo_CQty, .MaxRows, colRS.Item(lIndex).Item(gsCAT_TIML_CQTY)
            .SetText miSpdLotInfo_UserID, .MaxRows, colRS.Item(lIndex).Item(gsCAT_TIML_CREATEUSERID)
            .SetText miSpdLotInfo_MergeTime, .MaxRows, colRS.Item(lIndex).Item(gsCAT_TIML_CREATETIME)
            .SetText miSpdLotInfo_IPN, .MaxRows, colRS.Item(lIndex).Item(gsCAT_TIML_IPN)
            .SetText miSpdLotInfo_Owner, .MaxRows, colRS.Item(lIndex).Item(gsCAT_TIML_OWNER)
            .SetText miSpdLotInfo_Source, .MaxRows, colRS.Item(lIndex).Item(gsCAT_TIML_SOURCE)
            .SetText miSpdLotInfo_ErrMsg, .MaxRows, colRS.Item(lIndex).Item(gsCAT_TIML_ERRMSG)
            
            'Added by Jack on 2021/01/05 for CP測前併批專案 <Start>.
            .SetText miSpdLotInfo_SLotNO, .MaxRows, colRS.Item(lIndex).Item(gsCAT_TIML_SLOTNO)
            .SetText miSpdLotInfo_MergeStepId, .MaxRows, colRS.Item(lIndex).Item(gsCAT_TVMG_MERGESTEPID)
            .SetText miSpdLotInfo_DeleteTime, .MaxRows, colRS.Item(lIndex).Item(gsCAT_TIML_DELETETIME)
            .SetText miSpdLotInfo_MiddleLotId, .MaxRows, colRS.Item(lIndex).Item(gsCAT_TIML_MIDDLELOTID)
            .SetText miSpdLotInfo_WaferChipQty, .MaxRows, colRS.Item(lIndex).Item(gsCAT_TIML_WAFERCHIPQTY)
            .SetText miSpdLotInfo_FabCreateTime, .MaxRows, colRS.Item(lIndex).Item(gsCAT_TIML_FABCREATETIME)
            .SetText miSpdLotInfo_RiskLevel, .MaxRows, colRS.Item(lIndex).Item(gsCAT_TIML_RISKLEVEL)
            .SetText miSpdLotInfo_RowId, .MaxRows, colRS.Item(lIndex).Item("rowid")
            .SetText miSpdLotInfo_Table, .MaxRows, colRS.Item(lIndex).Item("updatetable")
            'Added by Jack on 2021/01/05 for CP測前併批專案 <End>.
        Next
    End With
    'Set up spd
    With Me.spdLotInfo
'        .Col = 1
'        Do While .Col <= .MaxCols
'            .ColWidth(.Col) = .MaxTextColWidth(.Col) + 1
'            .Col = .Col + 1
'        Loop
        
        'Sort it
        .UserColAction = UserColActionSort
        Call .SetCellBorder(0, 0, .MaxCols, .MaxRows, 15, -1, CellBorderStyleSolid)
        
    End With
    
    Call ResizeSpdColumn(Me.spdLotInfo, moAppLog)

'----
' Done
'----

ExitHandler:
    ' NOTE 1:
    ' MUST CALL GetErrInfo() here first before another action
    Call GetErrInfo(msMODULE_ID, sProcID, typErrInfo, Erl)
    Call LogProcOut(msMODULE_ID, sProcID, typErrInfo, moAppLog)
    ' <Your cleaning up codes goes here...>
    Screen.MousePointer = vbDefault
ErrorHandler:
    If typErrInfo.lErrNumber Then
        ' NOTE 2:
        ' If you have custom handling of some Errors, please
        ' UN-REMARED the following Select Case block!
        ' Also, modify if neccessarily!!!
        '---- Start of Select Case Block ----
        Select Case typErrInfo.lErrNumber
            Case glERR_INVALIDOBJECT
                ' Retry code goes here...
            Case Else
                typErrInfo.sUserText = "Fail to execute application, please call IT support!!" & vbCrLf & _
                                        "程式執行失敗, 請洽IT人員處理"
            End Select
        '---- Start of Select Case Block ----
        On Error GoTo ExitHandler:
        Call HandleError(False, typErrInfo, , moAppLog, True)
    End If
End Sub

Private Sub ExecQuery_SorterControl(ByVal sAction As String)
On Error GoTo ExitHandler:
    Dim sProcID             As String
    Dim typErrInfo          As tErrInfo

    '----
' Init
'----
    sProcID = "ExecQuery_SorterControl"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog)

'----
' Condition Checking
'----
    
    If Trim(txtLotID_MergeSplit.Text) = "" Then
            UtShowMsgBox "Pls input Query Criterial !!" & vbNewLine & "請輸入查詢條件 !!"
            GoTo ExitHandler
    End If
    

'----
' Action
'----
    Dim sSQL As String
    Dim colRS As Collection
    Dim colRS_2 As Collection
    
    Dim lIndex As Long
    
    Screen.MousePointer = vbHourglass

    Me.spdSorterControl.MaxRows = 0
    
    sSQL = "select " & gsCAT_TSTC_ACTION & ", " & gsCAT_TSTC_REFLOTID & ", " & gsCAT_TSTC_REFSLOTNO & ", " & gsCAT_TSTC_NEWLOTID & ", " & _
                 " " & gsCAT_TSTC_NEWSLOTNO & ", " & gsCAT_TSTC_ORIGINALLOTID & ", " & gsCAT_TSTC_WAFERID & ", " & gsCAT_TSTC_SORTERSTART & ", " & _
                 " " & gsCAT_TSTC_SORTERCOMPLETE & " , " & gsCAT_TSTC_CREATEUSERID & ", " & gsCAT_TSTC_CREATETIME & " " & _
            " from " & gsCAT_TBL_SORTER_CONTROL & " s " & _
           " where s." & gsCAT_TSTC_REFLOTID & " = '" & txtLotID_MergeSplit.Text & "' " & _
             " and s." & gsCAT_TSTC_DELETEFLAG & " = 'N' " & _
             " and s." & gsCAT_TSTC_ACTION & " = '" & sAction & "' " & _
             " and s.createtime in (select distinct " & gsCAT_TSTC_CREATETIME & " " & _
                                    " from " & gsCAT_TBL_SORTER_CONTROL & " s " & _
                                   " where s." & gsCAT_TSTC_REFLOTID & " = '" & txtLotID_MergeSplit.Text & "' " & _
                                     " and s." & gsCAT_TSTC_DELETEFLAG & " = 'N' " & _
                                     " and s." & gsCAT_TSTC_SORTERCOMPLETE & " = 'N' " & _
                                     " and s." & gsCAT_TSTC_ACTION & " = '" & sAction & "' ) " & _
           " order by createtime, REFSLOTNO "
    
    Set colRS = moProRawSql.QueryDatabase(sSQL)
    
    If colRS Is Nothing Or colRS.Count = 0 Then
        sSQL = "select * " & _
                " from " & gsCAT_TBL_SORTER_CONTROL & " s " & _
               " where s." & gsCAT_TSTC_REFLOTID & " = '" & txtLotID_MergeSplit.Text & "'" & _
               " and s." & gsCAT_TSTC_ACTION & " = '" & sAction & "' "
        
        Set colRS_2 = moProRawSql.QueryDatabase(sSQL)
        If colRS_2.Count > 0 Then
            UtShowMsgBox "Lot '" & Me.txtLotID_MergeSplit.Text & "' has finish Sorter '" & sAction & "' operation !!" & vbNewLine & _
                         "Lot '" & Me.txtLotID_MergeSplit.Text & "' 已完成 Sorter '" & sAction & "' 作業 !!"
        Else
            UtShowMsgBox "Lot '" & Me.txtLotID_MergeSplit.Text & "' is no '" & sAction & "' data found !!" & vbNewLine & _
                         "Lot '" & Me.txtLotID_MergeSplit.Text & "' 查無 '" & sAction & "' 資料 !!"
        End If
        GoTo ExitHandler
    End If
    
    With Me.spdSorterControl
        For lIndex = 1 To colRS.Count
            .MaxRows = .MaxRows + 1
            .SetText miSpdSorterControl_Action, .MaxRows, colRS.Item(lIndex).Item(gsCAT_TSTC_ACTION)
            .SetText miSpdSorterControl_RefLotId, .MaxRows, colRS.Item(lIndex).Item(gsCAT_TSTC_REFLOTID)
            .SetText miSpdSorterControl_RefSLotNo, .MaxRows, colRS.Item(lIndex).Item(gsCAT_TSTC_REFSLOTNO)
            .SetText miSpdSorterControl_NewLotId, .MaxRows, colRS.Item(lIndex).Item(gsCAT_TSTC_NEWLOTID)
            .SetText miSpdSorterControl_NewSLotNo, .MaxRows, colRS.Item(lIndex).Item(gsCAT_TSTC_NEWSLOTNO)
            .SetText miSpdSorterControl_OriginalLotId, .MaxRows, colRS.Item(lIndex).Item(gsCAT_TSTC_ORIGINALLOTID)
            .SetText miSpdSorterControl_WaferId, .MaxRows, colRS.Item(lIndex).Item(gsCAT_TSTC_WAFERID)
            .SetText miSpdSorterControl_SorterStart, .MaxRows, colRS.Item(lIndex).Item(gsCAT_TSTC_SORTERSTART)
            .SetText miSpdSorterControl_SorterComplete, .MaxRows, colRS.Item(lIndex).Item(gsCAT_TSTC_SORTERCOMPLETE)
            .SetText miSpdSorterControl_CreateUserId, .MaxRows, colRS.Item(lIndex).Item(gsCAT_TSTC_CREATEUSERID)
            .SetText miSpdSorterControl_CreateTime, .MaxRows, colRS.Item(lIndex).Item(gsCAT_TSTC_CREATETIME)
        Next
    End With

    With Me.spdSorterControl
        'Sort it
        .UserColAction = UserColActionSort
        Call .SetCellBorder(0, 0, .MaxCols, .MaxRows, 15, -1, CellBorderStyleSolid)
        Call .Sort(1, 1, .MaxCols, .MaxRows, SortByRow, Array(miSpdSorterControl_CreateTime, miSpdSorterControl_RefSLotNo), Array(1, 1))
    End With
    
    Call ResizeSpdColumn(Me.spdSorterControl, moAppLog)

'----
' Done
'----

ExitHandler:
    ' NOTE 1:
    ' MUST CALL GetErrInfo() here first before another action
    Call GetErrInfo(msMODULE_ID, sProcID, typErrInfo, Erl)
    Call LogProcOut(msMODULE_ID, sProcID, typErrInfo, moAppLog)
    ' <Your cleaning up codes goes here...>
    Screen.MousePointer = vbDefault
ErrorHandler:
    If typErrInfo.lErrNumber Then
        ' NOTE 2:
        ' If you have custom handling of some Errors, please
        ' UN-REMARED the following Select Case block!
        ' Also, modify if neccessarily!!!
        '---- Start of Select Case Block ----
        Select Case typErrInfo.lErrNumber
            Case glERR_INVALIDOBJECT
                ' Retry code goes here...
            Case Else
                typErrInfo.sUserText = "Fail to execute application, please call IT support!!" & vbCrLf & _
                                        "程式執行失敗, 請洽IT人員處理"
            End Select
        '---- Start of Select Case Block ----
        On Error GoTo ExitHandler:
        Call HandleError(False, typErrInfo, , moAppLog, True)
    End If
End Sub

namespace MesApi.Models
{
    // API 接收的參數
    public class QuerySorterControlRequest
    {
        public string LotId { get; set; } // 對應 txtLotID_MergeSplit.Text
        public string Action { get; set; } // 對應 sAction 參數
    }

    // 查詢回傳的資料結構
    public class SorterControlInfo
    {
        public string Action { get; set; }
        public string RefLotId { get; set; }
        public string RefSlotNo { get; set; }
        public string NewLotId { get; set; }
        public string NewSlotNo { get; set; }
        public string OriginalLotId { get; set; }
        public string WaferId { get; set; }
        public string SorterStart { get; set; }
        public string SorterComplete { get; set; }
        public string CreateUserId { get; set; }
        public string CreateTime { get; set; }
    }
}

using Microsoft.AspNetCore.Mvc;
using MesApi.Models;
using MesApi.Services;

namespace MesApi.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class LotController : ControllerBase
    {
        private readonly ILotService _lotService;

        public LotController(ILotService lotService)
        {
            _lotService = lotService;
        }

        // ... 之前寫的 ExecQuery [HttpGet("query")] ...

        // 新增：處理 Sorter Control 的查詢
        [HttpGet("sorter-control")]
        public async Task<IActionResult> ExecQuerySorterControl([FromQuery] QuerySorterControlRequest request)
        {
            // Condition Checking
            if (string.IsNullOrWhiteSpace(request.LotId) || string.IsNullOrWhiteSpace(request.Action))
            {
                return BadRequest(new { Message = "Pls input Query Criterial !! 請輸入查詢條件 !!" });
            }

            try
            {
                var result = await _lotService.GetSorterControlInfoAsync(request.LotId, request.Action);
                return Ok(result);
            }
            catch (InvalidOperationException ex)
            {
                // 攔截 Service 層拋出的特定商業邏輯例外 (對應 VB6 的 UtShowMsgBox)
                // 回傳 HTTP 400 Bad Request 給前端，並附帶錯誤訊息
                return BadRequest(new { Message = ex.Message });
            }
            catch (Exception ex)
            {
                // 系統發生預期外的錯誤
                return StatusCode(500, new { Message = "Fail to execute application, please call IT support!! 程式執行失敗, 請洽IT人員處理" });
            }
        }
    }
}

using MesApi.Models;
using MesApi.Repositories;

namespace MesApi.Services
{
    public interface ILotService
    {
        // ... 前一個方法 ...
        Task<IEnumerable<SorterControlInfo>> GetSorterControlInfoAsync(string lotId, string action);
    }

    public class LotService : ILotService
    {
        private readonly ILotRepository _lotRepo;

        public LotService(ILotRepository lotRepo)
        {
            _lotRepo = lotRepo;
        }

        // ... 前一個方法 ...

        public async Task<IEnumerable<SorterControlInfo>> GetSorterControlInfoAsync(string lotId, string action)
        {
            // 第一次查詢：查詢尚未完成的 Sorter 作業
            var data = await _lotRepo.GetActiveSorterControlsAsync(lotId, action);

            // 若第一段 SQL (colRS) 沒有資料
            if (data == null || !data.Any())
            {
                // 第二次查詢 (colRS_2)：檢查是否其實已經完成了？
                bool hasFinished = await _lotRepo.CheckAnySorterControlExistsAsync(lotId, action);

                if (hasFinished)
                {
                    // 對應 MsgBox: Has finish Sorter operation
                    throw new InvalidOperationException($"Lot '{lotId}' has finish Sorter '{action}' operation !! 已完成 Sorter '{action}' 作業 !!");
                }
                else
                {
                    // 對應 MsgBox: is no data found
                    throw new InvalidOperationException($"Lot '{lotId}' is no '{action}' data found !! 查無 '{action}' 資料 !!");
                }
            }

            return data;
        }
    }
}

using MesApi.Models;
using Dapper;
using System.Data;

namespace MesApi.Repositories
{
    public interface ILotRepository
    {
        // ... 前面的方法 ...
        Task<IEnumerable<SorterControlInfo>> GetActiveSorterControlsAsync(string lotId, string action);
        Task<bool> CheckAnySorterControlExistsAsync(string lotId, string action);
    }

    public class LotRepository : ILotRepository
    {
        private readonly IDbConnection _dbConnection;

        public LotRepository(IDbConnection dbConnection)
        {
            _dbConnection = dbConnection;
        }

        // ... 前面的方法 ...

        public async Task<IEnumerable<SorterControlInfo>> GetActiveSorterControlsAsync(string lotId, string action)
        {
            string sql = @"
                SELECT 
                    ACTION as Action, REFLOTID as RefLotId, REFSLOTNO as RefSlotNo, 
                    NEWLOTID as NewLotId, NEWSLOTNO as NewSlotNo, ORIGINALLOTID as OriginalLotId, 
                    WAFERID as WaferId, SORTERSTART as SorterStart, SORTERCOMPLETE as SorterComplete, 
                    CREATEUSERID as CreateUserId, CREATETIME as CreateTime
                FROM TBL_SORTER_CONTROL s
                WHERE s.REFLOTID = @LotId 
                  AND s.DELETEFLAG = 'N' 
                  AND s.ACTION = @Action 
                  AND s.CREATETIME IN (
                      SELECT DISTINCT CREATETIME 
                      FROM TBL_SORTER_CONTROL 
                      WHERE REFLOTID = @LotId 
                        AND DELETEFLAG = 'N' 
                        AND SORTERCOMPLETE = 'N' 
                        AND ACTION = @Action
                  )
                ORDER BY CREATETIME, REFSLOTNO";

            return await _dbConnection.QueryAsync<SorterControlInfo>(sql, new { LotId = lotId, Action = action });
        }

        public async Task<bool> CheckAnySorterControlExistsAsync(string lotId, string action)
        {
            // 對應第二段 colRS_2 的查詢，我們只需要知道「是否存在」，所以用 SELECT 1 即可，效能更好。
            string sql = @"
                SELECT 1 
                FROM TBL_SORTER_CONTROL 
                WHERE REFLOTID = @LotId 
                  AND ACTION = @Action";
            
            var result = await _dbConnection.QueryFirstOrDefaultAsync<int?>(sql, new { LotId = lotId, Action = action });
            return result != null;
        }
    }
}

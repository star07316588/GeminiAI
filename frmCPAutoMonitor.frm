Option Explicit

Private Const msMODULE_ID As String = "frmCPAutoMonitor"
'**********
' Variable Declaration
'**********
Private moCwMbx     As Object
Private moFwCAT     As Object
Private moFwMDL     As Object
Private moFwOPR     As Object
Private moFwPRP     As Object
Private moFwWF      As Object
Private moFwWIP     As Object
Private moProRawSql As Object
Private moAppLog    As Object

Private msRuleName  As String
Private msLotId     As String
Private msEqpId     As String

Private mbShowQuery As Boolean  'Add by Weilun on 20180808 for ReqNo.BE#201800559

'狀態檢查
Private Const msEnd                                 As String = "LotEnd"
Private Const msComplete                            As String = "LotComplete"
Private Const msFail                                As String = "Fail"
Private Const msExec                                As String = "Exec"
Private Const msTimeOut                             As String = "TimeOut"

'排序標籤
Private Const miLevel_Error                         As Integer = 1
Private Const miLevel_InitTimeOut                   As Integer = 2
Private Const miLevel_ExecTimeOut                   As Integer = 3
Private Const miLevel_MoveOutTimeOut                As Integer = 4
Private Const miLevel_SortRunNoCPAuto               As Integer = 5

Private Const miLevel_AutoExec                      As Integer = 98
Private Const miLevel_ManualExec                    As Integer = 99

Private Const miSPDMONITOR_MAXCOLS                  As Integer = 13

Private Const miSPDMONITOR_PRIORITY                 As Integer = 1
Private Const miSPDMONITOR_EQID                     As Integer = 2
Private Const miSPDMONITOR_LOTID                    As Integer = 3
Private Const miSPDMONITOR_STEPNAME                 As Integer = 4
Private Const miSPDMONITOR_ACTION                   As Integer = 5
Private Const miSPDMONITOR_STEPNO                   As Integer = 6
Private Const miSPDMONITOR_LOTSTATUS                As Integer = 7
Private Const miSPDMONITOR_CURRENTRULE              As Integer = 8
Private Const miSPDMONITOR_TDSTESTSUMTIME           As Integer = 9
Private Const miSPDMONITOR_CPAUTOSTATUSTIME         As Integer = 10
Private Const miSPDMONITOR_LOTENDSTATUS             As Integer = 11
Private Const miSPDMONITOR_LOTCOMPLETESTATUS        As Integer = 12
Private Const miSPDMONITOR_COMMENTS                 As Integer = 13

Private Const msSPDMONITOR_PRIORITY                 As String = "Priority"
Private Const msSPDMONITOR_EQID                     As String = "EqID"
Private Const msSPDMONITOR_LOTID                    As String = "LotID"
Private Const msSPDMONITOR_STEPNAME                 As String = "StepName"
Private Const msSPDMONITOR_ACTION                   As String = "Action"
Private Const msSPDMONITOR_STEPNO                   As String = "StepNo"
Private Const msSPDMONITOR_LOTSTATUS                As String = "LotStatus"
Private Const msSPDMONITOR_CURRENTRULE              As String = "CurrentRule"
Private Const msSPDMONITOR_TDSTESTSUMTIME           As String = "TdsTestSumTime"
Private Const msSPDMONITOR_CPAUTOSTATUSTIME         As String = "CPAutoStatusTime"
Private Const msSPDMONITOR_LOTENDSTATUS             As String = "LotEndStatus"
Private Const msSPDMONITOR_LOTCOMPLETESTATUS        As String = "LotCompleteStatus"
Private Const msSPDMONITOR_COMMENTS                 As String = "Comments"


Private miResult    As Integer
Public Property Let LotID(sLotId As String)
    On Error Resume Next
    msLotId = sLotId
End Property
Public Property Let EqpId(sEqpid As String)
    On Error Resume Next
    msEqpId = sEqpid
End Property
Public Property Set FwCATControl(oFwCAT As Object)
    On Error Resume Next
    Set moFwCAT = oFwCAT
End Property
Public Property Set FwWFControl(oFwWF As Object)
    On Error Resume Next
    Set moFwWF = oFwWF
End Property
Public Property Set FwOPRControl(oFwOPR As Object)
    On Error Resume Next
    Set moFwOPR = oFwOPR
End Property
Public Property Set FwMDLControl(oFwMDL As Object)
    On Error Resume Next
    Set moFwMDL = oFwMDL
End Property
Public Property Set FwPRPControl(oFwPRP As Object)
    On Error Resume Next
    Set moFwPRP = oFwPRP
End Property
Public Property Set FwWIPControl(oFwWIP As Object)
    On Error Resume Next
    Set moFwWIP = oFwWIP
End Property
Public Property Set CwMbxControl(oCwMbx As Object)
    On Error Resume Next
    Set moCwMbx = oCwMbx
End Property
Public Property Set ProRawSqlControl(oProRawSqlControl As Object)
    On Error Resume Next
    Set moProRawSql = oProRawSqlControl
End Property
Public Property Set MainTraceLog(oLogCtrl As Object)
    On Error Resume Next
    Set moAppLog = oLogCtrl
End Property

Public Property Let RuleName(sRuleName As String)
    On Error Resume Next
    msRuleName = sRuleName
End Property
Private Sub ResetFwControls()
'**************************************************
'**************************************************
    On Error Resume Next
    Set moCwMbx = Nothing
    Set moFwCAT = Nothing
    Set moFwMDL = Nothing
    Set moFwOPR = Nothing
    Set moFwPRP = Nothing
    Set moFwWF = Nothing
    Set moFwWIP = Nothing
    Set moProRawSql = Nothing
    Set moAppLog = Nothing

End Sub
Public Property Get Result() As Integer
'**************************************************
'**************************************************
    On Error Resume Next
    Result = miResult

End Property

Private Sub cmdCancel_Click()
    On Error Resume Next
    miResult = vbCancel
    Me.Hide
End Sub

'Add by Weilun on 20180808 for ReqNo.BE#201800559
Private Sub Form_Activate()
    On Error Resume Next
    If mbShowQuery = True Then
        Call cmdRefresh_Click
        mbShowQuery = False
    End If
End Sub

Private Sub Form_Load()
    On Error Resume Next
    miResult = vbCancel
    Call ResetFwControls
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Call ResetFwControls
End Sub




'================================================================================
' Function: Init()
'--------------------------------------------------------------------------------
' Description:  <Type your function description here...>
'--------------------------------------------------------------------------------
' Author:       Vencent Wei, CIT 2002/04/14
'--------------------------------------------------------------------------------
' Parameters:
'................................................................................
' ARGUMENTS
'   oLogCtrl            (I) [Optional] Valid reference to a object object
'   Argument2           (I) <Description goes here...>
'   Argument3           (O) <Description goes here...>
'
' NON-LOCAL ARGS
'   NonLoc1         (I) <Description goes here...>
'   NonLoc2         (I) <Description goes here...>
'
' RETURN TYPE
'   Boolean         (R) True/False
'
'--------------------------------------------------------------------------------
' Revision History:
'................................................................................
' [REV 01] <AuthorName>, <CompanyName> <YYYY/MM/DD>
' 1) <Description goes here...>
'    <Line 2...>
'
' [REV 02] <AuthorName>, <CompanyName> <YYYY/MM/DD>
' 1) <Description goes here...>
'    <Line 2...>
'
'================================================================================
Public Function Init() As Boolean
On Error GoTo ExitHandler:
Dim sProcID As String
Dim typErrInfo As tErrInfo

'----
' Init
'----
    sProcID = "Init"
    Init = True
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog) '"Entering Function...", oLogCtrl, glLOG_PROC, msMODULE_ID, sProcID)

'----
' Condition Checking
'----
    ' <Put your condition checking codes here>...

'----
' Action
'----
    ' <Put your Action codes here>...
    Call InitSpdMonitor
    
    'Modify by Weilun on 20180808 for ReqNo.BE#201800559 <start>
    '預先Query改為顯示後執行
    mbShowQuery = True
    'Call cmdRefresh_Click
    'Modify by Weilun on 20180808 for ReqNo.BE#201800559 <end>
'----
' Done
'----

ExitHandler:
    ' NOTE 1:
    ' MUST CALL GetErrInfo() here first before another action
    Call GetErrInfo(msMODULE_ID, sProcID, typErrInfo, Erl)
    Call LogProcOut(msMODULE_ID, sProcID, typErrInfo, moAppLog)
    ' <Your cleaning up codes goes here...>

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
        Init = False
        On Error GoTo ExitHandler:
        Call HandleError(True, typErrInfo, , moAppLog)
    End If
End Function

'================================================================================
' Sub: InitSpdMonitor()
'--------------------------------------------------------------------------------
' Description:  <Type your Sub description here...>
'--------------------------------------------------------------------------------
' Author:       Weilun Huang,MXIC 2018-04-20
'--------------------------------------------------------------------------------
' Parameters:
'................................................................................
' ARGUMENTS
'   oLogCtrl            (I) [Optional] Valid reference to a Object object
'   Argument2           (I) <Description goes here...>
'   Argument3           (O) <Description goes here...>
'
' NON-LOCAL ARGS
'   NonLoc1         (I) <Description goes here...>
'   NonLoc2         (I) <Description goes here...>
'
'--------------------------------------------------------------------------------
' Revision History:
'................................................................................
' [REV 01] <AuthorName>, <CompanyName> <YYYY/MM/DD>
' 1) <Description goes here...>
'    <Line 2...>
'
' [REV 02] <AuthorName>, <CompanyName> <YYYY/MM/DD>
' 1) <Description goes here...>
'    <Line 2...>
'
'================================================================================
Private Sub InitSpdMonitor()
On Error GoTo ExitHandler:
Dim sProcID As String
Dim typErrInfo As tErrInfo
Dim lCol As Long
'----
' Init
'----
    sProcID = "InitSpdMonitor"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog)

'----
' Condition Chking
'----
    ' <Put your condition checking codes here>...

'----
' Action
'----
    With Me.spdMonitor
        .MaxRows = 0
        .MaxCols = miSPDMONITOR_MAXCOLS
                
        .SetText miSPDMONITOR_EQID, 0, msSPDMONITOR_EQID
        .SetText miSPDMONITOR_LOTID, 0, msSPDMONITOR_LOTID
        .SetText miSPDMONITOR_STEPNO, 0, msSPDMONITOR_STEPNO
        .SetText miSPDMONITOR_STEPNAME, 0, msSPDMONITOR_STEPNAME
        .SetText miSPDMONITOR_LOTSTATUS, 0, msSPDMONITOR_LOTSTATUS
        .SetText miSPDMONITOR_CURRENTRULE, 0, msSPDMONITOR_CURRENTRULE
        .SetText miSPDMONITOR_TDSTESTSUMTIME, 0, msSPDMONITOR_TDSTESTSUMTIME
        .SetText miSPDMONITOR_ACTION, 0, msSPDMONITOR_ACTION
        .SetText miSPDMONITOR_CPAUTOSTATUSTIME, 0, msSPDMONITOR_CPAUTOSTATUSTIME
        .SetText miSPDMONITOR_LOTENDSTATUS, 0, msSPDMONITOR_LOTENDSTATUS
        .SetText miSPDMONITOR_LOTCOMPLETESTATUS, 0, msSPDMONITOR_LOTCOMPLETESTATUS
        .SetText miSPDMONITOR_COMMENTS, 0, msSPDMONITOR_COMMENTS
        .SetText miSPDMONITOR_PRIORITY, 0, msSPDMONITOR_PRIORITY
        
        .Col = miSPDMONITOR_PRIORITY
        .colHidden = True

        For lCol = 1 To .MaxCols
            Call LockSpdColumn(Me.spdMonitor, lCol)
        Next lCol
        
        .Col = -1
        .Row = -1
        .BackColor = vbWhite

        Call ResizeSpdColumn(Me.spdMonitor, moAppLog)
    End With
'----
' Done
'----

ExitHandler:
    ' NOTE 1:
    ' MUST CALL GetErrInfo() here first before another action
    Call GetErrInfo(msMODULE_ID, sProcID, typErrInfo, Erl)
    Call LogProcOut(msMODULE_ID, sProcID, typErrInfo, moAppLog)
    ' <Your cleaning up codes goes here...>
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

'================================================================================
' Sub: cmdRefresh_Click()
'--------------------------------------------------------------------------------
' Description:  <Type your Sub description here...>
'--------------------------------------------------------------------------------
' Author:       Weilun Huang,MXIC 2018-04-20
'--------------------------------------------------------------------------------
' Parameters:
'................................................................................
' ARGUMENTS
'   oLogCtrl            (I) [Optional] Valid reference to a Object object
'   Argument2           (I) <Description goes here...>
'   Argument3           (O) <Description goes here...>
'
' NON-LOCAL ARGS
'   NonLoc1         (I) <Description goes here...>
'   NonLoc2         (I) <Description goes here...>
'
'--------------------------------------------------------------------------------
' Revision History:
'................................................................................
' [REV 01] <AuthorName>, <CompanyName> <YYYY/MM/DD>
' 1) <Description goes here...>
'    <Line 2...>
'
' [REV 02] <AuthorName>, <CompanyName> <YYYY/MM/DD>
' 1) <Description goes here...>
'    <Line 2...>
'
'================================================================================
Private Sub cmdRefresh_Click()
On Error GoTo ExitHandler:
Dim sProcID As String
Dim typErrInfo As tErrInfo

Dim oLot                As FwLot
Dim sLotId              As String
Dim lIdx                As Long
Dim vPriority           As Variant

Dim sSQL                As String
Dim colRaws             As Collection
Dim oRaws               As Collection
Dim colTime             As Collection
Dim colData             As Collection
Dim sStatusList         As String
Dim sLotEndStatus       As String
Dim sLotCompleteStatus  As String

Dim bSearchSum          As Boolean
Dim sMaxUpdateTime      As String
Dim sUpdateTime         As String
Dim sMainStatus         As String
Dim sRuleName           As String

Dim sCheckResult        As String
Dim sScheduleTime       As String
Dim sCurrentTime        As String
Dim sExpiration         As String
Dim sExecutingLimit     As String
Dim sMoveOutLimit       As String

'進度條
Dim vPgbStage_1        As Variant
Dim vPgbStage_2        As Variant
Dim vPgbStage_3        As Variant

'----
' Init
'----
    sProcID = "cmdRefresh_Click"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog)
    
    sExecutingLimit = "1/144"   '10分鐘
    sMoveOutLimit = "1/288"     '5分鐘
    sExpiration = "1/720"       '2分鐘
    
    '減1避免超出範圍
    vPgbStage_1 = (pgbMonitor.Max - 1) * (1 / 3)
    vPgbStage_2 = (pgbMonitor.Max - 1) * (1 / 3)
    vPgbStage_3 = (pgbMonitor.Max - 1) * (1 / 3)
    
'----
' Condition Chking
'----
    ' <Put your condition checking codes here>...

'----
' Action
'----
    'Debug.Print
    
    '避免連點
    Me.cmdRefresh.Enabled = False 'Add by Weilun on 20180830 for ReqNo.BE#201800559
    
    With Me.spdMonitor
        DoEvents
        pgbMonitor.Value = 0
        .Visible = False
        .MaxRows = 0
        
        'Action欄位先解除Mutiline設定
        .Row = -1
        .Col = miSPDMONITOR_ACTION
        .CellType = CellTypeStaticText
        .TypeEditMultiLine = False '要配合CellTypeEdit
        .Lock = False
        
        '第一部份: 搜TBL_CP_AUTO_COMPLETE資料進行Query
        sStatusList = gsCPAUTO_STATUS_INIT & "," & _
                      gsCPAUTO_STATUS_AUTO_CHECKFAIL & "," & _
                      gsCPAUTO_STATUS_AUTO_ERROR & "," & _
                      gsCPAUTO_STATUS_AUTO_EXEC & "," & _
                      gsCPAUTO_STATUS_AUTO_FAIL & "," & _
                      gsCPAUTO_STATUS_MANUAL_EXEC
    
        sStatusList = "'" & Replace(sStatusList, ",", "','") & "'"
        
        'Modify by Weilun on 20180808 for ReqNo.BE#201800559
        '加入比對Lot_Attribute資訊, 濾掉不合時宜的資訊
        sSQL = " select TCAC." & gsCAT_TCAC_LOTID & ", " & _
                      " TCAC." & gsCAT_TCAC_STEPNAME & ", " & _
                      " TCAC." & gsCAT_TCAC_COMMENTS & ", " & _
                      " TCAC." & gsCAT_TCAC_LOTENDSTATUS & ", " & _
                      " TCAC." & gsCAT_TCAC_LOTCOMPLETESTATUS & ", " & _
                      " TCAC." & gsCAT_TCAC_CREATETIME & ", " & _
                      " TCAC." & gsCAT_TCAC_UPDATETIME & " "
                      
        sSQL = sSQL & _
                 " from " & gsCAT_TBL_CP_AUTO_COMPLETE & " TCAC , " & _
                            gsCAT_TBL_LOT_ATTRIBUTE & " TLATT "
                            
        sSQL = sSQL & _
                " where ( TCAC." & gsCAT_TCAC_LOTID & ", " & _
                        " TCAC." & gsCAT_TCAC_STEPNAME & ", " & _
                        " TCAC." & gsCAT_TCAC_CREATETIME & ") in " & _
                      " (select " & gsCAT_TCAC_LOTID & ", " & _
                                    gsCAT_TCAC_STEPNAME & ", " & _
                          " Max(" & gsCAT_TCAC_CREATETIME & ") " & _
                         " from " & gsCAT_TBL_CP_AUTO_COMPLETE & " " & _
                        " where " & gsCAT_TCAC_DELETEFLAG & " = 'N' " & _
                     " group by " & gsCAT_TCAC_LOTID & ", " & gsCAT_TCAC_STEPNAME & ") " & _
                  " and ( TCAC." & gsCAT_TCAC_LOTENDSTATUS & " in (" & sStatusList & ")" & _
                     " or TCAC." & gsCAT_TCAC_LOTCOMPLETESTATUS & " in (" & sStatusList & "))" & _
                  " and TCAC." & gsCAT_TCAC_DELETEFLAG & " = 'N' " & _
                  " and TCAC." & gsCAT_TCAC_LOTID & " = TLATT." & gsCAT_TLATT_LOTID & " " & _
                  " and TCAC." & gsCAT_TCAC_STEPNAME & " = TLATT." & gsCAT_TLATT_STEPNAME & " " & _
                  " and TLATT." & gsCAT_TLATT_LOTSTATUS1 & " in ('" & gsLOTSTATUS_RUN & "','" & _
                                                                      gsLOTSTATUS_ERUN & "','" & _
                                                                      gsLOTSTATUS_BREAK & "') " & _
             " order by TCAC." & gsCAT_TCAC_CREATETIME & " "
                  
'        sSQL = " select " & gsCAT_TCAC_LOTID & ", " & _
'                            gsCAT_TCAC_STEPNAME & ", " & _
'                            gsCAT_TCAC_COMMENTS & ", " & _
'                            gsCAT_TCAC_LOTENDSTATUS & ", " & _
'                            gsCAT_TCAC_LOTCOMPLETESTATUS & ", " & _
'                            gsCAT_TCAC_CREATETIME & ", " & _
'                            gsCAT_TCAC_UPDATETIME & " " & _
'                 " from " & gsCAT_TBL_CP_AUTO_COMPLETE & " " & _
'                " where (" & gsCAT_TCAC_LOTID & ", " & _
'                             gsCAT_TCAC_STEPNAME & ", " & _
'                             gsCAT_TCAC_CREATETIME & ") in " & _
'                      " (select " & gsCAT_TCAC_LOTID & ", " & _
'                                    gsCAT_TCAC_STEPNAME & ", " & _
'                          " Max(" & gsCAT_TCAC_CREATETIME & ") " & _
'                         " from " & gsCAT_TBL_CP_AUTO_COMPLETE & " " & _
'                        " where " & gsCAT_TCAC_DELETEFLAG & " = 'N' " & _
'                     " group by " & gsCAT_TCAC_LOTID & ", " & gsCAT_TCAC_STEPNAME & ") " & _
'                  " and ( " & gsCAT_TCAC_LOTENDSTATUS & " in (" & sStatusList & ")" & _
'                     " or " & gsCAT_TCAC_LOTCOMPLETESTATUS & " in (" & sStatusList & "))" & _
'                  " and " & gsCAT_TCAC_DELETEFLAG & " = 'N' " & _
'                  " order by " & gsCAT_TCAC_CREATETIME & " "
        'Modify by Weilun on 20180808 for ReqNo.BE#201800559
        
        Set colRaws = moProRawSql.QueryDatabase(sSQL)
        If colRaws.Count > 0 Then
            For Each oRaws In colRaws
                pgbMonitor.Value = pgbMonitor.Value + (1 / colRaws.Count * vPgbStage_1)
                DoEvents
                
                sCheckResult = ""
                sRuleName = ""
                sMaxUpdateTime = ""
                sMainStatus = ""
                
                Set oLot = moFwWIP.LotById(oRaws.Item(gsCAT_TCAC_LOTID))
                
                If Not oLot Is Nothing Then
                    bSearchSum = False
                    sLotEndStatus = oRaws.Item(gsCAT_TCAC_LOTENDSTATUS)
                    sLotCompleteStatus = oRaws.Item(gsCAT_TCAC_LOTCOMPLETESTATUS)
                    
                    'LotCompleteStatus優先, 並取得狀態
                    If sLotCompleteStatus <> "" Then
                        sMainStatus = sLotCompleteStatus
                        sRuleName = msComplete
                    Else
                        sMainStatus = sLotEndStatus
                        sRuleName = msEnd
                    End If
                    
                    '狀態分辨
                    '狀態1: 執行中
                    If sMainStatus = gsCPAUTO_STATUS_AUTO_EXEC Or _
                       sMainStatus = gsCPAUTO_STATUS_MANUAL_EXEC Then
                        sCheckResult = msExec
                        
                    '狀態2: 錯誤與失敗
                    ElseIf sMainStatus = gsCPAUTO_STATUS_AUTO_FAIL Or _
                           sMainStatus = gsCPAUTO_STATUS_AUTO_CHECKFAIL Or _
                           sMainStatus = gsCPAUTO_STATUS_AUTO_ERROR Then
                        sCheckResult = msFail
                       
                    '狀態3: 資料到齊已超過特定時間(sExpiration), 卻尚未自動化(狀態還是Init)
                    ElseIf sMainStatus = gsCPAUTO_STATUS_INIT Then
                        If CheckTdsTestSum(oLot, sMaxUpdateTime) = True Then
                            sSQL = " select to_char(to_date(SubStr('" & sMaxUpdateTime & "', 1, 15), 'YYYYMMDD HH24MISS') + " & sExpiration & ", " & _
                                                  " 'YYYYMMDD HH24MISS') || '000' as ScheduleTime, " & _
                                          " to_char(sysdate, 'YYYYMMDD HH24MISS') || '000' as CurrentTime " & _
                                     " From Dual "
                            Set colTime = moProRawSql.QueryDatabase(sSQL)
                            If colTime.Count = 1 Then
                                sScheduleTime = colTime.Item(1).Item("ScheduleTime")
                                sCurrentTime = colTime.Item(1).Item("CurrentTime")
                                If sCurrentTime > sScheduleTime Then
                                    sCheckResult = msTimeOut
                                End If
                            End If
                        End If
                    End If
                    
                    'Show資料, 經過處理後, 應該只有Init卻未TimeOut的不會進
                    If sCheckResult = msTimeOut Or _
                       sCheckResult = msExec Or _
                       sCheckResult = msFail Then
                        
                        .MaxRows = .MaxRows + 1
                        .SetText miSPDMONITOR_EQID, .MaxRows, oLot.CustomAttributes(gsLOT_CUSTOMATTR_CUR_EQP_ID)
                        .SetText miSPDMONITOR_LOTID, .MaxRows, oLot.Id
                        .SetText miSPDMONITOR_STEPNO, .MaxRows, oLot.CurrentStep.Steps.Item(1).Id
                        .SetText miSPDMONITOR_STEPNAME, .MaxRows, oLot.CurrentStep.Steps.Item(1).Description
                        .SetText miSPDMONITOR_LOTSTATUS, .MaxRows, oLot.CustomAttributes(gsLOT_CUSTOMATTR_STATUS)
                        .SetText miSPDMONITOR_CURRENTRULE, .MaxRows, oLot.CurrentStep.Steps.Item(1).CurrentRule.Id
                        .SetText miSPDMONITOR_LOTENDSTATUS, .MaxRows, oRaws.Item(gsCAT_TCAC_LOTENDSTATUS)
                        .SetText miSPDMONITOR_LOTCOMPLETESTATUS, .MaxRows, oRaws.Item(gsCAT_TCAC_LOTCOMPLETESTATUS)
                        .SetText miSPDMONITOR_COMMENTS, .MaxRows, oRaws.Item(gsCAT_TCAC_COMMENTS)
                        
                        If oRaws.Item(gsCAT_TCAC_UPDATETIME) <> "" Then
                            .SetText miSPDMONITOR_CPAUTOSTATUSTIME, .MaxRows, oRaws.Item(gsCAT_TCAC_UPDATETIME)
                        Else
                            .SetText miSPDMONITOR_CPAUTOSTATUSTIME, .MaxRows, oRaws.Item(gsCAT_TCAC_CREATETIME)
                        End If

                        '失敗的狀況
                        If sCheckResult = msFail Then
                            .SetText miSPDMONITOR_ACTION, .MaxRows, _
                                     "Auto " & sRuleName & " Fail, 請手動執行" & sRuleName & ". " & vbNewLine & _
                                     "Please execute " & sRuleName & "."
                            '優先度
                            .SetText miSPDMONITOR_PRIORITY, .MaxRows, miLevel_Error
                        End If

                        '資料已到齊並超過預定時間尚未自動化
                        If sCheckResult = msTimeOut Then
                            .SetText miSPDMONITOR_ACTION, .MaxRows, _
                                     "資料已到齊但是尚未Auto " & sRuleName & ", 請手動執行" & sRuleName & "." & vbNewLine & _
                                     "Please execute " & sRuleName & "."
                            .SetText miSPDMONITOR_TDSTESTSUMTIME, .MaxRows, sMaxUpdateTime
                            '優先度
                            .SetText miSPDMONITOR_PRIORITY, .MaxRows, miLevel_InitTimeOut
                        End If
                        
                        '執行中的狀況
                        If sCheckResult = msExec Then
                            If sMainStatus = gsCPAUTO_STATUS_AUTO_EXEC Then
                                'CPAuto已超過預定時間(sExecutingLimit), 卻還在執行中, 請介入
                                sUpdateTime = oRaws.Item(gsCAT_TCAC_UPDATETIME)
                                sSQL = " select to_char(to_date(SubStr('" & sUpdateTime & "', 1, 15), 'YYYYMMDD HH24MISS') + " & sExecutingLimit & ", " & _
                                                      " 'YYYYMMDD HH24MISS') || '000' as ScheduleTime, " & _
                                              " to_char(sysdate, 'YYYYMMDD HH24MISS') || '000' as CurrentTime " & _
                                         " From Dual "
                                Set colTime = moProRawSql.QueryDatabase(sSQL)
                                If colTime.Count = 1 Then
                                    sScheduleTime = colTime.Item(1).Item("ScheduleTime")
                                    sCurrentTime = colTime.Item(1).Item("CurrentTime")
                                    If sCurrentTime > sScheduleTime Then
                                        .SetText miSPDMONITOR_ACTION, .MaxRows, _
                                                 "Auto " & sRuleName & " 執行過久, 請手動介入." & vbNewLine & _
                                                 "Please execute " & sRuleName & "."
                                        '優先度
                                        .SetText miSPDMONITOR_PRIORITY, .MaxRows, miLevel_ExecTimeOut
                                    Else
                                        .SetText miSPDMONITOR_ACTION, .MaxRows, _
                                                 "Auto " & sRuleName & " 執行中, 請稍候." & vbNewLine & _
                                                 "Keep Monitor."
                                        '優先度
                                        .SetText miSPDMONITOR_PRIORITY, .MaxRows, miLevel_AutoExec
                                    End If
                                End If
                            Else
                                .SetText miSPDMONITOR_ACTION, .MaxRows, _
                                         "Manual " & sRuleName & " 執行中, 請注意是否執行過久. " & vbNewLine & _
                                         "Keep Monitor."
                                         
                                '優先度
                                .SetText miSPDMONITOR_PRIORITY, .MaxRows, miLevel_ManualExec
                            End If
                        End If
                    End If
                End If
                Set oLot = Nothing
            Next oRaws
        Else
            pgbMonitor.Value = pgbMonitor.Value + vPgbStage_1
            DoEvents
        End If

        '第二部份: MoveOut檢查
        sSQL = " select  t." & gsCAT_TLATT_LOTID & ", " & _
                      " cp." & gsCAT_TCAC_LOTENDSTATUS & ", " & _
                      " cp." & gsCAT_TCAC_LOTCOMPLETESTATUS & ", " & _
                      " cp." & gsCAT_TCAC_COMMENTS & ", " & _
                      " cp." & gsCAT_TCAC_UPDATETIME & " "
       
        sSQL = sSQL & _
                 " from " & gsCAT_TBL_LOT_ATTRIBUTE & " t, " & _
                          " fwlot                a, " & _
                          " fwlot_n2m            b, " & _
                          " fwwipstep            c, " & _
                          " fwwipstep_n2m        d, " & _
                          " fwwipsteprule        e, " & _
                            gsCAT_TBL_CP_AUTO_COMPLETE & " cp "
                          
        sSQL = sSQL & _
                " where t." & gsCAT_TLATT_STAGE & " = '" & gsSTAGE_WS & "' " & _
                  " and t." & gsCAT_TLATT_LOTSTATUS1 & " in ('" & gsLOTSTATUS_WAITD & "', " & _
                                                           " '" & gsLOTSTATUS_BREAK & "') " & _
                  " and t." & gsCAT_TLATT_RESOURCETYPE & " = '" & gsRESOURCE_WTESTER & "' " & _
                  " and t." & gsCAT_TLATT_LOTID & " = cp." & gsCAT_TCAC_LOTID & " " & _
                  " and t." & gsCAT_TLATT_STEPNAME & " = cp." & gsCAT_TCAC_STEPNAME & " " & _
                  " and (cp." & gsCAT_TCAC_LOTID & ", " & _
                       " cp." & gsCAT_TCAC_STEPNAME & ", " & _
                       " cp." & gsCAT_TCAC_CREATETIME & ") in " & _
                      " (select subcp." & gsCAT_TCAC_LOTID & ", " & _
                              " subcp." & gsCAT_TCAC_STEPNAME & ", " & _
                              " Max(subcp." & gsCAT_TCAC_CREATETIME & ") " & _
                         " from " & gsCAT_TBL_CP_AUTO_COMPLETE & " subcp " & _
                        " where subcp." & gsCAT_TCAC_DELETEFLAG & " = 'N' " & _
                     " group by subcp." & gsCAT_TCAC_LOTID & ", " & _
                              " subcp." & gsCAT_TCAC_STEPNAME & ") "
                              
        sSQL = sSQL & _
                  " and cp." & gsCAT_TCAC_LOTCOMPLETESTATUS & " in ('" & gsCPAUTO_STATUS_AUTO_OK & "', " & _
                                                                  " '" & gsCPAUTO_STATUS_MANUAL_OK & "') " & _
                  " and t." & gsCAT_TLATT_LOTID & " = a.appid " & _
                  " and a.sysid = b.fromid " & _
                  " and b.toid = c.sysid " & _
                  " and c.sysid = d.fromid " & _
                  " and d.toid = e.sysid " & _
                  " and c.currentruleindex = d.sequence " & _
                  " and e.rulename = 'MoveOut' "
                                 
        Set colRaws = moProRawSql.QueryDatabase(sSQL)
        If colRaws.Count > 0 Then
            For Each oRaws In colRaws
                pgbMonitor.Value = pgbMonitor.Value + (1 / colRaws.Count * vPgbStage_2)
                DoEvents
            
                Set oLot = moFwWIP.LotById(oRaws.Item(gsCAT_TLATT_LOTID))
            
                If Not oLot Is Nothing Then
                    'LotComplete已超過預定時間(sMoveOutLimit), 卻還在MoveOut
                    sUpdateTime = oRaws.Item(gsCAT_TCAC_UPDATETIME)
                    sSQL = " select to_char(to_date(SubStr('" & sUpdateTime & "', 1, 15), 'YYYYMMDD HH24MISS') + " & sMoveOutLimit & ", " & _
                                          " 'YYYYMMDD HH24MISS') || '000' as ScheduleTime, " & _
                                  " to_char(sysdate, 'YYYYMMDD HH24MISS') || '000' as CurrentTime " & _
                             " From Dual "
                    Set colTime = moProRawSql.QueryDatabase(sSQL)
                    If colTime.Count = 1 Then
                        sScheduleTime = colTime.Item(1).Item("ScheduleTime")
                        sCurrentTime = colTime.Item(1).Item("CurrentTime")
                        If sCurrentTime > sScheduleTime Then
                            .MaxRows = .MaxRows + 1
                            .SetText miSPDMONITOR_EQID, .MaxRows, oLot.CustomAttributes(gsLOT_CUSTOMATTR_CUR_EQP_ID)
                            .SetText miSPDMONITOR_LOTID, .MaxRows, oLot.Id
                            .SetText miSPDMONITOR_STEPNO, .MaxRows, oLot.CurrentStep.Steps.Item(1).Id
                            .SetText miSPDMONITOR_STEPNAME, .MaxRows, oLot.CurrentStep.Steps.Item(1).Description
                            .SetText miSPDMONITOR_LOTSTATUS, .MaxRows, oLot.CustomAttributes(gsLOT_CUSTOMATTR_STATUS)
                            .SetText miSPDMONITOR_CURRENTRULE, .MaxRows, oLot.CurrentStep.Steps.Item(1).CurrentRule.Id
                            .SetText miSPDMONITOR_LOTENDSTATUS, .MaxRows, oRaws.Item(gsCAT_TCAC_LOTENDSTATUS)
                            .SetText miSPDMONITOR_LOTCOMPLETESTATUS, .MaxRows, oRaws.Item(gsCAT_TCAC_LOTCOMPLETESTATUS)
                            .SetText miSPDMONITOR_CPAUTOSTATUSTIME, .MaxRows, oRaws.Item(gsCAT_TCAC_UPDATETIME)
                            .SetText miSPDMONITOR_COMMENTS, .MaxRows, oRaws.Item(gsCAT_TCAC_COMMENTS)
    
                            .SetText miSPDMONITOR_ACTION, .MaxRows, _
                                     "MoveOut超時，請確認CommMgm系統重工記錄." & vbNewLine & _
                                     "若排除上述異常，請手動執行強制MoveOut或Reposition Rule." & vbNewLine & _
                                     "Please execute move out forcibly after confirm CommMgm system ""rework"" record were done."

                            '優先度
                            .SetText miSPDMONITOR_PRIORITY, .MaxRows, miLevel_MoveOutTimeOut
                        End If
                    End If
                End If
                Set oLot = Nothing
                
            Next oRaws
        Else
            pgbMonitor.Value = pgbMonitor.Value + vPgbStage_2
            DoEvents
        End If

        '第三部份: 自動化紀錄初始化失敗查詢
        '狀態加入ERUN, Add by Weilun on 20180628 for fix
        sSQL = " select tlatt." & gsCAT_TLATT_LOTID & " " & _
                 " from " & gsCAT_TBL_LOT_ATTRIBUTE & " tlatt " & _
                " where tlatt." & gsCAT_TLATT_STAGE & " = '" & gsSTAGE_WS & "' " & _
                  " and tlatt." & gsCAT_TLATT_LOTSTATUS1 & " in ('" & gsLOTSTATUS_RUN & "', " & _
                                                                "'" & gsLOTSTATUS_ERUN & "') " & _
                  " and tlatt." & gsCAT_TLATT_RESOURCETYPE & " = '" & gsRESOURCE_WTESTER & "' "
                  
        Set colRaws = moProRawSql.QueryDatabase(sSQL)
        If colRaws.Count > 0 Then
            For Each oRaws In colRaws
                pgbMonitor.Value = pgbMonitor.Value + (1 / colRaws.Count * vPgbStage_3)
                DoEvents
                Set oLot = moFwWIP.LotById(oRaws.Item(gsCAT_TLATT_LOTID))
            
                '時間比對加入等於, Add by Weilun on 20180628 for fix
                If Not oLot Is Nothing Then
                    If oLot.CurrentStep.Steps.Item(1).CurrentRule.Id = "SortRun" Then '接下來的查詢是歷史紀錄, 還是得確認目前的Rule
                        sSQL = " select SRTime.lotid " & _
                                  " from (select A.WIPID as Lotid, " & _
                                               " max(A.Txntimestamp) as SortRunTime " & _
                                          " from FWWIPHISTORY B, " & _
                                               " FWWIPTRANSACTION A " & _
                                         " where B.Grouphistkey like 'SortRun-%' " & _
                                           " and A.SYSID = B.WIPTXN " & _
                                           " and A.WIPID = '" & oLot.Id & "' " & _
                                         " group by A.WIPID) SRTime, " & _
                                         gsCAT_TBL_CP_AUTO_COMPLETE & " cp " & _
                                 " where SRTime.lotid = cp." & gsCAT_TCAC_LOTID & " " & _
                                   " and SRTime.SortRunTime <= cp." & gsCAT_TCAC_CREATETIME & " "
                        
                        Set colData = moProRawSql.QueryDatabase(sSQL)
                        If colData.Count = 0 Then
                            .MaxRows = .MaxRows + 1
                            .SetText miSPDMONITOR_EQID, .MaxRows, oLot.CustomAttributes(gsLOT_CUSTOMATTR_CUR_EQP_ID)
                            .SetText miSPDMONITOR_LOTID, .MaxRows, oLot.Id
                            .SetText miSPDMONITOR_STEPNO, .MaxRows, oLot.CurrentStep.Steps.Item(1).Id
                            .SetText miSPDMONITOR_STEPNAME, .MaxRows, oLot.CurrentStep.Steps.Item(1).Description
                            .SetText miSPDMONITOR_LOTSTATUS, .MaxRows, oLot.CustomAttributes(gsLOT_CUSTOMATTR_STATUS)
                            .SetText miSPDMONITOR_CURRENTRULE, .MaxRows, oLot.CurrentStep.Steps.Item(1).CurrentRule.Id
    
                            .SetText miSPDMONITOR_ACTION, .MaxRows, _
                                     "沒有對應的SortRun自動化紀錄，請重新LotStart或手動執行後續作業." & vbNewLine & _
                                     "Please execute LotEnd and LotComplete."
    
                            '優先度
                            .SetText miSPDMONITOR_PRIORITY, .MaxRows, miLevel_SortRunNoCPAuto
                        End If
                    End If
                End If
            Next oRaws
        Else
            pgbMonitor.Value = pgbMonitor.Value + vPgbStage_3
            DoEvents
        End If
    
        '進行排版
        If .MaxRows > 0 Then
            '排序
            .Sort 1, 1, .MaxCols, .MaxRows, SortByRow, miSPDMONITOR_PRIORITY, 1
        
            '自動調整寬度
            Call ResizeSpdColumn(Me.spdMonitor, moAppLog)
            
            'Action欄位Mutiline設定
            .Row = -1
            .Col = miSPDMONITOR_ACTION
            .CellType = CellTypeEdit
            .TypeEditMultiLine = True '要配合CellTypeEdit
            .Lock = True

            '設定Action寬度以及自動調整高度
            .ColWidth(miSPDMONITOR_ACTION) = 50
            Call ResizeSpdRow(Me.spdMonitor, moAppLog)
            
            '標示
            For lIdx = 1 To .MaxRows
            
                .GetText miSPDMONITOR_PRIORITY, lIdx, vPriority
                If vPriority = miLevel_Error Or _
                   vPriority = miLevel_SortRunNoCPAuto Or _
                   vPriority = miLevel_InitTimeOut Or _
                   vPriority = miLevel_ExecTimeOut Or _
                   vPriority = miLevel_MoveOutTimeOut Then
                    .Row = lIdx
                    .Col = miSPDMONITOR_ACTION
                    .ForeColor = vbRed
                End If
            Next lIdx
        End If
        
        .Visible = True
    End With
    
    Me.cmdRefresh.Enabled = True 'Add by Weilun on 20180830 for ReqNo.BE#201800559
'----
' Done
'----

ExitHandler:
    ' NOTE 1:
    ' MUST CALL GetErrInfo() here first before another action
    Call GetErrInfo(msMODULE_ID, sProcID, typErrInfo, Erl)
    Call LogProcOut(msMODULE_ID, sProcID, typErrInfo, moAppLog)
    ' <Your cleaning up codes goes here...>
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
        
        '出問題也得把查詢狀況顯示, 雖然排版可能有問題
        Me.spdMonitor.Visible = True
        Me.cmdRefresh.Enabled = True 'Add by Weilun on 20180830 for ReqNo.BE#201800559
    End If
End Sub

'================================================================================
' Sub: CheckTdsTestSum()
'--------------------------------------------------------------------------------
' Description:  <Type your Sub description here...>
'--------------------------------------------------------------------------------
' Author:       Ivan, CIT 2002-03-20
'--------------------------------------------------------------------------------
' Parameters:
'................................................................................
' ARGUMENTS
'   oLogCtrl            (I) [Optional] Valid reference to a Object object
'   Argument2           (I) <Description goes here...>
'   Argument3           (O) <Description goes here...>
'
' NON-LOCAL ARGS
'   NonLoc1         (I) <Description goes here...>
'   NonLoc2         (I) <Description goes here...>
'
'--------------------------------------------------------------------------------
' Revision History:
'................................................................................
' [REV 01] <AuthorName>, <CompanyName> <YYYY/MM/DD>
' 1) <Description goes here...>
'    <Line 2...>
'
' [REV 02] <AuthorName>, <CompanyName> <YYYY/MM/DD>
' 1) <Description goes here...>
'    <Line 2...>
'
'================================================================================
Private Function CheckTdsTestSum(oLot As FwLot, ByRef sMaxUpdateTime As String) As Boolean
On Error GoTo ExitHandler:
Dim sProcID                 As String
Dim typErrInfo              As tErrInfo

Dim sSQL                    As String
Dim colRaws                 As Collection
Dim oRaws                   As FwStrings
Dim sTable                  As String
Dim sColumn                 As String
Dim sWhere                  As String
Dim sOther                  As String

Dim sParentLot              As String
Dim oComp                   As FwComponent
Dim iIndex                  As Integer
Dim sTimeHereSince          As String
Dim sStepNo                 As String
Dim sUpdateTime             As String
Dim sTdsQty                 As String
Dim sReviseQty              As String
Dim sReasonCode             As String
Dim bIsWaferIdInTdsSum      As Boolean
Dim sWaferID                As String
'----
' Init
'----
    sProcID = "CheckTdsTestSum"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog)

    sTimeHereSince = ConvertToString(oLot.CurrentStep.Steps.Item(1).TimeHereSince, moAppLog)
    sStepNo = oLot.CurrentStep.Steps.Item(1).Id
    sMaxUpdateTime = ""

    If InStr(1, oLot.Id, ".") > 0 Then
        sParentLot = Left(Left(oLot.Id, InStr(1, oLot.Id, ".") - 1), 8)
    Else
        sParentLot = Left(oLot.Id, 8)
    End If
'----
' Condition Checking
'----
    ' <Put your condition checking codes here>...

'----
' Action
'----

    'Define table name
    sTable = gsCAT_TBL_WS_TDS_SUM
    
    'Define column name
    sColumn = gsCAT_TWTS_WAFER_IDN & "," & _
              gsCAT_TWTS_TDS_GOOD_DIEN & "," & _
              gsCAT_TWTS_REVISE_GOOD_DIEN & "," & _
              gsCAT_TWTS_REASON_CODE & "," & _
              gsCAT_TWTS_CREATE_TIME & "," & _
              gsCAT_TWTS_UPDATE_TIME

    'Define where
    sWhere = "(" & gsCAT_TWTS_TEST_MODE & "," & gsCAT_TWTS_PARENT_LOT_ID & "," & _
                gsCAT_TWTS_WAFER_IDN & "," & gsCAT_TWTS_STEPNO & "," & gsCAT_TWTS_INSTEPTIME & "," & _
                "nvl(" & gsCAT_TWTS_TIME_STAMP & ",'NULL'))" & _
             " IN" & _
             " (SELECT " & gsCAT_TWTS_TEST_MODE & "," & gsCAT_TWTS_PARENT_LOT_ID & "," & _
                gsCAT_TWTS_WAFER_IDN & "," & gsCAT_TWTS_STEPNO & "," & gsCAT_TWTS_INSTEPTIME & _
                ",NVL(MAX(" & gsCAT_TWTS_TIME_STAMP & "),'NULL')" & _
             " FROM " & gsCAT_TBL_WS_TDS_SUM & _
             " WHERE " & _
            gsCAT_TWTS_INSTEPTIME & "='" & sTimeHereSince & "' and " & _
            gsCAT_TWTS_STEPNO & "='" & sStepNo & "' and " & _
            gsCAT_TWTS_PARENT_LOT_ID & "='" & sParentLot & "' and " & _
            gsCAT_TWTS_DELETE_FLAG & "='N'" & _
            " GROUP BY " & gsCAT_TWTS_TEST_MODE & "," & gsCAT_TWTS_PARENT_LOT_ID & "," & gsCAT_TWTS_WAFER_IDN & _
            "," & gsCAT_TWTS_STEPNO & "," & gsCAT_TWTS_INSTEPTIME & ")"

    sOther = "order by " & "nvl(" & gsCAT_TWTS_TIME_STAMP & ",' ') desc"
        
    Set colRaws = SelectRawSqlTxn(moAppLog, moFwWIP, moFwWF, moCwMbx, "CAT", sTable, sColumn, sWhere, sOther)
    If colRaws Is Nothing Then
        Call RaiseError(glERR_INVALIDOBJECT, _
                    FormatErrorText(gsETX_INVALIDOBJECT, "Collection"))
    End If
    
    For iIndex = 1 To 25
        bIsWaferIdInTdsSum = False
        sTdsQty = ""
        sReviseQty = ""
        sReasonCode = ""
        Set oComp = oLot.Components.Item(Format(iIndex, "00"))
        If Not oComp Is Nothing Then
            If oComp.Status <> modConstGlobal.gsCOMP_STATUS_SCRAPPED Then
                sWaferID = oComp.Id
                ' "SortRun" 條件 added by Jack 2010/02/10 for JC201000028.
                '若存在於 tbl_ws_tds_sum, 則將component ChipQty放入ReviseQty欄位,
                '否則放入TdsQty欄位.
                '需加一變數, check是否有存在?
                If colRaws.Count > 0 Then '此Lot是否有資料存在於 tbl_ws_tds_sum?
                    For Each oRaws In colRaws
                        If Trim(CStr(oRaws.Item(gsCAT_TWTS_WAFER_IDN))) = sWaferID Then
                            
                            bIsWaferIdInTdsSum = True
                            sTdsQty = oRaws.Item(gsCAT_TWTS_TDS_GOOD_DIEN)
                            sReviseQty = oRaws.Item(gsCAT_TWTS_REVISE_GOOD_DIEN)
                            sReasonCode = oRaws.Item(gsCAT_TWTS_REASON_CODE)
                            
                            '未到齊的情況:
                            '1.TdsQty和ReviseQty都空值
                            If (Not (Len(Trim(sTdsQty)) <> 0 Or _
                                     Len(Trim(sReviseQty)) <> 0)) Or _
                               (Len(Trim(sReviseQty)) <> 0 And Len(Trim(sReasonCode)) = 0) Then
                                GoTo ExitHandler 'False
                            End If
                            
                            If oRaws.Item(gsCAT_TWTS_UPDATE_TIME) <> "" Then
                                sUpdateTime = oRaws.Item(gsCAT_TWTS_UPDATE_TIME)
                            Else
                                sUpdateTime = oRaws.Item(gsCAT_TWTS_CREATE_TIME)
                            End If
                            
                            If sUpdateTime > sMaxUpdateTime Then
                                sMaxUpdateTime = sUpdateTime
                            End If
                            
                        End If
                    Next oRaws
                End If
                If Not bIsWaferIdInTdsSum Then 'added by Jack 2010/02/10 for JC201000028.
                    GoTo ExitHandler 'False
                End If
            End If
        End If
        Set oComp = Nothing
    Next iIndex
    
    CheckTdsTestSum = True
'----
' Done
'----

ExitHandler:
    ' NOTE 1:
    ' MUST CALL GetErrInfo() here first before another action
    Call GetErrInfo(msMODULE_ID, sProcID, typErrInfo, Erl)
    Call LogProcOut(msMODULE_ID, sProcID, typErrInfo, moAppLog)
    ' <Your cleaning up codes goes here...>
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
End Function

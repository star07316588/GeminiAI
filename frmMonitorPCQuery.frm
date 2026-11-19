Option Explicit

Private Const msMODULE_ID As String = "frmMonitorPCQuery"

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
Private msEngCostNo As String

Private miResult    As Integer

Private msType      As String

'**********
' Constant Declaration
'**********
'-----
' Spread spdLotList
'-----


Private Const msspdLotList_IP                    As String = "IP"
Private Const msspdLotList_AREA                  As String = "Area"
Private Const msspdLotList_MONITORSYSTEM         As String = "SystemName"
Private Const msspdLotList_LASTVERSION           As String = "Last Version Flag"
Private Const msspdLotList_VERSIONDETAIL         As String = "Version Detail"
Private Const msspdLotList_COMMENTS              As String = "Comment"


Private Const mispdLotList_IP                    As Integer = 1
Private Const mispdLotList_AREA                  As Integer = 2
Private Const mispdLotList_MONITORSYSTEM         As Integer = 3
Private Const mispdLotList_LASTVERSION           As Integer = 4
Private Const mispdLotList_VERSIONDETAIL         As Integer = 5
Private Const mispdLotList_COMMENTS              As Integer = 6

Private Const miSpdLotList_MaxCols               As Integer = 6

Public Property Let LotID(sLotId As String)
    On Error Resume Next
    msLotId = sLotId
End Property
Public Property Let EqpId(sEqpid As String)
    On Error Resume Next
    msEqpId = sEqpid
End Property
Public Property Let QueryType(sType As String)
    On Error Resume Next
    msType = sType
End Property

Public Property Let EngCostNo(sEngCostNo As String)
    On Error Resume Next
    msEngCostNo = sEngCostNo
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



Private Sub cmdClose_Click()
    On Error Resume Next
    miResult = vbCancel
    Me.Hide
End Sub

Private Sub cmdQuery_Click()
    Call ExecQuery
End Sub
'Add by Sam on 20200505 for ReqNo:202000138
Private Sub cmdStd_Click()
On Error GoTo ExitHandler:
Dim sProcID As String
Dim typErrInfo As tErrInfo

Dim ofrmMain As Object
Dim vIP As Variant
Dim vSystemName As Variant

'----
' Init
'----
    sProcID = "cmdStd_Click"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog)
'----
' Condition Checking
'----
    ' <Put your condition checking codes here>...

'----
' Action
'----
    Set ofrmMain = New frmMonitorPCQuery_Detail
    If (ofrmMain Is Nothing) Then
        Call RaiseError(glERR_INVALIDOBJECT, _
                        FormatErrorText(gsETX_INVALIDOBJECT, "Form"))
    End If
    Load ofrmMain
    With ofrmMain
       Set .CwMbxControl = moCwMbx
       Set .FwMDLControl = moFwMDL
       Set .FwWIPControl = moFwWIP
       Set .FwOPRControl = moFwOPR
       Set .FwPRPControl = moFwPRP
       Set .FwWFControl = moFwWF
       Set .FwCATControl = moFwCAT
       Set .MainTraceLog = moAppLog
       Set .CwMbxControl = moCwMbx
       Set .ProRawSqlControl = moProRawSql
           .RuleName = msRuleName
           .IP = "ALL"
           .SystemName = cboSystemName.Text
           .Init
           .Show vbModal
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

Private Sub Form_Load()
'**************************************************
'**************************************************
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
' Author:       Ivan, CIT 2001/01/14
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
Dim oLots      As FwLots
Dim oLot       As FwLot

Dim sSQL As String
Dim colRS As Collection
Dim lIdx As Long

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
    
    Call InitspdLotList
'    Call ExecQuery

    cboArea.Clear
    cboArea.AddItem ""
    sSQL = "select distinct " & gsCAT_TMSI_AREA & "  from " & _
           " " & gsCAT_TBL_MONITORPC_SYSTEMVERSION_IP & _
           " where " & gsCAT_TMSI_DELETEFLAG & " = 'N' "
    Set colRS = moProRawSql.QueryDatabase(sSQL)
    For lIdx = 1 To colRS.Count
        cboArea.AddItem colRS.Item(lIdx).Item(gsCAT_TMSI_AREA)
    Next
            
    cboSystemName.Clear
    cboSystemName.AddItem ""
    sSQL = "select distinct " & gsCAT_TMSI_MONITORSYSTEM & "  from " & _
           " " & gsCAT_TBL_MONITORPC_SYSTEMVERSION_IP & _
           " where " & gsCAT_TMSI_DELETEFLAG & " = 'N' "
    Set colRS = moProRawSql.QueryDatabase(sSQL)
    For lIdx = 1 To colRS.Count
        cboSystemName.AddItem colRS.Item(lIdx).Item(gsCAT_TMSI_MONITORSYSTEM)
    Next
            
            
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
' Function: InitspdLotList()
'--------------------------------------------------------------------------------
' Description:  <Type your function description here...>
'--------------------------------------------------------------------------------
' Author:       Ivan, CIT 2002-06-05
'--------------------------------------------------------------------------------
' Parameters:
'................................................................................
' ARGUMENTS
'   oLogCtrl            (I) [Optional] Valid reference to a clsLogTraceMsg object
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
Private Function InitspdLotList() As Boolean
On Error GoTo ExitHandler:
Dim sProcID As String
Dim typErrInfo As tErrInfo
Dim lCol       As Long
'----
' Init
'----
    sProcID = "InitspdLotList"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog) '"Entering Function...", moAppLog, glLOG_PROC, msMODULE_ID, sProcID)

'----
' Condition Checking
'----
    ' <Put your condition checking codes here>...

'----
' Action
'----
    With Me.spdLotList
        .MaxCols = miSpdLotList_MaxCols
        .MaxRows = 0
        
        Call .SetText(mispdLotList_IP, 0, msspdLotList_IP)
        Call .SetText(mispdLotList_AREA, 0, msspdLotList_AREA)
        Call .SetText(mispdLotList_MONITORSYSTEM, 0, msspdLotList_MONITORSYSTEM)
        Call .SetText(mispdLotList_LASTVERSION, 0, msspdLotList_LASTVERSION)
        Call .SetText(mispdLotList_VERSIONDETAIL, 0, msspdLotList_VERSIONDETAIL)
        Call .SetText(mispdLotList_COMMENTS, 0, msspdLotList_COMMENTS)
        
        .Col = mispdLotList_COMMENTS
        .Row = -1
        .CellType = CellTypeEdit
        .TypeMaxEditLen = 1024
        
        For lCol = 1 To .MaxCols
            Call LockSpdColumn(Me.spdLotList, lCol)
            .ColUserSortIndicator(lCol) = ColUserSortIndicatorAscending
            .UserColAction = UserColActionSort
        Next lCol
        
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
End Function

'================================================================================
' function: ExecQuery()
'--------------------------------------------------------------------------------
' Description:  <Type your function description here...>
'--------------------------------------------------------------------------------
' Author:       Sam, CIT 20090605
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
Public Sub ExecQuery()
On Error GoTo ExitHandler:
Dim sProcID As String
Dim typErrInfo As tErrInfo
Dim colRaws As Collection
Dim oRaws As Object
Dim sTable As String
Dim sColumn As String
Dim sWhere As String
Dim sOther As String

Dim lRow      As Long
Dim iCol      As Long

Dim sSQL As String
Dim sSql2 As String

Dim colMaxVersion As Collection
Dim colRS As Collection

Dim lIdx As Long
Dim sVersionFlag As String

'----
' Init
'----
    sProcID = "ExecQuery"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog)
    spdLotList.MaxRows = 0
'----
' Condition Checking
'----
    ' <Put your condition checking codes here>...

'----
' Action
'----
        
        
        sSQL = "select a." & gsCAT_TMSI_IP & " " & _
               " ,a." & gsCAT_TMSI_AREA & " " & _
               " ,a." & gsCAT_TMSI_MONITORSYSTEM & " " & _
               " ,a." & gsCAT_TMSI_COMMENTS & " " & _
               " from " & gsCAT_TBL_MONITORPC_SYSTEMVERSION_IP & " a " & _
               " where a." & gsCAT_TMSI_DELETEFLAG & "= 'N' "
        
        If Trim(cboArea.Text) <> "" Then
            sSQL = sSQL & " and a." & gsCAT_TMSI_AREA & " = '" & cboArea.Text & "' "
        End If
        
        If Trim(cboSystemName.Text) <> "" Then
            sSQL = sSQL & " and a." & gsCAT_TMSI_MONITORSYSTEM & " = '" & cboSystemName.Text & "' "
        End If
        
        Set colRaws = moProRawSql.QueryDatabase(sSQL)
        If colRaws Is Nothing Then
            Call RaiseError(glERR_INVALIDOBJECT, _
                        FormatErrorText(gsETX_INVALIDOBJECT, "Collection"))
        End If
        
        
        With Me.spdLotList
            For Each oRaws In colRaws
                .MaxRows = .MaxRows + 1
                sVersionFlag = "N"
                Call .SetText(mispdLotList_IP, .MaxRows, oRaws.Item(gsCAT_TMSI_IP))
                Call .SetText(mispdLotList_AREA, .MaxRows, oRaws.Item(gsCAT_TMSI_AREA))
                Call .SetText(mispdLotList_MONITORSYSTEM, .MaxRows, oRaws.Item(gsCAT_TMSI_MONITORSYSTEM))
                Call .SetText(mispdLotList_COMMENTS, .MaxRows, oRaws.Item(gsCAT_TMSI_COMMENTS))
                
                
                sSql2 = "select a." & gsCAT_TMS_FILENAME & ",max(a." & gsCAT_TMS_FILEUPDATETIME & ") as " & gsCAT_TMS_FILEUPDATETIME & "  " & _
                    " from " & gsCAT_TBL_MONITORPC_SYSTEMVERSION & " a," & gsCAT_TBL_MONITORPC_SYSTEMVERSION_IP & " b  " & _
                    " where a." & gsCAT_TMS_IP & " = b." & gsCAT_TMSI_IP & " " & _
                    " and a." & gsCAT_TMS_MONITORSYSTEM & " = b." & gsCAT_TMSI_MONITORSYSTEM & " " & _
                    " and a." & gsCAT_TMS_MONITORSYSTEM & " = '" & oRaws.Item(gsCAT_TMSI_MONITORSYSTEM) & "' " & _
                    " and a." & gsCAT_TMS_DELETEFLAG & " = 'N' " & _
                    " and b." & gsCAT_TMSI_DELETEFLAG & " = 'N' " & _
                    " group by a." & gsCAT_TMS_FILENAME & ""
                Set colMaxVersion = moProRawSql.QueryDatabase(sSql2)
                For lIdx = 1 To colMaxVersion.Count
'                    sVersionFlag = "N" 'Mark by Sam on 20200505 for ReqNO:202000138,移至下面,改為當有抓到檔案時再開始判斷
                    'Modify by Sam on 20200505 for ReqNo:202000138,將FileName條件改成不分大小寫
                    sSql2 = "select a." & gsCAT_TMS_FILENAME & ",a." & gsCAT_TMS_FILEUPDATETIME & " " & _
                        " from " & gsCAT_TBL_MONITORPC_SYSTEMVERSION & " a " & _
                        " where a." & gsCAT_TMS_IP & " = '" & oRaws.Item(gsCAT_TMSI_IP) & "' " & _
                        " and a." & gsCAT_TMS_MONITORSYSTEM & " = '" & oRaws.Item(gsCAT_TMSI_MONITORSYSTEM) & "' " & _
                        " and upper(a." & gsCAT_TMS_FILENAME & ") = '" & UCase(colMaxVersion.Item(lIdx).Item(gsCAT_TMS_FILENAME)) & "' " & _
                        " and a." & gsCAT_TMS_DELETEFLAG & " = 'N' "
                    Set colRS = moProRawSql.QueryDatabase(sSql2)
                    If colRS.Count > 0 Then
                        sVersionFlag = "N" 'Add by Sam on 20200505 for ReqNO:202000138,移至下面,改為當有抓到檔案再開始判斷 by Sam on 20200505 for ReqNO:202000138,移至下面,改為當有抓到檔案再開始判斷
                        If colRS.Item(1).Item(gsCAT_TMS_FILEUPDATETIME) >= colMaxVersion.Item(lIdx).Item(gsCAT_TMS_FILEUPDATETIME) Then
                            sVersionFlag = "Y"
                        Else
                            sVersionFlag = "N"
                            Exit For
                        End If
                    End If
                Next

                Call .SetText(mispdLotList_LASTVERSION, .MaxRows, sVersionFlag)

                
                .Col = mispdLotList_VERSIONDETAIL
                .Row = .MaxRows
                .CellType = CellTypeButton
                 .TypeButtonText = "Detail"
            Next oRaws
            Call .SetCellBorder(0, 0, .MaxCols, .MaxRows, 15, -1, CellBorderStyleSolid)
        End With
                
'        Call ResizeSpdRow(Me.spdLotList, moAppLog)
        Call ResizeSpdColumn(Me.spdLotList, moAppLog)
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


Private Sub spdLotList_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
On Error GoTo ExitHandler:
Dim sProcID As String
Dim typErrInfo As tErrInfo

Dim ofrmMain As Object
Dim vIP As Variant
Dim vSystemName As Variant

'----
' Init
'----
    sProcID = "spdLotList_ButtonClicked"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog)
'----
' Condition Checking
'----
    ' <Put your condition checking codes here>...

'----
' Action
'----
    Select Case Col
        Case mispdLotList_VERSIONDETAIL
            With Me.spdLotList
                .GetText mispdLotList_IP, Row, vIP
                .GetText mispdLotList_MONITORSYSTEM, Row, vSystemName
            End With
            Set ofrmMain = New frmMonitorPCQuery_Detail
            If (ofrmMain Is Nothing) Then
                Call RaiseError(glERR_INVALIDOBJECT, _
                                FormatErrorText(gsETX_INVALIDOBJECT, "Form"))
            End If
            Load ofrmMain
            With ofrmMain
               Set .CwMbxControl = moCwMbx
               Set .FwMDLControl = moFwMDL
               Set .FwWIPControl = moFwWIP
               Set .FwOPRControl = moFwOPR
               Set .FwPRPControl = moFwPRP
               Set .FwWFControl = moFwWF
               Set .FwCATControl = moFwCAT
               Set .MainTraceLog = moAppLog
               Set .CwMbxControl = moCwMbx
               Set .ProRawSqlControl = moProRawSql
                   .RuleName = msRuleName
                   .IP = CStr(vIP)
                   .SystemName = CStr(vSystemName)
                   .Init
                   .Show vbModal
            End With
        Case Else
    End Select
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

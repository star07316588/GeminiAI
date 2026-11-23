Option Explicit

Private Const msMODULE_ID As String = "frmYellowCardReSend"
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

Private Const msspdList_LotID             As String = "LotId"
Private Const msspdList_IPN               As String = "IPN"
Private Const msspdList_Qty               As String = "Qty"
Private Const msspdList_Location          As String = "Location"
Private Const msspdList_Datecode          As String = "Datecode"
Private Const msspdList_PackageName       As String = "PackageName"
Private Const msspdList_Bodysize          As String = "Bodysize"
Private Const msspdList_PackageCode       As String = "PackageCode"

Private Const mispdList_Check             As Integer = 1
Private Const miSpdList_LotID             As Integer = 2
Private Const miSpdList_IPN               As Integer = 3
Private Const mispdList_Qty               As Integer = 4
Private Const mispdList_Location          As Integer = 5
Private Const mispdList_Datecode          As Integer = 6
Private Const mispdList_PackageName       As Integer = 7
Private Const mispdList_Bodysize          As Integer = 8
Private Const mispdList_PackageCode       As Integer = 9

Private Const miSpdList_MAXCOLS          As Integer = 9

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


Private Sub cmdOk_Click()
On Error GoTo ExitHandler:
Dim sProcID     As String
Dim typErrInfo  As tErrInfo


Dim sSQL As String
Dim colRS As Collection
Dim lIdx As Long

Dim vLotID As Variant
Dim vChk As Variant

Dim vIPN As Variant
Dim vQty As Variant
Dim vLocation As Variant
Dim vDateCode As Variant
Dim vPkgName As Variant
Dim vBodySize As Variant
Dim vPkgCode As Variant
Dim sVendorNo As String
Dim sWQty As String

Dim oLot As FwLot

'----
' Init
'----
    sProcID = "cmdOK_Click"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog) '"Entering Function...", oLogCtrl, glLOG_PROC, msMODULE_ID, sProcID)

'----
' Condition Checking
'----
    ' <Put your condition checking codes here>...
    
    With spdList
        For lIdx = 1 To .MaxRows
            .GetText mispdList_Check, lIdx, vChk
            If Val(vChk) = 1 Then
                .GetText miSpdList_LotID, lIdx, vLotID
                Set oLot = FwuRetrieveLot(moFwWIP, CStr(vLotID), moAppLog)
                If oLot Is Nothing Then
                    UtShowMsgBox "LotId:" & vLotID & " 不存在."
                    GoTo ExitHandler
                End If
            End If
        Next
    End With
    
'----
' Action
'----

    With spdList
        For lIdx = 1 To .MaxRows
            .GetText mispdList_Check, lIdx, vChk
            If Val(vChk) = 1 Then
                .GetText miSpdList_LotID, lIdx, vLotID
                .GetText miSpdList_IPN, lIdx, vIPN
                .GetText mispdList_Qty, lIdx, vQty
                .GetText mispdList_Location, lIdx, vLocation
                .GetText mispdList_Datecode, lIdx, vDateCode
                .GetText mispdList_PackageName, lIdx, vPkgName
                .GetText mispdList_Bodysize, lIdx, vBodySize
                .GetText mispdList_PackageCode, lIdx, vPkgCode
                
                If Not (vPkgCode = "W" Or vPkgCode = "H") Then
                    sWQty = "0"
                Else
                    sWQty = CStr(vQty)
                End If
                
                sVendorNo = ""
'                sSQL = "select b." & gsCAT_TVM_VENDOR_NO & " from " & gsCAT_TBL_LOT_INFO & " a, " & gsCAT_TBL_VENDOR_MASTER & " b " & _
'                      " where a." & gsCAT_TLI_VENDOR & " = b." & gsCAT_TVM_VENDOR_NAME & " " & _
'                      " and a." & gsCAT_TLI_LOT_ID & " = '" & CStr(vLotID) & "' "
                '直接取Location的值轉NO
                sSQL = "select b." & gsCAT_TVM_VENDOR_NO & " from  " & gsCAT_TBL_VENDOR_MASTER & " b " & _
                      " where b." & gsCAT_TVM_DELETE_FLAG & " ='N' " & _
                      " and b." & gsCAT_TVM_VENDOR_NAME & " = '" & CStr(vLocation) & "' "
                Set colRS = moProRawSql.QueryDatabase(sSQL)
                If colRS.Count > 0 Then
                    sVendorNo = colRS.Item(1).Item(gsCAT_TVM_VENDOR_NO)
                End If
                
                If Trim(sVendorNo) = "" Then
                    sVendorNo = "0000101008"
                End If
                
                sSQL = "insert into " & gsCAT_TBL_YELLOW_CARD & " ( "
                sSQL = sSQL & " " & gsCAT_TYC_TKTYPE & " " & _
                              "," & gsCAT_TYC_NEWLOTNO & " " & _
                              "," & gsCAT_TYC_IPN & " " & _
                              "," & gsCAT_TYC_QTY & " " & _
                              "," & gsCAT_TYC_WQTY & " " & _
                              "," & gsCAT_TYC_TOVENDOR & " " & _
                              "," & gsCAT_TYC_DATECODE & " " & _
                              "," & gsCAT_TYC_PACKAGENAME & " " & _
                              "," & gsCAT_TYC_BODYSIZE & " " & _
                              "," & gsCAT_TYC_DATETIME & " " & _
                              "," & gsCAT_TYC_DATA_SOURCE & " " & _
                              "," & gsCAT_TYC_CREATEUSERID & " " & _
                              "," & gsCAT_TYC_CREATETIME & " "
                sSQL = sSQL & " ) values ( "
                
                sSQL = sSQL & " 'Manual' " & _
                              ",'" & CStr(vLotID) & "' " & _
                              ",'" & CStr(vIPN) & "' " & _
                              ",'" & CStr(vQty) & "' " & _
                              ",'" & sWQty & "' " & _
                              ",'" & sVendorNo & "' " & _
                              ",'" & CStr(vDateCode) & "' " & _
                              ",'" & CStr(vPkgName) & "' " & _
                              ",'" & CStr(vBodySize) & "' " & _
                              ",TO_CHAR(sysdate,'yyyymmdd') " & _
                              ",'Manual' " & _
                              ",'" & moFwOPR.ActiveUser.UserName & "' " & _
                              ",TO_CHAR(sysdate,'yyyymmdd hh24miss') || '000' "
                sSQL = sSQL & " ) "
                
                Call moProRawSql.QueryDatabase(sSQL)
            End If
        Next
    End With

    Me.Hide
    
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
        Call HandleError(True, typErrInfo, , moAppLog)
    End If

End Sub

Private Sub cmdSelect_Click(Index As Integer)
Dim oSpd As fpSpread
Dim lIdx As Long

    Set oSpd = spdList
    If oSpd Is Nothing Then
        Exit Sub
    End If
    
    With oSpd
        For lIdx = 1 To .MaxRows
            .Col = 1
            .Row = lIdx
            .Text = Index
        Next
    End With
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
Dim sProcID     As String
Dim typErrInfo  As tErrInfo
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
    Call mInitSpread
    Call spdList_Advance(True)
    
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
' Function: mInitSpread()
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
Public Function mInitSpread() As Boolean
On Error GoTo ExitHandler:
Dim sProcID     As String
Dim typErrInfo  As tErrInfo
Dim i           As Integer

'----
' Init
'----
    sProcID = "mInitSpread"
    mInitSpread = True
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog) '"Entering Function...", oLogCtrl, glLOG_PROC, msMODULE_ID, sProcID)

'----
' Condition Checking
'----
    ' <Put your condition checking codes here>...

'----
' Action
'----
    With spdList
        .MaxRows = 0
        .MaxCols = miSpdList_MAXCOLS
        .SetText mispdList_Check, 0, " "
        .SetText miSpdList_LotID, 0, msspdList_LotID
        .SetText miSpdList_IPN, 0, msspdList_IPN
        .SetText mispdList_Qty, 0, msspdList_Qty
        .SetText mispdList_Location, 0, msspdList_Location
        .SetText mispdList_Datecode, 0, msspdList_Datecode
        .SetText mispdList_PackageName, 0, msspdList_PackageName
        .SetText mispdList_Bodysize, 0, msspdList_Bodysize
        .SetText mispdList_PackageCode, 0, msspdList_PackageCode
        
        For i = 1 To .MaxCols
            .Col = i
            .Row = -1
            .Lock = True
            .BackColor = &HC0C0C0
            
            .ColUserSortIndicator(i) = ColUserSortIndicatorAscending
            .UserColAction = UserColActionSort
        Next
        
        .Col = mispdList_Check
        .CellType = CellTypeCheckBox
        .TypeCheckCenter = True
        .BackColor = &HC0C0C0
        .ColHeaderDisplay = DispBlank
        .Lock = False
        
        .Col = miSpdList_LotID
        .BackColor = &HFFFFFF
        .Lock = False
        
        .Col = mispdList_PackageCode
        .colHidden = True
        
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
        mInitSpread = False
        On Error GoTo ExitHandler:
        Call HandleError(True, typErrInfo, , moAppLog)
    End If
End Function


Private Sub spdList_Advance(ByVal AdvanceNext As Boolean)
    If AdvanceNext Then
        Me.spdList.MaxRows = Me.spdList.MaxRows + 1
        Call Me.spdList.SetCellBorder(0, 0, Me.spdList.MaxCols, Me.spdList.MaxRows, 15, -1, CellBorderStyleSolid)
        Me.spdList.SetActiveCell miSpdList_LotID, Me.spdList.MaxRows
        
        
        Me.spdList.Col = miSpdList_LotID
        Me.spdList.Row = spdList.MaxRows
        Me.spdList.CellType = CellTypeEdit
        Me.spdList.BackColor = &HFFFFFF
        Me.spdList.EditMode = True
                
    Else
        If Me.spdList.MaxRows > 1 Then
            Me.spdList.MaxRows = Me.spdList.MaxRows - 1
            Call Me.spdList.SetCellBorder(0, 0, Me.spdList.MaxCols, Me.spdList.MaxRows, 15, -1, CellBorderStyleSolid)
        End If
    End If

End Sub




'================================================================================
' Function: GetLotInfo()
'--------------------------------------------------------------------------------
' Description:  <Type your function description here...>
'--------------------------------------------------------------------------------
' Author:       Sam Chen , CIT 2023/12/05
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
Public Function GetLotInfo(ByVal lRow As Long, ByVal sLotId As String) As Boolean
On Error GoTo ExitHandler:
Dim sProcID     As String
Dim typErrInfo  As tErrInfo
Dim i           As Integer

Dim oLot As FwLot

Dim sIPN As String
Dim sLocation As String
Dim sPkgName  As String
Dim sBodySize As String

Dim sPkgCode As String
Dim sSQL As String
Dim colRS As Collection
Dim lIdx As Long

'----
' Init
'----
    sProcID = "GetLotInfo"
    GetLotInfo = True
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog) '"Entering Function...", oLogCtrl, glLOG_PROC, msMODULE_ID, sProcID)

'----
' Condition Checking
'----
    ' <Put your condition checking codes here>...

'----
' Action
'----
    Set oLot = FwuRetrieveLot(moFwWIP, sLotId, moAppLog)
    
    If oLot Is Nothing Then
        UtShowMsgBox "Lotid : " & sLotId & " 不存在."
        GetLotInfo = False
        With spdList
            .Row = lRow
            .SetText mispdList_Check, lRow, "0"
            
            .SetText miSpdList_LotID, lRow, ""
            .SetText miSpdList_IPN, lRow, ""
            .SetText mispdList_Qty, lRow, ""
            .SetText mispdList_Location, lRow, ""
            .SetText mispdList_Datecode, lRow, ""
            .SetText mispdList_PackageName, lRow, ""
            .SetText mispdList_Bodysize, lRow, ""
            .SetText mispdList_PackageCode, lRow, ""
        End With
        
    Else
        sSQL = "select FUN_GET_CUR_LOCATION('" & sLotId & "') as location from dual "
        Set colRS = moProRawSql.QueryDatabase(sSQL)
        If colRS.Count > 0 Then
            sLocation = colRS.Item(1).Item(1)
        End If
        
        sIPN = oLot.CustomAttributes(gsLOT_CUSTOMATTR_IPN)
        sSQL = "SELECT a." & gsCAT_TIM_PACKAGE_NAME & ",a." & gsCAT_TIM_BODY_SIZE & ",a." & gsCAT_TIM_PACKAGE_CODE & " " & _
              " FROM " & gsCAT_TBL_IPN_MASTER & " A " & _
              " where a." & gsCAT_TIM_DELETE_FLAG & " = 'N' " & _
              " and a." & gsCAT_TIM_IPN & " = '" & sIPN & "' "
        Set colRS = moProRawSql.QueryDatabase(sSQL)
        If colRS.Count > 0 Then
            sPkgName = colRS.Item(1).Item(gsCAT_TIM_PACKAGE_NAME)
            sBodySize = colRS.Item(1).Item(gsCAT_TIM_BODY_SIZE)
            sPkgCode = colRS.Item(1).Item(gsCAT_TIM_PACKAGE_CODE)
        End If
              
        With spdList
            .Row = lRow
            .SetText miSpdList_LotID, lRow, sLotId
            .SetText miSpdList_IPN, lRow, sIPN
            .SetText mispdList_Qty, lRow, oLot.ComponentQuantity
            .SetText mispdList_Location, lRow, sLocation
            .SetText mispdList_Datecode, lRow, oLot.CustomAttributes(gsLOT_CUSTOMATTR_DATECODE)
            .SetText mispdList_PackageName, lRow, sPkgName
            .SetText mispdList_Bodysize, lRow, sBodySize
            .SetText mispdList_PackageCode, lRow, sPkgCode
            
            .SetText mispdList_Check, lRow, "1"
        End With
    End If
        
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
        GetLotInfo = False
        On Error GoTo ExitHandler:
        Call HandleError(True, typErrInfo, , moAppLog)
    End If
End Function

Private Sub spdList_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
On Error GoTo ExitHandler:
Dim sProcID     As String
Dim typErrInfo  As tErrInfo

Dim vLotID As Variant
Dim oLot As FwLot

'----
' Init
'----
    sProcID = "spdList_ButtonClicked"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog) '"Entering Function...", oLogCtrl, glLOG_PROC, msMODULE_ID, sProcID)

'----
' Condition Checking
'----
    ' <Put your condition checking codes here>...

'----
' Action
'----
    If Col = mispdList_Check Then
        If ButtonDown Then
            With spdList
                .GetText miSpdList_LotID, Row, vLotID
                Set oLot = FwuRetrieveLot(moFwWIP, CStr(vLotID), moAppLog)
                If oLot Is Nothing Then
                    UtShowMsgBox "Lotid : " & CStr(vLotID) & " 不存在."
                    .SetText mispdList_Check, Row, "0"
                    GoTo ExitHandler
                End If
                
                If oLot.CustomAttributes(gsLOT_CUSTOMATTR_STATUS) = modConstGlobal.gsLOTSTATUS_TERMINATE Or _
                   oLot.CustomAttributes(gsLOT_CUSTOMATTR_STATUS) = modConstGlobal.gsLOTSTATUS_SHIP Then
                    UtShowMsgBox "Lotid : " & CStr(vLotID) & " Status is " & oLot.CustomAttributes(gsLOT_CUSTOMATTR_STATUS) & "."
                    .SetText mispdList_Check, Row, "0"
                    GoTo ExitHandler
                End If
                If oLot.CurrentStep.Steps.Item(1).Attributes(gsSTEP_CUSTOMATTR_STAGE) <> "FT" Then
                    UtShowMsgBox "Lotid : " & CStr(vLotID) & " Stage is not FT."
                    .SetText mispdList_Check, Row, "0"
                    GoTo ExitHandler
                End If
                
            End With
        End If
    End If
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
        Call HandleError(True, typErrInfo, , moAppLog)
    End If
End Sub

Private Sub spdList_Change(ByVal Col As Long, ByVal Row As Long)
On Error GoTo ExitHandler:
Dim sProcID     As String
Dim typErrInfo  As tErrInfo

Dim vLotID As Variant
'----
' Init
'----
    sProcID = "spdList_Change"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog) '"Entering Function...", oLogCtrl, glLOG_PROC, msMODULE_ID, sProcID)

'----
' Condition Checking
'----
    ' <Put your condition checking codes here>...

'----
' Action
'----
    If Col = miSpdList_LotID Then
        With spdList
            .GetText miSpdList_LotID, Row, vLotID
            Call GetLotInfo(Row, CStr(vLotID))
        End With
    End If
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
        Call HandleError(True, typErrInfo, , moAppLog)
    End If

End Sub

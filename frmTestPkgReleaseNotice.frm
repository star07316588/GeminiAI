Option Explicit

Private Const msMODULE_ID As String = "frmTestPkgReleaseNotice"
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
Private moRawSql    As clsRawSql

Private msRuleName  As String
Private msLotId     As String
Private msEqpId     As String

Private Const miSpdLotInfo_TdsProdCode      As Integer = 1
Private Const miSpdLotInfo_ProdGroup        As Integer = 2
Private Const miSpdLotInfo_PkgType          As Integer = 3
Private Const miSpdLotInfo_PinCount         As Integer = 4
Private Const miSpdLotInfo_StepName         As Integer = 5
Private Const miSpdLotInfo_EqType2          As Integer = 6
Private Const miSpdLotInfo_SubSystem        As Integer = 7
Private Const miSpdLotInfo_UPH              As Integer = 8

Private Const msSpdLotInfo_TdsProdCode      As String = "TdsProdCode"
Private Const msSpdLotInfo_ProdGroup        As String = "ProdGroup"
Private Const msSpdLotInfo_PkgType          As String = "PkgType"
Private Const msSpdLotInfo_PinCount         As String = "PinCount"
Private Const msSpdLotInfo_StepName         As String = "StepName"
Private Const msSpdLotInfo_EqType2          As String = "EqType2"
Private Const msSpdLotInfo_SubSystem        As String = "SubSystem"
Private Const msSpdLotInfo_UPH              As String = "UPH"

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
Private Sub cmdQuery_Click()
    Call ExecQuery
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
' Sub: ExecQuery()
'--------------------------------------------------------------------------------
' Description:  <Type your Sub description here...>
'--------------------------------------------------------------------------------
' Author:       Jeff Weng,  MXIC 2004/07/30
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
'================================================================================
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
    
    If (Trim(txtTdsProdCode.Text) = "" And Trim(txtProdGroup.Text) = "" And Trim(cboPkgType) = "" And _
        Trim(cboStepName) = "" And Trim(cboEqType2) = "" And Trim(cboSubSystem) = "") Then
            UtShowMsgBox "Pls input Query Criterial !!" & vbNewLine & "請輸入查詢條件 !!"
            GoTo ExitHandler
    End If
    

'----
' Action
'----
    Dim sSQL As String
    Dim oRawSql As New clsRawSql
    Dim oRS As clsRawSqlRecordset
    Dim mlOldChipQtySum As Long
    Dim mlNewChipQtySum As Long
    
    Dim colRS As Collection '2009/05/14 Vencent 加查詢筆數限制
    Dim sColumn As String

    Set oRawSql.CwMbx = moCwMbx
    Set oRawSql.AppLog = moAppLog
    oRawSql.RawSqlMailBox = moFwWF.ClientID & ".RawSql_Reply"
    
    Screen.MousePointer = vbHourglass

    Me.spdLotInfo.MaxRows = 0

    sColumn = "select a.prodgroup," & _
                 " a.stepno," & _
                 " a.stepname," & _
                 " a.eqtype2," & _
                 " a.subsystem," & _
                 " a.uph," & _
                 " b.packagecode," & _
                 " b.pincount," & _
                 " g.magnitude "
'----- User command modify rule on 040805-------------------------------------------------------------
'    sSql = sSql & " from " & _
'            "(select prodgroup,stepno,stepname,eqtype2,subsystem,uph from tbl_prod_step_eq_spec " & _
'                    " where docstatus ='Active' and " & _
'                            "((specifyeq ='N') or " & _
'                            "(specifyeq is null) or " & _
'                            "(specifyeq ='Y' and eqid is not null))) a," & _
'            "(select prodgroup,packagecode,pincount from tbl_prod_spec where docstatus='Active') b," & _
'            "(select c.productname,e.magnitude from " & _
'                    "(select * from fwproductversion where revstate='Active') c," & _
'                    "(select * from fwproductversion_n2m where linkname='attributes') d," & _
'                    "(select * from fwprpattributeinstance) e," & _
'                    "(select * from fwprpattribute where name='TdsProd') f " & _
'                    " where c.sysid=d.fromid and d.toid=e.sysid and e.attrclass=f.sysid) g"

    sSQL = " from " & _
            "(select prodgroup,stepno,stepname,eqtype2,subsystem,uph from tbl_prod_step_eq_spec " & _
                    " where docstatus ='Active') a," & _
            "(select prodgroup,packagecode,pincount from tbl_prod_spec where docstatus='Active') b," & _
            "(select c.productname,e.magnitude from " & _
                    "(select * from fwproductversion where revstate='Active') c," & _
                    "(select * from fwproductversion_n2m where linkname='attributes') d," & _
                    "(select * from fwprpattributeinstance) e," & _
                    "(select * from fwprpattribute where name='TdsProd') f " & _
                    " where c.sysid=d.fromid and d.toid=e.sysid and e.attrclass=f.sysid) g"
    sSQL = sSQL & " where " & _
            "a.prodgroup=b.prodgroup and a.prodgroup=g.productname(+)"
        
    
    'ProdGroup
    If Me.txtProdGroup <> "" Then
        If InStr(1, Me.txtProdGroup, "*") > 0 Then
            sSQL = sSQL & " and a.prodgroup like '" & Replace(Me.txtProdGroup, "*", "%") & "' "
        Else
            sSQL = sSQL & " and a.prodgroup ='" & Me.txtProdGroup & "' "
        End If
    End If
    
    
    'StepName
    
    If Me.cboStepName <> "" Then
        sSQL = sSQL & " and a.stepname like '" & Me.cboStepName & "' "
    End If
    
    'EqType2
    If Me.cboEqType2 <> "" Then
        If InStr(1, Me.cboEqType2, "*") > 0 Then
            sSQL = sSQL & " and a.eqtype2 like '" & Replace(Me.cboEqType2, "*", "%") & "' "
        Else
            sSQL = sSQL & " and a.eqtype2 ='" & Me.cboEqType2 & "' "
        End If
    End If
    
    'SubSystem
    If Trim(Me.cboSubSystem) <> "" Then
        If InStr(1, Me.cboSubSystem, "*") > 0 Then
            sSQL = sSQL & " and a.SubSystem like '" & Replace(Me.cboSubSystem, "*", "%") & "' "
        Else
            sSQL = sSQL & " and a.SubSystem ='" & Me.cboSubSystem & "' "
        End If
    End If
    
    
    'TdsProdGroup
    If Me.txtTdsProdCode <> "" Then
        If InStr(1, Me.txtTdsProdCode, "*") > 0 Then
            sSQL = sSQL & " and g.magnitude like '" & Replace(Me.txtTdsProdCode, "*", "%") & "' "
        Else
            sSQL = sSQL & " and g.magnitude ='" & Me.txtTdsProdCode & "' "
        End If
    End If
    
    
    'PkgType
    If Me.cboPkgType <> "" Then
        If InStr(1, Me.cboPkgType, "*") > 0 Then
            sSQL = sSQL & " and b.packagecode like '" & Replace(Me.cboPkgType, "*", "%") & "' "
        Else
            sSQL = sSQL & " and b.packagecode ='" & Me.cboPkgType & "' "
        End If
    End If
    
    '2009/05/18 Vencent 加查詢筆數限制
    Set colRS = moProRawSql.QueryDatabase("select count(*) " & sSQL)
    If Val(colRS.Item(1).Item(1)) > gsQueryLimitCnt Then
        UtShowRichBox gsQueryLimitMsg
        GoTo ExitHandler
    End If
    
    sSQL = sColumn & sSQL
    
    Debug.Print sSQL
    
    Set oRS = oRawSql.SelectRawSqlTxn1("CAT", sSQL)

    With Me.spdLotInfo
    If oRS.RecordCount <= 0 Then
        UtShowMsgBox "無符合條件之資料 !!"
    Else
        Do While oRS.EOF = False
            
            .MaxRows = .MaxRows + 1
            .SetText miSpdLotInfo_TdsProdCode, .MaxRows, oRS.Value("magnitude")
            .SetText miSpdLotInfo_ProdGroup, .MaxRows, oRS.Value("prodgroup")
            .SetText miSpdLotInfo_PkgType, .MaxRows, oRS.Value("packagecode")
            .SetText miSpdLotInfo_PinCount, .MaxRows, oRS.Value("pincount")
            .SetText miSpdLotInfo_StepName, .MaxRows, oRS.Value("stepno") & "," & oRS.Value("stepname")
            .SetText miSpdLotInfo_EqType2, .MaxRows, oRS.Value("eqtype2")
            .SetText miSpdLotInfo_SubSystem, .MaxRows, oRS.Value("SubSystem")
            .SetText miSpdLotInfo_UPH, .MaxRows, oRS.Value("UPH")
    
            oRS.MoveNext
        Loop
    End If
    End With

        
    If oRS.RecordCount > 0 Then
        'Set up spd
        With Me.spdLotInfo
            .Col = 1
            Do While .Col <= .MaxCols
                .ColWidth(.Col) = .MaxTextColWidth(.Col) + 1
                .Col = .Col + 1
            Loop
            
            'Sort it
            .SetActiveCell miSpdLotInfo_TdsProdCode, 1
            .UserColAction = UserColActionSort
            
            Call .SetCellBorder(0, 0, .MaxCols, .MaxRows, 15, -1, CellBorderStyleSolid)
            
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

'================================================================================
' Function: Init()
'--------------------------------------------------------------------------------
' Description:  <Type your function description here...>
'--------------------------------------------------------------------------------
' Author:       Jeff Weng, MXIC 2004/07/30
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
    Call Initspd
    Call InitComboBox
    
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
        Call HandleError(True, typErrInfo, , moAppLog, True)
    End If
End Function

'================================================================================
' Function: mInitSpread()
'--------------------------------------------------------------------------------
' Description:  <Type your function description here...>
'--------------------------------------------------------------------------------
' Author:       Jeff Weng, MXIC 2004/08/04
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
Public Function Initspd() As Boolean
On Error GoTo ExitHandler:
Dim sProcID     As String
Dim typErrInfo  As tErrInfo
Dim i           As Integer
Dim iIndex      As Long

'----
' Init
'----
    sProcID = "InitSpd"
    Initspd = True
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog) '"Entering Function...", oLogCtrl, glLOG_PROC, msMODULE_ID, sProcID)

'----
' Condition Checking
'----
    ' <Put your condition checking codes here>...

'----
' Action
'----


    With spdLotInfo
        .MaxRows = 0
        .MaxCols = 8
        .SetText miSpdLotInfo_TdsProdCode, 0, msSpdLotInfo_TdsProdCode
        .SetText miSpdLotInfo_ProdGroup, 0, msSpdLotInfo_ProdGroup
        .SetText miSpdLotInfo_PkgType, 0, msSpdLotInfo_PkgType
        .SetText miSpdLotInfo_PinCount, 0, msSpdLotInfo_PinCount
        .SetText miSpdLotInfo_StepName, 0, msSpdLotInfo_StepName
        .SetText miSpdLotInfo_EqType2, 0, msSpdLotInfo_EqType2
        .SetText miSpdLotInfo_SubSystem, 0, msSpdLotInfo_SubSystem
        .SetText miSpdLotInfo_UPH, 0, msSpdLotInfo_UPH
        
        For i = 1 To .MaxCols
            .ColUserSortIndicator(i) = ColUserSortIndicatorAscending
            .UserColAction = UserColActionSort
        
        Next
        
        .Row = -1
        For iIndex = 1 To .MaxCols
            Call LockSpdColumn(Me.spdLotInfo, iIndex)
        Next iIndex
        
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
        Initspd = False
        On Error GoTo ExitHandler:
        Call HandleError(True, typErrInfo, , moAppLog)
    End If
End Function


'================================================================================
' Sub: InitComboBox()
'--------------------------------------------------------------------------------
' Description:  <Type your Sub description here...>
'--------------------------------------------------------------------------------
' Author:       Jeff Weng, MXIC 2004/08/03
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
Private Sub InitComboBox()
On Error GoTo ExitHandler:
Dim sProcID As String
Dim typErrInfo As tErrInfo
'----
' Init
'----
    sProcID = "InitComboBox"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog)

'----
' Condition Checking
'----
    
'----
' Action
'----
    
    Dim sSQL As String, colSQLResult As Collection
    Dim iIdx As Integer
    
    'Pkg Type
    sSQL = "select distinct packagecode from tbl_prod_spec"
    Set colSQLResult = moProRawSql.QueryDatabase(sSQL)
    If colSQLResult Is Nothing Then
        Call RaiseError(glERR_RAWSQLERROR, _
                    FormatErrorText(gsETX_RAWSQLERROR, "Select", "PkgType"))
    ElseIf colSQLResult.Count > 0 Then
        Me.cboPkgType.Clear
        Me.cboPkgType.AddItem ""
        
        For iIdx = 1 To colSQLResult.Count
            Me.cboPkgType.AddItem colSQLResult.Item(iIdx).Item(1)
        Next iIdx
    End If
    
    
    'Step Name
    sSQL = "select distinct stepname from tbl_prod_step_eq_spec where substr(stepname,1,4)='SORT' or substr(stepname,1,2)='FT'"
    Set colSQLResult = moProRawSql.QueryDatabase(sSQL)
    If colSQLResult Is Nothing Then
        Call RaiseError(glERR_RAWSQLERROR, _
                    FormatErrorText(gsETX_RAWSQLERROR, "Select", "FwStep"))
    ElseIf colSQLResult.Count > 0 Then
        Me.cboStepName.Clear
        Me.cboStepName.AddItem ""
        
        For iIdx = 1 To colSQLResult.Count
            Me.cboStepName.AddItem colSQLResult.Item(iIdx).Item(1)
        Next iIdx
    End If
    
    
    'Eq Type2
    sSQL = "select distinct eqtype2 from tbl_prod_step_eq_spec"
    Set colSQLResult = moProRawSql.QueryDatabase(sSQL)
    If colSQLResult Is Nothing Then
        Call RaiseError(glERR_RAWSQLERROR, _
                    FormatErrorText(gsETX_RAWSQLERROR, "Select", "EqType2"))
    ElseIf colSQLResult.Count > 0 Then
        Me.cboEqType2.Clear
        Me.cboEqType2.AddItem ""
        
        For iIdx = 1 To colSQLResult.Count
            Me.cboEqType2.AddItem colSQLResult.Item(iIdx).Item(1)
        Next iIdx
    End If
    
    
    'Sub System
    sSQL = "select distinct subsystem from tbl_prod_step_eq_spec"
    Set colSQLResult = moProRawSql.QueryDatabase(sSQL)
    If colSQLResult Is Nothing Then
        Call RaiseError(glERR_RAWSQLERROR, _
                    FormatErrorText(gsETX_RAWSQLERROR, "Select", "SubSystem"))
    ElseIf colSQLResult.Count > 0 Then
        Me.cboSubSystem.Clear

        
        For iIdx = 1 To colSQLResult.Count
            Me.cboSubSystem.AddItem colSQLResult.Item(iIdx).Item(1)
        Next iIdx
        Me.cboSubSystem.RemoveItem (0)
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
        Call HandleError(False, typErrInfo, , moAppLog, True)
    End If
End Sub



Option Explicit

Private Const msMODULE_ID As String = "frmMergeSplitQuery"
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

Private Const msSpdCol_Activity         As String = "Activity"
Private Const msSpdCol_ParentLotId      As String = "ParentLotId"
Private Const msSpdCol_ParIPN           As String = "ParIPN"
Private Const msSpdCol_StepId           As String = "StepId"
Private Const msSpdCol_StepName         As String = "StepName"
Private Const msSpdCol_OriParWQty       As String = "OriParWQty"
Private Const msSpdCol_OriParCQty       As String = "OriParCQty"
Private Const msSpdCol_AftParWQty       As String = "AftParWQty"
Private Const msSpdCol_AftParCQty       As String = "AftParCQty"
Private Const msSpdCol_ChildLotId       As String = "ChildLotId"
Private Const msSpdCol_ChildWQty        As String = "ChildWQty"
Private Const msSpdCol_ChildCQty        As String = "ChildCQty"
Private Const msSpdCol_TimeStamp        As String = "TimeStamp"

Private Const miSpdCol_Activity         As Integer = 1
Private Const miSpdCol_ParentLotId      As Integer = 2
Private Const miSpdCol_ParIPN           As Integer = 3
Private Const miSpdCol_StepId           As Integer = 4
Private Const miSpdCol_StepName         As Integer = 5
Private Const miSpdCol_OriParWQty       As Integer = 6
Private Const miSpdCol_OriParCQty       As Integer = 7
Private Const miSpdCol_AftParWQty       As Integer = 8
Private Const miSpdCol_AftParCQty       As Integer = 9
Private Const miSpdCol_ChildLotId       As Integer = 10
Private Const miSpdCol_ChildWQty        As Integer = 11
Private Const miSpdCol_ChildCQty        As Integer = 12
Private Const miSpdCol_TimeStamp        As Integer = 13

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
    On Error Resume Next
    Result = miResult
End Property

Private Sub cmdCancel_Click()
    On Error Resume Next
    miResult = vbCancel
    Me.Hide
End Sub

Private Sub cmdPrint_Click()
    Me.spdList.PrintHeader = "Query by " & moFwOPR.ActiveUser.UserName _
        & " - " & Format(Now, "YYYYMMDD HHMMSS")
    Me.spdList.PrintSheet (1)
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

'----
' Action
'----
'    Me.dtpFromDate.Value = Date - 1
'    Me.dtpToDate.Value = Date
    Me.cboActivity.ListIndex = Me.cboActivity.ListCount - 1
    
    With Me.spdList
        .MaxRows = 1
        .MaxCols = 13
    
        .Col = 1
        Do While .Col <= .MaxCols
            .ColWidth(.Col) = .MaxTextColWidth(.Col) + 1
            .ColUserSortIndicator(.Col) = ColUserSortIndicatorAscending
            .UserColAction = UserColActionSort
            .Col = .Col + 1
        Loop
    End With
    
'----
' Done
'----

ExitHandler:
    ' NOTE 1:
    ' MUST CALL GetErrInfo() here first before another action
    Call GetErrInfo(msMODULE_ID, sProcID, typErrInfo, Erl)
    Call LogProcOut(msMODULE_ID, sProcID, typErrInfo, moAppLog)
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
        Init = False
        On Error GoTo ExitHandler:
        Call HandleError(True, typErrInfo, , moAppLog)
    End If
End Function

'================================================================================
' Sub: ExecQuery()
'--------------------------------------------------------------------------------
' Description:  <Type your Sub description here...>
'--------------------------------------------------------------------------------
' Author:       Vencent, CIT 2004/10/13
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
Private Sub ExecQuery()
On Error GoTo ExitHandler:
Dim sProcID As String
Dim typErrInfo As tErrInfo
Dim sSQL        As String
Dim sMergeSQL   As String
Dim sSplitSQL   As String

Dim sParentLotID    As String
Dim sChildLotID     As String

Dim oRawSqlRecordSet    As New clsRawSqlRecordset
Dim oRawSqlRecordSet2   As New clsRawSqlRecordset
Dim oRawSql             As New clsRawSql

Dim sSysID      As String
Dim colRS       As Collection
Dim iIdx        As Integer
Dim lCnt        As Long
'----
' Init
'----
    sProcID = "ExecQuery"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog)
    
    Screen.MousePointer = vbHourglass
    Set oRawSql.CwMbx = moCwMbx
    Set oRawSql.AppLog = moAppLog
    oRawSql.RawSqlMailBox = moFwWF.ClientID & ".RawSql_Reply"
'----
' Condition Checking
'----
    ' <Put your condition checking codes here>...
    If Me.txtParentLotID = "" And Me.txtIPN = "" And Me.txtChildLotId = "" _
    And Me.txtStepName = "" And Me.txtDateFrom = "" And Me.txtDateTo = "" Then
        UtShowMsgBox "請輸入至少一個條件"
        GoTo ExitHandler
    End If
'----
' Action
'----
    ' <Put your Action codes here>...
    If Me.cboActivity.Text = "Merge" Or Me.cboActivity.Text = "All" Then
        sMergeSQL = "select A.WIPID PARENTLOT," & vbNewLine & _
                    "      A.ACTIVITY," & vbNewLine & _
                    "      A.TXNTIMESTAMP TXNTIME," & vbNewLine & _
                    "      B.VALDATA CHILDLOT," & vbNewLine & _
                    "      A.MERGESTEPID STEPID," & vbNewLine & _
                    "      C.DESCRIPTION STEPNAME," & vbNewLine & _
                    "      D.VALDATA as IPN," & vbNewLine & _
                    "      0 as QTY, " & vbNewLine & _
                    "      A.SYSID" & vbNewLine & _
                    "  from FWMERGE A," & vbNewLine & _
                    "      FWMERGE_PN2M B," & vbNewLine & _
                    "      (select STEPNAME, DESCRIPTION from FWSTEPVERSION where REVSTATE = 'Active') C," & vbNewLine & _
                    "      (select FROMID, VALDATA from FWLOT_PN2M where KEYDATA = 'IPN') D" & vbNewLine & _
                    " where A.SYSID = B.FROMID and" & vbNewLine & _
                    "       B.LINKNAME = 'childLotCollection' and" & vbNewLine
                    '"      SUBSTR(A.MERGESTEPID, 1, 5) = C.STEPNAME and" & vbNewLine & _
                    '"      A.LOTOBJECT = D.FROMID" & vbNewLine
        
        'Start add by Janus, 20070827, for performance tuning
        If (Trim(Me.txtParentLotID.Text) <> "*" And Trim(Me.txtParentLotID.Text) <> "") Or _
           (Trim(Me.txtIPN.Text) <> "*" And Trim(Me.txtIPN.Text) <> "") Or _
           (Trim(Me.txtStepName.Text) <> "*" And Trim(Me.txtStepName.Text) <> "") Or _
           (Trim(Me.txtDateFrom.Text) <> "") Or (Trim(Me.txtDateTo.Text) <> "") Then
           sMergeSQL = sMergeSQL & "  SUBSTR(A.MERGESTEPID, 1, 5) = C.STEPNAME(+) and" & vbNewLine & _
                                   "  A.LOTOBJECT = D.FROMID(+) "
        Else
           sMergeSQL = sMergeSQL & "  SUBSTR(A.MERGESTEPID, 1, 5) = C.STEPNAME and" & vbNewLine & _
                                   "  A.LOTOBJECT = D.FROMID "
        End If
        'End add by Janus, 20070827, for performance tuning

        If Trim(Me.txtParentLotID.Text) <> "*" And Trim(Me.txtParentLotID.Text) <> "" Then
            If InStr(1, Me.txtParentLotID.Text, "*") > 0 Then
                sMergeSQL = sMergeSQL & " and A.WIPID like '" & Replace(Trim(Me.txtParentLotID.Text), "*", "%") & "'"
            Else
                sMergeSQL = sMergeSQL & " and A.WIPID = '" & Trim(Me.txtParentLotID.Text) & "'"
            End If
        End If
                
        If Trim(Me.txtIPN.Text) <> "*" And Trim(Me.txtIPN.Text) <> "" Then
            If InStr(1, Me.txtIPN.Text, "*") > 0 Then
                sMergeSQL = sMergeSQL & " and D.VALDATA like '" & Replace(Trim(Me.txtIPN.Text), "*", "%") & "'"
            Else
                sMergeSQL = sMergeSQL & " and D.VALDATA = '" & Trim(Me.txtIPN.Text) & "'"
            End If
        End If
        
        If Trim(Me.txtStepName.Text) <> "*" And Trim(Me.txtStepName.Text) <> "" Then
            If InStr(1, Me.txtStepName.Text, "*") > 0 Then
                sMergeSQL = sMergeSQL & " and C.DESCRIPTION like '" & Replace(Trim(Me.txtStepName.Text), "*", "%") & "'"
            Else
                sMergeSQL = sMergeSQL & " and C.DESCRIPTION = '" & Trim(Me.txtStepName.Text) & "'"
            End If
        End If
        
        If Trim(Me.txtDateFrom.Text) <> "" Then
            sMergeSQL = sMergeSQL & " and A.TXNTIMESTAMP > '" & Trim(Me.txtDateFrom.Text) & "'"
        End If
        If Trim(Me.txtDateTo.Text) <> "" Then
            sMergeSQL = sMergeSQL & " and substr(A.TXNTIMESTAMP,1,8) <= '" & Trim(Me.txtDateTo.Text) & "'"
        End If
        
        If Trim(Me.txtChildLotId.Text) <> "*" And Trim(Me.txtChildLotId.Text) <> "" Then
            sSQL = "select A.PARENTTXN" & vbNewLine & _
                    "from FWMERGE A" & vbNewLine & _
                    "where A.WIPID = '" & Trim(Me.txtChildLotId.Text) & "'" & vbNewLine & _
                    "and A.ACTIVITY = 'Merge'" & vbNewLine & _
                    "and A.PARENTTXN is not null"
            Set oRawSqlRecordSet2 = oRawSql.SelectRawSqlTxn1("CAT", sSQL)
            sSysID = ""
            Do While Not oRawSqlRecordSet2.EOF
                If sSysID = "" Then
                    sSysID = "'" & oRawSqlRecordSet2.Value("PARENTTXN") & "'"
                Else
                    sSysID = "'" & oRawSqlRecordSet2.Value("PARENTTXN") & "'" & "," & sSysID
                End If
                oRawSqlRecordSet2.MoveNext
            Loop
            If sSysID = "" Then
                sMergeSQL = ""
            Else
                sMergeSQL = sMergeSQL & vbNewLine & _
                        " and A.SYSID in (" & sSysID & ") " & _
                        " and B.VALDATA = '" & Trim(Me.txtChildLotId.Text) & "'"
            End If
        End If
    End If
    
    If Me.cboActivity.Text = "Split" Or Me.cboActivity.Text = "All" Then
        sSplitSQL = "select FSL.WIPID PARENTLOT," & vbNewLine & _
                    "      FSL.ACTIVITY," & vbNewLine & _
                    "      FSL.TXNTIMESTAMP TXNTIME," & vbNewLine & _
                    "      FS.CHILDLOTID CHILDLOT," & vbNewLine & _
                    "      FSLM.VALDATA STEPID," & vbNewLine & _
                    "      C.DESCRIPTION STEPNAME," & vbNewLine & _
                    "      D.VALDATA as IPN," & vbNewLine & _
                    "      FS.SPLITQTY QTY," & vbNewLine & _
                    "      FSL.SYSID" & vbNewLine & _
                    "  from FWSPLITLOT      FSL," & vbNewLine & _
                    "      FWSPLITLOT_N2M  FSLN," & vbNewLine & _
                    "      FWSPLIT         FS," & vbNewLine & _
                    "      FWSPLITLOT_PN2M FSLM," & vbNewLine & _
                    "      (select STEPNAME, DESCRIPTION from FWSTEPVERSION where REVSTATE = 'Active') C," & vbNewLine & _
                    "      (select FROMID, VALDATA from FWLOT_PN2M where KEYDATA = 'IPN') D" & vbNewLine & _
                    " where FSL.SYSID = FSLN.FROMID and" & vbNewLine & _
                    "       FSLN.TOID = FS.SYSID and" & vbNewLine & _
                    "      FSL.SYSID = FSLM.FROMID and" & vbNewLine & _
                    "       FSLM.KEYDATA = 'stepId' and" & vbNewLine
                    '"       substr(FSLM.VALDATA, 1, 5) = C.STEPNAME and" & vbNewLine & _
                    '"       FSL.LOTOBJECT = D.FROMID"
        
        'Start add by Janus, 20070827, for performance tuning
        If (Trim(Me.txtParentLotID.Text) <> "*" And Trim(Me.txtParentLotID.Text) <> "") Or _
           (Trim(Me.txtIPN.Text) <> "*" And Trim(Me.txtIPN.Text) <> "") Or _
           (Trim(Me.txtStepName.Text) <> "*" And Trim(Me.txtStepName.Text) <> "") Or _
           (Trim(Me.txtDateFrom.Text) <> "") Or (Trim(Me.txtDateTo.Text) <> "") Then
           sSplitSQL = sSplitSQL & "  substr(FSLM.VALDATA, 1, 5) = C.STEPNAME(+) and" & vbNewLine & _
                                   "  FSL.LOTOBJECT = D.FROMID(+)"
        Else
           sSplitSQL = sSplitSQL & "  substr(FSLM.VALDATA, 1, 5) = C.STEPNAME and" & vbNewLine & _
                                   "  FSL.LOTOBJECT = D.FROMID "
        End If
        'End add by Janus, 20070827, for performance tuning

        If Trim(Me.txtParentLotID.Text) <> "*" And Trim(Me.txtParentLotID.Text) <> "" Then
            If InStr(1, Me.txtParentLotID.Text, "*") > 0 Then
                sSplitSQL = sSplitSQL & " and FSL.WIPID like '" & Replace(Trim(Me.txtParentLotID.Text), "*", "%") & "'"
            Else
                sSplitSQL = sSplitSQL & " and FSL.WIPID = '" & Trim(Me.txtParentLotID.Text) & "'"
            End If
        End If
        
        If Trim(Me.txtIPN.Text) <> "*" And Trim(Me.txtIPN.Text) <> "" Then
            If InStr(1, Me.txtIPN.Text, "*") > 0 Then
                sSplitSQL = sSplitSQL & " and D.VALDATA like '" & Replace(Trim(Me.txtIPN.Text), "*", "%") & "'"
            Else
                sSplitSQL = sSplitSQL & " and D.VALDATA = '" & Trim(Me.txtIPN.Text) & "'"
            End If
        End If
        
        If Trim(Me.txtStepName.Text) <> "*" And Trim(Me.txtStepName.Text) <> "" Then
            If InStr(1, Me.txtStepName.Text, "*") > 0 Then
                sSplitSQL = sSplitSQL & " and C.DESCRIPTION like '" & Replace(Trim(Me.txtStepName.Text), "*", "%") & "'"
            Else
                sSplitSQL = sSplitSQL & " and C.DESCRIPTION = '" & Trim(Me.txtStepName.Text) & "'"
            End If
        End If
        
        If Trim(Me.txtDateFrom.Text) <> "" Then
            sSplitSQL = sSplitSQL & " and FSL.TXNTIMESTAMP > '" & Trim(Me.txtDateFrom.Text) & "'"
        End If
        If Trim(Me.txtDateTo.Text) <> "" Then
            sSplitSQL = sSplitSQL & " and substr(FSL.TXNTIMESTAMP,1,8) <= '" & Trim(Me.txtDateTo.Text) & "'"
        End If
        If Trim(Me.txtChildLotId.Text) <> "*" And Trim(Me.txtChildLotId.Text) <> "" Then
            sSQL = "select A.PARENTTXN" & vbNewLine & _
                    "from FWSPLITLOT A" & vbNewLine & _
                    "where A.WIPID = '" & Trim(Me.txtChildLotId.Text) & "'" & vbNewLine & _
                    "and A.ACTIVITY = 'Split'" & vbNewLine & _
                    "and A.PARENTTXN is not null"
            Set oRawSqlRecordSet2 = oRawSql.SelectRawSqlTxn1("CAT", sSQL)
            sSysID = ""
            Do While Not oRawSqlRecordSet2.EOF
                If sSysID = "" Then
                    sSysID = "'" & oRawSqlRecordSet2.Value("PARENTTXN") & "'"
                Else
                    sSysID = "'" & oRawSqlRecordSet2.Value("PARENTTXN") & "'" & "," & sSysID
                End If
                oRawSqlRecordSet2.MoveNext
            Loop
            
            If sSysID = "" Then
                sSplitSQL = ""
            Else
                sSplitSQL = sSplitSQL & vbNewLine & _
                        " and FSL.SYSID in (" & sSysID & ") " & _
                        " and FS.CHILDLOTID = '" & Trim(Me.txtChildLotId.Text) & "'"
            End If
        End If
     End If
    
    '2009/05/14 Vencent 加查詢筆數限制
    If sMergeSQL <> "" Then
        Set colRS = moProRawSql.QueryDatabase("select count(*) from (" & sMergeSQL & ")")
        lCnt = Val(colRS.Item(1).Item(1))
        If lCnt > gsQueryLimitCnt Then
            UtShowRichBox gsQueryLimitMsg
            GoTo ExitHandler
        End If
    End If
    If sSplitSQL <> "" Then
        Set colRS = moProRawSql.QueryDatabase("select count(*) from (" & sSplitSQL & ")")
        lCnt = lCnt + Val(colRS.Item(1).Item(1))
    End If
    If lCnt > gsQueryLimitCnt Then
        UtShowRichBox gsQueryLimitMsg
        GoTo ExitHandler
    End If
    
    If sMergeSQL <> "" And sSplitSQL <> "" Then
        sSQL = sMergeSQL & " union all " & sSplitSQL & " order by PARENTLOT, TXNTIME"
    ElseIf sMergeSQL <> "" Then
        sSQL = sMergeSQL & " order by PARENTLOT, TXNTIME"
    ElseIf sSplitSQL <> "" Then
        sSQL = sSplitSQL & " order by PARENTLOT, TXNTIME"
    Else
        sSQL = ""
    End If
    
    If sSQL <> "" Then Set oRawSqlRecordSet = oRawSql.SelectRawSqlTxn1("CAT", sSQL)
    
    With Me.spdList
        .MaxRows = 0
        Do While Not oRawSqlRecordSet.EOF
            .MaxRows = .MaxRows + 1
'            .SetActiveCell 1, .MaxRows
            sParentLotID = oRawSqlRecordSet.Value("PARENTLOT")
            sChildLotID = oRawSqlRecordSet.Value("CHILDLOT")
            .SetText miSpdCol_Activity, .MaxRows, oRawSqlRecordSet.Value("ACTIVITY")
            .SetText miSpdCol_ParentLotId, .MaxRows, sParentLotID
            .SetText miSpdCol_ChildLotId, .MaxRows, sChildLotID
            .SetText miSpdCol_ParIPN, .MaxRows, oRawSqlRecordSet.Value("IPN")
            .SetText miSpdCol_StepId, .MaxRows, Left(oRawSqlRecordSet.Value("STEPID"), 5)
            .SetText miSpdCol_StepName, .MaxRows, oRawSqlRecordSet.Value("STEPNAME")
            .SetText miSpdCol_TimeStamp, .MaxRows, oRawSqlRecordSet.Value("TXNTIME")
            
            sSQL = "select A.WIPID," & vbNewLine & _
                    "      A.ATTRIBUTE," & vbNewLine & _
                    "      A.ORIGINALVALUE," & vbNewLine & _
                    "      A.NEWVALUE" & vbNewLine & _
                    "  from FWWIPTRANSACTION A, FWWIPHISTORY B" & vbNewLine & _
                    " where A.WIPID in ('" & sParentLotID & "', '" & sChildLotID & "') and" & vbNewLine & _
                    "      A.ACTIVITY = 'ModifyAttribute' and" & vbNewLine & _
                    "      A.ATTRIBUTE in ('ChipQty', 'WaferQty') and" & vbNewLine & _
                    "      A.SYSID = B.WIPTXN and" & vbNewLine & _
                    "      B.GROUPHISTKEY =" & vbNewLine & _
                    "      (select GROUPHISTKEY" & vbNewLine & _
                    "         from FWWIPHISTORY" & vbNewLine & _
                    "        where WIPTXN = '" & oRawSqlRecordSet.Value("SYSID") & "')"
            
            sSQL = sSQL & " order by A.TXNTIMESTAMP desc, A.TXNSEQUENCE desc"
            
            'Set oRawSqlRecordSet2 = oRawSql.SelectRawSqlTxn1("CAT", sSql)
            'Do While Not oRawSqlRecordSet2.EOF
            Set colRS = moProRawSql.QueryDatabase(sSQL)
            For iIdx = 1 To colRS.Count
                'If oRawSqlRecordSet2.Value("WIPID") = sParentLotID Then
                If colRS.Item(iIdx).Item("WIPID") = sParentLotID Then
                    'If oRawSqlRecordSet2.Value("ATTRIBUTE") = gsLOT_CUSTOMATTR_CHIP_QTY Then
                    If colRS.Item(iIdx).Item("ATTRIBUTE") = gsLOT_CUSTOMATTR_CHIP_QTY Then
                        '.SetText miSpdCol_OriParCQty, .MaxRows, oRawSqlRecordSet2.Value("ORIGINALVALUE")
                        .SetText miSpdCol_OriParCQty, .MaxRows, colRS.Item(iIdx).Item("ORIGINALVALUE")
                        .Col = miSpdCol_AftParCQty
                        .Row = .MaxRows
                        'If .Text = "" Then .SetText miSpdCol_AftParCQty, .MaxRows, oRawSqlRecordSet2.Value("NEWVALUE")
                        If .Text = "" Then .SetText miSpdCol_AftParCQty, .MaxRows, colRS.Item(iIdx).Item("NEWVALUE")
                    Else
                        '.SetText miSpdCol_OriParWQty, .MaxRows, oRawSqlRecordSet2.Value("ORIGINALVALUE")
                        .SetText miSpdCol_OriParWQty, .MaxRows, colRS.Item(iIdx).Item("ORIGINALVALUE")
                        .Col = miSpdCol_AftParWQty
                        .Row = .MaxRows
                        'If .Text = "" Then .SetText miSpdCol_AftParWQty, .MaxRows, oRawSqlRecordSet2.Value("NEWVALUE")
                        If .Text = "" Then .SetText miSpdCol_AftParWQty, .MaxRows, colRS.Item(iIdx).Item("NEWVALUE")
                    End If
                Else
                    'If oRawSqlRecordSet2.Value("ATTRIBUTE") = gsLOT_CUSTOMATTR_CHIP_QTY Then
                    If colRS.Item(iIdx).Item("ATTRIBUTE") = gsLOT_CUSTOMATTR_CHIP_QTY Then
                        If oRawSqlRecordSet.Value("ACTIVITY") = "Merge" Then
                            '.SetText miSpdCol_ChildCQty, .MaxRows, oRawSqlRecordSet2.Value("ORIGINALVALUE")
                            .SetText miSpdCol_ChildCQty, .MaxRows, colRS.Item(iIdx).Item("ORIGINALVALUE")
                        Else
                            '.SetText miSpdCol_ChildCQty, .MaxRows, oRawSqlRecordSet2.Value("NEWVALUE")
                            .SetText miSpdCol_ChildCQty, .MaxRows, colRS.Item(iIdx).Item("NEWVALUE")
                        End If
                    Else
                        If oRawSqlRecordSet.Value("ACTIVITY") = "Merge" Then
                            '.SetText miSpdCol_ChildWQty, .MaxRows, oRawSqlRecordSet2.Value("ORIGINALVALUE")
                            .SetText miSpdCol_ChildWQty, .MaxRows, colRS.Item(iIdx).Item("ORIGINALVALUE")
                        Else
                            '.SetText miSpdCol_ChildWQty, .MaxRows, oRawSqlRecordSet2.Value("NEWVALUE")
                            .SetText miSpdCol_ChildWQty, .MaxRows, colRS.Item(iIdx).Item("NEWVALUE")
                        End If
                    End If
                End If
            '    oRawSqlRecordSet2.MoveNext
            'Loop
            Next iIdx
            
            .Col = miSpdCol_ChildCQty
            .Row = .MaxRows
            If .Text = "" Then
                sSQL = "select A.ATTRIBUTE, A.NEWVALUE" & vbNewLine & _
                        "  from FWWIPTRANSACTION A" & vbNewLine & _
                        " where A.WIPID = '" & sChildLotID & "' and" & vbNewLine & _
                        "       A.ACTIVITY = 'ModifyAttribute' and" & vbNewLine & _
                        "      A.ATTRIBUTE in ('ChipQty', 'WaferQty') and" & vbNewLine & _
                        "      A.TXNTIMESTAMP < '" & oRawSqlRecordSet.Value("TXNTIME") & "' and" & vbNewLine & _
                        "      A.TXNSEQUENCE in" & vbNewLine & _
                        "      (select max(B.TXNSEQUENCE)" & vbNewLine & _
                        "         from FWWIPTRANSACTION B" & vbNewLine & _
                        "        where B.WIPID = '" & sChildLotID & "' and B.ACTIVITY = 'ModifyAttribute' and" & vbNewLine & _
                        "              B.ATTRIBUTE = 'ChipQty' and" & vbNewLine & _
                        "              B.TXNTIMESTAMP < '" & oRawSqlRecordSet.Value("TXNTIME") & "'" & vbNewLine & _
                        "       union all" & vbNewLine & _
                        "       select max(B.TXNSEQUENCE)" & vbNewLine & _
                        "         from FWWIPTRANSACTION B" & vbNewLine & _
                        "        where B.WIPID = '" & sChildLotID & "' and B.ACTIVITY = 'ModifyAttribute' and" & vbNewLine & _
                        "              B.ATTRIBUTE = 'WaferQty' and" & vbNewLine & _
                        "              B.TXNTIMESTAMP < '" & oRawSqlRecordSet.Value("TXNTIME") & "')"

'                Set oRawSqlRecordSet2 = oRawSql.SelectRawSqlTxn1("CAT", sSql)
'                Do While Not oRawSqlRecordSet2.EOF
'                    If oRawSqlRecordSet2.Value("ATTRIBUTE") = gsLOT_CUSTOMATTR_CHIP_QTY Then
'                        .SetText miSpdCol_ChildCQty, .MaxRows, oRawSqlRecordSet2.Value("NEWVALUE")
'                    Else
'                        .SetText miSpdCol_ChildWQty, .MaxRows, oRawSqlRecordSet2.Value("NEWVALUE")
'                    End If
'                    oRawSqlRecordSet2.MoveNext
'                Loop
                Set colRS = moProRawSql.QueryDatabase(sSQL)
                For iIdx = 1 To colRS.Count
                    If colRS.Item(iIdx).Item("ATTRIBUTE") = gsLOT_CUSTOMATTR_CHIP_QTY Then
                        .SetText miSpdCol_ChildCQty, .MaxRows, colRS.Item(iIdx).Item("NEWVALUE")
                    Else
                        .SetText miSpdCol_ChildWQty, .MaxRows, colRS.Item(iIdx).Item("NEWVALUE")
                    End If
                Next
            End If
            
            oRawSqlRecordSet.MoveNext
        Loop
        
        If .MaxRows > 0 Then
            Me.cmdPrint.Enabled = True
        Else
            Me.cmdPrint.Enabled = False
        End If
        
        .Col = 1
        Do While .Col <= .MaxCols
            .ColWidth(.Col) = .MaxTextColWidth(.Col) + 1
            .Col = .Col + 1
        Loop
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
    Set oRawSqlRecordSet = Nothing
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

Public Sub GetPgIdAndNameFromFutAct(ByVal oLogCtrl As Object, _
        ByVal oFwWIP As Object, ByVal oFwWF As Object, ByVal oCwMbx As Object, _
        ByVal sLotID As String, ByVal sStepNo As String, ByVal sPath As String, _
        ByRef sPgId As String, ByRef sPgName As String, _
        ByRef sDocNo As String, _
        ByVal sEqType2 As String, ByVal sSubSysType As String, _
        Optional ByRef sPgMode As String)
'sPgMode added by Jack 2009/11/25 for CP Cuto Load Project.
    On Error GoTo ExitHandler:
    Dim sProcID As String
    Dim typErrInfo As tErrInfo
    Dim sTable      As String
    Dim sColumn     As String
    Dim sWhere      As String
    Dim sOther      As String
    Dim colRaws     As Collection
    Dim oRaws       As FwStrings
    '----
    ' Init
    '----
    sProcID = "GetPgIdAndNameFromFutAct"
    Call LogProcIn(msMODULE_ID, sProcID, oLogCtrl)
    sPgId = ""
    sPgName = ""
    sPgMode = "" 'sPgMode added by Jack 2009/11/25 for CP Cuto Load Project.
    '----
    ' Condition Checking
    '----
    '---------------------------------------------------------------------------------------------------------------
    '20070413 Vencent,M200701035
    If InStr(1, sLotID, ".") > 0 Then
        sLotID = Left(sLotID, InStr(1, sLotID, ".") - 1)
    End If
    '---------------------------------------------------------------------------------------------------------------
    
    '----
    ' Action
    '----
    '*************
    'Get PgId and PgName from table "Tbl_Lot_FutAct"
    '*************
    'Define table name
    sTable = gsCAT_TBL_LOT_FUTACT
        
    'Define column name
    'gsCAT_TLFA_PGMODE added by Jack 2009/11/25 for CP Cuto Load Project.
    sColumn = gsCAT_TLFA_PGID & "," & _
        gsCAT_TLFA_PGNAME & "," & _
        gsCAT_TLFA_DOC_NO & "," & _
        gsCAT_TLFA_PGMODE
        
    'Define where
    sWhere = gsCAT_TLFA_LOT_ID & "='" & Trim$(sLotID) & "' and " & _
        gsCAT_TLFA_STEP_NO & "='" & Trim$(sStepNo) & "' and " & _
        gsCAT_TLFA_DELETE_FLAG & "='N'"
        
    '20060906 Vencent,M200608049
    If sEqType2 <> "" Then
        sWhere = sWhere & " and nvl(" & gsCAT_TLFA_EQTYPE2 & ",'" & sEqType2 & "') = '" & sEqType2 & "'"
    End If
    
    '---------------------------------------------------------------------------------------------------------------
    '20070413 Vencent,M200701035
    Dim vSubSys As Variant, iIdx As Integer, sSubWhere As String
    If sSubSysType <> "" Then
        vSubSys = Split(sSubSysType, ",")
        For iIdx = LBound(vSubSys) To UBound(vSubSys)
            If sSubWhere <> "" Then sSubWhere = sSubWhere & " or "
            sSubWhere = sSubWhere & " nvl(" & gsCAT_TLFA_SUBSYSTEM & ",'" & vSubSys(iIdx) & "')='" & vSubSys(iIdx) & "'"
        Next iIdx
        sWhere = sWhere & " and (" & sSubWhere & ")"
    End If
    sWhere = sWhere & " and " & gsCAT_TLFA_HOLD_POSITION & " = '" & gsHOLD_POSITION_STEPIN & "'"
    '---------------------------------------------------------------------------------------------------------------
   
   'add by Sam start on 20090107 for ReqNo:M200812005
   If sPath <> "" Then
    sWhere = sWhere & " and " & gsCAT_TLFA_PATH & " = '" & sPath & "'"
   End If
   'add by Sam end on 20090107 for ReqNo:M200812005
   
   
    sOther = " order by " & gsCAT_TLFA_CREATE_TIME & " desc "
        
    Set colRaws = SelectRawSqlTxn(oLogCtrl, oFwWIP, oFwWF, oCwMbx, "CAT", sTable, sColumn, sWhere, sOther)
    If colRaws Is Nothing Then
        Call RaiseError(glERR_INVALIDOBJECT, _
            FormatErrorText(gsETX_INVALIDOBJECT, "Collection"))
    End If
    
    For Each oRaws In colRaws
        sPgId = Trim$(oRaws.Item(gsCAT_TLFA_PGID))
        sPgName = Trim$(oRaws.Item(gsCAT_TLFA_PGNAME))
        sDocNo = Trim$(oRaws.Item(gsCAT_TLFA_DOC_NO))
        sPgMode = Trim$(oRaws.Item(gsCAT_TLFA_PGMODE)) 'sPgMode added by Jack 2009/11/25 for CP Cuto Load Project.
        Exit For
    Next oRaws
    '*************
    '----
    ' Done
    '----

ExitHandler:
    ' NOTE 1:
    ' MUST CALL GetErrInfo() here first before another action
    Call GetErrInfo(msMODULE_ID, sProcID, typErrInfo, Erl)
    Call LogProcOut(msMODULE_ID, sProcID, typErrInfo, oLogCtrl)
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
        Call HandleError(False, typErrInfo, , oLogCtrl, True)
    End If
End Sub

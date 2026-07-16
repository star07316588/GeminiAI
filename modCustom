Public Function GetStepPath(ByRef moAppLog As Object, _
        ByVal moFwWIP As Object, _
        ByVal moFwWF As Object, _
        ByVal moCwMbx As Object, _
        ByVal sPlanId As String, _
        ByVal sStepSeq As String) As String
    On Error GoTo ExitHandler:
    Dim sProcID As String
    Dim typErrInfo As tErrInfo
    Dim sTable As String
    Dim sColumn As String
    Dim sWhere As String
    Dim sOrder As String
    Dim colRaws As Collection
    Dim oRaws As FwStrings
    '----
    ' Init
    '----
    sProcID = "GetStepPath"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog)
    '----
    ' Condition Chking
    '----


    '----
    ' Action
    '----
    sTable = "fwprocessplan a," & _
        "fwflatplan b," & _
        "fwflatplan_n2m c," & _
        "fwflatnode d," & _
        modConstCAT.gsCAT_TBL_ROUTE_STEP_LIST & " e"
             
    sColumn = "e." & modConstCAT.gsCAT_TRSL_PATH
              
    sWhere = "a.name = '" & sPlanId & "' and " & _
        "b.planname = a.name and " & _
        "b.planversion = a.activeversion and " & _
        "b.RevState ='Frozen' and " & _
        "b.sysid = c.fromid and " & _
        "c.linkname = 'planNodes' and " & _
        "c.toid = d.sysid and " & _
        "d.StepSeq = '" & sStepSeq & "' and " & _
        "e." & modConstCAT.gsCAT_TRSL_ROUTE_ID & "='" & sPlanId & "' and " & _
        "e." & modConstCAT.gsCAT_TRSL_STEP_HANDLE & "= d.handle and " & _
        "e." & modConstCAT.gsCAT_TRSL_DELETE_FLAG & "='N'"
             
    Set colRaws = SelectRawSqlTxn(moAppLog, moFwWIP, moFwWF, moCwMbx, "CAT", sTable, sColumn, sWhere)
    If colRaws Is Nothing Then
        Call RaiseError(glERR_INVALIDOBJECT, _
            FormatErrorText(gsETX_INVALIDOBJECT, "Collection"))
    End If
        
    For Each oRaws In colRaws
        GetStepPath = oRaws.Item(1)
    Next oRaws
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
End Function

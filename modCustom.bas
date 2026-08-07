Public Function NotStopTest(ByVal sLotID As String, ByVal sIPN As String, ByVal sProdgroup As String, ByVal sStopScope As String, _
                            ByVal bRunRule As Boolean, ByVal sStepName As String, ByVal sCurEqId As String, ByVal sEqType2 As String, _
                            ByVal sCurAccName As String, ByVal sStopTicNo As String, ByVal sSpecSpecifyEQ As String, ByVal sSpecEqId As String, _
                            ByVal sErunTicNO As String, ByVal sPgId As String, sPgName As String, ByVal sPgMode As String, ByVal sTemperature As String, _
                            ByRef moFwWIP As Object, ByRef moAppLog As Object, ByRef moProRawSql As Object, _
                            Optional ByRef sStopTicNoList As String, Optional ByRef sSubsystem As String, _
                            Optional ByRef sCheckWsDeviceFile As String, Optional ByRef sStopMessage As String, _
                            Optional ByVal sSpecifyAcc As String) As Boolean
                            
On Error GoTo ExitHandler:
Dim sProcID         As String
Dim typErrInfo      As tErrInfo
Dim sSQL            As String
Dim colRS    As Collection

Dim lIdx As Long

Dim sCurAccNo As String

Dim bIssue As Boolean
Dim bStopAcc As Boolean

Dim sStopAcc As String
Dim sStopTestP As String
Dim sStopTestNP As String

Dim sStage As String

Dim sFollowProd As String
Dim sContinueTest As String

Dim sStepID As String

'Added by Jack on 2017/02/06 for BE MES 生產資訊自動化 Phase 1 : 停復測自動化 <Start>
Dim sStopTestPTemp    As Variant '(將Array sStopTestP 內容拆解)
Dim lIdx1             As Integer
Dim sTECNLotID        As String
Dim sLotContinueTest  As String 'By Lot TECN ContinueTest ?
Dim sPath             As String
Dim bForSetEqNotice   As Boolean 'called by CAT.SetEqNotice ?? (Default : false / SetEqNotice直接呼叫GetProductStopTest( ).
Dim sStopAccNo        As String '停測的AccNo.
Dim bSotpDevice       As Boolean '是否因DeviceFile而停測?
Dim sBePeSet          As String '停測單資訊的BePeSet(BE PE處置 : A,M,A+M)
Dim sWsDeviceFile     As String 'Current WsDeviceFile.
Dim sStopDeviceFile   As String '停測的 WsDeviceFile.
'Added by Jack on 2017/02/06 for BE MES 生產資訊自動化 Phase 1 : 停復測自動化 <End>

'----
' Init
'----
    sProcID = "NotStopTest"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog)
    
    NotStopTest = True
    
    'Added by Jack on 2017/02/06 for BE MES 生產資訊自動化 Phase 1 : 停復測自動化 <Start>
    bForSetEqNotice = False 'called by CAT.SetEqNotice ?? (Default : false / SetEqNotice直接呼叫GetProductStopTest( ).
    sWsDeviceFile = sCheckWsDeviceFile
    sStopMessage = ""
    'Added by Jack on 2017/02/06 for BE MES 生產資訊自動化 Phase 1 : 停復測自動化 <End>
    
    '增加 gsCAT_TLATT_ROUTE : Added by Jack on 2017/02/06 for BE MES 生產資訊自動化 Phase 1 : 停復測自動化.
    sSQL = "select a." & gsCAT_TLATT_STEPID & " , a." & gsCAT_TLATT_STAGE & " " & _
                " ,a." & gsCAT_TLATT_ROUTE & " " & _
          " from " & gsCAT_TBL_LOT_ATTRIBUTE & " a " & _
          " where a." & gsCAT_TLATT_LOTID & " = '" & sLotID & "' "
    Set colRS = moProRawSql.QueryDatabase(sSQL)
    If colRS.Count > 0 Then
        sStage = colRS.Item(1).Item(gsCAT_TLATT_STAGE)
        sStepID = colRS.Item(1).Item(gsCAT_TLATT_STEPID)
        sPath = colRS.Item(1).Item(gsCAT_TLATT_ROUTE)
    End If
    
    'Added by Jack on 2017/02/06 for BE MES 生產資訊自動化 Phase 1 : 停復測自動化 <Start>
    sSQL = "select a." & gsCAT_TLI_TECN_LOT_ID & " " & _
          " from " & gsCAT_TBL_LOT_INFO & " a " & _
          " where a." & gsCAT_TLI_LOT_ID & " = '" & sLotID & "' "
    Set colRS = moProRawSql.QueryDatabase(sSQL)
    If colRS.Count > 0 Then
        sTECNLotID = colRS.Item(1).Item(gsCAT_TLI_TECN_LOT_ID)
    End If
    If Trim(sTECNLotID) <> "" Then
        sSQL = "select a." & gsCAT_TLSPES_CONTINUETEST & " " & _
              " from " & gsCAT_TBL_LOT_STEP_EQ_SPEC & " a " & _
              " where '" & sTECNLotID & "' like a." & gsCAT_TLSPES_TECNLOTID & " " & _
                " and a." & gsCAT_TLSPES_STEPNO & " = '" & sStepID & "' " & _
                " and a." & gsCAT_TLSPES_PATH & " = '" & sPath & "' " & _
                " and a." & gsCAT_TLSPES_EQTYPE2 & " = '" & sEqType2 & "' " & _
                " and nvl(trim(a." & gsCAT_TLSPES_SUBSYSTEM & "), ' ') = " & _
                "     nvl(trim('" & sSubsystem & "'), ' ') " & _
                " and a." & gsCAT_TLSPES_DELETEFLAG & " = 'N' "
        Set colRS = moProRawSql.QueryDatabase(sSQL)
        If colRS.Count > 0 Then
            sLotContinueTest = colRS.Item(1).Item(gsCAT_TLSPES_CONTINUETEST)
        End If
    End If
    'Added by Jack on 2017/02/06 for BE MES 生產資訊自動化 Phase 1 : 停復測自動化 <End>
    
    If bRunRule = True Then
        'gsCAT_TEI_WSDEVICE_FILE : Added by Jack on 2017/02/06 for BE MES 生產資訊自動化 Phase 1 : 停復測自動化
        sSQL = "select a." & gsCAT_TEI_PROBECARD_ID & " , a." & gsCAT_TEI_LOADBOARD_ID & " , a." & gsCAT_TEI_CONTACTBOARD_ID & " " & _
                      ", a." & gsCAT_TEI_WSDEVICE_FILE & " " & _
               " from " & gsCAT_TBL_EQ_INFO & " a " & _
               " where a." & gsCAT_TEI_EQ_ID & " = '" & sCurEqId & "' "
        Set colRS = moProRawSql.QueryDatabase(sSQL)
        If colRS.Count > 0 Then
            'Added by Jack on 2017/02/06 for BE MES 生產資訊自動化 Phase 1 : 停復測自動化 <Start>
            sWsDeviceFile = Trim(colRS.Item(1).Item(gsCAT_TEI_WSDEVICE_FILE))
            'Added by Jack on 2017/02/06 for BE MES 生產資訊自動化 Phase 1 : 停復測自動化 <End>
            
            If colRS.Item(1).Item(gsCAT_TEI_PROBECARD_ID) <> "" Then
                sCurAccName = Left(colRS.Item(1).Item(gsCAT_TEI_PROBECARD_ID), InStr(colRS.Item(1).Item(gsCAT_TEI_PROBECARD_ID), "-") - 1)
                sCurAccNo = sCurAccNo & "," & colRS.Item(1).Item(gsCAT_TEI_PROBECARD_ID)
            End If
                        
            If colRS.Item(1).Item(gsCAT_TEI_LOADBOARD_ID) <> "" Then
                sCurAccName = Left(colRS.Item(1).Item(gsCAT_TEI_LOADBOARD_ID), InStr(colRS.Item(1).Item(gsCAT_TEI_LOADBOARD_ID), "-") - 1)
                sCurAccNo = sCurAccNo & "," & colRS.Item(1).Item(gsCAT_TEI_LOADBOARD_ID)
            End If
            If colRS.Item(1).Item(gsCAT_TEI_CONTACTBOARD_ID) <> "" Then
                sCurAccName = Left(colRS.Item(1).Item(gsCAT_TEI_CONTACTBOARD_ID), InStr(colRS.Item(1).Item(gsCAT_TEI_CONTACTBOARD_ID), "-") - 1)
                sCurAccNo = sCurAccNo & "," & colRS.Item(1).Item(gsCAT_TEI_CONTACTBOARD_ID)
            End If
        End If
        If Left(sCurAccNo, 1) = "," Then
            sCurAccNo = Mid(sCurAccNo, 2)
        End If
        
        If Right(sCurAccNo, 1) = "," Then
            sCurAccNo = Left(sCurAccNo, Len(sCurAccNo) - 1)
        End If
    End If
    
    'Add by Sam start on 20170622 for SortRun Lot Start ,強制指定配件
    If sSpecifyAcc <> "" Then
        sCurAccNo = sSpecifyAcc
        If InStr(sSpecifyAcc, "-") > 0 Then
            sCurAccName = Left(sSpecifyAcc, InStr(sSpecifyAcc, "-") - 1)
        Else
            sCurAccName = sSpecifyAcc
        End If
    End If
    'Add by Sam end on 20170622 for SortRun Lot Start
    
    If sErunTicNO <> "" Then
       sSQL = " select a." & gsCAT_TER_FOLLOW_PRODUCT & ", " & _
            " a." & gsCAT_TER_CONTINUETEST & " " & _
            " from " & gsCAT_TBL_ERUN_REQ & " a " & _
            " where a." & gsCAT_TER_LOT_ID & "= '" & sLotID & "' " & _
            " and a." & gsCAT_TER_TICKET_NO & "= '" & sErunTicNO & "' " & _
            " and a." & gsCAT_TER_STAGE & "= '" & sStage & "' "
        Set colRS = moProRawSql.QueryDatabase(sSQL)
        If colRS.Count > 0 Then
           sFollowProd = colRS.Item(1).Item(gsCAT_TER_FOLLOW_PRODUCT)
           sContinueTest = colRS.Item(1).Item(gsCAT_TER_CONTINUETEST)
        End If
    End If
'----
' Condition Checking
'----
    ' <Put your condition checking codes here>...

'----
' Action
'----
      
    Select Case sStopScope
        Case gsCAT_STOP_TEST_ALL
            
            'Added by Jack on 2017/02/06 for BE MES 生產資訊自動化 Phase 1 : 停復測自動化 <Start>
            'LotContinueTest= 'Y' 則不需要檢核產品相關停測.
            If sLotContinueTest <> "Y" Then
            'Added by Jack on 2017/02/06 for BE MES 生產資訊自動化 Phase 1 : 停復測自動化 <End>
                If sContinueTest <> "Y" Then 'ContinueTest = 'Y' 則不需要檢核產品相關停測
                    If sFollowProd = "N" Then '有委測單但FollowProduction='N', ContinueTest = 'N' 則 call GetProdStopTest() ,工程品不卡產品相關
    '                    Call GetProductStopTest(sProdGroup, sEqType2, sPgId, sPGName, sPgMode, sStepName _
    '                                            , sStopTestP, sSpecSpecifyEQ, sSpecEQID, bIssue, sCurAccName, bStopAcc, sStopAcc, _
    '                                            moFwWIP, moAppLog, moProRawSql)
                                                
                    Else '無委測單或有委測單但FollowProduction='Y', ContinueTest = 'N' 則 call ChkProdSpecEqID()
                        If ChkProdSpecEqID(sSpecSpecifyEQ, sSpecEqId, sCurEqId, moAppLog) = False Then
                            sStopTicNoList = sStopTicNo
                            NotStopTest = False
                            'Added by Jack on 2017/02/06 for BE MES 生產資訊自動化 Phase 1 : 停復測自動化 <Start>
                            If Trim(sTECNLotID) <> "" Then
                                sStopMessage = "LotStop"
                            Else
                                sStopMessage = "ProdGroupStop(" & sProdgroup & ")"
                            End If
                            'Added by Jack on 2017/02/06 for BE MES 生產資訊自動化 Phase 1 : 停復測自動化 <End>
                        End If
                        
                        'Added by Jack on 2017/02/06 for BE MES 生產資訊自動化 Phase 1 : 停復測自動化 <Start>
                        '上面的 "產品相關停測" 檢查, 是檢核 SetEqNotice Update 的 tbl_prod_step_eq_spec 資料.
                        '這裡, 再加串 "產品相關停測".
                        'ConCatAcc 傳入 sCurAccName 即可(SetEqNotice 為 A.Probecardtype || A.Loadboardtype || A.Contactboardtype || A.Burninboard 加總).
                        '取回 sStopTestP '停測單號
                        '加傳 sIPN Added by Jack on 2017/08/09 for BE#201700328 requested by Livan (08/08 mail)
                        Call GetProductStopTest(moAppLog, sProdgroup, sEqType2, sPgId, sPgName, sPgMode, sStepName, _
                                                sStopTicNo, sSpecSpecifyEQ, sSpecEqId, bIssue, sCurAccName, bStopAcc, sStopAcc, _
                                                moProRawSql, bForSetEqNotice, sLotID, sCurAccNo, sStopAccNo, _
                                                bSotpDevice, sBePeSet, sWsDeviceFile, sStopDeviceFile, sCurEqId, sIPN)
                        
                        '記錄所有停測單號，重覆的不計入
                        sStopTestP = sStopTicNo
                        sStopTestPTemp = Split(sStopTestP, ",")
                        
                        For lIdx1 = LBound(sStopTestPTemp) To UBound(sStopTestPTemp)
                            If GetStrPosInArray(sStopTicNoList, sStopTestPTemp(lIdx1)) = 0 Then
                                sStopTicNoList = sStopTicNoList & "," & sStopTestPTemp(lIdx1)
                            End If
                        Next lIdx1
                        'Added by Jack on 2017/02/06 for BE MES 生產資訊自動化 Phase 1 : 停復測自動化 <End>
                    End If
                End If
            End If
                                                
            Call GetNonProdStopTest(sIPN, sEqType2, sCurEqId, sCurAccName, sCurAccNo, _
                                    sTemperature, sStopTestNP, moFwWIP, moAppLog, moProRawSql)
                                
            'Marked by Jack on 2017/02/06 for BE MES 生產資訊自動化 Phase 1 : 停復測自動化 <Start>
            'If sStopTestP <> "" Then
            '    sStopTicNoList = sStopTestP
            'End If
            'Marked by Jack on 2017/02/06 for BE MES 生產資訊自動化 Phase 1 : 停復測自動化 <End>
            
            If sStopTestNP <> "" Then
                sStopTicNoList = sStopTicNoList & "," & sStopTestNP
            End If
                                                            
            If Left(sStopTicNoList, 1) = "," Then
                sStopTicNoList = Mid(sStopTicNoList, 2)
            End If
                                                            
        Case gsCAT_STOP_TEST_NP
            
            Call GetNonProdStopTest(sIPN, sEqType2, sCurEqId, sCurAccName, sCurAccNo, _
                                    sTemperature, sStopTicNoList, moFwWIP, moAppLog, moProRawSql)
        Case gsCAT_STOP_TEST_P
            'Added by Jack on 2017/02/06 for BE MES 生產資訊自動化 Phase 1 : 停復測自動化 <Start>
            'LotContinueTest= 'Y' 則不需要檢核產品相關停測.
            If sLotContinueTest <> "Y" Then
            'Added by Jack on 2017/02/06 for BE MES 生產資訊自動化 Phase 1 : 停復測自動化 <End>
                If sContinueTest <> "Y" Then 'ContinueTest = 'Y' 則不需要檢核產品相關停測
                    If sFollowProd = "N" Then '有委測單但FollowProduction='N', ContinueTest = 'N' 則 call GetProdStopTest() ,工程品不卡產品相關
    '                    Call GetProductStopTest(sProdGroup, sEqType2, sPgId, sPGName, sPgMode, sStepName _
    '                                            , sStopTestP, sSpecSpecifyEQ, sSpecEQID, bIssue, sCurAccName, bStopAcc, sStopAcc, _
    '                                            moFwWIP, moAppLog, moProRawSql)
                                                
                    Else '無委測單或有委測單但FollowProduction='Y', ContinueTest = 'N' 則 call ChkProdSpecEqID()
                        If ChkProdSpecEqID(sSpecSpecifyEQ, sSpecEqId, sCurEqId, moAppLog) = False Then
                            sStopTicNoList = sStopTicNo
                        End If
                        
                        'Added by Jack on 2017/02/06 for BE MES 生產資訊自動化 Phase 1 : 停復測自動化 <Start>
                        '上面的 "產品相關停測" 檢查, 是檢核 SetEqNotice Update 的 tbl_prod_step_eq_spec 資料.
                        '這裡, 再加串 "產品相關停測".
                        'ConCatAcc 傳入 sCurAccName 即可(SetEqNotice 為 A.Probecardtype || A.Loadboardtype || A.Contactboardtype || A.Burninboard 加總).
                        'sStopTestP '停測單號
                        Call GetProductStopTest(moAppLog, sProdgroup, sEqType2, sPgId, sPgName, sPgMode, sStepName, _
                                                sStopTicNo, sSpecSpecifyEQ, sSpecEqId, bIssue, sCurAccName, bStopAcc, sStopAcc, _
                                                moProRawSql, bForSetEqNotice, sLotID, sCurAccNo, sStopAccNo, _
                                                bSotpDevice, sBePeSet, sWsDeviceFile, sStopDeviceFile, sCurEqId)
                        'Added by Jack on 2017/02/06 for BE MES 生產資訊自動化 Phase 1 : 停復測自動化 <End>
                    End If
                End If
            End If
            If Left(sStopTicNoList, 1) = "," Then
                sStopTicNoList = Mid(sStopTicNoList, 2)
            End If
            
    End Select
    
    If sStopTicNoList <> "" Then
        NotStopTest = False
    End If
    
    'Added by Jack on 2017/02/06 for BE MES 生產資訊自動化 Phase 1 : 停復測自動化 <Start>
    sStopMessage = Trim(sStopMessage) & "," & Trim(sStopTicNoList)
    If Left(sStopMessage, 1) = "," Then '原本sStopMessage為空值, 取停測單號.
        sStopMessage = Mid(sStopMessage, 2)
    End If
    If Right(sStopMessage, 1) = "," Then '原本sStopTicNoList為空值, 去掉多餘的 ",".
        sStopMessage = Mid(sStopMessage, 1, Len(sStopMessage) - 1)
    End If
    'Added by Jack on 2017/02/06 for BE MES 生產資訊自動化 Phase 1 : 停復測自動化 <End>
      
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

Public Function ChkProdSpecEqID(sSpecifyEQ As String, _
        sProdSpecEqId As String, _
        sCurEqId As String, _
        moAppLog As Object) As Boolean
    On Error GoTo ExitHandler:
    Dim sProcID As String
    Dim typErrInfo As tErrInfo

    Dim iIdx                    As Integer
    Dim sProdSpecEqIdArray()    As String
    Dim bExist                  As Boolean
    '----
    ' Init
    '----
    sProcID = "ChkProdSpecEqID"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog) '"Entering Function...", moAppLog, glLOG_PROC, msMODULE_ID, sProcID)

    '----
    ' Condition Checking
    '----
    ' <Put your condition checking codes here>...
    If sSpecifyEQ = "" Then
        ChkProdSpecEqID = True
        GoTo ExitHandler
    Else
        ChkProdSpecEqID = False
    End If
    '----
    ' Action
    '----
    ' <Put your Action codes here>...
    sProdSpecEqIdArray = Split(sProdSpecEqId, ",")
    bExist = False
    For iIdx = LBound(sProdSpecEqIdArray) To UBound(sProdSpecEqIdArray)
        If Trim(sProdSpecEqIdArray(iIdx)) = sCurEqId Then
            bExist = True
            Exit For
        End If
    Next iIdx
    
    If sSpecifyEQ = "Y" And bExist = True Then
        ChkProdSpecEqID = True
    ElseIf sSpecifyEQ = "N" And bExist = False Then
        ChkProdSpecEqID = True
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
End Function

Public Function GetProductStopTest(ByRef moAppLog As Object, ByVal sProdgroup As String, ByVal sEqType2 As String, _
                                    ByVal sPgId As String, ByVal sPgName As String, ByVal sPgMode As String, _
                                    ByVal sStepName As String, ByRef sStopTicNo As String, _
                                    ByRef sSpecifyEQ As String, ByRef sEqID As String, _
                                    ByRef bIssue As Boolean, ByVal ConCatAcc As String, _
                                    ByRef bStopAcc As Boolean, ByRef sAccName As String, _
                                    ByRef moProRawSql As Object, _
                                    Optional ByVal bForSetEqNotice As Boolean = False, Optional ByRef sLotID As String = "", _
                                    Optional ByRef sCurAccNo As String = "", Optional ByRef sStopAccNo As String = "", _
                                    Optional ByRef bStopDevice As Boolean = False, Optional ByRef sBePeSet As String = "", _
                                    Optional ByRef sWsDeviceFile As String = "", Optional ByRef sStopDeviceFile As String = "", _
                                    Optional ByRef sCurEqId As String = "", Optional ByRef sIPN As String = "") As Boolean
On Error GoTo ExitHandler:
Dim sProcID As String
Dim typErrInfo As tErrInfo
Dim sSQL As String
Dim colRaws As Collection
Dim lIdx As Long
Dim lIdx2 As Long

Dim sTmpEqID As String
Dim vEqID As Variant

Dim vRemoveEQ   As Variant

'Added by Jack on 2017/02/06 for BE MES 生產資訊自動化 Phase 1 : 停復測自動化 <Start>
'sCurEqId : run rule 傳入的Current EQID, 請勿修改它(也不要清空它)
Dim vStopTicNoAry   As Variant
Dim sBePeSet_Temp   As Variant
Dim sBePeSet_Total  As Variant
Dim bBePeSet_A      As Boolean
Dim bBePeSet_M      As Boolean
Dim dicTicNoVsEQ    As Dictionary
Dim sKey            As String
Dim sOldValue       As String
Dim sNewValue       As String
Dim vTmpTicNoAry    As Variant
Dim bRemoveStopTicY As Boolean
Dim bRemoveStopTicN As Boolean
'Added by Jack on 2017/02/06 for BE MES 生產資訊自動化 Phase 1 : 停復測自動化 <End>

'----
' Init
'----
    sProcID = "GetProductStopTest"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog)
    GetProductStopTest = True
    
    'Add by Sam start on 20250205 for  #203688,虛擬批號取.之前
    If InStr(sLotID, ".") > 0 Then
        sLotID = Left(sLotID, InStr(sLotID, ".") - 1)
    End If
    'Add by Sam end on 20250205 for  #203688
        
'----
' Condition Checking
'----
    ' <Put your condition checking codes here>...

'----
' Action
'----
    bIssue = False 'Check是否同時有指定可測及不可測機台
    sStopTicNo = ""
    sSpecifyEQ = ""
    sEqID = ""
    bStopAcc = False
    
    'Added by Jack on 2017/02/06 for BE MES 生產資訊自動化 Phase 1 : 停復測自動化 <Start>
    sStopAccNo = ""
    sStopDeviceFile = ""
    bStopDevice = False
    sBePeSet = ""
    sBePeSet_Total = ""
    sBePeSet_Temp = ""
    Set dicTicNoVsEQ = CreateObject("Scripting.Dictionary")
    'Added by Jack on 2017/02/06 for BE MES 生產資訊自動化 Phase 1 : 停復測自動化 <End>
    
    '機台配件停復測系統建構專案(新增串AccName)
    'Added by Jack on 2017/02/06 for BE MES 生產資訊自動化 Phase 1 : 停復測自動化 <Start>
    '增加欄位 :  d. gsCAT_TSTT_ACCNO & gsCAT_TSTT_DEVICEFILE, gsCAT_TSTT_BEPE_SET.
    '增加欄位 : dd. gsCAT_TIM_IPN, gsCAT_TIM_PRODBODY, gsCAT_TIM_PRODGROUPKEY
    'Added by Jack on 2017/02/06 for BE MES 生產資訊自動化 Phase 1 : 停復測自動化 <End>
    'Modified by Jack on 2017/07/10 for BE MES Phase 15 : [BE#201700328] CAT Set Eq Notice & MES spec query 停復測資訊連結 rule 變更 (SetEqNotice / SpecQuery) <Start>
    'before : " from (select aa.productname prodgroup, cc.magnitude prodcode, dd." & gsCAT_TIM_IPN & ", "
    '                      " dd." & gsCAT_TIM_PRODBODY & ", dd." & gsCAT_TIM_PRODGROUPKEY & " "
    'after  : " from (select aa.productname prodgroup, dd." & gsCAT_TIM_IPN & ", "
    '                      " substr(nvl(cc." & gsCAT_TPS_PROD_GROUP_KEY & ", dd." & gsCAT_TIM_PRODGROUPKEY & "), 1, 5) as prodcode, "
    '                      " substr(nvl(cc." & gsCAT_TPS_PROD_GROUP_KEY & ", dd." & gsCAT_TIM_PRODGROUPKEY & "), 1, 4) as prodbody, "
    '                      " nvl(cc." & gsCAT_TPS_PROD_GROUP_KEY & ", dd." & gsCAT_TIM_PRODGROUPKEY & ") as prodgroupkey "
    'Modified by Jack on 2017/07/10 for BE MES Phase 15 : [BE#201700328] CAT Set Eq Notice & MES spec query 停復測資訊連結 rule 變更 (SetEqNotice / SpecQuery) <End>
    sSQL = " select distinct d." & gsCAT_TSTT_STOPTICNO & ", d." & gsCAT_TSTT_SPECIFYEQ & ", d." & gsCAT_TSTT_EQID & ",d." & gsCAT_TSTT_ACCNAME & _
                    ",d." & gsCAT_TSTT_ACCNO & ",d." & gsCAT_TSTT_DEVICEFILE & ",d." & gsCAT_TSTT_BEPE_SET & _
           " from (select aa.productname prodgroup, dd." & gsCAT_TIM_IPN & ", " & _
                        " substr(nvl(cc." & gsCAT_TPS_PROD_GROUP_KEY & ", dd." & gsCAT_TIM_PRODGROUPKEY & "), 1, 5) as prodcode, " & _
                        " substr(nvl(cc." & gsCAT_TPS_PROD_GROUP_KEY & ", dd." & gsCAT_TIM_PRODGROUPKEY & "), 1, 4) as prodbody, " & _
                        " nvl(cc." & gsCAT_TPS_PROD_GROUP_KEY & ", dd." & gsCAT_TIM_PRODGROUPKEY & ") as prodgroupkey "
                        
    'Added by Jack on 2017/02/06 for BE MES 生產資訊自動化 Phase 1 : 停復測自動化 <Start>
    '增加欄位 :   gsCAT_TSTT_LOTID, gsCAT_TSTT_PRODGROUPKEY, gsCAT_TSTT_PRODBODY,
    '             gsCAT_TSTT_ACCNO, gsCAT_TSTT_DEVICEFILE, gsCAT_TSTT_BEPE_SET
    'Added by Jack on 2017/02/06 for BE MES 生產資訊自動化 Phase 1 : 停復測自動化 <End>
    'Modified by Jack on 2017/07/10 for BE MES Phase 15 : [BE#201700328] CAT Set Eq Notice & MES spec query 停復測資訊連結 rule 變更 (SetEqNotice / SpecQuery) <Start>
    'before : "   fwprpattributeinstance cc,"
    'after  : "   " & gsCAT_TBL_PROD_SPEC & " cc,"
    '---------------------------------------------------
    'before : "   and bb.toid = cc.sysid" --> 取消!!
    'after  : "   and aa.productname = cc.prodgroup(+) "
    '         "   and cc." & gsCAT_TPS_DOC_STATUS & " = 'Active' "
    '         "   and cc." & gsCAT_TPS_DELETE_FLAG & "='N' "
    'Modified by Jack on 2017/07/10 for BE MES Phase 15 : [BE#201700328] CAT Set Eq Notice & MES spec query 停復測資訊連結 rule 變更 (SetEqNotice / SpecQuery) <End>
    sSQL = sSQL & _
                 "   from fwproductversion       aa," & _
                      "   fwproductversion_n2m   bb," & _
                    "   " & gsCAT_TBL_PROD_SPEC & " cc," & _
                      "   " & gsCAT_TBL_IPN_MASTER & "        dd" & _
                "   where bb.linkname = 'attributes'" & _
                  "   and bb.fromid = aa.sysid" & _
                  "   and aa.productname = cc.prodgroup(+) " & _
                  "   and cc." & gsCAT_TPS_DOC_STATUS & " = 'Active' " & _
                  "   and cc." & gsCAT_TPS_DELETE_FLAG & "='N' " & _
                  "   and bb.keydata = 'TdsProd'" & _
                  "   and aa.revstate = 'Active'" & _
                  "   and aa.productname = dd." & gsCAT_TIM_PROD_GROUP & "(+)) c," & _
              "   (select " & gsCAT_TSTT_STOPTICNO & ", " & gsCAT_TSTT_PRODCODE & ", " & gsCAT_TSTT_IPN & "," & gsCAT_TSTT_STEPNAME & " " & _
                "," & gsCAT_TSTT_EQTYPE2 & ", " & gsCAT_TSTT_SPECIFYEQ & ", " & gsCAT_TSTT_EQID & _
                "," & gsCAT_TSTT_PGMODE & ", " & gsCAT_TSTT_PGID & ", " & gsCAT_TSTT_PGNAME & "," & gsCAT_TSTT_ACCNAME & _
                "," & gsCAT_TSTT_LOTID & ", " & gsCAT_TSTT_PRODGROUPKEY & ", " & gsCAT_TSTT_PRODBODY & _
                "," & gsCAT_TSTT_ACCNO & ", " & gsCAT_TSTT_DEVICEFILE & ", " & gsCAT_TSTT_BEPE_SET & " "

    sSQL = sSQL & _
                " from " & gsCAT_TBL_STOP_TEST & " " & _
                "   where " & gsCAT_TSTT_DELETEFLAG & " = 'N' "
                
        'Modified by Jack on 2017/02/06 for BE MES 生產資訊自動化 Phase 1 : 停復測自動化 <Start>
        ' 原本只判斷 DeleteFlag = "N"
        '   <a> 加傳 bForSetEqNotice : CAT : SetEqNotice   不管 BePe_Set ("M"/"A" 都看) 設定 (DeleteFlag = "N")
        '                              其他(MES RunRule, 包含SpecQuery)只看 BePe_Set = "A" 的設定 (DeleteFlag = "N")
        If Not bForSetEqNotice Then
            sSQL = sSQL & " and " & gsCAT_TSTT_BEPE_SET & " = 'A' "
        End If
                
        sSQL = sSQL & " ) d "
           
        'Modified by Jack on 2017/02/06 for BE MES 生產資訊自動化 Phase 1 : 停復測自動化 <End>
                        
        'Modified by Jack on 2017/02/06 for BE MES 生產資訊自動化 Phase 1 : 停復測自動化 <Start>
        '修改 <1> IPN 萬用字元 取代原則.(原本rule如下:)
        '"        and nvl(c.ipn, ' ') like nvl(REPLACE(d." & gsCAT_TSTT_IPN & ",'*','%'), nvl(c.ipn, ' '))"
        '     <2> accname 判斷增加 IN.
        '         and nvl('" & ConCatAcc & "', ' ') = nvl(d.accname, nvl('" & ConCatAcc & "', ' '))"
        'Modified by Jack on 2017/02/06 for BE MES 生產資訊自動化 Phase 1 : 停復測自動化 <End>
        'Modify by Sam on 20180503 for ReqNo:BE201800229,取消pgmode條件 ,增加pgname
'        sSQL = sSQL & " where '" & sEqType2 & "' = d.eqtype2  " & _
'          "   and '" & sProdGroup & "'= c.prodgroup" & _
'          "   and nvl(c.ipn, ' ') like nvl(REPLACE(REPLACE(d." & gsCAT_TSTT_IPN & ", '%', '_'),'*','%'), nvl(c.ipn, ' '))" & _
'          "   and nvl('" & sPgId & "', ' ') = nvl(d." & gsCAT_TSTT_PGID & ", nvl('" & sPgId & "', ' '))" & _
'          "   and nvl('" & sPgMode & "', ' ') = nvl(d." & gsCAT_TSTT_PGMODE & ", nvl('" & sPgMode & "', ' '))" & _
'          "   and nvl(c.prodcode, ' ') = nvl(d." & gsCAT_TSTT_PRODCODE & ", nvl(c.prodcode, ' '))" & _
'          "   and '" & sStepName & "' = nvl(d." & gsCAT_TSTT_STEPNAME & ", '" & sStepName & "')" & _
'          "   and ( nvl('" & ConCatAcc & "', ' ') = nvl(d.accname, nvl('" & ConCatAcc & "', ' ')) or " & _
'          "         InStr(d.accname, '" & ConCatAcc & "') > 0 ) "
        sSQL = sSQL & " where '" & sEqType2 & "' = d.eqtype2  " & _
          "   and '" & sProdgroup & "'= c.prodgroup" & _
          "   and nvl(c.ipn, ' ') like nvl(REPLACE(REPLACE(d." & gsCAT_TSTT_IPN & ", '%', '_'),'*','%'), nvl(c.ipn, ' '))" & _
          "   and nvl('" & sPgId & "', ' ') = nvl(d." & gsCAT_TSTT_PGID & ", nvl('" & sPgId & "', ' '))" & _
          "   and nvl('" & sPgName & "', ' ') = nvl(d." & gsCAT_TSTT_PGNAME & ", nvl('" & sPgName & "', ' '))" & _
          "   and nvl(c.prodcode, ' ') = nvl(d." & gsCAT_TSTT_PRODCODE & ", nvl(c.prodcode, ' '))" & _
          "   and '" & sStepName & "' = nvl(d." & gsCAT_TSTT_STEPNAME & ", '" & sStepName & "')" & _
          "   and ( nvl('" & ConCatAcc & "', ' ') = nvl(d.accname, nvl('" & ConCatAcc & "', ' ')) or " & _
          "         InStr(d.accname, '" & ConCatAcc & "') > 0 ) "
          
        'Added by Jack on 2017/02/06 for BE MES 生產資訊自動化 Phase 1 : 停復測自動化 <Start>
        '--LotID/IPN/ProdGroupKey/ProdCode/ProdBody/PGM --> 6選一
        '  "   and ( nvl('" & sCurAccNo & "', ' ') = nvl(d." & gsCAT_TSTT_ACCNO & ", nvl('" & sCurAccNo & "', ' ')) or "
        '  "         InStr(d." & gsCAT_TSTT_ACCNO & ", '" & sCurAccNo & "') > 0 ) "
        '-- 1碼萬用字元，請使用「%」，且不能放在最後一碼。
        '--全碼萬用字元，請使用「*」，且只能放在最後一碼。
        sSQL = sSQL & _
          "   and nvl('" & sLotID & "', ' ') like nvl(REPLACE(REPLACE(d." & gsCAT_TSTT_LOTID & ", '%', '_'),'*','%'), nvl('" & sLotID & "', ' '))" & _
          "   and nvl(c.prodgroupkey, ' ') = nvl(d." & gsCAT_TSTT_PRODGROUPKEY & ", nvl(c.prodgroupkey, ' '))" & _
          "   and nvl(c.ProdBody, ' ') = nvl(d." & gsCAT_TSTT_PRODBODY & ", nvl(c.ProdBody, ' '))" & _
          "   and ( nvl('" & sCurAccNo & "', ' ') = nvl(d." & gsCAT_TSTT_ACCNO & ", nvl('" & sCurAccNo & "', ' ')) or " & _
          "         InStr('" & sCurAccNo & "' || ',' ,d." & gsCAT_TSTT_ACCNO & " ||',') > 0 ) " & _
          "   and nvl('" & sWsDeviceFile & "', ' ') = nvl(d." & gsCAT_TSTT_DEVICEFILE & ", nvl('" & sWsDeviceFile & "', ' '))"
'InStr('LDDC85175T48(6542AD)-9,LDDC85175T48(6542AD)-11'||',', d.accno||',') > 0)
        'Added by Jack on 2017/02/06 for BE MES 生產資訊自動化 Phase 1 : 停復測自動化 <End>
        
        'Added by Jack on 2017/08/09 for BE#201700328 requested by Livan (08/08 mail) <Start>
        sSQL = sSQL & _
                  "   and nvl('" & sIPN & "', ' ') like nvl(REPLACE(REPLACE(d." & gsCAT_TSTT_IPN & ", '%', '_'),'*','%'), nvl('" & sIPN & "', ' '))"
        'Added by Jack on 2017/08/09 for BE#201700328 requested by Livan (08/08 mail) <End>
        
        'Modified by Jack on 2017/02/06 for BE MES 生產資訊自動化 Phase 1 : 停復測自動化
        '將 Order 指令 獨立出來.
        '20170929, bug fix by Robin, sSQL = sSQL & " order by d." & gsCAT_TSTT_ACCNAME & ", d." & gsCAT_TSTT_SPECIFYEQ & " desc, decode(d." & gsCAT_TSTT_EQID & ", null, 'A', 'B')||" & gsCAT_TSTT_STOPTICNO
        sSQL = sSQL & " order by d." & gsCAT_TSTT_SPECIFYEQ & " desc, decode(d." & gsCAT_TSTT_EQID & ", null, 'A', 'B')||" & gsCAT_TSTT_STOPTICNO & ",d." & gsCAT_TSTT_ACCNAME & " desc"

   
    Set colRaws = moProRawSql.QueryDatabase(sSQL)
    
    
    If colRaws.Count > 0 Then
        'Added by Jack on 2017/02/06 for BE MES 生產資訊自動化 Phase 1 : 停復測自動化 <Start>
        If sCurEqId <> "" Then
            Dim a As Dictionary
            For lIdx = 1 To colRaws.Count
                sStopTicNo = colRaws.Item(lIdx).Item(gsCAT_TSTT_STOPTICNO)
                sSpecifyEQ = colRaws.Item(lIdx).Item(gsCAT_TSTT_SPECIFYEQ)
                sEqID = colRaws.Item(lIdx).Item(gsCAT_TSTT_EQID)
                
                sKey = sStopTicNo & "/" & sSpecifyEQ
                If dicTicNoVsEQ.Exists(sKey) Then
                    sOldValue = dicTicNoVsEQ.Item(sKey)
                Else
                    sOldValue = ""
                End If
                
                sNewValue = sOldValue & "," & sEqID
                If Left(sNewValue, 1) = "," Then
                    sNewValue = Mid(sNewValue, 2)
                End If
                'Added by Jack on 2017/06/28 for bug fix <Start>
                If dicTicNoVsEQ.Exists(sKey) Then
                    dicTicNoVsEQ.Remove sKey
                End If
                'Added by Jack on 2017/06/28 for bug fix <End>
                dicTicNoVsEQ.Add sKey, sNewValue
            Next lIdx
        End If
        'Added by Jack on 2017/02/06 for BE MES 生產資訊自動化 Phase 1 : 停復測自動化 <End>
        
        sStopTicNo = colRaws.Item(1).Item(gsCAT_TSTT_STOPTICNO)
        sSpecifyEQ = colRaws.Item(1).Item(gsCAT_TSTT_SPECIFYEQ)
        sEqID = colRaws.Item(1).Item(gsCAT_TSTT_EQID)
        sAccName = colRaws.Item(1).Item(gsCAT_TSTT_ACCNAME)
        
        'Added by Jack on 2017/02/06 for BE MES 生產資訊自動化 Phase 1 : 停復測自動化 <Start>
        sStopAccNo = colRaws.Item(1).Item(gsCAT_TSTT_ACCNO)
        sStopDeviceFile = colRaws.Item(1).Item(gsCAT_TSTT_DEVICEFILE)
        
        If sStopDeviceFile <> "" Then
            bStopDevice = True
        End If
        'Added by Jack on 2017/02/06 for BE MES 生產資訊自動化 Phase 1 : 停復測自動化 <End>
        
        'Modified by Jack on 2017/02/06 for BE MES 生產資訊自動化 Phase 1 : 停復測自動化
        '新增 : sStopAccNo.
        '
        '表示此 recipe 停測
        If sSpecifyEQ = "Y" And sEqID = "" Then
            If sAccName <> "" Or sStopAccNo <> "" Then
                bStopAcc = True
            End If
        End If
        
        'Added by Jack on 2017/05/12 for BE MES 生產資訊自動化 Phase 1 : 停復測自動化 <Start>
        If sSpecifyEQ = "N" And (sAccName <> "" Or sStopAccNo <> "") Then
            bStopAcc = True
        End If
        'Added by Jack on 2017/05/12 for BE MES 生產資訊自動化 Phase 1 : 停復測自動化 <End>
        
'
'        If sSpecifyEq = "Y" And sEqID <> "" Then '先看第一筆
'            For lIdx = 2 To colRaws.Count '繼續處理第二筆之後的紀錄
'                If colRaws.Item(lIdx).Item(gsCAT_TSTT_SPECIFYEQ) = "N" Then '處理"不可測"機台,要從清單移除
'                    bIssue = True
'                    vRemoveEQ = Split(colRaws.Item(lIdx).Item(gsCAT_TSTT_EQID), ",")
'                    For lIdx2 = LBound(vRemoveEQ) To UBound(vRemoveEQ)
'                        sEqID = RemoveStrInArray(sEqID, vRemoveEQ(lIdx2))
'                    Next
''                    sEqID = RemoveStrInArray(sEqID, colRaws.Item(lIdx).Item(gsCAT_TSTT_EQID))
'                Else '處理"可測"機台,如果清單沒有，則加入清單內
'                    If GetStrPosInArray(sEqID, colRaws.Item(lIdx).Item(gsCAT_TSTT_EQID)) = 0 Then
'                        sEqID = sEqID & "," & colRaws.Item(lIdx).Item(gsCAT_TSTT_EQID)
'                    End If
'                End If
'            Next
'        End If
        
        'Add by Tony Start on 2014/06/30 for 產品停復測專案(判斷可測機台條件重寫,必須為機台交集)
        Dim vaTicNoAndEQ(100) As Variant
        Dim sEQList  As String
        Dim lCount   As Long
        Dim sFirstOneListNo As String
        
        If sSpecifyEQ = "Y" And sEqID <> "" Then '先看第一筆
            
            lCount = 0
            For lIdx = 1 To colRaws.Count '將所有的單號與機台先記錄起來(Ex:vaTicNoAndEQ(0):"ST201001008;Y;TM01,TM03,TM07")
                
                If lIdx = 1 Then
                    sFirstOneListNo = colRaws.Item(lIdx).Item(gsCAT_TSTT_STOPTICNO)
                End If
                    
                If sFirstOneListNo = colRaws.Item(lIdx).Item(gsCAT_TSTT_STOPTICNO) Then '還是同一張單
                    sEQList = sEQList & "," & colRaws.Item(lIdx).Item(gsCAT_TSTT_EQID)
                    
                    If VBA.Left(sEQList, 1) = "," Then
                        sEQList = VBA.Mid(sEQList, 2, Len(sEQList))
                    End If
                    
                Else '非同一張單
                    
                    lCount = lCount + 1
                    sFirstOneListNo = colRaws.Item(lIdx).Item(gsCAT_TSTT_STOPTICNO)
                    sEQList = colRaws.Item(lIdx).Item(gsCAT_TSTT_EQID)
                    
                End If
                
                vaTicNoAndEQ(lCount) = colRaws.Item(lIdx).Item(gsCAT_TSTT_STOPTICNO) & ";" & _
                                           colRaws.Item(lIdx).Item(gsCAT_TSTT_SPECIFYEQ) & ";" & _
                                           sEQList
                                           
            Next
            
            Dim asEQListDetail As Variant '(將Array內容拆解至內)
            Dim asEQIDDetail As Variant '(將EQID內容拆解至內)
            Dim sEQIDList_1 As String
            Dim sEQIDList_2 As String
            Dim lIdx1   As Long
            
            For lIdx = LBound(vaTicNoAndEQ) To lCount
                
                If lIdx = 0 Then '(第一筆一定是SpecifyEQ = 'Y')
                    asEQListDetail = Split(vaTicNoAndEQ(lIdx), ";")
                    sFirstOneListNo = asEQListDetail(0)
                    sEQIDList_2 = asEQListDetail(2)
                Else '第二張以後
                    asEQListDetail = Split(vaTicNoAndEQ(lIdx), ";")
                    asEQIDDetail = Split(asEQListDetail(2), ",")
                    
                    If asEQListDetail(1) = "Y" Then
                        
                        If sFirstOneListNo <> asEQListDetail(0) Then
                            If sEQIDList_2 <> "" Then
                                sEQIDList_1 = sEQIDList_2
                                sEQIDList_2 = ""
                            End If
                        End If
                        
                        For lIdx1 = LBound(asEQIDDetail) To UBound(asEQIDDetail)
                            If GetStrPosInArray(sEQIDList_1, asEQIDDetail(lIdx1)) >= 1 Then
                                sEQIDList_2 = sEQIDList_2 & "," & asEQIDDetail(lIdx1)
                            End If
                        Next lIdx1
                
                    Else 'SpecifyEQ = 'N'
                        bIssue = True
                        
                        For lIdx1 = LBound(asEQIDDetail) To UBound(asEQIDDetail)
                            sEQIDList_2 = RemoveStrInArray(sEQIDList_2, asEQIDDetail(lIdx1))
                        Next
                        
                    
                    End If
                End If
                    
            
            Next lIdx
            
        End If
        
        sEqID = sEQIDList_2
        
        'Add by Tony End on 2014/06/30 for 產品停復測專案
        
        If sSpecifyEQ = "N" Then '因已排序，假設第一筆為N，表示只有設定不可測機台的申請
            For lIdx = 1 To colRaws.Count
                If GetStrPosInArray(sEqID, colRaws.Item(lIdx).Item(gsCAT_TSTT_EQID)) = 0 Then
                    sEqID = sEqID & "," & colRaws.Item(lIdx).Item(gsCAT_TSTT_EQID)
                End If
            Next
        End If
        
        '記錄所有停測單號，重覆的不計入
        For lIdx = 1 To colRaws.Count
            If GetStrPosInArray(sStopTicNo, colRaws.Item(lIdx).Item(gsCAT_TSTT_STOPTICNO)) = 0 Then
                sStopTicNo = sStopTicNo & "," & colRaws.Item(lIdx).Item(gsCAT_TSTT_STOPTICNO)
            End If
        Next
        
        If sEqID <> "" Then
            If Left(sEqID, 1) = "," Then
                sEqID = Mid(sEqID, 2)
            End If
            sEqID = SortString(sEqID)
            
            vEqID = Split(sEqID, ",")
            For lIdx = LBound(vEqID) To UBound(vEqID)
                If GetStrPosInArray(sTmpEqID, vEqID(lIdx)) = 0 Then
                    sTmpEqID = sTmpEqID & "," & vEqID(lIdx)
                End If
            Next
            If Left(sTmpEqID, 1) = "," Then
                sTmpEqID = Mid(sTmpEqID, 2)
            End If
            If sTmpEqID <> "" Then
                sEqID = sTmpEqID
            End If
        End If
        
        'Added by Jack on 2017/02/06 for BE MES 生產資訊自動化 Phase 1 : 停復測自動化 <Start>
        '若同一停測單, SpecifyEq有Y也有N : 必需要兩者都符合可以移除單號, 才可移除之.
        '若同一停測單, SpecifyEq只有Y或N : 只要符合一者可以移除單號, 就可移除之.
        If sCurEqId <> "" Then
            vTmpTicNoAry = Split(sStopTicNo, ",")
            For lIdx = LBound(vTmpTicNoAry) To UBound(vTmpTicNoAry)
                bRemoveStopTicY = False
                bRemoveStopTicN = False
                
                sKey = vTmpTicNoAry(lIdx) & "/" & "Y"
                If dicTicNoVsEQ.Exists(sKey) Then
                    sOldValue = dicTicNoVsEQ.Item(sKey)
                    If sOldValue = "" Then '指定可測, 但無任何機台.
                        '不做該筆停測單移除.
                    Else '指定可測, 有機台.
                        '要做該筆停測單移除.
                        If GetStrPosInArray(sOldValue, sCurEqId) <> 0 Then
                            bRemoveStopTicY = True
                        End If
                    End If
                End If
                
                sKey = vTmpTicNoAry(lIdx) & "/" & "N"
                If dicTicNoVsEQ.Exists(sKey) Then
                    sOldValue = dicTicNoVsEQ.Item(sKey)
                    If sOldValue = "" Then '指定不可測, 但無任何機台.
                        '不做該筆停測單移除.
                    Else '指定不可測, 有機台.
                        '要做該筆停測單移除.
                        If GetStrPosInArray(sOldValue, sCurEqId) = 0 Then
                            bRemoveStopTicN = True
                        End If
                    End If
                End If
                
                If dicTicNoVsEQ.Exists(vTmpTicNoAry(lIdx) & "/" & "Y") And dicTicNoVsEQ.Exists(vTmpTicNoAry(lIdx) & "/" & "N") Then
                    If bRemoveStopTicY And bRemoveStopTicN Then
                        sStopTicNo = RemoveStrInArray(sStopTicNo, vTmpTicNoAry(lIdx))
                    End If
                Else
                    If bRemoveStopTicY Or bRemoveStopTicN Then
                        sStopTicNo = RemoveStrInArray(sStopTicNo, vTmpTicNoAry(lIdx))
                    End If
                End If
            Next lIdx
        End If
        'Added by Jack on 2017/02/06 for BE MES 生產資訊自動化 Phase 1 : 停復測自動化 <End>
        
        '把停測單後的"逗號"拿掉
        If sStopTicNo <> "" Then
            If Left(sStopTicNo, 1) = "," Then
                sStopTicNo = Mid(sStopTicNo, 2)
            End If
            sStopTicNo = SortString(sStopTicNo)
        End If
        
        'Added by Jack on 2017/02/06 for BE MES 生產資訊自動化 Phase 1 : 停復測自動化 <Start>
        '組合每一筆sStopTicNo的 BePeSet (A,M,A+M)
        'Ex:   ST201610014/A,ST201610014/A,ST201610014/A,ST201610024/M <-- sBePeSet_Total
        '輸出: A,M
        If sStopTicNo <> "" Then
            'Step-1 : 組合所有的停測單號/BePeSet.
            '         Ex:   ST201610014/A,ST201610014/A,ST201610014/A,ST201610024/M <-- sBePeSet_Total
            For lIdx = 1 To colRaws.Count
                sBePeSet_Total = sBePeSet_Total & "," & colRaws.Item(lIdx).Item(gsCAT_TSTT_STOPTICNO) & "/" & colRaws.Item(lIdx).Item(gsCAT_TSTT_BEPE_SET)
            Next
            If Left(sBePeSet_Total, 1) = "," Then
                sBePeSet_Total = Mid(sBePeSet_Total, 2)
            End If
            sBePeSet_Total = SortString(sBePeSet_Total)
                       
            vStopTicNoAry = Split(sStopTicNo, ",") 'sStopTicNo 上面已排序過.
            For lIdx = LBound(vStopTicNoAry) To UBound(vStopTicNoAry)
                bBePeSet_A = False
                bBePeSet_M = False
                sBePeSet_Temp = ""
                
                'Step-2 : 檢核每一個停測單是否有 A 或 M ??
                If GetStrPosInArray(sBePeSet_Total, vStopTicNoAry(lIdx) + "/A") <> 0 Then
                    bBePeSet_A = True
                End If
                If GetStrPosInArray(sBePeSet_Total, vStopTicNoAry(lIdx) + "/M") <> 0 Then
                    bBePeSet_M = True
                End If
                
                If bBePeSet_A And bBePeSet_M Then '同時有 A 及M.
                    sBePeSet_Temp = sBePeSet_Temp & "," & "A+M"
                ElseIf bBePeSet_A Then '只有 A.
                    sBePeSet_Temp = sBePeSet_Temp & "," & "A"
                ElseIf bBePeSet_M Then '只有 M.
                    sBePeSet_Temp = sBePeSet_Temp & "," & "M"
                End If
                
                If Left(sBePeSet_Temp, 1) = "," Then
                    sBePeSet_Temp = Mid(sBePeSet_Temp, 2)
                End If
                If sBePeSet_Temp = "" Then '_代表null(比較好識別)
                    sBePeSet_Temp = "_"
                End If
                sBePeSet = sBePeSet & "," & sBePeSet_Temp
            Next
            
            If Left(sBePeSet, 1) = "," Then
                sBePeSet = Mid(sBePeSet, 2)
            End If
        End If
        'Added by Jack on 2017/02/06 for BE MES 生產資訊自動化 Phase 1 : 停復測自動化 <End>
        
        
        If sSpecifyEQ = "" Then
            bIssue = True
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
        Call HandleError(False, typErrInfo, , moAppLog, True)
        GetProductStopTest = False
    End If
End Function

Public Function GetStrPosInArray(ByVal sArray As String, ByVal sStr As String, Optional ByVal sDelimiter As String) As Integer
    Dim lIndex As Long
    Dim sStrArray() As String
    
    GetStrPosInArray = 0
    If Trim(sArray) = "" Then
        Exit Function
    End If
    
    If Trim(sDelimiter) = "" Then
        sDelimiter = ","
    End If
    sStrArray = Split(sArray, sDelimiter)
    For lIndex = LBound(sStrArray) To UBound(sStrArray)
        If sStrArray(lIndex) <> "" And sStrArray(lIndex) Like sStr Then
            GetStrPosInArray = lIndex + 1
            Exit Function
        End If
    Next

End Function

Public Function GetNonProdStopTest(ByVal sIPN As String, ByVal sEqType2 As String, ByVal sCurEqId As String, _
                                    ByVal sCurAccName As String, ByVal sCurAccNo As String, ByVal sTemperature As String, _
                                    ByRef sStopTicNo As String, ByRef moFwWIP As Object, ByRef moAppLog As Object, ByRef moProRawSql As Object) As Boolean
On Error GoTo ExitHandler:
Dim sProcID         As String
Dim typErrInfo      As tErrInfo
Dim sSQL            As String
Dim colRS    As Collection

Dim sCarrierType As String
Dim sBodySize As String
Dim sPincount As String
Dim sPackageType As String

Dim iTemperature As Integer

Dim lIdx As Long

'Add by  Sam start on 20241224 for #199158
Dim sFormFactorName As String
Dim sModuleOption As String
'Add by  Sam end on 20241224 for #199158

'----
' Init
'----
    sProcID = "GetNonProdStopTest"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog)
    
    'Add by  Sam  on 20241224 for #199158,Add formfactor,moduleopiton
    sSQL = "select a." & gsCAT_TIM_PACKAGE_CODE & ", a." & gsCAT_TIM_BODY_SIZE & " " & _
           " , a." & gsCAT_TIM_PIN_COUNT & ",a." & gsCAT_TIM_CARRIER_TYPE & " " & _
           " , a." & gsCAT_TIM_FORMFACTORNAME & ",a." & gsCAT_TIM_MODULEOPTION & " " & _
           " from " & gsCAT_TBL_IPN_MASTER & " a " & _
           " where a." & gsCAT_TIM_IPN & " ='" & sIPN & "' "
    Set colRS = moProRawSql.QueryDatabase(sSQL)
    If colRS.Count > 0 Then
        sPackageType = colRS.Item(1).Item(gsCAT_TIM_PACKAGE_CODE)
        sBodySize = colRS.Item(1).Item(gsCAT_TIM_BODY_SIZE)
        sPincount = colRS.Item(1).Item(gsCAT_TIM_PIN_COUNT)
        sCarrierType = colRS.Item(1).Item(gsCAT_TIM_CARRIER_TYPE)
'Add by  Sam start on 20241224 for #199158
        sFormFactorName = colRS.Item(1).Item(gsCAT_TIM_FORMFACTORNAME)
        sModuleOption = colRS.Item(1).Item(gsCAT_TIM_MODULEOPTION)
'Add by  Sam end on 20241224 for #199158
    End If
    
    If sTemperature = gsDEFAULT_ROOM_TEMP Then
        iTemperature = giDEFAULT_ROOM_TEMP
    Else
        iTemperature = Val(sTemperature)
    End If
        
    sStopTicNo = ""
    GetNonProdStopTest = True
'----
' Condition Checking
'----
    ' <Put your condition checking codes here>...

'----
' Action
'----
    sSQL = " select distinct a." & gsCAT_TSTN_STOPTICNO & " " & vbNewLine & _
         " from " & gsCAT_TBL_STOP_TEST_NP & " a" & vbNewLine & _
        " where a." & gsCAT_TSTN_DELETEFLAG & " = 'N'" & vbNewLine & _
          " and nvl('" & sEqType2 & "', ' ') = nvl(a." & gsCAT_TSTN_EQTYPE2 & ", nvl('" & sEqType2 & "', ' '))" & vbNewLine & _
          " and nvl('" & sCurEqId & "', ' ') = nvl(a." & gsCAT_TSTN_TESTERID & ", nvl('" & sCurEqId & "', ' '))" & vbNewLine & _
          " and nvl('" & sCurAccName & "', ' ') = nvl(a." & gsCAT_TSTN_ACCNAME & ", nvl('" & sCurAccName & "', ' '))" & vbNewLine & _
          " and nvl('" & sPackageType & "', ' ') = nvl(a." & gsCAT_TSTN_PACKAGETYPE & ", nvl('" & sPackageType & "', ' '))" & vbNewLine & _
          " and nvl('" & sPincount & "', 0) = nvl(a." & gsCAT_TSTN_PINCOUNT & ", nvl('" & sPincount & "', 0))" & vbNewLine & _
          " and nvl('" & sCarrierType & "', ' ') = nvl(a." & gsCAT_TSTN_CARRIERTYPE & ", nvl('" & sCarrierType & "', ' '))" & vbNewLine & _
          " and nvl('" & sBodySize & "', ' ') = nvl(a." & gsCAT_TSTN_BODYSIZE & ", nvl('" & sBodySize & "', ' '))" & vbNewLine & _
          " and nvl('" & iTemperature & "', 0) >= nvl(a." & gsCAT_TSTN_TEMP_MIN & ", nvl('" & iTemperature & "', 0))" & vbNewLine & _
          " and nvl('" & iTemperature & "', 0) <= nvl(a." & gsCAT_TSTN_TEMP_MAX & ", nvl('" & iTemperature & "', 0))"
          
    If sCurAccNo <> "" Then
        sSQL = sSQL & " and ( FUN_ARRAYPOS('" & sCurAccNo & "',a." & gsCAT_TSTN_ACCNO & ") > 0 or a." & gsCAT_TSTN_ACCNO & " is null ) "
    Else
        sSQL = sSQL & " and a." & gsCAT_TSTN_ACCNO & " is null  "
    End If
          
'Add by  Sam start on 20241224 for #199158
        sSQL = sSQL & " and nvl('" & sFormFactorName & "', ' ') = nvl(a." & gsCAT_TSTN_FORM_FACTOR_NAME & ", nvl('" & sFormFactorName & "', ' '))" & vbNewLine
        sSQL = sSQL & " and nvl('" & sModuleOption & "', ' ') = nvl(a." & gsCAT_TSTN_MODULE_OPTION & ", nvl('" & sModuleOption & "', ' '))" & vbNewLine
'Add by  Sam end on 20241224 for #199158
          
    Set colRS = moProRawSql.QueryDatabase(sSQL)
    For lIdx = 1 To colRS.Count
        sStopTicNo = sStopTicNo & "," & colRS.Item(lIdx).Item(gsCAT_TSTN_STOPTICNO)
    Next
    
    If Left(sStopTicNo, 1) = "," Then
        sStopTicNo = Mid(sStopTicNo, 2)
    End If
    
    If sStopTicNo <> "" Then
        GetNonProdStopTest = False
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

End Function

Public Function SortString(ByVal sWaferID As String, Optional ByVal sDelimiter As String) As String
    Dim lIndex As Long
    Dim lIndex2 As Long
    Dim sStrArray() As String
    Dim sTemp As String
    Dim strWafer As String
    
    SortString = ""
    If Trim(sWaferID) = "" Then
        Exit Function
    End If
    
    If Trim(sDelimiter) = "" Then
        sDelimiter = ","
    End If
    
    sStrArray = Split(sWaferID, sDelimiter)
    
    For lIndex = LBound(sStrArray) To UBound(sStrArray)
        For lIndex2 = LBound(sStrArray) To UBound(sStrArray)
            If sStrArray(lIndex) < sStrArray(lIndex2) Then
                sTemp = sStrArray(lIndex2)
                sStrArray(lIndex2) = sStrArray(lIndex)
                sStrArray(lIndex) = sTemp
            End If
        Next
    Next

    strWafer = ""
    
    For lIndex = LBound(sStrArray) To UBound(sStrArray)
        strWafer = strWafer & sDelimiter & sStrArray(lIndex)
    Next

    If Left(strWafer, 1) = sDelimiter Then
        strWafer = Mid(strWafer, 2)
    End If
    
    SortString = strWafer
    
End Function

Public Function RemoveStrInArray(ByVal sArray As String, ByVal sStr As String, _
                                Optional ByVal sDelimiter As String, Optional ByVal sRemoveDelimiter As String) As String
    Dim lIndex As Long
    Dim lCnt As Long
    Dim sStrArray() As String
    Dim sRemoveArray() As String
    
    Dim sNewArray As String
           
    If Trim(sArray) = "" Then
        Exit Function
    End If
    
    If Trim(sStr) = "" Then
        RemoveStrInArray = sArray
        Exit Function
    End If
        
    If Trim(sDelimiter) = "" Then
        sDelimiter = ","
    End If
    
    If Trim(sRemoveDelimiter) = "" Then
        sRemoveDelimiter = ","
    End If
        
    sNewArray = ""
    
    sRemoveArray = Split(sStr, sRemoveDelimiter)
    
    sStrArray = Split(sArray, sDelimiter)
    
     For lCnt = LBound(sRemoveArray) To UBound(sRemoveArray)
    
        For lIndex = LBound(sStrArray) To UBound(sStrArray)
            If sStrArray(lIndex) <> "" And sStrArray(lIndex) <> sRemoveArray(lCnt) Then
                sNewArray = sNewArray & sDelimiter & sStrArray(lIndex)
            End If
        Next
        
        If sNewArray <> "" And lCnt < UBound(sRemoveArray) Then
         sStrArray = Split(sNewArray, sDelimiter)
         sNewArray = ""
        'Add by Weilun on 20190327 for BE#201900133 <start>
        Else '空值就要離開了, 不然跑接下來的迴圈會出問題
            Exit For
        'Add by Weilun on 20190327 for BE#201900133 <end>
        End If

    Next
    
    If sNewArray <> "" Then
        RemoveStrInArray = Mid(sNewArray, 2)
    End If
        
        
End Function

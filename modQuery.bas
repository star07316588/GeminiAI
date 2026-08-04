Public Function GetLotType(sLotID As String, moProRawSql As Object, moAppLog As Object) As String
On Error GoTo ExitHandler:
Dim sProcID             As String
Dim typErrInfo          As tErrInfo
Dim colRS               As Collection
Dim sSQL                As String
Dim sErunTicNO          As String
Dim sSapRwNo            As String
Dim sLotType            As String
'----
' Init
'----
    sProcID = "GetLotType"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog)
    GetLotType = ""
'----
' Condition Checking
'----
    ' <Put your condition checking codes here>...
'----
' Action
'----
    ' <Put your Action codes here>...
    'add by Nelson start on 2005/11/30 for ReqNo:M200505002
    'Get Lot's ERUNTICNO,SAPRWNO then check type of lot.
    sSQL = "SELECT " & _
                    gsCAT_TLI_ERUNTICNO & "," & _
                    gsCAT_TLI_SAPRWNO & " " & _
           "FROM " & _
                    gsCAT_TBL_LOT_INFO & " " & _
           "WHERE " & gsCAT_TLI_LOT_ID & "='" & sLotID & "'"
                    
    Set colRS = moProRawSql.QueryDatabase(sSQL)
    If colRS.Count > 0 Then
        sErunTicNO = colRS.Item(1).Item(1)
        sSapRwNo = colRS.Item(1).Item(2)
    End If
    If sErunTicNO = "" And sSapRwNo = "" Then
        sLotType = "Normal"
    ElseIf sErunTicNO <> "" Then
        sLotType = "Eng"
    ElseIf sSapRwNo <> "" Then
        sLotType = "Rework"
    End If
    GetLotType = sLotType
    'add by Nelson end on 2005/11/30 for ReqNo:M200505002
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

Public Function GetSwapPGName(ByVal moProRawSql As Object, _
                              ByVal moAppLog As Object, _
                              ByVal sLotID As String, _
                              ByVal sProdgroup As String, _
                              ByVal sStepName As String, _
                              ByVal sEqType2 As String, _
                              ByVal sOrigianlPGName As String, _
                              Optional ByVal bSpecQuery As Boolean) As String
On Error GoTo ExitHandler:
Dim sProcID         As String
Dim typErrInfo      As tErrInfo
Dim sSQL            As String
Dim colRS As Collection
Dim colRS2 As Collection

Dim colSeq As Collection
Dim colSeqs As Collection

Dim colPgm As Collection
Dim colPgms As Collection

Dim sSpecStepName As String
Dim sSpecPGName As String


Dim sWaferIds As String
Dim sSteps As String
Dim sPgms As String

Dim iCurStepIndex As Integer
Dim iChkStepIndex As Integer

Dim lIdx As Long
Dim lIdx2 As Long
Dim lRow As Long
    
Dim vSteps As Variant
Dim vPgms As Variant
Dim iCount As Integer

Dim iStepSt As Integer

Dim vSpecStep As Variant

Dim bMatch As Boolean

Dim colTemp As Collection
Dim colMatchCase As Collection

Dim sMatchPGName As String
Dim sMatchSwapFlag As String

Dim sTempSwapFlag As String
Dim sTempPGName As String

Dim sPgmList As String


Dim sProdCode As String
Dim sMaxsite As String
Dim sTestMode As String
Dim sTesterType As String
Dim sPcdType As String

Dim bFound As Boolean
'----
' Init
'----
    sProcID = "GetSwapPGName"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog) '"Entering Function...", moApplog, glLOG_PROC, msMODULE_ID, sProcID)
    GetSwapPGName = ""
    
'----
' Condition Checking
'----
    ' <Put your condition checking codes here>...

'----
' Action
'----

    iCurStepIndex = 0
    sSQL = " select distinct " & gsCAT_PSS_STEP_NAME & "  from (" & _
           "    select distinct t." & gsCAT_PSS_STEP_NO & " ,t." & gsCAT_PSS_STEP_NAME & " " & _
           " from " & gsCAT_TBL_PROD_STEP_SPEC & " t  " & _
           " where t." & gsCAT_PSS_PROD_GROUP & " ='" & sProdgroup & "' and t." & gsCAT_PSS_DOC_STATUS & "='Active' " & _
           " and t." & gsCAT_PSS_STEP_NAME & " like 'SORT%' order by t." & gsCAT_PSS_STEP_NO & " ) "
    Set colRS = moProRawSql.QueryDatabase(sSQL)
    sSpecStepName = ""
    For lIdx = 1 To colRS.Count
        If sStepName = colRS.Item(lIdx).Item(gsCAT_PSS_STEP_NAME) Then
            iCurStepIndex = lIdx
        End If
        sSpecStepName = sSpecStepName & "," & colRS.Item(lIdx).Item(gsCAT_PSS_STEP_NAME)
    Next
    
    If iCurStepIndex <= 0 Then       '未實測此Prodgroup的站點
        GetSwapPGName = ""
        GoTo ExitHandler
    End If
    
    If Left(sSpecStepName, 1) = "," Then
        sSpecStepName = Mid(sSpecStepName, 2)
    End If
    
    vSpecStep = Split(sSpecStepName, ",")
    
    Set colSeqs = New Collection
    sSQL = " select distinct a." & gsCAT_VWLPR_PRODGROUP & ", a." & gsCAT_VWLPR_STEP & ",a." & gsCAT_VWLPR_PGM & " " & _
           " from " & gsCAT_VIEW_WS_LOT_PGM_REC_SEQUENCE & " a " & _
           " where a." & gsCAT_VWLPR_CURLOTID & " = '" & sLotID & "' "
    Set colRS = moProRawSql.QueryDatabase(sSQL)
    For lIdx = 1 To colRS.Count
        If bSpecQuery = True Then
            If lIdx > 1 Then Exit For 'Spec Query只取一筆
        End If
        sSteps = colRS.Item(lIdx).Item(gsCAT_VWLPR_STEP)
        sPgms = colRS.Item(lIdx).Item(gsCAT_VWLPR_PGM)
        sWaferIds = ""
        Set colSeq = New Collection
        
        sSQL = " select distinct a." & gsCAT_VWLPR_WAFERID & " " & _
               " from " & gsCAT_VIEW_WS_LOT_PGM_REC_SEQUENCE & " a " & _
               " where a." & gsCAT_VWLPR_CURLOTID & " = '" & sLotID & "' " & _
               " and a." & gsCAT_VWLPR_STEP & "= '" & sSteps & "' " & _
               " and a." & gsCAT_VWLPR_PGM & "= '" & sPgms & "' "
        Set colRS2 = moProRawSql.QueryDatabase(sSQL)
        For lIdx2 = 1 To colRS2.Count
            sWaferIds = sWaferIds & ";" & colRS2.Item(lIdx2).Item(gsCAT_VWLPR_WAFERID)
        Next
        If Left(sWaferIds, 1) = ";" Then sWaferIds = Mid(sWaferIds, 2)
        
        colSeq.Add sWaferIds, "waferid" 'WaferID
        colSeq.Add sSteps, "step"    'Step
        colSeq.Add sPgms, "pgm" 'Pgm
        colSeqs.Add colSeq, CStr(lIdx)
    Next
    
    If colSeqs.Count = 0 Then  '無Sequence資料
        GetSwapPGName = ""
        GoTo ExitHandler
    End If
    
    
    sSQL = "select a." & gsCAT_TWSP_STEP1 & " as step1 " & _
            " ,a." & gsCAT_TWSP_STEP2 & " as step2 " & _
            " ,a." & gsCAT_TWSP_STEP3 & " as step3 " & _
            " ,a." & gsCAT_TWSP_STEP4 & " as step4 " & _
            " ,a." & gsCAT_TWSP_STEP5 & " as step5 " & _
            " ,a." & gsCAT_TWSP_STEP6 & " as step6 " & _
            " ,a." & gsCAT_TWSP_STEP7 & " as step7 " & _
            " ,a." & gsCAT_TWSP_STEP8 & " as step8 " & _
            " ,a." & gsCAT_TWSP_STEP9 & " as step9 " & _
            " ,a." & gsCAT_TWSP_STEP10 & " as step10 " & _
            " ,a." & gsCAT_TWSP_STEP11 & " as step11 " & _
            " ,a." & gsCAT_TWSP_STEP12 & " as step12 " & _
            " ,a." & gsCAT_TWSP_STEP13 & " as step13 " & _
            " ,a." & gsCAT_TWSP_STEP14 & " as step14 " & _
            " ,a." & gsCAT_TWSP_STEP15 & " as step15 "

    
    sSQL = sSQL & " ,a." & gsCAT_TWSP_PGM1 & " as pgm1 " & _
            " ,a." & gsCAT_TWSP_PGM2 & " as pgm2 " & _
            " ,a." & gsCAT_TWSP_PGM3 & " as pgm3 " & _
            " ,a." & gsCAT_TWSP_PGM4 & " as pgm4 " & _
            " ,a." & gsCAT_TWSP_PGM5 & " as pgm5 " & _
            " ,a." & gsCAT_TWSP_PGM6 & " as pgm6 " & _
            " ,a." & gsCAT_TWSP_PGM7 & " as pgm7 " & _
            " ,a." & gsCAT_TWSP_PGM8 & " as pgm8 " & _
            " ,a." & gsCAT_TWSP_PGM9 & " as pgm9 " & _
            " ,a." & gsCAT_TWSP_PGM10 & " as pgm10 " & _
            " ,a." & gsCAT_TWSP_PGM11 & " as pgm11 " & _
            " ,a." & gsCAT_TWSP_PGM12 & " as pgm12 " & _
            " ,a." & gsCAT_TWSP_PGM13 & " as pgm13 " & _
            " ,a." & gsCAT_TWSP_PGM14 & " as pgm14 " & _
            " ,a." & gsCAT_TWSP_PGM15 & " as pgm15 "
    
    sSQL = sSQL & " ,a." & gsCAT_TWSP_SWAPFLAG1 & " as swapflag1 " & _
            " ,a." & gsCAT_TWSP_SWAPFLAG2 & " as swapflag2 " & _
            " ,a." & gsCAT_TWSP_SWAPFLAG3 & " as swapflag3 " & _
            " ,a." & gsCAT_TWSP_SWAPFLAG4 & " as swapflag4 " & _
            " ,a." & gsCAT_TWSP_SWAPFLAG5 & " as swapflag5 " & _
            " ,a." & gsCAT_TWSP_SWAPFLAG6 & " as swapflag6 " & _
            " ,a." & gsCAT_TWSP_SWAPFLAG7 & " as swapflag7 " & _
            " ,a." & gsCAT_TWSP_SWAPFLAG8 & " as swapflag8 " & _
            " ,a." & gsCAT_TWSP_SWAPFLAG9 & " as swapflag9 " & _
            " ,a." & gsCAT_TWSP_SWAPFLAG10 & " as swapflag10 " & _
            " ,a." & gsCAT_TWSP_SWAPFLAG11 & " as swapflag11 " & _
            " ,a." & gsCAT_TWSP_SWAPFLAG12 & " as swapflag12 " & _
            " ,a." & gsCAT_TWSP_SWAPFLAG13 & " as swapflag13 " & _
            " ,a." & gsCAT_TWSP_SWAPFLAG14 & " as swapflag14 " & _
            " ,a." & gsCAT_TWSP_SWAPFLAG15 & " as swapflag15 "

    sSQL = sSQL & " from " & gsCAT_TBL_WS_PGM_REPLACE & "  a " & _
           " where a." & gsCAT_TWSP_PRODGROUP & "= '" & sProdgroup & "' "
    Set colRS = moProRawSql.QueryDatabase(sSQL)
    
    
    For lRow = 1 To colSeqs.Count
        Set colMatchCase = New Collection
        vSteps = Split(colSeqs.Item(lRow).Item("step"), ",")
        vPgms = Split(colSeqs.Item(lRow).Item("pgm"), ",")
        
        For lIdx = 1 To colRS.Count
            iStepSt = 0
            bMatch = True
            For iCount = LBound(vSteps) To UBound(vSteps)
                iStepSt = iStepSt + 1
                If Trim(vSteps(iCount)) = "" Then Exit For
                If vSteps(iCount) = sStepName Then Exit For
                
                If vSteps(iCount) <> colRS.Item(lIdx).Item("step" & iCount + 1) Or _
                   vPgms(iCount) <> colRS.Item(lIdx).Item("pgm" & iCount + 1) Then
                   bMatch = False
                   Exit For
                End If
            Next
                        
            If bMatch = True And iStepSt > 0 Then '取得相符設定的對應站點,EX:目前LOT為S1+S2 ,設定為S1+S2+S3 ,取得S3的站點及SWAPFLAG
                Set colTemp = New Collection
                colTemp.Add colRS.Item(lIdx).Item("pgm" & iStepSt), "pgm"
                colTemp.Add colRS.Item(lIdx).Item("swapflag" & iStepSt), "swap"
                colMatchCase.Add colTemp, CStr(colMatchCase.Count + 1)
                'Exit For
            End If
        Next
        colSeqs.Item(lRow).Add colMatchCase, "matchcase"
    Next
    
    sTempSwapFlag = ""
    sTempPGName = ""
    Set colPgms = New Collection
    For lRow = 1 To colSeqs.Count
        Set colPgm = New Collection
        If colSeqs.Item(lRow).Item("matchcase").Count > 0 Then
            If colSeqs.Item(lRow).Item("matchcase").Count > 1 Then '符合多組設定時,CHECK程式是否相同
                For lIdx = 1 To colSeqs.Item(lRow).Item("matchcase").Count
                    sTempSwapFlag = colSeqs.Item(lRow).Item("matchcase").Item(lIdx).Item("swap")
                    sTempPGName = colSeqs.Item(lRow).Item("matchcase").Item(lIdx).Item("pgm")
                    If sTempSwapFlag = "Y" Then
                        If CheckSwapPGNameAttr(sEqType2, sTempPGName, sOrigianlPGName, moProRawSql, moAppLog) = True Then
                            colPgm.Add sTempPGName, CStr(colPgm.Count + 1)
                        Else
                            colPgm.Add "", CStr(colPgm.Count + 1)
                        End If
                    End If
                Next
                colPgms.Add colPgm, CStr(colPgms.Count + 1)
            Else
                If colSeqs.Item(lRow).Item("matchcase").Item(1).Item("swap") = "Y" Then
                    If CheckSwapPGNameAttr(sEqType2, colSeqs.Item(lRow).Item("matchcase").Item(1).Item("pgm"), sOrigianlPGName, moProRawSql, moAppLog) = True Then
                        colPgm.Add colSeqs.Item(lRow).Item("matchcase").Item(1).Item("pgm"), CStr(colPgm.Count + 1)
                    Else
                        colPgm.Add "", CStr(colPgm.Count + 1)
                    End If
                End If
                colPgms.Add colPgm, CStr(colPgms.Count + 1)
            End If
        Else
            colPgm.Add "", "1"
            colPgms.Add colPgm, CStr(colPgms.Count + 1)
        End If
        
   Next
    
    sTempPGName = ""
    If colPgms.Count > 0 Then
        For lIdx = 1 To colPgms.Item(1).Count
            sPgmList = sPgmList & "," & colPgms.Item(1).Item(lIdx)
        Next
        If Left(sPgmList, 1) = "," Then sPgmList = Mid(sPgmList, 2)
        
        
        For lIdx = 2 To colPgms.Count
            If sPgmList = "" Then
                bFound = True
            Else
                bFound = False
            End If
            
            For lIdx2 = 1 To colPgms.Item(lIdx).Count
                If sPgmList = "" Then
                    If colPgms.Item(lIdx).Item(lIdx2) = "" Then
                        bFound = True
                        Exit For
                    Else
                        bFound = False
                    End If
                Else
                    If GetStrPosInArray(sPgmList, colPgms.Item(lIdx).Item(lIdx2)) > 0 Then
                        bFound = True
                        Exit For
                    End If
                End If
            Next
            If bFound = False Then
                GetSwapPGName = "X"     '有衝突需分批
                GoTo ExitHandler
            End If
        Next
    End If
    
    If colPgms.Count > 1 And bFound = False Then
        GetSwapPGName = "X"  '有衝突需分批
        GoTo ExitHandler
    End If
    If sPgmList = "" Then
        GetSwapPGName = ""
        GoTo ExitHandler
    End If

    vPgms = Split(sPgmList, ",")
    For lIdx = LBound(vPgms) To UBound(vPgms)
        If vPgms(lIdx) <> "" Then
            GetSwapPGName = vPgms(lIdx)
            Exit For
        End If
    Next
    


'    'Check 程式屬性
'    sSQL = "Select a." & gsCAT_TTDP_TDSPROD & " ,a." & gsCAT_TTDP_MAXSITE & " " & _
'          " ,a." & gsCAT_TTDP_TESTMODE & " ,a." & gsCAT_TTDP_TDSTESTERTYPE & " " & _
'          " ,a." & gsCAT_TTDP_PCDTYPE & " ,a." & gsCAT_TTDP_PGNAME & "  " & _
'          " from " & gscat_TBL_TDS_PGM & " a " & _
'          " where a." & gsCAT_TTDP_PGNAME & " = '" & sOrigianlPGName & "' " & _
'          " and a." & gsCAT_TTDP_TDSTESTERTYPE & " = '" & sEqType2 & "'  " & _
'          " and a." & gsCAT_TTDP_PKGCODE & " in ('W','H' ) " & _
'          " and a." & gsCAT_TTDP_PINCOUNT & " =0 "
'    Set colRS = moProRawSql.QueryDatabase(sSQL)
'    If colRS.Count > 0 Then
'        sProdCode = colRS.Item(1).Item(gsCAT_TTDP_TDSPROD)
'        sMaxsite = colRS.Item(1).Item(gsCAT_TTDP_MAXSITE)
'        sTestMode = colRS.Item(1).Item(gsCAT_TTDP_TESTMODE)
'        sTesterType = colRS.Item(1).Item(gsCAT_TTDP_TDSTESTERTYPE)
'        sPcdType = colRS.Item(1).Item(gsCAT_TTDP_PCDTYPE)
'    End If
'
'    vPgms = Split(sPgmList, ",")
'    For lIdx = LBound(vPgms) To UBound(vPgms)
'            sSQL = "Select a." & gsCAT_TTDP_TDSPROD & " ,a." & gsCAT_TTDP_MAXSITE & ", " & _
'                  " a." & gsCAT_TTDP_TESTMODE & " ,a." & gsCAT_TTDP_TDSTESTERTYPE & ", " & _
'                  " a." & gsCAT_TTDP_PCDTYPE & " ,a." & gsCAT_TTDP_PGNAME & "  " & _
'                  " from " & gscat_TBL_TDS_PGM & " a " & _
'                  " where a." & gsCAT_TTDP_PGNAME & " = '" & vPgms(lIdx) & "' " & _
'                  " and a." & gsCAT_TTDP_TDSTESTERTYPE & " = '" & sEqType2 & "'  " & _
'                  " and a." & gsCAT_TTDP_PKGCODE & " in ('W','H' ) " & _
'                  " and a." & gsCAT_TTDP_PINCOUNT & " =0 "
'
'            Set colRS = moProRawSql.QueryDatabase(sSQL)
'            If colRS.Count > 0 Then
'                If sProdCode = colRS.Item(1).Item(gsCAT_TTDP_TDSPROD) And _
'                   sMaxsite = colRS.Item(1).Item(gsCAT_TTDP_MAXSITE) And _
'                   sTestMode = colRS.Item(1).Item(gsCAT_TTDP_TESTMODE) And _
'                   sTesterType = colRS.Item(1).Item(gsCAT_TTDP_TDSTESTERTYPE) Then
'
'                    If sEqType2 Like "NEXTEST*" Then
'                        If sPcdType = colRS.Item(1).Item(gsCAT_TTDP_PCDTYPE) Then
'                            GetSwapPGName = vPgms(lIdx)
'                        End If
'                    Else
'                        GetSwapPGName = vPgms(lIdx)
'                    End If
'                End If
'            End If
'    Next
    
    
    
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
            Case glERR_INVALIDOBJECT, glERR_FAILTOUPDATE
                typErrInfo.sUserText = typErrInfo.sErrDescription
            Case Else
                typErrInfo.sUserText = "Fail to execute application, please call IT support!!" & vbCrLf & _
                                        "程式執行失敗, 請洽IT人員處理"
            End Select
        '---- Start of Select Case Block ----
        On Error GoTo ExitHandler:
        GetSwapPGName = False
        Call HandleError(False, typErrInfo, , moAppLog, True)
    End If
End Function

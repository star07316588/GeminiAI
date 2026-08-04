Public Function GetFTAccByPgm(ByVal moProRawSql As Object, ByVal moAppLog As Object, _
                              ByVal sStepName As String, ByVal sBodySize As String, _
                              ByVal sPgId As String, ByVal sPgName As String) As Collection
On Error GoTo ExitHandler:
Dim sProcID As String
Dim typErrInfo As tErrInfo

Dim sSQL As String
Dim colRS As Collection

Dim sProdCode As String
Dim sPkgCode As String
Dim sPinCount As String
Dim sEqType2 As String
Dim sWireVerssion As String
Dim sVendor As String
Dim sVendorNo As String
Dim sVendorList As String
Dim sTdsPkgCode As String

Dim sLoadBoard As String
Dim sContactBoard As String
Dim sBurnInBoard As String

'----
' Init
'----
    sProcID = "GetFTAccByPgm"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog)
    
    Set GetFTAccByPgm = New Collection
           
'----
' Condition Checking
'----
    ' <Put your condition checking codes here>...

'----
' Action
'----
    'Mark,統一參考DB FUNCTION FUN_GET_FT_ACC_BYPGM
'    If Left(sPGName, 1) = "Q" Then
'        sPGName = "F" & Mid(sPGName, 2)
'    ElseIf Left(sPGName, 2) = "TQ" Then
'        sPGName = "TF" & Mid(sPGName, 3)
'    End If
'    'PGM
'    sSQL = "select a." & gsCAT_TTDP_TDSPROD & "  " & _
'           ",a." & gsCAT_TTDP_PKGCODE & "  " & _
'           ",a." & gsCAT_TTDP_PINCOUNT & "  " & _
'           ",a." & gsCAT_TTDP_TDSTESTERTYPE & "  " & _
'           ",a." & gsCAT_TTDP_WIRE_VERSION & "  " & _
'           ",a." & gsCAT_TTDP_VENDORLIST & "  " & _
'           " from " & gsCAT_TBL_TDS_PGM & " a " & _
'           " where a." & gsCAT_TTDP_PGID & " = '" & sPgId & "' " & _
'           " and a." & gsCAT_TTDP_PGNAME & " = '" & sPGName & "' "
'    Set colRS = moProRawSql.QueryDatabase(sSQL)
'    If colRS.Count > 0 Then
'        sProdCode = colRS.Item(1).Item(gsCAT_TTDP_TDSPROD)
'        sPkgCode = colRS.Item(1).Item(gsCAT_TTDP_PKGCODE)
'        sPincount = colRS.Item(1).Item(gsCAT_TTDP_PINCOUNT)
'        sEqType2 = colRS.Item(1).Item(gsCAT_TTDP_TDSTESTERTYPE)
'        sWireVerssion = colRS.Item(1).Item(gsCAT_TTDP_WIRE_VERSION)
'    End If
'
'    'Vendor 固定MXIC
'    sSQL = "select a." & gsCAT_TVM_VENDOR_NO & " from " & gsCAT_TBL_VENDOR_MASTER & "  a " & _
'           " where a." & gsCAT_TVM_VENDOR_NAME & " = 'MXIC' "
'    Set colRS = moProRawSql.QueryDatabase(sSQL)
'    If colRS.Count > 0 Then
'        sVendorNo = colRS.Item(1).Item(gsCAT_TVM_VENDOR_NO)
'    End If
'
'    sSQL = "select " & gsCAT_TPT_TDSPKGDESC & " from " & gsCAT_TBL_PKGCODE_TYPE & " WHERE packagecode = '" & sPkgCode & "' "
'    Set colRS = moProRawSql.QueryDatabase(sSQL)
'    If colRS.Count > 0 Then
'        sTdsPkgCode = colRS.Item(1).Item(gsCAT_TPT_TDSPKGDESC)
'    End If
'
'    'Accessary
'    sSQL = "select a." & gsCAT_TAR_LOADBOARD & " " & _
'          " ,a." & gsCAT_TAR_CONTACTBOARD & " " & _
'          " ,a." & gsCAT_TAR_BURNINBOARD & " " & _
'          " from " & gsCAT_TBL_ACC_RELATION & " a " & _
'          " WHERE  a." & gsCAT_TAR_TESTER_TYPE & " = '" & sEqType2 & "' " & _
'          " and a." & gsCAT_TAR_TDS_PRODCODE & " = '" & sProdCode & "' " & _
'          " and a." & gsCAT_TAR_TDS_PKGCODE & " = '" & sTdsPkgCode & "' " & _
'          " and a." & gsCAT_TAR_PIN_COUNT & " = '" & sPincount & "' " & _
'          " and a." & gsCAT_TAR_WIRE_VERSION & " = '" & sWireVerssion & "' " & _
'          " and a." & gsCAT_TAR_BODYSIZE & " = '" & sBodysize & "' " & _
'          " and a." & gsCAT_TAR_VENDOR & " = '" & sVendorNo & "' " & _
'          " and a." & gsCAT_TAR_DELETE_FLAG & " = 'N' "
'
    sSQL = "select fun_get_ft_acc_bypgm('" & sStepName & "', '" & sBodySize & "', '" & sPgId & "', '" & sPgName & "' , 'loadboard') as loadboard " & _
            "     ,fun_get_ft_acc_bypgm('" & sStepName & "', '" & sBodySize & "', '" & sPgId & "', '" & sPgName & "' , 'contactboard') as contactboard " & _
            "     ,fun_get_ft_acc_bypgm('" & sStepName & "', '" & sBodySize & "', '" & sPgId & "', '" & sPgName & "' , 'burninboard') as burninboard " & _
            " from dual "
    Set colRS = moProRawSql.QueryDatabase(sSQL)
    If colRS.Count > 0 Then
        sLoadBoard = colRS.Item(1).Item("loadboard")
        sContactBoard = colRS.Item(1).Item("contactboard")
        sBurnInBoard = colRS.Item(1).Item("burninboard")
    End If
        
    GetFTAccByPgm.Add sLoadBoard, "loadboard"
    GetFTAccByPgm.Add sContactBoard, "contactboard"
    GetFTAccByPgm.Add sBurnInBoard, "burninboard"
    
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

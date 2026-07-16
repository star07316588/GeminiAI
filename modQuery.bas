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

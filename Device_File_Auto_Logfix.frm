Private Sub CwMailbox1_MessageReceived(message As String, replybox As String)
    On Error GoTo ExitHandler
    Dim typErrInfo As tErrInfo
    Dim sProcID As String
    Dim sStepName As String
    Dim sFormNo As String
    Dim sFormType As String
    Dim sEqID As String
    
    Dim sSql As String
    Dim colRS As Collection
    Dim sPath As String
    Dim sEqAbbr As String
    Dim sSubSysType As String
    Dim sProberVendor As String
    Dim sDeviceFile As String
    
    Dim iCount As Integer
    Dim oFs As Object
    Dim bFileExists As Boolean
    Dim oFile As Object
    Dim sResult As String
    
    Dim sMbxName As String
    Dim sType As String
    Dim sSystemTime As String
    Dim bDelFail As Boolean
    Dim lPid As Long
    
    '-----
    ' Init
    '-----
    sProcID = "CwMailbox1_MessageReceived"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog)
    
    sMbxName = FwWIP1.MbxName
    If sMbxName = "fw_mbx_fwadmin" Then
        sType = "PROD"
    Else
        sType = "ENG"
    End If
    
    sSql = "select to_char(sysdate,'yyyymmdd hh24miss') from dual "
    Set colRS = ProRawSQL.QueryDatabase(sSql)
    If colRS.Count > 0 Then
        sSystemTime = colRS.Item(1).Item(1)
    End If
    
    Dim sGroupHistory As String
    sGroupHistory = "DeviceFile_Auto-" & GetTxnSeq(moRawSQL, moAppLog)
    
    '--------------------
    ' Condition Checking
    '--------------------
    gbMail = False
    DoEvents
    
    sPath = "C:\device_filter\"
    On Error Resume Next
    Err = 0
    ChDir sPath
    If Err Then
        Err = 0
        MkDir sPath
        ' [ADD LOG] 記錄自動建立目錄
        Call LogTrace("Directory not found, created: " & sPath & " (Err=" & Err.Number & ")", moAppLog)
    End If
    Err = 0
    On Error GoTo ExitHandler
    
    Call LogTrace("------------------------------------------- ", moAppLog)
    Call LogTrace("ReceivedMessage: " & message & " Time:" & sSystemTime & " Env:" & sType & " ReplyBox:" & replybox, moAppLog)
    
    gsMailBody = "ReceivedMessage: " & message & " Time:" & sSystemTime
    If Trim(message) = "" Then
        ' [ADD LOG] 空訊息記錄
        Call LogError("Received empty message, skip processing.", moAppLog, sProcID)
        GoTo ExitHandler
    End If
    
    '-------
    ' Action
    '-------
    Dim vLotInfo As Variant
    '0:INS  1:FormNO 2:FormType  3:EqID
    vLotInfo = Split(message, "/")
    
    ' [ADD LOG] 檢查訊息格式避免 Subscript out of range 不明原因報錯
    If UBound(vLotInfo) < 3 Then
        Call LogError("Invalid message format (expected >= 4 fields): " & message, moAppLog, sProcID)
        GoTo ExitHandler
    End If
    
    sFormNo = vLotInfo(1)
    sFormType = vLotInfo(2)
    sEqID = vLotInfo(3)
    Call LogTrace("Parsed Msg -> FormNo:" & sFormNo & ", FormType:" & sFormType & ", EqID:" & sEqID, moAppLog)

    iCount = 0
    Do While (iCount < 30)
        sSql = "select a.eqabbr,a.subsysgroup,a.probevendor,a.devicefile_name from tbl_ws_eqform_actions a " & _
               " where a.deleteflag='N' and a.formno='" & sFormNo & "' and a.formtype= '" & sFormType & "' and a.eqid ='" & sEqID & "' "
        Set colRS = ProRawSQL.QueryDatabase(sSql)
        If colRS.Count > 0 Then
            sEqAbbr = colRS.Item(1).Item("eqabbr")
            sSubSysType = colRS.Item(1).Item("subsysgroup")
            sProberVendor = colRS.Item(1).Item("probevendor")
            sDeviceFile = colRS.Item(1).Item("devicefile_name")
            ' [ADD LOG] 記錄第幾次查到資料與欄位內容
            Call LogTrace("DB record found at retry #" & iCount & " -> EqAbbr:" & sEqAbbr & ", SubSys:" & sSubSysType & ", Vendor:" & sProberVendor & ", DeviceFile:" & sDeviceFile, moAppLog)
            Exit Do
        End If
        Sleep (1000)
        iCount = iCount + 1
    Loop
    
    ' [ADD LOG] 記錄 30 秒查無資料的逾時警告
    If iCount >= 30 Then
        Call LogError("DB query timeout (30s)! No record found in tbl_ws_eqform_actions. SQL: " & sSql, moAppLog, sProcID)
    End If
    
    ' [ADD LOG] 檢查舊檔刪除結果
    bDelFail = DeleteFile(sPath & sEqID & ".txt")
    If bDelFail = True Then
        Call LogError("Failed to delete old file: " & sPath & sEqID & ".txt (May read stale result!)", moAppLog, sProcID)
    End If
    
    Call LogTrace("Execute: " & sPath & "device_filter.exe " & sEqID & " " & sEqAbbr & " " & sSubSysType & " " & sProberVendor & " " & sDeviceFile & " " & sType, moAppLog)
    lPid = Shell(sPath & "device_filter.exe " & sEqID & " " & sEqAbbr & " " & sSubSysType & " " & sProberVendor & " " & sDeviceFile & " " & sType)
    ' [ADD LOG] 記錄外部程式 PID
    Call LogTrace("device_filter.exe started, PID=" & lPid, moAppLog)

    iCount = 0
    Set oFs = CreateObject("Scripting.FileSystemObject")
    Do Until oFs.FileExists(sPath & sEqID & ".txt") = True
        Sleep (1000)
        iCount = iCount + 1
        If iCount > 120 Then
            Exit Do
        End If
    Loop
    
    bFileExists = oFs.FileExists(sPath & sEqID & ".txt")
    sResult = ""
    If bFileExists = True Then
        ' [ADD LOG] 記錄等待耗時
        Call LogTrace("Result file detected after " & iCount & "s: " & sPath & sEqID & ".txt", moAppLog)
        Set oFile = oFs.OpenTextFile(sPath & sEqID & ".txt")
        Sleep (1000)
        If oFile.AtEndOfStream <> True Then
            sResult = oFile.ReadLine()
        End If
        oFile.Close ' 建議補上 Close，避免檔案被鎖定導致下次 DeleteFile 失敗
        
        ' [ADD LOG] 記錄讀到的結果
        Call LogTrace("Result file content: [" & sResult & "]", moAppLog)
        
        If sResult <> "" Then
            sSql = "update tbl_ws_eqform_actions set actionresult= '" & sResult & "',updatetime = to_char(sysdate,'yyyymmdd hh24miss') || '000'  " & _
                   " ,updateuserid ='SYS' where formno ='" & sFormNo & "' and formtype ='" & sFormType & "' and eqid ='" & sEqID & "' "
            Call LogTrace("Updating DB -> SQL: " & sSql, moAppLog)
            Call ProRawSQL.QueryDatabase(sSql)
            Call LogTrace("DB update completed.", moAppLog)
        Else
            ' [ADD LOG] 檔案存在但內容為空
            Call LogError("Result file is empty, skip DB update: " & sPath & sEqID & ".txt", moAppLog, sProcID)
        End If
    Else
        ' [ADD LOG] 等待 120 秒逾時未產生檔案
        Call LogError("Wait for result file timeout (>120s)! File not found: " & sPath & sEqID & ".txt", moAppLog, sProcID)
    End If
            
    '-------
    'Done
    '-------
ExitHandler:
    Call GetErrInfo(typErrInfo, Erl)
    Call LogProcOut(msMODULE_ID, sProcID, moAppLog)
    Screen.MousePointer = VBRUN.MousePointerConstants.vbDefault

ErrorHandler:
    If typErrInfo.lErrNumber Then
        Call LogTrace("ReceivedMessage: " & message & " Process Fail!", moAppLog)
        Call LogTrace("ErrNum: " & typErrInfo.lErrNumber & " Desc: " & typErrInfo.sErrDescription, moAppLog)
        Call HandlerError(typErrInfo, msMODULE_ID, sProcID, moAppLog, False)
        gbMail = True
    Else
        Call LogTrace("ReceivedMessage: " & message & " Process Complete.", moAppLog)
    End If
    
    If gbMail = True Then
        Call LogTrace("Sending error email to: " & Me.txtEMail(0), moAppLog)
        Call EMail(moSMTP, moAppLog, Me.txtSmtp, "Admin", Me.txtEMail(0), Me.txtCC(0), msMODULE_ID & " Error", gsMailBody)
    End If
    Call Restart("CwMailbox1_MessageReceived End")
End Sub

Private Sub CheckConnection_Timer()
    On Error GoTo ExitHandler:
    Dim sProcID As String
    Dim sSql As String
    Dim colRS As Collection
    Dim typErrInfo As tErrInfo

    ' [ADD LOG] 原本未賦值，導致 LogProcOut 函式名稱為空
    sProcID = "CheckConnection_Timer"
    
    'Check connection status
    sSql = "SELECT * FROM TBL_MBX_INFO WHERE MBX_NO <= 'MBX021'"
    Set colRS = ProRawSQL.QueryDatabase(sSql)

ExitHandler:
    Call GetErrInfo(typErrInfo, Erl)
    ' 註：每 30 秒觸發一次，若正常時不想洗版，可只留異常時記 Log
ErrorHandler:
    If typErrInfo.lErrNumber Then
        ' [ADD LOG] 記錄觸發重連的真正錯誤原因
        Call LogError("CheckConnection failed! ErrNum: " & typErrInfo.lErrNumber & " Desc: " & typErrInfo.sErrDescription, moAppLog, sProcID)
        Call HandlerError(typErrInfo, msMODULE_ID, sProcID, moAppLog, False)
        
        On Error GoTo ExitHandler:
        Call Restart("CheckConnection_Timer Failed")
    End If
End Sub

Private Sub Restart(Optional ByVal sReason As String = "Unknown")
    Dim pid As String
    
    ' [ADD LOG] 記錄重啟原因與新啟動的 PID
    Call LogTrace(DateTime.Now & " AutoReconnection Start. Reason: " & sReason, moAppLog)
    pid = Shell(App.Path & "\DeviceFile_Auto.exe")
    Call LogTrace("New instance launched. PID: " & pid, moAppLog)
    
    AutoRestartTimer.Enabled = False
    CheckConnection.Enabled = False
    Unload Me
End Sub

Private Function DeleteFile(ByVal sFileName As String) As Boolean
    On Error Resume Next
    Dim fs As New FileSystemObject
    DeleteFile = False
    
    If fs.FileExists(sFileName) = True Then
        FileSystem.Kill sFileName
        If Err.Number <> 0 Then
            ' [ADD LOG] 記錄刪檔失敗的系統錯誤訊息
            Call LogError("DeleteFile Kill failed (" & sFileName & "): " & Err.Number & " - " & Err.Description, moAppLog, "DeleteFile")
            DeleteFile = True
        Else
            Call LogTrace("Deleted old file: " & sFileName, moAppLog)
        End If
    End If
End Function

' [ADD LOG] 連線前
    Call LogTrace("Connecting -> Host:" & Trim(Me.txtHostName.Text) & ", MBX:" & Trim(Me.txtMBXName.Text) & ", PCMBX:" & Trim(Me.txtPCMBXSrvName.Text) & ", User:" & Me.txtUserName, moAppLog)

    ' ...原本的連線與初始化程式碼...

    ' [ADD LOG] 連線成功開始監聽
    Call LogTrace("Connected successfully. Listening on Mailbox: " & Me.CwMailbox1.MailboxName, moAppLog)

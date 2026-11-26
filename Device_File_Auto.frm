Option Explicit
'Object
Private moAppLog    As CITLogCtrl.clsLogTraceMsg
Private moRawSQL    As PrjRawSql.clsRawSql

'Default Connection Setting
Private Const msDefaultHostName As String = "172.18.16.101"
Private Const msDefaultMBXName  As String = "fw_mbx_mxicdev"
Private Const msDefaultPCMBX    As String = "PCMBXsrv.2"

Private Const msMODULE_ID       As String = "DeviceFile_Auto"
Private msUserName              As String
Private moSMTP                  As Object
Private miMbBoxIndex            As Integer

Private intIdx  As Integer

Private msSender As String
Private msRecipient As String
Private msSubject As String
Private msMessage As String

'Add by Samson on 20241129 AutoReconnetion
Private miSecond As Integer


'Private miErrMail As Integer

'================================================================================
'Name       : cmdClose_Click()
'Description: Close Form
'Author     : Vencent Wei, CIT 2003/10/23
'Parameters : N/A
'================================================================================
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdConfig_Click()
    Me.fraConnect.Visible = False
    Me.fraConfig.Visible = True
    
    txtEMail(0).Visible = True
    txtCC(0).Visible = True
    
    txtEMail(1).Visible = False
    txtCC(1).Visible = False
    txtEMail(2).Visible = False
    txtCC(2).Visible = False
    txtEMail(3).Visible = False
    txtCC(3).Visible = False

End Sub

'================================================================================
'Name       : cmdConnect_Click()
'Description: Connect to FW
'Author     : Vencent Wei, CIT 2003/10/23
'Parameters : N/A
'================================================================================
Private Sub cmdConnect_Click()
    On Error GoTo ExitHandler:
    Dim sProcID As String
    Dim typErrInfo As tErrInfo
    '-----
    ' Init
    '-----
    sProcID = "cmdConnect_Click"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog) '
    '--------------------
    ' Condition Checking
    '--------------------
    If Trim(Me.txtHostName.Text) = "" Or _
                Trim(Me.txtMBXName.Text) = "" Or _
                Trim(Me.txtPCMBXSrvName.Text) = "" Then
        MsgBox "Please Input Connection Setting", , "Msg"
        GoTo ExitHandler
    End If
    
    '-------
    ' Action
    '-------
    Screen.MousePointer = VBRUN.MousePointerConstants.vbHourglass

    With Me.CwMbx1
        .HostName = Trim(Me.txtHostName.Text)
        .MbxName = Trim(Me.txtMBXName.Text)
        .ServerName = Trim(Me.txtPCMBXSrvName.Text)
        .Connect
        Call .GetError(typErrInfo.lErrNumber, typErrInfo.sErrDescription)
        If typErrInfo.lErrNumber Then
            Err.Raise typErrInfo.lErrNumber, , typErrInfo.sErrDescription
        End If
    End With
    
    Me.FwOPR1.Initialize Trim(Me.txtHostName.Text), Trim(Me.txtMBXName.Text), Trim(Me.txtPCMBXSrvName.Text)
    Me.FwOPR1.Login Me.txtUserName, Me.txtPassword
    Me.FwWF1.Initialize Trim(Me.txtHostName.Text), Trim(Me.txtMBXName.Text), Trim(Me.txtPCMBXSrvName.Text)
    Call ProRawSQL.Initialize(Me.CwMbx1, _
            Me.FwWF1.ClientID & ".ProRawSql_Reply", _
            "C:\Rawsql.ini", _
            300)
    
    Set moRawSQL = New PrjRawSql.clsRawSql
    Set moRawSQL.CwMBX = Me.CwMbx1
    Set moRawSQL.AppLog = moAppLog
    moRawSQL.RawSqlMailBox = msMODULE_ID & Format(Now, "DDHHMMSS")
    

    msUserName = "Auto"
    
    Me.fraConnect.Visible = False
    Me.fraTrace.Visible = True
    Me.Refresh
    Me.CwMailbox1.MailboxName = "MBX_DEVICEFILE"
    Me.CwMailbox1.ReceiveMessages = Continuously
       
    
    Call ApIniWriteItem(App.Path & "\DeviceFile_Auto.ini", "ConnectSetting", "HostName", Me.txtHostName, moAppLog)
    Call ApIniWriteItem(App.Path & "\DeviceFile_Auto.ini", "ConnectSetting", "MBXName", Me.txtMBXName, moAppLog)
    Call ApIniWriteItem(App.Path & "\DeviceFile_Auto.ini", "ConnectSetting", "PCMBXSrv", Me.txtPCMBXSrvName, moAppLog)
        
    '-------
    'Done
    '-------
ExitHandler:
    Call GetErrInfo(typErrInfo, Erl)
    Call LogProcOut(msMODULE_ID, sProcID, moAppLog)
    Screen.MousePointer = VBRUN.MousePointerConstants.vbDefault
    
ErrorHandler:
    If typErrInfo.lErrNumber Then
        Call HandlerError(typErrInfo, msMODULE_ID, sProcID, moAppLog, False)
        Call Restart
    End If
End Sub

Private Sub cmdOK_Click()
Dim sValue  As String
    Call ApIniWriteItem(App.Path & "\DeviceFile_Auto.ini", "Config", "LogLevel", VBA.Left(Me.lblLogLevel, 1), moAppLog)
    Call ApIniWriteItem(App.Path & "\DeviceFile_Auto.ini", "Config", "EMAIL", Me.txtEMail(0), moAppLog)
    Call ApIniWriteItem(App.Path & "\DeviceFile_Auto.ini", "Config", "EMAIL2", Me.txtEMail(1), moAppLog)
    Call ApIniWriteItem(App.Path & "\DeviceFile_Auto.ini", "Config", "EMAIL3", Me.txtEMail(2), moAppLog)

    Call ApIniWriteItem(App.Path & "\DeviceFile_Auto.ini", "Config", "EMAIL4", Me.txtEMail(3), moAppLog)

    Call ApIniWriteItem(App.Path & "\DeviceFile_Auto.ini", "Config", "CC", Me.txtCC(0), moAppLog)
    Call ApIniWriteItem(App.Path & "\DeviceFile_Auto.ini", "Config", "CC2", Me.txtCC(1), moAppLog)
    Call ApIniWriteItem(App.Path & "\DeviceFile_Auto.ini", "Config", "CC3", Me.txtCC(2), moAppLog)
    
    Call ApIniWriteItem(App.Path & "\DeviceFile_Auto.ini", "Config", "CC4", Me.txtCC(3), moAppLog)

    Call ApIniWriteItem(App.Path & "\DeviceFile_Auto.ini", "Config", "SMTP", Me.txtSmtp, moAppLog)
    
    'Add by Samson on 20241129 AutoReconnetion
    Call ApIniWriteItem(App.Path & "\DeviceFile_Auto.ini", "Config", "AutoReconnectionAccount", Me.AutoReconnectionAccount, moAppLog)
    Call ApIniWriteItem(App.Path & "\DeviceFile_Auto.ini", "Config", "AutoReconnectionPassword", Me.AutoReconnetionPassword, moAppLog)

    If Me.AutoReconnetionCheck.Value = 1 Then
        Call ApIniWriteItem(App.Path & "\DeviceFile_Auto.ini", "Config", "AutoReconnection", "1", moAppLog)
        Me.txtUserName = Me.AutoReconnectionAccount
        Me.txtPassword = Me.AutoReconnetionPassword
        miSecond = 30
        Me.RestartAfterLable.Caption = "Restart After " & miSecond & " s."
        AutoRestartTimer.Interval = 1000
        AutoRestartTimer.Enabled = True
        CheckConnection.Interval = 30000
        CheckConnection.Enabled = True
    Else
        Call ApIniWriteItem(App.Path & "\DeviceFile_Auto.ini", "Config", "AutoReconnection", "0", moAppLog)
        Me.RestartAfterLable.Caption = ""
        AutoRestartTimer.Enabled = False
        CheckConnection.Enabled = False
    End If
    
    Me.fraConfig.Visible = False
    Me.fraConnect.Visible = True
End Sub

Private Sub CwMailbox1_ErrorOccurred(ErrorCode As Long, ErrorText As String)
    LogError "Error: " & ErrorText, moAppLog, "CwMailbox1_ErrorOccurred"
End Sub

'================================================================================
'Name       : CwMailbox1_MessageReceived()
'Author     : Vencent Wei, CIT 2004/11/22
'================================================================================
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
    
    '-----
    ' Init
    '-----
    sProcID = "CwMailbox1_MessageReceived"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog) '
    
    sMbxName = FwWIP1.MbxName
    If sMbxName = "fw_mbx_fwadmin" Then
        sType = "PROD"
    Else
        sType = "ENG"
    End If
    
    sSql = "select  to_char(sysdate,'yyyymmdd hh24miss')  from dual "
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
        ' We must create this directory
        Err = 0
        MkDir sPath
    End If
    Err = 0
    On Error GoTo ExitHandler
    
    Call LogTrace("------------------------------------------- ", moAppLog)
    Call LogTrace("ReceivedMessage: " & message & " Time:" & sSystemTime, moAppLog)
    
    gsMailBody = "ReceivedMessage: " & message & " Time:" & sSystemTime
    If Trim(message) = "" Then
        'LogError
        GoTo ExitHandler
    End If
    '-------
    ' Action
    '-------
    Dim vLotInfo    As Variant
    
    '0:INS  1:FormNO 2:FormType  3:EqID
    vLotInfo = Split(message, "/")
    
    sFormNo = vLotInfo(1)
    sFormType = vLotInfo(2)
    sEqID = vLotInfo(3)
        
    'Modify by Sam start on 20170329 for DeviceFile_AUTO 修改,改為sql取得這四個欄位,以迴圈包起來避免執行過快資料庫資料尚未進去
    'Modify start by Robin at 2017/02/08 for 上線調整, Direce retrieve from Mailbox
'    '4:EqAbbr 5:SubSysType 6:ProberVendor 7:DeviceFile
'    sEqAbbr = vLotInfo(4)
'    sSubSysType = vLotInfo(5)
'    sProberVendor = vLotInfo(6)
'    sDeviceFile = vLotInfo(7)
'    'Modify end by Robin at 2017/02/08 for 上線調整, Direce retrieve from Mailbox

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
            Exit Do
        End If
        Sleep (1000)
        iCount = iCount + 1
    Loop
    'Modify end by Robin at 2017/02/08 for 上線調整, Direce retrieve from Mailbox
    'Modify by Sam End on 20170329 for DeviceFile_AUTO 修改,改為sql取得這四個欄位,以迴圈包起來避免執行過快資料庫資料尚未進去
    
    Call DeleteFile(sPath & sEqID & ".txt")
    
    Call LogTrace("execute " & sPath & "device_filter.exe" & " " & sEqID & " " & sEqAbbr & " " & sSubSysType & " " & sProberVendor & " " & sDeviceFile & " " & sType, moAppLog)
    Call Shell(sPath & "device_filter.exe" & " " & sEqID & " " & sEqAbbr & " " & sSubSysType & " " & sProberVendor & " " & sDeviceFile & " " & sType)

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
        Set oFile = oFs.OpenTextFile(sPath & sEqID & ".txt")
        Sleep (1000)
        If oFile.AtEndOfStream <> True Then
            sResult = oFile.ReadLine()
        End If
        If sResult <> "" Then
            sSql = "update tbl_ws_eqform_actions set actionresult= '" & sResult & "',updatetime = to_char(sysdate,'yyyymmdd hh24miss') || '000'  " & _
                   " ,updateuserid ='SYS' where formno ='" & sFormNo & "' and formtype ='" & sFormType & "' and eqid ='" & sEqID & "' "
            Call ProRawSQL.QueryDatabase(sSql)
        End If
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
        Call LogTrace(typErrInfo.sErrDescription, moAppLog)
        Call HandlerError(typErrInfo, msMODULE_ID, sProcID, moAppLog, False)
        gbMail = True
    Else
        Call LogTrace("ReceivedMessage: " & message & " Process Complete.", moAppLog)
    End If
    
    If gbMail = True Then
        Call EMail(moSMTP, moAppLog, Me.txtSmtp, "Admin", Me.txtEMail(0), Me.txtCC(0), msMODULE_ID & " Error", gsMailBody)
    End If
    Call Restart
End Sub

'================================================================================
'Name       : Form_Load()
'Description: Init when Main Form Load
'Author     : Vencent Wei, CIT 2003/10/23
'Parameters : N/A
'================================================================================
Private Sub Form_Load()
    On Error GoTo ExitHandler:
    
    Dim sHostName   As String
    Dim sMbxName    As String
    Dim sPCMBX      As String
    Dim sValue      As String
    Dim sAutoReconnetionn   As String
    '-----
    ' Init
    '-----
    Me.fraConnect.Visible = True
    Me.fraTrace.Visible = False
    Me.cmdClear.Visible = False
    '-------
    ' Action
    '-------
    Set rtfLog = Me.rtfTrace
    
    Call ApIniGetItem(App.Path & "\DeviceFile_Auto.ini", "ConnectSetting", "HostName", sHostName)
    If sHostName = "" Then
        Me.txtHostName.Text = msDefaultHostName
        Me.txtMBXName.Text = msDefaultMBXName
        Me.txtPCMBXSrvName.Text = msDefaultPCMBX
    Else
        Call ApIniGetItem(App.Path & "\DeviceFile_Auto.ini", "ConnectSetting", "MBXName", sMbxName)
        Call ApIniGetItem(App.Path & "\DeviceFile_Auto.ini", "ConnectSetting", "PCMBXSrv", sPCMBX)
        Me.txtHostName.Text = sHostName
        Me.txtMBXName.Text = sMbxName
        Me.txtPCMBXSrvName.Text = sPCMBX
    End If
    
    
    Call ApIniGetItem(App.Path & "\DeviceFile_Auto.ini", "Config", "LogLevel", sValue)
    If sValue = "" Then
        Me.Slider1.Value = 9
    ElseIf Val(sValue) = 0 Then
        Me.Slider1.Value = 1
    ElseIf Val(sValue) = 1 Then
        Me.Slider1.Value = 2
    ElseIf Val(sValue) = 5 Then
        Me.Slider1.Value = 3
    Else
        Me.Slider1.Value = 4
    End If

    Call ApIniGetItem(App.Path & "\DeviceFile_Auto.ini", "Config", "EMAIL", sValue)
    Me.txtEMail(0) = sValue
    Call ApIniGetItem(App.Path & "\DeviceFile_Auto.ini", "Config", "EMAIL2", sValue)
    Me.txtEMail(1) = sValue
    Call ApIniGetItem(App.Path & "\DeviceFile_Auto.ini", "Config", "EMAIL3", sValue)
    Me.txtEMail(2) = sValue
    
    Call ApIniGetItem(App.Path & "\DeviceFile_Auto.ini", "Config", "EMAIL4", sValue)
    Me.txtEMail(3) = sValue
    
    Call ApIniGetItem(App.Path & "\DeviceFile_Auto.ini", "Config", "EMAIL5", sValue)
    Me.txtEMail(4) = sValue
    
    Call ApIniGetItem(App.Path & "\DeviceFile_Auto.ini", "Config", "CC", sValue)
    Me.txtCC(0) = sValue
    Call ApIniGetItem(App.Path & "\DeviceFile_Auto.ini", "Config", "CC2", sValue)
    Me.txtCC(1) = sValue
    Call ApIniGetItem(App.Path & "\DeviceFile_Auto.ini", "Config", "CC3", sValue)
    Me.txtCC(2) = sValue
    
    Call ApIniGetItem(App.Path & "\DeviceFile_Auto.ini", "Config", "CC4", sValue)
    Me.txtCC(3) = sValue
    
    Call ApIniGetItem(App.Path & "\DeviceFile_Auto.ini", "Config", "CC5", sValue)
    Me.txtCC(4) = sValue
    
    Call ApIniGetItem(App.Path & "\DeviceFile_Auto.ini", "Config", "SMTP", sValue)
    Me.txtSmtp = sValue
    
    Call ApIniGetItem(App.Path & "\DeviceFile_Auto.ini", "Config", "AutoReconnectionAccount", sValue)
    Me.AutoReconnectionAccount = sValue
    
    Call ApIniGetItem(App.Path & "\DeviceFile_Auto.ini", "Config", "AutoReconnectionPassword", sValue)
    Me.AutoReconnetionPassword = sValue
    On Error Resume Next
    
    Call ApIniGetItem(App.Path & "\DeviceFile_Auto.ini", "Config", "AutoReconnection", sAutoReconnetionn)
    If sAutoReconnetionn = "1" Then
        Me.AutoReconnetionCheck.Value = 1
        Call ApIniGetItem(App.Path & "\DeviceFile_Auto.ini", "Config", "AutoReconnectionAccount", sValue)
        Me.txtUserName = sValue
        Call ApIniGetItem(App.Path & "\DeviceFile_Auto.ini", "Config", "AutoReconnectionPassword", sValue)
        Me.txtPassword = sValue
        miSecond = 30
        Me.RestartAfterLable.Caption = "Restart After " & miSecond & " s."
        AutoRestartTimer.Interval = 1000
        AutoRestartTimer.Enabled = True
        CheckConnection.Interval = 30000
        CheckConnection.Enabled = True
    End If
    
    '需將jmail.dll放置到DeviceFile_Auto.exe執行檔之執行目錄下
    'Set moSMTP = CreateObject("OSSMTP.SMTPSession")
    Set moSMTP = CreateObject("jMail.Message")
    If Err.Number <> 0 Then
        Err.Clear
        'mark by Nelson on 2007/11/09 for 更新Email Object為jMail
        'Call Shell("c:\FASTech\FwClient\Etc\RegSvr32.exe /s " & App.Path & "\OSSMTP.dll")
        Call Shell("c:\FASTech\FwClient\Etc\RegSvr32.exe /s " & App.Path & "\jmail.dll")
        Set moSMTP = CreateObject("jMail.Message")
        If moSMTP Is Nothing Then
            UtShowMsgBox "Cat not Create Mail Object"
            GoTo ExitHandler
        End If
    End If
    On Error GoTo ExitHandler
    
    Call InitLogCtrl(moAppLog, LS_LOGLEVEL_9)
ExitHandler:
    If Err.Number Then
        Err.Clear
    End If
End Sub
'================================================================================
'Name       : Restart()
'Description: Close Form and Restart
'Author     : Samson Chiang, MES 2024/11/12
'Parameters : N/A
'================================================================================
Private Sub Restart()
    
    Dim pid      As String
    
    'ADD by Samson on 20241112 For AutoReconnection
    Call LogTrace(DateTime.Now & " AutoReconnection Start", moAppLog)
    
    pid = Shell(App.Path & "\DeviceFile_Auto.exe")
    AutoRestartTimer.Enabled = False
    CheckConnection.Enabled = False
    Unload Me
End Sub

'================================================================================
'Name       : Form_Unload()
'Description: Release Object when Main Form unLoad
'Author     : Vencent Wei, CIT 2003/10/24
'Parameters : N/A
'================================================================================
Private Sub Form_Unload(Cancel As Integer)
    Set moAppLog = Nothing
    Set moRawSQL = Nothing
    If Me.CwMbx1.Status = Connected Then Call Me.CwMbx1.Disconnect
End Sub

Private Sub optBin5_Click()
    Me.txtEMail(4).ZOrder 0
    Me.txtCC(4).ZOrder 0
    
    txtEMail(0).Visible = False
    txtCC(0).Visible = False
    txtEMail(1).Visible = False
    txtCC(1).Visible = False
    txtEMail(2).Visible = False
    txtCC(2).Visible = False
    
    txtEMail(3).Visible = False
    txtCC(3).Visible = False
    
    txtEMail(4).Visible = True
    txtCC(4).Visible = True

End Sub

Private Sub optBORROW_Click()
    Me.txtEMail(3).ZOrder 0
    Me.txtCC(3).ZOrder 0
    txtEMail(0).Visible = False
    txtCC(0).Visible = False
    txtEMail(1).Visible = False
    txtCC(1).Visible = False
    txtEMail(2).Visible = False
    txtCC(2).Visible = False
    
    txtEMail(3).Visible = True
    txtCC(3).Visible = True
        
    txtEMail(4).Visible = False
    txtCC(4).Visible = False
End Sub
Private Sub optFTSUBREC_Click()
    Dim sValue      As String
    Me.txtEMail(0).ZOrder 0
    Me.txtCC(0).ZOrder 0
    
    txtEMail(0).Visible = True
    txtCC(0).Visible = True
    
    txtEMail(1).Visible = False
    txtCC(1).Visible = False
    txtEMail(2).Visible = False
    txtCC(2).Visible = False
    txtEMail(3).Visible = False
    txtCC(3).Visible = False
    'Modify by Sam end on 20081216
    
    txtEMail(4).Visible = False
    txtCC(4).Visible = False
End Sub

Private Sub Option1_Click()
    Dim sValue      As String
    Me.txtEMail(2).ZOrder 0
    Me.txtCC(2).ZOrder 0
    
    txtEMail(0).Visible = False
    txtCC(0).Visible = False
    txtEMail(1).Visible = False
    txtCC(1).Visible = False
    
    txtEMail(2).Visible = True
    txtCC(2).Visible = True
    
    txtEMail(3).Visible = False
    txtCC(3).Visible = False
    
    txtEMail(4).Visible = False
    txtCC(4).Visible = False
End Sub

Private Sub optPBBANKSPLIT_Click()
    Dim sValue      As String
    Me.txtEMail(1).ZOrder 0
    Me.txtCC(1).ZOrder 0
    
    txtEMail(0).Visible = False
    txtCC(0).Visible = False
    
    txtEMail(1).Visible = True
    txtCC(1).Visible = True
    
    txtEMail(2).Visible = False
    txtCC(2).Visible = False
    txtEMail(3).Visible = False
    txtCC(3).Visible = False
    
    txtEMail(4).Visible = False
    txtCC(4).Visible = False
End Sub


'================================================================================
'Name       : Slider1_Change()
'Description: Display Log Level value and Desc in Label
'Author     : Vencent Wei, CIT 2003/10/29
'Parameters : N/A
'================================================================================
Private Sub Slider1_Change()
    Select Case Slider1.Value
        Case 1:
            Me.lblLogLevel.Caption = "0 - Log Error Only"
        Case 2:
            Me.lblLogLevel.Caption = "1 - Log Error and Coding Transcation"
        Case 3:
            Me.lblLogLevel.Caption = "5 - Log Error, Coding Transaction and Trace"
        Case 4:
            Me.lblLogLevel.Caption = "9 - Log All Message"
    End Select
    
End Sub

'================================================================================
'Name       : txtHostName_Change()
'Description: Auto change MBX name
'Author     : Vencent Wei, CIT 2003/10/24
'Parameters : N/A
'================================================================================
Private Sub txtHostName_Change()
    If Me.txtHostName = "192.9.101.16" Then Me.txtMBXName = "fw_mbx_dummy"
    If Me.txtHostName = "192.168.211.14" Then Me.txtMBXName = "fw_mbx_fwadmin"
End Sub

Private Function DeleteFile(ByVal sFileName As String) As Boolean
On Error Resume Next
    Dim fs As New FileSystemObject
    DeleteFile = False
    
    '檢查舊檔案是否存在，存在則刪除
'    Sleep (1000)
    If fs.FileExists(sFileName) = True Then
        FileSystem.Kill sFileName
        If Err.Number <> 0 Then
            '刪除失敗
            DeleteFile = True
        End If
    End If
        
End Function

'================================================================================
'Name       : AutoRestartTimer_Timer()
'Description: Auto change MBX name
'Author     : Samson Chiang, 2024/11/29
'Parameters : N/A
'================================================================================
Private Sub AutoRestartTimer_Timer()
        miSecond = miSecond - 1
        Me.RestartAfterLable.Caption = "Restart After " & miSecond & " s."
        If miSecond = 0 Then
            AutoRestartTimer = False
            Call cmdConnect_Click
        End If
End Sub
'================================================================================
'Name       : CheckConnection_Timer()
'Description: Auto change MBX name
'Author     : Samson Chiang, 2024/11/29
'Parameters : N/A
'================================================================================
Private Sub CheckConnection_Timer()
    On Error GoTo ExitHandler:
    Dim sProcID As String
    Dim sSql As String
    Dim colRS As Collection
    Dim typErrInfo As tErrInfo

    'Check connection status
    sSql = "SELECT * FROM TBL_MBX_INFO WHERE MBX_NO <= 'MBX021'"
    Set colRS = ProRawSQL.QueryDatabase(sSql)

ExitHandler:
    ' NOTE 1:
    ' MUST CALL GetErrInfo() here first before another action
       Call GetErrInfo(typErrInfo, Erl)
       Call LogProcOut(msMODULE_ID, sProcID, moAppLog)
    ' <Your cleaning up codes goes here...>
ErrorHandler:
    If typErrInfo.lErrNumber Then
      
        Select Case typErrInfo.lErrNumber
           Case Else
                typErrInfo.sUserText = typErrInfo.sErrDescription
        End Select
        '---- Start of Select Case Block ----
        On Error GoTo ExitHandler:
        Call Restart
    End If
End Sub


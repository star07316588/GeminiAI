Option Explicit

Private Const msMODULE_ID               As String = "CP_ETEST_MonitorCore"
'Object
Private moAppLog                        As CITLogCtrl.clsLogTraceMsg
Private moRawSql                        As PrjRawSql.clsRawSql
Private moProRawSql                     As ProRawSQL
Private moSMTP                          As Object

Private moCwMbx                         As Object
Private moFwCAT                         As Object
Private moFwMDL                         As Object
Private moFwOPR                         As Object
Private moFwPRP                         As Object
Private moFwWF                          As Object
Private moFwWIP                         As Object

'Add by Weilun on 20200819 for Req.VB 嘗試引用 C# DLL <start>
Private moClient                        As Object
'Private moClient                        As MXIC_BE_SignalRClient.clsMain //如果是Reference要用這個
'Add by Weilun on 20200819 for Req.VB 嘗試引用 C# DLL <end>

'Add by Weilun on 20221230 for Req.SorterControl_Monitor <start>
Private moSC_Main                       As Collection
Private moSC_Sub                        As Collection
'Add by Weilun on 20221230 for Req.SorterControl_Monitor <end>


'Default Connection Setting
Private Const msDefaultHostName         As String = "172.18.16.101"
Private Const msDefaultMBXName          As String = "fw_mbx_mxicdev"
Private Const msDefaultPCMBX            As String = "PCMBXsrv.5"

'Default Monitor Connecting Setting
Private Const msDefaultMonitorHost      As String = "http://localhost:1438"
Private Const msDefaultMonitorHub       As String = "MonitorHub"

'Locate
Private Const mlFormWidth               As Long = 8200
Private Const mlFormHeight              As Long = 6700
Private Const mlMajorButtonHeight       As Long = 400
Private Const mlMajorButtonWidth        As Long = 1500
Private Const mlBorderDistance          As Long = 50

'Folder
Private Const msSignalRComFolder        As String = "\SignalR COM" 'Add by Weilun on 20200819 for Req.VB 嘗試引用 C# DLL

'Ini File
Private Const msMainIniPath             As String = "\CP_ETEST_MonitorCore.ini"
Private Const msMonitorIniPath          As String = "\SignalRClient.ini"

'Monitor Exe & Dll File
Private Const msMonitorExe              As String = "\SignalRClient.exe"
Private Const msRegAsm                  As String = "\RegAsm.exe" 'Add by Weilun on 20200819 for Req.VB 嘗試引用 C# DLL
Private Const msMXICSignalR             As String = "\MXIC.BE.SignalRClient.dll" 'Add by Weilun on 20200819 for Req.VB 嘗試引用 C# DLL
Private Const msMonitorDll_1            As String = "\Microsoft.AspNet.SignalR.Client.dll"
Private Const msMonitorDll_2            As String = "\Newtonsoft.Json.dll"

'Other
Private msUserName                      As String
Private msTimeStamp                     As String
Private mlTimer                         As Long
Private mlSCTimer                       As Long    'Add by Weilun on 20221230 for Req.SorterControl_Monitor
Private mbDaily_Maintain                As Boolean
Private mbReadySendSignalR              As Boolean 'Add by Weilun on 20200203 for Req.BE MES Phase 41 - 手動停機修改, 系統背景執行SortLotEnd
Private Const msEnd                     As String = "LotEnd"
Private Const msComplete                As String = "LotComplete"

'------
' Spread spdVirtualMapping
'------
Private Const msSPDVIRTUALMAPPING_SLOT_NO                As String = "SlotNo"
Private Const msSPDVIRTUALMAPPING_LOT_ID                 As String = "LotID"
Private Const msSPDVIRTUALMAPPING_WAFER_ID               As String = "WaferID"

Private Const miSPDVIRTUALMAPPING_SLOT_NO                As Integer = 1
Private Const miSPDVIRTUALMAPPING_LOT_ID                 As Integer = 2
Private Const miSPDVIRTUALMAPPING_WAFER_ID               As Integer = 3

Private Const miSPDVIRTUALMAPPING_MAX_COLS               As Integer = 3
Private bReconnect                      As Boolean 'add by Ernest on 2023/07/28 for  #162501 TA monitor system -Message bus connection was lost 改善
Private unlaodflag                      As Boolean
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Const PROCESS_TERMINATE = &H1
Private Declare Function OpenProcess Lib "kernel32" ( _
    ByVal dwDesiredAccess As Long, _
    ByVal bInheritHandle As Long, _
    ByVal dwProcessId As Long) As Long
    
Private Declare Function TerminateProcess Lib "kernel32" ( _
    ByVal hProcess As Long, _
    ByVal uExitCode As Long) As Long



'================================================================================
'Name       : Form_Load()
'Description: Init when Main Form Load
'Author     : Weilun Huang, MXIC, 2018/02/21
'Parameters : N/A
'================================================================================
Private Sub Form_Load()
    On Error Resume Next
    Call CleanDuplicateExeInstances
    Call InitLogCtrl(moAppLog, LS_LOGLEVEL_5)
    Call Init
End Sub

'================================================================================
'Name       : Form_Unload()
'Description: Release Object when Main Form unLoad
'Author     : Weilun Huang, MXIC, 2018/02/21
'Parameters : N/A
'================================================================================
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Dim sExePath As String
    Dim sReUserName As String
    Dim sRePassword As String
    Dim pid   As Long
    
    Set moAppLog = Nothing
    Set moRawSql = Nothing
    If Not (moCwMbx Is Nothing) Then
        If moCwMbx.Status <> 1 Then 'DisConnected ?
            moCwMbx.Disconnect
        End If
    End If
    
    'ADD by Rubio on 20230907 For Mailbox AutoReconnection
    'Delete by Samson executing by Restart() function 20260304
    'If chkAutoReconnection.Value = vbChecked Then
       
    '    pid = Shell("C:\Windows\SysWOW64\FwClient\Exe\CP_ETEST_MonitorCore.exe")
        
    '    If pid > 0 Then

    '        Call CloseHandle(pid) '關掉背景程式
    '    Else
    '        Call Shell("C:\Windows\SysWOW64\FwClient\Exe\CP_ETEST_MonitorCore.exe")

    '    End If
        
        
    'End If
    
End Sub

'================================================================================
'Name       : Init()
'Description: 初始化轉帳程式
'Author     : Weilun Huang, MXIC, 2018/02/21
'Parameters : N/A
'================================================================================
Public Function Init() As Boolean
    On Error GoTo ExitHandler:
    
    Dim sProcID       As String
    Dim typErrInfo    As tErrInfo
    Dim sHostName     As String
    Dim sMbxName      As String
    Dim sPCMBX        As String
    Dim sValue        As String
    Dim sMonitorHost  As String
    Dim sMonitorHub   As String
    
    Dim sCheckMsg     As String 'Add by Weilun on 20200819 for Req.VB 嘗試引用 C# DLL
    Dim stxtReUserName As String
    Dim stxtRePassword As String
    '-----
    ' Init
    '-----
    sProcID = "Init"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog)
    
    Init = False
    
    Me.Height = mlFormHeight
    Me.Width = mlFormWidth
    Me.fraConnect.Visible = True
    
    Me.Height = mlFormHeight
    Me.Width = mlFormWidth

    Call LocateFrame(Me.fraConnect, True)
    Call LocateFrame(Me.fraConfig, False)
    Call LocateFrame(Me.fraTrace, False)
    '-------
    ' Action
    '-------
    txtEMail(0).Visible = True
    txtCC(0).Visible = True
    
    rtfTrace.SelColor = vbBlack
    rtfTrace.SelBold = False
    rtfTrace.SelFontSize = 8
    
    '對MES
    Call ApIniGetItem(App.Path & msMainIniPath, "ConnectSetting", "HostName", sHostName)
    If sHostName = "" Then
        Me.txtHostName.Text = msDefaultHostName
        Me.txtMBXName.Text = msDefaultMBXName
        Me.txtPCMBXSrvName.Text = msDefaultPCMBX
    Else
        Call ApIniGetItem(App.Path & msMainIniPath, "ConnectSetting", "MBXName", sMbxName)
        Call ApIniGetItem(App.Path & msMainIniPath, "ConnectSetting", "PCMBXSrv", sPCMBX)
        Me.txtHostName.Text = sHostName
        Me.txtMBXName.Text = sMbxName
        Me.txtPCMBXSrvName.Text = sPCMBX
    End If
    
    '對Monitor
    Call ApIniGetItem(App.Path & msMonitorIniPath, "ConnectSetting", "Host", sMonitorHost)
    If sMonitorHost = "" Then
        Me.txtMonitorHost.Text = msDefaultMonitorHost
        Me.txtMonitorHub.Text = msDefaultMonitorHub
    Else
        Call ApIniGetItem(App.Path & msMonitorIniPath, "ConnectSetting", "HubName", sMonitorHub)
        Me.txtMonitorHost.Text = sMonitorHost
        Me.txtMonitorHub.Text = sMonitorHub
    End If
    Call cmdMonitorConfigClose_Click '各Monitor項目設定
    
    Call ApIniGetItem(App.Path & msMainIniPath, "Config", "LogLevel", sValue)
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
    
    Call Slider1_Change
    Call LogLevelChange

    Call ApIniGetEmailItem(App.Path & msMainIniPath, "Config", "EMAIL", sValue, , moAppLog)
    Me.txtEMail(0) = sValue
     
    Call ApIniGetEmailItem(App.Path & msMainIniPath, "Config", "CC", sValue, , moAppLog)
    Me.txtCC(0) = sValue
    
    Call ApIniGetEmailItem(App.Path & msMainIniPath, "Config", "SMTP", sValue, , moAppLog)
    Me.txtSmtp = sValue
    

    'Add by Weilun on 20200819 fro Req.VB 嘗試引用 C# DLL <start>
    'Monitor相關檔案檢查
    sCheckMsg = ""
    If Dir(App.Path & msSignalRComFolder & msRegAsm) = "" Then
        sCheckMsg = sCheckMsg & "缺少" & App.Path & msSignalRComFolder & msRegAsm & vbNewLine
    End If
    
    If Dir(App.Path & msSignalRComFolder & msMXICSignalR) = "" Then
        sCheckMsg = sCheckMsg & "缺少" & App.Path & msSignalRComFolder & msMXICSignalR & vbNewLine
    End If
    
    If Dir(App.Path & msSignalRComFolder & msMonitorDll_1) = "" Then
        sCheckMsg = sCheckMsg & "缺少" & App.Path & msSignalRComFolder & msMonitorDll_1 & vbNewLine
    End If
    
    If Dir(App.Path & msSignalRComFolder & msMonitorDll_2) = "" Then
        sCheckMsg = sCheckMsg & "缺少" & App.Path & msSignalRComFolder & msMonitorDll_2 & vbNewLine
    End If
    
    If sCheckMsg = "" Then
        ShellExecute Me.hWnd, "open", _
                     Chr(34) & App.Path & msSignalRComFolder & msRegAsm & Chr(34), _
                     " /tlb /codebase " & Chr(34) & App.Path & msSignalRComFolder & msMXICSignalR & Chr(34), _
                     App.Path, SW_HIDE
    Else
        MsgBox sCheckMsg
        Me.cmdConnect.Enabled = False
    End If
    'Add by Weilun on 20200819 fro Req.VB 嘗試引用 C# DLL <end>
    
    On Error Resume Next 'Add by Weilun on 20200819
    
    Set moSMTP = CreateObject("jMail.Message")
    If Err.Number <> 0 Then
        Err.Clear
        
        Call Shell("c:\FASTech\FwClient\Etc\RegSvr32.exe /s " & App.Path & "\jmail.dll")
        Set moSMTP = CreateObject("jMail.Message")
        If moSMTP Is Nothing Then
            UtShowMsgBox "Cat not Create Mail Object"
            GoTo ExitHandler
        End If
    End If
    

    Init = True
    
    'ADD by Rubio on 20230907 For Mailbox AutoReconnection
    If Init = True And chkAutoReconnection.Value = vbChecked Then
        Call ApIniGetItem(App.Path & msMainIniPath, "MonitorConfig", "UserName", stxtReUserName)
        Call ApIniGetItem(App.Path & msMainIniPath, "MonitorConfig", "Password", stxtRePassword)
        Me.txtUserName = stxtReUserName
        Me.txtPassword = stxtRePassword
        TimerAutoReconnection.Interval = 30000
        TimerAutoReconnection.Enabled = True
    End If
    
    '-------
    'Done
    '-------
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
           Case Else
                typErrInfo.sUserText = typErrInfo.sErrDescription
        End Select
        '---- Start of Select Case Block ----
        On Error GoTo ExitHandler:
        'Call HandleError(True, typErrInfo, , moAppLog, , True)
        Call TransferLogError(typErrInfo.sUserText, _
                      moAppLog, Me.rtfTrace, msMODULE_ID, sProcID)
    End If
End Function

'================================================================================
'Name       : InitFwCtrls()
'Description: 初始化MES元件
'Author     : Weilun huang, MXIC 2018/03/19
'Parameters : N/A
'================================================================================
Private Function InitFwCtrls(ByRef sHostName As String, ByRef sMbxName As String, ByRef sSrvName As String) As Boolean
    On Error GoTo ExitHandler:
    Dim sProcID As String
    Dim typErrInfo As tErrInfo
    Dim lStartTick As Long
    Dim mlPrevPointer As Long

'----
' Init
'----
    sProcID = "InitFwCtrls"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog) '
    InitFwCtrls = False
'----
' Condition Checking
'----
    ' <Put your condition checking codes here>...
    If Len(sHostName) = 0 Then
        Call RaiseError(glERR_INVALIDVALUE, _
                FormatErrorText(gsETX_INVALIDVALUE, "sHostName::String", vbNullString))
    End If
    
    If Len(sSrvName) = 0 Then
        Call RaiseError(glERR_INVALIDVALUE, _
                FormatErrorText(gsETX_INVALIDVALUE, "sSrvName::String", vbNullString))
    End If
    
    If Len(sMbxName) = 0 Then
        Call RaiseError(glERR_INVALIDVALUE, _
                FormatErrorText(gsETX_INVALIDVALUE, "sMbxName::String", vbNullString))
    End If
        
'----
' Action
'----
    If Not (Me.CwMbxMain Is Nothing) Then
        If Me.CwMbxMain.Status <> 1 Then 'DisConnected ?
            Me.CwMbxMain.Disconnect
        End If
    End If

    'Initialize the CwMbx ctrl
    With Me.CwMbxMain
        Call LogTransBegin(msMODULE_ID, sProcID, "MBX.Init", moAppLog, lStartTick, "MbxName:=" & sMbxName, "HostName:=" & sHostName, "PCMBXsrvName:=" & sSrvName)
        .HostName = sHostName
        .MbxName = sMbxName
        .ServerName = sSrvName
        .Connect
        Call .GetError(typErrInfo.lErrNumber, typErrInfo.sErrDescription)
        Call LogTransEnd(msMODULE_ID, sProcID, "MBX.Init", moAppLog, lStartTick, "MbxName:=" & sMbxName, "HostName:=" & sHostName, "PCMBXsrvName:=" & sSrvName)
    End With
    
    ' Initialize the FwWF ctrl
    'FwWF初始化時, 可將同一個Form上FwOPR, FwPRP, FwMDL, FwWIP一併初始化
    With Me.FwWFMain
        Call LogTransBegin(msMODULE_ID, sProcID, "WF.Init", moAppLog, lStartTick, "MbxName:=" & sMbxName, "HostName:=" & sHostName, "PCMBXsrvName:=" & sSrvName)
        'modify by Ernest on 2023/07/28 for  #162501 TA monitor system -Message bus connection was lost 改善
        If bReconnect = False Then
            Call .Initialize(sHostName, sMbxName, sSrvName)
        End If
        Call LogTransEnd(msMODULE_ID, sProcID, "WF.Init", moAppLog, lStartTick, "MbxName:=" & sMbxName, "HostName:=" & sHostName, "PCMBXsrvName:=" & sSrvName)
    End With

    Set moCwMbx = Me.CwMbxMain
    Set moFwOPR = Me.FwOPRMain
    Set moFwPRP = Me.FwPRPMain
    Set moFwWIP = Me.FwWIPMain
    Set moFwWF = Me.FwWFMain
    Set moFwMDL = Me.FwMDLMain
 
    InitFwCtrls = True
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
            Case glERR_INVALIDOBJECT, glERR_FAILTOINSERT
                typErrInfo.sUserText = typErrInfo.sErrDescription
            Case Else
                typErrInfo.sUserText = "Error occurred while trying to establish MBX connection!" & vbCrLf & _
                "Error (" & typErrInfo.lErrNumber & ") " & typErrInfo.sErrDescription
        End Select
        '---- Start of Select Case Block ----
        On Error GoTo ExitHandler:
        'Call HandleError(False, typErrInfo, , moAppLog, , True)
        Call TransferLogError(typErrInfo.sUserText, _
                              moAppLog, Me.rtfTrace, msMODULE_ID, sProcID)
    End If
End Function

'================================================================================
'Name       : LocateMajorButton()
'Description: 隨視窗大小設定按鈕位置
'Author     : Weilun huang, MXIC 2018/03/19
'Parameters : N/A
'================================================================================
Private Sub LocateMajorButton()
    On Error GoTo ExitHandler:
    Dim sProcID As String
    Dim typErrInfo As tErrInfo
    '-----
    ' Init
    '-----
    sProcID = "LocateMajorButton"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog)

    '--------------------
    ' Condition Checking
    '--------------------
    ' <Put your condition checking codes here>...

    '-------
    ' Action
    '-------
    '定位主要按鈕
    Dim lBottonAreaHeight   As Long
    Dim lBottonAreaWidth    As Long
    
    lBottonAreaHeight = (mlBorderDistance * 2) + mlMajorButtonHeight
    lBottonAreaWidth = (mlBorderDistance * 2) + mlMajorButtonWidth
    
    Me.cmdClose.Height = mlMajorButtonHeight
    Me.cmdClose.Width = mlMajorButtonWidth
    Me.cmdClose.Top = Me.ScaleHeight - lBottonAreaHeight
    Me.cmdClose.Left = Me.ScaleWidth - lBottonAreaWidth
    
    Me.cmdClear.Height = mlMajorButtonHeight
    Me.cmdClear.Width = mlMajorButtonWidth
    Me.cmdClear.Top = Me.ScaleHeight - lBottonAreaHeight
    Me.cmdClear.Left = Me.ScaleWidth - (lBottonAreaWidth * 2)
    
    Me.cmdMonitorConfig.Height = mlMajorButtonHeight
    Me.cmdMonitorConfig.Width = mlMajorButtonWidth
    Me.cmdMonitorConfig.Top = Me.ScaleHeight - lBottonAreaHeight
    Me.cmdMonitorConfig.Left = Me.ScaleWidth - (lBottonAreaWidth * 3)
    '-------
    'Done
    '-------
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
           Case Else
                typErrInfo.sUserText = typErrInfo.sErrDescription
        End Select
        '---- Start of Select Case Block ----
        On Error GoTo ExitHandler:
        Call HandleError(False, typErrInfo, , moAppLog, , True)
    End If
End Sub

'================================================================================
'Name       : CwMailbox1_ErrorOccurred()
'Description: MailBox Error Log
'Author     : Weilun Huang, MXIC, 2019/09/03
'Parameters : N/A
'================================================================================
Private Sub CwMailbox1_ErrorOccurred(ErrorCode As Long, ErrorText As String)
    TransferLogError "Error: " & ErrorText, moAppLog, rtfTrace, "CwMailbox1_ErrorOccurred"
End Sub

'================================================================================
'Name       : CwMailbox1_MessageReceived()
'Description: MailBox1 Message Received
'Author     : Weilun Huang, MXIC, 2019/09/03
'Parameters : N/A
'================================================================================
Private Sub CwMailbox1_MessageReceived(message As String, replybox As String)
    On Error GoTo ExitHandler:
    
    Dim sProcID             As String
    Dim typErrInfo          As tErrInfo
    
    Dim sLotId              As String
    '-----
    ' Init
    '-----
    sProcID = "CwMailbox1_MessageReceived"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog)
    Dim asMsg()             As String
    Dim sMonitorType        As String
    
    sLotId = ""

    '----
    ' Condition Checking
    '----
    ' <Put your condition checking codes here>...
    
    '-------
    ' Action
    '-------
    Call TransferLogTrace("------------------------------------------- ", _
                          moAppLog, Me.rtfTrace)
    Call TransferLogTrace(DateTime.Now & ": CP_ETEST_MonitorCore ReceivedMessage: " & message, _
                          moAppLog, Me.rtfTrace)
                         
    
    asMsg = Split(message, "/")
    If UBound(asMsg) > -1 Then 'Type
    
        Select Case asMsg(0)
            Case "EqChange"
                If chkTriggerItem(0).Value = vbChecked Then
                    Call EqStateChnage(message)
                End If
            Case "LotEnd"
                If chkTriggerItem(1).Value = vbChecked Then
                    Call CPAutoChange(message)
                End If
            Case "WaferEnd"
                If chkTriggerItem(2).Value = vbChecked Then
                    Call WaferEndNotice(message)
                End If
            Case "TestAbort"
                If chkTriggerItem(3).Value = vbChecked Then
                    Call TestAbortAlarm(message)
                End If
            Case "PEISUpdate"
                If chkTriggerItem(4).Value = vbChecked And TimerPEISUpdate.Enabled = False Then
                    TimerPEISUpdate.Enabled = True '延遲後, 整體搜索並檢查
                End If
            Case "ReviewUpdate"
                If chkTriggerItem(5).Value = vbChecked Then
                    '更新該EQID的Level1情況, 但是不刷新各User Level1的畫面, 由Review User關掉畫面後再更新(前端呼叫)
                    
                    '為手動停機的Case時(ReviewFlag: M->S), 用SID進行手動停機
                    If asMsg(3) = "S" Then 'ReviewFlag
                        Call StopTestCheckAndExecute(asMsg(2)) 'SID
                    End If
                    
                    If UnloadTypeReview(message) = True Then 'Add by Weilunon 20211130 for Project.BE 工業 3.5 Phase 39 - CP AGV導入專案 - 建立 AGV 派工 & Status 監控系統
                        'Function 內部獨立更新畫面
                    Else
                        '原來的
                        Call TransferLogTrace(DateTime.Now & ": SortLevel1 for EqID: " & asMsg(1), _
                                              moAppLog, Me.rtfTrace)
        
                        Call Level1Sort(asMsg(1), asMsg(2), asMsg(3)) 'EqID/SID/ReviewFlag
                    End If
                End If
            'Add by Weilun on 20200512 for Project.BE MES Phase 44 – CP 產線增加自動化作業 <start>
            Case "UpdateETtestMoveOut"
                If chkTriggerItem(6).Value = vbChecked Then
                    Call UpdateETtestMoveOut(message)
                End If
            Case "AutoDgradeError"
                If chkTriggerItem(7).Value = vbChecked Then
                    Call InsertAutoProgramFail(message)
                End If
            Case "AutoHoldLotReleaseError"
                If chkTriggerItem(8).Value = vbChecked Then
                    Call InsertAutoProgramFail(message)
                End If
            'Add by Weilun on 20200512 for Project.BE MES Phase 44 – CP 產線增加自動化作業 <end>
            'Add by Weilun on 20211203 for Project.BE 工業 3.5 Phase 39 - CP AGV導入專案 - 建立 AGV 派工 & Status 監控系統 <start>
            Case "AGVLotStartError"
                If chkTriggerItem(9).Value = vbChecked Then
                    Call InsertAutoProgramFail(message)
                End If
            Case "AutoPrintSetupFormError"
                If chkTriggerItem(10).Value = vbChecked Then
                    Call InsertAutoProgramFail(message)
                End If
            Case "AutoPrintSetupFormNotice"
                If chkTriggerItem(11).Value = vbChecked Then
                    Call InsertAutoProgramFail(message)
                End If
            'Add by Weilun on 20211203 for Project.BE 工業 3.5 Phase 39 - CP AGV導入專案 - 建立 AGV 派工 & Status 監控系統 <end>
            'Add by Weilun on 20221230 for Req.SorterControl_Monitor <start>
            Case "InsertSorterControl"
                '可能一次刷25筆進來, 先整理
                If chkTriggerItem(12).Value = vbChecked Then
                    Call TransferLogTrace(DateTime.Now & ": Add '" & message & "' to SorterContrl colletcion.", _
                                          moAppLog, Me.rtfTrace)
                    Call ListSC(message)
                    TimerSorterControl.Enabled = True
                End If
            'Add by Weilun on 20221230 for Req.SorterControl_Monitor <end>
            Case Else
                Call TransferLogError(DateTime.Now & ": Unknown Type: " & asMsg(0), moAppLog, Me.rtfTrace)
        End Select
    End If
    
    'Mail Sample
    'Call EMail(moSMTP, moAppLog, Me.txtSmtp, "Admin", Me.txtEMail(0), Me.txtCC(0), msMODULE_ID & " " & " Receive Message. ", DateTime.Now & ": ReceivedMessage: " & message, , Me.rtfTrace)
    '-------
    'Done
    '-------
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
           Case Else
                typErrInfo.sUserText = typErrInfo.sErrDescription
        End Select
        '---- Start of Select Case Block ----
        On Error GoTo ExitHandler:
        Call TransferLogError(typErrInfo.sUserText, _
                              moAppLog, Me.rtfTrace, msMODULE_ID, sProcID)
        Call HandleError(False, typErrInfo, , moAppLog)
    End If
End Sub

'================================================================================
'Name       : CwMailbox2_ErrorOccurred()
'Description: MailBox Error Log
'Author     : Weilun Huang, MXIC, 2019/09/03
'Parameters : N/A
'================================================================================
Private Sub CwMailbox2_ErrorOccurred(ErrorCode As Long, ErrorText As String)
    TransferLogError "Error: " & ErrorText, moAppLog, rtfTrace, "CwMailbox1_ErrorOccurred"
End Sub

'================================================================================
'Name       : CwMailbox2_MessageReceived()
'Description: MailBox1 Message Received
'Author     : Weilun Huang, MXIC, 2019/09/03
'Parameters : N/A
'================================================================================
Private Sub CwMailbox2_MessageReceived(message As String, replybox As String)
    On Error GoTo ExitHandler:
    
    Dim sProcID             As String
    Dim typErrInfo          As tErrInfo
    
    Dim sLotId              As String
    '-----
    ' Init
    '-----
    sProcID = "CwMailbox1_MessageReceived"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog)
    
    sLotId = ""
    
    '----
    ' Condition Checking
    '----
    ' <Put your condition checking codes here>...
    
    '-------
    ' Action
    '-------
    Call TransferLogTrace("------------------------------------------- ", _
                          moAppLog, Me.rtfTrace)
    Call TransferLogTrace(DateTime.Now & ": CP_ETEST_MonitorCore ReceivedMessage: " & message, _
                          moAppLog, Me.rtfTrace)
                         
    'Mail Sample
    'Call EMail(moSMTP, moAppLog, Me.txtSmtp, "Admin", Me.txtEMail(0), Me.txtCC(0), msMODULE_ID & " " & " Receive Message. ", DateTime.Now & ": ReceivedMessage: " & message, , Me.rtfTrace)
                          
    '-------
    'Done
    '-------
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
           Case Else
                typErrInfo.sUserText = typErrInfo.sErrDescription
        End Select
        '---- Start of Select Case Block ----
        On Error GoTo ExitHandler:
        Call TransferLogError(typErrInfo.sUserText, _
                              moAppLog, Me.rtfTrace, msMODULE_ID, sProcID)
        Call HandleError(False, typErrInfo, , moAppLog)
    End If
End Sub

'================================================================================
'Name       : LocateFrame()
'Description: 設定Frame的位置, 如果Frame太大則以Frame為主, 過大的Frame顯示會出問題
'Author     : Weilun huang, MXIC 2018/03/19
'Parameters : N/A
'================================================================================
Private Sub LocateFrame(ByRef fraObject As Frame, ByVal bVisible)
    On Error GoTo ExitHandler:
    Dim sProcID As String
    Dim typErrInfo As tErrInfo
    
    '定位Frame
    Dim lButtonAreaHeight   As Long
    Dim lBroderHeight       As Long
    Dim lBroderWidth        As Long
    '-----
    ' Init
    '-----
    sProcID = "LocateFrame"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog)
    
    '--------------------
    ' Condition Checking
    '--------------------
    ' <Put your condition checking codes here>...

    '-------
    ' Action
    '-------
    
    fraObject.Visible = bVisible        '顯示或不顯示
    If bVisible Then
        
        fraObject.ZOrder vbBringToFront '最上層顯示
        lButtonAreaHeight = (mlBorderDistance * 2) + mlMajorButtonHeight
        lBroderHeight = Me.Height - Me.ScaleHeight
        lBroderWidth = Me.Width - Me.ScaleWidth
        
        '預設視窗大小
        Me.Height = mlFormHeight
        Me.Width = mlFormWidth
        
        '比較Frame與視窗的大小, 如果超過, 以Frame為主
        '不過Frame的比例太誇張的話, form的長寬接受不了, 請注意設計
        If fraObject.Height > Me.ScaleHeight - lButtonAreaHeight - (mlBorderDistance * 2) Then
            Me.Height = fraObject.Height + lBroderHeight + lButtonAreaHeight + (mlBorderDistance * 2)
        End If
        
        If fraObject.Width > Me.ScaleWidth - (mlBorderDistance * 2) Then
            Me.Width = fraObject.Width + lBroderWidth + (mlBorderDistance * 2)
        End If
        
        '置中(排除下方按鈕)
        fraObject.Top = ((Me.ScaleHeight - lButtonAreaHeight - (mlBorderDistance * 2)) / 2) - (fraObject.Height / 2)
        fraObject.Left = (Me.ScaleWidth / 2) - (fraObject.Width / 2)
        
        '重新設置按鈕位置
        Call LocateMajorButton
    End If
    '-------
    'Done
    '-------
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
           Case Else
                typErrInfo.sUserText = typErrInfo.sErrDescription
        End Select
        '---- Start of Select Case Block ----
        On Error GoTo ExitHandler:
        Call HandleError(False, typErrInfo, , moAppLog, , False)
    End If
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
    Dim sErrorConnectFlag As String
    Dim bInitFailRetryFlag As Boolean
    
    Dim sMsg       As String 'Add by Weilun 20200819 for Req.VB 嘗試引用 C# DLL
    '-----
    ' Init
    '-----
    bInitFailRetryFlag = False
    sProcID = "cmdConnect_Click"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog) '
    
    If CreateSignalRClient() = False Then
        GoTo ExitHandler
    End If
    
    '--------------------
    ' Condition Checking
    '--------------------
    If Trim(Me.txtHostName.Text) = "" Or _
                Trim(Me.txtMBXName.Text) = "" Or _
                Trim(Me.txtPCMBXSrvName.Text) = "" Then
        MsgBox "Please Input Connection Setting", , "Msg"
        GoTo ExitHandler
    End If
    
    '安全性卡關
    If Me.txtMBXName.Text = "fw_mbx_fwadmin" Then
        If Me.txtUserName <> "MIT01" Then
            MsgBox "Invalid User!", , "Msg"
            GoTo ExitHandler
        End If
    End If
    
    'Mark by Weilun on 20200819 for Req.VB 嘗試引用 C# DLL
'    'Monitor相關檔案檢查
'    If Dir(App.Path & msMonitorExe) = "" Then
'        MsgBox "缺少" & App.Path & msMonitorExe, , "Msg"
'        GoTo ExitHandler
'    End If
'
'    If Dir(App.Path & msMonitorDll_1) = "" Then
'        MsgBox "缺少" & App.Path & msMonitorDll_1, , "Msg"
'        GoTo ExitHandler
'    End If
'
'    If Dir(App.Path & msMonitorDll_2) = "" Then
'        MsgBox "缺少" & App.Path & msMonitorDll_2, , "Msg"
'        GoTo ExitHandler
'    End If
    
    
    '-------
    ' Action
    '-------
    mbReadySendSignalR = False 'Add by Weilun on 20200203 for Req.BE MES Phase 41 - 手動停機修改, 系統背景執行SortLotEnd

    Screen.MousePointer = VBRUN.MousePointerConstants.vbHourglass
    
    If InitFwCtrls(Trim(Me.txtHostName.Text), Trim(Me.txtMBXName.Text), Trim(Me.txtPCMBXSrvName.Text)) = False Then
        bInitFailRetryFlag = True
        GoTo ExitHandler
    End If
    
    'add by Rubio on 20230907 For MailBox
    If chkAutoReconnection.Value = vbChecked Then
        moFwOPR.Login Me.txtUserName, Me.txtPassword
        Call TransferLogTrace(DateTime.Now & ": MES AutoReconneciton Success." & "RetryUser :" & Me.txtUserName, _
                          moAppLog, Me.rtfTrace)

    Else
        moFwOPR.Login Me.txtUserName, Me.txtPassword '初始登入

    End If
    
    
    Call ProRawSQL.Initialize(Me.CwMbxMain, _
            Me.FwWFMain.ClientID & ".ProRawSql_Reply", _
            "C:\Rawsql.ini", _
            300)
    Set moProRawSql = Me.ProRawSQL
    
    Set moRawSql = New PrjRawSql.clsRawSql
    Set moRawSql.CwMBX = Me.CwMbxMain
    Set moRawSql.AppLog = moAppLog
    moRawSql.RawSqlMailBox = msMODULE_ID & Format(Now, "DDHHMMSS")
    
    msUserName = "MIT01"
    
    Call LocateFrame(Me.fraConnect, False)
    Call LocateFrame(Me.fraTrace, True)
    Me.Refresh
    
    Me.CwMailbox1.MailboxName = "MBX_CP_MONITOR_EVENT"
    Me.CwMailbox1.ReceiveMessages = Continuously
    
    Call ApIniWriteItem(App.Path & msMainIniPath, "ConnectSetting", "HostName", Me.txtHostName, moAppLog)
    Call ApIniWriteItem(App.Path & msMainIniPath, "ConnectSetting", "MBXName", Me.txtMBXName, moAppLog)
    Call ApIniWriteItem(App.Path & msMainIniPath, "ConnectSetting", "PCMBXSrv", Me.txtPCMBXSrvName, moAppLog)
    
    Call ApIniWriteItem(App.Path & msMonitorIniPath, "ConnectSetting", "Host", Me.txtMonitorHost.Text, moAppLog)
    Call ApIniWriteItem(App.Path & msMonitorIniPath, "ConnectSetting", "HubName", Me.txtMonitorHub.Text, moAppLog)
    
    Call cmdMonitorConfigClose_Click '寫Monitor Config相關
    
    'Timer Sample
    mlTimer = 0
    mbDaily_Maintain = True
    TimerMain.Interval = 60000         '1min  最大可設定 65535
    TimerMain.Enabled = True
    
    TimerPEISUpdate.Interval = 2000    '2秒延遲補PEIS資料
    TimerPEISBuffer.Interval = 6000    '6秒後再檢查一次PEIS資料
    
    'Add by Weilun on 20221230 for Req.SorterControl_Monitor<start>
    mlSCTimer = 0
    TimerSorterControl.Interval = 5000     '5秒動一次, 有兩次動作
    'Add by Weilun on 20221230 for Req.SorterControl_Monitor<end>

    'Add by Weilun on 20221230 for Req.SorterControl_Monitor<start>
    Set moSC_Main = New Collection
    Set moSC_Sub = New Collection
    'Add by Weilun on 20221230 for Req.SorterControl_Monitor<end>
        
    Screen.MousePointer = VBRUN.MousePointerConstants.vbDefault
    
    Call TransferLogTrace("MES連線成功.", _
                          moAppLog, Me.rtfTrace) '加上MES以供識別 ,Modify by Weilun on 20200819 fro Req.VB 嘗試引用 C# DLL
             
    'Add by Weilun on 20200819 fro Req.VB 嘗試引用 C# DLL
    If Not moClient Is Nothing Then
        sMsg = moClient.Connect(Me.txtMonitorHost.Text, Me.txtMonitorHub.Text)
        
        If sMsg = moClient.SignalRState_Connected Then
            Call TransferLogTrace("SignalR State: " & sMsg, _
                                  moAppLog, Me.rtfTrace)
        Else
            Call TransferLogError("SignalR State: " & sMsg, _
                                  moAppLog, Me.rtfTrace)
        End If
    End If
    'Add by Weilun on 20200819 fro Req.VB 嘗試引用 C# DLL
            
    'Add by Samson on 2026/03/04 check disconnection and restart
    If chkAutoReconnection.Value = vbChecked Then
        CheckConnection.Interval = 30000 '30 sec
        CheckConnection.Enabled = True
       
    End If
            
    'Monitor初始化
    Call TransferLogTrace("------------------------------------------- ", _
                          moAppLog, Me.rtfTrace)
    Call TransferLogTrace("啟動Monitor初始化與更新.", _
                          moAppLog, Me.rtfTrace)
    Call TransferLogTrace("------------------------------------------- ", _
                          moAppLog, Me.rtfTrace)
    
    Call MonitorMaintain
    Call MonitorInit
    Call TestTimeAbnormalCheck
    Call CPCompleteCheck
    Call NextLotPrepare
    Call CPAutoMonitor
    
    '檢查一下有沒有漏掉的PEIS未補資料, 必須要未被Trigger的情況下
    If TimerPEISUpdate.Enabled = False And TimerPEISBuffer.Enabled = False And chkTriggerItem(4).Value = Checked Then
        Call TransferLogTrace(DateTime.Now & ": Check PEIS data and update.", _
                              moAppLog, Me.rtfTrace)
        Call PEISUpdate
        
        '不檢查自動停機
    End If
    
    '全部整理
    Call Level1Sort
    
    '全部更新
    mbReadySendSignalR = True 'Add by Weilun on 20200203 for Req.BE MES Phase 41 - 手動停機修改, 系統背景執行SortLotEnd
    Call UpdateMonitorAllEq
    
    Call TransferLogTrace("------------------------------------------- ", _
                          moAppLog, Me.rtfTrace)
    Call TransferLogTrace("完成初始化與更新.", _
                          moAppLog, Me.rtfTrace)
    Call TransferLogTrace("------------------------------------------- ", _
                          moAppLog, Me.rtfTrace)
    'modify by Ernest on 2023/07/28 for  #162501 TA monitor system -Message bus connection was lost 改善
     bReconnect = False
    '-------
    'Done
    '-------
ExitHandler:
    ' NOTE 1:
    ' MUST CALL GetErrInfo() here first before another action
    Call GetErrInfo(msMODULE_ID, sProcID, typErrInfo, Erl)
    Call LogProcOut(msMODULE_ID, sProcID, typErrInfo, moAppLog)
    ' <Your cleaning up codes goes here...>
ErrorHandler:
    
    'ADD by Rubio Huang on 20230907 For Mailbox AutoReconnection,新增網路斷線 connect 初始化失敗 一律關掉畫面，讓物件init
    'Modified by Samson on 20260304 call cmdClose_Click -> call Restart
    If bInitFailRetryFlag = True Then
        Call TransferLogError(DateTime.Now & "MES Connection was lost,InitFailRetryFlag = Y ,AutoReconnection Start", _
                                  moAppLog, Me.rtfTrace, msMODULE_ID, sProcID)
        bInitFailRetryFlag = False
        Call Restart
    End If

    If typErrInfo.lErrNumber Then
        ' NOTE 2:
        ' If you have custom handling of some Errors, please
        ' UN-REMARED the following Select Case block!
        ' Also, modify if neccessarily!!!
        '---- Start of Select Case Block ----
        Select Case typErrInfo.lErrNumber
           Case Else
                typErrInfo.sUserText = typErrInfo.sErrDescription
        End Select
        '---- Start of Select Case Block ----
        On Error GoTo ExitHandler:
        'Call HandleError(False, typErrInfo, , moAppLog, , True)
        Call TransferLogError(typErrInfo.sUserText, _
              moAppLog, Me.rtfTrace, msMODULE_ID, sProcID)
        
    End If
End Sub

'================================================================================
'Name       : cmdClose_Click()
'Description: Close Form
'Author     : Vencent Wei, CIT 2003/10/23
'Parameters : N/A
'================================================================================
Private Sub cmdClose_Click()
        
    Unload Me
    
End Sub

'================================================================================
'Name       : cmdClear_Click()
'Description: Clear Msg
'Author     : Weilun Huang, MXIC, 2018/03/27
'Parameters : N/A
'================================================================================
Private Sub cmdClear_Click()
    On Error Resume Next
    Me.rtfTrace.Text = ""
End Sub

'================================================================================
'Name       : cmdConfig_Click()
'Description: Config Email, LogLevel or others
'Author     : Weilun Huang, MXIC, 2018/03/27
'Parameters : N/A
'================================================================================
Private Sub cmdConfig_Click()
    Call LocateFrame(Me.fraConnect, False)
    Call LocateFrame(Me.fraConfig, True)
    
    txtEMail(0).Visible = True
    txtCC(0).Visible = True
    
    'LogLevel
    If Not moAppLog Is Nothing Then
        Select Case moAppLog.LogLevel
            Case LS_LOGLEVEL_0:
                Slider1.Value = 1
            Case LS_LOGLEVEL_1:
                Slider1.Value = 2
            Case LS_LOGLEVEL_5:
                Slider1.Value = 3
            Case LS_LOGLEVEL_9:
                Slider1.Value = 4
        End Select
    End If
      
End Sub

'================================================================================
'Name       : cmdMonitorConfig_Click()
'Description: Config Email, LogLevel or others
'Author     : Weilun Huang, MXIC, 2018/03/27
'Parameters : N/A
'================================================================================
Private Sub cmdMonitorConfig_Click()

    If Me.fraMonitorConfig.Visible = False Then
        Call cmdMonitorConfigClose_Click '先刷新
        Call LocateFrame(Me.fraMonitorConfig, True)
    Else
        Call cmdMonitorConfigClose_Click '再按一次關閉的功能
    End If
End Sub

'================================================================================
'Name       : cmdMonitorConfigSave_Click()
'Description: Save Monitor Config
'Author     : Weilun Huang, MXIC, 2019/10/28
'Parameters : N/A
'================================================================================
Private Sub cmdMonitorConfigSave_Click()
On Error Resume Next

    Call LocateFrame(Me.fraMonitorConfig, False)
    
    Call ApIniWriteItem(App.Path & msMainIniPath, "MonitorConfig", "EqChange", Me.chkTriggerItem(0).Value, moAppLog)
    Call ApIniWriteItem(App.Path & msMainIniPath, "MonitorConfig", "LotEnd", Me.chkTriggerItem(1).Value, moAppLog)
    Call ApIniWriteItem(App.Path & msMainIniPath, "MonitorConfig", "WaferEnd", Me.chkTriggerItem(2).Value, moAppLog)
    Call ApIniWriteItem(App.Path & msMainIniPath, "MonitorConfig", "TestAbort", Me.chkTriggerItem(3).Value, moAppLog)
    Call ApIniWriteItem(App.Path & msMainIniPath, "MonitorConfig", "PEISUpdate", Me.chkTriggerItem(4).Value, moAppLog)
    Call ApIniWriteItem(App.Path & msMainIniPath, "MonitorConfig", "ReviewUpdate", Me.chkTriggerItem(5).Value, moAppLog)
    'Add by Weilun on 20200512 for Project.BE MES Phase 44 – CP 產線增加自動化作業 <start>
    Call ApIniWriteItem(App.Path & msMainIniPath, "MonitorConfig", Me.lblTriggerItem(6).Caption, Me.chkTriggerItem(6).Value, moAppLog) 'UpdateETtestMoveOut
    Call ApIniWriteItem(App.Path & msMainIniPath, "MonitorConfig", Me.lblTriggerItem(7).Caption, Me.chkTriggerItem(7).Value, moAppLog) 'AutoDgradeError
    Call ApIniWriteItem(App.Path & msMainIniPath, "MonitorConfig", Me.lblTriggerItem(8).Caption, Me.chkTriggerItem(8).Value, moAppLog) 'AutoHoldReleaseError
    'Add by Weilun on 20200512 for Project.BE MES Phase 44 – CP 產線增加自動化作業 <end>
    'Add by Weilun on 20211206 for Project.BE 工業 3.5 Phase 39 - CP AGV導入專案 - 建立 AGV 派工 & Status 監控系統 <start>
    Call ApIniWriteItem(App.Path & msMainIniPath, "MonitorConfig", Me.lblTriggerItem(9).Caption, Me.chkTriggerItem(9).Value, moAppLog) 'AGVLotStartError
    Call ApIniWriteItem(App.Path & msMainIniPath, "MonitorConfig", Me.lblTriggerItem(10).Caption, Me.chkTriggerItem(10).Value, moAppLog) 'AutoPrintSetupFormError
    Call ApIniWriteItem(App.Path & msMainIniPath, "MonitorConfig", Me.lblTriggerItem(11).Caption, Me.chkTriggerItem(11).Value, moAppLog) 'AutoPrintSetupFormNotice
    'Add by Weilun on 20211206 for Project.BE 工業 3.5 Phase 39 - CP AGV導入專案 - 建立 AGV 派工 & Status 監控系統 <end>
    Call ApIniWriteItem(App.Path & msMainIniPath, "MonitorConfig", Me.lblTriggerItem(12).Caption, Me.chkTriggerItem(12).Value, moAppLog) 'InsertSorterControl 'Add by Weilun on 20221230 for Req.SorterControl_Monitor
    
    Call ApIniWriteItem(App.Path & msMainIniPath, "MonitorConfig", "TestTimeAbnormalCheck", Me.chkSchedualItem(0).Value, moAppLog)
    Call ApIniWriteItem(App.Path & msMainIniPath, "MonitorConfig", "CPCompleteCheck", Me.chkSchedualItem(1).Value, moAppLog)
    Call ApIniWriteItem(App.Path & msMainIniPath, "MonitorConfig", "NextLotPrepare", Me.chkSchedualItem(2).Value, moAppLog)
    Call ApIniWriteItem(App.Path & msMainIniPath, "MonitorConfig", "CPAutoMonitor", Me.chkSchedualItem(3).Value, moAppLog)
    Call ApIniWriteItem(App.Path & msMainIniPath, "MonitorConfig", "UnloadTypeRefresh", Me.chkSchedualItem(4).Value, moAppLog) 'Add by Weilun on 20211126 for Project.BE 工業 3.5 Phase 39 - CP AGV導入專案 - 建立 AGV 派工 & Status 監控系統
    
    Call ApIniWriteItem(App.Path & msMainIniPath, "MonitorConfig", "StopTestCheckAndExecute", Me.chkStopTestItem.Value, moAppLog)
    
    'ADD By Rubio on 20230907 For MailBox
    Call ApIniWriteItem(App.Path & msMainIniPath, "MonitorConfig", "AutoReconnection", Me.chkAutoReconnection.Value, moAppLog)
    Call ApIniWriteItem(App.Path & msMainIniPath, "MonitorConfig", "UserName", Me.txtReUserName.Text, moAppLog)
    Call ApIniWriteItem(App.Path & msMainIniPath, "MonitorConfig", "Password", Me.txtRePassword.Text, moAppLog)
    
End Sub

'================================================================================
'Name       : cmdMonitorConfigClose_Click()
'Description: Close and reload Monitor Config
'Author     : Weilun Huang, MXIC, 2019/10/25
'Parameters : N/A
'================================================================================
Private Sub cmdMonitorConfigClose_Click()
On Error Resume Next
    Dim stxtReUserName As String
    Dim stxtRePassword As String
    Call LocateFrame(Me.fraMonitorConfig, False)
    
    '------
    Call LoadCheckBox(App.Path & msMainIniPath, "MonitorConfig", "EqChange", Me.chkTriggerItem(0))
    Call LoadCheckBox(App.Path & msMainIniPath, "MonitorConfig", "LotEnd", Me.chkTriggerItem(1))
    Call LoadCheckBox(App.Path & msMainIniPath, "MonitorConfig", "WaferEnd", Me.chkTriggerItem(2))
    Call LoadCheckBox(App.Path & msMainIniPath, "MonitorConfig", "TestAbort", Me.chkTriggerItem(3))
    Call LoadCheckBox(App.Path & msMainIniPath, "MonitorConfig", "PEISUpdate", Me.chkTriggerItem(4))
    Call LoadCheckBox(App.Path & msMainIniPath, "MonitorConfig", "ReviewUpdate", Me.chkTriggerItem(5))
    'Add by Weilun on 20200512 for Project.BE MES Phase 44 – CP 產線增加自動化作業 <start>
    Call LoadCheckBox(App.Path & msMainIniPath, "MonitorConfig", Me.lblTriggerItem(6).Caption, Me.chkTriggerItem(6)) 'UpdateETtestMoveOut
    Call LoadCheckBox(App.Path & msMainIniPath, "MonitorConfig", Me.lblTriggerItem(7).Caption, Me.chkTriggerItem(7)) 'AutoDgradeError
    Call LoadCheckBox(App.Path & msMainIniPath, "MonitorConfig", Me.lblTriggerItem(8).Caption, Me.chkTriggerItem(8)) 'AutoHoldReleaseError
    'Add by Weilun on 20200512 for Project.BE MES Phase 44 – CP 產線增加自動化作業 <end>
    'Add by Weilun on 20211206 for Project.BE 工業 3.5 Phase 39 - CP AGV導入專案 - 建立 AGV 派工 & Status 監控系統 <start>
    Call LoadCheckBox(App.Path & msMainIniPath, "MonitorConfig", Me.lblTriggerItem(9).Caption, Me.chkTriggerItem(9)) 'AGVLotStartError
    Call LoadCheckBox(App.Path & msMainIniPath, "MonitorConfig", Me.lblTriggerItem(10).Caption, Me.chkTriggerItem(10)) 'AutoPrintSetupFormError
    Call LoadCheckBox(App.Path & msMainIniPath, "MonitorConfig", Me.lblTriggerItem(11).Caption, Me.chkTriggerItem(11)) 'AutoPrintSetupFormNotice
    'Add by Weilun on 20211206 for Project.BE 工業 3.5 Phase 39 - CP AGV導入專案 - 建立 AGV 派工 & Status 監控系統 <end>
    Call LoadCheckBox(App.Path & msMainIniPath, "MonitorConfig", Me.lblTriggerItem(12).Caption, Me.chkTriggerItem(12)) 'InsertSorterControl 'Add by Weilun on 20221230 for Req.SorterControl_Monitor
    
    '------
    Call LoadCheckBox(App.Path & msMainIniPath, "MonitorConfig", "TestTimeAbnormalCheck", Me.chkSchedualItem(0))
    Call LoadCheckBox(App.Path & msMainIniPath, "MonitorConfig", "CPCompleteCheck", Me.chkSchedualItem(1))
    Call LoadCheckBox(App.Path & msMainIniPath, "MonitorConfig", "NextLotPrepare", Me.chkSchedualItem(2))
    Call LoadCheckBox(App.Path & msMainIniPath, "MonitorConfig", "CPAutoMonitor", Me.chkSchedualItem(3))
    Call LoadCheckBox(App.Path & msMainIniPath, "MonitorConfig", "UnloadTypeRefresh", Me.chkSchedualItem(4)) 'Add by Weilun on 20211126 for Project.BE 工業 3.5 Phase 39 - CP AGV導入專案 - 建立 AGV 派工 & Status 監控系統
    
    '------
    Call LoadCheckBox(App.Path & msMainIniPath, "MonitorConfig", "StopTestCheckAndExecute", Me.chkStopTestItem)
    
    Call LoadCheckBox(App.Path & msMainIniPath, "MonitorConfig", "AutoReconnection", Me.chkAutoReconnection)
    
    Call ApIniGetItem(App.Path & msMainIniPath, "MonitorConfig", "UserName", stxtReUserName)
    Call ApIniGetItem(App.Path & msMainIniPath, "MonitorConfig", "Password", stxtRePassword)
    Me.txtReUserName = stxtReUserName
    Me.txtRePassword = stxtRePassword
    
End Sub

'================================================================================
'Name       : LoadCheckBox()
'Description: Load CheckBox from Cfg
'Author     : Weilun Huang, MXIC, 2019/10/25
'Parameters : N/A
'================================================================================
Private Sub LoadCheckBox(ByVal sCfgFileName As String, ByVal sCfgSection As String, ByVal sCfgItem As String, ByRef chkObj As CheckBox)
On Error Resume Next
    Dim sGetValue   As String
    Dim iChkValue   As Integer
    Dim bLoad       As Boolean

    bLoad = False
    Call ApIniGetItem(sCfgFileName, sCfgSection, sCfgItem, sGetValue)
    If IsNumeric(sGetValue) = True Then
        iChkValue = CInt(sGetValue)
        If iChkValue = vbChecked Or iChkValue = vbUnchecked Then
            chkObj.Value = iChkValue
            bLoad = True
        End If
    End If
    
    If bLoad = False Then
        chkObj.Value = vbChecked
    End If
End Sub

'================================================================================
'Name       : cmdOK_Click()
'Description: Save Config and set LogLevel
'Author     : Weilun Huang, MXIC, 2018/03/27
'Parameters : N/A
'================================================================================
Private Sub cmdOK_Click()
On Error Resume Next
    Call ApIniWriteItem(App.Path & msMainIniPath, "Config", "LogLevel", VBA.Left(Me.lblLogLevel, 1), moAppLog)
    Call ApIniWriteItem(App.Path & msMainIniPath, "Config", "EMAIL", Me.txtEMail(0), moAppLog)

    Call ApIniWriteItem(App.Path & msMainIniPath, "Config", "CC", Me.txtCC(0), moAppLog)

    Call ApIniWriteItem(App.Path & msMainIniPath, "Config", "SMTP", Me.txtSmtp, moAppLog)
    
    Call LogLevelChange
    
    Call LocateFrame(Me.fraConfig, False)
    Call LocateFrame(Me.fraConnect, True)
End Sub

'================================================================================
'Name       : LogLevelChange()
'Description: 設定LogLevel
'Author     : Weilun huang, MXIC 2018/03/27
'Parameters : N/A
'================================================================================
Private Sub LogLevelChange()
On Error Resume Next
    'LogLevel
    If Not moAppLog Is Nothing Then
        Select Case Left(Me.lblLogLevel, 1)
            Case "0":
                moAppLog.LogLevel = LS_LOGLEVEL_0
            Case "1":
                moAppLog.LogLevel = LS_LOGLEVEL_1
            Case "5":
                moAppLog.LogLevel = LS_LOGLEVEL_5
            Case "9":
                moAppLog.LogLevel = LS_LOGLEVEL_9
        End Select
    End If
End Sub

'================================================================================
'Name       : optMailType_Click()
'Description:
'Author     : Weilun Huang, MXIC, 2018/03/27
'Parameters : N/A
'================================================================================
Private Sub optMailType_Click(Index As Integer)
On Error Resume Next
    'LogLevel
    txtEMail(0).Visible = False
    txtCC(0).Visible = False

    Select Case Index
        Case 0:
            Me.txtEMail(0).ZOrder 0
            Me.txtCC(0).ZOrder 0
            
            txtEMail(0).Visible = True
            txtCC(0).Visible = True
            
            txtEMail(1).Visible = False
            txtCC(1).Visible = False
        Case 1:
    End Select
End Sub

'================================================================================
'Name       : ApIniGetEmailItem()
'Description: 轉帳專屬取得設定的Email
'Author     : Weilun huang, MXIC 2018/03/19
'Parameters : N/A
'================================================================================
Public Sub ApIniGetEmailItem(ByRef sCfgFileName As String, ByRef sCfgSection As String, _
                        ByRef sCfgItem As String, ByRef sItemValue As String, _
                        Optional ByRef sDefValue As String, Optional ByRef oLogCtrl As Object)
    On Error GoTo ExitHandler
    Dim sProcID As String
    Dim typErrInfo As tErrInfo
    Dim dwInfo As Integer

    '----
    ' Init
    '----
    sProcID = "ApIniGetEmailItem"
    Call LogProcIn(msMODULE_ID, sProcID, oLogCtrl)
    'ps.msMODULE_ID = "modWinAPI"
    '----
    ' Condition Checking
    '----
    ' <Put your condition checking codes here>...

    '----
    ' Action
    '----
    sItemValue = Space$(1024)

    dwInfo = ApiGetProfStr(sCfgSection, sCfgItem, sDefValue, sItemValue, Len(sItemValue) + 1, sCfgFileName)

    If dwInfo = 0 And Len(sDefValue) <> 0 Then
        sItemValue = sDefValue
        Call LogTrace("Item [" & sCfgItem & "] assigned with DEFAULT value=[" & sItemValue & "]", oLogCtrl, glLOG_TRACE, msMODULE_ID, sProcID)
    ElseIf dwInfo <> 0 Then
        sItemValue = Mid$(sItemValue, 1, dwInfo)
        Call LogTrace("Item [" & sCfgItem & "] assigned with value=[" & sItemValue & "]", oLogCtrl, glLOG_TRACE, msMODULE_ID, sProcID)
    Else
        sItemValue = ""
        Call LogTrace("Item [" & sCfgItem & "] assigned with value=[" & sItemValue & "]", oLogCtrl, glLOG_TRACE, msMODULE_ID, sProcID)
    End If

    '-------
    'Done
    '-------
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
            Case Else
                typErrInfo.sUserText = typErrInfo.sErrDescription
        End Select
        '---- Start of Select Case Block ----
        On Error GoTo ExitHandler
        'Call HandleError(False, typErrInfo, , oLogCtrl, , True)
        Call TransferLogError(typErrInfo.sUserText, _
              moAppLog, Me.rtfTrace, msMODULE_ID, sProcID)
    End If
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

Private Sub TimerAutoReconnection_Timer()
    Dim sErrorConnectFlag As String
    'ADD By Rubio on 2023/9/7 For Mailbox AutoReconnection
    'Call ApIniGetItem(App.Path & msMainIniPath, "ConnectSetting", "ErrorConnectFlag", sErrorConnectFlag)
    Call cmdConnect_Click
    TimerAutoReconnection.Enabled = False

End Sub

Private Sub TimerMain_Timer()
    On Error GoTo ExitHandler:
    Dim sProcID                     As String
    Dim typErrInfo                  As tErrInfo
    
    Dim sSQL                        As String
    Dim colRaws                     As Collection
    Dim oRaws                       As Collection
    '-----
    ' Init
    '-----
    sProcID = "TimerMain_Timer"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog) '

    If Hour(Now) = 9 Then '每天早上9點進行
        If mbDaily_Maintain = True Then
            Call MonitorMaintain
            mbDaily_Maintain = False '執行過後在同一小時內不再執行
        End If
    Else
        mbDaily_Maintain = True
    End If


    mlTimer = mlTimer + 1

    If mlTimer Mod 5 = 0 Then      '5分鐘執行並歸0一次
'    If mlTimer Mod 1 = 0 Then      '1分鐘執行並歸0一次
        Call TransferLogTrace("------------------------------------------- ", _
                      moAppLog, Me.rtfTrace)
        Call TransferLogTrace(DateTime.Now & ": Start to process Schedual Check.", _
                              moAppLog, Me.rtfTrace)
                      
                      
        Call TestTimeAbnormalCheck
        Call CPCompleteCheck
        Call NextLotPrepare
        Call UnloadTypeRefresh 'Add by Weilun on 20211126 for Project.BE 工業 3.5 Phase 39 - CP AGV導入專案 - 建立 AGV 派工 & Status 監控系統
        Call CPAutoMonitor
        
        '順便檢查一下有沒有漏掉的PEIS未補資料
        If TimerPEISUpdate.Enabled = False And TimerPEISBuffer.Enabled = False And chkTriggerItem(4).Value = Checked Then
            Call TransferLogTrace(DateTime.Now & ": Search PEIS inputed data to update.", _
                                  moAppLog, Me.rtfTrace)
            Call PEISUpdate
            
            '檢查自動停機
            Call StopTestCheckAndExecute
        End If
        
        
    If Not moSC_Main Is Nothing And Not moSC_Sub Is Nothing Then
        If moSC_Main.Count = 0 And Not moSC_Sub.Count = 0 Then
            '待辦(Main)和Buffer(Sub)都沒資料可執行時, 這個Timer就可以先關了
            '直到有新資料進來或是整體檢查時特別檢查一下
            TimerSorterControl.Enabled = False
        End If
    End If
        
        'Add by Weilun on 20221230 for Req.SorterControl_Monitor <start>
        '檢查SorterControl資料
        If TimerSorterControl.Enabled = False And _
           (Not moSC_Main Is Nothing And Not moSC_Sub Is Nothing) And _
           chkTriggerItem(12).Value = Checked Then
            If moSC_Main.Count > 0 Or Not moSC_Sub.Count > 0 Then
                Call TransferLogTrace(DateTime.Now & ": Search SorterControl inputed data to insert.", _
                                      moAppLog, Me.rtfTrace)
                TimerSorterControl.Enabled = True
            End If
        End If
        'Add by Weilun on 20221230 for Req.SorterControl_Monitor <end>
        
        Call Level1Sort '全部整理
        Call UpdateMonitorAllEq
        
        Call TransferLogTrace(DateTime.Now & ": End of Schedual Check.", _
                              moAppLog, Me.rtfTrace)
        
        mlTimer = 0
    End If

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
           Case Else
                typErrInfo.sUserText = typErrInfo.sErrDescription
        End Select
        '---- Start of Select Case Block ----
        On Error GoTo ExitHandler:
        Call TransferLogError(typErrInfo.sUserText, _
                              moAppLog, Me.rtfTrace, msMODULE_ID, sProcID)
        Call HandleError(False, typErrInfo, , moAppLog)
    End If
End Sub

'主要的PEIS資訊更新
Private Sub TimerPEISUpdate_Timer()
    On Error GoTo ExitHandler:
    Dim sProcID                     As String
    Dim typErrInfo                  As tErrInfo
    
    '-----
    ' Init
    '-----
    sProcID = "TimerPEISUpdate_Timer"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog) '
    
    Call TransferLogTrace("------------------------------------------- ", _
                          moAppLog, Me.rtfTrace)
    Call TransferLogTrace(DateTime.Now & ": Search PEIS inputed data to update.", _
                          moAppLog, Me.rtfTrace)
    Call PEISUpdate
    
    '檢查自動停機
    Call StopTestCheckAndExecute
    
    Call Level1Sort '全部整理
    Call UpdateMonitorAllEq
    
    TimerPEISUpdate.Enabled = False '自己關閉
    
    TimerPEISBuffer.Enabled = True '延遲執行的開啟
    
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
           Case Else
                typErrInfo.sUserText = typErrInfo.sErrDescription
        End Select
        '---- Start of Select Case Block ----
        On Error GoTo ExitHandler:
        Call TransferLogError(typErrInfo.sUserText, _
                              moAppLog, Me.rtfTrace, msMODULE_ID, sProcID)
        Call HandleError(False, typErrInfo, , moAppLog)
    End If
End Sub

'避免第一次更新PEIS資訊更新時剛好有新的資料進來導致錯過補資料
Private Sub TimerPEISBuffer_Timer()
    On Error GoTo ExitHandler:
    Dim sProcID                     As String
    Dim typErrInfo                  As tErrInfo
    
    '-----
    ' Init
    '-----
    sProcID = "TimerPEISBuffer_Timer"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog) '
    
    Call PEISUpdate
    
    '檢查自動停機
    Call StopTestCheckAndExecute
    
    Call Level1Sort '全部整理
    Call UpdateMonitorAllEq
    
    TimerPEISBuffer.Enabled = False
    
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
           Case Else
                typErrInfo.sUserText = typErrInfo.sErrDescription
        End Select
        '---- Start of Select Case Block ----
        On Error GoTo ExitHandler:
        Call TransferLogError(typErrInfo.sUserText, _
                              moAppLog, Me.rtfTrace, msMODULE_ID, sProcID)
        Call HandleError(False, typErrInfo, , moAppLog)
    End If
End Sub

'================================================================================
'Name       : UpdateMonitorAllEq()
'Description: 下的指令為"UpdateAll", exe檔只負責轉送動作(UpdateAll)定義與程式都在Monitor網頁上
'Author     : Weilun huang, MXIC 2019/08/30
'Parameters : N/A
'================================================================================
Private Function UpdateMonitorAllEq() As Boolean
    On Error GoTo ExitHandler:
    Dim sProcID                     As String
    Dim typErrInfo                  As tErrInfo
    
    Dim sSQL                        As String
    Dim colRaws                     As Collection
    Dim oRaws                       As Collection
    
    Dim sMsg                        As String 'Add by Weilun on 20200819 for Req.VB 嘗試引用 C# DLL
    '-----
    ' Init
    '-----
    sProcID = "UpdateMonitorAllEq"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog) '
    
    UpdateMonitorAllEq = False
    '--------------------
    ' Condition Checking
    '--------------------
    
    '-------
    ' Action
    '-------
    If mbReadySendSignalR = True Then 'Add by Weilun on 20200203 for Req.BE MES Phase 41 - 手動停機修改, 系統背景執行SortLotEnd
    
        'Add by Weilun on 20200819 for Req.VB 嘗試引用 C# DLL <start>
        If Not moClient Is Nothing Then
            sMsg = moClient.SendFunctionAndCommand("UpdateAll", "")
            If sMsg = moClient.ClientState_SendSuccess Then
                Call TransferLogTrace(DateTime.Now & ": UpdateMonitorAllEq: " & sMsg, _
                                      moAppLog, Me.rtfTrace)
            Else
                Call TransferLogError(DateTime.Now & ": UpdateMonitorAllEq: " & sMsg, _
                                      moAppLog, Me.rtfTrace)
            End If
        End If
        'Add by Weilun on 20200819 for Req.VB 嘗試引用 C# DLL <end>
        
        'Mark by Weilun on 20200819 for Req.VB 嘗試引用 C# DLL
'        ShellExecute Me.hWnd, "open", App.Path & msMonitorExe, "UpdateAll", App.Path, SW_HIDE
    End If

    UpdateMonitorAllEq = True
    '-------
    'Done
    '-------
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
           Case Else
                typErrInfo.sUserText = typErrInfo.sErrDescription
        End Select
        '---- Start of Select Case Block ----
        On Error GoTo ExitHandler:
        Call TransferLogError(typErrInfo.sUserText, _
                              moAppLog, Me.rtfTrace, msMODULE_ID, sProcID)
        Call HandleError(False, typErrInfo, , moAppLog)
    End If
End Function

'================================================================================
'Name       : UpdateMonitorSingleEq()
'Description: 下的指令為"UpdateEq EqID", exe檔只負責轉送動作(UpdateEq)與唯一參數(EqID), 定義與程式都在Monitor網頁上
'Author     : Weilun huang, MXIC 2019/08/30/
'Parameters : N/A
'================================================================================
Private Function UpdateMonitorSingleEq(ByVal sEqId As String) As Boolean
    On Error GoTo ExitHandler:
    Dim sProcID                     As String
    Dim typErrInfo                  As tErrInfo
    
    Dim sSQL                        As String
    Dim colRaws                     As Collection
    Dim oRaws                       As Collection
        
    Dim sMsg                        As String 'Add by Weilun on 20200819 for Req.VB 嘗試引用 C# DLL
    '-----
    ' Init
    '-----
    sProcID = "UpdateMonitorSingleEq"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog) '
    
    UpdateMonitorSingleEq = False
    '--------------------
    ' Condition Checking
    '--------------------
    
    '-------
    ' Action
    '-------
    If mbReadySendSignalR = True Then 'Add by Weilun on 20200203 for Req.BE MES Phase 41 - 手動停機修改, 系統背景執行SortLotEnd
    
        'Add by Weilun on 20200819 for Req.VB 嘗試引用 C# DLL <start>
        If Not moClient Is Nothing Then
            sMsg = moClient.SendFunctionAndCommand("UpdateEq", sEqId)
            
            If sMsg = moClient.ClientState_SendSuccess Then
                Call TransferLogTrace(DateTime.Now & ": UpdateMonitorSingleEq: " & sMsg, _
                                      moAppLog, Me.rtfTrace)
            Else
                Call TransferLogError(DateTime.Now & ": UpdateMonitorSingleEq: " & sMsg, _
                                      moAppLog, Me.rtfTrace)
            End If
        End If
        'Add by Weilun on 20200819 for Req.VB 嘗試引用 C# DLL <end>
    
        'Mark by Weilun on 20200819 for Req.VB 嘗試引用 C# DLL
'        ShellExecute Me.hWnd, "open", App.Path & msMonitorExe, "UpdateEq " & sEqID, App.Path, SW_HIDE
    End If

    UpdateMonitorSingleEq = True
    '-------
    'Done
    '-------
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
           Case Else
                typErrInfo.sUserText = typErrInfo.sErrDescription
        End Select
        '---- Start of Select Case Block ----
        On Error GoTo ExitHandler:
        Call TransferLogError(typErrInfo.sUserText, _
                              moAppLog, Me.rtfTrace, msMODULE_ID, sProcID)
        Call HandleError(False, typErrInfo, , moAppLog)
    End If
End Function

'================================================================================
'Name       : MonitorInit()
'Description: 於Core啟動時更新所有機台狀態
'Author     : Weilun huang, MXIC 2019/09/20
'Parameters : N/A
'================================================================================
Private Sub MonitorInit()
    On Error GoTo ExitHandler
    Dim sProcID As String
    Dim typErrInfo As tErrInfo

    Dim sSQL As String
    Dim colRaws As Collection
    Dim oRaws As Collection

    Dim sModifySQL As String
    Dim sEqId As String

    Dim sEqTypeList         As String 'Add by Weilun on 20200512 for Project.BE MES Phase 44 – CP 產線增加自動化作業
    
    '-----
    ' Init
    '-----
    sProcID = "MonitorInit"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog) '

    '--------------------
    ' Condition Checking
    '--------------------

    '-------
    ' Action
    '-------
    Call TransferLogTrace(DateTime.Now & ": Monitor Init start.", _
                          moAppLog, Me.rtfTrace)

'    sSQL = " select equipmentname, state " & _
'             " from fweqpcurrentstate " & _
'            " where eqptype = 'WTESTER' " & _
'              " and activeversion >= 0 "
    'Modify by Weilun on 20200512 for Project.BE MES Phase 44 – CP 產線增加自動化作業 <start>
    '改用in List的方法, 以後能擴充.
    
    sEqTypeList = "WTESTER"                              '初版
    sEqTypeList = sEqTypeList & "," & "HoldLotRelease"   'Add for Project.BE MES Phase 44
    sEqTypeList = sEqTypeList & "," & "DGRADE"           'Add for Project.BE MES Phase 44
    sEqTypeList = sEqTypeList & "," & "WSORTER"          'Add by Weilun on 20230103 for Req.SorterControl_Monitor
    sEqTypeList = sEqTypeList & "," & "N2"           'Add by Weilun on 20230717 for Req.當 N2櫃異常時, 顯示在 CP 監控系統
    '需要加新機台
    
    
    
    sSQL = " select equipmentname, state " & _
             " from fweqpcurrentstate " & _
            " where eqptype in ('" & Replace(sEqTypeList, ",", "','") & "') " & _
              " and activeversion >= 0 " & _
         " order by equipmentname "
    'Modify by Weilun on 20200512 for Project.BE MES Phase 44 – CP 產線增加自動化作業 <end>

    Set colRaws = moProRawSql.QueryDatabase(sSQL)

    If colRaws.Count > 0 Then
        For Each oRaws In colRaws
            Call EqStateChnage(gsMONITOR_TYPE_INIT & "/" & oRaws.Item("equipmentname") & "/" & oRaws.Item("state"))
        Next
    End If

    Call TransferLogTrace(DateTime.Now & ": Monitor Init end.", _
                          moAppLog, Me.rtfTrace)
    '更新全部User, 由於該功能可被合併, 建議一起呼叫
    '-------
    'Done
    '-------
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
            Case Else
                typErrInfo.sUserText = typErrInfo.sErrDescription
        End Select
        '---- Start of Select Case Block ----
        On Error GoTo ExitHandler
        Call TransferLogError(typErrInfo.sUserText, _
                              moAppLog, Me.rtfTrace, msMODULE_ID, sProcID)
        Call HandleError(False, typErrInfo, , moAppLog)
    End If
End Sub

'================================================================================
'Name       : MonitorMaintain()
'Description: 機台維護
'Author     : Weilun huang, MXIC 2019/09/20
'Parameters : N/A
'================================================================================
Private Sub MonitorMaintain()
    On Error GoTo ExitHandler
    Dim sProcID As String
    Dim typErrInfo As tErrInfo

    Dim sSQL As String
    Dim colRaws As Collection
    Dim oRaws As Collection

    Dim sModifySQL As String
    Dim sEqId As String

    '-----
    ' Init
    '-----
    sProcID = "MonitorMaintain"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog) '

    '--------------------
    ' Condition Checking
    '--------------------

    '-------
    ' Action
    '-------
    '機台移除, 一天一次即可
    Call TransferLogTrace(DateTime.Now & ": Monitor maintain start.", _
                          moAppLog, Me.rtfTrace)

    sSQL = " select tcml1." & gsCAT_TCML1_EQID & " " & _
             " from " & gsCAT_TBL_CPETEST_MONITOR_LEVEL1 & " tcml1, fweqpcurrentstate eqp" & _
            " Where (eqp.ActiveVersion < 0 Or eqp.ActiveVersion Is Null) " & _
              " and tcml1." & gsCAT_TCML1_EQID & " = eqp.equipmentname(+) " & _
              " and tcml1." & gsCAT_TCML1_DELETEFLAG & " = 'N' "
    Set colRaws = moProRawSql.QueryDatabase(sSQL)
    If colRaws.Count > 0 Then
        For Each oRaws In colRaws
            sEqId = oRaws.Item(gsCAT_TCML1_EQID)
            Call TransferLogTrace(DateTime.Now & ": Remove EqID: " & sEqId, _
                                  moAppLog, Me.rtfTrace, msMODULE_ID, sProcID)

            '補上逗點, Modify by Weilun on 20200203 for Req.BE MES Phase 41 - 手動停機修改, 系統背景執行SortLotEnd
            sModifySQL = " update " & gsCAT_TBL_CPETEST_MONITOR_LEVEL1 & " " & _
                            " set " & gsCAT_TCML1_DELETEFLAG & " = 'Y', " & _
                                      gsCAT_TCML1_UPDATEUSERID & " = 'SYS', " & _
                                      gsCAT_TCML1_UPDATETIME & " = to_char(sysdate, 'yyyymmdd hh24miss') || '000' " & _
                          " where " & gsCAT_TCML1_EQID & " = '" & sEqId & "' " & _
                            " and " & gsCAT_TCML1_DELETEFLAG & " = 'N' "
            Call moProRawSql.QueryDatabase(sModifySQL) 'Add by Weilun on 20200203 for Req.BE MES Phase 41 - 手動停機修改, 系統背景執行SortLotEnd
                            
        Next
        '更新全部User, 由於該功能可被合併, 建議一起呼叫
    End If

    Call TransferLogTrace(DateTime.Now & ": Monitor maintain end.", _
                          moAppLog, Me.rtfTrace)
    '-------
    'Done
    '-------
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
            Case Else
                typErrInfo.sUserText = typErrInfo.sErrDescription
        End Select
        '---- Start of Select Case Block ----
        On Error GoTo ExitHandler
        Call TransferLogError(typErrInfo.sUserText, _
                              moAppLog, Me.rtfTrace, msMODULE_ID, sProcID)
        Call HandleError(False, typErrInfo, , moAppLog)
    End If
End Sub

'================================================================================
'Name       : EqMonitorDataCheck()
'Description: 檢查機台資料是否存在, 不存在便新增
'Author     : Weilun huang, MXIC 2019/09/04
'Parameters : N/A
'================================================================================
Private Sub EqMonitorDataCheck(ByVal oEquipment As FwEquipment)
    On Error GoTo ExitHandler
    Dim sProcID As String
    Dim typErrInfo As tErrInfo

    Dim sSQL As String
    Dim colRaws As Collection
    Dim oRaws As Collection
    '-----
    ' Init
    '-----
    sProcID = "EqMonitorDataCheck"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog) '

    '--------------------
    ' Condition Checking
    '--------------------

    '-------
    ' Action
    '-------
    If oEquipment Is Nothing Then
        Call TransferLogError("該Eq不存在.", _
                              moAppLog, Me.rtfTrace, msMODULE_ID, sProcID)
        GoTo ExitHandler
    '特定機台不要 'Add by Weilun on 20200203 for Req.BE MES Phase 41 - 手動停機修改, 系統背景執行SortLotEnd <start>
    ElseIf oEquipment.CustomAttributes(gsEQP_CUSTOMATTR_EQ_TYPE2) = "WTESTER(DUMMY)" Then
        Call TransferLogTrace("Skip Dummy Equipment.", _
                              moAppLog, Me.rtfTrace, msMODULE_ID, sProcID)
        GoTo ExitHandler
    '特定機台不要 'Add by Weilun on 20200203 for Req.BE MES Phase 41 - 手動停機修改, 系統背景執行SortLotEnd <end>
    'Add by Weilun on 20230103 for Req.SorterControl_Monitor <start>
    'WSORTER只要zWSORTER就好
    ElseIf oEquipment.Type = "WSORTER" And oEquipment.Id <> "zWSORTER" Then
        '這裡不顯示訊息
        GoTo ExitHandler
    'Add by Weilun on 20230103 for Req.SorterControl_Monitor <end>
    Else
        sSQL = " select " & gsCAT_TCML1_EQID & " " & _
                 " from " & gsCAT_TBL_CPETEST_MONITOR_LEVEL1 & " " & _
                " where " & gsCAT_TCML1_EQID & " = '" & oEquipment.Id & "' " & _
                  " and " & gsCAT_TCML1_DELETEFLAG & " = 'N' "

        Set colRaws = moProRawSql.QueryDatabase(sSQL)
        If colRaws.Count = 0 Then
            sSQL = " insert into " & gsCAT_TBL_CPETEST_MONITOR_LEVEL1 & " " & _
                           " (" & gsCAT_TCML1_EQID & ", " & _
                                  gsCAT_TCML1_CREATETIME & ", " & _
                                  gsCAT_TCML1_CREATEUSERID & ") " & _
                   " values ('" & oEquipment.Id & "', " & _
                                " to_char(sysdate, 'yyyymmdd hh24miss') || '000', " & _
                                " 'SYS')"
            Call moProRawSql.QueryDatabase(sSQL)
        End If
    End If
    '-------
    'Done
    '-------
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
            Case Else
                typErrInfo.sUserText = typErrInfo.sErrDescription
        End Select
        '---- Start of Select Case Block ----
        On Error GoTo ExitHandler
        Call TransferLogError(typErrInfo.sUserText, _
                              moAppLog, Me.rtfTrace, msMODULE_ID, sProcID)
        Call HandleError(False, typErrInfo, , moAppLog)
    End If
End Sub

'================================================================================
'Name       : InsertMonitorReceive()
'Description: 通用Receive寫入模版
'Author     : Weilun huang, MXIC 2019/09/25
'Parameters : N/A
'================================================================================
Private Sub InsertMonitorReceive(ByVal sEqId As String, _
                                 ByVal sLotId As String, _
                                 ByVal sWaferID As String, _
                                 ByVal sStepName As String, _
                                 ByVal sMonitorItem As String, _
                                 ByVal sMonitorItemCont As String, _
                                 ByVal sMonitorItemContDesc As String, _
                                 ByVal sAlarmLevel As String, _
                                 ByVal sReviewFlag As String)
    On Error GoTo ExitHandler
    Dim sProcID As String
    Dim typErrInfo As tErrInfo

    Dim sSQL As String
    Dim colRaws As Collection

    Dim sNextStepName As String
    Dim sNextStepRecipe As String
    Dim sHotFlag As String
    Dim sLotSeq As String
    
    Dim oEquipment              As FwEquipment '特定機台不要 'Add by Weilun on 20200203 for Req.BE MES Phase 41 - 手動停機修改, 系統背景執行SortLotEnd
    Dim sLotEndFlag             As String      'Add by Weilun on 20211125 for Project.BE 工業 3.5 Phase 39 - CP AGV導入專案 - 建立 AGV 派工 & Status 監控系統
    Dim sUnloadType             As String      'Add by Weilun on 20211125 for Project.BE 工業 3.5 Phase 39 - CP AGV導入專案 - 建立 AGV 派工 & Status 監控系統

    '----
    ' Init
    '----
    sProcID = "InsertMonitorReceive"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog)

    '----
    ' Condition Chking
    '----
    ' <Put your condition checking codes here>...
    
    '特定機台不要 'Add by Weilun on 20200203 for Req.BE MES Phase 41 - 手動停機修改, 系統背景執行SortLotEnd <start>
    Set oEquipment = moFwMDL.EquipmentById(sEqId)
    If oEquipment Is Nothing Then
        Call TransferLogError("該Eq不存在.", _
                              moAppLog, Me.rtfTrace, msMODULE_ID, sProcID)
        GoTo ExitHandler
    ElseIf oEquipment.CustomAttributes(gsEQP_CUSTOMATTR_EQ_TYPE2) = "WTESTER(DUMMY)" Then
        GoTo ExitHandler
    End If
    '特定機台不要 'Add by Weilun on 20200203 for Req.BE MES Phase 41 - 手動停機修改, 系統背景執行SortLotEnd <end>

    If sWaferID = "" Then
        sWaferID = "NA"
    End If

    If sAlarmLevel = "" Then
        sAlarmLevel = "99"
    End If

    If Not (sReviewFlag = gsCAT_TCMR_PRODTAFLAG Or _
            sReviewFlag = gsCAT_TCMR_HWFLAG Or _
            sReviewFlag = gsCAT_TCMR_PEFLAG) Then
        GoTo ExitHandler
    End If
    '----
    ' Action
    '----
    '取得LotSeq, 有就帶進去
    sLotSeq = GetWsLotSeq(sLotId, sEqId)

    '取得HotFlag
    sHotFlag = GetHotFlag(sLotId)

    '取得下一站, 如果是Bake會再取得Recipe
    Call GetNextStepNameAndRecipe(sLotId, sNextStepName, sNextStepRecipe)
    
    'Add by Weilun on 20211126 for Project.BE 工業 3.5 Phase 39 - CP AGV導入專案 - 建立 AGV 派工 & Status 監控系統 <start>
    sLotEndFlag = ""
    sUnloadType = ""
    If sMonitorItemContDesc = gsMONITOR_ITEM_DESC_UNLOAD Then
        sLotEndFlag = "Y"
        sUnloadType = GetUnloadType(sEqId, sLotId)
    End If
    'Add by Weilun on 20211126 for Project.BE 工業 3.5 Phase 39 - CP AGV導入專案 - 建立 AGV 派工 & Status 監控系統 <end>

    'LOTENDFLAG和UNLOADTYPE, Add by Weilun on 20211126 for Project.BE 工業 3.5 Phase 39 - CP AGV導入專案 - 建立 AGV 派工 & Status 監控系統
    sSQL = " insert into " & gsCAT_TBL_CPETEST_MONITOR_RECEIVE & " " & _
                      " (" & gsCAT_TCMR_SID & ", " & _
                             gsCAT_TCMR_EQID & ", " & _
                             gsCAT_TCMR_LOTID & ", " & _
                             gsCAT_TCMR_LOTSEQ & ", " & _
                             gsCAT_TCMR_WAFERID & ", " & _
                             gsCAT_TCMR_STEPNAME & ", " & _
                             gsCAT_TCMR_NEXTSTEPNAME & ", " & _
                             gsCAT_TCMR_RECIPE & ", " & _
                             gsCAT_TCMR_HOTFLAG & ", " & _
                             gsCAT_TCMR_MONITORITEM & ", " & _
                             gsCAT_TCMR_MONITORITEMCONT & ", " & _
                             gsCAT_TCMR_MONITORITEMCONTDESC & ", " & _
                             gsCAT_TCMR_ALARMLEVEL & ", " & _
                             sReviewFlag & ", " & _
                             gsCAT_TCMR_LOTENDFLAG & ", " & _
                             gsCAT_TCMR_UNLOADTYPE & ", " & _
                             gsCAT_TCMR_CREATETIME & ", " & _
                             gsCAT_TCMR_CREATEUSERID & ") "
    sSQL = sSQL & _
           " values ( CPETEST_MONIOTR_SEQ.NEXTVAL, " & _
                    " '" & sEqId & "', " & _
                    " '" & sLotId & "', " & _
                    " '" & sLotSeq & "', " & _
                    " '" & sWaferID & "', " & _
                    " '" & sStepName & "', " & _
                    " '" & sNextStepName & "', " & _
                    " '" & sNextStepRecipe & "', " & _
                    " '" & sHotFlag & "', " & _
                    " '" & sMonitorItem & "', " & _
                    " '" & sMonitorItemCont & "', " & _
                    " '" & sMonitorItemContDesc & "', " & _
                    " " & sAlarmLevel & ", " & _
                    " 'Y', " & _
                    " '" & sLotEndFlag & "', " & _
                    " '" & sUnloadType & "', " & _
                    " to_char(sysdate, 'yyyymmdd hh24miss') || '000', " & _
                    " 'SYS')"
    Call moProRawSql.QueryDatabase(sSQL)

    '-------
    'Done
    '-------
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
            Case Else
                typErrInfo.sUserText = typErrInfo.sErrDescription
        End Select
        '---- Start of Select Case Block ----
        On Error GoTo ExitHandler
        Call TransferLogError(typErrInfo.sUserText, _
                              moAppLog, Me.rtfTrace, msMODULE_ID, sProcID)
        Call HandleError(False, typErrInfo, , moAppLog)
    End If
End Sub

'================================================================================
'Name       : StopTestCheckAndExecute()
'Description: 搜尋停機資訊並進行停機
'Author     : Weilun huang, MXIC 2019/10/17
'Parameters : N/A
'================================================================================
Private Sub StopTestCheckAndExecute(Optional ByVal sSID As String = "")
    On Error GoTo ExitHandler
    Dim sProcID As String
    Dim typErrInfo As tErrInfo

    Dim sSQL As String
    Dim colRaws As Collection
    Dim oRaws As Collection

    Dim sModifySQL          As String
    
    Dim sMonitorItemCont    As String
    Dim sActionType         As String
    
    Dim sAutoValidTime      As String
    Dim sManualValidTime    As String
    Dim bValid              As Boolean
    
    Dim sHWFlag             As String
    Dim sPEFlag             As String
    
    Dim sMsg                As String
    Dim sMsgSQL             As String
    
    'Add by Weilun on 20200203 for Req.BE MES Phase 41 - 手動停機修改, 系統背景執行SortLotEnd <start>
    Dim sSubSQL             As String
    Dim colSubRaws          As Collection
    
    Dim sLotId              As String
    Dim sEqId               As String
    Dim sReviewUserID       As String
    Dim oLot                As FwLot
    Dim oEquipment          As FwEquipment
    'Add by Weilun on 20200203 for Req.BE MES Phase 41 - 手動停機修改, 系統背景執行SortLotEnd <end>

    Dim bOverReviewFlag     As Boolean 'Add by Weilun on 20200731 for ReqNo. BE#202000220
    Dim sExtraSQL           As String  'Add by Weilun on 20200731 for ReqNo. BE#202000220

    '-----
    ' Init
    '-----
    sProcID = "StopTestCheckAndExecute"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog) '
    
    sAutoValidTime = "1/288" '5分鐘
    sManualValidTime = "1/48" '30分鐘

    '--------------------
    ' Condition Checking
    '--------------------
    If chkStopTestItem.Value = vbUnchecked Then
        GoTo ExitHandler
    End If
    
    '-------
    ' Action
    '-------
    'REVIEWFLAG為S表示要停機, 包含自動停機先經由UpdatePEIS補完資料後改S以及User前端Review後
    
    If IsNumeric(sSID) = True Then '有傳SID, 無視ReviewFlag(還是舊的)
        '手動停機
        'HW和PE強制由N轉Y, 除了辨識用, 主要是之後還要重新Review, Modify by Weilun on 20191107 for 智慧監控整合
        Sleep (1000) ' 網頁同時在改Owner Flag, 加個一秒等他寫完, 不然好像會取到空值 'Add by Weilun on 20191122 for 智慧監控整合
        'EqID, LotID, updateuserid, Add by Weilun on 20200206 for Req.BE MES Phase 41 - 手動停機修改, 系統背景執行SortLotEnd
        sSQL = "select " & gsCAT_TCMR_SID & ", " & _
                           gsCAT_TCMR_EQID & ", " & _
                           gsCAT_TCMR_LOTID & ", " & _
                           gsCAT_TCMR_UPDATEUSERID & ", " & _
                         " case " & _
                           " when " & gsCAT_TCMR_HWFLAG & " = 'N' then 'Y' else " & gsCAT_TCMR_HWFLAG & " End As " & gsCAT_TCMR_HWFLAG & ", " & _
                         " case " & _
                           " when " & gsCAT_TCMR_PEFLAG & " = 'N' then 'Y' else " & gsCAT_TCMR_PEFLAG & " End As " & gsCAT_TCMR_PEFLAG & ", " & _
                           gsCAT_TCMR_MONITORITEMCONT & ", " & _
                           gsCAT_TCMR_EQSTOPTESTTYPE & ", " & _
                         " case " & _
                           " when " & gsCAT_TCMR_CREATETIME & " > " & _
                                " to_char(sysdate - " & sManualValidTime & ", 'yyyymmdd hh24miss') || '000' then 'Y' Else 'N' End As ValidFlag " & _
                " from " & gsCAT_TBL_CPETEST_MONITOR_RECEIVE & " " & _
               " where " & gsCAT_TCMR_EQSTOPTEST & " = 'Y' " & _
                 " and " & gsCAT_TCMR_EQSTOPTESTTYPE & " = '" & gsSTOPTEST_TYPE_MANUAL & "' " & _
                 " and " & gsCAT_TCMR_SID & " = " & sSID & " " & _
                 " and " & gsCAT_TCMR_CREATETIME & " is not null " & _
                 " and " & gsCAT_TCMR_DELETEFLAG & " = 'N' "
    Else
        '自動停機
        'EqID, LotID, updateuserid, Add by Weilun on 20200206 for Req.BE MES Phase 41 - 手動停機修改, 系統背景執行SortLotEnd
        sSQL = "select " & gsCAT_TCMR_SID & ", " & _
                           gsCAT_TCMR_EQID & ", " & _
                           gsCAT_TCMR_LOTID & ", " & _
                           gsCAT_TCMR_UPDATEUSERID & ", " & _
                           gsCAT_TCMR_HWFLAG & ", " & _
                           gsCAT_TCMR_PEFLAG & ", " & _
                           gsCAT_TCMR_MONITORITEMCONT & ", " & _
                           gsCAT_TCMR_EQSTOPTESTTYPE & ", " & _
                         " case " & _
                           " when " & gsCAT_TCMR_CREATETIME & " > " & _
                                " to_char(sysdate - " & sAutoValidTime & ", 'yyyymmdd hh24miss') || '000' then 'Y' Else 'N' End As ValidFlag " & _
                " from " & gsCAT_TBL_CPETEST_MONITOR_RECEIVE & " " & _
               " where " & gsCAT_TCMR_EQSTOPTEST & " = 'Y' " & _
                 " and " & gsCAT_TCMR_EQSTOPTESTTYPE & " = '" & gsSTOPTEST_TYPE_AUTO & "' " & _
                 " and " & gsCAT_TCMR_REVIEWFLAG & " = 'S' " & _
                 " and " & gsCAT_TCMR_CREATETIME & " is not null " & _
                 " and " & gsCAT_TCMR_DELETEFLAG & " = 'N' "
    End If
    
    Set colRaws = moProRawSql.QueryDatabase(sSQL)

    For Each oRaws In colRaws
        sMsg = "" 'Add by Weilun on 20200204 for Req.BE MES Phase 41 - 手動停機修改, 系統背景執行SortLotEnd
        bOverReviewFlag = False 'Add by Weilun on 20200731 for ReqNo. BE#202000220
    
        sSID = oRaws.Item(gsCAT_TCMR_SID)
        'Add by Weilun on 20200206 for Req.BE MES Phase 41 - 手動停機修改, 系統背景執行SortLotEnd <start>
        sEqId = oRaws.Item(gsCAT_TCMR_EQID)
        sLotId = oRaws.Item(gsCAT_TCMR_LOTID)
        sReviewUserID = oRaws.Item(gsCAT_TCMR_UPDATEUSERID)
        'Add by Weilun on 20200206 for Req.BE MES Phase 41 - 手動停機修改, 系統背景執行SortLotEnd <end>
        sMonitorItemCont = oRaws.Item(gsCAT_TCMR_MONITORITEMCONT)
        sHWFlag = oRaws.Item(gsCAT_TCMR_HWFLAG) '在手動Case, Review後會改為N, 藉由查詢強制改Y, 並更新回去需要Review
        sPEFlag = oRaws.Item(gsCAT_TCMR_PEFLAG) '同HWFlag
        
        bValid = False
        If oRaws.Item("ValidFlag") = "Y" Then
            bValid = True
        End If
        
        '機台和Lot一定要存在 Add by Weilun on 20200204 for Req.BE MES Phase 41 - 手動停機修改, 系統背景執行SortLotEnd
        Set oLot = moFwWIP.LotById(sLotId)
        Set oEquipment = moFwMDL.EquipmentById(sEqId)
        If oLot Is Nothing Then
            sMsg = oRaws.Item(gsCAT_TCMR_EQSTOPTESTTYPE) & " stop test event with nonexistent Lot: " & sLotId
        ElseIf oEquipment Is Nothing Then
            sMsg = oRaws.Item(gsCAT_TCMR_EQSTOPTESTTYPE) & " stop test event with nonexistent Equipment: " & sEqId
        Else
            '辨識類別
            Select Case UCase(Trim(sMonitorItemCont))
                Case UCase(gsMONITOR_ITEM_CONT_AUTOSTOP)
                    sActionType = gsSTOPTEST_ACTION_WREPAIR
                Case UCase(gsMONITOR_ITEM_CONT_COMMMGM)
                    If sHWFlag = "Y" Then 'HW優先
                        sActionType = gsSTOPTEST_ACTION_WREPAIR_HW
                    ElseIf sPEFlag = "Y" Then
                        'Modify by Weilun on 20200204 for Req.BE MES Phase 41 - 手動停機修改, 系統背景執行SortLotEnd <start>
                        'sActionType = gsSTOPTEST_ACTION_ENG
    
                        '轉ENG必須為IDLE, 做完LotEnd後再去轉ENG
                        If oEquipment.State = gsEQPSTATUS_RUN Or oEquipment.State = gsEQPSTATUS_ERUN Then
                            sActionType = gsSTOPTEST_ACTION_LOT_END & "," & gsSTOPTEST_ACTION_ENG
                        Else
                            sActionType = gsSTOPTEST_ACTION_ENG
                        End If
                        
                        'Modify by Weilun on 20200204 for Req.BE MES Phase 41 - 手動停機修改, 系統背景執行SortLotEnd <end>
                    Else
                        sActionType = ""
                    End If
    
                Case Else
                    sActionType = ""
            End Select
        
            'Add by Weilun on 20200204 for Req.BE MES Phase 41 - 手動停機修改, 系統背景執行SortLotEnd <start>
            '判定異常的停機, 包含不明的ActionType以及超時
            If bValid = False Then '主要
                '訊息更改, Marked by Weilun on 20200512 for Req.BE MES Phase 41 - 手動停機修改, 系統背景執行SortLotEnd <start>
'                If IsNumeric(sReviewUserID) Then '是工號的情況, 而不是系統的時候(另外在改這段的時候, 二階Review根本沒寫Update資訊)
'                    sMsg = oRaws.Item(gsCAT_TCMR_EQSTOPTESTTYPE) & " stop test event handle over limit time. Review UserID: " & sReviewUserID & "."
'                Else
'                    sMsg = oRaws.Item(gsCAT_TCMR_EQSTOPTESTTYPE) & " stop test event handle over limit time."
'                End If
                'Marked by Weilun on 20200512 for Req.BE MES Phase 41 - 手動停機修改, 系統背景執行SortLotEnd <end>

                '訊息更改, Add by Weilun on 20200512 for Req.BE MES Phase 41 - 手動停機修改, 系統背景執行SortLotEnd <start>
                If oRaws.Item(gsCAT_TCMR_EQSTOPTESTTYPE) = gsSTOPTEST_TYPE_MANUAL Then
                    sMsg = "手動停機異常確認時間超過30分鐘，請馬上執行。" & vbNewLine & _
                           "Review time was over 30 minutes for manual stop test event, please review immediately."
                    '以後要追加[事件+ReviewUserID]功能
                    
                    bOverReviewFlag = True 'Add by Weilun on 20200731 for ReqNo. BE#202000220
                Else
                    sMsg = "自動停機異常確認時間超過5分鐘，請馬上執行。" & vbNewLine & _
                           "Review time was over 5 minutes for auto stop test event, please review immediately."
                End If
                'Add by Weilun on 20200512 for Req.BE MES Phase 41 - 手動停機修改, 系統背景執行SortLotEnd <end>
            ElseIf sActionType = "" Then
                sMsg = oRaws.Item(gsCAT_TCMR_EQSTOPTESTTYPE) & " stop test event can not assign actionType."
                
            '機台上的LotID與要LotEnd的不符(基本上包含LotID不存在)
            'Modify by Weilun on 20230606 for Project.新增CP監控識別虛擬併批功能以及修改預測時間機制
            'ElseIf sActionType Like gsSTOPTEST_ACTION_LOT_END & "*" And oEquipment.CustomAttributes(gsEQP_CUSTOMATTR_CUR_LOT_ID) <> sLotID Then
            ElseIf sActionType Like gsSTOPTEST_ACTION_LOT_END & "*" And GetStrPosInArray(Replace(oEquipment.CustomAttributes(gsEQP_CUSTOMATTR_CUR_LOT_ID), ";", ","), sLotId, ",") < 1 Then
                sMsg = oRaws.Item(gsCAT_TCMR_EQSTOPTESTTYPE) & " stop test event start LotEnd with wrong Lot."
            End If
            'Add by Weilun on 20200204 for Req.BE MES Phase 41 - 手動停機修改, 系統背景執行SortLotEnd <end>
        End If

        
        'Mark by Weilun on 20200204 for Req.BE MES Phase 41 - 手動停機修改, 系統背景執行SortLotEnd <start>
'        '判定異常的停機, 包含不明的ActionType以及超時
'        sMsg = ""
'        If bValid = False Then '主要
'            Call TransferLogTrace(DateTime.Now & ": StopTest exceeds time limit of SID: " & sSID, moAppLog, Me.rtfTrace)
'            sMsg = "Stop Test fail : exceeds time limit"
'        ElseIf sActionType = "" Then
'            Call TransferLogTrace(DateTime.Now & ": Can not assign StopTest ActionType for SID: " & sSID, moAppLog, Me.rtfTrace)
'            sMsg = "Stop Test fail : can not assign Stop Test ActionType"
'        End If
        'Mark by Weilun on 20200204 for Req.BE MES Phase 41 - 手動停機修改, 系統背景執行SortLotEnd <end>
        
        
        If sMsg = "" Then '正常停機
        
            '進行停機
            
            'Review Flag:
            'P: PEIS需要補資料, 補完前不會Show在二階
            'M: 手動停機
            'S: 停機初始化
            'T: 正在停機中(已接手)
            '(取消)N: 需要Review(成功 + 失敗(寫對應的MSG))
            'N: 需要Review(成功) 'Add by Weilun on 20200204 for Req.BE MES Phase 41 - 手動停機修改, 系統背景執行SortLotEnd
            'Y: 已Review完
            'F: 停機發生問題(寫對應的MSG) 'Add by Weilun on 20200204 for Req.BE MES Phase 41 - 手動停機修改, 系統背景執行SortLotEnd
            
            '在手動停機, 是前端ReviewFlag改為M時Trigger, 並在這裡改T, 使用plsql改時建議用SQL+commit一次下修改指令
            '保險起見還是延遲個一秒再改
            '從peflag改停機類型,Fix by Weilun on 20200203 for Req.BE MES Phase 41 - 手動停機修改, 系統背景執行SortLotEnd
            If oRaws.Item(gsCAT_TCMR_EQSTOPTESTTYPE) = gsSTOPTEST_TYPE_MANUAL Then
                Call TransferLogTrace(DateTime.Now & ": Manual Stop Test delay 1 second for SID: " & sSID, _
                              moAppLog, Me.rtfTrace)
                Sleep 1000
            End If
            
            Call TransferLogTrace(DateTime.Now & ": Trigger Stop Test with [" & sActionType & "] for SID: " & sSID, _
                                  moAppLog, Me.rtfTrace)
            
            
            '先改ReviewFlag, 表示正在做
            '由於停機那邊也會回頭修改Receive, 所以要改Flag也得先做(需要的話)
            sModifySQL = " update " & gsCAT_TBL_CPETEST_MONITOR_RECEIVE & " " & _
                            " set " & gsCAT_TCMR_REVIEWFLAG & " = '" & "T" & "', " & _
                                      gsCAT_TCMR_HWFLAG & " = '" & sHWFlag & "', " & _
                                      gsCAT_TCMR_PEFLAG & " = '" & sPEFlag & "', " & _
                                      gsCAT_TCMR_UPDATEUSERID & " = 'SYS', " & _
                                      gsCAT_TCMR_UPDATETIME & " = to_char(sysdate, 'yyyymmdd hh24miss') || '000' " & _
                          " where " & gsCAT_TCMR_SID & " = " & sSID & " " & _
                            " and " & gsCAT_TCMR_REVIEWFLAG & " = 'S' " & _
                            " and " & gsCAT_TCMR_DELETEFLAG & " = 'N' "
            Call moProRawSql.QueryDatabase(sModifySQL)
            
            '然後再寫入停機
            sModifySQL = " insert into " & gsCAT_TBL_CPETEST_MONITOR_OUTPUT & " " & _
                              " (" & gsCAT_TCMO_ORISID & ", " & _
                                     gsCAT_TCMO_ACTIONTYPE & ", " & _
                                     gsCAT_TCMO_CREATETIME & ", " & _
                                     gsCAT_TCMO_CREATEUSERID & ") " & _
                   " values ( " & sSID & ", " & _
                            " '" & sActionType & "', " & _
                            " to_char(sysdate, 'yyyymmdd hh24miss') || '000', " & _
                            " 'SYS')"
            Call moProRawSql.QueryDatabase(sModifySQL)
        Else '不可停機
            'Add by Weilun on 20200203 for Req.BE MES Phase 41 - 手動停機修改, 系統背景執行SortLotEnd <start>
            Call TransferLogTrace(DateTime.Now & ": " & sMsg & " SID: " & sSID, moAppLog, Me.rtfTrace)
            
            Call InsertStopTestFail(sSID, _
                                    sMsg, _
                                    "1", _
                                    gsCAT_TCMR_PRODTAFLAG)
            
            'Add by Weilun on 20200731 for ReqNo.BE#202000220 <start>
            sExtraSQL = ""
            If bOverReviewFlag = True Then
                sExtraSQL = gsCAT_TCMR_OVERREVIEWFLAG & " = 'Y', "
            End If
            'Add by Weilun on 20200731 for ReqNo.BE#202000220 <end>
                                      
            '添加sExtraSQL, Modify by Weilun on 20200731 for ReqNo.BE#202000220
            sModifySQL = " update " & gsCAT_TBL_CPETEST_MONITOR_RECEIVE & " " & _
                            " set " & gsCAT_TCMR_REVIEWFLAG & " = '" & "F" & "', " & _
                                      gsCAT_TCMR_HWFLAG & " = '" & sHWFlag & "', " & _
                                      gsCAT_TCMR_PEFLAG & " = '" & sPEFlag & "', " & _
                                      gsCAT_TCMR_UPDATEUSERID & " = 'SYS', " & _
                                      sExtraSQL & _
                                      gsCAT_TCMR_UPDATETIME & " = to_char(sysdate, 'yyyymmdd hh24miss') || '000' " & _
                          " where " & gsCAT_TCMR_SID & " = " & sSID & " " & _
                            " and " & gsCAT_TCMR_REVIEWFLAG & " = 'S' " & _
                            " and " & gsCAT_TCMR_DELETEFLAG & " = 'N' "
            Call moProRawSql.QueryDatabase(sModifySQL)
                                      
            'Add by Weilun on 20200203 for Req.BE MES Phase 41 - 手動停機修改, 系統背景執行SortLotEnd <end>
            
            
            'Mark by Weilun on 20200203 for Req.BE MES Phase 41 - 手動停機修改, 系統背景執行SortLotEnd <start>
'            sMsgSQL = ""
'            If sHWFlag = "Y" Then
'                sMsgSQL = sMsgSQL & gsCAT_TCMR_HWMSG & " = '" & sMsg & "', "
'            End If
'            If sPEFlag = "Y" Then
'                sMsgSQL = sMsgSQL & gsCAT_TCMR_PEMSG & " = '" & sMsg & "', "
'            End If
'
'            sModifySQL = " update " & gsCAT_TBL_CPETEST_MONITOR_RECEIVE & " " & _
'                            " set " & gsCAT_TCMR_REVIEWFLAG & " = '" & "N" & "', " & _
'                                      gsCAT_TCMR_HWFLAG & " = '" & sHWFlag & "', " & _
'                                      gsCAT_TCMR_PEFLAG & " = '" & sPEFlag & "', " & _
'                                      sMsgSQL & _
'                                      gsCAT_TCMR_UPDATEUSERID & " = 'SYS', " & _
'                                      gsCAT_TCMR_UPDATETIME & " = to_char(sysdate, 'yyyymmdd hh24miss') || '000' " & _
'                          " where " & gsCAT_TCMR_SID & " = " & sSID & " " & _
'                            " and " & gsCAT_TCMR_REVIEWFLAG & " = 'S' " & _
'                            " and " & gsCAT_TCMR_DELETEFLAG & " = 'N' "
'            Call moProRawSql.QueryDatabase(sModifySQL)
            'Mark by Weilun on 20200203 for Req.BE MES Phase 41 - 手動停機修改, 系統背景執行SortLotEnd <end>
            
        End If
    Next

    '-------
    'Done
    '-------
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
            Case Else
                typErrInfo.sUserText = typErrInfo.sErrDescription
        End Select
        '---- Start of Select Case Block ----
        On Error GoTo ExitHandler
        Call TransferLogError(typErrInfo.sUserText, _
                              moAppLog, Me.rtfTrace, msMODULE_ID, sProcID)
        Call HandleError(False, typErrInfo, , moAppLog)
    End If
End Sub

'================================================================================
'Name       : InsertStopTestFail()
'Description: 停機發生錯誤時, 複製原有的二階資訊新增一筆異常二階資訊提供User Review Error Msg
'Author     : Weilun huang, MXIC 2020/02/07
'Parameters : N/A
'================================================================================
Private Sub InsertStopTestFail(ByVal sSID As String, _
                               ByVal sErrorMsg As String, _
                               ByVal sAlarmLevel As String, _
                               ByVal sReviewFlag As String)
    On Error GoTo ExitHandler
    Dim sProcID As String
    Dim typErrInfo As tErrInfo

    Dim sSQL As String
    Dim colRaws As Collection

    '----
    ' Init
    '----
    sProcID = "InsertStopTestFail"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog)

    '----
    ' Condition Chking
    '----
    ' <Put your condition checking codes here>...
    If Trim(sSID) = "" Then
        GoTo ExitHandler
    End If

    If sAlarmLevel = "" Then
        sAlarmLevel = "1"
    End If

    If Not (sReviewFlag = gsCAT_TCMR_PRODTAFLAG Or _
            sReviewFlag = gsCAT_TCMR_HWFLAG Or _
            sReviewFlag = gsCAT_TCMR_PEFLAG) Then
        GoTo ExitHandler
    End If
    '----
    ' Action
    '----
    sSQL = " insert into " & gsCAT_TBL_CPETEST_MONITOR_RECEIVE & " " & _
                      " (" & gsCAT_TCMR_SID & ", " & _
                             gsCAT_TCMR_EQID & ", " & _
                             gsCAT_TCMR_LOTID & ", " & _
                             gsCAT_TCMR_LOTSEQ & ", " & _
                             gsCAT_TCMR_WAFERID & ", " & _
                             gsCAT_TCMR_STEPNAME & ", " & _
                             gsCAT_TCMR_NEXTSTEPNAME & ", " & _
                             gsCAT_TCMR_RECIPE & ", " & _
                             gsCAT_TCMR_HOTFLAG & ", " & _
                             gsCAT_TCMR_MONITORITEM & ", " & _
                             gsCAT_TCMR_MONITORITEMCONT & ", " & _
                             gsCAT_TCMR_MONITORITEMCONTDESC & ", " & _
                             gsCAT_TCMR_ALARMLEVEL & ", " & _
                             sReviewFlag & ", " & _
                             gsCAT_TCMR_CREATETIME & ", " & _
                             gsCAT_TCMR_CREATEUSERID & ") "
    sSQL = sSQL & _
           " select CPETEST_MONIOTR_SEQ.NEXTVAL, " & _
                    gsCAT_TCMR_EQID & ", " & _
                    gsCAT_TCMR_LOTID & ", " & _
                    gsCAT_TCMR_LOTSEQ & ", " & _
                    gsCAT_TCMR_WAFERID & ", " & _
                    gsCAT_TCMR_STEPNAME & ", " & _
                    gsCAT_TCMR_NEXTSTEPNAME & ", " & _
                    gsCAT_TCMR_RECIPE & ", " & _
                    gsCAT_TCMR_HOTFLAG & ", " & _
                    gsCAT_TCMR_MONITORITEM & ", " & _
                    gsCAT_TCMR_MONITORITEMCONT & ", " & _
                    " '" & sErrorMsg & "', " & _
                    " " & sAlarmLevel & ", " & _
                    " 'Y', " & _
                    " to_char(sysdate, 'yyyymmdd hh24miss') || '000', " & _
                    " 'SYS' "
    sSQL = sSQL & _
           " from " & gsCAT_TBL_CPETEST_MONITOR_RECEIVE & " " & _
           " where " & gsCAT_TCMR_SID & " = '" & sSID & "' "
    Call moProRawSql.QueryDatabase(sSQL)

    '-------
    'Done
    '-------
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
            Case Else
                typErrInfo.sUserText = typErrInfo.sErrDescription
        End Select
        '---- Start of Select Case Block ----
        On Error GoTo ExitHandler
        Call TransferLogError(typErrInfo.sUserText, _
                              moAppLog, Me.rtfTrace, msMODULE_ID, sProcID)
        Call HandleError(False, typErrInfo, , moAppLog)
    End If
End Sub

'================================================================================
'Name       : Level1Sort()
'Description: 將二階資料整理為一階資料
'Author     : Weilun huang, MXIC 2019/09/04
'Parameters : N/A
'================================================================================
Private Sub Level1Sort(Optional ByVal sInputEqID As String = "", _
                       Optional ByVal sReviewSID As String = "", _
                       Optional ByVal sReviewFlag As String = "")
    On Error GoTo ExitHandler
    Dim sProcID As String
    Dim typErrInfo As tErrInfo

    Dim sSQL As String
    Dim colRaws As Collection
    Dim oRaws As Collection
    Dim sModifySQL As String

    Dim sEqIDSQL            As String
    Dim sIgnoreSIDSQL       As String

    Dim sEqId               As String
    Dim sEqStatus           As String
    Dim sAlarmLevel         As String
    Dim sHWFlag             As String
    Dim sPEFlag             As String
    Dim sProdTAFlag         As String
    Dim sLotEndFlag         As String 'Add by Weilun on 20211126 for Project.BE 工業 3.5 Phase 39 - CP AGV導入專案 - 建立 AGV 派工 & Status 監控系統
    Dim sEqStopTest         As String
    Dim sEqStopTestType     As String
    Dim sMonitorItem        As String
    Dim sUpdateTime         As String
    '-----
    ' Init
    '-----
    sProcID = "Level1Sort"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog) '

    '--------------------
    ' Condition Checking
    '--------------------

    sEqIDSQL = ""
    If sInputEqID <> "" Then 'Level1和Receive都是用eqid
        sEqIDSQL = " and " & gsCAT_TCMR_EQID & " = '" & sInputEqID & "' "
    End If

    If sReviewSID <> "" And IsNumeric(sReviewSID) And sReviewFlag = "Y" Then
        sIgnoreSIDSQL = " and " & gsCAT_TCMR_SID & " != " & sReviewSID & " "
    End If

    '-------
    ' Action
    '-------
    '主輸出部分
    'LOTENDFLAG (整個sSQL), Add by Weilun on 20211126 for Project.BE 工業 3.5 Phase 39 - CP AGV導入專案 - 建立 AGV 派工 & Status 監控系統
    sSQL = "select t." & gsCAT_TCML1_EQID & ", " & _
                 " t." & gsCAT_TCML1_EQSTATUS & ", " & _
                 " f1." & gsCAT_TCMR_ALARMLEVEL & ", " & _
                 " f1." & gsCAT_TCMR_HWFLAG & ", " & _
                 " f1." & gsCAT_TCMR_PEFLAG & ", " & _
                 " f1." & gsCAT_TCMR_PRODTAFLAG & ", " & _
                 " f1." & gsCAT_TCMR_LOTENDFLAG & ", " & _
                 " f1." & gsCAT_TCMR_EQSTOPTEST & ", " & _
                 " f1.ReceiveTime, " & _
                 " c3." & gsCAT_TCMR_EQSTOPTESTTYPE & ", " & _
                 " m3." & gsCAT_TCMR_MONITORITEM & " " & _
            " from "

    '從Level1取得機台狀態, 如果DeleteFlag為Y, 表示機台已不存在, 將不會被整理出來
    sSQL = sSQL & _
                  "(select " & gsCAT_TCML1_EQID & ", " & gsCAT_TCML1_EQSTATUS & " " & _
                    " from " & gsCAT_TBL_CPETEST_MONITOR_LEVEL1 & " " & _
                   " where " & gsCAT_TCML1_DELETEFLAG & " = 'N' " & _
                               sEqIDSQL & _
                 " ) t, "

    '取Recive的AlarmLevel, 並將Flag代換成數字搭配max進行標記
    'Modify by Weilun on 20200717 for ReqNo.BE#202000220 <start>
    'CreateTime/UpdateTime由max改min, 規則也由CT>UT改CT<UT, nvl(time,' ')改成nvl(time,'99999999 999999999')
    '事件計算時間也改成看AlarmLevel來確定是否要顯示(" Case when " & gsCAT_TCMR_ALARMLEVEL & " in (1) then gsCAT_TCMR_CREATETIME else '')
    'Modify by Weilun on 20200717 for ReqNo.BE#202000220 <end>
    sSQL = sSQL & _
                  "(select " & gsCAT_TCMR_EQID & ", " & _
                             " min(" & gsCAT_TCMR_ALARMLEVEL & ") as " & gsCAT_TCMR_ALARMLEVEL & ", " & _
                             " max(" & gsCAT_TCMR_HWFLAG & ") as " & gsCAT_TCMR_HWFLAG & ", " & _
                             " max(" & gsCAT_TCMR_PEFLAG & ") as " & gsCAT_TCMR_PEFLAG & ", " & _
                             " max(" & gsCAT_TCMR_PRODTAFLAG & ") as " & gsCAT_TCMR_PRODTAFLAG & ", " & _
                             " max(" & gsCAT_TCMR_LOTENDFLAG & ") as " & gsCAT_TCMR_LOTENDFLAG & ", " & _
                             " max(" & gsCAT_TCMR_EQSTOPTEST & ") as " & gsCAT_TCMR_EQSTOPTEST & ", " & _
                             " Case when min(nvl(" & gsCAT_TCMR_CREATETIME & ", '99999999 999999999')) < min(nvl(" & gsCAT_TCMR_UPDATETIME & ", '99999999 999999999')) then min(nvl(" & gsCAT_TCMR_CREATETIME & ", '99999999 999999999')) else min(nvl(" & gsCAT_TCMR_UPDATETIME & ", '99999999 999999999')) end as ReceiveTime "
    sSQL = sSQL & _
                    " from (select " & gsCAT_TCMR_EQID & ", " & _
                                       gsCAT_TCMR_ALARMLEVEL & ", " & _
                                     " Case when " & gsCAT_TCMR_HWFLAG & " = 'Y' then 1 else 0 End as " & gsCAT_TCMR_HWFLAG & ", " & _
                                     " Case when " & gsCAT_TCMR_PEFLAG & " = 'Y' then 1 else 0 End as " & gsCAT_TCMR_PEFLAG & ", " & _
                                     " Case when " & gsCAT_TCMR_PRODTAFLAG & " = 'Y' then 1 else 0 End as " & gsCAT_TCMR_PRODTAFLAG & ", " & _
                                     " Case when " & gsCAT_TCMR_LOTENDFLAG & " = 'Y' then 1 else 0 End as " & gsCAT_TCMR_LOTENDFLAG & ", " & _
                                     " Case when " & gsCAT_TCMR_EQSTOPTEST & " = 'Y' then 2 when " & gsCAT_TCMR_EQSTOPTEST & " = 'N' then 1 else 0 End as " & gsCAT_TCMR_EQSTOPTEST & ", " & _
                                     " Case when " & gsCAT_TCMR_ALARMLEVEL & " in (1) then " & gsCAT_TCMR_CREATETIME & " else '' End as " & gsCAT_TCMR_CREATETIME & ", " & _
                                     " Case when " & gsCAT_TCMR_ALARMLEVEL & " in (1) then " & gsCAT_TCMR_UPDATETIME & " else '' End as " & gsCAT_TCMR_UPDATETIME & " " & _
                            " from " & gsCAT_TBL_CPETEST_MONITOR_RECEIVE & " t " & _
                           " where ((" & gsCAT_TCMR_ALARMLEVEL & " = 99 and " & gsCAT_TCMR_REVIEWFLAG & " = 'N') or " & _
                                  " (" & gsCAT_TCMR_ALARMLEVEL & " = 1 and " & gsCAT_TCMR_REVIEWFLAG & " = 'M') or " & _
                                  " (" & gsCAT_TCMR_ALARMLEVEL & " = 1 and " & gsCAT_TCMR_REVIEWFLAG & " = 'N')) " & _
                                       sEqIDSQL & _
                                       sIgnoreSIDSQL & _
                             " and " & gsCAT_TCMR_DELETEFLAG & " = 'N') " & _
                " group by " & gsCAT_TCMR_EQID & ") f1, "

    '依機台將Receive的EqStopTestType串起來
    sSQL = sSQL & _
                  " (select c2." & gsCAT_TCMR_EQID & ", " & _
                          " substr(max(sys_connect_by_path(" & gsCAT_TCMR_EQSTOPTESTTYPE & ", ',')), 2) as " & gsCAT_TCMR_EQSTOPTESTTYPE & " " & _
                     " from (select c1.*, " & _
                                  " row_number() over(partition by c1." & gsCAT_TCMR_EQID & " order by c1." & gsCAT_TCMR_EQID & ") rnum " & _
                             " from (select distinct " & gsCAT_TCMR_EQID & ", " & gsCAT_TCMR_EQSTOPTESTTYPE & " " & _
                                     " from " & gsCAT_TBL_CPETEST_MONITOR_RECEIVE & " " & _
                                    " where ((" & gsCAT_TCMR_ALARMLEVEL & " = 99 and " & gsCAT_TCMR_REVIEWFLAG & " = 'N') or " & _
                                           " (" & gsCAT_TCMR_ALARMLEVEL & " = 1 and " & gsCAT_TCMR_REVIEWFLAG & " = 'M') or " & _
                                           " (" & gsCAT_TCMR_ALARMLEVEL & " = 1 and " & gsCAT_TCMR_REVIEWFLAG & " = 'N')) " & _
                                      " and " & gsCAT_TCMR_EQSTOPTESTTYPE & " is not null " & _
                                                sEqIDSQL & _
                                                sIgnoreSIDSQL & _
                                      " and " & gsCAT_TCMR_DELETEFLAG & " = 'N') c1) c2 " & _
                    " start with rnum = 1 " & _
                   " connect by prior rnum = rnum - 1 " & _
                          " and prior " & gsCAT_TCMR_EQID & " = " & gsCAT_TCMR_EQID & " " & _
                 " group by c2." & gsCAT_TCMR_EQID & ") c3, "

    '依機台將Receive的MonitorItem串起來(PS. 該功能為處要求, 但是SI和User都不想要)
    sSQL = sSQL & _
                  " (select m2." & gsCAT_TCMR_EQID & ", " & _
                          " substr(max(sys_connect_by_path(" & gsCAT_TCMR_MONITORITEM & ", ',')), 2) as " & gsCAT_TCMR_MONITORITEM & " " & _
                     " from (select m1.*, " & _
                                  " row_number() over(partition by m1." & gsCAT_TCMR_EQID & " order by m1." & gsCAT_TCMR_EQID & ") rnum " & _
                             " from (select distinct " & gsCAT_TCMR_EQID & ", " & gsCAT_TCMR_MONITORITEM & " " & _
                                     " from " & gsCAT_TBL_CPETEST_MONITOR_RECEIVE & " " & _
                                    " where ((" & gsCAT_TCMR_ALARMLEVEL & " = 99 and " & gsCAT_TCMR_REVIEWFLAG & " = 'N') or " & _
                                           " (" & gsCAT_TCMR_ALARMLEVEL & " = 1 and " & gsCAT_TCMR_REVIEWFLAG & " = 'M') or " & _
                                           " (" & gsCAT_TCMR_ALARMLEVEL & " = 1 and " & gsCAT_TCMR_REVIEWFLAG & " = 'N')) " & _
                                      " and " & gsCAT_TCMR_MONITORITEM & " is not null " & _
                                                sEqIDSQL & _
                                                sIgnoreSIDSQL & _
                                      " and " & gsCAT_TCMR_DELETEFLAG & " = 'N') m1) m2 " & _
                    " start with rnum = 1 " & _
                   " connect by prior rnum = rnum - 1 " & _
                          " and prior " & gsCAT_TCMR_EQID & " = " & gsCAT_TCMR_EQID & " " & _
                 " group by m2." & gsCAT_TCMR_EQID & ") m3 "

    '以Level1為主體, Receive都為附加
    sSQL = sSQL & _
           " where t." & gsCAT_TCML1_EQID & " = f1." & gsCAT_TCMR_EQID & "(+) " & _
             " and t." & gsCAT_TCML1_EQID & " = c3." & gsCAT_TCMR_EQID & "(+) " & _
             " and t." & gsCAT_TCML1_EQID & " = m3." & gsCAT_TCMR_EQID & "(+) " & _
        " order by t." & gsCAT_TCML1_EQID & " "

    Set colRaws = moProRawSql.QueryDatabase(sSQL)
    For Each oRaws In colRaws
    
        sEqId = oRaws.Item(gsCAT_TCML1_EQID) '由於Level1的EqID為必要值, 基本上不會是空的, 這裡就不再加一層if
        sEqStatus = oRaws.Item(gsCAT_TCML1_EQSTATUS)
        sUpdateTime = "'" & oRaws.Item("ReceiveTime") & "'"
        
        'Trim應該在Replace外, 並加上等於時間最大值的狀況, Modify by Weilun on 20200717 for ReqNo.BE#202000220
        'If Replace(Trim(sUpdateTime), "'", "") = "" Then
        If Trim(Replace(sUpdateTime, "'", "")) = "" Or _
           Trim(Replace(sUpdateTime, "'", "")) = "99999999 999999999" Then
            sUpdateTime = "to_char(sysdate, 'yyyymmdd hh24miss') || '000'"
        End If

        sAlarmLevel = "99" '預設值
        If Trim(oRaws.Item(gsCAT_TCMR_ALARMLEVEL)) <> "" And IsNumeric(oRaws.Item(gsCAT_TCMR_ALARMLEVEL)) = True Then
            sAlarmLevel = oRaws.Item(gsCAT_TCMR_ALARMLEVEL)
        End If

        'IDLE時至少黃底(2)
        If sEqStatus = gsEQPSTATUS_IDLE And sAlarmLevel > "2" Then
            sAlarmLevel = "2"
        End If

        sHWFlag = "" '預設值
        If oRaws.Item(gsCAT_TCMR_HWFLAG) = "1" Then
            sHWFlag = "Y"
        End If

        sPEFlag = "" '預設值
        If oRaws.Item(gsCAT_TCMR_PEFLAG) = "1" Then
            sPEFlag = "Y"
        End If

        sProdTAFlag = "" '預設值
        If oRaws.Item(gsCAT_TCMR_PRODTAFLAG) = "1" Then
            sProdTAFlag = "Y"
        End If
        
        'Add by Weilun on 20211126 for Project.BE 工業 3.5 Phase 39 - CP AGV導入專案 - 建立 AGV 派工 & Status 監控系統 <start>
        sLotEndFlag = "" '預設值
        If oRaws.Item(gsCAT_TCMR_LOTENDFLAG) = "1" Then
            sLotEndFlag = "Y"
        End If
        'Add by Weilun on 20211126 for Project.BE 工業 3.5 Phase 39 - CP AGV導入專案 - 建立 AGV 派工 & Status 監控系統 <end>
        
        sEqStopTest = "" '預設值
        If oRaws.Item(gsCAT_TCMR_EQSTOPTEST) = "2" Then
            sEqStopTest = "Y"
        ElseIf oRaws.Item(gsCAT_TCMR_EQSTOPTEST) = "1" Then
            sEqStopTest = "N"
        End If

        sEqStopTestType = Left(oRaws.Item(gsCAT_TCMR_EQSTOPTESTTYPE), 20) ' 依Level1欄位規格限制長度
        sMonitorItem = Left(oRaws.Item(gsCAT_TCMR_MONITORITEM), 1024)

        '更新相關欄位
        'LOTENDFLAG, Add by Weilun on 20211126 for Project.BE 工業 3.5 Phase 39 - CP AGV導入專案 - 建立 AGV 派工 & Status 監控系統
        sModifySQL = " update " & gsCAT_TBL_CPETEST_MONITOR_LEVEL1 & " " & _
                        " set " & gsCAT_TCML1_ALARMLEVEL & " = " & sAlarmLevel & ", " & _
                                  gsCAT_TCML1_HWFLAG & " = '" & sHWFlag & "', " & _
                                  gsCAT_TCML1_PEFLAG & " = '" & sPEFlag & "', " & _
                                  gsCAT_TCML1_PRODTAFLAG & " = '" & sProdTAFlag & "', " & _
                                  gsCAT_TCML1_LOTENDFLAG & " = '" & sLotEndFlag & "', " & _
                                  gsCAT_TCML1_EQSTOPTEST & " = '" & sEqStopTest & "', " & _
                                  gsCAT_TCML1_EQSTOPTESTTYPE & " = '" & sEqStopTestType & "', " & _
                                  gsCAT_TCML1_MONITORITEM & " = '" & sMonitorItem & "', " & _
                                  gsCAT_TCML1_UPDATEUSERID & " = 'SYS', " & _
                                  gsCAT_TCML1_UPDATETIME & " = " & sUpdateTime & " " & _
                      " where " & gsCAT_TCML1_EQID & " = '" & sEqId & "' " & _
                        " and " & gsCAT_TCML1_DELETEFLAG & " = 'N' "
        Call moProRawSql.QueryDatabase(sModifySQL)
    Next

    '-------
    'Done
    '-------
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
            Case Else
                typErrInfo.sUserText = typErrInfo.sErrDescription
        End Select
        '---- Start of Select Case Block ----
        On Error GoTo ExitHandler
        Call TransferLogError(typErrInfo.sUserText, _
                              moAppLog, Me.rtfTrace, msMODULE_ID, sProcID)
        Call HandleError(False, typErrInfo, , moAppLog)
    End If
End Sub

'================================================================================
'Name       : EqStateChnage()
'Description: 機台狀態變動更新Level1資料, 也用於初始化
'Author     : Weilun huang, MXIC 2019/09/04
'Parameters : N/A
'================================================================================
Private Sub EqStateChnage(ByVal sReciveMsg As String)
    On Error GoTo ExitHandler:
    Dim sProcID                     As String
    Dim typErrInfo                  As tErrInfo
    
    Dim asMsg()                     As String
    Dim sEventType                  As String
    Dim sEqId                       As String
    Dim sEqState                    As String
    
    Dim sOldLotID                   As String
    Dim sLotId                      As String
    Dim sIPN                        As String
    Dim sStepName                   As String
    Dim sWaferQty                   As String
    Dim dUph                        As Double
    Dim lWaitTest                   As Long
    Dim lUnTestedWQty               As Long 'Add by Weilun on 20200723 for ReqNo.BE#202000220
    'Add by Weilun on 20211126 for Project.BE 工業 3.5 Phase 39 - CP AGV導入專案 - 建立 AGV 派工 & Status 監控系統 <start>
    Dim sUnloadType                 As String
    Dim sUnloadTypeSQL              As String
    'Add by Weilun on 20211126 for Project.BE 工業 3.5 Phase 39 - CP AGV導入專案 - 建立 AGV 派工 & Status 監控系統 <end>
    Dim sEngDept As String 'Add by Weilun on 20230512 for ReqNo.BE#202300059
    
    Dim sStartTime                  As String
    Dim sExpectEndTime              As String
    Dim sRecentEndTime              As String
    
    Dim sSQL                        As String
    Dim colRaws                     As Collection
    Dim oRaws                       As Collection

    Dim oLot                        As FwLot
    Dim oEquipment                  As FwEquipment
    
    'Add by Weilun on 20230606 for Project.新增CP監控識別虛擬併批功能以及修改預測時間機制 <start>
    Dim sShowLotId                  As String
    Dim sEqpLotIDs                  As String
    Dim bIsVirtualMerge             As Boolean
    'Add by Weilun on 20230606 for Project.新增CP監控識別虛擬併批功能以及修改預測時間機制 <end>
    
    '-----
    ' Init
    '-----
    sProcID = "EqStateChnage"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog) '
    
    asMsg = Split(sReciveMsg, "/")
    '"-"
    sEventType = asMsg(0)
    sEqId = asMsg(1)
    sEqState = asMsg(2)
    
    '--------------------
    ' Condition Checking
    '--------------------
    
    '-------
    ' Action
    '-------


    Set oEquipment = moFwMDL.EquipmentById(sEqId)
    
    If oEquipment.Type = "N2" Then
        Debug.Print oEquipment.Id
    End If
    
    If oEquipment Is Nothing Then
        Call TransferLogError("該Eq不存在.", _
                              moAppLog, Me.rtfTrace, msMODULE_ID, sProcID)
        GoTo ExitHandler
    '特定機台不要 'Add by Weilun on 20200203 for Req.BE MES Phase 41 - 手動停機修改, 系統背景執行SortLotEnd <start>
    '在EqMonitorDataCheck外多攔一層, 省去Init時的一些誤會
    ElseIf oEquipment.CustomAttributes(gsEQP_CUSTOMATTR_EQ_TYPE2) = "WTESTER(DUMMY)" Then
        '這裡不顯示訊息
        GoTo ExitHandler
    '特定機台不要 'Add by Weilun on 20200203 for Req.BE MES Phase 41 - 手動停機修改, 系統背景執行SortLotEnd <end>
    'Add by Weilun on 20230103 for Req.SorterControl_Monitor <start>
    'WSORTER只要zWSORTER就好
    ElseIf oEquipment.Type = "WSORTER" And oEquipment.Id <> "zWSORTER" Then
        '這裡不顯示訊息
        GoTo ExitHandler
    'Add by Weilun on 20230103 for Req.SorterControl_Monitor <end>
''    'Add by Weilun on 20230717 for Req.當 N2櫃異常時, 顯示在 CP 監控系統 <start>
''    'zN2_BOX不顯示
''    ElseIf oEquipment.Type = "N2_BOX" And oEquipment.Id = "zN2_BOX" Then
''        '這裡不顯示訊息
''        GoTo ExitHandler
''    'Add by Weilun on 20230717 for Req.SorterControl_Monitor <end>
    Else
        
        Call TransferLogTrace(DateTime.Now & ": Refresh EqID: " & sEqId, _
                              moAppLog, Me.rtfTrace)
        
        Call EqMonitorDataCheck(oEquipment)
        '取得相關資料
        sStartTime = "''" '清空
        sExpectEndTime = "''" '清空
        sRecentEndTime = "''" '清空
        'Add by Weilun on 20211126 for Project.BE 工業 3.5 Phase 39 - CP AGV導入專案 - 建立 AGV 派工 & Status 監控系統 <start>
        sUnloadType = ""
        sUnloadTypeSQL = ""
        'Add by Weilun on 20211126 for Project.BE 工業 3.5 Phase 39 - CP AGV導入專案 - 建立 AGV 派工 & Status 監控系統 <end>
        
        '取得舊的LotID
        sSQL = " select " & gsCAT_TCML1_LOTID & _
                 " from " & gsCAT_TBL_CPETEST_MONITOR_LEVEL1 & " " & _
                " where " & gsCAT_TCML1_EQID & " = '" & oEquipment.Id & "' " & _
                  " and " & gsCAT_TCML1_DELETEFLAG & " = 'N' "
        Set colRaws = moProRawSql.QueryDatabase(sSQL)
        If colRaws.Count > 0 Then
            sOldLotID = colRaws.Item(1).Item(gsCAT_TCML1_LOTID)
        End If
        
        '取得現在LotID
        'Modify by Weilun on 20230606 for Project.新增CP監控識別虛擬併批功能以及修改預測時間機制
'        sLotID = oEquipment.CustomAttributes(gsEQP_CUSTOMATTR_CUR_LOT_ID)
        Call GetEqDetailLotList(oEquipment.Id, sLotId, sShowLotId, sEqpLotIDs, bIsVirtualMerge, "")
        Set oLot = moFwWIP.LotById(sLotId)
        If oLot Is Nothing Then
            sLotId = ""
        Else
            sIPN = oLot.CustomAttributes(gsLOT_CUSTOMATTR_IPN)
            sStepName = oLot.CurrentStep.Steps.Item(1).Description
            sWaferQty = oLot.CustomAttributes(gsLOT_CUSTOMATTR_WAFER_QTY)
            
            '轉Run或Erun時需要時間資訊, 如果是初始化遇到已經在跑的, 會試著重算時間
            '時間與UPH相關都改為"WTESTER"限定, Modify by Weilun on 20200512 for Project.BE MES Phase 44 – CP 產線增加自動化作業
            'If sEqState = gsEQPSTATUS_RUN Or sEqState = gsEQPSTATUS_ERUN Then
            If oEquipment.Type = "WTESTER" And _
               (sEqState = gsEQPSTATUS_RUN Or sEqState = gsEQPSTATUS_ERUN) Then
                '寫入起測與預計結束的時間
                dUph = GetUph(oLot, oEquipment, "WS", lWaitTest)
                If sEventType = gsMONITOR_TYPE_INIT Then '初始化的狀況
                    'Modify by Weilun on 20230606 for Project.新增CP監控識別虛擬併批功能以及修改預測時間機制
'                    If sOldLotID = sLotID Then 'LotID沒有變化時繼承時間欄位
                    If sOldLotID = sShowLotId Then 'LotID沒有變化時繼承時間欄位
                        '初始化的狀況, 這幾個欄位就維持
                        sStartTime = gsCAT_TCML1_STARTTIME
                        sRecentEndTime = gsCAT_TCML1_RECENTENDTIME
                            
                        '預計結束的時間, 由WaferEndTime往上加(監控掛太久可能導致不正確)
                        If dUph = 0 Then
                            sExpectEndTime = "''"
                        Else
                            '時間計算取消1.3倍, 並使用lUnTestedWQty, Modify by Weilun on 20200723 for ReqNo.BE#202000220 <start>
                            'sExpectEndTime = "to_char(to_date(SubStr(" & gsCAT_TCML1_RECENTENDTIME & ", 1, 15), " & _
                            '                                " 'YYYYMMDD HH24MISS') + (1*" & lWaitTest & "/" & dUph & "/24 * 1.3), " & _
                            '                        " 'YYYYMMDD HH24MISS') || SubStr(" & gsCAT_TCML1_RECENTENDTIME & ", 16, 3) " '後面那個是取000, 可相容全空白
                                                    
                            'Modify by Weilun on 20230606 for Project.新增CP監控識別虛擬併批功能以及修改預測時間機制
'                            lUnTestedWQty = GetUnTestedWQty(sLotId, sEqId, "", sStepName, "")

                            lUnTestedWQty = GetUnTestedTotalWQty(sEqpLotIDs, sEqId, sStepName)
                            sExpectEndTime = "to_char(to_date(SubStr(" & gsCAT_TCML1_RECENTENDTIME & ", 1, 15), " & _
                                                            " 'YYYYMMDD HH24MISS') + (1*" & lUnTestedWQty & "/" & dUph & "/24), " & _
                                                    " 'YYYYMMDD HH24MISS') || SubStr(" & gsCAT_TCML1_RECENTENDTIME & ", 16, 3) " '後面那個是取000, 可相容全空白
                            'Modify by Weilun on 20200723 for ReqNo.BE#202000220 <end>
                        End If
                    Else
                        '清空(預設), Lot更換, 舊的資料不能用, 也不能重新算時間
                    End If
                Else
                    '正常EqChange的狀況
                    
                    sStartTime = "to_char(sysdate, 'yyyymmdd hh24miss') || '000'"
                    sRecentEndTime = sStartTime
                    If dUph = 0 Then
                        sExpectEndTime = "''"
                    Else
                        '時間計算取消1.3倍, 並使用lUnTestedWQty, Modify by Weilun on 20200723 for ReqNo.BE#202000220 <start>
                        'sExpectEndTime = "to_char(sysdate + (1*" & lWaitTest & "/" & dUph & "/24 * 1.3), 'yyyymmdd hh24miss') || '000'"
                        
                        'Modify by Weilun on 20230606 for Project.新增CP監控識別虛擬併批功能以及修改預測時間機制
'                        lUnTestedWQty = GetUnTestedWQty(sLotId, sEqId, "", sStepName, "")
                        lUnTestedWQty = GetUnTestedTotalWQty(sEqpLotIDs, sEqId, sStepName)
                        sExpectEndTime = "to_char(sysdate + (1*" & lUnTestedWQty & "/" & dUph & "/24), 'yyyymmdd hh24miss') || '000'"
                        'Modify by Weilun on 20200723 for ReqNo.BE#202000220 <end>
                        
                        '寫AGV下貨資料與時間 'Add by Weilun on 20230606 for Project.新增CP監控識別虛擬併批功能以及修改預測時間機制
                        Call UpdateAGV_EqStatusUnloadInfo(sEqId, sLotId, sExpectEndTime) '請不要傳入虛擬併批母批
                    End If
                End If
                
                'Add by Weilun on 20211126 for Project.BE 工業 3.5 Phase 39 - CP AGV導入專案 - 建立 AGV 派工 & Status 監控系統 <start>
                '取得下貨方式
                sUnloadType = GetUnloadType(sEqId, sLotId)
                sUnloadTypeSQL = gsCAT_TCML1_UNLOADTYPE & " = '" & sUnloadType & "', "
                'Add by Weilun on 20211126 for Project.BE 工業 3.5 Phase 39 - CP AGV導入專案 - 建立 AGV 派工 & Status 監控系統 <end>
            End If
        End If

        'Add by Weilun on 20230512 for ReqNo.BE#202300059 <start>
        sEngDept = ""
        If oEquipment.Type = "WTESTER" And sEqState = gsEQPSTATUS_ENG Then
            sSQL = " select A.TIME, " & _
                          " A.STATE, " & _
                          " A.EVENTNAME, " & _
                          " to_char(B.DETAILDESCRIPTION) As DETAILDESCRIPTION " & _
                     " from FWEQPHISTORY A, FWEQPCOMMENT B " & _
                    " where (A.TIME) in (select max(TIME) " & _
                                         " from FwEqpHistory " & _
                                        " where EQUIPMENTNAME = '" & sEqId & "') " & _
                      " and A.STATE = '" & gsEQPSTATUS_ENG & "' " & _
                      " and A.EQUIPMENTNAME = '" & sEqId & "' " & _
                      " and A.EQPCOMMENT = B.SYSID(+) "
            Set colRaws = moProRawSql.QueryDatabase(sSQL)
            If colRaws.Count > 0 Then '最後一筆是ENG時才有資料
                sEngDept = Left(colRaws.Item(1).Item("DETAILDESCRIPTION"), 5) '部門僅5碼
            End If
        End If
        'Add by Weilun on 20230512 for ReqNo.BE#202300059 <end>

        '更新現況
        'Add sUnloadTypeSQL(有才更新) by Weilun on 20211126 for Project.BE 工業 3.5 Phase 39 - CP AGV導入專案 - 建立 AGV 派工 & Status 監控系統
        'Add sEngDept(有才更新) by Weilun on 20211126 for ReqNo.BE#202300059
        sSQL = " update " & gsCAT_TBL_CPETEST_MONITOR_LEVEL1 & " " & _
                  " set " & gsCAT_TCML1_EQSTATUS & " = '" & sEqState & "', " & _
                            gsCAT_TCML1_LOTID & " = '" & sShowLotId & "', " & _
                            gsCAT_TCML1_IPN & " = '" & sIPN & "', " & _
                            gsCAT_TCML1_STEPNAME & " = '" & sStepName & "', " & _
                            sUnloadTypeSQL & _
                            gsCAT_TCML1_ENGDEPT & " = '" & sEngDept & "', " & _
                            gsCAT_TCML1_STARTTIME & " = " & sStartTime & ", " & _
                            gsCAT_TCML1_EXPECTENDTIME & " = " & sExpectEndTime & ", " & _
                            gsCAT_TCML1_RECENTENDTIME & " = " & sRecentEndTime & ", " & _
                            gsCAT_TCML1_UPDATEUSERID & " = 'SYS', " & _
                            gsCAT_TCML1_UPDATETIME & " = to_char(sysdate, 'yyyymmdd hh24miss') || '000' " & _
                " where " & gsCAT_TCML1_EQID & " = '" & oEquipment.Id & "' " & _
                  " and " & gsCAT_TCML1_DELETEFLAG & " = 'N' "
        Call moProRawSql.QueryDatabase(sSQL)
        
        '然後通知前面更新(也是部分更新)
        If sEventType <> gsMONITOR_TYPE_INIT Then '初始化時不個別更新
            Call Level1Sort(oEquipment.Id) '單機台整理
            Call UpdateMonitorSingleEq(oEquipment.Id)
        End If
    End If
    '-------
    'Done
    '-------
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
           Case Else
                typErrInfo.sUserText = typErrInfo.sErrDescription
        End Select
        '---- Start of Select Case Block ----
        On Error GoTo ExitHandler:
        Call TransferLogError(typErrInfo.sUserText, _
                              moAppLog, Me.rtfTrace, msMODULE_ID, sProcID)
        Call HandleError(False, typErrInfo, , moAppLog)
    End If
End Sub

'================================================================================
'Name       : WaferEndNotice()
'Description: 更新Lot預計結束時間
'Author     : Weilun huang, MXIC 2019/09/04
'Parameters : N/A
'================================================================================
Private Sub WaferEndNotice(ByVal sReciveMsg As String)
    On Error GoTo ExitHandler:
    Dim sProcID                     As String
    Dim typErrInfo                  As tErrInfo
    
    Dim sSQL                        As String
    Dim colRaws                     As Collection
    Dim oRaws                       As Collection
    
    Dim sParentLotID                As String
    Dim sWaferID                    As String
    Dim sStepName                   As String
    Dim sWaferEndTime               As String

    Dim asMsg()                     As String
    Dim sEventType                  As String
    Dim sEqId                       As String
    Dim sEqState                    As String
    
    Dim sLotId                      As String
    Dim sTriggerLotID               As String
    Dim sWaferQty                   As String
    Dim dUph                        As Double
    Dim lWaitTest                   As Long
    Dim lUnTestedWQty               As Long 'Add by Weilun on 20200723 for ReqNo.BE#202000220
    Dim sExpectEndTime              As String
    Dim sOldWaferEndTime            As String

    Dim oLot                        As FwLot
    Dim oEquipment                  As FwEquipment
    
    'Add by Weilun on 20211117 for Project.BE 工業 3.5 Phase 39 - CP AGV導入專案 - 建立 AGV 派工 & Status 監控系統 <start>
    Dim bIsVirtualMerge             As Boolean
    Dim colVir                      As Collection
    'Add by Weilun on 20211117 for Project.BE 工業 3.5 Phase 39 - CP AGV導入專案 - 建立 AGV 派工 & Status 監控系統 <end>
    
    'Add by Weilun on 20230606 for Project.新增CP監控識別虛擬併批功能以及修改預測時間機制 <start>
    Dim sShowLotId                  As String
    Dim sEqpLotIDs                  As String
    'Add by Weilun on 20230606 for Project.新增CP監控識別虛擬併批功能以及修改預測時間機制 <end>
    '-----
    ' Init
    '-----
    sProcID = "WaferEndNotice"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog) '
    
    asMsg = Split(sReciveMsg, "/")
    '"5B224300/TM15/01/SORT2/20180505 2020200000"
    
    '一些欄位保留, 避免以後需要複雜計算
    sEventType = asMsg(0)
    sParentLotID = Left(asMsg(1), 8)
    sEqId = asMsg(2)
    sWaferID = Format(asMsg(3), "00")
'    sStepName = asMsg(4)
    sWaferEndTime = Replace(asMsg(5), "~", " ")
    
    '--------------------
    ' Condition Checking
    '--------------------
    
    '-------
    ' Action
    '-------
    
    Set oEquipment = moFwMDL.EquipmentById(sEqId)
    
    If oEquipment Is Nothing Then
        Call TransferLogError("該Eq不存在.", _
                              moAppLog, Me.rtfTrace, msMODULE_ID, sProcID)
        GoTo ExitHandler
    ElseIf oEquipment.CustomAttributes(gsEQP_CUSTOMATTR_EQ_TYPE2) = "WTESTER(DUMMY)" Then
        Call TransferLogTrace("Skip Dummy Equipment.", _
                              moAppLog, Me.rtfTrace, msMODULE_ID, sProcID)
        GoTo ExitHandler
    Else
        Call EqMonitorDataCheck(oEquipment)
         
        'Modify by Weilun on 20230606 for Project.新增CP監控識別虛擬併批功能以及修改預測時間機制 <start>
        'sLotId = oEquipment.CustomAttributes(gsEQP_CUSTOMATTR_CUR_LOT_ID) '用這個而不用MailBox的主要原因為PEIS傳過來LotID只有前八碼
        Call GetEqDetailLotList(oEquipment.Id, sLotId, sShowLotId, sEqpLotIDs, bIsVirtualMerge, "")
        
'        Set oLot = moFwWIP.LotById(sLotId)
'        If oLot Is Nothing Then
'            sLotId = ""
'        Else
        'Mark by Weilun on 20230606 for Project.新增CP監控識別虛擬併批功能以及修改預測時間機制 <end>
            '可能只接收到前8碼, 所以要找出正確的LotID
            '加上不能為Term以及Ship, Add by Weilun on 20180607 for CPAutoLotEnd First Edition
            sSQL = " select tli." & gsCAT_TLI_LOT_ID & " " & _
                     " from " & gsCAT_TBL_LOT_INFO & " tli, " & _
                                gsCAT_TBL_LOT_ATTRIBUTE & " tlatt " & _
                    " where tli." & gsCAT_TLI_LOT_ID & " = tlatt." & gsCAT_TLATT_LOTID & _
                      " and tli." & gsCAT_TLI_LOT_ID & " like '" & sParentLotID & "%' " & _
                      " and tli." & gsCAT_TLI_WAFERID & " like '%" & sWaferID & "%' " & _
                      " and tli." & gsCAT_TLI_DELETE_FLAG & " = 'N'" & _
                      " and tlatt." & gsCAT_TLATT_LOTSTATUS1 & " not in ('" & gsLOTSTATUS_TERMINATE & "','" & _
                                                                              gsLOTSTATUS_SHIP & "') "
        
            Set colRaws = moProRawSql.QueryDatabase(sSQL)
            If colRaws.Count = 1 Then
                sTriggerLotID = colRaws.Item(1).Item(gsCAT_TLI_LOT_ID)
            ElseIf colRaws.Count > 1 Then
                Call TransferLogError("在Tbl_Lot_Info搜尋到兩筆以上的資料!", _
                                      moAppLog, Me.rtfTrace, msMODULE_ID, sProcID)
                GoTo ExitHandler
            ElseIf colRaws.Count = 0 Then
                Call TransferLogError("在Tbl_Lot_Info找不到資料!", _
                                      moAppLog, Me.rtfTrace, msMODULE_ID, sProcID)
                GoTo ExitHandler
            End If
            
            '檢查Trigger的LotID與機台上的是否相同
'            If sTriggerLotID <> sLotId Then
            If GetStrPosInArray(Replace(sEqpLotIDs, ";", ","), sTriggerLotID, ",") > 0 Then 'Add by Weilun on 20230606 for Project.新增CP監控識別虛擬併批功能以及修改預測時間機制
                '重新指定
                sLotId = sTriggerLotID
            Else
                Call TransferLogTrace(DateTime.Now & ": WaferEnd的Lot與機台上的不相符, 不進行預計結束時間的計算.", _
                                      moAppLog, Me.rtfTrace)
                GoTo ExitHandler
            End If
            
            '取得物件
            'Add by Weilun on 20230606 for Project.新增CP監控識別虛擬併批功能以及修改預測時間機制 <start>
            Set oLot = moFwWIP.LotById(sLotId)
            If oLot Is Nothing Then
                sLotId = ""
                GoTo ExitHandler
            End If
            'Add by Weilun on 20230606 for Project.新增CP監控識別虛擬併批功能以及修改預測時間機制 <end>
        
            '由於KALOS的狀況可會送3次資料,頭尾可能長達10分鐘, 因此相同資訊就不要再重算以免算錯
            'sLotId -> sShowLotId, Modify by Weilun on 20230606 for Project.新增CP監控識別虛擬併批功能以及修改預測時間機制
            sSQL = " select " & gsCAT_TCML1_RECENTENDTIME & _
                     " from " & gsCAT_TBL_CPETEST_MONITOR_LEVEL1 & " " & _
                    " where " & gsCAT_TCML1_EQID & " = '" & oEquipment.Id & "' " & _
                      " and " & gsCAT_TCML1_LOTID & " = '" & sShowLotId & "' " & _
                      " and " & gsCAT_TCML1_RECENTENDTIME & " = '" & sWaferEndTime & "' " & _
                      " and " & gsCAT_TCML1_DELETEFLAG & " = 'N' "
            Set colRaws = moProRawSql.QueryDatabase(sSQL)
            If colRaws.Count = 0 Then '此機台的RecentEndTime與WaferEnd不相同時才重算
                sWaferQty = oLot.CustomAttributes(gsLOT_CUSTOMATTR_WAFER_QTY)
                sStepName = oLot.CurrentStep.Steps.Item(1).Description
                
                dUph = GetUph(oLot, oEquipment, "WS", lWaitTest)
                '取消乘以1.3, 並使用lUnTestedWQty, Mark by Weilun on 20200727 for ReqNo.BE#202000220 <start>

'                If (oEquipment.State = gsEQPSTATUS_RUN Or oEquipment.State = gsEQPSTATUS_ERUN) And _
'                   lWaitTest > 1 And _
'                   dUph <> 0 Then
'
'                    Call TransferLogTrace(DateTime.Now & ": Refresh ExpectEndTime.", _
'                          moAppLog, Me.rtfTrace)
'
'                    'WaferEnd的時候基本上Summary還沒到(TM系列也要30秒), 所以lWaitTest要多1片判斷
'                    '如果只剩1片而又來WaferEnd, 基本上要接著做LotEnd了, 就不更新了
'                    '預計結束的時間, 由WaferEndTime往上加
'                    '[& dUph & "/24 * 1.3)]取消* 1.3成為[& dUph & "/24)], Modify by Weilun on 20200723 for ReqNo.BE#202000220
'                    sExpectEndTime = "to_char(to_date(SubStr('" & sWaferEndTime & "', 1, 15), " & _
'                                                    " 'YYYYMMDD HH24MISS') + (1*" & (lWaitTest - 1) & "/" & dUph & "/24 * 1.3), " & _
'                                            " 'YYYYMMDD HH24MISS') || '000' "
                'Modify by Weilun on 20200727 for ReqNo.BE#202000220 <end>
                          
                '取消乘以1.3, 並使用lUnTestedWQty, Add by Weilun on 20200727 for ReqNo.BE#202000220 <start>
                'Modify by Weilun on 20230606 for Project.新增CP監控識別虛擬併批功能以及修改預測時間機制
'                lUnTestedWQty = GetUnTestedWQty(sLotId, sEqId, "", "", "", sReciveMsg)
                lUnTestedWQty = GetUnTestedTotalWQty(sEqpLotIDs, sEqId, sStepName, sLotId, sReciveMsg)
                
                'lUnTestedWQty從>1改為>0, Modify by Weilun on 20230627 for Project.新增CP監控識別虛擬併批功能以及修改預測時間機制
                If (oEquipment.State = gsEQPSTATUS_RUN Or oEquipment.State = gsEQPSTATUS_ERUN) And _
                   lUnTestedWQty > 0 And _
                   dUph <> 0 Then

                    Call TransferLogTrace(DateTime.Now & ": Refresh ExpectEndTime.", _
                          moAppLog, Me.rtfTrace)
                          
                    '預計結束的時間, 由WaferEndTime往上加
                    sExpectEndTime = "to_char(to_date(SubStr('" & sWaferEndTime & "', 1, 15), " & _
                                                    " 'YYYYMMDD HH24MISS') + (1*" & lUnTestedWQty & "/" & dUph & "/24), " & _
                                            " 'YYYYMMDD HH24MISS') || '000' "
                'Add by Weilun on 20200727 for ReqNo.BE#202000220 <end>
                                            
                    '更新預計結束時間, 這裡未提升效能可能不準, 但是不用跟CPAutoLotEnd搶資源(尤其真的要LotEnd時會一堆Trigger)
                    'sLotId -> sShowLotId, Modify by Weilun on 20230606 for Project.新增CP監控識別虛擬併批功能以及修改預測時間機制
                    sSQL = " update " & gsCAT_TBL_CPETEST_MONITOR_LEVEL1 & " " & _
                              " set " & gsCAT_TCML1_EXPECTENDTIME & " = " & sExpectEndTime & ", " & _
                                        gsCAT_TCML1_RECENTENDTIME & " = '" & sWaferEndTime & "', " & _
                                        gsCAT_TCML1_UPDATEUSERID & " = 'SYS', " & _
                                        gsCAT_TCML1_UPDATETIME & " = to_char(sysdate, 'yyyymmdd hh24miss') || '000' " & _
                            " where " & gsCAT_TCML1_EQID & " = '" & oEquipment.Id & "' " & _
                              " and " & gsCAT_TCML1_LOTID & " = '" & sShowLotId & "' " & _
                              " and " & gsCAT_TCML1_STEPNAME & " = '" & sStepName & "' " & _
                              " and " & gsCAT_TCML1_STARTTIME & " IS NOT NULL " & _
                              " and " & gsCAT_TCML1_DELETEFLAG & " = 'N' "
                    Call moProRawSql.QueryDatabase(sSQL)
                    
                    
                    
                    '然後再更新一階資訊(部分更新模組)
                    Call Level1Sort(oEquipment.Id)
                    
                    '然後通知前面更新(也是部分更新)
                    Call UpdateMonitorSingleEq(oEquipment.Id)
                    
                    'Add by Weilun on 20230606 for Project.新增CP監控識別虛擬併批功能以及修改預測時間機制 <start>
                    'If lUnTestedWQty = 1 Then 'Mark by Sam on 20230821 for ReqNo:202300189,改為每片都計算
                        Call UpdateAGV_EqStatusUnloadInfo(oEquipment.Id, sLotId, sExpectEndTime) '請不要傳入虛擬併批母批
                    'End If
                    Call UpdateAGVDispatchEndTime(oEquipment.Id, sExpectEndTime)
                    'Add by Weilun on 20230606 for Project.新增CP監控識別虛擬併批功能以及修改預測時間機制 <end>
                    
                    'Mark by Weilun on 20230606 for Project.新增CP監控識別虛擬併批功能以及修改預測時間機制 <start>
'                    'Add by Weilun on 20211117 for Project.BE 工業 3.5 Phase 39 - CP AGV導入專案 - 建立 AGV 派工 & Status 監控系統 <start>
'                    '虛擬併批因CurLotID, 除了最後一批應該走不到這裡
'                    '不過還是得把虛擬併批排除
'                    bIsVirtualMerge = False
'                    sSQL = "select " & gsCAT_TLI_VIRTUALLOTID & ", " & _
'                                           gsCAT_TLI_MERGETYPE & " " & _
'                                " from " & gsCAT_TBL_LOT_INFO & " " & _
'                               " where " & gsCAT_TLI_LOT_ID & " = '" & sLotId & "'  "
'                    Set colVir = moProRawSql.QueryDatabase(sSQL)
'                    If colVir.Count > 0 Then
'                        If colVir.Item(1).Item(gsCAT_TLI_VIRTUALLOTID) <> "" And _
'                           colVir.Item(1).Item(gsCAT_TLI_MERGETYPE) = "VirtualMerge" Then
'                            bIsVirtualMerge = True '為虛擬併批
'                        End If
'                    End If
'
'                    '只是非虛擬併批最後WaferEnd誰寫? (UnTestedWQty相關)
'                    If bIsVirtualMerge = False Then
'                        Call UpdateAGVDispatchEndTime(oEquipment.Id, sExpectEndTime)
'                    End If
'                    'Add by Weilun on 20211117 for Project.BE 工業 3.5 Phase 39 - CP AGV導入專案 - 建立 AGV 派工 & Status 監控系統 <end>
                    'Mark by Weilun on 20230606 for Project.新增CP監控識別虛擬併批功能以及修改預測時間機制 <end>
                End If
            Else
                If sWaferEndTime = colRaws.Item(1).Item(gsCAT_TCML1_RECENTENDTIME) Then
                    Call TransferLogTrace(DateTime.Now & ": Skip refresh ExpectEndTime, which receive same WaferEndTime.", _
                                          moAppLog, Me.rtfTrace)
                    GoTo ExitHandler
                End If
            End If
'        End If
    End If
    '-------
    'Done
    '-------
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
           Case Else
                typErrInfo.sUserText = typErrInfo.sErrDescription
        End Select
        '---- Start of Select Case Block ----
        On Error GoTo ExitHandler:
        Call TransferLogError(typErrInfo.sUserText, _
                              moAppLog, Me.rtfTrace, msMODULE_ID, sProcID)
        Call HandleError(False, typErrInfo, , moAppLog)
    End If
End Sub

'================================================================================
'Name       : TestAbortAlarm()
'Description: TestAbort訊息寫入
'Author     : Weilun huang, MXIC 2019/10/03
'Parameters : N/A
'================================================================================
Private Sub TestAbortAlarm(ByVal sReciveMsg As String)
    On Error GoTo ExitHandler
    Dim sProcID As String
    Dim typErrInfo As tErrInfo

    Dim asMsg() As String
    Dim sLotEndStatus As String
    Dim sLotCompleteStatus As String

    Dim sEventType As String
    Dim sLotId As String
    Dim sEqId As String
    Dim sWaferID As String
    Dim sStepName As String

    Dim sSQL As String
    Dim colRaws As Collection
    Dim oRaws As Collection

    '-----
    ' Init
    '-----
    sProcID = "CPAutoChange"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog) '

    asMsg = Split(sReciveMsg, "/")

    sEventType = asMsg(0)
    sLotId = asMsg(1)
    sEqId = asMsg(2)
    sWaferID = asMsg(3)
    sStepName = asMsg(4)

    '--------------------
    ' Condition Checking
    '--------------------

    '-------
    ' Action
    '-------
    If sEqId <> "" And sLotId <> "" And Trim(GetErunTicNO(sLotId)) = "" Then 'Add Trim(GetErunTicNO(sLotID)) = "" 量產品檢查 by Weilun on 20191125 for 智慧監控 fix
        Call TransferLogTrace(DateTime.Now & ": Add '" & gsMONITOR_ITEM_DESC_TESTABORT & "' for " & sLotId & "/" & sEqId, _
                              moAppLog, Me.rtfTrace)

        Call InsertMonitorReceive(sEqId, _
                                  sLotId, _
                                  sWaferID, _
                                  "", _
                                  gsMONITOR_ITEM_TESTDATA, _
                                  gsMONITOR_ITEM_CONT_SUMMARY, _
                                  gsMONITOR_ITEM_DESC_TESTABORT, _
                                  "1", _
                                  gsCAT_TCMR_PRODTAFLAG)

        '單機台整理
        Call Level1Sort(sEqId)

        '單機台更新
        Call UpdateMonitorSingleEq(sEqId)
    End If
    '-------
    'Done
    '-------
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
            Case Else
                typErrInfo.sUserText = typErrInfo.sErrDescription
        End Select
        '---- Start of Select Case Block ----
        On Error GoTo ExitHandler
        Call TransferLogError(typErrInfo.sUserText, _
                              moAppLog, Me.rtfTrace, msMODULE_ID, sProcID)
        Call HandleError(False, typErrInfo, , moAppLog)
    End If
End Sub

'================================================================================
'Name       : CPAutoChange()
'Description: 依CP_Auto_Complete的狀態變動進行相對應的Monitor作業
'Author     : Weilun huang, MXIC 2019/09/03
'Parameters : N/A
'================================================================================
Private Sub CPAutoChange(ByVal sReciveMsg As String)
    On Error GoTo ExitHandler
    Dim sProcID As String
    Dim typErrInfo As tErrInfo

    Dim asMsg() As String
    Dim sLotEndStatus As String
    Dim sLotCompleteStatus As String

    Dim sEventType As String
    Dim sLotId As String
    Dim sStepName As String
    Dim sMainStatus As String
    Dim sActionType As String
    Dim sAction As String

    Dim sSQL As String
    Dim colRaws As Collection
    Dim oRaws As Collection

    Dim oLot As FwLot
    Dim oEquipment As FwEquipment

    '-----
    ' Init
    '-----
    sProcID = "CPAutoChange"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog) '

    asMsg = Split(sReciveMsg, "/")

    sEventType = asMsg(0)
    sLotId = asMsg(1)
    sStepName = asMsg(2)
    sLotEndStatus = asMsg(3)
    sLotCompleteStatus = asMsg(4)

    '--------------------
    ' Condition Checking
    '--------------------

    '-------
    ' Action
    '-------

    Set oLot = moFwWIP.LotById(sLotId)
    If oLot Is Nothing Then
        Call TransferLogError("該Lot不存在.", _
                              moAppLog, Me.rtfTrace, msMODULE_ID, sProcID)
        GoTo ExitHandler
    End If

    Set oEquipment = moFwMDL.EquipmentById(oLot.CustomAttributes(gsLOT_CUSTOMATTR_CUR_EQP_ID))
    '分批合測併回母批後, 母批由於已經清空EqID, 必須另外抓 (在這個Function為防呆, 應對手動LotComplete強制LotEnd的Case)
'    If oEquipment Is Nothing And oLot.CustomAttributes(gsLOT_CUSTOMATTR_STATUS) = gsLOTSTATUS_BREAK Then
    If oEquipment Is Nothing Then 'LotComplete的可能性太多, 不再考慮Lot狀態, Modify by Weilun on 20191107 for 智慧監控 fix
        Set oEquipment = moFwMDL.EquipmentById(oLot.CustomAttributes(gsLOT_CUSTOMATTR_LAST_EQP_ID))
    End If

    If Not oEquipment Is Nothing Then
        If sLotEndStatus = gsCPAUTO_STATUS_INIT Then

            '搜尋並刪除同LotID和StepName的LotEnd紀錄, 有可能AutoCheckFail或是其他狀況退回去重跑
            sSQL = " select " & gsCAT_TCMR_CREATETIME & " " & _
                     " from " & gsCAT_TBL_CPETEST_MONITOR_RECEIVE & " " & _
                    " where " & gsCAT_TCMR_LOTID & " = '" & oLot.Id & "' " & _
                      " and " & gsCAT_TCMR_STEPNAME & " = '" & sStepName & "' " & _
                      " and " & gsCAT_TCMR_MONITORITEM & " = '" & gsMONITOR_ITEM_TESTDATA & "' " & _
                      " and " & gsCAT_TCMR_MONITORITEMCONT & " = '" & gsMONITOR_ITEM_CONT_SORTRUN & "' " & _
                      " and " & gsCAT_TCMR_MONITORITEMCONTDESC & " = '" & gsMONITOR_ITEM_DESC_WAITMOVE & "' " & _
                      " and " & gsCAT_TCMR_DELETEFLAG & " = 'N' "
            Set colRaws = moProRawSql.QueryDatabase(sSQL)
            If colRaws.Count > 0 Then
                Call TransferLogTrace(DateTime.Now & ": Clear '" & gsMONITOR_ITEM_DESC_WAITMOVE & "' for " & oLot.Id & "/" & sStepName, _
                          moAppLog, Me.rtfTrace)
                sSQL = " update " & gsCAT_TBL_CPETEST_MONITOR_RECEIVE & " " & _
                          " set " & gsCAT_TCMR_DELETEFLAG & " = 'Y', " & _
                                    gsCAT_TCMR_DELETEUSERID & " = 'SYS', " & _
                                    gsCAT_TCMR_DELETETIME & " = to_char(sysdate, 'yyyymmdd hh24miss') || '000' " & _
                        " where " & gsCAT_TCMR_LOTID & " = '" & oLot.Id & "' " & _
                          " and " & gsCAT_TCMR_STEPNAME & " = '" & sStepName & "' " & _
                          " and " & gsCAT_TCMR_MONITORITEM & " = '" & gsMONITOR_ITEM_TESTDATA & "' " & _
                          " and " & gsCAT_TCMR_MONITORITEMCONT & " = '" & gsMONITOR_ITEM_CONT_SORTRUN & "' " & _
                          " and " & gsCAT_TCMR_MONITORITEMCONTDESC & " = '" & gsMONITOR_ITEM_DESC_WAITMOVE & "' " & _
                          " and " & gsCAT_TCMR_DELETEFLAG & " = 'N' "
                Call moProRawSql.QueryDatabase(sSQL)
            Else
                Call TransferLogTrace(DateTime.Now & ": No Action.", _
                                      moAppLog, Me.rtfTrace)
                GoTo ExitHandler '跳過整理和更新
            End If

            '自動LotEnd或手動LotComplete的Case
        ElseIf (sLotEndStatus = gsCPAUTO_STATUS_AUTO_OK And sLotCompleteStatus = gsCPAUTO_STATUS_INIT) Or _
               (sLotEndStatus <> gsCPAUTO_STATUS_AUTO_OK And sLotCompleteStatus = gsCPAUTO_STATUS_MANUAL_OK) Then

            Call TransferLogTrace(DateTime.Now & ": Add '" & gsMONITOR_ITEM_DESC_WAITMOVE & "' for " & oLot.Id & "/" & oEquipment.Id & "/" & sStepName, _
                      moAppLog, Me.rtfTrace)

            Call InsertMonitorReceive(oEquipment.Id, _
                          oLot.Id, _
                          "NA", _
                          sStepName, _
                          gsMONITOR_ITEM_TESTDATA, _
                          gsMONITOR_ITEM_CONT_SORTRUN, _
                          gsMONITOR_ITEM_DESC_WAITMOVE, _
                          "99", _
                          gsCAT_TCMR_PRODTAFLAG)
                          
            'Add by Weilun on 20211125 for Project.BE 工業 3.5 Phase 39 - CP AGV導入專案 - 建立 AGV 派工 & Status 監控系統
            Call InsertMonitorReceive(oEquipment.Id, _
                                      oLot.Id, _
                                      "NA", _
                                      sStepName, _
                                      gsMONITOR_ITEM_TESTDATA, _
                                      gsMONITOR_ITEM_CONT_AUTOMONITOR, _
                                      gsMONITOR_ITEM_DESC_UNLOAD, _
                                      "1", _
                                      gsCAT_TCMR_PRODTAFLAG)

        Else
            '其他Case

            'LotCompleteStatus優先, 並取得狀態
            If sLotCompleteStatus <> "" Then
                sMainStatus = sLotCompleteStatus
                sActionType = msComplete
            Else
                sMainStatus = sLotEndStatus
                sActionType = msEnd
            End If

            If sMainStatus = gsCPAUTO_STATUS_AUTO_FAIL Or _
               sMainStatus = gsCPAUTO_STATUS_AUTO_CHECKFAIL Or _
               sMainStatus = gsCPAUTO_STATUS_AUTO_ERROR Then
                'CP自動化異常

                sAction = "Please execute " & sActionType & "."

                Call TransferLogTrace(DateTime.Now & ": Add '" & sAction & "' for " & oLot.Id & "/" & oEquipment.Id & "/" & sStepName, _
                                      moAppLog, Me.rtfTrace)

                Call InsertMonitorReceive(oEquipment.Id, _
                                          oLot.Id, _
                                          "NA", _
                                          sStepName, _
                                          gsMONITOR_ITEM_TESTDATA, _
                                          gsMONITOR_ITEM_CONT_SORTRUN, _
                                          sAction, _
                                          "1", _
                                          gsCAT_TCMR_PRODTAFLAG)
            Else
                Call TransferLogTrace(DateTime.Now & ": No Action.", _
                                      moAppLog, Me.rtfTrace)
                GoTo ExitHandler '跳過整理和更新
            End If
        End If
        
        'AGV相關, 這裡暫時保留與備忘, Comment by Weilun on 20211117 for Project.BE 工業 3.5 Phase 39 - CP AGV導入專案 - 建立 AGV 派工 & Status 監控系統
        '由於在AGV專案期間, 監控系統還沒有虛擬併批相關設計, 因此更新AGV相關資料分散至Auto/Manual LotEnd等
        '看狀況以後要不要整理至此

        '單機台整理
        Call Level1Sort(oEquipment.Id)

        '單機台更新
        Call UpdateMonitorSingleEq(oEquipment.Id)
    End If
    '-------
    'Done
    '-------
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
            Case Else
                typErrInfo.sUserText = typErrInfo.sErrDescription
        End Select
        '---- Start of Select Case Block ----
        On Error GoTo ExitHandler
        Call TransferLogError(typErrInfo.sUserText, _
                              moAppLog, Me.rtfTrace, msMODULE_ID, sProcID)
        Call HandleError(False, typErrInfo, , moAppLog)
    End If
End Sub

'================================================================================
'Name       : CPCompleteCheck()
'Description: MoveNext訊息滿30分鐘後檢查是否LotComplete
'Author     : Weilun huang, MXIC 2019/09/03
'Parameters : N/A
'================================================================================
Private Sub CPCompleteCheck()
    On Error GoTo ExitHandler
    Dim sProcID As String
    Dim typErrInfo As tErrInfo

    Dim sLotId As String
    Dim sStepName As String

    Dim sSQL As String
    Dim colRaws As Collection
    Dim oRaws As Collection

    Dim sCPSQL As String
    Dim sDeleteSQL As String
    Dim colCPRaws As Collection

    Dim sCompleteWait As String
    Dim bUpdateEqList As String

    '-----
    ' Init
    '-----
    sProcID = "CPCompleteCheck"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog)

    sCompleteWait = "1/48" '30分鐘
    bUpdateEqList = ""

    '--------------------
    ' Condition Checking
    '--------------------
    If chkSchedualItem(1).Value = vbUnchecked Then
        GoTo ExitHandler
    End If


    '-------
    ' Action
    '-------

    Call TransferLogTrace(DateTime.Now & ": Wait move lot check CP complete start.", _
                      moAppLog, Me.rtfTrace)

    '找出通知LotEnd已30分鐘的Lot
    sSQL = " select " & gsCAT_TCMR_LOTID & ", " & _
                        gsCAT_TCMR_STEPNAME & " " & _
             " from " & gsCAT_TBL_CPETEST_MONITOR_RECEIVE & " " & _
            " where " & gsCAT_TCMR_MONITORITEM & " = '" & gsMONITOR_ITEM_TESTDATA & "' " & _
              " and " & gsCAT_TCMR_MONITORITEMCONT & " = '" & gsMONITOR_ITEM_CONT_SORTRUN & "' " & _
              " and " & gsCAT_TCMR_MONITORITEMCONTDESC & " = '" & gsMONITOR_ITEM_DESC_WAITMOVE & "' " & _
              " and " & gsCAT_TCMR_DELETEFLAG & " = 'N' " & _
              " and to_char(to_date(SubStr(" & gsCAT_TCMR_CREATETIME & ", 1, 15), " & _
                             " 'YYYYMMDD HH24MISS') + " & sCompleteWait & ", " & _
                     " 'YYYYMMDD HH24MISS') || '000' < " & _
                  " to_char(sysdate, 'YYYYMMDD HH24MISS') || '000' "
    Set colRaws = moProRawSql.QueryDatabase(sSQL)

    If colRaws.Count > 0 Then
        For Each oRaws In colRaws
            sLotId = oRaws.Item(gsCAT_TCMR_LOTID)
            sStepName = oRaws.Item(gsCAT_TCMR_STEPNAME)

            '看該LotID和StepName的最新資料是否已LotComplete
            sCPSQL = " select " & gsCAT_TCAC_UPDATETIME & " " & _
                       " from " & gsCAT_TBL_CP_AUTO_COMPLETE & " " & _
                      " where (" & gsCAT_TCAC_LOTID & ", " & _
                                   gsCAT_TCAC_STEPNAME & ", " & _
                                   gsCAT_TCAC_CREATETIME & ") in " & _
                            " (select " & gsCAT_TCAC_LOTID & ", " & _
                                          gsCAT_TCAC_STEPNAME & ", " & _
                                " Max(" & gsCAT_TCAC_CREATETIME & ") " & _
                               " from " & gsCAT_TBL_CP_AUTO_COMPLETE & " " & _
                              " where " & gsCAT_TCAC_LOTID & " = '" & sLotId & "' " & _
                                " and " & gsCAT_TCAC_STEPNAME & " = '" & sStepName & "' " & _
                           " group by " & gsCAT_TCAC_LOTID & ", " & gsCAT_TCAC_STEPNAME & ") " & _
                        " and " & gsCAT_TCAC_DELETEFLAG & " = 'N' " & _
                        " and " & gsCAT_TCAC_LOTCOMPLETESTATUS & " in ('" & gsCPAUTO_STATUS_MANUAL_OK & "', " & _
                                                                     " '" & gsCPAUTO_STATUS_AUTO_OK & "' ) "
            Set colCPRaws = moProRawSql.QueryDatabase(sCPSQL)
            If colCPRaws.Count > 0 Then
                '刪除相關紀錄(基本上只保留最後一筆可以刪)
                Call TransferLogTrace(DateTime.Now & ": Clear '" & gsMONITOR_ITEM_DESC_WAITMOVE & "' for " & sLotId & "/" & sStepName, _
                          moAppLog, Me.rtfTrace)
                sDeleteSQL = " update " & gsCAT_TBL_CPETEST_MONITOR_RECEIVE & " " & _
                                " set " & gsCAT_TCMR_DELETEFLAG & " = 'Y', " & _
                                          gsCAT_TCMR_DELETEUSERID & " = 'SYS', " & _
                                          gsCAT_TCMR_DELETETIME & " = to_char(sysdate, 'yyyymmdd hh24miss') || '000' " & _
                              " where " & gsCAT_TCMR_LOTID & " = '" & sLotId & "' " & _
                                " and " & gsCAT_TCMR_STEPNAME & " = '" & sStepName & "' " & _
                                " and " & gsCAT_TCMR_MONITORITEM & " = '" & gsMONITOR_ITEM_TESTDATA & "' " & _
                                " and " & gsCAT_TCMR_MONITORITEMCONT & " = '" & gsMONITOR_ITEM_CONT_SORTRUN & "' " & _
                                " and " & gsCAT_TCMR_MONITORITEMCONTDESC & " = '" & gsMONITOR_ITEM_DESC_WAITMOVE & "' " & _
                                " and " & gsCAT_TCMR_DELETEFLAG & " = 'N' "
                Call moProRawSql.QueryDatabase(sDeleteSQL)
            End If
        Next
    End If

    '找出並刪除通知LotEnd已"一整天"的紀錄 (預防性自我維護)
    sSQL = " update " & gsCAT_TBL_CPETEST_MONITOR_RECEIVE & " " & _
              " set " & gsCAT_TCMR_DELETEFLAG & " = 'Y', " & _
                        gsCAT_TCMR_DELETEUSERID & " = 'SYS', " & _
                        gsCAT_TCMR_DELETETIME & " = to_char(sysdate, 'yyyymmdd hh24miss') || '000' " & _
            " where " & gsCAT_TCMR_MONITORITEM & " = '" & gsMONITOR_ITEM_TESTDATA & "' " & _
              " and " & gsCAT_TCMR_MONITORITEMCONT & " = '" & gsMONITOR_ITEM_CONT_SORTRUN & "' " & _
              " and " & gsCAT_TCMR_MONITORITEMCONTDESC & " = '" & gsMONITOR_ITEM_DESC_WAITMOVE & "' " & _
              " and " & gsCAT_TCMR_DELETEFLAG & " = 'N' " & _
              " and to_char(to_date(SubStr(" & gsCAT_TCMR_CREATETIME & ", 1, 15), " & _
                             " 'YYYYMMDD HH24MISS') + " & "1" & ", " & _
                     " 'YYYYMMDD HH24MISS') || '000' < " & _
                  " to_char(sysdate, 'YYYYMMDD HH24MISS') || '000' "
    Call moProRawSql.QueryDatabase(sSQL)

    Call TransferLogTrace(DateTime.Now & ": Wait move lot check CP complete end.", _
                          moAppLog, Me.rtfTrace)

    '各個更新檢查都完成後, 再更新一二階Table, 再通知前端

    '-------
    'Done
    '-------
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
            Case Else
                typErrInfo.sUserText = typErrInfo.sErrDescription
        End Select
        '---- Start of Select Case Block ----
        On Error GoTo ExitHandler
        Call TransferLogError(typErrInfo.sUserText, _
                              moAppLog, Me.rtfTrace, msMODULE_ID, sProcID)
        Call HandleError(False, typErrInfo, , moAppLog)
    End If
End Sub

'================================================================================
'Name       : TestTimeAbnormalCheck()
'Description: 檢查各機台的測試時間是否有異常
'Author     : Weilun huang, MXIC 2019/09/24
'Parameters : N/A
'================================================================================
Private Sub TestTimeAbnormalCheck()
    On Error GoTo ExitHandler
    Dim sProcID As String
    Dim typErrInfo As tErrInfo

    Dim sSQL As String
    Dim colRaws As Collection
    Dim oRaws As Collection

    Dim sCheckSQL As String
    Dim colCheckRaws As Collection
    Dim oCheckRaws As Collection

    Dim sCheckSQL2 As String
    Dim colCheckRaws2 As Collection
    Dim oCheckRaws2 As Collection

    Dim sModifySQL As String 'input, update, delete

    Dim sEqId As String
    Dim sLotId As String
    Dim sWaferID As String
    Dim sRecentEndTime As String
    Dim dUph As Double
    Dim lWaitTest As Long

    Dim oLot As FwLot
    Dim oEquipment As FwEquipment
    
    Dim sBuffer    As String
    Dim sMsg       As String
    
    'Add by Weilun on 20230606 for Project.新增CP監控識別虛擬併批功能以及修改預測時間機制 <start>
    Dim sShowLotId                  As String
    Dim sEqpLotIDs                  As String
    Dim sVirtualLotId               As String
    Dim sCurRunSlotNo               As String
    Dim sCurRunLotId                As String
    Dim sCurRunWaferID              As String
    Dim bIsVirtualMerge             As Boolean
    'Add by Weilun on 20230606 for Project.新增CP監控識別虛擬併批功能以及修改預測時間機制 <end>
    Dim sErrorConnectFlag           As String
    '-----
    ' Init
    '-----
    sProcID = "TestTimeAbnormalCheck"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog) '

    '--------------------
    ' Condition Checking
    '--------------------
    If chkSchedualItem(0).Value = vbUnchecked Then
        GoTo ExitHandler
    End If


    '-------
    ' Action
    '-------

    Call TransferLogTrace(DateTime.Now & ": Test time abnormal check start.", _
                          moAppLog, Me.rtfTrace)

    sSQL = "select " & gsCAT_TCML1_EQID & " " & _
            " from " & gsCAT_TBL_CPETEST_MONITOR_LEVEL1 & " " & _
           " where " & gsCAT_TCML1_LOTID & " Is Not Null " & _
             " and " & gsCAT_TCML1_EXPECTENDTIME & " is not null " & _
             " and " & gsCAT_TCML1_DELETEFLAG & " = 'N' "

    Set colRaws = moProRawSql.QueryDatabase(sSQL)
    If colRaws.Count > 0 Then
        For Each oRaws In colRaws
            '取得機台物件
            sEqId = oRaws.Item(gsCAT_TCML1_EQID)
            Set oEquipment = moFwMDL.EquipmentById(sEqId)
            If Not oEquipment Is Nothing Then
                
                'Add by Weilun on 20230606 for Project.新增CP監控識別虛擬併批功能以及修改預測時間機制
                Call GetEqDetailLotList(oEquipment.Id, sLotId, sShowLotId, sEqpLotIDs, bIsVirtualMerge, sVirtualLotId) '取得EQ上的LotInfo詳細資訊, 包含虛擬併批的代表批

                '取得Lot物件
'                sLotID = oEquipment.CustomAttributes(gsEQP_CUSTOMATTR_CUR_LOT_ID)
                Set oLot = moFwWIP.LotById(sLotId)
                If Not oLot Is Nothing Then

                    dUph = GetUph(oLot, oEquipment, "WS", lWaitTest) '取片Uph

                    If dUph > 0 And Trim(GetErunTicNO(sLotId)) = "" Then 'Add Trim(GetErunTicNO(sLotID)) = "" 量產品檢查 by Weilun on 20191125 for 智慧監控 fix
                        '針對每一機台去算
                        
                        'Add by Weilun on 20200723 for ReqNo.BE#202000220 <start>
                        If dUph > 0.5 Then
                            sBuffer = " + (1 / 48)" '30分鐘
                        Else
                            sBuffer = " * 1.3" '1.3倍
                        End If
                        'Add by Weilun on 20200723 for ReqNo.BE#202000220 <end>
                        
                        '固定乘以1.3倍的[dUph & " * 1.3) then]改為[dUph & sBuffer & ") then], Modify by Weilun on 20200723 for ReqNo.BE#202000220
                        'sLotID改sShowLotId, 'Modify by Weilun on 20230606 for Project.新增CP監控識別虛擬併批功能以及修改預測時間機制
                        sCheckSQL = " select " & gsCAT_TCML1_RECENTENDTIME & ", " & _
                                           " Case " & _
                                             " When (sysdate - to_date(SubStr(" & gsCAT_TCML1_RECENTENDTIME & ", 1, 15), " & _
                                                                     " 'YYYYMMDD HH24MISS')) > ((1/24) / " & dUph & sBuffer & ") then " & _
                                              " 'Y' " & _
                                             " Else " & _
                                              " 'N' " & _
                                           " End AS CheckResult " & _
                                      " from " & gsCAT_TBL_CPETEST_MONITOR_LEVEL1 & " " & _
                                     " where " & gsCAT_TCML1_LOTID & " = '" & sShowLotId & "' " & _
                                       " and " & gsCAT_TCML1_RECENTENDTIME & " is not null " & _
                                       " and " & gsCAT_TCML1_EXPECTENDTIME & " is not null " & _
                                       " and " & gsCAT_TCML1_EQID & " = '" & sEqId & "' " & _
                                       " and " & gsCAT_TCML1_DELETEFLAG & " = 'N' "

                        Set colCheckRaws = moProRawSql.QueryDatabase(sCheckSQL)
                        If colCheckRaws.Count > 0 Then
                            If colCheckRaws.Item(1).Item("CheckResult") = "Y" Then

                                sRecentEndTime = colCheckRaws.Item(1).Item(gsCAT_TCML1_RECENTENDTIME)

'                                'WaferID這裡取
                                If bIsVirtualMerge = True Then 'Add by Weilun on 20230606 for Project.新增CP監控識別虛擬併批功能以及修改預測時間機制
                                    Call GetVirtualLotCurRunInfo(sEqpLotIDs, sVirtualLotId, sCurRunSlotNo, "", "")
                                    sWaferID = sCurRunSlotNo
                                Else '原來的
                                    sWaferID = GetCurWaferIDByPgmList(sLotId)
                                End If

                                '是否有重複的資料(RecentEndTime之後)
                                'sLotID改sShowLotId, 'Modify by Weilun on 20230606 for Project.新增CP監控識別虛擬併批功能以及修改預測時間機制
                                sCheckSQL2 = " select " & gsCAT_TCMR_CREATETIME & " " & _
                                               " from " & gsCAT_TBL_CPETEST_MONITOR_RECEIVE & " " & _
                                              " where " & gsCAT_TCMR_CREATETIME & " > '" & sRecentEndTime & "' " & _
                                                " and " & gsCAT_TCMR_WAFERID & " = '" & sWaferID & "' " & _
                                                " and " & gsCAT_TCMR_EQID & " = '" & sEqId & "' " & _
                                                " and " & gsCAT_TCMR_LOTID & " = '" & sShowLotId & "' " & _
                                                " and " & gsCAT_TCMR_MONITORITEM & " = '" & gsMONITOR_ITEM_TESTDATA & "' " & _
                                                " and " & gsCAT_TCMR_MONITORITEMCONT & " = '" & gsMONITOR_ITEM_CONT_SUMMARY & "' " & _
                                                " and " & gsCAT_TCMR_MONITORITEMCONTDESC & " = '" & gsMONITOR_ITEM_DESC_TESTTIME & "' " & _
                                                " and " & gsCAT_TCMR_DELETEFLAG & " = 'N' "
                                Set colCheckRaws2 = moProRawSql.QueryDatabase(sCheckSQL2)

                                If colCheckRaws2.Count = 0 Then '沒有該時間區段的資料才加入
                                    'sLotID改sShowLotId, 'Modify by Weilun on 20230606 for Project.新增CP監控識別虛擬併批功能以及修改預測時間機制
                                    Call TransferLogTrace(DateTime.Now & ": " & sShowLotId & " test time abnormal on " & sEqId, _
                                                          moAppLog, Me.rtfTrace)
                                    'sLotID改sShowLotId, 'Modify by Weilun on 20230606 for Project.新增CP監控識別虛擬併批功能以及修改預測時間機制
                                    Call InsertMonitorReceive(sEqId, _
                                                              sShowLotId, _
                                                              sWaferID, _
                                                              "", _
                                                              gsMONITOR_ITEM_TESTDATA, _
                                                              gsMONITOR_ITEM_CONT_SUMMARY, _
                                                              gsMONITOR_ITEM_DESC_TESTTIME, _
                                                              "1", _
                                                              gsCAT_TCMR_PRODTAFLAG)
                                End If
                            End If
                        End If '沒有找到不管了, 或許是Lot已不再機台上甚至LotID是不一樣
                    End If 'dUph > 0
                End If 'Not oLot Is Nothing
            End If 'Not oEquipment Is Nothing
        Next
    End If

    Call TransferLogTrace(DateTime.Now & ": Test time abnormal check end.", _
                          moAppLog, Me.rtfTrace)
    '-------
    'Done
    '-------
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
        'modify by Ernest on 2023/07/23 for #162501 TA monitor system -Message bus connection was lost 改善
        
        'If typErrInfo.lErrNumber = 41224 Then
        '    Call TransferLogError("MES斷線", _
                                  moAppLog, Me.rtfTrace, msMODULE_ID, sProcID)
                                  
        '    Sleep 30000
            'Call EMail(moSMTP, moAppLog, Me.txtSmtp, "Admin", Me.txtEMail(0), Me.txtCC(0), msMODULE_ID & " " & typErrInfo.sErrDescription, DateTime.Now, , Me.rtfTrace)
        '    Set moAppLog = Nothing
        '    Set moRawSql = Nothing
        '    Set moProRawSql = Nothing
        '    Set moSMTP = Nothing
        
        '    Set moCwMbx = Nothing
        '    Set moFwCAT = Nothing
        '    Set moFwMDL = Nothing
        '    moFwOPR.Logout
        '    Set moFwOPR = Nothing
        '    Set moFwPRP = Nothing
        '    Set moFwWF = Nothing
        '    Set moFwWIP = Nothing
            
        '    Set moClient = Nothing
            
        '    Set moSC_Main = Nothing
        '    Set moSC_Sub = Nothing
            
        '    bReconnect = True
            
        '    Call cmdConnect_Click
        'Else
            ' NOTE 2:
            ' If you have custom handling of some Errors, please
            ' UN-REMARED the following Select Case block!
            ' Also, modify if neccessarily!!!
            '---- Start of Select Case Block ----
            
            'ADD by Rubio Huang on 20230907 For Mailbox AutoReconnection
            'Modified by Samson on 20260304 call cmdClose_Click -> call Restart
            If typErrInfo.lErrNumber = 41224 And chkAutoReconnection.Value = vbChecked Then
                Call TransferLogError(DateTime.Now & "MES Connection was lost,AutoReconnection Start", _
                                  moAppLog, Me.rtfTrace, msMODULE_ID, sProcID)
                
                Call Restart
            End If
            
            Select Case typErrInfo.lErrNumber
                Case Else
                    typErrInfo.sUserText = typErrInfo.sErrDescription
            End Select
            '---- Start of Select Case Block ----
            On Error GoTo ExitHandler
            Call TransferLogError(typErrInfo.sUserText, _
                                  moAppLog, Me.rtfTrace, msMODULE_ID, sProcID)
            Call HandleError(False, typErrInfo, , moAppLog)
        'End If
    End If
End Sub

'================================================================================
'Name       : NextLotPrepare()
'Description: 對於30分鐘內結束測試或是IDLE中的機台, 進行新增或刪除待測Lot資訊的功能
'Author     : Weilun huang, MXIC 2019/09/10
'Parameters : N/A
'================================================================================
Private Sub NextLotPrepare()
    On Error GoTo ExitHandler
    Dim sProcID As String
    Dim typErrInfo As tErrInfo

    Dim sSQL                As String
    Dim colRaws             As Collection
    Dim oRaws               As Collection

    Dim sLotSQL             As String
    Dim colLotRaws          As Collection
    Dim oLotRaws            As Collection

    Dim sCheckSQL           As String
    Dim colCheckRaws        As Collection
    Dim oCheckRaws          As Collection
    
    Dim sCPSQL              As String
    Dim colCPRaws           As Collection
    Dim oCPRaws             As Collection

    Dim sModifySQL          As String 'input, update, delete
    Dim sLotId              As String
    Dim sEqId               As String
    Dim sEqIDList           As String
    Dim sEqIDListSQL        As String 'Where in sEqIDListSQL

    Dim sExistLotIDlist     As String
    Dim sCurLotIdList       As String
    Dim sAddLotIDList       As String
    Dim sRemoveLotIDList    As String
    Dim asAddList()         As String
    Dim asRemoveList()      As String
    
    Dim sStepName           As String
    Dim bAddLotID           As Boolean

    Dim lIdx                As Long
    
    Dim sEqTypeList         As String 'Add by Weilun on 20200513 for Project.BE MES Phase 44 – CP 產線增加自動化作業
    '-----
    ' Init
    '-----
    sProcID = "NextLotPrepare"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog) '
    
    sEqTypeList = "WTESTER" 'Add by Weilun on 20200513 for Project.BE MES Phase 44 – CP 產線增加自動化作業

    '--------------------
    ' Condition Checking
    '--------------------

    '-------
    ' Action
    '-------

    If chkSchedualItem(2).Value = vbUnchecked Then
        GoTo ExitHandler
    End If

    '--------------------
    'Step1. 先清除已不符合條件機台上的待測資訊
    '--------------------

    '找出有待測資訊的機台
    sSQL = " select distinct " & gsCAT_TCMR_EQID & " " & _
             " from " & gsCAT_TBL_CPETEST_MONITOR_RECEIVE & " " & _
            " where " & gsCAT_TCMR_MONITORITEM & " = '" & gsMONITOR_ITEM_TESTDATA & "' " & _
              " and " & gsCAT_TCMR_MONITORITEMCONT & " = '" & gsMONITOR_ITEM_CONT_SORTRUN & "' " & _
              " and " & gsCAT_TCMR_MONITORITEMCONTDESC & " = '" & gsMONITOR_ITEM_DESC_CHANGELOT & "' " & _
              " and " & gsCAT_TCMR_DELETEFLAG & " = 'N' "
    Set colRaws = moProRawSql.QueryDatabase(sSQL)
    If colRaws.Count > 0 Then
        Call TransferLogTrace(DateTime.Now & ": Clear waits lots.....", _
                              moAppLog, Me.rtfTrace)
        For Each oRaws In colRaws
            sEqId = oRaws.Item(gsCAT_TCMR_EQID)
            '確定是否已不符合條件(主要是時間 & 查不到資料)
            '另外, 如果Level1的資料消失, 也會跟著清掉
            
            'Mark by Weilun on 20200513 for Project.BE MES Phase 44 – CP 產線增加自動化作業 <start>
'            sCheckSQL = " select " & gsCAT_TCML1_EQID & " " & _
'                          " from " & gsCAT_TBL_CPETEST_MONITOR_LEVEL1 & " t " & _
'                         " where ((" & gsCAT_TCML1_EXPECTENDTIME & " is not null and " & _
'                               "  (sysdate + 1 / 48) > " & _
'                               "  to_date(SubStr(" & gsCAT_TCML1_EXPECTENDTIME & ", 1, 15), 'YYYYMMDD HH24MISS')) " & _
'                               " or " & gsCAT_TCML1_EQSTATUS & " = '" & gsEQPSTATUS_IDLE & "') " & _
'                           " and " & gsCAT_TCML1_EQID & " = '" & sEqID & "' " & _
'                           " and " & gsCAT_TCML1_DELETEFLAG & " = 'N' "
            'Mark by Weilun on 20200513 for Project.BE MES Phase 44 – CP 產線增加自動化作業 <end>
                           
            'Add by Weilun on 20200513 for Project.BE MES Phase 44 – CP 產線增加自動化作業 <start>
            '加入eqptype檢查
            sCheckSQL = " select tcml1." & gsCAT_TCML1_EQID & " " & _
                          " from " & gsCAT_TBL_CPETEST_MONITOR_LEVEL1 & " tcml1, " & _
                                 " fweqpcurrentstate eqp " & _
                         " where ((tcml1." & gsCAT_TCML1_EXPECTENDTIME & " is not null and " & _
                                " (sysdate + 1 / 48) > " & _
                                " to_date(SubStr(tcml1." & gsCAT_TCML1_EXPECTENDTIME & ", 1, 15), 'YYYYMMDD HH24MISS')) " & _
                               " or tcml1." & gsCAT_TCML1_EQSTATUS & " = '" & gsEQPSTATUS_IDLE & "') " & _
                           " and eqp.activeversion >= 0 " & _
                           " and eqp.eqptype in ('" & Replace(sEqTypeList, ",", "','") & "') " & _
                           " and tcml1." & gsCAT_TCML1_EQID & " = eqp.equipmentname " & _
                           " and tcml1." & gsCAT_TCML1_EQID & " = '" & sEqId & "' " & _
                           " and tcml1." & gsCAT_TCML1_DELETEFLAG & " = 'N' "
            'Add by Weilun on 20200513 for Project.BE MES Phase 44 – CP 產線增加自動化作業 <end>
                           
            Set colCheckRaws = moProRawSql.QueryDatabase(sCheckSQL)
            If colCheckRaws.Count = 0 Then
                Call TransferLogTrace(DateTime.Now & ": Clear " & sEqId & " waits lots.", _
                                      moAppLog, Me.rtfTrace)

                '刪除該機台下的所有待測LOT資訊(另一種作法: 砍乾淨以避免積太多)
                sLotSQL = " update " & gsCAT_TBL_CPETEST_MONITOR_RECEIVE & " " & _
                             " set " & gsCAT_TCMR_DELETEFLAG & " = 'Y', " & _
                                       gsCAT_TCMR_DELETEUSERID & " = 'SYS', " & _
                                       gsCAT_TCMR_DELETETIME & " = to_char(sysdate, 'yyyymmdd hh24miss') || '000' " & _
                           " where " & gsCAT_TCMR_EQID & " = '" & sEqId & "' " & _
                             " and " & gsCAT_TCMR_MONITORITEM & " = '" & gsMONITOR_ITEM_TESTDATA & "' " & _
                             " and " & gsCAT_TCMR_MONITORITEMCONT & " = '" & gsMONITOR_ITEM_CONT_SORTRUN & "' " & _
                             " and " & gsCAT_TCMR_MONITORITEMCONTDESC & " = '" & gsMONITOR_ITEM_DESC_CHANGELOT & "' " & _
                             " and " & gsCAT_TCMR_DELETEFLAG & " = 'N' "
                Call moProRawSql.QueryDatabase(sLotSQL)
            End If
        Next
    End If

    '--------------------
    'Step2. 待測資訊新增與維護
    '--------------------

    '找出預計30分鐘內測完或是IDLE的機台
    
    'Mark by Weilun on 20200513 for Project.BE MES Phase 44 – CP 產線增加自動化作業 <start>
'    sEqIDList = ""
'    sSQL = " select " & gsCAT_TCML1_EQID & " " & _
'             " from " & gsCAT_TBL_CPETEST_MONITOR_LEVEL1 & " t " & _
'            " where ((" & gsCAT_TCML1_EXPECTENDTIME & " is not null and " & _
'                   " (sysdate + 1 / 48) > " & _
'                   " to_date(SubStr(" & gsCAT_TCML1_EXPECTENDTIME & ", 1, 15), 'YYYYMMDD HH24MISS')) " & _
'                  " or tcml1." & gsCAT_TCML1_EQSTATUS & " = '" & gsEQPSTATUS_IDLE & "') " & _
'              " and " & gsCAT_TCML1_DELETEFLAG & " = 'N' "
    'Mark by Weilun on 20200513 for Project.BE MES Phase 44 – CP 產線增加自動化作業 <end>

    'Add by Weilun on 20200513 for Project.BE MES Phase 44 – CP 產線增加自動化作業 <start>
    '加入eqptype檢查
    sEqIDList = ""

    sSQL = " select tcml1." & gsCAT_TCML1_EQID & " " & _
             " from " & gsCAT_TBL_CPETEST_MONITOR_LEVEL1 & " tcml1, " & _
                    " fweqpcurrentstate eqp " & _
            " where ((tcml1." & gsCAT_TCML1_EXPECTENDTIME & " is not null and " & _
                   " (sysdate + 1 / 48) > " & _
                   " to_date(SubStr(tcml1." & gsCAT_TCML1_EXPECTENDTIME & ", 1, 15), 'YYYYMMDD HH24MISS')) " & _
                  " or tcml1." & gsCAT_TCML1_EQSTATUS & " = '" & gsEQPSTATUS_IDLE & "') " & _
              " and eqp.activeversion >= 0 " & _
              " and eqp.eqptype in ('" & Replace(sEqTypeList, ",", "','") & "') " & _
              " and tcml1." & gsCAT_TCML1_EQID & " = eqp.equipmentname " & _
              " and tcml1." & gsCAT_TCML1_DELETEFLAG & " = 'N' " & _
         " order by tcml1." & gsCAT_TCML1_EQID & " "
    'Add by Weilun on 20200513 for Project.BE MES Phase 44 – CP 產線增加自動化作業 <end>
   
    Set colRaws = moProRawSql.QueryDatabase(sSQL)
    If colRaws.Count > 0 Then
        For Each oRaws In colRaws
            sEqIDList = sEqIDList & "," & oRaws.Item("EQID")
        Next
    End If
    If Left(sEqIDList, 1) = "," Then
        sEqIDList = Mid(sEqIDList, 2, Len(sEqIDList) - 1)
    End If

    If Trim(sEqIDList) <> "" Then 'EQ個別查詢時間是2~5秒 * n, 寫在一起能夠減少時間, 但是也有可能突然累加
        Call TransferLogTrace(DateTime.Now & ": Show waits lots for: " & sEqIDList, _
                              moAppLog, Me.rtfTrace)

        sEqIDListSQL = " ('" & Replace(sEqIDList, ",", "','") & "') "
        sLotSQL = " select tla." & gsCAT_TLATT_LOTID & ", " & _
                         " tla." & gsCAT_TLATT_STEPNAME & ", " & _
                         " tla." & gsCAT_TLATT_LOTSTATUS1 & ", " & _
                         " f19.SplitAssignEqId As EQID " & _
                    " from " & gsCAT_TBL_LOT_ATTRIBUTE & " tla, " & _
                           " fwlot b, " & _
                           " (SELECT FROMID, VALDATA SplitAssignEqId " & _
                              " From FWLOT_PN2M " & _
                             " WHERE KEYDATA = 'SplitAssignEqId') F19 " & _
                   " where f19.SplitAssignEqId in " & sEqIDListSQL & " " & _
                     " and b.sysid = f19.fromid(+) " & _
                     " and b.componentunits = '" & gsLOTUNIT_WAFER & "' " & _
                     " and tla." & gsCAT_TLATT_LOTID & " = b.appid " & _
                     " and tla." & gsCAT_TLATT_STEPNAME & " like 'SORT%' " & _
                     " and tla." & gsCAT_TLATT_RESOURCETYPE & " = 'WTESTER' " & _
                     " and tla." & gsCAT_TLATT_LOTSTATUS1 & " in ('" & gsLOTSTATUS_BREAK & "'," & _
                                                                " '" & gsLOTSTATUS_WAITR & "') " & _
                     " and tla." & gsCAT_TLATT_STAGE & " = '" & gsSTAGE_WS & "' "
        Set colLotRaws = moProRawSql.QueryDatabase(sLotSQL) '約3秒
        If colRaws.Count > 0 And colLotRaws.Count > 0 Then
            For Each oRaws In colRaws
                sEqId = oRaws.Item(gsCAT_TCML1_EQID)

                sExistLotIDlist = ""
                sCurLotIdList = ""
                sAddLotIDList = ""
                sRemoveLotIDList = ""
                Erase asAddList
                Erase asRemoveList

                '搜尋已在Monitor上的LotID並列表
                sCheckSQL = " select " & gsCAT_TCMR_LOTID & " " & _
                              " from " & gsCAT_TBL_CPETEST_MONITOR_RECEIVE & " " & _
                             " where " & gsCAT_TCMR_EQID & " = '" & sEqId & "' " & _
                               " and " & gsCAT_TCMR_MONITORITEM & " = '" & gsMONITOR_ITEM_TESTDATA & "' " & _
                               " and " & gsCAT_TCMR_MONITORITEMCONT & " = '" & gsMONITOR_ITEM_CONT_SORTRUN & "' " & _
                               " and " & gsCAT_TCMR_MONITORITEMCONTDESC & " = '" & gsMONITOR_ITEM_DESC_CHANGELOT & "' " & _
                               " and " & gsCAT_TCMR_DELETEFLAG & " = 'N' "
                Set colCheckRaws = moProRawSql.QueryDatabase(sCheckSQL)
                If colCheckRaws.Count > 0 Then
                    For Each oCheckRaws In colCheckRaws
                        sExistLotIDlist = sExistLotIDlist & "," & oCheckRaws.Item(gsCAT_TCMR_LOTID)
                    Next
                    If Left(sExistLotIDlist, 1) = "," Then
                        sExistLotIDlist = Mid(sExistLotIDlist, 2, Len(sExistLotIDlist) - 1)
                    End If
                End If

                '組建現在在機台上的LotList
                For Each oLotRaws In colLotRaws
                    If oLotRaws.Item("EQID") = sEqId Then
                        sLotId = oLotRaws.Item(gsCAT_TLATT_LOTID)
                        sStepName = oLotRaws.Item(gsCAT_TLATT_STEPNAME)
                    
                        bAddLotID = True
                    
                        If oLotRaws.Item(gsCAT_TLATT_LOTSTATUS1) = gsLOTSTATUS_BREAK Then '為Break時要檢查其CP結批的狀態
                            '檢查其CP結批的狀態
                            sCPSQL = " select " & gsCAT_TCAC_LOTENDSTATUS & ", " & _
                                                  gsCAT_TCAC_LOTCOMPLETESTATUS & ", " & _
                                                  gsCAT_TCAC_UPDATETIME & " " & _
                                       " from " & gsCAT_TBL_CP_AUTO_COMPLETE & " " & _
                                      " where (" & gsCAT_TCAC_LOTID & ", " & _
                                                   gsCAT_TCAC_STEPNAME & ", " & _
                                                   gsCAT_TCAC_CREATETIME & ") in " & _
                                            " (select " & gsCAT_TCAC_LOTID & ", " & _
                                                          gsCAT_TCAC_STEPNAME & ", " & _
                                                " Max(" & gsCAT_TCAC_CREATETIME & ") " & _
                                               " from " & gsCAT_TBL_CP_AUTO_COMPLETE & " " & _
                                              " where " & gsCAT_TCAC_LOTID & " = '" & sLotId & "' " & _
                                                " and " & gsCAT_TCAC_STEPNAME & " = '" & sStepName & "' " & _
                                           " group by " & gsCAT_TCAC_LOTID & ", " & gsCAT_TCAC_STEPNAME & ") " & _
                                        " and " & gsCAT_TCAC_DELETEFLAG & " = 'N' "
                            Set colCPRaws = moProRawSql.QueryDatabase(sCPSQL)
                            For Each oCPRaws In colCPRaws
                                '為都測完的Case, 自動以LotEnd為主, 手動以LotComplete為主
                                If oCPRaws.Item(gsCAT_TCAC_LOTENDSTATUS) = gsCPAUTO_STATUS_AUTO_OK Or _
                                   oCPRaws.Item(gsCAT_TCAC_LOTCOMPLETESTATUS) = gsCPAUTO_STATUS_MANUAL_OK Then
                                    bAddLotID = False '剛做完的不加入LotList
                                End If
                            Next
                        End If
                        
                        If bAddLotID = True Then '加入機台上的LotList
                            sCurLotIdList = sCurLotIdList & "," & sLotId
                        End If
                    End If
                Next
                If Left(sCurLotIdList, 1) = "," Then
                    sCurLotIdList = Mid(sCurLotIdList, 2, Len(sCurLotIdList) - 1)
                End If

                sAddLotIDList = RemoveStrInArray(sCurLotIdList, sExistLotIDlist, ",", ",") '要新增的
                sRemoveLotIDList = RemoveStrInArray(sExistLotIDlist, sCurLotIdList, ",", ",") '要移除的

                asAddList = Split(sAddLotIDList, ",")
                asRemoveList = Split(sRemoveLotIDList, ",")

                '要新增的
                If UBound(asAddList) > -1 Then
                    For lIdx = LBound(asAddList) To UBound(asAddList)
                        sLotId = asAddList(lIdx)

                        Call TransferLogTrace(DateTime.Now & ": List " & sLotId & " to " & sEqId, _
                                              moAppLog, Me.rtfTrace)

                        Call InsertMonitorReceive(sEqId, _
                                                  sLotId, _
                                                  "NA", _
                                                  "", _
                                                  gsMONITOR_ITEM_TESTDATA, _
                                                  gsMONITOR_ITEM_CONT_SORTRUN, _
                                                  gsMONITOR_ITEM_DESC_CHANGELOT, _
                                                  "99", _
                                                  gsCAT_TCMR_PRODTAFLAG)
                    Next
                End If

                '要移除的
                If UBound(asRemoveList) > -1 Then
                    For lIdx = LBound(asRemoveList) To UBound(asRemoveList)

                        sLotId = asRemoveList(lIdx)
                        Call TransferLogTrace(DateTime.Now & ": Clear " & sLotId & " from " & sEqId, _
                                              moAppLog, Me.rtfTrace)

                        sModifySQL = " update " & gsCAT_TBL_CPETEST_MONITOR_RECEIVE & " " & _
                                        " set " & gsCAT_TCMR_DELETEFLAG & " = 'Y', " & _
                                            " " & gsCAT_TCMR_DELETEUSERID & " = 'SYS', " & _
                                            " " & gsCAT_TCMR_DELETETIME & " = to_char(sysdate, 'yyyymmdd hh24miss') || '000' " & _
                                      " where " & gsCAT_TCMR_EQID & " = '" & sEqId & "' " & _
                                        " and " & gsCAT_TCMR_LOTID & " = '" & sLotId & "' " & _
                                        " and " & gsCAT_TCMR_MONITORITEM & " = '" & gsMONITOR_ITEM_TESTDATA & "' " & _
                                        " and " & gsCAT_TCMR_MONITORITEMCONT & " = '" & gsMONITOR_ITEM_CONT_SORTRUN & "' " & _
                                        " and " & gsCAT_TCMR_MONITORITEMCONTDESC & " = '" & gsMONITOR_ITEM_DESC_CHANGELOT & "' " & _
                                        " and " & gsCAT_TCMR_DELETEFLAG & " = 'N' "
                        Call moProRawSql.QueryDatabase(sModifySQL)
                    Next
                End If
            Next
        End If
        Call TransferLogTrace(DateTime.Now & ": Show waits lots complete.", _
                              moAppLog, Me.rtfTrace)
    End If

    '各個更新檢查都完成後, 再更新一二階Table, 再通知前端
    '-------
    'Done
    '-------
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
            Case Else
                typErrInfo.sUserText = typErrInfo.sErrDescription
        End Select
        '---- Start of Select Case Block ----
        On Error GoTo ExitHandler
        Call TransferLogError(typErrInfo.sUserText, _
                              moAppLog, Me.rtfTrace, msMODULE_ID, sProcID)
        Call HandleError(False, typErrInfo, , moAppLog)
    End If
End Sub

'================================================================================
'Name       : UnloadTypeRefresh()
'Description: 對於30分鐘內結束測試的機台, 更新其下貨資訊
'Author     : Weilun huang, MXIC 2020/11/26
'Parameters : N/A
'================================================================================
Private Sub UnloadTypeRefresh()
    On Error GoTo ExitHandler
    Dim sProcID As String
    Dim typErrInfo As tErrInfo

    Dim sSQL                As String
    Dim colRaws             As Collection
    Dim oRaws               As Collection

    Dim sModifySQL          As String 'input, update, delete
    Dim sLotId              As String
    Dim sEqId               As String
    Dim sEqTypeList         As String
    Dim sUnloadType         As String
    
    Dim oLot                As FwLot
    Dim oEquipment          As FwEquipment
    
    Dim bIsVirtualMerge     As Boolean
    '-----
    ' Init
    '-----
    sProcID = "NextLotPrepare"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog) '
    
    sEqTypeList = "WTESTER" 'Add by Weilun on 20200513 for Project.BE MES Phase 44 – CP 產線增加自動化作業

    '--------------------
    ' Condition Checking
    '--------------------

    '-------
    ' Action
    '-------

    If chkSchedualItem(4).Value = vbUnchecked Then
        GoTo ExitHandler
    End If

    Call TransferLogTrace(DateTime.Now & ": Unload Type Refresh start.", _
                          moAppLog, Me.rtfTrace)

    '找出預計30分鐘內測完的機台
    sSQL = " select tcml1." & gsCAT_TCML1_EQID & " " & _
             " from " & gsCAT_TBL_CPETEST_MONITOR_LEVEL1 & " tcml1, " & _
                    " fweqpcurrentstate eqp " & _
            " where tcml1." & gsCAT_TCML1_EXPECTENDTIME & " is not null and " & _
                  " (sysdate + 1 / 48) > " & _
                  " to_date(SubStr(tcml1." & gsCAT_TCML1_EXPECTENDTIME & ", 1, 15), 'YYYYMMDD HH24MISS') " & _
              " and eqp.activeversion >= 0 " & _
              " and eqp.eqptype in ('" & Replace(sEqTypeList, ",", "','") & "') " & _
              " and tcml1." & gsCAT_TCML1_EQID & " = eqp.equipmentname " & _
              " and tcml1." & gsCAT_TCML1_DELETEFLAG & " = 'N' " & _
         " order by tcml1." & gsCAT_TCML1_EQID & " "
   
    Set colRaws = moProRawSql.QueryDatabase(sSQL)
    If colRaws.Count > 0 Then
        For Each oRaws In colRaws
            '取得機台
            sEqId = oRaws.Item(gsCAT_TCML1_EQID)
            Set oEquipment = moFwMDL.EquipmentById(sEqId)
            If Not oEquipment Is Nothing Then

                '取得Lot物件
                'Modify by Weilun on 20230606 for Project.新增CP監控識別虛擬併批功能以及修改預測時間機制
'                sLotID = oEquipment.CustomAttributes(gsEQP_CUSTOMATTR_CUR_LOT_ID)
                Call GetEqDetailLotList(oEquipment.Id, sLotId, "", "", bIsVirtualMerge, "")
                
                Set oLot = moFwWIP.LotById(sLotId)
                If Not oLot Is Nothing Then
                    '取得下貨方式, 並更新一階資料
                    sUnloadType = GetUnloadType(sEqId, sLotId) '內建訊息(包含下一批的資訊)
                    
                    sModifySQL = " update " & gsCAT_TBL_CPETEST_MONITOR_LEVEL1 & " " & _
                                    " set " & gsCAT_TCML1_UNLOADTYPE & " = '" & sUnloadType & "', " & _
                                              gsCAT_TCML1_UPDATEUSERID & " = 'SYS', " & _
                                              gsCAT_TCML1_UPDATETIME & " = to_char(sysdate, 'yyyymmdd hh24miss') || '000' " & _
                                  " where " & gsCAT_TCML1_EQID & " = '" & sEqId & "' " & _
                                    " and " & gsCAT_TCML1_DELETEFLAG & " = 'N' "
                    Call moProRawSql.QueryDatabase(sModifySQL)
                End If
            End If
        Next
    End If
    
    Call TransferLogTrace(DateTime.Now & ": Unload Type Refresh end.", _
                          moAppLog, Me.rtfTrace)
    
    '各個更新檢查都完成後, 再更新一二階Table, 再通知前端
    '-------
    'Done
    '-------
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
            Case Else
                typErrInfo.sUserText = typErrInfo.sErrDescription
        End Select
        '---- Start of Select Case Block ----
        On Error GoTo ExitHandler
        Call TransferLogError(typErrInfo.sUserText, _
                              moAppLog, Me.rtfTrace, msMODULE_ID, sProcID)
        Call HandleError(False, typErrInfo, , moAppLog)
    End If
End Sub

'================================================================================
'Name       : CPAutoMonitor()
'Description: 原F5 CP Auto Monitor的紅字功能(步驟三不用)
'Author     : Weilun huang, MXIC 2019/09/25
'Parameters : N/A
'================================================================================
Private Sub CPAutoMonitor()
    On Error GoTo ExitHandler
    Dim sProcID As String
    Dim typErrInfo As tErrInfo

    Dim oLot As FwLot
    Dim sLotId As String
    Dim lIdx As Long
    Dim vPriority As Object

    Dim sSQL As String
    Dim colRaws As Collection
    Dim oRaws As Collection
    Dim colTime As Collection
    Dim sStatusList As String
    Dim sLotEndStatus As String
    Dim sLotCompleteStatus As String

    Dim sCheckSQL As String
    Dim colCheckRaws As Collection
    Dim oCheckRaws As Collection

    Dim sModifySQL As String 'input, update, delete

    Dim sMaxUpdateTime As String
    Dim sUpdateTime As String
    Dim sMainStatus As String
    Dim sRuleName As String

    Dim sEqId As String
    Dim sCheckResult As String
    Dim sScheduleTime As String
    Dim sCurrentTime As String
    Dim sLastAutoTime As String
    Dim sStepName As String
    Dim sExpiration As String
    Dim sExecutingLimit As String
    Dim sMoveOutLimit As String
    '----
    ' Init
    '----
    sProcID = "CPAutoMonitor"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog)

    sExecutingLimit = "1/144"   '10分鐘
    sMoveOutLimit = "1/288"     '5分鐘
    sExpiration = "1/720"       '2分鐘

    '----
    ' Condition Chking
    '----
    ' <Put your condition checking codes here>...

    If chkSchedualItem(3).Value = vbUnchecked Then
        GoTo ExitHandler
    End If


    '----
    ' Action
    '----

    Call TransferLogTrace(DateTime.Now & ": CP Auto Monitor check start.", _
                          moAppLog, Me.rtfTrace)

    '第一部份: 搜TBL_CP_AUTO_COMPLETE資料進行Query
    'Fail的部份已由CPAutoChange接手, 手動部份不管
    sStatusList = gsCPAUTO_STATUS_INIT & "," & _
                  gsCPAUTO_STATUS_AUTO_EXEC

    sStatusList = "'" & Replace(sStatusList, ",", "','") & "'"

    'Modify by Weilun on 20180808 for ReqNo.BE#201800559
    '加入比對Lot_Attribute資訊, 濾掉不合時宜的資訊
    sSQL = " select TCAC." & gsCAT_TCAC_LOTID & ", " & _
                  " TCAC." & gsCAT_TCAC_STEPNAME & ", " & _
                  " TCAC." & gsCAT_TCAC_COMMENTS & ", " & _
                  " TCAC." & gsCAT_TCAC_LOTENDSTATUS & ", " & _
                  " TCAC." & gsCAT_TCAC_LOTCOMPLETESTATUS & ", " & _
                  " TCAC." & gsCAT_TCAC_CREATETIME & ", " & _
                  " TCAC." & gsCAT_TCAC_UPDATETIME & " "

    sSQL = sSQL & _
             " from " & gsCAT_TBL_CP_AUTO_COMPLETE & " TCAC , " & _
                        gsCAT_TBL_LOT_ATTRIBUTE & " TLATT "

    sSQL = sSQL & _
            " where ( TCAC." & gsCAT_TCAC_LOTID & ", " & _
                    " TCAC." & gsCAT_TCAC_STEPNAME & ", " & _
                    " TCAC." & gsCAT_TCAC_CREATETIME & ") in " & _
                  " (select " & gsCAT_TCAC_LOTID & ", " & _
                                gsCAT_TCAC_STEPNAME & ", " & _
                      " Max(" & gsCAT_TCAC_CREATETIME & ") " & _
                     " from " & gsCAT_TBL_CP_AUTO_COMPLETE & " " & _
                    " where " & gsCAT_TCAC_DELETEFLAG & " = 'N' " & _
                 " group by " & gsCAT_TCAC_LOTID & ", " & gsCAT_TCAC_STEPNAME & ") " & _
              " and ( TCAC." & gsCAT_TCAC_LOTENDSTATUS & " in (" & sStatusList & ")" & _
                 " or TCAC." & gsCAT_TCAC_LOTCOMPLETESTATUS & " in (" & sStatusList & "))" & _
              " and TCAC." & gsCAT_TCAC_DELETEFLAG & " = 'N' " & _
              " and TCAC." & gsCAT_TCAC_LOTID & " = TLATT." & gsCAT_TLATT_LOTID & " " & _
              " and TCAC." & gsCAT_TCAC_STEPNAME & " = TLATT." & gsCAT_TLATT_STEPNAME & " " & _
              " and TLATT." & gsCAT_TLATT_LOTSTATUS1 & " in ('" & gsLOTSTATUS_RUN & "','" & _
                                                                  gsLOTSTATUS_ERUN & "','" & _
                                                                  gsLOTSTATUS_BREAK & "') " & _
         " order by TCAC." & gsCAT_TCAC_CREATETIME & " "


    Set colRaws = moProRawSql.QueryDatabase(sSQL)
    For Each oRaws In colRaws
        DoEvents

        sCheckResult = ""
        sRuleName = ""
        sMaxUpdateTime = ""
        sMainStatus = ""

        sLotEndStatus = oRaws.Item(gsCAT_TCAC_LOTENDSTATUS)
        sLotCompleteStatus = oRaws.Item(gsCAT_TCAC_LOTCOMPLETESTATUS)

        'LotCompleteStatus優先, 並取得狀態
        If sLotCompleteStatus <> "" Then
            sMainStatus = sLotCompleteStatus
            sRuleName = msComplete
        Else
            sMainStatus = sLotEndStatus
            sRuleName = msEnd
        End If

        '狀態分辨
        '狀態1: CPAuto已超過預定時間(sExecutingLimit), 卻還在執行中, 請介入
        If sMainStatus = gsCPAUTO_STATUS_AUTO_EXEC Then
            sUpdateTime = oRaws.Item(gsCAT_TCAC_UPDATETIME)
            sSQL = " select to_char(to_date(SubStr('" & sUpdateTime & "', 1, 15), 'YYYYMMDD HH24MISS') + " & sExecutingLimit & ", " & _
                                  " 'YYYYMMDD HH24MISS') || '000' as ScheduleTime, " & _
                          " to_char(sysdate, 'YYYYMMDD HH24MISS') || '000' as CurrentTime " & _
                     " From Dual "
            Set colTime = moProRawSql.QueryDatabase(sSQL)
            If colTime.Count = 1 Then
                sScheduleTime = colTime.Item(1).Item("ScheduleTime")
                sCurrentTime = colTime.Item(1).Item("CurrentTime")
                If sCurrentTime > sScheduleTime Then
                    sCheckResult = "Please execute " & sRuleName & "."
                End If
            End If


            '狀態2: 資料到齊已超過特定時間(sExpiration), 卻尚未自動化(狀態還是Init)
        ElseIf sMainStatus = gsCPAUTO_STATUS_INIT Then
            Set oLot = moFwWIP.LotById(oRaws.Item(gsCAT_TCAC_LOTID))

            If Not oLot Is Nothing Then
                If CheckTdsTestSumComplete(oLot, sMaxUpdateTime) = True Then
                    sSQL = " select to_char(to_date(SubStr('" & sMaxUpdateTime & "', 1, 15), 'YYYYMMDD HH24MISS') + " & sExpiration & ", " & _
                                          " 'YYYYMMDD HH24MISS') || '000' as ScheduleTime, " & _
                                  " to_char(sysdate, 'YYYYMMDD HH24MISS') || '000' as CurrentTime " & _
                             " From Dual "
                    Set colTime = moProRawSql.QueryDatabase(sSQL)
                    If colTime.Count = 1 Then
                        sScheduleTime = colTime.Item(1).Item("ScheduleTime")
                        sCurrentTime = colTime.Item(1).Item("CurrentTime")
                        If sCurrentTime > sScheduleTime Then
                            sCheckResult = "Please execute " & sRuleName & "."
                        End If
                    End If
                End If
            End If
        End If

        If sCheckResult <> "" Then
            Set oLot = moFwWIP.LotById(oRaws.Item(gsCAT_TCAC_LOTID))
            If Not oLot Is Nothing Then
                sLotId = oLot.Id
                sEqId = oLot.CustomAttributes(gsLOT_CUSTOMATTR_CUR_EQP_ID)
                If sEqId = "" And oLot.CustomAttributes(gsLOT_CUSTOMATTR_STATUS) = gsLOTSTATUS_BREAK Then 'Break的條件建議注意一下
                    sEqId = oLot.CustomAttributes(gsLOT_CUSTOMATTR_LAST_EQP_ID)
                End If
    
                If sEqId <> "" Then
                    If oRaws.Item(gsCAT_TCAC_UPDATETIME) <> "" Then
                        sLastAutoTime = oRaws.Item(gsCAT_TCAC_UPDATETIME)
                    Else
                        sLastAutoTime = oRaws.Item(gsCAT_TCAC_CREATETIME)
                    End If
    
                    sStepName = oRaws.Item(gsCAT_TCAC_STEPNAME)
    
                    '是否有重複的資料
                    sCheckSQL = " select " & gsCAT_TCMR_CREATETIME & " " & _
                                   " from " & gsCAT_TBL_CPETEST_MONITOR_RECEIVE & " " & _
                                  " where " & gsCAT_TCMR_CREATETIME & " > '" & sLastAutoTime & "' " & _
                                    " and " & gsCAT_TCMR_STEPNAME & " = '" & sStepName & "' " & _
                                    " and " & gsCAT_TCMR_EQID & " = '" & sEqId & "' " & _
                                    " and " & gsCAT_TCMR_LOTID & " = '" & sLotId & "' " & _
                                    " and " & gsCAT_TCMR_MONITORITEM & " = '" & gsMONITOR_ITEM_TESTDATA & "' " & _
                                    " and " & gsCAT_TCMR_MONITORITEMCONT & " = '" & gsMONITOR_ITEM_CONT_AUTOMONITOR & "' " & _
                                    " and " & gsCAT_TCMR_MONITORITEMCONTDESC & " = '" & sCheckResult & "' " & _
                                    " and " & gsCAT_TCMR_DELETEFLAG & " = 'N' "
                    Set colCheckRaws = moProRawSql.QueryDatabase(sCheckSQL)
    
                    If colCheckRaws.Count = 0 Then '沒有該時間區段的資料才加入
                        Call TransferLogTrace(DateTime.Now & ": " & sLotId & " " & sCheckResult & " on " & sEqId, _
                                              moAppLog, Me.rtfTrace)
    
                        Call InsertMonitorReceive(sEqId, _
                                                  sLotId, _
                                                  "NA", _
                                                  sStepName, _
                                                  gsMONITOR_ITEM_TESTDATA, _
                                                  gsMONITOR_ITEM_CONT_AUTOMONITOR, _
                                                  sCheckResult, _
                                                  "1", _
                                                  gsCAT_TCMR_PRODTAFLAG)
                    End If
                End If
            End If
        End If
    Next oRaws

    Call TransferLogTrace(DateTime.Now & ": CP Auto Monitor check: MoveOut timeout.", _
                          moAppLog, Me.rtfTrace)

    '第二部份: MoveOut檢查
    sSQL = " select  t." & gsCAT_TLATT_LOTID & ", " & _
                  " cp." & gsCAT_TCAC_LOTENDSTATUS & ", " & _
                  " cp." & gsCAT_TCAC_LOTCOMPLETESTATUS & ", " & _
                  " cp." & gsCAT_TCAC_COMMENTS & ", " & _
                  " cp." & gsCAT_TCAC_UPDATETIME & " "

    sSQL = sSQL & _
             " from " & gsCAT_TBL_LOT_ATTRIBUTE & " t, " & _
                      " fwlot                a, " & _
                      " fwlot_n2m            b, " & _
                      " fwwipstep            c, " & _
                      " fwwipstep_n2m        d, " & _
                      " fwwipsteprule        e, " & _
                        gsCAT_TBL_CP_AUTO_COMPLETE & " cp "

    sSQL = sSQL & _
            " where t." & gsCAT_TLATT_STAGE & " = '" & gsSTAGE_WS & "' " & _
              " and t." & gsCAT_TLATT_LOTSTATUS1 & " in ('" & gsLOTSTATUS_WAITD & "', " & _
                                                       " '" & gsLOTSTATUS_BREAK & "') " & _
              " and t." & gsCAT_TLATT_RESOURCETYPE & " = '" & "WTESTER" & "' " & _
              " and t." & gsCAT_TLATT_LOTID & " = cp." & gsCAT_TCAC_LOTID & " " & _
              " and t." & gsCAT_TLATT_STEPNAME & " = cp." & gsCAT_TCAC_STEPNAME & " " & _
              " and (cp." & gsCAT_TCAC_LOTID & ", " & _
                   " cp." & gsCAT_TCAC_STEPNAME & ", " & _
                   " cp." & gsCAT_TCAC_CREATETIME & ") in " & _
                  " (select subcp." & gsCAT_TCAC_LOTID & ", " & _
                          " subcp." & gsCAT_TCAC_STEPNAME & ", " & _
                          " Max(subcp." & gsCAT_TCAC_CREATETIME & ") " & _
                     " from " & gsCAT_TBL_CP_AUTO_COMPLETE & " subcp " & _
                    " where subcp." & gsCAT_TCAC_DELETEFLAG & " = 'N' " & _
                 " group by subcp." & gsCAT_TCAC_LOTID & ", " & _
                          " subcp." & gsCAT_TCAC_STEPNAME & ") "

    sSQL = sSQL & _
              " and cp." & gsCAT_TCAC_LOTCOMPLETESTATUS & " in ('" & gsCPAUTO_STATUS_AUTO_OK & "', " & _
                                                              " '" & gsCPAUTO_STATUS_MANUAL_OK & "') " & _
              " and t." & gsCAT_TLATT_LOTID & " = a.appid " & _
              " and a.sysid = b.fromid " & _
              " and b.toid = c.sysid " & _
              " and c.sysid = d.fromid " & _
              " and d.toid = e.sysid " & _
              " and c.currentruleindex = d.sequence " & _
              " and e.rulename = 'MoveOut' "

    Set colRaws = moProRawSql.QueryDatabase(sSQL)
    For Each oRaws In colRaws
        Set oLot = moFwWIP.LotById(oRaws.Item(gsCAT_TLATT_LOTID))
        If Not oLot Is Nothing Then
            sLotId = oLot.Id
            'LotComplete已超過預定時間(sMoveOutLimit), 卻還在MoveOut
            sUpdateTime = oRaws.Item(gsCAT_TCAC_UPDATETIME)
            sSQL = " select to_char(to_date(SubStr('" & sUpdateTime & "', 1, 15), 'YYYYMMDD HH24MISS') + " & sMoveOutLimit & ", " & _
                                  " 'YYYYMMDD HH24MISS') || '000' as ScheduleTime, " & _
                          " to_char(sysdate, 'YYYYMMDD HH24MISS') || '000' as CurrentTime " & _
                     " From Dual "
            Set colTime = moProRawSql.QueryDatabase(sSQL)
            If colTime.Count = 1 Then
                sScheduleTime = colTime.Item(1).Item("ScheduleTime")
                sCurrentTime = colTime.Item(1).Item("CurrentTime")
                If sCurrentTime > sScheduleTime Then
                    sCheckResult = "Please execute move out forcibly " & vbNewLine & _
                                   "after confirm CommMgm system [rework] " & vbNewLine & _
                                   "record were done."
                    sEqId = oLot.CustomAttributes(gsLOT_CUSTOMATTR_CUR_EQP_ID)
                    If sEqId = "" And (oLot.CustomAttributes(gsLOT_CUSTOMATTR_STATUS) = gsLOTSTATUS_BREAK Or oLot.CustomAttributes(gsLOT_CUSTOMATTR_STATUS) = gsLOTSTATUS_WAITD) Then
                        sEqId = oLot.CustomAttributes(gsLOT_CUSTOMATTR_LAST_EQP_ID)
                    End If

                    If sEqId <> "" Then
                        sStepName = oLot.CurrentStep.Steps.Item(1).Description

                        '是否有重複的資料
                        sCheckSQL = " select " & gsCAT_TCMR_CREATETIME & " " & _
                                       " from " & gsCAT_TBL_CPETEST_MONITOR_RECEIVE & " " & _
                                      " where " & gsCAT_TCMR_CREATETIME & " > '" & sUpdateTime & "' " & _
                                        " and " & gsCAT_TCMR_STEPNAME & " = '" & sStepName & "' " & _
                                        " and " & gsCAT_TCMR_EQID & " = '" & sEqId & "' " & _
                                        " and " & gsCAT_TCMR_LOTID & " = '" & sLotId & "' " & _
                                        " and " & gsCAT_TCMR_MONITORITEM & " = '" & gsMONITOR_ITEM_TESTDATA & "' " & _
                                        " and " & gsCAT_TCMR_MONITORITEMCONT & " = '" & gsMONITOR_ITEM_CONT_AUTOMONITOR & "' " & _
                                        " and " & gsCAT_TCMR_MONITORITEMCONTDESC & " = '" & sCheckResult & "' " & _
                                        " and " & gsCAT_TCMR_DELETEFLAG & " = 'N' "
                        Set colCheckRaws = moProRawSql.QueryDatabase(sCheckSQL)

                        If colCheckRaws.Count = 0 Then '沒有該時間區段的資料才加入
                            Call TransferLogTrace(DateTime.Now & ": " & sLotId & " " & sCheckResult & " on " & sEqId, _
                                                  moAppLog, Me.rtfTrace)

                            Call InsertMonitorReceive(sEqId, _
                                                      sLotId, _
                                                      "NA", _
                                                      sStepName, _
                                                      gsMONITOR_ITEM_TESTDATA, _
                                                      gsMONITOR_ITEM_CONT_AUTOMONITOR, _
                                                      sCheckResult, _
                                                      "1", _
                                                      gsCAT_TCMR_PRODTAFLAG)
                        End If
                    End If
                End If
            End If
        End If
        Set oLot = Nothing

    Next oRaws

    Call TransferLogTrace(DateTime.Now & ": CP Auto Monitor check end.", _
                          moAppLog, Me.rtfTrace)

    '-------
    'Done
    '-------
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
            Case Else
                typErrInfo.sUserText = typErrInfo.sErrDescription
        End Select
        '---- Start of Select Case Block ----
        On Error GoTo ExitHandler
        Call TransferLogError(typErrInfo.sUserText, _
                              moAppLog, Me.rtfTrace, msMODULE_ID, sProcID)
        Call HandleError(False, typErrInfo, , moAppLog)
    End If
End Sub

'================================================================================
'Name       : PEISUpdate()
'Description: 補齊PEIS的資料
'Author     : Weilun huang, MXIC 2019/10/15
'Parameters : N/A
'================================================================================
Private Sub PEISUpdate()
    On Error GoTo ExitHandler
    Dim sProcID As String
    Dim typErrInfo As tErrInfo

    Dim sModifySQL As String

    Dim sSID            As String
    Dim sLotId          As String
    Dim sEqId           As String
    Dim sWaferID        As String
    Dim sLotSeq         As String
    Dim sHotFlag        As String
    Dim sNextStepName   As String
    Dim sNextStepRecipe As String
    Dim sNewReviewFlag  As String

    Dim sSQL            As String
    Dim colRaws         As Collection
    Dim oRaws           As Collection
    Dim sOtherSet       As String
    Dim sExtraMsg       As String

    '-----
    ' Init
    '-----
    sProcID = "PEISUpdate"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog) '

    '--------------------
    ' Condition Checking
    '--------------------

    '-------
    ' Action
    '-------
    sSQL = " select " & gsCAT_TCMR_SID & ", " & _
                        gsCAT_TCMR_EQID & ", " & _
                        gsCAT_TCMR_LOTID & ", " & _
                        gsCAT_TCMR_WAFERID & ", " & _
                        gsCAT_TCMR_EQSTOPTEST & ", " & _
                        gsCAT_TCMR_EQSTOPTESTTYPE & " " & _
             " from " & gsCAT_TBL_CPETEST_MONITOR_RECEIVE & " " & _
            " where " & gsCAT_TCMR_DELETEFLAG & " = 'N' " & _
              " and " & gsCAT_TCMR_REVIEWFLAG & " = 'P' "
    Set colRaws = moProRawSql.QueryDatabase(sSQL)
    For Each oRaws In colRaws
        sSID = oRaws.Item(gsCAT_TCMR_SID)
        sEqId = oRaws.Item(gsCAT_TCMR_EQID)
        sLotId = oRaws.Item(gsCAT_TCMR_LOTID)
        
        sExtraMsg = ""
        sNewReviewFlag = "N" '預設值
        If Trim(oRaws.Item(gsCAT_TCMR_EQSTOPTEST)) = "Y" And _
           Trim(oRaws.Item(gsCAT_TCMR_EQSTOPTESTTYPE)) = gsSTOPTEST_TYPE_AUTO Then
           '自動停機的Case
            sNewReviewFlag = "S"
            sExtraMsg = ", Auto stop Test with ReviewFlag: " & sNewReviewFlag
        ElseIf Trim(oRaws.Item(gsCAT_TCMR_EQSTOPTEST)) = "Y" And _
               Trim(oRaws.Item(gsCAT_TCMR_EQSTOPTESTTYPE)) = gsSTOPTEST_TYPE_MANUAL Then
           '手動停機的Case
            sNewReviewFlag = "M"
            sExtraMsg = ", Manual stop Test with ReviewFlag: " & sNewReviewFlag
        End If

        '檢查是否有缺需要補值的狀況
        sOtherSet = ""
        If Trim(oRaws.Item(gsCAT_TCMR_WAFERID)) = "" Then
            sOtherSet = sOtherSet & gsCAT_TCMR_WAFERID & " = 'NA', "
        End If

        '取得LotSeq, 有就帶進去
        sLotSeq = GetWsLotSeq(sLotId, sEqId)

        '取得HotFlag
        sHotFlag = GetHotFlag(sLotId)

        '取得下一站, 如果是Bake會再取得Recipe
        Call GetNextStepNameAndRecipe(sLotId, sNextStepName, sNextStepRecipe)

        If IsNumeric(sSID) And sEqId <> "" And sLotId <> "" Then
            Call TransferLogTrace(DateTime.Now & ": Update PEIS data for SID:" & sSID & sExtraMsg, _
                                  moAppLog, Me.rtfTrace)

            sModifySQL = " update " & gsCAT_TBL_CPETEST_MONITOR_RECEIVE & " " & _
                            " set " & gsCAT_TCMR_LOTSEQ & " = '" & sLotSeq & "', " & _
                                      gsCAT_TCMR_HOTFLAG & " = '" & sHotFlag & "', " & _
                                      gsCAT_TCMR_NEXTSTEPNAME & " = '" & sNextStepName & "', " & _
                                      gsCAT_TCMR_RECIPE & " = '" & sNextStepRecipe & "', " & _
                                      sOtherSet & _
                                      gsCAT_TCMR_REVIEWFLAG & " = '" & sNewReviewFlag & "', " & _
                                      gsCAT_TCMR_UPDATEUSERID & " = 'SYS', " & _
                                      gsCAT_TCMR_UPDATETIME & " = to_char(sysdate, 'yyyymmdd hh24miss') || '000' " & _
                          " where " & gsCAT_TCMR_LOTID & " = '" & sLotId & "' " & _
                            " and " & gsCAT_TCMR_EQID & " = '" & sEqId & "' " & _
                            " and " & gsCAT_TCMR_DELETEFLAG & " = 'N' " & _
                            " and " & gsCAT_TCMR_SID & " = " & sSID & " "

            Call moProRawSql.QueryDatabase(sModifySQL)
        End If

    Next
    '-------
    'Done
    '-------
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
            Case Else
                typErrInfo.sUserText = typErrInfo.sErrDescription
        End Select
        '---- Start of Select Case Block ----
        On Error GoTo ExitHandler
        Call TransferLogError(typErrInfo.sUserText, _
                              moAppLog, Me.rtfTrace, msMODULE_ID, sProcID)
        Call HandleError(False, typErrInfo, , moAppLog)
    End If
End Sub

'================================================================================
'Name       : GetUph()
'Description: 取得每小時測是的數量
'Author     : Weilun huang, MXIC 2019/09/06
'Parameters : N/A
'================================================================================
Private Function GetUph(ByVal oLot As FwLot, _
                        ByVal oEquipment As FwEquipment, _
                        ByVal sStage As String, _
                        ByRef lWaitTest As Long) As Double
    On Error GoTo ExitHandler:
    Dim sProcID                     As String
    Dim typErrInfo                  As tErrInfo
    
    Dim sLotId                      As String
    Dim sEqId                       As String
    Dim dUph                        As Double
    Dim dProcTime                   As Double
    Dim sErunTicNO                  As String
    Dim sFollowProd                 As String
    
    Dim sEqType2                    As String
    Dim sProdgroup                  As String
    Dim sStepID                     As String
    Dim sWaferQty                   As String
    Dim sTimeHereSince              As String
    
    'Add by Weilun on 20191126 for Project. 電測自動化監控系統 <start>
    Dim sSubSysID                   As String
    Dim oSubSysID                   As FwEquipment
    Dim sSubSystem                  As String
    'Add by Weilun on 20191126 for Project. 電測自動化監控系統 <end>

    Dim sSQL                        As String
    Dim colRaws                     As Collection
    Dim oRaws                       As Collection
    
    Dim sPgName                     As String ' Add by Weilun on 20230608 for Project.新增CP監控識別虛擬併批功能以及修改預測時間機制
    
    Dim sProdcode As String 'Add by Sam on 20230821 for  ReqNo:202300189
    Dim sPkgCode As String 'Add by Sam on 20230821 for  ReqNo:202300189
    Dim sPincount As String 'Add by Sam on 20230821 for ReqNo:202300189
    Dim sStepName As String  'Add by Sam on 20230821 for ReqNo:202300189
    '-----
    ' Init
    '-----
    sProcID = "GetUph"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog) '
    
    GetUph = 0
    lWaitTest = 0
    
    '--------------------
    ' Condition Checking
    '--------------------
    
    '-------
    ' Action
    '-------
    If oEquipment Is Nothing Or oLot Is Nothing Then
        GoTo ExitHandler
    Else
        oEquipment.Refresh
        oLot.Refresh
        
        '基本資料取得
        sLotId = oLot.Id
        sEqId = oEquipment.Id
        sEqType2 = oEquipment.CustomAttributes(gsEQP_CUSTOMATTR_EQ_TYPE2)
        sPgName = oEquipment.CustomAttributes(gsEQP_CUSTOMATTR_CUR_PG_NAME) ' Add by Weilun on 20230608 for Project.新增CP監控識別虛擬併批功能以及修改預測時間機制
        sStepID = oLot.CurrentStep.Steps.Item(1).Id
        sWaferQty = oLot.CustomAttributes(gsLOT_CUSTOMATTR_WAFER_QTY)
        sTimeHereSince = ConvertToString(oLot.CurrentStep.Steps.Item(1).TimeHereSince, moAppLog)
        sProdgroup = oLot.CustomAttributes(gsLOT_CUSTOMATTR_PROD_GROUP)
        
        'Add by Weilun on 20191126 for Project. 電測自動化監控系統 <start>
        If oEquipment.CustomAttributes(gsEQP_CUSTOMATTR_SUBSYS1) <> "" Then '在CP只有一個, 基本也就拿第一個當SubSystem
            sSubSysID = oEquipment.CustomAttributes(gsEQP_CUSTOMATTR_SUBSYS1)
            Set oSubSysID = moFwMDL.EquipmentById(sSubSysID)
            If Not oSubSysID Is Nothing Then
                sSubSystem = oSubSysID.CustomAttributes(gsEQP_CUSTOMATTR_EQ_TYPE2)
            End If
        End If
        'Add by Weilun on 20191126 for Project. 電測自動化監控系統 <end>
        
        'Add by Sam Start  on 20230821 for ReqNo:202300189
        sStepName = oLot.CurrentStep.Steps.Item(1).Description
        sSQL = "select substr(a." & gsCAT_TIM_IPN & " ,1 ,4) || a." & gsCAT_TIM_MASK_OPTION & " as prodcode " & _
               " ,a." & gsCAT_TIM_PACKAGE_CODE & ",a." & gsCAT_TIM_PIN_COUNT & "  from " & gsCAT_TBL_IPN_MASTER & " a " & _
              " where a." & gsCAT_TIM_IPN & " = '" & oLot.CustomAttributes(gsLOT_CUSTOMATTR_IPN) & "' " & _
              " and a." & gsCAT_TIM_DELETE_FLAG & " ='N' "
        Set colRaws = moProRawSql.QueryDatabase(sSQL)
        If colRaws.Count > 0 Then
           sProdcode = colRaws.Item(1).Item("prodcode")
           sPkgCode = colRaws.Item(1).Item(gsCAT_TIM_PACKAGE_CODE)
           sPincount = colRaws.Item(1).Item(gsCAT_TIM_PIN_COUNT)
        End If
        'Add by Sam End on 20230821 for ReqNo:202300189
        
        sSQL = " select t." & gsCAT_TLI_ERUNTICNO & " " & _
                 " from " & gsCAT_TBL_LOT_INFO & " t " & _
                " where t." & gsCAT_TLI_LOT_ID & " = '" & sLotId & "' "
        Set colRaws = moProRawSql.QueryDatabase(sSQL)
        If colRaws.Count > 0 Then
            sErunTicNO = colRaws.Item(1).Item(gsCAT_TLI_ERUNTICNO)
        End If
        
        '取得未測片數
        lWaitTest = Val(sWaferQty) - GetTestedPcs(sStepID, sLotId, sEqId, sTimeHereSince)
        
        If sErunTicNO <> "" Then
            sSQL = " select t." & gsCAT_TER_FOLLOW_PRODUCT & " " & _
                     " from " & gsCAT_TBL_ERUN_REQ & " t " & _
                     "where t." & gsCAT_TER_LOT_ID & " = '" & sLotId & "' " & _
                      " and t." & gsCAT_TER_TICKET_NO & " = '" & sErunTicNO & "' " & _
                      " and t." & gsCAT_TER_STAGE & " = '" & sStage & "' " & _
                      " and t." & gsCAT_TER_DELETE_FLAG & " = 'N' "
            Set colRaws = moProRawSql.QueryDatabase(sSQL)
            If colRaws.Count > 0 Then
                sFollowProd = colRaws.Item(1).Item(gsCAT_TER_FOLLOW_PRODUCT)
            End If
        End If
        
        If sFollowProd = "N" Then
            '工程品不Follow暫時不計算(算的不正確), Modify by Weilun on 20191128 for Project. 電測自動化監控系統
            dUph = 0
'            Call GetErunRecipe(sErunTicNO, sStepID, sEqType2, "", "", "", "", "", "", dProcTime)
'            If dProcTime > 0 Then
'                If oLot.CurrentStep.Steps.Item(1).Description Like "SORT*" And _
'                   Mid(oLot.CurrentStep.Steps.Item(1).Id, 3, 1) <> "2" Then
'                    '抽測站
'
'                    dUph = 1 / dProcTime
'                Else
'                    '若狀態為WAITR且未測片數=0 表示此批可能是當站重工且已測試結束需重測的狀態,以整批數量計算
'                    If lWaitTest <= 0 And oLot.Status = gsLOTSTATUS_WAITR Then
'                       lWaitTest = Val(sWaferQty)
'                    End If
'
'                    dUph = lWaitTest / dProcTime
'                End If
'            Else
'                dUph = 0
'            End If
        Else
            '條件sPgName Add by Weilun on 20230608 for Project.新增CP監控識別虛擬併批功能以及修改預測時間機制
            sSQL = " SELECT A." & gsCAT_PSES_PG_NAME & ",A." & gsCAT_PSES_PROBECARD_TYPE & " " & _
                         " ,A." & gsCAT_PSES_TEMPERATURE & ",A." & gsCAT_PSES_WS_DEVICE_FILE & " " & _
                         " ,A." & gsCAT_PSES_SpecifyEq & ",A." & gsCAT_PSES_EqId & " " & _
                         " ,A." & gsCAT_PSES_UPH & " " & _
                 "     FROM " & gsCAT_TBL_PROD_STEP_EQ_SPEC & " A " & _
                 "    WHERE A." & gsCAT_PSES_PROD_GROUP & " = '" & sProdgroup & "'" & _
                     "  and A." & gsCAT_PSES_STEP_NO & " = '" & sStepID & "'" & _
                     "  and A." & gsCAT_PSES_EQTYPE2 & " = '" & sEqType2 & "'" & _
                     "  and A." & gsCAT_PSES_PG_NAME & " = '" & sPgName & "'" & _
                     "  and A." & gsCAT_PSES_DELETE_FLAG & " = 'N'" & _
                     "  and A." & gsCAT_PSES_DOC_STATUS & " = 'Active' "
            
            'Add by Weilun on 20191126 for Project. 電測自動化監控系統 <start>
            If Trim(sSubSystem) <> "" Then
                sSQL = sSQL & _
                     "  and A." & gsCAT_PSES_SUB_SYSTEM & " = '" & sSubSystem & "' "
            End If
            'Add by Weilun on 20191126 for Project. 電測自動化監控系統 <end>
    
            Set colRaws = moProRawSql.QueryDatabase(sSQL)
            If colRaws Is Nothing Then
                Call RaiseError(glERR_INVALIDOBJECT, _
                            FormatErrorText(gsETX_INVALIDOBJECT, "Collection"))
            End If
            
            If colRaws.Count > 0 Then
                dUph = Val(colRaws.Item(1).Item(gsCAT_PSES_UPH))
            Else
                dUph = 0
                'Add by Sam Start on 20230821 for ReqNo:202300189
                sSQL = "select A." & gsCAT_TPU_UPH & " from " & gsCAT_TBL_PGM_UPH & " a " & _
                       " where a." & gsCAT_TPU_TDSPROD & " = '" & sProdcode & "' " & _
                       " and  a." & gsCAT_TPU_PKGCODE & " = '" & sPkgCode & "' " & _
                       " and  a." & gsCAT_TPU_PINCOUNT & " = '" & sPincount & "' " & _
                       " and  a." & gsCAT_TPU_TDSTESTERTYPE & " = '" & sEqType2 & "' " & _
                       " and  a." & gsCAT_TPU_PGNAME & " = '" & sPgName & "' " & _
                       " and  a." & gsCAT_TPU_TESTMODE & " = '" & Replace(sStepName, "SORT", "S") & "' "
                If Trim(sSubSystem) <> "" Then
                    sSQL = sSQL & _
                         "  and A." & gsCAT_TPU_TDSSUBSYSTYPE & " = '" & sSubSystem & "' "
                End If
                Set colRaws = moProRawSql.QueryDatabase(sSQL)
                If colRaws Is Nothing Then
                    Call RaiseError(glERR_INVALIDOBJECT, _
                                FormatErrorText(gsETX_INVALIDOBJECT, "Collection"))
                End If
                            
                If colRaws.Count > 0 Then
                     dUph = Val(colRaws.Item(1).Item(gsCAT_TPU_UPH))
                End If
                'Add by Sam END on 20230821 for ReqNo:202300189
            End If
        End If
    End If
        
    GetUph = dUph
    '-------
    'Done
    '-------
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
           Case Else
                typErrInfo.sUserText = typErrInfo.sErrDescription
        End Select
        '---- Start of Select Case Block ----
        On Error GoTo ExitHandler:
        Call TransferLogError(typErrInfo.sUserText, _
                              moAppLog, Me.rtfTrace, msMODULE_ID, sProcID)
        Call HandleError(False, typErrInfo, , moAppLog)
    End If
End Function

'================================================================================
' Sub: GetErunRecipe()
'--------------------------------------------------------------------------------
' Description:  <Type your Sub description here...>
'--------------------------------------------------------------------------------
' Author:       Sam Chen, CIT 2011/10/20
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
''Modify by Sam on 20140708 for Project 機台vs配件停復測系統建構 ,增加回傳pgid pgmode
' Copy from WS.frmWSDailySkd by Weilun on 20190906
'================================================================================
Private Sub GetErunRecipe(ByVal sErunTicNO As String, ByVal sStepID As String, ByRef sEqType2 As String _
                              , ByRef sPgName As String, ByRef sPgId As String, ByRef sPgMode As String _
                              , ByRef sTemp As String, ByRef sProbecard As String, ByRef sWsDeviceFile As String _
                              , ByRef dProcTime As Double, Optional ByVal sSubSystem As String)
On Error GoTo ExitHandler:
Dim sProcID As String
Dim typErrInfo As tErrInfo
Dim colRaws As Collection

Dim sSQL As String
Dim colRS As Collection
Dim iIndex As Integer

Dim sProcUnit As String

'----
' Init
'----
    sProcID = "GetProdSpec"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog)

'----
' Condition Checking
'----
    ' <Put your condition checking codes here>...
    
'----
' Action
'----
'Modify by Sam on 20140708 for Project 機台vs配件停復測系統建構 ,增加回傳pgid pgmode
    sSQL = "SELECT " & gsCAT_TERE_EQTYPE2 & " ," & gsCAT_TERE_PGNAME & " ," & gsCAT_TERE_PGID & " ," & gsCAT_TERE_PGMODE & _
                      "," & gsCAT_TERE_TEMPERATURE & "," & gsCAT_TERE_WSDEVICEFILE & _
                      "," & gsCAT_TERE_PROBECARDTYPE & " " & _
                      "," & gsCAT_TERE_PROCESSTIME & "," & gsCAT_TERE_PROCTIMEUNIT & _
                      " from " & gsCAT_TBL_ERUN_RECIPE & _
                      " where " & gsCAT_TERE_STEPNO & " ='" & sStepID & "'" & _
                      " and " & gsCAT_TERE_DOCNO & "='" & sErunTicNO & "' " & _
                      " and " & gsCAT_TERE_DELETE_FLAG & " = 'N' "
                      
    If Trim(sEqType2) <> "" Then
        sSQL = sSQL & " and " & gsCAT_TERE_EQTYPE2 & " = '" & sEqType2 & "' "
    End If
    
    If Trim(sSubSystem) <> "" Then
        sSQL = sSQL & " and subsystem='" & Trim(sSubSystem) & "' "
    End If
                            
    sSQL = sSQL & " order by " & gsCAT_TERE_PROCESSTIME & " "
    
    Set colRaws = moProRawSql.QueryDatabase(sSQL)
    If colRaws Is Nothing Then
        Call RaiseError(glERR_INVALIDOBJECT, _
                    FormatErrorText(gsETX_INVALIDOBJECT, "Collection"))
    End If
    
    If colRaws.Count = 0 Then '無資料時去掉subsystem條件
'Modify by Sam on 20140708 for Project 機台vs配件停復測系統建構 ,增加回傳pgid pgmode
        sSQL = "SELECT " & gsCAT_TERE_EQTYPE2 & " ," & gsCAT_TERE_PGNAME & " ," & gsCAT_TERE_PGID & " ," & gsCAT_TERE_PGMODE & _
                          "," & gsCAT_TERE_TEMPERATURE & "," & gsCAT_TERE_WSDEVICEFILE & _
                          "," & gsCAT_TERE_PROBECARDTYPE & " " & _
                          "," & gsCAT_TERE_PROCESSTIME & "," & gsCAT_TERE_PROCTIMEUNIT & _
                          " from " & gsCAT_TBL_ERUN_RECIPE & _
                          " where " & gsCAT_TERE_STEPNO & " ='" & sStepID & "'" & _
                          " and " & gsCAT_TERE_DOCNO & "='" & sErunTicNO & "' " & _
                          " and " & gsCAT_TERE_DELETE_FLAG & " = 'N' "
                          
        If Trim(sEqType2) <> "" Then
            sSQL = sSQL & " and " & gsCAT_TERE_EQTYPE2 & " = '" & sEqType2 & "' "
        End If
                                
        sSQL = sSQL & " order by " & gsCAT_TERE_PROCESSTIME & " "
        
        Set colRaws = moProRawSql.QueryDatabase(sSQL)
        If colRaws Is Nothing Then
            Call RaiseError(glERR_INVALIDOBJECT, _
                    FormatErrorText(gsETX_INVALIDOBJECT, "Collection"))
        End If

    End If
    If colRaws.Count > 0 Then
        sEqType2 = colRaws.Item(1).Item(gsCAT_TERE_EQTYPE2)
        sPgName = colRaws.Item(1).Item(gsCAT_TERE_PGNAME)
        sTemp = colRaws.Item(1).Item(gsCAT_TERE_TEMPERATURE)
        sProbecard = colRaws.Item(1).Item(gsCAT_TERE_PROBECARDTYPE)
        sWsDeviceFile = colRaws.Item(1).Item(gsCAT_TERE_WSDEVICEFILE)
        dProcTime = Val(colRaws.Item(1).Item(gsCAT_TERE_PROCESSTIME))
        sProcUnit = colRaws.Item(1).Item(gsCAT_TERE_PROCTIMEUNIT)
        
        '轉成HR
        Select Case sProcUnit
            Case "MINS"
                dProcTime = dProcTime / 60
            Case "SECS"
                dProcTime = dProcTime / 60 / 60
        End Select
        
'Add by Sam start on 20140708 for Project 機台vs配件停復測系統建構 ,增加回傳pgid pgmode
        sPgId = colRaws.Item(1).Item(gsCAT_TERE_PGID)
        sPgMode = colRaws.Item(1).Item(gsCAT_TERE_PGMODE)
'Add by Sam end on 20140708 for Project 機台vs配件停復測系統建構 ,增加回傳pgid pgmode
        
    Else
        sPgName = ""
        sTemp = ""
        sProbecard = ""
        sWsDeviceFile = ""
        dProcTime = 0
'Add by Sam start on 20140708 for Project 機台vs配件停復測系統建構 ,增加回傳pgid pgmode
        sPgId = ""
        sPgMode = ""
'Add by Sam end on 20140708 for Project 機台vs配件停復測系統建構 ,增加回傳pgid pgmode
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

'================================================================================
' Sub: GetTestedPcs()
'--------------------------------------------------------------------------------
' Description:  <Type your Sub description here...>
'--------------------------------------------------------------------------------
' Author:       Sam Chen, CIT 2011/10/20
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
' Copy from WS.frmWSDailySkd by Weilun on 20190906
'================================================================================
Private Function GetTestedPcs(ByVal sStepID As String, ByVal sLotId As String, _
                              ByVal sEqId As String, ByVal sTimeHereSince As String) As Integer
On Error GoTo ExitHandler:
Dim sProcID As String
Dim typErrInfo As tErrInfo
Dim colRaws As Collection
Dim oRaws As FwStrings
Dim sTable As String
Dim sColumn As String
Dim sWhere As String

Dim sSQL As String
Dim colRS As Collection
Dim iIndex As Integer

Dim vGroupID As Variant
Dim vPrevGroupID As Variant

Dim sProducts As String
Dim vSets As Variant
Dim vIPN As Variant
Dim vTesterType As Variant

Dim iTestedPcs As Integer

'----
' Init
'----
    sProcID = "GetTestedPcs"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog)

'----
' Condition Checking
'----
    ' <Put your condition checking codes here>...
    
'----
' Action
'----

  sSQL = "   SELECT count(*) " & _
            "     FROM " & gsCAT_TBL_WS_TDS_SUM & " " & _
            "    WHERE (" & gsCAT_TWTS_WAFER_IDN & _
            "           ,nvl(" & gsCAT_TWTS_TIME_STAMP & ", " & gsCAT_TWTS_INSTEPTIME & ")) IN" & _
            "          (SELECT " & gsCAT_TWTS_WAFER_IDN & "," & _
            "                 MAX(nvl(" & gsCAT_TWTS_TIME_STAMP & ", " & gsCAT_TWTS_INSTEPTIME & ")) " & _
            "             FROM " & gsCAT_TBL_WS_TDS_SUM & "" & _
            "            WHERE " & gsCAT_TWTS_STEPNO & " = '" & sStepID & "'" & _
            "              and " & gsCAT_TWTS_LOT_ID & " = '" & sLotId & "'" & _
            "              and " & gsCAT_TWTS_INSTEPTIME & " = '" & sTimeHereSince & "'" & _
            "              and " & gsCAT_TWTS_DELETE_FLAG & " = 'N'" & _
            "            GROUP BY  " & gsCAT_TWTS_WAFER_IDN & " ) " & _
            "  and " & gsCAT_TWTS_STEPNO & " = '" & sStepID & "'" & _
            "  and " & gsCAT_TWTS_LOT_ID & " = '" & sLotId & "'" & _
            "  and " & gsCAT_TWTS_TESTED_PCS & "  is not null " & _
            "  and " & gsCAT_TWTS_DELETE_FLAG & " = 'N'"

'    Set colRaws = modRawSQL.SelectRawSqlTxn1(moAppLog, moFwWIP, moFwWF, moCwMbx, "CAT", sSQL)
    Set colRaws = moProRawSql.QueryDatabase(sSQL)
    If colRaws Is Nothing Then
        Call RaiseError(glERR_INVALIDOBJECT, _
                    FormatErrorText(gsETX_INVALIDOBJECT, "Collection"))
    End If
    
    If colRaws.Count > 0 Then
        iTestedPcs = Val(colRaws.Item(1).Item(1))
    End If
    
    GetTestedPcs = iTestedPcs
    
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

'================================================================================
'Name       : GetHotFlag()
'Description: 取得HotFlag
'Author     : Weilun huang, MXIC 2019/09/20
'Parameters : N/A
'================================================================================
Private Function GetHotFlag(ByVal sLotId As String) As String
    On Error GoTo ExitHandler
    Dim sProcID As String
    Dim typErrInfo As tErrInfo

    Dim sSQL As String
    Dim colSQLResult As Collection

    Dim oLot As FwLot
    Dim sIPN As String
    Dim sWaferLevel As String
    Dim sProdcode As String
    Dim sMainProd As String

    '-----
    ' Init
    '-----
    sProcID = "GetHotFlag"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog) '

    GetHotFlag = ""

    '--------------------
    ' Condition Checking
    '--------------------

    '-------
    ' Action
    '-------

    Set oLot = moFwWIP.LotById(sLotId)

    If Not oLot Is Nothing Then
        'Copy from Query.frmLotAttQuery.ExecQuery
        sSQL = "select " & gsCAT_TSL_LOCATION & " from " & gsCAT_TBL_STEP_LOCATION & " where " & gsCAT_TSL_STEPID & " ='" & oLot.CurrentStep.Steps.Item(1).Id & "'"
        Set colSQLResult = moProRawSql.QueryDatabase(sSQL)
        If colSQLResult.Count > 0 Then
            If colSQLResult.Item(1).Item(gsCAT_TSL_LOCATION) = "WS" Then

                sIPN = oLot.CustomAttributes(gsLOT_CUSTOMATTR_IPN)
                sWaferLevel = ""
                sProdcode = ""
                sMainProd = ""
                sSQL = "select fun_get_waferlevel(a." & gsCAT_TIM_IPN & ") as waferlevel " & _
                       " ,a." & gsCAT_TIM_MASK_OPTION & " , a." & gsCAT_TIM_BE_OPTION & " " & _
                       " from " & gsCAT_TBL_IPN_MASTER & " a " & _
                       " where a." & gsCAT_TIM_IPN & "= '" & sIPN & "' " & _
                       " and a." & gsCAT_TIM_DELETE_FLAG & "= 'N' "
                Set colSQLResult = moProRawSql.QueryDatabase(sSQL)
                If colSQLResult.Count > 0 Then
                    sWaferLevel = colSQLResult.Item(1).Item("waferlevel")
                    sProdcode = Left(sIPN, 4) & colSQLResult.Item(1).Item(gsCAT_TIM_MASK_OPTION)
                End If
                sSQL = "select a." & gsCAT_TPMP_PRODCODE & "   " & _
                      " from " & gsCAT_TBL_PC_MAIN_PROD & " a " & _
                      " where '" & sProdcode & "' like replace(replace(a." & gsCAT_TPMP_PRODCODE & ",'%','_'),'*','%')   " & _
                      " and '" & sIPN & "' like replace(replace(a." & gsCAT_TPMP_IPN & ",'%','_'),'*','%')  " & _
                      " and  '" & sWaferLevel & "' like  replace(a." & gsCAT_TPMP_WAFERLEVEL & ",'*','%')  " & _
                      " and a." & gsCAT_TPMP_DELETEFLAG & "= 'N' "

                Set colSQLResult = moProRawSql.QueryDatabase(sSQL)
                If colSQLResult.Count > 0 Then
                    sMainProd = "Y"
                End If

                sSQL = "select " & gsCAT_TWTL_WSOUT & " from " & gsCAT_TBL_WS_TRACKLOT & " where " & _
                        gsCAT_TWTL_LOTID & " ='" & sLotId & "' and " & _
                        gsCAT_TWTL_DELETEFLAG & " ='N' and " & gsCAT_TWTL_DOCSTATUS & "='Active'"
                Set colSQLResult = moProRawSql.QueryDatabase(sSQL)
                If colSQLResult.Count > 0 Then
                    If colSQLResult.Item(1).Item(gsCAT_TWTL_WSOUT) <> "" Then
                        sMainProd = "Y"
                    End If
                End If
                GetHotFlag = sMainProd
            End If
        End If
    End If

    '-------
    'Done
    '-------
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
            Case Else
                typErrInfo.sUserText = typErrInfo.sErrDescription
        End Select
        '---- Start of Select Case Block ----
        On Error GoTo ExitHandler
        Call HandleError(False, typErrInfo, , moAppLog)
    End If
End Function

'================================================================================
'Name       : GetNextStepNameAndRecipe()
'Description: 取得Lot的下一站以及為Bake時取出Recipe
'Author     : Weilun huang, MXIC 2019/09/18
'Parameters : N/A
'
'Add bAutoDGrade by Weilun on 20200630 for Project.BE MES Phase 44 – CP 產線增加自動化作業 UAT fix
'================================================================================
Private Function GetNextStepNameAndRecipe(ByVal sLotId As String, _
                                          ByRef sNextStepName As String, _
                                          ByRef sRecipe As String, _
                                          Optional ByVal bAutoDGrade As Boolean = False) As Boolean
    On Error GoTo ExitHandler:
    Dim sProcID                     As String
    Dim typErrInfo                  As tErrInfo
    
    Dim sSQL                        As String
    Dim colRaws                     As Collection
    Dim oRaws                       As Collection
    
    Dim oLot                        As FwLot
    Dim oAggNextStep                As FwAggregateStep
    Dim oLastStep                   As FwStep
    Dim sEqType2                    As String
    Dim sCurHandlerType2            As String
    Dim sProdgroup                  As String
    Dim sPath                       As String
    Dim sNextStepID                 As String
    
    Dim sTecn                       As String
    Dim sLotFutAct                  As String
    Dim sIPNFutAct                  As String
    Dim sProdGroupFutAct            As String
    Dim sPgName                     As String
    Dim sTemp                       As String
    Dim sProcTime                   As String
    Dim sProcUnit                   As String
    Dim sErunTicNO                  As String
    Dim sFollowOI                   As String
    Dim bTECNLotSpec                As String
    Dim colLotSpec                  As Collection
    
    '-----
    ' Init
    '-----
    sProcID = "GetNextStepNameAndRecipe"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog) '
    
    GetNextStepNameAndRecipe = False
    sNextStepName = ""
    sNextStepID = ""
    sRecipe = ""
    
    '--------------------
    ' Condition Checking
    '--------------------
    
    '-------
    ' Action
    '-------
    '測試用LotID: 8E449300 'ERun
    '             5D579500 'Normal
    
'    sLotID = "5D579500"
    Set oLot = moFwWIP.LotById(sLotId)
    
    If Not oLot Is Nothing Then

        'AutoDGrade要取得調完站的, 改用Current,Add by Weilun on 20200630 for Project.BE MES Phase 44 – CP 產線增加自動化作業 UAT fix
        If bAutoDGrade = True Then
            Set oAggNextStep = oLot.CurrentStep
        Else
            '無法取得下一站時跳開
            'PS: 另一種作法是If sErunTicNO <> "" And sFollowOI = "N" And sFollowProduct = "N" Then
            If oLot.CurrentStep.Steps.Item(1).AvailablePaths.Count = 0 Then
                GoTo ExitHandler
            End If
            
            '取得下一站
            Set oAggNextStep = FwuRetrieveNextStep(moFwPRP, _
                                                   oLot.PlanId, oLot.PlanVersion, _
                                                   oLot.CurrentStep.Steps.Item(1).Handle, "PASS", moAppLog)
        End If
                                               
        If oAggNextStep Is Nothing Then
            GoTo ExitHandler
        End If
                                               
        sNextStepName = oAggNextStep.Steps.Item(1).Description
        sNextStepID = oAggNextStep.Steps.Item(1).Id
        
        '僅BAKE站要倒出Recipe
        If sNextStepName Like "BAKE*" = True Then
            sEqType2 = "OVEN"
            sCurHandlerType2 = "" 'BAKE機台沒有副機
            
            '此SQL是以WsRun跑modCustom.RetrieveLotInfo所改寫的
            sSQL = " select distinct X.*, " & _
                                   " Y." & gsCAT_PSES_PG_NAME & ", " & _
                                   " Y." & gsCAT_PSES_TEMPERATURE & ", " & _
                                   " Y." & gsCAT_PSES_PROCESS_TIME & ", " & _
                                   " Y." & gsCAT_PSES_PROC_TIME_UNIT & ", " & _
                                   " decode(M." & gsCAT_TPS_PROD_GROUP & ", NULL, 'N', 'Y') TECN, " & _
                                   " Fun_ChkFutActByIPN(X." & gsCAT_TLATT_IPN & ", " & _
                                                      " X.StepNo, " & _
                                                      " X." & gsCAT_TLATT_ROUTE & ", " & _
                                                      " X.LOTTYPE, " & _
                                                      " X." & gsCAT_TLATT_LOTOWNER & ") as IPNFUTACT, " & _
                                   " Fun_ChkFutActByProdGroup(X." & gsCAT_TLATT_PRODGROUP & ", " & _
                                                            " X.StepNo, " & _
                                                            " X." & gsCAT_TLATT_ROUTE & ", " & _
                                                            " X.LotType, " & _
                                                            " X." & gsCAT_TLATT_LOTOWNER & ") As PRODGROUPFUTACT, " & _
                                   " Z." & gsCAT_TER_FOLLOW_OI & " "
                                   
            sSQL = sSQL & " from (select XX.*, trsl." & gsCAT_TRSL_PATH & " " & _
                                  " from (select tla." & gsCAT_TLATT_LOTID & ", " & _
                                               " tla." & gsCAT_TLATT_STAGE & ", " & _
                                               " tla." & gsCAT_TLATT_IPN & ", " & _
                                               " tla." & gsCAT_TLATT_LOTOWNER & ", " & _
                                               " tla." & gsCAT_TLATT_ROUTE & ", " & _
                                               " '" & sNextStepID & "' as StepNo, " & _
                                               " tla." & gsCAT_TLATT_PRODGROUP & ", " & _
                                               " e." & gsCAT_TLI_ERUNTICNO & ", " & _
                                               " b." & "planname" & ", " & _
                                               " d." & "handle" & ", " & _
                                               " decode(N." & gsCAT_TLFA_LOT_ID & ", NULL, 'N', 'Y') LotFutAct, " & _
                                               " DECODE(E." & gsCAT_TLI_ERUNTICNO & ", " & _
                                                      " NULL, " & _
                                                      " DECODE(E." & gsCAT_TLI_SAPRWNO & ", NULL, 'Normal', 'Rework'), " & _
                                                      " 'Eng') as LOTTYPE "
    
            sSQL = sSQL & " from (SELECT FROMID, TOID " & _
                                  " From FWLOT_N2M " & _
                                 " WHERE LINKNAME = 'stepCollection') c, " & _
                               " (select " & gsCAT_TLFA_LOT_ID & ", " & _
                                             gsCAT_TLFA_STEP_NO & ", " & _
                                             gsCAT_TLFA_PATH & " " & _
                                  " From " & gsCAT_TBL_LOT_FUTACT & " " & _
                                 " where " & gsCAT_TLFA_DELETE_FLAG & " = 'N' " & _
                              " group by " & gsCAT_TLFA_LOT_ID & ", " & _
                                             gsCAT_TLFA_STEP_NO & ", " & _
                                             gsCAT_TLFA_PATH & ") N, " & _
                                       " fwlot b, " & _
                                       " fwwipstep d, " & _
                                       gsCAT_TBL_LOT_INFO & " e, " & _
                                       gsCAT_TBL_LOT_ATTRIBUTE & " tla " & _
                         " Where tla." & gsCAT_TLATT_LOTID & " = b.appid " & _
                           " and b.sysid = c.fromid(+) " & _
                           " and c.toid = d.sysid " & _
                           " and tla." & gsCAT_TLATT_LOTID & " = e." & gsCAT_TLI_LOT_ID & "(+) " & _
                           " and tla." & gsCAT_TLATT_LOTID & " = N." & gsCAT_TLFA_LOT_ID & "(+) " & _
                           " and tla." & gsCAT_TLATT_STEPID & " = N." & gsCAT_TLFA_STEP_NO & "(+) " & _
                           " and tla." & gsCAT_TLATT_ROUTE & " = N." & gsCAT_TLFA_PATH & "(+) " & _
                           " and tla." & gsCAT_TLATT_LOTID & " = '" & sLotId & "') XX, " & _
                         gsCAT_TBL_ROUTE_STEP_LIST & " trsl "
    
            sSQL = sSQL & " where XX." & "PLANNAME" & " = trsl." & gsCAT_TRSL_ROUTE_ID & "(+) " & _
                            " and XX.HANDLE = trsl." & gsCAT_TRSL_FWSTEPSEQ & "(+)) X, " & _
                        " (select " & gsCAT_PSES_PROD_GROUP & ", " & _
                                      gsCAT_PSES_STEP_NO & ", " & _
                                      gsCAT_PSES_PATH & ", " & _
                                      gsCAT_PSES_PG_NAME & ", " & _
                                      gsCAT_PSES_TEMPERATURE & ", " & _
                                      gsCAT_PSES_PROCESS_TIME & ", " & _
                                      gsCAT_PSES_PROC_TIME_UNIT & " " & _
                           " From " & gsCAT_TBL_PROD_STEP_EQ_SPEC & " " & _
                          " where " & gsCAT_PSES_EQTYPE2 & " = '" & sEqType2 & "' " & _
                            " and " & gsCAT_PSES_DOC_STATUS & " = 'Active' " & _
                            " and " & gsCAT_PSES_DELETE_FLAG & " = 'N') Y, " & _
                        " (select * from " & gsCAT_TBL_ERUN_REQ & " where " & gsCAT_TER_DELETE_FLAG & " = 'N') Z, " & _
                        " (select " & gsCAT_TPS_PROD_GROUP & " " & _
                           " From " & gsCAT_TBL_PROD_SPEC & " " & _
                          " where " & gsCAT_TPS_DOC_STATUS & " = 'Active' " & _
                            " and " & gsCAT_TPS_DOC_NO & " like '-1' " & _
                            " and " & gsCAT_TPS_DELETE_FLAG & " = 'N') M "
    
            sSQL = sSQL & " where X." & "PRODGROUP" & " = Y." & gsCAT_PSES_PROD_GROUP & "(+) " & _
                            " and X." & "stepno" & " = Y." & gsCAT_PSES_STEP_NO & "(+) " & _
                            " and X." & "path" & " = Y." & gsCAT_PSES_PATH & "(+) " & _
                            " and X." & "PRODGROUP" & " = M." & gsCAT_TPS_PROD_GROUP & "(+) " & _
                            " and X." & "lotid" & " = Z." & gsCAT_TER_LOT_ID & "(+) " & _
                            " and X." & "erunticno" & " = Z." & gsCAT_TER_TICKET_NO & "(+) " & _
                            " and X." & "Stage" & " = Z." & gsCAT_TER_STAGE & "(+) "
    
            Set colRaws = moProRawSql.QueryDatabase(sSQL)
            If colRaws.Count > 0 Then

                sLotFutAct = colRaws.Item(1).Item("LotFutAct")
                sIPNFutAct = colRaws.Item(1).Item("IPNFutAct")
                sProdGroupFutAct = colRaws.Item(1).Item("ProdGroupFutAct")
                sErunTicNO = colRaws.Item(1).Item("ERunTicNo")
                sFollowOI = colRaws.Item(1).Item("FollowOI")

                sProdgroup = oLot.CustomAttributes(gsLOT_CUSTOMATTR_PROD_GROUP)
                sPath = oLot.CustomAttributes(gsLOT_CUSTOMATTR_ROUTE)
                
                'TECN相關
                bTECNLotSpec = CheckTECNLotInfo(moAppLog, moFwWIP, moFwWF, moCwMbx, sLotId)
                If bTECNLotSpec = True Then
                    Set colLotSpec = GetLotSpecRecipe(moFwWIP, moFwWF, moCwMbx, moAppLog, sLotId, sProdgroup, _
                                                     sPath, sNextStepID, sEqType2, sCurHandlerType2)
                End If
                
                If bTECNLotSpec = True Then
                    sTecn = "Y"
                Else
                    sTecn = colRaws.Item(1).Item("TECN")
                End If

                '----
                '- Get temperature,pgname, proc_time, proc_unit (Recipe Spec)- 先找 TECN , 若不合適再找 ECN
                '----
                If sTecn = "Y" Then
                    If bTECNLotSpec = True Then
                        If colLotSpec.Count > 0 Then
                            sPgName = colLotSpec.Item(1).Item(gsCAT_TLSPES_PGNAME)
                            sTemp = colLotSpec.Item(1).Item(gsCAT_TLSPES_TEMPERATURE)
                            sProcTime = colLotSpec.Item(1).Item(gsCAT_TLSPES_PROCESSTIME)
                            sProcUnit = colLotSpec.Item(1).Item(gsCAT_TLSPES_PROCTIMEUNIT)
                           
                        Else
                            sPgName = ""
                            sTemp = ""
                            sProcTime = ""
                            sProcUnit = ""
                           
                        End If
                    Else
                        Call GetEqSpecInfoByTECN(moAppLog, moFwWIP, moFwWF, moCwMbx, _
                                                 sProdgroup, sNextStepID, sEqType2, _
                                                 "", sPgName, sTemp, sProcTime, sProcUnit, "")
                    End If
                Else
                    '----
                    ' 搜尋Future Action , 若有的話, 用future action的PgId and PgName,
                    ' 若沒有的話,用Tbl_Prod_Step_Eq_Spec ECN 之PgId and Name
                    '----
                    If sLotFutAct = "Y" Or sIPNFutAct = "Y" Or sProdGroupFutAct = "Y" Then
                        Call GetPgIdAndNameFromFutAct(moAppLog, moFwWIP, moFwWF, moCwMbx, sLotId, sNextStepID, sPath, "", sPgName, "", "", "")
                        If Len(Trim$(sPgName)) = 0 Then
                            sPgName = colRaws.Item(1).Item(gsCAT_PSES_PG_NAME)
                        End If
                    Else
                        sPgName = colRaws.Item(1).Item(gsCAT_PSES_PG_NAME)
                    End If
                    '----
                    ' 用Tbl_Prod_Step_Eq_Spec ECN 之Temperature,Processtime,ProcTimeUnit,DocType
                    '----
                    sTemp = colRaws.Item(1).Item(gsCAT_PSES_TEMPERATURE)
                    sProcTime = colRaws.Item(1).Item(gsCAT_PSES_PROCESS_TIME)
                    sProcUnit = colRaws.Item(1).Item(gsCAT_PSES_PROC_TIME_UNIT)
                    '----
                End If
                
                '組合Recipe
                If sErunTicNO <> "" And sFollowOI <> "Y" Then
                    sRecipe = "委測:" & sErunTicNO
                Else
                    'Get Recipe
                    sRecipe = sPgName & Chr(32) & sTemp & Chr(32) & sProcTime & Chr(32) & sProcUnit
                End If
                            
                GetNextStepNameAndRecipe = True
            End If
        End If
    End If
    
    '-------
    'Done
    '-------
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
           Case Else
                typErrInfo.sUserText = typErrInfo.sErrDescription
        End Select
        '---- Start of Select Case Block ----
        On Error GoTo ExitHandler:
        Call HandleError(False, typErrInfo, , moAppLog)
    End If
End Function

'================================================================================
'Name       : GetCurWaferIDByPgmList()
'Description: 取得目前正在測的WaferID, 僅為預測
'Author     : Weilun huang, MXIC 2019/09/25
'Parameters : N/A
'================================================================================
Private Function GetCurWaferIDByPgmList(ByVal sLotId As String)
    On Error GoTo ExitHandler:
    Dim sProcID                     As String
    Dim typErrInfo                  As tErrInfo

    Dim sSQL                        As String
    Dim colRaws                     As Collection
    Dim oRaws                       As Collection

    Dim sTimeHereSince              As String
    Dim sWaferIDList                As String
    Dim sStepName                   As String
    Dim sLastWaferID                As String
    Dim sParentLotID                As String

    Dim oLot                        As FwLot
    Dim oEquipment                  As FwEquipment
    
    Dim iCur                        As Integer
    Dim iLast                       As Integer
    Dim asWaferIDList()             As String

    '-----
    ' Init
    '-----
    sProcID = "GetCurWaferIDByPgmList"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog) '

    GetCurWaferIDByPgmList = "NA" '預設值

    '--------------------
    ' Condition Checking
    '--------------------

    '-------
    ' Action
    '-------

    Set oLot = moFwWIP.LotById(sLotId)
    If Not oLot Is Nothing Then
        sTimeHereSince = ConvertToString(oLot.CurrentStep.Steps.Item(1).TimeHereSince, moAppLog)
        sStepName = oLot.CurrentStep.Steps.Item(1).Description
        sParentLotID = Left(sLotId, 8)

        '取得pgm_list的UserWaferID作為列表
        sSQL = " select " & gsCAT_TPL_LOT_ID & ", " & _
                            gsCAT_TPL_USERWAFERID & " " & _
                 " from " & gsCAT_TBL_PGM_LIST & " " & _
                " where (" & gsCAT_TPL_LOT_ID & ", " & _
                             gsCAT_TPL_CREATE_TIME & ") in " & _
                      " ( select " & gsCAT_TPL_LOT_ID & ", " & _
                           " max(" & gsCAT_TPL_CREATE_TIME & ") " & _
                          " from " & gsCAT_TBL_PGM_LIST & " " & _
                         " where " & gsCAT_TPL_LOT_ID & " ='" & sLotId & "' " & _
                           " and " & gsCAT_TPL_PROD_TYPE & " = 'WAFER' " & _
                           " and " & gsCAT_TPL_SOURCE & " <> 'Setup' " & _
                           " and " & gsCAT_TPL_CREATE_TIME & " >= '" & sTimeHereSince & "' " & _
                           " and " & gsCAT_TPL_DELETE_FLAG & " ='N' " & _
                      " group by " & gsCAT_TPL_LOT_ID & ") "
        Set colRaws = moProRawSql.QueryDatabase(sSQL)
        If colRaws.Count > 0 Then
            sWaferIDList = colRaws.Item(1).Item(gsCAT_TPL_USERWAFERID)
        Else
            GoTo ExitHandler
        End If

        If sWaferIDList = "" Then
            GoTo ExitHandler
        End If
        '以上有問題的維持NA
        
        sWaferIDList = Replace(sWaferIDList, ",", ";")
        asWaferIDList = Split(sWaferIDList, ";")
        
        '試著抓出該Lot最後一筆WaferEnd
        sSQL = " select distinct " & gsCAT_TCWE_WAFERID & " " & _
                 " from " & gsCAT_TBL_CP_WAFER_END & " " & _
                " where (" & gsCAT_TCWE_LOTID & ", " & _
                             gsCAT_TCWE_STEPNAME & ", " & _
                             gsCAT_TCWE_WAFERENDTIME & ") in "
        sSQL = sSQL & _
              " (select TCWE." & gsCAT_TCWE_LOTID & ", " & _
                      " TCWE." & gsCAT_TCWE_STEPNAME & ", " & _
                  " max(TCWE." & gsCAT_TCWE_WAFERENDTIME & ") " & _
                 " from " & gsCAT_TBL_CP_WAFER_END & " TCWE, " & _
                            gsCAT_TBL_LOT_INFO & " TLI " & _
                " where instr(TLI." & gsCAT_TLI_WAFERID & ", TCWE." & gsCAT_TCWE_WAFERID & ") > 0 " & _
                  " and substr(TCWE." & gsCAT_TCWE_LOTID & ",1,8) = substr(TLI." & gsCAT_TLI_LOT_ID & ",1,8) " & _
                  " and TCWE." & gsCAT_TCWE_WAFERENDTIME & " > '" & sTimeHereSince & "' " & _
                  " and TCWE." & gsCAT_TCWE_STEPNAME & " = '" & sStepName & "' " & _
                  " and TCWE." & gsCAT_TCWE_WAFERID & " is not null " & _
                  " and TCWE." & gsCAT_TCWE_DELETEFLAG & " = 'N' " & _
                  " and TCWE." & gsCAT_TCWE_LOTID & " like '" & sParentLotID & "%' " & _
                  " and TLI." & gsCAT_TLI_WAFERID & " is not null " & _
                  " and TLI." & gsCAT_TLI_LOT_ID & " = '" & sLotId & "' " & _
             " group by TCWE." & gsCAT_TCWE_LOTID & ", " & _
                      " TCWE." & gsCAT_TCWE_STEPNAME & ") "
                      
        Set colRaws = moProRawSql.QueryDatabase(sSQL)
        If colRaws.Count > 0 Then
            sLastWaferID = colRaws.Item(1).Item(gsCAT_TCWE_WAFERID)
            
            iLast = GetStrPosInArray(sWaferIDList, sLastWaferID, ";") - 1 '回傳1~n, 再調整為0~n-1, -1為找不到
            iCur = iLast + 1
        Else
            'Get First by list
            iCur = 0
        End If
        
        If iCur > -1 And iCur <= UBound(asWaferIDList) Then
            GetCurWaferIDByPgmList = asWaferIDList(iCur)
        End If '取不到也是維持NA

    End If
    '-------
    'Done
    '-------
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
           Case Else
                typErrInfo.sUserText = typErrInfo.sErrDescription
        End Select
        '---- Start of Select Case Block ----
        On Error GoTo ExitHandler:
        Call TransferLogError(typErrInfo.sUserText, _
                              moAppLog, Me.rtfTrace, msMODULE_ID, sProcID)
        Call HandleError(False, typErrInfo, , moAppLog)
    End If
End Function

'================================================================================
'Name       : CheckTdsTestSumComplete()
'Description: 檢查Tds Summary是否到齊, 相對於CPautoLotEnd用的是一個精簡版
'Author     : Weilun huang, MXIC 2019/10/01
'Parameters : N/A
'================================================================================
Private Function CheckTdsTestSumComplete(oLot As FwLot, ByRef sMaxUpdateTime As String) As Boolean
On Error GoTo ExitHandler:
Dim sProcID                 As String
Dim typErrInfo              As tErrInfo

Dim sSQL                    As String
Dim colRaws                 As Collection
Dim oRaws                   As Collection
Dim sTable                  As String
Dim sColumn                 As String
Dim sWhere                  As String
Dim sOther                  As String

Dim sWaferIDList            As String
Dim sParentLot              As String
Dim sTimeHereSince          As String
Dim sStepNo                 As String
Dim sUpdateTime             As String
Dim sTdsQty                 As String
Dim sReviseQty              As String
Dim sReasonCode             As String
Dim sWaferID                As String
'----
' Init
'----
    sProcID = "CheckTdsTestSumComplete"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog)

    CheckTdsTestSumComplete = False

    sTimeHereSince = ConvertToString(oLot.CurrentStep.Steps.Item(1).TimeHereSince, moAppLog)
    sStepNo = oLot.CurrentStep.Steps.Item(1).Id
    sMaxUpdateTime = ""

    If InStr(1, oLot.Id, ".") > 0 Then
        sParentLot = Left(Left(oLot.Id, InStr(1, oLot.Id, ".") - 1), 8)
    Else
        sParentLot = Left(oLot.Id, 8)
    End If
'----
' Condition Checking
'----
    ' <Put your condition checking codes here>...

'----
' Action
'----
    '取得WaferIDList進行Tds_Sum資料確認用
    sSQL = " select " & gsCAT_TLI_WAFERID & " " & _
             " from " & gsCAT_TBL_LOT_INFO & " " & _
            " where " & gsCAT_TLI_LOT_ID & " = '" & oLot.Id & "' " & _
              " and " & gsCAT_TLI_DELETE_FLAG & " = 'N'"
    Set colRaws = moProRawSql.QueryDatabase(sSQL)
    If colRaws.Count > 0 Then
        sWaferIDList = colRaws.Item(1).Item(gsCAT_TLI_WAFERID)
    Else
        Call RaiseError(glERR_ERRMSG, _
                        FormatErrorText(gsETX_ERRMSG, "Lot: " & oLot.Id & " infomation not exist!"))
    End If
    
    If sWaferIDList = "" Then
        Call RaiseError(glERR_ERRMSG, _
                        FormatErrorText(gsETX_ERRMSG, "Lot: " & oLot.Id & " waferid is null!"))
    End If
    sWaferIDList = Replace(sWaferIDList, ",", ";")

    '取得Tds Summary
    sSQL = " select " & gsCAT_TWTS_WAFER_IDN & ", " & _
                        gsCAT_TWTS_TDS_GOOD_DIEN & ", " & _
                        gsCAT_TWTS_REVISE_GOOD_DIEN & ", " & _
                        gsCAT_TWTS_REASON_CODE & ", " & _
                        gsCAT_TWTS_CREATE_TIME & ", " & _
                        gsCAT_TWTS_UPDATE_TIME & " " & _
             " from " & gsCAT_TBL_WS_TDS_SUM & " "
             
    sSQL = sSQL & _
            " where (" & gsCAT_TWTS_TEST_MODE & "," & gsCAT_TWTS_PARENT_LOT_ID & "," & _
                         gsCAT_TWTS_WAFER_IDN & "," & gsCAT_TWTS_STEPNO & "," & gsCAT_TWTS_INSTEPTIME & "," & _
                         "nvl(" & gsCAT_TWTS_TIME_STAMP & ",'NULL'))" & _
                  " IN" & _
                  " (SELECT " & gsCAT_TWTS_TEST_MODE & "," & gsCAT_TWTS_PARENT_LOT_ID & "," & _
                                gsCAT_TWTS_WAFER_IDN & "," & gsCAT_TWTS_STEPNO & "," & gsCAT_TWTS_INSTEPTIME & _
                                ",NVL(MAX(" & gsCAT_TWTS_TIME_STAMP & "),'NULL')" & _
                     " FROM " & gsCAT_TBL_WS_TDS_SUM & _
                    " WHERE " & gsCAT_TWTS_INSTEPTIME & "='" & sTimeHereSince & "' and " & _
                                 gsCAT_TWTS_STEPNO & "='" & sStepNo & "' and " & _
                                 gsCAT_TWTS_PARENT_LOT_ID & "='" & sParentLot & "' and " & _
                                 gsCAT_TWTS_DELETE_FLAG & "='N'" & _
                 " GROUP BY " & gsCAT_TWTS_TEST_MODE & "," & gsCAT_TWTS_PARENT_LOT_ID & "," & gsCAT_TWTS_WAFER_IDN & _
                                "," & gsCAT_TWTS_STEPNO & "," & gsCAT_TWTS_INSTEPTIME & ")"
        
    Set colRaws = moProRawSql.QueryDatabase(sSQL)
    If colRaws Is Nothing Then
        Call RaiseError(glERR_INVALIDOBJECT, _
                    FormatErrorText(gsETX_INVALIDOBJECT, "Collection"))
    End If
    
    For Each oRaws In colRaws
        sWaferID = Trim(CStr(oRaws.Item(gsCAT_TWTS_WAFER_IDN)))
        '以Tds_Sum為準, 再確定是否屬與這批Lot, 以此來處理一般與抽測的Case
        If InStr(sWaferIDList, sWaferID) > 0 Then
            
            sTdsQty = oRaws.Item(gsCAT_TWTS_TDS_GOOD_DIEN)
            sReviseQty = oRaws.Item(gsCAT_TWTS_REVISE_GOOD_DIEN)
            sReasonCode = oRaws.Item(gsCAT_TWTS_REASON_CODE)
            
            '未到齊的檢查:
            '1.TdsQty和ReviseQty都空值
            If (Not (Len(Trim(sTdsQty)) <> 0 Or _
                     Len(Trim(sReviseQty)) <> 0)) Or _
               (Len(Trim(sReviseQty)) <> 0 And Len(Trim(sReasonCode)) = 0) Then
                GoTo ExitHandler 'False
            End If
            
            If oRaws.Item(gsCAT_TWTS_UPDATE_TIME) <> "" Then
                sUpdateTime = oRaws.Item(gsCAT_TWTS_UPDATE_TIME)
            Else
                sUpdateTime = oRaws.Item(gsCAT_TWTS_CREATE_TIME)
            End If
            
            If sUpdateTime > sMaxUpdateTime Then
                sMaxUpdateTime = sUpdateTime
            End If
            
        End If
    Next oRaws

    CheckTdsTestSumComplete = True '都到齊
    '-------
    'Done
    '-------
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
           Case Else
                typErrInfo.sUserText = typErrInfo.sErrDescription
        End Select
        '---- Start of Select Case Block ----
        On Error GoTo ExitHandler:
        Call TransferLogError(typErrInfo.sUserText, _
                              moAppLog, Me.rtfTrace, msMODULE_ID, sProcID)
        Call HandleError(False, typErrInfo, , moAppLog)
    End If
End Function

'================================================================================
' Function: GetWsLotSeq()
'--------------------------------------------------------------------------------
' Description:  <Type your function description here...>
'--------------------------------------------------------------------------------
' Author:       Sam Chen, CIT 2011-11-23
'--------------------------------------------------------------------------------
' Parameters:
'................................................................................
' ARGUMENTS
'   moapplog            (I) [Optional] Valid reference to a clsLogTraceMsg object
'   Argument2           (I) <Description goes here...>
'   Argument3           (O) <Description goes here...>
'
' NON-LOCAL ARGS
'   NonLoc1         (I) <Description goes here...>
'   NonLoc2         (I) <Description goes here...>
'
' RETURN TYPE
'   String          (R) True/False
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
' Copy from Ws.frmSortRun, by Weilun on 20191008
'================================================================================
Private Function GetWsLotSeq(ByVal sLotId As String, ByVal sEqId As String) As String
On Error GoTo ExitHandler:
Dim sProcID As String
Dim typErrInfo As tErrInfo

Dim sSQL As String
Dim colRaws As Collection
Dim colRS As Collection
'-----
' Init
'----
    sProcID = "GetWsLotSeq"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog) '"Entering Function...", moapplog, glLOG_PROC, msMODULE_ID, sProcID)
    
    GetWsLotSeq = "999"  '預設999

'----
' Condition Checking
'----


'----
' Action
'----
    
    sSQL = " Select " & gsCAT_TWEQL_LOTSEQONEQID & " " & _
           " from " & gsCAT_TBL_WS_EQLOTLIST & " " & _
           " where " & gsCAT_TWEQL_LOTID & "='" & sLotId & "' "
        
    sSQL = sSQL & " and " & gsCAT_TWEQL_EQID & " ='" & sEqId & "' " 'Add by Sam on 20121227 for CP Scheduling
    
    'Add by Sam start on 20130925 for ReqNo:JC201300273
    sSQL = sSQL & " and " & gsCAT_TWEQL_COMPLETE & "  is null "
    'Add by Sam end on 20130925 for ReqNo:JC201300273
    
    Set colRaws = moProRawSql.QueryDatabase(sSQL)
    If colRaws Is Nothing Then
        Call RaiseError(glERR_INVALIDOBJECT, _
                    FormatErrorText(gsETX_INVALIDOBJECT, "Collection"))
    End If
    
    If colRaws.Count > 0 Then
        GetWsLotSeq = colRaws.Item(1).Item(gsCAT_TWEQL_LOTSEQONEQID)
    'Mark by Sam Start on 20121227 for CP Scheduling,不再使用
'    Else
'
'        If InStr(sLotId, ".") > 0 Then      '虛擬Lot的情況 EX:3A000000XX_01,3A000000XX_02,條件取母批及排測EqID
'            sLotId = Left(sLotId, InStr(sLotId, ".") - 1)
'
'            sSql = " Select " & gsCAT_TWEQL_LOTSEQONEQID & " " & _
'                   " from " & gsCAT_TBL_WS_EQLOTLIST & " " & _
'                   " where " & gsCAT_TWEQL_PARENTLOTID & "='" & sLotId & "' " & _
'                   " and " & gsCAT_TWEQL_EQID & "='" & sEqID & "' "
'
'            Set colRS = moProRawSql.QueryDatabase(sSql)
'            If colRS.Count > 0 Then
'                GetWsLotSeq = colRS.Item(1).Item(gsCAT_TWEQL_LOTSEQONEQID)
'            End If
'
'        End If
    'Mark by Sam End on 20121227 for CP Scheduling,不再使用
    
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
        GetWsLotSeq = False
    End If
End Function

'================================================================================
'Name       : GetErunTicNO()
'Description: 取得ErunTicNO
'Author     : Weilun huang, MXIC 2019/11/25
'Parameters : N/A
'================================================================================
Private Function GetErunTicNO(ByVal sLotId As String) As String
    On Error GoTo ExitHandler:
    Dim sProcID                     As String
    Dim typErrInfo                  As tErrInfo

    Dim sSQL                        As String
    Dim colRaws                     As Collection
    Dim oRaws                       As Collection

    '-----
    ' Init
    '-----
    sProcID = "GetErunTicNO"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog) '

    GetErunTicNO = "" '預設值

    '--------------------
    ' Condition Checking
    '--------------------

    '-------
    ' Action
    '-------
    sSQL = " select t." & gsCAT_TLI_ERUNTICNO & " " & _
             " from " & gsCAT_TBL_LOT_INFO & " t " & _
            " where t." & gsCAT_TLI_LOT_ID & " = '" & sLotId & "' "
    Set colRaws = moProRawSql.QueryDatabase(sSQL)
    If colRaws.Count > 0 Then
        GetErunTicNO = colRaws.Item(1).Item(gsCAT_TLI_ERUNTICNO)
    End If
    '-------
    'Done
    '-------

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
           Case Else
                typErrInfo.sUserText = typErrInfo.sErrDescription
        End Select
        '---- Start of Select Case Block ----
        On Error GoTo ExitHandler:
        Call TransferLogError(typErrInfo.sUserText, _
                              moAppLog, Me.rtfTrace, msMODULE_ID, sProcID)
        Call HandleError(False, typErrInfo, , moAppLog)
    End If
End Function

'================================================================================
'Name       : InsertAutoProgramFail()
'Description: MES自動化異常寫入CPMonitor
'Author     : Weilun huang, MXIC 2020/05/14 for Project.BE MES Phase 44 – CP 產線增加自動化作業
'Parameters : N/A
'================================================================================
Private Sub InsertAutoProgramFail(ByVal sReciveMsg As String)
    On Error GoTo ExitHandler:
    Dim sProcID                     As String
    Dim typErrInfo                  As tErrInfo

    Dim sSQL                        As String
    Dim colRaws                     As Collection
    Dim oRaws                       As Collection

    Dim asMsg()                     As String
    Dim sEventType                  As String
    Dim sEqId                       As String
    Dim sLotId                      As String
    Dim sErrorMsg                   As String
    Dim sMsgFlag                    As String
    Dim sItemCont                   As String

    '-----
    ' Init
    '-----
    sProcID = "InsertAutoProgramFail"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog) '
    
    sMsgFlag = "/Msg::"
    
    asMsg = Split(sReciveMsg, "/")
    '"AutoDgradeError/5F385600/zDGRADE/Msg::Message"
    
    If UBound(asMsg) < 3 Or InStr(sReciveMsg, sMsgFlag) < 1 Then
        'Error
        GoTo ExitHandler
    End If
    
    sEventType = asMsg(0)
    sLotId = asMsg(1)
    sEqId = asMsg(2)
    sErrorMsg = Mid(sReciveMsg, InStr(sReciveMsg, sMsgFlag) + Len(sMsgFlag))

    '--------------------
    ' Condition Checking
    '--------------------

    '-------
    ' Action
    '-------
    
    sItemCont = ""
    Select Case sEventType
        Case "AutoHoldLotReleaseError"
            sItemCont = gsMONITOR_ITEM_CONT_HLR_ERROR
        Case "AutoDgradeError"
            sItemCont = gsMONITOR_ITEM_CONT_DG_ERROR
        'Add by Weilun on 20211203 for Project.BE 工業 3.5 Phase 39 - CP AGV導入專案 - 建立 AGV 派工 & Status 監控系統 <start>
        Case "AGVLotStartError"
            sItemCont = gsMONITOR_ITEM_CONT_AGV_LS_ERROR
        Case "AutoPrintSetupFormError"
            sItemCont = gsMONITOR_ITEM_CONT_APSU_ERROR
        Case "AutoPrintSetupFormNotice"
            sItemCont = gsMONITOR_ITEM_CONT_APSU_NOTICE
        'Add by Weilun on 20211203 for Project.BE 工業 3.5 Phase 39 - CP AGV導入專案 - 建立 AGV 派工 & Status 監控系統 <end>
    End Select
    
    If sEqId <> "" And sLotId <> "" And sItemCont <> "" Then
        Call TransferLogTrace(DateTime.Now & ": Add '" & sEventType & "' for " & sLotId & "/" & sEqId, _
                              moAppLog, Me.rtfTrace)

        Call InsertMonitorReceive(sEqId, _
                                  sLotId, _
                                  "", _
                                  "", _
                                  gsMONITOR_ITEM_ONEKEYAUTO, _
                                  sItemCont, _
                                  sErrorMsg, _
                                  "1", _
                                  gsCAT_TCMR_PRODTAFLAG)

        '單機台整理
        Call Level1Sort(sEqId)

        '單機台更新
        Call UpdateMonitorSingleEq(sEqId)
    End If
    '-------
    'Done
    '-------

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
           Case Else
                typErrInfo.sUserText = typErrInfo.sErrDescription
        End Select
        '---- Start of Select Case Block ----
        On Error GoTo ExitHandler:
        Call TransferLogError(typErrInfo.sUserText, _
                              moAppLog, Me.rtfTrace, msMODULE_ID, sProcID)
        Call HandleError(False, typErrInfo, , moAppLog)
    End If
End Sub

'================================================================================
'Name       : UpdateETtestMoveOut()
'Description: MES自動化異常寫入CPMonitor
'Author     : Weilun huang, MXIC 2020/05/14 for Project.BE MES Phase 44 – CP 產線增加自動化作業
'Parameters : N/A
'================================================================================
Private Sub UpdateETtestMoveOut(ByVal sReciveMsg As String)
    On Error GoTo ExitHandler:
    Dim sProcID                     As String
    Dim typErrInfo                  As tErrInfo

    Dim sSQL                        As String
    Dim colRaws                     As Collection
    Dim oRaws                       As Collection

    Dim asMsg()                     As String
    Dim sEventType                  As String
    Dim sLotId                      As String
    Dim sNextStepName               As String
    Dim sRecipe                     As String
    Dim sEqId                       As String
    Dim bAutoDGrade                 As Boolean

    '-----
    ' Init
    '-----
    sProcID = "UpdateETtestMoveOut"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog) '
    
    'Modify by Weilun on 20200630 for Project.BE MES Phase 44 – CP 產線增加自動化作業 UAT fix <start>
    asMsg = Split(sReciveMsg, "/")
    ''"UpdateETtestMoveOut/5F385600/BAKE1/委測:T00001" '不再傳站點和Recipe進來
    '"UpdateETtestMoveOut/5F385600"
    
    '由3改為1
    If UBound(asMsg) <> 1 Then
        'Error
        GoTo ExitHandler
    End If
    
    sEventType = asMsg(0)
    sLotId = asMsg(1)
'    sNextStepName = asMsg(2)
'    sRecipe = asMsg(3) '可能是空值(或夾雜空白)
    'Modify by Weilun on 20200630 for Project.BE MES Phase 44 – CP 產線增加自動化作業 UAT fix <end>

    '--------------------
    ' Condition Checking
    '--------------------

    '-------
    ' Action
    '-------
    'Mark by Weilun on 20200630 for Project.BE MES Phase 44 – CP 產線增加自動化作業 UAT fix <start>
    ''需要LotID以及調站前取得下一站和其Recipe(後者為Bake才有)
    'If sLotID <> "" And sNextStepName <> "" Then
    'Mark by Weilun on 20200630 for Project.BE MES Phase 44 – CP 產線增加自動化作業 UAT fix <start>
        
    'Add by Weilun on 20200630 for Project.BE MES Phase 44 – CP 產線增加自動化作業 UAT fix <start>
    If sLotId <> "" Then
        'AutoDGrade改取得轉帳後的
        bAutoDGrade = True
        Call GetNextStepNameAndRecipe(sLotId, sNextStepName, sRecipe, bAutoDGrade)
        '資料更新
        Call TransferLogTrace(DateTime.Now & ": Get StepName: " & sNextStepName & " and Recipe: " & sRecipe, _
                              moAppLog, Me.rtfTrace)
    'Add by Weilun on 20200630 for Project.BE MES Phase 44 – CP 產線增加自動化作業 UAT fix <end>
    
    
        '資料更新
        Call TransferLogTrace(DateTime.Now & ": Update ETtestMoveOut information for " & sLotId, _
                              moAppLog, Me.rtfTrace)

        sSQL = " update " & gsCAT_TBL_CPETEST_MONITOR_RECEIVE & " " & _
                  " set " & gsCAT_TCMR_NEXTSTEPNAME & " = '" & sNextStepName & "', " & _
                            gsCAT_TCMR_RECIPE & " = '" & sRecipe & "', " & _
                            gsCAT_TCMR_UPDATEUSERID & " = 'SYS', " & _
                            gsCAT_TCMR_UPDATETIME & " = to_char(sysdate, 'yyyymmdd hh24miss') || '000' " & _
                " where " & gsCAT_TCMR_LOTID & " = '" & sLotId & "' " & _
                  " and " & gsCAT_TCMR_MONITORITEM & " = '" & gsMONITOR_ITEM_TESTDATA & "' " & _
                  " and " & gsCAT_TCMR_MONITORITEMCONT & " = '" & gsMONITOR_ITEM_CONT_SORTRUN & "' " & _
                  " and " & gsCAT_TCMR_MONITORITEMCONTDESC & " = '" & gsMONITOR_ITEM_DESC_WAITMOVE & "' " & _
                  " and " & gsCAT_TCMR_DELETEFLAG & " = 'N' "
        Call moProRawSql.QueryDatabase(sSQL)

        '機台更新
        sSQL = " select distinct " & gsCAT_TCMR_EQID & " " & _
                 " from " & gsCAT_TBL_CPETEST_MONITOR_RECEIVE & " " & _
                " where " & gsCAT_TCMR_LOTID & " = '" & sLotId & "' " & _
                  " and " & gsCAT_TCMR_MONITORITEM & " = '" & gsMONITOR_ITEM_TESTDATA & "' " & _
                  " and " & gsCAT_TCMR_MONITORITEMCONT & " = '" & gsMONITOR_ITEM_CONT_SORTRUN & "' " & _
                  " and " & gsCAT_TCMR_MONITORITEMCONTDESC & " = '" & gsMONITOR_ITEM_DESC_WAITMOVE & "' " & _
                  " and " & gsCAT_TCMR_DELETEFLAG & " = 'N' "
        Set colRaws = moProRawSql.QueryDatabase(sSQL)

        For Each oRaws In colRaws
            sEqId = oRaws.Item(gsCAT_TCMR_EQID)
        
            '單機台整理
            Call Level1Sort(sEqId)
    
            '單機台更新
            Call UpdateMonitorSingleEq(sEqId)
        Next
    End If
    '-------
    'Done
    '-------

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
           Case Else
                typErrInfo.sUserText = typErrInfo.sErrDescription
        End Select
        '---- Start of Select Case Block ----
        On Error GoTo ExitHandler:
        Call TransferLogError(typErrInfo.sUserText, _
                              moAppLog, Me.rtfTrace, msMODULE_ID, sProcID)
        Call HandleError(False, typErrInfo, , moAppLog)
    End If
End Sub

'================================================================================
'Name       : GetUnTestedWQty()
'Description: 以AutoLotEnd檢查WaferEnd的方式詳細調查剩餘片數(較精準),
'             複製來源是CPAutoLotEnd.CheckWaferEnd
'Author     : Weilun huang, MXIC 2020/07/24
'Parameters : N/A
'================================================================================
Private Function GetUnTestedWQty(ByVal sLotId As String, _
                                 ByVal sEqId As String, _
                                 ByVal sWaferID As String, _
                                 ByVal sStepName As String, _
                                 ByVal sWaferEndTime As String, _
                                 Optional ByVal sReciveMsg As String = "") As Long
    On Error GoTo ExitHandler
    Dim sProcID As String
    Dim typErrInfo As tErrInfo

    Dim asMsg()                     As String
    Dim sParentLotID                As String
    
    Dim sSQL                        As String
    Dim colRaws                     As Collection
    Dim oRaws                       As Collection
    Dim sTimeHereSince              As String
    Dim iIndex                      As Integer
    Dim sWaferIDList                As String
    Dim sCheckList                  As String
    Dim sGetWaferId                 As String
    
    Dim oComp                       As FwComponent
    Dim oLot                        As FwLot
    
    Dim colAllWafer                 As Collection
    Dim oWafer                      As Object
    
    Dim asUnTested()                As String
    Dim sId                         As Variant
    Dim oUnTested                   As Collection

    '-----
    ' Init
    '-----
    sProcID = "GetUnTestedWQty"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog) '

    GetUnTestedWQty = 0
    
    If Trim(sReciveMsg) <> "" Then
        asMsg = Split(sReciveMsg, "/")
        '"WaferEnd/5B224300/TM15/01/SORT2/20180505 2020200000"
        sParentLotID = Left(asMsg(1), 8)
        sEqId = asMsg(2)
        sWaferID = Format(asMsg(3), "00")
        sStepName = asMsg(4)
        sWaferEndTime = Replace(asMsg(5), "~", " ")
    Else '直接呼叫時, 至少需要LotID和StepName
        sParentLotID = Left(sLotId, 8)
        If Trim(sWaferID) <> "" Then
            sWaferID = Format(sWaferID, "00")
        End If
        sWaferEndTime = Replace(sWaferEndTime, "~", " ")
    End If
    
    If Trim(sLotId) = "" Then '必要項, From 機台
        GoTo ExitHandler
    End If
    
    If sParentLotID <> Left(sLotId, 8) Then '如果採用Message, 但與機台上不符時
        GoTo ExitHandler
    End If
    '--------------------
    ' Condition Checking
    '--------------------

    '-------
    ' Action
    '-------
            
    Set oLot = moFwWIP.LotById(sLotId) 'From Input or Message
    If oLot Is Nothing Then
        sLotId = ""
        GoTo ExitHandler
    End If

    sTimeHereSince = ConvertToString(oLot.CurrentStep.Steps.Item(1).TimeHereSince, moAppLog)

    '列出該Lot下的WaferId清單
    sWaferIDList = ""
    
    'Mark by Weilun on 20230607 for Project.新增CP監控識別虛擬併批功能以及修改預測時間機制 <start>
'    If oLot.CurrentStep.Steps.Item(1).Description Like "SORT*" And _
'        Mid(oLot.CurrentStep.Steps.Item(1).Id, 3, 1) <> "2" Then
'        '抽測以Pgm_List的為主
    'Mark by Weilun on 20230607 for Project.新增CP監控識別虛擬併批功能以及修改預測時間機制 <end>
        
        '抽測與一般批都以Pgm_List為主, 以應對產線特殊的操作Case for Project.新增CP監控識別虛擬併批功能以及修改預測時間機制
        sSQL = " select " & gsCAT_TPL_LOT_ID & ", " & _
                            gsCAT_TPL_USERWAFERID & " " & _
                 " from " & gsCAT_TBL_PGM_LIST & " " & _
                " where (" & gsCAT_TPL_LOT_ID & ", " & _
                             gsCAT_TPL_CREATE_TIME & ") in " & _
                       "( select " & gsCAT_TPL_LOT_ID & ", " & _
                           " max(" & gsCAT_TPL_CREATE_TIME & ") " & _
                          " from " & gsCAT_TBL_PGM_LIST & " " & _
                         " where " & gsCAT_TPL_LOT_ID & " ='" & sLotId & "' " & _
                           " and " & gsCAT_TPL_PROD_TYPE & " = 'WAFER' " & _
                           " and " & gsCAT_TPL_SOURCE & " <> 'Setup' " & _
                           " and " & gsCAT_TPL_CREATE_TIME & " >= '" & sTimeHereSince & "' " & _
                           " and " & gsCAT_TPL_DELETE_FLAG & " ='N' " & _
                      " group by " & gsCAT_TPL_LOT_ID & ") "
                           
        
        Set colRaws = moProRawSql.QueryDatabase(sSQL)
        If colRaws.Count > 0 Then
            sWaferIDList = colRaws.Item(1).Item(gsCAT_TPL_USERWAFERID)
        Else
            'Modify by Weilun on 20230607 for Project.新增CP監控識別虛擬併批功能以及修改預測時間機制
'            Call TransferLogError(DateTime.Now & ": 抽測Case, 在Tbl_Pgm_List找不到資料.", _
'                                  moAppLog, Me.rtfTrace, msMODULE_ID, sProcID)
            Call TransferLogError(DateTime.Now & ": " & sProcID & ", 在Tbl_Pgm_List找不到資料.", _
                                  moAppLog, Me.rtfTrace, msMODULE_ID, sProcID)
            GoTo ExitHandler
        End If
        
        If sWaferIDList = "" Then
            'Modify by Weilun on 20230607 for Project.新增CP監控識別虛擬併批功能以及修改預測時間機制
'            Call TransferLogError(DateTime.Now & ": 抽測Case, Tbl_Pgm_List有誤.", _
'                                  moAppLog, Me.rtfTrace, msMODULE_ID, sProcID)
            Call TransferLogError(DateTime.Now & ": " & sProcID & ", Tbl_Pgm_List有誤.", _
                                  moAppLog, Me.rtfTrace, msMODULE_ID, sProcID)
            GoTo ExitHandler
        End If
        
    'Mark by Weilun on 20230607 for Project.新增CP監控識別虛擬併批功能以及修改預測時間機制 <start>
'    Else '原本的
'        For iIndex = 1 To 25
'            Set oComp = oLot.Components.Item(Format(iIndex, "00"))
'            If Not oComp Is Nothing Then
'                If oComp.Status <> "Scrapped" Then
'                    sWaferIDList = sWaferIDList & "," & oComp.Id
'                End If
'            End If
'        Next iIndex
'    End If
    'Mark by Weilun on 20230607 for Project.新增CP監控識別虛擬併批功能以及修改預測時間機制 <end>
    
    If Left(sWaferIDList, 1) = "," Then
        sWaferIDList = Mid(sWaferIDList, 2)
    End If

    '準備檢查用清單
    sWaferIDList = SortWaferID(sWaferIDList, ",")
    sCheckList = Replace(sWaferIDList, ",", " ")    '最後搭配Trim

    '搜尋tbl_cp_wafer_end, 尋找該Lot在自動化開始後相關的資料
    'TBL_CP_WAFER_END 之 WaferEndTime 應大於 TBL_CP_AUTO_COMPLETE.CreateTime
    sSQL = " select " & gsCAT_TCWE_WAFERID & " " & _
             " from " & gsCAT_TBL_CP_WAFER_END & " " & _
            " where " & gsCAT_TCWE_LOTID & " like '" & sParentLotID & "%' " & _
              " and " & gsCAT_TCWE_WAFERID & " is not null " & _
              " and " & gsCAT_TCWE_STEPNAME & " = '" & sStepName & "' " & _
              " and " & gsCAT_TCWE_WAFERENDTIME & " > '" & sTimeHereSince & "' " & _
              " and " & gsCAT_TCWE_DELETEFLAG & " = 'N'"

    Set colRaws = moProRawSql.QueryDatabase(sSQL)
    If colRaws.Count > 0 Then
        For Each oRaws In colRaws
            'Item("WaferId")改Item(gsCAT_TCWE_WAFERID)' Modify by Weilun on 20180914 for ReqNo.BE#201800559
            sGetWaferId = Format(oRaws.Item(gsCAT_TCWE_WAFERID), "00")
            If IsNumeric(sGetWaferId) = True And Len(sGetWaferId) = 2 Then
                sCheckList = Replace(sCheckList, sGetWaferId, "")  '有找到, 從檢查清單中扣除
            End If
        Next oRaws
    End If

    '檢查清單扣掉傳入的WaferId
    If IsNumeric(sWaferID) = True And Len(sWaferID) = 2 And sWaferEndTime > sTimeHereSince Then
        sCheckList = Replace(sCheckList, sWaferID, "")
    End If
    
    'Add by Weilun on 20180912 for ReqNo.BE#201800559<start>
    '增加檢查階段, WaferEnd找不到的再去找Tds Summary找
    If Trim(sCheckList) <> "" Then
    
        Set colAllWafer = GetTdsTestSum(oLot.Id, sTimeHereSince, oLot.CurrentStep.Steps.Item(1).Id)
        
        '檢查Summary ReviseQty與ReasonCode
        For Each oWafer In colAllWafer
            'TDS Good Dien有值
            '或(Revise Good Dien有值且Reasoncode有值)
            If Len(Trim(oWafer.Item(gsCAT_TWTS_TDS_GOOD_DIEN))) <> 0 Or _
               (Len(Trim(oWafer.Item(gsCAT_TWTS_REVISE_GOOD_DIEN))) <> 0 And _
                Len(Trim(oWafer.Item(gsCAT_TWTS_REASON_CODE))) <> 0) Then
                
                sGetWaferId = Format(oWafer.Item(gsCAT_TWTS_WAFER_IDN), "00")
                If IsNumeric(sGetWaferId) = True And Len(sGetWaferId) = 2 Then
                    sCheckList = Replace(sCheckList, sGetWaferId, "")  '有找到, 從檢查清單中扣除
                End If

                
            End If
        Next oWafer
    End If
    'Add by Weilun on 20180912 for ReqNo.BE#201800559<end>
    
    '尚未到齊
    If Trim(sCheckList) <> "" Then
        Set oUnTested = New Collection
        For Each sId In Split(sCheckList)
            If Trim(sId) <> "" Then
                oUnTested.Add (sId)
            End If
        Next
    
        GetUnTestedWQty = oUnTested.Count
    End If
    '-------
    'Done
    '-------
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
           Case Else
                typErrInfo.sUserText = typErrInfo.sErrDescription
        End Select
        '---- Start of Select Case Block ----
        On Error GoTo ExitHandler:
        Call TransferLogError(typErrInfo.sUserText, _
                              moAppLog, Me.rtfTrace, msMODULE_ID, sProcID)
        Call HandleError(False, typErrInfo, , moAppLog)
    End If
End Function

'================================================================================
'Name       : GetTdsTestSum()
'Description: 取得WS_TDS_SUM顆數資料, 修改自WS.frmSortLotEnd.GetTdsTestSum()
'Author     : Weilun Huang, MXIC 2018/03/23
'Parameters :
'================================================================================
Private Function GetTdsTestSum(ByVal sLotId As String, ByVal sTimeHereSince As String, ByVal sStepNo As String) As Collection
On Error GoTo ExitHandler:
Dim sProcID                 As String
Dim typErrInfo              As tErrInfo
Dim colRaws                 As Collection
Dim oRaws                   As FwStrings
Dim sTable                  As String
Dim sColumn                 As String
Dim sWhere                  As String
Dim sOther                  As String
Dim oLot                    As FwLot
Dim iIndex                  As Integer
Dim oComp                   As FwComponent

Dim sParentLot              As String

Dim colAllWafer             As Collection
Dim oWafer                  As Collection
Dim iCount                  As Integer
Dim bFoundData              As Boolean
'----
' Init
'----
    sProcID = "GetTdsTestSum"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog)

'----
' Condition Checking
'----
    ' <Put your condition checking codes here>...

'----
' Action
'----
    If InStr(1, sLotId, ".") > 0 Then
        sParentLot = Left(Left(sLotId, InStr(1, sLotId, ".") - 1), 8)
    Else
        sParentLot = Left(sLotId, 8)
    End If
    
    'Define table name
    sTable = gsCAT_TBL_WS_TDS_SUM
    
    'Define column name
    sColumn = gsCAT_TWTS_WAFER_IDN & "," & _
              gsCAT_TWTS_TDS_GOOD_DIEN & "," & _
              gsCAT_TWTS_REVISE_GOOD_DIEN & "," & _
              gsCAT_TWTS_REASON_CODE & " "

    'Define where
    sWhere = "(" & gsCAT_TWTS_TEST_MODE & "," & gsCAT_TWTS_PARENT_LOT_ID & "," & _
                gsCAT_TWTS_WAFER_IDN & "," & gsCAT_TWTS_STEPNO & "," & gsCAT_TWTS_INSTEPTIME & "," & _
                "nvl(" & gsCAT_TWTS_TIME_STAMP & ",'NULL'))" & _
             " IN" & _
             " (SELECT " & gsCAT_TWTS_TEST_MODE & "," & gsCAT_TWTS_PARENT_LOT_ID & "," & _
                gsCAT_TWTS_WAFER_IDN & "," & gsCAT_TWTS_STEPNO & "," & gsCAT_TWTS_INSTEPTIME & _
                ",NVL(MAX(" & gsCAT_TWTS_TIME_STAMP & "),'NULL')" & _
             " FROM " & gsCAT_TBL_WS_TDS_SUM & _
             " WHERE " & _
            gsCAT_TWTS_INSTEPTIME & "='" & sTimeHereSince & "' and " & _
            gsCAT_TWTS_STEPNO & "='" & sStepNo & "' and " & _
            gsCAT_TWTS_PARENT_LOT_ID & "='" & sParentLot & "' and " & _
            gsCAT_TWTS_DELETE_FLAG & "='N'" & _
            " GROUP BY " & gsCAT_TWTS_TEST_MODE & "," & gsCAT_TWTS_PARENT_LOT_ID & "," & gsCAT_TWTS_WAFER_IDN & _
            "," & gsCAT_TWTS_STEPNO & "," & gsCAT_TWTS_INSTEPTIME & ")"
        
    sOther = "order by " & "nvl(" & gsCAT_TWTS_TIME_STAMP & ",' ') desc"
    
    Set colRaws = SelectRawSqlTxn(moAppLog, moFwWIP, moFwWF, moCwMbx, "CAT", sTable, sColumn, sWhere, sOther)
    If colRaws Is Nothing Then
        Call RaiseError(glERR_INVALIDOBJECT, _
                    FormatErrorText(gsETX_INVALIDOBJECT, "Collection"))
    End If
    
    Set oLot = moFwWIP.LotById(sLotId)
    Set colAllWafer = New Collection
    iCount = 0
    
    For iIndex = 1 To 25
        Set oComp = oLot.Components.Item(Format(iIndex, "00"))
        If Not oComp Is Nothing Then
            If oComp.Status <> "Scrapped" Then
                Set oWafer = New Collection
                bFoundData = False
                oWafer.Add oComp.Id, gsCAT_TWTS_WAFER_IDN

                '在搜尋結果裡尋找
                For Each oRaws In colRaws
                    If Trim(CStr(oRaws.Item(gsCAT_TWTS_WAFER_IDN))) = oComp.Id Then
                        bFoundData = True
                        oWafer.Add oRaws.Item(gsCAT_TWTS_TDS_GOOD_DIEN), gsCAT_TWTS_TDS_GOOD_DIEN
                        oWafer.Add oRaws.Item(gsCAT_TWTS_REVISE_GOOD_DIEN), gsCAT_TWTS_REVISE_GOOD_DIEN
                        oWafer.Add oRaws.Item(gsCAT_TWTS_REASON_CODE), gsCAT_TWTS_REASON_CODE
                        Exit For
                    End If
                Next oRaws
                
                '找不到時補齊欄位
                If bFoundData = False Then
                    oWafer.Add "", gsCAT_TWTS_TDS_GOOD_DIEN
                    oWafer.Add "", gsCAT_TWTS_REVISE_GOOD_DIEN
                    oWafer.Add "", gsCAT_TWTS_REASON_CODE
                End If
                
                iCount = iCount + 1
                colAllWafer.Add oWafer, CStr(iCount)
                Set oWafer = Nothing
            End If
        End If
        Set oComp = Nothing
    Next iIndex
    
    Set GetTdsTestSum = colAllWafer
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
                typErrInfo.sUserText = typErrInfo.sErrDescription
        End Select
        '---- Start of Select Case Block ----
        On Error GoTo ExitHandler:
        Call TransferLogError(typErrInfo.sUserText, _
                      moAppLog, Me.rtfTrace, msMODULE_ID, sProcID)
        Call HandleError(False, typErrInfo, , moAppLog)
    End If
End Function

'================================================================================
'Name       : CreateSignalRClient()
'Description: 建立SignalR以進行連線
'Author     : Weilun huang, MXIC 2020/08/19 for Req.VB 嘗試引用 C# DLL
'Parameters : N/A
'================================================================================
Private Function CreateSignalRClient() As Boolean
    On Error GoTo ExitHandler:
    Dim sProcID As String
    Dim typErrInfo As tErrInfo
    
    '定位Frame
    Dim lButtonAreaHeight   As Long
    Dim lBroderHeight       As Long
    Dim lBroderWidth        As Long
    '-----
    ' Init
    '-----
    sProcID = "CreateSignalRClient"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog)
    
    CreateSignalRClient = False
    
    '--------------------
    ' Condition Checking
    '--------------------
    ' <Put your condition checking codes here>...

    '-------
    ' Action
    '-------
    On Error Resume Next
    Set moClient = CreateObject("MXIC.BE.SignalRClient.clsMain")
    If moClient Is Nothing Then
        UtShowMsgBox "Cat not Create SignalR Object"
        GoTo ExitHandler
    Else
        moClient.WorkDir = App.Path
        If moClient.Init() = True Then
            CreateSignalRClient = True
        Else
            UtShowMsgBox "Cat not Init SignalR Object"
            GoTo ExitHandler
        End If
    End If
    
    
    '-------
    'Done
    '-------
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
           Case Else
                typErrInfo.sUserText = typErrInfo.sErrDescription
        End Select
        '---- Start of Select Case Block ----
        On Error GoTo ExitHandler:
        Call HandleError(False, typErrInfo, , moAppLog, , True)
    End If
End Function


'================================================================================
'Name       : UpdateAGVDispatchEndTime()
'Description:
'Author     : Weilun huang, MXIC 2021/11/17 for Project.BE 工業 3.5 Phase 39 - CP AGV導入專案 - 建立 AGV 派工 & Status 監控系統
'Parameters : N/A
'================================================================================
Private Sub UpdateAGVDispatchEndTime(ByVal sEqId As String, ByVal sEndTime As String)
    On Error GoTo ExitHandler:
    Dim sProcID                     As String
    Dim typErrInfo                  As tErrInfo

    Dim sSQL                        As String
    Dim colRaws                     As Collection
    Dim oRaws                       As Collection

    '-----
    ' Init
    '-----
    sProcID = "UpdateAGVDispatchEndTime"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog) '
    

    '--------------------
    ' Condition Checking
    '--------------------

    '-------
    ' Action
    '-------
    '機台ID必要, EndTime給空值表示清空
    If sEqId <> "" Then
        '資料更新
        Call TransferLogTrace(DateTime.Now & ": Update AGV Dispatch EndTime information for " & sEqId, _
                              moAppLog, Me.rtfTrace)

        'sEndTime使用toChar, 取消單引號,Modify by Weilun on 20220616 for AGV fix
        sSQL = " update " & gsCAT_TBL_AGV_DISPATCH_ORDER & " " & _
                  " set " & gsCAT_TADO_PREVIOUS_LOT_ENDTIME & " = " & sEndTime & ", " & _
                            gsCAT_TADO_UPDATEUSERID & " = 'SYS', " & _
                            gsCAT_TADO_UPDATETIME & " = to_char(sysdate, 'yyyymmdd hh24miss') || '000' " & _
                " where " & gsCAT_TADO_EQID & " = '" & sEqId & "' " & _
                  " and " & gsCAT_TADO_AGVID & " is null " & _
                  " and " & gsCAT_TADO_DELETEFLAG & " = 'N' "
        Call moProRawSql.QueryDatabase(sSQL)

'        For Each oRaws In colRaws
'            sEqID = oRaws.Item(gsCAT_TCMR_EQID)
'
'            '單機台整理
'            Call Level1Sort(sEqID)
'
'            '單機台更新
'            Call UpdateMonitorSingleEq(sEqID)
'        Next
    End If
    '-------
    'Done
    '-------

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
           Case Else
                typErrInfo.sUserText = typErrInfo.sErrDescription
        End Select
        '---- Start of Select Case Block ----
        On Error GoTo ExitHandler:
        Call TransferLogError(typErrInfo.sUserText, _
                              moAppLog, Me.rtfTrace, msMODULE_ID, sProcID)
        Call HandleError(False, typErrInfo, , moAppLog)
    End If
End Sub


'================================================================================
'Name       : GetUnloadType()
'Description:
'Author     : Weilun huang, MXIC 2021/11/25 for Project.BE 工業 3.5 Phase 39 - CP AGV導入專案 - 建立 AGV 派工 & Status 監控系統
'Parameters : N/A
'================================================================================
Private Function GetUnloadType(ByVal sEqId As String, ByVal sUnloadLotID As String) As String
    On Error GoTo ExitHandler:
    Dim sProcID                     As String
    Dim typErrInfo                  As tErrInfo

    Dim sSQL                        As String
    Dim colRaws                     As Collection
    Dim oRaws                       As Collection
    
    Dim sLotSQL                     As String
    Dim colLotRaws                  As Collection
    
    Dim sNextLotID                  As String
    Dim sUnloadType                 As String

    '-----
    ' Init
    '-----
    sProcID = "GetUnloadType"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog) '
    
    sUnloadType = "TA" '預設
    GetUnloadType = sUnloadType

    '--------------------
    ' Condition Checking
    '--------------------

    '-------
    ' Action
    '-------
    '機台ID必要, EndTime給空值表示清空
    If sEqId <> "" And sUnloadLotID <> "" Then
        '確認該機台有CAT維護
        sSQL = " select " & gsCAT_TSVC_VALUE & " " & _
                 " from " & gsCAT_TBL_SINGLE_VALUE_CONTROL & " " & _
                " where " & gsCAT_TSVC_CATEGORY & " = 'WS_AGV_TESTER' " & _
                  " and " & gsCAT_TSVC_ATTRIBUTE & " = 'EQID' " & _
                  " and " & gsCAT_TSVC_VALUE & " = '" & sEqId & "' " & _
                  " and " & gsCAT_TSVC_DELETEFLAG & " = 'N' "
        Set colRaws = moProRawSql.QueryDatabase(sSQL)
        
        '取得該機台最早派工的Lot, 且為WAITR
        If colRaws.Count > 0 Then
            sLotSQL = " select tla." & gsCAT_TLATT_LOTID & ", " & _
                             " tli." & gsCAT_TLI_SPLITASSIGNEQIDTIME & " " & _
                    " from " & gsCAT_TBL_LOT_ATTRIBUTE & " tla, " & _
                               gsCAT_TBL_LOT_INFO & " tli, " & _
                           " fwlot b, " & _
                           " (SELECT FROMID, VALDATA SplitAssignEqId " & _
                              " From FWLOT_PN2M " & _
                             " WHERE KEYDATA = 'SplitAssignEqId') F19 " & _
                   " where f19.SplitAssignEqId = '" & sEqId & "' " & _
                     " and b.sysid = f19.fromid(+) " & _
                     " and b.componentunits = '" & gsLOTUNIT_WAFER & "' " & _
                     " and tli." & gsCAT_TLI_SPLITASSIGNEQIDTIME & " is not null " & _
                     " and tli." & gsCAT_TLI_LOT_ID & " = tla." & gsCAT_TLATT_LOTID & " " & _
                     " and tla." & gsCAT_TLATT_LOTID & " = b.appid " & _
                     " and tla." & gsCAT_TLATT_STEPNAME & " like 'SORT%' " & _
                     " and tla." & gsCAT_TLATT_RESOURCETYPE & " = 'WTESTER' " & _
                     " and tla." & gsCAT_TLATT_LOTSTATUS1 & " = '" & gsLOTSTATUS_WAITR & "' " & _
                     " and tla." & gsCAT_TLATT_STAGE & " = '" & gsSTAGE_WS & "' " & _
                " order by tli." & gsCAT_TLI_SPLITASSIGNEQIDTIME & " asc "
            Set colLotRaws = moProRawSql.QueryDatabase(sLotSQL)
            
            If colLotRaws.Count > 0 Then
                sNextLotID = colLotRaws.Item(1).Item(gsCAT_TLATT_LOTID)
                Call LogTrace(DateTime.Now & ": Get " & sEqId & " NextLotID: " & sNextLotID, _
                              moAppLog)
            End If
            
            '比較WaferSize
            If sNextLotID <> "" Then
                sSQL = "select tlatt." & gsCAT_TLATT_LOTID & ", " & _
                               " tli." & gsCAT_TLI_ERUNTICNO & ", " & _
                               " tpb." & gsCAT_TPB_WAFER_SIZE & " " & _
                        " from " & gsCAT_TBL_LOT_ATTRIBUTE & " tlatt, " & _
                                   gsCAT_TBL_LOT_INFO & " tli, " & _
                                   gsCAT_TBL_IPN_MASTER & " tim, " & _
                                   gsCAT_TBL_PROD_BODY & " tpb " & _
                       " where tim." & gsCAT_TIM_PRODBODY & " = tpb." & gsCAT_TPB_PROD_BODY & " " & _
                         " and tlatt." & gsCAT_TLATT_LOTID & " = tli." & gsCAT_TLI_LOT_ID & " " & _
                         " and tlatt." & gsCAT_TLATT_IPN & " = tim." & gsCAT_TIM_IPN & " " & _
                         " and tlatt." & gsCAT_TLATT_LOTID & " in ('" & sUnloadLotID & "', '" & _
                                                                                sNextLotID & "') "
                Set colLotRaws = moProRawSql.QueryDatabase(sSQL)
                If colLotRaws.Count > 1 Then '需要比較兩筆的資料, 以後User會不會比較第三筆就不知道了
                    '無順序比較是否都為12吋
                    'gsCAT_TLI_ERUNTICNO都需要為空值
                    If colLotRaws.Item(1).Item(gsCAT_TPB_WAFER_SIZE) = "12" And _
                       colLotRaws.Item(1).Item(gsCAT_TLI_ERUNTICNO) = "" And _
                       colLotRaws.Item(2).Item(gsCAT_TPB_WAFER_SIZE) = "12" And _
                       colLotRaws.Item(2).Item(gsCAT_TLI_ERUNTICNO) = "" Then
                        sUnloadType = "AGV"
                    End If
                End If
            End If
        End If
    End If
    
    Call LogTrace(DateTime.Now & ": Get UnloadType: " & sUnloadType & " for " & sEqId, _
                  moAppLog)
                            
    GetUnloadType = sUnloadType
    '-------
    'Done
    '-------

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
           Case Else
                typErrInfo.sUserText = typErrInfo.sErrDescription
        End Select
        '---- Start of Select Case Block ----
        On Error GoTo ExitHandler:
        Call TransferLogError(typErrInfo.sUserText, _
                              moAppLog, Me.rtfTrace, msMODULE_ID, sProcID)
        Call HandleError(False, typErrInfo, , moAppLog)
    End If
End Function

'================================================================================
'Name       : UnloadTypeReview()
'Description:
'Author     : Weilun huang, MXIC 2021/11/30 for Project.BE 工業 3.5 Phase 39 - CP AGV導入專案 - 建立 AGV 派工 & Status 監控系統
'Parameters : N/A
'================================================================================
Private Function UnloadTypeReview(ByVal sReciveMsg As String) As Boolean
    On Error GoTo ExitHandler
    Dim sProcID As String
    Dim typErrInfo As tErrInfo

    Dim asMsg() As String

    Dim sEventType          As String
    Dim sReviewLotID        As String
    Dim sCurrentLotID       As String
    Dim sSID                As String
    Dim sEqId               As String
    Dim sReviewFlag         As String
    Dim oEquipment          As FwEquipment
    Dim bClearUnloadType    As Boolean

    Dim sSQL                As String
    Dim colRaws             As Collection
    Dim oRaws               As Collection

    '-----
    ' Init
    '-----
    sProcID = "UnloadTypeReview"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog) '
    
    UnloadTypeReview = False

    asMsg = Split(sReciveMsg, "/")

    sEventType = asMsg(0)
    sEqId = asMsg(1)
    sSID = asMsg(2)
    sReviewFlag = asMsg(3)
    
    bClearUnloadType = False

    '--------------------
    ' Condition Checking
    '--------------------

    '-------
    ' Action
    '-------
    If sEqId <> "" And sSID <> "" And sReviewFlag = "Y" Then
        sSQL = " select " & gsCAT_TCMR_LOTID & " " & _
                 " from " & gsCAT_TBL_CPETEST_MONITOR_RECEIVE & " " & _
                " where " & gsCAT_TCMR_SID & " = '" & sSID & "' " & _
                  " and " & gsCAT_TCMR_MONITORITEMCONTDESC & " = '" & gsMONITOR_ITEM_DESC_UNLOAD & "' " & _
                  " and " & gsCAT_TCMR_REVIEWFLAG & " = 'Y' " & _
                  " and " & gsCAT_TCMR_DELETEFLAG & " = 'N' "
        Set colRaws = moProRawSql.QueryDatabase(sSQL)
        If colRaws.Count > 0 Then
            sReviewLotID = colRaws.Item(1).Item(gsCAT_TCMR_LOTID)
        Else
            GoTo ExitHandler '非LotEnd Review事件
        End If
        
        
        '規則為, LotEnd Review事件的LotID與Level 1上的不一樣時才不清
        '理論上LotEnd後, Level 1上的LotID是清空或者換批了, 不過也可能有特殊狀況
        bClearUnloadType = True
        Set oEquipment = moFwMDL.EquipmentById(sEqId)
        If Not oEquipment Is Nothing Then
            'Mark by Weilun on 20230606 for Project.新增CP監控識別虛擬併批功能以及修改預測時間機制
'            sCurrentLotID = oEquipment.CustomAttributes(gsEQP_CUSTOMATTR_CUR_LOT_ID)
            
            'Modify by Weilun on 20230606 for Project.新增CP監控識別虛擬併批功能以及修改預測時間機制
            'If sReviewLotID <> "" And sReviewLotID = sCurrentLotID Then
            If sReviewLotID <> "" And GetStrPosInArray(Replace(oEquipment.CustomAttributes(gsEQP_CUSTOMATTR_CUR_LOT_ID), ";", ","), sReviewLotID, ",") > 0 Then
                bClearUnloadType = False 'LotID一致, 不清空UnloadType
            End If
        End If

        '清空
        If bClearUnloadType = True Then
            UnloadTypeReview = True
            Call TransferLogTrace(DateTime.Now & ": Clear UnloadType for " & sEqId, _
                                  moAppLog, Me.rtfTrace)
                                  
            sSQL = " update " & gsCAT_TBL_CPETEST_MONITOR_LEVEL1 & " " & _
                      " set " & gsCAT_TCML1_UNLOADTYPE & " = '', " & _
                                gsCAT_TCML1_UPDATEUSERID & " = 'SYS', " & _
                                gsCAT_TCML1_UPDATETIME & " = to_char(sysdate, 'yyyymmdd hh24miss') || '000' " & _
                    " where " & gsCAT_TCML1_EQID & " = '" & sEqId & "' " & _
                      " and " & gsCAT_TCML1_DELETEFLAG & " = 'N' "
            Call moProRawSql.QueryDatabase(sSQL)
            
            '雖然為Review事件, 不過LotEnd事件理論上不多, 事件比較需要即時反應
            '單機台整理
            Call Level1Sort(sEqId, sSID, sReviewFlag)
    
            '單機台更新
            Call UpdateMonitorSingleEq(sEqId)
        End If
    End If
    '-------
    'Done
    '-------
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
            Case Else
                typErrInfo.sUserText = typErrInfo.sErrDescription
        End Select
        '---- Start of Select Case Block ----
        On Error GoTo ExitHandler
        Call TransferLogError(typErrInfo.sUserText, _
                              moAppLog, Me.rtfTrace, msMODULE_ID, sProcID)
        Call HandleError(False, typErrInfo, , moAppLog)
    End If
End Function

Private Sub TimerSorterControl_Timer()
    Dim lIndex  As Long
    Dim oTemp   As Collection
    
    mlSCTimer = mlSCTimer + 1
    If mlSCTimer Mod 2 = 0 Then
        '搜尋並執行SC
        If Not moSC_Main Is Nothing Then
            For lIndex = 1 To moSC_Main.Count
                InsertSorterControl (moSC_Main.Item(lIndex))
            Next
        End If
        
        '做完清空, 接著Timer清0讓資料可以進來
        Set moSC_Main = New Collection
        mlSCTimer = 0
        
        Set oTemp = moSC_Sub
        Set moSC_Sub = New Collection 'Sub必須清空
        '這裡用一筆一筆個別加入Main, 而不用Set Main = Sub, 因為剛調完Flag說不定剛好有資料進來
        If Not oTemp Is Nothing Then
            For lIndex = 1 To oTemp.Count
                Call ListSC(oTemp.Item(lIndex))   '由於mlSCTimer歸0, 呼叫這個是加入moSC_Main
            Next
        End If
        
        If Not moSC_Main Is Nothing And Not moSC_Sub Is Nothing Then
            If moSC_Main.Count = 0 And moSC_Sub.Count = 0 Then
                '待辦(Main)和Buffer(Sub)都沒資料可執行時, 這個Timer就可以先關了
                '直到有新資料進來或是整體檢查時特別檢查一下
                TimerSorterControl.Enabled = False
            End If
        End If
    End If
    
End Sub

'================================================================================
'Name       : ListSC()
'Description: List SorterControl Info
'Author     : Weilun Huang, MXIC, 2022/12/30,
'Parameters : N/A
'================================================================================
Private Sub ListSC(ByVal sMessage As String)
    On Error GoTo ExitHandler:
    
    Dim sProcID             As String
    Dim typErrInfo          As tErrInfo
    
    Dim lIndex              As Long
    Dim bFound              As Boolean
    '-----
    ' Init
    '-----
    sProcID = "ListSC"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog)
    
    '----
    ' Condition Checking
    '----
    ' <Put your condition checking codes here>...
    
    '-------
    ' Action
    '-------
    
    '尋找是否已列入系統中
    bFound = False
    If Not moSC_Main Is Nothing Then
        For lIndex = 1 To moSC_Main.Count
            If moSC_Main.Item(lIndex) = sMessage Then
                bFound = True
            End If
        Next
    End If
    If bFound = False Then
        If Not moSC_Sub Is Nothing Then
            For lIndex = 1 To moSC_Sub.Count
                If moSC_Sub.Item(lIndex) = sMessage Then
                    bFound = True
                End If
            Next
        End If
    End If
    
    If bFound = False Then
        '每10秒執行一次, 執行前5秒改加入代辦清單
        '這裡Timer也會呼叫, 其中mlSCTimer會是0, 將加入主清單
        If mlSCTimer < 1 Then
            If Not moSC_Main Is Nothing Then
                moSC_Main.Add (sMessage)
            End If
        Else
            If Not moSC_Sub Is Nothing Then
                moSC_Sub.Add (sMessage)
            End If
        End If
    End If

                          
    '-------
    'Done
    '-------
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
           Case Else
                typErrInfo.sUserText = typErrInfo.sErrDescription
        End Select
        '---- Start of Select Case Block ----
        On Error GoTo ExitHandler:
        Call TransferLogError(typErrInfo.sUserText, _
                              moAppLog, Me.rtfTrace, msMODULE_ID, sProcID)
        Call HandleError(False, typErrInfo, , moAppLog)
    End If
End Sub

'================================================================================
'Name       : InsertSorterControl()
'Description: MES寫入SorterControl相關資料
'Author     : Weilun huang, MXIC 2022/12/30 for Req.SorterControl_Monitor
'Parameters : N/A
'================================================================================
Private Sub InsertSorterControl(ByVal sReciveMsg As String)
    On Error GoTo ExitHandler:
    Dim sProcID                     As String
    Dim typErrInfo                  As tErrInfo

    Dim sSQL                        As String
    Dim colRaws                     As Collection
    Dim oRaws                       As Collection

    Dim asMsg()                     As String
    Dim sEventType                  As String
    Dim sEqId                       As String
    Dim sLotId                      As String
    Dim sItemCont                   As String
    Dim sItemContDesc               As String
    Dim bInsert                     As String

    '-----
    ' Init
    '-----
    sProcID = "InsertSorterControl"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog) '

    asMsg = Split(sReciveMsg, "/")
    '"InsertSorterControl/5F385600/zSorter/20221230 112400000" '取消時間
    '"InsertSorterControl/5F385600/zSorter"
    
    'If UBound(asMsg) < 3 Then 'Modify by Weilun on 20230301 for Req.SorterControl_Monitor uat
    If UBound(asMsg) < 2 Then
        'Error
        GoTo ExitHandler
    End If
    
    sEventType = asMsg(0)

    '--------------------
    ' Condition Checking
    '--------------------

    '-------
    ' Action
    '-------
    
    sItemCont = ""
    Select Case sEventType
        Case "InsertSorterControl"
            bInsert = False
            sLotId = asMsg(1)
            sSQL = " select " & gsCAT_TLI_MERGETYPE & ", " & _
                                gsCAT_TLI_VIRTUALLOTID & " " & _
                     " from " & gsCAT_TBL_LOT_INFO & " " & _
                    " where " & gsCAT_TLI_LOT_ID & " = '" & sLotId & "' "
            Set colRaws = moProRawSql.QueryDatabase(sSQL)
            If colRaws.Count > 0 Then
                '虛擬併批更正為V開頭即可, 後續大量引用看要不要再調整
'                If colRaws.Item(1).Item(gsCAT_TLI_MERGETYPE) = "VirtualMerge" And _
'                    Trim(colRaws.Item(1).Item(gsCAT_TLI_VIRTUALLOTID)) <> "" Then
'                   bInsert = True
'                End If

                '這裡沒引用太多模組, 先把Type寫死
                If colRaws.Item(1).Item(gsCAT_TLI_MERGETYPE) = "CPSMALLLOT" Then
                   bInsert = True
                End If
            End If
            
            '虛擬併批為V開頭即可, Add by Weilun on 20230111 for Req.SorterControl_Monitor
            If sLotId Like "V*" Then
                bInsert = True
            End If
            
            If bInsert = False Then
                Call TransferLogTrace(DateTime.Now & ": LotId " & sLotId & "不為虛擬併批或測後併批, 不寫入Monitor", _
                                      moAppLog, Me.rtfTrace)
                GoTo ExitHandler
            End If
            
            sEqId = asMsg(2)
            sItemCont = gsCAT_TBL_SORTER_CONTROL
            sItemContDesc = "Lot: " & sLotId & " wait merge/split lot"
    End Select
    
    If sEqId <> "" And sLotId <> "" And sItemCont <> "" Then
        '搜尋已在Monitor上但是未結案的事件, 有的話不再加入
        sSQL = " select " & gsCAT_TCMR_LOTID & " " & _
                 " from " & gsCAT_TBL_CPETEST_MONITOR_RECEIVE & " " & _
                " where " & gsCAT_TCMR_EQID & " = '" & sEqId & "' " & _
                  " and " & gsCAT_TCMR_LOTID & " = '" & sLotId & "' " & _
                  " and " & gsCAT_TCMR_MONITORITEM & " = '" & gsMONITOR_ITEM_TESTDATA & "' " & _
                  " and " & gsCAT_TCMR_MONITORITEMCONT & " = '" & sItemCont & "' " & _
                  " and " & gsCAT_TCMR_MONITORITEMCONTDESC & " = '" & sItemContDesc & "' " & _
                  " and " & gsCAT_TCMR_REVIEWFLAG & " = 'N' " & _
                  " and " & gsCAT_TCMR_DELETEFLAG & " = 'N' "
        Set colRaws = moProRawSql.QueryDatabase(sSQL)
        If colRaws.Count > 0 Then
            Call TransferLogTrace(DateTime.Now & ": '" & sEventType & "' event for " & sLotId & "/" & sEqId & " already existed. Skip insert.", _
                                              moAppLog, Me.rtfTrace)
        Else '沒有才需要加入
            Call TransferLogTrace(DateTime.Now & ": Add '" & sEventType & "' for " & sLotId & "/" & sEqId, _
                                  moAppLog, Me.rtfTrace)
        
            Call InsertMonitorReceive(sEqId, _
                                      sLotId, _
                                      "", _
                                      "", _
                                      gsMONITOR_ITEM_TESTDATA, _
                                      sItemCont, _
                                      sItemContDesc, _
                                      "1", _
                                      gsCAT_TCMR_PRODTAFLAG)
    
            '單機台整理
            Call Level1Sort(sEqId)
    
            '單機台更新
            Call UpdateMonitorSingleEq(sEqId)
        End If
    End If
    '-------
    'Done
    '-------

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
           Case Else
                typErrInfo.sUserText = typErrInfo.sErrDescription
        End Select
        '---- Start of Select Case Block ----
        On Error GoTo ExitHandler:
        Call TransferLogError(typErrInfo.sUserText, _
                              moAppLog, Me.rtfTrace, msMODULE_ID, sProcID)
        Call HandleError(False, typErrInfo, , moAppLog)
    End If
End Sub
'
'================================================================================
'Name       : GetEqDetailLotList()
'Description: 取得EQ上LotID的詳細資訊
'Author     : Weilun huang, MXIC 2023/06/01
'Parameters : N/A
'================================================================================
Private Sub GetEqDetailLotList(ByVal sEqId As String, _
                               ByRef sSampleLotID As String, _
                               ByRef sMonitorShowLotId As String, _
                               ByRef sCurLotIdList As String, _
                               ByRef bIsVirtualMerge As Boolean, _
                               ByRef sVirtualLotId As String)

    '考慮到重工或中途停測Case, 以PgmList為主, (虛擬併批需轉WaferID), 並抓取PgmList CreateTime之後的Summmary以及WaferEnd資料

    '比對用LotID_WaferID取代


    On Error GoTo ExitHandler:
    Dim sProcID                     As String
    Dim typErrInfo                  As tErrInfo

    Dim sSQL                        As String
    Dim colRaws                     As Collection
    Dim oRaws                       As Collection

    Dim oLot                        As FwLot
    Dim sLotId                      As String
    Dim sGetLotID                   As String
    Dim bIsCPSmallLot               As Boolean
    Dim oEquipment                  As FwEquipment
    
    Dim sEqpLotIDs                  As String
    Dim asEqpLotIDs()               As String
    Dim lIndex                      As Long
    
    Dim colCurLotList               As Collection
    
    

    '-----
    ' Init
    '-----
    sProcID = "GetEqDetailLotList"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog) '

    bIsVirtualMerge = False
    sCurLotIdList = ""
    sSampleLotID = ""
    sMonitorShowLotId = ""
    sVirtualLotId = ""
    
    '--------------------
    ' Condition Checking
    '--------------------

    '-------
    ' Action
    '-------
    Set colCurLotList = New Collection
    Set oEquipment = moFwMDL.EquipmentById(sEqId)
    If oEquipment Is Nothing Then
        Call TransferLogError("Eq: " & sEqId & ", 不存在.", _
                              moAppLog, Me.rtfTrace, msMODULE_ID, sProcID)
        GoTo ExitHandler
    End If
    
    oEquipment.Refresh
    sEqpLotIDs = oEquipment.CustomAttributes(gsEQP_CUSTOMATTR_CUR_LOT_ID)
  
    sEqpLotIDs = Replace(sEqpLotIDs, ";", ",")
    
    sCurLotIdList = sEqpLotIDs

    asEqpLotIDs = Split(sEqpLotIDs, ",")
    
    '=====假設第一個LotID確實是正在測的=====
    '不過也有可能只是掛在機台上, 但是並非為Run
    
    If UBound(asEqpLotIDs) = -1 Then '空值Split導致
'        Call TransferLogError("Eq: " & sEqId & ", 沒有CurLotId資訊.", _
'                              moAppLog, Me.rtfTrace, msMODULE_ID, sProcID)
'        GoTo ExitHandler
    ElseIf UBound(asEqpLotIDs) > -1 Then
        '以LotID為主, IPN則會有重複
        For lIndex = LBound(asEqpLotIDs) To UBound(asEqpLotIDs)
            sGetLotID = asEqpLotIDs(lIndex)
            
            '取得Lot
            Set oLot = Nothing
            Set oLot = moFwWIP.LotById(sGetLotID)
            If Not (oLot Is Nothing) Then
                colCurLotList.Add sGetLotID
'                If (oLot.CustomAttributes(gsLOT_CUSTOMATTR_STATUS) = gsLOTSTATUS_RUN Or oLot.CustomAttributes(gsLOT_CUSTOMATTR_STATUS) = gsLOTSTATUS_ERUN) Then
                    sSampleLotID = sGetLotID '取得代表批, 並且結束搜尋 ---!!注意!!--為虛擬批時並非正在Run的虛擬子批
                    Exit For
'                End If
            End If
        Next
        
        If Trim(sSampleLotID) <> "" Then
            '檢查是否為虛擬併批
            Call CheckMergeType(moProRawSql, _
                                moAppLog, _
                                sSampleLotID, _
                                bIsVirtualMerge, _
                                bIsCPSmallLot, _
                                sVirtualLotId)
                                
            '呈現在監控Level 1的LotID
            If bIsVirtualMerge = True Then
                sMonitorShowLotId = sVirtualLotId '虛擬批
            Else
                sMonitorShowLotId = sSampleLotID '一般批
            End If
        Else
            sSampleLotID = ""
'            Call TransferLogError("Eq: " & sEqId & ", 無法取得CurRunLotId.", _
'                                  moAppLog, Me.rtfTrace, msMODULE_ID, sProcID)
'            GoTo ExitHandler
        End If

    End If
    '-------
    'Done
    '-------

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
           Case Else
                typErrInfo.sUserText = typErrInfo.sErrDescription
        End Select
        '---- Start of Select Case Block ----
        On Error GoTo ExitHandler:
        Call TransferLogError(typErrInfo.sUserText, _
                              moAppLog, Me.rtfTrace, msMODULE_ID, sProcID)
        Call HandleError(False, typErrInfo, , moAppLog)
    End If
End Sub


'================================================================================
'Name       : GetUnTestedTotalWQty()
'Description:
'Author     : Weilun huang, MXIC 2023/06/01
'Parameters : N/A
'================================================================================
Private Function GetUnTestedTotalWQty(ByVal sEqpLotIDs As String, _
                                      ByVal sEqId As String, _
                                      ByVal sStepName As String, _
                                      Optional ByVal sMsgLotID As String = "", _
                                      Optional ByVal sReciveMsg As String = "") As Long
    On Error GoTo ExitHandler
    Dim sProcID As String
    Dim typErrInfo As tErrInfo

    Dim sSQL                        As String
    Dim colRaws                     As Collection
    Dim oRaws                       As Collection
    Dim sTimeHereSince              As String
    
    Dim sLotId                      As String
    Dim asEqpLotIDs()               As String
    Dim lIndex                      As Long

    '-----
    ' Init
    '-----
    sProcID = "GetUnTestedTotalWQty"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog) '

    GetUnTestedTotalWQty = 0
    
    If Trim(sEqpLotIDs) = "" Then '必要項
        GoTo ExitHandler
    End If

    '--------------------
    ' Condition Checking
    '--------------------

    '-------
    ' Action
    '-------
            
    sEqpLotIDs = Replace(sEqpLotIDs, ";", ",")

    asEqpLotIDs = Split(sEqpLotIDs, ",")
            
    For lIndex = LBound(asEqpLotIDs) To UBound(asEqpLotIDs)
        sLotId = asEqpLotIDs(lIndex)
        
        If sLotId = sMsgLotID Then 'WaferEndNotice的Case
            GetUnTestedTotalWQty = GetUnTestedTotalWQty + GetUnTestedWQty(sLotId, sEqId, "", "", "", sReciveMsg)
        Else '一般EqChange的Case
            GetUnTestedTotalWQty = GetUnTestedTotalWQty + GetUnTestedWQty(sLotId, sEqId, "", sStepName, "")
        End If
    Next
  
    '-------
    'Done
    '-------

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
           Case Else
                typErrInfo.sUserText = typErrInfo.sErrDescription
        End Select
        '---- Start of Select Case Block ----
        On Error GoTo ExitHandler:
        Call TransferLogError(typErrInfo.sUserText, _
                              moAppLog, Me.rtfTrace, msMODULE_ID, sProcID)
        Call HandleError(False, typErrInfo, , moAppLog)
    End If
End Function

'================================================================================
'Name       : GetVirtualLotCurRunInfo()
'Description: 取得虛擬併批正在虛擬子批與其WaferID
'Author     : Weilun huang, MXIC 2023/06/06
'Parameters : N/A
'================================================================================
Private Sub GetVirtualLotCurRunInfo(ByVal sEqpLotIDs As String, _
                                    ByVal sVirtualLotId As String, _
                                    ByRef sCurRunSlotNo As String, _
                                    ByRef sCurRunLotId As String, _
                                    ByRef sCurRunWaferID As String)


    On Error GoTo ExitHandler:
    Dim sProcID                     As String
    Dim typErrInfo                  As tErrInfo

    Dim sSQL                        As String
    Dim colRaws                     As Collection
    Dim oRaws                       As Collection

    Dim spdVirtualMapping           As clsSandBoxSpread
    Dim sAllSlotNo                  As String
    Dim asAllSlotNo()               As String
    Dim sFristSlotNo                As String
    Dim vLotID                      As Variant
    Dim vSlotNo                     As Variant
    Dim vWaferID                    As Variant
    Dim lRow                        As Long
    Dim lIndex                      As Long

    '-----
    ' Init
    '-----
    sProcID = "GetVirtualLotCurRunInfo"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog) '

    sCurRunSlotNo = "NA"
    sCurRunLotId = ""
    sCurRunWaferID = ""
    
    '--------------------
    ' Condition Checking
    '--------------------

    '-------
    ' Action
    '-------

    If sVirtualLotId <> "" Then
        Set spdVirtualMapping = New clsSandBoxSpread
        
        With spdVirtualMapping
            .MaxRows = 0
            .MaxCols = miSPDVIRTUALMAPPING_MAX_COLS
    
            Call .SetText(miSPDVIRTUALMAPPING_SLOT_NO, 0, msSPDVIRTUALMAPPING_SLOT_NO)
            Call .SetText(miSPDVIRTUALMAPPING_LOT_ID, 0, msSPDVIRTUALMAPPING_LOT_ID)
            Call .SetText(miSPDVIRTUALMAPPING_WAFER_ID, 0, msSPDVIRTUALMAPPING_WAFER_ID)
        
            sSQL = " select tvmg." & gsCAT_TVMG_SLOTNO & ", " & _
                          " tvmg." & gsCAT_TVMG_LOTID & ", " & _
                          " tvmg." & gsCAT_TVMG_WAFERID & " " & _
                     " from " & gsCAT_TBL_VIRTUAL_MERGE & " tvmg, " & _
                                gsCAT_TBL_LOT_ATTRIBUTE & " tlatt " & _
                    " where tlatt." & gsCAT_TLATT_LOTSTATUS1 & " in  ('" & gsLOTSTATUS_RUN & "','" & gsLOTSTATUS_ERUN & "') " & _
                      " and tvmg." & gsCAT_TVMG_LOTID & " = tlatt." & gsCAT_TLATT_LOTID & " " & _
                      " and tvmg." & gsCAT_TVMG_LOTID & " in ('" & Replace(sEqpLotIDs, ",", "','") & "') " & _
                      " and tvmg." & gsCAT_TVMG_VIRTUALLOTID & " = '" & sVirtualLotId & "' " & _
                      " and tvmg." & gsCAT_TVMG_DELETEFLAG & " = 'N' " & _
                 " order by tvmg." & gsCAT_TVMG_SLOTNO & " "
            Set colRaws = moProRawSql.QueryDatabase(sSQL)
            For lIndex = 1 To colRaws.Count
                .MaxRows = spdVirtualMapping.MaxRows + 1
                .SetText (miSPDVIRTUALMAPPING_SLOT_NO), lIndex, colRaws.Item(lIndex).Item(gsCAT_TVMG_SLOTNO)
                .SetText (miSPDVIRTUALMAPPING_LOT_ID), lIndex, colRaws.Item(lIndex).Item(gsCAT_TVMG_LOTID)
                .SetText (miSPDVIRTUALMAPPING_WAFER_ID), lIndex, colRaws.Item(lIndex).Item(gsCAT_TVMG_WAFERID)
                
                If sCurRunLotId = "" Then
                    '經SlotNo排序, 第一個取得的LotID基本上就是目前在跑的虛擬子批
                    sCurRunLotId = colRaws.Item(lIndex).Item(gsCAT_TVMG_LOTID)
                End If
'               sAllSlotNo = sAllSlotNo & "," & Format(colRaws.Item(lIndex).Item(gsCAT_TVMG_SLOTNO), "00")
            Next
        
            '取得目前正在Run的WaferID
            sCurRunWaferID = GetCurWaferIDByPgmList(sCurRunLotId)
            '再轉回SlotNo
            If sCurRunWaferID <> "NA" And sCurRunWaferID <> "" Then
                For lRow = 1 To .MaxRows
                    .GetText miSPDVIRTUALMAPPING_WAFER_ID, lRow, vWaferID
                    
                    If sCurRunWaferID = CStr(vWaferID) Then
                        .GetText miSPDVIRTUALMAPPING_SLOT_NO, lRow, vSlotNo
                        sCurRunSlotNo = CStr(vSlotNo)
                    End If
                Next
            End If
        End With
    End If
    '-------
    'Done
    '-------

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
           Case Else
                typErrInfo.sUserText = typErrInfo.sErrDescription
        End Select
        '---- Start of Select Case Block ----
        On Error GoTo ExitHandler:
        Call TransferLogError(typErrInfo.sUserText, _
                              moAppLog, Me.rtfTrace, msMODULE_ID, sProcID)
        Call HandleError(False, typErrInfo, , moAppLog)
    End If
End Sub

'================================================================================
'Name       : UpdateAGV_EqStatusUnloadInfo()
'Description: Copy and Modify from modCustomUpdate.UpdateAGV_UnloadInfo
'Author     : Weilun huang, MXIC 2021/11/17 for Project.BE 工業 3.5 Phase 39 - CP AGV導入專案 - 建立 AGV 派工 & Status 監控系統
'Parameters : N/A
'================================================================================
Public Sub UpdateAGV_EqStatusUnloadInfo(ByVal sEqId As String, _
                                        ByVal sLotId As String, _
                                        ByVal sEndTime As String)
    On Error GoTo ExitHandler:
    Dim sProcID                     As String
    Dim typErrInfo                  As tErrInfo

    Dim sSQL                        As String
    Dim colRaws                     As Collection
    Dim oRaws                       As Collection
    Dim colLot                      As Collection
    Dim sWaferSize                  As String
    Dim sWaferQty                   As String
    
    Dim sUnloadLotID                As String
    Dim sUnloadEndTime              As String
    Dim sUnloadPcs                  As String
    
    Dim bIsVirtualMerge             As Boolean
    Dim bIsCPSmallLot               As Boolean
    Dim sVirtualLotId               As String
    Dim sMergeType                  As String
'    Dim colVirtualList              As Collection
    Dim sShowLotId                  As String
    Dim sEqpLotIDs                  As String
    Dim sVirtualWQty                As String

    '-----
    ' Init
    '-----
    sProcID = "UpdateAGV_EqStatusUnloadInfo"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog) '
    
    If sEndTime = "" Then
        sEndTime = "NA"
    End If
    

    '--------------------
    ' Condition Checking
    '--------------------

    '-------
    ' Action
    '-------
    '由於還要取WaferSize, LotId請不要使用虛擬併批母批, 'Comment by Weilun on 20230626 for Project.新增CP監控識別虛擬併批功能以及修改預測時間機制
    '還得轉虛擬併批
    If sEqId <> "" And sLotId <> "" Then
        '檢查是否為虛擬併批
        Call CheckMergeType(moProRawSql, _
                            moAppLog, _
                            sLotId, _
                            bIsVirtualMerge, _
                            bIsCPSmallLot, _
                            sVirtualLotId)
        

        '取得虛擬併批的總WQty
        sVirtualWQty = "0"
        If bIsVirtualMerge = True And sVirtualLotId <> "" Then
            sSQL = "select tli." & gsCAT_TLI_VIRTUALLOTID & ", " & _
                         " sum(tlatt." & gsCAT_TLATT_WAFERQTY & ") as VirtualWQty " & _
                    " from " & gsCAT_TBL_LOT_INFO & " tli, " & _
                               gsCAT_TBL_LOT_ATTRIBUTE & " tlatt " & _
                   " where tlatt." & gsCAT_TLATT_LOTID & " = tli." & gsCAT_TLI_LOT_ID & " " & _
                     " and tli." & gsCAT_TLI_VIRTUALLOTID & " = '" & sVirtualLotId & "'" & _
                     " and tli." & gsCAT_TLI_MERGETYPE & " = '" & gsLOT_MERGETYPE_VIRTUALMERGE & "' " & _
                " group by " & gsCAT_TLI_VIRTUALLOTID & " "
            
            Set colRaws = moProRawSql.QueryDatabase(sSQL)
            If colRaws.Count > 0 Then
                sVirtualWQty = colRaws.Item(1).Item("VirtualWQty")
            End If
        End If
        
        '取得WaferSize
        sWaferSize = ""
        sWaferQty = "0"
        sSQL = "select tlatt." & gsCAT_TLATT_LOTID & ", " & _
                     " tlatt." & gsCAT_TLATT_WAFERQTY & ", " & _
                     " tpb." & gsCAT_TPB_WAFER_SIZE & " " & _
                " from " & gsCAT_TBL_LOT_ATTRIBUTE & " tlatt, " & _
                     gsCAT_TBL_IPN_MASTER & " tim, " & _
                     gsCAT_TBL_PROD_BODY & " tpb " & _
               " where tim." & gsCAT_TIM_PRODBODY & " = tpb." & gsCAT_TPB_PROD_BODY & " " & _
                 " and tlatt." & gsCAT_TLATT_IPN & " = tim." & gsCAT_TIM_IPN & " " & _
                 " and tlatt." & gsCAT_TLATT_LOTID & " = '" & sLotId & "' "
        Set colRaws = moProRawSql.QueryDatabase(sSQL)
        If colRaws.Count > 0 Then
            sWaferSize = colRaws.Item(1).Item(gsCAT_TPB_WAFER_SIZE)
            sWaferQty = colRaws.Item(1).Item(gsCAT_TLATT_WAFERQTY)
        End If
    
        '寫入下貨資料
        If sWaferSize = "12" Then
            If bIsVirtualMerge = True And sVirtualLotId <> "" Then
                sUnloadLotID = sVirtualLotId
                sUnloadPcs = sVirtualWQty
            Else
                sUnloadLotID = sLotId
                sUnloadPcs = sWaferQty
            End If
            'Modify by Weilun on 20230721 for Project.新增CP監控識別虛擬併批功能以及修改預測時間機制 <start>
'            sUnloadEndTime = "' & sEndTime '" '含單引號, 預計測試結束時間
            If sEndTime = "" Then
                sUnloadEndTime = "'NA'"
            Else
                sUnloadEndTime = sEndTime '內含時間格式
            End If
            'Modify by Weilun on 20230721 for Project.新增CP監控識別虛擬併批功能以及修改預測時間機制 <end>
        Else
            sUnloadLotID = "NA"
            sUnloadPcs = "0"
            sUnloadEndTime = "'NA'" '含單引號
        End If
        
        Call TransferLogTrace(DateTime.Now & ": Update AGV Unload information for " & sEqId, _
                      moAppLog, Me.rtfTrace)
        
        '資料更新
        sSQL = " update " & gsCAT_TBL_AGV_EQ_STATUS & " " & _
                  " set " & gsCAT_TAES_LOTID & " = '" & sUnloadLotID & "', " & _
                            gsCAT_TAES_LOTPCS & " = '" & sUnloadPcs & "', " & _
                            gsCAT_TAES_LOTENDTIME & " = " & sUnloadEndTime & ", " & _
                            gsCAT_TAES_UPDATEUSERID & " = 'SYS', " & _
                            gsCAT_TAES_UPDATETIME & "  = to_char(sysdate, 'yyyymmdd hh24miss') || '000' " & _
                " where " & gsCAT_TAES_EQID & " = '" & sEqId & "' " & _
                  " and " & gsCAT_TAES_DELETEFLAG & " = 'N' "
        Call moProRawSql.QueryDatabase(sSQL)

    End If
    


    '-------
    'Done
    '-------

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
           Case Else
                typErrInfo.sUserText = typErrInfo.sErrDescription
        End Select
        '---- Start of Select Case Block ----
        On Error GoTo ExitHandler:
        Call HandleError(False, typErrInfo, , moAppLog, True)

    End If
End Sub

'Add by Samson on 2026/03/04 check disconnection and restart
Private Sub CheckConnection_Timer()
    On Error GoTo ExitHandler:
    Dim sProcID As String
    Dim sSQL As String
    Dim colRS As Collection
    Dim typErrInfo As tErrInfo

    '檢查與資料庫連線狀態
    sSQL = "SELECT * FROM " & gsCAT_TBL_LOT_INFO & " WHERE rownum <= 1"
    Set colRS = moProRawSql.QueryDatabase(sSQL)

ExitHandler:
    ' NOTE 1:
    ' MUST CALL GetErrInfo() here first before another action
    Call GetErrInfo(msMODULE_ID, sProcID, typErrInfo, Erl)
    Call LogProcOut(msMODULE_ID, sProcID, typErrInfo, moAppLog)
    ' <Your cleaning up codes goes here...>
ErrorHandler:
    If typErrInfo.lErrNumber Then
      
        Select Case typErrInfo.lErrNumber
           Case Else
                typErrInfo.sUserText = typErrInfo.sErrDescription
        End Select
        '---- Start of Select Case Block ----
        On Error GoTo ExitHandler:
        Call HandleError(False, typErrInfo, , moAppLog, False)
        Call Restart
    End If
End Sub



'================================================================================
'Name       : Restart()
'Description: Close Form and Restart
'Author     : Samson Chiang, MES 2026/03/04
'Parameters : N/A
'================================================================================
Private Sub Restart()
    
    Dim pid      As String
    
    'ADD by Samson on 20260304 For AutoReconnection
    Call LogTrace(DateTime.Now & " AutoReconnection Start", moAppLog)
    
    pid = Shell("C:\Windows\SysWOW64\FwClient\Exe\CP_ETEST_MonitorCore.exe")
    If pid > 0 Then

        Call CloseHandle(pid)
    Else
        Call Shell("C:\Windows\SysWOW64\FwClient\Exe\CP_ETEST_MonitorCore.exe")

    End If
    
    Unload Me
End Sub

Public Sub CleanDuplicateExeInstances()
'Add by Samson on 20260304 for clean Pid
    Dim objWMIService As Object
    Dim colProcessList As Object
    Dim objProcess As Object
    Dim strExeName As String
    Dim currentPid As Long
    Dim arrProcesses() As Variant
    Dim i As Long, j As Long
    Dim latestStartTime As Date
    Dim pidToKeep As Long
    Dim pid As Long
    Dim startTime As Date
    
    ' 取得目前 EXE 名稱
    strExeName = App.EXEName & ".exe"
    currentPid = GetCurrentProcessId()
    
    ' 連接 WMI
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    
    ' 取得同名 EXE 的所有進程
    Set colProcessList = objWMIService.ExecQuery("SELECT ProcessId, CreationDate, Name FROM Win32_Process WHERE Name='" & strExeName & "'")
    
    ' 暫存進程資料：pid 與啟動時間
    ReDim arrProcesses(1 To colProcessList.Count, 1 To 2) As Variant
    i = 0
    For Each objProcess In colProcessList
        i = i + 1
        pid = CLng(objProcess.ProcessId)
        startTime = WMIDateStringToDate(objProcess.CreationDate)
        
        arrProcesses(i, 1) = pid
        arrProcesses(i, 2) = startTime
    Next
    
    ' 找出最新時間的 PID
    latestStartTime = #1/1/1970#
    pidToKeep = 0
    For i = 1 To UBound(arrProcesses, 1)
        If arrProcesses(i, 2) > latestStartTime Then
            latestStartTime = arrProcesses(i, 2)
            pidToKeep = arrProcesses(i, 1)
        End If
    Next
    
    ' 關閉不是最新的 PID (保留 pidToKeep)
    For i = 1 To UBound(arrProcesses, 1)
        pid = arrProcesses(i, 1)
        If pid <> pidToKeep Then
            TerminateProcessByPID pid
        End If
    Next

End Sub

Private Function GetCurrentProcessId() As Long
'Add by Samson on 20260304 for clean Pid
    ' 使用 API 取得目前程序 PID
    ' VB6 沒有內建函式，要用 Windows API
    ' 透過 WMI 取得也可，但較複雜
    ' 這裡用 WMI 取得 PID 範例:
    Dim objWMIService As Object
    Dim colProcessList As Object
    Dim objProcess As Object
    Dim strExeName As String
    Dim currentExeFullPath As String
    
    strExeName = App.EXEName & ".exe"
    Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
    Set colProcessList = objWMIService.ExecQuery("SELECT ProcessId, ExecutablePath FROM Win32_Process WHERE Name='" & strExeName & "'")
    
    ' 找出與目前執行檔全路徑相同的 PID (因可能有多個同名)
    currentExeFullPath = App.Path & "\" & strExeName
    If Right(currentExeFullPath, 1) = "\" Then currentExeFullPath = Left(currentExeFullPath, Len(currentExeFullPath) - 1)
    
    For Each objProcess In colProcessList
        If StrComp(objProcess.ExecutablePath, currentExeFullPath, vbTextCompare) = 0 Then
            GetCurrentProcessId = objProcess.ProcessId
            Exit Function
        End If
    Next
    GetCurrentProcessId = 0
End Function

Private Sub TerminateProcessByPID(ByVal pid As Long)
'Add by Samson on 20260304 for clean Pid
    Dim hProcess As Long
    hProcess = OpenProcess(PROCESS_TERMINATE, 0, pid)
    If hProcess <> 0 Then
        TerminateProcess hProcess, 0
        CloseHandle hProcess
    End If
End Sub

Private Function WMIDateStringToDate(ByVal dtmIn As String) As Date
'Add by Samson on 20260304 for clean Pid
    ' WMI 的 CreationDate  格式: yyyymmddHHMMSS.mmmmmmsUUU
    ' 解析成 VB6 Date 型態
    WMIDateStringToDate = CDate(Mid(dtmIn, 5, 2) & "/" & Mid(dtmIn, 7, 2) & "/" & Left(dtmIn, 4) & " " & _
                              Mid(dtmIn, 9, 2) & ":" & Mid(dtmIn, 11, 2) & ":" & Mid(dtmIn, 13, 2))
End Function


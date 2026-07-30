VERSION 5.00
Begin VB.Form frmPrintFtFVISetupForm 
   Caption         =   "PrintFtFVISetupForm"
   ClientHeight    =   1635
   ClientLeft      =   3315
   ClientTop       =   4320
   ClientWidth     =   8700
   LinkTopic       =   "Form1"
   ScaleHeight     =   1635
   ScaleWidth      =   8700
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   2775
      TabIndex        =   8
      Top             =   1200
      Width           =   1440
   End
   Begin VB.TextBox txtLotID 
      Height          =   315
      Left            =   825
      TabIndex        =   2
      Top             =   525
      Width           =   1815
   End
   Begin VB.Frame fraLotCriteria 
      Caption         =   "Lot Criteria"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   990
      Left            =   75
      TabIndex        =   1
      Top             =   75
      Width           =   8490
      Begin VB.ComboBox cboSetupForm 
         Height          =   315
         Left            =   6375
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   450
         Width           =   1815
      End
      Begin VB.TextBox txtEqId 
         Height          =   315
         Left            =   3300
         TabIndex        =   4
         Top             =   450
         Width           =   1815
      End
      Begin VB.Label lblSetupForm 
         Caption         =   "Form Type:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5400
         TabIndex        =   6
         Top             =   450
         Width           =   990
      End
      Begin VB.Label Label1 
         Caption         =   "EqId:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2850
         TabIndex        =   5
         Top             =   450
         Width           =   615
      End
      Begin VB.Label lblLotId 
         Caption         =   "LotId:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   150
         TabIndex        =   3
         Top             =   450
         Width           =   840
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   4575
      TabIndex        =   0
      Top             =   1200
      Width           =   1440
   End
   Begin VB.Label lblStopInfo 
      Caption         =   "停測資訊"
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Top             =   1200
      Visible         =   0   'False
      Width           =   1935
   End
End
Attribute VB_Name = "frmPrintFtFVISetupForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const msMODULE_ID As String = "frmPrintFtFVISetupForm"
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
Private msOprId     As String

'Add by Tony Start on 2014/03/10 for HW SETUP RECIPE COMPARE project.
Private msPackageName As String
Private msPinCount    As String
Private msBodySize    As String

Private Const miCOMBO_RECIPE_ROW   As Integer = 20
Private Const miFTFVI_RECIPE_ROW   As Integer = 20
Private Const miLASER_RECIPE_ROW   As Integer = 20
Private Const mi_RECIPE_MAX_ROW    As Integer = 40 'SI要求超過40不要印出

'Add by Tony End on 2014/03/10 for HW SETUP RECIPE COMPARE project.

Private Const msFilePath        As String = "\system32\FwClient\Others\SetupForm.XLT"


Private miResult        As Integer
Private msDocNo         As String
Private msFormType      As String

Public Property Let LotID(sLotID As String)
    On Error Resume Next
    msLotId = sLotID
End Property
Public Property Let EqpId(sEqpID As String)
    On Error Resume Next
    msEqpId = sEqpID
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
'**************************************************
'**************************************************
    On Error Resume Next
    Result = miResult

End Property

Private Sub cmdCancel_Click()
    On Error Resume Next
    miResult = vbCancel
    Me.Hide
End Sub
'================================================================================
' Sub: cmdOk_Click()
'--------------------------------------------------------------------------------
' Description:  <M200804014>
'--------------------------------------------------------------------------------
' Author:       Tony Chang , MXIC 2008/04/09
'================================================================================
Private Sub cmdOK_Click()
On Error GoTo ExitHandler:
Dim sProcID     As String
Dim typErrInfo  As tErrInfo
Dim sFilePath   As String
Dim sFile       As String
Dim oLot        As FwLot
Dim oEqp        As FwEquipment

'Add by Sam start on 20170321 for Project 停復測自動化
Dim sSQL As String
Dim colRS As Collection
Dim lIdx As Long
Dim sERunTicNo As String
Dim sMsg As String
Dim sAccessary As String
Dim sPgID As String
Dim sPgName As String
Dim sPgMode As String
Dim sContactBoard As String
Dim sSubsystem As String
Dim sTemp As String
Dim sStopTicNo As String
Dim sSpecifyEQ As String
Dim sSpecifyEqID As String
Dim sStopTicNoList As String
Dim sStopInfo As String

'Add by Sam end on 20170321 for Project 停復測自動化

'----
' Init
'----
    sProcID = "cmdOk_Click"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog)
    Screen.MousePointer = vbHourglass
'----
' Condition Checking
'----
    

    If (moFwWIP Is Nothing) Then
        Call RaiseError(glERR_INVALIDOBJECT, _
                        FormatErrorText(gsETX_INVALIDOBJECT, "FwWIP"))
    End If
    
    If Me.txtLotID.Text = "" Then
        UtShowMsgBox "請輸入LOTID!" & vbCrLf & "Please keyin LotId!"
        GoTo ExitHandler
    End If
    
    If Me.txtEqId.Text = "" Then
        UtShowMsgBox "請輸入機台編號!" & vbCrLf & "Please keyin EqId!"
        GoTo ExitHandler
    End If
    
    If Me.cboSetupForm.Text = "" Then
        UtShowMsgBox "請選擇FormType!" & vbCrLf & "Please choose Form Type!"
        GoTo ExitHandler
    End If
    
    sFilePath = Environ$("windir") & msFilePath
    sFile = Dir(sFilePath)
    
    If sFile = "" Then
        UtShowMsgBox "SetupForm Excel template not found!" & vbCrLf & _
                     "找不到SetupForm Excel範本!"
        GoTo ExitHandler
    End If

    'Add by Sam start on 20170321 for Project 停復測自動化
    '與user確認過僅做機台停測的判斷
    Set oLot = FwuRetrieveLot(moFwWIP, txtLotID.Text, moAppLog)
    Set oEqp = FwuRetrieveEqp(moFwMDL, UCase(Me.txtEqId.Text), moAppLog)
    lblStopInfo.Caption = ""
    If (Not oLot Is Nothing) And (Not oEqp Is Nothing) Then
        msFormType = Me.cboSetupForm.Text
        sSQL = "select a." & gsCAT_TLI_ERUNTICNO & " from " & gsCAT_TBL_LOT_INFO & " a " & _
               " where a." & gsCAT_TLI_LOT_ID & " = '" & oLot.Id & "' "
        Set colRS = moProRawSql.QueryDatabase(sSQL)
        If colRS.Count > 0 Then
            sERunTicNo = colRS.Item(1).Item(gsCAT_TLI_ERUNTICNO)
        End If
        
        If msFormType = "COMBO" Then
            Call GetComboSpec(oLot, oEqp, sContactBoard, sPgID, sPgName, sPgMode, sSubsystem, sTemp, sStopTicNo, sSpecifyEQ, sSpecifyEqID)
        End If
        
        If NotStopTest(oLot.Id, oLot.CustomAttributes(gsLOT_CUSTOMATTR_IPN), oLot.CustomAttributes(gsLOT_CUSTOMATTR_PROD_GROUP), gsCAT_STOP_TEST_ALL, False, _
                                                             oLot.CurrentStep.Steps.Item(1).Description, oEqp.Id, oEqp.CustomAttributes(gsEQP_CUSTOMATTR_EQ_TYPE2) _
                                                             , sContactBoard, sStopTicNo, _
                                                             sSpecifyEQ, sSpecifyEqID, sERunTicNo, sPgID, sPgName, sPgMode, _
                                                             sTemp, moFwWIP, moAppLog, moProRawSql, "", sSubsystem, "", sMsg) = False Then
            Call UtShowMsgBox("停測中，停測資訊:" & vbNewLine & sMsg)
            GoTo ExitHandler
        End If
        
        If msFormType = "COMBO" Then
            Call GetStopInfoPByAcc(oLot.Id, oLot.CustomAttributes(gsLOT_CUSTOMATTR_PROD_GROUP), _
                                   oEqp.CustomAttributes(gsEQP_CUSTOMATTR_EQ_TYPE2), _
                                   oEqp.Id, oLot.CurrentStep.Steps.Item(1).Description, _
                                   sPgID, sPgName, sPgMode, "", sContactBoard, sStopTicNoList)
                                 
            Call GetStopInfoNPByAcc(oLot.CustomAttributes(gsLOT_CUSTOMATTR_IPN), oEqp.CustomAttributes(gsEQP_CUSTOMATTR_EQ_TYPE2), _
                                    oEqp.Id, sContactBoard, sTemp, sMsg, moFwWIP, moAppLog, moProRawSql)
            
            If sStopTicNoList <> "" Then
                sStopInfo = sStopTicNoList
            End If
            
            If sMsg <> "" Then
                sStopInfo = sStopTicNoList & "," & sMsg
            End If
            
            If sStopInfo <> "" Then
                lblStopInfo.Caption = sStopInfo
                Call UtShowMsgBox("請確認停測資訊:" & vbNewLine & sStopInfo, vbInformation)
            Else
                lblStopInfo.Caption = "NA"
            End If
        End If
        
    End If
    'Add by Sam en on 20170321 for Project 停復測自動化
    
'----
' Action
'----
    
    Dim appXL       As Object
    Dim iPosTitle   As Integer
    
    Set oLot = FwuRetrieveLot(moFwWIP, txtLotID.Text, moAppLog)
    Set oEqp = FwuRetrieveEqp(moFwMDL, UCase(Me.txtEqId.Text), moAppLog)
    msFormType = Me.cboSetupForm.Text
        
    Set appXL = CreateObject("Excel.Application")
    If appXL Is Nothing Then
        UtShowMsgBox "Excel not found"
        GoTo ExitHandler
    End If
    
    appXL.cursor = 2
    appXL.DisplayAlerts = False
    appXL.Workbooks.Add sFilePath
    appXL.ActiveWindow.WindowState = vbMaximized

    'ADD BY Sam start on 20100611 for ReqNo:JC201000123
    If msFormType = "LASER" Then
        iPosTitle = 2
        Call LaserSetupForm(appXL, oLot, iPosTitle)
    'Add by Sam Start on 20130328 for Project FVI Setup Recipe
    ElseIf msFormType = "COMBO" Then
        iPosTitle = 1
        Call ComboSetupForm(appXL, oLot, oEqp, iPosTitle)
    'Add by Sam End on 20130328 for Project FVI Setup Recipe
    Else
    'ADD BY Sam end on 20100611 for ReqNo:JC201000123
        iPosTitle = 1
    
        Call FtFVISetupForm(appXL, oLot, iPosTitle)
    End If
    appXL.Sheets("WS").Select
    appXL.ActiveWindow.SelectedSheets(1).Visible = False
    appXL.Sheets("FT").Select
    appXL.ActiveWindow.SelectedSheets(1).Visible = False
    appXL.Sheets("SCANNER").Select
    appXL.ActiveWindow.SelectedSheets(1).Visible = False
    appXL.Sheets("T&R").Select
    appXL.ActiveWindow.SelectedSheets(1).Visible = False
    
    'Add by Sam start on 20100611 for ReqNo:JC201000123
    appXL.Sheets(gsRESOURCE_COMBO).Select
    appXL.ActiveWindow.SelectedSheets(1).Visible = False
    appXL.Sheets("LASER").Select
    appXL.ActiveWindow.SelectedSheets(1).Visible = False
    'Add by Sam end on 20100611 for ReqNo:JC201000123
    
    'Add by Sam Start on 20130328 for Project FVI Setup Recipe
    appXL.Sheets("AT3-300AL").Select
    appXL.ActiveWindow.SelectedSheets(1).Visible = False
    appXL.Sheets("FT-940").Select
    appXL.ActiveWindow.SelectedSheets(1).Visible = False
    'Add by Sam END on 20130328 for Project FVI Setup Recipe
    
    appXL.Sheets(1).Select

    appXL.range("A1").Select
    appXL.cursor = 1
    appXL.DisplayAlerts = True
    
    '*****
    'Save
    '*****
'    ChDir "C:\FASTech"
'    appXL.ActiveWorkbook.SaveAs "C:\FASTech\" & msFormType & Format(Now, "YYYYMMDDHHMMSS") & ".xls", _
'        , "", "", False, False

    appXL.Visible = True
    appXL.ActiveWindow.SelectedSheets(1).PrintOut Copies:=1, Collate:=True
    Set appXL = Nothing
        
    miResult = vbOK
    Me.Hide
    
    
'----
' Done
'----

ExitHandler:
    ' NOTE 1:
    ' MUST CALL GetErrInfo() here first before another action
    Call GetErrInfo(msMODULE_ID, sProcID, typErrInfo, Erl)
    Call LogProcOut(msMODULE_ID, sProcID, typErrInfo, moAppLog)
    Screen.MousePointer = vbDefault
    ' Cleaning up
    
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
        Screen.MousePointer = vbDefault
        cmdOK.Enabled = True
        miResult = vbCancel
        On Error GoTo ExitHandler:
        Call HandleError(False, typErrInfo, , moAppLog, True)
    End If
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
' Description:  <M200804014>
'--------------------------------------------------------------------------------
' Author:       Tony Chang , MXIC 2008/04/10
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
    ' <Put your condition checking codes here>...

'----
' Action
'----
    ' <Put your Action codes here>...
         
    If msLotId <> "" Then
        Me.txtLotID.Text = msLotId
    End If
    
    If msEqpId <> "" Then
        Me.txtEqId.Text = msEqpId
    End If
    
    Me.cboSetupForm.AddItem gsRESOURCE_SCANNER
    Me.cboSetupForm.AddItem "T&R"
    
    'Add by Sam start on 20100611 for ReqNo:JC201000123
    Me.cboSetupForm.AddItem gsRESOURCE_COMBO
    Me.cboSetupForm.AddItem "LASER"
    'Add by Sam End on 20100611 for ReqNo:JC201000123
    
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
        Init = False
        On Error GoTo ExitHandler:
        Call HandleError(True, typErrInfo, , moAppLog)
    End If
End Function
'================================================================================
' Sub: FtFVISetupForm()
'--------------------------------------------------------------------------------
' Description:  <M200804014>
'--------------------------------------------------------------------------------
' Author:       Tony Chang , MXIC 2008/04/09
'================================================================================
Private Sub FtFVISetupForm(ByRef appXL As Object, ByRef oLot As FwLot, _
                           ByRef iPosTitle As Integer)
On Error GoTo ExitHandler:
Dim sProcID                 As String
Dim typErrInfo              As tErrInfo
Dim sSQL                    As String
Dim oRS1                    As Collection
Dim oRs2                    As Collection
Dim iPosBasicInfo           As Integer
Dim iPosLendAcc             As Integer
Dim iPosHWSetupQualityChk   As Integer
Dim sProdgroup              As String

Dim sPgm                    As String 'add by Sam start on 20081014 for ReqNO:M200809045
Dim sEqID                   As String 'add by Sam start on 20081014 for ReqNO:M200809045
Dim bSpecial_PGM            As Boolean 'add by Sam start on 20081014 for ReqNO:M200809045
Dim iIdx As Integer

'add by Sam start on 20100623 for ReqNO:JC201000123
Dim sERunTicNo              As String
Dim sFollowProd As String
Dim sFollowOI As String
Dim oEqp As FwEquipment
'add by Sam END on 20100623 for ReqNO:JC201000123

'----
' Init
'----
    sProcID = "FtFVISetupForm"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog)

    Set oEqp = moFwMDL.EquipmentById(msEqpId) 'add by Sam on 20100623 for ReqNO:JC201000123

'----
' Condition Checking
'----
    ' <Put your condition checking codes here>...

'----
' Action
'----
    ' <Put your Action codes here>...
    iPosBasicInfo = 6
    iPosLendAcc = 13
    iPosHWSetupQualityChk = 27
    'Modify by Sam start on 20081029 for ReqNo:M200810038 22 -> 27
'    If UCase(msFormType) = "SCANNER" Then iPosHWSetupQualityChk = 22
    If UCase(msFormType) = "SCANNER" Then iPosHWSetupQualityChk = 27
    'Modify by Sam end on 20081029 for ReqNo:M200810038
    
    With appXL
    
        '*****
        'COPY template
        '*****
        
        .Sheets(msFormType).Select
        .Cells.Select
        .selection.NumberFormatLocal = "@"
        .selection.Copy
        .Sheets(1).Select
        .ActiveSheet.Paste
        .Application.CutCopyMode = False
        
        '*****
        'Title部分
        '*****
        msOprId = moFwOPR.ActiveUser.UserName
        .range("D" & CStr(iPosTitle + 2)).Select
        .selection.horizontalalignment = -4131
        .activeCell.FormulaR1C1 = msOprId
        
        sSQL = "select shiftcode from tbl_mxic_emp where empno='" & msOprId & "'"
        Set oRS1 = moProRawSql.QueryDatabase(sSQL)
        If oRS1.Count > 0 Then
            .range("I" & CStr(iPosTitle + 2)).Select
            .selection.horizontalalignment = -4131
            .activeCell.FormulaR1C1 = oRS1.Item(1).Item(1)
        End If
        Set oRS1 = Nothing
        
        .range("S" & CStr(iPosTitle + 2)).Select
        .selection.horizontalalignment = -4131
        .activeCell.FormulaR1C1 = Format(Now, "YYYY/MM/DD HH:MM:SS")
        
        '*****
        'BasicInfo部分
        '*****
        
        .range("F" & CStr(iPosBasicInfo)).Select
        .selection.horizontalalignment = -4131
        .activeCell.FormulaR1C1 = oLot.Id
        
        
        .range("F" & CStr(iPosBasicInfo + 1)).Select
        .selection.horizontalalignment = -4131
        .activeCell.FormulaR1C1 = oLot.CustomAttributes(gsLOT_CUSTOMATTR_IPN)
        
        
        'Add by Sam start on 20100611 for ReqNo:JC201000123
        If UCase(msFormType) = gsRESOURCE_COMBO Then
            .range("F" & CStr(iPosBasicInfo + 2)).Select
            .selection.horizontalalignment = -4131
            .activeCell.FormulaR1C1 = oLot.CustomAttributes(gsLOT_CUSTOMATTR_STEPNAME)
        End If
        'Add by Sam End on 20100611 for ReqNo:JC201000123
        
        'Add Packagename/BodySize/PinCount by Tony on 2014/03/10 for HW SETUP RECIPE COMPARE project.
        sSQL = " select " & gsCAT_TIM_PACKAGE_NAME & "," & gsCAT_TIM_PROD_GROUP & "," & gsCAT_TIM_PACKAGE_NAME & "," & gsCAT_TIM_PIN_COUNT & "," & gsCAT_TIM_BODY_SIZE & _
               " from " & gsCAT_TBL_IPN_MASTER & _
               " where " & gsCAT_TIM_IPN & "='" & oLot.CustomAttributes(gsLOT_CUSTOMATTR_IPN) & "'"
        Set oRS1 = moProRawSql.QueryDatabase(sSQL)
        If oRS1.Count > 0 Then
            'Add by Sam start on 20100611 for ReqNo:JC201000123
            If UCase(msFormType) = gsRESOURCE_COMBO Then
                .range("F" & CStr(iPosBasicInfo + 3)).Select
            Else
            'Add by Sam End on 20100611 for ReqNo:JC201000123
                .range("F" & CStr(iPosBasicInfo + 2)).Select
            End If
            .selection.horizontalalignment = -4131
            .activeCell.FormulaR1C1 = oRS1.Item(1).Item(gsCAT_TIM_PACKAGE_NAME)
            sProdgroup = oRS1.Item(1).Item(gsCAT_TIM_PROD_GROUP)
            
            'Add by Tony Start on 2014/03/10 for HW SETUP RECIPE COMPARE project.
            '(先存放值,以便後續使用)
            msPackageName = oRS1.Item(1).Item("packagename")
            msPinCount = oRS1.Item(1).Item("pincount")
            msBodySize = oRS1.Item(1).Item("bodysize")
            'Add by Tony End on 2014/03/10 for HW SETUP RECIPE COMPARE project.
            
        End If
        
        'Marked by Tony Start on 20091207 for FtDailyScheduleAlarmProject.
        'add by Sam start on 20081014 for ReqNO:M200809045
        'Modify A1.customerno --> A1.value by Tony on 20091207 for FtDailyScheduleAlarmProject.
        
'        bSpecial_PGM = False
'            sSql = " SELECT DISTINCT D." & gsCAT_TSCS_PGM & ", D." & gsCAT_TSCI_EQ_ID & " " & _
'                           "from " & gsCAT_TBL_LOT_ATTRIBUTE & " A, " & _
'                                     gsCAT_TBL_IPN_MASTER & " B, " & _
'                                     gsCAT_TBL_PRM_BE_SPEC & " C, " & _
'                                     "( SELECT A1." & gsCAT_TSCS_PGM & ", " & _
'                                              "B1." & gsCAT_TSCI_EQ_ID & ", " & _
'                                              "A1." & gsCAT_TSCS_VALUE & ", " & _
'                                              "A1." & gsCAT_TSCS_CARRIERTYPE & ", " & _
'                                              "A1." & gsCAT_TSCS_PKGCODE & ", " & _
'                                              "A1." & gsCAT_TSCS_PINCONT & ", " & _
'                                              "A1." & gsCAT_TSCS_BODYSIZE & ", " & _
'                                              "A1." & gsCAT_TSCS_BRAND & " " & _
'                                       "from " & gsCAT_TBL_SPECIAL_CUSTOMER_SCANPGM & " A1 ," & _
'                                                 gsCAT_TBL_SCANNER_CONFI & " B1 " & _
'                                       "where A1." & gsCAT_TSCS_PGM & " =B1." & gsCAT_TSCI_PGM & "(+) " _
'                                                & " and A1." & gsCAT_TSCS_CARRIERTYPE & " =B1." & gsCAT_TSCI_CARRIER_TYPE & "(+)" _
'                                                & " and A1." & gsCAT_TSCS_PKGCODE & " =B1." & gsCAT_TSCI_PKG_CODE & "(+)" _
'                                                & " and A1." & gsCAT_TSCS_PINCONT & " =B1." & gsCAT_TSCI_PIN_CONT & "(+)" _
'                                                & " and A1." & gsCAT_TSCS_BODYSIZE & " =B1." & gsCAT_TSCI_BODY_SIZE & "(+)" _
'                                                & " and A1." & gsCAT_TSCS_BRAND & " =B1." & gsCAT_TSCI_BRAND & "(+)) " & " D "
'
'            sSql = sSql & " where A." & gsCAT_TLA_LOTID & "='" & oLot.Id & "'" & _
'                                    " and C." & gsCAT_TPBS_DELETE_FLAG & " <> 'Y'" & _
'                                    " and B." & gsCAT_TIM_DELETE_FLAG & " <> 'Y'" & _
'                                    " and C." & gsCAT_TPBS_DEFAULTS & " = 'Y'" & _
'                                    " and A." & gsCAT_TLATT_IPN & " =B." & gsCAT_TIM_IPN & _
'                                    " and B." & gsCAT_TIM_IPN & " =C." & gsCAT_TPBS_IPN & _
'                                    " and C." & gsCAT_TPBS_CUSTOMER_NO & " =D." & gsCAT_TSCS_VALUE & _
'                                    " and B." & gsCAT_TIM_CARRIER_TYPE & " =D." & gsCAT_TSCS_CARRIERTYPE & _
'                                    " and B." & gsCAT_TIM_PACKAGE_CODE & " =D." & gsCAT_TSCS_PKGCODE & _
'                                    " and B." & gsCAT_TIM_PIN_COUNT & " =D." & gsCAT_TSCS_PINCONT & _
'                                    " and B." & gsCAT_TIM_BODY_SIZE & " =D." & gsCAT_TSCS_BODYSIZE & _
'                                    " and B." & gsCAT_TIM_BRAND & " =D." & gsCAT_TSCS_BRAND & ""
'
'
'            Set oRS2 = moProRawSql.QueryDatabase(sSql)
'            If oRS2.Count > 0 Then      '有客戶指定
'                sEqID = ""
'                For iIdx = 1 To oRS2.Count  '尋找可執行的EQID
'                    If UCase(oRS2.Item(iIdx).Item(gsCAT_TSCI_EQ_ID)) = UCase(Me.txtEqId.Text) Then
'                        sPgm = oRS2.Item(iIdx).Item(gsCAT_TSCS_PGM)
'                        sEqID = oRS2.Item(iIdx).Item(gsCAT_TSCI_EQ_ID) '將客戶指定之EQID帶入
'                    End If
'                    '無可執行的EQID時，將EQID設為空值是為了不要帶出SCANNER_CONFI的其他參數
'                Next
'                bSpecial_PGM = True
'            End If
'
'            If bSpecial_PGM = False Then    '無客戶指定時將EQID設為使用者輸入之EQID
'                sEqID = Me.txtEqId.Text
'            End If
            
        'add by Sam end on 20081014 for ReqNO:M200809045
        'Marked by Tony End on 20091207 for FtDailyScheduleAlarmProject.
        
        .range("S" & CStr(iPosBasicInfo)).Select
        .activeCell.FormulaR1C1 = oLot.CustomAttributes(gsLOT_CUSTOMATTR_LOT_OWNER)

        .range("S" & CStr(iPosBasicInfo + 2)).Select
        .activeCell.FormulaR1C1 = UCase(Me.txtEqId.Text)

        sSQL = " select " & gsCAT_TLI_ERUNTICNO & "," & gsCAT_TLI_SAPRWNO & _
               " from " & gsCAT_TBL_LOT_INFO & _
               " where lotid='" & oLot.Id & "'"
        Set oRs2 = moProRawSql.QueryDatabase(sSQL)
        If oRs2.Count > 0 Then
            .range("S" & CStr(iPosBasicInfo + 1)).Select
            If oRs2.Item(1).Item("erunticno") = "" Then
                .activeCell.FormulaR1C1 = oRs2.Item(1).Item("saprwno")
            ElseIf oRs2.Item(1).Item("saprwno") = "" Then
                .activeCell.FormulaR1C1 = oRs2.Item(1).Item("erunticno")
            Else
                .activeCell.FormulaR1C1 = oRs2.Item(1).Item("erunticno") & ", " & oRs2.Item(1).Item("saprwno")
            End If
             sERunTicNo = oRs2.Item(1).Item("erunticno")    'Add by Sam on 20100623 for ReqNo:JC201000123
        End If
                        
        If Trim(sERunTicNo) <> "" Then
            sSQL = " SELECT A." & gsCAT_TER_FOLLOW_OI & " ,A." & gsCAT_TER_FOLLOW_PRODUCT & " " & _
                   " FROM " & gsCAT_TBL_ERUN_REQ & " A " & _
                   " WHERE A." & gsCAT_TER_TICKET_NO & " ='" & sERunTicNo & "' " & _
                   " AND A." & gsCAT_TER_DELETE_FLAG & " ='N' "
                   
            Set oRs2 = moProRawSql.QueryDatabase(sSQL)
            If oRs2.Count > 0 Then
               sFollowOI = oRs2.Item(1).Item(gsCAT_TER_FOLLOW_OI)
               sFollowProd = oRs2.Item(1).Item(gsCAT_TER_FOLLOW_PRODUCT)
            End If
        End If
        
        'Add by Sam start on 20100611 for ReqNo:JC201000123,COMBO的PGNAME取SPEC的
        If msFormType = gsRESOURCE_COMBO Then
            sSQL = " SELECT B." & gsCAT_PSS_DOC_NO & " , " & _
                   " B." & gsCAT_PSS_COMMENTS & " AS stepcomments , " & _
                   " C." & gsCAT_PSES_COMMENTS & "  as eqcomments , " & _
                   " C." & gsCAT_PSES_PG_NAME & " " & _
                   " FROM " & gsCAT_TBL_PROD_STEP_SPEC & " B ," & _
                   gsCAT_TBL_PROD_STEP_EQ_SPEC & " C " & _
                   " WHERE B." & gsCAT_PSS_PROD_GROUP & "=C." & gsCAT_PSES_PROD_GROUP & "(+) " & _
                   " AND B." & gsCAT_PSS_PATH & "=C." & gsCAT_PSES_PATH & "(+) " & _
                   " AND B." & gsCAT_PSS_STEP_NO & "=C." & gsCAT_PSES_STEP_NO & "(+) " & _
                   " AND B." & gsCAT_PSS_DOC_STATUS & "=C." & gsCAT_PSES_DOC_STATUS & "(+) " & _
                   " AND B." & gsCAT_PSS_DOC_STATUS & "='" & gsDOCSTATUS_ACTIVE & "' " & _
                   " AND B." & gsCAT_PSS_STEP_NO & "='" & oLot.CurrentStep.Steps.Item(1).Id & "' " & _
                   " AND B." & gsCAT_PSS_PROD_GROUP & "='" & sProdgroup & "' " & _
                   " AND B." & gsCAT_PSS_PATH & "='" & oLot.CustomAttributes(gsLOT_CUSTOMATTR_ROUTE) & "' "
            Set oRS1 = moProRawSql.QueryDatabase(sSQL)
            If oRS1.Count > 0 Then
                'TECNNO
                .range("F" & CStr(iPosBasicInfo + 5)).Select
                .activeCell.FormulaR1C1 = oRS1.Item(1).Item(gsCAT_PSS_DOC_NO)
                
                'PgName
                .range("F" & CStr(iPosBasicInfo + 4)).Select
                .activeCell.FormulaR1C1 = oRS1.Item(1).Item(gsCAT_PSES_PG_NAME)
                'EQCOMMENTS
                .range("F" & CStr(iPosBasicInfo + 9)).Select
                .activeCell.FormulaR1C1 = oRS1.Item(1).Item("eqcomments")
                'STEPCOMMENT
                .range("F" & CStr(iPosBasicInfo + 10)).Select
                .activeCell.FormulaR1C1 = oRS1.Item(1).Item("stepcomments")
              
            End If
            
            If sFollowOI = "N" And sFollowProd = "N" Then
               sSQL = " SELECT A." & gsCAT_TERE_PGNAME & "  " & _
                   " FROM " & gsCAT_TBL_ERUN_RECIPE & " A " & _
                   " WHERE A." & gsCAT_TERE_DOCNO & " ='" & sERunTicNo & "' " & _
                   " AND A." & gsCAT_TERE_PATH & " ='" & oLot.CustomAttributes(gsLOT_CUSTOMATTR_ROUTE) & "' " & _
                   " AND A." & gsCAT_TERE_STEPNO & " ='" & oLot.CurrentStep.Steps.Item(1).Id & "' " & _
                   " AND A." & gsCAT_TERE_EQTYPE2 & " ='" & oEqp.CustomAttributes(gsEQP_CUSTOMATTR_EQ_TYPE2) & "' " & _
                   " AND A." & gsCAT_TERE_DELETE_FLAG & " ='N' "
                Set oRs2 = moProRawSql.QueryDatabase(sSQL)
                If oRs2.Count > 0 Then
                    'PgName
                    .range("F" & CStr(iPosBasicInfo + 4)).Select
                    .activeCell.FormulaR1C1 = oRs2.Item(1).Item(gsCAT_TERE_PGNAME)
                End If
                'EQCOMMENTS
                .range("F" & CStr(iPosBasicInfo + 9)).Select
                .activeCell.FormulaR1C1 = ""
                'STEPCOMMENT
                .range("F" & CStr(iPosBasicInfo + 10)).Select
                .activeCell.FormulaR1C1 = ""

            End If
            
            
        Else
        'Add by Sam end on 20100611 for ReqNo:JC201000123,COMBO的PGNAME取SPEC的
            'Modified by Jack on 2016/02/02 for 2016 MES Phase-1 特殊客戶Scan需求攔截機制 專案 <Start>
            'Moved by Tony Start on 20100209 for Req.JC201000024
            'Call GetSpecialPGM(Me.txtEqId.Text, sPGM)
            'Moved by Tony End on 20100209 for Req.JC201000024
            sPgm = getRecipeSpecName1(Me.txtEqId.Text, oLot.Id, oLot.CustomAttributes(gsLOT_CUSTOMATTR_IPN), _
                            msPackageName, msPinCount, msBodySize, _
                            oRS1, moProRawSql, moAppLog, _
                            moCwMbx)
            'Modified by Jack on 2016/02/02 for 2016 MES Phase-1 特殊客戶Scan需求攔截機制 專案 <End>
            
            '<comment by Jack on 2016/02/02 for 2016 MES Phase-1 特殊客戶Scan需求攔截機制 專案>
            ' sSQL 在 以下 oRS1 部分 會用到, 不要誤mark.
            sSQL = " select C." & gsCAT_TSCI_PGM & ",C." & _
                    gsCAT_TSCI_VACUUMCUPZ1 & ",C." & _
                    gsCAT_TSCI_VACUUMCUPZ2 & ",C." & _
                    gsCAT_TSCI_Z1PITCH & ",C." & _
                    gsCAT_TSCI_LENS & ",C." & _
                    gsCAT_TSCI_SEALHEAD & ",C." & _
                    gsCAT_TSCI_SEALPITCH & ",C." & _
                    gsCAT_TSCI_SEALTEMP & ",C." & _
                    gsCAT_TSCI_SEALDWELL & ",C." & _
                    gsCAT_TSCI_SEALPRESSURE & ",C." & _
                    gsCAT_TSCI_LEADCO_TWEEZE & ",C." & _
                    gsCAT_TSCI_LEADPITCH_SPACING & ",C." & _
                    gsCAT_TSCI_LEADSKEW & ",C." & _
                    gsCAT_TSCI_LEADLD & ",C." & _
                    gsCAT_TSCI_LEADST & ",C." & _
                    gsCAT_TSCI_LEADTD & ",C." & _
                    gsCAT_TSCI_LEADBURR & ",C." & _
                    gsCAT_TSCI_MARKXO & ",C." & _
                    gsCAT_TSCI_MARKYO & ",C." & _
                    gsCAT_TSCI_MARKRO & ",C." & _
                    gsCAT_TSCI_MARKCO & ",C."
            sSQL = sSQL & gsCAT_TSCI_MARKOP & ",C." & _
                    gsCAT_TSCI_MARKUP & ",C." & _
                    gsCAT_TSCI_MARKBL & ",C." & _
                    gsCAT_TSCI_BALLCO & ",C." & _
                    gsCAT_TSCI_BALLPITCH & ",C." & _
                    gsCAT_TSCI_BALLXO & ",C." & _
                    gsCAT_TSCI_BALLYO & ",C." & _
                    gsCAT_TSCI_BALLRO & ",C." & _
                    gsCAT_TSCI_BALLQU & ",C." & _
                    gsCAT_TSCI_BALLCT & ",C." & _
                    gsCAT_TSCI_MARKLABELQUALITY & ",C." & _
                    gsCAT_TSCI_MARKQUALITY & ",C." & _
                    gsCAT_TSCI_LEADSTMAXTOLERANCE & ",C." & _
                    gsCAT_TSCI_LEADSTMINTOLERANCE & ",C." & _
                    gsCAT_TSCI_LEADSTMAX & ",C." & _
                    gsCAT_TSCI_LEADSTMIN & ",C." & _
                    gsCAT_TSCI_MKSCORETHRESHOLD & ",C." & _
                    gsCAT_TSCI_MKCONTRASTTHRESHOLD & ",C." & _
                    gsCAT_TSCI_MKFIRSTTEMPTHRESHOLD & ",C." & _
                    gsCAT_TSCI_LEADOFFSET & ",C." & _
                    gsCAT_TSCI_GLASSTYPE & ",C." & _
                    gsCAT_TSCI_Z2PITCH & ",C." & _
                    gsCAT_TSCI_CAMERAPITCH

        'ADD BY Sam start on 20081029 for ReqNo:M200810038
            sSQL = sSQL & ",C." & _
                    gsCAT_TSCI_LEAD_WIDTH & ",C." & _
                    gsCAT_TSCI_REMARK
        'ADD BY Sam end on 20081029 for ReqNo:M200810038


        'Modify by Sam start on 20081014 for ReqNo:M200809045
        '" and C." & gsCAT_TSCI_EQ_ID & "='" & UCase(Me.txtEqId.Text) & "'"

            'Modify sEqId to me.txtEqId.text by Tony on 20091208 for  FtDailyScheduleAlarmProject
            'Add sPgm by Tony on 20100209 for Req.JC201000024
            sSQL = sSQL & " from " & gsCAT_TBL_LOT_ATTRIBUTE & " a," & _
                               gsCAT_TBL_IPN_MASTER & " b," & _
                               gsCAT_TBL_SCANNER_CONFI & " C" & _
                   " where a." & gsCAT_TLA_LOTID & "='" & oLot.Id & "'" & _
                   " and a." & gsCAT_TLATT_IPN & " =b." & gsCAT_TIM_IPN & _
                   " and b." & gsCAT_TIM_CARRIER_TYPE & "=C." & gsCAT_TSCI_CARRIER_TYPE & "(+)" & _
                   " and b." & gsCAT_TIM_PACKAGE_CODE & "=C." & gsCAT_TSCI_PKG_CODE & "(+)" & _
                   " and b." & gsCAT_TIM_PIN_COUNT & "=C." & gsCAT_TSCI_PIN_CONT & "(+)" & _
                   " and b." & gsCAT_TIM_BODY_SIZE & "=C." & gsCAT_TSCI_BODY_SIZE & "(+)" & _
                   " and b." & gsCAT_TIM_BRAND & "=C." & gsCAT_TSCI_BRAND & "(+)" & _
                   " and c." & gsCAT_TSCI_PGM & "='" & sPgm & "'" & _
                   " and C." & gsCAT_TSCI_EQ_ID & "='" & UCase(Me.txtEqId.Text) & "'" & _
                   " and C." & gsCAT_TSCI_DELETE_FLAG & "='N'"

        'Modify by Sam end on 20081014 for ReqNo:M200809045

            'Marked by Tony Start on 20091207 for FtDailyScheduleAlarmProject.
            'add by Sam start on 20080925 for ReqNO:M200809045
    '        If bSpecial_PGM = False Then    '無客戶指定時增加判斷PGM DEFAULT='Y'
    '            sSql = sSql & " and C." & gsCAT_TSCI_PGMDEFAULT & "='Y' "
    '        End If
            'add by Sam end on 20080925 for ReqNO:M200809045
            'Marked by Tony End on 20091207 for FtDailyScheduleAlarmProject.
            
            '<comment by Jack on 2016/02/02 for 2016 MES Phase-1 特殊客戶Scan需求攔截機制 專案>
            ' oRS1 在 以下 "LendAcc部分" 會用到, 不要誤mark.
            Set oRS1 = moProRawSql.QueryDatabase(sSQL)
            
            'Marked by Tony Start on 20091207 for FtDailyScheduleAlarmProject.
    '        If oRS1.Count > 0 Then
            'add by Sam start on 20080925 for ReqNO:M200809045
    '            If bSpecial_PGM = False Then        '無客戶指定時帶入預設PGM
    '                sPgm = oRS1.Item(1).Item(gsCAT_TSCI_PGM)
    '            End If
            'add by Sam end on 20080925 for ReqNO:M200809045
            'Marked by Tony End on 20091207 for FtDailyScheduleAlarmProject.
    
                .range("F" & CStr(iPosBasicInfo + 3)).Select
                .selection.horizontalalignment = -4131
            
                'Add by Tony Start on 20091208 for FtDailyScheduleAlramProject.
                'Call GetSpecialPGM(Me.txtEqId.Text, sPgm) Moved by Tony on 20100209 for Req.JC201000024
                'Add by Tony End on 20091208 for FtDailyScheduleAlramProject.
            
    '            .activeCell.FormulaR1C1 = oRS1.Item(1).Item(gsCAT_TSCI_PGM)
                .activeCell.FormulaR1C1 = sPgm 'add by Sam on 20081014 for ReqNO:M200809045
                
                'add by Sam start on 20081029 for ReqNo:M200810038 增加PGM條碼
                .range("L" & CStr(iPosBasicInfo + 3)).Select
                .selection.horizontalalignment = -4131
                .activeCell.FormulaR1C1 = "*" & sPgm & "*"
                'add by Sam end on 20081029 for ReqNo:M200810038
            'End If'Marked by Tony on 20091207 for FtDailyScheduleAlarmProject.
            
        End If
            
       'Mark by Sam start on 20100623 for ReqNo:JC201000123,因為需要先取得Erunticno來印出ERUN RECIPE,所以整塊往上搬
'        .range("S" & CStr(iPosBasicInfo)).Select
'        .activeCell.FormulaR1C1 = oLot.CustomAttributes(gsLOT_CUSTOMATTR_LOT_OWNER)
'
'        .range("S" & CStr(iPosBasicInfo + 2)).Select
'        .activeCell.FormulaR1C1 = UCase(Me.txtEqId.Text)
'
'        sSql = " select " & gsCAT_TLI_ERUNTICNO & "," & gsCAT_TLI_SAPRWNO & _
'               " from " & gsCAT_TBL_LOT_INFO & _
'               " where lotid='" & oLot.Id & "'"
'        Set oRS2 = moProRawSql.QueryDatabase(sSql)
'        If oRS2.Count > 0 Then
'            .range("S" & CStr(iPosBasicInfo + 1)).Select
'            If oRS2.Item(1).Item("erunticno") = "" Then
'                .activeCell.FormulaR1C1 = oRS2.Item(1).Item("saprwno")
'            ElseIf oRS2.Item(1).Item("saprwno") = "" Then
'                .activeCell.FormulaR1C1 = oRS2.Item(1).Item("erunticno")
'            Else
'                .activeCell.FormulaR1C1 = oRS2.Item(1).Item("erunticno") & ", " & oRS2.Item(1).Item("saprwno")
'            End If
'        End If
        'Mark by Sam END on 20100623 for ReqNo:JC201000123
        
        If UCase(msFormType) = "T&R" Then
        
            sSQL = "select " & gsCAT_TPBS_CARRIER_QTY & " from " & gsCAT_TBL_PRM_BE_SPEC & _
                   " where " & gsCAT_TPBS_IPN & "='" & oLot.CustomAttributes(gsLOT_CUSTOMATTR_IPN) & "'" & _
                   " and " & gsCAT_TPBS_DELETE_FLAG & "='N'" & _
                   " and " & gsCAT_TPBS_DEFAULTS & "='Y'"
            Set oRs2 = moProRawSql.QueryDatabase(sSQL)
            If oRs2.Count > 0 Then
            'Modify by Sam start on 20081029 for ReqNo:M200810038 位置下移1格 3 -> 4
'                .range("S" & CStr(iPosBasicInfo + 3)).Select
                .range("S" & CStr(iPosBasicInfo + 4)).Select
            'Modify by Sam END on 20081029 for ReqNo:M200810038
                .activeCell.FormulaR1C1 = oRs2.Item(1).Item(gsCAT_TPBS_CARRIER_QTY)
            End If
        End If
                    
        'Add by Sam start on 20100611 for ReqNo:JC201000123
        If msFormType = gsRESOURCE_COMBO Then
            sSQL = " select A." & gsCAT_TEI_CONTACTBOARD_ID & " " & _
                   " FROM " & gsCAT_TBL_EQ_INFO & " A " & _
                   " WHERE A." & gsCAT_TEI_EQ_ID & "='" & msEqpId & "' "
            Set oRS1 = moProRawSql.QueryDatabase(sSQL)
            If oRS1.Count > 0 Then
                'CONTACTBOARD_ID
                .range("F" & CStr(iPosBasicInfo + 13)).Select
                .activeCell.FormulaR1C1 = oRS1.Item(1).Item(gsCAT_TEI_CONTACTBOARD_ID)
            End If
        Else
        '*****
        'LendAcc部分
        '*****
        'Add by Sam end on 20100611 for ReqNo:JC201000123
            If oRS1.Count > 0 Then
                
            'Modify by Sam start on 20081029 for ReqNo:M200810038  位置由T改為G
    '            .range("T" & CStr(iPosLendAcc)).Select
    '            .activeCell.FormulaR1C1 = oRS1.Item(1).Item(gsCAT_TSCI_GLASSTYPE)
    '
    '            .range("T" & CStr(iPosLendAcc + 1)).Select
    '            .activeCell.FormulaR1C1 = oRS1.Item(1).Item(gsCAT_TSCI_LENS)
    '
    '            .range("T" & CStr(iPosLendAcc + 2)).Select
    '            .activeCell.FormulaR1C1 = oRS1.Item(1).Item(gsCAT_TSCI_Z1PITCH)
    '
    '            .range("T" & CStr(iPosLendAcc + 3)).Select
    '            .activeCell.FormulaR1C1 = oRS1.Item(1).Item(gsCAT_TSCI_Z2PITCH)
    '
    '            .range("T" & CStr(iPosLendAcc + 4)).Select
    '            .activeCell.FormulaR1C1 = oRS1.Item(1).Item(gsCAT_TSCI_VACUUMCUPZ1)
    '
    '            .range("T" & CStr(iPosLendAcc + 5)).Select
    '            .activeCell.FormulaR1C1 = oRS1.Item(1).Item(gsCAT_TSCI_VACUUMCUPZ2)
                .range("G" & CStr(iPosLendAcc)).Select
                .activeCell.FormulaR1C1 = oRS1.Item(1).Item(gsCAT_TSCI_GLASSTYPE)
                
                .range("G" & CStr(iPosLendAcc + 1)).Select
                .activeCell.FormulaR1C1 = oRS1.Item(1).Item(gsCAT_TSCI_LENS)
                
                .range("G" & CStr(iPosLendAcc + 2)).Select
                .activeCell.FormulaR1C1 = oRS1.Item(1).Item(gsCAT_TSCI_Z1PITCH)
                
                .range("G" & CStr(iPosLendAcc + 3)).Select
                .activeCell.FormulaR1C1 = oRS1.Item(1).Item(gsCAT_TSCI_Z2PITCH)
                
                .range("G" & CStr(iPosLendAcc + 4)).Select
                .activeCell.FormulaR1C1 = oRS1.Item(1).Item(gsCAT_TSCI_VACUUMCUPZ1)
                
                .range("G" & CStr(iPosLendAcc + 5)).Select
                .activeCell.FormulaR1C1 = oRS1.Item(1).Item(gsCAT_TSCI_VACUUMCUPZ2)
    
            'Modify by Sam end on 20081029 for ReqNo:M200810038
                
                If UCase(msFormType) = "T&R" Then
            'Modify by Sam start on 20081029 for ReqNo:M200810038  位置由6789..上移至01234..
    
    '                .range("T" & CStr(iPosLendAcc + 6)).Select
    '                .activeCell.FormulaR1C1 = oRS1.Item(1).Item(gsCAT_TSCI_CAMERAPITCH)
    '
    '                .range("T" & CStr(iPosLendAcc + 7)).Select
    '                .activeCell.FormulaR1C1 = oRS1.Item(1).Item(gsCAT_TSCI_SEALHEAD)
    '
    '                .range("T" & CStr(iPosLendAcc + 8)).Select
    '                .activeCell.FormulaR1C1 = oRS1.Item(1).Item(gsCAT_TSCI_SEALPITCH)
    '
    '                .range("T" & CStr(iPosLendAcc + 9)).Select
    '                .activeCell.FormulaR1C1 = oRS1.Item(1).Item(gsCAT_TSCI_SEALTEMP)
    '
    '                .range("T" & CStr(iPosLendAcc + 10)).Select
    '                .activeCell.FormulaR1C1 = oRS1.Item(1).Item(gsCAT_TSCI_SEALDWELL)
    '
    '                .range("T" & CStr(iPosLendAcc + 11)).Select
    '                .activeCell.FormulaR1C1 = oRS1.Item(1).Item(gsCAT_TSCI_SEALPRESSURE)
                    .range("T" & CStr(iPosLendAcc)).Select
                    .activeCell.FormulaR1C1 = oRS1.Item(1).Item(gsCAT_TSCI_CAMERAPITCH)
                
                    .range("T" & CStr(iPosLendAcc + 1)).Select
                    .activeCell.FormulaR1C1 = oRS1.Item(1).Item(gsCAT_TSCI_SEALHEAD)
                    
                    .range("T" & CStr(iPosLendAcc + 2)).Select
                    .activeCell.FormulaR1C1 = oRS1.Item(1).Item(gsCAT_TSCI_SEALPITCH)
                    
                    .range("T" & CStr(iPosLendAcc + 3)).Select
                    .activeCell.FormulaR1C1 = oRS1.Item(1).Item(gsCAT_TSCI_SEALTEMP)
                    
                    .range("T" & CStr(iPosLendAcc + 4)).Select
                    .activeCell.FormulaR1C1 = oRS1.Item(1).Item(gsCAT_TSCI_SEALDWELL)
                    
                    .range("T" & CStr(iPosLendAcc + 5)).Select
                    .activeCell.FormulaR1C1 = oRS1.Item(1).Item(gsCAT_TSCI_SEALPRESSURE)
    
            'Modify by Sam END on 20081029 for ReqNo:M200810038
                
                End If
                
                'Add by Sam start on 20081029 for ReqNO:M200810038
                '****
                '特殊架機注意事項
                '****
                .range("B" & CStr(iPosLendAcc + 7)).Select
                .activeCell.FormulaR1C1 = oRS1.Item(1).Item(gsCAT_TSCI_REMARK)
    
                'Add by Sam end on 20081029 for ReqNO:M200810038
            
                '*****
                'HWSetupQualityCheck部分-Lead/BGA inspection recipe:
                '*****
                'Mark by Sam start on 20100624 for ReqNo:JC201000123,因為自動化所以這部份不再使用
'                .range("I" & CStr(iPosHWSetupQualityChk)).Select
'                .activeCell.FormulaR1C1 = oRS1.Item(1).Item(gsCAT_TSCI_LEADCO_TWEEZE)
'
'                .range("I" & CStr(iPosHWSetupQualityChk) + 1).Select
'                .activeCell.FormulaR1C1 = oRS1.Item(1).Item(gsCAT_TSCI_LEADPITCH_SPACING)
'
'                .range("I" & CStr(iPosHWSetupQualityChk) + 2).Select
'                .activeCell.FormulaR1C1 = oRS1.Item(1).Item(gsCAT_TSCI_LEADSKEW)
'
'                .range("I" & CStr(iPosHWSetupQualityChk) + 3).Select
'                .activeCell.FormulaR1C1 = oRS1.Item(1).Item(gsCAT_TSCI_LEADLD)
'
'                .range("I" & CStr(iPosHWSetupQualityChk) + 4).Select
'                .activeCell.FormulaR1C1 = oRS1.Item(1).Item(gsCAT_TSCI_LEADST)
'
'                .range("I" & CStr(iPosHWSetupQualityChk) + 5).Select
'                .activeCell.FormulaR1C1 = oRS1.Item(1).Item(gsCAT_TSCI_LEADTD)
'
'                .range("I" & CStr(iPosHWSetupQualityChk) + 6).Select
'                .activeCell.FormulaR1C1 = oRS1.Item(1).Item(gsCAT_TSCI_LEADBURR)
'
'                'Add by Sam start on 20081029 for ReqNo:m200810038 以下全都下移1格
'                .range("I" & CStr(iPosHWSetupQualityChk) + 7).Select
'                .activeCell.FormulaR1C1 = oRS1.Item(1).Item(gsCAT_TSCI_LEAD_WIDTH)
'                'Add by Sam END on 20081029 for ReqNo:m200810038
'
'                .range("I" & CStr(iPosHWSetupQualityChk) + 7 + 1).Select
'                .activeCell.FormulaR1C1 = oRS1.Item(1).Item(gsCAT_TSCI_LEADSTMAX)
'
'                .range("I" & CStr(iPosHWSetupQualityChk) + 8 + 1).Select
'                .activeCell.FormulaR1C1 = oRS1.Item(1).Item(gsCAT_TSCI_LEADSTMIN)
'
'                .range("I" & CStr(iPosHWSetupQualityChk) + 9 + 1).Select
'                .activeCell.FormulaR1C1 = oRS1.Item(1).Item(gsCAT_TSCI_LEADSTMAXTOLERANCE)
'
'                .range("I" & CStr(iPosHWSetupQualityChk) + 10 + 1).Select
'                .activeCell.FormulaR1C1 = oRS1.Item(1).Item(gsCAT_TSCI_LEADSTMINTOLERANCE)
'
'                .range("I" & CStr(iPosHWSetupQualityChk) + 11 + 1).Select
'                .activeCell.FormulaR1C1 = oRS1.Item(1).Item(gsCAT_TSCI_LEADOFFSET)
'
'                '*****
'                'HWSetupQualityCheck部分-Marking inspection recipe:
'                '*****
'
'                .range("W" & CStr(iPosHWSetupQualityChk)).Select
'                .activeCell.FormulaR1C1 = oRS1.Item(1).Item(gsCAT_TSCI_MARKXO)
'
'                .range("W" & CStr(iPosHWSetupQualityChk) + 1).Select
'                .activeCell.FormulaR1C1 = oRS1.Item(1).Item(gsCAT_TSCI_MARKYO)
'
'                .range("W" & CStr(iPosHWSetupQualityChk) + 2).Select
'                .activeCell.FormulaR1C1 = oRS1.Item(1).Item(gsCAT_TSCI_MARKRO)
'
'                .range("W" & CStr(iPosHWSetupQualityChk) + 3).Select
'                .activeCell.FormulaR1C1 = oRS1.Item(1).Item(gsCAT_TSCI_MARKCO)
'
'                .range("W" & CStr(iPosHWSetupQualityChk) + 4).Select
'                .activeCell.FormulaR1C1 = oRS1.Item(1).Item(gsCAT_TSCI_MARKOP)
'
'                .range("W" & CStr(iPosHWSetupQualityChk) + 5).Select
'
'                .activeCell.FormulaR1C1 = oRS1.Item(1).Item(gsCAT_TSCI_MARKUP)
'
'                .range("W" & CStr(iPosHWSetupQualityChk) + 6).Select
'                .activeCell.FormulaR1C1 = oRS1.Item(1).Item(gsCAT_TSCI_MARKBL)
'
'                'Modify by Sam start on 20081029 for ReqNo:M200810038  7以下全都下移一格
'                .range("W" & CStr(iPosHWSetupQualityChk) + 7 + 1).Select
'                .activeCell.FormulaR1C1 = oRS1.Item(1).Item(gsCAT_TSCI_MARKQUALITY)
'
'                .range("W" & CStr(iPosHWSetupQualityChk) + 8 + 1).Select
'                .activeCell.FormulaR1C1 = oRS1.Item(1).Item(gsCAT_TSCI_MARKLABELQUALITY)
'
'                .range("W" & CStr(iPosHWSetupQualityChk) + 9 + 1).Select
'                .activeCell.FormulaR1C1 = oRS1.Item(1).Item(gsCAT_TSCI_MKSCORETHRESHOLD)
'
'                .range("W" & CStr(iPosHWSetupQualityChk) + 10 + 1).Select
'                .activeCell.FormulaR1C1 = oRS1.Item(1).Item(gsCAT_TSCI_MKCONTRASTTHRESHOLD)
'
'                .range("W" & CStr(iPosHWSetupQualityChk) + 11 + 1).Select
'                .activeCell.FormulaR1C1 = oRS1.Item(1).Item(gsCAT_TSCI_MKFIRSTTEMPTHRESHOLD)
'                'Modify by Sam start on 20081029 for ReqNo:M200810038
                'Modify by SAM ON 20100624 FOR ReqNo:JC201000123
            End If
            
        End If
        
        'Add by Tony Start on 2014/02/25 for HW SETUP RECIPE COMPARE project
        '架機確認紀錄(H/W only):內容
        
        sEqID = Trim(Me.txtEqId.Text)
        
         Call getRecipeSpecName(sEqID, oLot.Id, oLot.CustomAttributes(gsLOT_CUSTOMATTR_IPN), _
                             msPackageName, msPinCount, msBodySize, _
                             oRS1, moProRawSql, moAppLog, _
                              moCwMbx)
        
        If Not oRS1 Is Nothing And oRS1.Count > 0 Then
            For iIdx = 1 To oRS1.Count
                If iIdx <= miFTFVI_RECIPE_ROW Then
                    .range("B" & CStr(iPosTitle + 28 + iIdx)).Select 'ID
                    .activeCell.FormulaR1C1 = oRS1.Item(iIdx).Item("col1")
                    
                    .range("C" & CStr(iPosTitle + 28 + iIdx)).Select 'Specname
                    .activeCell.FormulaR1C1 = oRS1.Item(iIdx).Item("specname")
                Else
                    If iIdx <= mi_RECIPE_MAX_ROW Then
                        .range("P" & CStr(iPosTitle + 28 + iIdx - miFTFVI_RECIPE_ROW)).Select 'ID
                        .activeCell.FormulaR1C1 = oRS1.Item(iIdx).Item("col1")
                        
                        .range("Q" & CStr(iPosTitle + 28 + iIdx - miFTFVI_RECIPE_ROW)).Select 'Specname
                        .activeCell.FormulaR1C1 = oRS1.Item(iIdx).Item("specname")
                    End If
                End If
            Next
        Else
            UtShowMsgBox "尚未維護，請確認架機需求或請維護相關資訊。請洽設備工程師。" & vbCrLf & "No data maintenance, make sure the equipment frame machine or maintain information. Please call HW to check."
        End If
        'Add by Tony End on 2014/02/25 for HW SETUP RECIPE COMPARE project
    
    End With
    Set oRS1 = Nothing
    Set oRs2 = Nothing
    
'----
' Done
'----

ExitHandler:
    ' NOTE 1:
    ' MUST CALL GetErrInfo() here first before another action
    Call GetErrInfo(msMODULE_ID, sProcID, typErrInfo, Erl)
    Call LogProcOut(msMODULE_ID, sProcID, typErrInfo, moAppLog)
    ' <Your cleaning up codes goes here...>
    Set oLot = Nothing
    
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
Private Sub txtEqId_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Me.cmdOK.SetFocus
End Sub

'================================================================================
' Sub: GetSpecialPGM()
'--------------------------------------------------------------------------------
' Description:  <FtDailyScheduleAlramProject>
'--------------------------------------------------------------------------------
' Author: Tony Chang , MXIC 2009/12/08
'--------------------------------------------------------------------------------
' Purpose:
' Tbl_special_customer_scanpgm有新增四種SpecialType,故要串出 tscs.pgm 要用
' carrierType + pincount + packagecode + bodySize + Brand + tscs.value
' 1.tscs.value => 利用 tli.vendorcode(AssyVendorCode) or BoxingSpecNo or IcDrawing or tpbs.customerno(defaults='Y')
'   去Tbl_special_customer_scanpgm所串出.
' 2.有了carrierType+pincount+packagecode+bodySize+Brand+tscs.value 後,則可以至 Tbl_special_customer_scanpgm串出
'   pgm.
'================================================================================
'Marked by Jack on 2016/02/02 for 2016 MES Phase-1 特殊客戶Scan需求攔截機制 專案 <Start>
'Private Sub GetSpecialPGM(ByVal sEqID As String, ByRef sPGM As String)
'On Error GoTo ExitHandler:
'Dim sProcID                 As String
'Dim typErrInfo              As tErrInfo
'Dim sSQL                    As String
'Dim colSQLResult            As Collection
'Dim iIdx                    As Integer
'Dim sSpecialType            As String
'Dim sCarrierType            As String
'Dim sPkgCode                As String
'Dim sPinCount               As String
'Dim sBodySize               As String
'Dim sBrand                  As String
'Dim sValue                  As String
'Dim aSpecialType()          As Variant
'Dim bFoundFlag              As Boolean
'
''----
'' Init
''----
'
'    sProcID = "GetSpecialPGM"
'    Call LogProcIn(msMODULE_ID, sProcID, moAppLog)
'
''----
'' Condition Checking
''----
'    ' <Put your condition checking codes here>...
'
''----
'' Action
''----
'
'    aSpecialType = Array("BOXINGSPECNO", "ICDRAWING", "CUSTOMERNO", "ASSYVENDORCODE")
'
'    For iIdx = 0 To 3
'
'        sSQL = "Select B." & gsCAT_TSCS_VALUE & "," & _
'                      "B." & gsCAT_TSCS_PGM & "," & _
'                      "B." & gsCAT_TSCS_CARRIERTYPE & "," & _
'                      "B." & gsCAT_TSCS_PKGCODE & "," & _
'                      "B." & gsCAT_TSCS_PINCONT & "," & _
'                      "B." & gsCAT_TSCS_BODYSIZE & "," & _
'                      "B." & gsCAT_TSCS_BRAND & _
'              " from (select tli." & gsCAT_TLI_VENDORCODE & "," & _
'                            "tim." & gsCAT_TIM_BOXING_SPEC_NO & "," & _
'                            "tim." & gsCAT_TIM_ICDRAWING & "," & _
'                            "TPBS." & gsCAT_TPBS_CUSTOMER_NO & "," & _
'                            "tim." & gsCAT_TIM_CARRIER_TYPE & "," & _
'                            "tim." & gsCAT_TIM_PACKAGE_CODE & "," & _
'                            "tim." & gsCAT_TIM_PIN_COUNT & "," & _
'                            "tim." & gsCAT_TIM_BODY_SIZE & "," & _
'                            "tim." & gsCAT_TIM_BRAND & _
'              "       from " & gsCAT_TBL_LOT_INFO & " tli," & _
'                               gsCAT_TBL_LOT_ATTRIBUTE & " tlatt," & _
'                               gsCAT_TBL_IPN_MASTER & " tim, " & _
'                             "(SELECT * FROM " & gsCAT_TBL_PRM_BE_SPEC & " WHERE " & gsCAT_TPBS_DEFAULTS & " = 'Y') TPBS "
'          sSQL = sSQL & " Where TLI." & gsCAT_TLI_LOT_ID & " = TLATT." & gsCAT_TLATT_LOTID & _
'              "       and tlatt." & gsCAT_TLATT_IPN & " = tim." & gsCAT_TIM_IPN & _
'              "       and TIM." & gsCAT_TIM_IPN & " = TPBS." & gsCAT_TPBS_IPN & _
'              "       and TLI." & gsCAT_TLI_LOT_ID & " like '" & Me.txtLotID.Text & "'" & _
'              "       and TLATT." & gsCAT_TLATT_IPN & " in (select F2.VALDATA " & _
'                                                          " from fwlot F1, fwlot_pn2m F2 " & _
'                                                          " where F2.KEYDATA = 'IPN' " & _
'                                                          " and F1.SYSID = F2.FROMID " & _
'                                                          " and F1.APPID like '" & Me.txtLotID.Text & "')) A,"
'           sSQL = sSQL & "(select t." & gsCAT_TSCS_VALUE & "," & _
'           "                      t." & gsCAT_TSCS_PGM & "," & _
'           "                      t." & gsCAT_TSCS_CARRIERTYPE & "," & _
'           "                      t." & gsCAT_TSCS_PKGCODE & "," & _
'           "                      t." & gsCAT_TSCS_PINCONT & "," & _
'           "                      t." & gsCAT_TSCS_BODYSIZE & "," & _
'           "                      t." & gsCAT_TSCS_BRAND & _
'           "               from " & gsCAT_TBL_SPECIAL_CUSTOMER_SCANPGM & " t " & _
'           "               where (UPPER(t.specialtype)) = '" & aSpecialType(iIdx) & "') B," & _
'           "              (SELECT * FROM " & gsCAT_TBL_SCANNER_CONFI & " TSCI WHERE TSCI." & gsCAT_TSCI_DELETE_FLAG & " = 'N') C"
'
'           sSQL = sSQL & " Where a." & gsCAT_TIM_CARRIER_TYPE & " = b." & gsCAT_TSCS_CARRIERTYPE & _
'           "         and a." & gsCAT_TIM_PACKAGE_CODE & " = b." & gsCAT_TSCS_PKGCODE & _
'           "         and a." & gsCAT_TIM_PIN_COUNT & " = b." & gsCAT_TSCS_PINCONT & _
'           "         and a." & gsCAT_TIM_BODY_SIZE & " = b." & gsCAT_TSCS_BODYSIZE & _
'           "         and a." & gsCAT_TIM_BRAND & " = b." & gsCAT_TSCS_BRAND & _
'           "         and (A." & gsCAT_TLI_VENDORCODE & "= B." & gsCAT_TSCS_VALUE & " OR A." & gsCAT_TIM_BOXING_SPEC_NO & "= B." & gsCAT_TSCS_VALUE & " OR " & _
'           "              A." & gsCAT_TIM_ICDRAWING & "= B." & gsCAT_TSCS_VALUE & " OR A." & gsCAT_TPBS_CUSTOMER_NO & "= B." & gsCAT_TSCS_VALUE & ")"
'
'
'           sSQL = sSQL & " and a." & gsCAT_TIM_CARRIER_TYPE & "= C." & gsCAT_TSCI_CARRIER_TYPE & _
'           "         and a." & gsCAT_TIM_PACKAGE_CODE & "= C." & gsCAT_TSCI_PKG_CODE & _
'           "         and a." & gsCAT_TIM_PIN_COUNT & "= C." & gsCAT_TSCI_PIN_CONT & _
'           "         and a." & gsCAT_TIM_BODY_SIZE & "= C." & gsCAT_TSCI_BODY_SIZE & _
'           "         and a." & gsCAT_TIM_BRAND & "= C." & gsCAT_TSCI_BRAND & _
'           "         AND C.EQID = '" & Me.txtEqId.Text & "'"
'
'        Set colSQLResult = moProRawSql.QueryDatabase(sSQL)
'
'        If Not colSQLResult Is Nothing Then
'            If colSQLResult.Count > 0 Then
'                sValue = colSQLResult.Item(1).Item(gsCAT_TSCS_VALUE)
'                sPGM = colSQLResult.Item(1).Item(gsCAT_TSCS_PGM)
'                If sValue <> "" And sPGM <> "" Then
'                    Exit For '已找到
'                End If
'            ElseIf iIdx = 3 Then '四個SpecialType都沒找到,找default.
'
'            sSQL = "Select c." & gsCAT_TSCI_PGM & _
'                   " from (select tim." & gsCAT_TIM_CARRIER_TYPE & "," & _
'                   "               tim." & gsCAT_TIM_PACKAGE_CODE & "," & _
'                   "               tim." & gsCAT_TIM_PIN_COUNT & "," & _
'                   "               tim." & gsCAT_TIM_BODY_SIZE & "," & _
'                   "               tim." & gsCAT_TIM_BRAND & _
'                   "          from " & gsCAT_TBL_LOT_INFO & " tli," & gsCAT_TBL_LOT_ATTRIBUTE & " tlatt, " & gsCAT_TBL_IPN_MASTER & " tim" & _
'                   "         Where TLI." & gsCAT_TLI_LOT_ID & " = TLATT." & gsCAT_TLATT_LOTID & _
'                   "           and tlatt." & gsCAT_TLATT_IPN & " = tim." & gsCAT_TIM_IPN & _
'                   "           and TLI." & gsCAT_TLI_LOT_ID & " like '" & Me.txtLotID.Text & "'" & _
'                   "           and TLATT." & gsCAT_TLATT_IPN & " in (select F2.VALDATA" & _
'                   "                               from fwlot F1, fwlot_pn2m F2" & _
'                   "                              where F2.KEYDATA = 'IPN'" & _
'                   "                                and F1.SYSID = F2.FROMID" & _
'                   "                                and F1.APPID like '" & Me.txtLotID.Text & "')) A, " & _
'                   "       (SELECT * FROM " & gsCAT_TBL_SCANNER_CONFI & " TSCI WHERE TSCI." & gsCAT_TSCI_DELETE_FLAG & " = 'N') C" & _
'                   " Where a." & gsCAT_TIM_CARRIER_TYPE & " = C." & gsCAT_TSCI_CARRIER_TYPE & _
'                   "   and a." & gsCAT_TIM_PACKAGE_CODE & " = C." & gsCAT_TSCI_PKG_CODE & _
'                   "   and a." & gsCAT_TIM_PIN_COUNT & " = C." & gsCAT_TSCI_PIN_CONT & _
'                   "   and a." & gsCAT_TIM_BODY_SIZE & " = C." & gsCAT_TSCI_BODY_SIZE & _
'                   "   and a." & gsCAT_TIM_BRAND & "= C." & gsCAT_TSCI_BRAND & _
'                   "   and c." & gsCAT_TSCI_EQ_ID & " = '" & Me.txtEqId.Text & "'" & _
'                   "   and c." & gsCAT_TSCI_PGMDEFAULT & " = 'Y'"
'
'                   Set colSQLResult = moProRawSql.QueryDatabase(sSQL)
'
'                   If Not colSQLResult Is Nothing Then
'                       If colSQLResult.Count > 0 Then
'                           sPGM = colSQLResult.Item(1).Item(gsCAT_TSCS_PGM)
'                       End If
'                   End If
'            End If
'        End If
'
'    Next iIdx
'
'
''----
'' Done
''----
'
'ExitHandler:
'    ' NOTE 1:
'    ' MUST CALL GetErrInfo() here first before another action
'    Call GetErrInfo(msMODULE_ID, sProcID, typErrInfo, Erl)
'    Call LogProcOut(msMODULE_ID, sProcID, typErrInfo, moAppLog)
'    ' <Your cleaning up codes goes here...>
'ErrorHandler:
'    If typErrInfo.lErrNumber Then
'        ' NOTE 2:
'        ' If you have custom handling of some Errors, please
'        ' UN-REMARED the following Select Case block!
'        ' Also, modify if neccessarily!!!
'        '---- Start of Select Case Block ----
'        Select Case typErrInfo.lErrNumber
'            Case glERR_INVALIDOBJECT
'                ' Retry code goes here...
'            Case Else
'                typErrInfo.sUserText = "Fail to execute application, please call IT support!!" & vbCrLf & _
'                                        "程式執行失敗, 請洽IT人員處理"
'            End Select
'        '---- Start of Select Case Block ----
'        On Error GoTo ExitHandler:
'        Call HandleError(False, typErrInfo, , moAppLog, True)
'    End If
'End Sub
'Marked by Jack on 2016/02/02 for 2016 MES Phase-1 特殊客戶Scan需求攔截機制 專案 <End>

'================================================================================
' Sub: LaserSetupForm()
'--------------------------------------------------------------------------------
' Description:  <JC201000123>
'--------------------------------------------------------------------------------
' Author:       Sam Chen, Cit 2010/06/11
'================================================================================
Private Sub LaserSetupForm(ByRef appXL As Object, ByRef oLot As FwLot, _
                           ByRef iPosTitle As Integer)
On Error GoTo ExitHandler:
Dim sProcID                 As String
Dim typErrInfo              As tErrInfo
Dim sSQL                    As String
Dim oRS1                    As Collection
Dim oRs2                    As Collection
Dim iPosBasicInfo           As Integer
Dim sProdgroup              As String

Dim iIdx As Integer
Dim sEqID   As String 'Add by Tony on 2014/03/05 for HW SETUP RECIPE COMPARE project.

'----
' Init
'----
    sProcID = "LaserSetupForm"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog)

'----
' Condition Checking
'----
    ' <Put your condition checking codes here>...

'----
' Action
'----
    ' <Put your Action codes here>...
    iPosBasicInfo = 8
    
    With appXL
    
        '*****
        'COPY template
        '*****
        
        .Sheets(msFormType).Select
        .Cells.Select
        .selection.NumberFormatLocal = "@"
        .selection.Copy
        .Sheets(1).Select
        .ActiveSheet.Paste
        .Application.CutCopyMode = False
        
        '*****
        'Title部分
        '*****
        msOprId = moFwOPR.ActiveUser.UserName
        '.range("F" & CStr(iPosTitle + 2)).Select'Modified by Tony on 2014/03/18 for HW SETUP RECIPE COMPARE project.
        .range("G" & CStr(iPosTitle + 2)).Select
        .selection.horizontalalignment = -4131
        .activeCell.FormulaR1C1 = msOprId
        
        sSQL = "select shiftcode from tbl_mxic_emp where empno='" & msOprId & "'"
        Set oRS1 = moProRawSql.QueryDatabase(sSQL)
        If oRS1.Count > 0 Then
            '.range("J" & CStr(iPosTitle + 2)).Select'Modified by Tony on 2014/03/18 for HW SETUP RECIPE COMPARE project.
            .range("L" & CStr(iPosTitle + 2)).Select
            .selection.horizontalalignment = -4131
            .activeCell.FormulaR1C1 = oRS1.Item(1).Item(1)
        End If
        Set oRS1 = Nothing
        
        .range("S" & CStr(iPosTitle + 2)).Select
        .selection.horizontalalignment = -4131
        .activeCell.FormulaR1C1 = Format(Now, "YYYY/MM/DD HH:MM:SS")
        
        '*****
        'BasicInfo部分
        '*****
        
        '.range("E" & CStr(iPosBasicInfo)).Select'Modified by Tony on 2014/03/18 for HW SETUP RECIPE COMPARE project.
        .range("H" & CStr(iPosBasicInfo)).Select
        .selection.horizontalalignment = -4131
        .activeCell.FormulaR1C1 = oLot.Id
        
        
        '.range("F" & CStr(iPosBasicInfo + 1)).Select'Modified by Tony on 2014/03/18 for HW SETUP RECIPE COMPARE project.
        .range("H" & CStr(iPosBasicInfo + 1)).Select
        .selection.horizontalalignment = -4131
        .activeCell.FormulaR1C1 = oLot.CustomAttributes(gsLOT_CUSTOMATTR_IPN)
                              
        .range("H" & CStr(iPosBasicInfo + 2)).Select
        .activeCell.FormulaR1C1 = UCase(Me.txtEqId.Text)
                              
        'Add by Tony BodySize/PinCount/PackageName on 2014/03/10 for HW SETUP RECIPE COMPARE project.
        sSQL = " select " & gsCAT_TIM_PACKAGE_NAME & "," & gsCAT_TIM_PIN_COUNT & "," & gsCAT_TIM_MARKING_SPEC_NO & "," & gsCAT_TIM_BODY_SIZE & "," & gsCAT_TIM_PACKAGE_NAME & "," & gsCAT_TIM_PIN_COUNT & _
               " from " & gsCAT_TBL_IPN_MASTER & _
               " where " & gsCAT_TIM_IPN & "='" & oLot.CustomAttributes(gsLOT_CUSTOMATTR_IPN) & "'"
        Set oRS1 = moProRawSql.QueryDatabase(sSQL)
        
        If oRS1.Count > 0 Then
            .range("H" & CStr(iPosBasicInfo + 3)).Select
            .selection.horizontalalignment = -4131
            .activeCell.FormulaR1C1 = oRS1.Item(1).Item(gsCAT_TIM_PACKAGE_NAME)
            
            '.range("G" & CStr(iPosBasicInfo + 4)).Select'Modified by Tony on 2014/03/18 for HW SETUP RECIPE COMPARE project.
            .range("H" & CStr(iPosBasicInfo + 4)).Select
            .selection.horizontalalignment = -4131
            .activeCell.FormulaR1C1 = oRS1.Item(1).Item(gsCAT_TIM_PIN_COUNT)

            .range("H" & CStr(iPosBasicInfo + 5)).Select
            .selection.horizontalalignment = -4131
            .activeCell.FormulaR1C1 = oRS1.Item(1).Item(gsCAT_TIM_MARKING_SPEC_NO)
            
            'Add by Tony Start on 2014/03/10 for HW SETUP RECIPE COMPARE project.
            '(先存放值,以便後續使用)
            msPackageName = oRS1.Item(1).Item("packagename")
            msPinCount = oRS1.Item(1).Item("pincount")
            msBodySize = oRS1.Item(1).Item("bodysize")
            'Add by Tony End on 2014/03/10 for HW SETUP RECIPE COMPARE project.
            
        End If
        
        '.range("E" & CStr(iPosBasicInfo + 6)).Select 'Modified by Tony on 2014/03/18 for HW SETUP RECIPE COMPARE project.
        .range("H" & CStr(iPosBasicInfo + 6)).Select
        .selection.horizontalalignment = -4131
        .activeCell.FormulaR1C1 = oLot.CustomAttributes(gsLOT_CUSTOMATTR_DATECODE)
        
        'Add by Tony Start on 2014/02/25 for HW SETUP RECIPE COMPARE project
        '架機確認紀錄(H/W only):內容
        
         sEqID = Trim(Me.txtEqId.Text)
        
         Call getRecipeSpecName(sEqID, oLot.Id, oLot.CustomAttributes(gsLOT_CUSTOMATTR_IPN), _
                             msPackageName, msPinCount, msBodySize, _
                             oRS1, moProRawSql, moAppLog, _
                              moCwMbx)
        
        If Not oRS1 Is Nothing And oRS1.Count > 0 Then
            For iIdx = 1 To oRS1.Count
                If iIdx <= miLASER_RECIPE_ROW Then
                    .range("B" & CStr(iPosTitle + 16 + iIdx)).Select 'ID
                    .activeCell.FormulaR1C1 = oRS1.Item(iIdx).Item("col1")
                    
                    .range("C" & CStr(iPosTitle + 16 + iIdx)).Select 'Specname
                    .activeCell.FormulaR1C1 = oRS1.Item(iIdx).Item("specname")
                 'Marked by Tony Start on 2014/04/09 for HW SETUP RECIPE COMPARE project'(超過20行就不印出)
'                Else
'                    If iIdx <= mi_RECIPE_MAX_ROW Then
'                        .range("N" & CStr(iPosTitle + 16 + iIdx - miLASER_RECIPE_ROW)).Select 'ID
'                        .activeCell.FormulaR1C1 = oRS1.Item(iIdx).Item("col1")
'
'                        .range("O" & CStr(iPosTitle + 16 + iIdx - miLASER_RECIPE_ROW)).Select 'Specname
'                        .activeCell.FormulaR1C1 = oRS1.Item(iIdx).Item("specname")
'                    End If
                'Marked by Tony End on 2014/04/09 for HW SETUP RECIPE COMPARE project
                End If
            Next
        Else
            UtShowMsgBox "尚未維護，請確認架機需求或請維護相關資訊。請洽設備工程師。" & vbCrLf & "No data maintenance, make sure the equipment frame machine or maintain information. Please call HW to check."
        End If
        'Add by Tony End on 2014/02/25 for HW SETUP RECIPE COMPARE project
        
      
    End With
    Set oRS1 = Nothing
    Set oRs2 = Nothing
    
'----
' Done
'----

ExitHandler:
    ' NOTE 1:
    ' MUST CALL GetErrInfo() here first before another action
    Call GetErrInfo(msMODULE_ID, sProcID, typErrInfo, Erl)
    Call LogProcOut(msMODULE_ID, sProcID, typErrInfo, moAppLog)
    ' <Your cleaning up codes goes here...>
    Set oLot = Nothing
    
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
' Sub: ComboSetupForm()
'--------------------------------------------------------------------------------
' Description:  Project FVI Setup Recipe
'--------------------------------------------------------------------------------
' Author:       Sam Chen , CIT 2013/03/27
'--------------------------------------------------------------------------------
'Modification history:
'1.For bug fix : add oEqp by Tony on 2014/03/03
'================================================================================
Private Sub ComboSetupForm(ByRef appXL As Object, ByRef oLot As FwLot, ByRef oEqp As FwEquipment, _
                           ByRef iPosTitle As Integer)
On Error GoTo ExitHandler:
Dim sProcID                 As String
Dim typErrInfo              As tErrInfo
Dim sSQL                    As String
Dim oRS1                    As Collection
Dim oRs2                    As Collection
Dim iPosBasicInfo           As Integer
Dim iPosLendAcc             As Integer
Dim iPosHWSetupQualityChk   As Integer
Dim sProdgroup              As String

Dim sPgm                    As String
Dim sEqID                   As String
Dim bSpecial_PGM            As Boolean
Dim iIdx As Integer

Dim sERunTicNo              As String
Dim sFollowProd             As String
Dim sFollowOI               As String
'Dim oEqp                    As FwEquipment'Marked by Tony on 2014/03/03
Dim sTapingQty              As String 'Add by Tony on 2014/09/05 for Req.JC201400285

'----
' Init
'----
    sProcID = "ComboSetupForm"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog)

    'Set oEqp = moFwMDL.EquipmentById(msEqpId)'Marked by Tony on 2014/03/03

'----
' Condition Checking
'----
    ' <Put your condition checking codes here>...

'----
' Action
'----
    ' <Put your Action codes here>...
    iPosBasicInfo = 6
    iPosLendAcc = 13
    iPosHWSetupQualityChk = 27
    If UCase(msFormType) = "SCANNER" Then iPosHWSetupQualityChk = 27
    
    With appXL
    
        '*****
        'COPY template
        '*****
        
        .Sheets(msFormType).Select
        .Cells.Select
        .selection.NumberFormatLocal = "@"
        .selection.Copy
        .Sheets(1).Select
        .ActiveSheet.Paste
        .Application.CutCopyMode = False
        
        '*****
        'Title部分
        '*****
        msOprId = moFwOPR.ActiveUser.UserName
        .range("D" & CStr(iPosTitle + 2)).Select
        .selection.horizontalalignment = -4131
        .activeCell.FormulaR1C1 = msOprId
        
        sSQL = "select shiftcode from tbl_mxic_emp where empno='" & msOprId & "'"
        Set oRS1 = moProRawSql.QueryDatabase(sSQL)
        If oRS1.Count > 0 Then
            .range("I" & CStr(iPosTitle + 2)).Select
            .selection.horizontalalignment = -4131
            .activeCell.FormulaR1C1 = oRS1.Item(1).Item(1)
        End If
        Set oRS1 = Nothing
        
        .range("S" & CStr(iPosTitle + 2)).Select
        .selection.horizontalalignment = -4131
        .activeCell.FormulaR1C1 = Format(Now, "YYYY/MM/DD HH:MM:SS")
        
        '*****
        'BasicInfo部分
        '*****
        
        'Add by Sam start on 20170322 for Project 停復測自動化
        .range("I" & CStr(iPosBasicInfo - 1)).Select
        .selection.horizontalalignment = -4131
        .activeCell.FormulaR1C1 = lblStopInfo.Caption
        'Add by Sam end on 20170322 for Project 停復測自動化
        
        .range("F" & CStr(iPosBasicInfo)).Select
        .selection.horizontalalignment = -4131
        .activeCell.FormulaR1C1 = oLot.Id
        
        
        .range("F" & CStr(iPosBasicInfo + 1)).Select
        .selection.horizontalalignment = -4131
        .activeCell.FormulaR1C1 = oLot.CustomAttributes(gsLOT_CUSTOMATTR_IPN)
        
        
        If UCase(msFormType) = gsRESOURCE_COMBO Then
            .range("F" & CStr(iPosBasicInfo + 2)).Select
            .selection.horizontalalignment = -4131
            .activeCell.FormulaR1C1 = oLot.CustomAttributes(gsLOT_CUSTOMATTR_STEPNAME)
        End If
        
        'Add BodySize/PackageName/Pincount by Tony on 2014/03/10 for HW SETUP RECIPE COMPARE project.
        'Add CarrierType by Tony on 2014/09/05 for Req.JC201400285
        sSQL = " select " & gsCAT_TIM_PACKAGE_NAME & "," & gsCAT_TIM_PROD_GROUP & "," & gsCAT_TIM_BODY_SIZE & "," & gsCAT_TIM_PACKAGE_NAME & "," & gsCAT_TIM_PIN_COUNT & "," & gsCAT_TIM_CARRIER_TYPE & _
               " from " & gsCAT_TBL_IPN_MASTER & _
               " where " & gsCAT_TIM_IPN & "='" & oLot.CustomAttributes(gsLOT_CUSTOMATTR_IPN) & "'"
               
        Set oRS1 = moProRawSql.QueryDatabase(sSQL)
        If oRS1.Count > 0 Then
            If UCase(msFormType) = gsRESOURCE_COMBO Then
                .range("F" & CStr(iPosBasicInfo + 3)).Select
            Else
                .range("F" & CStr(iPosBasicInfo + 2)).Select
            End If
            .selection.horizontalalignment = -4131
            .activeCell.FormulaR1C1 = oRS1.Item(1).Item(gsCAT_TIM_PACKAGE_NAME)
            sProdgroup = oRS1.Item(1).Item(gsCAT_TIM_PROD_GROUP)
            
            'Add by Tony Start on 2014/03/10 for HW SETUP RECIPE COMPARE project.
            '(先存放值,以便後續使用)
            msPackageName = oRS1.Item(1).Item("packagename")
            msPinCount = oRS1.Item(1).Item("pincount")
            msBodySize = oRS1.Item(1).Item("bodysize")
            'Add by Tony End on 2014/03/10 for HW SETUP RECIPE COMPARE project.
            
            'Add by Tony Start on 2014/09/05 for Req.JC201400285
            If UCase(oRS1.Item(1).Item(gsCAT_TPBS_CARRIER_TYPE)) = "TUBE" Then
                sTapingQty = "NA"
            End If
            'Add by Tony End on 2014/09/05 for Req.JC201400285
        End If
        
        .range("S" & CStr(iPosBasicInfo)).Select
        .activeCell.FormulaR1C1 = oLot.CustomAttributes(gsLOT_CUSTOMATTR_LOT_OWNER)

        .range("S" & CStr(iPosBasicInfo + 2)).Select
        .activeCell.FormulaR1C1 = UCase(Me.txtEqId.Text)

        sSQL = " select " & gsCAT_TLI_ERUNTICNO & "," & gsCAT_TLI_SAPRWNO & _
               " from " & gsCAT_TBL_LOT_INFO & _
               " where lotid='" & oLot.Id & "'"
        Set oRs2 = moProRawSql.QueryDatabase(sSQL)
        If oRs2.Count > 0 Then
            .range("S" & CStr(iPosBasicInfo + 1)).Select
            If oRs2.Item(1).Item("erunticno") = "" Then
                .activeCell.FormulaR1C1 = oRs2.Item(1).Item("saprwno")
            ElseIf oRs2.Item(1).Item("saprwno") = "" Then
                .activeCell.FormulaR1C1 = oRs2.Item(1).Item("erunticno")
            Else
                .activeCell.FormulaR1C1 = oRs2.Item(1).Item("erunticno") & ", " & oRs2.Item(1).Item("saprwno")
            End If
             sERunTicNo = oRs2.Item(1).Item("erunticno")    'Add by Sam on 20100623 for ReqNo:JC201000123
        End If
                        
        If Trim(sERunTicNo) <> "" Then
            sSQL = " SELECT A." & gsCAT_TER_FOLLOW_OI & " ,A." & gsCAT_TER_FOLLOW_PRODUCT & " " & _
                   " FROM " & gsCAT_TBL_ERUN_REQ & " A " & _
                   " WHERE A." & gsCAT_TER_TICKET_NO & " ='" & sERunTicNo & "' " & _
                   " AND A." & gsCAT_TER_DELETE_FLAG & " ='N' "
                   
            Set oRs2 = moProRawSql.QueryDatabase(sSQL)
            If oRs2.Count > 0 Then
               sFollowOI = oRs2.Item(1).Item(gsCAT_TER_FOLLOW_OI)
               sFollowProd = oRs2.Item(1).Item(gsCAT_TER_FOLLOW_PRODUCT)
            End If
        End If
        
        If msFormType = gsRESOURCE_COMBO Then
            'contactboardtype Added by Jack on 2014/01/13 for JC201400011
            sSQL = " SELECT B." & gsCAT_PSS_DOC_NO & " , " & _
                   " B." & gsCAT_PSS_COMMENTS & " AS stepcomments , " & _
                   " C." & gsCAT_PSES_COMMENTS & "  as eqcomments , " & _
                   " C." & gsCAT_PSES_CONTACTBOARD_TYPE & "  as contactboardtype , " & _
                   " C." & gsCAT_PSES_PG_NAME & " " & _
                   " FROM " & gsCAT_TBL_PROD_STEP_SPEC & " B ," & _
                   gsCAT_TBL_PROD_STEP_EQ_SPEC & " C " & _
                   " WHERE B." & gsCAT_PSS_PROD_GROUP & "=C." & gsCAT_PSES_PROD_GROUP & "(+) " & _
                   " AND B." & gsCAT_PSS_PATH & "=C." & gsCAT_PSES_PATH & "(+) " & _
                   " AND B." & gsCAT_PSS_STEP_NO & "=C." & gsCAT_PSES_STEP_NO & "(+) " & _
                   " AND B." & gsCAT_PSS_DOC_STATUS & "=C." & gsCAT_PSES_DOC_STATUS & "(+) " & _
                   " AND B." & gsCAT_PSS_DOC_STATUS & "='" & gsDOCSTATUS_ACTIVE & "' " & _
                   " AND B." & gsCAT_PSS_STEP_NO & "='" & oLot.CurrentStep.Steps.Item(1).Id & "' " & _
                   " AND B." & gsCAT_PSS_PROD_GROUP & "='" & sProdgroup & "' " & _
                   " AND B." & gsCAT_PSS_PATH & "='" & oLot.CustomAttributes(gsLOT_CUSTOMATTR_ROUTE) & "' " & _
                   " AND C." & gsCAT_PSES_EQTYPE2 & " ='" & oEqp.CustomAttributes(gsEQP_CUSTOMATTR_EQ_TYPE2) & "' "
            Set oRS1 = moProRawSql.QueryDatabase(sSQL)
            If oRS1.Count > 0 Then
                'TECNNO
                .range("F" & CStr(iPosBasicInfo + 5)).Select
                .activeCell.FormulaR1C1 = oRS1.Item(1).Item(gsCAT_PSS_DOC_NO)
                
                'PgName
                .range("F" & CStr(iPosBasicInfo + 4)).Select
                .activeCell.FormulaR1C1 = oRS1.Item(1).Item(gsCAT_PSES_PG_NAME)
                'EQCOMMENTS
                .range("F" & CStr(iPosBasicInfo + 9)).Select
                .activeCell.FormulaR1C1 = oRS1.Item(1).Item("eqcomments")
                'STEPCOMMENT
                .range("F" & CStr(iPosBasicInfo + 10)).Select
                .activeCell.FormulaR1C1 = oRS1.Item(1).Item("stepcomments")
              
                'Added by Jack on 2014/01/13 for JC201400011 <Start>
                'contactboardtype
                 .range("F" & CStr(iPosBasicInfo + 16)).Select
                 .activeCell.FormulaR1C1 = oRS1.Item(1).Item("contactboardtype")
                'Added by Jack on 2014/01/13 for JC201400011 <End>
            End If
            
            If sFollowOI = "N" And sFollowProd = "N" Then
               sSQL = " SELECT A." & gsCAT_TERE_PGNAME & "  " & _
                   " FROM " & gsCAT_TBL_ERUN_RECIPE & " A " & _
                   " WHERE A." & gsCAT_TERE_DOCNO & " ='" & sERunTicNo & "' " & _
                   " AND A." & gsCAT_TERE_PATH & " ='" & oLot.CustomAttributes(gsLOT_CUSTOMATTR_ROUTE) & "' " & _
                   " AND A." & gsCAT_TERE_STEPNO & " ='" & oLot.CurrentStep.Steps.Item(1).Id & "' " & _
                   " AND A." & gsCAT_TERE_EQTYPE2 & " ='" & oEqp.CustomAttributes(gsEQP_CUSTOMATTR_EQ_TYPE2) & "' " & _
                   " AND A." & gsCAT_TERE_DELETE_FLAG & " ='N' "
                Set oRs2 = moProRawSql.QueryDatabase(sSQL)
                If oRs2.Count > 0 Then
                    'PgName
                    .range("F" & CStr(iPosBasicInfo + 4)).Select
                    .activeCell.FormulaR1C1 = oRs2.Item(1).Item(gsCAT_TERE_PGNAME)
                End If
                'Marked by Jack on 2014/01/13 for JC201400011 <Start>
                'Chihhao要求不清空以下欄位內容.
                'EQCOMMENTS
                '.range("F" & CStr(iPosBasicInfo + 9)).Select
                '.activeCell.FormulaR1C1 = ""
                ''STEPCOMMENT
                '.range("F" & CStr(iPosBasicInfo + 10)).Select
                '.activeCell.FormulaR1C1 = ""
                'Marked by Jack on 2014/01/13 for JC201400011 <End>
            End If
            
        End If
                    
        'taping Qty
        sSQL = " select " & gsCAT_TPBS_CARRIER_QTY & _
               " from " & gsCAT_TBL_PRM_BE_SPEC & _
               " where " & gsCAT_TPBS_IPN & "='" & oLot.CustomAttributes(gsLOT_CUSTOMATTR_IPN) & "'" & _
               " and " & gsCAT_TPBS_DELETE_FLAG & "='N'" & _
               " and " & gsCAT_TPBS_DEFAULTS & "='Y'"
        Set oRs2 = moProRawSql.QueryDatabase(sSQL)
        
        'Add by Tony Start on 2014/09/05 for Req.JC201400285
        If sTapingQty = "NA" Then
            .range("S" & CStr(iPosBasicInfo + 5)).Select
            .activeCell.FormulaR1C1 = sTapingQty
        Else
        'Add by Tony End on 2014/09/05 for Req.JC201400285
            If oRs2.Count > 0 Then
                .range("S" & CStr(iPosBasicInfo + 5)).Select
                .activeCell.FormulaR1C1 = oRs2.Item(1).Item(gsCAT_TPBS_CARRIER_QTY)
            End If
        End If
        
       
        sSQL = " select A." & gsCAT_TEI_CONTACTBOARD_ID & " " & _
               " FROM " & gsCAT_TBL_EQ_INFO & " A " & _
               " WHERE A." & gsCAT_TEI_EQ_ID & "='" & msEqpId & "' "
        Set oRS1 = moProRawSql.QueryDatabase(sSQL)
        If oRS1.Count > 0 Then
            'CONTACTBOARD_ID
            .range("F" & CStr(iPosBasicInfo + 13)).Select
            .activeCell.FormulaR1C1 = oRS1.Item(1).Item(gsCAT_TEI_CONTACTBOARD_ID)
        End If
        
        'added by Jack on 2014/01/13 for JC201400011. <Start>
        If sFollowProd = "N" Then
            sSQL = " SELECT A." & gsCAT_TERE_CONTACTBOARDTYPE & ",  " & _
                                  gsCAT_TERE_PGNAME & " " & _
                " FROM " & gsCAT_TBL_ERUN_RECIPE & " A " & _
                " WHERE A." & gsCAT_TERE_DOCNO & " ='" & sERunTicNo & "' " & _
                " AND A." & gsCAT_TERE_PATH & " ='" & oLot.CustomAttributes(gsLOT_CUSTOMATTR_ROUTE) & "' " & _
                " AND A." & gsCAT_TERE_STEPNO & " ='" & oLot.CurrentStep.Steps.Item(1).Id & "' " & _
                " AND A." & gsCAT_TERE_EQTYPE2 & " ='" & oEqp.CustomAttributes(gsEQP_CUSTOMATTR_EQ_TYPE2) & "' " & _
                " AND A." & gsCAT_TERE_DELETE_FLAG & " ='N' "
             Set oRs2 = moProRawSql.QueryDatabase(sSQL)
             If oRs2.Count > 0 Then
                 'contactboardtype
                 .range("F" & CStr(iPosBasicInfo + 16)).Select
                 .activeCell.FormulaR1C1 = oRs2.Item(1).Item(gsCAT_TERE_CONTACTBOARDTYPE)
                 
                 'PgName
                .range("F" & CStr(iPosBasicInfo + 4)).Select
                .activeCell.FormulaR1C1 = oRs2.Item(1).Item(gsCAT_TERE_PGNAME)
             End If
        End If
        'added by Jack on 2014/01/13 for JC201400011. <End>
        
        'Add by Tony Start on 2014/02/25 for HW SETUP RECIPE COMPARE project
        '架機品質確認項目 (Check items of HW Setup Quality)內容
        
        sEqID = Trim(Me.txtEqId.Text)
        
        Call getRecipeSpecName(sEqID, oLot.Id, oLot.CustomAttributes(gsLOT_CUSTOMATTR_IPN), _
                            msPackageName, msPinCount, msBodySize, _
                            oRS1, moProRawSql, moAppLog, _
                            moCwMbx)
        
        If Not oRS1 Is Nothing And oRS1.Count > 0 Then
            For iIdx = 1 To oRS1.Count
                If iIdx <= miCOMBO_RECIPE_ROW Then
                    .range("B" & CStr(iPosTitle + 28 + iIdx)).Select 'ID
                    .activeCell.FormulaR1C1 = oRS1.Item(iIdx).Item("col1")
                    
                    .range("C" & CStr(iPosTitle + 28 + iIdx)).Select 'Specname
                    .activeCell.FormulaR1C1 = oRS1.Item(iIdx).Item("specname")
                Else
                    If iIdx <= mi_RECIPE_MAX_ROW Then
                        .range("P" & CStr(iPosTitle + 28 + iIdx - miCOMBO_RECIPE_ROW)).Select 'ID
                        .activeCell.FormulaR1C1 = oRS1.Item(iIdx).Item("col1")
                        
                        .range("Q" & CStr(iPosTitle + 28 + iIdx - miCOMBO_RECIPE_ROW)).Select 'Specname
                        .activeCell.FormulaR1C1 = oRS1.Item(iIdx).Item("specname")
                    End If
                End If
            Next
            
        Else
            UtShowMsgBox "尚未維護，請確認架機需求或請維護相關資訊。請洽設備工程師。" & vbCrLf & "No data maintenance, make sure the equipment frame machine or maintain information. Please call HW to check."
        End If
        'Add by Tony End on 2014/02/25 for HW SETUP RECIPE COMPARE project
        
        'Add by Tony Start on 2014/09/02 for Req.JC201400285
        'Handler File Name:
        .range("F" & CStr(iPosBasicInfo + 6)).Select
        .activeCell.FormulaR1C1 = getRecipeSpecName1(sEqID, oLot.Id, oLot.CustomAttributes(gsLOT_CUSTOMATTR_IPN), _
                            msPackageName, msPinCount, msBodySize, _
                            oRS1, moProRawSql, moAppLog, _
                            moCwMbx)
        
        'Add by Tony End on 2014/09/02 for Req.JC201400285
        
        'Add by Ernest on 2015/07/28 for Req.JC201500234---------start
       'taping sealing temp,sealing dwell time,sealing pressure,remark
       sPgm = getRecipeSpecName1(sEqID, oLot.Id, oLot.CustomAttributes(gsLOT_CUSTOMATTR_IPN), _
                            msPackageName, msPinCount, msBodySize, _
                            oRS1, moProRawSql, moAppLog, _
                            moCwMbx)
         sSQL = "select c.sealtemp,c.sealdwell,c.sealpressure,c.remark from " & gsCAT_TBL_LOT_ATTRIBUTE & " a," & _
                               gsCAT_TBL_IPN_MASTER & " b," & _
                               gsCAT_TBL_SCANNER_CONFI & " C" & _
                   " where a." & gsCAT_TLA_LOTID & "='" & oLot.Id & "'" & _
                   " and a." & gsCAT_TLATT_IPN & " =b." & gsCAT_TIM_IPN & _
                   " and b." & gsCAT_TIM_CARRIER_TYPE & "=C." & gsCAT_TSCI_CARRIER_TYPE & "(+)" & _
                   " and b." & gsCAT_TIM_PACKAGE_CODE & "=C." & gsCAT_TSCI_PKG_CODE & "(+)" & _
                   " and b." & gsCAT_TIM_PIN_COUNT & "=C." & gsCAT_TSCI_PIN_CONT & "(+)" & _
                   " and b." & gsCAT_TIM_BODY_SIZE & "=C." & gsCAT_TSCI_BODY_SIZE & "(+)" & _
                   " and b." & gsCAT_TIM_BRAND & "=C." & gsCAT_TSCI_BRAND & "(+)" & _
                   " and c." & gsCAT_TSCI_PGM & "='" & sPgm & "'" & _
                   " and C." & gsCAT_TSCI_EQ_ID & "='" & UCase(Me.txtEqId.Text) & "'" & _
                   " and C." & gsCAT_TSCI_DELETE_FLAG & "='N'"
        
        Set oRs2 = moProRawSql.QueryDatabase(sSQL)
        If oRs2.Count > 0 Then
                .range("S" & CStr(iPosBasicInfo + 6)).Select
                .activeCell.FormulaR1C1 = oRs2.Item(1).Item("sealtemp")
                .range("S" & CStr(iPosBasicInfo + 7)).Select
                .activeCell.FormulaR1C1 = oRs2.Item(1).Item("sealdwell")
                .range("S" & CStr(iPosBasicInfo + 8)).Select
                .activeCell.FormulaR1C1 = oRs2.Item(1).Item("sealpressure")
                .range("S" & CStr(iPosBasicInfo + 9)).Select
                .activeCell.FormulaR1C1 = oRs2.Item(1).Item("remark")
        End If
        'Add by Ernest on 2015/07/28 for Req.JC201500234-----------end
    End With
    Set oRS1 = Nothing
    Set oRs2 = Nothing
    
'----
' Done
'----

ExitHandler:
    ' NOTE 1:
    ' MUST CALL GetErrInfo() here first before another action
    Call GetErrInfo(msMODULE_ID, sProcID, typErrInfo, Erl)
    Call LogProcOut(msMODULE_ID, sProcID, typErrInfo, moAppLog)
    ' <Your cleaning up codes goes here...>
    Set oLot = Nothing
    
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
' Function: GetComboSpec()
'--------------------------------------------------------------------------------
' Description:  <Type your function description here...>
'--------------------------------------------------------------------------------
' Author:       Sam Chen, CIT 2017/03/21
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
'   Boolean          (R) True/False
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
Private Sub GetComboSpec(ByVal oLot As Object, ByVal oEqp As Object, _
                                 ByRef sRtnContactBoard As String, ByRef sRtnPgID As String, _
                                 ByRef sRtnPgName As String, ByRef sRtnPgMode As String, _
                                 ByRef sRtnSubsystem As String, ByRef sRtnTemp As String, ByRef sRtnStopTicNo As String, _
                                 ByRef sRtnSpecifyEQ As String, ByRef sRtnSpecifyEqID As String)
On Error GoTo ExitHandler:
Dim sProcID As String
Dim typErrInfo As tErrInfo
Dim colRS As Collection
Dim sSQL As String

Dim lIdx As Long
Dim sFollowProd As String
Dim sFollowOI As String

Dim sERunTicNo As String
Dim sTecnLotID As String

'----
' Init
'----
    sProcID = "GetComboSpec"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog) '"Entering Function...", moapplog, glLOG_PROC, msMODULE_ID, sProcID)
   
    
'----
' Condition Checking
'----
    ' <Put your condition checking codes here>...

'----
' Action
'----

    sSQL = "select a." & gsCAT_TLI_ERUNTICNO & ",a." & gsCAT_TLI_TECN_LOT_ID & " from " & gsCAT_TBL_LOT_INFO & " a " & _
           " where a." & gsCAT_TLI_LOT_ID & " = '" & oLot.Id & "' "
    Set colRS = moProRawSql.QueryDatabase(sSQL)
    If colRS.Count > 0 Then
        sERunTicNo = colRS.Item(1).Item(gsCAT_TLI_ERUNTICNO)
        sTecnLotID = colRS.Item(1).Item(gsCAT_TLI_TECN_LOT_ID)
    End If

    sSQL = " SELECT A." & gsCAT_TER_FOLLOW_OI & " ,A." & gsCAT_TER_FOLLOW_PRODUCT & " " & _
                   " FROM " & gsCAT_TBL_ERUN_REQ & " A " & _
                   " WHERE A." & gsCAT_TER_TICKET_NO & " ='" & sERunTicNo & "' " & _
                   " AND A." & gsCAT_TER_DELETE_FLAG & " ='N' "
                   
    Set colRS = moProRawSql.QueryDatabase(sSQL)
    If colRS.Count > 0 Then
       sFollowOI = colRS.Item(1).Item(gsCAT_TER_FOLLOW_OI)
       sFollowProd = colRS.Item(1).Item(gsCAT_TER_FOLLOW_PRODUCT)
    End If

    If sFollowOI = "N" And sFollowProd = "N" Then
               sSQL = " SELECT A." & gsCAT_TERE_CONTACTBOARDTYPE & "  " & _
                      " , A." & gsCAT_TERE_PGID & " " & _
                      " , A." & gsCAT_TERE_PGMODE & " " & _
                      " , A." & gsCAT_TERE_PGNAME & "  " & _
                      " , A." & gsCAT_TERE_SUBSYSTEM & "  " & _
                      " , A." & gsCAT_TERE_TEMPERATURE & "  " & _
                   " FROM " & gsCAT_TBL_ERUN_RECIPE & " A " & _
                   " WHERE A." & gsCAT_TERE_DOCNO & " ='" & sERunTicNo & "' " & _
                   " AND A." & gsCAT_TERE_PATH & " ='" & oLot.CustomAttributes(gsLOT_CUSTOMATTR_ROUTE) & "' " & _
                   " AND A." & gsCAT_TERE_STEPNO & " ='" & oLot.CurrentStep.Steps.Item(1).Id & "' " & _
                   " AND A." & gsCAT_TERE_EQTYPE2 & " ='" & oEqp.CustomAttributes(gsEQP_CUSTOMATTR_EQ_TYPE2) & "' " & _
                   " AND A." & gsCAT_TERE_DELETE_FLAG & " ='N' "
                Set colRS = moProRawSql.QueryDatabase(sSQL)
                If colRS.Count > 0 Then
                    sRtnContactBoard = colRS.Item(1).Item(gsCAT_TERE_CONTACTBOARDTYPE)
                    sRtnPgID = colRS.Item(1).Item(gsCAT_TERE_PGID)
                    sRtnPgMode = colRS.Item(1).Item(gsCAT_TERE_PGMODE)
                    sRtnPgName = colRS.Item(1).Item(gsCAT_TERE_PGNAME)
                    sRtnSubsystem = colRS.Item(1).Item(gsCAT_TERE_SUBSYSTEM)
                    sRtnTemp = colRS.Item(1).Item(gsCAT_TERE_TEMPERATURE)
                End If
    Else
        If sTecnLotID <> "" Then
            sSQL = " SELECT  C." & gsCAT_TLSPES_CONTACTBOARDTYPE & "  " & _
                   ", C." & gsCAT_TLSPES_PGID & " " & _
                   ", C." & gsCAT_TLSPES_PGNAME & " " & _
                   ", C." & gsCAT_TLSPES_PGMODE & " " & _
                   ", C." & gsCAT_TLSPES_SUBSYSTEM & " " & _
                   ", C." & gsCAT_TLSPES_TEMPERATURE & " " & _
                   ", C." & gsCAT_TLSPES_SPECIFYEQ & " " & _
                   ", C." & gsCAT_TLSPES_EQID & " " & _
                   " FROM " & gsCAT_TBL_LOT_STEP_SPEC & " B ," & _
                              gsCAT_TBL_LOT_STEP_EQ_SPEC & " C " & _
                   " WHERE C." & gsCAT_TLSPES_TECNLOTID & "=B." & gsCAT_TLSP_TECNLOTID & " and " & _
                 " C." & gsCAT_TLSPES_DOCNO & "=B." & gsCAT_TLSP_DOCNO & " and " & _
                 " C." & gsCAT_TLSPES_PRODGROUP & "=B." & gsCAT_TLSP_PRODGROUP & " and " & _
                 " '" & sTecnLotID & "' like C." & gsCAT_TLSPES_TECNLOTID & " and " & _
                 " C." & gsCAT_TLSPES_PATH & "='" & oLot.CustomAttributes(gsLOT_CUSTOMATTR_ROUTE) & "' and " & _
                 " C." & gsCAT_TLSPES_STEPNO & "='" & oLot.CurrentStep.Steps.Item(1).Id & "' and " & _
                 " C." & gsCAT_TLSPS_DELETEFLAG & "='N' and " & _
                 " C." & gsCAT_TLSPES_EQTYPE2 & "='" & oEqp.CustomAttributes(gsEQP_CUSTOMATTR_EQ_TYPE2) & "' and " & _
                " B." & gsCAT_TLSP_DELETEFLAG & "='N' "
                Set colRS = moProRawSql.QueryDatabase(sSQL)
                If colRS.Count > 0 Then
                    sRtnContactBoard = colRS.Item(1).Item(gsCAT_TLSPES_CONTACTBOARDTYPE)
                    sRtnPgID = colRS.Item(1).Item(gsCAT_TLSPES_PGID)
                    sRtnPgMode = colRS.Item(1).Item(gsCAT_TLSPES_PGMODE)
                    sRtnPgName = colRS.Item(1).Item(gsCAT_TLSPES_PGNAME)
                    sRtnSubsystem = colRS.Item(1).Item(gsCAT_TLSPES_SUBSYSTEM)
                    sRtnTemp = colRS.Item(1).Item(gsCAT_TLSPES_TEMPERATURE)
                    sRtnSpecifyEQ = colRS.Item(1).Item(gsCAT_TLSPES_SPECIFYEQ)
                    sRtnSpecifyEqID = colRS.Item(1).Item(gsCAT_TLSPES_EQID)
                End If
        
        Else
            sSQL = " SELECT  C." & gsCAT_PSES_CONTACTBOARD_TYPE & "  " & _
                   " ,C." & gsCAT_PSES_PG_ID & " " & _
                   " ,C." & gsCAT_PSES_PG_MODE & " " & _
                   " ,C." & gsCAT_PSES_PG_NAME & " " & _
                   " ,C." & gsCAT_PSES_SUB_SYSTEM & " " & _
                   " ,C." & gsCAT_PSES_TEMPERATURE & " " & _
                   " ,C." & gsCAT_PSES_STOPTICNO & " " & _
                   " ,C." & gsCAT_PSES_SpecifyEq & " " & _
                   " ,C." & gsCAT_PSES_EqId & " " & _
                   " FROM " & gsCAT_TBL_PROD_STEP_SPEC & " B ," & _
                   gsCAT_TBL_PROD_STEP_EQ_SPEC & " C " & _
                   " WHERE B." & gsCAT_PSS_PROD_GROUP & "=C." & gsCAT_PSES_PROD_GROUP & "(+) " & _
                   " AND B." & gsCAT_PSS_PATH & "=C." & gsCAT_PSES_PATH & "(+) " & _
                   " AND B." & gsCAT_PSS_STEP_NO & "=C." & gsCAT_PSES_STEP_NO & "(+) " & _
                   " AND B." & gsCAT_PSS_DOC_STATUS & "=C." & gsCAT_PSES_DOC_STATUS & "(+) " & _
                   " AND B." & gsCAT_PSS_DOC_STATUS & "='" & gsDOCSTATUS_ACTIVE & "' " & _
                   " AND B." & gsCAT_PSS_STEP_NO & "='" & oLot.CurrentStep.Steps.Item(1).Id & "' " & _
                   " AND B." & gsCAT_PSS_PROD_GROUP & "='" & oLot.CustomAttributes(gsLOT_CUSTOMATTR_PROD_GROUP) & "' " & _
                   " AND B." & gsCAT_PSS_PATH & "='" & oLot.CustomAttributes(gsLOT_CUSTOMATTR_ROUTE) & "' " & _
                   " AND C." & gsCAT_PSES_EQTYPE2 & " ='" & oEqp.CustomAttributes(gsEQP_CUSTOMATTR_EQ_TYPE2) & "' "
                Set colRS = moProRawSql.QueryDatabase(sSQL)
                If colRS.Count > 0 Then
                    sRtnContactBoard = colRS.Item(1).Item(gsCAT_PSES_CONTACTBOARD_TYPE)
                    sRtnPgID = colRS.Item(1).Item(gsCAT_PSES_PG_ID)
                    sRtnPgMode = colRS.Item(1).Item(gsCAT_PSES_PG_MODE)
                    sRtnPgName = colRS.Item(1).Item(gsCAT_PSES_PG_NAME)
                    sRtnSubsystem = colRS.Item(1).Item(gsCAT_PSES_SUB_SYSTEM)
                    sRtnTemp = colRS.Item(1).Item(gsCAT_PSES_TEMPERATURE)
                    sRtnStopTicNo = colRS.Item(1).Item(gsCAT_PSES_STOPTICNO)
                    sRtnSpecifyEQ = colRS.Item(1).Item(gsCAT_PSES_SpecifyEq)
                    sRtnSpecifyEqID = colRS.Item(1).Item(gsCAT_PSES_EqId)
                End If
                   
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
        Call HandleError(False, typErrInfo, True, moAppLog)
    End If
End Sub


'================================================================================
' Function: GetStopInfoPByAcc()
'--------------------------------------------------------------------------------
' Description:  <Type your Sub description here...>
'--------------------------------------------------------------------------------
' Author:       Sam Chen , CIT 2017/03/17
'--------------------------------------------------------------------------------
' Parameters:
'................................................................................
' ARGUMENTS
'
'
' NON-LOCAL ARGS
'   NonLoc1         (I) <Description goes here...>
'   NonLoc2         (I) <Description goes here...>
'
'--------------------------------------------------------------------------------
' Revision History:
'................................................................................
' [REV 01]
' 1)
'
'================================================================================
Private Function GetStopInfoPByAcc(ByVal sLotID As String, ByVal sProdgroup As String, ByVal sEqType2 As String, _
                                    ByVal sCurEqId As String, ByVal sStepName As String, _
                                    ByVal sPgID As String, sPgName As String, ByVal sPgMode As String, _
                                    ByVal sWsDeviceFile As String, ByVal sCurAccName As String, _
                                    ByRef sStopTicNo As String) As Boolean
On Error GoTo ExitHandler:
Dim sProcID         As String
Dim typErrInfo      As tErrInfo
Dim sSQL            As String
Dim colRS    As Collection

Dim sCarrierType As String
Dim sBodySize As String
Dim sPinCount As String
Dim sPackageType As String


Dim lIdx As Long

'----
' Init
'----
    sProcID = "GetStopInfoPByAcc"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog)
        
    sStopTicNo = ""
    GetStopInfoPByAcc = True
'----
' Condition Checking
'----
    ' <Put your condition checking codes here>...

'----
' Action
'----

    sSQL = " select distinct d." & gsCAT_TSTT_STOPTICNO & ", d." & gsCAT_TSTT_SPECIFYEQ & ", d." & gsCAT_TSTT_EQID & ",d." & gsCAT_TSTT_ACCNAME & _
                    ",d." & gsCAT_TSTT_ACCNO & ",d." & gsCAT_TSTT_DEVICEFILE & ",d." & gsCAT_TSTT_BEPE_SET
                    
    sSQL = sSQL & " from (select aa.productname prodgroup, cc.magnitude prodcode, dd." & gsCAT_TIM_IPN & ", " & _
                        " dd." & gsCAT_TIM_PRODBODY & ", dd." & gsCAT_TIM_PRODGROUPKEY & " " & _
                 "   from fwproductversion       aa," & _
                      "   fwproductversion_n2m   bb," & _
                    "   fwprpattributeinstance cc," & _
                      "   " & gsCAT_TBL_IPN_MASTER & "        dd" & _
                "   where bb.linkname = 'attributes'" & _
                  "   and bb.fromid = aa.sysid" & _
                  "   and bb.toid = cc.sysid" & _
                  "   and bb.keydata = 'TdsProd'" & _
                  "   and aa.revstate = 'Active'" & _
                  "   and aa.productname = dd." & gsCAT_TIM_PROD_GROUP & "(+)) c,"
                  
    sSQL = sSQL & "   (select " & gsCAT_TSTT_STOPTICNO & ", " & gsCAT_TSTT_PRODCODE & ", " & gsCAT_TSTT_IPN & "," & gsCAT_TSTT_STEPNAME & " " & _
                "," & gsCAT_TSTT_EQTYPE2 & ", " & gsCAT_TSTT_SPECIFYEQ & ", " & gsCAT_TSTT_EQID & _
                "," & gsCAT_TSTT_PGMODE & ", " & gsCAT_TSTT_PGID & ", " & gsCAT_TSTT_PGNAME & "," & gsCAT_TSTT_ACCNAME & _
                "," & gsCAT_TSTT_LOTID & ", " & gsCAT_TSTT_PRODGROUPKEY & ", " & gsCAT_TSTT_PRODBODY & _
                "," & gsCAT_TSTT_ACCNO & ", " & gsCAT_TSTT_DEVICEFILE & ", " & gsCAT_TSTT_BEPE_SET & " " & _
                " from " & gsCAT_TBL_STOP_TEST & " " & _
                "   where " & gsCAT_TSTT_DELETEFLAG & " = 'N' " & _
                "    and " & gsCAT_TSTT_BEPE_SET & " = 'A' ) d "

    sSQL = sSQL & " where '" & sEqType2 & "' = d.eqtype2  " & _
      "   and '" & sProdgroup & "'= c.prodgroup" & _
      "   and nvl(c.ipn, ' ') like nvl(REPLACE(REPLACE(d." & gsCAT_TSTT_IPN & ", '%', '_'),'*','%'), nvl(c.ipn, ' '))" & _
      "   and nvl('" & sPgID & "', ' ') = nvl(d." & gsCAT_TSTT_PGID & ", nvl('" & sPgID & "', ' '))" & _
      "   and nvl('" & sPgMode & "', ' ') = nvl(d." & gsCAT_TSTT_PGMODE & ", nvl('" & sPgMode & "', ' '))" & _
      "   and nvl(c.prodcode, ' ') = nvl(d." & gsCAT_TSTT_PRODCODE & ", nvl(c.prodcode, ' '))" & _
      "   and '" & sStepName & "' = nvl(d." & gsCAT_TSTT_STEPNAME & ", '" & sStepName & "')" & _
      "   and ( nvl('" & sCurAccName & "', ' ') = nvl(d.accname, nvl('" & sCurAccName & "', ' ')) or " & _
      "         InStr(d.accname, '" & sCurAccName & "') > 0 ) "
     
    sSQL = sSQL & _
      "   and nvl('" & sLotID & "', ' ') like nvl(REPLACE(REPLACE(d." & gsCAT_TSTT_LOTID & ", '%', '_'),'*','%'), nvl('" & sLotID & "', ' '))" & _
      "   and nvl(c.prodgroupkey, ' ') = nvl(d." & gsCAT_TSTT_PRODGROUPKEY & ", nvl(c.prodgroupkey, ' '))" & _
      "   and nvl(c.ProdBody, ' ') = nvl(d." & gsCAT_TSTT_PRODBODY & ", nvl(c.ProdBody, ' '))" & _
      "   and nvl('" & sWsDeviceFile & "', ' ') = nvl(d." & gsCAT_TSTT_DEVICEFILE & ", nvl('" & sWsDeviceFile & "', ' '))"
    
    sSQL = sSQL & " and d." & gsCAT_TSTT_ACCNO & " like '" & sCurAccName & "%' "
    
    sSQL = sSQL & " order by d." & gsCAT_TSTT_ACCNAME & ", d." & gsCAT_TSTT_SPECIFYEQ & " desc, decode(d." & gsCAT_TSTT_EQID & ", null, 'A', 'B')||" & gsCAT_TSTT_STOPTICNO
    
    Set colRS = moProRawSql.QueryDatabase(sSQL)
    
    For lIdx = 1 To colRS.Count
        If colRS.Item(lIdx).Item(gsCAT_TSTT_SPECIFYEQ) = "Y" Then
            If GetStrPosInArray(colRS.Item(lIdx).Item(gsCAT_TSTT_EQID), sCurEqId, ",") = 0 Then
                If sStopTicNo = "" Then
                    sStopTicNo = colRS.Item(lIdx).Item(gsCAT_TSTT_STOPTICNO)
                Else
                    If GetStrPosInArray(sStopTicNo, colRS.Item(lIdx).Item(gsCAT_TSTT_STOPTICNO), ",") = 0 Then
                        sStopTicNo = sStopTicNo & "," & colRS.Item(lIdx).Item(gsCAT_TSTT_STOPTICNO)
                    End If
                End If
            End If
        ElseIf colRS.Item(lIdx).Item(gsCAT_TSTT_SPECIFYEQ) = "N" Then
            If GetStrPosInArray(colRS.Item(lIdx).Item(gsCAT_TSTT_EQID), sCurEqId, ",") > 0 Then
                If sStopTicNo = "" Then
                    sStopTicNo = colRS.Item(lIdx).Item(gsCAT_TSTT_STOPTICNO)
                Else
                    If GetStrPosInArray(sStopTicNo, colRS.Item(lIdx).Item(gsCAT_TSTT_STOPTICNO), ",") = 0 Then
                        sStopTicNo = sStopTicNo & "," & colRS.Item(lIdx).Item(gsCAT_TSTT_STOPTICNO)
                    End If
                End If
            End If
        Else
            If sStopTicNo = "" Then
                sStopTicNo = colRS.Item(lIdx).Item(gsCAT_TSTT_STOPTICNO)
            Else
                If GetStrPosInArray(sStopTicNo, colRS.Item(lIdx).Item(gsCAT_TSTT_STOPTICNO), ",") = 0 Then
                    sStopTicNo = sStopTicNo & "," & colRS.Item(lIdx).Item(gsCAT_TSTT_STOPTICNO)
                End If
            End If
        End If
    Next
    If sStopTicNo <> "" Then
        GetStopInfoPByAcc = False
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


'================================================================================
' Function: GetStopInfoNPByAcc()
'--------------------------------------------------------------------------------
' Description:  <Type your Sub description here...>
'--------------------------------------------------------------------------------
' Author:       Sam Chen , CIT 2017/03/17
'--------------------------------------------------------------------------------
' Parameters:
'................................................................................
' ARGUMENTS
'
'
' NON-LOCAL ARGS
'   NonLoc1         (I) <Description goes here...>
'   NonLoc2         (I) <Description goes here...>
'
'--------------------------------------------------------------------------------
' Revision History:
'................................................................................
' [REV 01]
' 1)
'
'================================================================================
Private Function GetStopInfoNPByAcc(ByVal sIPN As String, ByVal sEqType2 As String, ByVal sCurEqId As String, _
                                    ByVal sCurAccName As String, ByVal sTemperature As String, _
                                    ByRef sStopTicNo As String, ByRef moFwWIP As Object, ByRef moAppLog As Object, ByRef moProRawSql As Object) As Boolean
On Error GoTo ExitHandler:
Dim sProcID         As String
Dim typErrInfo      As tErrInfo
Dim sSQL            As String
Dim colRS    As Collection

Dim sCarrierType As String
Dim sBodySize As String
Dim sPinCount As String
Dim sPackageType As String

Dim iTemperature As Integer

Dim lIdx As Long

'----
' Init
'----
    sProcID = "GetStopInfoNPByAcc"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog)
    
    If sCurAccName = "" Then    '無配件時不需要再做下去
        GetStopInfoNPByAcc = True
        GoTo ExitHandler
    End If
    
    sSQL = "select a." & gsCAT_TIM_PACKAGE_CODE & ", a." & gsCAT_TIM_BODY_SIZE & " " & _
           " , a." & gsCAT_TIM_PIN_COUNT & ",a." & gsCAT_TIM_CARRIER_TYPE & " " & _
           " from " & gsCAT_TBL_IPN_MASTER & " a " & _
           " where a." & gsCAT_TIM_IPN & " ='" & sIPN & "' "
    Set colRS = moProRawSql.QueryDatabase(sSQL)
    If colRS.Count > 0 Then
        sPackageType = colRS.Item(1).Item(gsCAT_TIM_PACKAGE_CODE)
        sBodySize = colRS.Item(1).Item(gsCAT_TIM_BODY_SIZE)
        sPinCount = colRS.Item(1).Item(gsCAT_TIM_PIN_COUNT)
        sCarrierType = colRS.Item(1).Item(gsCAT_TIM_CARRIER_TYPE)
    End If
    
    If sTemperature = gsDEFAULT_ROOM_TEMP Then
        iTemperature = giDEFAULT_ROOM_TEMP
    Else
        iTemperature = Val(sTemperature)
    End If
        
    sStopTicNo = ""
    GetStopInfoNPByAcc = True
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
          " and nvl('" & sPinCount & "', 0) = nvl(a." & gsCAT_TSTN_PINCOUNT & ", nvl('" & sPinCount & "', 0))" & vbNewLine & _
          " and nvl('" & sCarrierType & "', ' ') = nvl(a." & gsCAT_TSTN_CARRIERTYPE & ", nvl('" & sCarrierType & "', ' '))" & vbNewLine & _
          " and nvl('" & sBodySize & "', ' ') = nvl(a." & gsCAT_TSTN_BODYSIZE & ", nvl('" & sBodySize & "', ' '))" & vbNewLine & _
          " and nvl('" & iTemperature & "', 0) >= nvl(a." & gsCAT_TSTN_TEMP_MIN & ", nvl('" & iTemperature & "', 0))" & vbNewLine & _
          " and nvl('" & iTemperature & "', 0) <= nvl(a." & gsCAT_TSTN_TEMP_MAX & ", nvl('" & iTemperature & "', 0))"
          
    sSQL = sSQL & " and  a." & gsCAT_TSTN_ACCNO & " like '" & sCurAccName & "%'  "
          
    Set colRS = moProRawSql.QueryDatabase(sSQL)
    For lIdx = 1 To colRS.Count
        sStopTicNo = sStopTicNo & "," & colRS.Item(lIdx).Item(gsCAT_TSTN_STOPTICNO)
    Next
    
    If Left(sStopTicNo, 1) = "," Then
        sStopTicNo = Mid(sStopTicNo, 2)
    End If
    
    If sStopTicNo <> "" Then
        GetStopInfoNPByAcc = False
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



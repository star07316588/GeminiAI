VERSION 5.00
Begin VB.Form frmPrintSetupForm 
   BorderStyle     =   1  '單線固定
   Caption         =   "PrintSetupForm"
   ClientHeight    =   5580
   ClientLeft      =   3195
   ClientTop       =   1890
   ClientWidth     =   8880
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   8880
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
      Height          =   405
      Left            =   4680
      TabIndex        =   16
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Frame fraTop 
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
      Height          =   795
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8760
      Begin VB.CommandButton cmdQuery 
         Caption         =   "Query"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   4560
         TabIndex        =   3
         Top             =   240
         Width           =   1455
      End
      Begin VB.TextBox txtLotID 
         Height          =   315
         Left            =   1440
         TabIndex        =   0
         Top             =   300
         Width           =   2500
      End
      Begin VB.Label lblLotID 
         AutoSize        =   -1  'True
         Caption         =   "LotID:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   720
         TabIndex        =   4
         Top             =   360
         Width           =   540
      End
   End
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
      Height          =   405
      Left            =   3120
      TabIndex        =   2
      Top             =   5040
      Width           =   1095
   End
   Begin VB.Frame fraLotInformation 
      Caption         =   "PgID"
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
      Height          =   4095
      Left            =   0
      TabIndex        =   5
      Top             =   840
      Width           =   8775
      Begin VB.TextBox txtStopTicNo 
         Height          =   285
         Left            =   4680
         TabIndex        =   45
         Text            =   "StopTicNo"
         Top             =   2760
         Width           =   975
      End
      Begin VB.TextBox txtSpecifyEqID 
         Height          =   285
         Left            =   3480
         TabIndex        =   44
         Text            =   "SpecifyEQID"
         Top             =   2760
         Width           =   1095
      End
      Begin VB.TextBox txtSpecifyEq 
         Height          =   285
         Left            =   2640
         TabIndex        =   43
         Text            =   "SpecifyEQ"
         Top             =   2760
         Width           =   855
      End
      Begin VB.TextBox txtAccType 
         Height          =   285
         Left            =   6840
         TabIndex        =   42
         Text            =   "AccType"
         Top             =   1560
         Width           =   735
      End
      Begin VB.TextBox txtWsDeviceFile 
         Height          =   285
         Left            =   5400
         TabIndex        =   41
         Text            =   "WsDeviceFile"
         Top             =   1560
         Width           =   1215
      End
      Begin VB.TextBox txtTemp 
         Height          =   285
         Left            =   5760
         TabIndex        =   40
         Text            =   "Temperature"
         Top             =   1920
         Width           =   975
      End
      Begin VB.TextBox txtPgMode 
         Height          =   285
         Left            =   4920
         TabIndex        =   39
         Text            =   "PgMode"
         Top             =   1920
         Width           =   735
      End
      Begin VB.TextBox txtPgID 
         Height          =   285
         Left            =   4320
         TabIndex        =   38
         Text            =   "PgID"
         Top             =   1920
         Width           =   495
      End
      Begin VB.ComboBox cboTesterId 
         Height          =   315
         Left            =   6000
         TabIndex        =   37
         Top             =   600
         Width           =   2535
      End
      Begin VB.ComboBox cboStep 
         Height          =   315
         Left            =   1680
         TabIndex        =   36
         Top             =   600
         Width           =   2535
      End
      Begin VB.TextBox txtInstepTime 
         Height          =   285
         Left            =   6360
         TabIndex        =   35
         Top             =   3720
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox txtStepID 
         Height          =   285
         Left            =   6360
         TabIndex        =   34
         Top             =   3480
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox txtSetupWaferID 
         BackColor       =   &H8000000F&
         Height          =   315
         Index           =   2
         Left            =   6720
         Locked          =   -1  'True
         TabIndex        =   30
         Top             =   3080
         Width           =   2000
      End
      Begin VB.TextBox txtSetupWaferID 
         BackColor       =   &H8000000F&
         Height          =   315
         Index           =   1
         Left            =   3960
         Locked          =   -1  'True
         TabIndex        =   29
         Top             =   3080
         Width           =   2000
      End
      Begin VB.TextBox txtSetupWaferID 
         BackColor       =   &H8000000F&
         Height          =   315
         Index           =   0
         Left            =   1080
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   3080
         Width           =   2000
      End
      Begin VB.ComboBox cboSetupReason 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         ItemData        =   "frmPrintSetupForm.frx":0000
         Left            =   1680
         List            =   "frmPrintSetupForm.frx":0002
         Style           =   2  '單純下拉式
         TabIndex        =   25
         Top             =   3500
         Width           =   3525
      End
      Begin VB.TextBox txtWaferID 
         Height          =   315
         Left            =   1680
         TabIndex        =   24
         Top             =   2280
         Width           =   6825
      End
      Begin VB.ComboBox cboWaferID 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         ItemData        =   "frmPrintSetupForm.frx":0004
         Left            =   1680
         List            =   "frmPrintSetupForm.frx":0006
         Style           =   2  '單純下拉式
         TabIndex        =   22
         Top             =   2280
         Visible         =   0   'False
         Width           =   3525
      End
      Begin VB.TextBox txtPGName 
         Height          =   315
         Left            =   1680
         TabIndex        =   20
         Top             =   1920
         Width           =   2500
      End
      Begin VB.ComboBox cboSubSystem 
         BackColor       =   &H00FFFFFF&
         Height          =   315
         ItemData        =   "frmPrintSetupForm.frx":0008
         Left            =   1680
         List            =   "frmPrintSetupForm.frx":000F
         Style           =   2  '單純下拉式
         TabIndex        =   19
         Top             =   1320
         Width           =   3525
      End
      Begin VB.Label lblSetupWaferID 
         AutoSize        =   -1  'True
         Caption         =   "第三片"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   6120
         TabIndex        =   33
         Top             =   3120
         Width           =   585
      End
      Begin VB.Label lblSetupWaferID 
         AutoSize        =   -1  'True
         Caption         =   "第二片"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   3240
         TabIndex        =   32
         Top             =   3120
         Width           =   585
      End
      Begin VB.Label lblSetupWaferID 
         AutoSize        =   -1  'True
         Caption         =   "第一片"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   31
         Top             =   3120
         Width           =   585
      End
      Begin VB.Label lblSetupWafer 
         AutoSize        =   -1  'True
         Caption         =   "架機片刻號 Setup WaferID"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   28
         Top             =   2760
         Width           =   2295
      End
      Begin VB.Label lblSetupReason 
         AutoSize        =   -1  'True
         Caption         =   "Setup Reason:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   345
         TabIndex        =   26
         Top             =   3500
         Width           =   1260
      End
      Begin VB.Label lblWaferID 
         AutoSize        =   -1  'True
         Caption         =   "WaferID:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   840
         TabIndex        =   23
         Top             =   2400
         Width           =   765
      End
      Begin VB.Label lblPGName 
         AutoSize        =   -1  'True
         Caption         =   "Program Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   345
         TabIndex        =   21
         Top             =   1980
         Width           =   1290
      End
      Begin VB.Label IblLabelSubsystem 
         AutoSize        =   -1  'True
         Caption         =   "SubSystem , Site:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   18
         Top             =   1440
         Width           =   1515
      End
      Begin VB.Label lblLabelStage 
         AutoSize        =   -1  'True
         Caption         =   "Stage:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1080
         TabIndex        =   15
         Top             =   360
         Width           =   570
      End
      Begin VB.Label lblStage 
         BorderStyle     =   1  '單線固定
         Height          =   315
         Left            =   1680
         TabIndex        =   14
         Top             =   240
         Width           =   2505
      End
      Begin VB.Label lblLabelLotStatus 
         AutoSize        =   -1  'True
         Caption         =   "LotStatus:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5040
         TabIndex        =   13
         Top             =   360
         Width           =   885
      End
      Begin VB.Label lblLotStatus 
         BorderStyle     =   1  '單線固定
         Height          =   315
         Left            =   6000
         TabIndex        =   12
         Top             =   240
         Width           =   2505
      End
      Begin VB.Label lblLabelCurrentStep 
         AutoSize        =   -1  'True
         Caption         =   "CurrentStep:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   480
         TabIndex        =   11
         Top             =   720
         Width           =   1080
      End
      Begin VB.Label lblLabelTesterId 
         AutoSize        =   -1  'True
         Caption         =   "(主系統)TesterId:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4440
         TabIndex        =   10
         Top             =   720
         Width           =   1440
      End
      Begin VB.Label lblLabelIPN 
         AutoSize        =   -1  'True
         Caption         =   "IPN:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1200
         TabIndex        =   9
         Top             =   1080
         Width           =   390
      End
      Begin VB.Label lblIPN 
         BorderStyle     =   1  '單線固定
         Height          =   315
         Left            =   1680
         TabIndex        =   8
         Top             =   960
         Width           =   2505
      End
      Begin VB.Label lblLabelLotOwner 
         AutoSize        =   -1  'True
         Caption         =   "LotOwner:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5040
         TabIndex        =   7
         Top             =   1080
         Width           =   885
      End
      Begin VB.Label lblLotOwner 
         BorderStyle     =   1  '單線固定
         Height          =   315
         Left            =   6000
         TabIndex        =   6
         Top             =   960
         Width           =   2505
      End
   End
   Begin VB.Label lblStopInfo 
      Caption         =   "停測資訊"
      Height          =   375
      Left            =   120
      TabIndex        =   46
      Top             =   5040
      Width           =   1935
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "IPN:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   390
   End
End
Attribute VB_Name = "frmPrintSetupForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const msMODULE_ID As String = "frmSetupForm"
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

Private mbNewForm   As Boolean  'add by Sam on 20080826 for PDA_PROJECT

'Add by Tony Start on 2014/03/10 for HW SETUP RECIPE COMPARE project.
Private msPackageName As String
Private msPinCount    As String
Private msBodySize    As String
'add by Ernest on 2016/04/21 for ReqNo:JC201500378------start
Private mbChangestep As Boolean
Private moProduct    As FwProduct
Private mbChangetesterId As Boolean
'add by Ernest on 2016/04/21 for ReqNo:JC201500378---------end

Private msPGMTECNNo  As String 'Added by Jack on 2018/02/02 for Project TECN 自動化 <End>.

Private Const miAT3_RECIPE_ROW  As Integer = 20
Private Const miFT_RECIPE_ROW   As Integer = 10
Private Const miAT3_RECIPE_MAX_ROW As Integer = 40
Private Const miFT_RECIPE_MAX_ROW As Integer = 20

'Add by Tony End on 2014/03/10 for HW SETUP RECIPE COMPARE project.

Private Const msFilePath        As String = "\system32\FwClient\Others\SetupForm.XLT"

Private miResult    As Integer
Private msDocNo     As String
Public Property Let LotID(sLotID As String)
    On Error Resume Next
    msLotId = sLotID
End Property
Public Property Let EqpId(sEqpid As String)
    On Error Resume Next
    msEqpId = sEqpid
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



Private Sub cboStep_Click()
    mbChangestep = True
    Call ExecQuery
End Sub

'================================================================================
' Sub: cboSubSystem_Click()
'--------------------------------------------------------------------------------
' Description:  <Type your Sub description here...>
'--------------------------------------------------------------------------------
' Author:       Sam Chen, CIT 2008/08/26
'--------------------------------------------------------------------------------
' Parameters:
'................................................................................
' ARGUMENTS
'   oLogCtrl            (I) [Optional] Valid reference to a object object
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
'
'================================================================================
Private Sub cboSubSystem_Click()
On Error GoTo ExitHandler:
Dim sProcID As String
Dim typErrInfo As tErrInfo
Dim sFilePath As String

Dim sSwapPgName As String  'Add by Sam on 20150903 for Project 產品變更管理

'Add by Sam Start on 20171120 for Project TECN 自動化
Dim sProdgroupkey As String
Dim sProdCode As String
Dim sGrade As String
Dim sPGMTECNNo As String
Dim sGetPgm As String
Dim sGetPgID As String
Dim sPGMSource As String
Dim sTempTecnNo As String
Dim sGetTemp As String
Dim sTempSource As String
Dim sOverTime As String
Dim colRS As Collection
Dim sDocNos As String
Dim sTecnLevel As String

Dim sEqType2 As String
Dim sStepNo As String
Dim sStepName As String
Dim sIPN As String
Dim sStage As String
Dim sPgId As String
Dim sPgName As String
 'Add by Sam END on 20171120 for Project TECN 自動化

'Add by Sam start on 20200427 for Project FT配件資料標準化
Dim sLoadboard As String
Dim sContactBoard As String
Dim sBodySize As String
Dim colAcc As Collection
'Add by Sam end on 20200427 for Project FT配件資料標準化

'----
' Init
'----
    sProcID = "cboSubSystem_Click"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog)
    Screen.MousePointer = vbHourglass
    
    'Add by Sam start on 20170317 for Project 停復測自動化
    txtPgID = ""
    txtPgMode = ""
    txtTemp = ""
    txtWsDeviceFile = ""
    txtAccType = ""
    txtSpecifyEq = ""
    txtSpecifyEqID = ""
    txtStopTicNo = ""
    'Add by Sam end on 20170317 for Project 停復測自動化
    
    
'----
' Condition Checking
'----

'----
' Action
'----
'Modify by Sam on 20170322 for Project 停復測自動化,取消mbnewForm的條件判斷,因FT也需取得SPEC資訊
'If mbNewForm = True And cboSubSystem.Text <> "" Then
If cboSubSystem.Text <> "" Then
    Dim sSQL            As String
    Dim oRS             As Collection
    Dim oLot        As FwLot
    Dim oEqp    As FwEquipment
    Dim sProdGroup  As String
    Dim sSubSystemType As String
    Dim sMaxsite As String
    Dim sPath  As String
    Dim sWaferID() As String
    Dim iIdx As Integer
    
    Dim bTECNLotSpec            As Boolean  'Add by SAM ON 20100407 FOR Project By Lot TECN
    Dim sTecnLotID            As String   'Add by SAM ON 20100407 FOR Project By Lot TECN
    
    'Added by Jack on 2011/09/23 for JC201100228. 不FollowProduct之工程品 改為 帶出新維修單畫面. <Start>
    Dim sErunTicNO              As String
    Dim sFollowProd             As String
    'Added by Jack on 2011/09/23 for JC201100228. 不FollowProduct之工程品 改為 帶出新維修單畫面. <End>
    
    Me.txtPGName = ""
    
    msPGMTECNNo = "" 'Added by Jack on 2018/02/02 for Project TECN 自動化 <End>.

    Set oLot = FwuRetrieveLot(moFwWIP, txtLotID.Text, moAppLog)
    If oLot Is Nothing Then
       UtShowMsgBox "This Lot was not exist. !!" & vbNewLine & _
                        "此Lot不存在 !!"
        cmdOK.Enabled = False
        GoTo ExitHandler
    End If
    
    'Add by Sam start on 201703022 for Project  停復測自動化,FT時機台取cbotester,增加判斷Stage
    If lblStage = "FT" Then
        Set oEqp = FwuRetrieveEqp(moFwMDL, cboTesterId.Text, moAppLog)
    Else
    'Add by Sam End on 201703022 for Project  停復測自動化,FT時機台取cbotester,增加判斷Stage
        If oLot.CustomAttributes(gsLOT_CUSTOMATTR_CUR_EQP_ID) <> "" Then
            ' Get TesterId Object
            Set oEqp = FwuRetrieveEqp(moFwMDL, oLot.CustomAttributes(gsLOT_CUSTOMATTR_CUR_EQP_ID), moAppLog)
        Else
            Set oEqp = FwuRetrieveEqp(moFwMDL, oLot.CustomAttributes(gsLOT_CUSTOMATTR_SPLITASSIGNEQID), moAppLog)
        End If
    End If
    If oEqp Is Nothing Then
        UtShowMsgBox "The Equipment of this Lot was not exited. !!" & vbNewLine & _
                        "此Lot之測試機台不存在 !!"
        cmdOK.Enabled = False
        GoTo ExitHandler
    End If
    
    If InStr(cboSubSystem.Text, ",") > 0 Then
       sSubSystemType = Left(cboSubSystem.Text, InStr(cboSubSystem.Text, ",") - 1)
       sMaxsite = Mid$(cboSubSystem.Text, InStr(cboSubSystem.Text, ",") + 1, Len(cboSubSystem.Text) - InStr(cboSubSystem.Text, ","))
    Else
       sSubSystemType = cboSubSystem.Text
    End If
    
    sBodySize = "" 'Add by Sam on 20200427 for Project FT配件資料標準化
    'Modify by Sam on 20200427 for Project FT配件資料標準化,Add BodySize
    sSQL = "select " & gsCAT_TIM_PROD_GROUP & "," & gsCAT_TIM_BODY_SIZE & " from " & _
            gsCAT_TBL_IPN_MASTER & " where " & _
            gsCAT_TIM_IPN & "='" & oLot.CustomAttributes(gsLOT_CUSTOMATTR_IPN) & "'"
    Set oRS = moProRawSql.QueryDatabase(sSQL)
    If oRS.Count > 0 Then
        sProdGroup = oRS.Item(1).Item(gsCAT_TIM_PROD_GROUP)
        sBodySize = oRS.Item(1).Item(gsCAT_TIM_BODY_SIZE) 'Add by Sam on 20200427 for Project FT配件資料標準化
    End If
    
    sSQL = "SELECT " & gsCAT_TRSL_PATH & " FROM " & gsCAT_TBL_ROUTE_STEP_LIST & _
            " WHERE " & gsCAT_TRSL_ROUTE_ID & "='" & oLot.PlanId & "' AND " & _
            gsCAT_TRSL_FWSTEPSEQ & "='" & oLot.CurrentStep.Steps.Item(1).Handle & "'"
    Set oRS = moProRawSql.QueryDatabase(sSQL)
    If oRS.Count > 0 Then
        sPath = oRS.Item(1).Item(gsCAT_TRSL_PATH)
    End If
    
    'Added by Jack on 2011/09/23 for JC201100228. 不FollowProduct之工程品 改為 帶出新維修單畫面. <Start>
    sSQL = " SELECT " & gsCAT_TLI_ERUNTICNO & " " & _
           " FROM " & gsCAT_TBL_LOT_INFO & " A " & _
           " WHERE A." & gsCAT_TLI_LOT_ID & " ='" & oLot.Id & "' "
    Set oRS = moProRawSql.QueryDatabase(sSQL)
    If oRS.Count > 0 Then
        sErunTicNO = oRS.Item(1).Item(1)
        If Trim(sErunTicNO) <> "" Then
            'Modified by Jack on 2014/01/13 for JC201400011.
            'Stage='WS' 改為 畫面的Stage. (ori:" and " & gsCAT_TER_STAGE & "='WS' ")
            sSQL = "SELECT " & gsCAT_TER_FOLLOW_PRODUCT & _
                      " from " & gsCAT_TBL_ERUN_REQ & _
                      " where " & gsCAT_TER_LOT_ID & " ='" & oLot.Id & "'" & _
                      " and " & gsCAT_TER_TICKET_NO & "='" & sErunTicNO & "' " & _
                      " and " & gsCAT_TER_STAGE & "='" & lblStage & "' "
            Set oRS = moProRawSql.QueryDatabase(sSQL)
            If oRS.Count > 0 Then
                sFollowProd = oRS.Item(1).Item(1)
            End If
        End If
    End If
    
    'Added by Jack on 2011/09/23 for JC201100228. 不FollowProduct之工程品 改為 帶出新維修單畫面. <End>
    
'Add by Sam start on 20100513 for Project Lot TECN
    bTECNLotSpec = CheckTECNLotInfo(moAppLog, moFwWIP, moFwWF, moCwMbx, oLot.Id)
    
    'Added by Jack on 2011/09/23 for JC201100228. 不FollowProduct之工程品 改為 帶出新維修單畫面. <Start>
    If Trim(sErunTicNO) <> "" And Trim(sFollowProd) = "N" Then
        'Modify by Sam on 20170317 for Project 停復測自動化 ,增加欄位
'        sSql = "select pgname " & _
'                 "from tbl_erun_recipe a " & _
'                "where a.docno = '" & sErunTicNo & "' " & _
'                  "and a.stepno = '" & olot.CurrentStep.Steps.Item(1).Id & "' " & _
'                  "and a.eqtype2 = '" & oEqp.CustomAttributes(gsEQP_CUSTOMATTR_EQ_TYPE2) & "' "
        sSQL = "select " & gsCAT_TERE_PGNAME & " ," & gsCAT_TERE_PGID & " ," & gsCAT_TERE_PGMODE & "  " & _
                 " ," & gsCAT_TERE_TEMPERATURE & " ," & gsCAT_TERE_PROBECARDTYPE & " ," & gsCAT_TERE_LOADBOARDTYPE & "  " & _
                 " ," & gsCAT_TERE_CONTACTBOARDTYPE & " ," & gsCAT_TERE_WSDEVICEFILE & "   " & _
                 "from " & gsCAT_TBL_ERUN_RECIPE & " a " & _
                "where a." & gsCAT_TERE_DOCNO & " = '" & sErunTicNO & "' " & _
                  "and a." & gsCAT_TERE_STEPNO & " = '" & oLot.CurrentStep.Steps.Item(1).Id & "' " & _
                  "and a." & gsCAT_TERE_EQTYPE2 & " = '" & oEqp.CustomAttributes(gsEQP_CUSTOMATTR_EQ_TYPE2) & "' "
                  
        If Trim(sSubSystemType) = "" Then
            sSQL = sSQL & " and a.subsystem is null "
        Else
            sSQL = sSQL & " and a.subsystem='" & Trim(sSubSystemType) & "' "
        End If
        sSQL = sSQL & " and a.deleteflag = 'N' "
                  
        Set oRS = moProRawSql.QueryDatabase(sSQL)
        If oRS.Count > 0 Then
            txtPGName = oRS.Item(1).Item(gsCAT_TERE_PGNAME)
            'Add by Sam start on 20170317 for Project 停復測自動化
            txtPgID = oRS.Item(1).Item(gsCAT_TERE_PGID)
            txtPgMode = oRS.Item(1).Item(gsCAT_TERE_PGMODE)
            txtTemp = oRS.Item(1).Item(gsCAT_TERE_TEMPERATURE)
            txtWsDeviceFile = oRS.Item(1).Item(gsCAT_TERE_WSDEVICEFILE)
            If oRS.Item(1).Item(gsCAT_TERE_PROBECARDTYPE) <> "" Then
                txtAccType = oRS.Item(1).Item(gsCAT_TERE_PROBECARDTYPE)
            ElseIf oRS.Item(1).Item(gsCAT_TERE_LOADBOARDTYPE) <> "" Then
                txtAccType = oRS.Item(1).Item(gsCAT_TERE_LOADBOARDTYPE)
            ElseIf oRS.Item(1).Item(gsCAT_TERE_CONTACTBOARDTYPE) <> "" Then
                txtAccType = oRS.Item(1).Item(gsCAT_TERE_CONTACTBOARDTYPE)
            End If
            'Add by Sam end on 20170317 for Project 停復測自動化
        End If
        
    'Added by Jack on 2011/09/23 for JC201100228. 不FollowProduct之工程品 改為 帶出新維修單畫面. <End>
    ElseIf bTECNLotSpec = True Then
        sSQL = " SELECT A." & gsCAT_TLI_TECN_LOT_ID & " " & _
               " FROM " & gsCAT_TBL_LOT_INFO & " A " & _
               " WHERE A." & gsCAT_TLI_LOT_ID & " ='" & oLot.Id & "' "
     
        Set oRS = moProRawSql.QueryDatabase(sSQL)
        If oRS.Count > 0 Then
           sTecnLotID = oRS.Item(1).Item(1)
        End If
        'Modify by Sam start on 20170317 for Project 停復測自動化,增加欄位
'        sSql = "select " & gsCAT_TLSPES_PGNAME & "," & _
'                        gsCAT_TLSPES_REFPGNAME01 & "," & _
'                        gsCAT_TLSPES_REFPGNAME02 & "," & _
'                        gsCAT_TLSPES_REFPGNAME03 & "," & _
'                        gsCAT_TLSPES_REFSTEPNAME01 & "," & _
'                        gsCAT_TLSPES_REFSTEPNAME02 & "," & _
'                        gsCAT_TLSPES_REFSTEPNAME03 & "," & _
'                        gsCAT_TLSPES_REPLACEPGNAME01 & "," & _
'                        gsCAT_TLSPES_REPLACEPGNAME02 & "," & _
'                        gsCAT_TLSPES_REPLACEPGNAME03 & " " & _
'                        "from " & gsCAT_TBL_LOT_STEP_EQ_SPEC & " where " & _
'                        " '" & sTECNLotID & "' like " & gsCAT_TLSPES_TECNLOTID & " and " & _
'                        gsCAT_TLSPES_STEPNO & "='" & olot.CurrentStep.Steps.Item(1).Id & "' and " & _
'                        gsCAT_TLSPES_EQTYPE2 & "='" & oEqp.CustomAttributes(gsEQP_CUSTOMATTR_EQ_TYPE2) & "' and " & _
'                        gsCAT_TLSPES_DELETEFLAG & "='N' and " & gsCAT_TLSPES_SUBSYSTEM & " ='" & sSubSystemType & "' and " & _
'                        gsCAT_TLSPES_PATH & "='" & sPath & "' and nvl(" & gsCAT_TLSPES_MAXSITE & ",' ') = nvl('" & sMaxsite & "',' ')  "
        sSQL = "select " & gsCAT_TLSPES_PGNAME & "," & _
                        gsCAT_TLSPES_REFPGNAME01 & "," & _
                        gsCAT_TLSPES_REFPGNAME02 & "," & _
                        gsCAT_TLSPES_REFPGNAME03 & "," & _
                        gsCAT_TLSPES_REFSTEPNAME01 & "," & _
                        gsCAT_TLSPES_REFSTEPNAME02 & "," & _
                        gsCAT_TLSPES_REFSTEPNAME03 & "," & _
                        gsCAT_TLSPES_REPLACEPGNAME01 & "," & _
                        gsCAT_TLSPES_REPLACEPGNAME02 & "," & _
                        gsCAT_TLSPES_REPLACEPGNAME03 & ", " & _
                        gsCAT_TLSPES_PGID & "," & _
                        gsCAT_TLSPES_PGMODE & "," & _
                        gsCAT_TLSPES_TEMPERATURE & "," & _
                        gsCAT_TLSPES_PROBECARDTYPE & "," & _
                        gsCAT_TLSPES_LOADBOARDTYPE & ", " & _
                        gsCAT_TLSPES_CONTACTBOARDTYPE & ", " & _
                        gsCAT_TLSPES_SPECIFYEQ & ", " & _
                        gsCAT_TLSPES_EQID & ", " & _
                        gsCAT_TLSPES_WSDEVICEFILE & " "

        sSQL = sSQL & " , " & gsCAT_TLSPES_BURN_IN_BOARD & " " 'Add by Sam on 20200430 for PROJECT FT配件資料標準化

        sSQL = sSQL & " from " & gsCAT_TBL_LOT_STEP_EQ_SPEC & " where " & _
                        " '" & sTecnLotID & "' like " & gsCAT_TLSPES_TECNLOTID & " and " & _
                        gsCAT_TLSPES_STEPNO & "='" & oLot.CurrentStep.Steps.Item(1).Id & "' and " & _
                        gsCAT_TLSPES_EQTYPE2 & "='" & oEqp.CustomAttributes(gsEQP_CUSTOMATTR_EQ_TYPE2) & "' and " & _
                        gsCAT_TLSPES_DELETEFLAG & "='N' and " & gsCAT_TLSPES_SUBSYSTEM & " ='" & sSubSystemType & "' and " & _
                        gsCAT_TLSPES_PATH & "='" & sPath & "' and nvl(" & gsCAT_TLSPES_MAXSITE & ",' ') = nvl('" & sMaxsite & "',' ')  "
                        
        'Modify by Sam End on 20170317 for Project 停復測自動化,增加欄位
                        
        Set oRS = moProRawSql.QueryDatabase(sSQL)
        If oRS.Count > 0 Then
            txtPGName = GetPGM(oLot, oRS.Item(1).Item(gsCAT_TLSPES_PGNAME), _
                               oRS.Item(1).Item(gsCAT_TLSPES_REFSTEPNAME01), _
                               oRS.Item(1).Item(gsCAT_TLSPES_REFPGNAME01), _
                               oRS.Item(1).Item(gsCAT_TLSPES_REPLACEPGNAME01), _
                               oRS.Item(1).Item(gsCAT_TLSPES_REFSTEPNAME02), _
                               oRS.Item(1).Item(gsCAT_TLSPES_REFPGNAME02), _
                               oRS.Item(1).Item(gsCAT_TLSPES_REPLACEPGNAME02), _
                               oRS.Item(1).Item(gsCAT_TLSPES_REFSTEPNAME03), _
                               oRS.Item(1).Item(gsCAT_TLSPES_REFPGNAME03), _
                               oRS.Item(1).Item(gsCAT_TLSPES_REPLACEPGNAME03), _
                               oEqp.CustomAttributes(gsEQP_CUSTOMATTR_EQ_TYPE2), sSubSystemType)
            'Add by Sam start on 20170317 for Project 停復測自動化
            txtPgID = oRS.Item(1).Item(gsCAT_TLSPES_PGID)
            txtPgMode = oRS.Item(1).Item(gsCAT_TLSPES_PGMODE)
            txtTemp = oRS.Item(1).Item(gsCAT_TLSPES_TEMPERATURE)
            txtWsDeviceFile = oRS.Item(1).Item(gsCAT_TLSPES_WSDEVICEFILE)
            
            'Add by Sam Start on 20200427 for Project FT配件資料標準化
            sLoadboard = oRS.Item(1).Item(gsCAT_TLSPES_LOADBOARDTYPE)
            sContactBoard = oRS.Item(1).Item(gsCAT_TLSPES_CONTACTBOARDTYPE)
            If Trim(oRS.Item(1).Item(gsCAT_TLSPES_LOADBOARDTYPE)) = "" And _
                Trim(oRS.Item(1).Item(gsCAT_TLSPES_CONTACTBOARDTYPE)) = "" And _
                Trim(oRS.Item(1).Item(gsCAT_TLSPES_BURN_IN_BOARD)) = "" Then
                Set colAcc = GetFTAccByPgm(moProRawSql, moAppLog, sStepName, sBodySize, oRS.Item(1).Item(gsCAT_TLSPES_PGID), oRS.Item(1).Item(gsCAT_TLSPES_PGNAME))
                If colAcc.Count > 0 Then
                    sLoadboard = colAcc.Item("loadboard")
                    sContactBoard = colAcc.Item("contactboard")
                End If
            End If
            'Add by Sam End on 20200427 for Project FT配件資料標準化
            
            If oRS.Item(1).Item(gsCAT_TLSPES_PROBECARDTYPE) <> "" Then
                txtAccType = oRS.Item(1).Item(gsCAT_TLSPES_PROBECARDTYPE)
            'Modify by Sam start on 20200427 for Project FT配件資料標準化
'            ElseIf oRS.Item(1).Item(gsCAT_TLSPES_LOADBOARDTYPE) <> "" Then
'                txtAccType = oRS.Item(1).Item(gsCAT_TLSPES_LOADBOARDTYPE)
'            ElseIf oRS.Item(1).Item(gsCAT_TLSPES_CONTACTBOARDTYPE) <> "" Then
'                txtAccType = oRS.Item(1).Item(gsCAT_TLSPES_CONTACTBOARDTYPE)
'            End If
            ElseIf sLoadboard <> "" Then
                txtAccType = sLoadboard
            ElseIf sContactBoard <> "" Then
                txtAccType = sContactBoard
            End If
            'Modify by Sam End on 20200427 for Project FT配件資料標準化
            txtSpecifyEq = oRS.Item(1).Item(gsCAT_TLSPES_SPECIFYEQ)
            txtSpecifyEqID = oRS.Item(1).Item(gsCAT_TLSPES_EQID)
            
            'Add by Sam end on 20170317 for Project 停復測自動化
                                                              
            'Add by Sam start on 20171120 for Project TECN 自動化
            sProdgroupkey = ""
            sProdCode = ""
            sGrade = ""
            sPGMTECNNo = ""
            sGetPgm = ""
            sGetPgID = ""
            sPGMSource = ""
            sTempTecnNo = ""
            sGetTemp = ""
            sTempSource = ""
            sOverTime = ""
            sTecnLevel = ""
            sIPN = oLot.CustomAttributes(gsLOT_CUSTOMATTR_IPN)
            sStepNo = oLot.CurrentStep.Steps.Item(1).Id
            sStepName = oLot.CurrentStep.Steps.Item(1).Description
            sEqType2 = oEqp.CustomAttributes(gsEQP_CUSTOMATTR_EQ_TYPE2)
            sStage = oLot.CurrentStep.Steps.Item(1).Attributes(gsSTEP_CUSTOMATTR_STAGE)
            sPgId = txtPgID.Text
            sPgName = txtPGName.Text
            sSQL = "select A." & gsCAT_TIM_PRODGROUPKEY & ", a." & gsCAT_TIM_MASK_OPTION & ", a." & gsCAT_TIM_BE_OPTION & " ,a." & gsCAT_TIM_GRADE & " from " & gsCAT_TBL_IPN_MASTER & " a " & _
                  " where a." & gsCAT_TIM_IPN & " = '" & sIPN & "' "
            Set colRS = moProRawSql.QueryDatabase(sSQL)
            If colRS.Count > 0 Then
                sProdgroupkey = colRS.Item(1).Item(gsCAT_TIM_PRODGROUPKEY)
                sGrade = colRS.Item(1).Item(gsCAT_TIM_GRADE)
                If sStage = "WS" Then
                    sGrade = "C"
                    sProdCode = Left(sIPN, 4) & colRS.Item(1).Item(gsCAT_TIM_MASK_OPTION)
                ElseIf sStage = "FT" Then
                    sProdCode = Left(sIPN, 4) & colRS.Item(1).Item(gsCAT_TIM_BE_OPTION)
                End If
            End If
            'Modify by Sam on 20180417 for Project TECN 自動化,加傳PGID , PGNAME
            Call modTecn.GetTecnPgmRecipeAttr(moAppLog, moProRawSql, "", oLot.Id, "", _
                                          sPath, sStepNo, sStepName, _
                                          sEqType2, "", "", sGrade, "Y", sPgId, sPgName, _
                                          sPGMTECNNo, sGetPgm, sGetPgID, sPGMSource, _
                                          sTempTecnNo, sGetTemp, sTempSource, sOverTime, sTecnLevel)
                                          
            If sGetPgm <> "" Then
                'Mark by Sam Start on 20180417 for Project TECN 自動化,比對功能移至GetTecnPgmRecipeAttr內
'                If ComparePgmAttr(moAppLog, moProRawSql, _
'                                 sStage, sStepName, sEqType2, sPgId, sPgName, _
'                                 sGetPgID, sGetPgm) <> "PASS" Then
'                    sGetPgm = ""
'                    sGetPgID = ""
'                    sGetTemp = ""
'                    sOverTime = ""
'                End If
                'Mark by Sam eND on 20180417 for Project TECN 自動化,比對功能移至GetTecnPgmRecipeAttr內
                If sGetPgm <> "" Then
                    txtPGName = sGetPgm
                    txtPgID = sGetPgID
                End If
                If sGetTemp <> "" Then
                    txtTemp = sGetTemp
                End If
                
                msPGMTECNNo = sPGMTECNNo 'Added by Jack on 2018/02/02 for Project TECN 自動化.

            End If
            'Add by Sam End on 20171120 for Project TECN 自動化
                               
    '        txtPgName = oRS.Item(1).Item(gsCAT_PSES_PG_NAME)
        End If
    Else
'Add by Sam End on 20100513 for Project Lot TECN
    
        'Modify by Sam start on 20170317 for Project 停復測自動化,增加欄位
'        sSql = "select " & gsCAT_PSES_PG_NAME & "," & _
'                        gsCAT_PSES_REF_PG_NAME_01 & "," & _
'                        gsCAT_PSES_REF_PG_NAME_02 & "," & _
'                        gsCAT_PSES_REF_PG_NAME_03 & "," & _
'                        gsCAT_PSES_REF_STEP_NAME_01 & "," & _
'                        gsCAT_PSES_REF_STEP_NAME_02 & "," & _
'                        gsCAT_PSES_REF_STEP_NAME_03 & "," & _
'                        gsCAT_PSES_REPLACE_PG_NAME_01 & "," & _
'                        gsCAT_PSES_REPLACE_PG_NAME_02 & "," & _
'                        gsCAT_PSES_REPLACE_PG_NAME_03 & " " & _
'                        "from " & gsCAT_TBL_PROD_STEP_EQ_SPEC & " where " & _
'                        gsCAT_PSES_PROD_GROUP & "='" & sProdGroup & "' and " & _
'                        gsCAT_PSES_STEP_NO & "='" & olot.CurrentStep.Steps.Item(1).Id & "' and " & _
'                        gsCAT_PSES_EQTYPE2 & "='" & oEqp.CustomAttributes(gsEQP_CUSTOMATTR_EQ_TYPE2) & "' and " & _
'                        gsCAT_PSES_DOC_STATUS & "='Active' and " & gsCAT_PSES_SUB_SYSTEM & " ='" & sSubSystemType & "' and " & _
'                        gsCAT_PSES_PATH & "='" & sPath & "' and nvl(" & gsCAT_PSES_MAX_SITE & ",' ') = nvl('" & sMaxsite & "',' ')  "
        sSQL = "select " & gsCAT_PSES_PG_NAME & "," & _
                        gsCAT_PSES_REF_PG_NAME_01 & "," & _
                        gsCAT_PSES_REF_PG_NAME_02 & "," & _
                        gsCAT_PSES_REF_PG_NAME_03 & "," & _
                        gsCAT_PSES_REF_STEP_NAME_01 & "," & _
                        gsCAT_PSES_REF_STEP_NAME_02 & "," & _
                        gsCAT_PSES_REF_STEP_NAME_03 & "," & _
                        gsCAT_PSES_REPLACE_PG_NAME_01 & "," & _
                        gsCAT_PSES_REPLACE_PG_NAME_02 & "," & _
                        gsCAT_PSES_REPLACE_PG_NAME_03 & ", " & _
                        gsCAT_PSES_PG_ID & "," & _
                        gsCAT_PSES_PG_MODE & "," & _
                        gsCAT_PSES_TEMPERATURE & "," & _
                        gsCAT_PSES_PROBECARD_TYPE & "," & _
                        gsCAT_PSES_LOADBOARD_TYPE & ", " & _
                        gsCAT_PSES_CONTACTBOARD_TYPE & ", " & _
                        gsCAT_PSES_SpecifyEq & ", " & _
                        gsCAT_PSES_EqId & ", " & _
                        gsCAT_PSES_STOPTICNO & ", " & _
                        gsCAT_PSES_WS_DEVICE_FILE & " "
                        
        sSQL = sSQL & " , " & gsCAT_PSES_BURN_IN_BOARD & " " 'Add by Sam on 20200430 for PROJECT FT配件資料標準化

        sSQL = sSQL & " from " & gsCAT_TBL_PROD_STEP_EQ_SPEC & " where " & _
                        gsCAT_PSES_PROD_GROUP & "='" & sProdGroup & "' and " & _
                        gsCAT_PSES_STEP_NO & "='" & oLot.CurrentStep.Steps.Item(1).Id & "' and " & _
                        gsCAT_PSES_EQTYPE2 & "='" & oEqp.CustomAttributes(gsEQP_CUSTOMATTR_EQ_TYPE2) & "' and " & _
                        gsCAT_PSES_DOC_STATUS & "='Active' and " & gsCAT_PSES_SUB_SYSTEM & " ='" & sSubSystemType & "' and " & _
                        gsCAT_PSES_PATH & "='" & sPath & "' and nvl(" & gsCAT_PSES_MAX_SITE & ",' ') = nvl('" & sMaxsite & "',' ')  "
                
        'Modify by Sam end on 20170317 for Project 停復測自動化,增加欄位
                
        Set oRS = moProRawSql.QueryDatabase(sSQL)
        If oRS.Count > 0 Then
            'Mark by Sam Start on 20150903 for Project 產品變更管理,SWAP PGNAME改用新制
'            txtPGName = GetPGM(oLot, oRS.Item(1).Item(gsCAT_PSES_PG_NAME), _
'                               oRS.Item(1).Item(gsCAT_PSES_REF_STEP_NAME_01), _
'                               oRS.Item(1).Item(gsCAT_PSES_REF_PG_NAME_01), _
'                               oRS.Item(1).Item(gsCAT_PSES_REPLACE_PG_NAME_01), _
'                               oRS.Item(1).Item(gsCAT_PSES_REF_STEP_NAME_02), _
'                               oRS.Item(1).Item(gsCAT_PSES_REF_PG_NAME_02), _
'                               oRS.Item(1).Item(gsCAT_PSES_REPLACE_PG_NAME_02), _
'                               oRS.Item(1).Item(gsCAT_PSES_REF_STEP_NAME_03), _
'                               oRS.Item(1).Item(gsCAT_PSES_REF_PG_NAME_03), _
'                               oRS.Item(1).Item(gsCAT_PSES_REPLACE_PG_NAME_03), _
'                               oEqp.CustomAttributes(gsEQP_CUSTOMATTR_EQ_TYPE2), sSubSystemType)
            'Mark by Sam End on 20150903 for Project 產品變更管理,SWAP PGNAME改用新制
    '        txtPgName = oRS.Item(1).Item(gsCAT_PSES_PG_NAME)
            'Add by Sam Start on 20150903 for Project 產品變更管理
            sSwapPgName = GetSwapPGName(moProRawSql, moAppLog, oLot.Id, sProdGroup, oLot.CurrentStep.Steps.Item(1).Description _
                                    , oEqp.CustomAttributes(gsEQP_CUSTOMATTR_EQ_TYPE2), oRS.Item(1).Item(gsCAT_PSES_PG_NAME), False)
            txtPGName = oRS.Item(1).Item(gsCAT_PSES_PG_NAME)
            If sSwapPgName <> "" Then
                If sSwapPgName <> "X" Then
                    txtPGName = sSwapPgName
                End If
            End If
            'Mark by Sam End on 20150903 for Project 產品變更管理,SWAP PGNAME改用新制
            
            'Add by Sam start on 20170317 for Project 停復測自動化
            txtPgID = oRS.Item(1).Item(gsCAT_PSES_PG_ID)
            txtPgMode = oRS.Item(1).Item(gsCAT_PSES_PG_MODE)
            txtTemp = oRS.Item(1).Item(gsCAT_PSES_TEMPERATURE)
            txtWsDeviceFile = oRS.Item(1).Item(gsCAT_PSES_WS_DEVICE_FILE)
            
            'Add by Sam Start on 20200427 for Project FT配件資料標準化
            sLoadboard = oRS.Item(1).Item(gsCAT_PSES_LOADBOARD_TYPE)
            sContactBoard = oRS.Item(1).Item(gsCAT_PSES_CONTACTBOARD_TYPE)
            If Trim(oRS.Item(1).Item(gsCAT_PSES_LOADBOARD_TYPE)) = "" And _
                Trim(oRS.Item(1).Item(gsCAT_PSES_CONTACTBOARD_TYPE)) = "" And _
                Trim(oRS.Item(1).Item(gsCAT_PSES_BURN_IN_BOARD)) = "" Then
                Set colAcc = GetFTAccByPgm(moProRawSql, moAppLog, sStepName, sBodySize, oRS.Item(1).Item(gsCAT_PSES_PG_ID), oRS.Item(1).Item(gsCAT_PSES_PG_NAME))
                If colAcc.Count > 0 Then
                    sLoadboard = colAcc.Item("loadboard")
                    sContactBoard = colAcc.Item("contactboard")
                End If
            End If
            'Add by Sam End on 20200427 for Project FT配件資料標準化
            
            If oRS.Item(1).Item(gsCAT_PSES_PROBECARD_TYPE) <> "" Then
                txtAccType = oRS.Item(1).Item(gsCAT_PSES_PROBECARD_TYPE)
            'Modify by Sam start on 20200427 for Project FT配件資料標準化
'            ElseIf oRS.Item(1).Item(gsCAT_PSES_LOADBOARD_TYPE) <> "" Then
'                txtAccType = oRS.Item(1).Item(gsCAT_PSES_LOADBOARD_TYPE)
'            ElseIf oRS.Item(1).Item(gsCAT_PSES_CONTACTBOARD_TYPE) <> "" Then
'                txtAccType = oRS.Item(1).Item(gsCAT_PSES_CONTACTBOARD_TYPE)
'            End If
            ElseIf sLoadboard <> "" Then
                txtAccType = sLoadboard
            ElseIf sContactBoard <> "" Then
                txtAccType = sContactBoard
            End If
            'Modify by Sam End on 20200427 for Project FT配件資料標準化
            
            txtSpecifyEq = oRS.Item(1).Item(gsCAT_PSES_SpecifyEq)
            txtSpecifyEqID = oRS.Item(1).Item(gsCAT_PSES_EqId)
            txtStopTicNo = oRS.Item(1).Item(gsCAT_PSES_STOPTICNO)
            'Add by Sam end on 20170317 for Project 停復測自動化
            
            'Add by Sam start on 20171120 for Project TECN 自動化
            sProdgroupkey = ""
            sProdCode = ""
            sGrade = ""
            sPGMTECNNo = ""
            sGetPgm = ""
            sGetPgID = ""
            sPGMSource = ""
            sTempTecnNo = ""
            sGetTemp = ""
            sTempSource = ""
            sOverTime = ""
            sTecnLevel = ""
            sIPN = oLot.CustomAttributes(gsLOT_CUSTOMATTR_IPN)
            sEqType2 = oEqp.CustomAttributes(gsEQP_CUSTOMATTR_EQ_TYPE2)
            sStage = oLot.CurrentStep.Steps.Item(1).Attributes(gsSTEP_CUSTOMATTR_STAGE)
            sStepNo = oLot.CurrentStep.Steps.Item(1).Id
            sStepName = oLot.CurrentStep.Steps.Item(1).Description
            sPgId = txtPgID.Text
            sPgName = txtPGName.Text
            sSQL = "select A." & gsCAT_TIM_PRODGROUPKEY & ", a." & gsCAT_TIM_MASK_OPTION & ", a." & gsCAT_TIM_BE_OPTION & " ,a." & gsCAT_TIM_GRADE & " from " & gsCAT_TBL_IPN_MASTER & " a " & _
                  " where a." & gsCAT_TIM_IPN & " = '" & sIPN & "' "
            Set colRS = moProRawSql.QueryDatabase(sSQL)
            If colRS.Count > 0 Then
                sProdgroupkey = colRS.Item(1).Item(gsCAT_TIM_PRODGROUPKEY)
                sGrade = colRS.Item(1).Item(gsCAT_TIM_GRADE)
                If sStage = "WS" Then
                    sGrade = "C"
                    sProdCode = Left(sIPN, 4) & colRS.Item(1).Item(gsCAT_TIM_MASK_OPTION)
                ElseIf sStage = "FT" Then
                    sProdCode = Left(sIPN, 4) & colRS.Item(1).Item(gsCAT_TIM_BE_OPTION)
                End If
            End If
            'Modify by Sam on 20180417 for Project TECN 自動化,加傳pgid pgname
            Call modTecn.GetTecnPgmRecipeAttr(moAppLog, moProRawSql, "", oLot.Id, sIPN, _
                                          sPath, sStepNo, sStepName, _
                                          sEqType2, sProdgroupkey, sProdCode, sGrade, "Y", sPgId, sPgName, _
                                          sPGMTECNNo, sGetPgm, sGetPgID, sPGMSource, _
                                          sTempTecnNo, sGetTemp, sTempSource, sOverTime, sTecnLevel)
            
            If sGetPgm <> "" Then
                'Mark by Sam Start on 20180417 for Project TECN 自動化,比對功能移至GetTecnPgmRecipeAttr內
'                If ComparePgmAttr(moAppLog, moProRawSql, _
'                                 sStage, sStepName, sEqType2, sPgId, sPgName, _
'                                 sGetPgID, sGetPgm) <> "PASS" Then
'                    sGetPgm = ""
'                    sGetPgID = ""
'                    sGetTemp = ""
'                    sOverTime = ""
'                End If
                'Mark by Sam eND on 20180417 for Project TECN 自動化,比對功能移至GetTecnPgmRecipeAttr內
                If sGetPgm <> "" Then
                    txtPGName = sGetPgm
                    txtPgID = sGetPgID
                End If
                If sGetTemp <> "" Then
                    txtTemp = sGetTemp
                End If
                
                msPGMTECNNo = sPGMTECNNo 'Added by Jack on 2018/02/02 for Project TECN 自動化.
                
            End If
            'Add by Sam End on 20171120 for Project TECN 自動化
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
        On Error GoTo ExitHandler:
        Call HandleError(False, typErrInfo, , moAppLog, True)
    End If

End Sub

'add by ernest on 2016/04/21 for ReqNo:JC201500378---start
Private Sub cboTesterId_Click()
    mbChangetesterId = True
    mbChangestep = True
    'Call ExecQuery 'Mark by Sam End on 20190312 for Project FT ADVAN Tester,應User要求再改回直接取Assign機台,Mark此段否則會無限Loop
End Sub
'add by ernest on 2016/04/21 for ReqNo:JC201500378---end

Private Sub cmdCancel_Click()
    On Error Resume Next
    miResult = vbCancel
    Me.Hide
End Sub
'================================================================================
' Sub: cmdOk_Click()
'--------------------------------------------------------------------------------
' Description:  <Type your Sub description here...>
'--------------------------------------------------------------------------------
' Author:       Jeff Weng, MXIC 2005/03/30
'--------------------------------------------------------------------------------
' Parameters:
'................................................................................
' ARGUMENTS
'   oLogCtrl            (I) [Optional] Valid reference to a object object
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
'
'================================================================================
Private Sub cmdOK_Click()
On Error GoTo ExitHandler:
Dim sProcID As String
Dim typErrInfo As tErrInfo
Dim sFilePath As String

Dim oLot        As FwLot
Dim oEqp As FwEquipment
Dim sEqType2 As String
Dim sSQL As String 'add by Ernest on 2016/04/21 for ReqNo:JC201500378
Dim oRS As Collection 'add by Ernest on 2016/04/21 for ReqNo:JC201500378

'Add by Sam start on 20170317 for Project 停復測自動化
Dim sStepName As String
Dim sSubsystem As String
Dim sWsDeviceFile As String
Dim sStopTestNoList As String
Dim sMsg As String
Dim sErunTicNO  As String
Dim sStopInfo As String
'Add by Sam on 20170317 for Project 停復測自動化


'----
' Init
'----
    sProcID = "cmdOk_Click"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog)
    Screen.MousePointer = vbHourglass
'----
' Condition Checking
'----


    Dim sFile As String

    If (moFwWIP Is Nothing) Then
        Call RaiseError(glERR_INVALIDOBJECT, _
                        FormatErrorText(gsETX_INVALIDOBJECT, "FwWIP"))
    End If

    sFilePath = Environ$("windir") & msFilePath

    sFile = Dir(sFilePath)

    If sFile = "" Then
        UtShowMsgBox "RunCard Excel template not found "
        GoTo ExitHandler
    End If


'Mark by Sam start on 2008/05/19 for ReqNo:M200805020
'    If Len(cboSubSystem.Text) = 0 Then
'        UtShowMsgBox "Please Choise subsystem/site!!"
'        GoTo ExitHandler
'    End If
'Mark by Sam end on 2008/05/19 for ReqNo:M200805020

'add by Sam start on 20080826 for PDA_Project
    If mbNewForm = True Then
        If Len(txtWaferID.Text) = 0 Then
            UtShowMsgBox "WaferID 不可為空白。WaferID can't Empty."
            GoTo ExitHandler
        End If
        If Len(txtPGName.Text) = 0 Then
            UtShowMsgBox "Program Name 不可為空白。Program Name can't Empty."
            GoTo ExitHandler
        End If
        
        'Add by Sam start on 20120409 for ReqNo:JC201200096
        If Len(Trim(Me.cboSetupReason.Text)) = 0 And Me.lblStage <> "FT" Then
            UtShowMsgBox "Setup Reason 不可為空白。Setup Reason can't be Empty."
            GoTo ExitHandler
        End If
        'Add by Sam end on 20120409 for ReqNo:JC201200096
        
    End If
'add by Sam end on 20080826 for PDA_Project
'add by Ernest on 2016/04/21for ReqNo:JC201500378------start
    If Me.lblStage = "FT" Then
        If Me.cboTesterId.Text = "" Then
            UtShowMsgBox "TesterId 不可沒選"
            GoTo ExitHandler
        End If
    End If
'add by Ernest on 2016/04/21 for ReqNo:JC201500378------end

'Add by Sam start on 20170315 for project 停復測自動化
    Set oLot = FwuRetrieveLot(moFwWIP, txtLotID.Text, moAppLog)
    Set oEqp = FwuRetrieveEqp(moFwMDL, cboTesterId.Text, moAppLog)
    sEqType2 = oEqp.CustomAttributes(gsEQP_CUSTOMATTR_EQ_TYPE2)
    
    sSQL = "SELECT A." & gsCAT_TLI_ERUNTICNO & " FROM  " & gsCAT_TBL_LOT_INFO & " A " & _
           " WHERE A." & gsCAT_TLI_LOT_ID & " = '" & txtLotID.Text & "' "
    Set oRS = moProRawSql.QueryDatabase(sSQL)
    If oRS.Count > 0 Then
        sErunTicNO = oRS.Item(1).Item(gsCAT_TLI_ERUNTICNO)
    End If
           
    sStepName = Trim(Mid(cboStep.Text, 7))
    
    If InStr(cboSubSystem.Text, ",") > 0 Then
       sSubsystem = Left(cboSubSystem.Text, InStr(cboSubSystem.Text, ",") - 1)
    Else
       sSubsystem = cboSubSystem.Text
    End If
    
    sWsDeviceFile = txtWsDeviceFile.Text
    
    sMsg = ""
    If NotStopTest(oLot.Id, oLot.CustomAttributes(gsLOT_CUSTOMATTR_IPN), oLot.CustomAttributes(gsLOT_CUSTOMATTR_PROD_GROUP), gsCAT_STOP_TEST_ALL, False, _
                                                         sStepName, oEqp.Id, sEqType2, txtAccType.Text, txtStopTicNo.Text, _
                                                         txtSpecifyEq.Text, txtSpecifyEqID.Text, sErunTicNO, txtPgID.Text, txtPGName.Text, txtPgMode.Text, _
                                                         txtTemp.Text, moFwWIP, moAppLog, moProRawSql, sStopTestNoList, sSubsystem, sWsDeviceFile, sMsg) = False Then
        Call UtShowMsgBox("停測中，停測資訊:" & vbNewLine & sMsg)
        Call cmdCancel_Click
        GoTo ExitHandler
    End If
    
    Call GetStopInfoPByAcc(oLot.Id, oLot.CustomAttributes(gsLOT_CUSTOMATTR_PROD_GROUP), sEqType2, oEqp.Id, sStepName, _
                         txtPgID.Text, txtPGName.Text, txtPgMode.Text, txtWsDeviceFile, txtAccType.Text, sStopTestNoList)
                         
    Call GetStopInfoNPByAcc(oLot.CustomAttributes(gsLOT_CUSTOMATTR_IPN), sEqType2, oEqp.Id, _
                        txtAccType.Text, txtTemp.Text, sMsg, moFwWIP, moAppLog, moProRawSql)
    
    If sStopTestNoList <> "" Then
        sStopInfo = sStopTestNoList
    End If
    
    If sMsg <> "" Then
        sStopInfo = sStopTestNoList & "," & sMsg
    End If
    
    If Left(sStopInfo, 1) = "," Then
        sStopInfo = Mid(sStopInfo, 2)
    End If
    
    If sStopInfo <> "" Then
        lblStopInfo.Caption = sStopInfo
        Call UtShowMsgBox("請確認停測資訊:" & vbNewLine & sStopInfo, vbInformation)
    Else
        lblStopInfo.Caption = "NA"
    End If
    
'Add by Sam end on 20170315 for project 停復測自動化

'----
' Action
'----
    Dim appXL       As Object
    Dim iPosTitle   As Integer
    Set oLot = FwuRetrieveLot(moFwWIP, txtLotID.Text, moAppLog)
    
    If Me.lblStage = "FT" Then
        Set appXL = CreateObject("Excel.Application")
        If appXL Is Nothing Then
            UtShowMsgBox "Excel not found"
            GoTo ExitHandler
        End If
    
                
        'for debug
        If Left(moFwWF.ClientID, 8) = "jeffweng" Then appXL.Visible = True
        
        appXL.cursor = 2
        appXL.DisplayAlerts = False
        appXL.Workbooks.Add sFilePath
        appXL.ActiveWindow.WindowState = vbMaximized
    
        iPosTitle = 1
        'modify by Ernest on 2016/04/21 for ReqNo:JC201500378---start
        'Add by Sam start on 20130327 for Project FVI Setup Recipe
        'If oLot.CustomAttributes(gsLOT_CUSTOMATTR_CUR_EQP_ID) <> "" Then
        '    ' Get TesterId Object
        '    Set oEqp = FwuRetrieveEqp(moFwMDL, oLot.CustomAttributes(gsLOT_CUSTOMATTR_CUR_EQP_ID), moAppLog)
        '    sEqType2 = oEqp.CustomAttributes(gsEQP_CUSTOMATTR_EQ_TYPE2)
        'Else
        '    If oLot.CustomAttributes(gsLOT_CUSTOMATTR_SPLITASSIGNEQID) <> "" Then
        '        Set oEqp = FwuRetrieveEqp(moFwMDL, oLot.CustomAttributes(gsLOT_CUSTOMATTR_SPLITASSIGNEQID), moAppLog)
        '        sEqType2 = oEqp.CustomAttributes(gsEQP_CUSTOMATTR_EQ_TYPE2)
        '    End If
        'End If
        sSQL = "select distinct b.eqtype2 from view_b2b_fweqarea b where eqid='" & Me.cboTesterId.Text & "'"
        Set oRS = moProRawSql.QueryDatabase(sSQL)
        If oRS.Count > 0 Then
            sEqType2 = oRS.Item(1).Item("eqtype2")
        End If
        'modify by Ernest on 2016/04/21 for ReqNo:JC201500378-------end
        'PROGRAMMER : AT3-300AL / FT-940 --> comment added by Jack on 2014/01/13 for JC201400011.
        If sEqType2 = "AT3-300AL" Or sEqType2 = "FT-940" Then '目前這兩種位置相同
            Call AT3SetupForm(sEqType2, appXL, oLot, iPosTitle)
        Else
        'Add by Sam start on 20130327 for Project FVI Setup Recipe
            Call FtSetupForm(appXL, oLot, iPosTitle)
        End If
        
'將整段產生Excel程式碼移動到FT判斷中 by Dexter on 2022/11/21 for BE#202200310 CP MES  Printer setup/repair form rule 移除 呼叫 excel 功能 Start
    appXL.Sheets("WS").Select
    appXL.ActiveWindow.SelectedSheets(1).Visible = False
    appXL.Sheets("FT").Select
    appXL.ActiveWindow.SelectedSheets(1).Visible = False
    
    'Add by Sam Start on 20130327 for Project FVI Setup Recipe ,將所有頁面隱藏,只留setup頁面
    appXL.Sheets("AT3-300AL").Select
    appXL.ActiveWindow.SelectedSheets(1).Visible = False
    appXL.Sheets("FT-940").Select
    appXL.ActiveWindow.SelectedSheets(1).Visible = False
    appXL.Sheets("SCANNER").Select
    appXL.ActiveWindow.SelectedSheets(1).Visible = False
    appXL.Sheets("T&R").Select
    appXL.ActiveWindow.SelectedSheets(1).Visible = False
    appXL.Sheets(gsRESOURCE_COMBO).Select
    appXL.ActiveWindow.SelectedSheets(1).Visible = False
    appXL.Sheets("LASER").Select
    appXL.ActiveWindow.SelectedSheets(1).Visible = False
    'Add by Sam End on 20130327 for Project FVI Setup Recipe
    
    appXL.Sheets(1).Select

    appXL.range("A1").Select
    appXL.cursor = 1
    appXL.DisplayAlerts = True

    ChDir "C:\FASTech"
    appXL.ActiveWorkbook.SaveAs "C:\FASTech\" & Me.lblStage & Format(Now, "YYYYMMDDHHMMSS") & ".xls", _
        , "", "", False, False

    appXL.Visible = True
    
'    appXL.ActiveWindow.SelectedSheets(1).PrintOut Copies:=1, Collate:=True
'add by Sam start on 20080826 for PDA_Project
    If mbNewForm = True Then
        '新架機單不自動列印
        Call InsNewForm(oLot.Id)
    Else
        appXL.ActiveWindow.SelectedSheets(1).PrintOut Copies:=1, Collate:=True
    End If
'add by Sam end on 20080826 for PDA_Project
    Set appXL = Nothing
'將整段產生Excel程式碼移動到FT判斷中 by Dexter on 2022/11/21 for BE#202200310 CP MES  Printer setup/repair form rule 移除 呼叫 excel 功能 End
    
    Else
        'Call WsSetupForm(appXL, oLot, iPosTitle) 'Comment out by Dexter on 2022/11/21 for BE#202200310 CP MES  Printer setup/repair form rule 移除 呼叫 excel 功能
        If mbNewForm = True Then
            '新架機單不自動列印
            Call InsNewForm(oLot.Id)
        End If
    End If
        
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

Private Sub cmdQuery_Click()
    Call ExecQuery
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
' Description:  <Type your function description here...>
'--------------------------------------------------------------------------------
' Author:       Jeff Weng, MXIC 2005/03/30
'--------------------------------------------------------------------------------
' Parameters:
'................................................................................
' ARGUMENTS
'   oLogCtrl            (I) [Optional] Valid reference to a object object
'   Argument2           (I) <Description goes here...>
'   Argument3           (O) <Description goes here...>
'
' NON-LOCAL ARGS
'   NonLoc1         (I) <Description goes here...>
'   NonLoc2         (I) <Description goes here...>
'
' RETURN TYPE
'   Boolean         (R) True/False
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

    'Add by Sam start on 2017017 for Project  停復測自動化
    txtPgID = ""
    txtPgMode = ""
    txtTemp = ""
    txtWsDeviceFile = ""
    txtAccType = ""
    txtSpecifyEq = ""
    txtSpecifyEqID = ""
    txtStopTicNo = ""
    txtPgID.Visible = False
    txtPgMode.Visible = False
    txtTemp.Visible = False
    txtWsDeviceFile.Visible = False
    txtAccType.Visible = False
    txtSpecifyEq.Visible = False
    txtSpecifyEqID.Visible = False
    txtStopTicNo.Visible = False
    lblStopInfo.Caption = ""
    lblStopInfo.Visible = False
    'Add by Sam end on 2017017 for Project  停復測自動化
    Me.cboTesterId.Clear 'Add by HouYu for BE MES Phase 95 修改F1-列印架機單(E-Test)功能欄位值異常 on 20221128

'----
' Condition Checking
'----
    ' <Put your condition checking codes here>...

'----
' Action
'----
    ' <Put your Action codes here>...
         
    If msLotId <> "" Then
        Me.txtLotID = msLotId
        Call cmdQuery_Click
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
        Init = False
        On Error GoTo ExitHandler:
        Call HandleError(True, typErrInfo, , moAppLog)
    End If
End Function
'================================================================================
' Sub: ExecQuery()
'--------------------------------------------------------------------------------
' Description:  <Type your Sub description here...>
'--------------------------------------------------------------------------------
' Author:       Jeff Weng, MXIC 2005/03/30
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
'
'================================================================================
Private Sub ExecQuery()
On Error GoTo ExitHandler:
Dim sProcID As String
Dim typErrInfo As tErrInfo

'----
' Init
'----
    sProcID = "ExecQuery"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog)
    Screen.MousePointer = vbHourglass
'----
' Condition Checking
'----
    ' <Put your condition checking codes here>...
        
    Dim oLot    As FwLot
    Dim oEqp    As FwEquipment
    'add by Ernest on 2016/04/21 for ReqNo:JC201500378
    Dim sSQL        As String
    Dim oRS         As Collection
    
    Dim sBodySize As String 'Add by Sam on 20200430 for Project FT配件資料標準化
    
    'Added by Jack on 2021/03/12 for CP測前併批專案 <Start>.
    Dim sOriginalLotID         As String
    Dim bIsVirtualMerge        As Boolean
    Dim bIsCPSmallLot          As Boolean
    Dim sVirtualLotId          As String
    
    Dim iIdx_1                 As Integer
    Dim iIdx_2                 As Integer
    
    Dim sCheckLotId            As String
    Dim vCheckProbingQty       As Variant
    Dim vCheckWaferId          As Variant
    Dim sCheckWaferIdList      As String
    Dim sCheckResult           As String 'CPSamling回傳值.
    
    Dim iMinProbingQty         As Integer
    Dim sMinProbingLotID       As String
    Dim sMinProbingWferIdList  As String
    'Added by Jack on 2021/03/12 for CP測前併批專案 <End>.
    
    If Me.txtLotID.Text <> "" Then
        Set oLot = FwuRetrieveLot(moFwWIP, msLotId, moAppLog)
'        If oLot Is Nothing Then
'           UtShowMsgBox "This Lot was not exist. !!" & vbNewLine & _
'                            "此Lot不存在 !!"
'            cmdOK.Enabled = False
'            GoTo ExitHandler
'        End If
        If oLot Is Nothing Then
            Me.lblLotStatus = ""
            'Me.lblCurrentStep = "" 'mark by Ernest on 2016/04/21 for ReqNo:JC201500378
            Me.lblIPN = ""
            Me.lblLotOwner = ""
            Me.lblLotId = ""
            Me.lblStage = ""
           ' Me.lblTesterID = "" 'mark by Ernest on 2016/04/21 for ReqNo:JC201500378
            
            cmdOK.Enabled = False
            GoTo ExitHandler
        End If
        'add by Ernest on 2016/04/21 for ReqNo:JC201500378-------start
        sSQL = "select stage " & _
                " from fwadmin.tbl_lot_attribute where lotid='" & txtLotID.Text & "'"
        Set oRS = moProRawSql.QueryDatabase(sSQL)
        If oRS.Count > 0 Then
            Me.lblStage = oRS.Item(1).Item("stage")
        End If
        
        Dim mfrmMain As New frmCpSampling 'Added by Jack on 2021/03/12 for CP測前併批專案
        
        If Me.lblStage <> "FT" Then
        'add by Ernest on 2016/04/21 for ReqNo:JC201500378-------end
            'Added by Jack on 2021/03/12 for CP測前併批專案 <Start>.
            'for WS, 若為VirtualMerge Lot, 要判斷是否需要取代Lot ? (以整組按SLotNO大至少排序, 取針痕次數最小 & 尚有待測片者)
            sOriginalLotID = txtLotID.Text
            iMinProbingQty = -1
            
            Call modCPMerge.CheckMergeType(moProRawSql, moAppLog, sOriginalLotID, bIsVirtualMerge, bIsCPSmallLot, sVirtualLotId)
            If bIsVirtualMerge Then
                sSQL = "select vm." & gsCAT_TVMG_SLOTNO & ", vm." & gsCAT_TVMG_LOTID & " " & _
                        " from " & gsCAT_TBL_LOT_INFO & " li, " & _
                             " " & gsCAT_TBL_LOT_ATTRIBUTE & " la, " & _
                             " (select " & gsCAT_TVMG_LOTID & ", min(" & gsCAT_TVMG_SLOTNO & ") " & gsCAT_TVMG_SLOTNO & " " & _
                                " from " & gsCAT_TBL_VIRTUAL_MERGE & " vm " & _
                               " where " & gsCAT_TVMG_VIRTUALLOTID & " = '" & sVirtualLotId & "' " & _
                                 " and " & gsCAT_TVMG_DELETEFLAG & " = 'N' " & _
                               " group by " & gsCAT_TVMG_LOTID & ") vm " & _
                       " where li." & gsCAT_TLI_VIRTUALLOTID & " = '" & sVirtualLotId & "' " & _
                         " and li." & gsCAT_TLI_LOT_ID & " = la." & gsCAT_TLATT_LOTID & " " & _
                         " and (la." & gsCAT_TLATT_STEPID & ", la." & gsCAT_TLATT_STEPNAME & ") = " & _
                             " (select la." & gsCAT_TLATT_STEPID & ", la." & gsCAT_TLATT_STEPNAME & " " & _
                                " from " & gsCAT_TBL_LOT_ATTRIBUTE & " la " & _
                               " where " & gsCAT_TLATT_LOTID & " = '" & sOriginalLotID & "') " & _
                         " and la." & gsCAT_TLATT_LOTID & " = vm." & gsCAT_TVMG_LOTID & " " & _
                       " order by " & gsCAT_TVMG_SLOTNO & " desc, " & gsCAT_TVMG_LOTID & " "
                Set oRS = moProRawSql.QueryDatabase(sSQL)
                If oRS.Count > 0 Then
                
                    For iIdx_1 = 1 To oRS.Count
                        sCheckLotId = oRS.Item(iIdx_1).Item(gsCAT_TVMG_LOTID)
                        
                        Load mfrmMain
                        With mfrmMain
                            Set .CwMbxControl = moCwMbx
                            Set .FwMDLControl = moFwMDL
                            Set .FwWIPControl = moFwWIP
                            Set .FwOPRControl = moFwOPR
                            Set .FwPRPControl = moFwPRP
                            Set .FwWFControl = moFwWF
                            Set .FwCATControl = moFwCAT
                            Set .MainTraceLog = moAppLog
                            Set .CwMbxControl = moCwMbx
                            Set .ProRawSqlControl = moProRawSql
                            .LotID = sCheckLotId
                            .Init
                               
                            sCheckResult = .CheckResult
                            
                            sCheckWaferIdList = ""
                            vCheckProbingQty = ""
                            '一般SORT站 : AllLotFlag = "Y"
                            '抽測SORT站 : AllLotFlag = "N"
                            'Modified by Jack on 2021/03/26 for CP測前併批專案 UAT bug fix
                            'Lot有部份已測時, txtWaitTestPCs 會有空值的現象, 但是 txtUntestedWaferId有值.
                            'If (.txtWaitTestPCs = "" And .txtAllLotFlag = "Y") Or (Val(.txtWaitTestPCs) > 0 And .txtAllLotFlag = "N") Then
                            If (.txtWaitTestPCs = "" And .txtAllLotFlag = "Y") Or ((Trim(.txtUntestedWaferId) <> "" Or Val(.txtWaitTestPCs) > 0) And .txtAllLotFlag = "N") Then
                                For iIdx_2 = 1 To .spdUnTestedWaferId.MaxRows
                                    .spdUnTestedWaferId.GetText 1, iIdx_2, vCheckWaferId
                                    If iIdx_2 = 1 Then
                                        .spdUnTestedWaferId.GetText 2, iIdx_2, vCheckProbingQty
                                        sCheckWaferIdList = CStr(vCheckWaferId)
                                    Else
                                        sCheckWaferIdList = sCheckWaferIdList & "," & CStr(vCheckWaferId)
                                        If iIdx_2 >= 3 Then Exit For
                                    End If
                                Next
                                
                                If iMinProbingQty = -1 Or Val(CStr(vCheckProbingQty)) < iMinProbingQty Then
                                    iMinProbingQty = Val(CStr(vCheckProbingQty))
                                    sMinProbingLotID = sCheckLotId
                                    sMinProbingWferIdList = sCheckWaferIdList
                                End If
                            End If
                               
                            '.Show vbModal
                            Unload mfrmMain
                        End With
                    Next
                End If
                
            End If
            
            If sMinProbingLotID <> "" And sMinProbingLotID <> txtLotID.Text Then
                'UtShowMsgBox "LotID : " & txtLotID.Text & " 被同一VirtualLotID " & sMinProbingLotID & " 置換 !!"  'Modify by Yanwei on 20220609 for CP併批測試 Phase II 第一階段需求 => HW setup功能
                txtLotID.Text = sMinProbingLotID
            End If
            'Added by Jack on 2021/03/12 for CP測前併批專案 <End>.
            If Not (UCase(Left(oLot.CurrentStep.Steps.Item(1).Description, 4)) = "SORT" Or _
                UCase(Left(oLot.CurrentStep.Steps.Item(1).Description, 2)) = "FT" Or _
                UCase(Left(oLot.CurrentStep.Steps.Item(1).Description, 4)) = "TQAE" Or _
                UCase(Left(oLot.CurrentStep.Steps.Item(1).Description, 10)) = "PROGRAMMER") Then
                UtShowMsgBox "This Lot was not in Testing Step. !!" & vbNewLine & _
                                "此Lot未在測試站別, 無法列印 !!"
                cmdOK.Enabled = False
                GoTo ExitHandler
            End If
            
            
            If oLot.CustomAttributes(gsLOT_CUSTOMATTR_CUR_EQP_ID) <> "" Then
                ' Get TesterId Object
                Set oEqp = FwuRetrieveEqp(moFwMDL, oLot.CustomAttributes(gsLOT_CUSTOMATTR_CUR_EQP_ID), moAppLog)
            Else
                Set oEqp = FwuRetrieveEqp(moFwMDL, oLot.CustomAttributes(gsLOT_CUSTOMATTR_SPLITASSIGNEQID), moAppLog)
            End If
            If oEqp Is Nothing Then
                UtShowMsgBox "The Equipment of this Lot was not exited. !!" & vbNewLine & _
                                "此Lot之測試機台不存在 !!"
                cmdOK.Enabled = False
                GoTo ExitHandler
            End If
        End If
        Call CheckForm(Me.txtLotID.Text) 'add by Sam on 20080827 for PDA_PROJECT
        
    End If
    
    'Add by Sam start on 2017017 for Project  停復測自動化
    txtPgID = ""
    txtPgMode = ""
    txtTemp = ""
    txtWsDeviceFile = ""
    txtAccType = ""
    txtSpecifyEq = ""
    txtSpecifyEqID = ""
    txtStopTicNo = ""
    txtPgID.Visible = False
    txtPgMode.Visible = False
    txtTemp.Visible = False
    txtWsDeviceFile.Visible = False
    txtAccType.Visible = False
    txtSpecifyEq.Visible = False
    txtSpecifyEqID.Visible = False
    txtStopTicNo.Visible = False
    lblStopInfo.Caption = ""
    lblStopInfo.Visible = False
    'Add by Sam end on 2017017 for Project  停復測自動化
    
'----
' Action
'----
    ' <Put your Action codes here>...
    'modify by Ernest on 2016/04/21 for ReqNO:JC201500378---移至前方
    'Dim sSQL        As String
    'Dim oRS         As Collection
    Dim sWaferID() As String 'add by Sam on 20080827 for PDA_PROJECT
    Dim iIdx As Integer 'add by Sam on 20080827 for PDA_PROJECT
    
    Dim bTECNLotSpec            As Boolean  'Add by SAM ON 20100407 FOR Project By Lot TECN
    Dim sTecnLotID            As String   'Add by SAM ON 20100407 FOR Project By Lot TECN
    
    'Added by Jack on 2011/09/23 for JC201100228. 不FollowProduct之工程品 改為 帶出新維修單畫面. <Start>
    Dim sErunTicNO              As String
    Dim sFollowProd             As String
    Dim sSubSystemType          As String
    'Added by Jack on 2011/09/23 for JC201100228. 不FollowProduct之工程品 改為 帶出新維修單畫面. <End>
    
    Dim sAssignProbeCard As String 'Add by Sam on 20141128 for ReqNo:JC201400402
    Dim sAssignLoadboard As String 'Add by SAm on 20181206 for Project FT ADVAN Tester  Auto Lot Start
    
    'add by Ernest on 2016/04/21 for ReqNo:JC201500378----start
    Dim sStepName As String
    Dim sStepNo    As String
    Dim sProdGroup As String
    Dim sPath As String
    Dim sEqType2 As String
    Dim i As Integer
    Dim vStepName As Variant
    'add by Ernest on 2016/04/21 for ReqNo:JC201500378----end
    
    
    Me.lblLotStatus = oLot.CustomAttributes(gsLOT_CUSTOMATTR_STATUS)
    'mark by Ernest on 2016/04/21 for ReqNo:JC201500378
    'Me.lblCurrentStep = oLot.CurrentStep.Steps.Item(1).Id & " " & oLot.CurrentStep.Steps.Item(1).Description
    Me.lblIPN = oLot.CustomAttributes(gsLOT_CUSTOMATTR_IPN)
    Me.lblLotOwner = oLot.CustomAttributes(gsLOT_CUSTOMATTR_LOT_OWNER)
    
    'mark by Ernest on 2016/04/21 for ReqNo:JC201500378--start
    'If oLot.CustomAttributes(gsLOT_CUSTOMATTR_CUR_EQP_ID) <> "" Then
    '    Me.lblTesterID = oLot.CustomAttributes(gsLOT_CUSTOMATTR_CUR_EQP_ID)
    'Else
    '    Me.lblTesterID = oLot.CustomAttributes(gsLOT_CUSTOMATTR_SPLITASSIGNEQID)
    'End If
    'mark by Ernest on 2016/04/21 for ReqNo:JC201500378--end
    
    sSQL = "select stage " & _
            " from fwadmin.tbl_lot_attribute where lotid='" & txtLotID.Text & "'"
    Set oRS = moProRawSql.QueryDatabase(sSQL)
    If oRS.Count > 0 Then
        Me.lblStage = oRS.Item(1).Item("stage")
    End If
    Set oRS = Nothing
    Dim oItem As Object
    
    'Add by SAM START ON 20100514 FOR Project LOT TECN
    
    'Added by Jack on 2011/09/23 for JC201100228. 不FollowProduct之工程品 改為 帶出新維修單畫面.
    'Add : gsCAT_TLI_ERUNTICNO & sErunTicNO
    'Moify by Sam on 20141128 for ReqNo:JC201400402 ,增加AssignProbeCard
    sSQL = " SELECT A." & gsCAT_TLI_TECN_LOT_ID & ", " & _
             gsCAT_TLI_ERUNTICNO & ", " & _
             gsCAT_TLI_ASSIGNPROBECARD & ", " & gsCAT_TLI_ASSIGNLOADBOARD & " " & _
           " FROM " & gsCAT_TBL_LOT_INFO & " A " & _
           " WHERE A." & gsCAT_TLI_LOT_ID & " ='" & txtLotID.Text & "' "
    Set oRS = moProRawSql.QueryDatabase(sSQL)
    If oRS.Count > 0 Then
        sTecnLotID = oRS.Item(1).Item(1)
        sErunTicNO = oRS.Item(1).Item(2)
        sAssignProbeCard = oRS.Item(1).Item(gsCAT_TLI_ASSIGNPROBECARD) 'Add by Sam on 20141128 for ReqNo:JC201400402
        sAssignLoadboard = oRS.Item(1).Item(gsCAT_TLI_ASSIGNLOADBOARD) 'Add by Sam on 20181206 for Project FT ADVAN Tester  Auto Lot Start
    End If
    'Add by SAM END ON 20100514 FOR Project LOT TECN
    
    'Add by Sam Start on 20200430 for Project FT配件資料標準化,Add bodysize
    sSQL = "select i." & gsCAT_TIM_BODY_SIZE & " from " & gsCAT_TBL_IPN_MASTER & " i " & _
           "  where i." & gsCAT_TIM_IPN & "= '" & oLot.CustomAttributes(gsLOT_CUSTOMATTR_IPN) & "' "
    Set oRS = moProRawSql.QueryDatabase(sSQL)
    If oRS.Count > 0 Then
        sBodySize = oRS.Item(1).Item(gsCAT_TIM_BODY_SIZE)
    End If
    'Add by Sam End on 20200430 for Project FT配件資料標準化,Add bodysize

    'Added by Jack on 2011/09/23 for JC201100228. 不FollowProduct之工程品 改為 帶出新維修單畫面. <Start>
    If Trim(sErunTicNO) <> "" Then
        'Modified by Jack on 2014/01/13 for JC201400011.
        'Stage='WS' 改為 畫面的Stage. (ori:" and " & gsCAT_TER_STAGE & "='WS' ")
        sSQL = "SELECT " & gsCAT_TER_FOLLOW_PRODUCT & _
                  " from " & gsCAT_TBL_ERUN_REQ & _
                  " where " & gsCAT_TER_LOT_ID & " ='" & txtLotID.Text & "'" & _
                  " and " & gsCAT_TER_TICKET_NO & "='" & sErunTicNO & "' " & _
                  " and " & gsCAT_TER_STAGE & "='" & lblStage & "' "
                  
        '          " and " & gsCAT_TER_STAGE & "='WS' "
        Set oRS = moProRawSql.QueryDatabase(sSQL)
        If oRS.Count > 0 Then
            sFollowProd = oRS.Item(1).Item(1)
        End If
    End If
    'Added by Jack on 2011/09/23 for JC201100228. 不FollowProduct之工程品 改為 帶出新維修單畫面. <End>
    'add by Ernest on 2016/04/21 for ReqNo:JC201500378--start
    If Me.lblStage <> "FT" Then
        Me.cboStep.Clear
        Me.cboStep.Text = oLot.CurrentStep.Steps.Item(1).Id & " " & oLot.CurrentStep.Steps.Item(1).Description
        Me.cboTesterId.Clear
        If oLot.CustomAttributes(gsLOT_CUSTOMATTR_CUR_EQP_ID) <> "" Then
            Me.cboTesterId.Text = oLot.CustomAttributes(gsLOT_CUSTOMATTR_CUR_EQP_ID)
        Else
            Me.cboTesterId.Text = oLot.CustomAttributes(gsLOT_CUSTOMATTR_SPLITASSIGNEQID)
        End If
    Else
        sSQL = "select i." & gsCAT_TIM_PROD_GROUP & " from " & gsCAT_TBL_IPN_MASTER & _
               " i,tbl_lot_attribute a where i." & gsCAT_TIM_IPN & " = a.ipn and a.lotid='" & Me.txtLotID.Text & "'"
        Set oRS = moProRawSql.QueryDatabase(sSQL)
        If oRS.Count > 0 Then
            sProdGroup = oRS.Item(1).Item("Prodgroup")
        End If
        
        sPath = GetPath(Me.txtLotID.Text, moProRawSql, moAppLog)
        
        If mbChangestep = False Then
            Me.cboStep.Clear
             sSQL = "select distinct " & gsCAT_TRSL_STEP_NO & " ," & gsCAT_TRSL_STEP_NAME & " from " & gsCAT_TBL_ROUTE_STEP_LIST & _
                    " where " & gsCAT_TRSL_PATH & " = '" & sPath & "'" & _
                    "and (stepname like 'FT_' or stepname like 'FT__' or stepname like 'TQAE%')"
            Set oRS = moProRawSql.QueryDatabase(sSQL)
            If oRS.Count > 0 Then
                For i = 1 To oRS.Count
                    sStepNo = oRS.Item(i).Item(gsCAT_TRSL_STEP_NO)
                    sStepName = oRS.Item(i).Item(gsCAT_TRSL_STEP_NAME)
                    Me.cboStep.AddItem sStepNo & " " & sStepName
                    If oLot.CurrentStep.Steps.Item(1).Id = sStepNo Then
                        Me.cboStep.Text = sStepNo & " " & sStepName
                    End If
                Next
            End If
        End If
        
        If Me.cboStep.Text <> "" Then
            vStepName = Split(Me.cboStep.Text, " ")
            sStepName = Trim(vStepName(1))
            sStepNo = Trim(vStepName(0))
        Else
            Me.cboStep.Text = oLot.CurrentStep.Steps.Item(1).Id & " " & oLot.CurrentStep.Steps.Item(1).Description
            sStepName = oLot.CurrentStep.Steps.Item(1).Description
            sStepNo = oLot.CurrentStep.Steps.Item(1).Id
        End If
                
        'Mark by Sam Start on 20190312 for Project FT ADVAN Tester,應User要求再改回直接取Assign機台
'        If mbChangetesterId = False Then
'            Me.cboTesterId.Clear
'            sSQL = "select distinct eqid from view_b2b_fweqarea b,("
'            If Trim(sErunTicNO) <> "" And Trim(sFollowProd) = "N" Then
'                sSQL = sSQL & "select distinct d.EQTYPE2" & _
'                        " from tbl_erun_recipe d,tbl_lot_info tli,tbl_erun_req c " & _
'                        "where tli.lotid = '" & Me.txtLotID.Text & "' " & _
'                        "and tli.erunticno = c.ticketno(+) " & _
'                        " and c.routeid = d.path and c.ticketno = d.docno and d.stepname = '" _
'                        & sStepName & "' and d.stepno ='" & sStepNo & "'"
'            Else
'                If Trim(sTECNLotID) <> "" Then
'
'                    sSQL = sSQL & "select distinct " & gsCAT_PSES_EQTYPE2 & " from " & gsCAT_TBL_LOT_STEP_EQ_SPEC & _
'                            " where '" & sTECNLotID & "' like " & gsCAT_TLSPES_TECNLOTID & " and " & gsCAT_TLSPES_PATH & " ='" & sPath & _
'                            "' and " & gsCAT_TLSPES_STEPNAME & " = '" & sStepName & "' and " & gsCAT_TLSPES_DELETEFLAG & " ='N'"
'                Else
'                    sSQL = sSQL & "select distinct " & gsCAT_PSES_EQTYPE2 & " from " & gsCAT_TBL_PROD_STEP_EQ_SPEC & " where " & _
'                            gsCAT_PSES_PATH & " = '" & sPath & "' and " & gsCAT_PSES_PROD_GROUP & "='" & sProdgroup & _
'                            "' and " & gsCAT_PSES_STEP_NAME & " = '" & sStepName & "' and docstatus ='Active'"
'                End If
'            End If
'            sSQL = sSQL & ") a where b.eqtype2 = a.eqtype2"
'            Set oRS = moProRawSql.QueryDatabase(sSQL)
'            If oRS.Count > 0 Then
'                For i = 1 To oRS.Count
'                  Me.cboTesterId.AddItem oRS.Item(i).Item("eqid")
'                Next
'            End If
'        End If
        'Mark by Sam End on 20190312 for Project FT ADVAN Tester,應User要求再改回直接取Assign機台
        'Add by Sam start on 20190312 for Project FT ADVAN Tester,應User要求再改回直接取Assign機台
        Me.cboTesterId.Clear
        If oLot.CustomAttributes(gsLOT_CUSTOMATTR_CUR_EQP_ID) <> "" Then
            Me.cboTesterId.AddItem oLot.CustomAttributes(gsLOT_CUSTOMATTR_CUR_EQP_ID)
            Me.cboTesterId.Text = oLot.CustomAttributes(gsLOT_CUSTOMATTR_CUR_EQP_ID)
        Else
            Me.cboTesterId.AddItem oLot.CustomAttributes(gsLOT_CUSTOMATTR_SPLITASSIGNEQID)
            Me.cboTesterId.Text = oLot.CustomAttributes(gsLOT_CUSTOMATTR_SPLITASSIGNEQID)
        End If
        'Add by Sam End on 20190312 for Project FT ADVAN Tester,應User要求再改回直接取Assign機台
        If Me.cboTesterId.Text <> "" Then
            sSQL = "select distinct b.eqtype2 from view_b2b_fweqarea b where eqid='" & Me.cboTesterId.Text & "'"
            Set oRS = moProRawSql.QueryDatabase(sSQL)
            If oRS.Count > 0 Then
                sEqType2 = oRS.Item(1).Item("eqtype2")
            End If
        End If
    End If
    'add by Ernest on 2016/04/21 for ReqNo:JC201500378--end
    'Add by Sam start on 20100514 for Project Lot TECN
    'Added by Jack on 2011/09/23 for JC201100228. 不FollowProduct之工程品 改為 帶出新維修單畫面. <Start>
    If Me.lblStage <> "FT" Then 'add by Ernest on 2016/04/21 for ReqNo:JC201500378
        If Trim(sErunTicNO) <> "" And Trim(sFollowProd) = "N" Then
    '        If InStr(cboSubSystem.Text, ",") > 0 Then
    '           sSubSystemType = Left(cboSubSystem.Text, InStr(cboSubSystem.Text, ",") - 1)
    '        Else
    '           sSubSystemType = cboSubSystem.Text
    '        End If
            'Modify by Sam on 20200430 for Project FT配件資料標準化,Add Loadboard
            sSQL = "select a.subsystem,'' as MAXSITE ,a." & gsCAT_TERE_LOADBOARDTYPE & " as loadboard " & _
                     "from tbl_erun_recipe a " & _
                    "where a.docno = '" & sErunTicNO & "' " & _
                      "and a.stepno = '" & oLot.CurrentStep.Steps.Item(1).Id & "' " & _
                      "and a.eqtype2 = '" & oEqp.CustomAttributes(gsEQP_CUSTOMATTR_EQ_TYPE2) & "' " & _
                      "and a.deleteflag = 'N' "
            'Add by Sam on 20150312 for ReqNo:JC201500077,增加限制stage=WS的才有PROBECARD條件
            If lblStage.Caption = "WS" Then
                sSQL = sSQL & " and nvl(a." & gsCAT_TERE_PROBECARDTYPE & ",' ')= nvl('" & sAssignProbeCard & "',' ') " 'Add by Sam on 20141128 for ReqNo:JC201400402 ,增加Probecard條件
            End If
            
            'Add by Sam Start on 20181206 for Project FT ADVAN Tester  Auto Lot Start,增加Loadboard條件
            If lblStage.Caption = "FT" And sAssignLoadboard <> "" Then
                sSQL = sSQL & " and nvl(a." & gsCAT_TERE_LOADBOARDTYPE & ",' ')= nvl('" & sAssignLoadboard & "',' ') " 'Add by Sam on 20141128 for ReqNo:JC201400402 ,增加Probecard條件
            End If
            'Add by Sam End on 20181206 for Project FT ADVAN Tester  Auto Lot Start
              
            sSQL = sSQL & " and a." & gsCAT_TERE_PATH & "='" & oLot.CustomAttributes(gsLOT_CUSTOMATTR_ROUTE) & "' "      'Add by Sam on 20141128 for ReqNo:JC201400402 ,增加Path條件
            
        'Added by Jack on 2011/09/23 for JC201100228. 不FollowProduct之工程品 改為 帶出新維修單畫面. <End>
        ElseIf Trim(sTecnLotID) <> "" Then
            'Add by Sam Start on 20181206 for Project FT ADVAN Tester  Auto Lot Start,增加Loadboard條件
            sSQL = "select Subsystem,MAXSITE " & _
                   " ,decode(trim(" & gsCAT_TLSPES_LOADBOARDTYPE & " || " & gsCAT_TLSPES_CONTACTBOARDTYPE & " || " & gsCAT_TLSPES_BURN_IN_BOARD & "),null, " & _
                   "  fun_get_ft_acc_bypgm(" & gsCAT_TLSPES_STEPNAME & ", '" & sBodySize & "', " & gsCAT_TLSPES_PGID & ", " & gsCAT_TLSPES_PGNAME & " , 'loadboard') ,  " & gsCAT_TLSPES_LOADBOARDTYPE & " ) as loadboard  " & _
                   "from tbl_lot_step_eq_spec where " & _
                   " '" & sTecnLotID & "' like  tecnlotid and " & _
                   "stepno='" & oLot.CurrentStep.Steps.Item(1).Id & "' and " & _
                   "eqtype2='" & oEqp.CustomAttributes(gsEQP_CUSTOMATTR_EQ_TYPE2) & "' and " & _
                   "deleteflag ='N' "
            'Add by Sam on 20150312 for ReqNo:JC201500077,增加限制stage=WS的才有PROBECARD條件
            If lblStage.Caption = "WS" Then
                sSQL = sSQL & " and nvl(" & gsCAT_TLSPES_PROBECARDTYPE & ",' ')= nvl('" & sAssignProbeCard & "',' ') " 'Add by Sam on 20141128 for ReqNo:JC201400402 ,增加Probecard條件
            End If
'            'Mark by Sam start on 20200428 for Project FT配件資料標準化,不取消的話自動引用的情況會撈不出來
'            'Add by Sam Start on 20181206 for Project FT ADVAN Tester  Auto Lot Start,增加Loadboard條件
'            If lblStage.Caption = "FT" And sAssignLoadboard <> "" Then
'                sSQL = sSQL & " and nvl(a." & gsCAT_TLSPES_LOADBOARDTYPE & ",' ')= nvl('" & sAssignLoadboard & "',' ') " 'Add by Sam on 20141128 for ReqNo:JC201400402 ,增加Probecard條件
'            End If
'            'Add by Sam End on 20181206 for Project FT ADVAN Tester  Auto Lot Start
            'Mark by Sam End on 20200428 for Project FT配件資料標準化,不取消的話自動引用的情況會撈不出來
            sSQL = sSQL & " and " & gsCAT_TLSPES_PATH & "='" & oLot.CustomAttributes(gsLOT_CUSTOMATTR_ROUTE) & "' "      'Add by Sam on 20141128 for ReqNo:JC201400402 ,增加Path條件
        
        Else
        'Add by Sam end on 20100514 for Project Lot TECN
            'Modify by Sam on 20200430 for Project FT配件資料標準化,Add Loadboardtype
            sSQL = "select Subsystem,MAXSITE " & _
                   " ,decode(trim(" & gsCAT_PSES_LOADBOARD_TYPE & " || " & gsCAT_PSES_CONTACTBOARD_TYPE & " || " & gsCAT_PSES_BURN_IN_BOARD & "),null, " & _
                   "  fun_get_ft_acc_bypgm(" & gsCAT_PSES_STEP_NAME & ", '" & sBodySize & "', " & gsCAT_PSES_PG_ID & ", " & gsCAT_PSES_PG_NAME & " , 'loadboard') ,  " & gsCAT_PSES_LOADBOARD_TYPE & " ) as loadboard  " & _
                   "from tbl_prod_step_eq_spec where " & _
                   "prodgroup='" & oLot.CustomAttributes(gsLOT_CUSTOMATTR_PROD_GROUP) & "' and " & _
                   "stepno='" & oLot.CurrentStep.Steps.Item(1).Id & "' and " & _
                   "eqtype2='" & oEqp.CustomAttributes(gsEQP_CUSTOMATTR_EQ_TYPE2) & "' and " & _
                   "docstatus ='Active' "
            'Add by Sam on 20150312 for ReqNo:JC201500077,增加限制stage=WS的才有PROBECARD條件
            If lblStage.Caption = "WS" Then
                sSQL = sSQL & " and nvl(" & gsCAT_PSES_PROBECARD_TYPE & ",' ')= nvl('" & sAssignProbeCard & "',' ') " 'Add by Sam on 20141128 for ReqNo:JC201400402 ,增加Probecard條件
            End If
            'Mark by Sam start on 20200428 for Project FT配件資料標準化,不取消的話自動引用的情況會撈不出來
            'Add by Sam Start on 20181206 for Project FT ADVAN Tester  Auto Lot Start,增加Loadboard條件
'            If lblStage.Caption = "FT" And sAssignLoadboard <> "" Then
'                sSQL = sSQL & " and nvl(a." & gsCAT_PSES_LOADBOARD_TYPE & ",' ')= nvl('" & sAssignLoadboard & "',' ') "
'            End If
            'Add by Sam End on 20181206 for Project FT ADVAN Tester  Auto Lot Start
            'Mark by Sam End on 20200428 for Project FT配件資料標準化,不取消的話自動引用的情況會撈不出來
            sSQL = sSQL & " and path='" & oLot.CustomAttributes(gsLOT_CUSTOMATTR_ROUTE) & "' "      'Add by Sam on 20141128 for ReqNo:JC201400402 ,增加Path條件
        End If
    'add by Ernest on 2016/04/21 for ReqNo:JC201500378--start
    Else
        If Trim(sErunTicNO) <> "" And Trim(sFollowProd) = "N" Then
            'Modify by Sam on 20200430 for Project FT配件資料標準化,Add Loadboard
              sSQL = "select a.subsystem,'' as MAXSITE ,a." & gsCAT_TERE_LOADBOARDTYPE & " as loadboard  " & _
                     "from tbl_erun_recipe a " & _
                    "where a.docno = '" & sErunTicNO & "' " & _
                      "and a.stepno = '" & sStepNo & "' " & _
                      "and a.eqtype2 = '" & sEqType2 & "' " & _
                      "and a.deleteflag = 'N' "
              sSQL = sSQL & " and a." & gsCAT_TERE_PATH & "='" & sPath & "' "
              
            'Add by Sam Start on 20181206 for Project FT ADVAN Tester  Auto Lot Start,增加Loadboard條件
            If lblStage.Caption = "FT" And sAssignLoadboard <> "" Then
                sSQL = sSQL & " and nvl(a." & gsCAT_TERE_LOADBOARDTYPE & ",' ')= nvl('" & sAssignLoadboard & "',' ') "
            End If
            'Add by Sam End on 20181206 for Project FT ADVAN Tester  Auto Lot Start
              
        ElseIf Trim(sTecnLotID) <> "" Then
            'Modify by Sam on 20200430 for Project FT配件資料標準化,Add Loadboard
            sSQL = "select Subsystem,MAXSITE " & _
                   " ,decode(trim(" & gsCAT_TLSPES_LOADBOARDTYPE & " || " & gsCAT_TLSPES_CONTACTBOARDTYPE & " || " & gsCAT_TLSPES_BURN_IN_BOARD & "),null, " & _
                   "  fun_get_ft_acc_bypgm(" & gsCAT_TLSPES_STEPNAME & ", '" & sBodySize & "', " & gsCAT_TLSPES_PGID & ", " & gsCAT_TLSPES_PGNAME & " , 'loadboard') ,  " & gsCAT_TLSPES_LOADBOARDTYPE & " ) as loadboard  " & _
                   "from tbl_lot_step_eq_spec where " & _
                   " '" & sTecnLotID & "' like  tecnlotid and " & _
                   "stepno='" & sStepNo & "' and " & _
                   "eqtype2='" & sEqType2 & "' and " & _
                   "deleteflag ='N' "
            sSQL = sSQL & " and " & gsCAT_TLSPES_PATH & "='" & sPath & "' "
        
            'Mark by Sam start on 20200428 for Project FT配件資料標準化,不取消的話自動引用的情況會撈不出來
'            'Add by Sam Start on 20181206 for Project FT ADVAN Tester  Auto Lot Start,增加Loadboard條件
'            If lblStage.Caption = "FT" And sAssignLoadboard <> "" Then
'                sSQL = sSQL & " and nvl(" & gsCAT_TLSPES_LOADBOARDTYPE & ",' ')= nvl('" & sAssignLoadboard & "',' ') "
'            End If
'            'Add by Sam End on 20181206 for Project FT ADVAN Tester  Auto Lot Start
            'Mark by Sam End on 20200428 for Project FT配件資料標準化,不取消的話自動引用的情況會撈不出來
        Else
             'Modify by Sam on 20200430 for Project FT配件資料標準化,Add Loadboard
            sSQL = "select Subsystem,MAXSITE " & _
                   " ,decode(trim(" & gsCAT_TLSPES_LOADBOARDTYPE & " || " & gsCAT_TLSPES_CONTACTBOARDTYPE & " || " & gsCAT_TLSPES_BURN_IN_BOARD & "),null, " & _
                   "  fun_get_ft_acc_bypgm(" & gsCAT_TLSPES_STEPNAME & ", '" & sBodySize & "', " & gsCAT_TLSPES_PGID & ", " & gsCAT_TLSPES_PGNAME & " , 'loadboard') ,  " & gsCAT_TLSPES_LOADBOARDTYPE & " ) as loadboard  " & _
                   "from tbl_prod_step_eq_spec where " & _
                   "prodgroup='" & sProdGroup & "' and " & _
                   "stepno='" & sStepNo & "' and " & _
                   "eqtype2='" & sEqType2 & "' and " & _
                   "docstatus ='Active' "
            sSQL = sSQL & " and path='" & sPath & "' "
            
            'Mark by Sam start on 20200428 for Project FT配件資料標準化,不取消的話自動引用的情況會撈不出來
'            'Add by Sam Start on 20181206 for Project FT ADVAN Tester  Auto Lot Start,增加Loadboard條件
'            If lblStage.Caption = "FT" And sAssignLoadboard <> "" Then
'                sSQL = sSQL & " and nvl(" & gsCAT_PSES_LOADBOARD_TYPE & ",' ')= nvl('" & sAssignLoadboard & "',' ') "
'            End If
'            'Add by Sam End on 20181206 for Project FT ADVAN Tester  Auto Lot Start
            'Mark by Sam End on 20200428 for Project FT配件資料標準化,不取消的話自動引用的情況會撈不出來
        End If
    End If
    'add by Ernest on 2016/04/21 for ReqNo:JC201500378---end
    Me.cboSubSystem.Clear
    Set oRS = moProRawSql.QueryDatabase(sSQL)
    If oRS.Count > 0 Then
        For Each oItem In oRS
            'Add by Sam start on 20200430 for Project FT配件資料標準化
            If Me.lblStage = "FT" Then
                If sAssignLoadboard <> "" Then
                    If sAssignLoadboard = oItem.Item("loadboard") Then
                        Me.cboSubSystem.AddItem oItem.Item(1) + "," + oItem.Item(2)
                    End If
                Else
                    Me.cboSubSystem.AddItem oItem.Item(1) + "," + oItem.Item(2)
                End If
            Else
            'Add by Sam End on 20200430 for Project FT配件資料標準化
                Me.cboSubSystem.AddItem oItem.Item(1) + "," + oItem.Item(2)
            End If
            'Me.cboSubSystem.ListIndex = Me.cboSubSystem.ListCount - 1
        Next oItem
    End If
    
    'Add by Sam start on 20141128 for ReqNo:JC201400402
    If oRS.Count = 1 Then
        Me.cboSubSystem.ListIndex = 0
        Call cboSubSystem_Click
    End If
    'Add by Sam END on 20141128 for ReqNo:JC201400402
    
    'add by Sam start on 20080827 for PDA_PROJECT
'    If mbNewForm = True Then
'        Me.cboWaferID.Clear
'        sSql = "select " & gsCAT_TLI_WAFERID & " from " & gsCAT_TBL_LOT_INFO & " where " & gsCAT_TLI_LOT_ID & "='" & txtLotID.Text & "'"
'        Set oRS = moProRawSql.QueryDatabase(sSql)
'        If oRS.Count > 0 Then
'            sWaferId() = Split(oRS.Item(1).Item(gsCAT_TLI_WAFERID), ";")
'            For iIdx = LBound(sWaferId) To UBound(sWaferId)
'                cboWaferID.AddItem (sWaferId(iIdx))
'            Next
'        End If
'    End If
    'add by Sam end on 20080827 for PDA_PROJECT
    
    'Add by Sam start on 20120409 for ReqNo:JC201200096
    sSQL = "select " & gsCAT_TRCO_REASON_CODE & " " & _
           " from " & gsCAT_TBL_REASON_CODE & " where " & _
           " " & gsCAT_TRCO_CATEGORY & " = 'SetupFomReason'  and " & _
           " " & gsCAT_TRCO_DELETE_FLAG & " ='N' "
           
    Me.cboSetupReason.Clear
    Me.cboSetupReason.AddItem ""
    Set oRS = moProRawSql.QueryDatabase(sSQL)
    If oRS.Count > 0 Then
        For Each oItem In oRS
            Me.cboSetupReason.AddItem oItem.Item(gsCAT_TRCO_REASON_CODE)
        Next oItem
    End If
    'Add by Sam end on 20120409 for ReqNo:JC201200096
                    
    Call GetSetupWaferID(Me.txtLotID.Text) 'Add by Sam on 20130819 for ReqNo:JC201300238
    
    Set oRS = Nothing
    cmdOK.Enabled = True
    mbChangestep = False     'add by Ernest on 2016/04/21 for ReqNo:JC201500378
    mbChangetesterId = False 'add by Ernest on 2016/04/21 for ReqNo:JC201500378
        
    cboTesterId.Locked = True 'Add by Sam on 20190312 for Project FT ADVAN Tester,應User要求再改回直接取Assign機台
'----
' Done
'----

ExitHandler:
    ' NOTE 1:
    ' MUST CALL GetErrInfo() here first before another action
    Call GetErrInfo(msMODULE_ID, sProcID, typErrInfo, Erl)
    Call LogProcOut(msMODULE_ID, sProcID, typErrInfo, moAppLog)
    Screen.MousePointer = vbDefault
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
' Sub: FtSetupForm()
'--------------------------------------------------------------------------------
' Description:  <Type your Sub description here...>
'--------------------------------------------------------------------------------
' Author:       Jeff Weng, MXIC 2005/03/30
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
'
'================================================================================
Private Sub FtSetupForm(ByRef appXL As Object, ByRef oLot As FwLot, _
                      ByRef iPosTitle As Integer)
On Error GoTo ExitHandler:
Dim sProcID As String
Dim typErrInfo As tErrInfo
Dim sEqID   As String 'Add by Tony on 2014/03/05 for HW SETUP RECIPE COMPARE project.

'----
' Init
'----
    sProcID = "FtSetupForm"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog)

'----
' Condition Checking
'----
    ' <Put your condition checking codes here>...

'----
' Action
'----
    ' <Put your Action codes here>...

    Dim sSQL As String, oRS1 As Collection, oRs2 As Collection, oRs3 As Collection
    Dim iPosProcess As Integer, iPosComment As Integer, iPosNext As Integer
    Dim msOprId As String, oOpr As FwUser
    Dim msProdgroup As String
    Dim oEqp    As FwEquipment
    Dim msSubsystem1 As String, oSubsystem1   As FwEquipment
    Dim msSubsystem2 As String, oSubsystem2 As FwEquipment
    Dim msPGNAME             As String
    Dim msSubsystemType2     As String
    Dim iStartNum            As Integer
    Dim iEndNum              As Integer
    Dim msTemperature        As String
    Dim msPinCount           As String
    Dim msPackageCode        As String
    Dim msLoadBoardId        As String
    Dim msErunTicNo          As String
    Dim sPgName              As String
    Dim msREF_STEP_NAME_01   As String
    Dim msREF_PG_NAME_01     As String
    Dim msREPLACE_PG_NAME_01 As String
    Dim msREF_STEP_NAME_02   As String
    Dim msREF_PG_NAME_02     As String
    Dim msREPLACE_PG_NAME_02 As String
    Dim msREF_STEP_NAME_03   As String
    Dim msREF_PG_NAME_03     As String
    Dim msREPLACE_PG_NAME_03 As String
    Dim msPbFree             As String
    Dim msLoadBoardType      As String
    Dim msSubsystemType      As String
    Dim msMaxSite            As String
    Dim iIndex               As Integer
    Dim iIdx                 As Integer
    Dim sPath                As String
    Dim sMpType              As String 'Add by Tony on 20080218 for Req.M200801054
    Dim sLoadBoardId         As String 'Add by Tony on 20080218 for Req.M200801054
    Dim iPosDown            As Integer 'add by sam on 2008/05/19 for ReqNo:M200805020 ,
            
    Dim bTECNLotSpec          As Boolean  'Add by SAM ON 20100407 FOR Project By Lot TECN
    Dim sTecnLotID            As String   'Add by SAM ON 20100407 FOR Project By Lot TECN
            
    'added by Jack on 2014/01/13 for JC201400011. <Start>
    Dim sPureErunTicNo        As String
    Dim sFollowProd           As String
    'Dim sSubSystemType        As String 'Marked by Jack on 2016/03/21 for JC201500392 bug fix
    'added by Jack on 2014/01/13 for JC201400011. <End>
    'add by Ernest on 2016/04/27 for ReqNo:JC201500378-------start
    Dim sProdGroup  As String
    Dim vStepName As Variant
    Dim sStepName As String
    Dim sStepNo  As String
    Dim sEqType2 As String
    Dim oRS As Collection
    Dim sErunTicNO As String
    'add by Ernest on 2016/04/27 for ReqNo:JC201500378--------end
    'Add by Sam START on 20190628 for Project FT 配件檢核系統
    Dim sIPN As String
    Dim sKit As String
    Dim sLoadboard As String
    Dim sMatchPlate As String
    'Add by Sam END on 20190628 for Project FT 配件檢核系統
    
    'Add by Sam start on 20200427 for Project FT配件資料標準化
    Dim sContactBoard As String
    Dim sBodySize As String
    Dim colAcc As Collection
    Dim sBurnInBoard As String
    'Add by Sam end on 20200427 for Project FT配件資料標準化
    
    
    msDocNo = ""
    iPosProcess = iPosTitle + 9
    iPosComment = iPosTitle + 20
    
    'add by Sam start on 2008/5/19 for ReqNo:M200805020
    '--------------------------------------------------
    '新增欄位JUMPER及JUMPERPINNO於iPosTitle + 14,iPosTitle + 15的位置
    '原iPosTitle +14 及之後都再加上iposdown的數值
    iPosDown = 2
    'add by sam end on 2008/5/19 for ReqNo:M200805020
    
    msSubsystemType = ""
    msMaxSite = ""
    If InStr(cboSubSystem.Text, ",") > 0 Then
       msSubsystemType = Left(cboSubSystem.Text, InStr(cboSubSystem.Text, ",") - 1)
       msMaxSite = Mid$(cboSubSystem.Text, InStr(cboSubSystem.Text, ",") + 1, Len(cboSubSystem.Text) - InStr(cboSubSystem.Text, ","))
    Else
       msSubsystemType = cboSubSystem.Text
    End If
'Mark by Sam start on 2008/05/19 for ReqNo:M200805020
'    If msSubsystemType = "" Then Exit Sub
'Mark by Sam end on 2008/05/19 for ReqNo:M200805020
    With appXL
        If iPosTitle = 1 Then
            .Sheets("FT").Select
            .Cells.Select
            .selection.NumberFormatLocal = "@"
            .selection.Copy
            .Sheets(1).Select
            .ActiveSheet.Paste
            .Application.CutCopyMode = False
        Else
            .Sheets("FT").Select
            .Rows("1:49").Select
            .selection.NumberFormatLocal = "@"
            .selection.Copy
            .Sheets(1).Select
            .range("A" & CStr(iPosTitle)).Select
            .selection.Insert -4121
            .Application.CutCopyMode = False
            .range("A" & CStr(iPosTitle)).Select
            .ActiveWindow.SelectedSheets.HPageBreaks.Add appXL.activeCell

        End If
        
        Set oOpr = moFwOPR.ActiveUser
        msOprId = oOpr.UserName
        .range("C" & CStr(iPosTitle + 2)).Select
        .selection.horizontalalignment = -4131
        .activeCell.FormulaR1C1 = msOprId
        
        sSQL = "select shiftcode from tbl_mxic_emp where empno='" & msOprId & "'"
        Set oRS1 = moProRawSql.QueryDatabase(sSQL)
        If oRS1.Count > 0 Then
            .range("H" & CStr(iPosTitle + 2)).Select
            .selection.horizontalalignment = -4131
            .activeCell.FormulaR1C1 = oRS1.Item(1).Item(1)
        End If
        Set oRS1 = Nothing
        
        .range("R" & CStr(iPosTitle + 2)).Select
        .selection.horizontalalignment = -4131
        .activeCell.FormulaR1C1 = Format(Now, "YYYY/MM/DD HH:MM:SS")
        .range("T" & CStr(iPosTitle)).Select
        .activeCell.FormulaR1C1 = "SETUP FORM"
        
        'Add by Sam start on 20170317 for Project 停復測自動化,增加停測資訊
        .range("F" & CStr(iPosTitle + 4)).Select
        .activeCell.FormulaR1C1 = lblStopInfo.Caption
        'Add by Sam End on 20170317 for Project 停復測自動化
        
        .range("E" & CStr(iPosTitle + 6)).Select
        .activeCell.FormulaR1C1 = oLot.CustomAttributes(gsLOT_CUSTOMATTR_IPN)
        
        sIPN = oLot.CustomAttributes(gsLOT_CUSTOMATTR_IPN) 'Add by Sam on 20190628 for Project FT 配件檢核系統
        
        .range("E" & CStr(iPosTitle + 5)).Select
        .activeCell.FormulaR1C1 = oLot.Id
        .range("R" & CStr(iPosTitle + 5)).Select
        .activeCell.FormulaR1C1 = oLot.CustomAttributes(gsLOT_CUSTOMATTR_LOT_OWNER)
        .range("E" & CStr(iPosTitle + 7)).Select
        
        'modify by Ernest on 2016/04/27 for ReqNo:JC201500378-------start
        '.activeCell.FormulaR1C1 = oLot.CustomAttributes(gsLOT_CUSTOMATTR_STEPNAME)
        vStepName = Split(Me.cboStep.Text, " ")
        sStepName = Trim(vStepName(1))
        sStepNo = Trim(vStepName(0))
        .activeCell.FormulaR1C1 = sStepName
        'modify by Ernest on 2016/04/27 for ReqNo:JC201500378-------end
        
        .range("R" & CStr(iPosTitle + 7)).Select
        'modify by Ernest on 2016/04/21 for ReqNo:JC201500378----start
        'If oLot.CustomAttributes(gsLOT_CUSTOMATTR_CUR_EQP_ID) <> "" Then
        '    .activeCell.FormulaR1C1 = oLot.CustomAttributes(gsLOT_CUSTOMATTR_CUR_EQP_ID)
        '    ' Get TesterId Object
        '    Set oEqp = FwuRetrieveEqp(moFwMDL, oLot.CustomAttributes(gsLOT_CUSTOMATTR_CUR_EQP_ID), moAppLog)
        'Else
        '    .activeCell.FormulaR1C1 = oLot.CustomAttributes(gsLOT_CUSTOMATTR_SPLITASSIGNEQID)
        '    Set oEqp = FwuRetrieveEqp(moFwMDL, oLot.CustomAttributes(gsLOT_CUSTOMATTR_SPLITASSIGNEQID), moAppLog)
        'End If
        .activeCell.FormulaR1C1 = Me.cboTesterId.Text
        Set oEqp = FwuRetrieveEqp(moFwMDL, Me.cboTesterId.Text, moAppLog)
        'modify by Ernest on 2016/04/21 for ReqNo:JC201500378-----end
        sSQL = "select eruNticno,saprwno from tbl_lot_info where lotid='" & oLot.Id & "'"
        Set oRS1 = moProRawSql.QueryDatabase(sSQL)
        If oRS1.Count > 0 Then
            .range("R" & CStr(iPosTitle + 6)).Select
            sPureErunTicNo = oRS1.Item(1).Item("erunticno") 'added by Jack on 2014/01/13 for JC201400011.
            If oRS1.Item(1).Item("erunticno") = "" Then
                msErunTicNo = oRS1.Item(1).Item("saprwno")
                .activeCell.FormulaR1C1 = msErunTicNo
            ElseIf oRS1.Item(1).Item("saprwno") = "" Then
                msErunTicNo = oRS1.Item(1).Item("erunticno")
                .activeCell.FormulaR1C1 = msErunTicNo
            Else
                msErunTicNo = oRS1.Item(1).Item("erunticno") & ", " & oRS1.Item(1).Item("saprwno")
                .activeCell.FormulaR1C1 = msErunTicNo
            End If
        End If
        Set oRS1 = Nothing
                                      
        sBodySize = "" 'Add by Sam on 20200427 for Project FT配件資料標準化
        
        'Add BodySize by Tony on 2014/03/10 for HW SETUP RECIPE COMPARE project.
        sSQL = "select speed,code,checksum,prodgroup,packagename,pincount,packagecode,gptype,bodysize from tbl_ipn_master where ipn='" & oLot.CustomAttributes(gsLOT_CUSTOMATTR_IPN) & "'"
        
        Set oRS1 = moProRawSql.QueryDatabase(sSQL)
        If oRS1.Count > 0 Then
            
            'Add by Tony Start on 2014/03/10 for HW SETUP RECIPE COMPARE project.
            '(先存放值,以便後續使用)
            msPackageName = oRS1.Item(1).Item("packagename")
            msPinCount = oRS1.Item(1).Item("pincount")
            msBodySize = oRS1.Item(1).Item("bodysize")
            'Add by Tony Start on 2014/03/10 for HW SETUP RECIPE COMPARE project.
            
            .range("E" & CStr(iPosTitle + 9)).Select
            .activeCell.FormulaR1C1 = oRS1.Item(1).Item("speed")
            .range("R" & CStr(iPosTitle + 9)).Select
            .activeCell.FormulaR1C1 = Mid(oRS1.Item(1).Item("Code"), 3, 6)
            .range("R" & CStr(iPosTitle + 10)).Select
            .activeCell.FormulaR1C1 = oRS1.Item(1).Item("checksum")
            .range("E" & CStr(iPosTitle + 8)).Select
            .activeCell.FormulaR1C1 = oRS1.Item(1).Item("packagename")
            msProdgroup = oRS1.Item(1).Item("prodgroup")
            msPinCount = oRS1.Item(1).Item("PinCount")
            msPackageCode = oRS1.Item(1).Item("PackageCode")
            'PB Free
            If oRS1.Item(1).Item("gptype") = "XX" Then
               msPbFree = "N"
            Else
               If Len(Trim(oRS1.Item(1).Item("gptype"))) > 0 Then
                  msPbFree = "Y"
               Else
                  msPbFree = ""
               End If
            End If
            'Modify by Sam start on 20130327 for Project FVI Setup Recipe
'            'Modify by SAm on 2010624 for ReqNo:JC201000123
''            .range("R" & CStr(iPosTitle + 36 + iPosDown)).Select
'            .range("V" & CStr(iPosTitle + 39 + iPosDown)).Select
'            .activeCell.FormulaR1C1 = msPbFree
             If msPbFree = "Y" Then
                'Modify by Sam on 20190628 for Project FT 配件檢核系統
                '.range("E" & CStr(iPosTitle + 34 + iPosDown)).Select
                .range("E" & CStr(iPosTitle + 35 + iPosDown)).Select
                .activeCell.FormulaR1C1 = "▓"
            ElseIf msPbFree = "N" Then
                'Modify by Sam on 20190628 for Project FT 配件檢核系統
                '.range("G" & CStr(iPosTitle + 34 + iPosDown)).Select
                .range("G" & CStr(iPosTitle + 35 + iPosDown)).Select
                .activeCell.FormulaR1C1 = "▓"
            End If
            'Modify by Sam end on 20130327 for Project FVI Setup Recipe
                                   
        End If
        Set oRS1 = Nothing
        'modify by Ernest on 2016/04/21 for ReqNo:JC201500378----start
        ''add by Nelson start on 2007/3/20 for ReqNo:M200702009
        'sSQL = "SELECT PATH FROM TBL_ROUTE_STEP_LIST WHERE ROUTEID='" & oLot.PlanId & "' AND FWSTEPSEQ='" & oLot.CurrentStep.Steps.Item(1).Handle & "'"
        'Set oRS1 = moProRawSql.QueryDatabase(sSQL)
        'If oRS1.Count > 0 Then
        '    sPath = oRS1.Item(1).Item("PATH")
        'End If
        'Set oRS1 = Nothing
        ''add by Nelson end on 2007/3/20 for ReqNo:M200702009
        sSQL = " SELECT A." & gsCAT_TLI_TECN_LOT_ID & ", " & _
             gsCAT_TLI_ERUNTICNO & " " & _
           " FROM " & gsCAT_TBL_LOT_INFO & " A " & _
           " WHERE A." & gsCAT_TLI_LOT_ID & " ='" & txtLotID.Text & "' "
        Set oRS = moProRawSql.QueryDatabase(sSQL)
        If oRS.Count > 0 Then
            sTecnLotID = oRS.Item(1).Item(1)
            sErunTicNO = oRS.Item(1).Item(2)
        End If
        If Trim(sErunTicNO) <> "" Then
            sSQL = "SELECT " & gsCAT_TER_FOLLOW_PRODUCT & _
                      " from " & gsCAT_TBL_ERUN_REQ & _
                      " where " & gsCAT_TER_LOT_ID & " ='" & txtLotID.Text & "'" & _
                      " and " & gsCAT_TER_TICKET_NO & "='" & sErunTicNO & "' " & _
                      " and " & gsCAT_TER_STAGE & "='FT'"
            Set oRS = moProRawSql.QueryDatabase(sSQL)
            If oRS.Count > 0 Then
                sFollowProd = oRS.Item(1).Item(1)
            End If
        End If
        sBodySize = "" 'Add by Sam on 20200427 for Project FT配件資料標準化
        'Modify by Sam on 20200427 for Project FT配件資料標準化,Add BodySIze
        sSQL = "select i." & gsCAT_TIM_PROD_GROUP & ",i." & gsCAT_TIM_BODY_SIZE & " from " & gsCAT_TBL_IPN_MASTER & _
               " i,tbl_lot_attribute a where i." & gsCAT_TIM_IPN & " = a.ipn and a.lotid='" & Me.txtLotID.Text & "'"
        Set oRS = moProRawSql.QueryDatabase(sSQL)
        If oRS.Count > 0 Then
            sProdGroup = oRS.Item(1).Item("Prodgroup")
            sBodySize = oRS.Item(1).Item(gsCAT_TIM_BODY_SIZE) 'Add by Sam on 20200427 for Project FT配件資料標準化
        End If
        
        sPath = GetPath(Me.txtLotID.Text, moProRawSql, moAppLog)
        sSQL = "select distinct b.eqtype2 from view_b2b_fweqarea b where eqid='" & Me.cboTesterId.Text & "'"
        Set oRS = moProRawSql.QueryDatabase(sSQL)
        If oRS.Count > 0 Then
            sEqType2 = oRS.Item(1).Item("eqtype2")
        End If

        'modify by Ernest on 2016/04/21 for ReqNo:JC201500378-------end
        If Not oEqp Is Nothing Then
            'Get HandlerId Object
            If oEqp.CustomAttributes(gsEQP_CUSTOMATTR_SUBSYS2) = "" Then
                '.activecell.formulaR1C1 = oEqp.CustomAttributes(gsEQP_CUSTOMATTR_SUBSYS1)
                msSubsystem1 = oEqp.CustomAttributes(gsEQP_CUSTOMATTR_SUBSYS1)
'                Set oSubsystem1 = FwuRetrieveEqp(moFwMDL, msSubsystem1, moAppLog)
            Else
                msSubsystem1 = oEqp.CustomAttributes(gsEQP_CUSTOMATTR_SUBSYS1)
                msSubsystem2 = oEqp.CustomAttributes(gsEQP_CUSTOMATTR_SUBSYS2)
'                Set oSubsystem1 = FwuRetrieveEqp(moFwMDL, msSubsystem1, moAppLog)
'                Set oSubsystem2 = FwuRetrieveEqp(moFwMDL, msSubsystem2, moAppLog)
                '.activecell.formulaR1C1 = msSubsystem1 & " " & msSubsystem
            End If
            'msSubsystemType2 = oSubsystem1.CustomAttributes(gsEQP_CUSTOMATTR_EQ_TYPE2)
            '.Range("R" & CStr(iPosTitle + 8)).Select
            '.activecell.formulaR1C1 = msSubsystemType2
            
         
            'aDD BY Sam on start on 20100514 for Project Lot TECN
            bTECNLotSpec = CheckTECNLotInfo(moAppLog, moFwWIP, moFwWF, moCwMbx, oLot.Id)
            If bTECNLotSpec = True Then
                   sSQL = " SELECT A." & gsCAT_TLI_TECN_LOT_ID & " " & _
                          " FROM " & gsCAT_TBL_LOT_INFO & " A " & _
                          " WHERE A." & gsCAT_TLI_LOT_ID & " ='" & oLot.Id & "' "
                
                    Set oRS1 = moProRawSql.QueryDatabase(sSQL)
                    If oRS1.Count > 0 Then
                      sTecnLotID = oRS1.Item(1).Item(1)
                    End If
                   'modify by Ernest on 2016/04/21 for ReqNo:JC201500378-------start
                    'sSQL = "select Temperature, Comments,loadboardtype," & _
                    '        "CONTACTBOARDTYPE,KITTYPE," & _
                    '        "cabletype,pgname,Subsystem,eqtype2,'TECN' AS doctype,docno," & _
                    '        "pgid,pgname,refpgname01,refpgname02,refpgname03,refstepname01," & _
                    '        "refstepname02,refstepname03,replacepgname01,replacepgname02," & _
                    '        "replacepgname03,NEEDJUMPER,JUMPERPINNO " & _
                    '        "from tbl_lot_step_eq_spec where " & _
                    '        " '" & sTECNLotID & "' like tecnlotid and  " & _
                    '        "stepno='" & oLot.CurrentStep.Steps.Item(1).Id & "' and " & _
                    '        "eqtype2='" & oEqp.CustomAttributes(gsEQP_CUSTOMATTR_EQ_TYPE2) & "' and " & _
                    '        "subsystem ='" & msSubsystemType & "' and " & _
                    '        " deleteflag ='N' and nvl(maxsite,' ') = nvl('" & msMaxSite & "',' ')"
                    sSQL = "select Temperature, Comments,loadboardtype," & _
                            "CONTACTBOARDTYPE,KITTYPE," & _
                            "cabletype,pgname,Subsystem,eqtype2,'TECN' AS doctype,docno," & _
                            "pgid,pgname,refpgname01,refpgname02,refpgname03,refstepname01," & _
                            "refstepname02,refstepname03,replacepgname01,replacepgname02," & _
                            "replacepgname03,NEEDJUMPER,JUMPERPINNO,burninboard  " & _
                            "from tbl_lot_step_eq_spec where " & _
                            " '" & sTecnLotID & "' like tecnlotid and  " & _
                            "stepno='" & sStepNo & "' and " & _
                            "eqtype2='" & sEqType2 & "' and " & _
                            "subsystem ='" & msSubsystemType & "' and " & _
                            " deleteflag ='N' and nvl(maxsite,' ') = nvl('" & msMaxSite & "',' ')"
                    'modify by Ernest on 2016/04/21 for ReqNo:JC201500378-------end
            'Add by SAM END ON 20100514 FOR Project By Lot TECN
            Else
                
                'Modify by sam start on 2008/05/19 for ReqNO:M200805020
    '            sSQL = "select Temperature, Comments,loadboardtype," & _
    '                    "CONTACTBOARDTYPE,KITTYPE," & _
    '                    "cabletype,pgname,Subsystem,eqtype2,doctype,docno," & _
    '                    "pgid,pgname,refpgname01,refpgname02,refpgname03,refstepname01," & _
    '                    "refstepname02,refstepname03,replacepgname01,replacepgname02," & _
    '                    "replacepgname03 " & _
    '                    "from tbl_prod_step_eq_spec where " & _
    '                    "prodgroup='" & msProdgroup & "' and " & _
    '                    "stepno='" & oLot.CurrentStep.Steps.Item(1).Id & "' and " & _
    '                    "eqtype2='" & oEqp.CustomAttributes(gsEQP_CUSTOMATTR_EQ_TYPE2) & "' and " & _
    '                    "docstatus ='Active'" & " and subsystem ='" & msSubsystemType & "' and " & _
    '                    "maxsite = '" & msMaxSite & "' and path='" & sPath & "'"
                'modify by Ernest on 2016/04/21 for ReqNo:JC201500378-------start
                'sSQL = "select Temperature, Comments,loadboardtype," & _
                '        "CONTACTBOARDTYPE,KITTYPE," & _
                '        "cabletype,pgname,Subsystem,eqtype2,doctype,docno," & _
                '        "pgid,pgname,refpgname01,refpgname02,refpgname03,refstepname01," & _
                '        "refstepname02,refstepname03,replacepgname01,replacepgname02," & _
                '        "replacepgname03,NEEDJUMPER,JUMPERPINNO " & _
                '        "from tbl_prod_step_eq_spec where " & _
                '        "prodgroup='" & msProdGroup & "' and " & _
                '        "stepno='" & oLot.CurrentStep.Steps.Item(1).Id & "' and " & _
                '        "eqtype2='" & oEqp.CustomAttributes(gsEQP_CUSTOMATTR_EQ_TYPE2) & "' and " & _
                '        "docstatus ='Active'" & " and subsystem ='" & msSubsystemType & "' and " & _
                '        "maxsite = '" & msMaxSite & "' and path='" & sPath & "'"
                sSQL = "select Temperature, Comments,loadboardtype," & _
                        "CONTACTBOARDTYPE,KITTYPE," & _
                        "cabletype,pgname,Subsystem,eqtype2,doctype,docno," & _
                        "pgid,pgname,refpgname01,refpgname02,refpgname03,refstepname01," & _
                        "refstepname02,refstepname03,replacepgname01,replacepgname02," & _
                        "replacepgname03,NEEDJUMPER,JUMPERPINNO,burninboard " & _
                        "from tbl_prod_step_eq_spec where " & _
                        "prodgroup='" & sProdGroup & "' and " & _
                        "stepno='" & sStepNo & "' and " & _
                        "eqtype2='" & sEqType2 & "' and " & _
                        "docstatus ='Active'" & " and subsystem ='" & msSubsystemType & "' and " & _
                        "maxsite = '" & msMaxSite & "' and path='" & sPath & "'"

                'modify by Ernest on 2016/04/21 for ReqNo:JC201500378-------end
                'Modify by sam end on 2008/05/19 for ReqNO:M200805020
            End If
            
            Set oRS1 = moProRawSql.QueryDatabase(sSQL)
            If oRS1.Count > 0 Then
                
                msTemperature = oRS1.Item(1).Item("Temperature")
                'Add by Sam start on 20171121 for Project TECN 自動化,改成使用與停復測相同的來源
                If Trim(txtTemp.Text) <> "" Then
                    msTemperature = txtTemp.Text
                End If
                'Add by Sam END on 20171121 for Project TECN 自動化
                .range("R" & CStr(iPosTitle + 11)).Select
                '.activeCell.FormulaR1C1 = oRS1.Item(1).Item("Temperature")
                .activeCell.FormulaR1C1 = msTemperature
                
                .range("E" & CStr(iPosTitle + 14 + iPosDown)).Select
                .activeCell.FormulaR1C1 = oRS1.Item(1).Item("Comments")
                .range("R" & CStr(iPosTitle + 8)).Select
                msSubsystemType2 = oRS1.Item(1).Item("Subsystem")
                .activeCell.FormulaR1C1 = msSubsystemType2
                                
                msLoadBoardId = oRS1.Item(1).Item("Loadboardtype")
                'Add by Sam Start on 20200427 for Project FT配件資料標準化
                sContactBoard = oRS1.Item(1).Item("CONTACTBOARDTYPE")
                sBurnInBoard = oRS1.Item(1).Item("burninboard")
                If Trim(oRS1.Item(1).Item("Loadboardtype")) = "" And _
                    Trim(oRS1.Item(1).Item("CONTACTBOARDTYPE")) = "" And _
                    Trim(oRS1.Item(1).Item("burninboard")) = "" Then
                    Set colAcc = GetFTAccByPgm(moProRawSql, moAppLog, sStepName, sBodySize, oRS1.Item(1).Item("pgid"), oRS1.Item(1).Item("pgname"))
                    If colAcc.Count > 0 Then
                        msLoadBoardId = colAcc.Item("loadboard")
                        sContactBoard = colAcc.Item("contactboard")
                        sBurnInBoard = colAcc.Item("burninboard")
                    End If
                End If
                sKit = GetAccKitType(moProRawSql, moAppLog, sIPN, msLoadBoardId, sBurnInBoard, sEqType2)
                'Add by Sam End on 20200427 for Project FT配件資料標準化
                                                                
                'TECN NO
                'Modified by Jack on 2018/02/02 for Project TECN 自動化 <Start>.
                '改為DocType不為空值, 若 TECNNO有值, 要加到 DocNO前頭.
                'If Trim(oRS1.Item(1).Item("doctype")) = "TECN" Then
                If Trim(oRS1.Item(1).Item("doctype")) <> "" Then
                'Modified by Jack on 2018/02/02 for Project TECN 自動化 <End>.
                
                   .range("E" & CStr(iPosTitle + 12)).Select
                   '.activeCell.FormulaR1C1 = oRS1.Item(1).Item("docno")'Maked by Jack on 2018/02/02 for Project TECN 自動化.
                   'Modidied by Jack on 2018/02/02 for Project TECN 自動化 <Start>.
                   If msPGMTECNNo <> "" Then
                        .activeCell.FormulaR1C1 = msPGMTECNNo & "+" & oRS1.Item(1).Item("docno")
                   Else
                        .activeCell.FormulaR1C1 = oRS1.Item(1).Item("docno")
                   End If
                   'Added by Jack on 2018/02/02 for Project TECN 自動化 <End>.
                End If
                'PGM
                sPgName = oRS1.Item(1).Item("PGNAME")
                msREF_STEP_NAME_01 = oRS1.Item(1).Item("refstepname01")
                msREF_PG_NAME_01 = oRS1.Item(1).Item("refpgname01")
                msREPLACE_PG_NAME_01 = oRS1.Item(1).Item("replacepgname01")
                msREF_STEP_NAME_02 = oRS1.Item(1).Item("refstepname02")
                msREF_PG_NAME_02 = oRS1.Item(1).Item("refpgname02")
                msREPLACE_PG_NAME_02 = oRS1.Item(1).Item("replacepgname02")
                msREF_STEP_NAME_03 = oRS1.Item(1).Item("refstepname03")
                msREF_PG_NAME_03 = oRS1.Item(1).Item("refpgname03")
                msREPLACE_PG_NAME_03 = oRS1.Item(1).Item("replacepgname03")
                'modify by Ernest on 2016/04/21 for ReqNo:JC201500378-----start
                'msPGNAME = GetPGM(oLot, sPgName, msREF_STEP_NAME_01, msREF_PG_NAME_01, msREPLACE_PG_NAME_01, _
                '           msREF_STEP_NAME_02, msREF_PG_NAME_02, msREPLACE_PG_NAME_02, _
                '           msREF_STEP_NAME_03, msREF_PG_NAME_03, msREPLACE_PG_NAME_03, _
                '           oEqp.CustomAttributes(gsEQP_CUSTOMATTR_EQ_TYPE2), msSubsystemType)
                msPGNAME = GetPGM(oLot, sPgName, msREF_STEP_NAME_01, msREF_PG_NAME_01, msREPLACE_PG_NAME_01, _
                           msREF_STEP_NAME_02, msREF_PG_NAME_02, msREPLACE_PG_NAME_02, _
                           msREF_STEP_NAME_03, msREF_PG_NAME_03, msREPLACE_PG_NAME_03, _
                           sEqType2, msSubsystemType)
                
                'Add by Sam start on 20171121 for Project TECN 自動化,改成使用與停復測相同的來源
                If Trim(txtPGName.Text) <> "" Then
                    msPGNAME = txtPGName.Text
                End If
                'Add by Sam END on 20171121 for Project TECN 自動化
                
                'modify by Ernest on 2016/04/21 for ReqNo:JC201500378-------end
                .range("E" & CStr(iPosTitle + 10)).Select
                .activeCell.FormulaR1C1 = msPGNAME
                 '架機完成配件紀錄
                 
                 'add by sam start on 2008/05/19 for ReqNO:M200805020
                .range("E" & CStr(iPosTitle + 14)).Select
                .activeCell.FormulaR1C1 = oRS1.Item(1).Item("NEEDJUMPER")
                .range("E" & CStr(iPosTitle + 15)).Select
                .activeCell.FormulaR1C1 = oRS1.Item(1).Item("JUMPERPINNO")
                 'add by sam end on 2008/05/19 for ReqNO:M200805020
                 
                 'Modify by Sam start on 20100624 for ReqNo:JC201000123
'                .range("E" & CStr(iPosTitle + 24 + iPosDown)).Select
'                .activeCell.FormulaR1C1 = oRS1.Item(1).Item("Subsystem")
'                .range("R" & CStr(iPosTitle + 24 + iPosDown)).Select
'                .activeCell.FormulaR1C1 = oRS1.Item(1).Item("Subsystem")
'                .range("E" & CStr(iPosTitle + 25 + iPosDown)).Select
'                .activeCell.FormulaR1C1 = oRS1.Item(1).Item("loadboardtype")
'                sLoadBoardId = oRS1.Item(1).Item("loadboardtype") 'Add by Tony on 20080218 for Req.M200801054
'                .range("R" & CStr(iPosTitle + 25 + iPosDown)).Select
'                .activeCell.FormulaR1C1 = oRS1.Item(1).Item("loadboardtype")
'                .range("E" & CStr(iPosTitle + 26 + iPosDown)).Select
'                .activeCell.FormulaR1C1 = oRS1.Item(1).Item("cabletype")
'                .range("R" & CStr(iPosTitle + 26 + iPosDown)).Select
'                .activeCell.FormulaR1C1 = oRS1.Item(1).Item("cabletype")
'                .range("E" & CStr(iPosTitle + 27 + iPosDown)).Select
'                .activeCell.FormulaR1C1 = oRS1.Item(1).Item("CONTACTBOARDTYPE")
'                .range("R" & CStr(iPosTitle + 27 + iPosDown)).Select
'                .activeCell.FormulaR1C1 = oRS1.Item(1).Item("CONTACTBOARDTYPE")
'                .range("E" & CStr(iPosTitle + 28 + iPosDown)).Select
'                .activeCell.FormulaR1C1 = oRS1.Item(1).Item("KITTYPE")
'                .range("R" & CStr(iPosTitle + 28 + iPosDown)).Select
'                .activeCell.FormulaR1C1 = oRS1.Item(1).Item("KITTYPE")
                'Modify by Sam on 20190628 for Project FT 配件檢核系統
                '.range("E" & CStr(iPosTitle + 25 + iPosDown)).Select
                .range("E" & CStr(iPosTitle + 26 + iPosDown)).Select
                .activeCell.FormulaR1C1 = oRS1.Item(1).Item("Subsystem")
                'Modify by Sam on 20190628 for Project FT 配件檢核系統
                '.range("R" & CStr(iPosTitle + 25 + iPosDown)).Select
                .range("R" & CStr(iPosTitle + 26 + iPosDown)).Select
                .activeCell.FormulaR1C1 = oRS1.Item(1).Item("Subsystem")
                'Modify by Sam on 20190628 for Project FT 配件檢核系統
                '.range("E" & CStr(iPosTitle + 26 + iPosDown)).Select
                .range("E" & CStr(iPosTitle + 27 + iPosDown)).Select
                
                'Modify by Sam Start on 20200427 for project FT配件資料標準化
                '.activeCell.FormulaR1C1 = oRS1.Item(1).Item("loadboardtype")
'                sLoadBoardId = oRS1.Item(1).Item("loadboardtype") 'Add by Tony on 20080218 for Req.M200801054
                .activeCell.FormulaR1C1 = msLoadBoardId
                sLoadBoardId = msLoadBoardId
                'Modify by Sam End on 20200427 for project FT配件資料標準化
                
                'Modify by Sam on 20190628 for Project FT 配件檢核系統
                '.range("R" & CStr(iPosTitle + 26 + iPosDown)).Select
                .range("R" & CStr(iPosTitle + 27 + iPosDown)).Select
                
                'Modify by Sam Start on 20200427 for project FT配件資料標準化
'                .activeCell.FormulaR1C1 = oRS1.Item(1).Item("loadboardtype")
                .activeCell.FormulaR1C1 = msLoadBoardId
                'Modify by Sam End on 20200427 for project FT配件資料標準化
                
                'Modify by Sam on 20190628 for Project FT 配件檢核系統
                '.range("E" & CStr(iPosTitle + 27 + iPosDown)).Select
                .range("E" & CStr(iPosTitle + 28 + iPosDown)).Select
                .activeCell.FormulaR1C1 = oRS1.Item(1).Item("cabletype")
                'Modify by Sam on 20190628 for Project FT 配件檢核系統
                '.range("R" & CStr(iPosTitle + 27 + iPosDown)).Select
                .range("R" & CStr(iPosTitle + 28 + iPosDown)).Select
                .activeCell.FormulaR1C1 = oRS1.Item(1).Item("cabletype")
                'Modify by Sam on 20190628 for Project FT 配件檢核系統
                '.range("E" & CStr(iPosTitle + 28 + iPosDown)).Select
                .range("E" & CStr(iPosTitle + 29 + iPosDown)).Select
                
                'Modify by Sam on 20200427 for project FT配件資料標準化
                '.activeCell.FormulaR1C1 = oRS1.Item(1).Item("CONTACTBOARDTYPE")
                .activeCell.FormulaR1C1 = sContactBoard
                
                'Modify by Sam on 20190628 for Project FT 配件檢核系統
                '.range("R" & CStr(iPosTitle + 28 + iPosDown)).Select
                .range("R" & CStr(iPosTitle + 29 + iPosDown)).Select
                'Modify by Sam on 20200427 for project FT配件資料標準化
                '.activeCell.FormulaR1C1 = oRS1.Item(1).Item("CONTACTBOARDTYPE")
                .activeCell.FormulaR1C1 = sContactBoard
                
                'Modify by Sam on 20190628 for Project FT 配件檢核系統
                '.range("E" & CStr(iPosTitle + 29 + iPosDown)).Select
                .range("E" & CStr(iPosTitle + 30 + iPosDown)).Select
                'Modify by Sam on 20200427 for Project FT配件資料標準化
'                .activeCell.FormulaR1C1 = oRS1.Item(1).Item("KITTYPE")
                .activeCell.FormulaR1C1 = sKit
                'Modify by Sam on 20190628 for Project FT 配件檢核系統
                '.range("R" & CStr(iPosTitle + 29 + iPosDown)).Select
                .range("R" & CStr(iPosTitle + 30 + iPosDown)).Select
                'Modify by Sam on 20200427 for Project FT配件資料標準化
                '.activeCell.FormulaR1C1 = oRS1.Item(1).Item("KITTYPE")
                .activeCell.FormulaR1C1 = sKit
                
                'Add by Sam start on 20190628 for Project FT 配件檢核系統
'                sKit = oRS1.Item(1).Item("KITTYPE")
'                sLoadboard = oRS1.Item(1).Item("loadboardtype")  'Mark by Sam on 20200427 for Project FT配件資料標準化,上面已做
                
                'Modify by Sam on 20200427 for Project FT配件資料標準化
                'sMatchPlate = GetAccMatchPlate(moProRawSql, moAppLog, sIPN, sKit, sLoadboard)
                sMatchPlate = GetAccMatchPlate(moProRawSql, moAppLog, sIPN, sKit, msLoadBoardId, sEqType2)
                
                .range("E" & CStr(iPosTitle + 31 + iPosDown)).Select
                .activeCell.FormulaR1C1 = sMatchPlate
                .range("R" & CStr(iPosTitle + 31 + iPosDown)).Select
                .activeCell.FormulaR1C1 = sMatchPlate
                'Add by Sam end on 20190628 for Project FT 配件檢核系統
                
                'Modify by Sam end on 20100624 for ReqNo:JC201000123
'                msLoadBoardType = oRS1.Item(1).Item("loadboardtype") 'Mark by Sam on 20200427 for Project FT配件資料標準化,上面已做

            End If
            Set oRS1 = Nothing
      
            'Subsystem1,Subsystem2 Accessory Data
            'Modify by Sam on 20190628 for Project FT 配件檢核系統,Add Matchplate
            'sSQL = "select loadboardid, cableid, contactboardid, kitid from tbl_eq_info where eqid='" & oEqp.Id & "'"
            sSQL = "select loadboardid, cableid, contactboardid, kitid,matchplate from tbl_eq_info where eqid='" & oEqp.Id & "'"
            Set oRS1 = moProRawSql.QueryDatabase(sSQL)
            If oRS1.Count > 0 Then
                'Subsystem1 Accessory Data
                .range("E" & CStr(iPosTitle + 18 + iPosDown)).Select
                .activeCell.FormulaR1C1 = msSubsystem1
                
                .range("E" & CStr(iPosTitle + 19 + iPosDown)).Select
                If InStr(oRS1.Item(1).Item("loadboardid"), ",") > 0 Then
                    .activeCell.FormulaR1C1 = Left(oRS1.Item(1).Item("loadboardid"), InStr(oRS1.Item(1).Item("loadboardid"), ",") - 1)
                    msLoadBoardId = Left(oRS1.Item(1).Item("loadboardid"), InStr(oRS1.Item(1).Item("loadboardid"), ",") - 1)
                Else
                    .activeCell.FormulaR1C1 = oRS1.Item(1).Item("loadboardid")
                    msLoadBoardId = oRS1.Item(1).Item("loadboardid")
                End If
                
                .range("E" & CStr(iPosTitle + 20 + iPosDown)).Select
                If InStr(oRS1.Item(1).Item("cableid"), ",") > 0 Then
                    .activeCell.FormulaR1C1 = Left(oRS1.Item(1).Item("cableid"), InStr(oRS1.Item(1).Item("cableid"), ",") - 1)
                Else
                    .activeCell.FormulaR1C1 = oRS1.Item(1).Item("cableid")
                End If
                
                .range("E" & CStr(iPosTitle + 21 + iPosDown)).Select
                If InStr(oRS1.Item(1).Item("contactboardid"), ",") > 0 Then
                    .activeCell.FormulaR1C1 = Left(oRS1.Item(1).Item("contactboardid"), InStr(oRS1.Item(1).Item("contactboardid"), ",") - 1)
                Else
                    .activeCell.FormulaR1C1 = oRS1.Item(1).Item("contactboardid")
                End If
                
                .range("E" & CStr(iPosTitle + 22 + iPosDown)).Select
                If InStr(oRS1.Item(1).Item("kitid"), ",") > 0 Then
                    .activeCell.FormulaR1C1 = Left(oRS1.Item(1).Item("kitid"), InStr(oRS1.Item(1).Item("kitid"), ",") - 1)
                Else
                    .activeCell.FormulaR1C1 = oRS1.Item(1).Item("kitid")
                End If
                
                'Add by Sam start on 20190628 for Project FT 配件檢核系統
                .range("E" & CStr(iPosTitle + 23 + iPosDown)).Select
                If InStr(oRS1.Item(1).Item("matchplate"), ",") > 0 Then
                    .activeCell.FormulaR1C1 = Left(oRS1.Item(1).Item("matchplate"), InStr(oRS1.Item(1).Item("matchplate"), ",") - 1)
                Else
                    .activeCell.FormulaR1C1 = oRS1.Item(1).Item("matchplate")
                End If
                'Add by Sam end on 20190628 for Project FT 配件檢核系統
                
                'Add by Tony start on 20080130 for Req.M200801054
                '*****
                '架機完成配件紀錄
                '*****
                'Modify by Sam on 20100624 for ReqNo:JC201000123
                'Mark by Sam start on 20190628 for Project FT 配件檢核系統
'                .range("E" & CStr(iPosTitle + 30 + iPosDown)).Select
'                If sLoadBoardId <> "" Then
'                    sSQL = "Select " & gsCAT_TAB_MPTYPE & vbNewLine & _
'                           " from " & gsCAT_TBL_ACC_BASIC & vbNewLine & _
'                           " where " & gsCAT_TAB_ACC_NAME & "='" & sLoadBoardId & "'" & vbNewLine & _
'                           " group by " & gsCAT_TAB_MPTYPE
'                    Set oRs3 = moProRawSql.QueryDatabase(sSQL)
'                    If oRs3.Count > 0 Then
'                        For iIndex = 1 To oRs3.Count
'                            If oRs3.Item(iIndex).Item(1) = "A" Then sMpType = "A"
'                            If oRs3.Item(iIndex).Item(1) = "AD" Then sMpType = sMpType & "AD"
'                            If oRs3.Item(iIndex).Item(1) = "" Then sMpType = sMpType & "NULL"
'                        Next iIndex
'
'                        Select Case (sMpType)
'                        Case "A"
'                            .activeCell.FormulaR1C1 = Space(11) & "▓ A" & Space(16) & "□ AD" & Space(13) & "□ NA"
'                        Case "AD"
'                            .activeCell.FormulaR1C1 = Space(11) & "□ A" & Space(16) & "▓ AD" & Space(13) & "□ NA"
'                        Case "NULL"
'                            .activeCell.FormulaR1C1 = Space(11) & "□ A" & Space(16) & "□ AD" & Space(13) & "▓ NA"
'                        Case "AAD"
'                            .activeCell.FormulaR1C1 = Space(11) & "▓ A" & Space(16) & "▓ AD" & Space(13) & "□ NA"
'                        Case "ANULL"
'                            .activeCell.FormulaR1C1 = Space(11) & "▓ A" & Space(16) & "□ AD" & Space(13) & "▓ NA"
'                        Case "ADNULL"
'                            .activeCell.FormulaR1C1 = Space(11) & "□ A" & Space(16) & "▓ AD" & Space(13) & "▓ NA"
'                        Case "AADNULL"
'                            .activeCell.FormulaR1C1 = Space(11) & "▓ A" & Space(16) & "▓ AD" & Space(13) & "▓ NA"
'
'                        End Select
'                    Else
'                        .activeCell.FormulaR1C1 = Space(11) & "□ A" & Space(16) & "□ AD" & Space(13) & "▓ NA"
'                    End If
'
'                Else
'                    .activeCell.FormulaR1C1 = Space(11) & "□ A" & Space(16) & "□ AD" & Space(13) & "▓ NA"
'                End If
                'Mark by Sam End on 20190628 for Project FT 配件檢核系統
                'Add by Tony end on 20080130 for Req.M200801054
                
            'End If
            'Set oRS1 = Nothing
            
            'Subsystem2 Accessory Data
            'sSql = "select loadboardid, cableid, contactboardid, kitid from tbl_eq_info where eqid='" & oEqp.Id & "'"
            'Set oRS1 = moProRawSql.QueryDatabase(sSql)
            'If oRS1.Count > 0 Then
                .range("R" & CStr(iPosTitle + 18 + iPosDown)).Select
                .activeCell.FormulaR1C1 = msSubsystem2
                
                .range("R" & CStr(iPosTitle + 19 + iPosDown)).Select
                If InStr(oRS1.Item(1).Item("loadboardid"), ",") > 0 Then
                    iStartNum = 0: iEndNum = 0
                    iStartNum = InStr(oRS1.Item(1).Item("loadboardid"), ",") + 1
                    If InStr(iStartNum, oRS1.Item(1).Item("loadboardid"), ",") > 0 Then
                        iEndNum = InStr(iStartNum, oRS1.Item(1).Item("loadboardid"), ",")
                    Else
                        iEndNum = Len(oRS1.Item(1).Item("loadboardid")) + 1
                    End If
                    .activeCell.FormulaR1C1 = Mid$(oRS1.Item(1).Item("loadboardid"), iStartNum, iEndNum - iStartNum)
                Else
                    '.activecell.formulaR1C1 = oRS1.Item(1).Item("loadboardid")
                End If
                
                .range("R" & CStr(iPosTitle + 20 + iPosDown)).Select
                If InStr(oRS1.Item(1).Item("cableid"), ",") > 0 Then
                    iStartNum = 0: iEndNum = 0
                    iStartNum = InStr(oRS1.Item(1).Item("cableid"), ",") + 1
                    If InStr(iStartNum, oRS1.Item(1).Item("cableid"), ",") > 0 Then
                        iEndNum = InStr(iStartNum, oRS1.Item(1).Item("cableid"), ",")
                    Else
                        iEndNum = Len(oRS1.Item(1).Item("cableid")) + 1
                    End If
                    .activeCell.FormulaR1C1 = Mid$(oRS1.Item(1).Item("cableid"), iStartNum, iEndNum - iStartNum)
                Else
                    '.activecell.formulaR1C1 = oRS1.Item(1).Item("cableid")
                End If
                
                .range("R" & CStr(iPosTitle + 21 + iPosDown)).Select
                If InStr(oRS1.Item(1).Item("contactboardid"), ",") > 0 Then
                    iStartNum = 0: iEndNum = 0
                    iStartNum = InStr(oRS1.Item(1).Item("contactboardid"), ",") + 1
                    If InStr(iStartNum, oRS1.Item(1).Item("contactboardid"), ",") > 0 Then
                        iEndNum = InStr(iStartNum, oRS1.Item(1).Item("contactboardid"), ",")
                    Else
                        iEndNum = Len(oRS1.Item(1).Item("contactboardid")) + 1
                    End If
                    .activeCell.FormulaR1C1 = Mid$(oRS1.Item(1).Item("contactboardid"), iStartNum, iEndNum - iStartNum)
                Else
                    '.activecell.formulaR1C1 = oRS1.Item(1).Item("contactboardid")
                End If
                
                .range("R" & CStr(iPosTitle + 22 + iPosDown)).Select
                If InStr(oRS1.Item(1).Item("kitid"), ",") > 0 Then
                    iStartNum = 0: iEndNum = 0
                    iStartNum = InStr(oRS1.Item(1).Item("kitid"), ",") + 1
                    If InStr(iStartNum, oRS1.Item(1).Item("kitid"), ",") > 0 Then
                        iEndNum = InStr(iStartNum, oRS1.Item(1).Item("kitid"), ",")
                    Else
                        iEndNum = Len(oRS1.Item(1).Item("kitid")) + 1
                    End If
                    .activeCell.FormulaR1C1 = Mid$(oRS1.Item(1).Item("kitid"), iStartNum, iEndNum - iStartNum)
                Else
                    '.activecell.formulaR1C1 = oRS1.Item(1).Item("kitid")
                End If
                
                'Add by Sam start on 20190628 for Project FT 配件檢核系統
                .range("R" & CStr(iPosTitle + 23 + iPosDown)).Select
                If InStr(oRS1.Item(1).Item("matchplate"), ",") > 0 Then
                    iStartNum = 0: iEndNum = 0
                    iStartNum = InStr(oRS1.Item(1).Item("matchplate"), ",") + 1
                    If InStr(iStartNum, oRS1.Item(1).Item("matchplate"), ",") > 0 Then
                        iEndNum = InStr(iStartNum, oRS1.Item(1).Item("matchplate"), ",")
                    Else
                        iEndNum = Len(oRS1.Item(1).Item("matchplate")) + 1
                    End If
                    .activeCell.FormulaR1C1 = Mid$(oRS1.Item(1).Item("matchplate"), iStartNum, iEndNum - iStartNum)
                End If
                'Add by Sam end on 20190628 for Project FT 配件檢核系統
                
            End If
            Set oRS1 = Nothing
            
            ' Get Soaktime
            .range("R" & CStr(iPosTitle + 12)).Select
            If msTemperature = "25" Or msTemperature = "30" Or UCase(msTemperature) = "ROOM  TEMP" Or _
                UCase(msTemperature) = "ROM TEMP" Or UCase(msTemperature) = "ROOM REMP" Or _
                UCase(msTemperature) = "ROOMTEMP" Or UCase(msTemperature) = "ROOT TEMP" Or _
                UCase(msTemperature) = "ROOM TEMP." Or UCase(msTemperature) = "ROMM TEMP" Or _
                UCase(msTemperature) = "ROOM TEMP" Then
                .activeCell.FormulaR1C1 = 0
            Else
                sSQL = "select soaktime from tbl_subsys_confi where eqtype2='" & msSubsystemType2 & "' and pincount='" & msPinCount & "' and packagecode='" & msPackageCode & "'"
                Set oRS1 = moProRawSql.QueryDatabase(sSQL)
                If oRS1.Count > 0 Then
                    .activeCell.FormulaR1C1 = oRS1.Item(1).Item("soaktime")
                End If
                Set oRS1 = Nothing
                
            End If
            
            'Sthand File
            sSQL = "select sthand_file from tbl_acc_group where loadboard='" & msLoadBoardId & "'"
            Set oRS1 = moProRawSql.QueryDatabase(sSQL)
            If oRS1.Count > 0 Then
                .range("E" & CStr(iPosTitle + 11)).Select
                .activeCell.FormulaR1C1 = oRS1.Item(1).Item("sthand_file")
            End If
            Set oRS1 = Nothing
            
            'Step comments, good bin, dgrade bin
            Dim msGood As String
            Dim msDgrade As String
            Dim msBin As String
            Dim msPath As String

            msGood = ""
            msDgrade = ""
            
            sSQL = "select route " & _
                   " from fwadmin.tbl_lot_attribute where lotid='" & txtLotID.Text & "'"
            Set oRS1 = moProRawSql.QueryDatabase(sSQL)
            If oRS1.Count > 0 Then
               msPath = oRS1.Item(1).Item("route")
            End If
            Set oRS1 = Nothing


            'Add by SAM START ON 20100512 FOR Project By Lot TECN
            bTECNLotSpec = CheckTECNLotInfo(moAppLog, moFwWIP, moFwWF, moCwMbx, oLot.Id)
            If bTECNLotSpec = True Then
                'modify by Ernest on 2016/04/21 for ReqNo:JC201500378-------start
                'sSQL = "select Comments,'TECN' AS  doctype,docno ,bin1,bin2,bin3,bin4,bin5,bin6,bin7,bin8 " & _
                '       "from tbl_Lot_step_spec where " & _
                '       " '" & sTECNLotID & "' like tecnlotid and  " & _
                '       " stepno='" & oLot.CurrentStep.Steps.Item(1).Id & "' and " & _
                '       " deleteflag='N' and path = '" & msPath & "'"
                sSQL = "select Comments,'TECN' AS  doctype,docno ,bin1,bin2,bin3,bin4,bin5,bin6,bin7,bin8 " & _
                       "from tbl_Lot_step_spec where " & _
                       " '" & sTecnLotID & "' like tecnlotid and  " & _
                       " stepno='" & sStepNo & "' and " & _
                       " deleteflag='N' and path = '" & msPath & "'"
       
                'modify by Ernest on 2016/04/21 for ReqNo:JC201500378---------end
            Else
            'Add by SAM End ON 20100512 FOR Project By Lot TECN

                'modify by Ernest on 2016/04/21 for ReqNo:JC201500378-------start
                'sSQL = "select Comments,doctype,docno,bin1,bin2,bin3,bin4,bin5,bin6,bin7,bin8 " & _
                '       "from tbl_prod_step_spec where " & _
                '       "prodgroup='" & msProdGroup & "' and " & _
                '       "stepno='" & oLot.CurrentStep.Steps.Item(1).Id & "' and " & _
                '       "docstatus ='Active' and path = '" & msPath & "'"
                sSQL = "select Comments,doctype,docno,bin1,bin2,bin3,bin4,bin5,bin6,bin7,bin8 " & _
                       "from tbl_prod_step_spec where " & _
                       "prodgroup='" & sProdGroup & "' and " & _
                       "stepno='" & sStepNo & "' and " & _
                       "docstatus ='Active' and path = '" & msPath & "'"
                'modify by Ernest on 2016/04/21 for ReqNo:JC201500378-------start
            End If
            Set oRS1 = moProRawSql.QueryDatabase(sSQL)
            
            If oRS1.Count > 0 Then
               .range("E" & CStr(iPosTitle + 15 + iPosDown)).Select
                .activeCell.FormulaR1C1 = oRS1.Item(1).Item("Comments")
            'Mark by Sam start on 20160608 for Project FT DGRADE,取消BIN1~BIN8
'               For iIndex = 1 To 8
'                 msBin = "BIN" + CStr(iIndex)
'                 If oRS1.Item(1).Item(msBin) = "Good" Then
'                    If msGood = "" Then
'                       msGood = CStr(iIndex)
'                    Else
'                       msGood = msGood + "," + CStr(iIndex)
'                    End If
'                 End If
'                 If oRS1.Item(1).Item(msBin) = "Dgrade" Then
'                    If msDgrade = "" Then
'                       msDgrade = CStr(iIndex)
'                    Else
'                       msDgrade = msDgrade + "," + CStr(iIndex)
'                    End If
'                 End If
'               Next
'               .range("E" & CStr(iPosTitle + 13)).Select
'                .activeCell.FormulaR1C1 = msGood
'               .range("R" & CStr(iPosTitle + 13)).Select
'                .activeCell.FormulaR1C1 = msDgrade
            'Mark by Sam END on 20160608 for Project FT DGRADE
            End If
            Set oRS1 = Nothing

            'Clean Dut. Request
            Dim msCleanDut As String
            Dim msLastIPN  As String
            Dim msLastPbFree As String
            msCleanDut = ""
            msLastIPN = ""
            msLastPbFree = ""
            If msPbFree = "N" Then
               msCleanDut = "N"
            Else
               If Len(msPbFree) <= 0 Then
                  msCleanDut = "Current Lot or pre-Lot's PbFree is NULL!!"
               Else
                  msLastIPN = oEqp.CustomAttributes(gsEQP_CUSTOMATTR_LAST_IPN)
                  sSQL = "select gptype from tbl_ipn_master where ipn='" & msLastIPN & "'"
                  Set oRS1 = moProRawSql.QueryDatabase(sSQL)
                  If oRS1.Count > 0 Then
                     If oRS1.Item(1).Item("gptype") = "XX" Then
                        msLastPbFree = "N"
                     Else
                        If Len(Trim(oRS1.Item(1).Item("gptype"))) > 0 Then
                           msLastPbFree = "Y"
                        Else
                           msLastPbFree = ""
                        End If
                     End If
                  End If
                  Set oRS1 = Nothing
                  If msLastPbFree = "N" Then
                     msCleanDut = "Y"
                  ElseIf msLastPbFree = "Y" And InStr(msLoadBoardId, msLoadBoardType) = 0 Then
                     msCleanDut = "Y"
                  ElseIf msLastPbFree = "Y" And InStr(msLoadBoardId, msLoadBoardType) > 0 Then
                     msCleanDut = "N"
                  Else
                     msCleanDut = "Current Lot or pre-Lot's PbFree is NULL!!"
                  End If
               End If
            End If
            
            'Modify by Sam START on 20130327 for Project FVI Setup Recipe
'            'Modify by Sam on 20100624 for ReqNo:JC201000123
'            '.range("R" & CStr(iPosTitle + 37 + iPosDown)).Select
'            .range("V" & CStr(iPosTitle + 40 + iPosDown)).Select
'            .activeCell.FormulaR1C1 = msCleanDut
            If msCleanDut = "Y" Then
                'Modify by Sam on 20190628 for Project FT 配件檢核系統
                '.range("E" & CStr(iPosTitle + 35 + iPosDown)).Select
                .range("E" & CStr(iPosTitle + 36 + iPosDown)).Select
                .activeCell.FormulaR1C1 = "▓"
            ElseIf msCleanDut = "N" Then
                'Modify by Sam on 20190628 for Project FT 配件檢核系統
'                .range("G" & CStr(iPosTitle + 35 + iPosDown)).Select
                .range("G" & CStr(iPosTitle + 36 + iPosDown)).Select
                .activeCell.FormulaR1C1 = "▓"
            Else
                'Modify by Sam on 20190628 for Project FT 配件檢核系統
                '.range("I" & CStr(iPosTitle + 35 + iPosDown)).Select
                .range("I" & CStr(iPosTitle + 36 + iPosDown)).Select
                .activeCell.FormulaR1C1 = msCleanDut
            End If
            'Modify by Sam END on 20130327 for Project FVI Setup Recipe
            
        End If
        
        'added by Jack on 2014/01/13 for JC201400011. <Start>
        If Len(sPureErunTicNo) <> 0 Then
            sSQL = "SELECT " & gsCAT_TER_FOLLOW_PRODUCT & _
                      " from " & gsCAT_TBL_ERUN_REQ & _
                      " where " & gsCAT_TER_LOT_ID & " ='" & oLot.Id & "'" & _
                      " and " & gsCAT_TER_TICKET_NO & "='" & sPureErunTicNo & "' " & _
                      " and " & gsCAT_TER_STAGE & "='" & lblStage.Caption & "' " & _
                      " and " & gsCAT_TER_DELETE_FLAG & "='N' "
            Set oRS1 = moProRawSql.QueryDatabase(sSQL)
            If oRS1.Count > 0 Then
                sFollowProd = oRS1.Item(1).Item(1)
            End If
            
            
            If Trim(sPureErunTicNo) <> "" And Trim(sFollowProd) = "N" Then
                '<1> replace speed/code/checksum.
                sSQL = "SELECT " & gsCAT_TER_SPEED & "," & _
                                   gsCAT_TER_CHECKSUM & "," & _
                                   gsCAT_TER_CODE & " " & _
                          " from " & gsCAT_TBL_ERUN_REQ & _
                          " where " & gsCAT_TER_LOT_ID & " ='" & oLot.Id & "'" & _
                          " and " & gsCAT_TER_TICKET_NO & "='" & sPureErunTicNo & "' " & _
                          " and " & gsCAT_TER_STAGE & "='" & lblStage.Caption & "' " & _
                          " and " & gsCAT_TER_DELETE_FLAG & "='N' "
                Set oRS1 = moProRawSql.QueryDatabase(sSQL)
                If oRS1.Count > 0 Then
                    .range("E" & CStr(iPosTitle + 9)).Select 'speed
                    .activeCell.FormulaR1C1 = oRS1.Item(1).Item("speed")
                    
                    .range("R" & CStr(iPosTitle + 9)).Select 'Code
                    .activeCell.FormulaR1C1 = oRS1.Item(1).Item("code")
                
                    .range("R" & CStr(iPosTitle + 10)).Select 'checksum
                    .activeCell.FormulaR1C1 = oRS1.Item(1).Item("checksum")
                End If
                
                '<2> replace pgname/goodbin/jumperpinno/subsystem/temperature/loadboardtype/contactboardtype
                'modify by Ernest on 20160127 for ReqNo:JC201500378-------start
                'sSQL = "select " & gsCAT_TERE_PGNAME & ", " & _
                '                   gsCAT_TERE_GOODBIN & ", " & _
                '                   gsCAT_TERE_JUMPERPINNO & ", " & _
                '                   gsCAT_TERE_TEMPERATURE & ", " & _
                '                   gsCAT_TERE_LOADBOARDTYPE & ", " & _
                '                   gsCAT_TERE_SUBSYSTEM & ", " & _
                '                   gsCAT_TERE_CONTACTBOARDTYPE & " " & _
                '         "from " & gsCAT_TBL_ERUN_RECIPE & " a " & _
                '        "where a." & gsCAT_TERE_DOCNO & " = '" & sPureErunTicNo & "' " & _
                '          "and a." & gsCAT_TERE_STEPNO & " = '" & oLot.CurrentStep.Steps.Item(1).Id & "' " & _
                '          "and a." & gsCAT_TERE_EQTYPE2 & " = '" & oEqp.CustomAttributes(gsEQP_CUSTOMATTR_EQ_TYPE2) & "' "
                sSQL = "select " & gsCAT_TERE_PGNAME & ", " & _
                                   gsCAT_TERE_GOODBIN & ", " & _
                                   gsCAT_TERE_JUMPERPINNO & ", " & _
                                   gsCAT_TERE_TEMPERATURE & ", " & _
                                   gsCAT_TERE_LOADBOARDTYPE & ", " & _
                                   gsCAT_TERE_SUBSYSTEM & ", " & _
                                   gsCAT_TERE_CONTACTBOARDTYPE & " " & _
                         "from " & gsCAT_TBL_ERUN_RECIPE & " a " & _
                        "where a." & gsCAT_TERE_DOCNO & " = '" & sPureErunTicNo & "' " & _
                          "and a." & gsCAT_TERE_STEPNO & " = '" & sStepNo & "' " & _
                          "and a." & gsCAT_TERE_EQTYPE2 & " = '" & sEqType2 & "' "

                'modify by Ernest on 20160127 for ReqNo:JC201500378---------end
                'Modified by Jack on 2016/03/21 for JC201500392 <Start>
                'bug fix :sSubSystemType --> msSubSystemType.
                'If Trim(sSubSystemType) = "" Then
                '    sSQL = sSQL & " and a." & gsCAT_TERE_SUBSYSTEM & " is null "
                'Else
                '    sSQL = sSQL & " and a." & gsCAT_TERE_SUBSYSTEM & "='" & Trim(sSubSystemType) & "' "
                'End If
                If Trim(msSubsystemType) = "" Then
                    sSQL = sSQL & " and a." & gsCAT_TERE_SUBSYSTEM & " is null "
                Else
                    sSQL = sSQL & " and a." & gsCAT_TERE_SUBSYSTEM & "='" & Trim(msSubsystemType) & "' "
                End If
                'Modified by Jack on 2016/03/21 for JC201500392 <End>
                
                sSQL = sSQL & " and a.deleteflag = 'N' "
                          
                Set oRS1 = moProRawSql.QueryDatabase(sSQL)
                If oRS1.Count > 0 Then
                    .range("E" & CStr(iPosTitle + 10)).Select 'Program Name:
                    .activeCell.FormulaR1C1 = oRS1.Item(1).Item("pgname")
                    
                    .range("E" & CStr(iPosTitle + 13)).Select 'GoodBin
                    .activeCell.FormulaR1C1 = oRS1.Item(1).Item("goodbin")
                    
                    .range("E" & CStr(iPosTitle + 15)).Select 'jumperpinno
                    .activeCell.FormulaR1C1 = oRS1.Item(1).Item("jumperpinno")
                    
                    .range("E" & CStr(iPosTitle + 15)).Select 'jumperpinno
                    .activeCell.FormulaR1C1 = oRS1.Item(1).Item("jumperpinno")
                    
                    'Modify by Sam on 20190628 for Project FT 配件檢核系統
                    '.range("E" & CStr(iPosTitle + 25 + iPosDown)).Select 'subsystem1
                    .range("E" & CStr(iPosTitle + 26 + iPosDown)).Select 'subsystem1
                    .activeCell.FormulaR1C1 = oRS1.Item(1).Item("subsystem")
                    
                    'Modify by Sam on 20190628 for Project FT 配件檢核系統
                    '.range("R" & CStr(iPosTitle + 25 + iPosDown)).Select 'subsystem2
                    .range("R" & CStr(iPosTitle + 26 + iPosDown)).Select 'subsystem2
                    .activeCell.FormulaR1C1 = oRS1.Item(1).Item("subsystem")
                                        
                    .range("R" & CStr(iPosTitle + 8)).Select 'subsystem
                    .activeCell.FormulaR1C1 = oRS1.Item(1).Item("subsystem")
                    
                    .range("R" & CStr(iPosTitle + 11)).Select 'Temperature
                    .activeCell.FormulaR1C1 = oRS1.Item(1).Item("temperature")
                    
                    sLoadBoardId = oRS1.Item(1).Item("loadboardtype")
                    'Modify by Sam on 20190628 for Project FT 配件檢核系統
                    '.range("E" & CStr(iPosTitle + 26 + iPosDown)).Select
                    .range("E" & CStr(iPosTitle + 27 + iPosDown)).Select
                    .activeCell.FormulaR1C1 = oRS1.Item(1).Item("loadboardtype")
                    
                    'Modify by Sam on 20190628 for Project FT 配件檢核系統
                    '.range("R" & CStr(iPosTitle + 26 + iPosDown)).Select
                    .range("R" & CStr(iPosTitle + 27 + iPosDown)).Select
                    .activeCell.FormulaR1C1 = oRS1.Item(1).Item("loadboardtype")
                    
                    '*****
                    '架機完成配件紀錄
                    '*****
                    'Mark by Sam start on 20190628 for Project FT 配件檢核系統
'                    .range("E" & CStr(iPosTitle + 30 + iPosDown)).Select
'
'                    If sLoadBoardId <> "" Then
'                        sSQL = "Select " & gsCAT_TAB_MPTYPE & vbNewLine & _
'                               " from " & gsCAT_TBL_ACC_BASIC & vbNewLine & _
'                               " where " & gsCAT_TAB_ACC_NAME & "='" & sLoadBoardId & "'" & vbNewLine & _
'                               " group by " & gsCAT_TAB_MPTYPE
'                        Set oRs3 = moProRawSql.QueryDatabase(sSQL)
'                        If oRs3.Count > 0 Then
'                            For iIndex = 1 To oRs3.Count
'                                If oRs3.Item(iIndex).Item(1) = "A" Then sMpType = "A"
'                                If oRs3.Item(iIndex).Item(1) = "AD" Then sMpType = sMpType & "AD"
'                                If oRs3.Item(iIndex).Item(1) = "" Then sMpType = sMpType & "NULL"
'                            Next iIndex
'
'                            Select Case (sMpType)
'                            Case "A"
'                                .activeCell.FormulaR1C1 = Space(11) & "▓ A" & Space(16) & "□ AD" & Space(13) & "□ NA"
'                            Case "AD"
'                                .activeCell.FormulaR1C1 = Space(11) & "□ A" & Space(16) & "▓ AD" & Space(13) & "□ NA"
'                            Case "NULL"
'                                .activeCell.FormulaR1C1 = Space(11) & "□ A" & Space(16) & "□ AD" & Space(13) & "▓ NA"
'                            Case "AAD"
'                                .activeCell.FormulaR1C1 = Space(11) & "▓ A" & Space(16) & "▓ AD" & Space(13) & "□ NA"
'                            Case "ANULL"
'                                .activeCell.FormulaR1C1 = Space(11) & "▓ A" & Space(16) & "□ AD" & Space(13) & "▓ NA"
'                            Case "ADNULL"
'                                .activeCell.FormulaR1C1 = Space(11) & "□ A" & Space(16) & "▓ AD" & Space(13) & "▓ NA"
'                            Case "AADNULL"
'                                .activeCell.FormulaR1C1 = Space(11) & "▓ A" & Space(16) & "▓ AD" & Space(13) & "▓ NA"
'
'                            End Select
'                        Else
'                            .activeCell.FormulaR1C1 = Space(11) & "□ A" & Space(16) & "□ AD" & Space(13) & "▓ NA"
'                        End If
'
'                    Else
'                        .activeCell.FormulaR1C1 = Space(11) & "□ A" & Space(16) & "□ AD" & Space(13) & "▓ NA"
'                    End If
                    'Mark by Sam End on 20190628 for Project FT 配件檢核系統
                    'Modify by Sam on 20190628 for Project FT 配件檢核系統
                    '.range("E" & CStr(iPosTitle + 28 + iPosDown)).Select 'CONTACTBOARDTYPE
                    .range("E" & CStr(iPosTitle + 29 + iPosDown)).Select 'CONTACTBOARDTYPE
                    .activeCell.FormulaR1C1 = oRS1.Item(1).Item("contactboardtype")
                    
                    'Modify by Sam on 20190628 for Project FT 配件檢核系統
                    '.range("R" & CStr(iPosTitle + 28 + iPosDown)).Select 'CONTACTBOARDTYPE
                    .range("R" & CStr(iPosTitle + 29 + iPosDown)).Select 'CONTACTBOARDTYPE
                    .activeCell.FormulaR1C1 = oRS1.Item(1).Item("contactboardtype")
                End If
            End If
        End If
        'added by Jack on 2014/01/13 for JC201400011. <End>
        
        Set oRS1 = Nothing

        'End If

'         msPGNAME = GetPGM(oEqp.Id, oLot, msSubsystemType2, msDocNo, gsLOTUNIT_CHIP)
'        .Range("E" & CStr(iPosTitle + 10)).Select
'        .activecell.formulaR1C1 = msPGNAME
'        .Range("E" & CStr(iPosTitle + 12)).Select
'        .activecell.formulaR1C1 = msDocNo
        
        ' if msErunTicNo is not Enpity表重工品,則
        '設為空值 : ProgramName, DocumentNo, Temperature, subsys, Special Comment, Speed, Code, CheckSum
        'user直接查詢委測/重工單, 避免混淆
'Marked by Jack on 2014/01/13 for JC201400011 <Start>
'Chihhao要求不清空以下欄位內容.
'        If Len(msErunTicNo) <> 0 Then
'            .range("E" & CStr(iPosTitle + 10)).Select
'            .activeCell.FormulaR1C1 = ""
'            .range("E" & CStr(iPosTitle + 12)).Select
'            .activeCell.FormulaR1C1 = ""
'            .range("R" & CStr(iPosTitle + 11)).Select
'            .activeCell.FormulaR1C1 = ""
'            .range("R" & CStr(iPosTitle + 8)).Select
'            .activeCell.FormulaR1C1 = ""
'            .range("E" & CStr(iPosTitle + 16)).Select
'            .activeCell.FormulaR1C1 = ""
'            .range("E" & CStr(iPosTitle + 9)).Select
'            .activeCell.FormulaR1C1 = ""
'            .range("R" & CStr(iPosTitle + 9)).Select
'            .activeCell.FormulaR1C1 = ""
'            .range("R" & CStr(iPosTitle + 10)).Select
'            .activeCell.FormulaR1C1 = ""
'        End If
        
        'Add by Sam start on 20130327 for Project FVI Setup Recipe
         sSQL = "select " & gsCAT_TTFA_PITCH & "," & gsCAT_TTFA_VACUUMCUP & " from " & gsCAT_TBL_TESTER_FVI_ACC & _
                " where " & gsCAT_TTFA_PINCOUNT & "='" & msPinCount & "' " & _
                " and " & gsCAT_TTFA_PKGCODE & "='" & msPackageCode & "' " & _
                " and " & gsCAT_TTFA_DELETEFLAG & "='N' "
                
         Set oRS1 = moProRawSql.QueryDatabase(sSQL)
         If oRS1.Count > 0 Then
            'Modify by Sam on 20190628 for Project FT 配件檢核系統
'            .range("E" & CStr(iPosTitle + 32 + iPosDown)).Select
            .range("E" & CStr(iPosTitle + 33 + iPosDown)).Select
            .activeCell.FormulaR1C1 = oRS1.Item(1).Item(gsCAT_TTFA_PITCH)
            'Modify by Sam on 20190628 for Project FT 配件檢核系統
            '.range("E" & CStr(iPosTitle + 33 + iPosDown)).Select
            .range("E" & CStr(iPosTitle + 34 + iPosDown)).Select
            .activeCell.FormulaR1C1 = oRS1.Item(1).Item(gsCAT_TTFA_VACUUMCUP)
         End If
        'Add by Sam end on 20130327 for Project FVI Setup Recipe
        
        'Add by Tony Start on 2014/02/25 for HW SETUP RECIPE COMPARE project
        '(二)設備架機確認(HW Setup Quality check)內容
        'Modify in 2016/04/27 for ReqNO:JC201500378----start
        'sEqID = lblTesterId.Caption
        sEqID = Me.cboTesterId.Text
        'Modify in 2016/04/27 for ReqNO:JC201500378----start
        Call getRecipeSpecName(sEqID, oLot.Id, oLot.CustomAttributes(gsLOT_CUSTOMATTR_IPN), _
                            msPackageName, msPinCount, msBodySize, _
                            oRS1, moProRawSql, moAppLog, _
                            moCwMbx)
        
        If Not oRS1 Is Nothing And oRS1.Count > 0 Then
            For iIdx = 1 To oRS1.Count
                If iIdx <= miFT_RECIPE_ROW Then
                     'Modify by Sam on 20190628 for Project FT 配件檢核系統
                    '.range("A" & CStr(iPosTitle + 48 + iIdx)).Select 'ID
                    .range("A" & CStr(iPosTitle + 49 + iIdx)).Select 'ID
                    .activeCell.FormulaR1C1 = oRS1.Item(iIdx).Item("col1")
                    'Modify by Sam on 20190628 for Project FT 配件檢核系統
'                    .range("B" & CStr(iPosTitle + 48 + iIdx)).Select 'Specname
                    .range("B" & CStr(iPosTitle + 49 + iIdx)).Select 'Specname
                    .activeCell.FormulaR1C1 = oRS1.Item(iIdx).Item("specname")
                Else
                    If iIdx <= miFT_RECIPE_MAX_ROW Then
                        'Modify by Sam on 20190628 for Project FT 配件檢核系統
'                        .range("P" & CStr(iPosTitle + 48 + iIdx - miFT_RECIPE_ROW)).Select 'ID
                        .range("P" & CStr(iPosTitle + 49 + iIdx - miFT_RECIPE_ROW)).Select 'ID
                        .activeCell.FormulaR1C1 = oRS1.Item(iIdx).Item("col1")
                        'Modify by Sam on 20190628 for Project FT 配件檢核系統
                        '.range("Q" & CStr(iPosTitle + 48 + iIdx - miFT_RECIPE_ROW)).Select 'Specname
                        .range("Q" & CStr(iPosTitle + 49 + iIdx - miFT_RECIPE_ROW)).Select 'Specname
                        .activeCell.FormulaR1C1 = oRS1.Item(iIdx).Item("specname")
                    End If
                End If
            Next
        Else
            UtShowMsgBox "尚未維護，請確認架機需求或請維護相關資訊。請洽設備工程師。" & vbCrLf & "No data maintenance, make sure the equipment frame machine or maintain information. Please call HW to check."
        End If
        'Add by Tony End on 2014/02/25 for HW SETUP RECIPE COMPARE project
        
    End With

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
' Sub: WsSetupForm()
'--------------------------------------------------------------------------------
' Description:  <Type your Sub description here...>
'--------------------------------------------------------------------------------
' Author:       Jeff Weng, MXIC 2005/03/30
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
'
'================================================================================
Private Sub WsSetupForm(ByRef appXL As Object, _
                        ByRef oLot As FwLot, _
                        ByRef iPosTitle As Integer)
On Error GoTo ExitHandler:
Dim sProcID As String
Dim typErrInfo As tErrInfo
'----
' Init
'----
    sProcID = "WsSetupForm"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog)

'----
' Condition Checking
'----
    ' <Put your condition checking codes here>...

'----
' Action
'----
    ' <Put your Action codes here>...
    Dim sSQL As String, oRS1 As Collection, oRs2 As Collection
    Dim iPosProcess As Integer, iPosComment As Integer, iPosNext As Integer
    Dim msOprId As String, oOpr As FwUser
    Dim msProdgroup As String
    Dim oEqp    As FwEquipment, msSubsystem As String, oSubsystem   As FwEquipment
    Dim msPGNAME             As String
    Dim msSubsystemType2     As String
    Dim msErunTicNo          As String
    Dim msTestMode           As String
    Dim msStep               As String
    Dim sPgName              As String
    Dim msREF_STEP_NAME_01   As String
    Dim msREF_PG_NAME_01     As String
    Dim msREPLACE_PG_NAME_01 As String
    Dim msREF_STEP_NAME_02   As String
    Dim msREF_PG_NAME_02     As String
    Dim msREPLACE_PG_NAME_02 As String
    Dim msREF_STEP_NAME_03   As String
    Dim msREF_PG_NAME_03     As String
    Dim msREPLACE_PG_NAME_03 As String
    Dim msSubsystemType      As String
    Dim msMaxSite            As String
    Dim sPath                As String
    Dim iPosDown             As Integer 'add by sam on 2008/05/19 for ReqNo:M200805020
    
    Dim sFollowProd          As String  'add by sam on 2008/09/16 for PDA_PROJECT
    Dim sErunTicNO           As String  'add by sam on 2008/09/16 for PDA_PROJECT
    
    Dim bTECNLotSpec          As Boolean  'Add by SAM ON 20100407 FOR Project By Lot TECN
    Dim sTecnLotID            As String   'Add by SAM ON 20100407 FOR Project By Lot TECN
    
    'Added by Jack on 2011/09/23 for JC201100228. 不FollowProduct之工程品 改為 帶出新維修單畫面 <Start>
    Dim sEngNotFollowProd_Code      As String
    Dim sEngNotFollowProd_CheckSum  As String 'sEngNotFollowProd_Code,sEngNotFollowProd_CheckSum,sEngNotFollowProd_Speed
    Dim sEngNotFollowProd_Speed     As String
    Dim sStepNo                     As String
    Dim sSubSystemType              As String
    'Added by Jack on 2011/09/23 for JC201100228. 不FollowProduct之工程品 改為 帶出新維修單畫面 <End>
    
    Dim sSwapPgName As String 'Add y Sam on 20150904 for Project 產品變更管理
    Dim sTemp As String
    
    msDocNo = ""
    iPosProcess = iPosTitle + 9
    iPosComment = iPosTitle + 20
    
    'add by sam start on 2008/05/19 for ReqNO:M200805020
    '--------------------------------------------------
    '新增欄位PgmMatch於RANGE E,iPosTitle + 10的位置
    '原Range E,iPosTitle +10 及之後都再加上iposdown的數值
    iPosDown = 1
    'add by sam end on 2008/05/19 for ReqNO:M200805020
    
    msSubsystemType = ""
    msMaxSite = ""
    If InStr(cboSubSystem.Text, ",") > 0 Then
       msSubsystemType = Left(cboSubSystem.Text, InStr(cboSubSystem.Text, ",") - 1)
       msMaxSite = Mid$(cboSubSystem.Text, InStr(cboSubSystem.Text, ",") + 1, Len(cboSubSystem.Text) - InStr(cboSubSystem.Text, ","))
    Else
       msSubsystemType = cboSubSystem.Text
    End If
    
    'Mark by Sam start on 2008/05/19 for ReqNo:M200805020
'    If msSubsystemType = "" Then Exit Sub
    'Mark by Sam end on 2008/05/19 for ReqNo:M200805020

    With appXL
        If iPosTitle = 1 Then
            .Sheets("WS").Select
            .Cells.Select
            .selection.NumberFormatLocal = "@"
            .selection.Copy
            .Sheets(1).Select
            .ActiveSheet.Paste
            .Application.CutCopyMode = False
        Else
            .Sheets("WS").Select
            .Rows("1:21").Select
            .selection.NumberFormatLocal = "@"
            .selection.Copy
            .Sheets(1).Select
            .range("A" & CStr(iPosTitle)).Select
            .selection.Insert -4121
            .Application.CutCopyMode = False
            .range("A" & CStr(iPosTitle)).Select
            .ActiveWindow.SelectedSheets.HPageBreaks.Add appXL.activeCell

        End If
        
        Set oOpr = moFwOPR.ActiveUser
        msOprId = oOpr.UserName
        .range("C" & CStr(iPosTitle + 2)).Select
        .selection.horizontalalignment = -4131
        .activeCell.FormulaR1C1 = msOprId
        
        sSQL = "select shiftcode from tbl_mxic_emp where empno='" & msOprId & "'"
        Set oRS1 = moProRawSql.QueryDatabase(sSQL)
        If oRS1.Count > 0 Then
            .range("H" & CStr(iPosTitle + 2)).Select
            .selection.horizontalalignment = -4131
            .activeCell.FormulaR1C1 = oRS1.Item(1).Item(1)
        End If
        Set oRS1 = Nothing
        
        .range("N" & CStr(iPosTitle + 2)).Select
        .selection.horizontalalignment = -4131
        .activeCell.FormulaR1C1 = Format(Now, "YYYY/MM/DD HH:MM:SS")
        .range("T" & CStr(iPosTitle)).Select
        .activeCell.FormulaR1C1 = "SETUP FORM"
        
        .range("E" & CStr(iPosTitle + 5)).Select
        .activeCell.FormulaR1C1 = oLot.CustomAttributes(gsLOT_CUSTOMATTR_IPN)
        .range("E" & CStr(iPosTitle + 6)).Select
        .activeCell.FormulaR1C1 = oLot.Id
        .range("R" & CStr(iPosTitle + 5)).Select
        .activeCell.FormulaR1C1 = oLot.CustomAttributes(gsLOT_CUSTOMATTR_LOT_OWNER)
        .range("R" & CStr(iPosTitle + 8)).Select
        If oLot.CustomAttributes(gsLOT_CUSTOMATTR_CUR_EQP_ID) <> "" Then
            .activeCell.FormulaR1C1 = oLot.CustomAttributes(gsLOT_CUSTOMATTR_CUR_EQP_ID)
            Set oEqp = FwuRetrieveEqp(moFwMDL, oLot.CustomAttributes(gsLOT_CUSTOMATTR_CUR_EQP_ID), moAppLog)
        Else
            .activeCell.FormulaR1C1 = oLot.CustomAttributes(gsLOT_CUSTOMATTR_SPLITASSIGNEQID)
            Set oEqp = FwuRetrieveEqp(moFwMDL, oLot.CustomAttributes(gsLOT_CUSTOMATTR_SPLITASSIGNEQID), moAppLog)
        End If
        
        sSQL = "select waferid,eruNticno,saprwno from tbl_lot_info where lotid='" & oLot.Id & "'"
        Set oRS1 = moProRawSql.QueryDatabase(sSQL)
        If oRS1.Count > 0 Then
            
            sErunTicNO = oRS1.Item(1).Item("erunticno") 'add by sam on 2008/09/16 for PDA_PROJECT
            
            .range("E" & CStr(iPosTitle + 7)).Select
            .activeCell.FormulaR1C1 = oRS1.Item(1).Item("waferid")
            
            .range("R" & CStr(iPosTitle + 6)).Select
            If oRS1.Item(1).Item("erunticno") = "" Then
                msErunTicNo = oRS1.Item(1).Item("saprwno")
                .activeCell.FormulaR1C1 = msErunTicNo
            ElseIf oRS1.Item(1).Item("saprwno") = "" Then
                msErunTicNo = oRS1.Item(1).Item("erunticno")
                .activeCell.FormulaR1C1 = msErunTicNo
            Else
                msErunTicNo = oRS1.Item(1).Item("erunticno") & ", " & oRS1.Item(1).Item("saprwno")
                .activeCell.FormulaR1C1 = msErunTicNo
            End If
        End If
        
        'ADD BY Sam start on 20080916 for PDA_PROJECT
        'gsCAT_TER_SPEED,gsCAT_TER_CODE,gsCAT_TER_CHECKSUM Added by Jack on 2011/09/23 for JC201100228. 不FollowProduct之工程品 改為 帶出新維修單畫面.
        sSQL = "SELECT " & gsCAT_TER_FOLLOW_PRODUCT & _
                  "," & gsCAT_TER_SPEED & "," & gsCAT_TER_CODE & "," & gsCAT_TER_CHECKSUM & _
                  " from " & gsCAT_TBL_ERUN_REQ & _
                  " where " & gsCAT_TER_LOT_ID & " ='" & oLot.Id & "'" & _
                  " and " & gsCAT_TER_TICKET_NO & "='" & sErunTicNO & "' " & _
                  " and " & gsCAT_TER_STAGE & "='WS' "
                  
           Set oRS1 = moProRawSql.QueryDatabase(sSQL)
           If oRS1.Count > 0 Then
               sFollowProd = oRS1.Item(1).Item(gsCAT_TER_FOLLOW_PRODUCT)
               'Added by Jack on 2011/09/23 for JC201100228. 不FollowProduct之工程品 改為 帶出新維修單畫面 <Start>
               sEngNotFollowProd_Code = oRS1.Item(1).Item(gsCAT_TER_CODE)
               sEngNotFollowProd_CheckSum = oRS1.Item(1).Item(gsCAT_TER_CHECKSUM)
               sEngNotFollowProd_Speed = oRS1.Item(1).Item(gsCAT_TER_SPEED)
               'Added by Jack on 2011/09/23 for JC201100228. 不FollowProduct之工程品 改為 帶出新維修單畫面 <End>
           End If
        'ADD BY Sam start on 20080916 for PDA_PROJECT
        
        msStep = oLot.CurrentStep.Steps.Item(1).Description

        'msTestMode = Replace(UCase(Mid(CStr(msStep), 1)), "ORT", "")
        .range("E" & CStr(iPosTitle + 8)).Select
        .activeCell.FormulaR1C1 = msStep
 
        Set oRS1 = Nothing
               
        sSQL = "select speed,code,checksum,prodgroup from tbl_ipn_master where ipn='" & oLot.CustomAttributes(gsLOT_CUSTOMATTR_IPN) & "'"
        Set oRS1 = moProRawSql.QueryDatabase(sSQL)
        If oRS1.Count > 0 Then
            .range("R" & CStr(iPosTitle + 10)).Select
            .activeCell.FormulaR1C1 = Mid(oRS1.Item(1).Item("Code"), 3, 6)
            .range("R" & CStr(iPosTitle + 11)).Select
            .activeCell.FormulaR1C1 = oRS1.Item(1).Item("checksum")
            msProdgroup = oRS1.Item(1).Item("prodgroup")
            .range("R" & CStr(iPosTitle + 12)).Select
            .activeCell.FormulaR1C1 = oRS1.Item(1).Item("speed")
        End If
        Set oRS1 = Nothing
        'Add by Janus, 20051017, M200509025
        'Check MCD
        Set oRS1 = moProRawSql.QueryDatabase("select count(*)" & _
                   " from " & gsCAT_TBL_MULTICODE_IPN & _
                   " where " & gsCAT_TMI_IPN & " = '" & oLot.CustomAttributes(gsLOT_CUSTOMATTR_IPN) & "' " & _
                   " and   " & gsCAT_TMI_DELETEFLAG & " = 'N'")
        If oRS1.Count > 0 Then
           If Val(oRS1.Item(1).Item(1)) > 0 Then
              .range("R" & CStr(iPosTitle + 10)).Select
              .activeCell.FormulaR1C1 = "Check Code Server"
              .range("R" & CStr(iPosTitle + 11)).Select
              .activeCell.FormulaR1C1 = "Check Code Server"
           End If
        End If
        Set oRS1 = Nothing
        'Check MCP
        sSQL = "select ipn from tbl_ipn_bom where parent = '" & oLot.CustomAttributes(gsLOT_CUSTOMATTR_IPN) & "' " & _
               "and bom_level = 'FG' and deleteflag = 'N'"
        Set oRS1 = moProRawSql.QueryDatabase(sSQL)
        If oRS1.Count > 0 Then
           sSQL = "select PACKAGECOMPONENT from tbl_prod_body where prodbody = '" & Left(oRS1.Item(1).Item("ipn"), 4) & "' " & _
                  "and deleteflag = 'N' "
           Set oRs2 = moProRawSql.QueryDatabase(sSQL)
           If oRs2.Count > 0 Then
              If Left(oRs2.Item(1).Item("PACKAGECOMPONENT"), 3) = "MCP" Then
                 .range("R" & CStr(iPosTitle + 10)).Select
                 .activeCell.FormulaR1C1 = "Check Code Server"
                 .range("R" & CStr(iPosTitle + 11)).Select
                 .activeCell.FormulaR1C1 = "Check Code Server"
              End If
           End If
           Set oRs2 = Nothing
        End If
        Set oRS1 = Nothing
        'End add by Janus, 20051017,M200509025
        
        'add by Nelson start on 2007/3/20 for ReqNo:M200702009
        sSQL = "SELECT PATH FROM TBL_ROUTE_STEP_LIST WHERE ROUTEID='" & oLot.PlanId & "' AND FWSTEPSEQ='" & oLot.CurrentStep.Steps.Item(1).Handle & "'"
        Set oRS1 = moProRawSql.QueryDatabase(sSQL)
        If oRS1.Count > 0 Then
            sPath = oRS1.Item(1).Item("PATH")
        End If
        Set oRS1 = Nothing
        'add by Nelson end on 2007/3/20 for ReqNo:M200702009
        
        If Not oEqp Is Nothing Then
            .range("R" & CStr(iPosTitle + 9)).Select
'            .activecell.formulaR1C1 = msSubsystemType
             If oEqp.CustomAttributes(gsEQP_CUSTOMATTR_SUBSYS2) = "" Then
                .activeCell.FormulaR1C1 = oEqp.CustomAttributes(gsEQP_CUSTOMATTR_SUBSYS1)
            Else
                .activeCell.FormulaR1C1 = oEqp.CustomAttributes(gsEQP_CUSTOMATTR_SUBSYS1) & " " & oEqp.CustomAttributes(gsEQP_CUSTOMATTR_SUBSYS2)
            End If
'            If oEqp.CustomAttributes(gsEQP_CUSTOMATTR_SUBSYS2) = "" Then
'                msSubsystem = oEqp.CustomAttributes(gsEQP_CUSTOMATTR_SUBSYS1)
'            Else
'                msSubsystem = oEqp.CustomAttributes(gsEQP_CUSTOMATTR_SUBSYS2)
'            End If
'
'            Set oSubsystem = FwuRetrieveEqp(moFwMDL, msSubsystem, moAppLog)
'            msSubsystemType2 = oSubsystem.CustomAttributes(gsEQP_CUSTOMATTR_EQ_TYPE2)
            
            'aDD BY Sam on start on 20100514 for Project Lot TECN
            bTECNLotSpec = CheckTECNLotInfo(moAppLog, moFwWIP, moFwWF, moCwMbx, oLot.Id)
            If bTECNLotSpec = True Then
                   sSQL = " SELECT A." & gsCAT_TLI_TECN_LOT_ID & " " & _
                          " FROM " & gsCAT_TBL_LOT_INFO & " A " & _
                          " WHERE A." & gsCAT_TLI_LOT_ID & " ='" & oLot.Id & "' "
                
                    Set oRS1 = moProRawSql.QueryDatabase(sSQL)
                    If oRS1.Count > 0 Then
                      sTecnLotID = oRS1.Item(1).Item(1)
                    End If
                    
                    sSQL = "select Temperature, Comments,wsdevicefile,probecardtype,loadboardtype," & _
                            "cabletype,pgname,Subsystem,eqtype2,'TECN' AS doctype,docno," & _
                            "pgid,pgname,refpgname01,refpgname02,refpgname03,refstepname01," & _
                            "refstepname02,refstepname03,replacepgname01,replacepgname02," & _
                            "replacepgname03 ,PGMMATCH " & _
                            "from tbl_lot_step_eq_spec where " & _
                            " '" & sTecnLotID & "' like tecnlotid and  " & _
                            "stepno='" & oLot.CurrentStep.Steps.Item(1).Id & "' and " & _
                            "eqtype2='" & oEqp.CustomAttributes(gsEQP_CUSTOMATTR_EQ_TYPE2) & "' and " & _
                            "subsystem ='" & msSubsystemType & "' and " & _
                            " deleteflag ='N' and nvl(maxsite,' ') = nvl('" & msMaxSite & "',' ')"
                
            'Add by SAM END ON 20100514 FOR Project By Lot TECN
            Else

            'modify by sam start on 2008/05/19 for ReqNo:M200805020
    '            sSQL = "select Temperature, Comments,wsdevicefile,probecardtype,loadboardtype," & _
    '                    "cabletype,pgname,Subsystem,eqtype2,doctype,docno," & _
    '                    "pgid,pgname,refpgname01,refpgname02,refpgname03,refstepname01," & _
    '                    "refstepname02,refstepname03,replacepgname01,replacepgname02," & _
    '                    "replacepgname03 " & _
    '                    "from tbl_prod_step_eq_spec where " & _
    '                    "prodgroup='" & msProdgroup & "' and " & _
    '                    "stepno='" & oLot.CurrentStep.Steps.Item(1).Id & "' and " & _
    '                    "eqtype2='" & oEqp.CustomAttributes(gsEQP_CUSTOMATTR_EQ_TYPE2) & "' and " & _
    '                    "subsystem ='" & msSubsystemType & "' and " & _
    '                    "docstatus ='Active' and maxsite = '" & msMaxSite & "' and path='" & sPath & "'"
                sSQL = "select Temperature, Comments,wsdevicefile,probecardtype,loadboardtype," & _
                        "cabletype,pgname,Subsystem,eqtype2,doctype,docno," & _
                        "pgid,pgname,refpgname01,refpgname02,refpgname03,refstepname01," & _
                        "refstepname02,refstepname03,replacepgname01,replacepgname02," & _
                        "replacepgname03 ,PGMMATCH " & _
                        "from tbl_prod_step_eq_spec where " & _
                        "prodgroup='" & msProdgroup & "' and " & _
                        "stepno='" & oLot.CurrentStep.Steps.Item(1).Id & "' and " & _
                        "eqtype2='" & oEqp.CustomAttributes(gsEQP_CUSTOMATTR_EQ_TYPE2) & "' and " & _
                        "subsystem ='" & msSubsystemType & "' and " & _
                        "docstatus ='Active' and nvl(maxsite,' ') = nvl('" & msMaxSite & "',' ') and path='" & sPath & "'"
            'modify by sam END on 2008/05/19 for ReqNo:M200805020
            End If
                    
            Set oRS1 = moProRawSql.QueryDatabase(sSQL)
            If oRS1.Count > 0 Then
                'add by sam start on 2008/05/19 for ReqNo:M200805020
                .range("E" & CStr(iPosTitle + 10)).Select
                .activeCell.FormulaR1C1 = oRS1.Item(1).Item("PGMMATCH")
                'add by sam END on 2008/05/19 for ReqNo:M200805020
                
                'Add by Sam start on 20171121 for Project TECN 自動化,改成使用與停復測相同的來源
                sTemp = oRS1.Item(1).Item("Temperature")
                If Trim(txtTemp.Text) <> "" Then
                    sTemp = txtTemp.Text
                End If
                'Add by Sam END on 20171121 for Project TECN 自動化
                
                .range("E" & CStr(iPosTitle + 10 + iPosDown)).Select
                '.activeCell.FormulaR1C1 = oRS1.Item(1).Item("Temperature")
                .activeCell.FormulaR1C1 = sTemp
                
                .range("E" & CStr(iPosTitle + 13 + iPosDown)).Select
                .activeCell.FormulaR1C1 = oRS1.Item(1).Item("Comments")
                .range("E" & CStr(iPosTitle + 11 + iPosDown)).Select
                .activeCell.FormulaR1C1 = oRS1.Item(1).Item("wsdevicefile")
                .range("E" & CStr(iPosTitle + 16 + iPosDown)).Select
                .activeCell.FormulaR1C1 = oRS1.Item(1).Item("probecardtype")
                .range("E" & CStr(iPosTitle + 17 + iPosDown)).Select
                .activeCell.FormulaR1C1 = oRS1.Item(1).Item("loadboardtype")
                .range("E" & CStr(iPosTitle + 18 + iPosDown)).Select
                .activeCell.FormulaR1C1 = oRS1.Item(1).Item("cabletype")
                'Modified by Jack on 2018/02/02 for Project TECN 自動化 <Start>.
                '改為DocType不為空值, 若 TECNNO有值, 要加到 DocNO前頭.
                'If Trim(oRS1.Item(1).Item("doctype")) = "TECN" Then
                If Trim(oRS1.Item(1).Item("doctype")) <> "" Then
                'Modified by Jack on 2018/02/02 for Project TECN 自動化 <End>.
                   .range("E" & CStr(iPosTitle + 12 + iPosDown)).Select
                   '.activeCell.FormulaR1C1 = oRS1.Item(1).Item("docno") 'Maked by Jack on 2018/02/02 for Project TECN 自動化.
                   'Modidied by Jack on 2018/02/02 for Project TECN 自動化 <Start>.
                   If msPGMTECNNo <> "" Then
                        .activeCell.FormulaR1C1 = msPGMTECNNo & "+" & oRS1.Item(1).Item("docno")
                   Else
                        .activeCell.FormulaR1C1 = oRS1.Item(1).Item("docno")
                   End If
                   'Added by Jack on 2018/02/02 for Project TECN 自動化 <End>.
                End If
                sPgName = oRS1.Item(1).Item("PGNAME")
                msREF_STEP_NAME_01 = oRS1.Item(1).Item("refstepname01")
                msREF_PG_NAME_01 = oRS1.Item(1).Item("refpgname01")
                msREPLACE_PG_NAME_01 = oRS1.Item(1).Item("replacepgname01")
                msREF_STEP_NAME_02 = oRS1.Item(1).Item("refstepname02")
                msREF_PG_NAME_02 = oRS1.Item(1).Item("refpgname02")
                msREPLACE_PG_NAME_02 = oRS1.Item(1).Item("replacepgname02")
                msREF_STEP_NAME_03 = oRS1.Item(1).Item("refstepname03")
                msREF_PG_NAME_03 = oRS1.Item(1).Item("refpgname03")
                msREPLACE_PG_NAME_03 = oRS1.Item(1).Item("replacepgname03")
                
                
                'Add by Sam Start on 20150904 for Project 產品變更管理
                sSwapPgName = GetSwapPGName(moProRawSql, moAppLog, oLot.Id, msProdgroup, oLot.CurrentStep.Steps.Item(1).Description _
                                        , oEqp.CustomAttributes(gsEQP_CUSTOMATTR_EQ_TYPE2), sPgName, False)
                msPGNAME = sPgName
                If sSwapPgName <> "" Then
                    If sSwapPgName <> "X" Then
                        msPGNAME = sSwapPgName
                    End If
                End If
                'Mark by Sam End on 20150904 for Project 產品變更管理,SWAP PGNAME改用新制
                
                'Mark by Sam start on 20150904 for Project 產品變更管理
'                msPGNAME = GetPGM(oLot, sPgName, msREF_STEP_NAME_01, msREF_PG_NAME_01, msREPLACE_PG_NAME_01, _
'                           msREF_STEP_NAME_02, msREF_PG_NAME_02, msREPLACE_PG_NAME_02, _
'                           msREF_STEP_NAME_03, msREF_PG_NAME_03, msREPLACE_PG_NAME_03, _
'                           oEqp.CustomAttributes(gsEQP_CUSTOMATTR_EQ_TYPE2), msSubsystemType)
                'Mark by Sam End on 20150904 for Project 產品變更管理

                'Add by Sam start on 20171121 for Project TECN 自動化,改成使用與停復測相同的來源
                If Trim(txtPGName.Text) <> "" Then
                    msPGNAME = txtPGName
                End If
                'Add by Sam end on 20171121 for Project TECN 自動化,改成使用與停復測相同的來源

                .range("E" & CStr(iPosTitle + 9)).Select
                .activeCell.FormulaR1C1 = msPGNAME
            End If
            Set oRS1 = Nothing

            Dim msPath As String
           
            sSQL = "select route " & _
                   " from fwadmin.tbl_lot_attribute where lotid='" & txtLotID.Text & "'"
            Set oRS1 = moProRawSql.QueryDatabase(sSQL)
            If oRS1.Count > 0 Then
               msPath = oRS1.Item(1).Item("route")
            End If
            Set oRS1 = Nothing
            'Add by SAM START ON 20100512 FOR Project By Lot TECN
            bTECNLotSpec = CheckTECNLotInfo(moAppLog, moFwWIP, moFwWF, moCwMbx, oLot.Id)
            If bTECNLotSpec = True Then
                
                sSQL = "select Comments,'TECN' AS  doctype,docno " & _
                       "from tbl_Lot_step_spec where " & _
                       " '" & sTecnLotID & "' like tecnlotid and  " & _
                       " stepno='" & oLot.CurrentStep.Steps.Item(1).Id & "' and " & _
                       " deleteflag='N' and path = '" & msPath & "'"

            Else
            'Add by SAM End ON 20100512 FOR Project By Lot TECN

                sSQL = "select Comments,doctype,docno " & _
                       "from tbl_prod_step_spec where " & _
                       "prodgroup='" & msProdgroup & "' and " & _
                       "stepno='" & oLot.CurrentStep.Steps.Item(1).Id & "' and " & _
                       "docstatus ='Active' and path = '" & msPath & "'"
            End If
            Set oRS1 = moProRawSql.QueryDatabase(sSQL)
            If oRS1.Count > 0 Then
               .range("E" & CStr(iPosTitle + 14 + iPosDown)).Select
                .activeCell.FormulaR1C1 = oRS1.Item(1).Item("Comments")
            End If
            Set oRS1 = Nothing
        End If


'         msPGNAME = GetPGM(oEqp.Id, oLot, msSubsystemType2, msDocNo, gsLOTUNIT_WAFER)
'        .Range("E" & CStr(iPosTitle + 9)).Select
'        .activecell.formulaR1C1 = msPGNAME

        
        ' if msErunTicNo is not Enpity表重工品,則
        '設為空值 : PGM, Temperature, Speed, DeviceFile, DocumentNo, Special Request, Code, CheckSum, ProbeCard, LoadBoard, Cable
        'user直接查詢委測/重工單, 避免混淆
        'modify by sam on 20080916 for PDA_PROJECT  增加條件FOLLOW PRODUCT='N'的才清空
'        If Len(msErunTicNo) <> 0 Then
        If Len(msErunTicNo) <> 0 And sFollowProd = "N" Then
            .range("E" & CStr(iPosTitle + 9)).Select     'PGM
            .activeCell.FormulaR1C1 = ""
            .range("E" & CStr(iPosTitle + 10)).Select    'Pgm Match added by Jack 2008/08/05 [M200807052]
            .activeCell.FormulaR1C1 = ""
            .range("E" & CStr(iPosTitle + 10 + iPosDown)).Select   'Temperature
            .activeCell.FormulaR1C1 = ""
            .range("E" & CStr(iPosTitle + 11 + iPosDown)).Select   'DeviceFile
            .activeCell.FormulaR1C1 = ""
            .range("E" & CStr(iPosTitle + 12 + iPosDown)).Select  'Tecn No
            .activeCell.FormulaR1C1 = ""
            .range("E" & CStr(iPosTitle + 13 + iPosDown)).Select   'Eq Comment
            .activeCell.FormulaR1C1 = ""
            .range("E" & CStr(iPosTitle + 14 + iPosDown)).Select   'Step Comment
            .activeCell.FormulaR1C1 = ""
            'Modified by Jack on 2011/09/23 for JC201100228. 不FollowProduct之工程品 改為 帶出新維修單畫面. <Start>
            '原本為清空, 改為串Tbl_Erun_Req的資料.
            .range("R" & CStr(iPosTitle + 10)).Select    'Code
            '.activeCell.FormulaR1C1 = ""
            .activeCell.FormulaR1C1 = sEngNotFollowProd_Code
            .range("R" & CStr(iPosTitle + 11)).Select    'CheckSum
            '.activeCell.FormulaR1C1 = ""
            .activeCell.FormulaR1C1 = sEngNotFollowProd_CheckSum
            .range("R" & CStr(iPosTitle + 12)).Select    'Speed
            '.activeCell.FormulaR1C1 = ""
            .activeCell.FormulaR1C1 = sEngNotFollowProd_Speed
            'Modified by Jack on 2011/09/23 for JC201100228. 不FollowProduct之工程品 改為 帶出新維修單畫面. <End>
            .range("E" & CStr(iPosTitle + 16 + iPosDown)).Select   'ProbeCard
            .activeCell.FormulaR1C1 = ""
            .range("E" & CStr(iPosTitle + 17 + iPosDown)).Select   'LoadBoard
            .activeCell.FormulaR1C1 = ""
            .range("E" & CStr(iPosTitle + 18 + iPosDown)).Select   'Cable
            .activeCell.FormulaR1C1 = ""
            
            'Modified by Jack on 2011/09/23 for JC201100228. 不FollowProduct之工程品 改為 帶出新維修單畫面. <Start>
            '原本為清空, 改為串Tbl_Erun_Recipe的資料.
            If InStr(cboSubSystem.Text, ",") > 0 Then
               sSubSystemType = Left(cboSubSystem.Text, InStr(cboSubSystem.Text, ",") - 1)
            Else
               sSubSystemType = cboSubSystem.Text
            End If
            
            sStepNo = oLot.CurrentStep.Steps.Item(1).Id
            sSQL = "SELECT " & gsCAT_TERE_PGNAME & _
                      "," & gsCAT_TERE_TEMPERATURE & "," & gsCAT_TERE_WSDEVICEFILE & _
                      "," & gsCAT_TERE_PROBECARDTYPE & "," & gsCAT_TERE_LOADBOARDTYPE & _
                      " from " & gsCAT_TBL_ERUN_RECIPE & _
                      " where " & gsCAT_TERE_STEPNO & " ='" & sStepNo & "'" & _
                      " and " & gsCAT_TERE_DOCNO & "='" & sErunTicNO & "' " & _
                      " and eqtype2 = '" & oEqp.CustomAttributes(gsEQP_CUSTOMATTR_EQ_TYPE2) & "' "
            If Trim(sSubSystemType) = "" Then
                sSQL = sSQL & " and subsystem is null "
            Else
                sSQL = sSQL & " and subsystem='" & Trim(sSubSystemType) & "' "
            End If
            sSQL = sSQL & " and deleteflag = 'N' "
                      
            Set oRS1 = moProRawSql.QueryDatabase(sSQL)
            If oRS1.Count > 0 Then
                 .range("E" & CStr(iPosTitle + 9)).Select     'PGM
                 .activeCell.FormulaR1C1 = oRS1.Item(1).Item(gsCAT_TERE_PGNAME)
                 .range("E" & CStr(iPosTitle + 10 + iPosDown)).Select   'Temperature
                 .activeCell.FormulaR1C1 = oRS1.Item(1).Item(gsCAT_TERE_TEMPERATURE)
                 .range("E" & CStr(iPosTitle + 11 + iPosDown)).Select   'DeviceFile
                 .activeCell.FormulaR1C1 = oRS1.Item(1).Item(gsCAT_TERE_WSDEVICEFILE)
                 .range("E" & CStr(iPosTitle + 16 + iPosDown)).Select   'ProbeCard
                 .activeCell.FormulaR1C1 = oRS1.Item(1).Item(gsCAT_TERE_PROBECARDTYPE)
                 .range("E" & CStr(iPosTitle + 17 + iPosDown)).Select   'LoadBoard
                 .activeCell.FormulaR1C1 = oRS1.Item(1).Item(gsCAT_TERE_LOADBOARDTYPE)
            End If
            'Modified by Jack on 2011/09/23 for JC201100228. 不FollowProduct之工程品 改為 帶出新維修單畫面. <End>
            
        End If
                
        
    End With


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
' Function: GetPGM()
'--------------------------------------------------------------------------------
' Description:  <Type your function description here...>
'--------------------------------------------------------------------------------
' Author:       Jeff Weng, MXIC 2005/03/30
'--------------------------------------------------------------------------------
' Parameters:
'................................................................................
' ARGUMENTS
'   oLogCtrl            (I) [Optional] Valid reference to a clsLogTraceMsg object
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
'
'================================================================================
Private Function GetPGM(oLot As FwLot, sPgName As String, sREF_STEP_NAME_01 As String, sREF_PG_NAME_01 As String, _
    sREPLACE_PG_NAME_01 As String, sREF_STEP_NAME_02 As String, sREF_PG_NAME_02 As String, _
    sREPLACE_PG_NAME_02 As String, sREF_STEP_NAME_03 As String, sREF_PG_NAME_03 As String, _
    sREPLACE_PG_NAME_03 As String, sEqType2 As String, sSubSysType As String) As String
On Error GoTo ExitHandler:
Dim sProcID As String
Dim typErrInfo As tErrInfo
'----
' Init
'----
    sProcID = "GetPGM"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog) '"Entering Function...", moAppLog, glLOG_PROC, msMODULE_ID, sProcID)

'----
' Condition Checking
'----
    ' <Put your condition checking codes here>...

'----
' Action
'----
    ' <Put your Action codes here>...
    'Dim oEqp                As FwEquipment
    'Dim msEqType2           As String
    Dim oRawSqlRecordSet    As clsRawSqlRecordset
    Dim sLotID              As String
    'Dim sPath               As String
    'Dim sProdGroup          As String
    Dim sStepNo             As String
    'Dim sStepName           As String
    'Dim sLotOwner           As String
    'Dim sWaferQty           As String
    'Dim sChipQty            As String
    'Dim sIPN                As String
    'Dim sLotStatus1         As String
    'Dim sHotLotFlag         As String
    'Dim sERunTicNo          As String
    'Dim sFollowOI           As String
    'Dim sSapRwNo            As String
    'Dim sPlanId             As String
    'Dim sPlanVer            As String
    'Dim sCurHandle          As String
    'Dim sTimeStamp          As String
    'Dim sLastTestStep       As String
    'Dim sTECN               As String
    'Dim sLotFutAct          As String
    'Dim sIPNFutAct          As String
    Dim sPgId               As String
    Dim msPGNAME             As String
    'Dim sTemp               As String
    'Dim sProcTime           As String
    'Dim sProcUnit           As String
    'Dim sSiteDie            As String
    Dim sAttPG1             As String
    Dim sAttPG2             As String
    Dim sAttPG3             As String
    Dim sSQL                As String
    Dim colRawSQL           As Collection
    'Dim sTmpDocNo           As String
    Dim sDocNo1             As String
    Dim sPath               As String
    Dim vStepName           As Variant 'add by Ernest on 2016/04/21 for ReqNo:JC201500378
    'sTmpDocNo = ""
    'Set oEqp = FwuRetrieveEqp(moFwMDL, sEqID, moAppLog)
    
    'msEqType2 = oEqp.CustomAttributes(modConstFwAttr.gsEQP_CUSTOMATTR_EQ_TYPE2)
    
'    If oLot.CustomAttributes(gsLOT_CUSTOMATTR_CUR_EQP_ID) <> "" Then
'        Set oRawSqlRecordSet = RetrieveSortLotInfo(moAppLog, moCwMbx, moFwWF, oEqp.Type, msEqType2, _
'                                                gsLOTSTATUS_RUN & "','" & gsLOTSTATUS_ERUN & "','" & gsLOTSTATUS_WAITR & "','" & gsLOTSTATUS_BREAK, _
'                                                msRuleName, sUnit, sEqID, , , oLot, sSubsystemType2)
'    Else
'        Set oRawSqlRecordSet = RetrieveSortLotInfo(moAppLog, moCwMbx, moFwWF, oEqp.Type, msEqType2, _
'                                                gsLOTSTATUS_RUN & "','" & gsLOTSTATUS_ERUN & "','" & gsLOTSTATUS_WAITR & "','" & gsLOTSTATUS_BREAK, _
'                                                msRuleName, sUnit, , sEqID, , oLot, sSubsystemType2)
'    End If
        
'        oRawSqlRecordSet.MoveFirst
'        Do While Not oRawSqlRecordSet.EOF
            sLotID = oLot.Id
            'modify by Ernest on 2016/04/21 for ReqNo:JC201500378---------start
            'sStepNo = oLot.CurrentStep.Steps.Item(1).Id
            vStepName = Split(Me.cboStep.Text, " ")
            sStepNo = Trim(vStepName(0))
            'modify by Ernest on 2016/04/21 for ReqNo:JC201500378---------end
            'sLotId = oRawSqlRecordSet.Value("LotId")
            'sPath = oRawSqlRecordSet.Value("Route")
            'sProdGroup = oRawSqlRecordSet.Value("ProdGroup")
            'sStepNo = oRawSqlRecordSet.Value("StepNo")
            'sStepName = oRawSqlRecordSet.Value("StepName")
            'sWaferQty = oRawSqlRecordSet.Value("WaferQty")
            'sChipQty = oRawSqlRecordSet.Value("ChipQty")
            'sIPN = oRawSqlRecordSet.Value("IPN")
            'sLotStatus1 = oRawSqlRecordSet.Value("LotStatus1")
            'sERunTicNo = oRawSqlRecordSet.Value("ERunTicNo")
            'sFollowOI = oRawSqlRecordSet.Value("FollowOI")
            'sSapRwNo = oRawSqlRecordSet.Value("SapRwNo")
            'sPlanId = oRawSqlRecordSet.Value("PlanName")
            'sPlanVer = oRawSqlRecordSet.Value("PlanRevision")
            'sCurHandle = oRawSqlRecordSet.Value("Handle")
            'sTimeStamp = oRawSqlRecordSet.Value("FwTimeStamp")
            'sLastTestStep = oRawSqlRecordSet.Value("LastTestStep")
            'sTECN = oRawSqlRecordSet.Value("TECN")
            'sLotFutAct = oRawSqlRecordSet.Value("LotFutAct")
            'sIPNFutAct = oRawSqlRecordSet.Value("IPNFutAct")
            'sDocNo = oRawSqlRecordSet.Value("DocNo")
                '----
                ' 搜尋Future Action , 若有的話, 用future action的PgId and PgName,
                ' 若沒有的話,用Tbl_Prod_Step_Eq_Spec ECN 之PgId and Name
                '----
            sPath = oLot.CustomAttributes(gsLOT_CUSTOMATTR_ROUTE)
            Call GetPgIdAndNameFromFutAct(moAppLog, moFwWIP, moFwWF, moCwMbx, _
                sLotID, sStepNo, sPath, sPgId, msPGNAME, sDocNo1, sEqType2, sSubSysType)
'                If Len(Trim$(sDocNo1)) <> 0 Then
'                    If sDocNo = "" Then
'                        sDocNo = sDocNo1
'                    Else
'                        sDocNo = sDocNo & ", " & sDocNo1
'                    End If
'                End If
                
            If Len(Trim$(msPGNAME)) <> 0 Then
                    'Call GetPgIdAndNameFromFutAct(moAppLog, moFwWIP, moFwWF, moCwMbx, sLotId, sStepNo, sPgId, sPgName, sDocNo)
                    'If Len(Trim$(sPgName)) = 0 Then
                    '    sPgName = oRawSqlRecordSet.Value("PgName")
                    '    sDocNo = oRawSqlRecordSet.Value("DocNo")
                    'End If
            Else
               msPGNAME = sPgName
                    'sDocNo = oRawSqlRecordSet.Value("DocNo")
                    '--------------------------------------------------------------------
               If Val(sREF_STEP_NAME_01) > 0 Then
                  sAttPG1 = ""
                  sAttPG2 = ""
                  sAttPG3 = ""
                  sSQL = "select B.VALDATA PG1, C.VALDATA PG2, D.VALDATA PG3" & _
                        " from FWLOT A, FWLOT_PN2M B, FWLOT_PN2M C, FWLOT_PN2M D " & _
                         " where A.APPID = '" & sLotID & "' " & _
                         " and A.SYSID = B.FROMID " & _
                         " and B.KEYDATA = '" & gsLOT_CUSTOMATTR_TESTED1STPGNAME & "'" & _
                         " and A.SYSID = C.FROMID " & _
                         " and C.KEYDATA = '" & gsLOT_CUSTOMATTR_TESTED2NDPGNAME & "'" & _
                         " and A.SYSID = D.FROMID " & _
                         " and D.KEYDATA = '" & gsLOT_CUSTOMATTR_TESTED3RDPGNAME & "'"
                  Set colRawSQL = modRawSQL.SelectRawSqlTxn1(moAppLog, moFwWIP, moFwWF, moCwMbx, "CAT", sSQL)
                  If colRawSQL.Count > 0 Then
                     sAttPG1 = colRawSQL.Item(1).Item(1)
                     sAttPG2 = colRawSQL.Item(1).Item(2)
                     sAttPG3 = colRawSQL.Item(1).Item(3)
                  End If
                  If (Val(sREF_STEP_NAME_01) = 1 And sAttPG1 <> "" And _
                     sAttPG1 = sREF_PG_NAME_01) Or (Val(sREF_STEP_NAME_01) = 2 And _
                     sAttPG2 <> "" And sAttPG2 = sREF_PG_NAME_01) Or _
                     (Val(sREF_STEP_NAME_01) = 3 And sAttPG3 <> "" And _
                     sAttPG3 = sREF_PG_NAME_01) Then
                     msPGNAME = sREPLACE_PG_NAME_01
                  ElseIf (Val(sREF_STEP_NAME_02) = 1 And sAttPG1 <> "" And _
                          sAttPG1 = sREF_PG_NAME_02) Or _
                         (Val(sREF_STEP_NAME_02) = 2 And sAttPG2 <> "" And _
                          sAttPG2 = sREF_PG_NAME_02) Or _
                         (Val(sREF_STEP_NAME_02) = 3 And sAttPG3 <> "" And _
                          sAttPG3 = sREF_PG_NAME_02) Then
                          msPGNAME = sREPLACE_PG_NAME_02
                  ElseIf (Val(sREF_STEP_NAME_03) = 1 And sAttPG1 <> "" And _
                          sAttPG1 = sREF_PG_NAME_03) Or _
                         (Val(sREF_STEP_NAME_03) = 2 And sAttPG2 <> "" And _
                          sAttPG2 = sREF_PG_NAME_03) Or _
                         (Val(sREF_STEP_NAME_03) = 3 And sAttPG3 <> "" And _
                          sAttPG3 = sREF_PG_NAME_03) Then
                          msPGNAME = sREPLACE_PG_NAME_03 'Modified by Tony on 20071030 for Req.M200710069
                  End If
               End If
               '--------------------------------------------------------------------
            End If
'                If sDocNo = "" Then
'                    sDocNo = oRawSqlRecordSet.Value("DocNo")
'                Else
'                    sDocNo = sDocNo & ", " & oRawSqlRecordSet.Value("DocNo")
'                End If
                
                
            'oRawSqlRecordSet.MoveNext
        'Loop
        
    GetPGM = msPGNAME
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

Private Sub lblCurrentStep_Click()

End Sub

Private Sub txtLotID_Change()
    msLotId = UCase(txtLotID)
    Call ExecQuery
End Sub

Private Sub txtLotId_KeyPress(KeyAscii As Integer)
    
'    If KeyAscii = 13 And txtLotID.Text <> "" Then
'        Call ExecQuery
'    End If

End Sub

'================================================================================
' Sub: CheckForm()
'--------------------------------------------------------------------------------
' Description:  <Type your function description here...>
'--------------------------------------------------------------------------------
' Author:       Sam Chen ,Cit 2008/08/26
'--------------------------------------------------------------------------------
' Parameters:
'................................................................................
' ARGUMENTS
'   oLogCtrl            (I) [Optional] Valid reference to a clsLogTraceMsg object
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
'
'================================================================================
Private Sub CheckForm(ByVal sLotID As String)
On Error GoTo ExitHandler:
Dim sProcID         As String
Dim typErrInfo      As tErrInfo
Dim sSQL            As String
Dim oRS             As Collection
Dim sStage          As String
Dim sErunTicNO      As String
Dim sFollowProd     As String

'----
' Init
'----
    sProcID = "CheckForm"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog) '"Entering Function...", moAppLog, glLOG_PROC, msMODULE_ID, sProcID)
   
'----
' Condition Checking
'----
    ' <Put your condition checking codes here>...

'----
' Action
'----
    ' <Put your Action codes here>...
    sSQL = "SELECT " & gsCAT_TLATT_STAGE & _
           " from " & gsCAT_TBL_LOT_ATTRIBUTE & _
           " where " & gsCAT_TLATT_LOTID & " ='" & sLotID & "'"
    Set oRS = moProRawSql.QueryDatabase(sSQL)
    If oRS.Count > 0 Then
        sStage = oRS.Item(1).Item(gsCAT_TLATT_STAGE)
    End If
           
    sSQL = "SELECT " & gsCAT_TLI_ERUNTICNO & _
           " from " & gsCAT_TBL_LOT_INFO & _
           " where " & gsCAT_TLI_LOT_ID & " ='" & sLotID & "'"
    Set oRS = moProRawSql.QueryDatabase(sSQL)
    If oRS.Count > 0 Then
        sErunTicNO = oRS.Item(1).Item(gsCAT_TLI_ERUNTICNO)
    End If
           
    sSQL = "SELECT " & gsCAT_TER_FOLLOW_PRODUCT & _
           " from " & gsCAT_TBL_ERUN_REQ & _
           " where " & gsCAT_TER_LOT_ID & " ='" & sLotID & "'" & _
           " and " & gsCAT_TER_TICKET_NO & "='" & sErunTicNO & "' " & _
           " and " & gsCAT_TER_STAGE & "='WS' "
           
    Set oRS = moProRawSql.QueryDatabase(sSQL)
    If oRS.Count > 0 Then
        sFollowProd = oRS.Item(1).Item(gsCAT_TER_FOLLOW_PRODUCT)
    End If
                      
    If UCase(sStage) = "WS" Then
        'Modified by Jack on 2011/09/23 for JC201100228. 不FollowProduct之工程品 改為 帶出新維修單畫面 <Start>
        'If sERunTicNo <> "" And sFollowProd = "N" Then
            '維持舊版架機單
        
        'Else
            '新版架機單---------------------------------------------------
            'Modify by Sam start on 20130819 for ReqNo:JC201300238
'            Me.Height = 5115
'            cmdOK.Top = 4120
'            cmdCancel.Top = 4120
'            fraLotInformation.Height = 3135
            Me.Height = 6120
            cmdOK.Top = 5040
            cmdCancel.Top = 5040
            fraLotInformation.Height = 4095
            'Modify by Sam End on 20130819 for ReqNo:JC201300238
            mbNewForm = True
            Me.cboStep.Locked = True   'add by Ernest on 2016/04/21 for ReqNo"JC201500378
            Me.cboTesterId.Locked = True 'add by Ernest on 2016/04/21 for ReqNo"JC201500378
            '-------------------------------------------------------------
            GoTo ExitHandler
        'End If
    End If
    
'舊版---------------------------------------------------------
        Me.Height = 3795
        cmdOK.Top = 2760
        cmdCancel.Top = 2760
        fraLotInformation.Height = 1815
        mbNewForm = False
'-------------------------------------------------------------

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
' Sub: InsNewForm()
'--------------------------------------------------------------------------------
' Description:  <Type your function description here...>
'--------------------------------------------------------------------------------
' Author:       Sam Chen ,Cit 2008/08/26
'--------------------------------------------------------------------------------
' Parameters:
'................................................................................
' ARGUMENTS
'   oLogCtrl            (I) [Optional] Valid reference to a clsLogTraceMsg object
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
'
'================================================================================
Private Sub InsNewForm(ByVal sLotID As String)
On Error GoTo ExitHandler:
Dim sProcID         As String
Dim typErrInfo      As tErrInfo

Dim sSQL            As String
Dim oRS             As Collection
Dim sSQL2            As String
Dim oRs2            As Collection
Dim colSQLResult As Collection
Dim oLot        As FwLot
Dim oEqp    As FwEquipment
Dim sErunTicNO      As String
Dim sProdGroup  As String
Dim sSubSystemType As String
Dim sMaxsite    As String
Dim sPath  As String
Dim sFormNo As String
Dim sTemp As String
Dim sEqComments As String
Dim sWsDeviceFile As String
Dim sProberCard As String
Dim sLoadboard As String
Dim sCable  As String
Dim sTecnNO As String
Dim sPGMMatch As String
Dim sStepComments As String
Dim sProberID As String
Dim sCode As String
Dim sCheckSum As String
Dim sSpeed As String

Dim msOprId As String, oOpr As FwUser

Dim iIdx As Integer

    Dim bTECNLotSpec            As Boolean  'Add by SAM ON 20100407 FOR Project By Lot TECN
    Dim sTecnLotID            As String   'Add by SAM ON 20100407 FOR Project By Lot TECN

    'Added by Jack on 2011/09/23 for JC201100228. 不FollowProduct之工程品 改為 帶出新維修單畫面 <Start>
    Dim sFollowProd                    As String
    Dim sEngNotFollowProd_Code         As String
    Dim sEngNotFollowProd_CheckSum     As String
    Dim sEngNotFollowProd_Speed        As String
    Dim sStepNo                        As String
    Dim sPgName                        As String
    'Added by Jack on 2011/09/23 for JC201100228. 不FollowProduct之工程品 改為 帶出新維修單畫面 <End>

    Dim sPgId                          As String 'Added by Jack on 2019/07/04 for 201900365.

'----
' Init
'----
    sProcID = "InsNewForm"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog) '"Entering Function...", moAppLog, glLOG_PROC, msMODULE_ID, sProcID)
    
    Set oOpr = moFwOPR.ActiveUser
    msOprId = oOpr.UserName
'----
' Condition Checking
'----
    ' <Put your condition checking codes here>...
    Set oLot = FwuRetrieveLot(moFwWIP, sLotID, moAppLog)
    If oLot Is Nothing Then
       UtShowMsgBox "This Lot was not exist. !!" & vbNewLine & _
                        "此Lot不存在 !!"
        cmdOK.Enabled = False
        GoTo ExitHandler
    End If
    
    If oLot.CustomAttributes(gsLOT_CUSTOMATTR_CUR_EQP_ID) <> "" Then
        ' Get TesterId Object
        Set oEqp = FwuRetrieveEqp(moFwMDL, oLot.CustomAttributes(gsLOT_CUSTOMATTR_CUR_EQP_ID), moAppLog)
    Else
        Set oEqp = FwuRetrieveEqp(moFwMDL, oLot.CustomAttributes(gsLOT_CUSTOMATTR_SPLITASSIGNEQID), moAppLog)
    End If
    
    If oEqp Is Nothing Then
        UtShowMsgBox "The Equipment of this Lot was not exited. !!" & vbNewLine & _
                        "此Lot之測試機台不存在 !!"
        cmdOK.Enabled = False
        GoTo ExitHandler
    End If

'----
' Action
'----
    ' <Put your Action codes here>...
    If InStr(cboSubSystem.Text, ",") > 0 Then
       sSubSystemType = Left(cboSubSystem.Text, InStr(cboSubSystem.Text, ",") - 1)
       sMaxsite = Mid$(cboSubSystem.Text, InStr(cboSubSystem.Text, ",") + 1, Len(cboSubSystem.Text) - InStr(cboSubSystem.Text, ","))
    Else
       sSubSystemType = cboSubSystem.Text
    End If
    
    sSQL = "select " & gsCAT_TIM_PROD_GROUP & " from " & _
            gsCAT_TBL_IPN_MASTER & " where " & _
            gsCAT_TIM_IPN & "='" & oLot.CustomAttributes(gsLOT_CUSTOMATTR_IPN) & "'"
    Set oRS = moProRawSql.QueryDatabase(sSQL)
    If oRS.Count > 0 Then
        sProdGroup = oRS.Item(1).Item(gsCAT_TIM_PROD_GROUP)
    End If
    
    sSQL = "SELECT " & gsCAT_TRSL_PATH & " FROM " & gsCAT_TBL_ROUTE_STEP_LIST & _
            " WHERE " & gsCAT_TRSL_ROUTE_ID & "='" & oLot.PlanId & "' AND " & _
            gsCAT_TRSL_FWSTEPSEQ & "='" & oLot.CurrentStep.Steps.Item(1).Handle & "'"
    Set oRS = moProRawSql.QueryDatabase(sSQL)
    If oRS.Count > 0 Then
        sPath = oRS.Item(1).Item(gsCAT_TRSL_PATH)
    End If

'Modify by SAm on 02100514 for Project LOT TECN,GET tecnlotid
'    sSql = "SELECT " & gsCAT_TLI_ERUNTICNO & _
'           " from " & gsCAT_TBL_LOT_INFO & _
'           " where " & gsCAT_TLI_LOT_ID & " ='" & sLotId & "'"
    sSQL = "SELECT " & gsCAT_TLI_TECN_LOT_ID & " , " & gsCAT_TLI_ERUNTICNO & _
           " from " & gsCAT_TBL_LOT_INFO & _
           " where " & gsCAT_TLI_LOT_ID & " ='" & sLotID & "'"

    Set oRS = moProRawSql.QueryDatabase(sSQL)
    If oRS.Count > 0 Then
        sErunTicNO = oRS.Item(1).Item(gsCAT_TLI_ERUNTICNO)
        sTecnLotID = oRS.Item(1).Item(gsCAT_TLI_TECN_LOT_ID)
    End If

    If oEqp.CustomAttributes(gsEQP_CUSTOMATTR_SUBSYS2) = "" Then
        sProberID = oEqp.CustomAttributes(gsEQP_CUSTOMATTR_SUBSYS1)
    Else
        sProberID = oEqp.CustomAttributes(gsEQP_CUSTOMATTR_SUBSYS1) & " " & oEqp.CustomAttributes(gsEQP_CUSTOMATTR_SUBSYS2)
    End If

  'Add by Sam start on 20100514 for Project Lot TECN
    If sTecnLotID <> "" Then
        sSQL = "select " & gsCAT_TLSPES_TEMPERATURE & "," & gsCAT_TLSPES_COMMENTS & _
                "," & gsCAT_TLSPES_WSDEVICEFILE & "," & gsCAT_TLSPES_PROBECARDTYPE & _
                "," & gsCAT_TLSPES_LOADBOARDTYPE & "," & gsCAT_TLSPES_CABLETYPE & _
                "," & gsCAT_TLSPES_SUBSYSTEM & "," & gsCAT_TLSPES_EQTYPE2 & _
                "," & gsCAT_TLSPES_DOCNO & _
                "," & gsCAT_TLSPES_PGMMATCH & _
                " from " & gsCAT_TBL_LOT_STEP_EQ_SPEC & " where " & _
                " '" & sTecnLotID & "' LIKE " & gsCAT_TLSPES_TECNLOTID & " and " & _
                gsCAT_TLSPES_STEPNO & "='" & oLot.CurrentStep.Steps.Item(1).Id & "' and " & _
                gsCAT_TLSPES_EQTYPE2 & "='" & oEqp.CustomAttributes(gsEQP_CUSTOMATTR_EQ_TYPE2) & "' and " & _
                gsCAT_TLSPES_SUBSYSTEM & "='" & sSubSystemType & "' and " & _
                gsCAT_TLSPES_DELETEFLAG & "='N' and " & gsCAT_TLSPES_PATH & "='" & sPath & "' AND " & _
               " nvl(" & gsCAT_TLSPES_MAXSITE & ",' ') = nvl('" & sMaxsite & "',' ') "
    
        Set oRS = moProRawSql.QueryDatabase(sSQL)
        If oRS.Count > 0 Then
            sTemp = oRS.Item(1).Item(gsCAT_TLSPES_TEMPERATURE)
            sEqComments = oRS.Item(1).Item(gsCAT_TLSPES_COMMENTS)
            sWsDeviceFile = oRS.Item(1).Item(gsCAT_TLSPES_WSDEVICEFILE)
            sProberCard = oRS.Item(1).Item(gsCAT_TLSPES_PROBECARDTYPE)
            sLoadboard = oRS.Item(1).Item(gsCAT_TLSPES_LOADBOARDTYPE)
            sCable = oRS.Item(1).Item(gsCAT_TLSPES_CABLETYPE)
            sPGMMatch = oRS.Item(1).Item(gsCAT_TLSPES_PGMMATCH)
            sTecnNO = oRS.Item(1).Item(gsCAT_TLSPES_DOCNO)
           
        End If
    
        'Add by Sam Start on 2018/06/05 for  TECN 自動化
        If msPGMTECNNo <> "" Then
            sTecnNO = msPGMTECNNo & "+" & sTecnNO
        End If
        'Add by Sam End on 2018/06/05 for  TECN 自動化
    
        sSQL = "select " & gsCAT_TLSPS_COMMENTS & _
               "," & gsCAT_TLSPS_DOCNO & _
               " from " & gsCAT_TBL_LOT_STEP_SPEC & " where " & _
               " '" & sTecnLotID & "' LIKE " & gsCAT_TLSPS_TECNLOTID & " and " & _
               gsCAT_TLSPS_STEPNO & "='" & oLot.CurrentStep.Steps.Item(1).Id & "' and " & _
               gsCAT_TLSPS_DELETEFLAG & "='N' and " & gsCAT_TLSPS_PATH & " = '" & sPath & "'"
        Set oRS = moProRawSql.QueryDatabase(sSQL)
        If oRS.Count > 0 Then
            sStepComments = oRS.Item(1).Item(gsCAT_TLSPS_COMMENTS)
        End If
    
    
    Else
    'Add by Sam end on 20100514 for Project Lot TECN
    
        sSQL = "select " & gsCAT_PSES_TEMPERATURE & "," & gsCAT_PSES_COMMENTS & _
                "," & gsCAT_PSES_WS_DEVICE_FILE & "," & gsCAT_PSES_PROBECARD_TYPE & _
                "," & gsCAT_PSES_LOADBOARD_TYPE & "," & gsCAT_PSES_CABLE_TYPE & _
                "," & gsCAT_PSES_SUB_SYSTEM & "," & gsCAT_PSES_EQTYPE2 & _
                "," & gsCAT_PSES_DOC_TYPE & "," & gsCAT_PSES_DOC_NO & _
                "," & gsCAT_PSES_PG_MATCH & _
                " from " & gsCAT_TBL_PROD_STEP_EQ_SPEC & " where " & _
                gsCAT_PSES_PROD_GROUP & "='" & sProdGroup & "' and " & _
                gsCAT_PSES_STEP_NO & "='" & oLot.CurrentStep.Steps.Item(1).Id & "' and " & _
                gsCAT_PSES_EQTYPE2 & "='" & oEqp.CustomAttributes(gsEQP_CUSTOMATTR_EQ_TYPE2) & "' and " & _
                gsCAT_PSES_SUB_SYSTEM & "='" & sSubSystemType & "' and " & _
                gsCAT_PSES_DOC_STATUS & "='Active' and " & gsCAT_PSES_PATH & "='" & sPath & "' AND " & _
               " nvl(" & gsCAT_PSES_MAX_SITE & ",' ') = nvl('" & sMaxsite & "',' ') "
               
        Set oRS = moProRawSql.QueryDatabase(sSQL)
        If oRS.Count > 0 Then
            sTemp = oRS.Item(1).Item(gsCAT_PSES_TEMPERATURE)
            sEqComments = oRS.Item(1).Item(gsCAT_PSES_COMMENTS)
            sWsDeviceFile = oRS.Item(1).Item(gsCAT_PSES_WS_DEVICE_FILE)
            sProberCard = oRS.Item(1).Item(gsCAT_PSES_PROBECARD_TYPE)
            sLoadboard = oRS.Item(1).Item(gsCAT_PSES_LOADBOARD_TYPE)
            sCable = oRS.Item(1).Item(gsCAT_PSES_CABLE_TYPE)
            sPGMMatch = oRS.Item(1).Item(gsCAT_PSES_PG_MATCH)
            'Modified by Jack on 2018/02/02 for Project TECN 自動化 <Start>.
            '改為DocType不為空值, 若 TECNNO有值, 要加到 DocNO前頭.
            'If Trim(oRS.Item(1).Item(gsCAT_PSES_DOC_TYPE)) = "TECN" Then
            If Trim(oRS.Item(1).Item(gsCAT_PSES_DOC_TYPE)) <> "" Then
               sTecnNO = oRS.Item(1).Item(gsCAT_PSES_DOC_NO)
            End If
        End If
        
        'Modidied by Jack on 2018/02/02 for Project TECN 自動化 <Start>.
        If msPGMTECNNo <> "" Then
            sTecnNO = msPGMTECNNo & "+" & sTecnNO
        End If
        'Modidied by Jack on 2018/02/02 for Project TECN 自動化 <End>.
        
        sSQL = "select " & gsCAT_PSS_COMMENTS & _
               "," & gsCAT_PSS_DOC_TYPE & _
               "," & gsCAT_PSS_DOC_NO & _
               " from " & gsCAT_TBL_PROD_STEP_SPEC & " where " & _
               gsCAT_PSS_PROD_GROUP & "='" & sProdGroup & "' and " & _
               gsCAT_PSS_STEP_NO & "='" & oLot.CurrentStep.Steps.Item(1).Id & "' and " & _
               gsCAT_PSS_DOC_STATUS & "='Active' and " & gsCAT_PSS_PATH & " = '" & sPath & "'"
        Set oRS = moProRawSql.QueryDatabase(sSQL)
        If oRS.Count > 0 Then
            sStepComments = oRS.Item(1).Item(gsCAT_PSS_COMMENTS)
        End If
    End If
    
        sSQL = "select " & gsCAT_TIM_SPEED & "," & _
                gsCAT_TIM_CODE & "," & _
                gsCAT_TIM_CHECKSUM & " " & _
                " from " & gsCAT_TBL_IPN_MASTER & " where " & _
                gsCAT_TIM_IPN & "='" & oLot.CustomAttributes(gsLOT_CUSTOMATTR_IPN) & "'"
                
        Set oRS = moProRawSql.QueryDatabase(sSQL)
        If oRS.Count > 0 Then
            sCode = Mid(oRS.Item(1).Item(gsCAT_TIM_CODE), 3, 6)
            sCheckSum = oRS.Item(1).Item(gsCAT_TIM_CHECKSUM)
            sSpeed = oRS.Item(1).Item(gsCAT_TIM_SPEED)
        End If

        'Check MCD
        Set oRS = moProRawSql.QueryDatabase("select count(*)" & _
                   " from " & gsCAT_TBL_MULTICODE_IPN & _
                   " where " & gsCAT_TMI_IPN & " = '" & oLot.CustomAttributes(gsLOT_CUSTOMATTR_IPN) & "' " & _
                   " and   " & gsCAT_TMI_DELETEFLAG & " = 'N'")
        If oRS.Count > 0 Then
           If Val(oRS.Item(1).Item(1)) > 0 Then
            sCode = "Check Code Server"
            sCheckSum = "Check Code Server"
           End If
        End If
                
        'Check MCP
        sSQL = "select " & gsCAT_TIBOM_IPN & " from " & gsCAT_TBL_IPN_BOM & "  where " & gsCAT_TIBOM_PARENT & " = '" & oLot.CustomAttributes(gsLOT_CUSTOMATTR_IPN) & "' " & _
               "and " & gsCAT_TIBOM_BOM_LEVEL & " = 'FG' and " & gsCAT_TIBOM_DELETE_FLAG & " = 'N'"
        Set oRS = moProRawSql.QueryDatabase(sSQL)
        If oRS.Count > 0 Then
           sSQL2 = "select PACKAGECOMPONENT from " & gsCAT_TBL_PROD_BODY & " where " & gsCAT_TPB_PROD_BODY & " = '" & Left(oRS.Item(1).Item(gsCAT_TIBOM_IPN), 4) & "' " & _
                  "and " & gsCAT_TPB_DELETE_FLAG & " = 'N' "
           Set oRs2 = moProRawSql.QueryDatabase(sSQL2)
           If oRs2.Count > 0 Then
              If Left(oRs2.Item(1).Item("PACKAGECOMPONENT"), 3) = "MCP" Then
                 sCode = "Check Code Server"
                 sCheckSum = "Check Code Server"
              End If
           End If
           Set oRs2 = Nothing
        End If

'Mark by Sam start on 20130222 for ReqNo:JC201300050 ,往下搬
'    sSQL = " select 'WS' || to_char(sysdate, 'YYYYMMDD') || " & _
'            " decode(length(max(substr(" & gsCAT_TWEB_FORMNO & ", 11, 3)) + 1),1, " & _
'            " '00' || (max(substr(" & gsCAT_TWEB_FORMNO & ", 11, 3)) + 1),2, " & _
'            " '0' || (max(substr(" & gsCAT_TWEB_FORMNO & ", 11, 3)) + 1),null,'001', " & _
'            " max(substr(" & gsCAT_TWEB_FORMNO & ", 11, 3)) + 1) " & _
'            " From " & gsCAT_TBL_WS_EQFORM_BASIC & "  " & _
'            " where " & gsCAT_TWEB_FORMTYPE & " = 'SETUP' " & _
'            " and " & gsCAT_TWEB_FORMNO & " like 'WS' || to_char(sysdate,'yyyymmdd') ||'%' "
'    Set oRS = moProRawSql.QueryDatabase(sSQL)
'    If oRS.Count > 0 Then
'        sFormNo = oRS.Item(1).Item(1)
'    End If
'Mark by Sam end on 20130222 for ReqNo:JC201300050
    
    
    'Add by Sam start on 20171121 for Project TECN 自動化,改成使用與停復測相同的來源
    If Trim(txtTemp.Text) <> "" Then
        sTemp = txtTemp.Text
    End If
    'Add by Sam end on 20171121 for Project TECN 自動化,改成使用與停復測相同的來源
    
    'Added by Jack on 2011/09/23 for JC201100228. 不FollowProduct之工程品 改為 帶出新維修單畫面 <Start>
    sPgName = txtPGName.Text
    sPgId = txtPgID.Text                          'Added by Jack on 2019/07/04 for 201900365.
    If Trim(sErunTicNO) <> "" Then
        sSQL = "SELECT " & gsCAT_TER_FOLLOW_PRODUCT & _
                  "," & gsCAT_TER_SPEED & "," & gsCAT_TER_CODE & "," & gsCAT_TER_CHECKSUM & _
                  " from " & gsCAT_TBL_ERUN_REQ & _
                  " where " & gsCAT_TER_LOT_ID & " ='" & oLot.Id & "'" & _
                  " and " & gsCAT_TER_TICKET_NO & "='" & sErunTicNO & "' " & _
                  " and " & gsCAT_TER_STAGE & "='WS' "
        Set oRS = moProRawSql.QueryDatabase(sSQL)
        If oRS.Count > 0 Then
            sFollowProd = oRS.Item(1).Item(gsCAT_TER_FOLLOW_PRODUCT)
            sEngNotFollowProd_Code = oRS.Item(1).Item(gsCAT_TER_CODE)
            sEngNotFollowProd_CheckSum = oRS.Item(1).Item(gsCAT_TER_CHECKSUM)
            sEngNotFollowProd_Speed = oRS.Item(1).Item(gsCAT_TER_SPEED)
            '不FollowProduct之工程品資料改帶 Tbl_Erun_Req & Tbl_Erun_Recipe.
            If Trim(sFollowProd) = "N" Then
                '<1> 先將資料清空.
                sPgName = ""
                sPgId = "" 'Added by Jack on 2019/07/04 for 201900365.
                sPGMMatch = ""
                sTemp = ""
                sWsDeviceFile = ""
                sTecnNO = ""
                sEqComments = ""
                sStepComments = ""
                sCode = ""
                sCheckSum = ""
                sSpeed = ""
                sProberCard = ""
                sLoadboard = ""
                
                sStepNo = oLot.CurrentStep.Steps.Item(1).Id
                
                '<2> 再將資料以Tbl_Erun_Req & Tbl_Erun_Recipe取代.
                'sPGMMatch,sTecnNO,sEqComments,sStepComments,Cable空值 --> Request by AlanWu(吳宗勝 : 2011/09/23 15:18)
                'gsCAT_TERE_PGID 'Added by Jack on 2019/07/04 for 201900365.
                sCode = sEngNotFollowProd_Code
                sCheckSum = sEngNotFollowProd_CheckSum
                sSpeed = sEngNotFollowProd_Speed
                sSQL = "SELECT " & gsCAT_TERE_PGNAME & _
                          "," & gsCAT_TERE_TEMPERATURE & "," & gsCAT_TERE_WSDEVICEFILE & _
                          "," & gsCAT_TERE_PROBECARDTYPE & "," & gsCAT_TERE_LOADBOARDTYPE & _
                          "," & gsCAT_TERE_PGID & _
                          " from " & gsCAT_TBL_ERUN_RECIPE & _
                          " where " & gsCAT_TERE_STEPNO & " ='" & sStepNo & "'" & _
                          " and " & gsCAT_TERE_DOCNO & "='" & sErunTicNO & "' " & _
                          " and eqtype2 = '" & oEqp.CustomAttributes(gsEQP_CUSTOMATTR_EQ_TYPE2) & "' "
                If Trim(sSubSystemType) = "" Then
                    sSQL = sSQL & " and subsystem is null "
                Else
                    sSQL = sSQL & " and subsystem='" & Trim(sSubSystemType) & "' "
                End If
                sSQL = sSQL & " and deleteflag = 'N' "
                          
               Set oRS = moProRawSql.QueryDatabase(sSQL)
               If oRS.Count > 0 Then
                    sPgName = oRS.Item(1).Item(gsCAT_TERE_PGNAME)
                    sTemp = oRS.Item(1).Item(gsCAT_TERE_TEMPERATURE)
                    sWsDeviceFile = oRS.Item(1).Item(gsCAT_TERE_WSDEVICEFILE)
                    sProberCard = oRS.Item(1).Item(gsCAT_TERE_PROBECARDTYPE)
                    sPgId = oRS.Item(1).Item(gsCAT_TERE_PGID) 'Added by Jack on 2019/07/04 for 201900365.
               End If
            End If
        End If
    End If
    'Added by Jack on 2011/09/23 for JC201100228. 不FollowProduct之工程品 改為 帶出新維修單畫面 <End>


'Add by Sam start on 20130222 for ReqNo:JC201300050,若 Insert fail 則重新取單號寫入,最多20次,每失敗一次delay 1秒
'Modify by Sam on 20140819 for ReqNo:JC201400269,FormNO改為ST開頭
On Error Resume Next
    For iIdx = 1 To 20
'        sSQL = " select 'WS' || to_char(sysdate, 'YYYYMMDD') || " & _
'                " decode(length(max(substr(" & gsCAT_TWEB_FORMNO & ", 11, 3)) + 1),1, " & _
'                " '00' || (max(substr(" & gsCAT_TWEB_FORMNO & ", 11, 3)) + 1),2, " & _
'                " '0' || (max(substr(" & gsCAT_TWEB_FORMNO & ", 11, 3)) + 1),null,'001', " & _
'                " max(substr(" & gsCAT_TWEB_FORMNO & ", 11, 3)) + 1) " & _
'                " From " & gsCAT_TBL_WS_EQFORM_BASIC & "  " & _
'                " where " & gsCAT_TWEB_FORMTYPE & " = 'SETUP' " & _
'                " and " & gsCAT_TWEB_FORMNO & " like 'WS' || to_char(sysdate,'yyyymmdd') ||'%' "
        sSQL = " select 'ST' || to_char(sysdate, 'YYYYMMDD') || " & _
                " decode(length(max(substr(" & gsCAT_TWEB_FORMNO & ", 11, 3)) + 1),1, " & _
                " '00' || (max(substr(" & gsCAT_TWEB_FORMNO & ", 11, 3)) + 1),2, " & _
                " '0' || (max(substr(" & gsCAT_TWEB_FORMNO & ", 11, 3)) + 1),null,'001', " & _
                " max(substr(" & gsCAT_TWEB_FORMNO & ", 11, 3)) + 1) " & _
                " From " & gsCAT_TBL_WS_EQFORM_BASIC & "  " & _
                " where " & gsCAT_TWEB_FORMTYPE & " = 'SETUP' " & _
                " and " & gsCAT_TWEB_FORMNO & " like 'ST' || to_char(sysdate,'yyyymmdd') ||'%' "

        Set oRS = moProRawSql.QueryDatabase(sSQL)
        If oRS.Count > 0 Then
            sFormNo = oRS.Item(1).Item(1)
        End If
        Err.Clear
'Add by Sam end on 20130222 for ReqNo:JC201300050
        sSQL = "insert into " & gsCAT_TBL_WS_EQFORM_BASIC _
                & "(" & gsCAT_TWEB_FORMNO & " , " & _
                 gsCAT_TWEB_FORMTYPE & " , " & _
                 gsCAT_TWEB_IPN & " , " & _
                 gsCAT_TWEB_LOTID & " , " & _
                 gsCAT_TWEB_LOTOWNER & " , " & _
                 gsCAT_TWEB_ERUNTICNO & " , " & _
                 gsCAT_TWEB_WAFERID & " , " & _
                 gsCAT_TWEB_TESTERID & " , " & _
                 gsCAT_TWEB_PROBERID & " , " & _
                 gsCAT_TWEB_TESTMODE & " , " & _
                 gsCAT_TWEB_PGM & " , " & _
                 gsCAT_TWEB_ID & " , " & _
                 gsCAT_TWEB_PGMMATCH & " , " & _
                 gsCAT_TWEB_CODE & " , " & _
                 gsCAT_TWEB_CHECKSUM & " , " & _
                 gsCAT_TWEB_SPEED & " , " & _
                 gsCAT_TWEB_TEMPERATURE & " , " & _
                 gsCAT_TWEB_DEVICEFILE & " , " & _
                 gsCAT_TWEB_TECNNO & " , " & _
                 gsCAT_TWEB_PROBECARD & " , "
        sSQL = sSQL & _
                 gsCAT_TWEB_EQCOMMENT & " , " & _
                 gsCAT_TWEB_STEPCOMMENT & " , " & _
                 gsCAT_TWEB_CREATEUSERID & " , " & _
                 gsCAT_TWEB_CREATETIME & " "
                 
        'Add by Sam Start on 20120409 for ReqNo:JC201200096
        sSQL = sSQL & "," & _
                gsCAT_TWEB_SETUP_REASON
        'Add by Sam End on 20120409 for ReqNo:JC201200096
        
        'Add by Sam Start on 20130819 for ReqNo:JC201300238
        sSQL = sSQL & "," & _
                gsCAT_TWEB_ASSIGNWAFERID_1 & "," & _
                gsCAT_TWEB_ASSIGNWAFERID_2 & "," & _
                gsCAT_TWEB_ASSIGNWAFERID_3 & "," & _
                gsCAT_TWEB_STEPID & "," & _
                gsCAT_TWEB_INSTEPTIME & " "
        'Add by Sam End on 20130819 for ReqNo:JC201300238
        
        'Add by Sam Start on 20170317 for Project 停復測自動化
        sSQL = sSQL & "," & _
                gsCAT_TWEB_STOPINFORMATION
        'Add by Sam End on 20170317 for Project 停復測自動化
        
        sSQL = sSQL & " ) values ( "
        
        'Modified by Jack on 2011/09/23 for JC201100228. 不FollowProduct之工程品 改為 帶出新維修單畫面 <Start>
        'txtPGName.Text改用sPgName代替.
        'modify by Ernest on 2016/04/21 for ReqNo:JC201500378---------start
        'sSQL = sSQL _
        '       & "'" & sFormNo & "', " _
        '        & "'SETUP', " _
        '        & "'" & oLot.CustomAttributes(gsLOT_CUSTOMATTR_IPN) & "', " _
        '        & "'" & oLot.Id & "'," _
        '        & "'" & lblLotOwner.Caption & "', " _
        '        & "'" & sErunTicNo & "', " _
        '        & "'" & txtWaferID.Text & "', " _
        '        & "'" & lblTesterID.Caption & "', " _
        '        & "'" & sProberID & "', " _
        '        & "'" & oLot.CurrentStep.Steps.Item(1).Description & "', " _
        '        & "'" & sPgName & "', " _
        '        & "'" & sPGMMatch & "'," _
        '        & "'" & sCode & "'," _
        '        & "'" & sCheckSum & "'," _
        '        & "'" & sSpeed & "'," _
        '        & "'" & sTemp & "'," _
        '        & "'" & sWsDeviceFile & "'," _
        '        & "'" & sTecnNO & "'," _
        '        & "'" & sProberCard & "',"
                'sPGId Added by Jack on 2019/07/04 for 201900365.
        sSQL = sSQL _
                & "'" & sFormNo & "', " _
                & "'SETUP', " _
                & "'" & oLot.CustomAttributes(gsLOT_CUSTOMATTR_IPN) & "', " _
                & "'" & oLot.Id & "'," _
                & "'" & lblLotOwner.Caption & "', " _
                & "'" & sErunTicNO & "', " _
                & "'" & txtWaferID.Text & "', " _
                & "'" & Me.cboTesterId.Text & "', " _
                & "'" & sProberID & "', " _
                & "'" & oLot.CurrentStep.Steps.Item(1).Description & "', " _
                & "'" & sPgName & "', " _
                & "'" & sPgId & "', " _
                & "'" & sPGMMatch & "'," _
                & "'" & sCode & "'," _
                & "'" & sCheckSum & "'," _
                & "'" & sSpeed & "'," _
                & "'" & sTemp & "'," _
                & "'" & sWsDeviceFile & "'," _
                & "'" & sTecnNO & "'," _
                & "'" & sProberCard & "',"
        'modify by Ernest on 2016/04/21 for ReqNo:JC201500378---------end
        sSQL = sSQL _
                & "'" & Replace(sEqComments, "'", "''") & "'," _
                & "'" & Replace(sStepComments, "'", "''") & "'," _
                & "'" & msOprId & "'," _
                & "to_char(sysdate, 'YYYYMMDD HH24MISS')||'000'"
                
        'Add by Sam Start on 20120409 for ReqNo:JC201200096
        sSQL = sSQL & "," _
                & "'" & Replace(Me.cboSetupReason.Text, "'", "''") & "' "
                
        'Add by Sam End on 20120409 for ReqNo:JC201200096
    
        'Add by Sam Start on 20130819 for ReqNo:JC201300238
        sSQL = sSQL & "," & _
                "'" & Me.txtSetupWaferID(0) & "'," & _
                "'" & Me.txtSetupWaferID(1) & "'," & _
                "'" & Me.txtSetupWaferID(2) & "'," & _
                "'" & Me.txtStepID & "'," & _
                "'" & Me.txtInstepTime & "' "
                
        'Add by Sam End on 20130819 for ReqNo:JC201300238
    
        'Add by Sam Start on 20170317 for Project 停復測自動化
        sSQL = sSQL & "," _
                & "'" & lblStopInfo.Caption & "' "
        'Add by Sam End on 20170317 for Project 停復測自動化
    
        sSQL = sSQL & ") "
        Set colSQLResult = moProRawSql.QueryDatabase(sSQL)
        
'Add by Sam start on 20130222 for ReqNo:JC201300050
        If Err.Number = 0 Then
            Exit For
        End If
        
        Sleep (1000)
    Next
'Add by Sam end on 20130222 for ReqNo:JC201300050
    
On Error GoTo ExitHandler:
    
     'Add by Tony on 2014/03/03,訊息明確化
     Call LogTrace("Rule:" & msMODULE_ID & ",Function:" & sProcID & " insert SQL=" & sSQL, moAppLog)
    
    
    If colSQLResult Is Nothing Then
'        Call RaiseError(glERR_INVALIDOBJECT, _
'                    FormatErrorText(gsETX_INVALIDOBJECT, "Collection"))
        UtShowMsgBox "SETUP FORM 單號 : " & sFormNo & " 寫入失敗,請洽IT人員協助."
    End If
    'modify by Ernest on 2016/04/21 for ReqNo:JC201500378-------start
    'sSQL = " update " & gsCAT_TBL_EQ_INFO _
    '     & " set " & gsCAT_TEI_FORM_NO & "='" & sFormNo & "' ," _
    '     & gsCAT_TEI_UPDATE_USER_ID & " = '" & msOprId & "' , " _
    '     & gsCAT_TEI_UPDATE_TIME & " = to_char(sysdate, 'YYYYMMDD HH24MISS')||'000' " _
    '     & "where " & gsCAT_TEI_EQ_ID & "='" & lblTesterID.Caption & "'"
    sSQL = " update " & gsCAT_TBL_EQ_INFO _
         & " set " & gsCAT_TEI_FORM_NO & "='" & sFormNo & "' ," _
         & gsCAT_TEI_UPDATE_USER_ID & " = '" & msOprId & "' , " _
         & gsCAT_TEI_UPDATE_TIME & " = to_char(sysdate, 'YYYYMMDD HH24MISS')||'000' " _
         & "where " & gsCAT_TEI_EQ_ID & "='" & Me.cboTesterId.Text & "'"
    'modify by Ernest on 2016/04/21 for ReqNo:JC201500378---------end
    Set colSQLResult = moProRawSql.QueryDatabase(sSQL)

    If colSQLResult Is Nothing Then
        Call RaiseError(glERR_INVALIDOBJECT, _
                    FormatErrorText(gsETX_INVALIDOBJECT, "Collection"))
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
' Sub: AT3SetupForm()
'--------------------------------------------------------------------------------
' Description:  <Type your Sub description here...>
'--------------------------------------------------------------------------------
' Author:       Sam Chen, CIT 2013/03/27
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
'
'================================================================================
Private Sub AT3SetupForm(ByVal sEqType2 As String, ByRef appXL As Object, ByRef oLot As FwLot, _
                      ByRef iPosTitle As Integer)
On Error GoTo ExitHandler:
Dim sProcID As String
Dim typErrInfo As tErrInfo
Dim sEqID   As String 'Add by Tony on 2014/03/05 for HW SETUP RECIPE COMPARE project.

'----
' Init
'----
    sProcID = "AT3SetupForm"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog)

'----
' Condition Checking
'----
    ' <Put your condition checking codes here>...

'----
' Action
'----
    ' <Put your Action codes here>...

    Dim sSQL As String, oRS1 As Collection, oRs2 As Collection, oRs3 As Collection
    Dim iPosProcess As Integer, iPosComment As Integer, iPosNext As Integer
    Dim msOprId As String, oOpr As FwUser
    Dim msProdgroup As String
    Dim oEqp    As FwEquipment
    Dim msSubsystem1 As String, oSubsystem1   As FwEquipment
    Dim msSubsystem2 As String, oSubsystem2 As FwEquipment
    Dim msPGNAME             As String
    Dim msSubsystemType2     As String
    Dim iStartNum            As Integer
    Dim iEndNum              As Integer
    Dim msTemperature        As String
    Dim msPinCount           As String
    Dim msPackageCode        As String
    Dim msLoadBoardId        As String
    Dim msErunTicNo          As String
    Dim sPgName              As String
    Dim msREF_STEP_NAME_01   As String
    Dim msREF_PG_NAME_01     As String
    Dim msREPLACE_PG_NAME_01 As String
    Dim msREF_STEP_NAME_02   As String
    Dim msREF_PG_NAME_02     As String
    Dim msREPLACE_PG_NAME_02 As String
    Dim msREF_STEP_NAME_03   As String
    Dim msREF_PG_NAME_03     As String
    Dim msREPLACE_PG_NAME_03 As String
    Dim msPbFree             As String
    Dim msLoadBoardType      As String
    Dim msSubsystemType      As String
    Dim msMaxSite            As String
    Dim iIndex               As Integer
    Dim iIdx                 As Integer
    Dim sPath                As String
    Dim sMpType              As String
    Dim sLoadBoardId         As String
    Dim iPosDown            As Integer
            
    Dim bTECNLotSpec          As Boolean
    Dim sTecnLotID            As String
    
    'added by Jack on 2014/01/13 for JC201400011. <Start>
    Dim sPureErunTicNo        As String
    Dim sFollowProd           As String
    'added by Jack on 2014/01/13 for JC201400011. <End>
    
    'add by Ernest on 2016/04/27 for ReqNo:JC201500378-------start
    Dim sProdGroup  As String
    Dim vStepName As Variant
    Dim sStepName As String
    Dim sStepNo  As String
    Dim oRS As Collection
    Dim sErunTicNO As String
    'add by Ernest on 2016/04/27 for ReqNo:JC201500378--------end
    
    'Add by Sam start on 20200428 for Project FT配件資料標準化
    Dim sLoadboard As String
    Dim sContactBoard As String
    Dim sBurnInBoard As String
    Dim colAcc As Collection
    'Add by Sam END on 20200428 for Project FT配件資料標準化
    
    msDocNo = ""
    iPosProcess = iPosTitle + 9
    iPosComment = iPosTitle + 20
    
    iPosDown = 2
    
    msSubsystemType = ""
    msMaxSite = ""
    If InStr(cboSubSystem.Text, ",") > 0 Then
       msSubsystemType = Left(cboSubSystem.Text, InStr(cboSubSystem.Text, ",") - 1)
       msMaxSite = Mid$(cboSubSystem.Text, InStr(cboSubSystem.Text, ",") + 1, Len(cboSubSystem.Text) - InStr(cboSubSystem.Text, ","))
    Else
       msSubsystemType = cboSubSystem.Text
    End If
    
    With appXL
        
        .Sheets(sEqType2).Select
        .Cells.Select
        .selection.NumberFormatLocal = "@"
        .selection.Copy
        .Sheets(1).Select
        .ActiveSheet.Paste
        .Application.CutCopyMode = False
                    
        Set oOpr = moFwOPR.ActiveUser
        msOprId = oOpr.UserName
        .range("C" & CStr(iPosTitle + 2)).Select
        .selection.horizontalalignment = -4131
        .activeCell.FormulaR1C1 = msOprId
        
        sSQL = "select shiftcode from tbl_mxic_emp where empno='" & msOprId & "'"
        Set oRS1 = moProRawSql.QueryDatabase(sSQL)
        If oRS1.Count > 0 Then
            .range("H" & CStr(iPosTitle + 2)).Select
            .selection.horizontalalignment = -4131
            .activeCell.FormulaR1C1 = oRS1.Item(1).Item(1)
        End If
        Set oRS1 = Nothing
        
        .range("R" & CStr(iPosTitle + 2)).Select
        .selection.horizontalalignment = -4131
        .activeCell.FormulaR1C1 = Format(Now, "YYYY/MM/DD HH:MM:SS")
        .range("T" & CStr(iPosTitle)).Select
        .activeCell.FormulaR1C1 = "SETUP FORM"
        
        'Add by Sam start on 20170317 for Project 停復測自動化,增加停測資訊
        .range("G" & CStr(iPosTitle + 4)).Select
        .activeCell.FormulaR1C1 = lblStopInfo.Caption
        'Add by Sam End on 20170317 for Project 停復測自動化
        
        .range("E" & CStr(iPosTitle + 6)).Select
        .activeCell.FormulaR1C1 = oLot.CustomAttributes(gsLOT_CUSTOMATTR_IPN)
        .range("E" & CStr(iPosTitle + 5)).Select
        .activeCell.FormulaR1C1 = oLot.Id
        .range("R" & CStr(iPosTitle + 5)).Select
        .activeCell.FormulaR1C1 = oLot.CustomAttributes(gsLOT_CUSTOMATTR_LOT_OWNER)
        .range("E" & CStr(iPosTitle + 7)).Select
        'modify by Ernest on 2016/04/27 for ReqNo:JC201500378------start
        '.activeCell.FormulaR1C1 = oLot.CustomAttributes(gsLOT_CUSTOMATTR_STEPNAME)
        vStepName = Split(Me.cboStep.Text, " ")
        sStepName = Trim(vStepName(1))
        sStepNo = Trim(vStepName(0))
        .activeCell.FormulaR1C1 = sStepName
        'modify by Ernest on 2016/04/27 for ReqNo:JC201500378--------end
        .range("R" & CStr(iPosTitle + 7)).Select
        'modify by Ernest on 2016/04/27 for ReqNo:JC201500378------start
        'If oLot.CustomAttributes(gsLOT_CUSTOMATTR_CUR_EQP_ID) <> "" Then
        '    .activeCell.FormulaR1C1 = oLot.CustomAttributes(gsLOT_CUSTOMATTR_CUR_EQP_ID)
        '    ' Get TesterId Object
        '    Set oEqp = FwuRetrieveEqp(moFwMDL, oLot.CustomAttributes(gsLOT_CUSTOMATTR_CUR_EQP_ID), moAppLog)
        'Else
        '    .activeCell.FormulaR1C1 = oLot.CustomAttributes(gsLOT_CUSTOMATTR_SPLITASSIGNEQID)
        '    Set oEqp = FwuRetrieveEqp(moFwMDL, oLot.CustomAttributes(gsLOT_CUSTOMATTR_SPLITASSIGNEQID), moAppLog)
        'End If
        .activeCell.FormulaR1C1 = Me.cboTesterId.Text
        Set oEqp = FwuRetrieveEqp(moFwMDL, Me.cboTesterId.Text, moAppLog)
        'modify by Ernest on 2016/04/27 for ReqNo:JC201500378--------end
        sSQL = "select eruNticno,saprwno from tbl_lot_info where lotid='" & oLot.Id & "'"
        Set oRS1 = moProRawSql.QueryDatabase(sSQL)
        If oRS1.Count > 0 Then
            .range("R" & CStr(iPosTitle + 6)).Select
            sPureErunTicNo = oRS1.Item(1).Item("erunticno") 'added by Jack on 2014/01/13 for JC201400011.
            If oRS1.Item(1).Item("erunticno") = "" Then
                msErunTicNo = oRS1.Item(1).Item("saprwno")
                .activeCell.FormulaR1C1 = msErunTicNo
            ElseIf oRS1.Item(1).Item("saprwno") = "" Then
                msErunTicNo = oRS1.Item(1).Item("erunticno")
                .activeCell.FormulaR1C1 = msErunTicNo
            Else
                msErunTicNo = oRS1.Item(1).Item("erunticno") & ", " & oRS1.Item(1).Item("saprwno")
                .activeCell.FormulaR1C1 = msErunTicNo
            End If
        End If
        Set oRS1 = Nothing
                              
        'Add BodySize by Tony Start on 2014/03/10 for HW SETUP RECIPE COMPARE project.
        sSQL = "select speed,code,checksum,prodgroup,packagename,pincount,packagecode,gptype,bodysize from tbl_ipn_master where ipn='" & oLot.CustomAttributes(gsLOT_CUSTOMATTR_IPN) & "'"
        
        Set oRS1 = moProRawSql.QueryDatabase(sSQL)
        If oRS1.Count > 0 Then
        
            'Add by Tony Start on 2014/03/10 for HW SETUP RECIPE COMPARE project.
            '(先存放值,以便後續使用)
            msPackageName = oRS1.Item(1).Item("packagename")
            msPinCount = oRS1.Item(1).Item("pincount")
            msBodySize = oRS1.Item(1).Item("bodysize")
            'Add by Tony Start on 2014/03/10 for HW SETUP RECIPE COMPARE project.
            
            .range("E" & CStr(iPosTitle + 8)).Select
            .activeCell.FormulaR1C1 = oRS1.Item(1).Item("packagename")
                       
            .range("R" & CStr(iPosTitle + 9)).Select
            .activeCell.FormulaR1C1 = Mid(oRS1.Item(1).Item("Code"), 3, 6)
            .range("R" & CStr(iPosTitle + 10)).Select
            .activeCell.FormulaR1C1 = oRS1.Item(1).Item("checksum")
           
            msProdgroup = oRS1.Item(1).Item("prodgroup")
            msPinCount = oRS1.Item(1).Item("PinCount")
            msPackageCode = oRS1.Item(1).Item("PackageCode")
            
        End If
        Set oRS1 = Nothing
        'modify by Ernest on 2016/04/27 for ReqNo:JC201500378--------start
        'sSQL = "SELECT PATH FROM TBL_ROUTE_STEP_LIST WHERE ROUTEID='" & oLot.PlanId & "' AND FWSTEPSEQ='" & oLot.CurrentStep.Steps.Item(1).Handle & "'"
        'Set oRS1 = moProRawSql.QueryDatabase(sSQL)
        'If oRS1.Count > 0 Then
        '    sPath = oRS1.Item(1).Item("PATH")
        'End If
        'Set oRS1 = Nothing
        
        sSQL = " SELECT A." & gsCAT_TLI_TECN_LOT_ID & ", " & _
             gsCAT_TLI_ERUNTICNO & " " & _
           " FROM " & gsCAT_TBL_LOT_INFO & " A " & _
           " WHERE A." & gsCAT_TLI_LOT_ID & " ='" & txtLotID.Text & "' "
        Set oRS = moProRawSql.QueryDatabase(sSQL)
        If oRS.Count > 0 Then
            sTecnLotID = oRS.Item(1).Item(1)
            sErunTicNO = oRS.Item(1).Item(2)
        End If
        If Trim(sErunTicNO) <> "" Then
            sSQL = "SELECT " & gsCAT_TER_FOLLOW_PRODUCT & _
                      " from " & gsCAT_TBL_ERUN_REQ & _
                      " where " & gsCAT_TER_LOT_ID & " ='" & txtLotID.Text & "'" & _
                      " and " & gsCAT_TER_TICKET_NO & "='" & sErunTicNO & "' " & _
                      " and " & gsCAT_TER_STAGE & "='FT'"
            Set oRS = moProRawSql.QueryDatabase(sSQL)
            If oRS.Count > 0 Then
                sFollowProd = oRS.Item(1).Item(1)
            End If
        End If
        sSQL = "select i." & gsCAT_TIM_PROD_GROUP & " from " & gsCAT_TBL_IPN_MASTER & _
               " i,tbl_lot_attribute a where i." & gsCAT_TIM_IPN & " = a.ipn and a.lotid='" & Me.txtLotID.Text & "'"
        Set oRS = moProRawSql.QueryDatabase(sSQL)
        If oRS.Count > 0 Then
            sProdGroup = oRS.Item(1).Item("Prodgroup")
        End If
        sPath = GetPath(Me.txtLotID.Text, moProRawSql, moAppLog)
        sSQL = "select distinct b.eqtype2 from view_b2b_fweqarea b where eqid='" & Me.cboTesterId.Text & "'"
        Set oRS = moProRawSql.QueryDatabase(sSQL)
        If oRS.Count > 0 Then
            sEqType2 = oRS.Item(1).Item("eqtype2")
        End If
        'modify by Ernest on 2016/04/27 for ReqNo:JC201500378--------end
        If Not oEqp Is Nothing Then
            'Get HandlerId Object
'            If oEqp.CustomAttributes(gsEQP_CUSTOMATTR_SUBSYS2) = "" Then
'                '.activecell.formulaR1C1 = oEqp.CustomAttributes(gsEQP_CUSTOMATTR_SUBSYS1)
'                msSubsystem1 = oEqp.CustomAttributes(gsEQP_CUSTOMATTR_SUBSYS1)
'                Set oSubsystem1 = FwuRetrieveEqp(moFwMDL, msSubsystem1, moAppLog)
'            Else
'                msSubsystem1 = oEqp.CustomAttributes(gsEQP_CUSTOMATTR_SUBSYS1)
'                msSubsystem2 = oEqp.CustomAttributes(gsEQP_CUSTOMATTR_SUBSYS2)
'                Set oSubsystem1 = FwuRetrieveEqp(moFwMDL, msSubsystem1, moAppLog)
'                Set oSubsystem2 = FwuRetrieveEqp(moFwMDL, msSubsystem2, moAppLog)
'
'            End If

            bTECNLotSpec = CheckTECNLotInfo(moAppLog, moFwWIP, moFwWF, moCwMbx, oLot.Id)
            If bTECNLotSpec = True Then
                   sSQL = " SELECT A." & gsCAT_TLI_TECN_LOT_ID & " " & _
                          " FROM " & gsCAT_TBL_LOT_INFO & " A " & _
                          " WHERE A." & gsCAT_TLI_LOT_ID & " ='" & oLot.Id & "' "
                
                    Set oRS1 = moProRawSql.QueryDatabase(sSQL)
                    If oRS1.Count > 0 Then
                      sTecnLotID = oRS1.Item(1).Item(1)
                    End If
                    'modify by Ernest on 2016/04/27 for ReqNo:JC201500378--------start
                    'sSQL = "select Temperature, Comments,loadboardtype," & _
                    '        "CONTACTBOARDTYPE,KITTYPE," & _
                    '        "cabletype,pgname,Subsystem,eqtype2,'TECN' AS doctype,docno," & _
                    '        "pgid,pgname,refpgname01,refpgname02,refpgname03,refstepname01," & _
                    '        "refstepname02,refstepname03,replacepgname01,replacepgname02," & _
                    '        "replacepgname03,NEEDJUMPER,JUMPERPINNO " & _
                    '        "from tbl_lot_step_eq_spec where " & _
                    '        " '" & sTECNLotID & "' like tecnlotid and  " & _
                    '        "stepno='" & oLot.CurrentStep.Steps.Item(1).Id & "' and " & _
                    '        "eqtype2='" & oEqp.CustomAttributes(gsEQP_CUSTOMATTR_EQ_TYPE2) & "'  "
                    '        'and " & _
                    '        '"subsystem ='" & msSubsystemType & "' and " & _
                    '        '" deleteflag ='N' and nvl(maxsite,' ') = nvl('" & msMaxSite & "',' ')"
                    sSQL = "select Temperature, Comments,loadboardtype," & _
                            "CONTACTBOARDTYPE,KITTYPE," & _
                            "cabletype,pgname,Subsystem,eqtype2,'TECN' AS doctype,docno," & _
                            "pgid,pgname,refpgname01,refpgname02,refpgname03,refstepname01," & _
                            "refstepname02,refstepname03,replacepgname01,replacepgname02," & _
                            "replacepgname03,NEEDJUMPER,JUMPERPINNO,burninboard " & _
                            "from tbl_lot_step_eq_spec where " & _
                            " '" & sTecnLotID & "' like tecnlotid and  " & _
                            "stepno='" & sStepNo & "' and " & _
                            "eqtype2='" & sEqType2 & "'  "

                    'modify by Ernest on 2016/04/27 for ReqNo:JC201500378--------start

            Else
                'modify by Ernest on 2016/04/27 for ReqNo:JC201500378--------start
                'sSQL = "select Temperature, Comments,loadboardtype," & _
                '        "CONTACTBOARDTYPE,KITTYPE," & _
                '        "cabletype,pgname,Subsystem,eqtype2,doctype,docno," & _
                '        "pgid,pgname,refpgname01,refpgname02,refpgname03,refstepname01," & _
                '        "refstepname02,refstepname03,replacepgname01,replacepgname02," & _
                '        "replacepgname03,NEEDJUMPER,JUMPERPINNO " & _
                '        "from tbl_prod_step_eq_spec where " & _
                '        "prodgroup='" & msProdGroup & "' and " & _
                '        "stepno='" & oLot.CurrentStep.Steps.Item(1).Id & "' and " & _
                '        "eqtype2='" & oEqp.CustomAttributes(gsEQP_CUSTOMATTR_EQ_TYPE2) & "' and " & _
                '        "docstatus ='Active'" & "  and path='" & sPath & "' "
                '        'and subsystem ='" & msSubsystemType & "' and " & _
                '        '"maxsite = '" & msMaxSite & "'"
                sSQL = "select Temperature, Comments,loadboardtype," & _
                        "CONTACTBOARDTYPE,KITTYPE," & _
                        "cabletype,pgname,Subsystem,eqtype2,doctype,docno," & _
                        "pgid,pgname,refpgname01,refpgname02,refpgname03,refstepname01," & _
                        "refstepname02,refstepname03,replacepgname01,replacepgname02," & _
                        "replacepgname03,NEEDJUMPER,JUMPERPINNO,burninboard " & _
                        "from tbl_prod_step_eq_spec where " & _
                        "prodgroup='" & msProdgroup & "' and " & _
                        "stepno='" & sStepNo & "' and " & _
                        "eqtype2='" & sEqType2 & "' and " & _
                        "docstatus ='Active'" & "  and path='" & sPath & "' "
                'modify by Ernest on 2016/04/27 for ReqNo:JC201500378--------start
            End If
            
            Set oRS1 = moProRawSql.QueryDatabase(sSQL)
            If oRS1.Count > 0 Then
                .range("R" & CStr(iPosTitle + 8)).Select
                '.activeCell.FormulaR1C1 = oRS1.Item(1).Item("Temperature")'Marked by Jack on 2018/02/01 for Project TECN 自動化,改成使用與停復測相同的來源.
                msTemperature = oRS1.Item(1).Item("Temperature")
                'Added by Jack on 2018/02/01 for Project TECN 自動化,改成使用與停復測相同的來源 <Start>
                If Trim(txtTemp.Text) <> "" Then
                    msTemperature = txtTemp.Text
                End If
                .activeCell.FormulaR1C1 = msTemperature
                'Added by Jack on 2018/02/01 for Project TECN 自動化,改成使用與停復測相同的來源 <End>
                
                .range("E" & CStr(iPosTitle + 13 + iPosDown)).Select
                .activeCell.FormulaR1C1 = oRS1.Item(1).Item("Comments")
                
                msSubsystemType2 = oRS1.Item(1).Item("Subsystem")
                
                msLoadBoardId = oRS1.Item(1).Item("Loadboardtype")
                
                'Add by Sam start on 20200428 for Project FT配件資料標準化
                sLoadboard = oRS1.Item(1).Item("Loadboardtype")
                sContactBoard = oRS1.Item(1).Item("CONTACTBOARDTYPE")
                sBurnInBoard = oRS1.Item(1).Item("burninboard")
                If Trim(sLoadboard) = "" And Trim(sContactBoard) = "" And Trim(sBurnInBoard) = "" Then
                    Set colAcc = GetFTAccByPgm(moProRawSql, moAppLog, sStepName, msBodySize, oRS1.Item(1).Item("pgid"), oRS1.Item(1).Item("pgname"))
                    If colAcc.Count > 0 Then
                        sLoadboard = colAcc.Item("loadboard")
                        sContactBoard = colAcc.Item("contactboard")
                        sBurnInBoard = colAcc.Item("burninboard")
                    End If
                End If
                'Add by Sam End on 20200428 for Project FT配件資料標準化
                
                'TECN NO
                'Modified by Jack on 2018/02/02 for Project TECN 自動化 <Start>.
                '改為DocType不為空值, 若 TECNNO有值, 要加到 DocNO前頭.
                'If Trim(oRS1.Item(1).Item("doctype")) = "TECN" Then
                If Trim(oRS1.Item(1).Item("doctype")) <> "" Then
                'Modified by Jack on 2018/02/02 for Project TECN 自動化 <End>.
                   .range("E" & CStr(iPosTitle + 12)).Select
                   '.activeCell.FormulaR1C1 = oRS1.Item(1).Item("docno") 'Maked by Jack on 2018/02/02 for Project TECN 自動化.
                   'Modidied by Jack on 2018/02/02 for Project TECN 自動化 <Start>.
                   If msPGMTECNNo <> "" Then
                        .activeCell.FormulaR1C1 = msPGMTECNNo & "+" & oRS1.Item(1).Item("docno")
                   Else
                        .activeCell.FormulaR1C1 = oRS1.Item(1).Item("docno")
                   End If
                   'Added by Jack on 2018/02/02 for Project TECN 自動化 <End>.
                End If
                'PGM
                sPgName = oRS1.Item(1).Item("PGNAME")
                msREF_STEP_NAME_01 = oRS1.Item(1).Item("refstepname01")
                msREF_PG_NAME_01 = oRS1.Item(1).Item("refpgname01")
                msREPLACE_PG_NAME_01 = oRS1.Item(1).Item("replacepgname01")
                msREF_STEP_NAME_02 = oRS1.Item(1).Item("refstepname02")
                msREF_PG_NAME_02 = oRS1.Item(1).Item("refpgname02")
                msREPLACE_PG_NAME_02 = oRS1.Item(1).Item("replacepgname02")
                msREF_STEP_NAME_03 = oRS1.Item(1).Item("refstepname03")
                msREF_PG_NAME_03 = oRS1.Item(1).Item("refpgname03")
                msREPLACE_PG_NAME_03 = oRS1.Item(1).Item("replacepgname03")
                'modify by Ernest on 2016/04/27 for ReqNo:JC201500378--------start
                'msPGNAME = GetPGM(oLot, sPgName, msREF_STEP_NAME_01, msREF_PG_NAME_01, msREPLACE_PG_NAME_01, _
                '           msREF_STEP_NAME_02, msREF_PG_NAME_02, msREPLACE_PG_NAME_02, _
                '           msREF_STEP_NAME_03, msREF_PG_NAME_03, msREPLACE_PG_NAME_03, _
                '           oEqp.CustomAttributes(gsEQP_CUSTOMATTR_EQ_TYPE2), msSubsystemType)
                msPGNAME = GetPGM(oLot, sPgName, msREF_STEP_NAME_01, msREF_PG_NAME_01, msREPLACE_PG_NAME_01, _
                           msREF_STEP_NAME_02, msREF_PG_NAME_02, msREPLACE_PG_NAME_02, _
                           msREF_STEP_NAME_03, msREF_PG_NAME_03, msREPLACE_PG_NAME_03, _
                           sEqType2, msSubsystemType)

                'Added by Jack on 2018/02/01 for Project TECN 自動化,改成使用與停復測相同的來源 <Start>
                If Trim(txtPGName.Text) <> "" Then
                    msPGNAME = txtPGName.Text
                End If
                'Added by Jack on 2018/02/01 for Project TECN 自動化,改成使用與停復測相同的來源 <End>

                'modify by Ernest on 2016/04/27 for ReqNo:JC201500378----------end
                .range("E" & CStr(iPosTitle + 9)).Select
                .activeCell.FormulaR1C1 = msPGNAME
                
                 '架機完成配件紀錄                                          '
                .range("E" & CStr(iPosTitle + 25 + iPosDown)).Select
                'Modify by Sam on 20200428 for Project FT配件資料標準化
                '.activeCell.FormulaR1C1 = oRS1.Item(1).Item("CONTACTBOARDTYPE")
                .activeCell.FormulaR1C1 = sContactBoard
                 
            End If
            
            sSQL = "select " & gsCAT_TTFA_PITCH & "," & gsCAT_TTFA_VACUUMCUP & " from " & gsCAT_TBL_TESTER_FVI_ACC & _
                        " where " & gsCAT_TTFA_PINCOUNT & "='" & msPinCount & "' " & _
                        " and " & gsCAT_TTFA_PKGCODE & "='" & msPackageCode & "' " & _
                        " and " & gsCAT_TTFA_DELETEFLAG & "='N' "
                        
            Set oRs2 = moProRawSql.QueryDatabase(sSQL)
            If oRs2.Count > 0 Then
               .range("E" & CStr(iPosTitle + 26 + iPosDown)).Select
               .activeCell.FormulaR1C1 = oRs2.Item(1).Item(gsCAT_TTFA_VACUUMCUP)
            End If
            
            Set oRS1 = Nothing
      
            'Subsystem1,Subsystem2 Accessory Data
            sSQL = "select loadboardid, cableid, contactboardid, kitid,vacuumcupz1 from tbl_eq_info where eqid='" & oEqp.Id & "'"
            Set oRS1 = moProRawSql.QueryDatabase(sSQL)
            If oRS1.Count > 0 Then
                'Subsystem1 Accessory Data
                .range("E" & CStr(iPosTitle + 18 + iPosDown)).Select
                If InStr(oRS1.Item(1).Item("contactboardid"), ",") > 0 Then
                    .activeCell.FormulaR1C1 = Left(oRS1.Item(1).Item("contactboardid"), InStr(oRS1.Item(1).Item("contactboardid"), ",") - 1)
                Else
                    .activeCell.FormulaR1C1 = oRS1.Item(1).Item("contactboardid")
                End If
                .range("E" & CStr(iPosTitle + 19 + iPosDown)).Select
                If InStr(oRS1.Item(1).Item("vacuumcupz1"), ",") > 0 Then
                    .activeCell.FormulaR1C1 = Left(oRS1.Item(1).Item("vacuumcupz1"), InStr(oRS1.Item(1).Item("vacuumcupz1"), ",") - 1)
                Else
                    .activeCell.FormulaR1C1 = oRS1.Item(1).Item("vacuumcupz1")
                End If
                
            End If
            Set oRS1 = Nothing
            
            'Step comments, good bin, dgrade bin
            Dim msGood As String
            Dim msDgrade As String
            Dim msBin As String
            Dim msPath As String

            msGood = ""
            msDgrade = ""
            
            sSQL = "select route " & _
                   " from fwadmin.tbl_lot_attribute where lotid='" & txtLotID.Text & "'"
            Set oRS1 = moProRawSql.QueryDatabase(sSQL)
            If oRS1.Count > 0 Then
               msPath = oRS1.Item(1).Item("route")
            End If
            Set oRS1 = Nothing

            bTECNLotSpec = CheckTECNLotInfo(moAppLog, moFwWIP, moFwWF, moCwMbx, oLot.Id)
            If bTECNLotSpec = True Then
                'modify by Ernest on 2016/04/27 for ReqNo:JC201500378--------start
                'sSQL = "select Comments,'TECN' AS  doctype,docno ,bin1,bin2,bin3,bin4,bin5,bin6,bin7,bin8 " & _
                '       "from tbl_Lot_step_spec where " & _
                '       " '" & sTECNLotID & "' like tecnlotid and  " & _
                '       " stepno='" & oLot.CurrentStep.Steps.Item(1).Id & "' and " & _
                '       " deleteflag='N' and path = '" & msPath & "'"
                sSQL = "select Comments,'TECN' AS  doctype,docno ,bin1,bin2,bin3,bin4,bin5,bin6,bin7,bin8 " & _
                       "from tbl_Lot_step_spec where " & _
                       " '" & sTecnLotID & "' like tecnlotid and  " & _
                       " stepno='" & sStepNo & "' and " & _
                       " deleteflag='N' and path = '" & msPath & "'"
                'modify by Ernest on 2016/04/27 for ReqNo:JC201500378----------end
            Else

                'modify by Ernest on 2016/04/27 for ReqNo:JC201500378--------start
                'sSQL = "select Comments,doctype,docno,bin1,bin2,bin3,bin4,bin5,bin6,bin7,bin8 " & _
                '       "from tbl_prod_step_spec where " & _
                '       "prodgroup='" & msProdGroup & "' and " & _
                '       "stepno='" & oLot.CurrentStep.Steps.Item(1).Id & "' and " & _
                '       "docstatus ='Active' and path = '" & msPath & "'"
                sSQL = "select Comments,doctype,docno,bin1,bin2,bin3,bin4,bin5,bin6,bin7,bin8 " & _
                       "from tbl_prod_step_spec where " & _
                       "prodgroup='" & sProdGroup & "' and " & _
                       "stepno='" & sStepNo & "' and " & _
                       "docstatus ='Active' and path = '" & msPath & "'"
                'modify by Ernest on 2016/04/27 for ReqNo:JC201500378----------end
            End If
            Set oRS1 = moProRawSql.QueryDatabase(sSQL)
            
            If oRS1.Count > 0 Then
               .range("E" & CStr(iPosTitle + 14 + iPosDown)).Select
                .activeCell.FormulaR1C1 = oRS1.Item(1).Item("Comments")
               'Mark by Sam start on 20160608 for Project FT DGRADE,取消BIN1~BIN8
'               For iIndex = 1 To 8
'                 msBin = "BIN" + CStr(iIndex)
'                 If oRS1.Item(1).Item(msBin) = "Good" Then
'                    If msGood = "" Then
'                       msGood = CStr(iIndex)
'                    Else
'                       msGood = msGood + "," + CStr(iIndex)
'                    End If
'                 End If
'                 If oRS1.Item(1).Item(msBin) = "Dgrade" Then
'                    If msDgrade = "" Then
'                       msDgrade = CStr(iIndex)
'                    Else
'                       msDgrade = msDgrade + "," + CStr(iIndex)
'                    End If
'                 End If
'               Next
'               .range("R" & CStr(iPosTitle + 11)).Select
'                .activeCell.FormulaR1C1 = msGood
'               .range("R" & CStr(iPosTitle + 12)).Select
'                .activeCell.FormulaR1C1 = msDgrade
               'Mark by Sam END on 20160608 for Project FT DGRADE
            End If
            Set oRS1 = Nothing
        End If
        
        'added by Jack on 2014/01/13 for JC201400011. <Start>
        If Len(sPureErunTicNo) <> 0 Then
            sSQL = "SELECT " & gsCAT_TER_FOLLOW_PRODUCT & _
                      " from " & gsCAT_TBL_ERUN_REQ & _
                      " where " & gsCAT_TER_LOT_ID & " ='" & oLot.Id & "'" & _
                      " and " & gsCAT_TER_TICKET_NO & "='" & sPureErunTicNo & "' " & _
                      " and " & gsCAT_TER_STAGE & "='" & lblStage.Caption & "' " & _
                      " and " & gsCAT_TER_DELETE_FLAG & "='N' "
            Set oRS1 = moProRawSql.QueryDatabase(sSQL)
            If oRS1.Count > 0 Then
                sFollowProd = oRS1.Item(1).Item(1)
            End If
            
            
            If Trim(sPureErunTicNo) <> "" And Trim(sFollowProd) = "N" Then
                '<1> replace speed/code/checksum.
                sSQL = "SELECT " & gsCAT_TER_SPEED & "," & _
                                   gsCAT_TER_CHECKSUM & "," & _
                                   gsCAT_TER_CODE & " " & _
                          " from " & gsCAT_TBL_ERUN_REQ & _
                          " where " & gsCAT_TER_LOT_ID & " ='" & oLot.Id & "'" & _
                          " and " & gsCAT_TER_TICKET_NO & "='" & sPureErunTicNo & "' " & _
                          " and " & gsCAT_TER_STAGE & "='" & lblStage.Caption & "' " & _
                          " and " & gsCAT_TER_DELETE_FLAG & "='N' "
                Set oRS1 = moProRawSql.QueryDatabase(sSQL)
                If oRS1.Count > 0 Then
                    'PROGRAMMER : AT3-300AL / FT-940 no speed.
                    
                    .range("R" & CStr(iPosTitle + 9)).Select 'Code
                    .activeCell.FormulaR1C1 = oRS1.Item(1).Item("code")
                
                    .range("R" & CStr(iPosTitle + 10)).Select 'checksum
                    .activeCell.FormulaR1C1 = oRS1.Item(1).Item("checksum")
                End If
                
                '<2> replace pgname/goodbin/jumperpinno/subsystem/temperature/loadboardtype/contactboardtype
                'modify by Ernest on 2016/04/27 for ReqNo:JC201500378-------start
                'sSQL = "select " & gsCAT_TERE_PGNAME & ", " & _
                '                   gsCAT_TERE_GOODBIN & ", " & _
                '                   gsCAT_TERE_JUMPERPINNO & ", " & _
                '                   gsCAT_TERE_TEMPERATURE & ", " & _
                '                   gsCAT_TERE_LOADBOARDTYPE & ", " & _
                '                   gsCAT_TERE_SUBSYSTEM & ", " & _
                '                   gsCAT_TERE_CONTACTBOARDTYPE & " " & _
                '         "from " & gsCAT_TBL_ERUN_RECIPE & " a " & _
                '        "where a." & gsCAT_TERE_DOCNO & " = '" & sPureErunTicNo & "' " & _
                '          "and a." & gsCAT_TERE_STEPNO & " = '" & oLot.CurrentStep.Steps.Item(1).Id & "' " & _
                '          "and a." & gsCAT_TERE_EQTYPE2 & " = '" & oEqp.CustomAttributes(gsEQP_CUSTOMATTR_EQ_TYPE2) & "' "
                sSQL = "select " & gsCAT_TERE_PGNAME & ", " & _
                                   gsCAT_TERE_GOODBIN & ", " & _
                                   gsCAT_TERE_JUMPERPINNO & ", " & _
                                   gsCAT_TERE_TEMPERATURE & ", " & _
                                   gsCAT_TERE_LOADBOARDTYPE & ", " & _
                                   gsCAT_TERE_SUBSYSTEM & ", " & _
                                   gsCAT_TERE_CONTACTBOARDTYPE & " " & _
                         "from " & gsCAT_TBL_ERUN_RECIPE & " a " & _
                        "where a." & gsCAT_TERE_DOCNO & " = '" & sPureErunTicNo & "' " & _
                          "and a." & gsCAT_TERE_STEPNO & " = '" & sStepNo & "' " & _
                          "and a." & gsCAT_TERE_EQTYPE2 & " = '" & sEqType2 & "' "

                'modify by Ernest on 2016/04/27 for ReqNo:JC201500378----------end
                sSQL = sSQL & " and a.deleteflag = 'N' "
                          
                Set oRS1 = moProRawSql.QueryDatabase(sSQL)
                If oRS1.Count > 0 Then
                    .range("E" & CStr(iPosTitle + 9)).Select 'Program Name:
                    .activeCell.FormulaR1C1 = oRS1.Item(1).Item("pgname")
                    
                    .range("R" & CStr(iPosTitle + 11)).Select 'GoodBin
                    .activeCell.FormulaR1C1 = oRS1.Item(1).Item("goodbin")
                    
                    'PROGRAMMER : AT3-300AL / FT-940 no jumperpinno.
                    'PROGRAMMER : AT3-300AL / FT-940 no subsystem.
                    
                    .range("R" & CStr(iPosTitle + 8)).Select 'Temperature
                    .activeCell.FormulaR1C1 = oRS1.Item(1).Item("temperature")
                    
                    'PROGRAMMER : AT3-300AL / FT-940 no loadboardtype.
                    
                    .range("E" & CStr(iPosTitle + 25 + iPosDown)).Select 'CONTACTBOARDTYPE
                    .activeCell.FormulaR1C1 = oRS1.Item(1).Item("contactboardtype")
                End If
            End If
        End If
        'added by Jack on 2014/01/13 for JC201400011. <End>
        
        'Add by Tony Start on 2014/02/25 for HW SETUP RECIPE COMPARE project
        '(二)設備架機確認(HW Setup Quality check)內容
        'Modify in 2016/04/27 for ReqNO:JC201500378----start
        'sEqID = lblTesterId.Caption
        sEqID = Me.cboTesterId.Text
        'Modify in 2016/04/27 for ReqNO:JC201500378----start
        Call getRecipeSpecName(sEqID, oLot.Id, oLot.CustomAttributes(gsLOT_CUSTOMATTR_IPN), _
                             msPackageName, msPinCount, msBodySize, _
                             oRS1, moProRawSql, moAppLog, _
                              moCwMbx)
        
        If Not oRS1 Is Nothing And oRS1.Count > 0 Then
            For iIdx = 1 To oRS1.Count
                If iIdx <= miAT3_RECIPE_ROW Then
                    .range("A" & CStr(iPosTitle + 37 + iIdx)).Select 'ID
                    .activeCell.FormulaR1C1 = oRS1.Item(iIdx).Item("col1")
                    
                    .range("B" & CStr(iPosTitle + 37 + iIdx)).Select 'Specname
                    .activeCell.FormulaR1C1 = oRS1.Item(iIdx).Item("specname")
                Else
                    If iIdx <= miAT3_RECIPE_MAX_ROW Then
                        .range("P" & CStr(iPosTitle + 37 + iIdx - miAT3_RECIPE_ROW)).Select 'ID
                        .activeCell.FormulaR1C1 = oRS1.Item(iIdx).Item("col1")
                        
                        .range("Q" & CStr(iPosTitle + 37 + iIdx - miAT3_RECIPE_ROW)).Select 'Specname
                        .activeCell.FormulaR1C1 = oRS1.Item(iIdx).Item("specname")
                    End If
                End If
                
            Next
        Else
            UtShowMsgBox "尚未維護，請確認架機需求或請維護相關資訊。請洽設備工程師。" & vbCrLf & "No data maintenance, make sure the equipment frame machine or maintain information. Please call HW to check."
        End If
        'Add by Tony End on 2014/02/25 for HW SETUP RECIPE COMPARE project
        
        ' if msErunTicNo is not Enpity表重工品,則
        '設為空值 : ProgramName, DocumentNo, Temperature, Special Comment, Speed, Code, CheckSum
        'user直接查詢委測/重工單, 避免混淆
'Marked by Jack on 2014/01/13 for JC201400011 <Start>
'Chihhao要求不清空以下欄位內容.
'        If Len(msErunTicNo) <> 0 Then
'            .range("E" & CStr(iPosTitle + 9)).Select  'Program Name:
'            .activeCell.FormulaR1C1 = ""
'            .range("E" & CStr(iPosTitle + 12)).Select  'TECN NO
'            .activeCell.FormulaR1C1 = ""
'            .range("R" & CStr(iPosTitle + 8)).Select  'Temperature
'            .activeCell.FormulaR1C1 = ""
'            .range("E" & CStr(iPosTitle + 15)).Select  'EqComment
'            .activeCell.FormulaR1C1 = ""
'            .range("R" & CStr(iPosTitle + 9)).Select   'code
'            .activeCell.FormulaR1C1 = ""
'            .range("R" & CStr(iPosTitle + 10)).Select  'checksum
'            .activeCell.FormulaR1C1 = ""
'        End If
'Marked by Jack on 2014/01/13 for JC201400011 <End>

    End With

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
' Function: GetSetupWaferID()
'--------------------------------------------------------------------------------
' Description:  <Type your function description here...>
'--------------------------------------------------------------------------------
' Author:       Sam Chen, CIT 2013-08-19
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
Private Function GetSetupWaferID(ByVal sLotID As String) As Boolean
On Error GoTo ExitHandler:
Dim sProcID As String
Dim typErrInfo As tErrInfo
Dim colRS As Collection
Dim sSQL As String

Dim lIdx As Long

Dim oLot As FwLot
Dim sWaferID As String

Dim sTimeheresince As String

Dim oComp As FwComponent

Dim sParentLotID As String
Dim sStepNo As String

Dim iCount As Integer

'Added by Jack on 2016/03/21 for JC201500392 <Start>
Dim sTmpAssignWaferID    As String
Dim sAssignWaferID       As String
Dim sArysAssignWaferID() As String
'Added by Jack on 2016/03/21 for JC201500392 <End>
'add by Ernest on 2016/04/27 for ReqNo:JC201500378----------start
Dim vStepName As Variant
'add by Ernest on 2016/04/27 for ReqNo:JC201500378----------end
'----
' Init
'----
    sProcID = "GetSetupWaferID"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog) '"Entering Function...", moapplog, glLOG_PROC, msMODULE_ID, sProcID)
    GetSetupWaferID = True
    
    Set oLot = moFwWIP.LotById(sLotID)
    If oLot Is Nothing Then
        Call RaiseError(glERR_INVALIDOBJECT, _
                        FormatErrorText(gsETX_INVALIDOBJECT, "FwLot"))
        GoTo ExitHandler
    End If
    
    Me.txtStepID.Text = ""
    Me.txtInstepTime.Text = ""
    Me.txtSetupWaferID(0).Text = ""
    Me.txtSetupWaferID(1).Text = ""
    Me.txtSetupWaferID(2).Text = ""
    
    sParentLotID = Left(sLotID, 8)
    
    'Get Time Here Since
    sTimeheresince = ConvertToString(oLot.CurrentStep.Steps.Item(1).TimeHereSince, moAppLog)
    'Get StepNo
    'modify by Ernest on 2016/04/27 for ReqNo:JC201500378----------start
    'sStepNo = oLot.CurrentStep.Steps.Item(1).Id
    vStepName = Split(Me.cboStep.Text, " ")
    sStepNo = Trim(vStepName(0))
    'modify by Ernest on 2016/04/27 for ReqNo:JC201500378----------start
    Me.txtStepID = sStepNo
    Me.txtInstepTime = sTimeheresince
    
'----
' Condition Checking
'----
    ' <Put your condition checking codes here>...

'----
' Action
'----
    'Marked by Jack on 2016/03/21 for JC201500392 <Start>
    '改為呼叫模組化.
'    'Modified by Jack on 2014/02/14 for MES預警機制專案 <Start>
'    '改為呼叫模組化.
'    sSQL = "   select wafer.lotid , wafer.waferid, wt.TestFlag, " & _
'            "   nvl(a.total, 0) manualQty, " & _
'            "   nvl(b.total, 0) tdsQty, " & _
'            "   nvl(c.total, 0) subQty, " & _
'            "   nvl(d.total, 0) + nvl(e.total, 0) + nvl(f.total, 0) setupQty, " & _
'            "   nvl(a.total, 0) + nvl(b.total, 0) + nvl(c.total, 0) + " & _
'            "   nvl(d.total, 0) + nvl(e.total, 0) + nvl(f.total, 0) actQty "
'
'    sSQL = sSQL & " from " & _
'            "   (SELECT substr(a.appid,1,8) as lotid , c.componentid as waferid  " & _
'            "   FROM fwlot a, fwlot_n2m b, fwcomponent c " & _
'            "   where a.sysid = b.fromid " & _
'            "   and b.toid = c.sysid " & _
'            "   AND A.APPID = '" & sLotID & "') wafer, "
'    sSQL = sSQL & _
'            "   (select c." & gsCAT_TMT_WAFERID & " waferid, 'A' type, count(*) total " & _
'            "   from " & gsCAT_TBL_MANUAL_TESTQTY & " c " & _
'            "   where substr(c." & gsCAT_TMT_LOT_ID & ", 1, 8) = '" & sParentLotID & "' " & _
'            "   and c." & gsCAT_TMT_PROBING_FLAG & "  = 'Y' " & _
'            "   group by c." & gsCAT_TMT_WAFERID & ", 'A') a, " & _
'            "   (select d." & gsCAT_TWTS_WAFER_IDN & " waferid, 'B' type, count(*) total " & _
'            "   from " & gsCAT_TBL_WS_TDS_SUM & " d " & _
'            "   where d." & gsCAT_TWTS_PROBINGLOTID & " = '" & sParentLotID & "' " & _
'            "   and d." & gsCAT_TWTS_PROBING_FLAG & " = 'Y' " & _
'            "   group by d." & gsCAT_TWTS_WAFER_IDN & ", 'B') b, " & _
'            "   (select " & gsCAT_TWPC_WAFERID & ", 'C' type, sum(e.probingcount) total " & _
'            "   from " & gsCAT_TBL_WS_PROBING_COUNT & " e " & _
'            "   where substr(e." & gsCAT_TWPC_LOTID & ", 1, 8) = '" & sParentLotID & "' " & _
'            "   group by " & gsCAT_TWPC_WAFERID & ") c, "
'
'    sSQL = sSQL & _
'            "   (select distinct (" & gsCAT_TWEB_SETUPWAFERID_1 & ") waferid, " & _
'                        "   'D' type, " & _
'                        "   sum(" & gsCAT_TWEB_PROBINGQTY_1 & ") total " & _
'            "   from " & gsCAT_TBL_WS_EQFORM_BASIC & " B1 " & _
'            "   WHERE B1." & gsCAT_TWEB_FORMTYPE & " in ( 'SETUP','REPAIR') " & _
'            "   AND substr(b1." & gsCAT_TWEB_LOTID & ", 1, 8) = '" & sParentLotID & "' " & _
'            "   group by " & gsCAT_TWEB_SETUPWAFERID_1 & ") d, " & _
'            "   (select distinct (" & gsCAT_TWEB_SETUPWAFERID_2 & ") waferid, " & _
'                        "   'D' type, " & _
'                        "   sum(" & gsCAT_TWEB_PROBINGQTY_2 & ") total " & _
'            "   from " & gsCAT_TBL_WS_EQFORM_BASIC & " B2 " & _
'            "   WHERE B2." & gsCAT_TWEB_FORMTYPE & " in ( 'SETUP','REPAIR')  " & _
'            "   AND substr(b2." & gsCAT_TWEB_LOTID & ", 1, 8) = '" & sParentLotID & "' " & _
'            "   group by " & gsCAT_TWEB_SETUPWAFERID_2 & ") e, " & _
'            "   (select distinct (" & gsCAT_TWEB_SETUPWAFERID_3 & ") waferid, " & _
'                        "   'D' type, " & _
'                        "   sum(" & gsCAT_TWEB_PROBINGQTY_3 & ") total " & _
'            "   from " & gsCAT_TBL_WS_EQFORM_BASIC & " B3 " & _
'            "   WHERE B3." & gsCAT_TWEB_FORMTYPE & " in ( 'SETUP','REPAIR')  " & _
'            "   AND substr(b3." & gsCAT_TWEB_LOTID & ", 1, 8) = '" & sParentLotID & "' " & _
'            "   group by " & gsCAT_TWEB_SETUPWAFERID_3 & ") f, "
'
'    sSQL = sSQL & _
'            "   (select cc." & gsCAT_TMT_WAFERID & " AS waferid , substr(cc." & gsCAT_TMT_LOT_ID & ",1,8) as " & gsCAT_TMT_LOT_ID & "  " & _
'            "   from " & gsCAT_TBL_MANUAL_TESTQTY & " cc " & _
'            "   where substr(cc." & gsCAT_TMT_LOT_ID & ", 1, 8) = '" & sParentLotID & "' " & _
'            "   and cc." & gsCAT_TMT_PROBING_FLAG & " = 'Y' " & _
'            "   Union " & _
'            "   select dd." & gsCAT_TWTS_WAFER_IDN & " AS waferid,substr(dd." & gsCAT_TWTS_LOT_ID & ",1,8) as " & gsCAT_TWTS_LOT_ID & " " & _
'            "   from " & gsCAT_TBL_WS_TDS_SUM & " dd " & _
'            "   where " & gsCAT_TWTS_PROBINGLOTID & " = '" & sParentLotID & "' " & _
'            "   and dd." & gsCAT_TWTS_PROBING_FLAG & " = 'Y' " & _
'            "   union " & _
'            "   select ee." & gsCAT_TWPC_WAFERID & ",substr(ee." & gsCAT_TWPC_LOTID & ",1,8) as " & gsCAT_TWPC_LOTID & " " & _
'            "   from " & gsCAT_TBL_WS_PROBING_COUNT & "  ee " & _
'            "   where substr(ee." & gsCAT_TWPC_LOTID & ", 1, 8) = '" & sParentLotID & "' " & _
'            "   union "
'
'    sSQL = sSQL & _
'            "   select " & gsCAT_TWEB_SETUPWAFERID_1 & " , substr(" & gsCAT_TWEB_LOTID & ",1,8) as " & gsCAT_TWEB_LOTID & " " & _
'            "   from " & gsCAT_TBL_WS_EQFORM_BASIC & " " & _
'            "   WHERE " & gsCAT_TWEB_FORMTYPE & " in ( 'SETUP','REPAIR')  " & _
'            "   and substr(" & gsCAT_TWEB_LOTID & ", 1, 8) = '" & sParentLotID & "' " & _
'            "   and " & gsCAT_TWEB_SETUPWAFERID_1 & " is not null  " & _
'            "   union " & _
'            "   select " & gsCAT_TWEB_SETUPWAFERID_2 & " , substr(" & gsCAT_TWEB_LOTID & ",1,8) as " & gsCAT_TWEB_LOTID & "  " & _
'            "   from " & gsCAT_TBL_WS_EQFORM_BASIC & " " & _
'            "   WHERE " & gsCAT_TWEB_FORMTYPE & " in ( 'SETUP','REPAIR')  " & _
'            "   and substr(" & gsCAT_TWEB_LOTID & ", 1, 8) = '" & sParentLotID & "' " & _
'            "   and " & gsCAT_TWEB_SETUPWAFERID_2 & " is not null  " & _
'            "   union " & _
'            "   select " & gsCAT_TWEB_SETUPWAFERID_3 & " , substr(" & gsCAT_TWEB_LOTID & ",1,8) as " & gsCAT_TWEB_LOTID & "  " & _
'            "   from " & gsCAT_TBL_WS_EQFORM_BASIC & " " & _
'            "   WHERE " & gsCAT_TWEB_FORMTYPE & " in ( 'SETUP','REPAIR')  " & _
'            "   and substr(" & gsCAT_TWEB_LOTID & ", 1, 8) = '" & sParentLotID & "' " & _
'            "   and  " & gsCAT_TWEB_SETUPWAFERID_3 & " is not null " & _
'            "   ) ae       "
'    sSQL = sSQL & _
'            "   ,(select " & gsCAT_TWTS_PARENT_LOT_ID & " , " & gsCAT_TWTS_WAFER_IDN & "," & _
'            "   decode(nvl(" & gsCAT_TWTS_REVISE_GOOD_DIEN & ", " & gsCAT_TWTS_TDS_GOOD_DIEN & "),null,null,'Y')  TestFlag  " & _
'            "   from " & gsCAT_TBL_WS_TDS_SUM & " " & _
'            "   where (" & gsCAT_TWTS_TEST_MODE & "," & gsCAT_TWTS_PARENT_LOT_ID & " , " & _
'            " " & gsCAT_TWTS_WAFER_IDN & ", " & gsCAT_TWTS_STEPNO & ", " & gsCAT_TWTS_INSTEPTIME & ", " & _
'            "   nvl(" & gsCAT_TWTS_TIME_STAMP & ", 'NULL')) IN " & _
'            "   (SELECT " & gsCAT_TWTS_TEST_MODE & ", " & _
'            "   " & gsCAT_TWTS_PARENT_LOT_ID & ", " & _
'            "   " & gsCAT_TWTS_WAFER_IDN & ", " & _
'            "    " & gsCAT_TWTS_STEPNO & ", " & _
'            "    " & gsCAT_TWTS_INSTEPTIME & ", " & _
'            "   NVL(MAX(" & gsCAT_TWTS_TIME_STAMP & "), 'NULL') " & _
'            "   FROM " & gsCAT_TBL_WS_TDS_SUM & "  " & _
'            "   WHERE " & gsCAT_TWTS_INSTEPTIME & " = '" & sTimeheresince & "' " & _
'            "   and " & gsCAT_TWTS_STEPNO & " = '" & sStepNo & "' " & _
'            "   and " & gsCAT_TWTS_PARENT_LOT_ID & " = '" & sParentLotID & "' " & _
'            "   and " & gsCAT_TWTS_DELETE_FLAG & " = 'N' " & _
'            "   GROUP BY " & gsCAT_TWTS_TEST_MODE & ", " & gsCAT_TWTS_PARENT_LOT_ID & ", " & gsCAT_TWTS_WAFER_IDN & _
'            " , " & gsCAT_TWTS_STEPNO & ",  " & gsCAT_TWTS_INSTEPTIME & ")) wt  "
'
'    sSQL = sSQL & _
'            "   where wafer.lotid = ae.lotid(+)  and wafer.waferid = ae.waferid(+)  " & _
'            "   and wafer.waferid = a. waferid(+) " & _
'            "   and wafer.waferid = b.waferid(+) " & _
'            "   and wafer.waferid = c.waferid(+) " & _
'            "   and wafer.waferid = d.waferid(+) " & _
'            "   and wafer.waferid = e.waferid(+) " & _
'            "   and wafer.waferid = f.waferid(+) " & _
'            "   and wafer.lotid = wt. " & gsCAT_TWTS_PARENT_LOT_ID & "(+) " & _
'            "   and wafer.waferid  = wt." & gsCAT_TWTS_WAFER_IDN & "(+) " & _
'            "   and wt.TestFlag is null " & _
'            "   ORDER BY actQty,setupQty,WAFERID   "
'    '因為2014 Q1專案互卡, 先取消模組化的部份.
'    'Call modQuery.getActProbingQtySQLString(moProRawSql, sLotID, sSQL, "PrintSetupForm", sTimeheresince, sStepNo)
'    'Modified by Jack on 2014/02/14 for MES預警機制專案 <End>
    'Marked by Jack on 2016/03/21 for JC201500392 <Start>
    
    'Added by Jack on 2016/03/21 for JC201500392 <Start>
    '改為呼叫模組化.
    sSQL = "select Fun_GetAssignWaferID('" & sLotID & "') as GetAssignWaferID from dual"
    Set colRS = modRawSQL.SelectRawSqlTxn1(moAppLog, moFwWIP, moFwWF, moCwMbx, "CAT", sSQL)
    If colRS.Count > 0 Then
        sTmpAssignWaferID = colRS.Item(1).Item(1)
    End If
    '判斷是否KGD & <=3片 ??
    '回傳 : sTmpAssignWaferID = "Y;CORR(F8-26E),CORR(F8-18E),CORR(F8-8C WLCSP)"
    If Mid(sTmpAssignWaferID, 1, 1) = "Y" Then
        sAssignWaferID = Mid(sTmpAssignWaferID, 3) '"CORR(F8-26E),CORR(F8-18E),CORR(F8-8C WLCSP)"
        sArysAssignWaferID() = Split(sAssignWaferID, ",")
        iCount = 0
        For lIdx = LBound(sArysAssignWaferID) To UBound(sArysAssignWaferID)
            txtSetupWaferID(iCount).Text = sArysAssignWaferID(lIdx)
            iCount = iCount + 1
            If iCount >= 3 Then Exit For '最多3個, 不跳走會程式執行失敗(Index Out Of Bound).
        Next
    Else
        Call modQuery.getActProbingQtySQLString(moProRawSql, sLotID, sSQL, "PrintSetupForm", sTimeheresince, sStepNo)
    'Added by Jack on 2016/03/21 for JC201500392 <End>
    
        Set colRS = moProRawSql.QueryDatabase(sSQL)
        iCount = 0
        For lIdx = 1 To colRS.Count
            If colRS.Item(lIdx).Item("TestFlag") <> "Y" Then
                txtSetupWaferID(iCount).Text = colRS.Item(lIdx).Item("waferid")
                iCount = iCount + 1
            End If
            If iCount >= 3 Then Exit For
        Next
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
        GetSetupWaferID = False
    End If
End Function



'add by Ernest on 2016/04/27 for ReqNo:JC201500378----start
Private Function GetPath(sLotID As String, ByRef moProRawSql As Object, ByRef moAppLog As Object) As String
On Error GoTo ExitHandler:
Dim sProcID As String
Dim typErrInfo As tErrInfo

Dim sSQL As String
Dim colSQLResult As Collection
'----
' Init
'----
    sProcID = "GetPath"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog)
'----
' Condition Checking
'----
    ' <Put your condition checking codes here>...
    
'----
' Action
'----

    
    sSQL = "select route from tbl_lot_attribute where lotid = '" & sLotID & "'"
    Set colSQLResult = moProRawSql.QueryDatabase(sSQL)
    If colSQLResult.Count > 0 Then
       GetPath = colSQLResult.Item(1).Item("route")
    End If
    
'----
' Done
'----
ExitHandler:
    ' NOTE 1:
    ' MUST CALL GetErrInfo() here first before another action
    Call GetErrInfo(msMODULE_ID, sProcID, typErrInfo, Erl)
    Call LogProcOut(msMODULE_ID, sProcID, typErrInfo, moAppLog)
    ' cleaning up

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
            Case glERR_ERRMSG
                typErrInfo.sUserText = typErrInfo.sErrDescription
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
Private Function GetStopInfoPByAcc(ByVal sLotID As String, ByVal sProdGroup As String, ByVal sEqType2 As String, _
                                    ByVal sCurEqId As String, ByVal sStepName As String, _
                                    ByVal sPgId As String, sPgName As String, ByVal sPgMode As String, _
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
      "   and '" & sProdGroup & "'= c.prodgroup" & _
      "   and nvl(c.ipn, ' ') like nvl(REPLACE(REPLACE(d." & gsCAT_TSTT_IPN & ", '%', '_'),'*','%'), nvl(c.ipn, ' '))" & _
      "   and nvl('" & sPgId & "', ' ') = nvl(d." & gsCAT_TSTT_PGID & ", nvl('" & sPgId & "', ' '))" & _
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


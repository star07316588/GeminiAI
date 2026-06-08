VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmSimpleLotQuery 
   Caption         =   "SimpleLotQuery"
   ClientHeight    =   7905
   ClientLeft      =   150
   ClientTop       =   345
   ClientWidth     =   12030
   LinkTopic       =   "Form1"
   ScaleHeight     =   7905
   ScaleWidth      =   12030
   StartUpPosition =   1  '所屬視窗中央
   Begin VB.Frame fraSelCriteria 
      Caption         =   "Select Criteria to query"
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
      Height          =   1560
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   11835
      Begin VB.TextBox txtMonocotsLot 
         Height          =   315
         Left            =   6660
         TabIndex        =   28
         Top             =   1200
         Width           =   2500
      End
      Begin VB.TextBox txtVirtualLotID 
         Height          =   315
         Left            =   1980
         TabIndex        =   26
         Top             =   1200
         Width           =   2500
      End
      Begin VB.CommandButton cmdQueryTestFlow 
         Caption         =   "Query TestFlow"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   9360
         TabIndex        =   25
         Top             =   480
         Width           =   2205
      End
      Begin VB.ComboBox cboLocation 
         Height          =   300
         Left            =   6660
         TabIndex        =   23
         Top             =   900
         Width           =   2500
      End
      Begin VB.ComboBox cboHoldType 
         Height          =   315
         Left            =   1980
         TabIndex        =   21
         Top             =   900
         Width           =   2500
      End
      Begin VB.CommandButton cmdQuery 
         Caption         =   "Query Confirm"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   9360
         TabIndex        =   9
         Top             =   960
         Width           =   2205
      End
      Begin VB.ComboBox cboStep 
         Height          =   315
         Left            =   6660
         Style           =   2  '單純下拉式
         TabIndex        =   8
         Top             =   270
         Width           =   2505
      End
      Begin VB.TextBox txtLotID 
         Height          =   315
         Left            =   1980
         TabIndex        =   7
         Top             =   300
         Width           =   2500
      End
      Begin VB.TextBox txtIPN 
         Height          =   315
         Left            =   1980
         TabIndex        =   6
         Top             =   600
         Width           =   2500
      End
      Begin VB.ComboBox cboStage 
         Height          =   315
         Left            =   6660
         TabIndex        =   5
         Top             =   600
         Width           =   2500
      End
      Begin VB.Label lblMonocotsLot 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "Monocots ParentLotId:"
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
         Left            =   4635
         TabIndex        =   29
         Top             =   1215
         Width           =   1935
      End
      Begin VB.Label lblVirtualLotId 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "Virtual Lot Id:"
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
         Left            =   660
         TabIndex        =   27
         Top             =   1215
         Width           =   1155
      End
      Begin VB.Label lblLocation 
         AutoSize        =   -1  'True
         Caption         =   "Location:"
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
         Left            =   5775
         TabIndex        =   24
         Top             =   960
         Width           =   795
      End
      Begin VB.Label lblHoldType 
         AutoSize        =   -1  'True
         Caption         =   "Hold Type:"
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
         Left            =   855
         TabIndex        =   20
         Top             =   920
         Width           =   945
      End
      Begin VB.Label lblLotId 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         Caption         =   "Lot Id:"
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
         Left            =   900
         TabIndex        =   13
         Top             =   315
         Width           =   930
      End
      Begin VB.Label lblIPN 
         AutoSize        =   -1  'True
         Caption         =   "IPN(建議用全碼):"
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
         Left            =   405
         TabIndex        =   12
         Top             =   630
         Width           =   1410
      End
      Begin VB.Label lblStepTill 
         AutoSize        =   -1  'True
         Caption         =   "Step:"
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
         Left            =   6100
         TabIndex        =   11
         Top             =   315
         Width           =   465
      End
      Begin VB.Label lblStage 
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
         Left            =   6015
         TabIndex        =   10
         Top             =   600
         Width           =   570
      End
   End
   Begin VB.Frame fraLotInfo 
      Caption         =   "Wip Information"
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
      Height          =   5460
      Left            =   0
      TabIndex        =   2
      Top             =   1485
      Width           =   11835
      Begin FPSpreadADO.fpSpread spdHoldLot 
         Height          =   3075
         Left            =   7920
         TabIndex        =   22
         Top             =   2250
         Visible         =   0   'False
         Width           =   2850
         _Version        =   524288
         _ExtentX        =   5027
         _ExtentY        =   5424
         _StockProps     =   64
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   2
         MaxRows         =   0
         SpreadDesigner  =   "frmSimpleLotQuery.frx":0000
         AppearanceStyle =   0
      End
      Begin FPSpreadADO.fpSpread spdLotInfo 
         Height          =   5055
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   11595
         _Version        =   524288
         _ExtentX        =   20452
         _ExtentY        =   8916
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MaxCols         =   38
         MaxRows         =   5
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         ScrollBarExtMode=   -1  'True
         SpreadDesigner  =   "frmSimpleLotQuery.frx":0451
         StartingColNumber=   0
         AppearanceStyle =   0
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   10560
      TabIndex        =   1
      Top             =   7365
      Width           =   1095
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Query Print"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   45
      TabIndex        =   0
      Top             =   7365
      Width           =   1500
   End
   Begin VB.Label lblLabelTotalLotCount 
      Caption         =   "Total Lot Count:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   7005
      Width           =   1455
   End
   Begin VB.Label lblTotalLotCount 
      BorderStyle     =   1  '單線固定
      Height          =   255
      Left            =   1560
      TabIndex        =   18
      Top             =   7005
      Width           =   1095
   End
   Begin VB.Label lblLabelTotalWQty 
      Caption         =   "Total WQty:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      TabIndex        =   17
      Top             =   7005
      Width           =   1095
   End
   Begin VB.Label lblTotalWQty 
      BorderStyle     =   1  '單線固定
      Height          =   255
      Left            =   3960
      TabIndex        =   16
      Top             =   7005
      Width           =   975
   End
   Begin VB.Label lblLabelTotalCQty 
      Caption         =   "Total CQty:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5040
      TabIndex        =   15
      Top             =   7005
      Width           =   975
   End
   Begin VB.Label lblTotalCQty 
      BorderStyle     =   1  '單線固定
      Height          =   255
      Left            =   6120
      TabIndex        =   14
      Top             =   7005
      Width           =   1215
   End
End
Attribute VB_Name = "frmSimpleLotQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const msMODULE_ID As String = "frmSimpleLotQuery"
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

Private Const miSpdFieldPos_LotId       As Integer = 1
Private Const miSpdFieldPos_IPN         As Integer = 2
Private Const miSpdFieldPos_Status      As Integer = 3
Private Const miSpdFieldPos_HotFlag     As Integer = 4
Private Const miSpdFieldPos_LotOwner    As Integer = 5
Private Const miSpdFieldPos_WaferQty    As Integer = 6
Private Const miSpdFieldPos_ChipQty     As Integer = 7
Private Const miSpdFieldPos_Step        As Integer = 8

'Marked by Jack on 2014/04/07 for JC201400114 (9~12-->25~27) <Start>
'將LocId / Packing / CarrierType欄位移至最後
'Private Const miSpdFieldPos_LocID       As Integer = 9     'add by Sam on 20090310 for ReqNo:JC900022,add LocID
'
'Private Const miSpdFieldPos_Packing       As Integer = 10     'add by Sam on 20140304 for ReqNo:JC201400057
'Private Const miSpdFieldPos_CarrierType   As Integer = 11     'add by Sam on 20140304 for ReqNo:JC201400057
'Marked by Jack on 2014/04/07 for JC201400114 (9-->25) <End>

'Modified by Jack on 2014/04/07 for JC201400114 (12~27-->9-->24) <Start>
Private Const miSpdFieldPos_HoldCode    As Integer = 9
Private Const miSpdFieldPos_HoldReason  As Integer = 10
Private Const miSpdFieldPos_FabOutDate  As Integer = 11 'Add by Robin at 20150210 for JC201500043, 以下序號順延 AtDays to CarrierType 11~27 -> 12~28
Private Const miSpdFieldPos_AtDays      As Integer = 12
Private Const miSpdFieldPos_Stage       As Integer = 13
Private Const miSpdFieldPos_ErunTickNo            As Integer = 14
Private Const miSpdFieldPos_RwSapNo               As Integer = 15

Private Const miSpdFieldPos_MFGTICNO              As Integer = 16 'add by Sam on 20090310 for ReqNo:JC900022,add MFGTICNO

Private Const miSpdFieldPos_TestSite              As Integer = 17
Private Const miSpdFieldPos_TurnkeyVendor         As Integer = 18
Private Const miSpdFieldPos_FGIpn                 As Integer = 19

Private Const miSpdFieldPos_Location              As Integer = 20    'Added by Jack on 2010/05/05 for JC201000081. <Start>

Private Const miSpdFieldPos_UnpackTime            As Integer = 21    'Added by Sam on 20101011 for JC201000227
Private Const miSpdFieldPos_Remain_UnpackTime     As Integer = 22    'Added by Sam on 20101011 for JC201000227
Private Const miSpdFieldPos_WaferSize             As Integer = 23 'Add by Tony on 20110616 for Req.JC201100154.

'Added by Jack on 2013/02/25 for JC201300058. <Start>
'Modified by Jack on 2013/03/08 for PAMS Realignment Project
'GradeLevel --> LotGrade(加註解)
Private Const miSpdFieldPos_DateCode              As Integer = 24
Private Const miSpdFieldPos_GradeLevel            As Integer = 25
'Added by Jack on 2013/02/25 for JC201300058. <End>
'Modified by Jack on 2014/04/07 for JC201400114 (9-->25) <End>

'Added by Jack on 2014/04/07 for JC201400114 <Start>
'將LocId / Packing / CarrierType欄位移至最後
Private Const miSpdFieldPos_LocID       As Integer = 26
Private Const miSpdFieldPos_Packing     As Integer = 27     'add by Sam on 20140304 for ReqNo:JC201400057
Private Const miSpdFieldPos_CarrierType As Integer = 28     'add by Sam on 20140304 for ReqNo:JC201400057
'Added by Jack on 2014/04/07 for JC201400114 <End>

Private Const miSpdFieldPos_FailDgrade  As Integer = 29     'Added by Jack on 2015/07/09 for JC201500200.

Private Const miSpdFieldPos_RiskLevel  As Integer = 30  'Add by Sam on 20190904 for ReqNO:201900414
Private Const miSpdFieldPos_FOSBID  As Integer = 31  'Add by Sam on 20201023 for ReqNO:202000394
Private Const miSpdFieldPos_VirtualLotId    As Integer = 32  'Added by HouYu for BE MES Phase 55 - CP Merge Test專案，新增VirtualLotId欄位 on 2020/12/24

Private Const miSpdFieldPos_FABQT As Integer = 33 'add by Ernest on 2021/06/24 for BE MES Phase 61 - Fab real time process improvement monitor 專案
Private Const miSpdFieldPos_CPPriority As Integer = 34 'add by Ernest on 2021/06/24 for BE MES Phase 61 - Fab real time process improvement monitor 專案

Private Const miSpdFieldPos_CP_Dgrade_Info As Integer = 35 'add by Dexter on 2021/11/19 for BE 工業 3.5 Phase 38 - CP AGV導入專案
Private Const miSpdFieldPos_Dgrade_Detail As Integer = 36 'add by Dexter on 2021/11/19 for BE 工業 3.5 Phase 38 - CP AGV導入專案

Private Const miSpdFieldPos_AYInfo As Integer = 37  'Add by Sam on 20250826 for  #222244

Private Const miSpdFieldPos_MonocotsLot As Integer = 38  'Add by Sam on 20250909 for  #220288

Private miResult    As Integer
Public Property Let LotID(sLotId As String)
    On Error Resume Next
    msLotId = sLotId
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



Private Sub cmdCancel_Click()
    On Error Resume Next
    miResult = vbCancel
    Me.Hide
End Sub

Private Sub cmdPrint_Click()
    Me.spdLotInfo.PrintHeader = "Lot Query by " & moFwOPR.ActiveUser.UserName _
        & " - " & Format(Now, "YYYY/MM/DD HH:MM:SS")
    Me.spdLotInfo.PrintUseDataMax = False
    Me.spdLotInfo.PrintSheet (1)
End Sub

Private Sub cmdQuery_Click()
    Me.cmdQuery.Enabled = False
    Call ExecQuery
End Sub

Private Sub cmdQueryTestFlow_Click()
    Dim mfrmMain As New frmSimpleLotQuery_TestFlow
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
           .RuleName = msRuleName
           .LotID = Me.txtLotID
           .IPN = Me.txtIPN
           .sTep = Me.cboStep
           .Stage = Me.cboStage
           .Location = Me.cboLocation
           .HoldType = Me.cboHoldType
           .Init
        .Show vbModal
        Unload mfrmMain
    End With

End Sub

Private Sub Form_Load()
'**************************************************
'**************************************************
    'Add by Tony Start on 20110218 for Req.JC201100038
    'Marked by Tony on 20120130 for Req.JC201200007
'    Dim i As Integer
'    i = GetKeyState(vbKeyCapital)
'    If i = 0 Then
'        NTToggleKeyCaps
'    End If
    'Add by Tony End on 20110218 for Req.JC201100038
    
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
' Author:       Vencent Wei, CIT 2002/04/14
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

'----
' Condition Checking
'----
    ' <Put your condition checking codes here>...

'----
' Action
'----
    ' <Put your Action codes here>...


    Call InitComboBox
    Me.cmdQuery.Enabled = True
    
'----
' Done
'----
    With Me.spdLotInfo
    'add by Ernest on 20150917 for ReqNo:JC201500268--------start
    .SetText miSpdFieldPos_CarrierType, 0, "SFGLocCarriertype"
     'add by Ernest on 20150917 for ReqNo:JC201500268--------end
    End With
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
' Sub: InitComboBox()
'--------------------------------------------------------------------------------
' Description:  <Type your Sub description here...>
'--------------------------------------------------------------------------------
' Author:       Vencent Wei, CIT 2002/05/28
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
Private Sub InitComboBox()
On Error GoTo ExitHandler:
Dim sProcID As String
Dim typErrInfo As tErrInfo
'----
' Init
'----
    sProcID = "InitComboBox"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog)

'----
' Condition Checking
'----
    ' <Put your condition checking codes here>...
    
    
'----
' Action
'----

    Dim sSQL As String, colSQLResult As Collection
    Dim iIdx As Integer
    
    
    'Stage List
    Me.cboStage.Clear
    Me.cboStage.AddItem ""
    Me.cboStage.AddItem gsSTAGE_AY
    Me.cboStage.AddItem gsSTAGE_FT
    Me.cboStage.AddItem gsSTAGE_WS
    
    'Step
    sSQL = "select distinct A.name, C.DESCRIPTION" _
         & " from fwstep A, fwstep_n2m B, fwstepVersion C " _
         & " Where A.SysId = B.fromid " _
         & " and A.ACTIVEVERSION = C.REVISION " _
         & " and B.toid = C.sysid " _
         & " order by A.name"
         
    Set colSQLResult = moProRawSql.QueryDatabase(sSQL)
    If colSQLResult Is Nothing Then
        Call RaiseError(glERR_RAWSQLERROR, _
                    FormatErrorText(gsETX_RAWSQLERROR, "Select", "FwStep"))
    ElseIf colSQLResult.Count > 0 Then
        Me.cboStep.Clear
        Me.cboStep.AddItem ""
        
        For iIdx = 1 To colSQLResult.Count
            Me.cboStep.AddItem colSQLResult.Item(iIdx).Item(1) & " " & colSQLResult.Item(iIdx).Item(2)
        Next iIdx
    End If
    
    'Added by Jack on 2010/05/05 for JC201000081. <Start>
    'Location.
     sSQL = " select distinct " & gsCAT_TSL_LOCATION _
         & " from " & gsCAT_TBL_STEP_LOCATION _
         & " order by " & gsCAT_TSL_LOCATION
   
    Set colSQLResult = moProRawSql.QueryDatabase(sSQL)
    If colSQLResult Is Nothing Then
        Call RaiseError(glERR_RAWSQLERROR, _
                    FormatErrorText(gsETX_RAWSQLERROR, "Select", "Location"))
    ElseIf colSQLResult.Count > 0 Then
        Me.cboLocation.Clear
        Me.cboLocation.AddItem ""
        
        For iIdx = 1 To colSQLResult.Count
            Me.cboLocation.AddItem colSQLResult.Item(iIdx).Item(1)
        Next iIdx
    End If
    'Added by Jack on 2010/05/05 for JC201000081. <End>
    
    'Hold Type
    Me.cboHoldType.Clear
    Me.cboHoldType.AddItem ""
    Me.cboHoldType.AddItem "Hold"
    Me.cboHoldType.AddItem "Non-Hold"
    
    
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
' Sub: ExecQuery()
'--------------------------------------------------------------------------------
' Description:  <Type your Sub description here...>
'--------------------------------------------------------------------------------
' Author:       Vencent Wei, CIT 2002/05/28
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

'Add by Sam start on 20101011 for ReqNo:JC201000227
Dim sSystime As String

Dim sUnpackTime As String
Dim iMinutes  As Long

Dim lUnPackTat  As Double
Dim vUnPackTime As Variant
Dim vRemainUnpack As Variant
'Add by Sam End on 20101011 for ReqNo:JC201000227

'----
' Init
'----
    sProcID = "ExecQuery"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog)

    sSystime = GetSystemTime(moAppLog, moFwWIP, moFwWF, moCwMbx, "YYYY/MM/DD HH24:MI")

'----
' Condition Checking
'----
    ' <Put your condition checking codes here>...
    'cboLocation added by Jack on 2010/05/05 for JC201000081
    'txtVirtualLotID Added by Jack on 2021/05/18 for CP Merge Test Bumping
    If (cboStep.Text = "" And txtLotID.Text = "" And txtVirtualLotID.Text = "" And _
        txtIPN.Text = "" And cboStage.Text = "" And cboHoldType = "" And cboLocation = "" And _
        txtMonocotsLot = "") Then
            UtShowMsgBox "Pls input Query Criterial !!" & vbNewLine & "請輸入查詢條件 !!"
            GoTo ExitHandler
    End If
    
'----
' Action
'----
    ' <Put your Action codes here>...
    Dim sSQL As String
    Dim oRawSql As New clsRawSql
    Dim oRS As clsRawSqlRecordset
    Dim oRS2 As clsRawSqlRecordset
    Dim msWQtySum As Long
    Dim msCQtySum As Long
    Dim sHoldCode As String
    Dim sHoldReason As String
    
    Dim arrHoldLotIdStep() As String
    Dim sHoldLotId As String
    Dim lRow As Long, iIdx As Integer
    Const ciMaxLot As Integer = 1000
    Dim vLotRow As Variant
    
    Dim sColumn As String
    Dim colRS As Collection

    Set oRawSql.CwMbx = moCwMbx
    Set oRawSql.AppLog = moAppLog
    oRawSql.RawSqlMailBox = moFwWF.ClientID & ".RawSql_Reply"
    
    Screen.MousePointer = vbHourglass

    Me.spdLotInfo.MaxRows = 0
    
    sColumn = "select distinct " & _
                "A.lotid LotID, " & _
                "A.IPN IPN, " & _
                "A.lotstatus1 LotStatus, " & _
                "A.lotowner LotOwner, " & _
                "A.waferqty WaferQty, " & _
                "A.Chipqty ChipQty, " & _
                "A.Stepid StepId, " & _
                "A.stepname StepName, " & _
                "A.Stage Stage, " & _
                "A.hotlotflag HotLotFlag, " & _
                "A.processstatus ProcessStatus, " & _
                "A.timeheresince TimeHereSince, " & _
                "B.Erunticno ERunTicNo, " & _
                "B.Saprwno SapRwNo, " & _
                "B.turnkeyvendor, " & _
                "B.FGIPN, " & _
                "B.vendorname TestSite "
                
    'add by Sam start on 20090310 for ReqNo:JC900022
    sColumn = sColumn & ",B." & gsCAT_TLI_LOCID & " " & _
                  ",B." & gsCAT_TLI_MFGTICNO & " "
    'add by Sam END on 20090310 for ReqNo:JC900022
    
    'add by Sam start on 20101011 for ReqNo:JC201000227
    sColumn = sColumn & ",B." & gsCAT_TLI_UNPACKTIME & " " & _
                " , B." & gsCAT_TLI_REMAIN_UNPACK_TAT & " "
                
    'add by Sam END on 20101011 for ReqNo:JC201000227
            
    'Added by Jack on 2013/02/25 for JC201300058. <Start>
    'Modified by Jack on 2013/03/08 for PAMS Realignment Project
    'GradeLevel --> LotGrade(加註解)
    sColumn = sColumn & ",B.GradeLevel "
    sColumn = sColumn & ",A.DateCode "
    'Added by Jack on 2013/02/25 for JC201300058. <End>
    
    'gsCAT_TSL_LOCATION Added by Jack on 2010/05/05 for JC201000081.
    'turnkeyvendor2 for Location requirement.
    sColumn = sColumn & ",C." & gsCAT_TSL_LOCATION
    'Modify by Sam on 20100709 for ReqNo:JC201000138
    'sColumn = sColumn & ",nvl(decode(nvl(b.turnkeyvendor,' '),'MXIC','SFG',b.turnkeyvendor),'SFG') turnkeyvendor2 "
    'Add WaferSize by Tony on 20110621 for Req.JC201100154
    sColumn = sColumn & ",case when C." & gsCAT_TSL_LOCATION & "='SFG' AND A.stepname IN ('FTBANK','PBBANK','TRBANK') THEN decode(NVL(b.turnkeyvendor,'MXIC'),'MXIC','SFG',b.turnkeyvendor) " & _
                         " WHEN C." & gsCAT_TSL_LOCATION & "='SUB' THEN B.vendorname " & _
                         " ELSE C." & gsCAT_TSL_LOCATION & " END AS turnkeyvendor2 " & _
                         ",D." & gsCAT_TPB_WAFER_SIZE
    'gsCAT_TBL_STEP_LOCATION Added by Jack on 2010/05/05 for JC201000081.
    
    'Add by Sam start on 20140304 for ReqNO:JC201400057
    sColumn = sColumn & ",loc." & gsCAT_TLM_PACKING & ", loc." & gsCAT_TLM_CARRIERTYPE & " "
    'Add by Sam end on 20140304 for ReqNO:JC201400057
    
    'Add by Robin at 20150210 for JC201500043
    sColumn = sColumn & ", substr(E." & gsCAT_TLIT_FAB_OUT & ",1,8) " & gsCAT_TLIT_FAB_OUT & " "
    
    'Added by Jack on 2015/07/09 for JC201500200 <Start>.
    sColumn = sColumn & ",b." & gsCAT_TLI_FAILDGRADE & " "
    'Added by Jack on 2015/07/09 for JC201500200 <End>.
    
    'Add by Sam Start on 20190904 for ReqNo:201900414
    sColumn = sColumn & ",b." & gsCAT_TLI_RISKLEVEL & " "
    'Add by Sam end on 20190904 for ReqNo:201900414
    
    'Add by Sam Start on 20201023 for ReqNo:202000394
    sColumn = sColumn & ",b." & gsCAT_TLI_FOSBID & " "
    'Add by Sam end on 20201023 for ReqNo:202000394
    
    sColumn = sColumn & ",b." & gsCAT_TLI_VIRTUALLOTID & " " 'Added by HouYu for BE MES Phase 55 - CP Merge Test專案，新增VirtualLotId欄位 on 2020/12/24
    
    sColumn = sColumn & ",b." & gsCAT_TLI_FABQT & " " 'add by Ernest on 2021/06/24 for BE MES Phase 61 - Fab real time process improvement monitor 專案
    sColumn = sColumn & ",b." & gsCAT_TLI_CPPRIORITY & " " 'add by Ernest on 2021/06/24 for BE MES Phase 61 - Fab real time process improvement monitor 專案
    
    sColumn = sColumn & ",Fun_Get_DGrade_Info(A.lotid" & ",'0') CP_Dgrade_Info " 'add by Dexter on 2021/11/19 for BE 工業 3.5 Phase 38 - CP AGV導入專案
    
    sColumn = sColumn & ",b." & gsCAT_TLI_AY_INFO & " " 'Add by Sam on 20250826 for  #222244
    
    sColumn = sColumn & ",b." & gsCAT_TLI_MONOCOTSPARENTLOTID & " " 'Add by Sam on 20250909 for  #220288
    
    'Add Tbl_prod_body by Tony on 20110621 for Req.JC201100154
    'Modify start by Robin at 20150210 for JC201500043
    'sSQL = " from tbl_lot_attribute A, " & _
    '            "tbl_lot_info B,  " & _
    '        gsCAT_TBL_STEP_LOCATION & " C ," & _
    '        gsCAT_TBL_PROD_BODY & " D "
    sSQL = " from Tbl_Lot_InOut_Time E, " & _
                "tbl_lot_attribute A,  " & _
                "tbl_lot_info B,  " & _
            gsCAT_TBL_STEP_LOCATION & " C ," & _
            gsCAT_TBL_PROD_BODY & " D "
    'Modify end by Robin at 20150210 for JC201500043
            
    'Add by Sam start on 20140304 for ReqNO:JC201400057
    sSQL = sSQL & "," & gsCAT_TBL_LOC_MASTER & " loc "
    'Add by Sam end on 20140304 for ReqNO:JC201400057
    
    'gsCAT_TSL_STEPID Added by Jack on 2010/05/05 for JC201000081.
    sSQL = sSQL & _
            " where A.lotid=B.lotid(+) and " & _
                " A.lotstatus1 not like 'SHIP' and " & _
                " A.ProcessStatus in ('Active','Hold') " & _
                " and A.Stepid=C." & gsCAT_TSL_STEPID & "(+) "
                
    'Add by Sam start on 20140304 for ReqNO:JC201400057
    sSQL = sSQL & " and b." & gsCAT_TLI_LOCID & " =loc." & gsCAT_TLM_LOCID & "(+) "
    'Add by Sam end on 20140304 for ReqNO:JC201400057
                
                
    'Add by Tony on 20110621 for Req.JC201100154
    sSQL = sSQL & " AND SUBSTR(A." & gsCAT_TLATT_IPN & ",1,4) = D." & gsCAT_TPB_PROD_BODY & "(+)"
                
    'Added by Jack on 2010/05/05 for JC201000081 <Start>.
    'Location
    If Trim(Me.cboLocation) <> "" Then
        Select Case Trim(Me.cboLocation.Text)
               Case "SUB"
                    'Modify by Sam on 20100709 for ReqNo:201000138
'                    sSQL = sSQL & " and C.location = 'SUB' "
'                    sSQL = sSQL & " and nvl(decode(nvl(b.turnkeyvendor,' '),'MXIC','SFG',b.turnkeyvendor),'SFG') <> 'SFG'"
                    sSQL = sSQL & " and (C." & gsCAT_TSL_LOCATION & " = 'SUB' or (C." & gsCAT_TSL_LOCATION & "='SFG' AND A.stepname IN ('FTBANK','PBBANK','TRBANK') AND NVL(B.turnkeyvendor,'MXIC') <> 'MXIC') )"
                    
               Case "SFG"
                    'Modify by Sam on 20100709 for ReqNo:201000138
                    'sSQL = sSQL & " and (C.location = 'SFG' or C.location = 'SUB') "
                    'sSQL = sSQL & " and nvl(decode(nvl(b.turnkeyvendor,' '),'MXIC','SFG',b.turnkeyvendor),'SFG') = 'SFG' "
                     sSQL = sSQL & " and (C." & gsCAT_TSL_LOCATION & " = 'SFG' and (NVL(B.turnkeyvendor,'MXIC') = 'MXIC'  OR A.STEPNAME NOT IN ('FTBANK','PBBANK','TRBANK') )) "
               Case Else
                    sSQL = sSQL & " and C." & gsCAT_TSL_LOCATION & "='" & Me.cboLocation.Text & "' "
        End Select
    End If
    'Added by Jack on 2010/05/05 for JC201000081 <End>.
    
    'Step
    If Me.cboStep <> "" Then
        sSQL = sSQL & " and A.stepid like '" & Left(Me.cboStep, 5) & "' "
    End If

    'LotID
    If Me.txtLotID <> "" Then
        sSQL = sSQL & " and A.lotid like '" & Replace(Me.txtLotID, "*", "%") & "%' "
    End If
    
    'Added by Jack on 2021/05/18 for CP Merge Test Bumping <Start>
    'VirtualLotID==> 不允許萬用字元.
    If Me.txtVirtualLotID <> "" Then
        If InStr(1, Me.txtVirtualLotID, "*") > 0 Then '有輸入'*'
            UtShowMsgBox "虛擬批號，不允許萬用字元 !!"
        Else
            sSQL = sSQL & " and B." & gsCAT_TLI_VIRTUALLOTID & " = '" & Me.txtVirtualLotID & "' "
        End If
               
    End If
    'Added by Jack on 2021/05/18 for CP Merge Test Bumping <End>
    
    'Stage
    If Me.cboStage <> "" Then
        sSQL = sSQL & " and A.stage like '" & Me.cboStage.Text & "' "
    End If
    
    'IPN
    If Me.txtIPN <> "" Then
        If InStr(1, Me.txtIPN, "*") > 0 Then '有輸入'*'
            If Me.txtIPN.Text = "*" Then
                UtShowMsgBox "請輸入詳細的查詢條件 !!"
            Else
                sSQL = sSQL & " and A.IPN like '" & Replace(Me.txtIPN, "*", "%") & "' "
            End If
        Else
            sSQL = sSQL & " and A.IPN like '" & Me.txtIPN & "' "
        End If
    End If
    
    'Hold Type
    If Me.cboHoldType <> "" Then
        If Me.cboHoldType.Text = "Hold" Then
            sSQL = sSQL & " and A.lotstatus1 = 'HOLD'"
        Else
            sSQL = sSQL & " and A.lotstatus1 <> 'HOLD'"
        End If
    End If
    
    'Add by Sam start on 20250909 for  #220288
    If txtMonocotsLot <> "" Then
            sSQL = sSQL & " and b." & gsCAT_TLI_MONOCOTSPARENTLOTID & "  ='" & txtMonocotsLot & "'   "
    End If
    'Add by Sam End on 20250909 for  #220288
    
    'Add by Robin at 20150210 for JC201500043
    sSQL = sSQL & " and e.lotid = a.lotid"
    
    '2009/05/19 Vencent 加查詢筆數限制
    Set colRS = moProRawSql.QueryDatabase("select count(*) " & sSQL)
    If Val(colRS.Item(1).Item(1)) > gsQueryLimitCnt Then
        UtShowRichBox gsQueryLimitMsg
        Me.cmdQuery.Enabled = True
        GoTo ExitHandler
    End If
        
    sSQL = sColumn & sSQL & " order by A.lotid "
    Debug.Print sSQL
    
    Set oRS = oRawSql.SelectRawSqlTxn1("CAT", sSQL)
    msWQtySum = 0
    msCQtySum = 0
    
    ReDim arrHoldLotIdStep(Int(oRS.RecordCount / ciMaxLot) + 1)
    Me.spdHoldLot.MaxRows = 0
    
    With Me.spdLotInfo
    If oRS.RecordCount <= 0 Then
        UtShowMsgBox "無Lot符合條件 !!"
    Else
        Do While oRS.EOF = False
            
            .MaxRows = .MaxRows + 1
            .SetText miSpdFieldPos_LotId, .MaxRows, oRS.Value("LotID")
            .SetText miSpdFieldPos_IPN, .MaxRows, oRS.Value("IPN")
            .SetText miSpdFieldPos_Status, .MaxRows, oRS.Value("LotStatus")
            .SetText miSpdFieldPos_HotFlag, .MaxRows, oRS.Value("HotLotFlag")
            .SetText miSpdFieldPos_LotOwner, .MaxRows, oRS.Value("LotOwner")
            .SetText miSpdFieldPos_WaferQty, .MaxRows, oRS.Value("WaferQty")
             msWQtySum = msWQtySum + Val(oRS.Value("WaferQty"))
            .SetText miSpdFieldPos_ChipQty, .MaxRows, oRS.Value("ChipQty")
            msCQtySum = msCQtySum + Val(oRS.Value("ChipQty"))
            .SetText miSpdFieldPos_Step, .MaxRows, oRS.Value("StepId") & " " & oRS.Value("StepName")
            .SetText miSpdFieldPos_Stage, .MaxRows, oRS.Value("Stage")
            .SetText miSpdFieldPos_ErunTickNo, .MaxRows, oRS.Value("ERunTicNo")
            .SetText miSpdFieldPos_RwSapNo, .MaxRows, oRS.Value("SapRwNo")
            .SetText miSpdFieldPos_TestSite, .MaxRows, oRS.Value("TestSite")
            .SetText miSpdFieldPos_TurnkeyVendor, .MaxRows, oRS.Value("turnkeyvendor")
            .SetText miSpdFieldPos_FGIpn, .MaxRows, oRS.Value("FGIPN")
            
            'Added by Jack on 2010/05/05 for JC201000081. <Start>
            '若Location="SUB", 則以TurnkeyVendor取代之.
            .SetText miSpdFieldPos_Location, .MaxRows, oRS.Value("location")
            If oRS.Value("location") = "SUB" Or oRS.Value("location") = "SFG" Then
                .SetText miSpdFieldPos_Location, .MaxRows, oRS.Value("turnkeyvendor2")
'                If oRS.Value("turnkeyvendor") <> "" And UCase(oRS.Value("turnkeyvendor")) <> "MXIC" Then
'                    .SetText miSpdFieldPos_Location, .MaxRows, oRS.Value("turnkeyvendor")
'                Else
'                    .SetText miSpdFieldPos_Location, .MaxRows, "SFG"
'                End If
            End If
            'Added by Jack on 2010/05/05 for JC201000081. <End>

            'Add by Tony Start on 20110616 for Req.JC201100154.
            .SetText miSpdFieldPos_WaferSize, .MaxRows, oRS.Value("WAFERSIZE")
            'Add by Tony End on 20110616 for Req.JC201100154.

            'Added by Jack on 2013/02/25 for JC201300058. <Start>
            'Modified by Jack on 2013/03/08 for PAMS Realignment Project
            'GradeLevel --> LotGrade(加註解)
            .SetText miSpdFieldPos_DateCode, .MaxRows, oRS.Value("DateCode")
            .SetText miSpdFieldPos_GradeLevel, .MaxRows, oRS.Value("GradeLevel")
            'Added by Jack on 2013/02/25 for JC201300058. <End>

            'Add by Robin at 20150210 for JC201500043
            .SetText miSpdFieldPos_FabOutDate, .MaxRows, oRS.Value(gsCAT_TLIT_FAB_OUT)

'            sHoldReason = ""
'            sHoldCode = ""
            If oRS.Value("LotStatus") = gsLOTSTATUS_HOLD Then
'2009/05/04 Vencent marked, Query Hold Rec 改先紀錄 hold lot id & stepid
'                sSql = "select " & gsCAT_THR_HOLD_CODE & ", " _
'                                 & gsCAT_THR_HOLD_REASON _
'                     & " from " & gsCAT_TBL_HOLD_REC _
'                     & " where " & gsCAT_THR_LOT_ID & " = '" & oRS.Value("LotID") & "' " _
'                     & " and " & gsCAT_THR_STEP_ID & " = '" & oRS.Value("StepId") & "' " _
'                     & " and " & gsCAT_THR_REL_TIME & " is null "
'                Set oRS2 = oRawSql.SelectRawSqlTxn1("CAT", sSql)
'                Do While oRS2.EOF = False
'                    sHoldCode = sHoldCode & oRS2.Value(gsCAT_THR_HOLD_CODE) & ","
'                    sHoldReason = sHoldReason & oRS2.Value(gsCAT_THR_HOLD_REASON) & ","
'                    oRS2.MoveNext
'                Loop
'                If Len(sHoldCode) > 0 Then sHoldCode = Left(sHoldCode, Len(sHoldCode) - 1)
'                If Len(sHoldReason) > 0 Then sHoldReason = Left(sHoldReason, Len(sHoldReason) - 1)
                
                arrHoldLotIdStep(Int(.MaxRows / ciMaxLot)) = arrHoldLotIdStep(Int(.MaxRows / ciMaxLot)) & "('" & oRS.Value("LotID") & "','" & oRS.Value("StepId") & "'),"
                Me.spdHoldLot.MaxRows = Me.spdHoldLot.MaxRows + 1
                Me.spdHoldLot.SetText 1, Me.spdHoldLot.MaxRows, oRS.Value("LotID")
                Me.spdHoldLot.SetText 2, Me.spdHoldLot.MaxRows, .MaxRows
            End If
            
'            .SetText miSpdFieldPos_HoldCode, .MaxRows, sHoldCode
'            .SetText miSpdFieldPos_HoldReason, .MaxRows, sHoldReason
            
            .SetText miSpdFieldPos_AtDays, .MaxRows, CStr(Round(Now - ConvertToDate(oRS.Value("TimeHereSince"), moAppLog), 2))
            
            'add by Sam start on 20090310 for ReqNo:JC900022
            .SetText miSpdFieldPos_LocID, .MaxRows, oRS.Value(gsCAT_TLI_LOCID)
            .SetText miSpdFieldPos_MFGTICNO, .MaxRows, oRS.Value(gsCAT_TLI_MFGTICNO)
            'add by Sam END on 20090310 for ReqNo:JC900022
            
            
            'Add by Sam start on 20140304 for ReqNo:JC201400057
            .SetText miSpdFieldPos_Packing, .MaxRows, oRS.Value(gsCAT_TLM_PACKING)
            .SetText miSpdFieldPos_CarrierType, .MaxRows, oRS.Value(gsCAT_TLM_CARRIERTYPE)
            'Add by Sam end on 20140304 for ReqNo:JC201400057
            
            'Added by Jack on 2015/07/09 for JC201500200 <Start>
            .SetText miSpdFieldPos_FailDgrade, .MaxRows, oRS.Value(gsCAT_TLI_FAILDGRADE)
            'Added by Jack on 2015/07/09 for JC201500200 <End>
            
            .SetText miSpdFieldPos_RiskLevel, .MaxRows, oRS.Value(gsCAT_TLI_RISKLEVEL) 'Add by Sam on 20190904 for ReqNO:201900414
            .SetText miSpdFieldPos_FOSBID, .MaxRows, oRS.Value(gsCAT_TLI_FOSBID) 'Add by Sam on 20201023 for ReqNO:202000394
            
            .SetText miSpdFieldPos_VirtualLotId, .MaxRows, oRS.Value(gsCAT_TLI_VIRTUALLOTID) 'Added by HouYu for BE MES Phase 55 - CP Merge Test專案，新增VirtualLotId欄位 on 2020/12/24
            
            .SetText miSpdFieldPos_FABQT, .MaxRows, oRS.Value(gsCAT_TLI_FABQT) 'add by Ernest on 2021/06/24 for BE MES Phase 61 - Fab real time process improvement monitor 專案
            .SetText miSpdFieldPos_CPPriority, .MaxRows, oRS.Value(gsCAT_TLI_CPPRIORITY) 'add by Ernest on 2021/06/24 for BE MES Phase 61 - Fab real time process improvement monitor 專案
            
            .SetText miSpdFieldPos_CP_Dgrade_Info, .MaxRows, oRS.Value("CP_Dgrade_Info") 'add by Dexter on 2021/11/19 for BE 工業 3.5 Phase 38 - CP AGV導入專案
            .SetText miSpdFieldPos_Dgrade_Detail, .MaxRows, oRS.Value("Detail") 'add by Dexter on 2021/11/19 for BE 工業 3.5 Phase 38 - CP AGV導入專案
            
            'add by Sam start on 20101011 for ReqNo:JC201000227
            'If oRS.Value("location") = "FT" Then 'mark by Ernest on2018/8/21 for BE#201600344/BE#201800415---全開
            .Col = miSpdFieldPos_Remain_UnpackTime
            .Row = .MaxRows
            'Add by Sam start on 20141020 for ReqNo:JC201400349
            If oRS.Value(gsCAT_TLI_REMAIN_UNPACK_TAT) = "NA" Then
                .CellType = CellTypeStaticText
                .SetText miSpdFieldPos_UnpackTime, .MaxRows, oRS.Value(gsCAT_TLI_UNPACKTIME)
                .SetText miSpdFieldPos_Remain_UnpackTime, .MaxRows, oRS.Value(gsCAT_TLI_REMAIN_UNPACK_TAT)
            Else
            'Add by Sam end on 20141020 for ReqNo:JC201400349
                .CellType = CellTypeFloat
                
                .SetText miSpdFieldPos_UnpackTime, .MaxRows, oRS.Value(gsCAT_TLI_UNPACKTIME)
                    
                If oRS.Value(gsCAT_TLI_UNPACKTIME) <> "" Then
                
                    vUnPackTime = oRS.Value(gsCAT_TLI_UNPACKTIME)
                    vRemainUnpack = oRS.Value(gsCAT_TLI_REMAIN_UNPACK_TAT)
                    
                    sUnpackTime = Mid(vUnPackTime, 1, 4) & "/" & Mid(vUnPackTime, 5, 2) & "/" & Mid(vUnPackTime, 7, 2) _
                                  & " " & Mid(vUnPackTime, 10, 2) & ":" & Mid(vUnPackTime, 12, 2)

                    iMinutes = DateDiff("n", sUnpackTime, sSystime)
                    
                    lUnPackTat = Val(vRemainUnpack) - Format((iMinutes - iMinutes Mod 60) / 60 + (iMinutes Mod 60) / 60, "#######0.#")
                            
                    .SetText miSpdFieldPos_Remain_UnpackTime, .MaxRows, lUnPackTat
                Else
                    .SetText miSpdFieldPos_Remain_UnpackTime, .MaxRows, oRS.Value(gsCAT_TLI_REMAIN_UNPACK_TAT)
                End If
            End If
            'End If
            'add by Sam End on 20101011 for ReqNo:JC201000227
            
            'add by Dexter on 2021/11/19 for BE 工業 3.5 Phase 38 - CP AGV導入專案 Start
            .Col = miSpdFieldPos_Dgrade_Detail
            .Row = .MaxRows
            .CellType = CellTypeButton
            .TypeButtonText = "Detail"
            'add by Dexter on 2021/11/19 for BE 工業 3.5 Phase 38 - CP AGV導入專案 End
            
            .SetText miSpdFieldPos_AYInfo, .MaxRows, oRS.Value(gsCAT_TLI_AY_INFO)     'Add by Sam on 20250826 for  #222244
            
            .SetText miSpdFieldPos_MonocotsLot, .MaxRows, oRS.Value(gsCAT_TLI_MONOCOTSPARENTLOTID)       'Add by Sam on 20250909 for #220288

            oRS.MoveNext
        Loop
    End If
    End With
    

    lblTotalLotCount.Caption = oRS.RecordCount
    lblTotalWQty = msWQtySum
    lblTotalCQty = msCQtySum

    '2009/05/04 Vencent Hold rec
    Me.spdLotInfo.ReDraw = False
    For iIdx = 0 To UBound(arrHoldLotIdStep)
        If Len(arrHoldLotIdStep(iIdx)) > 0 Then
            arrHoldLotIdStep(iIdx) = VBA.Left(arrHoldLotIdStep(iIdx), Len(arrHoldLotIdStep(iIdx)) - 1)
            sSQL = "select " & gsCAT_THR_LOT_ID & "," _
                             & gsCAT_THR_HOLD_CODE & ", " _
                             & gsCAT_THR_HOLD_REASON _
                 & " from " & gsCAT_TBL_HOLD_REC _
                 & " where (" & gsCAT_THR_LOT_ID & "," & gsCAT_THR_STEP_ID & ") in (" & arrHoldLotIdStep(iIdx) & ")" _
                 & " and " & gsCAT_THR_REL_TIME & " is null " _
                 & " order by " & gsCAT_THR_LOT_ID ' & "," & gsCAT_THR_CREATE_TIME
            Set oRS2 = oRawSql.SelectRawSqlTxn1("CAT", sSQL)
            Do While oRS2.EOF = False
                If sHoldLotId = oRS2.Value(gsCAT_THR_LOT_ID) Then
                    sHoldCode = sHoldCode & oRS2.Value(gsCAT_THR_HOLD_CODE) & ","
                    sHoldReason = sHoldReason & oRS2.Value(gsCAT_THR_HOLD_REASON) & ","
                Else
                    'Lot Id changed, update spread
                    If sHoldLotId <> "" Then
                        If Len(sHoldCode) > 0 Then sHoldCode = Left(sHoldCode, Len(sHoldCode) - 1)
                        If Len(sHoldReason) > 0 Then sHoldReason = Left(sHoldReason, Len(sHoldReason) - 1)
                        For lRow = 1 To Me.spdHoldLot.MaxRows
                            Me.spdHoldLot.Col = 1
                            Me.spdHoldLot.Row = lRow
                            If Me.spdHoldLot.Text = sHoldLotId Then
                                Me.spdHoldLot.GetText 2, lRow, vLotRow
                                Me.spdLotInfo.SetText miSpdFieldPos_HoldCode, CLng(vLotRow), sHoldCode
                                Me.spdLotInfo.SetText miSpdFieldPos_HoldReason, CLng(vLotRow), sHoldReason
                                Me.spdHoldLot.DeleteRows lRow, 1
                                Me.spdHoldLot.MaxRows = Me.spdHoldLot.MaxRows - 1
                                Exit For
                            End If
                        Next
                    End If
                
                    sHoldLotId = oRS2.Value(gsCAT_THR_LOT_ID)
                    sHoldCode = oRS2.Value(gsCAT_THR_HOLD_CODE) & ","
                    sHoldReason = oRS2.Value(gsCAT_THR_HOLD_REASON) & ","
                End If
                oRS2.MoveNext
            Loop
            
            'update Last Lot
            If sHoldLotId <> "" Then
                If Len(sHoldCode) > 0 Then sHoldCode = Left(sHoldCode, Len(sHoldCode) - 1)
                If Len(sHoldReason) > 0 Then sHoldReason = Left(sHoldReason, Len(sHoldReason) - 1)
                For lRow = 1 To Me.spdHoldLot.MaxRows
                    Me.spdHoldLot.Col = 1
                    Me.spdHoldLot.Row = lRow
                    If Me.spdHoldLot.Text = sHoldLotId Then
                        Me.spdHoldLot.GetText 2, lRow, vLotRow
                        Me.spdLotInfo.SetText miSpdFieldPos_HoldCode, CLng(vLotRow), sHoldCode
                        Me.spdLotInfo.SetText miSpdFieldPos_HoldReason, CLng(vLotRow), sHoldReason
                        Me.spdHoldLot.DeleteRows lRow, 1
                        Me.spdHoldLot.MaxRows = Me.spdHoldLot.MaxRows - 1
                        Exit For
                    End If
                Next
            End If
        End If
    Next iIdx
    
    If Me.spdLotInfo.MaxRows > 0 Then
        'Set up spd
        With Me.spdLotInfo
            .Col = 1
            Do While .Col <= .MaxCols
                .ColWidth(.Col) = .MaxTextColWidth(.Col) + 1
                .Col = .Col + 1
            Loop
            
            'Sort it
            .SetActiveCell miSpdFieldPos_LotId, 1
            .UserColAction = UserColActionSort
            
            'Add by Tony Start on 2013/08/01 for Req.JC201300192
            .Col = miSpdFieldPos_AtDays
            .Row = -1
            .CellType = CellTypeFloat
            
            .Col = miSpdFieldPos_ChipQty
            .Row = -1
            .CellType = CellTypeInteger
            
            .Lock = True
            'Add by Tony End on 2013/08/01 for Req.JC201300192
        
        End With
        
        
        
        
        Me.cmdPrint.Enabled = True
    Else
        Me.cmdPrint.Enabled = False
    End If

    Me.spdLotInfo.ReDraw = True

'----
' Done
'----
    Me.cmdQuery.Enabled = True

ExitHandler:
    ' NOTE 1:
    ' MUST CALL GetErrInfo() here first before another action
    Call GetErrInfo(msMODULE_ID, sProcID, typErrInfo, Erl)
    Call LogProcOut(msMODULE_ID, sProcID, typErrInfo, moAppLog)
    ' <Your cleaning up codes goes here...>
    Screen.MousePointer = vbDefault
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
' Sub: DispLotInfo()
'--------------------------------------------------------------------------------
' Description:  <Type your Sub description here...>
'--------------------------------------------------------------------------------
' Author:       Vencent Wei, CIT 2002/05/28
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
Private Sub DispLotInfo(sLotId As String)
On Error GoTo ExitHandler:
Dim sProcID As String
Dim typErrInfo As tErrInfo
'----
' Init
'----
    sProcID = "DispLotInfo"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog)

'----
' Condition Checking
'----
    ' <Put your condition checking codes here>...
'----
' Action
'----
    ' <Put your Action codes here>...
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

Private Sub txtIPN_KeyUp(KeyCode As Integer, Shift As Integer)

    'Add by Tony Start on 20110218 for Req.JC201100038
    'Marked by Tony on 20120130 for Req.JC201200007
'    Dim i As Integer
'    i = GetKeyState(vbKeyCapital)
'    If i = 0 Then
'        NTToggleKeyCaps
'    End If
    'Add by Tony End on 20110218 for Req.JC201100038

End Sub

Private Sub txtLotID_KeyUp(KeyCode As Integer, Shift As Integer)

    'Add by Tony Start on 20110218 for Req.JC201100038
    'Marked by Tony on 20120130 for Req.JC201200007
'    Dim i As Integer
'    i = GetKeyState(vbKeyCapital)
'    If i = 0 Then
'        NTToggleKeyCaps
'    End If
    'Add by Tony End on 20110218 for Req.JC201100038
    
End Sub

'Dexter on 2021/11/19 for BE 工業 3.5 Phase 38 - CP AGV導入專案
Private Sub spdlotinfo_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    If Col = miSpdFieldPos_Dgrade_Detail Then
        Dim frmDetail As New frmDgDetail
        Dim oLot As FwLot
        Dim sSQL            As String
        Dim sLot As String
        Dim sWaferID As String
        Dim colRS As Collection
        
        Set oLot = FwuRetrieveLot(moFwWIP, Me.txtLotID, moAppLog)
        
        Load frmDetail
        With frmDetail
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
                 .RuleName = msRuleName
                 
           Me.spdLotInfo.Row = Row
           Me.spdLotInfo.Col = Col
           
           Me.spdLotInfo.Col = miSpdFieldPos_LotId
           .LotID = Me.spdLotInfo.Text
           sLot = Me.spdLotInfo.Text
           
           Me.spdLotInfo.Col = miSpdFieldPos_CP_Dgrade_Info
           sWaferID = Me.spdLotInfo.Text
           
        sSQL = "SELECT * " _
        & " FROM " & gsCAT_TBL_CP_DGRADE & " C, " & gsCAT_TBL_LOT_INFO & " LI " _
        & " WHERE LI." & gsCAT_TLI_LOT_ID & " = '" & sLot & "' " _
        & " AND SUBSTR(LI." & gsCAT_TLI_LOT_ID & ", 1, 8) = C." & gsCAT_TCPD_LOTID & " " _
        & " AND INSTR(LI." & gsCAT_TLI_WAFERID & ", C." & gsCAT_TCPD_WAFERID & ") > 0 " _
        & " AND C." & gsCAT_TCPD_DELETEFLAG & " = 'N' "
        
    Set colRS = moProRawSql.QueryDatabase(sSQL)
           
            Me.spdLotInfo.Col = miSpdFieldPos_IPN
           .IPN = Me.spdLotInfo.Text
           '.Ipn = oLot.CustomAttributes(gsLOT_CUSTOMATTR_IPN)
           Me.spdLotInfo.Col = miSpdFieldPos_Step
           .sTep = Me.spdLotInfo.Text
           '.Step = Me.cboStep
           Me.spdLotInfo.Col = miSpdFieldPos_WaferQty
           .WaferQty = Me.spdLotInfo.Text
           
           Me.spdLotInfo.Col = miSpdFieldPos_ChipQty
           .ChipQty = Me.spdLotInfo.Text
           
           
           .Init
           
                .Show vbModal
        End With
        
        Unload frmDetail
    End If
End Sub


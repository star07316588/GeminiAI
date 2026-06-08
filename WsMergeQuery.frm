VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmWsMergeQuery 
   Caption         =   "WsMergeQuery"
   ClientHeight    =   8145
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11130
   LinkTopic       =   "Form1"
   ScaleHeight     =   8145
   ScaleWidth      =   11130
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdResendSAP 
      Caption         =   "Resend SAP"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   32
      Top             =   7680
      Width           =   1695
   End
   Begin TabDlg.SSTab ssTabQuery 
      Height          =   1700
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   11055
      _ExtentX        =   19500
      _ExtentY        =   2990
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "Current"
      TabPicture(0)   =   "frmWsMergeQuery.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraQuery"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "History"
      TabPicture(1)   =   "frmWsMergeQuery.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Merge/Split Detail "
      TabPicture(2)   =   "frmWsMergeQuery.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame2"
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1200
         Left            =   -74880
         TabIndex        =   33
         Top             =   360
         Width           =   10500
         Begin VB.CommandButton cmdSplitQuery 
            Caption         =   "Split Query"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6480
            TabIndex        =   38
            Top             =   720
            Width           =   1695
         End
         Begin VB.TextBox txtLotID_MergeSplit 
            Height          =   405
            Left            =   1440
            TabIndex        =   35
            Top             =   360
            Width           =   2655
         End
         Begin VB.CommandButton cmdMergeQuery 
            Caption         =   "Merge Query"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   8400
            TabIndex        =   34
            Top             =   720
            Width           =   1695
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "(請輸入併批批號)"
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
            TabIndex        =   37
            Top             =   840
            Width           =   1485
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Merge LotID:"
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
            TabIndex        =   36
            Top             =   480
            Width           =   1110
         End
      End
      Begin VB.Frame Frame1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1200
         Left            =   -74880
         TabIndex        =   27
         Top             =   360
         Width           =   10500
         Begin VB.CommandButton cmdQuery_History 
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
            Height          =   375
            Left            =   8400
            TabIndex        =   29
            Top             =   720
            Width           =   1695
         End
         Begin VB.TextBox txtLotID_History 
            Height          =   405
            Left            =   1440
            TabIndex        =   28
            Top             =   360
            Width           =   2655
         End
         Begin VB.Label lblLotID_Histort 
            AutoSize        =   -1  'True
            Caption         =   "Merge LotID:"
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
            TabIndex        =   31
            Top             =   480
            Width           =   1110
         End
         Begin VB.Label lblLotID_Histort_Hint 
            AutoSize        =   -1  'True
            Caption         =   "(請輸入併批批號)"
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
            TabIndex        =   30
            Top             =   840
            Width           =   1485
         End
      End
      Begin VB.Frame fraQuery 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1200
         Left            =   120
         TabIndex        =   21
         Top             =   360
         Width           =   10500
         Begin VB.TextBox txtLotID 
            Height          =   405
            Left            =   1440
            TabIndex        =   24
            Top             =   360
            Width           =   2655
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
            Height          =   375
            Left            =   8400
            TabIndex        =   23
            Top             =   720
            Width           =   1695
         End
         Begin VB.Label lblLotIDHint 
            AutoSize        =   -1  'True
            Caption         =   "(請輸入併批批號)"
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
            TabIndex        =   26
            Top             =   840
            Width           =   1485
         End
         Begin VB.Label lblLotID 
            AutoSize        =   -1  'True
            Caption         =   "Merge LotID:"
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
            TabIndex        =   25
            Top             =   480
            Width           =   1110
         End
      End
      Begin VB.Frame fraReassignQuery 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1305
         Left            =   -74880
         TabIndex        =   8
         Top             =   480
         Width           =   10500
         Begin VB.CommandButton cmdReassignClose 
            Caption         =   "Close"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            Left            =   8715
            TabIndex        =   15
            Top             =   720
            Width           =   1335
         End
         Begin VB.CommandButton cmdReassignQuery 
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
            Height          =   450
            Left            =   8715
            TabIndex        =   14
            Top             =   210
            Width           =   1335
         End
         Begin VB.OptionButton optReassignStep 
            Caption         =   "Step"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   3240
            TabIndex        =   13
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton optReassignIPN 
            Caption         =   "Route/Product"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   270
            TabIndex        =   12
            Top             =   240
            Width           =   1635
         End
         Begin VB.OptionButton optReassignRoute 
            Caption         =   "Route"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   2160
            TabIndex        =   11
            Top             =   240
            Width           =   885
         End
         Begin VB.OptionButton optReassignAll 
            Caption         =   "All"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   4200
            TabIndex        =   10
            Top             =   240
            Value           =   -1  'True
            Width           =   735
         End
         Begin VB.ComboBox cboDeptNO 
            Height          =   315
            Left            =   6255
            Sorted          =   -1  'True
            TabIndex        =   9
            Top             =   240
            Width           =   1800
         End
         Begin MSComCtl2.DTPicker dtpReassignFromDate 
            Height          =   330
            Left            =   1485
            TabIndex        =   16
            Top             =   720
            Width           =   1680
            _ExtentX        =   2963
            _ExtentY        =   582
            _Version        =   393216
            Format          =   69402624
            CurrentDate     =   37410
         End
         Begin MSComCtl2.DTPicker dtpReassignToDate 
            Height          =   330
            Left            =   4365
            TabIndex        =   17
            Top             =   720
            Width           =   1680
            _ExtentX        =   2963
            _ExtentY        =   582
            _Version        =   393216
            Format          =   69402624
            CurrentDate     =   37411
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "From Date"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   360
            TabIndex        =   20
            Top             =   765
            Width           =   990
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "To Date"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   3495
            TabIndex        =   19
            Top             =   765
            Width           =   735
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "DeptNO :"
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
            Left            =   5160
            TabIndex        =   18
            Top             =   285
            Width           =   795
         End
      End
      Begin VB.Frame fraReassignLotList 
         Caption         =   "Lot List"
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
         Height          =   3300
         Left            =   -74880
         TabIndex        =   5
         Top             =   1890
         Width           =   10500
         Begin VB.TextBox txtReassignMessage 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   375
            Left            =   1350
            Locked          =   -1  'True
            TabIndex        =   6
            TabStop         =   0   'False
            Text            =   "No Matching Record Found, 無相符資料"
            Top             =   810
            Visible         =   0   'False
            Width           =   5325
         End
         Begin FPSpreadADO.fpSpread spdReassignLotList 
            Height          =   2895
            Left            =   75
            TabIndex        =   7
            Top             =   270
            Width           =   10305
            _Version        =   196613
            _ExtentX        =   18177
            _ExtentY        =   5106
            _StockProps     =   64
            ColHeaderDisplay=   0
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxCols         =   11
            MaxRows         =   10
            RetainSelBlock  =   0   'False
            ScrollBarExtMode=   -1  'True
            SpreadDesigner  =   "frmWsMergeQuery.frx":0054
         End
      End
      Begin VB.TextBox txtReassignMLotCount 
         BackColor       =   &H8000000A&
         Height          =   285
         Left            =   -66915
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   5280
         Width           =   2500
      End
      Begin VB.Label lblReassignTotalLotCount 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Total Lot Count :"
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
         Left            =   -68400
         TabIndex        =   22
         Top             =   5325
         Width           =   1455
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
      Height          =   375
      Left            =   9360
      TabIndex        =   1
      Top             =   7680
      Width           =   1575
   End
   Begin VB.Frame fraLotInfo 
      Caption         =   "Lot Information"
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
      Height          =   5775
      Left            =   0
      TabIndex        =   0
      Top             =   1800
      Width           =   11055
      Begin FPSpreadADO.fpSpread spdLotInfo 
         Height          =   5295
         Left            =   2040
         TabIndex        =   2
         Top             =   1680
         Width           =   10815
         _Version        =   196613
         _ExtentX        =   19076
         _ExtentY        =   9340
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
         SpreadDesigner  =   "frmWsMergeQuery.frx":03F9
      End
      Begin FPSpreadADO.fpSpread spdSorterControl 
         Height          =   5295
         Left            =   3000
         TabIndex        =   39
         Top             =   240
         Width           =   10815
         _Version        =   196613
         _ExtentX        =   19076
         _ExtentY        =   9340
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
         SpreadDesigner  =   "frmWsMergeQuery.frx":05CD
      End
   End
End
Attribute VB_Name = "frmWsMergeQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const msMODULE_ID As String = "frmWsMergeQuery"
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
Private moRawSql    As clsRawSql

Private msRuleName  As String
Private msLotId     As String
Private msEqpId     As String

'Modified by Jack on 2021/01/05 for CP測前併批專案, Maxcols : 11 --> 20.
Private Const miSpdLotInfo_Maxcols          As Integer = 20

Private Const miSpdLotInfo_ParentLotID      As Integer = 1
Private Const miSpdLotInfo_SLotNO           As Integer = 2 'Added by Jack on 2021/01/05 for CP測前併批專案
Private Const miSpdLotInfo_ChildLotID       As Integer = 3
Private Const miSpdLotInfo_WaferID          As Integer = 4
Private Const miSpdLotInfo_WQty             As Integer = 5
Private Const miSpdLotInfo_CQty             As Integer = 6
Private Const miSpdLotInfo_MergeStepId      As Integer = 7 'Added by Jack on 2021/01/05 for CP測前併批專案
Private Const miSpdLotInfo_UserID           As Integer = 8
Private Const miSpdLotInfo_MergeTime        As Integer = 9
Private Const miSpdLotInfo_IPN              As Integer = 10
Private Const miSpdLotInfo_Owner            As Integer = 11
Private Const miSpdLotInfo_Source           As Integer = 12
Private Const miSpdLotInfo_ErrMsg           As Integer = 13
'Added by Jack on 2021/01/05 for CP測前併批專案 <Start>
Private Const miSpdLotInfo_DeleteTime       As Integer = 14
Private Const miSpdLotInfo_MiddleLotId      As Integer = 15
Private Const miSpdLotInfo_WaferChipQty     As Integer = 16
Private Const miSpdLotInfo_FabCreateTime    As Integer = 17
Private Const miSpdLotInfo_RiskLevel        As Integer = 18
Private Const miSpdLotInfo_RowId            As Integer = 19
Private Const miSpdLotInfo_Table            As Integer = 20
'Added by Jack on 2021/01/05 for CP測前併批專案 <End>

Private Const msSpdLotInfo_ParentLotID      As String = "ParentLotID"
Private Const msSpdLotInfo_ChildLotID       As String = "ChildLotID"
Private Const msSpdLotInfo_WaferID          As String = "WaferID"
Private Const msSpdLotInfo_WQty             As String = "WQty"
Private Const msSpdLotInfo_CQty             As String = "CQty"
Private Const msSpdLotInfo_UserID           As String = "UserID"
Private Const msSpdLotInfo_MergeTime         As String = "MergeTime"
Private Const msSpdLotInfo_IPN             As String = "IPN"
Private Const msSpdLotInfo_Owner            As String = "Owner"
Private Const msSpdLotInfo_Source           As String = "Source"
Private Const msSpdLotInfo_ErrMsg           As String = "ErrMsg"
'Added by Jack on 2021/01/05 for CP測前併批專案 <Start>
Private Const msSpdLotInfo_SLotNO           As String = "SLotNo"
Private Const msSpdLotInfo_MergeStepId      As String = "MergeStepId"
Private Const msSpdLotInfo_DeleteTime       As String = "DeleteTime"
Private Const msSpdLotInfo_MiddleLotId      As String = "MiddleLotId"
Private Const msSpdLotInfo_WaferChipQty     As String = "WaferChipQty"
Private Const msSpdLotInfo_FabCreateTime    As String = "FabCreateTime"
Private Const msSpdLotInfo_RiskLevel        As String = "RiskLevel"
Private Const msSpdLotInfo_RowId            As String = "RowId"
Private Const msSpdLotInfo_Table            As String = "Table"
'Added by Jack on 2021/01/05 for CP測前併批專案 <End>

'Added by Jack on 2021/06/07 for BE MES Phase 59 - Bumping <Start>.
Private Const miSpdSorterControl_Maxcols          As Integer = 11

Private Const miSpdSorterControl_Action          As Integer = 1
Private Const miSpdSorterControl_RefLotId        As Integer = 2
Private Const miSpdSorterControl_RefSLotNo       As Integer = 3
Private Const miSpdSorterControl_NewLotId        As Integer = 4
Private Const miSpdSorterControl_NewSLotNo       As Integer = 5
Private Const miSpdSorterControl_OriginalLotId   As Integer = 6
Private Const miSpdSorterControl_WaferId         As Integer = 7
Private Const miSpdSorterControl_SorterStart     As Integer = 8
Private Const miSpdSorterControl_SorterComplete  As Integer = 9
Private Const miSpdSorterControl_CreateUserId    As Integer = 10
Private Const miSpdSorterControl_CreateTime      As Integer = 11

Private Const msSpdSorterControl_Action          As String = "Action"
Private Const msSpdSorterControl_RefLotId        As String = "RefLotId"
Private Const msSpdSorterControl_RefSLotNo       As String = "RefSLotNo"
Private Const msSpdSorterControl_NewLotId        As String = "NewLotId"
Private Const msSpdSorterControl_NewSLotNo       As String = "NewSLotNo"
Private Const msSpdSorterControl_OriginalLotId   As String = "OriginalLotId"
Private Const msSpdSorterControl_WaferId         As String = "WaferId"
Private Const msSpdSorterControl_SorterStart     As String = "SorterStart"
Private Const msSpdSorterControl_SorterComplete  As String = "SorterComplete"
Private Const msSpdSorterControl_CreateUserId    As String = "CreateUserId"
Private Const msSpdSorterControl_CreateTime      As String = "CreateTime"
'Added by Jack on 2021/06/07 for BE MES Phase 59 - Bumping <End>.

Private miResult    As Integer

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

Private Sub cmdCancel_Click()
    On Error Resume Next
    miResult = vbCancel
    Me.Hide
End Sub
'cmdQuery_Click Added by Jack on 2021/06/07 for BE MES Phase 59 - Bumping <Start>.
Private Sub cmdSplitQuery_Click()
    Call ExecQuery_SorterControl("Split")
End Sub
'cmdMergeQuery_Click( ) Added by Jack on 2021/06/07 for BE MES Phase 59 - Bumping <Start>.
Private Sub cmdMergeQuery_Click()
    Call ExecQuery_SorterControl("Merge")
End Sub
'Added by Jack on 2021/06/07 for BE MES Phase 59 - Bumping <Start>.
Private Sub cmdQuery_Click()
    Call ExecQuery
End Sub
'Added by Jack on 2021/01/05 for CP測前併批專案.
Private Sub cmdQuery_History_Click()
    Call ExecQuery
End Sub
'Added by Jack on 2021/01/05 for CP測前併批專案.
Private Sub cmdResendSAP_Click()
On Error GoTo ExitHandler:
    Dim sProcID             As String
    Dim typErrInfo          As tErrInfo
    
    Dim oUser               As FwUser
    
    Dim vUpdateTable        As Variant
    Dim vUpdateRowID        As Variant
    
    Dim sTable              As String
    Dim sColumnValue        As String
    Dim sWhere              As String
    
    '----
' Init
'----
    sProcID = "ExecQuery"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog)
    
    'Modified by Jack on 2021/03/03 for UAT bug fix. (加 "Set ")
    'oUser = moFwOPR.ActiveUser
    Set oUser = moFwOPR.ActiveUser
    
'----
' Condition Checking
'----
    
    If Trim(txtLotID.Text) = "" Then
            UtShowMsgBox "Pls input Query Criterial !!" & vbNewLine & "請輸入查詢條件 !!"
            GoTo ExitHandler
    End If
    
    If spdLotInfo.MaxRows = 0 Then
        UtShowMsgBox "Lot '" & Me.txtLotID.Text & "' is no data found !!" & vbNewLine & _
                     "Lot '" & Me.txtLotID.Text & "' 查無資料 !!"
        GoTo ExitHandler
    End If

'----
' Action
'----
    Dim sSQL As String
    Dim colRS As Collection
    
    Dim lIndex As Long
    
    Screen.MousePointer = vbHourglass

    With Me.spdLotInfo
        For lIndex = 1 To spdLotInfo.MaxRows
            .GetText miSpdLotInfo_Table, lIndex, vUpdateTable
            .GetText miSpdLotInfo_RowId, lIndex, vUpdateRowID
            
            vUpdateRowID = CStr(vUpdateRowID)
            
            'define table name
            sTable = CStr(vUpdateTable)
        
            'define column value
            sColumnValue = gsCAT_TVMG_SENDSAPFLAG & "='N', " & _
                           gsCAT_TVMG_SENDSAPTIME & "='', " & _
                           gsCAT_TVMG_UPDATEUSERID & "='" & oUser.UserName & "', " & _
                           gsCAT_TVMG_UPDATETIME & "=to_char(sysdate, 'yyyymmddhh24miss') || '00' "
                           
            sWhere = "rowid='" & vUpdateRowID & "' " & _
                      " and nvl(" & gsCAT_TVMG_SENDSAPFLAG & ", 'N') <> 'N'"
                
            'Modify table "Tbl_Lot_Info"
            If UpdateRawSqlTxn(moAppLog, moFwWIP, moFwWF, moCwMbx, "CAT", sTable, sColumnValue, sWhere) <> 0 Then
                Call RaiseError(glERR_FAILTOUPDATE, _
                    FormatErrorText(gsETX_FAILTOUPDATE, "data to database!"))
            End If
                        
        Next
    End With
    
    UtShowMsgBox "Lot '" & Me.txtLotID.Text & "' data is updated !!" & vbNewLine & _
                 "Lot '" & Me.txtLotID.Text & "' 資料已更新, 稍後將重傳給SAP !!"
    
'----
' Done
'----

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
Private Sub Form_Load()
'**************************************************
'**************************************************
    On Error Resume Next
    miResult = vbCancel
    Call ResetFwControls
    
    'Added by Jack on 2021/06/21 for Bumping Project UAT bug fixe <Start>
    With fraLotInfo
        ssTabQuery.Tab = 0
        spdLotInfo.Top = 360
        spdLotInfo.Left = 120
    
        spdSorterControl.Visible = False
        spdLotInfo.Visible = True
    End With
    'Added by Jack on 2021/06/21 for Bumping Project UAT bug fixe <End>
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Call ResetFwControls
End Sub


'================================================================================
' Sub: ExecQuery()
'--------------------------------------------------------------------------------
' Description:  <Type your Sub description here...>
'--------------------------------------------------------------------------------
' Author:       Jeff Weng,  MXIC 2004/07/30
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
'================================================================================
Private Sub ExecQuery()
On Error GoTo ExitHandler:
    Dim sProcID             As String
    Dim typErrInfo          As tErrInfo

    '----
' Init
'----
    sProcID = "ExecQuery"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog)

'----
' Condition Checking
'----
    
     'Modified by Jack on 2021/03/03 for UAT bug fix (兩頁查詢使用不同欄位LotID)
    If ssTabQuery.Tab = 0 Then
        If Trim(txtLotID.Text) = "" Then
                UtShowMsgBox "Pls input Query Criterial !!" & vbNewLine & "請輸入查詢條件 !!"
                GoTo ExitHandler
        End If
    Else
        If Trim(txtLotID_History.Text) = "" Then
                UtShowMsgBox "Pls input Query Criterial !!" & vbNewLine & "請輸入查詢條件 !!"
                GoTo ExitHandler
        End If
    End If
    

'----
' Action
'----
    Dim sSQL As String
    Dim colRS As Collection
    
    Dim lIndex As Long
    
    Screen.MousePointer = vbHourglass

    Me.spdLotInfo.MaxRows = 0
    
    'Added by Jack on 2021/01/05 for CP測前併批專案 <Start>.
    '注意 : 由於共用Spread, 所以查詢 Tbl_Virtual_Lot 或 Tbl_Inkless_Merge_List(_His) 欄位都要同步, 別名要一樣 !!
    'Step-1 : 先檢查是否為 Virtual Lot ?
     'Modified by Jack on 2021/03/03 for UAT bug fix (兩頁查詢使用不同欄位LotID)
    If ssTabQuery.Tab = 0 Then
        sSQL = "select  a.* " & _
                " from " & gsCAT_TBL_VIRTUAL_LOT & " a " & _
               " WHERE a." & gsCAT_TVL_LOTID & " = '" & Me.txtLotID.Text & "' "
    Else
        sSQL = "select  a.* " & _
                " from " & gsCAT_TBL_VIRTUAL_LOT & " a " & _
               " WHERE a." & gsCAT_TVL_LOTID & " = '" & Me.txtLotID_History.Text & "' "
    End If
    
    Set colRS = moProRawSql.QueryDatabase(sSQL)
    If colRS.Count > 0 Then
        sSQL = "select  v." & gsCAT_TVMG_VIRTUALLOTID & " as PARENTLOTID, v." & gsCAT_TVMG_SLOTNO & ", " & _
                      " v." & gsCAT_TVMG_LOTID & " as CHILDLOTID, v." & gsCAT_TVMG_WAFERID & ", " & _
                      " '1' as WQTY, '' as CQTY, '' as SOURCE, '' as ipn, '' as OWNER, '' as MIDDLELOTID, '' as WAFERCHIPQTY, " & _
                      " '' as ERRMSG, '' as FABCREATETIME, '' as RISKLEVEL, " & _
                      " v." & gsCAT_TVMG_MERGESTEPID & " , v." & gsCAT_TVMG_CREATEUSERID & ", " & _
                      " v." & gsCAT_TVMG_CREATETIME & " , v." & gsCAT_TVMG_DELETETIME & ", " & _
                      " v.rowid, '" & gsCAT_TBL_VIRTUAL_MERGE & "' as updatetable "
        sSQL = sSQL & " from " & gsCAT_TBL_VIRTUAL_MERGE & " v "
        
        'Modified by Jack on 2021/03/03 for UAT bug fix (兩頁查詢使用不同欄位LotID)
        If ssTabQuery.Tab = 0 Then
            sSQL = sSQL & " WHERE v." & gsCAT_TVMG_VIRTUALLOTID & " = '" & Me.txtLotID.Text & "' "
        Else
            sSQL = sSQL & " WHERE v." & gsCAT_TVMG_VIRTUALLOTID & " = '" & Me.txtLotID_History.Text & "' "
        End If
        
        'Current 只查 DeleteFlag <> 'Y' 的資料.
        If ssTabQuery.Tab = 0 Then
            sSQL = sSQL & " and nvl(v." & gsCAT_TVMG_DELETEFLAG & ", 'N') <> 'Y'  "
        End If
    Else
        'Modified by Jack on 2021/03/03 for UAT bug fix (兩頁查詢使用不同欄位LotID)
        If ssTabQuery.Tab = 0 Then
            sSQL = "select * from tbl_all_lot a where a.lotid = '" & Me.txtLotID.Text & "' "
        Else
            sSQL = "select * from tbl_all_lot a where a.lotid = '" & Me.txtLotID_History.Text & "' "
        End If
        
        Set colRS = moProRawSql.QueryDatabase(sSQL)
        If Not (colRS.Count > 0) Then
            'Modified by Jack on 2021/03/03 for UAT bug fix (兩頁查詢使用不同欄位LotID)
            If ssTabQuery.Tab = 0 Then
                UtShowMsgBox "Lot '" & Me.txtLotID.Text & "' is not exist in MES !!" & vbNewLine & _
                             "Lot '" & Me.txtLotID.Text & "' 不存在於 MES !!"
            Else
                UtShowMsgBox "Lot '" & Me.txtLotID_History.Text & "' is not exist in MES !!" & vbNewLine & _
                             "Lot '" & Me.txtLotID_History.Text & "' 不存在於 MES !!"
            End If
            GoTo ExitHandler
        End If
        '測後併批要查二個tables : Tbl_Inkless_Merge_List & Tbl_Inkless_Merge_List_His
        '更新SendSapFlag : 由於 Tbl_Inkless_Merge_List_His 不需要重送 SAP, 所以 Livan : 只更新 Tbl_Inkless_Merge_List 即可.
        sSQL = "select  a." & gsCAT_TIML_PARENTLOTID & ", a." & gsCAT_TIML_SLOTNO & ", " & _
                      " a." & gsCAT_TIML_CHILDLOTID & ", a." & gsCAT_TIML_WAFERID & ", " & _
                      " a." & gsCAT_TIML_WQTY & ", a." & gsCAT_TIML_CQTY & ", " & _
                      " a." & gsCAT_TIML_SOURCE & ", a." & gsCAT_TIML_IPN & ", " & _
                      " a." & gsCAT_TIML_OWNER & ", a." & gsCAT_TIML_MIDDLELOTID & ", " & _
                      " a." & gsCAT_TIML_WAFERCHIPQTY & ",a." & gsCAT_TIML_ERRMSG & ", " & _
                      " a." & gsCAT_TIML_FABCREATETIME & ",a." & gsCAT_TIML_RISKLEVEL & ", " & _
                      " '' as mergestepid, " & _
                      " a." & gsCAT_TIML_CREATEUSERID & ", a." & gsCAT_TIML_CREATETIME & ", " & _
                      " a." & gsCAT_TIML_DELETETIME & ", " & _
                      " a.rowid, '" & gsCAT_TBL_INKLESS_MERGE_LIST & "' as updatetable "
        
        sSQL = sSQL & " from " & gsCAT_TBL_INKLESS_MERGE_LIST & " a "
        'Modified by Jack on 2021/03/03 for UAT bug fix (兩頁查詢使用不同欄位LotID)
        If ssTabQuery.Tab = 0 Then
            sSQL = sSQL & " where a." & gsCAT_TIML_PARENTLOTID & " = '" & Me.txtLotID.Text & "' "
        Else
            sSQL = sSQL & " where a." & gsCAT_TIML_PARENTLOTID & " = '" & Me.txtLotID_History.Text & "' "
        End If
        
        'Current 只查 DeleteFlag <> 'Y' 的資料.
        If ssTabQuery.Tab = 0 Then
            sSQL = sSQL & " and nvl(a." & gsCAT_TIML_DELETEFLAG & ", 'N') <> 'Y'"
        End If
        
        sSQL = sSQL & " union "
        
        sSQL = sSQL & _
               "select  a." & gsCAT_TIML_PARENTLOTID & ", a." & gsCAT_TIML_SLOTNO & ", " & _
                      " a." & gsCAT_TIML_CHILDLOTID & ", a." & gsCAT_TIML_WAFERID & ", " & _
                      " a." & gsCAT_TIML_WQTY & ", a." & gsCAT_TIML_CQTY & ", " & _
                      " a." & gsCAT_TIML_SOURCE & ", a." & gsCAT_TIML_IPN & ", " & _
                      " a." & gsCAT_TIML_OWNER & ", a." & gsCAT_TIML_MIDDLELOTID & ", " & _
                      " a." & gsCAT_TIML_WAFERCHIPQTY & ",a." & gsCAT_TIML_ERRMSG & ", " & _
                      " a." & gsCAT_TIML_FABCREATETIME & ",a." & gsCAT_TIML_RISKLEVEL & ", " & _
                      " '' as mergestepid, " & _
                      " a." & gsCAT_TIML_CREATEUSERID & ", a." & gsCAT_TIML_CREATETIME & ", " & _
                      " a." & gsCAT_TIML_DELETETIME & ", " & _
                      " a.rowid, '" & gsCAT_TBL_INKLESS_MERGE_LIST & "' as updatetable "
        
        sSQL = sSQL & " from " & gsCAT_TBL_INKLESS_MERGE_LIST_HIS & " a "
        'Modified by Jack on 2021/03/03 for UAT bug fix (兩頁查詢使用不同欄位LotID)
        If ssTabQuery.Tab = 0 Then
            sSQL = sSQL & " where a." & gsCAT_TIML_PARENTLOTID & " = '" & Me.txtLotID.Text & "' "
        Else
            sSQL = sSQL & " where a." & gsCAT_TIML_PARENTLOTID & " = '" & Me.txtLotID_History.Text & "' "
        End If
        
        'Current 只查 DeleteFlag <> 'Y' 的資料.
        If ssTabQuery.TabIndex = 0 Then
            sSQL = sSQL & " and nvl(a." & gsCAT_TIML_DELETEFLAG & ", 'N') <> 'Y'"
        End If
        
    End If
    'Added by Jack on 2021/01/05 for CP測前併批專案 <End>.

    'Marked by Jack on 2021/01/05 for CP測前併批專案 <Start>.
    'sSQL = "select  a." & gsCAT_TIML_PARENTLOTID & " " & _
    '"  ,a." & gsCAT_TIML_CHILDLOTID & " " & _
    '"  ,a." & gsCAT_TIML_WAFERID & " " & _
    '"  ,a." & gsCAT_TIML_WQTY & " " & _
    '"  ,a." & gsCAT_TIML_CQTY & " " & _
    '"  ,a." & gsCAT_TIML_CREATEUSERID & " " & _
    '"  ,a." & gsCAT_TIML_CREATETIME & " " & _
    '"  ,a." & gsCAT_TIML_IPN & " " & _
    '"  ,a." & gsCAT_TIML_OWNER & " " & _
    '"  ,a." & gsCAT_TIML_SOURCE & " " & _
    '"  ,a." & gsCAT_TIML_ERRMSG & " "
    
    'sSQL = sSQL & " from " & gsCAT_TBL_INKLESS_MERGE_LIST & " a "
        
    'sSQL = sSQL & " WHERE a." & gsCAT_TIML_DELETEFLAG & " = 'N' "

    'If Trim(Me.txtLotID.Text) <> "" Then
    '    sSQL = sSQL & " and a." & gsCAT_TIML_PARENTLOTID & " = '" & Me.txtLotID.Text & "' "
    'End If
    'Marked by Jack on 2021/01/05 for CP測前併批專案 <End>.
    
    Set colRS = moProRawSql.QueryDatabase(sSQL)
    
    'Added by Jack on 2021/01/05 for CP測前併批專案 <Start>.
    If colRS Is Nothing Or colRS.Count = 0 Then
        'Modified by Jack on 2021/03/03 for UAT bug fix (兩頁查詢使用不同欄位LotID)
        If ssTabQuery.Tab = 0 Then
            UtShowMsgBox "Lot '" & Me.txtLotID.Text & "' is no data found !!" & vbNewLine & _
                         "Lot '" & Me.txtLotID.Text & "' 查無資料 !!"
        Else
            UtShowMsgBox "Lot '" & Me.txtLotID_History.Text & "' is no data found !!" & vbNewLine & _
                         "Lot '" & Me.txtLotID_History.Text & "' 查無資料 !!"
        End If
        GoTo ExitHandler
    End If
    'Added by Jack on 2021/01/05 for CP測前併批專案 <End>.
    
    With Me.spdLotInfo
        For lIndex = 1 To colRS.Count
            .MaxRows = .MaxRows + 1
            .SetText miSpdLotInfo_ParentLotID, .MaxRows, colRS.Item(lIndex).Item(gsCAT_TIML_PARENTLOTID)
            .SetText miSpdLotInfo_ChildLotID, .MaxRows, colRS.Item(lIndex).Item(gsCAT_TIML_CHILDLOTID)
            .SetText miSpdLotInfo_WaferID, .MaxRows, colRS.Item(lIndex).Item(gsCAT_TIML_WAFERID)
            .SetText miSpdLotInfo_WQty, .MaxRows, colRS.Item(lIndex).Item(gsCAT_TIML_WQTY)
            .SetText miSpdLotInfo_CQty, .MaxRows, colRS.Item(lIndex).Item(gsCAT_TIML_CQTY)
            .SetText miSpdLotInfo_UserID, .MaxRows, colRS.Item(lIndex).Item(gsCAT_TIML_CREATEUSERID)
            .SetText miSpdLotInfo_MergeTime, .MaxRows, colRS.Item(lIndex).Item(gsCAT_TIML_CREATETIME)
            .SetText miSpdLotInfo_IPN, .MaxRows, colRS.Item(lIndex).Item(gsCAT_TIML_IPN)
            .SetText miSpdLotInfo_Owner, .MaxRows, colRS.Item(lIndex).Item(gsCAT_TIML_OWNER)
            .SetText miSpdLotInfo_Source, .MaxRows, colRS.Item(lIndex).Item(gsCAT_TIML_SOURCE)
            .SetText miSpdLotInfo_ErrMsg, .MaxRows, colRS.Item(lIndex).Item(gsCAT_TIML_ERRMSG)
            
            'Added by Jack on 2021/01/05 for CP測前併批專案 <Start>.
            .SetText miSpdLotInfo_SLotNO, .MaxRows, colRS.Item(lIndex).Item(gsCAT_TIML_SLOTNO)
            .SetText miSpdLotInfo_MergeStepId, .MaxRows, colRS.Item(lIndex).Item(gsCAT_TVMG_MERGESTEPID)
            .SetText miSpdLotInfo_DeleteTime, .MaxRows, colRS.Item(lIndex).Item(gsCAT_TIML_DELETETIME)
            .SetText miSpdLotInfo_MiddleLotId, .MaxRows, colRS.Item(lIndex).Item(gsCAT_TIML_MIDDLELOTID)
            .SetText miSpdLotInfo_WaferChipQty, .MaxRows, colRS.Item(lIndex).Item(gsCAT_TIML_WAFERCHIPQTY)
            .SetText miSpdLotInfo_FabCreateTime, .MaxRows, colRS.Item(lIndex).Item(gsCAT_TIML_FABCREATETIME)
            .SetText miSpdLotInfo_RiskLevel, .MaxRows, colRS.Item(lIndex).Item(gsCAT_TIML_RISKLEVEL)
            .SetText miSpdLotInfo_RowId, .MaxRows, colRS.Item(lIndex).Item("rowid")
            .SetText miSpdLotInfo_Table, .MaxRows, colRS.Item(lIndex).Item("updatetable")
            'Added by Jack on 2021/01/05 for CP測前併批專案 <End>.
        Next
    End With
    'Set up spd
    With Me.spdLotInfo
'        .Col = 1
'        Do While .Col <= .MaxCols
'            .ColWidth(.Col) = .MaxTextColWidth(.Col) + 1
'            .Col = .Col + 1
'        Loop
        
        'Sort it
        .UserColAction = UserColActionSort
        Call .SetCellBorder(0, 0, .MaxCols, .MaxRows, 15, -1, CellBorderStyleSolid)
        
    End With
    
    Call ResizeSpdColumn(Me.spdLotInfo, moAppLog)

'----
' Done
'----

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
' Function: Init()
'--------------------------------------------------------------------------------
' Description:  <Type your function description here...>
'--------------------------------------------------------------------------------
' Author:       Jeff Weng, MXIC 2004/07/30
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
    Call InitSpd
    Call InitComboBox
    
    'Added by Jack on 2021/01/05 for CP測前併批專案 <Start>
    ssTabQuery.Tab = 0
    cmdResendSAP.Visible = True
    cmdResendSAP.Enabled = True
    'Added by Jack on 2021/01/05 for CP測前併批專案 <End>
    
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
        Call HandleError(True, typErrInfo, , moAppLog, True)
    End If
End Function

'================================================================================
' Function: mInitSpread()
'--------------------------------------------------------------------------------
' Description:  <Type your function description here...>
'--------------------------------------------------------------------------------
' Author:       Jeff Weng, MXIC 2004/08/04
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
Public Function InitSpd() As Boolean
On Error GoTo ExitHandler:
Dim sProcID     As String
Dim typErrInfo  As tErrInfo
Dim i           As Integer
Dim iIndex      As Long

Dim oUser As FwUser

'----
' Init
'----
    sProcID = "InitSpd"
    InitSpd = True
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog) '"Entering Function...", oLogCtrl, glLOG_PROC, msMODULE_ID, sProcID)

    'Get Login User
    Set oUser = moFwOPR.ActiveUser

'----
' Condition Checking
'----
    ' <Put your condition checking codes here>...

'----
' Action
'----


    With spdLotInfo
        .MaxRows = 0
        .MaxCols = miSpdLotInfo_Maxcols
        .SetText miSpdLotInfo_ParentLotID, 0, msSpdLotInfo_ParentLotID
        .SetText miSpdLotInfo_ChildLotID, 0, msSpdLotInfo_ChildLotID
        .SetText miSpdLotInfo_WaferID, 0, msSpdLotInfo_WaferID
        .SetText miSpdLotInfo_WQty, 0, msSpdLotInfo_WQty
        .SetText miSpdLotInfo_CQty, 0, msSpdLotInfo_CQty
        .SetText miSpdLotInfo_UserID, 0, msSpdLotInfo_UserID
        .SetText miSpdLotInfo_MergeTime, 0, msSpdLotInfo_MergeTime
        .SetText miSpdLotInfo_IPN, 0, msSpdLotInfo_IPN
        .SetText miSpdLotInfo_Owner, 0, msSpdLotInfo_Owner
        .SetText miSpdLotInfo_Source, 0, msSpdLotInfo_Source
        .SetText miSpdLotInfo_ErrMsg, 0, msSpdLotInfo_ErrMsg
        
        'Added by Jack on 2021/01/05 for CP測前併批專案 <Start>
        .SetText miSpdLotInfo_SLotNO, 0, msSpdLotInfo_SLotNO
        .SetText miSpdLotInfo_MergeStepId, 0, msSpdLotInfo_MergeStepId
        .SetText miSpdLotInfo_DeleteTime, 0, msSpdLotInfo_DeleteTime
        .SetText miSpdLotInfo_MiddleLotId, 0, msSpdLotInfo_MiddleLotId
        .SetText miSpdLotInfo_WaferChipQty, 0, msSpdLotInfo_WaferChipQty
        .SetText miSpdLotInfo_FabCreateTime, 0, msSpdLotInfo_FabCreateTime
        .SetText miSpdLotInfo_RiskLevel, 0, msSpdLotInfo_RiskLevel
        .SetText miSpdLotInfo_RowId, 0, msSpdLotInfo_RowId
        .SetText miSpdLotInfo_Table, 0, msSpdLotInfo_Table
        
        .Col = miSpdLotInfo_RowId
        .Row = -1
        .colHidden = True
        
        .Col = miSpdLotInfo_Table
        .Row = -1
        .colHidden = True
        'Added by Jack on 2021/01/05 for CP測前併批專案 <End>
        
        .Col = miSpdLotInfo_ErrMsg
        .Row = -1
        .TypeMaxEditLen = 256
        
        For i = 1 To .MaxCols
            .ColUserSortIndicator(i) = ColUserSortIndicatorAscending
            .UserColAction = UserColActionSort
        Next
        
        .Row = -1
        For iIndex = 1 To .MaxCols
            Call LockSpdColumn(Me.spdLotInfo, iIndex)
        Next iIndex
        
    End With
    
    'Added by Jack on 2021/06/07 for BE MES Phase 59 - Bumping <Start>.
    
    With spdSorterControl
        .MaxRows = 0
        .MaxCols = miSpdSorterControl_Maxcols
        .SetText miSpdSorterControl_Action, 0, msSpdSorterControl_Action
        .SetText miSpdSorterControl_RefLotId, 0, msSpdSorterControl_RefLotId
        .SetText miSpdSorterControl_RefSLotNo, 0, msSpdSorterControl_RefSLotNo
        .SetText miSpdSorterControl_NewLotId, 0, msSpdSorterControl_NewLotId
        .SetText miSpdSorterControl_NewSLotNo, 0, msSpdSorterControl_NewSLotNo
        .SetText miSpdSorterControl_OriginalLotId, 0, msSpdSorterControl_OriginalLotId
        .SetText miSpdSorterControl_WaferId, 0, msSpdSorterControl_WaferId
        .SetText miSpdSorterControl_SorterStart, 0, msSpdSorterControl_SorterStart
        .SetText miSpdSorterControl_SorterComplete, 0, msSpdSorterControl_SorterComplete
        .SetText miSpdSorterControl_CreateUserId, 0, msSpdSorterControl_CreateUserId
        .SetText miSpdSorterControl_CreateTime, 0, msSpdSorterControl_CreateTime
    
        For i = 1 To .MaxCols
            .ColUserSortIndicator(i) = ColUserSortIndicatorAscending
            .UserColAction = UserColActionSort
        Next
        
        .Row = -1
        For iIndex = 1 To .MaxCols
            Call LockSpdColumn(Me.spdSorterControl, iIndex)
        Next iIndex
    End With
    'Added by Jack on 2021/06/07 for BE MES Phase 59 - Bumping <End>.
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
        InitSpd = False
        On Error GoTo ExitHandler:
        Call HandleError(True, typErrInfo, , moAppLog)
    End If
End Function


'================================================================================
' Sub: InitComboBox()
'--------------------------------------------------------------------------------
' Description:  <Type your Sub description here...>
'--------------------------------------------------------------------------------
' Author:       Jeff Weng, MXIC 2004/08/03
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
    
'----
' Action
'----
  
    
    
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
'Added by Jack on 2021/01/05 for CP測前併批專案.
Private Sub ssTabQuery_Click(PreviousTab As Integer)
    If ssTabQuery.Tab = 0 Then
        cmdResendSAP.Visible = True
        cmdResendSAP.Enabled = True
        
        spdLotInfo.MaxRows = 0
        spdLotInfo.Visible = True
        spdSorterControl.MaxRows = 0
        spdSorterControl.Visible = False
        
        spdLotInfo.Top = 360
        spdLotInfo.Left = 120
    ElseIf ssTabQuery.Tab = 1 Then
        cmdResendSAP.Visible = False
        cmdResendSAP.Enabled = False
        
        spdLotInfo.MaxRows = 0
        spdLotInfo.Visible = True
        spdSorterControl.MaxRows = 0
        spdSorterControl.Visible = False
        
        spdLotInfo.Top = 360
        spdLotInfo.Left = 120
    ElseIf ssTabQuery.Tab = 2 Then
        cmdResendSAP.Visible = False
        cmdResendSAP.Enabled = False
        
        spdLotInfo.MaxRows = 0
        spdLotInfo.Visible = False
        spdSorterControl.MaxRows = 0
        spdSorterControl.Visible = True
        
        spdSorterControl.Top = 360
        spdSorterControl.Left = 120
    End If
End Sub

'================================================================================
' Sub: ExecQuery_SorterControl()
'--------------------------------------------------------------------------------
' Description:  <Type your Sub description here...>
'--------------------------------------------------------------------------------
' Author:       Jack Hsieh,  MXIC 2021/06/08
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
' [REV 01] Jack Hsieh,  MXIC 2021/06/08
' 1) Added by Jack on 2021/06/07 for BE MES Phase 59 - Bumping
'
'================================================================================
Private Sub ExecQuery_SorterControl(ByVal sAction As String)
On Error GoTo ExitHandler:
    Dim sProcID             As String
    Dim typErrInfo          As tErrInfo

    '----
' Init
'----
    sProcID = "ExecQuery_SorterControl"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog)

'----
' Condition Checking
'----
    
    If Trim(txtLotID_MergeSplit.Text) = "" Then
            UtShowMsgBox "Pls input Query Criterial !!" & vbNewLine & "請輸入查詢條件 !!"
            GoTo ExitHandler
    End If
    

'----
' Action
'----
    Dim sSQL As String
    Dim colRS As Collection
    Dim colRS_2 As Collection
    
    Dim lIndex As Long
    
    Screen.MousePointer = vbHourglass

    Me.spdSorterControl.MaxRows = 0
    
    sSQL = "select " & gsCAT_TSTC_ACTION & ", " & gsCAT_TSTC_REFLOTID & ", " & gsCAT_TSTC_REFSLOTNO & ", " & gsCAT_TSTC_NEWLOTID & ", " & _
                 " " & gsCAT_TSTC_NEWSLOTNO & ", " & gsCAT_TSTC_ORIGINALLOTID & ", " & gsCAT_TSTC_WAFERID & ", " & gsCAT_TSTC_SORTERSTART & ", " & _
                 " " & gsCAT_TSTC_SORTERCOMPLETE & " , " & gsCAT_TSTC_CREATEUSERID & ", " & gsCAT_TSTC_CREATETIME & " " & _
            " from " & gsCAT_TBL_SORTER_CONTROL & " s " & _
           " where s." & gsCAT_TSTC_REFLOTID & " = '" & txtLotID_MergeSplit.Text & "' " & _
             " and s." & gsCAT_TSTC_DELETEFLAG & " = 'N' " & _
             " and s." & gsCAT_TSTC_ACTION & " = '" & sAction & "' " & _
             " and s.createtime in (select distinct " & gsCAT_TSTC_CREATETIME & " " & _
                                    " from " & gsCAT_TBL_SORTER_CONTROL & " s " & _
                                   " where s." & gsCAT_TSTC_REFLOTID & " = '" & txtLotID_MergeSplit.Text & "' " & _
                                     " and s." & gsCAT_TSTC_DELETEFLAG & " = 'N' " & _
                                     " and s." & gsCAT_TSTC_SORTERCOMPLETE & " = 'N' " & _
                                     " and s." & gsCAT_TSTC_ACTION & " = '" & sAction & "' ) " & _
           " order by createtime, REFSLOTNO "
    
    Set colRS = moProRawSql.QueryDatabase(sSQL)
    
    If colRS Is Nothing Or colRS.Count = 0 Then
        sSQL = "select * " & _
                " from " & gsCAT_TBL_SORTER_CONTROL & " s " & _
               " where s." & gsCAT_TSTC_REFLOTID & " = '" & txtLotID_MergeSplit.Text & "'" & _
               " and s." & gsCAT_TSTC_ACTION & " = '" & sAction & "' "
        
        Set colRS_2 = moProRawSql.QueryDatabase(sSQL)
        If colRS_2.Count > 0 Then
            UtShowMsgBox "Lot '" & Me.txtLotID_MergeSplit.Text & "' has finish Sorter '" & sAction & "' operation !!" & vbNewLine & _
                         "Lot '" & Me.txtLotID_MergeSplit.Text & "' 已完成 Sorter '" & sAction & "' 作業 !!"
        Else
            UtShowMsgBox "Lot '" & Me.txtLotID_MergeSplit.Text & "' is no '" & sAction & "' data found !!" & vbNewLine & _
                         "Lot '" & Me.txtLotID_MergeSplit.Text & "' 查無 '" & sAction & "' 資料 !!"
        End If
        GoTo ExitHandler
    End If
    
    With Me.spdSorterControl
        For lIndex = 1 To colRS.Count
            .MaxRows = .MaxRows + 1
            .SetText miSpdSorterControl_Action, .MaxRows, colRS.Item(lIndex).Item(gsCAT_TSTC_ACTION)
            .SetText miSpdSorterControl_RefLotId, .MaxRows, colRS.Item(lIndex).Item(gsCAT_TSTC_REFLOTID)
            .SetText miSpdSorterControl_RefSLotNo, .MaxRows, colRS.Item(lIndex).Item(gsCAT_TSTC_REFSLOTNO)
            .SetText miSpdSorterControl_NewLotId, .MaxRows, colRS.Item(lIndex).Item(gsCAT_TSTC_NEWLOTID)
            .SetText miSpdSorterControl_NewSLotNo, .MaxRows, colRS.Item(lIndex).Item(gsCAT_TSTC_NEWSLOTNO)
            .SetText miSpdSorterControl_OriginalLotId, .MaxRows, colRS.Item(lIndex).Item(gsCAT_TSTC_ORIGINALLOTID)
            .SetText miSpdSorterControl_WaferId, .MaxRows, colRS.Item(lIndex).Item(gsCAT_TSTC_WAFERID)
            .SetText miSpdSorterControl_SorterStart, .MaxRows, colRS.Item(lIndex).Item(gsCAT_TSTC_SORTERSTART)
            .SetText miSpdSorterControl_SorterComplete, .MaxRows, colRS.Item(lIndex).Item(gsCAT_TSTC_SORTERCOMPLETE)
            .SetText miSpdSorterControl_CreateUserId, .MaxRows, colRS.Item(lIndex).Item(gsCAT_TSTC_CREATEUSERID)
            .SetText miSpdSorterControl_CreateTime, .MaxRows, colRS.Item(lIndex).Item(gsCAT_TSTC_CREATETIME)
        Next
    End With

    With Me.spdSorterControl
        'Sort it
        .UserColAction = UserColActionSort
        Call .SetCellBorder(0, 0, .MaxCols, .MaxRows, 15, -1, CellBorderStyleSolid)
        Call .Sort(1, 1, .MaxCols, .MaxRows, SortByRow, Array(miSpdSorterControl_CreateTime, miSpdSorterControl_RefSLotNo), Array(1, 1))
    End With
    
    Call ResizeSpdColumn(Me.spdSorterControl, moAppLog)

'----
' Done
'----

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



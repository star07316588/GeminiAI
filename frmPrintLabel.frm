VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form frmPrintLabel 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Print BarCode Label"
   ClientHeight    =   6315
   ClientLeft      =   1560
   ClientTop       =   5955
   ClientWidth     =   6525
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6315
   ScaleWidth      =   6525
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraLabelForm 
      Caption         =   "Label Pack Info. "
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
      Height          =   2895
      Index           =   9
      Left            =   0
      TabIndex        =   139
      Top             =   2520
      Visible         =   0   'False
      Width           =   6495
      Begin VB.TextBox txtLotId 
         Height          =   315
         Index           =   9
         Left            =   1560
         TabIndex        =   149
         Top             =   360
         Width           =   3480
      End
      Begin VB.ComboBox cbo_FT_TR_ReelId 
         Height          =   315
         ItemData        =   "frmPrintLabel.frx":0000
         Left            =   1560
         List            =   "frmPrintLabel.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   148
         Top             =   2280
         Width           =   3480
      End
      Begin VB.TextBox txt_FT_TR_EqId 
         Height          =   315
         Left            =   1560
         TabIndex        =   146
         Top             =   1800
         Width           =   3480
      End
      Begin VB.TextBox txt_FT_TR_QtyReel 
         Height          =   315
         Left            =   1560
         TabIndex        =   144
         Top             =   1320
         Width           =   3480
      End
      Begin VB.TextBox txt_FT_TR_IPN 
         Height          =   315
         Left            =   1560
         TabIndex        =   142
         Top             =   840
         Width           =   3480
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Reel ID:"
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
         Left            =   600
         TabIndex        =   147
         Top             =   2280
         Width           =   705
      End
      Begin VB.Label lbl_FT_TR_EqId 
         AutoSize        =   -1  'True
         Caption         =   "Eq ID:"
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
         Left            =   600
         TabIndex        =   145
         Top             =   1800
         Width           =   540
      End
      Begin VB.Label lbl_FT_TR_QtyReel 
         AutoSize        =   -1  'True
         Caption         =   "Qty/ Reel:"
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
         Left            =   600
         TabIndex        =   143
         Top             =   1320
         Width           =   885
      End
      Begin VB.Label lbl_FT_TR_IPN 
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
         Left            =   600
         TabIndex        =   141
         Top             =   840
         Width           =   375
      End
      Begin VB.Label lbl_FT_TR_LotId 
         AutoSize        =   -1  'True
         Caption         =   "Lot NO:"
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
         Left            =   600
         TabIndex        =   140
         Top             =   360
         Width           =   660
      End
   End
   Begin VB.Frame fraLabelForm 
      Caption         =   "Label Pack Info. "
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
      Height          =   2895
      Index           =   8
      Left            =   120
      TabIndex        =   116
      Top             =   2520
      Visible         =   0   'False
      Width           =   6375
      Begin TabDlg.SSTab tabLabelPackInfo 
         Height          =   2535
         Left            =   120
         TabIndex        =   117
         Top             =   240
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   4471
         _Version        =   393216
         Tabs            =   2
         TabHeight       =   520
         TabCaption(0)   =   "Auto"
         TabPicture(0)   =   "frmPrintLabel.frx":0004
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "lblLPI_LotNo"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "txtLotId(8)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "txtLPI_CarrierType"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).Control(3)=   "txtLPI_BoxingSpecNO"
         Tab(0).Control(3).Enabled=   0   'False
         Tab(0).Control(4)=   "txtLPI_Brand"
         Tab(0).Control(4).Enabled=   0   'False
         Tab(0).Control(5)=   "txtLPI_PinCount"
         Tab(0).Control(5).Enabled=   0   'False
         Tab(0).Control(6)=   "txtLPI_PackageCode"
         Tab(0).Control(6).Enabled=   0   'False
         Tab(0).Control(7)=   "txtLPI_LabelSpecNo"
         Tab(0).Control(7).Enabled=   0   'False
         Tab(0).Control(8)=   "txtLPI_IPN"
         Tab(0).Control(8).Enabled=   0   'False
         Tab(0).ControlCount=   9
         TabCaption(1)   =   "Manual"
         TabPicture(1)   =   "frmPrintLabel.frx":0020
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "cboLPI_PackageCode"
         Tab(1).Control(1)=   "cboLPI_PinCount"
         Tab(1).Control(2)=   "cboLPI_Brand"
         Tab(1).Control(3)=   "cboLPI_BoxingSpecNO"
         Tab(1).Control(4)=   "cboLPI_CarrierType"
         Tab(1).Control(5)=   "cboLPI_LabelSpecNO"
         Tab(1).Control(6)=   "lblLPI_PackageCode"
         Tab(1).Control(7)=   "lblLPI_PinCount"
         Tab(1).Control(8)=   "lblLPI_Brand"
         Tab(1).Control(9)=   "lblLPI_BoxingSpecNO"
         Tab(1).Control(10)=   "lblLPI_CarrierType"
         Tab(1).Control(11)=   "lblLPI_LabelSpecNO"
         Tab(1).ControlCount=   12
         Begin VB.TextBox txtLPI_IPN 
            Height          =   315
            Left            =   1080
            TabIndex        =   138
            Text            =   "txtLPI_IPN"
            Top             =   1080
            Visible         =   0   'False
            Width           =   2355
         End
         Begin VB.TextBox txtLPI_LabelSpecNo 
            Height          =   315
            Left            =   3600
            TabIndex        =   137
            Text            =   "txtLPI_LabelSpecNo"
            Top             =   2160
            Visible         =   0   'False
            Width           =   2355
         End
         Begin VB.TextBox txtLPI_PackageCode 
            Height          =   315
            Left            =   3600
            TabIndex        =   136
            Text            =   "txtLPI_PackageCode"
            Top             =   1800
            Visible         =   0   'False
            Width           =   2355
         End
         Begin VB.TextBox txtLPI_PinCount 
            Height          =   315
            Left            =   3600
            TabIndex        =   135
            Text            =   "txtLPI_PinCount"
            Top             =   1440
            Visible         =   0   'False
            Width           =   2355
         End
         Begin VB.TextBox txtLPI_Brand 
            Height          =   315
            Left            =   3600
            TabIndex        =   134
            Text            =   "txtLPI_Brand"
            Top             =   1080
            Visible         =   0   'False
            Width           =   2355
         End
         Begin VB.TextBox txtLPI_BoxingSpecNO 
            Height          =   315
            Left            =   3600
            TabIndex        =   133
            Text            =   "txtLPI_BoxingSpecNO"
            Top             =   720
            Visible         =   0   'False
            Width           =   2355
         End
         Begin VB.TextBox txtLPI_CarrierType 
            Height          =   315
            Left            =   3600
            TabIndex        =   132
            Text            =   "txtLPI_CarrierType"
            Top             =   360
            Visible         =   0   'False
            Width           =   2355
         End
         Begin VB.ComboBox cboLPI_PackageCode 
            Height          =   315
            ItemData        =   "frmPrintLabel.frx":003C
            Left            =   -73200
            List            =   "frmPrintLabel.frx":0043
            TabIndex        =   130
            Top             =   2160
            Width           =   3480
         End
         Begin VB.ComboBox cboLPI_PinCount 
            Height          =   315
            ItemData        =   "frmPrintLabel.frx":0055
            Left            =   -73200
            List            =   "frmPrintLabel.frx":005C
            TabIndex        =   128
            Top             =   1800
            Width           =   3480
         End
         Begin VB.ComboBox cboLPI_Brand 
            Height          =   315
            ItemData        =   "frmPrintLabel.frx":006E
            Left            =   -73200
            List            =   "frmPrintLabel.frx":0075
            TabIndex        =   126
            Top             =   1440
            Width           =   3480
         End
         Begin VB.ComboBox cboLPI_BoxingSpecNO 
            Height          =   315
            ItemData        =   "frmPrintLabel.frx":0087
            Left            =   -73200
            List            =   "frmPrintLabel.frx":008E
            TabIndex        =   124
            Top             =   1080
            Width           =   3480
         End
         Begin VB.ComboBox cboLPI_CarrierType 
            Height          =   315
            ItemData        =   "frmPrintLabel.frx":00A0
            Left            =   -73200
            List            =   "frmPrintLabel.frx":00A7
            TabIndex        =   122
            Top             =   720
            Width           =   3480
         End
         Begin VB.ComboBox cboLPI_LabelSpecNO 
            Height          =   315
            ItemData        =   "frmPrintLabel.frx":00B9
            Left            =   -73200
            List            =   "frmPrintLabel.frx":00C0
            TabIndex        =   120
            Top             =   360
            Width           =   3480
         End
         Begin VB.TextBox txtLotId 
            Height          =   315
            Index           =   8
            Left            =   1020
            TabIndex        =   118
            Top             =   480
            Width           =   2355
         End
         Begin VB.Label lblLPI_PackageCode 
            AutoSize        =   -1  'True
            Caption         =   "Package Code :"
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
            Left            =   -74760
            TabIndex        =   131
            Top             =   2205
            Width           =   1365
         End
         Begin VB.Label lblLPI_PinCount 
            AutoSize        =   -1  'True
            Caption         =   "Pin Count :"
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
            Left            =   -74760
            TabIndex        =   129
            Top             =   1845
            Width           =   945
         End
         Begin VB.Label lblLPI_Brand 
            AutoSize        =   -1  'True
            Caption         =   "Brand :"
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
            Left            =   -74760
            TabIndex        =   127
            Top             =   1485
            Width           =   615
         End
         Begin VB.Label lblLPI_BoxingSpecNO 
            AutoSize        =   -1  'True
            Caption         =   "Boxinf Spec NO :"
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
            Left            =   -74760
            TabIndex        =   125
            Top             =   1125
            Width           =   1470
         End
         Begin VB.Label lblLPI_CarrierType 
            AutoSize        =   -1  'True
            Caption         =   "Carrier Type :"
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
            Left            =   -74760
            TabIndex        =   123
            Top             =   765
            Width           =   1155
         End
         Begin VB.Label lblLPI_LabelSpecNO 
            AutoSize        =   -1  'True
            Caption         =   "Label Spec NO :"
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
            Left            =   -74760
            TabIndex        =   121
            Top             =   405
            Width           =   1410
         End
         Begin VB.Label lblLPI_LotNo 
            AutoSize        =   -1  'True
            Caption         =   "LOT NO:"
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
            TabIndex        =   119
            Top             =   525
            Width           =   765
         End
      End
   End
   Begin VB.CommandButton cmdPrintFail 
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
      Height          =   315
      Left            =   5160
      TabIndex        =   112
      Top             =   2160
      Width           =   1095
   End
   Begin VB.CommandButton cmdPrintPass 
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
      Height          =   315
      Left            =   5160
      TabIndex        =   111
      Top             =   1710
      Width           =   1095
   End
   Begin VB.TextBox txtPrintFailQty 
      Height          =   315
      Left            =   1560
      MaxLength       =   2
      TabIndex        =   109
      Top             =   2160
      Width           =   3480
   End
   Begin VB.Frame fraNoUse 
      Caption         =   "NoUse"
      Height          =   1575
      Left            =   120
      TabIndex        =   99
      Top             =   5880
      Visible         =   0   'False
      Width           =   5175
      Begin VB.TextBox txtProdNo 
         BackColor       =   &H80000000&
         Height          =   315
         Index           =   6
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   106
         Top             =   1080
         Width           =   2355
      End
      Begin VB.TextBox txtPartNo 
         Height          =   315
         Index           =   6
         Left            =   1080
         TabIndex        =   105
         Top             =   600
         Width           =   2355
      End
      Begin VB.TextBox txtPartNo 
         Height          =   315
         Index           =   5
         Left            =   2280
         TabIndex        =   104
         Top             =   240
         Width           =   2355
      End
      Begin VB.TextBox txtPartNo 
         Height          =   315
         Index           =   4
         Left            =   2040
         TabIndex        =   103
         Top             =   1200
         Width           =   2355
      End
      Begin VB.TextBox txtPartNo 
         Height          =   315
         Index           =   3
         Left            =   2520
         TabIndex        =   102
         Top             =   1080
         Width           =   2355
      End
      Begin VB.TextBox txtPartNo 
         Height          =   315
         Index           =   2
         Left            =   2160
         TabIndex        =   101
         Top             =   720
         Width           =   2355
      End
      Begin VB.TextBox txtPartNo 
         Height          =   315
         Index           =   0
         Left            =   1920
         TabIndex        =   100
         Top             =   360
         Width           =   2355
      End
   End
   Begin VB.ComboBox cboPrintServer 
      Height          =   315
      ItemData        =   "frmPrintLabel.frx":00D2
      Left            =   1575
      List            =   "frmPrintLabel.frx":00D4
      TabIndex        =   67
      Text            =   "cboPrintServer"
      Top             =   1320
      Width           =   3480
   End
   Begin VB.TextBox txtPrintQty 
      Height          =   315
      Left            =   1560
      MaxLength       =   2
      TabIndex        =   65
      Top             =   1710
      Width           =   3480
   End
   Begin VB.ComboBox cboStage 
      Height          =   315
      Left            =   1575
      TabIndex        =   0
      Text            =   "cboStage"
      Top             =   90
      Width           =   3480
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
      Height          =   315
      Left            =   4155
      TabIndex        =   29
      Top             =   5475
      Width           =   1095
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
      Height          =   315
      Left            =   5310
      TabIndex        =   30
      Top             =   5475
      Width           =   1095
   End
   Begin VB.TextBox txtBoxQty 
      Height          =   315
      Left            =   1575
      TabIndex        =   2
      Top             =   900
      Width           =   3480
   End
   Begin VB.ComboBox cboLabelSpec 
      Height          =   315
      ItemData        =   "frmPrintLabel.frx":00D6
      Left            =   1575
      List            =   "frmPrintLabel.frx":00DD
      TabIndex        =   1
      Text            =   "cboLabelSpec"
      Top             =   495
      Width           =   3480
   End
   Begin VB.Frame fraLabelForm 
      Caption         =   "Label Form"
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
      Height          =   2775
      Index           =   4
      Left            =   120
      TabIndex        =   57
      Top             =   2520
      Visible         =   0   'False
      Width           =   6375
      Begin VB.TextBox txtStepName 
         BackColor       =   &H80000000&
         Height          =   315
         Index           =   4
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   113
         Top             =   1320
         Width           =   1800
      End
      Begin VB.TextBox txtRouteId 
         BackColor       =   &H80000000&
         Height          =   315
         Index           =   4
         Left            =   1875
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   2250
         Width           =   4250
      End
      Begin VB.TextBox txtLotId 
         Height          =   315
         Index           =   4
         Left            =   1890
         TabIndex        =   23
         Top             =   270
         Width           =   2355
      End
      Begin VB.TextBox txtProdNo 
         BackColor       =   &H80000000&
         Height          =   315
         Index           =   4
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   675
         Width           =   2355
      End
      Begin VB.TextBox txtCqty 
         BackColor       =   &H80000000&
         Height          =   315
         Index           =   4
         Left            =   1875
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   1485
         Width           =   1335
      End
      Begin VB.TextBox txtLotOwner 
         BackColor       =   &H80000000&
         Height          =   315
         Index           =   4
         Left            =   1875
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   1890
         Width           =   1335
      End
      Begin VB.TextBox txtWqty 
         BackColor       =   &H80000000&
         Height          =   315
         Index           =   4
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   1080
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label lblStepName 
         AutoSize        =   -1  'True
         Caption         =   "STEP:"
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
         Index           =   4
         Left            =   3600
         TabIndex        =   114
         Top             =   1365
         Width           =   540
      End
      Begin VB.Label lblRouteId 
         AutoSize        =   -1  'True
         Caption         =   "ROUTE:"
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
         Index           =   4
         Left            =   1155
         TabIndex        =   63
         Top             =   2295
         Width           =   720
      End
      Begin VB.Label lblLotNo 
         AutoSize        =   -1  'True
         Caption         =   "LOT NO:"
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
         Index           =   4
         Left            =   1110
         TabIndex        =   62
         Top             =   315
         Width           =   765
      End
      Begin VB.Label lblProdNo 
         AutoSize        =   -1  'True
         Caption         =   "PRODUCT NO:"
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
         Index           =   4
         Left            =   570
         TabIndex        =   61
         Top             =   720
         Width           =   1305
      End
      Begin VB.Label lblCqty 
         AutoSize        =   -1  'True
         Caption         =   "CHIP QTY:"
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
         Left            =   930
         TabIndex        =   60
         Top             =   1515
         Width           =   945
      End
      Begin VB.Label lblLotOwner 
         AutoSize        =   -1  'True
         Caption         =   "LOT OWNER:"
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
         Left            =   675
         TabIndex        =   59
         Top             =   1935
         Width           =   1200
      End
      Begin VB.Label lblWqty 
         AutoSize        =   -1  'True
         Caption         =   "WAFER QTY:"
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
         Left            =   705
         TabIndex        =   58
         Top             =   1110
         Visible         =   0   'False
         Width           =   1170
      End
   End
   Begin VB.Frame fraLabelForm 
      Caption         =   "Label Form"
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
      Height          =   2745
      Index           =   7
      Left            =   120
      TabIndex        =   86
      Top             =   2520
      Visible         =   0   'False
      Width           =   6330
      Begin VB.TextBox txtUserId 
         BackColor       =   &H80000000&
         Height          =   315
         Index           =   7
         Left            =   4800
         Locked          =   -1  'True
         TabIndex        =   97
         Top             =   1860
         Width           =   1335
      End
      Begin VB.TextBox txtLotId 
         Height          =   315
         Index           =   7
         Left            =   2130
         TabIndex        =   91
         Top             =   270
         Width           =   2355
      End
      Begin VB.TextBox txtProdNo 
         BackColor       =   &H80000000&
         Height          =   315
         Index           =   7
         Left            =   2130
         Locked          =   -1  'True
         TabIndex        =   90
         Top             =   675
         Width           =   2355
      End
      Begin VB.TextBox txtCqty 
         BackColor       =   &H80000000&
         Height          =   315
         Index           =   7
         Left            =   2115
         Locked          =   -1  'True
         TabIndex        =   89
         Top             =   1905
         Width           =   1335
      End
      Begin VB.TextBox txtPartNo 
         BackColor       =   &H80000000&
         Height          =   315
         Index           =   7
         Left            =   2130
         Locked          =   -1  'True
         TabIndex        =   88
         Top             =   1080
         Width           =   2370
      End
      Begin VB.TextBox txtWqty 
         BackColor       =   &H80000000&
         Height          =   315
         Index           =   7
         Left            =   2130
         Locked          =   -1  'True
         TabIndex        =   87
         Top             =   1500
         Width           =   1335
      End
      Begin VB.Label lblUserId 
         AutoSize        =   -1  'True
         Caption         =   "USERID:"
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
         Index           =   4
         Left            =   3960
         TabIndex        =   98
         Top             =   1890
         Width           =   765
      End
      Begin VB.Label lblLotNo 
         AutoSize        =   -1  'True
         Caption         =   "LOT NO:"
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
         Index           =   7
         Left            =   1350
         TabIndex        =   96
         Top             =   315
         Width           =   765
      End
      Begin VB.Label lblProdNo 
         AutoSize        =   -1  'True
         Caption         =   "PRODUCT NO:"
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
         Index           =   6
         Left            =   840
         TabIndex        =   95
         Top             =   720
         Width           =   1290
      End
      Begin VB.Label lblQty 
         AutoSize        =   -1  'True
         Caption         =   "QUANTITY:"
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
         Index           =   3
         Left            =   1095
         TabIndex        =   94
         Top             =   1935
         Width           =   1020
      End
      Begin VB.Label lblEPN 
         AutoSize        =   -1  'True
         Caption         =   "PART NO:"
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
         Left            =   1245
         TabIndex        =   93
         Top             =   1095
         Width           =   885
      End
      Begin VB.Label lblWqty 
         AutoSize        =   -1  'True
         Caption         =   "WAFER QTY:"
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
         Index           =   5
         Left            =   945
         TabIndex        =   92
         Top             =   1530
         Width           =   1170
      End
   End
   Begin VB.Frame fraLabelForm 
      Caption         =   "Label Form"
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
      Height          =   2745
      Index           =   1
      Left            =   90
      TabIndex        =   39
      Top             =   2520
      Visible         =   0   'False
      Width           =   6330
      Begin VB.TextBox txtWqty 
         BackColor       =   &H80000000&
         Height          =   315
         Index           =   1
         Left            =   2130
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox txtPartNo 
         BackColor       =   &H80000000&
         Height          =   315
         Index           =   1
         Left            =   2115
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   1890
         Width           =   2370
      End
      Begin VB.TextBox txtCqty 
         BackColor       =   &H80000000&
         Height          =   315
         Index           =   1
         Left            =   2115
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   1485
         Width           =   1335
      End
      Begin VB.TextBox txtProdNo 
         BackColor       =   &H80000000&
         Height          =   315
         Index           =   1
         Left            =   2130
         Locked          =   -1  'True
         TabIndex        =   15
         Top             =   675
         Width           =   2355
      End
      Begin VB.TextBox txtLotId 
         Height          =   315
         Index           =   1
         Left            =   2130
         TabIndex        =   14
         Top             =   270
         Width           =   2355
      End
      Begin VB.Label lblWqty 
         AutoSize        =   -1  'True
         Caption         =   "WAFER QTY:"
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
         Index           =   3
         Left            =   945
         TabIndex        =   44
         Top             =   1110
         Width           =   1170
      End
      Begin VB.Label lblPacker 
         AutoSize        =   -1  'True
         Caption         =   "PART NO:"
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
         Left            =   1215
         TabIndex        =   43
         Top             =   1935
         Width           =   900
      End
      Begin VB.Label lblQty 
         AutoSize        =   -1  'True
         Caption         =   "QUANTITY:"
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
         Left            =   1095
         TabIndex        =   42
         Top             =   1515
         Width           =   1020
      End
      Begin VB.Label lblProdNo 
         AutoSize        =   -1  'True
         Caption         =   "PRODUCT NO:"
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
         Left            =   810
         TabIndex        =   41
         Top             =   720
         Width           =   1305
      End
      Begin VB.Label lblLotNo 
         AutoSize        =   -1  'True
         Caption         =   "LOT NO:"
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
         Left            =   1350
         TabIndex        =   40
         Top             =   315
         Width           =   765
      End
   End
   Begin VB.Frame fraLabelForm 
      Caption         =   "Label Form"
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
      Height          =   2745
      Index           =   5
      Left            =   75
      TabIndex        =   69
      Top             =   2520
      Visible         =   0   'False
      Width           =   6330
      Begin VB.TextBox txtUserId 
         BackColor       =   &H80000000&
         Height          =   315
         Index           =   5
         Left            =   2100
         Locked          =   -1  'True
         TabIndex        =   78
         Top             =   1890
         Width           =   2370
      End
      Begin VB.TextBox txtLotId 
         Height          =   315
         Index           =   5
         Left            =   2130
         TabIndex        =   73
         Top             =   270
         Width           =   2355
      End
      Begin VB.TextBox txtProdNo 
         BackColor       =   &H80000000&
         Height          =   315
         Index           =   5
         Left            =   2130
         Locked          =   -1  'True
         TabIndex        =   72
         Top             =   675
         Width           =   2355
      End
      Begin VB.TextBox txtCqty 
         BackColor       =   &H80000000&
         Height          =   315
         Index           =   5
         Left            =   2115
         Locked          =   -1  'True
         TabIndex        =   71
         Top             =   1485
         Width           =   1335
      End
      Begin VB.TextBox txtWqty 
         BackColor       =   &H80000000&
         Height          =   315
         Index           =   5
         Left            =   2130
         Locked          =   -1  'True
         TabIndex        =   70
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label lblUserId 
         AutoSize        =   -1  'True
         Caption         =   "USER ID:"
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
         Left            =   1215
         TabIndex        =   79
         Top             =   1935
         Width           =   840
      End
      Begin VB.Label lblLotNo 
         AutoSize        =   -1  'True
         Caption         =   "LOT NO:"
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
         Index           =   5
         Left            =   1350
         TabIndex        =   77
         Top             =   315
         Width           =   765
      End
      Begin VB.Label lblProdNo 
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
         Index           =   5
         Left            =   1710
         TabIndex        =   76
         Top             =   720
         Width           =   390
      End
      Begin VB.Label lblCqty 
         AutoSize        =   -1  'True
         Caption         =   "CHIP QTY:"
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
         Left            =   1170
         TabIndex        =   75
         Top             =   1515
         Width           =   945
      End
      Begin VB.Label lblWqty 
         AutoSize        =   -1  'True
         Caption         =   "WAFER QTY:"
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
         Index           =   4
         Left            =   945
         TabIndex        =   74
         Top             =   1110
         Width           =   1170
      End
   End
   Begin VB.Frame fraLabelForm 
      Caption         =   "Label Form"
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
      Height          =   2745
      Index           =   0
      Left            =   120
      TabIndex        =   31
      Top             =   2520
      Width           =   6330
      Begin VB.TextBox txtLotId 
         Height          =   315
         Index           =   0
         Left            =   2130
         TabIndex        =   19
         Top             =   450
         Width           =   2355
      End
      Begin VB.TextBox txtProdNo 
         BackColor       =   &H80000000&
         Height          =   315
         Index           =   0
         Left            =   2130
         Locked          =   -1  'True
         TabIndex        =   20
         Top             =   855
         Width           =   2355
      End
      Begin VB.TextBox txtQty 
         BackColor       =   &H80000000&
         Height          =   315
         Index           =   0
         Left            =   2115
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   1305
         Width           =   1335
      End
      Begin VB.TextBox txtUserId 
         BackColor       =   &H80000000&
         Height          =   315
         Index           =   0
         Left            =   2115
         Locked          =   -1  'True
         TabIndex        =   22
         Top             =   1710
         Width           =   1335
      End
      Begin VB.Label lblLotNo 
         AutoSize        =   -1  'True
         Caption         =   "(1T) LOT NO:"
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
         Left            =   945
         TabIndex        =   35
         Top             =   495
         Width           =   1170
      End
      Begin VB.Label lblProdNo 
         AutoSize        =   -1  'True
         Caption         =   "(1P) PRODUCT NO:"
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
         Left            =   405
         TabIndex        =   34
         Top             =   900
         Width           =   1710
      End
      Begin VB.Label lblQty 
         Caption         =   "(Q) QUANTITY:"
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
         Index           =   1
         Left            =   780
         TabIndex        =   33
         Top             =   1335
         Width           =   1335
      End
      Begin VB.Label lblPacker 
         AutoSize        =   -1  'True
         Caption         =   "PACKER:"
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
         Left            =   1305
         TabIndex        =   32
         Top             =   1755
         Width           =   810
      End
   End
   Begin VB.Frame fraLabelForm 
      Caption         =   "Label Form"
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
      Height          =   2715
      Index           =   6
      Left            =   120
      TabIndex        =   80
      Top             =   2640
      Width           =   6210
      Begin VB.TextBox txtCqty 
         BackColor       =   &H80000000&
         Height          =   315
         Index           =   6
         Left            =   4545
         Locked          =   -1  'True
         TabIndex        =   84
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox txtLotId 
         Height          =   315
         Index           =   6
         Left            =   1485
         TabIndex        =   82
         Top             =   240
         Width           =   2100
      End
      Begin FPSpreadADO.fpSpread spdChildLot 
         Height          =   1830
         Left            =   1110
         TabIndex        =   81
         Top             =   735
         Width           =   3810
         _Version        =   196613
         _ExtentX        =   6720
         _ExtentY        =   3228
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
         MaxRows         =   5
         RetainSelBlock  =   0   'False
         ScrollBarExtMode=   -1  'True
         SpreadDesigner  =   "frmPrintLabel.frx":00EF
      End
      Begin VB.Label lblCqty 
         AutoSize        =   -1  'True
         Caption         =   "QTY:"
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
         Left            =   3915
         TabIndex        =   85
         Top             =   300
         Width           =   450
      End
      Begin VB.Label lblLotNo 
         AutoSize        =   -1  'True
         Caption         =   "(1T) LOT NO:"
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
         Index           =   6
         Left            =   165
         TabIndex        =   83
         Top             =   300
         Width           =   1170
      End
   End
   Begin VB.Frame fraLabelForm 
      Caption         =   "Label Form"
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
      Height          =   2745
      Index           =   3
      Left            =   90
      TabIndex        =   51
      Top             =   2520
      Visible         =   0   'False
      Width           =   6330
      Begin VB.TextBox txtSpeed 
         BackColor       =   &H80000000&
         Height          =   315
         Index           =   3
         Left            =   4320
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   1890
         Width           =   1335
      End
      Begin VB.TextBox txtWqty 
         BackColor       =   &H80000000&
         Height          =   315
         Index           =   3
         Left            =   2130
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox txtLotOwner 
         BackColor       =   &H80000000&
         Height          =   315
         Index           =   3
         Left            =   2115
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   1890
         Width           =   1335
      End
      Begin VB.TextBox txtCqty 
         BackColor       =   &H80000000&
         Height          =   315
         Index           =   3
         Left            =   2115
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1485
         Width           =   1335
      End
      Begin VB.TextBox txtProdNo 
         BackColor       =   &H80000000&
         Height          =   315
         Index           =   3
         Left            =   2130
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   675
         Width           =   2355
      End
      Begin VB.TextBox txtLotId 
         Height          =   315
         Index           =   3
         Left            =   2130
         TabIndex        =   3
         Top             =   270
         Width           =   2355
      End
      Begin VB.Label lblSpeed 
         AutoSize        =   -1  'True
         Caption         =   "SPEED:"
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
         Index           =   3
         Left            =   3600
         TabIndex        =   64
         Top             =   1935
         Width           =   690
      End
      Begin VB.Label lblWqty 
         AutoSize        =   -1  'True
         Caption         =   "WAFER QTY:"
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
         Left            =   945
         TabIndex        =   56
         Top             =   1110
         Width           =   1170
      End
      Begin VB.Label lblLotOwner 
         AutoSize        =   -1  'True
         Caption         =   "LOT OWNER:"
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
         Left            =   915
         TabIndex        =   55
         Top             =   1935
         Width           =   1200
      End
      Begin VB.Label lblCqty 
         AutoSize        =   -1  'True
         Caption         =   "CHIP QTY:"
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
         Index           =   4
         Left            =   1170
         TabIndex        =   54
         Top             =   1515
         Width           =   945
      End
      Begin VB.Label lblProdNo 
         AutoSize        =   -1  'True
         Caption         =   "PRODUCT NO:"
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
         Index           =   3
         Left            =   810
         TabIndex        =   53
         Top             =   720
         Width           =   1305
      End
      Begin VB.Label lblLotNo 
         AutoSize        =   -1  'True
         Caption         =   "LOT NO:"
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
         Index           =   3
         Left            =   1350
         TabIndex        =   52
         Top             =   315
         Width           =   765
      End
   End
   Begin VB.Frame fraLabelForm 
      Caption         =   "Label Form"
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
      Height          =   2745
      Index           =   2
      Left            =   90
      TabIndex        =   45
      Top             =   2520
      Visible         =   0   'False
      Width           =   6330
      Begin VB.TextBox txtLotId 
         Height          =   315
         Index           =   2
         Left            =   2130
         TabIndex        =   9
         Top             =   270
         Width           =   2355
      End
      Begin VB.TextBox txtProdNo 
         BackColor       =   &H80000000&
         Height          =   315
         Index           =   2
         Left            =   2130
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   675
         Width           =   2355
      End
      Begin VB.TextBox txtCqty 
         BackColor       =   &H80000000&
         Height          =   315
         Index           =   2
         Left            =   2115
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1485
         Width           =   1335
      End
      Begin VB.TextBox txtUserId 
         BackColor       =   &H80000000&
         Height          =   315
         Index           =   2
         Left            =   2115
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   1890
         Width           =   2370
      End
      Begin VB.TextBox txtWqty 
         BackColor       =   &H80000000&
         Height          =   315
         Index           =   2
         Left            =   2130
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label lblLotNo 
         AutoSize        =   -1  'True
         Caption         =   "LOT NO:"
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
         Left            =   1350
         TabIndex        =   50
         Top             =   315
         Width           =   765
      End
      Begin VB.Label lblProdNo 
         AutoSize        =   -1  'True
         Caption         =   "PRODUCT NO:"
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
         Left            =   810
         TabIndex        =   49
         Top             =   720
         Width           =   1305
      End
      Begin VB.Label lblCqty 
         AutoSize        =   -1  'True
         Caption         =   "QUANTITY:"
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
         Index           =   3
         Left            =   1095
         TabIndex        =   48
         Top             =   1515
         Width           =   1020
      End
      Begin VB.Label lblUserId 
         AutoSize        =   -1  'True
         Caption         =   "USER ID:"
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
         Left            =   1275
         TabIndex        =   47
         Top             =   1935
         Width           =   840
      End
      Begin VB.Label lblWqty 
         AutoSize        =   -1  'True
         Caption         =   "WAFER QTY:"
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
         Left            =   945
         TabIndex        =   46
         Top             =   1110
         Width           =   1170
      End
   End
   Begin VB.Frame famLotList 
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
      Height          =   2760
      Left            =   120
      TabIndex        =   107
      Top             =   2520
      Visible         =   0   'False
      Width           =   6400
      Begin FPSpreadADO.fpSpread spdLotList 
         Height          =   2250
         Left            =   240
         TabIndex        =   108
         Top             =   360
         Width           =   6000
         _Version        =   196613
         _ExtentX        =   10583
         _ExtentY        =   3969
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
         MaxCols         =   7
         MaxRows         =   1
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "frmPrintLabel.frx":03B4
      End
      Begin FPSpreadADO.fpSpread spdWsEngLoc 
         Height          =   1890
         Left            =   240
         TabIndex        =   115
         Top             =   360
         Width           =   6000
         _Version        =   196613
         _ExtentX        =   10583
         _ExtentY        =   3334
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
         MaxCols         =   5
         MaxRows         =   1
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "frmPrintLabel.frx":0759
      End
   End
   Begin VB.Label lblPrintFailQty 
      AutoSize        =   -1  'True
      Caption         =   "Fail Sheet:"
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
      TabIndex        =   110
      Top             =   2190
      Width           =   915
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Printer Server:"
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
      Left            =   270
      TabIndex        =   68
      Top             =   1350
      Width           =   1245
   End
   Begin VB.Label lblPrintQty 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "列印份數:"
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
      Index           =   3
      Left            =   240
      TabIndex        =   66
      Top             =   1770
      Width           =   1155
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
      Left            =   945
      TabIndex        =   38
      Top             =   135
      Width           =   570
   End
   Begin VB.Label lblLabelForm 
      AutoSize        =   -1  'True
      Caption         =   "Label Type:"
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
      Left            =   495
      TabIndex        =   37
      Top             =   540
      Width           =   1020
   End
   Begin VB.Label lblQty 
      AutoSize        =   -1  'True
      Caption         =   "BoxQty:"
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
      Left            =   840
      TabIndex        =   36
      Top             =   945
      Width           =   675
   End
End
Attribute VB_Name = "frmPrintLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const msMODULE_ID As String = "frmPrintLabel"
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

Private msLotId     As String
Private msEqpId     As String
Private msRuleName  As String
Private miResult    As Integer

Private miIndex     As Integer

'**************************************************
'Comment by Sam on 20080910
'這四個變數僅供其他rule call form時使用
Private msProdNo       As String
Private msPartNo       As String
Private msLabelSpec    As String
Private msStage        As String
'**************************************************

'Added by Jack on 2010/04/12 for JC201000060. <Start>
Private Const miSPDLOTLIST_MAX_COLS    As Integer = 9 'Modified by Jack on 2011/04/15 8 --> 9
Private Const mispdLotList_CHECK       As Integer = 1
Private Const miSPDLOTLIST_LOT_ID      As Integer = 2
Private Const miSPDLOTLIST_PROD_NO     As Integer = 3
Private Const mispdLotList_WQTY        As Integer = 4
Private Const mispdLotList_CQTY        As Integer = 5
Private Const miSPDLOTLIST_LOT_OWNER   As Integer = 6
Private Const miSPDLOTLIST_SPEED       As Integer = 7
Private Const mispdLotList_ROUTE       As Integer = 8
Private Const mispdLotList_STEP        As Integer = 9 'Added by Jack on 2011/04/15
'Added by Jack on 2010/04/12 for JC201000060. <End>

'Add by Sam start on 20110823 for ReqNo:JC201100218
Private Const miSPDWSENGLOC_MAX_COLS    As Integer = 5
Private Const miSPDWSENGLOC_CHECK       As Integer = 1
Private Const miSPDWSENGLOC_LOT_ID      As Integer = 2
Private Const miSPDWSENGLOC_WQTY        As Integer = 3
Private Const miSPDWSENGLOC_OWNERID     As Integer = 4
Private Const miSPDWSENGLOC_OWNERDEP    As Integer = 5

'Add by Sam End on 20110823 for ReqNo:JC201100218

Public Property Let LotID(sLotId As String)
    msLotId = sLotId
End Property
Public Property Let EqpID(sEqpID As String)
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
Public Property Let ProdNo(sProdNo As String)
    On Error Resume Next
    msProdNo = sProdNo
End Property
Public Property Let PartNo(sPartNo As String)
    On Error Resume Next
    msPartNo = sPartNo
End Property
Public Property Let LabelSpec(sLabelSpec As String)
    On Error Resume Next
    msLabelSpec = sLabelSpec
End Property
Public Property Let Stage(sStage As String)
    On Error Resume Next
    msStage = sStage
End Property

Private Sub ResetForm()
    cmdOk.Enabled = False
    cmdCancel.Enabled = True
End Sub
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

Private Sub cboLabelSpec_Click()
    Dim iIndex As Integer
    Dim bLabelData  As Boolean
    Dim oLot        As FwLot 'Added by Jack on 2010/04/12 for JC201000060.
    For iIndex = Me.fraLabelForm.LBound To Me.fraLabelForm.UBound
        Me.fraLabelForm.Item(iIndex).Visible = False
    Next iIndex
    
    Me.spdWsEngLoc.Visible = False  'Add by Sam on 20110823 for ReqNo:JC201100218
    Me.spdLotList.Visible = True      'Add by Sam on 20110823 for ReqNo:JC201100218
    
    'Added by Jack on  2011/04/15 <Start>
    Me.cmdOk.Visible = True
    Me.cmdPrintPass.Enabled = False
    Me.cmdPrintFail.Enabled = False
    Me.txtPrintFailQty.Enabled = False
    With spdLotList
        .Col = mispdLotList_STEP
        .colHidden = True
    End With
    
    lblQty(0).Visible = True
    lblQty(1).Visible = True
    lblQty(2).Visible = True
    lblQty(3).Visible = True
    txtBoxQty.Visible = True
    lblPrintQty(3).Caption = "列印份數:"  'for "Pass Sheet".
    lblPrintFailQty.Visible = False
    txtPrintFailQty.Visible = False
    cmdPrintPass.Visible = False
    cmdPrintFail.Visible = False
    'Added by Jack on  2011/04/15 <Start>
    
    'Added by Jack on 2010/04/12 for JC201000060. <Start>
    famLotList.Visible = False
    With Me.spdLotList
        .MaxRows = 0
        .MaxRows = 1
        '將user於Dispatch Lot點選的Lot帶至Spread中.
        If Len(msLotId) <> 0 Then
            Set oLot = moFwWIP.LotById(msLotId)
            .SetText miSPDLOTLIST_LOT_ID, 1, Trim(msLotId)
            Call spdLotList_Change(miSPDLOTLIST_LOT_ID, 1) '以LotId將其他資料帶至Spread中.
            .SetText mispdLotList_CHECK, 1, 1 'LotId 存在, 打勾.
        End If
        Call .SetCellBorder(0, 0, .MaxCols, .MaxRows, 15, -1, CellBorderStyleSolid)
        Call ResizeSpdColumn(Me.spdLotList, moAppLog)
    End With
    'Added by Jack on 2010/04/12 for JC201000060. <End>
    
    Call RefreshPrinterServer(Me.cboPrintServer, Me.cboStage.Text, Me.cboLabelSpec.Text, moAppLog, moProRawSql) 'Add by Sam on 20121203 for ReqNo:JC201200343
    
'    bLabelData = GetLabelData 'add by Sam on 20080910 for ReqNO:M200807057 ,Mark by Sam on 20121203 for ReqNo:JC201200343
    
    If (Me.cboStage.Text = gsSTAGE_FT And Me.cboLabelSpec.Text = modPrint.gsLABEL_TO_SUBTR_NORMAL) Or _
        (Me.cboStage.Text = gsSTAGE_FT And Me.cboLabelSpec.Text = modPrint.gsLABEL_TO_SUBMK_NORMAL) Or _
        (Me.cboStage.Text = gsSTAGE_FT And Me.cboLabelSpec.Text = modPrint.gsLABEL_TO_SUBFT_NORMAL) Then
        Me.fraLabelForm.Item(0).Visible = True
        Me.txtBoxQty.Enabled = True: Me.txtBoxQty.BackColor = vbWhite
        miIndex = 0
    ElseIf (Me.cboStage.Text = gsSTAGE_FT And Me.cboLabelSpec.Text = modPrint.gsLABEL_TO_SUBTR_ENG_SAMPLE) Or _
           (Me.cboStage.Text = gsSTAGE_FT And Me.cboLabelSpec.Text = modPrint.gsLABEL_TO_SUBMK_ENG_SAMPLE) Or _
           (Me.cboStage.Text = gsSTAGE_FT And Me.cboLabelSpec.Text = modPrint.gsLABEL_TO_SUBFT_ENG_SAMPLE) Then
        Me.fraLabelForm.Item(0).Visible = True
        Me.txtBoxQty.Enabled = True: Me.txtBoxQty.BackColor = vbWhite
        miIndex = 0
    ElseIf (Me.cboStage.Text = gsSTAGE_WS And Me.cboLabelSpec.Text = modPrint.gsLABEL_WS_TO_SFG) Then
        Me.fraLabelForm.Item(1).Visible = True
        Me.txtBoxQty.Text = "" 'add by Sam on 20080910 for ReqNO:M200807057
        Me.txtBoxQty.Enabled = False: Me.txtBoxQty.BackColor = &H8000000F
        miIndex = 1
    ElseIf (Me.cboStage.Text = gsSTAGE_WS And Me.cboLabelSpec.Text = modPrint.gsLABEL_WS_SUMMARY) Then
        Me.fraLabelForm.Item(2).Visible = True
        Me.txtBoxQty.Text = "" 'add by Sam on 20080910 for ReqNO:M200807057
        Me.txtBoxQty.Enabled = False: Me.txtBoxQty.BackColor = &H8000000F
        miIndex = 2
        
    ElseIf (Me.cboStage.Text = gsSTAGE_WS And Me.cboLabelSpec.Text = modPrint.gsLABEL_WS_DGRADE_SUMMARY) Then
        Me.fraLabelForm.Item(2).Visible = True
        Me.txtBoxQty.Text = "" 'add by Sam on 20080910 for ReqNO:M200807057
        Me.txtBoxQty.Enabled = False: Me.txtBoxQty.BackColor = &H8000000F
        miIndex = 2
    'Modify by Sam on 20180814 for BE工業 3.5 BE MES Phase 24 -- CP Auto Dispatch, Add CP_SMALL_LABL
    'Modify by Sam on 20201203 for Project CP 測前併批,新增CP Virtual兩種label
    'Modify by Ernest on 20210315 for Project CP 測前併批 ,新增cp virtual label 新增AY
'    ElseIf (Me.cboStage.Text = gsSTAGE_WS And Me.cboLabelSpec.Text = modPrint.gsLABEL_WS_SMALL_LABEL) Or _
'           (Me.cboStage.Text = gsSTAGE_WS And Me.cboLabelSpec.Text = modPrint.gsLABEL_CP_SMALL_LABEL) Then
    ElseIf (Me.cboStage.Text = gsSTAGE_WS And Me.cboLabelSpec.Text = modPrint.gsLABEL_WS_SMALL_LABEL) Or _
           (Me.cboStage.Text = gsSTAGE_WS And Me.cboLabelSpec.Text = modPrint.gsLABEL_CP_VIRTUAL_LOT_LABEL) Or _
           (Me.cboStage.Text = gsSTAGE_WS And Me.cboLabelSpec.Text = modPrint.gsLABEL_CP_VIRTUAL_MERGE) Or _
           (Me.cboStage.Text = gsSTAGE_WS And Me.cboLabelSpec.Text = modPrint.gsLABEL_CP_SMALL_LABEL) Or _
           (Me.cboStage.Text = gsSTAGE_AY And Me.cboLabelSpec.Text = modPrint.gsLABEL_CP_VIRTUAL_LOT_LABEL) Or _
           (Me.cboStage.Text = gsSTAGE_AY And Me.cboLabelSpec.Text = modPrint.gsLABEL_CP_VIRTUAL_MERGE) Then
        'msLabelSpec added by Jack on 2010/04/12 for JC201000060.
        '有值 : 由frmAdltProd執行此程式, 只能單筆列印, 只可列印RENESAS_SHIPPING Label.
        '無值: 由F1執行此程式, "WS_SMALL_LABEL"/"FT_SMALL_LABEL"都能多筆列印.
        If msLabelSpec <> "" Then
            Me.fraLabelForm.Item(3).Visible = True
        Else
            famLotList.Visible = True
            With spdLotList
                .Col = mispdLotList_WQTY
                .colHidden = False
                .Col = miSPDLOTLIST_SPEED
                .colHidden = False
                .Col = mispdLotList_ROUTE
                .colHidden = True
                'Add by SAm start on 20201203 for Project CP 測前併批
                If Me.cboLabelSpec.Text = modPrint.gsLABEL_CP_VIRTUAL_LOT_LABEL Or _
                    Me.cboLabelSpec.Text = modPrint.gsLABEL_CP_VIRTUAL_MERGE Then
                    .Col = miSPDLOTLIST_LOT_OWNER
                    .colHidden = True
                    .Col = miSPDLOTLIST_SPEED
                    .colHidden = True
                End If
                'Add by SAm End on 20201203 for Project CP 測前併批
            End With
        End If
        Me.txtBoxQty.Text = "" 'add by Sam on 20080910 for ReqNO:M200807057
        Me.txtBoxQty.Enabled = False: Me.txtBoxQty.BackColor = &H8000000F
        miIndex = 3
    'Add by Sam start on 20110823 for ReqNo:JC201100218
    ElseIf (Me.cboStage.Text = gsSTAGE_WS And Me.cboLabelSpec.Text = modPrint.gsLABEL_WS_ENG_LOC_LABEL) Then
            famLotList.Visible = True
            spdLotList.Visible = False
            spdWsEngLoc.Visible = True
             With spdWsEngLoc
                .Col = miSPDWSENGLOC_CHECK
                .colHidden = True
            End With
            Me.txtBoxQty.Text = ""
            Me.txtBoxQty.Enabled = False: Me.txtBoxQty.BackColor = &H8000000F
            miIndex = 3
    'Add by Sam start on 20110823 for ReqNo:JC201100218
    ElseIf (Me.cboStage.Text = gsSTAGE_AY And Me.cboLabelSpec.Text = modPrint.gsLABEL_WS_SMALL_LABEL) Then
        'msLabelSpec added by Jack on 2010/04/12 for JC201000060.
        '有值 : 由frmAdltProd執行此程式, 只能單筆列印, 只可列印RENESAS_SHIPPING Label.
        '無值: 由F1執行此程式, "WS_SMALL_LABEL"/"FT_SMALL_LABEL"都能多筆列印.
        If msLabelSpec <> "" Then
            Me.fraLabelForm.Item(3).Visible = True
        Else
            famLotList.Visible = True
            With spdLotList
                .Col = mispdLotList_WQTY
                .colHidden = False
                .Col = miSPDLOTLIST_SPEED
                .colHidden = False
                .Col = mispdLotList_ROUTE
                .colHidden = True
            End With
        End If
        Me.txtBoxQty.Text = "" 'add by Sam on 20080910 for ReqNO:M200807057
        Me.txtBoxQty.Enabled = False: Me.txtBoxQty.BackColor = &H8000000F
        miIndex = 3
    ElseIf (Me.cboStage.Text = gsSTAGE_FT And Me.cboLabelSpec.Text = modPrint.gsLABEL_FT_SMALL_LABEL) Then
        'msLabelSpec added by Jack on 2010/04/12 for JC201000060.
        '有值 : 由frmAdltProd執行此程式, 只能單筆列印, 只可列印RENESAS_SHIPPING Label.
        '無值: 由F1執行此程式, "WS_SMALL_LABEL"/"FT_SMALL_LABEL"都能多筆列印.
        If msLabelSpec <> "" Then
            Me.fraLabelForm.Item(4).Visible = True
        Else
            famLotList.Visible = True 'added by Jack on 2010/04/12 for JC201000060.
            With spdLotList
                .Col = mispdLotList_WQTY
                .colHidden = True
                .Col = miSPDLOTLIST_SPEED
                .colHidden = True
                .Col = mispdLotList_ROUTE
                .colHidden = False
            End With
        End If
        Me.txtBoxQty.Text = "" 'add by Sam on 20080910 for ReqNO:M200807057
        Me.txtBoxQty.Enabled = False: Me.txtBoxQty.BackColor = &H8000000F
        miIndex = 4
    'Added by Jack on 2016/06/03 for 2016 MES Phase-5 Project (FT eRunCard Paperless Phase-2 Project) <Start>
    ElseIf (Me.cboStage.Text = gsSTAGE_FT And Me.cboLabelSpec.Text = modPrint.gsLABEL_FT_LOT_INFO) Then
        famLotList.Visible = True 'added by Jack on 2010/04/12 for JC201000060.
        With spdLotList
            .Col = mispdLotList_WQTY
            .colHidden = True
            .Col = miSPDLOTLIST_SPEED
            .colHidden = True
            .Col = mispdLotList_ROUTE
            .colHidden = False
        End With
        
        Me.txtBoxQty.Text = ""
        Me.txtBoxQty.Enabled = False
        Me.txtBoxQty.BackColor = &H8000000F
        Me.txtBoxQty.Visible = False
        Me.lblQty.Item(0).Visible = False
        miIndex = 4
    'Added by Jack on 2016/06/03 for 2016 MES Phase-5 Project (FT eRunCard Paperless Phase-2 Project) <End>
    
    'Added by Jack on 2016/06/03 for 2016 MES Phase-5 Project (FT eRunCard Paperless Phase-2 Project) <Start>
    ElseIf (Me.cboStage.Text = gsSTAGE_FT And Me.cboLabelSpec.Text = modPrint.gsLABEL_FT_Label_PACK_INFO) Then
        Me.fraLabelForm.Item(8).Visible = True
        Me.txtBoxQty.Text = ""
        Me.txtBoxQty.Enabled = False
        Me.txtBoxQty.BackColor = &H8000000F
        Me.txtBoxQty.Visible = False
        Me.lblQty.Item(0).Visible = False
        Me.tabLabelPackInfo.Tab = 0 'Auto
        miIndex = 8
    'Added by Jack on 2016/06/03 for 2016 MES Phase-5 Project (FT eRunCard Paperless Phase-2 Project) <End>
    
    'Add by Tony Start on 20121001 for Req.JC201200283
    ElseIf (Me.cboStage.Text = gsSTAGE_FT And Me.cboLabelSpec.Text = modPrint.gsLABEL_FT_BOX_COUNTING) Then
            
        Me.lblQty.Item(0).Visible = False
        Me.txtBoxQty.Visible = False
        
        With Me.fraLabelForm.Item(4)
            .Visible = True '(把fraLabelForm內的元件除了LotId都隱藏)
            Me.txtProdNo.Item(4).Visible = False
            Me.lblProdNo.Item(4).Visible = False
            Me.txtWQty.Item(4).Visible = False
            Me.lblWQty.Item(4).Visible = False
            Me.txtCqty.Item(4).Visible = False
            Me.lblCQty.Item(4).Visible = False
            Me.txtStepName.Item(4).Visible = False
            Me.lblStepName.Item(4).Visible = False
            Me.txtCqty.Item(4).Visible = False
            Me.lblCQty.Item(0).Visible = False
            Me.txtLotOwner.Item(4).Visible = False
            Me.lblLotOwner.Item(1).Visible = False
            Me.txtRouteId.Item(4).Visible = False
            Me.lblRouteId.Item(4).Visible = False
        End With
        miIndex = 4
    'Add by Tony End on 20121001 for Req.JC201200283
    ElseIf Me.cboStage.Text = gsSTAGE_WS And _
    (Me.cboLabelSpec.Text = modPrint.gsLABEL_WSMCD_TO_SFG Or _
    Me.cboLabelSpec.Text = modPrint.gsLABEL_WS_MULTILOT_TO_SFG) Then
        Me.fraLabelForm.Item(5).Visible = True
        Me.txtBoxQty.Text = "" 'add by Sam on 20080910 for ReqNO:M200807057
        Me.txtBoxQty.Enabled = False: Me.txtBoxQty.BackColor = &H8000000F
        miIndex = 5
    ElseIf (Me.cboStage.Text = gsSTAGE_FT And Me.cboLabelSpec.Text = modPrint.gsLABEL_FT_ETEST_MERGE) Then
        Me.fraLabelForm.Item(6).Visible = True
        Me.txtBoxQty.Text = "" 'add by Sam on 20080910 for ReqNO:M200807057
        Me.txtBoxQty.Enabled = False: Me.txtBoxQty.BackColor = &H8000000F
        miIndex = 6
    ElseIf (Me.cboStage.Text = gsSTAGE_WS And Me.cboLabelSpec.Text = modPrint.gsLABEL_RENESAS_SHIPPING) Then 'add by Nelson on 2008/06/03 for ReqNo:M200806003
        Me.fraLabelForm.Item(7).Visible = True
        Me.txtBoxQty.Text = "" 'add by Sam on 20080910 for ReqNO:M200807057
        Me.txtBoxQty.Enabled = False: Me.txtBoxQty.BackColor = &H8000000F
'        Me.txtPrintQty.Text = "2" 'mark by Sam on 20080910 for ReqNO:M200807057
        miIndex = 7
    'Added by Jack on  2011/04/15 <Start>
    ElseIf (Me.cboStage.Text = gsSTAGE_FT And Me.cboLabelSpec.Text = modPrint.gsLABEL_FT_BIN_CARD_LABEL) Then
        Me.cmdOk.Visible = False
        lblQty(0).Visible = False
        lblQty(1).Visible = False
        lblQty(2).Visible = False
        lblQty(3).Visible = False
        txtBoxQty.Visible = False
        lblPrintQty(3).Caption = "Pass Sheet:"
        lblPrintFailQty.Visible = True
        cmdPrintPass.Visible = True
        cmdPrintFail.Visible = True
    
        Me.cmdPrintPass.Enabled = True
        Me.cmdPrintFail.Enabled = True
        Me.txtPrintFailQty.Enabled = True
        Me.txtPrintFailQty.Visible = True
        
        '有值 : 由frmFtDispatch執行此程式, 只能單筆列印, 只可列印RENESAS_SHIPPING Label.
        '無值: 由F1執行此程式, "WS_SMALL_LABEL"/"FT_SMALL_LABEL"都能多筆列印.
        If msLabelSpec <> "" Then
            Me.fraLabelForm.Item(4).Visible = True
        Else
            famLotList.Visible = True
            With spdLotList
                .Col = mispdLotList_WQTY
                .colHidden = True
                .Col = miSPDLOTLIST_SPEED
                .colHidden = True
                .Col = mispdLotList_ROUTE
                .colHidden = False
                .Col = mispdLotList_STEP
                .colHidden = False
            End With
        End If
        '必需自動調整Spread寬度. --> 2011/04/15 待修改.
        Me.txtBoxQty.Text = ""
        Me.txtBoxQty.Enabled = False: Me.txtBoxQty.BackColor = &H8000000F
        miIndex = 4
    'Added by Jack on  2011/04/15 <Start>
    'Add gsLABEL_FT_TR_LABEL by HouYu for 202000413 on 20220422
    ElseIf (Me.cboStage.Text = gsSTAGE_FT And Me.cboLabelSpec.Text = modPrint.gsLABEL_FT_TR_LABEL) Then
        Me.txtBoxQty.Visible = False
        Me.lblQty.Item(0).Visible = False
        
        Me.txt_FT_TR_IPN.Locked = True
        Me.txt_FT_TR_IPN.BackColor = &H80000000
        Me.fraLabelForm.Item(9).Visible = True
        miIndex = 9
    'Add gsLABEL_FT_TR_LABEL by HouYu for 202000413 on 20220422
    End If

End Sub

Private Sub cboStage_Click()
    Call GetLabelSpec
    'Added by Jack on 2010/04/12 for JC201000060. <Start>
    With Me.spdLotList
        .MaxRows = 0
        .MaxRows = 1
        Call .SetCellBorder(0, 0, .MaxCols, .MaxRows, 15, -1, CellBorderStyleSolid)
        Call ResizeSpdColumn(Me.spdLotList, moAppLog)
    End With
    'Added by Jack on 2010/04/12 for JC201000060. <End>
End Sub

Private Sub cmdCancel_Click()
    On Error Resume Next
    miResult = vbCancel
    
    'added by Jack on 20111/04/19 for JC201100104. <Start>
    If fraLabelForm(4).Visible = True Then
        Me.txtLotID(4).Text = ""
        Me.txtLotID(4).Locked = False
        Me.txtLotID(4).BackColor = vbWhite
        Me.txtProdNo(4).Text = ""
        Me.txtWQty(4).Text = ""
        Me.txtCqty(4).Text = ""
        Me.txtLotOwner(4).Text = ""
        Me.txtRouteId(4).Text = ""
        Me.txtLotID(4).SetFocus
        Me.txtStepName(4).Text = ""  'added by Jack on 20111/04/19 for JC201100104.
    End If
    Screen.MousePointer = vbDefault
    'added by Jack on 20111/04/19 for JC201100104. <End>
    
    Me.Hide
End Sub

'================================================================================
' Sub: cmdOk_Click()
'--------------------------------------------------------------------------------
' Description:  <Type your Sub description here...>
'--------------------------------------------------------------------------------
' Author:       Ivan, CIT 2001/01/14
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
Dim bIsPartial As Boolean
Dim colCompIds As FwStrings
Dim oLot       As FwLot
Dim oComp      As FwComponent
Dim sSysTime   As String
Dim dteNow     As Date
Dim iIndex     As Integer

Dim colRS           As Collection
Dim colWSMCDInfo    As Collection
Dim sParentLotID    As String
Dim sParentIPN      As String
Dim sMcpFlag        As String
Dim oItem           As Object
Dim iCQty           As Long
Dim iWQty           As Long

Dim sSQL    As String

Dim sMsg    As String 'add by Sam on 2008/06/12 for ReqNo:M200806023

Dim sBizType As String 'add by Sam on 2008/08/05 for ReqNo:M200807064

'Added by Jack on 2010/04/12 for JC201000060. <Start>
Dim bLotSelect  As Boolean
Dim iRowIndex As Integer
Dim vLotID    As Variant
Dim vProdNo   As Variant
Dim vWqty     As Variant
Dim vCqty     As Variant
Dim vLotOwner As Variant
Dim vSpeed    As Variant
Dim vRoute    As Variant
'Added by Jack on 2010/04/12 for JC201000060. <End>

Dim i           As Integer 'Add by HouYu for 202000413 on 20220422

'Add by Weilun on 20220527 for ReqNo.BE#202000233 <start>
Dim sFuSa           As String
Dim sIPN_Data       As String
'Add by Weilun on 20220527 for ReqNo.BE#202000233 <end>
Dim sLocation       As String 'Add by Weilun on 20220816 for ReqNo.BE#202200157
'----
' Init
'----
    sProcID = "cmdOk_Click"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog)
    Screen.MousePointer = vbHourglass
    Set colCompIds = New FwStrings
'----
' Condition Checking
'----
    If Len(Me.cboPrintServer.Text) = 0 Then
        Call UtShowMsgBox("Please select Printer Server!")
        GoTo ExitHandler
    End If
    If Len(Trim(Me.txtPrintQty.Text)) = 0 Then
        Call UtShowMsgBox("請輸入列印份數!")
        GoTo ExitHandler
    End If
    If Val(Me.txtPrintQty.Text) = "0" Then
        Call UtShowMsgBox("列印份數不能為零!")
        GoTo ExitHandler
    End If
    If (moFwWIP Is Nothing) Then
        Call RaiseError(glERR_INVALIDOBJECT, _
                        FormatErrorText(gsETX_INVALIDOBJECT, "FwWIP"))
    End If
    
    'msLabelSpec added by Jack on 2010/04/12 for JC201000060.
    '有值 : 由frmAdltProd執行此程式, 只能單筆列印, 只可列印RENESAS_SHIPPING Label.
    '無值: 由F1執行此程式, "WS_SMALL_LABEL"/"FT_SMALL_LABEL"都能多筆列印.
    '多筆時, 不用check Lot是否存在, 在Spread打勾時, 已check; 只需check是否至少選了一筆以上.
    '20110831 增加gsLABEL_WS_ENG_LOC_LABEL
    'Add gsLABEL_FT_BOX_COUNTING by Tony on 20121001 for Req.JC201200283
    'gsLABEL_FT_LOT_INFO Added by Jack on 2016/06/03 for 2016 MES Phase-5 Project (FT eRunCard Paperless Phase-2 Project)
    'Modify by Sam on 20180814 for BE工業 3.5 BE MES Phase 24 -- CP Auto Dispatch ,Add cp_small_label
    'Modify by Sam on 20201203 for CP 測前併批,新增CP VIRTUAL LABEL
'    If (msLabelSpec = "") And (Me.cboLabelSpec.Text = modPrint.gsLABEL_WS_SMALL_LABEL Or _
'         Me.cboLabelSpec.Text = modPrint.gsLABEL_CP_SMALL_LABEL Or _
'         Me.cboLabelSpec.Text = modPrint.gsLABEL_FT_LOT_INFO Or _
'         Me.cboLabelSpec.Text = modPrint.gsLABEL_FT_SMALL_LABEL) Then
    If (msLabelSpec = "") And (Me.cboLabelSpec.Text = modPrint.gsLABEL_WS_SMALL_LABEL Or _
         Me.cboLabelSpec.Text = modPrint.gsLABEL_CP_SMALL_LABEL Or _
         Me.cboLabelSpec.Text = modPrint.gsLABEL_CP_VIRTUAL_LOT_LABEL Or _
         Me.cboLabelSpec.Text = modPrint.gsLABEL_CP_VIRTUAL_MERGE Or _
         Me.cboLabelSpec.Text = modPrint.gsLABEL_FT_LOT_INFO Or _
         Me.cboLabelSpec.Text = modPrint.gsLABEL_FT_SMALL_LABEL) Then
         bLotSelect = False        'Check lot selection or not
        With spdLotList
              .Col = mispdLotList_CHECK
              For iIndex = 1 To .MaxRows
                  .Row = iIndex
                  If .Value = 1 Then
                      bLotSelect = True
                      Exit For
                  End If
              Next
        End With
        If bLotSelect = False Then
            Call UtShowMsgBox("Please Select Lot(s)!!")
            GoTo ExitHandler
        End If
        
    ElseIf Me.cboLabelSpec.Text = modPrint.gsLABEL_WS_ENG_LOC_LABEL Then
    
    'Added by Jack on 2016/06/03 for 2016 MES Phase-5 Project (FT eRunCard Paperless Phase-2 Project) <Start>
    ElseIf Me.cboLabelSpec.Text = modPrint.gsLABEL_FT_Label_PACK_INFO Then
        'miIndex = 8
        If tabLabelPackInfo.Tab = 0 Then 'Auto mode才需檢核Lot是否存在?
            Set oLot = moFwWIP.LotById(Trim(Me.txtLotID(miIndex).Text))
            If oLot Is Nothing Then
                Call UtShowMsgBox("The Lot' " & (Trim(Me.txtLotID(miIndex).Text)) & "' is not exist!")
                GoTo ExitHandler
            End If
        End If
    'Added by Jack on 2016/06/03 for 2016 MES Phase-5 Project (FT eRunCard Paperless Phase-2 Project) <End>
    Else '原本的單筆check rule.
        Set oLot = moFwWIP.LotById(Trim(Me.txtLotID(miIndex).Text))
        If oLot Is Nothing Then
            Call UtShowMsgBox("The Lot is not exist!")
            GoTo ExitHandler
        End If
    End If
                
    'add by vencent wei, 2002/09/05
    If Me.cboLabelSpec.Text = "" Then
        UtShowMsgBox "Please select one label type"
        GoTo ExitHandler
        
    'Add gsLABEL_FT_BOX_COUNTING by Tony on 20121001 for Req.JC201200283(此label不用輸入BoxQty,且也隱藏了)
    'Add gsLABEL_FT_TR_LABEL by HouYu for 202000413 on 20220422
    'gsLABEL_FT_LOT_INFO, gsLABEL_FT_Label_PACK_INFO Added by Jack on 2016/06/03 for 2016 MES Phase-5 Project (FT eRunCard Paperless Phase-2 Project)
    ElseIf Me.cboStage.Text = gsSTAGE_FT And _
    Me.cboLabelSpec.Text <> modPrint.gsLABEL_FT_SMALL_LABEL And _
    Me.cboLabelSpec.Text <> modPrint.gsLABEL_FT_ETEST_MERGE And _
    Me.cboLabelSpec.Text <> modPrint.gsLABEL_FT_LOT_INFO And _
    Me.cboLabelSpec.Text <> modPrint.gsLABEL_FT_Label_PACK_INFO And _
    Me.cboLabelSpec.Text <> modPrint.gsLABEL_FT_TR_LABEL And _
    Me.cboLabelSpec.Text <> modPrint.gsLABEL_FT_BOX_COUNTING Then
    
        If IsNumeric(Me.txtBoxQty) = False Or Trim(Me.txtBoxQty) = "0" Then
            UtShowMsgBox "Please Input Correct BoxQty"
            GoTo ExitHandler
        ElseIf IsNumeric(Me.txtQty(miIndex).Text) = False Or Trim(Me.txtQty(miIndex).Text) = "0" Then
            UtShowMsgBox "Please Input Correct LotQty"
            GoTo ExitHandler
        End If
    End If
    
    If Me.cboStage.Text = gsSTAGE_WS And Me.cboLabelSpec.Text = modPrint.gsLABEL_WS_SUMMARY Then
        Set colRS = moProRawSql.QueryDatabase("select " & gsCAT_TLI_MCP_FLAG & _
                                              " from " & gsCAT_TBL_LOT_INFO & _
                                              " where " & gsCAT_TLI_LOT_ID & " = '" & Me.txtLotID(miIndex).Text & "'")
        If colRS.Count > 0 Then
            If colRS.Item(1).Item(1) = "D" Then
                UtShowMsgBox "This Lot is MCD production!!  此批Lot為MCD產品!!"
            End If
        End If
    End If
'----
' Action
'----
    'Get System Time
    sSysTime = GetSystemTime(moAppLog, moFwWIP, moFwWF, moCwMbx)
    
    dteNow = ConvertToDate(sSysTime, moAppLog)
    
    'update by vencent Wei 2002/09/05
'    If miIndex = 0 Then
'        If Len(Trim(Me.txtQty(miIndex).Text)) = 0 Or Len(Trim(Me.txtBoxQty.Text)) = 0 Then
'            GoTo ExitHandler
'        End If
'        If CLng(Me.txtQty(miIndex).Text) > CLng(Me.txtBoxQty.Text) Then
'            bIsPartial = True
'        Else
'            bIsPartial = False
'        End If
'    End If
    
    Dim iIdx As Integer, iPrintTimes As Integer
    Dim sPrintQty As String
    
    'Add gsLABEL_FT_TR_LABEL by HouYu for 202000413 on 20220422
    'Add gsLABEL_FT_BOX_COUNTING by Tony on 20121001 for Req.JC201200283
    'gsLABEL_FT_LOT_INFO, gsLABEL_FT_Label_PACK_INFO Added by Jack on 2016/06/03 for 2016 MES Phase-5 Project (FT eRunCard Paperless Phase-2 Project)
    If (Me.cboStage.Text = gsSTAGE_FT And _
    (Me.cboLabelSpec.Text = modPrint.gsLABEL_FT_SMALL_LABEL _
    Or Me.cboLabelSpec.Text = modPrint.gsLABEL_FT_ETEST_MERGE _
    Or Me.cboLabelSpec.Text = modPrint.gsLABEL_FT_LOT_INFO _
    Or Me.cboLabelSpec.Text = modPrint.gsLABEL_FT_Label_PACK_INFO _
    Or Me.cboLabelSpec.Text = modPrint.gsLABEL_FT_TR_LABEL _
    Or Me.cboLabelSpec.Text = modPrint.gsLABEL_FT_BOX_COUNTING)) _
    Or Me.cboStage.Text = gsSTAGE_WS _
    Or Me.cboStage.Text = gsSTAGE_AY Then
        iPrintTimes = 1
    Else
        iPrintTimes = CLng(Me.txtQty(miIndex).Text) \ CLng(Me.txtBoxQty)
        If CLng(Me.txtQty(miIndex).Text) Mod CLng(Me.txtBoxQty) > 0 Then
            iPrintTimes = iPrintTimes + 1
        End If
    End If
    
    bIsPartial = False
    
    For iIndex = 1 To Val(Me.txtPrintQty.Text)
        For iIdx = 1 To iPrintTimes
            'Partial and Qty
            'Add gsLABEL_FT_TR_LABEL by HouYu for 202000413 on 20220422
            'Add gsLABEL_FT_BOX_COUNTING by Tony on 20121001 for Req.JC201200283
            'gsLABEL_FT_LOT_INFO Added by Jack on 2016/06/03 for 2016 MES Phase-5 Project (FT eRunCard Paperless Phase-2 Project)
            If (Me.cboStage.Text = gsSTAGE_FT _
            And (Me.cboLabelSpec.Text = modPrint.gsLABEL_FT_SMALL_LABEL _
            Or Me.cboLabelSpec.Text = modPrint.gsLABEL_FT_ETEST_MERGE _
            Or Me.cboLabelSpec.Text = modPrint.gsLABEL_FT_LOT_INFO _
            Or Me.cboLabelSpec.Text = modPrint.gsLABEL_FT_Label_PACK_INFO _
            Or Me.cboLabelSpec.Text = modPrint.gsLABEL_FT_TR_LABEL _
            Or Me.cboLabelSpec.Text = modPrint.gsLABEL_FT_BOX_COUNTING)) _
            Or Me.cboStage.Text <> gsSTAGE_FT Then
                bIsPartial = False
            Else
                If iIdx = iPrintTimes Then
                    If CLng(Me.txtQty(miIndex).Text) Mod CLng(Me.txtBoxQty) > 0 Then
                        bIsPartial = True
                        sPrintQty = CStr(CLng(Me.txtQty(miIndex).Text) Mod CLng(Me.txtBoxQty))
                    Else
                        bIsPartial = False
                        sPrintQty = Me.txtBoxQty
                    End If
                Else
                    bIsPartial = False
                    sPrintQty = Me.txtBoxQty
                End If
            End If
        
            'print
            If Me.cboStage.Text = gsSTAGE_FT And Me.cboLabelSpec.Text = modPrint.gsLABEL_TO_SUBFT_NORMAL Then
                Call modPrint.Prt_FT_To_SubFT_Normal(Trim(Me.txtLotID(miIndex).Text), _
                                                      Trim(Me.txtProdNo(miIndex).Text), _
                                                      sPrintQty, _
                                                      Trim(Me.txtUserId(miIndex).Text), _
                                                      bIsPartial, Me.cboPrintServer.Text, _
                                                      moCwMbx, moAppLog)
        
            ElseIf Me.cboStage.Text = gsSTAGE_FT And Me.cboLabelSpec.Text = modPrint.gsLABEL_TO_SUBFT_ENG_SAMPLE Then
                Call modPrint.Prt_FT_To_SubFT_Eng_Sample(Trim(Me.txtLotID(miIndex).Text), _
                                                      Trim(Me.txtProdNo(miIndex).Text), _
                                                      sPrintQty, _
                                                      Trim(Me.txtUserId(miIndex).Text), _
                                                      bIsPartial, Me.cboPrintServer.Text, _
                                                      moCwMbx, moAppLog)
            ElseIf Me.cboStage.Text = gsSTAGE_FT And Me.cboLabelSpec.Text = modPrint.gsLABEL_TO_SUBMK_NORMAL Then
                Call modPrint.Prt_FT_To_SubMK_Normal(Trim(Me.txtLotID(miIndex).Text), _
                                                    Trim(Me.txtProdNo(miIndex).Text), _
                                                    sPrintQty, _
                                                    Trim(Me.txtUserId(miIndex).Text), _
                                                    bIsPartial, Me.cboPrintServer.Text, _
                                                    moCwMbx, moAppLog)
            ElseIf Me.cboStage.Text = gsSTAGE_FT And Me.cboLabelSpec.Text = modPrint.gsLABEL_TO_SUBMK_ENG_SAMPLE Then
                Call modPrint.Prt_FT_To_SubMK_Eng_Sample(Trim(Me.txtLotID(miIndex).Text), _
                                                      Trim(Me.txtProdNo(miIndex).Text), _
                                                      sPrintQty, _
                                                      Trim(Me.txtUserId(miIndex).Text), _
                                                      bIsPartial, Me.cboPrintServer.Text, _
                                                      moCwMbx, moAppLog)
            ElseIf Me.cboStage.Text = gsSTAGE_FT And Me.cboLabelSpec.Text = modPrint.gsLABEL_TO_SUBTR_NORMAL Then
                Call modPrint.Prt_FT_To_SubTR_Normal(Trim(Me.txtLotID(miIndex).Text), _
                                                    Trim(Me.txtProdNo(miIndex).Text), _
                                                    sPrintQty, _
                                                    Trim(Me.txtUserId(miIndex).Text), _
                                                    bIsPartial, Me.cboPrintServer.Text, _
                                                    moCwMbx, moAppLog)
            ElseIf Me.cboStage.Text = gsSTAGE_FT And Me.cboLabelSpec.Text = modPrint.gsLABEL_TO_SUBTR_ENG_SAMPLE Then
                Call modPrint.Prt_FT_To_SubTR_Eng_Sample(Trim(Me.txtLotID(miIndex).Text), _
                                                      Trim(Me.txtProdNo(miIndex).Text), _
                                                      sPrintQty, _
                                                      Trim(Me.txtUserId(miIndex).Text), _
                                                      bIsPartial, Me.cboPrintServer.Text, _
                                                      moCwMbx, moAppLog)
            ElseIf Me.cboStage.Text = gsSTAGE_WS And Me.cboLabelSpec.Text = modPrint.gsLABEL_WS_TO_SFG Then
                Call modPrint.Prt_WS_WS_To_SFG(Trim(Me.txtLotID(miIndex).Text), _
                                                Trim(Me.txtProdNo(miIndex).Text), _
                                                Trim(Me.txtWQty(miIndex).Text), _
                                                Trim(Me.txtCqty(miIndex).Text), _
                                                Trim(Me.txtPartNo(miIndex).Text), _
                                                Me.cboPrintServer.Text, _
                                                moCwMbx, moAppLog)
            ElseIf Me.cboStage.Text = gsSTAGE_WS And Me.cboLabelSpec.Text = modPrint.gsLABEL_WS_SUMMARY Then
                For Each oComp In oLot.Components
                    If oComp.Status <> modConstGlobal.gsCOMP_STATUS_SCRAPPED Then
                        colCompIds.Add oComp.Attributes(gsLOT_CUSTOMATTR_CHIP_QTY), oComp.Id
                    End If
                Next oComp
                'Modified by Tony on 20120120 for Req.JC201200010
                'mm/dd/yy hh:mm to yyyy/mm/dd
                Call modPrint.Prt_WS_WS_SUMMARY(Trim(Me.txtLotID(miIndex).Text), _
                                                Trim(Me.txtProdNo(miIndex).Text), _
                                                Trim(Me.txtWQty(miIndex).Text), _
                                                Trim(Me.txtCqty(miIndex).Text), _
                                                colCompIds, _
                                                Trim(Me.txtUserId(miIndex).Text), _
                                                Format(dteNow, "yyyy/mm/dd"), _
                                                Me.cboPrintServer.Text, _
                                                moCwMbx, moAppLog)
            'Modify by Sam on 20180814 for BE工業 3.5 BE MES Phase 24 -- CP Auto Dispatch ,Add CP_SMALL_LABEL
            'Modify by Sam on 20201203 for CP 測前併批,新增CP VIRTUAL LABEL
'            ElseIf (Me.cboStage.Text = gsSTAGE_WS And Me.cboLabelSpec.Text = modPrint.gsLABEL_WS_SMALL_LABEL) Or _
'                (Me.cboStage.Text = gsSTAGE_AY And Me.cboLabelSpec.Text = modPrint.gsLABEL_WS_SMALL_LABEL) Or _
'                (Me.cboStage.Text = gsSTAGE_WS And Me.cboLabelSpec.Text = modPrint.gsLABEL_CP_SMALL_LABEL) Then
            ElseIf (Me.cboStage.Text = gsSTAGE_WS And Me.cboLabelSpec.Text = modPrint.gsLABEL_WS_SMALL_LABEL) Or _
                (Me.cboStage.Text = gsSTAGE_AY And Me.cboLabelSpec.Text = modPrint.gsLABEL_WS_SMALL_LABEL) Or _
                (Me.cboStage.Text = gsSTAGE_WS And Me.cboLabelSpec.Text = modPrint.gsLABEL_CP_VIRTUAL_LOT_LABEL) Or _
                (Me.cboStage.Text = gsSTAGE_WS And Me.cboLabelSpec.Text = modPrint.gsLABEL_CP_VIRTUAL_MERGE) Or _
                (Me.cboStage.Text = gsSTAGE_WS And Me.cboLabelSpec.Text = modPrint.gsLABEL_CP_SMALL_LABEL) Or _
                (Me.cboStage.Text = gsSTAGE_AY And Me.cboLabelSpec.Text = modPrint.gsLABEL_CP_VIRTUAL_MERGE) Or _
                (Me.cboStage.Text = gsSTAGE_AY And Me.cboLabelSpec.Text = modPrint.gsLABEL_CP_VIRTUAL_LOT_LABEL) Then
                'msLabelSpec added by Jack on 2010/04/12 for JC201000060.
                '有值 : 由frmAdltProd執行此程式, 只能單筆列印, 只可列印RENESAS_SHIPPING Label.(程式碼照舊)
                '無值: 由F1執行此程式, "WS_SMALL_LABEL"/"FT_SMALL_LABEL"都能多筆列印.
                If msLabelSpec <> "" Then
                    'Modified by Jack on 2012/06/13 for support KGD_AEB Project [Hanno]
                    '統一呼叫 modPrint.GetBizType( ).
                    'sBizType = GetBizType(Me.txtProdNo(miIndex).Text, miIndex) 'Add by Sam on 2008/08/05 for ReqNo:M200807064
                    'Modified by Jack on 2016/08/01 for JC201600108 <Start>
                    'BizType改為ProdLevel串ProdType (直接在Prt_WS_WS_SMALL_LABEL取得).
                    'sBizType = modPrint.GetBizType(Me.txtProdNo(miIndex).Text, moAppLog, moProRawSql)
                    'Marked by Jack on 2016/08/31 for JC2010600173 <Start>
                    'sBizType此處用不到, Mark掉, 以後只要改Prt_WS_WS_SMALL_LABEL( )
                    'sBizType = modPrint.GetProdType(Me.txtProdNo(miIndex).Text, moAppLog, moProRawSql)
                    'Marked by Jack on 2016/08/31 for JC2010600173 <End>
                    'Modified by Jack on 2016/08/01 for JC201600108 <End>
                    
                    'Modify by Sam start on 2008/08/05 for ReqNo:M200807064
    '                Call modPrint.Prt_WS_WS_SMALL_LABEL(Trim(Me.txtLotId(miIndex).Text), _
                                                        Trim(Me.txtProdNo(miIndex).Text), _
                                                        Trim(Me.txtWqty(miIndex).Text), _
                                                        Trim(Me.txtCqty(miIndex).Text), _
                                                        Trim(Me.txtLotOwner(miIndex).Text), _
                                                        Trim(Me.txtSpeed(miIndex).Text), _
                                                        Me.cboPrintServer.Text, _
                                                        moCwMbx, moAppLog)
                    'Add by Sam Start on 20180814 for BE工業 3.5 BE MES Phase 24 -- CP Auto Dispatch
                    If Me.cboLabelSpec.Text = modPrint.gsLABEL_CP_SMALL_LABEL Then
                        Call modPrint.Prt_WS_CP_SMALL_LABEL(Trim(Me.txtLotID(miIndex).Text), _
                                                            Trim(Me.txtProdNo(miIndex).Text), _
                                                            Trim(Me.txtWQty(miIndex).Text), _
                                                            Trim(Me.txtCqty(miIndex).Text), _
                                                            Trim(Me.txtLotOwner(miIndex).Text), _
                                                            Trim(sBizType), _
                                                            Me.cboPrintServer.Text, _
                                                            moCwMbx, moAppLog, moProRawSql)
                    'Add by Sam start on 20201203 for Project CP 測前併批
                    ElseIf Me.cboLabelSpec.Text = modPrint.gsLABEL_CP_VIRTUAL_LOT_LABEL Then
                        Call modPrint.Prt_WS_CP_VIRTUAL_LOT_LABEL(Trim(Me.txtLotID(miIndex).Text), _
                                                            Trim(Me.txtProdNo(miIndex).Text), _
                                                            Trim(Me.txtWQty(miIndex).Text), _
                                                            Trim(Me.txtCqty(miIndex).Text), _
                                                            Me.cboPrintServer.Text, _
                                                            moCwMbx, moAppLog, moProRawSql)
                    ElseIf Me.cboLabelSpec.Text = modPrint.gsLABEL_CP_VIRTUAL_MERGE Then
                        Call modPrint.Prt_WS_CP_VIRTUAL_LOT_LABEL(Trim(Me.txtLotID(miIndex).Text), _
                                                            Trim(Me.txtProdNo(miIndex).Text), _
                                                            Trim(Me.txtWQty(miIndex).Text), _
                                                            Trim(Me.txtCqty(miIndex).Text), _
                                                            Me.cboPrintServer.Text, _
                                                            moCwMbx, moAppLog, moProRawSql, True)
                    'Add by Sam END on 20201203 for Project CP 測前併批
                    Else
                    'Add by Sam End on 20180814 for BE工業 3.5 BE MES Phase 24 -- CP Auto Dispatch
                        Call modPrint.Prt_WS_WS_SMALL_LABEL(Trim(Me.txtLotID(miIndex).Text), _
                                                            Trim(Me.txtProdNo(miIndex).Text), _
                                                            Trim(Me.txtWQty(miIndex).Text), _
                                                            Trim(Me.txtCqty(miIndex).Text), _
                                                            Trim(Me.txtLotOwner(miIndex).Text), _
                                                            Trim(sBizType), _
                                                            Me.cboPrintServer.Text, _
                                                            moCwMbx, moAppLog, moProRawSql)
                    End If
                    'Modify by Sam end on 2008/08/05 for ReqNo:M200807064
                Else
                    With Me.spdLotList
                        For iRowIndex = 1 To .MaxRows
                            .Col = mispdLotList_CHECK
                            .Row = iRowIndex
                            If .Text = 1 Then '打勾才要列印.
                                .GetText miSPDLOTLIST_LOT_ID, iRowIndex, vLotID
                                .GetText miSPDLOTLIST_PROD_NO, iRowIndex, vProdNo
                                .GetText mispdLotList_WQTY, iRowIndex, vWqty
                                .GetText mispdLotList_CQTY, iRowIndex, vCqty
                                .GetText miSPDLOTLIST_LOT_OWNER, iRowIndex, vLotOwner
                                .GetText miSPDLOTLIST_SPEED, iRowIndex, vSpeed
                                .GetText mispdLotList_ROUTE, iRowIndex, vRoute
                                
                                'Modified by Jack on 2012/06/13 for support KGD_AEB Project [Hanno]
                                '統一呼叫 modPrint.GetBizType( ).
                                'sBizType = GetBizType(Trim(vProdNo), miIndex) 'Add by Sam on 2008/08/05 for ReqNo:M200807064
                                'Modified by Jack on 2016/08/01 for JC201600108 <Start>
                                'BizType改為ProdLevel串ProdType.
                                'sBizType = modPrint.GetBizType(Trim(vProdNo), moAppLog, moProRawSql)
                                'Marked by Jack on 2016/08/31 for JC2010600173 <Start>
                                'sBizType此處用不到, Mark掉, 以後只要改Prt_WS_WS_SMALL_LABEL( )
                                'sBizType = modPrint.GetProdType(Trim(vProdNo), moAppLog, moProRawSql)
                                'Marked by Jack on 2016/08/31 for JC2010600173 <End>
                                'Modified by Jack on 2016/08/01 for JC201600108 <End>
                                'Add by Sam Start on 20180814 for BE工業 3.5 BE MES Phase 24 -- CP Auto Dispatch
                                If Me.cboLabelSpec.Text = modPrint.gsLABEL_CP_SMALL_LABEL Then
                                    Call modPrint.Prt_WS_CP_SMALL_LABEL(Trim(vLotID), _
                                                                        Trim(vProdNo), _
                                                                        Trim(vWqty), _
                                                                        Trim(vCqty), _
                                                                        Trim(vLotOwner), _
                                                                        Trim(sBizType), _
                                                                        Me.cboPrintServer.Text, _
                                                                        moCwMbx, moAppLog, moProRawSql)
                                'Add by Sam End on 20180814 for BE工業 3.5 BE MES Phase 24 -- CP Auto Dispatch
                                'Add by Sam start on 20201203 for Project CP 測前併批
                                ElseIf Me.cboLabelSpec.Text = modPrint.gsLABEL_CP_VIRTUAL_LOT_LABEL Then
                                    Call modPrint.Prt_WS_CP_VIRTUAL_LOT_LABEL(Trim(vLotID), _
                                                                        Trim(vProdNo), _
                                                                        Trim(vWqty), _
                                                                        Trim(vCqty), _
                                                                        Me.cboPrintServer.Text, _
                                                                        moCwMbx, moAppLog, moProRawSql)
                                ElseIf Me.cboLabelSpec.Text = modPrint.gsLABEL_CP_VIRTUAL_MERGE Then
                                    Call modPrint.Prt_WS_CP_VIRTUAL_LOT_LABEL(Trim(vLotID), _
                                                                        Trim(vProdNo), _
                                                                        Trim(vWqty), _
                                                                        Trim(vCqty), _
                                                                        Me.cboPrintServer.Text, _
                                                                        moCwMbx, moAppLog, moProRawSql, True)
                                'Add by Sam END on 20201203 for Project CP 測前併批
                                Else
                                    Call modPrint.Prt_WS_WS_SMALL_LABEL(Trim(vLotID), _
                                                                        Trim(vProdNo), _
                                                                        Trim(vWqty), _
                                                                        Trim(vCqty), _
                                                                        Trim(vLotOwner), _
                                                                        Trim(sBizType), _
                                                                        Me.cboPrintServer.Text, _
                                                                        moCwMbx, moAppLog, moProRawSql)
                                End If
                            End If
                            Next
                    End With
                End If
            ElseIf Me.cboStage.Text = gsSTAGE_FT And Me.cboLabelSpec.Text = modPrint.gsLABEL_FT_SMALL_LABEL Then
                'Add by Sky at 2005/05/18 for M200501046
                Dim sCustomer As String
                Dim sCsum As String
                Dim sGreen As String
                Dim sIPN As String
                Dim sGPType As String
                Dim sCargradeFlag As String 'Add by Tony on 20110215 for Req.JC201100038
                Dim sProdLine     As String 'Add by Jack on 2015/06/25 for Req.JC201500212
                    
                sCustomer = ""
                sCsum = ""
                sGreen = ""
                sIPN = ""
                sGPType = ""
                'Marked by Tony start on 20071102 requirement:M200710061
'                sSql = "SELECT tbl_customer_master.customername, tbl_prm_backend_spec.ipn  " & _
'                "  From tbl_lot_attribute, tbl_customer_master, tbl_prm_backend_spec " & _
'                "  Where tbl_lot_attribute.IPN = tbl_prm_backend_spec.IPN " & _
'                "   AND tbl_prm_backend_spec.customerno = tbl_customer_master.customerno " & _
'                "   AND tbl_prm_backend_spec.defaults = 'Y' " & _
'                "   AND lotid = '" & Trim(Me.txtLotId(miIndex).Text) & "'" & _
'                "    AND ROWNUM = 1 "
'                Set colRS = moProRawSql.QueryDatabase(sSql)
'                If Not colRS Is Nothing Then
'                    If colRS.Count > 0 Then
'                        sCustomer = colRS.Item(1).Item(1)
'                        sIPN = colRS.Item(1).Item(2)
'                    End If
'                End If
'                Set colRS = Nothing
'
'                sSql = "SELECT tbl_ipn_master.checksum, tbl_ipn_master.gptype " & _
'                "  From tbl_ipn_master " & _
'                "  where tbl_ipn_master.IPN = '" & sIPN & "'"
'                Set colRS = moProRawSql.QueryDatabase(sSql)
'                If Not colRS Is Nothing Then
'                    If colRS.Count > 0 Then
'                        sGPType = colRS.Item(1).Item(2)
'                        sCsum = colRS.Item(1).Item(1)
'                    End If
'                End If
                 'Marked by Tony end on 20071102 requirement:M200710061
                'msLabelSpec added by Jack on 2010/04/12 for JC201000060.
                '有值 : 由frmAdltProd執行此程式, 只能單筆列印, 只可列印RENESAS_SHIPPING Label.(程式碼照舊)
                '無值: 由F1執行此程式, "WS_SMALL_LABEL"/"FT_SMALL_LABEL"都能多筆列印.
                If msLabelSpec <> "" Then
                
                    'Add by Weilun on 20220816 for ReqNo.BE#202200157 <start>
                    sLocation = ""
                    sSQL = " select tsl." & gsCAT_TSL_LOCATION & " " & _
                             " from " & gsCAT_TBL_STEP_LOCATION & " tsl, " & _
                                        gsCAT_TBL_LOT_ATTRIBUTE & " tlatt " & _
                            " where tsl." & gsCAT_TSL_STEPID & " = tlatt." & gsCAT_TLATT_STEPID & " " & _
                              " and tlatt." & gsCAT_TLATT_LOTID & " = '" & Trim(Me.txtLotID(miIndex).Text) & "' "
                    Set colRS = moProRawSql.QueryDatabase(sSQL)
                    If Not colRS Is Nothing Then
                        If colRS.Count > 0 Then
                            sLocation = colRS.Item(1).Item(gsCAT_TSL_LOCATION)
                        End If
                    End If
                    'Add by Weilun on 20220816 for ReqNo.BE#202200157 <end>
                
                    sIPN_Data = "" 'Add by Weilun on 20220526 for ReqNo.BE#202000233
                
                     'Add by Tony start on 20071102 requirement:M200710061
                     'Add CarGradeFlag by Tony on 20110215 for Req.JC201100038
                     'Add FuSa,prodbody,pincount,package_code,bodysize by Weilun on 20220526 for ReqNo.BE#202000233
                    sSQL = "SELECT b." & gsCAT_TCUM_CUSTOMERNAME & ",c." & gsCAT_TPBS_IPN & _
                                 ",d." & gsCAT_TIM_GPTYPE & ",d." & gsCAT_TIM_CHECKSUM & _
                                 ",d." & gsCAT_TIM_CARGRADEFLAG & _
                                 ",d." & gsCAT_TIM_FUSA & _
                                 ",d." & gsCAT_TIM_PRODBODY & _
                                 ",d." & gsCAT_TIM_PINCOUNT & _
                                 ",d." & gsCAT_TIM_PACKAGE_CODE & _
                                 ",d." & gsCAT_TIM_BODY_SIZE & _
                          " From " & gsCAT_TBL_LOT_ATTRIBUTE & " a," & gsCAT_TBL_CUSTOMER_MASTER & " b," _
                          & gsCAT_TBL_PRM_BE_SPEC & " c," & gsCAT_TBL_IPN_MASTER & " d" _
                          & " Where a." & gsCAT_TLATT_IPN & " = c." & gsCAT_TPBS_IPN & "(+)" _
                          & " and c." & gsCAT_TPBS_CUSTOMER_NO & "= b." & gsCAT_TCUM_CUSTOMERNO & "(+)" _
                          & " and c." & gsCAT_TPBS_DEFAULTS & "(+) = 'Y'" _
                          & " and a." & gsCAT_TLATT_LOTID & " = '" & Trim(Me.txtLotID(miIndex).Text) & "'" _
                          & " and a." & gsCAT_TLATT_IPN & "=d." & gsCAT_TIM_IPN _
                          & " and ROWNUM = 1"
                    Set colRS = moProRawSql.QueryDatabase(sSQL)
                    If Not colRS Is Nothing Then
                        If colRS.Count > 0 Then
                            sCustomer = colRS.Item(1).Item(1)
                            sIPN = colRS.Item(1).Item(2)
                            sGPType = colRS.Item(1).Item(3)
                            sCsum = colRS.Item(1).Item(4)
                            sCargradeFlag = colRS.Item(1).Item(5) 'Add by Tony on 20110215 for Req.JC201100038
                            
                            'Add by Weilun on 20220526 for ReqNo.BE#202000233 <start>
                            sFuSa = colRS.Item(1).Item(gsCAT_TIM_FUSA)
                            If sLocation = "FT" Then 'FT的才用這種格式, 其餘的就用IPN(下方SQL抓), Add by Weilun on 20220816 for ReqNo.BE#202200157
                                sIPN_Data = colRS.Item(1).Item(gsCAT_TIM_PRODBODY) & "-" & _
                                            colRS.Item(1).Item(gsCAT_TIM_PINCOUNT) & _
                                            colRS.Item(1).Item(gsCAT_TIM_PACKAGE_CODE) & " " & _
                                            colRS.Item(1).Item(gsCAT_TIM_BODY_SIZE)
                            End If
                            'Add by Weilun on 20220526 for ReqNo.BE#202000233 <end>
                        End If
                    End If
                     'Add by Tony end on 20071102 requirement:M200710061
                     
                     'Add by Weilun on 20220526 for ReqNo.BE#202000233 <start>
                    '當TBL_IPN_MASTER不存在導致上面SQL取不到sIPN_Data時, 改用TBL_LOT_ATTRIBUTE的IPN
                    If sIPN_Data = "" Then
                        sSQL = "SELECT " & gsCAT_TLATT_IPN & " " & _
                                " From " & gsCAT_TBL_LOT_ATTRIBUTE & " " & _
                               " Where " & gsCAT_TLATT_LOTID & " = '" & Trim(Me.txtLotID(miIndex).Text) & "' "
                        Set colRS = moProRawSql.QueryDatabase(sSQL)
                        If colRS.Count > 0 Then
                            sIPN_Data = colRS.Item(1).Item(gsCAT_TLATT_IPN)
                        End If
                    End If
                    'Add by Weilun on 20220526 for ReqNo.BE#202000233 <end>
                     
                    Set colRS = Nothing
                    
                    If sGPType = "XX" Then
                        sGreen = "N"
                    ElseIf Len(sGPType) = 0 Then
                        sGreen = " "
                    Else
                        sGreen = "Y"
                    End If
                    
                    'Add by Jack on 2015/06/25 for Req.JC201500212 <Start>.
                    '以IPN 前4碼去串 tbl_prod_body.prodline
                    sProdLine = ""
                    sSQL = " SELECT A." & gsCAT_TPB_PROD_LINE & _
                           "   FROM " & gsCAT_TBL_PROD_BODY & " A " & _
                           "  WHERE A." & gsCAT_TPB_PROD_BODY & "= substr('" & Trim(Me.txtProdNo(miIndex).Text) & "', 1,4) "
                    Set colRS = moProRawSql.QueryDatabase(sSQL)
                    If colRS.Count > 0 Then
                        sProdLine = colRS.Item(1).Item(gsCAT_TPB_PROD_LINE)
                    End If
                    'Add by Jack on 2015/06/25 for Req.JC201500212 <End>.
                    
                    'Add CarGradeFlag by Tony on 20110215 for Req.JC201100038
                    'sProdLine Add by Jack on 2015/06/25 for Req.JC201500212
                    'Add sFuSa & 取消LotOwner & 取消Csum & Trim(Me.txtProdNo(miIndex).Text)-->sIPN_Data, Modify by Weilun on 20220527 for ReqNo.BE#202000233
                    'Add sLocation, 並恢復部分輸入, Modify by Weilun on 20220816 for ReqNo.BE#202200157
                    Call modPrint.Prt_FT_FT_SMALL_LABEL(Trim(Me.txtLotID(miIndex).Text), _
                                                        sIPN_Data, _
                                                        Trim(Me.txtWQty(miIndex).Text), _
                                                        Trim(Me.txtCqty(miIndex).Text), _
                                                        Trim(Me.txtLotOwner(miIndex).Text), _
                                                        Trim(Me.txtRouteId(miIndex).Text), _
                                                        sCustomer, sCsum, sGreen, _
                                                        Me.cboPrintServer.Text, _
                                                        moCwMbx, moAppLog, sCargradeFlag, sProdLine, "", sFuSa, sLocation)
'                    Call modPrint.Prt_FT_FT_SMALL_LABEL(Trim(Me.txtLotId(miIndex).Text), _
'                                                        sIPN_Data, _
'                                                        Trim(Me.txtWqty(miIndex).Text), _
'                                                        Trim(Me.txtCqty(miIndex).Text), _
'                                                        "", _
'                                                        Trim(Me.txtRouteId(miIndex).Text), _
'                                                        sCustomer, "", sGreen, _
'                                                        Me.cboPrintServer.Text, _
'                                                        moCwMbx, moAppLog, sCargradeFlag, sProdLine, "", sFusa)
'                    Call modPrint.Prt_FT_FT_SMALL_LABEL(Trim(Me.txtLotId(miIndex).Text), _
'                                                        Trim(Me.txtProdNo(miIndex).Text), _
'                                                        Trim(Me.txtWqty(miIndex).Text), _
'                                                        Trim(Me.txtCqty(miIndex).Text), _
'                                                        Trim(Me.txtLotOwner(miIndex).Text), _
'                                                        Trim(Me.txtRouteId(miIndex).Text), _
'                                                        sCustomer, sCsum, sGreen, _
'                                                        Me.cboPrintServer.Text, _
'                                                        moCwMbx, moAppLog, sCargradeFlag, sProdLine)
                Else
                    With Me.spdLotList
                        For iRowIndex = 1 To .MaxRows
                            .Col = mispdLotList_CHECK
                            .Row = iRowIndex
                            If .Text = 1 Then '打勾才要列印.
                                .GetText miSPDLOTLIST_LOT_ID, iRowIndex, vLotID
                                .GetText miSPDLOTLIST_PROD_NO, iRowIndex, vProdNo
                                .GetText mispdLotList_WQTY, iRowIndex, vWqty
                                .GetText mispdLotList_CQTY, iRowIndex, vCqty
                                .GetText miSPDLOTLIST_LOT_OWNER, iRowIndex, vLotOwner
                                .GetText miSPDLOTLIST_SPEED, iRowIndex, vSpeed
                                .GetText mispdLotList_ROUTE, iRowIndex, vRoute
                                
                                'Add by Weilun on 20220816 for ReqNo.BE#202200157 <start>
                                sLocation = ""
                                sSQL = " select tsl." & gsCAT_TSL_LOCATION & " " & _
                                         " from " & gsCAT_TBL_STEP_LOCATION & " tsl, " & _
                                                    gsCAT_TBL_LOT_ATTRIBUTE & " tlatt " & _
                                        " where tsl." & gsCAT_TSL_STEPID & " = tlatt." & gsCAT_TLATT_STEPID & " " & _
                                          " and tlatt." & gsCAT_TLATT_LOTID & " = '" & Trim(vLotID) & "' "
                                Set colRS = moProRawSql.QueryDatabase(sSQL)
                                If Not colRS Is Nothing Then
                                    If colRS.Count > 0 Then
                                        sLocation = colRS.Item(1).Item(gsCAT_TSL_LOCATION)
                                    End If
                                End If
                                'Add by Weilun on 20220816 for ReqNo.BE#202200157 <end>
                                
                                sIPN_Data = "" 'Add by Weilun on 20220526 for ReqNo.BE#202000233
                                
                                'Add CarGradeFlag by Tony on 20110215 for Req.JC201100038
                                'Add FuSa,prodbody,pincount,package_code,bodysize by Weilun on 20220526 for ReqNo.BE#202000233
                                sSQL = "SELECT b." & gsCAT_TCUM_CUSTOMERNAME & ",c." & gsCAT_TPBS_IPN & _
                                             ",d." & gsCAT_TIM_GPTYPE & ",d." & gsCAT_TIM_CHECKSUM & _
                                             ",d." & gsCAT_TIM_CARGRADEFLAG & _
                                             ",d." & gsCAT_TIM_FUSA & _
                                             ",d." & gsCAT_TIM_PRODBODY & _
                                             ",d." & gsCAT_TIM_PINCOUNT & _
                                             ",d." & gsCAT_TIM_PACKAGE_CODE & _
                                             ",d." & gsCAT_TIM_BODY_SIZE & _
                                      " From " & gsCAT_TBL_LOT_ATTRIBUTE & " a," & gsCAT_TBL_CUSTOMER_MASTER & " b," _
                                      & gsCAT_TBL_PRM_BE_SPEC & " c," & gsCAT_TBL_IPN_MASTER & " d" _
                                      & " Where a." & gsCAT_TLATT_IPN & " = c." & gsCAT_TPBS_IPN & "(+)" _
                                      & " and c." & gsCAT_TPBS_CUSTOMER_NO & "= b." & gsCAT_TCUM_CUSTOMERNO & "(+)" _
                                      & " and c." & gsCAT_TPBS_DEFAULTS & "(+) = 'Y'" _
                                      & " and a." & gsCAT_TLATT_LOTID & " = '" & Trim(vLotID) & "'" _
                                      & " and a." & gsCAT_TLATT_IPN & "=d." & gsCAT_TIM_IPN _
                                      & " and ROWNUM = 1"

                                Set colRS = moProRawSql.QueryDatabase(sSQL)
                                If Not colRS Is Nothing Then
                                    If colRS.Count > 0 Then
                                        sCustomer = colRS.Item(1).Item(1)
                                        sIPN = colRS.Item(1).Item(2)
                                        sGPType = colRS.Item(1).Item(3)
                                        sCsum = colRS.Item(1).Item(4)
                                        sCargradeFlag = colRS.Item(1).Item(5) 'Add by Tony on 20110215 for Req.JC201100038
                                        
                                        'Add by Weilun on 20220526 for ReqNo.BE#202000233 <start>
                                        sFuSa = colRS.Item(1).Item(gsCAT_TIM_FUSA)
                                        If sLocation = "FT" Then 'FT的才用這種格式, 其餘的就用IPN(下方SQL抓), Add by Weilun on 20220816 for ReqNo.BE#202200157
                                            sIPN_Data = colRS.Item(1).Item(gsCAT_TIM_PRODBODY) & "-" & _
                                                        colRS.Item(1).Item(gsCAT_TIM_PINCOUNT) & _
                                                        colRS.Item(1).Item(gsCAT_TIM_PACKAGE_CODE) & " " & _
                                                        colRS.Item(1).Item(gsCAT_TIM_BODY_SIZE)
                                        End If
                                        'Add by Weilun on 20220526 for ReqNo.BE#202000233 <end>
                                    End If
                                End If
                                '當TBL_IPN_MASTER不存在導致上面SQL取不到sIPN_Data時, 改用TBL_LOT_ATTRIBUTE的IPN
                                If sIPN_Data = "" Then
                                    sSQL = "SELECT " & gsCAT_TLATT_IPN & " " & _
                                            " From " & gsCAT_TBL_LOT_ATTRIBUTE & " " & _
                                           " Where " & gsCAT_TLATT_LOTID & " = '" & Trim(vLotID) & "' "
                                    Set colRS = moProRawSql.QueryDatabase(sSQL)
                                    If colRS.Count > 0 Then
                                        sIPN_Data = colRS.Item(1).Item(gsCAT_TLATT_IPN)
                                    End If
                                End If
                                'Add by Weilun on 20220526 for ReqNo.BE#202000233 <end>
                                Set colRS = Nothing
                                
                                If sGPType = "XX" Then
                                    sGreen = "N"
                                ElseIf Len(sGPType) = 0 Then
                                    sGreen = " "
                                Else
                                    sGreen = "Y"
                                End If
                                
                                'Add by Jack on 2015/06/25 for Req.JC201500212 <Start>.
                                '以IPN 前4碼去串 tbl_prod_body.prodline
                                sProdLine = ""
                                sSQL = " SELECT A." & gsCAT_TPB_PROD_LINE & _
                                       "   FROM " & gsCAT_TBL_PROD_BODY & " A " & _
                                       "  WHERE A." & gsCAT_TPB_PROD_BODY & "= substr('" & Trim(vProdNo) & "', 1,4) "
                                Set colRS = moProRawSql.QueryDatabase(sSQL)
                                If colRS.Count > 0 Then
                                    sProdLine = colRS.Item(1).Item(gsCAT_TPB_PROD_LINE)
                                End If
                                'Add by Jack on 2015/06/25 for Req.JC201500212 <End>.
                                
                                'Add CarGradeFlag by Tony on 20110215 for Req.JC201100038
                                'sProdLine Add by Jack on 2015/06/25 for Req.JC201500212
                                'Add sFuSa & 取消LotOwner & 取消Csum & Trim(vProdNo)-->sIPNData, Modify by Weilun on 20220526 for ReqNo.BE#202000233
                                'Add sLocation, 並恢復部分輸入, Modify by Weilun on 20220816 for ReqNo.BE#202200157
                                Call modPrint.Prt_FT_FT_SMALL_LABEL(Trim(vLotID), _
                                                                    sIPN_Data, _
                                                                    Trim(vWqty), _
                                                                    Trim(vCqty), _
                                                                    Trim(vLotOwner), _
                                                                    Trim(vRoute), _
                                                                    sCustomer, sCsum, sGreen, _
                                                                    Me.cboPrintServer.Text, _
                                                                    moCwMbx, moAppLog, sCargradeFlag, sProdLine, "", sFuSa, sLocation)
'                                Call modPrint.Prt_FT_FT_SMALL_LABEL(Trim(vLotID), _
'                                                                    sIPN_Data, _
'                                                                    Trim(vWqty), _
'                                                                    Trim(vCqty), _
'                                                                    "", _
'                                                                    Trim(vRoute), _
'                                                                    sCustomer, "", sGreen, _
'                                                                    Me.cboPrintServer.Text, _
'                                                                    moCwMbx, moAppLog, sCargradeFlag, sProdLine, "", sFusa)
'                                Call modPrint.Prt_FT_FT_SMALL_LABEL(Trim(vLotID), _
'                                                                    Trim(vProdNo), _
'                                                                    Trim(vWqty), _
'                                                                    Trim(vCqty), _
'                                                                    Trim(vLotOwner), _
'                                                                    Trim(vRoute), _
'                                                                    sCustomer, sCsum, sGreen, _
'                                                                    Me.cboPrintServer.Text, _
'                                                                    moCwMbx, moAppLog, sCargradeFlag, sProdLine)
                        End If
                        Next
                    End With
                End If
                '------------------------------------
            'Added by Jack on 2016/06/03 for 2016 MES Phase-5 Project (FT eRunCard Paperless Phase-2 Project) <Start>
            ElseIf Me.cboStage.Text = gsSTAGE_FT And Me.cboLabelSpec.Text = modPrint.gsLABEL_FT_LOT_INFO Then
                Dim sVendorCode       As String
                Dim sLotipn           As String
                Dim sCQty             As String
                Dim sRouteId          As String
                Dim sLotOwner         As String
                Dim sDateCode         As String
                Dim sIpnCarGradeFlag  As String
                Dim sIpnProdLine      As String
                Dim sCarrierType      As String
                Dim sExtraRomFlag     As String
                Dim sBrand            As String
                Dim sIcDrawing        As String
                Dim sMarkingSpecNo    As String
                Dim sIpnGreen         As String
                Dim sIpnCsum          As String
                Dim sNFlag            As String
                                
                With Me.spdLotList
                    For iRowIndex = 1 To .MaxRows
                        .Col = mispdLotList_CHECK
                        .Row = iRowIndex
                        If .Text = 1 Then '打勾才要列印.
                            .GetText miSPDLOTLIST_LOT_ID, iRowIndex, vLotID
                            .GetText miSPDLOTLIST_PROD_NO, iRowIndex, vProdNo
                            .GetText mispdLotList_WQTY, iRowIndex, vWqty
                            .GetText mispdLotList_CQTY, iRowIndex, vCqty
                            .GetText miSPDLOTLIST_LOT_OWNER, iRowIndex, vLotOwner
                            .GetText miSPDLOTLIST_SPEED, iRowIndex, vSpeed
                            .GetText mispdLotList_ROUTE, iRowIndex, vRoute
                            
                            Set oLot = moFwWIP.LotById(Trim(vLotID))
                            If oLot Is Nothing Then
                                Call UtShowMsgBox("The Lot '" & vLotID & "' is not exist!")
                                GoTo ExitHandler
                            End If

'Marked by Jack on 2016/07/15 for 2016 MES Phase-5 Project (FT eRunCard Paperless Phase-2 Project) <Start>
'                            'ps : 以IPN 前4碼去串 tbl_prod_body.prodline
'                            sSql = "select a." & gsCAT_TLI_VENDORCODE & " ," & _
'                                             " b." & gsCAT_TLATT_IPN & "," & _
'                                             " b." & gsCAT_TLATT_CHIPQTY & "," & _
'                                             " b." & gsCAT_TLATT_ROUTE & "," & _
'                                             " b." & gsCAT_TLATT_LOTOWNER & "," & _
'                                             " b." & gsCAT_TLATT_DATECODE & "," & _
'                                             " c." & gsCAT_TIM_CARGRADEFLAG & "," & _
'                                             " e." & gsCAT_TPB_PROD_LINE & "," & _
'                                             " c." & gsCAT_TIM_CARRIER_TYPE & "," & _
'                                             " c." & gsCAT_TIM_EXTRAROMFLAG & "," & _
'                                             " c." & gsCAT_TIM_BRAND & "," & _
'                                             " c." & gsCAT_TIM_ICDRAWING & "," & _
'                                             " c." & gsCAT_TIM_MARKING_SPEC_NO & "," & _
'                                             " decode(c." & gsCAT_TIM_GPTYPE & ", null, null, 'XX', 'N', 'Y') as Green," & _
'                                             " c." & gsCAT_TIM_CHECKSUM & "," & _
'                                             " decode(d." & gsCAT_TPBS_CUSTOMER_NO & ", '0000010750', 'Y', 'N') as NFlag"
'                            sSql = sSql _
'                                        & " from " & gsCAT_TBL_LOT_INFO & " a, " _
'                                                   & gsCAT_TBL_LOT_ATTRIBUTE & " b, " _
'                                                   & gsCAT_TBL_IPN_MASTER & " c, " _
'                                                   & gsCAT_TBL_PRM_BE_SPEC & " d, " _
'                                                   & gsCAT_TBL_PROD_BODY & " e " _
'                                        & " where a." & gsCAT_TLI_LOT_ID & " = '" & vLotID & "' " _
'                                        & " and a." & gsCAT_TLI_LOT_ID & " = b." & gsCAT_TLATT_LOTID & " " _
'                                        & " and b." & gsCAT_TLATT_IPN & " = c." & gsCAT_TIM_IPN & "(+) " _
'                                        & " and c." & gsCAT_TIM_IPN & " = d." & gsCAT_TPBS_IPN & "(+) " _
'                                        & " and substr(b." & gsCAT_TLATT_IPN & ", 1, 4) = e." & gsCAT_TPB_PROD_BODY & "(+) " _
'                                        & " and d." & gsCAT_TPBS_DEFAULTS & "(+) = 'Y' " _
'                                        & " and ROWNUM = 1 "
'
'                            Set colRS = moProRawSql.QueryDatabase(sSql)
'                            If Not colRS Is Nothing Then
'                                If colRS.Count > 0 Then
'                                    sVendorCode = colRS.Item(1).Item(1)
'                                    sLotipn = colRS.Item(1).Item(2)
'                                    sCqty = colRS.Item(1).Item(3)
'                                    sRouteId = colRS.Item(1).Item(4)
'                                    sLotOwner = colRS.Item(1).Item(5)
'                                    sDateCode = colRS.Item(1).Item(6)
'                                    sIpnCarGradeFlag = colRS.Item(1).Item(7)
'                                    sIpnProdLine = colRS.Item(1).Item(8)
'                                    sCarrierType = colRS.Item(1).Item(9)
'                                    sExtraRomFlag = colRS.Item(1).Item(10)
'                                    sBrand = colRS.Item(1).Item(11)
'                                    sIcDrawing = colRS.Item(1).Item(12)
'                                    sMarkingSpecNo = colRS.Item(1).Item(13)
'                                    sIpnGreen = colRS.Item(1).Item(14)
'                                    sIpnCsum = colRS.Item(1).Item(15)
'                                    sNFlag = colRS.Item(1).Item(16)
'
'                                    Call modPrint.Prt_FT_FT_LOT_INFO(Trim(vLotID), Trim(sLotipn), Trim(sCqty), _
'                                                                    Trim(sRouteId), Trim(vLotowner), Trim(sDateCode), _
'                                                                    Trim(sIpnCarGradeFlag), Trim(sIpnProdLine), Trim(sCarrierType), _
'                                                                    Trim(sExtraRomFlag), Trim(sBrand), Trim(sIcDrawing), _
'                                                                    Trim(sMarkingSpecNo), Trim(sIpnGreen), Trim(sIpnCsum), _
'                                                                    Trim(sVendorCode), Trim(sNFlag), Me.cboPrintServer.Text, _
'                                                                    moCwMbx, moAppLog)
'                                End If
'                            End If
'Marked by Jack on 2016/07/15 for 2016 MES Phase-5 Project (FT eRunCard Paperless Phase-2 Project) <End>

                            'Added by Jack on 2016/07/15 for 2016 MES Phase-5 Project (FT eRunCard Paperless Phase-2 Project) <Start>
                            'bug fix (for 不是由frmPrintLabel功能列印/統一呼叫列印.)
                            oLot.Refresh
                            Call modPrint.PrintBarcodeLabel(oLot.Id, _
                                                            oLot.CustomAttributes(modConstFwAttr.gsLOT_CUSTOMATTR_WAFER_QTY).Value, _
                                                            oLot.CustomAttributes(modConstFwAttr.gsLOT_CUSTOMATTR_CHIP_QTY).Value, _
                                                            oLot.CustomAttributes(modConstFwAttr.gsLOT_CUSTOMATTR_IPN).Value, _
                                                            oLot.CustomAttributes(modConstFwAttr.gsLOT_CUSTOMATTR_LOT_OWNER).Value, _
                                                            oLot.PlanId, _
                                                            Me.cboPrintServer.Text, _
                                                            Me.cboLabelSpec.Text, _
                                                            moFwOPR.ActiveUser.UserName, moAppLog, moFwWIP, moFwWF, _
                                                            moCwMbx, moProRawSql)
                            'Added by Jack on 2016/07/15 for 2016 MES Phase-5 Project (FT eRunCard Paperless Phase-2 Project) <End>

                            Set colRS = Nothing
                        End If
                    Next
                End With
            'Added by Jack on 2016/06/03 for 2016 MES Phase-5 Project (FT eRunCard Paperless Phase-2 Project) <End>
            'Added by Jack on 2016/06/03 for 2016 MES Phase-5 Project (FT eRunCard Paperless Phase-2 Project) <Start>
            ElseIf Me.cboStage.Text = gsSTAGE_FT And Me.cboLabelSpec.Text = modPrint.gsLABEL_FT_Label_PACK_INFO Then
                Dim sLPI_CarrierType  As String
                Dim sLPI_BoxingSpecNO As String
                Dim sLPI_Brand        As String
                Dim sLPI_PinCount     As String
                Dim sLPI_PackageCode  As String
                Dim sLPI_LabelSpecNo  As String
                
                Dim sLPI_LotNo        As String
                Dim bLPI_IsAutoMode   As Boolean
                
                Dim sLPI_IPN          As String
                
                sLPI_LotNo = ""
                bLPI_IsAutoMode = False
                
                If tabLabelPackInfo.Tab = 0 Then 'Auto mode.
                    sLPI_LotNo = Trim(txtLotID(8).Text) '上面的程式已檢核Lot是否存在!!
                    bLPI_IsAutoMode = True
                    
                    sLPI_CarrierType = txtLPI_CarrierType.Text
                    sLPI_BoxingSpecNO = txtLPI_BoxingSpecNO.Text
                    sLPI_Brand = txtLPI_Brand.Text
                    sLPI_PinCount = txtLPI_PinCount.Text
                    sLPI_PackageCode = txtLPI_PackageCode.Text
                    sLPI_LabelSpecNo = txtLPI_LabelSpecNo.Text
                    sLPI_IPN = txtLPI_IPN.Text 'Added by Jack on 2016/07/04
                    
                    'Added by Jack on 2016/07/15 for 2016 MES Phase-5 Project (FT eRunCard Paperless Phase-2 Project) <Start>
                    'bug fix (for 不是由frmPrintLabel功能列印.)
                    oLot.Refresh
                    Call modPrint.PrintBarcodeLabel(oLot.Id, _
                                                    oLot.CustomAttributes(modConstFwAttr.gsLOT_CUSTOMATTR_WAFER_QTY).Value, _
                                                    oLot.CustomAttributes(modConstFwAttr.gsLOT_CUSTOMATTR_CHIP_QTY).Value, _
                                                    oLot.CustomAttributes(modConstFwAttr.gsLOT_CUSTOMATTR_IPN).Value, _
                                                    oLot.CustomAttributes(modConstFwAttr.gsLOT_CUSTOMATTR_LOT_OWNER).Value, _
                                                    oLot.PlanId, _
                                                    Me.cboPrintServer.Text, _
                                                    Me.cboLabelSpec.Text, _
                                                    moFwOPR.ActiveUser.UserName, moAppLog, moFwWIP, moFwWF, _
                                                    moCwMbx, moProRawSql)
                    'Added by Jack on 2016/07/15 for 2016 MES Phase-5 Project (FT eRunCard Paperless Phase-2 Project) <End>
                Else 'Manual mode.
                    If cboLPI_LabelSpecNO.Text = "" Then
                        UtShowMsgBox ("The LabelSpecNO is not selected !")
                        GoTo ExitHandler
                    End If
                    If cboLPI_CarrierType.Text = "" Then
                        UtShowMsgBox ("The CarrierType is not selected !")
                        GoTo ExitHandler
                    End If
                    If cboLPI_BoxingSpecNO.Text = "" Then
                        UtShowMsgBox ("The BoxingSpecNO is not selected !")
                        GoTo ExitHandler
                    End If
                    If cboLPI_Brand.Text = "" Then
                        UtShowMsgBox ("The Brand is not selected !")
                        GoTo ExitHandler
                    End If
                    If cboLPI_PinCount.Text = "" Then
                        UtShowMsgBox ("The PinCount is not selected !")
                        GoTo ExitHandler
                    End If
                    If cboLPI_PackageCode.Text = "" Then
                        UtShowMsgBox ("The PackageCode is not selected !")
                        GoTo ExitHandler
                    End If
                    
                    sLPI_LabelSpecNo = cboLPI_LabelSpecNO.Text
                    sLPI_CarrierType = cboLPI_CarrierType.Text
                    sLPI_BoxingSpecNO = cboLPI_BoxingSpecNO.Text
                    sLPI_Brand = cboLPI_Brand.Text
                    sLPI_PinCount = cboLPI_PinCount.Text
                    sLPI_PackageCode = cboLPI_PackageCode.Text
                
                    'Added by Jack on 2016/07/26 for 2016 MES Phase-5 Project (FT eRunCard Paperless Phase-2 Project) <Start>
                    'bug fix (for 不是由frmPrintLabel功能列印.)
                    Call modPrint.Prt_FT_Label_PACK_INFO(sLPI_LotNo, sLPI_LabelSpecNo, _
                                                         sLPI_CarrierType, sLPI_BoxingSpecNO, _
                                                         sLPI_Brand, sLPI_PinCount, _
                                                         sLPI_PackageCode, bLPI_IsAutoMode, Me.cboPrintServer.Text, _
                                                         moCwMbx, moAppLog, moProRawSql, sLPI_IPN)
                    'Added by Jack on 2016/07/26 for 2016 MES Phase-5 Project (FT eRunCard Paperless Phase-2 Project) <Start>
                
                End If
                
                'Marked by Jack on 2016/07/26 for 2016 MES Phase-5 Project (FT eRunCard Paperless Phase-2 Project) <Start>
                'bug fix (Auto會多印一次.)
                'Auto & Maual 呼叫共用模組 :
                'sLPI_IPN added by Jack on 2016/07/04.
                'Call modPrint.Prt_FT_Label_PACK_INFO(sLPI_LotNo, sLPI_LabelSpecNo, _
                '                                     sLPI_CarrierType, sLPI_BoxingSpecNO, _
                '                                     sLPI_Brand, sLPI_PinCount, _
                '                                     sLPI_PackageCode, bLPI_IsAutoMode, Me.cboPrintServer.Text, _
                '                                     moCwMbx, moAppLog, moProRawSql, sLPI_IPN)
                'Marked by Jack on 2016/07/26 for 2016 MES Phase-5 Project (FT eRunCard Paperless Phase-2 Project) <Start>
            'Added by Jack on 2016/06/03 for 2016 MES Phase-5 Project (FT eRunCard Paperless Phase-2 Project) <End>
            ElseIf Me.cboStage.Text = gsSTAGE_WS And Me.cboLabelSpec.Text = modPrint.gsLABEL_WSMCD_TO_SFG Then
                Set colRS = moProRawSql.QueryDatabase( _
                    "select A." & gsCAT_TLI_MCDPARENT & _
                        ", C.VALDATA as IPN" & _
                        " from " & gsCAT_TBL_LOT_INFO & " A, FWLOT B, FWLOT_PN2M C " & _
                        " where " & gsCAT_TLI_LOT_ID & " = '" & Trim(Me.txtLotID(miIndex).Text) & "'" & _
                        " and A." & gsCAT_TLI_MCDPARENT & " = B.APPID" & _
                        " and C.KEYDATA = '" & gsLOT_CUSTOMATTR_IPN & "'" & _
                        " and B.SYSID = C.FROMID")
                
                If colRS.Count > 0 Then
                    sParentLotID = colRS.Item(1).Item(1)
                    sParentIPN = colRS.Item(1).Item(2)
                End If
                
                Set colRS = moProRawSql.QueryDatabase( _
                    "select A." & gsCAT_TLI_LOT_ID & "," & _
                           "C.VALDATA as IPN, B.COMPONENTQTY, " & _
                           "D." & gsCAT_TMI_RANK & "," & _
                           "C2.VALDATA as WQTY" & _
                    " from " & gsCAT_TBL_LOT_INFO & " A, " & _
                             " fwlot B, fwlot_pn2m C, fwlot_pn2m C2," & _
                               gsCAT_TBL_MULTICODE_IPN & " D " & _
                    " where A." & gsCAT_TLI_MCDPARENT & " = '" & sParentLotID & "'" & _
                    " and A." & gsCAT_TLI_LOT_ID & " = B.APPID " & _
                    " and C.KEYDATA = '" & gsLOT_CUSTOMATTR_IPN & "'" & _
                    " and B.SYSID = C.FROMID " & _
                    " and C2.KEYDATA = '" & gsLOT_CUSTOMATTR_WAFER_QTY & "'" & _
                    " and B.SYSID = C2.FROMID " & _
                    " and D." & gsCAT_TMI_IPN & " = '" & sParentIPN & "'" & _
                    " and D." & gsCAT_TMI_SUBIPN & " = C.VALDATA ")
                                                      
                If colRS.Count > 0 Then
                    Set colWSMCDInfo = New Collection
                    
                    iCQty = 0
                    iWQty = 0
                    For Each oItem In colRS
                        colWSMCDInfo.Add oItem.Item(1) & ";" & _
                                         oItem.Item(2) & ";" & _
                                         (Asc(oItem.Item(4)) - 64) & ";" & _
                                         oItem.Item(3)
                        iCQty = iCQty + Val(oItem.Item(3))
                        iWQty = iWQty + Val(oItem.Item(5))
                    Next oItem
                    
                    Call modPrint.Prt_WSMCD_TO_SFG(Trim(Me.txtLotID(miIndex).Text), _
                                                   Trim(Me.txtProdNo(miIndex).Text), _
                                                   iWQty, _
                                                   iCQty, _
                                                   Trim(Me.txtUserId(miIndex).Text), _
                                                   sSysTime, _
                                                   Me.cboPrintServer.Text, moCwMbx, moAppLog, _
                                                   colWSMCDInfo)
                
                End If
                
            ElseIf Me.cboStage.Text = gsSTAGE_FT And Me.cboLabelSpec.Text = modPrint.gsLABEL_FT_ETEST_MERGE Then
                Dim sLotInfo As String
                
                sLotInfo = Me.txtLotID(6).Text & "," & Me.txtCqty(6).Text
                With Me.spdChildLot
                    .Row = 1
                    Do While .Row <= .MaxRows
                        .Col = 1
                        sLotInfo = sLotInfo & "," & .Value
                        .Col = 2
                        sLotInfo = sLotInfo & "," & .Value
                        .Row = .Row + 1
                    Loop
                End With
                Call modPrint.Prt_FT_ETEST_MERGE(sLotInfo, Me.cboPrintServer.Text, moCwMbx, moAppLog)
            'Add by Tony Start on 20121001 for Req.JC201200283
            ElseIf Me.cboStage.Text = gsSTAGE_FT And Me.cboLabelSpec.Text = modPrint.gsLABEL_FT_BOX_COUNTING Then
            
                Call modPrint.Prt_FT_BOX_COUNTING_LABEL(oLot, iIndex & "/" & Me.txtPrintQty.Text, Me.cboPrintServer.Text, moCwMbx, moAppLog)
            
            'Add by Tony End on 20121001 for Req.JC201200283
            ElseIf Me.cboStage.Text = gsSTAGE_WS And Me.cboLabelSpec.Text = modPrint.gsLABEL_WS_DGRADE_SUMMARY Then
                
                'Modified by Tony on 20120120 for Req.JC201200010
                'mm/dd/yy hh:mm to yyyy/mm/dd
                Call modPrint.Prt_WS_WS_DGRADE_SUMMARY(Trim(Me.txtLotID(miIndex).Text), _
                                                Trim(Me.txtProdNo(miIndex).Text), _
                                                Trim(Me.txtWQty(miIndex).Text), _
                                                Trim(Me.txtCqty(miIndex).Text), _
                                                Trim(Me.txtUserId(miIndex).Text), _
                                                Format(dteNow, "yyyy/mm/dd"), _
                                                Me.cboPrintServer.Text, _
                                                moCwMbx, moAppLog, moProRawSql)
            
            '2007/08/02 add gsLABEL_WS_MULTILOT_TO_SFG
            ElseIf Me.cboStage.Text = gsSTAGE_WS And Me.cboLabelSpec.Text = modPrint.gsLABEL_WS_MULTILOT_TO_SFG Then
                'Modifid by Jack on 2014/03/24 for JC201400088. <Start>
                'Set colRS = moProRawSql.QueryDatabase( _
                '    "select A." & gsCAT_TLI_MCDPARENT & "," & _
                '          " A." & gsCAT_TLI_MCP_FLAG & "," & _
                '          " C.VALDATA as IPN" & _
                '        " from " & gsCAT_TBL_LOT_INFO & " A, FWLOT B, FWLOT_PN2M C " & _
                '        " where " & gsCAT_TLI_LOT_ID & " = '" & Trim(Me.txtLotID(miIndex).Text) & "'" & _
                '        " and A." & gsCAT_TLI_MCDPARENT & " = B.APPID" & _
                '        " and C.KEYDATA = '" & gsLOT_CUSTOMATTR_IPN & "'" & _
                '        " and B.SYSID = C.FROMID")
                Set colRS = moProRawSql.QueryDatabase( _
                    "select A." & gsCAT_TLI_MCDPARENT & "," & _
                          " A." & gsCAT_TLI_MCP_FLAG & " " & _
                        " from " & gsCAT_TBL_LOT_INFO & " A " & _
                        " where " & gsCAT_TLI_LOT_ID & " = '" & Trim(Me.txtLotID(miIndex).Text) & "'")
                'Modifid by Jack on 2014/03/24 for JC201400088. <End>
                
                If colRS.Count > 0 Then
                    sParentLotID = colRS.Item(1).Item(gsCAT_TLI_MCDPARENT)
                    sMcpFlag = colRS.Item(1).Item(gsCAT_TLI_MCP_FLAG)
                    'sParentIPN = colRS.Item(1).Item("IPN")'Marked by Jack on 2014/03/24 for JC201400088.
                End If
                
                'Added by Jack on 2014/03/24 for JC201400088. <Start>
                If sMcpFlag = "F" Or sMcpFlag = "D" Then
                    Set colRS = moProRawSql.QueryDatabase( _
                    "select A." & gsCAT_TLI_MCDPARENT & "," & _
                          " A." & gsCAT_TLI_MCP_FLAG & "," & _
                          " C.VALDATA as IPN" & _
                        " from " & gsCAT_TBL_LOT_INFO & " A, FWLOT B, FWLOT_PN2M C " & _
                        " where " & gsCAT_TLI_LOT_ID & " = '" & Trim(Me.txtLotID(miIndex).Text) & "'" & _
                        " and A." & gsCAT_TLI_MCDPARENT & " = B.APPID" & _
                        " and C.KEYDATA = '" & gsLOT_CUSTOMATTR_IPN & "'" & _
                        " and B.SYSID = C.FROMID")
                    If colRS.Count > 0 Then
                       sParentIPN = colRS.Item(1).Item("IPN")
                    End If
                End If
                'Added by Jack on 2014/03/24 for JC201400088. <End>
                
                'Marked by Jack on 2014/03/24 for JC201400088. <Start>
                '修改Label列印卡關
                '目前：不論McpFlag內容為何，當MCDParent為空值，則出現錯誤訊息，不允許列印。
                '變更：判斷McpFlag，決定列印Label內容。
                '      McpFlag='F'  依McdParent 串F13S 相關資料。
                '      McpFlag='D'  依McdParent 串MultiCode IPN 相關資料。
                '      McpFlag not ('F','D')  不需判斷McdParent，直接列印。
'                If sParentLotID = "" Then
'                    'Modify by Sam start on 2008/06/12 for ReqNO:M200806023
'                    'sMsg = "TBL_LOT_INFO.MCDPARENT Is Null" & vbCrLf & vbCrLf & _
'                    '       "此Label無法列印, 請用WS_SMALL_LABEL取代" & vbCrLf & vbCrLf & _
'                    '       "This label can not print, please replaced it with WS_SMALL_LABEL."
'                    'Modified by Jack 2009/09/08 [JC900117]
'                    sMsg = "TBL_LOT_INFO.MCDPARENT Is Null" & vbCrLf & vbCrLf & _
'                           "此Label無法列印, 請通知SFG重設MCDPARENT值." & vbCrLf & vbCrLf & _
'                           "This label can not print, please notice SFG to reset MCDPARENT."
'
'                    UtShowMsgBox sMsg
''                    UtShowMsgBox "TBL_LOT_INFO.MCDPARENT Is Null"
'                    'Modify by Sam end on 2008/06/12 for ReqNO:M200806023
'                    GoTo ExitHandler
'
'                End If
                'Marked by Jack on 2014/03/24 for JC201400088. <End>
                
                Set colWSMCDInfo = New Collection
                iCQty = 0
                iWQty = 0
                'Modify by Sam on 20140429 for ReqNo:JC201400140 ,增加C的判斷
                'If sMcpFlag = "F" Then 'F13S prod
                If sMcpFlag = "F" Or sMcpFlag = "C" Then 'F13S prod
                    'ADD B.PROCESSINGSTATUS <> 'Terminated' CRITERIA by Tony on 20080117 Req.M200801028
                    Set colRS = moProRawSql.QueryDatabase( _
                        "select A." & gsCAT_TLI_LOT_ID & "," & _
                               "C.VALDATA as IPN, B.COMPONENTQTY, " & _
                               "C2.VALDATA as WQTY" & _
                        " from " & gsCAT_TBL_LOT_INFO & " A, fwlot B, fwlot_pn2m C, fwlot_pn2m C2" & _
                        " where A." & gsCAT_TLI_MCDPARENT & " = '" & sParentLotID & "'" & _
                        " and A." & gsCAT_TLI_LOT_ID & " = B.APPID " & _
                        " and C.KEYDATA = '" & gsLOT_CUSTOMATTR_IPN & "'" & _
                        " and B.SYSID = C.FROMID " & _
                        " and C2.KEYDATA = '" & gsLOT_CUSTOMATTR_WAFER_QTY & "'" & _
                        " and B.SYSID = C2.FROMID " & _
                        " and B.PROCESSINGSTATUS <> '" & gsLOT_STATUS_TERMINATED & "'")
                    
                    For Each oItem In colRS
                        If oItem.Item(1) = sParentLotID Then
                            colWSMCDInfo.Add oItem.Item(1) & ";" & _
                                             oItem.Item(2) & ";" & _
                                             "BIN1;" & _
                                             oItem.Item(3)
                        Else
                            'Add by Sam Start on 20140429 for ReqNO:JC201400140
                            If sMcpFlag = "C" Then
                                colWSMCDInfo.Add oItem.Item(1) & ";" & _
                                                 oItem.Item(2) & ";" & _
                                                 "BIN7;" & _
                                                 oItem.Item(3)
                            Else
                            'Add by Sam End on 20140429 for ReqNO:JC201400140
                                colWSMCDInfo.Add oItem.Item(1) & ";" & _
                                                 oItem.Item(2) & ";" & _
                                                 "BIN3;" & _
                                                 oItem.Item(3)
                            End If
                        End If
                        iCQty = iCQty + Val(oItem.Item(3))
                        iWQty = iWQty + Val(oItem.Item("WQTY"))
                    Next oItem
                'Modified by Jack on 2014/03/24 for JC201400088. <Start>
                'Else '照原 mcd label
                'a. 原本 : Else '照原 mcd label
                '   修改 : ElseIf sMcpFlag = "D" Then
                'b. 原本 : where A." & gsCAT_TLI_MCDPARENT & " = '" & sParentLotID & "'"
                '   修改 : where A." & gsCAT_TLI_LOT_ID & " = '" & sParentLotID & "' "
                ElseIf sMcpFlag = "D" Then
                    'ADD B.PROCESSINGSTATUS <> 'Terminated' CRITERIA by Tony on 20080117 Req.M200801028
                    'Modified by Jack on 2014/03/24 for JC201400088. <Start>
                    Set colRS = moProRawSql.QueryDatabase( _
                        "select A." & gsCAT_TLI_LOT_ID & "," & _
                               "C.VALDATA as IPN, B.COMPONENTQTY, " & _
                               "D." & gsCAT_TMI_RANK & "," & _
                               "C2.VALDATA as WQTY" & _
                        " from " & gsCAT_TBL_LOT_INFO & " A, " & _
                                 " fwlot B, fwlot_pn2m C, fwlot_pn2m C2," & _
                                   gsCAT_TBL_MULTICODE_IPN & " D " & _
                        " where A." & gsCAT_TLI_LOT_ID & " = '" & sParentLotID & "'" & _
                        " and A." & gsCAT_TLI_LOT_ID & " = B.APPID " & _
                        " and C.KEYDATA = '" & gsLOT_CUSTOMATTR_IPN & "'" & _
                        " and B.SYSID = C.FROMID " & _
                        " and C2.KEYDATA = '" & gsLOT_CUSTOMATTR_WAFER_QTY & "'" & _
                        " and B.SYSID = C2.FROMID " & _
                        " and D." & gsCAT_TMI_IPN & " = '" & sParentIPN & "'" & _
                        " and D." & gsCAT_TMI_SUBIPN & " = C.VALDATA " & _
                        " and B.PROCESSINGSTATUS <> '" & gsLOT_STATUS_TERMINATED & "'")
                'Modified by Jack on 2014/03/24 for JC201400088. <End>
                        
                    For Each oItem In colRS
                        colWSMCDInfo.Add oItem.Item(1) & ";" & _
                                         oItem.Item(2) & ";" & _
                                         (Asc(oItem.Item(4)) - 64) & ";" & _
                                         oItem.Item(3)
                        iCQty = iCQty + Val(oItem.Item(3))
                        iWQty = iWQty + Val(oItem.Item(5))
                    Next oItem
                'Added by Jack on 2014/03/24 for JC201400088. <Start>
                'McpFlag not ('F','D')  不需判斷McdParent，直接列印。
                Else
                    Set colRS = moProRawSql.QueryDatabase( _
                        "select A." & gsCAT_TLI_LOT_ID & "," & _
                               "C.VALDATA as IPN, B.COMPONENTQTY, " & _
                               "C2.VALDATA as WQTY" & _
                        " from " & gsCAT_TBL_LOT_INFO & " A, fwlot B, fwlot_pn2m C, fwlot_pn2m C2" & _
                        " where A." & gsCAT_TLI_LOT_ID & " = '" & Trim(Me.txtLotID(miIndex).Text) & "'" & _
                        " and A." & gsCAT_TLI_LOT_ID & " = B.APPID " & _
                        " and C.KEYDATA = '" & gsLOT_CUSTOMATTR_IPN & "'" & _
                        " and B.SYSID = C.FROMID " & _
                        " and C2.KEYDATA = '" & gsLOT_CUSTOMATTR_WAFER_QTY & "'" & _
                        " and B.SYSID = C2.FROMID " & _
                        " and B.PROCESSINGSTATUS <> '" & gsLOT_STATUS_TERMINATED & "'")
                    
                    For Each oItem In colRS
                        colWSMCDInfo.Add oItem.Item(1) & ";" & _
                                         oItem.Item(2) & ";" & _
                                         "BIN1;" & _
                                         oItem.Item(3)
                        iCQty = iCQty + Val(oItem.Item(3))
                        iWQty = iWQty + Val(oItem.Item("WQTY"))
                    Next oItem
                'Added by Jack on 2014/03/24 for JC201400088. <End>
                End If
                
                If colRS.Count > 0 Then
                    'Modified by Tony on 20120120 for Req.JC201200010
                    'sSysTime to Format(dteNow, "yyyy/mm/dd")
                    Call modPrint.Prt_WSMULTILOT_TO_SFG(Trim(Me.txtUserId(miIndex).Text), _
                                                   Format(dteNow, "yyyy/mm/dd"), _
                                                   CStr(iWQty), _
                                                   CStr(iCQty), _
                                                   Me.cboPrintServer.Text, moCwMbx, moAppLog, _
                                                   colWSMCDInfo)
                End If
            ElseIf Me.cboStage.Text = gsSTAGE_WS And Me.cboLabelSpec.Text = modPrint.gsLABEL_RENESAS_SHIPPING Then 'add by Nelson on 2008/06/03 for ReqNo:M200806003
                Dim arrWaferId(25) As String
                Dim arrWaferIdChipQty(25) As String
                Dim iWIdIndex As Integer
                
                Set colRS = moProRawSql.QueryDatabase("select b.keydata waferid,d.valdata chipqty " & _
                                                      " from fwlot a," & _
                                                            "fwlot_n2m b," & _
                                                            "fwcomponent c," & _
                                                            "fwcomponent_pn2m d " & _
                                                      "where a.sysid=b.fromid and " & _
                                                            "b.toid=c.sysid and " & _
                                                            "c.sysid=d.fromid and " & _
                                                            "a.appid='" & Me.txtLotID.Item(miIndex).Text & "' order by b.keydata")
                
                If colRS.Count > 0 Then
                    For Each oItem In colRS
                        iWIdIndex = oItem.Item(1)
                        arrWaferId(iWIdIndex) = Left(Me.txtLotID(miIndex).Text, 6) & "-" & oItem.Item(1)
                        arrWaferIdChipQty(iWIdIndex) = oItem.Item(2)
                    Next oItem
                End If
                
                Call modPrint.Prt_WS_RENESAS_SHIPPING(Me.txtLotID(miIndex).Text, _
                                                      Me.txtProdNo(miIndex).Text, _
                                                      Me.txtWQty(miIndex).Text, _
                                                      Me.txtCqty(miIndex).Text, _
                                                      Me.txtPartNo(miIndex).Text, _
                                                      arrWaferId, _
                                                      arrWaferIdChipQty, _
                                                      Format(Now, "yyyy/mm/dd"), _
                                                      Me.txtUserId(miIndex).Text, _
                                                      Me.cboPrintServer.Text, _
                                                      moCwMbx, moAppLog)
                
            'Add by Sam on 20110823 for ReqNo:JC201100218
            ElseIf Me.cboStage.Text = gsSTAGE_WS And Me.cboLabelSpec.Text = modPrint.gsLABEL_WS_ENG_LOC_LABEL Then
                With Me.spdWsEngLoc
                    For iRowIndex = 1 To .MaxRows
                        .Row = iRowIndex
                            .GetText miSPDWSENGLOC_LOT_ID, iRowIndex, vLotID
                            If Trim(vLotID) <> "" Then
                                    Call modPrint.Prt_WS_WS_ENG_LOC_LABEL(Trim(CStr(vLotID)), _
                                                                          Me.cboPrintServer.Text, _
                                                                        moCwMbx, moAppLog, moProRawSql)
                            End If
                        Next
                End With
            'Add gsLABEL_FT_TR_LABEL by HouYu for 202000413 on 20220422
            ElseIf Me.cboStage.Text = gsSTAGE_FT And Me.cboLabelSpec.Text = modPrint.gsLABEL_FT_TR_LABEL Then
                If Trim(Me.txtLotID(miIndex).Text) = "" Then
                    UtShowMsgBox "Please input Lot Id"
                    GoTo ExitHandler
                End If
                If Trim(Me.txt_FT_TR_IPN.Text) = "" Then
                    UtShowMsgBox "Please input IPN"
                    GoTo ExitHandler
                End If
                If Trim(Me.txt_FT_TR_EqId.Text) = "" Then
                    UtShowMsgBox "Please input EQ Id"
                    GoTo ExitHandler
                End If
                If Trim(Me.txt_FT_TR_QtyReel.Text) = 0 Then
                        UtShowMsgBox "Please input Qty Reel"
                        GoTo ExitHandler
                End If
                If Trim(Me.cbo_FT_TR_ReelId.Text) = "" Then
                    UtShowMsgBox "Please select  Reel Id"
                    GoTo ExitHandler
                Else
                    If Trim(Me.cbo_FT_TR_ReelId.Text) = "All" Then
                        For i = 1 To Me.cbo_FT_TR_ReelId.ListCount - 1
                            Call modPrint.Prt_FT_TR_LABEL(Trim(Me.txtLotID(miIndex).Text), _
                                                    Me.txt_FT_TR_IPN.Text, _
                                                    Me.txt_FT_TR_QtyReel.Text, _
                                                    Me.txt_FT_TR_EqId.Text, _
                                                    i, _
                                                    Me.cboPrintServer.Text, _
                                                    moCwMbx, moAppLog)
                        Next i
                    Else
                        Call modPrint.Prt_FT_TR_LABEL(Trim(Me.txtLotID(miIndex).Text), _
                                                    Me.txt_FT_TR_IPN.Text, _
                                                    Me.txt_FT_TR_QtyReel.Text, _
                                                    Me.txt_FT_TR_EqId.Text, _
                                                    Me.cbo_FT_TR_ReelId, _
                                                    Me.cboPrintServer.Text, _
                                                    moCwMbx, moAppLog)
                    End If
                End If
            'Add gsLABEL_FT_TR_LABEL by HouYu for 202000413 on 20220422
            End If
        Next iIdx
    Next iIndex
'----
' Done
'----
' mark by Nelson start on 2004/2/11 for ReqNo:M200401006
    'Me.Hide
'    Me.cboStage.ListIndex = 0
'    Me.cboLabelSpec.ListIndex = 0
'    Me.cboPrintServer.ListIndex = 0
'    If Me.txtBoxQty.Enabled = True Then
'        Me.txtBoxQty.Text = ""
'    End If
'
'    Me.txtPrintQty.Text = ""
    
    If fraLabelForm(0).Visible = True Then
        Me.txtLotID(0).Text = ""
        Me.txtLotID(0).Locked = False
        Me.txtLotID(0).BackColor = vbWhite
        Me.txtProdNo(0).Text = ""
        Me.txtQty(0).Text = ""
'        Me.txtUserId(0).Text = ""
        Me.txtLotID(0).SetFocus
    End If
       
    If fraLabelForm(1).Visible = True Then
        Me.txtLotID(1).Text = ""
        Me.txtLotID(1).Locked = False
        Me.txtLotID(1).BackColor = vbWhite
        Me.txtProdNo(1).Text = ""
        Me.txtWQty(1).Text = ""
        Me.txtCqty(1).Text = ""
        Me.txtPartNo(1).Text = ""
        Me.txtLotID(1).SetFocus
    End If
    
    If fraLabelForm(2).Visible = True Then
        Me.txtLotID(2).Text = ""
        Me.txtLotID(2).Locked = False
        Me.txtLotID(2).BackColor = vbWhite
        Me.txtProdNo(2).Text = ""
        Me.txtWQty(2).Text = ""
        Me.txtCqty(2).Text = ""
'        Me.txtUserId(2).Text = ""
        Me.txtLotID(2).SetFocus
    End If
    
    If fraLabelForm(3).Visible = True Then
        Me.txtLotID(3).Text = ""
        Me.txtLotID(3).Locked = False
        Me.txtLotID(3).BackColor = vbWhite
        Me.txtProdNo(3).Text = ""
        Me.txtWQty(3).Text = ""
        Me.txtCqty(3).Text = ""
        Me.txtLotOwner(3).Text = ""
        Me.txtSpeed(3).Text = ""
        Me.txtLotID(3).SetFocus
    End If
    
    If fraLabelForm(4).Visible = True Then
        Me.txtLotID(4).Text = ""
        Me.txtLotID(4).Locked = False
        Me.txtLotID(4).BackColor = vbWhite
        Me.txtProdNo(4).Text = ""
        Me.txtWQty(4).Text = ""
        Me.txtCqty(4).Text = ""
        Me.txtLotOwner(4).Text = ""
        Me.txtRouteId(4).Text = ""
        Me.txtLotID(4).SetFocus
    End If
    
    If fraLabelForm(5).Visible = True Then
        Me.txtLotID(5).Text = ""
        Me.txtLotID(5).Locked = False
        Me.txtLotID(5).BackColor = vbWhite
        Me.txtProdNo(5).Text = ""
        Me.txtWQty(5).Text = ""
        Me.txtCqty(5).Text = ""
        
    End If
    
    If fraLabelForm(6).Visible = True Then
        Me.txtLotID(6).Text = ""
        Me.txtLotID(6).Locked = False
        Me.txtLotID(6).BackColor = vbWhite
        Me.txtCqty(6).Text = ""
        Me.spdChildLot.MaxRows = 0
    End If
    
    If fraLabelForm(7).Visible = True Then
        Me.txtLotID(7).Text = ""
        Me.txtLotID(7).Locked = False
        Me.txtLotID(7).BackColor = vbWhite
        Me.txtProdNo(7).Text = ""
        Me.txtWQty(7).Text = ""
        Me.txtCqty(7).Text = ""
        Me.txtPartNo(7).Text = ""
        'Me.txtUserId(7).Text = ""
        Me.txtLotID(7).SetFocus
    End If
    
    'Add gsLABEL_FT_TR_LABEL by HouYu for 202000413 on 20220422
    If Me.fraLabelForm(9).Visible = True Then
        Me.txtLotID(9).Text = ""
        Me.txtLotID(9).Locked = False
        Me.txtLotID(9).BackColor = vbWhite
        Me.txt_FT_TR_IPN.Text = ""
        Me.txt_FT_TR_EqId.Text = ""
        Me.cbo_FT_TR_ReelId.Clear
        Me.txt_FT_TR_QtyReel.Text = ""
    End If
    'Add gsLABEL_FT_TR_LABEL by HouYu for 202000413 on 20220422
    
    'Added by Jack on 2010/04/12 for JC201000060. <Start>
    '列印完成, 將資料清空.
    If spdLotList.Visible = True Then
        With spdLotList
            .MaxRows = 0
            .MaxRows = 1
            Call .SetCellBorder(0, 0, .MaxCols, .MaxRows, 15, -1, CellBorderStyleSolid)
        End With
    End If
    'Added by Jack on 2010/04/12 for JC201000060. <End>
    
'    fraLabelForm(0).Visible = True
' mark by Nelson end on 2004/2/11 for ReqNo:M200401006
ExitHandler:
    ' NOTE 1:
    ' MUST CALL GetErrInfo() here first before another action
    Call GetErrInfo(msMODULE_ID, sProcID, typErrInfo, Erl)
    Call LogProcOut(msMODULE_ID, sProcID, typErrInfo, moAppLog)
    Screen.MousePointer = vbDefault
    ' Cleaning up
    miResult = vbOK

    'add by Sam start on 20080910 for ReqNo:M200807057
    If msLabelSpec <> "" Then
        Me.Hide
    End If
    'add by Sam end on 20080910 for ReqNo:M200807057
    
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
        cmdOk.Enabled = True
        miResult = vbCancel
        On Error GoTo ExitHandler:
        Call HandleError(False, typErrInfo, , moAppLog, True)
    End If
End Sub
'================================================================================
' function: GetLabelSpec()
'--------------------------------------------------------------------------------
' Description:  <Type your function description here...>
'--------------------------------------------------------------------------------
' Author:       Ivan, CIT 2002-02-27
'--------------------------------------------------------------------------------
' Parameters:
'................................................................................
' ARGUMENTS
'   moapplog            (I) [Optional] Valid reference to a Object object
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
Private Function GetLabelSpec() As Boolean
On Error GoTo ExitHandler:
Dim sProcID As String
Dim typErrInfo As tErrInfo
Dim sTable As String
Dim sColumn As String
Dim sWhere As String
Dim sOther  As String
Dim colRaws As Collection
Dim oRaws As FwStrings
'----
' Init
'----
    sProcID = "GetLabelSpec"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog)
    GetLabelSpec = True
    Me.cboLabelSpec.Clear
'----
' Condition Chking
'----


'----
' Action
'----
    sTable = modConstCAT.gsCAT_TBL_LABEL_SPEC

    sColumn = modConstCAT.gsCAT_TLS_LABEL_SPECNO
    
    sWhere = modConstCAT.gsCAT_TLS_STAGE & "='" & Me.cboStage.Text & "' and " & _
              modConstCAT.gsCAT_TLS_DELETE_FLAG & "='N'"

    sOther = "order by " & modConstCAT.gsCAT_TLS_LABEL_SPECNO
             
    Set colRaws = SelectRawSqlTxn(moAppLog, moFwWIP, moFwWF, moCwMbx, "CAT", sTable, sColumn, sWhere, sOther)
    If colRaws Is Nothing Then
        Call RaiseError(glERR_INVALIDOBJECT, _
                    FormatErrorText(gsETX_INVALIDOBJECT, "Collection"))
    End If
    
    Me.cboLabelSpec.AddItem ""
    For Each oRaws In colRaws
        Me.cboLabelSpec.AddItem oRaws.Item(1)
    Next oRaws
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
        GetLabelSpec = False
        Call HandleError(True, typErrInfo, , moAppLog)
    End If
End Function

' function: GetLabelData()
'--------------------------------------------------------------------------------
' Description:  <Type your function description here...>
'--------------------------------------------------------------------------------
' Author:       Sam, CIT 2008-09-10
'--------------------------------------------------------------------------------
' Parameters:
'................................................................................
' ARGUMENTS
'   moapplog            (I) [Optional] Valid reference to a Object object
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
Private Function GetLabelData() As Boolean
On Error GoTo ExitHandler:
Dim sProcID As String
Dim typErrInfo As tErrInfo
Dim sTable As String
Dim sColumn As String
Dim sWhere As String
Dim sOther  As String
Dim colRaws As Collection
Dim oRaws As FwStrings
'----
' Init
'----
    sProcID = "GetLabelData"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog)
    GetLabelData = True
    txtBoxQty.Text = ""
    txtPrintQty = ""
    
    cboPrintServer.Text = ""
    
'----
' Condition Chking
'----

'----
' Action
'----
    sTable = modConstCAT.gsCAT_TBL_LABEL_SPEC

    sColumn = modConstCAT.gsCAT_TLS_BOX_QTY & "," & _
              modConstCAT.gsCAT_TLS_PRINT_QTY & "," & _
              modConstCAT.gsCAT_TLS_SERVER_NAME & "," & _
              modConstCAT.gsCAT_TLS_LABELW_SPECNO
    
    sWhere = modConstCAT.gsCAT_TLS_STAGE & "='" & Me.cboStage.Text & "' and " & _
             modConstCAT.gsCAT_TLS_LABEL_SPECNO & "='" & Me.cboLabelSpec.Text & "' and " & _
             modConstCAT.gsCAT_TLS_DELETE_FLAG & "='N'"
            
    Set colRaws = SelectRawSqlTxn(moAppLog, moFwWIP, moFwWF, moCwMbx, "CAT", sTable, sColumn, sWhere, sOther)
    If colRaws Is Nothing Then
        Call RaiseError(glERR_INVALIDOBJECT, _
                    FormatErrorText(gsETX_INVALIDOBJECT, "Collection"))
    End If
    
    If colRaws.Count > 0 Then
        For Each oRaws In colRaws
            txtBoxQty = oRaws.Item(1)
            
            If oRaws.Item(2) <> "" Then
                txtPrintQty = oRaws.Item(2)
            Else
                txtPrintQty = "1"
            End If
            
            Me.cboPrintServer.Text = oRaws.Item(3)
        Next oRaws
        GetLabelData = False
    Else
        txtPrintQty = "1"
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
        GetLabelData = True
        Call HandleError(True, typErrInfo, , moAppLog)
    End If
End Function

Private Sub printFtBinCardLabel(bPrintPass As Boolean)
'This method added by Jack on 2011/04/18 for JC201100104 Print FT_Bin_Card.
On Error GoTo ExitHandler:
Dim sProcID As String
Dim typErrInfo As tErrInfo
Dim bIsPartial As Boolean
Dim colCompIds As FwStrings
Dim oLot       As FwLot
Dim oComp      As FwComponent
Dim sSysTime   As String
Dim dteNow     As Date
Dim iIndex     As Integer

Dim colRS           As Collection
Dim colWSMCDInfo    As Collection
Dim sParentLotID    As String
Dim sParentIPN      As String
Dim sMcpFlag        As String
Dim oItem           As Object
Dim iCQty           As Long
Dim iWQty           As Long

Dim sSQL    As String

Dim sMsg    As String

Dim sBizType As String

Dim bLotSelect  As Boolean
Dim iRowIndex As Integer
Dim vLotID    As Variant
Dim vProdNo   As Variant
Dim vStepName As Variant
Dim vWqty     As Variant
Dim vCqty     As Variant
Dim vLotOwner As Variant
Dim vSpeed    As Variant
Dim vRoute    As Variant

Dim iSelectedLot As Integer
'----
' Init
'----
    sProcID = "cmdOk_Click"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog)
    Screen.MousePointer = vbHourglass
    Set colCompIds = New FwStrings
'----
' Condition Checking
'----
    If Len(Me.cboPrintServer.Text) = 0 Then
        Call UtShowMsgBox("Please select Printer Server!")
        GoTo ExitHandler
    End If
    If bPrintPass Then
        If Len(Trim(Me.txtPrintQty.Text)) = 0 Then
            Call UtShowMsgBox("請輸入列印份數!")
            GoTo ExitHandler
        End If
        If Val(Me.txtPrintQty.Text) = "0" Then
            Call UtShowMsgBox("列印份數不能為零!")
            GoTo ExitHandler
        End If
    Else
        If Len(Trim(Me.txtPrintFailQty.Text)) = 0 Then
            Call UtShowMsgBox("請輸入列印份數/Fail Sheet!")
            GoTo ExitHandler
        End If
        If Val(Me.txtPrintFailQty.Text) = "0" Then
            Call UtShowMsgBox("列印份數/Fail Sheet不能為零!")
            GoTo ExitHandler
        End If
    End If
    If (moFwWIP Is Nothing) Then
        Call RaiseError(glERR_INVALIDOBJECT, _
                        FormatErrorText(gsETX_INVALIDOBJECT, "FwWIP"))
    End If
    
    iSelectedLot = 0
    If (msLabelSpec = "") And (Me.cboLabelSpec.Text = modPrint.gsLABEL_FT_BIN_CARD_LABEL) Then
        bLotSelect = False 'Check lot selection or not
        With spdLotList
              .Col = mispdLotList_CHECK
              For iIndex = 1 To .MaxRows
                  .Row = iIndex
                  If .Value = 1 Then
                      iSelectedLot = iSelectedLot + 1
                      bLotSelect = True
                  End If
              Next
        End With
        If bLotSelect = False Then
            Call UtShowMsgBox("Please Select Lot(s)!!")
            GoTo ExitHandler
        ElseIf iSelectedLot > 1 Then '一次只能列印一筆.
            Call UtShowMsgBox("Only can print one Lot(s) once!!")
            GoTo ExitHandler
        End If
    Else '原本的單筆check rule.
        Set oLot = moFwWIP.LotById(Trim(Me.txtLotID(miIndex).Text))
        If oLot Is Nothing Then
            Call UtShowMsgBox("The Lot is not exist!")
            GoTo ExitHandler
        End If
    End If
                
    If Me.cboLabelSpec.Text = "" Then
        UtShowMsgBox "Please select one label type"
        GoTo ExitHandler
    End If
    
'----
' Action
'----
    'Get System Time
    sSysTime = GetSystemTime(moAppLog, moFwWIP, moFwWF, moCwMbx)
    
    dteNow = ConvertToDate(sSysTime, moAppLog)
      
    Dim iIdx As Integer, iPrintTimes As Integer
    Dim sPrintQty As String
    
    If bPrintPass Then
        iPrintTimes = Val(Me.txtPrintQty.Text)
    Else
        iPrintTimes = Val(Me.txtPrintFailQty.Text)
    End If
    
    bIsPartial = False
    
    For iIndex = 1 To iPrintTimes
        Dim sIPN      As String
        Dim sStepName As String
        Dim sStepID As String
        
        Dim sCustomer As String
        Dim sCsum As String
        Dim sGreen As String
        Dim sGPType As String
        Dim sCargradeFlag As String
        
        sCustomer = ""
        sCsum = ""
        sGreen = ""
        sIPN = ""
        sGPType = ""
        'msLabelSpec added by Jack on 2010/04/12 for JC201000060.
        '有值 : 由frmAdltProd執行此程式, 只能單筆列印, 只可列印RENESAS_SHIPPING Label.(程式碼照舊)
        '無值: 由F1執行此程式, "WS_SMALL_LABEL"/"FT_SMALL_LABEL"都能多筆列印.
        'StepID added by Jack on 2011/04/25 requested by Daniel.
        If msLabelSpec <> "" Then
            sSQL = "select a." & gsCAT_TLATT_IPN & ", " & _
                          "a." & gsCAT_TLATT_STEPNAME & ", " & _
                          "a." & gsCAT_TLATT_STEPID & " " & _
                     "from " & gsCAT_TBL_LOT_ATTRIBUTE & " a " & _
                    "where a." & gsCAT_TLATT_LOTID & " = '" & Trim(Me.txtLotID(miIndex).Text) & "'"
            
            Set colRS = moProRawSql.QueryDatabase(sSQL)
            If Not colRS Is Nothing Then
                If colRS.Count > 0 Then
                    sIPN = colRS.Item(1).Item(1)
                    sStepName = colRS.Item(1).Item(2)
                    sStepID = colRS.Item(1).Item(3)
                End If
            End If
            
            Set colRS = Nothing
            
            Call modPrint.Prt_FT_FT_BIN_CARD_LABEL(Trim(Me.txtLotID(miIndex).Text), _
                                                Trim(sIPN), _
                                                Trim(sStepName), _
                                                Trim(sStepID), _
                                                Me.cboPrintServer.Text, bPrintPass, _
                                                moCwMbx, moAppLog, sCargradeFlag)
        Else
            With Me.spdLotList
                For iRowIndex = 1 To .MaxRows
                    .Col = mispdLotList_CHECK
                    .Row = iRowIndex
                    If .Text = 1 Then '打勾才要列印.
                        .GetText miSPDLOTLIST_LOT_ID, iRowIndex, vLotID
                        .GetText miSPDLOTLIST_PROD_NO, iRowIndex, vProdNo
                        .GetText mispdLotList_STEP, iRowIndex, vStepName
                        .GetText mispdLotList_WQTY, iRowIndex, vWqty
                        .GetText mispdLotList_CQTY, iRowIndex, vCqty
                        .GetText miSPDLOTLIST_LOT_OWNER, iRowIndex, vLotOwner
                        .GetText miSPDLOTLIST_SPEED, iRowIndex, vSpeed
                        .GetText mispdLotList_ROUTE, iRowIndex, vRoute
                                
                        'StepID added by Jack on 2011/04/25 requested by Daniel.
                        sSQL = "select a." & gsCAT_TLATT_IPN & ", " & _
                                      "a." & gsCAT_TLATT_STEPNAME & ", " & _
                                      "a." & gsCAT_TLATT_STEPID & " " & _
                                 "from " & gsCAT_TBL_LOT_ATTRIBUTE & " a " & _
                                "where a." & gsCAT_TLATT_LOTID & " = '" & Trim(vLotID) & "'"
                        Set colRS = moProRawSql.QueryDatabase(sSQL)
                        If Not colRS Is Nothing Then
                            If colRS.Count > 0 Then
                                sIPN = colRS.Item(1).Item(1)
                                sStepName = colRS.Item(1).Item(2)
                                sStepID = colRS.Item(1).Item(3)
                            End If
                        End If
                        Set colRS = Nothing
                                               
                        Call modPrint.Prt_FT_FT_BIN_CARD_LABEL(Trim(vLotID), _
                                                            Trim(sIPN), _
                                                            Trim(sStepName), _
                                                            Trim(sStepID), _
                                                            Me.cboPrintServer.Text, bPrintPass, _
                                                            moCwMbx, moAppLog, sCargradeFlag)
                End If
                Next
            End With
        End If
    Next iIndex
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
    miResult = vbOK
    
'列印FT BIN CARD Label必需經由Cancel關閉Form.
'    'add by Sam start on 20080910 for ReqNo:M200807057
'    If msLabelSpec <> "" Then
'        Me.Hide
'    End If
'    'add by Sam end on 20080910 for ReqNo:M200807057
    
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
        cmdOk.Enabled = True
        miResult = vbCancel
        On Error GoTo ExitHandler:
        Call HandleError(False, typErrInfo, , moAppLog, True)
    End If
End Sub
'Added by Jack on 2011/04/18 for JC201100104.
Private Sub cmdPrintFail_Click()
    Call printFtBinCardLabel(False)
End Sub

'Added by Jack on 2011/04/18 for JC201100104.
Private Sub cmdPrintPass_Click()
    Call printFtBinCardLabel(True)
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
' Author:       Ivan, CIT 2001/01/14
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
Dim colStage   As Collection
Dim vStage     As Variant
Dim oUser      As FwUser
Dim iIndex     As Integer
Dim oLot       As FwLot
Dim sCurStage  As String
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
    'Init
    cboStage.Clear
    cboLabelSpec.Clear
    Me.txtPrintQty.Text = "1"
    Call GetPrinterServers(moAppLog, moFwWIP, moFwWF, moCwMbx, Me.cboPrintServer)
    
    'Added by Jack on 2011/04/15 <Start>
    Me.txtPrintFailQty.Text = "1"
    Me.txtPrintFailQty.Visible = False
    lblPrintFailQty.Visible = False
    Me.cmdPrintPass.Visible = False
    Me.cmdPrintFail.Visible = False
    
    'Added by Jack on 2011/04/15 <End>
    
    '******
    '* Get login user
    '******
        Set oUser = moFwOPR.ActiveUser
    '******
    
    '******
    '* Get All Stage
    '******
        Set colStage = modCustom.GetAllStage(moAppLog, _
                                            moFwWIP, _
                                            moFwWF, _
                                            moCwMbx)
        Me.cboStage.AddItem ""
        For Each vStage In colStage
            Me.cboStage.AddItem vStage
        Next vStage
    '******
    For iIndex = Me.txtUserId.LBound To Me.txtUserId.UBound
        On Error Resume Next
        Me.txtUserId(iIndex).Text = oUser.UserName
        If Err.Number <> 0 Then
            Err.Clear
        End If
        On Error GoTo ExitHandler
    Next iIndex

    If Len(msLotId) <> 0 Then
        Set oLot = moFwWIP.LotById(msLotId)
        sCurStage = oLot.CurrentStep.Steps.Item(1).Attributes(modConstFwAttr.gsSTEP_CUSTOMATTR_STAGE).Value
        For iIndex = 0 To Me.cboStage.ListCount - 1
            If Me.cboStage.List(iIndex) = sCurStage Then
                Me.cboStage.ListIndex = iIndex
                Exit For
            End If
        Next iIndex
        If Len(Me.cboStage.Text) <> 0 Then
            Me.cboLabelSpec.ListIndex = 0
            For iIndex = Me.fraLabelForm.LBound To Me.fraLabelForm.UBound
                'iIndex <> 8 --> Added by Jack on 2016/07/15 for bug fix.
                'WS 點選Lot 後, 再點F1 列印Label , 因為default 會把Lot 帶入所有Label , 造成出現FT Label的錯誤訊息
                '解法:FT_Label_PACK_INFO Label, 不要有default 帶入點選Lot 的.
                If iIndex <> 8 Then
'                If Me.fraLabelForm.Item(iIndex).Visible = True Then
                    Me.txtLotID.Item(iIndex).Text = msLotId
                    Call txtLotId_KeyPress(iIndex, 13)
'                    Exit For
'                End If
                End If
            Next iIndex
        End If
    End If
    
    'add by Sam start on 2008910 for ReqNO:M200807057
    If msLabelSpec <> "" Then
        Select Case msLabelSpec
            'Add by Sam START on 20210906 for Project FT 自動包裝
            Case "FT_Label_PACK_INFO"
                Me.cboStage.Text = msStage
                Call cboStage_Click
                Me.cboLabelSpec.Text = msLabelSpec
                Call cboLabelSpec_Click
            'Add by Sam END on 20210906 for Project FT 自動包裝
            Case Else 'modPrint.gsLABEL_RENESAS_SHIPPING
                Me.cboStage.Text = msStage
                Me.cboLabelSpec.Text = msLabelSpec
                Call cboLabelSpec_Click
                Me.cboStage.Locked = True
                Me.cboLabelSpec.Locked = True
                txtProdNo(miIndex) = msProdNo
                txtPartNo(miIndex) = msPartNo
        End Select
    End If
    'add by Sam end on 2008910 for ReqNO:M200807057
    
    'Added by Jack on 2010/04/12 for JC201000060. <Start>
    With spdLotList
        .MaxCols = miSPDLOTLIST_MAX_COLS
        '設定Headers
        .SetText mispdLotList_CHECK, 0, ""
        .SetText miSPDLOTLIST_LOT_ID, 0, "Lot NO"
        .SetText miSPDLOTLIST_PROD_NO, 0, "Product NO"
        .SetText mispdLotList_WQTY, 0, "Wafer Qty"
        .SetText mispdLotList_CQTY, 0, "Chip Qty"
        .SetText miSPDLOTLIST_LOT_OWNER, 0, "Lot Owner"
        .SetText miSPDLOTLIST_SPEED, 0, "Speed"
        .SetText mispdLotList_ROUTE, 0, "Route"
        .SetText mispdLotList_STEP, 0, "Step" 'Added by Jack on 2011/04/15
        
        '設定CheckBox欄位(置中).
        .Col = mispdLotList_CHECK
        .CellType = CellTypeCheckBox
        .TypeHAlign = TypeHAlignCenter

        '除了Check, Lot NO之外, 其餘欄位皆鎖住不可修改.
        For iIndex = miSPDLOTLIST_PROD_NO To spdLotList.MaxCols
            .Row = -1 '代表全部.
            .Col = iIndex
            .BackColor = &HC0C0C0
            .Lock = True
        Next
        Call .SetCellBorder(0, 0, .MaxCols, .MaxRows, 15, -1, CellBorderStyleSolid)
        Call ResizeSpdColumn(Me.spdLotList, moAppLog)
    End With
    'Added by Jack on 2010/04/12 for JC201000060. <End>
    
     'Added by Sam start on 2011/08/23 for JC201100218
    With spdWsEngLoc
        .MaxCols = miSPDWSENGLOC_MAX_COLS
               
        '設定CheckBox欄位(置中).
        .Col = miSPDWSENGLOC_CHECK
        .CellType = CellTypeCheckBox
        .TypeHAlign = TypeHAlignCenter

        '除了Check, Lot NO之外, 其餘欄位皆鎖住不可修改.
        For iIndex = miSPDWSENGLOC_WQTY To spdLotList.MaxCols
            .Row = -1 '代表全部.
            .Col = iIndex
            .BackColor = &HC0C0C0
            .Lock = True
        Next
        Call .SetCellBorder(0, 0, .MaxCols, .MaxRows, 15, -1, CellBorderStyleSolid)
        Call ResizeSpdColumn(Me.spdWsEngLoc, moAppLog)
    End With
    'Added by Sam end on 2011/08/23 for JC201100218
    
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

'Added by Jack on 2010/04/12 for JC201000060.
'加/減一筆的動作
Private Sub spdLotList_Advance(ByVal AdvanceNext As Boolean)
    If AdvanceNext Then
        Me.spdLotList.MaxRows = Me.spdLotList.MaxRows + 1
        Call Me.spdLotList.SetCellBorder(0, 0, Me.spdLotList.MaxCols, Me.spdLotList.MaxRows, 15, -1, CellBorderStyleSolid)
        Me.spdLotList.SetActiveCell miSPDLOTLIST_LOT_ID, Me.spdLotList.MaxRows
        
        Me.spdLotList.Col = miSPDLOTLIST_LOT_ID
        Me.spdLotList.Row = spdLotList.MaxRows
        Me.spdLotList.CellType = CellTypeEdit
        Me.spdLotList.BackColor = &HFFFFFF
        Me.spdLotList.EditMode = True
    Else
        If Me.spdLotList.MaxRows > 1 Then
            Me.spdLotList.MaxRows = Me.spdLotList.MaxRows - 1
            Call Me.spdLotList.SetCellBorder(0, 0, Me.spdLotList.MaxCols, Me.spdLotList.MaxRows, 15, -1, CellBorderStyleSolid)
        End If
    End If
End Sub
'Added by Jack on 2010/04/12 for JC201000060.
Private Sub spdLotList_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    Dim vLotID As Variant
    Dim oLot As FwLot
    
    If Col = mispdLotList_CHECK Then
        With spdLotList
         .Row = Row
         .Col = Col
         If .Value = 1 Then
            'Add by Sam start on 20201203 for Project CP 測前併批
            Select Case cboLabelSpec.Text
                Case gsLABEL_CP_VIRTUAL_MERGE, gsLABEL_CP_VIRTUAL_LOT_LABEL
                    'nothing
                Case Else
            'Add by Sam End on 20201203 for Project CP 測前併批
                    Call .GetText(miSPDLOTLIST_LOT_ID, Row, vLotID)
                    Set oLot = moFwWIP.LotById(vLotID)
                    If oLot Is Nothing Then
                        Call .SetText(1, Row, 0)
                    End If
            End Select
         End If
        End With
    End If
End Sub
'Added by Jack on 2010/04/12 for JC201000060.
Private Sub spdLotList_Change(ByVal Col As Long, ByVal Row As Long)
    Dim oLot       As FwLot
    Dim sSQL As String
    Dim colRS As Collection
    
    'gsLABEL_FT_BIN_CARD_LABEL added by Jack on 2011/04/18 for JC201100104.
    'gsLABEL_FT_LOT_INFO Added by Jack on 2016/06/03 for 2016 MES Phase-5 Project (FT eRunCard Paperless Phase-2 Project)
    'Modify by Sam on 20180814 for BE工業 3.5 BE MES Phase 24 -- CP Auto Dispatch,ADD cp_small_label
    'Modify by Sam on 20201203 for Project CP 測前併批 ,新增cp virtual label
    'Modify by Ernest on 20210315 for Project CP 測前併批 ,新增cp virtual label 新增AY
    If (Me.cboStage.Text = gsSTAGE_WS And Me.cboLabelSpec.Text = modPrint.gsLABEL_WS_SMALL_LABEL) Or _
       (Me.cboStage.Text = gsSTAGE_WS And Me.cboLabelSpec.Text = modPrint.gsLABEL_CP_SMALL_LABEL) Or _
        (Me.cboStage.Text = gsSTAGE_WS And Me.cboLabelSpec.Text = modPrint.gsLABEL_CP_VIRTUAL_LOT_LABEL) Or _
        (Me.cboStage.Text = gsSTAGE_WS And Me.cboLabelSpec.Text = modPrint.gsLABEL_CP_VIRTUAL_MERGE) Or _
       (Me.cboStage.Text = gsSTAGE_AY And Me.cboLabelSpec.Text = modPrint.gsLABEL_WS_SMALL_LABEL) Or _
       (Me.cboStage.Text = gsSTAGE_AY And Me.cboLabelSpec.Text = modPrint.gsLABEL_CP_VIRTUAL_LOT_LABEL) Or _
        (Me.cboStage.Text = gsSTAGE_AY And Me.cboLabelSpec.Text = modPrint.gsLABEL_CP_VIRTUAL_MERGE) Or _
       (Me.cboStage.Text = gsSTAGE_FT And Me.cboLabelSpec.Text = modPrint.gsLABEL_FT_BIN_CARD_LABEL) Or _
       (Me.cboStage.Text = gsSTAGE_FT And Me.cboLabelSpec.Text = modPrint.gsLABEL_FT_LOT_INFO) Or _
       (Me.cboStage.Text = gsSTAGE_FT And Me.cboLabelSpec.Text = modPrint.gsLABEL_FT_SMALL_LABEL) Then
       If Col = miSPDLOTLIST_LOT_ID Then
            With Me.spdLotList
                .Col = miSPDLOTLIST_LOT_ID
                .Row = Row
                'Add by Sam start on 20201203 for Project CP 測前併批
                Select Case cboLabelSpec.Text
                    Case gsLABEL_CP_VIRTUAL_LOT_LABEL, gsLABEL_CP_VIRTUAL_MERGE
                        'Add by Dexter on 20220809 for CP 併批測試 Phase II-第二階段需求 Key值取消IPN條件 <Start>
                        sSQL = "select substr(tla." & gsCAT_TLATT_IPN & ", 0, 4) as IPN, sum(tla." & gsCAT_TLATT_WAFERQTY & ") as wqty,sum(tla." & gsCAT_TLATT_CHIPQTY & ") as cqty " & _
                               " from (select distinct " & gsCAT_TVMG_VIRTUALLOTID & " ," & gsCAT_TVMG_LOTID & " " & _
                               "       from " & gsCAT_TBL_VIRTUAL_MERGE & " where " & gsCAT_TVMG_DELETEFLAG & "= 'N'  " & _
                               "        and " & gsCAT_TVMG_VIRTUALLOTID & " ='" & .Text & "' ) a " & _
                               " , " & gsCAT_TBL_LOT_ATTRIBUTE & " tla " & _
                               " where a." & gsCAT_TVMG_LOTID & " = tla." & gsCAT_TLATT_LOTID & " " & _
                               " group by substr(tla." & gsCAT_TLATT_IPN & ", 0, 4) "
                        'Add by Dexter on 20220809 for CP 併批測試 Phase II-第二階段需求 Key值取消IPN條件 <End>
                        Set colRS = moProRawSql.QueryDatabase(sSQL)
                        If colRS.Count = 0 Then
                            UtShowMsgBox ("The Virtual Lot [" & .Text & "] is not exist!")
                            Call .SetText(mispdLotList_CHECK, Row, 0) '取消勾消.
                            Call .SetText(miSPDLOTLIST_PROD_NO, Row, "")
                            Call .SetText(mispdLotList_WQTY, Row, "")
                            Call .SetText(mispdLotList_CQTY, Row, "")
                        Else
                            Call .SetText(mispdLotList_CHECK, Row, 1) '勾選.
                            Call .SetText(miSPDLOTLIST_PROD_NO, Row, colRS.Item(1).Item(gsCAT_TLATT_IPN))
                            Call .SetText(mispdLotList_WQTY, Row, colRS.Item(1).Item("wqty"))
                            Call .SetText(mispdLotList_CQTY, Row, colRS.Item(1).Item("cqty"))
                        End If
                    Case Else
                'Add by Sam End on 20201203 for Project CP 測前併批
                        Set oLot = moFwWIP.LotById(Trim(.Text))
                        If oLot Is Nothing Then
                            UtShowMsgBox ("The Lot [" & .Text & "] is not exist!")
                            Call .SetText(mispdLotList_CHECK, Row, 0) '取消勾消.
                            Call .SetText(miSPDLOTLIST_PROD_NO, Row, "")
                            Call .SetText(mispdLotList_WQTY, Row, "")
                            Call .SetText(mispdLotList_CQTY, Row, "")
                            Call .SetText(miSPDLOTLIST_LOT_OWNER, Row, "")
                            Call .SetText(miSPDLOTLIST_SPEED, Row, "")
                            Call .SetText(mispdLotList_ROUTE, Row, "") '
                            Call .SetText(mispdLotList_STEP, Row, "") 'added by Jack on 2011/04/18 for JC201100104.
                        Else
                            Call .SetText(mispdLotList_CHECK, Row, 1) '勾選.
                            Call .SetText(miSPDLOTLIST_PROD_NO, Row, oLot.CustomAttributes.Item(modConstFwAttr.gsLOT_CUSTOMATTR_IPN).Value)
                            Call .SetText(mispdLotList_WQTY, Row, oLot.CustomAttributes(modConstFwAttr.gsLOT_CUSTOMATTR_WAFER_QTY).Value)
                            Call .SetText(mispdLotList_CQTY, Row, oLot.CustomAttributes(modConstFwAttr.gsLOT_CUSTOMATTR_CHIP_QTY).Value)
                            Call .SetText(miSPDLOTLIST_LOT_OWNER, Row, oLot.CustomAttributes(modConstFwAttr.gsLOT_CUSTOMATTR_LOT_OWNER).Value)
                            Call .SetText(miSPDLOTLIST_SPEED, Row, GetSpeed(oLot.CustomAttributes(gsLOT_CUSTOMATTR_IPN), moProRawSql, moAppLog))
                            Call .SetText(mispdLotList_ROUTE, Row, oLot.PlanId)
                            Call .SetText(mispdLotList_STEP, Row, oLot.CustomAttributes.Item(modConstFwAttr.gsLOT_CUSTOMATTR_STEPNAME).Value) 'added by Jack on 2011/04/18 for JC201100104.
                        End If
                End Select
                Call .SetCellBorder(0, 0, .MaxCols, .MaxRows, 15, -1, CellBorderStyleSolid)
                Call ResizeSpdColumn(Me.spdLotList, moAppLog)
            End With
        End If
        
    End If
End Sub

Private Sub spdWsEngLoc_Advance(ByVal AdvanceNext As Boolean)
    If AdvanceNext Then
        Me.spdWsEngLoc.MaxRows = Me.spdWsEngLoc.MaxRows + 1
        Call Me.spdWsEngLoc.SetCellBorder(0, 0, Me.spdWsEngLoc.MaxCols, Me.spdWsEngLoc.MaxRows, 15, -1, CellBorderStyleSolid)
        Me.spdWsEngLoc.SetActiveCell miSPDWSENGLOC_LOT_ID, Me.spdWsEngLoc.MaxRows
        
        Me.spdWsEngLoc.Col = miSPDWSENGLOC_LOT_ID
        Me.spdWsEngLoc.Row = spdWsEngLoc.MaxRows
        Me.spdWsEngLoc.CellType = CellTypeEdit
        Me.spdWsEngLoc.BackColor = &HFFFFFF
        Me.spdWsEngLoc.EditMode = True
    Else
        If Me.spdWsEngLoc.MaxRows > 1 Then
            Me.spdWsEngLoc.MaxRows = Me.spdWsEngLoc.MaxRows - 1
            Call Me.spdWsEngLoc.SetCellBorder(0, 0, Me.spdWsEngLoc.MaxCols, Me.spdWsEngLoc.MaxRows, 15, -1, CellBorderStyleSolid)
        End If
    End If
End Sub

Private Sub spdWsEngLoc_Change(ByVal Col As Long, ByVal Row As Long)
    Dim sSQL As String
    Dim colRS As Collection
    Dim vLotID As Variant

'Add by Sam start on 20110823 for ReqNo:JC201100218
    If Me.cboStage.Text = gsSTAGE_WS And Me.cboLabelSpec.Text = modPrint.gsLABEL_WS_ENG_LOC_LABEL Then
        If Col = miSPDWSENGLOC_LOT_ID Then
            With Me.spdWsEngLoc
                .GetText Col, Row, vLotID
                sSQL = " select A." & gsCAT_TWELM_LOTID & " ,A." & gsCAT_TWELM_WQTY & _
                       " ,A." & gsCAT_TWELM_OWNERID & " ,A." & gsCAT_TWELM_OWNERDEP & _
                       " ,A." & gsCAT_TWELM_LOCID & _
                       " FROM " & gsCAT_TBL_WS_ENG_LOC_MASTER & " A " & _
                       " WHERE A." & gsCAT_TWELM_LOTID & "='" & CStr(vLotID) & "' "
                Set colRS = moProRawSql.QueryDatabase(sSQL)
                If colRS.Count > 0 Then
                    Call .SetText(miSPDWSENGLOC_CHECK, Row, 1) '勾選.
                    Call .SetText(miSPDWSENGLOC_LOT_ID, Row, colRS.Item(1).Item(gsCAT_TWELM_LOTID))
                    Call .SetText(miSPDWSENGLOC_WQTY, Row, colRS.Item(1).Item(gsCAT_TWELM_WQTY))
                    Call .SetText(miSPDWSENGLOC_OWNERID, Row, colRS.Item(1).Item(gsCAT_TWELM_OWNERID))
                    Call .SetText(miSPDWSENGLOC_OWNERDEP, Row, colRS.Item(1).Item(gsCAT_TWELM_OWNERDEP))
                Else
                    UtShowMsgBox ("The Lot is not exist in Tbl_Ws_Eng_Loc_Master!")
                    Call .SetText(miSPDWSENGLOC_CHECK, Row, "")
                    Call .SetText(miSPDWSENGLOC_LOT_ID, Row, "")
                    Call .SetText(miSPDWSENGLOC_WQTY, Row, "")
                    Call .SetText(miSPDWSENGLOC_OWNERID, Row, "")
                    Call .SetText(miSPDWSENGLOC_OWNERDEP, Row, "")
                End If
            End With
        End If
       
    End If
 'Add by Sam end on 20110823 for ReqNo:JC201100218
End Sub
'tabLabelPackInfo_Click( ) added by Jack on 2016/06/03 for 2016 MES Phase-5 Project (FT eRunCard Paperless Phase-2 Project).
Private Sub tabLabelPackInfo_Click(PreviousTab As Integer)
    If Me.tabLabelPackInfo.Tab = 1 Then
        cboLPI_LabelSpecNO.ListIndex = -1
        Me.cboLPI_LabelSpecNO.SetFocus
        cboLPI_CarrierType.Clear
        cboLPI_BoxingSpecNO.Clear
        cboLPI_Brand.Clear
        cboLPI_PinCount.Clear
        cboLPI_PackageCode.Clear
        
        Call GetLPI_LabelSpec
    Else
        Me.txtLotID(8).Text = ""
        Me.txtLotID(8).Text = ""
        Me.txtLotID(8).Locked = False
        Me.txtLotID(8).BackColor = &H80000005
        Me.txtLotID(8).SetFocus
    End If
End Sub

'Add gsLABEL_FT_TR_LABEL by HouYu for 202000413 on 20220422
Private Sub txt_FT_TR_QtyReel_Change()
On Error GoTo ExitHandler:
Dim sProcID As String
Dim typErrInfo As tErrInfo
Dim oLot       As FwLot
Dim sPartNo    As String

Dim sSQL    As String
Dim colRS   As Collection
Dim oItem   As Object

Dim i       As Integer
Dim sIPN    As String
Dim sCQty   As String
Dim iReel   As Integer

'----
' Init
'----
    sProcID = "txt_FT_TR_QtyReel_KeyPress"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog)

'----
' Condition Checking
'----
    If Me.txt_FT_TR_QtyReel.Text = "" Then
        GoTo ExitHandler
    End If
    If Not IsNumeric(Me.txt_FT_TR_QtyReel.Text) Then
        Me.txt_FT_TR_QtyReel.Text = ""
        GoTo ExitHandler
    End If
'----
' Action
'----
    sSQL = "SELECT * FROM " & gsCAT_TBL_LOT_ATTRIBUTE & _
            " WHERE " & gsCAT_TLATT_LOTID & " = '" & Me.txtLotID(miIndex).Text & "'"
    
    Set colRS = moProRawSql.QueryDatabase(sSQL)
    If colRS.Count > 0 Then
        sCQty = colRS.Item(1).Item(gsCAT_TLATT_CHIPQTY)
    End If
    
    Me.cbo_FT_TR_ReelId.Clear
    
    If Me.txt_FT_TR_QtyReel.Text = 0 Then
        Call UtShowMsgBox("Reel ID list can't be zero!")
        GoTo ExitHandler
    Else
        iReel = Fix(sCQty / Me.txt_FT_TR_QtyReel.Text)
    End If
    
    If iReel > 0 Then
        Me.cbo_FT_TR_ReelId.AddItem "All"
        For i = 1 To iReel
            Me.cbo_FT_TR_ReelId.AddItem i
        Next i
    Else
        Call UtShowMsgBox("QtyReel(" & Me.txt_FT_TR_QtyReel.Text & ") can not be large than CQty(" & sCQty & ")!")
        Me.txt_FT_TR_QtyReel.Text = ""
        GoTo ExitHandler
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

Private Sub txtBoxQty_Change()
    If Not IsNumeric(Me.txtBoxQty.Text) Then
        Me.txtBoxQty.Text = ""
    End If
End Sub

'================================================================================
' Sub: txtLotId_KeyPress()
'--------------------------------------------------------------------------------
' Description:  <Type your Sub description here...>
'--------------------------------------------------------------------------------
' Author:       Ivan, CIT 2002-06-05
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
Private Sub txtLotId_KeyPress(Index As Integer, KeyAscii As Integer)
On Error GoTo ExitHandler:
Dim sProcID As String
Dim typErrInfo As tErrInfo
Dim oLot       As FwLot
Dim sPartNo    As String

Dim sSQL    As String
Dim colRS   As Collection
Dim oItem   As Object

Dim i       As Integer
Dim sIPN    As String
Dim sCQty   As String
Dim iReel   As Integer

'----
' Init
'----
    sProcID = "txtLotId_KeyPress"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog)

'----
' Condition Checking
'----
    If KeyAscii <> 13 Then
        GoTo ExitHandler
    End If
    
    txtLotID(Index).Locked = True
    txtLotID(Index).BackColor = &H80000000
'----
' Action
'----
    
    Select Case Index
        Case 0
            Set oLot = moFwWIP.LotById(Trim(Me.txtLotID(Index).Text))
            If oLot Is Nothing Then
                UtShowMsgBox ("The Lot is not exist!")
                GoTo ExitHandler
            End If
            Me.txtProdNo(Index).Text = oLot.CustomAttributes.Item(modConstFwAttr.gsLOT_CUSTOMATTR_IPN).Value
            Me.txtQty(Index).Text = oLot.CustomAttributes.Item(modConstFwAttr.gsLOT_CUSTOMATTR_CHIP_QTY).Value
        Case 1
            Set oLot = moFwWIP.LotById(Trim(Me.txtLotID(Index).Text))
            If oLot Is Nothing Then
                UtShowMsgBox ("The Lot is not exist!")
                GoTo ExitHandler
            End If
            Me.txtProdNo(Index).Text = oLot.CustomAttributes.Item(modConstFwAttr.gsLOT_CUSTOMATTR_IPN).Value
            Me.txtCqty(Index).Text = oLot.CustomAttributes(modConstFwAttr.gsLOT_CUSTOMATTR_CHIP_QTY).Value
            Me.txtWQty(Index).Text = oLot.CustomAttributes(modConstFwAttr.gsLOT_CUSTOMATTR_WAFER_QTY).Value
            sPartNo = modCustom.RetrieveEPNbyIPN(moAppLog, _
                                                moFwWIP, _
                                                moFwWF, _
                                                moCwMbx, _
                                                oLot.CustomAttributes(modConstFwAttr.gsLOT_CUSTOMATTR_IPN).Value)
            Me.txtPartNo(Index).Text = sPartNo
        
        Case 2
            Set oLot = moFwWIP.LotById(Trim(Me.txtLotID(Index).Text))
            If oLot Is Nothing Then
                UtShowMsgBox ("The Lot is not exist!")
                GoTo ExitHandler
            End If
            Me.txtProdNo(Index).Text = oLot.CustomAttributes.Item(modConstFwAttr.gsLOT_CUSTOMATTR_IPN).Value
            Me.txtCqty(Index).Text = oLot.CustomAttributes(modConstFwAttr.gsLOT_CUSTOMATTR_CHIP_QTY).Value
            Me.txtWQty(Index).Text = oLot.CustomAttributes(modConstFwAttr.gsLOT_CUSTOMATTR_WAFER_QTY).Value
        Case 3
            Set oLot = moFwWIP.LotById(Trim(Me.txtLotID(Index).Text))
            If oLot Is Nothing Then
                UtShowMsgBox ("The Lot is not exist!")
                GoTo ExitHandler
            End If
            Me.txtProdNo(Index).Text = oLot.CustomAttributes.Item(modConstFwAttr.gsLOT_CUSTOMATTR_IPN).Value
            Me.txtCqty(Index).Text = oLot.CustomAttributes(modConstFwAttr.gsLOT_CUSTOMATTR_CHIP_QTY).Value
            Me.txtWQty(Index).Text = oLot.CustomAttributes(modConstFwAttr.gsLOT_CUSTOMATTR_WAFER_QTY).Value
            Me.txtLotOwner(Index).Text = oLot.CustomAttributes(modConstFwAttr.gsLOT_CUSTOMATTR_LOT_OWNER).Value
            Me.txtSpeed(Index).Text = GetSpeed(oLot.CustomAttributes(gsLOT_CUSTOMATTR_IPN), moProRawSql, moAppLog)
        Case 4
            Set oLot = moFwWIP.LotById(Trim(Me.txtLotID(Index).Text))
            If oLot Is Nothing Then
                UtShowMsgBox ("The Lot is not exist!")
                GoTo ExitHandler
            End If
            Me.txtProdNo(Index).Text = oLot.CustomAttributes.Item(modConstFwAttr.gsLOT_CUSTOMATTR_IPN).Value
            Me.txtCqty(Index).Text = oLot.CustomAttributes(modConstFwAttr.gsLOT_CUSTOMATTR_CHIP_QTY).Value
            Me.txtWQty(Index).Text = oLot.CustomAttributes(modConstFwAttr.gsLOT_CUSTOMATTR_WAFER_QTY).Value
            Me.txtLotOwner(Index).Text = oLot.CustomAttributes(modConstFwAttr.gsLOT_CUSTOMATTR_LOT_OWNER).Value
            Me.txtRouteId(Index).Text = oLot.PlanId
            'added by Jack on 2011/04/19 for JC201100104 <Start>
            Me.txtStepName(Index).Text = oLot.CustomAttributes(modConstFwAttr.gsLOT_CUSTOMATTR_STEPNAME).Value
            'added by Jack on 2011/04/19 for JC201100104 <End>
        Case 5
            Set oLot = moFwWIP.LotById(Trim(Me.txtLotID(Index).Text))
            If oLot Is Nothing Then
                UtShowMsgBox ("The Lot is not exist!")
                GoTo ExitHandler
            End If
            Me.txtProdNo(Index).Text = oLot.CustomAttributes.Item(modConstFwAttr.gsLOT_CUSTOMATTR_IPN).Value
            Me.txtCqty(Index).Text = oLot.CustomAttributes(modConstFwAttr.gsLOT_CUSTOMATTR_CHIP_QTY).Value
            Me.txtWQty(Index).Text = oLot.CustomAttributes(modConstFwAttr.gsLOT_CUSTOMATTR_WAFER_QTY).Value
        Case 6
            sSQL = "select A.WIPID         as PARENTLOTID," & vbNewLine & _
                    "      D.ORIGINALVALUE as PARENTLOTQTY," & vbNewLine & _
                    "      B.VALDATA       as CHILDLOTID," & vbNewLine & _
                    "      E.ORIGINALVALUE as CHILDLOTQTY" & vbNewLine
            sSQL = sSQL & _
                    "  from FWMERGE A," & vbNewLine & _
                    "      FWMERGE_PN2M B," & vbNewLine & _
                    "      FWWIPHISTORY C," & vbNewLine & _
                    "      (select D1.ORIGINALVALUE, D2.GROUPHISTKEY" & vbNewLine & _
                    "         from FWWIPTRANSACTION D1, FWWIPHISTORY D2" & vbNewLine & _
                    "        where D1.WIPID = '" & Me.txtLotID(6).Text & "'" & vbNewLine & _
                    "          and D1.ACTIVITY = 'ModifyAttribute'" & vbNewLine & _
                    "          and D1.ATTRIBUTE = 'ChipQty'" & vbNewLine & _
                    "          and D1.SYSID = D2.WIPTXN) D," & vbNewLine & _
                    "      (select E1.ORIGINALVALUE, E2.GROUPHISTKEY, E1.WIPID" & vbNewLine & _
                    "         from FWWIPTRANSACTION E1, FWWIPHISTORY E2" & vbNewLine & _
                    "        where E1.ACTIVITY = 'ModifyAttribute'" & vbNewLine & _
                    "          and E1.ATTRIBUTE = 'ChipQty'" & vbNewLine & _
                    "          and E1.SYSID = E2.WIPTXN) E" & vbNewLine
            sSQL = sSQL & _
                    " where MERGESTEPID between '51201' and '51299'" & vbNewLine & _
                    "   and A.WIPID = '" & Me.txtLotID(6).Text & "'" & vbNewLine & _
                    "   and A.SYSID = B.FROMID" & vbNewLine & _
                    "   and B.LINKNAME = 'childLotCollection'" & vbNewLine & _
                    "   and A.SYSID = C.WIPTXN" & vbNewLine & _
                    "   and C.GROUPHISTKEY = D.GROUPHISTKEY" & vbNewLine & _
                    "   and E.WIPID = B.VALDATA" & vbNewLine & _
                    "   and C.GROUPHISTKEY = E.GROUPHISTKEY"
            Set colRS = moProRawSql.QueryDatabase(sSQL)
            If colRS.Count > 0 Then
                Me.txtCqty(6).Text = colRS.Item(1).Item("PARENTLOTQTY")
                With Me.spdChildLot
                    .MaxRows = 0
                    For Each oItem In colRS
                        .MaxRows = .MaxRows + 1
                        .SetText 1, .MaxRows, oItem("CHILDLOTID")
                        .SetText 2, .MaxRows, oItem("CHILDLOTQTY")
                    Next oItem
                End With
            End If
        Case 7 'add by Nelson on 2008/06/03 for ReqNo:M200806003
            Set oLot = moFwWIP.LotById(Trim(Me.txtLotID(Index).Text))
            If oLot Is Nothing Then
                UtShowMsgBox ("The Lot is not exist!")
                GoTo ExitHandler
            End If
            Me.txtProdNo(Index).Text = oLot.CustomAttributes.Item(modConstFwAttr.gsLOT_CUSTOMATTR_IPN).Value
            Me.txtCqty(Index).Text = oLot.CustomAttributes(modConstFwAttr.gsLOT_CUSTOMATTR_CHIP_QTY).Value
            Me.txtWQty(Index).Text = oLot.CustomAttributes(modConstFwAttr.gsLOT_CUSTOMATTR_WAFER_QTY).Value
            sPartNo = modCustom.RetrieveEPNbyIPN(moAppLog, _
                                                moFwWIP, _
                                                moFwWF, _
                                                moCwMbx, _
                                                oLot.CustomAttributes(modConstFwAttr.gsLOT_CUSTOMATTR_IPN).Value)
            Me.txtPartNo(Index).Text = sPartNo
        'Added by Jack on 2016/06/03 for 2016 MES Phase-5 Project (FT eRunCard Paperless Phase-2 Project) <Start>
        Case 8
            Me.txtLPI_CarrierType = ""
            Me.txtLPI_BoxingSpecNO = ""
            Me.txtLPI_Brand = ""
            Me.txtLPI_PinCount = ""
            Me.txtLPI_PackageCode = ""
            Me.txtLPI_LabelSpecNo = ""
            
            Set oLot = moFwWIP.LotById(Trim(Me.txtLotID(Index).Text))
            If oLot Is Nothing Then
                UtShowMsgBox ("The Lot '" & Trim(Me.txtLotID(Index).Text) & "' is not exist!")
                Me.txtLotID(Index).Text = ""
                txtLotID(Index).Locked = False
                txtLotID(Index).BackColor = &H80000005
                GoTo ExitHandler
            End If
            
            'IPN added by Jack on 2016/07/04.
            sSQL = "select a." & gsCAT_TIM_CARRIER_TYPE & ", a." & gsCAT_TIM_BOXING_SPEC_NO & ", " _
                       & " a." & gsCAT_TIM_BRAND & ", a." & gsCAT_TIM_PIN_COUNT & ", " _
                       & " a." & gsCAT_TIM_PACKAGE_CODE & ", c." & gsCAT_TCUM_LABELSPECNO & ", " _
                       & " a." & gsCAT_TIM_IPN
            sSQL = sSQL _
                       & " from " _
                       & gsCAT_TBL_IPN_MASTER & " a, " & gsCAT_TBL_PRM_BE_SPEC & " b, " _
                       & gsCAT_TBL_CUSTOMER_MASTER & " c, " & gsCAT_TBL_LOT_ATTRIBUTE & " d "
            'defaults='Y' added by Jack on 2016/07/04
            sSQL = sSQL _
                       & " where d." & gsCAT_TLATT_LOTID & " = '" & oLot.Id & "' " _
                       & "   and d." & gsCAT_TLATT_IPN & " = a." & gsCAT_TIM_IPN & " " _
                       & "   and d." & gsCAT_TLATT_IPN & " = b." & gsCAT_TPBS_IPN & "(+) " _
                       & "   and b." & gsCAT_TPBS_CUSTOMER_NO & " = c." & gsCAT_TCUM_CUSTOMERNO & "(+) " _
                       & "   and b." & gsCAT_TPBS_DEFAULTS & "= 'Y' "
                       
            Set colRS = moProRawSql.QueryDatabase(sSQL)
            If colRS.Count > 0 Then
                Me.txtLPI_CarrierType = colRS.Item(1).Item(gsCAT_TIM_CARRIER_TYPE)
                Me.txtLPI_BoxingSpecNO = colRS.Item(1).Item(gsCAT_TIM_BOXING_SPEC_NO)
                Me.txtLPI_Brand = colRS.Item(1).Item(gsCAT_TIM_BRAND)
                Me.txtLPI_PinCount = colRS.Item(1).Item(gsCAT_TIM_PIN_COUNT)
                Me.txtLPI_PackageCode = colRS.Item(1).Item(gsCAT_TIM_PACKAGE_CODE)
                Me.txtLPI_LabelSpecNo = colRS.Item(1).Item(gsCAT_TCUM_LABELSPECNO)
                Me.txtLPI_IPN = colRS.Item(1).Item(gsCAT_TIM_IPN) 'IPN added by Jack on 2016/07/04.
            Else
                '為了因應工程品不follow PROD 的情況, 所以Rule  改成...
                '若IPN不存在, 且User 用的是auto  mode , 則顯示message
                '因工程品後段"是否follow PROD" 的欄位, 並沒有傳至MES, 所以產線印出Label, 還是會比對委測單, 以確保品質.
                'IPN 不存在, 或Tbl_prm_backend_spec.defaults<>'Y' ,  請以Manual 列印
                Call UtShowMsgBox("IPN 不存在 或Tbl_prm_backend_spec.defaults<>'Y', " & vbNewLine & _
                                  "請以Manual mode 列印" & vbNewLine & vbNewLine & _
                                  "IPN is not exist or Tbl_prm_backend_spec.defaults<>'Y'" & vbNewLine & _
                                  "please use Manual mode Print Label ")
                GoTo ExitHandler
            End If
        'Added by Jack on 2016/06/03 for 2016 MES Phase-5 Project (FT eRunCard Paperless Phase-2 Project) <End>
        'Add gsLABEL_FT_TR_LABEL by HouYu for 202000413 on 20220422
        Case 9
            sSQL = "SELECT * FROM " & gsCAT_TBL_LOT_ATTRIBUTE & _
                    " WHERE " & gsCAT_TLATT_LOTID & " = '" & Me.txtLotID(Index).Text & "'"
            
            Set colRS = moProRawSql.QueryDatabase(sSQL)
            If colRS.Count > 0 Then
                sIPN = colRS.Item(1).Item(gsCAT_TLATT_IPN)
                sCQty = colRS.Item(1).Item(gsCAT_TLATT_CHIPQTY)
            End If
            
            sSQL = "SELECT TLA." & gsCAT_TLATT_LOTID & "," & _
                          "TLA." & gsCAT_TIM_IPN & "," & _
                          "TIM." & gsCAT_TIM_IPN & "," & _
                          "TIM." & gsCAT_TIM_PRODBODY & "," & _
                          "TIM." & gsCAT_TIM_PIN_COUNT & "," & _
                          "TIM." & gsCAT_TIM_PACKAGE_CODE & "," & _
                          "TIM." & gsCAT_TIM_BODY_SIZE & "" & _
                     " FROM " & gsCAT_TBL_LOT_ATTRIBUTE & " TLA, " & gsCAT_TBL_IPN_MASTER & " TIM" & _
                    " WHERE TLA." & gsCAT_TLATT_IPN & " = TIM." & gsCAT_TIM_IPN & "(+)" & _
                      "AND TLA." & gsCAT_TLATT_LOTID & " = '" & Me.txtLotID(Index).Text & "'"
            
            Set colRS = moProRawSql.QueryDatabase(sSQL)
            
            If colRS.Count > 0 Then
                If colRS.Item(1).Item(3) <> "" Then
                    Me.txt_FT_TR_IPN.Text = colRS.Item(1).Item(gsCAT_TIM_PRODBODY) & "-" & colRS.Item(1).Item(gsCAT_TIM_PIN_COUNT) & colRS.Item(1).Item(gsCAT_TIM_PACKAGE_CODE) & " " & colRS.Item(1).Item(gsCAT_TIM_BODY_SIZE)
                Else
                    Me.txt_FT_TR_IPN.Text = colRS.Item(1).Item(2)
                End If
            End If
            
            sSQL = "SELECT " & gsCAT_TPBS_CARRIER_QTY & _
                    " FROM " & gsCAT_TBL_PRM_BE_SPEC & _
                    " WHERE " & gsCAT_TPBS_IPN & "='" & sIPN & _
                    "' AND " & gsCAT_TPBS_DEFAULTS & " ='Y'"
            Set colRS = moProRawSql.QueryDatabase(sSQL)
            If colRS.Count > 0 Then
                Me.txt_FT_TR_QtyReel.Text = colRS.Item(1).Item(gsCAT_TPBS_CARRIER_QTY)
            End If
        'Add gsLABEL_FT_TR_LABEL by HouYu for 202000413 on 20220422
    End Select
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

Private Sub txtPrintQty_Change()
    If Not IsNumeric(Me.txtPrintQty.Text) Then
        Me.txtPrintQty.Text = ""
    End If
End Sub
Private Sub txtQty_Change(Index As Integer)
    If Not IsNumeric(Me.txtQty(Index).Text) Then
        Me.txtQty(Index).Text = ""
    End If
End Sub

'================================================================================
' function: GetBizType()
'--------------------------------------------------------------------------------
' Description:  <Type your function description here...>
'--------------------------------------------------------------------------------
' Author:       Sam, CIT 2008-08-05
'--------------------------------------------------------------------------------
' Parameters:
'................................................................................
' ARGUMENTS
'   moapplog            (I) [Optional] Valid reference to a Object object
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
'Marked by Jack on 2012/06/13 for support KGD_AEB Project [Hanno]
'統一呼叫 modPrint.GetBizType( ).
'================================================================================
'Private Function GetBizType(ByVal sIPN As String, ByVal iIndex As Integer) As String
'On Error GoTo ExitHandler:
'Dim sProcID As String
'Dim typErrInfo As tErrInfo
'Dim sTable As String
'Dim sColumn As String
'Dim sWhere As String
'Dim sOther  As String
'Dim colRaws As Collection
'Dim oRaws As FwStrings
'Dim sSQL As String
''----
'' Init
''----
'    sProcID = "GetBizType"
'    Call LogProcIn(msMODULE_ID, sProcID, moAppLog)
'    GetBizType = ""
''----
'' Condition Chking
''----
'
'
''----
'' Action
''----
'    sTable = gsCAT_TBL_IPN_MASTER & " A," & _
'             gsCAT_TBL_PROD_BODY & " B "
'
'    sColumn = " B." & gsCAT_TPB_SBU_LINE & _
'              ",A." & gsCAT_TIM_BOM_LEVEL & _
'              ",A." & gsCAT_TIM_BIZTYPE
'
'    sColumn = sColumn & ",A." & gsCAT_TIM_PACKAGE_CODE  'Add by Sam on 20091014 for JC900116,KGD變更條件
'
'    'sIPN modified by Jack on 2010/04/12 for JC201000060.
''    sWhere = "A." & gsCAT_TIM_PRODBODY & "=B." & gsCAT_TPB_PROD_BODY & "(+) and " & _
''             "A." & gsCAT_TIM_IPN & "='" & Me.txtProdNo(iIndex).Text & "' and " & _
''             "A." & gsCAT_TIM_DELETE_FLAG & "='N'"
'    sWhere = "A." & gsCAT_TIM_PRODBODY & "=B." & gsCAT_TPB_PROD_BODY & "(+) and " & _
'             "A." & gsCAT_TIM_IPN & "='" & sIPN & "' and " & _
'             "A." & gsCAT_TIM_DELETE_FLAG & "='N'"
'
'    sSQL = "SELECT " & sColumn & " FROM " & sTable & " WHERE " & sWhere
'
''    Set colRaws = SelectRawSqlTxn(moAppLog, moFwWIP, moFwWF, moCwMbx, "CAT", sTable, sColumn, sWhere, sOther)
'    Set colRaws = SelectRawSqlTxn1(moAppLog, moFwWIP, moFwWF, moCwMbx, "CAT", sSQL)
'    If colRaws Is Nothing Then
'        Call RaiseError(glERR_INVALIDOBJECT, _
'                    FormatErrorText(gsETX_INVALIDOBJECT, "Collection"))
'    End If
'
'    If colRaws.Count > 0 Then
''Modify by Sam on 20091014 for JC900116,KGD變更條件
''        If colRaws.Item(1).Item(gsCAT_TPB_SBU_LINE) = "NVM" And _
''           colRaws.Item(1).Item(gsCAT_TIM_BOM_LEVEL) = "DB" And _
''           colRaws.Item(1).Item(gsCAT_TIM_BIZTYPE) = "K" Then
'
'        If colRaws.Item(1).Item(gsCAT_TPB_SBU_LINE) = "NVM" And _
'           colRaws.Item(1).Item(gsCAT_TIM_PACKAGE_CODE) = "W" And _
'           colRaws.Item(1).Item(gsCAT_TIM_BIZTYPE) = "K" Then
'
'           GetBizType = "KGD"
'        End If
'    End If
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
'
'
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
'        Call HandleError(True, typErrInfo, , moAppLog)
'    End If
'End Function
'Added by Jack on 2011/04/15
Private Sub txtPrintFailQty_Change()
    If Not IsNumeric(Me.txtPrintFailQty.Text) Then
        Me.txtPrintFailQty.Text = ""
    End If
End Sub
'================================================================================
' function: GetLPI_LabelSpec()
'--------------------------------------------------------------------------------
' Description:  <Type your function description here...>
'--------------------------------------------------------------------------------
' Author:       Jack Hsieh, MXIC 2016/06/03 for 2016 MES Phase-5 Project (FT eRunCard Paperless Phase-2 Project).
'--------------------------------------------------------------------------------
' Parameters:
'................................................................................
' ARGUMENTS
'   moapplog            (I) [Optional] Valid reference to a Object object
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
Private Function GetLPI_LabelSpec() As Boolean
On Error GoTo ExitHandler:
Dim sProcID As String
Dim typErrInfo As tErrInfo
Dim sTable As String
Dim sColumn As String
Dim sWhere As String
Dim sOther  As String
Dim colRaws As Collection
Dim oRaws As FwStrings
'----
' Init
'----
    sProcID = "GetLPI_LabelSpec"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog)
    GetLPI_LabelSpec = True
    Me.cboLPI_LabelSpecNO.Clear
'----
' Condition Chking
'----


'----
' Action
'----
    'select distinct LABELSPECNO from Tbl_Ft_Label_Spec_Info
    sTable = gsCAT_TBL_FT_LABEL_SPEC_INFO

    sColumn = " distinct " & gsCAT_TFLSI_LABELSPECNO
    
    sWhere = gsCAT_TFLSI_DELETEFLAG & "='N'"

    Set colRaws = SelectRawSqlTxn(moAppLog, moFwWIP, moFwWF, moCwMbx, "CAT", sTable, sColumn, sWhere, sOther)
    If colRaws Is Nothing Then
        Call RaiseError(glERR_INVALIDOBJECT, _
                    FormatErrorText(gsETX_INVALIDOBJECT, "Collection"))
    End If
    
    Me.cboLPI_LabelSpecNO.AddItem ""
    For Each oRaws In colRaws
        Me.cboLPI_LabelSpecNO.AddItem oRaws.Item(1)
    Next oRaws
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
        GetLPI_LabelSpec = False
        Call HandleError(True, typErrInfo, , moAppLog)
    End If
End Function


'================================================================================
' function: GetLPI_CarrierType()
'--------------------------------------------------------------------------------
' Description:  <Type your function description here...>
'--------------------------------------------------------------------------------
' Author:       Jack Hsieh, MXIC 2016/06/03 for 2016 MES Phase-5 Project (FT eRunCard Paperless Phase-2 Project).
'--------------------------------------------------------------------------------
' Parameters:
'................................................................................
' ARGUMENTS
'   moapplog            (I) [Optional] Valid reference to a Object object
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
Private Function GetLPI_CarrierType(ByVal sLabelSpec As String) As Boolean
On Error GoTo ExitHandler:
Dim sProcID As String
Dim typErrInfo As tErrInfo
Dim sTable As String
Dim sColumn As String
Dim sWhere As String
Dim sOther  As String
Dim colRaws As Collection
Dim oRaws As FwStrings
'----
' Init
'----
    sProcID = "GetLPI_CarrierType"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog)
    GetLPI_CarrierType = True
    Me.cboLPI_CarrierType.Clear
'----
' Condition Chking
'----


'----
' Action
'----
    'select distinct CARRIERTYPE from Tbl_Ft_Label_Spec_Info where LABELSPECNO = 'sLabelSpec'
    sTable = gsCAT_TBL_FT_LABEL_SPEC_INFO

    sColumn = "distinct " & gsCAT_TFLSI_CARRIERTYPE
    
    sWhere = gsCAT_TFLSI_LABELSPECNO & "='" & sLabelSpec & "' and " & _
             gsCAT_TFLSI_DELETEFLAG & "='N'"

    Set colRaws = SelectRawSqlTxn(moAppLog, moFwWIP, moFwWF, moCwMbx, "CAT", sTable, sColumn, sWhere, sOther)
    If colRaws Is Nothing Then
        Call RaiseError(glERR_INVALIDOBJECT, _
                    FormatErrorText(gsETX_INVALIDOBJECT, "Collection"))
    End If
    
    Me.cboLPI_CarrierType.AddItem ""
    For Each oRaws In colRaws
        Me.cboLPI_CarrierType.AddItem oRaws.Item(1)
    Next oRaws
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
        GetLPI_CarrierType = False
        Call HandleError(True, typErrInfo, , moAppLog)
    End If
End Function


'================================================================================
' function: GetLPI_BoxingSpecNO()
'--------------------------------------------------------------------------------
' Description:  <Type your function description here...>
'--------------------------------------------------------------------------------
' Author:       Jack Hsieh, MXIC 2016/06/03 for 2016 MES Phase-5 Project (FT eRunCard Paperless Phase-2 Project).
'--------------------------------------------------------------------------------
' Parameters:
'................................................................................
' ARGUMENTS
'   moapplog            (I) [Optional] Valid reference to a Object object
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
Private Function GetLPI_BoxingSpecNO(ByVal sCarrierType As String) As Boolean
On Error GoTo ExitHandler:
Dim sProcID As String
Dim typErrInfo As tErrInfo
Dim sTable As String
Dim sColumn As String
Dim sWhere As String
Dim sOther  As String
Dim colRaws As Collection
Dim oRaws As FwStrings
'----
' Init
'----
    sProcID = "GetLPI_BoxingSpecNO"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog)
    GetLPI_BoxingSpecNO = True
    Me.cboLPI_BoxingSpecNO.Clear
'----
' Condition Chking
'----


'----
' Action
'----
    'select distinct BOXINGSPECNO from TBL_FT_Label_Boxing_INFO where CARRIERTYPE = 'sCarrierType'
    sTable = gsCAT_TBL_FT_LABEL_BOXING_INFO

    sColumn = "distinct " & gsCAT_TFLBI_BOXINGSPECNO
    
    sWhere = gsCAT_TFLBI_CARRIERTYPE & "='" & sCarrierType & "' and " & _
             gsCAT_TFLBI_DELETEFLAG & "='N'"

    Set colRaws = SelectRawSqlTxn(moAppLog, moFwWIP, moFwWF, moCwMbx, "CAT", sTable, sColumn, sWhere, sOther)
    If colRaws Is Nothing Then
        Call RaiseError(glERR_INVALIDOBJECT, _
                    FormatErrorText(gsETX_INVALIDOBJECT, "Collection"))
    End If
    
    Me.cboLPI_BoxingSpecNO.AddItem ""
    For Each oRaws In colRaws
        Me.cboLPI_BoxingSpecNO.AddItem oRaws.Item(1)
    Next oRaws
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
        GetLPI_BoxingSpecNO = False
        Call HandleError(True, typErrInfo, , moAppLog)
    End If
End Function
'================================================================================
' function: GetLPI_Brand()
'--------------------------------------------------------------------------------
' Description:  <Type your function description here...>
'--------------------------------------------------------------------------------
' Author:       Jack Hsieh, MXIC 2016/06/03 for 2016 MES Phase-5 Project (FT eRunCard Paperless Phase-2 Project).
'--------------------------------------------------------------------------------
' Parameters:
'................................................................................
' ARGUMENTS
'   moapplog            (I) [Optional] Valid reference to a Object object
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
Private Function GetLPI_Brand(ByVal sCarrierType As String, ByVal sBoxingSpecNo As String) As Boolean
On Error GoTo ExitHandler:
Dim sProcID As String
Dim typErrInfo As tErrInfo
Dim sTable As String
Dim sColumn As String
Dim sWhere As String
Dim sOther  As String
Dim colRaws As Collection
Dim oRaws As FwStrings
'----
' Init
'----
    sProcID = "GetLPI_Brand"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog)
    GetLPI_Brand = True
    Me.cboLPI_Brand.Clear
'----
' Condition Chking
'----


'----
' Action
'----
    'select distinct BRAND from TBL_FT_Label_Boxing_INFO where CARRIERTYPE = 'Tape & Reel'
    '   and BOXINGSPECNO = '4040-0028'
    sTable = gsCAT_TBL_FT_LABEL_BOXING_INFO

    sColumn = "distinct " & gsCAT_TFLBI_BRAND
    
    sWhere = gsCAT_TFLBI_CARRIERTYPE & "='" & sCarrierType & "' and " & _
             gsCAT_TFLBI_BOXINGSPECNO & "='" & sBoxingSpecNo & "' and " & _
             gsCAT_TFLBI_DELETEFLAG & "='N'"

    Set colRaws = SelectRawSqlTxn(moAppLog, moFwWIP, moFwWF, moCwMbx, "CAT", sTable, sColumn, sWhere, sOther)
    If colRaws Is Nothing Then
        Call RaiseError(glERR_INVALIDOBJECT, _
                    FormatErrorText(gsETX_INVALIDOBJECT, "Collection"))
    End If
    
    Me.cboLPI_Brand.AddItem ""
    For Each oRaws In colRaws
        Me.cboLPI_Brand.AddItem oRaws.Item(1)
    Next oRaws
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
        GetLPI_Brand = False
        Call HandleError(True, typErrInfo, , moAppLog)
    End If
End Function

'================================================================================
' function: GetLPI_PinCount()
'--------------------------------------------------------------------------------
' Description:  <Type your function description here...>
'--------------------------------------------------------------------------------
' Author:       Jack Hsieh, MXIC 2016/06/03 for 2016 MES Phase-5 Project (FT eRunCard Paperless Phase-2 Project).
'--------------------------------------------------------------------------------
' Parameters:
'................................................................................
' ARGUMENTS
'   moapplog            (I) [Optional] Valid reference to a Object object
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
Private Function GetLPI_PinCount(ByVal sCarrierType As String, ByVal sBoxingSpecNo As String, _
                                 ByVal sBrand As String) As Boolean
On Error GoTo ExitHandler:
Dim sProcID As String
Dim typErrInfo As tErrInfo
Dim sTable As String
Dim sColumn As String
Dim sWhere As String
Dim sOther  As String
Dim colRaws As Collection
Dim oRaws As FwStrings
'----
' Init
'----
    sProcID = "GetLPI_PinCount"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog)
    GetLPI_PinCount = True
    Me.cboLPI_PinCount.Clear
'----
' Condition Chking
'----


'----
' Action
'----
    'select distinct PINCOUNT from TBL_FT_Label_Boxing_INFO where CARRIERTYPE = 'Tape & Reel'
    '   and BOXINGSPECNO = '4040-0028' and BRAND = 'MX'
    sTable = gsCAT_TBL_FT_LABEL_BOXING_INFO

    sColumn = "distinct " & gsCAT_TFLBI_PINCOUNT
    
    sWhere = gsCAT_TFLBI_CARRIERTYPE & "='" & sCarrierType & "' and " & _
             gsCAT_TFLBI_BOXINGSPECNO & "='" & sBoxingSpecNo & "' and " & _
             gsCAT_TFLBI_BRAND & "='" & sBrand & "' and " & _
             gsCAT_TFLBI_DELETEFLAG & "='N'"

    Set colRaws = SelectRawSqlTxn(moAppLog, moFwWIP, moFwWF, moCwMbx, "CAT", sTable, sColumn, sWhere, sOther)
    If colRaws Is Nothing Then
        Call RaiseError(glERR_INVALIDOBJECT, _
                    FormatErrorText(gsETX_INVALIDOBJECT, "Collection"))
    End If
    
    Me.cboLPI_PinCount.AddItem ""
    For Each oRaws In colRaws
        Me.cboLPI_PinCount.AddItem oRaws.Item(1)
    Next oRaws
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
        GetLPI_PinCount = False
        Call HandleError(True, typErrInfo, , moAppLog)
    End If
End Function

'================================================================================
' function: GetLPI_PackageCode()
'--------------------------------------------------------------------------------
' Description:  <Type your function description here...>
'--------------------------------------------------------------------------------
' Author:       Jack Hsieh, MXIC 2016/06/03 for 2016 MES Phase-5 Project (FT eRunCard Paperless Phase-2 Project).
'--------------------------------------------------------------------------------
' Parameters:
'................................................................................
' ARGUMENTS
'   moapplog            (I) [Optional] Valid reference to a Object object
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
Private Function GetLPI_PackageCode(ByVal sCarrierType As String, ByVal sBoxingSpecNo As String, _
                                    ByVal sBrand As String, ByVal sPincount As String) As Boolean
On Error GoTo ExitHandler:
Dim sProcID As String
Dim typErrInfo As tErrInfo
Dim sTable As String
Dim sColumn As String
Dim sWhere As String
Dim sOther  As String
Dim colRaws As Collection
Dim oRaws As FwStrings
'----
' Init
'----
    sProcID = "GetLPI_PackageCode"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog)
    GetLPI_PackageCode = True
    Me.cboLPI_PackageCode.Clear
'----
' Condition Chking
'----


'----
' Action
'----
    'select distinct PACKAGECODE from TBL_FT_Label_Boxing_INFO where CARRIERTYPE = 'Tape & Reel'
    '   and BOXINGSPECNO = '4040-0028' and BRAND = 'MX' and PINCOUNT = 'XXXX'
    sTable = gsCAT_TBL_FT_LABEL_BOXING_INFO

    sColumn = "distinct " & gsCAT_TFLBI_PACKAGECODE
    
    sWhere = gsCAT_TFLBI_CARRIERTYPE & "='" & sCarrierType & "' and " & _
             gsCAT_TFLBI_BOXINGSPECNO & "='" & sBoxingSpecNo & "' and " & _
             gsCAT_TFLBI_BRAND & "='" & sBrand & "' and " & _
             gsCAT_TFLBI_PINCOUNT & "='" & sPincount & "' and " & _
             gsCAT_TFLBI_DELETEFLAG & "='N'"

    Set colRaws = SelectRawSqlTxn(moAppLog, moFwWIP, moFwWF, moCwMbx, "CAT", sTable, sColumn, sWhere, sOther)
    If colRaws Is Nothing Then
        Call RaiseError(glERR_INVALIDOBJECT, _
                    FormatErrorText(gsETX_INVALIDOBJECT, "Collection"))
    End If
    
    Me.cboLPI_PackageCode.AddItem ""
    For Each oRaws In colRaws
        Me.cboLPI_PackageCode.AddItem oRaws.Item(1)
    Next oRaws
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
        GetLPI_PackageCode = False
        Call HandleError(True, typErrInfo, , moAppLog)
    End If
End Function


'Added by Jack on 2016/06/03 for 2016 MES Phase-5 Project (FT eRunCard Paperless Phase-2 Project)
Private Sub cboLPI_LabelSpecNO_Click()
    cboLPI_CarrierType.ListIndex = -1
    cboLPI_CarrierType.Clear
    
    If cboLPI_LabelSpecNO.Text <> "" Then
        Call GetLPI_CarrierType(cboLPI_LabelSpecNO.Text)
    End If
End Sub

'Added by Jack on 2016/06/03 for 2016 MES Phase-5 Project (FT eRunCard Paperless Phase-2 Project)
Private Sub cboLPI_CarrierType_click()
    cboLPI_BoxingSpecNO.ListIndex = -1
    cboLPI_BoxingSpecNO.Clear
    
    If cboLPI_CarrierType.Text <> "" Then
        Call GetLPI_BoxingSpecNO(cboLPI_CarrierType.Text)
    End If
End Sub

'Added by Jack on 2016/06/03 for 2016 MES Phase-5 Project (FT eRunCard Paperless Phase-2 Project)
Private Sub cboLPI_BoxingSpecNO_click()
    cboLPI_Brand.ListIndex = -1
    cboLPI_Brand.Clear
    
    If cboLPI_CarrierType.Text <> "" And cboLPI_BoxingSpecNO.Text <> "" Then
        Call GetLPI_Brand(cboLPI_CarrierType.Text, cboLPI_BoxingSpecNO.Text)
    End If
End Sub

'Added by Jack on 2016/06/03 for 2016 MES Phase-5 Project (FT eRunCard Paperless Phase-2 Project)
Private Sub cboLPI_Brand_click()
    cboLPI_PinCount.ListIndex = -1
    cboLPI_PinCount.Clear
    
    If cboLPI_CarrierType.Text <> "" And cboLPI_BoxingSpecNO.Text <> "" And cboLPI_Brand.Text <> "" Then
        Call GetLPI_PinCount(cboLPI_CarrierType.Text, cboLPI_BoxingSpecNO.Text, cboLPI_Brand.Text)
    End If
End Sub

'Added by Jack on 2016/06/03 for 2016 MES Phase-5 Project (FT eRunCard Paperless Phase-2 Project)
Private Sub cboLPI_PinCount_click()
    cboLPI_PackageCode.ListIndex = -1
    cboLPI_PackageCode.Clear
    
    If cboLPI_CarrierType.Text <> "" And cboLPI_BoxingSpecNO.Text <> "" And cboLPI_Brand.Text <> "" _
       And cboLPI_PinCount.Text <> "" Then
        Call GetLPI_PackageCode(cboLPI_CarrierType.Text, cboLPI_BoxingSpecNO.Text, cboLPI_Brand.Text, cboLPI_PinCount.Text)
    End If
End Sub


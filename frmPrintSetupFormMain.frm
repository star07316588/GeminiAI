VERSION 5.00
Begin VB.Form frmPrintSetupFormMain 
   Caption         =   "PrintSetupFormMain"
   ClientHeight    =   3195
   ClientLeft      =   5700
   ClientTop       =   4395
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   Begin VB.CommandButton cmdChooseMachine 
      Caption         =   "FT-FVI"
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
      Index           =   1
      Left            =   1125
      TabIndex        =   3
      Top             =   1575
      Width           =   2415
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
      Left            =   1575
      TabIndex        =   2
      Top             =   2700
      Width           =   1590
   End
   Begin VB.Frame fraPrintSetupForm 
      Caption         =   "PrintSetupForm"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2115
      Left            =   225
      TabIndex        =   0
      Top             =   300
      Width           =   4215
      Begin VB.CommandButton cmdChooseMachine 
         Caption         =   "E-Test"
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
         Index           =   0
         Left            =   900
         TabIndex        =   1
         Top             =   675
         Width           =   2415
      End
   End
End
Attribute VB_Name = "frmPrintSetupFormMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Const msMODULE_ID As String = "frmPrintSetupFormMain"
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

Private Const miETest   As Integer = 0
Private Const miFTFVI   As Integer = 1

Public Property Let LotID(sLotID As String)
    msLotId = sLotID
End Property
Public Property Let EqpId(sEqpid As String)
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
Private Sub ResetForm()
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

Private Sub cmdCancel_Click()
On Error Resume Next
    miResult = vbCancel
    Me.Hide
End Sub
'================================================================================
' Sub: cmdChooseMachine_Click()
'--------------------------------------------------------------------------------
' Description: Open correlate form from user's choice.<M200804014>
'--------------------------------------------------------------------------------
' Author: Tony Chang ,MXIC on 2008/04/09
'================================================================================
Private Sub cmdChooseMachine_Click(Index As Integer)
On Error GoTo ExitHandler:
Dim sProcID         As String
Dim typErrInfo      As tErrInfo
Dim mfrmMain        As Form
Dim sRuleName       As String
Dim sEqID           As String


'----
' Init
'----
    sProcID = "cmdChooseMachine_Click"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog) '"Entering Function...", oLogCtrl, glLOG_PROC, msMODULE_ID, sProcID)

'----
' Condition Checking
'----


'----
' Action
'----
    Select Case (Index)
        Case miETest:
            sRuleName = "PrintSetupForm"
            sEqID = msEqpId
            Set mfrmMain = New frmPrintSetupForm
        Case miFTFVI:
            sRuleName = "PrintFTFVISetupForm"
            sEqID = msEqpId
            Set mfrmMain = New frmPrintFtFVISetupForm
    End Select
    
    Load mfrmMain
    With mfrmMain
       Set .FwMDLControl = moFwMDL
       Set .FwWIPControl = moFwWIP
       Set .FwOPRControl = moFwOPR
       Set .FwPRPControl = moFwPRP
       Set .FwWFControl = moFwWF
       Set .FwCATControl = moFwCAT
       Set .MainTraceLog = moAppLog
       Set .CwMbxControl = moCwMbx
       Set .ProRawSqlControl = moProRawSql
           .RuleName = sRuleName
           .LotID = msLotId
           .EqpId = sEqID
           .Init
           Screen.MousePointer = vbDefault
           Me.Hide
           .Show vbModal
           
    End With


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
    Set mfrmMain = Nothing
ErrorHandler:
    If typErrInfo.lErrNumber Then
        typErrInfo.sUserText = "Fail to execute application, please call IT support!!" & vbCrLf & _
                                "程式執行失敗, 請洽IT人員處理"
        '---- Start of Select Case Block ----
        On Error GoTo ExitHandler:
        Call HandleError(False, typErrInfo, True, moAppLog)
    End If

End Sub
'================================================================================
' Sub: Init()
'--------------------------------------------------------------------------------
' Description:<M200804014>
'--------------------------------------------------------------------------------
' Author: Tony Chang ,MXIC on 2008/04/09
'================================================================================
Public Function Init() As Boolean
On Error GoTo ExitHandler:
Dim sProcID         As String
Dim typErrInfo      As tErrInfo


'----
' Init
'----
    sProcID = "Init"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog) '"Entering Function...", oLogCtrl, glLOG_PROC, msMODULE_ID, sProcID)
    

'----
' Condition Checking
'----


'----
' Action
'----

    '待User提Laser需求
    
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
        Init = False
        On Error GoTo ExitHandler:
        Call HandleError(False, typErrInfo, , moAppLog, True)
    End If
End Function




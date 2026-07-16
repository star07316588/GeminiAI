VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmRunCard 
   Caption         =   "RunCard"
   ClientHeight    =   6390
   ClientLeft      =   2730
   ClientTop       =   1860
   ClientWidth     =   8280
   LinkTopic       =   "Form1"
   ScaleHeight     =   6390
   ScaleWidth      =   8280
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox txtCount 
      BackColor       =   &H80000004&
      Height          =   330
      Left            =   1035
      Locked          =   -1  'True
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   5940
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
      Left            =   5895
      TabIndex        =   7
      Top             =   5985
      Width           =   1095
   End
   Begin VB.Frame fraList 
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
      Height          =   3465
      Left            =   135
      TabIndex        =   10
      Top             =   2385
      Width           =   8040
      Begin FPSpreadADO.fpSpread spdList 
         Height          =   2985
         Left            =   135
         TabIndex        =   5
         Top             =   315
         Width           =   7755
         _Version        =   196613
         _ExtentX        =   13679
         _ExtentY        =   5265
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
         MaxCols         =   5
         MaxRows         =   5
         RetainSelBlock  =   0   'False
         SpreadDesigner  =   "frmRunCard.frx":0000
         UserResize      =   2
      End
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
      Height          =   2115
      Left            =   135
      TabIndex        =   8
      Top             =   45
      Width           =   8040
      Begin VB.TextBox txtStep 
         Height          =   285
         Left            =   2475
         TabIndex        =   2
         Top             =   1215
         Width           =   1680
      End
      Begin VB.TextBox txtIPN 
         Height          =   285
         Left            =   2475
         TabIndex        =   0
         Top             =   315
         Width           =   2490
      End
      Begin VB.TextBox txtLotID 
         Height          =   285
         Left            =   2475
         TabIndex        =   3
         Top             =   1620
         Width           =   2490
      End
      Begin VB.ComboBox cboType 
         Height          =   315
         ItemData        =   "frmRunCard.frx":03E1
         Left            =   2475
         List            =   "frmRunCard.frx":03E3
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   720
         Width           =   2490
      End
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
         Height          =   330
         Left            =   6210
         TabIndex        =   4
         Top             =   1620
         Width           =   1455
      End
      Begin VB.Label lblStepName 
         BorderStyle     =   1  'Fixed Single
         Height          =   285
         Left            =   4275
         TabIndex        =   16
         Top             =   1215
         Width           =   1455
      End
      Begin VB.Label lblType 
         AutoSize        =   -1  'True
         Caption         =   "RunCard Type:"
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
         TabIndex        =   13
         Top             =   765
         Width           =   1290
      End
      Begin VB.Label lblStep 
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
         Left            =   1800
         TabIndex        =   12
         Top             =   1215
         Width           =   465
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
         Left            =   1710
         TabIndex        =   11
         Top             =   1665
         Width           =   540
      End
      Begin VB.Label lblPNo 
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
         Left            =   1845
         TabIndex        =   9
         Top             =   315
         Width           =   390
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
      Height          =   315
      Left            =   4410
      TabIndex        =   6
      Top             =   5985
      Width           =   1095
   End
   Begin VB.Label lblCount 
      AutoSize        =   -1  'True
      Caption         =   "Count:"
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
      Left            =   360
      TabIndex        =   15
      Top             =   6030
      Width           =   570
   End
End
Attribute VB_Name = "frmRunCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const msMODULE_ID As String = "frmRunCard"
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

Private Const msFilePath        As String = "\system32\FwClient\Others\runcard.XLT"

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
'================================================================================
' Sub: cmdOk_Click()
'--------------------------------------------------------------------------------
' Description:  <Type your Sub description here...>
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
'----
' Action
'----
    Dim oLot As FwLot
    Dim appXL As Object
    Dim iPosTitle As Integer
    
    With Me.spdList
        .Row = 1
        
        Set appXL = CreateObject("Excel.Application")
        If appXL Is Nothing Then
            UtShowMsgBox "Excel not found"
            GoTo ExitHandler
        End If
                
        'for debug
        If moFwWF.ClientID = "cit-vencent" Then appXL.Visible = True
        
        appXL.cursor = 2
        appXL.DisplayAlerts = False
        appXL.Workbooks.Add sFilePath
        appXL.ActiveWindow.WindowState = vbMaximized

        iPosTitle = 1
        
        Do While .Row <= .MaxRows
            .Col = 1
            If .Text = "1" Then
                .Col = 2
                Set oLot = FwuRetrieveLot(moFwWIP, .Text, moAppLog)
                If Not oLot Is Nothing Then
                    If Me.cboType = "FT" Then
                        Call FtExcelCell(appXL, oLot, iPosTitle)
                    Else
                        Call WsExcelCell(appXL, oLot, iPosTitle)
                    End If
                End If
            End If
            .Row = .Row + 1
        Loop
        
    End With
    
    appXL.Sheets("WS").Select
    appXL.ActiveWindow.SelectedSheets(1).Visible = False
    appXL.Sheets("FT").Select
    appXL.ActiveWindow.SelectedSheets(1).Visible = False
    appXL.Sheets(1).Select
'    appXL.ActiveSheet.PageSetup.LeftMargin = appXL.Application.InchesToPoints(0.62992125984252)
'    appXL.ActiveSheet.PageSetup.RightMargin = appXL.Application.InchesToPoints(0.590551181102362)

    appXL.range("A1").Select
    appXL.cursor = 1
    appXL.DisplayAlerts = True
    
    ChDir "C:\FASTech"
    appXL.ActiveWorkbook.SaveAs "C:\FASTech\" & Me.cboType & Format(Now, "YYYYMMDDHHMM") & ".xls", _
        , "", "", False, False

    appXL.Visible = True
'    appXL.Quit
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
        cmdOk.Enabled = True
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
Public Function init() As Boolean
On Error GoTo ExitHandler:
Dim sProcID As String
Dim typErrInfo As tErrInfo

'----
' Init
'----
    sProcID = "Init"
    init = True
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog) '"Entering Function...", oLogCtrl, glLOG_PROC, msMODULE_ID, sProcID)

'----
' Condition Checking
'----
    ' <Put your condition checking codes here>...

'----
' Action
'----
    ' <Put your Action codes here>...
    Dim sSQL As String, colSQLResult As Collection
    Dim iIdx As Integer
    Dim oLot As FwLot
    
'    Me.cboType.AddItem ""
    Me.cboType.AddItem "FT"
    Me.cboType.AddItem "WS"
    Me.cboType.Text = "FT"
    
    
    Me.spdList.MaxRows = 0
    Me.cmdOk.Enabled = False
    
    If msLotId <> "" Then
        Me.txtLotID = msLotId
        Set oLot = FwuRetrieveLot(moFwWIP, msLotId, moAppLog)
        If oLot.CurrentStep.Steps.Item(1).Attributes(gsSTEP_CUSTOMATTR_STAGE) = "FT" Then
            Me.cboType.ListIndex = 0
        Else
            Me.cboType.ListIndex = 1
        End If
        Set oLot = Nothing
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
        init = False
        On Error GoTo ExitHandler:
        Call HandleError(True, typErrInfo, , moAppLog)
    End If
End Function


'================================================================================
' Sub: ExecQuery()
'--------------------------------------------------------------------------------
' Description:  <Type your Sub description here...>
'--------------------------------------------------------------------------------
' Author:       Vencent Wei, CIT 2002/06/11
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

'----
' Action
'----
    ' <Put your Action codes here>...
    Dim sSQL As String, colSQLResult As Collection
    Dim iIdx As Integer
    Dim oLot As FwLot
    
    Me.spdList.MaxRows = 0
    Me.txtCount = "0"
    
    sSQL = " select A.APPID, C.STEPNAME, C.DESCRIPTION, A.COMPONENTQTY, E.VALDATA IPN" & _
           " from FWLOT A, FWLOT_N2M B, FWWIPSTEP C, FWWIPSTEP_PN2M D, FWLOT_PN2M E" & _
           " where A.PROCESSINGSTATUS in ('Active', 'Hold') and " & _
                 " A.SYSID = B.FROMID and " & _
                 " B.LINKNAME = 'stepCollection' and" & _
                 " B.TOID = C.SYSID and " & _
                 " C.SYSID = D.FROMID and " & _
                 " D.LINKNAME = 'attributes' and" & _
                 " D.KEYDATA = '" & gsSTEP_CUSTOMATTR_STAGE & "' and " & _
                 " A.SYSID = E.FROMID and" & _
                 " E.KEYDATA = '" & gsLOT_CUSTOMATTR_IPN & "'"
             
    If Me.txtStep <> "" Then
        sSQL = sSQL & " and C.STEPNAME = '" & Left(Me.txtStep, 5) & "' "
    End If
    
    If txtLotID <> "" Then
        sSQL = sSQL & " and A.APPID like '" & Replace(txtLotID, "*", "%") & "' "
    End If

    If Me.cboType = "FT" Then
        sSQL = sSQL & " and D.VALDATA = '" & gsSTAGE_FT & "' "
    ElseIf Me.cboType = "WS" Then
        sSQL = sSQL & " and D.VALDATA = '" & gsSTAGE_WS & "' "
    End If
    
    If Me.txtIPN <> "" Then
        sSQL = sSQL & " and E.valdata like '" & Replace(Me.txtIPN, "*", "%") & "%' "
    End If

    sSQL = sSQL & " order by A.appid "
    Set colSQLResult = moProRawSql.QueryDatabase(sSQL)
    If colSQLResult Is Nothing Then
        Call RaiseError(glERR_RAWSQLERROR, _
                    FormatErrorText(gsETX_RAWSQLERROR, "Select", "Fwlot"))
    ElseIf colSQLResult.Count > 0 Then
        With Me.spdList
            For iIdx = 1 To colSQLResult.Count
                .MaxRows = .MaxRows + 1
                .SetText 2, .MaxRows, colSQLResult.Item(iIdx).Item(1)
                .SetText 1, .MaxRows, 1
                .SetText 3, .MaxRows, colSQLResult.Item(iIdx).Item(5)
                .SetText 4, .MaxRows, colSQLResult.Item(iIdx).Item(4)
                .SetText 5, .MaxRows, colSQLResult.Item(iIdx).Item(2) & " " & colSQLResult.Item(iIdx).Item(3)

            Next iIdx
            
            Me.txtCount = .MaxRows
            Me.cmdOk.Enabled = True
        End With
        
        ResizeSpdColumn Me.spdList, moAppLog

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

Private Sub spdList_ButtonClicked(ByVal Col As Long, ByVal Row As Long, ByVal ButtonDown As Integer)
    If Col = 1 Then
        With Me.spdList
            .Col = Col
            .Row = Row
            If Not IsNumeric(Me.txtCount) Then Me.txtCount = 0
            If .Text = "1" Then
                Me.txtCount = CLng(Me.txtCount) + 1
                If CLng(Me.txtCount) > .MaxRows Then Me.txtCount = .MaxRows
                If CLng(Me.txtCount) > 0 Then Me.cmdOk.Enabled = True
            Else
                Me.txtCount = CLng(Me.txtCount) - 1
                If CLng(Me.txtCount) < 0 Then Me.txtCount = 0
                If CLng(Me.txtCount) = 0 Then Me.cmdOk.Enabled = False
            End If
        End With
    End If
End Sub

'================================================================================
' Sub: FtExcelCell()
'--------------------------------------------------------------------------------
' Description:  <Type your Sub description here...>
'--------------------------------------------------------------------------------
' Author:       Vencent Wei, CIT 2002/06/12
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
Private Sub FtExcelCell(ByRef appXL As Object, ByRef oLot As FwLot, _
                      ByRef iPosTitle As Integer)
On Error GoTo ExitHandler:
Dim sProcID             As String
Dim typErrInfo          As tErrInfo
Dim sIncludEngLot       As String
Dim sIncludeReworkLot   As String
Dim sLotType            As String
Dim sFutActType         As String
Dim sPackageCode        As String 'Add by Tony on 20070817 requirement:M200708010

'----
' Init
'----
    sProcID = "FtExcelCell"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog)

'----
' Condition Checking
'----
    ' <Put your condition checking codes here>...

'----
' Action
'----
    ' <Put your Action codes here>...
    Dim oPlan As FwProcessPlan, oAggrSteps As FwAggregateSteps, oAggrStep As FwAggregateStep
    Dim sSQL As String, colSQLResult As Collection, colSQL2 As Collection
    Dim iIdx As Integer, iIdx2 As Integer
    Dim iPosProcess As Integer, iPosComment As Integer, iPosNext As Integer
    Dim sHandler As String, sMergeID As String, sSplitID As String
    Dim iQtyOut As Long, iQtyIn As Long
    Dim sCurEqp As String, sSubSys1 As String, sSubSys2 As String
    Dim sLotOwner As String
    sLotOwner = oLot.CustomAttributes(gsLOT_CUSTOMATTR_LOT_OWNER)
    
    iPosProcess = iPosTitle + 12 + 7
    iPosComment = iPosTitle + 26 + 7
    With appXL
        If iPosTitle = 1 Then
            .Sheets("FT").Select
            .Cells.Select
            .selection.Copy
            .Sheets(1).Select
            .ActiveSheet.Paste
            .selection.NumberFormatLocal = "@"
            .Application.CutCopyMode = False
        Else
            .Sheets("FT").Select
            .Rows("1:36").Select
            .selection.NumberFormatLocal = "@"
            .selection.Copy
            .Sheets(1).Select
            .range("A" & CStr(iPosTitle)).Select
            .selection.Insert -4121
            .Application.CutCopyMode = False
            .range("A" & CStr(iPosTitle)).Select
            .ActiveWindow.SelectedSheets.HPageBreaks.Add appXL.activeCell

        End If
        .range("Q" & CStr(iPosTitle + 1)).Select
        .selection.horizontalalignment = -4131
        .activeCell.FormulaR1C1 = "Date: " & Format(Now, "YYYY/MM/DD HH:MM:SS")
        
        .range("E" & CStr(iPosTitle + 3)).Select
        .activeCell.FormulaR1C1 = oLot.CustomAttributes(gsLOT_CUSTOMATTR_IPN)
        .range("E" & CStr(iPosTitle + 4)).Select
        .activeCell.FormulaR1C1 = oLot.Id
        .range("E" & CStr(iPosTitle + 5)).Select
        .activeCell.FormulaR1C1 = CStr(oLot.ComponentQuantity)
        .range("E" & CStr(iPosTitle + 6)).Select
        .activeCell.FormulaR1C1 = oLot.CustomAttributes(gsLOT_CUSTOMATTR_LOT_OWNER)
        .range("Q" & CStr(iPosTitle + 3)).Select
        '.ActiveCell.FormulaR1C1 = olot.PlanId      ' -- remarked and modifyed by Haley on 2003/05/05 for PiLot Run bug No.33
        .activeCell.FormulaR1C1 = oLot.CustomAttributes(gsLOT_CUSTOMATTR_ROUTE)
        
        'Add Tbl_ipn_master.PackageCode by Tony on 2007/08/17 requirement:M200708010
        sSQL = " select A." & gsCAT_TIM_EPN & ", " _
                    & " A." & gsCAT_TIM_CHECKSUM & ", " _
                    & " A." & gsCAT_TIM_SECURITY & ", " _
                    & " A." & gsCAT_TIM_MARKING_SPEC_NO & ", " _
                    & " A." & gsCAT_TIM_BOXING_SPEC_NO & ", " _
                    & " A." & gsCAT_TIM_BOXING_TYPE & ", " _
                    & " A." & gsCAT_TIM_CARRIER_SPEC_NO & "," _
                    & " A." & gsCAT_TIM_ICDRAWING & "," _
                    & " A." & gsCAT_TIM_BRAND & "," _
                    & " A." & gsCAT_TIM_PACKAGE_CODE & "," _
                    & " B." & gsCAT_TPBS_CPN & ", " _
                    & " B." & gsCAT_TPBS_BAKE_TIME & " ||' '|| " _
                    & " B." & gsCAT_TPBS_BAKE_TEMP & " BAKE, " _
                    & " B." & gsCAT_TPBS_CUSTOMER_NO & "," _
                    & " B." & gsCAT_TPBS_CARRIER_TYPE & "," _
                    & " B." & gsCAT_TPBS_CARRIER_QTY & "," _
                    & " C." & gsCAT_TCUM_CUSTOMERNAME & "," _
                    & " C." & gsCAT_TCUM_LABELSPECNO & " LABEL"
                    
        sSQL = sSQL & ",A." & gsCAT_TIM_GRADE  'Add by Sam on 20131016 for ReqNo:JC201300301 ,增加IPN.GRADE
        
        sSQL = sSQL _
             & " from " & gsCAT_TBL_IPN_MASTER & " A, " _
                        & gsCAT_TBL_PRM_BE_SPEC & " B, " _
                        & gsCAT_TBL_CUSTOMER_MASTER & " C " _
             & " where A." & gsCAT_TIM_IPN & " = '" & oLot.CustomAttributes(gsLOT_CUSTOMATTR_IPN) & "' " _
             & " and A." & gsCAT_TIM_IPN & " = B." & gsCAT_TPBS_IPN _
             & " and B." & gsCAT_TPBS_CUSTOMER_NO & " = C." & gsCAT_TCUM_CUSTOMERNO & "(+)" _
             & " order by decode(B." & gsCAT_TPBS_DEFAULTS & ", 'Y', 0, 1)"
        Debug.Print sSQL
        Set colSQLResult = moProRawSql.QueryDatabase(sSQL)
        
        If colSQLResult.Count > 0 Then
            .range("E" & CStr(iPosTitle + 7)).Select
            .activeCell.FormulaR1C1 = colSQLResult.Item(1).Item(gsCAT_TIM_EPN)
            .range("E" & CStr(iPosTitle + 8)).Select
            .activeCell.FormulaR1C1 = colSQLResult.Item(1).Item(gsCAT_TPBS_CPN)
            .range("E" & CStr(iPosTitle + 9)).Select
            .activeCell.FormulaR1C1 = colSQLResult.Item(1).Item(gsCAT_TIM_CHECKSUM)
            .range("E" & CStr(iPosTitle + 10)).Select
            .activeCell.FormulaR1C1 = colSQLResult.Item(1).Item(gsCAT_TIM_SECURITY)
            .range("E" & CStr(iPosTitle + 11)).Select
            'Modify by Janus, 20050926, M200509017;by Tony 200708017,M200708010
            '.activecell.formulaR1C1 = colSQLResult.Item(1).Item("LABEL")
            sPackageCode = colSQLResult.Item(1).Item(gsCAT_TIM_PACKAGE_CODE)
            sPackageCode = UCase(Left(sPackageCode, 1))
            
            If Len(Trim(colSQLResult.Item(1).Item("LABEL"))) > 0 Then
               .activeCell.FormulaR1C1 = colSQLResult.Item(1).Item("LABEL")
            Else
                'Add by Sam start on 20131016 for ReqNo:JC201300301
               If Trim(colSQLResult.Item(1).Item("BRAND")) = "KH" And Trim(colSQLResult.Item(1).Item(gsCAT_TIM_GRADE)) = "Y" And _
                  sPackageCode <> "W" And sPackageCode <> "H" Then
                  .activeCell.FormulaR1C1 = "6130K-0807.1"
                'Add by Sam end on 20131016 for ReqNo:JC201300301
               ElseIf Trim(colSQLResult.Item(1).Item("BRAND")) = "KH" And _
                sPackageCode <> "W" And sPackageCode <> "H" Then
                  .activeCell.FormulaR1C1 = "6130K-0807"
               Else
                  .activeCell.FormulaR1C1 = "6130-0807"
               End If
            End If
            'End Modify By Janus, 20050926;by Tony 200708017,M200708010
            .range("E" & CStr(iPosTitle + 12)).Select
            .activeCell.FormulaR1C1 = colSQLResult.Item(1).Item(gsCAT_TIM_CARRIER_SPEC_NO)
            .range("E" & CStr(iPosTitle + 13)).Select
            .activeCell.FormulaR1C1 = colSQLResult.Item(1).Item(gsCAT_TPBS_CARRIER_TYPE)
            .range("E" & CStr(iPosTitle + 14)).Select
            .activeCell.FormulaR1C1 = colSQLResult.Item(1).Item(gsCAT_TPBS_CARRIER_QTY)
            .range("E" & CStr(iPosTitle + 15)).Select                                'Add by Jeff on 040915, Requirement : 200407004
            .activeCell.FormulaR1C1 = colSQLResult.Item(1).Item(gsCAT_TIM_ICDRAWING) 'Add by Jeff on 040915, Requirement : 200407004
            
            .range("Q" & CStr(iPosTitle + 4)).Select
            .activeCell.FormulaR1C1 = colSQLResult.Item(1).Item("BAKE")
            .range("Q" & CStr(iPosTitle + 5)).Select
            .activeCell.FormulaR1C1 = colSQLResult.Item(1).Item(gsCAT_TPBS_CUSTOMER_NO) & " " & colSQLResult.Item(1).Item(gsCAT_TCUM_CUSTOMERNAME)
            .range("Q" & CStr(iPosTitle + 6)).Select
            .activeCell.FormulaR1C1 = colSQLResult.Item(1).Item(gsCAT_TIM_BOXING_SPEC_NO)
            .range("Q" & CStr(iPosTitle + 7)).Select
            .activeCell.FormulaR1C1 = colSQLResult.Item(1).Item(gsCAT_TIM_BOXING_TYPE)
            .range("Q" & CStr(iPosTitle + 8)).Select
            .activeCell.FormulaR1C1 = colSQLResult.Item(1).Item(gsCAT_TIM_MARKING_SPEC_NO)
        End If
        
        'add by vencent, for ch031105-3
        sSQL = "select B." & gsCAT_TSP_MARKING_SPEC_NO & ",B." & gsCAT_TSP_MARKING_SPEC_REV & "," _
                        & gsCAT_TSP_TOP_LINE1 & ",B." & gsCAT_TSP_TOP_LINE2 & ",B." _
                        & gsCAT_TSP_TOP_LINE3 & ",B." & gsCAT_TSP_TOP_LINE4 & ",B." _
                        & gsCAT_TSP_TOP_LINE5 & ",B." & gsCAT_TSP_TOP_LINE6 & ",B." _
                        & gsCAT_TSP_TOP_LINE7 & ",B." & gsCAT_TSP_TOP_LINE8 & ",B." _
                        & gsCAT_TSP_TOP_LINE9 & ",B." & gsCAT_TSP_TOP_LINE10 & ",B." _
                        & gsCAT_TSP_BACK_SIDE1 & ",B." & gsCAT_TSP_BACK_SIDE2 & ",B." _
                        & gsCAT_TSP_BACK_SIDE3
        sSQL = sSQL & _
               " from " & gsCAT_TBL_LOT_INFO & " A," & _
                          gsCAT_TBL_SUB_PO & " B" & _
               " where A." & gsCAT_TLI_LOT_ID & " = '" & oLot.Id & "' " & _
                 " and A." & gsCAT_TLI_SASM_PONO & " = B." & gsCAT_TSP_PO_NO & _
                 " and A." & gsCAT_TLI_SASM_POITEM & " = B." & gsCAT_TSP_ITEM & _
                 " order by B." & gsCAT_TSP_PO_DATE_TIME & " desc"
        Set colSQLResult = moProRawSql.QueryDatabase(sSQL)
        
        If colSQLResult.Count > 0 Then
            .range("E" & CStr(iPosTitle + 16)).Select
            .activeCell.FormulaR1C1 = colSQLResult.Item(1).Item(gsCAT_TSP_BACK_SIDE1)
            .range("E" & CStr(iPosTitle + 17)).Select
            .activeCell.FormulaR1C1 = colSQLResult.Item(1).Item(gsCAT_TSP_BACK_SIDE2)
            .range("E" & CStr(iPosTitle + 18)).Select
            .activeCell.FormulaR1C1 = colSQLResult.Item(1).Item(gsCAT_TSP_BACK_SIDE3)
            
            .range("Q" & CStr(iPosTitle + 9)).Select
            .activeCell.FormulaR1C1 = colSQLResult.Item(1).Item(gsCAT_TSP_TOP_LINE1)
            .range("Q" & CStr(iPosTitle + 10)).Select
            .activeCell.FormulaR1C1 = colSQLResult.Item(1).Item(gsCAT_TSP_TOP_LINE2)
            .range("Q" & CStr(iPosTitle + 11)).Select
            .activeCell.FormulaR1C1 = colSQLResult.Item(1).Item(gsCAT_TSP_TOP_LINE3)
            .range("Q" & CStr(iPosTitle + 12)).Select
            .activeCell.FormulaR1C1 = colSQLResult.Item(1).Item(gsCAT_TSP_TOP_LINE4)
            .range("Q" & CStr(iPosTitle + 13)).Select
            .activeCell.FormulaR1C1 = colSQLResult.Item(1).Item(gsCAT_TSP_TOP_LINE5)
            .range("Q" & CStr(iPosTitle + 14)).Select
            .activeCell.FormulaR1C1 = colSQLResult.Item(1).Item(gsCAT_TSP_TOP_LINE6)
            .range("Q" & CStr(iPosTitle + 15)).Select
            .activeCell.FormulaR1C1 = colSQLResult.Item(1).Item(gsCAT_TSP_TOP_LINE7)
            .range("Q" & CStr(iPosTitle + 16)).Select
            .activeCell.FormulaR1C1 = colSQLResult.Item(1).Item(gsCAT_TSP_TOP_LINE8)
            .range("Q" & CStr(iPosTitle + 17)).Select
            .activeCell.FormulaR1C1 = colSQLResult.Item(1).Item(gsCAT_TSP_TOP_LINE9)
            .range("Q" & CStr(iPosTitle + 18)).Select
            .activeCell.FormulaR1C1 = colSQLResult.Item(1).Item(gsCAT_TSP_TOP_LINE10)
        End If
        'end of add
        
        'Lot History - Step
'        sSql = " select B.STEPNAME, B.DESCRIPTION, A.QUANTITYOUT, " _
'             & " A.TRACKINTIME, A.TXNTIMESTAMP, A.USERID USEROUT, A.LASTTRACKOUT, " _
'             & " C.USERID USERIN, C.QUANTITYIN " _
'             & " from FWTRACKOUT A, FWSTEPVERSION B, FWTRACKIN C " _
'             & " where A.WIPID = '" & olot.Id & "' " _
'             & " and A.OUTSTEP = B.STEPNAME || '.' || B.REVISION " _
'             & " and A.OUTSTEP = C.INSTEP(+) " _
'             & " and A.TRACKINTIME = C.TXNTIMESTAMP(+) " _
'             & " and A.WIPID = C.WIPID(+)" _
'             & " order by A.TXNTIMESTAMP "
 
        '改以TBL_ROUTE_STEP_LIST排序
        sSQL = "select F." & gsCAT_TRSL_STEP_NO & " STEPNAME," & _
                     " F." & gsCAT_TRSL_STEP_NAME & " DESCRIPTION," & _
                     " A.QUANTITYOUT," & _
                     " A.TRACKINTIME," & _
                     " A.TXNTIMESTAMP," & _
                     " A.USERID USEROUT," & _
                     " A.LASTTRACKOUT," & _
                     " B.USERID USERIN," & _
                     " B.QUANTITYIN"
                     
        sSQL = sSQL & _
               " from FWTRACKOUT A, " & _
                    " FWTRACKIN B, " & _
                    " FWWIPHISTORY C," & _
                    " FWWIPHISTORY_N2M D," & _
                    " FWWIPSTEPHISTORY E," & _
                    gsCAT_TBL_ROUTE_STEP_LIST & " F," & _
                    " (select n.handle, n.StepSeq" & _
                      " from FwFlatNode n, FwFlatPlan p, FwStepVersion s " & _
                      " where p.SysId = n.PlanSysId and " & _
                            " p.revstate = 'Frozen' and " & _
                            " p.PlanName = '" & oLot.PlanId & "' and" & _
                            " p.PlanVersion = '" & oLot.PlanVersion & "' and" & _
                            " s.SysId = n.NodeRef and" & _
                            " n.NodeType = 'S') G" & _
               " where A.WIPID = '2F048700A0' and" & _
                     " A.WIPID = B.WIPID(+) and" & _
                     " A.TRACKINTIME = B.TXNTIMESTAMP(+) and" & _
                     " A.SYSID = C.WIPTXN and" & _
                     " C.SYSID = D.FROMID and" & _
                     " D.TOID = E.SYSID and" & _
                     " E.HANDLE = G.StepSeq and" & _
                     " G.handle = F." & gsCAT_TRSL_STEP_HANDLE & " and " & _
                     " F." & gsCAT_TRSL_ROUTE_ID & " = '" & oLot.PlanId & "'"
                      
        sSQL = sSQL & " order by F." & gsCAT_TRSL_STEP_SEQ
       
        Set colSQLResult = moProRawSql.QueryDatabase(sSQL)
        
        If colSQLResult.Count > 0 Then
            'Copy
            For iIdx = 1 To colSQLResult.Count - 1
                .Rows(CStr(iPosProcess) & ":" & CStr(iPosComment - 1)).Select
                .selection.Copy
                .range("A" & CStr(iPosComment)).Select
                .selection.Insert -4121
                iPosComment = iPosComment + 13
            Next iIdx
            .Application.CutCopyMode = False
            
            'Fill cell
            For iIdx = 1 To colSQLResult.Count
'                .Range("E" & CStr(iPosTitle + iIDx * 13)).Select
                .range("E" & CStr(iPosProcess + 1 + (iIdx - 1) * 13)).Select
                .activeCell.FormulaR1C1 = colSQLResult.Item(iIdx).Item("DESCRIPTION")
'                .Range("Q" & CStr(iPosTitle + iIDx * 13)).Select
                .range("Q" & CStr(iPosProcess + 1 + (iIdx - 1) * 13)).Select
                .activeCell.FormulaR1C1 = colSQLResult.Item(iIdx).Item("STEPNAME")
                .range("E" & CStr(iPosProcess + 2 + (iIdx - 1) * 13)).Select
                .activeCell.FormulaR1C1 = Format(ConvertToDate(colSQLResult.Item(iIdx).Item("TRACKINTIME"), moAppLog), "YYYY/MM/DD HH:MM:SS")
                .range("Q" & CStr(iPosProcess + 2 + (iIdx - 1) * 13)).Select
                .activeCell.FormulaR1C1 = Format(ConvertToDate(colSQLResult.Item(iIdx).Item("TXNTIMESTAMP"), moAppLog), "YYYY/MM/DD HH:MM:SS")
                .range("E" & CStr(iPosProcess + 3 + (iIdx - 1) * 13)).Select
                .activeCell.FormulaR1C1 = colSQLResult.Item(iIdx).Item("USERIN")
                .range("Q" & CStr(iPosProcess + 3 + (iIdx - 1) * 13)).Select
                .activeCell.FormulaR1C1 = colSQLResult.Item(iIdx).Item("USEROUT")
                .range("E" & CStr(iPosProcess + 4 + (iIdx - 1) * 13)).Select
                .activeCell.FormulaR1C1 = colSQLResult.Item(iIdx).Item("QUANTITYIN")
                .range("Q" & CStr(iPosProcess + 4 + (iIdx - 1) * 13)).Select
                .activeCell.FormulaR1C1 = colSQLResult.Item(iIdx).Item("QUANTITYOUT")
                
                iQtyIn = 0
                iQtyOut = 0
                If IsNumeric(colSQLResult.Item(iIdx).Item("QUANTITYIN")) Then iQtyIn = CLng(colSQLResult.Item(iIdx).Item("QUANTITYIN"))
                If IsNumeric(colSQLResult.Item(iIdx).Item("QUANTITYOUT")) Then iQtyIn = CLng(colSQLResult.Item(iIdx).Item("QUANTITYOUT"))
                
                sCurEqp = GetLotAttr(oLot.Id, gsLOT_CUSTOMATTR_CUR_EQP_ID, colSQLResult.Item(iIdx).Item("TXNTIMESTAMP"))
                .range("Q" & CStr(iPosProcess + 7 + (iIdx - 1) * 13)).Select 'Q20
                .activeCell.FormulaR1C1 = sCurEqp
                .range("Q" & CStr(iPosProcess + 8 + (iIdx - 1) * 13)).Select
                
'                If colSqlResult.Item(iIdx).Item("SubSys1") = "" Then
'                    sHandler = colSqlResult.Item(iIdx).Item("SubSys2")
'                ElseIf colSqlResult.Item(iIdx).Item("SubSys2") = "" Then
'                    sHandler = colSqlResult.Item(iIdx).Item("SubSys1")
'                Else
'                    sHandler = colSqlResult.Item(iIdx).Item("SubSys1") & ";" & colSqlResult.Item(iIdx).Item("SubSys2")
'                End If
                If sCurEqp <> "" Then
                    sSubSys1 = GetEqpAttr(sCurEqp, gsEQP_CUSTOMATTR_SUBSYS1, colSQLResult.Item(iIdx).Item("TXNTIMESTAMP"))
                    sSubSys2 = GetEqpAttr(sCurEqp, gsEQP_CUSTOMATTR_SUBSYS2, colSQLResult.Item(iIdx).Item("TXNTIMESTAMP"))
                    If sSubSys1 = "" Then
                        sHandler = sSubSys2
                    ElseIf sSubSys2 = "" Then
                        sHandler = sSubSys1
                    Else
                        sHandler = sSubSys1 & ";" & sSubSys2
                    End If
                    .activeCell.FormulaR1C1 = sHandler
                End If
                
                .range("Q" & CStr(iPosProcess + 9 + (iIdx - 1) * 13)).Select 'Q22
'                .ActiveCell.FormulaR1C1 = colSqlResult.Item(iIdx).Item("PgName") & " " _
'                                        & colSqlResult.Item(iIdx).Item("Temperature") & " " _
'                                        & colSqlResult.Item(iIdx).Item("ProcessTime") & " " _
'                                        & colSqlResult.Item(iIdx).Item("ProcessUnit")
                .activeCell.FormulaR1C1 = GetLotAttr(oLot.Id, gsLOT_CUSTOMATTR_PGNAME, colSQLResult.Item(iIdx).Item("TXNTIMESTAMP")) & _
                                          GetLotAttr(oLot.Id, gsLOT_CUSTOMATTR_TEMPERATURE, colSQLResult.Item(iIdx).Item("TXNTIMESTAMP")) & _
                                          GetLotAttr(oLot.Id, gsLOT_CUSTOMATTR_PROCESS_HOUR, colSQLResult.Item(iIdx).Item("TXNTIMESTAMP")) & _
                                          GetLotAttr(oLot.Id, gsLOT_CUSTOMATTR_PROCESSTIMEUNIT, colSQLResult.Item(iIdx).Item("TXNTIMESTAMP"))

                If colSQLResult.Item(iIdx).Item("DESCRIPTION") Like "FT*" _
                Or colSQLResult.Item(iIdx).Item("DESCRIPTION") Like "TQAE*" Then
                    sSQL = " select PASSQTY, FAILQTY, BIN1, BIN2, BIN3, BIN4, BIN5, BIN6 " _
                         & " from TBL_FT_LOTCOMPLETE " _
                         & " where LOTID = '" & oLot.Id & "' " _
                         & " and STEPNO = '" & colSQLResult.Item(iIdx).Item("STEPNAME") & "' " _
                         & " and TESTERID = '" & sCurEqp & "' " _
                         & " and DELETEFLAG <> 'Y' "
                    Set colSQL2 = moProRawSql.QueryDatabase(sSQL)
                    
                    If colSQL2.Count > 0 Then
                        .range("E" & CStr(iPosProcess + 5 + (iIdx - 1) * 13)).Select 'E18
                        .activeCell.FormulaR1C1 = colSQL2.Item(1).Item("BIN1")
                        .range("E" & CStr(iPosProcess + 6 + (iIdx - 1) * 13)).Select
                        .activeCell.FormulaR1C1 = colSQL2.Item(1).Item("BIN2")
                        .range("E" & CStr(iPosProcess + 7 + (iIdx - 1) * 13)).Select
                        .activeCell.FormulaR1C1 = colSQL2.Item(1).Item("BIN3")
                        .range("E" & CStr(iPosProcess + 8 + (iIdx - 1) * 13)).Select
                        .activeCell.FormulaR1C1 = colSQL2.Item(1).Item("BIN4")
                        .range("E" & CStr(iPosProcess + 9 + (iIdx - 1) * 13)).Select
                        .activeCell.FormulaR1C1 = colSQL2.Item(1).Item("BIN5")
                        .range("E" & CStr(iPosProcess + 10 + (iIdx - 1) * 13)).Select
                        .activeCell.FormulaR1C1 = colSQL2.Item(1).Item("BIN6")
                        
                        .range("Q" & CStr(iPosProcess + 4 + (iIdx - 1) * 13)).Select 'Q17
                        .activeCell.FormulaR1C1 = colSQL2.Item(1).Item("PASSQTY")
                        .range("Q" & CStr(iPosProcess + 5 + (iIdx - 1) * 13)).Select
                        .activeCell.FormulaR1C1 = colSQL2.Item(1).Item("FAILQTY")
                        
                        If IsNumeric(colSQL2.Item(1).Item("PASSQTY")) Then
                            iQtyOut = CLng(colSQL2.Item(1).Item("PASSQTY"))
                        End If
                    End If
                    
                    If iQtyIn <> 0 Or iQtyOut <> 0 Then
                        .range("Q" & CStr(iPosProcess + 6 + (iIdx - 1) * 13)).Select 'Q19
                        .activeCell.FormulaR1C1 = Round(iQtyOut / iQtyIn, 2)
                    End If
                    
                    If colSQLResult.Item(iIdx).Item("QUANTITYIN") _
                     > colSQLResult.Item(iIdx).Item("QUANTITYOUT") Then
                        sSQL = " select B.COMMENTCODE, B.BRIEFDESCRIPTION " _
                             & " from FWSCRAPLOT A, FWCOMMENT B " _
                             & " where A.WIPID = '" & oLot.Id & "' " _
                             & " and A.TXNTIMESTAMP <= '" & colSQLResult.Item(iIdx).Item("TXNTIMESTAMP") & "' " _
                             & " and A.TXNTIMESTAMP >= '" & colSQLResult.Item(iIdx).Item("TRACKINTIME") & "' " _
                             & " and A.TXNCOMMENT = B.SYSID "
                             
                        Set colSQL2 = moProRawSql.QueryDatabase(sSQL)
                        If colSQL2.Count > 0 Then
                            .range("Q" & CStr(iPosProcess + 10 + (iIdx - 1) * 13)).Select
                            .activeCell.FormulaR1C1 = colSQL2.Item(1).Item("BRIEFDESCRIPTION")
                        End If
                        
                    End If
                    
                    ' for Merge ID
                    sSQL = " select B.VALDATA " _
                         & " from FWMERGE A, FWMERGE_PN2M B " _
                         & " where A.WIPID = '" & oLot.Id & "' " _
                         & " and A.TXNTIMESTAMP <= '" & colSQLResult.Item(iIdx).Item("TXNTIMESTAMP") & "' " _
                         & " and A.TXNTIMESTAMP >= '" & colSQLResult.Item(iIdx).Item("TRACKINTIME") & "' " _
                         & " and B.LINKNAME = 'childLotCollection' " _
                         & " and A.SYSID = B.FROMID "
                         
                    Set colSQL2 = moProRawSql.QueryDatabase(sSQL)
                    If colSQL2.Count > 0 Then
                        sMergeID = ""
                        For iIdx2 = 1 To colSQL2.Count
                            sMergeID = sMergeID & colSQL2.Item(iIdx2).Item(1) & ";"
                        Next iIdx2
                        If Right(sMergeID, 1) = ";" Then sMergeID = Left(sMergeID, Len(sMergeID) - 1)
                        .range("E" & CStr(iPosProcess + 11 + (iIdx - 1) * 13)).Select 'E24
                        .activeCell.FormulaR1C1 = sMergeID
                    End If
                    
                    ' for Split ID
                    sSQL = " select C.CHILDLOTID " _
                         & " from FWSPLITLOT A, FWSPLITLOT_N2M B, FWSPLIT C " _
                         & " where A.WIPID = '" & oLot.Id & "' " _
                         & " and A.TXNTIMESTAMP <= '" & colSQLResult.Item(iIdx).Item("TXNTIMESTAMP") & "' " _
                         & " and A.TXNTIMESTAMP >= '" & colSQLResult.Item(iIdx).Item("TRACKINTIME") & "' " _
                         & " and A.SYSID = B.FROMID " _
                         & " and B.LINKNAME = 'splitCollection' " _
                         & " and B.TOID = C.SYSID "
                         
                    Set colSQL2 = moProRawSql.QueryDatabase(sSQL)
                    If colSQL2.Count > 0 Then
                        sSplitID = ""
                        For iIdx2 = 1 To colSQL2.Count
                            sSplitID = sSplitID & colSQL2.Item(iIdx2).Item(1) & ";"
                        Next iIdx2
                        If Right(sSplitID, 1) = ";" Then sSplitID = Left(sSplitID, Len(sSplitID) - 1)
                        .range("E" & CStr(iPosProcess + 12 + (iIdx - 1) * 13)).Select
                        .activeCell.FormulaR1C1 = sSplitID
                    End If
                    
                End If
            Next iIdx
        End If
        
        ' Curr Step & Next Steps
        .Sheets("FT").Select
        .Rows("21:33").Select
        .selection.NumberFormatLocal = "@"
        .selection.Copy
        .Sheets(1).Select
        .range("A" & CStr(iPosComment)).Select
        
        .selection.Insert -4121
        .Application.CutCopyMode = False
        iPosProcess = iPosComment
        iPosComment = iPosComment + 13
        
        .range("E" & CStr(iPosProcess)).Select
        .activeCell.FormulaR1C1 = oLot.CurrentStep.Steps.Item(1).Description
        .range("Q" & CStr(iPosProcess)).Select
        .activeCell.FormulaR1C1 = oLot.CurrentStep.Steps.Item(1).Id
        
        If oLot.CurrentStep.Steps.Item(1).AvailablePaths.Count > 0 Then
            
            If oLot.CurrentStep.Steps.Item(1).AvailablePaths.Count = 1 _
            And oLot.CurrentStep.Steps.Item(1).AvailablePaths.Item(1) <> "PASS" Then GoTo FillCommentHandler
            
            Set oPlan = moFwPRP.PlanById(oLot.PlanId, oLot.PlanVersion)
            If oPlan Is Nothing Then GoTo FillCommentHandler
            
            Set oAggrSteps = oPlan.Steps(oLot.CurrentStep.Steps.Item(1).Handle, "+999")
            For Each oAggrStep In oAggrSteps
                .Sheets("FT").Select
                .Rows("21:33").Select
                .selection.Copy
                .Sheets(1).Select
                .range("A" & CStr(iPosComment)).Select
                .selection.Insert -4121
                
                iPosProcess = iPosProcess + 13
                iPosComment = iPosComment + 13
                .range("E" & CStr(iPosProcess)).Select
                .activeCell.FormulaR1C1 = oAggrStep.Steps.Item(1).Description
                .range("Q" & CStr(iPosProcess)).Select
                .activeCell.FormulaR1C1 = oAggrStep.Steps.Item(1).Id
            Next
            Set oPlan = Nothing
            Set oAggrSteps = Nothing
            Set oAggrStep = Nothing
        End If
                        
FillCommentHandler:
    'add by Nelson start on 2005/11/30 for ReqNo:M200505002
    sLotType = GetLotType(oLot.Id, moProRawSql, moAppLog)
    'add by Nelson end on 2005/11/30 for ReqNo:M200505002
    'Comment
    sSQL = " select * from ( " _
         & " select " & gsCAT_TLFA_STEP_NO & " Step, " & gsCAT_TLFA_COMMENTS & ", " _
                      & gsCAT_TLFA_USER_ID & ", " & gsCAT_TLFA_SET_TIME & " sTime,'' IncludeEngLot,'' IncludeReworkLot,'Lot' FutActType " _
         & " From " & gsCAT_TBL_LOT_FUTACT _
         & " where " & gsCAT_TLFA_LOT_ID & " = '" & oLot.Id & "' " _
         & " and " & gsCAT_TLFA_COMMENTS & " is not null " _
         & " and " & gsCAT_TLFA_DELETE_FLAG & " <> 'Y' " _
         & " Union " _
         & " select " & gsCAT_TIFA_STEP_NO & " Step, " & gsCAT_TIFA_COMMENTS & ", " _
                      & gsCAT_TIFA_USER_ID & ", " & gsCAT_TIFA_SET_TIME & " sTime, " & gsCAT_TIFA_INCLUDEENGLOT & "," & gsCAT_TIFA_INCLUDEREWORKLOT & ",'Ipn' FutActType" _
         & " From " & gsCAT_TBL_IPN_FUTACT _
         & " where " & gsCAT_TIFA_IPN & " = '" & oLot.CustomAttributes(gsLOT_CUSTOMATTR_IPN) & "' " _
         & " and " & gsCAT_TIFA_COMMENTS & " is not null " _
         & " and " & gsCAT_TIFA_DELETE_FLAG & " <> 'Y' "
    
    '-----------------------------------------------------------------------------------
    '2008/03/12 Vencent M200802020, add IncludeLotOwner
    sSQL = sSQL & " and (nvl(" & gsCAT_TIFA_INCLUDELOTOWNER & ", 'All')='All' or " & _
            gsCAT_TIFA_INCLUDELOTOWNER & " = '" & sLotOwner & "' or " & _
            gsCAT_TIFA_INCLUDELOTOWNER & " like '" & sLotOwner & ",%' or " & _
            gsCAT_TIFA_INCLUDELOTOWNER & " like '%," & sLotOwner & "' or " & _
            gsCAT_TIFA_INCLUDELOTOWNER & " like '%," & sLotOwner & ",%')"
    '-----------------------------------------------------------------------------------
         
    'add by Nelson start on 2005/11/24 for ReqNo:M200505002
    sSQL = sSQL & " Union " _
         & " select " & gsCAT_TPGFA_STEP_NO & " Step, " & gsCAT_TPGFA_COMMENTS & ", " _
                      & gsCAT_TPGFA_USER_ID & ", " & gsCAT_TPGFA_SET_TIME & " sTime, " & gsCAT_TPGFA_INCLUDEENGLOT & "," & gsCAT_TPGFA_INCLUDEREWORKLOT & ",'ProdGroup' FutActType" _
         & " From " & gsCAT_TBL_PRODGROUP_FUTACT _
         & " where " & gsCAT_TPGFA_PROD_GROUP & " = '" & oLot.CustomAttributes(gsLOT_CUSTOMATTR_PROD_GROUP) & "' " _
         & " and " & gsCAT_TPGFA_COMMENTS & " is not null " _
         & " and " & gsCAT_TPGFA_DELETE_FLAG & " <>'Y' "
    'add by Nelson end on 2005/11/24 for ReqNo:M200505002
    
    '-----------------------------------------------------------------------------------
    '2008/03/12 Vencent M200802020, add IncludeLotOwner
    sSQL = sSQL & " and (nvl(" & gsCAT_TPGFA_INCLUDELOTOWNER & ", 'All')='All' or " & _
            gsCAT_TPGFA_INCLUDELOTOWNER & " = '" & sLotOwner & "' or " & _
            gsCAT_TPGFA_INCLUDELOTOWNER & " like '" & sLotOwner & ",%' or " & _
            gsCAT_TPGFA_INCLUDELOTOWNER & " like '%," & sLotOwner & "' or " & _
            gsCAT_TPGFA_INCLUDELOTOWNER & " like '%," & sLotOwner & ",%')"
    '-----------------------------------------------------------------------------------
    
    sSQL = sSQL & " Union " _
         & " select " & gsCAT_THR_STEP_ID & " Step, " & gsCAT_THR_COMMENTS & ", " _
                      & gsCAT_THR_CREATE_USER_ID & ", " & gsCAT_THR_CREATE_TIME & " sTime,'' IncludeEngLot,'' IncludeReworkLot,'' FutActType" _
         & " From " & gsCAT_TBL_HOLD_REC _
         & " where " & gsCAT_THR_LOT_ID & " = '" & oLot.Id & "' " _
         & " and " & gsCAT_THR_COMMENTS & " is not null " _
         & " and " & gsCAT_THR_DELETE_FLAG & " <> 'Y' " _
         & " Union " _
         & " select " & gsCAT_TMT_STEPNO & " Step, " & gsCAT_TMT_COMMENTS & ", " _
                      & gsCAT_TMT_CREATE_USER_ID & ", " & gsCAT_TMT_CREATE_TIME & " sTime,'' IncludeEngLot,'' IncludeReworkLot,'' FutActType" _
         & " From " & gsCAT_TBL_MANUAL_TESTQTY _
         & " where " & gsCAT_TMT_LOT_ID & " = '" & oLot.Id & "' " _
         & " and " & gsCAT_TMT_COMMENTS & " is not null " _
         & " and " & gsCAT_TMT_DELETE_FLAG & " <> 'Y' " _
         & " ) order by sTime "
         
        Set colSQL2 = moProRawSql.QueryDatabase(sSQL)
        
        If colSQL2.Count > 0 Then
            iIdx2 = 1
            For iIdx = 1 To colSQL2.Count
                sIncludEngLot = colSQL2.Item(iIdx).Item(5)
                sIncludeReworkLot = colSQL2.Item(iIdx).Item(6)
                sFutActType = colSQL2.Item(iIdx).Item(7)
                If sFutActType = "Ipn" Or sFutActType = "ProdGroup" Then
                    Select Case sLotType
                        Case "Normal"
                            .range("C" & CStr(iPosComment + iIdx2 + 1)).Select
                            .activeCell.FormulaR1C1 = colSQL2.Item(iIdx).Item(1)
                            .range("Q" & CStr(iPosComment + iIdx2 + 1)).Select
                            .activeCell.FormulaR1C1 = colSQL2.Item(iIdx).Item(2)
                            .range("K" & CStr(iPosComment + iIdx2 + 1)).Select
                            .activeCell.FormulaR1C1 = colSQL2.Item(iIdx).Item(3)
                            .range("F" & CStr(iPosComment + iIdx2 + 1)).Select
                            .activeCell.FormulaR1C1 = Format(ConvertToDate(colSQL2.Item(iIdx).Item(4), moAppLog), "YYYY/MM/DD HH:MM:SS")
                            iIdx2 = iIdx2 + 1
                        Case "Eng"
                            If sIncludEngLot = "Y" Then
                                .range("C" & CStr(iPosComment + iIdx2 + 1)).Select
                                .activeCell.FormulaR1C1 = colSQL2.Item(iIdx).Item(1)
                                .range("Q" & CStr(iPosComment + iIdx2 + 1)).Select
                                .activeCell.FormulaR1C1 = colSQL2.Item(iIdx).Item(2)
                                .range("K" & CStr(iPosComment + iIdx2 + 1)).Select
                                .activeCell.FormulaR1C1 = colSQL2.Item(iIdx).Item(3)
                                .range("F" & CStr(iPosComment + iIdx2 + 1)).Select
                                .activeCell.FormulaR1C1 = Format(ConvertToDate(colSQL2.Item(iIdx).Item(4), moAppLog), "YYYY/MM/DD HH:MM:SS")
                                iIdx2 = iIdx2 + 1
                            End If
                        Case "Rework"
                            If sIncludeReworkLot = "Y" Then
                                .range("C" & CStr(iPosComment + iIdx2 + 1)).Select
                                .activeCell.FormulaR1C1 = colSQL2.Item(iIdx).Item(1)
                                .range("Q" & CStr(iPosComment + iIdx2 + 1)).Select
                                .activeCell.FormulaR1C1 = colSQL2.Item(iIdx).Item(2)
                                .range("K" & CStr(iPosComment + iIdx2 + 1)).Select
                                .activeCell.FormulaR1C1 = colSQL2.Item(iIdx).Item(3)
                                .range("F" & CStr(iPosComment + iIdx2 + 1)).Select
                                .activeCell.FormulaR1C1 = Format(ConvertToDate(colSQL2.Item(iIdx).Item(4), moAppLog), "YYYY/MM/DD HH:MM:SS")
                                iIdx2 = iIdx2 + 1
                            End If
                        Case Else
                    End Select
                Else
                    .range("C" & CStr(iPosComment + iIdx2 + 1)).Select
                    .activeCell.FormulaR1C1 = colSQL2.Item(iIdx).Item(1)
                    .range("Q" & CStr(iPosComment + iIdx2 + 1)).Select
                    .activeCell.FormulaR1C1 = colSQL2.Item(iIdx).Item(2)
                    .range("K" & CStr(iPosComment + iIdx2 + 1)).Select
                    .activeCell.FormulaR1C1 = colSQL2.Item(iIdx).Item(3)
                    .range("F" & CStr(iPosComment + iIdx2 + 1)).Select
                    .activeCell.FormulaR1C1 = Format(ConvertToDate(colSQL2.Item(iIdx).Item(4), moAppLog), "YYYY/MM/DD HH:MM:SS")
                    iIdx2 = iIdx2 + 1
                End If
            Next iIdx
        End If
        iPosTitle = iPosComment + 3 + colSQL2.Count
    
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
' Sub: WsExcelCell()
'--------------------------------------------------------------------------------
' Description:  <Type your Sub description here...>
'--------------------------------------------------------------------------------
' Author:       Vencent Wei, CIT 2002/06/13
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
Private Sub WsExcelCell(ByRef appXL As Object, _
                        ByRef oLot As FwLot, _
                        ByRef iPosTitle As Integer)
On Error GoTo ExitHandler:
Dim sProcID             As String
Dim typErrInfo          As tErrInfo
Dim sIncludEngLot       As String
Dim sIncludeReworkLot   As String
Dim sLotType            As String
Dim sFutActType         As String
'----
' Init
'----
    sProcID = "WsExcelCell"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog)

'----
' Condition Checking
'----
    ' <Put your condition checking codes here>...

'----
' Action
'----
    ' <Put your Action codes here>...
    Dim oPlan As FwProcessPlan, oAggrSteps As FwAggregateSteps, oAggrStep As FwAggregateStep
    Dim sSQL As String, colSQLResult As Collection, colSQL2 As Collection
    Dim iIdx As Integer, iIdx2 As Integer
    Dim iPosProcess As Integer, iPosComment As Integer, iPosNext As Integer
    Dim sTotalID As String, iTotalQty As Long, sId As String, sUser As String, sTester As String
    Dim sLotOwner As String
    sLotOwner = oLot.CustomAttributes(gsLOT_CUSTOMATTR_LOT_OWNER)
    
    iPosProcess = iPosTitle + 9
    iPosComment = iPosTitle + 20
    
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
            .Rows("1:28").Select
            .selection.NumberFormatLocal = "@"
            .selection.Copy
            .Sheets(1).Select
            .range("A" & CStr(iPosTitle)).Select
            .selection.Insert -4121
            .Application.CutCopyMode = False
            .range("A" & CStr(iPosTitle)).Select
            .ActiveWindow.SelectedSheets.HPageBreaks.Add appXL.activeCell

        End If
        
        .range("Q" & CStr(iPosTitle + 1)).Select
        .selection.horizontalalignment = -4131
        .activeCell.FormulaR1C1 = "Date: " & Format(Now, "YYYY/MM/DD HH:MM:SS")
        
        .range("E" & CStr(iPosTitle + 3)).Select
        .activeCell.FormulaR1C1 = oLot.CustomAttributes(gsLOT_CUSTOMATTR_IPN)
        .range("E" & CStr(iPosTitle + 4)).Select
        .activeCell.FormulaR1C1 = oLot.Id
        .range("E" & CStr(iPosTitle + 5)).Select
        .activeCell.FormulaR1C1 = CStr(oLot.CustomAttributes(gsLOT_CUSTOMATTR_WAFER_QTY))
        .range("E" & CStr(iPosTitle + 6)).Select
        .activeCell.FormulaR1C1 = CStr(oLot.CustomAttributes(gsLOT_CUSTOMATTR_CHIP_QTY))
        .range("E" & CStr(iPosTitle + 7)).Select
        .activeCell.FormulaR1C1 = oLot.CustomAttributes(gsLOT_CUSTOMATTR_LOT_OWNER)
        
        .range("Q" & CStr(iPosTitle + 3)).Select
        '.activecell.formulaR1C1 = olot.PlanId      '-- remarked and modified by Haley on 2003/05/05 for Pilot Run bug No.33
        .activeCell.FormulaR1C1 = oLot.CustomAttributes(gsLOT_CUSTOMATTR_ROUTE)
       
        .range("Q" & CStr(iPosTitle + 4)).Select
        .activeCell.FormulaR1C1 = oLot.StartDate
        .range("Q" & CStr(iPosTitle + 5)).Select
        .activeCell.FormulaR1C1 = oLot.CustomAttributes(gsLOT_CUSTOMATTR_STATUS)
        
        sSQL = " select " & gsCAT_TIM_GROSS_DIE _
             & " from " & gsCAT_TBL_IPN_MASTER _
             & " where " & gsCAT_TIM_IPN & " = '" & oLot.CustomAttributes(gsLOT_CUSTOMATTR_IPN) & "' "
        
        Set colSQLResult = moProRawSql.QueryDatabase(sSQL)
        
        If colSQLResult.Count > 0 Then
            .range("Q" & CStr(iPosTitle + 6)).Select
            .activeCell.FormulaR1C1 = colSQLResult.Item(1).Item(1)
'            .Range("E" & CStr(iPosTitle + 15)).Select
'            .activecell.formulaR1C1 = colSQLResult.Item(1).Item(2)
        End If
        
        'Lot History - Step
'        sSql = " select B.STEPNAME, B.DESCRIPTION, A.QUANTITYOUT, " _
'             & " A.TRACKINTIME, A.TXNTIMESTAMP, A.USERID USEROUT, A.LASTTRACKOUT, " _
'             & " C.USERID USERIN, C.QUANTITYIN " _
'             & " from FWTRACKOUT A, FWSTEPVERSION B, FWTRACKIN C " _
'             & " where A.WIPID = '" & oLot.Id & "' " _
'             & " and A.OUTSTEP = B.STEPNAME || '.' || B.REVISION " _
'             & " and A.OUTSTEP = C.INSTEP(+) " _
'             & " and A.TRACKINTIME = C.TXNTIMESTAMP(+) " _
'             & " and A.WIPID = C.WIPID(+) "

        '改以TBL_ROUTE_STEP_LIST排序
        sSQL = "select F." & gsCAT_TRSL_STEP_NO & " STEPNAME," & _
                     " F." & gsCAT_TRSL_STEP_NAME & " DESCRIPTION," & _
                     " A.QUANTITYOUT," & _
                     " A.TRACKINTIME," & _
                     " A.TXNTIMESTAMP," & _
                     " A.USERID USEROUT," & _
                     " A.LASTTRACKOUT," & _
                     " B.USERID USERIN," & _
                     " B.QUANTITYIN"
                     
        sSQL = sSQL & _
               " from FWTRACKOUT A, " & _
                    " FWTRACKIN B, " & _
                    " FWWIPHISTORY C," & _
                    " FWWIPHISTORY_N2M D," & _
                    " FWWIPSTEPHISTORY E," & _
                    gsCAT_TBL_ROUTE_STEP_LIST & " F," & _
                    " (select n.handle, n.StepSeq" & _
                      " from FwFlatNode n, FwFlatPlan p, FwStepVersion s " & _
                      " where p.SysId = n.PlanSysId and " & _
                            " p.revstate = 'Frozen' and " & _
                            " p.PlanName = '" & oLot.PlanId & "' and" & _
                            " p.PlanVersion = '" & oLot.PlanVersion & "' and" & _
                            " s.SysId = n.NodeRef and" & _
                            " n.NodeType = 'S') G" & _
               " where A.WIPID = '" & oLot.Id & "' and" & _
                     " A.WIPID = B.WIPID(+) and" & _
                     " A.TRACKINTIME = B.TXNTIMESTAMP(+) and" & _
                     " A.SYSID = C.WIPTXN and" & _
                     " C.SYSID = D.FROMID and" & _
                     " D.TOID = E.SYSID and" & _
                     " E.HANDLE = G.StepSeq and" & _
                     " G.handle = F." & gsCAT_TRSL_STEP_HANDLE & " and " & _
                     " F." & gsCAT_TRSL_ROUTE_ID & " = '" & oLot.PlanId & "'"
                      
        sSQL = sSQL & " order by F." & gsCAT_TRSL_STEP_SEQ
        
        Debug.Print "HisSQL= " & sSQL
        Set colSQLResult = moProRawSql.QueryDatabase(sSQL)
        
        If colSQLResult.Count > 0 Then
            For iIdx = 1 To colSQLResult.Count
                'copy null field
                If iIdx >= 2 Then
                    .Sheets("WS").Select
                    .Rows("11:20").Select
                    .selection.NumberFormatLocal = "@"
                    .selection.Copy
                    .Sheets(1).Select
                    .range("A" & CStr(iPosComment)).Select
                    .selection.Insert -4121
                    .Application.CutCopyMode = False
                    iPosComment = iPosComment + 10
                End If
                
                .range("E" & CStr(iPosTitle + iIdx * 10)).Select
                .activeCell.FormulaR1C1 = colSQLResult.Item(iIdx).Item("DESCRIPTION")
                .range("M" & CStr(iPosTitle + iIdx * 10)).Select
                .activeCell.FormulaR1C1 = colSQLResult.Item(iIdx).Item("STEPNAME")
                .range("E" & CStr(iPosTitle + iIdx * 10 + 1)).Select
                .activeCell.FormulaR1C1 = Format(ConvertToDate(colSQLResult.Item(iIdx).Item("TRACKINTIME"), moAppLog), "YYYY/MM/DD HH:MM:SS")
                .range("M" & CStr(iPosTitle + iIdx * 10 + 1)).Select
                .activeCell.FormulaR1C1 = Format(ConvertToDate(colSQLResult.Item(iIdx).Item("TXNTIMESTAMP"), moAppLog), "YYYY/MM/DD HH:MM:SS")
                .range("E" & CStr(iPosTitle + iIdx * 10) + 3).Select
                .activeCell.FormulaR1C1 = colSQLResult.Item(iIdx).Item("QUANTITYOUT")
                
                .range("E" & CStr(iPosTitle + iIdx * 10) + 2).Select
                .activeCell.FormulaR1C1 = colSQLResult.Item(iIdx).Item("USEROUT")
                
                If IsNumeric(colSQLResult.Item(iIdx).Item("QUANTITYOUT")) _
                And IsNumeric(colSQLResult.Item(iIdx).Item("QUANTITYIN")) Then
                    .range("M" & CStr(iPosTitle + iIdx * 10) + 3).Select
                    .activeCell.FormulaR1C1 = CLng(colSQLResult.Item(iIdx).Item("QUANTITYIN")) - CLng(colSQLResult.Item(iIdx).Item("QUANTITYOUT"))
                    .range("U" & CStr(iPosTitle + iIdx * 10) + 3).Select
                    .activeCell.FormulaR1C1 = CStr(Round(CLng(colSQLResult.Item(iIdx).Item("QUANTITYOUT")) / CLng(colSQLResult.Item(iIdx).Item("QUANTITYIN")), 2) * 100) & "%"
                End If
                
                sSQL = " select " & gsCAT_TWTS_TESTER_ID & ", " _
                                  & gsCAT_TWTS_USER_ID & ", " _
                                  & gsCAT_TWTS_LOT_GOOD_DIE & ", " _
                                  & " nvl(" & gsCAT_TWTS_REVISE_GOOD_DIEN & "," & gsCAT_TWTS_TDS_GOOD_DIEN & ") GOODDIE," _
                                  & gsCAT_TWTS_WAFER_IDN _
                     & " from " & gsCAT_TBL_WS_TDS_SUM _
                     & " where " & gsCAT_TWTS_LOT_ID & " = '" & oLot.Id & "' " _
                     & " and " & gsCAT_TWTS_STEPNO & " = '" & colSQLResult.Item(iIdx).Item("STEPNAME") & "' " _
                     & " and " & gsCAT_TWTS_TIME_STAMP & " <= '" & colSQLResult.Item(iIdx).Item("TXNTIMESTAMP") & "'" _
                     & " and " & gsCAT_TWTS_DELETE_FLAG & " <> 'Y' " _
                     & " order by " & gsCAT_TWTS_USER_ID & ", " & gsCAT_TWTS_TESTER_ID & ", " & gsCAT_TWTS_WAFER_IDN & ", " & gsCAT_TWTS_TIME_STAMP
                
                iPosProcess = iPosTitle + iIdx * 10 + 2
                Debug.Print "SSQL = " & sSQL
                Set colSQL2 = moProRawSql.QueryDatabase(sSQL)
                If colSQL2.Count > 0 Then
                    iPosProcess = iPosTitle + iIdx * 10 + 2
                    For iIdx2 = 1 To colSQL2.Count
                        If sUser = "" Then
                            sUser = colSQL2.Item(iIdx2).Item(gsCAT_TWTS_USER_ID)
                            sTester = colSQL2.Item(iIdx2).Item(gsCAT_TWTS_TESTER_ID)
                            sId = colSQL2.Item(iIdx2).Item(gsCAT_TWTS_WAFER_IDN) & ";"
                            If iIdx = colSQL2.Count Then
                                .range("E" & CStr(iPosProcess)).Select
                                .activeCell.FormulaR1C1 = sUser
                                .range("M" & CStr(iPosProcess)).Select
                                .activeCell.FormulaR1C1 = sTester
                                .range("U" & CStr(iPosProcess)).Select
                                If Right(sId, 1) = ";" Then sId = Left(sId, Len(sId) - 1)
                                .activeCell.FormulaR1C1 = sId
                            End If
                        ElseIf sUser <> colSQL2.Item(iIdx2).Item(gsCAT_TWTS_USER_ID) _
                            Or sTester <> colSQL2.Item(iIdx2).Item(gsCAT_TWTS_TESTER_ID) Then
                                .range("E" & CStr(iPosProcess)).Select
                                .activeCell.FormulaR1C1 = sUser
                                .range("M" & CStr(iPosProcess)).Select
                                .activeCell.FormulaR1C1 = sTester
                                .range("U" & CStr(iPosProcess)).Select
                                If Right(sId, 1) = ";" Then sId = Left(sId, Len(sId) - 1)
                                .activeCell.FormulaR1C1 = sId
                                
                                sUser = colSQL2.Item(iIdx2).Item(gsCAT_TWTS_USER_ID)
                                sTester = colSQL2.Item(iIdx2).Item(gsCAT_TWTS_TESTER_ID)
                                sId = colSQL2.Item(iIdx2).Item(gsCAT_TWTS_WAFER_IDN) & ";"
                                
                                If iIdx2 <> colSQL2.Count Then
                                    '再複製一分 Oprtator/Testid/TestWaferID
                                    .Sheets("WS").Select
                                    .Rows("13").Select
                                    .selection.Copy
                                    .Sheets(1).Select
                                    .range("A" & CStr(iPosProcess + 1)).Select
                                    .selection.Insert -4121
                                    iPosProcess = iPosProcess + 1
                                    iPosComment = iPosComment + 1
                                    iPosTitle = iPosTitle + 1
                                End If
                        Else
                            'then same user, tester
                            sId = sId & colSQL2.Item(iIdx2).Item(gsCAT_TWTS_WAFER_IDN) & ";"
                            If iIdx = colSQL2.Count Then
                                .range("E" & CStr(iPosProcess)).Select
                                .activeCell.FormulaR1C1 = sUser
                                .range("M" & CStr(iPosProcess)).Select
                                .activeCell.FormulaR1C1 = sTester
                                .range("U" & CStr(iPosProcess)).Select
                                If Right(sId, 1) = ";" Then sId = Left(sId, Len(sId) - 1)
                                .activeCell.FormulaR1C1 = sId
                            End If
                        End If
                        
                        If IsNumeric(colSQL2.Item(iIdx2).Item(gsCAT_TWTS_WAFER_IDN)) Then
                            .range(Chr(64 + CInt(colSQL2.Item(iIdx2).Item(gsCAT_TWTS_WAFER_IDN))) _
                                & CStr(iPosProcess + 4)).Select
                            .selection.NumberFormatLocal = "G/通用格式"
                            .selection.Font.Size = 10
                            .activeCell.FormulaR1C1 = colSQL2.Item(iIdx2).Item("GOODDIE")
                        End If
                    Next iIdx2
                    'TDS Qty
                    .range("E" & CStr(iPosProcess + 5)).Select
                    .selection.NumberFormatLocal = "G/通用格式"
'                    .ActiveCell.FormulaR1C1 = "=SUM(R[-1]C[-4]:R[-1]C[20])"
                    .activeCell.FormulaR1C1 = colSQL2.Item(iIdx2).Item(gsCAT_TWTS_LOT_GOOD_DIE)
                    .range("E" & CStr(iPosProcess + 5) & ":" & "J" & CStr(iPosProcess + 5)).Select
                    .selection.Merge
                    .selection.horizontalalignment = -4131
                End If
                
                ' 目檢片
                sSQL = " select distinct " & gsCAT_TMT_WAFERID _
                     & " from " & gsCAT_TBL_MANUAL_TESTQTY _
                     & " where " & gsCAT_TMT_LOT_ID & " = '" & oLot.Id & "' " _
                     & " and " & gsCAT_TMT_STEPNO & " = '" & colSQLResult.Item(iIdx).Item("STEPNAME") & "' " _
                     & " and " & gsCAT_TMT_DELETE_FLAG & " <> 'Y' " _
                     & " order by " & gsCAT_TMT_WAFERID
                Set colSQL2 = moProRawSql.QueryDatabase(sSQL)
                If colSQL2.Count > 0 Then
                    sId = ""
                    For iIdx2 = 1 To colSQL2.Count
                        sId = sId & colSQL2.Item(iIdx2).Item(1) & ";"
                    Next iIdx2
                    If Right(sId, 1) = ";" Then sId = Left(sId, Len(sId) - 1)
                    .range("E" & CStr(iPosProcess + 6)).Select
                    .activeCell.FormulaR1C1 = sId
                End If
            Next iIdx
        End If
        
        '--------------------------------------------
        ' Curr Step & Next Steps
        .Sheets("WS").Select
        .Rows("11:20").Select
        .selection.NumberFormatLocal = "@"
        .selection.Copy
        .Sheets(1).Select
        .range("A" & CStr(iPosComment)).Select
        
        .selection.Insert -4121
        .Application.CutCopyMode = False
        iPosProcess = iPosComment
        iPosComment = iPosComment + 10
        
        .range("E" & CStr(iPosProcess)).Select
        .activeCell.FormulaR1C1 = oLot.CurrentStep.Steps.Item(1).Description
        .range("M" & CStr(iPosProcess)).Select
        .activeCell.FormulaR1C1 = oLot.CurrentStep.Steps.Item(1).Id
        
        'Next Steps
        Dim sCurStepSeq As String
        Dim sRoute As String
        If oLot.CustomAttributes(gsLOT_CUSTOMATTR_ROUTE) Is Nothing Then
            sRoute = GetStepPath(moAppLog, moFwWIP, moFwWF, moCwMbx, oLot.PlanId, oLot.CurrentStep.Steps.Item(1).Handle)
        Else
            sRoute = oLot.CustomAttributes(gsLOT_CUSTOMATTR_ROUTE)
        End If
        sSQL = "select " & gsCAT_TRSL_STEP_NO & "," & _
                           gsCAT_TRSL_STEP_NAME & _
               " from " & gsCAT_TBL_ROUTE_STEP_LIST & _
               " where " & gsCAT_TRSL_PATH & " = '" & sRoute & "' and " & _
                           gsCAT_TRSL_STEP_SEQ & " > " & _
               "(select max(" & gsCAT_TRSL_STEP_SEQ & ") " & _
               " from " & gsCAT_TBL_ROUTE_STEP_LIST & _
               " where " & gsCAT_TRSL_PATH & " = '" & sRoute & "' and " & _
                           gsCAT_TRSL_STEP_NO & " = '" & oLot.CurrentStep.Steps.Item(1).Id & "') " & _
               " order by " & gsCAT_TRSL_STEP_SEQ
               
        Set colSQL2 = moProRawSql.QueryDatabase(sSQL)
        If colSQL2.Count > 0 Then
            For iIdx = 1 To colSQL2.Count
                .Sheets("WS").Select
                .Rows("11:20").Select
                .selection.Copy
                .Sheets(1).Select
                .range("A" & CStr(iPosComment)).Select
                .selection.Insert -4121
                
                iPosProcess = iPosProcess + 10
                iPosComment = iPosComment + 10
                .range("E" & CStr(iPosProcess)).Select
                .activeCell.FormulaR1C1 = colSQL2.Item(iIdx).Item(1)
                .range("M" & CStr(iPosProcess)).Select
                .activeCell.FormulaR1C1 = colSQL2.Item(iIdx).Item(2)
                
            Next iIdx
        Else
            GoTo FillCommentHandler
        End If
        
'        If olot.CurrentStep.Steps.Item(1).AvailablePaths.Count > 0 Then
'            If olot.CurrentStep.Steps.Item(1).AvailablePaths.Count = 1 _
'            And olot.CurrentStep.Steps.Item(1).AvailablePaths.Item(1) <> "PASS" Then GoTo FillCommentHandler
'
'            Set oPlan = moFwPRP.PlanById(olot.PlanId, olot.PlanVersion)
'            If oPlan Is Nothing Then GoTo FillCommentHandler
'
'            Set oAggrSteps = oPlan.Steps(olot.CurrentStep.Steps.Item(1).Handle, "+999")
'            For Each oAggrStep In oAggrSteps
'                .Sheets("WS").Select
'                .Rows("11:20").Select
'                .Selection.Copy
'                .Sheets(1).Select
'                .Range("A" & CStr(iPosComment)).Select
'                .Selection.Insert -4121
'
'                iPosProcess = iPosProcess + 10
'                iPosComment = iPosComment + 10
'                .Range("E" & CStr(iPosProcess)).Select
'                .ActiveCell.FormulaR1C1 = oAggrStep.Steps.Item(1).Description
'                .Range("M" & CStr(iPosProcess)).Select
'                .ActiveCell.FormulaR1C1 = oAggrStep.Steps.Item(1).Id
'            Next
'            Set oPlan = Nothing
'            Set oAggrSteps = Nothing
'            Set oAggrStep = Nothing
'        End If
        
FillCommentHandler:
    'add by Nelson Start on 2005/11/30 for ReqNo:M200505002
    'Get LotType
    sLotType = GetLotType(oLot.Id, moProRawSql, moAppLog)
    'add by Nelson end on 2005/11/30 for ReqNo:M200505002
    
    'Comment
    sSQL = " select * from ( " _
         & " select " & gsCAT_TLFA_STEP_NO & " Step, " & gsCAT_TLFA_COMMENTS & ", " _
                      & gsCAT_TLFA_USER_ID & ", " & gsCAT_TLFA_SET_TIME & " sTime,'' IncludeEngLot,'' IncludeReworkLot,'Lot' FutActType " _
         & " From " & gsCAT_TBL_LOT_FUTACT _
         & " where " & gsCAT_TLFA_LOT_ID & " = '" & oLot.Id & "' " _
         & " and " & gsCAT_TLFA_COMMENTS & " is not null " _
         & " and " & gsCAT_TLFA_DELETE_FLAG & " <> 'Y' " _
         & " Union " _
         & " select " & gsCAT_TIFA_STEP_NO & " Step, " & gsCAT_TIFA_COMMENTS & ", " _
                      & gsCAT_TIFA_USER_ID & ", " & gsCAT_TIFA_SET_TIME & " sTime, " & gsCAT_TIFA_INCLUDEENGLOT & "," & gsCAT_TIFA_INCLUDEREWORKLOT & ",'Ipn' FutActType" _
         & " From " & gsCAT_TBL_IPN_FUTACT _
         & " where " & gsCAT_TIFA_IPN & " = '" & oLot.CustomAttributes(gsLOT_CUSTOMATTR_IPN) & "' " _
         & " and " & gsCAT_TIFA_COMMENTS & " is not null " _
         & " and " & gsCAT_TIFA_DELETE_FLAG & " <> 'Y' "
    
    '-----------------------------------------------------------------------------------
    '2008/03/12 Vencent M200802020, add IncludeLotOwner
    sSQL = sSQL & " and (nvl(" & gsCAT_TIFA_INCLUDELOTOWNER & ", 'All')='All' or " & _
            gsCAT_TIFA_INCLUDELOTOWNER & " = '" & sLotOwner & "' or " & _
            gsCAT_TIFA_INCLUDELOTOWNER & " like '" & sLotOwner & ",%' or " & _
            gsCAT_TIFA_INCLUDELOTOWNER & " like '%," & sLotOwner & "' or " & _
            gsCAT_TIFA_INCLUDELOTOWNER & " like '%," & sLotOwner & ",%')"
    '-----------------------------------------------------------------------------------
    
    'add by Nelson start on 2005/11/24 for ReqNo:M200505002
    sSQL = sSQL & " Union " _
         & " select " & gsCAT_TPGFA_STEP_NO & " Step, " & gsCAT_TPGFA_COMMENTS & ", " _
                      & gsCAT_TPGFA_USER_ID & ", " & gsCAT_TPGFA_SET_TIME & " sTime, " & gsCAT_TPGFA_INCLUDEENGLOT & "," & gsCAT_TPGFA_INCLUDEREWORKLOT & ",'ProdGroup' FutActType" _
         & " From " & gsCAT_TBL_PRODGROUP_FUTACT _
         & " where " & gsCAT_TPGFA_PROD_GROUP & " = '" & oLot.CustomAttributes(gsLOT_CUSTOMATTR_PROD_GROUP) & "' " _
         & " and " & gsCAT_TPGFA_COMMENTS & " is not null " _
         & " and " & gsCAT_TPGFA_DELETE_FLAG & " <>'Y' "
    'add by Nelson end on 2005/11/24 for ReqNo:M200505002
    
    '-----------------------------------------------------------------------------------
    '2008/03/12 Vencent M200802020, add IncludeLotOwner
    sSQL = sSQL & " and (nvl(" & gsCAT_TPGFA_INCLUDELOTOWNER & ", 'All')='All' or " & _
            gsCAT_TPGFA_INCLUDELOTOWNER & " = '" & sLotOwner & "' or " & _
            gsCAT_TPGFA_INCLUDELOTOWNER & " like '" & sLotOwner & ",%' or " & _
            gsCAT_TPGFA_INCLUDELOTOWNER & " like '%," & sLotOwner & "' or " & _
            gsCAT_TPGFA_INCLUDELOTOWNER & " like '%," & sLotOwner & ",%')"
    '-----------------------------------------------------------------------------------

    sSQL = sSQL & " Union " _
         & " select " & gsCAT_THR_STEP_ID & " Step, " & gsCAT_THR_COMMENTS & ", " _
                      & gsCAT_THR_CREATE_USER_ID & ", " & gsCAT_THR_CREATE_TIME & " sTime,'' IncludeEngLot,'' IncludeReworkLot,'' FutActType" _
         & " From " & gsCAT_TBL_HOLD_REC _
         & " where " & gsCAT_THR_LOT_ID & " = '" & oLot.Id & "' " _
         & " and " & gsCAT_THR_COMMENTS & " is not null " _
         & " and " & gsCAT_THR_DELETE_FLAG & " <> 'Y' " _
         & " Union " _
         & " select " & gsCAT_TMT_STEPNO & " Step, " & gsCAT_TMT_COMMENTS & ", " _
                      & gsCAT_TMT_CREATE_USER_ID & ", " & gsCAT_TMT_CREATE_TIME & " sTime,'' IncludeEngLot,'' IncludeReworkLot,'' FutActType" _
         & " From " & gsCAT_TBL_MANUAL_TESTQTY _
         & " where " & gsCAT_TMT_LOT_ID & " = '" & oLot.Id & "' " _
         & " and " & gsCAT_TMT_COMMENTS & " is not null " _
         & " and " & gsCAT_TMT_DELETE_FLAG & " <> 'Y' " _
         & " ) order by sTime "
         
        Set colSQL2 = moProRawSql.QueryDatabase(sSQL)
        
        If colSQL2.Count > 0 Then
            iIdx2 = 1
            For iIdx = 1 To colSQL2.Count
                sIncludEngLot = colSQL2.Item(iIdx).Item(5)
                sIncludeReworkLot = colSQL2.Item(iIdx).Item(6)
                sFutActType = colSQL2.Item(iIdx).Item(7)
                If sFutActType = "Ipn" Or sFutActType = "ProdGroup" Then
                    Select Case sLotType
                        Case "Normal"
                            .range("C" & CStr(iPosComment + iIdx2 + 1)).Select
                            .activeCell.FormulaR1C1 = colSQL2.Item(iIdx).Item(1)
                            .range("Q" & CStr(iPosComment + iIdx2 + 1)).Select
                            .activeCell.FormulaR1C1 = colSQL2.Item(iIdx).Item(2)
                            .range("K" & CStr(iPosComment + iIdx2 + 1)).Select
                            .activeCell.FormulaR1C1 = colSQL2.Item(iIdx).Item(3)
                            .range("F" & CStr(iPosComment + iIdx2 + 1)).Select
                            .activeCell.FormulaR1C1 = Format(ConvertToDate(colSQL2.Item(iIdx).Item(4), moAppLog), "YYYY/MM/DD HH:MM:SS")
                            iIdx2 = iIdx2 + 1
                        Case "Eng"
                            If sIncludEngLot = "Y" Then
                                .range("C" & CStr(iPosComment + iIdx + 1)).Select
                                .activeCell.FormulaR1C1 = colSQL2.Item(iIdx).Item(1)
                                .range("Q" & CStr(iPosComment + iIdx + 1)).Select
                                .activeCell.FormulaR1C1 = colSQL2.Item(iIdx).Item(2)
                                .range("K" & CStr(iPosComment + iIdx + 1)).Select
                                .activeCell.FormulaR1C1 = colSQL2.Item(iIdx).Item(3)
                                .range("F" & CStr(iPosComment + iIdx + 1)).Select
                                .activeCell.FormulaR1C1 = Format(ConvertToDate(colSQL2.Item(iIdx).Item(4), moAppLog), "YYYY/MM/DD HH:MM:SS")
                                iIdx2 = iIdx2 + 1
                            End If
                        Case "Rework"
                            If sIncludeReworkLot = "Y" Then
                                .range("C" & CStr(iPosComment + iIdx + 1)).Select
                                .activeCell.FormulaR1C1 = colSQL2.Item(iIdx).Item(1)
                                .range("Q" & CStr(iPosComment + iIdx + 1)).Select
                                .activeCell.FormulaR1C1 = colSQL2.Item(iIdx).Item(2)
                                .range("K" & CStr(iPosComment + iIdx + 1)).Select
                                .activeCell.FormulaR1C1 = colSQL2.Item(iIdx).Item(3)
                                .range("F" & CStr(iPosComment + iIdx + 1)).Select
                                .activeCell.FormulaR1C1 = Format(ConvertToDate(colSQL2.Item(iIdx).Item(4), moAppLog), "YYYY/MM/DD HH:MM:SS")
                                iIdx2 = iIdx2 + 1
                            End If
                        Case Else
                    End Select
                Else
                    .range("C" & CStr(iPosComment + iIdx2 + 1)).Select
                    .activeCell.FormulaR1C1 = colSQL2.Item(iIdx).Item(1)
                    .range("Q" & CStr(iPosComment + iIdx2 + 1)).Select
                    .activeCell.FormulaR1C1 = colSQL2.Item(iIdx).Item(2)
                    .range("K" & CStr(iPosComment + iIdx2 + 1)).Select
                    .activeCell.FormulaR1C1 = colSQL2.Item(iIdx).Item(3)
                    .range("F" & CStr(iPosComment + iIdx2 + 1)).Select
                    .activeCell.FormulaR1C1 = Format(ConvertToDate(colSQL2.Item(iIdx).Item(4), moAppLog), "YYYY/MM/DD HH:MM:SS")
                    iIdx2 = iIdx2 + 1
                End If
            Next iIdx
        End If
        iPosTitle = iPosComment + 3 + colSQL2.Count


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

Private Sub txtStep_Change()
    On Error Resume Next
    Me.lblStepName = ""
    
    Dim sSQL As String, col1 As Collection
    If Len(txtStep.Text) = 5 Then
        sSQL = " select DESCRIPTION from FWSTEPVERSION " _
             & " where STEPNAME = '" & txtStep & "'" _
             & " and REVSTATE = 'Active' "
        Set col1 = moProRawSql.QueryDatabase(sSQL)
        If col1.Count > 0 Then
            Me.lblStepName = col1.Item(1).Item(1)
        End If
    End If
End Sub

'================================================================================
' Function: GetLotAttr()
'--------------------------------------------------------------------------------
' Description:  <Type your function description here...>
'--------------------------------------------------------------------------------
' Author:       Vencent, CIT 2003/03/20
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
'   string          (R) True/False
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
Private Function GetLotAttr(sLotID As String, sAttr As String, sTxnTime As String) As String
On Error GoTo ExitHandler:
Dim sProcID As String
Dim typErrInfo As tErrInfo
'----
' Init
'----
    sProcID = "GetLotAttr"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog) '"Entering Function...", moAppLog, glLOG_PROC, msMODULE_ID, sProcID)

'----
' Condition Checking
'----
    ' <Put your condition checking codes here>...

'----
' Action
'----
    ' <Put your Action codes here>...
    Dim sSQL As String, colSQLResult As Collection
    
    sSQL = " select NEWVALUE " & _
           " from FWWIPTRANSACTION " & _
           " where WIPID = '" & sLotID & "'" & _
           " and ATTRIBUTE = '" & sAttr & "'" & _
           " and TXNTIMESTAMP = " & _
           " (select max(TXNTIMESTAMP) from FWWIPTRANSACTION " & _
           " where WIPID = '" & sLotID & "' " & _
           " and ATTRIBUTE = '" & sAttr & "' " & _
           " and TXNTIMESTAMP <= '" & sTxnTime & "' )"
'        Debug.Print "Get " & sAttr & " Sql = " & sSQL
        Set colSQLResult = moProRawSql.QueryDatabase(sSQL)
        If colSQLResult Is Nothing Then
            Call RaiseError(glERR_RAWSQLERROR, _
                        FormatErrorText(gsETX_RAWSQLERROR, "Select", "NEWVALUE from FWWIPTRANSACTION"))
        ElseIf colSQLResult.Count > 0 Then
            GetLotAttr = colSQLResult.Item(1).Item(1)
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
' Function: GetEqpAttr()
'--------------------------------------------------------------------------------
' Description:  <Type your function description here...>
'--------------------------------------------------------------------------------
' Author:       Vencent, CIT 2003/03/21
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
Private Function GetEqpAttr(sEqpid As String, sAttr As String, sTxnTime As String) As String
On Error GoTo ExitHandler:
Dim sProcID As String
Dim typErrInfo As tErrInfo
'----
' Init
'----
    sProcID = "GetEqpAttr"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog) '"Entering Function...", moAppLog, glLOG_PROC, msMODULE_ID, sProcID)

'----
' Condition Checking
'----
    ' <Put your condition checking codes here>...

'----
' Action
'----
    ' <Put your Action codes here>...
    Dim sSQL As String, colSQLResult As Collection
    '20220517, CIT-Vincent, update sql use VIEW_FWEQPHISTORY for 2.8
    '" from FWEQPHISTORY "
    sSQL = " select ATTRIBUTEMAGNITUDE" & _
           " from VIEW_FWEQPHISTORY " & _
           " where EQUIPMENTNAME = '" & sEqpid & "'" & _
           " and EVENTNAME = 'FwEquipmentModifyAttributeTxn'" & _
           " and ATTRIBUTENAME = '" & sAttr & "'" & _
           " and TIME = ( select max(TIME) " & _
           " from FWEQPHISTORY " & _
           " where EQUIPMENTNAME = '" & sEqpid & "'" & _
           " and EVENTNAME = 'FwEquipmentModifyAttributeTxn'" & _
           " and ATTRIBUTENAME = '" & sAttr & "'" & _
           " and TIME < '" & sTxnTime & "')"

    Set colSQLResult = moProRawSql.QueryDatabase(sSQL)
    If colSQLResult.Count > 0 Then GetEqpAttr = colSQLResult.Item(1).Item(1)

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


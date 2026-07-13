VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#8.0#0"; "fpSPR80.OCX"
Begin VB.Form frmPrintISGForm 
   Caption         =   "列印 ISG 出貨單"
   ClientHeight    =   2895
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6150
   LinkTopic       =   "Form1"
   ScaleHeight     =   2895
   ScaleWidth      =   6150
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtShipNo 
      Height          =   315
      Left            =   1620
      TabIndex        =   0
      Top             =   600
      Width           =   2355
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok(確認)"
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
      Left            =   1200
      TabIndex        =   4
      Top             =   2400
      Width           =   1800
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel(放棄)"
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
      Left            =   3045
      TabIndex        =   3
      Top             =   2400
      Width           =   1800
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "列印出貨單"
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
      Left            =   4440
      TabIndex        =   2
      Top             =   540
      Width           =   1455
   End
   Begin FPSpreadADO.fpSpread spdPrint 
      Height          =   855
      Left            =   1140
      TabIndex        =   5
      Top             =   1260
      Visible         =   0   'False
      Width           =   3795
      _Version        =   196613
      _ExtentX        =   6694
      _ExtentY        =   1508
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
      MaxCols         =   50
      MaxRows         =   5
      RetainSelBlock  =   0   'False
      SpreadDesigner  =   "frmPrintISGForm.frx":0000
   End
   Begin VB.Label lblLabelShipNo 
      AutoSize        =   -1  'True
      Caption         =   "出貨單據號碼:"
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
      TabIndex        =   1
      Top             =   660
      Width           =   1230
   End
End
Attribute VB_Name = "frmPrintISGForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const msMODULE_ID As String = "frmPrintISGForm"
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

Private miResult    As Integer

Private Enum spdPrint_WS_Col
    'Modify by HouYu for 202000026
    'miSpdPrintCol_WS_MaxCols = 13
    miSpdPrintCol_WS_MaxCols = 9
    'miSpdPrintCol_WS_IsMerge = 1
    miSpdPrintCol_WS_LotNo = 1
    miSpdPrintCol_WS_IPN = 2
    miSpdPrintCol_WS_WQty = 3
    miSpdPrintCol_WS_CQty = 4
    miSpdPrintCol_WS_PoNo = 5
    miSpdPrintCol_WS_PoItem = 6
    miSpdPrintCol_WS_CustLotNo = 7
    miSpdPrintCol_WS_IsFail
    miSpdPrintCol_WS_MergeLotID
    miSpdPrintCol_WS_MergeCQty
    miSpdPrintCol_WS_OrPoNo
    miSpdPrintCol_WS_OrPoItem
    miSpdPrintCol_WS_WaferNo = 9
    'Added by HouYu for 202000026
End Enum

Private Enum spdPrint_FT_Col
    miSpdPrintCol_FT_MaxCols = 13
    miSpdPrintCol_FT_IsMerge = 1
    miSpdPrintCol_FT_LotNo
    miSpdPrintCol_FT_IPN
    miSpdPrintCol_FT_CQty
    miSpdPrintCol_FT_PoNo
    miSpdPrintCol_FT_PoItem
    miSpdPrintCol_FT_CustLotNo
    miSpdPrintCol_FT_PassQty
    miSpdPrintCol_FT_FailQty
    miSpdPrintCol_FT_MergeLotID
    miSpdPrintCol_FT_MergeCQty
    miSpdPrintCol_FT_OrPoNo
    miSpdPrintCol_FT_OrPoItem
End Enum

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
'================================================================================
' Sub: cmdOk_Click()
'--------------------------------------------------------------------------------
' Description:  ReqNo:M200707026
'--------------------------------------------------------------------------------
' Author:       Nelson, MXIC 2007/08/03
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
Private Sub cmdOk_Click()
On Error GoTo ExitHandler:
Dim sProcID     As String
Dim typErrInfo  As tErrInfo
Dim sFilePath   As String
Dim sShipNo     As String
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
   
    sShipNo = UCase(Me.txtShipNo.Text)
    If sShipNo = "" Then
        UtShowMsgBox ("請輸入出貨單號!" & vbNewLine & vbNewLine & "Please key in ShipNo!")
        GoTo ExitHandler
    End If
'----
' Action
'----
    Dim oLot        As FwLot
    Dim appXL       As Object
    Dim iPosTitle   As Integer
        
    Set appXL = CreateObject("Excel.Application")
    If appXL Is Nothing Then
        UtShowMsgBox "Excel not found"
        GoTo ExitHandler
    End If

    appXL.cursor = 2
    appXL.DisplayAlerts = False
    'appXL.ActiveWindow.WindowState = vbMaximized

    Call ISGForm(appXL)
   
    ChDir "C:\FASTech"
    appXL.ActiveWorkbook.SaveAs "C:\FASTech\" & sShipNo & "_" & Format(Now, "YYYYMMDDHHMMSS") & ".xls", _
        , "", "", False, False

    appXL.Visible = True
    Set appXL = Nothing
        
    miResult = vbOK
    

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

'================================================================================
' Sub: cmdPrint_Click()
'--------------------------------------------------------------------------------
' Description:  ReqNo:M200707026
'--------------------------------------------------------------------------------
' Author:       Nelson, MXIC 2007/08/03
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
Private Sub cmdPrint_Click()
On Error GoTo ExitHandler:
Dim sProcID As String
Dim typErrInfo As tErrInfo
Dim sShipNo     As String
Dim sLotId      As String
Dim sIPN        As String
Dim sWQty       As String
Dim sCQty       As String
Dim sPassQty    As String
Dim sFailQty    As String
Dim sPONO       As String
Dim sPOItem     As String
Dim sSQL        As String
Dim sShipType   As String
Dim lChildQty   As Long
Dim colRS       As Collection
Dim oItem       As Object

'----
' Init
'----
    sProcID = "cmdPrint_Click"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog)
    Screen.MousePointer = vbHourglass
'----
' Condition Checking
'----
    sShipNo = UCase(Me.txtShipNo.Text)
    
    If sShipNo = "" Then
        UtShowMsgBox ("請輸入出貨單號!" & vbNewLine & vbNewLine & "Please key in ShipNo!")
        GoTo ExitHandler
    End If
    
    If VBA.Left(sShipNo, 4) = "11TU" Then
        sShipType = "FT"
    Else
        sShipType = "WS"
    End If
    
    ' <Put your condition checking codes here>...
    sSQL = "SELECT " & gsCAT_TSF_LOTID & "," & _
                       gsCAT_TSF_IPN & "," & _
                       gsCAT_TSF_SHIPQTY & "," & _
                       gsCAT_TSF_WAFERQTY & "," & _
                       gsCAT_TSF_CHIPQTY & "," & _
                       gsCAT_TSF_PASSQTY & "," & _
                       gsCAT_TSF_FAILQTY & " " & _
           "FROM " & gsCAT_TBL_SHIP_FG & " " & _
           "WHERE " & gsCAT_TSF_TICKETNO & "='" & sShipNo & "' AND " & _
                      gsCAT_TSF_DELETE_FLAG & "='N'" & _
           "ORDER BY " & gsCAT_TSF_CREATE_TIME & " DESC "
           
    Set colRS = moProRawSql.QueryDatabase(sSQL)
    
    If colRS.Count > 0 Then
        sLotId = colRS.Item(1).Item(1)
        sIPN = colRS.Item(1).Item(2)
        sWQty = colRS.Item(1).Item(4)
        sCQty = colRS.Item(1).Item(5)
        sPassQty = colRS.Item(1).Item(6)
        sFailQty = colRS.Item(1).Item(7)
    Else
        UtShowMsgBox ("出貨單號(" & sShipNo & ")不存在!" & vbNewLine & vbNewLine & "ShipNo(" & sShipNo & ") is not exist in system!")
        GoTo ExitHandler
    End If
'----
' Action
'----
    ' <Put your Action codes here>...
    Me.spdPrint.Reset
    Me.spdPrint.MaxRows = 0
    If sShipType = "FT" Then
        Me.spdPrint.MaxCols = miSpdPrintCol_FT_MaxCols
    Else
        Me.spdPrint.MaxCols = miSpdPrintCol_WS_MaxCols
    End If
    With Me.spdPrint
        .MaxRows = .MaxRows + 1
        If sShipType = "FT" Then
            .SetText miSpdPrintCol_FT_LotNo, .MaxRows, sLotId
            .SetText miSpdPrintCol_FT_CQty, .MaxRows, sCQty
            .SetText miSpdPrintCol_FT_IPN, .MaxRows, sIPN
        Else
            .SetText miSpdPrintCol_WS_LotNo, .MaxRows, sLotId
            .SetText miSpdPrintCol_WS_CQty, .MaxRows, sCQty
            .SetText miSpdPrintCol_WS_WQty, .MaxRows, sWQty
            .SetText miSpdPrintCol_WS_IPN, .MaxRows, sIPN
        End If
        
        'Get PoNo, PoItem
        sSQL = "select " & gsCAT_TLI_LOT_ID & "," & _
                           gsCAT_TLI_CUR_PO_NO & "," & _
                           gsCAT_TLI_CUR_PO_ITEM & "," & _
                           gsCAT_TLI_CUSTLOTID & "," & _
                           gsCAT_TLI_CUSTPOITEM & "," & _
                           gsCAT_TLI_CUSTPONO & _
               " from " & gsCAT_TBL_LOT_INFO & _
               " where " & gsCAT_TLI_LOT_ID & " = '" & sLotId & "'"
        Set colRS = moProRawSql.QueryDatabase(sSQL)
        If colRS.Count > 0 Then
            sPONO = colRS.Item(1).Item(gsCAT_TLI_CUR_PO_NO)
            sPOItem = colRS.Item(1).Item(gsCAT_TLI_CUR_PO_ITEM)
            If sShipType = "FT" Then
                .SetText miSpdPrintCol_FT_PoNo, .MaxRows, colRS.Item(1).Item(gsCAT_TLI_CUR_PO_NO)
                .SetText miSpdPrintCol_FT_PoItem, .MaxRows, colRS.Item(1).Item(gsCAT_TLI_CUR_PO_ITEM)
                .SetText miSpdPrintCol_FT_CustLotNo, .MaxRows, colRS.Item(1).Item(gsCAT_TLI_CUSTLOTID)
                .SetText miSpdPrintCol_FT_OrPoNo, .MaxRows, colRS.Item(1).Item(gsCAT_TLI_CUSTPONO)
                .SetText miSpdPrintCol_FT_OrPoItem, .MaxRows, colRS.Item(1).Item(gsCAT_TLI_CUSTPOITEM)
            Else
                .SetText miSpdPrintCol_WS_PoNo, .MaxRows, colRS.Item(1).Item(gsCAT_TLI_CUR_PO_NO)
                .SetText miSpdPrintCol_WS_PoItem, .MaxRows, colRS.Item(1).Item(gsCAT_TLI_CUR_PO_ITEM)
                .SetText miSpdPrintCol_WS_CustLotNo, .MaxRows, colRS.Item(1).Item(gsCAT_TLI_CUSTLOTID)
                .SetText miSpdPrintCol_WS_OrPoNo, .MaxRows, colRS.Item(1).Item(gsCAT_TLI_CUSTPONO)
                .SetText miSpdPrintCol_WS_OrPoItem, .MaxRows, colRS.Item(1).Item(gsCAT_TLI_CUSTPOITEM)
            End If
        End If
        
        If sShipType = "FT" Then
            lChildQty = 0
            .SetText miSpdPrintCol_FT_MergeLotID, .MaxRows, sLotId
            .SetText miSpdPrintCol_FT_PassQty, .MaxRows, sPassQty
            .SetText miSpdPrintCol_FT_FailQty, .MaxRows, sFailQty
            
            sSQL = "select b.valdata lotid, d.valdata chipqty," & vbNewLine & _
                   gsCAT_TLI_CUSTLOTID & "," & _
                   gsCAT_TLI_CUSTPOITEM & "," & _
                   gsCAT_TLI_CUSTPONO & vbNewLine & _
                "from fwmerge a," & vbNewLine & _
                "     fwmerge_pn2m b," & vbNewLine & _
                "     fwlot c," & vbNewLine & _
                "     fwlot_pn2m d," & vbNewLine & _
                    gsCAT_TBL_LOT_INFO & " e" & vbNewLine & _
                "where a.wipid = '" & sLotId & "'" & vbNewLine & _
                "and a.sysid = b.fromid" & vbNewLine & _
                "and b.linkname = 'childLotCollection'" & vbNewLine & _
                "and substr(a.mergestepid,1,5) = '70710'" & vbNewLine & _
                "and b.valdata = c.appid" & vbNewLine & _
                "and c.sysid = d.fromid" & vbNewLine & _
                "and d.keydata = 'ChipQty'" & vbNewLine & _
                "and b.valdata = e.lotid(+)"
            Set colRS = moProRawSql.QueryDatabase(sSQL)
            
            'set merge info to spdPrint for print
            For Each oItem In colRS
                .SetText miSpdPrintCol_FT_IsMerge, .MaxRows, "Y"
                .MaxRows = .MaxRows + 1
                .SetText miSpdPrintCol_FT_LotNo, .MaxRows, sLotId
                .SetText miSpdPrintCol_FT_CQty, .MaxRows, sCQty
                .SetText miSpdPrintCol_FT_IPN, .MaxRows, sIPN
                
                .SetText miSpdPrintCol_FT_PoNo, .MaxRows, sPONO
                .SetText miSpdPrintCol_FT_PoItem, .MaxRows, sPOItem
                .SetText miSpdPrintCol_FT_CustLotNo, .MaxRows, oItem.Item(gsCAT_TLI_CUSTLOTID)
                .SetText miSpdPrintCol_FT_OrPoNo, .MaxRows, oItem.Item(gsCAT_TLI_CUSTPONO)
                .SetText miSpdPrintCol_FT_OrPoItem, .MaxRows, oItem.Item(gsCAT_TLI_CUSTPOITEM)
                
                .SetText miSpdPrintCol_FT_PassQty, .MaxRows, sPassQty
                .SetText miSpdPrintCol_FT_FailQty, .MaxRows, sFailQty
                
                .SetText miSpdPrintCol_FT_IsMerge, .MaxRows, "Y"
                .SetText miSpdPrintCol_FT_MergeLotID, .MaxRows, oItem.Item(1)
                .SetText miSpdPrintCol_FT_MergeCQty, .MaxRows, oItem.Item(2)
                lChildQty = lChildQty + Val(oItem.Item(2))
            Next oItem
            .SetText miSpdPrintCol_FT_MergeCQty, .MaxRows - colRS.Count, CStr(sCQty - lChildQty)
        'Added by HouYu for 202000026 on 20200218
        Else
            .Row = .MaxRows
            .Col = miSpdPrintCol_WS_WaferNo
            .TypeMaxEditLen = 80
            .SetText miSpdPrintCol_WS_WaferNo, .MaxRows, GetWaferNo(sLotId)
        End If
        'Added by HouYu for 202000026 on 20200218
    End With
    
    'Print 出貨單
    If Me.spdPrint.MaxRows > 0 Then Call PrintSheet(sShipType)
    
    Me.txtShipNo.Text = ""
    Set colRS = Nothing
    
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
        Init = False
        On Error GoTo ExitHandler:
        Call HandleError(True, typErrInfo, , moAppLog)
    End If
End Function

'================================================================================
' Sub: ISGForm()
'--------------------------------------------------------------------------------
' Description:  ReqNo:M200707026
'--------------------------------------------------------------------------------
' Author:       Nelson, MXIC 2007/08/03
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
Private Sub ISGForm(ByRef appXL As Object)
On Error GoTo ExitHandler:
Dim sProcID     As String
Dim typErrInfo  As tErrInfo
Dim sShipNo     As String
Dim sLotId      As String
Dim sIPN        As String
Dim sWQty       As String
Dim sCQty       As String
Dim sPassQty    As String
Dim sFailQty    As String
Dim sPONO       As String
Dim sPOItem     As String
Dim sSQL        As String
Dim sShipType   As String
Dim lChildQty   As Long
Dim colRS       As Collection
Dim oItem       As Object
Dim iPosTitle   As Integer
Dim oBook       As Object
Dim oSheet      As Object


'----
' Init
'----
    sProcID = "ISGForm"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog)

'----
' Condition Checking
'----
    ' <Put your condition checking codes here>...

'----
' Action
'----
    ' <Put your Action codes here>...
    iPosTitle = 1
    Set oBook = appXL.Workbooks.Add
    Set oSheet = oBook.Worksheets(1)
    
    appXL.Visible = True
    
    With oSheet
        sShipNo = UCase(Me.txtShipNo.Text)
        
        'setup columns width
        'Modify by HouYu for 202000026 on 20200218
        If VBA.Left(sShipNo, 4) = "11TU" Then 'FT
            .Columns("A").ColumnWidth = 12
            .Columns("B").ColumnWidth = 12
            .Columns("C").ColumnWidth = 25
            .Columns("D").ColumnWidth = 10
            .Columns("E").ColumnWidth = 10
            .Columns("F").ColumnWidth = 10
            .Columns("G").ColumnWidth = 10
            .Columns("H").ColumnWidth = 10
            .Columns("I").ColumnWidth = 10
            .Columns("J").ColumnWidth = 12
            .Columns("K").ColumnWidth = 12
            .Columns("L").ColumnWidth = 12
            .Columns("M").ColumnWidth = 12
        Else 'WS
            .Columns("A").ColumnWidth = 12 'LotNo
            .Columns("B").ColumnWidth = 25 'IPN
            .Columns("C").ColumnWidth = 10 'WQty
            .Columns("D").ColumnWidth = 10 'CQty
            .Columns("E").ColumnWidth = 10 'PoNo
            .Columns("F").ColumnWidth = 10 'PoItem
            .Columns("G").ColumnWidth = 10 'CustLotNo
            .Columns("H").ColumnWidth = 10 '好壞品
            .Columns("I").ColumnWidth = 80 'WaferNo
        End If
        
        'Modify by HouYu for 202000026 on 20200218
        'fill header
        .range("A" & CStr(iPosTitle)).Value = "MXIC/BE代工出貨單"
        .range("A" & CStr(iPosTitle)).Font.Bold = True
        .range("A" & CStr(iPosTitle)).Font.Size = 16
        
        .range("A" & CStr(iPosTitle + 2)).Value = "單據號碼："
        .range("A" & CStr(iPosTitle + 2)).Font.Size = 14
        .range("A" & CStr(iPosTitle + 2)).ShrinkToFit = True
        
        .range("B" & CStr(iPosTitle + 2)).Value = sShipNo
        .range("B" & CStr(iPosTitle + 2)).Font.Size = 14
        .range("B" & CStr(iPosTitle + 2)).ShrinkToFit = True
        
        .range("D" & CStr(iPosTitle + 2)).Value = "出貨簽名："
        .range("D" & CStr(iPosTitle + 2)).Font.Size = 14
        .range("D" & CStr(iPosTitle + 2)).ShrinkToFit = True
        
        .range("G" & CStr(iPosTitle + 2)).Value = "收貨簽名："
        .range("G" & CStr(iPosTitle + 2)).Font.Size = 14
        .range("G" & CStr(iPosTitle + 2)).ShrinkToFit = True
       
        If VBA.Left(sShipNo, 4) = "11TU" Then
            .range("A" & CStr(iPosTitle + 4)).Value = "併批"
           
            .range("B" & CStr(iPosTitle + 4)).Value = "LotNo"
           
            .range("C" & CStr(iPosTitle + 4)).Value = "IPN"
                        
            .range("D" & CStr(iPosTitle + 4)).Value = "CQty"
            
            .range("E" & CStr(iPosTitle + 4)).Value = "PoNo"
                        
            .range("F" & CStr(iPosTitle + 4)).Value = "PoItem"
            
            .range("G" & CStr(iPosTitle + 4)).Value = "CustLotNo"
            
            .range("H" & CStr(iPosTitle + 4)).Value = "PassQty"
            
            .range("I" & CStr(iPosTitle + 4)).Value = "FailQty"
            
            .range("J" & CStr(iPosTitle + 4)).Value = "併批前批號"
                        
            .range("K" & CStr(iPosTitle + 4)).Value = "併批前CQty"
            
            .range("L" & CStr(iPosTitle + 4)).Value = "前批號PoNo"
            
            .range("M" & CStr(iPosTitle + 4)).Value = "前批號PoItem"
        Else
            'Modify by HouYu for 202000026 on 20200218
            '.range("A" & CStr(iPosTitle + 4)).Value = "併批"
                        
            '.range("B" & CStr(iPosTitle + 4)).Value = "LotNo"
            
            '.range("C" & CStr(iPosTitle + 4)).Value = "IPN"
                        
            '.range("D" & CStr(iPosTitle + 4)).Value = "WQty"
            
            '.range("E" & CStr(iPosTitle + 4)).Value = "CQty"
            
            '.range("F" & CStr(iPosTitle + 4)).Value = "PoNo"
            
            '.range("G" & CStr(iPosTitle + 4)).Value = "PoItem"
            
            '.range("H" & CStr(iPosTitle + 4)).Value = "CustLotNo"
            
            '.range("I" & CStr(iPosTitle + 4)).Value = "好壞品"
            
            '.range("J" & CStr(iPosTitle + 4)).Value = "併批前批號"
            
            '.range("K" & CStr(iPosTitle + 4)).Value = "併批前CQty"
            
            '.range("L" & CStr(iPosTitle + 4)).Value = "前批號PoNo"
            
            '.range("M" & CStr(iPosTitle + 4)).Value = "前批號PoItem"
            
            .range("A" & CStr(iPosTitle + 4)).Value = "LotNo"
            
            .range("B" & CStr(iPosTitle + 4)).Value = "IPN"
                        
            .range("C" & CStr(iPosTitle + 4)).Value = "WQty"
            
            .range("D" & CStr(iPosTitle + 4)).Value = "CQty"
            
            .range("E" & CStr(iPosTitle + 4)).Value = "PoNo"
            
            .range("F" & CStr(iPosTitle + 4)).Value = "PoItem"
            
            .range("G" & CStr(iPosTitle + 4)).Value = "CustLotNo"
            
            .range("H" & CStr(iPosTitle + 4)).Value = "好壞品"
            
            .range("I" & CStr(iPosTitle + 4)).Value = "WaferNo"
            'Modify by HouYu for 202000026 on 20200218
        End If
        
        sSQL = "SELECT " & gsCAT_TSF_LOTID & "," & _
                           gsCAT_TSF_IPN & "," & _
                           gsCAT_TSF_SHIPQTY & "," & _
                           gsCAT_TSF_WAFERQTY & "," & _
                           gsCAT_TSF_CHIPQTY & "," & _
                           gsCAT_TSF_PASSQTY & "," & _
                           gsCAT_TSF_FAILQTY & " " & _
               "FROM " & gsCAT_TBL_SHIP_FG & " " & _
               "WHERE " & gsCAT_TSF_TICKETNO & "='" & sShipNo & "' AND " & _
                          gsCAT_TSF_DELETE_FLAG & "='N' " & _
               "ORDER BY " & gsCAT_TSF_CREATE_TIME & " DESC "
               
        Set colRS = moProRawSql.QueryDatabase(sSQL)
        
        If colRS.Count > 0 Then
            sLotId = colRS.Item(1).Item(1)
            sIPN = colRS.Item(1).Item(2)
            sWQty = colRS.Item(1).Item(4)
            sCQty = colRS.Item(1).Item(5)
            sPassQty = colRS.Item(1).Item(6)
            sFailQty = colRS.Item(1).Item(7)
        Else
            UtShowMsgBox ("出貨單號(" & sShipNo & ")不存在!" & vbNewLine & vbNewLine & "ShipNo(" & sShipNo & ") is not exist in system!")
            GoTo ExitHandler
        End If
        If VBA.Left(sShipNo, 4) = "11TU" Then
            .range("B" & CStr(iPosTitle + 5)).Value = sLotId
            .range("C" & CStr(iPosTitle + 5)).Value = sIPN
            .range("D" & CStr(iPosTitle + 5)).Value = sCQty
        Else
            'Modify by HouYu for 202000026 on 20200218
            '.range("B" & CStr(iPosTitle + 5)).Value = sLotId
            '.range("C" & CStr(iPosTitle + 5)).Value = sIPN
            '.range("D" & CStr(iPosTitle + 5)).Value = sWQty
            '.range("E" & CStr(iPosTitle + 5)).Value = sCQty
            
            .range("A" & CStr(iPosTitle + 5)).Value = sLotId
            .range("B" & CStr(iPosTitle + 5)).Value = sIPN
            .range("C" & CStr(iPosTitle + 5)).Value = sWQty
            .range("D" & CStr(iPosTitle + 5)).Value = sCQty
            'Modify by HouYu for 202000026 on 20200218
        End If
        
        'Get PoNo, PoItem
        sSQL = "select " & gsCAT_TLI_LOT_ID & "," & _
                           gsCAT_TLI_CUR_PO_NO & "," & _
                           gsCAT_TLI_CUR_PO_ITEM & "," & _
                           gsCAT_TLI_CUSTLOTID & "," & _
                           gsCAT_TLI_CUSTPOITEM & "," & _
                           gsCAT_TLI_CUSTPONO & _
               " from " & gsCAT_TBL_LOT_INFO & _
               " where " & gsCAT_TLI_LOT_ID & " = '" & sLotId & "'"
        Set colRS = moProRawSql.QueryDatabase(sSQL)
        If colRS.Count > 0 Then
            sPONO = colRS.Item(1).Item(gsCAT_TLI_CUR_PO_NO)
            sPOItem = colRS.Item(1).Item(gsCAT_TLI_CUR_PO_ITEM)
            If VBA.Left(sShipNo, 4) = "11TU" Then
                .range("E" & CStr(iPosTitle + 5)).Value = colRS.Item(1).Item(gsCAT_TLI_CUR_PO_NO)
                
                .range("F" & CStr(iPosTitle + 5)).Value = colRS.Item(1).Item(gsCAT_TLI_CUR_PO_ITEM)
                
                .range("G" & CStr(iPosTitle + 5)).Value = colRS.Item(1).Item(gsCAT_TLI_CUSTLOTID)
                
                .range("L" & CStr(iPosTitle + 5)).Value = colRS.Item(1).Item(gsCAT_TLI_CUSTPONO)

                .range("M" & CStr(iPosTitle + 5)).Value = colRS.Item(1).Item(gsCAT_TLI_CUSTPOITEM)
            Else
                'Modify by HouYu for 202000026 on 20200218
                '.range("F" & CStr(iPosTitle + 5)).Value = colRS.Item(1).Item(gsCAT_TLI_CUR_PO_NO)

                '.range("G" & CStr(iPosTitle + 5)).Value = colRS.Item(1).Item(gsCAT_TLI_CUR_PO_ITEM)

                '.range("H" & CStr(iPosTitle + 5)).Value = colRS.Item(1).Item(gsCAT_TLI_CUSTLOTID)

                '.range("L" & CStr(iPosTitle + 5)).Value = colRS.Item(1).Item(gsCAT_TLI_CUSTPONO)
                
                '.range("M" & CStr(iPosTitle + 5)).Value = colRS.Item(1).Item(gsCAT_TLI_CUSTPOITEM)
                
                .range("E" & CStr(iPosTitle + 5)).Value = colRS.Item(1).Item(gsCAT_TLI_CUR_PO_NO)

                .range("F" & CStr(iPosTitle + 5)).Value = colRS.Item(1).Item(gsCAT_TLI_CUR_PO_ITEM)

                .range("G" & CStr(iPosTitle + 5)).Value = colRS.Item(1).Item(gsCAT_TLI_CUSTLOTID)
                'Modify by HouYu for 202000026 on 20200218
            End If
        End If
        
        If VBA.Left(sShipNo, 4) = "11TU" Then
            lChildQty = 0
            .range("J" & CStr(iPosTitle + 5)).Value = sLotId
            
            .range("H" & CStr(iPosTitle + 5)).Value = sPassQty

            .range("I" & CStr(iPosTitle + 5)).Value = sFailQty
            
            sSQL = "select b.valdata lotid, d.valdata chipqty," & vbNewLine & _
                   gsCAT_TLI_CUSTLOTID & "," & _
                   gsCAT_TLI_CUSTPOITEM & "," & _
                   gsCAT_TLI_CUSTPONO & vbNewLine & _
                "from fwmerge a," & vbNewLine & _
                "     fwmerge_pn2m b," & vbNewLine & _
                "     fwlot c," & vbNewLine & _
                "     fwlot_pn2m d," & vbNewLine & _
                    gsCAT_TBL_LOT_INFO & " e" & vbNewLine & _
                "where a.wipid = '" & sLotId & "'" & vbNewLine & _
                "and a.sysid = b.fromid" & vbNewLine & _
                "and b.linkname = 'childLotCollection'" & vbNewLine & _
                "and substr(a.mergestepid,1,5) = '70710'" & vbNewLine & _
                "and b.valdata = c.appid" & vbNewLine & _
                "and c.sysid = d.fromid" & vbNewLine & _
                "and d.keydata = 'ChipQty'" & vbNewLine & _
                "and b.valdata = e.lotid(+)"
            Set colRS = moProRawSql.QueryDatabase(sSQL)
            
            'set merge info to spdPrint for print
            For Each oItem In colRS
                .range("A" & CStr(iPosTitle + 5)).Value = "Y"
                
                iPosTitle = iPosTitle + 1
                .range("B" & CStr(iPosTitle + 5)).Value = sLotId
                
                .range("C" & CStr(iPosTitle + 5)).Value = sIPN

                .range("D" & CStr(iPosTitle + 5)).Value = sCQty

                .range("E" & CStr(iPosTitle + 5)).Value = sPONO

                .range("F" & CStr(iPosTitle + 5)).Value = sPOItem

                .range("G" & CStr(iPosTitle + 5)).Value = oItem.Item(gsCAT_TLI_CUSTLOTID)
                
                .range("H" & CStr(iPosTitle + 5)).Value = sPassQty
    
                .range("I" & CStr(iPosTitle + 5)).Value = sFailQty
                
                .range("L" & CStr(iPosTitle + 5)).Value = oItem.Item(gsCAT_TLI_CUSTPONO)

                .range("M" & CStr(iPosTitle + 5)).Value = oItem.Item(gsCAT_TLI_CUSTPOITEM)

                .range("A" & CStr(iPosTitle + 5)).Value = "Y"
                
                .range("J" & CStr(iPosTitle + 5)).Value = oItem.Item(1)

                .range("K" & CStr(iPosTitle + 5)).Value = oItem.Item(2)

                lChildQty = lChildQty + Val(oItem.Item(2))
            Next oItem
            .range("K" & CStr(iPosTitle + 5 - colRS.Count)).Value = CStr(sCQty - lChildQty)
        'Added by HouYu for 202000026 on 20200218
        Else
            .range("I" & CStr(iPosTitle + 5)).Value = GetWaferNo(sLotId)
        'Added by HouYu for 202000026 on 20200218
        End If
        'Modify by HouYu for 202000026 on 20200218
        '.range("A5:M" & iPosTitle + 5).borders.LineStyle = 1
        .range("A5:I" & iPosTitle + 5).borders.LineStyle = 1
        'Modify by HouYu for 202000026 on 20200218
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
' Sub: PrintSheet()
' Author:       Vencent, CIT 2007/03/02
'================================================================================
Private Sub PrintSheet(sShipType As String)
On Error GoTo ExitHandler:
Dim sProcID As String
Dim typErrInfo As tErrInfo
'----
' Init
'----
    sProcID = "PrintSheet"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog)
'----
' Action
'----
    Dim buf As String
    
    With Me.spdPrint
        'modify by Nelson start on 2007/08/03 for ReqNo:M200707026
        If sShipType = "FT" Then
            .Col = spdPrint_FT_Col.miSpdPrintCol_FT_MaxCols
        Else
            .Col = spdPrint_WS_Col.miSpdPrintCol_WS_MaxCols
        End If
        'modify by Nelson end on 2007/08/03 for ReqNo:M200707026
        
        .Row = .MaxRows
        If .Text = "" Then .Text = " "
        'Spread Header
            'modify by Nelson start on 2007/08/03 for ReqNo:M200707026
            If sShipType = "FT" Then
                .Col = 1
                .col2 = spdPrint_FT_Col.miSpdPrintCol_FT_MaxCols
                .Row = 0
                .Row2 = 0
                .Clip = "併批" & Chr(9) & _
                        "LotNo" & Chr(9) & _
                        "IPN" & Chr(9) & _
                        "CQty" & Chr(9) & _
                        "PoNo" & Chr(9) & _
                        "PoItem" & Chr(9) & _
                        "CustLotNo" & Chr(9) & _
                        "PassQty" & Chr(9) & _
                        "FailQty" & Chr(9) & _
                        "併批前批號" & Chr(9) & _
                        "併批前CQty" & Chr(9) & _
                        "前批號PoNo" & Chr(9) & _
                        "前批號PoItem"
              
            Else
                .Col = 1
                .col2 = spdPrint_WS_Col.miSpdPrintCol_WS_MaxCols
                .Row = 0
                .Row2 = 0
                'Modify by HouYu for 202000026 on 20200218
                '.Clip = "併批" & Chr(9) & _
                '        "LotNo" & Chr(9) & _
                '        "IPN" & Chr(9) & _
                '        "WQty" & Chr(9) & _
                '        "CQty" & Chr(9) & _
                '        "PoNo" & Chr(9) & _
                '        "PoItem" & Chr(9) & _
                '        "CustLotNo" & Chr(9) & _
                '        "好壞品" & Chr(9) & _
                '        "併批前批號" & Chr(9) & _
                '        "併批前CQty" & Chr(9) & _
                '        "前批號PoNo" & Chr(9) & _
                '        "前批號PoItem"
                .Clip = "LotNo" & Chr(9) & _
                        "IPN" & Chr(9) & _
                        "WQty" & Chr(9) & _
                        "CQty" & Chr(9) & _
                        "PoNo" & Chr(9) & _
                        "PoItem" & Chr(9) & _
                        "CustLotNo" & Chr(9) & _
                        "好壞品" & Chr(9) & _
                        "WaferNo"
                'Modify by HouYu for 202000026 on 20200218
            End If
            'modify by Nelson end on 2007/08/03 for ReqNo:M200707026
        .Col = 1
        Do While .Col <= .MaxCols
            .ColWidth(.Col) = .MaxTextColWidth(.Col) + 5
            .Col = .Col + 1
        Loop
        
        .Row = -1
        .Col = -1
        .TypeHAlign = TypeHAlignCenter
        
        
        '.SetCellBorder 1, 0, .MaxCols, 0, 8, -1, CellBorderStyleSolid
        .SetCellBorder 0, 0, .MaxCols, .MaxRows, 15, -1, CellBorderStyleSolid
    
        .PrintBorder = False
        
        'Report Header
        buf = "/fn""Arial""/fz""16""/fb1/fi0/fu0/fk0/fs1MXIC/ BE代工出貨單/n/n/n"
        buf = buf & "/fn""Arial""/fz""12""/fb0/fi0/fu0/fk0/fs1單據號碼:   " & UCase(Me.txtShipNo.Text) & "      出貨簽名:                           收貨簽名:        /n/n"
        .PrintGrid = False
        .PrintHeader = buf
        .PrintGrid = False
        .PrintShadows = False
        .PrintSmartPrint = True
        .PrintRowHeaders = False
        .PrintOrientation = PrintOrientationLandscape
        .PrintUseDataMax = True
        .PrintSheet (0)
        .PrintSheet (0)
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
' Sub: GetWaferNo()
' Author:HouYu 20200218
'================================================================================
Function GetWaferNo(sLotId As String)
On Error GoTo ExitHandler:
Dim sProcID As String
Dim typErrInfo As tErrInfo

Dim iIndex    As Integer
Dim sSQL        As String
Dim colRS       As Collection
'----
' Init
'----
    sProcID = "GetWaferNo"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog)
'----
' Action
'----
        sSQL = "select a.appid," & _
                  "c.componentid waferno," & _
                  "d.valdata gooddie " & _
                "from fwlot a," & _
                     "fwlot_n2m b," & _
                     "fwcomponent c," & _
                     "fwcomponent_pn2m d " & _
                "where a.SysId = b.fromid " & _
                  "and b.toid = c.sysid " & _
                  "and c.sysid = d.fromid " & _
                  "and a.appid='" & sLotId & "' " & _
                "order by c.componentid"
        Set colRS = moProRawSql.QueryDatabase(sSQL)
        If colRS.Count > 0 Then
          For iIndex = 1 To colRS.Count
            If iIndex = 1 Then
              GetWaferNo = colRS.Item(iIndex).Item("waferno")
            Else
              GetWaferNo = GetWaferNo & ";" & colRS.Item(iIndex).Item("waferno")
            End If
          Next iIndex
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

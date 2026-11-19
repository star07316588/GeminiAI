Option Explicit

Private Const msMODULE_ID As String = "frmMonitorPCQuery_Detail"

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

Private msIp        As String
Private msSystemName As String

Private miResult    As Integer

Private msType      As String

'**********
' Constant Declaration
'**********
'-----
' Spread spdEngCost
'-----
Private Const msspdEngCost_FILENAME                           As String = "FileName"
Private Const msspdEngCost_FILEUPDATETIME                     As String = "FileUpdateTime"

Private Const mispdEngCost_FILENAME                           As Integer = 1
Private Const mispdEngCost_FILEUPDATETIME                     As Integer = 2

Public Property Let LotID(sLotId As String)
    On Error Resume Next
    msLotId = sLotId
End Property
Public Property Let EqpId(sEqpid As String)
    On Error Resume Next
    msEqpId = sEqpid
End Property
Public Property Let QueryType(sType As String)
    On Error Resume Next
    msType = sType
End Property

Public Property Let IP(sIp As String)
    On Error Resume Next
    msIp = sIp
End Property
Public Property Let SystemName(sSystemName As String)
    On Error Resume Next
    msSystemName = sSystemName
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



Private Sub cmdClose_Click()
    On Error Resume Next
    miResult = vbCancel
    Me.Hide
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
Dim oLots      As FwLots
Dim oLot       As FwLot
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
    Call InitspdEngCost
    Call ExecQuery
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
' Function: InitspdEngCost()
'--------------------------------------------------------------------------------
' Description:  <Type your function description here...>
'--------------------------------------------------------------------------------
' Author:       Ivan, CIT 2002-06-05
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
Private Function InitspdEngCost() As Boolean
On Error GoTo ExitHandler:
Dim sProcID As String
Dim typErrInfo As tErrInfo
Dim lCol       As Long
'----
' Init
'----
    sProcID = "InitspdEngCost"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog) '"Entering Function...", moAppLog, glLOG_PROC, msMODULE_ID, sProcID)

'----
' Condition Checking
'----
    ' <Put your condition checking codes here>...

'----
' Action
'----
    With Me.spdEngCost
        .MaxCols = mispdEngCost_FILEUPDATETIME
        .MaxRows = 0
        
        Call .SetText(mispdEngCost_FILENAME, 0, msspdEngCost_FILENAME)
        Call .SetText(mispdEngCost_FILEUPDATETIME, 0, msspdEngCost_FILEUPDATETIME)
        
        For lCol = 1 To .MaxCols
            Call LockSpdColumn(Me.spdEngCost, lCol)
            .ColUserSortIndicator(lCol) = ColUserSortIndicatorAscending
            .UserColAction = UserColActionSort
        Next lCol
        
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
End Function

'================================================================================
' function: ExecQuery()
'--------------------------------------------------------------------------------
' Description:  <Type your function description here...>
'--------------------------------------------------------------------------------
' Author:       Sam, CIT 20090605
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
Public Sub ExecQuery()
On Error GoTo ExitHandler:
Dim sProcID As String
Dim typErrInfo As tErrInfo
Dim colRaws As Collection
Dim oRaws As Object
Dim sTable As String
Dim sColumn As String
Dim sWhere As String
Dim vTesterId As Variant
Dim lRow      As Long
Dim dteLastWaferTime As Date
Dim sSQL As String
Dim oEqp As FwEquipment
'----
' Init
'----
    sProcID = "ExecQuery"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog)
   
'----
' Condition Checking
'----
    ' <Put your condition checking codes here>...

'----
' Action
'----
        'Add by Sam start on 20200505 for ReqNO:202000138
        If msIp = "ALL" Then
            sSQL = "select a." & gsCAT_TMS_FILENAME & ",max(a." & gsCAT_TMS_FILEUPDATETIME & ") as " & gsCAT_TMS_FILEUPDATETIME & "  " & _
                        " from " & gsCAT_TBL_MONITORPC_SYSTEMVERSION & " a," & gsCAT_TBL_MONITORPC_SYSTEMVERSION_IP & " b  " & _
                        " where a." & gsCAT_TMS_IP & " = b." & gsCAT_TMSI_IP & " " & _
                        " and a." & gsCAT_TMS_MONITORSYSTEM & " = b." & gsCAT_TMSI_MONITORSYSTEM & " " & _
                        " and a." & gsCAT_TMS_MONITORSYSTEM & " = '" & msSystemName & "' " & _
                        " and a." & gsCAT_TMS_DELETEFLAG & " = 'N' " & _
                        " and b." & gsCAT_TMSI_DELETEFLAG & " = 'N' " & _
                        " group by a." & gsCAT_TMS_FILENAME & ""
        Else
        'Add by Sam end on 20200505 for ReqNO:202000138
            sTable = gsCAT_TBL_MONITORPC_SYSTEMVERSION
            
            sColumn = gsCAT_TMS_FILENAME & "," & _
                    gsCAT_TMS_FILEUPDATETIME & " "

                     
            sWhere = gsCAT_TMS_IP & "='" & msIp & "' and " & gsCAT_TMS_MONITORSYSTEM & " = '" & msSystemName & "' "
                     
            
            sSQL = "select " & sColumn & " from " & sTable & " where " & sWhere
        End If
        
        Set colRaws = moProRawSql.QueryDatabase(sSQL)
        
        If colRaws Is Nothing Then
            Call RaiseError(glERR_INVALIDOBJECT, _
                        FormatErrorText(gsETX_INVALIDOBJECT, "Collection"))
        End If
    
        For Each oRaws In colRaws
            
            With Me.spdEngCost
                .MaxRows = .MaxRows + 1
                
                Call .SetText(mispdEngCost_FILENAME, .MaxRows, oRaws.Item(gsCAT_TMS_FILENAME))
                Call .SetText(mispdEngCost_FILEUPDATETIME, .MaxRows, oRaws.Item(gsCAT_TMS_FILEUPDATETIME))
                                
            End With
        Next oRaws

    Call ResizeSpdColumn(Me.spdEngCost, moAppLog)
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








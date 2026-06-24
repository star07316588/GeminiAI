'================================================================================
' Sub: modPrint()
' File Name: modPrint.bas
' Module Prefix: [Fwu]
' Description: <Description here...>
'--------------------------------------------------------------------------------
' Author: Ivan, CIT 2002/02/07
'--------------------------------------------------------------------------------
' Additional Remarks:
' 1. <...>
'--------------------------------------------------------------------------------
' Revision History:
'................................................................................
' Date Modified: 2002/02/07
' Modified By: Ivan Lee
' Details:
' 1) Modules created.
'--------------------------------------------------------------------------------
' Date Modified:
' Modified By:
' Details:
' 1) <Put your details here...>
'================================================================================
Option Explicit

Private Const msMODULE_ID                                As String = "modPrint"

'----
' Label Spec
'----
Public Const gsLABEL_WS_TO_SFG                           As String = "WS_TO_SFG"
Public Const gsLABEL_WS_SUMMARY                          As String = "WS_SUMMARY"
Public Const gsLABEL_WS_SMALL_LABEL                      As String = "WS_SMALL_LABEL"
Public Const gsLABEL_TO_SUBFT_NORMAL                     As String = "TO_SUBFT_NORMAL"
Public Const gsLABEL_TO_SUBFT_ENG_SAMPLE                 As String = "TO_SUBFT_ENG_SAMPLE"
Public Const gsLABEL_TO_SUBMK_NORMAL                     As String = "TO_SUBMK_NORMAL"
Public Const gsLABEL_TO_SUBMK_ENG_SAMPLE                 As String = "TO_SUBMK_ENG_SAMPLE"
Public Const gsLABEL_TO_SUBTR_NORMAL                     As String = "TO_SUBTR_NORMAL"
Public Const gsLABEL_TO_SUBTR_ENG_SAMPLE                 As String = "TO_SUBTR_ENG_SAMPLE"
Public Const gsLABEL_FT_SMALL_LABEL                      As String = "FT_SMALL_LABEL"
Public Const gsLABEL_WSMCD_TO_SFG                        As String = "WSMCD_TO_SFG"
Public Const gsLABEL_FT_ETEST_MERGE                      As String = "FT_ETEST_MERGE_LABEL"
Public Const gsLABEL_WS_DGRADE_SUMMARY                   As String = "WS_DGRADE_SUMMARY"
Public Const gsLABEL_WS_MULTILOT_TO_SFG                  As String = "WSMULTILOT_TO_SFG"
Public Const gsLABEL_RENESAS_SHIPPING                    As String = "RENESAS_SHIPPING"
Public Const gsLABEL_FT_BIN_CARD_LABEL                   As String = "FT_BIN_CARD" 'Added by Jack on 2011/04/15.
Public Const gsLABEL_WS_ENG_LOC_LABEL                    As String = "WS_ENG_SHELF" 'Added by Sam on 2011/08/23.
Public Const gsLABEL_FT_BOX_COUNTING                     As String = "FT_BOX_COUNTING" 'Added by Tony on 20121001 for Req.JC201200283
'Added by Jack on 2016/06/03 for 2016 MES Phase-5 Project (FT eRunCard Paperless Phase-2 Project) <Start>
Public Const gsLABEL_FT_LOT_INFO                         As String = "FT_LOT_INFO"
Public Const gsLABEL_FT_Label_PACK_INFO                  As String = "FT_Label_PACK_INFO"
'Added by Jack on 2016/06/03 for 2016 MES Phase-5 Project (FT eRunCard Paperless Phase-2 Project) <End>
Public Const gsLABEL_CP_SMALL_LABEL                      As String = "CP_SMALL_LABEL" 'Add by Sam on 20180814 for BE工業 3.5 BE MES Phase 24 -- CP Auto Dispatch
Public Const gsLABEL_CP_VIRTUAL_LOT_LABEL                As String = "CP_VIRTUAL_LOT_LABEL" 'Add by Sam on 20201203 For CP 測前併批
Public Const gsLABEL_CP_VIRTUAL_MERGE                    As String = "CP_VIRTUAL_MERGE" 'Add by Sam on 20201203 For CP 測前併批
Public Const gsLABEL_FT_TR_LABEL                         As String = "FT_TR_LABEL" 'Added by HouYu for 202000413 on 20220413

'---------------
' Label Spec End
'---------------

'================================================================================
' Sub: Prt_FT_To_SubTR_Normal()
'--------------------------------------------------------------------------------
' Description:  <This subroutine will print out a label for FT-> SFG >
'--------------------------------------------------------------------------------
' Author:       <Alex Chan>, <CIT> 2002/02/07
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
Public Sub Prt_FT_To_SubTR_Normal(ByVal sLotNo As String, _
                                ByVal sProdNo As String, _
                                ByVal sQty As String, _
                                ByVal sPacker As String, _
                                ByVal bIsPartial As Boolean, _
                                ByVal sPrinterServer As String, _
                                ByVal moCwMbx As Object, _
                                ByVal oLogCtrl As Object)


On Error GoTo ExitHandler:

Dim sMsg            As String
Dim sProcID         As String
Dim typErrInfo      As tErrInfo
Dim lErrCode        As Long
Dim sErrMsg         As String


'----
' Init
'----
    sProcID = "Prt_FT_To_SubTR_Normal"
    Call LogProcIn(msMODULE_ID, sProcID, oLogCtrl)

'----
' Condition Checking
'----

'----
' Action
'----
    '******
    'Print Label Info
    '******
        sMsg = sMsg & "^XA^LH00,00^LL550"
        sMsg = sMsg & "^FO10,05^GB780,475,3^FS"
        sMsg = sMsg & "^FO435,327^GB355,70,3^FS"
        sMsg = sMsg & "^FO435,397^GB0,80,3^FS"
        sMsg = sMsg & "^FO580,30^GB200,75,3^FS"
        sMsg = sMsg & "^FO640,55^A0N,35,50^CI0^FDT&R^FS"
        If bIsPartial Then
            sMsg = sMsg & "^FO580,102^GB200,75,3^FS"
            sMsg = sMsg & "^FO590,130^A0N,35,50^CI0^FDPARTIAL^FS"
        End If
'        sMsg = sMsg & "^FO580,250^GB200,75,3^FS"
'        sMsg = sMsg & "^FO640,275^A0N,35,50^CI0^FDENG^FS"
        sMsg = sMsg & "^FO40,58^A0N,33,28^FD(1T)LOT NO: " & sLotNo & "^FS"
        sMsg = sMsg & "^FO40,158^A0N,33,28^FD(1P)PRODUCT NO: " & sProdNo & "^FS"
        sMsg = sMsg & "^FO40,258^A0N,33,28^FD(Q)QUANTITY: " & sQty & "^FS"
        sMsg = sMsg & "^FO40,90^BY2,2,50^B3,,,N^FD1T" & sLotNo & "^FS"
        sMsg = sMsg & "^FO40,190^BY2,2,50^B3,,,N^FD1P" & sProdNo & "^FS"
        sMsg = sMsg & "^FO40,290^BY2,2,50^B3,,,N^FDQ" & sQty & "^FS"
        sMsg = sMsg & "^FO445,350^A0N,36,30^FDPACKER: " & sPacker & "^FS"
        sMsg = sMsg & "^FO445,420^A0N,36,30^FDQA:^FS"
        sMsg = sMsg & "^PQ1"
        sMsg = sMsg & "^XZ1"
    '******
    ' End of Label Info
    '******

    '******
    'Send message to TDS
    '******
        
        'Add by Tony Start on 20110117 for Req.JC201100003
        If InStr(1, sPrinterServer, "@") <> 0 Then
            sPrinterServer = VBA.Left(sPrinterServer, (InStr(1, sPrinterServer, "@")) - 1)
        End If
        'Add by Tony End on 20110117 for Req.JC201100003
        
        If Not moCwMbx.Send("MBX_" & sPrinterServer, sMsg) Then
            Call moCwMbx.GetError(lErrCode, sErrMsg)
            Call RaiseError(lErrCode, sErrMsg)
        End If
    '******

'----
' Done
'----
ExitHandler:
    ' NOTE 1:
    ' MUST CALL GetErrInfo() here first before another action
    Call GetErrInfo(msMODULE_ID, sProcID, typErrInfo, Erl)
    Call LogProcOut(msMODULE_ID, sProcID, typErrInfo, oLogCtrl)
    On Error Resume Next

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
        Call HandleError(False, typErrInfo, , oLogCtrl, True)
    End If
End Sub

'================================================================================
' Sub: Prt_FT_To_SubFT_Normal()
'--------------------------------------------------------------------------------
' Description:  <This subroutine will print out a label for FT-> SFG >
'--------------------------------------------------------------------------------
' Author:       <Alex Chan>, <CIT> 2002/02/07
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
Public Sub Prt_FT_To_SubFT_Normal(ByVal sLotNo As String, _
                                ByVal sProdNo As String, _
                                ByVal sQty As String, _
                                ByVal sPacker As String, _
                                ByVal bIsPartial As Boolean, _
                                ByVal sPrinterServer As String, _
                                ByVal moCwMbx As Object, _
                                ByVal oLogCtrl As Object)
On Error GoTo ExitHandler:

Dim sMsg            As String
Dim sProcID         As String
Dim typErrInfo      As tErrInfo
Dim lErrCode        As Long
Dim sErrMsg         As String

'----
' Init
'----
    sProcID = "Prt_FT_To_SubFT_Normal"
    Call LogProcIn(msMODULE_ID, sProcID, oLogCtrl)

'----
' Condition Checking
'----

'----
' Action
'----
    '******
    'Print Label Info
    '******
        sMsg = sMsg & "^XA^LH00,00^LL550"
        sMsg = sMsg & "^FO10,05^GB780,475,3^FS"
        sMsg = sMsg & "^FO435,327^GB355,70,3^FS"
        sMsg = sMsg & "^FO435,397^GB0,80,3^FS"
        sMsg = sMsg & "^FO580,30^GB200,75,3^FS"
        sMsg = sMsg & "^FO635,55^A0N,35,50^CI0^FDTEST^FS"
        If bIsPartial Then
            sMsg = sMsg & "^FO580,102^GB200,75,3^FS"
            sMsg = sMsg & "^FO590,130^A0N,35,50^CI0^FDPARTIAL^FS"
        End If
'        sMsg = sMsg & "^FO580,250^GB200,75,3^FS"
'        sMsg = sMsg & "^FO640,275^A0N,35,50^CI0^FDENG^FS"
        sMsg = sMsg & "^FO40,58^A0N,33,28^FD(1T)LOT NO: " & sLotNo & "^FS"
        sMsg = sMsg & "^FO40,158^A0N,33,28^FD(1P)PRODUCT NO: " & sProdNo & "^FS"
        sMsg = sMsg & "^FO40,258^A0N,33,28^FD(Q)QUANTITY: " & sQty & "^FS"
        sMsg = sMsg & "^FO40,90^BY2,2,50^B3,,,N^FD1T" & sLotNo & "^FS"
        sMsg = sMsg & "^FO40,190^BY2,2,50^B3,,,N^FD1P" & sProdNo & "^FS"
        sMsg = sMsg & "^FO40,290^BY2,2,50^B3,,,N^FDQ" & sQty & "^FS"
        sMsg = sMsg & "^FO445,350^A0N,36,30^FDPACKER: " & sPacker & "^FS"
        sMsg = sMsg & "^FO445,420^A0N,36,30^FDQA:^FS"
        sMsg = sMsg & "^PQ1"
        sMsg = sMsg & "^XZ1"
    '******
    ' End of Label Info
    '******

    '******
    'Send message to TDS
    '******
        
        'Add by Tony Start on 20110117 for Req.JC201100003
        If InStr(1, sPrinterServer, "@") <> 0 Then
            sPrinterServer = VBA.Left(sPrinterServer, (InStr(1, sPrinterServer, "@")) - 1)
        End If
        'Add by Tony End on 20110117 for Req.JC201100003
        
        If Not moCwMbx.Send("MBX_" & sPrinterServer, sMsg) Then
            Call moCwMbx.GetError(lErrCode, sErrMsg)
            Call RaiseError(lErrCode, sErrMsg)
        End If
    '******

'----
' Done
'----
ExitHandler:
    ' NOTE 1:
    ' MUST CALL GetErrInfo() here first before another action
    Call GetErrInfo(msMODULE_ID, sProcID, typErrInfo, Erl)
    Call LogProcOut(msMODULE_ID, sProcID, typErrInfo, oLogCtrl)
    On Error Resume Next

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
        Call HandleError(False, typErrInfo, , oLogCtrl, True)
    End If
End Sub
'================================================================================
' Sub: PrintBarcodeLabel()
'--------------------------------------------------------------------------------
' Description:  <Type your Sub description here...>
'--------------------------------------------------------------------------------
' Author:       Ivan, CIT 2002/11/22
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
'Modify by Sam on 20110823 ,增加外部傳入WaferID
'Modified by Jack on 2014/06/13 for JC21400147,增加外部傳入 sFabLotID
'Modify 增加外部傳入 sOriLotID by Ernest on 2017/11/03 for MES Phase 16  Remark 專案
'================================================================================
Public Sub PrintBarcodeLabel(ByVal sLotID As String, ByVal sWQty As String, _
        ByVal sCQty As String, ByVal sIPN As String, _
        ByVal sLotOwner As String, ByVal sRouteId As String, _
        ByVal sPrinterServer As String, ByVal sLabelType As String, _
        ByVal sUserName As String, ByVal oLogCtrl As Object, _
        ByVal oFwWIP As Object, ByVal oFwWF As Object, _
        ByVal oCwMbx As Object, ByVal oProRawSql As Object, _
        Optional oWSMCDInfo As Collection, Optional ByVal sWaferID As String, _
        Optional ByVal sFabLotID As String, Optional ByVal sOriLotID As String)
    On Error GoTo ExitHandler:
    Dim sProcID As String
    Dim typErrInfo As tErrInfo
    Dim iIdx As Integer, iPrintTimes As Integer
    Dim sPrintQty As String
    Dim sBoxQty    As String
    Dim sPartNo    As String
    Dim sSpeed     As String
    Dim ofrmBoxQty As frmBoxQty
    Dim bIsPartial  As Boolean
    Dim sCargradeFlag   As String 'Add by Tony on 20110215 for Req.JC201100038
    'Add by Weilun on 20220526 for ReqNo.BE#202000233 <start>
    Dim sFuSa           As String
    Dim sIPN_Data       As String
    'Add by Weilun on 20220526 for ReqNo.BE#202000233 <end>
    Dim sLocation       As String 'Add by Weilun on 20220816 for ReqNo.BE#202200157

    'Added by Jack on 2012/08/15 for JC201200228 <Start>.
    Dim oLot       As FwLot
    Dim oComp      As FwComponent
    Dim colCompIds As FwStrings
    'Added by Jack on 2012/08/15 for JC201200228 <End>.
    
    Dim sProdLine  As String 'Add by Jack on 2015/06/25 for Req.JC201500212
    
    'Added by Jack on 2016/07/15 for 2016 MES Phase-5 Project (FT eRunCard Paperless Phase-2 Project) <Start>
    Dim sVendorCode       As String
    Dim sLotipn           As String
    'Dim sCqty             As String
    'Dim sRouteId          As String
    'Dim sLotOwner         As String
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
    
    Dim sBoxingSpecNo     As String
    Dim sPinCount         As String
    Dim sPackageCode      As String
    Dim sLabelSpecNo      As String
    Dim sLPI_IPN          As String
    
    Dim sSQL As String
    Dim colRS As Collection
    'Added by Jack on 2016/07/15 for 2016 MES Phase-5 Project (FT eRunCard Paperless Phase-2 Project) <End>

    '----
    ' Init
    '----
    sProcID = "PrintBarcodeLabel"
    Call LogProcIn(msMODULE_ID, sProcID, oLogCtrl)
    Set colCompIds = New FwStrings 'Added by Jack on 2012/08/17 for bug fix.

    '----
    ' Condition Checking
    '----
    ' <Put your condition checking codes here>...

    '----
    ' Action
    '----
    'gsLABEL_WS_DGRADE_SUMMARY added by Jack on 2015/01/20 for JC201500013(不需跳frmBoxQty / 2015/02/03與Cherry確認過)
    'gsLABEL_FT_LOT_INFO & gsLABEL_FT_Label_PACK_INFO 'Added by Jack on 2016/07/15 for 2016 MES Phase-5 Project (FT eRunCard Paperless Phase-2 Project)
    'gsLABEL_CP_SMALL_LABEL 'Add by Sam on 20180814 for BE工業 3.5 BE MES Phase 24 -- CP Auto Dispatch
    'gsLABEL_CP_VIRTUAL_LOT_LABEL,gsLABEL_CP_VIRTUAL_MERGE 'Add by SAm on 20201207 for Project CP 測前併批
    If (sLabelType = modPrint.gsLABEL_FT_SMALL_LABEL) Or _
        (sLabelType = modPrint.gsLABEL_CP_SMALL_LABEL) Or _
        (sLabelType = modPrint.gsLABEL_WS_SMALL_LABEL) Or _
        (sLabelType = modPrint.gsLABEL_CP_VIRTUAL_LOT_LABEL) Or _
        (sLabelType = modPrint.gsLABEL_CP_VIRTUAL_MERGE) Or _
        (sLabelType = modPrint.gsLABEL_WS_SUMMARY) Or _
        (sLabelType = modPrint.gsLABEL_WS_TO_SFG) Or _
        (sLabelType = gsLABEL_WS_DGRADE_SUMMARY) Or _
        (sLabelType = gsLABEL_FT_LOT_INFO) Or _
        (sLabelType = gsLABEL_FT_Label_PACK_INFO) Or _
        (sLabelType = modPrint.gsLABEL_WS_MULTILOT_TO_SFG) Or _
        (sLabelType = modPrint.gsLABEL_WSMCD_TO_SFG) Then
        '2010/01/11 start, by Robin, add for WS_MULTILOT_TO_SFG ReqNo:JC201000012
        'WSMCD_TO_SFG應已不再使用
        iPrintTimes = 1
        
        '2010/01/11 start, by Robin, ReqNo:JC201000012
        '***************
        '* Let User to key in Qty
        '***************
        'Modify by Sam on 20131210 for ReqNo:JC201300363 ,增加ws_summary也可以輸入張數
'        If (sLabelType = modPrint.gsLABEL_WS_MULTILOT_TO_SFG) Or _
'            (sLabelType = modPrint.gsLABEL_WSMCD_TO_SFG) Then
        If (sLabelType = modPrint.gsLABEL_WS_MULTILOT_TO_SFG) Or _
            (sLabelType = modPrint.gsLABEL_WS_SUMMARY) Or _
            (sLabelType = modPrint.gsLABEL_WSMCD_TO_SFG) Then
            iPrintTimes = GetLabelPrintQty("WS", sLabelType, oLogCtrl, oProRawSql)
            Set ofrmBoxQty = New frmBoxQty
            Load ofrmBoxQty
            With ofrmBoxQty
                Set .MainTraceLog = oLogCtrl
                .Caption = "Input Label Qty"
                .lblQty.Item(0).Alignment = 1
                .lblQty.Item(0).Caption = "Label Copies:"
                .txtLotId.Text = sLotID
                .txtBoxQty = iPrintTimes 'initialze
                .Show vbModal
                sBoxQty = .txtBoxQty
            End With
            Unload ofrmBoxQty
            iPrintTimes = CLng(sBoxQty)
        End If
        '2010/01/11 end, by Robin, ReqNo:JC201000012
    Else
        '***************
        '* Let User to key in Box Qty
        '***************
        Set ofrmBoxQty = New frmBoxQty
        Load ofrmBoxQty
        With ofrmBoxQty
            Set .MainTraceLog = oLogCtrl
            .txtLotId.Text = sLotID
            .Show vbModal
            sBoxQty = .txtBoxQty
        End With
        Unload ofrmBoxQty
        '***************
        iPrintTimes = CLng(sCQty) \ CLng(sBoxQty)
        If CLng(sCQty) Mod CLng(sBoxQty) > 0 Then
            iPrintTimes = iPrintTimes + 1
        End If
    End If
    
    bIsPartial = False
    
    For iIdx = 1 To iPrintTimes
        'Partial and Qty
        'gsLABEL_WS_DGRADE_SUMMARY added by Jack on 2015/01/20 for JC201500013(不需跳frmBoxQty / 2015/02/03與Cherry確認過)
        'gsLABEL_FT_LOT_INFO & gsLABEL_FT_Label_PACK_INFO 'Added by Jack on 2016/07/15 for 2016 MES Phase-5 Project (FT eRunCard Paperless Phase-2 Project)
        'gsLABEL_CP_SMALL_LABEL 'Add by Sam on 20180814 for BE工業 3.5 BE MES Phase 24 -- CP Auto Dispatch
        'gsLABEL_CP_VIRTUAL_LOT_LABEL,gsLABEL_CP_VIRTUAL_MERGE 'Add by SAm on 20201207 for Project CP 測前併批
        If (sLabelType = modPrint.gsLABEL_FT_SMALL_LABEL) Or _
            (sLabelType = modPrint.gsLABEL_CP_SMALL_LABEL) Or _
            (sLabelType = modPrint.gsLABEL_WS_SMALL_LABEL) Or _
            (sLabelType = modPrint.gsLABEL_WS_SUMMARY) Or _
            (sLabelType = modPrint.gsLABEL_CP_VIRTUAL_LOT_LABEL) Or _
            (sLabelType = modPrint.gsLABEL_CP_VIRTUAL_MERGE) Or _
            (sLabelType = modPrint.gsLABEL_WS_TO_SFG) Or _
            (sLabelType = modPrint.gsLABEL_WS_MULTILOT_TO_SFG) Or _
            (sLabelType = modPrint.gsLABEL_WSMCD_TO_SFG) Or _
            (sLabelType = gsLABEL_FT_LOT_INFO) Or _
            (sLabelType = gsLABEL_FT_Label_PACK_INFO) Or _
            (sLabelType = modPrint.gsLABEL_WS_DGRADE_SUMMARY) Then
            bIsPartial = False
        Else
            '***************
            '* Calculate how many Label to be Printed
            '***************
            If iIdx = iPrintTimes Then
                If CLng(sCQty) Mod CLng(sBoxQty) > 0 Then
                    bIsPartial = True
                    sPrintQty = CStr(CLng(sCQty) Mod CLng(sBoxQty))
                Else
                    bIsPartial = False
                    sPrintQty = sBoxQty
                End If
            Else
                bIsPartial = False
                sPrintQty = sBoxQty
            End If
            '***************
        End If
    
        If sLabelType = modPrint.gsLABEL_WS_TO_SFG Then
            sPartNo = modCustom.RetrieveEPNbyIPN(oLogCtrl, _
                oFwWIP, _
                oFwWF, _
                oCwMbx, _
                sIPN)
                Call modPrint.Prt_WS_WS_To_SFG(sLotID, _
                sIPN, _
                sWQty, _
                sCQty, _
                sPartNo, sPrinterServer, oCwMbx, oLogCtrl)
        'Add by Sam start on 20180814 for BE工業 3.5 BE MES Phase 24 -- CP Auto Dispatch
        ElseIf sLabelType = modPrint.gsLABEL_CP_SMALL_LABEL Then
            sSpeed = modCustom.GetSpeed(sIPN, _
                oProRawSql, oLogCtrl)
            Call modPrint.Prt_WS_CP_SMALL_LABEL(sLotID, _
                sIPN, _
                sWQty, _
                sCQty, _
                sLotOwner, _
                sSpeed, sPrinterServer, oCwMbx, oLogCtrl, oProRawSql, sWaferID, sFabLotID)
        'Add by Sam END on 20180814 for BE工業 3.5 BE MES Phase 24 -- CP Auto Dispatch
        ElseIf sLabelType = modPrint.gsLABEL_WS_SMALL_LABEL Then
            sSpeed = modCustom.GetSpeed(sIPN, _
                oProRawSql, oLogCtrl)
            'Modify by Sam start on 20080902 for ReqNo:M200809003 add oProRawSQL
            '            Call modPrint.Prt_WS_WS_SMALL_LABEL(sLotId, _
            '                                                sIPN, _
            '                                                sWqty, _
            '                                                sCqty, _
            '                                                sLotOwner, _
            '                                                sSpeed, sPrinterServer, oCwMbx, oLogCtrl)
            'sFabLotID added by Jack on 2014/06/13 for JC201400147.
            Call modPrint.Prt_WS_WS_SMALL_LABEL(sLotID, _
                sIPN, _
                sWQty, _
                sCQty, _
                sLotOwner, _
                sSpeed, sPrinterServer, oCwMbx, oLogCtrl, oProRawSql, sWaferID, sFabLotID)
            'Modify by Sam end on 20080902 for ReqNo:M200809003 add oProRawSQL
        'Add by Sam Start on 20201207 for Project CP 測前併批
        ElseIf sLabelType = modPrint.gsLABEL_CP_VIRTUAL_LOT_LABEL Or _
               sLabelType = modPrint.gsLABEL_CP_VIRTUAL_MERGE Then
            sSQL = "select tla." & gsCAT_TLATT_IPN & ", sum(tla." & gsCAT_TLATT_WAFERQTY & ") as wqty,sum(tla." & gsCAT_TLATT_CHIPQTY & ") as cqty " & _
                   " from (select distinct " & gsCAT_TVMG_VIRTUALLOTID & " ," & gsCAT_TVMG_LOTID & " " & _
                   "       from " & gsCAT_TBL_VIRTUAL_MERGE & " where " & gsCAT_TVMG_DELETEFLAG & "= 'N'  " & _
                   "        and " & gsCAT_TVMG_VIRTUALLOTID & " ='" & sLotID & "' ) a " & _
                   " , " & gsCAT_TBL_LOT_ATTRIBUTE & " tla " & _
                   " where a." & gsCAT_TVMG_LOTID & " = tla." & gsCAT_TLATT_LOTID & " " & _
                   " group by tla." & gsCAT_TLATT_IPN & " "
            Set colRS = oProRawSql.QueryDatabase(sSQL)
            If colRS.Count > 0 Then
                sLotipn = colRS.Item(1).Item(gsCAT_TLATT_IPN)
                sWQty = colRS.Item(1).Item("wqty")
                sCQty = colRS.Item(1).Item("cqty")
                If sLabelType = modPrint.gsLABEL_CP_VIRTUAL_LOT_LABEL Then
                    Call modPrint.Prt_WS_CP_VIRTUAL_LOT_LABEL(sLotID, _
                                                        sLotipn, _
                                                        sWQty, _
                                                        sCQty, _
                                                        sPrinterServer, _
                                                        oCwMbx, oLogCtrl, oProRawSql)
                ElseIf sLabelType = modPrint.gsLABEL_CP_VIRTUAL_MERGE Then
                    Call modPrint.Prt_WS_CP_VIRTUAL_LOT_LABEL(sLotID, _
                                                        sLotipn, _
                                                        sWQty, _
                                                        sCQty, _
                                                        sPrinterServer, _
                                                        oCwMbx, oLogCtrl, oProRawSql, True)
                End If
            End If
        'Add by Sam End on 20201207 for Project CP 測前併批
        ElseIf sLabelType = modPrint.gsLABEL_TO_SUBFT_NORMAL Then
            Call modPrint.Prt_FT_To_SubFT_Normal(sLotID, _
                sIPN, _
                sPrintQty, sUserName, _
                bIsPartial, sPrinterServer, oCwMbx, oLogCtrl)
        'Added by Jack on 2016/07/15 for 2016 MES Phase-5 Project (FT eRunCard Paperless Phase-2 Project) <Start>
        ElseIf sLabelType = modPrint.gsLABEL_FT_LOT_INFO Then
            'ps : 以IPN 前4碼去串 tbl_prod_body.prodline
            sSQL = "select a." & gsCAT_TLI_VENDORCODE & " ," & _
                             " b." & gsCAT_TLATT_IPN & "," & _
                             " b." & gsCAT_TLATT_CHIPQTY & "," & _
                             " b." & gsCAT_TLATT_ROUTE & "," & _
                             " b." & gsCAT_TLATT_LOTOWNER & "," & _
                             " b." & gsCAT_TLATT_DATECODE & "," & _
                             " c." & gsCAT_TIM_CARGRADEFLAG & "," & _
                             " e." & gsCAT_TPB_PROD_LINE & "," & _
                             " c." & gsCAT_TIM_CARRIER_TYPE & "," & _
                             " c." & gsCAT_TIM_EXTRAROMFLAG & "," & _
                             " c." & gsCAT_TIM_BRAND & "," & _
                             " c." & gsCAT_TIM_ICDRAWING & "," & _
                             " c." & gsCAT_TIM_MARKING_SPEC_NO & "," & _
                             " decode(c." & gsCAT_TIM_GPTYPE & ", null, null, 'XX', 'N', 'Y') as Green," & _
                             " c." & gsCAT_TIM_CHECKSUM & "," & _
                             " decode(d." & gsCAT_TPBS_CUSTOMER_NO & ", '0000010750', 'Y', 'N') as NFlag"
            sSQL = sSQL _
                        & " from " & gsCAT_TBL_LOT_INFO & " a, " _
                                   & gsCAT_TBL_LOT_ATTRIBUTE & " b, " _
                                   & gsCAT_TBL_IPN_MASTER & " c, " _
                                   & gsCAT_TBL_PRM_BE_SPEC & " d, " _
                                   & gsCAT_TBL_PROD_BODY & " e " _
                        & " where a." & gsCAT_TLI_LOT_ID & " = '" & sLotID & "' " _
                        & " and a." & gsCAT_TLI_LOT_ID & " = b." & gsCAT_TLATT_LOTID & " " _
                        & " and b." & gsCAT_TLATT_IPN & " = c." & gsCAT_TIM_IPN & "(+) " _
                        & " and c." & gsCAT_TIM_IPN & " = d." & gsCAT_TPBS_IPN & "(+) " _
                        & " and substr(b." & gsCAT_TLATT_IPN & ", 1, 4) = e." & gsCAT_TPB_PROD_BODY & "(+) " _
                        & " and d." & gsCAT_TPBS_DEFAULTS & "(+) = 'Y' " _
                        & " and ROWNUM = 1 "
            
            Set colRS = oProRawSql.QueryDatabase(sSQL)
            If Not colRS Is Nothing Then
                If colRS.Count > 0 Then
                    sVendorCode = colRS.Item(1).Item(1)
                    sLotipn = colRS.Item(1).Item(2)
                    sCQty = colRS.Item(1).Item(3)
                    sRouteId = colRS.Item(1).Item(4)
                    sLotOwner = colRS.Item(1).Item(5)
                    sDateCode = colRS.Item(1).Item(6)
                    sIpnCarGradeFlag = colRS.Item(1).Item(7)
                    sIpnProdLine = colRS.Item(1).Item(8)
                    sCarrierType = colRS.Item(1).Item(9)
                    sExtraRomFlag = colRS.Item(1).Item(10)
                    sBrand = colRS.Item(1).Item(11)
                    sIcDrawing = colRS.Item(1).Item(12)
                    sMarkingSpecNo = colRS.Item(1).Item(13)
                    sIpnGreen = colRS.Item(1).Item(14)
                    sIpnCsum = colRS.Item(1).Item(15)
                    sNFlag = colRS.Item(1).Item(16)
                    
                    Call modPrint.Prt_FT_FT_LOT_INFO(sLotID, Trim(sLotipn), Trim(sCQty), _
                                                    Trim(sRouteId), Trim(sLotOwner), Trim(sDateCode), _
                                                    Trim(sIpnCarGradeFlag), Trim(sIpnProdLine), Trim(sCarrierType), _
                                                    Trim(sExtraRomFlag), Trim(sBrand), Trim(sIcDrawing), _
                                                    Trim(sMarkingSpecNo), Trim(sIpnGreen), Trim(sIpnCsum), _
                                                    Trim(sVendorCode), Trim(sNFlag), sPrinterServer, _
                                                    oCwMbx, oLogCtrl)
                End If
            End If
        'Added by Jack on 2016/07/15 for 2016 MES Phase-5 Project (FT eRunCard Paperless Phase-2 Project) <End>
        
        'Added by Jack on 2016/07/15 for 2016 MES Phase-5 Project (FT eRunCard Paperless Phase-2 Project) <Start>
        ElseIf sLabelType = modPrint.gsLABEL_FT_Label_PACK_INFO Then
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
                       & " where d." & gsCAT_TLATT_LOTID & " = '" & sLotID & "' " _
                       & "   and d." & gsCAT_TLATT_IPN & " = a." & gsCAT_TIM_IPN & " " _
                       & "   and d." & gsCAT_TLATT_IPN & " = b." & gsCAT_TPBS_IPN & "(+) " _
                       & "   and b." & gsCAT_TPBS_CUSTOMER_NO & " = c." & gsCAT_TCUM_CUSTOMERNO & "(+) " _
                       & "   and b." & gsCAT_TPBS_DEFAULTS & "= 'Y' "
                       
            Set colRS = oProRawSql.QueryDatabase(sSQL)
            If colRS.Count > 0 Then
                sCarrierType = colRS.Item(1).Item(gsCAT_TIM_CARRIER_TYPE)
                sBoxingSpecNo = colRS.Item(1).Item(gsCAT_TIM_BOXING_SPEC_NO)
                sBrand = colRS.Item(1).Item(gsCAT_TIM_BRAND)
                sPinCount = colRS.Item(1).Item(gsCAT_TIM_PIN_COUNT)
                sPackageCode = colRS.Item(1).Item(gsCAT_TIM_PACKAGE_CODE)
                sLabelSpecNo = colRS.Item(1).Item(gsCAT_TCUM_LABELSPECNO)
                sLPI_IPN = colRS.Item(1).Item(gsCAT_TIM_IPN)
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
        
            Call modPrint.Prt_FT_Label_PACK_INFO(sLotID, sLabelSpecNo, _
                                                 sCarrierType, sBoxingSpecNo, _
                                                 sBrand, sPinCount, _
                                                 sPackageCode, True, sPrinterServer, _
                                                 oCwMbx, oLogCtrl, oProRawSql, sLPI_IPN)
                
        'Added by Jack on 2016/07/15 for 2016 MES Phase-5 Project (FT eRunCard Paperless Phase-2 Project) <End>
    
        ElseIf sLabelType = modPrint.gsLABEL_TO_SUBFT_ENG_SAMPLE Then
            Call modPrint.Prt_FT_To_SubFT_Eng_Sample(sLotID, _
                sIPN, _
                sPrintQty, sUserName, _
                bIsPartial, sPrinterServer, oCwMbx, oLogCtrl)
        ElseIf sLabelType = modPrint.gsLABEL_TO_SUBMK_NORMAL Then
            Call modPrint.Prt_FT_To_SubMK_Normal(sLotID, _
                sIPN, _
                sPrintQty, sUserName, _
                bIsPartial, sPrinterServer, oCwMbx, oLogCtrl)
        ElseIf sLabelType = modPrint.gsLABEL_TO_SUBMK_ENG_SAMPLE Then
            Call modPrint.Prt_FT_To_SubMK_Eng_Sample(sLotID, _
                sIPN, _
                sPrintQty, sUserName, _
                bIsPartial, sPrinterServer, oCwMbx, oLogCtrl)
        ElseIf sLabelType = modPrint.gsLABEL_TO_SUBTR_NORMAL Then
            Call modPrint.Prt_FT_To_SubTR_Normal(sLotID, _
                sIPN, _
                sPrintQty, sUserName, _
                bIsPartial, sPrinterServer, oCwMbx, oLogCtrl)
        ElseIf sLabelType = modPrint.gsLABEL_TO_SUBTR_ENG_SAMPLE Then
            Call modPrint.Prt_FT_To_SubTR_Eng_Sample(sLotID, _
                sIPN, _
                sPrintQty, sUserName, _
                bIsPartial, sPrinterServer, oCwMbx, oLogCtrl)
        ElseIf sLabelType = modPrint.gsLABEL_FT_SMALL_LABEL Then
            'Add by Sky at 2005/05/18 for M200501046
            Dim sCustomer As String
            Dim sCsum As String
            Dim sGreen As String
            Dim sGPType As String
            'Marked by Jack on 2016/07/15 for 2016 MES Phase-5 Project (FT eRunCard Paperless Phase-2 Project) <Start>
            'Dim sSql As String
            'Dim colRS As Collection
            'Marked by Jack on 2016/07/15 for 2016 MES Phase-5 Project (FT eRunCard Paperless Phase-2 Project) <End>
            
            sCustomer = ""
            sCsum = ""
            sGreen = ""
            sGPType = ""
                
            sSQL = "SELECT tbl_customer_master.customername, tbl_prm_backend_spec.ipn  " & _
                "  From tbl_lot_attribute, tbl_customer_master, tbl_prm_backend_spec " & _
                "  Where tbl_lot_attribute.IPN = tbl_prm_backend_spec.IPN " & _
                "   AND tbl_prm_backend_spec.customerno = tbl_customer_master.customerno " & _
                "   AND tbl_prm_backend_spec.defaults = 'Y' " & _
                "   AND lotid = '" & sLotID & "'" & _
                "    AND ROWNUM = 1 "
            Set colRS = oProRawSql.QueryDatabase(sSQL)
            If Not colRS Is Nothing Then
                If colRS.Count > 0 Then
                    sCustomer = colRS.Item(1).Item(1)
                    sGreen = colRS.Item(1).Item(2) 'used here to get IPN
                End If
            End If
            Set colRS = Nothing
                
            sIPN_Data = sIPN '預設sIPN, 如果tbl_ipn_master有資料則改成IPN的資料, Add by Weilun on 20220526 for ReqNo.BE#202000233
            
            'Add by Weilun on 20220816 for ReqNo.BE#202200157 <start>
            sLocation = ""
            sSQL = " select tsl." & gsCAT_TSL_LOCATION & " " & _
                     " from " & gsCAT_TBL_STEP_LOCATION & " tsl, " & _
                                gsCAT_TBL_LOT_ATTRIBUTE & " tlatt " & _
                    " where tsl." & gsCAT_TSL_STEPID & " = tlatt." & gsCAT_TLATT_STEPID & " " & _
                      " and tlatt." & gsCAT_TLATT_LOTID & " = '" & sLotID & "' "
            Set colRS = oProRawSql.QueryDatabase(sSQL)
            If Not colRS Is Nothing Then
                If colRS.Count > 0 Then
                    sLocation = colRS.Item(1).Item(gsCAT_TSL_LOCATION)
                End If
            End If
            Set colRS = Nothing
            'Add by Weilun on 20220816 for ReqNo.BE#202200157 <end>
                
            'Add CarGradeFlag by Tony on 20110215 for Req.JC201100038
            'Add FuSa,prodbody,pincount,package_code,bodysize by Weilun on 20220526 for ReqNo.BE#202000233
            sSQL = "SELECT tbl_ipn_master.checksum, tbl_ipn_master.gptype," & gsCAT_TIM_CARGRADEFLAG & _
                    ", " & gsCAT_TIM_FUSA & _
                    ", " & gsCAT_TIM_PRODBODY & _
                    ", " & gsCAT_TIM_PINCOUNT & _
                    ", " & gsCAT_TIM_PACKAGE_CODE & _
                    ", " & gsCAT_TIM_BODY_SIZE & _
                "  From tbl_ipn_master " & _
                "  where tbl_ipn_master.IPN = '" & sGreen & "'"
            Set colRS = oProRawSql.QueryDatabase(sSQL)
            If Not colRS Is Nothing Then
                If colRS.Count > 0 Then
                    sCsum = colRS.Item(1).Item(1)
                    sGPType = colRS.Item(1).Item(2)
                    sCargradeFlag = colRS.Item(1).Item(3)
                    'Add by Weilun on 20220526 for ReqNo.BE#202000233 <start>
                    sFuSa = colRS.Item(1).Item(gsCAT_TIM_FUSA)
                    If sLocation = "FT" Then 'FT的才用這種格式, Add by Weilun on 20220816 for ReqNo.BE#202200157
                        sIPN_Data = colRS.Item(1).Item(gsCAT_TIM_PRODBODY) & "-" & _
                                    colRS.Item(1).Item(gsCAT_TIM_PINCOUNT) & _
                                    colRS.Item(1).Item(gsCAT_TIM_PACKAGE_CODE) & " " & _
                                    colRS.Item(1).Item(gsCAT_TIM_BODY_SIZE)
                    End If
                    'Add by Weilun on 20220526 for ReqNo.BE#202000233 <end>
                    
                End If
            End If
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
                   "  WHERE A." & gsCAT_TPB_PROD_BODY & "= substr('" & Trim(sIPN) & "', 1,4) "
            Set colRS = oProRawSql.QueryDatabase(sSQL)
            If colRS.Count > 0 Then
                sProdLine = colRS.Item(1).Item(gsCAT_TPB_PROD_LINE)
            End If
            'Add by Jack on 2015/06/25 for Req.JC201500212 <End>.
        
            'Add CarGradeFlag by Tony on 20110215 for Req.JC2011000038
            'Modify 增加外部傳入 sOriLotID by Ernest on 2017/11/03 for MES Phase 16  Remark 專案
            'Add sFuSa & sIPN->sIPN_Data, Modify by Weilun on 20220526 for ReqNo.BE#202000233
            'Add sLocation, 並恢復部分輸入, Modify by Weilun on 20220816 for ReqNo.BE#202200157
            Call modPrint.Prt_FT_FT_SMALL_LABEL(sLotID, _
                sIPN_Data, _
                sWQty, _
                sCQty, _
                sLotOwner, _
                sRouteId, _
                sCustomer, _
                sCsum, _
                sGreen, _
                sPrinterServer, oCwMbx, oLogCtrl, sCargradeFlag, sProdLine, sOriLotID, sFuSa, sLocation)
'            Call modPrint.Prt_FT_FT_SMALL_LABEL(sLotId, _
'                sIPN_Data, _
'                sWQty, _
'                sCQty, _
'                "", _
'                sRouteId, _
'                sCustomer, _
'                "", _
'                sGreen, _
'                sPrinterServer, oCwMbx, oLogCtrl, sCargradeFlag, sProdLine, sOriLotID, sFusa)
'            Call modPrint.Prt_FT_FT_SMALL_LABEL(sLotId, _
'                sIPN, _
'                sWQty, _
'                sCQty, _
'                sLotOwner, _
'                sRouteId, _
'                sCustomer, _
'                sCsum, _
'                sGreen, _
'                sPrinterServer, oCwMbx, oLogCtrl, sCargradeFlag, sProdLine, sOriLotID)
        ElseIf sLabelType = modPrint.gsLABEL_WSMCD_TO_SFG Then
            Call modPrint.Prt_WSMCD_TO_SFG(sLotID, _
                sIPN, _
                sWQty, _
                sCQty, _
                sUserName, _
                GetSystemTime(oLogCtrl, oFwWIP, oFwWF, oCwMbx), _
                sPrinterServer, oCwMbx, oLogCtrl, _
                oWSMCDInfo)
        ElseIf sLabelType = modPrint.gsLABEL_WS_DGRADE_SUMMARY Then
            
            Dim sSysTime    As String
            Dim dteNow      As Date
            
            sSysTime = GetSystemTime(oLogCtrl, oFwWIP, oFwWF, oCwMbx)
            dteNow = ConvertToDate(sSysTime, oLogCtrl)
            
            'Modified by Tony on 201200010 for Req.JC201200010
            'mm/dd/yy hh:mm to yyyy/mm/dd
            Call modPrint.Prt_WS_WS_DGRADE_SUMMARY(sLotID, _
                sIPN, _
                sWQty, _
                sCQty, _
                sUserName, _
                Format(dteNow, "yyyy/mm/dd"), _
                sPrinterServer, _
                oCwMbx, oLogCtrl, oProRawSql)

        ElseIf sLabelType = modPrint.gsLABEL_WS_MULTILOT_TO_SFG Then
            
            'Add by Tony Start on 20120120 for Req.JC201200010
            sSysTime = GetSystemTime(oLogCtrl, oFwWIP, oFwWF, oCwMbx)
            dteNow = ConvertToDate(sSysTime, oLogCtrl)
            'Add by Tony End on 20120120 for Req.JC201200010
            
            '            Call Prt_WSMULTILOT_TO_SFG(sUserName, GetSystemTime(oLogCtrl, oFwWIP, oFwWF, oCwMbx), _
            '                    sWqty, sCqty, sPrinterServer, oCwMbx, oLogCtrl, oWSMCDInfo)
            Call Prt_WSMULTILOT_TO_SFG(sUserName, Format(dteNow, "yyyy/mm/dd"), _
                sWQty, sCQty, sPrinterServer, oCwMbx, oLogCtrl, oWSMCDInfo)
        'Added by Jack on 2012/08/15 for JC201200228 <Start>.
        ElseIf sLabelType = modPrint.gsLABEL_WS_SUMMARY Then
            Set oLot = oFwWIP.LotById(Trim(sLotID))
            If oLot Is Nothing Then
                Call UtShowMsgBox("The Lot is not exist!")
                GoTo ExitHandler
            End If
            For Each oComp In oLot.Components
                If oComp.Status <> modConstGlobal.gsCOMP_STATUS_SCRAPPED Then
                    colCompIds.Add oComp.Attributes(gsLOT_CUSTOMATTR_CHIP_QTY), oComp.Id
                End If
            Next oComp
            sSysTime = GetSystemTime(oLogCtrl, oFwWIP, oFwWF, oCwMbx)
            dteNow = ConvertToDate(sSysTime, oLogCtrl)
            '若有修改, 記得一併改 frmPrintLabel.cmdOK()
            Call Prt_WS_WS_SUMMARY(Trim(Trim(oLot.Id)), _
                Trim(oLot.CustomAttributes(modConstFwAttr.gsLOT_CUSTOMATTR_IPN).Value), _
                Trim(oLot.CustomAttributes(modConstFwAttr.gsLOT_CUSTOMATTR_WAFER_QTY).Value), _
                Trim(oLot.CustomAttributes(modConstFwAttr.gsLOT_CUSTOMATTR_CHIP_QTY).Value), _
                colCompIds, _
                Trim(sUserName), _
                Format(dteNow, "yyyy/mm/dd"), _
                sPrinterServer, _
                oCwMbx, oLogCtrl)
        'Added by Jack on 2012/08/15 for JC201200228 <End>.
        End If
    Next iIdx
    '----
    ' Done
    '----

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
        Case glERR_INVALIDOBJECT
            ' Retry code goes here...
        Case Else
            typErrInfo.sUserText = "Fail to execute application, please call IT support!!" & vbCrLf & _
                "程式執行失敗, 請洽IT人員處理"
        End Select
        '---- Start of Select Case Block ----
        On Error GoTo ExitHandler:
        Call HandleError(False, typErrInfo, , oLogCtrl, True)
    End If
End Sub
'================================================================================
' Sub: Prt_WS_WS_To_SFG()
'--------------------------------------------------------------------------------
' Description:  <This subroutine will print out a label for FT-> SFG >
'--------------------------------------------------------------------------------
' Author:       <Alex Chan>, <CIT> 2002/02/07
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
Public Sub Prt_WS_WS_To_SFG(ByVal sLotNo As String, _
                                ByVal sProdNo As String, _
                                ByVal sWQty As String, _
                                ByVal sQty As String, _
                                ByVal sPartNo As String, _
                                ByVal sPrinterServer As String, _
                                ByVal moCwMbx As Object, _
                                ByVal oLogCtrl As Object)

On Error GoTo ExitHandler:

Dim sMsg            As String
Dim sProcID         As String
Dim typErrInfo      As tErrInfo
Dim lErrCode        As Long
Dim sErrMsg         As String

'----
' Init
'----
    sProcID = "Prt_WS_WS_To_SFG"
    Call LogProcIn(msMODULE_ID, sProcID, oLogCtrl)

'----
' Condition Checking
'----

'----
' Action
'----
    '******
    'Print Label Info
    '******
        sMsg = sMsg & "^XA^LH10,00^LL500"
        sMsg = sMsg & "^FO00,05^GB780,475,3^FS"
        sMsg = sMsg & "^FO30,20^A0,27,25^FDLOT NO: " & sLotNo & "^FS"
        sMsg = sMsg & "^FO30,130^A0,27,25^FDPRODUCT NO: " & sProdNo & "^FS"
        sMsg = sMsg & "^FO30,235^A0,27,25^FDWAFER PCS:  " & sWQty & "PCS^FS"
        sMsg = sMsg & "^FO30,340^A0,27,25^FDQUANTITY:  " & sQty & "EA^FS"
        sMsg = sMsg & "^FO30,445^A0,27,25^FDPART NO: " & sPartNo & "^FS"
        sMsg = sMsg & "^FO30,55^BY2,2,55^B3,,,N^FD" & sLotNo & "^FS"
        sMsg = sMsg & "^FO30,160^BY2,2,55^B3,,,N^FD" & sProdNo & "^FS"
        sMsg = sMsg & "^FO30,265^BY2,2,55^B3,,,N^FD" & sWQty & "^FS"
        sMsg = sMsg & "^FO30,370^BY2,2,55^B3,,,N^FD" & sQty & "^FS"
        sMsg = sMsg & "^FO425,385^GB355,50,3^FS"
        sMsg = sMsg & "^FO425,425^GB0,50,3^FS"
        sMsg = sMsg & "^FO440,400^A0,27,25^FDDATE CODE:^FS"
        sMsg = sMsg & "^FO440,440^A0,27,25^FDQA:^FS"
        sMsg = sMsg & "^PQ1"
        sMsg = sMsg & "^XZ"
    '******
    ' End of Label Info
    '******

    '******
    'Send message to TDS
    '******
        
        'Add by Tony Start on 20110117 for Req.JC201100003
        If InStr(1, sPrinterServer, "@") <> 0 Then
            sPrinterServer = VBA.Left(sPrinterServer, (InStr(1, sPrinterServer, "@")) - 1)
        End If
        'Add by Tony End on 20110117 for Req.JC201100003
        
        If Not moCwMbx.Send("MBX_" & sPrinterServer, sMsg) Then
            Call moCwMbx.GetError(lErrCode, sErrMsg)
            Call RaiseError(lErrCode, sErrMsg)
        End If
    '******
'----
' Done
'----
ExitHandler:
    ' NOTE 1:
    ' MUST CALL GetErrInfo() here first before another action
    Call GetErrInfo(msMODULE_ID, sProcID, typErrInfo, Erl)
    Call LogProcOut(msMODULE_ID, sProcID, typErrInfo, oLogCtrl)
    On Error Resume Next

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
        Call HandleError(False, typErrInfo, , oLogCtrl, True)
    End If
End Sub

'================================================================================
' Sub: Prt_WS_WS_SUMMARY()
'--------------------------------------------------------------------------------
' Description:  <This subroutine will print out a label for FT-> SFG >
'--------------------------------------------------------------------------------
' Author:       <Alex Chan>, <CIT> 2002/02/07
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
Public Sub Prt_WS_WS_SUMMARY(ByVal sLotNo As String, _
                            ByVal sProdNo As String, _
                            ByVal sWQty As String, _
                            ByVal sCQty As String, _
                            ByVal colCompIds As FwStrings, _
                            ByVal sUserName As String, _
                            ByVal sTimeStamp As String, _
                            ByVal sPrinterServer As String, _
                            ByVal moCwMbx As Object, _
                            ByVal oLogCtrl As Object)

On Error GoTo ExitHandler:

Dim sMsg            As String
Dim sProcID         As String
Dim typErrInfo      As tErrInfo
Dim lErrCode        As Long
Dim sErrMsg         As String


'----
' Init
'----
    sProcID = "Prt_WS_WS_SUMMARY"
    Call LogProcIn(msMODULE_ID, sProcID, oLogCtrl)

'----
' Condition Checking
'----

'----
' Action
'----
    On Error Resume Next
    '******
    'Print Label Info
    '******
'mark by Nelson start on 2007/09/14 for SI Request
'        sMsg = sMsg & "^XA"
'        sMsg = sMsg & "^LH0,0^FS"
'        sMsg = sMsg & "^LL1441"
'        sMsg = sMsg & "^MD0"
'        sMsg = sMsg & "^MNY"
'        sMsg = sMsg & "^FO360,365^A0R,40,32^CI0^FR^FDLOT : " & sLotNo & "^FS"
'        sMsg = sMsg & "^FO30,50^A0N,30,30^CI0^FR^FDLOT : " & sLotNo & "^FS"
'        sMsg = sMsg & "^BY2, 2.0^FO25,85^B3N,N,60,N,Y^FR^FD" & sLotNo & "^FS"
'        sMsg = sMsg & "^FO360,770^A0R,40,32^CI0^FR^FDPROD : " & sProdNo & "^FS"
'        sMsg = sMsg & "^FO30,160^A0N,30,32^CI0^FR^FDPROD : " & sProdNo & "^FS"
'        sMsg = sMsg & "^BY1, 3.0^FO25,195^B3N,N,60,N,Y^FR^FD" & sProdNo & "^FS"
'        sMsg = sMsg & "^FO304,300^A0R,25,32^CI0^FR^FD 1:   " & colCompIds.Item("01") & "^FS"
'        sMsg = sMsg & "^FO304,500^A0R,25,32^CI0^FR^FD 6:   " & colCompIds.Item("06") & "^FS"
'        sMsg = sMsg & "^FO304,700^A0R,25,32^CI0^FR^FD11:   " & colCompIds.Item("11") & "^FS"
'        sMsg = sMsg & "^FO304,900^A0R,25,32^CI0^FR^FD16:   " & colCompIds.Item("16") & "^FS"
'        sMsg = sMsg & "^FO304,1100^A0R,25,32^CI0^FR^FD21:   " & colCompIds.Item("21") & "^FS"
'        sMsg = sMsg & "^FO268,300^A0R,25,32^CI0^FR^FD 2:   " & colCompIds.Item("02") & "^FS"
'        sMsg = sMsg & "^FO268,500^A0R,25,32^CI0^FR^FD 7:   " & colCompIds.Item("07") & "^FS"
'        sMsg = sMsg & "^FO268,700^A0R,25,32^CI0^FR^FD12:   " & colCompIds.Item("12") & "^FS"
'        sMsg = sMsg & "^FO268,900^A0R,25,32^CI0^FR^FD17:   " & colCompIds.Item("17") & "^FS"
'        sMsg = sMsg & "^FO268,1100^A0R,25,32^CI0^FR^FD22:   " & colCompIds.Item("22") & "^FS"
'        sMsg = sMsg & "^FO232,300^A0R,25,32^CI0^FR^FD 3:   " & colCompIds.Item("03") & "^FS"
'        sMsg = sMsg & "^FO232,500^A0R,25,32^CI0^FR^FD 8:   " & colCompIds.Item("08") & "^FS"
'        sMsg = sMsg & "^FO232,700^A0R,25,32^CI0^FR^FD13:   " & colCompIds.Item("13") & "^FS"
'        sMsg = sMsg & "^FO232,900^A0R,25,32^CI0^FR^FD18:   " & colCompIds.Item("18") & "^FS"
'        sMsg = sMsg & "^FO232,1100^A0R,25,32^CI0^FR^FD23:   " & colCompIds.Item("23") & "^FS"
'        sMsg = sMsg & "^FO196,300^A0R,25,32^CI0^FR^FD 4:   " & colCompIds.Item("04") & "^FS"
'        sMsg = sMsg & "^FO196,500^A0R,25,32^CI0^FR^FD 9:   " & colCompIds.Item("09") & "^FS"
'        sMsg = sMsg & "^FO196,700^A0R,25,32^CI0^FR^FD14:   " & colCompIds.Item("14") & "^FS"
'        sMsg = sMsg & "^FO196,900^A0R,25,32^CI0^FR^FD19:   " & colCompIds.Item("19") & "^FS"
'        sMsg = sMsg & "^FO196,1100^A0R,25,32^CI0^FR^FD24:   " & colCompIds.Item("24") & "^FS"
'        sMsg = sMsg & "^FO160,300^A0R,25,32^CI0^FR^FD 5:   " & colCompIds.Item("05") & "^FS"
'        sMsg = sMsg & "^FO160,500^A0R,25,32^CI0^FR^FD10:   " & colCompIds.Item("10") & "^FS"
'        sMsg = sMsg & "^FO160,700^A0R,25,32^CI0^FR^FD15:   " & colCompIds.Item("15") & "^FS"
'        sMsg = sMsg & "^FO160,900^A0R,25,32^CI0^FR^FD20:   " & colCompIds.Item("20") & "^FS"
'        sMsg = sMsg & "^FO160,1100^A0R,25,32^CI0^FR^FD25:   " & colCompIds.Item("25") & "^FS"
'        sMsg = sMsg & "^FO105,300^A0R,30,32^CI0^FR^FDTTL QTY: " & sWqty & " PCS,   " & sCqty & "EA^FS"
'        sMsg = sMsg & "^FO105,720^A0R,30,32^CI0^FR^FDID No: " & sUserName & "^FS"
'        sMsg = sMsg & "^FO105,950^A0R,30,32^CI0^FR^FDDG: 0^FS"
'        sMsg = sMsg & "^FO50,300^A0R,30,32^CI0^FR^FDDATE: " & sTimeStamp & "^FS"
'        sMsg = sMsg & "^FO50,650^A0R,30,32^CI0^FR^FDSerial No: 0206000115^FS"
'        sMsg = sMsg & "^FO50,1000^A0R,30,32^CI0^FR^FDTQA:^FS"
'        sMsg = sMsg & "^XZ"
        sMsg = sMsg & "^XA"
        sMsg = sMsg & "^LH0,0^FS"
        sMsg = sMsg & "^LL440"
        sMsg = sMsg & "^MD0"
        sMsg = sMsg & "^MNY"
        sMsg = sMsg & "^FO368,36^A0N,40.32^CI0^FR^FDLOT : " & sLotNo & "^FS"
        
'        sMsg = sMsg & "^FO48,50^A0B,30,30^CI0^FR^FDLOT : " & sLotNo & "^FS"
        sMsg = sMsg & "^FO48,10^A0B,30,30^CI0^FR^FDLOT : " & sLotNo & "^FS"

        sMsg = sMsg & "^BY2,2.0^FO80,50^B3B,N,60,N,Y^FR^FD" & sLotNo & "^FS"
        sMsg = sMsg & "^FO775,36^A0N,40,32^CI0^FR^FDPROD : " & sProdNo & "^FS"
        
'        sMsg = sMsg & "^FO160,50^A0B,30,32^CI0^FR^FDPROD : " & sProdNo & "^FS"
        sMsg = sMsg & "^FO160,10^A0B,30,32^CI0^FR^FDPROD : " & sProdNo & "^FS"
        
        sMsg = sMsg & "^BY1, 3.0^FO192,50^B3B,N,60,N,Y^FR^FD" & sProdNo & "^FS"
        sMsg = sMsg & "^FO304,108^A0N,25,32^CI0^FR^FD 1:   " & colCompIds.Item("01") & "^FS"
        sMsg = sMsg & "^FO504,108^A0N,25,32^CI0^FR^FD 6:   " & colCompIds.Item("06") & "^FS"
        sMsg = sMsg & "^FO703,108^A0N,25,32^CI0^FR^FD11:   " & colCompIds.Item("11") & "^FS"
        sMsg = sMsg & "^FO903,108^A0N,25,32^CI0^FR^FD16:   " & colCompIds.Item("16") & "^FS"
        sMsg = sMsg & "^FO1103,108^A0N,25,32^CI0^FR^FD21:   " & colCompIds.Item("21") & "^FS"
        sMsg = sMsg & "^FO304,144^A0N,25,32^CI0^FR^FD 2:   " & colCompIds.Item("02") & "^FS"
        sMsg = sMsg & "^FO504,144^A0N,25,32^CI0^FR^FD 7:   " & colCompIds.Item("07") & "^FS"
        sMsg = sMsg & "^FO703,144^A0N,25,32^CI0^FR^FD12:   " & colCompIds.Item("12") & "^FS"
        sMsg = sMsg & "^FO903,144^A0N,25,32^CI0^FR^FD17:   " & colCompIds.Item("17") & "^FS"
        sMsg = sMsg & "^FO1103,144^A0N,25,32^CI0^FR^FD22:   " & colCompIds.Item("22") & "^FS"
        sMsg = sMsg & "^FO304,180^A0N,25,32^CI0^FR^FD 3:   " & colCompIds.Item("03") & "^FS"
        sMsg = sMsg & "^FO504,180^A0N,25,32^CI0^FR^FD 8:   " & colCompIds.Item("08") & "^FS"
        sMsg = sMsg & "^FO703,180^A0N,25,32^CI0^FR^FD13:   " & colCompIds.Item("13") & "^FS"
        sMsg = sMsg & "^FO903,180^A0N,25,32^CI0^FR^FD18:   " & colCompIds.Item("18") & "^FS"
        sMsg = sMsg & "^FO1103,180^A0N,25,32^CI0^FR^FD23:   " & colCompIds.Item("23") & "^FS"
        sMsg = sMsg & "^FO304,216^A0N,25,32^CI0^FR^FD 4:   " & colCompIds.Item("04") & "^FS"
        sMsg = sMsg & "^FO504,216^A0N,25,32^CI0^FR^FD 9:   " & colCompIds.Item("09") & "^FS"
        sMsg = sMsg & "^FO703,216^A0N,25,32^CI0^FR^FD14:   " & colCompIds.Item("14") & "^FS"
        sMsg = sMsg & "^FO903,216^A0N,25,32^CI0^FR^FD19:   " & colCompIds.Item("19") & "^FS"
        sMsg = sMsg & "^FO1103,216^A0N,25,32^CI0^FR^FD24:   " & colCompIds.Item("24") & "^FS"
        sMsg = sMsg & "^FO304,252^A0N,25,32^CI0^FR^FD 5:   " & colCompIds.Item("05") & "^FS"
        sMsg = sMsg & "^FO504,252^A0N,25,32^CI0^FR^FD10:   " & colCompIds.Item("10") & "^FS"
        sMsg = sMsg & "^FO703,252^A0N,25,32^CI0^FR^FD15:   " & colCompIds.Item("15") & "^FS"
        sMsg = sMsg & "^FO903,252^A0N,25,32^CI0^FR^FD20:   " & colCompIds.Item("20") & "^FS"
        sMsg = sMsg & "^FO1103,252^A0N,25,32^CI0^FR^FD25:   " & colCompIds.Item("25") & "^FS"
        '----- Marked by Jack 2008/07/23 [M200807038] <Start> -----
        'sMsg = sMsg & "^FO300,304^A0N,30,32^CI0^FR^FDTTL QTY: " & sWqty & " PCS,   " & sCqty & "EA^FS"
        'sMsg = sMsg & "^FO723,304^A0N,30,32^CI0^FR^FDID No: " & sUserName & "^FS"
        'sMsg = sMsg & "^FO951,304^A0N,30,32^CI0^FR^FDDG: 0^FS"
        'sMsg = sMsg & "^FO300,356^A0N,30,32^CI0^FR^FDDATE: " & sTimeStamp & "^FS"
        'sMsg = sMsg & "^FO651,356^A0N,30,32^CI0^FR^FDSerial No: 0206000115^FS"
        'sMsg = sMsg & "^FO999,356^A0N,30,32^CI0^FR^FDTQA:^FS"
        'sMsg = sMsg & "^XZ"
        '----- Marked by Jack 2008/07/23 [M200807038] <End> -----
        
        '----- Added by Jack 2008/07/23 [M200807038] <Start> -----change format
        sMsg = sMsg & "^FO250,318^A0N,30,32^CI0^FR^FDTTL WQTY: " & sWQty & " PCS ^FS"
'Modified by Jack 2009/07/15 JC900084
'        sMsg = sMsg & "^BY2,2.0^FO550,318^B3N,N,40,N,Y^FR^FD" & sWqty & "^FS"
'        sMsg = sMsg & "^FO250,370^A0N,30,32^CI0^FR^FDTTL CQTY: " & sCqty & "EA ^FS"
'        sMsg = sMsg & "^BY2,2.0^FO550,370^B3N,N,40,N,Y^FR^FD" & sCqty & "^FS"
'        sMsg = sMsg & "^FO750,318^A0N,30,32^CI0^FR^FDID No: " & sUserName & "^FS"
'        sMsg = sMsg & "^FO1000,318^A0N,30,32^CI0^FR^FDTQA:^FS"
'        sMsg = sMsg & "^FO830,370^A0N,30,32^CI0^FR^FDDATE: " & sTimeStamp & "^FS"
        
        sMsg = sMsg & "^BY2,2.0^FO540,318^B3N,N,40,N,Y^FR^FD" & sWQty & "^FS"
        sMsg = sMsg & "^FO250,370^A0N,30,32^CI0^FR^FDTTL CQTY: " & sCQty & "EA ^FS"
        sMsg = sMsg & "^BY2,2.0^FO540,370^B3N,N,40,N,Y^FR^FD" & sCQty & "^FS"
        sMsg = sMsg & "^FO690,318^A0N,30,32^CI0^FR^FDID No: " & sUserName & "^FS"
        'Modify by HouYu for 201700710 on 20201109
        'sMsg = sMsg & "^FO980,318^A0N,30,32^CI0^FR^FDTQA:^FS"
        sMsg = sMsg & "^FO980,318^A0N,30,32^CI0^FR^FDTQA:       ACC^FS"
        'Modify by HouYu for 201700710 on 20201109
        sMsg = sMsg & "^FO790,370^A0N,30,32^CI0^FR^FDDATE: " & sTimeStamp & "^FS"
        
        
'Added by Jack 2009/07/15 JC900084
        sMsg = sMsg & "^FO1070,284^GB190,140,4^FS"
        sMsg = sMsg & "^XZ"
        '----- Added by Jack 2008/07/23 [M200807038] <End> -----
'mark by Nelson end on 2007/09/14 for SI Request
    '******
    ' End of Label Info
    '******

    '******
    'Send message to TDS
    '******
        
        'Add by Tony Start on 20110117 for Req.JC201100003
        If InStr(1, sPrinterServer, "@") <> 0 Then
            sPrinterServer = VBA.Left(sPrinterServer, (InStr(1, sPrinterServer, "@")) - 1)
        End If
        'Add by Tony End on 20110117 for Req.JC201100003
        
        If Not moCwMbx.Send("MBX_" & sPrinterServer, sMsg) Then
            Call moCwMbx.GetError(lErrCode, sErrMsg)
            Call RaiseError(lErrCode, sErrMsg)
        End If
    '******

'----
' Done
'----
ExitHandler:
    ' NOTE 1:
    ' MUST CALL GetErrInfo() here first before another action
    Call GetErrInfo(msMODULE_ID, sProcID, typErrInfo, Erl)
    Call LogProcOut(msMODULE_ID, sProcID, typErrInfo, oLogCtrl)
    On Error Resume Next

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
        Call HandleError(False, typErrInfo, , oLogCtrl, True)
    End If
End Sub

'================================================================================
' Sub: Prt_FT_FT_SMALL_LABEL()
'--------------------------------------------------------------------------------
' Description:  <This subroutine will print out a label for FT-> SFG >
'--------------------------------------------------------------------------------
' Author:       <Alex Chan>, <CIT> 2002/02/07
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
' [REV 01] < Tony Chang >, < MXIC > < 2011/02/15>
' 1) <Description goes here...>
'    Add CarGradeFlag , 加optional因避免有舊rule使用之卻沒改到會立即不能使用.
'    於修改需求時已確認所有使用到之rule都已修改.
'    [註] sCarGradeFlag 為必要欄位
'
' 2) sProdLine Add by Jack on 2015/06/25 for Req.JC201500212.
'
' [REV 02] <AuthorName>, <CompanyName> <YYYY/MM/DD>
' 1) <Description goes here...>
'    <Line 2...>
'Modify 增加外部傳入 sOriLotID by Ernest on 2017/11/03 for MES Phase 16  Remark 專案
'Add sFuSa by Weilun on 20220526 for ReqNo.BE#202000233
'Add sLocation by Weilun on 20220816 for ReqNo.BE#202200157
'================================================================================
Public Sub Prt_FT_FT_SMALL_LABEL(ByVal sLotNo As String, _
                                ByVal sProdNo As String, _
                                ByVal sWQty As String, _
                                ByVal sCQty As String, _
                                ByVal sOwner As String, _
                                ByVal sRouteId As String, _
                                ByVal sCustomer As String, _
                                ByVal sCsum As String, _
                                ByVal sGreen As String, _
                                ByVal sPrinterServer As String, _
                                ByVal moCwMbx As Object, _
                                ByVal oLogCtrl As Object, _
                                Optional ByVal sCargradeFlag As String, _
                                Optional ByVal sProdLine As String, _
                                Optional ByVal sOriLotID As String, _
                                Optional ByVal sFuSa As String, _
                                Optional ByVal sLocation As String)
On Error GoTo ExitHandler:

Dim sMsg            As String
Dim sProcID         As String
Dim typErrInfo      As tErrInfo
Dim lErrCode        As Long
Dim sErrMsg         As String

'----
' Init
'----
    sProcID = "Prt_FT_FT_SMALL_LABEL"
    Call LogProcIn(msMODULE_ID, sProcID, oLogCtrl)

'----
' Condition Checking
'----

'----
' Action
'----
    On Error Resume Next
       
    '******
    'Print Label Info
    '******
    
    
    sMsg = ""
    sMsg = sMsg & "^XA"
    sMsg = sMsg & "^PRC"
    sMsg = sMsg & "^LH0,0^FS"
    sMsg = sMsg & "^LL280"
    sMsg = sMsg & "^MD0"
    sMsg = sMsg & "^LH0,0^FS"
    
    If sLocation = "FT" Then 'Add by Weilun on 20220816 for ReqNo.BE#202200157
        'FT跑新的
    
        'Add by Weilun on 20220526 for ReqNo.BE#202000233 <start>
        '更動Green位置, Modify by Weilun on 20220606 for for ReqNo.BE#202000233 fix
        '(sOriLotID <> "")400,161->400,85
        'else, 400,161->400,95
        If sOriLotID <> "" Then
            sMsg = sMsg & "^FO30,5^A0N,35,27^CI13^FR^FDCarGradeFlag:" & sCargradeFlag & "^FS"
            sMsg = sMsg & "^FO315,5^A0N,35,27^CI13^FR^FDFuSa:" & sFuSa & "^FS"
            sMsg = sMsg & "^FO30,45^A0N,35,27^CI13^FR^FDLot:" & sLotNo & "^FS"
            sMsg = sMsg & "^FO315,45^A0N,35,27^CI13^FR^FD" & sProdLine & "^FS"
            sMsg = sMsg & "^BY2,3.0^FO30,85^BCN,40,N,Y,N^FR^FD>:" & sLotNo & "^FS"
            sMsg = sMsg & "^FO30,138^A0N,35,27^CI13^FR^FD" & sProdNo & "^FS"
            sMsg = sMsg & "^FO400,85^A0N,35,27^CI13^FR^FDGreen:" & sGreen & "^FS"
            sMsg = sMsg & "^FO30,178^A0N,35,27^CI13^FR^FD" & sCQty & " EA^FS"
            sMsg = sMsg & "^FO265,178^A0N,35,27^CI13^FR^FDRoute:" & sRouteId & "^FS"
            sMsg = sMsg & "^FO30,215^A0N,35,27^CI13^FR^FDOrg_Lotid:" & sOriLotID & "^FS"
            sMsg = sMsg & "^PQ1,1,1,Y"
            sMsg = sMsg & "^XZ"
        Else
        sMsg = sMsg & "^FO30,11^A0N,37,33^CI13^FR^FDCarGradeFlag:" & sCargradeFlag & "^FS"
        sMsg = sMsg & "^FO315,11^A0N,37,33^CI13^FR^FDFuSa:" & sFuSa & "^FS"
        sMsg = sMsg & "^FO30,55^A0N,37,33^CI13^FR^FDLot:" & sLotNo & "^FS"
        sMsg = sMsg & "^FO315,55^A0N,37,33^CI13^FR^FD" & sProdLine & "^FS"
        sMsg = sMsg & "^BY2,3.0^FO30,95^BCN,50,N,Y,N^FR^FD>:" & sLotNo & "^FS"
        sMsg = sMsg & "^FO30,161^A0N,37,33^CI13^FR^FD" & sProdNo & "^FS"
        sMsg = sMsg & "^FO400,95^A0N,37,33^CI13^FR^FDGreen:" & sGreen & "^FS"
        sMsg = sMsg & "^FO30,206^A0N,37,33^CI13^FR^FD" & sCQty & " EA^FS"
        sMsg = sMsg & "^FO265,206^A0N,37,33^CI13^FR^FDRoute:" & sRouteId & "^FS"
        sMsg = sMsg & "^PQ1,1,1,Y"
        sMsg = sMsg & "^XZ"
        End If
        'Add by Weilun on 20220526 for ReqNo.BE#202000233 <end>
    Else
        '新舊規則依照Location並行,UnMark by Weilun on 20220816 for ReqNo.BE#202200157 <start>
        '非FT跑原本的
        
        'Mark by Weilun on 20220526 for ReqNo.BE#202000233 <start>
        'add by Ernest on 2017/11/03 for MES Phase 16  Remark 專案----start
        If sOriLotID <> "" Then
            sMsg = sMsg & "^FO30,5^A0N,35,27^CI13^FR^FDCarGradeFlag:" & sCargradeFlag & "^FS"
            sMsg = sMsg & "^FO315,5^A0N,35,27^CI13^FR^FD" & sProdLine & "^FS"
            sMsg = sMsg & "^FO30,45^A0N,35,27^CI13^FR^FDLot:" & sLotNo & "^FS"
            sMsg = sMsg & "^FO315,45^A0N,35,27^CI13^FR^FDOwner:" & sOwner & "^FS"
            sMsg = sMsg & "^BY2,3.0^FO30,85^BCN,40,N,Y,N^FR^FD>:" & sLotNo & "^FS"
            sMsg = sMsg & "^FO350,91^A0N,35,27^CI13^FR^FDCsum:" & sCsum & "^FS"
            sMsg = sMsg & "^FO30,138^A0N,35,27^CI13^FR^FD" & sProdNo & "^FS"
            sMsg = sMsg & "^FO400,138^A0N,35,27^CI13^FR^FDGreen:" & sGreen & "^FS"
            sMsg = sMsg & "^FO30,178^A0N,35,27^CI13^FR^FD" & sCQty & " EA^FS"
            sMsg = sMsg & "^FO265,178^A0N,35,27^CI13^FR^FDRoute:" & sRouteId & "^FS"
            sMsg = sMsg & "^FO30,215^A0N,35,27^CI13^FR^FDOrg_Lotid:" & sOriLotID & "^FS"
            sMsg = sMsg & "^PQ1,1,1,Y"
            sMsg = sMsg & "^XZ"
        Else
        'add by Ernest on 2017/11/03 for MES Phase 16  Remark 專案-------end
        'sMsg = sMsg & "^FO30,11^A0N,37,33^CI13^FR^FDCustomer:" & sCustomer & "^FS"'Marked by Tony on 99/9/9 for Req.JC201000203
        sMsg = sMsg & "^FO30,11^A0N,37,33^CI13^FR^FDCarGradeFlag:" & sCargradeFlag & "^FS" 'Add by Tony on 20110215 for Req.JC201100038
        sMsg = sMsg & "^FO315,11^A0N,37,33^CI13^FR^FD" & sProdLine & "^FS" 'Add by Jack on 2015/06/25 for Req.JC201500212.
        sMsg = sMsg & "^FO30,55^A0N,37,33^CI13^FR^FDLot:" & sLotNo & "^FS"
        sMsg = sMsg & "^FO315,55^A0N,37,33^CI13^FR^FDOwner:" & sOwner & "^FS"
        sMsg = sMsg & "^BY2,3.0^FO30,95^BCN,50,N,Y,N^FR^FD>:" & sLotNo & "^FS"
        sMsg = sMsg & "^FO350,106^A0N,37,33^CI13^FR^FDCsum:" & sCsum & "^FS"
        sMsg = sMsg & "^FO30,161^A0N,37,33^CI13^FR^FD" & sProdNo & "^FS"
        sMsg = sMsg & "^FO400,161^A0N,37,33^CI13^FR^FDGreen:" & sGreen & "^FS"
        sMsg = sMsg & "^FO30,206^A0N,37,33^CI13^FR^FD" & sCQty & " EA^FS"
        sMsg = sMsg & "^FO265,206^A0N,37,33^CI13^FR^FDRoute:" & sRouteId & "^FS"
        sMsg = sMsg & "^PQ1,1,1,Y"
        sMsg = sMsg & "^XZ"
        End If
        'Mark by Weilun on 20220526 for ReqNo.BE#202000233 <end>
        'UnMark by Weilun on 20220816 for ReqNo.BE#202200157 <end>
    End If
    
'        sMsg = sMsg & "^XA"
'        sMsg = sMsg & "^PRC"
'        sMsg = sMsg & "^LH0,0^FS"
'        sMsg = sMsg & "^LL280"
'        sMsg = sMsg & "^MD0"
'        sMsg = sMsg & "^LH0,0^FS"
''---- Modify by Jeff Weng on 03/09/04 by the change from Mac Tu
''        sMsg = sMsg & "^BY2,3.0^FO15,165^BCN,56,N,Y,N^FR^FD>:" & sProdNo & "^FS"
''        sMsg = sMsg & "^BY2,3.0^FO15,56^BCN,56,N,Y,N^FR^FD>:" & sLotNo & "^FS"
''        sMsg = sMsg & "^FO250,236^A0N,37,33^CI13^FR^FDRoute:" & sRouteId & "^FS"
''        sMsg = sMsg & "^FO15,11^A0N,37,33^CI13^FR^FDLot:" & sLotNo & "^FS"
''        sMsg = sMsg & "^FO310,11^A0N,37,33^CI13^FR^FDOwner:" & sOwner & "^FS"
''        sMsg = sMsg & "^FO15,126^A0N,37,33^CI13^FR^FD" & sProdNo & "^FS"
'''        sMsg = sMsg & "^FO15,240^A0N,37,33^CI13^FR^FD" & sCqty & " EA^FS"
''        sMsg = sMsg & "^BY2,3.0^FO30,165^BCN,56,N,Y,N^FR^FD>:" & sProdNo & "^FS"
''        sMsg = sMsg & "^BY2,3.0^FO30,56^BCN,56,N,Y,N^FR^FD>:" & sLotNo & "^FS"
''        sMsg = sMsg & "^FO265,236^A0N,37,33^CI13^FR^FDRoute:" & sRouteId & "^FS"
''        sMsg = sMsg & "^FO30,11^A0N,37,33^CI13^FR^FDLot:" & sLotNo & "^FS"
''        sMsg = sMsg & "^FO325,11^A0N,37,33^CI13^FR^FDOwner:" & sOwner & "^FS"
''        sMsg = sMsg & "^FO30,126^A0N,37,33^CI13^FR^FD" & sProdNo & "^FS"
''        sMsg = sMsg & "^FO30,240^A0N,37,33^CI13^FR^FD" & sCqty & " EA^FS"
''----- Enf of Modify
'        sMsg = sMsg & "^PQ1,1,1,Y"
'        sMsg = sMsg & "^XZ"
    '******
    ' End of Label Info
    '******

    '******
    'Send message to TDS
    '******
        'Add by Tony Start on 20110117 for Req.JC201100003
        If InStr(1, sPrinterServer, "@") <> 0 Then
            sPrinterServer = VBA.Left(sPrinterServer, (InStr(1, sPrinterServer, "@")) - 1)
        End If
        'Add by Tony End on 20110117 for Req.JC201100003
        
        If Not moCwMbx.Send("MBX_" & sPrinterServer, sMsg) Then
            Call moCwMbx.GetError(lErrCode, sErrMsg)
            Call RaiseError(lErrCode, sErrMsg)
        End If
    '******
'----
' Done
'----
ExitHandler:
    ' NOTE 1:
    ' MUST CALL GetErrInfo() here first before another action
    Call GetErrInfo(msMODULE_ID, sProcID, typErrInfo, Erl)
    Call LogProcOut(msMODULE_ID, sProcID, typErrInfo, oLogCtrl)
    On Error Resume Next

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
        Call HandleError(False, typErrInfo, , oLogCtrl, True)
    End If
End Sub
'================================================================================
' Sub: Prt_WS_WS_SMALL_LABEL()
'--------------------------------------------------------------------------------
' Description:  <This subroutine will print out a label for FT-> SFG >
'--------------------------------------------------------------------------------
' Author:       <Alex Chan>, <CIT> 2002/02/07
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
''Modify by Sam on 20080805 for ReqNo:M200807064 Add sBizType,remove sSpeed
'Modify by Sam on 20080902 for ReqNo:M200809003 ADD moProRawSql,Function GetBizType
'Modify by Sam on 20110823 for Project Hold Lot Release,增加外部傳入WaferID
'Modified by Jack on 2014/06/03 for JC201400147. 增加外部傳入 sFabLotId.
'================================================================================
Public Sub Prt_WS_WS_SMALL_LABEL(ByVal sLotNo As String, _
                                ByVal sProdNo As String, _
                                ByVal sWQty As String, _
                                ByVal sCQty As String, _
                                ByVal sOwner As String, _
                                ByVal sBizType As String, _
                                ByVal sPrinterServer As String, _
                                ByVal moCwMbx As Object, _
                                ByVal oLogCtrl As Object, _
                                ByVal moProRawSql As Object, _
                                Optional ByVal sWaferID As String, _
                                Optional ByVal sFabLotID As String)


On Error GoTo ExitHandler:

Dim sMsg            As String
Dim sProcID         As String
Dim typErrInfo      As tErrInfo
Dim lErrCode        As Long
Dim sErrMsg         As String

'Add by Sam start on 20101019 for Project RunCard Paperless
Dim sSQL As String
Dim colRS As Collection
Dim sHotLotFlag As String
Dim sErunTicNO As String
Dim sSapRwNo As String
Dim sMfgTicNo As String

Dim sWaferNo1 As String
Dim sWaferNo2 As String
'Add by Sam End on 20101019 for Project RunCard Paperless

'Add by Sam start on 20120217 for ReqNo:JC201200047
Dim iIdx As Integer
Dim iIdx2 As Integer
Dim sWaferNO(25) As String
Dim vWaferNO As Variant
Dim sSortWaferNo As String
Dim bFound As Boolean
'Add by Sam End on 20120217 for ReqNo:JC201200047

Dim sPrintFabLotId As String 'Added by Jack on 2014/06/03 for JC201400147.
'Added by Jack on 2016/08/30 for JC201600173 <Start>
Dim sCriteiraString   As String
Dim sCriteiraData     As String
'Added by Jack on 2016/08/30 for JC201600173 <End>

'----
' Init
'----
    sProcID = "Prt_WS_WS_SMALL_LABEL"
    Call LogProcIn(msMODULE_ID, sProcID, oLogCtrl)
   'Modified by Jack on 2016/08/01 for JC201600108 <Start>
    'BizType改為ProdLevel串ProdType.
    'sBizType = GetBizType(sProdNo, oLogCtrl, moProRawSql)
    'Modified by Jack on 2016/08/30 for JC201600173 <Start>
    'sBizType = GetProdType(sProdNo, oLogCtrl, moProRawSql)
    sCriteiraString = "" '此處只有列印, 不需傳入.
    sBizType = GetProdType(sProdNo, sCriteiraString, sCriteiraData, oLogCtrl, moProRawSql)
    'Modified by Jack on 2016/08/30 for JC201600173 <End>
    'Modified by Jack on 2016/08/01 for JC201600108 <End>
    
    'gsCAT_TLI_FABLOTID added by Jack on 2014/06/03 for JC201400147.
    sSQL = " select A." & gsCAT_TLATT_HOTLOTFLAG & " ,B." & gsCAT_TLI_ERUNTICNO & _
           " ,B." & gsCAT_TLI_SAPRWNO & " ,B." & gsCAT_TLI_MFGTICNO & _
           " ,B." & gsCAT_TLI_WAFERID & _
           " ,B." & gsCAT_TLI_FABLOTID & _
           " FROM " & gsCAT_TBL_LOT_ATTRIBUTE & " A, " & gsCAT_TBL_LOT_INFO & " B " & _
           " WHERE A." & gsCAT_TLATT_LOTID & "= B." & gsCAT_TLI_LOT_ID & " " & _
           " AND  A." & gsCAT_TLATT_LOTID & "='" & sLotNo & "' "
    Set colRS = moProRawSql.QueryDatabase(sSQL)
    If colRS.Count > 0 Then
        sHotLotFlag = colRS.Item(1).Item(gsCAT_TLATT_HOTLOTFLAG)
        sErunTicNO = colRS.Item(1).Item(gsCAT_TLI_ERUNTICNO)
        sSapRwNo = colRS.Item(1).Item(gsCAT_TLI_SAPRWNO)
        sMfgTicNo = colRS.Item(1).Item(gsCAT_TLI_MFGTICNO)
        
        'Modify by Sam on 20110803,如果有外部傳入Wafer ID時優先使用
'         sWaferNo1 = Replace(colRS.Item(1).Item(gsCAT_TLI_WAFERID), ";", ",")
        If sWaferID <> "" Then
            sWaferNo1 = Replace(sWaferID, ";", ",")
        Else
            sWaferNo1 = Replace(colRS.Item(1).Item(gsCAT_TLI_WAFERID), ";", ",")
        End If
        
        'Modified by Jack on 2014/06/03 for JC201400147. 增加外部傳入 sFabLotId <Start>.
        '如果有外部傳FabLotID時優先使用
        If sFabLotID <> "" Then
            sPrintFabLotId = sFabLotID
        Else
            sPrintFabLotId = colRS.Item(1).Item(gsCAT_TLI_FABLOTID)
        End If
        'Modified by Jack on 2014/06/03 for JC201400147. 增加外部傳入 sFabLotId <End>.
        
        'Add by Sam start on 20120217 for ReqNo:JC201200047
        sSortWaferNo = SortWaferID(sWaferNo1)
        vWaferNO = Split(sWaferNo1, ",")
        sWaferNo1 = ""
        'For iIdx = LBound(sWaferNo) To UBound(sWaferNo)
        For iIdx = 1 To 25
            bFound = False
            For iIdx2 = LBound(vWaferNO) To UBound(vWaferNO)
                If Val(vWaferNO(iIdx2)) = iIdx Then
                    bFound = True
                    Exit For
                End If
            Next
            If bFound = True Then
                sWaferNo1 = sWaferNo1 & "," & Format(iIdx, "00")
            Else
                sWaferNo1 = sWaferNo1 & ",__"
            End If
        Next
        sWaferNo1 = Mid(sWaferNo1, 2)
        'Add by Sam End on 20120217 for ReqNo:JC201200047
        
        If Len(sWaferNo1) > 38 Then
            sWaferNo2 = Mid(sWaferNo1, 40)
            sWaferNo1 = Mid(sWaferNo1, 1, 38)
        End If
    Else
        'Modified by Jack on 2014/06/03 for JC201400147. 增加外部傳入 sFabLotId <Start>.
        '如果有外部傳入Wafer ID時優先使用
        If sWaferID <> "" Then
            sWaferNo1 = Replace(sWaferID, ";", ",")
        End If
        sSortWaferNo = SortWaferID(sWaferNo1)
        vWaferNO = Split(sWaferNo1, ",")
        sWaferNo1 = ""
        'For iIdx = LBound(sWaferNo) To UBound(sWaferNo)
        For iIdx = 1 To 25
            bFound = False
            For iIdx2 = LBound(vWaferNO) To UBound(vWaferNO)
                If Val(vWaferNO(iIdx2)) = iIdx Then
                    bFound = True
                    Exit For
                End If
            Next
            If bFound = True Then
                sWaferNo1 = sWaferNo1 & "," & Format(iIdx, "00")
            Else
                sWaferNo1 = sWaferNo1 & ",__"
            End If
        Next
        sWaferNo1 = Mid(sWaferNo1, 2)
        'Add by Sam End on 20120217 for ReqNo:JC201200047
        
        If Len(sWaferNo1) > 38 Then
            sWaferNo2 = Mid(sWaferNo1, 40)
            sWaferNo1 = Mid(sWaferNo1, 1, 38)
        End If
        '如果有外部傳FabLotID時優先使用
        If sFabLotID <> "" Then
            sPrintFabLotId = sFabLotID
        End If
        'Modified by Jack on 2014/06/03 for JC201400147. 增加外部傳入 sFabLotId <End>.
    End If
'----
' Condition Checking
'----

'----
' Action
'----
    On Error Resume Next
    '******
    'Print Label Info
    '******
        'Modify by Sam on 20101019 for RunCard Paperless,換新Label
'
'        sMsg = sMsg & "^XA"
'        sMsg = sMsg & "^PRC"
'        sMsg = sMsg & "^LH0,0^FS"
'        sMsg = sMsg & "^LL280"
'        sMsg = sMsg & "^MD0"
'        sMsg = sMsg & "^LH0,0^FS"
'        sMsg = sMsg & "^BY2,3.0^FO15,165^BCN,56,N,Y,N^FR^FD>:" & sProdNo & "^FS"
'        sMsg = sMsg & "^BY2,3.0^FO15,56^BCN,56,N,Y,N^FR^FD>:" & sLotNo & "^FS"
''        sMsg = sMsg & "^FO326,240^A0N,37,33^CI13^FR^FDSpeed: " & sSpeed & "^FS" 'Mark by Sam on 20080805 for ReqNo:M200807064
'        sMsg = sMsg & "^FO15,11^A0N,37,33^CI13^FR^FDLot No: " & sLotNo & "^FS"
''        sMsg = sMsg & "^FO310,11^A0N,37,33^CI13^FR^FDOwner: " & sOwner & "^FS" 'Marked by Jack 2009/07/15 JC900084 取消LotOwner
'        sMsg = sMsg & "^FO320,240^A0N,37,33^CI13^FR^FDBiz Type: " & sBizType & "^FS" 'add by Sam on 20080805 for ReqNo:M200807064
'        sMsg = sMsg & "^FO15,126^A0N,37,33^CI13^FR^FDProd No: " & sProdNo & "^FS"
'        sMsg = sMsg & "^FO15,240^A0N,37,33^CI13^FR^FD" & sCqty & " EA / " & sWqty & " PCS^FS"
'        sMsg = sMsg & "^PQ1,1,1,Y"
'        sMsg = sMsg & "^XZ"

        sMsg = sMsg & "^XA"
        sMsg = sMsg & "^PRC"
        sMsg = sMsg & "^LH0,0^FS"
        sMsg = sMsg & "^LL280"
        sMsg = sMsg & "^MD0"
        sMsg = sMsg & "^LH0,0^FS"
        sMsg = sMsg & "^FO15,15^AFN,10,10^CI13^FDLotNo:^FS"
        sMsg = sMsg & "^FO130,15^A0N,37,33^CI13^FD" & sLotNo & "^FS"
        sMsg = sMsg & "^FO400,15^AFN,10,10^CI13^FDHotFlag:^FS"
        sMsg = sMsg & "^FO550,15^CF0N,37,33^CI13^FD" & sHotLotFlag & "^FS"
        'Modified by Jack on 2014/03/24 for JC201400088. <Start>
        'a. 原有LotNo 條碼高度減半 (56-->28; 120-->90; 170-->140)
        'sMsg = sMsg & "^FO15,50^BY2,3.0^BCN,56,N,Y,N^FR^FD>:" & sLotNo & "^FS"
        'sMsg = sMsg & "^FO15,120^AFN,10,10^CI13^FDProdNo:^FS"
        'sMsg = sMsg & "^FO145,120^A0N,37,33^CI13^FD" & sProdNo & "^FS"
        'sMsg = sMsg & "^FO15,170^AFN,10,10^CI13^FDQTY:^FS"
        sMsg = sMsg & "^FO15,50^BY2,3.0^BCN,28,N,Y,N^FR^FD>:" & sLotNo & "^FS"
                
        'Added by Jack on 2014/06/03 for JC201400147 <Start>
        sMsg = sMsg & "^FO15,90^AFN,10,10^CI13^FDFabLotId:^FS"
        sMsg = sMsg & "^FO160,90^A0N,37,33^CI13^FD" & sPrintFabLotId & "^FS"
        'Added by Jack on 2014/06/03 for JC201400147 <End>
        
        'Modified by Jack on 2014/06/03 for JC201400147 <Start>
        'sMsg = sMsg & "^FO15,90^AFN,10,10^CI13^FDProdNo:^FS"
        'sMsg = sMsg & "^FO145,90^A0N,37,33^CI13^FD" & sProdNo & "^FS"
        'sMsg = sMsg & "^FO15,140^AFN,10,10^CI13^FDQTY:^FS"
        sMsg = sMsg & "^FO15,140^AFN,10,10^CI13^FDProdNo:^FS"
        sMsg = sMsg & "^FO145,140^A0N,37,33^CI13^FD" & sProdNo & "^FS"
        sMsg = sMsg & "^FO15,190^AFN,10,10^CI13^FDQTY:^FS"
        'Modified by Jack on 2014/06/03 for JC201400147 <End>
        'Modified by Jack on 2014/03/24 for JC201400088. <End>
        
        'Modified by Jack on 2014/03/24 for JC201400088. <Start>
        'b. 變更QTY 的PCS與EA位置(拆成二行)
        'c. 新增PCS與EA條碼
        'sMsg = sMsg & "^FO90,170^A0N,37,33^CI13^FR^FD" & sCQty & " EA / " & sWQty & " PCS^FS"
        'Modified by Jack on 2014/06/03 for JC201400147 <Start>
        'sMsg = sMsg & "^FO90,140^A0N,37,33^CI13^FR^FD" & sWqty & " PCS^FS"
        'sMsg = sMsg & "^FO300,140^BY2,3.0^BCN,28,N,Y,N^FR^FD" & sWqty & "^FS"
        'sMsg = sMsg & "^FO90,180^A0N,37,33^CI13^FR^FD" & sCQty & " EA^FS"
        'sMsg = sMsg & "^FO300,180^BY2,3.0^BCN,28,N,Y,N^FR^FD" & sCQty & "^FS"
        sMsg = sMsg & "^FO90,190^A0N,37,33^CI13^FR^FD" & sWQty & " PCS^FS"
        sMsg = sMsg & "^FO300,190^BY2,3.0^BCN,28,N,Y,N^FR^FD" & sWQty & "^FS"
        sMsg = sMsg & "^FO90,230^A0N,37,33^CI13^FR^FD" & sCQty & " EA^FS"
        sMsg = sMsg & "^FO300,230^BY2,3.0^BCN,28,N,Y,N^FR^FD" & sCQty & "^FS"
        'Modified by Jack on 2014/06/03 for JC201400147 <End>
        'Modified by Jack on 2014/03/24 for JC201400088. <End>
        
        'Modified by Jack on 2014/06/03 for JC201400147 <Start>
        'sMsg = sMsg & "^FO15,220^AFN,10,10^CI13^FDErunTicNo:^FS"
        'sMsg = sMsg & "^FO170,220^A0N,37,33^CI13^FD" & sErunTicNo & "^FS"
        'sMsg = sMsg & "^FO400,220^AFN,10,10^CI13^FDBizType:^FS"
        'sMsg = sMsg & "^FO600,220^A0N,37,33^CI13^FR^FD" & sBizType & "^FS"
        sMsg = sMsg & "^FO15,270^AFN,10,10^CI13^FDErunTicNo:^FS"
        sMsg = sMsg & "^FO170,270^A0N,37,33^CI13^FD" & sErunTicNO & "^FS"
        
        'Modified by Jack on 2016/08/01 for JC201600108 <Start>
        'BizType改為ProdType.
        'sMsg = sMsg & "^FO400,270^AFN,10,10^CI13^FDBizType:^FS"
        sMsg = sMsg & "^FO400,270^AFN,10,10^CI13^FDProdType:^FS"
        'Modified by Jack on 2016/08/01 for JC201600108 <End>
        
        sMsg = sMsg & "^FO600,270^A0N,37,33^CI13^FR^FD" & sBizType & "^FS"
        'Modified by Jack on 2014/06/03 for JC201400147 <End>
        
        'Modified by Jack on 2014/06/03 for JC201400147 <Start>
        'sMsg = sMsg & "^FO15,270^AFN,10,10^CI13^FDSapRwNo:^FS"
        'sMsg = sMsg & "^FO170,270^A0N,37,33^CI13^FD" & sSapRwNo & "^FS"
        'sMsg = sMsg & "^FO15,320^AFN,10,10^CI13^FDMfgTicNo:^FS"
        'sMsg = sMsg & "^FO170,320^A0N,37,33^CI13^FD" & sMfgTicNo & "^FS"
        sMsg = sMsg & "^FO15,320^AFN,10,10^CI13^FDSapRwNo:^FS"
        sMsg = sMsg & "^FO170,320^A0N,37,33^CI13^FD" & sSapRwNo & "^FS"
        sMsg = sMsg & "^FO400,320^AFN,10,10^CI13^FDMfgTicNo:^FS"
        sMsg = sMsg & "^FO555,320^A0N,37,33^CI13^FD" & sMfgTicNo & "^FS"
        'Modified by Jack on 2014/06/03 for JC201400147 <End>
        
        sMsg = sMsg & "^FO15,370^AFN,10,10^CI13^FDID#:^FS"
        sMsg = sMsg & "^FO150,370^A0N,37,33^CI13^FD" & sWaferNo1 & "^FS"
        sMsg = sMsg & "^FO150,420^A0N,37,33^CI13^FD" & sWaferNo2 & "^FS"
        sMsg = sMsg & "^PQ1,1,1,Y"
        sMsg = sMsg & "^XZ"

    '******
    ' End of Label Info
    '******

    '******
    'Send message to TDS
    '******
        
        'Add by Tony Start on 20110117 for Req.JC201100003
        If InStr(1, sPrinterServer, "@") <> 0 Then
            sPrinterServer = VBA.Left(sPrinterServer, (InStr(1, sPrinterServer, "@")) - 1)
        End If
        'Add by Tony End on 20110117 for Req.JC201100003
        
        If Not moCwMbx.Send("MBX_" & sPrinterServer, sMsg) Then
            Call moCwMbx.GetError(lErrCode, sErrMsg)
            Call RaiseError(lErrCode, sErrMsg)
        End If
    '******

'----
' Done
'----
ExitHandler:
    ' NOTE 1:
    ' MUST CALL GetErrInfo() here first before another action
    Call GetErrInfo(msMODULE_ID, sProcID, typErrInfo, Erl)
    Call LogProcOut(msMODULE_ID, sProcID, typErrInfo, oLogCtrl)
    On Error Resume Next

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
        Call HandleError(False, typErrInfo, , oLogCtrl, True)
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
'================================================================================
'Marked by Jack on 2016/08/01 for JC201600108 <Start>
'BizType改為ProdLevel串ProdType.
'Public Function GetBizType(ByVal sIPN As String, ByVal oLogCtrl As Object, ByVal moProRawSql As Object) As String
'On Error GoTo ExitHandler:
'Dim sProcID As String
'Dim typErrInfo As tErrInfo
'Dim sTable As String
'Dim sColumn As String
'Dim sWhere As String
'Dim sOther  As String
'Dim colRaws As Collection
'Dim oRaws As FwStrings
'Dim sSql As String
''----
'' Init
''----
'    sProcID = "GetBizType"
'    Call LogProcIn(msMODULE_ID, sProcID, oLogCtrl)
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
'    sWhere = "A." & gsCAT_TIM_PRODBODY & "=B." & gsCAT_TPB_PROD_BODY & "(+) and " & _
'             "A." & gsCAT_TIM_IPN & "='" & sIPN & "' and " & _
'             "A." & gsCAT_TIM_DELETE_FLAG & "='N'"
'
'    sSql = "SELECT " & sColumn & " FROM " & sTable & " WHERE " & sWhere
'
''    Set colRaws = SelectRawSqlTxn(moAppLog, moFwWIP, moFwWF, moCwMbx, "CAT", sTable, sColumn, sWhere, sOther)
'    Set colRaws = moProRawSql.QueryDatabase(sSql)
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
'        'Modified by Jack on 2012/06/13 for support KGD_AEB Project [Hanno]
'        'KGD變更條件 : BIZTYPE = K or L
'        If colRaws.Item(1).Item(gsCAT_TPB_SBU_LINE) = "NVM" And _
'           colRaws.Item(1).Item(gsCAT_TIM_PACKAGE_CODE) = "W" And _
'           (colRaws.Item(1).Item(gsCAT_TIM_BIZTYPE) = "K" Or colRaws.Item(1).Item(gsCAT_TIM_BIZTYPE) = "L") Then
'           GetBizType = "KGD"
'        End If
'
'        'Add by Sam start on 20140729 for ReqNo:JC201400240 ,增加NVM,W,D->KTD
'        If colRaws.Item(1).Item(gsCAT_TPB_SBU_LINE) = "NVM" And _
'           colRaws.Item(1).Item(gsCAT_TIM_PACKAGE_CODE) = "W" And _
'           colRaws.Item(1).Item(gsCAT_TIM_BIZTYPE) = "D" Then
'           GetBizType = "KTD"
'        End If
'        'Add by Sam End on 20140729 for ReqNo:JC201400240
'    End If
'
''----
'' Done
''----
'
'ExitHandler:
'    ' NOTE 1:
'    ' MUST CALL GetErrInfo() here first before another action
'    Call GetErrInfo(msMODULE_ID, sProcID, typErrInfo, Erl)
'    Call LogProcOut(msMODULE_ID, sProcID, typErrInfo, oLogCtrl)
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
'        Call HandleError(True, typErrInfo, , oLogCtrl)
'    End If
'End Function
'Marked by Jack on 2016/08/01 for JC201600108 <End>

'================================================================================
' function: GetLabelPrintQty()
'--------------------------------------------------------------------------------
' Description:  <Type your function description here...>
'--------------------------------------------------------------------------------
' Author:       Robin, 2010-01-11, ReqNo:JC201000012
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
Private Function GetLabelPrintQty(ByVal sStage As String, ByVal sLabel As String, ByVal oLogCtrl As Object, ByVal moProRawSql As Object) As String
On Error GoTo ExitHandler:
Dim sProcID As String
Dim typErrInfo As tErrInfo
Dim sTable As String
Dim sColumn As String
Dim sWhere As String
Dim sOther  As String
Dim colRaws As Collection
Dim oRaws As FwStrings
Dim sSQL As String
'----
' Init
'----
    sProcID = "GetLabelPrintQty"
    Call LogProcIn(msMODULE_ID, sProcID, oLogCtrl)
    GetLabelPrintQty = 1
'----
' Condition Chking
'----


'----
' Action
'----
    sTable = gsCAT_TBL_LABEL_SPEC

    sColumn = gsCAT_TLS_PRINT_QTY
    
    sWhere = gsCAT_TLS_STAGE & " = '" & sStage & "' " & _
             " and " & gsCAT_TLS_LABEL_SPECNO & " = '" & sLabel & "'"

    sSQL = "SELECT " & sColumn & " FROM " & sTable & " WHERE " & sWhere
    
    Set colRaws = moProRawSql.QueryDatabase(sSQL)
    If colRaws Is Nothing Then
        Call RaiseError(glERR_INVALIDOBJECT, _
                    FormatErrorText(gsETX_INVALIDOBJECT, "Collection"))
    End If
    
    If colRaws.Count > 0 Then
        GetLabelPrintQty = colRaws.Item(1).Item(1)
    End If
'----
' Done
'----

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
            Case glERR_INVALIDOBJECT
                ' Retry code goes here...
            Case Else
                typErrInfo.sUserText = "Fail to execute application, please call IT support!!" & vbCrLf & _
                                        "程式執行失敗, 請洽IT人員處理"
            End Select
        '---- Start of Select Case Block ----
        On Error GoTo ExitHandler:
        Call HandleError(True, typErrInfo, , oLogCtrl)
    End If
End Function

'================================================================================
' Sub: Prt_FT_To_SubMK_Normal()
'--------------------------------------------------------------------------------
' Description:  <This subroutine will print out a label for FT-> SFG >
'--------------------------------------------------------------------------------
' Author:       <Alex Chan>, <CIT> 2002/02/07
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
Public Sub Prt_FT_To_SubMK_Normal(ByVal sLotNo As String, _
                                ByVal sProdNo As String, _
                                ByVal sQty As String, _
                                ByVal sPacker As String, _
                                ByVal bIsPartial As Boolean, _
                                ByVal sPrinterServer As String, _
                                ByVal moCwMbx As Object, _
                                ByVal oLogCtrl As Object)


On Error GoTo ExitHandler:

Dim sMsg            As String
Dim sProcID         As String
Dim typErrInfo      As tErrInfo
Dim lErrCode        As Long
Dim sErrMsg         As String

'----
' Init
'----
    sProcID = "Prt_FT_To_SubMK_Normal"
    Call LogProcIn(msMODULE_ID, sProcID, oLogCtrl)

'----
' Condition Checking
'----

'----
' Action
'----
    '******
    'Print Label Info
    '******
        sMsg = sMsg & "^XA^LH00,00^LL550"
        sMsg = sMsg & "^FO10,05^GB780,475,3^FS"
        sMsg = sMsg & "^FO435,327^GB355,70,3^FS"
        sMsg = sMsg & "^FO435,397^GB0,80,3^FS"
        sMsg = sMsg & "^FO580,30^GB200,75,3^FS"
        sMsg = sMsg & "^FO590,55^A0N,35,50^CI0^FDREMARK^FS"
        If bIsPartial Then
            sMsg = sMsg & "^FO580,102^GB200,75,3^FS"
            sMsg = sMsg & "^FO590,130^A0N,35,50^CI0^FDPARTIAL^FS"
        End If
'        sMsg = sMsg & "^FO580,250^GB200,75,3^FS"
'        sMsg = sMsg & "^FO640,275^A0N,35,50^CI0^FDENG^FS"
        sMsg = sMsg & "^FO40,58^A0N,33,28^FD(1T)LOT NO: " & sLotNo & "^FS"
        sMsg = sMsg & "^FO40,158^A0N,33,28^FD(1P)PRODUCT NO: " & sProdNo & "^FS"
        sMsg = sMsg & "^FO40,258^A0N,33,28^FD(Q)QUANTITY: " & sQty & "^FS"
        sMsg = sMsg & "^FO40,90^BY2,2,50^B3,,,N^FD1T" & sLotNo & "^FS"
        sMsg = sMsg & "^FO40,190^BY2,2,50^B3,,,N^FD1P" & sProdNo & "^FS"
        sMsg = sMsg & "^FO40,290^BY2,2,50^B3,,,N^FDQ" & sQty & "^FS"
        sMsg = sMsg & "^FO445,350^A0N,36,30^FDPACKER: " & sPacker & "^FS"
        sMsg = sMsg & "^FO445,420^A0N,36,30^FDQA:^FS"
        sMsg = sMsg & "^PQ1"
        sMsg = sMsg & "^XZ1"
    '******
    ' End of Label Info
    '******

    '******
    'Send message to TDS
    '******
        
        'Add by Tony Start on 20110117 for Req.JC201100003
        If InStr(1, sPrinterServer, "@") <> 0 Then
            sPrinterServer = VBA.Left(sPrinterServer, (InStr(1, sPrinterServer, "@")) - 1)
        End If
        'Add by Tony End on 20110117 for Req.JC201100003
        
        If Not moCwMbx.Send("MBX_" & sPrinterServer, sMsg) Then
            Call moCwMbx.GetError(lErrCode, sErrMsg)
            Call RaiseError(lErrCode, sErrMsg)
        End If
    '******

'----
' Done
'----
ExitHandler:
    ' NOTE 1:
    ' MUST CALL GetErrInfo() here first before another action
    Call GetErrInfo(msMODULE_ID, sProcID, typErrInfo, Erl)
    Call LogProcOut(msMODULE_ID, sProcID, typErrInfo, oLogCtrl)
    On Error Resume Next

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
        Call HandleError(False, typErrInfo, , oLogCtrl, True)
    End If
End Sub



'================================================================================
' Sub: Prt_FT_To_SubFT_Eng_Sample()
'--------------------------------------------------------------------------------
' Description:  <This subroutine will print out a label for FT-> SFG >
'--------------------------------------------------------------------------------
' Author:       <Alex Chan>, <CIT> 2002/02/07
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
Public Sub Prt_FT_To_SubFT_Eng_Sample(ByVal sLotNo As String, _
                                ByVal sProdNo As String, _
                                ByVal sQty As String, _
                                ByVal sPacker As String, _
                                ByVal bIsPartial As Boolean, _
                                ByVal sPrinterServer As String, _
                                ByVal moCwMbx As Object, _
                                ByVal oLogCtrl As Object)
On Error GoTo ExitHandler:

Dim sMsg            As String
Dim sProcID         As String
Dim typErrInfo      As tErrInfo
Dim lErrCode        As Long
Dim sErrMsg         As String

'----
' Init
'----
    sProcID = "Prt_FT_To_SubFT_Eng_Sample"
    Call LogProcIn(msMODULE_ID, sProcID, oLogCtrl)

'----
' Condition Checking
'----

'----
' Action
'----
    '******
    'Print Label Info
    '******
        sMsg = sMsg & "^XA^LH00,00^LL550"
        sMsg = sMsg & "^FO10,05^GB780,475,3^FS"
        sMsg = sMsg & "^FO435,327^GB355,70,3^FS"
        sMsg = sMsg & "^FO435,397^GB0,80,3^FS"
        sMsg = sMsg & "^FO580,30^GB200,75,3^FS"
'        sMsg = sMsg & "^FO635,55^A0N,35,50^CI0^FDTEST^FS"
        sMsg = sMsg & "^FO590,55^A0N,35,50^CI0^FDREMARK^FS"
        If bIsPartial Then
            sMsg = sMsg & "^FO580,102^GB200,75,3^FS"
            sMsg = sMsg & "^FO590,130^A0N,35,50^CI0^FDPARTIAL^FS"
        End If
        sMsg = sMsg & "^FO580,250^GB200,75,3^FS"
        sMsg = sMsg & "^FO640,275^A0N,35,50^CI0^FDENG^FS"
        sMsg = sMsg & "^FO40,58^A0N,33,28^FD(1T)LOT NO: 123456789012345^FS"
        sMsg = sMsg & "^FO40,158^A0N,33,28^FD(1P)PRODUCT NO: " & sProdNo & "^FS"
        sMsg = sMsg & "^FO40,258^A0N,33,28^FD(Q)QUANTITY: " & sQty & "^FS"
        sMsg = sMsg & "^FO40,90^BY2,2,50^B3,,,N^FD1T" & sLotNo & "^FS"
        sMsg = sMsg & "^FO40,190^BY2,2,50^B3,,,N^FD1P" & sProdNo & "^FS"
        sMsg = sMsg & "^FO40,290^BY2,2,50^B3,,,N^FDQ" & sQty & "^FS"
        sMsg = sMsg & "^FO445,350^A0N,36,30^FDPACKER: " & sPacker & "^FS"
        sMsg = sMsg & "^FO445,420^A0N,36,30^FDQA:^FS"
        sMsg = sMsg & "^PQ1"
        sMsg = sMsg & "^XZ1"
    '******
    ' End of Label Info
    '******

    '******
    'Send message to TDS
    '******
        
        'Add by Tony Start on 20110117 for Req.JC201100003
        If InStr(1, sPrinterServer, "@") <> 0 Then
            sPrinterServer = VBA.Left(sPrinterServer, (InStr(1, sPrinterServer, "@")) - 1)
        End If
        'Add by Tony End on 20110117 for Req.JC201100003
        
        If Not moCwMbx.Send("MBX_" & sPrinterServer, sMsg) Then
            Call moCwMbx.GetError(lErrCode, sErrMsg)
            Call RaiseError(lErrCode, sErrMsg)
        End If
    '******

'----
' Done
'----
ExitHandler:
    ' NOTE 1:
    ' MUST CALL GetErrInfo() here first before another action
    Call GetErrInfo(msMODULE_ID, sProcID, typErrInfo, Erl)
    Call LogProcOut(msMODULE_ID, sProcID, typErrInfo, oLogCtrl)
    On Error Resume Next

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
        Call HandleError(False, typErrInfo, , oLogCtrl, True)
    End If
End Sub



'================================================================================
' Sub: Prt_FT_To_SubMK_Eng_Sample()
'--------------------------------------------------------------------------------
' Description:  <This subroutine will print out a label for FT-> SFG >
'--------------------------------------------------------------------------------
' Author:       <Alex Chan>, <CIT> 2002/02/07
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
Public Sub Prt_FT_To_SubMK_Eng_Sample(ByVal sLotNo As String, _
                                ByVal sProdNo As String, _
                                ByVal sQty As String, _
                                ByVal sPacker As String, _
                                ByVal bIsPartial As Boolean, _
                                ByVal sPrinterServer As String, _
                                ByVal moCwMbx As Object, _
                                ByVal oLogCtrl As Object)
On Error GoTo ExitHandler:

Dim sMsg            As String
Dim sProcID         As String
Dim typErrInfo      As tErrInfo
Dim lErrCode        As Long
Dim sErrMsg         As String
'----
' Init
'----
    sProcID = "Prt_FT_To_SubMK_Eng_Sample"
    Call LogProcIn(msMODULE_ID, sProcID, oLogCtrl)

'----
' Condition Checking
'----

'----
' Action
'----
    '******
    'Print Label Info
    '******
        sMsg = sMsg & "^XA^LH00,00^LL550"
        sMsg = sMsg & "^FO10,05^GB780,475,3^FS"
        sMsg = sMsg & "^FO435,327^GB355,70,3^FS"
        sMsg = sMsg & "^FO435,397^GB0,80,3^FS"
        sMsg = sMsg & "^FO580,30^GB200,75,3^FS"
        sMsg = sMsg & "^FO590,55^A0N,35,50^CI0^FDREMARK^FS"
        If bIsPartial Then
            sMsg = sMsg & "^FO580,102^GB200,75,3^FS"
            sMsg = sMsg & "^FO590,130^A0N,35,50^CI0^FDPARTIAL^FS"
        End If
        sMsg = sMsg & "^FO580,250^GB200,75,3^FS"
        sMsg = sMsg & "^FO640,275^A0N,35,50^CI0^FDENG^FS"
        sMsg = sMsg & "^FO40,58^A0N,33,28^FD(1T)LOT NO: " & sLotNo & "^FS"
        sMsg = sMsg & "^FO40,158^A0N,33,28^FD(1P)PRODUCT NO: " & sProdNo & "^FS"
        sMsg = sMsg & "^FO40,258^A0N,33,28^FD(Q)QUANTITY: " & sQty & "^FS"
        sMsg = sMsg & "^FO40,90^BY2,2,50^B3,,,N^FD1T" & sLotNo & "^FS"
        sMsg = sMsg & "^FO40,190^BY2,2,50^B3,,,N^FD1P" & sProdNo & "^FS"
        sMsg = sMsg & "^FO40,290^BY2,2,50^B3,,,N^FDQ" & sQty & "^FS"
        sMsg = sMsg & "^FO445,350^A0N,36,30^FDPACKER: " & sPacker & "^FS"
        sMsg = sMsg & "^FO445,420^A0N,36,30^FDQA:^FS"
        sMsg = sMsg & "^PQ1"
        sMsg = sMsg & "^XZ1"
    '******
    ' End of Label Info
    '******

    '******
    'Send message to TDS
    '******
        
        'Add by Tony Start on 20110117 for Req.JC201100003
        If InStr(1, sPrinterServer, "@") <> 0 Then
            sPrinterServer = VBA.Left(sPrinterServer, (InStr(1, sPrinterServer, "@")) - 1)
        End If
        'Add by Tony End on 20110117 for Req.JC201100003
        
        If Not moCwMbx.Send("MBX_" & sPrinterServer, sMsg) Then
            Call moCwMbx.GetError(lErrCode, sErrMsg)
            Call RaiseError(lErrCode, sErrMsg)
        End If
    '******

'----
' Done
'----
ExitHandler:
    ' NOTE 1:
    ' MUST CALL GetErrInfo() here first before another action
    Call GetErrInfo(msMODULE_ID, sProcID, typErrInfo, Erl)
    Call LogProcOut(msMODULE_ID, sProcID, typErrInfo, oLogCtrl)
    On Error Resume Next

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
        Call HandleError(False, typErrInfo, , oLogCtrl, True)
    End If
End Sub




'================================================================================
' Sub: Prt_FT_To_SubTR_Eng_Sample()
'--------------------------------------------------------------------------------
' Description:  <This subroutine will print out a label for FT-> SFG >
'--------------------------------------------------------------------------------
' Author:       <Alex Chan>, <CIT> 2002/02/07
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
Public Sub Prt_FT_To_SubTR_Eng_Sample(ByVal sLotNo As String, _
                                    ByVal sProdNo As String, _
                                    ByVal sQty As String, _
                                    ByVal sPacker As String, _
                                    ByVal bIsPartial As Boolean, _
                                    ByVal sPrinterServer As String, _
                                    ByVal moCwMbx As Object, _
                                    ByVal oLogCtrl As Object)

On Error GoTo ExitHandler:

Dim sMsg            As String
Dim sProcID         As String
Dim typErrInfo      As tErrInfo
Dim lErrCode        As Long
Dim sErrMsg         As String

'----
' Init
'----
    sProcID = "Prt_FT_To_SubTR_Eng_Sample"
    Call LogProcIn(msMODULE_ID, sProcID, oLogCtrl)

'----
' Condition Checking
'----

'----
' Action
'----
    '******
    'Print Label Info
    '******
        sMsg = sMsg & "^XA^LH00,00^LL550"
        sMsg = sMsg & "^FO10,05^GB780,475,3^FS"
        sMsg = sMsg & "^FO435,327^GB355,70,3^FS"
        sMsg = sMsg & "^FO435,397^GB0,80,3^FS"
        sMsg = sMsg & "^FO580,30^GB200,75,3^FS"
        sMsg = sMsg & "^FO640,55^A0N,35,50^CI0^FDT&R^FS"
        If bIsPartial Then
            sMsg = sMsg & "^FO580,102^GB200,75,3^FS"
            sMsg = sMsg & "^FO590,130^A0N,35,50^CI0^FDPARTIAL^FS"
        End If
        sMsg = sMsg & "^FO580,250^GB200,75,3^FS"
        sMsg = sMsg & "^FO640,275^A0N,35,50^CI0^FDENG^FS"
        sMsg = sMsg & "^FO40,58^A0N,33,28^FD(1T)LOT NO: " & sLotNo & "^FS"
        sMsg = sMsg & "^FO40,158^A0N,33,28^FD(1P)PRODUCT NO: " & sProdNo & "^FS"
        sMsg = sMsg & "^FO40,258^A0N,33,28^FD(Q)QUANTITY: " & sQty & "^FS"
        sMsg = sMsg & "^FO40,90^BY2,2,50^B3,,,N^FD1T" & sLotNo & "^FS"
        sMsg = sMsg & "^FO40,190^BY2,2,50^B3,,,N^FD1P" & sProdNo & "^FS"
        sMsg = sMsg & "^FO40,290^BY2,2,50^B3,,,N^FDQ" & sQty & "^FS"
        sMsg = sMsg & "^FO445,350^A0N,36,30^FDPACKER: " & sPacker & "^FS"
        sMsg = sMsg & "^FO445,420^A0N,36,30^FDQA:^FS"
        sMsg = sMsg & "^PQ1"
        sMsg = sMsg & "^XZ1"
    '******
    ' End of Label Info
    '******

    '******
    'Send message to TDS
    '******
        
        'Add by Tony Start on 20110117 for Req.JC201100003
        If InStr(1, sPrinterServer, "@") <> 0 Then
            sPrinterServer = VBA.Left(sPrinterServer, (InStr(1, sPrinterServer, "@")) - 1)
        End If
        'Add by Tony End on 20110117 for Req.JC201100003
        
        If Not moCwMbx.Send("MBX_" & sPrinterServer, sMsg) Then
            Call moCwMbx.GetError(lErrCode, sErrMsg)
            Call RaiseError(lErrCode, sErrMsg)
        End If
    '******

'----
' Done
'----
ExitHandler:
    ' NOTE 1:
    ' MUST CALL GetErrInfo() here first before another action
    Call GetErrInfo(msMODULE_ID, sProcID, typErrInfo, Erl)
    Call LogProcOut(msMODULE_ID, sProcID, typErrInfo, oLogCtrl)
    On Error Resume Next

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
        Call HandleError(False, typErrInfo, , oLogCtrl, True)
    End If
End Sub

'================================================================================
' Sub: Prt_WSMCD_TO_SFG()
'--------------------------------------------------------------------------------
' Description:  <This subroutine will print out a label for FT-> SFG >
'--------------------------------------------------------------------------------
' Author:       <Alex Chan>, <CIT> 2002/02/07
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
Public Sub Prt_WSMCD_TO_SFG(ByVal sParLotNo As String, _
                            ByVal sParIPN As String, _
                            ByVal sParWqty As String, _
                            ByVal sParCQTY As String, _
                            ByVal sUserName As String, _
                            ByVal sTimeStamp As String, _
                            ByVal sPrinterServer As String, _
                            ByVal moCwMbx As Object, _
                            ByVal oLogCtrl As Object, _
                            Optional oWSMCDInfo As Collection)


On Error GoTo ExitHandler:

Dim sMsg            As String
Dim sProcID         As String
Dim typErrInfo      As tErrInfo
Dim lErrCode        As Long
Dim sErrMsg         As String

Dim sLot(1 To 5)    As String
Dim sIPN(1 To 5)    As String
Dim sCQty(1 To 5)   As String
Dim sCode(1 To 5)   As String


Dim iIdx            As Integer
Dim arrItem()       As String
'----
' Init
'----
    sProcID = "Prt_WSMCD_TO_SFG"
    Call LogProcIn(msMODULE_ID, sProcID, oLogCtrl)

'----
' Condition Checking
'----
    
'----
' Action
'----
    On Error Resume Next
    
    If oWSMCDInfo Is Nothing Then
        sLot(1) = sParLotNo
        sIPN(1) = sParIPN
        sCQty(1) = sParCQTY
    ElseIf oWSMCDInfo.Count = 0 Then
        sLot(1) = sParLotNo
        sIPN(1) = sParIPN
        sCQty(1) = sParCQTY
    Else
        For iIdx = 1 To oWSMCDInfo.Count
            arrItem = Split(oWSMCDInfo.Item(oWSMCDInfo.Count - iIdx + 1), ";")
            sLot(iIdx) = arrItem(0)
            sIPN(iIdx) = arrItem(1)
            sCode(iIdx) = arrItem(2)
            sCQty(iIdx) = arrItem(3)
        Next iIdx
    End If
    '******
    'Print Label Info
    '******

        sMsg = sMsg & "^XA"
        sMsg = sMsg & "^LH0,0^FS"
        sMsg = sMsg & "^LL1441"
        sMsg = sMsg & "^MD0"
        sMsg = sMsg & "^MNY"
        sMsg = sMsg & "^FO30,50^A0N,30,30^CI0^FR^FDLOT : " & sLot(1) & "^FS"
        sMsg = sMsg & "^BY2, 2.0^FO25,85^B3N,N,60,N,Y^FR^FD" & sLot(1) & "^FS"
        sMsg = sMsg & "^FO30,160^A0N,30,32^CI0^FR^FDPROD : " & sIPN(1) & "^FS"
        sMsg = sMsg & "^BY1, 3.0^FO25,195^B3N,N,60,N,Y^FR^FD" & sIPN(1) & "^FS"
        sMsg = sMsg & "^FO360,300^A0R,40,32^CI0^FR^FDLOT^FS"
        sMsg = sMsg & "^FO360,500^A0R,40,32^CI0^FR^FDPROD^FS"
        sMsg = sMsg & "^FO360,900^A0R,40,32^CI0^FR^FDQTY^FS"
        sMsg = sMsg & "^FO360,1100^A0R,40,32^CI0^FR^FDCODE^FS"
        sMsg = sMsg & "^FO304,300^A0R,25,32^CI0^FR^FD" & sLot(1) & "^FS"
        sMsg = sMsg & "^FO304,500^A0R,25,32^CI0^FR^FD" & sIPN(1) & "^FS"
        sMsg = sMsg & "^FO304,900^A0R,25,32^CI0^FR^FD" & sCQty(1) & "^FS"
        sMsg = sMsg & "^FO304,1100^A0R,25,32^CI0^FR^FD" & sCode(1) & "^FS"
        sMsg = sMsg & "^FO268,300^A0R,25,32^CI0^FR^FD" & sLot(2) & "^FS"
        sMsg = sMsg & "^FO268,500^A0R,25,32^CI0^FR^FD" & sIPN(2) & "^FS"
        sMsg = sMsg & "^FO268,900^A0R,25,32^CI0^FR^FD" & sCQty(2) & "^FS"
        sMsg = sMsg & "^FO268,1100^A0R,25,32^CI0^FR^FD" & sCode(2) & "^FS"
        sMsg = sMsg & "^FO232,300^A0R,25,32^CI0^FR^FD" & sLot(3) & "^FS"
        sMsg = sMsg & "^FO232,500^A0R,25,32^CI0^FR^FD" & sIPN(3) & "^FS"
        sMsg = sMsg & "^FO232,900^A0R,25,32^CI0^FR^FD" & sCQty(3) & "^FS"
        sMsg = sMsg & "^FO232,1100^A0R,25,32^CI0^FR^FD" & sCode(3) & "^FS"
        sMsg = sMsg & "^FO196,300^A0R,25,32^CI0^FR^FD" & sLot(4) & "^FS"
        sMsg = sMsg & "^FO196,500^A0R,25,32^CI0^FR^FD" & sIPN(4) & "^FS"
        sMsg = sMsg & "^FO196,900^A0R,25,32^CI0^FR^FD" & sCQty(4) & "^FS"
        sMsg = sMsg & "^FO196,1100^A0R,25,32^CI0^FR^FD" & sCode(4) & "^FS"
        sMsg = sMsg & "^FO160,300^A0R,25,32^CI0^FR^FD" & sLot(5) & "^FS"
        sMsg = sMsg & "^FO160,500^A0R,25,32^CI0^FR^FD" & sIPN(5) & "^FS"
        sMsg = sMsg & "^FO160,900^A0R,25,32^CI0^FR^FD" & sCQty(5) & "^FS"
        sMsg = sMsg & "^FO160,1100^A0R,25,32^CI0^FR^FD" & sCode(5) & "^FS"
        sMsg = sMsg & "^FO105,300^A0R,30,32^CI0^FR^FDTTL QTY: " & sParWqty & " PCS,   " & sParCQTY & "EA^FS"
        sMsg = sMsg & "^FO105,720^A0R,30,32^CI0^FR^FDID No: " & sUserName & "^FS"
        sMsg = sMsg & "^FO105,950^A0R,30,32^CI0^FR^FDDG: 0^FS"
        sMsg = sMsg & "^FO50,300^A0R,30,32^CI0^FR^FDDATE: " & sTimeStamp & "^FS"
        sMsg = sMsg & "^FO50,650^A0R,30,32^CI0^FR^FDSerial No: 0206000115^FS"
        sMsg = sMsg & "^FO50,1000^A0R,30,32^CI0^FR^FDTQA:^FS"
        sMsg = sMsg & "^XZ"


    '******
    ' End of Label Info
    '******

    '******
    'Send message to TDS
    '******
        
        'Add by Tony Start on 20110117 for Req.JC201100003
        If InStr(1, sPrinterServer, "@") <> 0 Then
            sPrinterServer = VBA.Left(sPrinterServer, (InStr(1, sPrinterServer, "@")) - 1)
        End If
        'Add by Tony End on 20110117 for Req.JC201100003
        
        If Not moCwMbx.Send("MBX_" & sPrinterServer, sMsg) Then
            Call moCwMbx.GetError(lErrCode, sErrMsg)
            Call RaiseError(lErrCode, sErrMsg)
        End If
    '******

'----
' Done
'----
ExitHandler:
    ' NOTE 1:
    ' MUST CALL GetErrInfo() here first before another action
    Call GetErrInfo(msMODULE_ID, sProcID, typErrInfo, Erl)
    Call LogProcOut(msMODULE_ID, sProcID, typErrInfo, oLogCtrl)
    On Error Resume Next

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
        Call HandleError(False, typErrInfo, , oLogCtrl, True)
    End If
End Sub



'================================================================================
' Sub: Prt_FT_ETEST_MERGE()
'--------------------------------------------------------------------------------
' Description:  <This subroutine will print out a label for FT-> SFG >
'--------------------------------------------------------------------------------
' Author:       Vencent Wei, CIT 2004/09/23
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
'================================================================================
Public Sub Prt_FT_ETEST_MERGE(ByVal sLotsInfo As String, _
                            ByVal sPrinterServer As String, _
                            ByVal moCwMbx As Object, _
                            ByVal oLogCtrl As Object)


On Error GoTo ExitHandler:

Dim sMsg            As String
Dim sProcID         As String
Dim typErrInfo      As tErrInfo
Dim lErrCode        As Long
Dim sErrMsg         As String

Dim iIdx            As Integer
Dim arrLotInfo      As Variant
'----
' Init
'----
    sProcID = "Prt_FT_ETEST_MERGE"
    Call LogProcIn(msMODULE_ID, sProcID, oLogCtrl)

'----
' Condition Checking
'----
    
'----
' Action
'----
   
    '******
    'Print Label Info
    '******
    arrLotInfo = Split(sLotsInfo, ",")
    sMsg = ""
    
    '每六批印一張
    For iIdx = LBound(arrLotInfo) To UBound(arrLotInfo)
        
'        If (iIdx - LBound(arrLotInfo)) Mod 2 = 0 Then
'            sMsg = sMsg & "^FO15,36^A0N,40,30^CI13^FR^FDLot No: " & arrLotInfo(iIdx) & "^FS"
'        Else
'            sMsg = sMsg & "^FO500,36^A0N,40,30^CI13^FR^FDQTY: " & arrLotInfo(iIdx) & "^FS"
'        End If
        '---
        'Modify by Janus, 20050831, request no : M200507046
        Select Case ((iIdx - LBound(arrLotInfo)) Mod 12)
        Case 0
            sMsg = "^XA"
            sMsg = sMsg & "^PRC"
            sMsg = sMsg & "^LH0,0^FS"
            sMsg = sMsg & "^LL280"
            sMsg = sMsg & "^MD0"
            sMsg = sMsg & "^LH0,0^FS"
            sMsg = sMsg & "^FO15,15^A0N,40,30^CI13^FR^FDLot No: " & arrLotInfo(iIdx) & "^FS"
            'sMsg = sMsg & "^FO15,36^A0N,40,30^CI13^FR^FDLot No: " & arrLotInfo(iIdx) & "^FS"
        Case 1
            'sMsg = sMsg & "^FO500,36^A0N,40,30^CI13^FR^FDQTY: " & arrLotInfo(iIdx) & "^FS"
            sMsg = sMsg & "^FO300,15^A0N,40,30^CI13^FR^FDQTY: " & arrLotInfo(iIdx) & "^FS"
        Case 2
            'sMsg = sMsg & "^FO15,100^A0N,40,30^CI13^FR^FDLot No: " & arrLotInfo(iIdx) & "^FS"
            sMsg = sMsg & "^FO15,60^A0N,40,30^CI13^FR^FDLot No: " & arrLotInfo(iIdx) & "^FS"
        Case 3
            'sMsg = sMsg & "^FO500,100^A0N,40,30^CI13^FR^FDQTY: " & arrLotInfo(iIdx) & "^FS"
            sMsg = sMsg & "^FO300,60^A0N,40,30^CI13^FR^FDQTY: " & arrLotInfo(iIdx) & "^FS"
        Case 4
            'sMsg = sMsg & "^FO15,164^A0N,40,30^CI13^FR^FDLot No: " & arrLotInfo(iIdx) & "^FS"
            sMsg = sMsg & "^FO15,105^A0N,40,30^CI13^FR^FDLot No: " & arrLotInfo(iIdx) & "^FS"
        Case 5
            'sMsg = sMsg & "^FO500,164^A0N,40,30^CI13^FR^FDQTY: " & arrLotInfo(iIdx) & "^FS"
            sMsg = sMsg & "^FO300,105^A0N,40,30^CI13^FR^FDQTY: " & arrLotInfo(iIdx) & "^FS"
        Case 6
            'sMsg = sMsg & "^FO15,228^A0N,40,30^CI13^FR^FDLot No: " & arrLotInfo(iIdx) & "^FS"
            sMsg = sMsg & "^FO15,150^A0N,40,30^CI13^FR^FDLot No: " & arrLotInfo(iIdx) & "^FS"
        Case 7
            'sMsg = sMsg & "^FO500,228^A0N,40,30^CI13^FR^FDQTY: " & arrLotInfo(iIdx) & "^FS"
            sMsg = sMsg & "^FO300,150^A0N,40,30^CI13^FR^FDQTY: " & arrLotInfo(iIdx) & "^FS"
        Case 8
            'sMsg = sMsg & "^FO15,292^A0N,40,30^CI13^FR^FDLot No: " & arrLotInfo(iIdx) & "^FS"
            sMsg = sMsg & "^FO15,195^A0N,40,30^CI13^FR^FDLot No: " & arrLotInfo(iIdx) & "^FS"
        Case 9
            'sMsg = sMsg & "^FO500,292^A0N,40,30^CI13^FR^FDQTY: " & arrLotInfo(iIdx) & "^FS"
            sMsg = sMsg & "^FO300,195^A0N,40,30^CI13^FR^FDQTY: " & arrLotInfo(iIdx) & "^FS"
        Case 10
            'sMsg = sMsg & "^FO15,356^A0N,40,30^CI13^FR^FDLot No: " & arrLotInfo(iIdx) & "^FS"
            sMsg = sMsg & "^FO15,235^A0N,40,30^CI13^FR^FDLot No: " & arrLotInfo(iIdx) & "^FS"
        Case 11
            'sMsg = sMsg & "^FO500,356^A0N,40,30^CI13^FR^FDQTY: " & arrLotInfo(iIdx) & "^FS"
            sMsg = sMsg & "^FO300,235^A0N,40,30^CI13^FR^FDQTY: " & arrLotInfo(iIdx) & "^FS"
            sMsg = sMsg & "^PQ1,1,1,Y"
            sMsg = sMsg & "^XZ"
            'End Modify by Janus, 20050831
            '---
            '******
            'Send message to TDS
            '******
            
            'Add by Tony Start on 20110117 for Req.JC201100003
            If InStr(1, sPrinterServer, "@") <> 0 Then
                sPrinterServer = VBA.Left(sPrinterServer, (InStr(1, sPrinterServer, "@")) - 1)
            End If
            'Add by Tony End on 20110117 for Req.JC201100003
            
            If Not moCwMbx.Send("MBX_" & sPrinterServer, sMsg) Then
                Call moCwMbx.GetError(lErrCode, sErrMsg)
                Call RaiseError(lErrCode, sErrMsg)
            End If
        End Select
        
        If iIdx = UBound(arrLotInfo) And ((iIdx - LBound(arrLotInfo)) Mod 12) <> 11 Then
            sMsg = sMsg & "^PQ1,1,1,Y"
            sMsg = sMsg & "^XZ"

            'Add by Tony Start on 20110117 for Req.JC201100003
            If InStr(1, sPrinterServer, "@") <> 0 Then
                sPrinterServer = VBA.Left(sPrinterServer, (InStr(1, sPrinterServer, "@")) - 1)
            End If
            'Add by Tony End on 20110117 for Req.JC201100003
            
            If Not moCwMbx.Send("MBX_" & sPrinterServer, sMsg) Then
                Call moCwMbx.GetError(lErrCode, sErrMsg)
                Call RaiseError(lErrCode, sErrMsg)
            End If
        End If
    Next iIdx


    '******
    
    '******
    ' End of Label Info
    '******


'----
' Done
'----
ExitHandler:
    ' NOTE 1:
    ' MUST CALL GetErrInfo() here first before another action
    Call GetErrInfo(msMODULE_ID, sProcID, typErrInfo, Erl)
    Call LogProcOut(msMODULE_ID, sProcID, typErrInfo, oLogCtrl)
    On Error Resume Next

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
        Call HandleError(False, typErrInfo, , oLogCtrl, True)
    End If
End Sub
'================================================================================
' Sub: Prt_WS_WS_SMALL_LABEL()
'--------------------------------------------------------------------------------
' Description:  <This subroutine will print out a label for FT-> SFG >
'--------------------------------------------------------------------------------
' Author:       Jeff Weng, MXIC 2006/11/30
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
Public Sub Prt_WS_WS_DGRADE_SUMMARY(ByVal sLotNo As String, _
                            ByVal sProdNo As String, _
                            ByVal sWQty As String, _
                            ByVal sCQty As String, _
                            ByVal sUserName As String, _
                            ByVal sTimeStamp As String, _
                            ByVal sPrinterServer As String, _
                            ByVal moCwMbx As Object, _
                            ByVal oLogCtrl As Object, _
                            ByVal moProRawSql As Object)


On Error GoTo ExitHandler:

Dim sMsg            As String
Dim sProcID         As String
Dim typErrInfo      As tErrInfo
Dim lErrCode        As Long
Dim sErrMsg         As String

'----
' Init
'----
    Dim sSQL            As String
    Dim colRS           As Collection
    Dim iIndex          As Integer
    Dim sItem(1 To 25)  As String
    Dim oItem           As Object
    Dim arrWaferIdStr     As String
    Dim sId             As String
    Dim sOriLotID       As String 'Added by Jack on 2016/09/09 for JC201600178
    'add by Ernest on 2021/02/02 for for CP 測前併批專案------start
    Dim sSlotNo         As String
    'add by Ernest on 2021/02/02 for for CP 測前併批專案-----end
        
    sProcID = "Prt_WS_WS_DGRADE_SUMMARY"
    Call LogProcIn(msMODULE_ID, sProcID, oLogCtrl)
    
    sOriLotID = sLotNo 'Added by Jack on 2016/09/09 for JC201600178 (保留列印原母批批號).
    
PRINT_JOB:  'Added by Jack on 2016/09/01 for JC201600178
    'modify  by Ernest on 2021/02/02 for for CP 測前併批專案
    'sSQL = "select childlotid,waferid from tbl_inkless_merge_list where parentlotid ='" & sLotNo & "'"
    sSQL = "select childlotid,waferid,slotno from tbl_inkless_merge_list where parentlotid ='" & sLotNo & "' and deleteflag='N' order by slotno"
    Set colRS = moProRawSql.QueryDatabase(sSQL)
      
    If colRS.Count > 0 Then
        
        sSlotNo = colRS.Item(1).Item("slotno")
        If sSlotNo = "" Then
            
            For iIndex = 1 To 25
                sItem(iIndex) = ""
            Next
        
            iIndex = 0
            
            For Each oItem In colRS
                arrWaferIdStr = oItem.Item("waferid")
                Do
                    iIndex = iIndex + 1
                    If InStr(arrWaferIdStr, ";") > 0 Then
                        sId = Left(arrWaferIdStr, InStr(arrWaferIdStr, ";") - 1)
                        sItem(iIndex) = Left(oItem.Item("childlotid"), 6) & "-" & sId
                        arrWaferIdStr = Mid(arrWaferIdStr, InStr(arrWaferIdStr, ";") + 1, Len(arrWaferIdStr) - (InStr(arrWaferIdStr, ";") - 1))
                    Else
                        sId = arrWaferIdStr
                        sItem(iIndex) = Left(oItem.Item("childlotid"), 6) & "-" & sId
                        arrWaferIdStr = ""
                    End If
                Loop While Len(arrWaferIdStr) > 0 Or iIndex > 25
                
            Next
        Else
            For iIndex = 1 To 25
                sItem(iIndex) = "______-__"
            Next
            For iIndex = 1 To colRS.Count
                sSlotNo = colRS.Item(iIndex).Item("slotno")
                If sSlotNo <> "" Then
                    sItem(Val(sSlotNo)) = Left(colRS.Item(iIndex).Item("childlotid"), 6) & "-" & colRS.Item(iIndex).Item("waferid")
                Else
                    sItem(iIndex) = Left(colRS.Item(iIndex).Item("childlotid"), 6) & "-" & colRS.Item(iIndex).Item("waferid")
                End If
            Next
        End If
        
    Else
        'Added by Jack on 2016/09/01 for JC201600178 <Start>
        '如果串不到併批資料且lot no.第9碼&10碼為"01~99"時, 提示訊息,
        '以輸入的"原始批號Lotid " 串tbl_inkless_merge_list.parentlotid 資料.
        If Len(sLotNo) = 10 Then 'LotID長度為10, 第9碼&10碼為"01~99".
            If Right(sLotNo, 1) <= "9" And Right(sLotNo, 1) >= "0" _
                And Mid(sLotNo, 9, 1) <= "9" And Mid(sLotNo, 9, 1) >= "0" Then
                
                sLotNo = InputBox("無母批merge 資訊, 請輸入原始批號 : " & vbNewLine & vbNewLine & "No parent lot merge information" & vbNewLine & "Please Input Original LotNo : ", "Error")
                If Trim(sLotNo) <> "" Then
                    GoTo PRINT_JOB '以原始批號再列印一次.
                End If
            End If
        End If
        'Added by Jack on 2016/09/01 for JC201600178 <End>
        GoTo ExitHandler
    End If
    sSQL = "": Set colRS = Nothing

    sLotNo = sOriLotID 'Added by Jack on 2016/09/09 for JC201600178 (還原列印原母批批號).

'----
' Condition Checking
'----

'----
' Action
'----
    On Error Resume Next
    '******
    'Print Label Info
    '******
    sMsg = ""
'mark by Nelson on 2007/09/14 for SI Request
'            sMsg = sMsg & "^XA"
'            sMsg = sMsg & "^LH0,0^FS"
'            sMsg = sMsg & "^LL1441"
'            sMsg = sMsg & "^MD0"
'            sMsg = sMsg & "^MNY"
'            sMsg = sMsg & "^FO360,365^A0R,40,32^CI0^FR^FDLOT : " & sLotNo & "^FS"
'            sMsg = sMsg & "^FO30,50^A0N,30,30^CI0^FR^FDLOT : " & sLotNo & "^FS"
'            sMsg = sMsg & "^BY2, 2.0^FO25,85^B3N,N,60,N,Y^FR^FD" & sLotNo & "^FS"
'            sMsg = sMsg & "^FO360,770^A0R,40,32^CI0^FR^FDPROD : " & sProdNo & "^FS"
'            sMsg = sMsg & "^FO30,160^A0N,30,32^CI0^FR^FDPROD : " & sProdNo & "^FS"
'            sMsg = sMsg & "^BY1, 3.0^FO25,195^B3N,N,60,N,Y^FR^FD" & sProdNo & "^FS"
'
'    For iIndex = 1 To 25
'        Select Case iIndex
'        Case 1
'            sMsg = sMsg & "^FO304,300^A0R,25,32^CI0^FR^FD " & sItem(iIndex) & "^FS" '01
'        Case 6
'            sMsg = sMsg & "^FO304,500^A0R,25,32^CI0^FR^FD " & sItem(iIndex) & "^FS" '06
'        Case 11
'            sMsg = sMsg & "^FO304,700^A0R,25,32^CI0^FR^FD " & sItem(iIndex) & "^FS" '11
'        Case 16
'            sMsg = sMsg & "^FO304,900^A0R,25,32^CI0^FR^FD " & sItem(iIndex) & "^FS" '16
'        Case 21
'            sMsg = sMsg & "^FO304,1100^A0R,25,32^CI0^FR^FD " & sItem(iIndex) & "^FS" '21
'        Case 2
'            sMsg = sMsg & "^FO268,300^A0R,25,32^CI0^FR^FD " & sItem(iIndex) & "^FS" '02
'        Case 7
'            sMsg = sMsg & "^FO268,500^A0R,25,32^CI0^FR^FD " & sItem(iIndex) & "^FS" '07
'        Case 12
'            sMsg = sMsg & "^FO268,700^A0R,25,32^CI0^FR^FD " & sItem(iIndex) & "^FS" '12
'        Case 17
'            sMsg = sMsg & "^FO268,900^A0R,25,32^CI0^FR^FD " & sItem(iIndex) & "^FS" '17
'        Case 22
'            sMsg = sMsg & "^FO268,1100^A0R,25,32^CI0^FR^FD " & sItem(iIndex) & "^FS" '22
'        Case 3
'            sMsg = sMsg & "^FO232,300^A0R,25,32^CI0^FR^FD " & sItem(iIndex) & "^FS" '03
'        Case 8
'            sMsg = sMsg & "^FO232,500^A0R,25,32^CI0^FR^FD " & sItem(iIndex) & "^FS" '08
'        Case 13
'            sMsg = sMsg & "^FO232,700^A0R,25,32^CI0^FR^FD " & sItem(iIndex) & "^FS" '13
'        Case 18
'            sMsg = sMsg & "^FO232,900^A0R,25,32^CI0^FR^FD " & sItem(iIndex) & "^FS" '18
'        Case 23
'            sMsg = sMsg & "^FO232,1100^A0R,25,32^CI0^FR^FD " & sItem(iIndex) & "^FS" '23
'        Case 4
'            sMsg = sMsg & "^FO196,300^A0R,25,32^CI0^FR^FD " & sItem(iIndex) & "^FS" '04
'        Case 9
'            sMsg = sMsg & "^FO196,500^A0R,25,32^CI0^FR^FD " & sItem(iIndex) & "^FS" '09
'        Case 14
'            sMsg = sMsg & "^FO196,700^A0R,25,32^CI0^FR^FD " & sItem(iIndex) & "^FS" '14
'        Case 19
'            sMsg = sMsg & "^FO196,900^A0R,25,32^CI0^FR^FD " & sItem(iIndex) & "^FS" '19
'        Case 24
'            sMsg = sMsg & "^FO196,1100^A0R,25,32^CI0^FR^FD " & sItem(iIndex) & "^FS" '24
'        Case 5
'            sMsg = sMsg & "^FO160,300^A0R,25,32^CI0^FR^FD " & sItem(iIndex) & "^FS" '05
'        Case 10
'            sMsg = sMsg & "^FO160,500^A0R,25,32^CI0^FR^FD " & sItem(iIndex) & "^FS" '10
'        Case 15
'            sMsg = sMsg & "^FO160,700^A0R,25,32^CI0^FR^FD " & sItem(iIndex) & "^FS" '15
'        Case 20
'            sMsg = sMsg & "^FO160,900^A0R,25,32^CI0^FR^FD " & sItem(iIndex) & "^FS" '20
'        Case 25
'            sMsg = sMsg & "^FO160,1100^A0R,25,32^CI0^FR^FD " & sItem(iIndex) & "^FS" '25
'        End Select
'    Next iIndex
'            sMsg = sMsg & "^FO105,300^A0R,30,32^CI0^FR^FDTTL QTY: " & sWqty & " PCS,   " & sCqty & "EA^FS"
'            sMsg = sMsg & "^FO105,720^A0R,30,32^CI0^FR^FDID No: " & sUserName & "^FS"
'            sMsg = sMsg & "^FO105,950^A0R,30,32^CI0^FR^FDDG: 0^FS"
'            sMsg = sMsg & "^FO50,300^A0R,30,32^CI0^FR^FDDATE: " & sTimeStamp & "^FS"
'            sMsg = sMsg & "^FO50,650^A0R,30,32^CI0^FR^FDSerial No: 0206000115^FS"
'            sMsg = sMsg & "^FO50,1000^A0R,30,32^CI0^FR^FDTQA:^FS"
'            sMsg = sMsg & "^XZ"
            sMsg = sMsg & "^XA"
            sMsg = sMsg & "^LH0,0^FS"
            sMsg = sMsg & "^LL440"
            sMsg = sMsg & "^MD0"
            sMsg = sMsg & "^MNY"
            sMsg = sMsg & "^FO368,36^A0N,40.32^CI0^FR^FDLOT : " & sLotNo & "^FS"
            
'            sMsg = sMsg & "^FO48,50^A0B,30,30^CI0^FR^FDLOT : " & sLotNo & "^FS"
            sMsg = sMsg & "^FO48,10^A0B,30,30^CI0^FR^FDLOT : " & sLotNo & "^FS"
            
            sMsg = sMsg & "^BY2, 2.0^FO80,50^B3B,N,60,N,Y^FR^FD" & sLotNo & "^FS"
            sMsg = sMsg & "^FO775,36^A0N,40,32^CI0^FR^FDPROD : " & sProdNo & "^FS"
            
'            sMsg = sMsg & "^FO160,50^A0B,30,32^CI0^FR^FDPROD : " & sProdNo & "^FS"
            sMsg = sMsg & "^FO160,10^A0B,30,32^CI0^FR^FDPROD : " & sProdNo & "^FS"
            
            sMsg = sMsg & "^BY1, 3.0^FO192,50^B3B,N,60,N,Y^FR^FD" & sProdNo & "^FS"
    
    For iIndex = 1 To 25
        Select Case iIndex
        Case 1
            sMsg = sMsg & "^FO300,108^A0N,25,32^CI0^FR^FD " & sItem(iIndex) & "^FS" '01
        Case 6
            sMsg = sMsg & "^FO480,108^A0N,25,32^CI0^FR^FD " & sItem(iIndex) & "^FS" '06
        Case 11
            sMsg = sMsg & "^FO680,108^A0N,25,32^CI0^FR^FD " & sItem(iIndex) & "^FS" '11
        Case 16
            sMsg = sMsg & "^FO880,108^A0N,25,32^CI0^FR^FD " & sItem(iIndex) & "^FS" '16
        Case 21
            sMsg = sMsg & "^FO1080,108^A0N,25,32^CI0^FR^FD " & sItem(iIndex) & "^FS" '21
        Case 2
            sMsg = sMsg & "^FO300,144^A0N,25,32^CI0^FR^FD " & sItem(iIndex) & "^FS" '02
        Case 7
            sMsg = sMsg & "^FO480,144^A0N,25,32^CI0^FR^FD " & sItem(iIndex) & "^FS" '07
        Case 12
            sMsg = sMsg & "^FO680,144^A0N,25,32^CI0^FR^FD " & sItem(iIndex) & "^FS" '12
        Case 17
            sMsg = sMsg & "^FO880,144^A0N,25,32^CI0^FR^FD " & sItem(iIndex) & "^FS" '17
        Case 22
            sMsg = sMsg & "^FO1080,144^A0N,25,32^CI0^FR^FD " & sItem(iIndex) & "^FS" '22
        Case 3
            sMsg = sMsg & "^FO300,180^A0N,25,32^CI0^FR^FD " & sItem(iIndex) & "^FS" '03
        Case 8
            sMsg = sMsg & "^FO480,180^A0N,25,32^CI0^FR^FD " & sItem(iIndex) & "^FS" '08
        Case 13
            sMsg = sMsg & "^FO680,180^A0N,25,32^CI0^FR^FD " & sItem(iIndex) & "^FS" '13
        Case 18
            sMsg = sMsg & "^FO880,180^A0N,25,32^CI0^FR^FD " & sItem(iIndex) & "^FS" '18
        Case 23
            sMsg = sMsg & "^FO1080,180^A0N,25,32^CI0^FR^FD " & sItem(iIndex) & "^FS" '23
        Case 4
            sMsg = sMsg & "^FO300,216^A0N,25,32^CI0^FR^FD " & sItem(iIndex) & "^FS" '04
        Case 9
            sMsg = sMsg & "^FO480,216^A0N,25,32^CI0^FR^FD " & sItem(iIndex) & "^FS" '09
        Case 14
            sMsg = sMsg & "^FO680,216^A0N,25,32^CI0^FR^FD " & sItem(iIndex) & "^FS" '14
        Case 19
            sMsg = sMsg & "^FO880,216^A0N,25,32^CI0^FR^FD " & sItem(iIndex) & "^FS" '19
        Case 24
            sMsg = sMsg & "^FO1080,216^A0N,25,32^CI0^FR^FD " & sItem(iIndex) & "^FS" '24
        Case 5
            sMsg = sMsg & "^FO300,252^A0N,25,32^CI0^FR^FD " & sItem(iIndex) & "^FS" '05
        Case 10
            sMsg = sMsg & "^FO480,252^A0N,25,32^CI0^FR^FD " & sItem(iIndex) & "^FS" '10
        Case 15
            sMsg = sMsg & "^FO680,252^A0N,25,32^CI0^FR^FD " & sItem(iIndex) & "^FS" '15
        Case 20
            sMsg = sMsg & "^FO880,252^A0N,25,32^CI0^FR^FD " & sItem(iIndex) & "^FS" '20
        Case 25
            sMsg = sMsg & "^FO1080,252^A0N,25,32^CI0^FR^FD " & sItem(iIndex) & "^FS" '25
        End Select
    Next iIndex
    
'Modify by Sam Start on 20080829 for ReqNo:M200808056
'            sMsg = sMsg & "^FO300,304^A0N,30,32^CI0^FR^FDTTL QTY: " & sWqty & " PCS,   " & sCqty & "EA^FS"
'            sMsg = sMsg & "^FO723,304^A0N,30,32^CI0^FR^FDID No: " & sUserName & "^FS"
'            sMsg = sMsg & "^FO951,304^A0N,30,32^CI0^FR^FDDG: 0^FS"
'            sMsg = sMsg & "^FO300,356^A0N,30,32^CI0^FR^FDDATE: " & sTimeStamp & "^FS"
'            sMsg = sMsg & "^FO651,356^A0N,30,32^CI0^FR^FDSerial No: 0206000115^FS"
'            sMsg = sMsg & "^FO999,356^A0N,30,32^CI0^FR^FDTQA:^FS"
'            sMsg = sMsg & "^XZ"
        sMsg = sMsg & "^FO250,318^A0N,30,32^CI0^FR^FDTTL WQTY: " & sWQty & " PCS ^FS"
        
'Marked by Jack 2009/07/15 JC900084
'        sMsg = sMsg & "^BY2,2.0^FO550,318^B3N,N,40,N,Y^FR^FD" & sWqty & "^FS"
'        sMsg = sMsg & "^FO250,370^A0N,30,32^CI0^FR^FDTTL CQTY: " & sCqty & "EA ^FS"
'        sMsg = sMsg & "^BY2,2.0^FO550,370^B3N,N,40,N,Y^FR^FD" & sCqty & "^FS"
'        sMsg = sMsg & "^FO750,318^A0N,30,32^CI0^FR^FDID No: " & sUserName & "^FS"
'        sMsg = sMsg & "^FO1000,318^A0N,30,32^CI0^FR^FDTQA:^FS"
'        sMsg = sMsg & "^FO830,370^A0N,30,32^CI0^FR^FDDATE: " & sTimeStamp & "^FS"
        
        sMsg = sMsg & "^BY2,2.0^FO540,318^B3N,N,40,N,Y^FR^FD" & sWQty & "^FS"
        sMsg = sMsg & "^FO250,370^A0N,30,32^CI0^FR^FDTTL CQTY: " & sCQty & "EA ^FS"
        sMsg = sMsg & "^BY2,2.0^FO540,370^B3N,N,40,N,Y^FR^FD" & sCQty & "^FS"
        sMsg = sMsg & "^FO690,318^A0N,30,32^CI0^FR^FDID No: " & sUserName & "^FS"
        sMsg = sMsg & "^FO980,318^A0N,30,32^CI0^FR^FDTQA:^FS"
        sMsg = sMsg & "^FO790,370^A0N,30,32^CI0^FR^FDDATE: " & sTimeStamp & "^FS"
        
'Added by Jack 2009/07/15 JC900084
        sMsg = sMsg & "^FO1070,284^GB190,140,4^FS"
        
        sMsg = sMsg & "^XZ"
'Modify by Sam end on 20080829 for ReqNo:M200808056

'mark by Nelson end on 2007/09/14 for SI Request
        

    '******
    ' End of Label Info
    '******

    '******
    'Send message to TDS
    '******
        
        'Add by Tony Start on 20110117 for Req.JC201100003
        If InStr(1, sPrinterServer, "@") <> 0 Then
            sPrinterServer = VBA.Left(sPrinterServer, (InStr(1, sPrinterServer, "@")) - 1)
        End If
        'Add by Tony End on 20110117 for Req.JC201100003
        
        If Not moCwMbx.Send("MBX_" & sPrinterServer, sMsg) Then
            Call moCwMbx.GetError(lErrCode, sErrMsg)
            Call RaiseError(lErrCode, sErrMsg)
        End If
    '******

'----
' Done
'----
ExitHandler:
    ' NOTE 1:
    ' MUST CALL GetErrInfo() here first before another action
    Call GetErrInfo(msMODULE_ID, sProcID, typErrInfo, Erl)
    Call LogProcOut(msMODULE_ID, sProcID, typErrInfo, oLogCtrl)
    On Error Resume Next

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
        Call HandleError(False, typErrInfo, , oLogCtrl, True)
    End If
End Sub


'================================================================================
' Sub: Prt_WSMULTILOT_TO_SFG()
'--------------------------------------------------------------------------------
' Description:  <This subroutine will print out a label for FT-> SFG >
'--------------------------------------------------------------------------------
' Author:       Vencent Wei, CIT 2007/07/31
'================================================================================
Public Sub Prt_WSMULTILOT_TO_SFG(ByVal sUserName As String, _
                                ByVal sTimeStamp As String, _
                                ByVal sParWqty As String, _
                                ByVal sParCQTY As String, _
                                ByVal sPrinterServer As String, _
                                ByVal moCwMbx As Object, _
                                ByVal oLogCtrl As Object, _
                                colLotInfo As Collection)


On Error GoTo ExitHandler:

Dim sMsg            As String
Dim sProcID         As String
Dim typErrInfo      As tErrInfo
Dim lErrCode        As Long
Dim sErrMsg         As String

Dim iIdx            As Integer
Dim arrLotInfo      As Variant

'----
' Init
'----
    sProcID = "Prt_WSMULTILOT_TO_SFG"
    Call LogProcIn(msMODULE_ID, sProcID, oLogCtrl)
'----
' Condition Checking
'----
    
'----
' Action
'----
   
    '******
    'Print Label Info
    '******
    sMsg = ""
    
'mark by Nelson start on 2007/09/14 for SI Request
'    sMsg = sMsg & "^XA"
'    sMsg = sMsg & "^LH0,0^FS"
'    sMsg = sMsg & "^LL1441"
'    sMsg = sMsg & "^MD0"
'    sMsg = sMsg & "^MNY"
'
'    sMsg = sMsg & "^FO360,120^A0R,40,32^CI0^FR^FDLOT^FS"
'    sMsg = sMsg & "^FO360,300^A0R,40,32^CI0^FR^FDPROD^FS"
'    sMsg = sMsg & "^FO360, 680^A0R,40,32^CI0^FR^FDQTY^FS"
'    sMsg = sMsg & "^FO360,810^A0R,40,32^CI0^FR^FDCODE^FS"
'
'
'    For iIdx = colLotInfo.Count To 1 Step -1
'        arrLotInfo = Split(colLotInfo(iIdx), ";")
'        'vChildLotId & ";" & vIPN & ";" & vCode & ";" & vCQty
'
'        sMsg = sMsg & "^FO" & CStr(304 - 36 * (iIdx - 1)) & ", 60^A0R,25,32^CI0^FR^FD" & arrLotInfo(0) & "^FS"
'        sMsg = sMsg & "^FO" & CStr(304 - 36 * (iIdx - 1)) & ", 300^A0R,25,32^CI0^FR^FD" & arrLotInfo(1) & "^FS"
'        sMsg = sMsg & "^FO" & CStr(304 - 36 * (iIdx - 1)) & ", 680^A0R,25,32^CI0^FR^FD" & arrLotInfo(3) & "^FS"
'        sMsg = sMsg & "^FO" & CStr(304 - 36 * (iIdx - 1)) & ", 810^A0R,25,32^CI0^FR^FD" & arrLotInfo(2) & "^FS"
'        sMsg = sMsg & "^BY2, 2.0^FO" & CStr(295 - 50 * (iIdx - 1)) & ",910^B3R,N,30,N,Y^FR^FD" & arrLotInfo(0) & "^FS"
'
'    Next iIdx
'
'    sMsg = sMsg & "^FO105,100^A0R,30,32^CI0^FR^FDTTL QTY:  " & sParWqty & " PCS,    " & sParCQTY & " EA^FS"
'    sMsg = sMsg & "^FO105,520^A0R,30,32^CI0^FR^FDID No:  " & sUserName & " ^FS"
'    sMsg = sMsg & "^FO105,770^A0R,30,32^CI0^FR^FDDG: 0^FS"
'    sMsg = sMsg & "^FO50,100^A0R,30,32^CI0^FR^FDDATE: " & sTimeStamp & " ^FS"
'    sMsg = sMsg & "^FO50,470^A0R,30,32^CI0^FR^FDSerial No: 0206000115^FS"
'    sMsg = sMsg & "^FO50,800^A0R,30,32^CI0^FR^FDTQA:^FS"
'    sMsg = sMsg & "^XZ"
    sMsg = sMsg & "^XA"
    sMsg = sMsg & "^LH0,0^FS"
    sMsg = sMsg & "^LL440"
    sMsg = sMsg & "^MD0"
    sMsg = sMsg & "^MNY"

    sMsg = sMsg & "^FO116,36^A0N,40,32^CI0^FR^FDLOT^FS"
    sMsg = sMsg & "^FO300,36^A0N,40,32^CI0^FR^FDPROD^FS"
    sMsg = sMsg & "^FO655,36^A0N,40,32^CI0^FR^FDQTY^FS"
    sMsg = sMsg & "^FO783,36^A0N,40,32^CI0^FR^FDCODE^FS"


    For iIdx = 1 To colLotInfo.Count
        arrLotInfo = Split(colLotInfo(iIdx), ";")
        'vChildLotId & ";" & vIPN & ";" & vCode & ";" & vCQty

        sMsg = sMsg & "^FO56, " & CStr(108 + 36 * (iIdx - 1)) & "^A0N,25,32^CI0^FR^FD" & arrLotInfo(0) & "^FS"
        sMsg = sMsg & "^FO296, " & CStr(108 + 36 * (iIdx - 1)) & "^A0N,25,32^CI0^FR^FD" & arrLotInfo(1) & "^FS"
        sMsg = sMsg & "^FO655, " & CStr(108 + 36 * (iIdx - 1)) & "^A0N,25,32^CI0^FR^FD" & arrLotInfo(3) & "^FS"
        sMsg = sMsg & "^FO783, " & CStr(108 + 36 * (iIdx - 1)) & "^A0N,25,32^CI0^FR^FD" & arrLotInfo(2) & "^FS"
        sMsg = sMsg & "^BY2, 2.0^FO875, " & CStr(108 + 36 * (iIdx - 1)) & "^B3N,N,30,N,Y^FR^FD" & arrLotInfo(0) & "^FS"
        
    Next iIdx
    
    sMsg = sMsg & "^FO96,304^A0N,30,32^CI0^FR^FDTTL QTY:  " & sParWqty & " PCS,    " & sParCQTY & " EA^FS"
    sMsg = sMsg & "^FO520,304^A0N,30,32^CI0^FR^FDID No:  " & sUserName & " ^FS"
    sMsg = sMsg & "^FO747,304^A0N,30,32^CI0^FR^FDDG: 0^FS"
    sMsg = sMsg & "^FO96,356^A0N,30,32^CI0^FR^FDDATE: " & sTimeStamp & " ^FS"
    sMsg = sMsg & "^FO480,356^A0N,30,32^CI0^FR^FDSerial No: 0206000115^FS"
    
    'Modified by Jack 2009/07/15 JC900084
    'sMsg = sMsg & "^FO795,356^A0N,30,32^CI0^FR^FDTQA:^FS"
    sMsg = sMsg & "^FO850,356^A0N,30,32^CI0^FR^FDTQA:^FS"
    
    'Added by Jack 2009/07/15 JC900084
    sMsg = sMsg & "^FO930,284^GB190,140,4^FS"
    
    sMsg = sMsg & "^XZ"
'add by Nelson end on 2007/09/14 for SI Request
    '******
    
    '******
    ' End of Label Info
    '******
    '******
    'Send message to prtserver
    '******
        
        'Add by Tony Start on 20110117 for Req.JC201100003
        If InStr(1, sPrinterServer, "@") <> 0 Then
            sPrinterServer = VBA.Left(sPrinterServer, (InStr(1, sPrinterServer, "@")) - 1)
        End If
        'Add by Tony End on 20110117 for Req.JC201100003
        
        If Not moCwMbx.Send("MBX_" & sPrinterServer, sMsg) Then
            Call moCwMbx.GetError(lErrCode, sErrMsg)
            Call RaiseError(lErrCode, sErrMsg)
        End If
    '******


'----
' Done
'----
ExitHandler:
    ' NOTE 1:
    ' MUST CALL GetErrInfo() here first before another action
    Call GetErrInfo(msMODULE_ID, sProcID, typErrInfo, Erl)
    Call LogProcOut(msMODULE_ID, sProcID, typErrInfo, oLogCtrl)
    On Error Resume Next

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
        Call HandleError(False, typErrInfo, , oLogCtrl, True)
    End If
End Sub

'================================================================================
' Sub: Prt_WS_RENESAS_SHIPPING()
'--------------------------------------------------------------------------------
' Description:  <This subroutine will print out a label for WS-> SFG >
'--------------------------------------------------------------------------------
' Author:       <Nelson Chen>, <MXIC> 2008/06/03
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
Public Sub Prt_WS_RENESAS_SHIPPING(ByVal sLotNo As String, _
                                ByVal sProdNo As String, _
                                ByVal sWQty As String, _
                                ByVal sQty As String, _
                                ByVal sPartNo As String, _
                                ByRef arrWaferId() As String, _
                                ByRef arrWaferIdChipQty() As String, _
                                ByVal sPrintDate As String, _
                                ByVal sUserID As String, _
                                ByVal sPrinterServer As String, _
                                ByVal moCwMbx As Object, _
                                ByVal oLogCtrl As Object)

On Error GoTo ExitHandler:

Dim sMsg            As String
Dim sProcID         As String
Dim typErrInfo      As tErrInfo
Dim lErrCode        As Long
Dim sErrMsg         As String
'----
' Init
'----
    sProcID = "Prt_WS_RENESAS_SHIPPING"
    Call LogProcIn(msMODULE_ID, sProcID, oLogCtrl)

'----
' Condition Checking
'----

'----
' Action
'----
    '******
    'Print Label Info
    '******
        sMsg = sMsg & "^XA"
        sMsg = sMsg & "^PRC"
        sMsg = sMsg & "^LH0,0^FS"
        sMsg = sMsg & "^LL1441"
        sMsg = sMsg & "^MD0"
        sMsg = sMsg & "^MNY"
        sMsg = sMsg & "^LH0,0^FS"
        sMsg = sMsg & "^FO38,370^A0B,30,0^CI0^FR^FDLOT:^FS"
        sMsg = sMsg & "^FO38,200^A0B,30,0^CI0^FR^FD" & sLotNo & "^FS"
        sMsg = sMsg & "^BY2,2.0^FO70,38^BAB,40,N,,^FR^FD" & sLotNo & "^FS"
        
        sMsg = sMsg & "^FO119,350^A0B,30,0^CI0^FR^FDPART:^FS"
        sMsg = sMsg & "^FO119,24^A0B,30,0^CI0^FR^FD" & sPartNo & "^FS"
        sMsg = sMsg & "^BY2,2,40^FO154,18^BAB,,N,,^FR^FD" & sPartNo & "^FS"
        sMsg = sMsg & "^FO202,350^A0B,30,0^CI0^FR^FDPROD:^FS"
        sMsg = sMsg & "^FO202,80^A0B,30,0^CI0^FR^FD" & sProdNo & "^FS"
        sMsg = sMsg & "^FO256,16^A0N,30,0^CI0^FR^FD1:^FS"
        sMsg = sMsg & "^FO284,16^A0N,30,0^CI0^FR^FD" & arrWaferId(1) & "^FS"
        sMsg = sMsg & "^FO465,16^A0N,30,0^CI0^FR^FD" & arrWaferIdChipQty(1) & "^FS"
        sMsg = sMsg & "^FO256,51^A0N,30,0^CI0^FR^FD2:^FS"
        sMsg = sMsg & "^FO284,51^A0N,30,0^CI0^FR^FD" & arrWaferId(2) & "^FS"
        sMsg = sMsg & "^FO465,51^A0N,30,0^CI0^FR^FD" & arrWaferIdChipQty(2) & "^FS"
        sMsg = sMsg & "^FO256,87^A0N,30,0^CI0^FR^FD3:^FS"
        sMsg = sMsg & "^FO284,87^A0N,30,0^CI0^FR^FD" & arrWaferId(3) & "^FS"
        sMsg = sMsg & "^FO465,87^A0N,30,0^CI0^FR^FD" & arrWaferIdChipQty(3) & "^FS"
        sMsg = sMsg & "^FO256,121^A0N,30,0^CI0^FR^FD4:^FS"
        sMsg = sMsg & "^FO284,121^A0N,30,0^CI0^FR^FD" & arrWaferId(4) & "^FS"
        sMsg = sMsg & "^FO465,121^A0N,30,0^CI0^FR^FD" & arrWaferIdChipQty(4) & "^FS"
        sMsg = sMsg & "^FO256,157^A0N,30,0^CI0^FR^FD5:^FS"
        sMsg = sMsg & "^FO284,157^A0N,30,0^CI0^FR^FD" & arrWaferId(5) & "^FS"
        sMsg = sMsg & "^FO465,157^A0N,30,0^CI0^FR^FD" & arrWaferIdChipQty(5) & "^FS"
        sMsg = sMsg & "^FO256,192^A0N,30,0^CI0^FR^FD6:^FS"
        sMsg = sMsg & "^FO284,192^A0N,30,0^CI0^FR^FD" & arrWaferId(6) & "^FS"
        sMsg = sMsg & "^FO466,192^A0N,30,0^CI0^FR^FD" & arrWaferIdChipQty(6) & "^FS"
        sMsg = sMsg & "^FO256,227^A0N,30,0^CI0^FR^FD7:^FS"
        sMsg = sMsg & "^FO284,227^A0N,30,0^CI0^FR^FD" & arrWaferId(7) & "^FS"
        sMsg = sMsg & "^FO466,227^A0N,30,0^CI0^FR^FD" & arrWaferIdChipQty(7) & "^FS"
        sMsg = sMsg & "^FO256,262^A0N,30,0^CI0^FR^FD8:^FS"
        sMsg = sMsg & "^FO284,262^A0N,30,0^CI0^FR^FD" & arrWaferId(8) & "^FS"
        sMsg = sMsg & "^FO466,262^A0N,30,0^CI0^FR^FD" & arrWaferIdChipQty(8) & "^FS"
        sMsg = sMsg & "^FO256,298^A0N,30,0^CI0^FR^FD9:^FS"
        sMsg = sMsg & "^FO284,298^A0N,30,0^CI0^FR^FD" & arrWaferId(9) & "^FS"
        sMsg = sMsg & "^FO466,298^A0N,30,0^CI0^FR^FD" & arrWaferIdChipQty(9) & "^FS"
        sMsg = sMsg & "^FO242,332^A0N,30,0^CI0^FR^FD10:^FS"
        sMsg = sMsg & "^FO284,332^A0N,30,0^CI0^FR^FD" & arrWaferId(10) & "^FS"
        sMsg = sMsg & "^FO466,332^A0N,30,0^CI0^FR^FD" & arrWaferIdChipQty(10) & "^FS"
        sMsg = sMsg & "^FO591,16^A0N,30,0^CI0^FR^FD11:^FS"
        sMsg = sMsg & "^FO631,16^A0N,30,0^CI0^FR^FD" & arrWaferId(11) & "^FS"
        sMsg = sMsg & "^FO813,16^A0N,30,0^CI0^FR^FD" & arrWaferIdChipQty(11) & "^FS"
        sMsg = sMsg & "^FO591,51^A0N,30,0^CI0^FR^FD12:^FS"
        sMsg = sMsg & "^FO631,51^A0N,30,0^CI0^FR^FD" & arrWaferId(12) & "^FS"
        sMsg = sMsg & "^FO813,51^A0N,30,0^CI0^FR^FD" & arrWaferIdChipQty(12) & "^FS"
        sMsg = sMsg & "^FO591,87^A0N,30,0^CI0^FR^FD13:^FS"
        sMsg = sMsg & "^FO631,87^A0N,30,0^CI0^FR^FD" & arrWaferId(13) & "^FS"
        sMsg = sMsg & "^FO813,87^A0N,30,0^CI0^FR^FD" & arrWaferIdChipQty(13) & "^FS"
        sMsg = sMsg & "^FO591,121^A0N,30,0^CI0^FR^FD14:^FS"
        sMsg = sMsg & "^FO631,121^A0N,30,0^CI0^FR^FD" & arrWaferId(14) & "^FS"
        sMsg = sMsg & "^FO813,121^A0N,30,0^CI0^FR^FD" & arrWaferIdChipQty(14) & "^FS"
        sMsg = sMsg & "^FO591,157^A0N,30,0^CI0^FR^FD15:^FS"
        sMsg = sMsg & "^FO631,157^A0N,30,0^CI0^FR^FD" & arrWaferId(15) & "^FS"
        sMsg = sMsg & "^FO813,157^A0N,30,0^CI0^FR^FD" & arrWaferIdChipQty(15) & "^FS"
        sMsg = sMsg & "^FO591,192^A0N,30,0^CI0^FR^FD16:^FS"
        sMsg = sMsg & "^FO631,192^A0N,30,0^CI0^FR^FD" & arrWaferId(16) & "^FS"
        sMsg = sMsg & "^FO813,192^A0N,30,0^CI0^FR^FD" & arrWaferIdChipQty(16) & "^FS"
        sMsg = sMsg & "^FO591,227^A0N,30,0^CI0^FR^FD17:^FS"
        sMsg = sMsg & "^FO631,227^A0N,30,0^CI0^FR^FD" & arrWaferId(17) & "^FS"
        sMsg = sMsg & "^FO813,227^A0N,30,0^CI0^FR^FD" & arrWaferIdChipQty(17) & "^FS"
        sMsg = sMsg & "^FO591,262^A0N,30,0^CI0^FR^FD18:^FS"
        sMsg = sMsg & "^FO631,262^A0N,30,0^CI0^FR^FD" & arrWaferId(18) & "^FS"
        sMsg = sMsg & "^FO813,262^A0N,30,0^CI0^FR^FD" & arrWaferIdChipQty(18) & "^FS"
        sMsg = sMsg & "^FO591,298^A0N,30,0^CI0^FR^FD19:^FS"
        sMsg = sMsg & "^FO631,298^A0N,30,0^CI0^FR^FD" & arrWaferId(19) & "^FS"
        sMsg = sMsg & "^FO813,298^A0N,30,0^CI0^FR^FD" & arrWaferIdChipQty(19) & "^FS"
        sMsg = sMsg & "^FO591,332^A0N,30,0^CI0^FR^FD20:^FS"
        sMsg = sMsg & "^FO631,332^A0N,30,0^CI0^FR^FD" & arrWaferId(20) & "^FS"
        sMsg = sMsg & "^FO813,332^A0N,30,0^CI0^FR^FD" & arrWaferIdChipQty(20) & "^FS"
        sMsg = sMsg & "^FO943,16^A0N,30,0^CI0^FR^FD21:^FS"
        sMsg = sMsg & "^FO983,16^A0N,30,0^CI0^FR^FD" & arrWaferId(21) & "^FS"
        sMsg = sMsg & "^FO1164,16^A0N,30,0^CI0^FR^FD" & arrWaferIdChipQty(21) & "^FS"
        sMsg = sMsg & "^FO943,51^A0N,30,0^CI0^FR^FD22:^FS"
        sMsg = sMsg & "^FO983,51^A0N,30,0^CI0^FR^FD" & arrWaferId(22) & "^FS"
        sMsg = sMsg & "^FO1164,51^A0N,30,0^CI0^FR^FD" & arrWaferIdChipQty(22) & "^FS"
        sMsg = sMsg & "^FO943,87^A0N,30,0^CI0^FR^FD23:^FS"
        sMsg = sMsg & "^FO983,87^A0N,30,0^CI0^FR^FD" & arrWaferId(23) & "^FS"
        sMsg = sMsg & "^FO1164,87^A0N,30,0^CI0^FR^FD" & arrWaferIdChipQty(23) & "^FS"
        sMsg = sMsg & "^FO943,121^A0N,30,0^CI0^FR^FD24:^FS"
        sMsg = sMsg & "^FO983,121^A0N,30,0^CI0^FR^FD" & arrWaferId(24) & "^FS"
        sMsg = sMsg & "^FO1164,121^A0N,30,0^CI0^FR^FD" & arrWaferIdChipQty(24) & "^FS"
        sMsg = sMsg & "^FO943,157^A0N,30,0^CI0^FR^FD25:^FS"
        sMsg = sMsg & "^FO983,157^A0N,30,0^CI0^FR^FD" & arrWaferId(25) & "^FS"
        sMsg = sMsg & "^FO1164,157^A0N,30,0^CI0^FR^FD" & arrWaferIdChipQty(25) & "^FS"
        
        'Added by Jack Hsieh 2008/07/11 [M200807015] <Start>
        'Add a text line : "On line"
        'Modify 220 to 200 by Tony on 20090612 for Req.JC900074
        sMsg = sMsg & "^FO943,200^A0N,35,0^CI0^FR^FDOn line^FS" 'Added by Jack
        'Added by Jack Hsieh 2008/07/11 [M200807015] <End>
        
        sMsg = sMsg & "^FO252,384^A0N,35,0^CI0^FR^FDTTL QTY:^FS"
        sMsg = sMsg & "^FO390,384^A0N,35,0^CI0^FR^FD" & sWQty & "PCS,^FS"
        sMsg = sMsg & "^FO506,384^A0N,35,0^CI0^FR^FD" & sQty & "EA,^FS"
        
        'Modify 264 to 244 by Tony on 20090612 for Req.JC900074
        sMsg = sMsg & "^FO943,244^A0N,35,0^CI0^FR^FDID No:^FS"
        
        'Modify 264 to 244 by Tony on 20090612 for Req.JC900074
        sMsg = sMsg & "^FO1039,244^A0N,35,0^CI0^FR^FD" & sUserID & "^FS"
        
        sMsg = sMsg & "^FO694,384^A0N,35,0^CI0^FR^FDDATE:^FS"
        sMsg = sMsg & "^FO794,384^A0N,35,0^CI0^FR^FD" & sPrintDate & "^FS" 'sPrintDate-->2008/5/28
        
        'Modify 308 to 288 by Tony on 20090612 for Req.JC900074
        sMsg = sMsg & "^FO943,288^A0N,35,0^CI0^FR^FDTQA:^FS"
        
        'Add by Tony on 20090612 for Req.JC900074
        sMsg = sMsg & "^FO1030,284^GB190,140,4^FS"
        
        sMsg = sMsg & "^PQ1,0,0,N"
        sMsg = sMsg & "^XZ"
    
'        sMsg = sMsg & "^XA^LH10,00^LL500"
'        sMsg = sMsg & "^FO00,05^GB780,475,3^FS"
'        sMsg = sMsg & "^FO30,20^A0,27,25^FDLOT NO: " & sLotNo & "^FS"
'        sMsg = sMsg & "^FO30,130^A0,27,25^FDPRODUCT NO: " & sProdNo & "^FS"
'        sMsg = sMsg & "^FO30,235^A0,27,25^FDWAFER PCS:  " & sWqty & "PCS^FS"
'        sMsg = sMsg & "^FO30,340^A0,27,25^FDQUANTITY:  " & sQty & "EA^FS"
'        sMsg = sMsg & "^FO30,445^A0,27,25^FDPART NO: " & sPartNo & "^FS"
'        sMsg = sMsg & "^FO30,55^BY2,2,55^B3,,,N^FD" & sLotNo & "^FS"
'        sMsg = sMsg & "^FO30,160^BY2,2,55^B3,,,N^FD" & sProdNo & "^FS"
'        sMsg = sMsg & "^FO30,265^BY2,2,55^B3,,,N^FD" & sWqty & "^FS"
'        sMsg = sMsg & "^FO30,370^BY2,2,55^B3,,,N^FD" & sQty & "^FS"
'        sMsg = sMsg & "^FO425,385^GB355,50,3^FS"
'        sMsg = sMsg & "^FO425,425^GB0,50,3^FS"
'        sMsg = sMsg & "^FO440,400^A0,27,25^FDDATE CODE:^FS"
'        sMsg = sMsg & "^FO440,440^A0,27,25^FDQA:^FS"
'        sMsg = sMsg & "^PQ1"
'        sMsg = sMsg & "^XZ"
    '******
    ' End of Label Info
    '******

    '******
    'Send message to TDS
    '******
        
        'Add by Tony Start on 20110117 for Req.JC201100003
        If InStr(1, sPrinterServer, "@") <> 0 Then
            sPrinterServer = VBA.Left(sPrinterServer, (InStr(1, sPrinterServer, "@")) - 1)
        End If
        'Add by Tony End on 20110117 for Req.JC201100003
        
        If Not moCwMbx.Send("MBX_" & sPrinterServer, sMsg) Then
            Call moCwMbx.GetError(lErrCode, sErrMsg)
            Call RaiseError(lErrCode, sErrMsg)
        End If
    '******
'----
' Done
'----
ExitHandler:
    ' NOTE 1:
    ' MUST CALL GetErrInfo() here first before another action
    Call GetErrInfo(msMODULE_ID, sProcID, typErrInfo, Erl)
    Call LogProcOut(msMODULE_ID, sProcID, typErrInfo, oLogCtrl)
    On Error Resume Next

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
        Call HandleError(False, typErrInfo, , oLogCtrl, True)
    End If
End Sub
'Added by Jack on 2011/04/18 for JC201100104.
'Papameter :
'<1> sLotNo     : LotID
'<2> sProdNo    : IPN
'<3> sTestMode  : StepName
'<4> bIsPassBin : true : PassBin; false : FailBin.
Public Sub Prt_FT_FT_BIN_CARD_LABEL(ByVal sLotNo As String, _
                                ByVal sProdNo As String, _
                                ByVal sTestMode As String, _
                                ByVal sStepID As String, _
                                ByVal sPrinterServer As String, _
                                ByVal bIsPassBin As Boolean, _
                                ByVal moCwMbx As Object, _
                                ByVal oLogCtrl As Object, _
                                Optional ByVal sCargradeFlag As String)
On Error GoTo ExitHandler:

Dim sMsg            As String
Dim sProcID         As String
Dim typErrInfo      As tErrInfo
Dim lErrCode        As Long
Dim sErrMsg         As String
'----
' Init
'----
    sProcID = "Prt_FT_FT_BIN_CARD_LABEL"
    Call LogProcIn(msMODULE_ID, sProcID, oLogCtrl)

'----
' Condition Checking
'----

'----
' Action
'----
    On Error Resume Next
    '******
    'Print Label Info
    '******
    sMsg = ""
    
    If bIsPassBin Then '列印Pass BIN
        sMsg = sMsg & "^XA"
        sMsg = sMsg & "^PRC "
        sMsg = sMsg & "^LH0,0^FS"
        sMsg = sMsg & "^LL751"
        sMsg = sMsg & "^MD0"
        sMsg = sMsg & "^MNM"
        sMsg = sMsg & "^LH144,0^FS"
        sMsg = sMsg & "^FO109,50^A0N,48,71^CI13^FR^FDMacronix ^FS"
        sMsg = sMsg & "^FO16,258^A0N,27,28^CI13^FR^FDLOT_NO   :" & sLotNo & "^FS"
        sMsg = sMsg & "^BY2,3.0^FO150,300^BCN,50,N,Y,N^FR^FD>:" & sLotNo & "^FS" 'LotID's BarCode.
        sMsg = sMsg & "^FO16,386^A0N,27,28^CI13^FR^FDTest Mode :" & sStepID & " " & sTestMode & "^FS"
        sMsg = sMsg & "^FO16,470^A0N,31,29^CI13^FR^FDTA-ID      :^FS"
        sMsg = sMsg & "^FO296,467^A0N,29,25^CI13^FR^FDDate   :^FS"
        sMsg = sMsg & "^FO16,533^A0N,26,37^CI13^FR^FDBIN ^FS"
        sMsg = sMsg & "^FO16,569^A0N,26,37^CI13^FR^FDMODE   :^FS"
        sMsg = sMsg & "^FO16,631^A0N,27,40^CI13^FR^FDQ'TY    :^FS"
        sMsg = sMsg & "^FO16,709^A0N,27,40^CI13^FR^FDNOTE   :^FS"
        sMsg = sMsg & "^FO88,109^A0N,48,60^CI13^FR^FDPASS BIN CARD^FS"
        sMsg = sMsg & "^FO16,184^A0N,27,32^CI13^FR^FDProd_ID  :" & sProdNo & "^FS"
        sMsg = sMsg & "^FO5,29^GB525,719,4^FS"
        sMsg = sMsg & "^FO11,234^GB512,0,4^FS"
        sMsg = sMsg & "^FO11,357^GB512,0,4^FS"
        sMsg = sMsg & "^FO11,437^GB512,0,4^FS"
        sMsg = sMsg & "^FO11,520^GB512,0,4^FS"
        sMsg = sMsg & "^FO13,599^GB512,0,4^FS"
        sMsg = sMsg & "^FO8,677^GB512,0,4^FS"
        sMsg = sMsg & "^PQ1,0,0,N"
        sMsg = sMsg & "^XZ"
        sMsg = sMsg & "^FX End of job"
        sMsg = sMsg & "^XA"
        sMsg = sMsg & "^IDR:ID*.*"
        sMsg = sMsg & "^XZ"
    Else '列印Fail BIN
        sMsg = sMsg & "^XA"
        sMsg = sMsg & "^PRC "
        sMsg = sMsg & "^LH0,0^FS"
        sMsg = sMsg & "^LL751"
        sMsg = sMsg & "^MD0"
        sMsg = sMsg & "^MNM"
        sMsg = sMsg & "^LH128,0^FS"
        sMsg = sMsg & "^FO123,42^A0N,48,71^CI13^FR^FDMacronix ^FS"
        sMsg = sMsg & "^FO21,517^A0N,27,40^CI13^FR^FDBIN2:______^FS"
        sMsg = sMsg & "^FO21,562^A0N,27,40^CI13^FR^FDBIN4:______^FS"
        sMsg = sMsg & "^FO19,258^A0N,27,28^CI13^FR^FDLOT_NO    :" & sLotNo & "^FS"
        sMsg = sMsg & "^BY2,3.0^FO150,300^BCN,50,N,Y,N^FR^FD>:" & sLotNo & "^FS" 'LotID's BarCode.
        sMsg = sMsg & "^FO19,368^A0N,27,28^CI13^FR^FDTest Mode :" & sStepID & " " & sTestMode & "^FS"
        sMsg = sMsg & "^FO19,445^A0N,27,31^CI13^FR^FDTA-ID    :^FS"
        sMsg = sMsg & "^FO264,446^A0N,31,28^CI13^FR^FDDate   :^FS"
        sMsg = sMsg & "^FO21,607^A0N,27,40^CI13^FR^FDBIN6:______^FS"
        sMsg = sMsg & "^FO21,647^A0N,27,40^CI13^FR^FDBIN_:______^FS"
        sMsg = sMsg & "^FO21,693^A0N,27,40^CI13^FR^FDNOTE   :^FS"
        sMsg = sMsg & "^FO88,109^A0N,48,60^CI13^FR^FDFAIL BIN CARD^FS"
        sMsg = sMsg & "^FO19,181^A0N,27,32^CI13^FR^FDProd_ID  :" & sProdNo & "^FS"
        sMsg = sMsg & "^FO282,647^A0N,27,40^CI13^FR^FDREJ :______^FS"
        sMsg = sMsg & "^FO282,607^A0N,27,40^CI13^FR^FDBIN7:______^FS"
        sMsg = sMsg & "^FO282,562^A0N,27,40^CI13^FR^FDBIN5:______^FS"
        sMsg = sMsg & "^FO282,517^A0N,27,40^CI13^FR^FDBIN3:______^FS"
        sMsg = sMsg & "^FO11,11^GB525,719,4^FS"
        sMsg = sMsg & "^FO13,237^GB520,0,4^FS"
        sMsg = sMsg & "^FO13,352^GB520,0,4^FS"
        sMsg = sMsg & "^FO13,415^GB520,0,4^FS"
        sMsg = sMsg & "^FO11,493^GB520,0,4^FS"
        sMsg = sMsg & "^PQ1,0,0,N"
        sMsg = sMsg & "^XZ"
        sMsg = sMsg & "^FX End of job"
        sMsg = sMsg & "^XA"
        sMsg = sMsg & "^IDR:ID*.*"
        sMsg = sMsg & "^XZ"
    End If
    
    '******
    ' End of Label Info
    '******

    '******
    'Send message to TDS
    '******
        If InStr(1, sPrinterServer, "@") <> 0 Then
            sPrinterServer = VBA.Left(sPrinterServer, (InStr(1, sPrinterServer, "@")) - 1)
        End If
        
        If Not moCwMbx.Send("MBX_" & sPrinterServer, sMsg) Then
            Call moCwMbx.GetError(lErrCode, sErrMsg)
            Call RaiseError(lErrCode, sErrMsg)
        End If
    '******
'----
' Done
'----
ExitHandler:
    ' NOTE 1:
    ' MUST CALL GetErrInfo() here first before another action
    Call GetErrInfo(msMODULE_ID, sProcID, typErrInfo, Erl)
    Call LogProcOut(msMODULE_ID, sProcID, typErrInfo, oLogCtrl)
    On Error Resume Next

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
        Call HandleError(False, typErrInfo, , oLogCtrl, True)
    End If
End Sub


'================================================================================
' Sub: Prt_WS_WS_ENG_LOC_LABEL()
'--------------------------------------------------------------------------------
' Description:  <This subroutine will print out a label for FT-> SFG >
'--------------------------------------------------------------------------------
' Author:       <Alex Chan>, <CIT> 2002/02/07
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
'================================================================================
Public Sub Prt_WS_WS_ENG_LOC_LABEL(ByVal sLotNo As String, _
                                ByVal sPrinterServer As String, _
                                ByVal moCwMbx As Object, _
                                ByVal oLogCtrl As Object, _
                                ByVal moProRawSql As Object)


On Error GoTo ExitHandler:

Dim sMsg            As String
Dim sProcID         As String
Dim typErrInfo      As tErrInfo
Dim lErrCode        As Long
Dim sErrMsg         As String

Dim sSQL As String
Dim colRS As Collection
Dim sWQty As String
Dim sOwnerId As String
Dim sOwnerDep As String
Dim sLocId As String

'----
' Init
'----
    sProcID = "Prt_WS_WS_ENG_LOC_LABEL"
    
    Call LogProcIn(msMODULE_ID, sProcID, oLogCtrl)
    
    
    sSQL = " select A." & gsCAT_TWELM_LOTID & " ,A." & gsCAT_TWELM_WQTY & _
           " ,A." & gsCAT_TWELM_OWNERID & " ,A." & gsCAT_TWELM_OWNERDEP & _
           " ,A." & gsCAT_TWELM_LOCID & _
           " FROM " & gsCAT_TBL_WS_ENG_LOC_MASTER & " A " & _
           " WHERE A." & gsCAT_TWELM_LOTID & "='" & sLotNo & "' "
           
    Set colRS = moProRawSql.QueryDatabase(sSQL)
    If colRS.Count > 0 Then
        sWQty = colRS.Item(1).Item(gsCAT_TWELM_WQTY)
        sOwnerId = colRS.Item(1).Item(gsCAT_TWELM_OWNERID)
        sOwnerDep = colRS.Item(1).Item(gsCAT_TWELM_OWNERDEP)
        sLocId = colRS.Item(1).Item(gsCAT_TWELM_LOCID)
    End If
'----
' Condition Checking
'----

'----
' Action
'----
    On Error Resume Next
    '******
    'Print Label Info
    '******
        sMsg = sMsg & "^XA"
        sMsg = sMsg & "^PRC "
        sMsg = sMsg & "^LH0,0^FS"
        sMsg = sMsg & "^LL480"
        sMsg = sMsg & "^MD0"
        sMsg = sMsg & "^MNY"
        sMsg = sMsg & "^LH0,0^FS"
        sMsg = sMsg & "^CWI,V10SPR_A.FNT^FS"
        sMsg = sMsg & "^FO29,37^AIN,45,0^CI0^FR^FDLOTID :" & sLotNo & "^FS"
        sMsg = sMsg & "^CWJ,U10SPR_A.FNT^FS"
        sMsg = sMsg & "^FO29,176^AJN,45,0^CI0^FR^FDWQTY :" & sWQty & "^FS"
        sMsg = sMsg & "^CWK,U1WSPR_A.FNT^FS"
        sMsg = sMsg & "^FO27,253^AKN,45,0^CI0^FR^FDOwnerID :" & sOwnerId & "^FS"
        sMsg = sMsg & "^CWJ,U10SPR_A.FNT^FS"
        sMsg = sMsg & "^FO29,330^AJN,45,0^CI0^FR^FDOwnerDep :" & sOwnerDep & "^FS"
        sMsg = sMsg & "^BY2,3.0^FO36,107^B3N,N,40,N,Y^FR^FD" & sLotNo & "^FS"
        sMsg = sMsg & "^FO13,19^GB757,434,4^FS"
        sMsg = sMsg & "^PQ1,0,0,N"
        sMsg = sMsg & "^XZ"
        sMsg = sMsg & "^FX End of job"
        sMsg = sMsg & "^XA"
        sMsg = sMsg & "^IDR:ID*.*"
        sMsg = sMsg & "^XZ"

    '******
    ' End of Label Info
    '******

    '******
    'Send message to TDS
    '******
        
        'Add by Tony Start on 20110117 for Req.JC201100003
        If InStr(1, sPrinterServer, "@") <> 0 Then
            sPrinterServer = VBA.Left(sPrinterServer, (InStr(1, sPrinterServer, "@")) - 1)
        End If
        'Add by Tony End on 20110117 for Req.JC201100003
        
        If Not moCwMbx.Send("MBX_" & sPrinterServer, sMsg) Then
            Call moCwMbx.GetError(lErrCode, sErrMsg)
            Call RaiseError(lErrCode, sErrMsg)
        End If
    '******

'----
' Done
'----
ExitHandler:
    ' NOTE 1:
    ' MUST CALL GetErrInfo() here first before another action
    Call GetErrInfo(msMODULE_ID, sProcID, typErrInfo, Erl)
    Call LogProcOut(msMODULE_ID, sProcID, typErrInfo, oLogCtrl)
    On Error Resume Next

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
        Call HandleError(False, typErrInfo, , oLogCtrl, True)
    End If
End Sub
'================================================================================
' Sub: Prt_FT_BOX_COUNTING_LABEL()
'--------------------------------------------------------------------------------
' Description:  Req.JC201200283
'--------------------------------------------------------------------------------
' Create Date: 20121001
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
'================================================================================
Public Sub Prt_FT_BOX_COUNTING_LABEL(ByVal oLot As FwLot, _
                                ByVal sPrintCount As String, _
                                ByVal sPrinterServer As String, _
                                ByVal moCwMbx As Object, _
                                ByVal oLogCtrl As Object)

On Error GoTo ExitHandler:
Dim sProcID         As String
Dim typErrInfo      As tErrInfo
Dim sMsg            As String
Dim lErrCode        As Long
Dim sErrMsg         As String
Dim sIPN    As String


'----
' Init
'----
    sProcID = "Prt_FT_BOX_COUNTING_LABEL"
    
    Call LogProcIn(msMODULE_ID, sProcID, oLogCtrl)
    
'----
' Condition Checking
'----

'----
' Action
'----
    On Error Resume Next
    
    sIPN = oLot.CustomAttributes.Item(gsLOT_CUSTOMATTR_IPN)
    
    '******
    'Print Label Info
    '******
        
'        sMsg = sMsg & "^XA"
'        sMsg = sMsg & "^PRC "
'        sMsg = sMsg & "^LH0,0^FS"
'        sMsg = sMsg & "^LL400"
'        sMsg = sMsg & "^MD0"
'        sMsg = sMsg & "^MNY"
'        sMsg = sMsg & "^LH0,0^FS"
'        sMsg = sMsg & "^FO157,38^A0N,43,37^CI13^FR^FD   Macronix^FS"
'        sMsg = sMsg & "^FO157,83^A0N,43,37^CI13^FR^FDBox Counting^FS"
'        sMsg = sMsg & "^FO29,159^A0N,31,28^CI13^FR^FDProd_ID:" & sIPN & "^FS"
'        sMsg = sMsg & "^FO29,217^A0N,31,28^CI13^FR^FDLOT_NO:" & oLot.Id & "^FS"
'         sMsg = sMsg & "^BY2,3.0^FO29,275^BCN,50,N,Y,N^FR^FD>" & oLot.Id & "^FS"  'LotID's BarCode.
'        'sMsg = sMsg & "^BY2,3.0^FO150,300^BCN,50,N,Y,N^FR^FD>" & oLot.Id & "^FS"  'LotID's BarCode.
'        sMsg = sMsg & "^FO370,313^A0N,31,28^CI13^FR^FDBOX:" & sPrintCount & "^FS"
'        sMsg = sMsg & "^FO11,19^GB525,354,4^FS"
'        sMsg = sMsg & "^FO16,126^GB512,0,4^FS"
'        sMsg = sMsg & "^PQ1,0,0,N"
'        sMsg = sMsg & "^XZ"
'        sMsg = sMsg & "^FX End of job"
'        sMsg = sMsg & "^XA"
'        sMsg = sMsg & "^IDR:ID*.*"
'        sMsg = sMsg & "^XZ"
        
        sMsg = sMsg & "^XA"
        sMsg = sMsg & "^PRC "
        sMsg = sMsg & "^LH0,0^FS"
        sMsg = sMsg & "^LL400"
        sMsg = sMsg & "^MD0"
        sMsg = sMsg & "^MNM"
        sMsg = sMsg & "^LH128,0^FS"
        sMsg = sMsg & "^FO168,30^A0N,43,37^CI13^FR^FDBox Counting^FS"
        sMsg = sMsg & "^FO24,129^A0N,31,28^CI13^FR^FDProd_ID:" & sIPN & "^FS"
        sMsg = sMsg & "^FO27,217^A0N,31,28^CI13^FR^FDLOT_NO:" & oLot.Id & "^FS"
        sMsg = sMsg & "^BY2,3.0^FO30,290^BCN,50,N,Y,N^FR^FD>:" & oLot.Id & "^FS" 'LotID's BarCode.
        '^FO30,表示由前面算過來的距離,愈大愈右邊
        sMsg = sMsg & "^FO290,196^A0N,69,61^CI13^FR^FDBOX:" & sPrintCount & "^FS"
        sMsg = sMsg & "^FO11,8^GB528,360,4^FS"
        sMsg = sMsg & "^FO13,83^GB512,0,4^FS"
        sMsg = sMsg & "^PQ1,0,0,N"
        sMsg = sMsg & "^XZ"
        sMsg = sMsg & "^FX End of job"
        sMsg = sMsg & "^XA"
        sMsg = sMsg & "^IDR:ID*.*"
        sMsg = sMsg & "^XZ"
        

    '******
    ' End of Label Info
    '******

    '******
    'Send message to TDS
    '******
        
        'Add by Tony Start on 20110117 for Req.JC201100003
        If InStr(1, sPrinterServer, "@") <> 0 Then
            sPrinterServer = VBA.Left(sPrinterServer, (InStr(1, sPrinterServer, "@")) - 1)
        End If
        'Add by Tony End on 20110117 for Req.JC201100003
        
        If Not moCwMbx.Send("MBX_" & sPrinterServer, sMsg) Then
            Call moCwMbx.GetError(lErrCode, sErrMsg)
            Call RaiseError(lErrCode, sErrMsg)
        End If
    '******

'----
' Done
'----
ExitHandler:
    ' NOTE 1:
    ' MUST CALL GetErrInfo() here first before another action
    Call GetErrInfo(msMODULE_ID, sProcID, typErrInfo, Erl)
    Call LogProcOut(msMODULE_ID, sProcID, typErrInfo, oLogCtrl)
    On Error Resume Next

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
        Call HandleError(False, typErrInfo, , oLogCtrl, True)
    End If
End Sub

'================================================================================
' function: RefreshPrinterServer()
'--------------------------------------------------------------------------------
' Description:  <Type your function description here...>
'--------------------------------------------------------------------------------
' Author:       Sam, CIT 2012-12-03
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
Public Sub RefreshPrinterServer(ByRef cboServer As ComboBox, ByVal sStage As String, ByVal sSpecNo As String, _
                                ByVal oLogCtrl As Object, ByVal moProRawSql As Object)
On Error GoTo ExitHandler:
Dim sProcID As String
Dim typErrInfo As tErrInfo
Dim sTable As String
Dim sColumn As String
Dim sWhere As String
Dim sOther  As String
Dim colRaws As Collection
Dim oRaws As FwStrings
Dim sSQL As String
Dim lIdx As Long
Dim iIdx As Integer
Dim vServer As Variant

Dim colRS As Collection

'----
' Init
'----
    sProcID = "RefreshPrinterServer"
    Call LogProcIn(msMODULE_ID, sProcID, oLogCtrl)
'----
' Condition Chking
'----


'----
' Action
'----
    sTable = gsCAT_TBL_LABEL_SPEC & " A "


    sColumn = " A." & gsCAT_TLS_SERVER_NAME
        
    sWhere = "A." & gsCAT_TLS_LABEL_SPECNO & "='" & sSpecNo & "' and " & _
             "A." & gsCAT_TLS_STAGE & "='" & sStage & "'"

    sSQL = "SELECT " & sColumn & " FROM " & sTable & " WHERE " & sWhere
    
'    Set colRaws = SelectRawSqlTxn(moAppLog, moFwWIP, moFwWF, moCwMbx, "CAT", sTable, sColumn, sWhere, sOther)
    Set colRaws = moProRawSql.QueryDatabase(sSQL)
    If colRaws Is Nothing Then
        Call RaiseError(glERR_INVALIDOBJECT, _
                    FormatErrorText(gsETX_INVALIDOBJECT, "Collection"))
    End If
    
    If colRaws.Count > 0 Then
        If Trim(colRaws.Item(1).Item(gsCAT_TLS_SERVER_NAME)) <> "" Then
            cboServer.Clear
            vServer = Split(colRaws.Item(1).Item(gsCAT_TLS_SERVER_NAME), ",")
            For lIdx = LBound(vServer) To UBound(vServer)
                sSQL = " select " & gsCAT_TPSR_SERVER_NAME & " , " & gsCAT_TPSR_DESCRIPTION & "  " & _
                        " from " & gsCAT_TBL_PRINTER_SERVER & " " & _
                        " where " & gsCAT_TPSR_SERVER_NAME & "= '" & vServer(lIdx) & "' " & _
                        " and " & gsCAT_TPSR_DELETE_FLAG & "= 'N' "
                Set colRS = moProRawSql.QueryDatabase(sSQL)
                For iIdx = 1 To colRS.Count
                    cboServer.AddItem UCase(colRS.Item(iIdx).Item(gsCAT_TPSR_SERVER_NAME) & "@" & colRS.Item(iIdx).Item(gsCAT_TPSR_DESCRIPTION))
                Next
            Next
            'Add by Sam Start on 20210505
            cboServer.ListIndex = 0
            'Add by Sam End on 20210505
        Else
            cboServer.Clear
            sSQL = " select " & gsCAT_TPSR_SERVER_NAME & " , " & gsCAT_TPSR_DESCRIPTION & "  " & _
                       " from " & gsCAT_TBL_PRINTER_SERVER & " " & _
                       " where  " & gsCAT_TPSR_DELETE_FLAG & "= 'N' "
            Set colRS = moProRawSql.QueryDatabase(sSQL)
            For iIdx = 1 To colRS.Count
                cboServer.AddItem UCase(colRS.Item(iIdx).Item(gsCAT_TPSR_SERVER_NAME) & "@" & colRS.Item(iIdx).Item(gsCAT_TPSR_DESCRIPTION))
            Next
                        
        End If
    End If
    
'----
' Done
'----

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
            Case glERR_INVALIDOBJECT
                ' Retry code goes here...
            Case Else
                typErrInfo.sUserText = "Fail to execute application, please call IT support!!" & vbCrLf & _
                                        "程式執行失敗, 請洽IT人員處理"
            End Select
        '---- Start of Select Case Block ----
        On Error GoTo ExitHandler:
        Call HandleError(True, typErrInfo, , oLogCtrl)
    End If
End Sub


'================================================================================
' Sub: ExportExcel()
'--------------------------------------------------------------------------------
' Description:
'--------------------------------------------------------------------------------
' Author:       Sam Chen, CIT 2014/09/25
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
Public Sub ExportExcel(ByRef spdTable As fpSpread, ByRef appXL As Object, ByRef moAppLog As Object)
On Error GoTo ExitHandler:
Dim sProcID     As String
Dim typErrInfo  As tErrInfo

Dim iPosTitle   As Integer
Dim oBook       As Object
Dim oSheet      As Object
Dim iRow        As Integer
Dim iCol        As Integer
Dim vTmp        As Variant

Dim intI As Integer
Dim vHeader As Variant

Dim iPosX As Integer
Dim iPosY As Integer
'----
' Init
'----
    sProcID = "ExportExcel"
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
    oSheet.Cells.Font.Size = 10
    With oSheet
        iPosX = 1
        iPosY = 1
        
        For intI = 1 To spdTable.MaxCols
            spdTable.Col = intI
            If spdTable.colHidden = False Then
                spdTable.GetText intI, 0, vHeader
                .Cells(iPosY, iPosX).Value = vHeader
                .Cells(iPosY, iPosX).borders.LineStyle = 1
                iPosX = iPosX + 1
            End If
        Next intI
        
        
        iPosY = 2
        For iRow = 1 To spdTable.MaxRows
            iPosX = 1
            For iCol = 1 To spdTable.MaxCols
            spdTable.Col = iCol
             If spdTable.colHidden = False Then
                spdTable.GetText iCol, iRow, vTmp
                .Cells(iPosY, iPosX).Value = CStr(vTmp)
                .Cells(iPosY, iPosX).borders.LineStyle = 1
                iPosX = iPosX + 1
             End If
            Next iCol
            iPosY = iPosY + 1
        Next iRow
        
        .Cells.Select
        .Cells.EntireColumn.AutoFit
        .range("A1").Select
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
' Sub: Prt_FT_FT_LOT_INFO()
'--------------------------------------------------------------------------------
' Description:  <This subroutine will print out a label for FT.
'--------------------------------------------------------------------------------
' Author:       <Jack Hsieh>, <MXIC> 2016/06/03
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
' [REV 01] Added by Jack on 2016/06/03 for 2016 MES Phase-5 Project (FT eRunCard Paperless Phase-2 Project)
' 1) Label Code provied by Daniel Wu.
'
'
' [REV 02] <AuthorName>, <CompanyName> <YYYY/MM/DD>
' 1) <Description goes here...>
'    <Line 2...>
'
'================================================================================
    Public Sub Prt_FT_FT_LOT_INFO(ByVal sLotNo As String, _
            ByVal sIPN As String, _
            ByVal sCQty As String, _
            ByVal sRouteId As String, _
            ByVal sLotOwner As String, _
            ByVal sDateCode As String, _
            ByVal sCargradeFlag As String, _
            ByVal sProdLine As String, _
            ByVal sCarrierType As String, _
            ByVal sExtraRomFlag As String, _
            ByVal sBrand As String, _
            ByVal sIcDrawing As String, _
            ByVal sMarkingSpecNo As String, _
            ByVal sGreen As String, _
            ByVal sCsum As String, _
            ByVal sVendorCode As String, _
            ByVal sNFflag As String, _
            ByVal sPrinterServer As String, _
            ByVal moCwMbx As Object, _
            ByVal oLogCtrl As Object)
        On Error GoTo ExitHandler:

        Dim sMsg                    As String
        Dim sProcID                 As String
        Dim typErrInfo              As tErrInfo
        Dim lErrCode                As Long
        Dim sErrMsg                 As String

        '----
        ' Init
        '----
1       sProcID = "Prt_FT_FT_LOT_INFO"
2       Call LogProcIn(msMODULE_ID, sProcID, oLogCtrl)

        '----
        ' Condition Checking
        '----

        '----
        ' Action
        '----
        On Error Resume Next
       
        '******
        'Print Label Info
        '******
        'ps : <1> Route的部份, user要求只要列印35碼, 以免超過長度.(目前平均約35碼)
        '         Mid(sRouteId, 1, 35)
        '     <2> 測試時需要 : sRouteId = sRouteId & "-ABCDEFGHIJKLMNOPQRSTUVWXYZ"
    
        'sRouteId = sRouteId & "-ABCDEFGHIJKLMNOPQRSTUVWXYZ"
        'TO-DO : 上線時, 這一行必需拿掉!!

3       sMsg = sMsg & "^XA"
4       sMsg = sMsg & "^MCY"
5       sMsg = sMsg & "^XZ"
6       sMsg = sMsg & "^XA"
7       sMsg = sMsg & "^FWN^CFD,24^LH0,0"
8       sMsg = sMsg & "^CI0^PR2^MNY^MTT^MMT^MD0^PON^PMN^LRN"
9       sMsg = sMsg & "^XZ"
10      sMsg = sMsg & "^XA"
11      sMsg = sMsg & "^DFR:TEMP_FMT.ZPL"
12      sMsg = sMsg & "^LRN"
13      sMsg = sMsg & "^XZ"
14      sMsg = sMsg & "^XA"
15      sMsg = sMsg & "^XFR:TEMP_FMT.ZPL"
16      sMsg = sMsg & "^FO9,3^GB783,464,4,B,0^FS"
17      sMsg = sMsg & "^A0N,28,28^FO20,17^FDIPN^FS"
18      sMsg = sMsg & "^A0N,28,28^FO20,52^FDLot no.^FS"
19      sMsg = sMsg & "^A0N,28,28^FO20,86^FDQty^FS"
20      sMsg = sMsg & "^A0N,28,28^FO20,123^FDRoute^FS"
21      sMsg = sMsg & "^A0N,28,28^FO20,156^FDOwner^FS"
22      sMsg = sMsg & "^A0N,28,28^FO20,190^FDCsum^FS"
23      sMsg = sMsg & "^A0N,28,28^FO20,224^FDGreen^FS"
24      sMsg = sMsg & "^A0N,28,28^FO21,258^FDBrand^FS"
25      sMsg = sMsg & "^A0N,28,28^FO22,430^FDCarrier Type^FS"
26      sMsg = sMsg & "^A0N,28,28^FO22,395^FDIC Outline^FS"
27      sMsg = sMsg & "^A0N,28,28^FO368,291^FDAEB^FS"
28      sMsg = sMsg & "^A0N,28,28^FO368,189^FDProdLine^FS"
29      sMsg = sMsg & "^A0N,28,28^FO368,224^FDXtraRom^FS"
30      sMsg = sMsg & "^A0N,28,28^FO368,258^FDN flag^FS"
31      sMsg = sMsg & "^A0N,28,28^FO22,291^FDAssy Code^FS"
32      sMsg = sMsg & "^A0N,28,28^FO22,327^FDDateCode^FS"
33      sMsg = sMsg & "^A0N,28,28^FO22,361^FDIC Marking^FS"
34      sMsg = sMsg & "^A0N,34,34^FO174,430^FD:^FS"
35      sMsg = sMsg & "^A0N,34,34^FO173,393^FD:^FS"
36      sMsg = sMsg & "^A0N,34,34^FO173,190^FD:^FS"
37      sMsg = sMsg & "^A0N,34,34^FO173,258^FD:^FS"
38      sMsg = sMsg & "^A0N,34,34^FO173,292^FD:^FS"
39      sMsg = sMsg & "^A0N,34,34^FO173,224^FD:^FS"
40      sMsg = sMsg & "^A0N,34,34^FO173,361^FD:^FS"
41      sMsg = sMsg & "^A0N,34,34^FO173,327^FD:^FS"
42      sMsg = sMsg & "^A0N,34,34^FO173,156^FD:^FS"
43      sMsg = sMsg & "^A0N,34,34^FO173,123^FD:^FS"
44      sMsg = sMsg & "^A0N,34,34^FO173,86^FD:^FS"
45      sMsg = sMsg & "^A0N,34,34^FO173,52^FD:^FS"
46      sMsg = sMsg & "^A0N,34,34^FO173,17^FD:^FS"
47      sMsg = sMsg & "^A0N,34,34^FO480,189^FD:^FS"
48      sMsg = sMsg & "^A0N,34,34^FO480,224^FD:^FS"
49      sMsg = sMsg & "^A0N,34,34^FO480,258^FD:^FS"
50      sMsg = sMsg & "^A0N,34,34^FO480,291^FD:^FS"
        
        'Modified by Jack on 2016/07/04 for eRunCard Paperless Project :
        '<1> CQty 增加 " EA".
        '<2> LotNo BarCode code-93 改 Code-128(字型).
'51      sMsg = sMsg & "^BY2^FO416,47^BAN,49,N,N,Y^FD" & sLotNo & "^FS"
'                       ^BY2^FO416,47^BCN,49,N,N,N^FD>:LOTID12345^FS
51      sMsg = sMsg & "^BY2^FO416,47^BCN,49,N,N,N^FD" & sLotNo & "^FS"
52      sMsg = sMsg & "^A0N,28,28^FO193,17^FD" & sIPN & "^FS"
53      sMsg = sMsg & "^A0N,28,28^FO193,52^FD" & sLotNo & "^FS"
54      sMsg = sMsg & "^A0N,28,28^FO193,86^FD" & sCQty & " EA^FS"
55      sMsg = sMsg & "^A0N,28,28^FO193,123^FD" & Mid(sRouteId, 1, 35) & "^FS"
56      sMsg = sMsg & "^A0N,28,28^FO193,156^FD" & sLotOwner & "^FS"
57      sMsg = sMsg & "^A0N,28,28^FO193,190^FD" & sCsum & "^FS"
58      sMsg = sMsg & "^A0N,28,28^FO193,224^FD" & sGreen & "^FS"
59      sMsg = sMsg & "^A0N,28,28^FO193,258^FD" & sBrand & "^FS"
60      sMsg = sMsg & "^A0N,28,28^FO193,291^FD" & sVendorCode & "^FS"
61      sMsg = sMsg & "^A0N,28,28^FO193,327^FD" & sDateCode & "^FS"
62      sMsg = sMsg & "^A0N,28,28^FO193,361^FD" & sMarkingSpecNo & "^FS"
63      sMsg = sMsg & "^A0N,28,28^FO193,395^FD" & sIcDrawing & "^FS"
64      sMsg = sMsg & "^A0N,28,28^FO193,430^FD" & sCarrierType & "^FS"
65      sMsg = sMsg & "^A0N,28,28^FO496,189^FD" & sProdLine & "^FS"
66      sMsg = sMsg & "^A0N,28,28^FO496,224^FD" & sExtraRomFlag & "^FS"
67      sMsg = sMsg & "^A0N,28,28^FO496,258^FD" & sNFflag & "^FS"
68      sMsg = sMsg & "^A0N,28,28^FO496,291^FD" & sCargradeFlag & "^FS"
69      sMsg = sMsg & "^PQ1,0,1,Y"
70      sMsg = sMsg & "^XZ"
71      sMsg = sMsg & "^XA"
72      sMsg = sMsg & "^IDR:TEMP_FMT.ZPL"
73      sMsg = sMsg & "^XZ"
    
        '******
        ' End of Label Info
        '******

        '******
        'Send message to MES Server MailBox.
        '******
74      If InStr(1, sPrinterServer, "@") <> 0 Then
75          sPrinterServer = VBA.Left(sPrinterServer, (InStr(1, sPrinterServer, "@")) - 1)
        End If
    
76      If Not moCwMbx.Send("MBX_" & sPrinterServer, sMsg) Then
77          Call moCwMbx.GetError(lErrCode, sErrMsg)
78          Call RaiseError(lErrCode, sErrMsg)
        End If
        '******
        '----
        ' Done
        '----
ExitHandler:
        ' NOTE 1:
        ' MUST CALL GetErrInfo() here first before another action
79      Call GetErrInfo(msMODULE_ID, sProcID, typErrInfo, Erl)
80      Call LogProcOut(msMODULE_ID, sProcID, typErrInfo, oLogCtrl)
        On Error Resume Next

ErrorHandler:
81      If typErrInfo.lErrNumber Then
            ' NOTE 2:
            ' If you have custom handling of some Errors, please
            ' UN-REMARED the following Select Case block!
            ' Also, modify if neccessarily!!!
            '---- Start of Select Case Block ----
            Select Case typErrInfo.lErrNumber
            Case glERR_INVALIDOBJECT
                ' Retry code goes here...
            Case Else
82              typErrInfo.sUserText = "Fail to execute application, please call IT support!!" & vbCrLf & _
                    "程式執行失敗, 請洽IT人員處理"
            End Select
            '---- Start of Select Case Block ----
            On Error GoTo ExitHandler:
83          Call HandleError(False, typErrInfo, , oLogCtrl, True)
        End If
    End Sub
'================================================================================
' Sub: Prt_FT_Label_PACK_INFO()
'--------------------------------------------------------------------------------
' Description:  <This subroutine will print out a label for FT.
'--------------------------------------------------------------------------------
' Author:       <Jack Hsieh>, <MXIC> 2016/06/03
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
' [REV 01] Added by Jack on 2016/06/03 for 2016 MES Phase-5 Project (FT eRunCard Paperless Phase-2 Project)
' 1) Label Code provied by Daniel Wu.
' [REV 02] Added by Jack on 2016/07/04 for 2016 MES Phase-5 Project (FT eRunCard Paperless Phase-2 Project)
'          <1> 加傳IPN (fot Auto Mode);
'          <2> LabelSpecNO串不到資料, 改用default.
'
'================================================================================
Public Sub Prt_FT_Label_PACK_INFO(ByVal sLotID As String, _
        ByVal sLabelSpecNo As String, _
        ByVal sCarrierType As String, _
        ByVal sBoxingSpecNo As String, _
        ByVal sBrand As String, _
        ByVal sPinCount As String, _
        ByVal sPackageCode As String, _
        ByVal bIsAutoPrint As Boolean, _
        ByVal sPrinterServer As String, _
        ByVal moCwMbx As Object, _
        ByVal oLogCtrl As Object, _
        ByVal moProRawSql As Object, _
        ByVal sIPN As String _
        )
    On Error GoTo ExitHandler:

    Dim sMsg                    As String
    Dim sProcID                 As String
    Dim typErrInfo              As tErrInfo
    Dim lErrCode                As Long
    Dim sErrMsg                 As String

    Dim sSQL                    As String
    Dim sTable                  As String
    Dim sColumn                 As String
    Dim sWhere                  As String
    Dim sOther                  As String
    Dim colRaws                 As Collection
    Dim oRaws                   As FwStrings

    '列印資料-1 : 欄位來自 Tbl_Ft_Label_Spec_Info.
    Dim sLabelSpecNoVer         As String
    Dim sSpecialLabelSize       As String
    Dim sSerial                 As String
    Dim sMergeLabelSpec         As String
    Dim sCopiesStdLabel         As String
    Dim sCopiesSpecialLabel     As String
    Dim sCopiesMergeLabel       As String
    Dim sSpecialStickReel       As String
    Dim sSpecialStickAlbag      As String
    Dim sSpecialPositionBox_1   As String
    Dim sSpecialPositionBox_2   As String

    '列印資料-2 : 欄位來自 Tbl_Ft_Label_Boxing_Info.
    Dim sBoxingSpecNoVer        As String
    Dim sVacuum                 As String
    Dim sHic                    As String
    Dim sDesiccant              As String
    Dim sAlbag                  As String

    '列印資料-3 : 其它欄位
    Dim sPrintLotNO             As String
    Dim sPrintMode              As String
    Dim sPackageCodePinCount    As String
    
    'Added by Jack on 2016/07/04 <Start>
    Dim sGrade                  As String
    Dim sBixType                As String
    'Added by Jack on 2016/07/04 <End>
    
    '----
    ' Init
    '----
    sProcID = "Prt_FT_Label_PACK_INFO"
    Call LogProcIn(msMODULE_ID, sProcID, oLogCtrl)
    sPackageCodePinCount = sPinCount & sPackageCode
    
    '----
    ' Condition Checking
    '----
    If bIsAutoPrint Then
        sPrintLotNO = sLotID
        sPrintMode = "Auto"
    Else
        sPrintLotNO = "NA" 'MK320 San Lee 要求帶NA, 不要空白(ps : SAP拉回, MES 沒有這筆Lot)
        sPrintMode = "Manual"
    End If

    'Check-Item-1 :
    sTable = gsCAT_TBL_FT_LABEL_SPEC_INFO

    sColumn = gsCAT_TFLSI_LABELSPECNOVER & ", " & gsCAT_TFLSI_SPECIALLABELSIZE & ", " & gsCAT_TFLSI_SERIAL & ", " & _
        gsCAT_TFLSI_COPIES_STD_LABEL & ", " & gsCAT_TFLSI_COPIES_SPECIAL_LABEL & ", " & gsCAT_TFLSI_COPIES_MERGE_LABEL & ", " & _
        gsCAT_TFLSI_MERGE_LABEL_SPEC & ", " & gsCAT_TFLSI_SPECIAL_STICK_REEL & ", " & gsCAT_TFLSI_SPECIAL_STICK_ALBAG & ", " & _
        gsCAT_TFLSI_SPECIAL_POSITION_BOX_1 & ", " & gsCAT_TFLSI_SPECIAL_POSITION_BOX_2 & " "

        sWhere = gsCAT_TFLSI_LABELSPECNO & "='" & sLabelSpecNo & "' and " & _
        gsCAT_TFLSI_CARRIERTYPE & "='" & sCarrierType & "' and " & _
        gsCAT_TFLSI_DELETEFLAG & "='N' "
    
    sSQL = "SELECT " & sColumn & " FROM " & sTable & " WHERE " & sWhere
    
    Set colRaws = moProRawSql.QueryDatabase(sSQL)
    If colRaws Is Nothing Then
        Call RaiseError(glERR_INVALIDOBJECT, _
            FormatErrorText(gsETX_INVALIDOBJECT, "Collection"))
    End If
    
    If colRaws.Count = 0 Then
        'Added by Jack on 2016/07/04 <Start>
        '若AutoMode的sLabelSpecNo串不到資料, 改用default的LabelSpecNo.
        If bIsAutoPrint Then
            sTable = gsCAT_TBL_IPN_MASTER

            sColumn = gsCAT_TIM_GRADE & ", " & gsCAT_TIM_BIZTYPE

            sWhere = gsCAT_TIM_IPN & "='" & sIPN & "' and " & _
                     gsCAT_TIM_DELETE_FLAG & "='N' "
    
            sSQL = "SELECT " & sColumn & " FROM " & sTable & " WHERE " & sWhere
            Set colRaws = moProRawSql.QueryDatabase(sSQL)
            If colRaws Is Nothing Then
                Call RaiseError(glERR_INVALIDOBJECT, _
                    FormatErrorText(gsETX_INVALIDOBJECT, "Collection"))
            End If
            If colRaws.Count <> 0 Then
                sGrade = colRaws.Item(1).Item(gsCAT_TIM_GRADE)
                sBixType = colRaws.Item(1).Item(gsCAT_TIM_BIZTYPE)
                If sBrand = "KH" And colRaws.Item(1).Item(gsCAT_TIM_GRADE) = "Y" And _
                    sPackageCode <> "H" And sPackageCode <> "W" Then
                    sLabelSpecNo = "6130K-0807.1"
                    ElseIf sBrand = "KH" And colRaws.Item(1).Item(gsCAT_TIM_BIZTYPE) = "D" And _
                        (sPackageCode = "H" Or sPackageCode = "W") Then
                    sLabelSpecNo = "6130K-0807.1"
                ElseIf sBrand = "KH" And sPackageCode <> "H" And sPackageCode <> "W" Then
                    sLabelSpecNo = "6130K-0807"
                Else
                    sLabelSpecNo = "6130-0807"
                End If
            End If
            sTable = gsCAT_TBL_FT_LABEL_SPEC_INFO

            sColumn = gsCAT_TFLSI_LABELSPECNOVER & ", " & gsCAT_TFLSI_SPECIALLABELSIZE & ", " & gsCAT_TFLSI_SERIAL & ", " & _
                gsCAT_TFLSI_COPIES_STD_LABEL & ", " & gsCAT_TFLSI_COPIES_SPECIAL_LABEL & ", " & gsCAT_TFLSI_COPIES_MERGE_LABEL & ", " & _
                gsCAT_TFLSI_MERGE_LABEL_SPEC & ", " & gsCAT_TFLSI_SPECIAL_STICK_REEL & ", " & gsCAT_TFLSI_SPECIAL_STICK_ALBAG & ", " & _
                gsCAT_TFLSI_SPECIAL_POSITION_BOX_1 & ", " & gsCAT_TFLSI_SPECIAL_POSITION_BOX_2 & " "
        
                sWhere = gsCAT_TFLSI_LABELSPECNO & "='" & sLabelSpecNo & "' and " & _
                gsCAT_TFLSI_CARRIERTYPE & "='" & sCarrierType & "' and " & _
                gsCAT_TFLSI_DELETEFLAG & "='N' "
            
            sSQL = "SELECT " & sColumn & " FROM " & sTable & " WHERE " & sWhere
            
            Set colRaws = moProRawSql.QueryDatabase(sSQL)
            If colRaws Is Nothing Then
                Call RaiseError(glERR_INVALIDOBJECT, _
                    FormatErrorText(gsETX_INVALIDOBJECT, "Collection"))
            End If
            
            If colRaws.Count = 0 Then
                Call UtShowMsgBox("FT_Label_Info(Default) 資訊異常, 請通知主任確認CAT資訊正確性. " & vbNewLine & vbNewLine & _
                                "FT_Label_Info error, Please inform the Supervisor confirmed " & vbNewLine & _
                                "that infomation of CAT system is correct. " & vbNewLine & vbNewLine & _
                                "(CAT:Tbl_Ft_Label_Spec_Info table) ")
                GoTo ExitHandler
            Else
                sLabelSpecNoVer = colRaws.Item(1).Item(gsCAT_TFLSI_LABELSPECNOVER)
                sSpecialLabelSize = colRaws.Item(1).Item(gsCAT_TFLSI_SPECIALLABELSIZE)
                sSerial = colRaws.Item(1).Item(gsCAT_TFLSI_SERIAL)
                sMergeLabelSpec = colRaws.Item(1).Item(gsCAT_TFLSI_MERGE_LABEL_SPEC)
                sCopiesStdLabel = colRaws.Item(1).Item(gsCAT_TFLSI_COPIES_STD_LABEL)
                sCopiesSpecialLabel = colRaws.Item(1).Item(gsCAT_TFLSI_COPIES_SPECIAL_LABEL)
                sCopiesMergeLabel = colRaws.Item(1).Item(gsCAT_TFLSI_COPIES_MERGE_LABEL)
                sSpecialStickReel = colRaws.Item(1).Item(gsCAT_TFLSI_SPECIAL_STICK_REEL)
                sSpecialStickAlbag = colRaws.Item(1).Item(gsCAT_TFLSI_SPECIAL_STICK_ALBAG)
                sSpecialPositionBox_1 = colRaws.Item(1).Item(gsCAT_TFLSI_SPECIAL_POSITION_BOX_1)
                sSpecialPositionBox_2 = colRaws.Item(1).Item(gsCAT_TFLSI_SPECIAL_POSITION_BOX_2)
            End If
            'Added by Jack on 2016/07/04 <End>
        Else
            Call UtShowMsgBox("FT_Label_Info 資訊異常, 請通知主任確認CAT資訊正確性. " & vbNewLine & vbNewLine & _
                "FT_Label_Info error, Please inform the Supervisor confirmed " & vbNewLine & _
                "that infomation of CAT system is correct. " & vbNewLine & vbNewLine & _
                "(CAT:Tbl_Ft_Label_Spec_Info table) ")
            GoTo ExitHandler
        End If
    Else
        sLabelSpecNoVer = colRaws.Item(1).Item(gsCAT_TFLSI_LABELSPECNOVER)
        sSpecialLabelSize = colRaws.Item(1).Item(gsCAT_TFLSI_SPECIALLABELSIZE)
        sSerial = colRaws.Item(1).Item(gsCAT_TFLSI_SERIAL)
        sMergeLabelSpec = colRaws.Item(1).Item(gsCAT_TFLSI_MERGE_LABEL_SPEC)
        sCopiesStdLabel = colRaws.Item(1).Item(gsCAT_TFLSI_COPIES_STD_LABEL)
        sCopiesSpecialLabel = colRaws.Item(1).Item(gsCAT_TFLSI_COPIES_SPECIAL_LABEL)
        sCopiesMergeLabel = colRaws.Item(1).Item(gsCAT_TFLSI_COPIES_MERGE_LABEL)
        sSpecialStickReel = colRaws.Item(1).Item(gsCAT_TFLSI_SPECIAL_STICK_REEL)
        sSpecialStickAlbag = colRaws.Item(1).Item(gsCAT_TFLSI_SPECIAL_STICK_ALBAG)
        sSpecialPositionBox_1 = colRaws.Item(1).Item(gsCAT_TFLSI_SPECIAL_POSITION_BOX_1)
        sSpecialPositionBox_2 = colRaws.Item(1).Item(gsCAT_TFLSI_SPECIAL_POSITION_BOX_2)
    End If

    'Check-Item-2 :
    sTable = gsCAT_TBL_FT_LABEL_BOXING_INFO

    sColumn = gsCAT_TFLBI_BOXINGSPECNOVER & ", " & gsCAT_TFLBI_VACUUM & ", " & gsCAT_TFLBI_HIC & ", " & _
        gsCAT_TFLBI_DESICCANT & ", " & gsCAT_TFLBI_ALBAG & " "

        sWhere = gsCAT_TFLBI_CARRIERTYPE & "='" & sCarrierType & "' and " & _
        gsCAT_TFLBI_BOXINGSPECNO & "='" & sBoxingSpecNo & "' and " & _
        gsCAT_TFLBI_BRAND & "='" & sBrand & "' and " & _
        gsCAT_TFLBI_PINCOUNT & "='" & sPinCount & "' and " & _
        gsCAT_TFLBI_PACKAGECODE & "='" & sPackageCode & "' and " & _
        gsCAT_TFLBI_DELETEFLAG & "='N' "
    
    sSQL = "SELECT " & sColumn & " FROM " & sTable & " WHERE " & sWhere
    
    Set colRaws = moProRawSql.QueryDatabase(sSQL)
    If colRaws Is Nothing Then
        Call RaiseError(glERR_INVALIDOBJECT, _
            FormatErrorText(gsETX_INVALIDOBJECT, "Collection"))
    End If
    
    If colRaws.Count = 0 Then
        'Auto mode : Package +Pincount 先以Lot’s IPN值為主, 若串不到資料, 再改為ALL + ALL 去串.
        If bIsAutoPrint Then
            sTable = gsCAT_TBL_FT_LABEL_BOXING_INFO
        
            sColumn = gsCAT_TFLBI_BOXINGSPECNOVER & ", " & gsCAT_TFLBI_VACUUM & ", " & gsCAT_TFLBI_HIC & ", " & _
                gsCAT_TFLBI_DESICCANT & ", " & gsCAT_TFLBI_ALBAG & " "
        
                sWhere = gsCAT_TFLBI_CARRIERTYPE & "='" & sCarrierType & "' and " & _
                gsCAT_TFLBI_BOXINGSPECNO & "='" & sBoxingSpecNo & "' and " & _
                gsCAT_TFLBI_BRAND & "='" & sBrand & "' and " & _
                gsCAT_TFLBI_PINCOUNT & "='ALL' and " & _
                gsCAT_TFLBI_PACKAGECODE & "='ALL' and " & _
                gsCAT_TFLBI_DELETEFLAG & "='N' "
            
            sSQL = "SELECT " & sColumn & " FROM " & sTable & " WHERE " & sWhere
            
            Set colRaws = moProRawSql.QueryDatabase(sSQL)
            If colRaws Is Nothing Then
                Call RaiseError(glERR_INVALIDOBJECT, _
                    FormatErrorText(gsETX_INVALIDOBJECT, "Collection"))
            End If
            
            If colRaws.Count = 0 Then
                Call UtShowMsgBox("FT_Boxing_Info 資訊異常, 請通知主任確認CAT資訊正確性. " & vbNewLine & vbNewLine & _
                    "FT_Boxing_Info error, Please inform the Supervisor confirmed " & vbNewLine & _
                    "that infomation of CAT system is correct. " & vbNewLine & vbNewLine & _
                    "(CAT:Tbl_Ft_Boxing_Spec_Info table) ")
                GoTo ExitHandler
            Else
                sPackageCodePinCount = "ALL" 'user要求只印一個ALL即可.
                
                sBoxingSpecNoVer = colRaws.Item(1).Item(gsCAT_TFLBI_BOXINGSPECNOVER)
                sVacuum = colRaws.Item(1).Item(gsCAT_TFLBI_VACUUM)
                sHic = colRaws.Item(1).Item(gsCAT_TFLBI_HIC)
                sDesiccant = colRaws.Item(1).Item(gsCAT_TFLBI_DESICCANT)
                sAlbag = colRaws.Item(1).Item(gsCAT_TFLBI_ALBAG)
            End If
        Else
            Call UtShowMsgBox("FT_Boxing_Info 資訊異常, 請通知主任確認CAT資訊正確性. " & vbNewLine & vbNewLine & _
                "FT_Boxing_Info error, Please inform the Supervisor confirmed " & vbNewLine & _
                "that infomation of CAT system is correct. " & vbNewLine & vbNewLine & _
                "(CAT:Tbl_Ft_Boxing_Spec_Info table) ")
            GoTo ExitHandler
        End If
    Else
        If sPinCount = "ALL" And sPackageCode = "ALL" Then
            sPackageCodePinCount = "ALL" 'user要求只印一個ALL即可.
        End If
        
        sBoxingSpecNoVer = colRaws.Item(1).Item(gsCAT_TFLBI_BOXINGSPECNOVER)
        sVacuum = colRaws.Item(1).Item(gsCAT_TFLBI_VACUUM)
        sHic = colRaws.Item(1).Item(gsCAT_TFLBI_HIC)
        sDesiccant = colRaws.Item(1).Item(gsCAT_TFLBI_DESICCANT)
        sAlbag = colRaws.Item(1).Item(gsCAT_TFLBI_ALBAG)
    End If

    '----
    ' Action
    '----
    On Error Resume Next
       
    '******
    'Print Label Info
    '******

    sMsg = sMsg & "^XA"
    sMsg = sMsg & "^MCY"
    sMsg = sMsg & "^XZ"
    sMsg = sMsg & "^XA"
    sMsg = sMsg & "^FWN^CFD,24^LH0,0"
    sMsg = sMsg & "^CI0^PR2^MNY^MTT^MMT^MD0^PON^PMN^LRN"
    sMsg = sMsg & "^XZ"
    sMsg = sMsg & "^XA"
    sMsg = sMsg & "^DFR:TEMP_FMT.ZPL"
    sMsg = sMsg & "^LRN"
    sMsg = sMsg & "^XZ"
    sMsg = sMsg & "^XA"
    sMsg = sMsg & "^XFR:TEMP_FMT.ZPL"
    sMsg = sMsg & "^FO15,17^GB1167,112,2,B,0^FS"
    sMsg = sMsg & "^FO15,71^GB1165,2,2,B,0^FS"
    sMsg = sMsg & "^FO191,18^GB2,110,2,B,0^FS"
    sMsg = sMsg & "^FO247,16^GB2,110,2,B,0^FS"
    sMsg = sMsg & "^FO799,18^GB2,110,2,B,0^FS"
    sMsg = sMsg & "^FO709,18^GB2,110,2,B,0^FS"
    sMsg = sMsg & "^FO988,15^GB2,110,2,B,0^FS"
    sMsg = sMsg & "^A0N,23,24^FO40,33^FDLabel Spec.^FS"
    sMsg = sMsg & "^A0N,23,24^FO425,33^FDLabel Size^FS"
    sMsg = sMsg & "^A0N,23,24^FO199,33^FDVer.^FS"
    sMsg = sMsg & "^A0N,23,24^FO723,33^FDSerial^FS"
    sMsg = sMsg & "^A0N,23,24^FO858,33^FDSource^FS"
    sMsg = sMsg & "^A0N,23,24^FO1047,33^FDLotNo.^FS"
    sMsg = sMsg & "^A0N,28,28^FO255,89^FD" & sSpecialLabelSize & "^FS"
    sMsg = sMsg & "^A0N,28,28^FO199,89^FD" & sLabelSpecNoVer & "^FS"
    sMsg = sMsg & "^A0N,28,28^FO20,89^FD" & sLabelSpecNo & "^FS"
    sMsg = sMsg & "^A0N,28,28^FO998,89^FD" & sPrintLotNO & "^FS" '注意!! 可能為 "NA" <-- Manual mode
    sMsg = sMsg & "^A0N,28,28^FO859,89^FD" & sPrintMode & "^FS" '注意!! Auto / Manual.
    sMsg = sMsg & "^A0N,28,28^FO741,89^FD" & sSerial & "^FS"
    sMsg = sMsg & "^FO13,315^GB1167,112,2,B,0^FS"
    sMsg = sMsg & "^FO15,368^GB1165,2,2,B,0^FS"
    sMsg = sMsg & "^FO519,315^GB2,110,2,B,0^FS"
    sMsg = sMsg & "^FO799,314^GB2,110,2,B,0^FS"
    sMsg = sMsg & "^FO398,314^GB2,110,2,B,0^FS"
    sMsg = sMsg & "^FO247,314^GB2,110,2,B,0^FS"
    sMsg = sMsg & "^FO191,316^GB2,110,2,B,0^FS"
    sMsg = sMsg & "^FO988,312^GB2,110,2,B,0^FS"
    sMsg = sMsg & "^A0N,23,24^FO36,333^FDBoxing Spec.^FS"
    sMsg = sMsg & "^A0N,23,24^FO198,333^FDVer.^FS"
    sMsg = sMsg & "^A0N,23,24^FO294,333^FDBrand^FS"
    sMsg = sMsg & "^A0N,23,24^FO419,333^FDPackage^FS"
    sMsg = sMsg & "^A0N,23,24^FO594,333^FDCarrier Type^FS"
    sMsg = sMsg & "^A0N,23,24^FO808,333^FDVacuum^FS"
    sMsg = sMsg & "^A0N,23,24^FO921,333^FDHIC^FS"
    sMsg = sMsg & "^A0N,23,24^FO992,333^FDDesiccant^FS"
    sMsg = sMsg & "^A0N,23,24^FO1093,333^FDAl. bag^FS"
    sMsg = sMsg & "^FO895,314^GB2,110,2,B,0^FS"
    sMsg = sMsg & "^FO1087,314^GB2,110,2,B,0^FS"
    sMsg = sMsg & "^A0N,28,28^FO19,387^FD" & sBoxingSpecNo & "^FS"
    sMsg = sMsg & "^A0N,28,28^FO199,388^FD" & sBoxingSpecNoVer & "^FS"
    sMsg = sMsg & "^A0N,28,28^FO296,388^FD" & sBrand & "^FS"
    sMsg = sMsg & "^A0N,28,28^FO411,388^FD" & sPackageCodePinCount & "^FS" '注意 : 可能為 "ALL"
    sMsg = sMsg & "^A0N,28,28^FO527,388^FD" & sCarrierType & "^FS"
    sMsg = sMsg & "^A0N,28,28^FO832,388^FD" & sVacuum & "^FS"
    sMsg = sMsg & "^A0N,28,28^FO934,388^FD" & sHic & "^FS"
    sMsg = sMsg & "^A0N,28,28^FO1032,388^FD" & sDesiccant & "^FS"
    sMsg = sMsg & "^A0N,28,28^FO1108,387^FD" & sAlbag & "^FS"
    sMsg = sMsg & "^FO13,136^GB1167,168,2,B,0^FS"
    sMsg = sMsg & "^FO13,190^GB1165,2,2,B,0^FS"
    sMsg = sMsg & "^FO14,248^GB1165,2,2,B,0^FS"
    sMsg = sMsg & "^FO190,190^GB2,110,2,B,0^FS"
    sMsg = sMsg & "^FO400,190^GB2,110,2,B,0^FS"
    sMsg = sMsg & "^FO520,140^GB2,162,2,B,0^FS"
    sMsg = sMsg & "^FO802,140^GB2,162,2,B,0^FS"
    sMsg = sMsg & "^FO988,137^GB2,162,2,B,0^FS"
    sMsg = sMsg & "^A0N,23,24^FO214,153^FDLabel  Copies^FS"
    sMsg = sMsg & "^A0N,23,24^FO69,209^FDSTD^FS"
    sMsg = sMsg & "^A0N,23,24^FO251,209^FDSpecial^FS"
    sMsg = sMsg & "^A0N,23,24^FO422,209^FDMerge^FS"
    sMsg = sMsg & "^A0N,23,24^FO578,153^FDMerge Label^FS"
    sMsg = sMsg & "^A0N,23,24^FO621,209^FDSpec.^FS"
    sMsg = sMsg & "^A0N,23,24^FO825,153^FDSpecial Stick^FS"
    sMsg = sMsg & "^A0N,23,24^FO1018,153^FDSpecial Pos.^FS"
    sMsg = sMsg & "^A0N,23,24^FO822,209^FDReel^FS"
    sMsg = sMsg & "^FO895,192^GB2,110,2,B,0^FS"
    sMsg = sMsg & "^FO1088,193^GB2,110,2,B,0^FS"
    sMsg = sMsg & "^A0N,23,24^FO903,209^FDAl. Bag^FS"
    sMsg = sMsg & "^A0N,23,24^FO1007,209^FDBox_1^FS"
    sMsg = sMsg & "^A0N,23,24^FO1098,208^FDBox_2^FS"
    sMsg = sMsg & "^A0N,28,28^FO78,265^FD" & sCopiesStdLabel & "^FS"
    sMsg = sMsg & "^A0N,28,28^FO270,265^FD" & sCopiesSpecialLabel & "^FS"
    sMsg = sMsg & "^A0N,28,28^FO442,265^FD" & sCopiesMergeLabel & "^FS"
    sMsg = sMsg & "^A0N,28,28^FO535,265^FD" & sMergeLabelSpec & "^FS"
    sMsg = sMsg & "^A0N,28,28^FO837,265^FD" & sSpecialStickReel & "^FS"
    sMsg = sMsg & "^A0N,28,28^FO926,265^FD" & sSpecialStickAlbag & "^FS"
    sMsg = sMsg & "^A0N,28,28^FO1000,265^FD" & sSpecialPositionBox_1 & "^FS"
    sMsg = sMsg & "^A0N,28,28^FO1094,265^FD" & sSpecialPositionBox_2 & "^FS"
    sMsg = sMsg & "^PQ1,0,1,Y"
    sMsg = sMsg & "^XZ"
    sMsg = sMsg & "^XA"
    sMsg = sMsg & "^IDR:TEMP_FMT.ZPL"
    sMsg = sMsg & "^XZ"
    '使用的欄位變數:
    'sLotID, sLabelSpecNo, sCarrierType, sBoxingSpecNo, sBrand, sPinCount, sPackageCode
    'sLabelSpecNoVer , sSpecialLabelSize, , sMergeLabelSpec, sCopiesStdLabel
    'sCopiesSpecialLabel , sCopiesMergeLabel, sSpecialStickReel, sSpecialStickAlbag
    'sSpecialPositionBox_1 , sSpecialPositionBox_2
    'sBoxingSpecNoVer, sVacuum, sHic, sDesiccant, sAlbag

    '******
    ' End of Label Info
    '******

    '******
    'Send message to MES Server MailBox.
    '******
    If InStr(1, sPrinterServer, "@") <> 0 Then
        sPrinterServer = VBA.Left(sPrinterServer, (InStr(1, sPrinterServer, "@")) - 1)
    End If
        
    If Not moCwMbx.Send("MBX_" & sPrinterServer, sMsg) Then
        Call moCwMbx.GetError(lErrCode, sErrMsg)
        Call RaiseError(lErrCode, sErrMsg)
    End If
    '******
    '----
    ' Done
    '----
ExitHandler:
    ' NOTE 1:
    ' MUST CALL GetErrInfo() here first before another action
    Call GetErrInfo(msMODULE_ID, sProcID, typErrInfo, Erl)
    Call LogProcOut(msMODULE_ID, sProcID, typErrInfo, oLogCtrl)
    On Error Resume Next

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
        Call HandleError(False, typErrInfo, , oLogCtrl, True)
    End If
End Sub


'================================================================================
' function: GetProdType()
'--------------------------------------------------------------------------------
' Description:  <Type your function description here...>
'--------------------------------------------------------------------------------
' Author:       Jack Hsieh, MXIC 2016/06/27 for JC201600108 (新增產品組合Rule).
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
' [REV 01] Jack Hsieh, MXIC, 2016/08/30 for JC201600173.
' 1) 加傳參數 : sCriteiraString (In), sCriteiraData (Out)
' 2) 呼叫 DB : GetProdLevel_Criteira(ipn, 'IpnRelease')
'    sCriteiraString = 'IpnRelease' (傳入), DB function 回傳 : "KTD,Y"
'    VB function 回傳 : GetProdType = 'KTD' ; sCriteiraData = 'Y'. <-- 以 "," 拆解.
' [REV 02] <AuthorName>, <CompanyName> <YYYY/MM/DD>
' 1) <Description goes here...>
'    <Line 2...>
'
'================================================================================
    Public Function GetProdType(ByVal sIPN As String, ByVal sCriteiraString As String, ByRef sCriteiraData As String, _
                                ByVal oLogCtrl As Object, ByVal moProRawSql As Object) As String
        On Error GoTo ExitHandler:
        Dim sProcID As String
        Dim typErrInfo As tErrInfo
        Dim sTable As String
        Dim sColumn As String
        Dim sWhere As String
        Dim sOther  As String
        Dim colRaws As Collection
        Dim oRaws As FwStrings
        Dim sSQL As String
        Dim sReturnData As String
        
        '----
        ' Init
        '----
1       sProcID = "GetProdType"
2       Call LogProcIn(msMODULE_ID, sProcID, oLogCtrl)
3       GetProdType = ""
        '--2016/08/31 Livan : 一切從嚴, 沒有值就預設為 'Y'.
        sCriteiraData = "Y"
        
        '----
        ' Condition Chking
        '----


        '----
        ' Action
        '----
        '    sTable = gsCAT_TBL_WS_PRODTYPE
        '
        '    sColumn = gsCAT_TWP_PRODLEVEL & "," & gsCAT_TWP_PRODTYPE
        '
        '    sWhere = gsCAT_TWP_PRODLEVEL & "=fun_get_waferlevel('" & sIPN & "', 'WaferLevel') " & _
        '        " and " & gsCAT_TWP_DELETEFLAG & " = 'N' "
        '
        '    sSQL = "SELECT " & sColumn & " FROM " & sTable & " WHERE " & sWhere
    
        '改用 資料庫的Function : GetProdType( ),日後 I/F & Web都可共用.
        
        'Modified by Jack on 2016/08/30 for JC201600173.
        '改用 資料庫的Function : GetProdType( )
'4       'sSQL = "select GetProdType('" & sIPN & "') as prodtype from dual "
4        sSQL = "select GetProdType('" & sIPN & "', '" & sCriteiraString & "') as prodtype from dual "
    
5       Set colRaws = moProRawSql.QueryDatabase(sSQL)
6       If colRaws Is Nothing Then
7           Call RaiseError(glERR_INVALIDOBJECT, _
                FormatErrorText(gsETX_INVALIDOBJECT, "Collection"))
        End If
    
8       If colRaws.Count > 0 Then
'9           GetProdType = colRaws.Item(1).Item(gsCAT_TWP_PRODTYPE)
9           sReturnData = colRaws.Item(1).Item(gsCAT_TWP_PRODTYPE)
            
            If InStr(sReturnData, ",") > 0 Then
                GetProdType = Trim(Mid(sReturnData, 1, InStr(sReturnData, ",") - 1))
                sCriteiraData = Trim(Mid(sReturnData, InStr(sReturnData, ",") + 1))
            End If
        End If
    
        '----
        ' Done
        '----

ExitHandler:
        ' NOTE 1:
        ' MUST CALL GetErrInfo() here first before another action
10      Call GetErrInfo(msMODULE_ID, sProcID, typErrInfo, Erl)
11      Call LogProcOut(msMODULE_ID, sProcID, typErrInfo, oLogCtrl)
        ' <Your cleaning up codes goes here...>
    
    
ErrorHandler:
12      If typErrInfo.lErrNumber Then
            ' NOTE 2:
            ' If you have custom handling of some Errors, please
            ' UN-REMARED the following Select Case block!
            ' Also, modify if neccessarily!!!
            '---- Start of Select Case Block ----
            Select Case typErrInfo.lErrNumber
            Case glERR_INVALIDOBJECT
                ' Retry code goes here...
            Case Else
13              typErrInfo.sUserText = "Fail to execute application, please call IT support!!" & vbCrLf & _
                    "程式執行失敗, 請洽IT人員處理"
            End Select
            '---- Start of Select Case Block ----
            On Error GoTo ExitHandler:
14          Call HandleError(True, typErrInfo, , oLogCtrl)
        End If
    End Function


'================================================================================
' Sub: Prt_WS_CP_SMALL_LABEL()
'--------------------------------------------------------------------------------
' Description:  <This subroutine will print out a label for FT-> SFG >
'--------------------------------------------------------------------------------
' Author:       <Sam Chan>, <CIT> 2018/08/13
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
'================================================================================
Public Sub Prt_WS_CP_SMALL_LABEL(ByVal sLotNo As String, _
                                ByVal sProdNo As String, _
                                ByVal sWQty As String, _
                                ByVal sCQty As String, _
                                ByVal sOwner As String, _
                                ByVal sBizType As String, _
                                ByVal sPrinterServer As String, _
                                ByVal moCwMbx As Object, _
                                ByVal oLogCtrl As Object, _
                                ByVal moProRawSql As Object, _
                                Optional ByVal sWaferID As String, _
                                Optional ByVal sFabLotID As String)


On Error GoTo ExitHandler:

Dim sMsg            As String
Dim sProcID         As String
Dim typErrInfo      As tErrInfo
Dim lErrCode        As Long
Dim sErrMsg         As String

Dim sSQL As String
Dim colRS As Collection
Dim colRS2 As Collection

Dim sHotLotFlag As String
Dim sErunTicNO As String
Dim sSapRwNo As String
Dim sMfgTicNo As String
Dim sWaferNo1 As String
Dim sWaferNo2 As String
Dim sWaferNo3 As String
Dim iIdx As Integer
Dim iIdx2 As Integer
Dim sWaferNO(25) As String
Dim vWaferNO As Variant
Dim sSortWaferNo As String
Dim bFound As Boolean

Dim sPrintFabLotId As String
Dim sCriteiraString   As String
Dim sCriteiraData     As String

Dim sPrintFlag As String
Dim sProdCode As String
Dim sWaferLevel As String

'Add by Weilun on 20220816 for ReqNo.BE#202200157 <start>
Dim sBomLevel         As String
Dim sFGIPN            As String
Dim sFusaIpn          As String
Dim sFuSa             As String
'Add by Weilun on 20220816 for ReqNo.BE#202200157 <end>

'----
' Init
'----
    sProcID = "Prt_WS_CP_SMALL_LABEL"
    Call LogProcIn(msMODULE_ID, sProcID, oLogCtrl)
    
    sCriteiraString = "" '此處只有列印, 不需傳入.
    sBizType = GetProdType(sProdNo, sCriteiraString, sCriteiraData, oLogCtrl, moProRawSql)
         
    'gsCAT_TIM_BOM_LEVEL, Add by Weilun on 20220816 for ReqNo.BE#202200157
    sSQL = "select fun_get_waferlevel(a." & gsCAT_TIM_IPN & ") as waferlevel " & _
           " ,a." & gsCAT_TIM_MASK_OPTION & " , a." & gsCAT_TIM_BE_OPTION & " " & _
           " ,a." & gsCAT_TIM_BOM_LEVEL & " " & _
           " from " & gsCAT_TBL_IPN_MASTER & " a " & _
           " where a." & gsCAT_TIM_IPN & "= '" & sProdNo & "' " & _
           " and a." & gsCAT_TIM_DELETE_FLAG & "= 'N' "
    Set colRS = moProRawSql.QueryDatabase(sSQL)
    If colRS.Count > 0 Then
        sWaferLevel = colRS.Item(1).Item("waferlevel")
        sProdCode = Left(sProdNo, 4) & colRS.Item(1).Item(gsCAT_TIM_MASK_OPTION)
        sBomLevel = colRS.Item(1).Item(gsCAT_TIM_BOM_LEVEL) 'Add by Weilun on 20220816 for ReqNo.BE#202200157
    End If
    If sProdCode = "" Then
        sProdCode = Left(sProdNo, 4)
    End If
    
    sHotLotFlag = ""
    sHotLotFlag = GetCpHot(sLotNo, sProdNo, moProRawSql, oLogCtrl)
    
    'gsCAT_TLI_FG_IPN, Add by Weilun on 20220816 for ReqNo.BE#202200157
    sSQL = " select A." & gsCAT_TLATT_HOTLOTFLAG & " ,B." & gsCAT_TLI_ERUNTICNO & _
           " ,B." & gsCAT_TLI_SAPRWNO & " ,B." & gsCAT_TLI_MFGTICNO & _
           " ,B." & gsCAT_TLI_WAFERID & _
           " ,B." & gsCAT_TLI_FABLOTID & _
           " ,B." & gsCAT_TLI_FG_IPN & _
           " FROM " & gsCAT_TBL_LOT_ATTRIBUTE & " A, " & gsCAT_TBL_LOT_INFO & " B " & _
           " WHERE A." & gsCAT_TLATT_LOTID & "= B." & gsCAT_TLI_LOT_ID & " " & _
           " AND  A." & gsCAT_TLATT_LOTID & "='" & sLotNo & "' "
    Set colRS = moProRawSql.QueryDatabase(sSQL)
    If colRS.Count > 0 Then
        sErunTicNO = colRS.Item(1).Item(gsCAT_TLI_ERUNTICNO)
        sSapRwNo = colRS.Item(1).Item(gsCAT_TLI_SAPRWNO)
        sMfgTicNo = colRS.Item(1).Item(gsCAT_TLI_MFGTICNO)
        sFGIPN = colRS.Item(1).Item(gsCAT_TLI_FG_IPN) 'Add by Weilun on 20220816 for ReqNo.BE#202200157
        
        sSQL = "select a." & gsCAT_TRR_PRINTFLAG & " from " & gsCAT_TBL_REWORK_REQ & " a " & _
                " where a." & gsCAT_TRR_SAPRWNO & " = '" & sSapRwNo & "' " & _
                " and a." & gsCAT_TRR_DELETEFLAG & " = 'N' "
        Set colRS2 = moProRawSql.QueryDatabase(sSQL)
        If colRS2.Count > 0 Then
            sPrintFlag = colRS2.Item(1).Item(gsCAT_TRR_PRINTFLAG)
        End If
        
        If sPrintFlag <> "Y" Then
            sSapRwNo = ""
        End If
        
        If sWaferID <> "" Then
            sWaferNo1 = Replace(sWaferID, ";", ",")
        Else
            sWaferNo1 = Replace(colRS.Item(1).Item(gsCAT_TLI_WAFERID), ";", ",")
        End If
        
        If sFabLotID <> "" Then
            sPrintFabLotId = sFabLotID
        Else
            sPrintFabLotId = colRS.Item(1).Item(gsCAT_TLI_FABLOTID)
        End If
        
        sSortWaferNo = SortWaferID(sWaferNo1)
        vWaferNO = Split(sWaferNo1, ",")
        sWaferNo1 = ""

        For iIdx = 1 To 25
            bFound = False
            For iIdx2 = LBound(vWaferNO) To UBound(vWaferNO)
                If Val(vWaferNO(iIdx2)) = iIdx Then
                    bFound = True
                    Exit For
                End If
            Next
            If bFound = True Then
                sWaferNo1 = sWaferNo1 & "," & Format(iIdx, "00")
            Else
                sWaferNo1 = sWaferNo1 & ",__"
            End If
        Next
        sWaferNo1 = Mid(sWaferNo1, 2)
        
        sWaferNo3 = Mid(sWaferNo1, 55) '19~25
        sWaferNo2 = Mid(sWaferNo1, 28, 26) '10~18
        sWaferNo1 = Mid(sWaferNo1, 1, 26)   '1~9
        
    Else
        '如果有外部傳WaferID 時優先使用
        If sWaferID <> "" Then
            sWaferNo1 = Replace(sWaferID, ";", ",")
        End If
        sSortWaferNo = SortWaferID(sWaferNo1)
        vWaferNO = Split(sWaferNo1, ",")
        sWaferNo1 = ""
        For iIdx = 1 To 25
            bFound = False
            For iIdx2 = LBound(vWaferNO) To UBound(vWaferNO)
                If Val(vWaferNO(iIdx2)) = iIdx Then
                    bFound = True
                    Exit For
                End If
            Next
            If bFound = True Then
                sWaferNo1 = sWaferNo1 & "," & Format(iIdx, "00")
            Else
                sWaferNo1 = sWaferNo1 & ",__"
            End If
        Next
        sWaferNo1 = Mid(sWaferNo1, 2)
        
        sWaferNo3 = Mid(sWaferNo1, 55) '19~25
        sWaferNo2 = Mid(sWaferNo1, 28, 26) '10~18
        sWaferNo1 = Mid(sWaferNo1, 1, 26)   '1~9
                        
        '如果有外部傳FabLotID時優先使用
        If sFabLotID <> "" Then
            sPrintFabLotId = sFabLotID
        End If
    End If
    
    'Add by Weilun on 20220816 for ReqNo.BE#202200157 <start>
    sFusaIpn = sProdNo '查詢Fusa預設為Lot的IPN
    If sBomLevel = "DB" And _
       (sWaferLevel = "KGD-AEB" Or sWaferLevel = "MXO1") And _
       sFGIPN <> "" Then
       
        sFusaIpn = sFGIPN
    End If
      
    sFuSa = ""
    sSQL = " select " & gsCAT_TIM_FUSA & " " & _
             " from " & gsCAT_TBL_IPN_MASTER & " " & _
            " where " & gsCAT_TIM_IPN & " = '" & sFGIPN & "' "
    Set colRS = moProRawSql.QueryDatabase(sSQL)
    If colRS.Count > 0 Then
        sFuSa = colRS.Item(1).Item(gsCAT_TIM_FUSA)
    End If
    'Add by Weilun on 20220816 for ReqNo.BE#202200157 <end>
'----
' Condition Checking
'----

'----
' Action
'----
    On Error Resume Next
    '******
    'Print Label Info
    '******
    
        sMsg = sMsg & "^XA"
        sMsg = sMsg & "^PRC"
        sMsg = sMsg & "^LH0,0^FS"
        sMsg = sMsg & "^LL280"
        sMsg = sMsg & "^MD0"
        sMsg = sMsg & "^LH0,0^FS"
        sMsg = sMsg & "^FO15,15^AFN,10,10^CI13^FDLotNo:^FS"
        sMsg = sMsg & "^FO130,15^A0N,37,33^CI13^FD" & sLotNo & "^FS"
        sMsg = sMsg & "^FO15,50^BY2,3.0^BCN,28,N,Y,N^FR^FD>:" & sLotNo & "^FS"
        sMsg = sMsg & "^FO15,105^AFN,10,10^CI13^FDFabLotId:^FS"
        sMsg = sMsg & "^FO170,105^A0N,37,33^CI13^FD" & sPrintFabLotId & "^FS"
        'Modify by Sam Start on 20201203 for Project CP 測前併批 ,Prodcode改成prodbody
'        sMsg = sMsg & "^FO15,150^AFN,10,10^CI13^FDProdCode:^FS"
'        sMsg = sMsg & "^FO170,150^A0N,37,33^CI13^FD" & sProdCode & "^FS"
        sMsg = sMsg & "^FO15,150^AFN,10,10^CI13^FDProdBody:^FS"
        sMsg = sMsg & "^FO170,150^A0N,37,33^CI13^FD" & Left(sProdCode, 4) & "^FS"
        'Modify by Sam End on 20201203 for Project CP 測前併批 ,Prodcode改成prodbody
        sMsg = sMsg & "^FO290,150^AFN,10,10^CI13^FDHot:^FS"
        sMsg = sMsg & "^FO360,150^CF0N,37,33^CI13^FD" & sHotLotFlag & "^FS"
        sMsg = sMsg & "^FO290,195^AFN,10,10^CI13^FDQTY:^FS"
        sMsg = sMsg & "^FO360,195^A0N,37,33^CI13^FR^FD" & sWQty & " PCS^FS"
        sMsg = sMsg & "^FO15,195^AFN,10,10^CI13^FDProdType:^FS"
        sMsg = sMsg & "^FO170,195^A0N,37,33^CI13^FR^FD" & sBizType & "^FS"
        sMsg = sMsg & "^FO15,240^AFN,10,10^CI13^FDErunTicNo:^FS"
        sMsg = sMsg & "^FO170,240^A0N,37,33^CI13^FD" & sErunTicNO & "^FS"
        'Add by Weilun on 20220816 for ReqNo.BE#202200157 <start>
        sMsg = sMsg & "^FO355,240^AFN,10,10^CI13^FDFuSa:^FS"
        sMsg = sMsg & "^FO435,240^A0N,37,33^CI13^FD" & sFuSa & "^FS"
        'Add by Weilun on 20220816 for ReqNo.BE#202200157 <end>
        sMsg = sMsg & "^FO15,285^AFN,10,10^CI13^FDSapRwNo:^FS"
        sMsg = sMsg & "^FO170,285^A0N,37,33^CI13^FD" & sSapRwNo & "^FS"
        sMsg = sMsg & "^FO15,330^AFN,10,10^CI13^FDID#:^FS"
        sMsg = sMsg & "^FO80,330^A0N,37,33^CI13^FD" & sWaferNo1 & "^FS"
        sMsg = sMsg & "^FO80,380^A0N,37,33^CI13^FD" & sWaferNo2 & "^FS"
        sMsg = sMsg & "^FO80,430^A0N,37,33^CI13^FD" & sWaferNo3 & "^FS"
        sMsg = sMsg & "^PQ1,1,1,Y"
        sMsg = sMsg & "^XZ"
    
    '******
    ' End of Label Info
    '******

    '******
    'Send message to TDS
    '******
        
        If InStr(1, sPrinterServer, "@") <> 0 Then
            sPrinterServer = VBA.Left(sPrinterServer, (InStr(1, sPrinterServer, "@")) - 1)
        End If
        
        If Not moCwMbx.Send("MBX_" & sPrinterServer, sMsg) Then
            Call moCwMbx.GetError(lErrCode, sErrMsg)
            Call RaiseError(lErrCode, sErrMsg)
        End If
    '******

'----
' Done
'----
ExitHandler:
    ' NOTE 1:
    ' MUST CALL GetErrInfo() here first before another action
    Call GetErrInfo(msMODULE_ID, sProcID, typErrInfo, Erl)
    Call LogProcOut(msMODULE_ID, sProcID, typErrInfo, oLogCtrl)
    On Error Resume Next

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
        Call HandleError(False, typErrInfo, , oLogCtrl, True)
    End If
End Sub


'================================================================================
' Function: GetCpHot()
'--------------------------------------------------------------------------------
' Description:  <Type your function description here...>
'--------------------------------------------------------------------------------
' Author:       Sam Chen, CIT 2018/09/12
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
Public Function GetCpHot(ByVal sLotID As String, ByVal sIPN As String, oProRawSql As Object, ByVal oLogCtrl As Object) As String
    On Error GoTo ExitHandler:
    Dim sProcID As String
    Dim typErrInfo As tErrInfo
    Dim sSQL As String, colRS As Collection
    Dim sMainProd As String
    
    '----
    ' Init
    '----
    sProcID = "GetCpHot"
    Call LogProcIn(msMODULE_ID, sProcID, oLogCtrl) '"Entering Function...", oLogCtrl, glLOG_PROC, msMODULE_ID, sProcID)
    GetCpHot = ""
    '----
    ' Condition Checking
    '----
    ' <Put your condition checking codes here>...

    '----
    ' Action
    '----
       
    sSQL = "select " & gsCAT_TWTL_WSOUT & "  " & _
            " from " & gsCAT_TBL_WS_TRACKLOT & _
            " where " & gsCAT_TWTL_LOTID & " = '" & sLotID & "'" & _
            " and " & gsCAT_TWTL_WSOUT & " is not null "
            
    Set colRS = oProRawSql.QueryDatabase(sSQL)
    If colRS.Count > 0 Then
        GetCpHot = "Y"
    End If
    
    If GetCpHot = "" Then
        sMainProd = GetMainProd(sIPN, oProRawSql, oLogCtrl)
        If sMainProd = "Y" Then
            GetCpHot = "Y"
        End If
    End If
   
    '----
    ' Done
    '----

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
        Case glERR_INVALIDOBJECT
            ' Retry code goes here...
        Case Else
            typErrInfo.sUserText = "Fail to execute application, please call IT support!!" & vbCrLf & _
                "程式執行失敗, 請洽IT人員處理"
        End Select
        '---- Start of Select Case Block ----
        On Error GoTo ExitHandler:
        Call HandleError(False, typErrInfo, , oLogCtrl)
    End If
End Function



'================================================================================
' Function: GetMainProd()
'--------------------------------------------------------------------------------
' Description:  <Type your function description here...>
'--------------------------------------------------------------------------------
' Author:       Sam Chen, CIT 2018/08/17
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
Public Function GetMainProd(ByVal sIPN As String, oProRawSql As Object, ByVal oLogCtrl As Object) As String
    On Error GoTo ExitHandler:
    Dim sProcID As String
    Dim typErrInfo As tErrInfo
    Dim sSQL As String, colRS As Collection
    Dim sWaferLevel As String
    Dim sProdCode As String
    
    '----
    ' Init
    '----
    sProcID = "GetMainProd"
    Call LogProcIn(msMODULE_ID, sProcID, oLogCtrl) '"Entering Function...", oLogCtrl, glLOG_PROC, msMODULE_ID, sProcID)
    GetMainProd = ""
    '----
    ' Condition Checking
    '----
    ' <Put your condition checking codes here>...

    '----
    ' Action
    '----
   
        sWaferLevel = ""
        sProdCode = ""
        
        sSQL = "select fun_get_waferlevel(a." & gsCAT_TIM_IPN & ") as waferlevel " & _
                       " ,a." & gsCAT_TIM_MASK_OPTION & " , a." & gsCAT_TIM_BE_OPTION & " " & _
                       " from " & gsCAT_TBL_IPN_MASTER & " a " & _
                       " where a." & gsCAT_TIM_IPN & "= '" & sIPN & "' " & _
                       " and a." & gsCAT_TIM_DELETE_FLAG & "= 'N' "
        Set colRS = oProRawSql.QueryDatabase(sSQL)
        If colRS.Count > 0 Then
            sWaferLevel = colRS.Item(1).Item("waferlevel")
            sProdCode = Left(sIPN, 4) & colRS.Item(1).Item(gsCAT_TIM_MASK_OPTION)
        End If
                
        sSQL = "select a." & gsCAT_TPMP_PRODCODE & "   " & _
                       " from " & gsCAT_TBL_PC_MAIN_PROD & " a " & _
                       " where '" & sProdCode & "' like replace(replace(a." & gsCAT_TPMP_PRODCODE & ",'%','_'),'*','%')   " & _
                       " and '" & sIPN & "' like replace(replace(a." & gsCAT_TPMP_IPN & ",'%','_'),'*','%')  " & _
                       " and  '" & sWaferLevel & "' like  replace(a." & gsCAT_TPMP_WAFERLEVEL & ",'*','%')  " & _
                       " and a." & gsCAT_TPMP_DELETEFLAG & "= 'N' "
        
        Set colRS = oProRawSql.QueryDatabase(sSQL)
        If colRS.Count > 0 Then
            GetMainProd = "Y"
        End If
   
    '----
    ' Done
    '----

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
        Case glERR_INVALIDOBJECT
            ' Retry code goes here...
        Case Else
            typErrInfo.sUserText = "Fail to execute application, please call IT support!!" & vbCrLf & _
                "程式執行失敗, 請洽IT人員處理"
        End Select
        '---- Start of Select Case Block ----
        On Error GoTo ExitHandler:
        Call HandleError(False, typErrInfo, , oLogCtrl)
    End If
End Function


'================================================================================
' function: GetLabelSpecByStage()
'--------------------------------------------------------------------------------
' Description:  <Type your function description here...>
'--------------------------------------------------------------------------------
' Author:       Sam Chen, 2020-12-02, Project CP 測前併批
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
Public Function GetLabelSpecByStage(ByVal moProRawSql As Object, ByVal moAppLog As Object, ByVal sStage As String) As Collection
On Error GoTo ExitHandler:
Dim sProcID As String
Dim typErrInfo As tErrInfo

Dim colRS As Collection
Dim sSQL As String

'----
' Init
'----
    sProcID = "GetLabelSpecByStage"
    Call LogProcIn(msMODULE_ID, sProcID, moAppLog)
    
    Set GetLabelSpecByStage = New Collection
    
'----
' Condition Chking
'----


'----
' Action
'----
    sSQL = "select distinct a." & gsCAT_TLS_LABEL_SPECNO & " from " & gsCAT_TBL_LABEL_SPEC & " a " & _
          " where " & gsCAT_TLS_STAGE & " = '" & sStage & "' " & _
          " and a." & gsCAT_TLS_DELETE_FLAG & " = 'N' " & _
          " order by a." & gsCAT_TLS_LABEL_SPECNO & " "
          
    Set GetLabelSpecByStage = moProRawSql.QueryDatabase(sSQL)
    
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
        Call HandleError(True, typErrInfo, , moAppLog)
    End If
End Function


'================================================================================
' Sub: Prt_WS_CP_VIRTUAL_LOT_LABEL()
'--------------------------------------------------------------------------------
' Description:  <This subroutine will print out a label for FT-> SFG >
'--------------------------------------------------------------------------------
' Author:       <Sam Chan>, <CIT> 2018/08/13
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
'================================================================================
Public Sub Prt_WS_CP_VIRTUAL_LOT_LABEL(ByVal sLotNo As String, _
                                ByVal sProdNo As String, _
                                ByVal sWQty As String, _
                                ByVal sCQty As String, _
                                ByVal sPrinterServer As String, _
                                ByVal moCwMbx As Object, _
                                ByVal oLogCtrl As Object, _
                                ByVal moProRawSql As Object, _
                                Optional ByVal bPrintLotList As Boolean = False)


On Error GoTo ExitHandler:

Dim sMsg            As String
Dim sProcID         As String
Dim typErrInfo      As tErrInfo
Dim lErrCode        As Long
Dim sErrMsg         As String

Dim sSQL As String
Dim colRS As Collection
Dim colRS2 As Collection

Dim sHotLotFlag As String
Dim sErunTicNO As String
Dim sSapRwNo As String
Dim sMfgTicNo As String
Dim sWaferNo1 As String
Dim sWaferNo2 As String
Dim sWaferNo3 As String
Dim iIdx As Integer
Dim iIdx2 As Integer
Dim sWaferNO(25) As String
Dim vWaferNO As Variant
Dim sSortWaferNo As String
Dim bFound As Boolean

Dim sPrintFabLotId As String
Dim sCriteiraString   As String
Dim sCriteiraData     As String

Dim sPrintFlag As String
Dim sProdCode As String
Dim sWaferLevel As String
Dim sProdBody As String

Dim intPosY As Integer
'----
' Init
'----
    sProcID = "Prt_WS_CP_VIRTUAL_LOT_LABEL"
    Call LogProcIn(msMODULE_ID, sProcID, oLogCtrl)
    
    sCriteiraString = "" '此處只有列印, 不需傳入.
    
    'Add by Dexter on 20220809 for CP 併批測試 Phase II-第二階段需求 Key值取消IPN條件 <Start>
    '重新取得完整IPN
    sSQL = "select tla." & gsCAT_TLATT_IPN & " as IPN, sum(tla." & gsCAT_TLATT_WAFERQTY & ") as wqty,sum(tla." & gsCAT_TLATT_CHIPQTY & ") as cqty " & _
                               " from (select distinct " & gsCAT_TVMG_VIRTUALLOTID & " ," & gsCAT_TVMG_LOTID & " " & _
                               "       from " & gsCAT_TBL_VIRTUAL_MERGE & " where " & gsCAT_TVMG_DELETEFLAG & "= 'N'  " & _
                               "        and " & gsCAT_TVMG_VIRTUALLOTID & " ='" & sLotNo & "' ) a " & _
                               " , " & gsCAT_TBL_LOT_ATTRIBUTE & " tla " & _
                               " where a." & gsCAT_TVMG_LOTID & " = tla." & gsCAT_TLATT_LOTID & " " & _
                               " group by tla." & gsCAT_TLATT_IPN & " "
   
    Set colRS = moProRawSql.QueryDatabase(sSQL)
    
    sProdNo = colRS.Item(1).Item(gsCAT_TLATT_IPN)
    'Add by Dexter on 20220809 for CP 併批測試 Phase II-第二階段需求 Key值取消IPN條件 <End>
         
    sSQL = "select fun_get_waferlevel(a." & gsCAT_TIM_IPN & ") as waferlevel " & _
           " ,a." & gsCAT_TIM_MASK_OPTION & " , a." & gsCAT_TIM_BE_OPTION & " " & _
           " from " & gsCAT_TBL_IPN_MASTER & " a " & _
           " where a." & gsCAT_TIM_IPN & "= '" & sProdNo & "' " & _
           " and a." & gsCAT_TIM_DELETE_FLAG & "= 'N' "
    Set colRS = moProRawSql.QueryDatabase(sSQL)
    If colRS.Count > 0 Then
        sWaferLevel = colRS.Item(1).Item("waferlevel")
        sProdCode = Left(sProdNo, 4) & colRS.Item(1).Item(gsCAT_TIM_MASK_OPTION)
    End If
    
    sProdBody = Left(sProdNo, 4)
    
       
    sSQL = "SELECT a." & gsCAT_TVMG_SLOTNO & ",a." & gsCAT_TVMG_LOTID & ",fun_splitlen(tli." & gsCAT_TLI_WAFERID & ",';') as pcs  " & _
            " from " & gsCAT_TBL_VIRTUAL_MERGE & " a ," & gsCAT_TBL_LOT_INFO & " tli " & _
            " where a." & gsCAT_TVMG_LOTID & " = tlI." & gsCAT_TLI_LOT_ID & " " & _
            " and a." & gsCAT_TVMG_DELETEFLAG & " = 'N' " & _
           " and a." & gsCAT_TVMG_VIRTUALLOTID & " = '" & sLotNo & "' " & _
           " order by a." & gsCAT_TVMG_SLOTNO & " "
    Set colRS = moProRawSql.QueryDatabase(sSQL)
    If colRS.Count > 0 Then
        For iIdx = 1 To colRS.Count
            sWaferNO(Val(colRS.Item(iIdx).Item(gsCAT_TVMG_SLOTNO))) = "Y"
        Next
        
        For iIdx = 1 To 13
            If Trim(sWaferNO(iIdx)) = "" Then
                sWaferNo1 = sWaferNo1 & "," & "__"
            Else
                sWaferNo1 = sWaferNo1 & "," & Format(iIdx, "00")
            End If
        Next
        For iIdx = 14 To 25
            If Trim(sWaferNO(iIdx)) = "" Then
                sWaferNo2 = sWaferNo2 & "," & "__"
            Else
                sWaferNo2 = sWaferNo2 & "," & Format(iIdx, "00")
            End If
        Next
        If Left(sWaferNo1, 1) = "," Then sWaferNo1 = Mid(sWaferNo1, 2)
        If Left(sWaferNo2, 1) = "," Then sWaferNo2 = Mid(sWaferNo2, 2)
        
    End If
                                                                  
'----
' Condition Checking
'----

'----
' Action
'----
    On Error Resume Next
    '******
    'Print Label Info
    '******
    
        '20211027 By Dexter for BE#202100279 : CP_VIRTUAL_MERGE 標籤格式修正 START
        sMsg = sMsg & "^XA"
        sMsg = sMsg & "^LH0,0^FS"
        sMsg = sMsg & "^LL1980"
        sMsg = sMsg & "^MD0"
        sMsg = sMsg & "^MNY"
        If bPrintLotList = True Then
            sMsg = sMsg & "^FO30,25^A0N,48.36^CI0^FR^FDVitualLotId : " & sLotNo & "^FS"
            sMsg = sMsg & "^BY2,2.0^FO540,30^B3N,N,60,N,N^FD" & sLotNo & "^FS"
        Else
            sMsg = sMsg & "^FO20,25^A0N,32.24^CI0^FR^FDVLotId:" & sLotNo & "^FS"
            sMsg = sMsg & "^BY2,2.0^FO270,20^BAN,40,N,N,N^FD" & sLotNo & "^FS"
        End If
        sMsg = sMsg & "^FO30,102^A0N,48.36^CI0^FR^FDProdBody : " & sProdBody & "^FS"
        sMsg = sMsg & "^FO405,102^A0N,48.36^CI0^FR^FDProdLevel : " & sWaferLevel & "^FS"
        sMsg = sMsg & "^FO30,174^A0N,48.36^CI0^FR^FDWaferQty : " & sWQty & "^FS"
        sMsg = sMsg & "^BY2,2.0^FO405,174^B3N,N,60,N,N^FD" & sWQty & "^FS"
        sMsg = sMsg & "^FO30,246^A0N,48.36^CI0^FR^FDChipQty : " & sCQty & "^FS"
        sMsg = sMsg & "^BY2,2.0^FO405,246^B3N,N,60,N,N^FD" & sCQty & "^FS"
        sMsg = sMsg & "^FO30,318^A0N,48.36^CI0^FR^FDID# :^FS"
        '20211027 By Dexter for BE#202100279 : CP_VIRTUAL_MERGE 標籤格式修正 END
    
    
'        sMsg = sMsg & "^XA"
'        sMsg = sMsg & "^LH0,0^FS"
'        sMsg = sMsg & "^LL280"
'        sMsg = sMsg & "^MD0"
'        sMsg = sMsg & "^MNY"
'        If bPrintLotList = True Then
'            sMsg = sMsg & "^FO20,25^A0N,32.24^CI0^FR^FDVitualLotId : " & sLotNo & "^FS"
'            sMsg = sMsg & "^BY2,2.0^FO360,20^B3N,N,40,N,N^FD" & sLotNo & "^FS"
'        Else
'            sMsg = sMsg & "^FO20,25^A0N,32.24^CI0^FR^FDVLotId:" & sLotNo & "^FS"
'            sMsg = sMsg & "^BY2,2.0^FO270,20^BAN,40,N,N,N^FD" & sLotNo & "^FS"
'        End If
'        sMsg = sMsg & "^FO20,68^A0N,32.24^CI0^FR^FDProdBody : " & sProdbody & "^FS"
'        sMsg = sMsg & "^FO270,68^A0N,32.24^CI0^FR^FDProdType : " & sWaferLevel & "^FS"
'        sMsg = sMsg & "^FO20,116^A0N,32.24^CI0^FR^FDWaferQty : " & sWQty & "^FS"
'        sMsg = sMsg & "^BY2,2.0^FO270,116^B3N,N,40,N,N^FD" & sWQty & "^FS"
'        sMsg = sMsg & "^FO20,164^A0N,32.24^CI0^FR^FDChipQty : " & sCQty & "^FS"
'        sMsg = sMsg & "^BY2,2.0^FO270,164^B3N,N,40,N,N^FD" & sCQty & "^FS"
'        sMsg = sMsg & "^FO20,212^A0N,32.24^CI0^FR^FDID#:^FS"
        
        '01~13
        sMsg = sMsg & "^FO136,318^A0N,48.36^CI0^FR^FD" & sWaferNo1 & "^FS"
        
        
'        sMsg = sMsg & "^FO84,212^A0N,32.24^CI0^FR^FD01^FS"
'        sMsg = sMsg & "^FO120,212^A0N,32.24^CI0^FR^FD02^FS"
'        sMsg = sMsg & "^FO156,212^A0N,32.24^CI0^FR^FD__^FS"
'        sMsg = sMsg & "^FO192,212^A0N,32.24^CI0^FR^FD04^FS"
'        sMsg = sMsg & "^FO228,212^A0N,32.24^CI0^FR^FD05^FS"
'        sMsg = sMsg & "^FO264,212^A0N,32.24^CI0^FR^FD06^FS"
'        sMsg = sMsg & "^FO300,212^A0N,32.24^CI0^FR^FD__^FS"
'        sMsg = sMsg & "^FO336,212^A0N,32.24^CI0^FR^FD08^FS"
'        sMsg = sMsg & "^FO372,212^A0N,32.24^CI0^FR^FD09^FS"
'        sMsg = sMsg & "^FO408,212^A0N,32.24^CI0^FR^FD10^FS"
'        sMsg = sMsg & "^FO444,212^A0N,32.24^CI0^FR^FD11^FS"
'        sMsg = sMsg & "^FO480,212^A0N,32.24^CI0^FR^FD__^FS"
'        sMsg = sMsg & "^FO516,212^A0N,32.24^CI0^FR^FD13^FS"
        
        '14~25
        sMsg = sMsg & "^FO136,378^A0N,48.36^CI0^FR^FD" & sWaferNo2 & "^FS"
        
'        sMsg = sMsg & "^FO84,252^A0N,32.24^CI0^FR^FD14^FS"
'        sMsg = sMsg & "^FO120,252^A0N,32.24^CI0^FR^FD15^FS"
'        sMsg = sMsg & "^FO156,252^A0N,32.24^CI0^FR^FD16^FS"
'        sMsg = sMsg & "^FO192,252^A0N,32.24^CI0^FR^FD17^FS"
'        sMsg = sMsg & "^FO228,252^A0N,32.24^CI0^FR^FD18^FS"
'        sMsg = sMsg & "^FO264,252^A0N,32.24^CI0^FR^FD__^FS"
'        sMsg = sMsg & "^FO300,252^A0N,32.24^CI0^FR^FD20^FS"
'        sMsg = sMsg & "^FO336,252^A0N,32.24^CI0^FR^FD21^FS"
'        sMsg = sMsg & "^FO372,252^A0N,32.24^CI0^FR^FD22^FS"
'        sMsg = sMsg & "^FO408,252^A0N,32.24^CI0^FR^FD__^FS"
'        sMsg = sMsg & "^FO444,252^A0N,32.24^CI0^FR^FD24^FS"
'        sMsg = sMsg & "^FO480,252^A0N,32.24^CI0^FR^FD25^FS"
'        sMsg = sMsg & "^FO112,212^A0N,32.24^CI0^FR^FD,^FS"
'        sMsg = sMsg & "^FO148,212^A0N,32.24^CI0^FR^FD,^FS"
'        sMsg = sMsg & "^FO184,212^A0N,32.24^CI0^FR^FD,^FS"
'        sMsg = sMsg & "^FO220,212^A0N,32.24^CI0^FR^FD,^FS"
'        sMsg = sMsg & "^FO256,212^A0N,32.24^CI0^FR^FD,^FS"
'        sMsg = sMsg & "^FO292,212^A0N,32.24^CI0^FR^FD,^FS"
'        sMsg = sMsg & "^FO328,212^A0N,32.24^CI0^FR^FD,^FS"
'        sMsg = sMsg & "^FO364,212^A0N,32.24^CI0^FR^FD,^FS"
'        sMsg = sMsg & "^FO400,212^A0N,32.24^CI0^FR^FD,^FS"
'        sMsg = sMsg & "^FO436,212^A0N,32.24^CI0^FR^FD,^FS"
'        sMsg = sMsg & "^FO472,212^A0N,32.24^CI0^FR^FD,^FS"
'        sMsg = sMsg & "^FO508,212^A0N,32.24^CI0^FR^FD,^FS"
'        sMsg = sMsg & "^FO112,252^A0N,32.24^CI0^FR^FD,^FS"
'        sMsg = sMsg & "^FO148,252^A0N,32.24^CI0^FR^FD,^FS"
'        sMsg = sMsg & "^FO184,252^A0N,32.24^CI0^FR^FD,^FS"
'        sMsg = sMsg & "^FO220,252^A0N,32.24^CI0^FR^FD,^FS"
'        sMsg = sMsg & "^FO256,252^A0N,32.24^CI0^FR^FD,^FS"
'        sMsg = sMsg & "^FO292,252^A0N,32.24^CI0^FR^FD,^FS"
'        sMsg = sMsg & "^FO328,252^A0N,32.24^CI0^FR^FD,^FS"
'        sMsg = sMsg & "^FO364,252^A0N,32.24^CI0^FR^FD,^FS"
'        sMsg = sMsg & "^FO400,252^A0N,32.24^CI0^FR^FD,^FS"
'        sMsg = sMsg & "^FO436,252^A0N,32.24^CI0^FR^FD,^FS"
'        sMsg = sMsg & "^FO472,252^A0N,32.24^CI0^FR^FD,^FS"

        If bPrintLotList = True Then
            sSQL = "SELECT distinct a." & gsCAT_TVMG_LOTID & ",fun_splitlen(tli." & gsCAT_TLI_WAFERID & ",';') as pcs  " & _
                    " from " & gsCAT_TBL_VIRTUAL_MERGE & " a ," & gsCAT_TBL_LOT_INFO & " tli " & _
                    " where a." & gsCAT_TVMG_LOTID & " = tlI." & gsCAT_TLI_LOT_ID & " " & _
                    " and a." & gsCAT_TVMG_DELETEFLAG & " = 'N' " & _
                   " and a." & gsCAT_TVMG_VIRTUALLOTID & " = '" & sLotNo & "' " & _
                   " order by a." & gsCAT_TVMG_LOTID & " "
            Set colRS = moProRawSql.QueryDatabase(sSQL)
        
            '20210930 By Dexter for BE#202100279 : CP_VIRTUAL_MERGE 標籤格式修正 START
        If colRS.Count > 0 Then
                intPosY = 480
                For iIdx = 1 To colRS.Count
                    sMsg = sMsg & "^FO30," & intPosY & "^A0N,48.36^CI0^FR^FDLotId" & Format(iIdx, "00") & " : " & colRS.Item(iIdx).Item(gsCAT_TVMG_LOTID) & "   " & colRS.Item(iIdx).Item("pcs") & " pc^FS"
                    sMsg = sMsg & "^BY2,2.0^FO624," & intPosY & "^B3N,N,48,N,N^FD" & colRS.Item(iIdx).Item(gsCAT_TVMG_LOTID) & "^FS"
                    intPosY = intPosY + 60
                Next
            End If
        '20210930 By Dexter for BE#202100279 : CP_VIRTUAL_MERGE 標籤格式修正 END
        
'            If colRS.Count > 0 Then
'                intPosY = 320
'                For iIdx = 1 To colRS.Count
'                    sMsg = sMsg & "^FO20," & intPosY & "^A0N,32.24^CI0^FR^FDLotId" & Format(iIdx, "00") & " : " & colRS.Item(iIdx).Item(gsCAT_TVMG_LOTID) & "   " & colRS.Item(iIdx).Item("pcs") & " pc^FS"
'                    sMsg = sMsg & "^BY2,2.0^FO416," & intPosY & "^B3N,N,32,N,N^FD" & colRS.Item(iIdx).Item(gsCAT_TVMG_LOTID) & "^FS"
'                    intPosY = intPosY + 40
'                Next
'            End If
        
'            sMsg = sMsg & "^FO20,320^A0N,32.24^CI0^FR^FDLotId01 : 3U81500BA1   1 pc^FS"
'            sMsg = sMsg & "^BY2,2.0^FO416,320^B3N,N,32,N,N^FD3U81500BA1^FS"
'            sMsg = sMsg & "^FO20,360^A0N,32.24^CI0^FR^FDLotId02 : 3U871703A1   1 pc^FS"
'            sMsg = sMsg & "^BY2,2.0^FO416,360^B3N,N,32,N,N^FD3U871703A1^FS"
'            sMsg = sMsg & "^FO20,400^A0N,32.24^CI0^FR^FDLotId03 : 3X089904A1   1 pc^FS"
'            sMsg = sMsg & "^BY2,2.0^FO416,400^B3N,N,32,N,N^FD3X089904A1^FS"
'            sMsg = sMsg & "^FO20,440^A0N,32.24^CI0^FR^FDLotId04 : 3X090006A1   1 pc^FS"
'            sMsg = sMsg & "^BY2,2.0^FO416,440^B3N,N,32,N,N^FD3X090006A1^FS"
'            sMsg = sMsg & "^FO20,480^A0N,32.24^CI0^FR^FDLotId05 : 3X131302A1   1 pc^FS"
'            sMsg = sMsg & "^BY2,2.0^FO416,480^B3N,N,32,N,N^FD3X131302A1^FS"
'            sMsg = sMsg & "^FO20,520^A0N,32.24^CI0^FR^FDLotId06 : 3X163507A1   1 pc^FS"
'            sMsg = sMsg & "^BY2,2.0^FO416,520^B3N,N,32,N,N^FD3X163507A1^FS"
        End If


        sMsg = sMsg & "^XZ"
    
    '******
    ' End of Label Info
    '******

    '******
    'Send message to TDS
    '******
        
        If InStr(1, sPrinterServer, "@") <> 0 Then
            sPrinterServer = VBA.Left(sPrinterServer, (InStr(1, sPrinterServer, "@")) - 1)
        End If
        
        If Not moCwMbx.Send("MBX_" & sPrinterServer, sMsg) Then
            Call moCwMbx.GetError(lErrCode, sErrMsg)
            Call RaiseError(lErrCode, sErrMsg)
        End If
    '******

'----
' Done
'----
ExitHandler:
    ' NOTE 1:
    ' MUST CALL GetErrInfo() here first before another action
    Call GetErrInfo(msMODULE_ID, sProcID, typErrInfo, Erl)
    Call LogProcOut(msMODULE_ID, sProcID, typErrInfo, oLogCtrl)
    On Error Resume Next

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
        Call HandleError(False, typErrInfo, , oLogCtrl, True)
    End If
End Sub
'================================================================================
' Sub: Prt_FT_TR_LABEL()
'--------------------------------------------------------------------------------
' Description:  Req.202000413 'Add gsLABEL_FT_TR_LABEL by HouYu for 202000413 on 20220422
'--------------------------------------------------------------------------------
' Create Date: 20220422
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
'================================================================================
Public Sub Prt_FT_TR_LABEL(ByVal sLotID As String, _
                            ByVal sIPN As String, _
                            ByVal sQtyReel As String, _
                            ByVal sEqID As String, _
                            ByVal sReelId As String, _
                            ByVal sPrinterServer As String, _
                            ByVal moCwMbx As Object, _
                            ByVal oLogCtrl As Object)

On Error GoTo ExitHandler:
Dim sProcID         As String
Dim typErrInfo      As tErrInfo
Dim sMsg            As String
Dim lErrCode        As Long
Dim sErrMsg         As String


'----
' Init
'----
    sProcID = "Prt_FT_BOX_COUNTING_LABEL"
    
    Call LogProcIn(msMODULE_ID, sProcID, oLogCtrl)
    
'----
' Condition Checking
'----

'----
' Action
'----
    On Error Resume Next
    
    '******
    'Print Label Info
    '******
        
    sMsg = sMsg & "^XA"
    sMsg = sMsg & "^LL440"
    sMsg = sMsg & "^LH0,0^FS"
    sMsg = sMsg & "^FO40,40^A0N,37,33^FH\^FDLOTID:" & sLotID & "^FS"
    sMsg = sMsg & "^BY2,3.0^FO 40,75^BCN,56,N,Y,N^FR^FD>:" & sLotID & "^FS"
    sMsg = sMsg & "^FO40,180^A0N,37,33^FH\^FDQty/Reel:" & sQtyReel & "^FS"
    sMsg = sMsg & "^FO40,220^A0N,37,33^FH\^FDEQID:" & sEqID & "^FS"
    sMsg = sMsg & "^FO320,220^A0N,37,33^FH\^FDReel ID:" & sReelId & "^FS"
    sMsg = sMsg & "^FO40,140^A0N,37,33^FH\^FDIPN Info:" & sIPN & "^FS"
    sMsg = sMsg & "^PQ1,0,1,Y^XZ"
    sMsg = sMsg & "^XZ"
    
    '******
    ' End of Label Info
    '******

    '******
    'Send message to TDS
    '******
    If InStr(1, sPrinterServer, "@") <> 0 Then
        sPrinterServer = VBA.Left(sPrinterServer, (InStr(1, sPrinterServer, "@")) - 1)
    End If
    
    If Not moCwMbx.Send("MBX_" & sPrinterServer, sMsg) Then
        Call moCwMbx.GetError(lErrCode, sErrMsg)
        Call RaiseError(lErrCode, sErrMsg)
    End If
    '******

'----
' Done
'----
ExitHandler:
    ' NOTE 1:
    ' MUST CALL GetErrInfo() here first before another action
    Call GetErrInfo(msMODULE_ID, sProcID, typErrInfo, Erl)
    Call LogProcOut(msMODULE_ID, sProcID, typErrInfo, oLogCtrl)
    On Error Resume Next

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
        Call HandleError(False, typErrInfo, , oLogCtrl, True)
    End If
End Sub

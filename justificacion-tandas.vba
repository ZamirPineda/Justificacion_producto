Option Explicit
Public i As Integer

Sub TD()

Hoja3.Range("R1").Value = "1"
Call cierre

End Sub
Sub PP_PT()

Hoja3.Range("R1").Value = "2"
Call cierre

End Sub

Sub cierre()
Dim application
Dim connection
Dim filas As Integer

Hoja3.Range("C4").Select
filas = Range(Selection, Selection.End(xlDown)).Count
filas = filas + 3

If Not IsObject(application) Then
   Set SapGuiAuto = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   Set session = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session, "on"
   WScript.ConnectObject application, "on"
End If

For i = 4 To filas
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").Text = "/ncooispi"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:PPIO_ENTRY:1100/ctxtPPIO_ENTRY_SC1100-ALV_VARIANT").Text = "//d GOMEZ"
session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_PAUFNR-LOW").Text = Hoja3.Cells(i, 3).Value
session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_PAUFNR-LOW").SetFocus
session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_PAUFNR-LOW").caretPosition = 8
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 8
session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").currentCellColumn = "AUFNR"
session.findById("wnd[0]").sendVKey 18
session.findById("wnd[0]/usr/tabsTABSTRIP_5115/tabpKOZE/ssubSUBSCR_5115:SAPLCOKO:5120/txtCAUFVD-GAMNG").Text = Hoja3.Cells(i, 11).Value
session.findById("wnd[0]/usr/tabsTABSTRIP_5115/tabpKOZE/ssubSUBSCR_5115:SAPLCOKO:5120/txtCAUFVD-GAMNG").caretPosition = 3
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 32
session.findById("wnd[0]/mbar/menu[2]/menu[9]/menu[0]").Select
session.findById("wnd[0]/tbar[1]/btn[45]").press
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").SetFocus
session.findById("wnd[1]/tbar[0]/btn[0]").press
session.findById("wnd[1]/usr/ctxtDY_PATH").Text = "\\Dncacusuge02\d\sip\Archivos Notificadores\Archivos Notificadores\Salchichas\"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = "costos.xls"
session.findById("wnd[1]/usr/ctxtDY_FILENAME").caretPosition = 10
session.findById("wnd[1]/tbar[0]/btn[11]").press

Call abrir

If Hoja3.Range("T11").Value <> "E" Then
    Hoja3.Range("T11").Select
    Selection.Copy
    Hoja3.Range("P8").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    
End If

If Hoja3.Range("T12").Value <> "E" Then
    Hoja3.Range("T12").Select
    Selection.Copy
    Hoja3.Range("Q8").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
    
End If

If Hoja3.Range("P9").Value >= 5 Or Hoja3.Range("Q9").Value >= 5 Then
    Hoja3.Cells(i, 12).Value = "VARIACION"
    UserForm1.Show
    
Else
    Hoja3.Cells(i, 12).Value = "OK"
    Call cierre1
End If
Next i
End Sub

Sub abrir()
Dim num As Byte
Dim fila As Range


application.Workbooks.Open ("\\Dncacusuge02\d\sip\Archivos Notificadores\Archivos Notificadores\Salchichas\costos.xls")

If application.Workbooks("inventario peliculas 2.0").Worksheets("CIERRE TECNICO").Range("R1").Value = "2" Then

    ActiveWorkbook.application.Columns("M:M").Select
    ActiveWorkbook.application.Selection.TextToColumns Destination:=Range("M1"), DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
            Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
            :=Array(1, 1), TrailingMinusNumbers:=True
    ActiveWorkbook.application.Columns("N:N").Select
    ActiveWorkbook.application.Selection.TextToColumns Destination:=Range("N1"), DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
            Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
            :=Array(1, 1), TrailingMinusNumbers:=True
    
    Set fila = ActiveWorkbook.Worksheets("COSTOS").Columns(3).Find("Liquidacion", lookat:=xlWhole)
    num = fila.Row
    ActiveWorkbook.Worksheets("COSTOS").Cells(num, 13).Select
    ActiveWorkbook.Worksheets("COSTOS").Cells(num, 13).Copy
    application.Workbooks("inventario peliculas 2.0").Activate
    Hoja3.Range("P5").Select
    ActiveSheet.Paste
    application.Workbooks("costos").Activate
    Selection.End(xlDown).End(xlDown).End(xlDown).Select
    Range(Selection, Selection.Offset(0, 1)).Select
    Selection.Copy
    application.Workbooks("inventario peliculas 2.0").Activate
    Hoja3.Range("P8").Select
    ActiveSheet.Paste
    application.Workbooks("costos").Close (False)
    ElseIf application.Workbooks("inventario peliculas 2.0").Worksheets("CIERRE TECNICO").Range("R1").Value = "1" Then
    ActiveWorkbook.application.Columns("K:K").Select
    ActiveWorkbook.application.Selection.TextToColumns Destination:=Range("K1"), DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=True, _
            Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
            :=Array(1, 1), TrailingMinusNumbers:=True
End If

End Sub

Sub cierre1()
Dim application
Dim connection

If Not IsObject(application) Then
   Set SapGuiAuto = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   Set session = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session, "on"
   WScript.ConnectObject application, "on"
End If
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").Text = "/ncooispi"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").resizeWorkingPane 171, 28, False
session.findById("wnd[0]/usr/ssub%_SUBSCREEN_TOPBLOCK:PPIO_ENTRY:1100/ctxtPPIO_ENTRY_SC1100-ALV_VARIANT").Text = "//d GOMEZ"
session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_PAUFNR-LOW").Text = Hoja3.Cells(i, 3).Value
session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_PAUFNR-LOW").SetFocus
session.findById("wnd[0]/usr/tabsTABSTRIP_SELBLOCK/tabpSEL_00/ssub%_SUBSCREEN_SELBLOCK:PPIO_ENTRY:1200/ctxtS_PAUFNR-LOW").caretPosition = 8
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 8
session.findById("wnd[0]/usr/cntlCUSTOM/shellcont/shell/shellcont/shell").currentCellColumn = "AUFNR"
session.findById("wnd[0]").sendVKey 18
session.findById("wnd[0]/usr/tabsTABSTRIP_5115/tabpKOWE").Select
session.findById("wnd[0]/usr/tabsTABSTRIP_5115/tabpKOWE/ssubSUBSCR_5115:SAPLCOKO:5190/chkAFPOD-ELIKZ").Selected = True
session.findById("wnd[0]/usr/tabsTABSTRIP_5115/tabpKOWE/ssubSUBSCR_5115:SAPLCOKO:5190/chkAFPOD-ELIKZ").SetFocus
session.findById("wnd[0]/mbar/menu[0]/menu[9]/menu[12]/menu[3]").Select
session.findById("wnd[0]").sendVKey 11
End Sub


Sub variacion()

Dim application
Dim connection

If Not IsObject(application) Then
   Set SapGuiAuto = GetObject("SAPGUI")
   Set application = SapGuiAuto.GetScriptingEngine
End If
If Not IsObject(connection) Then
   Set connection = application.Children(0)
End If
If Not IsObject(session) Then
   Set session = connection.Children(0)
End If
If IsObject(WScript) Then
   WScript.ConnectObject session, "on"
   WScript.ConnectObject application, "on"
End If
session.findById("wnd[0]").maximize
session.findById("wnd[0]/tbar[0]/okcd").Text = "/ncor6n"
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ssubSUB01:SAPLCORU_S:0010/subSLOT_HDR:SAPLCORU_S:5117/ctxtAFRUD-AUFNR").Text = Hoja3.Cells(i, 3).Value
session.findById("wnd[0]/usr/ssubSUB01:SAPLCORU_S:0010/subSLOT_HDR:SAPLCORU_S:5117/ctxtAFRUD-VORNR").Text = "011"
session.findById("wnd[0]/usr/ssubSUB01:SAPLCORU_S:0010/subSLOT_HDR:SAPLCORU_S:5117/ctxtAFRUD-VORNR").SetFocus
session.findById("wnd[0]/usr/ssubSUB01:SAPLCORU_S:0010/subSLOT_HDR:SAPLCORU_S:5117/ctxtAFRUD-VORNR").caretPosition = 3
session.findById("wnd[0]").sendVKey 0
If Hoja3.Cells(i, 8).Value > Hoja3.Cells(i, 7).Value Then
session.findById("wnd[1]/usr/btnOPTION2").press
session.findById("wnd[1]/usr/btnOPTION2").press
End If
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ssubSUB01:SAPLCORU_S:0010/subSLOT_DET1:SAPLCORU_S:0215/txtAFRUD-LMNGA").Text = ""
session.findById("wnd[0]/usr/ssubSUB01:SAPLCORU_S:0010/subSLOT_DET1:SAPLCORU_S:0215/txtAFRUD-LMNGA").SetFocus
session.findById("wnd[0]/usr/ssubSUB01:SAPLCORU_S:0010/subSLOT_DET1:SAPLCORU_S:0215/txtAFRUD-LMNGA").caretPosition = 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]").sendVKey 0
session.findById("wnd[0]/usr/ssubSUB01:SAPLCORU_S:0010/subSLOT_DET4:SAPLCORU_S:0800/cntlTEXTEDITOR1/shellcont/shell").Text = Hoja3.Cells(i, 13).Value
session.findById("wnd[0]/usr/ssubSUB01:SAPLCORU_S:0010/subSLOT_DET4:SAPLCORU_S:0800/cntlTEXTEDITOR1/shellcont/shell").setSelectionIndexes 30, 30
session.findById("wnd[0]").sendVKey 11
session.findById("wnd[0]").sendVKey 0

End Sub

Attribute VB_Name = "VBA_SAP"
'Function to call a Yes No Window that will not lock the Excel SAP.
Private Declare PtrSafe Function MessageBox _
    Lib "User32" Alias "MessageBoxA" _
       (ByVal hWnd As Long, _
        ByVal lpText As String, _
        ByVal lpCaption As String, _
        ByVal wType As Long) _
    As Long
' Function that return a Dictionary of the ditinct values given a Column
Private Function DistinctVals(a, Optional col = 1)
    Dim i&, v: v = a
    With CreateObject("Scripting.Dictionary")
        For i = 1 To UBound(v): .Item(v(i, col)) = 1: Next
        DistinctVals = application.Transpose(.keys)
    End With
End Function
'Function to connect to SAP's Scripting Appi
Function Connect_SAP()
    Dim SapGuiAuto
    Dim application
    Dim connection
    Dim session
    On Error Resume Next
    Set SapGuiAuto = GetObject("SAPGUI")
    If Err.Number <> 0 Then
        MsgBox "Unable to get SAPGUI", vbCritical
        Exit Function
    End If
    Set application = SapGuiAuto.GetScriptingEngine
    If Err.Number <> 0 Then
        MsgBox "Unable to get SAP Scripting Engine", vbCritical
        Exit Function
    End If
    Set connection = application.Children(0)
    If connection Is Nothing Then
        MsgBox "Unable to get connection", vbCritical
        Exit Function
    End If
    Set session = connection.Children(0)
    If session Is Nothing Then
        MsgBox "Unable to get session", vbCritical
        Exit Function
    End If
    If IsObject(WScript) Then
        WScript.ConnectObject session, "on"
        WScript.ConnectObject application, "on"
    End If
    ' Return the session object as needed
    Set Connect_SAP = session
End Function
Sub Pagare_Cliente()
'This sub using SAP Scripting Engine Applies a cheque into a cliente acount
'sensitive data is change by #@name of something

    Dim wbR As Workbook
    Dim wbP As Workbook
    Dim wsR As Worksheet
    Dim wsP As Worksheet
    Dim dicFusis As Object
    Set dicFusis = CreateObject("Scripting.Dictionary")
    fichero = application.GetOpenFilename("Archivos Excel (*.xlsx;*.xls),*.xlsx;*.xls", Title:="Abre la realción de pago")
    If fichero = False Then
        MsgBox ("No se ha selecionado fichero. Se cancela el proceso")
        Exit Sub
    Plantilla = application.GetOpenFilename("Archivos Excel (*.xlsx),*.xlsx", Title:="Abre la Plantilla Call Transaction")
    If Plantilla = False Then
        MsgBox ("No se ha selecionado fichero. Se cancela el proceso")
        Exit Sub
    End If
    Set wbR = Workbooks.Open(fichero)
    Set wsR = wbR.Sheets(1)
    Set wbP = Workbooks.Open(Plantilla)
    Set wsP = wbP.Sheets(1)
    wbR.Activate
    wsR.Columns("A:F").AutoFit
    importeReal = FormatNumber(CDbl(application.InputBox("Introduce el total del Pagare", Type:=1, Left:=(application.Width / 2), Top:=(application.Height / 2))), -1, vbUseDefault, vbUseDefault, vbUseDefault)
    importeRelacion = FormatNumber(CDbl(wsR.Range("B8").Value), -1, vbUseDefault, vbUseDefault, vbUseDefault)
    If importeReal = importeRelacion Then
        Finalrow = wsR.Cells(Rows.Count, 2).End(xlUp).Row ' Final Row in the "Relacion de pago"
        InicialRow = 10 ' Start Row in the "Relacion de pago"
        FPI = 10 ' Start Row in the Template
        FPF = wsP.Cells(Rows.Count, 4).End(xlUp).Row ' Final Row in the Template
        If FPI < FPF Then
            wsP.Range("D10:D" & FPF & "").Clear 'Cleans the Templeate of previous data (Batch input for SAP)
        End If
        wsR.Cells(8, 4).Replace what:="/", Replacement:=".", LookAt:=xlPart 'SAP uses time format as #dd.mm.aaaa
        wsR.Range("B:B").Replace what:="-", Replacement:="", LookAt:=xlPart 'Format Documentos will search and call in SAP
        vencimiento = wsR.Cells(8, 4).text
        asignacion = Right(vencimiento, 4) & Mid(vencimiento, 4, 2) & Left(vencimiento, 2)
        Fecha = Date
        Fecha = Format(Fecha, "dd.mm.yyyy")
        NPag = wsR.Cells(8, 1).Value
        Cargos = 0# ' Cargos will be a manualy note in the cliente count < 0
        Abonos = 0# ' Abonos will be a munaly note in the cliente count > 0
        Facturas = 0# 'Facturas will be search in SAP via Batch input Templete
        'Loop to categorize the lines in the "relacion de pago"
        'This loop is the one that can change dependeing of the client
        'This is a generic one.
        For i = InicialRow To Finalrow
            Largo = Len(wsR.Cells(i, 2).Value)
            Tipo = Left(wsR.Cells(i, 2).Value, 1)
            ValorC = wsR.Cells(i, 4).Value
            If Largo = 7 Then
                If Tipo = "4" Then
                    wsP.Cells(FPI, 4).Value = "X" & wsR.Cells(i, 2).Value
                    FPI = FPI + 1
                    Facturas = Facturas + wsR.Cells(i, 4).Value
                ElseIf Tipo = "5" Or Tipo = "6" Or Tipo = "7" Then
                    wsP.Cells(FPI, 4).Value = "V" & wsR.Cells(i, 2).Value
                    FPI = FPI + 1
                    Facturas = Facturas + wsR.Cells(i, 4).Value
                ElseIf Tipo = "7" Then
                    wsP.Cells(FPI, 4).Value = "Y" & wsR.Cells(i, 2).Value
                    FPI = FPI + 1
                    Facturas = Facturas + wsR.Cells(i, 4).Value
                End If
            ElseIf Tipo = "C" And ValorC < 0 Then
                Cargos = Cargos + wsR.Cells(i, 4).Value
            ElseIf (Tipo = "C" Or Tipo = "A") And ValorC > 0 Then
                Abonos = Abonos + wsR.Cells(i, 4).Value
            Else
                dicFusis.Add i, i 'Dictionay of "Special" Debits and Credits
            End If
        Next i
        wsP.Cells(2, 5).Value = Fecha
        wsP.Cells(2, 7).Value = Fecha
        wbP.Save
        wbP.Close
        cliente = InputBox("Introduce el codigo del cliente")
        session = Connect_SAP()
        session.findById("wnd[0]").resizeWorkingPane 92, 30, False
        session.findById("wnd[0]/tbar[0]/okcd").text = "z2s_k0021" 'Batch Input transaction
        session.findById("wnd[0]").sendVKey 0 'Send key ENTER
        session.findById("wnd[0]/usr/radP_CALLT").Select 'Select the type o Batch input (from file or from server)
        session.findById("wnd[0]/usr/ctxtP_FILE").text = Plantilla 'Enter the path to the file into SAP
        session.findById("wnd[0]/tbar[1]/btn[8]").press ' Run the Transaction in SAP
        session.findById("wnd[1]/usr/btnSPOP-OPTION1").press
        session.findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPDF05X:6102/btnICON_SELECT_ALL").press
        session.findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPDF05X:6102/btnIC_Z+").press
        Texto = session.findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPDF05X:6102/txtRF05A-NETTO").DisplayedText
        ImporteSAP = FormatNumber(CDbl(Format(Texto, General)), -1, vbUseDefault, vbUseDefault, vbUseDefault)
        Facturas = FormatNumber(CDbl(Format(Facturas, General)), -1, vbUseDefault, vbUseDefault, vbUseDefault)
        session.findById("wnd[0]/tbar[1]/btn[14]").press
        If ImporteSAP = Facturas Then
            GoTo Line1
        Else
            diff = ImporteSAP - Facturas
            MsgBox ("Hay diferencias entre las partidas y la relacion de facturas. " & diff)
            Confirmacion = MsgBox("¿Quiere ajustar por redondeo?", vbYesNo)
            If Confirmacion = vbYes Then
                GoTo Line2
            Else
                MsgBox ("Se Cancela el Proceso")
                Exit Sub
            End If
        End If
    Else
        MsgBox ("El importe no cuadra. Se cancela el proceso. Revise la relación")
        Exit Sub
    End If
    
Line1:
'Line to apply the payment into the cliente account
            session.findById("wnd[0]/usr/ctxtRF05A-NEWBS").text = "09" 'Cheque number in this SAP
            'Cliente = Application.InputBox("Introduce el codigo del CLIENTE")
            session.findById("wnd[0]/usr/ctxtRF05A-NEWKO").text = cliente
            'CME = Application.InputBox("Introduce la variable de CME")
            session.findById("wnd[0]/usr/ctxtRF05A-NEWUM").text = "W" 'Cheque CME in this SAP
            session.findById("wnd[0]").sendVKey 0
            session.findById("wnd[0]/usr/txtBSEG-WRBTR").text = importeReal
            session.findById("wnd[0]/usr/txtBSEG-ZUONR").text = Right(Fecha, 4) & Mid(Fecha, 4, 2) & Left(Fecha, 2)
            session.findById("wnd[0]/usr/ctxtBSEG-GSBER").text = "#CategoryAccount"
            session.findById("wnd[0]/usr/ctxtBSEG-SGTXT").text = "PAG. #NombreCliente " & NPag & " VTO. " & vencimiento
            session.findById("wnd[0]/usr/ctxtBSEG-ZFBDT").text = vencimiento
            session.findById("wnd[0]").sendVKey 0
            session.findById("wnd[0]/tbar[1]/btn[14]").press
            Cargos = Cargos * (-1)
            If Abonos <> 0 Then
                session.findById("wnd[0]/usr/ctxtRF05A-NEWBS").text = "16" 'Number of accounting entry for Credit in this SAP
                session.findById("wnd[0]/usr/ctxtRF05A-NEWKO").text = cliente
                session.findById("wnd[0]").sendVKey 0
                session.findById("wnd[0]/usr/txtBSEG-WRBTR").text = Abonos
                session.findById("wnd[0]/usr/ctxtBSEG-GSBER").text = "#CategoryAccount"
                session.findById("wnd[0]/usr/ctxtBSEG-ZFBDT").text = vencimiento
                session.findById("wnd[0]/usr/txtBSEG-ZUONR").text = asignacion
                session.findById("wnd[0]/usr/ctxtBSEG-SGTXT").text = "TOTAL ABONOS " & NPag & " VTO. " & vencimiento
                session.findById("wnd[0]").sendVKey 0
                session.findById("wnd[0]/tbar[1]/btn[14]").press
            End If
            If Cargos <> 0 Then
                session.findById("wnd[0]/usr/ctxtRF05A-NEWBS").text = "06" 'Number of accounting entry for Debit in this SAP
                session.findById("wnd[0]/usr/ctxtRF05A-NEWKO").text = cliente
                session.findById("wnd[0]").sendVKey 0
                session.findById("wnd[0]/usr/txtBSEG-WRBTR").text = Cargos
                session.findById("wnd[0]/usr/ctxtBSEG-GSBER").text = "#CategoryAccount"
                session.findById("wnd[0]/usr/ctxtBSEG-ZFBDT").text = vencimiento
                session.findById("wnd[0]/usr/txtBSEG-ZUONR").text = asignacion
                session.findById("wnd[0]/usr/ctxtBSEG-SGTXT").text = "TOTAL CARGOS " & NPag & " VTO. " & vencimiento
                session.findById("wnd[0]").sendVKey 0
                session.findById("wnd[0]/tbar[1]/btn[14]").press
            End If
            Dim FusiLinea
            Dim TipoCargo
            Dim VCargo
            Dim NomCargo
            Dim SddCargo
            NFusis = dicFusis.Count
            Matriz = dicFusis.Items
            For i = 0 To NFusis - 1
            'This loop applies the "Special" Debits and Credits
                FusiLinea = Matriz(i)
                VCargo = wsR.Cells(FusiLinea, 4).Value
                NomCargo = wsR.Cells(FusiLinea, 2).Value
                    If VCargo > 0 Then
                        session.findById("wnd[0]/usr/ctxtRF05A-NEWBS").text = "16"
                        session.findById("wnd[0]/usr/ctxtRF05A-NEWKO").text = cliente
                        session.findById("wnd[0]").sendVKey 0
                        session.findById("wnd[0]/usr/txtBSEG-WRBTR").text = VCargo
                        session.findById("wnd[0]/usr/ctxtBSEG-GSBER").text = "LI1E"
                        session.findById("wnd[0]/usr/ctxtBSEG-ZFBDT").text = vencimiento
                        session.findById("wnd[0]/usr/txtBSEG-ZUONR").text = asignacion
                        session.findById("wnd[0]/usr/ctxtBSEG-SGTXT").text = "CARGO " & NomCargo & " COSTES OPERATIVOS"
                        session.findById("wnd[0]").sendVKey 0
                        session.findById("wnd[0]/tbar[1]/btn[14]").press
                    ElseIf VCargo < 0 Then
                        VCargo = VCargo * (-1)
                        session.findById("wnd[0]/usr/ctxtRF05A-NEWBS").text = "06"
                        session.findById("wnd[0]/usr/ctxtRF05A-NEWKO").text = cliente
                        session.findById("wnd[0]").sendVKey 0
                        session.findById("wnd[0]/usr/txtBSEG-WRBTR").text = VCargo
                        session.findById("wnd[0]/usr/ctxtBSEG-GSBER").text = "LI1E"
                        session.findById("wnd[0]/usr/ctxtBSEG-ZFBDT").text = vencimiento
                        session.findById("wnd[0]/usr/txtBSEG-ZUONR").text = asignacion
                        session.findById("wnd[0]/usr/ctxtBSEG-SGTXT").text = "CARGO " & NomCargo & " COSTES OPERATIVOS"
                        session.findById("wnd[0]").sendVKey 0
                        session.findById("wnd[0]/tbar[1]/btn[14]").press
                    End If
            Next
            ini = session.findById("wnd[0]/usr/txtRF05A-ANZAZ").DisplayedText
            session.findById("wnd[0]/mbar/menu[0]/menu[3]").Select 'simulate in SAP
            fin = session.findById("wnd[0]/usr/txtRF05A-ANZAZ").DisplayedText
            ini = FormatNumber(CDbl(Format(ini, General)), -1, vbUseDefault, vbUseDefault, vbUseDefault)
            fin = FormatNumber(CDbl(Format(fin, General)), -1, vbUseDefault, vbUseDefault, vbUseDefault)
            For j = ini + 1 To fin - 1
                session.findById("wnd[0]/usr/txtRF05A-ANZAZ").SetFocus
                session.findById("wnd[0]").sendVKey 2
                session.findById("wnd[1]/usr/txt*BSEG-BUZEI").text = j
                session.findById("wnd[1]/tbar[0]/btn[13]").press
                session.findById("wnd[0]/usr/txtBSEG-ZUONR").text = asignacion
                session.findById("wnd[0]/usr/ctxtBSEG-SGTXT").text = "PAG. #NombreCliente " & NPag & " VTO. " & vencimiento
                session.findById("wnd[0]").sendVKey 0
                session.findById("wnd[0]/tbar[1]/btn[14]").press
            Next j
            session.findById("wnd[0]/mbar/menu[2]/menu[6]").Select
            session.findById("wnd[0]/usr/txtBSEG-ZUONR").text = asignacion
            session.findById("wnd[0]/usr/ctxtBSEG-SGTXT").text = "PAG. #NombreCliente " & NPag & " VTO. " & vencimiento
            session.findById("wnd[0]/tbar[1]/btn[14]").press
            'Ask the user to take a look to the results and if they seems correct to apply the accounting entry
            fin = MessageBox(&H0, "Comprueba los apuntes.¿Quieres aplicar el pago?", "CONFIRMACION", vbYesNo)
            If fin = vbNo Then
                MsgBox ("Se cancela el proceso sin aplicar el pago")
                Exit Sub
            End If
            session.findById("wnd[0]/tbar[0]/btn[11]").press
            session.findById("wnd[0]/tbar[0]/btn[15]").press
            session.findById("wnd[0]/tbar[0]/okcd").text = "fb03"
            session.findById("wnd[0]").sendVKey 0
            Nombre = session.findById("wnd[0]/usr/txtRF05L-BELNR").text
            session.findById("wnd[0]/tbar[0]/btn[15]").press
            RUTA = wbR.Path
            'Save the "Relacion de pago" trated by the program and with the SAP number of the Accounting entry
            wbR.SaveAs Filename:=RUTA & "\" & Nombre & " #NombreCliente " & importeReal & ".xlsx", FileFormat:=51
            MsgBox ("Se ha aplicado el pago con nº asieno " & Nombre & " Y se ha guardado el archivo")
            wbR.Close
            GoTo EndLine

Line2:
'Line to apply the diferences in the payment
            If diff > 1 Or diff < -1 Then
                If diff < 0 Then
                    diff = diff * (-1)
                    session.findById("wnd[0]").resizeWorkingPane 92, 30, False
                    session.findById("wnd[0]/usr/ctxtRF05A-NEWBS").text = "16"
                    session.findById("wnd[0]/usr/ctxtRF05A-NEWKO").text = cliente
                    session.findById("wnd[0]/usr/ctxtRF05A-NEWKO").SetFocus
                    session.findById("wnd[0]/usr/ctxtRF05A-NEWKO").caretPosition = 0
                    session.findById("wnd[0]").sendVKey 0
                    session.findById("wnd[0]/usr/txtBSEG-WRBTR").text = diff
                    session.findById("wnd[0]/usr/ctxtBSEG-GSBER").text = "li1e"
                    session.findById("wnd[0]/usr/txtBSEG-ZUONR").text = asignacion
                    Com = InputBox("Escribre el comentario de la diferencia")
                    session.findById("wnd[0]/usr/ctxtBSEG-SGTXT").text = Com
                    session.findById("wnd[0]/usr/ctxtBSEG-SGTXT").SetFocus
                    session.findById("wnd[0]/usr/ctxtBSEG-SGTXT").caretPosition = 0
                    session.findById("wnd[0]/tbar[1]/btn[14]").press 'Montañita
                    session.ActiveWindow.sendVKey 0
                    GoTo Line1
                ElseIf diff > 0 Then
                    session.findById("wnd[0]").resizeWorkingPane 92, 30, False
                    session.findById("wnd[0]/usr/ctxtRF05A-NEWBS").text = "06"
                    session.findById("wnd[0]/usr/ctxtRF05A-NEWKO").text = cliente
                    session.findById("wnd[0]/usr/ctxtRF05A-NEWKO").SetFocus
                    session.findById("wnd[0]/usr/ctxtRF05A-NEWKO").caretPosition = 0
                    session.findById("wnd[0]").sendVKey 0
                    session.findById("wnd[0]/usr/txtBSEG-WRBTR").text = diff
                    session.findById("wnd[0]/usr/ctxtBSEG-GSBER").text = "li1e"
                    session.findById("wnd[0]/usr/txtBSEG-ZUONR").text = asignacion
                    Com = InputBox("Escribre el comentario de la diferencia")
                    session.findById("wnd[0]/usr/ctxtBSEG-SGTXT").text = Com
                    session.findById("wnd[0]/usr/ctxtBSEG-SGTXT").SetFocus
                    session.findById("wnd[0]/usr/ctxtBSEG-SGTXT").caretPosition = 0
                    session.findById("wnd[0]/tbar[1]/btn[14]").press 'Montañita
                    session.ActiveWindow.sendVKey 0
                    GoTo Line1
                End If
            Else
                MsgBox ("No se ha ajustado por redondeo. Se pasa al siguiente pago")
                GoTo EndLine
            End If
            ElseIf diff < 0 Then
                session.findById("wnd[0]/usr/ctxtRF05A-NEWBS").text = "40"
                session.findById("wnd[0]/usr/ctxtRF05A-NEWKO").text = "6590101011"
                session.findById("wnd[0]").sendVKey 0
                session.findById("wnd[0]/usr/txtBSEG-WRBTR").text = diff
                session.findById("wnd[0]/usr/txtBSEG-ZUONR").text = asignacion
                session.findById("wnd[0]/usr/ctxtBSEG-SGTXT").text = "Dif. CONFIR. ECI " & NPag & " VTO. " & vencimiento
                session.findById("wnd[0]/tbar[1]/btn[14]").press
                session.findById("wnd[0]").sendVKey 0
                session.findById("wnd[1]/usr/ctxtCOBL-GSBER").text = "LI1E"
                session.findById("wnd[1]/usr/ctxtCOBL-KOSTL").text = "LCO3710020"
                session.findById("wnd[1]").sendVKey 0
                GoTo Line1
            ElseIf diff > 0 Then
                diff = diff * (-1)
                session.findById("wnd[0]/usr/ctxtRF05A-NEWBS").text = "50"
                session.findById("wnd[0]/usr/ctxtRF05A-NEWKO").text = "6590101011"
                session.findById("wnd[0]").sendVKey 0
                session.findById("wnd[0]/usr/txtBSEG-WRBTR").text = diff
                session.findById("wnd[0]/usr/txtBSEG-ZUONR").text = asignacion
                session.findById("wnd[0]/usr/ctxtBSEG-SGTXT").text = "Dif. CONFIR. ECI " & NPag & " VTO. " & vencimiento
                session.findById("wnd[0]/tbar[1]/btn[14]").press
                session.findById("wnd[0]").sendVKey 0
                session.findById("wnd[1]/usr/ctxtCOBL-GSBER").text = "LI1E"
                session.findById("wnd[1]/usr/ctxtCOBL-KOSTL").text = "LCO3710020"
                session.findById("wnd[1]").sendVKey 0
                GoTo Line1
            End If
EndLine:
    Kill fichero
    End Sub


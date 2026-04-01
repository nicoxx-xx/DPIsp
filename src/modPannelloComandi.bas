Attribute VB_Name = "modPannelloComandi"
' ==============================
' Copyright (C) 2026 Domenico Longo
'
' This program is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation, version 3.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with this program.  If not, see <https://www.gnu.org/licenses/>.
' ==============================

Option Explicit

Public Const CREDITS As String = "Under GNU GPLv3 (see LICENSE_GPL), Copyright (c) 2026 Domenico Longo"
Public Const APPVER As String = "v1.0.1"
Public Const MOD_BUTTON_PANEL_VERSION As String = "v2.7.0"
Public Const MOD_MOUSESCROLL_VERSION As String = "v1.0.8"

Public gRibbon As Object ' popolato da OnRibbonLoad

' =====================================================
'   VERSIONING
' =====================================================
Public Function GetAppVersion() As String
    GetAppVersion = "DPIsp " & APPVER & ";"
End Function

Public Function GetPanelGenEngineVersion() As String
    GetPanelGenEngineVersion = "command panel " & MOD_BUTTON_PANEL_VERSION & ";"
End Function

Public Function GetMouseScrollEngineVersion() As String
    GetMouseScrollEngineVersion = _
    "Under MIT License (see LICENSE_MIT), Copyright (c) 2019 Ion Cristian Buse" & vbCrLf & _
    "VBA UserForm MouseScroll " & MOD_MOUSESCROLL_VERSION & "; "
End Function

Public Function GetCredits() As String
    GetCredits = CREDITS
End Function

' === ENTRYPOINT: crea/rigenera il pannello ===
Public Sub CreaPannelloComandi()
    Dim sh As Worksheet
    Dim shp As Shape

    On Error Resume Next
    Set sh = ThisWorkbook.Worksheets("Pannello")
    On Error GoTo 0

    If sh Is Nothing Then
        Set sh = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        sh.Protect DrawingObjects:=True
        sh.Name = "Pannello"
    End If

    UnlockSheet sh  'utile solo se il pannello viene rigenerato dal pannello macro
    
    ' Pulisci
    For Each shp In sh.Shapes
        shp.Delete
    Next shp
    sh.Cells.ClearFormats
    sh.Cells.ClearContents

    ' Nascondi griglia nella finestra attiva (se possibile)
    On Error Resume Next
    sh.Activate
    If Not Application.ActiveWindow Is Nothing Then
        Application.ActiveWindow.DisplayGridlines = False
    ElseIf ThisWorkbook.Windows.count > 0 Then
        ThisWorkbook.Windows(1).DisplayGridlines = False
    End If
    On Error GoTo 0

    ' Titolo (usa EN DASH U+2013)
    Set shp = sh.Shapes.AddTextbox(msoTextOrientationHorizontal, 38, 18, 565, 42)
    shp.Name = "TitoloPannello"
    shp.Locked = True
    shp.line.Visible = msoFalse
    With shp.TextFrame2.TextRange
        .text = "Pannello di Controllo " & ChrW(8211) & " Ispezioni DPI (DPIsp)"
        .Font.Size = 20
        .Font.Bold = True
        .ParagraphFormat.Alignment = msoAlignCenter
    End With

    ' Cornic1
    Dim buttbox As Shape
    Set buttbox = sh.Shapes.AddShape(msoShapeRoundedRectangle, 20, 60, 600, 520)
    With buttbox
        .Locked = True
        .Name = "BoxComandi"
        .Fill.ForeColor.RGB = RGB(247, 247, 247)
        .line.ForeColor.RGB = RGB(210, 210, 210)
        On Error Resume Next: .Adjustments.Item(1) = 0.07: On Error GoTo 0
    End With
    
    Dim helpbox As Shape
    Set helpbox = sh.Shapes.AddShape(msoShapeRoundedRectangle, 650, 60, 600, 520)
    With helpbox
        .Locked = True
        .Name = "HelpBox"
        .Fill.ForeColor.RGB = RGB(247, 247, 247)
        .line.ForeColor.RGB = RGB(210, 210, 210)
        On Error Resume Next: .Adjustments.Item(1) = 0.07: On Error GoTo 0
    End With
    
    ' Pulsanti a sinistra (colonna 1)
    AddButton sh, 50, 88, 260, 44, "Rigenera Schede PDF", "EsportaPDF_perRiga", RGB(0, 113, 188)
    AddButton sh, 50, 144, 260, 44, "Rigenera Layout PDF", "CreaLayoutMockup", RGB(0, 158, 73)
    AddButton sh, 50, 200, 260, 44, "Rigenera Pannello Comandi", "CreaPannelloComandi", RGB(220, 0, 120) ' RGB(94, 94, 94)
    AddButton sh, 50, 256, 260, 44, "Rigenera Azioni Ispettive per DPI", "AggiornaDatiDaAzioniDPI", RGB(127, 0, 255)
    AddButton sh, 50, 312, 260, 44, "Incrementa Anno (Date + Next Inspection Date)", "IncrementaAnnoDate", RGB(20, 200, 90)
    AddButton sh, 50, 368, 260, 44, "Form inserimento/modifica dati", "MostraGestioneDPI", RGB(20, 90, 220)


    ' Pulsanti a destra (colonna 2)
    AddButton sh, 330, 88, 260, 44, "Apri cartella PDF", "ApriCartellaPDF", RGB(255, 140, 0)
    AddButton sh, 330, 144, 260, 44, "Esporta da foglio Dati in .xlsx", "Esporta_tblDati_in_XLSX", RGB(189, 16, 224)
    AddButton sh, 330, 200, 260, 44, "Importa in foglio Dati da .xlsx", "Importa_tblDati_da_XLSX", RGB(0, 153, 255)
    AddButton sh, 330, 256, 260, 44, "Esporta fogli ausiliari in .xlsx", "EsportaFogliAuxInXLSX", RGB(255, 99, 71)
    AddButton sh, 330, 312, 260, 44, "Importa fogli ausiliari da .xlsx", "ImportaFogliAuxDaXLSX", RGB(140, 80, 250)
    AddButton sh, 330, 368, 260, 44, "Blocca tutti i fogli con password", "BloccaFogli", RGB(215, 20, 150)
    
    ' ---- Separatore ----
    'Dim sep As Shape
    'Set sep = sh.Shapes.AddLine(40, 320, 580, 320)
    'sep.line.ForeColor.RGB = RGB(180, 180, 180)
    
    ' Help (bullet via U+2022)
    Dim helpTxt As Shape
    Set helpTxt = sh.Shapes.AddTextbox(msoTextOrientationHorizontal, 670, 80, 560, 380)
    helpTxt.Name = "Help"
    helpTxt.Locked = True
    helpTxt.line.Visible = msoFalse
    helpTxt.Fill.ForeColor.RGB = RGB(247, 247, 247)
    With helpTxt.TextFrame2
        .MarginLeft = 6: .MarginRight = 6: .MarginTop = 2: .MarginBottom = 2
        With .TextRange
            .text = "Suggerimenti:" & vbCrLf & _
                    ChrW(8226) & " Imposta OutputFolder e FileNamePattern per le schede PDF, in 'Impostazioni' (OutputFolder può essere un path assoluto o relativo)." & vbCrLf & _
                    ChrW(8226) & " Imposta i dati Ispettore in 'Impostazioni'." & vbCrLf & _
                    ChrW(8226) & " Tutti i fogli sono bloccati con la password scritta in 'Impostazioni' (tranne il foglio 'Pannello'). Il programma gestisce in automatico sblocco/blocco." & vbCrLf & _
                    ChrW(8226) & " Solo in caso di modifiche manuali sui fogli, questi vanno sbloccati con la password scritta in 'Impostazioni'." & vbCrLf & _
                    ChrW(8226) & " Filtra la tabella 'Dati' per generare schede PDF solo delle righe visibili." & vbCrLf & _
                    ChrW(8226) & " Se la data di prossima ispezione non deve essere inserita (p.e. se un DPI non ha superato il controllo), inserire 'nnn'" & vbCrLf & _
                    ChrW(8226) & " Nella colonna 'Result', le schede con esito positivo vanno marcate con 'ok', quelle con esito negativo, con 'ko'" & vbCrLf & _
                    ChrW(8226) & " La funzione 'Rigenera Schede PDF' effettua dei controlli formali sulle date e su campi vuoti." & vbCrLf & _
                    ChrW(8226) & " Per assegnare le attività ispettive per un nuovo DPI inserito in tebella, assegnare la tipologia nella colonna 'SCHEDA'" & vbCrLf & _
                                 " secondo le tipologie di DPI previste nel foglio 'Azioni_DPI'; basta inserire solo il numero iniziale della colonna ID." & vbCrLf & _
                                 " Utilizzare poi il pulsante 'Rigenera Azioni Ispettive per DPI' per generare e assegnare l'elenco delle azioni ispettive per tutti i DPI della tebella." & vbCrLf & _
                    ChrW(8226) & " Utilizzare le funzioni di import/export sia per il foglio Dati che per i fogli ausiliari, per fare copie di riserva o per" & vbCrLf & _
                                 " trasferire le informazioni su una nuova versione dell'applicativo." & vbCrLf & _
                    ChrW(8226) & " Utilizzare la funzione 'blocca fogli' per bloccare in automatico tutti i fogli (tranne 'Pannello')." & vbCrLf & _
                    ChrW(8226) & " Utilizzare la funzione 'Incrementa anno' per incrementare l'anno delle colonne 'Date' e 'Next inpsection Date'." & vbCrLf & _
                                 " Questa funzione è necessaria per predisporre la stampa delle schede a partire dai dati dell'ultima ispezione." & vbCrLf & _
                    ChrW(8226) & " Il form di inserimento/modifica dati (CRUD) permette di creare o modificare record sul foglio Dati." & vbCrLf & _
                                 " Il form effettua una serie di verifiche formali sui dati, sia separatamente che durante il processo di salvataggio del record." & vbCrLf & _
                                 " Permette di verificare l'esistenza di duplicati sul foglio Dati. E' possibile effettuare eliminazioni di un record." & vbCrLf & _
                    ChrW(8226) & " I fogli non possono essere rinominati o cancellati per nessun motivo." & vbCrLf & _
                    ChrW(8226) & " Leggere anche le note presenti nei singoli fogli." & vbCrLf & _
                                 " " & vbCrLf & _
                                 " " & vbCrLf & _
                                 "  ==> ATTENZIONE: l'applicativo genera in automatico documenti PDF firmati con firma olografa dell'Ispettore, che ha valore legale. La responsabilità della verifica della correttezza dei dati del documento PDF, rimane all'Ispettore. <=="
            .Font.Size = 10
            .ParagraphFormat.Alignment = msoAlignLeft
        End With
    End With

    ' Definisci i delimitatori
    Const OPEN_DELIM As String = "==>"
    Const CLOSE_DELIM As String = "<=="
    
    Dim helptext As String
    Dim startPos As Long, endPos As Long, formatLen As Long

    helptext = helpTxt.TextFrame2.TextRange.text
    
    ' Trova posizioni dei delimitatori
    startPos = InStr(1, helptext, OPEN_DELIM, vbTextCompare)
    startPos = startPos - 1 '+ Len(OPEN_DELIM)
    endPos = InStr(startPos + 1, helptext, CLOSE_DELIM, vbTextCompare)
    

    ' Verifica che entrambi esistano
    If startPos > 0 And endPos > startPos Then
        
        ' Calcolo della lunghezza del testo da formattare (esclusi i delimitatori)
        formatLen = ((endPos + Len(CLOSE_DELIM)) - startPos - 1)
        
        If formatLen > 0 Then
            ' Applica la formattazione
            With helpTxt.TextFrame2.TextRange.Characters(startPos + 1, formatLen).Font
                .Bold = msoTrue
                .Size = 14
                .Fill.ForeColor.RGB = RGB(255, 0, 0)
            End With
        End If
    Else
        'MsgBox "Delimitatori non trovati."
    End If

    
    ' credits
    Dim verTxt As Shape
    Set verTxt = sh.Shapes.AddTextbox(msoTextOrientationHorizontal, 670, 465, 560, 100)
    verTxt.Name = "versioning"
    verTxt.Locked = True
    verTxt.line.Visible = msoFalse
    verTxt.Fill.ForeColor.RGB = RGB(247, 247, 247)
    With verTxt.TextFrame2
        .MarginLeft = 2: .MarginRight = 2: .MarginTop = 2: .MarginBottom = 2
        With .TextRange
            .text = "Excel version: " & Application.Version & "; " & GetAppVersion() & vbCrLf & _
                    "modules version:" & vbCrLf & _
                    GetCredits() & vbCrLf & _
                    GetPdfExportEngineVersion() & " " & GetDataExportEngineVersion() & " " & GetImpExpAuxSheetVersion() & vbCrLf & _
                    GetLayoutGenEngineVersion() & " " & GetDataImportEngineVersion() & vbCrLf & _
                    GetPanelGenEngineVersion() & " " & GetInputFormPanelVersion() & " " & GetDPIActionBuilderVersion() & vbCrLf & _
                    vbCrLf & _
                    GetMouseScrollEngineVersion()
            .Font.Size = 8
            .ParagraphFormat.Alignment = msoAlignLeft
        End With
    End With

    ' logo
    Dim logoT As Shape
    Dim logoTPath As String
    Dim target As Shape, pic As Shape
    Dim sx As Double, sy As Double, scaleF As Double
    
    logoTPath = ThisWorkbook.path & "\LogoCNSAS.png"
    If Dir(logoTPath) = "" Then Exit Sub
    Set logoT = sh.Shapes.AddShape(msoShapeRectangle, 1130, 460, 120, 120): logoT.Name = "LogoThick"
    
    On Error Resume Next
        Set target = sh.Shapes("LogoThick")
    On Error GoTo 0
    
    If Not target Is Nothing Then
        Set pic = sh.Shapes.AddPicture(logoTPath, msoFalse, msoTrue, target.left, target.top, -1, -1)
        pic.LockAspectRatio = msoTrue
        sx = target.Width / pic.Width: sy = target.Height / pic.Height
        scaleF = IIf(sx < sy, sx, sy): If scaleF < 1 Then pic.Width = pic.Width * scaleF
        pic.left = target.left + (target.Width - pic.Width) / 2
        pic.top = target.top + (target.Height - pic.Height) / 2
        target.Visible = msoFalse: pic.Name = "LogoThickImg"

        With pic.Shadow
            .Visible = msoTrue
            .Style = msoShadowStyleOuterShadow
            .Blur = 12          ' morbidezza
            .OffsetX = 3        ' spostamento orizzontale
            .OffsetY = 3        ' spostamento verticale
            .Transparency = 0.2 ' 0 = opaco, 1 = invisibile
        End With
        
        With pic.ThreeD
            .Visible = msoTrue
            .BevelTopType = msoBevelRelaxedInset ' altri: msoBevelCircle, msoBevelSoftRound, ecc.
            .BevelTopInset = 4                   ' spessore del bordo (pt)
            .BevelTopDepth = 6                   ' profondità del rilievo (pt)
            .Depth = 0                           ' corpo 3D (lascia 0 per foto “piatta” con bordo rilievo)
            .PresetLightingSoftness = 3          ' luce morbida
        End With
        
    End If
    
    sh.Range("A1").Select
    MsgBox "Pannello comandi creato nel foglio 'Pannello'.", vbInformation
CleanExit:
End Sub

' === Utility per creare pulsanti ===
Private Sub AddButton(ByVal sh As Worksheet, _
                      ByVal x As Single, ByVal y As Single, _
                      ByVal W As Single, ByVal H As Single, _
                      ByVal testo As String, ByVal macroName As String, _
                      ByVal colorRGB As Long)
    Dim b As Shape
    Set b = sh.Shapes.AddShape(msoShapeRoundedRectangle, x, y, W, H)
    With b
        .Locked = True
        .Name = "Btn_" & Replace(testo, " ", "_")
        .Fill.ForeColor.RGB = colorRGB
        .line.Visible = msoFalse
        On Error Resume Next: .Adjustments.Item(1) = 0.2: On Error GoTo 0
        With .TextFrame2
            .MarginLeft = 6: .MarginRight = 6: .MarginTop = 4: .MarginBottom = 4
            With .TextRange
                .text = testo
                .Font.Size = 12
                .Font.Bold = True
                .Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
                .ParagraphFormat.Alignment = msoAlignCenter
            End With
        End With
        .OnAction = macroName
    End With
End Sub

' === Apri cartella PDF (senza dipendenze esterne) ===
Public Sub ApriCartellaPDF()
    Dim p As String
    p = ReadOutputFolder()
    'If Len(p) = 0 Then p = ThisWorkbook.Path
    p = ResolveExportPath(p)
    p = NormalizePathLite(p)
    If Right$(p, 1) = "\" Or Right$(p, 1) = "/" Then p = left$(p, Len(p) - 1)
    Shell "explorer.exe " & Chr$(34) & p & Chr$(34), vbNormalFocus
End Sub

Private Function ReadOutputFolder() As String
    On Error Resume Next
    Dim sh As Worksheet
    Dim f As Range
    Set sh = ThisWorkbook.Worksheets("Impostazioni")
    If Not sh Is Nothing Then
        Set f = sh.Columns(1).Find(What:="OutputFolder", LookAt:=xlWhole, LookIn:=xlValues)
        If Not f Is Nothing Then ReadOutputFolder = CStr(f.Offset(0, 1).value)
    End If
End Function

Private Function NormalizePathLite(ByVal p As String) As String
    Dim s As String
    s = Trim$(p)
    If Len(s) = 0 Then NormalizePathLite = "": Exit Function
    s = Replace(s, "/", "\")
    Do While InStr(s, "\\\\") > 0 And left$(s, 2) <> "\\\\"
        s = Replace(s, "\\\\", "\")
    Loop
    NormalizePathLite = s
End Function

Private Function ResolveExportPath(pathText As String) As String
    Dim T As String
    Dim basePath As String
    
    T = Trim(pathText)
    basePath = ThisWorkbook.path
    
    ' 1) Cella vuota ? path assoluto del file xlsm
    If Len(T) = 0 Then
        ResolveExportPath = basePath
        Exit Function
    End If
    
    ' Normalizza separatori per sicurezza minima
    T = Replace(T, "/", "\")
    
    ' 2) Path assoluto (drive letter tipo C:\ oppure D:\ ecc.)
    If Len(T) > 2 Then
        If Mid$(T, 2, 1) = ":" Then
            ResolveExportPath = T
            Exit Function
        End If
    End If
    
    ' 3) Path relativo
    ' Se inizia con "\" (equivalente a "/"), rimuovilo
    If left$(T, 1) = "\" Then
        T = Mid$(T, 2)
    End If
    
    ' Combina base + relativo
    ResolveExportPath = basePath & "\" & T
End Function

' === Controllo campi obbligatori ===
Public Function ControllaCampiObbligatori() As Boolean
    Dim sh As Worksheet
    Dim lo As ListObject
    Dim must() As Variant
    Dim miss As String
    Dim r As ListRow
    Dim i As Long
    Dim v As Variant

    Set sh = ThisWorkbook.Worksheets("Dati")
    On Error Resume Next
    Set lo = sh.ListObjects("tblDati")
    On Error GoTo 0
    If lo Is Nothing Then
        MsgBox "Tabella 'tblDati' non trovata.", vbExclamation, "Check error"
        Exit Function
    End If

    must = Array("Number", "Manufacturer", "Model", "Serial Number", "Date of Manufacture", "Date of Purchase", _
    "Date of First Use", "Next Ispection Date", "Date for retirement", "Required inspection activities")
    miss = ""

    For Each r In lo.ListRows
        If (Not r.Range.EntireRow.Hidden) And (Application.WorksheetFunction.CountIf(r.Range, "<>") <> 0) Then ' non verificare righe vuote
            For i = LBound(must) To UBound(must)
                v = GetValue(lo, r, CStr(must(i)))
                If Len(Trim$(CStr(v))) = 0 Then
                    miss = miss & "Riga " & (r.Index + 1) & ": manca -> " & CStr(must(i)) & vbCrLf   ' header row is not in ListRows
                End If
            Next i
        End If
    Next r

    If Len(miss) = 0 Then
        ControllaCampiObbligatori = True
        'MsgBox "OK: campi obbligatori presenti nelle righe visibili.", vbInformation
    Else
        ThisWorkbook.Worksheets("Dati").Activate
        ControllaCampiObbligatori = False
        MsgBox "Anomalie rilevate:" & vbCrLf & miss, vbExclamation, "Check error"
    End If
End Function

Private Function GetValue(ByVal lo As ListObject, ByVal r As ListRow, ByVal colName As String) As Variant
    Dim lc As ListColumn
    On Error Resume Next
    Set lc = lo.ListColumns(colName)
    If lc Is Nothing Then
        GetValue = ""
    Else
        GetValue = r.Range.Cells(1, lc.Index).value
    End If
End Function

' ======================
' VALIDAZIONE DATE FORMALI
' ======================
Public Function VerificaDateFormali() As Boolean
    Dim sh As Worksheet
    Dim lo As ListObject
    Dim r As ListRow
    Dim cols As Variant
    Dim rep As String
    Dim issues As Long

    Set sh = ThisWorkbook.Worksheets("Dati")
    On Error Resume Next
    Set lo = sh.ListObjects("tblDati")
    On Error GoTo 0
    If lo Is Nothing Then
        MsgBox "Tabella 'tblDati' non trovata.", vbExclamation, "Check error"
        Exit Function
    End If

    cols = Array("Date", "Date of Manufacture", "Date of Purchase", "Date of First Use", "Next Ispection Date", "Date for retirement")

    UnlockSheet sh
    
    ExcelLock
    
    ClearDateHighlights lo, cols

    For Each r In lo.ListRows
        If (Not r.Range.EntireRow.Hidden) And (Application.WorksheetFunction.CountIf(r.Range, "<>") <> 0) Then      ' non verificare righe vuote
            Dim i As Long
            Dim colName As String
            For i = LBound(cols) To UBound(cols)
                colName = CStr(cols(i))
                If Not IsValidDateCell(lo, r, colName) Then
                    HighlightCell lo, r, colName, RGB(255, 199, 206)
                    ' header row is not in ListRows
                    rep = rep & "Riga " & (r.Index + 1) & ": data non valida in '" & colName & "' -> '" & CStr(GetValue(lo, r, colName)) & "'" & vbCrLf
                    issues = issues + 1
                End If
            Next i
        End If
    Next r

    ExcelUnlock

    LockSheet sh
    
    If issues = 0 Then
        VerificaDateFormali = True
        'MsgBox "OK: tutte le date nelle righe visibili sono formalmente valide.", vbInformation
    Else
        ThisWorkbook.Worksheets("Dati").Activate
        VerificaDateFormali = False
        MsgBox IIf(Abs(issues) = 1, "E' stata rilevata [", "Sono state rilevate [") & issues & "] " & IIf(Abs(issues) = 1, "data non valida.", "date non valide.") & vbCrLf & _
               "Dettagli:" & vbCrLf & rep, vbExclamation, "Check completato"
    End If
End Function

Private Sub ClearDateHighlights(ByVal lo As ListObject, ByVal cols As Variant)
    Dim i As Long
    Dim lc As ListColumn
    Dim r As Long
    For i = LBound(cols) To UBound(cols)
        On Error Resume Next
        Set lc = lo.ListColumns(CStr(cols(i)))
        If Not lc Is Nothing Then
            For r = 1 To lo.ListRows.count
                lc.DataBodyRange.Cells(r, 1).Interior.ColorIndex = xlColorIndexNone
            Next r
        End If
        On Error GoTo 0
    Next i
End Sub

Private Sub HighlightCell(ByVal lo As ListObject, ByVal r As ListRow, ByVal colName As String, ByVal clr As Long)
    Dim lc As ListColumn
    On Error Resume Next
    Set lc = lo.ListColumns(colName)
    If Not lc Is Nothing Then lc.DataBodyRange.Cells(r.Index, 1).Interior.Color = clr
    On Error GoTo 0
End Sub

Private Function IsValidDateCell(ByVal lo As ListObject, ByVal r As ListRow, ByVal colName As String) As Boolean
    Dim v As Variant
    Dim d As Variant
    v = GetValue(lo, r, colName)
    If IsEmpty(v) Or Len(Trim$(CStr(v))) = 0 Then
        IsValidDateCell = False
        Exit Function
    End If
    d = TryParseDate(v)
    IsValidDateCell = Not IsEmpty(d) Or (v = "nnn")
End Function

Private Function TryParseDate(ByVal v As Variant) As Variant
    On Error GoTo Fail

    If IsNumeric(v) Then
        GoTo Fail
    End If

    Dim s As String
    
    s = Trim$(CStr(v))
    If Len(s) = 0 Then GoTo Fail

    s = Replace(s, "-", "/")
    s = Replace(s, ".", "/")

    If Not IsDate(s) Then GoTo Fail ' verifica validità data di anni bisestili, a meno di altri errori nei campi data

    Dim parts() As String
    parts = Split(s, "/")
    If UBound(parts) = 2 Then
        Dim dd As Long, mm As Long, yy As Long
        If IsNumeric(parts(0)) And IsNumeric(parts(1)) And IsNumeric(parts(2)) Then
            dd = CLng(parts(0)): mm = CLng(parts(1)): yy = CLng(parts(2))
            If yy < 100 And yy >= 0 Then yy = 2000 + yy
            If mm >= 1 And mm <= 12 And dd >= 1 And dd <= 31 And yy >= 1900 And yy <= 2200 Then
                TryParseDate = DateSerial(yy, mm, dd)
                Exit Function
            Else
                GoTo Fail
            End If
        Else
          GoTo Fail
        End If
    Else
        GoTo Fail
    End If

Fail:
    TryParseDate = Empty
End Function

' ======================
' INCREMENTO ANNO DATE (Date, Next Ispection Date)
' ======================
Public Sub IncrementaAnnoDate()
    Dim sh As Worksheet
    Dim lo As ListObject
    Dim r As ListRow
    Dim addYears As Long
    Dim resp As Variant

    resp = InputBox("Di quanti anni incrementare il campo 'Date' e 'Next Ispection Date'?", _
                    "Incrementa Anno", 1)
    If VarType(resp) = vbString Then
        If Trim$(resp) = "" Then Exit Sub
        If Not IsNumeric(resp) Then
            MsgBox "Valore non numerico.", vbExclamation, "Input error"
            Exit Sub
        End If
        addYears = CLng(resp)
    Else
        addYears = 1
    End If

    Set sh = ThisWorkbook.Worksheets("Dati")
    On Error Resume Next
    Set lo = sh.ListObjects("tblDati")
    On Error GoTo 0
    If lo Is Nothing Then
        MsgBox "Tabella 'tblDati' non trovata.", vbExclamation, "Check error"
        Exit Sub
    End If

    UnlockSheet sh

    Dim targetCols As Variant
    targetCols = Array("Date", "Next Ispection Date")

    ExcelLock

    Dim i As Long
    Dim colName As String
    For Each r In lo.ListRows
        If Not r.Range.EntireRow.Hidden Then
            For i = LBound(targetCols) To UBound(targetCols)
                colName = CStr(targetCols(i))
                AddYearsToCell lo, r, colName, addYears
            Next i
        End If
    Next r

    ExcelUnlock
    
    LockSheet sh

    MsgBox "Anno incrementato di " & addYears & " per le colonne 'Date' e 'Next Ispection Date' (righe visibili).", vbInformation, "Incremento date completato"
End Sub

Private Sub AddYearsToCell(ByVal lo As ListObject, ByVal r As ListRow, ByVal colName As String, ByVal nYears As Long)
    Dim lc As ListColumn
    Dim v As Variant
    Dim d As Variant

    On Error Resume Next
    Set lc = lo.ListColumns(colName)
    On Error GoTo 0
    If lc Is Nothing Then Exit Sub

    v = lc.DataBodyRange.Cells(r.Index, 1).value
    d = TryParseDate(v)
    If Not IsEmpty(d) Then
        lc.DataBodyRange.Cells(r.Index, 1).value = SafeAddYears(d, nYears)
    End If
End Sub

Private Function SafeAddYears(ByVal d As Date, ByVal nYears As Long) As Date
    Dim yy As Long, mm As Long, dd As Long
    Dim lastDay As Long
    yy = Year(d) + nYears
    mm = Month(d)
    dd = Day(d)
    lastDay = Day(DateSerial(yy, mm + 1, 0))
    If dd > lastDay Then dd = lastDay
    SafeAddYears = DateSerial(yy, mm, dd)
End Function

' ======================
' Check sheet protection and lock/unlock
' ======================
Private Function IsSheetProtected(ws As Worksheet) As Boolean
    IsSheetProtected = (ws.ProtectContents Or _
                        ws.ProtectDrawingObjects Or _
                        ws.ProtectScenarios)
End Function

Sub ExcelLock()
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
    End With
End Sub

Sub ExcelUnlock()
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
    End With
    
    DoEvents

    If Not Application.ActiveWindow Is Nothing Then
        With Application.ActiveWindow
            .SmallScroll Down:=1
            .SmallScroll Up:=1
        End With
    End If
    
End Sub

Public Sub LockSheet(ws As Worksheet)
    Dim pwd As String: pwd = GetImpostazione("Password_fogli")
    
    ws.EnableSelection = xlNoRestrictions
    If Not IsSheetProtected(ws) Then
        ws.Protect Password:=pwd, _
        UserInterfaceOnly:=True, _
        AllowFiltering:=True, _
        AllowSorting:=True
    End If
End Sub

Public Sub UnlockSheet(ws As Worksheet)
    Dim pwd As String: pwd = GetImpostazione("Password_fogli")
    
    If IsSheetProtected(ws) Then
        ' MsgBox "Il foglio del layout ('Layout') è protetto da password e le stampe PDF non possono essere create.", vbCritical
        ' Exit Sub
        ws.Unprotect Password:=pwd
    End If
End Sub

Public Sub SostituisciCodiciUbicazione()

    Dim wsDati As Worksheet
    Dim wsUbicazioni As Worksheet
    Dim lo As ListObject
    Dim dict As Object
    Dim r As ListRow
    Dim codice As String
    Dim colUbic As Long
    
    '--- fogli
    Set wsDati = ThisWorkbook.Worksheets("Dati")
    Set wsUbicazioni = ThisWorkbook.Worksheets("Ubicazioni")
    
    '--- tabella dati
    Set lo = wsDati.ListObjects("tblDati")
    
    '--- trova la colonna 'Ubicazione' nella tabella
    On Error Resume Next
    colUbic = lo.ListColumns("Ubicazione").Index
    On Error GoTo 0
    
    If colUbic = 0 Then
        MsgBox "La tabella tblDati non contiene una colonna chiamata 'Ubicazione'.", vbCritical
        Exit Sub
    End If
    
    '--- Dictionary: codice ? ubicazione estesa
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1   'TextCompare
    
    Dim lastRow As Long
    Dim i As Long
    Dim codiceFoglio As String
    Dim esteso As String
    
    ExcelLock
    
    '--- ultima riga foglio Ubicazioni
    lastRow = wsUbicazioni.Cells(wsUbicazioni.Rows.count, "A").End(xlUp).row
    
    '--- carica codici e ubicazioni estese
    ' Colonna A = ID? (ignorata)
    ' Colonna B = Ubicazione (codice)
    ' Colonna C = UbicazioneEstesa
    For i = 2 To lastRow
        codiceFoglio = Trim(wsUbicazioni.Cells(i, "B").value)
        esteso = wsUbicazioni.Cells(i, "C").value
        
        If Len(codiceFoglio) > 0 Then
            dict(codiceFoglio) = esteso
        End If
    Next i
    
    '--- scorre tutte le righe della tabella (anche filtrate)
    UnlockSheet wsDati
    For Each r In lo.ListRows
        
        codice = Trim(r.Range.Cells(1, colUbic).value)
        codice = UCase(codice)
        
        'se il codice esiste nel dizionario ? sostituisci
        If dict.exists(codice) Then
            r.Range.Cells(1, colUbic).value = dict(codice)
        End If
        
        'se NON esiste ? lascia il valore così com'è
    Next r
    
    LockSheet wsDati
    ExcelUnlock
    
    'MsgBox "Sostituzione ubicazioni completata!", vbInformation
End Sub

Sub BloccaFogli()

    Dim ws As Worksheet
    Dim i As Long
    
    ' Ciclo su TUTTI i fogli
    For i = 1 To ThisWorkbook.Worksheets.count
    
        Set ws = ThisWorkbook.Worksheets(i)
        
        If (ws.Name <> "Pannello") Then
            LockSheet ws
        End If
        
        ' Mostra il nome
        'MsgBox "Foglio: " & ws.Name, vbInformation
        
    Next i
    MsgBox "Tutti i fogli, tranne 'Pannello', sono stati bloccati con password.", vbInformation
End Sub

' ======================
' RIBBON CALLBACKS (multilinea)
' ======================
Public Sub OnRibbonLoad(ByVal ribbon As Object)
    Set gRibbon = ribbon
End Sub

Public Sub rb_GeneraPDF(ByVal control As Object)
    EsportaPDF_perRiga
End Sub

Public Sub rb_RigeneraLayout(ByVal control As Object)
    CreaLayoutMockup
End Sub

'Public Sub rb_VaiDati(ByVal control As Object)
    'VaiDati
'End Sub

'Public Sub rb_VaiLayout(ByVal control As Object)
    'VaiLayout
'End Sub

Public Sub rb_ApriCartellaPDF(ByVal control As Object)
    ApriCartellaPDF
End Sub

Public Sub rb_ControllaObbligatori(ByVal control As Object)
    ControllaCampiObbligatori
End Sub

Public Sub rb_VerificaDate(ByVal control As Object)
    VerificaDateFormali
End Sub

Public Sub rb_IncrementaAnno(ByVal control As Object)
    IncrementaAnnoDate
End Sub

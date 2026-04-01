Attribute VB_Name = "modLayoutMockup"
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

Public Const MOD_LAYOUT_VERSION As String = "v3.4.1"
Public Const DOC_VERSION As String = "v0.1.0"


' =====================================================
'   VERSIONING
' =====================================================
Public Function GetLayoutGenEngineVersion() As String
    GetLayoutGenEngineVersion = "layout generator " & MOD_LAYOUT_VERSION & "; pdf layout " & DOC_VERSION & ";"
End Function


Public Sub CreaLayoutMockup()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim sh As Worksheet
    Dim shp As Shape
    Dim rng As Range
    Dim addr As String

    On Error Resume Next
    Set sh = wb.Worksheets("Layout")
    On Error GoTo 0

    If sh Is Nothing Then
        Set sh = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count))
        sh.Name = "Layout"
        Set rng = sh.Range("A1:L62")
        ' addr = rng.Address(True, True, xlA1, False)
        addr = rng.Address
    Else
        UnlockSheet sh
        
        For Each shp In sh.Shapes: shp.Delete: Next shp
        sh.Cells.ClearFormats: sh.Cells.ClearContents
    End If

    ExcelLock

    ' ---- Impostazioni pagina ----
    Application.PrintCommunication = False
    With sh.PageSetup
        .Orientation = xlPortrait
        .PaperSize = xlPaperA4
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .LeftMargin = Application.CentimetersToPoints(1#)
        .RightMargin = Application.CentimetersToPoints(1#)
        .TopMargin = Application.CentimetersToPoints(1#)
        .BottomMargin = Application.CentimetersToPoints(1#)
        .HeaderMargin = Application.CentimetersToPoints(0)
        .FooterMargin = Application.CentimetersToPoints(0)
        .PrintArea = addr
    End With
    Application.PrintCommunication = True

    ' ---- Misure utili ----
    Dim pageW As Single, pageH As Single
    pageW = Application.CentimetersToPoints(21 - 1.3 - 1.3)
    pageH = Application.CentimetersToPoints(29.7 - 1# - 1#)

    Dim x0 As Single, y0 As Single
    x0 = sh.PageSetup.LeftMargin
    y0 = sh.PageSetup.TopMargin
    
    ' ---- versione documento ----
    Dim dver As Shape
    Set dver = sh.Shapes.AddTextbox(msoTextOrientationHorizontal, x0 - 18, y0 - 18, pageW - 80, 20)
    dver.Name = "DocVer"
    dver.TextFrame2.TextRange.text = "CNSAS-SASS 2026; Layout " & DOC_VERSION & "; UNI_EN365_2005 compliant" & vbCrLf & _
                                     "Per avere validità, questa scheda deve essere compilata in tutte le sue parti e " & _
                                     "firmata dall’ispettore che effettua l’ispezione, indicando anche la relativa matricola."
    dver.line.Visible = msoFalse
    dver.TextFrame2.MarginTop = 0
    dver.TextFrame2.TextRange.Font.Size = 6
    dver.TextFrame2.TextRange.Font.Bold = False
    dver.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignLeft

    ' ---- Loghi placeholder ----
    Dim logoL As Shape, logoR As Shape
    Set logoL = sh.Shapes.AddShape(msoShapeRectangle, x0, y0 + 5, 90, 50): logoL.Name = "LogoLeft"
    Set logoR = sh.Shapes.AddShape(msoShapeRectangle, x0 + pageW - 90, y0 + 5, 90, 50): logoR.Name = "LogoRight"

    ' ---- Titolo ----
    Dim ttl As Shape
    Set ttl = sh.Shapes.AddTextbox(msoTextOrientationHorizontal, x0 + 100, y0 + 5, pageW - 200, 28)
    ttl.Name = "Titolo"
    ttl.TextFrame2.TextRange.text = "SCHEDA DI ISPEZIONE DPI "
    ttl.TextFrame2.TextRange.Font.Size = 16
    ttl.TextFrame2.TextRange.Font.Bold = True
    ttl.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter

    ' ---- Sottotitolo ----
    Dim subttl As Shape
    Set subttl = sh.Shapes.AddTextbox(msoTextOrientationHorizontal, x0 + 100, y0 + 33, pageW - 200, 18)
    subttl.Name = "Sottotitolo"
    subttl.TextFrame2.TextRange.text = "{{SCHEDA}} " & ChrW(8211) & " N" & ChrW(176) & " {{Number}} " & ChrW(8211) & " {{Date}}"
    subttl.TextFrame2.TextRange.Font.Size = 11
    subttl.TextFrame2.TextRange.Font.Italic = True
    subttl.TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter

    ' ---- Separatore ----
    Dim sep As Shape
    Set sep = sh.Shapes.AddLine(x0, y0 + 61, x0 + pageW, y0 + 61)
    sep.line.ForeColor.RGB = RGB(180, 180, 180)

    ' ---- Sezione Dati DPI ----
    Dim yBase As Single: yBase = y0 + 70
    Dim lblW As Single: lblW = 180
    Dim valW As Single: valW = pageW - 200
    Dim rH As Single: rH = 18

    Call AddRow(sh, x0, yBase + rH * 0, lblW, valW, "Modello", "{{Model}}")
    Call AddRow(sh, x0, yBase + rH * 1, lblW, valW, "Produttore", "{{Manufacturer}}")
    Call AddRow(sh, x0, yBase + rH * 2, lblW, valW, "Serial Number", "{{Serial Number}}")
    Call AddRow(sh, x0, yBase + rH * 3, lblW, valW, "Data di fabbricazione", "{{Date of Manufacture}}")
    Call AddRow(sh, x0, yBase + rH * 4, lblW, valW, "Data di acquisto", "{{Date of Purchase}}")
    Call AddRow(sh, x0, yBase + rH * 5, lblW, valW, "Prima messa in servizio", "{{Date of First Use}}")
    Call AddRow(sh, x0, yBase + rH * 6, lblW, valW, "Prossima ispezione", "{{Next Ispection Date}}")
    Call AddRow(sh, x0, yBase + rH * 7, lblW, valW, "Data ritiro prevista", "{{Date for retirement}}")

    ' ---- Fascia esito ----
    Dim es As Shape
    Set es = sh.Shapes.AddShape(msoShapeRectangle, x0, yBase + rH * 8.5, pageW, 28)
    es.Name = "EsitoBar"
    es.TextFrame2.TextRange.text = "Esito ispezione:  {{Result}}"
    es.TextFrame2.TextRange.Font.Bold = True
    es.TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(0, 0, 0)
    es.TextFrame2.TextRange.Font.Size = 12
    es.TextFrame2.MarginLeft = 10
    es.TextFrame2.MarginTop = 6

    ' ---- Pannelli in basso ----
    Dim bottomPad As Single: bottomPad = 8
    Dim firmaH As Single: firmaH = 120
    Dim firmaW As Single: firmaW = pageW * 0.42
    Dim firmaLeft As Single: firmaLeft = x0 + pageW - firmaW
    ' +120 per spostare il blocco finale a fondo pagina A4
    Dim firmaTop As Single: firmaTop = y0 + pageH - firmaH - bottomPad + 120

    ' ---- BOX ISPETTORE/FIRMA (sfondo bianco) ----
    Dim signBox As Shape
    Set signBox = sh.Shapes.AddShape(msoShapeRectangle, firmaLeft, firmaTop, firmaW, firmaH)
    signBox.Name = "FirmaBox"
    signBox.Fill.Visible = msoTrue
    signBox.Fill.ForeColor.RGB = RGB(230, 230, 230)
    signBox.line.ForeColor.RGB = RGB(200, 200, 200)

    Dim isp As Shape
    Set isp = sh.Shapes.AddTextbox(msoTextOrientationHorizontal, firmaLeft + 8, firmaTop + 8, firmaW - 16, 18)
    isp.TextFrame2.TextRange.text = "Ispezionato da: {{Ispettore}}"
    isp.TextFrame2.MarginTop = 1

    Dim matric As Shape
    Set matric = sh.Shapes.AddTextbox(msoTextOrientationHorizontal, firmaLeft + 8, firmaTop + 28, firmaW - 16, 18)
    matric.TextFrame2.TextRange.text = "Matricola: {{Matricola_Ispettore}}"
    matric.TextFrame2.MarginTop = 1

    Dim dat As Shape
    Set dat = sh.Shapes.AddTextbox(msoTextOrientationHorizontal, firmaLeft + 8, firmaTop + 48, firmaW - 16, 18)
    dat.TextFrame2.TextRange.text = "Data: {{Date}}"
    dat.TextFrame2.MarginTop = 1

    Dim firmaPh As Shape
    Set firmaPh = sh.Shapes.AddShape(msoShapeRectangle, firmaLeft + 8, firmaTop + 72, firmaW - 16, firmaH - 80)
    firmaPh.Name = "FirmaSegnaposto"
    With firmaPh
        .Fill.Visible = msoTrue
        .Fill.ForeColor.RGB = RGB(255, 255, 255) ' sfondo bianco (non blu)
        .line.ForeColor.RGB = RGB(200, 200, 200)
        .TextFrame2.TextRange.text = "FIRMA"
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .TextFrame2.TextRange.Font.Size = 10
        .TextFrame2.TextRange.Font.Bold = True
    End With

    ' ---- COLONNA SINISTRA (Customer + Annotazioni) ----
    Dim leftW As Single: leftW = pageW - firmaW - 12

    ' Customer a UNA riga (altezza ridotta)
    Dim cliH As Single: cliH = 20
    Dim cliTop As Single: cliTop = firmaTop
    Dim cli As Shape
    Set cli = sh.Shapes.AddTextbox(msoTextOrientationHorizontal, x0, cliTop, leftW, cliH)
    cli.Name = "Cliente"
    cli.line.ForeColor.RGB = RGB(230, 230, 230)
    With cli.TextFrame2
        .MarginLeft = 8: .MarginRight = 8: .MarginTop = 2: .MarginBottom = 2
        .TextRange.text = "Cliente / Stazione:  {{Customer}}"
        .TextRange.Font.Size = 11
        .TextRange.ParagraphFormat.Alignment = msoAlignLeft
    End With

    ' Nuovo: ANNOTAZIONI occupa lo spazio rimanente sotto Customer
    Dim annTop As Single: annTop = cliTop + cliH + 6
    Dim annH As Single: annH = (firmaTop + firmaH) - annTop
    If annH < 40 Then annH = 40  ' garantisci un minimo
    Dim ann As Shape
    Set ann = sh.Shapes.AddTextbox(msoTextOrientationHorizontal, x0, annTop, leftW, annH)
    ann.Name = "AnnotazioniBox"
    With ann.TextFrame2
        .MarginLeft = 8: .MarginRight = 8: .MarginTop = 6: .MarginBottom = 6
        .TextRange.text = "Annotazioni:" & vbCrLf & "{{Annotazioni}}"
        .TextRange.Font.Size = 11
        .TextRange.ParagraphFormat.Alignment = msoAlignLeft
    End With
    ann.line.ForeColor.RGB = RGB(220, 220, 220)

    ' ---- Attività  richieste (resta invariata, sopra i pannelli in basso) ----
    Dim attTop As Single: attTop = es.top + es.Height + 12
    Dim attH As Single: attH = firmaTop - attTop - bottomPad
    If attH < 80 Then attH = 80

    Dim att As Shape
    ' +7 / -7 per accomodare barra decorativa
    Set att = sh.Shapes.AddTextbox(msoTextOrientationHorizontal, x0 + 7, attTop, pageW - 7, attH)
    att.Name = "AttivitaRichieste"
    att.TextFrame2.TextRange.text = "Attivit" & ChrW(224) & " di ispezione richieste:" & vbCrLf & "{{Required inspection activities}}"

    ' Barra decorativa
    Dim bar As Shape
    Set bar = sh.Shapes.AddShape(msoShapeRectangle, x0, att.top, 6, att.Height)
    bar.line.Visible = msoFalse
    bar.Fill.ForeColor.RGB = RGB(255, 199, 44)

    ' === Inserimento loghi CNSAS ===
    Call InserisciLoghiCNSAS(sh)

    ' Area di stampa (facoltativa)
    ' print area già impostata in page setup
    ' sh.PageSetup.PrintArea = sh.Range("A1:F62").Address

    ThisWorkbook.Worksheets("Layout").Activate
    sh.Range("A1").Select
    
    LockSheet sh
    
    ExcelUnlock
    
    MsgBox "Layout creato nel foglio 'Layout'.", vbInformation
CleanExit:
End Sub

' ===============================
' Inserimento loghi CNSAS
' ===============================
Public Sub InserisciLoghiCNSAS(ByVal sh As Worksheet)
    Dim logoPath As String
    logoPath = ThisWorkbook.path & "\LogoCNSAS.png"
    If Dir(logoPath) = "" Then Exit Sub

    Dim target As Shape, pic As Shape
    Dim sx As Double, sy As Double, scaleF As Double

    ' --- sinistra ---
    On Error Resume Next: Set target = sh.Shapes("LogoLeft"): On Error GoTo 0
    If Not target Is Nothing Then
        Set pic = sh.Shapes.AddPicture(logoPath, msoFalse, msoTrue, target.left, target.top, -1, -1)
        pic.LockAspectRatio = msoTrue
        sx = target.Width / pic.Width: sy = target.Height / pic.Height
        scaleF = IIf(sx < sy, sx, sy): If scaleF < 1 Then pic.Width = pic.Width * scaleF
        pic.left = target.left + (target.Width - pic.Width) / 2
        pic.top = target.top + (target.Height - pic.Height) / 2
        target.Visible = msoFalse: pic.Name = "LogoLeftImg"
    End If

    ' --- destra ---
    On Error Resume Next: Set target = sh.Shapes("LogoRight"): On Error GoTo 0
    If Not target Is Nothing Then
        Set pic = sh.Shapes.AddPicture(logoPath, msoFalse, msoTrue, target.left, target.top, -1, -1)
        pic.LockAspectRatio = msoTrue
        sx = target.Width / pic.Width: sy = target.Height / pic.Height
        scaleF = IIf(sx < sy, sx, sy): If scaleF < 1 Then pic.Width = pic.Width * scaleF
        pic.left = target.left + (target.Width - pic.Width) / 2
        pic.top = target.top + (target.Height - pic.Height) / 2
        target.Visible = msoFalse: pic.Name = "LogoRightImg"
    End If
End Sub

' ===============================
' Utility: riga etichetta/valore
' ===============================
Private Sub AddRow(ByVal sh As Worksheet, ByVal x As Single, ByVal y As Single, _
                   ByVal wLbl As Single, ByVal wVal As Single, _
                   ByVal labelText As String, ByVal valueText As String)
    Dim lb As Shape, vl As Shape
    Set lb = sh.Shapes.AddTextbox(msoTextOrientationHorizontal, x, y, wLbl, 16)
    lb.TextFrame2.TextRange.text = labelText
    lb.TextFrame2.MarginTop = 1
    lb.Fill.ForeColor.RGB = RGB(230, 230, 230)
    lb.line.ForeColor.RGB = RGB(200, 200, 200)

    Set vl = sh.Shapes.AddTextbox(msoTextOrientationHorizontal, x + wLbl + 10, y, wVal, 16)
    vl.TextFrame2.TextRange.text = valueText
    vl.TextFrame2.MarginTop = 1
    vl.line.Visible = msoFalse
End Sub

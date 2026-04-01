Attribute VB_Name = "modExportPDF"
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

Public Const MOD_EXPORTPDF_VERSION As String = "v3.6.2"

' === CONFIG ===
Private Const NOME_TAB As String = "tblDati"
Private Const NOME_LAYOUT As String = "Layout"
Private Const NOME_IMP As String = "Impostazioni"
Private Const NOME_LOG As String = "Log"
Private Const NOME_AZIONI As String = "Azioni_Ispettive"

Private Const FIRMA_PNG As String = "Test_firma.png"
Private Const FIRMA_JPG As String = "Test_firma.jpg"

Private gAzioni As Object  ' Scripting.Dictionary

' =====================================================
'   VERSIONING
' =====================================================
Public Function GetPdfExportEngineVersion() As String
    GetPdfExportEngineVersion = "PDF export engine " & MOD_EXPORTPDF_VERSION & ";"
End Function


' =====================================================
'   LETTURA IMPOSTAZIONI
' =====================================================
Public Function GetImpostazione(ByVal chiave As String) As String
    Dim sh As Worksheet, f As Range
    On Error Resume Next
    Set sh = ThisWorkbook.Worksheets(NOME_IMP)
    On Error GoTo 0
    If sh Is Nothing Then Exit Function

    Set f = sh.Columns(1).Find(What:=chiave, LookIn:=xlValues, LookAt:=xlWhole)
    If Not f Is Nothing Then
        GetImpostazione = CStr(f.Offset(0, 1).value)
    Else
        GetImpostazione = CStr(f.Offset(0, 2).value)    'try find value on second column (alternate value)
    End If
End Function

' =====================================================
'   PATH MANAGEMENT
' =====================================================
Private Function NormalizePath(ByVal p As String) As String
    Dim s As String
    s = Replace(Trim$(p), "/", "\")
    Do While InStr(s, "\\") > 0 And left$(s, 2) <> "\\"
        s = Replace(s, "\\", "\")
    Loop
    If Len(s) > 0 And Right$(s, 1) <> "\" Then s = s & "\"
    NormalizePath = s
End Function

Private Sub EnsureFolder(ByVal p As String)
    Dim base As String, parts() As String, cur As String, i As Long
    base = NormalizePath(p)
    If Len(base) = 0 Then Exit Sub

    On Error Resume Next
    If left$(base, 2) = "\\" Then
        parts = Split(Mid$(base, 3), "\")
        If UBound(parts) >= 1 Then
            cur = "\\" & parts(0) & "\" & parts(1)
            For i = 2 To UBound(parts)
                If Len(parts(i)) > 0 Then
                    cur = cur & "\" & parts(i)
                    If Dir(cur, vbDirectory) = "" Then MkDir cur
                End If
            Next i
        End If
    Else
        cur = left$(base, 3)
        parts = Split(Mid$(base, 4), "\")
        For i = 0 To UBound(parts)
            If Len(parts(i)) > 0 Then
                If Right$(cur, 1) <> "\" Then cur = cur & "\"
                cur = cur & parts(i)
                If Dir(cur, vbDirectory) = "" Then MkDir cur
            End If
        Next i
    End If
    On Error GoTo 0
End Sub

' =====================================================
'   FILE NAME
' =====================================================
Private Function Sanitize(ByVal s As String) As String
    Dim bad As Variant: bad = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    Dim i As Long
    For i = LBound(bad) To UBound(bad)
        s = Replace(s, bad(i), "_")
    Next i
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    Sanitize = Trim$(s)
End Function

Private Function TruncateFileName(ByVal s As String, Optional ByVal maxLen As Long = 180) As String
    If Len(s) <= maxLen Then
        TruncateFileName = s
    Else
        TruncateFileName = left$(s, maxLen)
    End If
End Function

Private Function GetPdfName(ByVal dict As Object, ByVal pattern As String) As String
    Dim T As String: T = pattern
    Dim p1 As Long, p2 As Long, key As String, val As String

    If pattern = "{{Filename}}" And dict.exists("Filename") Then
        T = CStr(dict("Filename"))
        If LCase$(Right$(T, 4)) = ".pdf" Then T = left$(T, Len(T) - 4)
        GetPdfName = Sanitize(T)
        Exit Function
    End If

    Do
        p1 = InStr(T, "{{")
        If p1 = 0 Then Exit Do
        p2 = InStr(p1 + 2, T, "}}")
        If p2 = 0 Then Exit Do
        key = Mid$(T, p1 + 2, p2 - (p1 + 2))
        If dict.exists(key) Then val = CStr(dict(key)) Else val = ""
        T = left$(T, p1 - 1) & val & Mid$(T, p2 + 2)
    Loop

    GetPdfName = Sanitize(T)
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

' =====================================================
'   DIZIONARIO RIGA
' =====================================================
Private Function DizionarioDaRiga(ByVal lo As ListObject, ByVal r As ListRow) As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim i As Long, campo As String, v As Variant
    For i = 1 To lo.ListColumns.count
        campo = lo.ListColumns(i).Name
        v = r.Range.Cells(1, i).value
        If IsNull(v) Or IsError(v) Then v = ""
        d(campo) = v
    Next i
    Set DizionarioDaRiga = d
End Function

' =====================================================
'   SOSTITUZIONE SEGNAPOSTO
' =====================================================
Private Sub SostituisciSegnaposto(ByVal sh As Worksheet, ByVal dict As Object)
    Dim k As Variant, shp As Shape
    For Each k In dict.Keys
        sh.UsedRange.Replace What:="{{" & k & "}}", Replacement:=dict(k), LookAt:=xlPart, _
                             SearchOrder:=xlByRows, MatchCase:=False

        For Each shp In sh.Shapes
            If shp.Type = msoTextBox Or shp.Type = msoShapeRectangle Or shp.Type = msoPlaceholder Then
                If InStr(1, shp.TextFrame2.TextRange.text, "{{" & k & "}}", vbTextCompare) > 0 Then
                    shp.TextFrame2.TextRange.text = Replace(shp.TextFrame2.TextRange.text, "{{" & k & "}}", CStr(dict(k)))
                End If
            End If
        Next shp
    Next k
End Sub

' =====================================================
'   COLORA ESITO
' =====================================================
Private Sub ColoraEsito(ByVal sh As Worksheet, ByVal dict As Object)
    On Error Resume Next
    Dim res As String, es As Shape, idRow As String
    If dict.exists("Result") Then res = LCase$(Trim$(CStr(dict("Result"))))
    Set es = sh.Shapes("EsitoBar")
    If Not es Is Nothing Then
        Select Case res
            Case "ok", "si", "buono"
                es.Fill.ForeColor.RGB = RGB(90, 250, 90)
                es.line.ForeColor.RGB = RGB(60, 250, 60)
                es.TextFrame2.TextRange.text = "IL PRODOTTO PUÒ CONTINUARE AD ESSERE USATO"
            Case "no", "ko", "no ok", "no buono"
                es.Fill.ForeColor.RGB = RGB(250, 90, 90)
                es.line.ForeColor.RGB = RGB(250, 60, 60)
                es.TextFrame2.TextRange.text = "IL PRODOTTO DEVE ESSERE MESSO FUORI SERVIZIO"
            Case Else
                es.Fill.ForeColor.RGB = RGB(90, 90, 250)
                es.line.ForeColor.RGB = RGB(60, 60, 250)
                es.TextFrame2.TextRange.text = "ISPEZIONE NON VALIDA"
                If dict.exists("Number") Then idRow = CStr(dict("Number")) Else idRow = ""
                MsgBox "La scheda [" & idRow & "] ha un campo 'Result' non valido.", vbExclamation
        End Select
    End If
    On Error GoTo 0
End Sub

' =====================================================
'   FIRMA IMMAGINE (PNG/JPG o override SignatureFile)
' =====================================================
Private Function ResolveSignaturePath() As String
    Dim base As String: base = NormalizePath(ThisWorkbook.path)
    Dim p As String, override As String

    override = Trim$(GetImpostazione("SignatureFile"))
    If Len(override) > 0 Then
        If InStr(override, "\") Or InStr(override, "/") Then p = override Else p = base & override
        If Dir(p) <> "" Then ResolveSignaturePath = p: Exit Function
    End If

    If Dir(base & FIRMA_PNG) <> "" Then ResolveSignaturePath = base & FIRMA_PNG: Exit Function
    If Dir(base & FIRMA_JPG) <> "" Then ResolveSignaturePath = base & FIRMA_JPG: Exit Function
End Function

Private Sub InserisciFirmaImmagine(ByVal sh As Worksheet)
    Dim target As Shape, pic As Shape, f As String
    Dim sx As Double, sy As Double, sc As Double
    Dim s As Shape

    On Error Resume Next
    Set target = sh.Shapes("FirmaSegnaposto")
    On Error GoTo 0
    If target Is Nothing Then Exit Sub

    f = ResolveSignaturePath()
    If Len(f) = 0 Then Exit Sub

    ' Rimuovi eventuale immagine precedente
    For Each s In sh.Shapes
        If s.Name = "FirmaImage" Then s.Delete
    Next s

    Set pic = sh.Shapes.AddPicture(fileName:=f, LinkToFile:=msoFalse, SaveWithDocument:=msoTrue, _
                                   left:=target.left, top:=target.top, Width:=-1, Height:=-1)
    If pic Is Nothing Then Exit Sub

    pic.LockAspectRatio = msoTrue
    sx = (target.Width - 5) / pic.Width     ' width and height are reduced to take into account box border
    sy = (target.Height - 5) / pic.Height
    sc = IIf(sx < sy, sx, sy)
    If sc < 1 Then pic.Width = pic.Width * sc

    pic.left = target.left + (target.Width - pic.Width) / 2
    pic.top = target.top + (target.Height - pic.Height) / 2
    pic.Name = "FirmaImage"
    pic.line.Visible = msoFalse
End Sub

' =====================================================
'   NORMALIZZA CODICE (per ID azioni: preserva zeri / dedup)
' =====================================================
Private Function NormalizeCode(ByVal code As String) As String
    Dim s As String, i As Long, ch As String, onlyDigits As Boolean
    s = UCase$(Trim$(code))
    s = Replace(s, " ", "")
    If Len(s) = 0 Then NormalizeCode = "": Exit Function

    onlyDigits = True
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch < "0" Or ch > "9" Then
            onlyDigits = False
            Exit For
        End If
    Next i

    If onlyDigits Then
        If Len(s) < 3 Then
            NormalizeCode = Right$("000" & s, 3)
        Else
            NormalizeCode = s
        End If
    Else
        NormalizeCode = s
    End If
End Function

' =====================================================
'   AZIONI ISPETTIVE (mappatura + build lista)
' =====================================================
Private Sub EnsureAzioniMap()
    If Not gAzioni Is Nothing Then Exit Sub
    Set gAzioni = CreateObject("Scripting.Dictionary")
    gAzioni.CompareMode = 1

    On Error Resume Next
    Dim sh As Worksheet: Set sh = ThisWorkbook.Worksheets(NOME_AZIONI)
    On Error GoTo 0
    If sh Is Nothing Then Exit Sub

    Dim hdrRow As Long: hdrRow = 1
    Dim lastRow As Long: lastRow = sh.Cells(sh.Rows.count, 1).End(xlUp).row
    If lastRow < hdrRow + 1 Then Exit Sub

    Dim colID As Long, colFase As Long, colAtt As Long
    colID = FindHeaderCol(sh, hdrRow, "ID")
    colFase = FindHeaderCol(sh, hdrRow, "Fase ispezione")
    colAtt = FindHeaderCol(sh, hdrRow, "AttivitÃ ")
    If colID = 0 Then colID = 1
    If colFase = 0 Then colFase = 2
    If colAtt = 0 Then colAtt = 3

    Dim r As Long, kText As String, kNorm As String, kNum As String
    Dim fase As String, att As String, desc As String

    For r = hdrRow + 1 To lastRow
        kText = Trim$(CStr(sh.Cells(r, colID).text)) ' preserva zeri iniziali
        If Len(kText) > 0 Then
            fase = Trim$(CStr(sh.Cells(r, colFase).value))
            att = Trim$(CStr(sh.Cells(r, colAtt).value))
            If Len(fase) > 0 Then
                desc = UCase$(kText) & " " & ChrW(8212) & " " & fase & ":" & vbCrLf & att
            Else
                desc = UCase$(kText) & " " & ChrW(8212) & " " & att
            End If
            kNorm = NormalizeCode(kText)
            If IsNumeric(kText) Then kNum = CStr(CLng(kText)) Else kNum = ""

            gAzioni(UCase$(kText)) = desc
            gAzioni(kNorm) = desc
            If Len(kNum) > 0 Then gAzioni(kNum) = desc
        End If
    Next r
End Sub

Private Function FindHeaderCol(ByVal sh As Worksheet, ByVal hdrRow As Long, ByVal headerText As String) As Long
    Dim c As Range, lastCol As Long
    lastCol = sh.Cells(hdrRow, sh.Columns.count).End(xlToLeft).Column
    For Each c In sh.Range(sh.Cells(hdrRow, 1), sh.Cells(hdrRow, lastCol))
        If StrComp(Trim$(CStr(c.value)), headerText, vbTextCompare) = 0 Then
            FindHeaderCol = c.Column
            Exit Function
        End If
    Next c
End Function

Private Function BuildAzioniList(ByVal codesRaw As String) As String
    EnsureAzioniMap

    Dim cleaned As String
    cleaned = codesRaw
    cleaned = Replace(cleaned, vbTab, " ")
    cleaned = Replace(cleaned, vbCr, " ")
    cleaned = Replace(cleaned, vbLf, " ")
    cleaned = Replace(cleaned, ",", " ")
    cleaned = Replace(cleaned, ";", " ")
    cleaned = Replace(cleaned, "/", " ")
    cleaned = Replace(cleaned, "-", " ")

    Dim tokens() As String: tokens = Split(cleaned, " ")

    Dim seen As Object: Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = 1

    Dim i As Long, T As String
    Dim k1 As String, k2 As String, k3 As String
    Dim acc As String: acc = ""
    Dim count As Long: count = 0
    Dim desc As String

    For i = LBound(tokens) To UBound(tokens)
        T = Trim$(tokens(i))
        If Len(T) > 0 Then
            k1 = UCase$(T)
            k2 = NormalizeCode(k1)
            If IsNumeric(k1) Then k3 = CStr(CLng(k1)) Else k3 = ""

            Dim key As String: key = IIf(Len(k2) > 0, k2, k1)
            If Not seen.exists(key) Then
                seen(key) = True

                desc = ""
                If gAzioni.exists(k1) Then desc = gAzioni(k1)
                If Len(desc) = 0 And gAzioni.exists(k2) Then desc = gAzioni(k2)
                If Len(desc) = 0 And Len(k3) > 0 And gAzioni.exists(k3) Then desc = gAzioni(k3)
                If Len(desc) = 0 Then desc = UCase$(T) & " " & ChrW(8212) & " [Azione non trovata]"

                If Len(acc) = 0 Then acc = ChrW(8226) & " " & desc _
                Else acc = acc & vbCrLf & ChrW(8226) & " " & desc

                count = count + 1
                If count >= 15 Then Exit For
            End If
        End If
    Next i

    BuildAzioniList = acc
End Function

Private Sub ApplyAzioniIspettive(ByVal sh As Worksheet, ByVal dict As Object)
    On Error Resume Next
    If Not dict.exists("Required inspection activities") Then Exit Sub
    Dim raw As String: raw = CStr(dict("Required inspection activities"))
    Dim elenco As String: elenco = BuildAzioniList(raw)

    Dim target As Shape, shp As Shape
    Set target = Nothing
    Set target = sh.Shapes("AttivitaRichieste")

    If target Is Nothing Then
        For Each shp In sh.Shapes
            If shp.Type = msoTextBox Or shp.Type = msoShapeRectangle Then
                If InStr(1, shp.TextFrame2.TextRange.text, "Attivit" & ChrW(224) & " di ispezione richieste", vbTextCompare) > 0 _
                   Or InStr(1, shp.TextFrame2.TextRange.text, "Attivit" & ChrW(224) & "  di ispezione richieste", vbTextCompare) > 0 Then
                    Set target = shp
                    Exit For
                End If
            End If
        Next shp
    End If

    If target Is Nothing Then Exit Sub

    Dim header As String: header = "Attivit" & ChrW(224) & " di ispezione richieste:"
    If Len(elenco) > 0 Then
        target.TextFrame2.TextRange.text = header & vbCrLf & vbCrLf & elenco
    Else
        target.TextFrame2.TextRange.text = header & vbCrLf & vbCrLf & ChrW(8226) & " [Nessuna azione indicata]"
    End If
    With target.TextFrame2
        .MarginLeft = 8: .MarginRight = 8: .MarginTop = 6: .MarginBottom = 6
        .TextRange.Font.Size = 9
    End With
End Sub

' =====================================================
'   LOG (BUFFER CIRCOLARE: max 1000 righe totali)
' =====================================================
Private Sub LogEsito(ByVal id As String, ByVal esito As String, ByVal pathPdf As String)
    On Error Resume Next
    Dim sh As Worksheet
    Dim lastRow As Long, del As Long

    Set sh = ThisWorkbook.Worksheets(NOME_LOG)
    lastRow = sh.Cells(sh.Rows.count, 1).End(xlUp).row
    UnlockSheet sh
    
    If lastRow >= 1001 Then
        del = lastRow - 999   ' mantieni header + 999 righe
        sh.Rows("2:" & (1 + del)).Delete
        lastRow = sh.Cells(sh.Rows.count, 1).End(xlUp).row
    End If

    sh.Cells(lastRow + 1, 1).value = Now
    sh.Cells(lastRow + 1, 2).value = id
    sh.Cells(lastRow + 1, 3).value = esito
    sh.Cells(lastRow + 1, 4).value = pathPdf
    
    LockSheet sh
End Sub

' =====================================================
'   SAFE DELETE ESTESO rimuove fogli temporanei/zombie
' =====================================================
Private Sub SafeDeleteSheet(ByVal tmpName As String)
    Dim sh As Worksheet
    Application.DisplayAlerts = False

    On Error Resume Next
    ' 1) Elimina esattamente il foglio temporaneo passato
    Set sh = Nothing
    Set sh = SheetByName(tmpName)
    If Not sh Is Nothing Then sh.Delete

    ' 2) Elimina eventuali fogli "zombie" con nome simile (es. "Layout (2)")
    For Each sh In ThisWorkbook.Worksheets
        If LCase(Replace(sh.Name, " ", "")) = LCase(Replace(tmpName, " ", "")) Then sh.Delete
        If InStr(1, LCase(sh.Name), LCase(Replace(tmpName, " ", ""))) > 0 Then sh.Delete
    Next sh

    On Error GoTo 0
    Application.DisplayAlerts = True
End Sub

Private Function SheetByName(ByVal nm As String) As Worksheet
    Dim sh As Worksheet
    For Each sh In ThisWorkbook.Worksheets
        If StrComp(sh.Name, nm, vbTextCompare) = 0 Then
            Set SheetByName = sh
            Exit Function
        End If
    Next sh
End Function

' =====================================================
'   PROPRIETA' DOCUMENTO
' =====================================================
Private Sub ImpostaProprietaDocumento(ByVal wb As Workbook)
    On Error Resume Next  ' alcune proprietà potrebbero non essere impostabili in ogni contesto

    With wb.BuiltinDocumentProperties
        .Item("Author").value = "Domenico Longo - 2026 - CNSAS-SASS"                        ' Autore
        .Item("Title").value = "Verbale Ispezione DPI"                                      ' Titolo
        .Item("Subject").value = "Generated by " & GetPdfExportEngineVersion()              ' Oggetto
        .Item("Revision Number").value = "v0.1.0"                                           ' Versione
    End With

    On Error GoTo 0
End Sub

Private Sub RimuoviProprietaDocumento(ByVal wb As Workbook)
    On Error Resume Next  ' alcune proprietà potrebbero non essere impostabili in ogni contesto

    With wb.BuiltinDocumentProperties
        .Item("Author").value = "Domenico Longo - 2026 - CNSAS-SASS"                        ' Autore
        .Item("Title").value = ""                                                           ' Titolo
        .Item("Subject").value = ""                                                         ' Oggetto
        .Item("Revision Number").value = ""                                           ' Versione
    End With

    On Error GoTo 0
End Sub

' =====================================================
'   EXPORT PDF con safe-delete in tutti i casi
' =====================================================
Public Sub EsportaPDF_perRiga()
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim shLayout As Worksheet, lo As ListObject
    Dim wsDati As Worksheet
    
    ' preliminary checks; these two contain excellock/unlock and must execute before lock/unlock section in pdf export
    If Not ControllaCampiObbligatori Then Exit Sub
    If Not VerificaDateFormali Then Exit Sub
    
    SostituisciCodiciUbicazione
    
    On Error Resume Next
    Set shLayout = wb.Worksheets(NOME_LAYOUT)
    Set lo = wb.Worksheets("Dati").ListObjects(NOME_TAB)
    On Error GoTo 0

    If shLayout Is Nothing Or lo Is Nothing Then
        MsgBox "Errore: foglio 'Layout' o tabella 'tblDati' mancanti.", vbCritical
        Exit Sub
    End If

    Set wsDati = wb.Worksheets("Dati")
    'protect 'Dati' sheet anyway
    LockSheet wsDati

    UnlockSheet shLayout

    Dim outF As String: outF = GetImpostazione("OutputFolder")
    'If Len(outF) = 0 Then outF = ThisWorkbook.Path
    outF = ResolveExportPath(outF)
    outF = NormalizePath(outF)
    EnsureFolder outF

    Dim pattern As String: pattern = GetImpostazione("FileNamePattern")
    If Len(pattern) = 0 Then pattern = "{{Filename}}"

    Dim r As ListRow, tmp As Worksheet
    Dim dict As Object
    Dim fName As String, fullpath As String
    Dim idLog As String
    Dim hadErr As Boolean
    Dim exportedOK As Long

    ExcelLock
    
    ImpostaProprietaDocumento wb

    On Error GoTo FatalErr

    exportedOK = 0
    
    For Each r In lo.ListRows
        If r.Range.EntireRow.Hidden Then GoTo NextR
        If Application.WorksheetFunction.CountIf(r.Range, "<>") = 0 Then GoTo NextR     ' non esportare righe vuote

        Set dict = DizionarioDaRiga(lo, r)

        ' Valori globali da Impostazioni:
        dict("Ispettore") = GetImpostazione("Ispettore")
        dict("Matricola_Ispettore") = GetImpostazione("Matricola Ispettore")

        ' Crea foglio temporaneo
        shLayout.Copy After:=shLayout
        Set tmp = wb.Worksheets(shLayout.Index + 1)

        SostituisciSegnaposto tmp, dict
        ColoraEsito tmp, dict
        ApplyAzioniIspettive tmp, dict
        InserisciFirmaImmagine tmp

        fName = TruncateFileName(GetPdfName(dict, pattern))
        fullpath = outF & fName & ".pdf"

        On Error GoTo ExportErr
        tmp.ExportAsFixedFormat Type:=xlTypePDF, fileName:=fullpath, _
                                IncludeDocProperties:=True, _
                                Quality:=xlQualityStandard, _
                                IgnorePrintAreas:=False, OpenAfterPublish:=False
        On Error GoTo FatalErr

        If dict.exists("Number") Then idLog = CStr(dict("Number")) Else idLog = ""
        LogEsito idLog, "OK", fullpath
        exportedOK = exportedOK + 1

CleanTmp:
        If Not tmp Is Nothing Then SafeDeleteSheet tmp.Name

NextR:
    Next r

CleanExit:
    ExcelUnlock
    RimuoviProprietaDocumento wb

    If hadErr Then
        ThisWorkbook.Worksheets("Log").Activate
        MsgBox "Esportazione completata con errori. Controlla il 'Log'.", vbExclamation
    Else
        ThisWorkbook.Worksheets("Pannello").Activate
        MsgBox "Esportazione di [" & exportedOK & "] " & IIf(Abs(exportedOK) = 1, "scheda", "schede") & " PDF completata.", vbInformation
    End If
    
    LockSheet shLayout

    Exit Sub

' === ERRORI ==============================================================
ExportErr:
    hadErr = True
    LogEsito idLog, "ERRORE EXPORT [" & Err.Number & "]: " & Err.Description, fullpath
    Resume CleanTmp

FatalErr:
    hadErr = True
    LogEsito idLog, "ERRORE GENERALE [" & Err.Number & "]: " & Err.Description, fullpath
    Resume CleanTmp

End Sub

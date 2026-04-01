Attribute VB_Name = "modUpdateDPIAction"
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

Public Const MOD_DPI_ACTION_BUILDER_VERSION As String = "v1.1.0"

' =====================================================
'   VERSIONING
' =====================================================
Public Function GetDPIActionBuilderVersion() As String
    GetDPIActionBuilderVersion = "DPI action builder " & MOD_DPI_ACTION_BUILDER_VERSION & ";"
End Function


Public Sub AggiornaDatiDaAzioniDPI()
    Dim wb As Workbook
    Dim wsDati As Worksheet, wsAzioni As Worksheet, wsLog As Worksheet
    Dim lo As ListObject
    Dim colScheda As ListColumn, colReq As ListColumn
    Dim idCol As Long, tipoCol As Long, azioniCol As Long
    Dim lastRowAz As Long
    Dim r As Long, nRows As Long
    Dim schedaCell As Range, reqCell As Range
    Dim valScheda As String, codice As String
    Dim arr As Variant
    Dim dictText As Object, dictNum As Object ' Scripting.Dictionary (late binding)
    Dim n As Long
    Dim hadLog As Boolean
    Dim duplicates As Long

    On Error GoTo CleanFail

    Set wb = ThisWorkbook
    Set wsDati = wb.Worksheets("Dati")
    Set wsAzioni = wb.Worksheets("Azioni_DPI")
    
    UnlockSheet wsDati

    ' Tuning prestazioni
    ExcelLock

    ' Recupero tabella e colonne
    Set lo = wsDati.ListObjects("tblDati")
    Set colScheda = lo.ListColumns("SCHEDA")
    Set colReq = lo.ListColumns("Required inspection activities")

    ' Trova le colonne per ID/Tipo DPI/Azioni Ispettive nel foglio Azioni_DPI
    idCol = FindHeaderColumn(wsAzioni, "ID")
    tipoCol = FindHeaderColumn(wsAzioni, "Tipo DPI")
    azioniCol = FindHeaderColumn(wsAzioni, "Azioni Ispettive")

    If idCol = 0 Or tipoCol = 0 Or azioniCol = 0 Then
        Err.Raise vbObjectError + 513, , _
            "Impossibile trovare una o più intestazioni nel foglio 'Azioni_DPI'." & vbCrLf & _
            "Richieste: 'ID', 'Tipo DPI', 'Azioni Ispettive' (riga 1)."
    End If

    lastRowAz = wsAzioni.Cells(wsAzioni.Rows.count, idCol).End(xlUp).row
    If lastRowAz < 2 Then
        Err.Raise vbObjectError + 514, , "Il foglio 'Azioni_DPI' non contiene righe dati."
    End If

    ' Crea/azzera foglio di log
    Set wsLog = PrepareLogSheet(wb, "Log_Update_AzioniDPI")

    ' Dizionari lookup: per testo (ID così com'è) e per numerico (equivalenza numerica fino a 3 cifre)
    Set dictText = CreateObject("Scripting.Dictionary")
    Set dictNum = CreateObject("Scripting.Dictionary")
    duplicates = 0

    Dim i As Long, idText As String, tipoText As String, azioniText As String
    Dim numKey As String, numVal As Long

    For i = 2 To lastRowAz
        idText = Trim$(CStr(wsAzioni.Cells(i, idCol).value))
        If Len(idText) > 0 Then
            tipoText = CStr(wsAzioni.Cells(i, tipoCol).value)
            azioniText = CStr(wsAzioni.Cells(i, azioniCol).value)

            ' Array dati: [0]=ID originale (testo, con eventuali zeri), [1]=Tipo DPI, [2]=Azioni Ispettive
            arr = Array(idText, tipoText, azioniText)

            ' Dizionario testuale (chiave: ID esatto come da cella)
            If dictText.exists(idText) Then
                duplicates = duplicates + 1
            Else
                dictText.Add idText, arr
            End If

            ' Dizionario numerico (chiave: valore numerico delle prime 3 cifre nell'ID)
            numKey = ExtractNumericCode(idText, 3)
            If Len(numKey) > 0 Then
                numVal = CLng(numKey) ' max 3 cifre ? safe in Long
                If Not dictNum.exists(numVal) Then
                    dictNum.Add numVal, arr
                End If
                ' Se ci sono duplicati numerici verrà usata la prima occorrenza
            End If
        End If
    Next i

    ' Itera righe della tabella tblDati
    nRows = 0
    If Not lo.DataBodyRange Is Nothing Then
        nRows = lo.DataBodyRange.Rows.count
    End If

    If nRows = 0 Then GoTo CleanExit ' Tabella vuota

    Dim processed As Long, notFound As Long, noCode As Long
    Dim found As Boolean
    Dim newSchedaText As String

    For r = 1 To nRows
        Set schedaCell = colScheda.DataBodyRange.Cells(r, 1)
        Set reqCell = colReq.DataBodyRange.Cells(r, 1)

        valScheda = Trim$(CStr(schedaCell.value))
        If lo.ListRows(r).Range.EntireRow.Hidden Then GoTo NextR
        If Len(valScheda) = 0 Then
            ' Riga considerata vuota per lo scopo: skip
            GoTo NextR
        End If

        codice = ExtractNumericCode(valScheda, 2) ' gli ID delle azioni ispettive hanno solo 2 cifre (1..99 max)
        If Len(codice) = 0 Then
            ' Nessun codice estraibile
            LogRow wsLog, Now, valScheda, "", "Nessun codice numerico", ""
            noCode = noCode + 1
            GoTo NextR
        End If

        found = False
        ' 1) Match testuale esatto su ID
        If dictText.exists(codice) Then
            arr = dictText(codice)
            found = True
        Else
            ' 2) Match per equivalenza numerica (ignora zeri iniziali etc.)
            n = CLng(codice)
            If dictNum.exists(n) Then
                arr = dictNum(n)
                found = True
            End If
        End If

        If found Then
            newSchedaText = CStr(arr(0)) & " - " & CStr(arr(1))   ' "ID - Tipo DPI"
            schedaCell.value = newSchedaText
            reqCell.value = CStr(arr(2))                          ' "Azioni Ispettive"
            processed = processed + 1
        Else
            LogRow wsLog, Now, valScheda, codice, "ID non trovato", ""
            notFound = notFound + 1
        End If

NextR:
    Next r

    ExcelUnlock
    
    ' Sommario nel log
    hadLog = (noCode > 0 Or notFound > 0 Or duplicates > 0)
    If hadLog Then
        LogRow wsLog, Now, "[Sommario]", "", "Elaborate", "OK: " & processed & " | Nessun codice: " & noCode & " | ID non trovati: " & notFound & " | ID duplicati (intestazioni ID ripetute in Azioni_DPI): " & duplicates
        wsLog.Columns.AutoFit
        ThisWorkbook.Worksheets("Log_Update_AzioniDPI").Activate
        MsgBox "Aggiornamento completato." & vbCrLf & _
               "OK: " & processed & vbCrLf & _
               "Nessun codice: " & noCode & vbCrLf & _
               "ID non trovati: " & notFound & vbCrLf & _
               "Duplicati ID (in 'Azioni_DPI'): " & duplicates & vbCrLf & _
               "Vedi foglio 'Log_AzioniDPI' per i dettagli.", vbInformation
    Else
        On Error Resume Next
        Application.DisplayAlerts = False
        wb.Worksheets("Log_Update_AzioniDPI").Delete
        ThisWorkbook.Worksheets("Pannello").Activate
        Application.DisplayAlerts = True
        On Error GoTo 0
        MsgBox "Aggiornamento completato." & vbCrLf & "OK: " & processed & " | Nessun codice: " & noCode & " | ID non trovati: " & notFound, vbInformation
    End If

CleanExit:
    ExcelUnlock
    
    LockSheet wsDati
    
    Exit Sub

CleanFail:
    ExcelUnlock
    
    LockSheet wsDati
    
    MsgBox "Errore: " & Err.Description, vbCritical
End Sub

' === Helpers ===

Private Function ExtractNumericCode(ByVal text As String, Optional ByVal maxDigits As Long = 3) As String
    ' Estrae fino a maxDigits cifre dal testo (ignorando il resto), procedendo da sinistra verso destra.
    Dim i As Long
    Dim ch As String
    Dim out As String

    text = CStr(text)
    For i = 1 To Len(text)
        ch = Mid$(text, i, 1)
        If ch >= "0" And ch <= "9" Then
            out = out & ch
            If Len(out) >= maxDigits Then Exit For
        End If
    Next i

    ExtractNumericCode = out
End Function

Private Function FindHeaderColumn(ByVal ws As Worksheet, ByVal headerName As String) As Long
    ' Cerca l'intestazione in riga 1 (case-insensitive, match esatto della cella).
    Dim f As Range
    Set f = ws.Rows(1).Find(What:=headerName, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False, SearchOrder:=xlByColumns)
    If Not f Is Nothing Then
        FindHeaderColumn = f.Column
    Else
        FindHeaderColumn = 0
    End If
End Function

Private Function PrepareLogSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count))
        ws.Name = sheetName
    Else
        ws.Cells.Clear
    End If

    ws.Range("A1:E1").value = Array("Timestamp", "SCHEDA (originale)", "Codice estratto", "Esito", "Nota")
    ws.Rows(1).Font.Bold = True
    Set PrepareLogSheet = ws
End Function

Private Sub LogRow(ByVal ws As Worksheet, ByVal ts As Date, ByVal schedaOriginale As String, ByVal codice As String, ByVal esito As String, ByVal nota As String)
    Dim nextRow As Long
    nextRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row + 1
    ws.Cells(nextRow, 1).value = ts
    ws.Cells(nextRow, 2).value = schedaOriginale
    ws.Cells(nextRow, 3).value = codice
    ws.Cells(nextRow, 4).value = esito
    ws.Cells(nextRow, 5).value = nota
End Sub


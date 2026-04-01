Attribute VB_Name = "modImportData"
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

Public Const MOD_IMPORT_DATA_VERSION As String = "v1.1.0"

' =====================================================
'   VERSIONING
' =====================================================
Public Function GetDataImportEngineVersion() As String
    GetDataImportEngineVersion = "Data import engine " & MOD_IMPORT_DATA_VERSION & ";"
End Function

Public Sub Importa_tblDati_da_XLSX()
    Dim pathXlsx As Variant
    Dim wbSrc As Workbook, wsSrc As Worksheet, loSrc As ListObject
    Dim wbDst As Workbook, wsDst As Worksheet, loDst As ListObject
    Dim arrSrc As Variant, arrOut As Variant
    Dim srcHeaders As Variant, dstHeaders As Variant
    Dim srcMap As Object, dstMap As Object
    Dim nSrcRows As Long, nSrcCols As Long, nDstCols As Long
    Dim i As Long, j As Long, srcCol As Long
    Dim scelta As VbMsgBoxResult
    Dim oldScr As Boolean, oldCalc As XlCalculation, oldEvt As Boolean
    Dim missing As String
    Dim basePath As String
    Dim textCols As Variant
    Dim textColsMap As Object
    Dim jText As Long
    
    
    On Error GoTo GestErr
    
    
    ' Elenco colonne da importare come TESTO (personalizza qui)
    textCols = Array("Serial Number", "Number")   ' aggiungi eventuali altri header
    Set textColsMap = CreateObject("Scripting.Dictionary")
    textColsMap.CompareMode = vbTextCompare
    For jText = LBound(textCols) To UBound(textCols)
        textColsMap(LCase$(Trim$(CStr(textCols(jText))))) = True
    Next jText

    
    '------------------ Selettore file (robusto)
    basePath = ThisWorkbook.path
    If Len(basePath) = 0 Then basePath = CreateObject("WScript.Shell").SpecialFolders("Desktop")
    ChDrive basePath
    ChDir basePath

    pathXlsx = Application.GetOpenFilename( _
        FileFilter:="Cartella di lavoro di Excel (*.xlsx), *.xlsx", _
        FilterIndex:=1, _
        Title:="Seleziona il file .xlsx da importare")
    If VarType(pathXlsx) = vbBoolean And pathXlsx = False Then Exit Sub  ' Annullato
    
    '------------------ Ottimizzazioni
    ExcelLock
    
    '------------------ Sorgente (.xlsx)
    Set wbSrc = Application.Workbooks.Open(fileName:=CStr(pathXlsx), ReadOnly:=True)
    On Error Resume Next
    Set wsSrc = wbSrc.Worksheets("Dati")
    If wsSrc Is Nothing Then
        MsgBox "Nel file sorgente non trovo il foglio 'Dati'.", vbCritical, "Foglio Dati mancante"
        GoTo GestErr
    End If
    On Error GoTo 0
    'On Error GoTo GestErr
    'If wsSrc Is Nothing Then Err.Raise vbObjectError + 100, , "Nel file sorgente non trovo il foglio 'Dati'."
    
    On Error Resume Next
    Set loSrc = wsSrc.ListObjects("tblDati")
    If loSrc Is Nothing Then
        MsgBox "Nel foglio 'Dati' del sorgente non trovo la tabella 'tblDati'.", vbCritical, "Tabella tblDati mancante"
        GoTo GestErr
    End If
    On Error GoTo 0
    'On Error GoTo GestErr
    'If loSrc Is Nothing Then Err.Raise vbObjectError + 101, , "Nel foglio 'Dati' del sorgente non trovo la tabella 'tblDati'."
    
    ' Headers sorgente
    srcHeaders = loSrc.HeaderRowRange.value   ' 1 x n
    nSrcCols = loSrc.Range.Columns.count
    
    ' Righe sorgente (tutte, indipendenti dai filtri)
    If loSrc.DataBodyRange Is Nothing Then
        nSrcRows = 0
    Else
        nSrcRows = loSrc.DataBodyRange.Rows.count
        arrSrc = loSrc.DataBodyRange.value
    End If
    
    '------------------ Destinazione (.xlsm corrente)
    Set wbDst = ThisWorkbook
    Set wsDst = wbDst.Worksheets("Dati")
    Set loDst = wsDst.ListObjects("tblDati")
    
    UnlockSheet wsDst
    
    ' rimuove eventuali filtri presenti
    On Error Resume Next
        If Not loDst.AutoFilter Is Nothing Then
            If loDst.AutoFilter.FilterMode Then loDst.AutoFilter.ShowAllData
        End If
        If wsDst.FilterMode Then wsDst.ShowAllData
    On Error GoTo 0
    
    ' Headers destinazione (ordine effettivo in tabella)
    dstHeaders = loDst.HeaderRowRange.value   ' 1 x n
    nDstCols = loDst.Range.Columns.count
    
    ' Applica NumberFormat "@" alle colonne di destinazione corrispondenti (anche se la tabella è vuota)
    For j = 1 To nDstCols
        Dim hdr As String
        hdr = Normalize(CStr(dstHeaders(1, j)))
        If textColsMap.exists(hdr) Then
            ' Formatto l'intera colonna della tabella (header + corpo) come testo
            loDst.ListColumns(j).Range.NumberFormat = "@"
        End If
    Next j

    '------------------ Mappe header (case-insensitive, trim)
    Set srcMap = CreateObject("Scripting.Dictionary")
    Set dstMap = CreateObject("Scripting.Dictionary")
    srcMap.CompareMode = vbTextCompare
    dstMap.CompareMode = vbTextCompare
    
    For j = 1 To nSrcCols
        srcMap(Normalize(CStr(srcHeaders(1, j)))) = j
    Next j
    For j = 1 To nDstCols
        dstMap(Normalize(CStr(dstHeaders(1, j)))) = j
    Next j
    
    ' Avviso colonne mancanti nel sorgente (continua riempiendo vuoto)
    missing = ""
    Dim key As Variant
    For Each key In dstMap.Keys
        If Not srcMap.exists(key) Then
            If Len(missing) > 0 Then missing = missing & ", "
            missing = missing & key
        End If
    Next key
    If Len(missing) > 0 Then
        MsgBox "Attenzione: nel file .xlsx mancano le seguenti colonne richieste in destinazione:" & vbCrLf & _
               missing & vbCrLf & vbCrLf & _
               "Proseguirò comunque; le colonne mancanti verranno lasciate vuote.", _
               vbExclamation, "Colonne mancanti"
    End If
    
    '------------------ Scelta utente: Sostituisci o Accoda
    If nSrcRows = 0 Then
        scelta = MsgBox("Il file sorgente non contiene righe nella tabella 'tblDati'." & vbCrLf & _
                        "Vuoi cancellare le righe esistenti in destinazione?", _
                        vbQuestion + vbYesNoCancel + vbDefaultButton2, "File sorgente vuoto")
        If scelta = vbCancel Then GoTo Uscita
        ' anche se non ci sono righe da importare, onora la scelta
    Else
        scelta = MsgBox("Vuoi CANCELLARE i dati esistenti e sostituirli con quelli importati?" & vbCrLf & _
                        vbCrLf & _
                        "Sì = Sostituisci" & vbCrLf & _
                        "No = Accoda ai dati presenti" & vbCrLf & _
                        "Annulla = Interrompi", _
                        vbQuestion + vbYesNoCancel + vbDefaultButton2, "Sostituisci/Accoda tblDati")
        If scelta = vbCancel Then GoTo Uscita
    End If
    
    '------------------ Componi l'array di output (righe x colonne destinazione)
    If nSrcRows > 0 Then
        ReDim arrOut(1 To nSrcRows, 1 To nDstCols)
        For i = 1 To nSrcRows
            For j = 1 To nDstCols
                Dim dstHdr As String
                dstHdr = Normalize(CStr(dstHeaders(1, j)))
    
                srcCol = -1
                If srcMap.exists(dstHdr) Then srcCol = srcMap(dstHdr)
    
                If srcCol > 0 Then
                    If textColsMap.exists(dstHdr) Then
                        ' Trattare come TESTO
                        If IsError(arrSrc(i, srcCol)) Or IsEmpty(arrSrc(i, srcCol)) Then
                            arrOut(i, j) = vbNullString
                        ElseIf IsNumeric(arrSrc(i, srcCol)) Then
                            ' Evita notazione E+ in caso di numeri grandi
                            arrOut(i, j) = Format$(arrSrc(i, srcCol), "0")
                        Else
                            ' Già stringa: normalizza come testo
                            arrOut(i, j) = CStr(arrSrc(i, srcCol))
                        End If
                    Else
                        ' Trattare normalmente (numero, data, testo…)
                        arrOut(i, j) = arrSrc(i, srcCol)
                    End If
                Else
                    arrOut(i, j) = Empty
                End If
            Next j
        Next i
    End If
    
    '------------------ Importazione (mantieni stile e filtri)
    ' Nota: non tocchiamo TableStyle né i criteri di filtro.
    ' Per "Sostituisci": eliminiamo tutte le righe della tabella (restano header, stile e filtro).
    
    Dim oldRows As Long, newRows As Long
    oldRows = 0
    If Not loDst.DataBodyRange Is Nothing Then oldRows = loDst.DataBodyRange.Rows.count
    
    If scelta = vbYes Then
        ' Sostituisci: elimina tutte le righe esistenti della tabella
        If oldRows > 0 Then
            loDst.DataBodyRange.Delete xlShiftUp
        End If
        oldRows = 0
    End If
    
    If nSrcRows > 0 Then
        ' Ridimensiona tabella a header + (oldRows + nSrcRows)
        newRows = oldRows + nSrcRows
        loDst.Resize loDst.Range.Resize(RowSize:=1 + newRows, ColumnSize:=nDstCols)
        
        ' Scrivi blocco in coda
        Dim rngWrite As Range
        Set rngWrite = loDst.DataBodyRange.Rows(oldRows + 1).Resize(nSrcRows, nDstCols)
        rngWrite.value = arrOut
    Else
        ' Nessuna riga da importare: se l'utente ha scelto "Sostituisci" rimani con sola riga header
        If scelta = vbYes Then
            loDst.Resize loDst.Range.Resize(RowSize:=1, ColumnSize:=nDstCols)
        End If
    End If
    
    ' Piccola cosmetica (facoltativa): niente autofit per non alterare layout utente
    ' wsDst.Columns.AutoFit
    
    MsgBox "Importazione di [" & nSrcRows & "] " & IIf(Abs(nSrcRows) = 1, "riga", "righe") & " completata.", vbInformation, "Importazione completata"
    GoTo Uscita

GestErr:
    'MsgBox "Errore durante l'importazione: " & Err.Description, vbCritical, "Importa tblDati"

Uscita:
    On Error Resume Next
    If Not wbSrc Is Nothing Then wbSrc.Close SaveChanges:=False
    ThisWorkbook.Worksheets("Dati").Activate
    wsDst.Range("A1").Select
    
    LockSheet wsDst
    
    ExcelUnlock
End Sub


'------------------ Helpers ------------------

Private Function Normalize(ByVal s As String) As String
    Normalize = LCase$(Trim$(s))
End Function


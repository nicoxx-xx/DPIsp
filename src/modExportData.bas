Attribute VB_Name = "modExportData"
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

Public Const MOD_EXPORT_DATA_VERSION As String = "v1.1.0"
Private Const msoFileDialogSaveAs As Long = 2


' =====================================================
'   VERSIONING
' =====================================================
Public Function GetDataExportEngineVersion() As String
    GetDataExportEngineVersion = "Data export engine " & MOD_EXPORT_DATA_VERSION & ";"
End Function

Public Sub Esporta_tblDati_in_XLSX()
    Dim wbSrc As Workbook
    Dim wsSrc As Worksheet
    Dim loSrc As ListObject
    Dim wbOut As Workbook
    Dim wsOut As Worksheet
    Dim loOut As ListObject
    Dim arr As Variant
    Dim nRows As Long, nCols As Long
    Dim savePath As String
    Dim okPath As Boolean
    
    On Error GoTo GestErr
      
    '--- ottimizzazioni
    ExcelLock
    
    Set wbSrc = ThisWorkbook
    On Error Resume Next
    Set wsSrc = wbSrc.Worksheets("Dati")
    On Error GoTo 0
    If wsSrc Is Nothing Then
        Err.Raise vbObjectError + 100, , "Non trovo il foglio 'Dati'."
    End If
    
    On Error Resume Next
    Set loSrc = wsSrc.ListObjects("tblDati")
    On Error GoTo 0
    If loSrc Is Nothing Then
        Err.Raise vbObjectError + 101, , "Non trovo la tabella 'tblDati' nel foglio 'Dati'."
    End If
    
    '--- calcolo dimensioni
    nCols = loSrc.Range.Columns.count
    If loSrc.DataBodyRange Is Nothing Then
        nRows = 0
    Else
        nRows = loSrc.DataBodyRange.Rows.count
        arr = loSrc.DataBodyRange.value  ' <- legge TUTTE le righe, anche se filtrate
    End If
    
    '--- crea nuovo file con un foglio
    Set wbOut = Application.Workbooks.Add(xlWBATWorksheet)
    Set wsOut = wbOut.Worksheets(1)
    wsOut.Name = "Dati"
    
    '--- scrive intestazioni
    wsOut.Range("A1").Resize(1, nCols).value = loSrc.HeaderRowRange.value
    
    '--- scrive corpo (se presente)
    If nRows > 0 Then
        wsOut.Range("A2").Resize(nRows, nCols).value = arr
    End If
    
    '--- trasforma in tabella, replica (se possibile) nome e stile
    Set loOut = wsOut.ListObjects.Add(xlSrcRange, wsOut.Range("A1").CurrentRegion, , xlYes)
    
    'Prova a chiamarla come l'originale, altrimenti mette _Export
    On Error Resume Next
    loOut.Name = "tblDati"
    If Err.Number <> 0 Then
        Err.Clear
        loOut.Name = "tblDati_Export"
    End If
    
    'Prova a replicare lo stile della tabella originale (se definito)
    If Len(loSrc.TableStyle) > 0 Then
        loOut.TableStyle = loSrc.TableStyle
    End If
    On Error GoTo GestErr
    
    '--- autofit colonne
    wsOut.Columns.AutoFit
    
    '--- finestra Salva con nome (stile Windows) + gestione esistenza file
    okPath = PickSavePathXlsx( _
                defaultName:="tblDati_export_" & Format(Now, "yyyymmdd_hhnnss") & ".xlsx", _
                outPath:=savePath)
    If Not okPath Then
        'utente ha annullato: chiude il file creato senza salvare
        Application.DisplayAlerts = False
        wbOut.Close SaveChanges:=False
        Application.DisplayAlerts = True
        GoTo Pulisci
    End If
    
    '--- salva in .xlsx (formato macro-free)
    Application.DisplayAlerts = False
    wbOut.SaveAs fileName:=savePath, FileFormat:=xlOpenXMLWorkbook  ' 51 = .xlsx
    wbOut.Close SaveChanges:=False
    Application.DisplayAlerts = True
    
    MsgBox "Esportazione di [" & nRows & "] " & IIf(Abs(nRows) = 1, "riga", "righe") & " completata!" & vbCrLf & savePath, vbInformation, "Export tblDati"
    GoTo Pulisci

GestErr:
    If Not wbOut Is Nothing Then
        Application.DisplayAlerts = False
        On Error Resume Next
        wbOut.Close SaveChanges:=False
        Application.DisplayAlerts = True
    End If
    MsgBox "Errore: " & Err.Description, vbCritical, "Export tblDati"

Pulisci:
    ExcelUnlock
End Sub


'===========================================================
' Finestra Salva con nome (Windows 11 style) che:
' - imposta filtro .xlsx
' - forza estensione .xlsx
' - se il file esiste chiede: Sovrascrivi / Cambia nome / Annulla
' Ritorna True se percorso scelto, False se annullato.
'===========================================================
Private Function PickSavePathXlsx(ByVal defaultName As String, ByRef outPath As String) As Boolean
    Dim fd As Object 'FileDialog
    Dim p As String
    Dim ans As VbMsgBoxResult
    Dim basePath As String
        
    On Error GoTo Fine
    
    basePath = ThisWorkbook.path
    If Len(basePath) = 0 Then basePath = CreateObject("WScript.Shell").SpecialFolders("Desktop")
    
    Do
        Set fd = Application.FileDialog(msoFileDialogSaveAs)
        With fd
            .Title = "Esporta tabella 'tblDati' in .xlsx"
            .InitialFileName = basePath & Application.PathSeparator & defaultName
            
            On Error Resume Next
            .Filters.Clear
            .Filters.Add "Cartella di lavoro di Excel (*.xlsx)", "*.xlsx"
            .FilterIndex = 1
            If Err.Number <> 0 Then
                ' Se i filtri non sono disponibili/gestiti, ignoriamo e continuiamo
                Err.Clear
                ' MsgBox "errore sui filtri", vbExclamation
            End If
            On Error GoTo 0

            If .Show <> -1 Then
                PickSavePathXlsx = False   'utente ha annullato
                Exit Function
            End If
            p = .SelectedItems(1)
        End With
        
        ' Forza estensione .xlsx
        If LCase$(Right$(p, 5)) <> ".xlsx" Then
            p = p & ".xlsx"
        End If
        
        ' Se esiste, chiedi cosa fare
        If Len(Dir$(p, vbNormal)) > 0 Then
            ans = MsgBox( _
                    "Il file esiste già:" & vbCrLf & p & vbCrLf & vbCrLf & _
                    "Vuoi sovrascriverlo?", _
                    vbExclamation + vbYesNoCancel + vbDefaultButton2, _
                    "Conferma sovrascrittura")
            Select Case ans
                Case vbYes
                    outPath = p
                    PickSavePathXlsx = True
                    Exit Function
                Case vbNo
                    ' ripresenta la finestra per cambiare nome
                    ' continua il Do
                Case vbCancel
                    PickSavePathXlsx = False
                    Exit Function
            End Select
        Else
            outPath = p
            PickSavePathXlsx = True
            Exit Function
        End If
    Loop
    
Fine:
End Function


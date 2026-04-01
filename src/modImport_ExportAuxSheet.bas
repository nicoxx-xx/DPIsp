Attribute VB_Name = "modImport_ExportAuxSheet"
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

Public Const MOD_IMPORTEXPORTAUXSHEET_VERSION As String = "v1.0.0"

' =====================================================
'   VERSIONING
' =====================================================
Public Function GetImpExpAuxSheetVersion() As String
    GetImpExpAuxSheetVersion = "Import/Export aux sheet engine " & MOD_IMPORTEXPORTAUXSHEET_VERSION & ";"
End Function


Sub EsportaFogliAuxInXLSX()

    Dim Percorso As Variant
    Dim wbTemp As Workbook
    Dim FogliDaEsportare As Variant
    Dim i As Long

    ' Fogli da esportare
    FogliDaEsportare = Array("Ubicazioni", "Produttori", "Modelli", _
                             "Azioni_Ispettive", "Azioni_DPI", _
                             "Impostazioni", "LICENZA")

    ' Finestra di dialogo per salvare il file .xlsx
    Percorso = Application.GetSaveAsFilename( _
                InitialFileName:="BackupFogliAusiliari.xlsx", _
                FileFilter:="Excel (*.xlsx), *.xlsx", _
                Title:="Seleziona la cartella in cui esportare")

    If Percorso = False Then Exit Sub   ' annullato dall’utente

    ExcelLock

    ' Crea nuovo file temporaneo
    Set wbTemp = Workbooks.Add(xlWBATWorksheet)


    ' Copia i fogli nel nuovo file
    For i = LBound(FogliDaEsportare) To UBound(FogliDaEsportare)
        ThisWorkbook.Worksheets(FogliDaEsportare(i)).Copy After:=wbTemp.Sheets(wbTemp.Sheets.count)
    Next i

    ' Rimuove il primo foglio creato
    Application.DisplayAlerts = False
    wbTemp.Worksheets(1).Delete
    Application.DisplayAlerts = True
    
    ExcelUnlock
    
    ' Salva come .xlsx
    wbTemp.SaveAs fileName:=Percorso, FileFormat:=xlOpenXMLWorkbook

    wbTemp.Close SaveChanges:=False

    MsgBox "Esportazione fogli ausiliari completata!", vbInformation

End Sub

Sub ImportaFogliAuxDaXLSX()

    Dim Percorso As Variant
    Dim wbOrigine As Workbook
    Dim FogliDaImportare As Variant
    Dim i As Long
    Dim NomeFoglio As String

    ' Fogli da importare
    FogliDaImportare = Array("Ubicazioni", "Produttori", "Modelli", _
                             "Azioni_Ispettive", "Azioni_DPI", _
                             "Impostazioni", "LICENZA")

    ' Seleziona file .xlsx
    Percorso = Application.GetOpenFilename( _
                FileFilter:="Excel (*.xlsx), *.xlsx", _
                Title:="Seleziona il file da importare")

    If Percorso = False Then Exit Sub   ' annullato

    ExcelLock

    ' Apre il file origine
    Set wbOrigine = Workbooks.Open(Percorso)

    ' Ciclo sui fogli da importare
    For i = LBound(FogliDaImportare) To UBound(FogliDaImportare)

        NomeFoglio = FogliDaImportare(i)

        ' Se il foglio esiste nel file origine
        If FoglioEsiste(wbOrigine, NomeFoglio) Then

            ' Elimina foglio esistente nel file xlsm (se presente)
            If FoglioEsiste(ThisWorkbook, NomeFoglio) Then
                Application.DisplayAlerts = False
                ThisWorkbook.Worksheets(NomeFoglio).Delete
                Application.DisplayAlerts = True
            End If

            ' Copia il foglio
            wbOrigine.Worksheets(NomeFoglio).Copy After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count)

        End If

    Next i

    wbOrigine.Close SaveChanges:=False

    ExcelUnlock

    MsgBox "Importazione fogli ausiliari completata!", vbInformation

End Sub


Function FoglioEsiste(wb As Workbook, NomeFoglio As String) As Boolean
    On Error Resume Next
    FoglioEsiste = Not wb.Worksheets(NomeFoglio) Is Nothing
    On Error GoTo 0
End Function

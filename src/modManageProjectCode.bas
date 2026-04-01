Attribute VB_Name = "modManageProjectCode"
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

'== Costanti VBIDE (late binding) ==
Const vbext_ct_StdModule As Long = 1
Const vbext_ct_ClassModule As Long = 2
Const vbext_ct_MSForm As Long = 3
Const vbext_ct_Document As Long = 100
' (Solo informativa) Protezione progetto:
' vbext_pp_none = 0, vbext_pp_locked = 1

' ============================
' Esporta:
'  - Moduli standard (.bas)
'  - UserForm (.frm + .frx)
' Opzionale: classi e moduli documento
' ============================

Public Sub AAA_Esporta_Moduli_e_UserForm()
    Dim wb As Workbook
    Dim exportRoot As String, exportFolder As String, baseName As String
    Dim vbProj As Object           ' VBIDE.VBProject (late binding)
    Dim vbc As Object              ' VBIDE.VBComponent (late binding)
    Dim ct As Long, fName As String
    Dim cntExported As Long, cntSkipped As Long
    Dim sep As String: sep = Application.PathSeparator
    Dim t0 As Single: t0 = Timer
    
    '== Opzioni ==
    Const ASK_FOR_FOLDER As Boolean = True   ' True = chiedi cartella di esportazione; False = salva accanto al file
    Const INCLUDE_CLASSES As Boolean = True ' True = esporta anche .cls
    Const INCLUDE_DOCUMENTS As Boolean = False ' True = esporta anche moduli fogli/ThisWorkbook come .cls
    
    On Error GoTo Err_Handler
    
    Set wb = ThisWorkbook  ' Cambia in ActiveWorkbook se preferisci
    baseName = NameWithoutExtension(wb.Name)
    
    ' Cartella radice
    If ASK_FOR_FOLDER Then
        exportRoot = PickFolder("Scegli la cartella di esportazione dei moduli")
        If Len(exportRoot) = 0 Then
            MsgBox "Operazione annullata.", vbInformation
            Exit Sub
        End If
    Else
        exportRoot = IIf(Len(wb.path) > 0, wb.path, CreateOnDesktop())
    End If
    
    ' Sottocartella specifica per questo file
    exportFolder = exportRoot & sep & "VBA_Export_" & baseName
    EnsureFolderExists exportFolder
    
    Set vbProj = wb.VBProject
    
    ' Se il progetto è protetto da password, .VBComponents può generare errore/limitazioni.
    ' Questo semplice check evita di procedere se il progetto è bloccato.
    If IsVBProjectLocked(vbProj) Then
        MsgBox "Il progetto VBA è protetto da password. Sbloccalo e riprova.", vbExclamation
        Exit Sub
    End If
    
    ' Cicla i componenti
    For Each vbc In vbProj.VBComponents
        ct = vbc.Type
        
        Select Case ct
            Case vbext_ct_StdModule
                fName = exportFolder & sep & vbc.Name & ".bas"
                SafeDelete fName
                vbc.Export fName
                cntExported = cntExported + 1
            
            Case vbext_ct_MSForm
                ' Esporta .frm e (automaticamente) il relativo .frx
                fName = exportFolder & sep & vbc.Name & ".frm"
                SafeDelete fName
                SafeDelete exportFolder & sep & vbc.Name & ".frx" ' risorse
                vbc.Export fName
                cntExported = cntExported + 1
            
            Case vbext_ct_ClassModule
                If INCLUDE_CLASSES Then
                    fName = exportFolder & sep & vbc.Name & ".cls"
                    SafeDelete fName
                    vbc.Export fName
                    cntExported = cntExported + 1
                Else
                    cntSkipped = cntSkipped + 1
                End If
            
            Case vbext_ct_Document
                ' (ThisWorkbook, Foglio1, ecc.)
                If INCLUDE_DOCUMENTS Then
                    fName = exportFolder & sep & vbc.Name & ".cls"
                    SafeDelete fName
                    vbc.Export fName
                    cntExported = cntExported + 1
                Else
                    cntSkipped = cntSkipped + 1
                End If
            
            Case Else
                ' altri tipi non gestiti
                cntSkipped = cntSkipped + 1
        End Select
    Next vbc
    
    MsgBox "Esportazione completata in:" & vbCrLf & exportFolder & vbCrLf & _
           "Esportati: " & cntExported & IIf((INCLUDE_CLASSES Or INCLUDE_DOCUMENTS) = False, _
           vbCrLf & "(Classi/Documenti ignorati per impostazione)", vbNullString) & vbCrLf & _
           "Tempo: " & Format(Timer - t0, "0.00") & " s", vbInformation
    Exit Sub

Err_Handler:
    ' Errore classico se non è abilitato l’accesso al modello a oggetti di progetto VBA
    If Err.Number = 1004 Then
        MsgBox "Errore 1004: accesso al progetto VBA non consentito." & vbCrLf & vbCrLf & _
               "Vai su: File > Opzioni > Centro protezione > Impostazioni Centro protezione > " & _
               "Impostazioni macro > spunta 'Considera attendibile l'accesso al modello a oggetti del progetto VBA'." & vbCrLf & _
               "Poi riprova.", vbCritical
    Else
        MsgBox "Errore " & Err.Number & ": " & Err.Description, vbCritical
    End If
End Sub

' ============================
' Rimuove:
'  - Moduli standard (.bas)
'  - UserForm (.frm + .frx)
'  - classi (.cls)
' ============================
Public Sub AAA_Rimuovi_Moduli_e_UserForm()
    Dim vbComp  As Object 'As VBIDE.VBComponent
    Dim nome As String

    On Error Resume Next

    For Each vbComp In ThisWorkbook.VBProject.VBComponents
        Select Case vbComp.Type

            'Moduli standard (.bas)
            Case vbext_ct_StdModule
                nome = vbComp.Name
                ThisWorkbook.VBProject.VBComponents.Remove vbComp

            'UserForm (.frm + .frx)
            Case vbext_ct_MSForm
                nome = vbComp.Name
                ThisWorkbook.VBProject.VBComponents.Remove vbComp

            'Classi (.cls)
            Case vbext_ct_ClassModule
                nome = vbComp.Name
                ThisWorkbook.VBProject.VBComponents.Remove vbComp

        End Select
    Next vbComp

    On Error GoTo 0

    MsgBox "Tutti i moduli, form e classi utente sono stati rimossi.", vbInformation
End Sub


' ============================
' Importa:
'  - Moduli standard (.bas)
'  - UserForm (.frm + .frx)
'  - classi (.cls)
' ============================
Public Sub AAA_Importa_Moduli_e_UserForm_DaCartella()

    Dim folderPath As String
    Dim fso As Object
    Dim folder As Object
    Dim file As Object
    Dim vbProj  As Object ' As VBIDE.VBProject
    Dim vbComp  As Object ' As VBIDE.VBComponent
    Dim nomeModulo As String
    Dim ext As String
    Dim cntImported As Long

    'Chiede la cartella da cui importare
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Seleziona la cartella contenente i moduli da importare"
        If .Show <> -1 Then Exit Sub
        folderPath = .SelectedItems(1)
    End With

    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"

    Set fso = CreateObject("Scripting.FileSystemObject")
    Set folder = fso.GetFolder(folderPath)
    Set vbProj = ThisWorkbook.VBProject

    cntImported = 0
    
    'Scansione dei file nella cartella
    For Each file In folder.Files

        ext = LCase(fso.GetExtensionName(file.Name))
        
        'Gestione solo dei file compatibili
        If ext = "bas" Or ext = "cls" Or ext = "frm" Then
            
            nomeModulo = fso.GetBaseName(file.Name)
            
            'Se un modulo con lo stesso nome esiste, non fare nulla, altrimenti importa modulo
            On Error Resume Next
            Set vbComp = vbProj.VBComponents(nomeModulo)
            
            If vbComp Is Nothing Then
                'vbProj.VBComponents.Remove vbComp
                'Importazione del modulo
                vbProj.VBComponents.Import file.path
            End If
            Set vbComp = Nothing
            On Error GoTo 0
            
            cntImported = cntImported + 1
            
        End If
    Next file

    MsgBox "Importazione completata di [" & cntImported & "] moduli.", vbInformation

End Sub

' ========== Helpers ==========

Private Function PickFolder(ByVal titleText As String) As String
    On Error GoTo fallback
    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = titleText
        .AllowMultiSelect = False
        If .Show = -1 Then
            PickFolder = .SelectedItems(1)
        Else
            PickFolder = vbNullString
        End If
    End With
    Exit Function
fallback:
    ' In caso di Office senza FileDialog (raro), ripiega sul Desktop
    PickFolder = CreateOnDesktop()
End Function

Private Sub EnsureFolderExists(ByVal path As String)
    If Len(Dir(path, vbDirectory)) = 0 Then
        MkDir path
    End If
End Sub

Private Sub SafeDelete(ByVal filePath As String)
    On Error Resume Next
    If Len(Dir(filePath, vbNormal)) > 0 Then Kill filePath
    On Error GoTo 0
End Sub

Private Function NameWithoutExtension(ByVal fileName As String) As String
    Dim p As Long
    p = InStrRev(fileName, ".")
    If p > 0 Then
        NameWithoutExtension = left$(fileName, p - 1)
    Else
        NameWithoutExtension = fileName
    End If
End Function

Private Function CreateOnDesktop() As String
    Dim path As String
    path = Environ$("USERPROFILE") & Application.PathSeparator & "Desktop"
    If Len(Dir(path, vbDirectory)) = 0 Then
        path = CurDir$ ' fallback
    End If
    CreateOnDesktop = path
End Function

Private Function IsVBProjectLocked(ByVal vbProj As Object) As Boolean
    On Error Resume Next
    ' Accesso a .VBComponents genera errore se bloccato in certi casi; qui proviamo una proprietà sicura.
    Dim dummy As Long
    dummy = vbProj.Protection ' 0 = none, 1 = locked (valori comuni)
    IsVBProjectLocked = (dummy <> 0)
    On Error GoTo 0
End Function


VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDPI 
   Caption         =   "Gestione dati di riga DPI"
   ClientHeight    =   7830
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14685
   OleObjectBlob   =   "frmDPI.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmDPI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False









Option Explicit

' ===========================
' frmDPI - Gestione tblDati
' ===========================

' Stato interno
Private mLo As ListObject                ' tabella destinazione: Dati!tblDati
Private mHdrMap As Object                ' Dictionary: header normalizzato -> indice colonna (1-based)
Private mEditIndex As Long               ' 0 = nuovo; >0 = indice riga (1-based) in DataBodyRange da modificare
Dim ws As Worksheet

' Intestazioni dei campi (come in tabella)
Private Const HDR_UBIC As String = "Ubicazione"
Private Const HDR_NUMBER As String = "Number"
Private Const HDR_DATE As String = "Date"
Private Const HDR_SCHEDA As String = "SCHEDA"
Private Const HDR_COMMENTS As String = "Comments"
Private Const HDR_SERIAL As String = "Serial Number"
Private Const HDR_MANUF As String = "Manufacturer"
Private Const HDR_MODEL As String = "Model"
Private Const HDR_DOM As String = "Date of Manufacture"
Private Const HDR_DOP As String = "Date of Purchase"
Private Const HDR_DOFU As String = "Date of First Use"
Private Const HDR_NEXTINSP As String = "Next Ispection Date"  ' ortografia come in tabella
Private Const HDR_RETIRE As String = "Date for retirement"
Private Const HDR_ANN As String = "Annotazioni"
Private Const HDR_CUST As String = "Customer"
Private Const HDR_RESULT As String = "Result"

' === Stato del DatePicker in-pagina ===
Private mDP_TargetTB As MSForms.TextBox   ' il TextBox che riceverà la data
Private mDP_Year As Long
Private mDP_Month As Long
Private mDP_Day As Long


' ============ Inizializzazione ============

Private Sub UserForm_Initialize()
    On Error GoTo ErrInit
    
    ' 1) Aggancia la tabella di destinazione
    'Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets("Dati")
    Set mLo = ws.ListObjects("tblDati")
       
    ' 2) Mappa intestazioni per nome (case-insensitive)
    Set mHdrMap = CreateObject("Scripting.Dictionary")
    mHdrMap.CompareMode = vbTextCompare
    Dim j As Long
    For j = 1 To mLo.ListColumns.count
        mHdrMap(Normalize(mLo.ListColumns(j).Name)) = j
    Next
    
    ' 3) Popola combobox statici
    With cboResult
        .Clear
        .AddItem "ok"
        .AddItem "ko"
        '.MatchRequired = True
    End With
    
    ' 4) Popola combobox da altri fogli
    LoadSchedeFromAzioniDPI cboScheda                 ' "ID - Tipo DPI" da Azioni_DPI
    LoadDistinctFromSheetColumn cboManufacturer, "Produttori", "Produttore"
    LoadDistinctFromSheetColumn cboModel, "Modelli", "Modello"
    LoadDistinctFromSheetColumn cboUbicazione, "Ubicazioni", "UbicazioneEstesa"
    'cboScheda.MatchRequired = True
    'cboManufacturer.MatchRequired = True
    'cboModel.MatchRequired = True
    
    ' 5) Formatta colonne "Number" e "Serial Number" come TESTO sul foglio
    EnsureTextFormatForColumns Array(HDR_NUMBER, HDR_SERIAL)
    
    ' 6) Stato: nuovo record
    mEditIndex = 0
    ClearForm
    
    ' 7 Credits:
    With txtCredits
        .text = GetCredits() & vbCrLf & GetInputFormPanelVersion() & vbCrLf & _
        GetMouseScrollEngineVersion()
        .Enabled = False
    End With
    
    ' 8 install mouse wheel hook
    EnableMouseScroll Me
    'Application.EnableCancelKey = xlDisabled

    Exit Sub

ErrInit:
    MsgBox "Errore inizializzazione form: " & Err.Description, vbCritical, "frmDPI"
End Sub

Private Sub UserForm_Terminate()
    ' uninstall mouse wheel hook
    DisableMouseScroll Me
End Sub

' ============ eventi utente ============
Private Sub UserForm_Click()
    CloseInlineDatePicker
End Sub

Private Sub cboDP_Month_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Button = 1 Then  ' solo click sinistro
        Me.cboDP_Month.DropDown
    End If
End Sub

Private Sub cboDP_Year_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Button = 1 Then  ' solo click sinistro
        Me.cboDP_Year.DropDown
    End If
End Sub

Private Sub cboManufacturer_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Button = 1 Then  ' solo click sinistro
        Me.cboManufacturer.DropDown
    End If
End Sub

Private Sub cboModel_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Button = 1 Then  ' solo click sinistro
        Me.cboModel.DropDown
    End If
End Sub

Private Sub cboResult_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Button = 1 Then  ' solo click sinistro
        Me.cboResult.DropDown
    End If
End Sub

Private Sub cboScheda_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Button = 1 Then  ' solo click sinistro
        Me.cboScheda.DropDown
    End If
End Sub

Private Sub cboUbicazione_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Button = 1 Then  ' solo click sinistro
        Me.cboUbicazione.DropDown
    End If
End Sub

Private Sub cmdNuovo_Click()
    mEditIndex = 0  ' 0 = nuovo; >0 = indice riga (1-based) in DataBodyRange da modificare
    ClearForm
    txtNumber.SetFocus
End Sub

Private Sub cmdModifica_Click()
    On Error GoTo ErrMod
    
    Dim s As String
    's = InputBox("Inserisci il valore ID Scheda da modificare", "Modifica DPI")
    'If Len(Trim$(s)) = 0 Then Exit Sub
    
    ' vincolo: solo cifre
    'If Not IsAllDigits(s) Then
       ' MsgBox "Inserisci solo cifre per ID Scheda.", vbExclamation, "Modifica DPI"
        'Exit Sub
    'End If
    
    Dim rng As Range, c As Range
    Set rng = GetColumnRange(HDR_NUMBER)
    If rng Is Nothing Then
        MsgBox "Colonna 'Number' non trovata nella tabella.", vbCritical
        Exit Sub
    End If
    
    ' Input numerico: Application.InputBox con Type:=1 accetta solo numeri
    s = Application.InputBox( _
            Prompt:="Inserisci il valore ID Scheda da modificare:", _
            Title:="Modifica riga per ID Scheda", _
            Default:=0, _
            Type:=1)
    ' Se annulla, s = False
    If (s = False) Then Exit Sub
    If (Len(Trim$(s)) = 0) Then
        MsgBox "Numero ID Scheda non valido.", vbInformation, "Modifica DPI"
        Exit Sub
    End If
    
        
    ' Cerca prima come stringa (poiché la colonna è formattata testo),
    ' poi in fallback come numero intero, nel caso di dati storici
    Set c = rng.Find(What:=s, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    If c Is Nothing Then
        Set c = rng.Find(What:=CLng(s), LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    End If
    
    If c Is Nothing Then
        MsgBox "Nessuna riga trovata con ID Scheda = " & s & ".", vbInformation, "Modifica DPI"
        Exit Sub
    End If
    
    mEditIndex = c.row - rng.Rows(1).row + 1
    LoadFromIndex mEditIndex
    txtNumber.SetFocus
    LockSheet ws
    Exit Sub

ErrMod:
    LockSheet ws
    MsgBox "Errore in Modifica: " & Err.Description, vbCritical
End Sub

Private Sub cmdSalva_Click()
    On Error GoTo ErrSave
    
    ' 1) Validazione completa (tutti obbligatori)
    If Not ValidateForm Then Exit Sub
     
    UnlockSheet ws
    
    ' 3) Riga di destinazione
    Dim idx As Long
    If mEditIndex = 0 Then
        Dim lr As ListRow
        Set lr = mLo.ListRows.Add
        idx = lr.Index
    Else
        idx = mEditIndex
    End If
    
    ' 4) Scrittura valori (Note: Number e Serial Number come TESTO)
    SetFieldText HDR_UBIC, Trim$(cboUbicazione.value), idx
    SetFieldText HDR_NUMBER, txtNumber.text, idx                                ' testo
    SetFieldDate HDR_DATE, txtDate.text, idx
    SetFieldText HDR_SCHEDA, Trim$(cboScheda.value), idx
    SetFieldText HDR_COMMENTS, txtComments.text, idx
    SetFieldText HDR_SERIAL, Trim$(txtSerialNumber.text), idx            ' testo
    SetFieldText HDR_MANUF, Trim$(cboManufacturer.value), idx
    SetFieldText HDR_MODEL, Trim$(cboModel.value), idx
    SetFieldDate HDR_DOM, txtDoM.text, idx
    SetFieldDate HDR_DOP, txtDoP.text, idx
    SetFieldDate HDR_DOFU, txtDoFU.text, idx
    SetFieldDateOrNNN HDR_NEXTINSP, txtNextInsp.text, idx
    SetFieldDate HDR_RETIRE, txtRetirement.text, idx
    SetFieldText HDR_ANN, txtAnnotazioni.text, idx
    SetFieldText HDR_CUST, Trim$(txtCustomer.text), idx
    SetFieldText HDR_RESULT, LCase$(Trim$(cboResult.value)), idx         ' "ok" / "ko"
    
    ' 5) Fine
    MsgBox IIf(mEditIndex = 0, "Nuovo DPI salvato.", "Modifica salvata."), vbInformation, "tblDati"
    mEditIndex = 0
    ClearForm
    txtNumber.SetFocus
    LockSheet ws
    Exit Sub

ErrSave:
    LockSheet ws
    MsgBox "Errore nel salvataggio: " & Err.Description, vbCritical, "Salva DPI"
End Sub

Private Sub cmdElimina_Click()

    Dim ws As Worksheet
    Dim lo As ListObject
    Dim numberCol As Long
    Dim hdr As Range
    Dim v As Variant
    Dim targetValue As Double
    Dim rngData As Range
    Dim r As Long
    Dim matches As Collection
    Dim msg As String
    Dim i As Long
    Dim resp As VbMsgBoxResult
    
    On Error GoTo CleanFail
    
    Set ws = ThisWorkbook.Worksheets("Dati")
    Set lo = ws.ListObjects("tblDati")
    
    ' Verifica che la tabella abbia righe dati
    If lo.DataBodyRange Is Nothing Then
        MsgBox "La tabella 'tblDati' non contiene dati.", vbInformation
        Exit Sub
    End If
    
    ' Trova la colonna "Number" indipendentemente dall'ordine
    numberCol = 0
    For Each hdr In lo.HeaderRowRange.Cells
        If Trim$(LCase$(hdr.Value2)) = "number" Then
            numberCol = hdr.Column - lo.Range.Columns(1).Column + 1 ' indice relativo nella tabella
            Exit For
        End If
    Next hdr
    
    If numberCol = 0 Then
        MsgBox "Impossibile trovare la colonna 'Number' nella tabella 'tblDati'.", vbExclamation
        Exit Sub
    End If
    
    ' Input numerico: Application.InputBox con Type:=1 accetta solo numeri
    v = Application.InputBox( _
            Prompt:="Inserisci il valore ID Scheda da eliminare:", _
            Title:="Elimina riga per ID Scheda", _
            Default:=0, _
            Type:=1)
    ' Se annulla, s = False
    If (v = False) Then Exit Sub
    If (Len(Trim$(v)) = 0) Then
        MsgBox "Numero ID Scheda non valido.", vbInformation, "Modifica DPI"
        Exit Sub
    End If
    
    targetValue = CDbl(v)
    
    ExcelLock
    UnlockSheet ws
    
    ' Raccoglie le righe che corrispondono
    Set matches = New Collection
    Set rngData = lo.DataBodyRange
    
    ' Scorri tutte le righe della tabella
    For r = 1 To rngData.Rows.count
        Dim cellVal As Variant
        cellVal = rngData.Cells(r, numberCol).Value2
        
        ' Confronto numerico robusto: prova a convertire a numero
        If IsNumeric(cellVal) Then
            If CDbl(cellVal) = targetValue Then
                matches.Add r ' indice di riga relativo a DataBodyRange
            End If
        Else
            ' Se il campo fosse testo numerico con spazi, prova a normalizzare
            Dim s As String
            s = Trim$(CStr(cellVal))
            If Len(s) > 0 And IsNumeric(s) Then
                If CDbl(s) = targetValue Then
                    matches.Add r
                End If
            End If
        End If
    Next r
    
    If matches.count = 0 Then
        MsgBox "Nessuna riga trovata con ID Scheda = " & targetValue & ".", vbInformation, "Elimina DPI"
        GoTo CleanExit
    End If
    
    If matches.count = 1 Then
        ' Conferma singola eliminazione
        msg = "Trovata 1 riga con ID Scheda = " & targetValue & "." & vbCrLf & _
              "Vuoi eliminarla definitivamente?"
        resp = MsgBox(msg, vbQuestion + vbYesNo + vbDefaultButton2)
        If resp <> vbYes Then GoTo CleanExit
        
        ' Elimina la riga corrispondente (dal basso verso l'alto è sicuro; qui è una sola)
        lo.ListRows(matches(1)).Delete
        
        MsgBox "Riga eliminata.", vbInformation, "Riga eliminata con successo"
    Else
        ' Più corrispondenze: conferma eliminazione di tutte
        msg = "Trovate " & matches.count & " righe con ID Scheda = " & targetValue & "." & vbCrLf & _
              "Vuoi eliminarle tutte definitivamente?"
        resp = MsgBox(msg, vbExclamation + vbYesNo + vbDefaultButton2)
        If resp <> vbYes Then GoTo CleanExit
        
        ' Elimina tutte le righe corrispondenti partendo dal basso
        For i = matches.count To 1 Step -1
            lo.ListRows(matches(i)).Delete
        Next i
        
        MsgBox matches.count & " righe eliminate.", vbInformation, "Righe eliminate con successo"
    End If

CleanExit:
    ExcelUnlock
    LockSheet ws
    Exit Sub

CleanFail:
    ExcelUnlock
    LockSheet ws
    MsgBox "Si è verificato un errore: " & Err.Number & " - " & Err.Description, vbExclamation
End Sub

Private Sub cmdDuplicati_Click()

    Dim ws As Worksheet
    Dim lo As ListObject
    Dim numberCol As Long
    Dim hdr As Range
    Dim dataRng As Range
    Dim r As Long
    Dim v As Variant
    Dim key As String
    Dim keyo As Variant
    
    Dim dictCounts As Object ' Scripting.Dictionary: key=normalizzazione numerica, item=conteggio
    Dim dictValues As Object ' Scripting.Dictionary: key=normalizzazione, item=Collection di rappresentazioni originali
    
    Dim dupKeys As Collection         ' normalizzazioni duplicate (conteggio >= 2)
    Dim criteriaList As Collection    ' valori originali da usare come filtro
    Dim critArr As Variant
    Dim i As Long, j As Long
    Dim dupRowsCount As Long
    Dim msg As String, preview As String
    
    On Error GoTo CleanFail
    
    Set ws = ThisWorkbook.Worksheets("Dati")
    Set lo = ws.ListObjects("tblDati")
    
    If lo.DataBodyRange Is Nothing Then
        ' Pulisce filtri se ci fossero, poi avvisa
        On Error Resume Next
        If Not lo.AutoFilter Is Nothing Then
            If lo.AutoFilter.FilterMode Then lo.AutoFilter.ShowAllData
        End If
        If ws.FilterMode Then ws.ShowAllData
        On Error GoTo 0
        
        MsgBox "La tabella 'tblDati' non contiene dati.", vbInformation
        Exit Sub
    End If
    
    ' Trova la colonna "Number" a prescindere dall'ordine
    numberCol = 0
    For Each hdr In lo.HeaderRowRange.Cells
        If Trim$(LCase$(hdr.Value2)) = "number" Then
            numberCol = hdr.Column - lo.Range.Columns(1).Column + 1 ' indice relativo alla tabella
            Exit For
        End If
    Next hdr
    
    If numberCol = 0 Then
        MsgBox "Impossibile trovare la colonna 'Number' nella tabella 'tblDati'.", vbExclamation
        Exit Sub
    End If
    
    Set dataRng = lo.DataBodyRange
    
    ' Dizionari (late binding, nessun riferimento da aggiungere)
    Set dictCounts = CreateObject("Scripting.Dictionary")
    Set dictValues = CreateObject("Scripting.Dictionary")
    
    ' Conta le occorrenze e memorizza tutte le rappresentazioni originali
    For r = 1 To dataRng.Rows.count
        v = dataRng.Cells(r, numberCol).Value2
        
        ' Ignora i blank
        If Len(Trim$(CStr(v))) > 0 Then
            If IsNumeric(v) Then
                key = CStr(CDbl(v)) ' normalizzazione coerente (gestisce 1, 1.0, "01", ecc.)
            Else
                Dim s As String
                s = Trim$(CStr(v))
                If IsNumeric(s) Then
                    key = CStr(CDbl(s))
                Else
                    ' Non numerico: lo consideriamo una chiave testuale (raro ma gestito)
                    key = s
                End If
            End If
            
            ' Conteggio
            If Not dictCounts.exists(key) Then
                dictCounts.Add key, 1
            Else
                dictCounts(key) = dictCounts(key) + 1
            End If
            
            ' Rappresentazioni originali (per includerle tutte nel filtro)
            If Not dictValues.exists(key) Then
                Dim col As Collection
                Set col = New Collection
                col.Add v
                dictValues.Add key, col
            Else
                Dim exists As Boolean: exists = False
                Dim col2 As Collection
                Set col2 = dictValues(key)
                For i = 1 To col2.count
                    If col2(i) = v Then
                        exists = True
                        Exit For
                    End If
                Next i
                If Not exists Then col2.Add v
            End If
        End If
    Next r
    
    ' Individua le chiavi duplicate (conteggio >= 2)
    Set dupKeys = New Collection
    For Each keyo In dictCounts.Keys
        If dictCounts(keyo) >= 2 Then dupKeys.Add keyo
    Next keyo
    
    ' PULIZIA: rimuove qualsiasi filtro già presente sulla tabella
    On Error Resume Next
    If Not lo.AutoFilter Is Nothing Then
        If lo.AutoFilter.FilterMode Then lo.AutoFilter.ShowAllData
    End If
    If ws.FilterMode Then ws.ShowAllData
    On Error GoTo 0
    
    If dupKeys.count = 0 Then
        ' Nessun duplicato: lascia la tabella senza filtri e avvisa
        MsgBox "Nessun DPI duplicato trovato.", vbInformation
        Exit Sub
    End If
    
    ' Costruisci i criteri per l'AutoFilter e il riepilogo
    Set criteriaList = New Collection
    dupRowsCount = 0
    preview = ""
    For i = 1 To dupKeys.count
        Dim k As String
        k = dupKeys(i)
        
        ' Riepilogo (Number = valore normalizzato, Occorrenze = N)
        preview = preview & "• Number = " & k & "  (occorrenze: " & dictCounts(k) & ")" & vbCrLf
        
        ' Aggiungi tutte le rappresentazioni originali di quel valore
        Dim reps As Collection
        Set reps = dictValues(k)
        For j = 1 To reps.count
            criteriaList.Add reps(j)
        Next j
        
        dupRowsCount = dupRowsCount + dictCounts(k)
    Next i
    
    ' Converte la collection in array per AutoFilter (0-based)
    ReDim critArr(0 To criteriaList.count - 1)
    For i = 1 To criteriaList.count
        critArr(i - 1) = criteriaList(i)
    Next i
    
    ExcelLock
    
    ' Applica il filtro su Number per mostrare solo i duplicati
    lo.Range.AutoFilter Field:=numberCol, Criteria1:=critArr, Operator:=xlFilterValues
    
    ExcelUnlock
    
    ' MsgBox di riepilogo (se molto lungo, viene troncato visivamente da MsgBox)
    msg = "Individuati " & dupKeys.count & " DPI duplicati'." & vbCrLf & _
          "Righe complessive corrispondenti mostrate tramite filtro: " & dupRowsCount & "." & vbCrLf & vbCrLf & _
          "Dettaglio:" & vbCrLf & preview
    MsgBox msg, vbInformation

    Exit Sub

CleanFail:
    ExcelUnlock
    MsgBox "Si è verificato un errore: " & Err.Number & " - " & Err.Description, vbExclamation
End Sub

Private Sub cmdValida_Click()

    If ValidateForm() Then
        MsgBox "Il form non contiene errori", vbInformation
    End If
End Sub

Private Sub cmdChiudi_Click()
    Unload Me
End Sub

Private Sub txtDate_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Button = 1 Then OpenInlineDatePicker Me.txtDate, "dd/mm/yyyy"
End Sub

Private Sub txtDoM_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Button = 1 Then OpenInlineDatePicker Me.txtDoM, "dd/mm/yyyy"
End Sub

Private Sub txtDoP_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Button = 1 Then OpenInlineDatePicker Me.txtDoP, "dd/mm/yyyy"
End Sub

Private Sub txtDoFU_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Button = 1 Then OpenInlineDatePicker Me.txtDoFU, "dd/mm/yyyy"
End Sub

Private Sub txtNextInsp_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Button = 1 Then OpenInlineDatePicker Me.txtNextInsp, "dd/mm/yyyy"
End Sub

Private Sub txtRetirement_MouseUp(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal y As Single)
    If Button = 1 Then OpenInlineDatePicker Me.txtRetirement, "dd/mm/yyyy"
End Sub


' ============ Validazioni ============

Private Function ValidateForm() As Boolean
    ValidateForm = False
    
    ' Tutti i campi obbligatori
    If Not HasValue(txtNumber.text) Then MsgReq "ID Scheda", txtNumber: Exit Function
    If Not IsAllDigits(txtNumber.text) Then
        MsgBox "'ID Scheda' deve contenere solo cifre.", vbExclamation: txtNumber.SetFocus: Exit Function
    End If
    
    Dim numText As String
    numText = Trim$(txtNumber.text)
    If Not IsUniqueNumberText(numText, mEditIndex) Then
        MsgBox "Il valore ID Scheda è già presente in tabella.", vbExclamation
        txtNumber.SetFocus
        Exit Function
    End If
    
    If (Not HasValue(cboUbicazione.value)) Or (Not ComboHasMatchedValue(cboUbicazione)) Then MsgReq "Ubicazione", cboUbicazione: Exit Function
    If Not HasValidDate(txtDate.text) Then MsgReqDate "Data odierna", txtDate: Exit Function
    If Not HasValue(txtCustomer.text) Then MsgReq "Cliente", txtCustomer: Exit Function
    If (Not HasValue(cboScheda.value)) Or (Not ComboHasMatchedValue(cboScheda)) Then MsgReq "Tipologia DPI", cboScheda: Exit Function
    If (Not HasValue(cboManufacturer.value)) Or (Not ComboHasMatchedValue(cboManufacturer)) Then MsgReq "Produttore", cboManufacturer: Exit Function
    If (Not HasValue(cboModel.value)) Or (Not ComboHasMatchedValue(cboModel)) Then MsgReq "Modello", cboModel: Exit Function
    If Not HasValue(txtSerialNumber.text) Then MsgReq "Serial Number", txtSerialNumber: Exit Function
    If Not HasValidDate(txtDoM.text) Then MsgReqDate "Data di produzione", txtDoM: Exit Function
    If Not HasValidDate(txtDoP.text) Then MsgReqDate "Data di acquisto", txtDoP: Exit Function
    If Not HasValidDate(txtDoFU.text) Then MsgReqDate "Data di primo utilizzo", txtDoFU: Exit Function
    If Not HasValidDateOrNNN(txtNextInsp.text) Then MsgBox "Il campo 'Data di prossima isp.' deve contenere una data valida (gg/mm/aaaa) oppure 'nnn'.", vbExclamation: Exit Function
    If Not HasValidDate(txtRetirement.text) Then MsgReqDate "Data di dismissione", txtRetirement: Exit Function
    If Not HasValue(txtComments.text) Then MsgReq "Commenti", txtComments: Exit Function
    'If Not HasValue(txtAnnotazioni.text) Then MsgReq "Annotazioni", txtAnnotazioni: Exit Function
    If (Not HasValue(cboResult.value)) Or (Not ComboHasMatchedValue(cboResult)) Then MsgReq "Risultato di ispezione (ok/ko)", cboResult: Exit Function
    Dim r As String: r = LCase$(Trim$(cboResult.value))
    If r <> "ok" And r <> "ko" Then
        MsgBox "'Risultato di ispezione' deve essere 'ok' oppure 'ko'.", vbExclamation
        cboResult.SetFocus: Exit Function
    End If
    

    ValidateForm = True
End Function

' ============ Helpers di validazione/UI ============

Private Sub MsgReq(ByVal fieldName As String, ctrl As Object)
    MsgBox "Il campo '" & fieldName & "' non è presente o non è valido.", vbExclamation
    ctrl.SetFocus
End Sub

Private Sub MsgReqDate(ByVal fieldName As String, ctrl As Object)
    If Len(Trim$(ctrl.text)) = 0 Then
        MsgReq fieldName, ctrl
    Else
        MsgBox "Il campo '" & fieldName & "' non è una data valida (gg/mm/aaaa).", vbExclamation
        ctrl.SetFocus
    End If
End Sub

Public Function ComboHasMatchedValue(ByVal cbo As MSForms.ComboBox) As Boolean
    Dim userValue As String
    Dim i As Long
    
    ' Normalizza il valore dell'utente
    userValue = Trim$(CStr(cbo.value))
    If Len(userValue) = 0 Then
        ComboHasMatchedValue = False
        Exit Function
    End If
    
    ' Scorre gli item del ComboBox
    For i = 0 To cbo.ListCount - 1
        If StrComp(userValue, CStr(cbo.List(i)), vbTextCompare) = 0 Then
            ComboHasMatchedValue = True
            Exit Function
        End If
    Next i
    
    ComboHasMatchedValue = False
End Function

Private Function HasValue(ByVal v As Variant) As Boolean
    HasValue = (Len(Trim$(CStr(v))) > 0)
End Function

Private Function HasValidDate(ByVal s As String) As Boolean
    s = Trim$(s)
    HasValidDate = (Len(s) > 0 And IsDate(s))
End Function

Private Function IsAllDigits(ByVal s As String) As Boolean
    Dim i As Long, ch As Integer
    s = Trim$(s)
    If Len(s) = 0 Then IsAllDigits = False: Exit Function
    For i = 1 To Len(s)
        ch = Asc(Mid$(s, i, 1))
        If ch < 48 Or ch > 57 Then IsAllDigits = False: Exit Function
    Next
    IsAllDigits = True
End Function

' Consente data valida OPPURE stringa speciale "nnn" (case-insensitive)
Private Function HasValidDateOrNNN(ByVal s As String) As Boolean
    s = Trim$(s)
    If Len(s) = 0 Then
        HasValidDateOrNNN = False
    ElseIf LCase$(s) = "nnn" Then
        HasValidDateOrNNN = True
    Else
        HasValidDateOrNNN = IsDate(s)
    End If
End Function

' Formatta per l'editing: se è data ? "dd/mm/yyyy"; altrimenti restituisce la stringa così com'è
Private Function FormatDateOrTextForEdit(ByVal v As Variant) As String
    If IsDate(v) Then
        FormatDateOrTextForEdit = Format(CDate(v), "dd/mm/yyyy")
    ElseIf IsError(v) Or IsEmpty(v) Then
        FormatDateOrTextForEdit = vbNullString
    Else
        FormatDateOrTextForEdit = CStr(v)
    End If
End Function

' Scrive una DATA oppure la stringa "nnn" nella cella della tabella
' - Se "nnn": forza NumberFormat="@" sulla cella e scrive testo
' - Se data: imposta NumberFormat data e scrive il valore Date
Private Sub SetFieldDateOrNNN(ByVal headerName As String, ByVal txt As String, ByVal index1Based As Long)
    Dim cIdx As Long: cIdx = GetColumnIndex(headerName)
    If cIdx = 0 Then Exit Sub
    If mLo.DataBodyRange Is Nothing Then Exit Sub
    
    Dim s As String: s = Trim$(txt)
    With mLo.DataBodyRange.Cells(index1Based, cIdx)
        If LCase$(s) = "nnn" Then
            .NumberFormat = "@"
            .value = "nnn"
        ElseIf Len(s) > 0 And IsDate(s) Then
            .NumberFormat = "dd/mm/yyyy"
            .value = CDate(s)
        Else
            .value = Empty
        End If
    End With
End Sub

' ============ Helpers di accesso tabella ============

Private Function GetColumnIndex(ByVal headerName As String) As Long
    Dim k As String: k = Normalize(headerName)
    If mHdrMap.exists(k) Then GetColumnIndex = mHdrMap(k) Else GetColumnIndex = 0
End Function

Private Function GetColumnRange(ByVal headerName As String) As Range
    Dim idx As Long: idx = GetColumnIndex(headerName)
    If idx = 0 Then Exit Function
    If mLo.DataBodyRange Is Nothing Then Exit Function
    Set GetColumnRange = mLo.DataBodyRange.Columns(idx)
End Function

Private Sub SetFieldText(ByVal headerName As String, ByVal value As String, ByVal index1Based As Long)
    ' Scrive SEMPRE TESTO
    Dim cIdx As Long: cIdx = GetColumnIndex(headerName)
    If cIdx = 0 Then Exit Sub
    If mLo.DataBodyRange Is Nothing Then Exit Sub
    mLo.DataBodyRange.Cells(index1Based, cIdx).NumberFormat = "@"   ' assicura testo per la cella
    mLo.DataBodyRange.Cells(index1Based, cIdx).value = CStr(value)
End Sub

Private Sub SetFieldDate(ByVal headerName As String, ByVal txt As String, ByVal index1Based As Long)
    Dim cIdx As Long: cIdx = GetColumnIndex(headerName)
    If cIdx = 0 Then Exit Sub
    If mLo.DataBodyRange Is Nothing Then Exit Sub
    If Len(Trim$(txt)) > 0 And IsDate(txt) Then
        mLo.DataBodyRange.Cells(index1Based, cIdx).value = CDate(txt)
    Else
        mLo.DataBodyRange.Cells(index1Based, cIdx).value = Empty
    End If
End Sub

Private Function ReadField(ByVal headerName As String, ByVal index1Based As Long) As Variant
    Dim cIdx As Long: cIdx = GetColumnIndex(headerName)
    If cIdx = 0 Then ReadField = Empty: Exit Function
    If mLo.DataBodyRange Is Nothing Then ReadField = Empty: Exit Function
    ReadField = mLo.DataBodyRange.Cells(index1Based, cIdx).value
End Function

' ============ Caricamento / Pulizia form ============

Private Sub ClearForm()
    cboUbicazione.value = vbNullString
    txtNumber.text = vbNullString
    txtDate.text = vbNullString
    cboScheda.value = vbNullString
    txtComments.text = vbNullString
    txtSerialNumber.text = vbNullString
    cboManufacturer.value = vbNullString
    cboModel.value = vbNullString
    txtDoM.text = vbNullString
    txtDoP.text = vbNullString
    txtDoFU.text = vbNullString
    txtNextInsp.text = vbNullString
    txtRetirement.text = vbNullString
    txtAnnotazioni.text = vbNullString
    txtCustomer.text = GetImpostazione("Cliente")
    cboResult.value = vbNullString
End Sub

Private Sub LoadFromIndex(ByVal index1Based As Long)
    On Error Resume Next
    cboUbicazione.value = NzStr(ReadField(HDR_UBIC, index1Based))
    txtNumber.text = NzStr(ReadField(HDR_NUMBER, index1Based))
    txtDate.text = FormatDateForEdit(ReadField(HDR_DATE, index1Based))
    cboScheda.value = EnsureInComboAndReturn(cboScheda, NzStr(ReadField(HDR_SCHEDA, index1Based)))
    txtComments.text = NzStr(ReadField(HDR_COMMENTS, index1Based))
    txtSerialNumber.text = NzStr(ReadField(HDR_SERIAL, index1Based))
    cboManufacturer.value = EnsureInComboAndReturn(cboManufacturer, NzStr(ReadField(HDR_MANUF, index1Based)))
    cboModel.value = EnsureInComboAndReturn(cboModel, NzStr(ReadField(HDR_MODEL, index1Based)))
    txtDoM.text = FormatDateForEdit(ReadField(HDR_DOM, index1Based))
    txtDoP.text = FormatDateForEdit(ReadField(HDR_DOP, index1Based))
    txtDoFU.text = FormatDateForEdit(ReadField(HDR_DOFU, index1Based))
    txtNextInsp.text = FormatDateOrTextForEdit(ReadField(HDR_NEXTINSP, index1Based))
    txtRetirement.text = FormatDateForEdit(ReadField(HDR_RETIRE, index1Based))
    txtAnnotazioni.text = NzStr(ReadField(HDR_ANN, index1Based))
    txtCustomer.text = NzStr(ReadField(HDR_CUST, index1Based))
    cboResult.value = EnsureInComboAndReturn(cboResult, LCase$(NzStr(ReadField(HDR_RESULT, index1Based))))
    On Error GoTo 0
End Sub

' ============ Eventi input ============

Private Sub txtNumber_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    ' Solo cifre e backspace
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii < 48 Or KeyAscii > 57 Then KeyAscii = 0
End Sub

Private Sub txtNumber_BeforeUpdate(ByVal Cancel As MSForms.ReturnBoolean)
    'If Len(Trim$(txtNumber.text)) = 0 Or Not IsAllDigits(txtNumber.text) Then
    '    MsgBox "'Number' deve contenere solo cifre.", vbExclamation
    '    Cancel = True
    'End If
End Sub

' ============ Utilità varie ============

Private Function Normalize(ByVal s As String) As String
    Normalize = LCase$(Trim$(s))
End Function

Private Function NzStr(ByVal v As Variant) As String
    If IsError(v) Or IsEmpty(v) Or v = "" Then NzStr = vbNullString Else NzStr = CStr(v)
End Function

Private Function FormatDateForEdit(ByVal v As Variant) As String
    If IsDate(v) Then FormatDateForEdit = Format(CDate(v), "dd/mm/yyyy") Else FormatDateForEdit = vbNullString
End Function

Private Function EnsureInComboAndReturn(cb As MSForms.ComboBox, ByVal val As String) As String
    Dim i As Long
    If Len(val) = 0 Then EnsureInComboAndReturn = vbNullString: Exit Function
    For i = 0 To cb.ListCount - 1
        If StrComp(cb.List(i), val, vbTextCompare) = 0 Then EnsureInComboAndReturn = val: Exit Function
    Next
    cb.AddItem val
    EnsureInComboAndReturn = val
End Function

Private Sub EnsureTextFormatForColumns(headers As Variant)
    On Error Resume Next
    Dim H As Variant, idx As Long
    For Each H In headers
        idx = GetColumnIndex(CStr(H))
        If idx > 0 Then
            ' Applica alla colonna della ListObject (header + corpo), così anche nuove righe ereditano il formato
            mLo.ListColumns(idx).Range.NumberFormat = "@"
        End If
    Next
    On Error GoTo 0
End Sub

' --- Unicità Number (stringa) ---
Private Function IsUniqueNumberText(ByVal numText As String, ByVal ignoreIndex As Long) As Boolean
    Dim rng As Range, c As Range
    Set rng = GetColumnRange(HDR_NUMBER)
    If rng Is Nothing Then IsUniqueNumberText = True: Exit Function
    
    ' 1) Cerca come testo
    Set c = rng.Find(What:=numText, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    If Not c Is Nothing Then
        Dim firstAddr As String: firstAddr = c.Address
        Do
            Dim idx As Long: idx = c.row - rng.Rows(1).row + 1
            If idx <> ignoreIndex Then IsUniqueNumberText = False: Exit Function
            Set c = rng.FindNext(c)
        Loop While Not c Is Nothing And c.Address <> firstAddr
    End If
    
    ' 2) Fallback: cerca come numero intero (caso storico)
    If IsNumeric(numText) Then
        Set c = rng.Find(What:=CLng(numText), LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
        If Not c Is Nothing Then
            Dim first2 As String: first2 = c.Address
            Do
                Dim idx2 As Long: idx2 = c.row - rng.Rows(1).row + 1
                If idx2 <> ignoreIndex Then IsUniqueNumberText = False: Exit Function
                Set c = rng.FindNext(c)
            Loop While Not c Is Nothing And c.Address <> first2
        End If
    End If
    
    IsUniqueNumberText = True
End Function

' ---- Loader da fogli esterni ----

Private Sub LoadSchedeFromAzioniDPI(cb As MSForms.ComboBox)
    On Error GoTo Fine
    cb.Clear
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets("Azioni_DPI")
    Dim rngH As Range: Set rngH = ws.Rows(1)
    Dim colID As Long, colTipo As Long, lastRow As Long, r As Long
    colID = FindHeaderColumn(rngH, "ID")
    colTipo = FindHeaderColumn(rngH, "Tipo DPI")
    If colID = 0 Or colTipo = 0 Then Exit Sub
    
    lastRow = ws.Cells(ws.Rows.count, colID).End(xlUp).row
    For r = 2 To lastRow
        If Len(Trim$(ws.Cells(r, colID).value)) > 0 Then
            cb.AddItem ws.Cells(r, colID).value & " - " & ws.Cells(r, colTipo).value
        End If
    Next
Fine:
End Sub

Private Sub LoadDistinctFromSheetColumn(cb As MSForms.ComboBox, ByVal sheetName As String, ByVal header As String)
    On Error GoTo Fine
    cb.Clear
    Dim ws As Worksheet: Set ws = ThisWorkbook.Worksheets(sheetName)
    Dim rngH As Range: Set rngH = ws.Rows(1)
    Dim col As Long: col = FindHeaderColumn(rngH, header)
    If col = 0 Then Exit Sub
    
    Dim dict As Object: Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare
    Dim lastRow As Long: lastRow = ws.Cells(ws.Rows.count, col).End(xlUp).row
    Dim r As Long, v As Variant, key As String
    For r = 2 To lastRow
        v = ws.Cells(r, col).value
        key = Trim$(LCase$(CStr(v)))
        If Len(key) > 0 Then If Not dict.exists(key) Then dict(key) = CStr(v)
    Next
    Dim k As Variant
    For Each k In dict.Keys
        cb.AddItem dict(k)
    Next
Fine:
End Sub

Private Function FindHeaderColumn(rngHeaderRow As Range, ByVal headerName As String) As Long
    Dim c As Range
    Set c = rngHeaderRow.Find(What:=headerName, LookIn:=xlValues, LookAt:=xlWhole, MatchCase:=False)
    If c Is Nothing Then FindHeaderColumn = 0 Else FindHeaderColumn = c.Column
End Function


'*************************
' date Picker
'*************************

' ---------- API per parsing basilare ----------
Private Function TryParseDate(ByVal s As String, ByRef outDate As Date) As Boolean
    Dim T As String: T = Trim$(s)
    If Len(T) = 0 Then Exit Function
    On Error Resume Next
    If IsDate(T) Then
        outDate = CDate(T)
        TryParseDate = True
        Exit Function
    End If
    ' Prova formati comuni: ISO e dd/mm/yyyy (accetta anche - e .)
    Dim p() As String, y As Long, m As Long, d As Long
    If T Like "####-##-##" Or T Like "####/##/##" Then
        T = Replace(T, "/", "-"): p = Split(T, "-")
        If UBound(p) = 2 Then
            y = val(p(0)): m = val(p(1)): d = val(p(2))
            outDate = DateSerial(y, m, d)
            If Err.Number = 0 Then TryParseDate = True
        End If
        Err.Clear
        Exit Function
    End If
    T = Replace(Replace(T, ".", "/"), "-", "/")
    p = Split(T, "/")
    If UBound(p) = 2 Then
        d = val(p(0)): m = val(p(1)): y = val(p(2))
        outDate = DateSerial(y, m, d)
        If Err.Number = 0 Then TryParseDate = True
        Err.Clear
    End If
    On Error GoTo 0
End Function

Private Function DaysInMonth(ByVal y As Long, ByVal m As Long) As Long
    DaysInMonth = Day(DateSerial(y, m + 1, 0))
End Function

' ---------- API pubblica: apri il datepicker per un TextBox ----------
Public Sub OpenInlineDatePicker(ByVal targetTB As MSForms.TextBox, _
                                Optional ByVal dateFormat As String = "dd/mm/yyyy")
    ' Salva il TextBox target
    Set mDP_TargetTB = targetTB
    pnlDate.Tag = dateFormat  ' memorizzo il formato scelto nel Tag del pannello

    ' Data base: dal TextBox se valida, altrimenti oggi
    Dim baseDate As Date
    If Not TryParseDate(targetTB.value, baseDate) Then baseDate = Date

    mDP_Year = Year(baseDate)
    mDP_Month = Month(baseDate)
    mDP_Day = Day(baseDate)

    ' Popola UI
    LoadDP_Months
    LoadDP_Years mDP_Year
    RefreshDP_Days
    UpdateDP_Title

    ' Seleziona mese/anno correnti nei combo
    cboDP_Month.ListIndex = mDP_Month - 1
    Dim yMin As Long: yMin = CLng(cboDP_Year.List(0))
    If mDP_Year >= yMin And mDP_Year <= CLng(cboDP_Year.List(cboDP_Year.ListCount - 1)) Then
        cboDP_Year.ListIndex = mDP_Year - yMin
    End If

    ' Posiziona e mostra pannello centrato nel form (semplice e robusto)
    'CenterPanel pnlDate, Me
    'pnlDate.Visible = True
    'pnlDate.ZOrder 0
        
    pnlDate.Visible = True            ' prima rendilo visibile per avere Width/Height effettivi
    pnlDate.ZOrder 0
    PlacePanelNearTextBox targetTB, pnlDate, 6

    ' (Opzionale) disattiva momentaneamente altri controlli:
    ' Me.Enabled = False : pnlDate.Enabled = True   ' sconsigliato: “ingrigisce” tutto
End Sub

' ---------- Costruzione contenuti ----------
Private Sub LoadDP_Months()
    Dim i As Long
    cboDP_Month.Clear
    For i = 1 To 12
        cboDP_Month.AddItem Format$(DateSerial(2000, i, 1), "mmmm")
    Next i
End Sub

Private Sub LoadDP_Years(ByVal aroundYear As Long)
    Dim y As Long, yMin As Long, yMax As Long
    yMin = aroundYear - 15: If yMin < 1900 Then yMin = 1900
    yMax = aroundYear + 150: If yMax > 2200 Then yMax = 2200

    cboDP_Year.Clear
    For y = yMin To yMax
        cboDP_Year.AddItem CStr(y)
    Next y
End Sub

Private Sub RefreshDP_Days()
    Dim i As Long, n As Long
    lstDP_Day.Clear
    n = DaysInMonth(mDP_Year, mDP_Month)
    For i = 1 To n
        lstDP_Day.AddItem CStr(i)
    Next i

    ' Preseleziona il giorno se presente
    If mDP_Day >= 1 And mDP_Day <= n Then
        lstDP_Day.ListIndex = mDP_Day - 1
    ElseIf n > 0 Then
        lstDP_Day.ListIndex = 0
        mDP_Day = 1
    End If
    
    'pnlDate.Caption = UCase$(Format$(DateSerial(mDP_Year, mDP_Month, mDP_Day), "dddd dd mmmm yyyy"))
    'lblDP_Title.Caption = UCase$(Format$(DateSerial(mDP_Year, mDP_Month, 1), "mmmm yyyy"))
End Sub

Private Sub UpdateDP_Title()
    Dim d As Date
    If mDP_Day >= 1 And mDP_Day <= DaysInMonth(mDP_Year, mDP_Month) Then
        d = DateSerial(mDP_Year, mDP_Month, mDP_Day)
    Else
        d = DateSerial(mDP_Year, mDP_Month, 1)
    End If

    pnlDate.Caption = UCase$(Format$(d, "dddd dd mmmm yyyy"))
End Sub


' Posiziona pnlDate accanto al TextBox: SOTTO ? SOPRA ? DESTRA ? SINISTRA
Private Sub PlacePanelNearTextBox(ByVal tb As MSForms.TextBox, ByVal panel As MSForms.Frame, _
                                  Optional ByVal margin As Single = 6)
    Dim ctlL As Single, ctlT As Single
    GetAbsPosInForm tb, ctlL, ctlT

    ' Candidate positions
    Dim candBelowL As Single, candBelowT As Single
    Dim candAboveL As Single, candAboveT As Single
    Dim candRightL As Single, candRightT As Single
    Dim candLeftL As Single, candLeftT As Single
    Dim belowOK As Boolean, aboveOK As Boolean, rightOK As Boolean, leftOK As Boolean
    
    ' SOTTO
    candBelowL = ctlL
    candBelowT = ctlT + tb.Height + margin
    belowOK = (candBelowT + panel.Height <= Me.InsideHeight)
    ' SOPRA
    candAboveL = ctlL
    candAboveT = ctlT - panel.Height - margin
    aboveOK = (candAboveT >= 0)
    ' DESTRA
    candRightL = ctlL + tb.Width + margin
    candRightT = ctlT
    rightOK = (candRightL + panel.Width <= Me.InsideWidth)
    ' SINISTRA
    candLeftL = ctlL - panel.Width - margin
    candLeftT = ctlT
    leftOK = (candLeftL >= 0)

    Dim l As Single, T As Single
    If belowOK Then
        l = candBelowL: T = candBelowT
    ElseIf aboveOK Then
        l = candAboveL: T = candAboveT
    ElseIf rightOK Then
        l = candRightL: T = candRightT
    ElseIf leftOK Then
        l = candLeftL: T = candLeftT
    Else
        ' fallback: sotto e poi clamp
        l = candBelowL: T = candBelowT
    End If

    ' Se eccede lateralmente, centra rispetto al TextBox
    If l < 0 Or l + panel.Width > Me.InsideWidth Then
        l = ctlL + (tb.Width - panel.Width) / 2
    End If

    ' Mantieni dentro i bordi
    ClampInsideForm l, T, panel.Width, panel.Height

    ' Applica
    panel.left = l
    panel.top = T
End Sub


' ---------- Eventi dei controlli del DatePicker ----------
Private Sub cboDP_Month_Change()
    If cboDP_Month.ListIndex >= 0 Then
        mDP_Month = cboDP_Month.ListIndex + 1
        RefreshDP_Days
        UpdateDP_Title
    End If
End Sub

Private Sub cboDP_Year_Change()
    If IsNumeric(cboDP_Year.value) Then
        mDP_Year = CLng(cboDP_Year.value)
        If mDP_Year < 1900 Then mDP_Year = 1900
        If mDP_Year > 2099 Then mDP_Year = 2099
        RefreshDP_Days
        UpdateDP_Title
    End If
End Sub

Private Sub lstDP_Day_Click()
    If lstDP_Day.ListIndex >= 0 Then
        mDP_Day = lstDP_Day.ListIndex + 1
        UpdateDP_Title
    End If
End Sub

Private Sub lstDP_Day_Change()
    If lstDP_Day.ListIndex >= 0 Then
        mDP_Day = lstDP_Day.ListIndex + 1
        UpdateDP_Title
    End If
End Sub

Private Sub lstDP_Day_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    ' Doppio click = conferma
    cmdDP_OK_Click
End Sub

Private Sub cmdDP_Today_Click()
    Dim T As Date: T = Date
    mDP_Year = Year(T): mDP_Month = Month(T): mDP_Day = Day(T)
    LoadDP_Years mDP_Year
    LoadDP_Months
    cboDP_Month.ListIndex = mDP_Month - 1
    Dim yMin As Long: yMin = CLng(cboDP_Year.List(0))
    cboDP_Year.ListIndex = mDP_Year - yMin
    RefreshDP_Days
End Sub

Private Sub cmdDP_Clear_Click()
    If Not mDP_TargetTB Is Nothing Then mDP_TargetTB.value = ""
    CloseInlineDatePicker
End Sub

Private Sub cmdDP_OK_Click()
    On Error Resume Next
    If mDP_Day < 1 Then mDP_Day = 1
    Dim sel As Date: sel = DateSerial(mDP_Year, mDP_Month, mDP_Day)
    Dim fmt As String: fmt = IIf(Len(pnlDate.Tag) > 0, CStr(pnlDate.Tag), "dd/mm/yyyy")
    If Not mDP_TargetTB Is Nothing Then mDP_TargetTB.value = Format$(sel, fmt)
    CloseInlineDatePicker
End Sub

Private Sub cmdDP_Cancel_Click()
    CloseInlineDatePicker
End Sub

Private Sub CloseInlineDatePicker()
    pnlDate.Visible = False
    ' Opzionale: restituisci focus al TextBox target
    On Error Resume Next
    If Not mDP_TargetTB Is Nothing Then mDP_TargetTB.SetFocus
    On Error GoTo 0
    ' Pulisci referenza
    Set mDP_TargetTB = Nothing
End Sub

' === Risalita robusta dei container (Frame, MultiPage.Page, ecc.) ===
Private Function TryGetParent(ByVal ctrl As Object, ByRef outParent As Object) As Boolean
    On Error Resume Next
    Dim p As Object
    Set p = CallByName(ctrl, "Parent", VbGet) ' non tutti i controlli espongono Parent: gestiamo con CallByName
    If Err.Number = 0 Then
        Set outParent = p
        TryGetParent = Not (p Is Nothing)
    Else
        Set outParent = Nothing
        TryGetParent = False
    End If
    Err.Clear
    On Error GoTo 0
End Function

' Somma Left/Top risalendo i container fino al form corrente (Me)
Private Sub GetAbsPosInForm(ByVal target As Object, ByRef absL As Single, ByRef absT As Single)
    Dim c As Object, parentObj As Object
    absL = 0: absT = 0
    Set c = target
    Do While Not (c Is Nothing) And Not (c Is Me)
        absL = absL + c.left
        absT = absT + c.top
        If Not TryGetParent(c, parentObj) Then Exit Do
        Set c = parentObj
    Loop
End Sub

' Mantiene un rettangolo (L,T,W,H) dentro ai bordi interni del form
Private Sub ClampInsideForm(ByRef l As Single, ByRef T As Single, ByVal W As Single, ByVal H As Single)
    If l < 0 Then l = 0
    If T < 0 Then T = 0
    If l + W > Me.InsideWidth Then l = Me.InsideWidth - W
    If T + H > Me.InsideHeight Then T = Me.InsideHeight - H
End Sub

' ===========================================
' POW Program Sequencer v6 - Modulo VBA
' ===========================================
' Genera un UNICO programma MDB combinando
' le funzioni dei programmi 30, 31, 32, 33
' nella sequenza desiderata
' ===========================================
' v6: LOGICA CORRETTA
'     Soudure: 1 sola riga per il programma
'     Script_Prog: tutte le funzioni unificate
'     so_CodProg = sp_CodProg = finalProgNum
' ===========================================

Option Explicit

' Costanti
Private Const PROG_30 As String = "30IGNIT"
Private Const PROG_31 As String = "31NOWELD"
Private Const PROG_32 As String = "32WELD"
Private Const PROG_33 As String = "33DWNSLP"

' Struttura per memorizzare le informazioni di un programma
Private Type ProgramInfo
    Number As Integer
    Name As String
    NumFunctions As Integer
    MaxLineNumber As Integer
End Type

' Array con info programmi
Private Programs(30 To 33) As ProgramInfo

Sub InitializePrograms()
    Programs(30).Number = 30
    Programs(30).Name = PROG_30
    Programs(30).NumFunctions = 12
    Programs(30).MaxLineNumber = 11

    Programs(31).Number = 31
    Programs(31).Name = PROG_31
    Programs(31).NumFunctions = 39
    Programs(31).MaxLineNumber = 38

    Programs(32).Number = 32
    Programs(32).Name = PROG_32
    Programs(32).NumFunctions = 49
    Programs(32).MaxLineNumber = 48

    Programs(33).Number = 33
    Programs(33).Name = PROG_33
    Programs(33).NumFunctions = 49
    Programs(33).MaxLineNumber = 48
End Sub

Sub GenerateMDB()
    ' ===========================================
    ' PROCEDURA PRINCIPALE
    ' Genera il file MDB dalla sequenza specificata
    ' ===========================================

    Dim ws As Worksheet
    Dim wsConfig As Worksheet
    Dim sequence() As Integer
    Dim i As Integer
    Dim lastRow As Long
    Dim progNum As Variant
    Dim outputPath As String
    Dim sourcePath As String
    Dim totalFunctions As Integer
    Dim fileDate As String
    Dim missingFiles As String
    Dim finalProgNumber As Integer
    Dim finalProgName As String
    Dim userInput As String

    ' Inizializza
    Call InitializePrograms

    Set ws = ThisWorkbook.Sheets("Sequenza")

    ' Leggi percorso sorgente dal foglio Configurazione
    On Error Resume Next
    Set wsConfig = ThisWorkbook.Sheets("Configurazione")
    If Not wsConfig Is Nothing Then
        sourcePath = wsConfig.Range("B2").Value
    End If
    On Error GoTo 0

    ' Se non configurato, usa cartella predefinita
    If sourcePath = "" Or sourcePath = "default" Then
        sourcePath = ThisWorkbook.Path & "\Sorgenti"
    End If

    ' Verifica che la cartella esista
    If Dir(sourcePath, vbDirectory) = "" Then
        sourcePath = ThisWorkbook.Path
    End If

    ' Verifica esistenza file sorgente
    missingFiles = ""
    For i = 30 To 33
        If Dir(sourcePath & "\" & Programs(i).Name & ".mdb") = "" Then
            missingFiles = missingFiles & "  - " & Programs(i).Name & ".mdb" & vbCrLf
        End If
    Next i

    If missingFiles <> "" Then
        MsgBox "File sorgente mancanti in:" & vbCrLf & sourcePath & vbCrLf & vbCrLf & _
               missingFiles & vbCrLf & _
               "Assicurati di salvare i file MDB da Powin-PC2 nella cartella Sorgenti.", vbCritical
        Exit Sub
    End If

    ' Trova l'ultima riga con dati
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    If lastRow < 2 Then
        MsgBox "Nessun programma specificato nella sequenza!", vbExclamation
        Exit Sub
    End If

    ' Leggi la sequenza
    ReDim sequence(1 To lastRow - 1)
    totalFunctions = 0

    For i = 2 To lastRow
        progNum = ws.Cells(i, 1).Value

        If Not IsNumeric(progNum) Then
            MsgBox "Errore riga " & i & ": valore non numerico", vbCritical
            Exit Sub
        End If

        If progNum < 30 Or progNum > 33 Then
            MsgBox "Errore riga " & i & ": programma " & progNum & " non valido." & vbCrLf & _
                   "Valori ammessi: 30, 31, 32, 33", vbCritical
            Exit Sub
        End If

        sequence(i - 1) = CInt(progNum)
        totalFunctions = totalFunctions + Programs(CInt(progNum)).NumFunctions
    Next i

    ' Chiedi il numero del programma finale
    userInput = InputBox("Inserisci il NUMERO del programma finale:" & vbCrLf & vbCrLf & _
                         "Questo sara' l'identificativo del programma" & vbCrLf & _
                         "quando lo importerai in Powin-PC2." & vbCrLf & vbCrLf & _
                         "(Es: 1, 50, 99, etc.)", _
                         "Numero Programma Finale", "1")

    If userInput = "" Then
        Exit Sub
    End If

    If Not IsNumeric(userInput) Then
        MsgBox "Il numero programma deve essere un valore numerico!", vbCritical
        Exit Sub
    End If

    finalProgNumber = CInt(userInput)

    ' Chiedi il nome del programma finale
    finalProgName = InputBox("Inserisci il NOME del programma finale:" & vbCrLf & vbCrLf & _
                             "Questo sara' il nome visualizzato in Powin-PC2." & vbCrLf & vbCrLf & _
                             "(Es: SEQUENZA, CUSTOM, etc.)", _
                             "Nome Programma Finale", finalProgNumber & "SEQ")

    If finalProgName = "" Then
        Exit Sub
    End If

    ' Mostra riepilogo
    Dim msg As String
    msg = "PROGRAMMA FINALE:" & vbCrLf
    msg = msg & "  Numero: " & finalProgNumber & vbCrLf
    msg = msg & "  Nome: " & finalProgName & vbCrLf & vbCrLf
    msg = msg & "CARTELLA SORGENTI:" & vbCrLf & sourcePath & vbCrLf & vbCrLf
    msg = msg & "SEQUENZA FUNZIONI:" & vbCrLf

    Dim lineCount As Integer
    lineCount = 0
    For i = 1 To UBound(sequence)
        Dim filePath As String
        filePath = sourcePath & "\" & Programs(sequence(i)).Name & ".mdb"
        fileDate = Format(FileDateTime(filePath), "dd/mm/yyyy hh:mm")
        msg = msg & "  " & i & ". " & Programs(sequence(i)).Name & " (linee " & (lineCount + 1) & "-" & (lineCount + Programs(sequence(i)).MaxLineNumber) & ")" & vbCrLf
        lineCount = lineCount + Programs(sequence(i)).MaxLineNumber
    Next i

    msg = msg & vbCrLf & "Totale linee: " & lineCount & vbCrLf & vbCrLf
    msg = msg & "Procedere con la generazione?"

    If MsgBox(msg, vbYesNo + vbQuestion, "Conferma Generazione") = vbNo Then
        Exit Sub
    End If

    ' Chiedi percorso di output
    outputPath = Application.GetSaveAsFilename( _
        InitialFileName:=finalProgName & ".mdb", _
        FileFilter:="Database Access (*.mdb), *.mdb", _
        Title:="Salva file MDB")

    If outputPath = "False" Then
        Exit Sub
    End If

    ' Genera il file MDB
    Call CreateUnifiedProgram(sequence, sourcePath, outputPath, finalProgNumber, finalProgName)

End Sub

Sub CreateUnifiedProgram(sequence() As Integer, sourcePath As String, outputPath As String, finalProgNum As Integer, finalProgName As String)
    ' ===========================================
    ' Crea un UNICO programma MDB unificato
    ' con tutte le funzioni in sequenza
    ' ===========================================
    ' LOGICA v6:
    ' - Soudure: UNA SOLA riga per il programma
    ' - Script_Prog: TUTTE le funzioni unificate
    ' ===========================================

    Dim connTarget As Object
    Dim connSource As Object
    Dim rsSource As Object
    Dim i As Integer
    Dim progNum As Integer
    Dim sourceFile As String
    Dim currentLineOffset As Integer
    Dim connStr As String
    Dim fld As Object
    Dim sql As String
    Dim fieldNames As String
    Dim fieldValues As String
    Dim insertSQL As String
    Dim fieldValue As Variant
    Dim stepNum As Integer

    ' Connection string per file MDB
    connStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source="

    ' STEP 1: Copia il primo file come base
    stepNum = 1
    On Error GoTo ErrorHandler
    sourceFile = sourcePath & "\" & Programs(sequence(1)).Name & ".mdb"

    ' Elimina file esistente se presente
    On Error Resume Next
    Kill outputPath
    On Error GoTo ErrorHandler

    FileCopy sourceFile, outputPath

    ' STEP 2: Apri connessione al target
    stepNum = 2
    Set connTarget = CreateObject("ADODB.Connection")
    connTarget.Mode = 3 ' adModeReadWrite

    On Error Resume Next
    connTarget.Open connStr & outputPath

    If Err.Number <> 0 Then
        MsgBox "STEP 2 - Errore apertura file target:" & vbCrLf & _
               outputPath & vbCrLf & vbCrLf & _
               "Errore: " & Err.Description, vbCritical
        Exit Sub
    End If
    On Error GoTo ErrorHandler

    ' STEP 3: Aggiorna UNICA riga in Soudure
    stepNum = 3
    On Error Resume Next

    ' Aggiorna la tabella Soudure (1 sola riga)
    sql = "UPDATE Soudure SET so_CodProg = " & finalProgNum & ", so_LibProg = '" & finalProgName & "'"
    connTarget.Execute sql

    If Err.Number <> 0 Then
        Debug.Print "UPDATE Soudure failed: " & Err.Description
        Err.Clear
    End If

    ' Aggiorna la tabella Script_Prog del primo programma
    sql = "UPDATE Script_Prog SET sp_CodProg = " & finalProgNum
    connTarget.Execute sql

    If Err.Number <> 0 Then
        Debug.Print "UPDATE Script_Prog sp_CodProg failed: " & Err.Description
        Err.Clear
    End If

    On Error GoTo ErrorHandler

    ' Offset iniziale per sp_NumLig (dopo le linee del primo programma)
    currentLineOffset = Programs(sequence(1)).MaxLineNumber

    ' STEP 4+: Per ogni programma successivo, copia SOLO Script_Prog (NON Soudure!)
    For i = 2 To UBound(sequence)
        stepNum = 3 + i
        progNum = sequence(i)
        sourceFile = sourcePath & "\" & Programs(progNum).Name & ".mdb"

        ' Apri connessione al sorgente
        Set connSource = CreateObject("ADODB.Connection")
        On Error Resume Next
        connSource.Open connStr & sourceFile

        If Err.Number <> 0 Then
            MsgBox "Errore apertura file: " & sourceFile & vbCrLf & Err.Description, vbExclamation
            Err.Clear
            On Error GoTo ErrorHandler
            GoTo NextProgram
        End If
        On Error GoTo ErrorHandler

        ' Copia le righe da Script_Prog (le funzioni del programma)
        Set rsSource = CreateObject("ADODB.Recordset")
        On Error Resume Next
        rsSource.Open "SELECT * FROM Script_Prog ORDER BY sp_Rang", connSource, 3, 1

        If Err.Number = 0 And Not rsSource.EOF Then
            Do While Not rsSource.EOF
                fieldNames = ""
                fieldValues = ""

                For Each fld In rsSource.Fields
                    ' Salta campi auto-incremento o ID
                    If LCase(fld.Name) = "id" Or _
                       LCase(fld.Name) = "sp_id" Or _
                       (fld.Attributes And &H10) = &H10 Then
                        ' Skip
                    Else
                        ' Aggiungi nome campo
                        If fieldNames <> "" Then fieldNames = fieldNames & ", "
                        fieldNames = fieldNames & "[" & fld.Name & "]"

                        ' Determina il valore
                        If fld.Name = "sp_Rang" Then
                            ' Rinumera sp_Rang con offset
                            fieldValue = fld.Value + currentLineOffset
                        ElseIf fld.Name = "sp_CodProg" Then
                            fieldValue = finalProgNum
                        Else
                            fieldValue = fld.Value
                        End If

                        ' Aggiungi valore con formattazione corretta
                        If fieldValues <> "" Then fieldValues = fieldValues & ", "

                        If IsNull(fieldValue) Then
                            fieldValues = fieldValues & "NULL"
                        ElseIf fld.Type = 202 Or fld.Type = 200 Or fld.Type = 201 Then ' String types
                            fieldValues = fieldValues & "'" & Replace(CStr(fieldValue), "'", "''") & "'"
                        ElseIf fld.Type = 7 Then ' Date
                            fieldValues = fieldValues & "#" & CStr(fieldValue) & "#"
                        Else
                            fieldValues = fieldValues & CStr(fieldValue)
                        End If
                    End If
                Next fld

                ' Esegui INSERT in Script_Prog
                insertSQL = "INSERT INTO Script_Prog (" & fieldNames & ") VALUES (" & fieldValues & ")"
                connTarget.Execute insertSQL
                If Err.Number <> 0 Then
                    Debug.Print "Errore INSERT Script_Prog: " & Err.Description
                    Err.Clear
                End If

                rsSource.MoveNext
            Loop
        End If

        If Err.Number <> 0 Then Err.Clear
        On Error GoTo ErrorHandler

        If Not rsSource Is Nothing Then
            rsSource.Close
            Set rsSource = Nothing
        End If

        ' Aggiorna offset per il prossimo programma
        currentLineOffset = currentLineOffset + Programs(progNum).MaxLineNumber

        ' Chiudi connessione sorgente
        connSource.Close
        Set connSource = Nothing

NextProgram:
    Next i

    ' Chiudi connessione target
    connTarget.Close
    Set connTarget = Nothing

    MsgBox "Programma unificato generato con successo!" & vbCrLf & vbCrLf & _
           "Numero: " & finalProgNum & vbCrLf & _
           "Nome: " & finalProgName & vbCrLf & _
           "File: " & outputPath, vbInformation

    Exit Sub

ErrorHandler:
    Dim errMsg As String
    errMsg = "Errore durante la generazione (STEP " & stepNum & "):" & vbCrLf & vbCrLf
    errMsg = errMsg & "Errore: " & Err.Description & vbCrLf
    errMsg = errMsg & "Numero: " & Err.Number & vbCrLf & vbCrLf

    Select Case stepNum
        Case 1
            errMsg = errMsg & "Problema: Copia file sorgente" & vbCrLf
        Case 2
            errMsg = errMsg & "Problema: Apertura connessione al file target" & vbCrLf
        Case 3
            errMsg = errMsg & "Problema: UPDATE tabella Soudure" & vbCrLf
        Case Else
            errMsg = errMsg & "Problema: Copia funzioni da programmi successivi" & vbCrLf
    End Select

    errMsg = errMsg & vbCrLf & "Possibili soluzioni:" & vbCrLf
    errMsg = errMsg & "1. Verifica che il file non sia aperto in Access" & vbCrLf
    errMsg = errMsg & "2. Installa Microsoft Access Database Engine 64-bit" & vbCrLf
    errMsg = errMsg & "   https://www.microsoft.com/en-us/download/details.aspx?id=54920"

    MsgBox errMsg, vbCritical

    ' Cleanup
    On Error Resume Next
    If Not rsSource Is Nothing Then rsSource.Close
    If Not connSource Is Nothing Then connSource.Close
    If Not connTarget Is Nothing Then connTarget.Close
End Sub

Sub ClearSequence()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sequenza")

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    If lastRow > 1 Then
        ws.Range("A2:A" & lastRow).ClearContents
    End If

    MsgBox "Sequenza cancellata.", vbInformation
End Sub

Sub AddDefaultSequence()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Sequenza")

    ws.Range("A2").Value = 30
    ws.Range("A3").Value = 31
    ws.Range("A4").Value = 32
    ws.Range("A5").Value = 33

    MsgBox "Sequenza standard aggiunta: 30, 31, 32, 33", vbInformation
End Sub

Sub ShowHelp()
    Dim msg As String

    msg = "POW PROGRAM SEQUENCER v3 - GUIDA" & vbCrLf & vbCrLf
    msg = msg & "FUNZIONAMENTO:" & vbCrLf
    msg = msg & "Questo tool crea UN SINGOLO PROGRAMMA che contiene" & vbCrLf
    msg = msg & "tutte le funzioni dei programmi 30-33 in sequenza." & vbCrLf & vbCrLf
    msg = msg & "COME USARE:" & vbCrLf
    msg = msg & "1. Salva i file MDB nella cartella Sorgenti" & vbCrLf
    msg = msg & "2. Inserisci la sequenza in colonna A (30, 31, 32, 33)" & vbCrLf
    msg = msg & "3. Esegui 'GenerateMDB'" & vbCrLf
    msg = msg & "4. Inserisci numero e nome del programma finale" & vbCrLf
    msg = msg & "5. Scegli dove salvare il file MDB" & vbCrLf & vbCrLf
    msg = msg & "RISULTATO:" & vbCrLf
    msg = msg & "Un file MDB con un unico programma contenente" & vbCrLf
    msg = msg & "tutte le funzioni nell'ordine specificato." & vbCrLf & vbCrLf
    msg = msg & "ESEMPIO:" & vbCrLf
    msg = msg & "Sequenza: 30, 32, 33" & vbCrLf
    msg = msg & "Programma finale N. 1 'MYSEQ'" & vbCrLf
    msg = msg & "  Linee 1-11: funzioni di 30IGNIT" & vbCrLf
    msg = msg & "  Linee 12-59: funzioni di 32WELD" & vbCrLf
    msg = msg & "  Linee 60-107: funzioni di 33DWNSLP"

    MsgBox msg, vbInformation, "Guida POW Sequencer v3"
End Sub

Sub SelectSourceFolder()
    Dim wsConfig As Worksheet
    Dim folderPath As String

    On Error Resume Next
    Set wsConfig = ThisWorkbook.Sheets("Configurazione")
    On Error GoTo 0

    If wsConfig Is Nothing Then
        MsgBox "Foglio Configurazione non trovato!", vbCritical
        Exit Sub
    End If

    With Application.FileDialog(msoFileDialogFolderPicker)
        .Title = "Seleziona cartella con i file MDB sorgente"
        .InitialFileName = ThisWorkbook.Path & "\Sorgenti\"
        If .Show = -1 Then
            folderPath = .SelectedItems(1)
            wsConfig.Range("B2").Value = folderPath

            Dim fileList As String
            fileList = ""
            If Dir(folderPath & "\30IGNIT.mdb") <> "" Then fileList = fileList & "  OK - 30IGNIT.mdb" & vbCrLf
            If Dir(folderPath & "\31NOWELD.mdb") <> "" Then fileList = fileList & "  OK - 31NOWELD.mdb" & vbCrLf
            If Dir(folderPath & "\32WELD.mdb") <> "" Then fileList = fileList & "  OK - 32WELD.mdb" & vbCrLf
            If Dir(folderPath & "\33DWNSLP.mdb") <> "" Then fileList = fileList & "  OK - 33DWNSLP.mdb" & vbCrLf

            If fileList = "" Then
                MsgBox "Cartella selezionata:" & vbCrLf & folderPath & vbCrLf & vbCrLf & _
                       "ATTENZIONE: Nessun file MDB trovato!", vbExclamation
            Else
                MsgBox "Cartella selezionata:" & vbCrLf & folderPath & vbCrLf & vbCrLf & _
                       "File trovati:" & vbCrLf & fileList, vbInformation
            End If
        End If
    End With
End Sub

Sub DiagnoseDatabase()
    ' ===========================================
    ' Diagnostica struttura database MDB
    ' Mostra tabelle e campi disponibili
    ' ===========================================

    Dim conn As Object
    Dim rs As Object
    Dim connStr As String
    Dim sourcePath As String
    Dim wsConfig As Worksheet
    Dim msg As String
    Dim filePath As String
    Dim fld As Object

    Call InitializePrograms

    ' Leggi percorso sorgente
    On Error Resume Next
    Set wsConfig = ThisWorkbook.Sheets("Configurazione")
    If Not wsConfig Is Nothing Then
        sourcePath = wsConfig.Range("B2").Value
    End If
    On Error GoTo 0

    If sourcePath = "" Or sourcePath = "default" Then
        sourcePath = ThisWorkbook.Path & "\Sorgenti"
    End If

    ' Usa il primo file disponibile per la diagnosi
    filePath = sourcePath & "\30IGNIT.mdb"
    If Dir(filePath) = "" Then
        filePath = sourcePath & "\31NOWELD.mdb"
    End If
    If Dir(filePath) = "" Then
        filePath = sourcePath & "\32WELD.mdb"
    End If
    If Dir(filePath) = "" Then
        filePath = sourcePath & "\33DWNSLP.mdb"
    End If

    If Dir(filePath) = "" Then
        MsgBox "Nessun file MDB trovato in:" & vbCrLf & sourcePath, vbCritical
        Exit Sub
    End If

    connStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & filePath

    Set conn = CreateObject("ADODB.Connection")
    On Error Resume Next
    conn.Open connStr

    If Err.Number <> 0 Then
        MsgBox "Errore connessione:" & vbCrLf & Err.Description, vbCritical
        Exit Sub
    End If
    On Error GoTo 0

    msg = "FILE: " & filePath & vbCrLf & vbCrLf

    ' Lista tabelle
    msg = msg & "TABELLE DISPONIBILI:" & vbCrLf
    Set rs = conn.OpenSchema(20) ' adSchemaTables
    Do While Not rs.EOF
        If rs("TABLE_TYPE") = "TABLE" Then
            msg = msg & "  - " & rs("TABLE_NAME") & vbCrLf
        End If
        rs.MoveNext
    Loop
    rs.Close

    ' Mostra campi della tabella Soudure (se esiste)
    msg = msg & vbCrLf & "CAMPI TABELLA SOUDURE:" & vbCrLf
    On Error Resume Next
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open "SELECT TOP 1 * FROM Soudure", conn, 3, 1

    If Err.Number <> 0 Then
        msg = msg & "  ERRORE: " & Err.Description & vbCrLf
        Err.Clear
    Else
        For Each fld In rs.Fields
            msg = msg & "  - " & fld.Name & " (" & fld.Type & ")" & vbCrLf
        Next fld
        rs.Close
    End If

    ' Mostra campi della tabella Script_Prog
    msg = msg & vbCrLf & "CAMPI TABELLA SCRIPT_PROG:" & vbCrLf
    Set rs = CreateObject("ADODB.Recordset")
    rs.Open "SELECT TOP 1 * FROM Script_Prog", conn, 3, 1

    If Err.Number <> 0 Then
        msg = msg & "  ERRORE: " & Err.Description & vbCrLf
        Err.Clear
    Else
        For Each fld In rs.Fields
            msg = msg & "  - " & fld.Name & " (" & fld.Type & ")" & vbCrLf
        Next fld
        rs.Close
    End If
    On Error GoTo 0

    conn.Close
    Set conn = Nothing

    MsgBox msg, vbInformation, "Diagnosi Database"
End Sub

Sub CheckSourceFiles()
    Dim wsConfig As Worksheet
    Dim sourcePath As String
    Dim msg As String
    Dim i As Integer

    Call InitializePrograms

    On Error Resume Next
    Set wsConfig = ThisWorkbook.Sheets("Configurazione")
    If Not wsConfig Is Nothing Then
        sourcePath = wsConfig.Range("B2").Value
    End If
    On Error GoTo 0

    If sourcePath = "" Or sourcePath = "default" Then
        sourcePath = ThisWorkbook.Path & "\Sorgenti"
    End If

    msg = "CARTELLA SORGENTI:" & vbCrLf & sourcePath & vbCrLf & vbCrLf
    msg = msg & "STATO FILE:" & vbCrLf

    For i = 30 To 33
        Dim filePath As String
        filePath = sourcePath & "\" & Programs(i).Name & ".mdb"

        If Dir(filePath) <> "" Then
            msg = msg & "  OK - " & Programs(i).Name & ".mdb" & vbCrLf
            msg = msg & "       Modificato: " & Format(FileDateTime(filePath), "dd/mm/yyyy hh:mm") & vbCrLf
        Else
            msg = msg & "  MANCANTE - " & Programs(i).Name & ".mdb" & vbCrLf
        End If
    Next i

    MsgBox msg, vbInformation, "Stato File Sorgente"
End Sub

' ===========================================
' POW Program Sequencer - Modulo VBA
' ===========================================
' Questo modulo genera un file MDB combinando
' i programmi 30, 31, 32, 33 nella sequenza desiderata
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
    ' Inizializza le informazioni dei programmi base

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
    Dim sequence() As Integer
    Dim i As Integer
    Dim lastRow As Long
    Dim progNum As Variant
    Dim outputPath As String
    Dim sourcePath As String
    Dim totalFunctions As Integer

    ' Inizializza
    Call InitializePrograms

    Set ws = ThisWorkbook.Sheets("Sequenza")

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

        ' Valida il numero programma
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

    ' Mostra riepilogo
    Dim msg As String
    msg = "Sequenza programmi:" & vbCrLf & vbCrLf

    For i = 1 To UBound(sequence)
        msg = msg & i & ". Programma " & sequence(i) & " (" & Programs(sequence(i)).Name & ")" & vbCrLf
    Next i

    msg = msg & vbCrLf & "Totale funzioni: " & totalFunctions & vbCrLf & vbCrLf
    msg = msg & "Procedere con la generazione del file MDB?"

    If MsgBox(msg, vbYesNo + vbQuestion, "Conferma Generazione") = vbNo Then
        Exit Sub
    End If

    ' Chiedi percorso di output
    outputPath = Application.GetSaveAsFilename( _
        InitialFileName:="ProgrammaSequenza.mdb", _
        FileFilter:="Database Access (*.mdb), *.mdb", _
        Title:="Salva file MDB")

    If outputPath = "False" Then
        Exit Sub
    End If

    ' Chiedi percorso sorgente (cartella con i file MDB originali)
    sourcePath = ThisWorkbook.Path
    If Dir(sourcePath & "\30IGNIT.mdb") = "" Then
        MsgBox "File sorgente 30IGNIT.mdb non trovato in:" & vbCrLf & sourcePath & vbCrLf & vbCrLf & _
               "Assicurati che i file MDB sorgente siano nella stessa cartella del file Excel.", vbCritical
        Exit Sub
    End If

    ' Genera il file MDB
    Call CreateSequencedMDB(sequence, sourcePath, outputPath)

End Sub

Sub CreateSequencedMDB(sequence() As Integer, sourcePath As String, outputPath As String)
    ' ===========================================
    ' Crea il file MDB combinato
    ' Usa DAO per manipolare i database Access
    ' ===========================================

    On Error GoTo ErrorHandler

    Dim dbSource As Object ' DAO.Database
    Dim dbTarget As Object ' DAO.Database
    Dim rsSource As Object ' DAO.Recordset
    Dim rsTarget As Object ' DAO.Recordset
    Dim daoEngine As Object
    Dim i As Integer
    Dim currentLineOffset As Integer
    Dim progNum As Integer
    Dim sourceFile As String

    ' Crea oggetto DAO
    Set daoEngine = CreateObject("DAO.DBEngine.36")

    ' Copia il primo file come base
    sourceFile = sourcePath & "\" & Programs(sequence(1)).Name & ".mdb"

    ' Copia file sorgente come base
    FileCopy sourceFile, outputPath

    ' Apri il database target
    Set dbTarget = daoEngine.OpenDatabase(outputPath, True)

    ' Offset iniziale per lineNumber
    currentLineOffset = Programs(sequence(1)).MaxLineNumber

    ' Per ogni programma successivo nella sequenza
    For i = 2 To UBound(sequence)
        progNum = sequence(i)
        sourceFile = sourcePath & "\" & Programs(progNum).Name & ".mdb"

        ' Apri database sorgente
        Set dbSource = daoEngine.OpenDatabase(sourceFile, False, True)

        ' Copia le funzioni con offset sui lineNumber
        Call CopyFunctionsWithOffset(dbSource, dbTarget, currentLineOffset)

        ' Aggiorna offset
        currentLineOffset = currentLineOffset + Programs(progNum).MaxLineNumber

        ' Chiudi sorgente
        dbSource.Close
        Set dbSource = Nothing
    Next i

    ' Chiudi target
    dbTarget.Close
    Set dbTarget = Nothing

    MsgBox "File MDB generato con successo!" & vbCrLf & vbCrLf & outputPath, vbInformation

    Exit Sub

ErrorHandler:
    MsgBox "Errore durante la generazione:" & vbCrLf & Err.Description, vbCritical

    ' Cleanup
    If Not dbSource Is Nothing Then dbSource.Close
    If Not dbTarget Is Nothing Then dbTarget.Close
End Sub

Sub CopyFunctionsWithOffset(dbSource As Object, dbTarget As Object, lineOffset As Integer)
    ' ===========================================
    ' Copia le funzioni da sorgente a target
    ' aggiungendo offset ai lineNumber
    ' ===========================================

    ' NOTA: Questa è una implementazione base.
    ' La struttura esatta delle tabelle POW deve essere
    ' verificata con Powin-PC2 per la corretta mappatura.

    Dim rsSource As Object
    Dim rsTarget As Object
    Dim sql As String
    Dim fld As Object

    ' Cerca la tabella principale (Soudure o simile)
    ' e copia i record aggiornando i numeri di linea

    On Error Resume Next

    ' Esempio: copia dalla tabella Soudure
    sql = "SELECT * FROM Soudure WHERE so_NumLigne > 0"
    Set rsSource = dbSource.OpenRecordset(sql)

    If Err.Number <> 0 Then
        ' Tabella non trovata, prova altre
        Err.Clear
        Exit Sub
    End If

    Set rsTarget = dbTarget.OpenRecordset("Soudure")

    Do While Not rsSource.EOF
        rsTarget.AddNew

        For Each fld In rsSource.Fields
            If fld.Name = "so_NumLigne" Then
                ' Aggiungi offset al numero di linea
                rsTarget(fld.Name) = fld.Value + lineOffset
            Else
                rsTarget(fld.Name) = fld.Value
            End If
        Next fld

        rsTarget.Update
        rsSource.MoveNext
    Loop

    rsSource.Close
    rsTarget.Close

    On Error GoTo 0
End Sub

Sub ClearSequence()
    ' Pulisce la sequenza corrente
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
    ' Aggiunge la sequenza standard 30-31-32-33
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

    msg = "POW PROGRAM SEQUENCER - GUIDA" & vbCrLf & vbCrLf
    msg = msg & "COME USARE:" & vbCrLf
    msg = msg & "1. Inserisci i numeri dei programmi nella colonna A" & vbCrLf
    msg = msg & "   (valori ammessi: 30, 31, 32, 33)" & vbCrLf & vbCrLf
    msg = msg & "2. L'ordine delle righe determina la sequenza" & vbCrLf
    msg = msg & "   di esecuzione delle funzioni" & vbCrLf & vbCrLf
    msg = msg & "3. Clicca 'Genera MDB' per creare il file" & vbCrLf & vbCrLf
    msg = msg & "PROGRAMMI DISPONIBILI:" & vbCrLf
    msg = msg & "  30 = IGNIT (Accensione) - 12 funzioni" & vbCrLf
    msg = msg & "  31 = NOWELD (No saldatura) - 39 funzioni" & vbCrLf
    msg = msg & "  32 = WELD (Saldatura) - 49 funzioni" & vbCrLf
    msg = msg & "  33 = DWNSLP (Downslope) - 49 funzioni" & vbCrLf & vbCrLf
    msg = msg & "ESEMPIO:" & vbCrLf
    msg = msg & "  Riga 2: 30  (prima IGNIT)" & vbCrLf
    msg = msg & "  Riga 3: 32  (poi WELD)" & vbCrLf
    msg = msg & "  Riga 4: 33  (infine DWNSLP)" & vbCrLf

    MsgBox msg, vbInformation, "Guida POW Sequencer"
End Sub

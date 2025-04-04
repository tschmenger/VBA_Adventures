Attribute VB_Name = "Vergleichen"
Sub ShowUserFormVergleichen()
    ' Show the user form
    UserFormVergleichen.Show
End Sub

Sub CompareWorkbooks()
    'Erstellt: ToS, 2025-03-07
    'Vergleicht alle gleichnamigen Arbeitsblätter zweier Arbeitsmappen und färbt Zellen, die unterschiedliche Werte haben, in beiden Arbeitsmappen gelb ein.
    ' Simpler als "Speardsheet Compare", sowohl in seiner Funktionalität aber eben auch in seiner Anwendung.
    ' Die zu vergleichenden Arbeitsmappen müssen vom Nutzer bereits geöffnet sein.
    ' ###################################################################################################
    'On Error GoTo err_handler: Const cPROC$ = "Vorlage"
    'Call mod_Logging.Logger(cMODULE, cPROC, etStart)
    ' ###################################################################################################
    Dim wb1 As Workbook
    Dim wb2 As Workbook
    Dim ws1 As Worksheet
    Dim ws2 As Worksheet
    Dim cell1 As Range
    Dim cell2 As Range
    Dim wsName As String
    Dim lastRow As Long
    Dim lastCol As Long
    ' ###################################################################################################
    Application.ScreenUpdating = False
    ' Debug-Ausgabe hinzufügen
    ' Debug.Print "CompareWorkbooks-Methode aufgerufen"
    
    ' Lege die ausgewählten Arbeitsmappen fest
    Set wb1 = Workbooks(UserFormVergleichen.ComboBoxVG1.Value)
    Set wb2 = Workbooks(UserFormVergleichen.ComboBoxVG2.Value)
    ' Überprüfen, ob die Arbeitsmappen korrekt geladen wurden
    If wb1 Is Nothing Then
        Debug.Print "Arbeitsmappe 1 nicht geladen"
        Exit Sub
    End If
    If wb2 Is Nothing Then
        Debug.Print "Arbeitsmappe 2 nicht geladen"
        Exit Sub
    End If
    ' ###################################################################################################
    Application.StatusBar = "Vergleich wird durchgeführt..."
    ' Durchlaufe alle Arbeitsblätter in der ersten Arbeitsmappe
    For Each ws1 In wb1.Worksheets
        wsName = ws1.Name
        Debug.Print "Vergleiche Arbeitsblatt: " & wsName
        ' Überprüfe, ob das Arbeitsblatt in der zweiten Arbeitsmappe existiert
        On Error Resume Next
        Set ws2 = wb2.Worksheets(wsName)
        On Error GoTo 0
        If Not ws2 Is Nothing Then
            ' Bestimme die letzte Zeile und die letzte Spalte
            lastRow = ws1.Cells(ws1.Rows.Count, 1).End(xlUp).Row
            lastCol = 20
            ' Debug-Ausgabe hinzufügen
            Debug.Print "Letzte Zeile: " & lastRow & ", Letzte Spalte: " & lastCol
            ' Durchlaufe alle Zellen und vergleiche die Werte
            For Each cell1 In ws1.Range(ws1.Cells(1, 1), ws1.Cells(lastRow, lastCol))
                Set cell2 = ws2.Cells(cell1.Row, cell1.Column)
                ' Debug.Print "Vergleiche Zelle: (" & cell1.Row & ", " & cell1.Column & ") Wert1: " & cell1.Value & " Wert2: " & cell2.Value
                ' Debug.Print cell1.Value & cell2.Value
                ' Check for errors before comparing cell values
                If Not IsError(cell1.Value) And Not IsError(cell2.Value) Then
                    If cell1.Value <> cell2.Value Then
                        cell1.Interior.Color = RGB(255, 255, 0)
                        cell2.Interior.Color = RGB(255, 255, 0)
                    End If
                Else
                    ' Färbe Zellen rot ein, falls diese Fehler verursachen (Ursprünglich schwierig bei Zellwerten wie "!Div/0" aber dieses Fehler sollte jetzt vermieden sein).
                    If IsError(cell1.Value) Then cell1.Interior.Color = RGB(255, 0, 0)
                    If IsError(cell2.Value) Then cell2.Interior.Color = RGB(255, 0, 0)
                End If
            Next cell1
        Else
            Debug.Print "Arbeitsblatt " & wsName & " nicht gefunden in der zweiten Arbeitsmappe"
        End If
        Set ws2 = Nothing
    Next ws1
    Application.ScreenUpdating = True
    Application.StatusBar = False
    MsgBox "Vergleich der Arbeitsmappen abgeschlossen."
    ' ###################################################################################################
    'Call mod_Logging.Logger(cMODULE, cPROC, etEnde)
    'Exit Sub
'err_handler:
    ' Call mod_Logging.Logger(cMODULE, cPROC, etFehler, Err.number, Err.Description)
    ' ###################################################################################################
End Sub

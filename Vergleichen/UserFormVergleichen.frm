VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormVergleichen 
   Caption         =   "Inhalte Vergleichen"
   ClientHeight    =   3300
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6450
   OleObjectBlob   =   "UserFormVergleichen.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "UserFormVergleichen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub UserForm_Initialize()
    Dim wb As Workbook
    Label_Vgl.Caption = "Achtung: Da jede Zelle verglichen wird, kann dieser Prozess bei gr��eren Tabellen etwas dauern." & vbNewLine & "Tip: Stelle sicher, dass beide Tabellen prinzipiell identisch formatiert sind."
    LabelVgl2.Caption = "W�hle aus bereits ge�ffneten Arbeitsmappen aus, um den Vergleich zu starten."
    For Each wb In Application.Workbooks
        If wb.Name <> "PERSONAL.XLSB" Then
            ComboBoxVG1.AddItem wb.Name
            ComboBoxVG2.AddItem wb.Name
        End If
    Next wb
    ' Set default selection
    If ComboBoxVG1.ListCount > 0 Then
        ComboBoxVG1.ListIndex = 0
    End If
    If ComboBoxVG2.ListCount > 0 Then
        ComboBoxVG2.ListIndex = 0
    End If
End Sub

Private Sub cmdVergleichen_Click()
    ' Debug-Ausgabe hinzuf�gen
    Debug.Print "Vergleichs-Button gedr�ckt"
    
    ' UserForm verstecken
    Me.Hide
    
    ' Vergleich starten
    CompareWorkbooks
End Sub

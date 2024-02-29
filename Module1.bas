Attribute VB_Name = "Module1"
Sub createSheetsBy_input()

'Notes: We can not create a sheet with no name.

    Dim sourceRange As Range 'Declare sourceRange as a Range for storing the range from where we are going to get the titles.
    Dim sourceCell As Range  'Declare sourceRange as a Range for storing the cells inside the range (sourceRange).
    On Error GoTo Errorhandling
    
    '---------- Spanish version ----------
    Set sourceRange = Application.InputBox(Prompt:="Selecciona el rango de celdas de donde vienen los títulos:", _
    Title:="Vamos a crear (sheet) hojas nuevas ", _
    Default:=Selection.Address, Type:=8)
    
    '---------- English version ----------
    'Set sourceRange = Application.InputBox(Prompt:="Select a cell range for the sheet titles.:", _
    Title:="Create sheets", _
    Default:=Selection.Address, Type:=8)
    
    For Each sourceCell In sourceRange 'Checks if the sourceCell variable is NOT empty. If the sourceCell variable is empty the procedure goes to "End If" line.
        If sourceCell <> "" Then
            Sheets.Add.Name = sourceCell 'Here we create new sheets named after the values stored in the sourceCell variable.
        End If
    Next sourceCell 'Go back to the "For each" statement and store a new single Cell in the sourceCell object.
Errorhandling:
End Sub



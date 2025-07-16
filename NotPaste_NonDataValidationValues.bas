Private Sub Worksheet_Change(ByVal Target As Range)

'Unallow changes in cell value, like copy or type data not belonging to the drop-down list
'Activated by trigger (any attempt to change cell value)

    Dim cell As Range
    Dim found As Boolean
    found = False
    Dim newValue As Variant
    Dim PreviousPoperties As Variant

  
    'Data validation lists
    ColumnA_Data = "A1:A3"
    
    
    'Column A: Data validation verification and not alllow pasting
    If Target.Column = 1 Then
    
        newValue = Target.Value 'Store the new value
    
        Application.EnableEvents = False 'Disable events to prevent infinite loop
        On Error GoTo ErrorHandler 'Handle errors
        Application.Undo 'Undo the change to get the previous properties
        On Error GoTo 0 'Reset error handling
        Target.Copy 'PreviousPoperties = Target.Copy 'Get previous properties
        Me.Range("ZZ1000000").PasteSpecial Paste:=xlPasteAll
        Application.CutCopyMode = False
    
    
        Target.Value = newValue  'Restore the new value
        Application.EnableEvents = True 'Re-enable events
        
        
        If Target.Value <> "" Then
            
            For Each cell In Worksheets(2).Range(ColumnA_Data)
                If Target.Value = cell.Value Then
                    found = True
                    Exit For
                End If
            Next cell
            
            If found = False Then
                MsgBox "Você não pode colar o valor """ & newValue & """ nesta célula. Use a lista suspensa.", vbExclamation
                Me.Range("ZZ1000000").Copy
                Me.Range("A" & Target.Row).PasteSpecial Paste:=xlPasteAll
                Application.CutCopyMode = False
            End If
            
            Me.Range("A" & Target.Row).Select
            
        End If
    
    End If
    
    Exit Sub
    
    
ErrorHandler:
    MsgBox "Recuperando valor da célula. Clique em OK.", vbCritical
    Application.EnableEvents = True
    
End Sub

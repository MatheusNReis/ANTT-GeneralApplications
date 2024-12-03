Sub MesclarCelulas()


'Merge cells vertically according to defined line steps and collumn

'Mescla células verticalmente conforme passos de linha e coluna configurados

'Created by Matheus Nunes Reis on 05/05/2024
'Copyright © 2024 Matheus Nunes Reis. All rights reserved.

'GitHub: MatheusNReis
'License: https://raw.githubusercontent.com/MatheusNReis/ANTT-GeneralApplications/3f27aa6f717d6861484d281eb74832db0c232e75/LICENSE
'MIT License. Copyright © 2024 MatheusNReis


    Dim ws As Worksheet
    Dim linhaInicial As Long
    Dim linhaFinal As Long
    Dim i As Long
    
    'Defina a planilha onde deseja mesclar as células
    Set ws = ThisWorkbook.ActiveSheet
    
    'Defina a linha final
    linhaFinal = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    
    'Loop para mesclar as células em pares
    For i = 7 To linhaFinal Step 2		' Modificar a linha conforme necessário
        ' Mesclar células
        ws.Range("A" & i & ":A" & i + 1).Merge	' Modificar a coluna conforme necessário
    Next i
    
    MsgBox "Células mescladas com sucesso!"
    
End Sub

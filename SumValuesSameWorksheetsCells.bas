Sub SomaValoresMesmaCelula()
    
'Sum values from the same defined cell for all existing worksheets in the workbook.

'Soma valores de uma mesma célula para todas as planilhas existentes na 
'pasta de trabalho.

'Created by Matheus Nunes Reis on 10/06/2024
'Copyright © 2024 Matheus Nunes Reis. All rights reserved.

'GitHub: MatheusNReis
'License: https://raw.githubusercontent.com/MatheusNReis/ANTT-GeneralApplications/3f27aa6f717d6861484d281eb74832db0c232e75/LICENSE
'MIT License. Copyright © 2024 MatheusNReis
    

    Dim ws As Worksheet
    Dim soma As Double

    somaM118 = 0
    somaM120 = 0
    QtdePlanilhas = 0
    QtdeKm = 0
    SomaLargura = 0

    ' Loop através de todas as planilhas
    For Each ws In ThisWorkbook.Worksheets
            
            QtdePlanilhas = QtdePlanilhas + 1
            QtdeKm = QtdeKm + Abs(ws.Range("C13").Value - ws.Range("E13").Value)
            SomaLargura = SomaLargura + ws.Range("A125")
            
        ' Verifique se a célula M120 contém um número
        If IsNumeric(ws.Range("M118").Value) Then
            ' Adicione o valor à soma
            somaM118 = somaM118 + ws.Range("M118").Value
        End If
        
        If IsNumeric(ws.Range("M120").Value) Then
            ' Adicione o valor à soma
            somaM120 = somaM120 + ws.Range("M120").Value
        End If
        
    Next ws

    ' Exibir resultados
    MsgBox "Soma todas  FC1+FC2+FC3 = " & somaM118 & vbNewLine & _
    "Soma todas FC2+FC3 = " & somaM120 & vbNewLine & _
    "Qtde Planilhas = " & QtdePlanilhas & vbNewLine & _
    "QtdeKm = " & QtdeKm & vbNewLine & _
    "SomaLargura = " & SomaLargura

End Sub

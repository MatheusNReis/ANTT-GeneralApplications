Sub CopiarConteudoOutroArquivo() 
   
    
'Copy and paste content from a file to the current file

'Copia e cola conteúdo de outro arquivo excel para o arquivo atual

'Created by Matheus Nunes Reis on 02/05/2024
'Copyright © 2024 Matheus Nunes Reis. All rights reserved.

'GitHub: MatheusNReis
'License: https://raw.githubusercontent.com/MatheusNReis/ANTT-GeneralApplications/3f27aa6f717d6861484d281eb74832db0c232e75/LICENSE
'MIT License. Copyright © 2024 MatheusNReis
    
    
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastColumn As Long
    
    ' Define o caminho do arquivo Excel a ser aberto
    Dim filePath As String
    filePath = "C:\Users\matheus.reis\Desktop\acompanhamento_fisico_mensal_concessionaria.xlsx" ' Substitua pelo caminho do seu arquivo

    ' Abre o arquivo Excel
    Set wb = Workbooks.Open(filePath)
    
    ' Define a planilha ativa do arquivo aberto
    Set ws = wb.Sheets(1) ' Altere para a planilha desejada se não for a primeira
    
    ' Encontra a última linha e a última coluna com dados na planilha
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    lastColumn = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    
    ' Copia todo o conteúdo da planilha
    ws.Range(ws.Cells(1, 1), ws.Cells(100, 2000)).Copy      'ATENÇÃO: observar as linhas de início ws.Cells(5, 5)
    
    ' Cola o conteúdo na célula A16 da pasta atual ativa
    ThisWorkbook.Sheets(1).Range("A16").PasteSpecial Paste:=xlPasteAll
    
    ' Fecha o arquivo Excel sem salvar alterações
    wb.Close False
    
    ' Limpa a área de transferência
    Application.CutCopyMode = False
End Sub

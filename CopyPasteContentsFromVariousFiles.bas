Sub CopiarConteudoDeVariosArquivos()
    
'Copy contents from different files

'Copia dados de diversos arquivos

'Created by Matheus Nunes Reis on 30/04/2024
'Copyright © 2024 Matheus Nunes Reis. All rights reserved.

'GitHub: MatheusNReis
'License: https://raw.githubusercontent.com/MatheusNReis/ANTT-GeneralApplications/3f27aa6f717d6861484d281eb74832db0c232e75/LICENSE
'MIT License. Copyright © 2024 MatheusNReis
    

    Dim wsDestino As Worksheet
    Dim lastRowDestino As Long
    Dim lastColumnDestino As Long
    Dim filePath As String
    Dim fileName As String
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim lastColumn As Long
    
    'Defina a planilha de destino onde deseja colar o conteúdo
    Set wsDestino = ThisWorkbook.Sheets(1) ' Altere conforme necessário
    
    'Encontre a última linha e a última coluna com dados na planilha de destino
    lastRowDestino = wsDestino.Cells(wsDestino.Rows.Count, "A").End(xlUp).Row
    lastColumnDestino = wsDestino.Cells(1, wsDestino.Columns.Count).End(xlToLeft).Column
    
    'Loop sobre os arquivos na pasta
    filePath = "Caminho\para\sua\pasta\" ' Substitua pelo caminho da sua pasta
    fileName = Dir(filePath & "*.xlsx")
    
    Do While fileName <> ""
        'Abre o arquivo Excel
        Set wb = Workbooks.Open(filePath & fileName)
        
        'Define a planilha ativa do arquivo aberto
        Set ws = wb.Sheets(1) ' Altere para a planilha desejada se não for a primeira
        
        'Encontra a última linha e a última coluna com dados na planilha
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        lastColumn = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        
        'Copia todo o conteúdo da planilha
        ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastColumn)).Copy
        
        'Cola o conteúdo na próxima linha vazia da planilha de destino
        wsDestino.Cells(lastRowDestino + 1, 1).PasteSpecial Paste:=xlPasteAll
        
        'Fecha o arquivo Excel sem salvar alterações
        wb.Close False
        
        'Atualiza a última linha na planilha de destino
        lastRowDestino = wsDestino.Cells(wsDestino.Rows.Count, "A").End(xlUp).Row
        
        'Limpa a área de transferência
        Application.CutCopyMode = False
        
        'Obtém o próximo arquivo na pasta
        fileName = Dir
    Loop
End Sub

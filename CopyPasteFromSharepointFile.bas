Sub CopiarDadosDePlanilhaParaArquivoConcessionaria()
    

'Copy and paste content from sharepoint to the current file using object excel method

'Copia e cola conteúdo de sharepoint para o arquivo atual usando o método de objeto excel

'Created by Matheus Nunes Reis on 12/05/2024
'Copyright © 2024 Matheus Nunes Reis. All rights reserved.

'GitHub: MatheusNReis
'License: https://raw.githubusercontent.com/MatheusNReis/ANTT-GeneralApplications/3f27aa6f717d6861484d281eb74832db0c232e75/LICENSE
'MIT License. Copyright © 2024 MatheusNReis
    

    Dim SharePointURL As String
    Dim SharePointFileName As String
    Dim objExcel As Object
    Dim objWorkbook As Object
    Dim objWorksheet As Object
    Dim DestWorkbook As Workbook
    Dim DestWorksheet As Worksheet
    Dim i As Integer
    Dim LastRow As Long
    
    'URL do SharePoint e nome do arquivo
    'SharePointURL = "https://anttgov.sharepoint.com/sites/LinKSharepoint"
    'SharePointFileName = "acompanhamento_obras_concessionaria"
    
    'Abre uma nova instância do Excel
    Set objExcel = CreateObject("Excel.Application")
    objExcel.Visible = True ' Torna o Excel visível
    
    'Abre a pasta de trabalho do SharePoint
    Set objWorkbook = objExcel.Workbooks.Open("https://anttgov.sharepoint.com/sites/LinKSharepoint.xlsx")
    
    'Define a primeira planilha da pasta de trabalho do Sharepoint
    Set objWorksheet = objWorkbook.Sheets(1)
    
    'Define a pasta de trabalho de destino e a planilha de destino
    Set DestWorkbook = Workbooks.Open("C:\Users\matheus.reis\Desktop\acompanhamento_obras_Concessionaria.xlsx")
    Set DestWorksheet = DestWorkbook.Sheets("Concessionaria")
    
    'Encontra a última linha com dados na planilha de destino
    LastRow = DestWorksheet.Cells(DestWorksheet.Rows.Count, "A").End(xlUp).Row
    
    'Copia os dados da primeira planilha para a planilha de destino
    For i = 1 To 15 ' Supondo que você queira copiar 10 linhas de dados
        ' Copia os dados da coluna A da primeira planilha para a coluna A da planilha de destino
        DestWorksheet.Cells(LastRow + i, 1).Value = objWorksheet.Cells(i, 1).Value               'arrumar para copiar a planilha toda
        ' Você pode expandir este loop para copiar mais colunas, se necessário
    Next i
    
    'Fecha a pasta de trabalho do SharePoint
    objWorkbook.Close SaveChanges:=False
    
    'Salva e fecha a pasta de trabalho de destino
    DestWorkbook.Save
    DestWorkbook.Close
    
    'Libera a memória
    Set DestWorksheet = Nothing
    Set DestWorkbook = Nothing
    Set objWorksheet = Nothing
    Set objWorkbook = Nothing
    objExcel.Quit
    Set objExcel = Nothing
    
End Sub

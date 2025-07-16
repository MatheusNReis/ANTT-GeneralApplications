Sub CombinarArquivosCSV()

'Combines .csv files from a folder into a single .csv.
'The data is presented with empty rows between entries, indicating different source files.

'Combina arquivos .csv contidos numa pasta em um único .csv. 
'Os dados são apresentados com linhas vazias entre eles indicando que são de arquivos de origem diferentes.
 
    Dim caminhoPasta As String
    Dim nomeArquivo As String
    Dim planilha As Worksheet
    Dim ultimaLinha As Long
    Dim dadosCSV As Variant
    Dim linhaCSV As Variant
    Dim i As Long, j As Long
    Dim fso As Object
    Dim arquivo As Object
    Dim texto As String
    
    ' Defina o caminho da pasta
    caminhoPasta = "C:\Users\matheus.reis\Desktop\Testecsv" ' Altere para o caminho da sua pasta
    
    ' Verifique se o caminho da pasta termina com uma barra invertida
    If Right(caminhoPasta, 1) <> "\" Then
        caminhoPasta = caminhoPasta & "\"
    End If
    
    ' Crie uma nova planilha para os dados combinados
    Set planilha = ThisWorkbook.Sheets.Add
    planilha.Name = "DadosCombinados"
    
    ' Inicialize a última linha
    ultimaLinha = 1
    
    ' Inicialize o FileSystemObject
    Set fso = CreateObject("Scripting.FileSystemObject")
    
    ' Loop através de todos os arquivos CSV na pasta
    nomeArquivo = Dir(caminhoPasta & "*.csv")
    If nomeArquivo = "" Then
        MsgBox "Nenhum arquivo CSV encontrado na pasta especificada."
        Exit Sub
    End If
    
    Do While nomeArquivo <> ""
        ' Abra o arquivo CSV
        Set arquivo = fso.OpenTextFile(caminhoPasta & nomeArquivo, 1)
        
        ' Leia o conteúdo do arquivo CSV
        texto = arquivo.ReadAll
        arquivo.Close
        
        ' Remova o BOM, se presente
        If Left(texto, 3) = "ï»¿" Then
            texto = Mid(texto, 4)
        End If
        
        ' Divida o conteúdo em linhas
        linhaCSV = Split(texto, vbCrLf)
        
        ' Loop através das linhas do arquivo CSV
        For i = 0 To UBound(linhaCSV)
            ' Divida cada linha em colunas
            dadosCSV = Split(linhaCSV(i), ";")
            
            ' Copie os dados para a planilha combinada
            For j = 0 To UBound(dadosCSV)
                planilha.Cells(ultimaLinha, j + 1).Value = dadosCSV(j)
            Next j
            ultimaLinha = ultimaLinha + 1
        Next i
        
        ' Obtenha o próximo arquivo CSV
        nomeArquivo = Dir
    Loop
    
    MsgBox "Combinação de arquivos CSV concluída!"
End Sub

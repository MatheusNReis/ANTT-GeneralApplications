Sub MesclarEmGrupos()


'Mesclar Células em Grupos.xlsm

'Merge cells in groups with number of lines defined by user and for specified column

'Mescla células em grupos de tamanho definido pelo usuário e para coluna especificada

'Created by Matheus Nunes Reis on 08/01/2025
'Copyright © 2025 Matheus Nunes Reis. All rights reserved.
    
'GitHub: MatheusNReis
'License:
'MIT License. Copyright © 2025 MatheusNReis


    Dim works As Worksheet
    Dim NomePlanilha As String
    Dim GrupoLinhas As Long
    Dim LastRowPlanWorks As Long
    Dim LinhaInicial As Long
    Dim Coluna As String
    
    
    NomePlanilha = ThisWorkbook.Sheets(1).Cells(3, "D").Value
    GrupoLinhas = ThisWorkbook.Sheets(1).Cells(4, "D").Value
    LinhaInicial = ThisWorkbook.Sheets(1).Cells(5, "D").Value
    Coluna = ThisWorkbook.Sheets(1).Cells(6, "D").Value
    
    
    
    If NomePlanilha = "" Then
        MsgBox "Informação 'Nome Planilha' não está preenchida."
        Exit Sub
    ElseIf GrupoLinhas = 0 Then
        MsgBox "Informação 'Mesclar linhas em grupos de' não está preenchida."
        Exit Sub
    ElseIf LinhaInicial = 0 Then
        MsgBox "Informação 'Linha inicial' não está preenchida."
        Exit Sub
    ElseIf Coluna = "" Then
        MsgBox "Informação 'Coluna' não está preenchida."
        Exit Sub
    End If
    
    
    Dim found As Boolean
    found = False
    For Each wb In Workbooks
        For Each ws In wb.Worksheets
            If ws.Name = NomePlanilha Then
            
                Dim resposta As VbMsgBoxResult
                resposta = MsgBox("'" & NomePlanilha & "' encontrado na planilha '" & wb.Name & "'", vbOKCancel + vbQuestion, "Confirmação de Planilha")
                If resposta = vbCancel Then
                    Exit Sub
                End If
                
                Set workb = wb
                Set works = wb.Sheets(NomePlanilha)
                found = True
                Exit For
                
            End If
        Next ws
        If found Then Exit For
    Next wb
        
    If Not found Then
        MsgBox "Planilha '" & NomePlanilha & "' não encontrada nas planilhas abertas."
        Exit Sub
    End If


    LastRowPlanWorks = works.Cells(Rows.Count, Coluna).End(xlUp).Row + GrupoLinhas



    Dim i As Long
    Dim LinhaFimMescla As Long
    
    For i = LinhaInicial To LastRowPlanWorks
        
        LinhaFimMescla = i + GrupoLinhas - 1 'i = linha inicial da mescla
        
         'Realiza mescla da linha i até LinhaFimMescla
         works.Range(Coluna & i, Coluna & LinhaFimMescla).Merge


        i = LinhaFimMescla 'O 'next i' ajusta i para a próxima LinhaInicial correta

    Next i
    
    
    MsgBox "Fim da mesclagem."
    
    
End Sub

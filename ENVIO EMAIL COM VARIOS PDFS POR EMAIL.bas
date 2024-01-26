Sub diretorio()
    Dim caminho As String
    Dim lin As Integer

    lin = Range("A7:A30").End(xlDown).Row

    'Laço para percorrer a coluna A começando da linha A7
    If Range("A7").Value = "" Then
        MsgBox "Insira um número da ficha para continuar."
        Exit Sub
    End If

    For linha = 7 To lin
        'Caminho = Variável com o endereço da pasta PDF, preenchida na célula B2, mas puxada já com regra de quebra de espaços da C3 _
        (bloqueada e oculta ao usuário)
        caminho = Cells(4, 3).Value & "\"
        'Procurar por todos os arquivos PDF na pasta com o padrão desejado
        Dim pdfFile As String
        pdfFile = Dir(caminho & Cells(linha, 1).Value & "*.pdf")

        'Verificar se é um arquivo PDF válido
        If InStr(1, pdfFile, Cells(linha, 1).Value, vbTextCompare) > 0 Then
            Cells(linha, "G").Value = "Ok"
            pdfFile = Dir
        
        ElseIf InStr(1, pdfFile, Cells(linha, 1).Value, vbTextCompare) = 0 Then
            Cells(linha, "G").Value = "Não tem"
            End If
           
            
        If Cells(linha, "A").Value = "" Then
            Cells(linha, "G").Value = ""
        End If
            
    Next

    MsgBox "PDFs verificados com sucesso!"
End Sub


Sub enviar_email()
    Dim lin As Integer
    Dim caminho As String
    Dim pdfFile As String
    Dim objeto_outlook As Object
    Dim Email As Object

    ' Verificar se a célula G7 está vazia
    If Range("G7").Value = "" Then
        MsgBox "Verifique os pdfs antes de continuar."
        Exit Sub

    ElseIf Err = -2147024894 Then
        MsgBox "Erro, reinicie o processo."
        Exit Sub

    Else
        MsgBox "Será preparado somente os e-mails com pdfs localizados!"
    End If

    lin = Range("A7:A30").End(xlDown).Row

    ' set porque esta recebendo um objeto/ pegar app outlook
    Set objeto_outlook = CreateObject("Outlook.Application")

    ' Laço para percorrer a linha A7 até limite de lin
    For linha = 7 To lin

        ' Verifica se há um valor na célula da coluna A
        If Cells(linha, 1).Value <> "" Then
            ' Criar um e-mail novo
            Set Email = objeto_outlook.createitem(0)

            ' caminho da pasta do pdf
            caminho = Cells(4, 3).Value & "\"

            ' Procurar por todos os arquivos PDF na pasta com o nome desejado
            pdfFile = Dir(caminho & Cells(linha, 1).Value & "*.pdf")

            ' Verificar se é um arquivo PDF válido
            If InStr(1, pdfFile, Cells(linha, 1).Value, vbTextCompare) > 0 Then
                Do While pdfFile <> ""
                    ' Anexar apenas os arquivos mencionados na coluna A
                    Email.Attachments.Add (caminho & pdfFile)
                    pdfFile = Dir
                Loop

                ' Restante do código permanece inalterado
                Email.display
                Email.To = Cells(2, 3).Value
                Email.cc = "LARI-LARN@grupofleury.com.br"
                Email.Subject = "Resultado Oncotype " & Cells(linha, 1).Value
                texto1 = "Olá Drs,<br><br> Segue resultado do exame oncotype <b> ficha: " & Cells(linha, 1).Value & " - " & Cells(linha, 2).Value & _
                         " </b> para liberação: <br><br> <b> Dados do médico: </b>"
                Email.htmlbody = texto1 & RangetoHTML(Range(Cells(6, 3), Cells(6, 6))) & RangetoHTML(Range(Cells(linha, 3), Cells(linha, 6))) & Email.htmlbody
                Call comparar_planilhas
                Cells(linha, "H").Value = "Pronto para Enviar"

            Else
                ' Tratamento caso não encontre PDFs correspondentes
                Cells(linha, "G").Value = "Não tem"
            End If
        End If
    Next

    ' Chamar a função que dá baixa nos e-mails com OK na coluna Email enviado
    Call SalvarPlanilha

    MsgBox "Programa encerrado!"
End Sub

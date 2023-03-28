Attribute VB_Name = "Módulo1"
Sub diretorio()
Dim caminho As String
Dim lin As Integer

'lin = limite da coluna A preenchida para o laço FOR

lin = Range("A7:A30").End(xlDown).Row

'Laço para percorrer a coluna A começando da linha A7

If Range("A7").Value = "" Then
    MsgBox "Insira um número da ficha para continuar."
    Exit Sub
    End If

        
For linha = 7 To lin
        
    'Caminho = Variavel com o endereço da pasta PDF, preenchida na celula B2, mas puxada já com regra de quebra de espaços da C3 _
    (bloqueada e oculta ao usuário)
            
    caminho = Cells(4, 3).Value & "\" & Cells(linha, 1).Value & ".pdf"
            
    'Se caminho vazio (sem pdf na pasta):
            
    If Dir(caminho) = "" Then
        If Cells(linha, "A").Value = "" Then
            Cells(linha, "H").Value = ""
        Else
            Cells(linha, "H").Value = "Não tem"
        End If
        
    'Se caminho com pdf:
    Else
       If Cells(linha, "A").Value = "" Then
            Cells(linha, "H").Value = ""
        Else
            Cells(linha, "H").Value = "Ok"
        End If

    End If

Next

MsgBox "Pdfs verificados com sucesso!"

End Sub

Sub enviar_email()
'variavel do primeiro laço
Dim lin As Integer
'On Error Resume Next
Dim caminho As String

'condição para ver se foi verificado os pdfs
If Cells(7, 7).Value = "" Then
    MsgBox " Verificar pdfs!"
    Exit Sub

'se tudo ok, ele segue com a elaboração do email
Else

'variavel "lin" que percorre a tabela
lin = Range("A7:A30").End(xlDown).Row

      
          
'set porque esta recebendo um objeto/ pegar app outlook
Set objeto_outlook = CreateObject("Outlook.Application")

'Criar um e-mail novo
Set Email = objeto_outlook.createitem(0)

'Laço para percorrer a linha A7 até limite de lin(que é a primeira coluna com fichas)

For linha = 7 To lin
        
    'Se houver algum pdf não localizado, ele não monta o email
    If Cells(linha, 7).Value = "Não tem" Then
        MsgBox "Envio cancelado, há pdfs não localizados"
        Exit Sub
    End If


            'condição para prencher coluna e-mail, se tiver vazio, coluna e-mail fica vazia, se não é preenchido o ok do email
    If Cells(linha, "A").Value = "" Then
        Cells(linha, "I").Value = ""
        
    Else
    'Caminho = Variavel com o endereço da pasta PDF, preenchida na celula B2, mas puxada já com regra de quebra de espaços da C3 _
    (bloqueada e oculta ao usuário)
    caminho = Cells(4, 3).Value & "\" & Cells(linha, 1).Value & ".pdf"
    
    'anexar pdfs localizados
    Email.Attachments.Add (caminho)
    
      'pegar dados da planilha da coluna 1 linha 6, até linha preenchida da coluna 7
    dados_cliente = RangetoHTML(Range(Cells(6, 1), Cells(linha, 7)))
    
    End If
    
            'condição para prencher coluna e-mail, se tiver vazio, coluna e-mail fica vazia, se não é preenchido o ok do email
       If Cells(linha, "A").Value = "" Then
            Cells(linha, "I").Value = ""
        Else
            Cells(linha, "I").Value = "Pronto para enviar"
        End If


Next

Email.display

Email.To = Cells(2, 3).Value
'Email.cc = "seumail@email.com.br"

Email.Subject = "Resultados exame xxx  - " & Cells(4, 4).Value

texto = "Olá a todos,<br><br> Caros, seguem resultados do exame xxx para a liberação: <br><br>"

Email.htmlbody = texto & dados_cliente & Email.htmlbody


MsgBox "Programa encerrado!"

End If


End Sub



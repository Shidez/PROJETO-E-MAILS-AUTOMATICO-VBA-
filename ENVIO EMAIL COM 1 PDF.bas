Attribute VB_Name = "Módulo1"
Sub diretorio()
Dim caminho As String
Dim lin As Integer

'lin = limite da coluna A preenchida para o laço FOR

'lin = Range("A7").CurrentRegion.Rows.Count '(NÃO PODE EM PLANILHA PROTEGIDA)
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
            Cells(linha, "G").Value = ""
        Else
            Cells(linha, "G").Value = "Não tem"
        End If
        
    'Se caminho com pdf:
    Else
       If Cells(linha, "A").Value = "" Then
            Cells(linha, "G").Value = ""
        Else
            Cells(linha, "G").Value = "Ok"
        End If

    End If

Next

MsgBox "Pdfs verificados com sucesso!"

End Sub


Sub enviar_email()
'variavel do primeiro laço
Dim lin As Integer
On Error Resume Next
Dim caminho As String

'Variavel para lembrar de verificar pdfs
   
If Range("G7").Value = "" Then
    MsgBox "Verifique os pdfs antes de continuar."
    Exit Sub
    
ElseIf Err = -2147024894 Then
    MsgBox "Erro, reinicie o processo."
    Exit Sub

Else
    MsgBox "Será preparado somente os e-mails com pdfs localizados!"
    
End If
    

'lin = limite da coluna A preenchida para o laço FOR
'lin = Range("A7").CurrentRegion.Rows.Count
lin = Range("A7:A30").End(xlDown).Row
    
'set porque esta recebendo um objeto/ pegar app outlook
Set objeto_outlook = CreateObject("Outlook.Application")


'Laço para percorrer a linha A7 até limite de lin
For linha = 7 To lin
    'Criar um e-mail novo
    Set Email = objeto_outlook.createitem(0)
    
    'Criar anexar o pdf primeiro para já dar o erro e bloquear o envio se não houver pdf
    
    Email.Attachments.Add (Cells(4, 3).Value & "\" & Cells(linha, 1).Value & ".pdf")
    
    'caminho da pasta do pdf
    
    caminho = Cells(4, 3).Value & "\" & Cells(linha, 1).Value & ".pdf"
    
    'Se caminho sem pdf na pasta:
    If Dir(caminho) = "" Then
        If Cells(linha, "A").Value = "" Then
            Cells(linha, "H").Value = ""
        Else
            Cells(linha, "H").Value = "Não tem"
        End If

    'Mensagem para avisar que acabou
    
    ElseIf Dir(caminho) = "" Then
        MsgBox "Programa encerrado!"
        Exit Sub
        
           
    'Se caminho com pdf:
      
    Else
    
        Email.display
        Email.To = Cells(2, 3).Value
        Email.cc = "seuemail@email.com.br"
        
        Email.Subject = "Resultado exame XXXX " & Cells(linha, 1).Value
                
        '& Cells(2, 2).Value & - incluir valor da celula no email
        
       
        texto1 = "Olá Drs,<br><br> Segue resultado do exame XXXX <b> ficha: " & Cells(linha, 1).Value & " - " & Cells(linha, 2).Value & _
         " </b> para liberação: <br><br> <b> Dados do médico: </b>"
                   
        Email.htmlbody = texto1 & RangetoHTML(Range(Cells(6, 3), Cells(6, 6))) & RangetoHTML(Range(Cells(linha, 3), Cells(linha, 6))) & Email.htmlbody

        'Email.send
        If Cells(linha, "A").Value = "" Then
            Cells(linha, "H").Value = ""
        Else
            Cells(linha, "H").Value = "Pronto para Enviar"
    
        End If

    End If
    
Next

MsgBox "Programa encerrado!"

End Sub





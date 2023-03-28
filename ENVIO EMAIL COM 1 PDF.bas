Attribute VB_Name = "M�dulo1"
Sub diretorio()
Dim caminho As String
Dim lin As Integer

'lin = limite da coluna A preenchida para o la�o FOR

'lin = Range("A7").CurrentRegion.Rows.Count '(N�O PODE EM PLANILHA PROTEGIDA)
lin = Range("A7:A30").End(xlDown).Row

'La�o para percorrer a coluna A come�ando da linha A7

If Range("A7").Value = "" Then
    MsgBox "Insira um n�mero da ficha para continuar."
    Exit Sub
    End If

        
For linha = 7 To lin
        
    'Caminho = Variavel com o endere�o da pasta PDF, preenchida na celula B2, mas puxada j� com regra de quebra de espa�os da C3 _
    (bloqueada e oculta ao usu�rio)
            
    caminho = Cells(4, 3).Value & "\" & Cells(linha, 1).Value & ".pdf"
            
    'Se caminho vazio (sem pdf na pasta):
            
    If Dir(caminho) = "" Then
        If Cells(linha, "A").Value = "" Then
            Cells(linha, "G").Value = ""
        Else
            Cells(linha, "G").Value = "N�o tem"
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
'variavel do primeiro la�o
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
    MsgBox "Ser� preparado somente os e-mails com pdfs localizados!"
    
End If
    

'lin = limite da coluna A preenchida para o la�o FOR
'lin = Range("A7").CurrentRegion.Rows.Count
lin = Range("A7:A30").End(xlDown).Row
    
'set porque esta recebendo um objeto/ pegar app outlook
Set objeto_outlook = CreateObject("Outlook.Application")


'La�o para percorrer a linha A7 at� limite de lin
For linha = 7 To lin
    'Criar um e-mail novo
    Set Email = objeto_outlook.createitem(0)
    
    'Criar anexar o pdf primeiro para j� dar o erro e bloquear o envio se n�o houver pdf
    
    Email.Attachments.Add (Cells(4, 3).Value & "\" & Cells(linha, 1).Value & ".pdf")
    
    'caminho da pasta do pdf
    
    caminho = Cells(4, 3).Value & "\" & Cells(linha, 1).Value & ".pdf"
    
    'Se caminho sem pdf na pasta:
    If Dir(caminho) = "" Then
        If Cells(linha, "A").Value = "" Then
            Cells(linha, "H").Value = ""
        Else
            Cells(linha, "H").Value = "N�o tem"
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
        
       
        texto1 = "Ol� Drs,<br><br> Segue resultado do exame XXXX <b> ficha: " & Cells(linha, 1).Value & " - " & Cells(linha, 2).Value & _
         " </b> para libera��o: <br><br> <b> Dados do m�dico: </b>"
                   
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





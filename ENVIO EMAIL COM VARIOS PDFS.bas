Attribute VB_Name = "M�dulo1"
Sub diretorio()
Dim caminho As String
Dim lin As Integer

'lin = limite da coluna A preenchida para o la�o FOR

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
            Cells(linha, "H").Value = ""
        Else
            Cells(linha, "H").Value = "N�o tem"
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
'variavel do primeiro la�o
Dim lin As Integer
'On Error Resume Next
Dim caminho As String

'condi��o para ver se foi verificado os pdfs
If Cells(7, 7).Value = "" Then
    MsgBox " Verificar pdfs!"
    Exit Sub

'se tudo ok, ele segue com a elabora��o do email
Else

'variavel "lin" que percorre a tabela
lin = Range("A7:A30").End(xlDown).Row

      
          
'set porque esta recebendo um objeto/ pegar app outlook
Set objeto_outlook = CreateObject("Outlook.Application")

'Criar um e-mail novo
Set Email = objeto_outlook.createitem(0)

'La�o para percorrer a linha A7 at� limite de lin(que � a primeira coluna com fichas)

For linha = 7 To lin
        
    'Se houver algum pdf n�o localizado, ele n�o monta o email
    If Cells(linha, 7).Value = "N�o tem" Then
        MsgBox "Envio cancelado, h� pdfs n�o localizados"
        Exit Sub
    End If


            'condi��o para prencher coluna e-mail, se tiver vazio, coluna e-mail fica vazia, se n�o � preenchido o ok do email
    If Cells(linha, "A").Value = "" Then
        Cells(linha, "I").Value = ""
        
    Else
    'Caminho = Variavel com o endere�o da pasta PDF, preenchida na celula B2, mas puxada j� com regra de quebra de espa�os da C3 _
    (bloqueada e oculta ao usu�rio)
    caminho = Cells(4, 3).Value & "\" & Cells(linha, 1).Value & ".pdf"
    
    'anexar pdfs localizados
    Email.Attachments.Add (caminho)
    
      'pegar dados da planilha da coluna 1 linha 6, at� linha preenchida da coluna 7
    dados_cliente = RangetoHTML(Range(Cells(6, 1), Cells(linha, 7)))
    
    End If
    
            'condi��o para prencher coluna e-mail, se tiver vazio, coluna e-mail fica vazia, se n�o � preenchido o ok do email
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

texto = "Ol� a todos,<br><br> Caros, seguem resultados do exame xxx para a libera��o: <br><br>"

Email.htmlbody = texto & dados_cliente & Email.htmlbody


MsgBox "Programa encerrado!"

End If


End Sub



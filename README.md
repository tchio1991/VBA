# VBA
Agenda completa em VBA integrada ao Excel
Private Sub CommandButton1_Click()
If Application.Visible = True Then
Application.Visible = False
Else
Application.Visible = True
End If
End Sub

Private Sub CommandButton2_Click()
If Application.Visible = True Then
End

Else
MsgBox "proibido sair do formulario"
End If
End Sub

Private Sub edit_Click()
Dim pergunta
Dim ul As Long
Dim myrange As Range
ul = Sheets("Banco").Range("A" & Rows.Count).End(xlUp).Row

If nome.Text = "" Then
        MsgBox "Contato invalido", vbCritical
        Exit Sub
    End If

pergunta = MsgBox("Atenção! isso irá alterar os valores atuais do contato, pelos valores dos campos desse formulário.." _
& vbNewLine & vbNewLine & "Deseja realmente fazer isso?", vbQuestion + vbYesNo)

If pergunta = vbYes Then
    
    For i = 1 To ul
        If Sheets("Banco").Range("A" & i).Text = nome Then
            With Worksheets("Banco").Range("A:A")
                Set myrange = Sheets("Banco").Range("A" & i)
                    myrange.Offset(0, 1).Value = endereco
                    myrange.Offset(0, 2).Value = estados
                    myrange.Offset(0, 4).Value = txt_fone
                If OptionButton1.Value = True Then
                myrange.Offset(0, 3).Value = "Masculino"
                Else
                myrange.Offset(0, 3).Value = "Feminino"
                End If
            End With
            MsgBox "Editado com Sucesso!", vbExclamation
            'Limpar as caixas de texto
             nome.Value = Empty
             endereco.Value = Empty
             estados.Value = Empty
             OptionButton1 = False
             OptionButton2 = False
             txt_fone.Value = Empty
             nome.SetFocus
            Exit Sub
        End If
Next i
    MsgBox "Contato Invalido", vbCritical
    Set myrange = Nothing
End If
End Sub

Private Sub endereco_Change()
If IsNumeric(endereco.Text) Then
endereco.Text = ""
endereco.SetFocus
End If
End Sub

Private Sub estados_Change()
If IsNumeric(estados.Text) Then
estados.Text = ""
estados.SetFocus
End If
End Sub

Private Sub excluir_Click()
If nome.Text = "" Then
MsgBox "Preencha o nome"
nome.SetFocus
Exit Sub
End If

If endereco.Text = "" Then
MsgBox "Preencha o endereço"
endereco.SetFocus
Exit Sub
End If

If estados.Text = "" Then
MsgBox "Preencha o estado"
estados.SetFocus
Exit Sub
End If

If OptionButton1 = False And OptionButton2 = False Then
MsgBox "Marque uma das alternativas"
nome.SetFocus
Exit Sub
End If

If txt_fone.Text = "" Then
MsgBox "Preencha o telefone"
txt_fone.SetFocus
Exit Sub
End If


'Declarar a variável Resp para receber uma resposta

Dim Resp As Integer


'Fazer a busca do registro digitado pelo usuário

With Worksheets("Banco").Range("A:A")

Set c = .Find(nome.Value, LookIn:=xlValues, LookAt:=xlWhole)


If Not c Is Nothing Then

    Resp = MsgBox("Tem certeza que deseja excluir o registro?", vbYesNo, "Confirmação")

    If Resp = vbYes Then

         c.Select

         Selection.EntireRow.Delete

         'Limpar as caixas de texto

         nome.Value = Empty

         endereco.Value = Empty

         estados.Value = Empty
         txt_fone.Value = Empty
    
         OptionButton1.Value = False

         OptionButton2.Value = False

         'Colocar o foco na primeira caixa de texto

         nome.SetFocus
MsgBox "Contato Excluido com Sucesso"
    Else

         MsgBox "O registro não será excluído!"

    End If

Else

     MsgBox "Cliente não encontrado!"
'Limpar as caixas de texto
nome.Value = Empty
endereco.Value = Empty
estados.Value = Empty
OptionButton1 = False
OptionButton2 = False
txt_fone.Value = Empty
nome.SetFocus
End If
End With
Exit Sub
End Sub

Private Sub fechar_Click()
If Application.Visible = False Then
Application.Quit
ActiveWorkbook.Save
Else
MsgBox "Feche o banco de dados"
Exit Sub
End If
MsgBox "Formulario fechado e salvo com sucesso", vbInformation, "Salvo"
End Sub

Private Sub gravar_Click()
If nome.Value = Empty Then
MsgBox "Preencha o nome", , "Nome"
nome.SetFocus
Exit Sub
End If

If endereco.Value = Empty Then
MsgBox "Preencha o endereço", , "Endereço"
endereco.SetFocus
Exit Sub
End If

If estados.Value = Empty Then
MsgBox "Preencha o estado", , "Estados"
estados.SetFocus
Exit Sub
End If

If OptionButton1 = False And OptionButton2 = False Then
MsgBox "Marque uma das opções: Masculino ou Feminino", , "Masculino ou Feminino"
nome.SetFocus
Exit Sub
End If

If txt_fone.Value = Empty Then
MsgBox "Preencha o telefone", , "Telefone"
txt_fone.SetFocus
Exit Sub
End If

'Ativar a primeira planilha
ThisWorkbook.Worksheets("Banco").Activate
'Selecionar a célula A2
Range("A2").Select
'Procurar a primeira célula vazia
Do
If Not (IsEmpty(ActiveCell)) Then
ActiveCell.Offset(1, 0).Select
End If
Loop Until IsEmpty(ActiveCell) = True

'Carregar os dados digitados nas caixas de texto para a planilha
ActiveCell.Value = nome.Value
ActiveCell.Offset(0, 1).Value = endereco.Value
ActiveCell.Offset(0, 2).Value = estados.Value
ActiveCell.Offset(0, 4).Value = txt_fone.Value

'Carregar o sexo do cliente dos botões de opção
If OptionButton1.Value = True Then
ActiveCell.Offset(0, 3).Value = "Masculino"
Else
ActiveCell.Offset(0, 3).Value = "Feminino"
End If

'Limpar as caixas de texto
nome.Value = Empty
endereco.Value = Empty
estados.Value = Empty
OptionButton1 = False
OptionButton2 = False
txt_fone.Value = Empty
nome.SetFocus

MsgBox "Registrado com sucesso", , "Sucesso"

End Sub

Private Sub nome_Change()
If IsNumeric(nome.Text) Then
nome.Text = ""
nome.SetFocus
End If
End Sub

Private Sub pesquisa_Click()
'Verificar se foi digitado um nome na primeira caixa de texto

If nome.Text = "" Then

     MsgBox "Digite um nome"

     nome.SetFocus

     Exit Sub
End If

With Worksheets("Banco").Range("A:A")

Set c = .Find(nome.Value, LookIn:=xlValues, LookAt:=xlPart)


If Not c Is Nothing Then

  c.Activate

  nome.Value = c.Value

  endereco.Value = c.Offset(0, 1).Value

  estados.Value = c.Offset(0, 2).Value

  txt_fone.Value = c.Offset(0, 4).Value


  'Carregando o botão de opção

  If c.Offset(0, 3) = "Masculino" Then

      OptionButton1.Value = True

  Else

       OptionButton2.Value = True

  End If

Else

    MsgBox "Cliente não localizado!"

End If

End With

End Sub

Private Sub txt_fone_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
 txt_fone.MaxLength = 14 '(45)3332-3333
 Select Case KeyAscii
      Case 8       'Aceita o BACK SPACE
      Case 14: SendKeys "{TAB}"    'Emula o TAB
      Case 48 To 57
         If txt_fone.SelStart = 0 Then txt_fone.SelText = "("
         If txt_fone.SelStart = 3 Then txt_fone.SelText = ")"
         If txt_fone.SelStart = 9 Then txt_fone.SelText = "-"
      Case Else: KeyAscii = 0     'Ignora os outros caracteres
   End Select
End Sub
Private Sub UserForm_Initialize()
Application.Visible = False
nome.SetFocus
nome.Font.Bold = True
nome.Font.Size = 10
endereco.Font.Bold = True
endereco.Font.Size = 10
estados.Font.Bold = True
estados.Font.Size = 10
txt_fone.Font.Bold = True
txt_fone.Font.Size = 10
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
 If CloseMode = vbFormControlMenu Then
        MsgBox "Feche pelo outro botão", vbCritical, "AVISO"
        Cancel = True
    End If
End Sub

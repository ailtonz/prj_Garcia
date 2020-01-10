Option Compare Database
Option Explicit

Private Sub cmdSalvar_Click()
On Error GoTo Err_cmdSalvar_Click
Dim strMsg As String
Dim strTitle As String
Dim Resposta As Variant


If Not ChecarCampos() Then

        strMsg = "Atenção: Existem campos obrigatórios não preenchidos! " & vbNewLine & _
                 "Deseja preenche-los?"
        strTitle = "Registro Inconsistente"
        
        Resposta = MsgBox(strMsg, vbExclamation + vbYesNo, strTitle)

        If Resposta = vbYes Then
            If Not TestaCampos() Then Exit Sub
        Else
            GoTo SALVAR
        End If

Else

SALVAR:
        'Salvar Registro
        DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70
'        'Gerar Historico
'        HistoricoDeUsuario Me.Caption, strUsuario, Me.Codigo.value & "|" & Me.Nome.value & "|" & Me.Apelido.value & "|" & Me.CNPJ.value & "|" & Me.IE.value & "|" & Me.Telefone.value & "|" & Me.Contato.value & "|" & Me.Email.value & "|" & Me.Endereco.value & "|" & Me.Bairro.value & "|" & Me.Cep.value & "|" & Me.Municipio.value & "|" & Me.Estado.value
        'Atualizar Pesquisa
        If EstaAberto("Pesquisar") Then Form_Pesquisar.lstCadastro.Requery
        'Fechar Formulário
        DoCmd.Close

End If
    
Exit_cmdSalvar_Click:
    Exit Sub

Err_cmdSalvar_Click:
    If Not (Err.Number = 2046 Or Err.Number = 0) Then MsgBox Err.Description
    DoCmd.Close
    Resume Exit_cmdSalvar_Click
End Sub

Private Sub cmdFechar_Click()
On Error GoTo Err_cmdFechar_Click

    DoCmd.DoMenuItem acFormBar, acEditMenu, acUndo, , acMenuVer70
    DoCmd.CancelEvent
    DoCmd.Close

Exit_cmdFechar_Click:
    Exit Sub

Err_cmdFechar_Click:
    If Not (Err.Number = 2046 Or Err.Number = 0) Then MsgBox Err.Description
    DoCmd.Close
    Resume Exit_cmdFechar_Click

End Sub

Private Sub Nome_Exit(Cancel As Integer)
    Me.Nome = UCase(Me.Nome)
End Sub

Private Sub Apelido_Exit(Cancel As Integer)
    Me.Apelido = UCase(Me.Apelido)
End Sub

Private Sub Contato_Exit(Cancel As Integer)
    Me.Contato = UCase(Me.Contato)
End Sub

Private Sub Email_Exit(Cancel As Integer)
    Me.Email = LCase(Me.Email)
End Sub

Private Sub Endereco_Exit(Cancel As Integer)
    Me.Endereco = UCase(Me.Endereco)
End Sub

Private Sub Bairro_Exit(Cancel As Integer)
    Me.Bairro = UCase(Me.Bairro)
End Sub

Private Sub Municipio_Exit(Cancel As Integer)
    Me.Municipio = UCase(Me.Municipio)
End Sub

Private Sub Estado_Exit(Cancel As Integer)
    Me.Estado = UCase(Me.Estado)
End Sub

Private Sub OBS_Exit(Cancel As Integer)
    Me.OBS = UCase(Me.OBS)
End Sub

Function TestaCampos() As Integer
Dim I As Integer
Dim strMsg As String
Dim strTitle As String

TestaCampos = True

For I = 0 To Me.Count - 1
    If Me(I).Tag = "-1" Then
        If IsNull(Me(I)) Then
            strMsg = "É obrigatório o preenchimento do campo '" & Me(I).Name & "'!"
            strTitle = "Registro Inconsistente"
    
            MsgBox strMsg, vbExclamation, strTitle
            Me(I).SetFocus
            TestaCampos = False
            Exit Function
            
        End If
    End If
Next I
    
End Function


Function ChecarCampos() As Boolean
Dim I As Integer
Dim strMsg As String
Dim strTitle As String

ChecarCampos = True

For I = 0 To Me.Count - 1
    If Me(I).Tag = "-1" Then
        If IsNull(Me(I)) Then
            ChecarCampos = False
            Exit Function
        End If
    End If
Next I
    
End Function


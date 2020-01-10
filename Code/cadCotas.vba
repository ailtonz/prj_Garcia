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
'        HistoricoDeUsuario Me.Caption, strUsuario, Me.codAterro.value & "|" & Me.dtCota.value & "|" & Me.Aterro.value & "|" & Me.QTD.value
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


Private Sub codAterro_Click()
    Me.Aterro = Me.codAterro.Column(1)
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


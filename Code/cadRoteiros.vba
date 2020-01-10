Option Compare Database
Option Explicit

Dim Viagem As Integer
Dim Ordem As Integer

Private Sub cmdRoteiro_Click()
    PrintPage 25, Me.codRoteiro
End Sub

Private Sub cmdSaldo_Click()
    Me.lstColoca.Requery
    Me.lstTroca.Requery
    Me.lstRetira.Requery
End Sub

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
'        HistoricoDeUsuario Me.Caption, strUsuario, Me.codRoteiro.value & "|" & Me.DTRoteiro & "|" & Me.Motorista & "|" & Me.Placa & "|" & Me.DTSaida
        'Atualizar Pesquisa
        If EstaAberto("Pesquisar") Then Form_Pesquisar.lstCadastro.Requery
        'Fechar Formulário
        DoCmd.Close
        'Pesquisar Roteiros
        Pesquisar 5

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
    Pesquisar 5
    
Exit_cmdFechar_Click:
    Exit Sub

Err_cmdFechar_Click:
    If Not (Err.Number = 2046 Or Err.Number = 0) Then MsgBox Err.Description
    DoCmd.Close
    Resume Exit_cmdFechar_Click

End Sub

Private Sub DTRoteiro_Exit(Cancel As Integer)
    Me.lstColoca.Requery
    Me.lstRetira.Requery
    Me.lstTroca.Requery
End Sub

Sub Form_Load()
'    Me.TimerInterval = 1000
End Sub

Private Sub Form_Open(Cancel As Integer)
    DoCmd.Maximize
    Viagem = 1
    Ordem = 1
    
    Me.lstColoca.Requery
    Me.lstTroca.Requery
    Me.lstRetira.Requery
    
End Sub

'Private Sub Form_Resize()
''Dim x
''
''x = RedimencionaControle(Me, [RoteirosItens])
'End Sub

'Sub Form_Timer()
'End Sub

Private Sub lstColoca_DblClick(Cancel As Integer)

Dim strSQL As String
Dim Obra As String
Dim Cliente As String
Dim codObra As String
Dim codCadastro As String
Dim Coloca As String
Dim dtColoca As String
Dim CTR As String

Dim Resposta As Boolean

Obra = Me.lstColoca.Column(0)
Cliente = Me.lstColoca.Column(1)
Coloca = Me.lstColoca.Column(2)
dtColoca = IIf((Me.lstColoca.Column(3) <> ""), Me.lstColoca.Column(3), Format(Now, "dd/mm/yy"))
codObra = Me.lstColoca.Column(4)
codCadastro = Me.lstColoca.Column(5)
CTR = Me.lstColoca.Column(6)

Resposta = VerificarCadastro(Me.codRoteiro, codObra, Viagem)

If Resposta Then
    AtualizarViagem "Coloca", codRoteiro, Viagem, codObra, Coloca, dtColoca
Else
    CadastrarViagem "Coloca", codRoteiro, Viagem, Ordem, Obra, Cliente, codObra, codCadastro, CTR, Coloca, dtColoca
End If

'strSQL = "INSERT INTO RoteirosItens ( codRoteiro, Viagem, Ordem, Obra, Cliente, codObra, codCadastro,CTR, C, DT_C ) " & _
'         "values (" & Me.codRoteiro & ",'" & Viagem & "','" & Ordem & "','" & Obra & "','" & Cliente & "'," & codObra & "," & codCadastro & ",'" & CTR & "'," & Coloca & ",'" & dtColoca & "')"

'ExecutarSQL strSQL

Me.lstColoca.Requery

Ordem = Ordem + 1

End Sub

Private Sub lstRetira_DblClick(Cancel As Integer)

Dim strSQL As String
Dim Obra As String
Dim Cliente As String
Dim codObra As String
Dim codCadastro As String
Dim Retira As String
Dim dtRetira As String
Dim CTR As String

Dim Resposta As Boolean

Obra = Me.lstRetira.Column(0)
Cliente = Me.lstRetira.Column(1)
Retira = Me.lstRetira.Column(2)
dtRetira = IIf((Me.lstRetira.Column(3) <> ""), Me.lstRetira.Column(3), Format(Now, "dd/mm/yy"))
codObra = Me.lstRetira.Column(4)
codCadastro = Me.lstRetira.Column(5)
CTR = Me.lstRetira.Column(6)

Resposta = VerificarCadastro(Me.codRoteiro, codObra, Viagem)

If Resposta Then
    AtualizarViagem "Retira", codRoteiro, Viagem, codObra, Retira, dtRetira
Else
    CadastrarViagem "Retira", codRoteiro, Viagem, Ordem, Obra, Cliente, codObra, codCadastro, CTR, Retira, dtRetira
End If

'strSQL = "INSERT INTO RoteirosItens ( codRoteiro, Viagem, Ordem, Obra, Cliente, codObra, codCadastro,CTR, R, DT_R ) " & _
'         "values (" & Me.codRoteiro & ",'" & Viagem & "','" & Ordem & "','" & Obra & "','" & Cliente & "'," & codObra & "," & codCadastro & ",'" & CTR & "'," & Retira & ",'" & dtRetira & "')"
'
'ExecutarSQL strSQL

Me.lstRetira.Requery

Ordem = Ordem + 1

End Sub

Private Sub lstTroca_DblClick(Cancel As Integer)

Dim strSQL As String
Dim Obra As String
Dim Cliente As String
Dim codObra As String
Dim codCadastro As String
Dim Troca As String
Dim dtTroca As String
Dim CTR As String

Dim Resposta As Boolean

Obra = Me.lstTroca.Column(0)
Cliente = Me.lstTroca.Column(1)
Troca = Me.lstTroca.Column(2)
dtTroca = IIf((Me.lstTroca.Column(3) <> ""), Me.lstTroca.Column(3), Format(Now, "dd/mm/yy"))
codObra = Me.lstTroca.Column(4)
codCadastro = Me.lstTroca.Column(5)
CTR = Me.lstTroca.Column(6)

Resposta = VerificarCadastro(Me.codRoteiro, codObra, Viagem)

If Resposta Then
    AtualizarViagem "Troca", codRoteiro, Viagem, codObra, Troca, dtTroca
Else
    CadastrarViagem "Troca", codRoteiro, Viagem, Ordem, Obra, Cliente, codObra, codCadastro, CTR, Troca, dtTroca
End If

'strSQL = "INSERT INTO RoteirosItens ( codRoteiro, Viagem, Ordem, Obra, Cliente, codObra, codCadastro,CTR, T, DT_T ) " & _
'         "values (" & Me.codRoteiro & ",'" & Viagem & "','" & Ordem & "','" & Obra & "','" & Cliente & "'," & codObra & "," & codCadastro & ",'" & CTR & "'," & Troca & ",'" & dtTroca & "')"
'
'ExecutarSQL strSQL

Me.lstTroca.Requery

Ordem = Ordem + 1

End Sub

Private Sub Motorista_Click()
    Me.codCadastro = Me.Motorista.Column(1)
    Me.RoteirosItens.Enabled = True
    Me.Recalc
End Sub
Private Sub cmdCTR_Click()
On Error GoTo Err_cmdCTR_Click

    Dim stDocName As String

    stDocName = "CTR"
    ImpressoraPadrao Categoria("ImpressoraPadrao")

    
    If Me.codRoteiro <> "" Then
        DoCmd.OpenReport stDocName, acPreview, , "codRoteiro = " & Me.codRoteiro
    End If

Exit_cmdCTR_Click:
    Exit Sub

Err_cmdCTR_Click:
    MsgBox Err.Description
    Resume Exit_cmdCTR_Click
    
End Sub
Private Sub cmdPedidos_Click()
On Error GoTo Err_cmdPedidos_Click

    Dim stDocName As String

    stDocName = "Pedidos"
'    ImpressoraPadrao Categoria("ImpressoraPadrao")
    
    If Me.codRoteiro <> "" Then
        DoCmd.OpenReport stDocName, acPreview, , "codRoteiro = " & Me.codRoteiro
    End If

Exit_cmdPedidos_Click:
    Exit Sub

Err_cmdPedidos_Click:
    MsgBox Err.Description
    Resume Exit_cmdPedidos_Click
    
End Sub

Private Sub optViagem_Click()

Select Case optViagem

    Case 1
        Viagem = 1
        Ordem = 1
    Case 2
        Viagem = 2
        Ordem = 1
    Case 3
        Viagem = 3
        Ordem = 1
    Case 4
        Viagem = 4
        Ordem = 1
    Case 5
        Viagem = 5
        Ordem = 1
        
End Select

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


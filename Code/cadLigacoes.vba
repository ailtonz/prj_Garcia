Option Compare Database
Option Explicit

Private Sub Cliente_Change()
    Me.lstPosicao.Requery
End Sub

Private Sub Cliente_Click()
    Me.codCadastro = Me.Cliente.Column(1)
    Me.lstPosicao.Requery
    Me.lstAntecipar.Requery
End Sub

Private Sub Cliente_Enter()
Dim strObras As String: strObras = "SELECT Obra, codObra, codCadastro, Razao, CTR FROM CadastrosObras WHERE (((CadastrosObras.codCadastro)=forms.cadLigacoes.codcadastro))"
    Me.Cliente.Requery
    Me.Obra.RowSource = strObras
    Me.Obra.Requery
End Sub

Private Sub Cliente_Exit(Cancel As Integer)
    Me.Obra.Requery
    Me.Contato.Requery
End Sub

Private Sub Cliente_NotInList(NewData As String, Response As Integer)
'Permite adicionar a editora à lista
Dim DB As DAO.Database
Dim rst As DAO.Recordset

On Error GoTo ErrHandler

'Pergunta se deseja acrescentar o novo item
If Confirmar("O Cliente: " & NewData & "  não faz parte da lista." & vbCrLf & "Deseja acrescentá-lo?") = True Then
    Set DB = CurrentDb()
    'Abre a tabela, adiciona o novo item e atualiza a combo
    Set rst = DB.OpenRecordset("Cadastros")
    With rst
        .AddNew
        '!codCadastro = NovoCodigo("cadGeral", "codCadastro")
        !Nome = NewData
        !TipoCadastro = "CLIENTE"
        .Update
        Response = acDataErrAdded
        .Close
    End With
        
    DoCmd.OpenForm "cadClientes", , , "Nome = '" & NewData & "'"
    
Else
    Response = acDataErrDisplay
End If

ExitHere:
Set rst = Nothing
Set DB = Nothing
Exit Sub

ErrHandler:
MsgBox Err.Description & vbCrLf & Err.Number & _
vbCrLf & Err.Source, , "EditoraID_fk_NotInList"
Resume ExitHere

End Sub

Private Sub Contato_Exit(Cancel As Integer)
    Me.Contato = UCase(Me.Contato)
End Sub

Private Sub lstAntecipar_DblClick(Cancel As Integer)
    
    Me.Obra = Me.lstAntecipar.Column(0)
    Me.R = Me.lstAntecipar.Column(2)
    Me.DT_R = CalcularVencimento(Day(Now) + 1, Month(Now), Year(Now))
    Me.codObra = Me.lstAntecipar.Column(9)

    strSQL = "INSERT INTO Ligacoes ( Obra, Cliente, codObra, codCadastro, R, DT_R, Data,codItem ) " & _
             "values ('" & Me.lstAntecipar.Column(0) & "','" & Me.Cliente & "'," & Me.codObra & "," & Me.codCadastro & "," & Me.lstAntecipar.Column(2) & ",'" & Me.lstAntecipar.Column(4) & "','" & Me.lstAntecipar.Column(4) & "'," & Me.lstAntecipar.Column(11) & ")"
    
    ExecutarSQL strSQL
    
    Me.lstAntecipar.Requery

End Sub

Private Sub lstPosicao_DblClick(Cancel As Integer)

    Me.Obra = Me.lstPosicao.Column(0)
    Me.R = Me.lstPosicao.Column(2)
    Me.DT_R = CalcularVencimento(Day(Now) + 1, Month(Now), Year(Now))
    Me.codObra = Me.lstPosicao.Column(9)

    strSQL = "INSERT INTO Ligacoes ( Obra, Cliente, codObra, codCadastro, R, DT_R, Data,codItem ) " & _
             "values ('" & Me.lstPosicao.Column(0) & "','" & Me.Cliente & "'," & Me.codObra & "," & Me.codCadastro & "," & Me.lstPosicao.Column(2) & ",'" & Me.lstPosicao.Column(4) & "','" & Me.lstPosicao.Column(4) & "'," & Me.lstPosicao.Column(11) & ")"
    
    ExecutarSQL strSQL
    
    Me.lstPosicao.Requery
    

End Sub

Private Sub Obra_Click()
    Me.codObra = Me.Obra.Column(1)
        
    If IsNull(Me.Cliente.value) Then
        Me.codCadastro = Me.Obra.Column(2)
        Me.Cliente = Me.Obra.Column(3)
        Me.CTR = Me.Obra.Column(4)
    End If
End Sub

Private Sub Obra_Enter()
Dim strObras As String: strObras = "SELECT Obra, codObra, codCadastro, Razao, CTR FROM CadastrosObras"

If IsNull(Me.Cliente.value) Then
    Me.Obra.RowSource = strObras
    Me.Obra.Requery
End If

End Sub

Private Sub Obra_Exit(Cancel As Integer)
    Me.Obra = UCase(Me.Obra)
End Sub

Private Sub OBS_Exit(Cancel As Integer)
    Me.OBS = UCase(Me.OBS)
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
        'Cadastrar data da ligação conforme o calendário
        Me.Data = Form_Calendario.Cal.value
        'Salvar Registro
        DoCmd.DoMenuItem acFormBar, acRecordsMenu, acSaveRecord, , acMenuVer70
'        'Gerar Historico
'        HistoricoDeUsuario Me.Caption, strUsuario, Me.codLigacao.value & "|" & Me.Cliente.value & "|" & Me.Obra.value & "|" & Me.Contato.value & "|" & Me.C.value & "|" & Me.DT_C.value & "|" & Me.R.value & "|" & Me.DT_R.value & "|" & Me.T.value & "|" & Me.T.value & "|" & Me.DT_T.value & "|" & Me.OBS.value & "|" & Me.Propaganda.value
        'Atualizar Listas do calendário
        If EstaAberto("Calendario") Then
            Form_Calendario.lstLigacoes.Requery
            Form_Calendario.lstAtrasos.Requery
        End If
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

Private Sub cmdNovoCliente_Click()
On Error GoTo Err_cmdNovoCliente_Click

    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "CadastroCliente"
    DoCmd.OpenForm stDocName, , , stLinkCriteria, acFormAdd

Exit_cmdNovoCliente_Click:
    Exit Sub

Err_cmdNovoCliente_Click:
    MsgBox Err.Description
    Resume Exit_cmdNovoCliente_Click
    
End Sub
Private Sub C_Exit(Cancel As Integer)
    If Me.C <> 0 Then Me.DT_C = CalcularVencimento(Day(Now) + 1, Month(Now), Year(Now))
End Sub

Private Sub R_Exit(Cancel As Integer)
    If Me.R <> 0 Then Me.DT_R = CalcularVencimento(Day(Now) + 1, Month(Now), Year(Now))
End Sub

Private Sub T_Exit(Cancel As Integer)
    If Me.T <> 0 Then Me.DT_T = CalcularVencimento(Day(Now) + 1, Month(Now), Year(Now))
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


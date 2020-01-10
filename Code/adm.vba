Option Compare Database
Option Explicit

Public strSQL As String
Public strTabela As String

Public codGrupo As Integer
Public codUsuario As Integer
Public strUsuario As String
Public codCadastro As String
Public strOBS As String
Public strOperacao As String
Public UsuarioOK As Boolean

Public Function AbrirLogin(Tabela As String, Optional Operacao As String)

strTabela = Tabela
strOperacao = Operacao
DoCmd.OpenForm "Login"

End Function

Public Function DescreverSolicitacao(codSolicitacao As Integer) As String
'Objetivo: Validar a senha do usuário.

Dim rSolicitacao As DAO.Recordset

Set rSolicitacao = CurrentDb.OpenRecordset("Select * from Formularios where codFormulario = " & codSolicitacao)

If Not rSolicitacao.EOF Then
    DescreverSolicitacao = rSolicitacao.Fields("NomeDoFormulario")
Else
    DescreverSolicitacao = ""
End If

rSolicitacao.Close

Set rSolicitacao = Nothing

End Function



Public Function HistoricoDeUsuario(strCadastro As String, strUsuario As String, strOBS As String)
'Objetivo: Historico de uso do sistema pelo usuário.

ExecutarSQL "insert into admGruposUsuariosHistorico (Cadastro,Usuario,OBS) Values ('" & strCadastro & "','" & strUsuario & "','" & strOBS & "')"

End Function


Public Function ValorCal()
On Error GoTo ValorCal_Err
   
      ' Testa se o form está aberto e em modo formulário
   If EstaAberto("Calendário") And IsFormView(Forms!Calendário) Then
      ' Captura o valor atual do calendário
      ValorCal = Forms!Calendário!Cal.value
   Else
      ValorCal = Now
   End If
   
   
ValorCal_Fim:
   Exit Function
ValorCal_Err:
   MsgBox Err.Description
   Resume ValorCal_Fim:
End Function


Public Function VerificarCadastro(codRoteiro, codObra, Viagem) As Boolean
Dim rDados As DAO.Recordset

Set rDados = CurrentDb.OpenRecordset("Select * from RoteirosItens where codRoteiro = " & codRoteiro & " and codObra = " & codObra & " and Viagem = " & Viagem & "")

If rDados.EOF Then
    VerificarCadastro = False
Else
    VerificarCadastro = True
End If

rDados.Close

Set rDados = Nothing

End Function


Public Function CadastrarViagem(TipoCadastro As String, codRoteiro, Viagem, Ordem, Obra, Cliente, codObra, codCadastro, CTR, Trabalho, dtTrabalho)

Select Case TipoCadastro

    Case "Coloca"
    
        strSQL = "INSERT INTO RoteirosItens ( codRoteiro, Viagem, Ordem, Obra, Cliente, codObra, codCadastro,CTR, C, DT_C ) " & _
                 "values (" & codRoteiro & ",'" & Viagem & "','" & Ordem & "','" & Obra & "','" & Cliente & "'," & codObra & "," & codCadastro & ",'" & CTR & "'," & Trabalho & ",'" & dtTrabalho & "')"

    Case "Retira"
    
        strSQL = "INSERT INTO RoteirosItens ( codRoteiro, Viagem, Ordem, Obra, Cliente, codObra, codCadastro,CTR, R, DT_R ) " & _
                 "values (" & codRoteiro & ",'" & Viagem & "','" & Ordem & "','" & Obra & "','" & Cliente & "'," & codObra & "," & codCadastro & ",'" & CTR & "'," & Trabalho & ",'" & dtTrabalho & "')"
    
    Case "Troca"
    
        strSQL = "INSERT INTO RoteirosItens ( codRoteiro, Viagem, Ordem, Obra, Cliente, codObra, codCadastro,CTR, T, DT_T ) " & _
                 "values (" & codRoteiro & ",'" & Viagem & "','" & Ordem & "','" & Obra & "','" & Cliente & "'," & codObra & "," & codCadastro & ",'" & CTR & "'," & Trabalho & ",'" & dtTrabalho & "')"

End Select

ExecutarSQL strSQL

End Function


Public Function AtualizarViagem(TipoCadastro As String, codRoteiro, Viagem, codObra, Trabalho, dtTrabalho)
Dim strSQL As String

Select Case TipoCadastro

    Case "Coloca"

        strSQL = "UPDATE RoteirosItens SET RoteirosItens.C = " & Trabalho & ", RoteirosItens.DT_C = '" & dtTrabalho & "' " & _
                 " WHERE (((RoteirosItens.codRoteiro)=" & codRoteiro & ") AND ((RoteirosItens.Viagem)=" & Viagem & ") AND ((RoteirosItens.codObra)=" & codObra & "))"
    
    Case "Retira"

        strSQL = "UPDATE RoteirosItens SET RoteirosItens.R = " & Trabalho & ", RoteirosItens.DT_R = '" & dtTrabalho & "' " & _
                 " WHERE (((RoteirosItens.codRoteiro)=" & codRoteiro & ") AND ((RoteirosItens.Viagem)=" & Viagem & ") AND ((RoteirosItens.codObra)=" & codObra & "))"
    
    Case "Troca"

        strSQL = "UPDATE RoteirosItens SET RoteirosItens.T = " & Trabalho & ", RoteirosItens.DT_T = '" & dtTrabalho & "' " & _
                 " WHERE (((RoteirosItens.codRoteiro)=" & codRoteiro & ") AND ((RoteirosItens.Viagem)=" & Viagem & ") AND ((RoteirosItens.codObra)=" & codObra & "))"
End Select

ExecutarSQL strSQL

End Function




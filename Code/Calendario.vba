'Option Compare Database
Option Explicit

Dim blnInicio As Boolean   ' indica inicialização do form
Public StringF2 As String

Private Sub Cal_AfterUpdate()
  ' Mudou a data, atualiza a agenda
  AtualizaAgenda
  Me!lstAtrasos.Requery
  
End Sub

Private Sub Cal_NewMonth()
   ' Calendario não deve aceitar datas nulas (sem data)
   Cal.ValueIsNull = False
   
   ' Põe o foco na caixa de texto DataLonga.
   ' Veja o comentário no procedimento
   ' DataAuxiliar_GotFocus.
   If Not blnInicio Then
      Me!DataAuxiliar.SetFocus
   End If
End Sub

Private Sub Cal_NewYear()
   ' Calendario não deve aceitar datas nulas (sem data)
   Cal.ValueIsNull = False
   
   ' Põe o foco na caixa de texto DataAuxiliar.
   ' Veja o comentário no procedimento
   ' DataAuxiliar_GotFocus.
   
   ' O If evita um erro quando o form está sendo
   ' carregado. Nesse momento, o controle ainda
   ' não pode receber o foco. A variável blnInicio
   ' indica que é o momento de abertura do form.
   If Not blnInicio Then
      Me!DataAuxiliar.SetFocus
   End If
End Sub

Private Sub AtualizaAgenda()
On Error GoTo Atualiza_Err

  ' Atualiza cx. texto DataLonga
  Me!DataAuxiliar.Requery
  ' Atualiza data na Agenda
  Me!lstLigacoes.Requery
  Me!lstAtrasos.Requery

Atualiza_Fim:
    Exit Sub
Atualiza_Err:
    MsgBox Err.Description
    Resume Atualiza_Fim
End Sub

Private Sub cmdHoje_Click()
On Error GoTo Err_cmdHoje_Click
Dim sqlLigacoes As String: sqlLigacoes = "SELECT DISTINCTROW Ligacoes.codLigacao, Ligacoes.Data, Ligacoes.Obra, Ligacoes.Cliente, Ligacoes.C, Ligacoes.R, Ligacoes.T, Ligacoes.DTLigacao FROM Ligacoes WHERE (((Ligacoes.Data)=Forms!Calendario!Cal.Value)) ORDER BY Ligacoes.DTLigacao DESC"
    
    ' Calendario: hoje
    Cal.Today
    ' Agenda: hoje
    AtualizaAgenda
    
    Me.lstLigacoes.RowSource = sqlLigacoes
    Me.lstLigacoes.ColumnWidths = "0cm;0cm;7cm;7cm;1cm;1cm;1cm"
    Me.lstLigacoes.Requery
    
Exit_cmdHoje_Click:
    Exit Sub
Err_cmdHoje_Click:
    MsgBox Err.Description
    Resume Exit_cmdHoje_Click
End Sub

Private Sub cmdImprimir_Click()
On Error GoTo cmdImprimir_Err

   ' Abre o form Imprimir como cx. de diálogo
   DoCmd.OpenForm "Imprimir", , , , , acDialog
   
cmdImprimir_Fim:
    Exit Sub
cmdImprimir_Err:
    MsgBox Err.Description
    Resume cmdImprimir_Fim
End Sub

Private Sub DataAuxiliar_GotFocus()
   ' Ao receber o foco, DataAuxiliar provoca a atualização
   ' do subform Agenda. Esta solução foi adotada porque o
   ' o objeto Calendario trava as caixas de combinação de
   ' mês e ano em janeiro e em 1900 se o método Agenda.Requery
   ' for chamado nos eventos Cal_NewMonth e Cal_NewYear,
   ' associados à escolha de mês ou ano nessas caixas.
   ' Esse método poderia ser aplicado à caixa DataLonga, mas
   ' para isso ela precisaria ficar como caixa de texto ativa,
   ' o que não faz sentido. Por isso criou-se o controle
   ' adicional DataAuxiliar, que fica ativo, mas tem tamanho
   ' bastante reduzido. Obs: DataAuxiliar não pode ser invisível,
   ' porque assim não tem condições de receber o foco.
   Me!lstLigacoes.Requery
   Me!lstAtrasos.Requery
End Sub

Private Sub Form_Load()
Dim blRet As Boolean



    ' Since we are not passing a filename of a Bitmap file
    ' the standard Window File Dialog will popup.
'    blRet = SetMDIBackGroundImage()
    
    blnInicio = True
    ' Calendario: hoje
    Cal.Today
    Cal.ValueIsNull = False
    blnInicio = False
    
    strTabela = "4"
    StringF2 = ""
'    Filtro strTabela
    Me.KeyPreview = True
    Me.lstLigacoes.SetFocus
    Me.lstLigacoes.Selected(1) = True
'    DoCmd.Maximize

    SetApplicationTitle Left(CurrentMDB, Len(CurrentMDB) - 4) '& " ( " & CaminhoDoBanco & " )"
    
    SysCmd acSysCmdSetStatus, "( " & CaminhoDoBanco & " )"


End Sub

Private Sub cmdFiltrar_Click()

    Dim txtFiltro As String
    txtFiltro = InputBox("Digite uma palavra para fazer o filtro:", "Filtro", StringF2, 0, 0)
    StringF2 = txtFiltro
    Filtro "4", txtFiltro
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
    
        Case vbKeyInsert
        
           cmdNovo_Click
           
        Case vbKeyReturn
        
            cmdAlterar_Click
            
        Case vbKeyDelete
           
            cmdExcluir_Click
        
        Case vbKeyF2
        
            cmdFiltrar_Click
            
    End Select
End Sub

Private Sub cmdNovo_Click()

    Manipulacao "4", "Novo"
'    AbrirLogin "4", "Novo"
    
End Sub

Private Sub cmdAlterar_Click()

    Manipulacao "4", "Alterar"
'    AbrirLogin "4", "Alterar"
    
End Sub

Private Sub cmdExcluir_Click()

    Manipulacao "4", "Excluir"
'    AbrirLogin "4", "Excluir"
    
End Sub

Private Sub lstAtrasos_DblClick(Cancel As Integer)
Dim strSQL As String

Dim Obra As String
Dim Cliente As String
Dim codObra As String
Dim codCadastro As String
Dim Retira As String
Dim dtRetira As String
Dim strData As String
Dim codItem As String
Dim CTR As String

Obra = Me.lstAtrasos.Column(0)
Cliente = Me.lstAtrasos.Column(1)
Retira = Me.lstAtrasos.Column(2)
codCadastro = Me.lstAtrasos.Column(8)
codObra = Me.lstAtrasos.Column(9)
CTR = Me.lstAtrasos.Column(10)
codItem = Me.lstAtrasos.Column(11)

dtRetira = CalcularVencimento(Day(Now) + 1, Month(Now), Year(Now))
strData = Format(Now, "dd/mm/yyyy")

strSQL = "INSERT INTO Ligacoes ( Obra, Cliente, codObra, codCadastro, R, DT_R, Data, codItem, CTR ) " & _
         "values ('" & Obra & "','" & Cliente & "'," & codObra & "," & codCadastro & "," & Retira & ",'" & dtRetira & "','" & strData & "'," & codItem & "," & CTR & ")"

ExecutarSQL strSQL

Me.lstAtrasos.Requery
Me.lstLigacoes.Requery

End Sub

Private Sub lstLigacoes_DblClick(Cancel As Integer)

    cmdAlterar_Click
    
End Sub

Private Sub cmdFechar_Click()
On Error GoTo Err_cmdFechar_Click

    DoCmd.Close

Exit_cmdFechar_Click:
    Exit Sub

Err_cmdFechar_Click:
    MsgBox Err.Description
    Resume Exit_cmdFechar_Click
    
End Sub

Public Sub Manipulacao(Tabela As String, Operacao As String)

If IsNull(Form_Calendario.lstLigacoes.value) And Operacao <> "Novo" Then
   Exit Sub
End If

Dim rstFormularios As DAO.Recordset

Set rstFormularios = _
    CurrentDb.OpenRecordset("Select * from Formularios where codFormulario = " & Tabela & "")

Select Case Operacao

 Case "Novo"
        
    DoCmd.OpenForm rstFormularios.Fields("NomeDoFormulario"), , , , acFormAdd
    
 Case "Alterar"

    DoCmd.OpenForm rstFormularios.Fields("NomeDoFormulario"), , , rstFormularios.Fields("Identificacao") & " = " & Form_Calendario.lstLigacoes.value

 Case "Excluir"

    If MsgBox("ATENÇÃO: Você deseja realmente excluir este registro ???", vbQuestion + vbOKCancel) = vbOK Then
       DoCmd.SetWarnings False
       DoCmd.RunSQL ("Delete from Ligacoes where codLigacao = " & Form_Calendario.lstLigacoes.value)
       DoCmd.SetWarnings True
    End If
    
    Me.lstAtrasos.Requery

End Select

Form_Calendario.lstLigacoes.Requery

Saida:

End Sub

Private Function Filtro(strTabela As String, Optional Procurar As String)

Dim rstFormularios As DAO.Recordset
Dim rstForm_Campos As DAO.Recordset
Dim rstForm_TabRelacionada As DAO.Recordset
Dim rstResultado As DAO.Recordset

Dim Sql As String
Dim SqlAux As String
Dim Contagem As Integer
Dim a, b, C As Integer
Dim Colunas As Integer

Dim Procuras(30) As String
Dim ProcurasAux As Integer
ProcurasAux = 1

For b = 1 To Len(Procurar)
   If Mid(Procurar, b, 1) = "+" Then
      ProcurasAux = ProcurasAux + 1
   Else
      Procuras(ProcurasAux) = Procuras(ProcurasAux) + Mid(Procurar, b, 1)
   End If
Next b

Set rstFormularios = _
    CurrentDb.OpenRecordset("Select * from Formularios " & _
                            " where codFormulario = " & _
                            strTabela & "")

Set rstForm_Campos = _
    CurrentDb.OpenRecordset("Select * from Formularios_Campos " & _
                            " where codFormulario = " & _
                            strTabela)

Set rstForm_TabRelacionada = _
    CurrentDb.OpenRecordset("Select * from Formularios_TabelaRelacionada " & _
                            " where codFormulario = " & _
                            strTabela)
Sql = "Select "

While Not rstForm_Campos.EOF
    If rstForm_Campos.Fields("Pesquisa") = True Then
        Sql = Sql & IIf(IsNull(rstForm_Campos.Fields("Nome")), _
                      rstForm_Campos.Fields("Campo"), _
                      rstForm_Campos.Fields("Campo") & _
                      " AS " & rstForm_Campos.Fields("Nome")) & ", "
    End If

    rstForm_Campos.MoveNext
Wend

Sql = Left(Sql, Len(Sql) - 2) & " "

Sql = Sql & " from "

If Not rstForm_TabRelacionada.EOF Then

    SqlAux = ""
    Contagem = 1
    rstForm_TabRelacionada.MoveFirst

    While Not rstForm_TabRelacionada.EOF

      SqlAux = "(" & SqlAux & IIf(Contagem <> 1, "", rstFormularios.Fields("TabelaPrincipal")) & " Left Join " & _
               rstForm_TabRelacionada.Fields("TabelaRelacionada") & " ON " & _
               rstFormularios.Fields("TabelaPrincipal") & "." & rstForm_TabRelacionada.Fields("CampoChave_Pai") & " = " & _
               rstForm_TabRelacionada.Fields("TabelaRelacionada") & "." & rstForm_TabRelacionada.Fields("CampoChave_Filho") & ")"

      rstForm_TabRelacionada.MoveNext
      Contagem = Contagem + 1

    Wend

    If SqlAux <> "" Then
       Sql = Sql & SqlAux
    End If

End If

If SqlAux = "" Then
   Sql = Sql & "" & rstFormularios.Fields("TabelaPrincipal") & " Where ( "
'Else
'   Sql = Sql & " Where ("
End If

rstForm_Campos.MoveFirst

For C = 1 To ProcurasAux

   rstForm_Campos.MoveFirst
   Sql = Sql & " ( "
   While Not rstForm_Campos.EOF
     If rstForm_Campos.Fields("Filtro") = True Then
        Sql = Sql & rstForm_Campos.Fields("Campo") & " Like '*" _
                  & LCase(Trim(Procuras(C))) & "*' OR "
     End If
     rstForm_Campos.MoveNext
   Wend
   Sql = Left(Sql, Len(Sql) - 3) & ") "
   If C <> ProcurasAux Then
      Sql = Sql + " And "
   End If

Next C

Sql = Sql + " ) "

Sql = Sql & "Order By "

rstForm_Campos.MoveFirst

While Not rstForm_Campos.EOF

  If rstForm_Campos.Fields("Ordem") <> "" Then
     Sql = Sql & rstForm_Campos.Fields("Campo") _
               & " " & rstForm_Campos.Fields("Ordem") & ", "
  End If

  rstForm_Campos.MoveNext

Wend

Sql = Left(Sql, Len(Sql) - 2) & " "

Sql = Sql & ";"

Me.lstLigacoes.RowSource = Sql
Me.lstLigacoes.ColumnHeads = True
Me.lstLigacoes.ColumnCount = rstForm_Campos.RecordCount
Me.Caption = rstFormularios.Fields("TituloDoFormulario")

Dim strTamanho As String

rstForm_Campos.MoveFirst
While Not rstForm_Campos.EOF
  If Not IsNull(rstForm_Campos.Fields("Tamanho")) Then
     strTamanho = strTamanho & str(rstForm_Campos.Fields("Tamanho")) & "cm;"
  End If
  rstForm_Campos.MoveNext
Wend

Me.lstLigacoes.ColumnWidths = strTamanho

'If IsNull(rstFormularios.Fields("campodesoma")) = True Then
'
'   Me.lblSoma.Caption = "Qtd: " & Me.lstLigacoes.ListCount - 1
'
'Else
'
'   Dim Soma
'   Set rstResultado = CurrentDb.OpenRecordset(Sql)
'   Do While Not rstResultado.EOF
'      Soma = Soma + rstResultado.Fields(rstFormularios.Fields("campodesoma"))
'      rstResultado.MoveNext
'   Loop
'   Me.lblSoma.Caption = "Qtd: " & Me.lstLigacoes.ListCount - 1 & "   Soma: " & FormatNumber(Soma, 2)
'   rstResultado.Close
'
'End If


rstFormularios.Close
rstForm_Campos.Close
rstForm_TabRelacionada.Close



End Function

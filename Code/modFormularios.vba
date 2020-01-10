Option Compare Database

Public Function Pesquisar(Tabela As String)
                                   
On Error GoTo Err_Pesquisar
  
    Dim stDocName As String
    Dim stLinkCriteria As String

    stDocName = "Pesquisar"
    strTabela = Tabela
       
    DoCmd.OpenForm stDocName, , , stLinkCriteria
        
Exit_Pesquisar:
    Exit Function

Err_Pesquisar:
    MsgBox Err.Description
    Resume Exit_Pesquisar
    
End Function

Public Function RedimencionaControle(frm As Form, ctl As Control)

Dim intAjuste As Integer
On Error Resume Next

intAjuste = frm.Section(acHeader).Height * frm.Section(acHeader).Visible

intAjuste = intAjuste + frm.Section(acFooter).Height * frm.Section(acFooter).Visible

On Error GoTo 0

intAjuste = Abs(intAjuste) + ctl.top

If intAjuste < frm.InsideHeight Then
    ctl.Height = frm.InsideHeight - intAjuste
'    ctl.Width = frm.InsideHeight + (intAjuste + intAjuste)
End If

End Function

Public Function EstaAberto(strName As String) As Boolean
On Error GoTo EstaAberto_Err
' Testa se o formulário está aberto

   Dim obj As AccessObject, dbs As Object
   Set dbs = Application.CurrentProject
   ' Procurar objetos AccessObject abertos na coleção AllForms.
   
   EstaAberto = False
   For Each obj In dbs.AllForms
        If obj.IsLoaded = True And obj.Name = strName Then
            ' Imprimir nome do obj.
            EstaAberto = True
            Exit For
        End If
   Next obj
    
EstaAberto_Fim:
  Exit Function
EstaAberto_Err:
  Resume EstaAberto_Fim
End Function

Public Function IsFormView(frm As Form) As Boolean
On Error GoTo IsFormView_Err
' Testa se o formulário está aberto em
' modo formulário (form view)

 IsFormView = False
 If frm.CurrentView = 1 Then
    IsFormView = True
 End If

IsFormView_Fim:
  Exit Function
IsFormView_Err:
  Resume IsFormView_Fim
End Function

Public Function AbrirArquivo(sTitulo As String, sDecricao As String, sTipo As String, SelecaoMultipla As Boolean) As String
Dim fd As Office.FileDialog

'Diálogo de selecionar arquivo - Office
Set fd = Application.FileDialog(msoFileDialogFilePicker)

'Título
fd.TITLE = sTitulo

'Filtros e descrição dos mesmos
fd.Filters.Add sDecricao, sTipo

'Premissões de selação
fd.AllowMultiSelect = SelecaoMultipla

If fd.Show = -1 Then
    AbrirArquivo = fd.SelectedItems(1)
End If

End Function

Public Function Confirmar(sMensagem As String) As _
Boolean
'Faz uma pergunta ao usuário e retorma True se a
'resposta for SIM, e false se a resposta for NÃO
Dim intResp As Integer

intResp = MsgBox(sMensagem, vbYesNo + vbQuestion, _
"Confirmação")

If intResp = vbYes Then
    Confirmar = True
Else
    Confirmar = False
End If
End Function




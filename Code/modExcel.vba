Option Compare Database


Public Function PrintPage(codRelatorio As Integer, Pagina As Integer)
Dim rRelatorio As DAO.Recordset
Dim sqlRelatorio: sqlRelatorio = "SELECT Modelo FROM admRelatorios WHERE codRelatorio = " & codRelatorio
Dim Modelo

    Set rRelatorio = CurrentDb.OpenRecordset(sqlRelatorio)
    
    If (Not rRelatorio.EOF) Then
        Modelo = rRelatorio("Modelo")
    End If
    
    rRelatorio.Close
    Set rRelatorio = Nothing
    
    Dim colQueries As New Collection
    GetQueries codRelatorio, Pagina, colQueries
    CreateRelatorioExcel codRelatorio, Modelo, colQueries
    
End Function

Public Function CreateRelatorioExcel(codRelatorio, Modelo, colQueries)
Dim XPlanilha As Object

'Historico "Carregar Modelo - Inicio: " & Now

    'Limpar Mascara
'    ExecutarSQL ("DELETE FROM admMascara")
    
    Set XPlanilha = CreateObject("Excel.Application")

    Dim arqModelo: arqModelo = Application.CurrentProject.Path & "\" & Modelo

    'Abre o arquivo modelo
    XPlanilha.Workbooks.Open (arqModelo)

    'Seleciona a primeira planilha
    XPlanilha.Workbooks(1).Sheets(1).Select
    
    Dim lastCell
    Dim currentCell
    With XPlanilha
        DoEvents
        lastCell = .Cells.Find(What:="[*]", LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False _
                , SearchFormat:=False).Address

    
        currentCell = .Cells.Find(What:="[*]", LookIn:=xlFormulas, LookAt:= _
                xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:=False _
                , SearchFormat:=False).Activate
    End With

    Do
        DoEvents
        With XPlanilha
            addressCell = .ActiveCell.Address
            InsertMascara codRelatorio, .ActiveCell.Text, .ActiveCell.Column, .ActiveCell.Row
            .Cells.FindNext(After:=.ActiveCell).Activate
        End With
    Loop While (Not IsEmpty(currentCell) And (addressCell <> lastCell))
    
    '' Collection
    For cq = 1 To colQueries.Count
        InsertModelo colQueries.Item(cq), codRelatorio, XPlanilha
    Next
    
    ''##################
    ''Formata Arquivo
    ''##################

    'Formata novo nome da planilha
    sTemp = "C:\Tmp\Rel" & Format(Now, "ddmmyy_hhnn") & ".xls"
    
    'Se o diretório não existe, cria
    If Dir$("C:\Tmp", vbDirectory) = "" Then MkDir "C:\Tmp"
    
    'Se o arquivo já existe, deleta
    If Dir$(sTemp) <> "" Then Kill sTemp
        
    ''##################
    ''Salva
    ''##################
        
    XPlanilha.ActiveWorkbook.SaveAs FileName:=sTemp, _
    FileFormat:=xlNormal, Password:="", WriteResPassword:="", _
    ReadOnlyRecommended:=False, CreateBackup:=False
    
    
    XPlanilha.Quit
    
    Set XPlanilha = Nothing
    
    'Limpar Mascara
    ExecutarSQL ("DELETE FROM admMascara")
    
    ImprimirExcel CStr(sTemp)
    
    'Se o arquivo já existe, deleta
    If Dir$(sTemp) <> "" Then Kill sTemp
    
'    Historico "Carregar Modelo - Terminio: " & Now
    
End Function

Sub InsertModelo(aryQuery, codRelatorio, XPlanilha)
Dim rQuery As DAO.Recordset
Dim rMascara As DAO.Recordset
Dim sqlMascara As String: sqlMascara = "SELECT * FROM admMascara WHERE Relatorio = '" & codRelatorio & "' AND origem = '" & aryQuery(1) & "'"
Dim sqlQuery As String: sqlQuery = aryQuery(2)
Dim sqlUpdate
Dim linha, coluna, value, expande
    
    Set rMascara = CurrentDb.OpenRecordset(sqlMascara)
    
    While Not rMascara.EOF
    
        Set rQuery = CurrentDb.OpenRecordset(sqlQuery)

            linha = rMascara("linha")
            coluna = rMascara("coluna")
            expande = rMascara("expande")

        If Not rQuery.EOF Then
            rQuery.MoveFirst
            
            If rQuery(rMascara("registro")) <> "" Then
                value = rQuery(rMascara("registro"))
            Else
                value = ""
            End If
            
            XPlanilha.Cells(linha, coluna).value = value
            
            If (expande > 0) Then
                rQuery.MoveNext
                Dim tmpLinha: tmpLinha = linha
                Dim ln: ln = 0
                Do While Not rQuery.EOF
                    
                    ln = ln + 1
                    tmpLinha = tmpLinha + 1
                    
                    If (CInt(ln) >= CInt(expande)) Then
                        Exit Do
                    End If
                    
                    
                    If rQuery(rMascara("registro")) <> "" Then
                        value = rQuery(rMascara("registro"))
                    Else
                        value = ""
                    End If
                    
                    XPlanilha.Cells(tmpLinha, coluna).value = value
                    rQuery.MoveNext
                Loop
            End If
            
        Else
            value = ""
            XPlanilha.Cells(linha, coluna).value = value
        End If
        
        rMascara.MoveNext
        
        rQuery.Close
    Set rQuery = Nothing
    Wend
    rMascara.Close
    Set rMascara = Nothing
End Sub

Sub GetQueries(ByVal codRelatorio, ByVal Pagina, ByRef colQueries)
Dim rConsultas As DAO.Recordset
Dim sqlQueries, sqlQueriesCount, origem, registro, lns, sqlPrimaria, sqlEstrangeira, sqlConsultas
Dim aryTmp(4) As String
    
    sqlConsultas = "SELECT * FROM admRelatoriosVinculos WHERE codRelatorio = " & codRelatorio
        
    Set rConsultas = CurrentDb.OpenRecordset(sqlConsultas)
    
    While Not rConsultas.EOF
        origem = rConsultas.Fields("Descricao")
        
        sqlQueries = "SELECT * FROM " & origem & " WHERE codRelatorio = '" & codRelatorio & "' AND Pagina = " & Pagina & ""
        
        aryTmp(1) = origem
        aryTmp(2) = sqlQueries
        
        colQueries.Add aryTmp
        rConsultas.MoveNext
    Wend
    rConsultas.Close
    Set rConsultas = Nothing
End Sub

Sub InsertMascara(codRelatorio, registro, coluna, linha)
Dim aryRegistro, sqlMascara, origem, Campo, primaria, estrangeira, expande
    
    primaria = 0
    estrangeira = 0
    expande = 0
    
    sqlMascara = "INSERT INTO admMascara(relatorio, origem, coluna, linha, registro, primaria, estrangeira, expande)VALUES('" & codRelatorio & "', "
    registro = Replace(registro, "[", "")
    registro = Replace(registro, "]", "")
    aryRegistro = Split(registro, "|")
    origem = aryRegistro(0)
    Campo = aryRegistro(1)
    Select Case (Mid(Campo, 1, 1))
        Case "!"
            primaria = 1
            Campo = Mid(Campo, 2, Len(Campo) - 1)
        Case "$"
            estrangeira = 1
            Campo = Mid(Campo, 2, Len(Campo) - 1)
    End Select
    
    If (UBound(aryRegistro) > 1) Then
        expande = aryRegistro(2)
    End If
    
    sqlMascara = sqlMascara & "'" & origem & "', " & coluna & ", " & linha & ", '" & Campo & "', " & primaria & ", " & estrangeira & ", " & expande & ")"
    ExecutarSQL (sqlMascara)
End Sub

Function FindLastRow(XPlanilha) As Variant
Dim LastRow As Variant

    If XPlanilha.WorksheetFunction.CountA(XPlanilha.Cells) > 0 Then
        LastRow = XPlanilha.Cells.Find(What:="*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    End If

    FindLastRow = LastRow
End Function

Function FindLastColumn(XPlanilha) As Variant
Dim LastColumn As Integer
    
    If XPlanilha.WorksheetFunction.CountA(XPlanilha.Cells) > 0 Then
        LastColumn = XPlanilha.Cells.Find(What:="*", SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    End If
    
    FindLastColumn = LastColumn
End Function

Public Function ExportarParaExcel(ssql As String, Modelo As String)
On Error GoTo Err_ExportarParaExcel

Dim DB As DAO.Database
Dim Rd As DAO.Recordset

Dim XPlanilha As Object

Dim iLinha As Integer
Dim intCampos As Integer
Dim I As Integer

Dim Count As Long

Dim sTemp As String
Dim arqModelo As String

Set DB = CurrentDb
Set Rd = DB.OpenRecordset(ssql)

If Not Rd.EOF Then

    Rd.MoveLast
    
    Count = Rd.RecordCount
    
    Rd.MoveFirst
    
    If Count > 0 Then
    
        DoEvents
        
        Dim s As Variant
        Dim C As Long
        
        'Cria referencia ao EXCEL
        Set XPlanilha = CreateObject("Excel.Application")
    
        ''##################
        ''Arquivo Modelo
        ''##################
    
            arqModelo = Application.CurrentProject.Path & "\" & Modelo
            
            'Abre o arquivo modelo
            XPlanilha.Workbooks.Open (arqModelo)
        
            'Seleciona a primeira planilha
            XPlanilha.Workbooks(1).Sheets(1).Select
        
            'Incrementa a linha
            iLinha = 6
        
        ''##################
        ''Transfere os dados
        ''##################
            
            intCampos = Rd.Fields.Count
            
            s = SysCmd(acSysCmdInitMeter, "Exportando " & Count & " Registros", Count)
        
            Do While Not Rd.EOF
                iLinha = iLinha + 1 'incrementa a linha
                I = 0
                    
                For I = 0 To intCampos - 1
                    XPlanilha.Cells(iLinha, I + 1).value = Rd(I)
                Next I
                
                s = SysCmd(acSysCmdUpdateMeter, iLinha)
                Rd.MoveNext
                
            Loop
            
            s = SysCmd(acSysCmdRemoveMeter)
            
'            'Linhas a repetir na parte superior - "CONTROLE DE CABEÇALHO"
'            XPlanilha.ActiveSheet.PageSetup.PrintTitleRows = "$1:$6"
'
'            'Área de Impressão - "CONTROLE DE COLUNAS"
'            XPlanilha.ActiveSheet.PageSetup.PrintArea = "$A$1:$" & rRelatorios.Fields("ColunasLimite") & "$" & iLinha
    
    
        ''##################
        ''Formata Arquivo
        ''##################
    
            'Formata novo nome da planilha
            sTemp = "C:\Tmp\Rel" & Format(Now, "ddmmyy_hhnn") & ".xls"
            'Se o diretório não existe, cria
            If Dir$("C:\Tmp", vbDirectory) = "" Then MkDir "C:\Tmp"
            'Se o arquivo já existe, deleta
            If Dir$(sTemp) <> "" Then Kill sTemp
            
      
        ''##################
        ''Salva
        ''##################
            
            XPlanilha.ActiveWorkbook.SaveAs FileName:=sTemp, _
            FileFormat:=xlNormal, Password:="", WriteResPassword:="", _
            ReadOnlyRecommended:=False, CreateBackup:=False
        
        ''##################
        ''Fecha o Excel
        ''##################
            XPlanilha.Quit
        
        ''######################
        ''Descarrega da memória
        ''######################
            Set XPlanilha = Nothing
            
            'Imprime direto
            ImprimirExcel CStr(sTemp)
                        
            'Se o arquivo já existe, deleta
            If Dir$(sTemp) <> "" Then Kill sTemp
            
    
    Else
    
        MsgBox "Não há dados para gerar a planilha.", vbInformation, "ATENÇÃO"
    
    End If

Else

    MsgBox "Não há Registros!", vbOKOnly + vbInformation, "Exportar para Excel"

End If

Rd.Close

Set Rd = Nothing
Set DB = Nothing

Set XPlanilha = Nothing

Exit_ExportarParaExcel:
    Exit Function

Err_ExportarParaExcel:
    Rd.Close
    
    Set Rd = Nothing
    Set DB = Nothing
    
    XPlanilha.ActiveWorkbook.SaveAs FileName:=sTemp, _
    FileFormat:=xlNormal, Password:="", WriteResPassword:="", _
    ReadOnlyRecommended:=False, CreateBackup:=False
    
    XPlanilha.Quit
    Set XPlanilha = Nothing
    
    
    MsgBox Err.Description
    Resume Exit_ExportarParaExcel
        
End Function

Public Function ImprimirExcel(Modelo As String)
Dim XPlanilha As Object

    Set XPlanilha = CreateObject("Excel.Application")
    
    With XPlanilha

        'Abre o arquivo modelo
        .Workbooks.Open (Modelo)

        'Seleciona a primeira planilha
        .Workbooks(1).Sheets(1).Select
    
        .ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True
        
        'Desativar mensagens de alerta
        .Application.DisplayAlerts = False

        'Fechar Excel
        .Quit
        
        'Desativar mensagens de alerta
        .Application.DisplayAlerts = True
    
    End With
    
'Desassociar a variável
Set XPlanilha = Nothing

End Function

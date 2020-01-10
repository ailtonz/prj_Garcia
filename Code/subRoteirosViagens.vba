Option Compare Database

Private Sub codAterro_Click()
    Me.Aterro = Me.codAterro.Column(1)
End Sub

Private Sub Obra_Click()
    
    Me.codObra = Me.Obra.Column(1)
    Me.codCadastro = Me.Obra.Column(2)
    Me.Cliente = Me.Obra.Column(3)
    Me.C = Me.Obra.Column(4)
    Me.R = Me.Obra.Column(5)
    Me.T = Me.Obra.Column(6)
    Me.CTR = Me.Obra.Column(7)
    
    Me.DT_C = Me.Obra.Column(8)
    Me.DT_R = Me.Obra.Column(9)
    Me.DT_T = Me.Obra.Column(10)

End Sub

Private Sub Obra_Enter()
    Me.Obra.Requery
End Sub

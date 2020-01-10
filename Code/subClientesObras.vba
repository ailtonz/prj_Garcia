Option Compare Database

Private Sub Contato_Click()
    Me.Telefone = Me.Contato.Column(1)
    Me.Email = Me.Contato.Column(2)
End Sub

Private Sub Contato_Exit(Cancel As Integer)
    If Me.Contato <> "" Then Me.Contato = UCase(Me.Contato)
End Sub

Private Sub Email_Exit(Cancel As Integer)
    If Me.Email <> "" Then Me.Email = LCase(Me.Email)
End Sub


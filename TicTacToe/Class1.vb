Public Class Cell
    Public btn As Button
    Public state As Char = ""
    Public clickable As Boolean = True

    Public Sub New(size As Size, pos As Point, frm As Form)
        'Init the button as a button and add it's handler
        btn = New Button
        btn.Text = ""
        AddHandler btn.Click, AddressOf clicked

        ' Set the size and pos
        btn.Location = pos
        btn.Size = size

        'Add to the form
        frm.Controls.Add(btn)
    End Sub

    Private Sub clicked()
        If (clickable) Then
            state = If(frmGame.turn, "X", "O")
            btn.Text = state

            frmGame.checkState()
            frmGame.turn = Not frmGame.turn
            clickable = False
        End If
    End Sub
End Class

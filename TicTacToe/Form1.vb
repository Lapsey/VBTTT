Public Class frmGame
    'Drune Rugland, March 11, 2022

    'This code is in DIRE need of refactoring 

    ' TODO:
    ' -Possible Make a game over sub that displays the message box
    ' -Manual restart button
    ' -Other general refactors 


    Public turn As Boolean = True ' True = X false = O
    Public board(2, 2) As Cell
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        resetBoard()
    End Sub

    Public Sub checkState()
        For row As Short = 0 To 2
            If (Char.IsLetter(board(row, 0).state) And Char.IsLetter(board(row, 1).state) And Char.IsLetter(board(row, 2).state)) Then
                If (board(row, 0).state = board(row, 1).state And board(row, 1).state = board(row, 2).state) Then
                    Dim result As DialogResult = MessageBox.Show(If(turn, "X", "O") & " Wins... Would you like to play again?", "Win", MessageBoxButtons.YesNo)

                    If (result = DialogResult.Yes) Then
                        resetBoard()
                    Else
                        End
                    End If
                End If
            End If
        Next

        For col As Short = 0 To 2
            If (Char.IsLetter(board(0, col).state) And Char.IsLetter(board(1, col).state) And Char.IsLetter(board(2, col).state)) Then
                If (board(0, col).state = board(1, col).state And board(1, col).state = board(2, col).state) Then
                    Dim result As DialogResult = MessageBox.Show(If(turn, "X", "O") & " Wins... Would you like to play again?", "Win", MessageBoxButtons.YesNo)

                    If (result = DialogResult.Yes) Then
                        resetBoard()
                    Else
                        End
                    End If
                End If
            End If
        Next

        If (Char.IsLetter(board(0, 0).state) And Char.IsLetter(board(1, 1).state) And Char.IsLetter(board(2, 2).state)) Then
            If (board(0, 0).state = board(1, 1).state And board(1, 1).state = board(2, 2).state) Then
                Dim result As DialogResult = MessageBox.Show(If(turn, "X", "O") & " Wins... Would you like to play again?", "Win", MessageBoxButtons.YesNo)

                If (result = DialogResult.Yes) Then
                    resetBoard()
                Else
                    End
                End If
            End If
        End If

        If (Char.IsLetter(board(0, 2).state) And Char.IsLetter(board(1, 1).state) And Char.IsLetter(board(2, 0).state)) Then
            If (board(0, 2).state = board(1, 1).state And board(1, 1).state = board(2, 0).state) Then
                Dim result As DialogResult = MessageBox.Show(If(turn, "X", "O") & " Wins... Would you like to play again?", "Win", MessageBoxButtons.YesNo)

                If (result = DialogResult.Yes) Then
                    resetBoard()
                Else
                    End
                End If
            End If
        End If

        If (isTie()) Then
            Dim result As DialogResult = MessageBox.Show("You tied... Would you like to play again", "Tie", MessageBoxButtons.YesNo)

            If (result = DialogResult.Yes) Then
                resetBoard()
            Else
                End
            End If
        End If
    End Sub

    Private Function isTie()
        For i As Short = 0 To 2
            For j As Short = 0 To 2
                If (Not Char.IsLetter(board(i, j).state)) Then
                    Return False
                End If
            Next
        Next

        Return True
    End Function

    Private Sub resetBoard()
        Me.Controls.Clear()

        For i As Short = 0 To 2
            For j As Short = 0 To 2
                Dim pos As New Point(i * 100, j * 100)
                Dim size = New Size(100, 100)

                Dim cell As New Cell(size, pos, Me)
                board(j, i) = cell
            Next
        Next
    End Sub
End Class

Public Class frmSpeedHome
    Private Sub btnNewGame_Click(sender As Object, e As EventArgs) Handles btnNewGame.Click
        frmNewGameOptions.Show()            'show game options form
        frmNewGameOptions.BringToFront()
    End Sub

    Private Sub btnHowToPlay_Click(sender As Object, e As EventArgs) Handles btnHowToPlay.Click
        'show how to play
        MessageBox.Show("SPEED" & vbCrLf & vbCrLf & howToPlay, "How To Play:", MessageBoxButtons.OK)
    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Application.Exit()                   'exit and close application
    End Sub
End Class

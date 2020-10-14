Public Class frmSpeedGame
    Public Property btnCancel_Click As Boolean
    Dim Paused As Boolean

    Private Sub frmSpeedGame_Load(sender As Object, e As EventArgs) Handles Me.Load
        Paused = False

        picCompPile.AllowDrop = True
        picPlayerPile.AllowDrop = True

        tmrGame.Start()

        If MouseButtons.Left Then               'once player left clicks anywhere on the form, the computer will start playing their cards
            If compEasy = True Then             'if computer is set to easy mode, it will place a card every 5 seconds (tmrEasy)
                tmrEasy.Start()
                tmrMedium.Stop()
                tmrHard.Stop()
            ElseIf compMedium = True Then       'if computer is set to medium mode, it will place a card every 3.5 seconds (tmrMedium)
                tmrEasy.Stop()
                tmrMedium.Start()
                tmrHard.Stop()
            ElseIf compHard = True Then         'if computer is set to hard mode, it will place a card every 2 seconds (tmrHard)
                tmrEasy.Stop()
                tmrMedium.Stop()
                tmrHard.Start()
            End If
        End If

        Me.Focus()
    End Sub

    Private Sub frmSpeedGame_KeyPress(sender As Object, e As KeyPressEventArgs) Handles Me.KeyPress

        If e.KeyChar = "P" Or e.KeyChar = "p" Then              'if P key is pressed, game is paused
            If Paused = True Then                               'if P key is pressed again, game is resumed
                tmrGame.Enabled = True
                If compEasy = True Then
                    tmrEasy.Start()
                    tmrMedium.Stop()
                    tmrHard.Stop()
                ElseIf compMedium = True Then
                    tmrEasy.Stop()
                    tmrMedium.Start()
                    tmrHard.Stop()
                ElseIf compHard = True Then
                    tmrEasy.Stop()
                    tmrMedium.Stop()
                    tmrHard.Start()
                End If

                btnShuffle.Enabled = True
                Paused = False

                lblPause.Text = "Press P to Pause"
            Else
                tmrGame.Enabled = False
                tmrEasy.Enabled = False
                tmrMedium.Enabled = False
                tmrHard.Enabled = False

                btnShuffle.Enabled = False
                Paused = True
                Me.Focus()

                lblPause.Text = "Game Paused" & vbCrLf & "Press P to Resume"
            End If
        End If

    End Sub

    Private Sub tmrGame_Tick(sender As Object, e As EventArgs) Handles tmrGame.Tick
        Dim back As Integer = -100
        Dim showButton1 As Boolean = True
        Dim showButton2 As Boolean = True
        Dim showButton3 As Boolean = True
        Dim showButton4 As Boolean = True

        'show button ONLY if player and computer cannot move anymore cards
        For num As Integer = 1 To 13
            For numCard As Integer = 0 To 4
                If compPile = num And (compCard(numCard) = num + 1 Or compCard(numCard) = num - 1) Then
                    showButton1 = False
                ElseIf compPile = 1 And (compCard(numCard) = 2 Or compCard(numCard) = 13) Then
                    showButton1 = False
                ElseIf compPile = 13 And (compCard(numCard) = 1 Or compCard(numCard) = 12) Then
                    showButton1 = False
                End If

                If playerPile = num And (compCard(numCard) = num + 1 Or compCard(numCard) = num - 1) Then
                    showButton2 = False
                ElseIf playerPile = 1 And (compCard(numCard) = 2 Or compCard(numCard) = 13) Then
                    showButton2 = False
                ElseIf playerPile = 13 And (compCard(numCard) = 1 Or compCard(numCard) = 12) Then
                    showButton2 = False
                End If

                If compPile = num And (playerCard(numCard) = num + 1 Or playerCard(numCard) = num - 1) Then
                    showButton3 = False
                ElseIf compPile = 1 And (playerCard(numCard) = 2 Or playerCard(numCard) = 13) Then
                    showButton3 = False
                ElseIf compPile = 13 And (playerCard(numCard) = 1 Or playerCard(numCard) = 12) Then
                    showButton3 = False
                End If

                If playerPile = num And (playerCard(numCard) = num + 1 Or playerCard(numCard) = num - 1) Then
                    showButton4 = False
                ElseIf playerPile = 1 And (playerCard(numCard) = 2 Or playerCard(numCard) = 13) Then
                    showButton4 = False
                ElseIf playerPile = 13 And (playerCard(numCard) = 1 Or playerCard(numCard) = 12) Then
                    showButton4 = False
                End If
            Next
        Next

        If showButton1 = True And showButton2 = True And showButton3 = True And showButton4 = True Then         'if player and computer cannot place card in either pile, then show button
            btnShuffle.Enabled = True
            btnShuffle.Visible = True
            btnShuffle.Show()
            btnShuffle.Focus()
            lblPause.Text = "Press P to Pause" & vbCrLf & "Press Space Bar to Shuffle"
        Else
            btnShuffle.Enabled = False
            btnShuffle.Visible = False
            btnShuffle.Hide()
            lblPause.Text = "Press P to Pause"
            Me.Focus()
        End If

        If frmNewGameOptions.radBlue.Checked = True Then
            If compCardsLeft1 = 0 Then                                  'if comp has 0 cards left, their compCard will be dislayed as a card back
                Me.picCompCard1.Image = My.Resources.CardBackBlue
                compCard(0) = back
                card(0) = back
                picCompCard1.Enabled = False
            End If
            If compCardsLeft2 = 0 Then
                Me.picCompCard2.Image = My.Resources.CardBackBlue
                compCard(1) = back
                card(1) = back
                picCompCard2.Enabled = False
            End If
            If compCardsLeft3 = 0 Then
                Me.picCompCard3.Image = My.Resources.CardBackBlue
                compCard(2) = back
                card(2) = back
                picCompCard3.Enabled = False
            End If
            If compCardsLeft4 = 0 Then
                Me.picCompCard4.Image = My.Resources.CardBackBlue
                compCard(3) = back
                card(3) = back
                picCompCard4.Enabled = False
            End If
            If compCardsLeft5 = 0 Then
                Me.picCompCard5.Image = My.Resources.CardBackBlue
                compCard(4) = back
                card(4) = back
                picCompCard5.Enabled = False
            End If

            If playerCardsLeft1 = 0 Then
                Me.picPlayerCard1.Image = My.Resources.CardBackBlue
                playerCard(0) = back
                card(5) = back
            End If
            If playerCardsLeft2 = 0 Then
                Me.picPlayerCard2.Image = My.Resources.CardBackBlue
                playerCard(1) = back
                card(6) = back
            End If
            If playerCardsLeft3 = 0 Then
                Me.picPlayerCard3.Image = My.Resources.CardBackBlue
                playerCard(2) = back
                card(7) = back
            End If
            If playerCardsLeft4 = 0 Then
                Me.picPlayerCard4.Image = My.Resources.CardBackBlue
                playerCard(3) = back
                card(8) = back
            End If
            If playerCardsLeft5 = 0 Then
                Me.picPlayerCard5.Image = My.Resources.CardBackBlue
                playerCard(4) = back
                card(9) = back
            End If
        ElseIf frmNewGameOptions.radRed.Checked = True Then
            If compCardsLeft1 = 0 Then
                Me.picCompCard1.Image = My.Resources.CardBackRed
                compCard(0) = back
                card(0) = back
            End If
            If compCardsLeft2 = 0 Then
                Me.picCompCard2.Image = My.Resources.CardBackRed
                compCard(1) = back
                card(1) = back
            End If
            If compCardsLeft3 = 0 Then
                Me.picCompCard3.Image = My.Resources.CardBackRed
                compCard(2) = back
                card(2) = back
            End If
            If compCardsLeft4 = 0 Then
                Me.picCompCard4.Image = My.Resources.CardBackRed
                compCard(3) = back
                card(3) = back
            End If
            If compCardsLeft5 = 0 Then
                Me.picCompCard5.Image = My.Resources.CardBackRed
                compCard(4) = back
                card(4) = back
            End If

            If playerCardsLeft1 = 0 Then
                Me.picPlayerCard1.Image = My.Resources.CardBackRed
                playerCard(0) = back
                card(5) = back
            End If
            If playerCardsLeft2 = 0 Then
                Me.picPlayerCard2.Image = My.Resources.CardBackRed
                playerCard(1) = back
                card(6) = back
            End If
            If playerCardsLeft3 = 0 Then
                Me.picPlayerCard3.Image = My.Resources.CardBackRed
                playerCard(2) = back
                card(7) = back
            End If
            If playerCardsLeft4 = 0 Then
                Me.picPlayerCard4.Image = My.Resources.CardBackRed
                playerCard(3) = back
                card(8) = back
            End If
            If playerCardsLeft5 = 0 Then
                Me.picPlayerCard5.Image = My.Resources.CardBackRed
                playerCard(4) = back
                card(9) = back
            End If
        ElseIf frmNewGameOptions.radBlack.Checked = True Then
            If compCardsLeft1 = 0 Then
                Me.picCompCard1.Image = My.Resources.CardBackBlack
                compCard(0) = back
                card(0) = back
            End If
            If compCardsLeft2 = 0 Then
                Me.picCompCard2.Image = My.Resources.CardBackBlack
                compCard(1) = back
                card(1) = back
            End If
            If compCardsLeft3 = 0 Then
                Me.picCompCard3.Image = My.Resources.CardBackBlack
                compCard(2) = back
                card(2) = back
            End If
            If compCardsLeft4 = 0 Then
                Me.picCompCard4.Image = My.Resources.CardBackBlack
                compCard(3) = back
                card(3) = back
            End If
            If compCardsLeft5 = 0 Then
                Me.picCompCard5.Image = My.Resources.CardBackBlack
                compCard(4) = back
                card(4) = back
            End If

            If playerCardsLeft1 = 0 Then
                Me.picPlayerCard1.Image = My.Resources.CardBackBlack
                playerCard(0) = back
                card(5) = back
            End If
            If playerCardsLeft2 = 0 Then
                Me.picPlayerCard2.Image = My.Resources.CardBackBlack
                playerCard(1) = back
                card(6) = back
            End If
            If playerCardsLeft3 = 0 Then
                Me.picPlayerCard3.Image = My.Resources.CardBackBlack
                playerCard(2) = back
                card(7) = back
            End If
            If playerCardsLeft4 = 0 Then
                Me.picPlayerCard4.Image = My.Resources.CardBackBlack
                playerCard(3) = back
                card(8) = back
            End If
            If playerCardsLeft5 = 0 Then
                Me.picPlayerCard5.Image = My.Resources.CardBackBlack
                playerCard(4) = back
                card(9) = back
            End If
        End If

        If playerCardsLeft1 = 0 And playerCardsLeft2 = 0 And playerCardsLeft3 = 0 And playerCardsLeft4 = 0 And playerCardsLeft5 = 0 Then
            tmrGame.Stop()                          'if all player cards have been placed, then the player has won
            tmrEasy.Stop()
            tmrMedium.Stop()
            tmrHard.Stop()
            btnShuffle.Enabled = False
            lblPause.Text = Nothing
            MessageBox.Show("Congratulations! You have Won!")
        End If

        If compCardsLeft1 = 0 And compCardsLeft2 = 0 And compCardsLeft3 = 0 And compCardsLeft4 = 0 And compCardsLeft5 = 0 Then
            tmrGame.Stop()                          'if all computer cards have been placed, then the computer has won
            tmrEasy.Stop()
            tmrMedium.Stop()
            tmrHard.Stop()
            btnShuffle.Enabled = False
            lblPause.Text = Nothing
            MessageBox.Show("Computer has won!")
        End If

    End Sub

    Private Sub tmrEasy_Tick(sender As Object, e As EventArgs) Handles tmrEasy.Tick
        Dim newNum As Integer
        Dim OK As Boolean = False

        For cardNum As Integer = 0 To 4                                                                'check for value of compCards [0 to 4]
            If compPile = compCard(cardNum) + 1 Or compPile = compCard(cardNum) - 1 Then
                Select Case cardNum
                    Case 0                                                  'if compCard(0) is 1 higher or 1 lower than compPile
                        Me.picCompPile.Image = Me.picCompCard1.Image        'set picCompPile to picCompCard1 image
                        compPile = compCard(0)                              'set compPile to compCard(0) integer [1 to 13]
                        card(10) = card(0)                                  'set card(10), compPileCard, to card(0) integer [1 to 52]
                        If compCardsLeft1 > 0 Then                          'if 1 or more cards left, 
                            compCardsLeft1 -= 1                             'subtract 1
                        End If
                    Case 1
                        Me.picCompPile.Image = Me.picCompCard2.Image
                        compPile = compCard(1)
                        card(10) = card(1)
                        If compCardsLeft2 > 0 Then
                            compCardsLeft2 -= 1
                        End If
                    Case 2
                        Me.picCompPile.Image = Me.picCompCard3.Image
                        compPile = compCard(2)
                        card(10) = card(2)
                        If compCardsLeft3 > 0 Then
                            compCardsLeft3 -= 1
                        End If
                    Case 3
                        Me.picCompPile.Image = Me.picCompCard4.Image
                        compPile = compCard(3)
                        card(10) = card(3)
                        If compCardsLeft4 > 0 Then
                            compCardsLeft4 -= 1
                        End If
                    Case 4
                        Me.picCompPile.Image = Me.picCompCard5.Image
                        compPile = compCard(4)
                        card(10) = card(4)
                        If compCardsLeft5 > 0 Then
                            compCardsLeft5 -= 1
                        End If
                End Select
                cardNum = 100

            ElseIf compPile = 1 And (compCard(cardNum) = 2 Or compCard(cardNum) = 13) Then
                Select Case cardNum
                    Case 0
                        Me.picCompPile.Image = Me.picCompCard1.Image
                        compPile = compCard(0)
                        card(10) = card(0)
                        If compCardsLeft1 > 0 Then
                            compCardsLeft1 -= 1
                        End If
                    Case 1
                        Me.picCompPile.Image = Me.picCompCard2.Image
                        compPile = compCard(1)
                        card(10) = card(1)
                        If compCardsLeft2 > 0 Then
                            compCardsLeft2 -= 1
                        End If
                    Case 2
                        Me.picCompPile.Image = Me.picCompCard3.Image
                        compPile = compCard(2)
                        card(10) = card(2)
                        If compCardsLeft3 > 0 Then
                            compCardsLeft3 -= 1
                        End If
                    Case 3
                        Me.picCompPile.Image = Me.picCompCard4.Image
                        compPile = compCard(3)
                        card(10) = card(3)
                        If compCardsLeft4 > 0 Then
                            compCardsLeft4 -= 1
                        End If
                    Case 4
                        Me.picCompPile.Image = Me.picCompCard5.Image
                        compPile = compCard(4)
                        card(10) = card(4)
                        If compCardsLeft5 > 0 Then
                            compCardsLeft5 -= 1
                        End If
                End Select
                cardNum = 100

            ElseIf compPile = 13 And (compCard(cardNum) = 1 Or compCard(cardNum) = 12) Then
                Select Case cardNum
                    Case 0
                        Me.picCompPile.Image = Me.picCompCard1.Image
                        compPile = compCard(0)
                        card(10) = card(0)
                        If compCardsLeft1 > 0 Then
                            compCardsLeft1 -= 1
                        End If
                    Case 1
                        Me.picCompPile.Image = Me.picCompCard2.Image
                        compPile = compCard(1)
                        card(10) = card(1)
                        If compCardsLeft2 > 0 Then
                            compCardsLeft2 -= 1
                        End If
                    Case 2
                        Me.picCompPile.Image = Me.picCompCard3.Image
                        compPile = compCard(2)
                        card(10) = card(2)
                        If compCardsLeft3 > 0 Then
                            compCardsLeft3 -= 1
                        End If
                    Case 3
                        Me.picCompPile.Image = Me.picCompCard4.Image
                        compPile = compCard(3)
                        card(10) = card(3)
                        If compCardsLeft4 > 0 Then
                            compCardsLeft4 -= 1
                        End If
                    Case 4
                        Me.picCompPile.Image = Me.picCompCard5.Image
                        compPile = compCard(4)
                        card(10) = card(4)
                        If compCardsLeft5 > 0 Then
                            compCardsLeft5 -= 1
                        End If
                End Select
                cardNum = 100

            ElseIf playerPile = compCard(cardNum) + 1 Or playerPile = compCard(cardNum) - 1 Then
                Select Case cardNum
                    Case 0
                        Me.picPlayerPile.Image = Me.picCompCard1.Image
                        playerPile = compCard(0)
                        card(11) = card(0)
                        If compCardsLeft1 > 0 Then
                            compCardsLeft1 -= 1
                        End If
                    Case 1
                        Me.picPlayerPile.Image = Me.picCompCard2.Image
                        playerPile = compCard(1)
                        card(11) = card(1)
                        If compCardsLeft2 > 0 Then
                            compCardsLeft2 -= 1
                        End If
                    Case 2
                        Me.picPlayerPile.Image = Me.picCompCard3.Image
                        playerPile = compCard(2)
                        card(11) = card(2)
                        If compCardsLeft3 > 0 Then
                            compCardsLeft3 -= 1
                        End If
                    Case 3
                        Me.picPlayerPile.Image = Me.picCompCard4.Image
                        playerPile = compCard(3)
                        card(11) = card(3)
                        If compCardsLeft4 > 0 Then
                            compCardsLeft4 -= 1
                        End If
                    Case 4
                        Me.picPlayerPile.Image = Me.picCompCard5.Image
                        playerPile = compCard(4)
                        card(11) = card(4)
                        If compCardsLeft5 > 0 Then
                            compCardsLeft5 -= 1
                        End If
                End Select
                cardNum = 100

            ElseIf playerPile = 1 And (compCard(cardNum) = 2 Or compCard(cardNum) = 13) Then
                Select Case cardNum
                    Case 0
                        Me.picPlayerPile.Image = Me.picCompCard1.Image
                        playerPile = compCard(0)
                        card(11) = card(0)
                        If compCardsLeft1 > 0 Then
                            compCardsLeft1 -= 1
                        End If
                    Case 1
                        Me.picPlayerPile.Image = Me.picCompCard2.Image
                        playerPile = compCard(1)
                        card(11) = card(1)
                        If compCardsLeft2 > 0 Then
                            compCardsLeft2 -= 1
                        End If
                    Case 2
                        Me.picPlayerPile.Image = Me.picCompCard3.Image
                        playerPile = compCard(2)
                        card(11) = card(2)
                        If compCardsLeft3 > 0 Then
                            compCardsLeft3 -= 1
                        End If
                    Case 3
                        Me.picPlayerPile.Image = Me.picCompCard4.Image
                        playerPile = compCard(3)
                        card(11) = card(3)
                        If compCardsLeft4 > 0 Then
                            compCardsLeft4 -= 1
                        End If
                    Case 4
                        Me.picPlayerPile.Image = Me.picCompCard5.Image
                        playerPile = compCard(4)
                        card(11) = card(4)
                        If compCardsLeft5 > 0 Then
                            compCardsLeft5 -= 1
                        End If
                End Select
                cardNum = 100

            ElseIf playerPile = 13 And (compCard(cardNum) = 1 Or compCard(cardNum) = 12) Then
                Select Case cardNum
                    Case 0
                        Me.picPlayerPile.Image = Me.picCompCard1.Image
                        playerPile = compCard(0)
                        card(11) = card(0)
                        If compCardsLeft1 > 0 Then
                            compCardsLeft1 -= 1
                        End If
                    Case 1
                        Me.picPlayerPile.Image = Me.picCompCard2.Image
                        playerPile = compCard(1)
                        card(11) = card(1)
                        If compCardsLeft2 > 0 Then
                            compCardsLeft2 -= 1
                        End If
                    Case 2
                        Me.picPlayerPile.Image = Me.picCompCard3.Image
                        playerPile = compCard(2)
                        card(11) = card(2)
                        If compCardsLeft3 > 0 Then
                            compCardsLeft3 -= 1
                        End If
                    Case 3
                        Me.picPlayerPile.Image = Me.picCompCard4.Image
                        playerPile = compCard(3)
                        card(11) = card(3)
                        If compCardsLeft4 > 0 Then
                            compCardsLeft4 -= 1
                        End If
                    Case 4
                        Me.picPlayerPile.Image = Me.picCompCard5.Image
                        playerPile = compCard(4)
                        card(11) = card(4)
                        If compCardsLeft5 > 0 Then
                            compCardsLeft5 -= 1
                        End If
                End Select
                cardNum = 100

            End If
        Next cardNum

        If card(10) = card(0) Or card(10) = card(1) Or card(10) = card(2) Or card(10) = card(3) Or card(10) = card(4) Then

            Do                                                      'generate a new card
                Randomize()
                newNum = Int(max * Rnd() + min)

                If Used.Contains(newNum) Then                       'if newNum has already been generated, do again
                    OK = False
                Else                                                'if newNum has not already been generated, show new card
                    OK = True

                    If card(10) = card(0) Then                      'if card(10), compPileCard, is the same as card(0), compPile1 integer
                        Select Case newNum
                            Case 1
                                Me.picCompCard1.Image = card1
                                compCard(0) = 1
                                card(0) = 1
                            Case 2
                                Me.picCompCard1.Image = card2
                                compCard(0) = 2
                                card(0) = 2
                            Case 3
                                Me.picCompCard1.Image = card3
                                compCard(0) = 3
                                card(0) = 3
                            Case 4
                                Me.picCompCard1.Image = card4
                                compCard(0) = 4
                                card(0) = 4
                            Case 5
                                Me.picCompCard1.Image = card5
                                compCard(0) = 5
                                card(0) = 5
                            Case 6
                                Me.picCompCard1.Image = card6
                                compCard(0) = 6
                                card(0) = 6
                            Case 7
                                Me.picCompCard1.Image = card7
                                compCard(0) = 7
                                card(0) = 7
                            Case 8
                                Me.picCompCard1.Image = card8
                                compCard(0) = 8
                                card(0) = 8
                            Case 9
                                Me.picCompCard1.Image = card9
                                compCard(0) = 9
                                card(0) = 9
                            Case 10
                                Me.picCompCard1.Image = card10
                                compCard(0) = 10
                                card(0) = 10
                            Case 11
                                Me.picCompCard1.Image = card11
                                compCard(0) = 11
                                card(0) = 11
                            Case 12
                                Me.picCompCard1.Image = card12
                                compCard(0) = 12
                                card(0) = 12
                            Case 13
                                Me.picCompCard1.Image = card13
                                compCard(0) = 13
                                card(0) = 13
                            Case 14
                                Me.picCompCard1.Image = card14
                                compCard(0) = 1
                                card(0) = 14
                            Case 15
                                Me.picCompCard1.Image = card15
                                compCard(0) = 2
                                card(0) = 15
                            Case 16
                                Me.picCompCard1.Image = card16
                                compCard(0) = 3
                                card(0) = 16
                            Case 17
                                Me.picCompCard1.Image = card17
                                compCard(0) = 4
                                card(0) = 17
                            Case 18
                                Me.picCompCard1.Image = card18
                                compCard(0) = 5
                                card(0) = 18
                            Case 19
                                Me.picCompCard1.Image = card19
                                compCard(0) = 6
                                card(0) = 19
                            Case 20
                                Me.picCompCard1.Image = card20
                                compCard(0) = 7
                                card(0) = 20
                            Case 21
                                Me.picCompCard1.Image = card21
                                compCard(0) = 8
                                card(0) = 21
                            Case 22
                                Me.picCompCard1.Image = card22
                                compCard(0) = 9
                                card(0) = 22
                            Case 23
                                Me.picCompCard1.Image = card23
                                compCard(0) = 10
                                card(0) = 23
                            Case 24
                                Me.picCompCard1.Image = card24
                                compCard(0) = 11
                                card(0) = 24
                            Case 25
                                Me.picCompCard1.Image = card25
                                compCard(0) = 12
                                card(0) = 25
                            Case 26
                                Me.picCompCard1.Image = card26
                                compCard(0) = 13
                                card(0) = 26
                            Case 27
                                Me.picCompCard1.Image = card27
                                compCard(0) = 1
                                card(0) = 27
                            Case 28
                                Me.picCompCard1.Image = card28
                                compCard(0) = 2
                                card(0) = 28
                            Case 29
                                Me.picCompCard1.Image = card29
                                compCard(0) = 3
                                card(0) = 29
                            Case 30
                                Me.picCompCard1.Image = card30
                                compCard(0) = 4
                                card(0) = 30
                            Case 31
                                Me.picCompCard1.Image = card31
                                compCard(0) = 5
                                card(0) = 31
                            Case 32
                                Me.picCompCard1.Image = card32
                                compCard(0) = 6
                                card(0) = 32
                            Case 33
                                Me.picCompCard1.Image = card33
                                compCard(0) = 7
                                card(0) = 33
                            Case 34
                                Me.picCompCard1.Image = card34
                                compCard(0) = 8
                                card(0) = 34
                            Case 35
                                Me.picCompCard1.Image = card35
                                compCard(0) = 9
                                card(0) = 35
                            Case 36
                                Me.picCompCard1.Image = card36
                                compCard(0) = 10
                                card(0) = 36
                            Case 37
                                Me.picCompCard1.Image = card37
                                compCard(0) = 11
                                card(0) = 37
                            Case 38
                                Me.picCompCard1.Image = card38
                                compCard(0) = 12
                                card(0) = 38
                            Case 39
                                Me.picCompCard1.Image = card39
                                compCard(0) = 13
                                card(0) = 39
                            Case 40
                                Me.picCompCard1.Image = card40
                                compCard(0) = 1
                                card(0) = 40
                            Case 41
                                Me.picCompCard1.Image = card41
                                compCard(0) = 2
                                card(0) = 41
                            Case 42
                                Me.picCompCard1.Image = card42
                                compCard(0) = 3
                                card(0) = 42
                            Case 43
                                Me.picCompCard1.Image = card43
                                compCard(0) = 4
                                card(0) = 43
                            Case 44
                                Me.picCompCard1.Image = card44
                                compCard(0) = 5
                                card(0) = 44
                            Case 45
                                Me.picCompCard1.Image = card45
                                compCard(0) = 6
                                card(0) = 45
                            Case 46
                                Me.picCompCard1.Image = card46
                                compCard(0) = 7
                                card(0) = 46
                            Case 47
                                Me.picCompCard1.Image = card47
                                compCard(0) = 8
                                card(0) = 47
                            Case 48
                                Me.picCompCard1.Image = card48
                                compCard(0) = 9
                                card(0) = 48
                            Case 49
                                Me.picCompCard1.Image = card49
                                compCard(0) = 10
                                card(0) = 49
                            Case 50
                                Me.picCompCard1.Image = card50
                                compCard(0) = 11
                                card(0) = 50
                            Case 51
                                Me.picCompCard1.Image = card51
                                compCard(0) = 12
                                card(0) = 51
                            Case 52
                                Me.picCompCard1.Image = card52
                                compCard(0) = 13
                                card(0) = 52
                        End Select
                    ElseIf card(10) = card(1) Then
                        Select Case newNum
                            Case 1
                                Me.picCompCard2.Image = card1
                                compCard(1) = 1
                                card(1) = 1
                            Case 2
                                Me.picCompCard2.Image = card2
                                compCard(1) = 2
                                card(1) = 2
                            Case 3
                                Me.picCompCard2.Image = card3
                                compCard(1) = 3
                                card(1) = 3
                            Case 4
                                Me.picCompCard2.Image = card4
                                compCard(1) = 4
                                card(1) = 4
                            Case 5
                                Me.picCompCard2.Image = card5
                                compCard(1) = 5
                                card(1) = 5
                            Case 6
                                Me.picCompCard2.Image = card6
                                compCard(1) = 6
                                card(1) = 6
                            Case 7
                                Me.picCompCard2.Image = card7
                                compCard(1) = 7
                                card(1) = 7
                            Case 8
                                Me.picCompCard2.Image = card8
                                compCard(1) = 8
                                card(1) = 8
                            Case 9
                                Me.picCompCard2.Image = card9
                                compCard(1) = 9
                                card(1) = 9
                            Case 10
                                Me.picCompCard2.Image = card10
                                compCard(1) = 10
                                card(1) = 10
                            Case 11
                                Me.picCompCard2.Image = card11
                                compCard(1) = 11
                                card(1) = 11
                            Case 12
                                Me.picCompCard2.Image = card12
                                compCard(1) = 12
                                card(1) = 12
                            Case 13
                                Me.picCompCard2.Image = card13
                                compCard(1) = 13
                                card(1) = 13
                            Case 14
                                Me.picCompCard2.Image = card14
                                compCard(1) = 1
                                card(1) = 14
                            Case 15
                                Me.picCompCard2.Image = card15
                                compCard(1) = 2
                                card(1) = 15
                            Case 16
                                Me.picCompCard2.Image = card16
                                compCard(1) = 3
                                card(1) = 16
                            Case 17
                                Me.picCompCard2.Image = card17
                                compCard(1) = 4
                                card(1) = 17
                            Case 18
                                Me.picCompCard2.Image = card18
                                compCard(1) = 5
                                card(1) = 18
                            Case 19
                                Me.picCompCard2.Image = card19
                                compCard(1) = 6
                                card(1) = 19
                            Case 20
                                Me.picCompCard2.Image = card20
                                compCard(1) = 7
                                card(1) = 20
                            Case 21
                                Me.picCompCard2.Image = card21
                                compCard(1) = 8
                                card(1) = 21
                            Case 22
                                Me.picCompCard2.Image = card22
                                compCard(1) = 9
                                card(1) = 22
                            Case 23
                                Me.picCompCard2.Image = card23
                                compCard(1) = 10
                                card(1) = 23
                            Case 24
                                Me.picCompCard2.Image = card24
                                compCard(1) = 11
                                card(1) = 24
                            Case 25
                                Me.picCompCard2.Image = card25
                                compCard(1) = 12
                                card(1) = 25
                            Case 26
                                Me.picCompCard2.Image = card26
                                compCard(1) = 13
                                card(1) = 26
                            Case 27
                                Me.picCompCard2.Image = card27
                                compCard(1) = 1
                                card(1) = 27
                            Case 28
                                Me.picCompCard2.Image = card28
                                compCard(1) = 2
                                card(1) = 28
                            Case 29
                                Me.picCompCard2.Image = card29
                                compCard(1) = 3
                                card(1) = 29
                            Case 30
                                Me.picCompCard2.Image = card30
                                compCard(1) = 4
                                card(1) = 30
                            Case 31
                                Me.picCompCard2.Image = card31
                                compCard(1) = 5
                                card(1) = 31
                            Case 32
                                Me.picCompCard2.Image = card32
                                compCard(1) = 6
                                card(1) = 32
                            Case 33
                                Me.picCompCard2.Image = card33
                                compCard(1) = 7
                                card(1) = 33
                            Case 34
                                Me.picCompCard2.Image = card34
                                compCard(1) = 8
                                card(1) = 34
                            Case 35
                                Me.picCompCard2.Image = card35
                                compCard(1) = 9
                                card(1) = 35
                            Case 36
                                Me.picCompCard2.Image = card36
                                compCard(1) = 10
                                card(1) = 36
                            Case 37
                                Me.picCompCard2.Image = card37
                                compCard(1) = 11
                                card(1) = 37
                            Case 38
                                Me.picCompCard2.Image = card38
                                compCard(1) = 12
                                card(1) = 38
                            Case 39
                                Me.picCompCard2.Image = card39
                                compCard(1) = 13
                                card(1) = 39
                            Case 40
                                Me.picCompCard2.Image = card40
                                compCard(1) = 1
                                card(1) = 40
                            Case 41
                                Me.picCompCard2.Image = card41
                                compCard(1) = 2
                                card(1) = 41
                            Case 42
                                Me.picCompCard2.Image = card42
                                compCard(1) = 3
                                card(1) = 42
                            Case 43
                                Me.picCompCard2.Image = card43
                                compCard(1) = 4
                                card(1) = 43
                            Case 44
                                Me.picCompCard2.Image = card44
                                compCard(1) = 5
                                card(1) = 44
                            Case 45
                                Me.picCompCard2.Image = card45
                                compCard(1) = 6
                                card(1) = 45
                            Case 46
                                Me.picCompCard2.Image = card46
                                compCard(1) = 7
                                card(1) = 46
                            Case 47
                                Me.picCompCard2.Image = card47
                                compCard(1) = 8
                                card(1) = 47
                            Case 48
                                Me.picCompCard2.Image = card48
                                compCard(1) = 9
                                card(1) = 48
                            Case 49
                                Me.picCompCard2.Image = card49
                                compCard(1) = 10
                                card(1) = 49
                            Case 50
                                Me.picCompCard2.Image = card50
                                compCard(1) = 11
                                card(1) = 50
                            Case 51
                                Me.picCompCard2.Image = card51
                                compCard(1) = 12
                                card(1) = 51
                            Case 52
                                Me.picCompCard2.Image = card52
                                compCard(1) = 13
                                card(1) = 52
                        End Select
                    ElseIf card(10) = card(2) Then
                        Select Case newNum
                            Case 1
                                Me.picCompCard3.Image = card1
                                compCard(2) = 1
                                card(2) = 1
                            Case 2
                                Me.picCompCard3.Image = card2
                                compCard(2) = 2
                                card(2) = 2
                            Case 3
                                Me.picCompCard3.Image = card3
                                compCard(2) = 3
                                card(2) = 3
                            Case 4
                                Me.picCompCard3.Image = card4
                                compCard(2) = 4
                                card(2) = 4
                            Case 5
                                Me.picCompCard3.Image = card5
                                compCard(2) = 5
                                card(2) = 5
                            Case 6
                                Me.picCompCard3.Image = card6
                                compCard(2) = 6
                                card(2) = 6
                            Case 7
                                Me.picCompCard3.Image = card7
                                compCard(2) = 7
                                card(2) = 7
                            Case 8
                                Me.picCompCard3.Image = card8
                                compCard(2) = 8
                                card(2) = 8
                            Case 9
                                Me.picCompCard3.Image = card9
                                compCard(2) = 9
                                card(2) = 9
                            Case 10
                                Me.picCompCard3.Image = card10
                                compCard(2) = 10
                                card(2) = 10
                            Case 11
                                Me.picCompCard3.Image = card11
                                compCard(2) = 11
                                card(2) = 11
                            Case 12
                                Me.picCompCard3.Image = card12
                                compCard(2) = 12
                                card(2) = 12
                            Case 13
                                Me.picCompCard3.Image = card13
                                compCard(2) = 13
                                card(2) = 13
                            Case 14
                                Me.picCompCard3.Image = card14
                                compCard(2) = 1
                                card(2) = 14
                            Case 15
                                Me.picCompCard3.Image = card15
                                compCard(2) = 2
                                card(2) = 15
                            Case 16
                                Me.picCompCard3.Image = card16
                                compCard(2) = 3
                                card(2) = 16
                            Case 17
                                Me.picCompCard3.Image = card17
                                compCard(2) = 4
                                card(2) = 17
                            Case 18
                                Me.picCompCard3.Image = card18
                                compCard(2) = 5
                                card(2) = 18
                            Case 19
                                Me.picCompCard3.Image = card19
                                compCard(2) = 6
                                card(2) = 19
                            Case 20
                                Me.picCompCard3.Image = card20
                                compCard(2) = 7
                                card(2) = 20
                            Case 21
                                Me.picCompCard3.Image = card21
                                compCard(2) = 8
                                card(2) = 21
                            Case 22
                                Me.picCompCard3.Image = card22
                                compCard(2) = 9
                                card(2) = 22
                            Case 23
                                Me.picCompCard3.Image = card23
                                compCard(2) = 10
                                card(2) = 23
                            Case 24
                                Me.picCompCard3.Image = card24
                                compCard(2) = 11
                                card(2) = 24
                            Case 25
                                Me.picCompCard3.Image = card25
                                compCard(2) = 12
                                card(2) = 25
                            Case 26
                                Me.picCompCard3.Image = card26
                                compCard(2) = 13
                                card(2) = 26
                            Case 27
                                Me.picCompCard3.Image = card27
                                compCard(2) = 1
                                card(2) = 27
                            Case 28
                                Me.picCompCard3.Image = card28
                                compCard(2) = 2
                                card(2) = 28
                            Case 29
                                Me.picCompCard3.Image = card29
                                compCard(2) = 3
                                card(2) = 29
                            Case 30
                                Me.picCompCard3.Image = card30
                                compCard(2) = 4
                                card(2) = 30
                            Case 31
                                Me.picCompCard3.Image = card31
                                compCard(2) = 5
                                card(2) = 31
                            Case 32
                                Me.picCompCard3.Image = card32
                                compCard(2) = 6
                                card(2) = 32
                            Case 33
                                Me.picCompCard3.Image = card33
                                compCard(2) = 7
                                card(2) = 33
                            Case 34
                                Me.picCompCard3.Image = card34
                                compCard(2) = 8
                                card(2) = 34
                            Case 35
                                Me.picCompCard3.Image = card35
                                compCard(2) = 9
                                card(2) = 35
                            Case 36
                                Me.picCompCard3.Image = card36
                                compCard(2) = 10
                                card(2) = 36
                            Case 37
                                Me.picCompCard3.Image = card37
                                compCard(2) = 11
                                card(2) = 37
                            Case 38
                                Me.picCompCard3.Image = card38
                                compCard(2) = 12
                                card(2) = 38
                            Case 39
                                Me.picCompCard3.Image = card39
                                compCard(2) = 13
                                card(2) = 39
                            Case 40
                                Me.picCompCard3.Image = card40
                                compCard(2) = 1
                                card(2) = 40
                            Case 41
                                Me.picCompCard3.Image = card41
                                compCard(2) = 2
                                card(2) = 41
                            Case 42
                                Me.picCompCard3.Image = card42
                                compCard(2) = 3
                                card(2) = 42
                            Case 43
                                Me.picCompCard3.Image = card43
                                compCard(2) = 4
                                card(2) = 43
                            Case 44
                                Me.picCompCard3.Image = card44
                                compCard(2) = 5
                                card(2) = 44
                            Case 45
                                Me.picCompCard3.Image = card45
                                compCard(2) = 6
                                card(2) = 45
                            Case 46
                                Me.picCompCard3.Image = card46
                                compCard(2) = 7
                                card(2) = 46
                            Case 47
                                Me.picCompCard3.Image = card47
                                compCard(2) = 8
                                card(2) = 47
                            Case 48
                                Me.picCompCard3.Image = card48
                                compCard(2) = 9
                                card(2) = 48
                            Case 49
                                Me.picCompCard3.Image = card49
                                compCard(2) = 10
                                card(2) = 49
                            Case 50
                                Me.picCompCard3.Image = card50
                                compCard(2) = 11
                                card(2) = 50
                            Case 51
                                Me.picCompCard3.Image = card51
                                compCard(2) = 12
                                card(2) = 51
                            Case 52
                                Me.picCompCard3.Image = card52
                                compCard(2) = 13
                                card(2) = 52
                        End Select
                    ElseIf card(10) = card(3) Then
                        Select Case newNum
                            Case 1
                                Me.picCompCard4.Image = card1
                                compCard(3) = 1
                                card(3) = 1
                            Case 2
                                Me.picCompCard4.Image = card2
                                compCard(3) = 2
                                card(3) = 2
                            Case 3
                                Me.picCompCard4.Image = card3
                                compCard(3) = 3
                                card(3) = 3
                            Case 4
                                Me.picCompCard4.Image = card4
                                compCard(3) = 4
                                card(3) = 4
                            Case 5
                                Me.picCompCard4.Image = card5
                                compCard(3) = 5
                                card(3) = 5
                            Case 6
                                Me.picCompCard4.Image = card6
                                compCard(3) = 6
                                card(3) = 6
                            Case 7
                                Me.picCompCard4.Image = card7
                                compCard(3) = 7
                                card(3) = 7
                            Case 8
                                Me.picCompCard4.Image = card8
                                compCard(3) = 8
                                card(3) = 8
                            Case 9
                                Me.picCompCard4.Image = card9
                                compCard(3) = 9
                                card(3) = 9
                            Case 10
                                Me.picCompCard4.Image = card10
                                compCard(3) = 10
                                card(3) = 10
                            Case 11
                                Me.picCompCard4.Image = card11
                                compCard(3) = 11
                                card(3) = 11
                            Case 12
                                Me.picCompCard4.Image = card12
                                compCard(3) = 12
                                card(3) = 12
                            Case 13
                                Me.picCompCard4.Image = card13
                                compCard(3) = 13
                                card(3) = 13
                            Case 14
                                Me.picCompCard4.Image = card14
                                compCard(3) = 1
                                card(3) = 14
                            Case 15
                                Me.picCompCard4.Image = card15
                                compCard(3) = 2
                                card(3) = 15
                            Case 16
                                Me.picCompCard4.Image = card16
                                compCard(3) = 3
                                card(3) = 16
                            Case 17
                                Me.picCompCard4.Image = card17
                                compCard(3) = 4
                                card(3) = 17
                            Case 18
                                Me.picCompCard4.Image = card18
                                compCard(3) = 5
                                card(3) = 18
                            Case 19
                                Me.picCompCard4.Image = card19
                                compCard(3) = 6
                                card(3) = 19
                            Case 20
                                Me.picCompCard4.Image = card20
                                compCard(3) = 7
                                card(3) = 20
                            Case 21
                                Me.picCompCard4.Image = card21
                                compCard(3) = 8
                                card(3) = 21
                            Case 22
                                Me.picCompCard4.Image = card22
                                compCard(3) = 9
                                card(3) = 22
                            Case 23
                                Me.picCompCard4.Image = card23
                                compCard(3) = 10
                                card(3) = 23
                            Case 24
                                Me.picCompCard4.Image = card24
                                compCard(3) = 11
                                card(3) = 24
                            Case 25
                                Me.picCompCard4.Image = card25
                                compCard(3) = 12
                                card(3) = 25
                            Case 26
                                Me.picCompCard4.Image = card26
                                compCard(3) = 13
                                card(3) = 26
                            Case 27
                                Me.picCompCard4.Image = card27
                                compCard(3) = 1
                                card(3) = 27
                            Case 28
                                Me.picCompCard4.Image = card28
                                compCard(3) = 2
                                card(3) = 28
                            Case 29
                                Me.picCompCard4.Image = card29
                                compCard(3) = 3
                                card(3) = 29
                            Case 30
                                Me.picCompCard4.Image = card30
                                compCard(3) = 4
                                card(3) = 30
                            Case 31
                                Me.picCompCard4.Image = card31
                                compCard(3) = 5
                                card(3) = 31
                            Case 32
                                Me.picCompCard4.Image = card32
                                compCard(3) = 6
                                card(3) = 32
                            Case 33
                                Me.picCompCard4.Image = card33
                                compCard(3) = 7
                                card(3) = 33
                            Case 34
                                Me.picCompCard4.Image = card34
                                compCard(3) = 8
                                card(3) = 34
                            Case 35
                                Me.picCompCard4.Image = card35
                                compCard(3) = 9
                                card(3) = 35
                            Case 36
                                Me.picCompCard4.Image = card36
                                compCard(3) = 10
                                card(3) = 36
                            Case 37
                                Me.picCompCard4.Image = card37
                                compCard(3) = 11
                                card(3) = 37
                            Case 38
                                Me.picCompCard4.Image = card38
                                compCard(3) = 12
                                card(3) = 38
                            Case 39
                                Me.picCompCard4.Image = card39
                                compCard(3) = 13
                                card(3) = 39
                            Case 40
                                Me.picCompCard4.Image = card40
                                compCard(3) = 1
                                card(3) = 40
                            Case 41
                                Me.picCompCard4.Image = card41
                                compCard(3) = 2
                                card(3) = 41
                            Case 42
                                Me.picCompCard4.Image = card42
                                compCard(3) = 3
                                card(3) = 42
                            Case 43
                                Me.picCompCard4.Image = card43
                                compCard(3) = 4
                                card(3) = 43
                            Case 44
                                Me.picCompCard4.Image = card44
                                compCard(3) = 5
                                card(3) = 44
                            Case 45
                                Me.picCompCard4.Image = card45
                                compCard(3) = 6
                                card(3) = 45
                            Case 46
                                Me.picCompCard4.Image = card46
                                compCard(3) = 7
                                card(3) = 46
                            Case 47
                                Me.picCompCard4.Image = card47
                                compCard(3) = 8
                                card(3) = 47
                            Case 48
                                Me.picCompCard4.Image = card48
                                compCard(3) = 9
                                card(3) = 48
                            Case 49
                                Me.picCompCard4.Image = card49
                                compCard(3) = 10
                                card(3) = 49
                            Case 50
                                Me.picCompCard4.Image = card50
                                compCard(3) = 11
                                card(3) = 50
                            Case 51
                                Me.picCompCard4.Image = card51
                                compCard(3) = 12
                                card(3) = 51
                            Case 52
                                Me.picCompCard4.Image = card52
                                compCard(3) = 13
                                card(3) = 52
                        End Select
                    ElseIf card(10) = card(4) Then
                        Select Case newNum
                            Case 1
                                Me.picCompCard5.Image = card1
                                compCard(4) = 1
                                card(4) = 1
                            Case 2
                                Me.picCompCard5.Image = card2
                                compCard(4) = 2
                                card(4) = 2
                            Case 3
                                Me.picCompCard5.Image = card3
                                compCard(4) = 3
                                card(4) = 3
                            Case 4
                                Me.picCompCard5.Image = card4
                                compCard(4) = 4
                                card(4) = 4
                            Case 5
                                Me.picCompCard5.Image = card5
                                compCard(4) = 5
                                card(4) = 5
                            Case 6
                                Me.picCompCard5.Image = card6
                                compCard(4) = 6
                                card(4) = 6
                            Case 7
                                Me.picCompCard5.Image = card7
                                compCard(4) = 7
                                card(4) = 7
                            Case 8
                                Me.picCompCard5.Image = card8
                                compCard(4) = 8
                                card(4) = 8
                            Case 9
                                Me.picCompCard5.Image = card9
                                compCard(4) = 9
                                card(4) = 9
                            Case 10
                                Me.picCompCard5.Image = card10
                                compCard(4) = 10
                                card(4) = 10
                            Case 11
                                Me.picCompCard5.Image = card11
                                compCard(4) = 11
                                card(4) = 11
                            Case 12
                                Me.picCompCard5.Image = card12
                                compCard(4) = 12
                                card(4) = 12
                            Case 13
                                Me.picCompCard5.Image = card13
                                compCard(4) = 13
                                card(4) = 13
                            Case 14
                                Me.picCompCard5.Image = card14
                                compCard(4) = 1
                                card(4) = 14
                            Case 15
                                Me.picCompCard5.Image = card15
                                compCard(4) = 2
                                card(4) = 15
                            Case 16
                                Me.picCompCard5.Image = card16
                                compCard(4) = 3
                                card(4) = 16
                            Case 17
                                Me.picCompCard5.Image = card17
                                compCard(4) = 4
                                card(4) = 17
                            Case 18
                                Me.picCompCard5.Image = card18
                                compCard(4) = 5
                                card(4) = 18
                            Case 19
                                Me.picCompCard5.Image = card19
                                compCard(4) = 6
                                card(4) = 19
                            Case 20
                                Me.picCompCard5.Image = card20
                                compCard(4) = 7
                                card(4) = 20
                            Case 21
                                Me.picCompCard5.Image = card21
                                compCard(4) = 8
                                card(4) = 21
                            Case 22
                                Me.picCompCard5.Image = card22
                                compCard(4) = 9
                                card(4) = 22
                            Case 23
                                Me.picCompCard5.Image = card23
                                compCard(4) = 10
                                card(4) = 23
                            Case 24
                                Me.picCompCard5.Image = card24
                                compCard(4) = 11
                                card(4) = 24
                            Case 25
                                Me.picCompCard5.Image = card25
                                compCard(4) = 12
                                card(4) = 25
                            Case 26
                                Me.picCompCard5.Image = card26
                                compCard(4) = 13
                                card(4) = 26
                            Case 27
                                Me.picCompCard5.Image = card27
                                compCard(4) = 1
                                card(4) = 27
                            Case 28
                                Me.picCompCard5.Image = card28
                                compCard(4) = 2
                                card(4) = 28
                            Case 29
                                Me.picCompCard5.Image = card29
                                compCard(4) = 3
                                card(4) = 29
                            Case 30
                                Me.picCompCard5.Image = card30
                                compCard(4) = 4
                                card(4) = 30
                            Case 31
                                Me.picCompCard5.Image = card31
                                compCard(4) = 5
                                card(4) = 31
                            Case 32
                                Me.picCompCard5.Image = card32
                                compCard(4) = 6
                                card(4) = 32
                            Case 33
                                Me.picCompCard5.Image = card33
                                compCard(4) = 7
                                card(4) = 33
                            Case 34
                                Me.picCompCard5.Image = card34
                                compCard(4) = 8
                                card(4) = 34
                            Case 35
                                Me.picCompCard5.Image = card35
                                compCard(4) = 9
                                card(4) = 35
                            Case 36
                                Me.picCompCard5.Image = card36
                                compCard(4) = 10
                                card(4) = 36
                            Case 37
                                Me.picCompCard5.Image = card37
                                compCard(4) = 11
                                card(4) = 37
                            Case 38
                                Me.picCompCard5.Image = card38
                                compCard(4) = 12
                                card(4) = 38
                            Case 39
                                Me.picCompCard5.Image = card39
                                compCard(4) = 13
                                card(4) = 39
                            Case 40
                                Me.picCompCard5.Image = card40
                                compCard(4) = 1
                                card(4) = 40
                            Case 41
                                Me.picCompCard5.Image = card41
                                compCard(4) = 2
                                card(4) = 41
                            Case 42
                                Me.picCompCard5.Image = card42
                                compCard(4) = 3
                                card(4) = 42
                            Case 43
                                Me.picCompCard5.Image = card43
                                compCard(4) = 4
                                card(4) = 43
                            Case 44
                                Me.picCompCard5.Image = card44
                                compCard(4) = 5
                                card(4) = 44
                            Case 45
                                Me.picCompCard5.Image = card45
                                compCard(4) = 6
                                card(4) = 45
                            Case 46
                                Me.picCompCard5.Image = card46
                                compCard(4) = 7
                                card(4) = 46
                            Case 47
                                Me.picCompCard5.Image = card47
                                compCard(4) = 8
                                card(4) = 47
                            Case 48
                                Me.picCompCard5.Image = card48
                                compCard(4) = 9
                                card(4) = 48
                            Case 49
                                Me.picCompCard5.Image = card49
                                compCard(4) = 10
                                card(4) = 49
                            Case 50
                                Me.picCompCard5.Image = card50
                                compCard(4) = 11
                                card(4) = 50
                            Case 51
                                Me.picCompCard5.Image = card51
                                compCard(4) = 12
                                card(4) = 51
                            Case 52
                                Me.picCompCard5.Image = card52
                                compCard(4) = 13
                                card(4) = 52
                        End Select
                    End If

                    Used.Add(newNum)

                End If
            Loop While OK = False
        End If

        If card(11) = card(0) Or card(11) = card(1) Or card(11) = card(2) Or card(11) = card(3) Or card(11) = card(4) Then
            Do                                          'generate a new unique random number
                Randomize()
                newNum = Int(max * Rnd() + min)

                If Used.Contains(newNum) Then
                    OK = False
                Else
                    OK = True

                    If card(11) = card(0) Then
                        Select Case newNum
                            Case 1
                                Me.picCompCard1.Image = card1
                                compCard(0) = 1
                                card(0) = 1
                            Case 2
                                Me.picCompCard1.Image = card2
                                compCard(0) = 2
                                card(0) = 2
                            Case 3
                                Me.picCompCard1.Image = card3
                                compCard(0) = 3
                                card(0) = 3
                            Case 4
                                Me.picCompCard1.Image = card4
                                compCard(0) = 4
                                card(0) = 4
                            Case 5
                                Me.picCompCard1.Image = card5
                                compCard(0) = 5
                                card(0) = 5
                            Case 6
                                Me.picCompCard1.Image = card6
                                compCard(0) = 6
                                card(0) = 6
                            Case 7
                                Me.picCompCard1.Image = card7
                                compCard(0) = 7
                                card(0) = 7
                            Case 8
                                Me.picCompCard1.Image = card8
                                compCard(0) = 8
                                card(0) = 8
                            Case 9
                                Me.picCompCard1.Image = card9
                                compCard(0) = 9
                                card(0) = 9
                            Case 10
                                Me.picCompCard1.Image = card10
                                compCard(0) = 10
                                card(0) = 10
                            Case 11
                                Me.picCompCard1.Image = card11
                                compCard(0) = 11
                                card(0) = 11
                            Case 12
                                Me.picCompCard1.Image = card12
                                compCard(0) = 12
                                card(0) = 12
                            Case 13
                                Me.picCompCard1.Image = card13
                                compCard(0) = 13
                                card(0) = 13
                            Case 14
                                Me.picCompCard1.Image = card14
                                compCard(0) = 1
                                card(0) = 14
                            Case 15
                                Me.picCompCard1.Image = card15
                                compCard(0) = 2
                                card(0) = 15
                            Case 16
                                Me.picCompCard1.Image = card16
                                compCard(0) = 3
                                card(0) = 16
                            Case 17
                                Me.picCompCard1.Image = card17
                                compCard(0) = 4
                                card(0) = 17
                            Case 18
                                Me.picCompCard1.Image = card18
                                compCard(0) = 5
                                card(0) = 18
                            Case 19
                                Me.picCompCard1.Image = card19
                                compCard(0) = 6
                                card(0) = 19
                            Case 20
                                Me.picCompCard1.Image = card20
                                compCard(0) = 7
                                card(0) = 20
                            Case 21
                                Me.picCompCard1.Image = card21
                                compCard(0) = 8
                                card(0) = 21
                            Case 22
                                Me.picCompCard1.Image = card22
                                compCard(0) = 9
                                card(0) = 22
                            Case 23
                                Me.picCompCard1.Image = card23
                                compCard(0) = 10
                                card(0) = 23
                            Case 24
                                Me.picCompCard1.Image = card24
                                compCard(0) = 11
                                card(0) = 24
                            Case 25
                                Me.picCompCard1.Image = card25
                                compCard(0) = 12
                                card(0) = 25
                            Case 26
                                Me.picCompCard1.Image = card26
                                compCard(0) = 13
                                card(0) = 26
                            Case 27
                                Me.picCompCard1.Image = card27
                                compCard(0) = 1
                                card(0) = 27
                            Case 28
                                Me.picCompCard1.Image = card28
                                compCard(0) = 2
                                card(0) = 28
                            Case 29
                                Me.picCompCard1.Image = card29
                                compCard(0) = 3
                                card(0) = 29
                            Case 30
                                Me.picCompCard1.Image = card30
                                compCard(0) = 4
                                card(0) = 30
                            Case 31
                                Me.picCompCard1.Image = card31
                                compCard(0) = 5
                                card(0) = 31
                            Case 32
                                Me.picCompCard1.Image = card32
                                compCard(0) = 6
                                card(0) = 32
                            Case 33
                                Me.picCompCard1.Image = card33
                                compCard(0) = 7
                                card(0) = 33
                            Case 34
                                Me.picCompCard1.Image = card34
                                compCard(0) = 8
                                card(0) = 34
                            Case 35
                                Me.picCompCard1.Image = card35
                                compCard(0) = 9
                                card(0) = 35
                            Case 36
                                Me.picCompCard1.Image = card36
                                compCard(0) = 10
                                card(0) = 36
                            Case 37
                                Me.picCompCard1.Image = card37
                                compCard(0) = 11
                                card(0) = 37
                            Case 38
                                Me.picCompCard1.Image = card38
                                compCard(0) = 12
                                card(0) = 38
                            Case 39
                                Me.picCompCard1.Image = card39
                                compCard(0) = 13
                                card(0) = 39
                            Case 40
                                Me.picCompCard1.Image = card40
                                compCard(0) = 1
                                card(0) = 40
                            Case 41
                                Me.picCompCard1.Image = card41
                                compCard(0) = 2
                                card(0) = 41
                            Case 42
                                Me.picCompCard1.Image = card42
                                compCard(0) = 3
                                card(0) = 42
                            Case 43
                                Me.picCompCard1.Image = card43
                                compCard(0) = 4
                                card(0) = 43
                            Case 44
                                Me.picCompCard1.Image = card44
                                compCard(0) = 5
                                card(0) = 44
                            Case 45
                                Me.picCompCard1.Image = card45
                                compCard(0) = 6
                                card(0) = 45
                            Case 46
                                Me.picCompCard1.Image = card46
                                compCard(0) = 7
                                card(0) = 46
                            Case 47
                                Me.picCompCard1.Image = card47
                                compCard(0) = 8
                                card(0) = 47
                            Case 48
                                Me.picCompCard1.Image = card48
                                compCard(0) = 9
                                card(0) = 48
                            Case 49
                                Me.picCompCard1.Image = card49
                                compCard(0) = 10
                                card(0) = 49
                            Case 50
                                Me.picCompCard1.Image = card50
                                compCard(0) = 11
                                card(0) = 50
                            Case 51
                                Me.picCompCard1.Image = card51
                                compCard(0) = 12
                                card(0) = 51
                            Case 52
                                Me.picCompCard1.Image = card52
                                compCard(0) = 13
                                card(0) = 52
                        End Select
                    ElseIf card(11) = card(1) Then
                        Select Case newNum
                            Case 1
                                Me.picCompCard2.Image = card1
                                compCard(1) = 1
                                card(1) = 1
                            Case 2
                                Me.picCompCard2.Image = card2
                                compCard(1) = 2
                                card(1) = 2
                            Case 3
                                Me.picCompCard2.Image = card3
                                compCard(1) = 3
                                card(1) = 3
                            Case 4
                                Me.picCompCard2.Image = card4
                                compCard(1) = 4
                                card(1) = 4
                            Case 5
                                Me.picCompCard2.Image = card5
                                compCard(1) = 5
                                card(1) = 5
                            Case 6
                                Me.picCompCard2.Image = card6
                                compCard(1) = 6
                                card(1) = 6
                            Case 7
                                Me.picCompCard2.Image = card7
                                compCard(1) = 7
                                card(1) = 7
                            Case 8
                                Me.picCompCard2.Image = card8
                                compCard(1) = 8
                                card(1) = 8
                            Case 9
                                Me.picCompCard2.Image = card9
                                compCard(1) = 9
                                card(1) = 9
                            Case 10
                                Me.picCompCard2.Image = card10
                                compCard(1) = 10
                                card(1) = 10
                            Case 11
                                Me.picCompCard2.Image = card11
                                compCard(1) = 11
                                card(1) = 11
                            Case 12
                                Me.picCompCard2.Image = card12
                                compCard(1) = 12
                                card(1) = 12
                            Case 13
                                Me.picCompCard2.Image = card13
                                compCard(1) = 13
                                card(1) = 13
                            Case 14
                                Me.picCompCard2.Image = card14
                                compCard(1) = 1
                                card(1) = 14
                            Case 15
                                Me.picCompCard2.Image = card15
                                compCard(1) = 2
                                card(1) = 15
                            Case 16
                                Me.picCompCard2.Image = card16
                                compCard(1) = 3
                                card(1) = 16
                            Case 17
                                Me.picCompCard2.Image = card17
                                compCard(1) = 4
                                card(1) = 17
                            Case 18
                                Me.picCompCard2.Image = card18
                                compCard(1) = 5
                                card(1) = 18
                            Case 19
                                Me.picCompCard2.Image = card19
                                compCard(1) = 6
                                card(1) = 19
                            Case 20
                                Me.picCompCard2.Image = card20
                                compCard(1) = 7
                                card(1) = 20
                            Case 21
                                Me.picCompCard2.Image = card21
                                compCard(1) = 8
                                card(1) = 21
                            Case 22
                                Me.picCompCard2.Image = card22
                                compCard(1) = 9
                                card(1) = 22
                            Case 23
                                Me.picCompCard2.Image = card23
                                compCard(1) = 10
                                card(1) = 23
                            Case 24
                                Me.picCompCard2.Image = card24
                                compCard(1) = 11
                                card(1) = 24
                            Case 25
                                Me.picCompCard2.Image = card25
                                compCard(1) = 12
                                card(1) = 25
                            Case 26
                                Me.picCompCard2.Image = card26
                                compCard(1) = 13
                                card(1) = 26
                            Case 27
                                Me.picCompCard2.Image = card27
                                compCard(1) = 1
                                card(1) = 27
                            Case 28
                                Me.picCompCard2.Image = card28
                                compCard(1) = 2
                                card(1) = 28
                            Case 29
                                Me.picCompCard2.Image = card29
                                compCard(1) = 3
                                card(1) = 29
                            Case 30
                                Me.picCompCard2.Image = card30
                                compCard(1) = 4
                                card(1) = 30
                            Case 31
                                Me.picCompCard2.Image = card31
                                compCard(1) = 5
                                card(1) = 31
                            Case 32
                                Me.picCompCard2.Image = card32
                                compCard(1) = 6
                                card(1) = 32
                            Case 33
                                Me.picCompCard2.Image = card33
                                compCard(1) = 7
                                card(1) = 33
                            Case 34
                                Me.picCompCard2.Image = card34
                                compCard(1) = 8
                                card(1) = 34
                            Case 35
                                Me.picCompCard2.Image = card35
                                compCard(1) = 9
                                card(1) = 35
                            Case 36
                                Me.picCompCard2.Image = card36
                                compCard(1) = 10
                                card(1) = 36
                            Case 37
                                Me.picCompCard2.Image = card37
                                compCard(1) = 11
                                card(1) = 37
                            Case 38
                                Me.picCompCard2.Image = card38
                                compCard(1) = 12
                                card(1) = 38
                            Case 39
                                Me.picCompCard2.Image = card39
                                compCard(1) = 13
                                card(1) = 39
                            Case 40
                                Me.picCompCard2.Image = card40
                                compCard(1) = 1
                                card(1) = 40
                            Case 41
                                Me.picCompCard2.Image = card41
                                compCard(1) = 2
                                card(1) = 41
                            Case 42
                                Me.picCompCard2.Image = card42
                                compCard(1) = 3
                                card(1) = 42
                            Case 43
                                Me.picCompCard2.Image = card43
                                compCard(1) = 4
                                card(1) = 43
                            Case 44
                                Me.picCompCard2.Image = card44
                                compCard(1) = 5
                                card(1) = 44
                            Case 45
                                Me.picCompCard2.Image = card45
                                compCard(1) = 6
                                card(1) = 45
                            Case 46
                                Me.picCompCard2.Image = card46
                                compCard(1) = 7
                                card(1) = 46
                            Case 47
                                Me.picCompCard2.Image = card47
                                compCard(1) = 8
                                card(1) = 47
                            Case 48
                                Me.picCompCard2.Image = card48
                                compCard(1) = 9
                                card(1) = 48
                            Case 49
                                Me.picCompCard2.Image = card49
                                compCard(1) = 10
                                card(1) = 49
                            Case 50
                                Me.picCompCard2.Image = card50
                                compCard(1) = 11
                                card(1) = 50
                            Case 51
                                Me.picCompCard2.Image = card51
                                compCard(1) = 12
                                card(1) = 51
                            Case 52
                                Me.picCompCard2.Image = card52
                                compCard(1) = 13
                                card(1) = 52
                        End Select
                    ElseIf card(11) = card(2) Then
                        Select Case newNum
                            Case 1
                                Me.picCompCard3.Image = card1
                                compCard(2) = 1
                                card(2) = 1
                            Case 2
                                Me.picCompCard3.Image = card2
                                compCard(2) = 2
                                card(2) = 2
                            Case 3
                                Me.picCompCard3.Image = card3
                                compCard(2) = 3
                                card(2) = 3
                            Case 4
                                Me.picCompCard3.Image = card4
                                compCard(2) = 4
                                card(2) = 4
                            Case 5
                                Me.picCompCard3.Image = card5
                                compCard(2) = 5
                                card(2) = 5
                            Case 6
                                Me.picCompCard3.Image = card6
                                compCard(2) = 6
                                card(2) = 6
                            Case 7
                                Me.picCompCard3.Image = card7
                                compCard(2) = 7
                                card(2) = 7
                            Case 8
                                Me.picCompCard3.Image = card8
                                compCard(2) = 8
                                card(2) = 8
                            Case 9
                                Me.picCompCard3.Image = card9
                                compCard(2) = 9
                                card(2) = 9
                            Case 10
                                Me.picCompCard3.Image = card10
                                compCard(2) = 10
                                card(2) = 10
                            Case 11
                                Me.picCompCard3.Image = card11
                                compCard(2) = 11
                                card(2) = 11
                            Case 12
                                Me.picCompCard3.Image = card12
                                compCard(2) = 12
                                card(2) = 12
                            Case 13
                                Me.picCompCard3.Image = card13
                                compCard(2) = 13
                                card(2) = 13
                            Case 14
                                Me.picCompCard3.Image = card14
                                compCard(2) = 1
                                card(2) = 14
                            Case 15
                                Me.picCompCard3.Image = card15
                                compCard(2) = 2
                                card(2) = 15
                            Case 16
                                Me.picCompCard3.Image = card16
                                compCard(2) = 3
                                card(2) = 16
                            Case 17
                                Me.picCompCard3.Image = card17
                                compCard(2) = 4
                                card(2) = 17
                            Case 18
                                Me.picCompCard3.Image = card18
                                compCard(2) = 5
                                card(2) = 18
                            Case 19
                                Me.picCompCard3.Image = card19
                                compCard(2) = 6
                                card(2) = 19
                            Case 20
                                Me.picCompCard3.Image = card20
                                compCard(2) = 7
                                card(2) = 20
                            Case 21
                                Me.picCompCard3.Image = card21
                                compCard(2) = 8
                                card(2) = 21
                            Case 22
                                Me.picCompCard3.Image = card22
                                compCard(2) = 9
                                card(2) = 22
                            Case 23
                                Me.picCompCard3.Image = card23
                                compCard(2) = 10
                                card(2) = 23
                            Case 24
                                Me.picCompCard3.Image = card24
                                compCard(2) = 11
                                card(2) = 24
                            Case 25
                                Me.picCompCard3.Image = card25
                                compCard(2) = 12
                                card(2) = 25
                            Case 26
                                Me.picCompCard3.Image = card26
                                compCard(2) = 13
                                card(2) = 26
                            Case 27
                                Me.picCompCard3.Image = card27
                                compCard(2) = 1
                                card(2) = 27
                            Case 28
                                Me.picCompCard3.Image = card28
                                compCard(2) = 2
                                card(2) = 28
                            Case 29
                                Me.picCompCard3.Image = card29
                                compCard(2) = 3
                                card(2) = 29
                            Case 30
                                Me.picCompCard3.Image = card30
                                compCard(2) = 4
                                card(2) = 30
                            Case 31
                                Me.picCompCard3.Image = card31
                                compCard(2) = 5
                                card(2) = 31
                            Case 32
                                Me.picCompCard3.Image = card32
                                compCard(2) = 6
                                card(2) = 32
                            Case 33
                                Me.picCompCard3.Image = card33
                                compCard(2) = 7
                                card(2) = 33
                            Case 34
                                Me.picCompCard3.Image = card34
                                compCard(2) = 8
                                card(2) = 34
                            Case 35
                                Me.picCompCard3.Image = card35
                                compCard(2) = 9
                                card(2) = 35
                            Case 36
                                Me.picCompCard3.Image = card36
                                compCard(2) = 10
                                card(2) = 36
                            Case 37
                                Me.picCompCard3.Image = card37
                                compCard(2) = 11
                                card(2) = 37
                            Case 38
                                Me.picCompCard3.Image = card38
                                compCard(2) = 12
                                card(2) = 38
                            Case 39
                                Me.picCompCard3.Image = card39
                                compCard(2) = 13
                                card(2) = 39
                            Case 40
                                Me.picCompCard3.Image = card40
                                compCard(2) = 1
                                card(2) = 40
                            Case 41
                                Me.picCompCard3.Image = card41
                                compCard(2) = 2
                                card(2) = 41
                            Case 42
                                Me.picCompCard3.Image = card42
                                compCard(2) = 3
                                card(2) = 42
                            Case 43
                                Me.picCompCard3.Image = card43
                                compCard(2) = 4
                                card(2) = 43
                            Case 44
                                Me.picCompCard3.Image = card44
                                compCard(2) = 5
                                card(2) = 44
                            Case 45
                                Me.picCompCard3.Image = card45
                                compCard(2) = 6
                                card(2) = 45
                            Case 46
                                Me.picCompCard3.Image = card46
                                compCard(2) = 7
                                card(2) = 46
                            Case 47
                                Me.picCompCard3.Image = card47
                                compCard(2) = 8
                                card(2) = 47
                            Case 48
                                Me.picCompCard3.Image = card48
                                compCard(2) = 9
                                card(2) = 48
                            Case 49
                                Me.picCompCard3.Image = card49
                                compCard(2) = 10
                                card(2) = 49
                            Case 50
                                Me.picCompCard3.Image = card50
                                compCard(2) = 11
                                card(2) = 50
                            Case 51
                                Me.picCompCard3.Image = card51
                                compCard(2) = 12
                                card(2) = 51
                            Case 52
                                Me.picCompCard3.Image = card52
                                compCard(2) = 13
                                card(2) = 52
                        End Select
                    ElseIf card(11) = card(3) Then
                        Select Case newNum
                            Case 1
                                Me.picCompCard4.Image = card1
                                compCard(3) = 1
                                card(3) = 1
                            Case 2
                                Me.picCompCard4.Image = card2
                                compCard(3) = 2
                                card(3) = 2
                            Case 3
                                Me.picCompCard4.Image = card3
                                compCard(3) = 3
                                card(3) = 3
                            Case 4
                                Me.picCompCard4.Image = card4
                                compCard(3) = 4
                                card(3) = 4
                            Case 5
                                Me.picCompCard4.Image = card5
                                compCard(3) = 5
                                card(3) = 5
                            Case 6
                                Me.picCompCard4.Image = card6
                                compCard(3) = 6
                                card(3) = 6
                            Case 7
                                Me.picCompCard4.Image = card7
                                compCard(3) = 7
                                card(3) = 7
                            Case 8
                                Me.picCompCard4.Image = card8
                                compCard(3) = 8
                                card(3) = 8
                            Case 9
                                Me.picCompCard4.Image = card9
                                compCard(3) = 9
                                card(3) = 9
                            Case 10
                                Me.picCompCard4.Image = card10
                                compCard(3) = 10
                                card(3) = 10
                            Case 11
                                Me.picCompCard4.Image = card11
                                compCard(3) = 11
                                card(3) = 11
                            Case 12
                                Me.picCompCard4.Image = card12
                                compCard(3) = 12
                                card(3) = 12
                            Case 13
                                Me.picCompCard4.Image = card13
                                compCard(3) = 13
                                card(3) = 13
                            Case 14
                                Me.picCompCard4.Image = card14
                                compCard(3) = 1
                                card(3) = 14
                            Case 15
                                Me.picCompCard4.Image = card15
                                compCard(3) = 2
                                card(3) = 15
                            Case 16
                                Me.picCompCard4.Image = card16
                                compCard(3) = 3
                                card(3) = 16
                            Case 17
                                Me.picCompCard4.Image = card17
                                compCard(3) = 4
                                card(3) = 17
                            Case 18
                                Me.picCompCard4.Image = card18
                                compCard(3) = 5
                                card(3) = 18
                            Case 19
                                Me.picCompCard4.Image = card19
                                compCard(3) = 6
                                card(3) = 19
                            Case 20
                                Me.picCompCard4.Image = card20
                                compCard(3) = 7
                                card(3) = 20
                            Case 21
                                Me.picCompCard4.Image = card21
                                compCard(3) = 8
                                card(3) = 21
                            Case 22
                                Me.picCompCard4.Image = card22
                                compCard(3) = 9
                                card(3) = 22
                            Case 23
                                Me.picCompCard4.Image = card23
                                compCard(3) = 10
                                card(3) = 23
                            Case 24
                                Me.picCompCard4.Image = card24
                                compCard(3) = 11
                                card(3) = 24
                            Case 25
                                Me.picCompCard4.Image = card25
                                compCard(3) = 12
                                card(3) = 25
                            Case 26
                                Me.picCompCard4.Image = card26
                                compCard(3) = 13
                                card(3) = 26
                            Case 27
                                Me.picCompCard4.Image = card27
                                compCard(3) = 1
                                card(3) = 27
                            Case 28
                                Me.picCompCard4.Image = card28
                                compCard(3) = 2
                                card(3) = 28
                            Case 29
                                Me.picCompCard4.Image = card29
                                compCard(3) = 3
                                card(3) = 29
                            Case 30
                                Me.picCompCard4.Image = card30
                                compCard(3) = 4
                                card(3) = 30
                            Case 31
                                Me.picCompCard4.Image = card31
                                compCard(3) = 5
                                card(3) = 31
                            Case 32
                                Me.picCompCard4.Image = card32
                                compCard(3) = 6
                                card(3) = 32
                            Case 33
                                Me.picCompCard4.Image = card33
                                compCard(3) = 7
                                card(3) = 33
                            Case 34
                                Me.picCompCard4.Image = card34
                                compCard(3) = 8
                                card(3) = 34
                            Case 35
                                Me.picCompCard4.Image = card35
                                compCard(3) = 9
                                card(3) = 35
                            Case 36
                                Me.picCompCard4.Image = card36
                                compCard(3) = 10
                                card(3) = 36
                            Case 37
                                Me.picCompCard4.Image = card37
                                compCard(3) = 11
                                card(3) = 37
                            Case 38
                                Me.picCompCard4.Image = card38
                                compCard(3) = 12
                                card(3) = 38
                            Case 39
                                Me.picCompCard4.Image = card39
                                compCard(3) = 13
                                card(3) = 39
                            Case 40
                                Me.picCompCard4.Image = card40
                                compCard(3) = 1
                                card(3) = 40
                            Case 41
                                Me.picCompCard4.Image = card41
                                compCard(3) = 2
                                card(3) = 41
                            Case 42
                                Me.picCompCard4.Image = card42
                                compCard(3) = 3
                                card(3) = 42
                            Case 43
                                Me.picCompCard4.Image = card43
                                compCard(3) = 4
                                card(3) = 43
                            Case 44
                                Me.picCompCard4.Image = card44
                                compCard(3) = 5
                                card(3) = 44
                            Case 45
                                Me.picCompCard4.Image = card45
                                compCard(3) = 6
                                card(3) = 45
                            Case 46
                                Me.picCompCard4.Image = card46
                                compCard(3) = 7
                                card(3) = 46
                            Case 47
                                Me.picCompCard4.Image = card47
                                compCard(3) = 8
                                card(3) = 47
                            Case 48
                                Me.picCompCard4.Image = card48
                                compCard(3) = 9
                                card(3) = 48
                            Case 49
                                Me.picCompCard4.Image = card49
                                compCard(3) = 10
                                card(3) = 49
                            Case 50
                                Me.picCompCard4.Image = card50
                                compCard(3) = 11
                                card(3) = 50
                            Case 51
                                Me.picCompCard4.Image = card51
                                compCard(3) = 12
                                card(3) = 51
                            Case 52
                                Me.picCompCard4.Image = card52
                                compCard(3) = 13
                                card(3) = 52
                        End Select
                    ElseIf card(11) = card(4) Then
                        Select Case newNum
                            Case 1
                                Me.picCompCard5.Image = card1
                                compCard(4) = 1
                                card(4) = 1
                            Case 2
                                Me.picCompCard5.Image = card2
                                compCard(4) = 2
                                card(4) = 2
                            Case 3
                                Me.picCompCard5.Image = card3
                                compCard(4) = 3
                                card(4) = 3
                            Case 4
                                Me.picCompCard5.Image = card4
                                compCard(4) = 4
                                card(4) = 4
                            Case 5
                                Me.picCompCard5.Image = card5
                                compCard(4) = 5
                                card(4) = 5
                            Case 6
                                Me.picCompCard5.Image = card6
                                compCard(4) = 6
                                card(4) = 6
                            Case 7
                                Me.picCompCard5.Image = card7
                                compCard(4) = 7
                                card(4) = 7
                            Case 8
                                Me.picCompCard5.Image = card8
                                compCard(4) = 8
                                card(4) = 8
                            Case 9
                                Me.picCompCard5.Image = card9
                                compCard(4) = 9
                                card(4) = 9
                            Case 10
                                Me.picCompCard5.Image = card10
                                compCard(4) = 10
                                card(4) = 10
                            Case 11
                                Me.picCompCard5.Image = card11
                                compCard(4) = 11
                                card(4) = 11
                            Case 12
                                Me.picCompCard5.Image = card12
                                compCard(4) = 12
                                card(4) = 12
                            Case 13
                                Me.picCompCard5.Image = card13
                                compCard(4) = 13
                                card(4) = 13
                            Case 14
                                Me.picCompCard5.Image = card14
                                compCard(4) = 1
                                card(4) = 14
                            Case 15
                                Me.picCompCard5.Image = card15
                                compCard(4) = 2
                                card(4) = 15
                            Case 16
                                Me.picCompCard5.Image = card16
                                compCard(4) = 3
                                card(4) = 16
                            Case 17
                                Me.picCompCard5.Image = card17
                                compCard(4) = 4
                                card(4) = 17
                            Case 18
                                Me.picCompCard5.Image = card18
                                compCard(4) = 5
                                card(4) = 18
                            Case 19
                                Me.picCompCard5.Image = card19
                                compCard(4) = 6
                                card(4) = 19
                            Case 20
                                Me.picCompCard5.Image = card20
                                compCard(4) = 7
                                card(4) = 20
                            Case 21
                                Me.picCompCard5.Image = card21
                                compCard(4) = 8
                                card(4) = 21
                            Case 22
                                Me.picCompCard5.Image = card22
                                compCard(4) = 9
                                card(4) = 22
                            Case 23
                                Me.picCompCard5.Image = card23
                                compCard(4) = 10
                                card(4) = 23
                            Case 24
                                Me.picCompCard5.Image = card24
                                compCard(4) = 11
                                card(4) = 24
                            Case 25
                                Me.picCompCard5.Image = card25
                                compCard(4) = 12
                                card(4) = 25
                            Case 26
                                Me.picCompCard5.Image = card26
                                compCard(4) = 13
                                card(4) = 26
                            Case 27
                                Me.picCompCard5.Image = card27
                                compCard(4) = 1
                                card(4) = 27
                            Case 28
                                Me.picCompCard5.Image = card28
                                compCard(4) = 2
                                card(4) = 28
                            Case 29
                                Me.picCompCard5.Image = card29
                                compCard(4) = 3
                                card(4) = 29
                            Case 30
                                Me.picCompCard5.Image = card30
                                compCard(4) = 4
                                card(4) = 30
                            Case 31
                                Me.picCompCard5.Image = card31
                                compCard(4) = 5
                                card(4) = 31
                            Case 32
                                Me.picCompCard5.Image = card32
                                compCard(4) = 6
                                card(4) = 32
                            Case 33
                                Me.picCompCard5.Image = card33
                                compCard(4) = 7
                                card(4) = 33
                            Case 34
                                Me.picCompCard5.Image = card34
                                compCard(4) = 8
                                card(4) = 34
                            Case 35
                                Me.picCompCard5.Image = card35
                                compCard(4) = 9
                                card(4) = 35
                            Case 36
                                Me.picCompCard5.Image = card36
                                compCard(4) = 10
                                card(4) = 36
                            Case 37
                                Me.picCompCard5.Image = card37
                                compCard(4) = 11
                                card(4) = 37
                            Case 38
                                Me.picCompCard5.Image = card38
                                compCard(4) = 12
                                card(4) = 38
                            Case 39
                                Me.picCompCard5.Image = card39
                                compCard(4) = 13
                                card(4) = 39
                            Case 40
                                Me.picCompCard5.Image = card40
                                compCard(4) = 1
                                card(4) = 40
                            Case 41
                                Me.picCompCard5.Image = card41
                                compCard(4) = 2
                                card(4) = 41
                            Case 42
                                Me.picCompCard5.Image = card42
                                compCard(4) = 3
                                card(4) = 42
                            Case 43
                                Me.picCompCard5.Image = card43
                                compCard(4) = 4
                                card(4) = 43
                            Case 44
                                Me.picCompCard5.Image = card44
                                compCard(4) = 5
                                card(4) = 44
                            Case 45
                                Me.picCompCard5.Image = card45
                                compCard(4) = 6
                                card(4) = 45
                            Case 46
                                Me.picCompCard5.Image = card46
                                compCard(4) = 7
                                card(4) = 46
                            Case 47
                                Me.picCompCard5.Image = card47
                                compCard(4) = 8
                                card(4) = 47
                            Case 48
                                Me.picCompCard5.Image = card48
                                compCard(4) = 9
                                card(4) = 48
                            Case 49
                                Me.picCompCard5.Image = card49
                                compCard(4) = 10
                                card(4) = 49
                            Case 50
                                Me.picCompCard5.Image = card50
                                compCard(4) = 11
                                card(4) = 50
                            Case 51
                                Me.picCompCard5.Image = card51
                                compCard(4) = 12
                                card(4) = 51
                            Case 52
                                Me.picCompCard5.Image = card52
                                compCard(4) = 13
                                card(4) = 52
                        End Select
                    End If

                    Used.Add(newNum)

                End If
            Loop While OK = False
        End If
        compCardsLeftTotal = compCardsLeft1 + compCardsLeft2 + compCardsLeft3 + compCardsLeft4 + compCardsLeft5
        Me.lblCompCardsLeftTotal.Text = "Computer's" & vbCrLf & "Total Cards Left:" & vbCrLf & compCardsLeftTotal
        Me.lblCompCardsLeft1.Text = "Cards Left: " & compCardsLeft1
        Me.lblCompCardsLeft2.Text = "Cards Left: " & compCardsLeft2
        Me.lblCompCardsLeft3.Text = "Cards Left: " & compCardsLeft3
        Me.lblCompCardsLeft4.Text = "Cards Left: " & compCardsLeft4
        Me.lblCompCardsLeft5.Text = "Cards Left: " & compCardsLeft5

    End Sub

    Private Sub tmrMedium_Tick(sender As Object, e As EventArgs) Handles tmrMedium.Tick
        Dim newNum As Integer
        Dim OK As Boolean = False

        For cardNum As Integer = 0 To 4                                                                'check for compCard [0 to 4]
            If compPile = compCard(cardNum) + 1 Or compPile = compCard(cardNum) - 1 Then
                Select Case cardNum
                    Case 0                                                  'if compCard(0) is 1 higher or 1 lower than compPile
                        Me.picCompPile.Image = Me.picCompCard1.Image        'set picCompPile to picCompCard1 image
                        compPile = compCard(0)                              'set compPile to compCard(0) integer [1 to 13]
                        card(10) = card(0)                                  'set card(10), compPileCard, to card(0) integer [1 to 52]
                        If compCardsLeft1 > 0 Then                          'if 1 or more cards left, 
                            compCardsLeft1 -= 1                             'subtract 1
                        End If
                    Case 1
                        Me.picCompPile.Image = Me.picCompCard2.Image
                        compPile = compCard(1)
                        card(10) = card(1)
                        If compCardsLeft2 > 0 Then
                            compCardsLeft2 -= 1
                        End If
                    Case 2
                        Me.picCompPile.Image = Me.picCompCard3.Image
                        compPile = compCard(2)
                        card(10) = card(2)
                        If compCardsLeft3 > 0 Then
                            compCardsLeft3 -= 1
                        End If
                    Case 3
                        Me.picCompPile.Image = Me.picCompCard4.Image
                        compPile = compCard(3)
                        card(10) = card(3)
                        If compCardsLeft4 > 0 Then
                            compCardsLeft4 -= 1
                        End If
                    Case 4
                        Me.picCompPile.Image = Me.picCompCard5.Image
                        compPile = compCard(4)
                        card(10) = card(4)
                        If compCardsLeft5 > 0 Then
                            compCardsLeft5 -= 1
                        End If
                End Select
                cardNum = 100

            ElseIf compPile = 1 And (compCard(cardNum) = 2 Or compCard(cardNum) = 13) Then
                Select Case cardNum
                    Case 0
                        Me.picCompPile.Image = Me.picCompCard1.Image
                        compPile = compCard(0)
                        card(10) = card(0)
                        If compCardsLeft1 > 0 Then
                            compCardsLeft1 -= 1
                        End If
                    Case 1
                        Me.picCompPile.Image = Me.picCompCard2.Image
                        compPile = compCard(1)
                        card(10) = card(1)
                        If compCardsLeft2 > 0 Then
                            compCardsLeft2 -= 1
                        End If
                    Case 2
                        Me.picCompPile.Image = Me.picCompCard3.Image
                        compPile = compCard(2)
                        card(10) = card(2)
                        If compCardsLeft3 > 0 Then
                            compCardsLeft3 -= 1
                        End If
                    Case 3
                        Me.picCompPile.Image = Me.picCompCard4.Image
                        compPile = compCard(3)
                        card(10) = card(3)
                        If compCardsLeft4 > 0 Then
                            compCardsLeft4 -= 1
                        End If
                    Case 4
                        Me.picCompPile.Image = Me.picCompCard5.Image
                        compPile = compCard(4)
                        card(10) = card(4)
                        If compCardsLeft5 > 0 Then
                            compCardsLeft5 -= 1
                        End If
                End Select
                cardNum = 100

            ElseIf compPile = 13 And (compCard(cardNum) = 1 Or compCard(cardNum) = 12) Then
                Select Case cardNum
                    Case 0
                        Me.picCompPile.Image = Me.picCompCard1.Image
                        compPile = compCard(0)
                        card(10) = card(0)
                        If compCardsLeft1 > 0 Then
                            compCardsLeft1 -= 1
                        End If
                    Case 1
                        Me.picCompPile.Image = Me.picCompCard2.Image
                        compPile = compCard(1)
                        card(10) = card(1)
                        If compCardsLeft2 > 0 Then
                            compCardsLeft2 -= 1
                        End If
                    Case 2
                        Me.picCompPile.Image = Me.picCompCard3.Image
                        compPile = compCard(2)
                        card(10) = card(2)
                        If compCardsLeft3 > 0 Then
                            compCardsLeft3 -= 1
                        End If
                    Case 3
                        Me.picCompPile.Image = Me.picCompCard4.Image
                        compPile = compCard(3)
                        card(10) = card(3)
                        If compCardsLeft4 > 0 Then
                            compCardsLeft4 -= 1
                        End If
                    Case 4
                        Me.picCompPile.Image = Me.picCompCard5.Image
                        compPile = compCard(4)
                        card(10) = card(4)
                        If compCardsLeft5 > 0 Then
                            compCardsLeft5 -= 1
                        End If
                End Select
                cardNum = 100

            ElseIf playerPile = compCard(cardNum) + 1 Or playerPile = compCard(cardNum) - 1 Then
                Select Case cardNum
                    Case 0
                        Me.picPlayerPile.Image = Me.picCompCard1.Image
                        playerPile = compCard(0)
                        card(11) = card(0)
                        If compCardsLeft1 > 0 Then
                            compCardsLeft1 -= 1
                        End If
                    Case 1
                        Me.picPlayerPile.Image = Me.picCompCard2.Image
                        playerPile = compCard(1)
                        card(11) = card(1)
                        If compCardsLeft2 > 0 Then
                            compCardsLeft2 -= 1
                        End If
                    Case 2
                        Me.picPlayerPile.Image = Me.picCompCard3.Image
                        playerPile = compCard(2)
                        card(11) = card(2)
                        If compCardsLeft3 > 0 Then
                            compCardsLeft3 -= 1
                        End If
                    Case 3
                        Me.picPlayerPile.Image = Me.picCompCard4.Image
                        playerPile = compCard(3)
                        card(11) = card(3)
                        If compCardsLeft4 > 0 Then
                            compCardsLeft4 -= 1
                        End If
                    Case 4
                        Me.picPlayerPile.Image = Me.picCompCard5.Image
                        playerPile = compCard(4)
                        card(11) = card(4)
                        If compCardsLeft5 > 0 Then
                            compCardsLeft5 -= 1
                        End If
                End Select
                cardNum = 100

            ElseIf playerPile = 1 And (compCard(cardNum) = 2 Or compCard(cardNum) = 13) Then
                Select Case cardNum
                    Case 0
                        Me.picPlayerPile.Image = Me.picCompCard1.Image
                        playerPile = compCard(0)
                        card(11) = card(0)
                        If compCardsLeft1 > 0 Then
                            compCardsLeft1 -= 1
                        End If
                    Case 1
                        Me.picPlayerPile.Image = Me.picCompCard2.Image
                        playerPile = compCard(1)
                        card(11) = card(1)
                        If compCardsLeft2 > 0 Then
                            compCardsLeft2 -= 1
                        End If
                    Case 2
                        Me.picPlayerPile.Image = Me.picCompCard3.Image
                        playerPile = compCard(2)
                        card(11) = card(2)
                        If compCardsLeft3 > 0 Then
                            compCardsLeft3 -= 1
                        End If
                    Case 3
                        Me.picPlayerPile.Image = Me.picCompCard4.Image
                        playerPile = compCard(3)
                        card(11) = card(3)
                        If compCardsLeft4 > 0 Then
                            compCardsLeft4 -= 1
                        End If
                    Case 4
                        Me.picPlayerPile.Image = Me.picCompCard5.Image
                        playerPile = compCard(4)
                        card(11) = card(4)
                        If compCardsLeft5 > 0 Then
                            compCardsLeft5 -= 1
                        End If
                End Select
                cardNum = 100

            ElseIf playerPile = 13 And (compCard(cardNum) = 1 Or compCard(cardNum) = 12) Then
                Select Case cardNum
                    Case 0
                        Me.picPlayerPile.Image = Me.picCompCard1.Image
                        playerPile = compCard(0)
                        card(11) = card(0)
                        If compCardsLeft1 > 0 Then
                            compCardsLeft1 -= 1
                        End If
                    Case 1
                        Me.picPlayerPile.Image = Me.picCompCard2.Image
                        playerPile = compCard(1)
                        card(11) = card(1)
                        If compCardsLeft2 > 0 Then
                            compCardsLeft2 -= 1
                        End If
                    Case 2
                        Me.picPlayerPile.Image = Me.picCompCard3.Image
                        playerPile = compCard(2)
                        card(11) = card(2)
                        If compCardsLeft3 > 0 Then
                            compCardsLeft3 -= 1
                        End If
                    Case 3
                        Me.picPlayerPile.Image = Me.picCompCard4.Image
                        playerPile = compCard(3)
                        card(11) = card(3)
                        If compCardsLeft4 > 0 Then
                            compCardsLeft4 -= 1
                        End If
                    Case 4
                        Me.picPlayerPile.Image = Me.picCompCard5.Image
                        playerPile = compCard(4)
                        card(11) = card(4)
                        If compCardsLeft5 > 0 Then
                            compCardsLeft5 -= 1
                        End If
                End Select
                cardNum = 100

            End If
        Next cardNum

        If card(10) = card(0) Or card(10) = card(1) Or card(10) = card(2) Or card(10) = card(3) Or card(10) = card(4) Then

            Do                                                      'generate a new card
                Randomize()
                newNum = Int(max * Rnd() + min)

                If Used.Contains(newNum) Then                       'if newNum has already been generated, do again
                    OK = False
                Else                                                'if newNum has not already been generated, show new card
                    OK = True

                    If card(10) = card(0) Then                      'if card(10), compPileCard, is the same as card(0), compPile1 integer
                        Select Case newNum
                            Case 1
                                Me.picCompCard1.Image = card1
                                compCard(0) = 1
                                card(0) = 1
                            Case 2
                                Me.picCompCard1.Image = card2
                                compCard(0) = 2
                                card(0) = 2
                            Case 3
                                Me.picCompCard1.Image = card3
                                compCard(0) = 3
                                card(0) = 3
                            Case 4
                                Me.picCompCard1.Image = card4
                                compCard(0) = 4
                                card(0) = 4
                            Case 5
                                Me.picCompCard1.Image = card5
                                compCard(0) = 5
                                card(0) = 5
                            Case 6
                                Me.picCompCard1.Image = card6
                                compCard(0) = 6
                                card(0) = 6
                            Case 7
                                Me.picCompCard1.Image = card7
                                compCard(0) = 7
                                card(0) = 7
                            Case 8
                                Me.picCompCard1.Image = card8
                                compCard(0) = 8
                                card(0) = 8
                            Case 9
                                Me.picCompCard1.Image = card9
                                compCard(0) = 9
                                card(0) = 9
                            Case 10
                                Me.picCompCard1.Image = card10
                                compCard(0) = 10
                                card(0) = 10
                            Case 11
                                Me.picCompCard1.Image = card11
                                compCard(0) = 11
                                card(0) = 11
                            Case 12
                                Me.picCompCard1.Image = card12
                                compCard(0) = 12
                                card(0) = 12
                            Case 13
                                Me.picCompCard1.Image = card13
                                compCard(0) = 13
                                card(0) = 13
                            Case 14
                                Me.picCompCard1.Image = card14
                                compCard(0) = 1
                                card(0) = 14
                            Case 15
                                Me.picCompCard1.Image = card15
                                compCard(0) = 2
                                card(0) = 15
                            Case 16
                                Me.picCompCard1.Image = card16
                                compCard(0) = 3
                                card(0) = 16
                            Case 17
                                Me.picCompCard1.Image = card17
                                compCard(0) = 4
                                card(0) = 17
                            Case 18
                                Me.picCompCard1.Image = card18
                                compCard(0) = 5
                                card(0) = 18
                            Case 19
                                Me.picCompCard1.Image = card19
                                compCard(0) = 6
                                card(0) = 19
                            Case 20
                                Me.picCompCard1.Image = card20
                                compCard(0) = 7
                                card(0) = 20
                            Case 21
                                Me.picCompCard1.Image = card21
                                compCard(0) = 8
                                card(0) = 21
                            Case 22
                                Me.picCompCard1.Image = card22
                                compCard(0) = 9
                                card(0) = 22
                            Case 23
                                Me.picCompCard1.Image = card23
                                compCard(0) = 10
                                card(0) = 23
                            Case 24
                                Me.picCompCard1.Image = card24
                                compCard(0) = 11
                                card(0) = 24
                            Case 25
                                Me.picCompCard1.Image = card25
                                compCard(0) = 12
                                card(0) = 25
                            Case 26
                                Me.picCompCard1.Image = card26
                                compCard(0) = 13
                                card(0) = 26
                            Case 27
                                Me.picCompCard1.Image = card27
                                compCard(0) = 1
                                card(0) = 27
                            Case 28
                                Me.picCompCard1.Image = card28
                                compCard(0) = 2
                                card(0) = 28
                            Case 29
                                Me.picCompCard1.Image = card29
                                compCard(0) = 3
                                card(0) = 29
                            Case 30
                                Me.picCompCard1.Image = card30
                                compCard(0) = 4
                                card(0) = 30
                            Case 31
                                Me.picCompCard1.Image = card31
                                compCard(0) = 5
                                card(0) = 31
                            Case 32
                                Me.picCompCard1.Image = card32
                                compCard(0) = 6
                                card(0) = 32
                            Case 33
                                Me.picCompCard1.Image = card33
                                compCard(0) = 7
                                card(0) = 33
                            Case 34
                                Me.picCompCard1.Image = card34
                                compCard(0) = 8
                                card(0) = 34
                            Case 35
                                Me.picCompCard1.Image = card35
                                compCard(0) = 9
                                card(0) = 35
                            Case 36
                                Me.picCompCard1.Image = card36
                                compCard(0) = 10
                                card(0) = 36
                            Case 37
                                Me.picCompCard1.Image = card37
                                compCard(0) = 11
                                card(0) = 37
                            Case 38
                                Me.picCompCard1.Image = card38
                                compCard(0) = 12
                                card(0) = 38
                            Case 39
                                Me.picCompCard1.Image = card39
                                compCard(0) = 13
                                card(0) = 39
                            Case 40
                                Me.picCompCard1.Image = card40
                                compCard(0) = 1
                                card(0) = 40
                            Case 41
                                Me.picCompCard1.Image = card41
                                compCard(0) = 2
                                card(0) = 41
                            Case 42
                                Me.picCompCard1.Image = card42
                                compCard(0) = 3
                                card(0) = 42
                            Case 43
                                Me.picCompCard1.Image = card43
                                compCard(0) = 4
                                card(0) = 43
                            Case 44
                                Me.picCompCard1.Image = card44
                                compCard(0) = 5
                                card(0) = 44
                            Case 45
                                Me.picCompCard1.Image = card45
                                compCard(0) = 6
                                card(0) = 45
                            Case 46
                                Me.picCompCard1.Image = card46
                                compCard(0) = 7
                                card(0) = 46
                            Case 47
                                Me.picCompCard1.Image = card47
                                compCard(0) = 8
                                card(0) = 47
                            Case 48
                                Me.picCompCard1.Image = card48
                                compCard(0) = 9
                                card(0) = 48
                            Case 49
                                Me.picCompCard1.Image = card49
                                compCard(0) = 10
                                card(0) = 49
                            Case 50
                                Me.picCompCard1.Image = card50
                                compCard(0) = 11
                                card(0) = 50
                            Case 51
                                Me.picCompCard1.Image = card51
                                compCard(0) = 12
                                card(0) = 51
                            Case 52
                                Me.picCompCard1.Image = card52
                                compCard(0) = 13
                                card(0) = 52
                        End Select
                    ElseIf card(10) = card(1) Then
                        Select Case newNum
                            Case 1
                                Me.picCompCard2.Image = card1
                                compCard(1) = 1
                                card(1) = 1
                            Case 2
                                Me.picCompCard2.Image = card2
                                compCard(1) = 2
                                card(1) = 2
                            Case 3
                                Me.picCompCard2.Image = card3
                                compCard(1) = 3
                                card(1) = 3
                            Case 4
                                Me.picCompCard2.Image = card4
                                compCard(1) = 4
                                card(1) = 4
                            Case 5
                                Me.picCompCard2.Image = card5
                                compCard(1) = 5
                                card(1) = 5
                            Case 6
                                Me.picCompCard2.Image = card6
                                compCard(1) = 6
                                card(1) = 6
                            Case 7
                                Me.picCompCard2.Image = card7
                                compCard(1) = 7
                                card(1) = 7
                            Case 8
                                Me.picCompCard2.Image = card8
                                compCard(1) = 8
                                card(1) = 8
                            Case 9
                                Me.picCompCard2.Image = card9
                                compCard(1) = 9
                                card(1) = 9
                            Case 10
                                Me.picCompCard2.Image = card10
                                compCard(1) = 10
                                card(1) = 10
                            Case 11
                                Me.picCompCard2.Image = card11
                                compCard(1) = 11
                                card(1) = 11
                            Case 12
                                Me.picCompCard2.Image = card12
                                compCard(1) = 12
                                card(1) = 12
                            Case 13
                                Me.picCompCard2.Image = card13
                                compCard(1) = 13
                                card(1) = 13
                            Case 14
                                Me.picCompCard2.Image = card14
                                compCard(1) = 1
                                card(1) = 14
                            Case 15
                                Me.picCompCard2.Image = card15
                                compCard(1) = 2
                                card(1) = 15
                            Case 16
                                Me.picCompCard2.Image = card16
                                compCard(1) = 3
                                card(1) = 16
                            Case 17
                                Me.picCompCard2.Image = card17
                                compCard(1) = 4
                                card(1) = 17
                            Case 18
                                Me.picCompCard2.Image = card18
                                compCard(1) = 5
                                card(1) = 18
                            Case 19
                                Me.picCompCard2.Image = card19
                                compCard(1) = 6
                                card(1) = 19
                            Case 20
                                Me.picCompCard2.Image = card20
                                compCard(1) = 7
                                card(1) = 20
                            Case 21
                                Me.picCompCard2.Image = card21
                                compCard(1) = 8
                                card(1) = 21
                            Case 22
                                Me.picCompCard2.Image = card22
                                compCard(1) = 9
                                card(1) = 22
                            Case 23
                                Me.picCompCard2.Image = card23
                                compCard(1) = 10
                                card(1) = 23
                            Case 24
                                Me.picCompCard2.Image = card24
                                compCard(1) = 11
                                card(1) = 24
                            Case 25
                                Me.picCompCard2.Image = card25
                                compCard(1) = 12
                                card(1) = 25
                            Case 26
                                Me.picCompCard2.Image = card26
                                compCard(1) = 13
                                card(1) = 26
                            Case 27
                                Me.picCompCard2.Image = card27
                                compCard(1) = 1
                                card(1) = 27
                            Case 28
                                Me.picCompCard2.Image = card28
                                compCard(1) = 2
                                card(1) = 28
                            Case 29
                                Me.picCompCard2.Image = card29
                                compCard(1) = 3
                                card(1) = 29
                            Case 30
                                Me.picCompCard2.Image = card30
                                compCard(1) = 4
                                card(1) = 30
                            Case 31
                                Me.picCompCard2.Image = card31
                                compCard(1) = 5
                                card(1) = 31
                            Case 32
                                Me.picCompCard2.Image = card32
                                compCard(1) = 6
                                card(1) = 32
                            Case 33
                                Me.picCompCard2.Image = card33
                                compCard(1) = 7
                                card(1) = 33
                            Case 34
                                Me.picCompCard2.Image = card34
                                compCard(1) = 8
                                card(1) = 34
                            Case 35
                                Me.picCompCard2.Image = card35
                                compCard(1) = 9
                                card(1) = 35
                            Case 36
                                Me.picCompCard2.Image = card36
                                compCard(1) = 10
                                card(1) = 36
                            Case 37
                                Me.picCompCard2.Image = card37
                                compCard(1) = 11
                                card(1) = 37
                            Case 38
                                Me.picCompCard2.Image = card38
                                compCard(1) = 12
                                card(1) = 38
                            Case 39
                                Me.picCompCard2.Image = card39
                                compCard(1) = 13
                                card(1) = 39
                            Case 40
                                Me.picCompCard2.Image = card40
                                compCard(1) = 1
                                card(1) = 40
                            Case 41
                                Me.picCompCard2.Image = card41
                                compCard(1) = 2
                                card(1) = 41
                            Case 42
                                Me.picCompCard2.Image = card42
                                compCard(1) = 3
                                card(1) = 42
                            Case 43
                                Me.picCompCard2.Image = card43
                                compCard(1) = 4
                                card(1) = 43
                            Case 44
                                Me.picCompCard2.Image = card44
                                compCard(1) = 5
                                card(1) = 44
                            Case 45
                                Me.picCompCard2.Image = card45
                                compCard(1) = 6
                                card(1) = 45
                            Case 46
                                Me.picCompCard2.Image = card46
                                compCard(1) = 7
                                card(1) = 46
                            Case 47
                                Me.picCompCard2.Image = card47
                                compCard(1) = 8
                                card(1) = 47
                            Case 48
                                Me.picCompCard2.Image = card48
                                compCard(1) = 9
                                card(1) = 48
                            Case 49
                                Me.picCompCard2.Image = card49
                                compCard(1) = 10
                                card(1) = 49
                            Case 50
                                Me.picCompCard2.Image = card50
                                compCard(1) = 11
                                card(1) = 50
                            Case 51
                                Me.picCompCard2.Image = card51
                                compCard(1) = 12
                                card(1) = 51
                            Case 52
                                Me.picCompCard2.Image = card52
                                compCard(1) = 13
                                card(1) = 52
                        End Select
                    ElseIf card(10) = card(2) Then
                        Select Case newNum
                            Case 1
                                Me.picCompCard3.Image = card1
                                compCard(2) = 1
                                card(2) = 1
                            Case 2
                                Me.picCompCard3.Image = card2
                                compCard(2) = 2
                                card(2) = 2
                            Case 3
                                Me.picCompCard3.Image = card3
                                compCard(2) = 3
                                card(2) = 3
                            Case 4
                                Me.picCompCard3.Image = card4
                                compCard(2) = 4
                                card(2) = 4
                            Case 5
                                Me.picCompCard3.Image = card5
                                compCard(2) = 5
                                card(2) = 5
                            Case 6
                                Me.picCompCard3.Image = card6
                                compCard(2) = 6
                                card(2) = 6
                            Case 7
                                Me.picCompCard3.Image = card7
                                compCard(2) = 7
                                card(2) = 7
                            Case 8
                                Me.picCompCard3.Image = card8
                                compCard(2) = 8
                                card(2) = 8
                            Case 9
                                Me.picCompCard3.Image = card9
                                compCard(2) = 9
                                card(2) = 9
                            Case 10
                                Me.picCompCard3.Image = card10
                                compCard(2) = 10
                                card(2) = 10
                            Case 11
                                Me.picCompCard3.Image = card11
                                compCard(2) = 11
                                card(2) = 11
                            Case 12
                                Me.picCompCard3.Image = card12
                                compCard(2) = 12
                                card(2) = 12
                            Case 13
                                Me.picCompCard3.Image = card13
                                compCard(2) = 13
                                card(2) = 13
                            Case 14
                                Me.picCompCard3.Image = card14
                                compCard(2) = 1
                                card(2) = 14
                            Case 15
                                Me.picCompCard3.Image = card15
                                compCard(2) = 2
                                card(2) = 15
                            Case 16
                                Me.picCompCard3.Image = card16
                                compCard(2) = 3
                                card(2) = 16
                            Case 17
                                Me.picCompCard3.Image = card17
                                compCard(2) = 4
                                card(2) = 17
                            Case 18
                                Me.picCompCard3.Image = card18
                                compCard(2) = 5
                                card(2) = 18
                            Case 19
                                Me.picCompCard3.Image = card19
                                compCard(2) = 6
                                card(2) = 19
                            Case 20
                                Me.picCompCard3.Image = card20
                                compCard(2) = 7
                                card(2) = 20
                            Case 21
                                Me.picCompCard3.Image = card21
                                compCard(2) = 8
                                card(2) = 21
                            Case 22
                                Me.picCompCard3.Image = card22
                                compCard(2) = 9
                                card(2) = 22
                            Case 23
                                Me.picCompCard3.Image = card23
                                compCard(2) = 10
                                card(2) = 23
                            Case 24
                                Me.picCompCard3.Image = card24
                                compCard(2) = 11
                                card(2) = 24
                            Case 25
                                Me.picCompCard3.Image = card25
                                compCard(2) = 12
                                card(2) = 25
                            Case 26
                                Me.picCompCard3.Image = card26
                                compCard(2) = 13
                                card(2) = 26
                            Case 27
                                Me.picCompCard3.Image = card27
                                compCard(2) = 1
                                card(2) = 27
                            Case 28
                                Me.picCompCard3.Image = card28
                                compCard(2) = 2
                                card(2) = 28
                            Case 29
                                Me.picCompCard3.Image = card29
                                compCard(2) = 3
                                card(2) = 29
                            Case 30
                                Me.picCompCard3.Image = card30
                                compCard(2) = 4
                                card(2) = 30
                            Case 31
                                Me.picCompCard3.Image = card31
                                compCard(2) = 5
                                card(2) = 31
                            Case 32
                                Me.picCompCard3.Image = card32
                                compCard(2) = 6
                                card(2) = 32
                            Case 33
                                Me.picCompCard3.Image = card33
                                compCard(2) = 7
                                card(2) = 33
                            Case 34
                                Me.picCompCard3.Image = card34
                                compCard(2) = 8
                                card(2) = 34
                            Case 35
                                Me.picCompCard3.Image = card35
                                compCard(2) = 9
                                card(2) = 35
                            Case 36
                                Me.picCompCard3.Image = card36
                                compCard(2) = 10
                                card(2) = 36
                            Case 37
                                Me.picCompCard3.Image = card37
                                compCard(2) = 11
                                card(2) = 37
                            Case 38
                                Me.picCompCard3.Image = card38
                                compCard(2) = 12
                                card(2) = 38
                            Case 39
                                Me.picCompCard3.Image = card39
                                compCard(2) = 13
                                card(2) = 39
                            Case 40
                                Me.picCompCard3.Image = card40
                                compCard(2) = 1
                                card(2) = 40
                            Case 41
                                Me.picCompCard3.Image = card41
                                compCard(2) = 2
                                card(2) = 41
                            Case 42
                                Me.picCompCard3.Image = card42
                                compCard(2) = 3
                                card(2) = 42
                            Case 43
                                Me.picCompCard3.Image = card43
                                compCard(2) = 4
                                card(2) = 43
                            Case 44
                                Me.picCompCard3.Image = card44
                                compCard(2) = 5
                                card(2) = 44
                            Case 45
                                Me.picCompCard3.Image = card45
                                compCard(2) = 6
                                card(2) = 45
                            Case 46
                                Me.picCompCard3.Image = card46
                                compCard(2) = 7
                                card(2) = 46
                            Case 47
                                Me.picCompCard3.Image = card47
                                compCard(2) = 8
                                card(2) = 47
                            Case 48
                                Me.picCompCard3.Image = card48
                                compCard(2) = 9
                                card(2) = 48
                            Case 49
                                Me.picCompCard3.Image = card49
                                compCard(2) = 10
                                card(2) = 49
                            Case 50
                                Me.picCompCard3.Image = card50
                                compCard(2) = 11
                                card(2) = 50
                            Case 51
                                Me.picCompCard3.Image = card51
                                compCard(2) = 12
                                card(2) = 51
                            Case 52
                                Me.picCompCard3.Image = card52
                                compCard(2) = 13
                                card(2) = 52
                        End Select
                    ElseIf card(10) = card(3) Then
                        Select Case newNum
                            Case 1
                                Me.picCompCard4.Image = card1
                                compCard(3) = 1
                                card(3) = 1
                            Case 2
                                Me.picCompCard4.Image = card2
                                compCard(3) = 2
                                card(3) = 2
                            Case 3
                                Me.picCompCard4.Image = card3
                                compCard(3) = 3
                                card(3) = 3
                            Case 4
                                Me.picCompCard4.Image = card4
                                compCard(3) = 4
                                card(3) = 4
                            Case 5
                                Me.picCompCard4.Image = card5
                                compCard(3) = 5
                                card(3) = 5
                            Case 6
                                Me.picCompCard4.Image = card6
                                compCard(3) = 6
                                card(3) = 6
                            Case 7
                                Me.picCompCard4.Image = card7
                                compCard(3) = 7
                                card(3) = 7
                            Case 8
                                Me.picCompCard4.Image = card8
                                compCard(3) = 8
                                card(3) = 8
                            Case 9
                                Me.picCompCard4.Image = card9
                                compCard(3) = 9
                                card(3) = 9
                            Case 10
                                Me.picCompCard4.Image = card10
                                compCard(3) = 10
                                card(3) = 10
                            Case 11
                                Me.picCompCard4.Image = card11
                                compCard(3) = 11
                                card(3) = 11
                            Case 12
                                Me.picCompCard4.Image = card12
                                compCard(3) = 12
                                card(3) = 12
                            Case 13
                                Me.picCompCard4.Image = card13
                                compCard(3) = 13
                                card(3) = 13
                            Case 14
                                Me.picCompCard4.Image = card14
                                compCard(3) = 1
                                card(3) = 14
                            Case 15
                                Me.picCompCard4.Image = card15
                                compCard(3) = 2
                                card(3) = 15
                            Case 16
                                Me.picCompCard4.Image = card16
                                compCard(3) = 3
                                card(3) = 16
                            Case 17
                                Me.picCompCard4.Image = card17
                                compCard(3) = 4
                                card(3) = 17
                            Case 18
                                Me.picCompCard4.Image = card18
                                compCard(3) = 5
                                card(3) = 18
                            Case 19
                                Me.picCompCard4.Image = card19
                                compCard(3) = 6
                                card(3) = 19
                            Case 20
                                Me.picCompCard4.Image = card20
                                compCard(3) = 7
                                card(3) = 20
                            Case 21
                                Me.picCompCard4.Image = card21
                                compCard(3) = 8
                                card(3) = 21
                            Case 22
                                Me.picCompCard4.Image = card22
                                compCard(3) = 9
                                card(3) = 22
                            Case 23
                                Me.picCompCard4.Image = card23
                                compCard(3) = 10
                                card(3) = 23
                            Case 24
                                Me.picCompCard4.Image = card24
                                compCard(3) = 11
                                card(3) = 24
                            Case 25
                                Me.picCompCard4.Image = card25
                                compCard(3) = 12
                                card(3) = 25
                            Case 26
                                Me.picCompCard4.Image = card26
                                compCard(3) = 13
                                card(3) = 26
                            Case 27
                                Me.picCompCard4.Image = card27
                                compCard(3) = 1
                                card(3) = 27
                            Case 28
                                Me.picCompCard4.Image = card28
                                compCard(3) = 2
                                card(3) = 28
                            Case 29
                                Me.picCompCard4.Image = card29
                                compCard(3) = 3
                                card(3) = 29
                            Case 30
                                Me.picCompCard4.Image = card30
                                compCard(3) = 4
                                card(3) = 30
                            Case 31
                                Me.picCompCard4.Image = card31
                                compCard(3) = 5
                                card(3) = 31
                            Case 32
                                Me.picCompCard4.Image = card32
                                compCard(3) = 6
                                card(3) = 32
                            Case 33
                                Me.picCompCard4.Image = card33
                                compCard(3) = 7
                                card(3) = 33
                            Case 34
                                Me.picCompCard4.Image = card34
                                compCard(3) = 8
                                card(3) = 34
                            Case 35
                                Me.picCompCard4.Image = card35
                                compCard(3) = 9
                                card(3) = 35
                            Case 36
                                Me.picCompCard4.Image = card36
                                compCard(3) = 10
                                card(3) = 36
                            Case 37
                                Me.picCompCard4.Image = card37
                                compCard(3) = 11
                                card(3) = 37
                            Case 38
                                Me.picCompCard4.Image = card38
                                compCard(3) = 12
                                card(3) = 38
                            Case 39
                                Me.picCompCard4.Image = card39
                                compCard(3) = 13
                                card(3) = 39
                            Case 40
                                Me.picCompCard4.Image = card40
                                compCard(3) = 1
                                card(3) = 40
                            Case 41
                                Me.picCompCard4.Image = card41
                                compCard(3) = 2
                                card(3) = 41
                            Case 42
                                Me.picCompCard4.Image = card42
                                compCard(3) = 3
                                card(3) = 42
                            Case 43
                                Me.picCompCard4.Image = card43
                                compCard(3) = 4
                                card(3) = 43
                            Case 44
                                Me.picCompCard4.Image = card44
                                compCard(3) = 5
                                card(3) = 44
                            Case 45
                                Me.picCompCard4.Image = card45
                                compCard(3) = 6
                                card(3) = 45
                            Case 46
                                Me.picCompCard4.Image = card46
                                compCard(3) = 7
                                card(3) = 46
                            Case 47
                                Me.picCompCard4.Image = card47
                                compCard(3) = 8
                                card(3) = 47
                            Case 48
                                Me.picCompCard4.Image = card48
                                compCard(3) = 9
                                card(3) = 48
                            Case 49
                                Me.picCompCard4.Image = card49
                                compCard(3) = 10
                                card(3) = 49
                            Case 50
                                Me.picCompCard4.Image = card50
                                compCard(3) = 11
                                card(3) = 50
                            Case 51
                                Me.picCompCard4.Image = card51
                                compCard(3) = 12
                                card(3) = 51
                            Case 52
                                Me.picCompCard4.Image = card52
                                compCard(3) = 13
                                card(3) = 52
                        End Select
                    ElseIf card(10) = card(4) Then
                        Select Case newNum
                            Case 1
                                Me.picCompCard5.Image = card1
                                compCard(4) = 1
                                card(4) = 1
                            Case 2
                                Me.picCompCard5.Image = card2
                                compCard(4) = 2
                                card(4) = 2
                            Case 3
                                Me.picCompCard5.Image = card3
                                compCard(4) = 3
                                card(4) = 3
                            Case 4
                                Me.picCompCard5.Image = card4
                                compCard(4) = 4
                                card(4) = 4
                            Case 5
                                Me.picCompCard5.Image = card5
                                compCard(4) = 5
                                card(4) = 5
                            Case 6
                                Me.picCompCard5.Image = card6
                                compCard(4) = 6
                                card(4) = 6
                            Case 7
                                Me.picCompCard5.Image = card7
                                compCard(4) = 7
                                card(4) = 7
                            Case 8
                                Me.picCompCard5.Image = card8
                                compCard(4) = 8
                                card(4) = 8
                            Case 9
                                Me.picCompCard5.Image = card9
                                compCard(4) = 9
                                card(4) = 9
                            Case 10
                                Me.picCompCard5.Image = card10
                                compCard(4) = 10
                                card(4) = 10
                            Case 11
                                Me.picCompCard5.Image = card11
                                compCard(4) = 11
                                card(4) = 11
                            Case 12
                                Me.picCompCard5.Image = card12
                                compCard(4) = 12
                                card(4) = 12
                            Case 13
                                Me.picCompCard5.Image = card13
                                compCard(4) = 13
                                card(4) = 13
                            Case 14
                                Me.picCompCard5.Image = card14
                                compCard(4) = 1
                                card(4) = 14
                            Case 15
                                Me.picCompCard5.Image = card15
                                compCard(4) = 2
                                card(4) = 15
                            Case 16
                                Me.picCompCard5.Image = card16
                                compCard(4) = 3
                                card(4) = 16
                            Case 17
                                Me.picCompCard5.Image = card17
                                compCard(4) = 4
                                card(4) = 17
                            Case 18
                                Me.picCompCard5.Image = card18
                                compCard(4) = 5
                                card(4) = 18
                            Case 19
                                Me.picCompCard5.Image = card19
                                compCard(4) = 6
                                card(4) = 19
                            Case 20
                                Me.picCompCard5.Image = card20
                                compCard(4) = 7
                                card(4) = 20
                            Case 21
                                Me.picCompCard5.Image = card21
                                compCard(4) = 8
                                card(4) = 21
                            Case 22
                                Me.picCompCard5.Image = card22
                                compCard(4) = 9
                                card(4) = 22
                            Case 23
                                Me.picCompCard5.Image = card23
                                compCard(4) = 10
                                card(4) = 23
                            Case 24
                                Me.picCompCard5.Image = card24
                                compCard(4) = 11
                                card(4) = 24
                            Case 25
                                Me.picCompCard5.Image = card25
                                compCard(4) = 12
                                card(4) = 25
                            Case 26
                                Me.picCompCard5.Image = card26
                                compCard(4) = 13
                                card(4) = 26
                            Case 27
                                Me.picCompCard5.Image = card27
                                compCard(4) = 1
                                card(4) = 27
                            Case 28
                                Me.picCompCard5.Image = card28
                                compCard(4) = 2
                                card(4) = 28
                            Case 29
                                Me.picCompCard5.Image = card29
                                compCard(4) = 3
                                card(4) = 29
                            Case 30
                                Me.picCompCard5.Image = card30
                                compCard(4) = 4
                                card(4) = 30
                            Case 31
                                Me.picCompCard5.Image = card31
                                compCard(4) = 5
                                card(4) = 31
                            Case 32
                                Me.picCompCard5.Image = card32
                                compCard(4) = 6
                                card(4) = 32
                            Case 33
                                Me.picCompCard5.Image = card33
                                compCard(4) = 7
                                card(4) = 33
                            Case 34
                                Me.picCompCard5.Image = card34
                                compCard(4) = 8
                                card(4) = 34
                            Case 35
                                Me.picCompCard5.Image = card35
                                compCard(4) = 9
                                card(4) = 35
                            Case 36
                                Me.picCompCard5.Image = card36
                                compCard(4) = 10
                                card(4) = 36
                            Case 37
                                Me.picCompCard5.Image = card37
                                compCard(4) = 11
                                card(4) = 37
                            Case 38
                                Me.picCompCard5.Image = card38
                                compCard(4) = 12
                                card(4) = 38
                            Case 39
                                Me.picCompCard5.Image = card39
                                compCard(4) = 13
                                card(4) = 39
                            Case 40
                                Me.picCompCard5.Image = card40
                                compCard(4) = 1
                                card(4) = 40
                            Case 41
                                Me.picCompCard5.Image = card41
                                compCard(4) = 2
                                card(4) = 41
                            Case 42
                                Me.picCompCard5.Image = card42
                                compCard(4) = 3
                                card(4) = 42
                            Case 43
                                Me.picCompCard5.Image = card43
                                compCard(4) = 4
                                card(4) = 43
                            Case 44
                                Me.picCompCard5.Image = card44
                                compCard(4) = 5
                                card(4) = 44
                            Case 45
                                Me.picCompCard5.Image = card45
                                compCard(4) = 6
                                card(4) = 45
                            Case 46
                                Me.picCompCard5.Image = card46
                                compCard(4) = 7
                                card(4) = 46
                            Case 47
                                Me.picCompCard5.Image = card47
                                compCard(4) = 8
                                card(4) = 47
                            Case 48
                                Me.picCompCard5.Image = card48
                                compCard(4) = 9
                                card(4) = 48
                            Case 49
                                Me.picCompCard5.Image = card49
                                compCard(4) = 10
                                card(4) = 49
                            Case 50
                                Me.picCompCard5.Image = card50
                                compCard(4) = 11
                                card(4) = 50
                            Case 51
                                Me.picCompCard5.Image = card51
                                compCard(4) = 12
                                card(4) = 51
                            Case 52
                                Me.picCompCard5.Image = card52
                                compCard(4) = 13
                                card(4) = 52
                        End Select
                    End If

                    Used.Add(newNum)

                End If
            Loop While OK = False
        End If

        If card(11) = card(0) Or card(11) = card(1) Or card(11) = card(2) Or card(11) = card(3) Or card(11) = card(4) Then
            Do                                          'generate a new unique random number
                Randomize()
                newNum = Int(max * Rnd() + min)

                If Used.Contains(newNum) Then
                    OK = False
                Else
                    OK = True

                    If card(11) = card(0) Then
                        Select Case newNum
                            Case 1
                                Me.picCompCard1.Image = card1
                                compCard(0) = 1
                                card(0) = 1
                            Case 2
                                Me.picCompCard1.Image = card2
                                compCard(0) = 2
                                card(0) = 2
                            Case 3
                                Me.picCompCard1.Image = card3
                                compCard(0) = 3
                                card(0) = 3
                            Case 4
                                Me.picCompCard1.Image = card4
                                compCard(0) = 4
                                card(0) = 4
                            Case 5
                                Me.picCompCard1.Image = card5
                                compCard(0) = 5
                                card(0) = 5
                            Case 6
                                Me.picCompCard1.Image = card6
                                compCard(0) = 6
                                card(0) = 6
                            Case 7
                                Me.picCompCard1.Image = card7
                                compCard(0) = 7
                                card(0) = 7
                            Case 8
                                Me.picCompCard1.Image = card8
                                compCard(0) = 8
                                card(0) = 8
                            Case 9
                                Me.picCompCard1.Image = card9
                                compCard(0) = 9
                                card(0) = 9
                            Case 10
                                Me.picCompCard1.Image = card10
                                compCard(0) = 10
                                card(0) = 10
                            Case 11
                                Me.picCompCard1.Image = card11
                                compCard(0) = 11
                                card(0) = 11
                            Case 12
                                Me.picCompCard1.Image = card12
                                compCard(0) = 12
                                card(0) = 12
                            Case 13
                                Me.picCompCard1.Image = card13
                                compCard(0) = 13
                                card(0) = 13
                            Case 14
                                Me.picCompCard1.Image = card14
                                compCard(0) = 1
                                card(0) = 14
                            Case 15
                                Me.picCompCard1.Image = card15
                                compCard(0) = 2
                                card(0) = 15
                            Case 16
                                Me.picCompCard1.Image = card16
                                compCard(0) = 3
                                card(0) = 16
                            Case 17
                                Me.picCompCard1.Image = card17
                                compCard(0) = 4
                                card(0) = 17
                            Case 18
                                Me.picCompCard1.Image = card18
                                compCard(0) = 5
                                card(0) = 18
                            Case 19
                                Me.picCompCard1.Image = card19
                                compCard(0) = 6
                                card(0) = 19
                            Case 20
                                Me.picCompCard1.Image = card20
                                compCard(0) = 7
                                card(0) = 20
                            Case 21
                                Me.picCompCard1.Image = card21
                                compCard(0) = 8
                                card(0) = 21
                            Case 22
                                Me.picCompCard1.Image = card22
                                compCard(0) = 9
                                card(0) = 22
                            Case 23
                                Me.picCompCard1.Image = card23
                                compCard(0) = 10
                                card(0) = 23
                            Case 24
                                Me.picCompCard1.Image = card24
                                compCard(0) = 11
                                card(0) = 24
                            Case 25
                                Me.picCompCard1.Image = card25
                                compCard(0) = 12
                                card(0) = 25
                            Case 26
                                Me.picCompCard1.Image = card26
                                compCard(0) = 13
                                card(0) = 26
                            Case 27
                                Me.picCompCard1.Image = card27
                                compCard(0) = 1
                                card(0) = 27
                            Case 28
                                Me.picCompCard1.Image = card28
                                compCard(0) = 2
                                card(0) = 28
                            Case 29
                                Me.picCompCard1.Image = card29
                                compCard(0) = 3
                                card(0) = 29
                            Case 30
                                Me.picCompCard1.Image = card30
                                compCard(0) = 4
                                card(0) = 30
                            Case 31
                                Me.picCompCard1.Image = card31
                                compCard(0) = 5
                                card(0) = 31
                            Case 32
                                Me.picCompCard1.Image = card32
                                compCard(0) = 6
                                card(0) = 32
                            Case 33
                                Me.picCompCard1.Image = card33
                                compCard(0) = 7
                                card(0) = 33
                            Case 34
                                Me.picCompCard1.Image = card34
                                compCard(0) = 8
                                card(0) = 34
                            Case 35
                                Me.picCompCard1.Image = card35
                                compCard(0) = 9
                                card(0) = 35
                            Case 36
                                Me.picCompCard1.Image = card36
                                compCard(0) = 10
                                card(0) = 36
                            Case 37
                                Me.picCompCard1.Image = card37
                                compCard(0) = 11
                                card(0) = 37
                            Case 38
                                Me.picCompCard1.Image = card38
                                compCard(0) = 12
                                card(0) = 38
                            Case 39
                                Me.picCompCard1.Image = card39
                                compCard(0) = 13
                                card(0) = 39
                            Case 40
                                Me.picCompCard1.Image = card40
                                compCard(0) = 1
                                card(0) = 40
                            Case 41
                                Me.picCompCard1.Image = card41
                                compCard(0) = 2
                                card(0) = 41
                            Case 42
                                Me.picCompCard1.Image = card42
                                compCard(0) = 3
                                card(0) = 42
                            Case 43
                                Me.picCompCard1.Image = card43
                                compCard(0) = 4
                                card(0) = 43
                            Case 44
                                Me.picCompCard1.Image = card44
                                compCard(0) = 5
                                card(0) = 44
                            Case 45
                                Me.picCompCard1.Image = card45
                                compCard(0) = 6
                                card(0) = 45
                            Case 46
                                Me.picCompCard1.Image = card46
                                compCard(0) = 7
                                card(0) = 46
                            Case 47
                                Me.picCompCard1.Image = card47
                                compCard(0) = 8
                                card(0) = 47
                            Case 48
                                Me.picCompCard1.Image = card48
                                compCard(0) = 9
                                card(0) = 48
                            Case 49
                                Me.picCompCard1.Image = card49
                                compCard(0) = 10
                                card(0) = 49
                            Case 50
                                Me.picCompCard1.Image = card50
                                compCard(0) = 11
                                card(0) = 50
                            Case 51
                                Me.picCompCard1.Image = card51
                                compCard(0) = 12
                                card(0) = 51
                            Case 52
                                Me.picCompCard1.Image = card52
                                compCard(0) = 13
                                card(0) = 52
                        End Select
                    ElseIf card(11) = card(1) Then
                        Select Case newNum
                            Case 1
                                Me.picCompCard2.Image = card1
                                compCard(1) = 1
                                card(1) = 1
                            Case 2
                                Me.picCompCard2.Image = card2
                                compCard(1) = 2
                                card(1) = 2
                            Case 3
                                Me.picCompCard2.Image = card3
                                compCard(1) = 3
                                card(1) = 3
                            Case 4
                                Me.picCompCard2.Image = card4
                                compCard(1) = 4
                                card(1) = 4
                            Case 5
                                Me.picCompCard2.Image = card5
                                compCard(1) = 5
                                card(1) = 5
                            Case 6
                                Me.picCompCard2.Image = card6
                                compCard(1) = 6
                                card(1) = 6
                            Case 7
                                Me.picCompCard2.Image = card7
                                compCard(1) = 7
                                card(1) = 7
                            Case 8
                                Me.picCompCard2.Image = card8
                                compCard(1) = 8
                                card(1) = 8
                            Case 9
                                Me.picCompCard2.Image = card9
                                compCard(1) = 9
                                card(1) = 9
                            Case 10
                                Me.picCompCard2.Image = card10
                                compCard(1) = 10
                                card(1) = 10
                            Case 11
                                Me.picCompCard2.Image = card11
                                compCard(1) = 11
                                card(1) = 11
                            Case 12
                                Me.picCompCard2.Image = card12
                                compCard(1) = 12
                                card(1) = 12
                            Case 13
                                Me.picCompCard2.Image = card13
                                compCard(1) = 13
                                card(1) = 13
                            Case 14
                                Me.picCompCard2.Image = card14
                                compCard(1) = 1
                                card(1) = 14
                            Case 15
                                Me.picCompCard2.Image = card15
                                compCard(1) = 2
                                card(1) = 15
                            Case 16
                                Me.picCompCard2.Image = card16
                                compCard(1) = 3
                                card(1) = 16
                            Case 17
                                Me.picCompCard2.Image = card17
                                compCard(1) = 4
                                card(1) = 17
                            Case 18
                                Me.picCompCard2.Image = card18
                                compCard(1) = 5
                                card(1) = 18
                            Case 19
                                Me.picCompCard2.Image = card19
                                compCard(1) = 6
                                card(1) = 19
                            Case 20
                                Me.picCompCard2.Image = card20
                                compCard(1) = 7
                                card(1) = 20
                            Case 21
                                Me.picCompCard2.Image = card21
                                compCard(1) = 8
                                card(1) = 21
                            Case 22
                                Me.picCompCard2.Image = card22
                                compCard(1) = 9
                                card(1) = 22
                            Case 23
                                Me.picCompCard2.Image = card23
                                compCard(1) = 10
                                card(1) = 23
                            Case 24
                                Me.picCompCard2.Image = card24
                                compCard(1) = 11
                                card(1) = 24
                            Case 25
                                Me.picCompCard2.Image = card25
                                compCard(1) = 12
                                card(1) = 25
                            Case 26
                                Me.picCompCard2.Image = card26
                                compCard(1) = 13
                                card(1) = 26
                            Case 27
                                Me.picCompCard2.Image = card27
                                compCard(1) = 1
                                card(1) = 27
                            Case 28
                                Me.picCompCard2.Image = card28
                                compCard(1) = 2
                                card(1) = 28
                            Case 29
                                Me.picCompCard2.Image = card29
                                compCard(1) = 3
                                card(1) = 29
                            Case 30
                                Me.picCompCard2.Image = card30
                                compCard(1) = 4
                                card(1) = 30
                            Case 31
                                Me.picCompCard2.Image = card31
                                compCard(1) = 5
                                card(1) = 31
                            Case 32
                                Me.picCompCard2.Image = card32
                                compCard(1) = 6
                                card(1) = 32
                            Case 33
                                Me.picCompCard2.Image = card33
                                compCard(1) = 7
                                card(1) = 33
                            Case 34
                                Me.picCompCard2.Image = card34
                                compCard(1) = 8
                                card(1) = 34
                            Case 35
                                Me.picCompCard2.Image = card35
                                compCard(1) = 9
                                card(1) = 35
                            Case 36
                                Me.picCompCard2.Image = card36
                                compCard(1) = 10
                                card(1) = 36
                            Case 37
                                Me.picCompCard2.Image = card37
                                compCard(1) = 11
                                card(1) = 37
                            Case 38
                                Me.picCompCard2.Image = card38
                                compCard(1) = 12
                                card(1) = 38
                            Case 39
                                Me.picCompCard2.Image = card39
                                compCard(1) = 13
                                card(1) = 39
                            Case 40
                                Me.picCompCard2.Image = card40
                                compCard(1) = 1
                                card(1) = 40
                            Case 41
                                Me.picCompCard2.Image = card41
                                compCard(1) = 2
                                card(1) = 41
                            Case 42
                                Me.picCompCard2.Image = card42
                                compCard(1) = 3
                                card(1) = 42
                            Case 43
                                Me.picCompCard2.Image = card43
                                compCard(1) = 4
                                card(1) = 43
                            Case 44
                                Me.picCompCard2.Image = card44
                                compCard(1) = 5
                                card(1) = 44
                            Case 45
                                Me.picCompCard2.Image = card45
                                compCard(1) = 6
                                card(1) = 45
                            Case 46
                                Me.picCompCard2.Image = card46
                                compCard(1) = 7
                                card(1) = 46
                            Case 47
                                Me.picCompCard2.Image = card47
                                compCard(1) = 8
                                card(1) = 47
                            Case 48
                                Me.picCompCard2.Image = card48
                                compCard(1) = 9
                                card(1) = 48
                            Case 49
                                Me.picCompCard2.Image = card49
                                compCard(1) = 10
                                card(1) = 49
                            Case 50
                                Me.picCompCard2.Image = card50
                                compCard(1) = 11
                                card(1) = 50
                            Case 51
                                Me.picCompCard2.Image = card51
                                compCard(1) = 12
                                card(1) = 51
                            Case 52
                                Me.picCompCard2.Image = card52
                                compCard(1) = 13
                                card(1) = 52
                        End Select
                    ElseIf card(11) = card(2) Then
                        Select Case newNum
                            Case 1
                                Me.picCompCard3.Image = card1
                                compCard(2) = 1
                                card(2) = 1
                            Case 2
                                Me.picCompCard3.Image = card2
                                compCard(2) = 2
                                card(2) = 2
                            Case 3
                                Me.picCompCard3.Image = card3
                                compCard(2) = 3
                                card(2) = 3
                            Case 4
                                Me.picCompCard3.Image = card4
                                compCard(2) = 4
                                card(2) = 4
                            Case 5
                                Me.picCompCard3.Image = card5
                                compCard(2) = 5
                                card(2) = 5
                            Case 6
                                Me.picCompCard3.Image = card6
                                compCard(2) = 6
                                card(2) = 6
                            Case 7
                                Me.picCompCard3.Image = card7
                                compCard(2) = 7
                                card(2) = 7
                            Case 8
                                Me.picCompCard3.Image = card8
                                compCard(2) = 8
                                card(2) = 8
                            Case 9
                                Me.picCompCard3.Image = card9
                                compCard(2) = 9
                                card(2) = 9
                            Case 10
                                Me.picCompCard3.Image = card10
                                compCard(2) = 10
                                card(2) = 10
                            Case 11
                                Me.picCompCard3.Image = card11
                                compCard(2) = 11
                                card(2) = 11
                            Case 12
                                Me.picCompCard3.Image = card12
                                compCard(2) = 12
                                card(2) = 12
                            Case 13
                                Me.picCompCard3.Image = card13
                                compCard(2) = 13
                                card(2) = 13
                            Case 14
                                Me.picCompCard3.Image = card14
                                compCard(2) = 1
                                card(2) = 14
                            Case 15
                                Me.picCompCard3.Image = card15
                                compCard(2) = 2
                                card(2) = 15
                            Case 16
                                Me.picCompCard3.Image = card16
                                compCard(2) = 3
                                card(2) = 16
                            Case 17
                                Me.picCompCard3.Image = card17
                                compCard(2) = 4
                                card(2) = 17
                            Case 18
                                Me.picCompCard3.Image = card18
                                compCard(2) = 5
                                card(2) = 18
                            Case 19
                                Me.picCompCard3.Image = card19
                                compCard(2) = 6
                                card(2) = 19
                            Case 20
                                Me.picCompCard3.Image = card20
                                compCard(2) = 7
                                card(2) = 20
                            Case 21
                                Me.picCompCard3.Image = card21
                                compCard(2) = 8
                                card(2) = 21
                            Case 22
                                Me.picCompCard3.Image = card22
                                compCard(2) = 9
                                card(2) = 22
                            Case 23
                                Me.picCompCard3.Image = card23
                                compCard(2) = 10
                                card(2) = 23
                            Case 24
                                Me.picCompCard3.Image = card24
                                compCard(2) = 11
                                card(2) = 24
                            Case 25
                                Me.picCompCard3.Image = card25
                                compCard(2) = 12
                                card(2) = 25
                            Case 26
                                Me.picCompCard3.Image = card26
                                compCard(2) = 13
                                card(2) = 26
                            Case 27
                                Me.picCompCard3.Image = card27
                                compCard(2) = 1
                                card(2) = 27
                            Case 28
                                Me.picCompCard3.Image = card28
                                compCard(2) = 2
                                card(2) = 28
                            Case 29
                                Me.picCompCard3.Image = card29
                                compCard(2) = 3
                                card(2) = 29
                            Case 30
                                Me.picCompCard3.Image = card30
                                compCard(2) = 4
                                card(2) = 30
                            Case 31
                                Me.picCompCard3.Image = card31
                                compCard(2) = 5
                                card(2) = 31
                            Case 32
                                Me.picCompCard3.Image = card32
                                compCard(2) = 6
                                card(2) = 32
                            Case 33
                                Me.picCompCard3.Image = card33
                                compCard(2) = 7
                                card(2) = 33
                            Case 34
                                Me.picCompCard3.Image = card34
                                compCard(2) = 8
                                card(2) = 34
                            Case 35
                                Me.picCompCard3.Image = card35
                                compCard(2) = 9
                                card(2) = 35
                            Case 36
                                Me.picCompCard3.Image = card36
                                compCard(2) = 10
                                card(2) = 36
                            Case 37
                                Me.picCompCard3.Image = card37
                                compCard(2) = 11
                                card(2) = 37
                            Case 38
                                Me.picCompCard3.Image = card38
                                compCard(2) = 12
                                card(2) = 38
                            Case 39
                                Me.picCompCard3.Image = card39
                                compCard(2) = 13
                                card(2) = 39
                            Case 40
                                Me.picCompCard3.Image = card40
                                compCard(2) = 1
                                card(2) = 40
                            Case 41
                                Me.picCompCard3.Image = card41
                                compCard(2) = 2
                                card(2) = 41
                            Case 42
                                Me.picCompCard3.Image = card42
                                compCard(2) = 3
                                card(2) = 42
                            Case 43
                                Me.picCompCard3.Image = card43
                                compCard(2) = 4
                                card(2) = 43
                            Case 44
                                Me.picCompCard3.Image = card44
                                compCard(2) = 5
                                card(2) = 44
                            Case 45
                                Me.picCompCard3.Image = card45
                                compCard(2) = 6
                                card(2) = 45
                            Case 46
                                Me.picCompCard3.Image = card46
                                compCard(2) = 7
                                card(2) = 46
                            Case 47
                                Me.picCompCard3.Image = card47
                                compCard(2) = 8
                                card(2) = 47
                            Case 48
                                Me.picCompCard3.Image = card48
                                compCard(2) = 9
                                card(2) = 48
                            Case 49
                                Me.picCompCard3.Image = card49
                                compCard(2) = 10
                                card(2) = 49
                            Case 50
                                Me.picCompCard3.Image = card50
                                compCard(2) = 11
                                card(2) = 50
                            Case 51
                                Me.picCompCard3.Image = card51
                                compCard(2) = 12
                                card(2) = 51
                            Case 52
                                Me.picCompCard3.Image = card52
                                compCard(2) = 13
                                card(2) = 52
                        End Select
                    ElseIf card(11) = card(3) Then
                        Select Case newNum
                            Case 1
                                Me.picCompCard4.Image = card1
                                compCard(3) = 1
                                card(3) = 1
                            Case 2
                                Me.picCompCard4.Image = card2
                                compCard(3) = 2
                                card(3) = 2
                            Case 3
                                Me.picCompCard4.Image = card3
                                compCard(3) = 3
                                card(3) = 3
                            Case 4
                                Me.picCompCard4.Image = card4
                                compCard(3) = 4
                                card(3) = 4
                            Case 5
                                Me.picCompCard4.Image = card5
                                compCard(3) = 5
                                card(3) = 5
                            Case 6
                                Me.picCompCard4.Image = card6
                                compCard(3) = 6
                                card(3) = 6
                            Case 7
                                Me.picCompCard4.Image = card7
                                compCard(3) = 7
                                card(3) = 7
                            Case 8
                                Me.picCompCard4.Image = card8
                                compCard(3) = 8
                                card(3) = 8
                            Case 9
                                Me.picCompCard4.Image = card9
                                compCard(3) = 9
                                card(3) = 9
                            Case 10
                                Me.picCompCard4.Image = card10
                                compCard(3) = 10
                                card(3) = 10
                            Case 11
                                Me.picCompCard4.Image = card11
                                compCard(3) = 11
                                card(3) = 11
                            Case 12
                                Me.picCompCard4.Image = card12
                                compCard(3) = 12
                                card(3) = 12
                            Case 13
                                Me.picCompCard4.Image = card13
                                compCard(3) = 13
                                card(3) = 13
                            Case 14
                                Me.picCompCard4.Image = card14
                                compCard(3) = 1
                                card(3) = 14
                            Case 15
                                Me.picCompCard4.Image = card15
                                compCard(3) = 2
                                card(3) = 15
                            Case 16
                                Me.picCompCard4.Image = card16
                                compCard(3) = 3
                                card(3) = 16
                            Case 17
                                Me.picCompCard4.Image = card17
                                compCard(3) = 4
                                card(3) = 17
                            Case 18
                                Me.picCompCard4.Image = card18
                                compCard(3) = 5
                                card(3) = 18
                            Case 19
                                Me.picCompCard4.Image = card19
                                compCard(3) = 6
                                card(3) = 19
                            Case 20
                                Me.picCompCard4.Image = card20
                                compCard(3) = 7
                                card(3) = 20
                            Case 21
                                Me.picCompCard4.Image = card21
                                compCard(3) = 8
                                card(3) = 21
                            Case 22
                                Me.picCompCard4.Image = card22
                                compCard(3) = 9
                                card(3) = 22
                            Case 23
                                Me.picCompCard4.Image = card23
                                compCard(3) = 10
                                card(3) = 23
                            Case 24
                                Me.picCompCard4.Image = card24
                                compCard(3) = 11
                                card(3) = 24
                            Case 25
                                Me.picCompCard4.Image = card25
                                compCard(3) = 12
                                card(3) = 25
                            Case 26
                                Me.picCompCard4.Image = card26
                                compCard(3) = 13
                                card(3) = 26
                            Case 27
                                Me.picCompCard4.Image = card27
                                compCard(3) = 1
                                card(3) = 27
                            Case 28
                                Me.picCompCard4.Image = card28
                                compCard(3) = 2
                                card(3) = 28
                            Case 29
                                Me.picCompCard4.Image = card29
                                compCard(3) = 3
                                card(3) = 29
                            Case 30
                                Me.picCompCard4.Image = card30
                                compCard(3) = 4
                                card(3) = 30
                            Case 31
                                Me.picCompCard4.Image = card31
                                compCard(3) = 5
                                card(3) = 31
                            Case 32
                                Me.picCompCard4.Image = card32
                                compCard(3) = 6
                                card(3) = 32
                            Case 33
                                Me.picCompCard4.Image = card33
                                compCard(3) = 7
                                card(3) = 33
                            Case 34
                                Me.picCompCard4.Image = card34
                                compCard(3) = 8
                                card(3) = 34
                            Case 35
                                Me.picCompCard4.Image = card35
                                compCard(3) = 9
                                card(3) = 35
                            Case 36
                                Me.picCompCard4.Image = card36
                                compCard(3) = 10
                                card(3) = 36
                            Case 37
                                Me.picCompCard4.Image = card37
                                compCard(3) = 11
                                card(3) = 37
                            Case 38
                                Me.picCompCard4.Image = card38
                                compCard(3) = 12
                                card(3) = 38
                            Case 39
                                Me.picCompCard4.Image = card39
                                compCard(3) = 13
                                card(3) = 39
                            Case 40
                                Me.picCompCard4.Image = card40
                                compCard(3) = 1
                                card(3) = 40
                            Case 41
                                Me.picCompCard4.Image = card41
                                compCard(3) = 2
                                card(3) = 41
                            Case 42
                                Me.picCompCard4.Image = card42
                                compCard(3) = 3
                                card(3) = 42
                            Case 43
                                Me.picCompCard4.Image = card43
                                compCard(3) = 4
                                card(3) = 43
                            Case 44
                                Me.picCompCard4.Image = card44
                                compCard(3) = 5
                                card(3) = 44
                            Case 45
                                Me.picCompCard4.Image = card45
                                compCard(3) = 6
                                card(3) = 45
                            Case 46
                                Me.picCompCard4.Image = card46
                                compCard(3) = 7
                                card(3) = 46
                            Case 47
                                Me.picCompCard4.Image = card47
                                compCard(3) = 8
                                card(3) = 47
                            Case 48
                                Me.picCompCard4.Image = card48
                                compCard(3) = 9
                                card(3) = 48
                            Case 49
                                Me.picCompCard4.Image = card49
                                compCard(3) = 10
                                card(3) = 49
                            Case 50
                                Me.picCompCard4.Image = card50
                                compCard(3) = 11
                                card(3) = 50
                            Case 51
                                Me.picCompCard4.Image = card51
                                compCard(3) = 12
                                card(3) = 51
                            Case 52
                                Me.picCompCard4.Image = card52
                                compCard(3) = 13
                                card(3) = 52
                        End Select
                    ElseIf card(11) = card(4) Then
                        Select Case newNum
                            Case 1
                                Me.picCompCard5.Image = card1
                                compCard(4) = 1
                                card(4) = 1
                            Case 2
                                Me.picCompCard5.Image = card2
                                compCard(4) = 2
                                card(4) = 2
                            Case 3
                                Me.picCompCard5.Image = card3
                                compCard(4) = 3
                                card(4) = 3
                            Case 4
                                Me.picCompCard5.Image = card4
                                compCard(4) = 4
                                card(4) = 4
                            Case 5
                                Me.picCompCard5.Image = card5
                                compCard(4) = 5
                                card(4) = 5
                            Case 6
                                Me.picCompCard5.Image = card6
                                compCard(4) = 6
                                card(4) = 6
                            Case 7
                                Me.picCompCard5.Image = card7
                                compCard(4) = 7
                                card(4) = 7
                            Case 8
                                Me.picCompCard5.Image = card8
                                compCard(4) = 8
                                card(4) = 8
                            Case 9
                                Me.picCompCard5.Image = card9
                                compCard(4) = 9
                                card(4) = 9
                            Case 10
                                Me.picCompCard5.Image = card10
                                compCard(4) = 10
                                card(4) = 10
                            Case 11
                                Me.picCompCard5.Image = card11
                                compCard(4) = 11
                                card(4) = 11
                            Case 12
                                Me.picCompCard5.Image = card12
                                compCard(4) = 12
                                card(4) = 12
                            Case 13
                                Me.picCompCard5.Image = card13
                                compCard(4) = 13
                                card(4) = 13
                            Case 14
                                Me.picCompCard5.Image = card14
                                compCard(4) = 1
                                card(4) = 14
                            Case 15
                                Me.picCompCard5.Image = card15
                                compCard(4) = 2
                                card(4) = 15
                            Case 16
                                Me.picCompCard5.Image = card16
                                compCard(4) = 3
                                card(4) = 16
                            Case 17
                                Me.picCompCard5.Image = card17
                                compCard(4) = 4
                                card(4) = 17
                            Case 18
                                Me.picCompCard5.Image = card18
                                compCard(4) = 5
                                card(4) = 18
                            Case 19
                                Me.picCompCard5.Image = card19
                                compCard(4) = 6
                                card(4) = 19
                            Case 20
                                Me.picCompCard5.Image = card20
                                compCard(4) = 7
                                card(4) = 20
                            Case 21
                                Me.picCompCard5.Image = card21
                                compCard(4) = 8
                                card(4) = 21
                            Case 22
                                Me.picCompCard5.Image = card22
                                compCard(4) = 9
                                card(4) = 22
                            Case 23
                                Me.picCompCard5.Image = card23
                                compCard(4) = 10
                                card(4) = 23
                            Case 24
                                Me.picCompCard5.Image = card24
                                compCard(4) = 11
                                card(4) = 24
                            Case 25
                                Me.picCompCard5.Image = card25
                                compCard(4) = 12
                                card(4) = 25
                            Case 26
                                Me.picCompCard5.Image = card26
                                compCard(4) = 13
                                card(4) = 26
                            Case 27
                                Me.picCompCard5.Image = card27
                                compCard(4) = 1
                                card(4) = 27
                            Case 28
                                Me.picCompCard5.Image = card28
                                compCard(4) = 2
                                card(4) = 28
                            Case 29
                                Me.picCompCard5.Image = card29
                                compCard(4) = 3
                                card(4) = 29
                            Case 30
                                Me.picCompCard5.Image = card30
                                compCard(4) = 4
                                card(4) = 30
                            Case 31
                                Me.picCompCard5.Image = card31
                                compCard(4) = 5
                                card(4) = 31
                            Case 32
                                Me.picCompCard5.Image = card32
                                compCard(4) = 6
                                card(4) = 32
                            Case 33
                                Me.picCompCard5.Image = card33
                                compCard(4) = 7
                                card(4) = 33
                            Case 34
                                Me.picCompCard5.Image = card34
                                compCard(4) = 8
                                card(4) = 34
                            Case 35
                                Me.picCompCard5.Image = card35
                                compCard(4) = 9
                                card(4) = 35
                            Case 36
                                Me.picCompCard5.Image = card36
                                compCard(4) = 10
                                card(4) = 36
                            Case 37
                                Me.picCompCard5.Image = card37
                                compCard(4) = 11
                                card(4) = 37
                            Case 38
                                Me.picCompCard5.Image = card38
                                compCard(4) = 12
                                card(4) = 38
                            Case 39
                                Me.picCompCard5.Image = card39
                                compCard(4) = 13
                                card(4) = 39
                            Case 40
                                Me.picCompCard5.Image = card40
                                compCard(4) = 1
                                card(4) = 40
                            Case 41
                                Me.picCompCard5.Image = card41
                                compCard(4) = 2
                                card(4) = 41
                            Case 42
                                Me.picCompCard5.Image = card42
                                compCard(4) = 3
                                card(4) = 42
                            Case 43
                                Me.picCompCard5.Image = card43
                                compCard(4) = 4
                                card(4) = 43
                            Case 44
                                Me.picCompCard5.Image = card44
                                compCard(4) = 5
                                card(4) = 44
                            Case 45
                                Me.picCompCard5.Image = card45
                                compCard(4) = 6
                                card(4) = 45
                            Case 46
                                Me.picCompCard5.Image = card46
                                compCard(4) = 7
                                card(4) = 46
                            Case 47
                                Me.picCompCard5.Image = card47
                                compCard(4) = 8
                                card(4) = 47
                            Case 48
                                Me.picCompCard5.Image = card48
                                compCard(4) = 9
                                card(4) = 48
                            Case 49
                                Me.picCompCard5.Image = card49
                                compCard(4) = 10
                                card(4) = 49
                            Case 50
                                Me.picCompCard5.Image = card50
                                compCard(4) = 11
                                card(4) = 50
                            Case 51
                                Me.picCompCard5.Image = card51
                                compCard(4) = 12
                                card(4) = 51
                            Case 52
                                Me.picCompCard5.Image = card52
                                compCard(4) = 13
                                card(4) = 52
                        End Select
                    End If

                    Used.Add(newNum)

                End If
            Loop While OK = False
        End If
        compCardsLeftTotal = compCardsLeft1 + compCardsLeft2 + compCardsLeft3 + compCardsLeft4 + compCardsLeft5
        Me.lblCompCardsLeftTotal.Text = "Computer's" & vbCrLf & "Total Cards Left:" & vbCrLf & compCardsLeftTotal
        Me.lblCompCardsLeft1.Text = "Cards Left: " & compCardsLeft1
        Me.lblCompCardsLeft2.Text = "Cards Left: " & compCardsLeft2
        Me.lblCompCardsLeft3.Text = "Cards Left: " & compCardsLeft3
        Me.lblCompCardsLeft4.Text = "Cards Left: " & compCardsLeft4
        Me.lblCompCardsLeft5.Text = "Cards Left: " & compCardsLeft5

    End Sub

    Private Sub tmrHard_Tick(sender As Object, e As EventArgs) Handles tmrHard.Tick
        Dim newNum As Integer
        Dim OK As Boolean = False

        For cardNum As Integer = 0 To 4                                                                'check for compCard [0 to 4]
            If compPile = compCard(cardNum) + 1 Or compPile = compCard(cardNum) - 1 Then
                Select Case cardNum
                    Case 0                                                  'if compCard(0) is 1 higher or 1 lower than compPile
                        Me.picCompPile.Image = Me.picCompCard1.Image        'set picCompPile to picCompCard1 image
                        compPile = compCard(0)                              'set compPile to compCard(0) integer [1 to 13]
                        card(10) = card(0)                                  'set card(10), compPileCard, to card(0) integer [1 to 52]
                        If compCardsLeft1 > 0 Then                          'if 1 or more cards left, 
                            compCardsLeft1 -= 1                             'subtract 1
                        End If
                    Case 1
                        Me.picCompPile.Image = Me.picCompCard2.Image
                        compPile = compCard(1)
                        card(10) = card(1)
                        If compCardsLeft2 > 0 Then
                            compCardsLeft2 -= 1
                        End If
                    Case 2
                        Me.picCompPile.Image = Me.picCompCard3.Image
                        compPile = compCard(2)
                        card(10) = card(2)
                        If compCardsLeft3 > 0 Then
                            compCardsLeft3 -= 1
                        End If
                    Case 3
                        Me.picCompPile.Image = Me.picCompCard4.Image
                        compPile = compCard(3)
                        card(10) = card(3)
                        If compCardsLeft4 > 0 Then
                            compCardsLeft4 -= 1
                        End If
                    Case 4
                        Me.picCompPile.Image = Me.picCompCard5.Image
                        compPile = compCard(4)
                        card(10) = card(4)
                        If compCardsLeft5 > 0 Then
                            compCardsLeft5 -= 1
                        End If
                End Select
                cardNum = 100

            ElseIf compPile = 1 And (compCard(cardNum) = 2 Or compCard(cardNum) = 13) Then
                Select Case cardNum
                    Case 0
                        Me.picCompPile.Image = Me.picCompCard1.Image
                        compPile = compCard(0)
                        card(10) = card(0)
                        If compCardsLeft1 > 0 Then
                            compCardsLeft1 -= 1
                        End If
                    Case 1
                        Me.picCompPile.Image = Me.picCompCard2.Image
                        compPile = compCard(1)
                        card(10) = card(1)
                        If compCardsLeft2 > 0 Then
                            compCardsLeft2 -= 1
                        End If
                    Case 2
                        Me.picCompPile.Image = Me.picCompCard3.Image
                        compPile = compCard(2)
                        card(10) = card(2)
                        If compCardsLeft3 > 0 Then
                            compCardsLeft3 -= 1
                        End If
                    Case 3
                        Me.picCompPile.Image = Me.picCompCard4.Image
                        compPile = compCard(3)
                        card(10) = card(3)
                        If compCardsLeft4 > 0 Then
                            compCardsLeft4 -= 1
                        End If
                    Case 4
                        Me.picCompPile.Image = Me.picCompCard5.Image
                        compPile = compCard(4)
                        card(10) = card(4)
                        If compCardsLeft5 > 0 Then
                            compCardsLeft5 -= 1
                        End If
                End Select
                cardNum = 100

            ElseIf compPile = 13 And (compCard(cardNum) = 1 Or compCard(cardNum) = 12) Then
                Select Case cardNum
                    Case 0
                        Me.picCompPile.Image = Me.picCompCard1.Image
                        compPile = compCard(0)
                        card(10) = card(0)
                        If compCardsLeft1 > 0 Then
                            compCardsLeft1 -= 1
                        End If
                    Case 1
                        Me.picCompPile.Image = Me.picCompCard2.Image
                        compPile = compCard(1)
                        card(10) = card(1)
                        If compCardsLeft2 > 0 Then
                            compCardsLeft2 -= 1
                        End If
                    Case 2
                        Me.picCompPile.Image = Me.picCompCard3.Image
                        compPile = compCard(2)
                        card(10) = card(2)
                        If compCardsLeft3 > 0 Then
                            compCardsLeft3 -= 1
                        End If
                    Case 3
                        Me.picCompPile.Image = Me.picCompCard4.Image
                        compPile = compCard(3)
                        card(10) = card(3)
                        If compCardsLeft4 > 0 Then
                            compCardsLeft4 -= 1
                        End If
                    Case 4
                        Me.picCompPile.Image = Me.picCompCard5.Image
                        compPile = compCard(4)
                        card(10) = card(4)
                        If compCardsLeft5 > 0 Then
                            compCardsLeft5 -= 1
                        End If
                End Select
                cardNum = 100

            ElseIf playerPile = compCard(cardNum) + 1 Or playerPile = compCard(cardNum) - 1 Then
                Select Case cardNum
                    Case 0
                        Me.picPlayerPile.Image = Me.picCompCard1.Image
                        playerPile = compCard(0)
                        card(11) = card(0)
                        If compCardsLeft1 > 0 Then
                            compCardsLeft1 -= 1
                        End If
                    Case 1
                        Me.picPlayerPile.Image = Me.picCompCard2.Image
                        playerPile = compCard(1)
                        card(11) = card(1)
                        If compCardsLeft2 > 0 Then
                            compCardsLeft2 -= 1
                        End If
                    Case 2
                        Me.picPlayerPile.Image = Me.picCompCard3.Image
                        playerPile = compCard(2)
                        card(11) = card(2)
                        If compCardsLeft3 > 0 Then
                            compCardsLeft3 -= 1
                        End If
                    Case 3
                        Me.picPlayerPile.Image = Me.picCompCard4.Image
                        playerPile = compCard(3)
                        card(11) = card(3)
                        If compCardsLeft4 > 0 Then
                            compCardsLeft4 -= 1
                        End If
                    Case 4
                        Me.picPlayerPile.Image = Me.picCompCard5.Image
                        playerPile = compCard(4)
                        card(11) = card(4)
                        If compCardsLeft5 > 0 Then
                            compCardsLeft5 -= 1
                        End If
                End Select
                cardNum = 100

            ElseIf playerPile = 1 And (compCard(cardNum) = 2 Or compCard(cardNum) = 13) Then
                Select Case cardNum
                    Case 0
                        Me.picPlayerPile.Image = Me.picCompCard1.Image
                        playerPile = compCard(0)
                        card(11) = card(0)
                        If compCardsLeft1 > 0 Then
                            compCardsLeft1 -= 1
                        End If
                    Case 1
                        Me.picPlayerPile.Image = Me.picCompCard2.Image
                        playerPile = compCard(1)
                        card(11) = card(1)
                        If compCardsLeft2 > 0 Then
                            compCardsLeft2 -= 1
                        End If
                    Case 2
                        Me.picPlayerPile.Image = Me.picCompCard3.Image
                        playerPile = compCard(2)
                        card(11) = card(2)
                        If compCardsLeft3 > 0 Then
                            compCardsLeft3 -= 1
                        End If
                    Case 3
                        Me.picPlayerPile.Image = Me.picCompCard4.Image
                        playerPile = compCard(3)
                        card(11) = card(3)
                        If compCardsLeft4 > 0 Then
                            compCardsLeft4 -= 1
                        End If
                    Case 4
                        Me.picPlayerPile.Image = Me.picCompCard5.Image
                        playerPile = compCard(4)
                        card(11) = card(4)
                        If compCardsLeft5 > 0 Then
                            compCardsLeft5 -= 1
                        End If
                End Select
                cardNum = 100

            ElseIf playerPile = 13 And (compCard(cardNum) = 1 Or compCard(cardNum) = 12) Then
                Select Case cardNum
                    Case 0
                        Me.picPlayerPile.Image = Me.picCompCard1.Image
                        playerPile = compCard(0)
                        card(11) = card(0)
                        If compCardsLeft1 > 0 Then
                            compCardsLeft1 -= 1
                        End If
                    Case 1
                        Me.picPlayerPile.Image = Me.picCompCard2.Image
                        playerPile = compCard(1)
                        card(11) = card(1)
                        If compCardsLeft2 > 0 Then
                            compCardsLeft2 -= 1
                        End If
                    Case 2
                        Me.picPlayerPile.Image = Me.picCompCard3.Image
                        playerPile = compCard(2)
                        card(11) = card(2)
                        If compCardsLeft3 > 0 Then
                            compCardsLeft3 -= 1
                        End If
                    Case 3
                        Me.picPlayerPile.Image = Me.picCompCard4.Image
                        playerPile = compCard(3)
                        card(11) = card(3)
                        If compCardsLeft4 > 0 Then
                            compCardsLeft4 -= 1
                        End If
                    Case 4
                        Me.picPlayerPile.Image = Me.picCompCard5.Image
                        playerPile = compCard(4)
                        card(11) = card(4)
                        If compCardsLeft5 > 0 Then
                            compCardsLeft5 -= 1
                        End If
                End Select
                cardNum = 100

            End If
        Next cardNum

        If card(10) = card(0) Or card(10) = card(1) Or card(10) = card(2) Or card(10) = card(3) Or card(10) = card(4) Then

            Do                                                      'generate a new card
                Randomize()
                newNum = Int(max * Rnd() + min)

                If Used.Contains(newNum) Then                       'if newNum has already been generated, do again
                    OK = False
                Else                                                'if newNum has not already been generated, show new card
                    OK = True

                    If card(10) = card(0) Then                      'if card(10), compPileCard, is the same as card(0), compPile1 integer
                        Select Case newNum
                            Case 1
                                Me.picCompCard1.Image = card1
                                compCard(0) = 1
                                card(0) = 1
                            Case 2
                                Me.picCompCard1.Image = card2
                                compCard(0) = 2
                                card(0) = 2
                            Case 3
                                Me.picCompCard1.Image = card3
                                compCard(0) = 3
                                card(0) = 3
                            Case 4
                                Me.picCompCard1.Image = card4
                                compCard(0) = 4
                                card(0) = 4
                            Case 5
                                Me.picCompCard1.Image = card5
                                compCard(0) = 5
                                card(0) = 5
                            Case 6
                                Me.picCompCard1.Image = card6
                                compCard(0) = 6
                                card(0) = 6
                            Case 7
                                Me.picCompCard1.Image = card7
                                compCard(0) = 7
                                card(0) = 7
                            Case 8
                                Me.picCompCard1.Image = card8
                                compCard(0) = 8
                                card(0) = 8
                            Case 9
                                Me.picCompCard1.Image = card9
                                compCard(0) = 9
                                card(0) = 9
                            Case 10
                                Me.picCompCard1.Image = card10
                                compCard(0) = 10
                                card(0) = 10
                            Case 11
                                Me.picCompCard1.Image = card11
                                compCard(0) = 11
                                card(0) = 11
                            Case 12
                                Me.picCompCard1.Image = card12
                                compCard(0) = 12
                                card(0) = 12
                            Case 13
                                Me.picCompCard1.Image = card13
                                compCard(0) = 13
                                card(0) = 13
                            Case 14
                                Me.picCompCard1.Image = card14
                                compCard(0) = 1
                                card(0) = 14
                            Case 15
                                Me.picCompCard1.Image = card15
                                compCard(0) = 2
                                card(0) = 15
                            Case 16
                                Me.picCompCard1.Image = card16
                                compCard(0) = 3
                                card(0) = 16
                            Case 17
                                Me.picCompCard1.Image = card17
                                compCard(0) = 4
                                card(0) = 17
                            Case 18
                                Me.picCompCard1.Image = card18
                                compCard(0) = 5
                                card(0) = 18
                            Case 19
                                Me.picCompCard1.Image = card19
                                compCard(0) = 6
                                card(0) = 19
                            Case 20
                                Me.picCompCard1.Image = card20
                                compCard(0) = 7
                                card(0) = 20
                            Case 21
                                Me.picCompCard1.Image = card21
                                compCard(0) = 8
                                card(0) = 21
                            Case 22
                                Me.picCompCard1.Image = card22
                                compCard(0) = 9
                                card(0) = 22
                            Case 23
                                Me.picCompCard1.Image = card23
                                compCard(0) = 10
                                card(0) = 23
                            Case 24
                                Me.picCompCard1.Image = card24
                                compCard(0) = 11
                                card(0) = 24
                            Case 25
                                Me.picCompCard1.Image = card25
                                compCard(0) = 12
                                card(0) = 25
                            Case 26
                                Me.picCompCard1.Image = card26
                                compCard(0) = 13
                                card(0) = 26
                            Case 27
                                Me.picCompCard1.Image = card27
                                compCard(0) = 1
                                card(0) = 27
                            Case 28
                                Me.picCompCard1.Image = card28
                                compCard(0) = 2
                                card(0) = 28
                            Case 29
                                Me.picCompCard1.Image = card29
                                compCard(0) = 3
                                card(0) = 29
                            Case 30
                                Me.picCompCard1.Image = card30
                                compCard(0) = 4
                                card(0) = 30
                            Case 31
                                Me.picCompCard1.Image = card31
                                compCard(0) = 5
                                card(0) = 31
                            Case 32
                                Me.picCompCard1.Image = card32
                                compCard(0) = 6
                                card(0) = 32
                            Case 33
                                Me.picCompCard1.Image = card33
                                compCard(0) = 7
                                card(0) = 33
                            Case 34
                                Me.picCompCard1.Image = card34
                                compCard(0) = 8
                                card(0) = 34
                            Case 35
                                Me.picCompCard1.Image = card35
                                compCard(0) = 9
                                card(0) = 35
                            Case 36
                                Me.picCompCard1.Image = card36
                                compCard(0) = 10
                                card(0) = 36
                            Case 37
                                Me.picCompCard1.Image = card37
                                compCard(0) = 11
                                card(0) = 37
                            Case 38
                                Me.picCompCard1.Image = card38
                                compCard(0) = 12
                                card(0) = 38
                            Case 39
                                Me.picCompCard1.Image = card39
                                compCard(0) = 13
                                card(0) = 39
                            Case 40
                                Me.picCompCard1.Image = card40
                                compCard(0) = 1
                                card(0) = 40
                            Case 41
                                Me.picCompCard1.Image = card41
                                compCard(0) = 2
                                card(0) = 41
                            Case 42
                                Me.picCompCard1.Image = card42
                                compCard(0) = 3
                                card(0) = 42
                            Case 43
                                Me.picCompCard1.Image = card43
                                compCard(0) = 4
                                card(0) = 43
                            Case 44
                                Me.picCompCard1.Image = card44
                                compCard(0) = 5
                                card(0) = 44
                            Case 45
                                Me.picCompCard1.Image = card45
                                compCard(0) = 6
                                card(0) = 45
                            Case 46
                                Me.picCompCard1.Image = card46
                                compCard(0) = 7
                                card(0) = 46
                            Case 47
                                Me.picCompCard1.Image = card47
                                compCard(0) = 8
                                card(0) = 47
                            Case 48
                                Me.picCompCard1.Image = card48
                                compCard(0) = 9
                                card(0) = 48
                            Case 49
                                Me.picCompCard1.Image = card49
                                compCard(0) = 10
                                card(0) = 49
                            Case 50
                                Me.picCompCard1.Image = card50
                                compCard(0) = 11
                                card(0) = 50
                            Case 51
                                Me.picCompCard1.Image = card51
                                compCard(0) = 12
                                card(0) = 51
                            Case 52
                                Me.picCompCard1.Image = card52
                                compCard(0) = 13
                                card(0) = 52
                        End Select
                    ElseIf card(10) = card(1) Then
                        Select Case newNum
                            Case 1
                                Me.picCompCard2.Image = card1
                                compCard(1) = 1
                                card(1) = 1
                            Case 2
                                Me.picCompCard2.Image = card2
                                compCard(1) = 2
                                card(1) = 2
                            Case 3
                                Me.picCompCard2.Image = card3
                                compCard(1) = 3
                                card(1) = 3
                            Case 4
                                Me.picCompCard2.Image = card4
                                compCard(1) = 4
                                card(1) = 4
                            Case 5
                                Me.picCompCard2.Image = card5
                                compCard(1) = 5
                                card(1) = 5
                            Case 6
                                Me.picCompCard2.Image = card6
                                compCard(1) = 6
                                card(1) = 6
                            Case 7
                                Me.picCompCard2.Image = card7
                                compCard(1) = 7
                                card(1) = 7
                            Case 8
                                Me.picCompCard2.Image = card8
                                compCard(1) = 8
                                card(1) = 8
                            Case 9
                                Me.picCompCard2.Image = card9
                                compCard(1) = 9
                                card(1) = 9
                            Case 10
                                Me.picCompCard2.Image = card10
                                compCard(1) = 10
                                card(1) = 10
                            Case 11
                                Me.picCompCard2.Image = card11
                                compCard(1) = 11
                                card(1) = 11
                            Case 12
                                Me.picCompCard2.Image = card12
                                compCard(1) = 12
                                card(1) = 12
                            Case 13
                                Me.picCompCard2.Image = card13
                                compCard(1) = 13
                                card(1) = 13
                            Case 14
                                Me.picCompCard2.Image = card14
                                compCard(1) = 1
                                card(1) = 14
                            Case 15
                                Me.picCompCard2.Image = card15
                                compCard(1) = 2
                                card(1) = 15
                            Case 16
                                Me.picCompCard2.Image = card16
                                compCard(1) = 3
                                card(1) = 16
                            Case 17
                                Me.picCompCard2.Image = card17
                                compCard(1) = 4
                                card(1) = 17
                            Case 18
                                Me.picCompCard2.Image = card18
                                compCard(1) = 5
                                card(1) = 18
                            Case 19
                                Me.picCompCard2.Image = card19
                                compCard(1) = 6
                                card(1) = 19
                            Case 20
                                Me.picCompCard2.Image = card20
                                compCard(1) = 7
                                card(1) = 20
                            Case 21
                                Me.picCompCard2.Image = card21
                                compCard(1) = 8
                                card(1) = 21
                            Case 22
                                Me.picCompCard2.Image = card22
                                compCard(1) = 9
                                card(1) = 22
                            Case 23
                                Me.picCompCard2.Image = card23
                                compCard(1) = 10
                                card(1) = 23
                            Case 24
                                Me.picCompCard2.Image = card24
                                compCard(1) = 11
                                card(1) = 24
                            Case 25
                                Me.picCompCard2.Image = card25
                                compCard(1) = 12
                                card(1) = 25
                            Case 26
                                Me.picCompCard2.Image = card26
                                compCard(1) = 13
                                card(1) = 26
                            Case 27
                                Me.picCompCard2.Image = card27
                                compCard(1) = 1
                                card(1) = 27
                            Case 28
                                Me.picCompCard2.Image = card28
                                compCard(1) = 2
                                card(1) = 28
                            Case 29
                                Me.picCompCard2.Image = card29
                                compCard(1) = 3
                                card(1) = 29
                            Case 30
                                Me.picCompCard2.Image = card30
                                compCard(1) = 4
                                card(1) = 30
                            Case 31
                                Me.picCompCard2.Image = card31
                                compCard(1) = 5
                                card(1) = 31
                            Case 32
                                Me.picCompCard2.Image = card32
                                compCard(1) = 6
                                card(1) = 32
                            Case 33
                                Me.picCompCard2.Image = card33
                                compCard(1) = 7
                                card(1) = 33
                            Case 34
                                Me.picCompCard2.Image = card34
                                compCard(1) = 8
                                card(1) = 34
                            Case 35
                                Me.picCompCard2.Image = card35
                                compCard(1) = 9
                                card(1) = 35
                            Case 36
                                Me.picCompCard2.Image = card36
                                compCard(1) = 10
                                card(1) = 36
                            Case 37
                                Me.picCompCard2.Image = card37
                                compCard(1) = 11
                                card(1) = 37
                            Case 38
                                Me.picCompCard2.Image = card38
                                compCard(1) = 12
                                card(1) = 38
                            Case 39
                                Me.picCompCard2.Image = card39
                                compCard(1) = 13
                                card(1) = 39
                            Case 40
                                Me.picCompCard2.Image = card40
                                compCard(1) = 1
                                card(1) = 40
                            Case 41
                                Me.picCompCard2.Image = card41
                                compCard(1) = 2
                                card(1) = 41
                            Case 42
                                Me.picCompCard2.Image = card42
                                compCard(1) = 3
                                card(1) = 42
                            Case 43
                                Me.picCompCard2.Image = card43
                                compCard(1) = 4
                                card(1) = 43
                            Case 44
                                Me.picCompCard2.Image = card44
                                compCard(1) = 5
                                card(1) = 44
                            Case 45
                                Me.picCompCard2.Image = card45
                                compCard(1) = 6
                                card(1) = 45
                            Case 46
                                Me.picCompCard2.Image = card46
                                compCard(1) = 7
                                card(1) = 46
                            Case 47
                                Me.picCompCard2.Image = card47
                                compCard(1) = 8
                                card(1) = 47
                            Case 48
                                Me.picCompCard2.Image = card48
                                compCard(1) = 9
                                card(1) = 48
                            Case 49
                                Me.picCompCard2.Image = card49
                                compCard(1) = 10
                                card(1) = 49
                            Case 50
                                Me.picCompCard2.Image = card50
                                compCard(1) = 11
                                card(1) = 50
                            Case 51
                                Me.picCompCard2.Image = card51
                                compCard(1) = 12
                                card(1) = 51
                            Case 52
                                Me.picCompCard2.Image = card52
                                compCard(1) = 13
                                card(1) = 52
                        End Select
                    ElseIf card(10) = card(2) Then
                        Select Case newNum
                            Case 1
                                Me.picCompCard3.Image = card1
                                compCard(2) = 1
                                card(2) = 1
                            Case 2
                                Me.picCompCard3.Image = card2
                                compCard(2) = 2
                                card(2) = 2
                            Case 3
                                Me.picCompCard3.Image = card3
                                compCard(2) = 3
                                card(2) = 3
                            Case 4
                                Me.picCompCard3.Image = card4
                                compCard(2) = 4
                                card(2) = 4
                            Case 5
                                Me.picCompCard3.Image = card5
                                compCard(2) = 5
                                card(2) = 5
                            Case 6
                                Me.picCompCard3.Image = card6
                                compCard(2) = 6
                                card(2) = 6
                            Case 7
                                Me.picCompCard3.Image = card7
                                compCard(2) = 7
                                card(2) = 7
                            Case 8
                                Me.picCompCard3.Image = card8
                                compCard(2) = 8
                                card(2) = 8
                            Case 9
                                Me.picCompCard3.Image = card9
                                compCard(2) = 9
                                card(2) = 9
                            Case 10
                                Me.picCompCard3.Image = card10
                                compCard(2) = 10
                                card(2) = 10
                            Case 11
                                Me.picCompCard3.Image = card11
                                compCard(2) = 11
                                card(2) = 11
                            Case 12
                                Me.picCompCard3.Image = card12
                                compCard(2) = 12
                                card(2) = 12
                            Case 13
                                Me.picCompCard3.Image = card13
                                compCard(2) = 13
                                card(2) = 13
                            Case 14
                                Me.picCompCard3.Image = card14
                                compCard(2) = 1
                                card(2) = 14
                            Case 15
                                Me.picCompCard3.Image = card15
                                compCard(2) = 2
                                card(2) = 15
                            Case 16
                                Me.picCompCard3.Image = card16
                                compCard(2) = 3
                                card(2) = 16
                            Case 17
                                Me.picCompCard3.Image = card17
                                compCard(2) = 4
                                card(2) = 17
                            Case 18
                                Me.picCompCard3.Image = card18
                                compCard(2) = 5
                                card(2) = 18
                            Case 19
                                Me.picCompCard3.Image = card19
                                compCard(2) = 6
                                card(2) = 19
                            Case 20
                                Me.picCompCard3.Image = card20
                                compCard(2) = 7
                                card(2) = 20
                            Case 21
                                Me.picCompCard3.Image = card21
                                compCard(2) = 8
                                card(2) = 21
                            Case 22
                                Me.picCompCard3.Image = card22
                                compCard(2) = 9
                                card(2) = 22
                            Case 23
                                Me.picCompCard3.Image = card23
                                compCard(2) = 10
                                card(2) = 23
                            Case 24
                                Me.picCompCard3.Image = card24
                                compCard(2) = 11
                                card(2) = 24
                            Case 25
                                Me.picCompCard3.Image = card25
                                compCard(2) = 12
                                card(2) = 25
                            Case 26
                                Me.picCompCard3.Image = card26
                                compCard(2) = 13
                                card(2) = 26
                            Case 27
                                Me.picCompCard3.Image = card27
                                compCard(2) = 1
                                card(2) = 27
                            Case 28
                                Me.picCompCard3.Image = card28
                                compCard(2) = 2
                                card(2) = 28
                            Case 29
                                Me.picCompCard3.Image = card29
                                compCard(2) = 3
                                card(2) = 29
                            Case 30
                                Me.picCompCard3.Image = card30
                                compCard(2) = 4
                                card(2) = 30
                            Case 31
                                Me.picCompCard3.Image = card31
                                compCard(2) = 5
                                card(2) = 31
                            Case 32
                                Me.picCompCard3.Image = card32
                                compCard(2) = 6
                                card(2) = 32
                            Case 33
                                Me.picCompCard3.Image = card33
                                compCard(2) = 7
                                card(2) = 33
                            Case 34
                                Me.picCompCard3.Image = card34
                                compCard(2) = 8
                                card(2) = 34
                            Case 35
                                Me.picCompCard3.Image = card35
                                compCard(2) = 9
                                card(2) = 35
                            Case 36
                                Me.picCompCard3.Image = card36
                                compCard(2) = 10
                                card(2) = 36
                            Case 37
                                Me.picCompCard3.Image = card37
                                compCard(2) = 11
                                card(2) = 37
                            Case 38
                                Me.picCompCard3.Image = card38
                                compCard(2) = 12
                                card(2) = 38
                            Case 39
                                Me.picCompCard3.Image = card39
                                compCard(2) = 13
                                card(2) = 39
                            Case 40
                                Me.picCompCard3.Image = card40
                                compCard(2) = 1
                                card(2) = 40
                            Case 41
                                Me.picCompCard3.Image = card41
                                compCard(2) = 2
                                card(2) = 41
                            Case 42
                                Me.picCompCard3.Image = card42
                                compCard(2) = 3
                                card(2) = 42
                            Case 43
                                Me.picCompCard3.Image = card43
                                compCard(2) = 4
                                card(2) = 43
                            Case 44
                                Me.picCompCard3.Image = card44
                                compCard(2) = 5
                                card(2) = 44
                            Case 45
                                Me.picCompCard3.Image = card45
                                compCard(2) = 6
                                card(2) = 45
                            Case 46
                                Me.picCompCard3.Image = card46
                                compCard(2) = 7
                                card(2) = 46
                            Case 47
                                Me.picCompCard3.Image = card47
                                compCard(2) = 8
                                card(2) = 47
                            Case 48
                                Me.picCompCard3.Image = card48
                                compCard(2) = 9
                                card(2) = 48
                            Case 49
                                Me.picCompCard3.Image = card49
                                compCard(2) = 10
                                card(2) = 49
                            Case 50
                                Me.picCompCard3.Image = card50
                                compCard(2) = 11
                                card(2) = 50
                            Case 51
                                Me.picCompCard3.Image = card51
                                compCard(2) = 12
                                card(2) = 51
                            Case 52
                                Me.picCompCard3.Image = card52
                                compCard(2) = 13
                                card(2) = 52
                        End Select
                    ElseIf card(10) = card(3) Then
                        Select Case newNum
                            Case 1
                                Me.picCompCard4.Image = card1
                                compCard(3) = 1
                                card(3) = 1
                            Case 2
                                Me.picCompCard4.Image = card2
                                compCard(3) = 2
                                card(3) = 2
                            Case 3
                                Me.picCompCard4.Image = card3
                                compCard(3) = 3
                                card(3) = 3
                            Case 4
                                Me.picCompCard4.Image = card4
                                compCard(3) = 4
                                card(3) = 4
                            Case 5
                                Me.picCompCard4.Image = card5
                                compCard(3) = 5
                                card(3) = 5
                            Case 6
                                Me.picCompCard4.Image = card6
                                compCard(3) = 6
                                card(3) = 6
                            Case 7
                                Me.picCompCard4.Image = card7
                                compCard(3) = 7
                                card(3) = 7
                            Case 8
                                Me.picCompCard4.Image = card8
                                compCard(3) = 8
                                card(3) = 8
                            Case 9
                                Me.picCompCard4.Image = card9
                                compCard(3) = 9
                                card(3) = 9
                            Case 10
                                Me.picCompCard4.Image = card10
                                compCard(3) = 10
                                card(3) = 10
                            Case 11
                                Me.picCompCard4.Image = card11
                                compCard(3) = 11
                                card(3) = 11
                            Case 12
                                Me.picCompCard4.Image = card12
                                compCard(3) = 12
                                card(3) = 12
                            Case 13
                                Me.picCompCard4.Image = card13
                                compCard(3) = 13
                                card(3) = 13
                            Case 14
                                Me.picCompCard4.Image = card14
                                compCard(3) = 1
                                card(3) = 14
                            Case 15
                                Me.picCompCard4.Image = card15
                                compCard(3) = 2
                                card(3) = 15
                            Case 16
                                Me.picCompCard4.Image = card16
                                compCard(3) = 3
                                card(3) = 16
                            Case 17
                                Me.picCompCard4.Image = card17
                                compCard(3) = 4
                                card(3) = 17
                            Case 18
                                Me.picCompCard4.Image = card18
                                compCard(3) = 5
                                card(3) = 18
                            Case 19
                                Me.picCompCard4.Image = card19
                                compCard(3) = 6
                                card(3) = 19
                            Case 20
                                Me.picCompCard4.Image = card20
                                compCard(3) = 7
                                card(3) = 20
                            Case 21
                                Me.picCompCard4.Image = card21
                                compCard(3) = 8
                                card(3) = 21
                            Case 22
                                Me.picCompCard4.Image = card22
                                compCard(3) = 9
                                card(3) = 22
                            Case 23
                                Me.picCompCard4.Image = card23
                                compCard(3) = 10
                                card(3) = 23
                            Case 24
                                Me.picCompCard4.Image = card24
                                compCard(3) = 11
                                card(3) = 24
                            Case 25
                                Me.picCompCard4.Image = card25
                                compCard(3) = 12
                                card(3) = 25
                            Case 26
                                Me.picCompCard4.Image = card26
                                compCard(3) = 13
                                card(3) = 26
                            Case 27
                                Me.picCompCard4.Image = card27
                                compCard(3) = 1
                                card(3) = 27
                            Case 28
                                Me.picCompCard4.Image = card28
                                compCard(3) = 2
                                card(3) = 28
                            Case 29
                                Me.picCompCard4.Image = card29
                                compCard(3) = 3
                                card(3) = 29
                            Case 30
                                Me.picCompCard4.Image = card30
                                compCard(3) = 4
                                card(3) = 30
                            Case 31
                                Me.picCompCard4.Image = card31
                                compCard(3) = 5
                                card(3) = 31
                            Case 32
                                Me.picCompCard4.Image = card32
                                compCard(3) = 6
                                card(3) = 32
                            Case 33
                                Me.picCompCard4.Image = card33
                                compCard(3) = 7
                                card(3) = 33
                            Case 34
                                Me.picCompCard4.Image = card34
                                compCard(3) = 8
                                card(3) = 34
                            Case 35
                                Me.picCompCard4.Image = card35
                                compCard(3) = 9
                                card(3) = 35
                            Case 36
                                Me.picCompCard4.Image = card36
                                compCard(3) = 10
                                card(3) = 36
                            Case 37
                                Me.picCompCard4.Image = card37
                                compCard(3) = 11
                                card(3) = 37
                            Case 38
                                Me.picCompCard4.Image = card38
                                compCard(3) = 12
                                card(3) = 38
                            Case 39
                                Me.picCompCard4.Image = card39
                                compCard(3) = 13
                                card(3) = 39
                            Case 40
                                Me.picCompCard4.Image = card40
                                compCard(3) = 1
                                card(3) = 40
                            Case 41
                                Me.picCompCard4.Image = card41
                                compCard(3) = 2
                                card(3) = 41
                            Case 42
                                Me.picCompCard4.Image = card42
                                compCard(3) = 3
                                card(3) = 42
                            Case 43
                                Me.picCompCard4.Image = card43
                                compCard(3) = 4
                                card(3) = 43
                            Case 44
                                Me.picCompCard4.Image = card44
                                compCard(3) = 5
                                card(3) = 44
                            Case 45
                                Me.picCompCard4.Image = card45
                                compCard(3) = 6
                                card(3) = 45
                            Case 46
                                Me.picCompCard4.Image = card46
                                compCard(3) = 7
                                card(3) = 46
                            Case 47
                                Me.picCompCard4.Image = card47
                                compCard(3) = 8
                                card(3) = 47
                            Case 48
                                Me.picCompCard4.Image = card48
                                compCard(3) = 9
                                card(3) = 48
                            Case 49
                                Me.picCompCard4.Image = card49
                                compCard(3) = 10
                                card(3) = 49
                            Case 50
                                Me.picCompCard4.Image = card50
                                compCard(3) = 11
                                card(3) = 50
                            Case 51
                                Me.picCompCard4.Image = card51
                                compCard(3) = 12
                                card(3) = 51
                            Case 52
                                Me.picCompCard4.Image = card52
                                compCard(3) = 13
                                card(3) = 52
                        End Select
                    ElseIf card(10) = card(4) Then
                        Select Case newNum
                            Case 1
                                Me.picCompCard5.Image = card1
                                compCard(4) = 1
                                card(4) = 1
                            Case 2
                                Me.picCompCard5.Image = card2
                                compCard(4) = 2
                                card(4) = 2
                            Case 3
                                Me.picCompCard5.Image = card3
                                compCard(4) = 3
                                card(4) = 3
                            Case 4
                                Me.picCompCard5.Image = card4
                                compCard(4) = 4
                                card(4) = 4
                            Case 5
                                Me.picCompCard5.Image = card5
                                compCard(4) = 5
                                card(4) = 5
                            Case 6
                                Me.picCompCard5.Image = card6
                                compCard(4) = 6
                                card(4) = 6
                            Case 7
                                Me.picCompCard5.Image = card7
                                compCard(4) = 7
                                card(4) = 7
                            Case 8
                                Me.picCompCard5.Image = card8
                                compCard(4) = 8
                                card(4) = 8
                            Case 9
                                Me.picCompCard5.Image = card9
                                compCard(4) = 9
                                card(4) = 9
                            Case 10
                                Me.picCompCard5.Image = card10
                                compCard(4) = 10
                                card(4) = 10
                            Case 11
                                Me.picCompCard5.Image = card11
                                compCard(4) = 11
                                card(4) = 11
                            Case 12
                                Me.picCompCard5.Image = card12
                                compCard(4) = 12
                                card(4) = 12
                            Case 13
                                Me.picCompCard5.Image = card13
                                compCard(4) = 13
                                card(4) = 13
                            Case 14
                                Me.picCompCard5.Image = card14
                                compCard(4) = 1
                                card(4) = 14
                            Case 15
                                Me.picCompCard5.Image = card15
                                compCard(4) = 2
                                card(4) = 15
                            Case 16
                                Me.picCompCard5.Image = card16
                                compCard(4) = 3
                                card(4) = 16
                            Case 17
                                Me.picCompCard5.Image = card17
                                compCard(4) = 4
                                card(4) = 17
                            Case 18
                                Me.picCompCard5.Image = card18
                                compCard(4) = 5
                                card(4) = 18
                            Case 19
                                Me.picCompCard5.Image = card19
                                compCard(4) = 6
                                card(4) = 19
                            Case 20
                                Me.picCompCard5.Image = card20
                                compCard(4) = 7
                                card(4) = 20
                            Case 21
                                Me.picCompCard5.Image = card21
                                compCard(4) = 8
                                card(4) = 21
                            Case 22
                                Me.picCompCard5.Image = card22
                                compCard(4) = 9
                                card(4) = 22
                            Case 23
                                Me.picCompCard5.Image = card23
                                compCard(4) = 10
                                card(4) = 23
                            Case 24
                                Me.picCompCard5.Image = card24
                                compCard(4) = 11
                                card(4) = 24
                            Case 25
                                Me.picCompCard5.Image = card25
                                compCard(4) = 12
                                card(4) = 25
                            Case 26
                                Me.picCompCard5.Image = card26
                                compCard(4) = 13
                                card(4) = 26
                            Case 27
                                Me.picCompCard5.Image = card27
                                compCard(4) = 1
                                card(4) = 27
                            Case 28
                                Me.picCompCard5.Image = card28
                                compCard(4) = 2
                                card(4) = 28
                            Case 29
                                Me.picCompCard5.Image = card29
                                compCard(4) = 3
                                card(4) = 29
                            Case 30
                                Me.picCompCard5.Image = card30
                                compCard(4) = 4
                                card(4) = 30
                            Case 31
                                Me.picCompCard5.Image = card31
                                compCard(4) = 5
                                card(4) = 31
                            Case 32
                                Me.picCompCard5.Image = card32
                                compCard(4) = 6
                                card(4) = 32
                            Case 33
                                Me.picCompCard5.Image = card33
                                compCard(4) = 7
                                card(4) = 33
                            Case 34
                                Me.picCompCard5.Image = card34
                                compCard(4) = 8
                                card(4) = 34
                            Case 35
                                Me.picCompCard5.Image = card35
                                compCard(4) = 9
                                card(4) = 35
                            Case 36
                                Me.picCompCard5.Image = card36
                                compCard(4) = 10
                                card(4) = 36
                            Case 37
                                Me.picCompCard5.Image = card37
                                compCard(4) = 11
                                card(4) = 37
                            Case 38
                                Me.picCompCard5.Image = card38
                                compCard(4) = 12
                                card(4) = 38
                            Case 39
                                Me.picCompCard5.Image = card39
                                compCard(4) = 13
                                card(4) = 39
                            Case 40
                                Me.picCompCard5.Image = card40
                                compCard(4) = 1
                                card(4) = 40
                            Case 41
                                Me.picCompCard5.Image = card41
                                compCard(4) = 2
                                card(4) = 41
                            Case 42
                                Me.picCompCard5.Image = card42
                                compCard(4) = 3
                                card(4) = 42
                            Case 43
                                Me.picCompCard5.Image = card43
                                compCard(4) = 4
                                card(4) = 43
                            Case 44
                                Me.picCompCard5.Image = card44
                                compCard(4) = 5
                                card(4) = 44
                            Case 45
                                Me.picCompCard5.Image = card45
                                compCard(4) = 6
                                card(4) = 45
                            Case 46
                                Me.picCompCard5.Image = card46
                                compCard(4) = 7
                                card(4) = 46
                            Case 47
                                Me.picCompCard5.Image = card47
                                compCard(4) = 8
                                card(4) = 47
                            Case 48
                                Me.picCompCard5.Image = card48
                                compCard(4) = 9
                                card(4) = 48
                            Case 49
                                Me.picCompCard5.Image = card49
                                compCard(4) = 10
                                card(4) = 49
                            Case 50
                                Me.picCompCard5.Image = card50
                                compCard(4) = 11
                                card(4) = 50
                            Case 51
                                Me.picCompCard5.Image = card51
                                compCard(4) = 12
                                card(4) = 51
                            Case 52
                                Me.picCompCard5.Image = card52
                                compCard(4) = 13
                                card(4) = 52
                        End Select
                    End If

                    Used.Add(newNum)

                End If
            Loop While OK = False
        End If

        If card(11) = card(0) Or card(11) = card(1) Or card(11) = card(2) Or card(11) = card(3) Or card(11) = card(4) Then
            Do                                          'generate a new unique random number
                Randomize()
                newNum = Int(max * Rnd() + min)

                If Used.Contains(newNum) Then
                    OK = False
                Else
                    OK = True

                    If card(11) = card(0) Then
                        Select Case newNum
                            Case 1
                                Me.picCompCard1.Image = card1
                                compCard(0) = 1
                                card(0) = 1
                            Case 2
                                Me.picCompCard1.Image = card2
                                compCard(0) = 2
                                card(0) = 2
                            Case 3
                                Me.picCompCard1.Image = card3
                                compCard(0) = 3
                                card(0) = 3
                            Case 4
                                Me.picCompCard1.Image = card4
                                compCard(0) = 4
                                card(0) = 4
                            Case 5
                                Me.picCompCard1.Image = card5
                                compCard(0) = 5
                                card(0) = 5
                            Case 6
                                Me.picCompCard1.Image = card6
                                compCard(0) = 6
                                card(0) = 6
                            Case 7
                                Me.picCompCard1.Image = card7
                                compCard(0) = 7
                                card(0) = 7
                            Case 8
                                Me.picCompCard1.Image = card8
                                compCard(0) = 8
                                card(0) = 8
                            Case 9
                                Me.picCompCard1.Image = card9
                                compCard(0) = 9
                                card(0) = 9
                            Case 10
                                Me.picCompCard1.Image = card10
                                compCard(0) = 10
                                card(0) = 10
                            Case 11
                                Me.picCompCard1.Image = card11
                                compCard(0) = 11
                                card(0) = 11
                            Case 12
                                Me.picCompCard1.Image = card12
                                compCard(0) = 12
                                card(0) = 12
                            Case 13
                                Me.picCompCard1.Image = card13
                                compCard(0) = 13
                                card(0) = 13
                            Case 14
                                Me.picCompCard1.Image = card14
                                compCard(0) = 1
                                card(0) = 14
                            Case 15
                                Me.picCompCard1.Image = card15
                                compCard(0) = 2
                                card(0) = 15
                            Case 16
                                Me.picCompCard1.Image = card16
                                compCard(0) = 3
                                card(0) = 16
                            Case 17
                                Me.picCompCard1.Image = card17
                                compCard(0) = 4
                                card(0) = 17
                            Case 18
                                Me.picCompCard1.Image = card18
                                compCard(0) = 5
                                card(0) = 18
                            Case 19
                                Me.picCompCard1.Image = card19
                                compCard(0) = 6
                                card(0) = 19
                            Case 20
                                Me.picCompCard1.Image = card20
                                compCard(0) = 7
                                card(0) = 20
                            Case 21
                                Me.picCompCard1.Image = card21
                                compCard(0) = 8
                                card(0) = 21
                            Case 22
                                Me.picCompCard1.Image = card22
                                compCard(0) = 9
                                card(0) = 22
                            Case 23
                                Me.picCompCard1.Image = card23
                                compCard(0) = 10
                                card(0) = 23
                            Case 24
                                Me.picCompCard1.Image = card24
                                compCard(0) = 11
                                card(0) = 24
                            Case 25
                                Me.picCompCard1.Image = card25
                                compCard(0) = 12
                                card(0) = 25
                            Case 26
                                Me.picCompCard1.Image = card26
                                compCard(0) = 13
                                card(0) = 26
                            Case 27
                                Me.picCompCard1.Image = card27
                                compCard(0) = 1
                                card(0) = 27
                            Case 28
                                Me.picCompCard1.Image = card28
                                compCard(0) = 2
                                card(0) = 28
                            Case 29
                                Me.picCompCard1.Image = card29
                                compCard(0) = 3
                                card(0) = 29
                            Case 30
                                Me.picCompCard1.Image = card30
                                compCard(0) = 4
                                card(0) = 30
                            Case 31
                                Me.picCompCard1.Image = card31
                                compCard(0) = 5
                                card(0) = 31
                            Case 32
                                Me.picCompCard1.Image = card32
                                compCard(0) = 6
                                card(0) = 32
                            Case 33
                                Me.picCompCard1.Image = card33
                                compCard(0) = 7
                                card(0) = 33
                            Case 34
                                Me.picCompCard1.Image = card34
                                compCard(0) = 8
                                card(0) = 34
                            Case 35
                                Me.picCompCard1.Image = card35
                                compCard(0) = 9
                                card(0) = 35
                            Case 36
                                Me.picCompCard1.Image = card36
                                compCard(0) = 10
                                card(0) = 36
                            Case 37
                                Me.picCompCard1.Image = card37
                                compCard(0) = 11
                                card(0) = 37
                            Case 38
                                Me.picCompCard1.Image = card38
                                compCard(0) = 12
                                card(0) = 38
                            Case 39
                                Me.picCompCard1.Image = card39
                                compCard(0) = 13
                                card(0) = 39
                            Case 40
                                Me.picCompCard1.Image = card40
                                compCard(0) = 1
                                card(0) = 40
                            Case 41
                                Me.picCompCard1.Image = card41
                                compCard(0) = 2
                                card(0) = 41
                            Case 42
                                Me.picCompCard1.Image = card42
                                compCard(0) = 3
                                card(0) = 42
                            Case 43
                                Me.picCompCard1.Image = card43
                                compCard(0) = 4
                                card(0) = 43
                            Case 44
                                Me.picCompCard1.Image = card44
                                compCard(0) = 5
                                card(0) = 44
                            Case 45
                                Me.picCompCard1.Image = card45
                                compCard(0) = 6
                                card(0) = 45
                            Case 46
                                Me.picCompCard1.Image = card46
                                compCard(0) = 7
                                card(0) = 46
                            Case 47
                                Me.picCompCard1.Image = card47
                                compCard(0) = 8
                                card(0) = 47
                            Case 48
                                Me.picCompCard1.Image = card48
                                compCard(0) = 9
                                card(0) = 48
                            Case 49
                                Me.picCompCard1.Image = card49
                                compCard(0) = 10
                                card(0) = 49
                            Case 50
                                Me.picCompCard1.Image = card50
                                compCard(0) = 11
                                card(0) = 50
                            Case 51
                                Me.picCompCard1.Image = card51
                                compCard(0) = 12
                                card(0) = 51
                            Case 52
                                Me.picCompCard1.Image = card52
                                compCard(0) = 13
                                card(0) = 52
                        End Select
                    ElseIf card(11) = card(1) Then
                        Select Case newNum
                            Case 1
                                Me.picCompCard2.Image = card1
                                compCard(1) = 1
                                card(1) = 1
                            Case 2
                                Me.picCompCard2.Image = card2
                                compCard(1) = 2
                                card(1) = 2
                            Case 3
                                Me.picCompCard2.Image = card3
                                compCard(1) = 3
                                card(1) = 3
                            Case 4
                                Me.picCompCard2.Image = card4
                                compCard(1) = 4
                                card(1) = 4
                            Case 5
                                Me.picCompCard2.Image = card5
                                compCard(1) = 5
                                card(1) = 5
                            Case 6
                                Me.picCompCard2.Image = card6
                                compCard(1) = 6
                                card(1) = 6
                            Case 7
                                Me.picCompCard2.Image = card7
                                compCard(1) = 7
                                card(1) = 7
                            Case 8
                                Me.picCompCard2.Image = card8
                                compCard(1) = 8
                                card(1) = 8
                            Case 9
                                Me.picCompCard2.Image = card9
                                compCard(1) = 9
                                card(1) = 9
                            Case 10
                                Me.picCompCard2.Image = card10
                                compCard(1) = 10
                                card(1) = 10
                            Case 11
                                Me.picCompCard2.Image = card11
                                compCard(1) = 11
                                card(1) = 11
                            Case 12
                                Me.picCompCard2.Image = card12
                                compCard(1) = 12
                                card(1) = 12
                            Case 13
                                Me.picCompCard2.Image = card13
                                compCard(1) = 13
                                card(1) = 13
                            Case 14
                                Me.picCompCard2.Image = card14
                                compCard(1) = 1
                                card(1) = 14
                            Case 15
                                Me.picCompCard2.Image = card15
                                compCard(1) = 2
                                card(1) = 15
                            Case 16
                                Me.picCompCard2.Image = card16
                                compCard(1) = 3
                                card(1) = 16
                            Case 17
                                Me.picCompCard2.Image = card17
                                compCard(1) = 4
                                card(1) = 17
                            Case 18
                                Me.picCompCard2.Image = card18
                                compCard(1) = 5
                                card(1) = 18
                            Case 19
                                Me.picCompCard2.Image = card19
                                compCard(1) = 6
                                card(1) = 19
                            Case 20
                                Me.picCompCard2.Image = card20
                                compCard(1) = 7
                                card(1) = 20
                            Case 21
                                Me.picCompCard2.Image = card21
                                compCard(1) = 8
                                card(1) = 21
                            Case 22
                                Me.picCompCard2.Image = card22
                                compCard(1) = 9
                                card(1) = 22
                            Case 23
                                Me.picCompCard2.Image = card23
                                compCard(1) = 10
                                card(1) = 23
                            Case 24
                                Me.picCompCard2.Image = card24
                                compCard(1) = 11
                                card(1) = 24
                            Case 25
                                Me.picCompCard2.Image = card25
                                compCard(1) = 12
                                card(1) = 25
                            Case 26
                                Me.picCompCard2.Image = card26
                                compCard(1) = 13
                                card(1) = 26
                            Case 27
                                Me.picCompCard2.Image = card27
                                compCard(1) = 1
                                card(1) = 27
                            Case 28
                                Me.picCompCard2.Image = card28
                                compCard(1) = 2
                                card(1) = 28
                            Case 29
                                Me.picCompCard2.Image = card29
                                compCard(1) = 3
                                card(1) = 29
                            Case 30
                                Me.picCompCard2.Image = card30
                                compCard(1) = 4
                                card(1) = 30
                            Case 31
                                Me.picCompCard2.Image = card31
                                compCard(1) = 5
                                card(1) = 31
                            Case 32
                                Me.picCompCard2.Image = card32
                                compCard(1) = 6
                                card(1) = 32
                            Case 33
                                Me.picCompCard2.Image = card33
                                compCard(1) = 7
                                card(1) = 33
                            Case 34
                                Me.picCompCard2.Image = card34
                                compCard(1) = 8
                                card(1) = 34
                            Case 35
                                Me.picCompCard2.Image = card35
                                compCard(1) = 9
                                card(1) = 35
                            Case 36
                                Me.picCompCard2.Image = card36
                                compCard(1) = 10
                                card(1) = 36
                            Case 37
                                Me.picCompCard2.Image = card37
                                compCard(1) = 11
                                card(1) = 37
                            Case 38
                                Me.picCompCard2.Image = card38
                                compCard(1) = 12
                                card(1) = 38
                            Case 39
                                Me.picCompCard2.Image = card39
                                compCard(1) = 13
                                card(1) = 39
                            Case 40
                                Me.picCompCard2.Image = card40
                                compCard(1) = 1
                                card(1) = 40
                            Case 41
                                Me.picCompCard2.Image = card41
                                compCard(1) = 2
                                card(1) = 41
                            Case 42
                                Me.picCompCard2.Image = card42
                                compCard(1) = 3
                                card(1) = 42
                            Case 43
                                Me.picCompCard2.Image = card43
                                compCard(1) = 4
                                card(1) = 43
                            Case 44
                                Me.picCompCard2.Image = card44
                                compCard(1) = 5
                                card(1) = 44
                            Case 45
                                Me.picCompCard2.Image = card45
                                compCard(1) = 6
                                card(1) = 45
                            Case 46
                                Me.picCompCard2.Image = card46
                                compCard(1) = 7
                                card(1) = 46
                            Case 47
                                Me.picCompCard2.Image = card47
                                compCard(1) = 8
                                card(1) = 47
                            Case 48
                                Me.picCompCard2.Image = card48
                                compCard(1) = 9
                                card(1) = 48
                            Case 49
                                Me.picCompCard2.Image = card49
                                compCard(1) = 10
                                card(1) = 49
                            Case 50
                                Me.picCompCard2.Image = card50
                                compCard(1) = 11
                                card(1) = 50
                            Case 51
                                Me.picCompCard2.Image = card51
                                compCard(1) = 12
                                card(1) = 51
                            Case 52
                                Me.picCompCard2.Image = card52
                                compCard(1) = 13
                                card(1) = 52
                        End Select
                    ElseIf card(11) = card(2) Then
                        Select Case newNum
                            Case 1
                                Me.picCompCard3.Image = card1
                                compCard(2) = 1
                                card(2) = 1
                            Case 2
                                Me.picCompCard3.Image = card2
                                compCard(2) = 2
                                card(2) = 2
                            Case 3
                                Me.picCompCard3.Image = card3
                                compCard(2) = 3
                                card(2) = 3
                            Case 4
                                Me.picCompCard3.Image = card4
                                compCard(2) = 4
                                card(2) = 4
                            Case 5
                                Me.picCompCard3.Image = card5
                                compCard(2) = 5
                                card(2) = 5
                            Case 6
                                Me.picCompCard3.Image = card6
                                compCard(2) = 6
                                card(2) = 6
                            Case 7
                                Me.picCompCard3.Image = card7
                                compCard(2) = 7
                                card(2) = 7
                            Case 8
                                Me.picCompCard3.Image = card8
                                compCard(2) = 8
                                card(2) = 8
                            Case 9
                                Me.picCompCard3.Image = card9
                                compCard(2) = 9
                                card(2) = 9
                            Case 10
                                Me.picCompCard3.Image = card10
                                compCard(2) = 10
                                card(2) = 10
                            Case 11
                                Me.picCompCard3.Image = card11
                                compCard(2) = 11
                                card(2) = 11
                            Case 12
                                Me.picCompCard3.Image = card12
                                compCard(2) = 12
                                card(2) = 12
                            Case 13
                                Me.picCompCard3.Image = card13
                                compCard(2) = 13
                                card(2) = 13
                            Case 14
                                Me.picCompCard3.Image = card14
                                compCard(2) = 1
                                card(2) = 14
                            Case 15
                                Me.picCompCard3.Image = card15
                                compCard(2) = 2
                                card(2) = 15
                            Case 16
                                Me.picCompCard3.Image = card16
                                compCard(2) = 3
                                card(2) = 16
                            Case 17
                                Me.picCompCard3.Image = card17
                                compCard(2) = 4
                                card(2) = 17
                            Case 18
                                Me.picCompCard3.Image = card18
                                compCard(2) = 5
                                card(2) = 18
                            Case 19
                                Me.picCompCard3.Image = card19
                                compCard(2) = 6
                                card(2) = 19
                            Case 20
                                Me.picCompCard3.Image = card20
                                compCard(2) = 7
                                card(2) = 20
                            Case 21
                                Me.picCompCard3.Image = card21
                                compCard(2) = 8
                                card(2) = 21
                            Case 22
                                Me.picCompCard3.Image = card22
                                compCard(2) = 9
                                card(2) = 22
                            Case 23
                                Me.picCompCard3.Image = card23
                                compCard(2) = 10
                                card(2) = 23
                            Case 24
                                Me.picCompCard3.Image = card24
                                compCard(2) = 11
                                card(2) = 24
                            Case 25
                                Me.picCompCard3.Image = card25
                                compCard(2) = 12
                                card(2) = 25
                            Case 26
                                Me.picCompCard3.Image = card26
                                compCard(2) = 13
                                card(2) = 26
                            Case 27
                                Me.picCompCard3.Image = card27
                                compCard(2) = 1
                                card(2) = 27
                            Case 28
                                Me.picCompCard3.Image = card28
                                compCard(2) = 2
                                card(2) = 28
                            Case 29
                                Me.picCompCard3.Image = card29
                                compCard(2) = 3
                                card(2) = 29
                            Case 30
                                Me.picCompCard3.Image = card30
                                compCard(2) = 4
                                card(2) = 30
                            Case 31
                                Me.picCompCard3.Image = card31
                                compCard(2) = 5
                                card(2) = 31
                            Case 32
                                Me.picCompCard3.Image = card32
                                compCard(2) = 6
                                card(2) = 32
                            Case 33
                                Me.picCompCard3.Image = card33
                                compCard(2) = 7
                                card(2) = 33
                            Case 34
                                Me.picCompCard3.Image = card34
                                compCard(2) = 8
                                card(2) = 34
                            Case 35
                                Me.picCompCard3.Image = card35
                                compCard(2) = 9
                                card(2) = 35
                            Case 36
                                Me.picCompCard3.Image = card36
                                compCard(2) = 10
                                card(2) = 36
                            Case 37
                                Me.picCompCard3.Image = card37
                                compCard(2) = 11
                                card(2) = 37
                            Case 38
                                Me.picCompCard3.Image = card38
                                compCard(2) = 12
                                card(2) = 38
                            Case 39
                                Me.picCompCard3.Image = card39
                                compCard(2) = 13
                                card(2) = 39
                            Case 40
                                Me.picCompCard3.Image = card40
                                compCard(2) = 1
                                card(2) = 40
                            Case 41
                                Me.picCompCard3.Image = card41
                                compCard(2) = 2
                                card(2) = 41
                            Case 42
                                Me.picCompCard3.Image = card42
                                compCard(2) = 3
                                card(2) = 42
                            Case 43
                                Me.picCompCard3.Image = card43
                                compCard(2) = 4
                                card(2) = 43
                            Case 44
                                Me.picCompCard3.Image = card44
                                compCard(2) = 5
                                card(2) = 44
                            Case 45
                                Me.picCompCard3.Image = card45
                                compCard(2) = 6
                                card(2) = 45
                            Case 46
                                Me.picCompCard3.Image = card46
                                compCard(2) = 7
                                card(2) = 46
                            Case 47
                                Me.picCompCard3.Image = card47
                                compCard(2) = 8
                                card(2) = 47
                            Case 48
                                Me.picCompCard3.Image = card48
                                compCard(2) = 9
                                card(2) = 48
                            Case 49
                                Me.picCompCard3.Image = card49
                                compCard(2) = 10
                                card(2) = 49
                            Case 50
                                Me.picCompCard3.Image = card50
                                compCard(2) = 11
                                card(2) = 50
                            Case 51
                                Me.picCompCard3.Image = card51
                                compCard(2) = 12
                                card(2) = 51
                            Case 52
                                Me.picCompCard3.Image = card52
                                compCard(2) = 13
                                card(2) = 52
                        End Select
                    ElseIf card(11) = card(3) Then
                        Select Case newNum
                            Case 1
                                Me.picCompCard4.Image = card1
                                compCard(3) = 1
                                card(3) = 1
                            Case 2
                                Me.picCompCard4.Image = card2
                                compCard(3) = 2
                                card(3) = 2
                            Case 3
                                Me.picCompCard4.Image = card3
                                compCard(3) = 3
                                card(3) = 3
                            Case 4
                                Me.picCompCard4.Image = card4
                                compCard(3) = 4
                                card(3) = 4
                            Case 5
                                Me.picCompCard4.Image = card5
                                compCard(3) = 5
                                card(3) = 5
                            Case 6
                                Me.picCompCard4.Image = card6
                                compCard(3) = 6
                                card(3) = 6
                            Case 7
                                Me.picCompCard4.Image = card7
                                compCard(3) = 7
                                card(3) = 7
                            Case 8
                                Me.picCompCard4.Image = card8
                                compCard(3) = 8
                                card(3) = 8
                            Case 9
                                Me.picCompCard4.Image = card9
                                compCard(3) = 9
                                card(3) = 9
                            Case 10
                                Me.picCompCard4.Image = card10
                                compCard(3) = 10
                                card(3) = 10
                            Case 11
                                Me.picCompCard4.Image = card11
                                compCard(3) = 11
                                card(3) = 11
                            Case 12
                                Me.picCompCard4.Image = card12
                                compCard(3) = 12
                                card(3) = 12
                            Case 13
                                Me.picCompCard4.Image = card13
                                compCard(3) = 13
                                card(3) = 13
                            Case 14
                                Me.picCompCard4.Image = card14
                                compCard(3) = 1
                                card(3) = 14
                            Case 15
                                Me.picCompCard4.Image = card15
                                compCard(3) = 2
                                card(3) = 15
                            Case 16
                                Me.picCompCard4.Image = card16
                                compCard(3) = 3
                                card(3) = 16
                            Case 17
                                Me.picCompCard4.Image = card17
                                compCard(3) = 4
                                card(3) = 17
                            Case 18
                                Me.picCompCard4.Image = card18
                                compCard(3) = 5
                                card(3) = 18
                            Case 19
                                Me.picCompCard4.Image = card19
                                compCard(3) = 6
                                card(3) = 19
                            Case 20
                                Me.picCompCard4.Image = card20
                                compCard(3) = 7
                                card(3) = 20
                            Case 21
                                Me.picCompCard4.Image = card21
                                compCard(3) = 8
                                card(3) = 21
                            Case 22
                                Me.picCompCard4.Image = card22
                                compCard(3) = 9
                                card(3) = 22
                            Case 23
                                Me.picCompCard4.Image = card23
                                compCard(3) = 10
                                card(3) = 23
                            Case 24
                                Me.picCompCard4.Image = card24
                                compCard(3) = 11
                                card(3) = 24
                            Case 25
                                Me.picCompCard4.Image = card25
                                compCard(3) = 12
                                card(3) = 25
                            Case 26
                                Me.picCompCard4.Image = card26
                                compCard(3) = 13
                                card(3) = 26
                            Case 27
                                Me.picCompCard4.Image = card27
                                compCard(3) = 1
                                card(3) = 27
                            Case 28
                                Me.picCompCard4.Image = card28
                                compCard(3) = 2
                                card(3) = 28
                            Case 29
                                Me.picCompCard4.Image = card29
                                compCard(3) = 3
                                card(3) = 29
                            Case 30
                                Me.picCompCard4.Image = card30
                                compCard(3) = 4
                                card(3) = 30
                            Case 31
                                Me.picCompCard4.Image = card31
                                compCard(3) = 5
                                card(3) = 31
                            Case 32
                                Me.picCompCard4.Image = card32
                                compCard(3) = 6
                                card(3) = 32
                            Case 33
                                Me.picCompCard4.Image = card33
                                compCard(3) = 7
                                card(3) = 33
                            Case 34
                                Me.picCompCard4.Image = card34
                                compCard(3) = 8
                                card(3) = 34
                            Case 35
                                Me.picCompCard4.Image = card35
                                compCard(3) = 9
                                card(3) = 35
                            Case 36
                                Me.picCompCard4.Image = card36
                                compCard(3) = 10
                                card(3) = 36
                            Case 37
                                Me.picCompCard4.Image = card37
                                compCard(3) = 11
                                card(3) = 37
                            Case 38
                                Me.picCompCard4.Image = card38
                                compCard(3) = 12
                                card(3) = 38
                            Case 39
                                Me.picCompCard4.Image = card39
                                compCard(3) = 13
                                card(3) = 39
                            Case 40
                                Me.picCompCard4.Image = card40
                                compCard(3) = 1
                                card(3) = 40
                            Case 41
                                Me.picCompCard4.Image = card41
                                compCard(3) = 2
                                card(3) = 41
                            Case 42
                                Me.picCompCard4.Image = card42
                                compCard(3) = 3
                                card(3) = 42
                            Case 43
                                Me.picCompCard4.Image = card43
                                compCard(3) = 4
                                card(3) = 43
                            Case 44
                                Me.picCompCard4.Image = card44
                                compCard(3) = 5
                                card(3) = 44
                            Case 45
                                Me.picCompCard4.Image = card45
                                compCard(3) = 6
                                card(3) = 45
                            Case 46
                                Me.picCompCard4.Image = card46
                                compCard(3) = 7
                                card(3) = 46
                            Case 47
                                Me.picCompCard4.Image = card47
                                compCard(3) = 8
                                card(3) = 47
                            Case 48
                                Me.picCompCard4.Image = card48
                                compCard(3) = 9
                                card(3) = 48
                            Case 49
                                Me.picCompCard4.Image = card49
                                compCard(3) = 10
                                card(3) = 49
                            Case 50
                                Me.picCompCard4.Image = card50
                                compCard(3) = 11
                                card(3) = 50
                            Case 51
                                Me.picCompCard4.Image = card51
                                compCard(3) = 12
                                card(3) = 51
                            Case 52
                                Me.picCompCard4.Image = card52
                                compCard(3) = 13
                                card(3) = 52
                        End Select
                    ElseIf card(11) = card(4) Then
                        Select Case newNum
                            Case 1
                                Me.picCompCard5.Image = card1
                                compCard(4) = 1
                                card(4) = 1
                            Case 2
                                Me.picCompCard5.Image = card2
                                compCard(4) = 2
                                card(4) = 2
                            Case 3
                                Me.picCompCard5.Image = card3
                                compCard(4) = 3
                                card(4) = 3
                            Case 4
                                Me.picCompCard5.Image = card4
                                compCard(4) = 4
                                card(4) = 4
                            Case 5
                                Me.picCompCard5.Image = card5
                                compCard(4) = 5
                                card(4) = 5
                            Case 6
                                Me.picCompCard5.Image = card6
                                compCard(4) = 6
                                card(4) = 6
                            Case 7
                                Me.picCompCard5.Image = card7
                                compCard(4) = 7
                                card(4) = 7
                            Case 8
                                Me.picCompCard5.Image = card8
                                compCard(4) = 8
                                card(4) = 8
                            Case 9
                                Me.picCompCard5.Image = card9
                                compCard(4) = 9
                                card(4) = 9
                            Case 10
                                Me.picCompCard5.Image = card10
                                compCard(4) = 10
                                card(4) = 10
                            Case 11
                                Me.picCompCard5.Image = card11
                                compCard(4) = 11
                                card(4) = 11
                            Case 12
                                Me.picCompCard5.Image = card12
                                compCard(4) = 12
                                card(4) = 12
                            Case 13
                                Me.picCompCard5.Image = card13
                                compCard(4) = 13
                                card(4) = 13
                            Case 14
                                Me.picCompCard5.Image = card14
                                compCard(4) = 1
                                card(4) = 14
                            Case 15
                                Me.picCompCard5.Image = card15
                                compCard(4) = 2
                                card(4) = 15
                            Case 16
                                Me.picCompCard5.Image = card16
                                compCard(4) = 3
                                card(4) = 16
                            Case 17
                                Me.picCompCard5.Image = card17
                                compCard(4) = 4
                                card(4) = 17
                            Case 18
                                Me.picCompCard5.Image = card18
                                compCard(4) = 5
                                card(4) = 18
                            Case 19
                                Me.picCompCard5.Image = card19
                                compCard(4) = 6
                                card(4) = 19
                            Case 20
                                Me.picCompCard5.Image = card20
                                compCard(4) = 7
                                card(4) = 20
                            Case 21
                                Me.picCompCard5.Image = card21
                                compCard(4) = 8
                                card(4) = 21
                            Case 22
                                Me.picCompCard5.Image = card22
                                compCard(4) = 9
                                card(4) = 22
                            Case 23
                                Me.picCompCard5.Image = card23
                                compCard(4) = 10
                                card(4) = 23
                            Case 24
                                Me.picCompCard5.Image = card24
                                compCard(4) = 11
                                card(4) = 24
                            Case 25
                                Me.picCompCard5.Image = card25
                                compCard(4) = 12
                                card(4) = 25
                            Case 26
                                Me.picCompCard5.Image = card26
                                compCard(4) = 13
                                card(4) = 26
                            Case 27
                                Me.picCompCard5.Image = card27
                                compCard(4) = 1
                                card(4) = 27
                            Case 28
                                Me.picCompCard5.Image = card28
                                compCard(4) = 2
                                card(4) = 28
                            Case 29
                                Me.picCompCard5.Image = card29
                                compCard(4) = 3
                                card(4) = 29
                            Case 30
                                Me.picCompCard5.Image = card30
                                compCard(4) = 4
                                card(4) = 30
                            Case 31
                                Me.picCompCard5.Image = card31
                                compCard(4) = 5
                                card(4) = 31
                            Case 32
                                Me.picCompCard5.Image = card32
                                compCard(4) = 6
                                card(4) = 32
                            Case 33
                                Me.picCompCard5.Image = card33
                                compCard(4) = 7
                                card(4) = 33
                            Case 34
                                Me.picCompCard5.Image = card34
                                compCard(4) = 8
                                card(4) = 34
                            Case 35
                                Me.picCompCard5.Image = card35
                                compCard(4) = 9
                                card(4) = 35
                            Case 36
                                Me.picCompCard5.Image = card36
                                compCard(4) = 10
                                card(4) = 36
                            Case 37
                                Me.picCompCard5.Image = card37
                                compCard(4) = 11
                                card(4) = 37
                            Case 38
                                Me.picCompCard5.Image = card38
                                compCard(4) = 12
                                card(4) = 38
                            Case 39
                                Me.picCompCard5.Image = card39
                                compCard(4) = 13
                                card(4) = 39
                            Case 40
                                Me.picCompCard5.Image = card40
                                compCard(4) = 1
                                card(4) = 40
                            Case 41
                                Me.picCompCard5.Image = card41
                                compCard(4) = 2
                                card(4) = 41
                            Case 42
                                Me.picCompCard5.Image = card42
                                compCard(4) = 3
                                card(4) = 42
                            Case 43
                                Me.picCompCard5.Image = card43
                                compCard(4) = 4
                                card(4) = 43
                            Case 44
                                Me.picCompCard5.Image = card44
                                compCard(4) = 5
                                card(4) = 44
                            Case 45
                                Me.picCompCard5.Image = card45
                                compCard(4) = 6
                                card(4) = 45
                            Case 46
                                Me.picCompCard5.Image = card46
                                compCard(4) = 7
                                card(4) = 46
                            Case 47
                                Me.picCompCard5.Image = card47
                                compCard(4) = 8
                                card(4) = 47
                            Case 48
                                Me.picCompCard5.Image = card48
                                compCard(4) = 9
                                card(4) = 48
                            Case 49
                                Me.picCompCard5.Image = card49
                                compCard(4) = 10
                                card(4) = 49
                            Case 50
                                Me.picCompCard5.Image = card50
                                compCard(4) = 11
                                card(4) = 50
                            Case 51
                                Me.picCompCard5.Image = card51
                                compCard(4) = 12
                                card(4) = 51
                            Case 52
                                Me.picCompCard5.Image = card52
                                compCard(4) = 13
                                card(4) = 52
                        End Select
                    End If

                    Used.Add(newNum)

                End If
            Loop While OK = False
        End If

        compCardsLeftTotal = compCardsLeft1 + compCardsLeft2 + compCardsLeft3 + compCardsLeft4 + compCardsLeft5
        Me.lblCompCardsLeftTotal.Text = "Computer's" & vbCrLf & "Total Cards Left:" & vbCrLf & compCardsLeftTotal
        Me.lblCompCardsLeft1.Text = "Cards Left: " & compCardsLeft1
        Me.lblCompCardsLeft2.Text = "Cards Left: " & compCardsLeft2
        Me.lblCompCardsLeft3.Text = "Cards Left: " & compCardsLeft3
        Me.lblCompCardsLeft4.Text = "Cards Left: " & compCardsLeft4
        Me.lblCompCardsLeft5.Text = "Cards Left: " & compCardsLeft5

    End Sub

    Private Sub picPlayerCard1_MouseDown(sender As Object, e As MouseEventArgs) Handles picPlayerCard1.MouseDown

        'allow card to be dragged if it is 1 higher or 1 lower than either of the piles
        For num As Integer = 1 To 13
            If compPile = num And (playerCard(0) = num - 1 Or playerCard(0) = num + 1) Then
                cardCheck = 1
                If e.Button = MouseButtons.Left Then picPlayerCard1.DoDragDrop(picPlayerCard1.Image, DragDropEffects.Copy)
            ElseIf compPile = 1 And (playerCard(0) = 2 Or playerCard(0) = 13) Then
                cardCheck = 1
                If e.Button = MouseButtons.Left Then picPlayerCard1.DoDragDrop(picPlayerCard1.Image, DragDropEffects.Copy)
            ElseIf compPile = 13 And (playerCard(0) = 12 Or playerCard(0) = 1) Then
                cardCheck = 1
                If e.Button = MouseButtons.Left Then picPlayerCard1.DoDragDrop(picPlayerCard1.Image, DragDropEffects.Copy)
            ElseIf playerPile = num And (playerCard(0) = num - 1 Or playerCard(0) = num + 1) Then
                cardCheck = 1
                If e.Button = MouseButtons.Left Then picPlayerCard1.DoDragDrop(picPlayerCard1.Image, DragDropEffects.Copy)
            ElseIf playerPile = 1 And (playerCard(0) = 2 Or playerCard(0) = 13) Then
                cardCheck = 1
                If e.Button = MouseButtons.Left Then picPlayerCard1.DoDragDrop(picPlayerCard1.Image, DragDropEffects.Copy)
            ElseIf playerPile = 13 And (playerCard(0) = 12 Or playerCard(0) = 1) Then
                cardCheck = 1
                If e.Button = MouseButtons.Left Then picPlayerCard1.DoDragDrop(picPlayerCard1.Image, DragDropEffects.Copy)
            End If
        Next num

    End Sub

    Private Sub picPlayerCard2_MouseDown(sender As Object, e As MouseEventArgs) Handles picPlayerCard2.MouseDown

        For num As Integer = 1 To 13
            If compPile = num And (playerCard(1) = num - 1 Or playerCard(1) = num + 1) Then
                cardCheck = 2
                If e.Button = MouseButtons.Left Then picPlayerCard2.DoDragDrop(picPlayerCard2.Image, DragDropEffects.Copy)
            ElseIf compPile = 1 And (playerCard(1) = 2 Or playerCard(1) = 13) Then
                cardCheck = 2
                If e.Button = MouseButtons.Left Then picPlayerCard2.DoDragDrop(picPlayerCard2.Image, DragDropEffects.Copy)
            ElseIf compPile = 13 And (playerCard(1) = 12 Or playerCard(1) = 1) Then
                cardCheck = 2
                If e.Button = MouseButtons.Left Then picPlayerCard2.DoDragDrop(picPlayerCard2.Image, DragDropEffects.Copy)
            ElseIf playerPile = num And (playerCard(1) = num - 1 Or playerCard(1) = num + 1) Then
                cardCheck = 2
                If e.Button = MouseButtons.Left Then picPlayerCard2.DoDragDrop(picPlayerCard2.Image, DragDropEffects.Copy)
            ElseIf playerPile = 1 And (playerCard(1) = 2 Or playerCard(1) = 13) Then
                cardCheck = 2
                If e.Button = MouseButtons.Left Then picPlayerCard2.DoDragDrop(picPlayerCard2.Image, DragDropEffects.Copy)
            ElseIf playerPile = 13 And (playerCard(1) = 12 Or playerCard(1) = 1) Then
                cardCheck = 2
                If e.Button = MouseButtons.Left Then picPlayerCard2.DoDragDrop(picPlayerCard2.Image, DragDropEffects.Copy)
            End If
        Next num

    End Sub

    Private Sub picPlayerCard3_MouseDown(sender As Object, e As MouseEventArgs) Handles picPlayerCard3.MouseDown

        For num As Integer = 1 To 13
            If compPile = num And (playerCard(2) = num - 1 Or playerCard(2) = num + 1) Then
                cardCheck = 3
                If e.Button = MouseButtons.Left Then picPlayerCard3.DoDragDrop(picPlayerCard3.Image, DragDropEffects.Copy)
            ElseIf compPile = 1 And (playerCard(2) = 2 Or playerCard(2) = 13) Then
                cardCheck = 3
                If e.Button = MouseButtons.Left Then picPlayerCard3.DoDragDrop(picPlayerCard3.Image, DragDropEffects.Copy)
            ElseIf compPile = 13 And (playerCard(2) = 12 Or playerCard(2) = 1) Then
                cardCheck = 3
                If e.Button = MouseButtons.Left Then picPlayerCard3.DoDragDrop(picPlayerCard3.Image, DragDropEffects.Copy)
            ElseIf playerPile = num And (playerCard(2) = num - 1 Or playerCard(2) = num + 1) Then
                cardCheck = 3
                If e.Button = MouseButtons.Left Then picPlayerCard3.DoDragDrop(picPlayerCard3.Image, DragDropEffects.Copy)
            ElseIf playerPile = 1 And (playerCard(2) = 2 Or playerCard(2) = 13) Then
                cardCheck = 3
                If e.Button = MouseButtons.Left Then picPlayerCard3.DoDragDrop(picPlayerCard3.Image, DragDropEffects.Copy)
            ElseIf playerPile = 13 And (playerCard(2) = 12 Or playerCard(2) = 1) Then
                cardCheck = 3
                If e.Button = MouseButtons.Left Then picPlayerCard3.DoDragDrop(picPlayerCard3.Image, DragDropEffects.Copy)
            End If
        Next num

    End Sub

    Private Sub picPlayerCard4_MouseDown(sender As Object, e As MouseEventArgs) Handles picPlayerCard4.MouseDown

        For num As Integer = 1 To 13
            If compPile = num And (playerCard(3) = num - 1 Or playerCard(3) = num + 1) Then
                cardCheck = 4
                If e.Button = MouseButtons.Left Then picPlayerCard4.DoDragDrop(picPlayerCard4.Image, DragDropEffects.Copy)
            ElseIf compPile = 1 And (playerCard(3) = 2 Or playerCard(3) = 13) Then
                cardCheck = 4
                If e.Button = MouseButtons.Left Then picPlayerCard4.DoDragDrop(picPlayerCard4.Image, DragDropEffects.Copy)
            ElseIf compPile = 13 And (playerCard(3) = 12 Or playerCard(3) = 1) Then
                cardCheck = 4
                If e.Button = MouseButtons.Left Then picPlayerCard4.DoDragDrop(picPlayerCard4.Image, DragDropEffects.Copy)
            ElseIf playerPile = num And (playerCard(3) = num - 1 Or playerCard(3) = num + 1) Then
                cardCheck = 4
                If e.Button = MouseButtons.Left Then picPlayerCard4.DoDragDrop(picPlayerCard4.Image, DragDropEffects.Copy)
            ElseIf playerPile = 1 And (playerCard(3) = 2 Or playerCard(3) = 13) Then
                cardCheck = 4
                If e.Button = MouseButtons.Left Then picPlayerCard4.DoDragDrop(picPlayerCard4.Image, DragDropEffects.Copy)
            ElseIf playerPile = 13 And (playerCard(3) = 12 Or playerCard(3) = 1) Then
                cardCheck = 4
                If e.Button = MouseButtons.Left Then picPlayerCard4.DoDragDrop(picPlayerCard4.Image, DragDropEffects.Copy)
            End If
        Next num

    End Sub

    Private Sub picPlayerCard5_MouseDown(sender As Object, e As MouseEventArgs) Handles picPlayerCard5.MouseDown

        For num As Integer = 1 To 13
            If compPile = num And (playerCard(4) = num - 1 Or playerCard(4) = num + 1) Then
                cardCheck = 5
                If e.Button = MouseButtons.Left Then picPlayerCard5.DoDragDrop(picPlayerCard5.Image, DragDropEffects.Copy)
            ElseIf compPile = 1 And (playerCard(4) = 2 Or playerCard(4) = 13) Then
                cardCheck = 5
                If e.Button = MouseButtons.Left Then picPlayerCard5.DoDragDrop(picPlayerCard5.Image, DragDropEffects.Copy)
            ElseIf compPile = 13 And (playerCard(4) = 12 Or playerCard(4) = 1) Then
                cardCheck = 5
                If e.Button = MouseButtons.Left Then picPlayerCard5.DoDragDrop(picPlayerCard5.Image, DragDropEffects.Copy)
            ElseIf playerPile = num And (playerCard(4) = num - 1 Or playerCard(4) = num + 1) Then
                cardCheck = 5
                If e.Button = MouseButtons.Left Then picPlayerCard5.DoDragDrop(picPlayerCard5.Image, DragDropEffects.Copy)
            ElseIf playerPile = 1 And (playerCard(4) = 2 Or playerCard(4) = 13) Then
                cardCheck = 5
                If e.Button = MouseButtons.Left Then picPlayerCard5.DoDragDrop(picPlayerCard5.Image, DragDropEffects.Copy)
            ElseIf playerPile = 13 And (playerCard(4) = 12 Or playerCard(4) = 1) Then
                cardCheck = 5
                If e.Button = MouseButtons.Left Then picPlayerCard5.DoDragDrop(picPlayerCard5.Image, DragDropEffects.Copy)
            End If
        Next num

    End Sub

    Private Sub picCompPile_DragEnter(sender As Object, e As DragEventArgs) Handles picCompPile.DragEnter
        If e.Data.GetDataPresent(DataFormats.Bitmap) Then
            e.Effect = DragDropEffects.Copy
        Else
            e.Effect = DragDropEffects.None
        End If
    End Sub

    Private Sub picPlayerPile_DragEnter(sender As Object, e As DragEventArgs) Handles picPlayerPile.DragEnter
        If e.Data.GetDataPresent(DataFormats.Bitmap) Then
            e.Effect = DragDropEffects.Copy
        Else
            e.Effect = DragDropEffects.None
        End If
    End Sub

    Private Sub picCompPile_DragDrop(sender As Object, e As DragEventArgs) Handles picCompPile.DragDrop
        Dim imageMove As Bitmap = DirectCast(e.Data.GetData(DataFormats.Bitmap, True), Bitmap)
        Dim newNum As Integer
        Dim OK As Boolean = False

        If cardCheck = 1 Then                                                                       'if playerCard0 was dragged, 
            For num As Integer = 1 To 13                                                            'check if it is 1 higher or 1 lower than compPile
                If compPile = num And (playerCard(0) = num + 1 Or playerCard(0) = num - 1) Then
                    picCompPile.Image = imageMove                   'change compPile card image
                    card(10) = card(5)                              'change card [A to K]
                    compPile = playerCard(0)                        'change card [1 to 52]
                    If playerCardsLeft1 > 0 Then
                        playerCardsLeft1 -= 1                       'subtract 1 from number of player cards left
                    End If
                ElseIf compPile = 1 And (playerCard(0) = 2 Or playerCard(0) = 13) Then
                    picCompPile.Image = imageMove
                    card(10) = card(5)
                    compPile = playerCard(0)
                    If playerCardsLeft1 > 0 Then
                        playerCardsLeft1 -= 1
                    End If
                ElseIf compPile = 13 And (playerCard(0) = 1 Or playerCard(0) = 12) Then
                    picCompPile.Image = imageMove
                    card(10) = card(5)
                    compPile = playerCard(0)
                    If playerCardsLeft1 > 0 Then
                        playerCardsLeft1 -= 1
                    End If
                End If
            Next
        ElseIf cardCheck = 2 Then
            For num As Integer = 1 To 13
                If compPile = num And (playerCard(1) = num + 1 Or playerCard(1) = num - 1) Then
                    picCompPile.Image = imageMove
                    card(10) = card(6)
                    compPile = playerCard(1)
                    If playerCardsLeft2 > 0 Then
                        playerCardsLeft2 -= 1
                    End If
                ElseIf compPile = 1 And (playerCard(1) = 2 Or playerCard(1) = 13) Then
                    picCompPile.Image = imageMove
                    card(10) = card(6)
                    compPile = playerCard(1)
                    If playerCardsLeft2 > 0 Then
                        playerCardsLeft2 -= 1
                    End If
                ElseIf compPile = 13 And (playerCard(1) = 1 Or playerCard(1) = 12) Then
                    picCompPile.Image = imageMove
                    card(10) = card(6)
                    compPile = playerCard(1)
                    If playerCardsLeft2 > 0 Then
                        playerCardsLeft2 -= 1
                    End If
                End If
            Next
        ElseIf cardCheck = 3 Then
            For num As Integer = 1 To 13
                If compPile = num And (playerCard(2) = num + 1 Or playerCard(2) = num - 1) Then
                    picCompPile.Image = imageMove
                    card(10) = card(7)
                    compPile = playerCard(2)
                    If playerCardsLeft3 > 0 Then
                        playerCardsLeft3 -= 1
                    End If
                ElseIf compPile = 1 And (playerCard(2) = 2 Or playerCard(2) = 13) Then
                    picCompPile.Image = imageMove
                    card(10) = card(7)
                    compPile = playerCard(2)
                    If playerCardsLeft3 > 0 Then
                        playerCardsLeft3 -= 1
                    End If
                ElseIf compPile = 13 And (playerCard(2) = 1 Or playerCard(2) = 12) Then
                    picCompPile.Image = imageMove
                    card(10) = card(7)
                    compPile = playerCard(2)
                    If playerCardsLeft3 > 0 Then
                        playerCardsLeft3 -= 1
                    End If
                End If
            Next
        ElseIf cardCheck = 4 Then
            For num As Integer = 1 To 13
                If compPile = num And (playerCard(3) = num + 1 Or playerCard(3) = num - 1) Then
                    picCompPile.Image = imageMove
                    card(10) = card(8)
                    compPile = playerCard(3)
                    If playerCardsLeft4 > 0 Then
                        playerCardsLeft4 -= 1
                    End If
                ElseIf compPile = 1 And (playerCard(3) = 2 Or playerCard(3) = 13) Then
                    picCompPile.Image = imageMove
                    card(10) = card(8)
                    compPile = playerCard(3)
                    If playerCardsLeft4 > 0 Then
                        playerCardsLeft4 -= 1
                    End If
                ElseIf compPile = 13 And (playerCard(3) = 1 Or playerCard(3) = 12) Then
                    picCompPile.Image = imageMove
                    card(10) = card(8)
                    compPile = playerCard(3)
                    If playerCardsLeft4 > 0 Then
                        playerCardsLeft4 -= 1
                    End If
                End If
            Next
        ElseIf cardCheck = 5 Then
            For num As Integer = 1 To 13
                If compPile = num And (playerCard(4) = num + 1 Or playerCard(4) = num - 1) Then
                    picCompPile.Image = imageMove
                    card(10) = card(9)
                    compPile = playerCard(4)
                    If playerCardsLeft5 > 0 Then
                        playerCardsLeft5 -= 1
                    End If
                ElseIf compPile = 1 And (playerCard(4) = 2 Or playerCard(4) = 13) Then
                    picCompPile.Image = imageMove
                    card(10) = card(9)
                    compPile = playerCard(4)
                    If playerCardsLeft5 > 0 Then
                        playerCardsLeft5 -= 1
                    End If
                ElseIf compPile = 13 And (playerCard(4) = 1 Or playerCard(4) = 12) Then
                    picCompPile.Image = imageMove
                    card(10) = card(9)
                    compPile = playerCard(4)
                    If playerCardsLeft5 > 0 Then
                        playerCardsLeft5 -= 1
                    End If
                End If
            Next
        End If

        If card(10) = card(5) Or card(10) = card(6) Or card(10) = card(7) Or card(10) = card(8) Or card(10) = card(9) Then
            'allow generate new card if compPile is same as any of the playerCards
            If playerCardsLeft1 > 0 Or playerCardsLeft2 > 0 Or playerCardsLeft3 > 0 Or playerCardsLeft4 > 0 Or playerCardsLeft5 > 0 Then
                Do                                                      'generate new card only if playerCardsLeft is more than 0
                    Randomize()
                    newNum = Int(max * Rnd() + min)

                    If Used.Contains(newNum) Then
                        OK = False
                    Else
                        OK = True

                        If card(10) = card(5) Then
                            Select Case newNum
                                Case 1
                                    Me.picPlayerCard1.Image = card1
                                    playerCard(0) = 1
                                    card(5) = 1
                                Case 2
                                    Me.picPlayerCard1.Image = card2
                                    playerCard(0) = 2
                                    card(5) = 2
                                Case 3
                                    Me.picPlayerCard1.Image = card3
                                    playerCard(0) = 3
                                    card(5) = 3
                                Case 4
                                    Me.picPlayerCard1.Image = card4
                                    playerCard(0) = 4
                                    card(5) = 4
                                Case 5
                                    Me.picPlayerCard1.Image = card5
                                    playerCard(0) = 5
                                    card(5) = 5
                                Case 6
                                    Me.picPlayerCard1.Image = card6
                                    playerCard(0) = 6
                                    card(5) = 6
                                Case 7
                                    Me.picPlayerCard1.Image = card7
                                    playerCard(0) = 7
                                    card(5) = 7
                                Case 8
                                    Me.picPlayerCard1.Image = card8
                                    playerCard(0) = 8
                                    card(5) = 8
                                Case 9
                                    Me.picPlayerCard1.Image = card9
                                    playerCard(0) = 9
                                    card(5) = 9
                                Case 10
                                    Me.picPlayerCard1.Image = card10
                                    playerCard(0) = 10
                                    card(5) = 10
                                Case 11
                                    Me.picPlayerCard1.Image = card11
                                    playerCard(0) = 11
                                    card(5) = 11
                                Case 12
                                    Me.picPlayerCard1.Image = card12
                                    playerCard(0) = 12
                                    card(5) = 12
                                Case 13
                                    Me.picPlayerCard1.Image = card13
                                    playerCard(0) = 13
                                    card(5) = 13
                                Case 14
                                    Me.picPlayerCard1.Image = card14
                                    playerCard(0) = 1
                                    card(5) = 14
                                Case 15
                                    Me.picPlayerCard1.Image = card15
                                    playerCard(0) = 2
                                    card(5) = 15
                                Case 16
                                    Me.picPlayerCard1.Image = card16
                                    playerCard(0) = 3
                                    card(5) = 16
                                Case 17
                                    Me.picPlayerCard1.Image = card17
                                    playerCard(0) = 4
                                    card(5) = 17
                                Case 18
                                    Me.picPlayerCard1.Image = card18
                                    playerCard(0) = 5
                                    card(5) = 18
                                Case 19
                                    Me.picPlayerCard1.Image = card19
                                    playerCard(0) = 6
                                    card(5) = 19
                                Case 20
                                    Me.picPlayerCard1.Image = card20
                                    playerCard(0) = 7
                                    card(5) = 20
                                Case 21
                                    Me.picPlayerCard1.Image = card21
                                    playerCard(0) = 8
                                    card(5) = 21
                                Case 22
                                    Me.picPlayerCard1.Image = card22
                                    playerCard(0) = 9
                                    card(5) = 22
                                Case 23
                                    Me.picPlayerCard1.Image = card23
                                    playerCard(0) = 10
                                    card(5) = 23
                                Case 24
                                    Me.picPlayerCard1.Image = card24
                                    playerCard(0) = 11
                                    card(5) = 24
                                Case 25
                                    Me.picPlayerCard1.Image = card25
                                    playerCard(0) = 12
                                    card(5) = 25
                                Case 26
                                    Me.picPlayerCard1.Image = card26
                                    playerCard(0) = 13
                                    card(5) = 26
                                Case 27
                                    Me.picPlayerCard1.Image = card27
                                    playerCard(0) = 1
                                    card(5) = 27
                                Case 28
                                    Me.picPlayerCard1.Image = card28
                                    playerCard(0) = 2
                                    card(5) = 28
                                Case 29
                                    Me.picPlayerCard1.Image = card29
                                    playerCard(0) = 3
                                    card(5) = 29
                                Case 30
                                    Me.picPlayerCard1.Image = card30
                                    playerCard(0) = 4
                                    card(5) = 30
                                Case 31
                                    Me.picPlayerCard1.Image = card31
                                    playerCard(0) = 5
                                    card(5) = 31
                                Case 32
                                    Me.picPlayerCard1.Image = card32
                                    playerCard(0) = 6
                                    card(5) = 32
                                Case 33
                                    Me.picPlayerCard1.Image = card33
                                    playerCard(0) = 7
                                    card(5) = 33
                                Case 34
                                    Me.picPlayerCard1.Image = card34
                                    playerCard(0) = 8
                                    card(5) = 34
                                Case 35
                                    Me.picPlayerCard1.Image = card35
                                    playerCard(0) = 9
                                    card(5) = 35
                                Case 36
                                    Me.picPlayerCard1.Image = card36
                                    playerCard(0) = 10
                                    card(5) = 36
                                Case 37
                                    Me.picPlayerCard1.Image = card37
                                    playerCard(0) = 11
                                    card(5) = 37
                                Case 38
                                    Me.picPlayerCard1.Image = card38
                                    playerCard(0) = 12
                                    card(5) = 38
                                Case 39
                                    Me.picPlayerCard1.Image = card39
                                    playerCard(0) = 13
                                    card(5) = 39
                                Case 40
                                    Me.picPlayerCard1.Image = card40
                                    playerCard(0) = 1
                                    card(5) = 40
                                Case 41
                                    Me.picPlayerCard1.Image = card41
                                    playerCard(0) = 2
                                    card(5) = 41
                                Case 42
                                    Me.picPlayerCard1.Image = card42
                                    playerCard(0) = 3
                                    card(5) = 42
                                Case 43
                                    Me.picPlayerCard1.Image = card43
                                    playerCard(0) = 4
                                    card(5) = 43
                                Case 44
                                    Me.picPlayerCard1.Image = card44
                                    playerCard(0) = 5
                                    card(5) = 44
                                Case 45
                                    Me.picPlayerCard1.Image = card45
                                    playerCard(0) = 6
                                    card(5) = 45
                                Case 46
                                    Me.picPlayerCard1.Image = card46
                                    playerCard(0) = 7
                                    card(5) = 46
                                Case 47
                                    Me.picPlayerCard1.Image = card47
                                    playerCard(0) = 8
                                    card(5) = 47
                                Case 48
                                    Me.picPlayerCard1.Image = card48
                                    playerCard(0) = 9
                                    card(5) = 48
                                Case 49
                                    Me.picPlayerCard1.Image = card49
                                    playerCard(0) = 10
                                    card(5) = 49
                                Case 50
                                    Me.picPlayerCard1.Image = card50
                                    playerCard(0) = 11
                                    card(5) = 50
                                Case 51
                                    Me.picPlayerCard1.Image = card51
                                    playerCard(0) = 12
                                    card(5) = 51
                                Case 52
                                    Me.picPlayerCard1.Image = card52
                                    playerCard(0) = 13
                                    card(5) = 52
                            End Select
                        ElseIf card(10) = card(6) Then
                            Select Case newNum
                                Case 1
                                    Me.picPlayerCard2.Image = card1
                                    playerCard(1) = 1
                                    card(6) = 1
                                Case 2
                                    Me.picPlayerCard2.Image = card2
                                    playerCard(1) = 2
                                    card(6) = 2
                                Case 3
                                    Me.picPlayerCard2.Image = card3
                                    playerCard(1) = 3
                                    card(6) = 3
                                Case 4
                                    Me.picPlayerCard2.Image = card4
                                    playerCard(1) = 4
                                    card(6) = 4
                                Case 5
                                    Me.picPlayerCard2.Image = card5
                                    playerCard(1) = 5
                                    card(6) = 5
                                Case 6
                                    Me.picPlayerCard2.Image = card6
                                    playerCard(1) = 6
                                    card(6) = 6
                                Case 7
                                    Me.picPlayerCard2.Image = card7
                                    playerCard(1) = 7
                                    card(6) = 7
                                Case 8
                                    Me.picPlayerCard2.Image = card8
                                    playerCard(1) = 8
                                    card(6) = 8
                                Case 9
                                    Me.picPlayerCard2.Image = card9
                                    playerCard(1) = 9
                                    card(6) = 9
                                Case 10
                                    Me.picPlayerCard2.Image = card10
                                    playerCard(1) = 10
                                    card(6) = 10
                                Case 11
                                    Me.picPlayerCard2.Image = card11
                                    playerCard(1) = 11
                                    card(6) = 11
                                Case 12
                                    Me.picPlayerCard2.Image = card12
                                    playerCard(1) = 12
                                    card(6) = 12
                                Case 13
                                    Me.picPlayerCard2.Image = card13
                                    playerCard(1) = 13
                                    card(6) = 13
                                Case 14
                                    Me.picPlayerCard2.Image = card14
                                    playerCard(1) = 1
                                    card(6) = 14
                                Case 15
                                    Me.picPlayerCard2.Image = card15
                                    playerCard(1) = 2
                                    card(6) = 15
                                Case 16
                                    Me.picPlayerCard2.Image = card16
                                    playerCard(1) = 3
                                    card(6) = 16
                                Case 17
                                    Me.picPlayerCard2.Image = card17
                                    playerCard(1) = 4
                                    card(6) = 17
                                Case 18
                                    Me.picPlayerCard2.Image = card18
                                    playerCard(1) = 5
                                    card(6) = 18
                                Case 19
                                    Me.picPlayerCard2.Image = card19
                                    playerCard(1) = 6
                                    card(6) = 19
                                Case 20
                                    Me.picPlayerCard2.Image = card20
                                    playerCard(1) = 7
                                    card(6) = 20
                                Case 21
                                    Me.picPlayerCard2.Image = card21
                                    playerCard(1) = 8
                                    card(6) = 21
                                Case 22
                                    Me.picPlayerCard2.Image = card22
                                    playerCard(1) = 9
                                    card(6) = 22
                                Case 23
                                    Me.picPlayerCard2.Image = card23
                                    playerCard(1) = 10
                                    card(6) = 23
                                Case 24
                                    Me.picPlayerCard2.Image = card24
                                    playerCard(1) = 11
                                    card(6) = 24
                                Case 25
                                    Me.picPlayerCard2.Image = card25
                                    playerCard(1) = 12
                                    card(6) = 25
                                Case 26
                                    Me.picPlayerCard2.Image = card26
                                    playerCard(1) = 13
                                    card(6) = 26
                                Case 27
                                    Me.picPlayerCard2.Image = card27
                                    playerCard(1) = 1
                                    card(6) = 27
                                Case 28
                                    Me.picPlayerCard2.Image = card28
                                    playerCard(1) = 2
                                    card(6) = 28
                                Case 29
                                    Me.picPlayerCard2.Image = card29
                                    playerCard(1) = 3
                                    card(6) = 29
                                Case 30
                                    Me.picPlayerCard2.Image = card30
                                    playerCard(1) = 4
                                    card(6) = 30
                                Case 31
                                    Me.picPlayerCard2.Image = card31
                                    playerCard(1) = 5
                                    card(6) = 31
                                Case 32
                                    Me.picPlayerCard2.Image = card32
                                    playerCard(1) = 6
                                    card(6) = 32
                                Case 33
                                    Me.picPlayerCard2.Image = card33
                                    playerCard(1) = 7
                                    card(6) = 33
                                Case 34
                                    Me.picPlayerCard2.Image = card34
                                    playerCard(1) = 8
                                    card(6) = 34
                                Case 35
                                    Me.picPlayerCard2.Image = card35
                                    playerCard(1) = 9
                                    card(6) = 35
                                Case 36
                                    Me.picPlayerCard2.Image = card36
                                    playerCard(1) = 10
                                    card(6) = 36
                                Case 37
                                    Me.picPlayerCard2.Image = card37
                                    playerCard(1) = 11
                                    card(6) = 37
                                Case 38
                                    Me.picPlayerCard2.Image = card38
                                    playerCard(1) = 12
                                    card(6) = 38
                                Case 39
                                    Me.picPlayerCard2.Image = card39
                                    playerCard(1) = 13
                                    card(6) = 39
                                Case 40
                                    Me.picPlayerCard2.Image = card40
                                    playerCard(1) = 1
                                    card(6) = 40
                                Case 41
                                    Me.picPlayerCard2.Image = card41
                                    playerCard(1) = 2
                                    card(6) = 41
                                Case 42
                                    Me.picPlayerCard2.Image = card42
                                    playerCard(1) = 3
                                    card(6) = 42
                                Case 43
                                    Me.picPlayerCard2.Image = card43
                                    playerCard(1) = 4
                                    card(6) = 43
                                Case 44
                                    Me.picPlayerCard2.Image = card44
                                    playerCard(1) = 5
                                    card(6) = 44
                                Case 45
                                    Me.picPlayerCard2.Image = card45
                                    playerCard(1) = 6
                                    card(6) = 45
                                Case 46
                                    Me.picPlayerCard2.Image = card46
                                    playerCard(1) = 7
                                    card(6) = 46
                                Case 47
                                    Me.picPlayerCard2.Image = card47
                                    playerCard(1) = 8
                                    card(6) = 47
                                Case 48
                                    Me.picPlayerCard2.Image = card48
                                    playerCard(1) = 9
                                    card(6) = 48
                                Case 49
                                    Me.picPlayerCard2.Image = card49
                                    playerCard(1) = 10
                                    card(6) = 49
                                Case 50
                                    Me.picPlayerCard2.Image = card50
                                    playerCard(1) = 11
                                    card(6) = 50
                                Case 51
                                    Me.picPlayerCard2.Image = card51
                                    playerCard(1) = 12
                                    card(6) = 51
                                Case 52
                                    Me.picPlayerCard2.Image = card52
                                    playerCard(1) = 13
                                    card(6) = 52
                            End Select
                        ElseIf card(10) = card(7) Then
                            Select Case newNum
                                Case 1
                                    Me.picPlayerCard3.Image = card1
                                    playerCard(2) = 1
                                    card(7) = 1
                                Case 2
                                    Me.picPlayerCard3.Image = card2
                                    playerCard(2) = 2
                                    card(7) = 2
                                Case 3
                                    Me.picPlayerCard3.Image = card3
                                    playerCard(2) = 3
                                    card(7) = 3
                                Case 4
                                    Me.picPlayerCard3.Image = card4
                                    playerCard(2) = 4
                                    card(7) = 4
                                Case 5
                                    Me.picPlayerCard3.Image = card5
                                    playerCard(2) = 5
                                    card(7) = 5
                                Case 6
                                    Me.picPlayerCard3.Image = card6
                                    playerCard(2) = 6
                                    card(7) = 6
                                Case 7
                                    Me.picPlayerCard3.Image = card7
                                    playerCard(2) = 7
                                    card(7) = 7
                                Case 8
                                    Me.picPlayerCard3.Image = card8
                                    playerCard(2) = 8
                                    card(7) = 8
                                Case 9
                                    Me.picPlayerCard3.Image = card9
                                    playerCard(2) = 9
                                    card(7) = 9
                                Case 10
                                    Me.picPlayerCard3.Image = card10
                                    playerCard(2) = 10
                                    card(7) = 10
                                Case 11
                                    Me.picPlayerCard3.Image = card11
                                    playerCard(2) = 11
                                    card(7) = 11
                                Case 12
                                    Me.picPlayerCard3.Image = card12
                                    playerCard(2) = 12
                                    card(7) = 12
                                Case 13
                                    Me.picPlayerCard3.Image = card13
                                    playerCard(2) = 13
                                    card(7) = 13
                                Case 14
                                    Me.picPlayerCard3.Image = card14
                                    playerCard(2) = 1
                                    card(7) = 14
                                Case 15
                                    Me.picPlayerCard3.Image = card15
                                    playerCard(2) = 2
                                    card(7) = 15
                                Case 16
                                    Me.picPlayerCard3.Image = card16
                                    playerCard(2) = 3
                                    card(7) = 16
                                Case 17
                                    Me.picPlayerCard3.Image = card17
                                    playerCard(2) = 4
                                    card(7) = 17
                                Case 18
                                    Me.picPlayerCard3.Image = card18
                                    playerCard(2) = 5
                                    card(7) = 18
                                Case 19
                                    Me.picPlayerCard3.Image = card19
                                    playerCard(2) = 6
                                    card(7) = 19
                                Case 20
                                    Me.picPlayerCard3.Image = card20
                                    playerCard(2) = 7
                                    card(7) = 20
                                Case 21
                                    Me.picPlayerCard3.Image = card21
                                    playerCard(2) = 8
                                    card(7) = 21
                                Case 22
                                    Me.picPlayerCard3.Image = card22
                                    playerCard(2) = 9
                                    card(7) = 22
                                Case 23
                                    Me.picPlayerCard3.Image = card23
                                    playerCard(2) = 10
                                    card(7) = 23
                                Case 24
                                    Me.picPlayerCard3.Image = card24
                                    playerCard(2) = 11
                                    card(7) = 24
                                Case 25
                                    Me.picPlayerCard3.Image = card25
                                    playerCard(2) = 12
                                    card(7) = 25
                                Case 26
                                    Me.picPlayerCard3.Image = card26
                                    playerCard(2) = 13
                                    card(7) = 26
                                Case 27
                                    Me.picPlayerCard3.Image = card27
                                    playerCard(2) = 1
                                    card(7) = 27
                                Case 28
                                    Me.picPlayerCard3.Image = card28
                                    playerCard(2) = 2
                                    card(7) = 28
                                Case 29
                                    Me.picPlayerCard3.Image = card29
                                    playerCard(2) = 3
                                    card(7) = 29
                                Case 30
                                    Me.picPlayerCard3.Image = card30
                                    playerCard(2) = 4
                                    card(7) = 30
                                Case 31
                                    Me.picPlayerCard3.Image = card31
                                    playerCard(2) = 5
                                    card(7) = 31
                                Case 32
                                    Me.picPlayerCard3.Image = card32
                                    playerCard(2) = 6
                                    card(7) = 32
                                Case 33
                                    Me.picPlayerCard3.Image = card33
                                    playerCard(2) = 7
                                    card(7) = 33
                                Case 34
                                    Me.picPlayerCard3.Image = card34
                                    playerCard(2) = 8
                                    card(7) = 34
                                Case 35
                                    Me.picPlayerCard3.Image = card35
                                    playerCard(2) = 9
                                    card(7) = 35
                                Case 36
                                    Me.picPlayerCard3.Image = card36
                                    playerCard(2) = 10
                                    card(7) = 36
                                Case 37
                                    Me.picPlayerCard3.Image = card37
                                    playerCard(2) = 11
                                    card(7) = 37
                                Case 38
                                    Me.picPlayerCard3.Image = card38
                                    playerCard(2) = 12
                                    card(7) = 38
                                Case 39
                                    Me.picPlayerCard3.Image = card39
                                    playerCard(2) = 13
                                    card(7) = 39
                                Case 40
                                    Me.picPlayerCard3.Image = card40
                                    playerCard(2) = 1
                                    card(7) = 40
                                Case 41
                                    Me.picPlayerCard3.Image = card41
                                    playerCard(2) = 2
                                    card(7) = 41
                                Case 42
                                    Me.picPlayerCard3.Image = card42
                                    playerCard(2) = 3
                                    card(7) = 42
                                Case 43
                                    Me.picPlayerCard3.Image = card43
                                    playerCard(2) = 4
                                    card(7) = 43
                                Case 44
                                    Me.picPlayerCard3.Image = card44
                                    playerCard(2) = 5
                                    card(7) = 44
                                Case 45
                                    Me.picPlayerCard3.Image = card45
                                    playerCard(2) = 6
                                    card(7) = 45
                                Case 46
                                    Me.picPlayerCard3.Image = card46
                                    playerCard(2) = 7
                                    card(7) = 46
                                Case 47
                                    Me.picPlayerCard3.Image = card47
                                    playerCard(2) = 8
                                    card(7) = 47
                                Case 48
                                    Me.picPlayerCard3.Image = card48
                                    playerCard(2) = 9
                                    card(7) = 48
                                Case 49
                                    Me.picPlayerCard3.Image = card49
                                    playerCard(2) = 10
                                    card(7) = 49
                                Case 50
                                    Me.picPlayerCard3.Image = card50
                                    playerCard(2) = 11
                                    card(7) = 50
                                Case 51
                                    Me.picPlayerCard3.Image = card51
                                    playerCard(2) = 12
                                    card(7) = 51
                                Case 52
                                    Me.picPlayerCard3.Image = card52
                                    playerCard(2) = 13
                                    card(7) = 52
                            End Select
                        ElseIf card(10) = card(8) Then
                            Select Case newNum
                                Case 1
                                    Me.picPlayerCard4.Image = card1
                                    playerCard(3) = 1
                                    card(8) = 1
                                Case 2
                                    Me.picPlayerCard4.Image = card2
                                    playerCard(3) = 2
                                    card(8) = 2
                                Case 3
                                    Me.picPlayerCard4.Image = card3
                                    playerCard(3) = 3
                                    card(8) = 3
                                Case 4
                                    Me.picPlayerCard4.Image = card4
                                    playerCard(3) = 4
                                    card(8) = 4
                                Case 5
                                    Me.picPlayerCard4.Image = card5
                                    playerCard(3) = 5
                                    card(8) = 5
                                Case 6
                                    Me.picPlayerCard4.Image = card6
                                    playerCard(3) = 6
                                    card(8) = 6
                                Case 7
                                    Me.picPlayerCard4.Image = card7
                                    playerCard(3) = 7
                                    card(8) = 7
                                Case 8
                                    Me.picPlayerCard4.Image = card8
                                    playerCard(3) = 8
                                    card(8) = 8
                                Case 9
                                    Me.picPlayerCard4.Image = card9
                                    playerCard(3) = 9
                                    card(8) = 9
                                Case 10
                                    Me.picPlayerCard4.Image = card10
                                    playerCard(3) = 10
                                    card(8) = 10
                                Case 11
                                    Me.picPlayerCard4.Image = card11
                                    playerCard(3) = 11
                                    card(8) = 11
                                Case 12
                                    Me.picPlayerCard4.Image = card12
                                    playerCard(3) = 12
                                    card(8) = 12
                                Case 13
                                    Me.picPlayerCard4.Image = card13
                                    playerCard(3) = 13
                                    card(8) = 13
                                Case 14
                                    Me.picPlayerCard4.Image = card14
                                    playerCard(3) = 1
                                    card(8) = 14
                                Case 15
                                    Me.picPlayerCard4.Image = card15
                                    playerCard(3) = 2
                                    card(8) = 15
                                Case 16
                                    Me.picPlayerCard4.Image = card16
                                    playerCard(3) = 3
                                    card(8) = 16
                                Case 17
                                    Me.picPlayerCard4.Image = card17
                                    playerCard(3) = 4
                                    card(8) = 17
                                Case 18
                                    Me.picPlayerCard4.Image = card18
                                    playerCard(3) = 5
                                    card(8) = 18
                                Case 19
                                    Me.picPlayerCard4.Image = card19
                                    playerCard(3) = 6
                                    card(8) = 19
                                Case 20
                                    Me.picPlayerCard4.Image = card20
                                    playerCard(3) = 7
                                    card(8) = 20
                                Case 21
                                    Me.picPlayerCard4.Image = card21
                                    playerCard(3) = 8
                                    card(8) = 21
                                Case 22
                                    Me.picPlayerCard4.Image = card22
                                    playerCard(3) = 9
                                    card(8) = 22
                                Case 23
                                    Me.picPlayerCard4.Image = card23
                                    playerCard(3) = 10
                                    card(8) = 23
                                Case 24
                                    Me.picPlayerCard4.Image = card24
                                    playerCard(3) = 11
                                    card(8) = 24
                                Case 25
                                    Me.picPlayerCard4.Image = card25
                                    playerCard(3) = 12
                                    card(8) = 25
                                Case 26
                                    Me.picPlayerCard4.Image = card26
                                    playerCard(3) = 13
                                    card(8) = 26
                                Case 27
                                    Me.picPlayerCard4.Image = card27
                                    playerCard(3) = 1
                                    card(8) = 27
                                Case 28
                                    Me.picPlayerCard4.Image = card28
                                    playerCard(3) = 2
                                    card(8) = 28
                                Case 29
                                    Me.picPlayerCard4.Image = card29
                                    playerCard(3) = 3
                                    card(8) = 29
                                Case 30
                                    Me.picPlayerCard4.Image = card30
                                    playerCard(3) = 4
                                    card(8) = 30
                                Case 31
                                    Me.picPlayerCard4.Image = card31
                                    playerCard(3) = 5
                                    card(8) = 31
                                Case 32
                                    Me.picPlayerCard4.Image = card32
                                    playerCard(3) = 6
                                    card(8) = 32
                                Case 33
                                    Me.picPlayerCard4.Image = card33
                                    playerCard(3) = 7
                                    card(8) = 33
                                Case 34
                                    Me.picPlayerCard4.Image = card34
                                    playerCard(3) = 8
                                    card(8) = 34
                                Case 35
                                    Me.picPlayerCard4.Image = card35
                                    playerCard(3) = 9
                                    card(8) = 35
                                Case 36
                                    Me.picPlayerCard4.Image = card36
                                    playerCard(3) = 10
                                    card(8) = 36
                                Case 37
                                    Me.picPlayerCard4.Image = card37
                                    playerCard(3) = 11
                                    card(8) = 37
                                Case 38
                                    Me.picPlayerCard4.Image = card38
                                    playerCard(3) = 12
                                    card(8) = 38
                                Case 39
                                    Me.picPlayerCard4.Image = card39
                                    playerCard(3) = 13
                                    card(8) = 39
                                Case 40
                                    Me.picPlayerCard4.Image = card40
                                    playerCard(3) = 1
                                    card(8) = 40
                                Case 41
                                    Me.picPlayerCard4.Image = card41
                                    playerCard(3) = 2
                                    card(8) = 41
                                Case 42
                                    Me.picPlayerCard4.Image = card42
                                    playerCard(3) = 3
                                    card(8) = 42
                                Case 43
                                    Me.picPlayerCard4.Image = card43
                                    playerCard(3) = 4
                                    card(8) = 43
                                Case 44
                                    Me.picPlayerCard4.Image = card44
                                    playerCard(3) = 5
                                    card(8) = 44
                                Case 45
                                    Me.picPlayerCard4.Image = card45
                                    playerCard(3) = 6
                                    card(8) = 45
                                Case 46
                                    Me.picPlayerCard4.Image = card46
                                    playerCard(3) = 7
                                    card(8) = 46
                                Case 47
                                    Me.picPlayerCard4.Image = card47
                                    playerCard(3) = 8
                                    card(8) = 47
                                Case 48
                                    Me.picPlayerCard4.Image = card48
                                    playerCard(3) = 9
                                    card(8) = 48
                                Case 49
                                    Me.picPlayerCard4.Image = card49
                                    playerCard(3) = 10
                                    card(8) = 49
                                Case 50
                                    Me.picPlayerCard4.Image = card50
                                    playerCard(3) = 11
                                    card(8) = 50
                                Case 51
                                    Me.picPlayerCard4.Image = card51
                                    playerCard(3) = 12
                                    card(8) = 51
                                Case 52
                                    Me.picPlayerCard4.Image = card52
                                    playerCard(3) = 13
                                    card(8) = 52
                            End Select
                        ElseIf card(10) = card(9) Then
                            Select Case newNum
                                Case 1
                                    Me.picPlayerCard5.Image = card1
                                    playerCard(4) = 1
                                    card(9) = 1
                                Case 2
                                    Me.picPlayerCard5.Image = card2
                                    playerCard(4) = 2
                                    card(9) = 2
                                Case 3
                                    Me.picPlayerCard5.Image = card3
                                    playerCard(4) = 3
                                    card(9) = 3
                                Case 4
                                    Me.picPlayerCard5.Image = card4
                                    playerCard(4) = 4
                                    card(9) = 4
                                Case 5
                                    Me.picPlayerCard5.Image = card5
                                    playerCard(4) = 5
                                    card(9) = 5
                                Case 6
                                    Me.picPlayerCard5.Image = card6
                                    playerCard(4) = 6
                                    card(9) = 6
                                Case 7
                                    Me.picPlayerCard5.Image = card7
                                    playerCard(4) = 7
                                    card(9) = 7
                                Case 8
                                    Me.picPlayerCard5.Image = card8
                                    playerCard(4) = 8
                                    card(9) = 8
                                Case 9
                                    Me.picPlayerCard5.Image = card9
                                    playerCard(4) = 9
                                    card(9) = 9
                                Case 10
                                    Me.picPlayerCard5.Image = card10
                                    playerCard(4) = 10
                                    card(9) = 10
                                Case 11
                                    Me.picPlayerCard5.Image = card11
                                    playerCard(4) = 11
                                    card(9) = 11
                                Case 12
                                    Me.picPlayerCard5.Image = card12
                                    playerCard(4) = 12
                                    card(9) = 12
                                Case 13
                                    Me.picPlayerCard5.Image = card13
                                    playerCard(4) = 13
                                    card(9) = 13
                                Case 14
                                    Me.picPlayerCard5.Image = card14
                                    playerCard(4) = 1
                                    card(9) = 14
                                Case 15
                                    Me.picPlayerCard5.Image = card15
                                    playerCard(4) = 2
                                    card(9) = 15
                                Case 16
                                    Me.picPlayerCard5.Image = card16
                                    playerCard(4) = 3
                                    card(9) = 16
                                Case 17
                                    Me.picPlayerCard5.Image = card17
                                    playerCard(4) = 4
                                    card(9) = 17
                                Case 18
                                    Me.picPlayerCard5.Image = card18
                                    playerCard(4) = 5
                                    card(9) = 18
                                Case 19
                                    Me.picPlayerCard5.Image = card19
                                    playerCard(4) = 6
                                    card(9) = 19
                                Case 20
                                    Me.picPlayerCard5.Image = card20
                                    playerCard(4) = 7
                                    card(9) = 20
                                Case 21
                                    Me.picPlayerCard5.Image = card21
                                    playerCard(4) = 8
                                    card(9) = 21
                                Case 22
                                    Me.picPlayerCard5.Image = card22
                                    playerCard(4) = 9
                                    card(9) = 22
                                Case 23
                                    Me.picPlayerCard5.Image = card23
                                    playerCard(4) = 10
                                    card(9) = 23
                                Case 24
                                    Me.picPlayerCard5.Image = card24
                                    playerCard(4) = 11
                                    card(9) = 24
                                Case 25
                                    Me.picPlayerCard5.Image = card25
                                    playerCard(4) = 12
                                    card(9) = 25
                                Case 26
                                    Me.picPlayerCard5.Image = card26
                                    playerCard(4) = 13
                                    card(9) = 26
                                Case 27
                                    Me.picPlayerCard5.Image = card27
                                    playerCard(4) = 1
                                    card(9) = 27
                                Case 28
                                    Me.picPlayerCard5.Image = card28
                                    playerCard(4) = 2
                                    card(9) = 28
                                Case 29
                                    Me.picPlayerCard5.Image = card29
                                    playerCard(4) = 3
                                    card(9) = 29
                                Case 30
                                    Me.picPlayerCard5.Image = card30
                                    playerCard(4) = 4
                                    card(9) = 30
                                Case 31
                                    Me.picPlayerCard5.Image = card31
                                    playerCard(4) = 5
                                    card(9) = 31
                                Case 32
                                    Me.picPlayerCard5.Image = card32
                                    playerCard(4) = 6
                                    card(9) = 32
                                Case 33
                                    Me.picPlayerCard5.Image = card33
                                    playerCard(4) = 7
                                    card(9) = 33
                                Case 34
                                    Me.picPlayerCard5.Image = card34
                                    playerCard(4) = 8
                                    card(9) = 34
                                Case 35
                                    Me.picPlayerCard5.Image = card35
                                    playerCard(4) = 9
                                    card(9) = 35
                                Case 36
                                    Me.picPlayerCard5.Image = card36
                                    playerCard(4) = 10
                                    card(9) = 36
                                Case 37
                                    Me.picPlayerCard5.Image = card37
                                    playerCard(4) = 11
                                    card(9) = 37
                                Case 38
                                    Me.picPlayerCard5.Image = card38
                                    playerCard(4) = 12
                                    card(9) = 38
                                Case 39
                                    Me.picPlayerCard5.Image = card39
                                    playerCard(4) = 13
                                    card(9) = 39
                                Case 40
                                    Me.picPlayerCard5.Image = card40
                                    playerCard(4) = 1
                                    card(9) = 40
                                Case 41
                                    Me.picPlayerCard5.Image = card41
                                    playerCard(4) = 2
                                    card(9) = 41
                                Case 42
                                    Me.picPlayerCard5.Image = card42
                                    playerCard(4) = 3
                                    card(9) = 42
                                Case 43
                                    Me.picPlayerCard5.Image = card43
                                    playerCard(4) = 4
                                    card(9) = 43
                                Case 44
                                    Me.picPlayerCard5.Image = card44
                                    playerCard(4) = 5
                                    card(9) = 44
                                Case 45
                                    Me.picPlayerCard5.Image = card45
                                    playerCard(4) = 6
                                    card(9) = 45
                                Case 46
                                    Me.picPlayerCard5.Image = card46
                                    playerCard(4) = 7
                                    card(9) = 46
                                Case 47
                                    Me.picPlayerCard5.Image = card47
                                    playerCard(4) = 8
                                    card(9) = 47
                                Case 48
                                    Me.picPlayerCard5.Image = card48
                                    playerCard(4) = 9
                                    card(9) = 48
                                Case 49
                                    Me.picPlayerCard5.Image = card49
                                    playerCard(4) = 10
                                    card(9) = 49
                                Case 50
                                    Me.picPlayerCard5.Image = card50
                                    playerCard(4) = 11
                                    card(9) = 50
                                Case 51
                                    Me.picPlayerCard5.Image = card51
                                    playerCard(4) = 12
                                    card(9) = 51
                                Case 52
                                    Me.picPlayerCard5.Image = card52
                                    playerCard(4) = 13
                                    card(9) = 52
                            End Select
                        End If

                        Used.Add(newNum)

                    End If
                Loop While OK = False
            End If
        End If

        playerCardsLeftTotal = playerCardsLeft1 + playerCardsLeft2 + playerCardsLeft3 + playerCardsLeft4 + playerCardsLeft5
        Me.lblPlayerCardsLeftTotal.Text = playerName & "'s" & vbCrLf & "Total Cards Left:" & vbCrLf & playerCardsLeftTotal
        Me.lblPlayerCardsLeft1.Text = "Cards Left:" & playerCardsLeft1
        Me.lblPlayerCardsLeft2.Text = "Cards Left:" & playerCardsLeft2
        Me.lblPlayerCardsLeft3.Text = "Cards Left:" & playerCardsLeft3
        Me.lblPlayerCardsLeft4.Text = "Cards Left:" & playerCardsLeft4
        Me.lblPlayerCardsLeft5.Text = "Cards Left:" & playerCardsLeft5

    End Sub

    Private Sub picPlayerPile_DragDrop(sender As Object, e As DragEventArgs) Handles picPlayerPile.DragDrop
        Dim imageMove As Bitmap = DirectCast(e.Data.GetData(DataFormats.Bitmap, True), Bitmap)
        Dim newNum As Integer
        Dim OK As Boolean = False

        If cardCheck = 1 Then
            For num As Integer = 1 To 13
                If playerPile = num And (playerCard(0) = num + 1 Or playerCard(0) = num - 1) Then
                    picPlayerPile.Image = imageMove
                    card(11) = card(5)
                    playerPile = playerCard(0)
                    If playerCardsLeft1 > 0 Then
                        playerCardsLeft1 -= 1
                    End If
                ElseIf playerPile = 1 And (playerCard(0) = 2 Or playerCard(0) = 13) Then
                    picPlayerPile.Image = imageMove
                    card(11) = card(5)
                    playerPile = playerCard(0)
                    If playerCardsLeft1 > 0 Then
                        playerCardsLeft1 -= 1
                    End If
                ElseIf playerPile = 13 And (playerCard(0) = 1 Or playerCard(0) = 12) Then
                    picPlayerPile.Image = imageMove
                    card(11) = card(5)
                    playerPile = playerCard(0)
                    If playerCardsLeft1 > 0 Then
                        playerCardsLeft1 -= 1
                    End If
                End If
            Next
        ElseIf cardCheck = 2 Then
            For num As Integer = 1 To 13
                If playerPile = num And (playerCard(1) = num + 1 Or playerCard(1) = num - 1) Then
                    picPlayerPile.Image = imageMove
                    card(11) = card(6)
                    playerPile = playerCard(1)
                    If playerCardsLeft2 > 0 Then
                        playerCardsLeft2 -= 1
                    End If
                ElseIf playerPile = 1 And (playerCard(1) = 2 Or playerCard(1) = 13) Then
                    picPlayerPile.Image = imageMove
                    card(11) = card(6)
                    playerPile = playerCard(1)
                    If playerCardsLeft2 > 0 Then
                        playerCardsLeft2 -= 1
                    End If
                ElseIf playerPile = 13 And (playerCard(1) = 1 Or playerCard(1) = 12) Then
                    picPlayerPile.Image = imageMove
                    card(11) = card(6)
                    playerPile = playerCard(1)
                    If playerCardsLeft2 > 0 Then
                        playerCardsLeft2 -= 1
                    End If
                End If
            Next
        ElseIf cardCheck = 3 Then
            For num As Integer = 1 To 13
                If playerPile = num And (playerCard(2) = num + 1 Or playerCard(2) = num - 1) Then
                    picPlayerPile.Image = imageMove
                    card(11) = card(7)
                    playerPile = playerCard(2)
                    If playerCardsLeft3 > 0 Then
                        playerCardsLeft3 -= 1
                    End If
                ElseIf playerPile = 1 And (playerCard(2) = 2 Or playerCard(2) = 13) Then
                    picPlayerPile.Image = imageMove
                    card(11) = card(7)
                    compPile = playerCard(2)
                    If playerCardsLeft3 > 0 Then
                        playerCardsLeft3 -= 1
                    End If
                ElseIf playerPile = 13 And (playerCard(2) = 1 Or playerCard(2) = 12) Then
                    picPlayerPile.Image = imageMove
                    card(11) = card(7)
                    playerPile = playerCard(2)
                    If playerCardsLeft3 > 0 Then
                        playerCardsLeft3 -= 1
                    End If
                End If
            Next
        ElseIf cardCheck = 4 Then
            For num As Integer = 1 To 13
                If playerPile = num And (playerCard(3) = num + 1 Or playerCard(3) = num - 1) Then
                    picPlayerPile.Image = imageMove
                    card(11) = card(8)
                    playerPile = playerCard(3)
                    If playerCardsLeft4 > 0 Then
                        playerCardsLeft4 -= 1
                    End If
                ElseIf playerPile = 1 And (playerCard(3) = 2 Or playerCard(3) = 13) Then
                    picPlayerPile.Image = imageMove
                    card(11) = card(8)
                    playerPile = playerCard(3)
                    If playerCardsLeft4 > 0 Then
                        playerCardsLeft4 -= 1
                    End If
                ElseIf playerPile = 13 And (playerCard(3) = 1 Or playerCard(3) = 12) Then
                    picPlayerPile.Image = imageMove
                    card(11) = card(8)
                    playerPile = playerCard(3)
                    If playerCardsLeft4 > 0 Then
                        playerCardsLeft4 -= 1
                    End If
                End If
            Next
        ElseIf cardCheck = 5 Then
            For num As Integer = 1 To 13
                If playerPile = num And (playerCard(4) = num + 1 Or playerCard(4) = num - 1) Then
                    picPlayerPile.Image = imageMove
                    card(11) = card(9)
                    playerPile = playerCard(4)
                    If playerCardsLeft5 > 0 Then
                        playerCardsLeft5 -= 1
                    End If
                ElseIf playerPile = 1 And (playerCard(4) = 2 Or playerCard(4) = 13) Then
                    picPlayerPile.Image = imageMove
                    card(11) = card(9)
                    playerPile = playerCard(4)
                    If playerCardsLeft5 > 0 Then
                        playerCardsLeft5 -= 1
                    End If
                ElseIf playerPile = 13 And (playerCard(4) = 1 Or playerCard(4) = 12) Then
                    picPlayerPile.Image = imageMove
                    card(11) = card(9)
                    playerPile = playerCard(4)
                    If playerCardsLeft5 > 0 Then
                        playerCardsLeft5 -= 1
                    End If
                End If
            Next
        End If

        If card(11) = card(5) Or card(11) = card(6) Or card(11) = card(7) Or card(11) = card(8) Or card(11) = card(9) Then
            'allow generate new card if playerPile is same as any of the playerCards
            If playerCardsLeft1 > 0 Or playerCardsLeft2 > 0 Or playerCardsLeft3 > 0 Or playerCardsLeft4 > 0 Or playerCardsLeft5 > 0 Then
                Do                                                      'generate new card only if playerCardsLeft is more than 0
                    Randomize()
                    newNum = Int(max * Rnd() + min)

                    If Used.Contains(newNum) Then
                        OK = False
                    Else
                        OK = True

                        If card(11) = card(5) Then
                            Select Case newNum
                                Case 1
                                    Me.picPlayerCard1.Image = card1
                                    playerCard(0) = 1
                                    card(5) = 1
                                Case 2
                                    Me.picPlayerCard1.Image = card2
                                    playerCard(0) = 2
                                    card(5) = 2
                                Case 3
                                    Me.picPlayerCard1.Image = card3
                                    playerCard(0) = 3
                                    card(5) = 3
                                Case 4
                                    Me.picPlayerCard1.Image = card4
                                    playerCard(0) = 4
                                    card(5) = 4
                                Case 5
                                    Me.picPlayerCard1.Image = card5
                                    playerCard(0) = 5
                                    card(5) = 5
                                Case 6
                                    Me.picPlayerCard1.Image = card6
                                    playerCard(0) = 6
                                    card(5) = 6
                                Case 7
                                    Me.picPlayerCard1.Image = card7
                                    playerCard(0) = 7
                                    card(5) = 7
                                Case 8
                                    Me.picPlayerCard1.Image = card8
                                    playerCard(0) = 8
                                    card(5) = 8
                                Case 9
                                    Me.picPlayerCard1.Image = card9
                                    playerCard(0) = 9
                                    card(5) = 9
                                Case 10
                                    Me.picPlayerCard1.Image = card10
                                    playerCard(0) = 10
                                    card(5) = 10
                                Case 11
                                    Me.picPlayerCard1.Image = card11
                                    playerCard(0) = 11
                                    card(5) = 11
                                Case 12
                                    Me.picPlayerCard1.Image = card12
                                    playerCard(0) = 12
                                    card(5) = 12
                                Case 13
                                    Me.picPlayerCard1.Image = card13
                                    playerCard(0) = 13
                                    card(5) = 13
                                Case 14
                                    Me.picPlayerCard1.Image = card14
                                    playerCard(0) = 1
                                    card(5) = 14
                                Case 15
                                    Me.picPlayerCard1.Image = card15
                                    playerCard(0) = 2
                                    card(5) = 15
                                Case 16
                                    Me.picPlayerCard1.Image = card16
                                    playerCard(0) = 3
                                    card(5) = 16
                                Case 17
                                    Me.picPlayerCard1.Image = card17
                                    playerCard(0) = 4
                                    card(5) = 17
                                Case 18
                                    Me.picPlayerCard1.Image = card18
                                    playerCard(0) = 5
                                    card(5) = 18
                                Case 19
                                    Me.picPlayerCard1.Image = card19
                                    playerCard(0) = 6
                                    card(5) = 19
                                Case 20
                                    Me.picPlayerCard1.Image = card20
                                    playerCard(0) = 7
                                    card(5) = 20
                                Case 21
                                    Me.picPlayerCard1.Image = card21
                                    playerCard(0) = 8
                                    card(5) = 21
                                Case 22
                                    Me.picPlayerCard1.Image = card22
                                    playerCard(0) = 9
                                    card(5) = 22
                                Case 23
                                    Me.picPlayerCard1.Image = card23
                                    playerCard(0) = 10
                                    card(5) = 23
                                Case 24
                                    Me.picPlayerCard1.Image = card24
                                    playerCard(0) = 11
                                    card(5) = 24
                                Case 25
                                    Me.picPlayerCard1.Image = card25
                                    playerCard(0) = 12
                                    card(5) = 25
                                Case 26
                                    Me.picPlayerCard1.Image = card26
                                    playerCard(0) = 13
                                    card(5) = 26
                                Case 27
                                    Me.picPlayerCard1.Image = card27
                                    playerCard(0) = 1
                                    card(5) = 27
                                Case 28
                                    Me.picPlayerCard1.Image = card28
                                    playerCard(0) = 2
                                    card(5) = 28
                                Case 29
                                    Me.picPlayerCard1.Image = card29
                                    playerCard(0) = 3
                                    card(5) = 29
                                Case 30
                                    Me.picPlayerCard1.Image = card30
                                    playerCard(0) = 4
                                    card(5) = 30
                                Case 31
                                    Me.picPlayerCard1.Image = card31
                                    playerCard(0) = 5
                                    card(5) = 31
                                Case 32
                                    Me.picPlayerCard1.Image = card32
                                    playerCard(0) = 6
                                    card(5) = 32
                                Case 33
                                    Me.picPlayerCard1.Image = card33
                                    playerCard(0) = 7
                                    card(5) = 33
                                Case 34
                                    Me.picPlayerCard1.Image = card34
                                    playerCard(0) = 8
                                    card(5) = 34
                                Case 35
                                    Me.picPlayerCard1.Image = card35
                                    playerCard(0) = 9
                                    card(5) = 35
                                Case 36
                                    Me.picPlayerCard1.Image = card36
                                    playerCard(0) = 10
                                    card(5) = 36
                                Case 37
                                    Me.picPlayerCard1.Image = card37
                                    playerCard(0) = 11
                                    card(5) = 37
                                Case 38
                                    Me.picPlayerCard1.Image = card38
                                    playerCard(0) = 12
                                    card(5) = 38
                                Case 39
                                    Me.picPlayerCard1.Image = card39
                                    playerCard(0) = 13
                                    card(5) = 39
                                Case 40
                                    Me.picPlayerCard1.Image = card40
                                    playerCard(0) = 1
                                    card(5) = 40
                                Case 41
                                    Me.picPlayerCard1.Image = card41
                                    playerCard(0) = 2
                                    card(5) = 41
                                Case 42
                                    Me.picPlayerCard1.Image = card42
                                    playerCard(0) = 3
                                    card(5) = 42
                                Case 43
                                    Me.picPlayerCard1.Image = card43
                                    playerCard(0) = 4
                                    card(5) = 43
                                Case 44
                                    Me.picPlayerCard1.Image = card44
                                    playerCard(0) = 5
                                    card(5) = 44
                                Case 45
                                    Me.picPlayerCard1.Image = card45
                                    playerCard(0) = 6
                                    card(5) = 45
                                Case 46
                                    Me.picPlayerCard1.Image = card46
                                    playerCard(0) = 7
                                    card(5) = 46
                                Case 47
                                    Me.picPlayerCard1.Image = card47
                                    playerCard(0) = 8
                                    card(5) = 47
                                Case 48
                                    Me.picPlayerCard1.Image = card48
                                    playerCard(0) = 9
                                    card(5) = 48
                                Case 49
                                    Me.picPlayerCard1.Image = card49
                                    playerCard(0) = 10
                                    card(5) = 49
                                Case 50
                                    Me.picPlayerCard1.Image = card50
                                    playerCard(0) = 11
                                    card(5) = 50
                                Case 51
                                    Me.picPlayerCard1.Image = card51
                                    playerCard(0) = 12
                                    card(5) = 51
                                Case 52
                                    Me.picPlayerCard1.Image = card52
                                    playerCard(0) = 13
                                    card(5) = 52
                            End Select
                        ElseIf card(11) = card(6) Then
                            Select Case newNum
                                Case 1
                                    Me.picPlayerCard2.Image = card1
                                    playerCard(1) = 1
                                    card(6) = 1
                                Case 2
                                    Me.picPlayerCard2.Image = card2
                                    playerCard(1) = 2
                                    card(6) = 2
                                Case 3
                                    Me.picPlayerCard2.Image = card3
                                    playerCard(1) = 3
                                    card(6) = 3
                                Case 4
                                    Me.picPlayerCard2.Image = card4
                                    playerCard(1) = 4
                                    card(6) = 4
                                Case 5
                                    Me.picPlayerCard2.Image = card5
                                    playerCard(1) = 5
                                    card(6) = 5
                                Case 6
                                    Me.picPlayerCard2.Image = card6
                                    playerCard(1) = 6
                                    card(6) = 6
                                Case 7
                                    Me.picPlayerCard2.Image = card7
                                    playerCard(1) = 7
                                    card(6) = 7
                                Case 8
                                    Me.picPlayerCard2.Image = card8
                                    playerCard(1) = 8
                                    card(6) = 8
                                Case 9
                                    Me.picPlayerCard2.Image = card9
                                    playerCard(1) = 9
                                    card(6) = 9
                                Case 10
                                    Me.picPlayerCard2.Image = card10
                                    playerCard(1) = 10
                                    card(6) = 10
                                Case 11
                                    Me.picPlayerCard2.Image = card11
                                    playerCard(1) = 11
                                    card(6) = 11
                                Case 12
                                    Me.picPlayerCard2.Image = card12
                                    playerCard(1) = 12
                                    card(6) = 12
                                Case 13
                                    Me.picPlayerCard2.Image = card13
                                    playerCard(1) = 13
                                    card(6) = 13
                                Case 14
                                    Me.picPlayerCard2.Image = card14
                                    playerCard(1) = 1
                                    card(6) = 14
                                Case 15
                                    Me.picPlayerCard2.Image = card15
                                    playerCard(1) = 2
                                    card(6) = 15
                                Case 16
                                    Me.picPlayerCard2.Image = card16
                                    playerCard(1) = 3
                                    card(6) = 16
                                Case 17
                                    Me.picPlayerCard2.Image = card17
                                    playerCard(1) = 4
                                    card(6) = 17
                                Case 18
                                    Me.picPlayerCard2.Image = card18
                                    playerCard(1) = 5
                                    card(6) = 18
                                Case 19
                                    Me.picPlayerCard2.Image = card19
                                    playerCard(1) = 6
                                    card(6) = 19
                                Case 20
                                    Me.picPlayerCard2.Image = card20
                                    playerCard(1) = 7
                                    card(6) = 20
                                Case 21
                                    Me.picPlayerCard2.Image = card21
                                    playerCard(1) = 8
                                    card(6) = 21
                                Case 22
                                    Me.picPlayerCard2.Image = card22
                                    playerCard(1) = 9
                                    card(6) = 22
                                Case 23
                                    Me.picPlayerCard2.Image = card23
                                    playerCard(1) = 10
                                    card(6) = 23
                                Case 24
                                    Me.picPlayerCard2.Image = card24
                                    playerCard(1) = 11
                                    card(6) = 24
                                Case 25
                                    Me.picPlayerCard2.Image = card25
                                    playerCard(1) = 12
                                    card(6) = 25
                                Case 26
                                    Me.picPlayerCard2.Image = card26
                                    playerCard(1) = 13
                                    card(6) = 26
                                Case 27
                                    Me.picPlayerCard2.Image = card27
                                    playerCard(1) = 1
                                    card(6) = 27
                                Case 28
                                    Me.picPlayerCard2.Image = card28
                                    playerCard(1) = 2
                                    card(6) = 28
                                Case 29
                                    Me.picPlayerCard2.Image = card29
                                    playerCard(1) = 3
                                    card(6) = 29
                                Case 30
                                    Me.picPlayerCard2.Image = card30
                                    playerCard(1) = 4
                                    card(6) = 30
                                Case 31
                                    Me.picPlayerCard2.Image = card31
                                    playerCard(1) = 5
                                    card(6) = 31
                                Case 32
                                    Me.picPlayerCard2.Image = card32
                                    playerCard(1) = 6
                                    card(6) = 32
                                Case 33
                                    Me.picPlayerCard2.Image = card33
                                    playerCard(1) = 7
                                    card(6) = 33
                                Case 34
                                    Me.picPlayerCard2.Image = card34
                                    playerCard(1) = 8
                                    card(6) = 34
                                Case 35
                                    Me.picPlayerCard2.Image = card35
                                    playerCard(1) = 9
                                    card(6) = 35
                                Case 36
                                    Me.picPlayerCard2.Image = card36
                                    playerCard(1) = 10
                                    card(6) = 36
                                Case 37
                                    Me.picPlayerCard2.Image = card37
                                    playerCard(1) = 11
                                    card(6) = 37
                                Case 38
                                    Me.picPlayerCard2.Image = card38
                                    playerCard(1) = 12
                                    card(6) = 38
                                Case 39
                                    Me.picPlayerCard2.Image = card39
                                    playerCard(1) = 13
                                    card(6) = 39
                                Case 40
                                    Me.picPlayerCard2.Image = card40
                                    playerCard(1) = 1
                                    card(6) = 40
                                Case 41
                                    Me.picPlayerCard2.Image = card41
                                    playerCard(1) = 2
                                    card(6) = 41
                                Case 42
                                    Me.picPlayerCard2.Image = card42
                                    playerCard(1) = 3
                                    card(6) = 42
                                Case 43
                                    Me.picPlayerCard2.Image = card43
                                    playerCard(1) = 4
                                    card(6) = 43
                                Case 44
                                    Me.picPlayerCard2.Image = card44
                                    playerCard(1) = 5
                                    card(6) = 44
                                Case 45
                                    Me.picPlayerCard2.Image = card45
                                    playerCard(1) = 6
                                    card(6) = 45
                                Case 46
                                    Me.picPlayerCard2.Image = card46
                                    playerCard(1) = 7
                                    card(6) = 46
                                Case 47
                                    Me.picPlayerCard2.Image = card47
                                    playerCard(1) = 8
                                    card(6) = 47
                                Case 48
                                    Me.picPlayerCard2.Image = card48
                                    playerCard(1) = 9
                                    card(6) = 48
                                Case 49
                                    Me.picPlayerCard2.Image = card49
                                    playerCard(1) = 10
                                    card(6) = 49
                                Case 50
                                    Me.picPlayerCard2.Image = card50
                                    playerCard(1) = 11
                                    card(6) = 50
                                Case 51
                                    Me.picPlayerCard2.Image = card51
                                    playerCard(1) = 12
                                    card(6) = 51
                                Case 52
                                    Me.picPlayerCard2.Image = card52
                                    playerCard(1) = 13
                                    card(6) = 52
                            End Select
                        ElseIf card(11) = card(7) Then
                            Select Case newNum
                                Case 1
                                    Me.picPlayerCard3.Image = card1
                                    playerCard(2) = 1
                                    card(7) = 1
                                Case 2
                                    Me.picPlayerCard3.Image = card2
                                    playerCard(2) = 2
                                    card(7) = 2
                                Case 3
                                    Me.picPlayerCard3.Image = card3
                                    playerCard(2) = 3
                                    card(7) = 3
                                Case 4
                                    Me.picPlayerCard3.Image = card4
                                    playerCard(2) = 4
                                    card(7) = 4
                                Case 5
                                    Me.picPlayerCard3.Image = card5
                                    playerCard(2) = 5
                                    card(7) = 5
                                Case 6
                                    Me.picPlayerCard3.Image = card6
                                    playerCard(2) = 6
                                    card(7) = 6
                                Case 7
                                    Me.picPlayerCard3.Image = card7
                                    playerCard(2) = 7
                                    card(7) = 7
                                Case 8
                                    Me.picPlayerCard3.Image = card8
                                    playerCard(2) = 8
                                    card(7) = 8
                                Case 9
                                    Me.picPlayerCard3.Image = card9
                                    playerCard(2) = 9
                                    card(7) = 9
                                Case 10
                                    Me.picPlayerCard3.Image = card10
                                    playerCard(2) = 10
                                    card(7) = 10
                                Case 11
                                    Me.picPlayerCard3.Image = card11
                                    playerCard(2) = 11
                                    card(7) = 11
                                Case 12
                                    Me.picPlayerCard3.Image = card12
                                    playerCard(2) = 12
                                    card(7) = 12
                                Case 13
                                    Me.picPlayerCard3.Image = card13
                                    playerCard(2) = 13
                                    card(7) = 13
                                Case 14
                                    Me.picPlayerCard3.Image = card14
                                    playerCard(2) = 1
                                    card(7) = 14
                                Case 15
                                    Me.picPlayerCard3.Image = card15
                                    playerCard(2) = 2
                                    card(7) = 15
                                Case 16
                                    Me.picPlayerCard3.Image = card16
                                    playerCard(2) = 3
                                    card(7) = 16
                                Case 17
                                    Me.picPlayerCard3.Image = card17
                                    playerCard(2) = 4
                                    card(7) = 17
                                Case 18
                                    Me.picPlayerCard3.Image = card18
                                    playerCard(2) = 5
                                    card(7) = 18
                                Case 19
                                    Me.picPlayerCard3.Image = card19
                                    playerCard(2) = 6
                                    card(7) = 19
                                Case 20
                                    Me.picPlayerCard3.Image = card20
                                    playerCard(2) = 7
                                    card(7) = 20
                                Case 21
                                    Me.picPlayerCard3.Image = card21
                                    playerCard(2) = 8
                                    card(7) = 21
                                Case 22
                                    Me.picPlayerCard3.Image = card22
                                    playerCard(2) = 9
                                    card(7) = 22
                                Case 23
                                    Me.picPlayerCard3.Image = card23
                                    playerCard(2) = 10
                                    card(7) = 23
                                Case 24
                                    Me.picPlayerCard3.Image = card24
                                    playerCard(2) = 11
                                    card(7) = 24
                                Case 25
                                    Me.picPlayerCard3.Image = card25
                                    playerCard(2) = 12
                                    card(7) = 25
                                Case 26
                                    Me.picPlayerCard3.Image = card26
                                    playerCard(2) = 13
                                    card(7) = 26
                                Case 27
                                    Me.picPlayerCard3.Image = card27
                                    playerCard(2) = 1
                                    card(7) = 27
                                Case 28
                                    Me.picPlayerCard3.Image = card28
                                    playerCard(2) = 2
                                    card(7) = 28
                                Case 29
                                    Me.picPlayerCard3.Image = card29
                                    playerCard(2) = 3
                                    card(7) = 29
                                Case 30
                                    Me.picPlayerCard3.Image = card30
                                    playerCard(2) = 4
                                    card(7) = 30
                                Case 31
                                    Me.picPlayerCard3.Image = card31
                                    playerCard(2) = 5
                                    card(7) = 31
                                Case 32
                                    Me.picPlayerCard3.Image = card32
                                    playerCard(2) = 6
                                    card(7) = 32
                                Case 33
                                    Me.picPlayerCard3.Image = card33
                                    playerCard(2) = 7
                                    card(7) = 33
                                Case 34
                                    Me.picPlayerCard3.Image = card34
                                    playerCard(2) = 8
                                    card(7) = 34
                                Case 35
                                    Me.picPlayerCard3.Image = card35
                                    playerCard(2) = 9
                                    card(7) = 35
                                Case 36
                                    Me.picPlayerCard3.Image = card36
                                    playerCard(2) = 10
                                    card(7) = 36
                                Case 37
                                    Me.picPlayerCard3.Image = card37
                                    playerCard(2) = 11
                                    card(7) = 37
                                Case 38
                                    Me.picPlayerCard3.Image = card38
                                    playerCard(2) = 12
                                    card(7) = 38
                                Case 39
                                    Me.picPlayerCard3.Image = card39
                                    playerCard(2) = 13
                                    card(7) = 39
                                Case 40
                                    Me.picPlayerCard3.Image = card40
                                    playerCard(2) = 1
                                    card(7) = 40
                                Case 41
                                    Me.picPlayerCard3.Image = card41
                                    playerCard(2) = 2
                                    card(7) = 41
                                Case 42
                                    Me.picPlayerCard3.Image = card42
                                    playerCard(2) = 3
                                    card(7) = 42
                                Case 43
                                    Me.picPlayerCard3.Image = card43
                                    playerCard(2) = 4
                                    card(7) = 43
                                Case 44
                                    Me.picPlayerCard3.Image = card44
                                    playerCard(2) = 5
                                    card(7) = 44
                                Case 45
                                    Me.picPlayerCard3.Image = card45
                                    playerCard(2) = 6
                                    card(7) = 45
                                Case 46
                                    Me.picPlayerCard3.Image = card46
                                    playerCard(2) = 7
                                    card(7) = 46
                                Case 47
                                    Me.picPlayerCard3.Image = card47
                                    playerCard(2) = 8
                                    card(7) = 47
                                Case 48
                                    Me.picPlayerCard3.Image = card48
                                    playerCard(2) = 9
                                    card(7) = 48
                                Case 49
                                    Me.picPlayerCard3.Image = card49
                                    playerCard(2) = 10
                                    card(7) = 49
                                Case 50
                                    Me.picPlayerCard3.Image = card50
                                    playerCard(2) = 11
                                    card(7) = 50
                                Case 51
                                    Me.picPlayerCard3.Image = card51
                                    playerCard(2) = 12
                                    card(7) = 51
                                Case 52
                                    Me.picPlayerCard3.Image = card52
                                    playerCard(2) = 13
                                    card(7) = 52
                            End Select
                        ElseIf card(11) = card(8) Then
                            Select Case newNum
                                Case 1
                                    Me.picPlayerCard4.Image = card1
                                    playerCard(3) = 1
                                    card(8) = 1
                                Case 2
                                    Me.picPlayerCard4.Image = card2
                                    playerCard(3) = 2
                                    card(8) = 2
                                Case 3
                                    Me.picPlayerCard4.Image = card3
                                    playerCard(3) = 3
                                    card(8) = 3
                                Case 4
                                    Me.picPlayerCard4.Image = card4
                                    playerCard(3) = 4
                                    card(8) = 4
                                Case 5
                                    Me.picPlayerCard4.Image = card5
                                    playerCard(3) = 5
                                    card(8) = 5
                                Case 6
                                    Me.picPlayerCard4.Image = card6
                                    playerCard(3) = 6
                                    card(8) = 6
                                Case 7
                                    Me.picPlayerCard4.Image = card7
                                    playerCard(3) = 7
                                    card(8) = 7
                                Case 8
                                    Me.picPlayerCard4.Image = card8
                                    playerCard(3) = 8
                                    card(8) = 8
                                Case 9
                                    Me.picPlayerCard4.Image = card9
                                    playerCard(3) = 9
                                    card(8) = 9
                                Case 10
                                    Me.picPlayerCard4.Image = card10
                                    playerCard(3) = 10
                                    card(8) = 10
                                Case 11
                                    Me.picPlayerCard4.Image = card11
                                    playerCard(3) = 11
                                    card(8) = 11
                                Case 12
                                    Me.picPlayerCard4.Image = card12
                                    playerCard(3) = 12
                                    card(8) = 12
                                Case 13
                                    Me.picPlayerCard4.Image = card13
                                    playerCard(3) = 13
                                    card(8) = 13
                                Case 14
                                    Me.picPlayerCard4.Image = card14
                                    playerCard(3) = 1
                                    card(8) = 14
                                Case 15
                                    Me.picPlayerCard4.Image = card15
                                    playerCard(3) = 2
                                    card(8) = 15
                                Case 16
                                    Me.picPlayerCard4.Image = card16
                                    playerCard(3) = 3
                                    card(8) = 16
                                Case 17
                                    Me.picPlayerCard4.Image = card17
                                    playerCard(3) = 4
                                    card(8) = 17
                                Case 18
                                    Me.picPlayerCard4.Image = card18
                                    playerCard(3) = 5
                                    card(8) = 18
                                Case 19
                                    Me.picPlayerCard4.Image = card19
                                    playerCard(3) = 6
                                    card(8) = 19
                                Case 20
                                    Me.picPlayerCard4.Image = card20
                                    playerCard(3) = 7
                                    card(8) = 20
                                Case 21
                                    Me.picPlayerCard4.Image = card21
                                    playerCard(3) = 8
                                    card(8) = 21
                                Case 22
                                    Me.picPlayerCard4.Image = card22
                                    playerCard(3) = 9
                                    card(8) = 22
                                Case 23
                                    Me.picPlayerCard4.Image = card23
                                    playerCard(3) = 10
                                    card(8) = 23
                                Case 24
                                    Me.picPlayerCard4.Image = card24
                                    playerCard(3) = 11
                                    card(8) = 24
                                Case 25
                                    Me.picPlayerCard4.Image = card25
                                    playerCard(3) = 12
                                    card(8) = 25
                                Case 26
                                    Me.picPlayerCard4.Image = card26
                                    playerCard(3) = 13
                                    card(8) = 26
                                Case 27
                                    Me.picPlayerCard4.Image = card27
                                    playerCard(3) = 1
                                    card(8) = 27
                                Case 28
                                    Me.picPlayerCard4.Image = card28
                                    playerCard(3) = 2
                                    card(8) = 28
                                Case 29
                                    Me.picPlayerCard4.Image = card29
                                    playerCard(3) = 3
                                    card(8) = 29
                                Case 30
                                    Me.picPlayerCard4.Image = card30
                                    playerCard(3) = 4
                                    card(8) = 30
                                Case 31
                                    Me.picPlayerCard4.Image = card31
                                    playerCard(3) = 5
                                    card(8) = 31
                                Case 32
                                    Me.picPlayerCard4.Image = card32
                                    playerCard(3) = 6
                                    card(8) = 32
                                Case 33
                                    Me.picPlayerCard4.Image = card33
                                    playerCard(3) = 7
                                    card(8) = 33
                                Case 34
                                    Me.picPlayerCard4.Image = card34
                                    playerCard(3) = 8
                                    card(8) = 34
                                Case 35
                                    Me.picPlayerCard4.Image = card35
                                    playerCard(3) = 9
                                    card(8) = 35
                                Case 36
                                    Me.picPlayerCard4.Image = card36
                                    playerCard(3) = 10
                                    card(8) = 36
                                Case 37
                                    Me.picPlayerCard4.Image = card37
                                    playerCard(3) = 11
                                    card(8) = 37
                                Case 38
                                    Me.picPlayerCard4.Image = card38
                                    playerCard(3) = 12
                                    card(8) = 38
                                Case 39
                                    Me.picPlayerCard4.Image = card39
                                    playerCard(3) = 13
                                    card(8) = 39
                                Case 40
                                    Me.picPlayerCard4.Image = card40
                                    playerCard(3) = 1
                                    card(8) = 40
                                Case 41
                                    Me.picPlayerCard4.Image = card41
                                    playerCard(3) = 2
                                    card(8) = 41
                                Case 42
                                    Me.picPlayerCard4.Image = card42
                                    playerCard(3) = 3
                                    card(8) = 42
                                Case 43
                                    Me.picPlayerCard4.Image = card43
                                    playerCard(3) = 4
                                    card(8) = 43
                                Case 44
                                    Me.picPlayerCard4.Image = card44
                                    playerCard(3) = 5
                                    card(8) = 44
                                Case 45
                                    Me.picPlayerCard4.Image = card45
                                    playerCard(3) = 6
                                    card(8) = 45
                                Case 46
                                    Me.picPlayerCard4.Image = card46
                                    playerCard(3) = 7
                                    card(8) = 46
                                Case 47
                                    Me.picPlayerCard4.Image = card47
                                    playerCard(3) = 8
                                    card(8) = 47
                                Case 48
                                    Me.picPlayerCard4.Image = card48
                                    playerCard(3) = 9
                                    card(8) = 48
                                Case 49
                                    Me.picPlayerCard4.Image = card49
                                    playerCard(3) = 10
                                    card(8) = 49
                                Case 50
                                    Me.picPlayerCard4.Image = card50
                                    playerCard(3) = 11
                                    card(8) = 50
                                Case 51
                                    Me.picPlayerCard4.Image = card51
                                    playerCard(3) = 12
                                    card(8) = 51
                                Case 52
                                    Me.picPlayerCard4.Image = card52
                                    playerCard(3) = 13
                                    card(8) = 52
                            End Select
                        ElseIf card(11) = card(9) Then
                            Select Case newNum
                                Case 1
                                    Me.picPlayerCard5.Image = card1
                                    playerCard(4) = 1
                                    card(9) = 1
                                Case 2
                                    Me.picPlayerCard5.Image = card2
                                    playerCard(4) = 2
                                    card(9) = 2
                                Case 3
                                    Me.picPlayerCard5.Image = card3
                                    playerCard(4) = 3
                                    card(9) = 3
                                Case 4
                                    Me.picPlayerCard5.Image = card4
                                    playerCard(4) = 4
                                    card(9) = 4
                                Case 5
                                    Me.picPlayerCard5.Image = card5
                                    playerCard(4) = 5
                                    card(9) = 5
                                Case 6
                                    Me.picPlayerCard5.Image = card6
                                    playerCard(4) = 6
                                    card(9) = 6
                                Case 7
                                    Me.picPlayerCard5.Image = card7
                                    playerCard(4) = 7
                                    card(9) = 7
                                Case 8
                                    Me.picPlayerCard5.Image = card8
                                    playerCard(4) = 8
                                    card(9) = 8
                                Case 9
                                    Me.picPlayerCard5.Image = card9
                                    playerCard(4) = 9
                                    card(9) = 9
                                Case 10
                                    Me.picPlayerCard5.Image = card10
                                    playerCard(4) = 10
                                    card(9) = 10
                                Case 11
                                    Me.picPlayerCard5.Image = card11
                                    playerCard(4) = 11
                                    card(9) = 11
                                Case 12
                                    Me.picPlayerCard5.Image = card12
                                    playerCard(4) = 12
                                    card(9) = 12
                                Case 13
                                    Me.picPlayerCard5.Image = card13
                                    playerCard(4) = 13
                                    card(9) = 13
                                Case 14
                                    Me.picPlayerCard5.Image = card14
                                    playerCard(4) = 1
                                    card(9) = 14
                                Case 15
                                    Me.picPlayerCard5.Image = card15
                                    playerCard(4) = 2
                                    card(9) = 15
                                Case 16
                                    Me.picPlayerCard5.Image = card16
                                    playerCard(4) = 3
                                    card(9) = 16
                                Case 17
                                    Me.picPlayerCard5.Image = card17
                                    playerCard(4) = 4
                                    card(9) = 17
                                Case 18
                                    Me.picPlayerCard5.Image = card18
                                    playerCard(4) = 5
                                    card(9) = 18
                                Case 19
                                    Me.picPlayerCard5.Image = card19
                                    playerCard(4) = 6
                                    card(9) = 19
                                Case 20
                                    Me.picPlayerCard5.Image = card20
                                    playerCard(4) = 7
                                    card(9) = 20
                                Case 21
                                    Me.picPlayerCard5.Image = card21
                                    playerCard(4) = 8
                                    card(9) = 21
                                Case 22
                                    Me.picPlayerCard5.Image = card22
                                    playerCard(4) = 9
                                    card(9) = 22
                                Case 23
                                    Me.picPlayerCard5.Image = card23
                                    playerCard(4) = 10
                                    card(9) = 23
                                Case 24
                                    Me.picPlayerCard5.Image = card24
                                    playerCard(4) = 11
                                    card(9) = 24
                                Case 25
                                    Me.picPlayerCard5.Image = card25
                                    playerCard(4) = 12
                                    card(9) = 25
                                Case 26
                                    Me.picPlayerCard5.Image = card26
                                    playerCard(4) = 13
                                    card(9) = 26
                                Case 27
                                    Me.picPlayerCard5.Image = card27
                                    playerCard(4) = 1
                                    card(9) = 27
                                Case 28
                                    Me.picPlayerCard5.Image = card28
                                    playerCard(4) = 2
                                    card(9) = 28
                                Case 29
                                    Me.picPlayerCard5.Image = card29
                                    playerCard(4) = 3
                                    card(9) = 29
                                Case 30
                                    Me.picPlayerCard5.Image = card30
                                    playerCard(4) = 4
                                    card(9) = 30
                                Case 31
                                    Me.picPlayerCard5.Image = card31
                                    playerCard(4) = 5
                                    card(9) = 31
                                Case 32
                                    Me.picPlayerCard5.Image = card32
                                    playerCard(4) = 6
                                    card(9) = 32
                                Case 33
                                    Me.picPlayerCard5.Image = card33
                                    playerCard(4) = 7
                                    card(9) = 33
                                Case 34
                                    Me.picPlayerCard5.Image = card34
                                    playerCard(4) = 8
                                    card(9) = 34
                                Case 35
                                    Me.picPlayerCard5.Image = card35
                                    playerCard(4) = 9
                                    card(9) = 35
                                Case 36
                                    Me.picPlayerCard5.Image = card36
                                    playerCard(4) = 10
                                    card(9) = 36
                                Case 37
                                    Me.picPlayerCard5.Image = card37
                                    playerCard(4) = 11
                                    card(9) = 37
                                Case 38
                                    Me.picPlayerCard5.Image = card38
                                    playerCard(4) = 12
                                    card(9) = 38
                                Case 39
                                    Me.picPlayerCard5.Image = card39
                                    playerCard(4) = 13
                                    card(9) = 39
                                Case 40
                                    Me.picPlayerCard5.Image = card40
                                    playerCard(4) = 1
                                    card(9) = 40
                                Case 41
                                    Me.picPlayerCard5.Image = card41
                                    playerCard(4) = 2
                                    card(9) = 41
                                Case 42
                                    Me.picPlayerCard5.Image = card42
                                    playerCard(4) = 3
                                    card(9) = 42
                                Case 43
                                    Me.picPlayerCard5.Image = card43
                                    playerCard(4) = 4
                                    card(9) = 43
                                Case 44
                                    Me.picPlayerCard5.Image = card44
                                    playerCard(4) = 5
                                    card(9) = 44
                                Case 45
                                    Me.picPlayerCard5.Image = card45
                                    playerCard(4) = 6
                                    card(9) = 45
                                Case 46
                                    Me.picPlayerCard5.Image = card46
                                    playerCard(4) = 7
                                    card(9) = 46
                                Case 47
                                    Me.picPlayerCard5.Image = card47
                                    playerCard(4) = 8
                                    card(9) = 47
                                Case 48
                                    Me.picPlayerCard5.Image = card48
                                    playerCard(4) = 9
                                    card(9) = 48
                                Case 49
                                    Me.picPlayerCard5.Image = card49
                                    playerCard(4) = 10
                                    card(9) = 49
                                Case 50
                                    Me.picPlayerCard5.Image = card50
                                    playerCard(4) = 11
                                    card(9) = 50
                                Case 51
                                    Me.picPlayerCard5.Image = card51
                                    playerCard(4) = 12
                                    card(9) = 51
                                Case 52
                                    Me.picPlayerCard5.Image = card52
                                    playerCard(4) = 13
                                    card(9) = 52
                            End Select
                        End If

                        Used.Add(newNum)

                    End If
                Loop While OK = False
            End If
        End If

        playerCardsLeftTotal = playerCardsLeft1 + playerCardsLeft2 + playerCardsLeft3 + playerCardsLeft4 + playerCardsLeft5
        Me.lblPlayerCardsLeftTotal.Text = playerName & "'s" & vbCrLf & "Total Cards Left:" & vbCrLf & playerCardsLeftTotal
        Me.lblPlayerCardsLeft1.Text = "Cards Left:" & playerCardsLeft1
        Me.lblPlayerCardsLeft2.Text = "Cards Left:" & playerCardsLeft2
        Me.lblPlayerCardsLeft3.Text = "Cards Left:" & playerCardsLeft3
        Me.lblPlayerCardsLeft4.Text = "Cards Left:" & playerCardsLeft4
        Me.lblPlayerCardsLeft5.Text = "Cards Left:" & playerCardsLeft5

    End Sub

    Private Sub NewGameToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NewGameToolStripMenuItem.Click
        tmrGame.Stop()                       'pause game
        tmrEasy.Stop()
        tmrMedium.Stop()
        tmrHard.Stop()

        frmNewGameOptions.Show()            'show options form
        frmNewGameOptions.BringToFront()

        If btnCancel_Click = True Then      'if player clicks cancel, 
            frmNewGameOptions.Hide()        'hide options form and 
            tmrGame.Start()                 'resume previous game

            If compEasy = True Then
                tmrEasy.Start()
                tmrMedium.Stop()
                tmrHard.Stop()
            ElseIf compMedium = True Then
                tmrEasy.Stop()
                tmrMedium.Start()
                tmrHard.Stop()
            ElseIf compHard = True Then
                tmrEasy.Stop()
                tmrMedium.Stop()
                tmrHard.Start()
            End If
        End If

    End Sub

    Private Sub HowToPlayToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles HowToPlayToolStripMenuItem.Click
        tmrGame.Stop()              'pause game
        tmrEasy.Stop()
        tmrMedium.Stop()
        tmrHard.Stop()

        'show messagebox with rules/how to play
        MessageBox.Show("SPEED" & vbCrLf & vbCrLf & howToPlay, "How to Play", MessageBoxButtons.OK)

        If DialogResult.OK Then                 'once player clicks okay, game resumes
            tmrGame.Start()

            If compEasy = True Then
                tmrEasy.Start()
                tmrMedium.Stop()
                tmrHard.Stop()
            ElseIf compMedium = True Then
                tmrEasy.Stop()
                tmrMedium.Start()
                tmrHard.Stop()
            ElseIf compHard = True Then
                tmrEasy.Stop()
                tmrMedium.Stop()
                tmrHard.Start()
            End If

        End If

    End Sub

    Private Sub QuitToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles QuitToolStripMenuItem.Click
        tmrGame.Stop()              'end game
        tmrEasy.Stop()
        tmrMedium.Stop()
        tmrHard.Stop()

        frmNewGameOptions.txtName.Text = Nothing        'set options to default
        frmNewGameOptions.radEasy.Checked = True
        frmNewGameOptions.radMedium.Checked = False
        frmNewGameOptions.radHard.Checked = False
        frmNewGameOptions.radBlue.Checked = True
        frmNewGameOptions.radRed.Checked = False
        frmNewGameOptions.radBlack.Checked = False

        Me.Close()                      'close form and return to main menu
        frmSpeedHome.Activate()
        frmSpeedHome.Show()
    End Sub

    Private Sub btnShuffle_Click(sender As Object, e As EventArgs) Handles btnShuffle.Click
        Dim OK As Boolean = False
        Dim newNum As Integer

        For pile As Integer = 10 To 11
            For cardNum As Integer = 0 To 9
                If card(pile) <> (card(cardNum) + 1 Or card(cardNum) - 1) Then
                    For counter As Integer = 1 To 2
                        Do
                            Randomize()
                            newNum = Int(max * Rnd() + min)

                            If Used.Contains(newNum) Then
                                OK = False
                            Else
                                OK = True

                                Select Case counter
                                    Case 1
                                        Select Case newNum
                                            Case 1
                                                Me.picCompPile.Image = card1
                                                compPile = 1
                                                card(10) = 1
                                            Case 2
                                                Me.picCompPile.Image = card2
                                                compPile = 2
                                                card(10) = 2
                                            Case 3
                                                Me.picCompPile.Image = card3
                                                compPile = 3
                                                card(10) = 3
                                            Case 4
                                                Me.picCompPile.Image = card4
                                                compPile = 4
                                                card(10) = 4
                                            Case 5
                                                Me.picCompPile.Image = card5
                                                compPile = 5
                                                card(10) = 5
                                            Case 6
                                                Me.picCompPile.Image = card6
                                                compPile = 6
                                                card(10) = 6
                                            Case 7
                                                Me.picCompPile.Image = card7
                                                compPile = 7
                                                card(10) = 7
                                            Case 8
                                                Me.picCompPile.Image = card8
                                                compPile = 8
                                                card(10) = 8
                                            Case 9
                                                Me.picCompPile.Image = card9
                                                compPile = 9
                                                card(10) = 9
                                            Case 10
                                                Me.picCompPile.Image = card10
                                                compPile = 10
                                                card(10) = 10
                                            Case 11
                                                Me.picCompPile.Image = card11
                                                compPile = 11
                                                card(10) = 11
                                            Case 12
                                                Me.picCompPile.Image = card12
                                                compPile = 12
                                                card(10) = 12
                                            Case 13
                                                Me.picCompPile.Image = card13
                                                compPile = 13
                                                card(10) = 13
                                            Case 14
                                                Me.picCompPile.Image = card14
                                                compPile = 1
                                                card(10) = 14
                                            Case 15
                                                Me.picCompPile.Image = card15
                                                compPile = 2
                                                card(10) = 15
                                            Case 16
                                                Me.picCompPile.Image = card16
                                                compPile = 3
                                                card(10) = 16
                                            Case 17
                                                Me.picCompPile.Image = card17
                                                compPile = 4
                                                card(10) = 17
                                            Case 18
                                                Me.picCompPile.Image = card18
                                                compPile = 5
                                                card(10) = 18
                                            Case 19
                                                Me.picCompPile.Image = card19
                                                compPile = 6
                                                card(10) = 19
                                            Case 20
                                                Me.picCompPile.Image = card20
                                                compPile = 7
                                                card(10) = 20
                                            Case 21
                                                Me.picCompPile.Image = card21
                                                compPile = 8
                                                card(10) = 21
                                            Case 22
                                                Me.picCompPile.Image = card22
                                                compPile = 9
                                                card(10) = 22
                                            Case 23
                                                Me.picCompPile.Image = card23
                                                compPile = 10
                                                card(10) = 23
                                            Case 24
                                                Me.picCompPile.Image = card24
                                                compPile = 11
                                                card(10) = 24
                                            Case 25
                                                Me.picCompPile.Image = card25
                                                compPile = 12
                                                card(10) = 25
                                            Case 26
                                                Me.picCompPile.Image = card26
                                                compPile = 13
                                                card(10) = 26
                                            Case 27
                                                Me.picCompPile.Image = card27
                                                compPile = 1
                                                card(10) = 27
                                            Case 28
                                                Me.picCompPile.Image = card28
                                                compPile = 2
                                                card(10) = 28
                                            Case 29
                                                Me.picCompPile.Image = card29
                                                compPile = 3
                                                card(10) = 29
                                            Case 30
                                                Me.picCompPile.Image = card30
                                                compPile = 4
                                                card(10) = 30
                                            Case 31
                                                Me.picCompPile.Image = card31
                                                compPile = 5
                                                card(10) = 31
                                            Case 32
                                                Me.picCompPile.Image = card32
                                                compPile = 6
                                                card(10) = 32
                                            Case 33
                                                Me.picCompPile.Image = card33
                                                compPile = 7
                                                card(10) = 33
                                            Case 34
                                                Me.picCompPile.Image = card34
                                                compPile = 8
                                                card(10) = 34
                                            Case 35
                                                Me.picCompPile.Image = card35
                                                compPile = 9
                                                card(10) = 35
                                            Case 36
                                                Me.picCompPile.Image = card36
                                                compPile = 10
                                                card(10) = 36
                                            Case 37
                                                Me.picCompPile.Image = card37
                                                compPile = 11
                                                card(10) = 37
                                            Case 38
                                                Me.picCompPile.Image = card38
                                                compPile = 12
                                                card(10) = 38
                                            Case 39
                                                Me.picCompPile.Image = card39
                                                compPile = 13
                                                card(10) = 39
                                            Case 40
                                                Me.picCompPile.Image = card40
                                                compPile = 1
                                                card(10) = 40
                                            Case 41
                                                Me.picCompPile.Image = card41
                                                compPile = 2
                                                card(10) = 41
                                            Case 42
                                                Me.picCompPile.Image = card42
                                                compPile = 3
                                                card(10) = 42
                                            Case 43
                                                Me.picCompPile.Image = card43
                                                compPile = 4
                                                card(10) = 43
                                            Case 44
                                                Me.picCompPile.Image = card44
                                                compPile = 5
                                                card(10) = 44
                                            Case 45
                                                Me.picCompPile.Image = card45
                                                compPile = 6
                                                card(10) = 45
                                            Case 46
                                                Me.picCompPile.Image = card46
                                                compPile = 7
                                                card(10) = 46
                                            Case 47
                                                Me.picCompPile.Image = card47
                                                compPile = 8
                                                card(10) = 47
                                            Case 48
                                                Me.picCompPile.Image = card48
                                                compPile = 9
                                                card(10) = 48
                                            Case 49
                                                Me.picCompPile.Image = card49
                                                compPile = 10
                                                card(10) = 49
                                            Case 50
                                                Me.picCompPile.Image = card50
                                                compPile = 11
                                                card(10) = 50
                                            Case 51
                                                Me.picCompPile.Image = card51
                                                compPile = 12
                                                card(10) = 51
                                            Case 52
                                                Me.picCompPile.Image = card52
                                                compPile = 13
                                                card(10) = 52
                                        End Select
                                    Case 2
                                        Select Case newNum
                                            Case 1
                                                Me.picPlayerPile.Image = card1
                                                playerPile = 1
                                                card(11) = 1
                                            Case 2
                                                Me.picPlayerPile.Image = card2
                                                playerPile = 2
                                                card(11) = 2
                                            Case 3
                                                Me.picPlayerPile.Image = card3
                                                playerPile = 3
                                                card(11) = 3
                                            Case 4
                                                Me.picPlayerPile.Image = card4
                                                playerPile = 4
                                                card(11) = 4
                                            Case 5
                                                Me.picPlayerPile.Image = card5
                                                playerPile = 5
                                                card(11) = 5
                                            Case 6
                                                Me.picPlayerPile.Image = card6
                                                playerPile = 6
                                                card(11) = 6
                                            Case 7
                                                Me.picPlayerPile.Image = card7
                                                playerPile = 7
                                                card(11) = 7
                                            Case 8
                                                Me.picPlayerPile.Image = card8
                                                playerPile = 8
                                                card(11) = 8
                                            Case 9
                                                Me.picPlayerPile.Image = card9
                                                playerPile = 9
                                                card(11) = 9
                                            Case 10
                                                Me.picPlayerPile.Image = card10
                                                playerPile = 10
                                                card(11) = 10
                                            Case 11
                                                Me.picPlayerPile.Image = card11
                                                playerPile = 11
                                                card(11) = 11
                                            Case 12
                                                Me.picPlayerPile.Image = card12
                                                playerPile = 12
                                                card(11) = 12
                                            Case 13
                                                Me.picPlayerPile.Image = card13
                                                playerPile = 13
                                                card(11) = 13
                                            Case 14
                                                Me.picPlayerPile.Image = card14
                                                playerPile = 1
                                                card(11) = 14
                                            Case 15
                                                Me.picPlayerPile.Image = card15
                                                playerPile = 2
                                                card(11) = 15
                                            Case 16
                                                Me.picPlayerPile.Image = card16
                                                playerPile = 3
                                                card(11) = 16
                                            Case 17
                                                Me.picPlayerPile.Image = card17
                                                playerPile = 4
                                                card(11) = 17
                                            Case 18
                                                Me.picPlayerPile.Image = card18
                                                playerPile = 5
                                                card(11) = 18
                                            Case 19
                                                Me.picPlayerPile.Image = card19
                                                playerPile = 6
                                                card(11) = 19
                                            Case 20
                                                Me.picPlayerPile.Image = card20
                                                playerPile = 7
                                                card(11) = 20
                                            Case 21
                                                Me.picPlayerPile.Image = card21
                                                playerPile = 8
                                                card(11) = 21
                                            Case 22
                                                Me.picPlayerPile.Image = card22
                                                playerPile = 9
                                                card(11) = 22
                                            Case 23
                                                Me.picPlayerPile.Image = card23
                                                playerPile = 10
                                                card(11) = 23
                                            Case 24
                                                Me.picPlayerPile.Image = card24
                                                playerPile = 11
                                                card(11) = 24
                                            Case 25
                                                Me.picPlayerPile.Image = card25
                                                playerPile = 12
                                                card(11) = 25
                                            Case 26
                                                Me.picPlayerPile.Image = card26
                                                playerPile = 13
                                                card(11) = 26
                                            Case 27
                                                Me.picPlayerPile.Image = card27
                                                playerPile = 1
                                                card(11) = 27
                                            Case 28
                                                Me.picPlayerPile.Image = card28
                                                playerPile = 2
                                                card(11) = 28
                                            Case 29
                                                Me.picPlayerPile.Image = card29
                                                playerPile = 3
                                                card(11) = 29
                                            Case 30
                                                Me.picPlayerPile.Image = card30
                                                playerPile = 4
                                                card(11) = 30
                                            Case 31
                                                Me.picPlayerPile.Image = card31
                                                playerPile = 5
                                                card(11) = 31
                                            Case 32
                                                Me.picPlayerPile.Image = card32
                                                playerPile = 6
                                                card(11) = 32
                                            Case 33
                                                Me.picPlayerPile.Image = card33
                                                playerPile = 7
                                                card(11) = 33
                                            Case 34
                                                Me.picPlayerPile.Image = card34
                                                playerPile = 8
                                                card(11) = 34
                                            Case 35
                                                Me.picPlayerPile.Image = card35
                                                playerPile = 9
                                                card(11) = 35
                                            Case 36
                                                Me.picPlayerPile.Image = card36
                                                playerPile = 10
                                                card(11) = 36
                                            Case 37
                                                Me.picPlayerPile.Image = card37
                                                playerPile = 11
                                                card(11) = 37
                                            Case 38
                                                Me.picPlayerPile.Image = card38
                                                playerPile = 12
                                                card(11) = 38
                                            Case 39
                                                Me.picPlayerPile.Image = card39
                                                playerPile = 13
                                                card(11) = 39
                                            Case 40
                                                Me.picPlayerPile.Image = card40
                                                playerPile = 1
                                                card(11) = 40
                                            Case 41
                                                Me.picPlayerPile.Image = card41
                                                playerPile = 2
                                                card(11) = 41
                                            Case 42
                                                Me.picPlayerPile.Image = card42
                                                playerPile = 3
                                                card(11) = 42
                                            Case 43
                                                Me.picPlayerPile.Image = card43
                                                playerPile = 4
                                                card(11) = 43
                                            Case 44
                                                Me.picPlayerPile.Image = card44
                                                playerPile = 5
                                                card(11) = 44
                                            Case 45
                                                Me.picPlayerPile.Image = card45
                                                playerPile = 6
                                                card(11) = 45
                                            Case 46
                                                Me.picPlayerPile.Image = card46
                                                playerPile = 7
                                                card(11) = 46
                                            Case 47
                                                Me.picPlayerPile.Image = card47
                                                playerPile = 8
                                                card(11) = 47
                                            Case 48
                                                Me.picPlayerPile.Image = card48
                                                playerPile = 9
                                                card(11) = 48
                                            Case 49
                                                Me.picPlayerPile.Image = card49
                                                playerPile = 10
                                                card(11) = 49
                                            Case 50
                                                Me.picPlayerPile.Image = card50
                                                playerPile = 11
                                                card(11) = 50
                                            Case 51
                                                Me.picPlayerPile.Image = card51
                                                playerPile = 12
                                                card(11) = 51
                                            Case 52
                                                Me.picPlayerPile.Image = card52
                                                playerPile = 13
                                                card(11) = 52
                                        End Select
                                End Select
                            End If
                        Loop While OK = False
                    Next counter
                End If
            Next cardNum
        Next pile

    End Sub

End Class
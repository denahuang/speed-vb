Public Class frmNewGameOptions

    Private Sub frmNewGameOptions_Load(sender As Object, e As EventArgs) Handles Me.Load
        'clear and set options to default
        txtName.Text = Nothing
        radEasy.Checked = True
        radMedium.Checked = False
        radHard.Checked = False
        radBlue.Checked = True
        radRed.Checked = False
        radBlack.Checked = False
    End Sub

    Private Sub txtName_KeyDown(sender As Object, e As KeyEventArgs) Handles txtName.KeyDown
        'set focus to start button when player presses enter key after entering their name
        If e.KeyCode = Keys.Return Then
            btnStart.Focus()
        End If
    End Sub

    Private Sub btnStart_Click(sender As Object, e As EventArgs) Handles btnStart.Click
        frmSpeedGame.Close()                                'reset game form by closing

        playerName = Me.txtName.Text                        'set player's name based on what they entered

        If radEasy.Checked = True Then                      'set computer difficulty
            compEasy = True
            compMedium = False
            compHard = False
        ElseIf radMedium.Checked = True Then
            compEasy = False
            compMedium = True
            compHard = False
        ElseIf radHard.Checked = True Then
            compEasy = False
            compMedium = False
            compHard = True
        End If

        compCardsLeftTotal = 15             'give the value of how many cards the player and computer have
        compCardsLeft1 = 1
        compCardsLeft2 = 2
        compCardsLeft3 = 3
        compCardsLeft4 = 4
        compCardsLeft5 = 5
        playerCardsLeftTotal = 15
        playerCardsLeft1 = 1
        playerCardsLeft2 = 2
        playerCardsLeft3 = 3
        playerCardsLeft4 = 4
        playerCardsLeft5 = 5

        If radBlue.Checked = True Then                                          'set card back color based on player preference
            frmSpeedGame.picCompPile.Image = My.Resources.CardBackBlue
            frmSpeedGame.picCompDeck.Image = My.Resources.CardBackBlue
            frmSpeedGame.picCompCard1.Image = My.Resources.CardBackBlue
            frmSpeedGame.picCompCard2.Image = My.Resources.CardBackBlue
            frmSpeedGame.picCompCard3.Image = My.Resources.CardBackBlue
            frmSpeedGame.picCompCard4.Image = My.Resources.CardBackBlue
            frmSpeedGame.picCompCard5.Image = My.Resources.CardBackBlue

            frmSpeedGame.picPlayerPile.Image = My.Resources.CardBackBlue
            frmSpeedGame.picPlayerDeck.Image = My.Resources.CardBackBlue
            frmSpeedGame.picPlayerCard1.Image = My.Resources.CardBackBlue
            frmSpeedGame.picPlayerCard2.Image = My.Resources.CardBackBlue
            frmSpeedGame.picPlayerCard3.Image = My.Resources.CardBackBlue
            frmSpeedGame.picPlayerCard4.Image = My.Resources.CardBackBlue
            frmSpeedGame.picPlayerCard5.Image = My.Resources.CardBackBlue
        ElseIf radRed.Checked = True Then
            frmSpeedGame.picCompPile.Image = My.Resources.CardBackRed
            frmSpeedGame.picCompDeck.Image = My.Resources.CardBackRed
            frmSpeedGame.picCompCard1.Image = My.Resources.CardBackRed
            frmSpeedGame.picCompCard2.Image = My.Resources.CardBackRed
            frmSpeedGame.picCompCard3.Image = My.Resources.CardBackRed
            frmSpeedGame.picCompCard4.Image = My.Resources.CardBackRed
            frmSpeedGame.picCompCard5.Image = My.Resources.CardBackRed

            frmSpeedGame.picPlayerPile.Image = My.Resources.CardBackRed
            frmSpeedGame.picPlayerDeck.Image = My.Resources.CardBackRed
            frmSpeedGame.picPlayerCard1.Image = My.Resources.CardBackRed
            frmSpeedGame.picPlayerCard2.Image = My.Resources.CardBackRed
            frmSpeedGame.picPlayerCard3.Image = My.Resources.CardBackRed
            frmSpeedGame.picPlayerCard4.Image = My.Resources.CardBackRed
            frmSpeedGame.picPlayerCard5.Image = My.Resources.CardBackRed
        ElseIf radBlack.Checked = True Then
            frmSpeedGame.picCompPile.Image = My.Resources.CardBackBlack
            frmSpeedGame.picCompDeck.Image = My.Resources.CardBackBlack
            frmSpeedGame.picCompCard1.Image = My.Resources.CardBackBlack
            frmSpeedGame.picCompCard2.Image = My.Resources.CardBackBlack
            frmSpeedGame.picCompCard3.Image = My.Resources.CardBackBlack
            frmSpeedGame.picCompCard4.Image = My.Resources.CardBackBlack
            frmSpeedGame.picCompCard5.Image = My.Resources.CardBackBlack

            frmSpeedGame.picPlayerPile.Image = My.Resources.CardBackBlack
            frmSpeedGame.picPlayerDeck.Image = My.Resources.CardBackBlack
            frmSpeedGame.picPlayerCard1.Image = My.Resources.CardBackBlack
            frmSpeedGame.picPlayerCard2.Image = My.Resources.CardBackBlack
            frmSpeedGame.picPlayerCard3.Image = My.Resources.CardBackBlack
            frmSpeedGame.picPlayerCard4.Image = My.Resources.CardBackBlack
            frmSpeedGame.picPlayerCard5.Image = My.Resources.CardBackBlack
        End If

        'display card labels and how many cards left
        frmSpeedGame.lblCompCardsLeftTotal.Text = "Computer's" & vbCrLf & "Total Cards Left:" & vbCrLf & compCardsLeftTotal
        frmSpeedGame.lblCompCardsLeft1.Text = "Cards Left: " & compCardsLeft1
        frmSpeedGame.lblCompCardsLeft2.Text = "Cards Left: " & compCardsLeft2
        frmSpeedGame.lblCompCardsLeft3.Text = "Cards Left: " & compCardsLeft3
        frmSpeedGame.lblCompCardsLeft4.Text = "Cards Left: " & compCardsLeft4
        frmSpeedGame.lblCompCardsLeft5.Text = "Cards Left: " & compCardsLeft5

        frmSpeedGame.lblPlayerCardsLeftTotal.Text = playerName & "'s" & vbCrLf & "Total Cards Left:" & vbCrLf & playerCardsLeftTotal
        frmSpeedGame.lblPlayerCardsLeft1.Text = "Cards Left:" & playerCardsLeft1
        frmSpeedGame.lblPlayerCardsLeft2.Text = "Cards Left:" & playerCardsLeft2
        frmSpeedGame.lblPlayerCardsLeft3.Text = "Cards Left:" & playerCardsLeft3
        frmSpeedGame.lblPlayerCardsLeft4.Text = "Cards Left:" & playerCardsLeft4
        frmSpeedGame.lblPlayerCardsLeft5.Text = "Cards Left:" & playerCardsLeft5

        'clear list of cards already generated
        Used.Clear()
        Used.Add(0)

        For counter As Integer = 1 To 12                    'generate 12 different cards to display for game
            Dim OK As Boolean = False
            Dim num As Integer

            Do                                              'generate random unique numbers
                Randomize()
                num = Int(max * Rnd() + min)

                If Used.Contains(num) Then                  'if card has already been generated, generate a new card
                    OK = False

                Else                                        'if card has not already been generated, allow card to be displayed
                    OK = True

                    Select Case counter
                        Case 1
                            Select Case num
                                Case 1
                                    frmSpeedGame.picCompCard1.Image = card1             'set image
                                    compCard(0) = 1                                     'set card [1 to 13: A to K]
                                    card(0) = 1                                         'set card [1 to 52: A of Spades to K of Diamonds]
                                Case 2
                                    frmSpeedGame.picCompCard1.Image = card2
                                    compCard(0) = 2
                                    card(0) = 2
                                Case 3
                                    frmSpeedGame.picCompCard1.Image = card3
                                    compCard(0) = 3
                                    card(0) = 3
                                Case 4
                                    frmSpeedGame.picCompCard1.Image = card4
                                    compCard(0) = 4
                                    card(0) = 4
                                Case 5
                                    frmSpeedGame.picCompCard1.Image = card5
                                    compCard(0) = 5
                                    card(0) = 5
                                Case 6
                                    frmSpeedGame.picCompCard1.Image = card6
                                    compCard(0) = 6
                                    card(0) = 6
                                Case 7
                                    frmSpeedGame.picCompCard1.Image = card7
                                    compCard(0) = 7
                                    card(0) = 7
                                Case 8
                                    frmSpeedGame.picCompCard1.Image = card8
                                    compCard(0) = 8
                                    card(0) = 8
                                Case 9
                                    frmSpeedGame.picCompCard1.Image = card9
                                    compCard(0) = 9
                                    card(0) = 9
                                Case 10
                                    frmSpeedGame.picCompCard1.Image = card10
                                    compCard(0) = 10
                                    card(0) = 10
                                Case 11
                                    frmSpeedGame.picCompCard1.Image = card11
                                    compCard(0) = 11
                                    card(0) = 11
                                Case 12
                                    frmSpeedGame.picCompCard1.Image = card12
                                    compCard(0) = 12
                                    card(0) = 12
                                Case 13
                                    frmSpeedGame.picCompCard1.Image = card13
                                    compCard(0) = 13
                                    card(0) = 13
                                Case 14
                                    frmSpeedGame.picCompCard1.Image = card14
                                    compCard(0) = 1
                                    card(0) = 14
                                Case 15
                                    frmSpeedGame.picCompCard1.Image = card15
                                    compCard(0) = 2
                                    card(0) = 15
                                Case 16
                                    frmSpeedGame.picCompCard1.Image = card16
                                    compCard(0) = 3
                                    card(0) = 16
                                Case 17
                                    frmSpeedGame.picCompCard1.Image = card17
                                    compCard(0) = 4
                                    card(0) = 17
                                Case 18
                                    frmSpeedGame.picCompCard1.Image = card18
                                    compCard(0) = 5
                                    card(0) = 18
                                Case 19
                                    frmSpeedGame.picCompCard1.Image = card19
                                    compCard(0) = 6
                                    card(0) = 19
                                Case 20
                                    frmSpeedGame.picCompCard1.Image = card20
                                    compCard(0) = 7
                                    card(0) = 20
                                Case 21
                                    frmSpeedGame.picCompCard1.Image = card21
                                    compCard(0) = 8
                                    card(0) = 21
                                Case 22
                                    frmSpeedGame.picCompCard1.Image = card22
                                    compCard(0) = 9
                                    card(0) = 22
                                Case 23
                                    frmSpeedGame.picCompCard1.Image = card23
                                    compCard(0) = 10
                                    card(0) = 23
                                Case 24
                                    frmSpeedGame.picCompCard1.Image = card24
                                    compCard(0) = 11
                                    card(0) = 24
                                Case 25
                                    frmSpeedGame.picCompCard1.Image = card25
                                    compCard(0) = 12
                                    card(0) = 25
                                Case 26
                                    frmSpeedGame.picCompCard1.Image = card26
                                    compCard(0) = 13
                                    card(0) = 26
                                Case 27
                                    frmSpeedGame.picCompCard1.Image = card27
                                    compCard(0) = 1
                                    card(0) = 27
                                Case 28
                                    frmSpeedGame.picCompCard1.Image = card28
                                    compCard(0) = 2
                                    card(0) = 28
                                Case 29
                                    frmSpeedGame.picCompCard1.Image = card29
                                    compCard(0) = 3
                                    card(0) = 29
                                Case 30
                                    frmSpeedGame.picCompCard1.Image = card30
                                    compCard(0) = 4
                                    card(0) = 30
                                Case 31
                                    frmSpeedGame.picCompCard1.Image = card31
                                    compCard(0) = 5
                                    card(0) = 31
                                Case 32
                                    frmSpeedGame.picCompCard1.Image = card32
                                    compCard(0) = 6
                                    card(0) = 32
                                Case 33
                                    frmSpeedGame.picCompCard1.Image = card33
                                    compCard(0) = 7
                                    card(0) = 33
                                Case 34
                                    frmSpeedGame.picCompCard1.Image = card34
                                    compCard(0) = 8
                                    card(0) = 34
                                Case 35
                                    frmSpeedGame.picCompCard1.Image = card35
                                    compCard(0) = 9
                                    card(0) = 35
                                Case 36
                                    frmSpeedGame.picCompCard1.Image = card36
                                    compCard(0) = 10
                                    card(0) = 36
                                Case 37
                                    frmSpeedGame.picCompCard1.Image = card37
                                    compCard(0) = 11
                                    card(0) = 37
                                Case 38
                                    frmSpeedGame.picCompCard1.Image = card38
                                    compCard(0) = 12
                                    card(0) = 38
                                Case 39
                                    frmSpeedGame.picCompCard1.Image = card39
                                    compCard(0) = 13
                                    card(0) = 39
                                Case 40
                                    frmSpeedGame.picCompCard1.Image = card40
                                    compCard(0) = 1
                                    card(0) = 40
                                Case 41
                                    frmSpeedGame.picCompCard1.Image = card41
                                    compCard(0) = 2
                                    card(0) = 41
                                Case 42
                                    frmSpeedGame.picCompCard1.Image = card42
                                    compCard(0) = 3
                                    card(0) = 42
                                Case 43
                                    frmSpeedGame.picCompCard1.Image = card43
                                    compCard(0) = 4
                                    card(0) = 43
                                Case 44
                                    frmSpeedGame.picCompCard1.Image = card44
                                    compCard(0) = 5
                                    card(0) = 44
                                Case 45
                                    frmSpeedGame.picCompCard1.Image = card45
                                    compCard(0) = 6
                                    card(0) = 45
                                Case 46
                                    frmSpeedGame.picCompCard1.Image = card46
                                    compCard(0) = 7
                                    card(0) = 46
                                Case 47
                                    frmSpeedGame.picCompCard1.Image = card47
                                    compCard(0) = 8
                                    card(0) = 47
                                Case 48
                                    frmSpeedGame.picCompCard1.Image = card48
                                    compCard(0) = 9
                                    card(0) = 48
                                Case 49
                                    frmSpeedGame.picCompCard1.Image = card49
                                    compCard(0) = 10
                                    card(0) = 49
                                Case 50
                                    frmSpeedGame.picCompCard1.Image = card50
                                    compCard(0) = 11
                                    card(0) = 50
                                Case 51
                                    frmSpeedGame.picCompCard1.Image = card51
                                    compCard(0) = 12
                                    card(0) = 51
                                Case 52
                                    frmSpeedGame.picCompCard1.Image = card52
                                    compCard(0) = 13
                                    card(0) = 52
                            End Select
                        Case 2
                            Select Case num
                                Case 1
                                    frmSpeedGame.picCompCard2.Image = card1
                                    compCard(1) = 1
                                    card(1) = 1
                                Case 2
                                    frmSpeedGame.picCompCard2.Image = card2
                                    compCard(1) = 2
                                    card(1) = 2
                                Case 3
                                    frmSpeedGame.picCompCard2.Image = card3
                                    compCard(1) = 3
                                    card(1) = 3
                                Case 4
                                    frmSpeedGame.picCompCard2.Image = card4
                                    compCard(1) = 4
                                    card(1) = 4
                                Case 5
                                    frmSpeedGame.picCompCard2.Image = card5
                                    compCard(1) = 5
                                    card(1) = 5
                                Case 6
                                    frmSpeedGame.picCompCard2.Image = card6
                                    compCard(1) = 6
                                    card(1) = 6
                                Case 7
                                    frmSpeedGame.picCompCard2.Image = card7
                                    compCard(1) = 7
                                    card(1) = 7
                                Case 8
                                    frmSpeedGame.picCompCard2.Image = card8
                                    compCard(1) = 8
                                    card(1) = 8
                                Case 9
                                    frmSpeedGame.picCompCard2.Image = card9
                                    compCard(1) = 9
                                    card(1) = 9
                                Case 10
                                    frmSpeedGame.picCompCard2.Image = card10
                                    compCard(1) = 10
                                    card(1) = 10
                                Case 11
                                    frmSpeedGame.picCompCard2.Image = card11
                                    compCard(1) = 11
                                    card(1) = 11
                                Case 12
                                    frmSpeedGame.picCompCard2.Image = card12
                                    compCard(1) = 12
                                    card(1) = 12
                                Case 13
                                    frmSpeedGame.picCompCard2.Image = card13
                                    compCard(1) = 13
                                    card(1) = 13
                                Case 14
                                    frmSpeedGame.picCompCard2.Image = card14
                                    compCard(1) = 1
                                    card(1) = 14
                                Case 15
                                    frmSpeedGame.picCompCard2.Image = card15
                                    compCard(1) = 2
                                    card(1) = 15
                                Case 16
                                    frmSpeedGame.picCompCard2.Image = card16
                                    compCard(1) = 3
                                    card(1) = 16
                                Case 17
                                    frmSpeedGame.picCompCard2.Image = card17
                                    compCard(1) = 4
                                    card(1) = 17
                                Case 18
                                    frmSpeedGame.picCompCard2.Image = card18
                                    compCard(1) = 5
                                    card(1) = 18
                                Case 19
                                    frmSpeedGame.picCompCard2.Image = card19
                                    compCard(1) = 6
                                    card(1) = 19
                                Case 20
                                    frmSpeedGame.picCompCard2.Image = card20
                                    compCard(1) = 7
                                    card(1) = 20
                                Case 21
                                    frmSpeedGame.picCompCard2.Image = card21
                                    compCard(1) = 8
                                    card(1) = 21
                                Case 22
                                    frmSpeedGame.picCompCard2.Image = card22
                                    compCard(1) = 9
                                    card(1) = 22
                                Case 23
                                    frmSpeedGame.picCompCard2.Image = card23
                                    compCard(1) = 10
                                    card(1) = 23
                                Case 24
                                    frmSpeedGame.picCompCard2.Image = card24
                                    compCard(1) = 11
                                    card(1) = 24
                                Case 25
                                    frmSpeedGame.picCompCard2.Image = card25
                                    compCard(1) = 12
                                    card(1) = 25
                                Case 26
                                    frmSpeedGame.picCompCard2.Image = card26
                                    compCard(1) = 13
                                    card(1) = 26
                                Case 27
                                    frmSpeedGame.picCompCard2.Image = card27
                                    compCard(1) = 1
                                    card(1) = 27
                                Case 28
                                    frmSpeedGame.picCompCard2.Image = card28
                                    compCard(1) = 2
                                    card(1) = 28
                                Case 29
                                    frmSpeedGame.picCompCard2.Image = card29
                                    compCard(1) = 3
                                    card(1) = 29
                                Case 30
                                    frmSpeedGame.picCompCard2.Image = card30
                                    compCard(1) = 4
                                    card(1) = 30
                                Case 31
                                    frmSpeedGame.picCompCard2.Image = card31
                                    compCard(1) = 5
                                    card(1) = 31
                                Case 32
                                    frmSpeedGame.picCompCard2.Image = card32
                                    compCard(1) = 6
                                    card(1) = 32
                                Case 33
                                    frmSpeedGame.picCompCard2.Image = card33
                                    compCard(1) = 7
                                    card(1) = 33
                                Case 34
                                    frmSpeedGame.picCompCard2.Image = card34
                                    compCard(1) = 8
                                    card(1) = 34
                                Case 35
                                    frmSpeedGame.picCompCard2.Image = card35
                                    compCard(1) = 9
                                    card(1) = 35
                                Case 36
                                    frmSpeedGame.picCompCard2.Image = card36
                                    compCard(1) = 10
                                    card(1) = 36
                                Case 37
                                    frmSpeedGame.picCompCard2.Image = card37
                                    compCard(1) = 11
                                    card(1) = 37
                                Case 38
                                    frmSpeedGame.picCompCard2.Image = card38
                                    compCard(1) = 12
                                    card(1) = 38
                                Case 39
                                    frmSpeedGame.picCompCard2.Image = card39
                                    compCard(1) = 13
                                    card(1) = 39
                                Case 40
                                    frmSpeedGame.picCompCard2.Image = card40
                                    compCard(1) = 1
                                    card(1) = 40
                                Case 41
                                    frmSpeedGame.picCompCard2.Image = card41
                                    compCard(1) = 2
                                    card(1) = 41
                                Case 42
                                    frmSpeedGame.picCompCard2.Image = card42
                                    compCard(1) = 3
                                    card(1) = 42
                                Case 43
                                    frmSpeedGame.picCompCard2.Image = card43
                                    compCard(1) = 4
                                    card(1) = 43
                                Case 44
                                    frmSpeedGame.picCompCard2.Image = card44
                                    compCard(1) = 5
                                    card(1) = 44
                                Case 45
                                    frmSpeedGame.picCompCard2.Image = card45
                                    compCard(1) = 6
                                    card(1) = 45
                                Case 46
                                    frmSpeedGame.picCompCard2.Image = card46
                                    compCard(1) = 7
                                    card(1) = 46
                                Case 47
                                    frmSpeedGame.picCompCard2.Image = card47
                                    compCard(1) = 8
                                    card(1) = 47
                                Case 48
                                    frmSpeedGame.picCompCard2.Image = card48
                                    compCard(1) = 9
                                    card(1) = 48
                                Case 49
                                    frmSpeedGame.picCompCard2.Image = card49
                                    compCard(1) = 10
                                    card(1) = 49
                                Case 50
                                    frmSpeedGame.picCompCard2.Image = card50
                                    compCard(1) = 11
                                    card(1) = 50
                                Case 51
                                    frmSpeedGame.picCompCard2.Image = card51
                                    compCard(1) = 12
                                    card(1) = 51
                                Case 52
                                    frmSpeedGame.picCompCard2.Image = card52
                                    compCard(1) = 13
                                    card(1) = 52
                            End Select
                        Case 3
                            Select Case num
                                Case 1
                                    frmSpeedGame.picCompCard3.Image = card1
                                    compCard(2) = 1
                                    card(2) = 1
                                Case 2
                                    frmSpeedGame.picCompCard3.Image = card2
                                    compCard(2) = 2
                                    card(2) = 2
                                Case 3
                                    frmSpeedGame.picCompCard3.Image = card3
                                    compCard(2) = 3
                                    card(2) = 3
                                Case 4
                                    frmSpeedGame.picCompCard3.Image = card4
                                    compCard(2) = 4
                                    card(2) = 4
                                Case 5
                                    frmSpeedGame.picCompCard3.Image = card5
                                    compCard(2) = 5
                                    card(2) = 5
                                Case 6
                                    frmSpeedGame.picCompCard3.Image = card6
                                    compCard(2) = 6
                                    card(2) = 6
                                Case 7
                                    frmSpeedGame.picCompCard3.Image = card7
                                    compCard(2) = 7
                                    card(2) = 7
                                Case 8
                                    frmSpeedGame.picCompCard3.Image = card8
                                    compCard(2) = 8
                                    card(2) = 8
                                Case 9
                                    frmSpeedGame.picCompCard3.Image = card9
                                    compCard(2) = 9
                                    card(2) = 9
                                Case 10
                                    frmSpeedGame.picCompCard3.Image = card10
                                    compCard(2) = 10
                                    card(2) = 10
                                Case 11
                                    frmSpeedGame.picCompCard3.Image = card11
                                    compCard(2) = 11
                                    card(2) = 11
                                Case 12
                                    frmSpeedGame.picCompCard3.Image = card12
                                    compCard(2) = 12
                                    card(2) = 12
                                Case 13
                                    frmSpeedGame.picCompCard3.Image = card13
                                    compCard(2) = 13
                                    card(2) = 13
                                Case 14
                                    frmSpeedGame.picCompCard3.Image = card14
                                    compCard(2) = 1
                                    card(2) = 14
                                Case 15
                                    frmSpeedGame.picCompCard3.Image = card15
                                    compCard(2) = 2
                                    card(2) = 15
                                Case 16
                                    frmSpeedGame.picCompCard3.Image = card16
                                    compCard(2) = 3
                                    card(2) = 16
                                Case 17
                                    frmSpeedGame.picCompCard3.Image = card17
                                    compCard(2) = 4
                                    card(2) = 17
                                Case 18
                                    frmSpeedGame.picCompCard3.Image = card18
                                    compCard(2) = 5
                                    card(2) = 18
                                Case 19
                                    frmSpeedGame.picCompCard3.Image = card19
                                    compCard(2) = 6
                                    card(2) = 19
                                Case 20
                                    frmSpeedGame.picCompCard3.Image = card20
                                    compCard(2) = 7
                                    card(2) = 20
                                Case 21
                                    frmSpeedGame.picCompCard3.Image = card21
                                    compCard(2) = 8
                                    card(2) = 21
                                Case 22
                                    frmSpeedGame.picCompCard3.Image = card22
                                    compCard(2) = 9
                                    card(2) = 22
                                Case 23
                                    frmSpeedGame.picCompCard3.Image = card23
                                    compCard(2) = 10
                                    card(2) = 23
                                Case 24
                                    frmSpeedGame.picCompCard3.Image = card24
                                    compCard(2) = 11
                                    card(2) = 24
                                Case 25
                                    frmSpeedGame.picCompCard3.Image = card25
                                    compCard(2) = 12
                                    card(2) = 25
                                Case 26
                                    frmSpeedGame.picCompCard3.Image = card26
                                    compCard(2) = 13
                                    card(2) = 26
                                Case 27
                                    frmSpeedGame.picCompCard3.Image = card27
                                    compCard(2) = 1
                                    card(2) = 27
                                Case 28
                                    frmSpeedGame.picCompCard3.Image = card28
                                    compCard(2) = 2
                                    card(2) = 28
                                Case 29
                                    frmSpeedGame.picCompCard3.Image = card29
                                    compCard(2) = 3
                                    card(2) = 29
                                Case 30
                                    frmSpeedGame.picCompCard3.Image = card30
                                    compCard(2) = 4
                                    card(2) = 30
                                Case 31
                                    frmSpeedGame.picCompCard3.Image = card31
                                    compCard(2) = 5
                                    card(2) = 31
                                Case 32
                                    frmSpeedGame.picCompCard3.Image = card32
                                    compCard(2) = 6
                                    card(2) = 32
                                Case 33
                                    frmSpeedGame.picCompCard3.Image = card33
                                    compCard(2) = 7
                                    card(2) = 33
                                Case 34
                                    frmSpeedGame.picCompCard3.Image = card34
                                    compCard(2) = 8
                                    card(2) = 34
                                Case 35
                                    frmSpeedGame.picCompCard3.Image = card35
                                    compCard(2) = 9
                                    card(2) = 35
                                Case 36
                                    frmSpeedGame.picCompCard3.Image = card36
                                    compCard(2) = 10
                                    card(2) = 36
                                Case 37
                                    frmSpeedGame.picCompCard3.Image = card37
                                    compCard(2) = 11
                                    card(2) = 37
                                Case 38
                                    frmSpeedGame.picCompCard3.Image = card38
                                    compCard(2) = 12
                                    card(2) = 38
                                Case 39
                                    frmSpeedGame.picCompCard3.Image = card39
                                    compCard(2) = 13
                                    card(2) = 39
                                Case 40
                                    frmSpeedGame.picCompCard3.Image = card40
                                    compCard(2) = 1
                                    card(2) = 40
                                Case 41
                                    frmSpeedGame.picCompCard3.Image = card41
                                    compCard(2) = 2
                                    card(2) = 41
                                Case 42
                                    frmSpeedGame.picCompCard3.Image = card42
                                    compCard(2) = 3
                                    card(2) = 42
                                Case 43
                                    frmSpeedGame.picCompCard3.Image = card43
                                    compCard(2) = 4
                                    card(2) = 43
                                Case 44
                                    frmSpeedGame.picCompCard3.Image = card44
                                    compCard(2) = 5
                                    card(2) = 44
                                Case 45
                                    frmSpeedGame.picCompCard3.Image = card45
                                    compCard(2) = 6
                                    card(2) = 45
                                Case 46
                                    frmSpeedGame.picCompCard3.Image = card46
                                    compCard(2) = 7
                                    card(2) = 46
                                Case 47
                                    frmSpeedGame.picCompCard3.Image = card47
                                    compCard(2) = 8
                                    card(2) = 47
                                Case 48
                                    frmSpeedGame.picCompCard3.Image = card48
                                    compCard(2) = 9
                                    card(2) = 48
                                Case 49
                                    frmSpeedGame.picCompCard3.Image = card49
                                    compCard(2) = 10
                                    card(2) = 49
                                Case 50
                                    frmSpeedGame.picCompCard3.Image = card50
                                    compCard(2) = 11
                                    card(2) = 50
                                Case 51
                                    frmSpeedGame.picCompCard3.Image = card51
                                    compCard(2) = 12
                                    card(2) = 51
                                Case 52
                                    frmSpeedGame.picCompCard3.Image = card52
                                    compCard(2) = 13
                                    card(2) = 52
                            End Select
                        Case 4
                            Select Case num
                                Case 1
                                    frmSpeedGame.picCompCard4.Image = card1
                                    compCard(3) = 1
                                    card(3) = 1
                                Case 2
                                    frmSpeedGame.picCompCard4.Image = card2
                                    compCard(3) = 2
                                    card(3) = 2
                                Case 3
                                    frmSpeedGame.picCompCard4.Image = card3
                                    compCard(3) = 3
                                    card(3) = 3
                                Case 4
                                    frmSpeedGame.picCompCard4.Image = card4
                                    compCard(3) = 4
                                    card(3) = 4
                                Case 5
                                    frmSpeedGame.picCompCard4.Image = card5
                                    compCard(3) = 5
                                    card(3) = 5
                                Case 6
                                    frmSpeedGame.picCompCard4.Image = card6
                                    compCard(3) = 6
                                    card(3) = 6
                                Case 7
                                    frmSpeedGame.picCompCard4.Image = card7
                                    compCard(3) = 7
                                    card(3) = 7
                                Case 8
                                    frmSpeedGame.picCompCard4.Image = card8
                                    compCard(3) = 8
                                    card(3) = 8
                                Case 9
                                    frmSpeedGame.picCompCard4.Image = card9
                                    compCard(3) = 9
                                    card(3) = 9
                                Case 10
                                    frmSpeedGame.picCompCard4.Image = card10
                                    compCard(3) = 10
                                    card(3) = 10
                                Case 11
                                    frmSpeedGame.picCompCard4.Image = card11
                                    compCard(3) = 11
                                    card(3) = 11
                                Case 12
                                    frmSpeedGame.picCompCard4.Image = card12
                                    compCard(3) = 12
                                    card(3) = 12
                                Case 13
                                    frmSpeedGame.picCompCard4.Image = card13
                                    compCard(3) = 13
                                    card(3) = 13
                                Case 14
                                    frmSpeedGame.picCompCard4.Image = card14
                                    compCard(3) = 1
                                    card(3) = 14
                                Case 15
                                    frmSpeedGame.picCompCard4.Image = card15
                                    compCard(3) = 2
                                    card(3) = 15
                                Case 16
                                    frmSpeedGame.picCompCard4.Image = card16
                                    compCard(3) = 3
                                    card(3) = 16
                                Case 17
                                    frmSpeedGame.picCompCard4.Image = card17
                                    compCard(3) = 4
                                    card(3) = 17
                                Case 18
                                    frmSpeedGame.picCompCard4.Image = card18
                                    compCard(3) = 5
                                    card(3) = 18
                                Case 19
                                    frmSpeedGame.picCompCard4.Image = card19
                                    compCard(3) = 6
                                    card(3) = 19
                                Case 20
                                    frmSpeedGame.picCompCard4.Image = card20
                                    compCard(3) = 7
                                    card(3) = 20
                                Case 21
                                    frmSpeedGame.picCompCard4.Image = card21
                                    compCard(3) = 8
                                    card(3) = 21
                                Case 22
                                    frmSpeedGame.picCompCard4.Image = card22
                                    compCard(3) = 9
                                    card(3) = 22
                                Case 23
                                    frmSpeedGame.picCompCard4.Image = card23
                                    compCard(3) = 10
                                    card(3) = 23
                                Case 24
                                    frmSpeedGame.picCompCard4.Image = card24
                                    compCard(3) = 11
                                    card(3) = 24
                                Case 25
                                    frmSpeedGame.picCompCard4.Image = card25
                                    compCard(3) = 12
                                    card(3) = 25
                                Case 26
                                    frmSpeedGame.picCompCard4.Image = card26
                                    compCard(3) = 13
                                    card(3) = 26
                                Case 27
                                    frmSpeedGame.picCompCard4.Image = card27
                                    compCard(3) = 1
                                    card(3) = 27
                                Case 28
                                    frmSpeedGame.picCompCard4.Image = card28
                                    compCard(3) = 2
                                    card(3) = 28
                                Case 29
                                    frmSpeedGame.picCompCard4.Image = card29
                                    compCard(3) = 3
                                    card(3) = 29
                                Case 30
                                    frmSpeedGame.picCompCard4.Image = card30
                                    compCard(3) = 4
                                    card(3) = 30
                                Case 31
                                    frmSpeedGame.picCompCard4.Image = card31
                                    compCard(3) = 5
                                    card(3) = 31
                                Case 32
                                    frmSpeedGame.picCompCard4.Image = card32
                                    compCard(3) = 6
                                    card(3) = 32
                                Case 33
                                    frmSpeedGame.picCompCard4.Image = card33
                                    compCard(3) = 7
                                    card(3) = 33
                                Case 34
                                    frmSpeedGame.picCompCard4.Image = card34
                                    compCard(3) = 8
                                    card(3) = 34
                                Case 35
                                    frmSpeedGame.picCompCard4.Image = card35
                                    compCard(3) = 9
                                    card(3) = 35
                                Case 36
                                    frmSpeedGame.picCompCard4.Image = card36
                                    compCard(3) = 10
                                    card(3) = 36
                                Case 37
                                    frmSpeedGame.picCompCard4.Image = card37
                                    compCard(3) = 11
                                    card(3) = 37
                                Case 38
                                    frmSpeedGame.picCompCard4.Image = card38
                                    compCard(3) = 12
                                    card(3) = 38
                                Case 39
                                    frmSpeedGame.picCompCard4.Image = card39
                                    compCard(3) = 13
                                    card(3) = 39
                                Case 40
                                    frmSpeedGame.picCompCard4.Image = card40
                                    compCard(3) = 1
                                    card(3) = 40
                                Case 41
                                    frmSpeedGame.picCompCard4.Image = card41
                                    compCard(3) = 2
                                    card(3) = 41
                                Case 42
                                    frmSpeedGame.picCompCard4.Image = card42
                                    compCard(3) = 3
                                    card(3) = 42
                                Case 43
                                    frmSpeedGame.picCompCard4.Image = card43
                                    compCard(3) = 4
                                    card(3) = 43
                                Case 44
                                    frmSpeedGame.picCompCard4.Image = card44
                                    compCard(3) = 5
                                    card(3) = 44
                                Case 45
                                    frmSpeedGame.picCompCard4.Image = card45
                                    compCard(3) = 6
                                    card(3) = 45
                                Case 46
                                    frmSpeedGame.picCompCard4.Image = card46
                                    compCard(3) = 7
                                    card(3) = 46
                                Case 47
                                    frmSpeedGame.picCompCard4.Image = card47
                                    compCard(3) = 8
                                    card(3) = 47
                                Case 48
                                    frmSpeedGame.picCompCard4.Image = card48
                                    compCard(3) = 9
                                    card(3) = 48
                                Case 49
                                    frmSpeedGame.picCompCard4.Image = card49
                                    compCard(3) = 10
                                    card(3) = 49
                                Case 50
                                    frmSpeedGame.picCompCard4.Image = card50
                                    compCard(3) = 11
                                    card(3) = 50
                                Case 51
                                    frmSpeedGame.picCompCard4.Image = card51
                                    compCard(3) = 12
                                    card(3) = 51
                                Case 52
                                    frmSpeedGame.picCompCard4.Image = card52
                                    compCard(3) = 13
                                    card(3) = 52
                            End Select
                        Case 5
                            Select Case num
                                Case 1
                                    frmSpeedGame.picCompCard5.Image = card1
                                    compCard(4) = 1
                                    card(4) = 1
                                Case 2
                                    frmSpeedGame.picCompCard5.Image = card2
                                    compCard(4) = 2
                                    card(4) = 2
                                Case 3
                                    frmSpeedGame.picCompCard5.Image = card3
                                    compCard(4) = 3
                                    card(4) = 3
                                Case 4
                                    frmSpeedGame.picCompCard5.Image = card4
                                    compCard(4) = 4
                                    card(4) = 4
                                Case 5
                                    frmSpeedGame.picCompCard5.Image = card5
                                    compCard(4) = 5
                                    card(4) = 5
                                Case 6
                                    frmSpeedGame.picCompCard5.Image = card6
                                    compCard(4) = 6
                                    card(4) = 6
                                Case 7
                                    frmSpeedGame.picCompCard5.Image = card7
                                    compCard(4) = 7
                                    card(4) = 7
                                Case 8
                                    frmSpeedGame.picCompCard5.Image = card8
                                    compCard(4) = 8
                                    card(4) = 8
                                Case 9
                                    frmSpeedGame.picCompCard5.Image = card9
                                    compCard(4) = 9
                                    card(4) = 9
                                Case 10
                                    frmSpeedGame.picCompCard5.Image = card10
                                    compCard(4) = 10
                                    card(4) = 10
                                Case 11
                                    frmSpeedGame.picCompCard5.Image = card11
                                    compCard(4) = 11
                                    card(4) = 11
                                Case 12
                                    frmSpeedGame.picCompCard5.Image = card12
                                    compCard(4) = 12
                                    card(4) = 12
                                Case 13
                                    frmSpeedGame.picCompCard5.Image = card13
                                    compCard(4) = 13
                                    card(4) = 13
                                Case 14
                                    frmSpeedGame.picCompCard5.Image = card14
                                    compCard(4) = 1
                                    card(4) = 14
                                Case 15
                                    frmSpeedGame.picCompCard5.Image = card15
                                    compCard(4) = 2
                                    card(4) = 15
                                Case 16
                                    frmSpeedGame.picCompCard5.Image = card16
                                    compCard(4) = 3
                                    card(4) = 16
                                Case 17
                                    frmSpeedGame.picCompCard5.Image = card17
                                    compCard(4) = 4
                                    card(4) = 17
                                Case 18
                                    frmSpeedGame.picCompCard5.Image = card18
                                    compCard(4) = 5
                                    card(4) = 18
                                Case 19
                                    frmSpeedGame.picCompCard5.Image = card19
                                    compCard(4) = 6
                                    card(4) = 19
                                Case 20
                                    frmSpeedGame.picCompCard5.Image = card20
                                    compCard(4) = 7
                                    card(4) = 20
                                Case 21
                                    frmSpeedGame.picCompCard5.Image = card21
                                    compCard(4) = 8
                                    card(4) = 21
                                Case 22
                                    frmSpeedGame.picCompCard5.Image = card22
                                    compCard(4) = 9
                                    card(4) = 22
                                Case 23
                                    frmSpeedGame.picCompCard5.Image = card23
                                    compCard(4) = 10
                                    card(4) = 23
                                Case 24
                                    frmSpeedGame.picCompCard5.Image = card24
                                    compCard(4) = 11
                                    card(4) = 24
                                Case 25
                                    frmSpeedGame.picCompCard5.Image = card25
                                    compCard(4) = 12
                                    card(4) = 25
                                Case 26
                                    frmSpeedGame.picCompCard5.Image = card26
                                    compCard(4) = 13
                                    card(4) = 26
                                Case 27
                                    frmSpeedGame.picCompCard5.Image = card27
                                    compCard(4) = 1
                                    card(4) = 27
                                Case 28
                                    frmSpeedGame.picCompCard5.Image = card28
                                    compCard(4) = 2
                                    card(4) = 28
                                Case 29
                                    frmSpeedGame.picCompCard5.Image = card29
                                    compCard(4) = 3
                                    card(4) = 29
                                Case 30
                                    frmSpeedGame.picCompCard5.Image = card30
                                    compCard(4) = 4
                                    card(4) = 30
                                Case 31
                                    frmSpeedGame.picCompCard5.Image = card31
                                    compCard(4) = 5
                                    card(4) = 31
                                Case 32
                                    frmSpeedGame.picCompCard5.Image = card32
                                    compCard(4) = 6
                                    card(4) = 32
                                Case 33
                                    frmSpeedGame.picCompCard5.Image = card33
                                    compCard(4) = 7
                                    card(4) = 33
                                Case 34
                                    frmSpeedGame.picCompCard5.Image = card34
                                    compCard(4) = 8
                                    card(4) = 34
                                Case 35
                                    frmSpeedGame.picCompCard5.Image = card35
                                    compCard(4) = 9
                                    card(4) = 35
                                Case 36
                                    frmSpeedGame.picCompCard5.Image = card36
                                    compCard(4) = 10
                                    card(4) = 36
                                Case 37
                                    frmSpeedGame.picCompCard5.Image = card37
                                    compCard(4) = 11
                                    card(4) = 37
                                Case 38
                                    frmSpeedGame.picCompCard5.Image = card38
                                    compCard(4) = 12
                                    card(4) = 38
                                Case 39
                                    frmSpeedGame.picCompCard5.Image = card39
                                    compCard(4) = 13
                                    card(4) = 39
                                Case 40
                                    frmSpeedGame.picCompCard5.Image = card40
                                    compCard(4) = 1
                                    card(4) = 40
                                Case 41
                                    frmSpeedGame.picCompCard5.Image = card41
                                    compCard(4) = 2
                                    card(4) = 41
                                Case 42
                                    frmSpeedGame.picCompCard5.Image = card42
                                    compCard(4) = 3
                                    card(4) = 42
                                Case 43
                                    frmSpeedGame.picCompCard5.Image = card43
                                    compCard(4) = 4
                                    card(4) = 43
                                Case 44
                                    frmSpeedGame.picCompCard5.Image = card44
                                    compCard(4) = 5
                                    card(4) = 44
                                Case 45
                                    frmSpeedGame.picCompCard5.Image = card45
                                    compCard(4) = 6
                                    card(4) = 45
                                Case 46
                                    frmSpeedGame.picCompCard5.Image = card46
                                    compCard(4) = 7
                                    card(4) = 46
                                Case 47
                                    frmSpeedGame.picCompCard5.Image = card47
                                    compCard(4) = 8
                                    card(4) = 47
                                Case 48
                                    frmSpeedGame.picCompCard5.Image = card48
                                    compCard(4) = 9
                                    card(4) = 48
                                Case 49
                                    frmSpeedGame.picCompCard5.Image = card49
                                    compCard(4) = 10
                                    card(4) = 49
                                Case 50
                                    frmSpeedGame.picCompCard5.Image = card50
                                    compCard(4) = 11
                                    card(4) = 50
                                Case 51
                                    frmSpeedGame.picCompCard5.Image = card51
                                    compCard(4) = 12
                                    card(4) = 51
                                Case 52
                                    frmSpeedGame.picCompCard5.Image = card52
                                    compCard(4) = 13
                                    card(4) = 52
                            End Select
                        Case 6
                            Select Case num
                                Case 1
                                    frmSpeedGame.picPlayerCard1.Image = card1
                                    playerCard(0) = 1
                                    card(5) = 1
                                Case 2
                                    frmSpeedGame.picPlayerCard1.Image = card2
                                    playerCard(0) = 2
                                    card(5) = 2
                                Case 3
                                    frmSpeedGame.picPlayerCard1.Image = card3
                                    playerCard(0) = 3
                                    card(5) = 3
                                Case 4
                                    frmSpeedGame.picPlayerCard1.Image = card4
                                    playerCard(0) = 4
                                    card(5) = 4
                                Case 5
                                    frmSpeedGame.picPlayerCard1.Image = card5
                                    playerCard(0) = 5
                                    card(5) = 5
                                Case 6
                                    frmSpeedGame.picPlayerCard1.Image = card6
                                    playerCard(0) = 6
                                    card(5) = 6
                                Case 7
                                    frmSpeedGame.picPlayerCard1.Image = card7
                                    playerCard(0) = 7
                                    card(5) = 7
                                Case 8
                                    frmSpeedGame.picPlayerCard1.Image = card8
                                    playerCard(0) = 8
                                    card(5) = 8
                                Case 9
                                    frmSpeedGame.picPlayerCard1.Image = card9
                                    playerCard(0) = 9
                                    card(5) = 9
                                Case 10
                                    frmSpeedGame.picPlayerCard1.Image = card10
                                    playerCard(0) = 10
                                    card(5) = 10
                                Case 11
                                    frmSpeedGame.picPlayerCard1.Image = card11
                                    playerCard(0) = 11
                                    card(5) = 11
                                Case 12
                                    frmSpeedGame.picPlayerCard1.Image = card12
                                    playerCard(0) = 12
                                    card(5) = 12
                                Case 13
                                    frmSpeedGame.picPlayerCard1.Image = card13
                                    playerCard(0) = 13
                                    card(5) = 13
                                Case 14
                                    frmSpeedGame.picPlayerCard1.Image = card14
                                    playerCard(0) = 1
                                    card(5) = 14
                                Case 15
                                    frmSpeedGame.picPlayerCard1.Image = card15
                                    playerCard(0) = 2
                                    card(5) = 15
                                Case 16
                                    frmSpeedGame.picPlayerCard1.Image = card16
                                    playerCard(0) = 3
                                    card(5) = 16
                                Case 17
                                    frmSpeedGame.picPlayerCard1.Image = card17
                                    playerCard(0) = 4
                                    card(5) = 17
                                Case 18
                                    frmSpeedGame.picPlayerCard1.Image = card18
                                    playerCard(0) = 5
                                    card(5) = 18
                                Case 19
                                    frmSpeedGame.picPlayerCard1.Image = card19
                                    playerCard(0) = 6
                                    card(5) = 19
                                Case 20
                                    frmSpeedGame.picPlayerCard1.Image = card20
                                    playerCard(0) = 7
                                    card(5) = 20
                                Case 21
                                    frmSpeedGame.picPlayerCard1.Image = card21
                                    playerCard(0) = 8
                                    card(5) = 21
                                Case 22
                                    frmSpeedGame.picPlayerCard1.Image = card22
                                    playerCard(0) = 9
                                    card(5) = 22
                                Case 23
                                    frmSpeedGame.picPlayerCard1.Image = card23
                                    playerCard(0) = 10
                                    card(5) = 23
                                Case 24
                                    frmSpeedGame.picPlayerCard1.Image = card24
                                    playerCard(0) = 11
                                    card(5) = 24
                                Case 25
                                    frmSpeedGame.picPlayerCard1.Image = card25
                                    playerCard(0) = 12
                                    card(5) = 25
                                Case 26
                                    frmSpeedGame.picPlayerCard1.Image = card26
                                    playerCard(0) = 13
                                    card(5) = 26
                                Case 27
                                    frmSpeedGame.picPlayerCard1.Image = card27
                                    playerCard(0) = 1
                                    card(5) = 27
                                Case 28
                                    frmSpeedGame.picPlayerCard1.Image = card28
                                    playerCard(0) = 2
                                    card(5) = 28
                                Case 29
                                    frmSpeedGame.picPlayerCard1.Image = card29
                                    playerCard(0) = 3
                                    card(5) = 29
                                Case 30
                                    frmSpeedGame.picPlayerCard1.Image = card30
                                    playerCard(0) = 4
                                    card(5) = 30
                                Case 31
                                    frmSpeedGame.picPlayerCard1.Image = card31
                                    playerCard(0) = 5
                                    card(5) = 31
                                Case 32
                                    frmSpeedGame.picPlayerCard1.Image = card32
                                    playerCard(0) = 6
                                    card(5) = 32
                                Case 33
                                    frmSpeedGame.picPlayerCard1.Image = card33
                                    playerCard(0) = 7
                                    card(5) = 33
                                Case 34
                                    frmSpeedGame.picPlayerCard1.Image = card34
                                    playerCard(0) = 8
                                    card(5) = 34
                                Case 35
                                    frmSpeedGame.picPlayerCard1.Image = card35
                                    playerCard(0) = 9
                                    card(5) = 35
                                Case 36
                                    frmSpeedGame.picPlayerCard1.Image = card36
                                    playerCard(0) = 10
                                    card(5) = 36
                                Case 37
                                    frmSpeedGame.picPlayerCard1.Image = card37
                                    playerCard(0) = 11
                                    card(5) = 37
                                Case 38
                                    frmSpeedGame.picPlayerCard1.Image = card38
                                    playerCard(0) = 12
                                    card(5) = 38
                                Case 39
                                    frmSpeedGame.picPlayerCard1.Image = card39
                                    playerCard(0) = 13
                                    card(5) = 39
                                Case 40
                                    frmSpeedGame.picPlayerCard1.Image = card40
                                    playerCard(0) = 1
                                    card(5) = 40
                                Case 41
                                    frmSpeedGame.picPlayerCard1.Image = card41
                                    playerCard(0) = 2
                                    card(5) = 41
                                Case 42
                                    frmSpeedGame.picPlayerCard1.Image = card42
                                    playerCard(0) = 3
                                    card(5) = 42
                                Case 43
                                    frmSpeedGame.picPlayerCard1.Image = card43
                                    playerCard(0) = 4
                                    card(5) = 43
                                Case 44
                                    frmSpeedGame.picPlayerCard1.Image = card44
                                    playerCard(0) = 5
                                    card(5) = 44
                                Case 45
                                    frmSpeedGame.picPlayerCard1.Image = card45
                                    playerCard(0) = 6
                                    card(5) = 45
                                Case 46
                                    frmSpeedGame.picPlayerCard1.Image = card46
                                    playerCard(0) = 7
                                    card(5) = 46
                                Case 47
                                    frmSpeedGame.picPlayerCard1.Image = card47
                                    playerCard(0) = 8
                                    card(5) = 47
                                Case 48
                                    frmSpeedGame.picPlayerCard1.Image = card48
                                    playerCard(0) = 9
                                    card(5) = 48
                                Case 49
                                    frmSpeedGame.picPlayerCard1.Image = card49
                                    playerCard(0) = 10
                                    card(5) = 49
                                Case 50
                                    frmSpeedGame.picPlayerCard1.Image = card50
                                    playerCard(0) = 11
                                    card(5) = 50
                                Case 51
                                    frmSpeedGame.picPlayerCard1.Image = card51
                                    playerCard(0) = 12
                                    card(5) = 51
                                Case 52
                                    frmSpeedGame.picPlayerCard1.Image = card52
                                    playerCard(0) = 13
                                    card(5) = 52
                            End Select
                        Case 7
                            Select Case num
                                Case 1
                                    frmSpeedGame.picPlayerCard2.Image = card1
                                    playerCard(1) = 1
                                    card(6) = 1
                                Case 2
                                    frmSpeedGame.picPlayerCard2.Image = card2
                                    playerCard(1) = 2
                                    card(6) = 2
                                Case 3
                                    frmSpeedGame.picPlayerCard2.Image = card3
                                    playerCard(1) = 3
                                    card(6) = 3
                                Case 4
                                    frmSpeedGame.picPlayerCard2.Image = card4
                                    playerCard(1) = 4
                                    card(6) = 4
                                Case 5
                                    frmSpeedGame.picPlayerCard2.Image = card5
                                    playerCard(1) = 5
                                    card(6) = 5
                                Case 6
                                    frmSpeedGame.picPlayerCard2.Image = card6
                                    playerCard(1) = 6
                                    card(6) = 6
                                Case 7
                                    frmSpeedGame.picPlayerCard2.Image = card7
                                    playerCard(1) = 7
                                    card(6) = 7
                                Case 8
                                    frmSpeedGame.picPlayerCard2.Image = card8
                                    playerCard(1) = 8
                                    card(6) = 8
                                Case 9
                                    frmSpeedGame.picPlayerCard2.Image = card9
                                    playerCard(1) = 9
                                    card(6) = 9
                                Case 10
                                    frmSpeedGame.picPlayerCard2.Image = card10
                                    playerCard(1) = 10
                                    card(6) = 10
                                Case 11
                                    frmSpeedGame.picPlayerCard2.Image = card11
                                    playerCard(1) = 11
                                    card(6) = 11
                                Case 12
                                    frmSpeedGame.picPlayerCard2.Image = card12
                                    playerCard(1) = 12
                                    card(6) = 12
                                Case 13
                                    frmSpeedGame.picPlayerCard2.Image = card13
                                    playerCard(1) = 13
                                    card(6) = 13
                                Case 14
                                    frmSpeedGame.picPlayerCard2.Image = card14
                                    playerCard(1) = 1
                                    card(6) = 14
                                Case 15
                                    frmSpeedGame.picPlayerCard2.Image = card15
                                    playerCard(1) = 2
                                    card(6) = 15
                                Case 16
                                    frmSpeedGame.picPlayerCard2.Image = card16
                                    playerCard(1) = 3
                                    card(6) = 16
                                Case 17
                                    frmSpeedGame.picPlayerCard2.Image = card17
                                    playerCard(1) = 4
                                    card(6) = 17
                                Case 18
                                    frmSpeedGame.picPlayerCard2.Image = card18
                                    playerCard(1) = 5
                                    card(6) = 18
                                Case 19
                                    frmSpeedGame.picPlayerCard2.Image = card19
                                    playerCard(1) = 6
                                    card(6) = 19
                                Case 20
                                    frmSpeedGame.picPlayerCard2.Image = card20
                                    playerCard(1) = 7
                                    card(6) = 20
                                Case 21
                                    frmSpeedGame.picPlayerCard2.Image = card21
                                    playerCard(1) = 8
                                    card(6) = 21
                                Case 22
                                    frmSpeedGame.picPlayerCard2.Image = card22
                                    playerCard(1) = 9
                                    card(6) = 22
                                Case 23
                                    frmSpeedGame.picPlayerCard2.Image = card23
                                    playerCard(1) = 10
                                    card(6) = 23
                                Case 24
                                    frmSpeedGame.picPlayerCard2.Image = card24
                                    playerCard(1) = 11
                                    card(6) = 24
                                Case 25
                                    frmSpeedGame.picPlayerCard2.Image = card25
                                    playerCard(1) = 12
                                    card(6) = 25
                                Case 26
                                    frmSpeedGame.picPlayerCard2.Image = card26
                                    playerCard(1) = 13
                                    card(6) = 26
                                Case 27
                                    frmSpeedGame.picPlayerCard2.Image = card27
                                    playerCard(1) = 1
                                    card(6) = 27
                                Case 28
                                    frmSpeedGame.picPlayerCard2.Image = card28
                                    playerCard(1) = 2
                                    card(6) = 28
                                Case 29
                                    frmSpeedGame.picPlayerCard2.Image = card29
                                    playerCard(1) = 3
                                    card(6) = 29
                                Case 30
                                    frmSpeedGame.picPlayerCard2.Image = card30
                                    playerCard(1) = 4
                                    card(6) = 30
                                Case 31
                                    frmSpeedGame.picPlayerCard2.Image = card31
                                    playerCard(1) = 5
                                    card(6) = 31
                                Case 32
                                    frmSpeedGame.picPlayerCard2.Image = card32
                                    playerCard(1) = 6
                                    card(6) = 32
                                Case 33
                                    frmSpeedGame.picPlayerCard2.Image = card33
                                    playerCard(1) = 7
                                    card(6) = 33
                                Case 34
                                    frmSpeedGame.picPlayerCard2.Image = card34
                                    playerCard(1) = 8
                                    card(6) = 34
                                Case 35
                                    frmSpeedGame.picPlayerCard2.Image = card35
                                    playerCard(1) = 9
                                    card(6) = 35
                                Case 36
                                    frmSpeedGame.picPlayerCard2.Image = card36
                                    playerCard(1) = 10
                                    card(6) = 36
                                Case 37
                                    frmSpeedGame.picPlayerCard2.Image = card37
                                    playerCard(1) = 11
                                    card(6) = 37
                                Case 38
                                    frmSpeedGame.picPlayerCard2.Image = card38
                                    playerCard(1) = 12
                                    card(6) = 38
                                Case 39
                                    frmSpeedGame.picPlayerCard2.Image = card39
                                    playerCard(1) = 13
                                    card(6) = 39
                                Case 40
                                    frmSpeedGame.picPlayerCard2.Image = card40
                                    playerCard(1) = 1
                                    card(6) = 40
                                Case 41
                                    frmSpeedGame.picPlayerCard2.Image = card41
                                    playerCard(1) = 2
                                    card(6) = 41
                                Case 42
                                    frmSpeedGame.picPlayerCard2.Image = card42
                                    playerCard(1) = 3
                                    card(6) = 42
                                Case 43
                                    frmSpeedGame.picPlayerCard2.Image = card43
                                    playerCard(1) = 4
                                    card(6) = 43
                                Case 44
                                    frmSpeedGame.picPlayerCard2.Image = card44
                                    playerCard(1) = 5
                                    card(6) = 44
                                Case 45
                                    frmSpeedGame.picPlayerCard2.Image = card45
                                    playerCard(1) = 6
                                    card(6) = 45
                                Case 46
                                    frmSpeedGame.picPlayerCard2.Image = card46
                                    playerCard(1) = 7
                                    card(6) = 46
                                Case 47
                                    frmSpeedGame.picPlayerCard2.Image = card47
                                    playerCard(1) = 8
                                    card(6) = 47
                                Case 48
                                    frmSpeedGame.picPlayerCard2.Image = card48
                                    playerCard(1) = 9
                                    card(6) = 48
                                Case 49
                                    frmSpeedGame.picPlayerCard2.Image = card49
                                    playerCard(1) = 10
                                    card(6) = 49
                                Case 50
                                    frmSpeedGame.picPlayerCard2.Image = card50
                                    playerCard(1) = 11
                                    card(6) = 50
                                Case 51
                                    frmSpeedGame.picPlayerCard2.Image = card51
                                    playerCard(1) = 12
                                    card(6) = 51
                                Case 52
                                    frmSpeedGame.picPlayerCard2.Image = card52
                                    playerCard(1) = 13
                                    card(6) = 52
                            End Select
                        Case 8
                            Select Case num                                               'continue
                                Case 1
                                    frmSpeedGame.picPlayerCard3.Image = card1
                                    playerCard(2) = 1
                                    card(7) = 1
                                Case 2
                                    frmSpeedGame.picPlayerCard3.Image = card2
                                    playerCard(2) = 2
                                    card(7) = 2
                                Case 3
                                    frmSpeedGame.picPlayerCard3.Image = card3
                                    playerCard(2) = 3
                                    card(7) = 3
                                Case 4
                                    frmSpeedGame.picPlayerCard3.Image = card4
                                    playerCard(2) = 4
                                    card(7) = 4
                                Case 5
                                    frmSpeedGame.picPlayerCard3.Image = card5
                                    playerCard(2) = 5
                                    card(7) = 5
                                Case 6
                                    frmSpeedGame.picPlayerCard3.Image = card6
                                    playerCard(2) = 6
                                    card(7) = 6
                                Case 7
                                    frmSpeedGame.picPlayerCard3.Image = card7
                                    playerCard(2) = 7
                                    card(7) = 7
                                Case 8
                                    frmSpeedGame.picPlayerCard3.Image = card8
                                    playerCard(2) = 8
                                    card(7) = 8
                                Case 9
                                    frmSpeedGame.picPlayerCard3.Image = card9
                                    playerCard(2) = 9
                                    card(7) = 9
                                Case 10
                                    frmSpeedGame.picPlayerCard3.Image = card10
                                    playerCard(2) = 10
                                    card(7) = 10
                                Case 11
                                    frmSpeedGame.picPlayerCard3.Image = card11
                                    playerCard(2) = 11
                                    card(7) = 11
                                Case 12
                                    frmSpeedGame.picPlayerCard3.Image = card12
                                    playerCard(2) = 12
                                    card(7) = 12
                                Case 13
                                    frmSpeedGame.picPlayerCard3.Image = card13
                                    playerCard(2) = 13
                                    card(7) = 13
                                Case 14
                                    frmSpeedGame.picPlayerCard3.Image = card14
                                    playerCard(2) = 1
                                    card(7) = 14
                                Case 15
                                    frmSpeedGame.picPlayerCard3.Image = card15
                                    playerCard(2) = 2
                                    card(7) = 15
                                Case 16
                                    frmSpeedGame.picPlayerCard3.Image = card16
                                    playerCard(2) = 3
                                    card(7) = 16
                                Case 17
                                    frmSpeedGame.picPlayerCard3.Image = card17
                                    playerCard(2) = 4
                                    card(7) = 17
                                Case 18
                                    frmSpeedGame.picPlayerCard3.Image = card18
                                    playerCard(2) = 5
                                    card(7) = 18
                                Case 19
                                    frmSpeedGame.picPlayerCard3.Image = card19
                                    playerCard(2) = 6
                                    card(7) = 19
                                Case 20
                                    frmSpeedGame.picPlayerCard3.Image = card20
                                    playerCard(2) = 7
                                    card(7) = 20
                                Case 21
                                    frmSpeedGame.picPlayerCard3.Image = card21
                                    playerCard(2) = 8
                                    card(7) = 21
                                Case 22
                                    frmSpeedGame.picPlayerCard3.Image = card22
                                    playerCard(2) = 9
                                    card(7) = 22
                                Case 23
                                    frmSpeedGame.picPlayerCard3.Image = card23
                                    playerCard(2) = 10
                                    card(7) = 23
                                Case 24
                                    frmSpeedGame.picPlayerCard3.Image = card24
                                    playerCard(2) = 11
                                    card(7) = 24
                                Case 25
                                    frmSpeedGame.picPlayerCard3.Image = card25
                                    playerCard(2) = 12
                                    card(7) = 25
                                Case 26
                                    frmSpeedGame.picPlayerCard3.Image = card26
                                    playerCard(2) = 13
                                    card(7) = 26
                                Case 27
                                    frmSpeedGame.picPlayerCard3.Image = card27
                                    playerCard(2) = 1
                                    card(7) = 27
                                Case 28
                                    frmSpeedGame.picPlayerCard3.Image = card28
                                    playerCard(2) = 2
                                    card(7) = 28
                                Case 29
                                    frmSpeedGame.picPlayerCard3.Image = card29
                                    playerCard(2) = 3
                                    card(7) = 29
                                Case 30
                                    frmSpeedGame.picPlayerCard3.Image = card30
                                    playerCard(2) = 4
                                    card(7) = 30
                                Case 31
                                    frmSpeedGame.picPlayerCard3.Image = card31
                                    playerCard(2) = 5
                                    card(7) = 31
                                Case 32
                                    frmSpeedGame.picPlayerCard3.Image = card32
                                    playerCard(2) = 6
                                    card(7) = 32
                                Case 33
                                    frmSpeedGame.picPlayerCard3.Image = card33
                                    playerCard(2) = 7
                                    card(7) = 33
                                Case 34
                                    frmSpeedGame.picPlayerCard3.Image = card34
                                    playerCard(2) = 8
                                    card(7) = 34
                                Case 35
                                    frmSpeedGame.picPlayerCard3.Image = card35
                                    playerCard(2) = 9
                                    card(7) = 35
                                Case 36
                                    frmSpeedGame.picPlayerCard3.Image = card36
                                    playerCard(2) = 10
                                    card(7) = 36
                                Case 37
                                    frmSpeedGame.picPlayerCard3.Image = card37
                                    playerCard(2) = 11
                                    card(7) = 37
                                Case 38
                                    frmSpeedGame.picPlayerCard3.Image = card38
                                    playerCard(2) = 12
                                    card(7) = 38
                                Case 39
                                    frmSpeedGame.picPlayerCard3.Image = card39
                                    playerCard(2) = 13
                                    card(7) = 39
                                Case 40
                                    frmSpeedGame.picPlayerCard3.Image = card40
                                    playerCard(2) = 1
                                    card(7) = 40
                                Case 41
                                    frmSpeedGame.picPlayerCard3.Image = card41
                                    playerCard(2) = 2
                                    card(7) = 41
                                Case 42
                                    frmSpeedGame.picPlayerCard3.Image = card42
                                    playerCard(2) = 3
                                    card(7) = 42
                                Case 43
                                    frmSpeedGame.picPlayerCard3.Image = card43
                                    playerCard(2) = 4
                                    card(7) = 43
                                Case 44
                                    frmSpeedGame.picPlayerCard3.Image = card44
                                    playerCard(2) = 5
                                    card(7) = 44
                                Case 45
                                    frmSpeedGame.picPlayerCard3.Image = card45
                                    playerCard(2) = 6
                                    card(7) = 45
                                Case 46
                                    frmSpeedGame.picPlayerCard3.Image = card46
                                    playerCard(2) = 7
                                    card(7) = 46
                                Case 47
                                    frmSpeedGame.picPlayerCard3.Image = card47
                                    playerCard(2) = 8
                                    card(7) = 47
                                Case 48
                                    frmSpeedGame.picPlayerCard3.Image = card48
                                    playerCard(2) = 9
                                    card(7) = 48
                                Case 49
                                    frmSpeedGame.picPlayerCard3.Image = card49
                                    playerCard(2) = 10
                                    card(7) = 49
                                Case 50
                                    frmSpeedGame.picPlayerCard3.Image = card50
                                    playerCard(2) = 11
                                    card(7) = 50
                                Case 51
                                    frmSpeedGame.picPlayerCard3.Image = card51
                                    playerCard(2) = 12
                                    card(7) = 51
                                Case 52
                                    frmSpeedGame.picPlayerCard3.Image = card52
                                    playerCard(2) = 13
                                    card(7) = 52
                            End Select
                        Case 9
                            Select Case num
                                Case 1
                                    frmSpeedGame.picPlayerCard4.Image = card1
                                    playerCard(3) = 1
                                    card(8) = 1
                                Case 2
                                    frmSpeedGame.picPlayerCard4.Image = card2
                                    playerCard(3) = 2
                                    card(8) = 2
                                Case 3
                                    frmSpeedGame.picPlayerCard4.Image = card3
                                    playerCard(3) = 3
                                    card(8) = 3
                                Case 4
                                    frmSpeedGame.picPlayerCard4.Image = card4
                                    playerCard(3) = 4
                                    card(8) = 4
                                Case 5
                                    frmSpeedGame.picPlayerCard4.Image = card5
                                    playerCard(3) = 5
                                    card(8) = 5
                                Case 6
                                    frmSpeedGame.picPlayerCard4.Image = card6
                                    playerCard(3) = 6
                                    card(8) = 6
                                Case 7
                                    frmSpeedGame.picPlayerCard4.Image = card7
                                    playerCard(3) = 7
                                    card(8) = 7
                                Case 8
                                    frmSpeedGame.picPlayerCard4.Image = card8
                                    playerCard(3) = 8
                                    card(8) = 8
                                Case 9
                                    frmSpeedGame.picPlayerCard4.Image = card9
                                    playerCard(3) = 9
                                    card(8) = 9
                                Case 10
                                    frmSpeedGame.picPlayerCard4.Image = card10
                                    playerCard(3) = 10
                                    card(8) = 10
                                Case 11
                                    frmSpeedGame.picPlayerCard4.Image = card11
                                    playerCard(3) = 11
                                    card(8) = 11
                                Case 12
                                    frmSpeedGame.picPlayerCard4.Image = card12
                                    playerCard(3) = 12
                                    card(8) = 12
                                Case 13
                                    frmSpeedGame.picPlayerCard4.Image = card13
                                    playerCard(3) = 13
                                    card(8) = 13
                                Case 14
                                    frmSpeedGame.picPlayerCard4.Image = card14
                                    playerCard(3) = 1
                                    card(8) = 14
                                Case 15
                                    frmSpeedGame.picPlayerCard4.Image = card15
                                    playerCard(3) = 2
                                    card(8) = 15
                                Case 16
                                    frmSpeedGame.picPlayerCard4.Image = card16
                                    playerCard(3) = 3
                                    card(8) = 16
                                Case 17
                                    frmSpeedGame.picPlayerCard4.Image = card17
                                    playerCard(3) = 4
                                    card(8) = 17
                                Case 18
                                    frmSpeedGame.picPlayerCard4.Image = card18
                                    playerCard(3) = 5
                                    card(8) = 18
                                Case 19
                                    frmSpeedGame.picPlayerCard4.Image = card19
                                    playerCard(3) = 6
                                    card(8) = 19
                                Case 20
                                    frmSpeedGame.picPlayerCard4.Image = card20
                                    playerCard(3) = 7
                                    card(8) = 20
                                Case 21
                                    frmSpeedGame.picPlayerCard4.Image = card21
                                    playerCard(3) = 8
                                    card(8) = 21
                                Case 22
                                    frmSpeedGame.picPlayerCard4.Image = card22
                                    playerCard(3) = 9
                                    card(8) = 22
                                Case 23
                                    frmSpeedGame.picPlayerCard4.Image = card23
                                    playerCard(3) = 10
                                    card(8) = 23
                                Case 24
                                    frmSpeedGame.picPlayerCard4.Image = card24
                                    playerCard(3) = 11
                                    card(8) = 24
                                Case 25
                                    frmSpeedGame.picPlayerCard4.Image = card25
                                    playerCard(3) = 12
                                    card(8) = 25
                                Case 26
                                    frmSpeedGame.picPlayerCard4.Image = card26
                                    playerCard(3) = 13
                                    card(8) = 26
                                Case 27
                                    frmSpeedGame.picPlayerCard4.Image = card27
                                    playerCard(3) = 1
                                    card(8) = 27
                                Case 28
                                    frmSpeedGame.picPlayerCard4.Image = card28
                                    playerCard(3) = 2
                                    card(8) = 28
                                Case 29
                                    frmSpeedGame.picPlayerCard4.Image = card29
                                    playerCard(3) = 3
                                    card(8) = 29
                                Case 30
                                    frmSpeedGame.picPlayerCard4.Image = card30
                                    playerCard(3) = 4
                                    card(8) = 30
                                Case 31
                                    frmSpeedGame.picPlayerCard4.Image = card31
                                    playerCard(3) = 5
                                    card(8) = 31
                                Case 32
                                    frmSpeedGame.picPlayerCard4.Image = card32
                                    playerCard(3) = 6
                                    card(8) = 32
                                Case 33
                                    frmSpeedGame.picPlayerCard4.Image = card33
                                    playerCard(3) = 7
                                    card(8) = 33
                                Case 34
                                    frmSpeedGame.picPlayerCard4.Image = card34
                                    playerCard(3) = 8
                                    card(8) = 34
                                Case 35
                                    frmSpeedGame.picPlayerCard4.Image = card35
                                    playerCard(3) = 9
                                    card(8) = 35
                                Case 36
                                    frmSpeedGame.picPlayerCard4.Image = card36
                                    playerCard(3) = 10
                                    card(8) = 36
                                Case 37
                                    frmSpeedGame.picPlayerCard4.Image = card37
                                    playerCard(3) = 11
                                    card(8) = 37
                                Case 38
                                    frmSpeedGame.picPlayerCard4.Image = card38
                                    playerCard(3) = 12
                                    card(8) = 38
                                Case 39
                                    frmSpeedGame.picPlayerCard4.Image = card39
                                    playerCard(3) = 13
                                    card(8) = 39
                                Case 40
                                    frmSpeedGame.picPlayerCard4.Image = card40
                                    playerCard(3) = 1
                                    card(8) = 40
                                Case 41
                                    frmSpeedGame.picPlayerCard4.Image = card41
                                    playerCard(3) = 2
                                    card(8) = 41
                                Case 42
                                    frmSpeedGame.picPlayerCard4.Image = card42
                                    playerCard(3) = 3
                                    card(8) = 42
                                Case 43
                                    frmSpeedGame.picPlayerCard4.Image = card43
                                    playerCard(3) = 4
                                    card(8) = 43
                                Case 44
                                    frmSpeedGame.picPlayerCard4.Image = card44
                                    playerCard(3) = 5
                                    card(8) = 44
                                Case 45
                                    frmSpeedGame.picPlayerCard4.Image = card45
                                    playerCard(3) = 6
                                    card(8) = 45
                                Case 46
                                    frmSpeedGame.picPlayerCard4.Image = card46
                                    playerCard(3) = 7
                                    card(8) = 46
                                Case 47
                                    frmSpeedGame.picPlayerCard4.Image = card47
                                    playerCard(3) = 8
                                    card(8) = 47
                                Case 48
                                    frmSpeedGame.picPlayerCard4.Image = card48
                                    playerCard(3) = 9
                                    card(8) = 48
                                Case 49
                                    frmSpeedGame.picPlayerCard4.Image = card49
                                    playerCard(3) = 10
                                    card(8) = 49
                                Case 50
                                    frmSpeedGame.picPlayerCard4.Image = card50
                                    playerCard(3) = 11
                                    card(8) = 50
                                Case 51
                                    frmSpeedGame.picPlayerCard4.Image = card51
                                    playerCard(3) = 12
                                    card(8) = 51
                                Case 52
                                    frmSpeedGame.picPlayerCard4.Image = card52
                                    playerCard(3) = 13
                                    card(8) = 52
                            End Select
                        Case 10
                            Select Case num
                                Case 1
                                    frmSpeedGame.picPlayerCard5.Image = card1
                                    playerCard(4) = 1
                                    card(9) = 1
                                Case 2
                                    frmSpeedGame.picPlayerCard5.Image = card2
                                    playerCard(4) = 2
                                    card(9) = 2
                                Case 3
                                    frmSpeedGame.picPlayerCard5.Image = card3
                                    playerCard(4) = 3
                                    card(9) = 3
                                Case 4
                                    frmSpeedGame.picPlayerCard5.Image = card4
                                    playerCard(4) = 4
                                    card(9) = 4
                                Case 5
                                    frmSpeedGame.picPlayerCard5.Image = card5
                                    playerCard(4) = 5
                                    card(9) = 5
                                Case 6
                                    frmSpeedGame.picPlayerCard5.Image = card6
                                    playerCard(4) = 6
                                    card(9) = 6
                                Case 7
                                    frmSpeedGame.picPlayerCard5.Image = card7
                                    playerCard(4) = 7
                                    card(9) = 7
                                Case 8
                                    frmSpeedGame.picPlayerCard5.Image = card8
                                    playerCard(4) = 8
                                    card(9) = 8
                                Case 9
                                    frmSpeedGame.picPlayerCard5.Image = card9
                                    playerCard(4) = 9
                                    card(9) = 9
                                Case 10
                                    frmSpeedGame.picPlayerCard5.Image = card10
                                    playerCard(4) = 10
                                    card(9) = 10
                                Case 11
                                    frmSpeedGame.picPlayerCard5.Image = card11
                                    playerCard(4) = 11
                                    card(9) = 11
                                Case 12
                                    frmSpeedGame.picPlayerCard5.Image = card12
                                    playerCard(4) = 12
                                    card(9) = 12
                                Case 13
                                    frmSpeedGame.picPlayerCard5.Image = card13
                                    playerCard(4) = 13
                                    card(9) = 13
                                Case 14
                                    frmSpeedGame.picPlayerCard5.Image = card14
                                    playerCard(4) = 1
                                    card(9) = 14
                                Case 15
                                    frmSpeedGame.picPlayerCard5.Image = card15
                                    playerCard(4) = 2
                                    card(9) = 15
                                Case 16
                                    frmSpeedGame.picPlayerCard5.Image = card16
                                    playerCard(4) = 3
                                    card(9) = 16
                                Case 17
                                    frmSpeedGame.picPlayerCard5.Image = card17
                                    playerCard(4) = 4
                                    card(9) = 17
                                Case 18
                                    frmSpeedGame.picPlayerCard5.Image = card18
                                    playerCard(4) = 5
                                    card(9) = 18
                                Case 19
                                    frmSpeedGame.picPlayerCard5.Image = card19
                                    playerCard(4) = 6
                                    card(9) = 19
                                Case 20
                                    frmSpeedGame.picPlayerCard5.Image = card20
                                    playerCard(4) = 7
                                    card(9) = 20
                                Case 21
                                    frmSpeedGame.picPlayerCard5.Image = card21
                                    playerCard(4) = 8
                                    card(9) = 21
                                Case 22
                                    frmSpeedGame.picPlayerCard5.Image = card22
                                    playerCard(4) = 9
                                    card(9) = 22
                                Case 23
                                    frmSpeedGame.picPlayerCard5.Image = card23
                                    playerCard(4) = 10
                                    card(9) = 23
                                Case 24
                                    frmSpeedGame.picPlayerCard5.Image = card24
                                    playerCard(4) = 11
                                    card(9) = 24
                                Case 25
                                    frmSpeedGame.picPlayerCard5.Image = card25
                                    playerCard(4) = 12
                                    card(9) = 25
                                Case 26
                                    frmSpeedGame.picPlayerCard5.Image = card26
                                    playerCard(4) = 13
                                    card(9) = 26
                                Case 27
                                    frmSpeedGame.picPlayerCard5.Image = card27
                                    playerCard(4) = 1
                                    card(9) = 27
                                Case 28
                                    frmSpeedGame.picPlayerCard5.Image = card28
                                    playerCard(4) = 2
                                    card(9) = 28
                                Case 29
                                    frmSpeedGame.picPlayerCard5.Image = card29
                                    playerCard(4) = 3
                                    card(9) = 29
                                Case 30
                                    frmSpeedGame.picPlayerCard5.Image = card30
                                    playerCard(4) = 4
                                    card(9) = 30
                                Case 31
                                    frmSpeedGame.picPlayerCard5.Image = card31
                                    playerCard(4) = 5
                                    card(9) = 31
                                Case 32
                                    frmSpeedGame.picPlayerCard5.Image = card32
                                    playerCard(4) = 6
                                    card(9) = 32
                                Case 33
                                    frmSpeedGame.picPlayerCard5.Image = card33
                                    playerCard(4) = 7
                                    card(9) = 33
                                Case 34
                                    frmSpeedGame.picPlayerCard5.Image = card34
                                    playerCard(4) = 8
                                    card(9) = 34
                                Case 35
                                    frmSpeedGame.picPlayerCard5.Image = card35
                                    playerCard(4) = 9
                                    card(9) = 35
                                Case 36
                                    frmSpeedGame.picPlayerCard5.Image = card36
                                    playerCard(4) = 10
                                    card(9) = 36
                                Case 37
                                    frmSpeedGame.picPlayerCard5.Image = card37
                                    playerCard(4) = 11
                                    card(9) = 37
                                Case 38
                                    frmSpeedGame.picPlayerCard5.Image = card38
                                    playerCard(4) = 12
                                    card(9) = 38
                                Case 39
                                    frmSpeedGame.picPlayerCard5.Image = card39
                                    playerCard(4) = 13
                                    card(9) = 39
                                Case 40
                                    frmSpeedGame.picPlayerCard5.Image = card40
                                    playerCard(4) = 1
                                    card(9) = 40
                                Case 41
                                    frmSpeedGame.picPlayerCard5.Image = card41
                                    playerCard(4) = 2
                                    card(9) = 41
                                Case 42
                                    frmSpeedGame.picPlayerCard5.Image = card42
                                    playerCard(4) = 3
                                    card(9) = 42
                                Case 43
                                    frmSpeedGame.picPlayerCard5.Image = card43
                                    playerCard(4) = 4
                                    card(9) = 43
                                Case 44
                                    frmSpeedGame.picPlayerCard5.Image = card44
                                    playerCard(4) = 5
                                    card(9) = 44
                                Case 45
                                    frmSpeedGame.picPlayerCard5.Image = card45
                                    playerCard(4) = 6
                                    card(9) = 45
                                Case 46
                                    frmSpeedGame.picPlayerCard5.Image = card46
                                    playerCard(4) = 7
                                    card(9) = 46
                                Case 47
                                    frmSpeedGame.picPlayerCard5.Image = card47
                                    playerCard(4) = 8
                                    card(9) = 47
                                Case 48
                                    frmSpeedGame.picPlayerCard5.Image = card48
                                    playerCard(4) = 9
                                    card(9) = 48
                                Case 49
                                    frmSpeedGame.picPlayerCard5.Image = card49
                                    playerCard(4) = 10
                                    card(9) = 49
                                Case 50
                                    frmSpeedGame.picPlayerCard5.Image = card50
                                    playerCard(4) = 11
                                    card(9) = 50
                                Case 51
                                    frmSpeedGame.picPlayerCard5.Image = card51
                                    playerCard(4) = 12
                                    card(9) = 51
                                Case 52
                                    frmSpeedGame.picPlayerCard5.Image = card52
                                    playerCard(4) = 13
                                    card(9) = 52
                            End Select
                        Case 11
                            Select Case num
                                Case 1
                                    frmSpeedGame.picCompPile.Image = card1
                                    compPile = 1
                                    card(10) = 1
                                Case 2
                                    frmSpeedGame.picCompPile.Image = card2
                                    compPile = 2
                                    card(10) = 2
                                Case 3
                                    frmSpeedGame.picCompPile.Image = card3
                                    compPile = 3
                                    card(10) = 3
                                Case 4
                                    frmSpeedGame.picCompPile.Image = card4
                                    compPile = 4
                                    card(10) = 4
                                Case 5
                                    frmSpeedGame.picCompPile.Image = card5
                                    compPile = 5
                                    card(10) = 5
                                Case 6
                                    frmSpeedGame.picCompPile.Image = card6
                                    compPile = 6
                                    card(10) = 6
                                Case 7
                                    frmSpeedGame.picCompPile.Image = card7
                                    compPile = 7
                                    card(10) = 7
                                Case 8
                                    frmSpeedGame.picCompPile.Image = card8
                                    compPile = 8
                                    card(10) = 8
                                Case 9
                                    frmSpeedGame.picCompPile.Image = card9
                                    compPile = 9
                                    card(10) = 9
                                Case 10
                                    frmSpeedGame.picCompPile.Image = card10
                                    compPile = 10
                                    card(10) = 10
                                Case 11
                                    frmSpeedGame.picCompPile.Image = card11
                                    compPile = 11
                                    card(10) = 11
                                Case 12
                                    frmSpeedGame.picCompPile.Image = card12
                                    compPile = 1
                                    card(10) = 12
                                Case 13
                                    frmSpeedGame.picCompPile.Image = card13
                                    compPile = 13
                                    card(10) = 13
                                Case 14
                                    frmSpeedGame.picCompPile.Image = card14
                                    compPile = 1
                                    card(10) = 14
                                Case 15
                                    frmSpeedGame.picCompPile.Image = card15
                                    compPile = 2
                                    card(10) = 15
                                Case 16
                                    frmSpeedGame.picCompPile.Image = card16
                                    compPile = 3
                                    card(10) = 16
                                Case 17
                                    frmSpeedGame.picCompPile.Image = card17
                                    compPile = 4
                                    card(10) = 17
                                Case 18
                                    frmSpeedGame.picCompPile.Image = card18
                                    compPile = 5
                                    card(10) = 18
                                Case 19
                                    frmSpeedGame.picCompPile.Image = card19
                                    compPile = 6
                                    card(10) = 19
                                Case 20
                                    frmSpeedGame.picCompPile.Image = card20
                                    compPile = 7
                                    card(10) = 20
                                Case 21
                                    frmSpeedGame.picCompPile.Image = card21
                                    compPile = 8
                                    card(10) = 21
                                Case 22
                                    frmSpeedGame.picCompPile.Image = card22
                                    compPile = 9
                                    card(10) = 22
                                Case 23
                                    frmSpeedGame.picCompPile.Image = card23
                                    compPile = 10
                                    card(10) = 23
                                Case 24
                                    frmSpeedGame.picCompPile.Image = card24
                                    compPile = 11
                                    card(10) = 24
                                Case 25
                                    frmSpeedGame.picCompPile.Image = card25
                                    compPile = 12
                                    card(10) = 25
                                Case 26
                                    frmSpeedGame.picCompPile.Image = card26
                                    compPile = 13
                                    card(10) = 26
                                Case 27
                                    frmSpeedGame.picCompPile.Image = card27
                                    compPile = 1
                                    card(10) = 27
                                Case 28
                                    frmSpeedGame.picCompPile.Image = card28
                                    compPile = 2
                                    card(10) = 28
                                Case 29
                                    frmSpeedGame.picCompPile.Image = card29
                                    compPile = 3
                                    card(10) = 29
                                Case 30
                                    frmSpeedGame.picCompPile.Image = card30
                                    compPile = 4
                                    card(10) = 30
                                Case 31
                                    frmSpeedGame.picCompPile.Image = card31
                                    compPile = 5
                                    card(10) = 31
                                Case 32
                                    frmSpeedGame.picCompPile.Image = card32
                                    compPile = 6
                                    card(10) = 32
                                Case 33
                                    frmSpeedGame.picCompPile.Image = card33
                                    compPile = 7
                                    card(10) = 33
                                Case 34
                                    frmSpeedGame.picCompPile.Image = card34
                                    compPile = 8
                                    card(10) = 34
                                Case 35
                                    frmSpeedGame.picCompPile.Image = card35
                                    compPile = 9
                                    card(10) = 35
                                Case 36
                                    frmSpeedGame.picCompPile.Image = card36
                                    compPile = 10
                                    card(10) = 36
                                Case 37
                                    frmSpeedGame.picCompPile.Image = card37
                                    compPile = 11
                                    card(10) = 37
                                Case 38
                                    frmSpeedGame.picCompPile.Image = card38
                                    compPile = 12
                                    card(10) = 38
                                Case 39
                                    frmSpeedGame.picCompPile.Image = card39
                                    compPile = 13
                                    card(10) = 39
                                Case 40
                                    frmSpeedGame.picCompPile.Image = card40
                                    compPile = 1
                                    card(10) = 40
                                Case 41
                                    frmSpeedGame.picCompPile.Image = card41
                                    compPile = 2
                                    card(10) = 41
                                Case 42
                                    frmSpeedGame.picCompPile.Image = card42
                                    compPile = 3
                                    card(10) = 42
                                Case 43
                                    frmSpeedGame.picCompPile.Image = card43
                                    compPile = 4
                                    card(10) = 43
                                Case 44
                                    frmSpeedGame.picCompPile.Image = card44
                                    compPile = 5
                                    card(10) = 44
                                Case 45
                                    frmSpeedGame.picCompPile.Image = card45
                                    compPile = 6
                                    card(10) = 45
                                Case 46
                                    frmSpeedGame.picCompPile.Image = card46
                                    compPile = 7
                                    card(10) = 46
                                Case 47
                                    frmSpeedGame.picCompPile.Image = card47
                                    compPile = 8
                                    card(10) = 47
                                Case 48
                                    frmSpeedGame.picCompPile.Image = card48
                                    compPile = 9
                                    card(10) = 48
                                Case 49
                                    frmSpeedGame.picCompPile.Image = card49
                                    compPile = 10
                                    card(10) = 49
                                Case 50
                                    frmSpeedGame.picCompPile.Image = card50
                                    compPile = 11
                                    card(10) = 50
                                Case 51
                                    frmSpeedGame.picCompPile.Image = card51
                                    compPile = 12
                                    card(10) = 51
                                Case 52
                                    frmSpeedGame.picCompPile.Image = card52
                                    compPile = 13
                                    card(10) = 52
                            End Select
                        Case 12
                            Select Case num
                                Case 1
                                    frmSpeedGame.picPlayerPile.Image = card1
                                    playerPile = 1
                                    card(11) = 1
                                Case 2
                                    frmSpeedGame.picPlayerPile.Image = card2
                                    playerPile = 2
                                    card(11) = 2
                                Case 3
                                    frmSpeedGame.picPlayerPile.Image = card3
                                    playerPile = 3
                                    card(11) = 3
                                Case 4
                                    frmSpeedGame.picPlayerPile.Image = card4
                                    playerPile = 4
                                    card(11) = 4
                                Case 5
                                    frmSpeedGame.picPlayerPile.Image = card5
                                    playerPile = 5
                                    card(11) = 5
                                Case 6
                                    frmSpeedGame.picPlayerPile.Image = card6
                                    playerPile = 6
                                    card(11) = 6
                                Case 7
                                    frmSpeedGame.picPlayerPile.Image = card7
                                    playerPile = 7
                                    card(11) = 7
                                Case 8
                                    frmSpeedGame.picPlayerPile.Image = card8
                                    playerPile = 8
                                    card(11) = 8
                                Case 9
                                    frmSpeedGame.picPlayerPile.Image = card9
                                    playerPile = 9
                                    card(11) = 9
                                Case 10
                                    frmSpeedGame.picPlayerPile.Image = card10
                                    playerPile = 10
                                    card(11) = 10
                                Case 11
                                    frmSpeedGame.picPlayerPile.Image = card11
                                    playerPile = 11
                                    card(11) = 11
                                Case 12
                                    frmSpeedGame.picPlayerPile.Image = card12
                                    playerPile = 12
                                    card(11) = 12
                                Case 13
                                    frmSpeedGame.picPlayerPile.Image = card13
                                    playerPile = 13
                                    card(11) = 13
                                Case 14
                                    frmSpeedGame.picPlayerPile.Image = card14
                                    playerPile = 1
                                    card(11) = 14
                                Case 15
                                    frmSpeedGame.picPlayerPile.Image = card15
                                    playerPile = 2
                                    card(11) = 15
                                Case 16
                                    frmSpeedGame.picPlayerPile.Image = card16
                                    playerPile = 3
                                    card(11) = 16
                                Case 17
                                    frmSpeedGame.picPlayerPile.Image = card17
                                    playerPile = 4
                                    card(11) = 17
                                Case 18
                                    frmSpeedGame.picPlayerPile.Image = card18
                                    playerPile = 5
                                    card(11) = 18
                                Case 19
                                    frmSpeedGame.picPlayerPile.Image = card19
                                    playerPile = 6
                                    card(11) = 19
                                Case 20
                                    frmSpeedGame.picPlayerPile.Image = card20
                                    playerPile = 7
                                    card(11) = 20
                                Case 21
                                    frmSpeedGame.picPlayerPile.Image = card21
                                    playerPile = 8
                                    card(11) = 21
                                Case 22
                                    frmSpeedGame.picPlayerPile.Image = card22
                                    playerPile = 9
                                    card(11) = 22
                                Case 23
                                    frmSpeedGame.picPlayerPile.Image = card23
                                    playerPile = 10
                                    card(11) = 23
                                Case 24
                                    frmSpeedGame.picPlayerPile.Image = card24
                                    playerPile = 11
                                    card(11) = 24
                                Case 25
                                    frmSpeedGame.picPlayerPile.Image = card25
                                    playerPile = 12
                                    card(11) = 25
                                Case 26
                                    frmSpeedGame.picPlayerPile.Image = card26
                                    playerPile = 13
                                    card(11) = 26
                                Case 27
                                    frmSpeedGame.picPlayerPile.Image = card27
                                    playerPile = 1
                                    card(11) = 27
                                Case 28
                                    frmSpeedGame.picPlayerPile.Image = card28
                                    playerPile = 2
                                    card(11) = 28
                                Case 29
                                    frmSpeedGame.picPlayerPile.Image = card29
                                    playerPile = 3
                                    card(11) = 29
                                Case 30
                                    frmSpeedGame.picPlayerPile.Image = card30
                                    playerPile = 4
                                    card(11) = 30
                                Case 31
                                    frmSpeedGame.picPlayerPile.Image = card31
                                    playerPile = 5
                                    card(11) = 31
                                Case 32
                                    frmSpeedGame.picPlayerPile.Image = card32
                                    playerPile = 6
                                    card(11) = 32
                                Case 33
                                    frmSpeedGame.picPlayerPile.Image = card33
                                    playerPile = 7
                                    card(11) = 33
                                Case 34
                                    frmSpeedGame.picPlayerPile.Image = card34
                                    playerPile = 8
                                    card(11) = 34
                                Case 35
                                    frmSpeedGame.picPlayerPile.Image = card35
                                    playerPile = 9
                                    card(11) = 35
                                Case 36
                                    frmSpeedGame.picPlayerPile.Image = card36
                                    playerPile = 10
                                    card(11) = 36
                                Case 37
                                    frmSpeedGame.picPlayerPile.Image = card37
                                    playerPile = 11
                                    card(11) = 37
                                Case 38
                                    frmSpeedGame.picPlayerPile.Image = card38
                                    playerPile = 12
                                    card(11) = 38
                                Case 39
                                    frmSpeedGame.picPlayerPile.Image = card39
                                    playerPile = 13
                                    card(11) = 39
                                Case 40
                                    frmSpeedGame.picPlayerPile.Image = card40
                                    playerPile = 1
                                    card(11) = 40
                                Case 41
                                    frmSpeedGame.picPlayerPile.Image = card41
                                    playerPile = 2
                                    card(11) = 41
                                Case 42
                                    frmSpeedGame.picPlayerPile.Image = card42
                                    playerPile = 3
                                    card(11) = 42
                                Case 43
                                    frmSpeedGame.picPlayerPile.Image = card43
                                    playerPile = 4
                                    card(11) = 43
                                Case 44
                                    frmSpeedGame.picPlayerPile.Image = card44
                                    playerPile = 5
                                    card(11) = 44
                                Case 45
                                    frmSpeedGame.picPlayerPile.Image = card45
                                    playerPile = 6
                                    card(11) = 45
                                Case 46
                                    frmSpeedGame.picPlayerPile.Image = card46
                                    playerPile = 7
                                    card(11) = 46
                                Case 47
                                    frmSpeedGame.picPlayerPile.Image = card47
                                    playerPile = 8
                                    card(11) = 47
                                Case 48
                                    frmSpeedGame.picPlayerPile.Image = card48
                                    playerPile = 9
                                    card(11) = 48
                                Case 49
                                    frmSpeedGame.picPlayerPile.Image = card49
                                    playerPile = 10
                                    card(11) = 49
                                Case 50
                                    frmSpeedGame.picPlayerPile.Image = card50
                                    playerPile = 11
                                    card(11) = 50
                                Case 51
                                    frmSpeedGame.picPlayerPile.Image = card51
                                    playerPile = 12
                                    card(11) = 51
                                Case 52
                                    frmSpeedGame.picPlayerPile.Image = card52
                                    playerPile = 13
                                    card(11) = 52
                            End Select
                    End Select

                    Used.Add(num)                       'add number to list so that it cannot be generated again

                End If
            Loop While OK = False                       'if card has already been generated, then get another number
        Next counter                                    'generate until all the spaces have been filled


        Me.Hide()                               'hide options form
        frmSpeedHome.Hide()                     'hide main menu form
        frmSpeedGame.Show()                     'display game form

    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Me.Hide()                               'hide options form

        Me.txtName.Text = Nothing               'reset options to default
        Me.radEasy.Checked = True
        Me.radMedium.Checked = False
        Me.radHard.Checked = False
        Me.radBlue.Checked = True
        Me.radRed.Checked = False
        Me.radBlack.Checked = False

    End Sub

    Private Sub picBlueCard_Click(sender As Object, e As EventArgs) Handles picBlueCard.Click
        radBlue.Checked = True                  'allow player to click radio button or card to set option
    End Sub

    Private Sub picRedCard_Click(sender As Object, e As EventArgs) Handles picRedCard.Click
        radRed.Checked = True                   'allow player to click radio button or card to set option
    End Sub

    Private Sub picBlackCard_Click(sender As Object, e As EventArgs) Handles picBlackCard.Click
        radBlack.Checked = True                 'allow player to click radio button or card to set option
    End Sub

    Friend Shared Function btnCancel_Click() As Boolean
        Throw New NotImplementedException()             'allow buttonCancel to be used in other forms
    End Function
End Class
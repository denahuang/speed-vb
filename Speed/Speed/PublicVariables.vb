Module PublicVariables
    Public playerName As String

    Public compEasy As Boolean = False
    Public compMedium As Boolean = False
    Public compHard As Boolean = False

    Public compCardsLeftTotal As Integer = 15
    Public compCardsLeft1 As Integer = 1
    Public compCardsLeft2 As Integer = 2
    Public compCardsLeft3 As Integer = 3
    Public compCardsLeft4 As Integer = 4
    Public compCardsLeft5 As Integer = 5

    Public playerCardsLeftTotal As Integer = 15
    Public playerCardsLeft1 As Integer = 1
    Public playerCardsLeft2 As Integer = 2
    Public playerCardsLeft3 As Integer = 3
    Public playerCardsLeft4 As Integer = 4
    Public playerCardsLeft5 As Integer = 5

    Public Used As New List(Of Integer)

    Public card1 As Bitmap = My.Resources.SpadeA
    Public card2 As Bitmap = My.Resources.Spade2
    Public card3 As Bitmap = My.Resources.Spade3
    Public card4 As Bitmap = My.Resources.Spade4
    Public card5 As Bitmap = My.Resources.Spade5
    Public card6 As Bitmap = My.Resources.Spade6
    Public card7 As Bitmap = My.Resources.Spade7
    Public card8 As Bitmap = My.Resources.Spade8
    Public card9 As Bitmap = My.Resources.Spade9
    Public card10 As Bitmap = My.Resources.Spade10
    Public card11 As Bitmap = My.Resources.SpadeJ
    Public card12 As Bitmap = My.Resources.SpadeQ
    Public card13 As Bitmap = My.Resources.SpadeK
    Public card14 As Bitmap = My.Resources.HeartA
    Public card15 As Bitmap = My.Resources.Heart2
    Public card16 As Bitmap = My.Resources.Heart3
    Public card17 As Bitmap = My.Resources.Heart4
    Public card18 As Bitmap = My.Resources.Heart5
    Public card19 As Bitmap = My.Resources.Heart6
    Public card20 As Bitmap = My.Resources.Heart7
    Public card21 As Bitmap = My.Resources.Heart8
    Public card22 As Bitmap = My.Resources.Heart9
    Public card23 As Bitmap = My.Resources.Heart10
    Public card24 As Bitmap = My.Resources.HeartJ
    Public card25 As Bitmap = My.Resources.HeartQ
    Public card26 As Bitmap = My.Resources.HeartK
    Public card27 As Bitmap = My.Resources.ClubA
    Public card28 As Bitmap = My.Resources.Club2
    Public card29 As Bitmap = My.Resources.Club3
    Public card30 As Bitmap = My.Resources.Club4
    Public card31 As Bitmap = My.Resources.Club5
    Public card32 As Bitmap = My.Resources.Club6
    Public card33 As Bitmap = My.Resources.Club7
    Public card34 As Bitmap = My.Resources.Club8
    Public card35 As Bitmap = My.Resources.Club9
    Public card36 As Bitmap = My.Resources.Club10
    Public card37 As Bitmap = My.Resources.ClubJ
    Public card38 As Bitmap = My.Resources.ClubQ
    Public card39 As Bitmap = My.Resources.ClubK
    Public card40 As Bitmap = My.Resources.DiamondA
    Public card41 As Bitmap = My.Resources.Diamond2
    Public card42 As Bitmap = My.Resources.Diamond3
    Public card43 As Bitmap = My.Resources.Diamond4
    Public card44 As Bitmap = My.Resources.Diamond5
    Public card45 As Bitmap = My.Resources.Diamond6
    Public card46 As Bitmap = My.Resources.Diamond7
    Public card47 As Bitmap = My.Resources.Diamond8
    Public card48 As Bitmap = My.Resources.Diamond9
    Public card49 As Bitmap = My.Resources.Diamond10
    Public card50 As Bitmap = My.Resources.DiamondJ
    Public card51 As Bitmap = My.Resources.DiamondQ
    Public card52 As Bitmap = My.Resources.DiamondK

    Public card(11) As Integer

    Public compCard(4) As Integer
    Public playerCard(4) As Integer
    Public compPile As Integer
    Public playerPile As Integer

    Public cardCheck As Integer

    Public max As Integer = 52
    Public min As Integer = 1


    Public howToPlay As String = "Each player (player and computer) is given 15 random cards, spread out in 5 piles. From left to right, the piles have an increase in number of cards by 1 (1,2,3,4,5). There are two piles in the middle, side by side, with a random card facing up: these are the computer’s (left) and player’s (right) pile. There are two other piles on either side of the two center piles, which contain the remaining cards that will be used as shuffle cards." _
        & vbCrLf & vbCrLf & "Once the player clicks anywhere on the form, the timer will begin and the computer will begin to play his cards. The player will have to place a card that is either one above or one below onto either pile. Ace is considered as a value above King, but below 2, which creates a loop between all the cards. They cannot place a card that is the same value. Once the card has been placed, a new card from the pile is shown. If neither the player or computer can play their cards, the player will have to click his/her shuffle pile, which will simultaneously generate a new random card for both the player and the computer. " _
        & vbCrLf & vbCrLf & "The game ends when one has finished playing all their cards. "
End Module

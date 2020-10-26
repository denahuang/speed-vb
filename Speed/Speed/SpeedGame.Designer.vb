<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmSpeedGame
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmSpeedGame))
        Me.MenuToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.NewGameToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.HelpToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.HowToPlayToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.QuitToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.MenuStrip1 = New System.Windows.Forms.MenuStrip()
        Me.tmrGame = New System.Windows.Forms.Timer(Me.components)
        Me.tmrEasy = New System.Windows.Forms.Timer(Me.components)
        Me.tmrMedium = New System.Windows.Forms.Timer(Me.components)
        Me.tmrHard = New System.Windows.Forms.Timer(Me.components)
        Me.lblCompCardsLeft5 = New System.Windows.Forms.Label()
        Me.lblCompCardsLeft4 = New System.Windows.Forms.Label()
        Me.lblCompCardsLeft3 = New System.Windows.Forms.Label()
        Me.lblCompCardsLeft2 = New System.Windows.Forms.Label()
        Me.lblCompCardsLeft1 = New System.Windows.Forms.Label()
        Me.lblCompCardsLeftTotal = New System.Windows.Forms.Label()
        Me.lblPlayerCardsLeftTotal = New System.Windows.Forms.Label()
        Me.lblPlayerCardsLeft1 = New System.Windows.Forms.Label()
        Me.lblPlayerCardsLeft2 = New System.Windows.Forms.Label()
        Me.lblPlayerCardsLeft3 = New System.Windows.Forms.Label()
        Me.lblPlayerCardsLeft4 = New System.Windows.Forms.Label()
        Me.lblPlayerCardsLeft5 = New System.Windows.Forms.Label()
        Me.picPlayerDeck = New System.Windows.Forms.PictureBox()
        Me.picCompPile = New System.Windows.Forms.PictureBox()
        Me.picCompDeck = New System.Windows.Forms.PictureBox()
        Me.picPlayerPile = New System.Windows.Forms.PictureBox()
        Me.picCompCard1 = New System.Windows.Forms.PictureBox()
        Me.picCompCard5 = New System.Windows.Forms.PictureBox()
        Me.picCompCard4 = New System.Windows.Forms.PictureBox()
        Me.picCompCard3 = New System.Windows.Forms.PictureBox()
        Me.picCompCard2 = New System.Windows.Forms.PictureBox()
        Me.picPlayerCard4 = New System.Windows.Forms.PictureBox()
        Me.picPlayerCard3 = New System.Windows.Forms.PictureBox()
        Me.picPlayerCard1 = New System.Windows.Forms.PictureBox()
        Me.picPlayerCard5 = New System.Windows.Forms.PictureBox()
        Me.picPlayerCard2 = New System.Windows.Forms.PictureBox()
        Me.btnShuffle = New System.Windows.Forms.Button()
        Me.lblPause = New System.Windows.Forms.Label()
        Me.MenuStrip1.SuspendLayout()
        CType(Me.picPlayerDeck, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.picCompPile, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.picCompDeck, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.picPlayerPile, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.picCompCard1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.picCompCard5, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.picCompCard4, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.picCompCard3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.picCompCard2, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.picPlayerCard4, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.picPlayerCard3, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.picPlayerCard1, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.picPlayerCard5, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.picPlayerCard2, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'MenuToolStripMenuItem
        '
        Me.MenuToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.NewGameToolStripMenuItem, Me.HelpToolStripMenuItem, Me.QuitToolStripMenuItem})
        Me.MenuToolStripMenuItem.Font = New System.Drawing.Font("Kristen ITC", 20.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MenuToolStripMenuItem.ForeColor = System.Drawing.Color.MidnightBlue
        Me.MenuToolStripMenuItem.Name = "MenuToolStripMenuItem"
        Me.MenuToolStripMenuItem.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.MenuToolStripMenuItem.Size = New System.Drawing.Size(150, 58)
        Me.MenuToolStripMenuItem.Text = "Menu"
        '
        'NewGameToolStripMenuItem
        '
        Me.NewGameToolStripMenuItem.Font = New System.Drawing.Font("Kristen ITC", 16.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.NewGameToolStripMenuItem.ForeColor = System.Drawing.Color.MidnightBlue
        Me.NewGameToolStripMenuItem.Name = "NewGameToolStripMenuItem"
        Me.NewGameToolStripMenuItem.Size = New System.Drawing.Size(285, 48)
        Me.NewGameToolStripMenuItem.Text = "New Game"
        '
        'HelpToolStripMenuItem
        '
        Me.HelpToolStripMenuItem.DropDownItems.AddRange(New System.Windows.Forms.ToolStripItem() {Me.HowToPlayToolStripMenuItem})
        Me.HelpToolStripMenuItem.Font = New System.Drawing.Font("Kristen ITC", 16.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.HelpToolStripMenuItem.ForeColor = System.Drawing.Color.MidnightBlue
        Me.HelpToolStripMenuItem.Name = "HelpToolStripMenuItem"
        Me.HelpToolStripMenuItem.Size = New System.Drawing.Size(285, 48)
        Me.HelpToolStripMenuItem.Text = "Help"
        '
        'HowToPlayToolStripMenuItem
        '
        Me.HowToPlayToolStripMenuItem.Name = "HowToPlayToolStripMenuItem"
        Me.HowToPlayToolStripMenuItem.Size = New System.Drawing.Size(311, 48)
        Me.HowToPlayToolStripMenuItem.Text = "How To Play"
        '
        'QuitToolStripMenuItem
        '
        Me.QuitToolStripMenuItem.Font = New System.Drawing.Font("Kristen ITC", 16.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.QuitToolStripMenuItem.ForeColor = System.Drawing.Color.MidnightBlue
        Me.QuitToolStripMenuItem.Name = "QuitToolStripMenuItem"
        Me.QuitToolStripMenuItem.Size = New System.Drawing.Size(285, 48)
        Me.QuitToolStripMenuItem.Text = "Quit"
        '
        'MenuStrip1
        '
        Me.MenuStrip1.BackColor = System.Drawing.Color.Transparent
        Me.MenuStrip1.Font = New System.Drawing.Font("Kristen ITC", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.MenuStrip1.ImageScalingSize = New System.Drawing.Size(24, 24)
        Me.MenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.MenuToolStripMenuItem})
        Me.MenuStrip1.Location = New System.Drawing.Point(0, 0)
        Me.MenuStrip1.Name = "MenuStrip1"
        Me.MenuStrip1.Size = New System.Drawing.Size(1237, 62)
        Me.MenuStrip1.TabIndex = 38
        Me.MenuStrip1.Text = "MenuStrip1"
        '
        'tmrGame
        '
        '
        'tmrEasy
        '
        Me.tmrEasy.Enabled = True
        Me.tmrEasy.Interval = 5000
        '
        'tmrMedium
        '
        Me.tmrMedium.Enabled = True
        Me.tmrMedium.Interval = 3500
        '
        'tmrHard
        '
        Me.tmrHard.Enabled = True
        Me.tmrHard.Interval = 2000
        '
        'lblCompCardsLeft5
        '
        Me.lblCompCardsLeft5.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.lblCompCardsLeft5.AutoSize = True
        Me.lblCompCardsLeft5.BackColor = System.Drawing.Color.Transparent
        Me.lblCompCardsLeft5.Font = New System.Drawing.Font("Century Gothic", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCompCardsLeft5.Location = New System.Drawing.Point(158, 99)
        Me.lblCompCardsLeft5.Name = "lblCompCardsLeft5"
        Me.lblCompCardsLeft5.Size = New System.Drawing.Size(0, 30)
        Me.lblCompCardsLeft5.TabIndex = 57
        '
        'lblCompCardsLeft4
        '
        Me.lblCompCardsLeft4.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.lblCompCardsLeft4.AutoSize = True
        Me.lblCompCardsLeft4.BackColor = System.Drawing.Color.Transparent
        Me.lblCompCardsLeft4.Font = New System.Drawing.Font("Century Gothic", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCompCardsLeft4.Location = New System.Drawing.Point(351, 99)
        Me.lblCompCardsLeft4.Name = "lblCompCardsLeft4"
        Me.lblCompCardsLeft4.Size = New System.Drawing.Size(0, 30)
        Me.lblCompCardsLeft4.TabIndex = 58
        '
        'lblCompCardsLeft3
        '
        Me.lblCompCardsLeft3.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.lblCompCardsLeft3.AutoSize = True
        Me.lblCompCardsLeft3.BackColor = System.Drawing.Color.Transparent
        Me.lblCompCardsLeft3.Font = New System.Drawing.Font("Century Gothic", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCompCardsLeft3.Location = New System.Drawing.Point(566, 99)
        Me.lblCompCardsLeft3.Name = "lblCompCardsLeft3"
        Me.lblCompCardsLeft3.Size = New System.Drawing.Size(0, 30)
        Me.lblCompCardsLeft3.TabIndex = 59
        '
        'lblCompCardsLeft2
        '
        Me.lblCompCardsLeft2.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.lblCompCardsLeft2.AutoSize = True
        Me.lblCompCardsLeft2.BackColor = System.Drawing.Color.Transparent
        Me.lblCompCardsLeft2.Font = New System.Drawing.Font("Century Gothic", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCompCardsLeft2.Location = New System.Drawing.Point(765, 99)
        Me.lblCompCardsLeft2.Name = "lblCompCardsLeft2"
        Me.lblCompCardsLeft2.Size = New System.Drawing.Size(0, 30)
        Me.lblCompCardsLeft2.TabIndex = 60
        '
        'lblCompCardsLeft1
        '
        Me.lblCompCardsLeft1.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.lblCompCardsLeft1.AutoSize = True
        Me.lblCompCardsLeft1.BackColor = System.Drawing.Color.Transparent
        Me.lblCompCardsLeft1.Font = New System.Drawing.Font("Century Gothic", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCompCardsLeft1.Location = New System.Drawing.Point(979, 99)
        Me.lblCompCardsLeft1.Name = "lblCompCardsLeft1"
        Me.lblCompCardsLeft1.Size = New System.Drawing.Size(0, 30)
        Me.lblCompCardsLeft1.TabIndex = 61
        '
        'lblCompCardsLeftTotal
        '
        Me.lblCompCardsLeftTotal.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.lblCompCardsLeftTotal.AutoSize = True
        Me.lblCompCardsLeftTotal.BackColor = System.Drawing.Color.Transparent
        Me.lblCompCardsLeftTotal.Font = New System.Drawing.Font("Century Gothic", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCompCardsLeftTotal.Location = New System.Drawing.Point(12, 457)
        Me.lblCompCardsLeftTotal.Name = "lblCompCardsLeftTotal"
        Me.lblCompCardsLeftTotal.Size = New System.Drawing.Size(0, 30)
        Me.lblCompCardsLeftTotal.TabIndex = 62
        '
        'lblPlayerCardsLeftTotal
        '
        Me.lblPlayerCardsLeftTotal.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.lblPlayerCardsLeftTotal.AutoSize = True
        Me.lblPlayerCardsLeftTotal.BackColor = System.Drawing.Color.Transparent
        Me.lblPlayerCardsLeftTotal.Font = New System.Drawing.Font("Century Gothic", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPlayerCardsLeftTotal.Location = New System.Drawing.Point(1037, 457)
        Me.lblPlayerCardsLeftTotal.Name = "lblPlayerCardsLeftTotal"
        Me.lblPlayerCardsLeftTotal.Size = New System.Drawing.Size(0, 30)
        Me.lblPlayerCardsLeftTotal.TabIndex = 63
        '
        'lblPlayerCardsLeft1
        '
        Me.lblPlayerCardsLeft1.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.lblPlayerCardsLeft1.AutoSize = True
        Me.lblPlayerCardsLeft1.BackColor = System.Drawing.Color.Transparent
        Me.lblPlayerCardsLeft1.Font = New System.Drawing.Font("Century Gothic", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPlayerCardsLeft1.Location = New System.Drawing.Point(158, 874)
        Me.lblPlayerCardsLeft1.Name = "lblPlayerCardsLeft1"
        Me.lblPlayerCardsLeft1.Size = New System.Drawing.Size(0, 30)
        Me.lblPlayerCardsLeft1.TabIndex = 64
        '
        'lblPlayerCardsLeft2
        '
        Me.lblPlayerCardsLeft2.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.lblPlayerCardsLeft2.AutoSize = True
        Me.lblPlayerCardsLeft2.BackColor = System.Drawing.Color.Transparent
        Me.lblPlayerCardsLeft2.Font = New System.Drawing.Font("Century Gothic", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPlayerCardsLeft2.Location = New System.Drawing.Point(351, 874)
        Me.lblPlayerCardsLeft2.Name = "lblPlayerCardsLeft2"
        Me.lblPlayerCardsLeft2.Size = New System.Drawing.Size(0, 30)
        Me.lblPlayerCardsLeft2.TabIndex = 65
        '
        'lblPlayerCardsLeft3
        '
        Me.lblPlayerCardsLeft3.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.lblPlayerCardsLeft3.AutoSize = True
        Me.lblPlayerCardsLeft3.BackColor = System.Drawing.Color.Transparent
        Me.lblPlayerCardsLeft3.Font = New System.Drawing.Font("Century Gothic", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPlayerCardsLeft3.Location = New System.Drawing.Point(566, 874)
        Me.lblPlayerCardsLeft3.Name = "lblPlayerCardsLeft3"
        Me.lblPlayerCardsLeft3.Size = New System.Drawing.Size(0, 30)
        Me.lblPlayerCardsLeft3.TabIndex = 66
        '
        'lblPlayerCardsLeft4
        '
        Me.lblPlayerCardsLeft4.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.lblPlayerCardsLeft4.AutoSize = True
        Me.lblPlayerCardsLeft4.BackColor = System.Drawing.Color.Transparent
        Me.lblPlayerCardsLeft4.Font = New System.Drawing.Font("Century Gothic", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPlayerCardsLeft4.Location = New System.Drawing.Point(765, 874)
        Me.lblPlayerCardsLeft4.Name = "lblPlayerCardsLeft4"
        Me.lblPlayerCardsLeft4.Size = New System.Drawing.Size(0, 30)
        Me.lblPlayerCardsLeft4.TabIndex = 67
        '
        'lblPlayerCardsLeft5
        '
        Me.lblPlayerCardsLeft5.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.lblPlayerCardsLeft5.AutoSize = True
        Me.lblPlayerCardsLeft5.BackColor = System.Drawing.Color.Transparent
        Me.lblPlayerCardsLeft5.Font = New System.Drawing.Font("Century Gothic", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPlayerCardsLeft5.Location = New System.Drawing.Point(979, 874)
        Me.lblPlayerCardsLeft5.Name = "lblPlayerCardsLeft5"
        Me.lblPlayerCardsLeft5.Size = New System.Drawing.Size(0, 30)
        Me.lblPlayerCardsLeft5.TabIndex = 68
        '
        'picPlayerDeck
        '
        Me.picPlayerDeck.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.picPlayerDeck.BackColor = System.Drawing.Color.Transparent
        Me.picPlayerDeck.Image = CType(resources.GetObject("picPlayerDeck.Image"), System.Drawing.Image)
        Me.picPlayerDeck.Location = New System.Drawing.Point(855, 385)
        Me.picPlayerDeck.Name = "picPlayerDeck"
        Me.picPlayerDeck.Size = New System.Drawing.Size(160, 228)
        Me.picPlayerDeck.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.picPlayerDeck.TabIndex = 54
        Me.picPlayerDeck.TabStop = False
        '
        'picCompPile
        '
        Me.picCompPile.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.picCompPile.BackColor = System.Drawing.Color.Transparent
        Me.picCompPile.Image = CType(resources.GetObject("picCompPile.Image"), System.Drawing.Image)
        Me.picCompPile.Location = New System.Drawing.Point(441, 385)
        Me.picCompPile.Name = "picCompPile"
        Me.picCompPile.Size = New System.Drawing.Size(167, 228)
        Me.picCompPile.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.picCompPile.TabIndex = 53
        Me.picCompPile.TabStop = False
        '
        'picCompDeck
        '
        Me.picCompDeck.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.picCompDeck.BackColor = System.Drawing.Color.Transparent
        Me.picCompDeck.Image = CType(resources.GetObject("picCompDeck.Image"), System.Drawing.Image)
        Me.picCompDeck.Location = New System.Drawing.Point(241, 385)
        Me.picCompDeck.Name = "picCompDeck"
        Me.picCompDeck.Size = New System.Drawing.Size(161, 228)
        Me.picCompDeck.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.picCompDeck.TabIndex = 52
        Me.picCompDeck.TabStop = False
        '
        'picPlayerPile
        '
        Me.picPlayerPile.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.picPlayerPile.BackColor = System.Drawing.Color.Transparent
        Me.picPlayerPile.Image = CType(resources.GetObject("picPlayerPile.Image"), System.Drawing.Image)
        Me.picPlayerPile.Location = New System.Drawing.Point(647, 385)
        Me.picPlayerPile.Name = "picPlayerPile"
        Me.picPlayerPile.Size = New System.Drawing.Size(167, 228)
        Me.picPlayerPile.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.picPlayerPile.TabIndex = 51
        Me.picPlayerPile.TabStop = False
        '
        'picCompCard1
        '
        Me.picCompCard1.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.picCompCard1.BackColor = System.Drawing.Color.Transparent
        Me.picCompCard1.Image = Global.Speed.My.Resources.Resources.CardBackBlue
        Me.picCompCard1.Location = New System.Drawing.Point(964, 137)
        Me.picCompCard1.Name = "picCompCard1"
        Me.picCompCard1.Size = New System.Drawing.Size(164, 223)
        Me.picCompCard1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.picCompCard1.TabIndex = 50
        Me.picCompCard1.TabStop = False
        '
        'picCompCard5
        '
        Me.picCompCard5.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.picCompCard5.BackColor = System.Drawing.Color.Transparent
        Me.picCompCard5.Image = CType(resources.GetObject("picCompCard5.Image"), System.Drawing.Image)
        Me.picCompCard5.Location = New System.Drawing.Point(137, 137)
        Me.picCompCard5.Name = "picCompCard5"
        Me.picCompCard5.Size = New System.Drawing.Size(158, 223)
        Me.picCompCard5.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.picCompCard5.TabIndex = 49
        Me.picCompCard5.TabStop = False
        '
        'picCompCard4
        '
        Me.picCompCard4.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.picCompCard4.BackColor = System.Drawing.Color.Transparent
        Me.picCompCard4.Image = CType(resources.GetObject("picCompCard4.Image"), System.Drawing.Image)
        Me.picCompCard4.Location = New System.Drawing.Point(338, 137)
        Me.picCompCard4.Name = "picCompCard4"
        Me.picCompCard4.Size = New System.Drawing.Size(161, 223)
        Me.picCompCard4.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.picCompCard4.TabIndex = 48
        Me.picCompCard4.TabStop = False
        '
        'picCompCard3
        '
        Me.picCompCard3.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.picCompCard3.BackColor = System.Drawing.Color.Transparent
        Me.picCompCard3.Image = CType(resources.GetObject("picCompCard3.Image"), System.Drawing.Image)
        Me.picCompCard3.Location = New System.Drawing.Point(553, 137)
        Me.picCompCard3.Name = "picCompCard3"
        Me.picCompCard3.Size = New System.Drawing.Size(155, 223)
        Me.picCompCard3.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.picCompCard3.TabIndex = 47
        Me.picCompCard3.TabStop = False
        '
        'picCompCard2
        '
        Me.picCompCard2.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.picCompCard2.BackColor = System.Drawing.Color.Transparent
        Me.picCompCard2.Image = CType(resources.GetObject("picCompCard2.Image"), System.Drawing.Image)
        Me.picCompCard2.Location = New System.Drawing.Point(751, 137)
        Me.picCompCard2.Name = "picCompCard2"
        Me.picCompCard2.Size = New System.Drawing.Size(164, 223)
        Me.picCompCard2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.picCompCard2.TabIndex = 46
        Me.picCompCard2.TabStop = False
        '
        'picPlayerCard4
        '
        Me.picPlayerCard4.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.picPlayerCard4.BackColor = System.Drawing.Color.Transparent
        Me.picPlayerCard4.Image = CType(resources.GetObject("picPlayerCard4.Image"), System.Drawing.Image)
        Me.picPlayerCard4.Location = New System.Drawing.Point(753, 641)
        Me.picPlayerCard4.Name = "picPlayerCard4"
        Me.picPlayerCard4.Size = New System.Drawing.Size(162, 221)
        Me.picPlayerCard4.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.picPlayerCard4.TabIndex = 45
        Me.picPlayerCard4.TabStop = False
        '
        'picPlayerCard3
        '
        Me.picPlayerCard3.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.picPlayerCard3.BackColor = System.Drawing.Color.Transparent
        Me.picPlayerCard3.Image = CType(resources.GetObject("picPlayerCard3.Image"), System.Drawing.Image)
        Me.picPlayerCard3.Location = New System.Drawing.Point(554, 641)
        Me.picPlayerCard3.Name = "picPlayerCard3"
        Me.picPlayerCard3.Size = New System.Drawing.Size(155, 221)
        Me.picPlayerCard3.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.picPlayerCard3.TabIndex = 44
        Me.picPlayerCard3.TabStop = False
        '
        'picPlayerCard1
        '
        Me.picPlayerCard1.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.picPlayerCard1.BackColor = System.Drawing.Color.Transparent
        Me.picPlayerCard1.Image = CType(resources.GetObject("picPlayerCard1.Image"), System.Drawing.Image)
        Me.picPlayerCard1.Location = New System.Drawing.Point(137, 641)
        Me.picPlayerCard1.Name = "picPlayerCard1"
        Me.picPlayerCard1.Size = New System.Drawing.Size(158, 221)
        Me.picPlayerCard1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.picPlayerCard1.TabIndex = 42
        Me.picPlayerCard1.TabStop = False
        '
        'picPlayerCard5
        '
        Me.picPlayerCard5.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.picPlayerCard5.BackColor = System.Drawing.Color.Transparent
        Me.picPlayerCard5.Image = CType(resources.GetObject("picPlayerCard5.Image"), System.Drawing.Image)
        Me.picPlayerCard5.Location = New System.Drawing.Point(967, 641)
        Me.picPlayerCard5.Name = "picPlayerCard5"
        Me.picPlayerCard5.Size = New System.Drawing.Size(161, 221)
        Me.picPlayerCard5.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.picPlayerCard5.TabIndex = 41
        Me.picPlayerCard5.TabStop = False
        '
        'picPlayerCard2
        '
        Me.picPlayerCard2.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.picPlayerCard2.BackColor = System.Drawing.Color.Transparent
        Me.picPlayerCard2.Image = CType(resources.GetObject("picPlayerCard2.Image"), System.Drawing.Image)
        Me.picPlayerCard2.Location = New System.Drawing.Point(339, 641)
        Me.picPlayerCard2.Name = "picPlayerCard2"
        Me.picPlayerCard2.Size = New System.Drawing.Size(160, 221)
        Me.picPlayerCard2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.picPlayerCard2.TabIndex = 43
        Me.picPlayerCard2.TabStop = False
        '
        'btnShuffle
        '
        Me.btnShuffle.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.btnShuffle.BackColor = System.Drawing.Color.White
        Me.btnShuffle.Font = New System.Drawing.Font("Century Gothic", 10.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btnShuffle.Location = New System.Drawing.Point(884, 474)
        Me.btnShuffle.Name = "btnShuffle"
        Me.btnShuffle.Size = New System.Drawing.Size(101, 47)
        Me.btnShuffle.TabIndex = 69
        Me.btnShuffle.Text = " Shuffle"
        Me.btnShuffle.UseVisualStyleBackColor = False
        '
        'lblPause
        '
        Me.lblPause.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lblPause.BackColor = System.Drawing.Color.Transparent
        Me.lblPause.Font = New System.Drawing.Font("Kristen ITC", 16.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPause.ForeColor = System.Drawing.Color.MidnightBlue
        Me.lblPause.Location = New System.Drawing.Point(696, 23)
        Me.lblPause.Name = "lblPause"
        Me.lblPause.Size = New System.Drawing.Size(508, 111)
        Me.lblPause.TabIndex = 83
        Me.lblPause.Text = "Press P to Pause"
        Me.lblPause.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'frmSpeedGame
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackgroundImage = Global.Speed.My.Resources.Resources.wood_background
        Me.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.ClientSize = New System.Drawing.Size(1237, 955)
        Me.Controls.Add(Me.lblPause)
        Me.Controls.Add(Me.btnShuffle)
        Me.Controls.Add(Me.lblPlayerCardsLeft5)
        Me.Controls.Add(Me.lblPlayerCardsLeft4)
        Me.Controls.Add(Me.lblPlayerCardsLeft3)
        Me.Controls.Add(Me.lblPlayerCardsLeft2)
        Me.Controls.Add(Me.lblPlayerCardsLeft1)
        Me.Controls.Add(Me.lblPlayerCardsLeftTotal)
        Me.Controls.Add(Me.lblCompCardsLeftTotal)
        Me.Controls.Add(Me.lblCompCardsLeft1)
        Me.Controls.Add(Me.lblCompCardsLeft2)
        Me.Controls.Add(Me.lblCompCardsLeft3)
        Me.Controls.Add(Me.lblCompCardsLeft4)
        Me.Controls.Add(Me.lblCompCardsLeft5)
        Me.Controls.Add(Me.picPlayerDeck)
        Me.Controls.Add(Me.picCompPile)
        Me.Controls.Add(Me.picCompDeck)
        Me.Controls.Add(Me.picPlayerPile)
        Me.Controls.Add(Me.picCompCard1)
        Me.Controls.Add(Me.picCompCard5)
        Me.Controls.Add(Me.picCompCard4)
        Me.Controls.Add(Me.picCompCard3)
        Me.Controls.Add(Me.picCompCard2)
        Me.Controls.Add(Me.picPlayerCard4)
        Me.Controls.Add(Me.picPlayerCard3)
        Me.Controls.Add(Me.picPlayerCard1)
        Me.Controls.Add(Me.picPlayerCard5)
        Me.Controls.Add(Me.picPlayerCard2)
        Me.Controls.Add(Me.MenuStrip1)
        Me.KeyPreview = True
        Me.Name = "frmSpeedGame"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "|"
        Me.WindowState = System.Windows.Forms.FormWindowState.Maximized
        Me.MenuStrip1.ResumeLayout(False)
        Me.MenuStrip1.PerformLayout()
        CType(Me.picPlayerDeck, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.picCompPile, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.picCompDeck, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.picPlayerPile, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.picCompCard1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.picCompCard5, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.picCompCard4, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.picCompCard3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.picCompCard2, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.picPlayerCard4, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.picPlayerCard3, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.picPlayerCard1, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.picPlayerCard5, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.picPlayerCard2, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents picPlayerDeck As PictureBox
    Friend WithEvents picCompPile As PictureBox
    Friend WithEvents picCompDeck As PictureBox
    Friend WithEvents picPlayerPile As PictureBox
    Friend WithEvents picCompCard1 As PictureBox
    Friend WithEvents picCompCard5 As PictureBox
    Friend WithEvents picCompCard4 As PictureBox
    Friend WithEvents picCompCard3 As PictureBox
    Friend WithEvents picCompCard2 As PictureBox
    Friend WithEvents MenuToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents NewGameToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents HelpToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents HowToPlayToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents QuitToolStripMenuItem As ToolStripMenuItem
    Friend WithEvents picPlayerCard4 As PictureBox
    Friend WithEvents picPlayerCard3 As PictureBox
    Friend WithEvents picPlayerCard1 As PictureBox
    Friend WithEvents picPlayerCard5 As PictureBox
    Friend WithEvents picPlayerCard2 As PictureBox
    Friend WithEvents MenuStrip1 As MenuStrip
    Friend WithEvents tmrGame As Timer
    Friend WithEvents tmrEasy As Timer
    Friend WithEvents tmrMedium As Timer
    Friend WithEvents tmrHard As Timer
    Friend WithEvents lblCompCardsLeft5 As Label
    Friend WithEvents lblCompCardsLeft4 As Label
    Friend WithEvents lblCompCardsLeft3 As Label
    Friend WithEvents lblCompCardsLeft2 As Label
    Friend WithEvents lblCompCardsLeft1 As Label
    Friend WithEvents lblCompCardsLeftTotal As Label
    Friend WithEvents lblPlayerCardsLeftTotal As Label
    Friend WithEvents lblPlayerCardsLeft1 As Label
    Friend WithEvents lblPlayerCardsLeft2 As Label
    Friend WithEvents lblPlayerCardsLeft3 As Label
    Friend WithEvents lblPlayerCardsLeft4 As Label
    Friend WithEvents lblPlayerCardsLeft5 As Label
    Friend WithEvents btnShuffle As Button
    Friend WithEvents lblPause As Label
End Class

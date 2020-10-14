<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmNewGameOptions
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmNewGameOptions))
        Me.btnStart = New System.Windows.Forms.Button()
        Me.grpCardColor = New System.Windows.Forms.GroupBox()
        Me.picBlackCard = New System.Windows.Forms.PictureBox()
        Me.picRedCard = New System.Windows.Forms.PictureBox()
        Me.picBlueCard = New System.Windows.Forms.PictureBox()
        Me.radBlack = New System.Windows.Forms.RadioButton()
        Me.radRed = New System.Windows.Forms.RadioButton()
        Me.radBlue = New System.Windows.Forms.RadioButton()
        Me.grpDifficulty = New System.Windows.Forms.GroupBox()
        Me.radHard = New System.Windows.Forms.RadioButton()
        Me.radMedium = New System.Windows.Forms.RadioButton()
        Me.radEasy = New System.Windows.Forms.RadioButton()
        Me.txtName = New System.Windows.Forms.TextBox()
        Me.lblNamePrompt = New System.Windows.Forms.Label()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.grpCardColor.SuspendLayout()
        CType(Me.picBlackCard, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.picRedCard, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.picBlueCard, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpDifficulty.SuspendLayout()
        Me.SuspendLayout()
        '
        'btnStart
        '
        Me.btnStart.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.btnStart.Location = New System.Drawing.Point(385, 15)
        Me.btnStart.Name = "btnStart"
        Me.btnStart.Size = New System.Drawing.Size(103, 34)
        Me.btnStart.TabIndex = 15
        Me.btnStart.Text = "Start"
        Me.btnStart.UseVisualStyleBackColor = True
        '
        'grpCardColor
        '
        Me.grpCardColor.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.grpCardColor.Controls.Add(Me.picBlackCard)
        Me.grpCardColor.Controls.Add(Me.picRedCard)
        Me.grpCardColor.Controls.Add(Me.picBlueCard)
        Me.grpCardColor.Controls.Add(Me.radBlack)
        Me.grpCardColor.Controls.Add(Me.radRed)
        Me.grpCardColor.Controls.Add(Me.radBlue)
        Me.grpCardColor.Location = New System.Drawing.Point(29, 191)
        Me.grpCardColor.Name = "grpCardColor"
        Me.grpCardColor.Size = New System.Drawing.Size(459, 254)
        Me.grpCardColor.TabIndex = 14
        Me.grpCardColor.TabStop = False
        Me.grpCardColor.Text = "Select a Card Color:"
        '
        'picBlackCard
        '
        Me.picBlackCard.Image = CType(resources.GetObject("picBlackCard.Image"), System.Drawing.Image)
        Me.picBlackCard.Location = New System.Drawing.Point(313, 56)
        Me.picBlackCard.Name = "picBlackCard"
        Me.picBlackCard.Size = New System.Drawing.Size(128, 192)
        Me.picBlackCard.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.picBlackCard.TabIndex = 5
        Me.picBlackCard.TabStop = False
        '
        'picRedCard
        '
        Me.picRedCard.Image = CType(resources.GetObject("picRedCard.Image"), System.Drawing.Image)
        Me.picRedCard.Location = New System.Drawing.Point(168, 56)
        Me.picRedCard.Name = "picRedCard"
        Me.picRedCard.Size = New System.Drawing.Size(128, 192)
        Me.picRedCard.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.picRedCard.TabIndex = 4
        Me.picRedCard.TabStop = False
        '
        'picBlueCard
        '
        Me.picBlueCard.Image = CType(resources.GetObject("picBlueCard.Image"), System.Drawing.Image)
        Me.picBlueCard.Location = New System.Drawing.Point(21, 56)
        Me.picBlueCard.Name = "picBlueCard"
        Me.picBlueCard.Size = New System.Drawing.Size(128, 192)
        Me.picBlueCard.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.picBlueCard.TabIndex = 3
        Me.picBlueCard.TabStop = False
        '
        'radBlack
        '
        Me.radBlack.AutoSize = True
        Me.radBlack.Location = New System.Drawing.Point(313, 26)
        Me.radBlack.Name = "radBlack"
        Me.radBlack.Size = New System.Drawing.Size(73, 24)
        Me.radBlack.TabIndex = 2
        Me.radBlack.Text = "Black"
        Me.radBlack.UseVisualStyleBackColor = True
        '
        'radRed
        '
        Me.radRed.AutoSize = True
        Me.radRed.Location = New System.Drawing.Point(168, 26)
        Me.radRed.Name = "radRed"
        Me.radRed.Size = New System.Drawing.Size(64, 24)
        Me.radRed.TabIndex = 1
        Me.radRed.Text = "Red"
        Me.radRed.UseVisualStyleBackColor = True
        '
        'radBlue
        '
        Me.radBlue.AutoSize = True
        Me.radBlue.Checked = True
        Me.radBlue.Location = New System.Drawing.Point(21, 26)
        Me.radBlue.Name = "radBlue"
        Me.radBlue.Size = New System.Drawing.Size(66, 24)
        Me.radBlue.TabIndex = 0
        Me.radBlue.TabStop = True
        Me.radBlue.Text = "Blue"
        Me.radBlue.UseVisualStyleBackColor = True
        '
        'grpDifficulty
        '
        Me.grpDifficulty.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.grpDifficulty.Controls.Add(Me.radHard)
        Me.grpDifficulty.Controls.Add(Me.radMedium)
        Me.grpDifficulty.Controls.Add(Me.radEasy)
        Me.grpDifficulty.Location = New System.Drawing.Point(29, 89)
        Me.grpDifficulty.Name = "grpDifficulty"
        Me.grpDifficulty.Size = New System.Drawing.Size(459, 84)
        Me.grpDifficulty.TabIndex = 13
        Me.grpDifficulty.TabStop = False
        Me.grpDifficulty.Text = "Select a Difficulty:"
        '
        'radHard
        '
        Me.radHard.AutoSize = True
        Me.radHard.Location = New System.Drawing.Point(313, 38)
        Me.radHard.Name = "radHard"
        Me.radHard.Size = New System.Drawing.Size(69, 24)
        Me.radHard.TabIndex = 2
        Me.radHard.Text = "Hard"
        Me.radHard.UseVisualStyleBackColor = True
        '
        'radMedium
        '
        Me.radMedium.AutoSize = True
        Me.radMedium.Location = New System.Drawing.Point(168, 38)
        Me.radMedium.Name = "radMedium"
        Me.radMedium.Size = New System.Drawing.Size(90, 24)
        Me.radMedium.TabIndex = 1
        Me.radMedium.Text = "Medium"
        Me.radMedium.UseVisualStyleBackColor = True
        '
        'radEasy
        '
        Me.radEasy.AutoSize = True
        Me.radEasy.Checked = True
        Me.radEasy.Location = New System.Drawing.Point(21, 38)
        Me.radEasy.Name = "radEasy"
        Me.radEasy.Size = New System.Drawing.Size(69, 24)
        Me.radEasy.TabIndex = 0
        Me.radEasy.TabStop = True
        Me.radEasy.Text = "Easy"
        Me.radEasy.UseVisualStyleBackColor = True
        '
        'txtName
        '
        Me.txtName.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.txtName.Location = New System.Drawing.Point(29, 45)
        Me.txtName.Name = "txtName"
        Me.txtName.Size = New System.Drawing.Size(324, 26)
        Me.txtName.TabIndex = 12
        '
        'lblNamePrompt
        '
        Me.lblNamePrompt.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.lblNamePrompt.AutoSize = True
        Me.lblNamePrompt.Location = New System.Drawing.Point(25, 22)
        Me.lblNamePrompt.Name = "lblNamePrompt"
        Me.lblNamePrompt.Size = New System.Drawing.Size(130, 20)
        Me.lblNamePrompt.TabIndex = 11
        Me.lblNamePrompt.Text = "Enter your name:"
        '
        'btnCancel
        '
        Me.btnCancel.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.btnCancel.Location = New System.Drawing.Point(385, 55)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(103, 34)
        Me.btnCancel.TabIndex = 16
        Me.btnCancel.Text = "Cancel"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'frmNewGameOptions
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(9.0!, 20.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.SystemColors.Control
        Me.ClientSize = New System.Drawing.Size(519, 458)
        Me.Controls.Add(Me.btnCancel)
        Me.Controls.Add(Me.btnStart)
        Me.Controls.Add(Me.grpCardColor)
        Me.Controls.Add(Me.grpDifficulty)
        Me.Controls.Add(Me.txtName)
        Me.Controls.Add(Me.lblNamePrompt)
        Me.ForeColor = System.Drawing.SystemColors.ControlText
        Me.KeyPreview = True
        Me.Name = "frmNewGameOptions"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "New Game: Options"
        Me.grpCardColor.ResumeLayout(False)
        Me.grpCardColor.PerformLayout()
        CType(Me.picBlackCard, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.picRedCard, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.picBlueCard, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpDifficulty.ResumeLayout(False)
        Me.grpDifficulty.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents btnStart As Button
    Friend WithEvents grpCardColor As GroupBox
    Friend WithEvents picBlackCard As PictureBox
    Friend WithEvents picRedCard As PictureBox
    Friend WithEvents picBlueCard As PictureBox
    Friend WithEvents radBlack As RadioButton
    Friend WithEvents radRed As RadioButton
    Friend WithEvents radBlue As RadioButton
    Friend WithEvents grpDifficulty As GroupBox
    Friend WithEvents radHard As RadioButton
    Friend WithEvents radMedium As RadioButton
    Friend WithEvents radEasy As RadioButton
    Friend WithEvents txtName As TextBox
    Friend WithEvents lblNamePrompt As Label
    Friend WithEvents btnCancel As Button
End Class

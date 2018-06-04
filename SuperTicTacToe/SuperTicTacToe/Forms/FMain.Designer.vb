<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FMain
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

#Region "f"
    Private WithEvents chkAiGoesFirst As System.Windows.Forms.CheckBox
    Private WithEvents btnNewGame As System.Windows.Forms.Button
    Private WithEvents bgbBigBoard As CBigBoard



#End Region

    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FMain))
        Me.chkAiGoesFirst = New System.Windows.Forms.CheckBox()
        Me.btnNewGame = New System.Windows.Forms.Button()
        Me.btnLoadGame = New System.Windows.Forms.Button()
        Me.rdoEasy = New System.Windows.Forms.RadioButton()
        Me.rdoHard = New System.Windows.Forms.RadioButton()
        Me.btnExit = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.chkAddToDatabase = New System.Windows.Forms.CheckBox()
        Me.rdoHuman = New System.Windows.Forms.RadioButton()
        Me.bgbBigBoard = New SuperTicTacToe.CBigBoard()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'chkAiGoesFirst
        '
        Me.chkAiGoesFirst.AutoSize = True
        Me.chkAiGoesFirst.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.chkAiGoesFirst.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.chkAiGoesFirst.ForeColor = System.Drawing.Color.Black
        Me.chkAiGoesFirst.Location = New System.Drawing.Point(238, 22)
        Me.chkAiGoesFirst.Name = "chkAiGoesFirst"
        Me.chkAiGoesFirst.Size = New System.Drawing.Size(81, 17)
        Me.chkAiGoesFirst.TabIndex = 2
        Me.chkAiGoesFirst.Tag = ""
        Me.chkAiGoesFirst.Text = "AI goes first"
        Me.chkAiGoesFirst.UseVisualStyleBackColor = False
        '
        'btnNewGame
        '
        Me.btnNewGame.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.btnNewGame.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.btnNewGame.ForeColor = System.Drawing.Color.Black
        Me.btnNewGame.Location = New System.Drawing.Point(6, 22)
        Me.btnNewGame.Name = "btnNewGame"
        Me.btnNewGame.Size = New System.Drawing.Size(70, 35)
        Me.btnNewGame.TabIndex = 3
        Me.btnNewGame.Tag = ""
        Me.btnNewGame.Text = "New Game"
        Me.btnNewGame.UseVisualStyleBackColor = False
        '
        'btnLoadGame
        '
        Me.btnLoadGame.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.btnLoadGame.Location = New System.Drawing.Point(82, 21)
        Me.btnLoadGame.Name = "btnLoadGame"
        Me.btnLoadGame.Size = New System.Drawing.Size(70, 35)
        Me.btnLoadGame.TabIndex = 1
        Me.btnLoadGame.Tag = ""
        Me.btnLoadGame.Text = "Load Game"
        Me.btnLoadGame.UseVisualStyleBackColor = False
        '
        'rdoEasy
        '
        Me.rdoEasy.AutoSize = True
        Me.rdoEasy.Location = New System.Drawing.Point(410, 26)
        Me.rdoEasy.Name = "rdoEasy"
        Me.rdoEasy.Size = New System.Drawing.Size(48, 17)
        Me.rdoEasy.TabIndex = 4
        Me.rdoEasy.Text = "Easy"
        Me.rdoEasy.UseVisualStyleBackColor = True
        '
        'rdoHard
        '
        Me.rdoHard.AutoSize = True
        Me.rdoHard.Checked = True
        Me.rdoHard.Location = New System.Drawing.Point(410, 43)
        Me.rdoHard.Name = "rdoHard"
        Me.rdoHard.Size = New System.Drawing.Size(48, 17)
        Me.rdoHard.TabIndex = 5
        Me.rdoHard.TabStop = True
        Me.rdoHard.Text = "Hard"
        Me.rdoHard.UseVisualStyleBackColor = True
        '
        'btnExit
        '
        Me.btnExit.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.btnExit.Location = New System.Drawing.Point(159, 21)
        Me.btnExit.Name = "btnExit"
        Me.btnExit.Size = New System.Drawing.Size(70, 35)
        Me.btnExit.TabIndex = 6
        Me.btnExit.Tag = ""
        Me.btnExit.Text = "Exit"
        Me.btnExit.UseVisualStyleBackColor = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(354, 28)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(50, 13)
        Me.Label1.TabIndex = 7
        Me.Label1.Text = "Difficulty:"
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.chkAddToDatabase)
        Me.GroupBox1.Controls.Add(Me.btnExit)
        Me.GroupBox1.Controls.Add(Me.btnNewGame)
        Me.GroupBox1.Controls.Add(Me.rdoHuman)
        Me.GroupBox1.Controls.Add(Me.chkAiGoesFirst)
        Me.GroupBox1.Controls.Add(Me.btnLoadGame)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.rdoEasy)
        Me.GroupBox1.Controls.Add(Me.rdoHard)
        Me.GroupBox1.Location = New System.Drawing.Point(6, 557)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(537, 62)
        Me.GroupBox1.TabIndex = 1
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Settings"
        '
        'chkAddToDatabase
        '
        Me.chkAddToDatabase.AutoSize = True
        Me.chkAddToDatabase.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.chkAddToDatabase.Checked = True
        Me.chkAddToDatabase.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkAddToDatabase.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!)
        Me.chkAddToDatabase.ForeColor = System.Drawing.Color.Black
        Me.chkAddToDatabase.Location = New System.Drawing.Point(238, 38)
        Me.chkAddToDatabase.Name = "chkAddToDatabase"
        Me.chkAddToDatabase.Size = New System.Drawing.Size(110, 17)
        Me.chkAddToDatabase.TabIndex = 9
        Me.chkAddToDatabase.Tag = ""
        Me.chkAddToDatabase.Text = "Add To Database"
        Me.chkAddToDatabase.UseVisualStyleBackColor = False
        '
        'rdoHuman
        '
        Me.rdoHuman.AutoSize = True
        Me.rdoHuman.Location = New System.Drawing.Point(410, 10)
        Me.rdoHuman.Name = "rdoHuman"
        Me.rdoHuman.Size = New System.Drawing.Size(92, 17)
        Me.rdoHuman.TabIndex = 8
        Me.rdoHuman.Text = "Human(No AI)"
        Me.rdoHuman.UseVisualStyleBackColor = True
        '
        'bgbBigBoard
        '
        Me.bgbBigBoard.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.bgbBigBoard.ImeMode = System.Windows.Forms.ImeMode.NoControl
        Me.bgbBigBoard.Location = New System.Drawing.Point(6, 0)
        Me.bgbBigBoard.Name = "bgbBigBoard"
        Me.bgbBigBoard.Size = New System.Drawing.Size(537, 551)
        Me.bgbBigBoard.TabIndex = 0
        Me.bgbBigBoard.TabStop = False
        Me.bgbBigBoard.Tag = "0"
        '
        'FMain
        '
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(549, 626)
        Me.Controls.Add(Me.GroupBox1)
        Me.Controls.Add(Me.bgbBigBoard)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.Name = "FMain"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Ultimate Tic Tac Toe"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents btnLoadGame As System.Windows.Forms.Button
    Friend WithEvents rdoEasy As System.Windows.Forms.RadioButton
    Friend WithEvents rdoHard As System.Windows.Forms.RadioButton
    Friend WithEvents btnExit As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents rdoHuman As System.Windows.Forms.RadioButton
    Private WithEvents chkAddToDatabase As System.Windows.Forms.CheckBox
End Class

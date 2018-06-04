<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FLoadGame
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
#End Region

    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.VCompletedGamesBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.DbUltimateTicTacToeDataSetBindingSource = New System.Windows.Forms.BindingSource(Me.components)
        Me.DbUltimateTicTacToeDataSet = New SuperTicTacToe.dbUltimateTicTacToeDataSet()
        Me.dgvGames = New System.Windows.Forms.DataGridView()
        Me.VCompletedGamesBindingSource1 = New System.Windows.Forms.BindingSource(Me.components)
        Me.Label1 = New System.Windows.Forms.Label()
        Me.VCompletedGamesTableAdapter = New SuperTicTacToe.dbUltimateTicTacToeDataSetTableAdapters.VCompletedGamesTableAdapter()
        Me.btnSelect = New System.Windows.Forms.Button()
        Me.IntGameIDDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.StrPlayerNameDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        Me.IntMoveCountDataGridViewTextBoxColumn = New System.Windows.Forms.DataGridViewTextBoxColumn()
        CType(Me.VCompletedGamesBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DbUltimateTicTacToeDataSetBindingSource, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.DbUltimateTicTacToeDataSet, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.dgvGames, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.VCompletedGamesBindingSource1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'VCompletedGamesBindingSource
        '
        Me.VCompletedGamesBindingSource.DataMember = "VCompletedGames"
        Me.VCompletedGamesBindingSource.DataSource = Me.DbUltimateTicTacToeDataSetBindingSource
        '
        'DbUltimateTicTacToeDataSetBindingSource
        '
        Me.DbUltimateTicTacToeDataSetBindingSource.DataSource = Me.DbUltimateTicTacToeDataSet
        Me.DbUltimateTicTacToeDataSetBindingSource.Position = 0
        '
        'DbUltimateTicTacToeDataSet
        '
        Me.DbUltimateTicTacToeDataSet.DataSetName = "dbUltimateTicTacToeDataSet"
        Me.DbUltimateTicTacToeDataSet.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema
        '
        'dgvGames
        '
        Me.dgvGames.AllowUserToAddRows = False
        Me.dgvGames.AllowUserToDeleteRows = False
        Me.dgvGames.AutoGenerateColumns = False
        Me.dgvGames.BackgroundColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.dgvGames.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dgvGames.Columns.AddRange(New System.Windows.Forms.DataGridViewColumn() {Me.IntGameIDDataGridViewTextBoxColumn, Me.StrPlayerNameDataGridViewTextBoxColumn, Me.IntMoveCountDataGridViewTextBoxColumn})
        Me.dgvGames.DataSource = Me.VCompletedGamesBindingSource
        Me.dgvGames.Location = New System.Drawing.Point(14, 39)
        Me.dgvGames.Name = "dgvGames"
        Me.dgvGames.ReadOnly = True
        Me.dgvGames.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvGames.Size = New System.Drawing.Size(344, 208)
        Me.dgvGames.TabIndex = 1
        '
        'VCompletedGamesBindingSource1
        '
        Me.VCompletedGamesBindingSource1.DataMember = "VCompletedGames"
        Me.VCompletedGamesBindingSource1.DataSource = Me.DbUltimateTicTacToeDataSetBindingSource
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(48, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(278, 20)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Select the game you would like to load"
        '
        'VCompletedGamesTableAdapter
        '
        Me.VCompletedGamesTableAdapter.ClearBeforeFill = True
        '
        'btnSelect
        '
        Me.btnSelect.Location = New System.Drawing.Point(152, 253)
        Me.btnSelect.Name = "btnSelect"
        Me.btnSelect.Size = New System.Drawing.Size(75, 23)
        Me.btnSelect.TabIndex = 3
        Me.btnSelect.Text = "Select"
        Me.btnSelect.UseVisualStyleBackColor = True
        '
        'IntGameIDDataGridViewTextBoxColumn
        '
        Me.IntGameIDDataGridViewTextBoxColumn.DataPropertyName = "intGameID"
        Me.IntGameIDDataGridViewTextBoxColumn.HeaderText = "Game ID"
        Me.IntGameIDDataGridViewTextBoxColumn.Name = "IntGameIDDataGridViewTextBoxColumn"
        Me.IntGameIDDataGridViewTextBoxColumn.ReadOnly = True
        '
        'StrPlayerNameDataGridViewTextBoxColumn
        '
        Me.StrPlayerNameDataGridViewTextBoxColumn.DataPropertyName = "strPlayerName"
        Me.StrPlayerNameDataGridViewTextBoxColumn.HeaderText = "Player Name"
        Me.StrPlayerNameDataGridViewTextBoxColumn.Name = "StrPlayerNameDataGridViewTextBoxColumn"
        Me.StrPlayerNameDataGridViewTextBoxColumn.ReadOnly = True
        '
        'IntMoveCountDataGridViewTextBoxColumn
        '
        Me.IntMoveCountDataGridViewTextBoxColumn.DataPropertyName = "intMoveCount"
        Me.IntMoveCountDataGridViewTextBoxColumn.HeaderText = "Move Count"
        Me.IntMoveCountDataGridViewTextBoxColumn.Name = "IntMoveCountDataGridViewTextBoxColumn"
        Me.IntMoveCountDataGridViewTextBoxColumn.ReadOnly = True
        '
        'FLoadGame
        '
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(192, Byte), Integer), CType(CType(192, Byte), Integer), CType(CType(255, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(373, 281)
        Me.Controls.Add(Me.btnSelect)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.dgvGames)
        Me.Name = "FLoadGame"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Load Game"
        CType(Me.VCompletedGamesBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DbUltimateTicTacToeDataSetBindingSource, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.DbUltimateTicTacToeDataSet, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.dgvGames, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.VCompletedGamesBindingSource1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents DbUltimateTicTacToeDataSetBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents DbUltimateTicTacToeDataSet As SuperTicTacToe.dbUltimateTicTacToeDataSet
    Friend WithEvents VCompletedGamesBindingSource As System.Windows.Forms.BindingSource
    Friend WithEvents VCompletedGamesTableAdapter As SuperTicTacToe.dbUltimateTicTacToeDataSetTableAdapters.VCompletedGamesTableAdapter
    Friend WithEvents dgvGames As System.Windows.Forms.DataGridView
    Friend WithEvents VCompletedGamesBindingSource1 As System.Windows.Forms.BindingSource
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents btnSelect As System.Windows.Forms.Button
    Friend WithEvents IntGameIDDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents StrPlayerNameDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
    Friend WithEvents IntMoveCountDataGridViewTextBoxColumn As System.Windows.Forms.DataGridViewTextBoxColumn
End Class


' -------------------------------------------------------------------------
' Name: Stephen Ulm
' Class: Capstone
' Abstract: Tic Tac Toe with 3 levels of AI
' -------------------------------------------------------------------------

Imports System.Threading
Partial Public Class FMain

    Private m_clsGame As CGame
    Private m_blnIsGameWon As Boolean = False

    ' -------------------------------------------------------------------------
    ' Name: FMain_Load
    ' Abstract: Opens the database connection
    ' -------------------------------------------------------------------------
    Private Sub FMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            modDatabaseUtilities.OpenDatabaseConnectionSQLServer()
        Catch excError As Exception

            MessageBox.Show("Could not connect to database, locking the connect to database option and the load game button")

            chkAddToDatabase.Checked = False
            chkAddToDatabase.Enabled = False
            btnLoadGame.Enabled = False

            WriteLog(excError, False)
        End Try


    End Sub



    ' -------------------------------------------------------------------------
    ' Name: FMain_FormClosed
    ' Abstract: Close the database connection
    ' -------------------------------------------------------------------------
    Private Sub FMain_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed

        Try
            modDatabaseUtilities.CloseDatabaseConnection()
        Catch excError As Exception
            WriteLog(excError)
        End Try

    End Sub

    ' -------------------------------------------------------------------------
    ' Name: bgbBigBoard_GameIsOver
    ' Abstract: Tells the game to make the end game button a new game button
    ' -------------------------------------------------------------------------
    Private Sub bgbBigBoard_GameIsOver(sender As Object, e As EventArgs) Handles bgbBigBoard.GameIsOver

        btnNewGame.Text = "New game"

    End Sub


    ' -------------------------------------------------------------------------
    ' Name: btnNewGame_Click
    ' Abstract: As a new game button prompts for a name and if it gets one starts a new game and
    ' becomes an end game button, which asks if the user is sure and if they are resets the board and
    ' becomes a new game button again
    ' -------------------------------------------------------------------------
    Private Sub btnNewGame_Click(sender As Object, e As EventArgs) Handles btnNewGame.Click

        Dim strName As String
        Dim enuDifficulty As enuDifficulties
        Dim frmInputBox As New FInputBox
        Dim dgrEndGame As DialogResult
        Dim blnAddToDatabase As Boolean

        If btnNewGame.Text = "End Game" Then

            dgrEndGame = MessageBox.Show("Are you sure you would like to end the game?", "End game?", MessageBoxButtons.YesNo)
            If dgrEndGame = Windows.Forms.DialogResult.Yes Then
                bgbBigBoard.ResetBoard()
                btnNewGame.Text = "New Game"
            End If


        Else
            frmInputBox.ShowDialog()

            If frmInputBox.WasCanceled = False Then

                strName = frmInputBox.GetInput()

                m_blnIsGameWon = False

                blnAddToDatabase = chkAddToDatabase.Checked

                If rdoHuman.Checked = True Then
                    enuDifficulty = enuDifficulties.ReplayOrHuman
                    blnAddToDatabase = False
                ElseIf rdoEasy.Checked = True Then
                    enuDifficulty = enuDifficulties.Easy
                ElseIf rdoHard.Checked = True Then
                    enuDifficulty = enuDifficulties.Hard
                End If

                btnNewGame.Text = "End Game"

                bgbBigBoard.StartGame(strName, chkAiGoesFirst.Checked, enuDifficulty, blnAddToDatabase)

            End If
        End If
    End Sub



    ' -------------------------------------------------------------------------
    ' Name: btnLoadGame_Click
    ' Abstract: Brings up the window to pick which game to load, then loads it
    ' -------------------------------------------------------------------------
    Private Sub btnLoadGame_Click(sender As Object, e As EventArgs) Handles btnLoadGame.Click
        Dim frmLoadGame As New FLoadGame
        Dim intGameID As Integer

        intGameID = frmLoadGame.ShowDialog()

        If intGameID <> -1 Then
            'Replay game
            bgbBigBoard.ReplayGame(intGameID)
        End If

    End Sub



    ' -------------------------------------------------------------------------
    ' Name: btnExit_Click
    ' Abstract: Closes the window
    ' -------------------------------------------------------------------------
    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click

        Me.Close()

    End Sub

    ' -------------------------------------------------------------------------
    ' Name: rdoHuman_CheckedChanged
    ' Abstract: Warns the player that add to database is disabled for human vs human games
    ' -------------------------------------------------------------------------
    Private Sub rdoHuman_CheckedChanged(sender As Object, e As EventArgs) Handles rdoHuman.CheckedChanged

        MessageBox.Show("Warning:  Even if ""Add to Database"" is checked, adding to the database is disabled for human vs human games.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Warning)

    End Sub

End Class
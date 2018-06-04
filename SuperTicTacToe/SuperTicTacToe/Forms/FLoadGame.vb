
' -------------------------------------------------------------------------
' Name: Stephen Ulm
' Class: Capstone
' Abstract: Tic Tac Toe with 3 levels of AI
' -------------------------------------------------------------------------

Imports System.Threading
Partial Public Class FLoadGame
    Private m_intGameID = -1
    Private m_intReturnedGameID = -1

    Public Overloads Function ShowDialog() As Integer
        MyBase.ShowDialog()
        Return m_intReturnedGameID
    End Function



    Private Sub FLoadGame_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        'This line of code loads data into the 'DbUltimateTicTacToeDataSet.VCompletedGames' table
        Me.VCompletedGamesTableAdapter.ClearBeforeFill = True
        Me.VCompletedGamesTableAdapter.Fill(Me.DbUltimateTicTacToeDataSet.VCompletedGames)

        If dgvGames.Rows.Count > 0 Then

            m_intGameID = dgvGames.Rows(0).Cells(0).Value
        End If

    End Sub


    Private Sub lstGames_Format(sender As Object, e As ListControlConvertEventArgs)


        Dim strName As String = e.ListItem(1)
        Dim intMoveCount As String = e.ListItem(3)

        e.Value = "Name: " + strName + " Move count: " + intMoveCount

    End Sub

    Private Sub DataGridView1_CellMouseClick(sender As Object, e As DataGridViewCellMouseEventArgs) Handles dgvGames.CellMouseClick

        m_intGameID = dgvGames.Rows(e.RowIndex).Cells(0).Value 'Game ID

    End Sub

    Private Sub btnSelect_Click(sender As Object, e As EventArgs) Handles btnSelect.Click
        m_intReturnedGameID = m_intGameID
        Me.Close()
    End Sub
End Class

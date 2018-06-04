
' -------------------------------------------------------------------------
' Name: Stephen Ulm
' Class: Capstone
' Abstract: Tic Tac Toe with 3 levels of AI
' -------------------------------------------------------------------------

Imports System.Threading
Partial Public Class FInputBox
    Private m_blnOkayWasClicked As Boolean = False
    Private m_strUserInput As String


    Private Sub FInputBox_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        m_blnOkayWasClicked = False
        m_strUserInput = ""
        txtUserInput.Text = ""
    End Sub

    Private Sub btnOkay_Click(sender As Object, e As EventArgs) Handles btnOkay.Click
        m_strUserInput = txtUserInput.Text
        m_blnOkayWasClicked = True
        Me.Close()
    End Sub

    Public Function WasCanceled() As String
        Return Not m_blnOkayWasClicked
    End Function

    Public Function GetInput() As String
        Return m_strUserInput
    End Function

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Me.Close()
    End Sub

End Class

Public Class CGame
    Private m_intGameID
    Private m_strName As String
    Private m_blnAIGoesFirst As Boolean
    Private m_enuDifficulty As enuDifficulties
    Private m_enuGameOutcome As enuGameOutcomes
    Private m_dtmPlayed As DateTime

    Private m_blnXTurn As Boolean
    Private m_blnIsComputerTurn As Boolean
    Private m_blnIsGameOver As Boolean
    Private m_blnAddToDatabase As Boolean

    Private m_intMoveIndex As Integer
    Public plrHuman As CPlayer
    Public plrComputer As CPlayer

    Public ReadOnly m_clrHuman As Color = System.Drawing.Color.FromArgb(128, 255, 128)
    Public ReadOnly m_clrComputer As Color = System.Drawing.Color.FromArgb(128, 255, 255)



    ' -------------------------------------------------------------------------
    ' Name: New
    ' Abstract: Constructor
    ' -------------------------------------------------------------------------
    Public Sub New(strName As String, blnAiGoesFirst As Boolean, enuDifficulty As enuDifficulties, blnAddToDatabase As Boolean)

        Initialize(strName, blnAiGoesFirst, enuDifficulty, blnAddToDatabase)

    End Sub



    ' -------------------------------------------------------------------------
    ' Name: Initialize
    ' Abstract: Initializes all the values
    ' -------------------------------------------------------------------------
    Public Sub Initialize(strName As String, blnAIGoesFirst As Boolean, enuDifficulty As enuDifficulties, blnConnectToDatabase As Boolean)

        SetName(strName)

        m_blnAIGoesFirst = blnAIGoesFirst

        m_enuDifficulty = enuDifficulty

        m_enuGameOutcome = enuGameOutcomes.GameEndedEarly

        m_dtmPlayed = DateTime.Now

        m_blnXTurn = True
        m_blnIsGameOver = False

        m_blnIsComputerTurn = blnAIGoesFirst

        m_intMoveIndex = 1

        m_blnAddToDatabase = blnConnectToDatabase

        If m_blnAddToDatabase = True Then

            m_intGameID = modDatabaseUtilities.StartGame(strName, blnAIGoesFirst, m_enuDifficulty, m_enuGameOutcome, m_dtmPlayed)

        End If

        If blnAIGoesFirst = True Then
            plrComputer = New CPlayer("X", m_clrComputer, m_clrHuman)
            plrHuman = New CPlayer("O", m_clrHuman, m_clrComputer)
        Else

            plrComputer = New CPlayer("O", m_clrComputer, m_clrHuman)
            plrHuman = New CPlayer("X", m_clrHuman, m_clrComputer)
        End If



    End Sub




    ' -------------------------------------------------------------------------
    ' Name: SetName
    ' Abstract: Sets the name
    ' -------------------------------------------------------------------------
    Public Sub SetName(strName As String)

        If strName.Length > 50 Then
            strName.Remove(50)
        End If

        m_strName = strName

    End Sub



    ' -------------------------------------------------------------------------
    ' Name: GetName
    ' Abstract: Returns the name
    ' -------------------------------------------------------------------------
    Public Function GetName() As String

        Return m_strName

    End Function



    ' -------------------------------------------------------------------------
    ' Name: GetDifficulty
    ' Abstract: Returns the difficulty
    ' -------------------------------------------------------------------------
    Public Function GetDifficulty() As enuDifficulties

        Return m_enuDifficulty

    End Function



    ' -------------------------------------------------------------------------
    ' Name: IsAiFirst() 
    ' Abstract: Returns if the ai went first
    ' -------------------------------------------------------------------------
    Public Function IsAiFirst() As Boolean

        Return m_blnAIGoesFirst

    End Function



    ' -------------------------------------------------------------------------
    ' Name: IsXTurn
    ' Abstract: Returns if it's x's turn
    ' -------------------------------------------------------------------------
    Public Function IsXTurn() As Boolean

        Return m_blnXTurn

    End Function



    ' -------------------------------------------------------------------------
    ' Name: IsComputerTurn
    ' Abstract: Returns if it's computer's turn
    ' -------------------------------------------------------------------------
    Public Function IsComputerTurn() As Boolean

        Return m_blnIsComputerTurn

    End Function



    ' -------------------------------------------------------------------------
    ' Name: IsGameOver
    ' Abstract: Returns if the game is over
    ' -------------------------------------------------------------------------
    Public Function IsGameOver() As Boolean

        Return m_blnIsGameOver

    End Function



    ' -------------------------------------------------------------------------
    ' Name: DoMove
    ' Abstract: Does a move
    ' -------------------------------------------------------------------------
    Public Sub DoMove(gspSender As CGameSpace)

        Dim lcbParent As CLocalBoard

        If m_blnAddToDatabase = True Then

            lcbParent = gspSender.Parent

            modDatabaseUtilities.EnterMove(m_intGameID, m_intMoveIndex, lcbParent.GetBoardID, gspSender.GetGameSpaceID)

            m_intMoveIndex += 1

        End If
        m_blnXTurn = Not m_blnXTurn
        m_blnIsComputerTurn = Not m_blnIsComputerTurn

    End Sub



    ' -------------------------------------------------------------------------
    ' Name: SetOutcome
    ' Abstract: Sets the outcome
    ' -------------------------------------------------------------------------
    Public Sub SetOutcome(intGameOutcome As enuGameOutcomes)

        If m_blnAddToDatabase = True Then
            modDatabaseUtilities.SetGameOutcome(m_intGameID, intGameOutcome)

            m_blnIsGameOver = True
        End If
    End Sub



    ' -------------------------------------------------------------------------
    ' Name: GetCurrentPlayer
    ' Abstract: Gets the current player
    ' -------------------------------------------------------------------------
    Public Function GetCurrentPlayer() As CPlayer

        If IsComputerTurn() = True Then
            Return plrComputer
        Else
            Return plrHuman
        End If

    End Function

End Class

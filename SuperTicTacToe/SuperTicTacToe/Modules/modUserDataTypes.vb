' ----------------------------------------------------------------------
' Name: modUserDataTypes
' Abstract: Suitcases for traveling with data
' ----------------------------------------------------------------------

' ----------------------------------------------------------------------
' Options
' ----------------------------------------------------------------------
Option Explicit On



Public Module modUserDataTypes


    Public Enum enuGameOutcomes

        PlayerWon = 1
        ComputerWon = 2
        Draw = 3
        GameEndedEarly = 4

    End Enum


    Public Enum enuBoardStates

        PlayerWon = 1
        ComputerWon = 2
        Draw = 3
        StillPlayable = 4

    End Enum

    Public Enum enuDifficulties

        Easy = 1
        Hard = 2
        ReplayOrHuman = 500

    End Enum

    Public Enum enuMoveRanks

        WinningMove = 1
        BlockingMove = 2
        AmazingMove = 3
        GreatMove = 4
        GoodMove = 5
        OkayMove = 6
        NeutralMove = 7
        UnintelligentMove = 8
        BadMove = 9
        TerribleMove = 10
        HorrendousMove = 11
        Taken = 12

    End Enum

    Public Enum enuOutcomesOfSendingToWonBoard

        HumanWinsTheGame = 1
        HumanWinsMyGameWinningBoard = 2
        HumanBlocksMyGameWinningBoard = 2
        HumanWinsABoard = 3
        HumanBlocksMyBoardWin = 4
        NothingHappens = 5

    End Enum

    Public Structure udtGameReplays
        Dim intGameID As Integer
        Dim strPlayerName As String
        Dim blnComputerMovesFirst As Boolean
        Dim intMoveIndex As Integer
        Dim lintMoves As List(Of Integer())
    End Structure

End Module

Imports System.Threading
Public Class CBigBoard

    Inherits CBoard(Of CLocalBoard)
    Private WithEvents m_lcbGroup0 As New CLocalBoard()
    Private WithEvents m_lcbGroup1 As New CLocalBoard()
    Private WithEvents m_lcbGroup2 As New CLocalBoard()
    Private WithEvents m_lcbGroup3 As New CLocalBoard()
    Private WithEvents m_lcbGroup4 As New CLocalBoard()
    Private WithEvents m_lcbGroup5 As New CLocalBoard()
    Private WithEvents m_lcbGroup6 As New CLocalBoard()
    Private WithEvents m_lcbGroup7 As New CLocalBoard()
    Private WithEvents m_lcbGroup8 As New CLocalBoard()
    Public Event GameIsOver(bgbMe As CBigBoard, e As EventArgs)
    Private m_blnIsGameOver As Boolean = False



    ' -------------------------------------------------------------------------
    ' Name: New
    ' Abstract: Initializes the class
    ' -------------------------------------------------------------------------
    Sub New()
        Me.m_lcbGroup0.Location = New System.Drawing.Point(9, 13)
        Me.m_lcbGroup0.Name = "m_lcbGroup0"

        Me.m_lcbGroup1.Location = New System.Drawing.Point(184, 13)
        Me.m_lcbGroup1.Name = "m_lcbGroup1"

        Me.m_lcbGroup2.Location = New System.Drawing.Point(359, 13)
        Me.m_lcbGroup2.Name = "m_lcbGroup2"

        Me.m_lcbGroup3.Location = New System.Drawing.Point(9, 190)
        Me.m_lcbGroup3.Name = "m_lcbGroup3"

        Me.m_lcbGroup4.Location = New System.Drawing.Point(184, 190)
        Me.m_lcbGroup4.Name = "m_lcbGroup4"

        Me.m_lcbGroup5.Location = New System.Drawing.Point(359, 190)
        Me.m_lcbGroup5.Name = "m_lcbGroup5"

        Me.m_lcbGroup6.Location = New System.Drawing.Point(9, 367)
        Me.m_lcbGroup6.Name = "m_lcbGroup6"

        Me.m_lcbGroup7.Location = New System.Drawing.Point(184, 367)
        Me.m_lcbGroup7.Name = "m_lcbGroup7"

        Me.m_lcbGroup8.Location = New System.Drawing.Point(359, 367)
        Me.m_lcbGroup8.Name = "m_lcbGroup8"

        Me.Controls.Add(m_lcbGroup0)
        Me.Controls.Add(m_lcbGroup1)
        Me.Controls.Add(m_lcbGroup2)
        Me.Controls.Add(m_lcbGroup3)
        Me.Controls.Add(m_lcbGroup4)
        Me.Controls.Add(m_lcbGroup5)
        Me.Controls.Add(m_lcbGroup6)
        Me.Controls.Add(m_lcbGroup7)
        Me.Controls.Add(m_lcbGroup8)

        ResetBoard()

        AddHandler m_lcbGroup0.BoardIsWon, AddressOf gen0_SpaceIsWon
        AddHandler m_lcbGroup1.BoardIsWon, AddressOf gen1_SpaceIsWon
        AddHandler m_lcbGroup2.BoardIsWon, AddressOf gen2_SpaceIsWon
        AddHandler m_lcbGroup3.BoardIsWon, AddressOf gen3_SpaceIsWon
        AddHandler m_lcbGroup4.BoardIsWon, AddressOf gen4_SpaceIsWon
        AddHandler m_lcbGroup5.BoardIsWon, AddressOf gen5_SpaceIsWon
        AddHandler m_lcbGroup6.BoardIsWon, AddressOf gen6_SpaceIsWon
        AddHandler m_lcbGroup7.BoardIsWon, AddressOf gen7_SpaceIsWon
        AddHandler m_lcbGroup8.BoardIsWon, AddressOf gen8_SpaceIsWon
        AddHandler Me.BoardIsWon, AddressOf Me_BoardIsWon


        For Each lcb As CLocalBoard In Me.Controls
            lcb.BackColor = System.Drawing.Color.FromArgb(192, 192, 255)
            lcb.ForeColor = System.Drawing.Color.Black
            lcb.Size = New System.Drawing.Size(169, 171)
            lcb.TabIndex = 0
            lcb.TabStop = False
            lcb.Tag = "0"
            m_lgenMoves.Add(lcb)
        Next

        Me.Invalidate()
        Me.Refresh()

    End Sub



    ' -------------------------------------------------------------------------
    ' Name: Me_BoardIsWon
    ' Abstract: The game is won
    ' -------------------------------------------------------------------------
    Private Sub Me_BoardIsWon(bgbme As CBigBoard, e As EventArgs)
        If m_clsGame.IsComputerTurn = True Then
            SetGameOutcome(enuGameOutcomes.ComputerWon)
        Else
            SetGameOutcome(enuGameOutcomes.PlayerWon)
        End If
    End Sub



    ' -------------------------------------------------------------------------
    ' Name: SetGame
    ' Abstract: Sets the game for the board
    ' -------------------------------------------------------------------------
    Public Overloads Sub SetGame(clsGame As CGame)

        m_clsGame = clsGame

        ResetBoard()

    End Sub



    ' -------------------------------------------------------------------------
    ' Name: gspGeneric_GameSpaceEnter
    ' Abstract: Highlights the board you're sending them to
    ' -------------------------------------------------------------------------
    Public Sub gspGeneric_GameSpaceEnter(gspTarget As CGameSpace, e As EventArgs) Handles m_lcbGroup0.GameSpaceEnter, m_lcbGroup1.GameSpaceEnter, m_lcbGroup2.GameSpaceEnter, m_lcbGroup3.GameSpaceEnter, m_lcbGroup4.GameSpaceEnter, m_lcbGroup5.GameSpaceEnter, m_lcbGroup6.GameSpaceEnter, m_lcbGroup7.GameSpaceEnter, m_lcbGroup8.GameSpaceEnter

        Dim lcbTarget As CLocalBoard
        ' Dim sbrTransparentLightGreen As New SolidBrush(Color.FromArgb(80, 50, 255, 50))
        Dim sbrTransparentLightBlue As New SolidBrush(Color.FromArgb(80, 50, 50, 255))

        lcbTarget = CType(Me.Controls(gspTarget.GetGameSpaceID), CLocalBoard)

        lcbTarget.CreateGraphics.FillRectangle(sbrTransparentLightBlue, 2, 2, lcbTarget.Size.Width - 5, lcbTarget.Size.Height - 5)

    End Sub



    ' -------------------------------------------------------------------------
    ' Name: gspGeneric_GameSpaceLeave
    ' Abstract: Highlights the board you're sending them to
    ' -------------------------------------------------------------------------
    Public Sub gspGeneric_GameSpaceLeave(gspTarget As CGameSpace, e As EventArgs) Handles m_lcbGroup0.GameSpaceLeave, m_lcbGroup1.GameSpaceLeave, m_lcbGroup2.GameSpaceLeave, m_lcbGroup3.GameSpaceLeave, m_lcbGroup4.GameSpaceLeave, m_lcbGroup5.GameSpaceLeave, m_lcbGroup6.GameSpaceLeave, m_lcbGroup7.GameSpaceLeave, m_lcbGroup8.GameSpaceLeave

        Me.Invalidate()
        Me.Refresh()

    End Sub


    ' -------------------------------------------------------------------------
    ' Name: StartGame
    ' Abstract: Starts a game
    ' -------------------------------------------------------------------------
    Public Sub StartGame(strName As String, blnAiGoesFirst As Boolean, enuDifficulty As enuDifficulties, blnAddToDatabase As Boolean)

        m_clsGame = New CGame(strName, blnAiGoesFirst, enuDifficulty, blnAddToDatabase)
        ResetBoard()

        For Each clsBoard As CLocalBoard In Me.Controls

            clsBoard.SetGame(m_clsGame)

        Next

        If blnAiGoesFirst = True Then

            'Insert into the top left board
            DoAITurn(m_lcbGroup0)

        End If

        m_blnIsGameOver = False

    End Sub



    ' -------------------------------------------------------------------------
    ' Name: btnGeneric_SpaceIsWon()
    ' Abstract: Puts an x or an o into a box
    ' -------------------------------------------------------------------------
    Private Sub gspGeneric_SpaceIsWon(gspSender As CGameSpace, e As EventArgs) Handles m_lcbGroup0.GameSpaceClick, m_lcbGroup1.GameSpaceClick, m_lcbGroup2.GameSpaceClick, m_lcbGroup3.GameSpaceClick, m_lcbGroup4.GameSpaceClick, m_lcbGroup5.GameSpaceClick, m_lcbGroup6.GameSpaceClick, m_lcbGroup7.GameSpaceClick, m_lcbGroup8.GameSpaceClick

        Dim lcbNext As CLocalBoard

        lcbNext = InsertXOrO(gspSender)

        If m_blnIsGameOver = False OrElse m_clsGame.GetDifficulty = enuDifficulties.ReplayOrHuman Then

            If (m_clsGame.IsComputerTurn = True AndAlso m_clsGame.GetDifficulty <> enuDifficulties.ReplayOrHuman) Then
                DoAITurn(lcbNext)
            End If

        End If

    End Sub




    ' -------------------------------------------------------------------------
    ' Name: ResetBoard
    ' Abstract: Resets the board
    ' -------------------------------------------------------------------------
    Public Sub ResetBoard()

        m_lcbGroup0.SetRank(enuMoveRanks.OkayMove, True)
        m_lcbGroup1.SetRank(enuMoveRanks.NeutralMove, True)
        m_lcbGroup2.SetRank(enuMoveRanks.OkayMove, True)
        m_lcbGroup3.SetRank(enuMoveRanks.NeutralMove, True)
        m_lcbGroup4.SetRank(enuMoveRanks.NeutralMove, True)
        m_lcbGroup5.SetRank(enuMoveRanks.NeutralMove, True)
        m_lcbGroup6.SetRank(enuMoveRanks.OkayMove, True)
        m_lcbGroup7.SetRank(enuMoveRanks.NeutralMove, True)
        m_lcbGroup8.SetRank(enuMoveRanks.OkayMove, True)

        m_lcbGroup0.SetHumanRank(enuMoveRanks.OkayMove, True)
        m_lcbGroup1.SetHumanRank(enuMoveRanks.NeutralMove, True)
        m_lcbGroup2.SetHumanRank(enuMoveRanks.OkayMove, True)
        m_lcbGroup3.SetHumanRank(enuMoveRanks.NeutralMove, True)
        m_lcbGroup4.SetHumanRank(enuMoveRanks.NeutralMove, True)
        m_lcbGroup5.SetHumanRank(enuMoveRanks.NeutralMove, True)
        m_lcbGroup6.SetHumanRank(enuMoveRanks.OkayMove, True)
        m_lcbGroup7.SetHumanRank(enuMoveRanks.NeutralMove, True)
        m_lcbGroup8.SetHumanRank(enuMoveRanks.OkayMove, True)

        For Each lcb As CLocalBoard In Me.Controls

            lcb.ResetBoard()

        Next

        Me.Enabled = True
        Me.BackColor = m_clrEnabled
        Me.Text = ""
        Me.m_enuBoardState = enuBoardStates.StillPlayable

        Me.Invalidate()
        Me.Refresh()

    End Sub



    ' -------------------------------------------------------------------------
    ' Name: ReplayGame
    ' Abstract: Replays an old game
    ' -------------------------------------------------------------------------
    Public Sub ReplayGame(intGameID As Integer)

        Dim udtGameReplay As udtGameReplays
        Dim gspTarget As CGameSpace

        ResetBoard()

        udtGameReplay = modDatabaseUtilities.GetGameReplayInformation(intGameID)

        With udtGameReplay

            'Start a game
            m_clsGame = New CGame(.strPlayerName, .blnComputerMovesFirst, enuDifficulties.ReplayOrHuman, False)

            For Each clsBoard As CLocalBoard In Me.Controls

                clsBoard.SetGame(m_clsGame)

            Next

            'Go through the move list
            For intIndex As Integer = 0 To .lintMoves.Count - 1

                'Get the game space
                gspTarget = GetGameSpaceFromBoardSquarePair(.lintMoves, intIndex)

                'Click it
                gspTarget.PerformClick()

                'Wait 1/10 of a second
                Thread.Sleep(100)
            Next

            m_blnIsGameOver = False

        End With

    End Sub



    ' -------------------------------------------------------------------------
    ' Name: GetGameSpaceFromBoardSquarePair
    ' Abstract: Gets the game space from the board and the square
    ' -------------------------------------------------------------------------
    Public Function GetGameSpaceFromBoardSquarePair(lintMoves As List(Of Integer()), intMoveIndex As Integer) As CGameSpace
        Dim intGroupID As Integer
        Dim intSquareID As Integer
        Dim lcbTarget As CLocalBoard
        Dim gspTarget As CGameSpace

        intGroupID = lintMoves(intMoveIndex)(0)
        intSquareID = lintMoves(intMoveIndex)(1)

        lcbTarget = CType(Me.Controls(intGroupID), CLocalBoard)
        gspTarget = CType(lcbTarget.Controls(intSquareID), CGameSpace)

        Return gspTarget

    End Function



    ' -------------------------------------------------------------------------
    ' Name: DoAITurn
    ' Abstract: Does the AI's turn
    ' -------------------------------------------------------------------------
    Private Sub DoAITurn(grpTarget As CLocalBoard)

        If m_clsGame.IsGameOver = False AndAlso m_clsGame.GetDifficulty <> enuDifficulties.ReplayOrHuman Then

            If m_clsGame.GetDifficulty = enuDifficulties.Hard Then

                DoHardAITurn(grpTarget)

            Else
                DoEasyAITurn(grpTarget)

            End If
        End If
    End Sub



    ' -------------------------------------------------------------------------
    ' Name: DoEasyAITurn
    ' Abstract: Plays randomly for the easy AI
    ' -------------------------------------------------------------------------
    Private Sub DoEasyAITurn(lcbTarget As CLocalBoard)
        Dim rnd As New Random()
        Dim gspTarget As CGameSpace
        Dim lgspMoves As List(Of CGameSpace)
        Dim intMoveIndex As Integer

        'Get all moves
        lgspMoves = lcbTarget.GetAllNotTakenMoves()

        'Is the board full or won?
        If (lgspMoves.Count = 0) Then

            'Yes, get all moves from all boards
            lgspMoves = GetAllMovesFromAllBoards(m_clsGame.plrComputer)
        End If

        'Are all boards full or won?
        If (lgspMoves.Count = 0) Then
            SetGameOutcome(enuGameOutcomes.Draw)
        End If

        intMoveIndex = rnd.Next(0, lgspMoves.Count)

        'Play one of the spaces
        gspTarget = lgspMoves(intMoveIndex)

        gspTarget.PerformClick()
    End Sub



    ' -------------------------------------------------------------------------
    ' Name: GetHighestRankInList
    ' Abstract: Gets the Highest Rank In the List
    ' -------------------------------------------------------------------------
    Private Function GetHighestRankedItemsInList(lgspList As List(Of CGameSpace)) As List(Of CGameSpace)

        Dim enuBestRank As enuMoveRanks = enuMoveRanks.Taken
        Dim lgspTargets As List(Of CGameSpace)

        'Get the highest rank
        For Each genMove In lgspList

            If genMove.GetRank < enuBestRank Then

                enuBestRank = genMove.GetRank

            End If

        Next
        lgspTargets = GetItemsInListOfRank(lgspList, enuBestRank)

        Return lgspTargets

    End Function



    ' -------------------------------------------------------------------------
    ' Name: GetItemsInListOfRank
    ' Abstract: Gets items in the list of the designated rank
    ' -------------------------------------------------------------------------
    Public Function GetItemsInListOfRank(lgspList As List(Of CGameSpace), enuTargetRank As enuMoveRanks) As List(Of CGameSpace)

        Dim lgspTargetsOfRank As New List(Of CGameSpace)

        For Each gspMove In lgspList

            If gspMove.GetRank = enuTargetRank Then

                lgspTargetsOfRank.Add(gspMove)

            End If

        Next

        Return lgspTargetsOfRank

    End Function



    ' -------------------------------------------------------------------------
    ' Name: UpgradeToHigherOutcome
    ' Abstract: Changes to the new outcome if it's higher
    ' -------------------------------------------------------------------------
    Private Sub UpgradeToHigherOutcome(ByRef enuOldOutcome As enuOutcomesOfSendingToWonBoard, ByVal enuNewOutcome As enuOutcomesOfSendingToWonBoard)

        If enuNewOutcome < enuOldOutcome Then

            enuOldOutcome = enuNewOutcome
        End If

    End Sub



    ' -------------------------------------------------------------------------
    ' Name: DoHardAITurn
    ' Abstract: Tactically plays for the hard AI
    ' -------------------------------------------------------------------------
    Private Sub DoHardAITurn(lcbCurrentBoard As CLocalBoard)
        Dim rnd As New Random()
        Dim lgspMoves As List(Of CGameSpace)
        Dim intMoveIndex As Integer
        Dim enuOutcomeOfSendingToWonBoard As enuOutcomesOfSendingToWonBoard = enuOutcomesOfSendingToWonBoard.NothingHappens

        lgspMoves = lcbCurrentBoard.GetAllNotTakenMoves()

        enuOutcomeOfSendingToWonBoard = GetResultOfSendingToWonBoard()

        lgspMoves = FilterMoves(lgspMoves, lcbCurrentBoard, enuOutcomeOfSendingToWonBoard)

        'Are there moves left?
        If lgspMoves.Count > 0 Then

            'Get the best ones
            lgspMoves = GetHighestRankedItemsInList(lgspMoves)

        Else

            lgspMoves.AddRange(lcbCurrentBoard.GetBestMoves())

            If lgspMoves.Count = 0 Then

                lgspMoves = GetBestMovesFromAllBoards(m_clsGame.plrComputer)

                'Filter the moves from all boards
                lgspMoves = FilterMoves(lgspMoves, lcbCurrentBoard, enuOutcomeOfSendingToWonBoard)

            End If

        End If

        intMoveIndex = rnd.Next(0, lgspMoves.Count - 1)

        'Play one of the best spaces
        lgspMoves(intMoveIndex).PerformClick()

    End Sub



    ' -------------------------------------------------------------------------
    ' Name: GetResultOfSendingToWonBoard
    ' Abstract: Gets what happens if the human is sent to a won or filled board
    ' -------------------------------------------------------------------------
    Private Function GetResultOfSendingToWonBoard() As enuOutcomesOfSendingToWonBoard

        Dim enuOutcomeOfSendingToWonBoard As enuOutcomesOfSendingToWonBoard

        For Each lcbBoard As CLocalBoard In Me.Controls

            'Does this board win the game?
            If lcbBoard.GetHumanRank() = enuMoveRanks.WinningMove Then

                'Yes, Can they win this board if sent there?
                If lcbBoard.GetBestMoveRankForHuman = enuMoveRanks.WinningMove Then

                    'Yes, stop looking, if you send to a taken board you lose
                    UpgradeToHigherOutcome(enuOutcomeOfSendingToWonBoard, enuOutcomesOfSendingToWonBoard.HumanWinsTheGame)
                    Exit For

                End If

                'No, does this board block my win?
            ElseIf lcbBoard.GetHumanRank() = enuMoveRanks.BlockingMove Then

                'Can they win this board?
                If lcbBoard.GetBestMoveRankForHuman = enuMoveRanks.WinningMove Then

                    UpgradeToHigherOutcome(enuOutcomeOfSendingToWonBoard, enuOutcomesOfSendingToWonBoard.HumanWinsMyGameWinningBoard)

                    'Can they block me on this board?
                ElseIf lcbBoard.GetBestMoveRankForHuman = enuMoveRanks.BlockingMove Then

                    UpgradeToHigherOutcome(enuOutcomeOfSendingToWonBoard, enuOutcomesOfSendingToWonBoard.HumanBlocksMyGameWinningBoard)

                End If

                'No, can they win this board?
            ElseIf lcbBoard.GetBestMoveRankForHuman = enuMoveRanks.WinningMove Then

                UpgradeToHigherOutcome(enuOutcomeOfSendingToWonBoard, enuOutcomesOfSendingToWonBoard.HumanWinsABoard)

                'No, can they block on this board?
            ElseIf lcbBoard.GetBestMoveRankForHuman = enuMoveRanks.BlockingMove Then

                UpgradeToHigherOutcome(enuOutcomeOfSendingToWonBoard, enuOutcomesOfSendingToWonBoard.HumanBlocksMyBoardWin)

            End If

        Next

        Return enuOutcomeOfSendingToWonBoard

    End Function



    ' -------------------------------------------------------------------------
    ' Name: FilterMoves()
    ' Abstract: Filters the moves to get the best moves
    ' -------------------------------------------------------------------------
    Private Function FilterMoves(lgspMoves As List(Of CGameSpace), lcbCurrentBoard As CLocalBoard, enuOutcomeOfSendingToWonBoard As enuOutcomesOfSendingToWonBoard) As List(Of CGameSpace)

        Dim lcbNextBoard As CLocalBoard
        Dim lgspNewMoves As New List(Of CGameSpace)
        Dim enuBestHumanRank As enuMoveRanks
        Dim enuCurrentBoardRank As enuMoveRanks
        Dim enuNextBoardHumanRank As enuMoveRanks
        Dim enuCurrentGameSpaceRank As enuMoveRanks
        Dim blnAddNonWinningMoves As Boolean = True


        enuCurrentBoardRank = lcbCurrentBoard.GetRank

        'Loop through and get all the spaces that win immediately/prevent losing immediately
        For Each gspSpace As CGameSpace In lgspMoves

            lcbNextBoard = GetNextLocalBoard(gspSpace)

            enuNextBoardHumanRank = lcbNextBoard.GetHumanRank()
            enuCurrentGameSpaceRank = gspSpace.GetRank

            'Does this move win the game?
            If lcbCurrentBoard.GetRank = enuMoveRanks.WinningMove AndAlso gspSpace.GetRank = enuMoveRanks.WinningMove Then

                'Yes, make it the only move and stop looking
                lgspNewMoves.Clear()
                lgspNewMoves.Add(gspSpace)
                Exit For

                'No, does stop you from losing?
            ElseIf lcbCurrentBoard.GetRank = enuMoveRanks.BlockingMove AndAlso gspSpace.GetRank <= enuMoveRanks.BlockingMove Then

                'Yes, clear the nonwinning/nonblocking moves but keep looking for winning/blocking moves
                If blnAddNonWinningMoves = True Then
                    lgspNewMoves.Clear()
                End If
                lgspNewMoves.Add(gspSpace)
                blnAddNonWinningMoves = False

            Else : lgspNewMoves.Add(gspSpace)
            End If
        Next

        'Are there any left?
        If lgspNewMoves.Count > 0 Then

            'Yes, use the filtered options
            lgspMoves.Clear()
            lgspMoves.AddRange(lgspNewMoves)
            lgspNewMoves.Clear()

        End If

        'Loop through and get all the spaces that don't let the human block or win unless the AI does the same or better
        For Each gspSpace As CGameSpace In lgspMoves

            lcbNextBoard = GetNextLocalBoard(gspSpace)

            enuNextBoardHumanRank = lcbNextBoard.GetHumanRank()
            enuCurrentGameSpaceRank = gspSpace.GetRank

            'Is the next board taken, or does this board return the human to the same board after you win it?
            If enuNextBoardHumanRank = enuMoveRanks.Taken OrElse
            (gspSpace.GetHumanRank = enuMoveRanks.WinningMove AndAlso lcbNextBoard.Name = lcbCurrentBoard.Name) Then

                Select Case enuOutcomeOfSendingToWonBoard

                    'If any of these are true, add it
                    Case Is = enuOutcomesOfSendingToWonBoard.HumanWinsTheGame
                    Case Is = enuOutcomesOfSendingToWonBoard.HumanWinsMyGameWinningBoard
                    Case Is = enuOutcomesOfSendingToWonBoard.HumanBlocksMyGameWinningBoard AndAlso enuCurrentGameSpaceRank > enuMoveRanks.BlockingMove AndAlso enuCurrentBoardRank > enuMoveRanks.BlockingMove
                    Case Is = enuOutcomesOfSendingToWonBoard.HumanWinsABoard AndAlso enuCurrentGameSpaceRank <> enuMoveRanks.WinningMove
                    Case Is = enuOutcomesOfSendingToWonBoard.HumanBlocksMyBoardWin AndAlso enuCurrentGameSpaceRank > enuMoveRanks.BlockingMove
                        ' OrElse (enuOutcomesOfSendingToWonBoard.NothingHappens AndAlso enuCurrentGameSpaceRank >= enuMoveRanks.BlockingMove)

                    Case Else : lgspNewMoves.Add(gspSpace)

                End Select

            Else : lgspNewMoves.Add(gspSpace)

            End If

        Next


        'Are there any left?
        If lgspNewMoves.Count > 0 Then

            'Yes, use the filtered options
            lgspMoves.Clear()
            lgspMoves.AddRange(lgspNewMoves)
            lgspNewMoves.Clear()

        End If

        'Loop through and get all the spaces that are better than the rank of the board you send them to
        For Each gspSpace As CGameSpace In lgspMoves

            lcbNextBoard = GetNextLocalBoard(gspSpace)

            enuNextBoardHumanRank = lcbNextBoard.GetHumanRank()
            enuCurrentGameSpaceRank = gspSpace.GetRank

            If (enuCurrentBoardRank - 1 <= enuNextBoardHumanRank) Then

                lgspNewMoves.Add(gspSpace)
            End If
        Next

        'Are there any?
        If lgspNewMoves.Count > 0 Then

            'Yes, use the filtered options
            lgspMoves.Clear()
            lgspMoves.AddRange(lgspNewMoves)
            lgspNewMoves.Clear()

        End If


        'Loop through and find any that don't give a win or block without getting one
        For Each gspSpace As CGameSpace In lgspMoves

            lcbNextBoard = GetNextLocalBoard(gspSpace)
            enuBestHumanRank = lcbNextBoard.GetBestMoveRankForHuman

            'Can the opponent block or win where you send them, not counting sending to same board?
            If lcbNextBoard.Name <> lcbCurrentBoard.Name AndAlso enuBestHumanRank <= enuMoveRanks.BlockingMove Then

                'Yes, can we do the same or better?
                If gspSpace.GetRank <= enuMoveRanks.BlockingMove Then

                    'Yes, add it
                    lgspNewMoves.Add(gspSpace)

                End If

                'No, do it
            Else : lgspNewMoves.Add(gspSpace)

            End If
        Next

        'Are there any?
        If lgspNewMoves.Count > 0 Then

            'Yes, use the filtered options
            lgspMoves.Clear()
            lgspMoves.AddRange(lgspNewMoves)
            lgspNewMoves.Clear()

        End If

        Return lgspMoves
    End Function


    ' -------------------------------------------------------------------------
    ' Name: GetAllMovesFromAllBoards
    ' Abstract: Gets all moves from all boards
    ' -------------------------------------------------------------------------
    Private Function GetAllMovesFromAllBoards(plrTarget As CPlayer) As List(Of CGameSpace)

        Dim lgspTargets As New List(Of CGameSpace)

        For Each lcbBoards As CLocalBoard In Me.Controls

            lgspTargets.AddRange(lcbBoards.GetAllNotTakenMoves())

        Next

        Return lgspTargets

    End Function



    ' -------------------------------------------------------------------------
    ' Name: GetBestMovesFromAllBoards
    ' Abstract: Gets the best moves from all boards
    ' -------------------------------------------------------------------------
    Private Function GetBestMovesFromAllBoards(plrTarget As CPlayer) As List(Of CGameSpace)

        Dim lgspTargets As New List(Of CGameSpace)
        Dim enuCurrentRank As enuMoveRanks = enuMoveRanks.WinningMove

        While (lgspTargets.Count = 0 AndAlso enuCurrentRank < enuMoveRanks.Taken)

            For Each lcbBoards As CLocalBoard In Me.Controls

                lgspTargets.AddRange(lcbBoards.GetMovesOfRank(enuCurrentRank))

            Next

            enuCurrentRank = CType(enuCurrentRank + 1, enuMoveRanks)

        End While

        Return lgspTargets

    End Function



    ' -------------------------------------------------------------------------
    ' Name: SetGameOutcome
    ' Abstract: Sets the game outcome
    ' -------------------------------------------------------------------------
    Private Sub SetGameOutcome(enuGameOutcome As enuGameOutcomes)

        For intBoardIndex As Integer = 0 To 8
            For intButtonIndex As Integer = 0 To 8
                Me.Controls(intBoardIndex).Controls(intButtonIndex).Enabled = False
            Next
        Next

        'set text
        m_clsGame.SetOutcome(enuGameOutcome)

        m_blnIsGameOver = True

        Me.Invalidate()
        Me.Refresh()

        RaiseEvent GameIsOver(Me, EventArgs.Empty)

        If enuGameOutcome = enuGameOutcomes.ComputerWon Then
            MessageBox.Show("The computer wins.")

        ElseIf enuGameOutcome = enuGameOutcomes.PlayerWon Then
            MessageBox.Show(m_clsGame.GetName() + " wins.")

        ElseIf enuGameOutcome = enuGameOutcomes.Draw Then
            MessageBox.Show("Cat's game")
        End If

    End Sub



    ' -------------------------------------------------------------------------
    ' Name: DisableAllBoardsButOne
    ' Abstract: Disables all buttons but the one on the selected board
    ' -------------------------------------------------------------------------
    Private Sub DisableAllBoardsButOne(grpBoardToEnable As CLocalBoard)

        For intIndex As Integer = 0 To 8

            Me.Controls(intIndex).Enabled = False

        Next

        grpBoardToEnable.Enabled = True
        Thread.Sleep(100)

    End Sub



    ' -------------------------------------------------------------------------
    ' Name: MakeControlFlash
    ' Abstract:  Makes the designated control Flash with a transparent rectangle
    ' -------------------------------------------------------------------------
    Private Sub MakeControlFlash(ctrTarget As Control)

        'Dim m_clrOld As Color
        Dim sbrTransparentLightBlue As New SolidBrush(Color.FromArgb(80, 50, 50, 255))
        Dim sbrTransparentLightGreen As New SolidBrush(Color.FromArgb(80, 50, 255, 50))
        Dim sbrTarget As SolidBrush

        If m_blnIsGameOver = False Then

            If m_clsGame.GetCurrentPlayer = m_clsGame.plrComputer Then
                sbrTarget = sbrTransparentLightBlue
            Else
                sbrTarget = sbrTransparentLightGreen
            End If

            'Draw a transparent rectangle over the control
            ctrTarget.CreateGraphics.FillRectangle(sbrTarget, 2, 2, ctrTarget.Size.Width - 5, ctrTarget.Size.Height - 5)

            Thread.Sleep(100)
            Me.Invalidate()
            Me.Refresh()
            Thread.Sleep(100)

            ctrTarget.CreateGraphics.FillRectangle(sbrTarget, 2, 2, ctrTarget.Size.Width - 5, ctrTarget.Size.Height - 5)

            Thread.Sleep(100)

            'Remove it again
            Me.Invalidate()
            Me.Refresh()

        End If
    End Sub



    ' -------------------------------------------------------------------------
    ' Name: MakeAllBoardsFlash
    ' Abstract:  Makes all boards flash with a transparent rectangle
    ' -------------------------------------------------------------------------
    Private Sub MakeAllBoardsFlash()
        Dim sbrTransparentLightBlue As New SolidBrush(Color.FromArgb(80, 50, 50, 255))
        Dim sbrTransparentLightGreen As New SolidBrush(Color.FromArgb(80, 50, 255, 50))
        Dim sbrTarget As SolidBrush

        If m_blnIsGameOver = False Then

            If m_clsGame.GetCurrentPlayer = m_clsGame.plrComputer Then
                sbrTarget = sbrTransparentLightBlue
            Else
                sbrTarget = sbrTransparentLightGreen
            End If

            'Add a transparent rectangle over each board
            For Each lcbBoard As CLocalBoard In Me.Controls
                lcbBoard.CreateGraphics.FillRectangle(sbrTarget, 2, 2, lcbBoard.Size.Width - 5, lcbBoard.Size.Height - 5)
            Next

            Thread.Sleep(125)

            'Remove the rectangle and wait 1/4th second
            Me.Invalidate()
            Me.Refresh()
            Thread.Sleep(125)

            'return it
            For Each lcbBoard As CLocalBoard In Me.Controls
                lcbBoard.CreateGraphics.FillRectangle(sbrTarget, 2, 2, lcbBoard.Size.Width - 5, lcbBoard.Size.Height - 5)
            Next

            Thread.Sleep(125)

            'Remove it again
            Me.Invalidate()
            Me.Refresh()
        End If
    End Sub



    ' -------------------------------------------------------------------------
    ' Name: EnableAllBoards
    ' Abstract:  Enables all boards
    ' -------------------------------------------------------------------------
    Private Sub EnableAllBoards()

        For intIndex As Integer = 0 To 8

            Me.Controls(intIndex).Enabled = True

        Next

    End Sub



    ' -------------------------------------------------------------------------
    ' Name: EnableAllButtons()
    ' Abstract: Enable all buttons
    ' -------------------------------------------------------------------------
    Private Sub EnableAllButtons()

        For intBoardIndex As Integer = 0 To 8
            Me.Controls(intBoardIndex).Enabled = True
            For intButtonIndex As Integer = 0 To 8
                Me.Controls(intBoardIndex).Controls(intButtonIndex).Enabled = True
            Next
        Next
    End Sub



    ' -------------------------------------------------------------------------
    ' Name: InsertXOrO
    ' Abstract: Puts an x or an o into a box
    ' -------------------------------------------------------------------------
    Private Function InsertXOrO(gspSender As CGameSpace) As CLocalBoard

        Dim lcbNext As CLocalBoard
        Static Dim blnFlashNextInsert As Boolean = False

        MakeControlFlash(gspSender)

        If blnFlashNextInsert = True Then
            MakeControlFlash(CType(gspSender.Parent, CLocalBoard))
            blnFlashNextInsert = False

        End If

        m_clsGame.DoMove(gspSender)

        lcbNext = GetNextLocalBoard(gspSender)

        MakeControlFlash(lcbNext)

        If lcbNext.IsBoardPlayable = True Then

            DisableAllBoardsButOne(lcbNext)

        Else
            EnableAllBoards()
            MakeAllBoardsFlash()
            blnFlashNextInsert = True
        End If

        Return lcbNext

    End Function



    ' -------------------------------------------------------------------------
    ' Name: GetNextLocalBoard
    ' Abstract: Gets the next local board
    ' -------------------------------------------------------------------------
    Private Function GetNextLocalBoard(gspSender As CGameSpace) As CLocalBoard

        Dim grpNextLocalBoard As CLocalBoard
        Dim intBoardNumber As Integer

        intBoardNumber = gspSender.GetGameSpaceID

        grpNextLocalBoard = CType(Me.Controls(intBoardNumber), CLocalBoard)

        Return grpNextLocalBoard

    End Function
End Class
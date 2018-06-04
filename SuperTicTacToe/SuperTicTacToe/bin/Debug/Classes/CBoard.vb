Public Class CBoard(Of Generic As {Control, IRankedControl, New})
    Inherits GroupBox
    Protected m_lgenMoves As New List(Of Generic)
    Protected m_enuBoardState As enuBoardStates
    Protected m_clsGame As CGame
    Public Event BoardIsWon(brdMe As CBoard(Of Generic), e As EventArgs)
    Protected m_clrCorrect As Color
    Public Event GameSpaceClick(genTarget As Generic, e As EventArgs)
    Protected ReadOnly m_clrHuman As Color = System.Drawing.Color.FromArgb(128, 255, 128)
    Protected ReadOnly m_clrComputer As Color = System.Drawing.Color.FromArgb(128, 255, 255)
    Protected ReadOnly m_clrEnabled As Color = System.Drawing.Color.FromArgb(192, 192, 255)
    Protected ReadOnly m_clrDisabled As Color = System.Drawing.Color.FromArgb(128, 128, 255)


    ' -------------------------------------------------------------------------
    ' Name: GetBoardState
    ' Abstract: Returns the Board State
    ' -------------------------------------------------------------------------
    Public Function GetBoardState() As enuBoardStates

        Return m_enuBoardState

    End Function



    ' -------------------------------------------------------------------------
    ' Name: IsBoardPlayable
    ' Abstract: Returns if the board is playable
    ' -------------------------------------------------------------------------
    Public Function IsBoardPlayable() As Boolean

        If m_enuBoardState = enuBoardStates.StillPlayable Then
            Return True
        Else
            Return False
        End If

    End Function



    ' -------------------------------------------------------------------------
    ' Name: GetAllMoves
    ' Abstract: Gets the best moves that are available
    ' -------------------------------------------------------------------------
    Public Function GetAllNotTakenMoves() As List(Of Generic)

        Dim lgenNotTaken As New List(Of Generic)

        For Each gen In m_lgenMoves

            If gen.GetRank <> enuMoveRanks.Taken Then

                lgenNotTaken.Add(gen)
            End If
        Next
        Return lgenNotTaken
    End Function



    ' -------------------------------------------------------------------------
    ' Name: GetBestMovesForHuman
    ' Abstract: Gets the best moves that are available
    ' -------------------------------------------------------------------------
    Public Function GetBestMovesForHuman(Optional enuStartingRank As enuMoveRanks = enuMoveRanks.WinningMove) As List(Of Generic)

        Dim lgenTargets As New List(Of Generic)
        Dim enuCurrentRank As enuMoveRanks
        Dim enuBestRank As enuMoveRanks = enuMoveRanks.Taken

        If IsBoardPlayable() = True Then

            'Get the highest rank
            For Each genMove In m_lgenMoves

                enuCurrentRank = genMove.GetHumanRank

                If enuCurrentRank >= enuStartingRank AndAlso enuCurrentRank < enuBestRank Then

                    enuBestRank = enuCurrentRank

                End If

            Next

            'Is the highest rank taken?
            If enuBestRank <> enuMoveRanks.Taken Then

                'No, get all moves of  the highest rank
                lgenTargets = GetHumanMovesOfRank(enuBestRank)

            End If

        End If

        Return lgenTargets


    End Function



    ' -------------------------------------------------------------------------
    ' Name: GetBestMoves
    ' Abstract: Gets the best moves that are available
    ' -------------------------------------------------------------------------
    Public Function GetBestMoves(Optional enuStartingRank As enuMoveRanks = enuMoveRanks.WinningMove) As List(Of Generic)

        Dim lgenTargets As New List(Of Generic)
        Dim enuBestRank As enuMoveRanks = enuMoveRanks.Taken

        If IsBoardPlayable() = True Then

            'Get the highest rank
            For Each genMove In m_lgenMoves

                If genMove.GetRank >= enuStartingRank AndAlso genMove.GetRank < enuBestRank Then

                    enuBestRank = genMove.GetRank

                End If

            Next

            'Is the highest rank taken?
            If enuBestRank <> enuMoveRanks.Taken Then

                'No, get all moves of  the highest rank
                lgenTargets = GetMovesOfRank(enuBestRank)

            End If

        End If

        Return lgenTargets


    End Function



    ' -------------------------------------------------------------------------
    ' Name: GetMovesOfHumanRank
    ' Abstract: Gets moves of the designated rank
    ' -------------------------------------------------------------------------
    Public Function GetHumanMovesOfRank(enuTargetRank As enuMoveRanks) As List(Of Generic)

        Dim lgenTargets As New List(Of Generic)

        For Each genMove In m_lgenMoves

            If genMove.GetHumanRank = enuTargetRank Then

                lgenTargets.Add(genMove)

            End If

        Next

        Return lgenTargets

    End Function



    ' -------------------------------------------------------------------------
    ' Name: GetMovesOfRank
    ' Abstract: Gets moves of the designated rank
    ' -------------------------------------------------------------------------
    Public Function GetMovesOfRank(enuTargetRank As enuMoveRanks) As List(Of Generic)

        Dim lgenTargets As New List(Of Generic)

        For Each genMove In m_lgenMoves

            If genMove.GetRank = enuTargetRank Then

                lgenTargets.Add(genMove)

            End If

        Next

        Return lgenTargets

    End Function



    ' -------------------------------------------------------------------------
    ' Name: SetBoardState()
    ' Abstract: Disables all buttons AndAlso sets the board as won/drawn
    ' -------------------------------------------------------------------------
    Protected Sub SetBoardState(enuNewBoardState As enuBoardStates)

        m_enuBoardState = enuNewBoardState

        'Is it a draw?
        If enuNewBoardState = enuBoardStates.Draw Then

            Me.BackColor = m_clrDisabled

        Else
            'No, one of the players won

            'Which player won?
            If enuNewBoardState = enuBoardStates.PlayerWon Then

                Me.BackColor = m_clrHuman

            ElseIf enuNewBoardState = enuBoardStates.ComputerWon Then

                Me.BackColor = m_clrComputer

            End If

        End If

        m_clrCorrect = Me.BackColor

        For Each gen As Generic In Me.Controls

            gen.Enabled = False
            gen.SetRank(enuMoveRanks.Taken)
            gen.SetHumanRank(enuMoveRanks.Taken)

        Next

        Me.Invalidate()
        Me.Refresh()

        RaiseEvent BoardIsWon(Me, EventArgs.Empty)

    End Sub



    ' -------------------------------------------------------------------------
    ' Name: SetBoxColor
    ' Abstract: Sets the box's color to the current player's
    ' -------------------------------------------------------------------------
    Protected Sub SetBoxColor(genTarget As Generic)
        If m_clsGame.IsComputerTurn = True Then
            genTarget.BackColor = m_clrComputer
        Else
            genTarget.BackColor = m_clrHuman
        End If

        genTarget.Invalidate()
        genTarget.Refresh()
    End Sub



    ' -------------------------------------------------------------------------
    ' Name: gen0_SpaceIsWon
    ' Abstract: Puts an x or an o into a box
    ' -------------------------------------------------------------------------
    Protected Sub gen0_SpaceIsWon(genTarget As Generic, e As EventArgs)

        Dim intIncrementAmount As Integer
        Dim clrTarget As Color
        Dim enuAIRank As enuMoveRanks
        Dim enuHumanRank As enuMoveRanks

        If m_clsGame.IsComputerTurn = True Then
            intIncrementAmount = 1
            clrTarget = m_clsGame.m_clrComputer
            enuAIRank = enuMoveRanks.WinningMove
            enuHumanRank = enuMoveRanks.BlockingMove

            'The same player wouldn't want to play both this button and these buttons
            CType(Me.Controls(5), Generic).IncrementRank(-1)
            CType(Me.Controls(7), Generic).IncrementRank(-1)

        Else
            intIncrementAmount = -1
            clrTarget = m_clsGame.m_clrHuman
            enuAIRank = enuMoveRanks.BlockingMove
            enuHumanRank = enuMoveRanks.WinningMove

            'The same player wouldn't want to play both this button and these buttons
            CType(Me.Controls(5), Generic).IncrementHumanRank(-1)
            CType(Me.Controls(7), Generic).IncrementHumanRank(-1)

        End If

        IncrementTopRowRanks(intIncrementAmount)
        IncrementLeftColumnRank(intIncrementAmount)
        IncrementTopLeftDiagonalRank(intIncrementAmount)

        CheckTopRow(genTarget)
        CheckLeftColumn(genTarget)

        'Diagonals
        If Me.Controls(4).BackColor = clrTarget Then

            CType(Me.Controls(8), Generic).SetRank(enuAIRank)
            CType(Me.Controls(8), Generic).SetHumanRank(enuHumanRank)

        ElseIf Me.Controls(8).BackColor = clrTarget Then

            CType(Me.Controls(4), Generic).SetRank(enuAIRank)
            CType(Me.Controls(4), Generic).SetHumanRank(enuHumanRank)

        End If

        GenericClick(genTarget)

    End Sub



    ' -------------------------------------------------------------------------
    ' Name: gen1_SpaceIsWon
    ' Abstract: Puts an x or an o into a box
    ' -------------------------------------------------------------------------
    Protected Sub gen1_SpaceIsWon(genTarget As Generic, e As EventArgs)
        Dim intIncrementAmount As Integer

        If m_clsGame.IsComputerTurn = True Then
            intIncrementAmount = 1
            CType(Me.Controls(3), Generic).IncrementRank(-1)
            CType(Me.Controls(5), Generic).IncrementRank(-1)
            CType(Me.Controls(6), Generic).IncrementRank(-1)
            CType(Me.Controls(8), Generic).IncrementRank(-1)
        Else
            intIncrementAmount = -1
            CType(Me.Controls(3), Generic).IncrementHumanRank(-1)
            CType(Me.Controls(5), Generic).IncrementHumanRank(-1)
            CType(Me.Controls(6), Generic).IncrementHumanRank(-1)
            CType(Me.Controls(8), Generic).IncrementHumanRank(-1)
        End If

        IncrementTopRowRanks(intIncrementAmount)
        IncrementMiddleColumnRank(intIncrementAmount)

        CheckTopRow(genTarget)
        CheckMiddleColumn(genTarget)

        GenericClick(genTarget)
    End Sub



    ' -------------------------------------------------------------------------
    ' Name: gen2_SpaceIsWon
    ' Abstract: Puts an x or an o into a box
    ' -------------------------------------------------------------------------
    Protected Sub gen2_SpaceIsWon(genTarget As Generic, e As EventArgs)

        Dim intIncrementAmount As Integer
        Dim clrTarget As Color
        Dim enuAIRank As enuMoveRanks
        Dim enuHumanRank As enuMoveRanks

        If m_clsGame.IsComputerTurn = True Then
            intIncrementAmount = 1
            clrTarget = m_clsGame.m_clrComputer
            enuAIRank = enuMoveRanks.WinningMove
            enuHumanRank = enuMoveRanks.BlockingMove
            CType(Me.Controls(3), Generic).IncrementRank(-1)
            CType(Me.Controls(7), Generic).IncrementRank(-1)
        Else
            intIncrementAmount = -1
            clrTarget = m_clsGame.m_clrHuman
            enuAIRank = enuMoveRanks.BlockingMove
            enuHumanRank = enuMoveRanks.WinningMove
            CType(Me.Controls(3), Generic).IncrementHumanRank(-1)
            CType(Me.Controls(7), Generic).IncrementHumanRank(-1)
        End If


        IncrementTopRowRanks(intIncrementAmount)
        IncrementRightColumnRank(intIncrementAmount)
        IncrementTopRightDiagonalRank(intIncrementAmount)

        CheckTopRow(genTarget)
        CheckRightColumn(genTarget)

        If Me.Controls(4).BackColor = clrTarget Then

            CType(Me.Controls(6), Generic).SetRank(enuAIRank)
            CType(Me.Controls(6), Generic).SetHumanRank(enuHumanRank)

        ElseIf Me.Controls(6).BackColor = clrTarget Then

            CType(Me.Controls(4), Generic).SetRank(enuAIRank)
            CType(Me.Controls(4), Generic).SetHumanRank(enuHumanRank)
        End If

        GenericClick(genTarget)

    End Sub



    ' -------------------------------------------------------------------------
    ' Name: gen3_SpaceIsWon
    ' Abstract: Puts an x or an o into a box
    ' -------------------------------------------------------------------------
    Protected Sub gen3_SpaceIsWon(genTarget As Generic, e As EventArgs)
        Dim intIncrementAmount As Integer

        If m_clsGame.IsComputerTurn = True Then
            intIncrementAmount = 1

            CType(Me.Controls(1), Generic).IncrementRank(-1)
            CType(Me.Controls(2), Generic).IncrementRank(-1)
            CType(Me.Controls(7), Generic).IncrementRank(-1)
            CType(Me.Controls(8), Generic).IncrementRank(-1)
        Else
            intIncrementAmount = -1

            CType(Me.Controls(1), Generic).IncrementHumanRank(-1)
            CType(Me.Controls(2), Generic).IncrementHumanRank(-1)
            CType(Me.Controls(7), Generic).IncrementHumanRank(-1)
            CType(Me.Controls(8), Generic).IncrementHumanRank(-1)
        End If

        IncrementMiddleRowRanks(intIncrementAmount)
        IncrementLeftColumnRank(intIncrementAmount)

        CheckMiddleRow(genTarget)
        CheckLeftColumn(genTarget)

        GenericClick(genTarget)

    End Sub



    ' -------------------------------------------------------------------------
    ' Name: gen4_SpaceIsWon
    ' Abstract: Puts an x or an o into a box
    ' -------------------------------------------------------------------------
    Protected Sub gen4_SpaceIsWon(genTarget As Generic, e As EventArgs)

        Dim intIncrementAmount As Integer
        Dim clrTarget As Color
        Dim enuAIRank As enuMoveRanks
        Dim enuHumanRank As enuMoveRanks

        If m_clsGame.IsComputerTurn = True Then
            intIncrementAmount = 1
            clrTarget = m_clsGame.m_clrComputer
            enuAIRank = enuMoveRanks.WinningMove
            enuHumanRank = enuMoveRanks.BlockingMove
        Else
            intIncrementAmount = -1
            clrTarget = m_clsGame.m_clrHuman
            enuAIRank = enuMoveRanks.BlockingMove
            enuHumanRank = enuMoveRanks.WinningMove
        End If

        IncrementMiddleRowRanks(intIncrementAmount)
        IncrementMiddleColumnRank(intIncrementAmount)
        IncrementTopLeftDiagonalRank(intIncrementAmount)
        IncrementTopRightDiagonalRank(intIncrementAmount)

        CheckMiddleRow(genTarget)
        CheckMiddleColumn(genTarget)

        If Me.Controls(0).BackColor = clrTarget Then

            CType(Me.Controls(8), Generic).SetRank(enuAIRank)
            CType(Me.Controls(8), Generic).SetHumanRank(enuHumanRank)

        ElseIf Me.Controls(8).BackColor = clrTarget Then

            CType(Me.Controls(0), Generic).SetRank(enuAIRank)
            CType(Me.Controls(0), Generic).SetHumanRank(enuHumanRank)

        End If

        If Me.Controls(2).BackColor = clrTarget Then

            CType(Me.Controls(6), Generic).SetRank(enuAIRank)
            CType(Me.Controls(6), Generic).SetHumanRank(enuHumanRank)

        ElseIf Me.Controls(6).BackColor = clrTarget Then

            CType(Me.Controls(2), Generic).SetRank(enuAIRank)
            CType(Me.Controls(2), Generic).SetHumanRank(enuHumanRank)

        End If

        GenericClick(genTarget)

    End Sub



    ' -------------------------------------------------------------------------
    ' Name: gen5_SpaceIsWon
    ' Abstract: Puts an x or an o into a box
    ' -------------------------------------------------------------------------
    Protected Sub gen5_SpaceIsWon(genTarget As Generic, e As EventArgs)
        Dim intIncrementAmount As Integer

        If m_clsGame.IsComputerTurn = True Then
            intIncrementAmount = 1
            CType(Me.Controls(0), Generic).IncrementRank(-1)
            CType(Me.Controls(1), Generic).IncrementRank(-1)
            CType(Me.Controls(6), Generic).IncrementRank(-1)
            CType(Me.Controls(7), Generic).IncrementRank(-1)
        Else
            intIncrementAmount = -1
            CType(Me.Controls(0), Generic).IncrementHumanRank(-1)
            CType(Me.Controls(1), Generic).IncrementHumanRank(-1)
            CType(Me.Controls(6), Generic).IncrementHumanRank(-1)
            CType(Me.Controls(7), Generic).IncrementHumanRank(-1)
        End If

        IncrementMiddleRowRanks(intIncrementAmount)
        IncrementRightColumnRank(intIncrementAmount)

        CheckMiddleRow(genTarget)
        CheckRightColumn(genTarget)

        GenericClick(genTarget)

    End Sub



    ' -------------------------------------------------------------------------
    ' Name: gen6_SpaceIsWon
    ' Abstract: Puts an x or an o into a box
    ' -------------------------------------------------------------------------
    Protected Sub gen6_SpaceIsWon(genTarget As Generic, e As EventArgs)

        Dim intIncrementAmount As Integer
        Dim clrTarget As Color
        Dim enuAIRank As enuMoveRanks
        Dim enuHumanRank As enuMoveRanks

        If m_clsGame.IsComputerTurn = True Then
            intIncrementAmount = 1
            clrTarget = m_clsGame.m_clrComputer
            enuAIRank = enuMoveRanks.WinningMove
            enuHumanRank = enuMoveRanks.BlockingMove
            CType(Me.Controls(1), Generic).IncrementRank(-1)
            CType(Me.Controls(5), Generic).IncrementRank(-1)
        Else
            intIncrementAmount = -1
            clrTarget = m_clsGame.m_clrHuman
            enuAIRank = enuMoveRanks.BlockingMove
            enuHumanRank = enuMoveRanks.WinningMove
            CType(Me.Controls(1), Generic).IncrementHumanRank(-1)
            CType(Me.Controls(5), Generic).IncrementHumanRank(-1)
        End If

        IncrementBottomRowRanks(intIncrementAmount)
        IncrementLeftColumnRank(intIncrementAmount)
        IncrementTopRightDiagonalRank(intIncrementAmount)

        CheckBottomRow(genTarget)
        CheckLeftColumn(genTarget)

        If Me.Controls(4).BackColor = clrTarget Then

            CType(Me.Controls(2), Generic).SetRank(enuAIRank)
            CType(Me.Controls(2), Generic).SetHumanRank(enuHumanRank)

        ElseIf Me.Controls(2).BackColor = clrTarget Then

            CType(Me.Controls(4), Generic).SetRank(enuAIRank)
            CType(Me.Controls(4), Generic).SetHumanRank(enuHumanRank)

        End If

        GenericClick(genTarget)

    End Sub



    ' -------------------------------------------------------------------------
    ' Name: gen7_SpaceIsWon
    ' Abstract: Puts an x or an o into a box
    ' -------------------------------------------------------------------------
    Protected Sub gen7_SpaceIsWon(genTarget As Generic, e As EventArgs)

        Dim intIncrementAmount As Integer

        If m_clsGame.IsComputerTurn = True Then
            intIncrementAmount = 1
            CType(Me.Controls(0), Generic).IncrementRank(-1)
            CType(Me.Controls(2), Generic).IncrementRank(-1)
            CType(Me.Controls(3), Generic).IncrementRank(-1)
            CType(Me.Controls(5), Generic).IncrementRank(-1)
        Else
            intIncrementAmount = -1

            CType(Me.Controls(0), Generic).IncrementHumanRank(-1)
            CType(Me.Controls(2), Generic).IncrementHumanRank(-1)
            CType(Me.Controls(3), Generic).IncrementHumanRank(-1)
            CType(Me.Controls(5), Generic).IncrementHumanRank(-1)
        End If

        IncrementBottomRowRanks(intIncrementAmount)
        IncrementMiddleColumnRank(intIncrementAmount)

        CheckBottomRow(genTarget)
        CheckMiddleColumn(genTarget)

        GenericClick(genTarget)

    End Sub



    ' -------------------------------------------------------------------------
    ' Name: gen8_SpaceIsWon
    ' Abstract: Puts an x or an o into a box
    ' -------------------------------------------------------------------------
    Protected Sub gen8_SpaceIsWon(genTarget As Generic, e As EventArgs)

        Dim intIncrementAmount As Integer
        Dim clrTarget As Color
        Dim enuAIRank As enuMoveRanks
        Dim enuHumanRank As enuMoveRanks

        If m_clsGame.IsComputerTurn = True Then
            intIncrementAmount = 1
            clrTarget = m_clsGame.m_clrComputer
            enuAIRank = enuMoveRanks.WinningMove
            enuHumanRank = enuMoveRanks.BlockingMove
            CType(Me.Controls(1), Generic).IncrementRank(-1)
            CType(Me.Controls(3), Generic).IncrementRank(-1)
        Else
            intIncrementAmount = -1
            clrTarget = m_clsGame.m_clrHuman
            enuAIRank = enuMoveRanks.BlockingMove
            enuHumanRank = enuMoveRanks.WinningMove
            CType(Me.Controls(1), Generic).IncrementHumanRank(-1)
            CType(Me.Controls(3), Generic).IncrementHumanRank(-1)
        End If

        IncrementBottomRowRanks(intIncrementAmount)
        IncrementRightColumnRank(intIncrementAmount)
        IncrementTopLeftDiagonalRank(intIncrementAmount)

        CheckBottomRow(genTarget)
        CheckRightColumn(genTarget)

        If Me.Controls(4).BackColor = clrTarget Then

            CType(Me.Controls(0), Generic).SetRank(enuAIRank)
            CType(Me.Controls(0), Generic).SetHumanRank(enuHumanRank)

        ElseIf Me.Controls(0).BackColor = clrTarget Then

            CType(Me.Controls(4), Generic).SetRank(enuAIRank)
            CType(Me.Controls(4), Generic).SetHumanRank(enuHumanRank)

        End If

        GenericClick(genTarget)

    End Sub



    ' -------------------------------------------------------------------------
    ' Name: GenericClick
    ' Abstract: Changes the color, sets the rank, AndAlso sees if that click wins the board
    ' -------------------------------------------------------------------------
    Protected Sub GenericClick(genTarget As Generic)

        Dim enuNewBoardState As enuBoardStates = enuBoardStates.StillPlayable

        SetBoxColor(genTarget)
        genTarget.Text = m_clsGame.GetCurrentPlayer

        'Is it the computer's turn?
        If m_clsGame.IsComputerTurn = True Then

            'Yes, this move win?
            If genTarget.GetRank = enuMoveRanks.WinningMove Then

                'Yes, set the board state as the computer's win
                enuNewBoardState = enuBoardStates.ComputerWon
            End If

            'No, did the human win?
        ElseIf genTarget.GetHumanRank = enuMoveRanks.WinningMove Then

            'Yes, set the board state as the player's win
            enuNewBoardState = enuBoardStates.PlayerWon

            'No, did this move fill the board?
        ElseIf GetMovesOfRank(enuMoveRanks.Taken).Count = 8 Then

            'Yes, set the board state as a draw
            enuNewBoardState = enuBoardStates.Draw
        End If


        genTarget.SetRank(enuMoveRanks.Taken)
        genTarget.SetHumanRank(enuMoveRanks.Taken)
        genTarget.Enabled = False

        If enuNewBoardState <> enuBoardStates.StillPlayable Then
            SetBoardState(enuNewBoardState)
        End If

        'If this is a button, set it's color and text and raise a click event
        If genTarget.Name.StartsWith("m_gsp") Then

            RaiseEvent GameSpaceClick(genTarget, EventArgs.Empty)

        Else
            genTarget.Text = ""

        End If


    End Sub



    ' -------------------------------------------------------------------------
    ' Name: IncrementTopRowRanks
    ' Abstract: Increments the Top Row's Rank
    ' -------------------------------------------------------------------------
    Protected Sub IncrementTopRowRanks(intAmount As Integer)

        Dim clrOtherPlayer As Color

        clrOtherPlayer = m_clsGame.GetCurrentPlayer.GetOtherPlayerColor

        If Me.Controls(0).BackColor <> clrOtherPlayer AndAlso Me.Controls(1).BackColor <> clrOtherPlayer AndAlso Me.Controls(2).BackColor <> clrOtherPlayer Then
            CType(Me.Controls(0), Generic).IncrementRank(intAmount)
            CType(Me.Controls(1), Generic).IncrementRank(intAmount)
            CType(Me.Controls(2), Generic).IncrementRank(intAmount)
            CType(Me.Controls(0), Generic).IncrementHumanRank(-intAmount)
            CType(Me.Controls(1), Generic).IncrementHumanRank(-intAmount)
            CType(Me.Controls(2), Generic).IncrementHumanRank(-intAmount)
        End If


    End Sub



    ' -------------------------------------------------------------------------
    ' Name: IncrementMiddleRowRanks
    ' Abstract: Increments the Middle Row's Rank
    ' -------------------------------------------------------------------------
    Protected Sub IncrementMiddleRowRanks(intAmount As Integer)

        Dim clrOtherPlayer As Color

        clrOtherPlayer = m_clsGame.GetCurrentPlayer.GetOtherPlayerColor

        If Me.Controls(3).BackColor <> clrOtherPlayer AndAlso Me.Controls(4).BackColor <> clrOtherPlayer AndAlso Me.Controls(5).BackColor <> clrOtherPlayer Then
            CType(Me.Controls(3), Generic).IncrementRank(intAmount)
            CType(Me.Controls(4), Generic).IncrementRank(intAmount)
            CType(Me.Controls(5), Generic).IncrementRank(intAmount)
            CType(Me.Controls(3), Generic).IncrementHumanRank(-intAmount)
            CType(Me.Controls(4), Generic).IncrementHumanRank(-intAmount)
            CType(Me.Controls(5), Generic).IncrementHumanRank(-intAmount)
        End If

    End Sub



    ' -------------------------------------------------------------------------
    ' Name: IncrementBottomRowRanks
    ' Abstract: Increments the Bottom Row's Rank
    ' -------------------------------------------------------------------------
    Protected Sub IncrementBottomRowRanks(intAmount As Integer)

        Dim clrOtherPlayer As Color

        clrOtherPlayer = m_clsGame.GetCurrentPlayer.GetOtherPlayerColor

        If Me.Controls(6).BackColor <> clrOtherPlayer AndAlso Me.Controls(7).BackColor <> clrOtherPlayer AndAlso Me.Controls(8).BackColor <> clrOtherPlayer Then
            CType(Me.Controls(6), Generic).IncrementRank(intAmount)
            CType(Me.Controls(7), Generic).IncrementRank(intAmount)
            CType(Me.Controls(8), Generic).IncrementRank(intAmount)
            CType(Me.Controls(6), Generic).IncrementHumanRank(-intAmount)
            CType(Me.Controls(7), Generic).IncrementHumanRank(-intAmount)
            CType(Me.Controls(8), Generic).IncrementHumanRank(-intAmount)
        End If

    End Sub



    ' -------------------------------------------------------------------------
    ' Name: IncrementLeftColumnRank
    ' Abstract: Increments the Left Column's Rank
    ' -------------------------------------------------------------------------
    Protected Sub IncrementLeftColumnRank(intAmount As Integer)

        Dim clrOtherPlayer As Color

        clrOtherPlayer = m_clsGame.GetCurrentPlayer.GetOtherPlayerColor

        If Me.Controls(0).BackColor <> clrOtherPlayer AndAlso Me.Controls(3).BackColor <> clrOtherPlayer AndAlso Me.Controls(6).BackColor <> clrOtherPlayer Then
            CType(Me.Controls(0), Generic).IncrementRank(intAmount)
            CType(Me.Controls(3), Generic).IncrementRank(intAmount)
            CType(Me.Controls(6), Generic).IncrementRank(intAmount)
            CType(Me.Controls(0), Generic).IncrementHumanRank(-intAmount)
            CType(Me.Controls(3), Generic).IncrementHumanRank(-intAmount)
            CType(Me.Controls(6), Generic).IncrementHumanRank(-intAmount)
        End If

    End Sub



    ' -------------------------------------------------------------------------
    ' Name: IncrementMiddleColumnRank
    ' Abstract: Increments the Middle Column's Rank
    ' -------------------------------------------------------------------------
    Protected Sub IncrementMiddleColumnRank(intAmount As Integer)

        Dim clrOtherPlayer As Color

        clrOtherPlayer = m_clsGame.GetCurrentPlayer.GetOtherPlayerColor

        If Me.Controls(1).BackColor <> clrOtherPlayer AndAlso Me.Controls(4).BackColor <> clrOtherPlayer AndAlso Me.Controls(7).BackColor <> clrOtherPlayer Then
            CType(Me.Controls(1), Generic).IncrementRank(intAmount)
            CType(Me.Controls(4), Generic).IncrementRank(intAmount)
            CType(Me.Controls(7), Generic).IncrementRank(intAmount)
            CType(Me.Controls(1), Generic).IncrementHumanRank(-intAmount)
            CType(Me.Controls(4), Generic).IncrementHumanRank(-intAmount)
            CType(Me.Controls(7), Generic).IncrementHumanRank(-intAmount)
        End If


    End Sub



    ' -------------------------------------------------------------------------
    ' Name: IncrementRightColumnRank
    ' Abstract: Increments the Right Column's Rank
    ' -------------------------------------------------------------------------
    Protected Sub IncrementRightColumnRank(intAmount As Integer)

        Dim clrOtherPlayer As Color

        clrOtherPlayer = m_clsGame.GetCurrentPlayer.GetOtherPlayerColor

        If Me.Controls(2).BackColor <> clrOtherPlayer AndAlso Me.Controls(5).BackColor <> clrOtherPlayer AndAlso Me.Controls(8).BackColor <> clrOtherPlayer Then
            CType(Me.Controls(2), Generic).IncrementRank(intAmount)
            CType(Me.Controls(5), Generic).IncrementRank(intAmount)
            CType(Me.Controls(8), Generic).IncrementRank(intAmount)
            CType(Me.Controls(2), Generic).IncrementHumanRank(-intAmount)
            CType(Me.Controls(5), Generic).IncrementHumanRank(-intAmount)
            CType(Me.Controls(8), Generic).IncrementHumanRank(-intAmount)
        End If

    End Sub



    ' -------------------------------------------------------------------------
    ' Name: IncrementTopLeftDiagonalRank
    ' Abstract: Increments the Top Left Diagonal's Rank
    ' -------------------------------------------------------------------------
    Protected Sub IncrementTopLeftDiagonalRank(intAmount As Integer)

        Dim clrOtherPlayer As Color

        clrOtherPlayer = m_clsGame.GetCurrentPlayer.GetOtherPlayerColor

        If Me.Controls(0).BackColor <> clrOtherPlayer AndAlso Me.Controls(4).BackColor <> clrOtherPlayer AndAlso Me.Controls(8).BackColor <> clrOtherPlayer Then
            CType(Me.Controls(0), Generic).IncrementRank(intAmount)
            CType(Me.Controls(4), Generic).IncrementRank(intAmount)
            CType(Me.Controls(8), Generic).IncrementRank(intAmount)
            CType(Me.Controls(0), Generic).IncrementHumanRank(-intAmount)
            CType(Me.Controls(4), Generic).IncrementHumanRank(-intAmount)
            CType(Me.Controls(8), Generic).IncrementHumanRank(-intAmount)
        End If


    End Sub



    ' -------------------------------------------------------------------------
    ' Name: IncrementTopRightDiagonalRank
    ' Abstract: Increments the Top Right Diagonal's Rank
    ' -------------------------------------------------------------------------
    Protected Sub IncrementTopRightDiagonalRank(intAmount As Integer)

        Dim clrOtherPlayer As Color

        clrOtherPlayer = m_clsGame.GetCurrentPlayer.GetOtherPlayerColor

        If Me.Controls(2).BackColor <> clrOtherPlayer AndAlso Me.Controls(4).BackColor <> clrOtherPlayer AndAlso Me.Controls(6).BackColor <> clrOtherPlayer Then
            CType(Me.Controls(2), Generic).IncrementRank(intAmount)
            CType(Me.Controls(4), Generic).IncrementRank(intAmount)
            CType(Me.Controls(6), Generic).IncrementRank(intAmount)
            CType(Me.Controls(2), Generic).IncrementHumanRank(-intAmount)
            CType(Me.Controls(4), Generic).IncrementHumanRank(-intAmount)
            CType(Me.Controls(6), Generic).IncrementHumanRank(-intAmount)
        End If


    End Sub



    ' -------------------------------------------------------------------------
    ' Name: CheckTopRow
    ' Abstract: Checks if a Generic in the top row can be won next turn
    ' -------------------------------------------------------------------------
    Protected Sub CheckTopRow(genTarget As Generic)

        Dim intTargetIndex As Integer
        Dim clrTarget As Color
        Dim enuAIRank As enuMoveRanks
        Dim enuHumanRank As enuMoveRanks

        intTargetIndex = Me.Controls.IndexOf(genTarget)

        If m_clsGame.IsComputerTurn = True Then
            clrTarget = m_clsGame.m_clrComputer
            enuAIRank = enuMoveRanks.WinningMove
            enuHumanRank = enuMoveRanks.BlockingMove
        Else
            clrTarget = m_clsGame.m_clrHuman
            enuAIRank = enuMoveRanks.BlockingMove
            enuHumanRank = enuMoveRanks.WinningMove
        End If

        If Me.Controls(intTargetIndex + 3).BackColor = clrTarget Then

            CType(Me.Controls(intTargetIndex + 6), Generic).SetRank(enuAIRank)
            CType(Me.Controls(intTargetIndex + 6), Generic).SetHumanRank(enuHumanRank)

        ElseIf Me.Controls(intTargetIndex + 6).BackColor = clrTarget Then

            CType(Me.Controls(intTargetIndex + 3), Generic).SetRank(enuAIRank)
            CType(Me.Controls(intTargetIndex + 3), Generic).SetHumanRank(enuHumanRank)

        End If

    End Sub



    ' -------------------------------------------------------------------------
    ' Name: CheckMiddleRow
    ' Abstract: Checks if a Generic in the middle row can be won next turn
    ' -------------------------------------------------------------------------
    Protected Sub CheckMiddleRow(genTarget As Generic)

        Dim intTargetIndex As Integer
        Dim clrTarget As Color
        Dim enuAIRank As enuMoveRanks
        Dim enuHumanRank As enuMoveRanks

        intTargetIndex = Me.Controls.IndexOf(genTarget)

        If m_clsGame.IsComputerTurn = True Then
            clrTarget = m_clsGame.m_clrComputer
            enuAIRank = enuMoveRanks.WinningMove
            enuHumanRank = enuMoveRanks.BlockingMove
        Else
            clrTarget = m_clsGame.m_clrHuman
            enuAIRank = enuMoveRanks.BlockingMove
            enuHumanRank = enuMoveRanks.WinningMove
        End If

        If Me.Controls(intTargetIndex - 3).BackColor = clrTarget Then

            CType(Me.Controls(intTargetIndex + 3), Generic).SetRank(enuAIRank)
            CType(Me.Controls(intTargetIndex + 3), Generic).SetHumanRank(enuHumanRank)

        ElseIf Me.Controls(intTargetIndex + 3).BackColor = clrTarget Then

            CType(Me.Controls(intTargetIndex - 3), Generic).SetRank(enuAIRank)
            CType(Me.Controls(intTargetIndex - 3), Generic).SetHumanRank(enuHumanRank)
        End If
    End Sub



    ' -------------------------------------------------------------------------
    ' Name: CheckBottomRow
    ' Abstract: Checks if a Generic in the bottom row can be won next turn
    ' -------------------------------------------------------------------------
    Protected Sub CheckBottomRow(genTarget As Generic)

        Dim intTargetIndex As Integer
        Dim clrTarget As Color
        Dim enuAIRank As enuMoveRanks
        Dim enuHumanRank As enuMoveRanks

        intTargetIndex = Me.Controls.IndexOf(genTarget)

        If m_clsGame.IsComputerTurn = True Then
            clrTarget = m_clsGame.m_clrComputer
            enuAIRank = enuMoveRanks.WinningMove
            enuHumanRank = enuMoveRanks.BlockingMove
        Else
            clrTarget = m_clsGame.m_clrHuman
            enuAIRank = enuMoveRanks.BlockingMove
            enuHumanRank = enuMoveRanks.WinningMove
        End If

        If Me.Controls(intTargetIndex - 6).BackColor = clrTarget Then

            CType(Me.Controls(intTargetIndex - 3), Generic).SetRank(enuAIRank)
            CType(Me.Controls(intTargetIndex - 3), Generic).SetHumanRank(enuHumanRank)

        ElseIf Me.Controls(intTargetIndex - 3).BackColor = clrTarget Then

            CType(Me.Controls(intTargetIndex - 6), Generic).SetRank(enuAIRank)
            CType(Me.Controls(intTargetIndex - 6), Generic).SetHumanRank(enuHumanRank)
        End If
    End Sub



    ' -------------------------------------------------------------------------
    ' Name: CheckLeftColumn
    ' Abstract: Checks if a Generic in the left column can be won next turn
    ' -------------------------------------------------------------------------
    Protected Sub CheckLeftColumn(genTarget As Generic)

        Dim intTargetIndex As Integer
        Dim clrTarget As Color
        Dim enuAIRank As enuMoveRanks
        Dim enuHumanRank As enuMoveRanks

        intTargetIndex = Me.Controls.IndexOf(genTarget)

        If m_clsGame.IsComputerTurn = True Then
            clrTarget = m_clsGame.m_clrComputer
            enuAIRank = enuMoveRanks.WinningMove
            enuHumanRank = enuMoveRanks.BlockingMove
        Else
            clrTarget = m_clsGame.m_clrHuman
            enuAIRank = enuMoveRanks.BlockingMove
            enuHumanRank = enuMoveRanks.WinningMove
        End If

        If Me.Controls(intTargetIndex + 1).BackColor = clrTarget Then

            CType(Me.Controls(intTargetIndex + 2), Generic).SetRank(enuAIRank)
            CType(Me.Controls(intTargetIndex + 2), Generic).SetHumanRank(enuHumanRank)

        ElseIf Me.Controls(intTargetIndex + 2).BackColor = clrTarget Then

            CType(Me.Controls(intTargetIndex + 1), Generic).SetRank(enuAIRank)
            CType(Me.Controls(intTargetIndex + 1), Generic).SetHumanRank(enuHumanRank)
        End If

    End Sub



    ' -------------------------------------------------------------------------
    ' Name: CheckMiddleColumn
    ' Abstract: Checks if a Generic in the middle column can be won next turn
    ' -------------------------------------------------------------------------
    Protected Sub CheckMiddleColumn(genTarget As Generic)

        Dim intTargetIndex As Integer
        Dim clrTarget As Color
        Dim enuAIRank As enuMoveRanks
        Dim enuHumanRank As enuMoveRanks

        intTargetIndex = Me.Controls.IndexOf(genTarget)

        If m_clsGame.IsComputerTurn = True Then
            clrTarget = m_clsGame.m_clrComputer
            enuAIRank = enuMoveRanks.WinningMove
            enuHumanRank = enuMoveRanks.BlockingMove
        Else
            clrTarget = m_clsGame.m_clrHuman
            enuAIRank = enuMoveRanks.BlockingMove
            enuHumanRank = enuMoveRanks.WinningMove
        End If

        If Me.Controls(intTargetIndex - 1).BackColor = clrTarget Then

            CType(Me.Controls(intTargetIndex + 1), Generic).SetRank(enuAIRank)
            CType(Me.Controls(intTargetIndex + 1), Generic).SetHumanRank(enuHumanRank)

        ElseIf Me.Controls(intTargetIndex + 1).BackColor = clrTarget Then

            CType(Me.Controls(intTargetIndex - 1), Generic).SetRank(enuAIRank)
            CType(Me.Controls(intTargetIndex - 1), Generic).SetHumanRank(enuHumanRank)
        End If
    End Sub



    ' -------------------------------------------------------------------------
    ' Name: CheckRightColumn
    ' Abstract: Checks if a Generic in the right column can be won next turn
    ' -------------------------------------------------------------------------
    Protected Sub CheckRightColumn(genTarget As Generic)

        Dim intTargetIndex As Integer
        Dim clrTarget As Color
        Dim enuAIRank As enuMoveRanks
        Dim enuHumanRank As enuMoveRanks

        intTargetIndex = Me.Controls.IndexOf(genTarget)

        If m_clsGame.IsComputerTurn = True Then
            clrTarget = m_clsGame.m_clrComputer
            enuAIRank = enuMoveRanks.WinningMove
            enuHumanRank = enuMoveRanks.BlockingMove
        Else
            clrTarget = m_clsGame.m_clrHuman
            enuAIRank = enuMoveRanks.BlockingMove
            enuHumanRank = enuMoveRanks.WinningMove
        End If

        If Me.Controls(intTargetIndex - 2).BackColor = clrTarget Then

            CType(Me.Controls(intTargetIndex - 1), Generic).SetRank(enuAIRank)
            CType(Me.Controls(intTargetIndex - 1), Generic).SetHumanRank(enuHumanRank)

        ElseIf Me.Controls(intTargetIndex - 1).BackColor = clrTarget Then

            CType(Me.Controls(intTargetIndex - 2), Generic).SetRank(enuAIRank)
            CType(Me.Controls(intTargetIndex - 2), Generic).SetHumanRank(enuHumanRank)
        End If

    End Sub
End Class

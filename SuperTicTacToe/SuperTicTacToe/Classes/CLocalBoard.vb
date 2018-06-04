Public Class CLocalBoard
    Inherits CBoard(Of CGameSpace)
    Implements IRankedControl
    Private Property m_intBoardID As Integer Implements IRankedControl.m_intID
    Private m_gsp0 As New CGameSpace()
    Private m_gsp1 As New CGameSpace()
    Private m_gsp2 As New CGameSpace()
    Private m_gsp3 As New CGameSpace()
    Private m_gsp4 As New CGameSpace()
    Private m_gsp5 As New CGameSpace()
    Private m_gsp6 As New CGameSpace()
    Private m_gsp7 As New CGameSpace()
    Private m_gsp8 As New CGameSpace()
    Private Property m_mrkRank As New CMoveRank Implements IRankedControl.m_mrkRank
    Private Property m_mrkHumanRank As New CMoveRank Implements IRankedControl.m_mrkHumanRank
    Public Event GameSpaceEnter(gspTarget As CGameSpace, e As EventArgs)
    Public Event GameSpaceLeave(gspTarget As CGameSpace, e As EventArgs)



    ' -------------------------------------------------------------------------
    ' Name: New
    ' Abstract: Initializes the class
    ' -------------------------------------------------------------------------
    Sub New()

        Me.m_gsp0.Location = New System.Drawing.Point(18, 19)
        Me.m_gsp0.Name = "m_gsp0"

        Me.m_gsp1.Location = New System.Drawing.Point(64, 19)
        Me.m_gsp1.Name = "m_gsp1"

        Me.m_gsp2.Location = New System.Drawing.Point(110, 19)
        Me.m_gsp2.Name = "m_gsp2"

        Me.m_gsp3.Location = New System.Drawing.Point(18, 65)
        Me.m_gsp3.Name = "m_gsp3"

        Me.m_gsp4.Location = New System.Drawing.Point(64, 65)
        Me.m_gsp4.Name = "m_gsp4"

        Me.m_gsp5.Location = New System.Drawing.Point(110, 65)
        Me.m_gsp5.Name = "m_gsp5"

        Me.m_gsp6.Location = New System.Drawing.Point(18, 111)
        Me.m_gsp6.Name = "m_gsp6"

        Me.m_gsp7.Location = New System.Drawing.Point(64, 111)
        Me.m_gsp7.Name = "m_gsp7"

        Me.m_gsp8.Location = New System.Drawing.Point(110, 111)
        Me.m_gsp8.Name = "m_gsp8"

        Me.Controls.Add(m_gsp0)
        Me.Controls.Add(m_gsp1)
        Me.Controls.Add(m_gsp2)
        Me.Controls.Add(m_gsp3)
        Me.Controls.Add(m_gsp4)
        Me.Controls.Add(m_gsp5)
        Me.Controls.Add(m_gsp6)
        Me.Controls.Add(m_gsp7)
        Me.Controls.Add(m_gsp8)

        AddHandler m_gsp0.Click, AddressOf gen0_SpaceIsWon
        AddHandler m_gsp1.Click, AddressOf gen1_SpaceIsWon
        AddHandler m_gsp2.Click, AddressOf gen2_SpaceIsWon
        AddHandler m_gsp3.Click, AddressOf gen3_SpaceIsWon
        AddHandler m_gsp4.Click, AddressOf gen4_SpaceIsWon
        AddHandler m_gsp5.Click, AddressOf gen5_SpaceIsWon
        AddHandler m_gsp6.Click, AddressOf gen6_SpaceIsWon
        AddHandler m_gsp7.Click, AddressOf gen7_SpaceIsWon
        AddHandler m_gsp8.Click, AddressOf gen8_SpaceIsWon

        ResetBoard()

        For Each gsp As CGameSpace In Me.Controls
            gsp.BackColor = m_clrDisabled
            gsp.Enabled = False
            gsp.FlatAppearance.BorderSize = 4
            gsp.FlatStyle = System.Windows.Forms.FlatStyle.Flat
            gsp.Font = New System.Drawing.Font("Microsoft Sans Serif", 18.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
            gsp.ForeColor = System.Drawing.Color.Black
            gsp.Size = New System.Drawing.Size(40, 40)
            gsp.TabIndex = 0
            gsp.UseVisualStyleBackColor = False
            AddHandler gsp.EnabledChanged, AddressOf gspGeneric_EnabledChanged
            AddHandler gsp.MouseEnter, AddressOf gspGeneric_MouseEnter
            AddHandler gsp.MouseLeave, AddressOf gspGeneric_MouseLeave
            m_lgenMoves.Add(gsp)
        Next

    End Sub

    ' -------------------------------------------------------------------------
    ' Name: gspGeneric_MouseEnter
    ' Abstract: Highlights the board you're sending them to when you Enter
    ' -------------------------------------------------------------------------
    Private Sub gspGeneric_MouseEnter(gspTarget As CGameSpace, e As EventArgs)

        RaiseEvent GameSpaceEnter(gspTarget, e)

    End Sub



    ' -------------------------------------------------------------------------
    ' Name: gspGeneric_MouseLeave
    ' Abstract: Changes the color back when you leave a button
    ' -------------------------------------------------------------------------
    Private Sub gspGeneric_MouseLeave(gspTarget As CGameSpace, e As EventArgs)

        RaiseEvent GameSpaceLeave(gspTarget, e)

    End Sub



    ' -------------------------------------------------------------------------
    ' Name: gspGeneric_EnabledChanged
    ' Abstract: Changes the color of buttons when they are disabled
    ' -------------------------------------------------------------------------
    Private Sub gspGeneric_EnabledChanged(gspTarget As CGameSpace, e As EventArgs)

        gspTarget.Invalidate()
        gspTarget.Refresh()

        'Has this box been played in?
        If gspTarget.Text = "" Then

            'No, is it enabled?
            If gspTarget.Enabled = True Then

                'Yes, Set it to the enabled color
                gspTarget.BackColor = m_clrEnabled
            Else

                'No, Set it to the disabled color
                gspTarget.BackColor = m_clrDisabled
            End If
        End If


        gspTarget.Invalidate()
        gspTarget.Refresh()
    End Sub



    ' -------------------------------------------------------------------------
    ' Name: SetGame
    ' Abstract: Sets the game for the board
    ' -------------------------------------------------------------------------
    Public Overloads Sub SetGame(clsGame As CGame)

        m_clsGame = clsGame
        For Each gs As CGameSpace In Me.Controls

            gs.Enabled = True
            gs.BackColor = m_clrEnabled
            gs.Text = ""
        Next

    End Sub



    ' -------------------------------------------------------------------------
    ' Name: ResetBoard
    ' Abstract: Resets the board
    ' -------------------------------------------------------------------------
    Public Sub ResetBoard()

        m_gsp0.SetRank(enuMoveRanks.OkayMove, True)
        m_gsp1.SetRank(enuMoveRanks.NeutralMove, True)
        m_gsp2.SetRank(enuMoveRanks.OkayMove, True)
        m_gsp3.SetRank(enuMoveRanks.NeutralMove, True)
        m_gsp4.SetRank(enuMoveRanks.NeutralMove, True)
        m_gsp5.SetRank(enuMoveRanks.NeutralMove, True)
        m_gsp6.SetRank(enuMoveRanks.OkayMove, True)
        m_gsp7.SetRank(enuMoveRanks.NeutralMove, True)
        m_gsp8.SetRank(enuMoveRanks.OkayMove, True)

        m_gsp0.SetHumanRank(enuMoveRanks.OkayMove, True)
        m_gsp1.SetHumanRank(enuMoveRanks.NeutralMove, True)
        m_gsp2.SetHumanRank(enuMoveRanks.OkayMove, True)
        m_gsp3.SetHumanRank(enuMoveRanks.NeutralMove, True)
        m_gsp4.SetHumanRank(enuMoveRanks.NeutralMove, True)
        m_gsp5.SetHumanRank(enuMoveRanks.NeutralMove, True)
        m_gsp6.SetHumanRank(enuMoveRanks.OkayMove, True)
        m_gsp7.SetHumanRank(enuMoveRanks.NeutralMove, True)
        m_gsp8.SetHumanRank(enuMoveRanks.OkayMove, True)

        'Make the button that sends them to the board you're playing on's rank one less
        CType(Me.Controls(m_intBoardID), CGameSpace).IncrementRank(-1)

        For Each gs As CGameSpace In Me.Controls

            gs.Enabled = False
            gs.BackColor = m_clrDisabled
            gs.Text = ""

        Next

        Me.Enabled = True
        Me.BackColor = m_clrEnabled
        m_clrCorrect = m_clrEnabled
        Me.Text = ""
        Me.m_enuBoardState = enuBoardStates.StillPlayable

        Me.Invalidate()
        Me.Refresh()

    End Sub



    ' -------------------------------------------------------------------------
    ' Name: GetValueToHuman
    ' Abstract: Returns the value this board has to the human
    ' -------------------------------------------------------------------------
    Public Function GetBestMoveRankForHuman() As enuMoveRanks
        Dim enuBestRank As enuMoveRanks
        Dim lgspBestMoves As List(Of CGameSpace)

        'Get the highest available move, AndAlso set the board's value to that
        If GetBoardState() = enuBoardStates.StillPlayable Then

            lgspBestMoves = GetBestMovesForHuman()

            enuBestRank = lgspBestMoves(0).GetHumanRank()

        Else

            enuBestRank = enuMoveRanks.Taken
        End If


        Return enuBestRank

    End Function



    ' -------------------------------------------------------------------------
    ' Name: GetValueToAI
    ' Abstract: Returns the value this board has to the human
    ' -------------------------------------------------------------------------
    Public Function GetBestMoveRankForAI() As enuMoveRanks

        Dim enuBestRank As enuMoveRanks
        Dim lgspBestMoves As List(Of CGameSpace)

        'Get the highest available move, AndAlso set the board's value to that
        If GetBoardState() = enuBoardStates.StillPlayable Then

            lgspBestMoves = GetBestMoves()

            enuBestRank = lgspBestMoves(0).GetRank()

        Else

            enuBestRank = enuMoveRanks.Taken
        End If


        Return enuBestRank

    End Function



    ' -------------------------------------------------------------------------
    ' Name: GetHumanRank
    ' Abstract: Gets the HumanRank
    ' -------------------------------------------------------------------------
    Public Function GetHumanRank() As enuMoveRanks Implements IRankedControl.GetHumanRank

        Return m_mrkHumanRank.GetRank

    End Function



    ' -------------------------------------------------------------------------
    ' Name: SetHumanRank
    ' Abstract: Sets the Rank for the Human
    ' -------------------------------------------------------------------------
    Public Sub SetHumanRank(enuNewHumanRank As enuMoveRanks, Optional blnOverrideProtections As Boolean = False) Implements IRankedControl.SetHumanRank

        m_mrkHumanRank.SetRank(enuNewHumanRank, blnOverrideProtections)

    End Sub



    ' -------------------------------------------------------------------------
    ' Name: GetRank
    ' Abstract: Gets the rank
    ' -------------------------------------------------------------------------
    Public Function GetRank() As enuMoveRanks Implements IRankedControl.GetRank

        Return m_mrkRank.GetRank

    End Function



    ' -------------------------------------------------------------------------
    ' Name: SetRank
    ' Abstract: Sets the rank, but only if it's not already a winning,
    ' blocking, or taken move
    ' -------------------------------------------------------------------------
    Public Sub SetRank(enuNewRank As enuMoveRanks, Optional blnOverrideProtections As Boolean = False) Implements IRankedControl.SetRank

        m_mrkRank.SetRank(enuNewRank, blnOverrideProtections)

    End Sub



    ' -------------------------------------------------------------------------
    ' Name: IncrementHumanRank
    ' Abstract: Increases the rank by the specified amount
    ' -------------------------------------------------------------------------
    Public Sub IncrementHumanRank(intAmount As Integer) Implements IRankedControl.IncrementHumanRank

        m_mrkHumanRank.IncrementRank(intAmount)

    End Sub



    ' -------------------------------------------------------------------------
    ' Name: IncrementRank
    ' Abstract: Increases the rank by the specified amount
    ' -------------------------------------------------------------------------
    Public Sub IncrementRank(intAmount As Integer) Implements IRankedControl.IncrementRank

        m_mrkRank.IncrementRank(intAmount)

    End Sub



    ' -------------------------------------------------------------------------
    ' Name: GetBoardID
    ' Abstract: Returns the Board ID
    ' -------------------------------------------------------------------------
    Public Function GetBoardID() As Integer Implements IRankedControl.GetID

        m_intBoardID = Val(Me.Name.Last())

        Return m_intBoardID

    End Function



    ' -------------------------------------------------------------------------
    ' Name: GetCorrectColor
    ' Abstract: Gets the board's correct color for when hovering changes the backcolor
    ' -------------------------------------------------------------------------
    Public Function GetCorrectColor() As Color

        Return m_clrCorrect

    End Function

End Class

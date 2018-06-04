Public Class CGameSpace
    Inherits Button
    Implements IRankedControl
    Private Property m_mrkRank As New CMoveRank Implements IRankedControl.m_mrkRank
    Private Property m_mrkHumanRank As New CMoveRank Implements IRankedControl.m_mrkHumanRank
    Private Property m_intGameSpaceID As Integer Implements IRankedControl.m_intID



    ' -------------------------------------------------------------------------
    ' Name: GetGameSpaceID
    ' Abstract: Gets the GameSpaceID
    ' -------------------------------------------------------------------------
    Public Function GetGameSpaceID() As Integer Implements IRankedControl.GetID

        m_intGameSpaceID = Val(Me.Name.Last())
        Return m_intGameSpaceID

    End Function



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
End Class

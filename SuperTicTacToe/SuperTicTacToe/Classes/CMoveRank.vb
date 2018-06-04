Public Class CMoveRank
    Private m_enuMoveRank As enuMoveRanks

    ' -------------------------------------------------------------------------
    ' Name: GetRank
    ' Abstract: Gets the rank
    ' -------------------------------------------------------------------------
    Public Function GetRank() As enuMoveRanks

        Return m_enuMoveRank

    End Function



    ' -------------------------------------------------------------------------
    ' Name: SetRank
    ' Abstract: Sets the rank, but only if it's not already a winning,
    ' blocking, or taken move
    ' -------------------------------------------------------------------------
    Public Sub SetRank(enuNewRank As enuMoveRanks, Optional blnOverrideProtections As Boolean = False)

        'Is override protections on, or is the new rank taken?
        If blnOverrideProtections = True OrElse enuNewRank = enuMoveRanks.Taken Then

            'Yes, Let it through
            m_enuMoveRank = enuNewRank

            'No, is the current rank taken or winning?
        ElseIf m_enuMoveRank <> enuMoveRanks.Taken AndAlso m_enuMoveRank <> enuMoveRanks.WinningMove Then

            'No, let it through
            m_enuMoveRank = enuNewRank

        End If

    End Sub



    ' -------------------------------------------------------------------------
    ' Name: IncrementRank
    ' Abstract: Increases the rank by the specified amount
    ' -------------------------------------------------------------------------
    Public Sub IncrementRank(intAmount As Integer)

        'Is the rank taken, or higher than blocking?
        If m_enuMoveRank <> modUserDataTypes.enuMoveRanks.Taken _
        AndAlso m_enuMoveRank > modUserDataTypes.enuMoveRanks.BlockingMove Then

            'Increase it
            m_enuMoveRank -= intAmount

        End If

    End Sub

End Class

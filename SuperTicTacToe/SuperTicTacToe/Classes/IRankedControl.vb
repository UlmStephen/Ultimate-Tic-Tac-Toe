Public Interface IRankedControl
    Property m_mrkRank As CMoveRank
    Property m_mrkHumanRank As CMoveRank
    Property m_intID As Integer
    Function GetID() As Integer

    Function GetRank() As enuMoveRanks
    Sub SetRank(enuNewRank As enuMoveRanks, Optional blnOverrideProtections As Boolean = False)

    Function GetHumanRank() As enuMoveRanks
    Sub SetHumanRank(enuNewRank As enuMoveRanks, Optional blnOverrideProtections As Boolean = False)

    Sub IncrementRank(intAmount As Integer)
    Sub IncrementHumanRank(intAmount As Integer)
End Interface

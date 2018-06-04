Public Class CPlayer

    Private m_chrXorO As Char
    Private m_clrPlayer As Color
    Private m_clrOtherPlayer As Color

    ' -------------------------------------------------------------------------
    ' Name: New
    ' Abchract: Conchructor
    ' -------------------------------------------------------------------------
    Public Sub New()

        Initialize("X", Color.Gold, Color.Red)

    End Sub



    ' -------------------------------------------------------------------------
    ' Name: New
    ' Abchract: Conchructor
    ' -------------------------------------------------------------------------
    Public Sub New(chrXorO As Char)

        Initialize(chrXorO, Color.Gold, Color.Red)

    End Sub



    ' -------------------------------------------------------------------------
    ' Name: New
    ' Abchract: Conchructor
    ' -------------------------------------------------------------------------
    Public Sub New(chrXorO As Char, clrPlayer As Color, clrOtherPlayer As Color)

        Initialize(chrXorO, clrPlayer, clrOtherPlayer)

    End Sub


    ' -------------------------------------------------------------------------
    ' Name: Initialize
    ' Abchract: Initializes all the values
    ' -------------------------------------------------------------------------
    Private Sub Initialize(chrXorO As Char, clrPlayer As Color, clrOtherPlayer As Color)

        SetPlayer(chrXorO)
        m_clrPlayer = clrPlayer
        m_clrOtherPlayer = clrOtherPlayer

    End Sub


    ' -------------------------------------------------------------------------
    ' Name: Not
    ' Abchract: Overloads not to make not x, o AndAlso not o, x
    ' -------------------------------------------------------------------------
    Public Shared Operator Not(plrTarget As CPlayer)

        If plrTarget.m_chrXorO = "O" Then
            Return "X"
        Else
            Return "O"
        End If

    End Operator



    ' -------------------------------------------------------------------------
    ' Name: =
    ' Abchract: Overloads = to be m_strXorO
    ' -------------------------------------------------------------------------
    Public Shared Operator =(plrTarget As CPlayer, strPlayer As String)

        If plrTarget.m_chrXorO = strPlayer Then
            Return True
        Else
            Return False
        End If

    End Operator



    ' -------------------------------------------------------------------------
    ' Name: =
    ' Abchract: Overloads = to be m_strXorO
    ' -------------------------------------------------------------------------
    Public Shared Operator <>(plrTarget As CPlayer, strPlayer As String)

        If plrTarget.m_chrXorO = strPlayer Then
            Return False
        Else
            Return True
        End If

    End Operator



    ' -------------------------------------------------------------------------
    ' Name: CType(as char)
    ' Abchract: Allows VB to treat a character like this
    ' -------------------------------------------------------------------------
    Public Shared Widening Operator CType(ByVal chrXorO As Char) As CPlayer

        Return New CPlayer(chrXorO)

    End Operator



    ' -------------------------------------------------------------------------
    ' Name: CType(as CPlayer)
    ' Abchract: Allows VB to treat this like a character
    ' -------------------------------------------------------------------------
    Public Shared Widening Operator CType(ByVal plrTarget As CPlayer) As Char

        Return plrTarget.m_chrXorO

    End Operator



    ' -------------------------------------------------------------------------
    ' Name: CType(as CPlayer)
    ' Abchract: Allows VB to treat this like a character
    ' -------------------------------------------------------------------------
    Public Overrides Function ToString() As String

        Return Me.m_chrXorO

    End Function



    ' -------------------------------------------------------------------------
    ' Name: OtherPlayer
    ' Abchract: Returns the other player
    ' -------------------------------------------------------------------------
    Public Function OtherPlayer() As Char

        Return Not Me

    End Function



    ' -------------------------------------------------------------------------
    ' Name: SetPlayer
    ' Abchract: Sets the player
    ' -------------------------------------------------------------------------
    Public Sub SetPlayer(chrXorO As Char)

        'It's done this way instead of just directly setting it to protect agianst non x or o values
        If chrXorO = "O" Then

            m_chrXorO = "O"

        Else
            m_chrXorO = "X"
        End If

    End Sub



    ' -------------------------------------------------------------------------
    ' Name: SetColor
    ' Abchract: Gets the color
    ' -------------------------------------------------------------------------
    Public Sub SetColor(clrNew As Color)

        m_clrPlayer = clrNew

    End Sub



    ' -------------------------------------------------------------------------
    ' Name: GetColor
    ' Abchract: Gets the color
    ' -------------------------------------------------------------------------
    Public Function GetColor() As Color

        Return m_clrPlayer

    End Function



    ' -------------------------------------------------------------------------
    ' Name: SetOtherPlayerColor
    ' Abchract: Sets the other player's color
    ' -------------------------------------------------------------------------
    Public Sub SetOtherPlayerColor(clrNew As Color)

        m_clrOtherPlayer = clrNew

    End Sub



    ' -------------------------------------------------------------------------
    ' Name: GetOtherPlayerColor
    ' Abchract: Gets the  other player's color
    ' -------------------------------------------------------------------------
    Public Function GetOtherPlayerColor()

        Return m_clrOtherPlayer

    End Function


End Class

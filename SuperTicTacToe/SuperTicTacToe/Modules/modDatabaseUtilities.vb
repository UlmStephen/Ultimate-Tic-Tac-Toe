' ----------------------------------------------------------------------
' Name: modDatabaseUtilities
' Abstract: Some general purpose database utilities
' ----------------------------------------------------------------------

' ----------------------------------------------------------------------
' Options
' ----------------------------------------------------------------------
Option Explicit On

Imports System.Data.SqlClient



Public Module modDatabaseUtilities

    ' ----------------------------------------------------------------------
    ' Name: Module Variables
    ' ----------------------------------------------------------------------

    ' SQL Server Connection string with integrated login v1
    Private m_strDatabaseConnectionStringSQLServerV1 As String = "Provider=SQLOLEDB;" & _
                                                                 "Server=(Local);" & _
                                                                 "Database=dbUltimateTicTacToe;" & _
                                                                 "Integrated Security=SSPI;"

    ' SQL Server Connection string with integrated login v2
    Private m_strDatabaseConnectionStringSQLServerV2 As String = "Provider=SQLOLEDB;" & _
                                                                 "Server=(Local);" & _
                                                                 "Database=dbTeamsAndPlayers;" & _
                                                                 "Trusted_Connection=True;"

    ' SQL Express Connection string                             
    'Private m_strDatabaseConnectionString As String = "Provider=SQLOLEDB;" & _
    '                                                  "Server=(Local)\SQLEXPRESS;" & _
    '                                                  "Database=dbTeamsAndPlayers;" & _
    '                                                  "User ID=sa;" & _
    '                                                  "Password=;"

    ' Access 2000 / Windows XP Connection string                             
    Private m_strDatabaseConnectionStringMSAccessV1 As String = "Provider=Microsoft.Jet.OLEDB.4.0;" & _
                                                                "Data Source=" & Application.StartupPath & "\..\..\Database\TeamsAndPlayers.mdb;" & _
                                                                "User ID=Admin;" & _
                                                                "Password=;"

    ' Access 2007 / Windows 7 Connection string                             
    Private m_strDatabaseConnectionStringMSAccessV2 As String = "Provider=Microsoft.ACE.OLEDB.12.0;" & _
                                                                "Data Source=" & Application.StartupPath & "\..\..\Database\TeamsAndPlayers.accdb;" & _
                                                                "User ID=Admin;" & _
                                                                "Password=;"


    ' In a 2-Tier app we connect once during FMain_Load AndAlso hold
    ' The connection open while until the application closes
    Private m_conAdministrator As OleDb.OleDbConnection

#Region "Open/Close"

    ' ----------------------------------------------------------------------
    ' Name: OpenDatabaseConnection
    ' Abstract: Open a connection to the database.
    '           In a 2-Tier (client server) applicaiton we connect once in FMain
    '           AndAlso hold the connection open until FMain closes
    ' ----------------------------------------------------------------------
    Public Function OpenDatabaseConnectionSQLServer() As Boolean

        Dim blnResult As Boolean = False

        Try

            ' Open a connection to the database
            m_conAdministrator = New OleDb.OleDbConnection
            m_conAdministrator.ConnectionString = m_strDatabaseConnectionStringSQLServerV1
            m_conAdministrator.Open()

            ' Success
            blnResult = True

        Catch excError As Exception

            WriteLog(excError)

        End Try

        Return blnResult

    End Function



    ' ----------------------------------------------------------------------
    ' Name: OpenDatabaseConnectionMSAccess
    ' Abstract: Open a connection to the database.
    '           In a 2-Tier (client server) applicaiton we connect once in FMain
    '           AndAlso hold the connection open until FMain closes
    ' ----------------------------------------------------------------------
    Public Function OpenDatabaseConnectionMSAccess() As Boolean

        Dim blnResult As Boolean = False

        Try

            ' Open a connection to the database
            m_conAdministrator = New OleDb.OleDbConnection
            m_conAdministrator.ConnectionString = m_strDatabaseConnectionStringMSAccessV2
            m_conAdministrator.Open()

            ' Success
            blnResult = True

        Catch excError As Exception

            WriteLog(excError)

        End Try

        Return blnResult

    End Function



    ' ----------------------------------------------------------------------
    ' Name: CloseDatabaseConnection
    ' Abstract: Closes the Database Connection
    ' ----------------------------------------------------------------------
    Public Function CloseDatabaseConnection() As Boolean

        Dim blnResult As Boolean = False

        Try

            ' Anything there?
            If m_conAdministrator IsNot Nothing Then

                ' Yes, is it open?
                If m_conAdministrator.State <> ConnectionState.Closed Then

                    ' Yes, close it
                    m_conAdministrator.Close()

                End If

                ' Clean up
                m_conAdministrator = Nothing

            End If

            blnResult = True

        Catch excError As Exception

            WriteLog(excError)

        End Try

        Return blnResult

    End Function

#End Region
    ' ----------------------------------------------------------------------
    ' Name: StartGame
    ' Abstract: Starts a game
    ' ----------------------------------------------------------------------
    Public Function StartGame(strPlayerName As String, blnComputerMovesFirst As Boolean, intGameDifficultyID As Integer, intGameOutcomeID As Integer, dtmPlayed As Date) As Integer

        Dim intGameID As Integer

        Try

            Dim strInsert As String
            Dim cmdInsert As OleDb.OleDbCommand



            'Build the INSERT command. Never build Command with raw user input to prevent SQL injections
            strInsert = "INSERT INTO TGames ( strPlayerName, blnComputerMovesFirst, intGameDifficultyID, intGameOutcomeID, dtmPlayed ) " + _
                        "VALUES ( ?, ?, ?, ?, ?)" + _
                        "SELECT SCOPE_IDENTITY()"

            'Make the Command instance
            cmdInsert = New OleDb.OleDbCommand(strInsert, m_conAdministrator)

            'Add column values here instead of above to prevent SQL injection attacks
            With cmdInsert.Parameters

                .AddWithValue("1", strPlayerName)
                .AddWithValue("2", blnComputerMovesFirst)
                .AddWithValue("3", intGameDifficultyID)
                .AddWithValue("4", intGameOutcomeID)
                .AddWithValue("5", dtmPlayed)

            End With

            'Insert the row
            intGameID = cmdInsert.ExecuteScalar()

        Catch excError As Exception

            WriteLog(excError)

        End Try

        Return intGameID

    End Function



    ' ----------------------------------------------------------------------
    ' Name: EnterMove
    ' Abstract: Enters A Move into the database
    ' ----------------------------------------------------------------------
    Public Function EnterMove(intGameID As Integer, intMoveIndex As Integer, intGroup As Integer, intSquare As Integer) As Boolean

        Dim blnResult As Boolean = False

        Try

            Dim strInsert As String
            Dim cmdInsert As OleDb.OleDbCommand

            
                'Build the INSERT command. Never build Command with raw user input to prevent SQL injections
            strInsert = "EXECUTE uspAddMove ?, ?, ?, ?"

                'Make the Command instance
                cmdInsert = New OleDb.OleDbCommand(strInsert, m_conAdministrator)

                'Add column values here instead of above to prevent SQL injection attacks
                With cmdInsert.Parameters

                .AddWithValue("1", intGameID)
                .AddWithValue("2", intMoveIndex)
                .AddWithValue("3", intGroup)
                .AddWithValue("4", intSquare)

                End With

                'Insert the row
                cmdInsert.ExecuteNonQuery()

                'Success
                blnResult = True



        Catch excError As Exception

            WriteLog(excError)

        End Try

        Return blnResult

    End Function



    ' ----------------------------------------------------------------------
    ' Name: GetGameReplayInformation
    ' Abstract: Gets all the informated needed for a replay
    ' ----------------------------------------------------------------------
    Public Function GetGameReplayInformation(intGameID As Integer) As udtGameReplays

        Dim udtGameReplay As New udtGameReplays

        Try

            Dim strSelect As String
            Dim cmdSelect As OleDb.OleDbCommand
            Dim drReader As OleDb.OleDbDataReader


            'Build the Select command. Never build Command with raw user input to prevent SQL injections
            strSelect = "SELECT" &
                            " intGameID " &
                            ",strPlayerName " &
                            ",blnComputerMovesFirst " &
                            ",intMoveIndex " &
                            ",intGroup " &
                            ",intSquare " &
                        "FROM VGameReplayInformation WHERE intGameID = " & intGameID

            'Make the Command instance
            cmdSelect = New OleDb.OleDbCommand(strSelect, m_conAdministrator)

            'Get the rows
            drReader = cmdSelect.ExecuteReader()
            With udtGameReplay

                .lintMoves = New List(Of Integer())

                drReader.Read()
                .intGameID = drReader(0) 'intGameID
                .strPlayerName = drReader(1) 'strPlayerName
                .blnComputerMovesFirst = drReader(2) 'blnComputerMovesFirst

                Do
                    .intMoveIndex = drReader(3) 'intMoveIndex
                    .lintMoves.Add({drReader(4), drReader(5)}) 'intGroup, intSquare
                Loop While drReader.Read() = True
            End With
        Catch excError As Exception

            WriteLog(excError)

        End Try

        Return udtGameReplay

    End Function



    ' ----------------------------------------------------------------------
    ' Name: SetGameOutcome
    ' Abstract: Sets the game's outcome
    ' ----------------------------------------------------------------------
    Public Function SetGameOutcome(intGameID As Integer, intGameOutcome As Integer) As Boolean

        Dim blnResult As Boolean = False

        Try

            Dim strInsert As String
            Dim cmdInsert As OleDb.OleDbCommand


            'Build the INSERT command. Never build Command with raw user input to prevent SQL injections
            strInsert = _
            "UPDATE TGames " + _
            "SET intGameOutcomeID = ? " + _
            "WHERE " + _
                "intGameID = ?"

            'Make the Command instance
            cmdInsert = New OleDb.OleDbCommand(strInsert, m_conAdministrator)

            'Add column values here instead of above to prevent SQL injection attacks
            With cmdInsert.Parameters

                .AddWithValue("1", intGameOutcome)
                .AddWithValue("2", intGameID)


            End With

            'Insert the row
            cmdInsert.ExecuteNonQuery()

            'Success
            blnResult = True



        Catch excError As Exception

            WriteLog(excError)

        End Try

        Return blnResult

    End Function

End Module

' -------------------------------------------------------------------------
' Module: modWriteLog
' Author: Patrick Callahan
' Abstract: Log messages to disk
'
' Revision        Owner   Changes:
' 2013/04/07      P.C.    For Book
' -------------------------------------------------------------------------

' -------------------------------------------------------------------------
' Options
' -------------------------------------------------------------------------
Option Explicit On


' -------------------------------------------------------------------------
' Imports
' -------------------------------------------------------------------------
Imports System
Imports System.IO


Public Module modWriteLog

    ' -------------------------------------------------------------------------
    '  Module constants
    ' -------------------------------------------------------------------------
    ' What log file should we use
    Private Const strLOG_FILE_EXTENSION As String = ".Log"

    ' -------------------------------------------------------------------------
    '  Module variables
    ' -------------------------------------------------------------------------
    Private m_strOldLogFilePath As String           ' Name of the last log file opened
    Private m_fsLogFile As FileStream = Nothing     ' File handle of the last log file opened


    ' -------------------------------------------------------------------------
    ' Name: WriteLog
    ' Abstract: Overload withd blnDisplay set to true
    ' -------------------------------------------------------------------------
    Public Sub WriteLog(ByVal excErrorToLog As Exception, _
               Optional ByVal blnDisplayWarning As Boolean = True)

        Try

            WriteLog(excErrorToLog.ToString(), blnDisplayWarning)

        Catch excError As Exception

            ' Display error message
            MessageBox.Show("Error:" & vbNewLine & excError.ToString(), _
                            Application.ProductName, _
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)

        End Try

    End Sub



    ' -------------------------------------------------------------------------
    ' Name: WriteLog
    ' Abstract: Write a message to the error log.
    ' -------------------------------------------------------------------------
    Public Sub WriteLog(ByVal strMessageToLog As String, _
               Optional ByVal blnDisplayWarning As Boolean = True)

        Try

            Dim fsLogFile As FileStream = Nothing
            Dim encConvertToByteArray As New System.Text.UTF8Encoding

            ' Warn the user?
            If blnDisplayWarning = True Then

                ' Yes( ProductName is set in AssemblyInfo )
                MessageBox.Show(strMessageToLog, Application.ProductName, _
                                MessageBoxButtons.OK, MessageBoxIcon.Warning)

            End If

            ' Append a date/time stamp
            strMessageToLog = (DateTime.Now).ToString("yyyy/MM/dd HH:mm:ss") _
                              & " - " & strMessageToLog & vbNewLine & _
                              vbNewLine

            ' Get a free file handle
            fsLogFile = GetLogFile()

            ' Is the file OK?
            If Not fsLogFile Is Nothing Then

                ' Yes, Log it
                fsLogFile.Write(encConvertToByteArray.GetBytes(strMessageToLog), _
                                0, strMessageToLog.Length)

                ' Flush the buffer so we can immediately see results in file.  Very important.
                ' Otherwise we have to wait for flush which might be when application closes
                ' or we get another error.  Waiting for the application to close may not be
                ' a good idea if the application is in a production environment (e.g. a web
                '  app running on a remote server)
                fsLogFile.Flush()

            End If

        Catch excError As Exception

            ' Display error message
            MessageBox.Show("Error:" & vbNewLine & excError.ToString(), _
                            Application.ProductName, _
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)

        End Try

    End Sub


    ' -------------------------------------------------------------------------
    ' Name: DeleteOldFiles
    ' Abstract: Delete any files older than 10 days.
    ' -------------------------------------------------------------------------
    Private Sub DeleteOldFiles()

        Try

            Dim strLogFilePath As String = ""
            Dim dirLogDirectory As DirectoryInfo = Nothing
            Dim dtmFileCreated As DateTime = Now
            Dim intDaysOld As Integer = 0

            ' Path
            strLogFilePath = Application.StartupPath & "\Log\"

            ' Look for any files
            dirLogDirectory = New DirectoryInfo(strLogFilePath)

            ' Are there any?
            For Each finLogFile As FileInfo _
                In dirLogDirectory.GetFiles("*" & strLOG_FILE_EXTENSION)

                ' When was the file created?
                dtmFileCreated = finLogFile.CreationTime

                ' How old is the file?
                intDaysOld = (dtmFileCreated.Subtract(DateTime.Now)).Days

                ' Is the file older than 10 days?
                If intDaysOld > 10 Then

                    ' Yes.  Delete it.
                    finLogFile.Delete()

                End If

            Next

        Catch excError As Exception

            ' Display error message
            MessageBox.Show("Error:" & vbNewLine & excError.ToString(), _
                            Application.ProductName, _
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)

        End Try

    End Sub


    ' -------------------------------------------------------------------------
    ' Name: GetLogFile
    ' Abstract: Open the log file for writing.  Use today's date as part of
    '           the file name.  Each day a new log file will be created.
    '           Makes debug easier.
    '           Use a filestream object so we can specify file read share
    '           during the open call.
    ' -------------------------------------------------------------------------
    Private Function GetLogFile() As FileStream

        Try
            Dim strToday As String = (DateTime.Now).ToString("yyyyMMdd")
            Dim strLogFilePath As String = ""

            ' Log everything in a log directory off of the current application directory
            strLogFilePath = Application.StartupPath & _
                             "\Log\" & strToday & strLOG_FILE_EXTENSION

            ' Is this a new day?
            If m_strOldLogFilePath <> strLogFilePath Then

                ' Save the log file name
                m_strOldLogFilePath = strLogFilePath

                ' Does the log directory exist?
                If Directory.Exists(Application.StartupPath & "\Log") = False Then

                    ' No, so create it
                    Directory.CreateDirectory(Application.StartupPath & "\Log")

                End If

                ' Close old log file( if there is one )
                If Not m_fsLogFile Is Nothing Then m_fsLogFile.Close()

                ' Delete old log files
                DeleteOldFiles()

                ' Does the file exist?
                If File.Exists(strLogFilePath) = False Then

                    ' No, create with shared read access so it can be read while application has it open
                    m_fsLogFile = New FileStream(strLogFilePath, FileMode.Create, _
                                                 FileAccess.Write, FileShare.Read)

                Else

                    ' Yes, append with shared read access so it can be read while application has it open
                    m_fsLogFile = New FileStream(strLogFilePath, FileMode.Append, _
                                                 FileAccess.Write, FileShare.Read)

                End If

            End If

        Catch excError As Exception

            ' Display error message
            MessageBox.Show("Error:" & vbNewLine & excError.ToString(), _
                            Application.ProductName, _
                            MessageBoxButtons.OK, MessageBoxIcon.Warning)

        End Try

        ' Return result
        Return m_fsLogFile

    End Function

End Module

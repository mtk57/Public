Attribute VB_Name = "SimpleTranscriptionLogging"
Option Explicit

Public Const DEBUG_LOG_ENABLED As Boolean = False

Private m_logFileNumber As Integer
Private m_logFileOpened As Boolean
Private m_logFilePath As String

Public Sub StartLog(ByVal filePath As String)
    If DEBUG_LOG_ENABLED = False Then Exit Sub
    If Len(filePath) = 0 Then Exit Sub

    On Error GoTo ErrHandler

    If m_logFileOpened Then
        Close #m_logFileNumber
        m_logFileOpened = False
    End If

    m_logFileNumber = FreeFile
    m_logFilePath = filePath

    Open m_logFilePath For Append As #m_logFileNumber
    m_logFileOpened = True

    Print #m_logFileNumber, String$(40, "-")
    Print #m_logFileNumber, Format$(Now, "yyyy-mm-dd HH:nn:ss") & " : ログ開始"
    Exit Sub

ErrHandler:
    m_logFileOpened = False
    m_logFilePath = vbNullString
End Sub

Public Sub WriteLog(ByVal message As String)
    If DEBUG_LOG_ENABLED = False Then Exit Sub
    If m_logFileOpened = False Then Exit Sub

    On Error Resume Next
    Print #m_logFileNumber, Format$(Now, "yyyy-mm-dd HH:nn:ss") & " : " & message
End Sub

Public Sub StopLog()
    If DEBUG_LOG_ENABLED = False Then Exit Sub

    On Error Resume Next
    If m_logFileOpened Then
        Print #m_logFileNumber, Format$(Now, "yyyy-mm-dd HH:nn:ss") & " : ログ終了"
        Close #m_logFileNumber
    End If

    m_logFileOpened = False
    m_logFilePath = vbNullString
End Sub

Public Function LogFilePath() As String
    LogFilePath = m_logFilePath
End Function

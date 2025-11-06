Attribute VB_Name = "ModuleMain"
Option Explicit

Private Const VER = "1.0.0"
Private Const DEBUG_LOG_ENABLED As Boolean = False
Private Const LOG_FILE_NAME As String = "xls2xlsx_debug.log"

Private Type ConversionSettings
    FolderPath As String
    IncludeSubfolders As Boolean
    KeepXls As Boolean
End Type

Public Sub RunConversion()
    Dim originalScreenUpdating As Boolean
    Dim originalDisplayAlerts As Boolean
    Dim originalEnableEvents As Boolean
    Dim originalCalculation As XlCalculation
    
    On Error GoTo ErrorHandler
    
    Dim settings As ConversionSettings
    settings = GetConversionSettings()
    
    PrepareDebugLog
    WriteDebugLog "処理開始: バージョン " & VER
    WriteDebugLog "設定: フォルダ=" & settings.FolderPath & ", サブフォルダ=" & CStr(settings.IncludeSubfolders) & ", xls保持=" & CStr(settings.KeepXls)
    
    originalScreenUpdating = Application.ScreenUpdating
    originalDisplayAlerts = Application.DisplayAlerts
    originalEnableEvents = Application.EnableEvents
    originalCalculation = Application.Calculation
    
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    
    ProcessFolder settings, settings.FolderPath
    
    WriteDebugLog "処理完了"
    MsgBox "変換が完了しました。", vbInformation
    GoTo Cleanup
ErrorHandler:
    WriteDebugLog "エラー: " & Err.Number & " - " & Err.Description
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical
Cleanup:
    Application.ScreenUpdating = originalScreenUpdating
    Application.DisplayAlerts = originalDisplayAlerts
    Application.EnableEvents = originalEnableEvents
    Application.Calculation = originalCalculation
End Sub

Private Function GetConversionSettings() As ConversionSettings
    Dim ws As Worksheet
    On Error GoTo NotFound
    Set ws = ThisWorkbook.Worksheets("main")
    On Error GoTo 0
    
    Dim result As ConversionSettings
    Dim rawFolder As String
    rawFolder = Trim$(CStr(ws.Range("B7").Value))
    If Len(rawFolder) = 0 Then
        Err.Raise vbObjectError + 1, , "フォルダパスを入力してください。"
    End If
    
    rawFolder = NormalizeFolderPath(rawFolder)
    If Dir$(rawFolder, vbDirectory) = vbNullString Then
        Err.Raise vbObjectError + 2, , "指定したフォルダが存在しません: " & rawFolder
    End If
    
    Dim rawIncludeSub As String
    rawIncludeSub = UCase$(Trim$(CStr(ws.Range("B8").Value)))
    If rawIncludeSub = vbNullString Or rawIncludeSub = "YES" Then
        result.IncludeSubfolders = True
    ElseIf rawIncludeSub = "NO" Then
        result.IncludeSubfolders = False
    Else
        Err.Raise vbObjectError + 3, , "サブフォルダも対象にはYESまたはNOを入力してください。"
    End If
    
    Dim rawKeepXls As String
    rawKeepXls = UCase$(Trim$(CStr(ws.Range("B9").Value)))
    If rawKeepXls = vbNullString Or rawKeepXls = "YES" Then
        result.KeepXls = True
    ElseIf rawKeepXls = "NO" Then
        result.KeepXls = False
    Else
        Err.Raise vbObjectError + 4, , "xlsは残すにはYESまたはNOを入力してください。"
    End If
    
    result.FolderPath = rawFolder
    GetConversionSettings = result
    Exit Function
NotFound:
    Err.Raise vbObjectError + 5, , "mainシートが見つかりません。"
End Function

Private Function NormalizeFolderPath(ByVal folderPath As String) As String
    Dim trimmedPath As String
    trimmedPath = Replace(folderPath, "/", "\")
    Do While Len(trimmedPath) > 3 And Right$(trimmedPath, 1) = "\"
        trimmedPath = Left$(trimmedPath, Len(trimmedPath) - 1)
    Loop
    If Len(trimmedPath) = 2 And Mid$(trimmedPath, 2, 1) = ":" Then
        trimmedPath = trimmedPath & "\"
    End If
    NormalizeFolderPath = trimmedPath
End Function

Private Sub ProcessFolder(ByVal settings As ConversionSettings, ByVal currentFolder As String)
    WriteDebugLog "フォルダ処理開始: " & currentFolder
    
    Dim fileName As String
    fileName = Dir$(currentFolder & "\*.xls", vbNormal)
    Do While fileName <> vbNullString
        On Error GoTo ConvertError
        ConvertSingleFile currentFolder & "\" & fileName, settings.KeepXls
        On Error GoTo 0
        fileName = Dir$
        GoTo ContinueLoop
ConvertError:
        WriteDebugLog "変換失敗: " & currentFolder & "\" & fileName & " - " & Err.Number & " - " & Err.Description
        Err.Clear
        fileName = Dir$
ContinueLoop:
    Loop
    
    If settings.IncludeSubfolders Then
        Dim subEntry As String
        subEntry = Dir$(currentFolder & "\*", vbDirectory)
        Do While subEntry <> vbNullString
            If subEntry <> "." And subEntry <> ".." Then
                Dim subPath As String
                subPath = currentFolder & "\" & subEntry
                Dim attr As Long
                On Error Resume Next
                attr = GetAttr(subPath)
                If Err.Number = 0 Then
                    On Error GoTo 0
                    If (attr And vbDirectory) = vbDirectory Then
                        ProcessFolder settings, subPath
                    End If
                Else
                    WriteDebugLog "サブフォルダ属性取得失敗: " & subPath & " - " & Err.Number & " - " & Err.Description
                    Err.Clear
                    On Error GoTo 0
                End If
            End If
            subEntry = Dir$
        Loop
    End If
    
    WriteDebugLog "フォルダ処理完了: " & currentFolder
End Sub

Private Sub ConvertSingleFile(ByVal sourcePath As String, ByVal keepXls As Boolean)
    WriteDebugLog "変換開始: " & sourcePath
    
    Dim targetPath As String
    targetPath = Left$(sourcePath, Len(sourcePath) - 4) & ".xlsx"
    
    Dim wb As Workbook
    Set wb = Application.Workbooks.Open(Filename:=sourcePath, UpdateLinks:=False, ReadOnly:=False)
    On Error GoTo CloseWorkbook
    
    wb.SaveAs Filename:=targetPath, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    WriteDebugLog "保存完了: " & targetPath
    
CloseWorkbook:
    Dim closeErr As Long
    Dim closeDesc As String
    closeErr = Err.Number
    closeDesc = Err.Description
    Err.Clear
    wb.Close SaveChanges:=False
    If closeErr <> 0 Then
        Err.Raise closeErr, , closeDesc
    End If
    
    If Not keepXls Then
        On Error GoTo DeleteError
        Kill sourcePath
        WriteDebugLog "削除完了: " & sourcePath
        On Error GoTo 0
    End If
    
    WriteDebugLog "変換完了: " & sourcePath
    Exit Sub
DeleteError:
    WriteDebugLog "削除失敗: " & sourcePath & " - " & Err.Number & " - " & Err.Description
    Err.Clear
End Sub

Private Sub PrepareDebugLog()
    If Not DEBUG_LOG_ENABLED Then
        Exit Sub
    End If
    
    Dim logPath As String
    logPath = GetLogPath()
    On Error Resume Next
    Kill logPath
    Dim handle As Integer
    handle = FreeFile
    Open logPath For Append As #handle
    Close #handle
    On Error GoTo 0
End Sub

Private Sub WriteDebugLog(ByVal message As String)
    If Not DEBUG_LOG_ENABLED Then
        Exit Sub
    End If
    
    Dim logPath As String
    logPath = GetLogPath()
    Dim handle As Integer
    handle = FreeFile
    On Error Resume Next
    Open logPath For Append As #handle
    If Err.Number = 0 Then
        Print #handle, Format$(Now, "yyyy-mm-dd hh:nn:ss") & " - " & message
    End If
    Close #handle
    On Error GoTo 0
End Sub

Private Function GetLogPath() As String
    If Len(ThisWorkbook.Path) > 0 Then
        GetLogPath = ThisWorkbook.Path & Application.PathSeparator & LOG_FILE_NAME
    Else
        GetLogPath = Application.DefaultFilePath & Application.PathSeparator & LOG_FILE_NAME
    End If
End Function

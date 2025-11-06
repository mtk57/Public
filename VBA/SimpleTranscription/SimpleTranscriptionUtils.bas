Attribute VB_Name = "SimpleTranscriptionUtils"
Option Explicit

Public Function ColumnLetterToNumber(ByVal columnLetters As String) As Long
    Dim letters As String
    letters = Trim$(UCase$(columnLetters))
    
    If Len(letters) = 0 Then Exit Function
    
    Dim i As Long
    Dim resultValue As Long
    For i = 1 To Len(letters)
        Dim codePoint As Long
        codePoint = Asc(Mid$(letters, i, 1)) - Asc("A") + 1
        If codePoint < 1 Or codePoint > 26 Then
            resultValue = 0
            Exit For
        End If
        resultValue = resultValue * 26 + codePoint
    Next i
    
    ColumnLetterToNumber = resultValue
End Function

Public Function NormalizeTextForCompare(ByVal valueData As Variant, ByVal caseSensitive As Boolean, ByVal distinguishWidth As Boolean) As String
    If IsError(valueData) Then Exit Function
    If IsEmpty(valueData) Then Exit Function
    If IsNull(valueData) Then Exit Function
    
    Dim workText As String
    workText = CStr(valueData)
    
    If distinguishWidth = False Then
        On Error Resume Next
        workText = StrConv(workText, vbNarrow)
        On Error GoTo 0
    End If
    
    If caseSensitive = False Then
        workText = UCase$(workText)
    End If
    
    NormalizeTextForCompare = Trim$(workText)
End Function

Public Function StringsMatch(ByVal baseValue As Variant, ByVal targetValue As Variant, ByVal allowPartial As Boolean, ByVal caseSensitive As Boolean, ByVal distinguishWidth As Boolean) As Boolean
    Dim baseText As String
    Dim targetText As String
    
    baseText = NormalizeTextForCompare(baseValue, caseSensitive, distinguishWidth)
    targetText = NormalizeTextForCompare(targetValue, caseSensitive, distinguishWidth)
    
    Dim compareMode As VbCompareMethod
    If caseSensitive Then
        compareMode = vbBinaryCompare
    Else
        compareMode = vbTextCompare
    End If
    
    If allowPartial Then
        If Len(baseText) = 0 Then
            StringsMatch = (Len(targetText) = 0)
        Else
            StringsMatch = (InStr(1, targetText, baseText, compareMode) > 0)
        End If
    Else
        StringsMatch = (StrComp(targetText, baseText, compareMode) = 0)
    End If
End Function

Public Function IsCellYellow(ByVal targetCell As Range) As Boolean
    If targetCell Is Nothing Then Exit Function
    
    Dim colorIndexValue As Variant
    On Error Resume Next
    colorIndexValue = targetCell.Interior.ColorIndex
    On Error GoTo 0
    
    If Not IsEmpty(colorIndexValue) Then
        If colorIndexValue = 6 Then
            IsCellYellow = True
            Exit Function
        End If
    End If
    
    Dim colorValue As Variant
    On Error Resume Next
    colorValue = targetCell.Interior.Color
    On Error GoTo 0
    
    If Not IsEmpty(colorValue) Then
        If colorValue = vbYellow Then
            IsCellYellow = True
        End If
    End If
End Function

Public Function IsBlankValue(ByVal valueData As Variant) As Boolean
    If IsError(valueData) Then Exit Function
    If IsNull(valueData) Then
        IsBlankValue = True
        Exit Function
    End If
    
    If IsEmpty(valueData) Then
        IsBlankValue = True
        Exit Function
    End If
    
    Select Case VarType(valueData)
        Case vbString
            IsBlankValue = (Len(Trim$(valueData)) = 0)
        Case Else
            IsBlankValue = False
    End Select
End Function

Public Function ResolveWorkbookPath(ByVal fileNameOrPath As String) As String
    Dim candidate As String
    candidate = Trim$(fileNameOrPath)
    If Len(candidate) = 0 Then Exit Function
    
    Dim fso As Object
    On Error Resume Next
    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error GoTo 0
    
    If fso Is Nothing Then
        ResolveWorkbookPath = candidate
        Exit Function
    End If
    
    If fso.FileExists(candidate) Then
        ResolveWorkbookPath = fso.GetAbsolutePathName(candidate)
        Exit Function
    End If
    
    Dim basePath As String
    basePath = ThisWorkbook.Path
    If Len(basePath) > 0 Then
        Dim combined As String
        combined = fso.BuildPath(basePath, candidate)
        If fso.FileExists(combined) Then
            ResolveWorkbookPath = fso.GetAbsolutePathName(combined)
            Exit Function
        End If
    End If
    
    ResolveWorkbookPath = vbNullString
End Function

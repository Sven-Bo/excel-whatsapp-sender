Attribute VB_Name = "mEmojis"
Option Explicit

#If VBA7 Then 'https://developer.apple.com/library/archive/documentation/System/Conceptual/ManPages_iPhoneOS/man3/iconv.3.html
    Private Declare PtrSafe Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
    Private Declare PtrSafe Function EmptyLongArray Lib "oleaut32" Alias "SafeArrayCreateVector" (Optional ByVal vt As VbVarType = vbLong, Optional ByVal lLow As Long = 0, Optional ByVal lCount As Long = 0) As Long()

#Else
    Private Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
    Private Declare Function EmptyLongArray Lib "oleaut32" Alias "SafeArrayCreateVector" (Optional ByVal vt As VbVarType = vbLong, Optional ByVal lLow As Long = 0, Optional ByVal lCount As Long = 0) As Long()
#End If

Public Function HasEmoji(nText As String) As Boolean
    On Error Resume Next
    
    ' If text is empty or null, return false
    If nText = "" Or IsNull(nText) Then
        HasEmoji = False
        Exit Function
    End If
    
    Dim iCharsU() As Long
    Dim c As Long
    
    iCharsU = StringToArrayAscU(nText)
    
    ' Check if array was created successfully
    If Not IsArray(iCharsU) Then
        HasEmoji = False
        Exit Function
    End If
    
    ' Check if array is empty
    If UBound(iCharsU) < 0 Then
        HasEmoji = False
        Exit Function
    End If
    
    For c = 0 To UBound(iCharsU)
        If IsEmojiChar(iCharsU(c)) Then
            HasEmoji = True
            Exit Function
        End If
    Next
    
    HasEmoji = False
    On Error GoTo 0
End Function

Private Function IsEmojiChar(ByVal nCharU As Long) As Boolean
    ' Based on emoji info: https://unicode.org/Public/emoji/12.1/emoji-data.txt
    Select Case nCharU
        Case &H203C, &H2049, &H2122, &H2139, &H2194 To &H2199, &H21A9 To &H21AA, &H231A To &H231B, &H2328, &H23CF, &H23E9 To &H23F3, &H23F8 To &H23FA, &H24C2, &H25AA To &H25AB, &H25B6, &H25C0, &H25FB To &H25FE, &H2600 To &H2604, &H260E, &H2611, &H2614 To &H2615, &H2618, &H261D, &H2620, &H2622 To &H2623, &H2626, &H262A, &H262E To &H262F, &H2638 To &H263A, &H2640, &H2642, &H2648 To &H2653, &H265F To &H2660, &H2663, &H2665 To &H2666, &H2668, &H267B, &H267E To &H267F, &H2692 To &H2697, &H2699, &H269B To &H269C, &H26A0 To &H26A1, &H26AA To &H26AB, &H26B0 To &H26B1, &H26BD To &H26BE, &H26C4 To &H26C5, &H26C8, &H26CE To &H26CF, &H26D1, &H26D3 To &H26D4, &H26E9 To &H26EA, &H26F0 To &H26F5, &H26F7 To &H26FA, &H26FD, &H2702, &H2705, &H2708 To &H270D, &H270F, &H2712, &H2714, &H2716, &H271D, &H2721, &H2728, &H2733 To &H2734, &H2744, _
            &H2747, &H274C, &H274E, &H2753 To &H2755, &H2757, &H2763 To &H2764, &H2795 To &H2797, &H27A1, &H27B0, &H27BF, &H2934 To &H2935, &H2B05 To &H2B07, &H2B1B To &H2B1C, &H2B50, &H2B55, &H3030, &H303D, &H3297, &H3299, &H1F004, &H1F0CF, &H1F170 To &H1F171, &H1F17E To &H1F17F, &H1F18E, &H1F191 To &H1F19A, &H1F1E6 To &H1F1FF, &H1F201 To &H1F202, &H1F21A, &H1F22F, &H1F232 To &H1F23A, &H1F250 To &H1F251, &H1F300 To &H1F321, &H1F324 To &H1F393, &H1F396 To &H1F397, &H1F399 To &H1F39B, &H1F39E To &H1F3F0, &H1F3F3 To &H1F3F5, &H1F3F7 To &H1F4FD, &H1F4FF To &H1F53D, &H1F549 To &H1F54E, &H1F550 To &H1F567, &H1F56F To &H1F570, &H1F573 To &H1F57A, &H1F587, &H1F58A To &H1F58D, &H1F590, &H1F595 To &H1F596, &H1F5A4 To &H1F5A5, &H1F5A8, &H1F5B1 To &H1F5B2, &H1F5BC, &H1F5C2 To &H1F5C4, &H1F5D1 To &H1F5D3, &H1F5DC To &H1F5DE, &H1F5E1, &H1F5E3, &H1F5E8, &H1F5EF, &H1F5F3, &H1F5FA To &H1F64F, _
            &H1F680 To &H1F6C5, &H1F6CB To &H1F6D2, &H1F6D5, &H1F6E0 To &H1F6E5, &H1F6E9, &H1F6EB To &H1F6EC, &H1F6F0, &H1F6F3 To &H1F6FA, &H1F7E0 To &H1F7EB, &H1F90D To &H1F93A, &H1F93C To &H1F945, &H1F947 To &H1F971, &H1F973 To &H1F976, &H1F97A To &H1F9A2, &H1F9A5 To &H1F9AA, &H1F9AE To &H1F9CA, &H1F9CD To &H1F9FF, &H1FA70 To &H1FA73, &H1FA78 To &H1FA7A, &H1FA80 To &H1FA82, &H1FA90
            
            IsEmojiChar = True
    End Select
End Function

Private Function StringToArrayAscU(ByRef nText As String) As Long()
    On Error Resume Next
    
    Dim iRet() As Long
    Dim iLen As Long
    Dim c As Long
    Dim iCharsW() As Integer
    Dim c2 As Long
    
    ' Handle empty or null text
    If nText = "" Or IsNull(nText) Then
        StringToArrayAscU = EmptyLongArray
        Exit Function
    End If
    
    iLen = Len(nText)
    If iLen Then
        ReDim iCharsW(iLen - 1)
        ReDim iRet(iLen - 1)
        
        ' Use error handling for memory operations
        CopyMemory iCharsW(0), ByVal StrPtr(nText), iLen * 2
        
        ' Check if memory operation succeeded
        If Err.Number <> 0 Then
            StringToArrayAscU = EmptyLongArray
            Exit Function
        End If
        
        c2 = -1
        For c = 0 To iLen - 1
            'Debug.Assert iCharsW(c) = AscW(Mid$(nText, c + 1))
            If ((iCharsW(c) < &HD800&) Or (iCharsW(c) > &HDBFF&)) And (iCharsW(c) > 0) Then
                c2 = c2 + 1
                iRet(c2) = iCharsW(c)
            ElseIf c < iLen - 1 Then
                c2 = c2 + 1
                iRet(c2) = &H10000 + (((iCharsW(c) And &H3FF&) * 1024&) Or (iCharsW(c + 1) And &H3FF&))
                c = c + 1
            End If
        Next
        
        If c2 >= 0 Then
            ReDim Preserve iRet(c2)
        Else
            ' If no valid characters were found, return an empty array
            iRet = EmptyLongArray
        End If
    Else
        iRet = EmptyLongArray
    End If
    StringToArrayAscU = iRet
    
    On Error GoTo 0
End Function

Private Function InStrEmoji(nText As String, nLength As Long, Optional StartPos As Long = 1) As Long
    On Error Resume Next
    
    ' Initialize return values
    InStrEmoji = 0
    nLength = 0
    
    ' Handle empty or null text
    If nText = "" Or IsNull(nText) Then
        Exit Function
    End If
    
    ' Handle invalid start position
    If StartPos < 1 Then StartPos = 1
    
    Dim iLen As Long
    Dim c As Long
    Dim iCharsW() As Integer
    Dim iLastEmojiChar As Long
    
    iLen = Len(nText)
    If iLen Then
        ReDim iCharsW(iLen - 1)
        
        ' Use error handling for memory operations
        CopyMemory iCharsW(0), ByVal StrPtr(nText), iLen * 2
        
        ' Check if memory operation succeeded
        If Err.Number <> 0 Then
            Exit Function
        End If
        
        For c = StartPos - 1 To iLen - 1
            If ((iCharsW(c) < &HD800&) Or (iCharsW(c) > &HDBFF&)) And (iCharsW(c) > 0) Then
                If IsEmojiChar(iCharsW(c)) Then
                    If InStrEmoji = 0 Then InStrEmoji = c + 1
                    iLastEmojiChar = c
                ElseIf InStrEmoji > 0 Then
                    Exit Function
                End If
            ElseIf c < (iLen - 1) Then
                If IsEmojiChar(&H10000 + (((iCharsW(c) And &H3FF&) * 1024&) Or (iCharsW(c + 1) And &H3FF&))) Then
                    If InStrEmoji = 0 Then InStrEmoji = c + 1
                    iLastEmojiChar = c + 1
                ElseIf InStrEmoji > 0 Then
                    Exit Function
                End If
                c = c + 1
            End If
        Next
    End If
    
    If InStrEmoji > 0 And iLastEmojiChar >= InStrEmoji - 1 Then
        nLength = iLastEmojiChar - InStrEmoji + 1
    Else
        nLength = 0
    End If
    
    On Error GoTo 0
End Function

Private Function ChrU(ByVal nCharCodeU As Long) As String
    On Error Resume Next
    
    Const cPOW10 As Long = 2 ^ 10
    
    If nCharCodeU <= &HFFFF& Then
        ChrU = ChrW$(nCharCodeU)
    Else
        ChrU = ChrW$(&HD800& + (nCharCodeU And &HFFFF&) \ cPOW10) & ChrW$(&HDC00& + (nCharCodeU And (cPOW10 - 1)))
    End If
    
    ' If there was an error, return empty string
    If Err.Number <> 0 Then
        ChrU = ""
    End If
    
    On Error GoTo 0
End Function


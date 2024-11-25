Attribute VB_Name = "BB2RTF"
Option Explicit
'
' BbCodes supported:
'   b, i, u
'   size={FontSize}
'   color={ColorNameOrValue}
'   font={FontName}
'   table={Col1_Width},{Col2_Width}, ...[;[TableLeftOffset],[ColumnLeftOffset]]
'   row={Col1_BackColor},{Col2_BackColor}, ...;<<col1_border>>;<<col2_border>>;...]
'      <<colN_border>>:=[BorderLeftColor] [BorderLeftWidth],[BorderTopColor] [BorderTopWidth],[BorderRightColor] [BorderRightWidth],[BorderBottomColor] [BorderBottomWidth]
'   col
'   url={Url}
'
' e.g.
'   [table=100,200,300]
'   [row]--A--[col]--B--[col]--C--[/row]
'   [row]1[col]test[col]value[/row]
'   [/table]
'
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long

Private Const STR_BBCODE_TAGS               As String = "[[]|[b]|[/b]|[i]|[/i]|[u]|[/u]|[size=|[/size]|[color=|[/color]|[url=|[/url]|[font=|[/font]|[right]|[/right]|[center]|[/center]|[table=|[/table]|[row]|[row=|[/row]|[col]"
Private Const STR_BBCODE_COLOR_NAMES        As String = "black|red|green|blue|cyan|magenta|yellow|grey|white"
Private Const STR_BBCODE_COLOR_RGBS         As String = "&H000000|&HFF0000|&H00FF00|&H0000FF|&H00FFFF|&HFF00FF|&HFFFF00|&HC0C0C0|&HFFFFFF"
Private Const STR_BBCODE_RTF_PREFIX         As String = "{\rtf1"
Private Const STR_BBCODE_RTF_SUFFIX         As String = "}"
Private Const STR_BBCODE_FONTS_PREFIX       As String = "{\fonttbl "
Private Const STR_BBCODE_FONTS_SUFFIX       As String = "}"
Private Const STR_BBCODE_COLORS_PREFIX      As String = "{\colortbl "
Private Const STR_BBCODE_COLORS_SUFFIX      As String = "}"

Private Enum UcsBBCodeTags
    ucsTagBracket
    ucsTagBold
    ucsTagBoldEnd
    ucsTagItalic
    ucsTagItalicEnd
    ucsTagUnderline
    ucsTagUnderlineEnd
    ucsTagSize
    ucsTagSizeEnd
    ucsTagColor
    ucsTagColorEnd
    ucsTagUrl
    ucsTagUrlEnd
    ucsTagFont
    ucsTagFontEnd
    ucsTagRight
    ucsTagRightEnd
    ucsTagCenter
    ucsTagCenterEnd
    ucsTagTable
    ucsTagTableEnd
    ucsTagRow
    ucsTagRowPlain
    ucsTagRowEnd
    ucsTagCol
End Enum

Private Sub PrintError(sFunc As String)
    Debug.Print sFunc, Err.Description
End Sub

Public Function BbCode2Rtf(sText As String, oFont As StdFont) As String
    Const FUNC_NAME     As String = "BbCode2Rtf"
    Dim vTags           As Variant
    Dim vColorNames     As Variant
    Dim vColorRGBs      As Variant
    Dim cStack          As Collection
    Dim cFonts          As Collection
    Dim cColors         As Collection
    Dim lPos            As Long
    Dim lTagStart       As Long
    Dim lTagEnd         As Long
    Dim sRetVal         As String
    Dim eTag            As UcsBBCodeTags
    Dim lCurSize        As Long
    Dim lCurColor       As Long
    Dim lCurFont        As Long
    Dim lTemp           As Long
    Dim lIdx            As Long
    Dim lJdx            As Long
    Dim sValue          As String
    Dim sParFmt         As String
    Dim vColumns        As Variant
    Dim vOffsets        As Variant
    Dim vSplit          As Variant
    Dim vBgColors       As Variant
    Dim vBorders        As Variant
    Dim vBorderDef      As Variant
    
    On Error GoTo EH
    '--- prepare lookup arrays
    vTags = Split(STR_BBCODE_TAGS, "|")
    vColorNames = Split(STR_BBCODE_COLOR_NAMES, "|")
    vColorRGBs = Split(STR_BBCODE_COLOR_RGBS, "|")
    '--- prepare collections
    Set cStack = New Collection
    Set cFonts = New Collection
    Set cColors = New Collection
    cFonts.Add oFont.name, oFont.name
    cColors.Add 0, "#0"
    '--- init default current values
    lCurFont = 1
    lCurSize = Round(oFont.size * 2)
    lCurColor = 0
    sRetVal = "\f" & lCurFont & "\fs" & lCurSize & "\cf" & lCurColor & vbCrLf
    '--- parse
    lPos = 1
    Do While lPos <= Len(sText)
        lTagStart = InStr(lPos, sText, "[")
        If lTagStart > 0 Then
            lTagEnd = InStr(lTagStart, sText, "]")
        Else
            lTagEnd = 0
            lTagStart = Len(sText) + 1
        End If
        sRetVal = sRetVal & RtfEscape(Mid$(sText, lPos, lTagStart - lPos), sParFmt)
        lPos = lTagStart + 1
        If lTagStart > 0 And lTagEnd > 0 Then
            For eTag = 0 To UBound(vTags)
                If LCase$(Mid$(sText, lTagStart, Len(vTags(eTag)))) = vTags(eTag) And _
                        (Right$(vTags(eTag), 1) <> "]" Or lTagEnd = lTagStart + Len(vTags(eTag))) - 1 Then
                    Exit For
                End If
            Next
            Select Case eTag
            Case ucsTagBracket
                sRetVal = sRetVal & RtfEscape("[", sParFmt)
            Case ucsTagBold
                sRetVal = sRetVal & "\b "
            Case ucsTagBoldEnd
                sRetVal = sRetVal & "\b0 "
            Case ucsTagItalic
                sRetVal = sRetVal & "\i "
            Case ucsTagItalicEnd
                sRetVal = sRetVal & "\i0 "
            Case ucsTagUnderline
                sRetVal = sRetVal & "\ul "
            Case ucsTagUnderlineEnd
                sRetVal = sRetVal & "\ul0 "
            Case ucsTagSize
                cStack.Add Array(ucsTagSize, lCurSize)
                sValue = Trim$(pvBbCodeGetValue(Mid$(sText, lTagStart, lTagEnd - lTagStart + 1)))
                If Right$(sValue, 1) = "%" Then
                    lTemp = Round(2 * C_Val(sValue) * oFont.size / 100, 0)
                Else
                    lTemp = Round(2 * C_Val(sValue), 0)
                End If
                If lTemp > 0 Then
                    lCurSize = lTemp
                    sRetVal = sRetVal & "\fs" & lCurSize & " "
                End If
            Case ucsTagSizeEnd
                For lIdx = cStack.count To 1 Step -1
                    If cStack(lIdx)(0) = ucsTagSize Then
                        lCurSize = cStack(lIdx)(1)
                        sRetVal = sRetVal & "\fs" & lCurSize & " "
                        cStack.Remove lIdx
                        Exit For
                    End If
                Next
                If lIdx < 1 Then
                    GoTo UnknownTag
                End If
            Case ucsTagColor
                sValue = LCase$(pvBbCodeGetValue(Mid$(sText, lTagStart, lTagEnd - lTagStart + 1)))
                cStack.Add Array(ucsTagColor, lCurColor)
                lCurColor = pvBbCodeGetColorIdx(sValue, cColors, vColorNames, vColorRGBs)
                sRetVal = sRetVal & "\cf" & lCurColor & " "
            Case ucsTagColorEnd
                For lIdx = cStack.count To 1 Step -1
                    If cStack(lIdx)(0) = ucsTagColor Then
                        lCurColor = cStack(lIdx)(1)
                        sRetVal = sRetVal & "\cf" & lCurColor & " "
                        cStack.Remove lIdx
                        Exit For
                    End If
                Next
                If lIdx < 1 Then
                    GoTo UnknownTag
                End If
            Case ucsTagUrl
                sValue = pvBbCodeGetValue(Mid$(sText, lTagStart, lTagEnd - lTagStart + 1))
                If InStr(sValue, ":") <= 2 Then
                    If InStr(sValue, "\") > 0 Then
                        sValue = "file:" & IIf(Left$(sValue, 2) <> "\\", "///", vbNullString) & Replace(sValue, "\", "/")
                    End If
                End If
                cStack.Add Array(ucsTagColor, lCurColor)
                lCurColor = pvBbCodeGetColorIdx("blue", cColors, vColorNames, vColorRGBs)
                sRetVal = sRetVal & "{\field{\*\fldinst{HYPERLINK """ & Replace(Replace(Replace(sValue, "\", "\\\\"), "{", "\{"), "}", "\}") & """}}{\fldrslt{\ul\cf" & lCurColor & " "
            Case ucsTagUrlEnd
                For lIdx = cStack.count To 1 Step -1
                    If cStack(lIdx)(0) = ucsTagColor Then
                        lCurColor = cStack(lIdx)(1)
                        cStack.Remove lIdx
                        Exit For
                    End If
                Next
                sRetVal = sRetVal & "}}}"
            Case ucsTagFont
                sValue = pvBbCodeGetValue(Mid$(sText, lTagStart, lTagEnd - lTagStart + 1))
                On Error Resume Next
                cFonts.Add sValue, sValue
                On Error GoTo EH
                For lIdx = 1 To cFonts.count
                    If LCase$(cFonts(lIdx)) = LCase$(sValue) Then
                        Exit For
                    End If
                Next
                cStack.Add Array(ucsTagFont, lCurFont)
                lCurFont = lIdx
                sRetVal = sRetVal & "\f" & lCurFont & " "
            Case ucsTagFontEnd
                For lIdx = cStack.count To 1 Step -1
                    If cStack(lIdx)(0) = ucsTagFont Then
                        lCurFont = cStack(lIdx)(1)
                        sRetVal = sRetVal & "\f" & lCurFont & " "
                        cStack.Remove lIdx
                        Exit For
                    End If
                Next
                If lIdx < 1 Then
                    GoTo UnknownTag
                End If
            Case ucsTagRight
                sParFmt = "\qr "
                sRetVal = sRetVal & sParFmt
            Case ucsTagRightEnd
                sParFmt = vbNullString
            Case ucsTagCenter
                sParFmt = "\qc "
                sRetVal = sRetVal & sParFmt
            Case ucsTagCenterEnd
                sParFmt = vbNullString
            Case ucsTagTable
                sValue = pvBbCodeGetValue(Mid$(sText, lTagStart, lTagEnd - lTagStart + 1))
                cStack.Add Array(ucsTagTable, sValue)
                sParFmt = vbNullString
                If Mid$(sText, lTagEnd + 1, 2) = vbCrLf Then
                    lTagEnd = lTagEnd + 2
                End If
            Case ucsTagTableEnd
                For lIdx = cStack.count To 1 Step -1
                    If cStack(lIdx)(0) = ucsTagTable Then
                        sRetVal = sRetVal & "\pard "
                        cStack.Remove lIdx
                        Exit For
                    End If
                Next
                If lIdx < 1 Then
                    GoTo UnknownTag
                End If
                If Mid$(sText, lTagEnd + 1, 2) = vbCrLf Then
                    lTagEnd = lTagEnd + 2
                End If
            Case ucsTagRow, ucsTagRowPlain
                For lIdx = cStack.count To 1 Step -1
                    If cStack(lIdx)(0) = ucsTagTable Then
                        sValue = cStack(lIdx)(1)
                        Exit For
                    End If
                Next
                If lIdx < 1 Then
                    GoTo UnknownTag
                End If
                '--- columns
                vSplit = Split(sValue, ";")
                vColumns = Split(At(vSplit, 0), ",")
                vOffsets = Split(At(vSplit, 1), ",")
                sRetVal = sRetVal & "\trowd\trgaph" & At(vOffsets, 0, "70") & "\trleft" & At(vOffsets, 1, "0")
                'sRetVal = sRetVal & "\trbrdrl\brdrs\brdrw10\brdrcf0 \trbrdrt\brdrs\brdrw10\brdrcf0 \trbrdrr\brdrs\brdrw10\brdrcf0 \trbrdrb\brdrs\brdrw10\brdrcf0" & vbCrLf
                '--- borders
                sValue = pvBbCodeGetValue(Mid$(sText, lTagStart, lTagEnd - lTagStart + 1))
                vSplit = Split(sValue, ";")
                vBgColors = Split(At(vSplit, 0), ",")
                lTemp = C_Lng(At(vOffsets, 1, "0"))
                For lIdx = 0 To UBound(vColumns)
                    If LenB(At(vBgColors, lIdx)) <> 0 Then
                        sRetVal = sRetVal & "\clcbpat" & pvBbCodeGetColorIdx(At(vBgColors, lIdx), cColors, vColorNames, vColorRGBs)
                    End If
                    lTemp = lTemp + C_Lng(vColumns(lIdx))
                    vBorders = Split(At(vSplit, lIdx + 1), ",")
                    For lJdx = 0 To 3
                        If LenB(At(vBorders, lJdx)) <> 0 Then
                            vBorderDef = Split(vBorders(lJdx), " ")
                            sRetVal = sRetVal & "\clbrdr" & Mid$("ltrb", lJdx + 1, 1) & "\brdrs\brdrw" & C_Lng(At(vBorderDef, 1, "10")) & "\brdrcf" & pvBbCodeGetColorIdx(At(vBorderDef, 0), cColors, vColorNames, vColorRGBs) & vbCrLf
                        End If
                    Next
                    sRetVal = sRetVal & "\cellx" & lTemp & vbCrLf
                Next
                sRetVal = sRetVal & "\pard\intbl "
            Case ucsTagRowEnd
                sRetVal = sRetVal & "\cell\row" & vbCrLf
                If Mid$(sText, lTagEnd + 1, 2) = vbCrLf Then
                    lTagEnd = lTagEnd + 2
                End If
            Case ucsTagCol
                sRetVal = sRetVal & "\cell\pard\intbl "
            Case Else
UnknownTag:
                '--- unknown tag
                sRetVal = sRetVal & RtfEscape(Mid$(sText, lTagStart, lTagEnd - lTagStart + 1), sParFmt)
            End Select
            lPos = lTagEnd + 1
        ElseIf lPos <= Len(sText) Then
            sRetVal = sRetVal & "["
        End If
    Loop
    BbCode2Rtf = STR_BBCODE_RTF_PREFIX
    '--- dump fonts table
    BbCode2Rtf = BbCode2Rtf & vbCrLf & STR_BBCODE_FONTS_PREFIX
    For lIdx = 1 To cFonts.count
        BbCode2Rtf = BbCode2Rtf & "{\f" & lIdx & "\fcharset204 " & cFonts(lIdx) & ";}"
    Next
    BbCode2Rtf = BbCode2Rtf & STR_BBCODE_FONTS_SUFFIX
    '--- dump colors table
    BbCode2Rtf = BbCode2Rtf & vbCrLf & STR_BBCODE_COLORS_PREFIX
    For lIdx = 1 To cColors.count
        lTemp = cColors(lIdx)
        If lTemp = 0 Then
            BbCode2Rtf = BbCode2Rtf & ";"
        Else
            BbCode2Rtf = BbCode2Rtf & _
                "\red" & ((lTemp \ &H10000) And &HFF&) & _
                "\green" & ((lTemp \ &H100&) And &HFF&) & _
                "\blue" & (lTemp And &HFF&) & ";"
        End If
    Next
    BbCode2Rtf = BbCode2Rtf & STR_BBCODE_COLORS_SUFFIX
    '--- insert body & suffix
    BbCode2Rtf = BbCode2Rtf & vbCrLf & sRetVal & vbCrLf & STR_BBCODE_RTF_SUFFIX
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Public Function RtfEscape(sText As String, Optional sParFmt As String) As String
'    RtfEscape = Replace(Replace(Replace(Replace(Replace(Replace(Replace(sText, "\" & vbCrLf, Chr(127)), "\", "\\"), "{", "\{"), "}", "\}"), vbTab, "\tab "), vbCrLf, "\par" & vbCrLf & sParFmt), Chr(127), "\line" & vbCrLf)
    Dim lSize       As Long
    Dim baBuffer()  As Byte
    Dim lIdx        As Long
    Dim nNext       As Byte
    
    lSize = WideCharToMultiByte(1251, 0, StrPtr(sText), Len(sText), 0, 0, 0, 0)
    If lSize > 0 Then
        ReDim baBuffer(0 To lSize - 1) As Byte
        Call WideCharToMultiByte(1251, 0, StrPtr(sText), Len(sText), VarPtr(baBuffer(0)), lSize, 0, 0)
        Do While lIdx <= UBound(baBuffer)
            If lIdx < UBound(baBuffer) Then
                nNext = baBuffer(lIdx + 1)
            Else
                nNext = 0
            End If
            Select Case baBuffer(lIdx)
            Case 92     ' "\"
                If nNext = 13 Then ' vbCr
                    RtfEscape = RtfEscape & "\line" & vbCrLf
                    lIdx = lIdx + 2
                Else
                    RtfEscape = RtfEscape & "\\"
                End If
            Case 123, 125 ' "{", "}"
                RtfEscape = RtfEscape & "\" & Chr$(baBuffer(lIdx))
            Case 9      ' vbTab
                RtfEscape = RtfEscape & "\tab "
            Case 10     ' vbLf
                RtfEscape = RtfEscape & "\par" & vbCrLf & sParFmt
            Case 13     ' vbCr
                RtfEscape = RtfEscape & "\par" & vbCrLf & sParFmt
                If nNext = 10 Then ' vbLf
                    lIdx = lIdx + 1
                End If
            Case Else
                If baBuffer(lIdx) < &H80 Then
                    RtfEscape = RtfEscape & Chr$(baBuffer(lIdx))
                Else
                    RtfEscape = RtfEscape & "\'" & Hex(baBuffer(lIdx))
                End If
            End Select
            lIdx = lIdx + 1
        Loop
    End If
End Function

Private Function pvBbCodeGetValue(sTag As String) As String
    If InStr(Mid(sTag, 2, Len(sTag) - 2), "=") > 0 Then
        pvBbCodeGetValue = Split(Mid(sTag, 2, Len(sTag) - 2), "=")(1)
    End If
End Function

Private Function pvBbCodeGetColorIdx(sValue As String, cColors As Collection, vColorNames As Variant, vColorRGBs As Variant) As Long
    Dim lTemp           As Long
    Dim lIdx            As Long
    
    If Left(sValue, 1) = "#" Then
        lTemp = C_Lng("&H" & Mid(sValue, 2))
    Else
        lTemp = 0
        For lIdx = 0 To UBound(vColorNames)
            If vColorNames(lIdx) = sValue Then
                lTemp = vColorRGBs(lIdx)
                Exit For
            End If
        Next
    End If
    On Error Resume Next
    cColors.Add lTemp, "#" & lTemp
    On Error GoTo 0
    For lIdx = 1 To cColors.count
        If cColors(lIdx) = lTemp Then
            pvBbCodeGetColorIdx = lIdx - 1
            Exit For
        End If
    Next
End Function

Public Function At(Data As Variant, ByVal Index As Long, Optional Default As String) As String
    On Error Resume Next
    At = Default
    At = C_Str(Data(Index))
    On Error GoTo 0
End Function

Public Function C_Str(Value As Variant) As String
    On Error Resume Next
    C_Str = CStr(Value)
    On Error GoTo 0
End Function

Public Function C_Lng(Value As Variant) As Long
    On Error Resume Next
    C_Lng = CLng(Value)
    On Error GoTo 0
End Function

Private Function C_Val(ByVal Value As String) As Double
    On Error Resume Next
    C_Val = Val(Replace(Replace(Value, "e", "_"), "d", "_"))
    On Error GoTo 0
End Function


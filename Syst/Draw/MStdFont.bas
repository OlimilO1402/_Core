Attribute VB_Name = "MStdFont"
Option Explicit

' v ' ############################## ' v '    StdFont    ' v ' ############################## ' v '
Public Function StdFont(ByVal Name As String, ByVal Size As Single, _
                        Optional ByVal bBold As Boolean = False, _
                        Optional ByVal bItalic As Boolean = False, _
                        Optional ByVal aWeight As Long = 1, _
                        Optional ByVal aCharSet As Long = 1, _
                        Optional ByVal bUnderline As Boolean = False, _
                        Optional ByVal bStrikeThru As Boolean = False) As StdFont
    Set StdFont = New StdFont
    With StdFont
        .Name = Name
        .Size = Size
        .Bold = bBold
        .Italic = bItalic
        .Weight = aWeight
        .Charset = aCharSet
        .Underline = bUnderline
        .Strikethrough = bStrikeThru
    End With
End Function

Function StdFontC(ByVal other As StdFont) As StdFont
    With other
        Set StdFontC = MNew.StdFont(.Name, .Size, .Bold, .Italic, .Weight, .Charset, .Underline, .Strikethrough)
    End With
End Function

Public Function StdFont_Clone(this As StdFont) As StdFont
    Set StdFont_Clone = New StdFont
    With StdFont_Clone
        .Name = this.Name
        .Bold = this.Bold
        .Charset = this.Charset
        .Italic = this.Italic
        .Size = this.Size
        .Strikethrough = this.Strikethrough
        .Underline = this.Underline
        .Weight = this.Weight
    End With
End Function

Public Function StdFont_Parse(s As String) As StdFont
    Set StdFont_Parse = New StdFont
    Dim sLines() As String: sLines = Split(s, "{")
    If VBA.Strings.LCase(Strings.Trim(sLines(0))) <> "stdfont" Then Exit Function
    sLines = Split(sLines(1), vbCrLf)
    Dim i As Long, sElems() As String, sKey As String, sVal As String
    For i = LBound(sLines) + 1 To UBound(sLines) - 1
        sElems() = Split(sLines(i), "=")
        sKey = VBA.LCase(VBA.Trim(sElems(0)))
        If UBound(sElems) > 0 Then
            sVal = VBA.Trim(sElems(1))
            With StdFont_Parse
                Select Case sKey
                Case "name":                   .Name = sVal
                Case "bold":                   .Bold = Boolean_Parse(sVal)
                Case "charset":             .Charset = CInt(sVal)
                Case "italic":               .Italic = Boolean_Parse(sVal)
                Case "size":                   .Size = CCur(sVal)
                Case "strikethrough": .Strikethrough = Boolean_Parse(sVal)
                Case "underline":         .Underline = Boolean_Parse(sVal)
                Case "weight":               .Weight = CInt(sVal)
                End Select
            End With
        End If
    Next
End Function

Public Function StdFont_ToStr(this As StdFont) As String
    Dim s As String: s = "StdFont {" & vbCrLf
    s = s & "    Name          = " & this.Name & vbCrLf
    s = s & "    Bold          = " & this.Bold & vbCrLf
    s = s & "    Charset       = " & this.Charset & vbCrLf
    s = s & "    Italic        = " & this.Italic & vbCrLf
    s = s & "    Size          = " & this.Size & vbCrLf
    s = s & "    Strikethrough = " & this.Strikethrough & vbCrLf
    s = s & "    Underline     = " & this.Underline & vbCrLf
    s = s & "    Weight        = " & this.Weight & vbCrLf
    s = s & "}"
    StdFont_ToStr = s
End Function

Private Function Boolean_Parse(ByVal sVal As String) As Boolean
'also see in MString
    Dim b As Boolean
    sVal = VBA.LCase(sVal)
    Select Case True
    Case sVal = "falsch": b = False
    Case sVal = "false":  b = False
    Case sVal = "nein":   b = False
    Case sVal = "no":     b = False
    Case sVal = "wahr":   b = True
    Case sVal = "true":   b = True
    Case sVal = "yes":    b = True
    Case sVal = "ja":     b = True
    End Select
    Boolean_Parse = b
End Function

Public Sub StdFont_Assign(fntDst As StdFont, fntSrc As StdFont)
    With fntDst
        .Name = fntSrc.Name
        .Size = fntSrc.Size
        .Bold = fntSrc.Bold
        .Italic = fntSrc.Italic
        .Weight = fntSrc.Weight
        .Charset = fntSrc.Charset
        .Underline = fntSrc.Underline
        .Strikethrough = fntSrc.Strikethrough
    End With
End Sub


Attribute VB_Name = "modRTF"
Global Heading As ColorConstants
Global Color As ColorConstants

Public Sub Text(Text As String, Optional Colour As ColorConstants, Optional Bold As Boolean, Optional Italic As Boolean, Optional Underline As Boolean, Optional Size As Integer, Optional Alignment As AlignmentConstants, Optional Font As String)
On Error Resume Next
' Based off Chris Stratford's Mercury Chat code
If Text = "" Then Exit Sub
With frmChat.bandeja
    .SelStart = Len(.Text)
    .SelLength = Len(.Text)
    .SelBold = Bold
    .SelItalic = Italic
    .SelUnderline = Underline
    .SelFontSize = Size
    .SelAlignment = Alignment
    .SelColor = Colour
    .SelText = Text
    .SelStart = Len(.Text)
    .SelLength = 0
End With
End Sub

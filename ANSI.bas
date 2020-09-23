Attribute VB_Name = "ANSI"
Option Explicit

Public Const Sv = "[S"      'Save Cursor Position
Public Const Rs = "[U"      'Restore Cursor Position
Public Const Sva = "[7"     'Save Cursor Position & Attributes
Public Const Rsa = "[8"     'Restore Cursor Position & Attributes

Public Const iReset = "[C"  'Reset All Terminal Settings
Public Const eLW = "[7H"    'Enable Line Wrapping
Public Const dLW = "[7l"    'Disable Line Wrapping

Public Const Scr = "[R"     'Scroll Entire Screen
Public Const sDn = "[D"     'Scroll Display Down 1 Line
Public Const sUp = "[M"     'Scroll Display Up 1 Line

Public Const sTab = "[H"    'Sets Tab At Current Position
Public Const cTab = "[G"    'Clear Tab At Current Position
Public Const xTab = "[3G"   'Clear All Tabs

Public Const xEOL = "[K"    'Erases Current Cursor Position To End Of Current Line
Public Const xSOL = "[1K"   'Erases Current Cursor Position To Start Of Current Line
Public Const xLine = "[2K"  'Erases Current Line
Public Const xDn = "[J"     'Erases Screen From Current Line Down To Bottom Of Screen
Public Const xUp = "[1J"    'Erases Screen From Current Line Up To Top Of Screen
Public Const xCLS = "[2J"   'Clear Entire Screen And move Cursor To Home

Public Const PrSc = "[2J"   'Print Screen
Public Const PrLn = "[2J"   'Print Line
Public Const sPrLog = "[2J" 'Start Printing Log
Public Const xPrLog = "[2J" 'Stop printing Log

Dim i As Integer

' $$$ -[ Begin Simple ANSI Conversion ]- $$$

Function Mv(xPos As Integer, yPos As Integer)

Mv = "[" & xPos & ";" & yPos & "H"

End Function

Function Up(UpCount As Integer)

Up = "[" & UpCount & "A"

End Function

Function Dn(DnCount As Integer)

Dn = "[" & DnCount & "B"

End Function

Function Fw(FwCount As Integer)

Fw = "[" & FwCount & "C"

End Function

Function Bk(BkCount As Integer)

Bk = "[" & BkCount & "D"

End Function

Function Scroll(sStart As Integer, sEnd As Integer)

Scroll = "[" & sStart & ";" & sEnd & "R"

End Function

Function DefKey(KeyCodeOrKeyASCII As String, iVal As String)

If IsNumeric(KeyCodeOrKeyASCII) = False Then
    DefKey = "[" & Asc(KeyCodeOrKeyASCII) & ";" & Chr(34) & iVal & Chr(34) & "P"
Else
    If KeyCodeOrKeyASCII > 255 Then
        MsgBox "Invalid ASCII Value Provided For The DefKey Function."
    End If
        DefKey = "[" & KeyCodeOrKeyASCII & ";" & iVal & "P"
End If
   
End Function

' $$$ -[ End Simple ANSI Conversion ]- $$$

Public Function Box(xPos As Integer, yPos As Integer, Width As Integer, Height As Integer, Border As Integer)

Dim TopLeft As String: Dim TopRight As String
Dim BottomLeft As String: Dim BottomRight As String
Dim AcrossTop As String: Dim AcrossBottom As String
Dim DownLeft As String: Dim DownRight As String
Dim TheBoxData

If Border = 1 Then
    TopLeft = Chr(218): TopRight = Chr(191)
    BottomLeft = Chr(192): BottomRight = Chr(217)
    AcrossTop = Chr(196): AcrossBottom = Chr(196)
    DownLeft = Chr(179): DownRight = Chr(179)
End If

If Border = 2 Then
    TopLeft = Chr(201): TopRight = Chr(187)
    BottomLeft = Chr(200): BottomRight = Chr(188)
    AcrossTop = Chr(205): AcrossBottom = Chr(205)
    DownLeft = Chr(186): DownRight = Chr(186)
End If

If Border = 3 Then
    TopLeft = Chr(220): TopRight = Chr(220)
    BottomLeft = Chr(223): BottomRight = Chr(223)
    AcrossTop = Chr(220): AcrossBottom = Chr(223)
    DownLeft = Chr(221): DownRight = Chr(222)
End If

If Border = 4 Then
    TopLeft = Chr(219): TopRight = Chr(219)
    BottomLeft = Chr(219): BottomRight = Chr(219)
    AcrossTop = Chr(219): AcrossBottom = Chr(219)
    DownLeft = Chr(219): DownRight = Chr(219)
End If

    TheBoxData = TheBoxData & Mv(xPos, yPos) & TopLeft

        For i = 1 To Width - 2
            TheBoxData = TheBoxData & AcrossTop
            DoEvents
        Next

    TheBoxData = TheBoxData & Mv(xPos, Width + yPos - 1) & TopRight

    TheBoxData = TheBoxData & Mv(Height + xPos - 1, yPos) & BottomLeft

        For i = 1 To Width - 2
            TheBoxData = TheBoxData & AcrossBottom
            DoEvents
        Next

    TheBoxData = TheBoxData & Mv(Height + xPos - 1, yPos + Width - 1) & BottomRight

        For i = 1 To Height - 2
            TheBoxData = TheBoxData & Mv(xPos + i, yPos) & DownLeft
            DoEvents
        Next

        For i = 1 To Height - 2
            TheBoxData = TheBoxData & Mv(xPos + i, Width + yPos - 1) & DownRight
            DoEvents
        Next

Box = TheBoxData

End Function

Public Function Color(Forground As Integer, Background As Integer, Attr As String)

Forground = Forground + 40: Background = Background + 30

If Attr = 3 Then Attr = 333: If Attr = 6 Then Attr = 333

Color = "[" & Forground & ";" & Background
If Not Attr = 333 Then Color = Color & ";" & Attr
Color = Color & "M"

End Function

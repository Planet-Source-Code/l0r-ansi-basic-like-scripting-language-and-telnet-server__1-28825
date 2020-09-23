VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3375
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   4335
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   15.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   14.063
   ScaleMode       =   4  'Character
   ScaleWidth      =   36.125
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Play Script"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2880
      Width           =   4095
   End
   Begin VB.TextBox txtInData 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1815
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   1
      Top             =   480
      Width           =   4095
   End
   Begin VB.TextBox txtTheData 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   2400
      Width           =   4095
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Left            =   3240
      Top             =   3360
   End
   Begin MSWinsockLib.Winsock Sock1 
      Left            =   2760
      Top             =   3360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim TheData As String: Dim rxData As String: Dim tmpLoop As Integer
Dim i As Integer: Dim Scriptlet As String

Dim tmpParam0 As String: Dim tmpParam1 As String
Dim tmpParam2 As String: Dim tmpParam3 As String
Dim tmpParam4 As String: Dim tmpParam5 As String
Dim tmpParam6 As String: Dim tmpParam7 As String
Dim tmpParam8 As String: Dim tmpParam9 As String

Dim Params As String

Private Sub Command3_Click()

Open App.Path & "\script.ini" For Input As #1

    Do While Not EOF(1)

    Line Input #1, Scriptlet

    If Scriptlet = "" Then GoTo LoopIt

    Scriptlet = Replace(Scriptlet, vbTab, "")
    
    If UCase(Mid(Scriptlet, 1, 3)) = "CLS" Then GoTo eval
    If UCase(Mid(Scriptlet, 1, 3)) = "END" Then GoTo eval
    
        Params = Replace(Scriptlet, Split(Scriptlet)(0), "")

eval:

        Select Case UCase(Split(Scriptlet, " ")(0))
            
            Case "CLS"
                Sock1.SendData xCLS

            Case "BOX"
                tmpParam0 = Split(Params, ", ")(0)
                tmpParam1 = Split(Params, ", ")(1)
                tmpParam2 = Split(Params, ", ")(2)
                tmpParam3 = Split(Params, ", ")(3)
                tmpParam4 = Split(Params, ", ")(4)
                Sock1.SendData Box(Val(tmpParam0), Val(tmpParam1), Val(tmpParam2), Val(tmpParam3), Val(tmpParam4))

            Case "WAIT"
                Timer1.Enabled = True
                Timer1.Interval = Val(Params)

                While Timer1.Interval > 0
                    DoEvents
                Wend

                Timer1.Enabled = False

            Case "MOVE"
                tmpParam0 = Split(Params, ", ")(0)
                tmpParam1 = Split(Params, ", ")(1)
                Sock1.SendData Mv(Val(tmpParam0), Val(tmpParam1))

            Case "PRINT"
                Params = Mid(Params, 3, Len(Params) - 3)
                Params = Replace(Params, "$IP", Sock1.RemoteHostIP)
                Params = Replace(Params, "$HOST", Sock1.RemoteHost)
                Params = Replace(Params, "$TIME", Time)
                Params = Replace(Params, "$DATE", Date)
                Sock1.SendData Params
            
            Case "COLOR"
                tmpParam0 = Split(Params, ", ")(0)
                tmpParam1 = Split(Params, ", ")(1)
                tmpParam2 = Split(Params, ", ")(2)
                Sock1.SendData Color(Val(tmpParam0), Val(tmpParam1), Val(tmpParam2))

            Case "TERMINATE"
                Sock1.Close

            Case "END"
                GoTo EndIt

            Case "GO_UP"
                Sock1.SendData Up(0)
            
            Case "GO_DN"
                Sock1.SendData Dn(0)
            
            Case "GO_FW"
                Sock1.SendData Fw(0)
            
            Case "GO_BK"
                Sock1.SendData Bk(0)

            Case "SCROLL"
                For i = 1 To Params
                    Sock1.SendData sUp
                    DoEvents
                Next i

            Case "SAVE_CURSOR"
                Sock1.SendData Sv
                
            Case "RESTORE_CURSOR"
                Sock1.SendData Rs
            
            Case "SAVE_ATTRIB"
                Sock1.SendData Sva
            
            Case "RESTORE_ATTRIB"
                Sock1.SendData Rsa

            Case "RESET"
                Sock1.SendData iReset
                
            Case "WRAP"
                If UCase(Params) = "ON" Then Sock1.SendData eLW
                If UCase(Params) = "OFF" Then Sock1.SendData dLW
    
        End Select

LoopIt:
    Loop

EndIt:
    Close #1

End Sub

Private Sub Form_Load()

Sock1.LocalPort = 23
Sock1.Listen

End Sub

Private Sub Sock1_DataArrival(ByVal bytesTotal As Long)

Sock1.GetData TheData: txtTheData.Text = TheData

rxData = rxData & TheData: txtInData = rxData

With Sock1: Select Case UCase(TheData)

    Case "[A"
        .SendData Up(0)

    Case "[B"
        .SendData Dn(0)

    Case "[C"
        .SendData Fw(0)

    Case "[D"
        .SendData Bk(0)

    'Case "[P"
    '        .SendData Mv(1, 1)
    '    For tmpLoop = 1 To 26
    '        Wait (0.5)
    '        .SendData sUp
    '    Next tmpLoop
        
    Case "[Q"
        .SendData sDn
        
    Case "[R"
        .SendData sUp
        
    Case "[S"
        .SendData xCLS

    Case Chr(8)
        .SendData " " & Bk(0)
    
    Case Chr(127)
        .SendData xLine

 End Select: End With
  
If TheData = Chr(10) Or TheData = Chr(13) Or TheData = vbCrLf Then

    With Sock1: Select Case UCase(Replace(rxData, vbCrLf, ""))

        Case "CLEAR"
          .SendData xCLS
        
    End Select: End With

    rxData = "": txtInData = ""

End If

End Sub

Private Sub Sock1_Close()

Sock1.LocalPort = 23
Sock1.Listen
Cls

End Sub

Private Sub Sock1_ConnectionRequest(ByVal requestID As Long)

Sock1.Close
Sock1.Accept requestID

Print "Current IP: " & Sock1.RemoteHostIP

End Sub

Private Sub Sock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)

Sock1.Close

End Sub

Private Sub Timer1_Timer()
    Timer1.Interval = 0
End Sub

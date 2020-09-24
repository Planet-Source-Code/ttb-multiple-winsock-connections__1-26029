VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Server"
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5700
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   5700
   StartUpPosition =   3  'Windows Default
   Begin MSWinsockLib.Winsock main 
      Left            =   3000
      Top             =   2400
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Start"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4800
      TabIndex        =   3
      Top             =   2520
      Width           =   855
   End
   Begin VB.TextBox txtchat 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2415
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   0
      Width           =   5655
   End
   Begin VB.TextBox txtmsg 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Text            =   "ttb 0wn3r1z3rz"
      Top             =   2520
      Width           =   3735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3840
      TabIndex        =   0
      Top             =   2520
      Width           =   855
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Index           =   0
      Left            =   2400
      Top             =   1800
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
Public users As Long
Option Explicit
Private Sub Command1_Click()
Dim i As Integer
For i = 0 To 200
If Winsock1(i).State = sckConnected Then Winsock1(i).SendData "Server :" & txtmsg.Text
Next i
txtchat.Text = txtchat.Text & vbCrLf & "Server: " & txtmsg.Text

End Sub

Private Sub Command2_Click()

If Command2.Caption = "Start" Then
main.LocalPort = 806
main.Listen
Command2.Caption = "Stop"
Exit Sub
End If
If Command2.Caption = "Stop" Then
main.Close
Command2.Caption = "Start"
Exit Sub
End If
End Sub

Private Sub Form_Load()
Dim i As Integer
For i = 1 To 200
 Load Winsock1(i)
Next i
End Sub





Private Sub main_ConnectionRequest(ByVal requestID As Long)
Dim i As Integer, j As Integer
For i = 0 To 200
If Winsock1(i).State = sckClosed Then
users = users + 1
Winsock1(i).Accept requestID
For j = 0 To 200
If Winsock1(j).State = sckConnected Then Winsock1(j).SendData "NUSER" & users
Next j
Exit Sub
End If
Next i
main.Close
main.Listen
End Sub

Private Sub Winsock1_Close(Index As Integer)
users = users - 1
Dim i As Integer
For i = 0 To 200
If Winsock1(i).State = sckConnected Then Winsock1(i).SendData "DUSER" & users
Next i
End Sub

Private Sub Winsock1_DataArrival(Index As Integer, ByVal bytesTotal As Long)
Dim i As Integer
Dim strdata As String
Winsock1(Index).GetData strdata
txtchat.Text = txtchat.Text & vbCrLf & strdata
For i = 0 To 200
If Winsock1(i).State = sckConnected Then Winsock1(i).SendData strdata
Next i

End Sub


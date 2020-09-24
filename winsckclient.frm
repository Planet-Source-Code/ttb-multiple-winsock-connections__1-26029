VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Client"
   ClientHeight    =   3690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4785
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3690
   ScaleWidth      =   4785
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
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
      Left            =   600
      MaxLength       =   20
      TabIndex        =   7
      Text            =   "Client"
      Top             =   480
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Connect"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   5
      Top             =   240
      Width           =   975
   End
   Begin VB.TextBox Text1 
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
      Left            =   840
      TabIndex        =   4
      Top             =   120
      Width           =   2775
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
      TabIndex        =   2
      Top             =   3360
      Width           =   855
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
      Top             =   3360
      Width           =   3735
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
      TabIndex        =   0
      Top             =   840
      Width           =   4695
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   2400
      Top             =   2280
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label4 
      Caption         =   "0"
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
      Left            =   3480
      TabIndex        =   9
      Top             =   480
      Width           =   135
   End
   Begin VB.Label Label3 
      Caption         =   "Users:"
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
      Left            =   2640
      TabIndex        =   8
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Nick:"
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
      Left            =   0
      TabIndex        =   6
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Server:"
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
      Left            =   0
      TabIndex        =   3
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If txtmsg.Text = "" Then Exit Sub
If Winsock1.State = sckConnected Then
Winsock1.SendData Text2.Text & ": " & txtmsg.Text

End If
End Sub

Private Sub Command2_Click()
If Command2.Caption = "Connect" Then
If Text1.Text = "" Then Exit Sub
Winsock1.RemoteHost = Text1
Winsock1.RemotePort = 806
Winsock1.Connect
Command2.Caption = "Disconnect"
Exit Sub
End If
If Command2.Caption = "Disconnect" Then
Winsock1.Close
Command2.Caption = "Connect"
Exit Sub
End If
End Sub

Private Sub Winsock1_Connect()
Winsock1.SendData Text2.Text & " connected: " & Winsock1.LocalIP
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
Dim strdata As String
Winsock1.GetData strdata
If Left(strdata, 5) = "DUSER" Then
Label4.Caption = Right(strdata, Len(strdata) - 5)
Exit Sub
End If
If Left(strdata, 5) = "NUSER" Then
Label4.Caption = Right(strdata, Len(strdata) - 5)
Exit Sub
End If
txtchat.Text = txtchat & vbCrLf & strdata
End Sub


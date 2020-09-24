VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "                                   NetBus Scanner"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4785
   FillColor       =   &H00FFFFFF&
   BeginProperty Font 
      Name            =   "Comic Sans MS"
      Size            =   6
      Charset         =   162
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Trojan Scanner.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   4785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   3840
      Top             =   1080
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Disconnect 
      Caption         =   "DisConnect"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   6
      Top             =   360
      Width           =   1455
   End
   Begin VB.CommandButton Connect 
      Caption         =   "Connect"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   5
      Top             =   120
      Width           =   1455
   End
   Begin VB.TextBox finds 
      BackColor       =   &H80000001&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   735
      Left            =   720
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   480
      Width           =   2415
   End
   Begin VB.TextBox Port 
      BackColor       =   &H80000001&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   285
      Left            =   720
      TabIndex        =   1
      Top             =   480
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.TextBox IP 
      BackColor       =   &H80000001&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000005&
      Height          =   285
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   2415
   End
   Begin VB.Label Explane2 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "DigitSmall"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   840
      TabIndex        =   9
      Top             =   1560
      Width           =   3015
   End
   Begin VB.Label Explane3 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   3240
      TabIndex        =   8
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "Finds:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Explane1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "DigitSmall"
         Size            =   7.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   720
      TabIndex        =   4
      Top             =   1320
      Width           =   3015
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "IP:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000001&
      Height          =   255
      Left            =   360
      TabIndex        =   2
      Top             =   120
      Width           =   255
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False





' First you must prepare your Form like this;
' 2 Button, 3 Text Box, 3 Label and 1 Winsock .


'Labels:
'1 Label for; when you are connected you must see the Connection IP.
'1 Label for; when you are connected you must see the Connection Port.
'1 Label for; Connected or Disconnected.
' Buttons:
' 1 Button for; Connect to Remote Host.
' 1 Button Disconnect from Remote Host.
' Text Boxs:
' 1 Text Box for; entering your Ip address.
' 1 Text Box for; using Port.
' And the latest Text Box for; when you run this code Helping the commands.


' Some Changes for This Code;

' If you can't see any WinSock in Your Visual Basic Tool Box then click right in you tool box and choose WinSock,
' Change Names:
' Text1.Text -----> IP.Text
' Text2.Text -----> Port.Text (Visible =False)
' Text3.Text -----> finds.Text (mutiline & locked = True)
' Command1 -----> Connect
' Command2 -----> Disconnect
' Label1 ------> Explane1
' Label2 ------> Explane2
' Label3 ------> Explane3


' Cool, Time To Coding...
' Starting....

' ©opy®ight 2001 - Yigit Aktan
' I hate School, Hack The Planet!
' My Personel e mail: yigitaktan@yahoo.com

Private Sub Connect_Click()
Call Disconnect_Click

'in here if you blank you IP address Text area then error message for you.
If IP.Text = "" Then
MsgBox "You Must entering your IP Address..."
Else


'if Error...
On Error GoTo Error
'Now Winsock reading our IP & Port Text, then connected.
Winsock1.Connect IP.Text, Port.Text


'Then finds.Text says it to us.
If finds.Text = "" Then
finds.Text = Winsock1.RemoteHost & " Can't find any NetBus Server..."
'Then closing our If :-D
End If
End If
Exit Sub


'If any Error our Program then Error Message Box and Close our Program.
Error:
MsgBox "Pls try Again!"
End
'End Sub, The End Of The Connect Button Code.
End Sub

'Preapare our Disconnect Button Code...
Private Sub Disconnect_Click()
'When we click this button, winsock is closing.
Winsock1.Close
'When Winscok closed change our label Caption and write winsock closing.
Explane3.Caption = "Disconnected."
'When Winsock closed then reset all. (IP,Port,)
Explane1.Caption = ""
finds.Text = ""
Explane2.Caption = ""
End Sub

'Form is Loading with our choosing.
Private Sub Form_Load()
'IP.Text = Our IP.
IP.Text = Winsock1.LocalIP
'Port.Text = NetBus Port (12345)
'Don't Forget; Our Port.Text is invisible. :-D
Port.Text = "12345"
'When we open our Program then all right message for Label1 = Explane1...
Explane1.Caption = "Running, time to Connection."
End Sub

' ***********************************************************
'Only your choosing (IF YOU WANT).

Private Sub IP_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
KeyAscii = 0
Call Connect_Click
Call Disconnect_Click
End If
End Sub
' ***********************************************************

' Prepare for Winsock...
Private Sub Winsock1_Connect()
'When connect the Server then right Connected message to Explane3.
Explane3.Caption = "Connected."
'Entering the Connection IP to Explane1.
Explane1.Caption = " Ip: " & Winsock1.RemoteHost
'Entering the Connection Port to Explane2
Explane2.Caption = "Port: " & Winsock1.RemotePort
End Sub

'getting data with Winsock...
Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)
'this is string.
Dim Data As String
'Winsock is connect the Server and get Data there.
Winsock1.GetData Data, vbString
'blank for Data.
finds.Text = ""
'When Winsock get Data from Remote Host then writing it to finds.Text.
finds.Text = finds.Text & Data
finds.SelStart = Len(finds)


'If finds.Text is blank the we understand we can't connect to remote Host then we understand
'can't find NetBus Server.
If finds.Text = "" Then
finds.Text = Winsock1.RemoteHost & " Can't find any NetBus Server..."
End If
End Sub



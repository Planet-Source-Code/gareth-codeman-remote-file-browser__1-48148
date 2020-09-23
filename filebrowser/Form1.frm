VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form1 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "                                                                                     REMOTE FILE BROWSER"
   ClientHeight    =   8220
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   11910
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   11910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   1680
      Top             =   6120
   End
   Begin VB.Frame Frame2 
      Height          =   1575
      Left            =   4800
      TabIndex        =   8
      Top             =   6600
      Width           =   6975
      Begin VB.CommandButton Command3 
         Caption         =   "change drive"
         Height          =   255
         Left            =   5760
         TabIndex        =   14
         Top             =   240
         Width           =   1095
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Width           =   5535
      End
      Begin VB.CommandButton Command4 
         Caption         =   "change dir"
         Height          =   255
         Left            =   5760
         TabIndex        =   12
         Top             =   600
         Width           =   1095
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Top             =   600
         Width           =   5535
      End
      Begin VB.CommandButton Command6 
         Caption         =   "open prog"
         Height          =   255
         Left            =   5760
         TabIndex        =   10
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Top             =   960
         Width           =   5535
      End
   End
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   6600
      Width           =   4455
      Begin VB.CommandButton Command7 
         Caption         =   "X"
         Height          =   255
         Left            =   3240
         TabIndex        =   7
         ToolTipText     =   "Disconnect"
         Top             =   240
         Width           =   375
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   6
         Text            =   "127.0.0.1"
         Top             =   240
         Width           =   1335
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1560
         TabIndex        =   5
         Text            =   "1987"
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "O"
         Height          =   255
         Left            =   2760
         TabIndex        =   4
         ToolTipText     =   "Connect"
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Clear"
      Height          =   255
      Left            =   1080
      TabIndex        =   2
      Top             =   120
      Width           =   855
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   1200
      Top             =   720
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.CommandButton Command1 
      Caption         =   "get drives"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000003&
      Height          =   6030
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   11655
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "MADE BY GARETH"
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   7800
      Width           =   4455
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   7320
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
On Error Resume Next
Winsock1.SendData "getdrives" & "~nothing"
End Sub

Private Sub Command2_Click()
On Error Resume Next
Winsock1.Connect Text1, Text2
End Sub


Private Sub Command3_Click()
On Error Resume Next
Winsock1.SendData "cdrive" & "~" & Text3
End Sub

Private Sub Command5_Click()
List1.Clear
End Sub

Private Sub Command6_Click()
On Error Resume Next
Winsock1.SendData "opennorm" & "~" & Text5
End Sub


Private Sub Command7_Click()
On Error Resume Next
Winsock1.SendData "byebye~server"
DoEvents
Winsock1.Close

End Sub

Private Sub Form_Terminate()
On Error Resume Next
Winsock1.SendData "byebye~server"
DoEvents
Winsock1.Close
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Winsock1.SendData "byebye~server"
DoEvents
Winsock1.Close
End Sub

Private Sub List1_Click()
Text3.Text = List1.Text
Text4.Text = List1.Text
Text5.Text = List1.Text
End Sub

Private Sub Timer1_Timer()
If Winsock1.State = sckConnected Then Label1.Caption = "Connected"
If Winsock1.State <> sckConnected Then Label1.Caption = "Not Connected"
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)


Dim xx As String
Winsock1.GetData xx
DoEvents
List1.AddItem xx
DoEvents

End Sub

Private Sub Command4_Click()
On Error Resume Next
Winsock1.SendData "cdir" & "~" & Text4
End Sub

VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Paltalk Whisper Bot  Source Made by nWo"
   ClientHeight    =   2190
   ClientLeft      =   3615
   ClientTop       =   3990
   ClientWidth     =   5730
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   5730
   Begin VB.TextBox txtRoom 
      Height          =   375
      Left            =   720
      TabIndex        =   4
      Top             =   3600
      Width           =   3975
   End
   Begin VB.ComboBox Combo1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00FF0000&
      Height          =   315
      ItemData        =   "Form1.frx":0000
      Left            =   1680
      List            =   "Form1.frx":0025
      TabIndex        =   3
      Text            =   "Enter Name of person here"
      Top             =   240
      Width           =   3735
   End
   Begin VB.TextBox Text3 
      Height          =   1215
      Left            =   1680
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Text            =   "Form1.frx":00CD
      Top             =   660
      Width           =   3735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clear"
      Height          =   315
      Left            =   60
      TabIndex        =   1
      Top             =   660
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Send"
      Height          =   315
      Left            =   60
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
  Dim richedita As Long
    Dim chatclass As Long
    Dim Button As Long
    Dim sContruct As String
   sConstruct = "/w " & Combo1.Text & ": " & Text3.Text 'build the whisper string
    txtRoom = (GetCaption$(FindAnyWindow&(Me, "Group")))
    chatclass = FindWindow("#32770", txtRoom)
    chatclass = FindWindowEx(chatclass, 0&, "#32770", vbNullString)
    richedita = FindWindowEx(chatclass, 0&, "richedit20a", vbNullString)
    richedita = FindWindowEx(chatclass, richedita, "richedit20a", vbNullString)
    Call SendMessageByString(richedita, WM_SETTEXT, 0&, sConstruct)
    chatclass = FindWindow("#32770", txtRoom)
    Button = FindWindowEx(chatclass, 0&, "button", "button1")
     On Error Resume Next
AppActivate "Group - Voice Conference"
On Error Resume Next
AppActivate "Group - Private Voice Conference"
    Call SendMessageLong(Button, WM_KEYDOWN, VK_SPACE, 0&)
    Call SendMessageLong(Button, WM_KEYUP, VK_SPACE, 0&)
   
   txtRoom.Text = ""
   Text3.Text = ""
End Sub

Private Sub Command2_Click()
Text3.Text = ""
End Sub

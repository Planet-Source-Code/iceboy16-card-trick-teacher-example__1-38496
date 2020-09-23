VERSION 5.00
Begin VB.Form Main 
   Caption         =   "Card Trick Teacher Example"
   ClientHeight    =   3270
   ClientLeft      =   1500
   ClientTop       =   1230
   ClientWidth     =   5400
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   3270
   ScaleWidth      =   5400
   Begin VB.CommandButton Command3 
      Caption         =   "SHOW ME A REAL TRICK DUDE!"
      Height          =   255
      Left            =   2640
      TabIndex        =   9
      Top             =   3000
      Width           =   2655
   End
   Begin VB.CommandButton Command2 
      Height          =   135
      Left            =   5280
      TabIndex        =   7
      Top             =   0
      Width           =   135
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Next >>"
      Height          =   375
      Left            =   0
      TabIndex        =   6
      Top             =   2880
      Width           =   1095
   End
   Begin VB.PictureBox Picture5 
      AutoSize        =   -1  'True
      Height          =   1440
      Left            =   4320
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   1380
      ScaleWidth      =   1020
      TabIndex        =   4
      Top             =   240
      Width           =   1080
   End
   Begin VB.PictureBox Picture4 
      AutoSize        =   -1  'True
      Height          =   1440
      Left            =   2160
      Picture         =   "Form1.frx":4992
      ScaleHeight     =   1380
      ScaleWidth      =   1005
      TabIndex        =   3
      Top             =   240
      Width           =   1065
   End
   Begin VB.PictureBox Picture3 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   1425
      Left            =   3240
      Picture         =   "Form1.frx":9324
      ScaleHeight     =   1365
      ScaleWidth      =   1005
      TabIndex        =   2
      Top             =   240
      Width           =   1065
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      Height          =   1440
      Left            =   1080
      Picture         =   "Form1.frx":DBEA
      ScaleHeight     =   1380
      ScaleWidth      =   1005
      TabIndex        =   1
      Top             =   240
      Width           =   1065
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      Height          =   1440
      Left            =   0
      Picture         =   "Form1.frx":1257C
      ScaleHeight     =   1380
      ScaleWidth      =   1005
      TabIndex        =   0
      Top             =   240
      Width           =   1065
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "CTT - Make your CTT's freely!"
      Height          =   195
      Left            =   1560
      TabIndex        =   8
      Top             =   0
      Width           =   2130
   End
   Begin VB.Image Image10 
      Height          =   495
      Left            =   2160
      Top             =   1440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image Image9 
      Height          =   495
      Left            =   2160
      Top             =   1440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image Image8 
      Height          =   495
      Left            =   2160
      Top             =   1440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image Image7 
      Height          =   495
      Left            =   2160
      Top             =   1440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image Image6 
      Height          =   495
      Left            =   2160
      Top             =   1440
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Image Image3 
      Height          =   1365
      Left            =   240
      Picture         =   "Form1.frx":16F0E
      Top             =   1680
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Image Image1 
      Height          =   1380
      Left            =   1440
      Picture         =   "Form1.frx":1B7D4
      Top             =   1680
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Image Image2 
      Height          =   1380
      Left            =   2760
      Picture         =   "Form1.frx":20166
      Top             =   1680
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.Image Image4 
      Height          =   1380
      Left            =   3960
      Picture         =   "Form1.frx":24AF8
      Top             =   1680
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.Label Label1 
      Caption         =   $"Form1.frx":2948A
      Height          =   1215
      Left            =   0
      TabIndex        =   5
      Top             =   1680
      Width           =   5415
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a
Private Sub Command1_Click()
a = a + 1
If a = 1 Then
Picture1.Top = 0
Label1.Caption = "Say he chooses the 6 of hearts. Know let him look at his card and while he does so, you flip the rest of the cards. Now the most of the symbols look down. Then let him put the card in and make a quick mix up. Now as you put the cards down you'll notice that a card (in this case 6 of hearts) is looking up. Yes! That's his card...   Feel free to make your own card trick teacher with tricks you know... following my idea."
Command1.Caption = "Restart"
Image6.Picture = Picture2.Picture
Image7.Picture = Picture3.Picture
Image8.Picture = Picture4.Picture
Image9.Picture = Picture5.Picture
Picture2.Picture = Image1.Picture
Picture3.Picture = Image2.Picture
Picture4.Picture = Image3.Picture
Picture5.Picture = Image4.Picture
Else
Image1.Picture = Picture2.Picture
Image2.Picture = Picture3.Picture
Image3.Picture = Picture4.Picture
Image4.Picture = Picture5.Picture
Picture2.Picture = Image6.Picture
Picture3.Picture = Image7.Picture
Picture4.Picture = Image8.Picture
Picture5.Picture = Image9.Picture
Picture1.Top = 240
Label1.Caption = "First get a bunch of cards. Then put the cards in right placement so the most symbols look up. For example in 6 of hearts the 4 big hearts look up and the 2 look down. The the 7 of spades 5 spades look up and 2 down. Do that with every card you use and then offer one card to someone. Let him choose one and NEVER look at it. This way he'll trust you. Select 'Next'"
Command1.Caption = "Next >>"
a = 0
End If
End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Command3_Click()
Form1.Show
Me.Hide
End Sub

Private Sub Form_Load()
a = 0
End Sub

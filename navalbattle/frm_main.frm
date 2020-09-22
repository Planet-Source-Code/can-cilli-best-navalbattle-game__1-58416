VERSION 5.00
Begin VB.Form frm_main 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Howæbout NavalBattle"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4260
   Icon            =   "frm_main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   183
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   284
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2745
      Left            =   0
      ScaleHeight     =   181
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   282
      TabIndex        =   0
      Top             =   0
      Width           =   4260
      Begin VB.PictureBox marqback 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   240
         ScaleHeight     =   65
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   249
         TabIndex        =   6
         Top             =   1560
         Width           =   3735
         Begin VB.Label marqview 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   $"frm_main.frx":0CCA
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   1935
            Left            =   0
            TabIndex        =   7
            Top             =   900
            Width           =   3735
         End
      End
      Begin VB.Timer marquee 
         Enabled         =   0   'False
         Interval        =   250
         Left            =   0
         Top             =   0
      End
      Begin VB.CommandButton lannet 
         BackColor       =   &H0080C0FF&
         Caption         =   "Lan/Net Game"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2685
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   1170
         Width           =   1215
      End
      Begin VB.CommandButton twoplayer 
         BackColor       =   &H0080C0FF&
         Caption         =   "Two on Same"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1485
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1170
         Width           =   1215
      End
      Begin VB.CommandButton singleplayer 
         BackColor       =   &H0080C0FF&
         Caption         =   "Single Player"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   285
         MaskColor       =   &H8000000F&
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1170
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "®"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   5.25
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Left            =   1320
         TabIndex        =   5
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Howæbout   NavalBattle - version 1.2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   375
         Left            =   0
         TabIndex        =   1
         Top             =   750
         Width           =   4215
      End
   End
End
Attribute VB_Name = "frm_main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    'initializing window graphics
    graphics Me.pic, "botback"
    graphics Me.pic, "bigtopback"
    graphics Me.pic, "mainicon"
    graphics Me.marqback, "botback"
    marquee.Enabled = True
    marqview = "All logos, items and codes (except Sid Meier's Civilization II ship images, Tim Miron's BitBlt tutorial codes, Peter Verburgh's IP finder code and Huw Wilkins' Network neighbourhood codes which can be downloaded freely from Planetsourcecode.com) are originally made and designed by Can Çilli (jOhNChiLLy). Turkish patent laws reserved all rights of this product. Special thanks to game tester Emre Kul (CooL). For any help please contact me at acaza@superonline.com"
End Sub

Private Sub Form_Unload(Cancel As Integer)
    'if you close the window, then application will close
    On Error Resume Next
    If frm_lannet.sock.state <> sckClosed Then frm_lannet.sock.SendData "hnb|close": frm_lannet.sock.Close
    If frm_game.sock.state <> sckClosed Then frm_game.sock.SendData "hnb|close": frm_game.sock.Close
    deletebuffer
    End
End Sub

Private Sub lannet_Click()
    'opens lan/net game window
    frm_lannet.Show: Me.Hide
End Sub

Private Sub marquee_Timer()
    If marqview.Top <= -124 Then marqview.Top = 65: DoEvents
    marqview.Top = marqview.Top - 3
End Sub

Private Sub twoplayer_Click()
    'opens new game window for two player
    playernum = 2
    frm_newgame.Show: Me.Hide
End Sub

Private Sub singleplayer_Click()
    'opens new game window for single player
    playernum = 1
    frm_newgame.Show: Me.Hide
End Sub

VERSION 5.00
Begin VB.Form frm_newgame 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Howæbout NavalBattle"
   ClientHeight    =   2295
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3375
   Icon            =   "frm_newgame.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   153
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   225
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox back 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   3075
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   1
      ToolTipText     =   "Back"
      Top             =   60
      Width           =   240
   End
   Begin VB.PictureBox pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2295
      Left            =   0
      ScaleHeight     =   151
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   223
      TabIndex        =   0
      Top             =   0
      Width           =   3375
      Begin VB.CommandButton start 
         BackColor       =   &H0080C0FF&
         Caption         =   "Start Game"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1800
         Width           =   2895
      End
      Begin VB.TextBox name2 
         BackColor       =   &H00C0E0FF&
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
         Height          =   285
         Left            =   720
         MaxLength       =   22
         TabIndex        =   4
         Top             =   1320
         Width           =   2295
      End
      Begin VB.TextBox name1 
         BackColor       =   &H00C0E0FF&
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
         Height          =   285
         Left            =   720
         MaxLength       =   22
         TabIndex        =   3
         Top             =   960
         Width           =   2295
      End
      Begin VB.Label plyrname 
         BackStyle       =   0  'Transparent
         Caption         =   "Enter Player's Names"
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
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   600
         Width           =   3135
      End
   End
End
Attribute VB_Name = "frm_newgame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim backvar, lastname(2)

Private Sub Form_Load()
    'initializing window graphics
    graphics Me.pic, "botback"
    graphics Me.pic, "smalltopback"
    graphics Me.pic, "smallicon"
    graphics Me.pic, "player1"
    'initializing window display
    If playernum = 2 Then
        frm_newgame.Height = 2670
        start.Top = 120
        name2.Visible = True
        graphics Me.pic, "player2"
        plyrname = "Enter Players' Names"
    End If
    If playernum = 1 Or playernum = 3 Then
        frm_newgame.Height = 2270
        start.Top = 94
        name2.Visible = False
        plyrname = "Enter Your Name"
        If playernum = 3 Then start.Caption = "Continue"
    End If
    back.Picture = LoadResPicture(201, 0): backvar = False
    For i = 1 To 2: lastname(i) = GetSetting("Navalbattle", "Settings", "Player" & i): Next i
    name1 = IIf(lastname(1) <> "", lastname(1), name1): name2 = IIf(lastname(2) <> "", lastname(2), name2)
End Sub

Private Sub start_Click()
    'manages the function of the frame
    If Not Len(name1) <> 0 Then Beep: name1.SetFocus: Exit Sub
    If playernum = 2 And Not Len(name2) <> 0 Then Beep: name2.SetFocus: Exit Sub
    If InStr(1, name1, "|") > 0 Or InStr(1, name2, "|") Then Beep: plyrname = "Invalid character | at position " & IIf(InStr(1, name1, "|") > 0, InStr(1, name1, "|"), InStr(1, name2, "|")): Exit Sub
    If playernum = 2 And name2 = name1 Then Beep: name2.SetFocus: Exit Sub
    If name1 <> lastname(1) Then SaveSetting "Navalbattle", "Settings", "Player1", name1
    If name2 <> lastname(2) Then SaveSetting "Navalbattle", "Settings", "Player2", name2
    If playernum < 3 Then mymode = "": frm_game.Show: Unload Me
    playername = name1: Unload Me
End Sub

Private Sub Form_Activate()
    'activates player name to change on load
    name1.SetFocus: name1.SelStart = Len(name1)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'controls who wanted to close the window,
    'if user wants it ends the application, if code wants closes the window
    If UnloadMode = 1 Or playernum = 3 Then Unload Me Else deletebuffer: End
End Sub

Private Sub pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'if mouse is out the back button change the graphic
    If backvar = True Then back.Picture = LoadResPicture(201, 0): backvar = False
End Sub

Private Sub back_Click()
    'back button is clicked, close window
    If playernum = 3 Then Unload Me: Exit Sub
    Unload Me: frm_main.Show
End Sub

Private Sub back_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'if mouse is over the back button change the graphic
    back.Picture = LoadResPicture(202, 0): backvar = True
End Sub


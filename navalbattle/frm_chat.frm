VERSION 5.00
Begin VB.Form frm_chat 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Chat Window"
   ClientHeight    =   4050
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7575
   Icon            =   "frm_chat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   270
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   505
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox back 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   7275
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
      Height          =   4050
      Left            =   0
      ScaleHeight     =   268
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   503
      TabIndex        =   0
      Top             =   0
      Width           =   7575
      Begin VB.TextBox messages 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   2895
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   600
         Width           =   7110
      End
      Begin VB.CommandButton send 
         BackColor       =   &H0080C0FF&
         Caption         =   "Send"
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
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   3600
         Width           =   615
      End
      Begin VB.TextBox msg 
         BackColor       =   &H00C0E0FF&
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   9
            Charset         =   162
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   285
         Left            =   240
         TabIndex        =   2
         Top             =   3600
         Width           =   6375
      End
   End
End
Attribute VB_Name = "frm_chat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim backvar, lastmsg

Private Sub Form_Activate()
    'switchs off message blinker
    If frm_game.msgwink.Enabled = True Then frm_game.msgwink.Enabled = False: frm_game.msgicon.Visible = False: frm_game.infoicon.Visible = True
    messages.SelStart = Len(messages): msg.SetFocus
End Sub

Private Sub Form_Load()
    'initializing window graphics
    graphics Me.pic, "botback"
    graphics Me.pic, "smalltopback"
    graphics Me.pic, "smallicon"
    back.Picture = LoadResPicture(201, 0): backvar = False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'hides window instead of closing it, to save the messages
    If UnloadMode = 0 Then Me.Hide: Cancel = True
End Sub

Private Sub msg_KeyUp(KeyCode As Integer, Shift As Integer)
    'message entry commands
    If KeyCode = 13 Then send_Click: Exit Sub
    If KeyCode = 38 And Not Len(msg) > 0 Then msg = lastmsg: msg.SelStart = Len(msg): Exit Sub
End Sub

Private Sub send_Click()
    'sends the message
    If Not Len(msg) > 0 Then Beep: Exit Sub
    If frm_game.sock.state = sckClosed Or frm_game.sock.state = sckClosing Then
        addmsg "* Other side has left game.": Beep
        msg = "": msg.SetFocus: Exit Sub
    End If
    msg = Replace(msg, "|", "¦")
    frm_game.sock.SendData "hnb|message|" & msg
    addmsg "<" & playername & "> " & msg
    lastmsg = msg: msg = "": msg.SetFocus
End Sub

Public Sub addmsg(txt As String)
    'adds message to the messages list
    messages = messages & IIf(Len(messages) > 0, vbCrLf, "") & txt
    messages.SelStart = Len(messages)
End Sub

Private Sub pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'if mouse is out the back button change the graphic
    If backvar = True Then back.Picture = LoadResPicture(201, 0): backvar = False
End Sub

Private Sub back_Click()
    'back button is clicked, hide window
    Me.Hide
End Sub

Private Sub back_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'if mouse is over the back button change the graphic
    back.Picture = LoadResPicture(202, 0): backvar = True
End Sub

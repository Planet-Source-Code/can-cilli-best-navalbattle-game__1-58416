VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frm_game 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Howæbout NavalBattle"
   ClientHeight    =   7245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7275
   Icon            =   "frm_game.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   483
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   485
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox back 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   6975
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
      Height          =   7245
      Left            =   0
      ScaleHeight     =   481
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   483
      TabIndex        =   0
      Top             =   0
      Width           =   7275
      Begin VB.PictureBox infoicon 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   360
         ScaleHeight     =   20
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   20
         TabIndex        =   29
         Top             =   1200
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.Timer looserships 
         Enabled         =   0   'False
         Interval        =   2000
         Left            =   1800
         Top             =   0
      End
      Begin VB.Timer ordertime 
         Enabled         =   0   'False
         Interval        =   1500
         Left            =   1440
         Top             =   0
      End
      Begin VB.PictureBox wait 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   5025
         Left            =   360
         ScaleHeight     =   333
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   436
         TabIndex        =   26
         Top             =   1560
         Visible         =   0   'False
         Width           =   6570
         Begin VB.Label click 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Click, if you are ready!"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404040&
            Height          =   255
            Left            =   0
            TabIndex        =   28
            Top             =   4440
            Width           =   6555
         End
         Begin VB.Label waittext 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "It's other player's turn!"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   11.25
               Charset         =   162
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   735
            Left            =   0
            TabIndex        =   27
            Top             =   2280
            Width           =   6555
         End
      End
      Begin VB.PictureBox missileview 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   585
         Left            =   360
         ScaleHeight     =   37
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   79
         TabIndex        =   22
         Top             =   6000
         Width           =   1215
         Begin VB.Line missileline 
            BorderColor     =   &H00FFFFFF&
            Index           =   1
            X1              =   0
            X2              =   80
            Y1              =   12
            Y2              =   12
         End
         Begin VB.Label missiletext 
            Alignment       =   2  'Center
            BackColor       =   &H00808080&
            Caption         =   "missiles"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   162
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   180
            Left            =   0
            TabIndex        =   25
            Top             =   0
            Width           =   1215
         End
         Begin VB.Line missileline 
            BorderColor     =   &H00808080&
            Index           =   0
            X1              =   61
            X2              =   61
            Y1              =   13
            Y2              =   40
         End
         Begin VB.Label missilenum 
            BackStyle       =   0  'Transparent
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   162
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   0
            Left            =   1020
            TabIndex        =   24
            Top             =   360
            Width           =   255
         End
         Begin VB.Label missilenum0 
            BackStyle       =   0  'Transparent
            Caption         =   "¥"
            BeginProperty Font 
               Name            =   "Symbol"
               Size            =   11.25
               Charset         =   2
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   0
            Left            =   975
            TabIndex        =   23
            Top             =   135
            Width           =   255
         End
      End
      Begin VB.PictureBox shipview 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   4170
         Left            =   360
         ScaleHeight     =   276
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   79
         TabIndex        =   12
         Top             =   1560
         Width           =   1215
         Begin VB.PictureBox shipnums 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   3960
            Left            =   900
            ScaleHeight     =   264
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   20
            TabIndex        =   15
            Top             =   195
            Width           =   300
            Begin VB.Label ship1 
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   162
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   495
               Index           =   0
               Left            =   105
               TabIndex        =   21
               Top             =   75
               Visible         =   0   'False
               Width           =   255
            End
            Begin VB.Label ship2 
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   162
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   495
               Index           =   0
               Left            =   105
               TabIndex        =   20
               Top             =   765
               Visible         =   0   'False
               Width           =   255
            End
            Begin VB.Label ship3 
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   162
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   495
               Index           =   0
               Left            =   105
               TabIndex        =   19
               Top             =   1410
               Visible         =   0   'False
               Width           =   255
            End
            Begin VB.Label ship4 
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   162
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   495
               Index           =   0
               Left            =   105
               TabIndex        =   18
               Top             =   2115
               Visible         =   0   'False
               Width           =   255
            End
            Begin VB.Label ship5 
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   162
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   495
               Index           =   0
               Left            =   105
               TabIndex        =   17
               Top             =   2760
               Visible         =   0   'False
               Width           =   255
            End
            Begin VB.Label ship6 
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   8.25
                  Charset         =   162
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   495
               Index           =   0
               Left            =   105
               TabIndex        =   16
               Top             =   3435
               Visible         =   0   'False
               Width           =   255
            End
         End
         Begin VB.PictureBox ships 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   3960
            Index           =   0
            Left            =   0
            ScaleHeight     =   264
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   59
            TabIndex        =   14
            Top             =   195
            Width           =   885
         End
         Begin VB.Label shiptext 
            Alignment       =   2  'Center
            BackColor       =   &H00808080&
            Caption         =   "ships"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   162
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   180
            Left            =   0
            TabIndex        =   13
            Top             =   0
            Width           =   1230
         End
      End
      Begin VB.Timer connection 
         Enabled         =   0   'False
         Interval        =   15000
         Left            =   1080
         Top             =   0
      End
      Begin VB.PictureBox mapview 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   5025
         Left            =   2100
         ScaleHeight     =   333
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   320
         TabIndex        =   9
         Top             =   1560
         Width           =   4830
         Begin VB.PictureBox map 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   4800
            Index           =   0
            Left            =   0
            MouseIcon       =   "frm_game.frx":0CCA
            MousePointer    =   99  'Custom
            ScaleHeight     =   320
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   320
            TabIndex        =   11
            Top             =   195
            Width           =   4800
         End
         Begin VB.Label maptext 
            Alignment       =   2  'Center
            BackColor       =   &H00808080&
            Caption         =   "map"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   162
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   180
            Left            =   0
            TabIndex        =   10
            Top             =   0
            Width           =   4800
         End
      End
      Begin VB.PictureBox msgicon 
         AutoRedraw      =   -1  'True
         BorderStyle     =   0  'None
         Height          =   300
         Left            =   360
         ScaleHeight     =   20
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   20
         TabIndex        =   8
         Top             =   1200
         Visible         =   0   'False
         Width           =   300
      End
      Begin VB.CommandButton leavegame 
         BackColor       =   &H0080C0FF&
         Caption         =   "Leave Game"
         Enabled         =   0   'False
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
         Left            =   3630
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   6750
         Visible         =   0   'False
         Width           =   3300
      End
      Begin VB.CommandButton chatwindow 
         BackColor       =   &H0080C0FF&
         Caption         =   "Chat Window"
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
         Left            =   345
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   6750
         Visible         =   0   'False
         Width           =   3300
      End
      Begin VB.Timer targetarrows 
         Enabled         =   0   'False
         Interval        =   50
         Left            =   720
         Top             =   0
      End
      Begin VB.Timer msgwink 
         Enabled         =   0   'False
         Interval        =   350
         Left            =   360
         Top             =   0
      End
      Begin MSWinsockLib.Winsock sock 
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
         RemotePort      =   41670
         LocalPort       =   41670
      End
      Begin VB.Label info 
         BackStyle       =   0  'Transparent
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
         Left            =   720
         TabIndex        =   5
         Top             =   1230
         Width           =   6135
      End
      Begin VB.Label gametype 
         BackStyle       =   0  'Transparent
         Caption         =   "Game Type"
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
         Left            =   5880
         TabIndex        =   4
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label clientname 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   3240
         TabIndex        =   3
         Top             =   840
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label hostname 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
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
         Left            =   600
         TabIndex        =   2
         Top             =   840
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frm_game"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const MAPHORZSQR = 15
Const MAPVERTSQR = 15
Const SQRSIZE = 20
Dim backvar, playorder, shiporiginalpos, myorder, placedship, otherplayer
Dim shootresult, shootresult2, compnuclear, nuclearresult, nuclearresult2
Dim lastbox(7) As String
Dim weapon As String
Dim arrows(6) As String
Dim mapinfo(4, (MAPHORZSQR + 1) * 100 + MAPVERTSQR + 1) As String
Dim shippos(3, 6, 4) As String
Dim shipshot(6) As Shot
Private Type Shot
    'for computer's deliberate shots
    lastbox As String
    shotnum As Integer
    direction As Integer
    code As String
    done As Integer
    try As Integer
End Type

Private Sub Form_Load()
    'initializing window graphics
    graphics Me.pic, "botback"
    graphics Me.pic, "bot2back"
    graphics Me.pic, "bigtopback"
    graphics Me.pic, "bigicon"
    graphics Me.pic, "player1g"
    graphics Me.pic, "singleplyr"
    graphics shipnums, "blueback"
    graphics msgicon, "message"
    graphics infoicon, "infoicon"
    graphics wait, "waitback"
    graphics wait, "waitmidback"
    graphics wait, "waiticon"
    'initializing game settings
    resetgame
    gametype = "Single Player": myorder = 0: playorder = 1: loadmap 0: shippos(1, 0, 0) = "1"
    pic.Height = 459: Me.Height = 484 * Screen.TwipsPerPixelY
    back.Picture = LoadResPicture(201, 0): backvar = False
    If playernum <= 2 Then hostname = Space(2) & frm_newgame.name1: info = Trim(hostname) & ": Please place your ships on the map."
    If playernum > 1 Then
        If playernum = 2 Then
            playorder = 3: loadmap 2: graphics Me.pic, "twoplayer": shippos(3, 0, 0) = "1"
            gametype = "Two on Same": clientname = Space(2) & frm_newgame.name2: otherplayer = Trim(clientname): playorder = 1
        Else
            chatwindow.Visible = True: leavegame.Visible = True: placedship = False
            pic.Height = 481: Me.Height = 506 * Screen.TwipsPerPixelY: connection.Enabled = True
            If mymode = "lanhost" Or mymode = "lanclient" Then
                gametype = "LAN Game": hostname = Space(2) & frm_lannet.lanhostname
                clientname = Space(2) & frm_lannet.lanclientname: graphics Me.pic, "networkg"
            Else: gametype = "Internet Game": hostname = Space(2) & frm_lannet.nethostname
                clientname = Space(2) & frm_lannet.netclientname: graphics Me.pic, "internetg"
            End If
            If mymode = "lanhost" Or mymode = "nethost" Then
                sock.LocalPort = frm_lannet.sock.LocalPort + 10: sock.RemotePort = sock.LocalPort
                sock.Listen: myorder = 0: playorder = 1: otherplayer = Trim(clientname): waitfor "Starting Game..."
            ElseIf mymode = "lanclient" Or mymode = "netclient" Then
                sock.LocalPort = frm_lannet.sock.RemotePort + 10: waitfor "Starting Game..."
                sock.Connect IIf(Len(frm_lannet.sock.RemoteHostIP) <> 0, frm_lannet.sock.RemoteHostIP, frm_lannet.sock.RemoteHost), sock.LocalPort
                myorder = 2: playorder = 3: shippos(3, 0, 0) = "1": loadmap 2: otherplayer = Trim(hostname)
            End If
        End If
        graphics Me.pic, "player2g": clientname.Visible = True
    End If
    viewmaps
    If playorder = 1 And myorder = 0 Then ships(0).Visible = True: ships(2).Visible = False: ship1(2).Visible = False: ship2(2).Visible = False: ship3(2).Visible = False: ship4(2).Visible = False: ship5(2).Visible = False: ship6(2).Visible = False
    shootresult = "": nuclearresult = "": weapon = "missile": putships -1
End Sub

Private Sub missileview_Click()
    'switches from missile to nuclear and vice versa
    If targetarrows.Enabled = True Or ordertime.Enabled = True Then Exit Sub
    If (playorder = 0 And myorder = 0) Or (playorder = 2 And playernum >= 2) Then
        If weapon = "nuclear" Then weapon = "missile": missilenum(playorder).ForeColor = QBColor(8): Exit Sub
        If weapon = "missile" And Val(missilenum(playorder)) > 0 Then weapon = "nuclear": missilenum(playorder).ForeColor = QBColor(4): Exit Sub
    End If
End Sub

Private Sub sock_Connect()
    'reports that client has connected successfully
    If mymode = "lanclient" Or mymode = "netclient" Then sock.SendData "hnb|started"
End Sub

Private Sub targetarrows_Timer()
    'puts and removes the target arrows on the map
    If arrows(4) > 0 Then
        For i = 0 To 3
            If arrows(i) <> 0 And arrows(i) <> -5 Then
                maploc = locarrows(CStr(arrows(6)), i, arrows(i))
                If Len(mapinfo(playorder, Val(maploc))) = 0 Then mapicon map(playorder), CStr(maploc), "way" & i, 1
                If Val(arrows(5)) > Val(arrows(4)) Then
                    maploc2 = locarrows(CStr(arrows(6)), i, arrows(i) + 2)
                    If mapinfo(playorder, Val(maploc2)) = "" Then mapicon map(playorder), CStr(maploc2), "blue", 1
                End If
                arrows(i) = arrows(i) - 2
            ElseIf arrows(i) = 0 Then arrows(i) = -5
            End If
        Next i
        BitBlt map(playorder).hdc, 0, 0, map(playorder).ScaleWidth, map(playorder).ScaleHeight, myMapBackBuffer(playorder), 0, 0, vbSrcCopy
        map(playorder).Picture = map(playorder).Image: arrows(4) = arrows(4) - 2: DoEvents
    ElseIf arrows(4) <= 0 Then
        targetarrows.Enabled = False
        For i = 0 To 3
            If arrows(i) <= 0 And arrows(i) <> -5 Then
                maploc2 = locarrows(CStr(arrows(6)), i, arrows(i) + 2)
                If mapinfo(playorder, Val(maploc2)) = "" Then mapicon map(playorder), CStr(maploc2), "blue", 1
            End If
        Next i
        temp = Split(shootresult, ":")
        mapinfo(playorder, temp(0)) = IIf(Val(temp(1)) > 0, "shot", temp(1))
        If Not (playernum = 2 And (playorder = 1 Or playorder = 3)) Then putships Val(temp(1))
        mapicon map(playorder), CStr(temp(0)), mapinfo(playorder, temp(0)), 1
        If (weapon = "nuclear" Or playorder = 1 Or playorder = 3) And nuclearresult <> "" Then
            temp0 = Split(nuclearresult, "|"): temp1 = Split(temp0(1), ":"): temp2 = Split(temp0(0), ":")
            boxx = Val(Left(temp(0), 2)): boxy = Val(Right(temp(0), 2)): i = 0
            For Each coord In temp2
                coords = Split(coord, "."): box = twolen(CStr(boxx + coords(0))) & twolen(CStr(boxy + coords(1)))
                mapinfo(playorder, box) = IIf(Val(temp1(i)) > 0, "shot", temp1(i))
                If mapinfo(playorder, box) = "shot" And (Not (playernum = 2 And (playorder = 1 Or playorder = 3))) Then putships Val(temp1(i))
                mapicon map(playorder), CStr(box), mapinfo(playorder, box), 1: i = i + 1
            Next
            nuclearresult = "": nuclearresult2 = IIf((playorder = 1 Or playorder = 3) And playernum >= 2, "", nuclearresult2)
        End If
        BitBlt map(playorder).hdc, 0, 0, map(playorder).ScaleWidth, map(playorder).ScaleHeight, myMapBackBuffer(playorder), 0, 0, vbSrcCopy
        map(playorder).Picture = map(playorder).Image: weapon = "missile": ordertime.Enabled = True
    End If
End Sub

Private Sub msgwink_Timer()
    'incoming message icon winks
    msgicon.Visible = IIf(msgicon.Visible = True, False, True)
End Sub

Private Sub connection_Timer()
    'closes game window, in case connection time out
    connection.Enabled = False
    frm_lannet.sock.SendData "hnb|error"
    Select Case mymode
        Case "lanhost": frm_lannet.sock.Close: frm_lannet.lancreate_Click
        Case "lanclient": frm_lannet.sock.Close: frm_lannet.lanjoin_Click
        Case "nethost": frm_lannet.sock.Close: frm_lannet.netrecreate_Click
        Case "netclient": frm_lannet.sock.Close: frm_lannet.netrejoin_Click
    End Select
    frm_lannet.puttext "Game connection timed out. Please try again."
    Unload frm_game: frm_lannet.Show
End Sub

Private Sub looserships_Timer()
    'switches between winner's and looser's maps
    If map(myorder).Visible = True Then
        playorder = myorder + 1: viewmaps
        missileview.Visible = False: Exit Sub
    Else: playorder = myorder: viewmaps
        missileview.Visible = False: Exit Sub
    End If
End Sub

Private Sub ordertime_Timer()
    'causes a break when the maps change
    ordertime.Enabled = False
    For i = 0 To 2 Step 2
        If ship1(i).ForeColor = QBColor(12) And ship2(i).ForeColor = QBColor(12) And ship3(i).ForeColor = QBColor(12) And ship4(i).ForeColor = QBColor(12) And ship5(i).ForeColor = QBColor(12) And ship6(i).ForeColor = QBColor(12) Then endofplay CInt(i): Exit Sub
    Next i
    If playernum >= 2 Then
        If playorder = 3 Then playorder = 2: viewmaps: info = Trim(clientname) & ": Shoot!": Exit Sub
        If playorder = 1 Then playorder = 0: viewmaps: info = Trim(hostname) & ": Shoot!": Exit Sub
        If playernum = 3 And (playorder = 0 Or playorder = 2) Then waitfor "Waiting " & otherplayer & " to shoot!": viewmaps: Exit Sub
    End If
    changeorder
End Sub

Private Sub map_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    'moves pointer for shooting and ships for placing
    Dim box As String, ship As Integer
    If InStr(1, mymode, "winner") > 0 Or InStr(1, mymode, "looser") > 0 Then Exit Sub
    If (playorder = myorder + 1 Or (playorder = 3 And playernum = 2)) And Val(shippos(playorder, 0, 0)) < 7 Then
        box = twolen(Int((IIf(X = 0, X, X - 1)) / SQRSIZE)) & twolen(Int((IIf(Y = 0, Y, Y - 1)) / SQRSIZE)): ship = Val(shippos(playorder, 0, 0))
        If shipplace(box, Val(shippos(playorder, ship, 0)), 2) = False Then Exit Sub
        If lastbox(ship) <> "" Then mapmouse map(playorder), X, Y, 0 - ship
        mapmouse map(playorder), X, Y, ship
    ElseIf playorder = myorder Or (playorder = 2 And playernum = 2) Then
        If targetarrows.Enabled = True Or ordertime.Enabled = True Then Exit Sub
        mapmouse map(Index), X, Y, 0
    End If
End Sub

Private Sub map_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim box As String
    Dim boxx, boxy, boxnew, info
    If X < 0 Or Y < 0 Or X > SQRSIZE * (MAPHORZSQR + 1) Or Y > SQRSIZE * (MAPVERTSQR + 1) Or InStr(1, mymode, "winner") > 0 Or InStr(1, mymode, "looser") > 0 Then Exit Sub
    box = twolen(Int((IIf(X = 0, X, X - 1)) / SQRSIZE)) & twolen(Int((IIf(Y = 0, Y, Y - 1)) / SQRSIZE))
    boxx = Val(Left(box, 2)): boxy = Val(Right(box, 2))
    If Button = 1 Then
        'ship shooting
        If playorder = myorder Or (playorder = 2 And playernum = 2) Then
            If mapinfo(playorder, Val(box)) <> "" Or targetarrows.Enabled = True Or ordertime.Enabled = True Then Exit Sub
            If playorder = 0 And myorder = 0 And playernum <= 2 Then
                info = 3
            ElseIf playorder = 2 And playernum = 2 Then info = 1
            ElseIf playernum = 3 And (playorder = 0 Or playorder = 2) Then info = IIf(playorder = 0, 3, IIf(playorder = 2, 1, 0))
            End If
            shootresult = box & ":" & IIf(mapinfo(info, Val(box)) = "", "miss", mapinfo(info, Val(box))): shootresult2 = shootresult
            If weapon = "nuclear" Then
                missilenum(playorder) = Val(missilenum(playorder)) - 1: missilenum(playorder).ForeColor = QBColor(8)
                temp0 = Array("-1.0", "-1.-1", "0.-1", "1.-1", "1.0", "1.1", "0.1", "-1.1"): temp1 = Array("", "")
                For Each coord In temp0
                    coords = Split(coord, ".")
                    If boxx + coords(0) < 0 Or boxx + coords(0) > MAPHORZSQR Or boxy + coords(1) < 0 Or boxy + coords(1) > MAPVERTSQR Then GoTo 10
                    boxnew = twolen(CStr(boxx + coords(0))) & twolen(CStr(boxy + coords(1)))
                    If mapinfo(playorder, Val(boxnew)) = "" Then temp1(0) = temp1(0) & coords(0) & "." & coords(1) & ":": temp1(1) = temp1(1) & IIf(mapinfo(info, Val(boxnew)) = "", "miss", mapinfo(info, Val(boxnew))) & ":"
10              Next
                If temp1(0) <> "" Then nuclearresult = Left(temp1(0), Len(temp1(0)) - 1) & "|" & Left(temp1(1), Len(temp1(1)) - 1): nuclearresult2 = nuclearresult
            End If
            If playernum = 3 Then sock.SendData "hnb|shot|" & box & "|" & weapon & IIf(nuclearresult <> "", "|" & nuclearresult, "")
            lastbox(0) = "0": putarrows box: targetarrows.Enabled = True
        'ship placing
        ElseIf playorder = myorder + 1 Or (playorder = 3 And playernum = 2) And Val(shippos(playorder, 0, 0)) < 7 Then
            If targetarrows.Enabled = True Or ordertime.Enabled = True Then Exit Sub
            If shipplace(box, Val(shippos(playorder, Val(shippos(playorder, 0, 0)), 0))) = False Then Beep: Exit Sub
            shipplace box, Val(shippos(playorder, Val(shippos(playorder, 0, 0)), 0)), 1
            mapmouse map(playorder), X, Y, Val(shippos(playorder, 0, 0)), 10
            shippos(playorder, 0, 0) = CStr(Val(shippos(playorder, 0, 0)) + 1)
            putships 0 - Val(shippos(playorder, 0, 0)): lastbox(0) = "0"
            'end of placing
            If playorder = 1 And myorder = 0 And Val(shippos(playorder, 0, 0)) = 7 Then
                If playernum = 1 Then
                    playorder = 3: comp_placeship
                ElseIf playernum = 2 Then waitfor "It is" & clientname & "'s turn to place ships!"
                End If
            ElseIf playorder = 3 And playernum = 2 And Val(shippos(playorder, 0, 0)) = 7 Then waitfor "It is" & hostname & "'s turn to shoot!"
            End If
            If playernum = 3 And Val(shippos(playorder, 0, 0)) = 7 Then
                sendships
                If placedship <> True Then waitfor "Waiting " & otherplayer & " to place ships!"
                Exit Sub
            End If
        End If
    ElseIf Button = 2 Then
        'ship rotating
        If (playorder = myorder + 1 Or (playorder = 3 And playernum = 2)) And Val(shippos(playorder, 0, 0)) < 7 Then
            i = Val(shippos(playorder, Val(shippos(playorder, 0, 0)), 0)): j = Val(i): k = 0
            Do: j = base(j + 1, 4)
                If j = Val(i) Then Exit Sub
            Loop Until shipplace(box, Val(j), 2) = True
            mapmouse map(playorder), X, Y, 0 - Val(shippos(playorder, 0, 0)), Val(i)
            lastbox(shippos(playorder, 0, 0)) = "0"
            shippos(playorder, Val(shippos(playorder, 0, 0)), 0) = CStr(j)
            mapmouse map(playorder), X, Y, Val(shippos(playorder, 0, 0))
        End If
    End If
End Sub

Private Sub wait_Click()
    'player changer
    If playernum = 2 Then
        If playorder = 1 Then
            playorder = 3: viewmaps
            For i = 2 To 6: putships CInt(0 - i): Next i
            info = clientname & ": Please place your ships on the map.": waitfor
        ElseIf playorder = 3 Then playorder = 0: viewmaps: info = Trim(hostname) & ": Shoot!": waitfor
        ElseIf playorder = 0 Then
            playorder = 3: viewmaps: shootresult = shootresult2: nuclearresult = nuclearresult2
            putarrows CStr(Left(shootresult, 4)): waitfor
            If nuclearresult <> "" Then info = Trim(hostname) & " uses nuclear!": Beep
            targetarrows.Enabled = True
        ElseIf playorder = 2 Then
            playorder = 1: viewmaps: shootresult = shootresult2: nuclearresult = nuclearresult2
            putarrows CStr(Left(shootresult, 4)): waitfor
            If nuclearresult <> "" Then info = clientname & " uses nuclear!": Beep
            targetarrows.Enabled = True
        End If
    End If
End Sub

Private Function changeorder()
    'changes the play order and screen
    If playorder = 0 Then
        If playernum = 1 Then playorder = 2: viewmaps: comp_shoot: Exit Function
        If playernum = 2 Then waitfor "It is" & clientname & "'s turn to shoot!": Exit Function
        If playernum = 3 Then playorder = 2: Exit Function
    ElseIf playorder = 1 Then
        If playernum = 1 Then playorder = 0: info = Trim(hostname) & ": You shoot!": viewmaps: Exit Function
    ElseIf playorder = 2 Then
        If playernum = 1 Then playorder = 1: info = "Computer Shoots!": viewmaps: targetarrows.Enabled = True: Exit Function
        If playernum = 2 Then waitfor "It is" & hostname & "'s turn to shoot!": Exit Function
    ElseIf playorder = 3 Then
        If playernum = 1 Then playorder = 0: info = Trim(hostname) & ": You shoot!": viewmaps: Exit Function
    End If
End Function

Private Function mapmouse(obj As PictureBox, px As Single, py As Single, ship As Integer, Optional rotate As Integer)
    'draws items on map
    Dim box As String
    Dim boxx, boxy, boxnew, scode, coord
    box = twolen(Int((IIf(px = 0, px, px - 1)) / SQRSIZE)) & twolen(Int((IIf(py = 0, py, py - 1)) / SQRSIZE))
    If lastbox(0) <> box And ship = 0 Then
        If lastbox(0) <> "" And lastbox(0) <> "0" Then mapicon obj, lastbox(0), IIf(mapinfo(playorder, lastbox(0)) = "", "blue", mapinfo(playorder, lastbox(0))), 1
        lastbox(0) = box: mapicon obj, box, weapon, 1
        BitBlt obj.hdc, 0, 0, obj.ScaleWidth, obj.ScaleHeight, myMapBackBuffer(obj.Index), 0, 0, vbSrcCopy: obj.Picture = obj.Image
    ElseIf (lastbox(Abs(ship)) <> box And ship > 0) Or (lastbox(Abs(ship)) = box And rotate = 10) Then
        boxx = Val(Left(box, 2)): boxy = Val(Right(box, 2))
        scode = shippos(playorder, ship, Val(shippos(playorder, ship, 0)))
        args = Split(scode, ":")
        For Each coord In args
            temp = Split(coord, ".")
            If boxx + Val(temp(0)) < 0 Or boxx + Val(temp(0)) > MAPHORZSQR Or boxy + Val(temp(1)) < 0 Or boxy + Val(temp(1)) > MAPVERTSQR Then GoTo 10
            boxnew = twolen(CStr(boxx + Val(temp(0)))) & twolen(CStr(boxy + Val(temp(1))))
            mapicon obj, CStr(boxnew), IIf(rotate = 10, "ship", "ship2"), 1
10      Next
        BitBlt obj.hdc, 0, 0, obj.ScaleWidth, obj.ScaleHeight, myMapBackBuffer(obj.Index), 0, 0, vbSrcCopy
        obj.Picture = obj.Image: lastbox(ship) = box
    ElseIf (lastbox(Abs(ship)) <> box And ship < 0) Or (rotate > 0 And ship < 0) Then
        box = lastbox(Abs(ship)): boxx = Val(Left(box, 2)): boxy = Val(Right(box, 2))
        scode = shippos(playorder, Abs(ship), Val(shippos(playorder, Abs(ship), 0)))
        args = Split(scode, ":")
        For Each coord In args
            temp = Split(coord, ".")
            If boxx + Val(temp(0)) < 0 Or boxx + Val(temp(0)) > MAPHORZSQR Or boxy + Val(temp(1)) < 0 Or boxy + Val(temp(1)) > MAPVERTSQR Then GoTo 20
            boxnew = twolen(CStr(boxx + Val(temp(0)))) & twolen(CStr(boxy + Val(temp(1))))
            mapicon obj, CStr(boxnew), IIf(mapinfo(playorder, Val(boxnew)) = "", "blue", IIf(Val(mapinfo(playorder, Val(boxnew))) > 0, "ship", mapinfo(playorder, Val(boxnew)))), 1
20      Next
        BitBlt obj.hdc, 0, 0, obj.ScaleWidth, obj.ScaleHeight, myMapBackBuffer(obj.Index), 0, 0, vbSrcCopy: obj.Picture = obj.Image
    End If
End Function

Private Function shipplace(box As String, rotate As Integer, Optional func As Integer)
    'reports if it is possible to place the ship on that coords
    Dim boxx, boxy, scode, args
    boxx = Val(Left(box, 2)): boxy = Val(Right(box, 2))
    scode = shippos(playorder, Val(shippos(playorder, 0, 0)), rotate)
    args = Split(scode, ":"): temp0 = Array("-1.0", "0.-1", "1.0", "0.1")
    For Each coord In args
        temp = Split(coord, ".")
        If Val(boxx) + Val(temp(0)) < 0 Or Val(boxx) + Val(temp(0)) > MAPHORZSQR Or Val(boxy) + Val(temp(1)) < 0 Or Val(boxy) + Val(temp(1)) > MAPVERTSQR Then shipplace = False: Exit Function
        If func <> 2 And func <> 1 Then
            If mapinfo(playorder, Val(boxx) + Val(temp(0)) & twolen(CStr(Val(boxy) + Val(temp(1))))) <> "" Then shipplace = False: Exit Function
            For i = 0 To 3
                temp2 = Split(temp0(i), ".")
                If (Val(boxx) + Val(temp(0)) + Val(temp2(0)) < 0 Or Val(boxx) + Val(temp(0)) + Val(temp2(0)) > MAPHORZSQR) Or (Val(boxy) + Val(temp(1)) + Val(temp2(1)) < 0 Or Val(boxy) + Val(temp(1)) + Val(temp2(1)) > MAPVERTSQR) Then GoTo 10
                If Val(mapinfo(playorder, twolen(CStr(Val(boxx) + Val(temp(0)) + Val(temp2(0)))) & twolen(CStr(Val(boxy) + Val(temp(1)) + Val(temp2(1)))))) > 0 Then shipplace = False: Exit Function
10          Next i
        End If
        If func = 1 Then mapinfo(playorder, twolen(CStr(Val(boxx) + Val(temp(0)))) & twolen(CStr(Val(boxy) + Val(temp(1))))) = shippos(playorder, 0, 0)
    Next
    shipplace = True
End Function

Private Function mapicon(obj As PictureBox, box As String, icon As String, Optional complete As Integer)
    'manages map pictures
    Dim pos
    Select Case icon
        Case "blue": pos = Array(350, 0, 20, 20)
        Case "miss": pos = Array(270, 0, 20, 20)
        Case "shot": pos = Array(290, 0, 20, 20)
        Case "missile": pos = Array(310, 0, 20, 20)
        Case "nuclear": pos = Array(330, 0, 20, 20)
        Case "way0": pos = Array(330, 20, 20, 20)
        Case "way1": pos = Array(270, 20, 20, 20)
        Case "way2": pos = Array(290, 20, 20, 20)
        Case "way3": pos = Array(310, 20, 20, 20)
        Case "ship": pos = Array(350, 20, 20, 20)
        Case "ship2": pos = Array(350, 40, 20, 20)
    End Select
    If IsEmpty(pos) = False And box <> "" Then
        BitBlt myMapBackBuffer(obj.Index), Val(Left(box, 2)) * SQRSIZE, Val(Right(box, 2)) * SQRSIZE, pos(2), pos(3), mySprite, pos(0), pos(1), vbSrcCopy
        If Val(complete) <> 1 Then BitBlt obj.hdc, 0, 0, obj.ScaleWidth, obj.ScaleHeight, myMapBackBuffer(obj.Index), 0, 0, vbSrcCopy: obj.Picture = obj.Image
    End If
End Function

Private Function comp_placeship()
    'computer places its ships
    Randomize Timer
    Dim boxx, boxy, box, rotate
    For i = 1 To 6
        Do: boxx = Int(Rnd * (MAPHORZSQR + 1)): boxy = Int(Rnd * (MAPVERTSQR + 1)): rotate = Int(Rnd * 4) + 1
            box = twolen(CStr(boxx)) & twolen(CStr(boxy))
        Loop Until shipplace(CStr(box), CInt(rotate)) = True
        shipplace CStr(box), CInt(rotate), 1: shippos(playorder, 0, 0) = shippos(playorder, 0, 0) + 1: lastbox(i) = ""
    Next i
    compnuclear = 1: changeorder
End Function

Private Function comp_shoot()
    'computer shoots
    Randomize Timer
    Dim boxx, boxy, box, boxnew, boxold, result, lastnbox
    Dim basicdir, coord, coords(1), done, direction
    basicdir = Array("-1.0", "0.-1", "1.0", "0.1")
    'decide for nuclear using
    If compnuclear = 1 Then
        For i = 0 To 2 Step 2: j = 0
            If ship1(i).ForeColor = QBColor(12) Then j = j + 1
            If ship2(i).ForeColor = QBColor(12) Then j = j + 1
            If ship3(i).ForeColor = QBColor(12) Then j = j + 1
            If ship4(i).ForeColor = QBColor(12) Then j = j + 1
            If ship5(i).ForeColor = QBColor(12) Then j = j + 1
            If ship6(i).ForeColor = QBColor(12) Then j = j + 1
            If j = 5 Then weapon = "nuclear": compnuclear = 0: GoTo 5
        Next i
        If Val(gettok(lastbox(7), 1, ":")) = 6 And shipshot(Val(gettok(lastbox(7), 1, ":"))).shotnum < 4 Then weapon = "nuclear": compnuclear = 0: GoTo 5
    End If
    'decide for random target
5   If Not Val(gettok(lastbox(7), 1, ":")) > 0 Then
        Do
            'this line is for control mechanism
            'you have to place a checkbox named 'Check1' to execute
            'If Check1.Value = vbChecked Then box = CStr(InputBox("box")): GoTo 7
            Do: boxx = Int(Rnd * (MAPHORZSQR + 1)): boxy = (Int(Rnd * (Int(MAPVERTSQR / 2) + 1)) * 2) + (boxx Mod 2): Loop Until boxy < MAPVERTSQR + 1
            box = twolen(CStr(boxx)) & twolen(CStr(boxy)): result = 0
            For i = 0 To 3
                coord = Split(basicdir(i), "."): coords(0) = boxx + coord(0): coords(1) = boxy + coord(1)
                If coords(0) >= 0 And coords(0) <= MAPHORZSQR And coords(1) >= 0 And coords(1) <= MAPVERTSQR Then
                    boxnew = twolen(CStr(coords(0))) & twolen(CStr(coords(1)))
                    If Val(mapinfo(2, CStr(boxnew))) > 0 Then result = result + 1
                End If
            Next i
        Loop Until mapinfo(2, Val(box)) = "" And result = 0
    Else
    'decide for target deliberately
7       ship = Val(gettok(lastbox(7), 1, ":")): boxold = shipshot(ship).lastbox: i = 0
        If shipshot(ship).shotnum = 1 Then
            Do: result = 0: boxx = Val(Left(boxold, 2)): boxy = Val(Right(boxold, 2))
                coord = Split(basicdir(i), "."): coords(0) = boxx + coord(0): coords(1) = boxy + coord(1)
                If coords(0) >= 0 And coords(0) <= MAPHORZSQR And coords(1) >= 0 And coords(1) <= MAPVERTSQR Then
                    box = twolen(CStr(coords(0))) & twolen(CStr(coords(1)))
                    If Val(mapinfo(2, CStr(box))) > 0 Then result = result + 1
                End If: i = i + 1
            Loop Until mapinfo(2, Val(box)) = "" And result = 0 And box <> ""
        ElseIf shipshot(ship).shotnum > 1 Then
10          direction = verthorz(sortbox(shipshot(ship).code))
            If (ship = 4 Or ship = 6) And shipshot(ship).shotnum = 4 And numtok(sortbox(shipshot(ship).code), ":") = numtok(shipshot(ship).code, ":") Then
                shipshot(ship).lastbox = gettok(sortbox(shipshot(ship).code), IIf(shipshot(ship).try Mod 2 = 0, 2, 3), ":"): shipshot(ship).done = 1
                shipshot(ship).direction = IIf(verthorz(sortbox(shipshot(ship).code)) Mod 2 = 0, IIf(Val(shipshot(ship).try) Mod 2 = 0, 1, 3), IIf(Val(shipshot(ship).try) Mod 2 = 0, 2, 0))
                shipshot(ship).try = shipshot(ship).try + 1
            ElseIf ship = 4 And shipshot(ship).shotnum = 2 And numtok(sortbox(shipshot(ship).code), ":") = numtok(shipshot(ship).code, ":") And shipshot(ship).done = 1 Then
                shipshot(ship).lastbox = gettok(sortbox(shipshot(ship).code), IIf(shipshot(ship).try Mod 2 = 0, 2, 1), ":"): shipshot(ship).try = shipshot(ship).try + 1
            ElseIf ship = 4 And shipshot(ship).shotnum = 3 And numtok(sortbox(shipshot(ship).code), ":") <> numtok(shipshot(ship).code, ":") Then
                shipshot(ship).direction = changedir(shipshot(ship).direction)
            ElseIf ship = 6 And shipshot(ship).shotnum = 3 And numtok(sortbox(shipshot(ship).code), ":") < numtok(shipshot(ship).code, ":") Then
                shipshot(ship).lastbox = gettok(sortbox(shipshot(ship).code), IIf(shipshot(ship).try <= 1, 1, 2), ":")
                shipshot(ship).direction = verthorz(sortbox(shipshot(ship).code)): shipshot(ship).try = shipshot(ship).try + 1
            ElseIf ship = 6 And shipshot(ship).shotnum = 5 And shipshot(ship).try = 0 And numtok(sortbox(shipshot(ship).code, 2), ":") = 1 Then
                shipshot(ship).direction = IIf(shipshot(ship).lastbox = sortbox(shipshot(ship).code, 2), changedir(shipshot(ship).direction, 1), shipshot(ship).direction)
                shipshot(ship).lastbox = sortbox(shipshot(ship).code, 2): shipshot(ship).try = shipshot(ship).try + 1: shipshot(ship).done = 1
            ElseIf ship = 6 And shipshot(ship).shotnum >= 5 And (numtok(sortbox(shipshot(ship).code), ":") = 3 Or numtok(sortbox(shipshot(ship).code, 2), ":") = 3) Then
                If numtok(sortbox(shipshot(ship).code), ":") = numtok(sortbox(shipshot(ship).code, 2), ":") Then
                    shipshot(ship).lastbox = gettok(sortbox(shipshot(ship).code, IIf(Int(shipshot(ship).try / 2) Mod 2 = 0, 1, 2)), IIf(shipshot(ship).try Mod 2 = 0, 1, numtok(sortbox(shipshot(ship).code, IIf(shipshot(ship).try < 2, 1, 2)), ":")), ":")
                Else: shipshot(ship).lastbox = gettok(sortbox(shipshot(ship).code, 2), IIf(shipshot(ship).try Mod 2 = 0, 1, numtok(sortbox(shipshot(ship).code, 2), ":")), ":")
                End If
                shipshot(ship).direction = verthorz(gettok(sortbox(shipshot(ship).code), 1, ":") & ":" & gettok(sortbox(shipshot(ship).code), 2, ":"))
                shipshot(ship).direction = IIf(shipshot(ship).try Mod 2 = 0, changedir(shipshot(ship).direction), shipshot(ship).direction)
                shipshot(ship).try = shipshot(ship).try + 1: shipshot(ship).done = 1
            End If
            done = 0: temp1 = Split(sortbox(shipshot(ship).code), ":")
            temp2 = numtok(sortbox(shipshot(ship).code), ":"): direction = shipshot(ship).direction
            For i = 1 To 2
                coord = Split(basicdir(shipshot(ship).direction), ".")
                If i = 1 Then boxnew = IIf(direction < changedir(CInt(direction)), temp1(0), temp1(temp2 - 1))
                If i = 2 Then boxnew = IIf(direction < changedir(CInt(direction)), temp1(temp2 - 1), temp1(0))
                If shipshot(ship).done = 1 Then boxnew = shipshot(ship).lastbox: i = 2
                coords(0) = Val(Left(boxnew, 2)) + coord(0): coords(1) = Val(Right(boxnew, 2)) + coord(1)
                box = Val(twolen(CStr(coords(0))) & twolen(CStr(coords(1))))
                If coords(0) < 0 Or coords(0) > MAPHORZSQR Or coords(1) < 0 Or coords(1) > MAPVERTSQR Then shipshot(ship).direction = changedir(shipshot(ship).direction): done = done + 1
                If coords(0) >= 0 And coords(0) <= MAPHORZSQR And coords(1) >= 0 And coords(1) <= MAPVERTSQR Then
                    If mapinfo(2, box) = "miss" Then shipshot(ship).direction = changedir(shipshot(ship).direction): done = done + 1
                End If
                coords(0) = coords(0) + coord(0): coords(1) = coords(1) + coord(1)
                box = Val(twolen(CStr(coords(0))) & twolen(CStr(coords(1))))
                If coords(0) >= 0 And coords(0) <= MAPHORZSQR And coords(1) >= 0 And coords(1) <= MAPVERTSQR Then
                    If Val(mapinfo(2, box)) > 0 Then shipshot(ship).direction = changedir(shipshot(ship).direction): done = done + 1
                End If
                If done = 2 Then
                    shipshot(ship).done = shipshot(ship).done + 1
                    shipshot(ship).direction = changedir(shipshot(ship).direction, 1): GoTo 10
                End If
                If done = 0 Then GoTo 20
            Next i
20          coord = Split(basicdir(shipshot(ship).direction), "."): box = shipshot(ship).lastbox
            Do
                boxx = Val(Left(box, 2)) + coord(0): boxy = Val(Right(box, 2)) + coord(1)
                If boxx < 0 Or boxx > MAPHORZSQR Or boxy < 0 Or boxy > MAPVERTSQR Then shipshot(ship).direction = changedir(shipshot(ship).direction): GoTo 10
                box = twolen(CStr(boxx)) & twolen(CStr(boxy))
                If mapinfo(2, Val(box)) = "miss" Then shipshot(ship).direction = changedir(shipshot(ship).direction): GoTo 20
                If Val(mapinfo(2, Val(box))) > 0 Then shipshot(ship).lastbox = box
            Loop Until mapinfo(2, Val(box)) = ""
        End If
    End If
    mapinfo(2, Val(box)) = IIf(mapinfo(1, Val(box)) = "", "miss", mapinfo(1, Val(box)))
    If weapon = "nuclear" Then
        boxx = Val(Left(box, 2)): boxy = Val(Right(box, 2))
        temp0 = Array("-1.0", "-1.-1", "0.-1", "1.-1", "1.0", "1.1", "0.1", "-1.1"): temp1 = Array("", "")
        For Each coord In temp0
            coord2 = Split(coord, ".")
            If boxx + coord2(0) < 0 Or boxx + coord2(0) > MAPHORZSQR Or boxy + coord2(1) < 0 Or boxy + coord2(1) > MAPVERTSQR Then GoTo 30
            boxnew = twolen(CStr(boxx + coord2(0))) & twolen(CStr(boxy + coord2(1)))
            If mapinfo(2, Val(boxnew)) <> "" Then GoTo 30
            temp1(0) = temp1(0) & coord2(0) & "." & coord2(1) & ":": temp1(1) = temp1(1) & IIf(mapinfo(1, Val(boxnew)) = "", "miss", mapinfo(1, Val(boxnew))) & ":"
            mapinfo(2, Val(boxnew)) = IIf(mapinfo(1, Val(boxnew)) = "", "miss", mapinfo(1, Val(boxnew)))
            If Val(mapinfo(2, Val(boxnew))) > 0 Then
                lastbox(7) = addtok(lastbox(7), mapinfo(2, Val(boxnew)), ":"): shipshot(Val(gettok(lastbox(7), numtok(lastbox(7), ":"), ":"))).shotnum = shipshot(Val(gettok(lastbox(7), numtok(lastbox(7), ":"), ":"))).shotnum + 1
                If shipshot(Val(gettok(lastbox(7), numtok(lastbox(7), ":"), ":"))).shotnum = 2 Then lastnbox = shipshot(Val(gettok(lastbox(7), numtok(lastbox(7), ":"), ":"))).code
                shipshot(Val(gettok(lastbox(7), numtok(lastbox(7), ":"), ":"))).code = addtok(shipshot(Val(gettok(lastbox(7), numtok(lastbox(7), ":"), ":"))).code, CStr(boxnew), ":"): shipshot(Val(gettok(lastbox(7), numtok(lastbox(7), ":"), ":"))).lastbox = CStr(boxnew)
                If lastnbox <> "" Then shipshot(Val(gettok(lastbox(7), numtok(lastbox(7), ":"), ":"))).direction = IIf(verthorz(lastnbox & ":" & CStr(boxnew)) = "", shipshot(Val(gettok(lastbox(7), numtok(lastbox(7), ":"), ":"))).direction, verthorz(lastnbox & ":" & CStr(boxnew)))
                lastnbox = CStr(boxnew)
            End If
30      Next
        nuclearresult = Left(temp1(0), Len(temp1(0)) - 1) & "|" & Left(temp1(1), Len(temp1(1)) - 1)
        info = "Computer uses nuclear!": Beep
    End If
    If Val(mapinfo(2, Val(box))) > 0 Then
        lastbox(7) = addtok(lastbox(7), mapinfo(2, Val(box)), ":"): ship = Val(gettok(lastbox(7), 1, ":"))
        shipshot(ship).shotnum = shipshot(ship).shotnum + 1
        If shipshot(ship).lastbox <> "" Then shipshot(ship).direction = IIf(verthorz(shipshot(ship).lastbox & ":" & box) = "", shipshot(ship).direction, verthorz(shipshot(ship).lastbox & ":" & box))
        shipshot(ship).lastbox = box: shipshot(ship).code = addtok(shipshot(ship).code, CStr(box), ":"): shipshot(ship).try = 0
    End If
    shootresult = box & ":" & mapinfo(2, Val(box)): putarrows CStr(box)
    changeorder
End Function

Private Function rotateship()
    'rotates ship code
    Dim scode, turn
    shippos(1, 0, 0) = "1": shippos(3, 0, 0) = "1"
    For i = 1 To 6
        shippos(1, i, 1) = shiporiginalpos(i): shippos(3, i, 1) = shiporiginalpos(i)
        shippos(1, i, 0) = "1": shippos(3, i, 0) = "1"
        For j = 2 To 4
            scode = Split(shippos(1, i, j - 1), ":"): turn = ""
            For Each coord In scode: turn = turn & changecoord(CStr(coord), 0) & ":": Next
            turn = Left(turn, Len(turn) - 1): shippos(1, i, j) = turn: shippos(3, i, j) = turn
        Next j
    Next i
    shippos(1, 6, 3) = "-1.0:0.0:1.0:2.0:-2.-1:-1.-1:0.-1:1.-1": shippos(3, 6, 3) = shippos(1, 6, 3)
    shippos(1, 6, 4) = "0.2:0.1:0.0:0.-1:1.1:1.0:1.-1:1.-2": shippos(3, 6, 4) = shippos(1, 6, 4)
End Function

Private Function sendships()
    'sends locations of the ships to the other player
    Dim scode As String
    Dim ncode As String
    For i = 0 To MAPHORZSQR
        For j = 0 To MAPVERTSQR
            box = twolen(CStr(i)) & twolen(CStr(j))
            If Val(mapinfo(playorder, Val(box))) > 0 Then scode = scode & mapinfo(playorder, Val(box)) & "." & CStr(box) & ":"
        Next j
    Next i
    If placedship = True Then
        If mymode = "lanhost" Or mymode = "nethost" Then ncode = "shooting": playorder = 0: viewmaps: waitfor
        If mymode = "lanclient" Or mymode = "netclient" Then ncode = "youshoot": waitfor "Waiting " & otherplayer & " to shoot!"
    End If
    sock.SendData "hnb|sendship|" & Left(scode, Len(scode) - 1) & IIf(Len(ncode) > 0, "|" & ncode, "")
End Function

Private Function endofplay(winnerplayer As Integer)
    'reports end of game and shows maps to players
    Dim box, winner, winnername, sourcemap
    If myorder = winnerplayer Then
        winnername = Trim(playername): mymode = mymode & "winner"
    Else: winnername = IIf(playernum = 1, "Computer", otherplayer): mymode = mymode & "looser"
    End If
    MsgBox "Game Over!" & vbCrLf & winnername & " has won the game!", vbInformation + vbOKOnly, "Game Over"
    info = winnername & " has won the game!"
    If InStr(1, mymode, "looser") > 0 Then
        sourcemap = IIf(myorder = 0, 3, 1)
        For i = 0 To MAPHORZSQR: For j = 0 To MAPVERTSQR
            box = twolen(CStr(i)) & twolen(CStr(j))
            If Val(mapinfo(sourcemap, Val(box))) > 0 And mapinfo(myorder, Val(box)) = "" Then mapicon map(myorder), CStr(box), "ship", 1
        Next j: Next i
        BitBlt map(myorder).hdc, 0, 0, map(myorder).ScaleWidth, map(myorder).ScaleHeight, myMapBackBuffer(myorder), 0, 0, vbSrcCopy
        map(myorder).Picture = map(myorder).Image
    End If
    looserships.Enabled = True
End Function

Private Function putarrows(box As String)
    'initializes first positions of target arrows
    Dim boxx, boxy, args, tmps
    boxx = Val(Left(box, 2)): boxy = Val(Right(box, 2))
    args = Array(boxy, MAPHORZSQR + 1 - boxx, MAPVERTSQR + 1 - boxy, boxx)
    tmps = Array(boxy, MAPHORZSQR + 1 - boxx, MAPVERTSQR + 1 - boxy, boxx)
    For i = 0 To 2: For j = 0 To 2
        If args(j) > args(j + 1) Then temp = args(j): args(j) = args(j + 1): args(j + 1) = temp
    Next j: Next i
    i = 1: While args(0) < 4: args(0) = args(i): i = i + 1: Wend
    args(0) = IIf(args(0) > 10, 10, args(0))
    arrows(4) = args(0): arrows(5) = args(0): arrows(6) = box
    For i = 0 To 3: arrows(i) = IIf(tmps(i) < 4, 0, args(0)): Next i
End Function

Private Function locarrows(box As String, way, dist)
    'locates target arrows for any moment
    Dim boxx, boxy, c
    boxx = Val(Left(box, 2)): boxy = Val(Right(box, 2))
    If way = 1 Or way = 2 Then c = 1 Else c = -1
    If way = 1 Or way = 3 Then locarrows = twolen(CStr(boxx + c * dist)) & twolen(CStr(boxy))
    If way = 0 Or way = 2 Then locarrows = twolen(CStr(boxx)) & twolen(CStr(boxy + c * dist))
End Function

Private Function loadmap(mapn As Integer)
    'loads map, ship, missile pictures for any player
    Dim box As String
    If mapn = 0 Then
        Load map(1)
        For i = 0 To MAPHORZSQR + 1: For j = 0 To MAPVERTSQR + 1
                box = twolen(CStr(i)) & twolen(CStr(j))
                mapinfo(0, Val(box)) = "": mapinfo(1, Val(box)) = "": mapinfo(2, Val(box)) = "": mapinfo(3, Val(box)) = ""
                If i <= MAPHORZSQR And j <= MAPVERTSQR Then mapicon map(0), box, "blue", 1: mapicon map(1), box, "blue", 1
        Next j: Next i
        BitBlt map(0).hdc, 0, 0, map(0).ScaleWidth, map(0).ScaleHeight, myMapBackBuffer(0), 0, 0, vbSrcCopy
        BitBlt map(1).hdc, 0, 0, map(1).ScaleWidth, map(1).ScaleHeight, myMapBackBuffer(1), 0, 0, vbSrcCopy
        map(0).Picture = map(0).Image: map(1).Picture = map(1).Image
        Load ships(2): Load ship1(2): Load ship2(2): Load ship3(2): Load ship4(2): Load ship5(2): Load ship6(2)
        graphics ships(2), "blueback": ships(2).CurrentX = 0: ships(2).Line (60, 0)-(60, 270), QBColor(15)
        graphics ships(0), "blueback"
        graphics missileview, "pinkback"
        graphics missileview, "missile1"
        graphics missileview, "missile2"
        For i = 0 To 5: graphics ships(2), "ships", CInt(i): Next i
        putships 11: putships 13
    ElseIf mapn = 2 Then
        Load map(2): Load map(3)
        For i = 0 To MAPHORZSQR: For j = 0 To MAPVERTSQR
                box = twolen(CStr(i)) & twolen(CStr(j))
                mapicon map(2), box, "blue", 1: mapicon map(3), box, "blue", 1
        Next j: Next i
        BitBlt map(2).hdc, 0, 0, map(2).ScaleWidth, map(2).ScaleHeight, myMapBackBuffer(2), 0, 0, vbSrcCopy
        BitBlt map(3).hdc, 0, 0, map(3).ScaleWidth, map(3).ScaleHeight, myMapBackBuffer(3), 0, 0, vbSrcCopy
        map(2).Picture = map(2).Image: map(3).Picture = map(3).Image
        Load missilenum(2): Load missilenum0(2)
    End If
End Function

Private Function viewmaps()
    'shows/hides maps for players
    map(0).Visible = False: map(1).Visible = False: ships(0).Visible = False: ships(2).Visible = False: missilenum(0).Visible = False: missileview.Visible = False: missilenum0(0).Visible = False
    ship1(0).Visible = False: ship2(0).Visible = False: ship3(0).Visible = False: ship4(0).Visible = False: ship5(0).Visible = False: ship6(0).Visible = False
    ship1(2).Visible = False: ship2(2).Visible = False: ship3(2).Visible = False: ship4(2).Visible = False: ship5(2).Visible = False: ship6(2).Visible = False
    If playernum = 2 Or (playorder > 1 And myorder = 2) Then map(2).Visible = False: map(3).Visible = False: missilenum(2).Visible = False: missilenum0(2).Visible = False
    If playorder = 0 Then
        map(0).Visible = True: ships(0).Visible = True: missileview.Visible = True: missilenum(0).Visible = True: missilenum0(0).Visible = True
        ship1(0).Visible = True: ship2(0).Visible = True: ship3(0).Visible = True: ship4(0).Visible = True: ship5(0).Visible = True: ship6(0).Visible = True
    ElseIf playorder = 1 Then
        map(1).Visible = True: ships(2).Visible = True
        ship1(2).Visible = True: ship2(2).Visible = True: ship3(2).Visible = True: ship4(2).Visible = True: ship5(2).Visible = True: ship6(2).Visible = True
    ElseIf playorder = 2 Then
        If playernum >= 2 Then map(2).Visible = True: ships(2).Visible = True: missileview.Visible = True: missilenum(2).Visible = True: missilenum0(2).Visible = True
        ship1(2).Visible = True: ship2(2).Visible = True: ship3(2).Visible = True: ship4(2).Visible = True: ship5(2).Visible = True: ship6(2).Visible = True
    ElseIf playorder = 3 Then
        map(3).Visible = True: ships(0).Visible = True
        ship1(0).Visible = True: ship2(0).Visible = True: ship3(0).Visible = True: ship4(0).Visible = True: ship5(0).Visible = True: ship6(0).Visible = True
    End If
End Function

Private Sub sock_DataArrival(ByVal bytesTotal As Long)
    'controls all the connections between host and client
    Dim data As String
    Dim command
    sock.GetData data
    If Left(data, 3) <> "hnb" Then Exit Sub
    command = Split(data, "|")
10  If mymode = "lanhost" Or mymode = "nethost" Then
        If command(1) = "started" Then
            info = "Game is successfully started. Please place your ships..."
            frm_lannet.Hide: connection.Enabled = False: playorder = 5: shippos(1, 0, 0) = "1": viewmaps: playorder = 1: leavegame.Enabled = True
            map(1).Visible = True: ships(0).Visible = True: ship1(0).Visible = True: waitfor: Exit Sub
        ElseIf command(1) = "youshoot" Then playorder = 0: viewmaps: waitfor: Exit Sub
        ElseIf command(1) = "cheat" Then info = playername & ": Your chance to win has decreased " & IIf(command(2) = "m", "at last turn", "at this turn") & ".": Exit Sub
        ElseIf command(1) = "message" Then
            If frm_chat.WindowState = 1 Or Screen.ActiveForm.Name <> "frm_chat" Then msgwink.Enabled = True: info = "Incoming message..."
            frm_chat.addmsg "<" & otherplayer & "> " & command(2)
        ElseIf command(1) = "sendship" Then
            temp0 = Split(command(2), ":")
            For Each temp1 In temp0
                temp2 = Split(temp1, "."): mapinfo(3, Val(temp2(1))) = CStr(temp2(0))
            Next
            placedship = True
            If numtok(data, "|") = 3 Then Exit Sub Else command(1) = command(3): GoTo 10
        ElseIf command(1) = "shot" Then
            If command(3) = "nuclear" Then nuclearresult = command(4) & "|" & command(5): info = otherplayer & " uses nuclear!": Beep Else info = otherplayer & " shoots!"
            playorder = 1: shootresult = command(2) & ":" & IIf(mapinfo(1, Val(command(2))) = "", "miss", mapinfo(1, Val(command(2)))): viewmaps: waitfor
            putarrows CStr(command(2)): targetarrows.Enabled = True
        End If
    ElseIf mymode = "lanclient" Or mymode = "netclient" Then
        If command(1) = "started" Then
            info = "Game is successfully started. Please place your ships...": frm_lannet.Hide
            connection.Enabled = False: playorder = 3: shippos(3, 0, 0) = "1": viewmaps: leavegame.Enabled = True
            ship1(0).Visible = False: putships -1: putships -2: putships -3: putships -4: putships -5: putships -6: waitfor: Exit Sub
        ElseIf command(1) = "shooting" Then waitfor "Waiting " & otherplayer & " to shoot!": Exit Sub
        ElseIf command(1) = "cheat" Then info = playername & ": Your chance to win has decreased " & IIf(command(2) = "m", "at last turn", "at this turn") & ".": Exit Sub
        ElseIf command(1) = "message" Then
            If frm_chat.WindowState = 1 Or Screen.ActiveForm.Name <> "frm_chat" Then msgwink.Enabled = True: info = "Incoming message..."
            frm_chat.addmsg "<" & otherplayer & "> " & command(2)
        ElseIf command(1) = "sendship" Then
            temp0 = Split(command(2), ":")
            For Each temp1 In temp0
                temp2 = Split(temp1, "."): mapinfo(1, Val(temp2(1))) = CStr(temp2(0))
            Next
            placedship = True
            If numtok(data, "|") = 3 Then Exit Sub Else command(1) = command(3): GoTo 10
        ElseIf command(1) = "shot" Then
            If command(3) = "nuclear" Then nuclearresult = command(4) & "|" & command(5): info = otherplayer & " uses nuclear!": Beep Else info = otherplayer & " shoots!"
            playorder = 3: shootresult = command(2) & ":" & IIf(mapinfo(3, Val(command(2))) = "", "miss", mapinfo(3, Val(command(2)))): viewmaps: waitfor
            putarrows CStr(command(2)): targetarrows.Enabled = True
        End If
    End If
End Sub

Private Sub sock_ConnectionRequest(ByVal requestID As Long)
    'accept game request if you are host
    If mymode = "lanhost" Or mymode = "nethost" Then sock.Close: sock.Accept requestID: sock.SendData "hnb|started"
End Sub

Private Sub sock_Close()
    'reports player that game connection is terminated
    If mymode <> "closer" And InStr(1, mymode, "winner") = 0 And InStr(1, mymode, "looser") = 0 Then
        Unload Me: Unload frm_chat: frm_main.Show: frm_lannet.sock.Close
        MsgBox "Otherside has left the game.", vbInformation, "Game Terminated"
    End If
End Sub

Private Sub leavegame_Click()
    'leaves game securely
    If sock.state = 7 And InStr(1, mymode, "winner") = 0 And InStr(1, mymode, "looser") = 0 Then
        reply = MsgBox("You have an open game connection. If you leave now, it will be terminated." & vbCrLf & "Are you sure to leave?", vbExclamation + vbYesNo + vbDefaultButton2, "Leaving Game")
        If reply = vbNo Then Exit Sub
    End If
    mymode = IIf(InStr(1, mymode, "winner") > 0 Or InStr(1, mymode, "looser") > 0, Left(mymode, Len(mymode) - 6), mymode)
    Select Case mymode
        Case "lanhost": mymode = "closer": frm_lannet.sock.Close: frm_lannet.lancreate_Click
        Case "lanclient": mymode = "closer": frm_lannet.sock.Close: frm_lannet.lanjoin_Click
        Case "nethost": mymode = "closer": frm_lannet.netrecreate.Caption = "Close Game": frm_lannet.sock.Close: frm_lannet.netrecreate_Click
        Case "netclient": mymode = "closer": frm_lannet.sock.Close: frm_lannet.netrejoin_Click
    End Select
    sock.Close: Unload Me: Unload frm_chat: frm_main.Show
End Sub

Private Function resetgame()
    'resets all the variables about game
    shiporiginalpos = Array("", "0.0:1.0", "-1.0:0.0:1.0", "-1.0:0.0:1.0:2.0", "-1.0:0.0:1.0:2.0:0.-1", "-2.0:-1.0:0.0:1.0:2.0:3.0", "-2.0:-1.0:0.0:1.0:-1.-1:0.-1:1.-1:2.-1")
    shootresult = "": shootresult2 = "": nuclearresult = "": nuclearresult2 = "": placedship = ""
    For i = 0 To 7: lastbox(i) = "": Next i
    For i = 0 To 3: For j = 0 To 6: For k = 0 To 4
        shippos(i, j, k) = ""
    Next k: Next j: Next i
    For i = 0 To 6
        With shipshot(i)
            .lastbox = ""
            .shotnum = 0
            .direction = 0
            .code = ""
            .done = 0
            .try = 0
        End With
    Next i
    rotateship
End Function

Private Function putships(ship As Integer)
    'draws ships and manages their shot numbers
    If ship > 10 Then
        ship = ship - 10
        ship1(ship - 1) = "0" & vbCrLf & "2": ship2(ship - 1) = "0" & vbCrLf & "3"
        ship3(ship - 1) = "0" & vbCrLf & "4": ship4(ship - 1) = "0" & vbCrLf & "5"
        ship5(ship - 1) = "0" & vbCrLf & "6": ship6(ship - 1) = "0" & vbCrLf & "8"
    ElseIf ship < 0 Then
        Select Case Abs(ship)
            Case 1: If ship1(0).Visible = True Then ship1(0).Visible = False: temp = "hide" Else ship1(0).Visible = True: temp = "show"
            Case 2: If ship2(0).Visible = True Then ship2(0).Visible = False: temp = "hide" Else ship2(0).Visible = True: temp = "show"
            Case 3: If ship3(0).Visible = True Then ship3(0).Visible = False: temp = "hide" Else ship3(0).Visible = True: temp = "show"
            Case 4: If ship4(0).Visible = True Then ship4(0).Visible = False: temp = "hide" Else ship4(0).Visible = True: temp = "show"
            Case 5: If ship5(0).Visible = True Then ship5(0).Visible = False: temp = "hide" Else ship5(0).Visible = True: temp = "show"
            Case 6: If ship6(0).Visible = True Then ship6(0).Visible = False: temp = "hide" Else ship6(0).Visible = True: temp = "show"
        End Select
        If temp = "hide" Then graphics ships(0), "shipback", Abs(ship + 1) Else graphics ships(0), "ships", Abs(ship + 1)
    ElseIf ship > 0 And ship < 7 Then
        temp2 = IIf(playorder = 1, 2, IIf(playorder = 3, 0, playorder))
        Select Case ship
            Case 1
                temp = Split(ship1(temp2), vbCrLf)
                If temp(0) < 2 Then ship1(temp2) = temp(0) + 1 & vbCrLf & "2": ship1(temp2).ForeColor = IIf(temp(0) = 0, &H80C0FF, ship1(temp2).ForeColor)
                If temp(0) + 1 = 2 Then temp1 = 1: ship1(temp2).ForeColor = &HFF&: lastbox(7) = IIf(playorder = 1 Or playorder = 3, remtok(lastbox(7), gettok(lastbox(7), 1, ":"), ":"), lastbox(7))
            Case 2
                temp = Split(ship2(temp2), vbCrLf)
                If temp(0) < 3 Then ship2(temp2) = temp(0) + 1 & vbCrLf & "3": ship2(temp2).ForeColor = IIf(temp(0) = 0, &H80C0FF, IIf(temp(0) = 1, &H80FF&, ship2(temp2).ForeColor))
                If temp(0) + 1 = 3 Then temp1 = 2: ship2(temp2).ForeColor = &HFF&: lastbox(7) = IIf(playorder = 1 Or playorder = 3, remtok(lastbox(7), gettok(lastbox(7), 1, ":"), ":"), lastbox(7))
            Case 3
                temp = Split(ship3(temp2), vbCrLf)
                If temp(0) < 4 Then ship3(temp2) = temp(0) + 1 & vbCrLf & "4": ship3(temp2).ForeColor = IIf(temp(0) = 0, &H80C0FF, IIf(temp(0) = 1, &H80FF&, IIf(temp(0) = 2, &H40C0&, ship3(temp2).ForeColor)))
                If temp(0) + 1 = 4 Then temp1 = 3: ship3(temp2).ForeColor = &HFF&: lastbox(7) = IIf(playorder = 1 Or playorder = 3, remtok(lastbox(7), gettok(lastbox(7), 1, ":"), ":"), lastbox(7))
            Case 4
                temp = Split(ship4(temp2), vbCrLf)
                If temp(0) < 5 Then ship4(temp2) = temp(0) + 1 & vbCrLf & "5": ship4(temp2).ForeColor = IIf(temp(0) = 0, &H80C0FF, IIf(temp(0) = 1 Or temp(0) = 2, &H80FF&, IIf(temp(0) = 3, &H40C0&, ship4(temp2).ForeColor)))
                If temp(0) + 1 = 5 Then temp1 = 4: ship4(temp2).ForeColor = &HFF&: lastbox(7) = IIf(playorder = 1 Or playorder = 3, remtok(lastbox(7), gettok(lastbox(7), 1, ":"), ":"), lastbox(7))
            Case 5
                temp = Split(ship5(temp2), vbCrLf)
                If temp(0) < 6 Then ship5(temp2) = temp(0) + 1 & vbCrLf & "6": ship5(temp2).ForeColor = IIf(temp(0) = 0 Or temp(0) = 1, &H80C0FF, IIf(temp(0) = 2 Or temp(0) = 3, &H80FF&, IIf(temp(0) = 4, &H40C0&, ship5(temp2).ForeColor)))
                If temp(0) + 1 = 6 Then temp1 = 5: ship5(temp2).ForeColor = &HFF&: lastbox(7) = IIf(playorder = 1 Or playorder = 3, remtok(lastbox(7), gettok(lastbox(7), 1, ":"), ":"), lastbox(7))
            Case 6
                temp = Split(ship6(temp2), vbCrLf)
                If temp(0) < 8 Then ship6(temp2) = temp(0) + 1 & vbCrLf & "8": ship6(temp2).ForeColor = IIf(temp(0) = 0 Or temp(0) = 1, &H80C0FF, IIf(temp(0) = 2 Or temp(0) = 3 Or temp(0) = 4, &H80FF&, IIf(temp(0) = 5 Or temp(0) = 6, &H40C0&, ship6(temp2).ForeColor)))
                If temp(0) + 1 = 8 Then temp1 = 6: ship6(temp2).ForeColor = &HFF&: lastbox(7) = IIf(playorder = 1 Or playorder = 3, remtok(lastbox(7), gettok(lastbox(7), 1, ":"), ":"), lastbox(7))
        End Select
        If temp1 > 0 Then graphics ships(temp2), "skull", temp1 - 1
    End If
End Function

Private Sub map_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    'for cheaters, little cheat codes
    'if it's your order, right click on your map, then
    'shows opponent's map, if you press Shift+M(ap)
    'gives 9 nuclear missiles, if you press Shift+N(uclear)
    If KeyCode = 77 And Shift = 1 Then
        Dim cheat(MAPVERTSQR, MAPHORZSQR), info, info2
        temp = "": info = IIf(playernum = 1, 3, IIf(playorder = 0 And playernum >= 2, 3, IIf(playorder = 2 And playernum >= 2, 1, "")))
        If info = "" Then Exit Sub Else info2 = IIf(info = 1, 2, 0)
        For i = 0 To MAPHORZSQR: For j = 0 To MAPVERTSQR
            box = Val(twolen(CStr(i)) & twolen(CStr(j)))
            cheat(j, i) = IIf(mapinfo(info, box) = "", "o", mapinfo(info, box))
            If playernum <> 2 Then cheat(j, i) = IIf(mapinfo(info2, box) = "miss", "=", IIf(mapinfo(info2, box) = "shot", "+", cheat(j, i)))
        Next j: Next i
        For i = 0 To MAPVERTSQR: For j = 0 To MAPHORZSQR
            temp = temp & IIf(cheat(i, j) = "miss", "=", IIf(cheat(i, j) = "shot", "+", cheat(i, j))): Next j
        temp = temp & vbCrLf: Next i
        If playernum = 3 Then sock.SendData "hnb|cheat|m"
        MsgBox "Opponent's Current Map" & vbCrLf & "(=:miss, +:shot, o:free)" & vbCrLf & String(MAPHORZSQR + 1, "_") & vbCrLf & temp, vbOKOnly, "You Cheater!"
    ElseIf KeyCode = 78 And Shift = 1 And (playorder = myorder Or (playorder = 2 And playernum >= 2)) Then
        If Val(missilenum(playorder)) < 9 Then
            If playernum = 3 Then sock.SendData "hnb|cheat|n"
            missilenum(playorder) = 9: Beep
        End If
    Else: Exit Sub
    End If
End Sub

Private Function sortbox(scode As String, Optional code As Integer)
    'sorts the ship shot boxes with respect to cols and rows
    Dim cols, rows, code1, code2, command, coord, basiccoord, codetmp
    temp = Split(scode, ":"): k = 0
    For Each coord In temp
        k = IIf(coord <> "", k + 1, k)
        cols = addtok(CStr(cols), CStr(Left(coord, 2)), ":") ': cols = Right(cols, Len(cols) - 1)
        rows = addtok(CStr(rows), CStr(Right(coord, 2)), ":") ': rows = Right(rows, Len(rows) - 1)
    Next
    command = IIf(numtok(CStr(cols), ":") > numtok(CStr(rows), ":"), "horz", "vert")
    For i = 0 To k
        For j = 0 To k - 2
            If command = "vert" And ((Val(Right(temp(j), 2)) > Val(Right(temp(j + 1), 2)) And Val(Left(temp(j), 2)) = Val(Left(temp(j + 1), 2))) Or (Val(Left(temp(j), 2)) > Val(Left(temp(j + 1), 2)))) Then coord = temp(j): temp(j) = temp(j + 1): temp(j + 1) = coord
            If command = "horz" And ((Val(Right(temp(j), 2)) = Val(Right(temp(j + 1), 2)) And Val(Left(temp(j), 2)) > Val(Left(temp(j + 1), 2))) Or (Val(Right(temp(j), 2)) > Val(Right(temp(j + 1), 2)))) Then coord = temp(j): temp(j) = temp(j + 1): temp(j + 1) = coord
        Next j
    Next i
    basiccoord = IIf(command = "horz", Right(temp(0), 2), Left(temp(0), 2))
    For i = 0 To k - 1
        coord = IIf(command = "horz", Right(temp(i), 2), Left(temp(i), 2))
        If coord = basiccoord Then code1 = code1 & ":" & temp(i) Else code2 = code2 & ":" & temp(i)
    Next i
    If Len(code1) > 0 Then code1 = Right(code1, Len(code1) - 1)
    If Len(code2) > 0 Then code2 = Right(code2, Len(code2) - 1)
    If numtok(CStr(code2), ":") > numtok(CStr(code1), ":") Then codetmp = code1: code1 = code2: code2 = codetmp
    sortbox = IIf(Len(code1) = 4 And Len(code2) = 4, code1 & ":" & code2, IIf(code = "2", code2, code1))
End Function

Private Function addtok(scode As String, token As String, id As String)
    'adds an item, in a token list, if it doesnot exist
    temp = Split(scode, id): n = numtok(scode, id): m = 0
    For i = 1 To n: m = IIf(temp(i - 1) = token, m + 1, m): Next i
    If m = 0 Then scode = scode & IIf(scode <> "", id, "") & token
    addtok = scode
End Function

Public Function remtok(scode As String, token As String, id As String)
    'removes the specified token form token list
    temp = Split(scode, id): temp2 = ""
    For Each temp1 In temp: temp2 = IIf(token <> temp1, addtok(CStr(temp2), CStr(temp1), ":"), CStr(temp2)): Next
    remtok = CStr(temp2)
End Function

Public Function numtok(scode As String, id As String)
    'returns the number of tokens in a token list
    temp = Split(scode, id): n = 0
    For Each temp1 In temp: n = n + 1: Next
    numtok = n
End Function

Public Function gettok(scode As String, token As Integer, id As String)
    'returns desired token
    If Len(scode) = 0 Then Exit Function
    temp = Split(scode, id): gettok = temp(token - 1)
End Function

Private Function verthorz(scode As String)
    'returns the direction of ship
    temp1 = Split(scode, "|"): temp2 = Split(temp1(0), ":")
    If Left(temp2(0), 2) = Left(temp2(1), 2) Then
        If Val(Right(temp2(0), 2)) > Val(Right(temp2(1), 2)) Then verthorz = 1 Else verthorz = 3
    ElseIf Right(temp2(0), 2) = Right(temp2(1), 2) Then
        If Val(Left(temp2(0), 2)) > Val(Left(temp2(1), 2)) Then verthorz = 0 Else verthorz = 2
    Else: verthorz = ""
    End If
End Function

Private Function changedir(direction As Integer, Optional degree As Integer)
    'returns the opposite direction
    Dim newdir As Integer
    Select Case direction
        Case 0: newdir = IIf(degree = 1, 1, 2)
        Case 1: newdir = IIf(degree = 1, 0, 3)
        Case 2: newdir = IIf(degree = 1, 1, 0)
        Case 3: newdir = IIf(degree = 1, 0, 1)
    End Select
    changedir = newdir
End Function

Private Function waitfor(Optional txt As String)
    'shows/hides wait screen
    If Not Len(txt) > 0 Then wait.Visible = False: Exit Function
    If Len(txt) > 0 Then wait.Visible = True: info = "Waiting for other player!": waittext = txt: click.Visible = IIf(playernum = 3, False, True)
End Function

Private Function twolen(text As String)
    'lengthen length of text to two
    twolen = IIf(Len(text) = 1, "0" & text, text)
End Function

Private Function changecoord(coord As String, rotate As Integer)
    'turns the coordinates 90 degree CW or CCW
    temp = Split(coord, ".")
    changecoord = IIf(rotate = 0, 0 - Val(temp(1)) & "." & temp(0), temp(1) & "." & 0 - Val(temp(0)))
End Function

Private Sub waittext_Click()
    'player changer
    wait_Click
End Sub

Private Sub chatwindow_click()
    'opens the chat window
    frm_chat.Show: frm_chat.WindowState = 0
End Sub

Private Sub msgicon_Dblclick()
    'opens the chat window
    frm_chat.Show: frm_chat.WindowState = 0
End Sub

Private Sub info_Change()
    'shows/hides info icon
    infoicon.Visible = IIf(info = "Incoming message..." Or msgwink.Enabled = True, False, True)
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'controls who wanted to close the window,
    'if user wants it ends the application, if code wants closes the window
    If UnloadMode = 1 Then
        frm_lannet.sock.Close: sock.Close: Unload Me: Unload frm_chat
    ElseIf UnloadMode = 0 Then
        If sock.state = 7 And InStr(1, mymode, "winner") = 0 And InStr(1, mymode, "looser") = 0 Then
            reply = MsgBox("You have an open game connection. If you close window, it will be terminated." & vbCrLf & "Are you sure to close window?", vbExclamation + vbYesNo + vbDefaultButton2, "Closing Window")
            If reply = vbNo Then Cancel = True Else frm_lannet.sock.Close: sock.Close: deletebuffer: End
        Else: deletebuffer: End
        End If
    Else: deletebuffer: End
    End If
End Sub

Private Sub pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'if mouse is out the back button or map change the graphic
    If backvar = True Then back.Picture = LoadResPicture(201, 0): backvar = False
    If lastbox(0) <> "0" And lastbox(0) <> "" Then mapicon map(playorder), lastbox(0), IIf(mapinfo(playorder, Val(lastbox(0))) = "", "blue", mapinfo(playorder, Val(lastbox(0)))): lastbox(0) = "0"
End Sub

Private Sub maptext_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'if mouse is out the map change the graphic
    If lastbox(0) <> "0" And lastbox(0) <> "" Then mapicon map(playorder), lastbox(0), IIf(mapinfo(playorder, Val(lastbox(0))) = "", "blue", mapinfo(playorder, Val(lastbox(0)))): lastbox(0) = "0"
End Sub

Private Sub back_Click()
    'back button is clicked, close window
    If sock.state = 7 And InStr(1, mymode, "winner") = 0 And InStr(1, mymode, "looser") = 0 Then
        reply = MsgBox("You have an open game connection. If you close window, it will be terminated." & vbCrLf & "Are you sure to close window?", vbExclamation + vbYesNo + vbDefaultButton2, "Closing Window")
        If reply = vbNo Then Exit Sub
    End If
    If InStr(1, mymode, "winner") > 0 Or InStr(1, mymode, "looser") > 0 Then mymode = Left(mymode, Len(mymode) - 6)
    Select Case mymode
        Case "lanhost": mymode = "closer": frm_lannet.sock.Close: frm_lannet.lancreate_Click
        Case "lanclient": mymode = "closer": frm_lannet.sock.Close: frm_lannet.lanjoin_Click
        Case "nethost": mymode = "closer": frm_lannet.netrecreate.Caption = "Close Game": frm_lannet.sock.Close: frm_lannet.netrecreate_Click
        Case "netclient": mymode = "closer": frm_lannet.sock.Close: frm_lannet.netrejoin_Click
    End Select
    sock.Close: Unload Me: Unload frm_chat: frm_main.Show
End Sub

Private Sub back_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'if mouse is over the back button change the graphic
    back.Picture = LoadResPicture(202, 0): backvar = True
End Sub


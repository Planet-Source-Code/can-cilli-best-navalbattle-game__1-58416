VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frm_lannet 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Howæbout NavalBattle"
   ClientHeight    =   3615
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5175
   Icon            =   "frm_lannet.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   241
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   345
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox back 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Left            =   4875
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
      Height          =   3615
      Left            =   0
      ScaleHeight     =   239
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   343
      TabIndex        =   0
      Top             =   0
      Width           =   5175
      Begin MSComctlLib.ImageList imglist 
         Left            =   720
         Top             =   0
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   7
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_lannet.frx":0CCA
               Key             =   "domain"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_lannet.frx":1266
               Key             =   "group"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_lannet.frx":1802
               Key             =   "network"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_lannet.frx":1B9E
               Key             =   "root"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_lannet.frx":213A
               Key             =   "server"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_lannet.frx":26D6
               Key             =   "ndscontainer"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frm_lannet.frx":2A2E
               Key             =   "tree"
            EndProperty
         EndProperty
      End
      Begin VB.Timer porttimer 
         Enabled         =   0   'False
         Interval        =   15000
         Left            =   360
         Top             =   0
      End
      Begin VB.CommandButton netshow 
         BackColor       =   &H0080C0FF&
         Caption         =   "Internet Game"
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
         Left            =   2550
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   600
         Width           =   2325
      End
      Begin VB.CommandButton lanshow 
         BackColor       =   &H0080C0FF&
         Caption         =   "LAN Game"
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
         TabIndex        =   4
         Top             =   600
         Width           =   2325
      End
      Begin VB.PictureBox picnet 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2535
         Left            =   240
         ScaleHeight     =   169
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   313
         TabIndex        =   3
         Top             =   960
         Visible         =   0   'False
         Width           =   4695
         Begin VB.CommandButton lastip 
            BackColor       =   &H0080C0FF&
            Caption         =   "Last IP"
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
            Left            =   3720
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   960
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.CommandButton netrejoin 
            BackColor       =   &H0080C0FF&
            Caption         =   "Join"
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
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   31
            Top             =   2160
            Visible         =   0   'False
            Width           =   4455
         End
         Begin VB.CommandButton netrecreate 
            BackColor       =   &H0080C0FF&
            Caption         =   "Create"
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
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   2160
            Visible         =   0   'False
            Width           =   4455
         End
         Begin VB.ComboBox ip 
            BackColor       =   &H00C0E0FF&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   330
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   960
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.TextBox destip 
            BackColor       =   &H00C0E0FF&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   315
            Index           =   3
            Left            =   3120
            MaxLength       =   3
            TabIndex        =   20
            Top             =   960
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox destip 
            BackColor       =   &H00C0E0FF&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   315
            Index           =   2
            Left            =   2520
            MaxLength       =   3
            TabIndex        =   19
            Top             =   960
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox destip 
            BackColor       =   &H00C0E0FF&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   315
            Index           =   1
            Left            =   1920
            MaxLength       =   3
            TabIndex        =   18
            Top             =   960
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.TextBox destip 
            BackColor       =   &H00C0E0FF&
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   315
            Index           =   0
            Left            =   1320
            MaxLength       =   3
            TabIndex        =   17
            Top             =   960
            Visible         =   0   'False
            Width           =   495
         End
         Begin VB.CommandButton netjoin 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            Caption         =   "Join a Game"
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
            Left            =   2340
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   480
            Width           =   2235
         End
         Begin VB.CommandButton netcreate 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            Caption         =   "Create a Game"
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
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   480
            Width           =   2235
         End
         Begin VB.Label netclientname 
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
            TabIndex        =   29
            Top             =   1755
            Width           =   3255
         End
         Begin VB.Label nethostname 
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
            TabIndex        =   28
            Top             =   1395
            Width           =   3255
         End
         Begin VB.Label destipdot 
            BackStyle       =   0  'Transparent
            Caption         =   "."
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   2
            Left            =   3000
            TabIndex        =   23
            Top             =   1020
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label destipdot 
            BackStyle       =   0  'Transparent
            Caption         =   "."
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   1
            Left            =   2400
            TabIndex        =   22
            Top             =   1020
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label destipdot 
            BackStyle       =   0  'Transparent
            Caption         =   "."
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   255
            Index           =   0
            Left            =   1800
            TabIndex        =   21
            Top             =   1020
            Visible         =   0   'False
            Width           =   135
         End
         Begin VB.Label textnet 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   495
            Left            =   120
            TabIndex        =   16
            Top             =   30
            Width           =   3135
         End
         Begin VB.Label txtdestip 
            BackStyle       =   0  'Transparent
            Caption         =   "Destination IP:"
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
            Left            =   240
            TabIndex        =   15
            Top             =   1080
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label txtyourip 
            BackStyle       =   0  'Transparent
            Caption         =   "Your IP:"
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
            Left            =   240
            TabIndex        =   14
            Top             =   1080
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label Label2 
            BackStyle       =   0  'Transparent
            Caption         =   "Internet Game"
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
            Left            =   3525
            TabIndex        =   11
            Top             =   30
            Width           =   1215
         End
      End
      Begin VB.PictureBox piclan 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2535
         Left            =   240
         ScaleHeight     =   169
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   313
         TabIndex        =   2
         Top             =   960
         Visible         =   0   'False
         Width           =   4695
         Begin VB.CommandButton lanstart 
            BackColor       =   &H0080C0FF&
            Caption         =   "Start Game"
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
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   2160
            Visible         =   0   'False
            Width           =   4455
         End
         Begin VB.CommandButton lanjoin 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            Caption         =   "Join a Game on Selected Computer"
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
            Left            =   2160
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   480
            Width           =   2415
         End
         Begin VB.CommandButton lancreate 
            Appearance      =   0  'Flat
            BackColor       =   &H0080C0FF&
            Caption         =   "Create Game && Wait Player"
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
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   480
            Width           =   2055
         End
         Begin MSComctlLib.TreeView tree 
            Height          =   1455
            Left            =   120
            TabIndex        =   7
            Top             =   960
            Visible         =   0   'False
            Width           =   4455
            _ExtentX        =   7858
            _ExtentY        =   2566
            _Version        =   393217
            Indentation     =   265
            LabelEdit       =   1
            Style           =   7
            ImageList       =   "imglist"
            BorderStyle     =   1
            Appearance      =   0
         End
         Begin VB.Label lanclientname 
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
            Left            =   690
            TabIndex        =   25
            Top             =   1485
            Width           =   2655
         End
         Begin VB.Label lanhostname 
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
            Left            =   690
            TabIndex        =   24
            Top             =   1110
            Width           =   2775
         End
         Begin VB.Label textlan 
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   162
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   495
            Left            =   120
            TabIndex        =   10
            Top             =   30
            Width           =   3375
         End
         Begin VB.Label label1 
            BackStyle       =   0  'Transparent
            Caption         =   "LAN Game"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00808080&
            Height          =   255
            Left            =   3840
            TabIndex        =   6
            Top             =   30
            Width           =   855
         End
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
   End
End
Attribute VB_Name = "frm_lannet"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private NetRoot As NetResource
Dim backvar, connection
Private Const MAX_IP = 5   'To make a buffer... i dont think you have more than 5 ip on your pc..
Private Type IPINFO
    dwAddr As Long   ' IP address
    dwIndex As Long '  interface index
    dwMask As Long ' subnet mask
    dwBCastAddr As Long ' broadcast address
    dwReasmSize  As Long ' assembly size
End Type
Private Type MIB_IPADDRTABLE
    dEntrys As Long   'number of entries in the table
    mIPInfo(MAX_IP) As IPINFO  'array of IP address entries
End Type
Private Type IP_Array
    mBuffer As MIB_IPADDRTABLE
    BufferLen As Long
End Type
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)
Private Declare Function GetIpAddrTable Lib "IPHlpApi" (pIPAdrTable As Byte, pdwSize As Long, ByVal Sort As Long) As Long

Private Sub Form_Load()
    'initializing window graphics
    graphics Me.pic, "botback"
    graphics Me.pic, "smalltopback"
    graphics Me.pic, "smallicon"
    graphics Me.picnet, "botback"
    graphics Me.picnet, "internet"
    graphics Me.piclan, "botback"
    graphics Me.piclan, "network"
    back.Picture = LoadResPicture(201, 0): backvar = False: timeout = False
    If sock.state <> sckClosed Then sock.Close
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    'controls who wanted to close the window,
    'if user wants it ends the application, if code wants closes the window
    If UnloadMode = 1 Then sock.Close: Unload Me Else deletebuffer: End
End Sub

Public Sub lancreate_Click()
    If lancreate.Caption = "Close Game" Then
        If sock.state <> sckClosing And sock.state <> sckClosed Then
            reply = MsgBox("Are you sure to close the game?", vbQuestion + vbDefaultButton2 + vbYesNo, "Close Game?")
            If reply = vbNo Then Exit Sub
            mymode = "closer"
        End If
        sock.Close: lancreate.Caption = "Create Game && Wait Player": lannetshow True: lanstart.Visible = False
        textlan = IIf(mymode = "closer", "Game closed.", textlan): lanjoin.Enabled = True: lanhostname = "-": lanclientname = "-"
    Else
        playernum = 3: playername = "": frm_newgame.Show 1
        If playername <> "" Then
            tree.Visible = False: lanjoin.Enabled = False: lancreate.Caption = "Close Game"
            graphics Me.piclan, "player1l": lanhostname = playername: lannetshow False
            graphics Me.piclan, "player2l": lanclientname = "<Waiting for player>"
            sock.LocalPort = port("new"): sock.Close: sock.Listen: mymode = "lanhost"
            textlan = "Game created on port #" & sock.LocalPort - 41670 & "!" & vbCrLf & "Waiting for a player..."
        End If
    End If
End Sub

Public Sub lanjoin_Click()
    If lanjoin.Caption = "Close Connection" Then
        If sock.state <> sckClosing And sock.state <> sckClosed Then
            reply = MsgBox("Are you sure to close the connection?", vbQuestion + vbDefaultButton2 + vbYesNo, "Close Connection?")
            If reply = vbNo Then Exit Sub
            mymode = "closer": connection = ""
        End If
        sock.Close: lanjoin.Caption = "Join a Game on Selected Computer": porttimer.Enabled = False
        textlan = IIf(mymode = "closer", "Connection closed.", textlan): lannetshow True: lancreate.Enabled = True: Me.MousePointer = 1
    Else
        tree.Visible = True: tree.Nodes.Clear: tree.Refresh: textlan = "Please be patient...": DoEvents
        Dim nX As NetResource, nodX As Node
        Set NetRoot = New NetResource   ' Create a new NetResource object. By default it will be the network root
        Set nodX = tree.Nodes.Add(, , "_ROOT", "Entire Network", "root", "root")  ' Add a node into the tree for it
        nodX.Tag = "Y" ' Set populated flag to "Y" since we populate this one immediately
        ' Populate the top level of objects under "Entire Network"
        For Each nX In NetRoot.Children
            Set nodX = tree.Nodes.Add("_ROOT", tvwChild, nX.RemoteName, StrConv(nX.ShortName, vbProperCase), LCase(nX.ResourceTypeName), LCase(nX.ResourceTypeName))
            nodX.Tag = "N"  ' We haven't populated the nodes underneath this one yet, so set its flag to "N"
            tree.Nodes.Add nodX.Key, tvwChild, nodX.Key + "_FAKE", "FAKE", "server", "server" ' Create a fake node under it so that the treeview gives the "+" symbol
            nodX.EnsureVisible: DoEvents
        Next
        DoEvents: textlan = "Double click on a remote computer to join...": timeout = False
     End If
End Sub

Private Sub lanshow_Click()
    picnet.Visible = False
    piclan.Visible = True
End Sub

Private Sub lanstart_Click()
    puttext "Starting Game...": sock.SendData "hnb|start": DoEvents: frm_game.Show: frm_lannet.Hide
End Sub

Private Sub lastip_Click()
    'restores last ip
    lastipnum = GetSetting("Navalbattle", "Settings", "Last IP")
    lastiptok = frm_game.numtok(CStr(lastipnum), ".")
    For i = 0 To 3
        If lastiptok = 4 Then destip.Item(i) = frm_game.gettok(CStr(lastipnum), i + 1, ".")
    Next i
End Sub

Public Sub netrecreate_Click()
    If netrecreate.Caption = "Create" Then
        playernum = 3: playername = "": frm_newgame.Show 1
        If playername <> "" Then
            netrecreate.Caption = "Close Game": netcreate.Enabled = False: netjoin.Enabled = False
            nethostname = playername: netclientname = "<Waiting for player>": lannetshow False
            sock.LocalPort = port("new"): sock.Close: sock.Listen: mymode = "nethost"
            textnet = "Created on port #" & sock.LocalPort - 41670 & "!" & vbCrLf & "You have to tell your IP to your friend!"
        End If
    ElseIf netrecreate.Caption = "Close Game" Then
        If sock.state <> sckClosing And sock.state <> sckClosed Then
            reply = MsgBox("Are you sure to close the game?", vbQuestion + vbDefaultButton2 + vbYesNo, "Close Game?")
            If reply = vbNo Then Exit Sub
            mymode = "closer"
        End If
        sock.Close: netrecreate.Caption = "Create": netcreate.Enabled = True: netjoin.Enabled = True: lannetshow True
        textnet = IIf(mymode = "closer", "Game closed.", textnet): nethostname = "-": netclientname = "-"
    ElseIf netrecreate.Caption = "Start Game" Then
        puttext "Starting Game...": sock.SendData "hnb|start": frm_game.Show: frm_lannet.Hide
    End If
End Sub

Public Sub netrejoin_Click()
    On Error Resume Next
    If netrejoin.Caption = "Join" Then
        If destip(0) > 255 Or destip(0) < 0 Or IsNumeric(destip(0)) = False Then Beep: destip(0).SetFocus: destip(0).SelStart = 0: destip(0).SelLength = Len(destip(0)): Exit Sub
        If destip(1) > 255 Or destip(1) < 0 Or IsNumeric(destip(1)) = False Then Beep: destip(1).SetFocus: destip(1).SelStart = 0: destip(1).SelLength = Len(destip(1)): Exit Sub
        If destip(2) > 255 Or destip(2) < 0 Or IsNumeric(destip(2)) = False Then Beep: destip(2).SetFocus: destip(2).SelStart = 0: destip(2).SelLength = Len(destip(2)): Exit Sub
        If destip(3) > 255 Or destip(3) < 0 Or IsNumeric(destip(3)) = False Then Beep: destip(3).SetFocus: destip(3).SelStart = 0: destip(3).SelLength = Len(destip(3)): Exit Sub
        playernum = 3: playername = "": frm_newgame.Show 1
        If playername <> "" Then
            destip(0).Enabled = False: destip(1).Enabled = False: destip(2).Enabled = False: destip(3).Enabled = False
            SaveSetting "Navalbattle", "Settings", "Last IP", destip(0) & "." & destip(1) & "." & destip(2) & "." & destip(3)
            netrejoin.Caption = "Close Connection": netcreate.Enabled = False: netjoin.Enabled = False: lannetshow False: lastip.Enabled = False
            netclientname = playername: nethostname = "<Waiting to Connect>": sock.Close: mymode = "netclient": connection = "active"
            sock.Connect destip(0) & "." & destip(1) & "." & destip(2) & "." & destip(3), 41670: porttimer.Enabled = True
            textnet = "Trying to connect on port #0..." & vbCrLf & "Please be patient...": Me.MousePointer = 13
        End If
    ElseIf netrejoin.Caption = "Close Connection" Then
        If sock.state <> sckClosing And sock.state <> sckClosed Then
            reply = MsgBox("Are you sure to close the connection?", vbQuestion + vbDefaultButton2 + vbYesNo, "Close Connection?")
            If reply = vbNo Then Exit Sub
            mymode = "closer": connection = "": porttimer.Enabled = False
        End If
        destip(0).Enabled = True: destip(1).Enabled = True: destip(2).Enabled = True: destip(3).Enabled = True: lannetshow True
        sock.Close: netrejoin.Caption = "Join": netcreate.Enabled = True: Me.MousePointer = 1: lastip.Enabled = True
        textnet = IIf(mymode = "closer", "Connection closed.", textnet): netjoin.Enabled = True: nethostname = "-": netclientname = "-"
    End If
End Sub

Private Sub netcreate_Click()
    'shows the related controls about net host
    txtyourip.Visible = True: ip.Visible = True: netrejoin.Visible = False
    txtdestip.Visible = False: nethostname = "": netclientname = "": lastip.Visible = False
    graphics Me.picnet, "player1n": graphics Me.picnet, "player2n"
    For i = 0 To 3
        destip.Item(i).Visible = False
        If i < 3 Then destipdot.Item(i).Visible = False
    Next i
    findip
End Sub

Private Sub netjoin_Click()
    'shows the related controls about net client
    txtyourip.Visible = False: ip.Visible = False: txtdestip.Visible = True
    netrecreate.Visible = False: nethostname = "": netclientname = "": netrejoin.Visible = True
    graphics Me.picnet, "player1n": graphics Me.picnet, "player2n"
    lastipnum = GetSetting("Navalbattle", "Settings", "Last IP"): lastiptok = frm_game.numtok(CStr(lastipnum), ".")
    If lastiptok = 4 Then lastip.Visible = True
    For i = 0 To 3
        destip.Item(i).Visible = True
        If i < 3 Then destipdot.Item(i).Visible = True
    Next i
    textnet = "You must enter the host's IP as destination IP."
End Sub

Private Sub netshow_Click()
    piclan.Visible = False
    picnet.Visible = True
End Sub

Private Sub NodeExpand(Node As MSComctlLib.Node)
    ' Distinguish between expansion of a network object or a file system folder as seen over the network
    Dim FSO As Scripting.FileSystemObject
    Dim NWFolder As Scripting.Folder
    Dim FilX As Scripting.File, DirX As Scripting.Folder
    Dim tNod As Node, isFSFolder As Boolean
    ' Remove the fake node used to force the treeview to show the "+" icon
    tree.Nodes.Remove Node.Key + "_FAKE": DoEvents
    ' Search up through the tree, noting the node keys so that we can then locate the NetResource object
    ' under NetRoot.
    Dim pS As String, kPath() As String, nX As NetResource, i As Integer, tX As NetResource
    Set tNod = Node ' Start at the node that was expanded
    Do While Not tNod.Parent Is Nothing ' Proceed up the tree using parent references, each time saving the node key to the string pS
        pS = tNod.Key + "|" + pS
        Set tNod = tNod.Parent
    Loop
    ' String pS is now of the form "<Node Key>|<Node Key>|<Node Key>"
    ' Split this into an array using the VB6 Split function
    kPath = Split(pS, "|"): DoEvents
    Set nX = NetRoot
    ' Now loop through this array, this time following down the tree of NetResource objects from NetRoot to the child NetResource object that corresponds to
    ' the node the user clicked
    For i = 0 To UBound(kPath) - 1
        Set nX = nX.Children(kPath(i)): DoEvents
    Next
    ' Now that we know both the node and the corresponding NetResource we can enumerate the children and add the nodes
    For Each tX In nX.Children
        Set tNod = tree.Nodes.Add(nX.RemoteName, tvwChild, tX.RemoteName, StrConv(tX.ShortName, vbProperCase), LCase(tX.ResourceTypeName), LCase(tX.ResourceTypeName))
        tNod.Tag = "N": DoEvents
        ' Add fake nodes to all new nodes except when they're printers (you can always be sure a printer never has children)
        If tX.ResourceType <> Printer And tX.ResourceType <> Server Then tree.Nodes.Add tX.RemoteName, tvwChild, tX.RemoteName + "_FAKE", "FAKE", "server", "server"
    Next
    tree.Refresh  ' Refresh the view
    Node.Tag = "Y"  ' Set the tag to "Y" to denote that this node has been expanded and populated
End Sub

Private Sub porttimer_Timer()
    'changes port after 15 second waiting
    porttimer.Enabled = False: portchange
End Sub

Private Function portchange()
    'changes port if a connection is not set
    On Error GoTo 10
    Dim hostaddress, hostport
    hostaddress = IIf(sock.RemoteHost <> "", sock.RemoteHost, IIf(sock.RemoteHostIP <> "", sock.RemoteHostIP, "connectionerror"))
    If hostaddress = "connectionerror" Then GoTo 10
5   sock.Close: hostport = Val(sock.RemotePort) + 1: puttext "Trying to connect on port #" & Val(hostport) - 41670 & "..." & vbCrLf & "Please be patient...": DoEvents
    If hostport = 41676 Then connection = "": puttext "Remote computer hasn't created a game or couldn't find address.": GoTo 20
    sock.Connect hostaddress, hostport: porttimer.Enabled = True: Exit Function
10  If Err = 10048 Then GoTo 5
    sock.Close: puttext "An error occured about IP! Please try to connect again soon!"
20  If mymode = "netclient" Then sock.Close: netrejoin_Click: Exit Function
    If mymode = "lanclient" Then sock.Close: lanjoin_Click: Exit Function
End Function

Private Sub sock_Close()
    'returns controls & buttons to active
    If mymode <> "closer" And connection <> "active" Then
        puttext "Other side closed connection."
        Select Case mymode
            Case "lanhost": sock.Close: lancreate_Click
            Case "lanclient": sock.Close: lanjoin_Click
            Case "nethost": netrecreate.Caption = "Close Game": sock.Close: netrecreate_Click
            Case "netclient": sock.Close: netrejoin_Click
        End Select
    End If
End Sub

Private Sub sock_Connect()
    'stops port search and records the port
    If mymode = "netclient" Or mymode = "lanclient" Then Me.MousePointer = 1: connection = "": porttimer.Enabled = False: port sock.RemotePort + 1
End Sub

Private Sub sock_ConnectionRequest(ByVal requestID As Long)
    'starts connection
    If mymode = "lanhost" Or mymode = "nethost" Then sock.Close: sock.Accept requestID: sock.SendData "hnb|requestname"
End Sub

Private Sub sock_DataArrival(ByVal bytesTotal As Long)
    'manages all data transfer
    Dim data As String
    Dim command
    sock.GetData data
    If Left(data, 3) <> "hnb" Then Exit Sub
    command = Split(data, "|")
    If mymode = "lanhost" Or mymode = "nethost" Then
        If command(1) = "requestname" Then
            sock.SendData "hnb|myname|" & playername: Exit Sub
        ElseIf command(1) = "myname" Then
            If command(2) = playername Then sock.SendData "hnb|changename": Exit Sub
            puttext "A player joined your game!"
            If mymode = "lanhost" Then lanclientname = command(2): lanstart.Visible = True Else netclientname = command(2): netrecreate.Caption = "Start Game"
            sock.SendData "hnb|welcome|" & playername: Exit Sub
        ElseIf command(1) = "error" Then
            Unload frm_game: frm_lannet.Show: sock.Close
            If mymode = "lanhost" Then lancreate_Click Else netrecreate_Click
            puttext "Connection attempt timed out. Please try again.": Exit Sub
        End If
    ElseIf mymode = "lanclient" Or mymode = "netclient" Then
        If command(1) = "requestname" Then
            sock.SendData "hnb|myname|" & playername: Exit Sub
        ElseIf command(1) = "changename" Then
            MsgBox "You have the same name with host." & vbCrLf & "Please change!", vbExclamation, "Change Name": oldname = playername
            Do: playernum = 3: frm_newgame.Show 1: Loop Until playername <> oldname
            sock.SendData "hnb|myname|" & playername
            If mymode = "lanclient" Then lanclientname = playername Else netclientname = playername
            Exit Sub
        ElseIf command(1) = "welcome" Then
            If mymode = "lanclient" Then lanhostname = command(2) Else nethostname = command(2)
            puttext "Successfully joined game. Waiting host to start game...": Exit Sub
        ElseIf command(1) = "start" Then puttext "Starting Game...": DoEvents: frm_game.Show: Exit Sub
        ElseIf command(1) = "error" Then
            Unload frm_game: frm_lannet.Show: sock.Close
            If mymode = "lanclient" Then lanjoin_Click Else netrejoin_Click
            puttext "Connection attempt timed out. Please try again.": Exit Sub
        End If
    End If
End Sub

Private Sub sock_Error(ByVal number As Integer, Description As String, ByVal scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    'manages all connection errors
    On Error GoTo 10
    If connection = "active" Then porttimer.Enabled = False: portchange: Exit Sub
    If lanjoin.Caption = "Close Connection" Then
        sock.Close: lanjoin_Click: GoTo 10
    ElseIf lancreate.Caption = "Close Game" Then sock.Close: lancreate_Click: GoTo 10
    ElseIf netrejoin.Caption = "Close Connection" Then sock.Close: netrejoin_Click: GoTo 10
    ElseIf netrecreate.Caption = "Close Game" Then sock.Close: netrecreate_Click: GoTo 10
    End If: Exit Sub
10  If number = 10061 Then
        puttext "Error! Remote computer hasn't created a game!": Exit Sub
    ElseIf number = 11001 Then puttext "Error! Couldn't established a connection!": Exit Sub
    ElseIf number = 10048 Then puttext "Error! Remote computer's address is in use!": Exit Sub
    Else: puttext "Error!": MsgBox number & ":" & Description
    End If
End Sub

Private Sub tree_DblClick()
    'joins to a network client
    If tree.SelectedItem.Image = "server" Then
        remoteaddr = Mid(tree.SelectedItem.FullPath, InStrRev(tree.SelectedItem.FullPath, "\") + 1, Len(tree.SelectedItem.FullPath))
        If StrConv(remoteaddr, vbLowerCase) = sock.LocalHostName Then textlan = "It is your computer. Choose another...": Exit Sub
        playernum = 3: playername = "": frm_newgame.Show 1
        If playername <> "" Then
            lancreate.Enabled = False: lanjoin.Caption = "Close Connection": lannetshow False: lanclientname = playername
            sock.Close: sock.Connect StrConv(remoteaddr, vbLowerCase), 41670: mymode = "lanclient": connection = "active"
            textlan = "Trying to connect on port #0..." & vbCrLf & "Please be patient...": Me.MousePointer = 13: porttimer.Enabled = True
        End If
    End If
End Sub

Private Sub tree_Expand(ByVal Node As MSComctlLib.Node)
    'expands nodes in tree
    If Node.Tag = "N" Then NodeExpand Node
End Sub

Public Function ConvertAddressToString(longAddr As Long) As String
    'converts a Long to a string
    Dim myByte(3) As Byte
    Dim Cnt As Long
    CopyMemory myByte(0), longAddr, 4
    For Cnt = 0 To 3
        ConvertAddressToString = ConvertAddressToString + CStr(myByte(Cnt)) + "."
    Next Cnt
    ConvertAddressToString = Left$(ConvertAddressToString, Len(ConvertAddressToString) - 1)
End Function

Public Sub findip()
    'returns available IPs on your computer
    Dim Ret As Long
    Dim bBytes() As Byte
    Dim Listing As MIB_IPADDRTABLE
    On Error GoTo 10
    GetIpAddrTable ByVal 0&, Ret, True
    If Ret <= 0 Then Exit Sub
    ReDim bBytes(0 To Ret - 1) As Byte
    'retrieve the data
    GetIpAddrTable bBytes(0), Ret, False
    'Get the first 4 bytes to get the entry's.. ip installed
    CopyMemory Listing.dEntrys, bBytes(0), 4: ip.Clear
    For tel = 0 To Listing.dEntrys - 1
        'Copy whole structure to Listing..
        CopyMemory Listing.mIPInfo(tel), bBytes(4 + (tel * Len(Listing.mIPInfo(0)))), Len(Listing.mIPInfo(tel))
        If ConvertAddressToString(Listing.mIPInfo(tel).dwMask) <> "0.0.0.0" And ConvertAddressToString(Listing.mIPInfo(tel).dwMask) <> "255.0.0.0" Then ip.AddItem ConvertAddressToString(Listing.mIPInfo(tel).dwAddr)
    Next
    If ip.ListCount = 0 Then GoTo 10 Else ip.ListIndex = 0
    textnet = "": netrecreate.Visible = True: netrecreate.Caption = "Create"
    Exit Sub
10  textnet = "Cannot find an IP. You won't be able to create a game."
End Sub

Public Function puttext(ptext As String)
    'Decides where (lan/net) to report
    If Left(mymode, 3) = "lan" Then textlan = ptext: Exit Function
    If Left(mymode, 3) = "net" Then textnet = ptext
End Function

Private Function lannetshow(state As Boolean)
    'enables/disables main buttons
    netshow.Enabled = state: lanshow.Enabled = state
End Function

Private Sub destip_KeyUp(Index As Integer, KeyCode As Integer, Shift As Integer)
    'moves cursor to the next space
    If Index < 3 And Len(destip(Index)) = 3 Then destip(Index + 1).SetFocus: destip(Index + 1).SelStart = 0: destip(Index + 1).SelLength = Len(destip(Index + 1))
    If Index > 0 And Len(destip(Index)) = 0 And KeyCode = 8 Then destip(Index - 1).SetFocus
    If Index = 3 And Len(destip(Index)) = 3 Then netrejoin.SetFocus
End Sub

Private Sub pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'if mouse is out the back button change the graphic
    If backvar = True Then back.Picture = LoadResPicture(201, 0): backvar = False
End Sub

Private Sub back_Click()
    'back button is clicked, close window
    If lanshow.Enabled = False Then
        reply = MsgBox("You have an open connection." & vbCrLf & "Are you sure to close the window?", vbQuestion + vbDefaultButton2 + vbYesNo, "Close Window")
        If reply = vbNo Then Exit Sub
        sock.Close
    End If
    Unload Me: frm_main.Show
End Sub

Private Sub back_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'if mouse is over the back button change the graphic
    back.Picture = LoadResPicture(202, 0): backvar = True
End Sub


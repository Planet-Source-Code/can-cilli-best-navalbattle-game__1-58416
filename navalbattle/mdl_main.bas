Attribute VB_Name = "mdl_main"
'general vars, backpage is used for the windows
'playernum is used for identifying single or multi player game
'playername is used for identifying player on lan/net games
'api calls for graphical design
Public backpage, playernum, playername, mymode, timeout, lastobj
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hdc As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Public myBackBuffer As Long, myBufferBMP As Long
Public myMapBackBuffer(3) As Long, myMapBufferBMP(3) As Long, mySprite As Long

Sub Main()
    On Error Resume Next
    'load the image settings which we use everywhere for graphics
    mySprite = CreateCompatibleDC(GetDC(0))
    SelectObject mySprite, LoadResPicture(101, 0)
    myBackBuffer = CreateCompatibleDC(GetDC(0))
    myBufferBMP = CreateCompatibleBitmap(GetDC(0), 510, 500)
    SelectObject myBackBuffer, myBufferBMP
    BitBlt myBackBuffer, 0, 0, 510, 500, 0, 0, 0, vbWhiteness
    For i = 0 To 3
        myMapBackBuffer(i) = CreateCompatibleDC(GetDC(0))
        myMapBufferBMP(i) = CreateCompatibleBitmap(GetDC(0), 320, 320)
        SelectObject myMapBackBuffer(i), myMapBufferBMP(i)
        BitBlt myMapBackBuffer(i), 0, 0, 320, 320, 0, 0, 0, vbWhiteness
    Next i
    'set first registry entries
    If GetSetting("Navalbattle", "Settings", "Player1") = "" Then SaveSetting "Navalbattle", "Settings", "Player1", "Player 1"
    If GetSetting("Navalbattle", "Settings", "Player2") = "" Then SaveSetting "Navalbattle", "Settings", "Player2", "Player 2"
    'shows the main window
    frm_main.Show
End Sub

Public Function port(command As String)
    'manages registry entries
    Dim today, portnum
    today = GetSetting("Navalbattle", "Settings", "Day")
    If Val(today) <> Weekday(Date) Then SaveSetting "Navalbattle", "Settings", "Day", Weekday(Date): SaveSetting "Navalbattle", "Settings", "Port", "41670": port = 41670: Exit Function
    portnum = GetSetting("Navalbattle", "Settings", "Port"): portnum = IIf(portnum = "", 41670, portnum)
    If command = "new" Then
        SaveSetting "Navalbattle", "Settings", "Port", IIf(Val(portnum) = 41675, "41670", Val(portnum) + 1): port = portnum
    Else: SaveSetting "Navalbattle", "Settings", "Port", IIf(Val(command) = 41675, "41670", Val(command))
    End If
End Function

Public Function base(number, cbase)
    'mathematical base arithmetic
    base = IIf(Abs(number) Mod cbase = 0, cbase, Abs(number) Mod cbase)
End Function

Public Function graphics(obj As PictureBox, img As String, Optional arg As Integer)
    'window graphic initializer
    Dim pos
    Select Case img
        Case "bigicon": pos = Array(5, 2, 269, 38, 0, 0, 269, 38)
        Case "mainicon": pos = Array((obj.ScaleWidth - 269) / 2, 2, 269, 38, 0, 0, 269, 38)
        Case "waiticon": pos = Array((obj.ScaleWidth - 269) / 2, 95, 269, 38, 0, 0, 269, 38)
        Case "smallicon": pos = Array(3, 2, 87, 22, 0, 39, 87, 22)
        Case "player1": pos = Array(15, 63, 20, 20, 188, 40, 20, 20)
        Case "player2": pos = Array(15, 88, 20, 20, 208, 40, 20, 20)
        Case "player1l": pos = Array(15, 73, 20, 20, 188, 40, 20, 20)
        Case "player2l": pos = Array(15, 98, 20, 20, 208, 40, 20, 20)
        Case "player1g": pos = Array(19, 53, 20, 20, 188, 40, 20, 20)
        Case "player2g": pos = Array(195, 53, 20, 20, 208, 40, 20, 20)
        Case "player1n": pos = Array(15, 91, 20, 20, 188, 40, 20, 20)
        Case "player2n": pos = Array(15, 116, 20, 20, 208, 40, 20, 20)
        Case "message": pos = Array(0, 0, 20, 20, 228, 40, 20, 20)
        Case "internet": pos = Array(212, 0, 20, 20, 268, 40, 20, 20)
        Case "network": pos = Array(232, 0, 20, 20, 248, 40, 20, 20)
        Case "internetg": pos = Array(369, 53, 20, 20, 268, 40, 20, 20)
        Case "networkg": pos = Array(369, 53, 20, 20, 248, 40, 20, 20)
        Case "singleplyr": pos = Array(369, 53, 20, 20, 288, 40, 20, 20)
        Case "twoplayer": pos = Array(369, 53, 20, 20, 308, 40, 20, 20)
        Case "infoicon": pos = Array(0, 0, 20, 20, 328, 40, 20, 20)
        Case "missile1": pos = Array(0, 14, 58, 12, 88, 39, 58, 10)
        Case "missile2": pos = Array(0, 26, 58, 12, 88, 50, 58, 10)
        Case "ships": pos = Array(0, arg * 44, 58, 44, arg * 59, 62, 58, 44)
        Case "skull": pos = Array(0, arg * 44, 19, 19, 168, 40, 19, 19)
        Case "bigtopback": pos = Array(0, 0, obj.ScaleWidth, 42, 147, 40, 10, 10)
        Case "smalltopback": pos = Array(0, 0, obj.ScaleWidth, 25, 147, 40, 10, 10)
        Case "waitback": pos = Array(0, 0, obj.ScaleWidth, obj.ScaleHeight, 147, 40, 10, 10)
        Case "waitmidback": pos = Array(0, 140, obj.ScaleWidth, obj.ScaleHeight - 180, 147, 50, 10, 10)
        Case "botback": pos = Array(0, 0, obj.ScaleWidth, IIf(obj.ScaleHeight < 380, obj.ScaleHeight + 25, 400), 147, 50, 9, 9)
        Case "bot2back": pos = Array(0, 400, obj.ScaleWidth, obj.ScaleHeight - 400, 147, 50, 10, 10)
        Case "pinkback": pos = Array(0, 13, obj.ScaleWidth, obj.ScaleHeight, 157, 40, 10, 10)
        Case "blueback": pos = Array(0, 0, obj.ScaleWidth, obj.ScaleHeight, 157, 50, 10, 10)
        Case "shipback": pos = Array(0, arg * 44, 58, 44, 157, 50, 10, 10)
    End Select
    If IsEmpty(pos) = False Then
        If lastobj <> obj.Parent.Name & obj.Name & obj.TabIndex Then
            lastobj = obj.Parent.Name & obj.Name & obj.TabIndex
            SelectObject myBackBuffer, myBufferBMP: SelectObject myBackBuffer, obj.Picture
        End If
        If Right(img, 4) = "back" Then GoTo 10
        BitBlt myBackBuffer, pos(0), pos(1), pos(6), pos(7), mySprite, pos(4), pos(5), vbSrcCopy
        BitBlt obj.hdc, 0, 0, obj.ScaleWidth, obj.ScaleHeight, myBackBuffer, 0, 0, vbSrcCopy: obj.Picture = obj.Image
        Exit Function
10      i = 0: While i < pos(2)
            j = 0: While j < pos(3)
                BitBlt myBackBuffer, pos(0) + i, pos(1) + j, IIf(pos(2) - i < pos(6), pos(2) - i, pos(6)), IIf(pos(3) - j < pos(7), pos(3) - j, pos(7)), mySprite, pos(4), pos(5), vbSrcCopy
            j = j + pos(7): Wend
        i = i + pos(6): Wend
        BitBlt obj.hdc, 0, 0, obj.ScaleWidth, obj.ScaleHeight, myBackBuffer, 0, 0, vbSrcCopy: obj.Picture = obj.Image
    End If
End Function

Public Function deletebuffer()
    'clears memory
    DeleteObject myBufferBMP: DeleteDC myBackBuffer: DeleteDC mySprite
    For i = 0 To 3
        DeleteObject myMapBufferBMP(i): DeleteDC myMapBackBuffer(i)
    Next i
End Function

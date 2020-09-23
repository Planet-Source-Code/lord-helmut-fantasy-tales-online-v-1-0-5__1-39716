Attribute VB_Name = "GameFunctions"
'this is where the game functions will be stored such as the sub to create the
'levels from the files, the subs to draw the players, text, menus, etc
Option Explicit
Private Keys(1 To 5) As Boolean '1 - 4 are dirs, 5 is escape

Public Sub MainLoop()

    Do
        DoEvents
        
        BodyRECT.left = ((frame * PlayerData.BodyWidth) - PlayerData.BodyWidth)
        BodyRECT.Right = (frame * PlayerData.BodyWidth)
        HeadRECT.top = ((frame * PlayerData.HeadHeight) - PlayerData.HeadHeight)
        HeadRECT.Bottom = (frame * PlayerData.HeadHeight)
        
        'Call GameFunctions.DrawLevel(TileSet, ddsBackBuffer)
        Call DrawLevel(TileSet, ddsBackBuffer)
        Call DrawPlayerData
        Call DrawText
        Call DrawFPS
        
        Call CheckKeys
        Call ExecKeys

        ddsPrimary.Flip Nothing, DDFLIP_WAIT
        
    Loop Until Running = False
    
End Sub

Public Sub LoadData() 'Loads the game level into the game
  Dim x As Integer, y As Integer 'the x and y coords of each tile

  Open App.Path & "\Files\start.fto" For Binary As #1 'opens the level up for inputting the info
  For x = 0 To 63
    For y = 0 To 63
      Get #1, , Map(x, y) 'gets the coords of the tile
    Next y 'loops
  Next x 'loops
Close 'closes the level file

End Sub

Public Sub LoadPlayerStats(Nick As String, Class As String, PlayerX As Single, PlayerY As Single)
Open App.Path & "\Files\PlayerStats.dat" For Input As #1
Input #1, Nick 'loads the players nickname in from a textfile
Input #1, Class 'loads in the players class
Input #1, PlayerX, PlayerY 'loads in the players coordinates
Close #1
End Sub

Public Sub ExecKeys()

If Keys(5) = True Then
    Running = False
End If
If Keys(1) = True Then
    PlayerData.y = PlayerData.y + (PlayerSpeed * Elapsed * 16)
    Call PlaySound(Steps)
    frame = 1
End If
If Keys(2) = True Then
    PlayerData.x = PlayerData.x - (PlayerSpeed * Elapsed * 16)
    Call PlaySound(Steps)
    frame = 2
End If
If Keys(3) = True Then
    PlayerData.y = PlayerData.y - (PlayerSpeed * Elapsed * 16)
    Call PlaySound(Steps)
    frame = 3
End If
If Keys(4) = True Then
    PlayerData.x = PlayerData.x + (PlayerSpeed * Elapsed * 16)
    Call PlaySound(Steps)
    frame = 4
End If

End Sub

Public Sub CheckKeys()
DIDEV.GetDeviceStateKeyboard DIState

If DIState.Key(DIK_ESCAPE) <> 0 Then
    Keys(5) = True
Else
    Keys(5) = False
End If
If DIState.Key(DIK_DOWN) <> 0 Then
    Keys(1) = True
Else
    Keys(1) = False
End If
If DIState.Key(DIK_LEFT) <> 0 Then
    Keys(2) = True
Else
    Keys(2) = False
End If
If DIState.Key(DIK_UP) <> 0 Then
    Keys(3) = True
Else
    Keys(3) = False
End If
If DIState.Key(DIK_RIGHT) <> 0 Then
    Keys(4) = True
Else
    Keys(4) = False
End If

End Sub

Public Sub DrawPlayerData()
        'line to display character
        
        BltPlayer Body, BodyRECT, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT          'puts the body image at the players x and y coords
        BltPlayer Head, HeadRECT, DDBLTFAST_SRCCOLORKEY Or DDBLTFAST_WAIT, 3, -21        'places the head a little about the x and y coords
End Sub

Public Sub DrawText()
Dim DrawSize As Size 'This is the UDT to hold the pixel size of the text string
Dim x As Single, y As Single
    Call GetPlayerPosReMap(x, y)
        'line to print text out
        GetTextExtentPoint32 Main.hdc, Trim$(PlayerStats.Nick), Len(Trim$(PlayerStats.Nick)), DrawSize 'Loads the length of the string into the UDT called Size
        Call ddsBackBuffer.DrawText(Fix((x + (PlayerData.BodyWidth / 2)) - (DrawSize.cx / 2)), y + 36, Trim$(PlayerStats.Nick), False) 'Tells where to draw the nickname string
End Sub

Public Sub PlaySound(Sound As DirectSoundBuffer)
Sound.Play DSBPLAY_DEFAULT
End Sub

Public Sub DrawFPS()

    Static ElapsedTimer As Long
    Static Timer As Long
    Static FPSCounter As Integer
    Static FPS As Integer

    Elapsed = Dx.TickCount - ElapsedTimer
    ElapsedTimer = Dx.TickCount
    If Timer + 1000 <= Dx.TickCount Then
        Timer = Dx.TickCount
        FPS = FPSCounter + 1
        FPSCounter = 0
    Else
        FPSCounter = FPSCounter + 1
    End If

    ddsBackBuffer.SetForeColor vbWhite
    ddsBackBuffer.SetFontBackColor vbBlue
    ddsBackBuffer.SetFontTransparency False

    ddsBackBuffer.DrawText 1, 440, "FPS: " & FPS, False

    ddsBackBuffer.SetFontTransparency True
    ddsBackBuffer.SetForeColor RGB(0, 0, 0)

End Sub

Public Sub SyncFont()
    Main.Font.Name = fontinfo.Name
    Main.Font.Bold = fontinfo.Bold
    Main.Font.Size = fontinfo.Size
End Sub

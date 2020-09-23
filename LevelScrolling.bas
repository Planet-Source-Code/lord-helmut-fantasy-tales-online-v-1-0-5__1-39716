Attribute VB_Name = "LevelScrolling"
Public Sub DrawLevel(ByRef TileSourceSurf As DirectDrawSurface7, ByRef TileDestSurf As DirectDrawSurface7)

    'RhaEngine scrolling level code
    'Copyright (c) 2001-2002 Boco Soft.
    'Please don't claim credit for my code.
    'Programmed by Boco

    Dim SrcRect As RECT
    Dim I As Integer
    Dim J As Integer
    Dim DrawLoc(0 To 1) As Integer
    Dim StartTile(0 To 3) As Integer
    Dim TCount(0 To 1) As Integer
    Dim OffSet(0 To 1) As Integer
    Dim tile As Integer

    'This specifies the "window" the player looks through.  By multiplying
    'the player tile coordinate by the width of the tiles, you get the position
    'of the player in pixels.
    'NOTE: Since FTO Character coordinates are alread in pixels, I deleted the
    '  multiplication by the tile width and height.
    

    viewWindow.left = PlayerData.x - (640 / 2)
    viewWindow.Right = PlayerData.x + (640 / 2)
    viewWindow.top = PlayerData.y - (480 / 2)
    viewWindow.Bottom = PlayerData.y + (480 / 2)
    
    'Now, we see if the "window" is viewing outside of the level.  If it is,
    'reposition it (has the affect of the player staying in the center of the
    'screen until you come to the edge of the map.

    If viewWindow.left < 0 Then
        viewWindow.left = 0: viewWindow.Right = 640
    End If

    'If the map is smaller than the screen, then the SmallMap variable is set
    'to True.  It then appropriatly adjusts the player viewing "window".

    If viewWindow.Right > (UBound(Map, 1) * 16) + 16 Then
        viewWindow.Right = (UBound(Map, 1) * 16) + 16
        viewWindow.left = viewWindow.Right - 640
        If viewWindow.left < 0 Then
            SmallMap = True
            viewWindow.left = 0: viewWindow.Right = UBound(Map, 1) * 16
        End If
    End If
    If viewWindow.top < 0 Then
        viewWindow.top = 0: viewWindow.Bottom = 480
    End If
    If viewWindow.Bottom > (UBound(Map, 2) * 16) + 16 Then
        viewWindow.Bottom = (UBound(Map, 2) * 16) + 16
        viewWindow.top = viewWindow.Bottom - 480
        If viewWindow.top < 0 Then
            SmallMap = True
            viewWindow.top = 0: viewWindow.Bottom = UBound(Map, 2) * 16
        End If
    End If

    'Determine the tile offset.  When you "smooth scroll" a map, the tiles don't
    'move one tile at a time.  You can see maybe 25% of the tile.  The offset
    'determines how much of the tile we can see.

    OffSet(0) = viewWindow.left Mod 16
    OffSet(1) = viewWindow.top Mod 16

    'Now, we find which tiles we are going to display.  We use our "window" and
    'convert the value back into tiles.

    StartTile(0) = Fix(viewWindow.left / 16)
    StartTile(1) = Fix(viewWindow.Right / 16)
    StartTile(2) = Fix(viewWindow.top / 16)
    StartTile(3) = Fix(viewWindow.Bottom / 16)

    'This fixes my problem!  What this will do is show the extreme right and bottom
    'edge tiles.  They wouldn't show before.

    If StartTile(1) = (UBound(Map, 1) + 1) Then StartTile(1) = StartTile(1) - 1
    If StartTile(3) = (UBound(Map, 2) + 1) Then StartTile(3) = StartTile(3) - 1

    'Best to set it to zero.

    TCount(0) = 0
    TCount(1) = 0

    For I = StartTile(2) To StartTile(3)
        For J = StartTile(0) To StartTile(1)

            'This determines which tile we are going to use out of our tileset.  We
            'will pass SrcRect to the BltFast DirectX command to specify the window on
            'tileset that we will copy to the screen.

            tile = Map(J, I)
            With SrcRect
                .left = Fix((tile \ 63) * 16)
                .top = Fix((tile Mod 63) * 16)
                .Right = .left + 16
                .Bottom = .top + 16
            End With

            'If a portion of the tile is not shown, then adjust the source window.
            'If you try to blit outside of the screen in DirectX, it will not blit.

            If J * 16 < viewWindow.left Then SrcRect.left = SrcRect.left + (viewWindow.left Mod 16)
            If I * 16 < viewWindow.top Then SrcRect.top = SrcRect.top + (viewWindow.top Mod 16)
            If J * 16 + 16 > viewWindow.Right Then SrcRect.Right = SrcRect.Right - (J * 16 + 16 - viewWindow.Right)
            If I * 16 + 16 > viewWindow.Bottom Then SrcRect.Bottom = SrcRect.Bottom - (I * 16 + 16 - viewWindow.Bottom)

            'Get the location on the screen where we are going to draw to.

            DrawLoc(0) = TCount(0) * 16 - OffSet(0)
            DrawLoc(1) = TCount(1) * 16 - OffSet(1)
            If DrawLoc(0) < 0 Then DrawLoc(0) = 0
            If DrawLoc(1) < 0 Then DrawLoc(1) = 0

            'After a little check, blit the tile!

            If I < StartTile(3) Or J < StartTile(1) Or SmallMap = False Then
                TileDestSurf.BltFast DrawLoc(0), DrawLoc(1), TileSourceSurf, SrcRect, DDBLTFAST_WAIT
            End If

            'Increment our tile count

            TCount(0) = TCount(0) + 1
        Next J

        'Increment our tile count again

        TCount(0) = 0
        TCount(1) = TCount(1) + 1
    Next I

End Sub

Public Sub BltPlayer(ByRef surf As DirectDrawSurface7, ByRef SrcRect As RECT, _
        ByVal trans As CONST_DDBLTFASTFLAGS, Optional xOffset As Single, Optional yOffset As Single)
        Dim x As Single, y As Single
        Dim temprect As RECT
        
        Call GetPlayerPosReMap(x, y)
        x = x + xOffset: y = y + yOffset
        temprect = SrcRect 'because we can't pass rects byval _
        we need to store it to reset it later on
        
        'Here we check if the persons too far left and adjust
        If x <= 0 Then SrcRect.left = SrcRect.left - x
        If (x + SrcRect.Right - SrcRect.left) >= 640 Then _
        SrcRect.Right = SrcRect.left + (640 - x)
        x = IIf(x < 0, 0, x)
      
        'Here we check if the persons too far right and adjust
        If y <= 0 Then SrcRect.top = SrcRect.top - y
        If (y + SrcRect.Bottom - SrcRect.top) >= 480 Then _
        SrcRect.Bottom = SrcRect.top + (480 - y)
        y = IIf(y < 0, 0, y)
        
        'Blt the whole thing
        ddsBackBuffer.BltFast x, y, surf, SrcRect, trans
        
        'Then copy rect rect back over
        SrcRect = temprect
End Sub

Public Sub GetPlayerPosReMap(ByRef x As Single, y As Single)
        x = PlayerData.x: y = PlayerData.y
        
        'If the map stopped scrolling, move the player
        If Int(x / 16) > 42 Then
            x = x - (22 * 16)
        'If the maps still scrolling, don't move the player
        ElseIf Int(x / 16) > 20 And Int(x / 16) < 43 Then
            x = (21 * 16)
        End If
        
        'If the map stopped scrolling, move the player
        If Int(y / 16) > 47 Then
            y = y - (32 * 16)
        'If the maps still scrolling, don't move the player
        ElseIf Int(y / 16) > 15 And Int(y / 16) < 48 Then
            y = (16 * 16)
        End If
End Sub

Public Sub GetPlayerMap(ByRef left As Double, ByRef top As Double)
   'If the player has moved past the middle of the map
   If (PlayerData.x / 16) > 20 Then
        'Check that they're not too close to the edge
        If (PlayerData.x / 16) > 43 Then
            'If they are, stop the map scrolling
            left = 23
        Else
            'Otherwise figure out the correct scrolling distance
            left = (PlayerData.x / 16) - 20
        End If
    End If
    
       'If the player has moved past the middle of the map
    If (PlayerData.y / 16) > 15 Then
        'Check that they're not too close to the edge
        If (PlayerData.y / 16) > 48 Then
            'If they are, stop the map scrolling
            top = 33
        Else
            'Otherwise figure out the correct scrolling distance
            top = (PlayerData.y / 16) - 15
        End If
    End If
    
End Sub


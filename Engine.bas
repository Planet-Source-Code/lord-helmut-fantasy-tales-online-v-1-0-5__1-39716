Attribute VB_Name = "Engine"
Option Explicit 'Makes sure every variable is declared

'DX Part Declarations
'-------------------------
Public Dx As New DirectX7 'Declares DX
Public DDraw As DirectDraw7 'Declares DD
Public DSound As DirectSound 'Declares DS
Public DInput As DirectInput 'Declares DI

'Main Surfaces
'-------------------------
Public ddsPrimary As DirectDrawSurface7 'Primary surface everything is shown on
Public ddsBackBuffer As DirectDrawSurface7 'Back surface everything is blitted to first

'Sound Buffers
'-------------------------
Public Steps As DirectSoundBuffer 'Buffer that holds walking sound
Public StepsBufferDesc As DSBUFFERDESC
Public Sword As DirectSoundBuffer 'buffer that holds sword sound
Public SwordBufferDesc As DSBUFFERDESC

'Other surfaces
'-------------------------
Public Level As DirectDrawSurface7 'surface that holds the level
Public LevelRECT As RECT 'size of the level
Public Body As DirectDrawSurface7 'surface that holds the body sprite
Public BodyRECT As RECT 'size of sprite sheet for body
Public Head As DirectDrawSurface7 'surface that holds the head sprite
Public HeadRECT As RECT 'size of the sprite sheet for head
Public TileSet As DirectDrawSurface7 'Surface that holds tileset
Public TileSetRECT As RECT 'size of the tileset surface

'Surface Descriptions
'-------------------------
Public ddsdPrimary As DDSURFACEDESC2
Public ddsdBackbuffer As DDSURFACEDESC2

'Stuff for Direct Sound
'-------------------------
Public DSDesc As DSBUFFERDESC 'describes settings for the buffer
Public DSWave As WAVEFORMATEX 'describes settings for the sound buffer

'Stuff For Direct Input
'-------------------------
Public DIDEV As DirectInputDevice
Public DIState As DIKEYBOARDSTATE

Public fontinfo As New StdFont 'declares the font data

'Public Types
'-------------------------
Public Type PlayerData 'These are variables that only the game will be able to edit
    x As Single 'The players X coordinate
    y As Single 'The Players Y coordinate
    BodyHeight As Integer 'The pixel height of the body
    BodyWidth As Integer 'The pixel width of the body
    HeadHeight As Integer 'The pixel height of the head
    HeadWidth As Integer 'The pixel width of the head
End Type 'ends the playerdata udt
Public Type PlayerStats 'Stuff that the player will be able to edit in later versions
    Nick As String * 15 'The players nickname, this can be a max of 15 chars
    Class As String 'This is the name of the players class
End Type 'Ends the UDT
Type Size
    cx As Long
    cy As Long
End Type

'API Declerations
'-------------------------
Declare Function GetTextExtentPoint32 Lib "gdi32" Alias "GetTextExtentPoint32A" (ByVal hdc As Long, ByVal lpsz As String, ByVal cbString As Long, lpSize As Size) As Long
Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long

'General Declerations
'-------------------------
Public frame As Integer 'what frame is current displayed
Public Running As Boolean  'if program is running
Public Map(63, 63) As Integer
Public xOffset As Integer, yOffset As Integer
Public PlayerData As PlayerData, PlayerStats As PlayerStats
Public Elapsed As Long

Public viewWindow As RECT
Public SmallMap As Boolean

'Generalized Constants
'-------------------------
Public Const ScreenWidth = 640 'This is the width of the screen
Public Const screenHeight = 480 'This is the height of the screen
Public Const TileWidth = 16 'size of the tiles width
Public Const TileHeight = 16 'Size of the tiles height

Public Const PlayerSpeed = 0.005 'how fast player runs

Public Sub Terminate() 'Unsets everything and closes

    Set Level = Nothing 'Takes everything out of the level surface
    Set Body = Nothing 'Erases all the body data
    Set Head = Nothing 'Erases all the head data
    Set TileSet = Nothing 'Erases the tileset surface
    
    Set ddsPrimary = Nothing 'Clears off the primary surface
    Set ddsBackBuffer = Nothing 'Cleans out the backbuffer
    
    DDraw.RestoreDisplayMode 'Gives player back there display mode
    DDraw.SetCooperativeLevel Main.hWnd, DDSCL_NORMAL 'Sets everything back to normal
    
    DIDEV.Unacquire 'Gets rid of the stuff needed for DirectInput
    ShowCursor 1 'Shows the players mouse cursor again
    End 'ends the program
    
End Sub

Public Sub InitDirectDraw() 'All the initilization for Direct Draw
Set DDraw = Dx.DirectDrawCreate("")
    Main.Show

    DDraw.SetCooperativeLevel Main.hWnd, DDSCL_FULLSCREEN Or DDSCL_ALLOWMODEX Or DDSCL_EXCLUSIVE
    DDraw.SetDisplayMode ScreenWidth, screenHeight, 16, 0, DDSDM_DEFAULT
    

    ddsdPrimary.lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
    ddsdPrimary.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_FLIP Or DDSCAPS_COMPLEX
    ddsdPrimary.lBackBufferCount = 1
    
    Set ddsPrimary = DDraw.CreateSurface(ddsdPrimary)
    
    Dim Caps As DDSCAPS2
    Caps.lCaps = DDSCAPS_BACKBUFFER
    
    Set ddsBackBuffer = ddsPrimary.GetAttachedSurface(Caps)

    ddsBackBuffer.GetSurfaceDesc ddsdBackbuffer
End Sub

Public Sub DDCreateSurface(surface As DirectDrawSurface7, BmpPath As String, RECTvar As RECT, Optional TransCol As Integer = 0, Optional UseSystemMemory As Boolean = True)
'This sub is called when it needs to create a surface
    Dim tempddsd As DDSURFACEDESC2
    
    Set surface = Nothing
    
    tempddsd.lFlags = DDSD_CAPS
    If UseSystemMemory = True Then
        tempddsd.ddsCaps.lCaps = DDSCAPS_SYSTEMMEMORY Or DDSCAPS_OFFSCREENPLAIN
    Else
        tempddsd.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
    End If
    Set surface = DDraw.CreateSurfaceFromFile(BmpPath, tempddsd)
    
    RECTvar.Right = tempddsd.lWidth
    RECTvar.Bottom = tempddsd.lHeight
    
    Dim ddckColourKey As DDCOLORKEY
    ddckColourKey.low = 0
    ddckColourKey.high = 0
    surface.SetColorKey DDCKEY_SRCBLT, ddckColourKey
    
End Sub

Public Sub InitDirectSound()
Set DSound = Dx.DirectSoundCreate("")
If Err.Number <> 0 Then
    Exit Sub
End If
DSound.SetCooperativeLevel Main.hWnd, DSSCL_EXCLUSIVE
DSDesc.lFlags = DSBCAPS_CTRLFREQUENCY Or DSBCAPS_CTRLPAN Or DSBCAPS_CTRLVOLUME Or DSBCAPS_STATIC
DSWave.nFormatTag = WAVE_FORMAT_PCM 'Sound Must be PCM otherwise we get errors
DSWave.nChannels = 2    '1= Mono, 2 = Stereo
DSWave.lSamplesPerSec = 22050
DSWave.nBitsPerSample = 16 '16 =16bit, 8=8bit
DSWave.nBlockAlign = DSWave.nBitsPerSample / 8 * DSWave.nChannels
DSWave.lAvgBytesPerSec = DSWave.lSamplesPerSec * DSWave.nBlockAlign

Call DSCreateSoundBuffer(Steps, App.Path & "\Sounds\steps.wav", DSDesc, DSWave)
Call DSCreateSoundBuffer(Sword, App.Path & "\Sounds\sword.wav", DSDesc, DSWave)

End Sub

Public Sub DSCreateSoundBuffer(Buffer As DirectSoundBuffer, filename As String, BufferDesc As DSBUFFERDESC, wFormat As WAVEFORMATEX)
If DSound Is Nothing Then Exit Sub
    Set Buffer = DSound.CreateSoundBufferFromFile(filename, BufferDesc, wFormat)
End Sub

Public Sub InitTextDisplay()
    ddsBackBuffer.SetFontTransparency True
    ddsBackBuffer.SetForeColor RGB(0, 0, 0)
    fontinfo.Bold = True
    fontinfo.Size = 10
    fontinfo.Name = "verdana"
    ddsBackBuffer.SetFont fontinfo
    
    Call SyncFont
End Sub

Public Sub InitDirectInput()
    Set DInput = Dx.DirectInputCreate()
    If Err.Number <> 0 Then
        Exit Sub
    End If
    Set DIDEV = DInput.CreateDevice("GUID_SysKeyboard")
    DIDEV.SetCommonDataFormat DIFORMAT_KEYBOARD
    DIDEV.SetCooperativeLevel Main.hWnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE
    DIDEV.Acquire
End Sub



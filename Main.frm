VERSION 5.00
Begin VB.Form Main 
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
On Local Error GoTo errhandler

'starting frame
frame = 1
'PlayerData Sprite Sizes
PlayerData.BodyHeight = 43 'how tall the body sprite is
PlayerData.BodyWidth = 38 'how wide the body sprite is
PlayerData.HeadHeight = 32 'how tall the head sprite is
PlayerData.HeadWidth = 32 'how wide the head sprite is

    Running = True 'sets the boolean to true so we know the games running
On Local Error GoTo errhandler 'if an error occurs, go to the error handler

    Engine.InitDirectDraw 'Initilizes the Direct Draw part of DX
    Engine.InitDirectSound 'Initilizes the Direct Sound part of DX
    Engine.InitTextDisplay 'Initilizes the part to show text
    Engine.InitDirectInput 'Initilizes the Direct Input portion
    
    'LoadPlayerStats loads the players nickname, there class, there x and y coords
    Call LoadPlayerStats(PlayerStats.Nick, PlayerStats.Class, PlayerData.x, PlayerData.y)

    DDCreateSurface TileSet, App.Path & "\Files\TileSet.til", TileSetRECT 'creates a surface to hold tileset image
    DDCreateSurface Head, App.Path & "\Characters\" & PlayerStats.Class & "-Head.gfx", HeadRECT 'creates surface to hold the players head
    DDCreateSurface Body, App.Path & "\Characters\" & PlayerStats.Class & "-Body.gfx", BodyRECT 'creates surface to hold the players body

    Call LoadData 'loads in level

'variables values for the PlayerData image
BodyRECT.left = 0
BodyRECT.Right = 38
BodyRECT.top = 0
BodyRECT.Bottom = 43

HeadRECT.left = 0
HeadRECT.Right = 32
HeadRECT.top = 0
HeadRECT.Bottom = 32

ShowCursor 0 'hides the mouse cursor when game starts
MainLoop 'start the games loop
Terminate 'Calls the terminate sub after main loop ends

errhandler: 'If an error occurs, it goes to this
Running = False 'Stops the main loop
Terminate 'Calls the Terminate sub

End Sub


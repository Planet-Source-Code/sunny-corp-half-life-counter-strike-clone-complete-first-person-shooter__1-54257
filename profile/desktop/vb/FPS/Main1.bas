Attribute VB_Name = "ModMain"
'---------Credits----------'
'Cyborg mesh from Fabio
'star wars meshes from JK2 but entirely animated by me, AA, and reskined to suit DX
'Collison, and world subs from Fernando

Option Explicit
Dim ob As Boolean
'main DX variables
Global DX_Main As New DirectX7
Global D3D_Main As Direct3DRM3
Global DD_Main As DirectDraw4
Dim i 'normal counter used in various loops
Dim bcol1
Global move As Boolean
'DirectInput variables
Dim DI_Main As DirectInput 'main 'parent' of direct input
Dim DI_Device As DirectInputDevice 'will hold state info about keyboard
Dim DI_Device2 As DirectInputDevice 'will hold state info about mouse
Dim DI_State As DIKEYBOARDSTATE 'gets current state of keyboard, i.e.: what keys r pressed
Dim DI_State2 As DIMOUSESTATE 'gets current state of mouse, i.e.:what mouse button is pressed
Global posx, posy, posz, ptheta 'useless for now, DELETE THEM!!!!!
Dim obm
Dim boro 'BOT rotation, before
Dim curro 'bot rotation, now

'angles 5 deg.
Const Cos5 = 0.996194698091746
Const Sin5 = 8.71557427476582E-02

'DirectDraw surfaces
Dim DS_Front As DirectDrawSurface4 'front buffer , what you see
Global DS_Back As DirectDrawSurface4 'back buffer, what everything is drawn on
Dim SD_Front As DDSURFACEDESC2  'Surface Description
Dim DD_Back As DDSCAPS2  'General Surface Info

'Direct3d Devices
Dim D3D_Device As Direct3DRMDevice3 'will hold and acquire all D3D related events
Dim D3D_Viewport As Direct3DRMViewport2 'what u see, in the world sense
Dim D3D_View2 As Direct3DRMViewport2

'Declare frames(places to put the meshes)
Dim FR_Root As Direct3DRMFrame3 'our main frame
Dim FR_Camera As Direct3DRMFrame3 'our camera
Dim FR_Light As Direct3DRMFrame3 'our light
Dim FR_Gun As Direct3DRMFrame3 'our gun
Dim FR_Objects As Direct3DRMFrame3 ' Normal game objects
Dim FR_Ammo(4) As Direct3DRMFrame3 '5 free clips at one time
Dim FR_Bullet(30) As Direct3DRMFrame3 'will have 31 bullets per time avialable, holds the
Dim CurFrameG As Single 'current frame of the gun loading

'Enemies: Cyborg
Global FR_Mech1 As Direct3DRMFrame3
Dim AN_Mech1 As Direct3DRMAnimationSet2
Dim AN_Gun As Direct3DRMAnimationSet2
Dim CurFrameM1 As Single

Global Health As Integer 'our health



'Hostage
Dim AN_Host As Direct3DRMAnimationSet2
Dim FR_Host As Direct3DRMFrame3
Dim CurFrameHo





'yoda!
Dim AN_Yoda As Direct3DRMAnimationSet2
Global FR_Yoda As Direct3DRMFrame3
Dim CurFrameY 'Yoda's frame count
Dim ypos As D3DVECTOR

Dim AN_Battle As Direct3DRMAnimationSet2
Dim FR_Battle As Direct3DRMFrame3
Dim BaPos As D3DVECTOR
Dim CurFrameB

'various states of different people/objects/etc.
Private Enum Enemy_StateEnum 'enemy states, self explanatory
    Normal
    Flee
    Look
    die
    Fire
End Enum

Private Enum Player_StateEnum 'player states
    dead
    Player_Normal
    AlmostDead
    Firep
    Reload
End Enum

Public Enum Yoda
    NormalY
    Talk
    Hand
    Saber
End Enum

'Global dead As Boolean
Dim AI_Timer As Integer 'timer to start looking for the player unless he/she is not found

'enemy bullets
Dim ActiveB(50) As Boolean 'checks if the selected bullet has been shot
Dim BB 'used to pass the value on the number of current bullet and so forth
Dim m_Frame(50) As Direct3DRMFrame3 'enemy will have an ammo of 51 bullets, recharged automatically
Dim m_meshBuilder(50) As Direct3DRMMeshBuilder3 'bullets r created by the mesh builder and the meshes are loaded in


Global FR_World As Direct3DRMFrame3 'world objects
Global Pos As D3DVECTOR 'general obj. positions/cam. pos.
Global meyes As D3DVECTOR 'eyes of player
Dim heyes 'height of player; might add crouch, and etc.


Global GPos As D3DVECTOR 'pos. of gun
'Meshes
Dim MS_World As Direct3DRMMeshBuilder3
Dim MS_Objects As Direct3DRMMeshBuilder3
Dim MS_Ammo As Direct3DRMMeshBuilder3
Dim MS_Gun As Direct3DRMMeshBuilder3
Dim T_Sky As Direct3DRMTexture3 'texture of the gun;revise
Dim MS_Bullet(30) As Direct3DRMMeshBuilder3
Dim ActiveM(30) As Boolean 'state of each bullet being fired
Global Clips As Integer 'clips for the gun
Global Bullets As Integer 'bullets for the gun
Global Msg As String 'gun messsages
Global BPos As D3DVECTOR 'holds the position of the battle scence
Dim BCol 'checks bullet collision
Dim Mech1_Health 'health of the mech
Dim BM 'counter for shooting loops, BM: Bullet Mine

'Lights
Dim LT_Ambient As Direct3DRMLight 'will light the entire world
Dim LT_Spotlight As Direct3DRMLight ' will be a spot light
Dim LT_GPoint As Direct3DRMLight 'gun light

'camera positions
Dim XX, YY, ZZ As Long

Dim GDistance, GAngle 'variables to store data obtained from functions requesting distance between, and angle

'Music & sound related variables; Only for DirectMusic, DirectSound is in ModSound!
Dim Pref As DirectMusicPerformance 'Performance; quality, better wording could be ur preferences
Dim Seg As DirectMusicSegment 'what the loader(see below) puts in here 'segment' of a music file
Dim SegState As DirectMusicSegmentState 'what is current state of the above segment; has the file finsihed, or is it still playing
Dim Loader1 As DirectMusicLoader 'loads the first file

'use windows mouse api
Public Declare Function ShowCursor Lib "user32" (ByVal bShow As Long) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
'enumerate
Private Enemy_State As Enemy_StateEnum 'declare the enemy state to a type, declared above
Private Player_State As Player_StateEnum ' declare player state to a type, declared above
Private Yoda_State As Yoda
Public Type POINTAPI 'track mouse position
    x As Long
    y As Long
End Type

Global MPos As POINTAPI 'mouse postion( MousePosition) is declared to the type point api as above
Global esc As Boolean 'global var to check if the user quit
Dim m As String 'normal variable; revise
Global CTime As Integer

Public Sub DX_Init()
ShowCursor False 'hide mouse
heyes = 2.8
Set DD_Main = DX_Main.DirectDraw4Create("")

DD_Main.SetCooperativeLevel frmMain.hWnd, DDSCL_FULLSCREEN Or DDSCL_EXCLUSIVE

DD_Main.SetDisplayMode 640, 480, 16, 0, DDSDM_DEFAULT

SD_Front.lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
SD_Front.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_3DDEVICE Or DDSCAPS_FLIP Or DDSCAPS_COMPLEX
SD_Front.lBackBufferCount = 1
Set DS_Front = DD_Main.CreateSurface(SD_Front)
'previous set created the screen

DD_Back.lCaps = DDSCAPS_BACKBUFFER
Set DS_Back = DS_Front.GetAttachedSurface(DD_Back)
DS_Back.SetForeColor RGB(255, 0, 0)
'previous set the area to create text

'Set up direct3d and devices
Set D3D_Main = DX_Main.Direct3DRMCreate() 'set up d3d in rm
Set D3D_Device = D3D_Main.CreateDeviceFromSurface("IID_IDirect3DHalDevice", DD_Main, DS_Back, D3DRMDEVICE_DEFAULT) 'hardware rendering

D3D_Device.SetBufferCount 2
D3D_Device.SetQuality D3DRMRENDER_GOURAUD  'render quality
D3D_Device.SetTextureQuality D3DRMTEXTURE_LINEAR
D3D_Device.SetRenderMode D3DRMRENDERMODE_BLENDEDTRANSPARENCY

'start initializing directinput
'set up keyboard
Set DI_Main = DX_Main.DirectInputCreate
Set DI_Device = DI_Main.CreateDevice("Guid_SYSKeyboard")
DI_Device.SetCommonDataFormat DIFORMAT_KEYBOARD
DI_Device.SetCooperativeLevel frmMain.hWnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE
DI_Device.Acquire 'get the info
'set up mouse, only for mouse click
Set DI_Device2 = DI_Main.CreateDevice("Guid_SYSMouse")
DI_Device2.SetCommonDataFormat DIFORMAT_MOUSE
DI_Device2.SetCooperativeLevel frmMain.hWnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE
DI_Device2.Acquire


End Sub

Public Sub DX_MusicInit()

  On Local Error GoTo MusicInit
  
  'sets the loader to a position to load .mid files, in this case main music,background file
  Set Loader1 = DX_Main.DirectMusicLoaderCreate()
  
  'set the preferance to our main dx7 object
  Set Pref = DX_Main.DirectMusicPerformanceCreate
  
  'initialize preferences
  Call Pref.Init(Nothing, 0)
  
  'set standard options
  Pref.SetPort -1, 80
  Call Pref.SetMasterAutoDownload(True)
  'set our volume and tempo
  Call SetVolume(40)
  Call SetTempo(1)
  
  Set Seg = Loader1.LoadSegment(App.Path & "/Sounds/Battle.mid")  'load our main background music
  
  Seg.SetStandardMidiFile 'set file format to standard file type
  
  Set SegState = Pref.PlaySegment(Seg, 0, 0) '0 means start from start
  
MusicInit:
  frmMain.ProgressBar1.Value = 50
  frmMain.Label2.Caption = "Loading Music Files... Setting Up Files..."
End Sub

Public Sub SetTempo(nTempo As Single)
'the tempo must be between 0.25 and 2
'the default is 1; 2 would double this, 0.5 would half this
Pref.SetMasterTempo nTempo
End Sub

Public Sub SetVolume(nVolume As Byte)
'This formula allows you to specify a volume between
'0-100; similiar to a percentage
Pref.SetMasterVolume (nVolume * 42 - 3000)
End Sub

Public Sub DX_MakeObjects()
'create frames
Set FR_Root = D3D_Main.CreateFrame(Nothing)
Set FR_Camera = D3D_Main.CreateFrame(FR_Root)
Set FR_Light = D3D_Main.CreateFrame(FR_Root)
Set FR_World = D3D_Main.CreateFrame(FR_Root)
Set FR_Mech1 = D3D_Main.CreateFrame(FR_Root)
Set FR_Gun = D3D_Main.CreateFrame(FR_Root)
Set FR_Objects = D3D_Main.CreateFrame(FR_Root)
Set FR_Yoda = D3D_Main.CreateFrame(FR_Root)
Set FR_Battle = D3D_Main.CreateFrame(FR_Root)
Set FR_Host = D3D_Main.CreateFrame(FR_Root)

For i = 0 To 4
    Set FR_Ammo(i) = D3D_Main.CreateFrame(FR_Root)
Next i

For i = 0 To 30
Set FR_Bullet(i) = D3D_Main.CreateFrame(FR_Root)
Next i

For i = 0 To 50
       Set m_Frame(i) = D3D_Main.CreateFrame(FR_Root)
Next i

CTime = 0 'time line set to 0 the start

FR_Root.SetSceneBackgroundRGB 0, 0, 0 'back ground color to black

FR_Camera.SetPosition Nothing, -130, 6, -130 'start position of player

Set D3D_Viewport = D3D_Main.CreateViewport(D3D_Device, FR_Camera, 0, 0, 640, 480)
Health = 1000 'player health
D3D_Viewport.SetBack 1000 'the distance you can see from where ur standing

FR_Light.SetPosition Nothing, 0, 25, 0 'spotlight position
Set LT_Spotlight = D3D_Main.CreateLightRGB(D3DRMLIGHT_POINT, 1, 1, 1)
FR_Light.AddLight LT_Spotlight

Set LT_Ambient = D3D_Main.CreateLightRGB(D3DRMLIGHT_AMBIENT, 0.5, 0.5, 0.5) 'ambient light, e.g. the main light, it lights the entire world
FR_Root.AddLight LT_Ambient

FR_Camera.GetPosition Nothing, Pos 'obtain position of the player


'fog setup
FR_Root.SetSceneFogEnable D_FALSE
FR_Root.SetSceneFogColor 126 * 65536 + 125 * 256 + 125
FR_Root.SetSceneBackground 125 * 65536 + 125 * 256 + 125
FR_Root.SetSceneFogMethod D3DRMFOGMETHOD_VERTEX
FR_Root.SetSceneFogParams 1, 500, 1

'SET THE SKY
Set MS_World = D3D_Main.CreateMeshBuilder()
Set T_Sky = D3D_Main.LoadTexture(App.Path & "/ciel.bmp")
MS_World.LoadFromFile App.Path & "/ciel2.x", 0, 0, Nothing, Nothing
MS_World.ScaleMesh 40, 40, 40
MS_World.SetTexture T_Sky
FR_World.AddVisual MS_World




'Load the Bot 1:cyborg
    Set AN_Mech1 = D3D_Main.CreateAnimationSet()
    AN_Mech1.LoadFromFile App.Path & "/cyborg.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing, FR_Mech1
    FR_Mech1.AddScale D3DRMCOMBINE_BEFORE, 0.1, 0.1, 0.1
    FR_Mech1.SetPosition Nothing, 100, -0.1, 40
    Enemy_State = Normal
    AI_Timer = 0
    Cyborg_Sp = False
'load its Bullets
        For i = 0 To 50
            Set m_meshBuilder(i) = D3D_Main.CreateMeshBuilder()
            m_meshBuilder(i).LoadFromFile App.Path & "/ene_bullet.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
            m_meshBuilder(i).ScaleMesh 1, 1, 1
            m_Frame(i).AddVisual m_meshBuilder(i)
            m_Frame(i).SetPosition FR_Mech1, -50, 90, -15
        Next i
     Mech1_Health = 500




    'load player gun
    Set AN_Gun = D3D_Main.CreateAnimationSet()
    AN_Gun.LoadFromFile App.Path & "/wep.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing, FR_Gun
    FR_Gun.AddScale D3DRMCOMBINE_BEFORE, 0.01, 0.01, 0.01
    FR_Gun.SetPosition Nothing, Pos.x, Pos.y, Pos.z
    FR_Gun.AddRotation D3DRMCOMBINE_BEFORE, -90, 0, 0, 0

        'load the bullets
            For i = 0 To 30
                Set MS_Bullet(i) = D3D_Main.CreateMeshBuilder()
                MS_Bullet(i).LoadFromFile App.Path & "/ene_bullet.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
                MS_Bullet(i).ScaleMesh 0.1, 0.1, 0.1
                FR_Gun.GetPosition Nothing, GPos
                FR_Bullet(i).AddVisual MS_Bullet(i)
                FR_Bullet(i).SetPosition Nothing, GPos.x, GPos.y, GPos.z
                ActiveM(i) = False
            Next i

    'set the ammo clips
    Set MS_Ammo = D3D_Main.CreateMeshBuilder()
    MS_Ammo.LoadFromFile App.Path & "/clip.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing
    MS_Ammo.ScaleMesh 1, 1, 1

        For i = 0 To 4 'set the position of each particular clip's position
            FR_Ammo(i).AddVisual MS_Ammo 'also set the mesh of the clip to the frame
            Dim x, z 'variables to hold x and y value
            x = Randomizex_Clip 'procedure 'Randomizex_clip' which randmoizes the x value and then sets it to the variable 'x'
            z = Randomizey_Clip  'procedure 'Randomizey_Clip' which randomizes the y value and then sets the value to variable 'y'
            FR_Ammo(i).SetPosition Nothing, x, 6, z 'set each particualar clip to its x and y value
        Next i

'internal specifications/Player Stats
Bullets = 30 'starts with 30 bullets
Clips = 2 'starts with 2 clips
Player_State = Player_Normal 'current player state

'create the main world
Call AddFloor(1000, 1000, -450, -0.1, -550, App.Path & "/floor.bmp", 40, 40) 'we only need to add a visible floor

'add collision to the surrounding world(prevent walking into no no's land ;)
Call AddCollision(True, 400, -150, -150)
Call AddCollision(True, 400, -150, 400)
Call AddCollision(False, 550, -150, -150)
Call AddCollision(False, 550, 250, -150)

'add game settings, details

    'add yoda
    Set AN_Yoda = D3D_Main.CreateAnimationSet()
    AN_Yoda.LoadFromFile App.Path & "\yoda.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing, FR_Yoda
    FR_Yoda.AddScale D3DRMCOMBINE_BEFORE, 1, 1, 1
    FR_Yoda.SetPosition Nothing, 0, -0.1, 480
    FR_Yoda.SetOrientation FR_Yoda, Sin5, 0, -Cos5, 0, 1, 0
    Yoda_State = NormalY

    'add the battle
    Set AN_Battle = D3D_Main.CreateAnimationSet()
    AN_Battle.LoadFromFile App.Path & "\battle\bafinal.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing, FR_Battle
    FR_Battle.AddScale D3DRMCOMBINE_BEFORE, 2, 2, 2
    FR_Battle.SetPosition Nothing, 900, 0, 0

    'add the hostage
    Set AN_Host = D3D_Main.CreateAnimationSet()
    AN_Host.LoadFromFile App.Path & "\hostage\hostage.x", 0, D3DRMLOAD_FROMFILE, Nothing, Nothing, FR_Host
    FR_Host.SetPosition Nothing, 200, 0, 0
    move = False

'show progress
frmMain.ProgressBar1.Value = 100
frmMain.Label2.Caption = "Initializing Objects...Setting Up Objects...Going IN!!"
End Sub
Public Sub DX_Render()

Do While esc = False
    On Local Error Resume Next
    'timeline counter
    CTime = CTime + 1
    DoEvents 'let windows do its stuff
    Check_Clip 'calls the procedure to check if clip is near
    FR_Mech1.GetPosition Nothing, Pos 'obtains position of mech
    TimeLine 'in game messages
    DX_Mouse ' procedure to use the mouse
    DX_Keyboard 'procedure to use keyboard
    FR_Camera.GetPosition Nothing, GPos 'gets camera/player position
    GDistance = GetDist(Pos.x, Pos.z, GPos.x, GPos.z) 'obtains distance of mech and player
    GAngle = FindAngle(Pos.x, Pos.z, GPos.x, GPos.z) 'obtains angle of mech and player
    D3D_Viewport.Clear D3DRMCLEAR_ZBUFFER Or D3DRMCLEAR_TARGET 'clear view port and sets up everything again
    D3D_Device.Update 'updates 3d visuals
    D3D_Viewport.Render FR_Root 'renders main frame
    DS_Back.DrawText 5, 10, "Clips Available: " & Clips, False 'draws text, in this # of clips
    DS_Back.DrawText 5, 25, "Bullets Available: " & Bullets, False
    DS_Back.DrawText 5, 86, "Mech Health: " & Mech1_Health / 500 * 100 & "%", False 'draws mech health, in percentage
    DS_Back.DrawText 5, 50, "Health: " & Health / 1000 * 100 & "%", False 'draws player health, in percentage
    'draw crosshhair
    DS_Back.DrawLine 640 / 2 - 9, 480 / 2 - 3, 640 / 2 + 9, 480 / 2 - 3
    DS_Back.DrawLine 640 / 2, 480 / 2 - 10, 640 / 2, 480 / 2 + 8
       
    
    DS_Back.DrawText 60, 0, obm, False 'draws in game messages
   
    CurFrameM1 = CurFrameM1 + 1 'gun frames
    
    'checks if player is reloading or not
    If Not Player_State = Reload Then
        Player_State = Firep
        CurFrameG = 150
        AN_Gun.SetTime CurFrameG
    Else
        ' if not then start counter
        CurFrameG = CurFrameG + 50
        If CurFrameG > 1350 Then CurFrameG = 150
        If CurFrameG < 150 Then CurFrameG = 1350
        If CurFrameG > 1350 Then AN_Gun.SetTime 1350 Else AN_Gun.SetTime CurFrameG 'set the animation to currunt frames
        If CurFrameG = 1350 Then Player_State = Firep
    End If
    
    AI_BOTS 'call the ai
    Check_State 'check state of different objects
 
 If CurFrameM1 > 6090 Then CurFrameM1 = 0 'frame control
 If CurFrameM1 < 0 Then CurFrameM1 = 6090

If CurFrameM1 > 6090 Then AN_Mech1.SetTime 6090 Else AN_Mech1.SetTime CurFrameM1 'assign gun frame to animation

DS_Front.Flip Nothing, DDFLIP_WAIT 'flip the surfaces (i.e. front, and back buffers)

    

   
   If Mech1_Health <= 0 Then 'checks if mech is dead
        Enemy_State = die
        
    End If
    
    'checks if mech's bullet has hit the player
    Dim Bpos1 As D3DVECTOR
    For i = 0 To 50
        If i <= BB And ActiveB(i) = True Then
            m_Frame(i).GetPosition Nothing, Bpos1
            
            bcol1 = BulletCol(Bpos1, GPos)
            If bcol1 = True Then
                ActiveB(i) = False
                Health = Health - 3
                m_Frame(i).SetPosition FR_Mech1, -50, 90, -20
            Else
                m_Frame(i).AddTranslation D3DRMCOMBINE_BEFORE, 0, 0, -10
            End If
        End If
    Next i
    
    'checks if player is firing or not
     If Player_State = Firep Then
     For i = 0 To 30
        If i <= BM And ActiveM(i) = True Then
            FR_Bullet(i).GetPosition Nothing, BPos
            FR_Bullet(i).AddTranslation D3DRMCOMBINE_BEFORE, 0, 0, 5
            BCol = BulletCol(BPos, Pos)
            If BCol = True Then
                If Mech1_Health > 0 Then
                    ActiveM(i) = False
                    Mech1_Health = Mech1_Health - 3
                End If
            End If
        End If
    Next i
    End If
    
    'checks if player is dead
  If Health <= 0 Then
    Player_State = dead 'if yes then player state is dead
  End If

 
Loop
End Sub

Public Sub TimeLine()

'THIS PROCEDURE CONTROLS THE STORY OF THE ENTIRE GAME,
'MEANING IT CONTROLS THE EVENTS BEING EXECUTED WHEN CERTAIN REQUIREMENTS ARE MET
    
    If CTime = 10 Then 'Game has started, play objectives
        Object.Play DSBPLAY_DEFAULT
        obm = objective
    End If
    
    If CTime = 180 Then 'this if statement ereases the messages shown, unless otherwise changed
        obm = ""
        FR_Battle.GetPosition Nothing, BaPos 'obtain battle position
        CTime = 11 'set the time back to 11
        If BaPos.y > -200 And CurFrameB >= 160 Then 'verifies conditions regarding the position of the battle and the current battle frame
            FR_Battle.SetPosition Nothing, 0, -200, 0 'hide the battle scene
            CTime = 19000 'continue the story
            obm = "" 'clear messages
        End If
    End If
    
    'if statements below change, and show different messages
    If CTime = 5000 Then
        obm = Soldier1_Text
        CTime = 6000
    End If
    
    If CTime = 6100 Then
        obm = Computer1_Text
        CTime = 6350
    End If
    
    
    
    If CTime = 6400 Then
        obm = Computer1_Texta
        CTime = 11
    End If
    
    'this is used when the mech has died
    If ob = True And CTime = 100 Then
           CTime = 10000
           ob = False
           Enemy_State = Flee
    End If
    
    'gets the mech out of here when it is dead
    If CTime = 10100 Then
        obm = Computer2_Text
        FR_Mech1.SetPosition Nothing, -900, 0, 0
        CTime = 11000
        ob = False
    End If
    
    'if statements below change, and show different messages
    If CTime = 11100 Then
        obm = Computer2_Texta
    End If
    
    If CTime = 11200 Then
        obm = Soldier2_Text
    End If
    
    'enables fog because mech has died, and proceeds through the story
    If CTime = 11250 Then
        obm = Soldier2_Texta
        CTime = 11890
        FR_Root.SetSceneFogEnable D_TRUE
    End If
    
    'another timeline procedure to eliminate messages
    If CTime = 12000 Then
        obm = ""
    End If
        
    'start up yoda
    If CTime = 13090 Then
        obm = Yoda1_Text
        Yoda1.Play DSBPLAY_DEFAULT
    End If
    
    'lines below control the messages and sometimes voice files for yoda
    If CTime = 13130 Then
        obm = Yoda1_Texta
    End If
    
    If CTime = 13230 Then
        obm = Yoda1_Textb
    End If
    
    If CTime = 13310 Then
        obm = Yoda2_Text
    End If
    
    If CTime = 13410 Then
        obm = Yoda2_Textb
    End If
    
    If CTime = 13550 Then
        obm = Yoda2_Textc
    End If
    
    
    
    
    'main battle(jedi) flow, controls the flow of the animation
    If CTime >= 13670 And CTime < 19000 Then
        CurFrameB = CurFrameB + 0.1
        AN_Battle.SetTime CurFrameB
        FR_Battle.SetPosition Nothing, -180, 0, 0
        
        'battle has ended, proceed to ctime 11 to eliminate all battle related events
        If CurFrameB >= 160 Then
            AN_Battle.SetTime 159
            Maul.Play DSBPLAY_DEFAULT
            obm = Maul_Text
            CTime = 11
        End If
        
        'plays different saber sounds corresponding with the animation
        If CurFrameB = 0 Then
            Saber_on.Play DSBPLAY_DEFAULT
        End If
        
        If CurFrameB = 25 Then
            saber_Swing2.Play DSBPLAY_DEFAULT
        End If
        
        If CurFrameB = 30 Then
            Saber_Swing1.Play DSBPLAY_DEFAULT
        End If
        
        If CurFrameB = 60 Then
            Saber_BHit.Play DSBPLAY_DEFAULT
        End If
        
        If CurFrameB = 28 Then
            Saber_Hit.Play DSBPLAY_DEFAULT
        End If
        
        'plays voices of different characters in the animation
        If CurFrameB >= 73 And CurFrameB <= 81 Then
            Darth_Vader1.Play DSBPLAY_DEFAULT
            obm = Dart1_Text
        End If
        
        If CurFrameB >= 127 And CurFrameB <= 132 Then
            Qui.Play DSBPLAY_DEFAULT
            obm = Qui_Text
        End If
        
        
    
    End If
    
    'the final set of events for the timeline
    
    If CTime = 19000 Then
        obm = ""
        Yoda2.Play DSBPLAY_DEFAULT
        obm = Yoda3_Text
        move = True
        CTime = 29000
    End If
    
    If CTime = 29200 Then
        obm = Yoda3_Texta
    End If
    
    If CTime = 29300 Then
        obm = Soldier3_text
    End If
    
    'game has ended, exit!
    If CTime = 29500 Then
        Call Pref.Stop(Seg, SegState, 0, 0)
        Call DX_Exit
    End If
    
    
    
End Sub
Public Sub DX_Keyboard()

'THIS PROCEDURE CONTROLS ALL KEYBOARD INPUT USING DIRECTINPUT

Dim OldPos As D3DVECTOR 'used for collision detection
  FR_Camera.GetPosition Nothing, OldPos
  
DI_Device.GetDeviceStateKeyboard DI_State

If DI_State.Key(DIK_ESCAPE) <> 0 Then 'check if escape key is pressed, if it is, exit
    Call Pref.Stop(Seg, SegState, 0, 0)
    Call DX_Exit
End If

If DI_State.Key(DIK_E) <> 0 Then 'cheat key, speeds up the timeline, plus gives more clips, and lowers the mech substantially
    Clips = Clips + 6
    Mech1_Health = 1
End If

If DI_State.Key(DIK_T) <> 0 Then 'used to talk with yoda
    If Yoda_State = Talk Then
        Player.Play DSBPLAY_DEFAULT
        obm = Player_Text
        CTime = 13000
    End If
End If

If DI_State.Key(DIK_O) <> 0 Then 'used to open hostage door
        move = True
        CurFrameHo = 60 'set the hostage animation to point where it is going away
End If

If DI_State.Key(DIK_D) <> 0 Then 'strafe right
    FR_Camera.SetPosition FR_Camera, 1, 0, 0
    FR_Gun.SetPosition FR_Gun, 1, 0, 0
    For i = 0 To 30
        If ActiveM(i) = False Then
            FR_Bullet(i).SetPosition FR_Bullet(i), 1, 0, 0
        End If
    Next i
End If

If DI_State.Key(DIK_A) <> 0 Then 'strafe left
    FR_Camera.SetPosition FR_Camera, -1, 0, 0
    FR_Gun.SetPosition FR_Gun, -1, 0, 0
    For i = 0 To 30
        If ActiveM(i) = False Then
            FR_Bullet(i).SetPosition FR_Bullet(i), -1, 0, 0
        End If
    Next i
End If

If DI_State.Key(DIK_W) <> 0 Then 'move up
    If DI_State.Key(DIK_LSHIFT) <> 0 Or DI_State.Key(DIK_RSHIFT) <> 0 Then 'run if shift key is pressed
      Walk.Play DSBPLAY_DEFAULT
      FR_Camera.SetPosition FR_Camera, 0, 0, 2.5 ' (Run) Move the viewport forward
      FR_Gun.SetPosition FR_Gun, 0, -1, 2.5
      FR_Gun.SetPosition FR_Gun, 0, 0, 2.5
      For i = 0 To 30
        If ActiveM(i) = False Then
            FR_Bullet(i).SetPosition FR_Bullet(i), 0, 0, 2.5
        End If
      Next i
    Else
      Walk.Play DSBPLAY_DEFAULT
      FR_Camera.SetPosition FR_Camera, 0, 0, 1.5 ' Move the viewport forward
      FR_Gun.SetPosition FR_Gun, 0, 0, 1.5
      For i = 0 To 30
        If ActiveM(i) = False Then
            FR_Bullet(i).SetPosition FR_Bullet(i), 0, 0, 1.5
        End If
      Next i
    End If
  End If
  
If DI_State.Key(DIK_S) <> 0 Then 'move backwards
    Walk.Play DSBPLAY_DEFAULT
    FR_Camera.SetPosition FR_Camera, 0, 0, -1
    FR_Gun.SetPosition FR_Gun, 0, 0, -1
    For i = 0 To 30
         If ActiveM(i) = False Then
            FR_Bullet(i).SetPosition FR_Bullet(i), 0, 0, -1
         End If
    Next i
End If
    'CODE BELOW IS USED FOR COLLISION DETECTION, MAINLY CALLS TO CHECKCOLLISION
  Dim CVector As D3DVECTOR
  FR_Camera.GetPosition Nothing, CVector
  
  CVector = CheckCollision(OldPos, CVector) ' Colision Detection
  
  FR_Camera.SetPosition Nothing, CVector.x, 6, CVector.z
  FR_Gun.SetPosition Nothing, CVector.x, 6, CVector.z
  
  For i = 0 To 30
    If ActiveM(i) = False Then
        FR_Bullet(i).SetPosition Nothing, CVector.x, 6, CVector.z
    End If
  Next i
 
End Sub
Public Sub DX_Mouse()

    'THIS PROCEDURE ENABLES, AND CHECKS MOUSE POSITIONS, AND MOUSE CLICKS, USING POINTAPI FOR MOUSE POSITIONS, AND DINPUT FOR MOUSE CLICKS
    Dim rspeed
    Const RDiv = 100
    Dim MPosNow As POINTAPI
    GetCursorPos MPosNow
    
    'ROTATION OF THE VIEWPORT FOR DIFFERENT REASONS
    If MPos.x > MPosNow.x Then
    rspeed = Abs(MPosNow.x - MPos.x) / RDiv
        If rspeed < 2 Then
            FR_Camera.AddRotation D3DRMCOMBINE_BEFORE, 0, 1, 0, -rspeed
            FR_Gun.AddRotation D3DRMCOMBINE_BEFORE, 0, 1, 0, -rspeed
            
            For i = 0 To 30
                If BM > 1 And ActiveM(i) = False Then
                    FR_Bullet(i).AddRotation D3DRMCOMBINE_BEFORE, 0, 1, 0, -rspeed
                End If
            Next i
        Else
            FR_Camera.AddRotation D3DRMCOMBINE_BEFORE, 1, 0, 0, -1 / RDiv
            FR_Gun.AddRotation D3DRMCOMBINE_BEFORE, 1, 0, 0, -1 / RDiv
            For i = 0 To 30
                If BM > 1 And ActiveM(i) = False Then
                    FR_Bullet(i).AddRotation D3DRMCOMBINE_BEFORE, 1, 0, 0, -1 / RDiv
                End If
            Next i
            
        End If
    ElseIf MPos.x < MPosNow.x Then
    rspeed = Abs(MPos.x - MPosNow.x) / RDiv
        If rspeed < 2 Then
            FR_Camera.AddRotation D3DRMCOMBINE_BEFORE, 0, 1, 0, rspeed
            FR_Gun.AddRotation D3DRMCOMBINE_BEFORE, 0, 1, 0, rspeed
            For i = 0 To 30
                If BM > 1 And ActiveM(i) = False Then
                    FR_Bullet(i).AddRotation D3DRMCOMBINE_BEFORE, 0, 1, 0, rspeed
                End If
            Next i
        Else
            FR_Camera.AddRotation D3DRMCOMBINE_BEFORE, 1, 0, 0, 1 / RDiv
            FR_Gun.AddRotation D3DRMCOMBINE_BEFORE, 1, 0, 0, 1 / RDiv
            For i = 0 To 30
                If BM > 1 And ActiveM(i) = False Then
                    FR_Bullet(i).AddRotation D3DRMCOMBINE_BEFORE, 0, 1, 0, rspeed
                End If
            Next i
        End If
  End If
  
  If MPos.y > MPosNow.y Then
    rspeed = Abs(MPosNow.y - MPos.y) / RDiv
    rspeed = rspeed / 2
    If rspeed < 1.8 Then
    FR_Camera.AddRotation D3DRMCOMBINE_BEFORE, 1, 0, 0, -rspeed
    FR_Gun.AddRotation D3DRMCOMBINE_BEFORE, 1, 0, 0, -rspeed
    For i = 0 To 30
        If BM > 1 And ActiveM(i) = False Then
            FR_Bullet(i).AddRotation D3DRMCOMBINE_BEFORE, 1, 0, 0, -rspeed
        End If
    Next i
    Else
    FR_Camera.AddRotation D3DRMCOMBINE_BEFORE, 1, 0, 0, -1 / RDiv
    FR_Gun.AddRotation D3DRMCOMBINE_BEFORE, 1, 0, 0, -1 / RDiv
    For i = 0 To 30
        If BM > 1 And ActiveM(i) = False Then
            FR_Bullet(i).AddRotation D3DRMCOMBINE_BEFORE, 1, 0, 0, -1 / RDiv
        End If
    Next i
    End If
  ElseIf MPos.y < MPosNow.y Then
    rspeed = Abs(MPosNow.y - MPos.y) / RDiv
    rspeed = rspeed / 2
    If rspeed < 1.8 Then
    FR_Camera.AddRotation D3DRMCOMBINE_BEFORE, 1, 0, 0, rspeed
    FR_Gun.AddRotation D3DRMCOMBINE_BEFORE, 1, 0, 0, rspeed
    For i = 0 To 30
        If BM > 1 And ActiveM(i) = False Then
            FR_Bullet(i).AddRotation D3DRMCOMBINE_BEFORE, 1, 0, 0, rspeed
        End If
    Next i
    Else
    FR_Camera.AddRotation D3DRMCOMBINE_BEFORE, 1, 0, 0, 1 / RDiv
    FR_Gun.AddRotation D3DRMCOMBINE_BEFORE, 1, 0, 0, 1 / RDiv
    For i = 0 To 30
        If BM > 1 And ActiveM(i) = False Then
            FR_Bullet(i).AddRotation D3DRMCOMBINE_BEFORE, 1, 0, 0, 1 / RDiv
        End If
    Next i
    End If
End If
  'if the mouse is too far right, put it on the left
  If MPos.x + 1 = 640 Then SetCursorPos 1, MPos.y
  If MPos.x - 1 = -1 Then SetCursorPos 639, MPos.y
  If MPos.y + 1 = 480 Then SetCursorPos MPos.x, 1
  If MPos.y - 1 = -1 Then SetCursorPos MPos.x, 479
  
  'RESET POSITION VALUES
  MPos.x = MPosNow.x
  MPos.y = MPosNow.y
        
        
  'Makes the camera not barrel rolled
  Dim RC As D3DVECTOR
  Dim RCU As D3DVECTOR
  FR_Camera.GetOrientation Nothing, RC, RCU
  FR_Camera.SetOrientation Nothing, RC.x, RC.y, RC.z, 0, 1, 0
  FR_Gun.SetOrientation Nothing, RC.x, RC.y, RC.z, 0, 1, 0
  For i = 0 To 30
    If BM > 1 And ActiveM(i) = False Then
        FR_Bullet(i).SetOrientation Nothing, RC.x, RC.y, RC.z, 0, 1, 0
    End If
  Next i
  
  DI_Device2.GetDeviceStateMouse DI_State2 'gets mouse status; if key is pressed or not

    If Not DI_State2.Buttons(0) = 0 Then ' if left key is pressed, if so then shoot, also checks if player can shoot
        If BM < 31 Then BM = BM + 1
        If BM < 31 Then ActiveM(BM) = True
        If Bullets > 0 And Bullets < 31 And Player_State = Firep Then
            Bullets = Bullets - 1
            MShoot.Play DSBPLAY_DEFAULT
            Player_State = Firep
        ElseIf Bullets = 0 And Clips > 0 Then
            Msg = "RELOAD!!"
        ElseIf Bullets = 0 And Clips = 0 Then
            Msg = "NO more Ammo!"
        End If
        
    ElseIf DI_State2.Buttons(0) = 0 Then
            Msg = " "
    End If
    
    If DI_State2.Buttons(1) <> 0 Then 'checks if right key is pressed, and reloads
        If Bullets <= 0 And Clips > 0 Then
            Player_State = Reload
            Bullets = 30
            Clips = Clips - 1
            BM = 0
            For i = 1 To 31
                ActiveM(i) = False
            Next i
        ElseIf Bullets > 0 And Bullets < 30 And Not DI_State2.Buttons(1) = 0 Then
            Msg = "You don't need to RELOAD!"
        ElseIf Bullets = 30 Then
            Msg = " "
        Else
            Msg = "NO More Clips!"
        End If
    End If
End Sub
Public Sub DX_Exit()

'THIS PROCEDURE EXITS THE PROGRAM
If Not (Pref Is Nothing) Then Pref.CloseDown 'end BG music
Call DSSound_Exi 'end directsound
Call DD_Main.RestoreDisplayMode
Call DD_Main.SetCooperativeLevel(frmMain.hWnd, DDSCL_NORMAL)
Call DI_Device.Unacquire
Call DI_Device2.Unacquire
ShowCursor True
End
End Sub

Public Sub AI_BOTS()
'--------------------Artificail Intelligence of Mech-------------'
'NOTE: After the normal state procedure, most other code is just a copy and paste, with just some changes
Dim BT_DirX, BT_DirZ
Dim inrang As Boolean
Dim check_area
Dim rotation As Boolean

'obtains position of mech and player
FR_Mech1.GetPosition Nothing, Pos
FR_Camera.GetPosition Nothing, GPos
  
 GAngle = FindAngle(Pos.x, Pos.z, GPos.x, GPos.z) 'calculate angle difference between two objects
  
 GDistance = GetDist(Pos.x, Pos.z, GPos.x, GPos.z) 'calculates the distance between player and mech
 BT_DirX = Pos.x + 20
 BT_DirZ = Pos.z + 20

'the starting state of mech
If Enemy_State = Normal Then
  Dim m, n
  FR_Mech1.SetPosition FR_Mech1, 0, 0, -1
  For i = 0 To 50
    If i <= BB And ActiveB(i) = False Then
    m_Frame(i).SetPosition FR_Mech1, -50, 90, -15
    End If
  Next i
  If GPos.z = BT_DirX - 5 Then
    inrang = True
  End If
  
    Randomize 'randomize?, seems to work sometimes only
    boro = Rnd * 360
    n = Rnd * 3
    If n < 1 Then n = 1
    If n > 2 Then n = 2
    If n > 1 And n < 2 Then n = 1
    
    
'collision detection for the mech
Dim CMVector As D3DVECTOR

  FR_Mech1.GetPosition Nothing, CMVector
  
  CMVector = CheckCollision(Pos, CMVector) ' Colision Detection
  

  FR_Mech1.SetPosition Nothing, CMVector.x, CMVector.y, CMVector.z
  
 m = BOTColl(CMVector)

'lines below check if object is in a certain pos. if so  then rotate
 If CMVector.z = -148.5 Then
    FR_Mech1.SetOrientation FR_Mech1, -boro, 0, -boro, 0, 1, 0
    For i = 0 To 50
         If i <= BB And ActiveB(i) = False Then
            m_Frame(i).SetOrientation m_Frame(i), -boro, 0, -boro, 0, 1, 0
        End If
    Next i
    
    curro = boro
       
End If

If CMVector.z >= 398.5 Then
    FR_Mech1.SetOrientation FR_Mech1, boro, 0, -boro, 0, 1, 0
    
    For i = 0 To 50
         If i <= BB And ActiveB(i) = False Then
            m_Frame(i).SetOrientation m_Frame(i), -boro, 0, -boro, 0, 1, 0
        End If
    Next i
End If

If CMVector.x >= 248.5 Then
 FR_Mech1.SetOrientation FR_Mech1, -boro, 0, 3 * boro, 0, 1, 0
    curro = boro
        
        For i = 0 To 50
         If i <= BB And ActiveB(i) = False Then
            m_Frame(i).SetOrientation m_Frame(i), -boro, 0, -boro, 0, 1, 0
        End If
    Next i
End If

If CMVector.x = -148.5 Then
FR_Mech1.SetOrientation FR_Mech1, -boro, 0, -boro, 0, 1, 0
    curro = boro
        
        For i = 0 To 50
         If i <= BB And ActiveB(i) = False Then
            m_Frame(i).SetOrientation m_Frame(i), -boro, 0, -boro, 0, 1, 0
        End If
    Next i
End If
'--------------checks if the player is in range

    If GDistance <= 100 Then
        Enemy_State = Look
        Objective1.Play DSBPLAY_DEFAULT
        obm = Cyborg1_Text
        CTime = 4800
    End If
    
    
End If

 
If Enemy_State = Look Then 'checks if the enemy is looking
   Clips = Clips + 25 'just to make the game easier ;)
   If BT_DirX > GPos.x And BT_DirZ > GPos.z Then
   FR_Mech1.SetOrientation FR_Mech1, boro, 0, boro + GAngle, 0, 1, 0
   For i = 0 To 50
        If ActiveB(i) = False Then
            m_Frame(i).SetOrientation m_Frame(i), boro, 0, boro + GAngle, 0, 1, 0
        End If
    Next i
   FR_Mech1.SetPosition FR_Mech1, 0, 0, -50
   
   For i = 0 To 50 'sets the bullets to mechs gun if not being fired
    If ActiveB(i) = False Then
        m_Frame(i).SetPosition FR_Mech1, -50, 90, -20
    End If
    Next i
   Enemy_State = Fire
   End If
    
   If BT_DirX > GPos.x And BT_DirZ < GPos.z Then
   FR_Mech1.SetOrientation FR_Mech1, boro, 0, -boro - GAngle, 0, 1, 0
   For i = 0 To 50
        If ActiveB(i) = False Then
            m_Frame(i).SetOrientation m_Frame(i), boro, 0, -boro - GAngle, 0, 1, 0
        End If
    Next i
   FR_Mech1.SetPosition FR_Mech1, 0, 0, -100
   
   For i = 0 To 50
    If ActiveB(i) = False Then
        m_Frame(i).SetPosition FR_Mech1, -50, 90, -20
    End If
    Next i
   Enemy_State = Fire
   End If

If BT_DirX < GPos.x And BT_DirZ > GPos.z Then
   FR_Mech1.SetOrientation FR_Mech1, -boro, 0, -boro - GAngle, 0, 1, 0
   For i = 0 To 50
        If ActiveB(i) = False Then
            m_Frame(i).SetOrientation m_Frame(i), -boro, 0, -boro - GAngle, 0, 1, 0
        End If
    Next i
   FR_Mech1.SetPosition FR_Mech1, 0, 0, -100
   
   For i = 0 To 50
    If ActiveB(i) = False Then
        m_Frame(i).SetPosition FR_Mech1, -50, 90, -20
    End If
    Next i
   Enemy_State = Fire
   End If
   
End If

If Enemy_State = Fire Then
    
  'collision detection, and majoritly the same code as above, except for firing procedure

  FR_Mech1.GetPosition Nothing, CMVector
  
  CMVector = CheckCollision(Pos, CMVector) ' Colision Detection
    Randomize
    boro = Rnd * 360
    n = Rnd * 3
    If n < 1 Then n = 1
    If n > 2 Then n = 2
    If n > 1 And n < 2 Then n = 1
    
    
     For i = 0 To 50
    If BB > i And ActiveB(i) = False Then
        m_Frame(i).SetPosition FR_Mech1, -50, 90, -20
    End If
  Next i
  
 If CMVector.z = -148.5 Then
    FR_Mech1.SetOrientation FR_Mech1, -boro, 0, 3 * -boro, 0, 1, 0
    For i = 0 To 50
        
            m_Frame(i).SetOrientation m_Frame(i), -boro, 0, 3 * -boro, 0, 1, 0
        
    Next i
    
    curro = boro
        
End If

If CMVector.z >= 398.5 Then
    FR_Mech1.SetOrientation FR_Mech1, boro, 0, -boro, 0, 1, 0
    
    For i = 0 To 50
        If BB >= i And ActiveB(i) = False Then
            m_Frame(i).SetOrientation m_Frame(i), boro, 0, -boro, 0, 1, 0
        End If
    Next i
End If

If CMVector.x >= 248.5 Then
 FR_Mech1.SetOrientation FR_Mech1, -boro, 0, 3 * boro, 0, 1, 0
    curro = boro
        
        For i = 0 To 50
        If BB >= i And ActiveB(i) = False Then
            m_Frame(i).SetOrientation m_Frame(i), -boro, 0, 3 * boro, 0, 1, 0
        End If
    Next i
End If

If CMVector.x = -148.5 Then
FR_Mech1.SetOrientation FR_Mech1, -boro, 0, -boro, 0, 1, 0
    curro = boro
        
        For i = 0 To 50
        If BB >= i And ActiveB(i) = False Then
            m_Frame(i).SetOrientation m_Frame(i), -boro, 0, -boro, 0, 1, 0
        End If
    Next i
End If
        

  
  
  FR_Mech1.SetPosition Nothing, CMVector.x, CMVector.y, CMVector.z
    If BB < 51 Then BB = BB + 5 'check bullet status counter
    If BB < 51 Then ActiveB(BB) = True
    If BB >= 51 Then BB = 0
    If BB <= 0 Then
        For i = 0 To 50
            ActiveB(i) = False 'sets all bullets back to mech
        Next i
    End If
    
    
    FR_Mech1.SetPosition FR_Mech1, 0, 0, -4
    
   
End If

If Health <= 0 Then 'checks again if player is dead
    Player_State = dead
End If

If Enemy_State = die Then 'if player is dead
    Mech1_Health = "dead"
    EndA.Play DSBPLAY_DEFAULT
    
    obm = Cyborg2_Text
    ob = True
    FR_Mech1.SetPosition FR_Mech1, 0, 0, 0
    Set AN_Mech1 = Nothing
End If

End Sub


Public Sub Check_State()
'------------This procedure checks different states of different objects, and acting from it----------'
Dim BPosM1 As D3DVECTOR
Dim MPos As D3DVECTOR

If Player_State = dead Then 'if player is dead then exit
    Call Pref.Stop(Seg, SegState, 0, 0)
    Call DX_Exit
End If

For i = 0 To 50 'bullet is set back to mech if outside of world
    m_Frame(i).GetPosition Nothing, BPosM1
    If ActiveB(i) = True And BPosM1.x <= -148.5 Or BPosM1.x >= 238.5 Or BPosM1.z <= -148.5 Or BPosM1.z >= 380 Then
        ActiveB(i) = False
        FR_Mech1.GetPosition Nothing, MPos
        m_Frame(i).SetPosition Nothing, MPos.x, MPos.y, MPos.z
        m_Frame(i).SetPosition FR_Mech1, -50, 90, -15
    End If
Next i

If BB >= 0 And ActiveB(BB) = False Then 'sets bullets back to mech
    For i = 0 To 50
        FR_Mech1.GetPosition Nothing, MPos
        m_Frame(i).SetPosition Nothing, MPos.x, MPos.y, MPos.z
        m_Frame(i).SetPosition FR_Mech1, -50, 90, -15
    Next i
End If

If CTime >= 12100 Then 'yoda's entry

    If Yoda_State = NormalY Then
    CurFrameY = CurFrameY + 1
    
    If CurFrameY > 40 Then CurFrameY = 0
    
    If CurFrameY < 40 Then AN_Yoda.SetTime CurFrameY Else CurFrameY = 0
    FR_Yoda.SetPosition FR_Yoda, 0, 0, 5
    End If
    
    
    
    FR_Yoda.GetPosition Nothing, ypos
    
    
    If ypos.z <= -1 Then
        FR_Yoda.SetPosition FR_Yoda, 0, 0, 0
        CurFrameY = 0
        AN_Yoda.SetTime CurFrameY
        Yoda_State = Talk 'talk yoada
        
    End If
    
End If
If CTime >= 19000 And move = True Then 'sets the hostage to the animation sequence when it is exiting
    
    CurFrameHo = CurFrameHo + 1
    If CurFrameHo < 60 Then CurFrameHo = 60
    If CurFrameHo > 229 Then CurFrameHo = 212
    AN_Host.SetTime CurFrameHo
    
    

    
Else
    CurFrameHo = CurFrameHo + 1
    If CurFrameHo > 60 Then CurFrameHo = 0
    AN_Host.SetTime CurFrameHo
End If
    
End Sub

Public Sub Check_Clip()
'Proccedure checks if clip is in area, if so increase clips
Dim cpos As D3DVECTOR
Dim dist

FR_Camera.GetPosition Nothing, GPos

For i = 0 To 4
    FR_Ammo(i).GetPosition Nothing, cpos

    dist = Sqr((cpos.x - GPos.x) ^ 2 + (cpos.z - GPos.z) ^ 2)

    If dist <= 10 Then
        Clips = Clips + 1
        Dim x, z
        x = Randomizex_Clip
        z = Randomizey_Clip
        FR_Ammo(i).SetPosition Nothing, x, 6, z
    End If
Next i
End Sub

Public Function Randomizex_Clip()
' randomize a value between 0 and 350
Dim rx

Randomize
rx = 350 * Rnd
Randomizex_Clip = rx
End Function

Public Function Randomizey_Clip()
' randomize a value between 0 and 230

Dim ry

Randomize
ry = 230 * Rnd
Randomizey_Clip = ry
End Function


VERSION 5.00
Begin VB.Form MainFrm 
   BorderStyle     =   0  'None
   Caption         =   "Retained Mode Lighting example"
   ClientHeight    =   6165
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6795
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6165
   ScaleWidth      =   6795
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "MainFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Direct3D RM example
'by Jack Hoxley
'www.dx4vb.da.ru

Option Explicit

'These are our three main objects.
Dim DX As New DirectX7
'Directdraw goes with D3D - You always need it.
Dim DD As DirectDraw4
Dim RM As Direct3DRM3

'These are Directdraw variables
Dim SurfPrimary As DirectDrawSurface4
Dim SurfBack As DirectDrawSurface4
Dim DDSDPrimary As DDSURFACEDESC2
Dim DDCapsBack As DDSCAPS2

'These are d3dRM variables
Dim rmDevice As Direct3DRMDevice3
'The viewport represents what you see on screen.
Dim rmViewport As Direct3DRMViewport2
'Your scene is created from lots of frames:
Dim rootFrame As Direct3DRMFrame3
Dim lightFrame As Direct3DRMFrame3
Dim cameraFrame As Direct3DRMFrame3
Dim objectFrame As Direct3DRMFrame3
Dim RGBLightFrame(2) As Direct3DRMFrame3
'These are the lights
Dim Light As Direct3DRMLight
Dim RGBlight(2) As Direct3DRMLight
'A Meshbuilder loads a binary X file (NOT ascii) into your program
'Use a convertor for this.
Dim meshBuilder As Direct3DRMMeshBuilder3

'These 3 are DirectDraw variables
Dim bRunning As Boolean
Dim CurModeActiveStatus As Boolean
Dim bRestore As Boolean

'Frame rate info
Dim LastTime As Long
Dim NumFramesDone As Integer
Dim FrameText As String

'Light vars. We use D3DVECTOR
'because they have the built in values that we want: x,y,z
Dim BlueLight As D3DVECTOR
Dim BlueDir As Boolean
Dim GreenLight As D3DVECTOR
Dim GreenDir As Boolean
Dim RedLight As D3DVECTOR
Dim RedDir As Boolean




Sub InitDX()

    ' create the ddraw object and set the cooperative level
    Set DD = DX.DirectDraw4Create("")
    MainFrm.Show
    DD.SetCooperativeLevel MainFrm.hWnd, DDSCL_FULLSCREEN Or DDSCL_EXCLUSIVE
    
    ' this will be full-screen, so set the display mode
    DD.SetDisplayMode 640, 480, 16, 0, DDSDM_DEFAULT
    
    ' create the primary surface
    DDSDPrimary.lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
    DDSDPrimary.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_3DDEVICE Or DDSCAPS_COMPLEX Or DDSCAPS_FLIP
    DDSDPrimary.lBackBufferCount = 1
    Set SurfPrimary = DD.CreateSurface(DDSDPrimary)
           
    ' get the back buffer
    DDCapsBack.lCaps = DDSCAPS_BACKBUFFER
    Set SurfBack = SurfPrimary.GetAttachedSurface(DDCapsBack)
    
    ' Create the Retained Mode object
    Set RM = DX.Direct3DRMCreate()
    
    SurfBack.SetForeColor RGB(0, 255, 0)
    
    ' Now, create the device from the full screen DD surface
    
    'Set the next line to be either Software rendered (and NOT use 3D Card, use processor)
    'Set rmDevice = RM.CreateDeviceFromSurface("IID_IDirect3DRGBDevice", DD, SurfBack, D3DRMDEVICE_DEFAULT)
    'Set the next line to be either Hardware rendered (and use 3D Card)
    Set rmDevice = RM.CreateDeviceFromSurface("IID_IDirect3DHALDevice", DD, SurfBack, D3DRMDEVICE_DEFAULT)
    
    rmDevice.SetBufferCount 2
    'Keep these next settings for the best all-round results.
    'rmDevice.SetQuality D3DRMLIGHT_ON Or D3DRMFILL_SOLID Or D3DRMRENDER_GOURAUD
    'Use these settings for the best quality when you use lighting and textures
    rmDevice.SetQuality D3DRMLIGHT_ON Or D3DRMRENDER_GOURAUD
    rmDevice.SetTextureQuality D3DRMTEXTURE_NEAREST
    rmDevice.SetRenderMode D3DRMRENDERMODE_BLENDEDTRANSPARENCY
    
    'This is the frame rate info.
    FrameText = "Still Checking"
End Sub



Sub InitScene(sMesh As String)
    ' Create the scene frames
    Set rootFrame = RM.CreateFrame(Nothing)
    Set cameraFrame = RM.CreateFrame(rootFrame)
    Set lightFrame = RM.CreateFrame(rootFrame)
    Set objectFrame = RM.CreateFrame(rootFrame)
    
    ' Set the background color to black.
    'It's more obvious to see this way.
    rootFrame.SetSceneBackgroundRGB 0, 0, 0
    
    'If you move the Z axis value, you can effectively create zooming...
    cameraFrame.SetPosition Nothing, 0, 0, -450
    Set rmViewport = RM.CreateViewport(rmDevice, cameraFrame, 0, 0, 640, 480)
       
    'Set these values between 0 and 1 for the RGB, instead of
    '0 to 255
    'Note, this also affects the brightness of the object. It also blends with
    'the default colour. In this example the object is white; so these colours
    'mix with white for the final colour
    
    'If you want more dramatic lighting, turn down, or turn off the AMBIENT
    'light.
    Set Light = RM.CreateLightRGB(D3DRMLIGHT_AMBIENT, 0.1, 0.1, 0.1)
    rootFrame.AddLight Light
    
    ' create the mesh and load the x file
    Set meshBuilder = RM.CreateMeshBuilder()
    meshBuilder.LoadFromFile sMesh, 0, 0, Nothing, Nothing
    meshBuilder.SetQuality D3DRMLIGHT_ON Or D3DRMRENDER_GOURAUD
    

    ' add the visuals
    meshBuilder.ScaleMesh 4, 4, 1
    objectFrame.AddVisual meshBuilder
    
    'The X is left to right
    'The Y is up/down
    'The Z is + for backwards
    'The Z is - for closer to the screen
    objectFrame.SetPosition Nothing, -60, 90, 5
    
    'If you find your objects disappearing when they go along way away
    'from you, set this number to be larger.
    'It goes this many units away from the camera
    rmViewport.SetBack 600
    
    Call LightStuff
End Sub

Sub LightStuff()
    Set RGBLightFrame(0) = RM.CreateFrame(rootFrame)
    Set RGBLightFrame(1) = RM.CreateFrame(rootFrame)
    Set RGBLightFrame(2) = RM.CreateFrame(rootFrame)
    
    'These lights are currently set to be as bright as possible. If you
    'want more subtle lighting, set the '1' to be '0.5'
    Set RGBlight(0) = RM.CreateLightRGB(D3DRMLIGHT_POINT, 1, 0, 0)
    Set RGBlight(1) = RM.CreateLightRGB(D3DRMLIGHT_POINT, 0, 1, 0)
    Set RGBlight(2) = RM.CreateLightRGB(D3DRMLIGHT_POINT, 0, 0, 1)
    
    'POINT LIGHT
    'These just appear to make a dot of colour on the object. A point light
    'casts light out from every point of it.
    'SPOT
    'These give much more precise lighting; with much less area-of-effect
    'than a point light
    'DIRECTIONAL
    'These just make the object go very bright white - which is all the lights
    'merged together - R+G+B=white
    'AMBIENT
    'Same as the Directional light in this case.
    
    RGBLightFrame(0).AddLight RGBlight(0)
    RGBLightFrame(1).AddLight RGBlight(1)
    RGBLightFrame(2).AddLight RGBlight(2)
    
    UpdateLights
End Sub

Sub RenderLoop()
    Dim t1 As Long
    Dim Pos As D3DVECTOR
    Dim LastTimeV As Long
    
    On Local Error Resume Next
    
    bRunning = True
    t1 = DX.TickCount()
    Do While bRunning = True
        DoEvents
        
        ' this will keep us from trying to blt in case we lose the surfaces (alt-tab)
        bRestore = False
        Do Until ExModeActive
            DoEvents
            bRestore = True
        Loop
    
        ' if we lost and got back the surfaces, then restore them
        DoEvents
        If bRestore Then
            bRestore = False
            DD.RestoreAllSurfaces
        End If
        
        
        'This is retained mode stuff.
        'This line is needed ONLY if you are constantly moving the object
        'It updates it faster the higher the number
        rootFrame.Move 0.5
        UpdateLights
        rmViewport.Clear D3DRMCLEAR_TARGET Or D3DRMCLEAR_ZBUFFER ' clear the viewport
        rmDevice.Update   'blt the image to the screen
        rmViewport.Render rootFrame 'render to the device
        
        
        'This is DirectDraw Stuff
        Call SurfBack.DrawText(10, 10, "D3DRM Full Screen, Click screen to exit", False)
        Call SurfBack.DrawText(10, 30, "Current frame rate: " & FrameText & " fps", False)
        Call SurfBack.DrawText(10, 50, "Vertex count in Object: " & CStr(meshBuilder.GetVertexCount), False)
        SurfPrimary.Flip Nothing, DDFLIP_WAIT
        
        'This basically counts how many frames have been done in the last
        'second. It'll give a fairly accurate idea of how fast it renders. This example
        'does a steady 60fps on an AMD 500 + Savage4 AGP32Mb.....
        'Don't be fooled if it is high if you have a good computer....
            NumFramesDone = NumFramesDone + 1
        If DX.TickCount >= LastTime + 1000 Then
            LastTime = DX.TickCount
            FrameText = CStr(NumFramesDone)
            NumFramesDone = 0
        End If
        
    Loop
End Sub
Function ExModeActive() As Boolean
    Dim TestCoopRes As Long
    
    TestCoopRes = DD.TestCooperativeLevel
    
    If (TestCoopRes = DD_OK) Then
        ExModeActive = True
    Else
        ExModeActive = False
    End If
End Function

Sub UpdateLights()
With RedLight
    If RedDir = False Then
    .x = 0: .y = .y + 3: .z = -25
    Else
    .x = 0: .y = .y - 3: .z = -25
    End If
    
    If .y > 190 Then
    RedDir = True
    End If
    If .y < -210 Then
    RedDir = False
    End If
End With
With GreenLight
    If GreenDir = False Then
    .x = .x + 3: .y = .y + 3: .z = -25
    Else
    .x = .x - 3: .y = .y - 3: .z = -25
    End If
    If .x > 190 And .y > 190 Then
    GreenDir = True
    End If
    If .x < -210 And .y < -210 Then
    GreenDir = False
    End If
End With
With BlueLight
    If BlueDir = False Then
    .x = .x + 3: .y = 0: .z = -25
    Else
    .x = .x - 3: .y = 0: .z = -25
    End If
    
    If .x > 190 Then
        BlueDir = True
    End If
    If .x < -210 Then
        BlueDir = False
    End If
End With

    RGBLightFrame(0).SetPosition Nothing, RedLight.x, RedLight.y, RedLight.z
    RGBLightFrame(1).SetPosition Nothing, GreenLight.x, GreenLight.y, GreenLight.z
    RGBLightFrame(2).SetPosition Nothing, BlueLight.x, BlueLight.y, BlueLight.z
End Sub


Private Sub Form_Load()
On Error GoTo ErrOut:
    Show
    DoEvents
    InitDX
    InitScene App.Path & "\Surface.x"
    RenderLoop
ErrOut:
End
End Sub


Sub EndIT()
    bRunning = False
    Call DD.RestoreDisplayMode
    Call DD.SetCooperativeLevel(Me.hWnd, DDSCL_NORMAL)
    End
End Sub

Private Sub Form_Click()
    EndIT
End Sub


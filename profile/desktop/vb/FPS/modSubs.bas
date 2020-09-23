Attribute VB_Name = "modSubs"
' By Frederico Machado (indiofu@bol.com.br)


Global TexturePath As String ' Our texture directory
Global MeshPath As String ' Our mesh directory

Public Const W_FRONT = 0 ' Means wall in the front
Public Const W_BACK = 1  ' Means wall in the back
Public Const W_LEFT = 2  ' wall in the left
Public Const W_RIGHT = 3 ' wall in the right

Public Const PI = 3.14159265358979

' Adds a wall in any point in the scene
' W = wall, where it will be placed. You can use the consts W_FRONT, W_BACK, ...
' SX means the Size in the X coord
' SY means the Size in the Y coord, easy...
' X = position in the X coord where it will be placed
' Y = position in the Y coord where it will be placed
' Z = position in the Z coord where it will be placed
' Texture = just the filename of the texture
' FU means how many times the texture will be tiled in the X coord
' FV means how many times the texture will be tiled in the Y coord
Public Sub AddWall(W As Integer, SX As Single, SY As Single, x As Single, y As Single, z As Single, Texture As String, FU As Integer, FV As Integer)
  
  Dim WallFace As Direct3DRMFace2 ' Just the Wall Face
  Dim WallTexture As Direct3DRMTexture3 ' It will contain the texture of the wall
  Dim MS_Wall As Direct3DRMMeshBuilder3 ' It will contain the wall
  
  Set MS_Wall = D3D_Main.CreateMeshBuilder() ' Lets create the meshbuilder
  Set WallTexture = D3D_Main.LoadTexture(Texture) ' Loads the texture
  
  Set WallFace = D3D_Main.CreateFace ' Lets create the face
  
  If W = W_FRONT Then
    WallFace.AddVertex x, y, z: WallFace.AddVertex x, y + SY, z: WallFace.AddVertex x + SX, y + SY, z: WallFace.AddVertex x + SX, y, z ' We create the front of the wall
    WallFace.AddVertex x + SX, y, z: WallFace.AddVertex x + SX, y + SY, z: WallFace.AddVertex x, y + SY, z: WallFace.AddVertex x, y, z ' and the back of the wall
  ElseIf W = W_BACK Then
    WallFace.AddVertex x + SX, y, z: WallFace.AddVertex x + SX, y + SY, z: WallFace.AddVertex x, y + SY, z: WallFace.AddVertex x, y, z ' We create the front of the wall
    WallFace.AddVertex x, y, z: WallFace.AddVertex x, y + SY, z: WallFace.AddVertex x + SX, y + SY, z: WallFace.AddVertex x + SX, y, z ' and the back of the wall
  ElseIf W = W_LEFT Then
    WallFace.AddVertex x, y, z: WallFace.AddVertex x, y + SY, z: WallFace.AddVertex x, y + SY, z + SX: WallFace.AddVertex x, y, z + SX ' We create the front of the wall
    WallFace.AddVertex x, y, z + SX: WallFace.AddVertex x, y + SY, z + SX: WallFace.AddVertex x, y + SY, z: WallFace.AddVertex x, y, z ' and the back of the wall
  ElseIf W = W_RIGHT Then
    WallFace.AddVertex x, y, z + SX: WallFace.AddVertex x, y + SY, z + SX: WallFace.AddVertex x, y + SY, z: WallFace.AddVertex x, y, z ' We create the front of the wall
    WallFace.AddVertex x, y, z: WallFace.AddVertex x, y + SY, z: WallFace.AddVertex x, y + SY, z + SX: WallFace.AddVertex x, y, z + SX ' and the back of the wall
  Else
    Exit Sub ' this type of wall doesn't exist for us
  End If
  
  MS_Wall.AddFace WallFace ' Add the face to the meshbuilder
  
  MS_Wall.SetTextureCoordinates 0, 0, FV ' set the texture coordinates in the front of the wall
  MS_Wall.SetTextureCoordinates 1, 0, 0
  MS_Wall.SetTextureCoordinates 2, FU, 0
  MS_Wall.SetTextureCoordinates 3, FU, FV
  MS_Wall.SetTextureCoordinates 4, FU, FV ' set the texture coordinates in the back of the wall
  MS_Wall.SetTextureCoordinates 5, FU, 0
  MS_Wall.SetTextureCoordinates 6, 0, 0
  MS_Wall.SetTextureCoordinates 7, 0, FV
  
  MS_Wall.SetPerspective 1 ' Set the perspective
  
  Set WallFace = MS_Wall.GetFace(0) ' Get the face
  WallFace.SetTexture WallTexture ' set the texture

  FR_World.AddVisual MS_Wall ' Add to its frame

End Sub

' Adds a floor in any point in the scene
' SX means the Size in the X coord
' SY means the Size in the Y coord, easy...
' X = position in the X coord where it will be placed
' Y = position in the Y coord where it will be placed
' Z = position in the Z coord where it will be placed
' Texture = just the filename of the texture
' FU means how many times the texture will be tiled in the X coord
' FV means how many times the texture will be tiled in the Y coord
Public Sub AddFloor(SX As Single, SY As Single, x As Single, y As Single, z As Single, Texture As String, FU As Integer, FV As Integer)

  Dim FloorFace As Direct3DRMFace2 ' The floor face
  Dim FloorTexture As Direct3DRMTexture3 ' It will contain the texture of the wall
  Dim MS_Floor As Direct3DRMMeshBuilder3 ' Will contain the floor
  
  Set MS_Floor = D3D_Main.CreateMeshBuilder() ' Create our meshbuilder
  Set FloorTexture = D3D_Main.LoadTexture(Texture) ' Load our texture
  
  Set FloorFace = D3D_Main.CreateFace ' Lets create our face
  
  FloorFace.AddVertex x, y, z: FloorFace.AddVertex x, y, z + SY: FloorFace.AddVertex x + SX, y, z + SY: FloorFace.AddVertex x + SX, y, z ' Create the floor
  
  MS_Floor.AddFace FloorFace ' Add the floor face
  
  MS_Floor.SetTextureCoordinates 0, 0, FV ' Lets set the texture coordinates
  MS_Floor.SetTextureCoordinates 1, 0, 0
  MS_Floor.SetTextureCoordinates 2, FU, 0
  MS_Floor.SetTextureCoordinates 3, FU, FV
  
  MS_Floor.SetPerspective 1 ' Set the perspective
  
  Set FloorFace = MS_Floor.GetFace(0) ' Get the face
  FloorFace.SetTexture FloorTexture ' Set the texture

  FR_World.AddVisual MS_Floor ' Add the floor to its frame

End Sub

' Adds a floor in any point in the scene
' SX means the Size in the X coord
' SY means the Size in the Y coord, easy...
' X = position in the X coord where it will be placed
' Y = position in the Y coord where it will be placed
' Z = position in the Z coord where it will be placed
' Texture = just the filename of the texture
' FU means how many times the texture will be tiled in the X coord
' FV means how many times the texture will be tiled in the Y coord
Public Sub AddRoof(SX As Single, SY As Single, x As Single, y As Single, z As Single, Texture As String, FU As Integer, FV As Integer)

  Dim RoofFace As Direct3DRMFace2 ' The Roof face
  Dim RoofTexture As Direct3DRMTexture3 ' Will contain the texture
  Dim MS_Roof As Direct3DRMMeshBuilder3 ' The roof meshbuilder
  
  Set MS_Roof = D3D_Main.CreateMeshBuilder() ' Create our meshbuilder
  Set RoofTexture = D3D_Main.LoadTexture(Texture) 'Load the roof texture
  
  Set RoofFace = D3D_Main.CreateFace ' Create the roof face
  
  RoofFace.AddVertex x + SX, y, z: RoofFace.AddVertex x + SX, y, z + SY: RoofFace.AddVertex x, y, z + SY: RoofFace.AddVertex x, y, z ' Create the roof
  
  MS_Roof.AddFace RoofFace ' Add the face
  
  MS_Roof.SetTextureCoordinates 0, 0, FV ' Lets set the texture coordinates
  MS_Roof.SetTextureCoordinates 1, 0, 0
  MS_Roof.SetTextureCoordinates 2, FU, 0
  MS_Roof.SetTextureCoordinates 3, FU, FV
  
  MS_Roof.SetPerspective 1 ' Set the perspective
  
  Set RoofFace = MS_Roof.GetFace(0) ' Get the face
  RoofFace.SetTexture RoofTexture ' Set the texture

  FR_World.AddVisual MS_Roof ' Add the roof to its frame

End Sub


Public Function GetRandom(ByVal Min As Single, ByVal Max As Single) As Single
'by me
    GetRandom = Rnd * (Max - Min) + Min
End Function

Public Function GetDist(intX1 As Single, intZ1 As Single, intX2 As Single, intZ2 As Single) As Single
'by me
    GetDist = Sqr((intX1 - intX2) ^ 2 + (intZ1 - intZ2) ^ 2)

End Function
Public Function FindAngle(intX1 As Single, intZ1 As Single, intX2 As Single, intZ2 As Single) As Single
'by me, partly
Dim sngXComp As Single
Dim sngZComp As Single

    sngXComp = intX2 - intX1
    sngZComp = intZ1 - intZ2
    If sngZComp > 0 Then FindAngle = Atn(sngXComp / sngZComp)
    If sngZComp < 0 Then FindAngle = Atn(sngXComp / sngZComp) + PI
    
    'FindAngle = FindAngle * (180 / PI)
    'Debug.Print FindAngle
End Function


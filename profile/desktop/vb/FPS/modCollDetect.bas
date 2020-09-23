Attribute VB_Name = "modCollDetect"

' By Frederico Machado (indiofu@bol.com.br)
'Editied by AA


Public Type OBJ_COORDS
  XX As Boolean
  XSize As Single
  x As Single
  z As Single
End Type

Global CollObjects() As OBJ_COORDS ' It will contains our walls
Global ObjCount As Integer ' The number of objects that we have to test

' This Sub add a wall to the collision detection
' XX = it is true if the wall is a front or a back wall
' other case it is false, easy!
' XSize is the size of the wall in the X coord
' X = position in the X coord where it will be placed
' Z = position in the Z coord where it will be placed
Public Sub AddCollision(XX As Boolean, XSize As Single, x As Single, z As Single)
  
  ObjCount = ObjCount + 1 ' Add one to the number of objects
  ReDim Preserve CollObjects(ObjCount - 1) As OBJ_COORDS ' Add one to the objects
  
  ' Set the parameters
  CollObjects(ObjCount - 1).XX = XX
  CollObjects(ObjCount - 1).XSize = XSize
  CollObjects(ObjCount - 1).x = x
  CollObjects(ObjCount - 1).z = z
  
End Sub

Public Function CheckCollision(OldPos As D3DVECTOR, NewPos As D3DVECTOR) As D3DVECTOR
  
  For i = 0 To ObjCount - 1
  
    If CollObjects(i).XX = True Then ' It is a front or a back wall
      If NewPos.z < (CollObjects(i).z + 1.5) And NewPos.z > (CollObjects(i).z - 1.5) Then ' Test if we are hitting the wall
        If NewPos.x >= CollObjects(i).x - 1 And NewPos.x <= (CollObjects(i).x + CollObjects(i).XSize) + 1 Then ' Verify if we are between the start and the end of the wall
          If OldPos.z > CollObjects(i).z Then ' Verify what side of the wall we are
            NewPos.z = (CollObjects(i).z + 1.5)
          ElseIf OldPos.z < CollObjects(i).z Then ' Verify what side of the wall we are
            NewPos.z = (CollObjects(i).z - 1.5)
          End If
          GoTo Jump ' Lets verify the next wall
        End If
      End If
    Else ' It is a left or a right wall
      If NewPos.x < (CollObjects(i).x + 1.5) And NewPos.x > (CollObjects(i).x - 1.5) Then ' Test if we are hitting the wall
        If NewPos.z >= CollObjects(i).z - 1 And NewPos.z <= (CollObjects(i).z + CollObjects(i).XSize) + 1 Then ' Verify if we are between the start and the end of the wall
          If OldPos.x > CollObjects(i).x Then ' Verify what side of the wall we are
            NewPos.x = (CollObjects(i).x + 1.5)
          ElseIf OldPos.x < CollObjects(i).x Then ' Verify what side of the wall we are
            NewPos.x = (CollObjects(i).x - 1.5)
          End If
          GoTo Jump ' Lets verify the next wall
        End If
      End If
    End If
    
Jump:
    
  Next ' Test the next wall
  
  ' Set our new position
  CheckCollision.x = NewPos.x: CheckCollision.z = NewPos.z
  
End Function
Public Function BOTColl(CurrPos As D3DVECTOR) As Boolean
 For i = 1 To ObjCount - 1
    If CollObjects(i).XX = True And CurrPos.x = CollObjects(i).x Or CurrPos.z = CollObjects(i).z Then
        
    End If
 Next
End Function

Public Function BulletCol(BullPos As D3DVECTOR, Enemy As D3DVECTOR) As Boolean
    'by me
    Dim dist
    
    dist = GetDist(BullPos.x, BullPos.z, Enemy.x, Enemy.z)
    
    If dist <= 5 Then BulletCol = True
End Function

Attribute VB_Name = "Module1"
Public Coordenadas(100) As Long 'aca estan las X y Y del poligono 1
Public Coordenadas1(100) As Long 'aca estan las X y Y del poligono 2
Public Vertices1 As Integer 'numero de vertices del poligono1
Public Vertices2 As Integer 'numero de vertices del poligono2
Public Vert1(1, 1) As Long
Public Vert2(1, 1) As Long

Public Sub CalculaDiferenciaDeVertices()
If Vertices1 > Vertices2 Then
    ReDim Vert1(Vertices1 / 2, Vertices1 / 2) As Long
    ReDim Vert2(Vertices1 / 2, Vertices1 / 2) As Long
Else
    ReDim Vert1(Vertices2 / 2, Vertices2 / 2) As Long
    ReDim Vert2(Vertices2 / 2, Vertices2 / 2) As Long
End If

End Sub

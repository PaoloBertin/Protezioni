Attribute VB_Name = "Modulo1"
Option Explicit
Option Base 1

Sub DisegnaPolilineaInterattiva()
    Dim oDoc As AcadDocument
    Dim oSpace As AcadModelSpace
    Dim oPline As AcadLWPolyline
    Dim pt() As Double
    Dim pts() As Double
    Dim nPunti As Integer
    
    ' Set oDoc = ThisDrawing
    ' Set oSpace = oDoc.ModelSpace
    
    ' punti
    nPunti = 4
    ReDim pt(1 To nPunti, 1 To 2)
    ReDim pts(1 To nPunti * 2)
    
    pt(1, 1) = 0#
    pt(1, 2) = 0#
    pt(2, 1) = 50#
    pt(2, 2) = 0#
    pt(3, 1) = 50#
    pt(3, 2) = 50#
    pt(4, 1) = 0#
    pt(4, 2) = 50#
    
    pts = MatrixToVector(nPunti, pt)
    
    'Dim oPline As AcadLWPolyline
    Set oPline = ThisDrawing.ModelSpace.AddLightWeightPolyline(pts)
  
    ' Set pl2D = oSpace.AddLightWeightPolyline(pts)
    ' Set oPline = oSpace.AddPolyline(pts)
    oPline.Closed = True
    
    ThisDrawing.Regen acActiveViewport
End Sub

Function MatrixToVector(nPunti As Integer, pt() As Double) As Double()

        ' alloca vettore monodimensionale
        Dim vet() As Double
        ReDim vet(1 To nPunti * 2)
        
        Dim i As Integer
        
        ' costruisce vettore
        For i = 1 To nPunti
            vet(2 * i - 1) = pt(i, 1)  ' X
            vet(i * 2) = pt(i, 2)      ' Y
        Next i
        
        MatrixToVector = vet
End Function

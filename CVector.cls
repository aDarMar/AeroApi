VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CVector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' Classe che rappresenta un vettore
Option Explicit 'It forces to define every parameter in the functions

' Array that stores the components of vector. It is private so that it can be accessed only by the class methods
Private pVec() As Double
Private Pa() As Double
Private Pb() As Double
Private lnt As Double
Private nel As Long
'====== METHODS ======
' Kind of Constructor
Public Sub VecInit(vals() As Double)
'This subroutine assigns the values to the vector
'INPUT
'vals: array containing the coordinates
    ReDim pVec(0 To UBound(vals) - LBound(vals))
    Dim i As Integer
    For i = 0 To UBound(vals) - LBound(vals)
        pVec(i) = vals(LBound(vals) + i)
    Next i
    lnt = Norm(Me)
    nel = VecDims()
End Sub

Public Sub VecInitByPts(Pai() As Double, Pbi() As Double, IsVers As Boolean)
'This module defines the class giving the two points. The vector points from Pa to Pb
' x--------------->
' Pa              Pb
' If IsVers = True the components are scaled so that the vector ha unitary modulus but lnt retains the original value given by the distance bettween the two points
    If Not (UBound(Pai) - LBound(Pai) = UBound(Pbi) - LBound(Pbi)) Then
        Err.Raise 0, "VecInitByPts", "Points must be of same dimensions"
    Else
        Dim i As Long
        ReDim Pa(0 To UBound(Pai) - LBound(Pai))
        ReDim Pb(0 To UBound(Pai) - LBound(Pai))
        ReDim pVec(0 To UBound(Pai) - LBound(Pai))
    
        For i = 0 To UBound(Pai) - LBound(Pai)
            pVec(i) = Pbi(LBound(Pbi) + i) - Pai(LBound(Pai) + i)
            Pa(i) = Pai(LBound(Pai) + i)
            Pb(i) = Pbi(LBound(Pbi) + i)
        Next i
        lnt = Norm(Me)
        nel = VecDims()
        If IsVers Then
            For i = 0 To nel - 1
                pVec(i) = pVec(i) / lnt
            Next i
            'lnt = Norm(Me)
            'If Not (lnt = 1) Then
                'Err.Raise 0, "VecInitByPts", "Norm of the Versor is not Unitary"
            'End If
        End If
    End If
End Sub

Public Sub ApplPointDef(Pai() As Double)
'This module assgns the application point of the vector
    If Not (UBound(Pai) - LBound(Pai) + 1 = nel) Then
        Err.Raise 513, "ApplPointDef", "Application Point must be of teh same dimension of vector"
    Else
    Dim i As Long
    ReDim Pa(0 To UBound(Pai) - LBound(Pai))
        For i = 0 To UBound(Pai) - LBound(Pai)
            Pa(i) = Pai(LBound(Pai) + i)
        Next i
    End If
End Sub

Public Function VecDims() As Integer
'Method that returns the vector dimension
    VecDims = UBound(pVec) - LBound(pVec) + 1
End Function

Public Function ReturnVec() As Double() 'Variant
'Methods that returns private property for evaluation
    ReturnVec = pVec
End Function

Public Function RetPtA() As Double() 'Not Double() [array] becase otherwise one cannot access to a single elemnt(?) v.RetPtAArray(2)
    RetPtA = Pa
End Function

Public Function RetPtB() As Double()
    RetPtB = Pb
End Function

Public Function RetLen() As Double
    RetLen = lnt
End Function

Public Function VecSum(a As CVector, b As CVector) As CVector
'Method that returns the sum of two vectors
    Dim la As Double
    Dim lb As Double
    Dim vb() As Double
    Dim va() As Double
    Dim vs() As Double
    Dim vecS As New CVector
    Dim i As Long
    i = 0
    va = a.ReturnVec()
    vb = b.ReturnVec()  'Otherwise, we cannot use the private property
    lb = b.VecDims()
    la = a.VecDims()
    If Not (la = lb) Then
        Err.Raise 1, "DotProd", "Vectors must be of same dimensions"
    End If
    ReDim vs(0 To la - 1)
    For i = 0 To la - 1
        vs(i) = va(i) + vb(i)
    Next i
    vecS.VecInit vs
    Set VecSum = vecS 'When assigning an object from a function Set is required
End Function

Public Function ScalProd(a As Double) As CVector
'Method that performs the product between a vector and a scalar
    Dim i As Long
    Dim vm() As Double
    Dim vetm As New CVector
    ReDim vm(0 To nel - 1)
    For i = 0 To nel - 1
        vm(i) = pVec(i) * a
    Next i
    vetm.VecInit vm
    Set ScalProd = vetm
End Function

Public Function DotProd(a As CVector, b As CVector) As Double
'Method that performs scalar product between two vectors
    Dim la As Double
    Dim lb As Double
    Dim vb As Variant
    Dim va As Variant
    Dim i As Long
    i = 0
    'dpr = 0
    va = a.ReturnVec()
    vb = b.ReturnVec()  'Otherwise, we cannot use the private property
    lb = b.VecDims()
    la = a.VecDims()
    If Not (la = lb) Then
        Err.Raise 1, "DotProd", "Vectors must be of same dimensions"
    End If
    DotProd = 0
    For i = 0 To la - 1
        DotProd = DotProd + va(i) * vb(i)
    Next i
End Function

Public Function CrossProd(a As CVector, b As CVector) As CVector
'Method that returns the cross product of two vectors
    Dim la As Double
    Dim lb As Double
    Dim vb() As Double
    Dim va() As Double
    Dim vs() As Double
    Dim vecS As New CVector
    Dim i As Long
    i = 0
    va = a.ReturnVec()
    vb = b.ReturnVec()  'Otherwise, we cannot use the private property
    lb = b.VecDims()
    la = a.VecDims()
    If Not (la = lb) And la = 3 Then
        Err.Raise 1, "DotProd", "Vectors must be three dimensional"
    End If
    ReDim vs(0 To 2)
    vs(0) = va(1) * vb(2) - va(2) * vb(1)
    vs(1) = va(2) * vb(0) - va(0) * vb(2)
    vs(2) = va(0) * vb(1) - va(1) * vb(0)
    vecS.VecInit vs
    Set CrossProd = vecS 'When assigning an object from a function Set is required
End Function

Public Function Norm(a As CVector) As Double
'Methods that returns the Euclidean norm of the vector
    Norm = Sqr(a.DotProd(a, a))
End Function

Public Function PtAlAx(lnt As Double) As Double()
'Function that returns the coordinates of a point located at a length lnt along the vector.
'INPUT
'   lnt: distance of the point along the vector. NOTE that lnt must be calculated with the same components of the vector, meaning that if pVec ha unitary norm, then lnt is calculated with the components of the versor
    Dim PtC() As Double
    Dim i As Long
    ReDim PtC(nel - 1)
    
    For i = 0 To nel - 1
        PtC(i) = Pa(i) + lnt * pVec(i)
    Next i
    PtAlAx = PtC
End Function

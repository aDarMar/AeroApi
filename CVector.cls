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
Public Sub VecInit(vals() As Double, Optional IsVers As Boolean = False)
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
    If IsVers Then
            MakeVersor
    End If
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
            MakeVersor
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

Public Sub MakeVersor()
'Transforms the object into a versor
    Dim i As Long
    For i = 0 To nel - 1
        pVec(i) = pVec(i) / lnt
    Next i
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

Public Function NormTo2Vecs(r As CVector, q As CVector) As CVector
'Function that returns a vector that is perpendicular to the first vector (r) and lies in the plane identified by the vectors r,q
'x Function that returns the position along the elastic axis of the projection of a point C, defined as the head of vector q  x
'x
'x          x C
'x         /| n
'x        / v
'x       /  |
'x      /   |
'x     /    |
'x  q /     |
'x   x-->---x--------------------x
'x   A r    C'                   B
'x   |--dn->|
'x n must be perpendiular to r and it must lie in a plane that contains C and the elastic axis. The quations are:
'x   (1) n dot r = 0
'x   (2) n = at * q + bt * r
'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
    Dim n As CVector    'versor that points from point C outside the elastic axis to C', its projection on it
    Dim bt As Double    'Linear combination coefficient for r
    Dim at As Double    'Linear combination coefficient for qr
    Dim i As Long
    Dim j As Long
        
    at = r.DotProd(r, q)
    If at = 0 Then
    'r and s are already perpendicular
        Set NormTo2Vecs = q
    Else
    'finds the vector perpendicular to r and s
        bt = r.DotProd(r, r)
        bt = 1 / bt
        at = -1 / at
        Set n = r.VecSum(q.ScalProd(at), r) 'Assigns vector n components
        n.ApplPointDef q.RetPtB()           'Assigns C as application point
        n.MakeVersor                        'Turns n into a versor
        Set NormTo2Vecs = n
    End If
End Function

Public Function StrLineItsct(r As CVector, n As CVector, Optional tol As Double = 0.00001) As Double()
'Function that finds the intersection between two straight lines. The straight lines are defined by the directional versor and its point of application.
' 3D ONLY
' OUTPUT
'   - Pint: [X,Y,Z,lnt,lnt2] point C' given as the intersection of the two straight lines (A,r) and (C,n). lnt = |C' - A| and lnt2 = |C' - C|
 
    'Finds the projection on the elastic axis of the point. We find the intersection between the two straight lines defined by (r,A)  and (n,C)
    ' Checks if n is a versor, otherwise it will give a wrong length -> TODO REMOVE THIS RESTICTION
    If Abs(n.Norm(n) - 1) > 0.00001 Then
        Err.Raise 501, "StrLineItsct", "n is not a versor"
    End If
    Dim sol As Boolean
    Dim dn As Double
    Dim i As Long
    Dim j As Long
    
    sol = False
    ' Checks if vectors are parallel
    If r.Norm(r.CrossProd(r, n)) < tol Then
        Err.Raise 601, "StrLineItsct", "Straight Lines are parallel"
    End If
    i = 1
    Do
        j = i + 1
        Do
            dn = n.ReturnVec()(j) * r.ReturnVec()(i) - n.ReturnVec()(i) * r.ReturnVec()(j)  ' It is 0 only if n//r
            If Not (dn = 0) Then
                sol = True
                j = j - 1
            End If
            j = j + 1
        Loop Until j > 2 Or sol
        i = i + 1
    Loop Until sol Or i > 2
    ' If sol = False a solutions hasn't been found
    If sol = False Then
        Err.Raise 600, "StrLineItsct", "Intersection not found"
    End If
    i = i - 1   ' Reducing i by 1 to bring it back at the correct value before exiting the cycle
    Dim lnt As Double
    Dim Pint() As Double   ' Intersection Point + Length from r application point
    
    lnt = (n.ReturnVec()(j) * (n.RetPtA()(i) - r.RetPtA()(i)) - n.ReturnVec()(i) * (n.RetPtA()(j) - r.RetPtA()(j))) / dn
    Pint = r.PtAlAx(lnt)
    ReDim Preserve Pint(0 To 4)
    Pint(3) = lnt               ' Distance of the intersection point from the application poin of vector r
    Pint(4) = (lnt * n.ReturnVec()(j) * r.ReturnVec()(i) + n.ReturnVec()(j) * (r.RetPtA()(i) - n.RetPtA()(i))) / (n.ReturnVec()(i) * n.ReturnVec()(j)) ' Distance of the intersection point from the application poin of vector n
    StrLineItsct = Pint
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

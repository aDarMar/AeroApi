Attribute VB_Name = "Elastic_Axis_Import_Module"
Function FindNodNum(wst As Long, rst As Long) As Long
'Funzione che calcola il numero di nodi riportati in una scheda
'INPUT
'wst: worksheet: indica il foglio su cui cercare
'rst: row start:indica la riga da cui iniziare a contare
'OUTPUT
'FindNodNum: numero di elementi nella scheda
    'Dim nval As Long
    Dim ID As Long
    'Dim Row As Long
    'Row = rs
    FindNodNum = 0
    ID = Worksheets(wst).Cells(rst, 1).Value
    Do While ID > 0
        FindNodNum = FindNodNum + 1
        ID = Worksheets(wst).Cells(rst + FindNodNum, 1).Value
    Loop
End Function

Function DefineNode(femap As Object, wst As Long, rst As Long, Optional flg As String = "no", Optional deflg = "yes") As Variant
'Funzione che definisce i nodi in una scheda e fornsice in output una matrice contenente ID x y z
'INPUT
'femap: oggetto che punta alla sessione corrente di femap
'wst: [worksheet] indica il foglio su cui cercare
'rst: [row start] indica la riga da cui iniziare a contare
'flg:           [Beam/Mass] flag that defines the type of input expected from the excel worksheet. If omitted the function will not save additional property data
'deflg:         [yes/no] flag to decide if the function has to also define the nodes contained in the excel file
'OUTPUT
'out: matrice che contiene per ogni riga le seguenti informazioni del nodo :[ID riga x y z -1 [Element Prop.] ]
    If deflg = "yes" Then
        'Create a Node object.
        Dim nd As Object
        Dim rc As Long    '?
        
        Set nd = femap.feNode
    End If
    'Dichiarazioni
    Dim npts As Long     'Numero di nodi definiti nella Worksheet
    Dim Row As Long      'Contatore per le righe
    Dim addR As Long     'Additional properties counter
    Dim x As Double      'x
    Dim y As Double      'y
    Dim z As Double      'z
    Dim l As Long        'layer ID
    Dim co As Long       'color definition
    Dim d As Long        'def CSys
    Dim oc As Long       'output CSys
    Dim e As Long        'Node Type Node = 0 non viene inizializzata quindi � posta automaticamente 0
    Dim ID As Long       'Entity ID
    Dim Pb As Variant    'Vincoli: Indica se il corrispettivo DoF � Libero o No quind Bool
    Dim p(6) As Long     'Vettore per inizializzare pb
    Dim out() As Variant 'Vettore di Output
    'Inizializzazioni
    npts = FindNodNum(wst, rst)
    'Adding more columns to save properties Data for Element definition later on
    If flg = "Beam" Then
        addR = 4        'Saves A Imax,Imin, J
    Else
        If flg = "Mass" Then
            addR = 7    'Saves M Ixx Iyy Izz Ixy Iyz Ixz
        End If
    End If
    ReDim out(0 To npts - 1, 0 To 5 + addR) 'in VSA non si pu� definire direttamente un array con un valore non costante. Inoltre posso decidere il range degli indici arbitrariamente
    Pb = p              '4.     Arrays are passed as variants. Dimension both, assign the array to the variant, and pass the variant.
    Row = rst
    'Legge i Dati per i Nodi
    ID = Worksheets(wst).Cells(Row, 1).Value
    Do While ID > 0
    '5.     Load the local variables with worksheet values.
        ' Node Data
        l = Worksheets(wst).Cells(Row, 2).Value
        co = Worksheets(wst).Cells(Row, 3).Value
        d = Worksheets(wst).Cells(Row, 4).Value
        oc = Worksheets(wst).Cells(Row, 5).Value
        x = Worksheets(wst).Cells(Row, 6).Value
        y = Worksheets(wst).Cells(Row, 7).Value
        z = Worksheets(wst).Cells(Row, 8).Value
        Pb(0) = Worksheets(wst).Cells(Row, 9).Value
        Pb(1) = Worksheets(wst).Cells(Row, 10).Value
        Pb(2) = Worksheets(wst).Cells(Row, 11).Value
        Pb(3) = Worksheets(wst).Cells(Row, 12).Value
        Pb(4) = Worksheets(wst).Cells(Row, 13).Value
        Pb(5) = Worksheets(wst).Cells(Row, 14).Value
        'Salviamo i dati necessari successivamente nell'array out
        out(Row - rst, 0) = ID
        out(Row - rst, 1) = Row
        out(Row - rst, 2) = x
        out(Row - rst, 3) = y
        out(Row - rst, 4) = z
        out(Row - rst, 5) = -1
        'Storing Additional Property Data based on addR
        If Not (flg = "No") Then
            For j = 1 To addR
                out(Row - rst, 5 + j) = Worksheets(wst).Cells(Row, 14 + j).Value
            Next j
        End If
        'Checks if the deflg = yes. In that case the program defines also the nodes in femap
        If deflg = "yes" Then
        'Definiamo l'elemento Node e avanziamo con la riga
            rc = nd.PutAll(ID, x, y, z, l, co, e, d, oc, Pb)    '6.     Put all data back into FEMAP with one call.
        End If
        Row = Row + 1
        ID = Worksheets(wst).Cells(Row, 1).Value
    Loop
    DefineNode = out
End Function

Function CocktailSort(vect As Variant, colr As Integer) As Variant
'Funzione che ordina un vettore in base alla colonna specificata
'INPUT
'vect: vettore da ordinare in senso crescente
'colr: indice della colonna da usare come riferimento da ordinare
    Dim swp As Boolean
    Dim nRow As Long
    Dim nCol As Long
    Dim i As Long
    
    nRow = UBound(vect, 1) - LBound(vect, 1) + 1
    nCol = UBound(vect, 2) - LBound(vect, 1) + 1
    swp = False
    Do
        For i = 0 To nRow - 2
            If vect(i, colr) > vect(i + 1, colr) Then
                vect = swap(vect, i, i + 1, 2)
                swp = True
            End If
        Next i
        If swp = False Then
            Exit Do
        End If
        swp = False
        For i = 0 To nRow - 2
            If vect(nRow - i - 1, colr) < vect(nRow - (i + 2), colr) Then
                vect = swap(vect, nRow - i - 1, nRow - (i + 2), 2)
                swp = True
            End If
        Next i
    Loop While swp = True
    CocktailSort = vect 'Vedi commento funzione swap
End Function

Function swap(vect As Variant, i As Long, j As Long, dm As Integer) As Variant
    'Funzione che preso un vettore e due indici copia il contenuto di un indice nell'altro indice
    'INPUT
    'vect: vettore o matrice da ordinare - pu� avere dimensione massima 2
    'i,j : indici delle posizioni da scambiare
    'dm  : dimensione - 1 swaps columns 2 swaps row
    Dim temp As Variant
    Dim n As Long
    If dm = 1 Then
    'Scambia le colonne
        For n = LBound(vect, dm) To UBound(vect, dm)
            temp = vect(n, i)
            vect(n, i) = vect(n, j)
            vect(n, j) = temp
        Next n
    Else
    'Swaps two rows
        For n = LBound(vect, dm) To UBound(vect, dm)
            temp = vect(i, n)
            vect(i, n) = vect(j, n)
            vect(j, n) = temp
        Next n
    End If
    swap = vect 'vect dovrebbe essere passato per riferimento, ma VSA si aspetta sempre un valore da function senn� dovrei usare sub e avere l'accortezza di avviare il codice selezionando la Sub principale per�
End Function

Function ProjPts(femap As Object, vec1 As Variant, vec2 As Variant) As Variant
'Funzione che proietta i punti di vec2 sull'asse individuato dal primo e ultimo punto di vec1
'   INPUT
'       - vec1,vec2 [ID,row,x,y,z,ds,dX,dY,dZ]
'   OUTPUT
'       - out: [ID -1 x,y,z,lnt,dX,dY,dZ] coordnates of the projected points and offset from the corresponding CG

    Dim nd As Object
    Dim r As New CVector        'versore Ps-Pe ovvero versore dell'asse elastico
    Dim q As New CVector        'versore Ps-Q
    Dim n As CVector            'versore tale da essere perpendicolare all'asse elastico
    Dim Ps() As Double          'Punto iniziale Asse elastico
    Dim Pe() As Double          'Punto finale asse elastico
    Dim Pi() As Double          'Points on the elastic axis
    Dim lnt As Double           'Position of Projected point measured from elastic axis
' Variables for Node Definition
    Dim l As Long               'layer ID
    Dim co As Long              'color definition
    Dim d As Long               'def CSys
    Dim oc As Long              'output CSys
    Dim e As Long               'Node Type Node = 0 non viene inizializzata quindi � posta automaticamente 0
    Dim rc As Long              '?
    Dim ID As Long              'Entity ID
    Dim Pb As Variant           'Vincoli: Indica se il corrispettivo DoF � Libero o No quind Bool
    Dim p(6) As Long            'Vettore per inizializzare pb
    
    Dim out() As Variant
    
    Set nd = femap.feNode
    ID = vec2(LBound(vec2, 1), 0) + 1000 'The starting ID is the ID of the elastic Axis + 1000
    ReDim out(UBound(vec2, 1) - LBound(vec2, 1), 8)
    ReDim Pe(0 To 2)
    ReDim Ps(0 To 2)
    Pb = p
    l = 1
    co = 46
    
' Get first and last node of elastic axis
    For i = 0 To 2
        Ps(i) = vec1(0, 2 + i)
        Pe(i) = vec1(UBound(vec1, 1), 2 + i)
    Next i
    r.VecInitByPts Ps, Pe, True     'Elastic axis versor
' Loop for projecting points along the elastic axis
    For i = 0 To UBound(vec2, 1)
        For j = 0 To 2
            Pe(j) = vec2(i, 2 + j)  'Point to project
        Next j
        q.VecInitByPts Ps, Pe, True
        Set n = r.NormTo2Vecs(r, q)     ' Vector normal to the elastic axis
        Pi = r.StrLineItsct(r, n)   ' Calculates the point given by the projection of the local CG onto the elastic axis
        'lnt = NormVec(r, q)         'Calculates the length of the point from the beginning of the elastic axis
        If Pi(3) > 0 And Not (Pi(3) > r.RetLen()) Then  ' |C' - A| > 0 and < total length
            'Pe = r.PtAlAx(lnt)
            
            rc = nd.PutAll(ID, Pi(0), Pi(1), Pi(2), l, co, e, d, oc, Pb)    '6.     Put all data back into FEMAP with one call.
            out(i, 0) = ID
            out(i, 1) = -1  ' There are no rows
        ' Components of the Projected Vector
            out(i, 2) = Pi(0)
            out(i, 3) = Pi(1)
            out(i, 4) = Pi(2)
            out(i, 5) = Pi(3)
        ' Components of C-C' vector (from C' to C). The - sign is used to revert vector's verse. In this way XCG = X + dX
            out(i, 6) = -n.ReturnVec()(0) * Pi(4)
            out(i, 7) = -n.ReturnVec()(1) * Pi(4)
            out(i, 8) = -n.ReturnVec()(2) * Pi(4)
        
            ID = ID + 1
        End If
    Next i
    ProjPts = out
End Function

Function OrderPoints(eax As Variant, pax As Variant) As Variant
'Function that creates a matrix containing informations about all the points defined along the elastic axis ordered with increasing distance from the fuselage point
'INPUT
'   eax: [ID riga x y z -1 [Element Prop.] ] matrix of points of the elastic axis
'   pax: [ID -1 x,y,z,lnt,dX,dY,dZ] matrix of the mass points projected along the elastic axis.
' OUTPUT
'   tpts: [ ID <ExcelRow|-1> x y z |Pi - P0| <A|dX> <Imax|dY> <Imin|dZ> <J|-1>  ]
'
' xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
    Dim tpts() As Variant       'Matrix containing data of all the points that lie on the elastic axis [ID,Row(-1 if is not defined in the Sheet1),x,y,z,s,A or X_CG_offset,I1 or Y_CG_offset,I2 or Z_CG_Offset,J or -1 (if the three previous entries are the offsets)]
    Dim tvec As New CVector     'Cvector object that defines the segment P0 - Pi
    Dim P0(2) As Double         'Elastic axis starting point (the one enarest the fuselage)
    Dim Pi(2) As Double         'Points on the elastic axis
    
    ReDim tpts(0 To UBound(eax, 1) - LBound(eax, 1) + UBound(pax, 1) - LBound(pax, 1) + 1, 0 To 9)
    'Initializzation
    For i = 0 To 2
        P0(i) = eax(0, 2 + i)   ' El.ax. root node
    Next i

    For i = 0 To UBound(eax, 1) - LBound(eax, 1)
    'Copying data from matrix eax to new storage matrix tpts
        For j = 0 To 9
            tpts(i, j) = eax(i, j)      ' [ID ExcelRow x y z -1 A Imax,Imin, J ]
            If j > 1 And j < 5 Then
                Pi(j - 2) = eax(i, j)   ' Copy coordinates of ith node
            End If
        Next j
    'Computing the distances of points Pi from P0
        tvec.VecInitByPts P0, Pi, False
        tpts(i, 5) = tvec.Norm(tvec)
        ifin = i + 1
    Next i
    'Copying data from matrix pax to new storage matrix tpts
    For i = 0 To UBound(pax, 1) - LBound(pax, 1)
        For j = 0 To 8
            tpts(i + ifin, j) = pax(i, j)   ' Saving in the new matrix the data from the projected vector [ID -1 x,y,z,lnt,dX,dY,dZ]
        Next j
        ' Setting the last row as -1 to indicate that the previous e columns are  X,Y,Z offsets and not A I1 I2
        j = 9
        'For j = 6 To 9
        tpts(i + ifin, j) = -1              ' [ID -1 x,y,z,lnt,dX,dY,dZ,-1]
        'Next j
    Next i
    'Ordering the Vector with respect to the distances column
    tpts = CocktailSort(tpts, 5)
    OrderPoints = tpts
End Function

Function NormVec(r As CVector, q As CVector) As Double
'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
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
    Dim dn As Double    'Lenght of C'A
    Dim sol As Boolean
    Dim i As Long
    Dim j As Long
        
    Set n = r.NormTo2Vecs(r, q)         'Defines the n vector
    sol = False
    'Finds the projection on the elastic axis of the point. We find the intersection between the two straight lines defined by (r,A)  and (n,C)
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
    i = i - 1
    NormVec = (n.ReturnVec()(j) * (n.RetPtA()(i) - r.RetPtA()(i)) - n.ReturnVec()(i) * (n.RetPtA()(j) - r.RetPtA()(j))) / dn
        
End Function

Function InterpMod(f1 As Variant, f2 As Variant, x1 As Variant, x2 As Variant, x0 As Variant) As Variant()
'Function that performs linear interpolation between two values
'
'                             o f2
'        f1 o     ?     ?     |
'           |     |     |     |
'           x-----x-----x-----x
'           x1  x0(0)  x0(1)  x2
'NOTe to Self: An array with just one element is not an array IsArray(x0) = False if ReDim x0(0) ?????
    If IsArray(x0) Then
        Dim m As Double
        Dim i As Long
        Dim out() As Variant
        
        ReDim out(0 To UBound(x0) - LBound(x0))
        m = (f2 - f1) / (x2 - x1)
        
        For i = 0 To UBound(x0) - LBound(x0)
            out(i) = m * (x0(LBound(x0) + i) - x2) + f2
        Next i
        InterpMod = out
    Else
        InterpMod = (f2 - f1) / (x2 - x1) * (x0 - x2) + f2
    End If
End Function

Function Interp(i1 As Long, i2 As Long, j() As Long, apts As Variant, Optional xidx As Long = 0, Optional fidi As Long = 1, Optional fide As Long = -1) As Variant()
'Function that performs linear interpolation on points of matrix apts identified by points j, based on sample points identified by i1,i2 of the same matrix
'INPUT
'   - i1,i2: sample points used for interpolation
'   - j: points where the function is interpolated                                 xidx fidx
'   - apts: matrix containing x,f(x). The matrix must be ordered by columns       [x(0),f(0)]
'   - xidx: column where x-values are stored. Default choice is the first colum   [x(1),f(1)]
'   - fidi: starting colum where the f-values are stored.
'   - fide: ending column where the f-values are stored
'
'               o
'   o           |
'   |           |
'   x-----x-----x
'   i1    j     i2
'
'If IsArray(j) Then
'More than one point to interpolate
    Dim k As Long           ' x index
    Dim out() As Variant    ' interpolated values
    Dim inp() As Variant
    
    If fide = -1 Then
    'It means that data to be interpolated is stored only in one row
        ReDim inp(0 To UBound(j) - LBound(j))
        For k = 0 To UBound(j) - LBound(j)
            inp(k) = apts(k, xidx)
        Next k
        out = InterpMod(apts(i1, fidi), apts(i2, fidi), apts(i1, xidx), apts(i2, xidx), inp)
    Else
    'The function must interpolate more columns
    ReDim inp(0)
        ReDim out(0 To UBound(j) - LBound(j), 0 To fide - fidi)
        For k = 0 To UBound(j) - LBound(j)  ' Number of x0s
        inp(0) = apts(j(LBound(j) + k), xidx)  ' x0 value
            For i = 0 To fide - fidi        ' Number of functions to be interpolated at x0s
                out(k, i) = InterpMod(apts(i1, fidi + i), apts(i2, fidi + i), apts(i1, xidx), apts(i2, xidx), inp)(0) 'TODO: make it better
            Next i
        Next k
    End If
    Interp = out
    
End Function

Function ElAxAxes(P0() As Double, Pe() As Double, ax As String) As CVector
' Function that calculates the axes of a reference system that is obtained from the constructive reference frame with a 3-2 rotation sequence
' with angles Sweep, Dihedral. Sweep here is actually the opposite of the normally defined sweep, meaning that sweep >0 indicates a forward swept wing.
'
'
'
'
'
'             /
'            /
'           o--------------->
'           |\
'           | \
'           |  \
'           |   \
'           |    \
'           v
'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx

        Dim y As New CVector
        Dim x As New CVector
        Dim vec(2) As Double
        Dim Lam As Double
        Dim Dih As Double
        
        y.VecInitByPts P0, Pe, True                                 ' Defining the y_el versor
        Lam = Atn(-y.ReturnVec()(0) / y.ReturnVec()(1))             ' Elastinc Axis sweep
        Dih = Atn(y.ReturnVec()(2) / Sqr(y.ReturnVec()(1) _
        * y.ReturnVec()(1) + y.ReturnVec()(0) * y.ReturnVec()(0)))  ' Elastic Axis Dihedral
        ' x_el is obtained as a rotation around Z axis.
        vec(0) = Cos(Lam)
        vec(1) = Sin(Lam)
        vec(2) = 0
        x.VecInit vec                                               ' x_el CVector object definition
        If ax = "z" Then
            'Dim z As CVector
            Set x = x.CrossProd(x, y)                               ' In this case we calculate z
            'ElAxAxes = z
        End If
        
        Set ElAxAxes = x                                                ' x can be x_el versor or z_el versor
        
        
End Function

Public Sub DefineElems(femap As Object, apts As Variant, mpts As Variant)
'   Input
'       - apts: [ ID <ExcelRow|-1> x y z |Pi - P0| <A|dX> <Imax|dY> <Imin|dZ> <J|-1>  ]
'       - mpts: [ ID ExcelRow x y z M Ixx Iyy Izz Ixy Iyz Ixz ] Mass Properties
'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx
    Dim pr As Object            ' property object
    Dim i As Long               ' index
    Dim j As Long               ' index
    Dim lst() As Long           ' List on indices of points to interpolate
    Dim idx(3) As Long          ' list on indices to access pamt for BEAM end B
    Dim out As Variant          ' interpolation vector for beam end B
    Dim temp() As Double        ' Temp vector used to initialize Variant arrays and other uses
    Dim P0(2) As Double         ' Vector to store the origin of elastic axis
    Dim i1 As Long              ' index of the leftmost sampling point
    Dim i2 As Long              ' index of the rightmost sampling point
    Dim rc As Long              ' error index variable
    
    ReDim lst(0)                ' initializing lst as an array of 1 element because the Interp function takes as input for j an array in all cases
'====== Material Data =======
' Dummy Material
    Dim DMat As Object
    Dim mval As Variant
    
    ReDim temp(0 To 171)
    Set DMat = femap.feMatl     ' Define DMat as a Femap material
    mval = DMat.mmat
    For i = 0 To 5
        mval(i) = 1             '  Ei(1:3) Gi(1:3) = 1
    Next i
    mval(49) = 0                ' rho = 0
    DMat.ID = 1                 ' MatID
    DMat.layer = 1
    DMat.Type = 0               ' Isotropic
    DMat.mmat = mval
    DMat.AutoComplete           ' Auto fill all the other fields
    DMat.Put (1)
'======= Element Data =======
' Beam
    Dim beam As Object
    Dim nds As Variant          ' vnode per BEAM
    Dim orv As Variant          ' vorient TODO: choose correct orientation vector
    Dim nrax As New CVector     ' Cvectir object that defines the orientation vector for BEAM
    
    ReDim temp(2)               ' Gives temp 3 zeros
    
    'orv = temp
    For j = 0 To 2
        P0(j) = apts(0, 2 + j) ' Storing P0: root node
        temp(j) = apts(UBound _
        (apts, 1), 2 + j)       ' Storing Pe: tip node
    Next j

    Set nrax = ElAxAxes(P0, temp, "x") ' Vector containing the x axis of the elastic axis reference frame TODOOOOO x or z?????
    orv = nrax.ReturnVec()
    Set beam = femap.feElem
    
'Rigid
'Mass
'====== Properties Data ======
'Variabile per aggiornate le propriet�
    Dim vflg As Variant          ' vflag
    Dim mat As Variant          ' pmat per BEAM
    Dim mat2 As Variant         ' pmat per MASS
    'Dim temp() As Double        ' Dichiarato senza dimensione iniziale

    ReDim temp(78)              ' Inizializzazione
    mat = temp                  ' Assegnazione all'array mat per l'elemento BEAM
    mat2 = temp                 ' Assegnazione all'array mat per l'elemento MASS
    Dim flg As Boolean          ' Interpolation flag
    
    'Dim temp2() As Long
    'ReDim temp2(4)              ' Ridimensionamento dell'array
    'flg = temp2                 ' Assegnazione a mat
    
    Set pr = femap.feProp
    ReDim out(0 To 3)
    i = 0
    idx(0) = 20
    idx(1) = 21
    idx(2) = 22
    idx(3) = 24
' Assigning data for first node
    For j = 0 To 3
        mat(idx(j)) = apts(0, 6 + j) ' A Imax Imin J
    Next j
' Element and Property definition cycle
    Do While i < UBound(apts, 1)
        pr.Get (999 + i)        ' Get the Property 999+i
        pr.Type = 5             ' Set property as Beam
        mattemp = pr.pmat       ' Copy the material variable
        vflg = pr.vflagI()      ' Get the vflag array
        vflg(0) = 1             ' Sets the Tapered flag to true
        pr.vflagI() = vflg      ' Put the modified vfalg array back
        i1 = i
        i2 = i + 1
    ' Interpolate A I1 I2 J at the Projected Nodes: the following loop tries to find two nodes that have their propreties defined and interpolates
        Do
            flg = False
            If apts(i1, 1) = -1 Then    ' Element i1 is a projected point
                i1 = i1 - 1             ' Go back by 1 position
                flg = True
            End If
            If apts(i2, 1) = -1 Then    ' Element i2 is a projected point
                i2 = i2 + 1
                flg = True
            End If
        Loop Until flg
        lst(0) = i + 1                  ' It is defined as an array because Interp can accept also more points to interpolate
        out = Interp(i1, i2, lst, apts, 5, 6, 9) ' Interpolates column 5 as x0
        
    'End A Properties: taken from the previous element's end B
        mat(0) = mat(20)        ' Area
        mat(1) = mat(21)        ' I1
        mat(2) = mat(22)        ' I2
        mat(4) = mat(24)        ' J
    'End B Properties
        mat(20) = out(0, 0)     ' Area
        mat(21) = out(0, 1)     ' I1
        mat(22) = out(0, 2)     ' I2
        mat(24) = out(0, 3)     ' J
    'BEAM Property Definition
        'pr.Type = 5             ' Beam
        pr.matlID = 1           ' MatID: dummy material, muste be already defined when assigned
        pr.pmat = mat
        'pr.vflag = flg
        pr.Put (1000 + i)
    ' BEAM Element Definition
        beam.Get (1000 + i)
        beam.layer = 1
        beam.Type = 5 'Beam
        beam.propID = 1000 + i
        beam.topology = 0 'Line2
        'nds(0) = olID CAPIRE COME FAR LEGGERE UNA VARIANT
        'nds(1) = ID
        beam.Node(0) = apts(i, 0) 'Questa � una soluzione temporanea
        beam.Node(1) = apts(i + 1, 0)
        'beam.Nodes = nds    'Assegna gli IDs dei nodi che compongono l'elemento
        beam.orientID = 0   'interpreta vorient come componenti vettore che orienta l'asse della sezione
        beam.vorient = orv  'Assegna le componenti del vettore orientamento
        beam.Put (1000 + i)
        i = i + 1
    Loop
' ====== ELEMENT DEFINITION ======
    Dim MassEl As Object
    'Dim vecP As New CVector             ' Vector used to declare offset between CG and its projection
    
    Set MassEl = femap.feElem
    
    'Assigning the DoFs for the single node
    'For i = 0 To 5:
    '    vdof(i) = -1
    'Next i
'Loop for creating Mass and Rigid Elements
    For i = 0 To UBound(mpts, 1) - LBound(mpts, 1)
    ' MASS Property Definition
        pr.Get (2000 + i)
        pr.Type = 27                        ' Mass
        mat = temp                          ' Reset mat to all zeros array
        For j = 1 To 7                      ' Assigns mass values read from excel
            mat(j) = mpts(i, j + 5)         ' M Ixx Iyy Izz Ixy Iyz Ixz with respect to the CG
        Next j
        For j = 8 To 10
            mat(j) = apts(2 * i + 1, j - 2) ' X,Y,Z Offsets from grid point to CG
        Next j
        ' Assign to My and Mz the value of Mx because CONM2 has only one mass definition. If MY and MZ are not equal to MX, Femap uses a CONM1 element
        ' BUT shows a mass with rigid link when running this macro.
        mat(11) = mat(7)
        mat(12) = mat(7)
        
        pr.pmat = mat                       ' Assigns the correct mass data
        pr.Put (2000 + i)                   ' Saves the Element
    ' MASS Element Definition: CONM2, coordinate ID is the default 0, so global retangular. This means that X,Y,Z are offsets
        MassEl.Get (3000 + i)               ' Takes an untouched element
        MassEl.layer = 1
        MassEl.Type = 27                    ' Mass
        MassEl.topology = 9                 ' 9 is Point because the mass element is attached only to a point
        MassEl.propID = 2000 + i            ' Mass Element ID
        MassEl.Node(0) = apts(2 * i + 1, 0) ' Node associated with mass ID
        MassEl.Put (3000 + i)
    Next i
End Sub




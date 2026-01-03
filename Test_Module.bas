Attribute VB_Name = "Test_Module"
Sub TestVettore()
    ' Creazione di due oggetti Vettore
    Dim v1 As New CVector
    Dim v2 As New CVector
    Dim v3 As New CVector
    
    Dim Pa(2) As Double
    Dim Pb(2) As Double
    Dim Pt() As Double
    
    Dim lenv As Double
    Pa(0) = 0
    Pa(1) = 1.2
    Pa(2) = 3.12
    Pb(0) = 1.1
    Pb(1) = 1.32
    Pb(2) = -0.32
    
    v3.VecInitByPts Pa, Pb, False
    
    Dim vsum As CVector
    Dim vprod As CVector
    Dim ScalProd As Double
    
    Dim risultatoProdotto As New CVector
    'Dim prodottoScalare As Double
    Dim insd(2) As Double
    insd(0) = 0#
    insd(1) = 1#
    insd(2) = 2#
    ' Impostazione dei vettori
    v1.VecInit insd
    v2.VecInit insd
    
    lenv = v1.VecDims()
    nrm = v1.DotProd(v1, v1)
    Set vsum = v1.VecSum(v1, v1) 'When assigning an object from a function Set is required
    Set vprod = v1.ScalProd(-2)
    Set vprod = v1.VecSum(v1.ScalProd(-1), v1)
    'risultatoProdotto.VecInit risultatoSomma
    
    v1.ApplPointDef Pa
    v1.PtAlAx (0.5)
    Pt = v1.RetPtA()
    
    ' Test Cross Prod
    Dim x1 As New CVector
    Dim x2 As New CVector
    Dim x3 As New CVector
    
    Dim v(2) As Double
    
    v(1) = 1
    x1.VecInit v
    v(1) = 0
    v(0) = 1
    x2.VecInit v
    
    Set x3 = x3.CrossProd(x1, x2)
    Debug.Print "x3 --"; "nx: "; x3.ReturnVec()(0); "ny: "; x3.ReturnVec()(1); "nz: "; x3.ReturnVec()(2)
    
    
    ' Somma di due vettori usando l'operatore +
    'Set risultatoSomma = v1 + v2
    'MsgBox "Somma dei vettori: " & Join(risultatoSomma.pVettore, ", ")
    
    ' Prodotto di un vettore per uno scalare usando l'operatore *
    'Set risultatoProdotto = v1 * 2
    'MsgBox "Vettore moltiplicato per 2: " & Join(risultatoProdotto.pVettore, ", ")
    
    ' Prodotto scalare di due vettori
    'prodottoScalare = v1.ProdottoScalareVettoriale(v2)
    'MsgBox "Prodotto scalare tra v1 e v2: " & prodottoScalare
End Sub



Private Sub UpdateNodalDataFast()

'1.      Attach to the model in a FEMAP session that is already running.

    Dim femap As Object
    Set femap = GetObject(, "femap.model")

    Dim bpts() As Variant   'Array con le informaizoni dei nodi sull'asse elastico
    Dim mpts() As Variant   'Array con informazioni sui nodi dei centri di massa
    Dim ppts() As Variant   'Array with informations about projected points
    Dim apts() As Variant
    'Definiamo gli Array con i valori dei nodi
    bpts = DefineNode(femap, 1, 2)
    mpts = DefineNode(femap, 2, 2)
    'Ordiniamo gli array e calcoliamo la distanza dal nodo di riferimento
    bpts = CocktailSort(bpts, 3)
    'CAlcoliamo le proiezioni dei punti mpts sull'asse elastico
    ppts = ProjPts(femap, bpts, mpts)
    'Ordering the Points
    apts = OrderPoints(bpts, ppts)
    'Defining the Elements
    
    Dim tst As Long
    
    tst = -12
    Test (tst)
    Test (add(tst))
    
    
    
    
'====== Dimension local variables to receive the data ======
'====== Node Data ======
    Dim x As Double 'x
    Dim y As Double 'y
    Dim yold As Double 'variabile temporanea che salva y dell'iterazione precedente
    Dim z As Double 'z
    Dim l As Long   'layer ID
    Dim co As Long  'color definition
    Dim d As Long   'def CSys
    Dim oc As Long  'output CSys
    Dim e As Long   'Node Type Node = 0 non viene inizializzata quindi è posta automaticamente 0
    Dim rc As Long  '?
    Dim ID As Long  'Entity ID
    Dim olID As Long 'temp ID
    Dim Pb As Variant   'Vincoli: Indica se il corrispettivo DoF è Libero o No quind Bool
    Dim p(6) As Long
    Pb = p              '4.     Arrays are passed as variants. Dimension both, assign the array to the variant, and pass the variant.
'====== PROPERTIES ======
    Dim pr As Object
    Set pr = femap.feProp
' Properties Data
    'Variabile per aggiornate le proprietà
    Dim flg As Variant          ' vflag
    Dim mat As Variant          ' pmat per BEAM
    Dim mat2 As Variant         ' pmat per MASS
    Dim temp() As Double        ' Dichiarato senza dimensione iniziale
    ReDim temp(78)              ' Inizializzazione
    mat = temp                  ' Assegnazione all'array mat per l'elemento BEAM
    mat2 = temp                 ' Assegnazione all'array mat per l'elemento MASS
    Dim temp2() As Long
    ReDim temp2(4)              ' Ridimensionamento dell'array
    flg = temp2                 ' Assegnazione a mat
' Element Data
    Dim beam As Object
    Set beam = femap.feElem
    Dim nds As Variant          ' vnode per BEAM
    Dim orv As Variant          ' vorient
    'Inizializzazione
    Dim temp3(2) As Long
    Dim temp4(3) As Double
    nds = temp3
    orv = temp4
    'DEEBUG
    orv(0) = 1.2
    orv(1) = 0
    orv(2) = 0
    
    Dim k As Integer
    Dim m As Integer
    Dim n As Long
    k = 0
    m = 0
    
    Row = 2

    ID = Worksheets(1).Cells(Row, 1).Value
'====== LETTURA DATI ======
    If ID > 0 Then
    'Inizializziamo i valori di mat altrimenti darà errore quando assegna i valori di B ad A
        mat(20) = Worksheets(1).Cells(Row, 15).Value 'Area
        mat(21) = Worksheets(1).Cells(Row, 16).Value 'I1
        mat(22) = Worksheets(1).Cells(Row, 17).Value 'I2
        mat(24) = Worksheets(1).Cells(Row, 18).Value 'J
    ' Node Data
        l = Worksheets(1).Cells(Row, 2).Value
        co = Worksheets(1).Cells(Row, 3).Value
        d = Worksheets(1).Cells(Row, 4).Value
        oc = Worksheets(1).Cells(Row, 5).Value
        x = Worksheets(1).Cells(Row, 6).Value
        y = Worksheets(1).Cells(Row, 7).Value
        z = Worksheets(1).Cells(Row, 8).Value
        Pb(0) = Worksheets(1).Cells(Row, 9).Value
        Pb(1) = Worksheets(1).Cells(Row, 10).Value
        Pb(2) = Worksheets(1).Cells(Row, 11).Value
        Pb(3) = Worksheets(1).Cells(Row, 12).Value
        Pb(4) = Worksheets(1).Cells(Row, 13).Value
        Pb(5) = Worksheets(1).Cells(Row, 14).Value
        rc = nd.PutAll(ID, x, y, z, l, co, e, d, oc, Pb)    '6.     Put all data back into FEMAP with one call.
        olID = ID
        yold = y
        Row = Row + 1
        ID = Worksheets(1).Cells(Row, 1).Value
        
    End If
    
    While ID > 0
        'Legge i Dati per i Nodi
'5.     Load the local variables with worksheet values.
        ' Node Data
        l = Worksheets(1).Cells(Row, 2).Value
        co = Worksheets(1).Cells(Row, 3).Value
        d = Worksheets(1).Cells(Row, 4).Value
        oc = Worksheets(1).Cells(Row, 5).Value
        x = Worksheets(1).Cells(Row, 6).Value
        y = Worksheets(1).Cells(Row, 7).Value
        z = Worksheets(1).Cells(Row, 8).Value
        Pb(0) = Worksheets(1).Cells(Row, 9).Value
        Pb(1) = Worksheets(1).Cells(Row, 10).Value
        Pb(2) = Worksheets(1).Cells(Row, 11).Value
        Pb(3) = Worksheets(1).Cells(Row, 12).Value
        Pb(4) = Worksheets(1).Cells(Row, 13).Value
        Pb(5) = Worksheets(1).Cells(Row, 14).Value
        
        rc = nd.PutAll(ID, x, y, z, l, co, e, d, oc, Pb)    '6.     Put all data back into FEMAP with one call.
        
        'DEBUG: carica la proprietà di ID 1
        'pr.Get (1)
        'v = pr.pmat
        'Debug.Print "Il valore di x è: "; v(1)
        'v = pr.vflag
        'Debug.Print "Il valore di x è: "; v(1)
        'v = pr.Type
        'Debug.Print "Il valore di x è: "; v(1)
        'If k = 0 Then
        'Proprietà Estremo A
         '   mat(0) = Worksheets(1).Cells(Row, 15).Value 'Area
         '   mat(1) = Worksheets(1).Cells(Row, 16).Value 'I1
         '   mat(2) = Worksheets(1).Cells(Row, 17).Value 'I2
        '    mat(4) = Worksheets(1).Cells(Row, 18).Value 'J
        'Flags
            flg(0) = 1
        '    k = k + 1 'Aggiorniamo k
        
        'Proprietà Estremo A: coincide con quella dell'estremo B dell'elemento precedente
            mat(0) = mat(20)
            mat(1) = mat(21) 'I1
            mat(2) = mat(22) 'I2
            mat(4) = mat(24) 'J
        'Proprietà Estremo B
            mat(20) = Worksheets(1).Cells(Row, 15).Value 'Area
            mat(21) = Worksheets(1).Cells(Row, 16).Value 'I1
            mat(22) = Worksheets(1).Cells(Row, 17).Value 'I2
            mat(24) = Worksheets(1).Cells(Row, 18).Value 'J
        'Definizione della Proprietà BEAM
            pr.Type = 5 'Beam
            pr.matlID = 1 'MatID -> deve essere già definito
            pr.pmat = mat
            pr.vflag = flg
            pr.Put (1000 + m)
            'pr.Get (m) 'Prendiamo la proprietà m-esima (se non esiste la crea?)
            'v = pr.pmat
        'Definizione dell'Elemento BEAM
            beam.layer = 1
            beam.Type = 5 'Beam
            beam.propID = 1000 + m
            beam.topology = 0 'Line2
            'nds(0) = olID CAPIRE COME FAR LEGGERE UNA VARIANT
            'nds(1) = ID
            beam.Node(0) = olID 'Questa è una soluzione temporanea
            beam.Node(1) = ID
            'beam.Nodes = nds    'Assegna gli IDs dei nodi che compongono l'elemento
            beam.orientID = 0   'interpreta vorient come componenti vettore che orienta l'asse della sezione
            beam.vorient = orv  'Assegna le componenti del vettore orientamento
            beam.Put (1000 + m)
        '====== Creazione Proiezioni sull'Asse Elastico ======
            
            
            For n = nold To n_max
            ym = Worksheets(2).Cells(n, 7).Value    'Legge il valore y del nodo con massa
                If y > ym And yold < ym Then    'Questo deve trovarsi tra
                    'Funzione Proiezione punti
                End If
            Next n
            '====== Temporaneo, fare in modo che sia la funzione stessa a inserire i punti sulla Beam
        'Definizione della Proprietà MASS
        ' MASS Element Data
        'mat2(7) = Worksheets(1).Cells(Row, 19).Value ' Se la massa è > 0 allora assegnamo gli altri valori
        'If mat2(7) > 0 Then
         '   For i = 1 To 6
         '       mat2(i) = Worksheets(1).Cells(Row, 19 + i).Value
         '   Next i
         '   pr.Type = 27 'Mass
         '   pr.pmat = mat
         '   pr.Put (2000 + m)
        'End If
            
        
        
        
        
        
            'TODO
            'k = 0   'Resettiamo k a 0
            m = m + 1
'7.     Go to the next row.

        Row = Row + 1
        yold = y        'Salva la coordinata y del punto
        olID = ID
        ID = Worksheets(1).Cells(Row, 1).Value

    Wend

 

End Sub

Public Sub testProp(ID As Integer, pr As Object)
    pr.Get (ID)
    Debug.Print "Property Type: "; pr.Type
    mat = pr.pmat
    For j = LBound(mat) To UBound(mat)
        Debug.Print "Property pmat("; j; "): "; mat(j)
    Next j
End Sub

Public Sub testElem(ID As Integer, MassEl As Object)
    MassEl.Get (ID)
    Debug.Print "Element Type: "; MassEl.Type
    Debug.Print "Element Topology: "; MassEl.topology
    Debug.Print "Element Property ID: "; MassEl.propID
    Debug.Print "Element formulation"
    For j = LBound(MassEl.vformulation) To UBound(MassEl.vformulation)
        Debug.Print "-"; j; ": "; MassEl.vformulation()(j)
    Next j
    Debug.Print "Element Node Referenced"
    For j = LBound(MassEl.vnode) To UBound(MassEl.vnode)
        Debug.Print "-"; j; ": "; MassEl.vnode()(j)
    Next j
End Sub

Private Sub Test()

    Dim femap As Object
    Set femap = GetObject(, "femap.model")
    
    Dim pr As Object
    Dim MassEl As Object

    Set pr = femap.feProp
    Set MassEl = femap.feElem

    Dim mat() As Double
    Dim j As Integer
    
    testProp 2000, pr
    Debug.Print "======================"
    testElem 3001, MassEl
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    
    pr.Get (100)
    Debug.Print "Property Type: "; pr.Type
    mat = pr.pmat
    For j = LBound(mat) To UBound(mat)
        Debug.Print "Property pmat("; j; "): "; mat(j)
    Next j
    Debug.Print "== CONM1 MASS MATRIX PROPERTY ID 2 =="
    pr.Get (2)
    Debug.Print "Property Type: "; pr.Type
    mat = pr.pmat
    For j = LBound(mat) To UBound(mat)
        Debug.Print "Property pmat("; j; "): "; mat(j)
    Next j
    
    
    MassEl.Get (2)               ' Takes an untouched element
    Debug.Print "=== CONM2 MASS ELEMENT: ID 2 ==="
    Debug.Print "Element Type: "; MassEl.Type
    Debug.Print "Element Topology: "; MassEl.topology
    Debug.Print "Element Property ID: "; MassEl.propID
    Debug.Print "Element formulation"
    For j = LBound(MassEl.vformulation) To UBound(MassEl.vformulation)
        Debug.Print "-"; j; ": "; MassEl.vformulation()(j)
    Next j
    Debug.Print "Element Node Referenced"
    For j = LBound(MassEl.vnode) To UBound(MassEl.vnode)
        Debug.Print "-"; j; ": "; MassEl.vnode()(j)
    Next j

    MassEl.Get (3000)               ' Takes an untouched element
    Debug.Print "=== CONM1 MASS ELEMENT: ID 2 ==="
    Debug.Print "Element Type: "; MassEl.Type
    Debug.Print "Element Topology: "; MassEl.topology
    Debug.Print "Element Property ID: "; MassEl.propID
    Debug.Print "Element formulation"
    For j = LBound(MassEl.vformulation) To UBound(MassEl.vformulation)
        Debug.Print "-"; j; ": "; MassEl.vformulation()(j)
    Next j
    Debug.Print "Element Node Referenced"
    For j = LBound(MassEl.vnode) To UBound(MassEl.vnode)
        Debug.Print "-"; j; ": "; MassEl.vnode()(j)
    Next j
    
    
    
    'pr.Type = 27                        ' Mass
    'pr.pmat = mat                       ' Assigns the correct mass data
    'pr.Put (2000 + i)                   ' Saves the Element

    ' MASS Element Definition: CONM2, coordinate ID is the default 0, so global retangular. This means that X,Y,Z are offsets
    
    'MassEl.layer = 1
    'MassEl.Type = 27                    ' Mass
    'MassEl.topology = 9                 ' 9 is Point because the mass element is attached only to a point
    'MassEl.propID = 2000 + i            ' Mass Element ID
    'MassEl.Node(0) = 6100  ' Node associated with mass ID
    'MassEl.Put (3000 + i)
    
End Sub


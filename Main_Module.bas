Attribute VB_Name = "Main_Module"
Public Sub Test(a As Long)
    Debug.Print a
End Sub

Function add(a As Long) As Long
    add = a - 1
End Function

Sub ExportFilesToRepo()
    'Taken from https://minerupset.com/2022/Git-and-Excel-VBA/
    ' Tool for exporting VBA-based components to an external folder (say for storing in a git repo)
    ' Make sure to go to the Toolbar Menu -> Tools -> References -> Select "Microsoft Visual Basic For Applications Extensibility 5.3"
    ' Without this, you will not be able to access the VBA Project of the workbook

    ' Adjust your path name here
    Dim pathName As String: pathName = "d:\Programmes\Git\AeroApi\"

    ' The VBComponent Class represents those objects that make up an Excel Workbook
    Dim vbModule As VBComponent

    ' This loops through each of those VBComponents in the Active Workbook
    For Each vbModule In ActiveWorkbook.VBProject.VBComponents

        ' Some Debug.Print statements for easy testing during development
        Debug.Print vbModule.Name

        ' Runs a selection based on the type of module the component is and either exports it
        ' to the specified path (or doesn't) based on that type. It also adds the correct file extension
        ' based on that type. For a reference on types go to:
        'https://docs.microsoft.com/en-us/office/vba/language/reference/visual-basic-add-in-model/properties-visual-basic-add-in-model#type

        Select Case vbModule.Type
            Case 1
                vbModule.Export pathName & vbModule.Name & ".bas"
                Debug.Print "Exported"
            Case 2
                vbModule.Export pathName & vbModule.Name & ".cls"
                Debug.Print "Exported"
            Case 3
                vbModule.Export pathName & vbModule.Name & ".frm"
                Debug.Print "Exported"
            Case Else
                Debug.Print "Not exporting " & vbModule.Name
        End Select
    Next vbModule
End Sub

Sub ImportFilesToRepo()
    'Taken from https://minerupset.com/2022/Git-and-Excel-VBA/
    ' Tool for importing VBA files from a given folder destination
    ' Make sure to go to the Toolbar Menu -> Tools -> References -> Select "Microsoft Visual Basic For Applications Extensibility 5.3"
    ' Without this, you will not be able to access the VBA Project of the workbook
    ' This module works only if the Modules are already present in the session
    Dim pathName As String: pathName = "d:\Programmes\Git\AeroApi\"

    'Dir is a function that allows you to iterate through files in a directory
    Dim filePath As Variant: filePath = Dir(pathName)
    Dim vbModule As VBComponent

    Do While Len(filePath) > 0

        'Accounting for modules, classes, and forms
        If Right(filePath, 3) = "bas" Or Right(filePath, 3) = "cls" Or Right(filePath, 3) = "frm" Then

            'Remove the edge case of this file
            If filePath <> "gitConnector.bas" Then

                'Need to remove the existing module
                For Each vbModule In ThisWorkbook.VBProject.VBComponents

                    'If the name of the module matches the name of the file (without its extension)
                    If vbModule.Name = Left(filePath, Len(filePath) - 4) Then

                        'Delete the module
                        ThisWorkbook.VBProject.VBComponents.Remove vbModule

                        'Import the new module from the path
                        ActiveWorkbook.VBProject.VBComponents.Import (pathName & filePath)
                    End If
                Next
            End If
        End If

        'This is how Dir iterates to the next file
        filePath = Dir
    Loop
End Sub

Private Sub UpdateNodalDataFast()

'1.      Attach to the model in a FEMAP session that is already running.

    Dim femap As Object
    Set femap = GetObject(, "femap.model")
    
    Dim tst1() As Double
    ReDim tst1(0)   '<- Not an array
    Debug.Print IsArray(tst1)
    Dim tst2() As Variant
    ReDim tst2(0)
    Debug.Print IsArray(tst2)
    
    Dim bpts() As Variant   'Array con le informaizoni dei nodi sull'asse elastico
    Dim mpts() As Variant   'Array con informazioni sui nodi dei centri di massa
    Dim ppts() As Variant   'Array with informations about projected points
    Dim apts() As Variant
    'Definiamo gli Array con i valori dei nodi
    bpts = DefineNode(femap, 1, 2, "Beam")
    mpts = DefineNode(femap, 2, 2, "Mass", "no")
    'Ordiniamo gli array e calcoliamo la distanza dal nodo di riferimento
    bpts = CocktailSort(bpts, 3)
    'CAlcoliamo le proiezioni dei punti mpts sull'asse elastico
    ppts = ProjPts(femap, bpts, mpts)
    'Ordering the Points
    apts = OrderPoints(bpts, ppts)
    'Defining the Elements
    DefineElems femap, apts, mpts
    
    
    
    Dim tst As Long
    
    tst = -12
    Test (tst)
    Test (add(tst))
    
End Sub



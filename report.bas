Attribute VB_Name = "report"
Dim nameSpace As String
Dim nameSpaceURI As String


Sub main()
    
    Dim SwApp As SldWorks.SldWorks
    Dim COSMOSWORKS As COSMOSWORKS
    Dim COSMOSObject As Object
    Dim swVersion As Long
    Dim cwVersion As Long
    Dim cwProgID As String

    'Connect to SOLIDWORKS
    If SwApp Is Nothing Then Set SwApp = Application.SldWorks
    
    'Determine host SOLIDWORKS major version
    swVersion = Left(SwApp.RevisionNumber, 2)

    'Calculate the version-specific ProgID of the Simulation add-in that is compatible with this version of SOLIDWORKS
    cwVersion = swVersion - 15
    cwProgID = "SldWorks.Simulation." & cwVersion
    Debug.Print (cwProgID)

    
    'Get the SOLIDWORKS Simulation object

    Set COSMOSObject = SwApp.GetAddInObject(cwProgID)
    If COSMOSObject Is Nothing Then MsgBox "COSMOSObject object not found"

    Set COSMOSWORKS = COSMOSObject.COSMOSWORKS
    If COSMOSWORKS Is Nothing Then MsgBox "COSMOSWORKS object not found"
    
   'Open the active document and use the COSMOSWORKS API
    Dim theStudy As CWStudy, activeStudyNumber As Integer
    
    activeStudyNumber = COSMOSWORKS.ActiveDoc.StudyManager.ActiveStudy
    Set theStudy = COSMOSWORKS.ActiveDoc.StudyManager.GetStudy(activeStudyNumber)
    
''---------------------------------------------------------''
    '' CHANGING THIS CAUSES CRASHES!! ''
    Dim studyOptions As CWStaticStudyOptions
    Set studyOptions = theStudy.StaticStudyOptions
    
    outputpath = theStudy.StaticStudyOptions.ResultFolder
''---------------------------------------------------------''
    
    Dim xml As DOMDocument60
    Set xml = New DOMDocument60
    
    TemplatePath = SwApp.GetCurrentMacroPathFolder & "\"
    templatename = "FEAreportDataTemplate.xml"
    
    c = Dir(TemplatePath & templatename)
    If c = "" Then
        MsgBox "template not found"
        End
    End If
    
    xml.Load TemplatePath & templatename
    
    q = xml.xml
    l = Len(q)
    If q = "" Then
        MsgBox "template not loaded"
        End
    End If
    
    nameSpaceURI = "foley10x.com/schemas/feaData"
    nameSpace = "xmlns:fea='foley10x.com/schemas/feaData'"
    
    xml.setProperty "SelectionNamespaces", nameSpace
    
    dataToXML xml, "fea:studyName", theStudy.Name
    dataToXML xml, "fea:analysisType", analysisStudyTypeText(theStudy.AnalysisType)
    dataToXML xml, "fea:meshType", theStudy.MeshType
    
    getStudyOptions theStudy, xml
    getLoadAndRestraints theStudy, xml
    getMesh theStudy, xml
    getMaterials theStudy, xml
     
    savefilename = theStudy.Name & "-FEAreportData.xml"
    xml.Save outputpath & "\" & savefilename '"\FEAreportData.xml"
    
    'theStudy.Results.SetPlotDisplayOptions
    'theStudy.Results.SavePlotsAseDrawings
    'theStudy.Results.SetPlotPositionFormatOptions
     
    
End Sub

Sub getStudyOptions(theStudy As CWStudy, theXMLDoc As DOMDocument60)
'Sub getStudyOptions(studyOptions As CWStaticStudyOptions, theXMLDoc As DOMDocument60)

    Dim studyOptions As CWStaticStudyOptions
    Set studyOptions = theStudy.StaticStudyOptions
    
    With studyOptions
    
        dataToXML theXMLDoc, "fea:solverType", solverTypeText(.SolverType)
        dataToXML theXMLDoc, "fea:useInplaneEffect", onOffText(.UseInPlaneEffect)
        dataToXML theXMLDoc, "fea:useSoftSpring", onOffText(.UseSoftSpring)
        dataToXML theXMLDoc, "fea:useIniterialRelief", onOffText(.UseInertialRelief)
        dataToXML theXMLDoc, "fea:incompatibleBonding", incompatibleBondingOptionText(.IncompatibleBondingOption)
        dataToXML theXMLDoc, "fea:useLargeDisplacement", onOffText(.LargeDisplacement)
        dataToXML theXMLDoc, "fea:computeFreeBodyForces", onOffText(.ComputeFreeBodyForce)
        dataToXML theXMLDoc, "fea:useFriction", onOffText(.IncludeGlobalFriction)
        dataToXML theXMLDoc, "fea:useAdaptiveMethod", adaptiveMethodText(.AdaptiveMethodType)
        
    End With
    
End Sub

Sub getLoadAndRestraints(theStudy As CWStudy, theXMLDoc As DOMDocument60)
    
    Dim landrmgr As CWLoadsAndRestraintsManager
    Dim landr As CWLoadsAndRestraints
    Dim errCode As Long
    Dim f1 As CWBearingLoad
    Dim r1 As CWRestraint
    
    Set landrmgr = theStudy.LoadsAndRestraintsManager
    i = landrmgr.Count
 
    'swsLoadsAndRestraintsType_e.swsLoadsAndRestraintsTypeBearingLoads
    'swsLoadsAndRestraintsType_e.swsLoadsAndRestraintsTypeConnectors
    'swsLoadsAndRestraintsType_e.swsLoadsAndRestraintsTypeForce
    'swsLoadsAndRestraintsType_e.swsLoadsAndRestraintsTypeGravity
    'swsLoadsAndRestraintsType_e.swsLoadsAndRestraintsTypeMeshControl
    'swsLoadsAndRestraintsType_e.swsLoadsAndRestraintsTypePressure
    'swsLoadsAndRestraintsType_e.swsLoadsAndRestraintsTypeRemoteLoad
    'swsLoadsAndRestraintsType_e.swsLoadsAndRestraintsTypeRemoteMass
    'swsLoadsAndRestraintsType_e.swsLoadsAndRestraintsTypeRestraint
    
    For j = 0 To i - 1
    
        Set landr = landrmgr.GetLoadsAndRestraints(j, errCode)
        
        Debug.Print j, landr.Name,
        
        Select Case landr.Type
        
            Case swsLoadsAndRestraintsTypeRestraint
                addRestraintToXML theXMLDoc, landr
             
            Case swsLoadsAndRestraintsTypeForce
                addForceToXML theXMLDoc, landr
                 
            Case swsLoadsAndRestraintsTypePressure
                'addPressureToXML theXMLDoc, landr
                 
            Case swsLoadsAndRestraintsTypeBearingLoads
                addBearingLoadToXML theXMLDoc, landr
                
            Case swsLoadsAndRestraintsTypeGravity
                'addGravityToXML theXMLDoc, landr
                
            Case swsLoadsAndRestraintsTypeRemoteLoad
                'addRemoteLoadToXML theXMLDoc, landr
                
            Case swsLoadsAndRestraintsTypeRemoteMass
                'addremotemasstoxml theXMLDoc, landr
                
            Case swsLoadsAndRestraintsTypeConnectors
                'addconnectorstoxml theXMLDoc, landr
                
            Case swsLoadsAndRestraintsTypeMeshControl
                'addmeshcontroltoxml theXMLDoc, landr
            
            Case 34
                'pin connector, does not seem to be part of enumeration
                
            Case Else
                
                MsgBox "Load and Restraints Type not found" & vbLf & landr.Name
                
                End
                
        End Select
        
    Next
    
End Sub

Sub addRestraintToXML(theXMLDoc As DOMDocument60, theRestraint As CWLoadsAndRestraints)
    
    Dim parentElement As IXMLDOMElement
    Dim restraintElement As IXMLDOMElement
    Dim typeElement As IXMLDOMElement, nameElement As IXMLDOMElement
    Dim translationElement As IXMLDOMElement, rotationElement As IXMLDOMElement
    
    Dim b1 As Long, b2 As Long, b3 As Long
    Dim d1 As Double, d2 As Double, d3 As Double
    
    Set parentElement = theXMLDoc.selectNodes("//fea:restraints").Item(0)
        
    Set restraintElement = theXMLDoc.createNode(NODE_ELEMENT, "restraint", nameSpaceURI)
    
    Set nameElement = theXMLDoc.createNode(NODE_ELEMENT, "restraintName", nameSpaceURI)
    Set typeElement = theXMLDoc.createNode(NODE_ELEMENT, "restraintType", nameSpaceURI)
    
    Set translationElement = theXMLDoc.createNode(NODE_ELEMENT, "translation", nameSpaceURI)
    Set rotationElement = theXMLDoc.createNode(NODE_ELEMENT, "rotation", nameSpaceURI)
    
    
    With theRestraint
    
        typeElement.setAttribute "displayLabel", "Type"
        typeElement.Text = restraintTypeText(.RestraintType)
        
        nameElement.setAttribute "displayLabel", "Name"
        nameElement.Text = .Name
        
        'translationElement = .GetTranslationComponentsValues(b1, b2, b3, d1, d2, d3)
        'rotationElement = .GetRotationComponentsValues(b1,b2,b3,d1,d2,d3
        
    End With
    
    restraintElement.appendChild typeElement
    restraintElement.appendChild nameElement
    'restraintElement.appendChild translationElement
    'restraintElement.appendChild rotationElement
    
    parentElement.appendChild restraintElement

End Sub

Sub addForceToXML(theXMLDoc As DOMDocument60, theLoad As CWLoadsAndRestraints)
'force.Equation
'force.ForceType
'force.GetForceComponentValues
'force.GetMomentComponentValues
'force.NormalForceOrTorqueValue
'force.Unit

    Dim force As CWForce
    Dim fb1 As Long, fb2 As Long, fb3 As Long
    Dim fd1 As Double, fd2 As Double, fd3 As Double
    
    Dim parentElement As IXMLDOMElement, loadElement As IXMLDOMElement, nameElement As IXMLDOMElement
    
    Set parentElement = theXMLDoc.selectSingleNode("//fea:loads")
    Set loadElement = theXMLDoc.createNode(NODE_ELEMENT, "load", nameSpaceURI)
    
    Set nameElement = theXMLDoc.createNode(NODE_ELEMENT, "loadName", nameSpaceURI)
    nameElement.Text = theLoad.Name
    nameElement.setAttribute "displayLabel", "Name"
    
    
    Set force = theLoad
    theType = force.ForceType
    
    'swsForceType_e.swsForceTypeForceOrMoment
    'swsForceType_e.swsForceTypeNormal
    'swsForceType_e.swsForceTypeTorque
    
    Select Case theType
        
        Case swsForceTypeNormal
        
            v = force.NormalForceOrTorqueValue
            u = force.Unit
            
            'swsForceUnit_e.swsForceUnitkgfOrkgfcm
            'swsForceUnit_e.swsForceUnitlbOrlbin
            'swsForceUnit_e.swsForceUnitNOrNm
            
            eqn = force.Equation
            eqnu = force.EquationLinearUnit
            
            Set typeElement = theXMLDoc.createNode(NODE_ELEMENT, "loadType", nameSpaceURI)
            Set valueElement = theXMLDoc.createNode(NODE_ELEMENT, "loadValue", nameSpaceURI)
            
            typeElement.setAttribute "displayLabel", "Type"
            valueElement.setAttribute "displayLabel", "Value"
            
            typeElement.Text = "Normal force"
            valueElement.Text = force.NormalForceOrTorqueValue
            
            loadElement.appendChild nameElement
            loadElement.appendChild typeElement
            loadElement.appendChild valueElement
            
            parentElement.appendChild loadElement
            
        Case swsForceTypeTorque
        Case swsForceTypeForceOrMoment
            
            force.GetForceComponentValues fb1, fb2, fb3, fd1, fd2, fd3
            forcex = fb1 * fd1
            forcey = fb2 * fd2
            forcez = fb3 * fd3
            
            Set typeElement = theXMLDoc.createNode(NODE_ELEMENT, "loadType", nameSpaceURI)
            typeElement.setAttribute "displayLabel", "Type"
            
            Set dir1element = theXMLDoc.createNode(NODE_ELEMENT, "dir1", nameSpaceURI)
            Set dir2element = theXMLDoc.createNode(NODE_ELEMENT, "dir2", nameSpaceURI)
            Set dir3element = theXMLDoc.createNode(NODE_ELEMENT, "dir3", nameSpaceURI)
            
            typeElement.Text = "Force"
            dir1element.Text = forcex
            dir2element.Text = forcey
            dir3element.Text = forcez
            
            loadElement.appendChild nameElement
            loadElement.appendChild typeElement
            loadElement.appendChild dir1element
            loadElement.appendChild dir2element
            loadElement.appendChild dir3element
            
            parentElement.appendChild loadElement
            
    End Select
    
      
End Sub


'pressure.Equation
'pressure.EquationLinearUnit
'pressure.PressureType
'pressure.Unit
'pressure.Value
'pressure.IncludeNonUniformDistribution
        
        
Sub addBearingLoadToXML(theXMLDoc As DOMDocument60, theLoad As CWLoadsAndRestraints)
'bearingLoad.BearingLoadUnit
'bearingLoad.Direction
'bearingLoad.XDirectionValue
'bearingLoad.YDirectionValue

    'swsUnit_e.swsUnitEnglish
    'swsUnit_e.swsUnitMetric
    'swsUnit_e.swsUnitEnglish
    
    Dim bload As CWBearingLoad
    Dim parentElement As IXMLDOMElement
    Dim loadElement As IXMLDOMElement
    
    Set parentElement = theXMLDoc.selectNodes("//fea:loads").Item(0)
    
    Set loadElement = theXMLDoc.createNode(NODE_ELEMENT, "load", nameSpaceURI)
    
    Set nameElement = theXMLDoc.createNode(NODE_ELEMENT, "loadName", nameSpaceURI)
    Set typeElement = theXMLDoc.createNode(NODE_ELEMENT, "loadType", nameSpaceURI)
    Set valueElement = theXMLDoc.createNode(NODE_ELEMENT, "loadValue", nameSpaceURI)
    'Set transverseelement = theXMLDoc.createElement("transverseValue")
    
    Set bload = theLoad
    
    nameElement.Text = theLoad.Name
    typeElement.Text = "Bearing load"
    valueElement.Text = bload.XDirectionValue
    
    nameElement.setAttribute "displayLabel", "Name"
    typeElement.setAttribute "displayLabel", "Type"
    valueElement.setAttribute "displayLabel", "Value"

    loadElement.appendChild nameElement
    loadElement.appendChild typeElement
    loadElement.appendChild valueElement
    
    parentElement.appendChild loadElement
    
End Sub

Sub getMesh(theStudy As CWStudy, theXMLDoc As DOMDocument60)
              
    Dim dmax As Double, dmin As Double
    Dim defaultSize As Double, defaultTolerance As Double
    
    With theStudy.mesh
      
        'standard mesh
        .GetDefaultElementSizeAndTolerance swsLinearUnitInches, defaultSize, defaultTolerance
        .GetDefaultMaxAndMinElementSize swsLinearUnitInches, dmax, dmin
        c = dmax - dmin
        
        'default size does not change, even if a size is entered directly
        'this is the starting size.
        'setting spinner to "course" doubles it
        'setting spinner to "fine" half's it
        
        a = .ElementSize / 0.0254
        t = .Tolerance / 0.0254 '(min size if curvature based)
        
        'for curvature based mesh only: (0 if not curvature based)
        'if changed from curvature back to standard, previous values stay
        x = .MaxElementSize / 0.0254
        n = .MinElementSize / 0.0254
        
        
        dataToXML theXMLDoc, "fea:meshType", meshTypeText(.MeshType)
        dataToXML theXMLDoc, "fea:mesherUsed", mesherTypeText(.MesherType)
        dataToXML theXMLDoc, "fea:elementCount", .ElementCount
        dataToXML theXMLDoc, "fea:jacobianPoints", jacobianPointsText(.UseJacobianCheckForSolids)
        dataToXML theXMLDoc, "fea:jacobianCheckForShell", jacobianPointsText(.UseJacobianCheckForShells)
        
        
        .GetDefaultElementSizeAndTolerance swsLinearUnitInches, defaultSize, defaultTolerance
        
        dataToXML theXMLDoc, "fea:defaultElementSize", Round(defaultSize, 4)
        dataToXML theXMLDoc, "fea:elementSize", Round(.ElementSize / 0.0254, 4)
        
        dataToXML theXMLDoc, "fea:meshQuality", meshQualityText(.Quality)
        dataToXML theXMLDoc, "fea:worstJacobian", Round(.GetWorstJacobianRatio, 4)
        
    End With
    
End Sub

Sub getMaterials(theStudy As CWStudy, theXMLDoc As DOMDocument60)
    
    Dim theComponent As CWSolidComponent
    Dim theSolidBody As CWSolidBody
    Dim theMaterial As CWMaterial
    Dim nUnit As Long, E As String, nu As String, Su As String, Se As String
    Dim materials As Dictionary
    
    Dim errCode As Long
    
    Set materials = New Dictionary
    
    ComponentCount = theStudy.SolidManager.ComponentCount
    
    For j = 0 To ComponentCount - 1
    
        Set theComponent = theStudy.SolidManager.GetComponentAt(j, errCode)
        
        bodycount = theComponent.SolidBodyCount
        
        For i = 0 To bodycount - 1
            
            Set theSolidBody = theComponent.GetSolidBodyAt(i, errCode)
            Set theMaterial = theSolidBody.GetSolidBodyMaterial
            
            hasmaterial = Not theMaterial Is Nothing
            If hasmaterial Then
                
                If Not materials.Exists(theMaterial.MaterialName) Then
                    materials.Add theMaterial.MaterialName, theMaterial
                End If
            End If
            
        Next
    
    Next
    
    materialsToXML theXMLDoc, materials
    
End Sub

Sub materialsToXML(theXMLDoc As DOMDocument60, theMaterials As Dictionary)
'output materials to XML file
    
    Dim parentElement As IXMLDOMElement
    Dim childElement As IXMLDOMElement
    Dim curMaterial As CWMaterial
    
    'get parent element and save child template
    Set parentElement = theXMLDoc.selectNodes("//fea:materials").Item(0)
    Set childElement = parentElement.firstChild
    
    parentElement.removeChild childElement
    
    For Each curkey In theMaterials.Keys
    
        'Keys = theMaterials.Keys
        
        Set curMaterial = theMaterials.Item(curkey)
        Set childElement = childElement.CloneNode(True)
        
        nUnit = swsUnitSystem_e.swsUnitSystemIPS
        
        E = "EX"        'elastic modulus
        nu = "NUXY"     'NUXY for poisson's ratio
        Su = "SIGXT"    'ultimate
        Sy = "SIGYLD"   'yield
        G = "GXY"       'shear modulus
        rho = "DENS"    'density
        
        With curMaterial
        
            childElement.selectNodes("//fea:materialName").Item(0).Text = .MaterialName
            childElement.selectSingleNode("fea:materialName").Text = .MaterialName
            childElement.selectSingleNode("fea:materialType").Text = .Category
            childElement.selectSingleNode("fea:materialModelType").Text = materialModelTypeText(.ModelType)
            childElement.selectSingleNode("fea:yield").Text = Format(.GetPropertyByName(nUnit, Sy, 0), "#")
            childElement.selectSingleNode("fea:tensile").Text = Format(.GetPropertyByName(nUnit, Su, 0), "#")
            childElement.selectSingleNode("fea:E").Text = Format(.GetPropertyByName(nUnit, E, 0), "#")
            childElement.selectSingleNode("fea:nu").Text = Format(.GetPropertyByName(nUnit, nu, 0), ".##")
            childElement.selectSingleNode("fea:density").Text = Format(.GetPropertyByName(nUnit, rho, 0), "#.##")
            childElement.selectSingleNode("fea:G").Text = Format(.GetPropertyByName(nUnit, G, 0), "#")
            
        End With
        
        parentElement.appendChild childElement
        
    Next
     
End Sub


Sub dataToXML(theXMLDoc As DOMDocument60, theTagName As String, theData)
    
    Dim theElement As IXMLDOMElement
    Set theElement = theXMLDoc.selectNodes("//" & theTagName).Item(0)
    
    theElement.Text = theData
    
End Sub


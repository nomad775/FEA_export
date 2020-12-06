Attribute VB_Name = "report"
Sub main()
    
    Dim swApp As SldWorks.SldWorks
    Dim COSMOSWORKS As COSMOSWORKS
    Dim COSMOSObject As Object
    Dim swVersion As Long
    Dim cwVersion As Long
    Dim cwProgID As String

    'Connect to SOLIDWORKS
    If swApp Is Nothing Then Set swApp = Application.SldWorks
 
    'Determine host SOLIDWORKS major version
    swVersion = Left(swApp.RevisionNumber, 2)

    'Calculate the version-specific ProgID of the Simulation add-in that is compatible with this version of SOLIDWORKS
    cwVersion = swVersion - 15
    cwProgID = "SldWorks.Simulation." & cwVersion
    Debug.Print (cwProgID)

    
    'Get the SOLIDWORKS Simulation object

    Set COSMOSObject = swApp.GetAddInObject(cwProgID)
    If COSMOSObject Is Nothing Then MsgBox "COSMOSObject object not found"

    Set COSMOSWORKS = COSMOSObject.COSMOSWORKS
    If COSMOSWORKS Is Nothing Then MsgBox "COSMOSWORKS object not found"
    
   'Open the active document and use the COSMOSWORKS API
    Dim theStudy As CWStudy
    Set theStudy = COSMOSWORKS.ActiveDoc.StudyManager.GetStudy(0)
    
    Debug.Print COSMOSWORKS.ActiveDoc.StudyManager.StudyCount
    
    
    Dim xml As DOMDocument60
    Set xml = New DOMDocument60
    
    xml.Load "D:\xmlTemplate.xml"
    
    dataToXML xml, "studyName", theStudy.name
    dataToXML xml, "analysisType", analysisStudyTypeText(theStudy.AnalysisType)
    dataToXML xml, "meshType", theStudy.MeshType
    
    getStudyOptions theStudy, xml
    getLoadAndRestraints theStudy, xml
    getMesh theStudy, xml
    getMaterials theStudy, xml
     
    xml.Save "D:\reportData.xml"
    
End Sub

Sub getStudyOptions(theStudy As CWStudy, theXMLDoc As DOMDocument60)
      
    Dim studyOptions As CWStaticStudyOptions
    Set studyOptions = theStudy.StaticStudyOptions
    
    With studyOptions
        dataToXML theXMLDoc, "solverType", solverTypeText(.SolverType)
        dataToXML theXMLDoc, "useInplaneEffect", onOffText(.UseInPlaneEffect)
        dataToXML theXMLDoc, "useSoftSpring", onOffText(.UseSoftSpring)
        dataToXML theXMLDoc, "useIniterialRelief", onOffText(.UseInertialRelief)
        dataToXML theXMLDoc, "incompatibleBonding", incompatibleBondingOptionText(.IncompatibleBondingOption)
        dataToXML theXMLDoc, "useLargeDisplacement", onOffText(.LargeDisplacement)
        dataToXML theXMLDoc, "computeFreeBodyForce", onOffText(.ComputeFreeBodyForce)
        dataToXML theXMLDoc, "useFriction", onOffText(.IncludeGlobalFriction)
        dataToXML theXMLDoc, "useAdaptiveMethod", adaptiveMethodText(.AdaptiveMethodType)
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
        
        Debug.Print j, landr.name,
        
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
                'addmeshcontroltoxml theXMLDoc, land
                
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
    
    Set parentElement = theXMLDoc.getElementsByTagName("restraints").Item(0)
    Set restraintElement = theXMLDoc.createElement("restraint")
    Set typeElement = theXMLDoc.createElement("type")
    Set nameElement = theXMLDoc.createElement("name")
    
    Set translationElement = theXMLDoc.createElement("translation")
    Set rotationElement = theXMLDoc.createElement("rotation")
    
    With theRestraint
        
        typeElement.Text = restraintTypeText(.RestraintType)
        nameElement.Text = .name
        
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
    
    Dim parentElement As IXMLDOMElement, loadElement As IXMLDOMElement
    
    Set parentElement = theXMLDoc.getElementsByTagName("loads").Item(0)
    Set loadElement = theXMLDoc.createElement("load")
    
    Set nameElement = theXMLDoc.createElement("loadName")
    nameElement.Text = theLoad.name
    
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
            
            Set typeElement = theXMLDoc.createElement("loadType")
            Set valueelement = theXMLDoc.createElement("loadValue")
            
            typeElement.Text = "Normal force"
            valueelement.Text = force.NormalForceOrTorqueValue
            
            loadElement.appendChild nameElement
            loadElement.appendChild typeElement
            loadElement.appendChild valueelement
            
            parentElement.appendChild loadElement
            
        Case swsForceTypeTorque
        Case swsForceTypeForceOrMoment
            
            force.GetForceComponentValues fb1, fb2, fb3, fd1, fd2, fd3
            forcex = fb1 * fd1
            forcey = fb2 * fd2
            forcez = fb3 * fd3
            
            Set typeElement = theXMLDoc.createElement("loadType")
            Set dir1element = theXMLDoc.createElement("dir1")
            Set dir2element = theXMLDoc.createElement("dir2")
            Set dir3element = theXMLDoc.createElement("dir3")
            
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
    
    Set parentElement = theXMLDoc.getElementsByTagName("loads").Item(0)
    
    Set loadElement = theXMLDoc.createElement("load")
    
    Set nameElement = theXMLDoc.createElement("loadName")
    Set typeElement = theXMLDoc.createElement("loadType")
    Set valueelement = theXMLDoc.createElement("loadValue")
    'Set transverseelement = theXMLDoc.createElement("transverseValue")
    
    Set bload = theLoad
    
    nameElement.Text = theLoad.name
    typeElement.Text = "Bearing load"
    valueelement.Text = bload.XDirectionValue
    
    loadElement.appendChild nameElement
    loadElement.appendChild typeElement
    loadElement.appendChild valueelement
    
    parentElement.appendChild loadElement
    
End Sub


'gravity.GetGravitationalAcclerationValues
'gravity.Unit
        

Sub getMesh(theStudy As CWStudy, theXMLDoc As DOMDocument60)
    
    With theStudy.mesh
        Debug.Print "element count " & .ElementCount
        Debug.Print "element size  " & .ElementSize
        Debug.Print "max size      " & .MaxElementSize
        Debug.Print "controls count" & .MeshControlCount
        Debug.Print "mesh type     " & .MesherType
        Debug.Print "min size      " & .MinElementSize
        Debug.Print "node count    " & .NodeCount
        Debug.Print "quality       " & .Quality
        Debug.Print "unit          " & .Unit
        Debug.Print "chk jac shell " & .UseJacobianCheckForShells
        Debug.Print "chk jac solid " & .UseJacobianCheckForSolids
        Debug.Print "worst jac ratio" & .GetWorstJacobianRatio
        
        dataToXML theXMLDoc, "meshType", meshTypeText(.MeshType)
        dataToXML theXMLDoc, "mesherUsed", mesherTypeText(.MesherType)
        dataToXML theXMLDoc, "elementCount", .ElementCount
        dataToXML theXMLDoc, "jacobianPoints", jacobianPointsText(.UseJacobianCheckForSolids)
        dataToXML theXMLDoc, "jacobianCheckForShell", jacobianPointsText(.UseJacobianCheckForShells)
        dataToXML theXMLDoc, "maxElementSize", .MaxElementSize
        dataToXML theXMLDoc, "minElementSize", .MinElementSize
        dataToXML theXMLDoc, "meshQuality", meshQualityText(.Quality)
        
    End With
    
End Sub

Sub getMaterials(theStudy As CWStudy, theXMLDoc As DOMDocument60)
    
    Dim theComponent As CWSolidComponent
    Dim theSolidBody As CWSolidBody
    Dim theMaterial As CWMaterial
    Dim nUnit As Long, E As String, nu As String, Su As String, Se As String
    Dim materials As Dictionary
    
    Dim errCode As Long
    
    Set theComponent = theStudy.SolidManager.GetComponentAt(0, errCode)
    Set materials = New Dictionary
    
    bodycount = theComponent.SolidBodyCount
    
    For i = 0 To bodycount - 1
        
        Set theSolidBody = theComponent.GetSolidBodyAt(i, errCode)
        Set theMaterial = theSolidBody.GetSolidBodyMaterial
        
        If Not materials.Exists(theMaterial.MaterialName) Then
            materials.add theMaterial.MaterialName, theMaterial
        End If
        
    Next
    
    materialsToXML theXMLDoc, materials
    
End Sub

Sub materialsToXML(theXMLDoc As DOMDocument60, theMaterials As Dictionary)
'output materials to XML file
    
    Dim parentElement As IXMLDOMElement
    Dim childElement As IXMLDOMElement
    Dim curMaterial As CWMaterial
    
    'get parent element and save child template
    Set parentElement = theXMLDoc.getElementsByTagName("materials").Item(0)
    Set childElement = parentElement.FirstChild
    
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
        
        With curMaterial
        
            childElement.getElementsByTagName("materialName").Item(0).Text = .MaterialName
            'materialFragment.selectSingleNode("\defaultFailureCriterion") = .xx
            
            'a = .Category
            'b = .ModelType  'swsMaterialModelType_e; 0 = Linear Elastic Isotropic
            
            'c = .Source     'swsMaterialSource_e    0Solidworks, 1 custer,2centor library,3library files
            '.Category 'text field
            '.Source     'text field
            
            childElement.selectSingleNode("materialName").Text = .MaterialName
            
            childElement.selectSingleNode("yield").Text = .GetPropertyByName(nUnit, Sy, 0)
            childElement.selectSingleNode("tensile").Text = .GetPropertyByName(nUnit, Su, 0)
            childElement.selectSingleNode("E").Text = .GetPropertyByName(nUnit, E, 0)
            childElement.selectSingleNode("nu").Text = .GetPropertyByName(nUnit, nu, 0)
            
        End With
        
        parentElement.appendChild childElement
        
    Next
     
End Sub



Function onOffText(theOption As Integer) As String
    Select Case theOption
        Case 0
            onOffText = "Off"
        Case 1
            onOffText = "On"
        Case Else
            onOffText = "Unknown"
    End Select
End Function

Function analysisStudyTypeText(theType As Integer) As String
    'incomplete list
    'swsAnalysisStudyType_e.swsAnalysisStudyTypeDropTest
    'swsAnalysisStudyType_e.swsAnalysisStudyTypeDynamic
    'swsAnalysisStudyType_e.swsAnalysisStudyTypeNonlinear
    'swsAnalysisStudyType_e.swsAnalysisStudyTypeStatic
    
    Select Case theType
        Case swsAnalysisStudyTypeStatic
            analysisStudyTypeText = "Static"
        Case swsAnalysisStudyTypeDynamic
            analysisStudyTypeText = "Dynamic"
        Case swsAnalysisStudyTypeDropTest
            analysisStudyTypeText = "Drop test"
        Case swsAnalysisStudyTypeNonlinear
            analysisStudyTypeText = "Non-linear"
    End Select
    
End Function

Function solverTypeText(theSolverType As Integer) As String
    
    Select Case theSolverType
        Case swsSolverType_e.swsSolverTypeAbaqus
            solverTypeText = "Abaqus"
        Case swsSolverType_e.swsSolverTypeAutomatic
            solverTypeText = "Automatic"
        Case swsSolverType_e.swsSolverTypeCASI
            solverTypeText = "CASI"
        Case swsSolverType_e.swsSolverTypeDirectSparse
            solverTypeText = "Direct Sparse"
        Case swsSolverType_e.swsSolverTypeFFEPlus
            solverTypeText = "FFE Plus"
        Case swsSolverType_e.swsSolverTypeINTEL
            solverTypeText = "Intel"
        Case swsSolverType_e.swsSolverTypeINTELCluster
            solverTypeText = "Intel cluster"
        Case Else
            solverTypeText = "Unknown"
    End Select
    
End Function

Function incompatibleBondingOptionText(theOption As Integer) As String
    'swsIncompatibleBondingOption_e
    'swsIncompatibleBondingOption_Automatic 0 = Automatic; let solver automatically switch from surface-based bonding contact to node-based bonding contact, if surface-based bonding contact slows down solution convergence
    'swsIncompatibleBondingOption_MoreAccurate 2 = More accurate (slower); use surface-based bonding contact method to produce continuous and more accurate stresses in contact regions
    'swsIncompatibleBondingOption_Simplified 1 = Simplified; use node-based bonding contact method on models with extensive contact surfaces to quickly reach solution convergence

    Select Case theOption
        Case swsIncompatibleBondingOption_Automatic
            incompatibleBondingOptionText = "Automatic"
        Case swsIncompatibleBondingOption_MoreAccurate
            incompatibleBondingOptionText = "More Accurate"
        Case swsIncompatibleBondingOption_Simplified
            incompatibleBondingOptionText = "Simplified"
        Case Else
            incompatibleBondingOptionText = "Unknown"
    End Select
        
End Function

Function adaptiveMethodText(theType As Integer) As String
    '0 = None
    '1 = h-adaptive; iteratively adjusts the size of the mesh cell in areas of the model where a smaller mesh is needed
    '2 = p-adaptive; iteratively adjusts the polynomial order of the mesh to improve accuracy
    
    Select Case theType
        Case 0
            adaptiveMethodText = "None"
        Case 1
            adaptiveMethodText = "H-adaptive"
        Case 2
            adaptiveMethodText = "P-adaptive"
    End Select
        
End Function

Function loadAndRestraintTypeText(theType As Integer) As String
    'swsLoadsAndRestraintsType_e
    'swsLoadsAndRestraintsType_e.swsLoadsAndRestraintsTypeBearingLoads
    'swsLoadsAndRestraintsType_e.swsLoadsAndRestraintsTypeConnectors
    'swsLoadsAndRestraintsType_e.swsLoadsAndRestraintsTypeForce
    'swsLoadsAndRestraintsType_e.swsLoadsAndRestraintsTypeGravity
    'swsLoadsAndRestraintsType_e.swsLoadsAndRestraintsTypeMeshControl
    'swsLoadsAndRestraintsType_e.swsLoadsAndRestraintsTypePressure
    'swsLoadsAndRestraintsType_e.swsLoadsAndRestraintsTypeRemoteLoad
    'swsLoadsAndRestraintsType_e.swsLoadsAndRestraintsTypeRemoteMass
    'swsLoadsAndRestraintsType_e.swsLoadsAndRestraintsTypeRestraint
    
    Select Case theType
        Case swsLoadsAndRestraintsTypeBearingLoads
            loadAndRestraintTypeText = "Bearing load"
        Case swsLoadsAndRestraintsTypeConnectors
            loadAndRestraintTypeText = "Connector"
        Case swsLoadsAndRestraintsTypeForce
            loadAndRestraintTypeText = "Force"
        Case swsLoadsAndRestraintsTypeGravity
            loadAndRestraintTypeText = "Gravity"
        Case swsLoadsAndRestraintsTypeMeshControl
            loadAndRestraintTypeText = "Mesh control"
        Case swsLoadsAndRestraintsTypePressure
            loadAndRestraintTypeText = "Pressure"
        Case swsLoadsAndRestraintsTypeRemoteLoad
            loadAndRestraintTypeText = "Remote load"
        Case swsLoadsAndRestraintsTypeRemoteMass
            loadAndRestraintTypeText = "Remote mass"
        Case swsLoadsAndRestraintsTypeRestraint
            loadAndRestraintTypeText = "Restraint"
    End Select
    
End Function

Function restraintTypeText(theType As Integer) As String
    Select Case theType
        Case swsRestraintType_e.swsRestraintTypeCyclicSymmetry
            restraintTypeText = "Cyclic symmetry"
        Case swsRestraintType_e.swsRestraintTypeCylindricalFaces
            restraintTypeText = "Cylindrical faces"
        Case swsRestraintType_e.swsRestraintTypeFixed
            restraintTypeText = "Fixed"
        Case swsRestraintType_e.swsRestraintTypeFlatFace
            restraintTypeText = "Flat faces"
        Case swsRestraintType_e.swsRestraintTypeHinge
            restraintTypeText = "Hinge"
        Case swsRestraintType_e.swsRestraintTypeImmovable
            restranttypetext = "Immovable"
        Case swsRestraintType_e.swsRestraintTypeReferenceGeometry
            restraintTypeText = "Reference geometry"
        Case swsRestraintType_e.swsRestraintTypeRoller
            RestraintType = "Roller"
        Case swsRestraintType_e.swsRestraintTypeSphericalSurface
            restraintTypeText = "Spherical surface"
        Case swsRestraintType_e.swsRestraintTypeSymmetric
            restraintTypeText = "Symmetric"
    End Select
End Function

Function mesherTypeText(theType As Integer) As String

    'mesh type swsMesherType_e
    'swsMesherType_e.swsMesherTypeAlternate = 1     curvature-based
    'swsMesherType_e.swsMesherTypeAlternateCB = 2   blended curvature-based
    'swsMesherType_e.swsMesherTypeStandard = 0      standard
    
    Select Case theType
        Case swsMesherTypeAlternate
            mesherTypeText = "Curvature based"
        Case swsMesherTypeAlternateCB
            mesherTypeText = "Blended curvature-based"
        Case swsMesherTypeStandard
            mesherTypeText = "Standard"
    End Select
End Function

Function meshTypeText(theType As Integer) As String
    'swsMeshType_e.swsMeshTypeBeam
    'swsMeshType_e.swsMeshTypeMidSurface
    'swsMeshType_e.swsMeshTypeMixed
    'swsMeshType_e.swsMeshTypeSolid
    'swsMeshType_e.swsMeshTypeSurfaces
    
    Select Case theType
        Case swsMeshTypeBeam
            meshTypeText = "Beam"
        Case swsMeshTypeMidSurface
            meshTypeText = "Mid-surface"
        Case swsMeshTypeMixed
            meshTypeText = "Mixed"
        Case swsMeshTypeSolid
            meshTypeText = "Solid"
        Case swsMeshTypeSurfaces
            meshTypeText = "Surfaces"
    End Select
    
End Function

Function meshQualityText(theType As Integer) As String
    'quality
    'swsMeshQuality_e.swsMeshQualityDraft
    'swsMeshQuality_e.swsMeshQualityHigh
    
    Select Case theType
        Case swsMeshQualityDraft
            meshQualityText = "Draft"
        Case swsMeshQualityHigh
            meshQualityText = "High"
    End Select
        
End Function

Function linearUnitText(theType As Integer) As String
    'unit
    'swsLinearUnit_e.swsLinearUnitInches = 3
    'swsLinearUnit_e.swsLinearUnitFeet = 4
    'swsLinearUnit_e.swsLinearUnitCentimeters = 1
    'swsLinearUnit_e.swsLinearUnitMeters = 2
    'swsLinearUnit_e.swsLinearUnitMillimeters = 0
    
    Select Case theType
        Case swsLinearUnitInches
            linearUnitText = "Inches"
        Case swsLinearUnitFeet
            linearUnitText = "Feet"
        Case swsLinearUnitCentimeters
            linearUnitText = "Centimeters"
        Case swsLinearUnitMeters
            linearUnitText = "Meters"
        Case swsLinearUnitMillimeters
            linearunitext = "Millimeters"
    End Select
    
End Function

Function jacobianPointsText(theType As Integer) As String

    'useJacobianCheckForSolids has 4 options
    '1 =  4 points
    '2 = 16 points
    '3 = 29 points
    '4 = at nodes
    
    Select Case theType
        Case 0
            jacobianPointsText = "Off"
        Case 1
            jacobianPointsText = "4 points"
        Case 2
            jacobianPointsText = "16 points"
        Case 3
            jacobianPointsText = "29 points"
        Case 4
            jacobianPointsText = "At Nodes"
    End Select
    
End Function


Sub dataToXML(theXMLDoc As DOMDocument60, theTagName As String, theData)
    
    Dim theElement As IXMLDOMElement
    Set theElement = theXMLDoc.getElementsByTagName(theTagName).Item(0)
    theElement.Text = theData
    
End Sub

Sub z_restraintValue()
 
                'Set r1 = landr
            
                'If r1.RestraintType = swsRestraintType_e.swsRestraintTypeFixed Then
                    
                '    Debug.Print " (fixed restraint)"
                    
                    'Dim b1 As Long, b2 As Long, b3 As Long
                    'Dim d1 As Double, d2 As Double, d3 As Double
                    
                    'r1.GetTranslationComponentsValues b1, b2, b3, d1, d2, d3
                    'Debug.Print "restaints", d1, d2, d3
                    
                    'Dim errCode As Long
                    
                    'swSelectType_e.swSelPLANESECTIONS
                    'swSelectType_e.swSelFACES
                    
               '     Set theFace = landr.GetEntityAt(0, swSelFACES)
                    
                    'theValue = theStudy.Results.GetReactionForcesAndMomentsWithSelections(1, Nothing, swsForceUnitlbOrlbin, theFace, total, eachobject, errCode)
                    
                    'Debug.Print "TEST REACTION", total(0), total(1), total(2), total(3), total(4), total(5)
                    
                'End If
                
End Sub

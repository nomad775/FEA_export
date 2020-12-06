Attribute VB_Name = "enumToText"
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
        Case Else
            loadAndRestraintTypeText = "Not found"
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
            restraintTypeText = "Immovable"
        Case swsRestraintType_e.swsRestraintTypeReferenceGeometry
            restraintTypeText = "Reference geometry"
        Case swsRestraintType_e.swsRestraintTypeRoller
            restraintTypeText = "Roller"
        Case swsRestraintType_e.swsRestraintTypeSphericalSurface
            restraintTypeText = "Spherical surface"
        Case swsRestraintType_e.swsRestraintTypeSymmetric
            restraintTypeText = "Symmetric"
        Case Else
            x = 0
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

Function materialModelTypeText(theType As String) As String
'swsMaterialModelType_e
'swsMaterialModelType_e.swsMaterialModelTypeLinearElasticAnisotropic
'swsMaterialModelType_e.swsMaterialModelTypeLinearElasticIsotropic
'swsMaterialModelType_e.swsMaterialModelTypeLinearElasticOrthtropic
'swsMaterialModelType_e.swsMaterialModelTypeNonlinearElastic

    Select Case theType
        
        Case swsMaterialModelTypeLinearElasticIsotropic
            materialModelTypeText = "Linear isotropic"
            
        Case swsMaterialModelTypeLinearElasticOrthtropic
            materialModelTypeText = "Linear orthropic"
            
        Case swsMaterialModelTypeLinearElasticAnisotropic
            materialModelTypeText = "Linear Anisotropic"
            
        Case swsMaterialModelTypeNonlinearElastic
            materialModelTypeText = "Non-Linear elastic"
            
        Case Else
            materialModelTypeText = "Other non-linear"
            
    End Select
    
End Function

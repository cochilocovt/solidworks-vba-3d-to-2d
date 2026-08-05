Option Explicit

' R23 typed record for one resolved model feature.
'
' This class stores what a feature IS, never where its instances sit on the
' drawing. Physical placement belongs to CPhysicalHoleLocation and drawing
' placement to CViewHoleProjection.
'
' Every semantic manufacturing field is paired with a proof-source string.
' A field whose proof source is empty has not been established and must not
' be published as manufacturing intent.

Public ModelFeature As SldWorks.Feature
Public SeedFeature As SldWorks.Feature
Public DefinitionObject As Object
Public OwnedFaces As Collection

' Diagnostics only. Feature names must never take part in identity keys.
Public FeatureName As String
Public SeedFeatureName As String

' Raw and normalized type evidence. RawTypeName2 of "ICE" is an Instant3D
' wrapper and must be resolved through RawTypeName.
Public RawTypeName2 As String
Public RawTypeName As String
Public EffectiveType As String
Public TypeResolutionSource As String

Public ConfigurationName As String
Public SuppressionState As String
Public SuppressionProven As Boolean

Public DefinitionInterfaceName As String
Public DefinitionReadStatus As String
Public SelectionAccessGranted As Boolean
Public SelectionAccessReleased As Boolean

Public OperationKind As String

Public NominalDiameterM As Double
Public DiameterProofSource As String

Public DepthM As Double
Public DepthProofSource As String

Public EndConditionCode As Long
Public EndConditionProofSource As String

Public IsBossFeature As Boolean
Public BossProofSource As String

Public CounterBoreDiameterM As Double
Public CounterBoreDepthM As Double
Public CounterBoreProofSource As String

Public ThreadDescription As String
Public ThreadDepthM As Double
Public ThreadProofSource As String

Public FitDescription As String
Public FitProofSource As String

Public ToleranceMinM As Double
Public ToleranceMaxM As Double
Public ToleranceProofSource As String

Public Accepted As Boolean
Public RejectionReason As String

Private Sub Class_Initialize()
    Set OwnedFaces = New Collection

    RawTypeName2 = vbNullString
    RawTypeName = vbNullString
    EffectiveType = vbNullString
    TypeResolutionSource = "Unresolved"
    SuppressionState = "Unproven"
    SuppressionProven = False
    DefinitionInterfaceName = vbNullString
    DefinitionReadStatus = "NotRead"
    OperationKind = "Unclassified"
    EndConditionCode = -1
    Accepted = False
    RejectionReason = "NotEvaluated"
End Sub

' Identity of the feature for graph indexing. Composed of the effective type
' plus both raw type readings and the diagnostic name, because the installed
' VBA host reuses COM wrapper addresses and ObjPtr alone is not a safe key.
Public Function FeatureIdentityKey() As String
    FeatureIdentityKey = _
        "type=" & Module11_GeometryIdentity.IdentityToken(EffectiveType) & _
        "|raw2=" & Module11_GeometryIdentity.IdentityToken(RawTypeName2) & _
        "|raw1=" & Module11_GeometryIdentity.IdentityToken(RawTypeName) & _
        "|name=" & Module11_GeometryIdentity.IdentityToken(FeatureName)
End Function

' The semantic family this feature contributes to. Built only from typed
' manufacturing values, so two identically defined features in different
' tree positions belong to one family.
Public Function SemanticFamilyKey() As String
    SemanticFamilyKey = Module11_GeometryIdentity.BuildSemanticFamilyKey( _
        OperationKind, NominalDiameterM, DepthM, _
        CounterBoreDiameterM, CounterBoreDepthM, ThreadDescription)
End Function

Public Function HasOwnedFaces() As Boolean
    If OwnedFaces Is Nothing Then Exit Function
    HasOwnedFaces = (OwnedFaces.Count > 0)
End Function

' Lists the semantic fields that carry a value but no proof source. QA must
' treat a non-empty result as a manufacturing-definition failure rather than
' publishing the value.
Public Function UnprovenSemanticFields() As String
    Dim missing As String

    missing = AppendUnproven(missing, "Diameter", _
        NominalDiameterM <> 0#, DiameterProofSource)
    missing = AppendUnproven(missing, "Depth", _
        DepthM <> 0#, DepthProofSource)
    missing = AppendUnproven(missing, "EndCondition", _
        EndConditionCode >= 0, EndConditionProofSource)
    missing = AppendUnproven(missing, "CounterBoreDiameter", _
        CounterBoreDiameterM <> 0#, CounterBoreProofSource)
    missing = AppendUnproven(missing, "CounterBoreDepth", _
        CounterBoreDepthM <> 0#, CounterBoreProofSource)
    missing = AppendUnproven(missing, "Thread", _
        Len(Trim$(ThreadDescription)) > 0, ThreadProofSource)
    missing = AppendUnproven(missing, "Fit", _
        Len(Trim$(FitDescription)) > 0, FitProofSource)
    missing = AppendUnproven(missing, "Tolerance", _
        (ToleranceMinM <> 0#) Or (ToleranceMaxM <> 0#), _
        ToleranceProofSource)

    UnprovenSemanticFields = missing
End Function

Private Function AppendUnproven( _
    ByVal current As String, _
    ByVal fieldName As String, _
    ByVal fieldHasValue As Boolean, _
    ByVal proofSource As String) As String

    AppendUnproven = current

    If Not fieldHasValue Then Exit Function
    If Len(Trim$(proofSource)) > 0 Then Exit Function

    If Len(current) > 0 Then
        AppendUnproven = current & "," & fieldName
    Else
        AppendUnproven = fieldName
    End If
End Function

Public Function SemanticSummary() As String
    SemanticSummary = _
        "operation=" & OperationKind & _
        "|effectiveType=" & EffectiveType & _
        "|diameterM=" & Format$(NominalDiameterM, "0.000000000") & _
        "|depthM=" & Format$(DepthM, "0.000000000") & _
        "|cboreDiameterM=" & _
            Format$(CounterBoreDiameterM, "0.000000000") & _
        "|cboreDepthM=" & Format$(CounterBoreDepthM, "0.000000000") & _
        "|thread=" & ThreadDescription & _
        "|fit=" & FitDescription & _
        "|unprovenFields=" & UnprovenSemanticFields()
End Function

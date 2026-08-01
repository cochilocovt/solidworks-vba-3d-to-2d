Option Explicit

' R23 typed record for one physical manufacturing location.
'
' A physical location is an infinite axis line plus the axial interval that
' the material actually occupies. Three invariants follow, and every one of
' them is enforced by CanConsolidateWith:
'
'   1. Coaxial cylinders created by SEPARATE features, for example a
'      counterbore and its through hole, or two independent Cut-Extrudes
'      forming a stepped bore, consolidate into ONE location when their
'      axial intervals meet.
'   2. Opposite blind holes sharing ONE infinite axis stay SEPARATE, because
'      their axial intervals do not meet.
'   3. Feature names never take part in identity. Renaming a feature cannot
'      change, split or merge a physical location.

Public ConfigurationName As String
Public ComponentName As String
Public ModelPathName As String

' Sign-normalized unit direction and its line moment. Together these
' identify the infinite line independently of the sampled origin.
Public AxisX As Double
Public AxisY As Double
Public AxisZ As Double
Public MomentX As Double
Public MomentY As Double
Public MomentZ As Double

' Axial extent of the consolidated stack along the normalized axis.
Public AxialMinM As Double
Public AxialMaxM As Double

Public OpeningSide As String

' Coaxial members of this location. StackFeatures holds CFeatureDefinition
' references; StackRadiiM holds the matching cylinder radii.
Public StackFeatures As Collection
Public StackRadiiM As Collection

' The model cylindrical faces that produced this location, retained so a
' drawing projection can reach their boundary edges. Phase 0 proved mapping
' is per-EDGE: IView.GetCorrespondingEntity maps only one of a counterbore's
' two circular edges, so a projection must be free to try every edge of every
' contributing face before failing the location.
Public SourceFaces As Collection

Public PrimaryRadiusM As Double
Public MaximumRadiusM As Double

Public AggregateDefinitionSummary As String
Public LineKey As String
Public PhysicalInstanceKey As String
Public IndexedPhysicalKey As String
Public SemanticFamilyKey As String

Public IdentityProven As Boolean
Public RejectionReason As String

Private mIntervalInitialized As Boolean

Private Sub Class_Initialize()
    Set StackFeatures = New Collection
    Set StackRadiiM = New Collection
    Set SourceFaces = New Collection

    OpeningSide = "Unproven"
    LineKey = vbNullString
    PhysicalInstanceKey = vbNullString
    IndexedPhysicalKey = vbNullString
    SemanticFamilyKey = vbNullString
    IdentityProven = False
    RejectionReason = "NotEvaluated"
    mIntervalInitialized = False
End Sub

' Establishes the infinite-line identity from a sampled axis point and
' direction. Must succeed before the location can be consolidated or keyed.
Public Function SetAxisFromSample( _
    ByVal pointX As Double, _
    ByVal pointY As Double, _
    ByVal pointZ As Double, _
    ByVal directionX As Double, _
    ByVal directionY As Double, _
    ByVal directionZ As Double) As Boolean

    If Not Module11_GeometryIdentity.NormalizeAxisAndMoment( _
        pointX, pointY, pointZ, _
        directionX, directionY, directionZ, _
        AxisX, AxisY, AxisZ, _
        MomentX, MomentY, MomentZ) Then

        RejectionReason = "DegenerateAxisDirection"
        Exit Function
    End If

    LineKey = Module11_GeometryIdentity.BuildAxisLineKey( _
        AxisX, AxisY, AxisZ, MomentX, MomentY, MomentZ)

    IdentityProven = True
    RejectionReason = vbNullString
    SetAxisFromSample = True
End Function

' Extends the axial interval to include the supplied span.
Public Sub ExtendAxialInterval( _
    ByVal firstValueM As Double, _
    ByVal secondValueM As Double)

    Dim lowValue As Double
    Dim highValue As Double

    If firstValueM <= secondValueM Then
        lowValue = firstValueM
        highValue = secondValueM
    Else
        lowValue = secondValueM
        highValue = firstValueM
    End If

    If Not mIntervalInitialized Then
        AxialMinM = lowValue
        AxialMaxM = highValue
        mIntervalInitialized = True
    Else
        If lowValue < AxialMinM Then AxialMinM = lowValue
        If highValue > AxialMaxM Then AxialMaxM = highValue
    End If
End Sub

Public Function HasAxialInterval() As Boolean
    HasAxialInterval = mIntervalInitialized
End Function

Public Sub AddStackMember( _
    ByRef featureDefinition As CFeatureDefinition, _
    ByVal radiusM As Double, _
    ByVal axialMinM As Double, _
    ByVal axialMaxM As Double)

    If Not featureDefinition Is Nothing Then
        StackFeatures.Add featureDefinition
    End If

    StackRadiiM.Add Abs(radiusM)
    ExtendAxialInterval axialMinM, axialMaxM
    RefreshRadii
End Sub

Private Sub RefreshRadii()
    Dim i As Long
    Dim currentRadius As Double

    PrimaryRadiusM = 0#
    MaximumRadiusM = 0#

    For i = 1 To StackRadiiM.Count
        currentRadius = Abs(CDbl(StackRadiiM(i)))

        If i = 1 Then
            PrimaryRadiusM = currentRadius
            MaximumRadiusM = currentRadius
        Else
            ' The functional size of a counterbored or stepped hole is its
            ' smallest coaxial cylinder; the largest is the entry feature.
            If currentRadius < PrimaryRadiusM Then
                PrimaryRadiusM = currentRadius
            End If
            If currentRadius > MaximumRadiusM Then
                MaximumRadiusM = currentRadius
            End If
        End If
    Next i
End Sub

' Same infinite line, same configuration and same component.
Public Function IsCoaxialWith( _
    ByRef other As CPhysicalHoleLocation) As Boolean

    If other Is Nothing Then Exit Function
    If Not IdentityProven Then Exit Function
    If Not other.IdentityProven Then Exit Function

    If StrComp(LineKey, other.LineKey, vbBinaryCompare) <> 0 Then Exit Function

    If StrComp(Module11_GeometryIdentity.IdentityToken(ConfigurationName), _
        Module11_GeometryIdentity.IdentityToken(other.ConfigurationName), _
        vbBinaryCompare) <> 0 Then Exit Function

    If StrComp(Module11_GeometryIdentity.IdentityToken(ComponentName), _
        Module11_GeometryIdentity.IdentityToken(other.ComponentName), _
        vbBinaryCompare) <> 0 Then Exit Function

    IsCoaxialWith = True
End Function

Public Function AxialIntervalMeets( _
    ByRef other As CPhysicalHoleLocation) As Boolean

    If other Is Nothing Then Exit Function
    If Not mIntervalInitialized Then Exit Function
    If Not other.HasAxialInterval() Then Exit Function

    AxialIntervalMeets = _
        Module11_GeometryIdentity.AxialIntervalsOverlap( _
            AxialMinM, AxialMaxM, other.AxialMinM, other.AxialMaxM)
End Function

' The consolidation decision. Coaxial alone is not sufficient: the axial
' intervals must also meet, which is what keeps opposite blind holes on one
' infinite line separate.
Public Function CanConsolidateWith( _
    ByRef other As CPhysicalHoleLocation) As Boolean

    If Not IsCoaxialWith(other) Then Exit Function
    If Not AxialIntervalMeets(other) Then Exit Function

    CanConsolidateWith = True
End Function

Public Sub AddSourceFace(ByVal modelFace As Object)
    If modelFace Is Nothing Then Exit Sub
    SourceFaces.Add modelFace
End Sub

Public Sub MergeFrom(ByRef other As CPhysicalHoleLocation)
    If other Is Nothing Then Exit Sub

    Dim i As Long

    For i = 1 To other.StackFeatures.Count
        StackFeatures.Add other.StackFeatures(i)
    Next i

    For i = 1 To other.StackRadiiM.Count
        StackRadiiM.Add Abs(CDbl(other.StackRadiiM(i)))
    Next i

    ' Consolidation must not lose reachable geometry: a stepped bore's
    ' projection may only be mappable through the merged member's edges.
    For i = 1 To other.SourceFaces.Count
        SourceFaces.Add other.SourceFaces(i)
    Next i

    If other.HasAxialInterval() Then
        ExtendAxialInterval other.AxialMinM, other.AxialMaxM
    End If

    If Len(Trim$(OpeningSide)) = 0 Or OpeningSide = "Unproven" Then
        OpeningSide = other.OpeningSide
    End If

    RefreshRadii
    RecomputeKeys
End Sub

' Rebuilds both identity keys from current state. The physical key contains
' no feature name by construction.
Public Sub RecomputeKeys()
    If Not IdentityProven Then Exit Sub

    PhysicalInstanceKey = _
        Module11_GeometryIdentity.BuildPhysicalLocationKey( _
            ConfigurationName, ComponentName, LineKey, _
            Module11_GeometryIdentity.CanonicalIntervalToken( _
                AxialMinM, AxialMaxM))

    SemanticFamilyKey = BuildFamilyKeyFromStack()
    AggregateDefinitionSummary = BuildAggregateSummary()
End Sub

' Family identity comes from the first accepted stack member's typed
' definition, so that six identically defined counterbores share one family.
Private Function BuildFamilyKeyFromStack() As String
    Dim i As Long
    Dim member As CFeatureDefinition

    For i = 1 To StackFeatures.Count
        Set member = StackFeatures(i)
        If Not member Is Nothing Then
            BuildFamilyKeyFromStack = member.SemanticFamilyKey()
            Exit Function
        End If
    Next i

    BuildFamilyKeyFromStack = _
        "op=UNCLASSIFIED|radius=" & _
        Module11_GeometryIdentity.CanonicalRadiusToken(PrimaryRadiusM)
End Function

Private Function BuildAggregateSummary() As String
    Dim radiiSummary As String
    Dim i As Long

    For i = 1 To StackRadiiM.Count
        If Len(radiiSummary) > 0 Then radiiSummary = radiiSummary & "/"
        radiiSummary = radiiSummary & _
            Module11_GeometryIdentity.CanonicalRadiusToken( _
                CDbl(StackRadiiM(i)))
    Next i

    BuildAggregateSummary = _
        "stackMembers=" & CStr(StackRadiiM.Count) & _
        "|radiiM=" & radiiSummary & _
        "|primaryRadiusM=" & _
            Module11_GeometryIdentity.CanonicalRadiusToken(PrimaryRadiusM) & _
        "|maximumRadiusM=" & _
            Module11_GeometryIdentity.CanonicalRadiusToken(MaximumRadiusM) & _
        "|axialIntervalM=" & _
            Module11_GeometryIdentity.CanonicalIntervalToken( _
                AxialMinM, AxialMaxM) & _
        "|openingSide=" & OpeningSide
End Function

Public Function StackMemberCount() As Long
    StackMemberCount = StackRadiiM.Count
End Function

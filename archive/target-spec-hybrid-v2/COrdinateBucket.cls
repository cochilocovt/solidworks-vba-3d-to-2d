Option Explicit

' R23 Phase 5 typed record for ONE ordinate coordinate in one direction of
' one view.
' R23-505 exists because a bucket is not a hole. Two coaxial holes viewed
' along their shared axis produce exactly one drawing entity at exactly one
' page coordinate, which Phase 3 proved live and recorded as
' OneDrawingEntityForTwoCoaxialHoles. Dimensioning that coordinate twice
' would be a duplicate; crediting only one of the holes would silently drop
' the other from coverage. A bucket therefore holds ONE selectable entity
' and the LIST of physical locations it represents, and coverage is counted
' over locations while dimensioning is counted over buckets.

Public CoordinateKey As String
Public Direction As Long
Public DirectionName As String

' The page coordinate this bucket sits at, in the Page frame Phase 3
' established. Both are kept even though only one is ordinated, because the
' unused one is what makes a bucket identifiable in evidence.
Public PageX As Double
Public PageY As Double

' The single drawing entity that will be appended to the ordinate group.
Public Anchor As SldWorks.Entity
Public AnchorSourceKey As String
Public AnchorProof As String

Public RepresentedLocationKeys As Collection
Public RepresentedProjections As Collection

Public IsDatum As Boolean

Public Created As Boolean
Public CreatedDimension As SldWorks.DisplayDimension
Public CreationProof As String

Private Sub Class_Initialize()
    Set RepresentedLocationKeys = New Collection
    Set RepresentedProjections = New Collection
    Direction = 0
    IsDatum = False
    Created = False
    CreationProof = "NotAttempted"
End Sub

' Adds a physical location to this bucket's credit list. Duplicate keys are
' ignored so a location cannot be counted twice by being reached along two
' routes.
Public Sub CreditLocation( _
    ByVal physicalKey As String, _
    ByRef projection As CViewHoleProjection)

    If Len(Trim$(physicalKey)) = 0 Then Exit Sub
    If AlreadyCredited(physicalKey) Then Exit Sub

    RepresentedLocationKeys.Add physicalKey
    If Not projection Is Nothing Then
        RepresentedProjections.Add projection
        RecordSchemeDirectionalCoverage projection
    End If
End Sub

' A bucket represents one deliberate ordinate coordinate. Credit the
' projection in this bucket's direction as planned-scheme evidence. Actual
' coverage is only recorded by MarkCreatedDirectionalCoverage after a
' successful ordinate transaction and read-back.
Private Sub RecordSchemeDirectionalCoverage( _
    ByRef projection As CViewHoleProjection)

    If projection Is Nothing Then Exit Sub

    If Direction = Module15_OrdinateScheme.ORD_HORIZONTAL Then
        projection.SchemeCoveredX = True
        projection.SchemeCoverageSourceX = _
            "OrdinateBucket:" & CoordinateKey
    ElseIf Direction = Module15_OrdinateScheme.ORD_VERTICAL Then
        projection.SchemeCoveredY = True
        projection.SchemeCoverageSourceY = _
            "OrdinateBucket:" & CoordinateKey
    End If
End Sub

' Records actual coverage only after the caller has proved that SOLIDWORKS
' created an ordinate and IView.GetDisplayDimensions read it back.
Public Sub MarkCreatedDirectionalCoverage(ByVal proof As String)
    Dim i As Long

    Created = True
    CreationProof = proof

    For i = 1 To RepresentedProjections.Count
        Dim projection As CViewHoleProjection
        Set projection = RepresentedProjections(i)
        If projection Is Nothing Then GoTo ContinueProjection

        If Direction = Module15_OrdinateScheme.ORD_HORIZONTAL Then
            projection.CoveredX = True
            projection.CoverageSourceX = _
                "CreatedOrdinateBucket:" & CoordinateKey
        ElseIf Direction = Module15_OrdinateScheme.ORD_VERTICAL Then
            projection.CoveredY = True
            projection.CoverageSourceY = _
                "CreatedOrdinateBucket:" & CoordinateKey
        End If

ContinueProjection:
    Next i
End Sub

Public Function AlreadyCredited(ByVal physicalKey As String) As Boolean
    Dim i As Long
    For i = 1 To RepresentedLocationKeys.Count
        If StrComp(CStr(RepresentedLocationKeys(i)), physicalKey, _
            vbBinaryCompare) = 0 Then

            AlreadyCredited = True
            Exit Function
        End If
    Next i
End Function

Public Function CreditedCount() As Long
    CreditedCount = RepresentedLocationKeys.Count
End Function

Public Function HasSelectableAnchor() As Boolean
    HasSelectableAnchor = (Not Anchor Is Nothing)
End Function

Public Function CreditedKeyList() As String
    Dim result As String
    Dim i As Long

    For i = 1 To RepresentedLocationKeys.Count
        If Len(result) > 0 Then result = result & ";"
        result = result & CStr(RepresentedLocationKeys(i))
    Next i

    CreditedKeyList = result
End Function

Public Function Summary() As String
    Summary = "coordinateKey=" & CoordinateKey & _
        "|direction=" & DirectionName & _
        "|pageX=" & Format$(PageX, "0.000000000") & _
        "|pageY=" & Format$(PageY, "0.000000000") & _
        "|isDatum=" & CStr(IsDatum) & _
        "|anchor=" & CStr(HasSelectableAnchor()) & _
        "|creditedLocations=" & CStr(CreditedCount()) & _
        "|created=" & CStr(Created)
End Function

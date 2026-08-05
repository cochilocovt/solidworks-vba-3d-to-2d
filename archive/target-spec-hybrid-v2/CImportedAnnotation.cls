Option Explicit

' R23 typed record for one imported model annotation.
'
' The import transaction returns a count; a count is not coverage. This
' record carries the identity and provenance needed to decide whether an
' imported item actually satisfies a required semantic key, and to keep a
' reference-supplied tolerance distinguishable from a model-supplied one.

Public DisplayDimension As SldWorks.DisplayDimension
Public Annotation As SldWorks.Annotation
Public Dimension As SldWorks.Dimension

' SourceIdentity is the model-side identity, for example
' "D1@Sketch4@P-0251-14A-001.Part". SelectionName is the drawing-side name.
Public SourceIdentity As String
Public SelectionName As String
Public OwnerViewName As String

' Where the value came from. "Model" for imported model data, or
' "TargetSpecification" for an approved reference-supplied value. Production
' must never record a reference-supplied tolerance as model provenance.
Public ProvenanceSource As String

' Live-proven note: imported section diameters arrive as
' swDiameterDimension = 6, not swDiametricLinearDimension = 15. Created
' ordinates were observed as Type2 1 and 7, so ordinate classification must
' accept 1, 7 and 8.
Public DimensionType2 As Long
Public IsHoleCallout As Boolean
Public IsReference As Boolean

Public NominalAvailable As Boolean
Public NominalM As Double

' Which member supplied the nominal, and what the alternative returned. A
' drawing document has no configurations, so a configuration-scoped read
' returns 0 for every drawing dimension; keeping both readings visible stops
' that being mistaken for a genuine zero.
Public NominalSource As String

Public ToleranceType As Long
Public FitType As Long
Public FitDisplayStyle As Long
Public HoleFit As String
Public ShaftFit As String
Public FitValues As String
Public ToleranceMinM As Double
Public ToleranceMaxM As Double
Public ToleranceMinStatus As Long
Public ToleranceMaxStatus As Long
Public ToleranceReadStatus As String

' How much authority this tolerance carries. Standing instruction from the
' user (2026-08-01): tolerances in the designers' existing drawings were
' added manually to signal that some tolerance is acceptable, NOT to state
' that the part holds them. Reading one proves a designer typed it, nothing
' more. R23 records them as evidence and must never promote them to
' manufacturing requirements or use them to classify a feature. The rule for
' deciding when a tolerance SHOULD be added is still open.
Public ToleranceAuthority As String

Public AttachedEntityCount As Long
Public AttachmentFingerprint As String

Public PageX As Double
Public PageY As Double
Public PositionReadStatus As String

' Requirement key this annotation was reconciled against, when any.
Public CoversRequirementKey As String
Public ReconciliationStatus As String

Private Sub Class_Initialize()
    ProvenanceSource = "Unproven"
    DimensionType2 = -1
    NominalAvailable = False
    NominalSource = "NotRead"
    ToleranceType = -1
    FitType = -1
    FitDisplayStyle = -1
    ToleranceMinStatus = -1
    ToleranceMaxStatus = -1
    ToleranceReadStatus = "NotRead"
    ToleranceAuthority = "Unevaluated"
    AttachedEntityCount = 0
    AttachmentFingerprint = "None"
    PositionReadStatus = "NotRead"
    CoversRequirementKey = vbNullString
    ReconciliationStatus = "Unreconciled"
End Sub

Public Function AnnotationIdentityKey() As String
    AnnotationIdentityKey = _
        "source=" & _
            Module11_GeometryIdentity.IdentityToken(SourceIdentity) & _
        "|view=" & Module11_GeometryIdentity.IdentityToken(OwnerViewName)
End Function

Public Function HasAttachments() As Boolean
    HasAttachments = (AttachedEntityCount > 0)
End Function

' True only when the annotation carries a real nonzero tolerance. An empty
' tolerance block is not evidence of a fit.
Public Function HasNonZeroTolerance() As Boolean
    If ToleranceType <= 0 Then Exit Function

    HasNonZeroTolerance = _
        (Not Module11_GeometryIdentity.ToleranceValuesMatch( _
            ToleranceMinM, 0#)) Or _
        (Not Module11_GeometryIdentity.ToleranceValuesMatch( _
            ToleranceMaxM, 0#))
End Function

Public Function HasFitData() As Boolean
    HasFitData = _
        (Len(Trim$(HoleFit)) > 0) Or _
        (Len(Trim$(ShaftFit)) > 0) Or _
        (Len(Trim$(FitValues)) > 0)
End Function

Public Function NominalMatches(ByVal expectedM As Double) As Boolean
    If Not NominalAvailable Then Exit Function

    NominalMatches = Module11_GeometryIdentity.ValuesMatchWithin( _
        NominalM, expectedM, Module11_GeometryIdentity.TOLERANCE_MATCH_M)
End Function

Public Function ProvenanceSummary() As String
    ProvenanceSummary = _
        "source=" & SourceIdentity & _
        "|selection=" & SelectionName & _
        "|view=" & OwnerViewName & _
        "|provenance=" & ProvenanceSource & _
        "|type2=" & CStr(DimensionType2) & _
        "|holeCallout=" & CStr(IsHoleCallout) & _
        "|reference=" & CStr(IsReference) & _
        "|nominalAvailable=" & CStr(NominalAvailable) & _
        "|nominalM=" & Format$(NominalM, "0.000000000") & _
        "|toleranceType=" & CStr(ToleranceType) & _
        "|nonZeroTolerance=" & CStr(HasNonZeroTolerance()) & _
        "|fitData=" & CStr(HasFitData()) & _
        "|attachments=" & CStr(AttachedEntityCount) & _
        "|reconciliation=" & ReconciliationStatus
End Function

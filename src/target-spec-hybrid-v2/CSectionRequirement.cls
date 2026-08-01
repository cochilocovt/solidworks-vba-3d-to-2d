Option Explicit

' R23 Phase 8 typed record for ONE section-dimension requirement.
'
' A requirement is a semantic statement about the part - "the overall
' thickness is 18.00" - not a drawing object. It is reconciled against the
' dimensions already imported into the section view first, and only creates
' something when nothing already satisfies it.
'
' The fields divide into three groups, and keeping them apart is the point
' of the record:
'
'   REQUIRED    what the target specification demands
'   OBSERVED    what was actually read back from the drawing
'   VERDICT     whether the two agree, and precisely why not
'
' Nothing writes an OBSERVED field from a REQUIRED one. A requirement that
' reports its own nominal back as the observed nominal proves nothing, and
' that is the failure mode this separation exists to prevent.

' --- REQUIRED -----------------------------------------------------------

Public Key As String
Public Role As String
Public NominalM As Double

' Semicolon-separated swDimensionType_e values this requirement accepts.
' R23-804: the fixture has been seen in two states - type-6 diameter records
' in Phase 0, and plain linear records in the 2026-08-01 run - so a diameter
' requirement accepts both and no single type is ever required.
Public AcceptedTypes As String

' True for the two bore diameters. The nominal identifies them; this asks
' the separate question of whether the drawing DISPLAYS the match as a
' diameter, which IDisplayDimension.Diametric answers.
Public RequiresDiameterDisplay As Boolean

' R23-808 lane. The lane NAMES where this requirement belongs; it does not
' carry coordinates. Choosing page positions is layout, a later phase.
Public Lane As String

Public ToleranceRequired As Boolean
Public RequiredHoleFit As String
Public RequiredToleranceMinM As Double
Public RequiredToleranceMaxM As Double

' R23-806. Where the required tolerance comes from. For the 47 H7 this is
' the approved reference specification, NEVER the model: Phase 0 read
' D1@Sketch4 in the part directly and found toleranceType=0, fitType=-1 and
' empty fit strings. QA has to be able to say that out loud.
Public ToleranceAuthority As String

' --- OBSERVED -----------------------------------------------------------

Public MatchCount As Long
Public MatchedName As String
Public MatchedFullName As String
Public MatchedTypeCode As Long
Public MatchedNominalAvailable As Boolean
Public MatchedNominalM As Double
Public MatchedToleranceType As Long
Public MatchedFitType As Long
Public MatchedHoleFit As String
Public MatchedShaftFit As String
Public MatchedToleranceMinM As Double
Public MatchedToleranceMaxM As Double
Public MatchedToleranceMinStatus As Long
Public MatchedToleranceMaxStatus As Long
Public AttachedEntityCount As Long
Public AttachedEntityTypes As String
Public ToleranceProof As String
Public MatchedDiametric As Boolean
Public MatchedDiametricKnown As Boolean

' Which of the nominal-read routes actually answered. The supported
' configuration route declines for drawing-authored reference dimensions,
' so naming the route that worked is what lets a later run drop the ones
' that were never needed.
Public NominalRoute As String

' --- VERDICT ------------------------------------------------------------

Public Matched As Boolean
Public Origin As String
Public ToleranceSatisfied As Boolean
Public ToleranceProvenance As String
Public Failures As String

Private Sub Class_Initialize()
    MatchCount = 0
    Matched = False
    Origin = "NotReconciled"
    MatchedTypeCode = -1
    MatchedToleranceType = -1
    MatchedFitType = -1
    MatchedToleranceMinStatus = -1
    MatchedToleranceMaxStatus = -1
    MatchedDiametric = False
    MatchedDiametricKnown = False
    NominalRoute = "nominalRoute=NotRead"
    RequiresDiameterDisplay = False
    ToleranceSatisfied = False
    ToleranceProvenance = "NotApplied"
    ToleranceProof = "NotRead"
    AttachedEntityTypes = "NotRead"
    Failures = "NotReconciled"
End Sub

Public Function AcceptsTypeCode(ByVal typeCode As Long) As Boolean
    AcceptsTypeCode = (InStr(1, ";" & AcceptedTypes & ";", _
        ";" & CStr(typeCode) & ";", vbBinaryCompare) > 0)
End Function

Public Function RequirementSummary() As String
    RequirementSummary = "key=" & Key & _
        "|role=" & Role & _
        "|nominalM=" & Format$(NominalM, "0.000000") & _
        "|acceptedTypes=" & AcceptedTypes & _
        "|requiresDiameterDisplay=" & CStr(RequiresDiameterDisplay) & _
        "|lane=" & Lane & _
        "|toleranceRequired=" & CStr(ToleranceRequired) & _
        "|toleranceAuthority=" & ToleranceAuthority
End Function

Public Function ObservationSummary() As String
    ObservationSummary = "matched=" & CStr(Matched) & _
        "|origin=" & Origin & _
        "|matchCount=" & CStr(MatchCount) & _
        "|name=" & MatchedName & _
        "|fullName=" & MatchedFullName & _
        "|type2=" & CStr(MatchedTypeCode) & _
        "|nominalAvailable=" & CStr(MatchedNominalAvailable) & _
        "|nominalM=" & Format$(MatchedNominalM, "0.000000") & _
        "|" & NominalRoute & _
        "|diametricKnown=" & CStr(MatchedDiametricKnown) & _
        "|diametric=" & CStr(MatchedDiametric) & _
        "|attachedEntities=" & CStr(AttachedEntityCount) & _
        "|attachedTypes=" & AttachedEntityTypes & _
        "|" & ToleranceProof
End Function

Public Function Summary() As String
    Summary = RequirementSummary() & _
        "|" & ObservationSummary() & _
        "|toleranceSatisfied=" & CStr(ToleranceSatisfied) & _
        "|toleranceProvenance=" & ToleranceProvenance & _
        "|failures=" & Failures
End Function

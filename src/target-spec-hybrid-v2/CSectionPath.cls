Option Explicit

' R23 Phase 7 typed record for ONE section-line path built from model
' intent.
'
' The path is four waypoints and three segments, every coordinate taken from
' a proved projection rather than from a percentage of the view outline. The
' old strategy is disproved: it placed the upper label in the zone region
' and the lower arrow in the part-identification band, because a fraction of
' an outline knows nothing about where the holes are.
'
' Waypoint order, from R23-702:
'   1  the stepped-bore centre
'   2  the same X, at the highest face-hole row
'   3  the minimum-X face-hole column, at that row
'   4  the same column, at the lowest face-hole row
'
' All coordinates are PAGE coordinates until the single conversion in
' R23-705. Mixing frames is the defect that section work has hit before, so
' the frame is named in every field and the record never holds both.

Public SourceViewName As String
Public SourceView As SldWorks.View
Public SectionLabel As String

' Waypoints, page frame.
Public W1X As Double
Public W1Y As Double
Public W2X As Double
Public W2Y As Double
Public W3X As Double
Public W3Y As Double
Public W4X As Double
Public W4Y As Double

' The projections the path was built from, kept so the crossing proof cites
' geometry rather than repeating the numbers.
Public BoreProjection As CViewHoleProjection
Public ColumnProjections As Collection

Public DistinctColumns As Long
Public DistinctRows As Long

Public CrossingProofs As Collection
Public CrossingFailures As String

Public Resolved As Boolean
Public RejectionReason As String

' Set only by the mutating path.
Public SegmentsCreated As Long
Public SectionViewName As String
Public SectionLineInfoProof As String

Private Sub Class_Initialize()
    Set ColumnProjections = New Collection
    Set CrossingProofs = New Collection
    Resolved = False
    RejectionReason = "NotAttempted"
    CrossingFailures = "NotAttempted"
    SegmentsCreated = 0
    SectionLineInfoProof = "NotRead"
End Sub

' Three segments: 1-2, 2-3, 3-4.
Public Function SegmentCount() As Long
    SegmentCount = 3
End Function

Public Function WaypointSummary() As String
    WaypointSummary = "frame=Page" & _
        "|w1=" & Format$(W1X, "0.000000000") & _
            "," & Format$(W1Y, "0.000000000") & _
        "|w2=" & Format$(W2X, "0.000000000") & _
            "," & Format$(W2Y, "0.000000000") & _
        "|w3=" & Format$(W3X, "0.000000000") & _
            "," & Format$(W3Y, "0.000000000") & _
        "|w4=" & Format$(W4X, "0.000000000") & _
            "," & Format$(W4Y, "0.000000000")
End Function

Public Function Summary() As String
    Summary = "view=" & SourceViewName & _
        "|label=" & SectionLabel & _
        "|resolved=" & CStr(Resolved) & _
        "|reason=" & RejectionReason & _
        "|segments=" & CStr(SegmentCount()) & _
        "|distinctColumns=" & CStr(DistinctColumns) & _
        "|distinctRows=" & CStr(DistinctRows) & _
        "|columnHoles=" & CStr(ColumnProjections.Count) & _
        "|crossingsProven=" & CStr(CrossingProofs.Count) & _
        "|crossingFailures=" & CrossingFailures & _
        "|" & WaypointSummary()
End Function

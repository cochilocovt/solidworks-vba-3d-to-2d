Option Explicit

' R23 Phase 9 typed record for ONE page-frame content envelope.
'
' An envelope is the rectangle that everything belonging to a view actually
' occupies on the sheet: the model outline plus every dimension primitive,
' text box, note extent, leader point, callout, section segment, arrow and
' J-label that travels with it. Layout that reasons about the model outline
' alone moves a view into a place its annotations do not fit, which is the
' defect this record exists to remove.
'
' The same record type also describes a PROTECTED region - title block,
' part-identification band, content border, zone strips - so clearance is
' one comparison between two rectangles rather than two different kinds of
' check that can disagree.
'
' EVERY coordinate here is PAGE frame, in metres, relative to the sheet
' origin at the lower-left corner. Nothing in this record is ever view
' frame. Mixing the two is the defect this project has hit repeatedly, so
' the frame is named in the record and the module converts before
' contributing rather than storing both.

Public Kind As String
Public Name As String
Public Frame As String

Public Seeded As Boolean
Public MinX As Double
Public MinY As Double
Public MaxX As Double
Public MaxY As Double

' Per-source contribution counts. An envelope that grew only from the model
' outline is not a content envelope, and these counts are what proves it
' grew from the annotations too.
Public OutlinePoints As Long
Public DimensionPoints As Long
Public TextPoints As Long
Public NotePoints As Long
Public LeaderPoints As Long
Public ArrowPoints As Long
Public SectionPoints As Long
Public OtherPoints As Long

' Points discarded because they fell outside the sheet. A non-zero count is
' a frame smell, not a rounding artefact: page-frame geometry belonging to a
' view on the sheet lies on the sheet.
'
' The first rejected point is kept with its source. A bare count cannot tell
' a frame error from geometry that really does hang off the page, and the
' 2026-08-01 run reported 34 rejections on the section view without saying
' where a single one of them was.
Public RejectedOffSheet As Long
Public FirstRejectedX As Double
Public FirstRejectedY As Double
Public FirstRejectedSource As String

Public SourceFailures As String

Private Sub Class_Initialize()
    Kind = "Unknown"
    Frame = "Page"
    Seeded = False
    RejectedOffSheet = 0
    FirstRejectedSource = "None"
    SourceFailures = vbNullString
End Sub

Public Sub RecordRejection( _
    ByVal x As Double, _
    ByVal y As Double, _
    ByVal source As String)

    If RejectedOffSheet = 0 Then
        FirstRejectedX = x
        FirstRejectedY = y
        FirstRejectedSource = source
    End If

    RejectedOffSheet = RejectedOffSheet + 1
End Sub

Public Sub AddPoint( _
    ByVal x As Double, _
    ByVal y As Double, _
    ByVal source As String)

    If Not Seeded Then
        MinX = x
        MaxX = x
        MinY = y
        MaxY = y
        Seeded = True
    Else
        If x < MinX Then MinX = x
        If x > MaxX Then MaxX = x
        If y < MinY Then MinY = y
        If y > MaxY Then MaxY = y
    End If

    Select Case source
        Case "Outline"
            OutlinePoints = OutlinePoints + 1
        Case "Dimension"
            DimensionPoints = DimensionPoints + 1
        Case "Text"
            TextPoints = TextPoints + 1
        Case "Note"
            NotePoints = NotePoints + 1
        Case "Leader"
            LeaderPoints = LeaderPoints + 1
        Case "Arrow"
            ArrowPoints = ArrowPoints + 1
        Case "Section"
            SectionPoints = SectionPoints + 1
        Case Else
            OtherPoints = OtherPoints + 1
    End Select
End Sub

Public Sub AddRectangle( _
    ByVal leftX As Double, _
    ByVal bottomY As Double, _
    ByVal rightX As Double, _
    ByVal topY As Double, _
    ByVal source As String)

    AddPoint leftX, bottomY, source
    AddPoint rightX, topY, source
End Sub

Public Function ContributionCount() As Long
    ContributionCount = OutlinePoints + DimensionPoints + TextPoints + _
        NotePoints + LeaderPoints + ArrowPoints + SectionPoints + _
        OtherPoints
End Function

' True when the envelope grew from something other than the model outline.
' R23-900 asks for a CONTENT envelope; an outline-only rectangle is the old
' behaviour wearing a new name.
Public Function HasAnnotationContent() As Boolean
    HasAnnotationContent = _
        (ContributionCount() - OutlinePoints) > 0
End Function

Public Function Width() As Double
    If Seeded Then Width = MaxX - MinX
End Function

Public Function Height() As Double
    If Seeded Then Height = MaxY - MinY
End Function

Public Function CentreX() As Double
    If Seeded Then CentreX = (MinX + MaxX) / 2#
End Function

Public Function CentreY() As Double
    If Seeded Then CentreY = (MinY + MaxY) / 2#
End Function

Public Function BoundsToken() As String
    If Not Seeded Then
        BoundsToken = "bounds=Unseeded"
        Exit Function
    End If

    BoundsToken = "bounds=" & Format$(MinX, "0.000000") & _
        "," & Format$(MinY, "0.000000") & _
        "," & Format$(MaxX, "0.000000") & _
        "," & Format$(MaxY, "0.000000")
End Function

Public Function Summary() As String
    Dim failures As String
    failures = SourceFailures
    If Len(failures) = 0 Then failures = "None"

    Summary = "kind=" & Kind & _
        "|name=" & Name & _
        "|frame=" & Frame & _
        "|" & BoundsToken() & _
        "|width=" & Format$(Width(), "0.000000") & _
        "|height=" & Format$(Height(), "0.000000") & _
        "|outline=" & CStr(OutlinePoints) & _
        "|dimension=" & CStr(DimensionPoints) & _
        "|text=" & CStr(TextPoints) & _
        "|note=" & CStr(NotePoints) & _
        "|leader=" & CStr(LeaderPoints) & _
        "|arrow=" & CStr(ArrowPoints) & _
        "|section=" & CStr(SectionPoints) & _
        "|other=" & CStr(OtherPoints) & _
        "|rejectedOffSheet=" & CStr(RejectedOffSheet) & _
        "|firstRejected=" & FirstRejectedSource & _
            ":" & Format$(FirstRejectedX, "0.000000") & _
            "," & Format$(FirstRejectedY, "0.000000") & _
        "|annotationContent=" & CStr(HasAnnotationContent()) & _
        "|sourceFailures=" & failures
End Function

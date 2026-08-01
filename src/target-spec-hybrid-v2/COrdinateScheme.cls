Option Explicit

' R23 Phase 5 typed record for ONE ordinate scheme.
'
' R23-500 replaces feature-family grouping. A family says what a hole IS; it
' says nothing about which ordinate group the hole belongs to. Two holes of
' the same family machined from opposite faces belong to different groups,
' and two holes of different families machined from the same face in the
' same view belong to the same one. The scheme key is therefore
'
'   view role + machining face + datum policy + direction
'
' every part of which is measured rather than named. View role comes from
' the view type and from Phase 3's axis-normal measurement; machining face
' comes from the physical location's sign-normalized axis; the datum policy
' is chosen per direction; the direction is a swAddOrdinateDims_e member.
' A renamed or reoriented view therefore still lands in the right scheme.

Public ViewName As String
Public ViewType As Long
Public ViewRole As String
Public MachiningFaceKey As String
Public DatumPolicyId As String

Public Direction As Long
Public DirectionName As String

Public DrawingView As SldWorks.View

' R23-501 and R23-502. The datum is the FIRST selection of the transaction
' and everything else is measured from it, so it is recorded with its own
' proof rather than assumed.
Public DatumBucket As COrdinateBucket
Public DatumKind As String
Public DatumProof As String
Public DatumResolved As Boolean

Public Buckets As Collection

' R23-507. Locations that belong to the profile/reference ledger rather than
' the small-hole ledger are held separately and never enter Buckets, so a
' bore diameter can never be counted as small-hole coverage.
Public ProfileProjections As Collection

Public TransactionAttempted As Boolean
Public TransactionProof As String
Public CreatedDimensionCount As Long

Private Sub Class_Initialize()
    Set Buckets = New Collection
    Set ProfileProjections = New Collection
    DatumResolved = False
    DatumKind = "Unresolved"
    DatumProof = "NotAttempted"
    TransactionAttempted = False
    TransactionProof = "NotAttempted"
    CreatedDimensionCount = 0
End Sub

Public Function SchemeKey() As String
    SchemeKey = "view=" & ViewName & _
        "|role=" & ViewRole & _
        "|face=" & MachiningFaceKey & _
        "|datumPolicy=" & DatumPolicyId & _
        "|direction=" & DirectionName
End Function

Public Function BucketCount() As Long
    BucketCount = Buckets.Count
End Function

Public Function BucketByKey( _
    ByVal coordinateKey As String) As COrdinateBucket

    Dim i As Long
    For i = 1 To Buckets.Count
        Dim candidate As COrdinateBucket
        Set candidate = Buckets(i)
        If StrComp(candidate.CoordinateKey, coordinateKey, _
            vbBinaryCompare) = 0 Then

            Set BucketByKey = candidate
            Exit Function
        End If
    Next i

    Set BucketByKey = Nothing
End Function

Public Sub AddBucket(ByRef bucket As COrdinateBucket)
    If bucket Is Nothing Then Exit Sub
    Buckets.Add bucket
End Sub

' Total physical locations credited across every bucket. This is the number
' coverage is judged on; BucketCount is the number of dimensions required.
Public Function CreditedLocationCount() As Long
    Dim total As Long
    Dim i As Long

    For i = 1 To Buckets.Count
        Dim bucket As COrdinateBucket
        Set bucket = Buckets(i)
        total = total + bucket.CreditedCount()
    Next i

    CreditedLocationCount = total
End Function

Public Function AnchoredBucketCount() As Long
    Dim total As Long
    Dim i As Long

    For i = 1 To Buckets.Count
        Dim bucket As COrdinateBucket
        Set bucket = Buckets(i)
        If bucket.HasSelectableAnchor() Then total = total + 1
    Next i

    AnchoredBucketCount = total
End Function

Public Function CreatedBucketCount() As Long
    Dim total As Long
    Dim i As Long

    For i = 1 To Buckets.Count
        Dim bucket As COrdinateBucket
        Set bucket = Buckets(i)
        If bucket.Created Then total = total + 1
    Next i

    CreatedBucketCount = total
End Function

Public Function Summary() As String
    Summary = SchemeKey() & _
        "|viewType=" & CStr(ViewType) & _
        "|buckets=" & CStr(BucketCount()) & _
        "|anchoredBuckets=" & CStr(AnchoredBucketCount()) & _
        "|creditedLocations=" & CStr(CreditedLocationCount()) & _
        "|profileEntries=" & CStr(ProfileProjections.Count) & _
        "|datumResolved=" & CStr(DatumResolved) & _
        "|datumKind=" & DatumKind
End Function

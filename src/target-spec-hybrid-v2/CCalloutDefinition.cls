Option Explicit

' R23 Phase 6 typed record for ONE manufacturing callout definition covering
' ONE semantic family.
'
' A definition is either NATIVE - a SOLIDWORKS hole callout that already
' carries the Hole Wizard's own data - or a CONTROLLED FALLBACK built field
' by field from typed feature data. It is never free text. Every field
' carries the source that proved it, so a definition that looks complete can
' still be shown to be unproven.
'
' Quantity is the count of unique PHYSICAL LOCATIONS in the family. It is
' never a feature count and never an edge count: one Hole Wizard feature can
' produce six holes, a mirror can double them, and a counterbore contributes
' several edges per hole.

Public FamilyKey As String
Public MachiningFaceKey As String

' NativeCallout or ControlledFallback.
Public DefinitionSource As String

Public OwnerViewName As String
Public NativeAnnotation As SldWorks.DisplayDimension
Public NativeVariableCount As Long

' R23-606.
Public Quantity As Long
Public QuantityProofSource As String

Public NominalDiameterM As Double
Public DiameterProofSource As String

Public DepthM As Double
Public DepthProofSource As String

Public EndConditionCode As Long
Public EndConditionProofSource As String

' R23-607.
Public CounterBoreDiameterM As Double
Public CounterBoreDepthM As Double
Public CounterBoreProofSource As String

Public ThreadDescription As String
Public ThreadDepthM As Double
Public ThreadProofSource As String

' R23-608. Fit and tolerance are read, never invented, and never promoted
' from a drawing-authored value.
Public HoleFit As String
Public ShaftFit As String
Public ToleranceType As Long
Public ToleranceMinM As Double
Public ToleranceMaxM As Double
Public ToleranceProofSource As String

Public AttachmentProven As Boolean
Public AttachmentProofSource As String

Public MissingFields As String
Public RejectionReason As String

Private Sub Class_Initialize()
    DefinitionSource = "Unresolved"
    QuantityProofSource = "Unproven"
    DiameterProofSource = "Unproven"
    DepthProofSource = "Unproven"
    EndConditionProofSource = "Unproven"
    CounterBoreProofSource = "Unproven"
    ThreadProofSource = "Unproven"
    ToleranceProofSource = "Unproven"
    AttachmentProofSource = "Unproven"
    EndConditionCode = -1
    ToleranceType = -1
End Sub

Public Function IsNative() As Boolean
    IsNative = (StrComp(DefinitionSource, "NativeCallout", _
        vbBinaryCompare) = 0)
End Function

' R23-603 and R23-610. A definition is complete only when every field the
' hole's own geometry requires is present AND proved. A counterbored hole
' without counterbore data is incomplete; a plain hole without it is not.
Public Function CompletenessFailureReason() As String
    Dim missing As String

    If Quantity <= 0 Then
        missing = AppendMissing(missing, "Quantity")
    ElseIf StrComp(QuantityProofSource, "Unproven", vbBinaryCompare) = 0 Then
        missing = AppendMissing(missing, "QuantityUnproven")
    End If

    If NominalDiameterM <= 0# Then
        missing = AppendMissing(missing, "NominalDiameter")
    ElseIf StrComp(DiameterProofSource, "Unproven", vbBinaryCompare) = 0 Then
        missing = AppendMissing(missing, "NominalDiameterUnproven")
    End If

    If StrComp(EndConditionProofSource, "Unproven", _
        vbBinaryCompare) = 0 Then

        missing = AppendMissing(missing, "EndConditionUnproven")
    End If

    ' Depth is required only for a BLIND hole. swEndConditions_e verified
    ' against the 2025 table: swEndCondBlind = 0, and every through/up-to
    ' member is non-zero (ThroughAll 1, ThroughNext 2, UpToSurface 4,
    ' MidPlane 6, ThroughAllBoth 9, UpToNext 11).
    '
    ' The first live run proved this was written inverted. It demanded a
    ' depth from the counterbore family (EndCondition 2, ThroughNext, which
    ' has no depth to give) and let the M5 family pass with EndCondition 0
    ' and no depth at all - the dangerous direction, because a genuinely
    ' missing depth on a blind tapped hole went unreported.
    If EndConditionCode = 0 And DepthM <= 0# Then
        missing = AppendMissing(missing, "Depth")
    End If

    If CounterBoreDiameterM > 0# Then
        If CounterBoreDepthM <= 0# Then
            missing = AppendMissing(missing, "CounterBoreDepth")
        End If
        If StrComp(CounterBoreProofSource, "Unproven", _
            vbBinaryCompare) = 0 Then

            missing = AppendMissing(missing, "CounterBoreUnproven")
        End If
    End If

    If Len(Trim$(ThreadDescription)) > 0 Then
        If StrComp(ThreadProofSource, "Unproven", vbBinaryCompare) = 0 Then
            missing = AppendMissing(missing, "ThreadUnproven")
        End If
    End If

    If Not AttachmentProven Then
        missing = AppendMissing(missing, "Attachment")
    End If

    If Len(missing) = 0 Then missing = "None"

    MissingFields = missing
    CompletenessFailureReason = missing
End Function

Public Function IsComplete() As Boolean
    IsComplete = (StrComp(CompletenessFailureReason(), "None", _
        vbBinaryCompare) = 0)
End Function

Private Function AppendMissing( _
    ByVal existing As String, _
    ByVal field As String) As String

    If Len(existing) = 0 Then
        AppendMissing = field
    Else
        AppendMissing = existing & ";" & field
    End If
End Function

Public Function Summary() As String
    Summary = "family=" & _
        Module16_CalloutDefinition.EvidenceToken(FamilyKey) & _
        "|face=" & _
        Module16_CalloutDefinition.EvidenceToken(MachiningFaceKey) & _
        "|source=" & DefinitionSource & _
        "|view=" & OwnerViewName & _
        "|quantity=" & CStr(Quantity) & _
        "|quantitySource=" & QuantityProofSource & _
        "|diameterM=" & Format$(NominalDiameterM, "0.000000000") & _
        "|diameterSource=" & DiameterProofSource & _
        "|depthM=" & Format$(DepthM, "0.000000000") & _
        "|endCondition=" & CStr(EndConditionCode) & _
        "|cboreDiameterM=" & Format$(CounterBoreDiameterM, "0.000000000") & _
        "|cboreDepthM=" & Format$(CounterBoreDepthM, "0.000000000") & _
        "|thread=" & ThreadDescription & _
        "|threadDepthM=" & Format$(ThreadDepthM, "0.000000000") & _
        "|holeFit=" & HoleFit & _
        "|shaftFit=" & ShaftFit & _
        "|toleranceType=" & CStr(ToleranceType) & _
        "|toleranceSource=" & ToleranceProofSource & _
        "|nativeVariables=" & CStr(NativeVariableCount) & _
        "|attachmentProven=" & CStr(AttachmentProven) & _
        "|missingFields=" & CompletenessFailureReason()
End Function

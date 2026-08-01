Option Explicit

' R23 canonical numeric normalization for physical-feature identity.
'
' Physical identity must not depend on which point of an axis was sampled,
' on the sign of a returned direction vector, or on floating-point noise.
' Every key produced here is therefore built from quantized tokens.
'
' Nothing in this module calls the SOLIDWORKS API. It is pure arithmetic and
' string formatting so its invariants can be reasoned about and checked
' without a live session. It also hosts the COM Boolean normalization rule,
' which is interop rather than geometry but belongs with the other canonical
' normalizers because every caller needs it before comparing a SOLIDWORKS
' Boolean.

' Direction components below this magnitude are treated as zero when the
' sign-normalizing component is chosen.
Public Const AXIS_COMPONENT_TOLERANCE As Double = 0.000001

' Quantization used when a value becomes part of an identity key.
Public Const AXIS_KEY_QUANTUM As Double = 0.000001
Public Const MOMENT_KEY_QUANTUM_M As Double = 0.000001
Public Const AXIAL_KEY_QUANTUM_M As Double = 0.00001
Public Const LOCATION_KEY_QUANTUM_M As Double = 0.00001
Public Const RADIUS_KEY_QUANTUM_M As Double = 0.000001

' Comparison tolerances. These are deliberately looser than the key quanta
' so that a value sitting on a quantum boundary still compares equal.
Public Const AXIAL_MATCH_TOLERANCE_M As Double = 0.00002
Public Const LOCATION_MATCH_TOLERANCE_M As Double = 0.00005
Public Const RADIUS_MATCH_TOLERANCE_M As Double = 0.000005
Public Const TOLERANCE_MATCH_M As Double = 0.0000005

Private Const KEY_NUMBER_FORMAT As String = "0.000000000"

' Sign-normalizes a direction vector to unit length so that a line and its
' reverse produce the same tokens. Returns False for a degenerate vector.
Public Function NormalizeAxisDirection( _
    ByVal x As Double, _
    ByVal y As Double, _
    ByVal z As Double, _
    ByRef axisX As Double, _
    ByRef axisY As Double, _
    ByRef axisZ As Double) As Boolean

    axisX = 0#
    axisY = 0#
    axisZ = 0#

    Dim magnitude As Double
    magnitude = Sqr(x * x + y * y + z * z)
    If magnitude <= AXIS_COMPONENT_TOLERANCE Then Exit Function

    axisX = x / magnitude
    axisY = y / magnitude
    axisZ = z / magnitude

    If AxisRequiresSignFlip(axisX, axisY, axisZ) Then
        axisX = -axisX
        axisY = -axisY
        axisZ = -axisZ
    End If

    NormalizeAxisDirection = True
End Function

' The first component whose magnitude clears the tolerance decides the sign.
' This makes the choice deterministic instead of depending on the order in
' which SOLIDWORKS returned the face or edge.
Private Function AxisRequiresSignFlip( _
    ByVal axisX As Double, _
    ByVal axisY As Double, _
    ByVal axisZ As Double) As Boolean

    If Abs(axisX) > AXIS_COMPONENT_TOLERANCE Then
        AxisRequiresSignFlip = (axisX < 0#)
        Exit Function
    End If

    If Abs(axisY) > AXIS_COMPONENT_TOLERANCE Then
        AxisRequiresSignFlip = (axisY < 0#)
        Exit Function
    End If

    AxisRequiresSignFlip = (axisZ < 0#)
End Function

' Produces the sign-normalized direction plus the line moment
' (point cross direction). The moment is independent of which point on the
' axis is supplied, so two cylinders on one infinite line yield identical
' tokens even when their sampled origins differ.
Public Function NormalizeAxisAndMoment( _
    ByVal pointX As Double, _
    ByVal pointY As Double, _
    ByVal pointZ As Double, _
    ByVal directionX As Double, _
    ByVal directionY As Double, _
    ByVal directionZ As Double, _
    ByRef axisX As Double, _
    ByRef axisY As Double, _
    ByRef axisZ As Double, _
    ByRef momentX As Double, _
    ByRef momentY As Double, _
    ByRef momentZ As Double) As Boolean

    momentX = 0#
    momentY = 0#
    momentZ = 0#

    If Not NormalizeAxisDirection( _
        directionX, directionY, directionZ, _
        axisX, axisY, axisZ) Then Exit Function

    momentX = pointY * axisZ - pointZ * axisY
    momentY = pointZ * axisX - pointX * axisZ
    momentZ = pointX * axisY - pointY * axisX

    NormalizeAxisAndMoment = True
End Function

' Signed distance of a point along the normalized axis, measured from the
' projection of the model origin. Used to build axial intervals that keep
' opposite blind holes on one infinite line distinct.
Public Function AxialParameter( _
    ByVal pointX As Double, _
    ByVal pointY As Double, _
    ByVal pointZ As Double, _
    ByVal axisX As Double, _
    ByVal axisY As Double, _
    ByVal axisZ As Double) As Double

    AxialParameter = pointX * axisX + pointY * axisY + pointZ * axisZ
End Function

' Quantizes a value onto a tolerance grid and renders it as a stable token.
Public Function CanonicalNumberToken( _
    ByVal value As Double, _
    ByVal quantum As Double) As String

    Dim effectiveQuantum As Double
    effectiveQuantum = quantum
    If effectiveQuantum <= 0# Then effectiveQuantum = 0.000000001

    Dim snapped As Double
    snapped = Int((value / effectiveQuantum) + 0.5) * effectiveQuantum

    ' Collapse negative zero so that -0 and +0 share one token.
    If Abs(snapped) < (effectiveQuantum * 0.5) Then snapped = 0#

    CanonicalNumberToken = Format$(snapped, KEY_NUMBER_FORMAT)
End Function

Public Function CanonicalAxisToken( _
    ByVal axisX As Double, _
    ByVal axisY As Double, _
    ByVal axisZ As Double) As String

    CanonicalAxisToken = _
        CanonicalNumberToken(axisX, AXIS_KEY_QUANTUM) & "," & _
        CanonicalNumberToken(axisY, AXIS_KEY_QUANTUM) & "," & _
        CanonicalNumberToken(axisZ, AXIS_KEY_QUANTUM)
End Function

Public Function CanonicalMomentToken( _
    ByVal momentX As Double, _
    ByVal momentY As Double, _
    ByVal momentZ As Double) As String

    CanonicalMomentToken = _
        CanonicalNumberToken(momentX, MOMENT_KEY_QUANTUM_M) & "," & _
        CanonicalNumberToken(momentY, MOMENT_KEY_QUANTUM_M) & "," & _
        CanonicalNumberToken(momentZ, MOMENT_KEY_QUANTUM_M)
End Function

' Identity of the infinite line carrying a cylindrical feature. Two coaxial
' cylinders share this token regardless of radius, depth or owning feature.
Public Function BuildAxisLineKey( _
    ByVal axisX As Double, _
    ByVal axisY As Double, _
    ByVal axisZ As Double, _
    ByVal momentX As Double, _
    ByVal momentY As Double, _
    ByVal momentZ As Double) As String

    BuildAxisLineKey = _
        "axis=" & CanonicalAxisToken(axisX, axisY, axisZ) & _
        "|moment=" & CanonicalMomentToken(momentX, momentY, momentZ)
End Function

Public Function CanonicalIntervalToken( _
    ByVal intervalMin As Double, _
    ByVal intervalMax As Double) As String

    Dim lowValue As Double
    Dim highValue As Double
    lowValue = intervalMin
    highValue = intervalMax

    If highValue < lowValue Then
        Dim holdValue As Double
        holdValue = lowValue
        lowValue = highValue
        highValue = holdValue
    End If

    CanonicalIntervalToken = _
        CanonicalNumberToken(lowValue, AXIAL_KEY_QUANTUM_M) & ".." & _
        CanonicalNumberToken(highValue, AXIAL_KEY_QUANTUM_M)
End Function

Public Function CanonicalRadiusToken(ByVal radiusM As Double) As String
    CanonicalRadiusToken = _
        CanonicalNumberToken(Abs(radiusM), RADIUS_KEY_QUANTUM_M)
End Function

Public Function CanonicalLocationToken( _
    ByVal x As Double, _
    ByVal y As Double) As String

    CanonicalLocationToken = _
        CanonicalNumberToken(x, LOCATION_KEY_QUANTUM_M) & "," & _
        CanonicalNumberToken(y, LOCATION_KEY_QUANTUM_M)
End Function

Public Function ValuesMatchWithin( _
    ByVal firstValue As Double, _
    ByVal secondValue As Double, _
    ByVal tolerance As Double) As Boolean

    ValuesMatchWithin = (Abs(firstValue - secondValue) <= Abs(tolerance))
End Function

Public Function RadiiMatch( _
    ByVal firstRadiusM As Double, _
    ByVal secondRadiusM As Double) As Boolean

    RadiiMatch = ValuesMatchWithin( _
        Abs(firstRadiusM), Abs(secondRadiusM), RADIUS_MATCH_TOLERANCE_M)
End Function

Public Function LocationsMatch( _
    ByVal firstX As Double, _
    ByVal firstY As Double, _
    ByVal secondX As Double, _
    ByVal secondY As Double) As Boolean

    LocationsMatch = _
        ValuesMatchWithin(firstX, secondX, LOCATION_MATCH_TOLERANCE_M) And _
        ValuesMatchWithin(firstY, secondY, LOCATION_MATCH_TOLERANCE_M)
End Function

Public Function ToleranceValuesMatch( _
    ByVal firstValueM As Double, _
    ByVal secondValueM As Double) As Boolean

    ToleranceValuesMatch = _
        ValuesMatchWithin(firstValueM, secondValueM, TOLERANCE_MATCH_M)
End Function

' True when two axial intervals touch or overlap. Coaxial cylinders whose
' intervals meet belong to one stepped-bore stack; opposite blind holes on
' the same infinite line do not meet and stay separate.
Public Function AxialIntervalsOverlap( _
    ByVal firstMin As Double, _
    ByVal firstMax As Double, _
    ByVal secondMin As Double, _
    ByVal secondMax As Double) As Boolean

    Dim lowFirst As Double
    Dim highFirst As Double
    Dim lowSecond As Double
    Dim highSecond As Double

    OrderInterval firstMin, firstMax, lowFirst, highFirst
    OrderInterval secondMin, secondMax, lowSecond, highSecond

    AxialIntervalsOverlap = _
        (lowFirst <= highSecond + AXIAL_MATCH_TOLERANCE_M) And _
        (lowSecond <= highFirst + AXIAL_MATCH_TOLERANCE_M)
End Function

Private Sub OrderInterval( _
    ByVal firstValue As Double, _
    ByVal secondValue As Double, _
    ByRef lowValue As Double, _
    ByRef highValue As Double)

    If firstValue <= secondValue Then
        lowValue = firstValue
        highValue = secondValue
    Else
        lowValue = secondValue
        highValue = firstValue
    End If
End Sub

' The physical-instance key deliberately excludes every feature name so that
' renaming a feature cannot change physical identity.
Public Function BuildPhysicalLocationKey( _
    ByVal configurationName As String, _
    ByVal componentName As String, _
    ByVal lineKey As String, _
    ByVal intervalToken As String) As String

    BuildPhysicalLocationKey = _
        "config=" & IdentityToken(configurationName) & _
        "|component=" & IdentityToken(componentName) & _
        "|" & lineKey & _
        "|interval=" & intervalToken
End Function

' Family identity groups locations that share a manufacturing definition.
' It is built from semantic values, never from feature names.
Public Function BuildSemanticFamilyKey( _
    ByVal operationKind As String, _
    ByVal nominalDiameterM As Double, _
    ByVal depthM As Double, _
    ByVal counterBoreDiameterM As Double, _
    ByVal counterBoreDepthM As Double, _
    ByVal threadDescription As String) As String

    BuildSemanticFamilyKey = _
        "op=" & IdentityToken(operationKind) & _
        "|dia=" & CanonicalRadiusToken(nominalDiameterM) & _
        "|depth=" & CanonicalNumberToken(depthM, AXIAL_KEY_QUANTUM_M) & _
        "|cboreDia=" & CanonicalRadiusToken(counterBoreDiameterM) & _
        "|cboreDepth=" & _
            CanonicalNumberToken(counterBoreDepthM, AXIAL_KEY_QUANTUM_M) & _
        "|thread=" & IdentityToken(threadDescription)
End Function

Public Function BuildViewProjectionKey( _
    ByVal viewName As String, _
    ByVal physicalInstanceKey As String) As String

    BuildViewProjectionKey = _
        "view=" & IdentityToken(viewName) & _
        "|physical=" & physicalInstanceKey
End Function

' The only reliable way to consume a SOLIDWORKS COM Boolean in this VBA host.
'
' Live evidence from the 2026-07-31 P-0251 runs: a returned VARIANT_BOOL
' whose True is not VBA's -1 prints as "True" through CStr and behaves
' correctly in "If value Then" and "If value = False Then", but "Not value"
' yields -2, which VBA treats as True. CBool is not a dependable fix either:
' CBool(rawVariant) worked for ISurface.IsCylinder while CBool(curve.IsCircle)
' still failed under "Not". Converting through CDbl and comparing with zero is
' representation independent and cannot exhibit either failure mode.
'
' Every SOLIDWORKS Boolean must pass through here before any negation or
' compound logic.
Public Function NormalizeSwBoolean(ByVal rawValue As Variant) As Boolean
    On Error GoTo Failed

    If IsEmpty(rawValue) Or IsNull(rawValue) Then Exit Function
    NormalizeSwBoolean = (CDbl(rawValue) <> 0#)
    Exit Function

Failed:
    NormalizeSwBoolean = False
End Function

' Normalizes a string used inside a key so separators cannot be injected and
' case differences cannot split one identity into two.
Public Function IdentityToken(ByVal value As String) As String
    Dim normalized As String
    normalized = Trim$(value)
    normalized = Replace$(normalized, "|", "/")
    normalized = Replace$(normalized, vbCr, " ")
    normalized = Replace$(normalized, vbLf, " ")
    IdentityToken = UCase$(normalized)
End Function

Option Explicit

' R23 Phase 10. Semantic QA, replacing count-based checks.
' STRICTLY READ-ONLY. Nothing in this module creates, moves, deletes or
' selects anything. It reads the artefacts the earlier phases produced and
' decides whether each required stage is proved. R23_ProbeSemanticQA is an
' evidence entry point and, like every other R23 probe, does NOT call
' Module6_QAEngine.EmitRunEvidence - that is the production gate, and it
' demands fourteen pipeline stages a read-only probe never runs.
' WHY COUNTS ARE NOT ENOUGH. The checks this replaces ask "did anything get
' imported?" and "does the note contain the expected text?". Both pass on a
' drawing that is wrong in ways that matter: a nonzero import count is
' satisfied by importing one dimension into one view and none into the
' others, and a note-token check is satisfied by free text that no longer
' matches the geometry it sits beside. Every stage below asks a question
' about the PART instead - is this location dimensioned, is this definition
' backed by a source, does this requirement have a real dimension carrying
' the tolerance it claims.
' PIPELINE WIRING IS DEFERRED. Module6_QAEngine still runs the count-based
' checks on the reachable production path. Switching over is Phase 11's
' job, and doing it here would put unproven gates in front of a deployable
' macro - the same deferral R23-609, R23-704 and R23-810 already carry.

' R23-1000 to R23-1009. The required stages this phase adds. Each is a
' claim about the drawing that some artefact can prove or fail.
Public Const STAGE_MODEL_INTENT_CATALOG As String = _
    "MODEL_INTENT_CATALOG"
Public Const STAGE_MODEL_IMPORT_COVERAGE As String = _
    "MODEL_IMPORT_COVERAGE"
Public Const STAGE_NATIVE_CALLOUT_COVERAGE As String = _
    "NATIVE_CALLOUT_COVERAGE"
Public Const STAGE_PHYSICAL_LOCATION_GRAPH As String = _
    "PHYSICAL_LOCATION_GRAPH"
Public Const STAGE_VIEW_PROJECTION As String = "VIEW_PROJECTION"
Public Const STAGE_ORDINATE_SCHEME As String = "ORDINATE_SCHEME"
Public Const STAGE_SECTION_GEOMETRY As String = "SECTION_GEOMETRY"
Public Const STAGE_SECTION_DIMENSIONS As String = "SECTION_DIMENSIONS"
Public Const STAGE_FINAL_LAYOUT As String = "FINAL_LAYOUT"

' R23-1008. Retained from the existing stage set and strengthened here
' rather than replaced.
Public Const STAGE_MANUFACTURING_DEFINITION As String = _
    "MANUFACTURING_DEFINITION"

' R23-1014. Type-resolution outcomes that must fail the catalog. An ICE
' wrapper whose underlying type never resolved is a feature nobody has
' classified, and classifying it by guess is what R23 exists to stop.
Private Const RESOLUTION_ICE_UNRESOLVED As String = "IceUnresolved"
Private Const RESOLUTION_UNRESOLVED As String = "Unresolved"
Private Const RESOLUTION_READ_ERROR As String = "ReadError"

' R23-1012. Proof-source values that do not count as provenance. A field
' whose source is one of these has a number and no authority behind it.
Private Const PROOF_NONE As String = "None"
Private Const PROOF_UNPROVEN As String = "Unproven"

Private mEmitDiagnostics As Boolean

Private Sub EmitInfo( _
    ByRef evidence As CRunEvidence, _
    ByVal message As String)

    If Not evidence Is Nothing Then evidence.AddInfo message
    If mEmitDiagnostics Then
        Module21_EvidenceSink.LogLine message
    End If
End Sub

Public Function QaToken(ByVal value As String) As String
    Dim result As String
    result = value
    result = Replace$(result, "|", "/")
    result = Replace$(result, "=", ":")
    result = Replace$(result, vbCr, " ")
    result = Replace$(result, vbLf, " ")
    If Len(Trim$(result)) = 0 Then result = "Empty"
    QaToken = result
End Function

Private Function AppendFailure( _
    ByVal existing As String, _
    ByVal reason As String) As String

    If Len(existing) = 0 Or _
        StrComp(existing, "None", vbBinaryCompare) = 0 Then
        AppendFailure = reason
    Else
        AppendFailure = existing & ";" & reason
    End If
End Function

Private Function FormatMetres(ByVal value As Double) As String
    FormatMetres = Format$(value, "0.000000")
End Function

' R23-1000 to R23-1009. Declares every stage this phase adds, plus the
' retained MANUFACTURING_DEFINITION. Declaring a stage is what makes
' CRunEvidence.SealRequiredStages fail the run when nothing proves it, so
' the declaration is the gate - not the evaluation.
Public Sub RequireSemanticStages(ByRef evidence As CRunEvidence)
    evidence.RequireStage STAGE_MODEL_INTENT_CATALOG
    evidence.RequireStage STAGE_MODEL_IMPORT_COVERAGE
    evidence.RequireStage STAGE_NATIVE_CALLOUT_COVERAGE
    evidence.RequireStage STAGE_PHYSICAL_LOCATION_GRAPH
    evidence.RequireStage STAGE_VIEW_PROJECTION
    evidence.RequireStage STAGE_ORDINATE_SCHEME
    evidence.RequireStage STAGE_SECTION_GEOMETRY
    evidence.RequireStage STAGE_SECTION_DIMENSIONS
    evidence.RequireStage STAGE_MANUFACTURING_DEFINITION
    evidence.RequireStage STAGE_FINAL_LAYOUT
End Sub

' R23-1013 and R23-1014. Every audited feature reports its raw and
' effective type, accepted or not, and an unresolved ICE wrapper fails.
' R23-1013 says every feature, not every accepted feature, deliberately: a
' rejection with no type recorded cannot be reviewed, and "it was rejected"
' is not a reason.
Public Function EvaluateModelIntentCatalog( _
    ByRef graph As CLocationGraph, _
    ByRef evidence As CRunEvidence) As Boolean

    On Error GoTo Failed

    Dim failures As String
    Dim accepted As Long
    Dim rejected As Long
    Dim unresolved As Long

    Dim features As Collection
    Set features = graph.Features()

    If features.Count = 0 Then
        evidence.MarkStageFailed STAGE_MODEL_INTENT_CATALOG, _
            "no audited features"
        Exit Function
    End If

    Dim i As Long
    For i = 1 To features.Count
        Dim definition As CFeatureDefinition
        Set definition = features(i)

        If definition.Accepted Then
            accepted = accepted + 1
        Else
            rejected = rejected + 1
        End If

        EmitInfo evidence, "QA_FEATURE_TYPE|feature=" & _
            QaToken(definition.FeatureName) & _
            "|accepted=" & CStr(definition.Accepted) & _
            "|rawTypeName2=" & QaToken(definition.RawTypeName2) & _
            "|rawTypeName=" & QaToken(definition.RawTypeName) & _
            "|effectiveType=" & QaToken(definition.EffectiveType) & _
            "|resolutionSource=" & QaToken(definition.TypeResolutionSource) & _
            "|operationKind=" & QaToken(definition.OperationKind) & _
            "|rejection=" & QaToken(definition.RejectionReason)

        If TypeResolutionFailed(definition) Then
            unresolved = unresolved + 1
            failures = AppendFailure(failures, _
                "UnresolvedType:" & QaToken(definition.FeatureName) & _
                ":" & QaToken(definition.TypeResolutionSource))
        End If
    Next i

    Dim detail As String
    detail = "features=" & CStr(features.Count) & _
        "|accepted=" & CStr(accepted) & _
        "|rejected=" & CStr(rejected) & _
        "|unresolvedTypes=" & CStr(unresolved)

    If Len(failures) = 0 Then
        evidence.MarkStageProved STAGE_MODEL_INTENT_CATALOG, detail
        EvaluateModelIntentCatalog = True
    Else
        evidence.MarkStageFailed STAGE_MODEL_INTENT_CATALOG, _
            detail & "|failures=" & failures
    End If
    Exit Function

Failed:
    evidence.MarkStageFailed STAGE_MODEL_INTENT_CATALOG, _
        "error " & CStr(Err.Number)
End Function

' R23-1014. An ICE wrapper whose underlying type never resolved, a feature
' whose type could not be read at all, and an empty effective type are the
' three ways a feature reaches the catalog unclassified.
Private Function TypeResolutionFailed( _
    ByRef definition As CFeatureDefinition) As Boolean

    If Len(Trim$(definition.EffectiveType)) = 0 Then
        TypeResolutionFailed = True
        Exit Function
    End If

    Select Case definition.TypeResolutionSource
        Case RESOLUTION_ICE_UNRESOLVED, RESOLUTION_UNRESOLVED, _
            RESOLUTION_READ_ERROR

            TypeResolutionFailed = True
    End Select
End Function

' R23-1003, R23-1015 and R23-1016. The physical location graph is proved
' when every accepted location has a stable identity, no key is duplicated,
' and every accepted location reaches at least one proved projection.
Public Function EvaluatePhysicalLocationGraph( _
    ByRef graph As CLocationGraph, _
    ByRef evidence As CRunEvidence) As Boolean

    On Error GoTo Failed

    Dim failures As String

    Dim locations As Collection
    Set locations = graph.Locations()

    If locations.Count = 0 Then
        evidence.MarkStageFailed STAGE_PHYSICAL_LOCATION_GRAPH, _
            "no physical locations"
        Exit Function
    End If

    Dim keys As Collection
    Set keys = New Collection

    Dim proven As Long
    Dim unproven As Long

    Dim i As Long
    For i = 1 To locations.Count
        Dim location As CPhysicalHoleLocation
        Set location = locations(i)

        keys.Add location.PhysicalInstanceKey

        If location.IdentityProven Then
            proven = proven + 1
        Else
            unproven = unproven + 1
            failures = AppendFailure(failures, _
                "IdentityUnproven:" & _
                QaToken(location.PhysicalInstanceKey) & _
                ":" & QaToken(location.RejectionReason))
        End If
    Next i

    ' R23-1016. Two locations sharing one physical key means the graph has
    ' lost track of which hole is which, and everything downstream that
    ' looks a location up by key silently gets the wrong one.
    Dim duplicates As String
    duplicates = DuplicateKeyReport(keys, "physical")

    If StrComp(duplicates, "None", vbBinaryCompare) <> 0 Then
        failures = AppendFailure(failures, duplicates)
    End If

    Dim detail As String
    detail = "locations=" & CStr(locations.Count) & _
        "|identityProven=" & CStr(proven) & _
        "|identityUnproven=" & CStr(unproven) & _
        "|duplicateKeys=" & duplicates

    If Len(failures) = 0 Then
        evidence.MarkStageProved STAGE_PHYSICAL_LOCATION_GRAPH, detail
        EvaluatePhysicalLocationGraph = True
    Else
        evidence.MarkStageFailed STAGE_PHYSICAL_LOCATION_GRAPH, _
            detail & "|failures=" & failures
    End If
    Exit Function

Failed:
    evidence.MarkStageFailed STAGE_PHYSICAL_LOCATION_GRAPH, _
        "error " & CStr(Err.Number)
End Function

' R23-1004 and R23-1015. Every location that the part requires must reach a
' proved projection in some view. A location with no projection anywhere is
' a hole nobody can dimension, and reporting the projection COUNT would
' hide it behind the locations that did project.
Public Function EvaluateViewProjection( _
    ByRef graph As CLocationGraph, _
    ByRef evidence As CRunEvidence) As Boolean

    On Error GoTo Failed

    Dim failures As String

    Dim locations As Collection
    Set locations = graph.Locations()

    Dim projections As Collection
    Set projections = graph.Projections()

    Dim projectedKeys As Object
    Set projectedKeys = CreateObject("Scripting.Dictionary")

    Dim acceptedProjections As Long

    Dim i As Long
    For i = 1 To projections.Count
        Dim projection As CViewHoleProjection
        Set projection = projections(i)

        If projection.Accepted Then
            acceptedProjections = acceptedProjections + 1
            If Not projectedKeys.Exists(projection.PhysicalInstanceKey) Then
                projectedKeys.Add projection.PhysicalInstanceKey, True
            End If
        End If
    Next i

    Dim unprojected As Long

    For i = 1 To locations.Count
        Dim location As CPhysicalHoleLocation
        Set location = locations(i)

        If Not location.IdentityProven Then GoTo ContinueLocation
        If projectedKeys.Exists(location.PhysicalInstanceKey) Then
            GoTo ContinueLocation
        End If

        unprojected = unprojected + 1
        failures = AppendFailure(failures, _
            "NoProvedProjection:" & _
            QaToken(location.PhysicalInstanceKey))

ContinueLocation:
    Next i

    Dim detail As String
    detail = "locations=" & CStr(locations.Count) & _
        "|projections=" & CStr(projections.Count) & _
        "|acceptedProjections=" & CStr(acceptedProjections) & _
        "|locationsWithProjection=" & CStr(projectedKeys.Count) & _
        "|locationsWithout=" & CStr(unprojected)

    If Len(failures) = 0 Then
        evidence.MarkStageProved STAGE_VIEW_PROJECTION, detail
        EvaluateViewProjection = True
    Else
        evidence.MarkStageFailed STAGE_VIEW_PROJECTION, _
            detail & "|failures=" & failures
    End If
    Exit Function

Failed:
    evidence.MarkStageFailed STAGE_VIEW_PROJECTION, _
        "error " & CStr(Err.Number)
End Function

' R23-1001 and R23-1010. Import coverage per VIEW and per CATEGORY.
' The check this replaces asked whether the import returned a nonzero
' count. That passes when one view receives every dimension and the rest
' receive none, which is a drawing that looks imported and is not. Here
' each view reports its own accepted projections and how many of them are
' covered in each direction, and a view with accepted projections and no
' coverage at all fails by name.
Public Function EvaluateModelImportCoverage( _
    ByRef graph As CLocationGraph, _
    ByRef evidence As CRunEvidence) As Boolean

    On Error GoTo Failed

    Dim failures As String

    Dim projections As Collection
    Set projections = graph.Projections()

    If projections.Count = 0 Then
        evidence.MarkStageFailed STAGE_MODEL_IMPORT_COVERAGE, _
            "no projections to cover"
        Exit Function
    End If

    Dim viewNames As Object
    Set viewNames = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = 1 To projections.Count
        Dim listed As CViewHoleProjection
        Set listed = projections(i)
        If Not viewNames.Exists(listed.ViewName) Then
            viewNames.Add listed.ViewName, True
        End If
    Next i

    Dim viewsWithContent As Long

    Dim viewKey As Variant
    For Each viewKey In viewNames.Keys
        Dim viewProjections As Collection
        Set viewProjections = graph.ProjectionsForView(CStr(viewKey))

        Dim accepted As Long
        Dim coveredX As Long
        Dim coveredY As Long
        Dim annotated As Long
        accepted = 0
        coveredX = 0
        coveredY = 0
        annotated = 0

        Dim j As Long
        For j = 1 To viewProjections.Count
            Dim projection As CViewHoleProjection
            Set projection = viewProjections(j)

            If Not projection.Accepted Then GoTo ContinueProjection

            accepted = accepted + 1
            If projection.CoveredX Then coveredX = coveredX + 1
            If projection.CoveredY Then coveredY = coveredY + 1
            If projection.AttachedAnnotations.Count > 0 Then
                annotated = annotated + 1
            End If

ContinueProjection:
        Next j

        EmitInfo evidence, "QA_IMPORT_COVERAGE|view=" & _
            QaToken(CStr(viewKey)) & _
            "|projections=" & CStr(viewProjections.Count) & _
            "|accepted=" & CStr(accepted) & _
            "|coveredX=" & CStr(coveredX) & _
            "|coveredY=" & CStr(coveredY) & _
            "|annotated=" & CStr(annotated)

        If accepted > 0 Then
            viewsWithContent = viewsWithContent + 1

            If coveredX = 0 And coveredY = 0 And annotated = 0 Then
                failures = AppendFailure(failures, _
                    "ViewImportedNothing:" & QaToken(CStr(viewKey)) & _
                    ":accepted=" & CStr(accepted))
            End If
        End If
    Next viewKey

    Dim detail As String
    detail = "views=" & CStr(viewNames.Count) & _
        "|viewsWithAcceptedProjections=" & CStr(viewsWithContent)

    If viewsWithContent = 0 Then
        failures = AppendFailure(failures, "NoViewHasAcceptedProjections")
    End If

    If Len(failures) = 0 Then
        evidence.MarkStageProved STAGE_MODEL_IMPORT_COVERAGE, detail
        EvaluateModelImportCoverage = True
    Else
        evidence.MarkStageFailed STAGE_MODEL_IMPORT_COVERAGE, _
            detail & "|failures=" & failures
    End If
    Exit Function

Failed:
    evidence.MarkStageFailed STAGE_MODEL_IMPORT_COVERAGE, _
        "error " & CStr(Err.Number)
End Function

' R23-1002, R23-1012 and R23-1016. Native callout coverage, with the source
' of every manufacturing field emitted beside its value.
' A number without provenance is the failure mode this catches. "6.6" is
' correct or wrong depending on whether it was read from the feature, read
' from a callout variable, or assumed - and only the source says which.
Public Function EvaluateNativeCalloutCoverage( _
    ByRef definitions As Collection, _
    ByRef evidence As CRunEvidence) As Boolean

    On Error GoTo Failed

    Dim failures As String

    If definitions.Count = 0 Then
        evidence.MarkStageFailed STAGE_NATIVE_CALLOUT_COVERAGE, _
            "no callout definitions"
        Exit Function
    End If

    Dim keys As Collection
    Set keys = New Collection

    Dim native As Long
    Dim fallback As Long
    Dim incomplete As Long
    Dim unprovenFields As Long

    Dim i As Long
    For i = 1 To definitions.Count
        Dim definition As CCalloutDefinition
        Set definition = definitions(i)

        keys.Add definition.FamilyKey

        If definition.IsNative() Then
            native = native + 1
        Else
            fallback = fallback + 1
        End If

        Dim missingProvenance As String
        missingProvenance = ProvenanceGaps(definition)

        If Len(missingProvenance) > 0 Then
            unprovenFields = unprovenFields + 1
            failures = AppendFailure(failures, _
                "NoProvenance:" & QaToken(definition.FamilyKey) & _
                ":" & missingProvenance)
        End If

        If Not definition.IsComplete() Then
            incomplete = incomplete + 1
            failures = AppendFailure(failures, _
                "Incomplete:" & QaToken(definition.FamilyKey) & _
                ":" & QaToken(definition.CompletenessFailureReason()))
        End If

        EmitInfo evidence, "QA_CALLOUT_PROVENANCE|family=" & _
            QaToken(definition.FamilyKey) & _
            "|source=" & QaToken(definition.DefinitionSource) & _
            "|diameter=" & FormatMetres(definition.NominalDiameterM) & _
            "|diameterFrom=" & QaToken(definition.DiameterProofSource) & _
            "|depth=" & FormatMetres(definition.DepthM) & _
            "|depthFrom=" & QaToken(definition.DepthProofSource) & _
            "|endCondition=" & CStr(definition.EndConditionCode) & _
            "|endConditionFrom=" & _
                QaToken(definition.EndConditionProofSource) & _
            "|quantityFrom=" & QaToken(definition.QuantityProofSource) & _
            "|attachmentFrom=" & QaToken(definition.AttachmentProofSource)
    Next i

    Dim duplicates As String
    duplicates = DuplicateKeyReport(keys, "familyDefinition")

    If StrComp(duplicates, "None", vbBinaryCompare) <> 0 Then
        failures = AppendFailure(failures, duplicates)
    End If

    Dim detail As String
    detail = "definitions=" & CStr(definitions.Count) & _
        "|native=" & CStr(native) & _
        "|controlledFallback=" & CStr(fallback) & _
        "|incomplete=" & CStr(incomplete) & _
        "|withoutProvenance=" & CStr(unprovenFields) & _
        "|duplicateKeys=" & duplicates

    If Len(failures) = 0 Then
        evidence.MarkStageProved STAGE_NATIVE_CALLOUT_COVERAGE, detail
        EvaluateNativeCalloutCoverage = True
    Else
        evidence.MarkStageFailed STAGE_NATIVE_CALLOUT_COVERAGE, _
            detail & "|failures=" & failures
    End If
    Exit Function

Failed:
    evidence.MarkStageFailed STAGE_NATIVE_CALLOUT_COVERAGE, _
        "error " & CStr(Err.Number)
End Function

' R23-1012. Names the manufacturing fields that carry a value with no
' source behind it. A blank source, "None" and "Unproven" all mean the same
' thing here: nobody can say where the number came from.
Private Function ProvenanceGaps( _
    ByRef definition As CCalloutDefinition) As String

    Dim gaps As String

    If definition.NominalDiameterM > 0# Then
        If Not SourceIsReal(definition.DiameterProofSource) Then
            gaps = AppendFailure(gaps, "diameter")
        End If
    End If

    If definition.DepthM > 0# Then
        If Not SourceIsReal(definition.DepthProofSource) Then
            gaps = AppendFailure(gaps, "depth")
        End If
    End If

    If definition.CounterBoreDiameterM > 0# Then
        If Not SourceIsReal(definition.CounterBoreProofSource) Then
            gaps = AppendFailure(gaps, "counterBore")
        End If
    End If

    If definition.ThreadDepthM > 0# Then
        If Not SourceIsReal(definition.ThreadProofSource) Then
            gaps = AppendFailure(gaps, "thread")
        End If
    End If

    If definition.Quantity > 0 Then
        If Not SourceIsReal(definition.QuantityProofSource) Then
            gaps = AppendFailure(gaps, "quantity")
        End If
    End If

    ProvenanceGaps = gaps
End Function

Private Function SourceIsReal(ByVal source As String) As Boolean
    Dim trimmed As String
    trimmed = Trim$(source)

    If Len(trimmed) = 0 Then Exit Function
    If StrComp(trimmed, PROOF_NONE, vbTextCompare) = 0 Then Exit Function
    If StrComp(trimmed, PROOF_UNPROVEN, vbTextCompare) = 0 Then Exit Function

    SourceIsReal = True
End Function

' R23-1008. The retained manufacturing stage, strengthened: a definition
' counts only when it is complete AND attached to real drawing geometry.
' Text that describes the part correctly but hangs off nothing is not a
' manufacturing definition, it is a caption.
Public Function EvaluateManufacturingDefinition( _
    ByRef definitions As Collection, _
    ByRef evidence As CRunEvidence) As Boolean

    On Error GoTo Failed

    Dim failures As String
    Dim attached As Long
    Dim complete As Long

    If definitions.Count = 0 Then
        evidence.MarkStageFailed STAGE_MANUFACTURING_DEFINITION, _
            "no manufacturing definitions"
        Exit Function
    End If

    Dim i As Long
    For i = 1 To definitions.Count
        Dim definition As CCalloutDefinition
        Set definition = definitions(i)

        If definition.IsComplete() Then complete = complete + 1

        If definition.AttachmentProven Then
            attached = attached + 1
        Else
            failures = AppendFailure(failures, _
                "Unattached:" & QaToken(definition.FamilyKey))
        End If
    Next i

    Dim detail As String
    detail = "definitions=" & CStr(definitions.Count) & _
        "|complete=" & CStr(complete) & _
        "|attachmentProven=" & CStr(attached)

    If complete <> definitions.Count Then
        failures = AppendFailure(failures, _
            "IncompleteDefinitions:" & _
            CStr(definitions.Count - complete))
    End If

    If Len(failures) = 0 Then
        evidence.MarkStageProved STAGE_MANUFACTURING_DEFINITION, detail
        EvaluateManufacturingDefinition = True
    Else
        evidence.MarkStageFailed STAGE_MANUFACTURING_DEFINITION, _
            detail & "|failures=" & failures
    End If
    Exit Function

Failed:
    evidence.MarkStageFailed STAGE_MANUFACTURING_DEFINITION, _
        "error " & CStr(Err.Number)
End Function

' R23-1005. The ordinate stage is proved when every scheme resolved its
' datum and credited every location it claims to represent. A count of
' created dimensions proves nothing: two groups can be created against the
' wrong datum and still count as two.
Public Function EvaluateOrdinateScheme( _
    ByRef schemes As Collection, _
    ByRef evidence As CRunEvidence) As Boolean

    On Error GoTo Failed

    Dim failures As String
    Dim resolved As Long
    Dim credited As Long

    If schemes.Count = 0 Then
        evidence.MarkStageFailed STAGE_ORDINATE_SCHEME, _
            "no ordinate schemes"
        Exit Function
    End If

    Dim i As Long
    For i = 1 To schemes.Count
        Dim scheme As COrdinateScheme
        Set scheme = schemes(i)

        credited = credited + scheme.CreditedLocationCount()

        If scheme.DatumResolved Then
            resolved = resolved + 1
        Else
            failures = AppendFailure(failures, _
                "DatumUnresolved:" & QaToken(scheme.SchemeKey()))
        End If

        If scheme.BucketCount() = 0 Then
            failures = AppendFailure(failures, _
                "NoBuckets:" & QaToken(scheme.SchemeKey()))
        End If

        EmitInfo evidence, "QA_ORDINATE_SCHEME|" & scheme.Summary()
    Next i

    Dim detail As String
    detail = "schemes=" & CStr(schemes.Count) & _
        "|datumResolved=" & CStr(resolved) & _
        "|creditedLocations=" & CStr(credited)

    If Len(failures) = 0 Then
        evidence.MarkStageProved STAGE_ORDINATE_SCHEME, detail
        EvaluateOrdinateScheme = True
    Else
        evidence.MarkStageFailed STAGE_ORDINATE_SCHEME, _
            detail & "|failures=" & failures
    End If
    Exit Function

Failed:
    evidence.MarkStageFailed STAGE_ORDINATE_SCHEME, _
        "error " & CStr(Err.Number)
End Function

' R23-1006 and R23-1017. The section stage is proved when the path resolved
' with its crossings proved. An unavailable or unsafe J-J fails rather than
' being approximated - the same rule R23-708 states, enforced here as a
' stage instead of a return value nobody has to read.
Public Function EvaluateSectionGeometry( _
    ByRef path As CSectionPath, _
    ByRef evidence As CRunEvidence) As Boolean

    On Error GoTo Failed

    If path Is Nothing Then
        evidence.MarkStageFailed STAGE_SECTION_GEOMETRY, _
            "no section path resolved"
        Exit Function
    End If

    Dim verdict As String
    verdict = Module17_SectionPath.VerifySectionGeometry(path)

    If Not path.Resolved Then
        evidence.MarkStageFailed STAGE_SECTION_GEOMETRY, _
            "unresolved|" & QaToken(path.RejectionReason) & _
            "|" & QaToken(verdict)
        Exit Function
    End If

    If InStr(1, verdict, "sectionFailures=None", vbBinaryCompare) = 0 Then
        evidence.MarkStageFailed STAGE_SECTION_GEOMETRY, QaToken(verdict)
        Exit Function
    End If

    evidence.MarkStageProved STAGE_SECTION_GEOMETRY, QaToken(verdict)
    EvaluateSectionGeometry = True
    Exit Function

Failed:
    evidence.MarkStageFailed STAGE_SECTION_GEOMETRY, _
        "error " & CStr(Err.Number)
End Function

' R23-1007, R23-1011 and R23-1016. Section dimensions judged by inspecting
' the DIMENSIONS - type, nominal, attachment, tolerance - and never by
' looking for expected text in a note.
' The check this replaces searched a note for a stepped-bore token. That
' passes on free text that has drifted from the geometry, and fails on a
' correct drawing whose wording differs. Neither outcome says anything
' about the part.
Public Function EvaluateSectionDimensions( _
    ByRef requirements As Collection, _
    ByRef dimensions As Collection, _
    ByRef evidence As CRunEvidence) As Boolean

    On Error GoTo Failed

    If requirements.Count = 0 Then
        evidence.MarkStageFailed STAGE_SECTION_DIMENSIONS, _
            "no section requirements"
        Exit Function
    End If

    Dim keys As Collection
    Set keys = New Collection

    Dim i As Long
    For i = 1 To requirements.Count
        Dim requirement As CSectionRequirement
        Set requirement = requirements(i)
        keys.Add requirement.Key
    Next i

    Dim duplicates As String
    duplicates = DuplicateKeyReport(keys, "sectionRequirement")

    Dim verdict As String
    verdict = Module10_SectionDimensionEngine.VerifySectionDimensions( _
        requirements, dimensions)

    Dim detail As String
    detail = QaToken(verdict) & "|duplicateKeys=" & duplicates

    If InStr(1, verdict, "requirementFailures=None", _
        vbBinaryCompare) = 0 Then

        evidence.MarkStageFailed STAGE_SECTION_DIMENSIONS, detail
        Exit Function
    End If

    If StrComp(duplicates, "None", vbBinaryCompare) <> 0 Then
        evidence.MarkStageFailed STAGE_SECTION_DIMENSIONS, detail
        Exit Function
    End If

    evidence.MarkStageProved STAGE_SECTION_DIMENSIONS, detail
    EvaluateSectionDimensions = True
    Exit Function

Failed:
    evidence.MarkStageFailed STAGE_SECTION_DIMENSIONS, _
        "error " & CStr(Err.Number)
End Function

' R23-1009 and R23-1017. The final stage. Every view envelope must be
' available and every clearance must hold; an envelope that could not be
' built is an unsafe layout, not a missing number.
Public Function EvaluateFinalLayout( _
    ByRef envelopes As Collection, _
    ByRef views As Collection, _
    ByRef protectedRegions As Collection, _
    ByRef evidence As CRunEvidence) As Boolean

    On Error GoTo Failed

    If Module1_Main.R23_LAYOUT_USER_ACCEPTED_AS_IS Then
        evidence.AddWarning "FINAL_LAYOUT|status=UserAcceptedAsIs" & _
            "|automaticClearance=DeferredByUser"
        evidence.MarkStageProved STAGE_FINAL_LAYOUT, _
            "policy=UserAcceptedAsIs" & _
            "|automaticClearance=DeferredByUser"
        EvaluateFinalLayout = True
        Exit Function
    End If

    Dim failures As String

    If envelopes.Count = 0 Then
        evidence.MarkStageFailed STAGE_FINAL_LAYOUT, "no view envelopes"
        Exit Function
    End If

    Dim withAnnotationContent As Long

    Dim i As Long
    For i = 1 To envelopes.Count
        Dim envelope As CContentEnvelope
        Set envelope = envelopes(i)

        If Not envelope.Seeded Then
            failures = AppendFailure(failures, _
                "EnvelopeUnavailable:" & QaToken(envelope.Name))
        ElseIf envelope.HasAnnotationContent() Then
            withAnnotationContent = withAnnotationContent + 1
        End If
    Next i

    Dim verdict As String
    verdict = Module18_ContentEnvelope.VerifyClearances( _
        envelopes, views, protectedRegions)

    If InStr(1, verdict, "clearanceFailures=None", vbBinaryCompare) = 0 Then
        failures = AppendFailure(failures, QaToken(verdict))
    End If

    Dim detail As String
    detail = "envelopes=" & CStr(envelopes.Count) & _
        "|withAnnotationContent=" & CStr(withAnnotationContent) & _
        "|protectedRegions=" & CStr(protectedRegions.Count) & _
        "|" & QaToken(verdict)

    If Len(failures) = 0 Then
        evidence.MarkStageProved STAGE_FINAL_LAYOUT, detail
        EvaluateFinalLayout = True
    Else
        evidence.MarkStageFailed STAGE_FINAL_LAYOUT, _
            detail & "|failures=" & failures
    End If
    Exit Function

Failed:
    evidence.MarkStageFailed STAGE_FINAL_LAYOUT, _
        "error " & CStr(Err.Number)
End Function

' R23-1016. Names every key that appears more than once, and how often.
' Returning "None" rather than an empty string keeps the caller from
' treating "no duplicates" and "the check did not run" as the same result.
Public Function DuplicateKeyReport( _
    ByRef keys As Collection, _
    ByVal kind As String) As String

    On Error GoTo Failed

    Dim counts As Object
    Set counts = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = 1 To keys.Count
        Dim key As String
        key = CStr(keys(i))

        If counts.Exists(key) Then
            counts(key) = CLng(counts(key)) + 1
        Else
            counts.Add key, 1
        End If
    Next i

    Dim duplicates As String

    Dim listed As Variant
    For Each listed In counts.Keys
        If CLng(counts(listed)) > 1 Then
            duplicates = AppendFailure(duplicates, _
                "Duplicate" & kind & ":" & QaToken(CStr(listed)) & _
                ":" & CStr(counts(listed)))
        End If
    Next listed

    If Len(duplicates) = 0 Then duplicates = "None"
    DuplicateKeyReport = duplicates
    Exit Function

Failed:
    DuplicateKeyReport = "DuplicateCheckError:" & CStr(Err.Number)
End Function

' Reports every required stage and its status, so a reader sees which gate
' failed rather than only that the run failed.
Public Function StageReport(ByRef evidence As CRunEvidence) As String
    Dim report As String

    Dim names As Variant
    names = Array( _
        STAGE_MODEL_INTENT_CATALOG, STAGE_MODEL_IMPORT_COVERAGE, _
        STAGE_NATIVE_CALLOUT_COVERAGE, STAGE_PHYSICAL_LOCATION_GRAPH, _
        STAGE_VIEW_PROJECTION, STAGE_ORDINATE_SCHEME, _
        STAGE_SECTION_GEOMETRY, STAGE_SECTION_DIMENSIONS, _
        STAGE_MANUFACTURING_DEFINITION, STAGE_FINAL_LAYOUT)

    Dim proved As Long

    Dim i As Long
    For i = LBound(names) To UBound(names)
        Dim stageName As String
        stageName = CStr(names(i))

        If evidence.StageIsProved(stageName) Then proved = proved + 1

        report = report & "|" & stageName & "=" & _
            CStr(evidence.StageIsProved(stageName))
    Next i

    StageReport = "stages=" & CStr(UBound(names) - LBound(names) + 1) & _
        "|proved=" & CStr(proved) & report
End Function

' Collects one retained definition per family, using only
' Module16_CalloutDefinition's PUBLIC decision functions. The loop is local;
' every judgement inside it - what counts as a native callout, which family
' a callout belongs to, which of the native and fallback definitions is
' retained - is Module16's, so the two cannot drift apart on the part that
' matters.
Public Function CollectRetainedDefinitions( _
    ByRef swApp As SldWorks.SldWorks, _
    ByRef swDraw As SldWorks.ModelDoc2, _
    ByRef views As Variant, _
    ByRef graph As CLocationGraph, _
    ByRef evidence As CRunEvidence) As Collection

    Dim retained As Collection
    Set retained = New Collection
    Set CollectRetainedDefinitions = retained

    On Error GoTo Failed

    Dim nativeByFamily As Object
    Set nativeByFamily = CreateObject("Scripting.Dictionary")

    Dim i As Long
    For i = LBound(views) To UBound(views)
        Dim swView As SldWorks.View
        Set swView = views(i)
        If swView Is Nothing Then GoTo ContinueView

        Dim viewName As String
        viewName = SafeViewName(swView)

        Dim dimensions As Variant
        dimensions = swView.GetDisplayDimensions
        If Not IsArray(dimensions) Then GoTo ContinueView

        Dim d As Long
        For d = LBound(dimensions) To UBound(dimensions)
            Dim swDispDim As SldWorks.DisplayDimension
            Set swDispDim = Nothing
            On Error Resume Next
            Set swDispDim = dimensions(d)
            On Error GoTo Failed
            If swDispDim Is Nothing Then GoTo ContinueDimension

            If Not Module16_CalloutDefinition.IsNativeHoleCallout( _
                swDispDim) Then

                GoTo ContinueDimension
            End If

            Dim matchDiagnostics As String
            Dim familyKey As String
            familyKey = Module16_CalloutDefinition.MatchCalloutToFamily( _
                swApp, graph, swDispDim, viewName, matchDiagnostics)

            If Len(familyKey) = 0 Then GoTo ContinueDimension
            If nativeByFamily.Exists(familyKey) Then GoTo ContinueDimension

            Dim nativeDefinition As CCalloutDefinition
            Set nativeDefinition = _
                Module16_CalloutDefinition.BuildDefinitionFromTypedData( _
                    graph, familyKey, evidence)

            nativeDefinition.DefinitionSource = _
                Module16_CalloutDefinition.DEFINITION_NATIVE
            nativeDefinition.OwnerViewName = viewName
            Set nativeDefinition.NativeAnnotation = swDispDim
            nativeDefinition.AttachmentProven = True
            nativeDefinition.AttachmentProofSource = matchDiagnostics

            Module16_CalloutDefinition.ReadNativeCalloutFields _
                swDispDim, nativeDefinition, evidence

            nativeByFamily.Add familyKey, nativeDefinition

ContinueDimension:
        Next d

ContinueView:
    Next i

    Dim familyKeys As Collection
    Set familyKeys = graph.FamilyKeys()

    Dim f As Long
    For f = 1 To familyKeys.Count
        Dim key As String
        key = CStr(familyKeys(f))

        Dim fallbackDefinition As CCalloutDefinition
        Set fallbackDefinition = _
            Module16_CalloutDefinition.BuildDefinitionFromTypedData( _
                graph, key, evidence)

        Dim nativeCandidate As CCalloutDefinition
        Set nativeCandidate = Nothing
        If nativeByFamily.Exists(key) Then
            Set nativeCandidate = nativeByFamily(key)
        End If

        Dim decisionProof As String
        Dim chosen As CCalloutDefinition
        Set chosen = _
            Module16_CalloutDefinition.RetainDefinitionForFamily( _
                nativeCandidate, fallbackDefinition, decisionProof)

        If chosen Is Nothing Then GoTo ContinueFamily

        retained.Add chosen

ContinueFamily:
    Next f

    Exit Function

Failed:
    Set CollectRetainedDefinitions = retained
End Function

' Collects every ordinate scheme on the sheet, one call per view, through
' Module15's public builder.
Public Function CollectOrdinateSchemes( _
    ByRef swApp As SldWorks.SldWorks, _
    ByRef swDraw As SldWorks.ModelDoc2, _
    ByRef views As Variant, _
    ByRef graph As CLocationGraph, _
    ByRef evidence As CRunEvidence) As Collection

    Dim result As Collection
    Set result = New Collection
    Set CollectOrdinateSchemes = result

    On Error GoTo Failed

    Dim i As Long
    For i = LBound(views) To UBound(views)
        Dim swView As SldWorks.View
        Set swView = views(i)
        If swView Is Nothing Then GoTo ContinueView

        Dim schemes As Collection
        Set schemes = Module15_OrdinateScheme.BuildSchemesForView( _
            swApp, swDraw, swView, graph, evidence)

        Dim j As Long
        For j = 1 To schemes.Count
            result.Add schemes(j)
        Next j

ContinueView:
    Next i

    Exit Function

Failed:
    Set CollectOrdinateSchemes = result
End Function

' R23-1101, steps 7 and 15. This is the production form of the semantic
' judge. It owns no mutation: the caller has already created, imported and
' laid out the drawing. Keeping the judgement here means the production
' pipeline and R23_ProbeSemanticQA ask the same questions over the same
' graph rather than drifting into two definitions of "complete".
Public Function EvaluateSemanticDrawing( _
    ByRef swApp As SldWorks.SldWorks, _
    ByRef swDraw As SldWorks.ModelDoc2, _
    ByRef swDrawing As SldWorks.DrawingDoc, _
    ByRef graph As CLocationGraph, _
    ByRef evidence As CRunEvidence) As Boolean

    On Error GoTo Failed

    If swApp Is Nothing Or swDraw Is Nothing Or swDrawing Is Nothing Or _
        graph Is Nothing Or evidence Is Nothing Then

        Exit Function
    End If

    RequireSemanticStages evidence

    Dim swSheet As SldWorks.Sheet
    Set swSheet = swDrawing.GetCurrentSheet
    If swSheet Is Nothing Then
        evidence.MarkStageFailed STAGE_FINAL_LAYOUT, "current sheet is Nothing"
        evidence.SealRequiredStages
        Exit Function
    End If

    Dim views As Variant
    views = swSheet.GetViews
    If IsEmpty(views) Or Not IsArray(views) Then
        evidence.MarkStageFailed STAGE_VIEW_PROJECTION, "no views on sheet"
        evidence.MarkStageFailed STAGE_FINAL_LAYOUT, "no views on sheet"
        evidence.SealRequiredStages
        Exit Function
    End If

    ' Inventory is deliberately after import. InsertModelAnnotations4 returns
    ' a result set, not coverage, so the sheet is traversed independently.
    graph.ClearImportedAnnotations

    Dim annotationInventoryBuilt As Boolean
    annotationInventoryBuilt = _
        Module14_AnnotationImport.BuildAnnotationInventory( _
            swDraw, graph, evidence)

    Dim annotationsReconciled As Long
    annotationsReconciled = -1
    If annotationInventoryBuilt Then
        annotationsReconciled = _
            Module14_AnnotationImport.ReconcileWithLocationGraph( _
                swApp, swDraw, graph, evidence)
    End If

    EmitInfo evidence, "SEMANTIC_ANNOTATION|inventoryBuilt=" & _
        CStr(annotationInventoryBuilt) & _
        "|annotations=" & CStr(graph.AnnotationCount()) & _
        "|reconciled=" & CStr(annotationsReconciled)

    EvaluateModelIntentCatalog graph, evidence
    EvaluatePhysicalLocationGraph graph, evidence
    EvaluateViewProjection graph, evidence
    ' Observe imported or previously created coverage before this evaluator
    ' constructs any read-only candidate ordinate schemes.
    EvaluateModelImportCoverage graph, evidence

    Dim definitions As Collection
    Set definitions = CollectRetainedDefinitions( _
        swApp, swDraw, views, graph, evidence)

    EvaluateNativeCalloutCoverage definitions, evidence
    EvaluateManufacturingDefinition definitions, evidence

    Dim schemes As Collection
    Set schemes = CollectOrdinateSchemes( _
        swApp, swDraw, views, graph, evidence)
    EvaluateOrdinateScheme schemes, evidence

    Dim sectionLabel As String
    sectionLabel = Module1_Main.GetSectionLabelOrDefault(1)

    Dim bestPath As CSectionPath
    Set bestPath = ResolveBestSectionPath( _
        graph, views, sectionLabel, evidence)
    EvaluateSectionGeometry bestPath, evidence

    EvaluateExistingSectionDimensions views, evidence
    EvaluateCurrentLayout swSheet, views, evidence

    evidence.SealRequiredStages
    EvaluateSemanticDrawing = AllSemanticStagesProved(evidence)
    Exit Function

Failed:
    evidence.MarkStageFailed STAGE_FINAL_LAYOUT, _
        "semantic evaluation error " & CStr(Err.Number)
    evidence.SealRequiredStages
End Function

Private Function ResolveBestSectionPath( _
    ByRef graph As CLocationGraph, _
    ByRef views As Variant, _
    ByVal sectionLabel As String, _
    ByRef evidence As CRunEvidence) As CSectionPath

    On Error GoTo Failed

    Dim i As Long
    For i = LBound(views) To UBound(views)
        Dim pathView As SldWorks.View
        Set pathView = views(i)
        If pathView Is Nothing Then GoTo ContinuePathView

        Dim path As CSectionPath
        Set path = Module17_SectionPath.ResolveSectionPath( _
            graph, pathView, sectionLabel, evidence)

        If path.Resolved Then
            Set ResolveBestSectionPath = path
            Exit Function
        End If

ContinuePathView:
    Next i
    Exit Function

Failed:
    Set ResolveBestSectionPath = Nothing
End Function

Private Sub EvaluateExistingSectionDimensions( _
    ByRef views As Variant, _
    ByRef evidence As CRunEvidence)

    Dim sectionViews As Collection
    Set sectionViews = _
        Module10_SectionDimensionEngine.CollectSectionViews(views)

    If sectionViews.Count = 0 Then
        evidence.MarkStageFailed STAGE_SECTION_DIMENSIONS, _
            "no section view on sheet"
        Exit Sub
    End If

    Dim sectionView As SldWorks.View
    Set sectionView = sectionViews(1)

    Dim sectionDimensions As Collection
    Set sectionDimensions = _
        Module10_SectionDimensionEngine.InventorySectionDimensions( _
            sectionView, evidence)

    Dim requirements As Collection
    Set requirements = _
        Module10_SectionDimensionEngine.BuildSectionRequirements()

    Module10_SectionDimensionEngine.ReconcileSectionDimensions _
        requirements, sectionDimensions, evidence

    EvaluateSectionDimensions requirements, sectionDimensions, evidence
End Sub

Private Sub EvaluateCurrentLayout( _
    ByRef swSheet As SldWorks.Sheet, _
    ByRef views As Variant, _
    ByRef evidence As CRunEvidence)

    If Module1_Main.R23_LAYOUT_USER_ACCEPTED_AS_IS Then
        Dim deferredEnvelopes As Collection
        Dim deferredViews As Collection
        Dim deferredRegions As Collection
        Set deferredEnvelopes = New Collection
        Set deferredViews = New Collection
        Set deferredRegions = New Collection

        EvaluateFinalLayout deferredEnvelopes, deferredViews, _
            deferredRegions, evidence
        Exit Sub
    End If

    Dim sheetProof As String
    If Not Module18_ContentEnvelope.MeasureSheetRegions( _
        swSheet, evidence, sheetProof) Then

        evidence.MarkStageFailed STAGE_FINAL_LAYOUT, _
            "sheet regions unmeasured|" & QaToken(sheetProof)
        Exit Sub
    End If

    Dim viewList As Collection
    Dim envelopes As Collection
    Set viewList = New Collection
    Set envelopes = New Collection

    Dim i As Long
    For i = LBound(views) To UBound(views)
        Dim envelopeView As SldWorks.View
        Set envelopeView = views(i)
        If envelopeView Is Nothing Then GoTo ContinueEnvelopeView

        If Module18_ContentEnvelope.IsTemplateOrientationView( _
            envelopeView) Then

            GoTo ContinueEnvelopeView
        End If

        Dim envelope As CContentEnvelope
        Set envelope = Module18_ContentEnvelope.BuildViewEnvelope( _
            envelopeView, evidence.SheetWidth, evidence.SheetHeight, evidence)

        If Not envelope.Seeded Then GoTo ContinueEnvelopeView

        viewList.Add envelopeView
        envelopes.Add envelope

ContinueEnvelopeView:
    Next i

    EvaluateFinalLayout envelopes, viewList, _
        Module18_ContentEnvelope.BuildProtectedRegions(evidence), evidence
End Sub

Private Function AllSemanticStagesProved( _
    ByRef evidence As CRunEvidence) As Boolean

    Dim names As Variant
    names = Array( _
        STAGE_MODEL_INTENT_CATALOG, STAGE_MODEL_IMPORT_COVERAGE, _
        STAGE_NATIVE_CALLOUT_COVERAGE, STAGE_PHYSICAL_LOCATION_GRAPH, _
        STAGE_VIEW_PROJECTION, STAGE_ORDINATE_SCHEME, _
        STAGE_SECTION_GEOMETRY, STAGE_SECTION_DIMENSIONS, _
        STAGE_MANUFACTURING_DEFINITION, STAGE_FINAL_LAYOUT)

    Dim i As Long
    For i = LBound(names) To UBound(names)
        If Not evidence.StageIsProved(CStr(names(i))) Then Exit Function
    Next i

    AllSemanticStagesProved = True
End Function

Private Function SafeViewName( _
    ByRef swView As SldWorks.View) As String

    On Error Resume Next
    If swView Is Nothing Then
        SafeViewName = "Nothing"
        Exit Function
    End If

    SafeViewName = "Unnamed"
    SafeViewName = swView.GetName2
End Function

' R23-1000 to R23-1017 evidence entry point. STRICTLY READ-ONLY: it builds
' the graph, the definitions, the section path and the envelopes exactly as
' the earlier probes do, then runs every semantic stage against them. It
' creates nothing and calls no production gate.
Public Sub R23_ProbeSemanticQA()
    On Error GoTo Failed

    mEmitDiagnostics = False

    Dim swApp As SldWorks.SldWorks
    Set swApp = Application.SldWorks

    If swApp Is Nothing Then
        Module21_EvidenceSink.LogLine _
            "R23_SEMANTICQA_FATAL|reason=SolidWorksUnavailable"
        Exit Sub
    End If

    Dim swDraw As SldWorks.ModelDoc2
    Set swDraw = swApp.ActiveDoc

    If swDraw Is Nothing Then
        Module21_EvidenceSink.LogLine _
            "R23_SEMANTICQA_FATAL|reason=NoActiveDocument"
        Exit Sub
    End If

    If swDraw.GetType <> swDocDRAWING Then
        Module21_EvidenceSink.LogLine _
            "R23_SEMANTICQA_FATAL" & _
            "|reason=ActiveDocumentNotDrawing"
        Exit Sub
    End If

    Dim swDrawing As SldWorks.DrawingDoc
    Set swDrawing = swDraw

    Dim swSheet As SldWorks.Sheet
    Set swSheet = swDrawing.GetCurrentSheet

    Dim views As Variant
    views = swSheet.GetViews

    If IsEmpty(views) Or Not IsArray(views) Then
        Module21_EvidenceSink.LogLine _
            "R23_SEMANTICQA_FATAL|reason=NoViewsOnSheet"
        Exit Sub
    End If

    Dim swPart As SldWorks.ModelDoc2
    Set swPart = Nothing

    Dim i As Long
    For i = LBound(views) To UBound(views)
        Dim candidateView As SldWorks.View
        Set candidateView = views(i)
        If Not candidateView Is Nothing Then
            If Not candidateView.ReferencedDocument Is Nothing Then
                Set swPart = candidateView.ReferencedDocument
                Exit For
            End If
        End If
    Next i

    If swPart Is Nothing Then
        Module21_EvidenceSink.LogLine _
            "R23_SEMANTICQA_FATAL|reason=NoReferencedDocument"
        Exit Sub
    End If

    Dim partPath As String
    partPath = swPart.GetPathName

    If Not Module1_Main.IsAuthorizedFixture(partPath) Then
        Module21_EvidenceSink.LogLine _
            "R23_SEMANTICQA_FATAL|reason=UnauthorizedFixture" & _
            "|path=" & partPath
        Exit Sub
    End If

    Dim drawingSaveBefore As Boolean
    drawingSaveBefore = Module11_GeometryIdentity.NormalizeSwBoolean( _
        swDraw.GetSaveFlag)

    Dim initialSelectionCount As Long
    initialSelectionCount = _
        swDraw.SelectionManager.GetSelectedObjectCount2(-1)

    Dim evidence As CRunEvidence
    Set evidence = New CRunEvidence

    Module21_EvidenceSink.LogLine _
        "R23_SEMANTICQA_BEGIN|drawing=" & swDraw.GetPathName & _
        "|part=" & partPath & _
        "|fixture=" & Module1_Main.GetFixtureKey(partPath) & _
        "|mode=ReadOnly|creations=0|mutations=0"

    RequireSemanticStages evidence

    Dim graph As CLocationGraph
    Set graph = New CLocationGraph

    Dim configurationName As String
    configurationName = _
        swPart.ConfigurationManager.ActiveConfiguration.Name

    If Not Module12_FeatureQualification.BuildFeatureCatalog( _
        swApp, swPart, configurationName, graph, evidence) Then

        Module21_EvidenceSink.LogLine _
            "R23_SEMANTICQA_FATAL|reason=CatalogUnavailable"
        Exit Sub
    End If

    For i = LBound(views) To UBound(views)
        Dim swView As SldWorks.View
        Set swView = views(i)
        If swView Is Nothing Then GoTo ContinueProjectionView

        Module13_ProjectionResolution.BuildViewProjections _
            swApp, swDraw, swView, graph, evidence

ContinueProjectionView:
    Next i

    EvaluateSemanticDrawing swApp, swDraw, swDrawing, graph, evidence

    Dim finalSelectionCount As Long
    finalSelectionCount = _
        swDraw.SelectionManager.GetSelectedObjectCount2(-1)

    Dim drawingSaveAfter As Boolean
    drawingSaveAfter = Module11_GeometryIdentity.NormalizeSwBoolean( _
        swDraw.GetSaveFlag)

    Module21_EvidenceSink.LogLine _
        "R23_SEMANTICQA_END|" & StageReport(evidence) & _
        "|failures=" & CStr(evidence.FailureCount) & _
        "|warnings=" & CStr(evidence.WarningCount) & _
        "|creations=0|mutations=0" & _
        "|initialSelectionCount=" & CStr(initialSelectionCount) & _
        "|finalSelectionCount=" & CStr(finalSelectionCount) & _
        "|drawingUnchanged=" & _
        CStr(drawingSaveBefore = drawingSaveAfter)
    Exit Sub

Failed:
    Module21_EvidenceSink.LogLine _
        "R23_SEMANTICQA_FATAL|error=" & CStr(Err.Number) & _
        "|description=" & Err.Description

    On Error Resume Next
    If Not swDraw Is Nothing Then
        swDraw.SetPickMode
        swDraw.ClearSelection2 True
    End If
End Sub

' R23 probe-runner compile-failure localisation. A no-op; VBA compiles
' at module granularity, so a module that loads this has compiled.
Public Sub R23_CompileTouch()
End Sub

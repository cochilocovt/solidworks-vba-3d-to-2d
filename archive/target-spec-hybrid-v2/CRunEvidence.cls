Option Explicit

Private mInfo As Collection
Private mWarnings As Collection
Private mFailures As Collection
Private mStageRequired As Object
Private mStageStatus As Object
Private mStageDetail As Object
Private mStageSealLogged As Object
Private mPhysicalLocations As Object
Private mProjectionLocations As Object
Private mPhysicalLocationRecords As Collection
Private mProjectionLocationRecords As Collection
Private mTitlePropertyRecords As Collection
Private mManufacturingCandidates As Collection

Public RunId As String
Public PartPath As String
Public ConfigurationName As String
Public SolidWorksRevision As String
Public MacroRevision As String
Public TemplatePath As String
Public SheetTemplateName As String
Public SheetFormatName As String
Public DatumOrigin As String
Public SheetScale As Double
Public SettingsSummary As String
Public RequestedScaleNumerator As Long
Public RequestedScaleDenominator As Long
Public ActualScaleNumerator As Double
Public ActualScaleDenominator As Double
Public SheetScaleReadbackProven As Boolean

Public SheetWidth As Double
Public SheetHeight As Double
Public UsableLeft As Double
Public UsableBottom As Double
Public UsableRight As Double
Public UsableTop As Double
Public ContentBorderLeft As Double
Public ContentBorderBottom As Double
Public ContentBorderRight As Double
Public ContentBorderTop As Double
Public TitleBlockLeft As Double
Public TitleBlockBottom As Double
Public TitleBlockRight As Double
Public TitleBlockTop As Double
Public PartIdentificationLeft As Double
Public PartIdentificationBottom As Double
Public PartIdentificationRight As Double
Public PartIdentificationTop As Double
Public PartIdentificationBoundsProven As Boolean
Public LayoutBoundariesProven As Boolean
Public FinalSelectionCount As Long
Public SheetContextRestored As Boolean
Public FinalCleanupVerified As Boolean
Public FinalCleanupInvalidated As Boolean
Public SolidWorksMutationSequence As Long
Public LastSolidWorksMutation As String
Public EvidenceReportPath As String
Public EvidenceWriteAttempted As Boolean
Public EvidenceWriteSucceeded As Boolean
Public EvidenceWriteVerified As Boolean

Public ViewsCreated As Long
Public ImportedAnnotations As Long
Public ImportedDisplayDimensions As Long
Public CandidatesAccepted As Long
Public CandidatesRejected As Long
Public HorizontalOrdinateGroups As Long
Public VerticalOrdinateGroups As Long
Public OrdinateEntitiesSelected As Long
Public LayoutMoves As Long
Public TitlePropertiesWritten As Long

Private Sub Class_Initialize()
    Set mInfo = New Collection
    Set mWarnings = New Collection
    Set mFailures = New Collection
    Set mStageRequired = CreateTextDictionary()
    Set mStageStatus = CreateTextDictionary()
    Set mStageDetail = CreateTextDictionary()
    Set mStageSealLogged = CreateTextDictionary()
    Set mPhysicalLocations = CreateTextDictionary()
    Set mProjectionLocations = CreateTextDictionary()
    Set mPhysicalLocationRecords = New Collection
    Set mProjectionLocationRecords = New Collection
    Set mTitlePropertyRecords = New Collection
    Set mManufacturingCandidates = New Collection
End Sub

Public Sub RequireStage(ByVal stageName As String)
    Dim key As String
    key = NormalizeStageName(stageName)
    If Len(key) = 0 Then Exit Sub

    If Not mStageRequired.Exists(key) Then mStageRequired.Add key, True
    If Not mStageStatus.Exists(key) Then mStageStatus.Add key, "NOT_STARTED"
    If Not mStageDetail.Exists(key) Then mStageDetail.Add key, vbNullString
End Sub

Public Sub MarkStageProved( _
    ByVal stageName As String, _
    ByVal detail As String)

    Dim key As String
    key = NormalizeStageName(stageName)
    If Len(key) = 0 Then Exit Sub

    RequireStage key
    If StrComp(StageStatusText(key), "FAILED", vbTextCompare) <> 0 Then
        mStageStatus(key) = "PROVED"
        mStageDetail(key) = SanitizeEvidenceValue(detail)
    Else
        AppendStageDetail key, "later proof attempt: " & detail
    End If
End Sub

Public Sub MarkStageFailed( _
    ByVal stageName As String, _
    ByVal detail As String)

    Dim key As String
    key = NormalizeStageName(stageName)
    If Len(key) = 0 Then Exit Sub

    RequireStage key
    mStageStatus(key) = "FAILED"
    AppendStageDetail key, detail
End Sub

Public Sub MarkStageNotApplicable( _
    ByVal stageName As String, _
    ByVal detail As String)

    Dim key As String
    key = NormalizeStageName(stageName)
    If Len(key) = 0 Then Exit Sub

    If Not mStageStatus.Exists(key) Then mStageStatus.Add key, "NOT_APPLICABLE"
    If StrComp(StageStatusText(key), "FAILED", vbTextCompare) = 0 Then _
        Exit Sub
    mStageStatus(key) = "NOT_APPLICABLE"
    If Not mStageDetail.Exists(key) Then mStageDetail.Add key, vbNullString
    mStageDetail(key) = SanitizeEvidenceValue(detail)
End Sub

Public Function StageIsProved(ByVal stageName As String) As Boolean
    Dim key As String
    key = NormalizeStageName(stageName)
    If Len(key) = 0 Then Exit Function
    If Not mStageStatus.Exists(key) Then Exit Function

    StageIsProved = (StrComp( _
        CStr(mStageStatus(key)), "PROVED", vbTextCompare) = 0)
End Function

Public Sub SealRequiredStages()
    Dim key As Variant
    For Each key In mStageRequired.Keys
        If Not StageIsProved(CStr(key)) Then
            If Not mStageSealLogged.Exists(CStr(key)) Then
                AddFailure "STAGE_GATE|stage=" & CStr(key) & _
                    "|status=" & StageStatusText(CStr(key)) & _
                    "|detail=" & StageDetailText(CStr(key))
                mStageSealLogged.Add CStr(key), True
            End If
        End If
    Next key
End Sub

Public Function RegisterPhysicalLocation( _
    ByVal physicalKey As String, _
    ByVal componentName As String, _
    ByVal featureName As String, _
    ByVal configurationName As String, _
    ByVal modelCentreKey As String, _
    ByVal axisKey As String, _
    ByVal sourceView As String) As Boolean

    Dim key As String
    key = LCase$(Trim$(physicalKey))

    If Len(key) = 0 Or Len(Trim$(modelCentreKey)) = 0 Or _
       Len(Trim$(axisKey)) = 0 Then

        AddFailure _
            "PHYSICAL_LOCATION_GATE|status=FAILED|reason=incomplete-canoni" & _
            "cal-identity"
        MarkStageFailed "PHYSICAL_LOCATIONS", _
            "physical key, model centre, or axis is empty"
        Exit Function
    End If

    Dim identitySignature As String
    identitySignature = _
        LCase$(Trim$(componentName)) & "|" & _
        LCase$(Trim$(featureName)) & "|" & _
        LCase$(Trim$(configurationName)) & "|" & _
        LCase$(Trim$(modelCentreKey)) & "|" & _
        LCase$(Trim$(axisKey))

    If Not mPhysicalLocations.Exists(key) Then
        mPhysicalLocations.Add key, identitySignature
        mPhysicalLocationRecords.Add _
            "PHYSICAL_LOCATION|key=" & SanitizeEvidenceValue(physicalKey) & _
            "|component=" & SanitizeEvidenceValue(componentName) & _
            "|feature=" & SanitizeEvidenceValue(featureName) & _
            "|configuration=" & SanitizeEvidenceValue(configurationName) & _
            "|model_centre=" & SanitizeEvidenceValue(modelCentreKey) & _
            "|axis=" & SanitizeEvidenceValue(axisKey) & _
            "|first_view=" & SanitizeEvidenceValue(sourceView)
    ElseIf StrComp( _
        CStr(mPhysicalLocations(key)), identitySignature, _
        vbTextCompare) <> 0 Then

        AddFailure _
            "PHYSICAL_LOCATION_GATE|status=FAILED|reason=key-collision|key" & _
            "=" & _
            SanitizeEvidenceValue(physicalKey)
        MarkStageFailed "PHYSICAL_LOCATIONS", _
            "one physical key resolved to conflicting canonical identities"
        Exit Function
    End If

    RegisterPhysicalLocation = True
End Function

Public Sub AddProjectionLocationRecord( _
    ByVal physicalKey As String, _
    ByVal viewName As String, _
    ByVal viewX As Double, _
    ByVal viewY As Double, _
    ByVal drawingEntityType As String)

    Dim physicalLookupKey As String
    physicalLookupKey = LCase$(Trim$(physicalKey))

    If Len(physicalLookupKey) = 0 Or Len(Trim$(viewName)) = 0 Then
        AddFailure _
            "PROJECTION_LOCATION_GATE|status=FAILED|reason=incomplete-key-" & _
            "or-view"
        MarkStageFailed "ORDINATE_COVERAGE", _
            "projection key or view is empty"
        Exit Sub
    End If

    If Not mPhysicalLocations.Exists(physicalLookupKey) Then
        AddFailure _
            "PROJECTION_LOCATION_GATE|status=FAILED|reason=unregistered-ph" & _
            "ysical-key|key=" & _
            SanitizeEvidenceValue(physicalKey)
        MarkStageFailed "ORDINATE_COVERAGE", _
            "projection references an unregistered physical location"
        Exit Sub
    End If

    Dim projectionKey As String
    projectionKey = physicalLookupKey & "|" & LCase$(Trim$(viewName))

    Dim projectionSignature As String
    projectionSignature = Format$(viewX, "0.000000") & "|" & _
        Format$(viewY, "0.000000") & "|" & _
        LCase$(Trim$(drawingEntityType))

    If mProjectionLocations.Exists(projectionKey) Then
        If StrComp( _
            CStr(mProjectionLocations(projectionKey)), projectionSignature, _
            vbTextCompare) <> 0 Then

            AddFailure _
                "PROJECTION_LOCATION_GATE|status=FAILED|reason=view-key-co" & _
                "llision|key=" & _
                SanitizeEvidenceValue(projectionKey)
            MarkStageFailed "ORDINATE_COVERAGE", _
                "one physical/view key resolved to conflicting projected " & _
                    "locations"
        End If
        Exit Sub
    End If

    mProjectionLocations.Add projectionKey, projectionSignature
    mProjectionLocationRecords.Add _
        "PROJECTION_LOCATION|physical_key=" & _
        SanitizeEvidenceValue(physicalKey) & _
        "|view=" & SanitizeEvidenceValue(viewName) & _
        "|view_x_m=" & Format$(viewX, "0.000000") & _
        "|view_y_m=" & Format$(viewY, "0.000000") & _
        "|entity=" & SanitizeEvidenceValue(drawingEntityType)
End Sub

Public Property Get UniquePhysicalLocationCount() As Long
    UniquePhysicalLocationCount = mPhysicalLocations.Count
End Property

Public Property Get ProjectionLocationRecordCount() As Long
    ProjectionLocationRecordCount = mProjectionLocations.Count
End Property

Public Sub RegisterManufacturingCandidate( _
    ByRef candidate As CHoleCandidate)

    If candidate Is Nothing Then Exit Sub
    If Not candidate.Accepted Then Exit Sub

    Dim i As Long
    For i = 1 To mManufacturingCandidates.Count
        Dim existing As CHoleCandidate
        Set existing = mManufacturingCandidates(i)

        If StrComp(existing.ViewName, candidate.ViewName, vbTextCompare) = _
            0 And _
           StrComp(existing.PhysicalInstanceKey, _
               candidate.PhysicalInstanceKey, _
                   vbTextCompare) = 0 Then Exit Sub
    Next i

    mManufacturingCandidates.Add candidate
End Sub

Public Property Get ManufacturingCandidateCount() As Long
    ManufacturingCandidateCount = mManufacturingCandidates.Count
End Property

Public Function ManufacturingCandidateAt( _
    ByVal candidateIndex As Long) As CHoleCandidate

    If candidateIndex < 1 Or _
       candidateIndex > mManufacturingCandidates.Count Then Exit Function

    Set ManufacturingCandidateAt = mManufacturingCandidates(candidateIndex)
End Function

Public Sub AddTitlePropertyEvidence( _
    ByVal propertyName As String, _
    ByVal sourceDescription As String, _
    ByVal linkedToken As String, _
    ByVal renderedValue As String, _
    ByVal targetRegion As String, _
    ByVal extentText As String, _
    ByVal isRequired As Boolean, _
    ByVal verified As Boolean)

    mTitlePropertyRecords.Add _
        "TITLE_PROPERTY|name=" & SanitizeEvidenceValue(propertyName) & _
        "|source=" & SanitizeEvidenceValue(sourceDescription) & _
        "|link=" & SanitizeEvidenceValue(linkedToken) & _
        "|rendered=" & SanitizeEvidenceValue(renderedValue) & _
        "|region=" & SanitizeEvidenceValue(targetRegion) & _
        "|extent=" & SanitizeEvidenceValue(extentText) & _
        "|required=" & CStr(isRequired) & _
        "|verified=" & CStr(verified)
End Sub

Public Sub RecordSolidWorksMutation(ByVal operationName As String)
    SolidWorksMutationSequence = SolidWorksMutationSequence + 1
    LastSolidWorksMutation = operationName

    If FinalCleanupVerified Then
        FinalCleanupVerified = False
        FinalCleanupInvalidated = True
        MarkStageFailed "FINAL_CLEANUP", _
            "post-cleanup mutation=" & operationName
        AddFailure "FINAL_CLEANUP_INVALIDATED|mutation=" & _
            SanitizeEvidenceValue(operationName)
    End If
End Sub

Public Sub AddInfo(ByVal message As String)
    If Len(Trim$(message)) = 0 Then Exit Sub
    mInfo.Add message
    Module21_EvidenceSink.LogLine "QA INFO: " & message
End Sub

Public Sub AddWarning(ByVal message As String)
    If Len(Trim$(message)) = 0 Then Exit Sub
    mWarnings.Add message
    Module21_EvidenceSink.LogLine "QA WARNING: " & message
End Sub

Public Sub AddFailure(ByVal message As String)
    If Len(Trim$(message)) = 0 Then Exit Sub
    mFailures.Add message
    Module21_EvidenceSink.LogLine "QA FAILURE: " & message
End Sub

Public Property Get HasFailures() As Boolean
    HasFailures = (mFailures.Count > 0)
End Property

Public Property Get FailureCount() As Long
    FailureCount = mFailures.Count
End Property

Public Property Get WarningCount() As Long
    WarningCount = mWarnings.Count
End Property

Public Function BuildText() As String
    Dim text As String
    text = "TARGET-SPEC HYBRID V2" & vbCrLf & _
           "Run: " & RunId & vbCrLf & _
           "Part: " & PartPath & vbCrLf & _
           "Configuration: " & ConfigurationName & vbCrLf & _
           "SOLIDWORKS: " & SolidWorksRevision & vbCrLf & _
           "Macro/export identity: " & MacroRevision & vbCrLf & _
           "Template: " & TemplatePath & vbCrLf & _
           "Sheet template: " & SheetTemplateName & vbCrLf & _
           "Sheet format: " & SheetFormatName & vbCrLf & _
           "Datum: " & DatumOrigin & vbCrLf & _
           "Scale: " & Format$(SheetScale, "0.###") & vbCrLf & _
           "Requested scale ratio: " & CStr(RequestedScaleNumerator) & _
           ":" & CStr(RequestedScaleDenominator) & vbCrLf & _
           "Actual sheet scale ratio: " & Format$(ActualScaleNumerator, _
               "0.###") & _
           ":" & Format$(ActualScaleDenominator, "0.###") & vbCrLf & _
           "Sheet scale readback: " & CStr(SheetScaleReadbackProven) & _
               vbCrLf & _
           "Settings: " & SettingsSummary & vbCrLf

    text = text & _
           "Sheet size m: " & Format$(SheetWidth, "0.000000") & _
           " x " & Format$(SheetHeight, "0.000000") & vbCrLf & _
           "Usable bounds m: " & Format$(UsableLeft, "0.000000") & _
           "," & Format$(UsableBottom, "0.000000") & " to " & _
           Format$(UsableRight, "0.000000") & "," & _
           Format$(UsableTop, "0.000000") & vbCrLf & _
           "Content border bounds m: " & _
           Format$(ContentBorderLeft, "0.000000") & _
           "," & Format$(ContentBorderBottom, "0.000000") & " to " & _
           Format$(ContentBorderRight, "0.000000") & "," & _
           Format$(ContentBorderTop, "0.000000") & vbCrLf

    text = text & "Title block bounds m: " & _
           Format$(TitleBlockLeft, "0.000000") & "," & _
           Format$(TitleBlockBottom, "0.000000") & " to " & _
           Format$(TitleBlockRight, "0.000000") & "," & _
           Format$(TitleBlockTop, "0.000000") & vbCrLf
    text = text & "Part identification bounds m: " & _
           Format$(PartIdentificationLeft, "0.000000") & "," & _
           Format$(PartIdentificationBottom, "0.000000") & " to " & _
           Format$(PartIdentificationRight, "0.000000") & "," & _
           Format$(PartIdentificationTop, "0.000000") & "; proved=" & _
           CStr(PartIdentificationBoundsProven) & vbCrLf
    text = text & "Measured layout boundaries: " & _
           CStr(LayoutBoundariesProven) & vbCrLf
    text = text & "Final selection count: " & CStr(FinalSelectionCount) & _
           vbCrLf
    text = text & "Sheet context restored: " & CStr(SheetContextRestored) & _
           vbCrLf
    text = text & "Final cleanup verified: " & CStr(FinalCleanupVerified) & _
           vbCrLf
    text = text & "Final cleanup invalidated: " & _
           CStr(FinalCleanupInvalidated) & vbCrLf
    text = text & "SOLIDWORKS mutation sequence: " & _
           CStr(SolidWorksMutationSequence) & vbCrLf
    text = text & "Last SOLIDWORKS mutation: " & LastSolidWorksMutation & _
           vbCrLf
    text = text & "Evidence report: " & EvidenceReportPath & vbCrLf
    text = text & "Evidence write attempted: " & _
           CStr(EvidenceWriteAttempted) & vbCrLf
    text = text & "Evidence write succeeded: " & _
           CStr(EvidenceWriteSucceeded) & vbCrLf
    text = text & "Evidence write verified: " & _
           CStr(EvidenceWriteVerified) & vbCrLf & vbCrLf

    text = text & "COUNTS" & vbCrLf & _
           "Views created: " & CStr(ViewsCreated) & vbCrLf & _
           "Annotations imported: " & CStr(ImportedAnnotations) & vbCrLf & _
           "Visible display dimensions: " & CStr(ImportedDisplayDimensions) _
               & vbCrLf & _
           "Owned hole candidates: " & CStr(CandidatesAccepted) & vbCrLf & _
           "Unique physical locations: " & _
               CStr(UniquePhysicalLocationCount) & vbCrLf & _
           "Rejected circular entities: " & CStr(CandidatesRejected) & _
               vbCrLf & _
           "Horizontal ordinate groups: " & CStr(HorizontalOrdinateGroups) _
               & vbCrLf & _
           "Vertical ordinate groups: " & CStr(VerticalOrdinateGroups) & _
               vbCrLf & _
           "Ordinate feature selections: " & CStr(OrdinateEntitiesSelected) _
               & vbCrLf & _
           "Layout moves: " & CStr(LayoutMoves) & vbCrLf & _
           "Title properties written: " & CStr(TitlePropertiesWritten) & _
               vbCrLf & _
           "Manufacturing candidates retained: " & _
           CStr(ManufacturingCandidateCount) & vbCrLf

    text = text & BuildStageText()
    text = text & BuildCollectionText("PHYSICAL LOCATIONS", _
        mPhysicalLocationRecords)
    text = text & BuildCollectionText("PROJECTED LOCATIONS", _
        mProjectionLocationRecords)
    text = text & BuildCollectionText("TITLE PROPERTY EVIDENCE", _
        mTitlePropertyRecords)
    text = text & BuildCollectionText("FAILURES", mFailures)
    text = text & BuildCollectionText("WARNINGS", mWarnings)
    text = text & BuildCollectionText("EVIDENCE", mInfo)

    If mFailures.Count > 0 Then
        text = text & vbCrLf & "RESULT: FAIL"
    ElseIf mWarnings.Count > 0 Then
        text = text & vbCrLf & _
            "RESULT: PASS WITH WARNINGS - VISUAL ACCEPTANCE STILL REQUIRED"
    Else
        text = text & vbCrLf & "RESULT: PASS"
    End If

    BuildText = text
End Function

Private Function BuildStageText() As String
    Dim text As String
    text = vbCrLf & "REQUIRED STAGES (" & _
        CStr(mStageRequired.Count) & ")" & vbCrLf

    Dim key As Variant
    For Each key In mStageStatus.Keys
        text = text & "- STAGE|name=" & CStr(key) & _
            "|required=" & CStr(mStageRequired.Exists(CStr(key))) & _
            "|status=" & StageStatusText(CStr(key)) & _
            "|detail=" & StageDetailText(CStr(key)) & vbCrLf
    Next key

    BuildStageText = text
End Function

Private Function StageStatusText(ByVal key As String) As String
    If mStageStatus.Exists(key) Then
        StageStatusText = CStr(mStageStatus(key))
    Else
        StageStatusText = "NOT_STARTED"
    End If
End Function

Private Function StageDetailText(ByVal key As String) As String
    If mStageDetail.Exists(key) Then
        StageDetailText = CStr(mStageDetail(key))
    End If
End Function

Private Sub AppendStageDetail( _
    ByVal key As String, _
    ByVal detail As String)

    Dim cleanDetail As String
    cleanDetail = SanitizeEvidenceValue(detail)

    If Len(cleanDetail) = 0 Then Exit Sub

    If Len(StageDetailText(key)) = 0 Then
        mStageDetail(key) = cleanDetail
    ElseIf InStr(1, StageDetailText(key), cleanDetail, vbTextCompare) = 0 Then
        mStageDetail(key) = StageDetailText(key) & "; " & cleanDetail
    End If
End Sub

Private Function NormalizeStageName(ByVal stageName As String) As String
    NormalizeStageName = UCase$(Trim$(stageName))
    NormalizeStageName = Replace$(NormalizeStageName, " ", "_")
End Function

Private Function SanitizeEvidenceValue(ByVal value As String) As String
    SanitizeEvidenceValue = Replace$(value, "|", "/")
    SanitizeEvidenceValue = Replace$(SanitizeEvidenceValue, vbCr, " ")
    SanitizeEvidenceValue = Replace$(SanitizeEvidenceValue, vbLf, " ")
    SanitizeEvidenceValue = Trim$(SanitizeEvidenceValue)
End Function

Private Function CreateTextDictionary() As Object
    Dim dictionary As Object
    Set dictionary = CreateObject("Scripting.Dictionary")
    dictionary.CompareMode = vbTextCompare
    Set CreateTextDictionary = dictionary
End Function

Private Function BuildCollectionText( _
    ByVal heading As String, _
    ByVal items As Collection) As String

    Dim text As String
    text = vbCrLf & heading & " (" & CStr(items.Count) & ")" & vbCrLf

    Dim i As Long
    For i = 1 To items.Count
        text = text & "- " & CStr(items(i)) & vbCrLf
    Next i

    BuildCollectionText = text
End Function

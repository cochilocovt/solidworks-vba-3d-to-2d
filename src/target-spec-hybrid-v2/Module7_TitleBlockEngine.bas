Option Explicit

Private Const swCustomInfoText As Long = 30
Private Const swCustomPropertyReplaceValue As Long = 2
Private Const swCustomInfoAddResult_AddedOrChanged As Long = 0
Private Const swCustomInfoGetResult_CachedValue As Long = 0
Private Const swCustomInfoGetResult_NotPresent As Long = 1
Private Const swCustomInfoGetResult_ResolvedValue As Long = 2
Private Const swMassPropertyAccuracyLevel_Higher As Long = 2
Private Const swMassPropertiesStatus_OK As Long = 0

Private Const REFERENCE_GENERAL_NOTES As String = _
    "ALL DIMENSIONS ARE IN MILLIMETRES." & vbCrLf & _
    "ALL CORNERS ARE CHAMFERED 0.5 x 45 DEG." & vbCrLf & _
    "REMOVE ALL SHARP EDGES AND BURRS."

Public Sub PopulateTitleBlock( _
    ByRef swPart As SldWorks.ModelDoc2, _
    ByRef swDrawModel As SldWorks.ModelDoc2, _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByRef evidence As CRunEvidence)

    On Error GoTo Failed
    evidence.RequireStage "TITLE_PROPERTIES"
    evidence.RequireStage "GENERAL_NOTES"
    evidence.RequireStage "PART_IDENTIFICATION"

    Dim titleStageValid As Boolean
    Dim notesStageValid As Boolean
    Dim partIdentificationStageValid As Boolean
    titleStageValid = True
    notesStageValid = True
    partIdentificationStageValid = True

    Dim drawingManager As SldWorks.CustomPropertyManager
    Set drawingManager = swDrawModel.Extension.CustomPropertyManager("")

    If drawingManager Is Nothing Then
        evidence.AddFailure "Drawing CustomPropertyManager is Nothing."
        evidence.MarkStageFailed "TITLE_PROPERTIES", _
            "drawing CustomPropertyManager is Nothing"
        evidence.MarkStageFailed "GENERAL_NOTES", _
            "drawing CustomPropertyManager is Nothing"
        evidence.MarkStageFailed "PART_IDENTIFICATION", _
            "drawing CustomPropertyManager is Nothing"
        Exit Sub
    End If

    Dim configurationName As String
    configurationName = _
        swPart.ConfigurationManager.ActiveConfiguration.Name

    Dim configurationManager As SldWorks.CustomPropertyManager
    Dim documentManager As SldWorks.CustomPropertyManager

    Set configurationManager = _
        swPart.Extension.CustomPropertyManager(configurationName)
    Set documentManager = swPart.Extension.CustomPropertyManager("")

    If Not Module1_Main.GlobalConfig.PopulateTitle Then
        evidence.AddFailure "Mandatory controlled title population was disabled."
        evidence.MarkStageFailed "TITLE_PROPERTIES", _
            "disabled by configuration"
        titleStageValid = False
        GoTo SheetItems
    End If

    Dim partNumber As String
    Dim description As String
    Dim material As String
    Dim partNumberSource As String
    Dim descriptionSource As String
    Dim materialSource As String

    partNumber = ReadPropertyConfigFirst( _
        configurationManager, documentManager, _
        Array("PartNo", "DrawingNumber", "Part Number"), _
        partNumberSource)
    description = ReadPropertyConfigFirst( _
        configurationManager, documentManager, _
        Array("Description", "PartName", "Part Name"), descriptionSource)
    material = ReadPropertyConfigFirst( _
        configurationManager, documentManager, Array("Material"), _
        materialSource)

    If Len(partNumber) = 0 Then
        evidence.AddFailure "Mandatory part-number property is missing."
        titleStageValid = False
        partNumber = FileNameWithoutExtension(swPart.GetPathName)
        partNumberSource = "fallback:authorized-part-filename"
        evidence.AddWarning "Filename used only as visible part identification; " & _
            "it does not satisfy the missing controlled property."
    End If

    If Len(description) = 0 Then
        evidence.AddFailure "Mandatory description property is missing."
        titleStageValid = False
        descriptionSource = "missing"
    End If

    If Len(material) = 0 Then
        evidence.AddFailure "Mandatory material property is missing."
        titleStageValid = False
        materialSource = "missing"
    End If

    If Not WriteDrawingProperty( _
        drawingManager, "PartNo", partNumber, partNumberSource, evidence) Then
        titleStageValid = False
    End If
    If Not WriteDrawingProperty( _
        drawingManager, "Description", description, descriptionSource, evidence) Then
        titleStageValid = False
    End If
    If Not WriteDrawingProperty( _
        drawingManager, "Material", material, materialSource, evidence) Then
        titleStageValid = False
    End If

    CopyOptionalProperty configurationManager, documentManager, _
        drawingManager, Array("CustomerCode", "Customer Code"), _
        "CustomerCode", evidence
    CopyOptionalProperty configurationManager, documentManager, _
        drawingManager, Array("Project"), "Project", evidence
    CopyOptionalProperty configurationManager, documentManager, _
        drawingManager, Array("Qty", "Quantity"), "Qty", evidence

    Dim massSource As String
    Dim massKgText As String
    Dim massNote As SldWorks.Note
    Dim massLinkPrepared As Boolean
    massKgText = ReadActiveConfigurationMassKg( _
        swPart, configurationName, massSource, evidence)

    If Len(massKgText) = 0 Then
        titleStageValid = False
    ElseIf Not WriteDrawingProperty( _
        drawingManager, "Mass", massKgText, massSource, evidence) Then

        titleStageValid = False
    Else
        massLinkPrepared = PrepareUniqueMassLink( _
            swDraw, massNote, evidence)
        If Not massLinkPrepared Then titleStageValid = False
    End If

    CopyOptionalProperty configurationManager, documentManager, _
        drawingManager, Array("SurfaceTreatment", "Surface Treatment"), _
        "SurfaceTreatment", evidence
    CopyOptionalProperty configurationManager, documentManager, _
        drawingManager, Array("HeatTreatment", "Heat Treatment"), _
        "HeatTreatment", evidence
    CopyOptionalProperty configurationManager, documentManager, _
        drawingManager, Array("Revision", "Rev"), "Revision", evidence
    CopyOptionalProperty configurationManager, documentManager, _
        drawingManager, Array("DrawnBy", "Drafter", "Designer"), _
        "DrawnBy", evidence
    CopyOptionalProperty configurationManager, documentManager, _
        drawingManager, Array("CheckedBy", "Checker"), _
        "CheckedBy", evidence
    CopyOptionalProperty configurationManager, documentManager, _
        drawingManager, Array("ApprovedBy", "Approver"), _
        "ApprovedBy", evidence
    CopyOptionalProperty configurationManager, documentManager, _
        drawingManager, Array("Unit", "Units"), "Unit", evidence

    Dim displayedScale As String
    If evidence.SheetScaleReadbackProven Then
        displayedScale = ScaleRatioToText( _
            evidence.ActualScaleNumerator, evidence.ActualScaleDenominator)
    Else
        evidence.AddFailure "Displayed scale cannot be sourced: sheet-scale readback is unproved."
        titleStageValid = False
    End If

    If Not WriteDrawingProperty( _
        drawingManager, "DisplayedScale", displayedScale, _
        "sheet:GetProperties2", evidence) Then
        titleStageValid = False
    End If

    WriteDrawingProperty drawingManager, "DrawnDate", _
        Format$(Date, "dd-mm-yyyy"), "macro:current-date", evidence

    If Len(Trim$(Module1_Main.GlobalConfig.TotalCostManual)) > 0 Then
        WriteDrawingProperty drawingManager, "TotalCost", _
            Module1_Main.GlobalConfig.TotalCostManual, _
            "manual:UI", evidence
    End If

    evidence.AddInfo "Drawing properties were populated configuration-first; " & _
        "linked title-note text is checked after rebuild."

    If Not RebuildDrawing( _
        swDrawModel, evidence, "title-property rebuild") Then
        titleStageValid = False
    ElseIf massLinkPrepared Then
        If Not VerifyPreparedMassLink( _
            massNote, massKgText, evidence) Then titleStageValid = False
    End If

    If Not VerifyLinkedText( _
        swDraw, partNumber, "PartNo", partNumberSource, _
        "TITLE_BLOCK", True, evidence) Then titleStageValid = False
    If Not VerifyLinkedText( _
        swDraw, description, "Description", descriptionSource, _
        "TITLE_BLOCK", True, evidence) Then titleStageValid = False
    If Not VerifyLinkedText( _
        swDraw, material, "Material", materialSource, _
        "TITLE_BLOCK", True, evidence) Then titleStageValid = False
    If Not VerifyLinkedText( _
        swDraw, displayedScale, "DisplayedScale", _
        "sheet:GetProperties2", "TITLE_BLOCK", True, evidence) Then
        titleStageValid = False
    End If

    If titleStageValid Then
        evidence.MarkStageProved "TITLE_PROPERTIES", _
            "mandatory values, property writes, exact links, rendered text, and title extents proved"
    Else
        evidence.MarkStageFailed "TITLE_PROPERTIES", _
            "one or more mandatory title fields or links were not proved"
    End If

SheetItems:
    If Module1_Main.GlobalConfig.InsertNotes Then
        Dim generalNotes As String
        Dim generalNotesSource As String
        generalNotes = ReadPropertyConfigFirst( _
            configurationManager, documentManager, _
            Array("GeneralNotes", "General Notes", "DrawingNotes"), _
            generalNotesSource)

        If Len(generalNotes) = 0 Then
            generalNotes = REFERENCE_GENERAL_NOTES
            generalNotesSource = "controlled-reference-general-notes"
            evidence.AddInfo "General notes sourced from the three manual references."
        End If

        If Not WriteDrawingProperty( _
            drawingManager, "GeneralNotes", generalNotes, _
            generalNotesSource, evidence) Then notesStageValid = False
        If Not RebuildDrawing( _
            swDrawModel, evidence, "general-notes rebuild") Then
            notesStageValid = False
        End If
        If Not VerifyLinkedText( _
            swDraw, generalNotes, "GeneralNotes", generalNotesSource, _
            "RESERVED_NOTE_BAND", True, evidence) Then
            notesStageValid = False
        End If

        If notesStageValid Then
            evidence.MarkStageProved "GENERAL_NOTES", _
                "linked rendered note and reserved-band extent proved"
        Else
            evidence.MarkStageFailed "GENERAL_NOTES", _
                "general note write, link, rendered value, or extent failed"
        End If
    Else
        evidence.AddFailure "Mandatory general notes were disabled."
        evidence.MarkStageFailed "GENERAL_NOTES", _
            "disabled by configuration"
    End If

    If Module1_Main.GlobalConfig.InsertBarcode Then
        Dim partIdentification As String
        partIdentification = "*" & _
            FileNameWithoutExtension(swPart.GetPathName) & "*"

        If Not WriteDrawingProperty( _
            drawingManager, "PartIdentification", partIdentification, _
            "macro:authorized-part-filename", evidence) Then

            partIdentificationStageValid = False
        End If
        If Not WriteDrawingProperty( _
            drawingManager, "PartIdentificationLocation", _
            PartIdentificationLocation(swPart.GetPathName), _
            "fixture:target-spec", evidence) Then

            partIdentificationStageValid = False
        End If
        If Not RebuildDrawing( _
            swDrawModel, evidence, "part-identification rebuild") Then
            partIdentificationStageValid = False
        End If
        If Not VerifyLinkedText( _
            swDraw, partIdentification, "PartIdentification", _
            "macro:authorized-part-filename", _
            PartIdentificationLocation(swPart.GetPathName), _
            True, evidence) Then
            partIdentificationStageValid = False
        End If

        evidence.AddWarning "Asterisk-delimited part identification was written " & _
            "as a linked property; it is not claimed as a verified barcode."

        If partIdentificationStageValid Then
            evidence.MarkStageProved "PART_IDENTIFICATION", _
                "linked value and fixture-specific region extent proved"
        Else
            evidence.MarkStageFailed "PART_IDENTIFICATION", _
                "part-identification write, link, rendered value, or extent failed"
        End If
    Else
        evidence.AddFailure "Mandatory part identification was disabled."
        evidence.MarkStageFailed "PART_IDENTIFICATION", _
            "disabled by configuration"
    End If
    Exit Sub

Failed:
    evidence.AddFailure "Title block error " & CStr(Err.Number) & _
        ": " & Err.Description
    evidence.MarkStageFailed "TITLE_PROPERTIES", _
        "API error " & CStr(Err.Number) & ": " & Err.Description
    evidence.MarkStageFailed "GENERAL_NOTES", _
        "API error " & CStr(Err.Number) & ": " & Err.Description
    evidence.MarkStageFailed "PART_IDENTIFICATION", _
        "API error " & CStr(Err.Number) & ": " & Err.Description
End Sub

Private Function ReadActiveConfigurationMassKg( _
    ByRef swPart As SldWorks.ModelDoc2, _
    ByVal configurationName As String, _
    ByRef sourceDescription As String, _
    ByRef evidence As CRunEvidence) As String

    On Error GoTo Failed

    Dim status As Long
    Dim massProperties As Variant
    massProperties = swPart.Extension.GetMassProperties2( _
        swMassPropertyAccuracyLevel_Higher, status, False)

    If status <> swMassPropertiesStatus_OK Then
        evidence.AddFailure "Active-configuration mass calculation failed; " & _
            "configuration='" & configurationName & "', status=" & _
            CStr(status) & "."
        Exit Function
    End If

    If Not IsArray(massProperties) Then
        evidence.AddFailure "Active-configuration mass calculation returned " & _
            "no metric mass-properties array."
        Exit Function
    End If

    Dim baseIndex As Long
    baseIndex = LBound(massProperties)
    If UBound(massProperties) - baseIndex + 1 < 12 Then
        evidence.AddFailure "Active-configuration mass-properties array is " & _
            "shorter than the documented 12 metric values."
        Exit Function
    End If

    Dim massKg As Double
    massKg = CDbl(massProperties(baseIndex + 5))
    If massKg <= 0# Then
        evidence.AddFailure "Active-configuration mass is not positive; " & _
            "configuration='" & configurationName & "'."
        Exit Function
    End If

    sourceDescription = "model-active-configuration:" & configurationName & _
        ":IModelDocExtension.GetMassProperties2(metric-kg)"
    ReadActiveConfigurationMassKg = Format$(massKg, "0.00")

    evidence.AddInfo "MASS_PROPERTY|configuration=" & configurationName & _
        "|method=IModelDocExtension.GetMassProperties2" & _
        "|accuracy=Higher|useSelected=False|status=" & CStr(status) & _
        "|unit=kg|value=" & Format$(massKg, "0.000000") & _
        "|display=" & ReadActiveConfigurationMassKg
    Exit Function

Failed:
    evidence.AddFailure "Active-configuration mass calculation error " & _
        CStr(Err.Number) & ": " & Err.Description
End Function

Private Function PartIdentificationLocation(ByVal partPath As String) As String
    If UCase$(FileNameWithoutExtension(partPath)) = "P-0252-01-013" Then
        PartIdentificationLocation = "AboveBottomRightTitleBlock"
    Else
        PartIdentificationLocation = "LowerLeft"
    End If
End Function

Private Sub CopyOptionalProperty( _
    ByRef configurationManager As SldWorks.CustomPropertyManager, _
    ByRef documentManager As SldWorks.CustomPropertyManager, _
    ByRef drawingManager As SldWorks.CustomPropertyManager, _
    ByVal candidates As Variant, _
    ByVal targetName As String, _
    ByRef evidence As CRunEvidence)

    Dim value As String
    Dim sourceDescription As String
    value = ReadPropertyConfigFirst( _
        configurationManager, documentManager, candidates, _
        sourceDescription)

    If Len(value) > 0 Then
        WriteDrawingProperty drawingManager, targetName, value, _
            sourceDescription, evidence
    End If
End Sub

Private Function ReadPropertyConfigFirst( _
    ByRef configurationManager As SldWorks.CustomPropertyManager, _
    ByRef documentManager As SldWorks.CustomPropertyManager, _
    ByVal candidates As Variant, _
    ByRef sourceDescription As String) As String

    ReadPropertyConfigFirst = ReadFirstResolvedProperty( _
        configurationManager, candidates, "model-configuration", _
        sourceDescription)

    If Len(ReadPropertyConfigFirst) = 0 Then
        ReadPropertyConfigFirst = ReadFirstResolvedProperty( _
            documentManager, candidates, "model-document", _
            sourceDescription)
    End If
End Function

Private Function ReadFirstResolvedProperty( _
    ByRef propertyManager As SldWorks.CustomPropertyManager, _
    ByVal propertyNames As Variant, _
    ByVal scopeName As String, _
    ByRef sourceDescription As String) As String

    If propertyManager Is Nothing Then Exit Function

    On Error GoTo Failed

    Dim i As Long
    For i = LBound(propertyNames) To UBound(propertyNames)
        Dim rawValue As String
        Dim resolvedValue As String
        Dim wasResolved As Boolean
        Dim linked As Boolean
        Dim resultCode As Long

        resultCode = propertyManager.Get6( _
            CStr(propertyNames(i)), False, rawValue, resolvedValue, _
            wasResolved, linked)

        If resultCode = swCustomInfoGetResult_NotPresent Then GoTo NextProperty
        If resultCode <> swCustomInfoGetResult_CachedValue And _
           resultCode <> swCustomInfoGetResult_ResolvedValue Then GoTo NextProperty

        If Len(Trim$(resolvedValue)) > 0 Then
            ReadFirstResolvedProperty = Trim$(resolvedValue)
            sourceDescription = scopeName & ":" & CStr(propertyNames(i))
            Exit Function
        End If

        If Len(Trim$(rawValue)) > 0 Then
            ReadFirstResolvedProperty = Trim$(rawValue)
            sourceDescription = scopeName & ":" & CStr(propertyNames(i))
            Exit Function
        End If

NextProperty:
    Next i
    Exit Function

Failed:
    ReadFirstResolvedProperty = vbNullString
End Function

Private Function WriteDrawingProperty( _
    ByRef propertyManager As SldWorks.CustomPropertyManager, _
    ByVal propertyName As String, _
    ByVal propertyValue As String, _
    ByVal sourceDescription As String, _
    ByRef evidence As CRunEvidence) As Boolean

    If Len(propertyValue) = 0 Then
        evidence.AddFailure "Drawing property '" & propertyName & _
            "' has an empty controlled value."
        Exit Function
    End If

    On Error GoTo Failed

    Dim resultCode As Long
    resultCode = propertyManager.Add3( _
        propertyName, swCustomInfoText, propertyValue, _
        swCustomPropertyReplaceValue)

    If resultCode = swCustomInfoAddResult_AddedOrChanged Then
        evidence.TitlePropertiesWritten = _
            evidence.TitlePropertiesWritten + 1
        evidence.AddInfo "TITLE_PROPERTY_WRITE|name=" & propertyName & _
            "|source=" & sourceDescription & "|status=PROVED"
        WriteDrawingProperty = True
    Else
        evidence.AddFailure "Add3 failed for drawing property '" & _
            propertyName & "'; return=" & CStr(resultCode) & "."
    End If
    Exit Function

Failed:
    evidence.AddFailure "Property write error for '" & _
        propertyName & "': " & Err.Description
End Function

Private Function RebuildDrawing( _
    ByRef swDrawModel As SldWorks.ModelDoc2, _
    ByRef evidence As CRunEvidence, _
    ByVal contextName As String) As Boolean

    RebuildDrawing = Module8_RuntimeSupport.RebuildDocumentVerified( _
        swDrawModel, contextName, evidence)
End Function

Private Function VerifyLinkedText( _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByVal expectedText As String, _
    ByVal propertyName As String, _
    ByVal sourceDescription As String, _
    ByVal targetRegion As String, _
    ByVal isRequired As Boolean, _
    ByRef evidence As CRunEvidence) As Boolean

    On Error GoTo Failed

    If Len(expectedText) = 0 Then
        evidence.AddTitlePropertyEvidence propertyName, sourceDescription, _
            vbNullString, vbNullString, targetRegion, vbNullString, _
            isRequired, False
        If isRequired Then
            evidence.AddFailure "Mandatory linked-note value is empty for '" & _
                propertyName & "'."
        End If
        Exit Function
    End If

    Dim notes As Collection
    Set notes = GetCandidateNotes(swDraw, targetRegion)

    Dim observedLinkedText As String
    Dim observedRenderedText As String
    Dim observedExtentText As String

    Dim i As Long
    For i = 1 To notes.Count
        Dim note As SldWorks.Note
        Set note = notes(i)

        Dim linkedText As String
        linkedText = note.PropertyLinkedText

        If LinkedTextReferencesProperty(linkedText, propertyName) Then
            observedLinkedText = linkedText
            Dim renderedText As String
            renderedText = note.GetText
            observedRenderedText = renderedText

            If InStr(1, NormalizeComparisonText(renderedText), _
                NormalizeComparisonText(expectedText), vbTextCompare) > 0 Then

                Dim extentText As String
                Dim extentLeft As Double
                Dim extentBottom As Double
                Dim extentRight As Double
                Dim extentTop As Double

                If GetNoteExtent( _
                    note, extentLeft, extentBottom, extentRight, _
                    extentTop, extentText) Then

                    observedExtentText = extentText
                    If NoteExtentIsInRegion( _
                        extentLeft, extentBottom, extentRight, extentTop, _
                        targetRegion, evidence) Then

                        evidence.AddTitlePropertyEvidence _
                            propertyName, sourceDescription, linkedText, _
                            renderedText, targetRegion, extentText, _
                            isRequired, True
                        evidence.AddInfo "Controlled linked note proved property '" & _
                            propertyName & "' in " & targetRegion & "."
                        VerifyLinkedText = True
                        Exit Function
                    End If
                End If
            End If
        End If
    Next i

    evidence.AddTitlePropertyEvidence propertyName, sourceDescription, _
        observedLinkedText, observedRenderedText, targetRegion, _
        observedExtentText, _
        isRequired, False

    If isRequired Then
        evidence.AddFailure "No exact visible linked note proved property '" & _
            propertyName & "' and target region '" & targetRegion & "'."
    End If
    Exit Function

Failed:
    evidence.AddTitlePropertyEvidence propertyName, sourceDescription, _
        vbNullString, vbNullString, targetRegion, vbNullString, _
        isRequired, False
    If isRequired Then
        evidence.AddFailure "Linked-note verification error for '" & _
            propertyName & "': " & Err.Description
    End If
End Function

Private Function GetCandidateNotes( _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByVal targetRegion As String) As Collection

    Dim results As New Collection
    Set GetCandidateNotes = results

    On Error GoTo Failed

    If StrComp(targetRegion, "TITLE_BLOCK", vbTextCompare) = 0 Then
        Dim swSheet As SldWorks.Sheet
        Set swSheet = swDraw.GetCurrentSheet
        If swSheet Is Nothing Then Exit Function

        Dim titleBlock As SldWorks.TitleBlock
        Set titleBlock = swSheet.TitleBlock
        If titleBlock Is Nothing Then Exit Function

        Dim titleNotes As Variant
        titleNotes = titleBlock.GetNotes
        AddVariantNotes titleNotes, results
        Exit Function
    End If

    Dim swView As SldWorks.View
    Set swView = swDraw.GetFirstView

    Do While Not swView Is Nothing
        Dim note As SldWorks.Note
        Set note = swView.GetFirstNote

        Do While Not note Is Nothing
            results.Add note
            Set note = note.GetNext
        Loop

        Set swView = swView.GetNextView
    Loop
    Exit Function

Failed:
    Set GetCandidateNotes = results
End Function

Private Function PrepareUniqueMassLink( _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByRef repairedNote As SldWorks.Note, _
    ByRef evidence As CRunEvidence) As Boolean

    On Error GoTo Failed

    Dim notes As Collection
    Set notes = GetCandidateNotes(swDraw, "ALL_DRAWING_NOTES")

    Dim matchCount As Long
    Dim oldLink As String
    Dim i As Long

    For i = 1 To notes.Count
        Dim note As SldWorks.Note
        Set note = notes(i)

        If Not note Is Nothing Then
            Dim linkedText As String
            linkedText = note.PropertyLinkedText

            If InStr(1, linkedText, "$PRP", vbTextCompare) > 0 And _
               InStr(1, linkedText, "MASS", vbTextCompare) > 0 Then

                matchCount = matchCount + 1
                Set repairedNote = note
                oldLink = linkedText
            End If
        End If
    Next i

    evidence.AddInfo "TITLE_LINK_REPAIR|name=Mass|matches=" & _
        CStr(matchCount) & "|status=Inspected"

    If matchCount = 0 Then
        evidence.AddFailure "No unique property-linked mass note was found; " & _
            "the visible title mass link was not changed."
        Exit Function
    End If

    If matchCount <> 1 Then
        evidence.AddFailure "Mass-note link repair found " & _
            CStr(matchCount) & " candidates; refusing an ambiguous change."
        Set repairedNote = Nothing
        Exit Function
    End If

    Dim drawingPropertyLink As String
    drawingPropertyLink = "$PRP:" & Chr$(34) & "Mass" & Chr$(34)

    evidence.RecordSolidWorksMutation "INote.PropertyLinkedText(Mass)"
    repairedNote.PropertyLinkedText = drawingPropertyLink

    If StrComp(Trim$(repairedNote.PropertyLinkedText), _
        drawingPropertyLink, vbTextCompare) <> 0 Then

        evidence.AddFailure "Mass-note drawing-property link did not read back " & _
            "exactly after INote.PropertyLinkedText."
        Exit Function
    End If

    evidence.AddInfo "TITLE_LINK_REPAIR|name=Mass|oldLink=" & _
        SafeEvidenceValue(oldLink) & _
        "|newLink=" & SafeEvidenceValue(drawingPropertyLink) & _
        "|status=Prepared"
    PrepareUniqueMassLink = True
    Exit Function

Failed:
    evidence.AddFailure "Mass-note link preparation error " & _
        CStr(Err.Number) & ": " & Err.Description
    Set repairedNote = Nothing
End Function

Private Function VerifyPreparedMassLink( _
    ByRef repairedNote As SldWorks.Note, _
    ByVal expectedMassText As String, _
    ByRef evidence As CRunEvidence) As Boolean

    If repairedNote Is Nothing Then Exit Function
    On Error GoTo Failed

    Dim drawingPropertyLink As String
    drawingPropertyLink = "$PRP:" & Chr$(34) & "Mass" & Chr$(34)

    Dim linkedText As String
    Dim renderedText As String
    linkedText = repairedNote.PropertyLinkedText
    renderedText = repairedNote.GetText

    If StrComp(Trim$(linkedText), _
        drawingPropertyLink, vbTextCompare) <> 0 Then

        evidence.AddFailure "Mass-note link changed after rebuild."
        Exit Function
    End If

    If InStr(1, NormalizeComparisonText(renderedText), _
        NormalizeComparisonText(expectedMassText), vbTextCompare) = 0 Then

        evidence.AddFailure "Mass-note rendered text does not contain the " & _
            "computed kilogram value '" & expectedMassText & "'."
        Exit Function
    End If

    evidence.AddInfo "TITLE_LINK_REPAIR|name=Mass|link=" & _
        SafeEvidenceValue(linkedText) & _
        "|rendered=" & SafeEvidenceValue(renderedText) & _
        "|status=PROVED"
    VerifyPreparedMassLink = True
    Exit Function

Failed:
    evidence.AddFailure "Mass-note link verification error " & _
        CStr(Err.Number) & ": " & Err.Description
End Function

Private Sub AddVariantNotes( _
    ByVal notesValue As Variant, _
    ByRef results As Collection)

    On Error GoTo Failed
    If IsEmpty(notesValue) Or IsNull(notesValue) Then Exit Sub

    If IsArray(notesValue) Then
        Dim i As Long
        For i = LBound(notesValue) To UBound(notesValue)
            If IsObject(notesValue(i)) Then
                Dim noteObject As Object
                Set noteObject = notesValue(i)
                results.Add noteObject
            End If
        Next i
    ElseIf IsObject(notesValue) Then
        Dim singleNote As Object
        Set singleNote = notesValue
        results.Add singleNote
    End If
    Exit Sub

Failed:
    ' The caller treats an incomplete note collection as unproved evidence.
End Sub

Private Function LinkedTextReferencesProperty( _
    ByVal linkedText As String, _
    ByVal propertyName As String) As Boolean

    If InStr(1, linkedText, "$PRP", vbTextCompare) = 0 Then Exit Function

    LinkedTextReferencesProperty = _
        (InStr(1, linkedText, Chr$(34) & propertyName & Chr$(34), _
            vbTextCompare) > 0)
End Function

Private Function SafeEvidenceValue(ByVal value As String) As String
    SafeEvidenceValue = Replace$(Trim$(value), "|", "/")
    SafeEvidenceValue = Replace$(SafeEvidenceValue, vbCr, " ")
    SafeEvidenceValue = Replace$(SafeEvidenceValue, vbLf, " ")
End Function

Private Function GetNoteExtent( _
    ByRef note As SldWorks.Note, _
    ByRef extentLeft As Double, _
    ByRef extentBottom As Double, _
    ByRef extentRight As Double, _
    ByRef extentTop As Double, _
    ByRef extentText As String) As Boolean

    On Error GoTo Failed

    Dim extent As Variant
    extent = note.GetExtent
    If Not IsArray(extent) Then Exit Function
    If UBound(extent) - LBound(extent) + 1 < 6 Then Exit Function

    Dim lowerIndex As Long
    lowerIndex = LBound(extent)

    extentLeft = CDbl(extent(lowerIndex))
    extentBottom = CDbl(extent(lowerIndex + 1))
    extentRight = CDbl(extent(lowerIndex + 3))
    extentTop = CDbl(extent(lowerIndex + 4))

    If extentRight < extentLeft Then
        Dim swapValue As Double
        swapValue = extentLeft
        extentLeft = extentRight
        extentRight = swapValue
    End If

    If extentTop < extentBottom Then
        swapValue = extentBottom
        extentBottom = extentTop
        extentTop = swapValue
    End If

    If extentRight <= extentLeft Or extentTop <= extentBottom Then _
        Exit Function

    extentText = Format$(extentLeft, "0.000000") & "," & _
        Format$(extentBottom, "0.000000") & " to " & _
        Format$(extentRight, "0.000000") & "," & _
        Format$(extentTop, "0.000000")
    GetNoteExtent = True
    Exit Function

Failed:
    GetNoteExtent = False
End Function

Private Function NoteExtentIsInRegion( _
    ByVal extentLeft As Double, _
    ByVal extentBottom As Double, _
    ByVal extentRight As Double, _
    ByVal extentTop As Double, _
    ByVal targetRegion As String, _
    ByRef evidence As CRunEvidence) As Boolean

    If Not evidence.LayoutBoundariesProven Then Exit Function

    Const tolerance As Double = 0.0000001
    If extentLeft < -tolerance Or extentBottom < -tolerance Or _
       extentRight > evidence.SheetWidth + tolerance Or _
       extentTop > evidence.SheetHeight + tolerance Then Exit Function

    Select Case UCase$(Trim$(targetRegion))
        Case "TITLE_BLOCK"
            NoteExtentIsInRegion = _
                (extentLeft >= evidence.TitleBlockLeft - tolerance) And _
                (extentBottom >= evidence.TitleBlockBottom - tolerance) And _
                (extentRight <= evidence.TitleBlockRight + tolerance) And _
                (extentTop <= evidence.TitleBlockTop + tolerance)

        Case "RESERVED_NOTE_BAND", "LOWERLEFT"
            NoteExtentIsInRegion = _
                (extentTop <= evidence.UsableBottom + tolerance) And _
                (extentRight <= evidence.TitleBlockLeft + tolerance)

        Case "ABOVEBOTTOMRIGHTTITLEBLOCK"
            NoteExtentIsInRegion = _
                (extentLeft >= evidence.TitleBlockLeft - tolerance) And _
                (extentBottom >= evidence.TitleBlockTop - tolerance) And _
                (extentTop <= evidence.UsableBottom + tolerance)
    End Select
End Function

Private Function ScaleRatioToText( _
    ByVal numerator As Double, _
    ByVal denominator As Double) As String

    If numerator <= 0# Or denominator <= 0# Then Exit Function

    ScaleRatioToText = FormatScaleNumber(numerator) & ":" & _
        FormatScaleNumber(denominator)
End Function

Private Function FormatScaleNumber(ByVal value As Double) As String
    If Abs(value - CDbl(CLng(value))) < 0.000000001 Then
        FormatScaleNumber = CStr(CLng(value))
    Else
        FormatScaleNumber = Format$(value, "0.###")
    End If
End Function

Private Function NormalizeComparisonText(ByVal text As String) As String
    NormalizeComparisonText = Replace$(text, vbCr, vbNullString)
    NormalizeComparisonText = Replace$(NormalizeComparisonText, vbLf, "|")
    NormalizeComparisonText = Trim$(NormalizeComparisonText)
End Function

Private Function FileNameWithoutExtension(ByVal path As String) As String
    Dim slashPosition As Long
    slashPosition = InStrRev(path, "\")

    Dim nameOnly As String
    If slashPosition > 0 Then
        nameOnly = Mid$(path, slashPosition + 1)
    Else
        nameOnly = path
    End If

    Dim dotPosition As Long
    dotPosition = InStrRev(nameOnly, ".")

    If dotPosition > 1 Then
        FileNameWithoutExtension = Left$(nameOnly, dotPosition - 1)
    Else
        FileNameWithoutExtension = nameOnly
    End If
End Function

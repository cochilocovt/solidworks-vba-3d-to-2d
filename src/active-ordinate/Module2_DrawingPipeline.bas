Attribute VB_Name = "Module2_DrawingPipeline"
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" _
        (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" _
        (ByVal dwMilliseconds As Long)
#End If

Private Const swCreateSectionView_NotAligned As Long = 1

Private Const swDisplayMode_HiddenLinesVisible As Long = 1
Private Const swDisplayMode_HiddenLinesRemoved As Long = 3
Private Const swDisplayMode_ShadedWithEdges As Long = 6

Public Sub CreateDrawing( _
    ByRef swApp As SldWorks.SldWorks, _
    ByRef swPart As SldWorks.ModelDoc2, _
    ByVal partPath As String, _
    ByVal templatePath As String)

    RunDrawingPipeline swApp, swPart, partPath, templatePath
End Sub

Public Sub RunDrawingPipeline( _
    ByRef swApp As SldWorks.SldWorks, _
    ByRef swPart As SldWorks.ModelDoc2, _
    ByVal partPath As String, _
    ByVal templatePath As String)

    On Error GoTo FailRun

    Dim holes As Collection
    Set holes = Module3_ModelAudit.GetAllHoleLikeFeatures(swPart)

    Dim swDrawModel As SldWorks.ModelDoc2
    Set swDrawModel = swApp.NewDocument(templatePath, 0, 0, 0)

    If swDrawModel Is Nothing Then
        Err.Raise vbObjectError + 701, , "Failed to create drawing document."
    End If

    Dim swDraw As SldWorks.DrawingDoc
    Set swDraw = swDrawModel

    Dim swFrontView As SldWorks.View
    CreateViews swPart, swDrawModel, swDraw, partPath, swFrontView

    swDrawModel.ForceRebuild3 False
    DoEvents
    Sleep 300

    Dim importedModelItems As Long
    importedModelItems = 0

    If Module1_Main.GlobalConfig.UseModelDimensions Then
        If Not swFrontView Is Nothing Then
            importedModelItems = _
                Module4_ModelItemImporter.ImportModelItemsAcrossDrawing( _
                    swDrawModel, _
                    swDraw, _
                    swFrontView.Name)
        End If
    End If

    If Module1_Main.GlobalConfig.UseOrdinateDims Then
        Module5_FallbackDimensionEngine.AddOrdinateDimensionsToAllViews _
            swDrawModel, _
            swDraw, _
            Module1_Main.GlobalConfig.DatumOrigin
    End If

    If Module1_Main.GlobalConfig.AutoArrange Then
        Module4_ModelItemImporter.AutoArrangeAllDrawingDimensions _
            swDrawModel, _
            swDraw
    End If

    If Module1_Main.GlobalConfig.PopulateTitle Then
        Module7_TitleBlockEngine.PopulateTitleBlock _
            swPart, _
            swDrawModel, _
            swDraw
    End If

    swDrawModel.ForceRebuild3 False
    swDrawModel.ViewZoomtofit2

    If Module1_Main.GlobalConfig.GenerateQAReport Then
        MsgBox Module6_QAEngine.BuildRunSummary( _
            swDrawModel, _
            swDraw, _
            holes, _
            importedModelItems), _
            vbInformation, _
            "Drawing QA Summary"
    End If

    Exit Sub

FailRun:
    MsgBox "Drawing pipeline failed: " & Err.Description, _
           vbCritical, _
           "Pipeline Error"
End Sub

Private Sub CreateViews( _
    ByRef swPart As SldWorks.ModelDoc2, _
    ByRef swDrawModel As SldWorks.ModelDoc2, _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByVal partPath As String, _
    ByRef swFrontView As SldWorks.View)

    On Error GoTo FailCreateViews

    Dim swPartDoc As SldWorks.PartDoc
    Set swPartDoc = swPart

    Dim bbox As Variant
    bbox = swPartDoc.GetPartBox(True)

    Dim partW As Double
    Dim partH As Double

    partW = Abs(CDbl(bbox(3)) - CDbl(bbox(0)))
    partH = Abs(CDbl(bbox(4)) - CDbl(bbox(1)))

    Dim swSheet As SldWorks.Sheet
    Set swSheet = swDraw.GetCurrentSheet

    If swSheet Is Nothing Then
        Err.Raise vbObjectError + 702, , "Could not obtain current drawing sheet."
    End If

    Dim sheetW As Double
    Dim sheetH As Double
    Dim paperSize As Long

    paperSize = swSheet.GetSize(sheetW, sheetH)

    Dim scaleVal As Double
    scaleVal = Module1_Main.GlobalConfig.SheetScale

    If scaleVal <= 0# Then scaleVal = 1#

    Dim displayMode As Long

    If Module1_Main.GlobalConfig.UseHLR Then
        displayMode = swDisplayMode_HiddenLinesRemoved
    Else
        displayMode = swDisplayMode_HiddenLinesVisible
    End If

    Dim xFront As Double
    Dim yFront As Double

    xFront = sheetW * 0.42
    yFront = sheetH * 0.48

    Set swFrontView = swDraw.CreateDrawViewFromModelView3( _
        partPath, _
        "*Front", _
        xFront, _
        yFront, _
        0#)

    If swFrontView Is Nothing Then
        Err.Raise vbObjectError + 711, , "Failed to create front view."
    End If

    ConfigureView swFrontView, displayMode, scaleVal

    If Module1_Main.GlobalConfig.CreateTop Then
        CreateNamedView swDraw, partPath, "*Top", _
                        xFront, _
                        yFront + (partH * scaleVal * 0.75), _
                        displayMode, scaleVal
    End If

    If Module1_Main.GlobalConfig.CreateBottom Then
        CreateNamedView swDraw, partPath, "*Bottom", _
                        xFront, _
                        yFront - (partH * scaleVal * 0.75), _
                        displayMode, scaleVal
    End If

    If Module1_Main.GlobalConfig.CreateRight Then
        CreateNamedView swDraw, partPath, "*Right", _
                        xFront + (partW * scaleVal * 0.55), _
                        yFront, _
                        displayMode, scaleVal
    End If

    If Module1_Main.GlobalConfig.CreateLeft Then
        CreateNamedView swDraw, partPath, "*Left", _
                        xFront - (partW * scaleVal * 0.55), _
                        yFront, _
                        displayMode, scaleVal
    End If

    If Module1_Main.GlobalConfig.CreateBack Then
        CreateNamedView swDraw, partPath, "*Back", _
                        xFront + (partW * scaleVal * 0.85), _
                        yFront, _
                        displayMode, scaleVal
    End If

    If Module1_Main.GlobalConfig.CreateIso Then
        CreateNamedView swDraw, partPath, "*Isometric", _
                        sheetW * 0.8, _
                        sheetH * 0.72, _
                        swDisplayMode_ShadedWithEdges, _
                        scaleVal * 0.55
    End If

    swDrawModel.ForceRebuild3 False
    DoEvents
    Sleep 300

    If Module1_Main.GlobalConfig.CreateSection Then
        CreateSectionFromConfig swDrawModel, swDraw, swFrontView, displayMode, scaleVal
    End If

    Exit Sub

FailCreateViews:
    Err.Raise Err.Number, , Err.Description
End Sub

Private Sub CreateNamedView( _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByVal partPath As String, _
    ByVal modelViewName As String, _
    ByVal xPos As Double, _
    ByVal yPos As Double, _
    ByVal displayMode As Long, _
    ByVal scaleVal As Double)

    Dim swView As SldWorks.View

    Set swView = swDraw.CreateDrawViewFromModelView3( _
        partPath, _
        modelViewName, _
        xPos, _
        yPos, _
        0#)

    If Not swView Is Nothing Then
        ConfigureView swView, displayMode, scaleVal
    End If
End Sub

Private Sub ConfigureView( _
    ByRef swView As SldWorks.View, _
    ByVal displayMode As Long, _
    ByVal scaleVal As Double)

    On Error Resume Next

    If swView Is Nothing Then Exit Sub

    swView.ScaleDecimal = scaleVal
    swView.SetDisplayMode3 False, displayMode, False, True
    swView.SetLightweightToResolved
End Sub

Private Sub CreateSectionFromConfig( _
    ByRef swDrawModel As SldWorks.ModelDoc2, _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByRef swFrontView As SldWorks.View, _
    ByVal displayMode As Long, _
    ByVal scaleVal As Double)

    On Error GoTo SafeExit

    If swDrawModel Is Nothing Then Exit Sub
    If swDraw Is Nothing Then Exit Sub
    If swFrontView Is Nothing Then Exit Sub

    Dim secLabel As String
    Dim secVertical As Boolean

    GetPrimarySectionSettings secLabel, secVertical

    If Len(secLabel) = 0 Then secLabel = "J"

    swDraw.ActivateView swFrontView.Name
    swDrawModel.ClearSelection2 True

    Dim outline As Variant
    outline = swFrontView.GetOutline

    If IsEmpty(outline) Then GoTo SafeExit

    Dim leftX As Double
    Dim bottomY As Double
    Dim rightX As Double
    Dim topY As Double

    leftX = CDbl(outline(0))
    bottomY = CDbl(outline(1))
    rightX = CDbl(outline(2))
    topY = CDbl(outline(3))

    Dim width As Double
    Dim height As Double

    width = rightX - leftX
    height = topY - bottomY

    If width <= 0# Or height <= 0# Then GoTo SafeExit

    Dim midX As Double
    Dim midY As Double

    midX = (leftX + rightX) / 2#
    midY = (bottomY + topY) / 2#

    Dim extension As Double
    extension = 0.02

    Dim swLine As SldWorks.SketchLine

    If secVertical Then
        Set swLine = swDrawModel.SketchManager.CreateLine( _
            midX, _
            bottomY - extension, _
            0#, _
            midX, _
            topY + extension, _
            0#)
    Else
        Set swLine = swDrawModel.SketchManager.CreateLine( _
            leftX - extension, _
            midY, _
            0#, _
            rightX + extension, _
            midY, _
            0#)
    End If

    If swLine Is Nothing Then GoTo SafeExit

    swDrawModel.ClearSelection2 True

    If Not swLine.Select4(True, Nothing) Then GoTo SafeExit

    Dim targetX As Double
    Dim targetY As Double

    targetX = rightX + 0.08
    targetY = midY

    Dim excludedComponents As Variant
    excludedComponents = Empty

    Dim swSectionView As SldWorks.View

    Set swSectionView = swDraw.CreateSectionViewAt5( _
        targetX, _
        targetY, _
        0#, _
        secLabel, _
        swCreateSectionView_NotAligned, _
        excludedComponents, _
        0#)

    If Not swSectionView Is Nothing Then
        ConfigureView swSectionView, displayMode, scaleVal
    End If

SafeExit:
    On Error Resume Next
    swDrawModel.ClearSelection2 True
    swDraw.ActivateView ""
    On Error GoTo 0
End Sub

Private Sub GetPrimarySectionSettings( _
    ByRef secLabel As String, _
    ByRef secVertical As Boolean)

    secLabel = vbNullString
    secVertical = False

    If Module1_Main.GlobalSectionCount > 0 Then
        secLabel = Module1_Main.GetSectionLabelOrDefault(1)
        secVertical = Module1_Main.GlobalSections(1).Vertical
    Else
        secLabel = UCase$(Left$(Trim$(Module1_Main.GlobalConfig.SectionLabel), 2))

        If Len(secLabel) = 0 Then secLabel = "J"

        secVertical = Module1_Main.GlobalConfig.SectionVertical
    End If
End Sub


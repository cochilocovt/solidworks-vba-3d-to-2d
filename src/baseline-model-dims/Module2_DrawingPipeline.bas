Attribute VB_Name = "Module2_DrawingPipeline"
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Private Const swCreateSectionView_NotAligned As Long = 1
Private Const swDisplayMode_HiddenLinesRemoved As Long = 3
Private Const swDisplayMode_HiddenLinesVisible As Long = 1
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
        Err.Raise vbObjectError + 701, , "Failed to create drawing document"
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
        importedModelItems = Module4_ModelItemImporter.ImportModelItemsAcrossDrawing(swDrawModel, swDraw, swFrontView.Name)
    End If

    If Module1_Main.GlobalConfig.UseOrdinateDims Then
        AddFallbackOrdinateDimensions swDrawModel, swDraw
    End If

    If Module1_Main.GlobalConfig.AutoArrange Then
        Module4_ModelItemImporter.AutoArrangeAllDrawingDimensions swDrawModel, swDraw
    End If

    If Module1_Main.GlobalConfig.PopulateTitle Then
        Module7_TitleBlockEngine.PopulateTitleBlock swPart, swDrawModel, swDraw
    End If

    swDrawModel.ForceRebuild3 False
    swDrawModel.ViewZoomtofit2

    MsgBox Module6_QAEngine.BuildRunSummary(swDrawModel, swDraw, holes, importedModelItems), vbInformation, "Drawing QA Summary"
    Exit Sub

FailRun:
    MsgBox "Drawing pipeline failed: " & Err.Description, vbCritical, "Pipeline Error"
End Sub

Private Sub AddFallbackOrdinateDimensions(ByRef swDrawModel As SldWorks.ModelDoc2, ByRef swDraw As SldWorks.DrawingDoc)
    On Error GoTo SafeExit

    Dim swView As SldWorks.View
    Set swView = swDraw.GetFirstView
    If Not swView Is Nothing Then Set swView = swView.GetNextView

    Do While Not swView Is Nothing
        If Not IsIsoView(swView.Name) Then
            Module5_FallbackDimensionEngine.CreateHoleOrdinateDims swDraw, swView, Module1_Main.GlobalConfig.DatumOrigin
        End If
        Set swView = swView.GetNextView
    Loop

SafeExit:
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
    partW = Abs(bbox(3) - bbox(0))
    partH = Abs(bbox(4) - bbox(1))

    Dim swSheet As SldWorks.Sheet
    Set swSheet = swDraw.GetCurrentSheet

    Dim sheetW As Double
    Dim sheetH As Double
    swSheet.GetSize sheetW, sheetH

    Dim scaleVal As Double
    scaleVal = Module1_Main.GlobalConfig.SheetScale
    If scaleVal <= 0 Then scaleVal = 1#

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

    Set swFrontView = swDraw.CreateDrawViewFromModelView3(partPath, "*Front", xFront, yFront, 0)
    If swFrontView Is Nothing Then
        Err.Raise vbObjectError + 711, , "Failed to create front view"
    End If
    ConfigureView swFrontView, displayMode, scaleVal

    If Module1_Main.GlobalConfig.CreateTop Then
        CreateNamedView swDraw, partPath, "*Top", xFront, yFront + (partH * scaleVal) * 0.75, displayMode, scaleVal
    End If

    If Module1_Main.GlobalConfig.CreateBottom Then
        CreateNamedView swDraw, partPath, "*Bottom", xFront, yFront - (partH * scaleVal) * 0.75, displayMode, scaleVal
    End If

    If Module1_Main.GlobalConfig.CreateRight Then
        CreateNamedView swDraw, partPath, "*Right", xFront + (partW * scaleVal) * 0.55, yFront, displayMode, scaleVal
    End If

    If Module1_Main.GlobalConfig.CreateLeft Then
        CreateNamedView swDraw, partPath, "*Left", xFront - (partW * scaleVal) * 0.55, yFront, displayMode, scaleVal
    End If

    If Module1_Main.GlobalConfig.CreateBack Then
        CreateNamedView swDraw, partPath, "*Back", xFront + (partW * scaleVal) * 0.85, yFront, displayMode, scaleVal
    End If

    If Module1_Main.GlobalConfig.CreateIso Then
        CreateNamedView swDraw, partPath, "*Isometric", sheetW * 0.8, sheetH * 0.72, swDisplayMode_ShadedWithEdges, scaleVal * 0.55
    End If

    swDrawModel.ForceRebuild3 False
    DoEvents
    Sleep 300

    If Module1_Main.GlobalConfig.CreateSection Then
        CreateSectionFromConfig swDrawModel, swDraw, bbox, swFrontView, displayMode, scaleVal
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
    Set swView = swDraw.CreateDrawViewFromModelView3(partPath, modelViewName, xPos, yPos, 0)

    If Not swView Is Nothing Then
        ConfigureView swView, displayMode, scaleVal
    End If
End Sub

Private Sub ConfigureView(ByRef swView As SldWorks.View, ByVal displayMode As Long, ByVal scaleVal As Double)
    On Error Resume Next

    swView.ScaleDecimal = scaleVal
    swView.SetDisplayMode3 False, displayMode, False, True
    swView.SetLightweightToResolved
End Sub

Private Sub CreateSectionFromConfig( _
    ByRef swDrawModel As SldWorks.ModelDoc2, _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByVal bbox As Variant, _
    ByRef swFrontView As SldWorks.View, _
    ByVal displayMode As Long, _
    ByVal scaleVal As Double)

    On Error GoTo SafeExit

    If swFrontView Is Nothing Then Exit Sub

    swDraw.ActivateView swFrontView.Name

    Dim midX As Double
    Dim midY As Double
    midX = (bbox(0) + bbox(3)) / 2#
    midY = (bbox(1) + bbox(4)) / 2#

    Dim spanX As Double
    Dim spanY As Double
    spanX = Abs(bbox(3) - bbox(0))
    spanY = Abs(bbox(4) - bbox(1))

    Dim secLabel As String
    Dim secVertical As Boolean
    GetPrimarySectionSettings secLabel, secVertical

    Dim swLine As SldWorks.SketchLine
    If secVertical Then
        Set swLine = swDrawModel.SketchManager.CreateLine(midX, bbox(1) - spanY * 0.15, 0, midX, bbox(4) + spanY * 0.15, 0)
    Else
        Set swLine = swDrawModel.SketchManager.CreateLine(bbox(0) - spanX * 0.15, midY, 0, bbox(3) + spanX * 0.15, midY, 0)
    End If

    If swLine Is Nothing Then GoTo SafeExit

    swDrawModel.ClearSelection2 True
    swLine.Select4 True, Nothing

    Dim frontPos As Variant
    frontPos = swFrontView.Position

    Dim targetX As Double
    Dim targetY As Double
    targetX = frontPos(0) + 0.18
    targetY = frontPos(1)

    Dim swSectionView As SldWorks.View
    Set swSectionView = swDraw.CreateSectionViewAt5(targetX, targetY, 0, secLabel, swCreateSectionView_NotAligned, Nothing, 0)

    If Not swSectionView Is Nothing Then
        ConfigureView swSectionView, displayMode, scaleVal
    End If

SafeExit:
    swDrawModel.ClearSelection2 True
    swDraw.ActivateView ""
End Sub

Private Sub GetPrimarySectionSettings(ByRef secLabel As String, ByRef secVertical As Boolean)
    secLabel = vbNullString
    secVertical = False

    If Module1_Main.GlobalSectionCount > 0 Then
        secLabel = Module1_Main.GetSectionLabelOrDefault(1)
        secVertical = Module1_Main.GlobalSections(1).Vertical
    Else
        secLabel = Trim$(Module1_Main.GlobalConfig.SectionLabel)
        If Len(secLabel) = 0 Then secLabel = "J"
        secVertical = Module1_Main.GlobalConfig.SectionVertical
    End If
End Sub

Private Function IsIsoView(ByVal viewName As String) As Boolean
    IsIsoView = (InStr(1, viewName, "ISO", vbTextCompare) > 0 Or InStr(1, viewName, "ISOMETRIC", vbTextCompare) > 0)
End Function


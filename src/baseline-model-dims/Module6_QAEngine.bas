Attribute VB_Name = "Module6_QAEngine"

Option Explicit

Public Function BuildRunSummary( _
    ByRef swDrawModel As SldWorks.ModelDoc2, _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByVal holes As Collection, _
    ByVal importedModelItems As Long) As String

    Dim report As String
    Dim totalDims As Long
    totalDims = CountAllViewDimensions(swDraw)

    report = "Drawing QA Summary" & vbCrLf & String(26, "-") & vbCrLf
    report = report & "Detected hole-like features: " & Module3_ModelAudit.CountHoles(holes) & vbCrLf
    report = report & "Imported model items: " & importedModelItems & vbCrLf
    report = report & "Total drawing view dimensions: " & totalDims & vbCrLf
    report = report & BuildPerViewSummary(swDraw)

    If importedModelItems = 0 Then
        report = report & "WARNING: Model item import returned zero items." & vbCrLf
    End If

    If totalDims = 0 Then
        report = report & "WARNING: No drawing dimensions were found. Review model-item import and fallback logic." & vbCrLf
    ElseIf totalDims < 6 Then
        report = report & "WARNING: Drawing has very few dimensions. Manual review required." & vbCrLf
    Else
        report = report & "PASS: Drawing contains dimensions." & vbCrLf
    End If

    BuildRunSummary = report
End Function

Public Function CountAllViewDimensions(ByRef swDraw As SldWorks.DrawingDoc) As Long
    Dim swView As SldWorks.View
    Set swView = swDraw.GetFirstView
    If Not swView Is Nothing Then Set swView = swView.GetNextView

    Do While Not swView Is Nothing
        CountAllViewDimensions = CountAllViewDimensions + Module4_ModelItemImporter.CountDisplayDimensionsInView(swView)
        Set swView = swView.GetNextView
    Loop
End Function

Private Function BuildPerViewSummary(ByRef swDraw As SldWorks.DrawingDoc) As String
    Dim swView As SldWorks.View
    Set swView = swDraw.GetFirstView
    If Not swView Is Nothing Then Set swView = swView.GetNextView

    Do While Not swView Is Nothing
        BuildPerViewSummary = BuildPerViewSummary & swView.Name & ": " & Module4_ModelItemImporter.CountDisplayDimensionsInView(swView) & " dims" & vbCrLf
        Set swView = swView.GetNextView
    Loop
End Function



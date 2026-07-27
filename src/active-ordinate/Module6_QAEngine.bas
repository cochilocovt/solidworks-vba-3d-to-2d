Attribute VB_Name = "Module6_QAEngine"
Option Explicit

Public Function BuildRunSummary( _
    ByRef swDrawModel As SldWorks.ModelDoc2, _
    ByRef swDraw As SldWorks.DrawingDoc, _
    ByVal holes As Collection, _
    ByVal importedModelItems As Long) As String

    Dim report As String
    report = "Drawing QA Summary" & vbCrLf
    report = report & "-------------------" & vbCrLf

    Dim holeLikeCount As Long
    holeLikeCount = Module3_ModelAudit.CountHoles(holes)

    report = report & "Detected hole-like feature records: " & _
             CStr(holeLikeCount) & vbCrLf
    report = report & "Imported model items: " & _
             CStr(importedModelItems) & vbCrLf

    Dim totalDims As Long
    totalDims = CountAllViewDimensions(swDraw)
    report = report & "Total drawing view dimensions: " & _
             CStr(totalDims) & vbCrLf

    Dim swView As SldWorks.View
    Set swView = swDraw.GetFirstView

    If Not swView Is Nothing Then
        Set swView = swView.GetNextView
    End If

    Do While Not swView Is Nothing
        Dim dimCount As Long
        dimCount = Module4_ModelItemImporter.CountDisplayDimensionsInView(swView)

        report = report & "View '" & swView.Name & "': " & _
                 CStr(dimCount) & " dims" & vbCrLf

        Set swView = swView.GetNextView
    Loop

    report = report & vbCrLf

    If importedModelItems <= 0 Then
        report = report & _
            "FAIL: Model annotation import returned zero items." & vbCrLf
    End If

    If totalDims <= 0 Then
        report = report & _
            "FAIL: No drawing dimensions were created." & vbCrLf
    End If

    If importedModelItems > 0 And totalDims > 0 Then
        report = report & _
            "PASS WITH WARNINGS: model annotations and drawing dimensions exist." & vbCrLf
    End If

    report = report & _
        "NOT VERIFIED BY VBA QA: feature-proven hole ownership, actual datum entity, " & _
        "ordinate API result per group, duplicate model/ordinate coverage, view eligibility, " & _
        "annotation collisions, sheet bounds, title-block bounds, and linked properties."

    BuildRunSummary = report
End Function

Private Function CountAllViewDimensions( _
    ByRef swDraw As SldWorks.DrawingDoc) As Long

    If swDraw Is Nothing Then Exit Function

    Dim total As Long
    total = 0

    Dim swView As SldWorks.View
    Set swView = swDraw.GetFirstView

    If Not swView Is Nothing Then
        Set swView = swView.GetNextView
    End If

    Do While Not swView Is Nothing
        total = total + _
            Module4_ModelItemImporter.CountDisplayDimensionsInView(swView)
        Set swView = swView.GetNextView
    Loop

    CountAllViewDimensions = total
End Function


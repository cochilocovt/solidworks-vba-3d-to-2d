Attribute VB_Name = "Module7_TitleBlockEngine"
Option Explicit

Private Const swCustomInfoText As Long = 30

Public Sub PopulateTitleBlock(ByRef swPart As SldWorks.ModelDoc2, ByRef swDrawModel As SldWorks.ModelDoc2, ByRef swDraw As SldWorks.DrawingDoc)
    On Error GoTo Failed

    Dim swPartMgr As SldWorks.CustomPropertyManager
    Dim swDrawMgr As SldWorks.CustomPropertyManager

    Set swPartMgr = swPart.extension.CustomPropertyManager("")
    Set swDrawMgr = swDrawModel.extension.CustomPropertyManager("")

    If swPartMgr Is Nothing Or swDrawMgr Is Nothing Then Exit Sub

    CopyFirstResolvedProp swPartMgr, swDrawMgr, Array("Description", "PartName"), "Description"
    CopyFirstResolvedProp swPartMgr, swDrawMgr, Array("PartNo", "DrawingNumber", "Part Number"), "PartNo"
    CopyFirstResolvedProp swPartMgr, swDrawMgr, Array("Material"), "Material"
    CopyFirstResolvedProp swPartMgr, swDrawMgr, Array("CustomerCode", "Customer Code"), "CustomerCode"
    CopyFirstResolvedProp swPartMgr, swDrawMgr, Array("Project"), "Project"
    CopyFirstResolvedProp swPartMgr, swDrawMgr, Array("Qty", "Quantity"), "Qty"
    CopyFirstResolvedProp swPartMgr, swDrawMgr, Array("Mass", "Weight"), "Mass"

    WriteDrawingProperty swDrawMgr, "DrawnDate", Format$(Date, "dd-mm-yyyy")

    If Len(Trim$(Module1_Main.GlobalConfig.TotalCostManual)) > 0 Then
        WriteDrawingProperty swDrawMgr, "TotalCost", Module1_Main.GlobalConfig.TotalCostManual
    End If

    If Module1_Main.GlobalConfig.InsertNotes Then InsertGeneralNotes swDraw
    If Module1_Main.GlobalConfig.InsertBarcode Then InsertBarcode swPart, swDraw
    Exit Sub

Failed:
    Debug.Print "Title block warning: " & Err.Description
End Sub

Private Sub CopyFirstResolvedProp(ByRef srcMgr As SldWorks.CustomPropertyManager, ByRef dstMgr As SldWorks.CustomPropertyManager, ByVal candidates As Variant, ByVal targetName As String)
    Dim valText As String
    valText = ReadPartProperty(srcMgr, candidates)
    If Len(valText) > 0 Then WriteDrawingProperty dstMgr, targetName, valText
End Sub

Private Function ReadPartProperty(ByRef custPropMgr As SldWorks.CustomPropertyManager, ByVal propertyNames As Variant) As String
    On Error Resume Next

    Dim i As Long
    Dim valOut As String
    Dim resolvedOut As String
    Dim wasResolved As Boolean
    Dim linked As Boolean

    ReadPartProperty = vbNullString

    For i = LBound(propertyNames) To UBound(propertyNames)
        custPropMgr.Get6 CStr(propertyNames(i)), False, valOut, resolvedOut, wasResolved, linked

        If Len(resolvedOut) > 0 Then
            ReadPartProperty = resolvedOut
            Exit Function
        ElseIf Len(valOut) > 0 Then
            ReadPartProperty = valOut
            Exit Function
        End If
    Next i
End Function

Private Sub WriteDrawingProperty(ByRef custPropMgr As SldWorks.CustomPropertyManager, ByVal propName As String, ByVal propValue As String)
    On Error Resume Next

    custPropMgr.Add3 propName, swCustomInfoText, propValue, 0
    custPropMgr.Set2 propName, propValue
End Sub

Private Sub InsertGeneralNotes(ByRef swDraw As SldWorks.DrawingDoc)
    On Error Resume Next

    If DrawingContainsText(swDraw, "GENERAL NOTES") Then Exit Sub

    Dim swModel As SldWorks.ModelDoc2
    Set swModel = swDraw

    swDraw.ActivateView ""

    Dim noteText As String
    noteText = "GENERAL NOTES" & vbCrLf & _
               "1. All dimensions are in mm" & vbCrLf & _
               "2. All corners are chamfered 0.5 x 45 deg" & vbCrLf & _
               "3. Remove all sharp edges and burrs"

    Dim swNote As SldWorks.Note
    Set swNote = swModel.InsertNote(noteText)

    If Not swNote Is Nothing Then
        Dim swAnn As SldWorks.Annotation
        Set swAnn = swNote.GetAnnotation
        If Not swAnn Is Nothing Then swAnn.SetPosition2 0.22, 0.03, 0
    End If

    swModel.ClearSelection2 True
End Sub

Private Sub InsertBarcode(ByRef swPart As SldWorks.ModelDoc2, ByRef swDraw As SldWorks.DrawingDoc)
    On Error Resume Next

    Dim swModel As SldWorks.ModelDoc2
    Set swModel = swDraw

    swDraw.ActivateView ""

    Dim labelText As String
    labelText = FileNameWithoutExtension(swPart.GetPathName)
    If Len(labelText) = 0 Then Exit Sub

    If DrawingContainsText(swDraw, "*" & labelText & "*") Then Exit Sub

    Dim swNote As SldWorks.Note
    Set swNote = swModel.InsertNote("*" & labelText & "*")

    If Not swNote Is Nothing Then
        Dim swAnn As SldWorks.Annotation
        Set swAnn = swNote.GetAnnotation
        If Not swAnn Is Nothing Then swAnn.SetPosition2 0.02, 0.015, 0
    End If

    swModel.ClearSelection2 True
End Sub

Private Function DrawingContainsText(ByRef swDraw As SldWorks.DrawingDoc, ByVal findText As String) As Boolean
    On Error Resume Next

    Dim swView As SldWorks.View
    Dim swNote As SldWorks.Note
    Dim txt As String

    Set swView = swDraw.GetFirstView

    Do While Not swView Is Nothing
        Set swNote = swView.GetFirstNote
        Do While Not swNote Is Nothing
            txt = swNote.GetText
            If InStr(1, txt, findText, vbTextCompare) > 0 Then
                DrawingContainsText = True
                Exit Function
            End If
            Set swNote = swNote.GetNext
        Loop
        Set swView = swView.GetNextView
    Loop
End Function

Private Function FileNameWithoutExtension(ByVal path As String) As String
    Dim p As Long
    Dim d As Long
    Dim fileNameOnly As String

    p = InStrRev(path, "\")
    If p > 0 Then
        fileNameOnly = Mid$(path, p + 1)
    Else
        fileNameOnly = path
    End If

    d = InStrRev(fileNameOnly, ".")
    If d > 1 Then
        FileNameWithoutExtension = Left$(fileNameOnly, d - 1)
    Else
        FileNameWithoutExtension = fileNameOnly
    End If
End Function


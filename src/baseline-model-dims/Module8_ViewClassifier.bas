
Option Explicit

' View classification, API-backed.
'
' Replaces the name substring test that Module2_DrawingPipeline used to decide
' which views the ordinate engine may touch. That test looked for "ISO" in
' IView.Name, but SOLIDWORKS auto-names views "Drawing View1".."Drawing View5",
' so it never matched and the r7 live run put ordinate chains on the isometric
' view. Gap A6 in docs/BASELINE_TO_REFERENCE_DRAWING_GAP.md.
'
' Classification uses two documented members:
'   IView.Type              -> swDrawingViewTypes_e, isolates section/detail/
'                              projected views and the sheet itself.
'   IView.GetOrientationName -> the predefined orientation ("*Front", "*Top",
'                              "*Isometric", ...) or "" for section, detail,
'                              projected and unfolded views.
' These are the same strings Module2_DrawingPipeline passes to
' CreateDrawViewFromModelView3, so the round trip should be exact. What each
' actually returns for this pipeline has NOT been probed on this build; every
' classification is therefore printed into the QA report by DescribeView so the
' live run doubles as the probe.

Private Const swDrawingSheetType As Long = 1
Private Const swDrawingSectionViewType As Long = 2
Private Const swDrawingDetailViewType As Long = 3
Private Const swDrawingProjectedViewType As Long = 4

Public Const VIEW_ROLE_UNKNOWN As Long = 0
Public Const VIEW_ROLE_FRONT As Long = 1
Public Const VIEW_ROLE_TOP As Long = 2
Public Const VIEW_ROLE_BOTTOM As Long = 3
Public Const VIEW_ROLE_RIGHT As Long = 4
Public Const VIEW_ROLE_LEFT As Long = 5
Public Const VIEW_ROLE_BACK As Long = 6
Public Const VIEW_ROLE_PICTORIAL As Long = 7
Public Const VIEW_ROLE_SECTION As Long = 8
Public Const VIEW_ROLE_DETAIL As Long = 9
Public Const VIEW_ROLE_PROJECTED As Long = 10
Public Const VIEW_ROLE_SHEET As Long = 11

Public Function ClassifyView(ByRef swView As SldWorks.View) As Long
    ClassifyView = VIEW_ROLE_UNKNOWN
    If swView Is Nothing Then Exit Function

    Dim viewType As Long
    viewType = ReadViewType(swView)

    Select Case viewType
        Case swDrawingSheetType
            ClassifyView = VIEW_ROLE_SHEET
            Exit Function
        Case swDrawingSectionViewType
            ClassifyView = VIEW_ROLE_SECTION
            Exit Function
        Case swDrawingDetailViewType
            ClassifyView = VIEW_ROLE_DETAIL
            Exit Function
        Case swDrawingProjectedViewType
            ClassifyView = VIEW_ROLE_PROJECTED
            Exit Function
    End Select

    ClassifyView = RoleFromOrientationName(ReadOrientationName(swView))
    If ClassifyView <> VIEW_ROLE_UNKNOWN Then Exit Function

    ' Last resort only. The orientation name is the contract; this keeps a
    ' build whose GetOrientationName returns "" from silently treating a
    ' pictorial view as a candidate for dimensions.
    ClassifyView = RoleFromViewName(swView.Name)
End Function

' True when the ordinate engine may run on this view.
'
' The reference drawing P-0251-14A-001 carries ordinate chains on the front
' view only (gap A5), so that is the policy. Two deliberate asymmetries:
'   - VIEW_ROLE_PICTORIAL is refused unconditionally. An isometric view has no
'     true-length axes; ordinates on it are always wrong.
'   - VIEW_ROLE_UNKNOWN is allowed. If classification breaks, producing extra
'     ordinates is recoverable and visible; producing none looks like a working
'     run that quietly did nothing. The QA roster names every UNKNOWN view.
Public Function AllowsOrdinateDimensions(ByVal role As Long) As Boolean
    Select Case role
        Case VIEW_ROLE_PICTORIAL, VIEW_ROLE_SHEET
            AllowsOrdinateDimensions = False
        Case VIEW_ROLE_FRONT, VIEW_ROLE_UNKNOWN
            AllowsOrdinateDimensions = True
        Case Else
            AllowsOrdinateDimensions = False
    End Select
End Function

' One line of run evidence per view: what the API returned and what it was
' classified as. This is the probe for IView.Type / IView.GetOrientationName.
Public Function DescribeView(ByRef swView As SldWorks.View) As String
    If swView Is Nothing Then
        DescribeView = "(Nothing)"
        Exit Function
    End If

    Dim orientation As String
    orientation = ReadOrientationName(swView)
    If Len(orientation) = 0 Then orientation = "(empty)"

    DescribeView = swView.Name & _
        " | Type=" & CStr(ReadViewType(swView)) & _
        " | Orientation=" & orientation & _
        " | Role=" & RoleName(ClassifyView(swView))
End Function

Public Function RoleName(ByVal role As Long) As String
    Select Case role
        Case VIEW_ROLE_FRONT: RoleName = "Front"
        Case VIEW_ROLE_TOP: RoleName = "Top"
        Case VIEW_ROLE_BOTTOM: RoleName = "Bottom"
        Case VIEW_ROLE_RIGHT: RoleName = "Right"
        Case VIEW_ROLE_LEFT: RoleName = "Left"
        Case VIEW_ROLE_BACK: RoleName = "Back"
        Case VIEW_ROLE_PICTORIAL: RoleName = "Pictorial"
        Case VIEW_ROLE_SECTION: RoleName = "Section"
        Case VIEW_ROLE_DETAIL: RoleName = "Detail"
        Case VIEW_ROLE_PROJECTED: RoleName = "Projected"
        Case VIEW_ROLE_SHEET: RoleName = "Sheet"
        Case Else: RoleName = "Unknown"
    End Select
End Function

Private Function ReadViewType(ByRef swView As SldWorks.View) As Long
    On Error Resume Next
    ReadViewType = -1
    ReadViewType = swView.Type
    On Error GoTo 0
End Function

Private Function ReadOrientationName(ByRef swView As SldWorks.View) As String
    On Error Resume Next
    ReadOrientationName = vbNullString
    ReadOrientationName = swView.GetOrientationName
    On Error GoTo 0
End Function

' GetOrientationName returns the predefined orientation with a leading "*",
' e.g. "*Front". Match on the stripped, upper-cased text so a build that omits
' the marker still classifies. Anything unrecognised stays UNKNOWN so the
' caller can fall back rather than silently mis-role a view.
Private Function RoleFromOrientationName(ByVal orientation As String) As Long
    Dim key As String
    key = UCase$(Trim$(orientation))
    Do While Left$(key, 1) = "*"
        key = Mid$(key, 2)
    Loop
    key = Trim$(key)

    Select Case key
        Case "FRONT": RoleFromOrientationName = VIEW_ROLE_FRONT
        Case "TOP": RoleFromOrientationName = VIEW_ROLE_TOP
        Case "BOTTOM": RoleFromOrientationName = VIEW_ROLE_BOTTOM
        Case "RIGHT": RoleFromOrientationName = VIEW_ROLE_RIGHT
        Case "LEFT": RoleFromOrientationName = VIEW_ROLE_LEFT
        Case "BACK", "REAR": RoleFromOrientationName = VIEW_ROLE_BACK
        Case "ISOMETRIC", "DIMETRIC", "TRIMETRIC"
            RoleFromOrientationName = VIEW_ROLE_PICTORIAL
        Case Else: RoleFromOrientationName = VIEW_ROLE_UNKNOWN
    End Select
End Function

Private Function RoleFromViewName(ByVal viewName As String) As Long
    Dim upper As String
    upper = UCase$(Trim$(viewName))

    If InStr(1, upper, "ISOMETRIC", vbBinaryCompare) > 0 Then
        RoleFromViewName = VIEW_ROLE_PICTORIAL
    ElseIf InStr(1, upper, "DIMETRIC", vbBinaryCompare) > 0 Then
        RoleFromViewName = VIEW_ROLE_PICTORIAL
    ElseIf InStr(1, upper, "TRIMETRIC", vbBinaryCompare) > 0 Then
        RoleFromViewName = VIEW_ROLE_PICTORIAL
    Else
        RoleFromViewName = VIEW_ROLE_UNKNOWN
    End If
End Function


' Compile-failure localisation no-op called by
' Module20_ProbeRunner.R23_TouchAllModules.
Public Sub R23_CompileTouch()
End Sub

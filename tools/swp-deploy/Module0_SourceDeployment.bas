Option Explicit

' This bootstrap is imported once into Fable.swp. It is deliberately excluded
' from the managed component manifest so that it never removes itself while it
' is running. The project must reference Microsoft Visual Basic for
' Applications Extensibility 5.3 so that the intrinsic VBE object is available.

Private Const REQUEST_PATH As String = _
    "C:\Users\V.T\Documents\VBA 3D TO 2D\tools\swp-deploy\deployment-request.txt"

Private Const FOR_READING As Long = 1
Private Const FOR_WRITING As Long = 2
Private Const TRISTATE_FALSE As Long = 0
Private Const VBEXT_CT_STDMODULE As Long = 1
Private Const VBEXT_CT_CLASSMODULE As Long = 2

Public Sub DeployFromRequest()
    On Error GoTo Failed

    Dim currentStage As String
    currentStage = "READ_REQUEST"

    Dim requestText As String
    requestText = ReadAllText(REQUEST_PATH)

    Dim expectedProject As String
    Dim expectedInput As String
    Dim outputSwp As String
    Dim resultFile As String

    expectedProject = ReadRequestValue(requestText, "EXPECTED_PROJECT")
    expectedInput = ReadRequestValue(requestText, "INPUT_SWP")
    outputSwp = ReadRequestValue(requestText, "OUTPUT_SWP")
    resultFile = ReadRequestValue(requestText, "RESULT_FILE")

    If Len(expectedProject) = 0 Or Len(expectedInput) = 0 Or _
       Len(outputSwp) = 0 Or Len(resultFile) = 0 Then

        Err.Raise vbObjectError + 7100, "DeployFromRequest", _
            "The deployment request is missing a required path or project name."
    End If

    Dim swApp As SldWorks.SldWorks
    Set swApp = Application.SldWorks

    If swApp Is Nothing Then
        Err.Raise vbObjectError + 7101, "DeployFromRequest", _
            "The SOLIDWORKS application interface is unavailable."
    End If

    Dim currentMacroPath As String
    currentMacroPath = swApp.GetCurrentMacroPathName

    If Not PathsMatch(currentMacroPath, expectedInput) Then
        Err.Raise vbObjectError + 7102, "DeployFromRequest", _
            "Refusing deployment because the running macro is not the candidate input. " & _
            "Actual='" & currentMacroPath & "'; expected='" & expectedInput & "'."
    End If

    Dim targetProject As Object
    Set targetProject = VBE.ActiveVBProject

    If targetProject Is Nothing Then
        Err.Raise vbObjectError + 7103, "DeployFromRequest", _
            "VBE.ActiveVBProject is unavailable."
    End If

    If StrComp(CStr(targetProject.Name), expectedProject, vbTextCompare) <> 0 Then
        Err.Raise vbObjectError + 7104, "DeployFromRequest", _
            "Active VBA project mismatch. Actual='" & CStr(targetProject.Name) & _
            "'; expected='" & expectedProject & "'."
    End If

    If FileExists(outputSwp) Then
        Err.Raise vbObjectError + 7105, "DeployFromRequest", _
            "The requested output already exists: " & outputSwp
    End If

    Dim componentLines As Collection
    Set componentLines = ReadComponentLines(requestText)

    If componentLines.Count = 0 Then
        Err.Raise vbObjectError + 7106, "DeployFromRequest", _
            "The deployment request contains no managed components."
    End If

    currentStage = "VALIDATE_SOURCES"
    ValidateAllSources componentLines

    Dim logText As String
    logText = "DEPLOYMENT|status=STARTED" & vbCrLf & _
        "INPUT_SWP=" & expectedInput & vbCrLf & _
        "OUTPUT_SWP=" & outputSwp & vbCrLf

    Dim componentLine As Variant
    For Each componentLine In componentLines
        currentStage = "REPLACE_COMPONENT:" & _
            CStr(Split(CStr(componentLine), "|")(0))
        ReplaceOneComponent targetProject, CStr(componentLine), logText
    Next componentLine

    currentStage = "SAVE_CANDIDATE_INPUT"
    SaveActiveMacroProject targetProject

    logText = logText & "DEPLOYMENT|status=SUCCESS" & vbCrLf & _
        "SAVED_PATH=" & expectedInput & vbCrLf & _
        "REQUESTED_OUTPUT=" & outputSwp & vbCrLf
    WriteAllText resultFile, logText
    Exit Sub

Failed:
    Dim failureNumber As Long
    Dim failureDescription As String
    failureNumber = Err.Number
    failureDescription = Err.Description

    Dim failureResult As String
    On Error Resume Next
    failureResult = resultFile
    If Len(failureResult) = 0 Then
        failureResult = ReadRequestValue(ReadAllText(REQUEST_PATH), "RESULT_FILE")
    End If
    If Len(failureResult) > 0 Then
        WriteAllText failureResult, _
            "DEPLOYMENT|status=FAILED|error=" & CStr(failureNumber) & _
            "|stage=" & CleanLogText(currentStage) & _
            "|detail=" & CleanLogText(failureDescription) & vbCrLf
    End If
    On Error GoTo 0
End Sub

Public Sub CompileProbe()
    On Error GoTo Failed

    Dim requestText As String
    requestText = ReadAllText(REQUEST_PATH)

    Dim resultFile As String
    resultFile = ReadRequestValue(requestText, "COMPILE_RESULT_FILE")

    If Len(resultFile) = 0 Then
        Err.Raise vbObjectError + 7120, "CompileProbe", _
            "COMPILE_RESULT_FILE is missing from the deployment request."
    End If

    WriteAllText resultFile, "COMPILE_PROBE|status=SUCCESS" & vbCrLf
    Exit Sub

Failed:
    Dim failureNumber As Long
    Dim failureDescription As String
    failureNumber = Err.Number
    failureDescription = Err.Description

    On Error Resume Next
    If Len(resultFile) > 0 Then
        WriteAllText resultFile, _
            "COMPILE_PROBE|status=FAILED|error=" & CStr(failureNumber) & _
            "|detail=" & CleanLogText(failureDescription) & vbCrLf
    End If
    On Error GoTo 0
End Sub

Private Sub ValidateAllSources(ByRef componentLines As Collection)
    Dim componentLine As Variant
    For Each componentLine In componentLines
        Dim fields As Variant
        fields = Split(CStr(componentLine), "|")

        If UBound(fields) <> 2 Then
            Err.Raise vbObjectError + 7130, "ValidateAllSources", _
                "Invalid COMPONENT record: " & CStr(componentLine)
        End If

        Dim componentName As String
        Dim componentKind As String
        Dim sourcePath As String
        componentName = Trim$(CStr(fields(0)))
        componentKind = Trim$(CStr(fields(1)))
        sourcePath = Trim$(CStr(fields(2)))

        If componentKind <> "StdModule" And componentKind <> "ClassModule" Then
            Err.Raise vbObjectError + 7131, "ValidateAllSources", _
                "Unsupported component kind for '" & componentName & "': " & _
                componentKind
        End If

        If Not FileExists(sourcePath) Then
            Err.Raise vbObjectError + 7132, "ValidateAllSources", _
                "Source file is missing for '" & componentName & "': " & sourcePath
        End If

        Dim sourceText As String
        sourceText = ReadAllText(sourcePath)

        If ContainsAttributeMetadata(sourceText) Then
            Err.Raise vbObjectError + 7133, "ValidateAllSources", _
                "VBA Attribute metadata is not allowed in deployable source '" & _
                componentName & "'."
        End If

        If InStr(1, sourceText, "Option Explicit", vbTextCompare) = 0 Then
            Err.Raise vbObjectError + 7134, "ValidateAllSources", _
                "Option Explicit is missing from component '" & componentName & "'."
        End If
    Next componentLine
End Sub

Private Sub SaveActiveMacroProject(ByRef targetProject As Object)
    ' VBProject.SaveAs raises error 748 for a SOLIDWORKS host-managed .swp.
    ' Execute the VBE's built-in Save command (control ID 3) instead. The
    ' candidate already has its final input path, so no Save As dialog is due.
    Dim saveControl As Object
    Set saveControl = targetProject.VBE.CommandBars.FindControl(1, 3, "", False)

    If saveControl Is Nothing Then
        Err.Raise vbObjectError + 7107, "SaveActiveMacroProject", _
            "The VBE Save command (ID 3) was not found."
    End If

    If Not CBool(saveControl.Enabled) Then
        Err.Raise vbObjectError + 7108, "SaveActiveMacroProject", _
            "The VBE Save command is disabled while deploying the candidate."
    End If

    saveControl.Execute
    DoEvents
End Sub

Private Function ContainsAttributeMetadata(ByVal sourceText As String) As Boolean
    Dim normalizedText As String
    normalizedText = Replace(sourceText, vbCrLf, vbLf)

    Dim sourceLines As Variant
    sourceLines = Split(normalizedText, vbLf)

    Dim sourceLine As Variant
    For Each sourceLine In sourceLines
        If StrComp(Left$(Trim$(CStr(sourceLine)), 10), _
            "Attribute ", vbTextCompare) = 0 Then

            ContainsAttributeMetadata = True
            Exit Function
        End If
    Next sourceLine
End Function

Private Sub ReplaceOneComponent( _
    ByRef targetProject As Object, _
    ByVal componentLine As String, _
    ByRef logText As String)

    Dim fields As Variant
    fields = Split(componentLine, "|")

    Dim componentName As String
    Dim componentKind As String
    Dim sourcePath As String
    componentName = Trim$(CStr(fields(0)))
    componentKind = Trim$(CStr(fields(1)))
    sourcePath = Trim$(CStr(fields(2)))

    Dim existingComponent As Object
    Set existingComponent = TryGetComponent(targetProject, componentName)

    If Not existingComponent Is Nothing Then
        targetProject.VBComponents.Remove existingComponent
    End If

    Dim componentType As Long
    If componentKind = "StdModule" Then
        componentType = VBEXT_CT_STDMODULE
    ElseIf componentKind = "ClassModule" Then
        componentType = VBEXT_CT_CLASSMODULE
    Else
        Err.Raise vbObjectError + 7140, "ReplaceOneComponent", _
            "Unsupported component kind for '" & componentName & "': " & _
            componentKind
    End If

    Dim importedComponent As Object
    Set importedComponent = targetProject.VBComponents.Add(componentType)

    If importedComponent Is Nothing Then
        Err.Raise vbObjectError + 7141, "ReplaceOneComponent", _
            "VBComponents.Add returned Nothing for '" & componentName & "'."
    End If

    importedComponent.Name = componentName

    If StrComp(CStr(importedComponent.Name), componentName, vbTextCompare) <> 0 Then
        Err.Raise vbObjectError + 7142, "ReplaceOneComponent", _
            "Created component could not be renamed. Actual='" & _
            CStr(importedComponent.Name) & "'; expected='" & componentName & "'."
    End If

    Dim sourceText As String
    sourceText = ReadAllText(sourcePath)
    importedComponent.CodeModule.AddFromString sourceText

    If importedComponent.CodeModule.CountOfLines = 0 Then
        Err.Raise vbObjectError + 7143, "ReplaceOneComponent", _
            "CodeModule.AddFromString added no code for '" & componentName & "'."
    End If

    logText = logText & "COMPONENT|name=" & componentName & _
        "|kind=" & componentKind & "|status=CREATED" & vbCrLf
End Sub

Private Function TryGetComponent( _
    ByRef targetProject As Object, _
    ByVal componentName As String) As Object

    On Error Resume Next
    Set TryGetComponent = targetProject.VBComponents.Item(componentName)
    On Error GoTo 0
End Function

Private Function ReadComponentLines(ByVal requestText As String) As Collection
    Dim result As New Collection
    Dim lines As Variant
    lines = Split(Replace(requestText, vbCrLf, vbLf), vbLf)

    Dim lineValue As Variant
    For Each lineValue In lines
        If StrComp(Left$(CStr(lineValue), 10), "COMPONENT=", vbTextCompare) = 0 Then
            result.Add Mid$(CStr(lineValue), 11)
        End If
    Next lineValue

    Set ReadComponentLines = result
End Function

Private Function ReadRequestValue( _
    ByVal requestText As String, _
    ByVal keyName As String) As String

    Dim prefix As String
    prefix = keyName & "="

    Dim lines As Variant
    lines = Split(Replace(requestText, vbCrLf, vbLf), vbLf)

    Dim lineValue As Variant
    For Each lineValue In lines
        If StrComp(Left$(CStr(lineValue), Len(prefix)), _
            prefix, vbTextCompare) = 0 Then

            ReadRequestValue = Mid$(CStr(lineValue), Len(prefix) + 1)
            Exit Function
        End If
    Next lineValue
End Function

Private Function PathsMatch(ByVal firstPath As String, ByVal secondPath As String) As Boolean
    On Error GoTo NotMatched

    Dim fileSystem As Object
    Set fileSystem = CreateObject("Scripting.FileSystemObject")

    PathsMatch = (StrComp( _
        fileSystem.GetAbsolutePathName(firstPath), _
        fileSystem.GetAbsolutePathName(secondPath), _
        vbTextCompare) = 0)
    Exit Function

NotMatched:
    PathsMatch = False
End Function

Private Function FileExists(ByVal filePath As String) As Boolean
    On Error Resume Next
    FileExists = CreateObject("Scripting.FileSystemObject").FileExists(filePath)
    On Error GoTo 0
End Function

Private Function ReadAllText(ByVal filePath As String) As String
    Dim fileSystem As Object
    Set fileSystem = CreateObject("Scripting.FileSystemObject")

    Dim textStream As Object
    Set textStream = fileSystem.OpenTextFile( _
        filePath, FOR_READING, False, TRISTATE_FALSE)
    ReadAllText = textStream.ReadAll
    textStream.Close
End Function

Private Sub WriteAllText(ByVal filePath As String, ByVal textValue As String)
    Dim fileSystem As Object
    Set fileSystem = CreateObject("Scripting.FileSystemObject")

    Dim textStream As Object
    Set textStream = fileSystem.OpenTextFile( _
        filePath, FOR_WRITING, True, TRISTATE_FALSE)
    textStream.Write textValue
    textStream.Close
End Sub

Private Function CleanLogText(ByVal value As String) As String
    CleanLogText = Replace(Replace(value, vbCr, " "), vbLf, " ")
End Function

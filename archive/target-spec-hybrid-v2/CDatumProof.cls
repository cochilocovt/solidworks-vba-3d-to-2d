Option Explicit

Public DrawingEntity As SldWorks.Entity
Public ModelEntity As SldWorks.Entity
Public Component As SldWorks.Component2
Public DrawingView As SldWorks.View
Public RequestedOrigin As String
Public ProofSource As String
Public StableKey As String
Public EntityKind As String
Public ViewName As String
Public ComponentName As String
Public ConfigurationName As String
Public ModelX As Double
Public ModelY As Double
Public ModelZ As Double
Public ViewX As Double
Public ViewY As Double
Public SheetX As Double
Public SheetY As Double
Public Proven As Boolean
Public SelectionSucceeded As Boolean
Public SelectionAttempts As Long
Public FailureReason As String

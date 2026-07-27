Option Explicit

Public DrawingEdge As SldWorks.Edge
Public ModelEdge As SldWorks.Edge
Public DrawingEntityAliases As Collection
Public ModelEntityAliases As Collection
Public Component As SldWorks.Component2
Public OwnerFeature As SldWorks.Feature
Public DrawingView As SldWorks.View

Public ViewName As String
Public ComponentName As String
Public FeatureName As String
Public FeatureType As String
Public ConfigurationName As String
Public RadiiStack As String
Public SemanticsSummary As String
Public MachiningSide As String
Public SemanticKey As String
Public FamilyKey As String
Public PhysicalInstanceKey As String
Public CoverageScopeKey As String
Public OwnershipProof As String

Public ModelX As Double
Public ModelY As Double
Public ModelZ As Double
Public ViewX As Double
Public ViewY As Double
Public SheetX As Double
Public SheetY As Double
Public AxisX As Double
Public AxisY As Double
Public AxisZ As Double
Public ModelAxisX As Double
Public ModelAxisY As Double
Public ModelAxisZ As Double
Public Radius As Double

Public Accepted As Boolean
Public RejectionReason As String
Public CoveredX As Boolean
Public CoveredY As Boolean
Public CoverageInspectionSucceeded As Boolean
Public CoverageSourceX As String
Public CoverageSourceY As String

Private Sub Class_Initialize()
    Set DrawingEntityAliases = New Collection
    Set ModelEntityAliases = New Collection
End Sub

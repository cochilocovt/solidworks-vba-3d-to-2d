\# Architecture



\## Execution Flow



```text

Module1\_Main

&#x20; -> UserForm1

&#x20; -> Module2\_DrawingPipeline

&#x20;      -> Create drawing and views

&#x20;      -> Module4\_ModelItemImporter

&#x20;      -> Module5\_FallbackDimensionEngine

&#x20;      -> Module4 dimension arrangement

&#x20;      -> Module7\_TitleBlockEngine

&#x20;      -> Module6\_QAEngine

&#x20;           -> Module3\_ModelAudit

&#x20;           -> Module4 displayed-dimension count




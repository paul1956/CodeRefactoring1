; Unshipped analyzer release
; https://github.com/dotnet/roslyn-analyzers/blob/master/src/Microsoft.CodeAnalysis.Analyzers/ReleaseTrackingAnalyzers.Help.md

### New Rules
Rule ID | Category | Severity | Notes
--------|----------|----------|-------
CC00063 | Usage | Error | UriAnalyzer
VBF0002 | Naming | Warning | ArgumentExceptionAnalyzer
VBF0003 | Design | Warning | CatchEmptyAnalyzer
VBF0004 | Design | Warning | EmptyCatchBlockAnalyzer
VBF0011 | Performance | Warning | RemoveWhereWhenItIsPossibleAnalyzer
VBF0023 | Performance | Warning | SealedAttributeAnalyzer
VBF0024 | Design | Warning | StaticConstructorExceptionAnalyzer
VBF0030 | Performance | Info | MakeLocalVariableConstWhenPossibleAnalyzer
VBF0039 | Performance | Warning | StringBuilderInLoopAnalyzer
VBF0043 | Refactoring | Hidden | ChangeAnyToAllAnalyzer
VBF0044 | Refactoring | Hidden | ParameterRefactoryAnalyzer
VBF0057 | Usage | Warning | UnusedParametersAnalyzer
VBF0060 | Usage | Warning | MustInheritClassShouldNotHavePublicConstructorsAnalyzer
VBF0068 | Usage | Info | RemovePrivateMethodNeverUsedAnalyzer
VBF0069 | Usage | Info | NeverUsedAnalyzer
VBF0070 | Reliability | Hidden | UseConfigureAwaitFalseAnalyzer
VBF0092 | Refactoring | Hidden | ChangeAnyToAllAnalyzer
VBF0101 | Style | Error | AddAsClauseAnalyzer
VBF0103 | Style | Warning | AddAsClauseForLambdasAnalyzer
VBF0104 | Style | Error | AddAsClauseAsObjectAnalyzer

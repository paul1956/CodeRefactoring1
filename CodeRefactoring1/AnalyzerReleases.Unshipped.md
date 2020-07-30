; Unshipped analyzer release
; https://github.com/dotnet/roslyn-analyzers/blob/master/src/Microsoft.CodeAnalysis.Analyzers/ReleaseTrackingAnalyzers.Help.md

### New Rules
Rule ID | Category | Severity | Notes
--------|----------|----------|-------
ByValAnalyzer_Fix | Style | Info | ByValAnalyzer_FixAnalyzer
CC00063 | Usage | Error | UriAnalyzer
CC0064 | Usage | Error | IPAddressAnalyzer
VBF0002 | Naming | Warning | ArgumentExceptionAnalyzer
VBF0003 | Design | Warning | CatchEmptyAnalyzer
VBF0004 | Design | Warning | EmptyCatchBlockAnalyzer
VBF0011 | Performance | Warning | RemoveWhereWhenItIsPossibleAnalyzer
VBF0021 | Design | Warning | NameOfAnalyzer
VBF0023 | Performance | Warning | SealedAttributeAnalyzer
VBF0024 | Design | Warning | StaticConstructorExceptionAnalyzer
VBF0030 | Performance | Info | MakeLocalVariableConstWhenPossibleAnalyzer
VBF0035 | Refactoring | Hidden | AllowMembersOrderingAnalyzer
VBF0039 | Performance | Warning | StringBuilderInLoopAnalyzer
VBF0043 | Refactoring | Hidden | ChangeAnyToAllAnalyzer
VBF0044 | Refactoring | Hidden | ParameterRefactoryAnalyzer
VBF0054 | Usage | Error | JsonNetAnalyzer
VBF0057 | Usage | Warning | UnusedParametersAnalyzer
VBF0060 | Usage | Warning | MustInheritClassShouldNotHavePublicConstructorsAnalyzer
VBF0068 | Usage | Info | RemovePrivateMethodNeverUsedAnalyzer
VBF0069 | Usage | Info | NeverUsedAnalyzer
VBF0070 | Reliability | Hidden | UseConfigureAwaitFalseAnalyzer
VBF0092 | Refactoring | Hidden | ChangeAnyToAllAnalyzer
VBF0103 | Style | Error | AddAsClauseForLambdasAnalyzer
VBF0104 | Style | Error | AddAsClauseAsObjectAnalyzer
VBF0108 | Design | Warning | NameOfAnalyzer
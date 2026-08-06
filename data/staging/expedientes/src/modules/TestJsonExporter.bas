Attribute VB_Name = "TestJsonExporter"
Option Compare Database
Option Explicit

Public Function TestJsonExporter_RetiredLegacyModuleHasNoExecutableTests() As String
    Dim logs(0 To 0) As String

    logs(0) = "Retired legacy ad-hoc JSON exporter tests; covered by Test_ExpedienteJsonExporter and harness contracts"
    TestJsonExporter_RetiredLegacyModuleHasNoExecutableTests = BuildJsonOk("retired", logs)
End Function


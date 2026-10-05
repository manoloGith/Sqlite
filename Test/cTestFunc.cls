VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTestFunc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=========================================================================
' cTestFunc - scalar UDFs used by frmTest
'   VB_UPPER(text)   -> UCase$ (checks UTF-8 round-trip through a callback)
'   VB_ADD(a, b)     -> a + b as double
'   VB_FAIL()        -> raises a VB error (must surface as an SQL error)
'=========================================================================
Option Explicit

Implements IFunction

Private Property Get IFunction_DefinedNames() As String
    IFunction_DefinedNames = "VB_UPPER,VB_ADD,VB_FAIL"
End Property

Private Sub IFunction_Callback(ByVal ZeroBasedNameIndex As Long, ByVal ParamCount As Long, UDF As cUDFMethods)
    Select Case ZeroBasedNameIndex
    Case 0
        UDF.SetResultText UCase$(UDF.GetText(1))
    Case 1
        UDF.SetResultDouble UDF.GetDouble(1) + UDF.GetDouble(2)
    Case 2
        Err.Raise vbObjectError + 1, "cTestFunc", "fallo intencionado"
    End Select
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTestColl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=========================================================================
' cTestColl - collation used by frmTest
'   VB_REVERSE       -> case-insensitive, descending order
'=========================================================================
Option Explicit

Implements ICollation

Private Property Get ICollation_DefinedNames() As String
    ICollation_DefinedNames = "VB_REVERSE"
End Property

Private Function ICollation_CallbackCollate(ByVal ZeroBasedNameIndex As Long, S1 As String, S2 As String) As Long
    ICollation_CallbackCollate = -StrComp(S1, S2, vbTextCompare)
End Function

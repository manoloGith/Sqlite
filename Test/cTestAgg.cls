VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTestAgg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'=========================================================================
' cTestAgg - aggregate UDF used by frmTest
'   VB_SUM(x)        -> sum of x as double
'=========================================================================
Option Explicit

Implements IAggregateFunction

Private m_dblSum                    As Double

Private Property Get IAggregateFunction_DefinedNames() As String
    IAggregateFunction_DefinedNames = "VB_SUM"
End Property

Private Sub IAggregateFunction_CallbackStep(ByVal ZeroBasedNameIndex As Long, ByVal ParamCount As Long, UDF As cUDFMethods)
    m_dblSum = m_dblSum + UDF.GetDouble(1)
End Sub

Private Sub IAggregateFunction_CallbackFinal(ByVal ZeroBasedNameIndex As Long, UDF As cUDFMethods)
    UDF.SetResultDouble m_dblSum
    m_dblSum = 0
End Sub

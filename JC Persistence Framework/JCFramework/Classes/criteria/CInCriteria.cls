VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CInCriteria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements CSelectionCriteria
Private m_selCriteria As CSelectionCriteria

Private m_Attributes As Collection

Public Property Get Attributes() As Collection
  Set Attributes = m_Attributes
End Property

Public Property Set Attributes(ByVal attrs As Collection)
  Set m_Attributes = attrs
End Property

Public Property Get AttrMap() As CAttributeMap
  Set AttrMap = CSelectionCriteria_AttrMap
End Property

Public Property Set AttrMap(ByVal anAttrMap As CAttributeMap)
  Set CSelectionCriteria_AttrMap = anAttrMap
End Property


Private Sub Class_Initialize()
  Set m_selCriteria = New CSelectionCriteria
End Sub

Private Property Set CSelectionCriteria_AttrMap(ByVal RHS As CAttributeMap)
  Set m_selCriteria.AttrMap = RHS
End Property

Private Property Get CSelectionCriteria_AttrMap() As CAttributeMap
  Set CSelectionCriteria_AttrMap = m_selCriteria.AttrMap
End Property

Private Sub CSelectionCriteria_fillSqlStatement(ByVal strValue As String, ByVal pStatement As CSqlStatement, ByVal clm As CClassMap)
  'TODO: Implement this functionality in CCriteriaCondition in method: "fillStatement"
End Sub



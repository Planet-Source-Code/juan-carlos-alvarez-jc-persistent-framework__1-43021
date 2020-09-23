VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSelectionCriteria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'**************************************************
'Class: CSelectionCriteria
'Author: Juan Carlos Alvarez
'**************************************************

Private m_attrMap As CAttributeMap

Public Property Get AttrMap() As CAttributeMap
  Set AttrMap = m_attrMap
End Property

Public Property Set AttrMap(ByVal anAttrMap As CAttributeMap)
  Set m_attrMap = anAttrMap
End Property

Public Sub fillSqlStatement(ByVal strValue As String, ByVal pStatement As CSqlStatement, ByVal clm As CClassMap)
  'TODO: Implement this functionality
End Sub

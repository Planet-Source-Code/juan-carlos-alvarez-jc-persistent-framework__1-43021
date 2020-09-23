VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "COrderEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'**************************************************
'Class: COrderEntry
'Author: Juan Carlos Alvarez
'**************************************************

'Entry for building ORDER BY clause.
Private m_attributeMap As CAttributeMap
Private m_ascend As Boolean

'Returns true if order is ascending or
'false if descending.
Public Property Get Ascend() As Boolean
    Ascend = m_ascend
End Property

Public Property Let Ascend(a As Boolean)
    m_ascend = a
End Property

Public Property Get AttributeMap() As CAttributeMap
    Set AttributeMap = m_attributeMap
End Property

Public Property Set AttributeMap(AttrMap As CAttributeMap)
    Set m_attributeMap = AttrMap
End Property

'COrderEntry constructor.
Private Sub Class_Initialize()
  m_ascend = True
End Sub

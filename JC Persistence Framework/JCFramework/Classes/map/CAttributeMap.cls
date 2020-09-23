VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAttributeMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'**************************************************
'Class: CAttributeMap
'Author: Juan Carlos Alvarez
'**************************************************

Private m_Name As String
Private m_columnMap As CColumnMap
Private m_attributeMap As CAttributeMap

Public Property Get Name() As String
    Name = m_Name
End Property

Public Property Let Name(n As String)
    m_Name = n
End Property

Public Property Get ColumnMap() As CColumnMap
    Set ColumnMap = m_columnMap
End Property

Public Property Set ColumnMap(clmMap As CColumnMap)
    Set m_columnMap = clmMap
End Property

Public Property Get AttributeMap() As CAttributeMap
    Set AttributeMap = m_attributeMap
End Property

Public Property Set AttributeMap(attrMap As CAttributeMap)
    Set m_attributeMap = attrMap
End Property

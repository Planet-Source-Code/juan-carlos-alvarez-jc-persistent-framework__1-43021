VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CUDAMapEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'**************************************************
'Class: CUDAMapEntry
'Author: Juan Carlos Alvarez
'**************************************************

'This class represents entry of UniDirectionalAssociationMap.

Private m_fromAttrMap As CAttributeMap
Private m_toAttrMap As CAttributeMap

'Class Properties
Public Property Get FromAttrMap() As CAttributeMap
  Set FromAttrMap = m_fromAttrMap
End Property

Public Property Set FromAttrMap(ByVal f As CAttributeMap)
  Set m_fromAttrMap = f
End Property

Public Property Get ToAttrMap() As CAttributeMap
  Set ToAttrMap = m_toAttrMap
End Property

Public Property Set ToAttrMap(ByVal t As CAttributeMap)
  Set m_toAttrMap = t
End Property

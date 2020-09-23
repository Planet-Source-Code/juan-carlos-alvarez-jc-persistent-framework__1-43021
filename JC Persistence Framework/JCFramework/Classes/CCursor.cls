VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCursor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'**************************************************
'Class: CCursor
'Author: Juan Carlos Alvarez
'**************************************************

Private m_rs As CResultset
Private m_classMap As CClassMap

Public Property Get ResultSet() As CResultset
  Set ResultSet = m_rs
End Property

Public Property Set ResultSet(ByVal rs As CResultset)
  Set m_rs = rs
End Property

Public Property Get ClassMap() As CClassMap
  Set ClassMap = m_classMap
End Property

Public Property Set ClassMap(ByVal clm As CClassMap)
  Set m_classMap = clm
End Property

Public Sub loadObject(ByRef obj As CPersistentObject)
'TO DO: Handle errors
  m_classMap.retrieveObject obj, m_rs
End Sub

Public Sub nextCursor()
  m_rs.ResultSet.MoveNext
End Sub

Public Function hasElements() As Boolean
  If m_rs.ResultSet.EOF Then
    hasElements = False
  Else
    hasElements = True
  End If
End Function


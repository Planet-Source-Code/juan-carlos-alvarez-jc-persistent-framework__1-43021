VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CColumnMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'**************************************************
'Class: CColumnMap
'Author: Juan Carlos Alvarez
'**************************************************

Public Enum ColumnKeyTypes
  KEY_NONE = 0
  PRIMARY_KEY = 1
  FOREIGN_KEY = 2
End Enum
Private m_name As String
Private m_keyType As Integer
Private m_tableMap As CTableMap

Public Property Get Name() As String
    Name = m_name
End Property

Public Property Let Name(n As String)
    m_name = n
End Property

Public Property Get KeyType() As Integer
    KeyType = m_keyType
End Property

Public Property Let KeyType(k As Integer)
    m_keyType = k
End Property

Public Property Get TableMap() As CTableMap
    Set TableMap = m_tableMap
End Property

Public Property Set TableMap(tblMap As CTableMap)
    Set m_tableMap = tblMap
End Property

Private Sub Class_Initialize()
    m_keyType = KEY_NONE
End Sub

Public Function getFullyQualifiedName() As String
    getFullyQualifiedName = Me.TableMap.Name + "." + Me.Name
End Function

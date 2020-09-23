VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CTableMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'**************************************************
'Class: CTableMap
'Author: Juan Carlos Alvarez
'**************************************************

Private m_name As String
Private m_databaseMap As CDatabaseMap

Public Property Get Name() As String
  Name = m_name
End Property

Public Property Let Name(ByVal n As String)
  m_name = n
End Property

Public Property Get DatabaseMap() As CDatabaseMap
  Set DatabaseMap = m_databaseMap
End Property

Public Property Let DatabaseMap(ByVal dbMap As CDatabaseMap)
  Set m_databaseMap = dbMap
End Property

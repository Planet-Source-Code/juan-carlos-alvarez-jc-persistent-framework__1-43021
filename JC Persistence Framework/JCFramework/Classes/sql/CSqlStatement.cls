VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CSqlStatement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'**************************************************
'Class: CSqlStatement
'Author: Juan Carlos Alvarez
'**************************************************

Private m_sqlString As String
Private m_rs As CResultset

Public Property Get SqlString() As String
    SqlString = m_sqlString
End Property

Public Property Let SqlString(ByVal sql As String)
    m_sqlString = sql
End Property

Private Sub Class_Initialize()
  Set m_rs = New CResultset
End Sub

Public Sub addSqlClause(ByVal sqlClause As String)
  SqlString = SqlString & sqlClause
End Sub

Public Sub addSqlStatement(ByVal statement As CSqlStatement)

End Sub

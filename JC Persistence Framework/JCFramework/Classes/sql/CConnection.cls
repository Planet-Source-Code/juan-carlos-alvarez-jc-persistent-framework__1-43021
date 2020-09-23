VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CConnection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'**************************************************
'Class: CConnection
'Author: Juan Carlos Alvarez
'**************************************************

Private m_connection As ADODB.Connection
Private m_autoCommit As Boolean
Private m_started As Boolean
Private m_manageTransactions As Boolean
Private Const DEBUG_MODE As Boolean = True

Public Property Get Connection() As ADODB.Connection
  Set Connection = m_connection
End Property

Public Property Set Connection(ByVal conn As ADODB.Connection)
  Set m_connection = conn
End Property

Public Property Get AutoCommit() As Boolean
  AutoCommit = m_autoCommit
End Property

Public Property Let AutoCommit(ByVal auto As Boolean)
  m_autoCommit = auto
End Property

Public Property Get Started() As Boolean
  Started = m_started
End Property

Public Property Let Started(ByVal st As Boolean)
  m_started = st
End Property

Public Property Get ManageTransactions() As Boolean
  ManageTransactions = m_manageTransactions
End Property

Public Property Let ManageTransactions(ByVal manT As Boolean)
  m_manageTransactions = manT
End Property

Public Sub commit()
  If DEBUG_MODE Then
    Debug.Print "COMMIT TRANSACTION"
  End If
  If Me.ManageTransactions Then
    m_connection.CommitTrans
  End If
End Sub

Public Sub startTransaction()
  If Not Me.Started Then
    If DEBUG_MODE Then
      Debug.Print "BEGIN TRANSACTION"
    End If
    If Me.ManageTransactions Then
      m_connection.BeginTrans
    End If
    Me.Started = True
  End If
End Sub

Public Sub processStatement(ByVal statement As CSqlStatement)
'TO DO: Handle errors
  Me.startTransaction
  If DEBUG_MODE Then
    Debug.Print statement.SqlString
  End If
  Me.Connection.Execute statement.SqlString
End Sub

Public Function processSelectStatement(ByVal statement As CSqlStatement) As CResultset
'TO DO: Handle errors
  If DEBUG_MODE Then
    Debug.Print statement.SqlString
  End If
  'Me.startTransaction
  Dim rs As New CResultset
  Set rs.ResultSet = Me.Connection.Execute(statement.SqlString)

  Set processSelectStatement = rs
End Function

Public Sub rollback()
  If DEBUG_MODE Then
    Debug.Print "ROLLBACK TRANSACTION"
  End If
  If Me.ManageTransactions Then
    m_connection.RollbackTrans
  End If
End Sub

Private Sub Class_Initialize()
  'By default we manage tranactions
  'The DB that not manage transactions should
  'set this attribute to false
  m_manageTransactions = True
End Sub

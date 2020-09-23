VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CMsAccessDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'**************************************************
'Class: CMsAccessDatabase
'Author: Juan Carlos Alvarez
'**************************************************

Implements CRelationalDatabase

Private m_relDB As CRelationalDatabase

Private m_name As String
Private m_user As String
Private m_password As String

Private Sub Class_Initialize()
  Set m_relDB = New CRelationalDatabase
End Sub

Private Property Set CRelationalDatabase_ConnectionPool(ByVal RHS As CStack)
  Set m_relDB.ConnectionPool = RHS
End Property

Private Property Get CRelationalDatabase_ConnectionPool() As CStack
  Set CRelationalDatabase_ConnectionPool = m_relDB.ConnectionPool
End Property

Private Sub CRelationalDatabase_freeConnection(ByVal conn As CConnection)
  m_relDB.freeConnection conn
End Sub

Private Function CRelationalDatabase_getClauseStringAnd() As String
  CRelationalDatabase_getClauseStringAnd = m_relDB.getClauseStringAnd
End Function

Private Function CRelationalDatabase_getClauseStringAscend() As String
  CRelationalDatabase_getClauseStringAscend = m_relDB.getClauseStringAscend
End Function

Private Function CRelationalDatabase_getClauseStringBetween() As String
  CRelationalDatabase_getClauseStringBetween = m_relDB.getClauseStringBetween
End Function

Private Function CRelationalDatabase_getClauseStringDelete() As String
  CRelationalDatabase_getClauseStringDelete = m_relDB.getClauseStringDelete
End Function

Private Function CRelationalDatabase_getClauseStringDescend() As String
  CRelationalDatabase_getClauseStringDescend = m_relDB.getClauseStringDescend
End Function

Private Function CRelationalDatabase_getClauseStringEqualTo() As String
  CRelationalDatabase_getClauseStringEqualTo = m_relDB.getClauseStringEqualTo
End Function

Private Function CRelationalDatabase_getClauseStringForUpdate() As String
  CRelationalDatabase_getClauseStringForUpdate = m_relDB.getClauseStringForUpdate
End Function

Private Function CRelationalDatabase_getClauseStringFrom() As String
  CRelationalDatabase_getClauseStringFrom = m_relDB.getClauseStringFrom
End Function

Private Function CRelationalDatabase_getClauseStringGroupBy() As String
  CRelationalDatabase_getClauseStringGroupBy = m_relDB.getClauseStringGroupBy
End Function

Private Function CRelationalDatabase_getClauseStringHaving() As String
  CRelationalDatabase_getClauseStringHaving = m_relDB.getClauseStringHaving
End Function

Private Function CRelationalDatabase_getClauseStringIn() As String
  CRelationalDatabase_getClauseStringIn = m_relDB.getClauseStringIn
End Function

Private Function CRelationalDatabase_getClauseStringInsert() As String
  CRelationalDatabase_getClauseStringInsert = m_relDB.getClauseStringInsert
End Function

Private Function CRelationalDatabase_getClauseStringIs() As String
  CRelationalDatabase_getClauseStringIs = m_relDB.getClauseStringIs
End Function

Private Function CRelationalDatabase_getClauseStringLike() As String
  CRelationalDatabase_getClauseStringLike = m_relDB.getClauseStringLike
End Function

Private Function CRelationalDatabase_getClauseStringLimit() As String
  CRelationalDatabase_getClauseStringLimit = m_relDB.getClauseStringLimit
End Function

Private Function CRelationalDatabase_getClauseStringNot() As String
  CRelationalDatabase_getClauseStringNot = m_relDB.getClauseStringNot
End Function

Private Function CRelationalDatabase_getClauseStringOr() As String
  CRelationalDatabase_getClauseStringOr = m_relDB.getClauseStringOr
End Function

Private Function CRelationalDatabase_getClauseStringOrderBy() As String
  CRelationalDatabase_getClauseStringOrderBy = m_relDB.getClauseStringOrderBy
End Function

Private Function CRelationalDatabase_getClauseStringSelect() As String
  CRelationalDatabase_getClauseStringSelect = m_relDB.getClauseStringSelect
End Function

Private Function CRelationalDatabase_getClauseStringSet() As String
  CRelationalDatabase_getClauseStringSet = m_relDB.getClauseStringSet
End Function

Private Function CRelationalDatabase_getClauseStringUpdate() As String
  CRelationalDatabase_getClauseStringUpdate = m_relDB.getClauseStringUpdate
End Function

Private Function CRelationalDatabase_getClauseStringValues() As String
  CRelationalDatabase_getClauseStringValues = m_relDB.getClauseStringValues
End Function

Private Function CRelationalDatabase_getClauseStringWhere() As String
  CRelationalDatabase_getClauseStringWhere = m_relDB.getClauseStringWhere
End Function

Private Function CRelationalDatabase_getConnection(ByVal relDb As CRelationalDatabase) As CConnection
  Set CRelationalDatabase_getConnection = m_relDB.getConnection(Me)
End Function

Private Function CRelationalDatabase_getNewConnection() As CConnection
  'Se establece una nueva conexion con la base de datos Access97
  Dim conn As New CConnection
  Set conn.Connection = New ADODB.Connection
  conn.Connection.Provider = "Microsoft.Jet.OLEDB.4.0"
  conn.Connection.properties("Data Source") = m_name
  conn.Connection.properties("Jet OLEDB:Database Password") = m_password
  conn.Connection.Open
  Set CRelationalDatabase_getNewConnection = conn
End Function

Private Sub CRelationalDatabase_init(ByVal properties As Scripting.IDictionary)
  m_name = properties.Item("name")
  m_user = properties.Item("user")
  m_password = properties.Item("password")
End Sub

Private Property Let CRelationalDatabase_Name(ByVal RHS As String)
  m_relDB.Name = RHS
End Property

Private Property Get CRelationalDatabase_Name() As String
  CRelationalDatabase_Name = m_relDB.Name
End Property

Private Function CRelationalDatabase_getValueFor(ByVal tempValue As Variant) As String
  Dim tempRetorno As String
  If IsNullAlias(tempValue) Then
    tempRetorno = "NULL"
  Else
    Select Case VarType(tempValue)
      Case vbString
        ' Chr(34) = "
        tempRetorno = Chr(34) + Replace(tempValue, Chr(34), "'") + Chr(34)
      Case vbDate
        ' Chr(35) = #
        tempRetorno = Chr(35) + Format(tempValue, "mm/dd/yy hh:mm:ss") + Chr(35)
      Case adBoolean
        tempRetorno = IIf(tempValue, "TRUE", "FALSE")
      Case vbSingle, vbDouble, vbCurrency, vbDecimal
        tempRetorno = Replace(tempValue, ",", ".")
      Case Else
        tempRetorno = tempValue
    End Select
  End If
  CRelationalDatabase_getValueFor = tempRetorno
End Function

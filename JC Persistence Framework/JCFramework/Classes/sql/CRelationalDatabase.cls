VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRelationalDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'**************************************************
'Class: CRelationalDatabase
'Author: Juan Carlos Alvarez
'**************************************************

Private m_connectionPool As CStack
Private m_name As String

Public Property Get Name() As String
  Name = m_name
End Property

Public Property Let Name(ByVal n As String)
  m_name = n
End Property

Public Property Get ConnectionPool() As CStack
  Set ConnectionPool = m_connectionPool
End Property

Public Property Set ConnectionPool(ByVal c As CStack)
  Set m_connectionPool = c
End Property

Private Sub Class_Initialize()
  Set m_connectionPool = New CStack
End Sub

Private Sub Class_Terminate()
  Set m_connectionPool = Nothing
End Sub

Public Sub freeConnection(ByVal conn As CConnection)
  m_connectionPool.Push conn
  conn.Connection.Close
End Sub

Public Function getClauseStringAnd() As String
  getClauseStringAnd = "AND"
End Function

Public Function getClauseStringAscend() As String
  getClauseStringAscend = "ASC"
End Function

Public Function getClauseStringBetween() As String
  getClauseStringBetween = "BETWEEN"
End Function
    
Public Function getClauseStringDelete() As String
  getClauseStringDelete = "DELETE"
End Function
    
Public Function getClauseStringDescend() As String
  getClauseStringDescend = "DESC"
End Function
    
Public Function getClauseStringEqualTo() As String
  getClauseStringEqualTo = "="
End Function
   
Public Function getClauseStringForUpdate() As String
  getClauseStringForUpdate = "FOR UPDATE"
End Function
   
Public Function getClauseStringFrom() As String
  getClauseStringFrom = "FROM"
End Function

Public Function getClauseStringGroupBy() As String
  getClauseStringGroupBy = "GROUP BY"
End Function

Public Function getClauseStringHaving() As String
  getClauseStringHaving = "HAVING"
End Function

Public Function getClauseStringIn() As String
  getClauseStringIn = "IN"
End Function

Public Function getClauseStringInsert() As String
  getClauseStringInsert = "INSERT INTO"
End Function

Public Function getClauseStringIs() As String
  getClauseStringIs = "IS"
End Function

Public Function getClauseStringLike() As String
  getClauseStringLike = "LIKE"
End Function

Public Function getClauseStringLimit() As String
  getClauseStringLimit = "LIMIT"
End Function

Public Function getClauseStringNot() As String
  getClauseStringNot = "NOT"
End Function

Public Function getClauseStringOr() As String
  getClauseStringOr = "OR"
End Function

Public Function getClauseStringOrderBy() As String
  getClauseStringOrderBy = "ORDER BY"
End Function

Public Function getClauseStringSelect() As String
  getClauseStringSelect = "SELECT"
End Function

Public Function getClauseStringSet() As String
  getClauseStringSet = "SET"
End Function

Public Function getClauseStringUpdate() As String
  getClauseStringUpdate = "UPDATE"
End Function

Public Function getClauseStringValues() As String
  getClauseStringValues = "VALUES"
End Function

Public Function getClauseStringWhere() As String
  getClauseStringWhere = "WHERE"
End Function

Public Function getConnection(ByVal relDb As CRelationalDatabase) As CConnection
  Dim conn As CConnection
  Dim i As Integer
  
  Set conn = Nothing
  'La conección sera tomada del pool o sera creada si el pool esta vacio
  Do While ((Not (m_connectionPool.isEmpty)) And (conn Is Nothing))
    Set conn = m_connectionPool.Pop
    'Si esta ya comiteada entonces la libero
    If (conn.Connection.State = adStateClosed) Then
      Set conn = Nothing
    End If
  Loop
  If conn Is Nothing Then
    Set conn = relDb.getNewConnection()
  End If
  Set getConnection = conn
End Function
    
'Crea y retorna la nueva coneccion a esta base de datos.
'Esta funcion es abstracta (se definirá en las clases
'que implementen esta clase)
Public Function getNewConnection() As CConnection

End Function

Public Sub init(ByVal properties As Scripting.Dictionary)

End Sub

'Metodo abstracto que permite obtener los valores formateados
'para la base de datos en cuestión
Public Function getValueFor(ByVal tempValue As Variant) As String

End Function

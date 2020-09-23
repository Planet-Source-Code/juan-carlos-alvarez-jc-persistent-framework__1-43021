VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CRetrieveCriteria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'**************************************************
'Class: CRetrieveCriteria
'Author: Juan Carlos Alvarez
'**************************************************

Implements CPersistentCriteria
Private m_persCriteria As CPersistentCriteria

Private m_groupAttributes As Collection
Private m_orderAttributes As Collection
Private m_offset As Integer
Private m_rows As Integer
Private m_Having As CCriteriaCondition

Public Property Get Offset() As Integer
  Offset = m_offset
End Property

Public Property Let Offset(ByVal offs As Integer)
  m_offset = offs
End Property

Public Property Get Rows() As Integer
  Rows = m_rows
End Property

Public Property Let Rows(ByVal r As Integer)
  m_rows = r
End Property

'Properties para implementar la herencia
Public Property Get ClassName() As String
    ClassName = CPersistentCriteria_ClassName
End Property

Public Property Let ClassName(ByVal n As String)
    CPersistentCriteria_ClassName = n
End Property

Public Property Set ClassMap(ByVal clm As CClassMap)
  If WhereCondition Is Nothing Then
    'Agrego las tablas de este classmap
    Dim cm As CClassMap
    Set cm = clm
    Do
      'Agrego una sola tabla por cada ClassMap
      Tables.Add cm.Tables.Item(1)
      Set cm = cm.SuperClass
    Loop While Not cm Is Nothing
    'Crear la condicion para el WHERE de este criteria
    Set WhereCondition = getNewCondition
    Set WhereCondition.ClassMap = clm
  End If
  Set CPersistentCriteria_ClassMap = clm
End Property

Public Property Get ClassMap() As CClassMap
  Set ClassMap = CPersistentCriteria_ClassMap
End Property

Public Property Set Criteria(ByVal cri As Collection)
  Set CPersistentCriteria_Criteria = cri
End Property

Public Property Get Criteria() As Collection
  Set Criteria = CPersistentCriteria_Criteria
End Property

Public Property Set Tables(ByVal ta As Collection)
  Set CPersistentCriteria_Tables = ta
End Property

Public Property Get Tables() As Collection
  Set Tables = CPersistentCriteria_Tables
End Property

Public Property Set Associations(ByVal assoc As Collection)
  Set CPersistentCriteria_Associations = assoc
End Property

Public Property Get Associations() As Collection
  Set Associations = CPersistentCriteria_Associations
End Property

Public Property Set WhereCondition(ByVal cri As CCriteriaCondition)
  Set CPersistentCriteria_WhereCondition = cri
End Property

Public Property Get WhereCondition() As CCriteriaCondition
  Set WhereCondition = CPersistentCriteria_WhereCondition
End Property

Private Sub Class_Initialize()
  Set m_persCriteria = New CPersistentCriteria
  Set m_groupAttributes = New Collection
  Set m_orderAttributes = New Collection
  m_offset = -1
  m_rows = -1
  Set m_Having = Nothing
End Sub

Private Sub Class_Terminate()
  Set m_persCriteria = Nothing
  Set m_groupAttributes = Nothing
  Set m_orderAttributes = Nothing
  If Not m_Having Is Nothing Then
    Set m_Having = Nothing
  End If
End Sub

'Adds attribute to GROUP BY part of the statement.
Public Sub addGroupAttribute(ByVal attributeName As String)
  m_groupAttributes.Add m_persCriteria.ClassMap.getAttributeMapByString(attributeName, True)
End Sub

'Adds attribute to ORDER BY part of the statement with ascending direction of sorting.
Public Sub addOrderAttribute(ByVal attributeName As String)
  addOrderAttributeByAscend attributeName, True
End Sub

'Adds attribute to ORDER BY part of the statement with the specified direction of sorting.
Public Sub addOrderAttributeByAscend(ByVal attributeName As String, ByVal pAscend As Boolean)
  Dim orderEntry As COrderEntry
  Set orderEntry = New COrderEntry
  orderEntry.AttributeMap = m_persCriteria.ClassMap.getAttributeMapByString(attributeName, True)
  orderEntry.Ascend = pAscend
  m_orderAttributes.Add orderEntry
End Sub

Public Function getNewCondition() As CCriteriaCondition
  Dim criCond As CCriteriaCondition
  Set criCond = New CCriteriaCondition
  Set criCond.ClassMap = Me.ClassMap
  Set criCond.Tables = Me.Tables
  Set criCond.Associations = Me.Associations
  Set getNewCondition = criCond
End Function

Public Function getSqlStatementParameters(ByVal colParameters As Collection) As CSqlStatement
  Dim statement As CSqlStatement
  Dim i As Integer
    
  Set statement = Me.ClassMap.getSelectSql
  'Agregar la clausula FROM de acuerdo a las tablas involucradas
  statement.addSqlClause " " + Me.ClassMap.RelationalDatabase.getClauseStringFrom + " "
  Dim tm As CTableMap
  For i = 1 To Me.Tables.Count
    Set tm = Me.Tables.Item(i)
    If i > 1 Then
      statement.addSqlClause ", "
    End If
    statement.addSqlClause tm.Name
  Next i
  
  fillStatementWithWhere statement, colParameters
  
  Dim attMap As CAttributeMap
  If m_groupAttributes.Count > 0 Then
    'Agregar la clausula GROUP BY
    statement.addSqlClause " " + Me.ClassMap.RelationalDatabase.getClauseStringGroupBy
    For i = 0 To i < m_groupAttributes.Count
      Set attMap = m_groupAttributes.Item(i)
      If i > 0 Then
        statement.addSqlClause ", "
      Else
        statement.addSqlClause " "
      End If
      statement.addSqlClause attMap.ColumnMap.getFullyQualifiedName
    Next i
  End If
  
  Dim entry As COrderEntry
  If m_orderAttributes.Count > 0 Then
    'Agregar la clausula ORDER BY
    statement.addSqlClause " " + Me.ClassMap.RelationalDatabase.getClauseStringOrderBy
    For i = 0 To i < m_orderAttributes.Count
      Set entry = m_orderAttributes.Item(i)
      If i > 0 Then
        statement.addSqlClause ", "
      Else
        statement.addSqlClause " "
      End If
      statement.addSqlClause entry.AttributeMap.ColumnMap.getFullyQualifiedName
      If entry.Ascend Then
        statement.addSqlClause Me.ClassMap.RelationalDatabase.getClauseStringAscend
      Else
        statement.addSqlClause Me.ClassMap.RelationalDatabase.getClauseStringDescend
      End If
    Next i
  End If
  
  If m_rows > 0 Then
    'Agregar la clausula LIMIT
    statement.addSqlClause " " + Me.ClassMap.RelationalDatabase.getClauseStringLimit
    If m_offset >= 0 Then
      statement.addSqlClause Str(m_offset) + ", "
    End If
    statement.addSqlClause Str(m_rows)
  End If
  
  Set getSqlStatementParameters = statement
End Function

Public Function perform(ByVal obj As CPersistentObject, ByVal colParameters As Collection) As CCursor
  Static persistentBroker As CPersistenceBroker
  Set persistentBroker = getPersistenceBrokerInstance
  Set perform = persistentBroker.Instance.processRetrieveCriteria(obj, Me, colParameters)
End Function

Public Sub fillStatementWithWhere(ByVal statement As CSqlStatement, ByVal colParameters As Collection)
  m_persCriteria.fillStatementWithWhere statement, colParameters
End Sub

'Metodos heredados
Private Property Set CPersistentCriteria_Associations(ByVal RHS As Collection)
  Set m_persCriteria.Associations = RHS
End Property

Private Property Get CPersistentCriteria_Associations() As Collection
  Set CPersistentCriteria_Associations = m_persCriteria.Associations
End Property

Private Property Let CPersistentCriteria_ClassName(ByVal RHS As String)
  Let m_persCriteria.ClassName = RHS
End Property

Private Property Get CPersistentCriteria_ClassName() As String
  CPersistentCriteria_ClassName = m_persCriteria.ClassName
End Property

Private Property Set CPersistentCriteria_ClassMap(ByVal RHS As CClassMap)
  Set m_persCriteria.ClassMap = RHS
End Property

Private Property Get CPersistentCriteria_ClassMap() As CClassMap
  Set CPersistentCriteria_ClassMap = m_persCriteria.ClassMap
End Property

Private Property Set CPersistentCriteria_Criteria(ByVal RHS As Collection)
  Set m_persCriteria.Criteria = RHS
End Property

Private Property Get CPersistentCriteria_Criteria() As Collection
  Set CPersistentCriteria_Criteria = m_persCriteria.Criteria
End Property

Private Sub CPersistentCriteria_fillStatementWithWhere(statement As CSqlStatement, ByVal colParameters As Collection)
  m_persCriteria.fillStatementWithWhere statement, colParameters
End Sub

Private Property Set CPersistentCriteria_Tables(ByVal RHS As Collection)
  Set m_persCriteria.Tables = RHS
End Property

Private Property Get CPersistentCriteria_Tables() As Collection
  Set CPersistentCriteria_Tables = m_persCriteria.Tables
End Property

Private Sub CPersistentCriteria_fillStatementWithCondition(statement As CSqlStatement)
  m_persCriteria.fillStatementWithCondition statement
End Sub

Private Function CPersistentCriteria_getSize() As Integer
  CPersistentCriteria_getSize = m_persCriteria.getSize
End Function

Private Sub CPersistentCriteria_setClassMap(pClassName As String)
  m_persCriteria.setClassMap pClassName
End Sub

Private Sub CPersistentCriteria_addOrCriteria(ByVal orCriteria As CPersistentCriteria)
  m_persCriteria.addOrCriteria orCriteria
End Sub

Private Sub CPersistentCriteria_addSelectEqualTo(ByVal theAttrMap As CAttributeMap)
  m_persCriteria.addSelectEqualTo theAttrMap
End Sub

Private Sub CPersistentCriteria_addSelectGreaterThan(ByVal theAttrMap As CAttributeMap)
  m_persCriteria.addSelectGreaterThan theAttrMap
End Sub

Private Sub CPersistentCriteria_addSelectGreaterThanOrEqualTo(ByVal theAttrMap As CAttributeMap)
  m_persCriteria.addSelectGreaterThanOrEqualTo theAttrMap
End Sub

Private Sub CPersistentCriteria_addSelectIn(ByVal attrs As Collection)
  m_persCriteria.addSelectIn attrs
End Sub

Private Property Set CPersistentCriteria_WhereCondition(ByVal RHS As CCriteriaCondition)
  Set m_persCriteria.WhereCondition = RHS
End Property

Private Property Get CPersistentCriteria_WhereCondition() As CCriteriaCondition
  Set CPersistentCriteria_WhereCondition = m_persCriteria.WhereCondition
End Property

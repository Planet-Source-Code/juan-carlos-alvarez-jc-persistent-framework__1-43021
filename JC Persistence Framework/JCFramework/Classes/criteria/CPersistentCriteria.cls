VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPersistentCriteria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'**************************************************
'Class: CPersistentCriteria
'Author: Juan Carlos Alvarez
'**************************************************

Private m_criteria As Collection
Private m_className As String
Private m_classMap As CClassMap
Private m_Tables As Collection
Private m_Associations As Collection
Private m_WhereCondition As CCriteriaCondition

Public Property Get Criteria() As Collection
  Set Criteria = m_criteria
End Property

Public Property Set Criteria(ByVal cri As Collection)
  Set m_criteria = cri
End Property

Public Property Get Tables() As Collection
  Set Tables = m_Tables
End Property

Public Property Set Tables(ByVal ta As Collection)
  Set m_Tables = ta
End Property

Public Property Get Associations() As Collection
  Set Associations = m_Associations
End Property

Public Property Set Associations(ByVal assoc As Collection)
  Set m_Associations = assoc
End Property

Public Property Get ClassName() As String
    ClassName = m_className
End Property

Public Property Let ClassName(ByVal n As String)
    m_className = n
End Property

Public Property Get ClassMap() As CClassMap
  Set ClassMap = m_classMap
End Property

Public Property Set ClassMap(ByVal clm As CClassMap)
  Set m_classMap = clm
End Property

Public Property Get WhereCondition() As CCriteriaCondition
  Set WhereCondition = m_WhereCondition
End Property

Public Property Set WhereCondition(ByVal cri As CCriteriaCondition)
  Set m_WhereCondition = cri
End Property

Public Sub setClassMap(pClassName As String)
  Dim persistentBroker As CPersistenceBroker
  Set persistentBroker = getPersistenceBrokerInstance
  
  m_className = pClassName
  Set m_classMap = persistentBroker.getClassMap(m_className)
End Sub

Public Sub addOrCriteria(ByVal orCriteria As CPersistentCriteria)
  m_criteria.Add orCriteria
End Sub

Public Sub addSelectEqualTo(ByVal theAttrMap As CAttributeMap)
  Dim equalTo As CEqualToCriteria
  Set equalTo = New CEqualToCriteria
  equalTo.AttrMap = theAttrMap
  m_criteria.Add equalTo
End Sub

Public Sub addSelectGreaterThan(ByVal theAttrMap As CAttributeMap)
  Dim greaterThan As CEqualToCriteria
  Set greaterThan = New CEqualToCriteria
  greaterThan.AttrMap = theAttrMap
  m_criteria.Add greaterThan
End Sub

Public Sub addSelectGreaterThanOrEqualTo(ByVal theAttrMap As CAttributeMap)
  Dim greaterThanOrEqualTo As CGreaterThanOrEqualTo
  Set greaterThanOrEqualTo = New CGreaterThanOrEqualTo
  greaterThanOrEqualTo.AttrMap = theAttrMap
  m_criteria.Add greaterThanOrEqualTo
End Sub

Public Sub addSelectIn(ByVal attrs As Collection)
  Dim selectIn As CInCriteria
  Set selectIn = New CInCriteria
  Set selectIn.Attributes = attrs
  m_criteria.Add selectIn
End Sub

Public Sub fillStatementWithCondition(ByRef statement As CSqlStatement)

End Sub

Public Sub fillStatementWithWhere(ByRef statement As CSqlStatement, ByVal colParameters As Collection)
  'Consruimos el 'WHERE'
  Dim inheritedAssoc As String
  inheritedAssoc = Me.ClassMap.getInheritedAssociations
  If ((m_WhereCondition.getSize > 0) Or (Len(inheritedAssoc) > 0)) Then
    'Agregar la clave WHERE al statement
    statement.addSqlClause " " + Me.ClassMap.RelationalDatabase.getClauseStringWhere + " "
    'Agregar partes de herencia y asociaciones
    If ((m_Associations.Count > 0) Or (Len(inheritedAssoc) > 0)) Then
      'Agregar la parte de herencia
      If Len(inheritedAssoc) > 0 Then
        statement.addSqlClause inheritedAssoc
      End If
      'Agregar la parte de asociaciones
      If m_Associations.Count > 0 Then
        If Len(inheritedAssoc) > 0 Then
          statement.addSqlClause " " + Me.ClassMap.RelationalDatabase.getClauseStringAnd + " "
        End If
      End If
      If m_WhereCondition.getSize > 0 Then
        statement.addSqlClause " " + Me.ClassMap.RelationalDatabase.getClauseStringAnd + " "
      End If
    ElseIf m_WhereCondition.getSize > 0 Then
      statement.addSqlClause " "
    End If
    WhereCondition.fillStatement statement, colParameters
  End If
End Sub

Public Function getSize() As Integer
  getSize = m_criteria.Count
End Function

Private Sub Class_Initialize()
  Set m_criteria = New Collection
  Set m_Tables = New Collection
  Set m_WhereCondition = Nothing
  Set m_Associations = New Collection
End Sub

Private Sub Class_Terminate()
  Set m_criteria = Nothing
  Set m_Tables = Nothing
  Set m_Associations = Nothing
  If Not m_WhereCondition Is Nothing Then
    Set m_WhereCondition = Nothing
  End If
End Sub

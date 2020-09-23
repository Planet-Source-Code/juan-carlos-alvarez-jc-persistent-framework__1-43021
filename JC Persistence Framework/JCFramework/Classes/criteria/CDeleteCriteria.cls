VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CDeleteCriteria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'**************************************************
'Class: CDeleteCriteria
'Author: Juan Carlos Alvarez
'**************************************************

Implements CPersistentCriteria
Private m_persCriteria As CPersistentCriteria

'Properties para implementar la herencia
Public Property Get ClassName() As String
    ClassName = CPersistentCriteria_ClassName
End Property

Public Property Let ClassName(ByVal n As String)
    CPersistentCriteria_ClassName = n
End Property

Public Property Set ClassMap(ByVal clm As CClassMap)
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

Private Sub Class_Initialize()
  Set m_persCriteria = New CPersistentCriteria
End Sub

Public Function perform(ByVal colParameters As Collection) As CCursor
  Static persistentBroker As CPersistenceBroker
  Set persistentBroker = getPersistenceBrokerInstance
  persistentBroker.Instance.processDeleteCriteria Me, colParameters
End Function

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

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCriteriaCondition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Condition for criteria's WHERE and HAVING parts

Private m_Simple As Boolean
Private m_Parts As Collection
Private m_Associations As Collection
Private m_classMap As CClassMap
Private m_Tables As Collection

Public Property Set ClassMap(ByVal clm As CClassMap)
  Set m_classMap = clm
End Property

Public Property Get ClassMap() As CClassMap
  Set ClassMap = m_classMap
End Property

Public Property Set Tables(ByVal ta As Collection)
  Set m_Tables = ta
End Property

Public Property Get Tables() As Collection
  Set Tables = m_Tables
End Property

Public Property Set Associations(ByVal assoc As Collection)
  Set m_Associations = assoc
End Property

Public Property Get Associations() As Collection
  Set Associations = m_Associations
End Property

Private Sub Class_Initialize()
  m_Simple = True
  Set m_Parts = New Collection
End Sub

Public Sub addOrCriteria(ByVal orCriteria As CCriteriaCondition)
  m_Simple = False
  m_Parts.Add orCriteria
End Sub

Public Sub addSelectEqualTo(ByVal Name As String)
  Dim am As CAttributeMap
  Set am = getAttributeMap(Name)
  Dim eqToCri As CEqualToCriteria
  Set eqToCri = New CEqualToCriteria
  Set eqToCri.AttrMap = am
  m_Parts.Add eqToCri
End Sub

Private Function getAttributeMap(ByVal attributeName As String) As CAttributeMap
  Dim AttrMap As CAttributeMap
  Dim myArrayStrings As Variant
  myArrayStrings = Split(attributeName, ".", -1, vbTextCompare)
  Dim clMap As CClassMap
  Set clMap = m_classMap
  
  Dim i As Integer, j As Integer, strName As String
  Dim udaMap As CUDAMap
  For i = 0 To UBound(myArrayStrings) - 1
    strName = myArrayStrings(i)
    Set udaMap = clMap.AssociationMaps(strName)
    m_Associations.Add udaMap
    
    Set clMap = udaMap.ForClass
    If Not clMap Is Nothing Then
      For j = 1 To clMap.Tables.Count
        m_Tables.Add clMap.Tables.Item(j)
      Next j
    Else
      Exit For
    End If
  Next i
  
  If Not clMap Is Nothing Then
    Set AttrMap = clMap.getAttributeMapByString(myArrayStrings(UBound(myArrayStrings)), True)
  End If
  
  Set getAttributeMap = AttrMap
End Function

Public Sub addSelectGreaterThan(ByVal Name As String)
  Dim am As CAttributeMap
  Set am = getAttributeMap(Name)
  Dim grThanCri As New CGreaterThanCriteria
  Set grThanCri.AttrMap = am
  m_Parts.Add grThanCri
End Sub

Public Sub addSelectGreaterThanOrEqualTo(ByVal Name As String)
  Dim am As CAttributeMap
  Set am = getAttributeMap(Name)
  Dim grThanOrEqToCri As New CGreaterThanOrEqualTo
  Set grThanOrEqToCri.AttrMap = am
  m_Parts.Add grThanOrEqToCri
End Sub

Public Sub addSelectIn(ByVal Name As String)
  Dim am As CAttributeMap
  Set am = getAttributeMap(Name)
  Dim selectIn As New CInCriteria
  Set selectIn.AttrMap = am
  m_Parts.Add selectIn
End Sub

Public Sub addSelectLessThan(ByVal Name As String)
  Dim am As CAttributeMap
  Set am = getAttributeMap(Name)
  Dim selectLessThan As New CLessThanCriteria
  Set selectLessThan.AttrMap = am
  m_Parts.Add selectLessThan
End Sub

Public Sub addSelectLessThanOrEqualTo(ByVal Name As String)
  Dim am As CAttributeMap
  Set am = getAttributeMap(Name)
  Dim selLessThanOrEq As New CLessThanOrEqualToCriteria
  Set selLessThanOrEq.AttrMap = am
  m_Parts.Add selLessThanOrEq
End Sub

Public Sub addSelectLike(ByVal Name As String)
  Dim am As CAttributeMap
  Set am = getAttributeMap(Name)
  Dim selectLike As New CLikeCriteria
  Set selectLike.AttrMap = am
  m_Parts.Add selectLike
End Sub

Public Sub addSelectNotEqualTo(ByVal Name As String)
  Dim am As CAttributeMap
  Set am = getAttributeMap(Name)
  Dim selectNotEqualTo As New CNotEqualToCriteria
  Set selectNotEqualTo.AttrMap = am
  m_Parts.Add selectNotEqualTo
End Sub

Public Function getSize() As Integer
  getSize = m_Parts.Count
End Function

Public Sub fillStatement(ByRef statement As CSqlStatement, ByVal colParameters As Collection)
  If Not m_Simple Then
    statement.addSqlClause "("
  End If
  'Agrego el criterio
  Dim contador As Integer, i As Integer
  Dim objPart As Object, selCri As CSelectionCriteria
  For i = 1 To Me.getSize
    Set objPart = m_Parts.Item(i)
    If TypeName(objPart) = "CSelectionCriteria" Or _
      TypeName(objPart) = "CEqualToCriteria" Or _
      TypeName(objPart) = "CNotEqualToCriteria" Then
      ' o es uno de sus hijos (debo hardcodear por no soportar herencia)
      If contador = 0 Then
        statement.addSqlClause "("
      Else
        statement.addSqlClause " " + Me.ClassMap.RelationalDatabase.getClauseStringAnd + " "
      End If
      Set selCri = m_Parts.Item(i)
      selCri.fillSqlStatement colParameters(i), statement, Me.ClassMap
      contador = contador + 1
      'MsgBox "selection criteria"
    ElseIf TypeName(objPart) = "CCriteriaCondition" Then
      'MsgBox "criteria condition"
    End If
    'MsgBox TypeName(objPart)
    
  Next i
  
  If contador > 0 Then
    statement.addSqlClause ")"
  End If
  If Not m_Simple Then
    statement.addSqlClause ")"
  End If
End Sub

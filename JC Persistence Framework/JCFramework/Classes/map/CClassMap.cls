VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CClassMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'**************************************************
'Class: CClassMap
'Author: Juan Carlos Alvarez
'**************************************************

Private m_name As String
Private m_selectStatement As CSqlStatement
Private m_insertStatement As CSqlStatement
Private m_updateStatement As CSqlStatement
Private m_deleteStatement As CSqlStatement
Private m_attributeMaps As Collection
Private m_hashedAttributeMaps As Scripting.Dictionary
Private m_keyAttributeMaps As Collection
Private m_referenceAttributeMaps As Collection
Private m_Tables As Collection
Private m_mapObjectClass As CPersistentObject
Private m_associationMaps As Scripting.Dictionary
Private m_relationalDatabase As CRelationalDatabase
Private m_superClass As CClassMap
Private m_inverseAssociationMaps As Collection
Private m_straightAssociationMaps As Collection

Public Property Get Name() As String
    Name = m_name
End Property

Public Property Let Name(ByVal n As String)
    m_name = n
End Property

Public Property Get SelectStatement() As CSqlStatement
  Set SelectStatement = m_selectStatement
End Property

Public Property Set SelectStatement(ByVal statement As CSqlStatement)
  Set m_selectStatement = statement
End Property

Public Property Get InsertStatement() As CSqlStatement
  Set InsertStatement = m_insertStatement
End Property

Public Property Set InsertStatement(ByVal statement As CSqlStatement)
  Set m_insertStatement = statement
End Property

Public Property Get DeleteStatement() As CSqlStatement
  Set DeleteStatement = m_deleteStatement
End Property

Public Property Set DeleteStatement(ByVal statement As CSqlStatement)
  Set m_deleteStatement = statement
End Property

Public Property Get KeyAttributeMaps() As Collection
  Set KeyAttributeMaps = m_keyAttributeMaps
End Property

Public Property Set KeyAttributeMaps(ByVal keyAttrMaps As Collection)
  Set m_keyAttributeMaps = keyAttrMaps
End Property

Public Property Get AssociationMaps() As Scripting.Dictionary
  Set AssociationMaps = m_associationMaps
End Property

Public Property Set AssociationMaps(ByVal assocMaps As Scripting.Dictionary)
  Set m_associationMaps = assocMaps
End Property

Public Property Get HashedAttributeMaps() As Scripting.Dictionary
  Set HashedAttributeMaps = m_hashedAttributeMaps
End Property

Public Property Set HashedAttributeMaps(ByVal hashAttrMaps As Scripting.Dictionary)
  Set m_hashedAttributeMaps = hashAttrMaps
End Property

Public Property Get ReferenceAttributeMaps() As Collection
  Set ReferenceAttributeMaps = m_referenceAttributeMaps
End Property

Public Property Set ReferenceAttributeMaps(ByVal refAttrMaps As Collection)
  Set m_referenceAttributeMaps = refAttrMaps
End Property

Public Property Get Tables() As Collection
  Set Tables = m_Tables
End Property

Public Property Set Tables(ByVal tbls As Collection)
  Set m_Tables = tbls
End Property

Public Property Get AttributeMaps() As Collection
  Set AttributeMaps = m_attributeMaps
End Property

Public Property Set AttributeMaps(ByVal attrMaps As Collection)
  Set m_attributeMaps = attrMaps
End Property

Public Property Get MapObjectClass() As CPersistentObject
  Set MapObjectClass = m_mapObjectClass
End Property

Public Property Set MapObjectClass(ByVal mapClass As CPersistentObject)
  Set m_mapObjectClass = mapClass
End Property

Public Property Get SuperClass() As CClassMap
  Set SuperClass = m_superClass
End Property

Public Property Set SuperClass(ByVal supClass As CClassMap)
  Set m_superClass = supClass
End Property

Public Property Get RelationalDatabase() As CRelationalDatabase
  Dim persistenceBroker As CPersistenceBroker
  If m_relationalDatabase Is Nothing Then
    If Me.AttributeMaps.Count > 0 Then
      Dim AttributeMap As CAttributeMap
      'Get the first attribute
      Set AttributeMap = m_attributeMaps.Item(1)
      Set persistenceBroker = getPersistenceBrokerInstance
      Set m_relationalDatabase = persistenceBroker.getRelationalDatabase( _
        AttributeMap.ColumnMap.TableMap.DatabaseMap.Name)
    End If
  End If
  Set RelationalDatabase = m_relationalDatabase
End Property

Public Property Set RelationalDatabase(ByVal relDb As CRelationalDatabase)
  Set m_relationalDatabase = relDb
End Property

Public Property Get InverseAssociationMaps() As Collection
  Set InverseAssociationMaps = m_inverseAssociationMaps
End Property

Public Property Set InverseAssociationMaps(ByVal inv As Collection)
  Set m_inverseAssociationMaps = inv
End Property

Public Property Get StraightAssociationMaps() As Collection
  Set StraightAssociationMaps = m_straightAssociationMaps
End Property

Public Property Set StraightAssociationMaps(ByVal straight As Collection)
  Set m_straightAssociationMaps = straight
End Property

Private Sub Class_Initialize()
  Set m_attributeMaps = New Collection
  Set m_keyAttributeMaps = New Collection
  Set m_hashedAttributeMaps = CreateObject("Scripting.Dictionary")
  Set m_associationMaps = CreateObject("Scripting.Dictionary")
  Set m_referenceAttributeMaps = New Collection
  Set m_Tables = New Collection
  Set m_inverseAssociationMaps = New Collection
  Set m_straightAssociationMaps = New Collection
End Sub

Public Sub addAssociationMap(ByVal map As CUDAMap)
  m_associationMaps.Add map.Target.Name, map
  If map.Inverse Then
    m_inverseAssociationMaps.Add map
  Else
    m_straightAssociationMaps.Add map
  End If
End Sub

Public Sub addAttributeMap(ByVal AttrMap As CAttributeMap)
  m_hashedAttributeMaps.Add AttrMap.Name, AttrMap
  If Not AttrMap.ColumnMap Is Nothing Then
    m_attributeMaps.Add AttrMap
    If AttrMap.ColumnMap.KeyType <> ColumnKeyTypes.KEY_NONE Then
      m_keyAttributeMaps.Add AttrMap
    End If
    If Not AttrMap.AttributeMap Is Nothing Then
      m_referenceAttributeMaps.Add AttrMap
    End If
    m_Tables.Add AttrMap.ColumnMap.TableMap
  End If
End Sub

Public Function getAttributeMap(ByVal index As Integer) As CAttributeMap
  Dim AttrMap As CAttributeMap
  Set AttrMap = Me.AttributeMaps.Item(index)
  Set getAttributeMap = AttrMap
End Function

Public Function getAttributeMapByString(ByVal Name As String, ByVal areSuperClassesIncluded As Boolean) As CAttributeMap
  Dim am As CAttributeMap
  Dim cm As CClassMap
  Dim dictTemp As Scripting.Dictionary
  Set cm = Me
  Do
    Set dictTemp = cm.HashedAttributeMaps
    Set am = dictTemp.Item(Name)
    Set cm = cm.SuperClass
  Loop While ((areSuperClassesIncluded) And (am Is Nothing) And (Not cm Is Nothing))
  Set getAttributeMapByString = am
End Function

Public Function getDeleteSqlFor(ByVal obj As CPersistentObject) As CSqlStatement
  If Not m_deleteStatement Is Nothing Then
    Set m_deleteStatement = Nothing
  End If
  If m_deleteStatement Is Nothing Then
    'Create new SqlStatement
    Set m_deleteStatement = New CSqlStatement
    'Add 'DELETE FROM' clause to the select statement
    m_deleteStatement.addSqlClause Me.RelationalDatabase.getClauseStringDelete + " " + Me.RelationalDatabase.getClauseStringFrom + " "
    Dim map As CAttributeMap
    Set map = Me.getAttributeMap(1)
    If Not map Is Nothing Then
      m_deleteStatement.addSqlClause map.ColumnMap.TableMap.Name + " "
    End If
    'Add 'WHERE key= "some value"' to the select statement
    m_deleteStatement.addSqlClause Me.RelationalDatabase.getClauseStringWhere + " "
    Dim i As Integer
    Dim AttrMap As CAttributeMap
    For i = 1 To Me.getKeySize
      Set AttrMap = Me.getKeyAttributeMap(i)
      Set map = Me.AttributeMaps.Item(i)
      m_deleteStatement.addSqlClause map.ColumnMap.Name + "=" + _
        m_relationalDatabase.getValueFor(obj.getValueByAttribute(AttrMap.Name))
    Next i
  End If
  Set getDeleteSqlFor = m_deleteStatement
End Function

Public Function getKeyAttributeMap(ByVal index As Integer) As CAttributeMap
  Dim AttrMap As CAttributeMap
  Set AttrMap = m_keyAttributeMaps.Item(index)
  Set getKeyAttributeMap = AttrMap
End Function

Public Function getKeySize() As Integer
  getKeySize = Me.KeyAttributeMaps.Count
End Function

Public Function getReferenceAttributeMap(ByVal index As Integer) As CAttributeMap
  Dim AttrMap As CAttributeMap
  Set AttrMap = m_referenceAttributeMaps.Item(index)
  Set getReferenceAttributeMap = AttrMap
End Function

Public Function getReferenceSize() As Integer
  getReferenceSize = m_referenceAttributeMaps.Count
End Function

Public Function getSelectSql() As CSqlStatement
  'Create new SqlStatement
  Dim statement As New CSqlStatement
  'Add the select clause
  statement.addSqlClause Me.RelationalDatabase.getClauseStringSelect + " "
  'Add clauses for all attributes
  Dim isFirst As Boolean
  isFirst = True
  Dim ClassMap As CClassMap
  Set ClassMap = Me
  Dim i As Integer
  Do
    For i = 1 To ClassMap.getSize
      If isFirst Then
        statement.addSqlClause "" + ClassMap.getAttributeMap(i).ColumnMap.getFullyQualifiedName
      Else
        statement.addSqlClause ", " + ClassMap.getAttributeMap(i).ColumnMap.getFullyQualifiedName
      End If
      isFirst = False
    Next i
    Set ClassMap = ClassMap.SuperClass
  Loop While Not ClassMap Is Nothing
  Set getSelectSql = statement
End Function

Public Function getSelectSqlFor(ByVal obj As CPersistentObject) As CSqlStatement
  If Not m_selectStatement Is Nothing Then
    Set m_selectStatement = Nothing
  End If
  If m_selectStatement Is Nothing Then
    'Create new SqlStatement
    Set m_selectStatement = getSelectSql()
    
    'Add the FROM clause to the statement
    m_selectStatement.addSqlClause " " + RelationalDatabase.getClauseStringFrom + " "
    Dim isFirst As Boolean
    isFirst = True
    Dim ClassMap As CClassMap
    Set ClassMap = Me
    Dim AttrMap As CAttributeMap
    Do
      Set AttrMap = ClassMap.getAttributeMap(1)
      If Not AttrMap Is Nothing Then
        If isFirst Then
          m_selectStatement.addSqlClause "" + AttrMap.ColumnMap.TableMap.Name
        Else
          m_selectStatement.addSqlClause ", " + AttrMap.ColumnMap.TableMap.Name
        End If
      End If
      Set ClassMap = ClassMap.SuperClass
      isFirst = False
    Loop While Not ClassMap Is Nothing
    m_selectStatement.addSqlClause " "
      
    'Add the clause "WHERE key=?" to the selectStatement
    m_selectStatement.addSqlClause RelationalDatabase.getClauseStringWhere + " "
    Dim i As Integer
    For i = 1 To Me.getKeySize
      Set AttrMap = Me.getKeyAttributeMap(i)
      If i > 1 Then
        m_selectStatement.addSqlClause " " + RelationalDatabase.getClauseStringAnd + " " + _
          Me.getKeyAttributeMap(i).ColumnMap.getFullyQualifiedName + "=" + _
            m_relationalDatabase.getValueFor(obj.getValueByAttribute(AttrMap.Name))
      Else
        m_selectStatement.addSqlClause "" + _
          Me.getKeyAttributeMap(i).ColumnMap.getFullyQualifiedName + "=" + _
            m_relationalDatabase.getValueFor(obj.getValueByAttribute(AttrMap.Name))
      End If
    Next i
    'Add the clause "WHERE tabla.columna=superclass_table.columna to the selectStatement
    Set ClassMap = Me
    Do
      For i = 1 To ClassMap.getReferenceSize
        m_selectStatement.addSqlClause " " + RelationalDatabase.getClauseStringAnd + " " + _
          ClassMap.getReferenceAttributeMap(i).ColumnMap.getFullyQualifiedName + "=" + _
          ClassMap.getReferenceAttributeMap(i).AttributeMap.ColumnMap.getFullyQualifiedName
      Next i
      Set ClassMap = ClassMap.SuperClass
    Loop While Not ClassMap Is Nothing
  End If
  Set getSelectSqlFor = m_selectStatement
End Function

Public Function getInsertSqlFor(ByVal obj As CPersistentObject) As CSqlStatement
  If Not m_insertStatement Is Nothing Then
    Set m_insertStatement = Nothing
  End If
  If m_insertStatement Is Nothing Then
    'Create new SqlStatement
    Set m_insertStatement = New CSqlStatement
    'Add the clause "INSERT INTO" to the insertStatement
    m_insertStatement.addSqlClause RelationalDatabase.getClauseStringInsert + " "
    Dim AttrMap As CAttributeMap
    Set AttrMap = getAttributeMap(1)
    If Not AttrMap Is Nothing Then
      m_insertStatement.addSqlClause AttrMap.ColumnMap.TableMap.Name + " "
    End If
    'Add clauses for all attributes
    m_insertStatement.addSqlClause "("
    Dim i As Integer
    For i = 1 To Me.getSize
      If i > 1 Then
        m_insertStatement.addSqlClause ", " + _
          Me.getAttributeMap(i).ColumnMap.Name
      Else
        'No agrego la coma si es el primer atributo
        'Don´t add a coma if it is the first attribute
        m_insertStatement.addSqlClause "" + _
          Me.getAttributeMap(i).ColumnMap.Name
      End If
    Next i
    m_insertStatement.addSqlClause ") "
    'Agrego la clausula "VALUES" al insertStatement
    'Add clause "VALUES" to the insertStatement
    m_insertStatement.addSqlClause RelationalDatabase.getClauseStringValues + " "
    m_insertStatement.addSqlClause "("
    For i = 1 To Me.getSize
      Set AttrMap = Me.getAttributeMap(i)
      If i > 1 Then
        m_insertStatement.addSqlClause ", " + _
          m_relationalDatabase.getValueFor(obj.getValueByAttribute(AttrMap.Name))
      Else
        'No agrego la coma si es el primer atributo
        'Don´t add a coma if it is the first attribute
        m_insertStatement.addSqlClause "" + _
          m_relationalDatabase.getValueFor(obj.getValueByAttribute(AttrMap.Name))
      End If
    Next i
    m_insertStatement.addSqlClause ")"
  End If
  Set getInsertSqlFor = m_insertStatement
End Function

Public Function getUpdateSqlFor(ByVal obj As CPersistentObject) As CSqlStatement
  If Not m_updateStatement Is Nothing Then
    Set m_updateStatement = Nothing
  End If
  If m_updateStatement Is Nothing Then
    'Creo un nuevo statement
    Set m_updateStatement = New CSqlStatement
    'Agrego la clausula "UPDATE " al insertStatement
    m_updateStatement.addSqlClause RelationalDatabase.getClauseStringUpdate + " "
    Dim AttrMap As CAttributeMap
    Set AttrMap = getAttributeMap(1)
    If Not AttrMap Is Nothing Then
      m_updateStatement.addSqlClause AttrMap.ColumnMap.TableMap.Name + " "
    End If
    'Agrego las clausulas para todos los atributos
    m_updateStatement.addSqlClause RelationalDatabase.getClauseStringSet + " "
    Dim i As Integer
    For i = 1 To Me.getSize
      Set AttrMap = Me.getAttributeMap(i)
      If i > 1 Then
        m_updateStatement.addSqlClause ", " + _
          Me.getAttributeMap(i).ColumnMap.Name + "=" + _
            m_relationalDatabase.getValueFor(obj.getValueByAttribute(AttrMap.Name))
      Else
        'No agrego la coma si es el primer atributo
        m_updateStatement.addSqlClause "" + _
          Me.getAttributeMap(i).ColumnMap.Name + "=" + _
            m_relationalDatabase.getValueFor(obj.getValueByAttribute(AttrMap.Name))
      End If
    Next i
    'Agrego la clausula WHERE
    m_updateStatement.addSqlClause " " + RelationalDatabase.getClauseStringWhere + " "
    For i = 1 To Me.getKeySize
      Set AttrMap = Me.getKeyAttributeMap(i)
      If i > 1 Then
        m_updateStatement.addSqlClause " " + RelationalDatabase.getClauseStringAnd + " " + _
          Me.getKeyAttributeMap(i).ColumnMap.getFullyQualifiedName + "=" + _
            m_relationalDatabase.getValueFor(obj.getValueByAttribute(AttrMap.Name))
      Else
        m_updateStatement.addSqlClause "" + _
          Me.getKeyAttributeMap(i).ColumnMap.getFullyQualifiedName + "=" + _
            m_relationalDatabase.getValueFor(obj.getValueByAttribute(AttrMap.Name))
      End If
    Next i
  End If
  Set getUpdateSqlFor = m_updateStatement
End Function

Public Sub retrieveObject(ByRef obj As CPersistentObject, ByVal rs As CResultset)
  Dim ClassMap As CClassMap
  Set ClassMap = Me
  Dim i As Integer, AttrMap As CAttributeMap
  Do
    For i = 1 To ClassMap.getSize
      Set AttrMap = ClassMap.getAttributeMap(i)
      obj.setAttributeValue AttrMap.Name, rs.ResultSet.Fields(i - 1).Value
    Next i
    Set ClassMap = ClassMap.SuperClass
  Loop While Not ClassMap Is Nothing
  obj.Persistent = True
End Sub

Public Function getSize() As Integer
  getSize = m_attributeMaps.Count
End Function

Public Function getStraightAssociationMapSize() As Integer
  getStraightAssociationMapSize = m_straightAssociationMaps.Count
End Function

Public Function getInverseAssociationMapSize() As Integer
  getInverseAssociationMapSize = m_inverseAssociationMaps.Count
End Function

Public Function getStraightAssociationMap(ByVal index As Integer) As CUDAMap
  Dim udaMap As CUDAMap
  Set udaMap = Me.StraightAssociationMaps.Item(index)
  Set getStraightAssociationMap = udaMap
End Function

Public Function getInverseAssociationMap(ByVal index As Integer) As CUDAMap
  Dim udaMap As CUDAMap
  Set udaMap = Me.InverseAssociationMaps.Item(index)
  Set getInverseAssociationMap = udaMap
End Function

'Retorna 'tabla.columna=superclase_tabla.columna'
Public Function getInheritedAssociations() As String
  Dim strTemp As String
  Dim i As Integer
  Dim clMap As CClassMap
  Set clMap = Me
  Do
    For i = 1 To clMap.getReferenceSize
      If i > 1 Then
        strTemp = strTemp + " " + Me.RelationalDatabase.getClauseStringAnd + " "
      End If
    Next i
    Set clMap = clMap.SuperClass
  Loop While Not clMap Is Nothing
End Function

Public Function getValueForRelationalDatabase(ByVal tempValue As Variant) As String
  getValueForRelationalDatabase = m_relationalDatabase.getValueFor(tempValue)
End Function

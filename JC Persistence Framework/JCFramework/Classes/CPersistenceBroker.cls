VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPersistenceBroker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'**************************************************
'Class: CPersistenceBroker
'Author: Juan Carlos Alvarez
'**************************************************

Private m_instance As CPersistenceBroker
Private m_classMaps As Scripting.Dictionary
Private m_databases As Scripting.Dictionary

Public Static Property Get Instance() As CPersistenceBroker
  Set Instance = getPersistenceBrokerInstance
End Property

Public Static Property Set Instance(ByVal inst As CPersistenceBroker)
  Set m_instance = inst
End Property

Public Property Get ClassMaps() As Scripting.Dictionary
  Set ClassMaps = m_classMaps
End Property

Public Property Set ClassMaps(ByVal clMaps As Scripting.Dictionary)
  Set m_classMaps = clMaps
End Property

Public Property Get Databases() As Scripting.Dictionary
  Set Databases = m_databases
End Property

Public Property Set Databases(ByVal dbs As Scripting.Dictionary)
  Set m_databases = dbs
End Property

Private Static Sub Class_Initialize()
  Set m_instance = Nothing
  Set m_classMaps = CreateObject("Scripting.Dictionary")
  Set m_databases = CreateObject("Scripting.Dictionary")
End Sub

Private Sub Class_Terminate()
  Set m_instance = Nothing
  Set m_classMaps = Nothing
  Set m_databases = Nothing
End Sub

Public Sub retrieveObject(ByRef obj As CPersistentObject)
  On Error GoTo rollBack_Trans
  Dim cm As CClassMap
  Set cm = obj.getClassMap(obj)
  Dim conn As CConnection
  Set conn = cm.RelationalDatabase.getConnection(Nothing)
  conn.AutoCommit = False
  retrievePrivateObject obj, conn
  'conn.commit
  cm.RelationalDatabase.freeConnection conn
  Exit Sub
rollBack_Trans:
  'conn.rollback
  cm.RelationalDatabase.freeConnection conn
End Sub

Private Sub retrievePrivateObject(ByRef obj As CPersistentObject, ByVal conn As CConnection)
  Dim cm As CClassMap
  Set cm = obj.getClassMap(obj)
  Dim statement As CSqlStatement
  Set statement = cm.getSelectSqlFor(obj)
  Dim rs As CResultset
  Set rs = conn.processSelectStatement(statement)
  'Si encontro el objeto buscado ...
  If Not rs.ResultSet.EOF Then
    cm.retrieveObject obj, rs
  End If
  'Traer los datos asociados al objeto (sus referencias)
  Dim i As Integer, j As Integer, k As Integer
  Dim colCriteriaParameters As New Collection
  Dim udaMap As CUDAMap
  Dim aCriteria As CRetrieveCriteria
  Dim cursor As CCursor
  Dim claves() As Variant
  Dim Value As CPersistentObject
  Dim ValueVar As String
  Dim col As Collection
  claves = cm.AssociationMaps.Keys
  For i = 0 To cm.AssociationMaps.Count - 1
    Set udaMap = cm.AssociationMaps.Item(claves(i))
    If Not udaMap.RetrieveAutomatic Then
      Exit For
    End If
    
    'Se agregan los parametros para los criterios
    Set aCriteria = New CRetrieveCriteria
    Set aCriteria.ClassMap = udaMap.ForClass
    For j = 1 To udaMap.getSize
      If udaMap.Inverse Then
        'TODO: Implement this functionality
        'Set anAttr = udaMap.getEntry(j).FromAttrMap
        'aCriteria.WhereCondition.addSelectEqualTo udaMap.getEntry(j).FromAttrMap.Name
        'Set Value = obj.getObjectByAttribute(udaMap.getEntry(j).ToAttrMap.Name)
        'colCriteriaParameters.Add Value
      Else
        aCriteria.WhereCondition.addSelectEqualTo udaMap.getEntry(j).ToAttrMap.Name
        If udaMap.Cardinality = CardinalityEnum.ONE_TO_ONE Then
          Set Value = obj.getObjectByAttribute(udaMap.Target.Name)
          If Not Value Is Nothing Then
            Value.retrieve Value
            obj.setAttributeValue udaMap.Target.Name, Value
          End If
        ElseIf udaMap.Cardinality = CardinalityEnum.ONE_TO_MANY Then
          ValueVar = cm.getValueForRelationalDatabase( _
            obj.getValueByAttribute(udaMap.getEntry(j).FromAttrMap.Name))
          colCriteriaParameters.Add ValueVar
          'Se obtiene el cursor dados todos los parametros y los criterios
          Set cursor = processCriteria(aCriteria, colCriteriaParameters, conn)
          'Se agregan los objetos a la coleccion
          Set col = obj.getCollectionByAttribute(udaMap.Target.Name)
          Set Value = col.Item(1)
          Dim anObjPers As CPersistentObject
          While cursor.hasElements
            Set anObjPers = Value.getNewObject
            cursor.loadObject anObjPers
            col.Add anObjPers
            cursor.nextCursor
          Wend
          col.Remove (1)
        End If
      End If
    Next j
  Next i
  Set colCriteriaParameters = Nothing
End Sub

Public Sub saveObject(ByRef obj As CPersistentObject)
  On Error GoTo rollBack_Trans
  Dim cm As CClassMap
  Set cm = obj.getClassMap(obj)
  Dim conn As CConnection
  Set conn = cm.RelationalDatabase.getConnection(Nothing)
  conn.AutoCommit = False
  savePrivateObject obj, conn
  conn.commit
  cm.RelationalDatabase.freeConnection conn
  Exit Sub
rollBack_Trans:
  conn.rollback
  cm.RelationalDatabase.freeConnection conn
End Sub

Private Sub savePrivateObject(ByRef obj As CPersistentObject, ByVal conn As CConnection)
  Dim clMap As CClassMap
  Set clMap = obj.getClassMap(obj)
  Dim classMapStack As CStack
  Set classMapStack = New CStack
  Dim cm As CClassMap
  Set cm = clMap
  
  'Save straight relationships
  Dim i, j, k, m As Integer
  Dim udaMap As CUDAMap
  Dim Value As CPersistentObject, col As Collection
  For i = 1 To cm.getStraightAssociationMapSize
    Set udaMap = cm.getStraightAssociationMap(i)
    If Not udaMap.SaveAutomatic Then
      Exit For
    End If
    If udaMap.Cardinality = CardinalityEnum.ONE_TO_ONE Then
      Set Value = obj.getObjectByAttribute(udaMap.Target.Name)
      If Not Value Is Nothing Then
        savePrivateObject Value, conn
        For j = 1 To udaMap.getSize
          Set Value = obj.getObjectByAttribute(udaMap.Target.Name)
        Next j
      End If
    ElseIf udaMap.Cardinality = CardinalityEnum.ONE_TO_MANY Then
      Set col = obj.getCollectionByAttribute(udaMap.Target.Name)
      If Not col Is Nothing Then
        For k = 1 To col.Count
          Set Value = col.Item(k)
          savePrivateObject Value, conn
        Next k
      End If
    End If
  Next i
  
  Dim statement As CSqlStatement
  If obj.Persistent Then
    'Actualizo el objeto en la base
    Do
      classMapStack.Push cm
      Set cm = cm.SuperClass
    Loop While Not cm Is Nothing
    Do While Not classMapStack.isEmpty
      Set cm = classMapStack.Pop
      Set statement = clMap.getUpdateSqlFor(obj)
      Dim rs As CResultset
      conn.processStatement statement
    Loop
  Else
    'Insertar nueva entrada en la base de datos
    Do
      classMapStack.Push cm
      Set cm = cm.SuperClass
    Loop While Not cm Is Nothing
    Do While Not classMapStack.isEmpty
      Set cm = classMapStack.Pop
      Set statement = clMap.getInsertSqlFor(obj)
      conn.processStatement statement
    Loop
  End If
  
  'Save inverse relationships
  For i = 1 To cm.getInverseAssociationMapSize
    Set udaMap = cm.getInverseAssociationMap(i)
    If Not udaMap.SaveAutomatic Then
      Exit For
    End If
    If udaMap.Cardinality = CardinalityEnum.ONE_TO_ONE Then
      Set Value = obj.getObjectByAttribute(udaMap.Target.Name)
      If Not Value Is Nothing Then
        savePrivateObject Value, conn
        For j = 1 To udaMap.getSize
          Set Value = obj.getObjectByAttribute(udaMap.getEntry(j).FromAttrMap.Name)
        Next j
      End If
    ElseIf udaMap.Cardinality = CardinalityEnum.ONE_TO_MANY Then
      Set col = obj.getCollectionByAttribute(udaMap.Target.Name)
      If Not col Is Nothing Then
        For k = 1 To col.Count
          Set Value = col.Item(k)
          savePrivateObject Value, conn
          For m = 1 To udaMap.getSize
            Set Value = obj.getObjectByAttribute(udaMap.getEntry(m).FromAttrMap.Name)
          Next m
        Next k
      End If
    End If
  Next i
  'El objeto ya esta en la base entonces es Persistente.
  obj.Persistent = True
End Sub

Public Sub deleteObject(ByVal obj As CPersistentObject)
  On Error GoTo rollBack_Trans
  Dim cm As CClassMap
  Set cm = obj.getClassMap(obj)
  Dim conn As CConnection
  Set conn = cm.RelationalDatabase.getConnection(Nothing)
  Dim statement As CSqlStatement
  Set statement = cm.getDeleteSqlFor(obj)
  conn.AutoCommit = False
  deletePrivateObject obj, conn
  conn.commit
  cm.RelationalDatabase.freeConnection conn
  Exit Sub
rollBack_Trans:
  conn.rollback
  cm.RelationalDatabase.freeConnection conn
End Sub

Private Sub deletePrivateObject(ByRef obj As CPersistentObject, ByVal conn As CConnection)
  Dim clMap As CClassMap
  Set clMap = obj.getClassMap(obj)
  'Eliminar el objeto de la base
  Dim cm As CClassMap
  Set cm = clMap
  
  'Retrieve associations and delete if deleteAutomatic
  Dim udaMap As CUDAMap
  Dim aCriteria As CRetrieveCriteria
  Dim cursor As CCursor
  Dim claves() As Variant
  Dim Value As CPersistentObject
  Dim ValueVar As String
  Dim col As Collection
  Dim colCriteriaParameters As New Collection
  Dim i As Integer, j As Integer
  claves = cm.AssociationMaps.Keys
  For i = 0 To cm.AssociationMaps.Count - 1
    Set udaMap = cm.AssociationMaps.Item(claves(i))
    If Not udaMap.DeleteAutomatic Then
      Exit For
    End If
    
    'Se agregan los parametros para los criterios
    Set aCriteria = New CRetrieveCriteria
    Set aCriteria.ClassMap = udaMap.ForClass
    For j = 1 To udaMap.getSize
      If udaMap.Inverse Then
        'TODO: Implement this functionality
        'Set anAttr = udaMap.getEntry(j).FromAttrMap
        'aCriteria.WhereCondition.addSelectEqualTo udaMap.getEntry(j).FromAttrMap.Name
        'Set Value = obj.getObjectByAttribute(udaMap.getEntry(j).ToAttrMap.Name)
        'colCriteriaParameters.Add Value
      Else
        aCriteria.WhereCondition.addSelectEqualTo udaMap.getEntry(j).ToAttrMap.Name
        If udaMap.Cardinality = CardinalityEnum.ONE_TO_ONE Then
          Set Value = obj.getObjectByAttribute(udaMap.Target.Name)
          If Not Value Is Nothing Then
            Value.retrieve Value
            deletePrivateObject Value, conn
          End If
        ElseIf udaMap.Cardinality = CardinalityEnum.ONE_TO_MANY Then
          ValueVar = cm.getValueForRelationalDatabase( _
            obj.getValueByAttribute(udaMap.getEntry(j).FromAttrMap.Name))
          colCriteriaParameters.Add ValueVar
          'Se obtiene el cursor dados todos los parametros y los criterios
          Set cursor = processCriteria(aCriteria, colCriteriaParameters, conn)
          'Se agregan los objetos a la coleccion
          Set col = obj.getCollectionByAttribute(udaMap.Target.Name)
          If Not col Is Nothing Then
            If col.Count > 0 Then
              Set Value = col.Item(1)
              Dim anObjPers As CPersistentObject
              While cursor.hasElements
                Set anObjPers = Value.getNewObject
                cursor.loadObject anObjPers
                deletePrivateObject anObjPers, conn
                cursor.nextCursor
              Wend
            End If
            Set col = Nothing
          End If
        End If
      End If
    Next j
  Next i
  
  Do
    'Elimino el objeto
    Dim statement As CSqlStatement
    Set statement = clMap.getDeleteSqlFor(obj)
    conn.processStatement statement
    Set cm = cm.SuperClass
  Loop While Not cm Is Nothing
  'El objeto ya no esta más en la base de datos por lo que no es persistente
  obj.Persistent = False
  
  Set colCriteriaParameters = Nothing
End Sub

Public Function getClassMap(ByVal ClassName As String) As CClassMap
  Dim ClassMap As CClassMap
  Set ClassMap = m_classMaps.Item(ClassName)
  Set getClassMap = ClassMap
End Function

Public Function getRelationalDatabase(ByVal Name As String) As CRelationalDatabase
  Dim relDb As CRelationalDatabase
  Set relDb = Me.Databases.Item(Name)
  Set getRelationalDatabase = relDb
End Function

Public Sub init()
  m_classMaps.RemoveAll
  m_databases.RemoveAll
  Set m_instance = getPersistenceBrokerInstance
End Sub

Public Sub loadConfig(ByVal configLoader As IConfigLoader)
  configLoader.loadConfig m_databases, m_classMaps
End Sub

' This method is called from the DeleteCriteria perform method
' to delete multiple objects from the database.
Public Function processDeleteCriteria(ByRef pCriteria As CDeleteCriteria, ByVal colParameters As Collection) As Long
  Dim conn As CConnection
  Set conn = pCriteria.ClassMap.RelationalDatabase.getConnection(Nothing)
  
  'Creamos un RetrieveCriteria para obtener el objeto a eliminar
  Dim retrieveCriteria As CRetrieveCriteria
  Set retrieveCriteria = New CRetrieveCriteria
  Set retrieveCriteria.Criteria = pCriteria
    
  Dim statement As CSqlStatement
  'Set statement = retrieveCriteria.getSqlStatement
  
  conn.AutoCommit = False
  
  Dim cursor As CCursor
  Set cursor = processPrivateCriteria(retrieveCriteria, conn, colParameters)

  Dim contador As Long
  contador = 0
  'Se borran los objetos
On Error GoTo rollBack_Trans
  Dim persObj As CPersistentObject
'  Do While cursor.nextCursor
'    Set persObj = cursor.getObject
    deletePrivateObject persObj, conn
    contador = contador + 1
'  Loop
  conn.commit
  pCriteria.ClassMap.RelationalDatabase.freeConnection conn
  processDeleteCriteria = contador
  Exit Function
rollBack_Trans:
  conn.rollback
End Function

' This method is called from the RetrieveCriteria perform method
' to retrieve multiple objects from the database. Returns Cursor with the result.
Public Function processRetrieveCriteria(ByRef obj As CPersistentObject, ByRef pCriteria As CRetrieveCriteria, ByVal colParameters As Collection) As CCursor
  Dim clMap As CClassMap
  Set clMap = obj.getClassMap(obj)
    
  Dim conn As CConnection
  Set conn = clMap.RelationalDatabase.getConnection(Nothing)
  Set pCriteria.ClassMap = clMap
  
On Error GoTo rollBack_Trans
  Dim statement As CSqlStatement
  Set statement = pCriteria.getSqlStatementParameters(colParameters)
  
  conn.AutoCommit = False
  
  Dim cursor As CCursor
  Set cursor = processPrivateCriteria(pCriteria, colParameters, conn)
  
  Set processRetrieveCriteria = cursor
  Exit Function
rollBack_Trans:
  'conn.rollback
  clMap.RelationalDatabase.freeConnection conn
End Function

Private Function processPrivateCriteria(ByVal pCriteria As CRetrieveCriteria, ByVal colParameters As Collection, ByVal conn As CConnection) As CCursor
  Dim statement As CSqlStatement
  Set statement = pCriteria.getSqlStatementParameters(colParameters)
  
  Dim rs As CResultset
  Set rs = conn.processSelectStatement(statement)
  
  Dim cursor As CCursor
  Set cursor = New CCursor
  Set cursor.ResultSet = rs
  Set cursor.ClassMap = pCriteria.ClassMap
  
  Set processPrivateCriteria = cursor
End Function

Private Function processCriteria(ByVal pCriteria As CRetrieveCriteria, ByVal colCriteriaParameters As Collection, ByVal conn As CConnection) As CCursor
  Dim statement As CSqlStatement
  Set statement = pCriteria.getSqlStatementParameters(colCriteriaParameters)
  
  Dim rs As CResultset
  Set rs = conn.processSelectStatement(statement)
  
  Dim cursor As CCursor
  Set cursor = New CCursor
  Set cursor.ResultSet = rs
  Set cursor.ClassMap = pCriteria.ClassMap
  
  Set processCriteria = cursor
End Function

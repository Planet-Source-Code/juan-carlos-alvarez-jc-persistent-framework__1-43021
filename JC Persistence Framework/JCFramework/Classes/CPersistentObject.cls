VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CPersistentObject"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'**************************************************
'Class: CPersistentObject
'Author: Juan Carlos Alvarez
'**************************************************

Private m_persistent As Boolean

Public Property Get Persistent() As Boolean
  Persistent = m_persistent
End Property

Public Property Let Persistent(ByVal boolValue As Boolean)
  m_persistent = boolValue
End Property

' Funcion que permite que cada objeto pueda setear sus atributos
Public Sub setAttributeValue(ByVal pName As String, ByRef Value As Variant)

End Sub

' Funcion que permite obtener el valor de un objeto dado su nombre de atributo
Public Function getValueByAttribute(ByVal pName As String) As Variant

End Function

' Funcion que permite obtener el objeto dado su nombre de atributo
Public Function getObjectByAttribute(ByVal pName As String) As CPersistentObject

End Function

' Funcion que permite crear un nuevo objeto (se aplica el polimorfismo)
' Es para crear objetos que serán contenidos dentro de una colección
Public Function getNewObject() As CPersistentObject

End Function

' Funcion que permite obtener la coleccion de un objeto dado su nombre de atributo
Public Function getCollectionByAttribute(ByVal pName As String) As Collection

End Function

Public Function getClassMap(ByVal obj As CPersistentObject) As CClassMap
  Dim ClassMap As CClassMap
  Static persistenceBroker As CPersistenceBroker
  Set persistenceBroker = getPersistenceBrokerInstance
  Set ClassMap = persistenceBroker.getClassMap(TypeName(obj))

  If (ClassMap Is Nothing) Then
    MsgBox ("No class map for " + TypeName(obj) + " class")
  End If
  Set getClassMap = ClassMap
End Function

Private Sub Class_Initialize()
  m_persistent = False
End Sub

Public Sub retrieve(ByRef obj As CPersistentObject)
  Static persistentBroker As CPersistenceBroker
  Set persistentBroker = getPersistenceBrokerInstance
  persistentBroker.retrieveObject obj
End Sub

Public Sub save(ByVal obj As CPersistentObject)
  Static persistentBroker As CPersistenceBroker
  Set persistentBroker = getPersistenceBrokerInstance
  persistentBroker.saveObject obj
End Sub

Public Sub delete(ByVal obj As CPersistentObject)
  Static persistentBroker As CPersistenceBroker
  Set persistentBroker = getPersistenceBrokerInstance
  persistentBroker.deleteObject obj
End Sub

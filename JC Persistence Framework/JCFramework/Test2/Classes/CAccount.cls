VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'****************************************************************
'* Code that all classes should has for supports the JCFramework
'****************************************************************

Implements CPersistentObject

Private m_objPers As CPersistentObject

'*****************************************************
'* Personalized Code for this class
'*****************************************************

'Propiedades
Private m_AccountId As Long
Private m_UserId As Long
Private m_Description As String

'---------------------------------------------
'Metodos de acceso y modificacion
'AccountId
Public Property Get AccountId() As Long
    AccountId = m_AccountId
End Property
Public Property Let AccountId(ByVal a As Long)
    m_AccountId = a
End Property

' UserId
Public Property Get UserId() As Long
    UserId = m_UserId
End Property
Public Property Let UserId(ByVal c As Long)
    m_UserId = c
End Property

'Description
Public Property Get Description() As String
    Description = m_Description
End Property
Public Property Let Description(ByVal d As String)
    m_Description = d
End Property

'****************************************************************
'* Code that all classes should has for supports the JCFramework
'****************************************************************
Private Sub Class_Initialize()
  Set m_objPers = New CPersistentObject
End Sub

Private Sub Class_Terminate()
  Set m_objPers = Nothing
End Sub

Public Property Get Persistent() As Boolean
  Persistent = CPersistentObject_Persistent
End Property

Public Property Let Persistent(boolValue As Boolean)
  Let CPersistentObject_Persistent = boolValue
End Property

Public Sub retrieve()
  CPersistentObject_retrieve Me
End Sub

Public Sub save()
  CPersistentObject_save Me
End Sub

Public Sub delete()
  CPersistentObject_delete Me
End Sub

Public Function getClassMap() As JCFramework.CClassMap
  Set getClassMap = CPersistentObject_getClassMap(Nothing)
End Function

Private Function CPersistentObject_getClassMap(ByVal object As JCFramework.CPersistentObject) As JCFramework.CClassMap
  Set CPersistentObject_getClassMap = m_objPers.getClassMap(Me)
End Function

Private Function CPersistentObject_getCollectionByAttribute(ByVal pName As String) As Collection
  Dim tempRetorno As Collection
  Set tempRetorno = CallByName(Me, pName, VbGet)
  Set CPersistentObject_getCollectionByAttribute = tempRetorno
End Function

Private Function CPersistentObject_getNewObject() As JCFramework.CPersistentObject
  Dim anObj As CAccount
  Set anObj = New CAccount
  Set CPersistentObject_getNewObject = anObj
End Function

Private Function CPersistentObject_getObjectFromCollectionByAttribute(ByVal pName As String) As JCFramework.CPersistentObject
  Dim tempRetorno As CPersistentObject
  Dim colAux As Collection
  Set colAux = CallByName(Me, pName, VbGet)
  Set tempRetorno = colAux.Item(0)
  Set CPersistentObject_getObjectFromCollectionByAttribute = tempRetorno
End Function

Private Function CPersistentObject_getObjectByAttribute(ByVal pName As String) As JCFramework.CPersistentObject
  Dim tempRetorno As CPersistentObject
  Set tempRetorno = CallByName(Me, pName, VbGet)
  Set CPersistentObject_getObjectByAttribute = tempRetorno
End Function

Private Function CPersistentObject_getValueByAttribute(ByVal pName As String) As Variant
  CPersistentObject_getValueByAttribute = CallByName(Me, pName, VbGet)
End Function

Private Property Let CPersistentObject_Persistent(ByVal RHS As Boolean)
  m_objPers.Persistent = RHS
End Property

Private Property Get CPersistentObject_Persistent() As Boolean
  CPersistentObject_Persistent = m_objPers.Persistent
End Property

Private Sub CPersistentObject_retrieve(ByRef obj As JCFramework.CPersistentObject)
  m_objPers.retrieve obj
End Sub

Private Sub CPersistentObject_save(ByVal obj As JCFramework.CPersistentObject)
  m_objPers.save obj
End Sub

Private Sub CPersistentObject_delete(ByVal obj As JCFramework.CPersistentObject)
  m_objPers.delete Me
End Sub

Private Sub CPersistentObject_setAttributeValue(ByVal pName As String, ByRef Value As Variant)
  setAttributeValue Me, pName, Value
End Sub



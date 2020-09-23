VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CUser"
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

'****************************************************************
'Personalized Code for this class
'****************************************************************
'Propiedades
Private m_UserId As Long
Private m_Firstname As String
Private m_Lastname As String
Private m_Address As String
Private m_Telephone As String
Private m_Username As String
Private m_Password As String
Private m_DateOfBirth As Date
Private m_DateOfAdded As Date
Private m_EMail As String
Private m_CountryId As Long
Private m_Country As CCountry
Private m_Accounts As Collection

'---------------------------------------------
'Metodos de acceso y modificacion
' UserId
Public Property Get UserId() As Long
    UserId = m_UserId
End Property
Public Property Let UserId(ByVal c As Long)
    m_UserId = c
End Property

'Firstname
Public Property Get Firstname() As String
    Firstname = m_Firstname
End Property
Public Property Let Firstname(ByVal n As String)
    m_Firstname = n
End Property

'Lastname
Public Property Get Lastname() As String
    Lastname = m_Lastname
End Property
Public Property Let Lastname(ByVal n As String)
    m_Lastname = n
End Property

'Address
Public Property Get Address() As String
    Address = m_Address
End Property
Public Property Let Address(ByVal d As String)
    m_Address = d
End Property

'Telephone
Public Property Get Telephone() As String
    Telephone = m_Telephone
End Property
Public Property Let Telephone(ByVal t As String)
    m_Telephone = t
End Property

'Fecha Nacimiento
Public Property Get DateOfBirth() As Date
    DateOfBirth = m_DateOfBirth
End Property
Public Property Let DateOfBirth(ByVal f As Date)
    m_DateOfBirth = f
End Property

'Username
Public Property Get Username() As String
    Username = m_Username
End Property
Public Property Let Username(ByVal u As String)
    m_Username = u
End Property

'Password
Public Property Get Password() As String
    Password = m_Password
End Property
Public Property Let Password(ByVal c As String)
    m_Password = c
End Property

'Fecha de Ingreso
Public Property Get DateOfAdded() As Date
    DateOfAdded = m_DateOfAdded
End Property
Public Property Let DateOfAdded(ByVal f As Date)
    m_DateOfAdded = f
End Property

'EMail
Public Property Get EMail() As String
    EMail = m_EMail
End Property
Public Property Let EMail(ByVal e As String)
    m_EMail = e
End Property

'Accounts
Public Property Get Accounts() As Collection
  '****************************************************************
  'Personalized Code for this class
  '****************************************************************
  'FOR THE MOMENT (NOT DEFINED YET):
  'If you have a ONE_TO_MANY relation
  'you should create the collection if
  'it is not created and then
  'add a new object of the collection type
  If m_Accounts Is Nothing Then
    Set m_Accounts = New Collection
    Dim anAccount As CAccount
    Set anAccount = New CAccount
    anAccount.UserId = Me.UserId
    m_Accounts.Add anAccount
  End If
  Set Accounts = m_Accounts
End Property
Public Property Set Accounts(ByVal a As Collection)
    Set m_Accounts = a
End Property

'CountryId
Public Property Get CountryId() As Long
    CountryId = m_CountryId
End Property
Public Property Let CountryId(ByVal c As Long)
    m_CountryId = c
End Property

'Country
Public Property Get Country() As CCountry
    ' If you have a reference and it would retrieve automatic
    ' Put the reference value before return the object
    ' It then could be done automatic by the framework
    If m_Country Is Nothing Then
      Set m_Country = New CCountry
    End If
    m_Country.CountryId = Me.CountryId
    Set Country = m_Country
End Property
Public Property Set Country(ByVal c As CCountry)
    Set m_Country = c
End Property

'****************************************************************
'* Code that all classes should has for supports the JCFramework
'****************************************************************
Private Sub Class_Initialize()
  Set m_objPers = New CPersistentObject
  '****************************************************************
  'Personalized Code for this class
  '****************************************************************
  Set m_Country = Nothing
  Set m_Accounts = Nothing
End Sub

Private Sub Class_Terminate()
  Set m_objPers = Nothing
  Set m_Country = Nothing
  Set m_Accounts = Nothing
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
  Dim anObj As CUser
  Set anObj = New CUser
  Set CPersistentObject_getNewObject = anObj
End Function

Private Function CPersistentObject_getObjectByAttribute(ByVal pName As String) As JCFramework.CPersistentObject
  Dim tempRetorno As CPersistentObject
  Set tempRetorno = CallByName(Me, pName, VbGet)
  Set CPersistentObject_getObjectByAttribute = tempRetorno
End Function

Private Function CPersistentObject_getObjectFromCollectionByAttribute(ByVal pName As String) As JCFramework.CPersistentObject
  Dim tempRetorno As CPersistentObject
  Dim colAux As Collection
   Set colAux = CallByName(Me, pName, VbGet)
  Set tempRetorno = colAux.Item(1)
  Set CPersistentObject_getObjectFromCollectionByAttribute = tempRetorno
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


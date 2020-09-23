VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CUDAMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'**************************************************
'Class: CUDAMap
'Author: Juan Carlos Alvarez
'**************************************************

'This class contains information about association map.

Public Enum CardinalityEnum
  ONE_TO_ONE = 0
  ONE_TO_MANY = 1
End Enum

Private m_cardinality As Integer
Private m_must As Boolean
Private m_saveAutomatic As Boolean
Private m_deleteAutomatic As Boolean
Private m_retrieveAutomatic As Boolean
Private m_forClass As CClassMap
Private m_target As CAttributeMap
Private m_entries As Collection
Private m_inverse As Boolean

'Class Properties
Public Property Get Cardinality() As Integer
  Cardinality = m_cardinality
End Property

Public Property Let Cardinality(ByVal c As Integer)
  m_cardinality = c
End Property

Public Property Get Must() As Boolean
  Must = m_must
End Property

Public Property Let Must(ByVal m As Boolean)
  m_must = m
End Property

Public Property Get SaveAutomatic() As Boolean
  SaveAutomatic = m_saveAutomatic
End Property

Public Property Let SaveAutomatic(ByVal s As Boolean)
  m_saveAutomatic = s
End Property

Public Property Get DeleteAutomatic() As Boolean
  DeleteAutomatic = m_deleteAutomatic
End Property

Public Property Let DeleteAutomatic(ByVal d As Boolean)
  m_deleteAutomatic = d
End Property

Public Property Get RetrieveAutomatic() As Boolean
  RetrieveAutomatic = m_retrieveAutomatic
End Property

Public Property Let RetrieveAutomatic(ByVal r As Boolean)
  m_retrieveAutomatic = r
End Property

Public Property Get ForClass() As CClassMap
  Set ForClass = m_forClass
End Property

Public Property Set ForClass(ByVal clm As CClassMap)
  Set m_forClass = clm
End Property

Public Property Get Target() As CAttributeMap
  Set Target = m_target
End Property

Public Property Set Target(ByVal t As CAttributeMap)
  Set m_target = t
End Property

Public Property Get Entries() As Collection
  Set Entries = m_entries
End Property

Public Property Set Entries(ByVal t As Collection)
  Set m_entries = t
End Property

Public Property Get Inverse() As Boolean
  Inverse = m_inverse
End Property

Public Property Let Inverse(ByVal i As Boolean)
  m_inverse = i
End Property

Private Sub Class_Initialize()
  m_cardinality = ONE_TO_ONE
  'Set m_forClass = New CClassMap
  Set m_entries = New Collection
End Sub

Private Sub Class_Terminate()
  Set m_entries = Nothing
End Sub

'Adds new entry to this association map.
Public Sub addEntry(ByVal entry As CUDAMapEntry)
  m_entries.Add entry
End Sub

'Returns entry with the given index.
Public Function getEntry(ByVal index As Integer) As CUDAMapEntry
  Set getEntry = m_entries.Item(index)
End Function

'Returns number of entries.
Public Function getSize() As Integer
  getSize = m_entries.Count
End Function

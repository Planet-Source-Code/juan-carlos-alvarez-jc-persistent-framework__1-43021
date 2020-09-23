VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CResultset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'**************************************************
'Class: CResultset
'Author: Juan Carlos Alvarez
'**************************************************

Private m_rs As ADODB.Recordset
Private m_cmd As ADODB.Command
Private m_modo As ADODB.CursorTypeEnum
Private m_location As ADODB.CursorLocationEnum
Private m_lockType As ADODB.LockTypeEnum
Private m_options As Long

Public Property Get ResultSet() As ADODB.Recordset
  Set ResultSet = m_rs
End Property

Public Property Set ResultSet(ByVal rs As ADODB.Recordset)
  Set m_rs = rs
End Property

Public Property Get Command() As ADODB.Command
  Set Command = m_cmd
End Property

Public Property Set Command(ByVal cmd As ADODB.Command)
  Set m_cmd = cmd
End Property

Public Property Get Modo() As ADODB.CursorTypeEnum
  Modo = m_modo
End Property

Public Property Let Modo(ByVal m As ADODB.CursorTypeEnum)
   m_modo = m
End Property

Public Property Get Location() As ADODB.CursorLocationEnum
  Location = m_location
End Property

Public Property Let Location(ByVal m As ADODB.CursorLocationEnum)
   m_location = m
End Property

Public Property Get LockType() As ADODB.LockTypeEnum
  LockType = m_lockType
End Property

Public Property Let LockType(ByVal m As ADODB.LockTypeEnum)
   m_lockType = m
End Property

Public Property Get Options() As Long
  Options = m_options
End Property

Public Property Let Options(ByVal opts As Long)
   m_options = opts
End Property

Private Sub Class_Initialize()
  Modo = adOpenKeyset
  Location = adUseClient
  LockType = adLockUnspecified
  Options = adCmdText
End Sub

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'**************************************************
'Class: CStack
'Author: Juan Carlos Alvarez
'**************************************************

Private m_stackCollection As Collection

Private Sub Class_Initialize()
    Set m_stackCollection = New Collection
End Sub

Private Sub Class_Terminate()
    Set m_stackCollection = Nothing
End Sub

Public Sub Push(ByVal obj As Object)
  ' Colocamos el item en el Stack
  ' Si el stack no esta vacio colocamos el item
  ' antes del primer item
  If m_stackCollection.Count > 0 Then
    m_stackCollection.Add obj, , 1
  Else
    ' sino lo agregamos al final
    m_stackCollection.Add obj
  End If
End Sub

Public Function Pop() As Object
  ' Quitamos el item del Stack
  If m_stackCollection.Count > 0 Then
    If IsObject(m_stackCollection.Item(1)) Then
      Set Pop = m_stackCollection.Item(1)
    Else
      Pop = m_stackCollection.Item(1)
    End If
    m_stackCollection.Remove 1
  Else
    Err.Raise vbObjectError + 4999, "CStack::Pop", "Stack is empty"
  End If
End Function

Public Property Get StackCount() As Long
  ' Cantidad de items en el Stack
  StackCount = m_stackCollection.Count
End Property

Public Function isEmpty() As Boolean
  If m_stackCollection.Count = 0 Then
    isEmpty = True
  Else
    isEmpty = False
  End If
End Function

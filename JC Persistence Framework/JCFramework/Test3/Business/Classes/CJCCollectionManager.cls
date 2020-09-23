VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CJCCollectionManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function getCollectionFor(ByVal anObjPers As CPersistentObject) As Collection
  'Fill all persistent objects in a Collection
  Dim colObjects As New Collection
  
  Dim retrieveCriteria As CRetrieveCriteria
  Dim cursor As CCursor
  Dim objPersTemp As CPersistentObject
  Set objPersTemp = anObjPers.getNewObject
  
  Dim objPers As CPersistentObject
  Dim indice As Long
  
  Set retrieveCriteria = New CRetrieveCriteria
  indice = 1
  Dim colParams As New Collection
  Set cursor = retrieveCriteria.perform(objPersTemp, colParams)
  While cursor.hasElements
    Set objPers = anObjPers.getNewObject
    cursor.loadObject objPers
    colObjects.Add objPers
    cursor.nextCursor
  Wend
  Set objPersTemp = Nothing
  Set colParams = Nothing
  Set retrieveCriteria = Nothing
  
  Set getCollectionFor = colObjects
End Function

Attribute VB_Name = "modPersistenceBrokerSingleton"
Option Explicit

'**************************************************
'Module: PersistenceBrokerSingleton
'Author: Juan Carlos Alvarez
'**************************************************

Public Static Function getPersistenceBrokerInstance() As CPersistenceBroker
  Static staticInstance As CPersistenceBroker
  If staticInstance Is Nothing Then
    Set staticInstance = New CPersistenceBroker
    Set staticInstance.Instance = staticInstance
    Dim XMLConfigLoader As New CXMLConfigLoader
    staticInstance.loadConfig XMLConfigLoader
  End If
  Set getPersistenceBrokerInstance = staticInstance
End Function


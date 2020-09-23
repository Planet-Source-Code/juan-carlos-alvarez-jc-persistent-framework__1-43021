VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IConfigLoader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'**************************************************
'Interface: IConfigLoader
'Just now we read the ConfigLoader from an XML file
'but we then maybe want to read it from a database
'This is the reason for the interface
'**************************************************

Public Sub loadConfig(ByRef dbMap As Scripting.Dictionary, ByRef clMap As Scripting.Dictionary)
  
End Sub


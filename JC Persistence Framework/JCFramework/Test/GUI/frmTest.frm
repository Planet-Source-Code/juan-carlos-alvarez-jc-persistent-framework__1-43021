VERSION 5.00
Begin VB.Form frmTest 
   Caption         =   "Testing JCFramework"
   ClientHeight    =   3555
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3555
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCollections 
      Caption         =   "Test JCFramework - retrieveObjectsCollection"
      Height          =   555
      Left            =   1140
      TabIndex        =   4
      Top             =   2520
      Width           =   2295
   End
   Begin VB.CommandButton cmdTest 
      Caption         =   "Test JCFramework - retrieve"
      Height          =   375
      Left            =   1140
      TabIndex        =   3
      Top             =   360
      Width           =   2295
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Test JCFramework - delete"
      Height          =   375
      Left            =   1140
      TabIndex        =   2
      Top             =   1980
      Width           =   2295
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Test JCFramework - update"
      Height          =   375
      Left            =   1140
      TabIndex        =   1
      Top             =   1440
      Width           =   2295
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Test JCFramework - save"
      Height          =   375
      Left            =   1140
      TabIndex        =   0
      Top             =   900
      Width           =   2295
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCollections_Click()
  Dim retrieveCriteria As New CRetrieveCriteria
  Dim cursor As CCursor
  Dim persona As New CPersona
  Dim colAux As New Collection
  Set cursor = retrieveCriteria.perform(persona, colAux)
  While cursor.hasElements
    cursor.loadObject persona
    MsgBox Str(persona.Id) + ", " + persona.Name
    cursor.nextCursor
  Wend
  Set colAux = Nothing
  Set cursor = Nothing
  Set persona = Nothing
End Sub

Private Sub cmdDelete_Click()
  Dim persona As CPersona
  Set persona = New CPersona
  persona.Id = 1
  persona.delete
  MsgBox "Objeto eliminado"
End Sub

Private Sub cmdSave_Click()
  Dim persona As CPersona
  Set persona = New CPersona
  persona.Id = 5
  persona.Name = "Pedro"
  'Para controlar que no exista
  'Se intenta traer el objeto
  persona.retrieve
  If persona.Persistent Then
    MsgBox "Ya existe esa persona con ese id"
    Exit Sub
  End If
  persona.save
  persona.retrieve
  MsgBox "Objeto grabado con nombre = " + persona.Name
End Sub

Private Sub cmdTest_Click()
  Dim persona As CPersona
  Set persona = New CPersona
  persona.Id = 3
  persona.retrieve
  MsgBox persona.Name
  
'  If persona.Persistent Then
'    MsgBox "persistente"
'  Else
'    MsgBox "no persistente"
'  End If
End Sub

Private Sub cmdUpdate_Click()
  Dim persona As CPersona
  Set persona = New CPersona
  persona.Id = 3
  persona.retrieve
  MsgBox persona.Name
  persona.Name = "Cambio de nombre"
  persona.save
  persona.retrieve
  MsgBox persona.Name
End Sub

Private Sub Form_Load()
  'Para iniciar la carga del XML automáticamente
  'y no cuando se intenta realizar la primera operación
  Dim persBroker As New CPersistenceBroker
  persBroker.init
End Sub

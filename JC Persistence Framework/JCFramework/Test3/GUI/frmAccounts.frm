VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmAccounts 
   Caption         =   "Accounts"
   ClientHeight    =   3585
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5625
   LinkTopic       =   "Form1"
   ScaleHeight     =   3585
   ScaleWidth      =   5625
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "New"
      Height          =   360
      Left            =   4260
      TabIndex        =   2
      Top             =   180
      Width           =   1200
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   2820
      TabIndex        =   0
      Top             =   3120
      Width           =   1200
   End
   Begin MSFlexGridLib.MSFlexGrid grillaCuentas 
      Height          =   2835
      Left            =   120
      TabIndex        =   1
      ToolTipText     =   "Seleccione un cliente para modificar, eliminar o consultar"
      Top             =   120
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   5001
      _Version        =   393216
      Rows            =   1
      FixedCols       =   0
      AllowBigSelection=   -1  'True
      FocusRect       =   0
      HighLight       =   2
      FillStyle       =   1
      ScrollBars      =   2
      SelectionMode   =   1
      FormatString    =   "<Account id      |<Description                          "
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmAccounts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_Accounts As Collection
Private m_User As CUser

'User
Public Property Get User() As CUser
  Set User = m_User
End Property
Public Property Set User(ByVal u As CUser)
  Set m_User = u
End Property

'Accounts
Public Property Get Accounts() As Collection
  Set Accounts = m_Accounts
End Property
Public Property Set Accounts(ByVal a As Collection)
  Set m_Accounts = a
End Property

Private Sub cmdNuevo_Click()
  frmAccount.Show vbModal
  If Not frmAccount.Account Is Nothing Then
    frmAccount.Account.UserId = m_User.UserId
    m_Accounts.Add frmAccount.Account
    ActualizarAsignadas
  End If
  Unload frmAccount
End Sub

Private Sub Form_Load()
  ActualizarAsignadas
End Sub

Private Sub Form_Unload(Cancel As Integer)
  'Se liberan los objetos creados para las colecciones
  If Not m_Accounts Is Nothing Then
    Set m_Accounts = Nothing
  End If
End Sub

Private Sub cmdOk_Click()
  If m_User.UserId = 0 Then
    MsgBox "Changes don´t take effect until you click on Add button"
  Else
    MsgBox "Changes don´t take effect until you click on Modify button"
  End If
  Me.Hide
End Sub

Private Sub ActualizarAsignadas()
    grillaCuentas.Clear
    grillaCuentas.FormatString = "<Account id      |<Description                          "
    grillaCuentas.Rows = 1
    
    CargarAsignadas
End Sub

Private Sub CargarAsignadas()
  Dim anAccount As CAccount
  Dim indice As Integer
  
  indice = 1
  For Each anAccount In Accounts
    grillaCuentas.AddItem Trim(Str(anAccount.AccountId)) + Chr(9) + anAccount.Description
    grillaCuentas.RowData(indice) = anAccount.AccountId
    indice = indice + 1
  Next
End Sub

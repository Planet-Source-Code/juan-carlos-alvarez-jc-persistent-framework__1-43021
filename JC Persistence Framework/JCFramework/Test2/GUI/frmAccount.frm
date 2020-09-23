VERSION 5.00
Begin VB.Form frmAccount 
   Caption         =   "Form1"
   ClientHeight    =   1680
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   3450
   LinkTopic       =   "Form1"
   ScaleHeight     =   1680
   ScaleWidth      =   3450
   StartUpPosition =   2  'CenterScreen
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
      Left            =   2100
      TabIndex        =   4
      Top             =   1200
      Width           =   1200
   End
   Begin VB.TextBox txtDescription 
      Height          =   315
      Left            =   1260
      TabIndex        =   3
      Top             =   720
      Width           =   1995
   End
   Begin VB.TextBox txtAccountId 
      Height          =   315
      Left            =   1260
      TabIndex        =   1
      Top             =   300
      Width           =   1995
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Description:"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   780
      Width           =   840
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Account id:"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   810
   End
End
Attribute VB_Name = "frmAccount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_Account As CAccount

'Account
Public Property Get Account() As CAccount
  Set Account = m_Account
End Property
Public Property Set Account(ByVal c As CAccount)
  Set m_Account = c
End Property

Private Sub cmdOk_Click()
  If Val(txtAccountId.Text) <= 0 Then
    MsgBox "Account id must be greater than 0"
    Exit Sub
  End If
  Set m_Account = New CAccount
  'This project is to show how the framework works and not
  'to control if you insert a number or a text
  'Please fill a number in the text field accountId
  m_Account.AccountId = txtAccountId.Text
  m_Account.Description = txtDescription.Text
  m_Account.retrieve
  If m_Account.Persistent Then
    MsgBox "The accountId=" & Str(m_Account.AccountId) & " is already in use. Choose another one."
    Set m_Account = Nothing
  Else
    If frmAccounts.User.UserId = 0 Then
      MsgBox "Changes don´t take effect until you click on Add button"
    Else
      MsgBox "Changes don´t take effect until you click on Modify button"
    End If
    Me.Hide
  End If
End Sub

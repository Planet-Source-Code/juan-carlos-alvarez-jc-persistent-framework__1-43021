VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Begin VB.Form frmTest2 
   Caption         =   "Test2 JC Persistent Framework"
   ClientHeight    =   7860
   ClientLeft      =   2835
   ClientTop       =   555
   ClientWidth     =   6675
   LinkTopic       =   "Form1"
   ScaleHeight     =   7860
   ScaleWidth      =   6675
   Begin VB.CommandButton cmdAccounts 
      Caption         =   "Accounts"
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
      Left            =   5340
      TabIndex        =   25
      Top             =   3480
      Width           =   1200
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "New"
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
      Left            =   5340
      TabIndex        =   9
      Top             =   360
      Width           =   1200
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "Delete"
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
      Left            =   5340
      TabIndex        =   12
      Top             =   1800
      Width           =   1200
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "Modify"
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
      Left            =   5340
      TabIndex        =   11
      Top             =   1320
      Width           =   1200
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "Add"
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
      Left            =   5340
      TabIndex        =   10
      Top             =   840
      Width           =   1200
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Exit"
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
      Left            =   5340
      TabIndex        =   13
      Top             =   3960
      Width           =   1200
   End
   Begin VB.Frame frameCliente 
      Caption         =   "User"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4635
      Left            =   180
      TabIndex        =   14
      Top             =   180
      Width           =   4935
      Begin VB.ComboBox cmbCountries 
         Height          =   315
         Left            =   1380
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   4140
         Width           =   2355
      End
      Begin VB.TextBox txtDir 
         Height          =   330
         Left            =   1380
         TabIndex        =   3
         Top             =   1560
         Width           =   3360
      End
      Begin VB.TextBox txtUserId 
         Height          =   330
         Left            =   1380
         TabIndex        =   0
         Top             =   300
         Width           =   1800
      End
      Begin VB.TextBox txtTel 
         Height          =   330
         Left            =   1380
         TabIndex        =   4
         Top             =   1980
         Width           =   1620
      End
      Begin VB.TextBox txtApe 
         Height          =   330
         Left            =   1380
         TabIndex        =   2
         Top             =   1140
         Width           =   3360
      End
      Begin VB.TextBox txtNom 
         Height          =   330
         Left            =   1380
         TabIndex        =   1
         Top             =   720
         Width           =   3360
      End
      Begin VB.TextBox txtUsu 
         Height          =   330
         Left            =   1380
         TabIndex        =   5
         Top             =   2400
         Width           =   1620
      End
      Begin VB.TextBox txtClave 
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   1380
         MaxLength       =   10
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   2820
         Width           =   1620
      End
      Begin VB.TextBox txtEMail 
         Height          =   330
         Left            =   1380
         TabIndex        =   7
         Top             =   3240
         Width           =   3360
      End
      Begin MSComCtl2.DTPicker dtpFechaNac 
         Height          =   360
         Left            =   1380
         TabIndex        =   8
         Top             =   3660
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   635
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CustomFormat    =   "XXXXXXXXXX"
         Format          =   23658497
         CurrentDate     =   36673
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Country:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   26
         Top             =   4140
         Width           =   720
      End
      Begin VB.Label lblFechaNac 
         AutoSize        =   -1  'True
         Caption         =   "Date of birth:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   23
         Top             =   3720
         Width           =   1110
      End
      Begin VB.Label lblUserId 
         AutoSize        =   -1  'True
         Caption         =   "User id:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   22
         Top             =   360
         Width           =   690
      End
      Begin VB.Label lblTelefono 
         AutoSize        =   -1  'True
         Caption         =   "Telephone:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   21
         Top             =   2040
         Width           =   1035
      End
      Begin VB.Label lblDireccion 
         AutoSize        =   -1  'True
         Caption         =   "Address:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   20
         Top             =   1620
         Width           =   810
      End
      Begin VB.Label lblApellido 
         AutoSize        =   -1  'True
         Caption         =   "Last name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   19
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label lblNombre 
         AutoSize        =   -1  'True
         Caption         =   "First name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   18
         Top             =   780
         Width           =   975
      End
      Begin VB.Label lblNomUsu 
         AutoSize        =   -1  'True
         Caption         =   "Username:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   17
         Top             =   2460
         Width           =   990
      End
      Begin VB.Label lblClave 
         AutoSize        =   -1  'True
         Caption         =   "Password:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   16
         Top             =   2880
         Width           =   945
      End
      Begin VB.Label lblEmail 
         AutoSize        =   -1  'True
         Caption         =   "E-mail:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   15
         Top             =   3300
         Width           =   615
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grillaClientes 
      Height          =   2715
      Left            =   180
      TabIndex        =   24
      ToolTipText     =   "Seleccione un cliente para modificar, eliminar o consultar"
      Top             =   4980
      Width           =   6315
      _ExtentX        =   11139
      _ExtentY        =   4789
      _Version        =   393216
      Rows            =   1
      Cols            =   3
      FixedCols       =   0
      AllowBigSelection=   -1  'True
      FocusRect       =   0
      HighLight       =   2
      FillStyle       =   1
      ScrollBars      =   2
      SelectionMode   =   1
      FormatString    =   "<Last name                             |<First name                          |<Telephone        "
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
Attribute VB_Name = "frmTest2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Corresponde al cliente seleccionado en la grilla
'To manage the selected user in the grid
Private m_cliSel As CUser

Private Sub cmdAccounts_Click()
  Set frmAccounts.User = m_cliSel
  'if a new user, remove his item from collection
  'for the moment this is the solution
  If m_cliSel.UserId = 0 Then
    If m_cliSel.Accounts.Count > 0 Then
      m_cliSel.Accounts.Remove (1)
    End If
  End If
  Set frmAccounts.Accounts = m_cliSel.Accounts
  frmAccounts.Show vbModal
  Set m_cliSel.Accounts = frmAccounts.Accounts
  Unload frmAccounts
End Sub

Private Sub Form_Load()
  'Importante: En este proyecto "Test2" no hay referencia
  'alguna a ADO (a la base de datos).
  'Si se quiere manejar otra base solo es cuestión de modificar
  'en el archivo ini el path del archivo xml que se utilizará,
  'no hay que tocar absolutamente ni una línea de código,
  'ni en la parte gráfica de esta aplicación, ni en la dll de business,
  'ni en la dll del JCFramework.
  'Esto queda probado en esta nueva versión con el manejo de
  'MsAccess y MySQL (esto es nuevo en esta version).
  
  'Important: In this project "Test2" there aren´t any references
  'to ADO (to the database).
  'If you want to work with another database you only need
  'to change in the ini file the xml path that you are going to use,
  'but anything in the GUI part, nor in the business dll, nor in the
  'jcframework dll.
  
  
  'Para iniciar la carga del XML automáticamente
  'y no cuando se intenta realizar la primera operación
  
  Dim initializer As New CJCFrameworkInit
  initializer.init
  
  
  'Común al proyecto test2
  LimpiarFecha dtpFechaNac
  Set m_cliSel = New CUser
  InicioBotones
  CargarPaises
  ActualizarGrilla
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set m_cliSel = Nothing
End Sub

Private Sub cmdSalir_Click()
  Unload Me
End Sub

Private Sub CargarGrilla()
  'Traigo los datos de la base
  '(esto lo hace el JCFramework)
  Dim colUsers As Collection
  
  Dim manager As New CJCCollectionManager
  Dim aUserTemp As New CUser
  Set colUsers = manager.getCollectionFor(aUserTemp)
  Set aUserTemp = Nothing
  Set manager = Nothing
  
  Dim cliente As CUser
  Dim indice As Long
  indice = 1
  For Each cliente In colUsers
    ' Cargo los datos del cliente a la grilla
    grillaClientes.AddItem cliente.Lastname + Chr(9) + cliente.Firstname + Chr(9) + cliente.Telephone
    grillaClientes.RowData(indice) = cliente.UserId
    indice = indice + 1
  Next
  Set colUsers = Nothing
End Sub

Private Sub CargarPaises()
  Dim colCountries As Collection
  
  Dim manager As New CJCCollectionManager
  Dim aCountryTemp As New CCountry
  Set colCountries = manager.getCollectionFor(aCountryTemp)
  Set aCountryTemp = Nothing
  Set manager = Nothing
  
  Dim aCountry As CCountry
  Dim indice As Long
  indice = 0
  For Each aCountry In colCountries
    ' Cargo los datos del cliente al combo
    cmbCountries.AddItem aCountry.Description
    cmbCountries.ItemData(indice) = aCountry.CountryId
    indice = indice + 1
  Next
  Set colCountries = Nothing
End Sub

Private Sub ActualizarGrilla()
    grillaClientes.Clear
    'Esto se debe leer de un archivo ini
    grillaClientes.FormatString = "<Last name                             |<First name                          |<Telephone        "
    grillaClientes.Rows = 1
    
    CargarGrilla
End Sub

Private Sub LimpiarCampos()
    txtUserId.Text = ""
    txtNom.Text = ""
    txtApe.Text = ""
    txtDir.Text = ""
    txtTel.Text = ""
    txtUsu.Text = ""
    txtClave.Text = ""
    txtEMail.Text = ""
    LimpiarFecha dtpFechaNac
End Sub

Private Sub ObjetoAInterface(ByVal cliente As CUser)
    txtUserId.Text = cliente.UserId
    txtNom.Text = cliente.Firstname
    txtApe.Text = cliente.Lastname
    txtDir.Text = cliente.Address
    txtTel.Text = cliente.Telephone
    txtUsu.Text = cliente.Username
    txtClave.Text = cliente.Password
    txtEMail.Text = cliente.EMail
    If Not cliente.DateOfBirth = #12:00:00 AM# Then
        PrepararFecha dtpFechaNac
        dtpFechaNac.Value = cliente.DateOfBirth
    End If
    cmbCountries.Text = cliente.Country.Description
End Sub

Private Sub InterfaceAObjeto(ByRef cliente As CUser)
    cliente.UserId = txtUserId.Text
    cliente.Firstname = txtNom.Text
    cliente.Lastname = txtApe.Text
    cliente.Address = txtDir.Text
    cliente.Telephone = txtTel.Text
    cliente.Username = txtUsu.Text
    cliente.Password = txtClave.Text
    cliente.EMail = txtEMail.Text
    cliente.DateOfBirth = dtpFechaNac.Value
    cliente.DateOfAdded = Now
End Sub

Private Sub InicioBotones()
    cmdNuevo.Enabled = False
    cmdAgregar.Enabled = True
    cmdModificar.Enabled = False
    cmdEliminar.Enabled = False
    Me.txtUserId.Enabled = True
End Sub

Private Sub BotonesClickGrilla()
    cmdNuevo.Enabled = True
    cmdAgregar.Enabled = False
    cmdModificar.Enabled = True
    cmdEliminar.Enabled = True
    Me.txtUserId.Enabled = False
End Sub

Private Sub cmdNuevo_Click()
    LimpiarCampos
    'Manejo de botones
    InicioBotones
    'Manejo de otros controles
    txtUserId.SetFocus
End Sub

Private Sub cmdAgregar_Click()
  If Not ControlesOk Then
    Exit Sub
  End If
  
'  ' Prueba relacion multiple
'  Dim colCuentas As Collection
'  Set colCuentas = New Collection
'
'  Dim retrieveCriteria As CRetrieveCriteria
'  Dim cursor As CCursor
'  Dim cuenta As New CAccount
'  Set retrieveCriteria = New CRetrieveCriteria
'  Dim colParams As New Collection
'  Set cursor = retrieveCriteria.perform(cuenta, colParams)
'  While cursor.hasElements
'    cursor.loadObject cuenta
'    colCuentas.Add cuenta
'    Set cuenta = Nothing
'    Set cuenta = New CAccount
'    cursor.nextCursor
'  Wend
'  Set retrieveCriteria = Nothing
'  Set cuenta = Nothing
'  ' Fin Prueba relacion multiple
  
'  'Prueba relacion simple
'  Dim aCountry As CCountry
'  Set aCountry = New CCountry
'  aCountry.CountryId = 1
'  aCountry.retrieve
'  'Fin Prueba relacion simple
  
  Dim cliente As CUser
  Set cliente = New CUser
  InterfaceAObjeto cliente
  'Para controlar que no exista
  'Se intenta traer el objeto
  cliente.retrieve
  If cliente.Persistent Then
    MsgBox "The user with id=" & Trim(Str(cliente.UserId)) & " already exist."
    Exit Sub
  End If
  
  'Set the selected country for the user
  Dim aCountry As CCountry, idCountry As Long
  Set aCountry = New CCountry
  idCountry = cmbCountries.ItemData(cmbCountries.ListIndex)
  aCountry.CountryId = idCountry
  aCountry.retrieve
  
  cliente.CountryId = aCountry.CountryId
  Set cliente.Country = aCountry
  
  'Set in the selected accounts for the user, the userid
  'for the moment this is the solution
  Dim anAccount As CAccount
  For Each anAccount In m_cliSel.Accounts
    anAccount.UserId = cliente.UserId
  Next
  
  If m_cliSel.Accounts.Count > 0 Then
    If m_cliSel.Accounts.Item(1).AccountId = 0 Then
      m_cliSel.Accounts.Remove (1)
    End If
  End If
  
  Set cliente.Accounts = m_cliSel.Accounts
  cliente.save
  
  Set cliente = Nothing
  Set aCountry = Nothing
  ActualizarGrilla
  cmdNuevo_Click
End Sub

Private Sub cmdModificar_Click()
  If Not ControlesOk Then Exit Sub
  'Guardo el cliente seleccionado
  InterfaceAObjeto m_cliSel
  
  'Set the selected country for the user
  Dim aCountry As CCountry, idCountry As Long
  Set aCountry = New CCountry
  idCountry = cmbCountries.ItemData(cmbCountries.ListIndex)
  aCountry.CountryId = idCountry
  aCountry.retrieve
  m_cliSel.CountryId = aCountry.CountryId
  Set m_cliSel.Country = aCountry
  
  m_cliSel.save
  Set aCountry = Nothing
  'Actualizo cambios (por ejemplo si cambio el Firstname)
  ActualizarGrilla
  cmdNuevo_Click
End Sub

Private Sub cmdEliminar_Click()
  'Elimino el cliente seleccionado
  m_cliSel.Delete
  ActualizarGrilla
  cmdNuevo_Click
End Sub

Private Sub grillaClientes_Click()
  If grillaClientes.Rows <= 1 Then Exit Sub
  'Busco la clave del cliente, o sea su UserId
  Dim indice As Integer, Ced As Long
  indice = grillaClientes.Row
  Ced = grillaClientes.RowData(indice)

  'El tener al cliente seleccionado facilita
  'la realizacion de metodos como eliminar y modificar
  Set m_cliSel = Nothing
  Set m_cliSel = New CUser
  m_cliSel.UserId = Ced
  m_cliSel.retrieve
  ObjetoAInterface m_cliSel
  'Habilito botones necesarios
  BotonesClickGrilla
End Sub

Private Sub grillaClientes_SelChange()
  'grillaClientes_Click
End Sub

Private Sub dtpFechaNac_GotFocus()
    PrepararFecha dtpFechaNac
End Sub

Private Sub dtpFechaNac_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If cmdAgregar.Enabled Then
            cmdAgregar.SetFocus
        Else
            cmdModificar.SetFocus
        End If
    ElseIf KeyCode = vbKeyDelete Then
        'Se coloca el foco en el siguiente control
        If cmdAgregar.Enabled Then
            cmdAgregar.SetFocus
        Else
            cmdModificar.SetFocus
        End If
        LimpiarFecha dtpFechaNac
    End If
End Sub

Private Function ControlesOk() As Boolean
    'Validaciones de las cajas de texto
    If (txtUserId = "") Then
        MsgBox "Fill in userId"
        txtUserId.SetFocus
        txtUserId.SelStart = 0
        txtUserId.SelLength = Len(txtUserId)
        ControlesOk = False
        Exit Function
    End If
    If (txtNom = "") Then
        MsgBox "Fill in Firstname"
        txtNom.SetFocus
        txtNom.SelStart = 0
        txtNom.SelLength = Len(txtNom)
        ControlesOk = False
        Exit Function
    End If
    If (txtApe = "") Then
        MsgBox "Fill in Lastname"
        txtApe.SetFocus
        txtApe.SelStart = 0
        txtApe.SelLength = Len(txtApe)
        ControlesOk = False
        Exit Function
    End If
    If (txtUsu = "") Then
        MsgBox "Fill in Username"
        txtUsu.SetFocus
        txtUsu.SelStart = 0
        txtUsu.SelLength = Len(txtUsu)
        ControlesOk = False
        Exit Function
    End If
    If (txtClave = "") Then
        MsgBox "Fill in Password"
        txtClave.SetFocus
        txtClave.SelStart = 0
        txtClave.SelLength = Len(txtClave)
        ControlesOk = False
        Exit Function
    End If
    If (cmbCountries.ListIndex < 0) Then
        MsgBox "Choose a country"
        cmbCountries.SetFocus
        ControlesOk = False
        Exit Function
    End If
    ControlesOk = True
End Function

Private Sub LimpiarFecha(dtPick As DTPicker)
    dtPick.Format = dtpCustom
    dtPick.CustomFormat = "XXXXXXXXXX"
    dtPick.Value = #12:00:00 AM#
End Sub

Private Sub PrepararFecha(dtPick As DTPicker)
    dtPick.Format = dtpShortDate
    dtPick.Value = #1/1/2001#
End Sub

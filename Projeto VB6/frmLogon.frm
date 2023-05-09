VERSION 5.00
Begin VB.Form frmLogon 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ProCam Caprind 2002 - [Logon]"
   ClientHeight    =   2280
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8124
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.4
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmLogon.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2280
   ScaleWidth      =   8124
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   2175
      Left            =   120
      TabIndex        =   6
      Top             =   0
      Width           =   7815
      Begin VB.TextBox txtUsuario 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Height          =   285
         Left            =   2400
         TabIndex        =   0
         Top             =   1080
         Width           =   2775
      End
      Begin VB.TextBox txtSenha 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   2400
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   1440
         Width           =   2775
      End
      Begin VB.CommandButton cmdAcessar 
         Caption         =   "Acessar"
         Height          =   660
         Left            =   5280
         TabIndex        =   2
         Top             =   1060
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Para efetuar seu logon, você deve se identificar. Preencha os campos abaixo e clique no botão <acessar>:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   615
         Left            =   1560
         TabIndex        =   9
         Top             =   480
         Width           =   4935
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Usuário:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   8
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Senha:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.4
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   7
         Top             =   1440
         Width           =   735
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         Height          =   1815
         Left            =   960
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   6135
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   2295
      Left            =   -360
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   8175
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         Height          =   1620
         IntegralHeight  =   0   'False
         Left            =   600
         TabIndex        =   4
         Top             =   120
         Width           =   3375
      End
      Begin VB.Label Label5 
         Caption         =   "Inicializando..."
         Height          =   255
         Left            =   600
         TabIndex        =   5
         Top             =   0
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmLogon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAcessar_Click()
    List1.Clear
    Frame1.Visible = True
    Frame1.Refresh
    
    List1.AddItem "Abrindo base de dados...", 0
    List1.ListIndex = 0
    
    If AbreBD = True Then
        List1.AddItem "Ok!", 0
        List1.ListIndex = 0
        
        List1.AddItem "Validando usuário...", 0
        List1.ListIndex = 0
        
        If procValidaSenha(txtUsuario.Text, txtSenha.Text) = True Then
            frmMDI.Show
'            procAtualizaBancoTabelas
            
            Unload Me
        Else
            List1.AddItem "*** Erro ***", 0
            List1.ListIndex = 0
            
            MsgBox "Nome de usuário ou senha inválidos!", vbCritical
        
            Frame1.Visible = False
            txtUsuario.SetFocus
        End If
    Else
        List1.AddItem "*** Erro ***", 0
        List1.ListIndex = 0
        
        Frame1.Visible = False
        txtUsuario.SetFocus
    End If
End Sub

'Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'    Dim X As String
'
'    If KeyCode = vbKeyF8 Then
'        X = InputBox("Digite na linha abaixo local onde se encontra a base de dados, exemplo: """ & App.Path & "\Caprind.MDB"".", , LocalBD)
'        If X <> "" Then
'            LocalBD = X
'            SaveSetting "ProCam", "Caprind", "LocalBD", X
'            MsgBox "Endereço da base de dados alterado com sucesso!", vbExclamation
'        End If
'    End If
'End Sub

Private Sub Form_Resize()
txtUsuario.SetFocus
End Sub

Private Sub txtSenha_Change()
    Call procDefaultAcesso
End Sub

'Private Sub txtUsuario_Change()
'    Call procDefaultAcesso
'End Sub


Function procDefaultAcesso()
    If txtSenha.Text <> "" And txtUsuario.Text <> "" Then
        cmdAcessar.Default = True
    Else
        cmdAcessar.Default = False
    End If
End Function

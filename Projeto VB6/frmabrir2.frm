VERSION 5.00
Begin VB.Form frmabrir2 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Caprind ( Gerprod  - Coletor de dados no chão de fábrica )"
   ClientHeight    =   6960
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9315
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmabrir2.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6960
   ScaleWidth      =   9315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Height          =   2325
      Left            =   30
      TabIndex        =   0
      Top             =   810
      Width           =   9270
      Begin VB.TextBox txtdescricao 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         IMEMode         =   3  'DISABLE
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   6
         ToolTipText     =   "Descrição do item em fabricação."
         Top             =   1500
         Width           =   8775
      End
      Begin VB.TextBox txtdesenho 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         IMEMode         =   3  'DISABLE
         Left            =   3450
         MaxLength       =   20
         TabIndex        =   5
         ToolTipText     =   "Senha do usuário."
         Top             =   645
         Width           =   1905
      End
      Begin VB.TextBox txtquant 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         IMEMode         =   3  'DISABLE
         Left            =   5370
         MaxLength       =   20
         TabIndex        =   4
         ToolTipText     =   "Senha do usuário."
         Top             =   645
         Width           =   1785
      End
      Begin VB.TextBox txtprazo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         IMEMode         =   3  'DISABLE
         Left            =   7170
         MaxLength       =   20
         TabIndex        =   3
         ToolTipText     =   "Senha do usuário."
         Top             =   645
         Width           =   1845
      End
      Begin VB.TextBox TXTOF 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         IMEMode         =   3  'DISABLE
         Left            =   2100
         MaxLength       =   20
         TabIndex        =   2
         ToolTipText     =   "Nº Ordem"
         Top             =   645
         Width           =   1335
      End
      Begin VB.TextBox cmbos 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         IMEMode         =   3  'DISABLE
         Left            =   240
         MaxLength       =   20
         TabIndex        =   1
         ToolTipText     =   "Numero da ordem de serviço"
         Top             =   645
         Width           =   1845
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N° O.F"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   2340
         TabIndex        =   12
         Top             =   270
         Width           =   945
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição do item"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   960
         TabIndex        =   11
         Top             =   1155
         Width           =   2730
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   3840
         TabIndex        =   10
         Top             =   270
         Width           =   1080
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quant."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   5730
         TabIndex        =   9
         Top             =   270
         Width           =   1005
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Prazo final"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   7290
         TabIndex        =   8
         Top             =   270
         Width           =   1590
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N° O.S"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   720
         TabIndex        =   7
         ToolTipText     =   "N° da ordem de serviço"
         Top             =   270
         Width           =   990
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Frame2"
      Height          =   2895
      Left            =   30
      TabIndex        =   13
      Top             =   2880
      Width           =   9270
      Begin VB.TextBox txtdescmaq 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   440
         IMEMode         =   3  'DISABLE
         Left            =   3690
         MaxLength       =   20
         TabIndex        =   17
         ToolTipText     =   "Senha do usuário."
         Top             =   765
         Width           =   5295
      End
      Begin VB.TextBox txtmaquina 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   440
         IMEMode         =   3  'DISABLE
         Left            =   1770
         MaxLength       =   20
         TabIndex        =   16
         ToolTipText     =   "código da máquina"
         Top             =   765
         Width           =   1905
      End
      Begin VB.TextBox txtfase 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   440
         IMEMode         =   3  'DISABLE
         Left            =   210
         MaxLength       =   20
         TabIndex        =   15
         ToolTipText     =   "Senha do usuário."
         Top             =   765
         Width           =   1545
      End
      Begin VB.TextBox txtis 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         IMEMode         =   3  'DISABLE
         Left            =   210
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   14
         ToolTipText     =   "Senha do usuário."
         Top             =   1725
         Width           =   8775
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fase"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   480
         Left            =   630
         TabIndex        =   21
         Top             =   405
         Width           =   735
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Posto Trab."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   1860
         TabIndex        =   20
         Top             =   405
         Width           =   1755
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição do posto de trabalho"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   3960
         TabIndex        =   19
         Top             =   405
         Width           =   4755
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Instrução do serviço"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   690
         TabIndex        =   18
         Top             =   1365
         Width           =   3120
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Frame3"
      Height          =   1335
      Left            =   30
      TabIndex        =   22
      Top             =   5610
      Width           =   9270
      Begin VB.TextBox txtusuario 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         IMEMode         =   3  'DISABLE
         Left            =   2880
         MaxLength       =   20
         TabIndex        =   24
         ToolTipText     =   "Senha do usuário."
         Top             =   600
         Width           =   6105
      End
      Begin VB.TextBox txtSenha 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         IMEMode         =   3  'DISABLE
         Left            =   210
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   23
         ToolTipText     =   "Senha do usuário."
         Top             =   600
         Width           =   2655
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Operador"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   5340
         TabIndex        =   26
         Top             =   270
         Width           =   1425
      End
      Begin VB.Label lblsenha 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Senha"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   360
         Left            =   1050
         TabIndex        =   25
         Top             =   270
         Width           =   960
      End
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter - Aceitar Dados   /   Esc - Voltar a tela anterior"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   435
      Left            =   120
      TabIndex        =   27
      Top             =   180
      Width           =   8820
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      Height          =   735
      Left            =   30
      Top             =   30
      Width           =   9240
   End
End
Attribute VB_Name = "frmabrir2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub procchecacampos()
On Error GoTo tratar_erro

If txtusuario.Text = "" Then
    MsgBox ("Digite sua senha para acesso."), vbInformation
    txtSenha.SetFocus
    Exit Sub
End If
If cmbos.Text = "" Then
    MsgBox ("Escolha uma O.s para acesso."), vbInformation
    cmbos.SetFocus
    Exit Sub
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub cmbos_Change()
On Error GoTo tratar_erro

If cmbos.Text <> "" Then
    Set TBMaquinas = BD.OpenRecordset("Select * from ordemservico where idproducao = " & cmbos.Text & " and pronto = 'NÃO';")
    If TBMaquinas.EOF = False Then
        txtof.Text = IIf(IsNull(TBMaquinas!OF) = False, TBMaquinas!OF, "")
        txtFase.Text = IIf(IsNull(TBMaquinas!fase) = False, TBMaquinas!fase, "")
        txtis.Text = IIf(IsNull(TBMaquinas!descfase) = False, TBMaquinas!descfase, "")
        txtdescricao.Text = IIf(IsNull(TBMaquinas("descricao")) = False, TBMaquinas!descricao, "")
        txtdesenho.Text = IIf(IsNull(TBMaquinas("desenho")) = False, TBMaquinas!desenho, "")
        txtprazo.Text = IIf(IsNull(TBMaquinas("prazofinal")) = False, TBMaquinas!PRAZOFINAL, "")
        txtquant.Text = IIf(IsNull(TBMaquinas("quantidade")) = False, TBMaquinas!Quantidade, "")
        txtMaquina.Text = IIf(IsNull(TBMaquinas!maquina) = False, TBMaquinas!maquina, "")
        txtSenha.Enabled = True
    Else
        txtSenha.Enabled = False
        txtof.Text = ""
        txtFase.Text = ""
        txtis.Text = ""
        txtdescricao.Text = ""
        txtdesenho.Text = ""
        txtprazo.Text = ""
        txtquant.Text = ""
        txtMaquina.Text = ""
        txtdescmaq.Text = ""
        cmbos.SetFocus

End If
        Else
        txtof.Text = ""
        txtFase.Text = ""
        txtis.Text = ""
        txtdescricao.Text = ""
        txtdesenho.Text = ""
        txtprazo.Text = ""
        txtquant.Text = ""
        txtMaquina.Text = ""
        txtdescmaq.Text = ""

TBMaquinas.Close
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro


If txtSenha.Text <> "" And txtusuario.Text <> "" And cmbos.Text <> "" Then
    procchecacampos
    If KeyCode = 13 Then
        frmProducao.procabrir
        If OrdemExiste = True Then
            Unload Me
        End If
    End If
End If
If KeyCode = vbKeyEscape Then
    Unload Me
    frmfundo.Show
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo tratar_erro
    
Call procKeyascii(KeyAscii)

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro
    
txtSenha.Enabled = False

If Leitor = True Then
lblsenha.Caption = "Código"
cmbos.SetFocus
Else
lblsenha.Caption = "Senha"
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub txtmaquina_Change()
On Error GoTo tratar_erro

Set TBabrir = BD.OpenRecordset("Select * from cadmaquinas where maquina = '" & txtMaquina.Text & "';")
If TBabrir.EOF = False Then
    txtdescmaq = IIf(IsNull(TBabrir!descricao) = False, TBabrir!descricao, "")
End If
TBabrir.Close

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub txtSenha_Change()
On Error GoTo tratar_erro


If txtSenha.Text <> "" Then
    Set tbusuario = BD.OpenRecordset("Select * from usuarios where senha = '" & txtSenha.Text & "'")
    If tbusuario.EOF = False Then
        txtusuario.Text = tbusuario("usuario")
    Else
        txtusuario.Text = ""
        Exit Sub
    End If
End If


Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

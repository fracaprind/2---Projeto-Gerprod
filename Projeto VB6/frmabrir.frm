VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmabrir 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Caprind ( Gerprod coletor de dados )"
   ClientHeight    =   5595
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   9375
   ClipControls    =   0   'False
   Icon            =   "frmabrir.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5595
   ScaleWidth      =   9375
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Dados da Ordem"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3345
      Left            =   0
      TabIndex        =   6
      Top             =   780
      Width           =   9375
      Begin VB.TextBox txtitem 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   475
         IMEMode         =   3  'DISABLE
         Left            =   2100
         MaxLength       =   20
         TabIndex        =   23
         ToolTipText     =   "Numero da ordem"
         Top             =   900
         Width           =   3015
      End
      Begin VB.TextBox txtordem 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   475
         IMEMode         =   3  'DISABLE
         Left            =   180
         MaxLength       =   20
         TabIndex        =   22
         ToolTipText     =   "Numero da ordem"
         Top             =   900
         Width           =   1905
      End
      Begin VB.TextBox txtdescricao 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   150
         TabIndex        =   2
         ToolTipText     =   "Descrição do item"
         Top             =   1830
         Width           =   8955
      End
      Begin VB.TextBox txtusuario 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         IMEMode         =   3  'DISABLE
         Left            =   3900
         MaxLength       =   20
         TabIndex        =   5
         ToolTipText     =   "Senha do usuário."
         Top             =   2730
         Width           =   5205
      End
      Begin VB.TextBox txtSenha 
         Alignment       =   2  'Center
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         IMEMode         =   3  'DISABLE
         Left            =   1800
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   4
         ToolTipText     =   "Senha do usuário."
         Top             =   2730
         Width           =   2085
      End
      Begin VB.ComboBox cmbPT 
         BackColor       =   &H80000014&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   150
         MouseIcon       =   "frmabrir.frx":030A
         MousePointer    =   99  'Custom
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         ToolTipText     =   "Posto de trabalho"
         Top             =   2730
         Width           =   1635
      End
      Begin VB.TextBox txtquant 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   475
         IMEMode         =   3  'DISABLE
         Left            =   5130
         MaxLength       =   20
         TabIndex        =   0
         ToolTipText     =   "Quantidade de itens a serem fabricados"
         Top             =   900
         Width           =   1905
      End
      Begin MSMask.MaskEdBox mskprazofinal 
         Height          =   480
         Left            =   7050
         TabIndex        =   1
         ToolTipText     =   "Digite a data de entrega do item"
         Top             =   900
         Width           =   2040
         _ExtentX        =   3598
         _ExtentY        =   847
         _Version        =   393216
         Appearance      =   0
         BackColor       =   14737632
         AllowPrompt     =   -1  'True
         AutoTab         =   -1  'True
         Enabled         =   0   'False
         MaxLength       =   8
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yy"
         Mask            =   "##/##/##"
         PromptChar      =   "_"
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição do item"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   3240
         TabIndex        =   16
         Top             =   1470
         Width           =   2235
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Operador"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   5910
         TabIndex        =   15
         Top             =   2370
         Width           =   1185
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Senha"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   2430
         TabIndex        =   14
         Top             =   2370
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Posto"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   570
         TabIndex        =   13
         Top             =   2370
         Width           =   660
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Prazo final"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   7350
         TabIndex        =   12
         Top             =   540
         Width           =   1290
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   3330
         TabIndex        =   11
         Top             =   3420
         Width           =   1005
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quantidade"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   5400
         TabIndex        =   10
         Top             =   540
         Width           =   1455
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código item"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   2850
         TabIndex        =   9
         Top             =   540
         Width           =   1485
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Ordem"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   480
         TabIndex        =   7
         Top             =   540
         Width           =   1275
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Height          =   1485
      Left            =   0
      TabIndex        =   8
      Top             =   4110
      Width           =   9375
      Begin VB.TextBox txtos 
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
         Left            =   1290
         MaxLength       =   20
         TabIndex        =   25
         ToolTipText     =   "Senha do usuário."
         Top             =   675
         Width           =   1305
      End
      Begin VB.TextBox txtcliente 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         IMEMode         =   3  'DISABLE
         Left            =   2610
         TabIndex        =   20
         ToolTipText     =   "Código do item"
         Top             =   675
         Width           =   6525
      End
      Begin VB.TextBox txtfase 
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
         Left            =   210
         MaxLength       =   20
         TabIndex        =   17
         ToolTipText     =   "Senha do usuário."
         Top             =   675
         Width           =   1065
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente"
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
         Height          =   390
         Left            =   5550
         TabIndex        =   24
         Top             =   330
         Width           =   1035
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "O.s"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   1620
         TabIndex        =   19
         Top             =   330
         Width           =   465
      End
      Begin VB.Label Label14 
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
         Height          =   360
         Left            =   330
         TabIndex        =   18
         Top             =   330
         Width           =   735
      End
   End
   Begin VB.Label Label15 
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
      Left            =   330
      TabIndex        =   21
      Top             =   150
      Width           =   8820
   End
   Begin VB.Shape Shape2 
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      Height          =   735
      Left            =   30
      Top             =   30
      Width           =   9330
   End
End
Attribute VB_Name = "frmabrir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbmaquina_Change()
On Error GoTo tratar_erro

If cmbmaquina.Text <> "" Then
    Set TBFases = BD.OpenRecordset("Select fase from ordemservico where maquina = '" & cmbmaquina.Text & "' and prioridade =1")
    If TBFases.EOF = False Then
        TBFases.MoveFirst
        txtordem.Text = TBFases("of")
        txtfase.Text = ""
        txtfase.Text = TBFases("fase")
        txtdescricao.Text = TBFases("descricao")
    End If
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub cmbmaquina_Click()
On Error GoTo tratar_erro
Dim condicao As String

condicao = "NÃO"
If cmbmaquina.Text <> "" Then
    Set TBFases = BD.OpenRecordset("Select * from ordemservico where maquina = '" & cmbmaquina.Text & "' and prioridade =1 and pronto = '" & condicao & "'")
    If TBFases.EOF = False Then
        TBFases.MoveFirst
        txtordem.Text = TBFases("of")
        txtfase.Text = ""
        txtfase.Text = TBFases("fase")
        txtdescricao.Text = TBFases("descricao")
    End If
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub cmbmaquina_Scroll()
On Error GoTo tratar_erro

If cmbmaquina.Text <> "" Then
    Set TBFases = BD.OpenRecordset("Select * from ordemservico where maquina = '" & cmbmaquina.Text & "' and prioridade =1")
    If TBFases.EOF = False Then
        TBFases.MoveFirst
        txtordem.Text = TBFases("of")
        txtfase.Text = ""
        txtfase.Text = TBFases("fase")
    End If
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub txtordem_Change()
On Error GoTo tratar_erro

If IsNumeric(txtordem.Text) = False Then
Exit Sub
End If

Set TBMaquinas = BD.OpenRecordset("Select * from producao where of = " & txtordem.Text & " order by LISTA")
If TBMaquinas.EOF = False Then
    txtitem.Enabled = False
    txtitem.Text = TBMaquinas!desenho
    txtquant.Enabled = False
    txtquant.Text = TBMaquinas!quant
    mskprazofinal.Enabled = False
    mskprazofinal.Text = Format(TBMaquinas!prazoentrega, "dd/mm/yy")
    txtcliente.Enabled = False
    txtcliente.Text = TBMaquinas!cliente
    txtdescricao.Enabled = False
    txtdescricao.Text = TBMaquinas!nomelista
    TBMaquinas.MoveFirst
Else
    txtitem.Text = ""
    txtquant.Enabled = True
    txtquant.Text = ""
    mskprazofinal.Enabled = True
    mskprazofinal.Text = "__/__/__"
    txtcliente.Enabled = True
    txtcliente.Text = ""
    txtdescricao.Enabled = True
    txtdescricao.Text = ""
End If
TBMaquinas.Close

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub cmbPT_Click()
On Error GoTo tratar_erro

If txtordem.Text <> "" Then
    Set TBProducao = BD.OpenRecordset("Select * from ordemservico where of= " & txtordem.Text & " and maquina = '" & cmbPT & "' and pronto = 'NÃO' order by idproducao;")
    If TBProducao.EOF = False Then
        txtfase.Text = TBProducao!fase
        txtos.Text = TBProducao!IDProducao
        TBProducao.Close
        Exit Sub
    End If
    Set TBProducao = BD.OpenRecordset("Select * from ordemservico where of= " & txtordem.Text & " order by idproducao;")
    If TBProducao.EOF = False Then
        TBProducao.MoveLast
        txtfase = TBProducao!fase + 10
        txtos.Text = ""
        TBProducao.Close
        Exit Sub
    End If
End If
txtfase = 10
txtos.Text = ""
mskprep.Text = "00:00:00"
mskexecucao.Text = "00:00:00"

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub txtitem_Change()
On Error GoTo tratar_erro


Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Set tbusuario = BD.OpenRecordset("Select * from usuarios where senha = '" & txtSenha.Text & "'")
If tbusuario.EOF = False Then
    txtusuario.Text = tbusuario("usuario")
Else
    txtusuario.Text = ""
End If
    
If KeyCode = 13 Then
    If txtordem.Text = "" Then
        MsgBox ("Escolha um pedido interno."), vbCritical
        txtordem.SetFocus
        Exit Sub
    End If
    If txtitem.Text = "" Then
        MsgBox ("Escolha um item."), vbCritical
        txtitem.SetFocus
        Exit Sub
    End If
    If txtquant.Text = "" Then
        MsgBox ("Digite a quantidade."), vbCritical
        txtquant.SetFocus
        Exit Sub
    End If
    If mskprazofinal.Text = "__/__/__" Then
        MsgBox ("Digite um prazo de fabricação."), vbCritical
        mskprazofinal.SetFocus
        Exit Sub
    End If
    If txtdescricao.Text = "" Then
        MsgBox ("Digite a descrição do item."), vbCritical
        txtdescricao.SetFocus
        Exit Sub
    End If
    If txtusuario.Text = "" Then
        MsgBox ("Digite a sua senha."), vbCritical
        txtSenha.SetFocus
        Exit Sub
    End If
    If cmbPT.Text = "" Then
        MsgBox ("Escolha um posto de trabalho."), vbCritical
        cmbPT.SetFocus
        Exit Sub
    End If
    frmProducao.procabrirNovo
    Unload Me
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

Set TBMaquinas = BD.OpenRecordset("Select * from cadmaquinas order by maquina")
cmbPT.Clear
Do Until TBMaquinas.EOF
    cmbPT.AddItem TBMaquinas("maquina")
    TBMaquinas.MoveNext
Loop
TBMaquinas.Close

txtordem.SetFocus
frmabrir.Refresh

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub txtSenha_Change()
On Error GoTo tratar_erro

Set tbusuario = BD.OpenRecordset("Select * from usuarios where senha = '" & txtSenha.Text & "'")
If tbusuario.EOF = False Then
    txtusuario.Text = tbusuario("usuario")
Else
    txtusuario.Text = ""
    Exit Sub
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

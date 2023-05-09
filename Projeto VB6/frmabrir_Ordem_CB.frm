VERSION 5.00
Begin VB.Form frmabrir_Ordem_CB 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Gerprod  - Coletor de dados no chão de fábrica - Código de barras"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9315
   ClipControls    =   0   'False
   Icon            =   "frmabrir_Ordem_CB.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   9315
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Dados da ordem"
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
      Height          =   3345
      Left            =   30
      TabIndex        =   12
      Top             =   780
      Width           =   9240
      Begin VB.TextBox txtcliente 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   450
         IMEMode         =   3  'DISABLE
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Cliente."
         Top             =   2685
         Width           =   8745
      End
      Begin VB.TextBox mskprazofinal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   475
         IMEMode         =   3  'DISABLE
         Left            =   7050
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Prazo final."
         Top             =   750
         Width           =   1920
      End
      Begin VB.TextBox txtitem 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   475
         IMEMode         =   3  'DISABLE
         Left            =   2100
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "Código interno."
         Top             =   750
         Width           =   3015
      End
      Begin VB.TextBox txtordem 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   475
         IMEMode         =   3  'DISABLE
         Left            =   240
         MaxLength       =   20
         TabIndex        =   0
         ToolTipText     =   "Número da ordem."
         Top             =   750
         Width           =   1845
      End
      Begin VB.TextBox txtdescricao 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   675
         IMEMode         =   3  'DISABLE
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Descrição."
         Top             =   1620
         Width           =   8745
      End
      Begin VB.TextBox txtquant 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   475
         IMEMode         =   3  'DISABLE
         Left            =   5130
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade."
         Top             =   750
         Width           =   1905
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente"
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
         Left            =   4192
         TabIndex        =   23
         Top             =   2340
         Width           =   840
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição"
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
         Left            =   4012
         TabIndex        =   19
         Top             =   1260
         Width           =   1200
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
         Left            =   7365
         TabIndex        =   18
         Top             =   390
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
         TabIndex        =   17
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
         Left            =   5355
         TabIndex        =   16
         Top             =   390
         Width           =   1455
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código interno"
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
         Left            =   2692
         TabIndex        =   15
         Top             =   390
         Width           =   1830
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nº ordem"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   382
         TabIndex        =   13
         Top             =   390
         Width           =   1500
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Height          =   2145
      Left            =   30
      TabIndex        =   14
      Top             =   4140
      Width           =   9240
      Begin VB.TextBox txtDescricaoPT 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         IMEMode         =   3  'DISABLE
         Left            =   1920
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         ToolTipText     =   "Descrição do posto de trabalho."
         Top             =   510
         Width           =   7035
      End
      Begin VB.TextBox txtusuario 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   480
         IMEMode         =   3  'DISABLE
         Left            =   2940
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Operador."
         Top             =   1440
         Width           =   3615
      End
      Begin VB.TextBox txtSenha 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
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
         ForeColor       =   &H00000000&
         Height          =   480
         IMEMode         =   3  'DISABLE
         Left            =   240
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   8
         ToolTipText     =   "Número do cracha do operador."
         Top             =   1440
         Width           =   2655
      End
      Begin VB.TextBox TxtPT 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   475
         IMEMode         =   3  'DISABLE
         Left            =   240
         MaxLength       =   20
         TabIndex        =   6
         ToolTipText     =   "Posto de trabalho."
         Top             =   510
         Width           =   1635
      End
      Begin VB.ComboBox txtfase 
         BackColor       =   &H80000014&
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
         ForeColor       =   &H00000000&
         Height          =   450
         Left            =   6570
         MouseIcon       =   "frmabrir_Ordem_CB.frx":030A
         MousePointer    =   99  'Custom
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   10
         ToolTipText     =   "Fase."
         Top             =   1440
         Width           =   1065
      End
      Begin VB.TextBox txtos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   450
         IMEMode         =   3  'DISABLE
         Left            =   7650
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Número da OS."
         Top             =   1440
         Width           =   1305
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição do posto de trabalho"
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
         Left            =   3487
         TabIndex        =   27
         Top             =   180
         Width           =   3900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Posto"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   607
         TabIndex        =   26
         Top             =   180
         Width           =   900
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
         Left            =   4155
         TabIndex        =   25
         Top             =   1080
         Width           =   1185
      End
      Begin VB.Label lblsenha 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código operador"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   270
         TabIndex        =   24
         Top             =   1080
         Width           =   2595
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "OS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   8085
         TabIndex        =   21
         Top             =   1080
         Width           =   435
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fase"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   6742
         TabIndex        =   20
         Top             =   1080
         Width           =   720
      End
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   600
      Left            =   180
      Picture         =   "frmabrir_Ordem_CB.frx":0614
      Top             =   120
      Width           =   1245
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Enter - Aceitar dados   /   Esc - Voltar a tela anterior"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   330
      Left            =   1830
      TabIndex        =   22
      Top             =   240
      Width           =   7005
   End
   Begin VB.Shape Shape2 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00000000&
      BorderWidth     =   3
      Height          =   735
      Left            =   30
      Top             =   30
      Width           =   9240
   End
End
Attribute VB_Name = "frmabrir_Ordem_CB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Maquina As String

Private Sub cmbPT_Change()
On Error GoTo tratar_erro

If txtordem.Text <> "" Then
    If IsNumeric(txtordem) = True Then
        If cmbPT.Text <> "" Then
            If IsNumeric(cmbPT.Text) = True Then
                txtfase.Clear
                txtos.Text = ""
                Set TBMaquinas = BD.OpenRecordset("Select maquina from CadMaquinas where IdMaquina = " & cmbPT.Text & "")
                If TBMaquinas.EOF = False Then Maquina = TBMaquinas!Maquina Else Maquina = ""
                TBMaquinas.Close
                Set TBProducao = BD.OpenRecordset("Select * from ordemservico where of = " & txtordem.Text & " and maquina = '" & Maquina & "' and pronto = 'NÃO' order by idproducao")
                If TBProducao.EOF = False Then
                    Do While TBProducao.EOF = False
                        txtfase.AddItem TBProducao!Fase
                        TBProducao.MoveNext
                    Loop
                Else
                    Set TBProducao = BD.OpenRecordset("Select * from ordemservico where of = " & txtordem.Text & " order by idproducao")
                    If TBProducao.EOF = False Then
                        TBProducao.MoveLast
                        txtfase.AddItem TBProducao!Fase + 10
                        txtos.Text = ""
                    Else
                        txtfase.AddItem 10
                    End If
                End If
                TBProducao.Close
                txtSenha.Enabled = True
            Else
                txtSenha.Enabled = False
            End If
        Else
            txtSenha.Enabled = False
        End If
    End If
End If
If cmbPT <> "" Then
    If IsNumeric(cmbPT) = True Then
        Set TBMaquinas = BD.OpenRecordset("Select Descricao from CadMaquinas where IDMaquina = " & cmbPT & "")
        If TBMaquinas.EOF = False Then
            txtDescricaoPT = TBMaquinas!Descricao
        End If
        TBMaquinas.Close
    Else
        MsgBox ("Só é permitido número neste campo."), vbExclamation
        cmbPT = ""
        txtDescricaoPT = ""
        cmbPT.SetFocus
        Exit Sub
    End If
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

Caption = "Gerprod  - Coletor de dados no chão de fábrica - Código de barras - Empresa : " & Empresa & ""

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub txtfase_Click()
On Error GoTo tratar_erro

If txtordem.Text <> "" Then
    txtos.Text = ""
    Set TBProducao = BD.OpenRecordset("Select * from ordemservico where of = " & txtordem.Text & " and maquina = '" & Maquina & "' and fase = " & txtfase.Text & " and pronto = 'NÃO' order by idproducao")
    If TBProducao.EOF = False Then
        txtos.Text = TBProducao!IDProducao
    End If
    TBProducao.Close
    ProcVerificaDisponMaquina
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub txtordem_Change()
On Error GoTo tratar_erro

If txtordem.Text <> "" Then
    If IsNumeric(txtordem.Text) = True Then
        Set TBMaquinas = BD.OpenRecordset("Select * from producao where of = " & txtordem.Text & " and Concluida = False and Status <> 'Cancelada'")
        If TBMaquinas.EOF = False Then
            txtitem.Text = TBMaquinas!desenho
            txtquant.Text = TBMaquinas!Quant
            mskprazofinal.Text = Format(TBMaquinas!prazoentrega, "dd/mm/yy")
            txtcliente.Text = IIf(IsNull(TBMaquinas!cliente) = False, TBMaquinas!cliente, "")
            txtdescricao.Text = TBMaquinas!produto
            cmbPT.Enabled = True
        Else
            ProcLimpaCampos
            cmbPT.Enabled = False
            txtSenha.Enabled = False
            txtfase.Enabled = False
        End If
    Else
        ProcLimpaCampos
        cmbPT.Enabled = False
        txtSenha.Enabled = False
        txtfase.Enabled = False
    End If
Else
    ProcLimpaCampos
    cmbPT.Enabled = False
    txtSenha.Enabled = False
    txtfase.Enabled = False
End If

Exit Sub
tratar_erro:
    If Err.Number = "91" Then
        MsgBox ("Não foi possivel encontrar o arquivo " & Localbd & ""), vbExclamation
        frmOpcoesGeral2.Show 1
        Exit Sub
    End If
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcLimpaCampos()
On Error GoTo tratar_erro

txtitem.Text = ""
txtquant.Text = ""
mskprazofinal.Text = "__/__/__"
txtcliente.Text = ""
txtdescricao.Text = ""
cmbPT.Text = ""
txtDescricaoPT = ""
txtSenha.Text = ""
txtusuario.Text = ""
txtfase.Clear
txtos.Text = ""
txtcliente.Text = ""

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcVerificaDisponMaquina()
On Error GoTo tratar_erro

If txtos.Text <> "" Then
    Set TBMaquinas = BD.OpenRecordset("Select * from CadMaquinas where maquina = '" & Maquina & "' and Liberada = false and OS <> " & txtos.Text & "")
Else
    Set TBMaquinas = BD.OpenRecordset("Select * from CadMaquinas where maquina = '" & Maquina & "' and liberada = false")
End If
If TBMaquinas.EOF = False Then
    MsgBox ("Não é permitido utilizar essa máquina, pois a mesma já está sendo utilizada na OS: " & TBMaquinas!OS & "."), vbExclamation
    TBMaquinas.Close
    Exit Sub
End If
TBMaquinas.Close

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

If KeyCode = 13 Then
    If txtordem.Text = "" Then
        MsgBox ("Informe o número da ordem antes de acessar."), vbExclamation
        txtordem.SetFocus
        Exit Sub
    End If
    If cmbPT.Enabled = False Then Exit Sub
    If cmbPT.Text = "" Then
        MsgBox ("Informe o posto de trabalho antes de acessar."), vbExclamation
        cmbPT.SetFocus
        Exit Sub
    End If
    If txtusuario.Text = "" Then
        MsgBox ("Informe a senha antes de acessar."), vbExclamation
        txtSenha.SetFocus
        Exit Sub
    End If
    If txtfase.Text = "" Then
        MsgBox ("Informe a fase antes de acessar."), vbExclamation
        txtfase.SetFocus
        Exit Sub
    End If
    frmProducao_CB.procabrirNovo
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
    
'Call procKeyascii(KeyAscii)

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub TxtPT_Change()

End Sub

Private Sub txtSenha_Change()
On Error GoTo tratar_erro

Set tbusuario = BD.OpenRecordset("Select * from usuarios where codigo = '" & txtSenha.Text & "' and Bloqueado = False")
If tbusuario.EOF = False Then
    txtusuario.Text = tbusuario("usuario")
    txtfase.Enabled = True
Else
    txtusuario.Text = ""
    txtfase.Enabled = False
    Exit Sub
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub


VERSION 5.00
Begin VB.Form frmabrir_CB 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Gerprod  - Coletor de dados no chão de fábrica - Código de barras"
   ClientHeight    =   7005
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   9315
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmabrir_CB.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7005
   ScaleWidth      =   9315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2325
      Left            =   30
      TabIndex        =   12
      Top             =   810
      Width           =   9270
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
         MultiLine       =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         ToolTipText     =   "Descrição."
         Top             =   1500
         Width           =   8775
      End
      Begin VB.TextBox txtdesenho 
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
         Left            =   3450
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Código interno."
         Top             =   645
         Width           =   1905
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
         Height          =   450
         IMEMode         =   3  'DISABLE
         Left            =   5370
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Quantidade."
         Top             =   645
         Width           =   1785
      End
      Begin VB.TextBox txtprazo 
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
         Left            =   7170
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Prazo final."
         Top             =   645
         Width           =   1815
      End
      Begin VB.TextBox TXTOF 
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
         Left            =   2100
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "Nº da ordem."
         Top             =   645
         Width           =   1335
      End
      Begin VB.TextBox cmbos 
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
         Height          =   450
         IMEMode         =   3  'DISABLE
         Left            =   240
         MaxLength       =   20
         TabIndex        =   0
         ToolTipText     =   "Número da ordem de serviço."
         Top             =   645
         Width           =   1845
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ordem"
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
         Left            =   2340
         TabIndex        =   18
         Top             =   300
         Width           =   855
      End
      Begin VB.Label Label7 
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
         Left            =   4027
         TabIndex        =   17
         Top             =   1155
         Width           =   1200
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cód. interno"
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
         Left            =   3630
         TabIndex        =   16
         Top             =   300
         Width           =   1545
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quant."
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
         Left            =   5835
         TabIndex        =   15
         Top             =   300
         Width           =   855
      End
      Begin VB.Label Label10 
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
         Left            =   7432
         TabIndex        =   14
         Top             =   300
         Width           =   1290
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N° OS"
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
         Left            =   697
         TabIndex        =   13
         ToolTipText     =   "N° da ordem de serviço"
         Top             =   300
         Width           =   930
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Frame2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   30
      TabIndex        =   19
      Top             =   2880
      Width           =   9270
      Begin VB.TextBox txtdescmaq 
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
         Left            =   3690
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Descrição do posto de trabalho."
         Top             =   765
         Width           =   5325
      End
      Begin VB.TextBox txtfase 
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
         MaxLength       =   20
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Fase."
         Top             =   765
         Width           =   1515
      End
      Begin VB.TextBox txtis 
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
         Height          =   915
         IMEMode         =   3  'DISABLE
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         TabStop         =   0   'False
         ToolTipText     =   "Instrução do serviço."
         Top             =   1725
         Width           =   8775
      End
      Begin VB.ComboBox cmbmaquina 
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
         Left            =   1770
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   7
         ToolTipText     =   "Código do posto de trabalho."
         Top             =   765
         Width           =   1905
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fase"
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
         Left            =   712
         TabIndex        =   23
         Top             =   435
         Width           =   570
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Posto trab."
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
         Left            =   1845
         TabIndex        =   22
         Top             =   435
         Width           =   1755
      End
      Begin VB.Label Label11 
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
         Left            =   4402
         TabIndex        =   21
         Top             =   435
         Width           =   3900
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Instrução do serviço"
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
         Left            =   3360
         TabIndex        =   20
         Top             =   1395
         Width           =   2535
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Frame3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   30
      TabIndex        =   24
      Top             =   5610
      Width           =   9270
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
         Height          =   450
         IMEMode         =   3  'DISABLE
         Left            =   2880
         Locked          =   -1  'True
         MaxLength       =   20
         TabIndex        =   11
         TabStop         =   0   'False
         ToolTipText     =   "Operador."
         Top             =   630
         Width           =   6135
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
         Height          =   450
         IMEMode         =   3  'DISABLE
         Left            =   240
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   10
         ToolTipText     =   "Número do cracha do operador."
         Top             =   630
         Width           =   2625
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
         Left            =   5355
         TabIndex        =   26
         Top             =   270
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
         Left            =   255
         TabIndex        =   25
         Top             =   270
         Width           =   2595
      End
   End
   Begin VB.Label Label9 
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
      TabIndex        =   27
      Top             =   240
      Width           =   7005
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   600
      Left            =   180
      Picture         =   "frmabrir_CB.frx":0E42
      Top             =   120
      Width           =   1245
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
Attribute VB_Name = "frmabrir_CB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub ProcChecaCampos()
On Error GoTo tratar_erro

If cmbos.Text = "" Then
    MsgBox ("Informa a OS antes de acessar."), vbExclamation
    cmbos.SetFocus
    Exit Sub
End If
If txtusuario.Text = "" Then
    MsgBox ("Informe sua senha antes de acessar."), vbExclamation
    txtSenha.SetFocus
    Exit Sub
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub cmbmaquina_Click()
On Error GoTo tratar_erro

Set TBMaquinas = BD.OpenRecordset("Select Descricao from CadMaquinas where Maquina = '" & cmbmaquina & "'")
If TBMaquinas.EOF = False Then
    txtdescmaq = IIf(IsNull(TBMaquinas!Descricao), "", TBMaquinas!Descricao)
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub cmbos_Change()
On Error GoTo tratar_erro

If cmbos.Text <> "" Then
    If IsNumeric(cmbos) = True Then
        Set TBMaquinas = BD.OpenRecordset("Select ordemservico.* FROM ordemservico INNER JOIN producao ON ordemservico.OF = producao.OF where ordemservico.idproducao = " & cmbos.Text & " and ordemservico.pronto = 'NÃO' and producao.status <> 'Cancelada'")
        If TBMaquinas.EOF = False Then
            TXTOF.Text = IIf(IsNull(TBMaquinas!OF) = False, TBMaquinas!OF, "")
            txtfase.Text = IIf(IsNull(TBMaquinas!Fase) = False, TBMaquinas!Fase, "")
            txtis.Text = IIf(IsNull(TBMaquinas!descfase) = False, TBMaquinas!descfase, "")
            txtdescricao.Text = IIf(IsNull(TBMaquinas("descricao")) = False, TBMaquinas!Descricao, "")
            txtdesenho.Text = IIf(IsNull(TBMaquinas("desenho")) = False, TBMaquinas!desenho, "")
            txtprazo.Text = IIf(IsNull(TBMaquinas("prazofinal")) = False, TBMaquinas!Prazofinal, "")
            txtquant.Text = IIf(IsNull(TBMaquinas("quantidade")) = False, TBMaquinas!Quantidade, "")
            OSControlada = TBMaquinas!OSControlada
            Processo_controlado = TBMaquinas!Processo_controlado
            Set TBProducao = BD.OpenRecordset("Select * from ProducaoFases where OS = " & TBMaquinas!IDProducao & "")
            If TBProducao.EOF = False Then
                cmbmaquina.Enabled = False
            End If
            TBProducao.Close
            txtSenha.Enabled = True
            cmbmaquina.Clear
            If IsNull(TBMaquinas!Maquina) = False And TBMaquinas!Maquina <> "" Then
                Set TBMaquinas = BD.OpenRecordset("Select Maquina, Grupo from CadMaquinas where Maquina = '" & TBMaquinas!Maquina & "'")
                If TBMaquinas.EOF = False Then
                    cmbmaquina.AddItem TBMaquinas!Maquina
                    If IsNull(TBMaquinas!Grupo) = False And TBMaquinas!Grupo <> "" Then
                        Set TBOrdem = BD.OpenRecordset("Select Maquina from CadMaquinas where Maquina <> '" & TBMaquinas!Maquina & "' and Grupo = '" & TBMaquinas!Grupo & "'")
                        If TBOrdem.EOF = False Then
                            Do While TBOrdem.EOF = False
                                cmbmaquina.AddItem TBOrdem!Maquina
                                TBOrdem.MoveNext
                            Loop
                        End If
                        TBOrdem.Close
                    End If
                    cmbmaquina = TBMaquinas!Maquina
                End If
            End If
        Else
            ProcLimpaCampos
            cmbmaquina.Enabled = False
            txtSenha.Enabled = False
        End If
        TBMaquinas.Close
    Else
        ProcLimpaCampos
        cmbmaquina.Enabled = False
        txtSenha.Enabled = False
    End If
Else
    ProcLimpaCampos
    cmbmaquina.Enabled = False
    txtSenha.Enabled = False
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

TXTOF.Text = ""
txtfase.Text = ""
txtis.Text = ""
txtdescricao.Text = ""
txtdesenho.Text = ""
txtprazo.Text = ""
txtquant.Text = ""
cmbmaquina.Clear
txtdescmaq.Text = ""
txtSenha.Text = ""
txtusuario.Text = ""

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

If txtSenha.Text <> "" And txtusuario.Text <> "" And cmbos.Text <> "" Then
    ProcChecaCampos
    If KeyCode = 13 Then
        frmProducao_CB.procabrir
        frmProducao_CB.ProcLista12Ultimos
        If OrdemExiste = True Then Unload Me
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
    
Call ProcKeyAscii(KeyAscii)

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

Private Sub txtSenha_Change()
On Error GoTo tratar_erro

If txtSenha.Text <> "" Then
    Set TBUsuarios = BD.OpenRecordset("Select * from usuarios where codigo = '" & txtSenha.Text & "' and Bloqueado = False")
    If TBUsuarios.EOF = False Then
        txtusuario.Text = TBUsuarios("usuario")
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

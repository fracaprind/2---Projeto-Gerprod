VERSION 5.00
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmabrir_plano_prod 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   Caption         =   "Gerprod  - Coletor de dados no chão de fábrica"
   ClientHeight    =   9045
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   12645
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmabrir_plano_prod.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   Picture         =   "frmabrir_plano_prod.frx":0E42
   ScaleHeight     =   9045
   ScaleWidth      =   12645
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtos 
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
      Left            =   450
      MaxLength       =   20
      TabIndex        =   23
      TabStop         =   0   'False
      ToolTipText     =   "Número da ordem de serviço."
      Top             =   9135
      Visible         =   0   'False
      Width           =   1845
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
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
      Left            =   1725
      TabIndex        =   11
      Top             =   2595
      Width           =   9000
      Begin VB.TextBox Txt_ID_apontamento 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   180
         MaxLength       =   20
         TabIndex        =   0
         ToolTipText     =   "Número para apontamento."
         Top             =   540
         Width           =   1815
      End
      Begin VB.TextBox txt_data 
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
         Left            =   4185
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         ToolTipText     =   "Data."
         Top             =   540
         Width           =   1275
      End
      Begin VB.TextBox Txt_responsavel 
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
         Left            =   5520
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         ToolTipText     =   "Responsável."
         Top             =   540
         Width           =   3255
      End
      Begin VB.TextBox Txt_plano 
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
         Left            =   2025
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         ToolTipText     =   "Número do plano."
         Top             =   540
         Width           =   2115
      End
      Begin VB.TextBox Txt_obs 
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
         Height          =   720
         IMEMode         =   3  'DISABLE
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         ToolTipText     =   "Observações."
         Top             =   1395
         Width           =   8565
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Data"
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
         Left            =   4530
         TabIndex        =   22
         Top             =   180
         Width           =   585
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Responsável"
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
         Left            =   6345
         TabIndex        =   21
         Top             =   180
         Width           =   1545
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N° plano"
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
         Left            =   2527
         TabIndex        =   20
         Top             =   180
         Width           =   1110
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Observações"
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
         Left            =   3675
         TabIndex        =   19
         Top             =   1020
         Width           =   1575
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N° apont."
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
         Left            =   330
         TabIndex        =   12
         ToolTipText     =   "N° da ordem de serviço"
         Top             =   180
         Width           =   1515
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1725
      TabIndex        =   13
      Top             =   5085
      Width           =   9000
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
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         ToolTipText     =   "Descrição do posto de trabalho."
         Top             =   540
         Width           =   6465
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
         Left            =   180
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         ToolTipText     =   "Código do posto de trabalho."
         Top             =   540
         Width           =   2085
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
         Left            =   345
         TabIndex        =   15
         Top             =   180
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
         Left            =   3562
         TabIndex        =   14
         Top             =   180
         Width           =   3900
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   1725
      TabIndex        =   16
      Top             =   6495
      Width           =   9000
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
         Left            =   2850
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "Operador."
         Top             =   570
         Width           =   5955
      End
      Begin VB.TextBox txtSenha 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
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
         Left            =   180
         MaxLength       =   20
         PasswordChar    =   "*"
         TabIndex        =   7
         Top             =   570
         Width           =   2625
      End
      Begin VB.Label lblcodigosenha 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código/Senha"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   382
         TabIndex        =   24
         Top             =   180
         Visible         =   0   'False
         Width           =   2220
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
         Left            =   5235
         TabIndex        =   18
         Top             =   180
         Width           =   1185
      End
      Begin VB.Label lblsenha 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Senha"
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
         Left            =   1012
         TabIndex        =   17
         Top             =   180
         Width           =   960
      End
   End
   Begin DrawSuite2022.USButton Cmd_enter 
      Height          =   600
      Left            =   2700
      TabIndex        =   9
      Top             =   1545
      Width           =   3150
      _ExtentX        =   5556
      _ExtentY        =   1058
      BorderColor     =   14404026
      BorderColorDown =   11632444
      BorderColorOver =   11632444
      ButtonShape     =   1
      Caption         =   "Enter - Aceitar dados"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GradientColor2  =   16777215
      GradientColor3  =   16777215
      GradientColorDown2=   16246986
      GradientColorDown3=   15189380
      GradientColorDown4=   14596208
      GradientColorOver1=   16643560
      GradientColorOver2=   16576988
      GradientColorOver3=   16441780
      GradientColorOver4=   16178091
      PicAlign        =   8
      PicSize         =   4
      PicSizeH        =   48
      PicSizeW        =   48
   End
   Begin DrawSuite2022.USButton Cmd_esc 
      Height          =   600
      Left            =   5910
      TabIndex        =   10
      Top             =   1545
      Width           =   3150
      _ExtentX        =   5556
      _ExtentY        =   1058
      BorderColor     =   14404026
      BorderColorDown =   11632444
      BorderColorOver =   11632444
      ButtonShape     =   1
      Caption         =   "Esc - Voltar a tela anterior"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   192
      GradientColor2  =   16777215
      GradientColor3  =   16777215
      GradientColorDown2=   16246986
      GradientColorDown3=   15189380
      GradientColorDown4=   14596208
      GradientColorOver1=   16643560
      GradientColorOver2=   16576988
      GradientColorOver3=   16441780
      GradientColorOver4=   16178091
      PicAlign        =   8
      PicSize         =   4
      PicSizeH        =   48
      PicSizeW        =   48
   End
   Begin ActiveResizeCtl.ActiveResize ActiveResize1 
      Left            =   135
      Top             =   315
      _ExtentX        =   847
      _ExtentY        =   847
      Resolution      =   99
      ScreenHeight    =   768
      ScreenWidth     =   1366
      ScreenHeightDT  =   1080
      ScreenWidthDT   =   1920
      AutoResizeOnLoad=   0   'False
      ApplicationName =   "Active Resize Control Professional"
      FormHeightDT    =   9510
      FormWidthDT     =   12765
      FormScaleHeightDT=   9045
      FormScaleWidthDT=   12645
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin VB.Label lblVersaoatual 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2.5"
      BeginProperty Font 
         Name            =   "Bodoni Bk BT"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11000
      TabIndex        =   28
      Top             =   400
      Width           =   450
   End
   Begin VB.Label lblano 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Copyright 1999 - 2018 Caprind Sistemas ®. Todos os direitos reservados."
      BeginProperty Font 
         Name            =   "@Malgun Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   3675
      TabIndex        =   27
      Top             =   8505
      Width           =   5400
   End
   Begin VB.Label lbldireitos 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   $"frmabrir_plano_prod.frx":7E6B4
      BeginProperty Font 
         Name            =   "@Malgun Gothic"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   225
      Left            =   1110
      TabIndex        =   26
      Top             =   8730
      Width           =   10320
   End
   Begin VB.Label Image1 
      Caption         =   "Label14"
      BeginProperty Font 
         Name            =   "C39HrP36DlTt"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   9180
      TabIndex        =   25
      Top             =   1530
      Width           =   1500
   End
   Begin DrawSuite2022.USAlphaImage USAlphaImage1 
      Height          =   9060
      Left            =   0
      Top             =   0
      Width           =   12645
      _ExtentX        =   22304
      _ExtentY        =   15981
      Image           =   "frmabrir_plano_prod.frx":7E745
   End
End
Attribute VB_Name = "frmabrir_plano_prod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbmaquina_Click()
On Error GoTo tratar_erro

txtSenha = ""
txtSenha.Enabled = False
Set TBMaquinas = CreateObject("adodb.recordset")
TBMaquinas.Open "Select Descricao from CadMaquinas where Maquina = '" & cmbmaquina & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBMaquinas.EOF = False Then
    txtdescmaq = IIf(IsNull(TBMaquinas!Descricao), "", TBMaquinas!Descricao)
    txtSenha.Enabled = True
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub cmbmaquina_KeyPress(KeyAscii As Integer)
On Error GoTo tratar_erro
    
If KeyAscii = vbKeyReturn Then Sendkeys "{TAB}": KeyAscii = 0

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcLimpaCampos()
On Error GoTo tratar_erro

Txt_plano = ""
txt_data = ""
Txt_responsavel = ""
Txt_obs = ""
cmbmaquina.Clear
txtdescmaq = ""
txtSenha = ""
txtusuario = ""

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Cmd_enter_Click()
On Error GoTo tratar_erro

Call Form_KeyDown(13, 0)

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Cmd_esc_Click()
On Error GoTo tratar_erro

Call Form_KeyDown(27, 0)

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

WindowState = 2

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Txt_ID_apontamento_KeyPress(KeyAscii As Integer)
On Error GoTo tratar_erro
    
If KeyAscii = vbKeyReturn Then Sendkeys "{TAB}": KeyAscii = 0

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

If Atualizando = True Then Exit Sub

Select Case KeyCode
    Case vbKeyReturn:
        ProcVerifUsuario
        If Txt_ID_apontamento = "" Or cmbmaquina = "" Or txtusuario = "" Then Exit Sub
        ProcLogonOut (txtusuario)
        If TemInternet = True And ErroDriverMYSQL = False Then
            If FunValidarCliente(txtusuario) = False Then
                Unload Me
                frmfundo.Show
                Exit Sub
            End If
            If FunValidarClienteSemInternet = False Then
                Unload Me
                frmfundo.Show
                Exit Sub
            End If
        Else
            If ErroDriverMYSQL = True Then MsgBox ("O driver MySQL não foi instalado corretamente, favor verificar."), vbInformation
        End If
        If FunLogonIn(txtusuario) = False Then
            Unload Me
            frmfundo.Show
            Exit Sub
        End If
        
        ProcVerifAPCodigo True, Txt_ID_apontamento
        With frmProducao
            .ProcAbrir
            .ProcLista12Ultimos
        End With
        If OrdemExiste = True Then Unload Me
    Case vbKeyEscape:
        Unload Me
        frmfundo.Show
End Select

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo tratar_erro
    
'If Codigo_Barras = False Then Call FunKeyAscii(KeyAscii)

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

'If xPixels = 800 Then Img_800x600.Visible = True Else Img_1024x768.Visible = True
If Codigo_Barras = True Then
    Caption = "Gerprod - Coletor de dados no chão de fábrica - Código de barras - Nome do banco de dados: " & Nome_banco & " | Licenças Gerprod: " & Qtlicencas_gerprod & IIf(TemInternet = False, " | Sem internet", "")
    Image1.Visible = True
    lblsenha.Visible = False
    lblcodigosenha.Visible = True
    txtSenha.ToolTipText = "Código/Senha do usuário"
Else
    Caption = "Gerprod - Coletor de dados no chão de fábrica - Nome do banco de dados: " & Nome_banco & " | Licenças Gerprod: " & Qtlicencas_gerprod & IIf(TemInternet = False, " | Sem internet", "")
    Image1.Visible = False
    lblsenha.Visible = True
    lblcodigosenha.Visible = False
    txtSenha.ToolTipText = "Senha do usuário"
End If

lblVersaoatual.Caption = "v" & App.Major & "." & App.Minor & "." & App.Revision
lblano.Caption = "Copyright 1999 - " & Year(Date) & " Caprind Sistemas ®. Todos os direitos reservados."

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Txt_ID_apontamento_Change()
On Error GoTo tratar_erro

ProcLimpaCampos
If Txt_ID_apontamento <> "" Then
    VerifNumero = Txt_ID_apontamento
    ProcVerificaNumero
    If VerifNumero = False Then
        Txt_ID_apontamento = ""
        Txt_ID_apontamento.SetFocus
        Exit Sub
    End If
    
    Set TBMaquinas = CreateObject("adodb.recordset")
    TBMaquinas.Open "Select PFO.Planotexto, PFO.Data, PFO.Responsavel, PFO.Observacao, OS.IDproducao, OS.OSControlada, OS.Processo_controlado, OS.IDFase, OS.Maquina, P.ID_empresa from (ProducaoFases_OS PFO INNER JOIN ordemservico OS ON OS.ID_apontamento = PFO.ID) INNER JOIN Producao P ON P.Ordem = OS.Ordem where PFO.ID = " & Txt_ID_apontamento & " and OS.pronto = 'NÃO' and P.status <> 'Cancelada' and PFO.DtValidacao IS NOT NULL and P.DtValidacao IS NOT NULL and P.DtValidacao_Custo IS NULL", Conexao, adOpenKeyset, adLockOptimistic
    If TBMaquinas.EOF = False Then
        txtOS = TBMaquinas!IDProducao
        Txt_plano = IIf(IsNull(TBMaquinas!Planotexto), "", TBMaquinas!Planotexto)
        txt_data = IIf(IsNull(TBMaquinas!Data), "", Format(TBMaquinas!Data, "dd/mm/yy"))
        Txt_responsavel = IIf(IsNull(TBMaquinas!Responsavel), "", TBMaquinas!Responsavel)
        Txt_obs = IIf(IsNull(TBMaquinas!Observacao), "", TBMaquinas!Observacao)
        
        OSControlada = TBMaquinas!OSControlada
        Processo_controlado = TBMaquinas!Processo_controlado
        
        'Versao do processo
        Versao_processo = "A"
        Set TBFases = CreateObject("adodb.recordset")
        TBFases.Open "Select Versao from Fases where IDFase = " & TBMaquinas!IDFase, Conexao, adOpenKeyset, adLockOptimistic
        If TBFases.EOF = False Then
            Versao_processo = IIf(IsNull(TBFases!Versao), "A", TBFases!Versao)
        End If
        TBFases.Close
        
        With cmbmaquina
            .Enabled = True
            .Clear
            If IsNull(TBMaquinas!Maquina) = False And TBMaquinas!Maquina <> "" Then
                'Verifica se a opção de carregar somente as maquinas do grupo da OS esta liberada
                TextoFiltro = ""
                Set TBFiltro = CreateObject("adodb.recordset")
                TBFiltro.Open "Select * from empresa where CNPJ = '" & CNPJEmpresa & "' and Codigo = " & TBMaquinas!ID_empresa & " and Grupo_Gerprod = 'True'", Conexao, adOpenKeyset, adLockOptimistic
                If TBFiltro.EOF = False Then
                    Set TBCFOP = CreateObject("adodb.recordset")
                    TBCFOP.Open "Select Grupo from CadMaquinas where Maquina = '" & TBMaquinas!Maquina & "'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBCFOP.EOF = False Then TextoFiltro = "and grupo = '" & TBCFOP!Grupo & "'"
                    TBCFOP.Close
                End If
                TBFiltro.Close
                
                Set TBOrdem = CreateObject("adodb.recordset")
                TBOrdem.Open "Select * from CadMaquinas where Bloqueado = 'False' " & TextoFiltro & " order by Maquina", Conexao, adOpenKeyset, adLockOptimistic
                If TBOrdem.EOF = False Then
                    Do While TBOrdem.EOF = False
                        If TBOrdem!Liberada = False Then
                            Set TBAbrir = CreateObject("adodb.recordset")
                            TBAbrir.Open "Select Maquina from CadMaquinas_Monitor where Maquina = '" & TBOrdem!Maquina & "' and OS = " & txtOS, Conexao, adOpenKeyset, adLockOptimistic
                            If TBAbrir.EOF = False Then
                                .AddItem TBOrdem!Maquina
                            End If
                            TBAbrir.Close
                        Else
                            .AddItem TBOrdem!Maquina
                        End If
                        TBOrdem.MoveNext
                    Loop
                End If
                
                Set TBOrdem = CreateObject("adodb.recordset")
                TBOrdem.Open "Select Maquina from Ordemservico_maq_utilizadas where OS = " & txtOS & " order by ID", Conexao, adOpenKeyset, adLockOptimistic
                If TBOrdem.EOF = False Then
                    TBOrdem.MoveLast
                    
                    Set TBAbrir = CreateObject("adodb.recordset")
                    TBAbrir.Open "Select CadMaquinas.Maquina from CadMaquinas INNER JOIN CadMaquinas_Monitor ON CadMaquinas.Maquina = CadMaquinas_Monitor.Maquina where CadMaquinas.Maquina = '" & TBOrdem!Maquina & "' and CadMaquinas.Liberada = 'False' and CadMaquinas.Bloqueado = 'False' and CadMaquinas_Monitor.OS <> " & txtOS, Conexao, adOpenKeyset, adLockOptimistic
                    If TBAbrir.EOF = False Then
                        .AddItem TBOrdem!Maquina
                    End If
                    TBAbrir.Close
                    
                    .text = TBOrdem!Maquina
                Else
                    'Verifica se a máquina prevista esta sendo utilizada pela OS a ser apontada
                    Set TBAbrir = CreateObject("adodb.recordset")
                    TBAbrir.Open "Select CadMaquinas.Maquina from CadMaquinas INNER JOIN CadMaquinas_Monitor ON CadMaquinas.Maquina = CadMaquinas_Monitor.Maquina where CadMaquinas.Maquina = '" & TBMaquinas!Maquina & "' and CadMaquinas.Liberada = 'False' and CadMaquinas.Bloqueado = 'False' and CadMaquinas_Monitor.OS = " & txtOS, Conexao, adOpenKeyset, adLockOptimistic
                    If TBAbrir.EOF = False Then
                        .text = TBMaquinas!Maquina
                    Else
                        'Verifica se a máquina prevista esta liberada
                        Set TBAbrir = CreateObject("adodb.recordset")
                        TBAbrir.Open "Select * from CadMaquinas where Maquina = '" & TBMaquinas!Maquina & "' and Bloqueado = 'False' and Liberada = 'True'", Conexao, adOpenKeyset, adLockOptimistic
                        If TBAbrir.EOF = False Then
                            .text = TBMaquinas!Maquina
                        End If
                    End If
                    TBAbrir.Close
                End If
                TBOrdem.Close
            End If
        End With
    Else
        cmbmaquina.Enabled = False
        txtSenha.Enabled = False
    End If
    TBMaquinas.Close
Else
    cmbmaquina.Enabled = False
    txtSenha.Enabled = False
End If

Exit Sub
tratar_erro:
    If Err.Number = -2147467259 Or Err.Number = 3709 Then
        MsgBox ("Não foi possível estabelecer a conexão com o banco de dados, o sistema será fechado."), vbCritical
        End
    End If
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub txtSenha_Change()
On Error GoTo tratar_erro

txtusuario.text = ""

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub txtSenha_LostFocus()
On Error GoTo tratar_erro

ProcVerifUsuario

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcVerifUsuario()
On Error GoTo tratar_erro

If txtSenha.text <> "" Then
    Set TBUsuarios = CreateObject("adodb.recordset")
    If Codigo_Barras = True Then TextoFiltro = "(Senha = '" & txtSenha.text & "' or Codigo = '" & txtSenha.text & "')" Else TextoFiltro = "Senha = '" & txtSenha.text & "'"
    TBUsuarios.Open "Select Usuario from usuarios where Bloqueado = 'False' and " & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
    If TBUsuarios.EOF = False Then
        txtusuario.text = TBUsuarios!Usuario
        Operador = txtusuario
    End If
    TBUsuarios.Close
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub



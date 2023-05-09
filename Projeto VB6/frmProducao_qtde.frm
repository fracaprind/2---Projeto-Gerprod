VERSION 5.00
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{50D37AD9-8D3C-43DD-ADD5-7C957C951843}#1.9#0"; "FlexCell.ocx"
Begin VB.Form frmProducao_qtde 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Gerprod | Informe as quantidades"
   ClientHeight    =   10575
   ClientLeft      =   0
   ClientTop       =   30
   ClientWidth     =   6540
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProducao_qtde.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10575
   ScaleWidth      =   6540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   26
      Top             =   10170
      Width           =   6540
      _ExtentX        =   11536
      _ExtentY        =   714
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   540
      Left            =   0
      TabIndex        =   25
      Top             =   0
      Width           =   6540
      _ExtentX        =   11536
      _ExtentY        =   953
      DibPicture      =   "frmProducao_qtde.frx":0E42
      CaptionDelimiter=   "|"
      CaptionOnCenter =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Icon            =   "frmProducao_qtde.frx":82E7
      IconSize        =   1
      IconSizeX       =   24
      IconSizeY       =   24
      ShowMaximizeButton=   0   'False
      ShowMinimizeButton=   0   'False
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ordem produção"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   825
      Left            =   180
      TabIndex        =   23
      Top             =   660
      Width           =   1875
      Begin VB.TextBox txtOP 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         ToolTipText     =   "Número da OS."
         Top             =   300
         Width           =   1515
      End
   End
   Begin FlexCell.Grid GridRE 
      Height          =   6885
      Left            =   180
      TabIndex        =   22
      Top             =   3210
      Width           =   6195
      _ExtentX        =   10927
      _ExtentY        =   12144
      AllowUserReorderColumn=   -1  'True
      AllowUserResizing=   0   'False
      Appearance      =   0
      BackColor2      =   14737632
      BackColorBkg    =   -2147483644
      DefaultFontSize =   8.25
      GridColor       =   12632256
      Rows            =   1
      MultiSelect     =   0   'False
      AllowUserPaste  =   3
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Condicional"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   855
      Left            =   2910
      TabIndex        =   20
      Top             =   1500
      Width           =   1305
      Begin VB.TextBox txtQTCD 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   405
         Left            =   120
         TabIndex        =   2
         Text            =   "0"
         ToolTipText     =   "Quantidade não conforme."
         Top             =   300
         Width           =   1050
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   2415
      Left            =   4230
      TabIndex        =   19
      Top             =   660
      Width           =   2115
      Begin VB.CommandButton Cmd_backspace 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "<-------"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   825
         MaskColor       =   &H00404040&
         MouseIcon       =   "frmProducao_qtde.frx":9139
         Style           =   1  'Graphical
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   1800
         Width           =   990
      End
      Begin VB.CommandButton Number 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   0
         Left            =   285
         MaskColor       =   &H00404040&
         MouseIcon       =   "frmProducao_qtde.frx":9443
         Style           =   1  'Graphical
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   1800
         Width           =   480
      End
      Begin VB.CommandButton Number 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   3
         Left            =   1335
         MaskColor       =   &H00404040&
         MouseIcon       =   "frmProducao_qtde.frx":974D
         Style           =   1  'Graphical
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   1260
         Width           =   480
      End
      Begin VB.CommandButton Number 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   2
         Left            =   825
         MaskColor       =   &H00404040&
         MouseIcon       =   "frmProducao_qtde.frx":9A57
         Style           =   1  'Graphical
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1260
         Width           =   480
      End
      Begin VB.CommandButton Number 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   1
         Left            =   285
         MaskColor       =   &H00404040&
         MouseIcon       =   "frmProducao_qtde.frx":9D61
         Style           =   1  'Graphical
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   1260
         Width           =   480
      End
      Begin VB.CommandButton Number 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   6
         Left            =   1335
         MaskColor       =   &H00404040&
         MouseIcon       =   "frmProducao_qtde.frx":A06B
         Style           =   1  'Graphical
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   750
         Width           =   480
      End
      Begin VB.CommandButton Number 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   5
         Left            =   825
         MaskColor       =   &H00404040&
         MouseIcon       =   "frmProducao_qtde.frx":A375
         Style           =   1  'Graphical
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   750
         Width           =   480
      End
      Begin VB.CommandButton Number 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   4
         Left            =   285
         MaskColor       =   &H00404040&
         MouseIcon       =   "frmProducao_qtde.frx":A67F
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   750
         Width           =   480
      End
      Begin VB.CommandButton Number 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   9
         Left            =   1335
         MaskColor       =   &H00404040&
         MouseIcon       =   "frmProducao_qtde.frx":A989
         Style           =   1  'Graphical
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   240
         Width           =   480
      End
      Begin VB.CommandButton Number 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   8
         Left            =   825
         MaskColor       =   &H00404040&
         MouseIcon       =   "frmProducao_qtde.frx":AC93
         Style           =   1  'Graphical
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   240
         Width           =   480
      End
      Begin VB.CommandButton Number 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Index           =   7
         Left            =   285
         MaskColor       =   &H00404040&
         MouseIcon       =   "frmProducao_qtde.frx":AF9D
         Style           =   1  'Graphical
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   240
         Width           =   480
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Ordem de serviço"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   825
      Left            =   2070
      TabIndex        =   18
      Top             =   660
      Width           =   2145
      Begin VB.TextBox Txt_OS 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   270
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         ToolTipText     =   "Número da OS."
         Top             =   300
         Width           =   1575
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Aprovado"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   180
      TabIndex        =   17
      Top             =   1500
      Width           =   1245
      Begin VB.TextBox txtTOK 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   405
         Left            =   120
         TabIndex        =   0
         ToolTipText     =   "Quantidade conforme."
         Top             =   300
         Width           =   1020
      End
   End
   Begin VB.Frame Frame15 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Não conforme"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   1440
      TabIndex        =   16
      Top             =   1500
      Width           =   1455
      Begin VB.TextBox txtTNC 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   405
         Left            =   180
         TabIndex        =   1
         Text            =   "0"
         ToolTipText     =   "Quantidade não conforme."
         Top             =   300
         Width           =   1140
      End
   End
   Begin DrawSuite2022.USButton Cmd_F3 
      Height          =   645
      Left            =   150
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   2430
      Width           =   1830
      _ExtentX        =   3228
      _ExtentY        =   1138
      Caption         =   "F3 - GRAVAR"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   5263559
      BorderColorDisabled=   13160660
      BorderColorDown =   4013465
      BorderColorOver =   4408288
      GradientColor1  =   5263559
      GradientColor2  =   5263559
      GradientColor3  =   5263559
      GradientColor4  =   5263559
      GradientColorDisabled1=   13160660
      GradientColorDisabled2=   13160660
      GradientColorDisabled3=   13160660
      GradientColorDisabled4=   13160660
      GradientColorOver1=   4408288
      GradientColorOver2=   4408288
      GradientColorOver3=   4408288
      GradientColorOver4=   4408288
      GradientColorDown1=   4013465
      GradientColorDown2=   4013465
      GradientColorDown3=   4013465
      GradientColorDown4=   4013465
      PicAlign        =   8
      PicSize         =   4
      PicSizeH        =   48
      PicSizeW        =   48
      ShowFocusRect   =   0   'False
      Theme           =   4
   End
   Begin DrawSuite2022.USButton Cmd_esc 
      Height          =   645
      Left            =   2070
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   2430
      Width           =   2130
      _ExtentX        =   3757
      _ExtentY        =   1138
      Caption         =   "ESC - VOLTAR"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   4960354
      BorderColorDisabled=   13160660
      BorderColorDown =   4210752
      BorderColorOver =   49152
      GradientColor1  =   4960354
      GradientColor2  =   4960354
      GradientColor3  =   4960354
      GradientColor4  =   4960354
      GradientColorDisabled1=   14215660
      GradientColorDisabled2=   14215660
      GradientColorDisabled3=   14215660
      GradientColorDisabled4=   14215660
      GradientColorOver1=   49152
      GradientColorOver2=   49152
      GradientColorOver3=   49152
      GradientColorOver4=   49152
      GradientColorDown1=   32768
      GradientColorDown2=   32768
      GradientColorDown3=   32768
      GradientColorDown4=   32768
      PicAlign        =   8
      PicSize         =   4
      PicSizeH        =   48
      PicSizeW        =   48
      ShowFocusRect   =   0   'False
      Theme           =   3
   End
End
Attribute VB_Name = "frmProducao_qtde"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Calc_OK As Boolean 'OK
Dim Linha As Integer 'OK
Public QtdeDescricaoNC As Integer 'OK

Private Sub btnNS_Click()
On Error GoTo tratar_erro

Contador = txtTNC.Text
frmNumeroSerie.Show 1

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

Private Sub Cmd_F3_Click()
On Error GoTo tratar_erro
Contador = 0

Call Form_KeyDown(114, 0)

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcCalculaTotalNC()
On Error GoTo tratar_erro
Dim QTNC As Double
Dim QTCD As Double
TOK = LOTE 'txtTOK
Contador = 0
QTNC = 0
QTCD = 0

With GridRE
Linha = .Rows - 1
    For initfor = 1 To Linha
        QTNC = QTNC + IIf(IsNumeric(GridRE.Cell(Contador, 2).Text) = True, GridRE.Cell(Contador, 2).Text, 0)
        QTCD = QTCD + IIf(IsNumeric(GridRE.Cell(Contador, 3).Text) = True, GridRE.Cell(Contador, 3).Text, 0)
        Contador = Contador + 1
    Next initfor
End With

txtTNC = QTNC
txtQTCD = QTCD
txtTOK = TOK - (QTNC + QTCD)

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro


Select Case KeyCode
    Case vbKeyF3:
    ProcGravar
    Case vbKeyEscape:
        Gravar = False
        Unload Me
End Select

 If KeyCode = 13 Then
 Sendkeys "{tab}" '13 = Enter Key
 End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Public Sub ProcGravar()
On Error GoTo tratar_erro

Acao = "salvar"
TOK = IIf(txtTOK = "", 0, txtTOK)
If txtTOK = "" Or TOK < 0 Then
    NomeCampo = "quantidade conforme"
    ProcVerificaAcao
    txtTOK.SetFocus
    Exit Sub
End If

TNC = IIf(txtTNC = "", 0, txtTNC)
If TNC < 0 Then
    NomeCampo = "quantidade não conforme"
    ProcVerificaAcao
    txtTNC.SetFocus
    Exit Sub
End If
If txtTNC = "" Then txtTNC = 0

TCD = IIf(txtQTCD = "", 0, txtQTCD)
If TCD < 0 Then
    NomeCampo = "quantidade não conforme"
    ProcVerificaAcao
    txtQTCD.SetFocus
    Exit Sub
End If

If txtQTCD = "" Then txtQTCD = 0


Gravar = True
If FunVerifQtdeOSProcControlado(Txt_OS) = False Then
    Gravar = False
    Exit Sub
End If

If FunVerificaEstoqueAP = False Then Exit Sub

If FunVerifNCPorDescricao(Txt_OS) = True Then
    ReDim ArrayQtdeDescNC(1 To QtdeDescricaoNC, 1 To 2)
    With GridRE
        For i = 1 To (.Rows - 1)
        If .Cell(i, 2).Text <> "" Or .Cell(i, 3).Text <> "" Then ProcGravaNCCQ
        Next
    End With
Else
ProcGravaNCCQ
End If

If OrdemRastreavel = True Then
 ProcQTRastreavel
End If


Unload Me

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Public Sub procGravarQTOKEtiqueta()
On Error GoTo tratar_erro

Contador = txtTOK
Status = "APROVADO"
IDApontamento = frmProducao.Lista.SelectedItem

'Verifica se é primeira OS
Set TBAbrir = CreateObject("adodb.recordset")
StrSQL = "Select * from Producao_rastreavel where Ordem = '" & Ordem & "' and OS = '" & OS - 1 & "' ORDER BY OS, N_serie"
TBAbrir.Open StrSQL, Conexao, adOpenKeyset, adLockOptimistic
  If TBAbrir.EOF = True Then
    StrSQL = "Select * from Producao_rastreavel where Ordem = '" & Ordem & "' AND Status = 'DISPONIVEL' AND OS = '" & OS & "' ORDER BY OS, N_serie"
  Else
    StrSQL = "Select * from Producao_rastreavel where Status <> 'NÃO CONFORME' and data is not null and OS = " & OS - 1
  End If
TBAbrir.Close

    Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open StrSQL, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Do While Contador > 0
            NumeroSerie = TBAbrir!N_Serie
                    
                Set TBAbrir_Status = CreateObject("adodb.recordset")
                    TBAbrir_Status.Open "Select * from ProducaoFases_Codigos", Conexao, adOpenKeyset, adLockOptimistic
                        TBAbrir_Status.AddNew
                            TBAbrir_Status!Data = Now
                            TBAbrir_Status!OS = OS
                            TBAbrir_Status!Codigo = NumeroSerie
                            TBAbrir_Status!Responsavel = Operador
                            TBAbrir_Status!Status = Status
                            TBAbrir_Status!IDProducao = IDApontamento
                        TBAbrir_Status.Update
                    TBAbrir_Status.Close
                Contador = Contador - 1
                TBAbrir.MoveNext
            Loop
    End If


Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Public Sub ProcBuscaDisponivel()
On Error GoTo tratar_erro
NumeroSerie = ""

StrSQL = "Select * from Producao_rastreavel where Ordem = '" & Ordem & "' AND Status = 'DISPONIVEL' AND OS = '" & OS & "' ORDER BY OS, N_serie"

Set TBFiltro = CreateObject("adodb.recordset")
'Debug.Print StrSQL

TBFiltro.Open StrSQL, Conexao, adOpenKeyset, adLockOptimistic
  If TBFiltro.EOF = False Then
    NumeroSerie = TBFiltro!N_Serie
  End If


TBFiltro.Close

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Public Sub procGravarQTNCEtiqueta()
On Error GoTo tratar_erro

With frmNumeroSerieOK
    .txtOK = txtTOK
    .txtNC = txtTNC
    .chkAprovar.Value = 1
    .chkAprovar.Enabled = False
    .ChkRegistro.Value = 1
    .ChkRegistro.Enabled = False
    .Show 1
    Gravar = False
End With

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Public Sub ProcQTRastreavel()
On Error GoTo tratar_erro


If IsNumeric(txtTNC) Then

    If txtTNC > 0 Then
     procGravarQTNCEtiqueta
    End If

End If

If IsNumeric(txtTOK) Then

    If txtTOK > 0 Then
     'procGravarQTOKEtiqueta
    End If

End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Public Sub ProcGravaNCCQ()
On Error GoTo tratar_erro
    
If TNC > 0 Or TCD > 0 Then
    qtdeNC_ordem = TNC
    CamposLoop = ""
    
    OperadorTexto = ""
    Set TBUsuarios = CreateObject("adodb.recordset")
    TBUsuarios.Open "Select CODIGO from Usuarios where Usuario = '" & Operador & "' and Bloqueado = 'False'", Conexao, adOpenKeyset, adLockOptimistic
    If TBUsuarios.EOF = False Then
        If IsNull(TBUsuarios!Codigo) = False And TBUsuarios!Codigo <> "" Then OperadorTexto = TBUsuarios!Codigo & "-" & Operador Else OperadorTexto = Operador
    End If
    TBUsuarios.Close
    
    Set TBProcessos = CreateObject("adodb.recordset")
    TBProcessos.Open "Select * from " & NomeTabelaAp & " where idfase = " & TBOS!IDProducao & " and Data <= '" & Format(Date, "Short Date") & "' and Tempoinicio < '" & Hora_apontamento & "' order by Data, Tempoinicio", Conexao, adOpenKeyset, adLockOptimistic
    If TBProcessos.EOF = False Then
        TBProcessos.MoveLast
        
        SetorNC = ""
        Set TBMaquinas = CreateObject("adodb.recordset")
        TBMaquinas.Open "Select Setor from CadMaquinas where Maquina = '" & TBProcessos!Maquina & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBMaquinas.EOF = False Then
            SetorNC = TBMaquinas!Setor
        End If
        TBMaquinas.Close
        
        Dim ParecerCQ As String 'OK
        If FunVerifNCRejeitado(TBOS!IDProducao) = True Then ParecerCQ = "Rejeitar" Else ParecerCQ = "Nada consta"
        
        If FunVerifNCPorDescricao(TBOS!IDProducao) = True Then
                    QTNC = IIf(GridRE.Cell(i, 2).Text <> "", GridRE.Cell(i, 2).Text, 0)
                    QTCD = IIf(GridRE.Cell(i, 3).Text <> "", GridRE.Cell(i, 3).Text, 0)
                    
                    If QTCD <> "0" Then
                    ParecerCQ = "Aprovado c/ desvio"
                    End If

                    CamposLoop = "(" & TBProcessos!IDProducao & "," & TBOS!Ordem & ", " & TBOS!IDProducao & ", " & Replace(QTNC, ",", ".") & ", " & Replace(TBOS!Quantidade, ",", ".") & ", '" & TBProcessos!Data & "', '" & Format(TBProcessos!TempoInicio, "hh:mm:ss") & "', '" & TBProcessos!Maquina & "', " & TBProcessos!Turno & ", '" & ParecerCQ & "', '" & OperadorTexto & "', '" & SetorNC & "', '" & GridRE.Cell(i, 1).Text & "', 1," & QTCD & ")"
                    Conexao.Execute "INSERT INTO CQ_NC_FABRICA (IDProducao, Ordem, OS, TTNC, LOTE, Data, Hora, Maquina, Turno, PARECERCQ, Operador, Setor, obsFab, Analizada,QTCD) VALUES " & CamposLoop
          
        Else
        
        If Individual = False Then
        

             Set TBAbrir = CreateObject("adodb.recordset")
                 TBAbrir.Open "Select * from CQ_NC_FABRICA", Conexao, adOpenKeyset, adLockOptimistic
                 TBAbrir.AddNew
                 TBAbrir!IDProducao = TBProcessos!IDProducao
                 TBAbrir!Ordem = TBOS!Ordem
                 TBAbrir!OS = TBOS!IDProducao
                 TBAbrir!TTNC = 1 'Replace(TNC, ",", ".")
                 TBAbrir!LOTE = Replace(TBOS!Quantidade, ",", ".")
                 TBAbrir!Data = TBProcessos!Data
                 TBAbrir!Hora = Format(TBProcessos!TempoInicio, "hh:mm:ss")
                 TBAbrir!Maquina = TBProcessos!Maquina
                 TBAbrir!Turno = TBProcessos!Turno
                 TBAbrir!ParecerCQ = ParecerCQ
                 TBAbrir!Operador = OperadorTexto
                 TBAbrir!Setor = SetorNC
                 TBAbrir.Update
                 Codigo = TBAbrir!Codigo
                 IDProducao = TBProcessos!IDProducao
                 TBAbrir.Close
    '            'CamposLoop = "(" & TBProcessos!IDProducao & ", " & TBOS!Ordem & ", " & TBOS!IDProducao & ", " & Replace(TNC, ",", ".") & ", " & Replace(TBOS!Quantidade, ",", ".") & " , '" & TBProcessos!Data & "', '" & Format(TBProcessos!TempoInicio, "hh:mm:ss") & "', '" & TBProcessos!Maquina & "', " & TBProcessos!Turno & ", '" & ParecerCQ & "', '" & OperadorTexto & "', '" & SetorNC & "', '" & NumeroSerie & "')"
     '           SQL = "INSERT INTO CQ_NC_FABRICA (IDProducao, Ordem, OS, TTNC, LOTE, Data, Hora, Maquina, Turno, PARECERCQ, Operador, SetorSetor, NumeroSerie) VALUES " & CamposLoop
     '           'Debug.Print SQL
     '           Conexao.Execute "INSERT INTO CQ_NC_FABRICA (IDProducao, Ordem, OS, TTNC, LOTE, Data, Hora, Maquina, Turno, PARECERCQ, Operador, Setor, NumeroSerie) VALUES " & CamposLoop
        End If
        End If
    End If
    TBProcessos.Close
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo tratar_erro
    
If FunVerifNCPorDescricao(Txt_OS) = False Then Call FunKeyAscii(KeyAscii)

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro
ProcAjustaGridRE

Calc_OK = True
Txt_OS = TBOS!IDProducao
txtOP = OF

'If Evento = 3 Then
If Processo_controlado = True Then
    ProcVerifQtdePCSOSAnt TBOS!Ordem, TBOS!Fase, TBOS!IDProducao, TBOS!Quantidade
    If qtdeliberada - TBOS!Totalprod > 0 Then
        txtTOK = qtdeliberada - TBOS!Totalprod
        txtTNC = 0
    End If
End If

If FunVerifNCPorDescricao(Txt_OS) = True Then
    With txtTNC
        .Locked = True
        .TabStop = False
    End With
    ProcCarregaListaNC
    Height = 10600
Else
    Height = 3600
End If


Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub



Private Sub ProcCarregaListaNC()
On Error GoTo tratar_erro
Dim L As Long

With GridRE
    
    L = 1
 Contador = 1
GridRE.Rows = 1
   
    Set TBLista = CreateObject("adodb.recordset")
    TBLista.Open "Select ID, Causa from CQ_NC_FABRICA_causa where DtValidacao IS NOT NULL order by Causa", Conexao, adOpenKeyset, adLockOptimistic
    If TBLista.EOF = False Then
        Contador = 1
        QtdeDescricaoNC = TBLista.RecordCount

        Do While TBLista.EOF = False


            GridRE.AddItem TBLista!Causa
            GridRE.Cell(Contador, 3).ForeColor = vbBlue
            GridRE.Cell(Contador, 2).ForeColor = vbRed
                 
        Contador = Contador + 1
        TBLista.MoveNext
        Loop
    End If
    TBLista.Close
End With
    
Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Public Sub ProcAjustaGridRE()
On Error GoTo tratar_erro

With GridRE

    .AllowUserPaste = cellTextOnly
    .AllowUserResizing = False
    .ExtendLastCol = True
    .BoldFixedCell = False
    .DisplayDateTimeMask = True
    .DisplayFocusRect = False
    .SelectionMode = cellSelectionFree

    .DrawMode = cellOwnerDraw
    
    .Appearance = Flat
    .ScrollBarStyle = Flat
    .FixedRowColStyle = Flat
    .Cell(0, 1).ForeColor = vbRed
    .Cell(0, 1).Text = "Não conformidade"

    .Cell(0, 2).ForeColor = vbRed
    .Cell(0, 2).Text = "Não conforme"
    
    .Cell(0, 3).ForeColor = vbBlue
    .Cell(0, 3).Text = "Condicional"
        
    .Column(1).CellType = cellTextBox
    .Column(1).Alignment = cellCenterCenter
    .Column(1).Locked = True
    
    .Column(2).CellType = cellTextBox
    .Column(2).Alignment = cellCenterCenter
    
    .Column(3).CellType = cellTextBox
    .Column(3).Alignment = cellCenterCenter
    
 
    .Column(0).Width = 10
    .Column(1).Width = 200
    .Column(2).Width = 80
    .Column(3).Width = 80
End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub GridRE_CellChange(ByVal Row As Long, ByVal Col As Long)
On Error GoTo tratar_erro

ProcCalculaTotalNC

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub GridRE_KeyPress(KeyAscii As Integer)
On Error GoTo tratar_erro

If KeyAscii = 46 Then KeyAscii = 44
    If KeyAscii = 44 And InStr(GridRE.ActiveCell.Text, ",") <> 0 Then
        KeyAscii = 0
        Exit Sub
    End If
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 44 Then
        KeyAscii = 0
    End If
Exit Sub

tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub




Private Sub txtQTCD_Change()
On Error GoTo tratar_erro
  
If IsNumeric(txtQTCD.Text) = True Then
TotalCondicional = txtQTCD
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub txtTNC_Change()
On Error GoTo tratar_erro
  
If IsNumeric(txtTNC.Text) = False Then Exit Sub
If txtTNC.Text <> "" Or txtTNC.Text > 0 Then
TotalNaoconforme = txtTNC
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub txtTNC_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus txtTNC
Calc_OK = False

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub


Private Sub txtTOK_Change()
On Error GoTo tratar_erro
  
If IsNumeric(txtTOK.Text) = False Then Exit Sub
If txtTOK.Text <> "" Or txtTOK.Text > 0 Then
TotalAprovado = txtTOK
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub txtTOK_GotFocus()
On Error GoTo tratar_erro
  
FunGotFocus txtTOK
Calc_OK = True

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Function FunVerificaEstoqueAP() As Boolean
On Error GoTo tratar_erro

'Verifica se existe quantidade em estoque para fazer a baixa, quando a ordem tem baixa de material automatico, se não tiver não deixa apontar
'=========================================================================================================================================================================================================================================================
If Varias_OS = True Then
    TextoFiltro = "Select OS.*, P.Desenho, P.Quant, P.Ordem from (OrdemServico OS INNER JOIN ProducaoFases_OS PFOS ON OS.ID_apontamento = PFOS.ID) INNER JOIN Producao P ON P.Ordem = OS.Ordem where PFOS.ID = " & frmProducao.Txt_ID_apontamento & " AND P.Retirar_estoque = 'True' order by OS.IDproducao"
Else
    TextoFiltro = "Select OS.*, P.Desenho, P.Quant, P.Ordem from OrdemServico OS INNER JOIN Producao P ON P.Ordem = OS.Ordem where OS.IDProducao = " & frmProducao.ListaOS.SelectedItem.ListSubItems(1) & " AND P.Retirar_estoque = 'True'"
End If

Set TBUN = CreateObject("adodb.recordset")
TBUN.Open TextoFiltro, Conexao, adOpenKeyset, adLockReadOnly
If TBUN.EOF = False Then
    Do While TBUN.EOF = False
    
        Set TBUN1 = CreateObject("adodb.recordset")
        TBUN1.Open "Select * from ordemservico where Ordem = " & TBUN!Ordem & " order by fase, retrabalho, IDproducao", Conexao, adOpenKeyset, adLockReadOnly
        If TBUN1.EOF = False Then
            TBUN1.MoveFirst
            If TBUN1!IDProducao = TBUN!IDProducao Then
                QtdeSaida = 0
                Set TBUN2 = CreateObject("adodb.recordset")
                TBUN2.Open "Select PM.Requisitado, ISNULL(QSEP.Saida, 0) as QtdeSaida, PM.CODIGO, P.IDcliente from ((Producao P INNER JOIN Producaomaterial PM ON PM.Ordem = P.Ordem) LEFT JOIN Qtde_saida_estoque_produto QSEP ON QSEP.Ordem = PM.Ordem and QSEP.Desenho = PM.Codigo) where PM.Ordem = " & TBUN!Ordem, Conexao, adOpenKeyset, adLockReadOnly
                If TBUN2.EOF = False Then
                    Do While TBUN2.EOF = False
                    
                        QtdeSaida = IIf(IsNull(TBUN2!QtdeSaida), 0, TBUN2!QtdeSaida)
                        Peso = TBUN2!Requisitado / TBUN!Quant
                        SaqueSaldo = (Peso * (TOK + TNC + TBUN!Totalprod))
                        SaqueSaldo = SaqueSaldo - QtdeSaida
                        
                        Set TBFI = CreateObject("adodb.recordset")
                        TBFI.Open "Select SUM(EC.Estoque_real) as TotalEstoque from estoque_controle EC INNER JOIN Estoque_produtos EP ON EP.IDestoque = EC.IDestoque where EC.Desenho = '" & TBUN2!Codigo & "' and EC.Estoque_real > 0 and EP.Liberado = 'SIM' and (EC.Consignacao = 'False' or EC.Consignacao = 'True' and EC.id_cliente = " & TBUN2!IDCliente & " and EC.Tipodest_NFcons = 'C' or EC.Consignacao = 'True' and EC.Tipodest_NFcons = 'F')", Conexao, adOpenKeyset, adLockReadOnly
                        If TBFI.EOF = False Then qtdeliberar = IIf(IsNull(TBFI!TotalEstoque), 0, TBFI!TotalEstoque)
                        TBFI.Close
                        
                        If SaqueSaldo > qtdeliberar Then
                            MsgBox ("Não é permitido apontar, pois a quantidade requisitada do item " & TBUN2!Codigo & " é maior que a disponivel em estoque."), vbExclamation
                            FunVerificaEstoqueAP = False
                            TBUN.Close
                            TBUN1.Close
                            TBUN2.Close
                            Exit Function
                        End If
                        
                        TBUN2.MoveNext
                    Loop
                End If
                TBUN2.Close
            End If
        End If
        TBUN1.Close
        TBUN.MoveNext
    Loop
End If
TBUN.Close
FunVerificaEstoqueAP = True

Exit Function
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Function

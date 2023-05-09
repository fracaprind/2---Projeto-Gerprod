VERSION 5.00
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmabrir_Ordem 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Gerprod  - Coletor de dados no chão de fábrica"
   ClientHeight    =   9045
   ClientLeft      =   60
   ClientTop       =   315
   ClientWidth     =   12645
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmabrir_Ordem.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Moveable        =   0   'False
   Picture         =   "frmabrir_Ordem.frx":0E42
   ScaleHeight     =   9045
   ScaleWidth      =   12645
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
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
      Left            =   3705
      Locked          =   -1  'True
      TabIndex        =   7
      TabStop         =   0   'False
      ToolTipText     =   "Descrição do posto de trabalho."
      Top             =   6075
      Width           =   6885
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
      Left            =   4320
      Locked          =   -1  'True
      TabIndex        =   9
      TabStop         =   0   'False
      ToolTipText     =   "Operador."
      Top             =   6990
      Width           =   3825
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
      Height          =   480
      IMEMode         =   3  'DISABLE
      Left            =   1815
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   8
      Top             =   6990
      Width           =   2475
   End
   Begin VB.ComboBox cmbPT 
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
      Left            =   1815
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   6
      ToolTipText     =   "Posto de trabalho."
      Top             =   6075
      Width           =   1875
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
      Left            =   8190
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   10
      ToolTipText     =   "Fase."
      Top             =   6990
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
      Left            =   9285
      Locked          =   -1  'True
      TabIndex        =   11
      TabStop         =   0   'False
      ToolTipText     =   "Número da OS."
      Top             =   6990
      Width           =   1305
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
      Left            =   6735
      Locked          =   -1  'True
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Quantidade."
      Top             =   2895
      Width           =   1905
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
      Left            =   1815
      Locked          =   -1  'True
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Descrição."
      Top             =   3765
      Width           =   8775
   End
   Begin VB.TextBox txtordem 
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
      Height          =   475
      IMEMode         =   3  'DISABLE
      Left            =   1815
      TabIndex        =   0
      ToolTipText     =   "Número da ordem."
      Top             =   2895
      Width           =   1845
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
      Left            =   3690
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Código interno."
      Top             =   2895
      Width           =   3015
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
      Left            =   8670
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      ToolTipText     =   "Prazo final."
      Top             =   2895
      Width           =   1920
   End
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
      Left            =   1815
      Locked          =   -1  'True
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Cliente."
      Top             =   4830
      Width           =   8775
   End
   Begin DrawSuite2022.USButton Cmd_enter 
      Height          =   600
      Left            =   2850
      TabIndex        =   12
      Top             =   1545
      Width           =   3150
      _ExtentX        =   5556
      _ExtentY        =   1058
      BorderColor     =   5263559
      BorderColorDisabled=   13160660
      BorderColorDown =   4013465
      BorderColorOver =   4408288
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
      ForeColor       =   16777215
      ForeColorDown   =   16777215
      ForeColorOver   =   16777215
      GradientColor1  =   5263559
      GradientColor2  =   5263559
      GradientColor3  =   5263559
      GradientColor4  =   5263559
      GradientColorDisabled1=   13160660
      GradientColorDisabled2=   13160660
      GradientColorDisabled3=   13160660
      GradientColorDisabled4=   13160660
      GradientColorDown1=   4013465
      GradientColorDown2=   4013465
      GradientColorDown3=   4013465
      GradientColorDown4=   4013465
      GradientColorOver1=   4408288
      GradientColorOver2=   4408288
      GradientColorOver3=   4408288
      GradientColorOver4=   4408288
      PicAlign        =   8
      PicSize         =   4
      PicSizeH        =   48
      PicSizeW        =   48
      ShowFocusRect   =   0   'False
      ShowFocusRectDown=   0   'False
      Theme           =   4
   End
   Begin DrawSuite2022.USButton Cmd_esc 
      Height          =   600
      Left            =   6060
      TabIndex        =   13
      Top             =   1545
      Width           =   3150
      _ExtentX        =   5556
      _ExtentY        =   1058
      BorderColor     =   5263559
      BorderColorDisabled=   13160660
      BorderColorDown =   4013465
      BorderColorOver =   4408288
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
      ForeColor       =   16777215
      ForeColorDown   =   16777215
      ForeColorOver   =   16777215
      GradientColor1  =   5263559
      GradientColor2  =   5263559
      GradientColor3  =   5263559
      GradientColor4  =   5263559
      GradientColorDisabled1=   13160660
      GradientColorDisabled2=   13160660
      GradientColorDisabled3=   13160660
      GradientColorDisabled4=   13160660
      GradientColorDown1=   4013465
      GradientColorDown2=   4013465
      GradientColorDown3=   4013465
      GradientColorDown4=   4013465
      GradientColorOver1=   4408288
      GradientColorOver2=   4408288
      GradientColorOver3=   4408288
      GradientColorOver4=   4408288
      PicAlign        =   8
      PicSize         =   4
      PicSizeH        =   48
      PicSizeW        =   48
      ShowFocusRect   =   0   'False
      ShowFocusRectDown=   0   'False
      Theme           =   4
   End
   Begin ActiveResizeCtl.ActiveResize ActiveResize1 
      Left            =   0
      Top             =   0
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
      TabIndex        =   28
      Top             =   8505
      Width           =   5400
   End
   Begin VB.Label lblVersaoatual 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2.5"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11000
      TabIndex        =   27
      Top             =   400
      Width           =   450
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
      Left            =   1935
      TabIndex        =   26
      Top             =   6615
      Visible         =   0   'False
      Width           =   2220
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
      Left            =   2475
      TabIndex        =   25
      Top             =   6615
      Width           =   960
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
      Left            =   8355
      TabIndex        =   24
      Top             =   6615
      Width           =   720
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
      Left            =   9720
      TabIndex        =   23
      Top             =   6615
      Width           =   435
   End
   Begin VB.Label Label9 
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
      Left            =   5190
      TabIndex        =   22
      Top             =   5715
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
      Left            =   2070
      TabIndex        =   21
      Top             =   5715
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
      Left            =   5640
      TabIndex        =   20
      Top             =   6615
      Width           =   1185
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
      Left            =   1980
      TabIndex        =   19
      Top             =   2535
      Width           =   1500
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
      Left            =   4275
      TabIndex        =   18
      Top             =   2535
      Width           =   1830
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
      Left            =   6960
      TabIndex        =   17
      Top             =   2535
      Width           =   1455
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
      Left            =   8985
      TabIndex        =   16
      Top             =   2535
      Width           =   1290
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
      Left            =   5595
      TabIndex        =   15
      Top             =   3405
      Width           =   1200
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
      Left            =   5775
      TabIndex        =   14
      Top             =   4485
      Width           =   840
   End
   Begin DrawSuite2022.USAlphaImage USAlphaImage1 
      Height          =   9060
      Left            =   0
      Top             =   0
      Width           =   12645
      _ExtentX        =   22304
      _ExtentY        =   15981
      Image           =   "frmabrir_Ordem.frx":7E6B4
   End
   Begin VB.Label lbldireitos 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Este programa é protegido por leis de direitos autorais (Copyright) e tratados internacionais."
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
      Left            =   2850
      TabIndex        =   29
      Top             =   8730
      Width           =   6840
   End
End
Attribute VB_Name = "frmabrir_Ordem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

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

Private Sub Form_Load()
On Error GoTo tratar_erro

'If xPixels = 800 Then Img_800x600.Visible = True Else Img_1024x768.Visible = True
If Codigo_Barras = True Then
    Caption = "Gerprod - Coletor de dados no chão de fábrica - Código de barras - Nome do banco de dados: " & Nome_banco & " | Licenças Gerprod: " & Qtlicencas_gerprod & IIf(TemInternet = False, " | Sem internet", "")
'    Image1.Visible = True
    lblsenha.Visible = False
    lblcodigosenha.Visible = True
    txtSenha.ToolTipText = "Código/Senha do usuário"
Else
    Caption = "Gerprod - Coletor de dados no chão de fábrica - Nome do banco de dados: " & Nome_banco & " | Licenças Gerprod: " & Qtlicencas_gerprod & IIf(TemInternet = False, " | Sem internet", "")
 '   Image1.Visible = False
    lblsenha.Visible = True
    lblcodigosenha.Visible = False
    txtSenha.ToolTipText = "Senha do usuário"
End If

cmbPT.Clear
Set TBMaquinas = CreateObject("adodb.recordset")
TBMaquinas.Open "Select maquina from cadmaquinas where Bloqueado = 'False' order by maquina", Conexao, adOpenKeyset, adLockOptimistic
If TBMaquinas.EOF = False Then
    cmbPT.AddItem ""
    Do While TBMaquinas.EOF = False
        cmbPT.AddItem TBMaquinas!Maquina
        TBMaquinas.MoveNext
    Loop
End If
TBMaquinas.Close

lblVersaoatual.Caption = "v" & App.Major & "." & App.Minor & "." & App.Revision
lblano.Caption = "Copyright 1999 - " & Year(Date) & " Caprind Sistemas ®. Todos os direitos reservados."

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

Private Sub txtfase_Click()
On Error GoTo tratar_erro

If txtordem.text <> "" Then
    txtOS.text = ""
    Set TBProducao = CreateObject("adodb.recordset")
    TBProducao.Open "Select * from ordemservico where Ordem = " & txtordem.text & " and maquina = '" & cmbPT & "' and fase = " & txtfase.text & " and pronto = 'NÃO' order by idproducao", Conexao, adOpenKeyset, adLockOptimistic
    If TBProducao.EOF = False Then
        txtOS.text = TBProducao!IDProducao
        OSControlada = TBProducao!OSControlada
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

ProcLimpaCampos
If txtordem.text <> "" Then
    VerifNumero = txtordem
    ProcVerificaNumero
    If VerifNumero = False Then
        txtordem = ""
        txtordem.SetFocus
        Exit Sub
    End If
    Set TBMaquinas = CreateObject("adodb.recordset")
    TBMaquinas.Open "Select * from producao where Ordem = " & txtordem.text & " and Status <> 'Cancelada' and DtValidacao IS NOT NULL and DtValidacao_Custo IS NULL", Conexao, adOpenKeyset, adLockOptimistic
    If TBMaquinas.EOF = False Then
        txtitem.text = TBMaquinas!Desenho
        txtquant.text = TBMaquinas!Quant
        mskprazofinal.text = Format(TBMaquinas!prazoentrega, "dd/mm/yy")
        txtcliente.text = IIf(IsNull(TBMaquinas!cliente) = False, TBMaquinas!cliente, "")
        txtdescricao.text = TBMaquinas!Produto
        cmbPT.Enabled = True
    Else
        cmbPT.Enabled = False
        txtSenha.Enabled = False
        txtfase.Enabled = False
    End If
    TBMaquinas.Close
Else
    cmbPT.Enabled = False
    txtSenha.Enabled = False
    txtfase.Enabled = False
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

Private Sub ProcLimpaCampos()
On Error GoTo tratar_erro

txtitem.text = ""
txtquant.text = ""
mskprazofinal.text = "__/__/__"
txtcliente.text = ""
txtdescricao.text = ""
cmbPT.ListIndex = -1
txtSenha.text = ""
txtusuario.text = ""
txtfase.Clear
txtOS.text = ""
txtcliente.text = ""

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub cmbPT_Click()
On Error GoTo tratar_erro

If txtordem.text <> "" Then
    If IsNumeric(txtordem) = True Then
        txtfase.Clear
        txtOS.text = ""
        Set TBProducao = CreateObject("adodb.recordset")
        TBProducao.Open "Select * from ordemservico where Ordem = " & txtordem.text & " and maquina = '" & cmbPT & "' and pronto = 'NÃO' order by idproducao", Conexao, adOpenKeyset, adLockOptimistic
        If TBProducao.EOF = False Then
            Do While TBProducao.EOF = False
                txtfase.AddItem TBProducao!Fase
                TBProducao.MoveNext
            Loop
        Else
            Set TBProducao = CreateObject("adodb.recordset")
            TBProducao.Open "Select * from ordemservico where Ordem = " & txtordem.text & " order by idproducao", Conexao, adOpenKeyset, adLockOptimistic
            If TBProducao.EOF = False Then
                TBProducao.MoveLast
                txtfase.AddItem TBProducao!Fase + 10
                txtOS.text = ""
            Else
                txtfase.AddItem 10
            End If
        End If
        TBProducao.Close
        If cmbPT.text <> "" Then
            txtSenha.Enabled = True
        Else
            txtSenha.Enabled = False
            txtSenha = ""
        End If
    End If
End If
Set TBMaquinas = CreateObject("adodb.recordset")
TBMaquinas.Open "Select Descricao from CadMaquinas where Maquina = '" & cmbPT & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBMaquinas.EOF = False Then
    txtDescricaoPT = TBMaquinas!Descricao
End If
TBMaquinas.Close

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcVerificaDisponMaquina()
On Error GoTo tratar_erro

Set TBMaquinas = CreateObject("adodb.recordset")
If txtOS.text <> "" Then TextoFiltro = "and Ordemservico_maq_utilizadas.OS <> " & txtOS.text Else TextoFiltro = ""
TBMaquinas.Open "Select * from CadMaquinas INNER JOIN Ordemservico_maq_utilizadas on CadMaquinas.Maquina = Ordemservico_maq_utilizadas.Maquina where CadMaquinas.maquina = '" & Maquina & "' and CadMaquinas.Liberada = 'false' " & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
If TBMaquinas.EOF = False Then
    MsgBox ("Não é permitido utilizar essa máquina, pois a mesma já está sendo utilizada na OS: " & TBMaquinas!OS & "."), vbExclamation
    cmbPT.ListIndex = -1
    txtDescricaoPT = ""
    cmbPT.SetFocus
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

If Atualizando = True Then Exit Sub

Select Case KeyCode
    Case vbKeyReturn:
        ProcVerifUsuario
        Acao = "acessar"
        If txtordem = "" Or cmbPT = "" Or txtusuario = "" Or txtfase = "" Then Exit Sub
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
        
        ProcVerifAPCodigo False, txtordem
        frmProducao.ProcAbrirNovo
        Unload Me
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
    
If Codigo_Barras = False Then Call FunKeyAscii(KeyAscii)

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub txtSenha_Change()
On Error GoTo tratar_erro

txtusuario.text = ""
txtfase.Enabled = False

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
        txtfase.Enabled = True
    End If
    TBUsuarios.Close
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub



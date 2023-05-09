VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Begin VB.Form frmabrir_OS 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H00C0C0C0&
   BorderStyle     =   0  'None
   Caption         =   "GERPROD | Coletor de dados no chão de fábrica"
   ClientHeight    =   10035
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15480
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmabrir_OS.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   10035
   ScaleWidth      =   15480
   WindowState     =   2  'Maximized
   Begin DrawSuite2022.USCheckBox chkQT_Final 
      Height          =   405
      Left            =   9330
      TabIndex        =   41
      Top             =   2820
      Width           =   5805
      _ExtentX        =   10239
      _ExtentY        =   714
      BackColor       =   16382457
      BackStyle       =   0
      Caption         =   "Apontar quantidades produzidas no fechamento da  O.S"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   128
      ShowFocusRect   =   -1  'True
      Theme           =   1
   End
   Begin VB.TextBox Txt_caminho 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   450
      IMEMode         =   3  'DISABLE
      Left            =   450
      Locked          =   -1  'True
      TabIndex        =   39
      TabStop         =   0   'False
      ToolTipText     =   "Caminho do arquivo."
      Top             =   8400
      Width           =   14505
   End
   Begin VB.ComboBox cmbmaquina 
      BackColor       =   &H00FFFFFF&
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   465
      Left            =   2010
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   2
      ToolTipText     =   "Código do posto de trabalho."
      Top             =   6390
      Width           =   2145
   End
   Begin VB.TextBox txtfase 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   450
      IMEMode         =   3  'DISABLE
      Left            =   450
      Locked          =   -1  'True
      TabIndex        =   38
      TabStop         =   0   'False
      ToolTipText     =   "Fase."
      Top             =   6390
      Width           =   1545
   End
   Begin VB.TextBox txtdescmaq 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
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
      Left            =   4170
      Locked          =   -1  'True
      TabIndex        =   37
      TabStop         =   0   'False
      ToolTipText     =   "Descrição do posto de trabalho."
      Top             =   6390
      Width           =   10785
   End
   Begin RichTextLib.RichTextBox Txt_instrucao 
      Height          =   735
      Left            =   450
      TabIndex        =   40
      TabStop         =   0   'False
      ToolTipText     =   "Instrução do serviço."
      Top             =   7290
      Width           =   14505
      _ExtentX        =   25585
      _ExtentY        =   1296
      _Version        =   393217
      BackColor       =   16777215
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"frmabrir_OS.frx":0E42
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox TXTOF 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   450
      IMEMode         =   3  'DISABLE
      Left            =   420
      Locked          =   -1  'True
      TabIndex        =   36
      TabStop         =   0   'False
      ToolTipText     =   "Número da ordem."
      Top             =   4830
      Width           =   1305
   End
   Begin VB.TextBox txtprazo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   450
      IMEMode         =   3  'DISABLE
      Left            =   4830
      Locked          =   -1  'True
      TabIndex        =   35
      TabStop         =   0   'False
      ToolTipText     =   "Prazo final."
      Top             =   4830
      Width           =   1335
   End
   Begin VB.TextBox txtquant 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   450
      IMEMode         =   3  'DISABLE
      Left            =   3660
      Locked          =   -1  'True
      TabIndex        =   34
      TabStop         =   0   'False
      ToolTipText     =   "Quantidade."
      Top             =   4830
      Width           =   1155
   End
   Begin VB.TextBox txtdesenho 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   450
      IMEMode         =   3  'DISABLE
      Left            =   1740
      Locked          =   -1  'True
      TabIndex        =   33
      TabStop         =   0   'False
      ToolTipText     =   "Código interno."
      Top             =   4830
      Width           =   1905
   End
   Begin VB.TextBox txtdescricao 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   450
      IMEMode         =   3  'DISABLE
      Left            =   6210
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   32
      TabStop         =   0   'False
      ToolTipText     =   "Descrição."
      Top             =   4830
      Width           =   8715
   End
   Begin VB.TextBox txtusuario 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   450
      IMEMode         =   3  'DISABLE
      Left            =   3980
      Locked          =   -1  'True
      TabIndex        =   29
      TabStop         =   0   'False
      ToolTipText     =   "Operador."
      Top             =   3420
      Width           =   5055
   End
   Begin DrawSuite2022.USCheckBox chkIndividual 
      Height          =   405
      Left            =   9330
      TabIndex        =   30
      Top             =   3390
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   714
      BackColor       =   16382457
      BackStyle       =   0
      Caption         =   "Ordem com rastreabilidade individual"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   128
      ShowFocusRect   =   -1  'True
      Theme           =   1
   End
   Begin DrawSuite2022.USCheckBox chkOrdemControlada 
      Height          =   405
      Left            =   9330
      TabIndex        =   31
      Top             =   3105
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   714
      BackColor       =   16382457
      BackStyle       =   0
      Caption         =   "Ordem com quantidades controladas"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   128
      ShowFocusRect   =   -1  'True
      Theme           =   1
   End
   Begin VB.TextBox txtOS 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   450
      IMEMode         =   3  'DISABLE
      Left            =   420
      TabIndex        =   0
      ToolTipText     =   "Numero da ordem de serviço"
      Top             =   3420
      Width           =   1755
   End
   Begin VB.TextBox txtSenha 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      ForeColor       =   &H00000000&
      Height          =   450
      IMEMode         =   3  'DISABLE
      Left            =   2190
      MaxLength       =   20
      PasswordChar    =   "*"
      TabIndex        =   1
      ToolTipText     =   "Número para apontamento."
      Top             =   3420
      Width           =   1770
   End
   Begin VB.Frame Frame1 
      Caption         =   "Informe os dados para apontamento de produção"
      Height          =   1485
      Left            =   270
      TabIndex        =   24
      Top             =   2610
      Width           =   15015
      Begin DrawSuite2022.USCheckBox chkOSIndividual 
         Height          =   405
         Left            =   9060
         TabIndex        =   42
         Top             =   1080
         Width           =   4905
         _ExtentX        =   8652
         _ExtentY        =   714
         BackColor       =   16382457
         BackStyle       =   0
         Caption         =   "Ordem de serviço com rastreabilidade individual"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   128
         ShowFocusRect   =   -1  'True
         Theme           =   1
      End
      Begin VB.Label lblcodigosenha 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   2400
         TabIndex        =   28
         Top             =   450
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Operador"
         Height          =   345
         Left            =   5640
         TabIndex        =   27
         Top             =   450
         Width           =   1185
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N° OS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   495
         TabIndex        =   26
         ToolTipText     =   "N° da ordem de serviço"
         Top             =   450
         Width           =   870
      End
      Begin VB.Label lblsenha 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Senha"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   375
         Left            =   2400
         TabIndex        =   25
         Top             =   450
         Width           =   870
      End
   End
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   22
      Top             =   9630
      Width           =   15480
      _ExtentX        =   27305
      _ExtentY        =   714
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   555
      Left            =   0
      TabIndex        =   21
      Top             =   0
      Width           =   15480
      _ExtentX        =   27305
      _ExtentY        =   979
      CaptionDelimiter=   "|"
      CaptionOnCenter =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Icon            =   "frmabrir_OS.frx":0EBD
      ShowControlBox  =   0   'False
   End
   Begin DrawSuite2022.USGroupBox USGroupBox3 
      Height          =   3495
      Left            =   270
      TabIndex        =   15
      Top             =   5610
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   6165
      Appearance      =   0
      Caption         =   "Informações do processo de produção"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      GradientColor1  =   14737632
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Fase"
         Height          =   345
         Left            =   705
         TabIndex        =   20
         Top             =   390
         Width           =   570
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Caminho do arquivo"
         Height          =   345
         Left            =   6127
         TabIndex        =   19
         Top             =   2430
         Width           =   2610
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição do posto de trabalho"
         Height          =   345
         Left            =   7222
         TabIndex        =   18
         Top             =   435
         Width           =   3900
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Posto trabalho"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   15.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1807
         TabIndex        =   17
         Top             =   390
         Width           =   2010
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Instrução do serviço"
         Height          =   345
         Left            =   6165
         TabIndex        =   16
         Top             =   1335
         Width           =   2535
      End
   End
   Begin DrawSuite2022.USGroupBox USGroupBox2 
      Height          =   1485
      Left            =   270
      TabIndex        =   9
      Top             =   4110
      Width           =   15015
      _ExtentX        =   26485
      _ExtentY        =   2619
      Appearance      =   0
      Caption         =   "Informações da ordem de produção"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      GradientColor1  =   14737632
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição"
         Height          =   345
         Left            =   9727
         TabIndex        =   14
         Top             =   360
         Width           =   1200
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Prazo final"
         Height          =   345
         Left            =   4515
         TabIndex        =   13
         Top             =   360
         Width           =   1290
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Quant."
         Height          =   345
         Left            =   3465
         TabIndex        =   12
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cód. interno"
         Height          =   345
         Left            =   1650
         TabIndex        =   11
         Top             =   360
         Width           =   1545
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ordem"
         Height          =   345
         Left            =   375
         TabIndex        =   10
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.TextBox Txt_ID_apontamento 
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
      Left            =   1305
      MaxLength       =   20
      TabIndex        =   6
      Top             =   8295
      Visible         =   0   'False
      Width           =   1035
   End
   Begin DrawSuite2022.USButton Cmd_enter 
      Height          =   660
      Left            =   8010
      TabIndex        =   3
      Top             =   1815
      Width           =   2730
      _ExtentX        =   4815
      _ExtentY        =   1164
      Caption         =   "(Enter) Aceitar dados"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColorDown =   15048022
      BorderColorOver =   15381630
      PicAlign        =   8
      PicSize         =   4
      PicSizeH        =   48
      PicSizeW        =   48
   End
   Begin DrawSuite2022.USButton Cmd_esc 
      Height          =   660
      Left            =   13590
      TabIndex        =   5
      Top             =   1815
      Width           =   1710
      _ExtentX        =   3016
      _ExtentY        =   1164
      Caption         =   "(Esc) Voltar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColorDown =   15048022
      BorderColorOver =   15381630
      PicAlign        =   8
      PicSize         =   4
      PicSizeH        =   48
      PicSizeW        =   48
   End
   Begin DrawSuite2022.USButton Cmd_visualizar 
      Height          =   660
      Left            =   10800
      TabIndex        =   4
      Top             =   1815
      Width           =   2730
      _ExtentX        =   4815
      _ExtentY        =   1164
      Caption         =   "(F2) Visualizar arquivo"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColorDown =   15048022
      BorderColorOver =   15381630
      PicAlign        =   8
      PicSize         =   4
      PicSizeH        =   48
      PicSizeW        =   48
   End
   Begin ActiveResizeCtl.ActiveResize ActiveResize1 
      Left            =   0
      Top             =   420
      _ExtentX        =   847
      _ExtentY        =   847
      Resolution      =   99
      ScreenHeight    =   1080
      ScreenWidth     =   2560
      ScreenHeightDT  =   1080
      ScreenWidthDT   =   1920
      AutoCenterForm  =   -1  'True
      ApplicationName =   "Active Resize Control Professional"
      FormHeightDT    =   10035
      FormWidthDT     =   15480
      FormScaleHeightDT=   10035
      FormScaleWidthDT=   15480
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin DrawSuite2022.USAlphaImage USAlphaImage2 
      Height          =   990
      Left            =   2430
      TabIndex        =   43
      Top             =   1140
      Width           =   5250
      _ExtentX        =   9260
      _ExtentY        =   1746
      Image           =   "frmabrir_OS.frx":1D0F
   End
   Begin VB.Label lbldireitos 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmabrir_OS.frx":4F10
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
      Height          =   195
      Left            =   60
      TabIndex        =   23
      Top             =   9390
      Width           =   15360
   End
   Begin DrawSuite2022.USAlphaImage USAlphaImage1 
      Height          =   1620
      Left            =   510
      TabIndex        =   44
      Top             =   840
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   2858
      Image           =   "frmabrir_OS.frx":4FA1
      Props           =   5
   End
   Begin VB.Label lblano 
      Alignment       =   2  'Center
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
      Height          =   195
      Left            =   45
      TabIndex        =   8
      Top             =   9165
      Width           =   15390
   End
   Begin VB.Label lblVersaoatual 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "2.5"
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
      Left            =   13920
      TabIndex        =   7
      Top             =   720
      Width           =   495
   End
End
Attribute VB_Name = "frmabrir_OS"
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
    If Codigo_Barras = True Then txtSenha.SetFocus
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

Private Sub TXTOF_Change()
On Error GoTo tratar_erro

If IsNumeric(TXTOF.Text) Then
    Set TBAbrir = CreateObject("adodb.recordset")
    StrSQL = "Select individual from producao where ordem = '" & TXTOF.Text & "'"
    
    TBAbrir.Open StrSQL, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            If TBAbrir!Individual = True Then
                OrdemRastreavel = True
            Else
                OrdemRastreavel = False
            End If
        End If
    TBAbrir.Close
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub txtOS_Change()
On Error GoTo tratar_erro

ProcLimpaCampos


Select Case Codigo_Barras
        Case True: GoTo CB
        Case False: GoTo Inicio
    End Select

CB:
    If Len(txtos.Text) = 6 Then
Inicio:
If txtos.Text <> "" Then
    

    VerifNumero = txtos
    ProcVerificaNumero
    If VerifNumero = False Then
        txtos = ""
        txtos.SetFocus
        Exit Sub
    End If
    
    Set TBMaquinas = CreateObject("adodb.recordset")
    TBMaquinas.Open "Select OS.*,OS.Rastreavel as OSIndividual, P.ID_empresa, P.Desenho, P.Produto,P.Individual as OrdemIndividual FROM ordemservico OS INNER JOIN producao P ON OS.ordem = P.ordem where OS.idproducao = " & txtos.Text & " and OS.pronto = 'NÃO' and P.status <> 'Cancelada' and P.DtValidacao IS NOT NULL and P.DtValidacao_Custo IS NULL", Conexao, adOpenKeyset, adLockOptimistic
    

    If TBMaquinas.EOF = False Then
    
    Set TBAbrir = CreateObject("adodb.recordset")
    SQL = "Select CNPJ,Razao from Empresa Where codigo = " & TBMaquinas!ID_empresa
    TBAbrir.Open SQL, Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
    CNPJEmpresa = TBAbrir!CNPJ
'    lblEmpresa.Caption = "EMPRESA : " & TBAbrir!Razao
    End If
    TBAbrir.Close
    
        Txt_ID_apontamento = IIf(IsNull(TBMaquinas!ID_apontamento), "", TBMaquinas!ID_apontamento)
        TXTOF.Text = IIf(IsNull(TBMaquinas!Ordem), "", TBMaquinas!Ordem)
        OF = TXTOF.Text
        Ordem = TXTOF.Text
        OS = txtos.Text
        
        txtfase.Text = IIf(IsNull(TBMaquinas!Fase), "", TBMaquinas!Fase)
        Txt_instrucao.TextRTF = IIf(IsNull(TBMaquinas!descFase), "", TBMaquinas!descFase)
        txtdescricao.Text = IIf(IsNull(TBMaquinas!Produto), "", TBMaquinas!Produto)
        txtdesenho.Text = IIf(IsNull(TBMaquinas!Desenho), "", TBMaquinas!Desenho)
        txtprazo.Text = IIf(IsNull(TBMaquinas!Prazofinal), "", TBMaquinas!Prazofinal)
        If TBMaquinas!ordemIndividual <> "" Then
        
            If TBMaquinas!OSIndividual = True Then
                Individual = True
                chkOSIndividual.Value = Checked
            Else
                Individual = False
                chkOSIndividual.Value = Unchecked
            End If
            
            If TBMaquinas!ordemIndividual = True Then
                'Individual = False
                chkIndividual = Checked
            Else
                'Individual = True
                chkIndividual.Value = Unchecked
            End If
            
            
            If TBMaquinas!Processo_controlado = 0 Then
                Processo_controlado = False
                chkOrdemControlada.Value = Unchecked
            Else
                Processo_controlado = True
                chkOrdemControlada.Value = Checked
            End If
            
             If TBMaquinas!QT_Final = 0 Then
                QT_Final = False
                chkQT_Final.Value = Unchecked
            Else
                QT_Final = True
                chkQT_Final.Value = Checked
            End If
           
        
        End If
        
        'Verifica qtde. de pçs apontadas na OS anterior
        If TBMaquinas!Processo_controlado = True Then
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select * from ordemservico where Ordem = " & TBMaquinas!Ordem & " and Retrabalho = 'False' and Fase < '" & TBMaquinas!Fase & "' order by Fase desc", Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                txtquant.Text = IIf(IsNull(TBAbrir!QTOK), 0, TBAbrir!QTOK) + FunVerifQtdeRetrabalhoFase(TBMaquinas!Ordem, TBAbrir!Fase)
            Else
                txtquant.Text = TBMaquinas!Quantidade
                LOTE = txtquant.Text
            End If
            TBAbrir.Close
        Else
            Set TBAbrir = CreateObject("adodb.recordset")
            TBAbrir.Open "Select Sum(QTNC) as TTNC from ordemservico where Ordem = " & TBMaquinas!Ordem & " and Idproducao <> " & TBMaquinas!IDProducao, Conexao, adOpenKeyset, adLockOptimistic
            If TBAbrir.EOF = False Then
                QtdeRefugo = IIf(IsNull(TBAbrir!TTNC), 0, TBAbrir!TTNC)
            End If
            TBAbrir.Close
            txtquant.Text = TBMaquinas!Quantidade - IIf(TBMaquinas!Retrabalho = False, QtdeRefugo, 0)
        End If
        
        '========================================
        ' Variaveis de controle
        '========================================
        OSControlada = TBMaquinas!OSControlada
        Processo_controlado = TBMaquinas!Processo_controlado
        QT_Final = IIf(IsNull(TBMaquinas!QT_Final), 0, TBMaquinas!QT_Final)
        '========================================
        'Versao do processo
        '========================================
        Versao_processo = "A"
        Txt_caminho = ""
        Set TBFases = CreateObject("adodb.recordset")
        TBFases.Open "Select Versao, Caminho from Fases where IDFase = " & TBMaquinas!IDFase, Conexao, adOpenKeyset, adLockOptimistic
        If TBFases.EOF = False Then
            Versao_processo = IIf(IsNull(TBFases!Versao), "A", TBFases!Versao)
            Txt_caminho = IIf(IsNull(TBFases!Caminho), "", TBFases!Caminho)
        End If
        TBFases.Close
        '========================================

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
                    TBCFOP.Open "Select Grupo from CadMaquinas where Maquina = '" & TBMaquinas!Maquina & "' and bloqueado = 'false'", Conexao, adOpenKeyset, adLockOptimistic
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
                            TBAbrir.Open "Select Maquina from CadMaquinas_Monitor where Maquina = '" & TBOrdem!Maquina & "' and OS = " & txtos, Conexao, adOpenKeyset, adLockOptimistic
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
                TBOrdem.Open "Select Maquina from Ordemservico_maq_utilizadas where OS = " & txtos & " order by ID", Conexao, adOpenKeyset, adLockOptimistic
                If TBOrdem.EOF = False Then
                    TBOrdem.MoveLast
                    
                    Set TBAbrir = CreateObject("adodb.recordset")
                    TBAbrir.Open "Select CadMaquinas.Maquina from CadMaquinas INNER JOIN CadMaquinas_Monitor ON CadMaquinas.Maquina = CadMaquinas_Monitor.Maquina where CadMaquinas.Maquina = '" & TBOrdem!Maquina & "' and CadMaquinas.Liberada = 'False' and CadMaquinas.Bloqueado = 'False' and CadMaquinas_Monitor.OS <> " & txtos, Conexao, adOpenKeyset, adLockOptimistic
                    If TBAbrir.EOF = False Then
                        .AddItem TBOrdem!Maquina
                    End If
                    TBAbrir.Close
                    
                    .Text = TBOrdem!Maquina
                Else
                    'Verifica se a máquina prevista esta sendo utilizada pela OS a ser apontada
                    Set TBAbrir = CreateObject("adodb.recordset")
                    TBAbrir.Open "Select CadMaquinas.Maquina from CadMaquinas INNER JOIN CadMaquinas_Monitor ON CadMaquinas.Maquina = CadMaquinas_Monitor.Maquina where CadMaquinas.Maquina = '" & TBMaquinas!Maquina & "' and CadMaquinas.Liberada = 'False' and CadMaquinas.Bloqueado = 'False' and CadMaquinas_Monitor.OS = " & txtos, Conexao, adOpenKeyset, adLockOptimistic
                    If TBAbrir.EOF = False Then
                        .Text = TBMaquinas!Maquina
                    Else
                        'Verifica se a máquina prevista esta liberada
                        Set TBAbrir = CreateObject("adodb.recordset")
                        TBAbrir.Open "Select * from CadMaquinas where Maquina = '" & TBMaquinas!Maquina & "' and Bloqueado = 'False' and Liberada = 'True'", Conexao, adOpenKeyset, adLockOptimistic
                        If TBAbrir.EOF = False Then
                            .Text = TBMaquinas!Maquina
                        End If
                    End If
                    TBAbrir.Close
                End If
                TBOrdem.Close
            End If
        End With
        
        If Codigo_Barras = True Then
           If cmbmaquina = "" Then cmbmaquina.SetFocus Else txtSenha.SetFocus
        End If
   
    Else
        cmbmaquina.Enabled = False
        txtSenha.Enabled = False
    End If
   
    TBMaquinas.Close
End If
End If
LOTE = IIf(txtquant.Text <> "", txtquant, 0)
 
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

TXTOF.Text = ""
txtdesenho.Text = ""
txtquant.Text = ""
txtprazo.Text = ""
txtdescricao.Text = ""
txtfase.Text = ""
cmbmaquina.Clear
txtdescmaq.Text = ""
Txt_instrucao.Text = ""
txtSenha.Text = ""
txtusuario.Text = ""
Txt_caminho = ""

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub txtOS_LostFocus()
On Error GoTo tratar_erro

If txtos = "" Then Exit Sub
If Varias_OS = False Then
    Set TBMaquinas = CreateObject("adodb.recordset")
    TBMaquinas.Open "Select ordemservico.* FROM ordemservico INNER JOIN producao ON ordemservico.ordem = producao.ordem where ordemservico.idproducao = " & txtos.Text & " and ordemservico.pronto = 'NÃO' and producao.status <> 'Cancelada'", Conexao, adOpenKeyset, adLockOptimistic
    If TBMaquinas.EOF = False Then
        If IsNull(TBMaquinas!ID_apontamento) = False And TBMaquinas!ID_apontamento <> "" Then
            MsgBox ("Só é permitido apontar esta OS pelas opções: " & vbCrLf & " F9 - ABRIR VÁRIAS OS's " & vbCrLf & " F7 - ABRIR PLANO " & vbCrLf & " F11 - ABRIR VÁRIAS OS's " & vbCrLf & " F12 - ABRIR PLANO"), vbExclamation
            txtos = ""
            txtos.SetFocus
            ProcLimpaCampos
            cmbmaquina.Enabled = False
            txtSenha.Enabled = False
        End If
    End If
    TBMaquinas.Close
End If

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

Private Sub Cmd_visualizar_Click()
On Error GoTo tratar_erro

If Txt_caminho <> "" Then ProcAbrirArquivo Txt_caminho

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
        If txtos = "" Or cmbmaquina = "" Or txtusuario = "" Then Exit Sub
        
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
        
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select E.Desbloquear_prim_apont_OS_proc_controlado, E.Bloquear_apontamento_simultaneo from (Empresa E INNER JOIN Producao P ON P.ID_empresa = E.Codigo) INNER JOIN OrdemServico OS ON OS.Ordem = P.Ordem where OS.IDProducao = " & txtos, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            If Varias_OS = True And TBAbrir!Bloquear_apontamento_simultaneo = True Then
                MsgBox ("O recurso de ABRIR VÁRIAS OS's está bloqueado para essa empresa."), vbExclamation
                TBAbrir.Close
                Exit Sub
            End If
            If TBAbrir!Desbloquear_prim_apont_OS_proc_controlado = False And Processo_controlado = True And txtquant = "0" Then
                'MsgBox ("PROCESSO CONTROLADO, essa OS não está disponível para apontamento, a OS anterior deve ser apontada primeiro."), vbExclamation
                'TBAbrir.Close
                'Exit Sub
            End If
        End If
        TBAbrir.Close
                
        'Verifica se é apontamento de várias OS
        Set TBOrdem = CreateObject("adodb.recordset")
        TBOrdem.Open "Select OrdemServico.ID_apontamento FROM ProducaoFases_OS INNER JOIN OrdemServico ON ProducaoFases_OS.ID = OrdemServico.ID_apontamento where OrdemServico.IDProducao = " & txtos, Conexao, adOpenKeyset, adLockOptimistic
        If TBOrdem.EOF = True Then
            If Varias_OS = True Then
                Set TBAbrir = CreateObject("adodb.recordset")
                TBAbrir.Open "Select * from ProducaoFases_OS", Conexao, adOpenKeyset, adLockOptimistic
                TBAbrir.AddNew
                TBAbrir!Planotexto = ProcCriarNovoNumeroPP
                TBAbrir!Data = Date
                TBAbrir!Responsavel = txtusuario
                TBAbrir!DtValidacao = Now
                TBAbrir!RespValidacao = txtusuario
                TBAbrir.Update
                Txt_ID_apontamento = TBAbrir!ID
                Conexao.Execute "Update OrdemServico Set ID_apontamento = " & Txt_ID_apontamento & " where IDproducao = " & txtos
                TBAbrir.Close
                
                Permitido = True
                Contador = 1
                Do While Permitido = True
                    If Contador > 1 Then
                        If MsgBox("Deseja adicionar outra OS neste apontamento?", vbYesNo) = vbNo Then Permitido = False
                    End If
                    If Permitido = True Then
                        OS_texto = ""
Mensagem1:
                        OS_texto = InputBox("Favor informar o número da outra OS que será apontada.")
                        If OS_texto = "" Then
                            If Contador > 1 Then GoTo Prosseguir
                            If MsgBox("Deseja cancelar esta operação?", vbYesNo) = vbYes Then
                                Conexao.Execute "Update OrdemServico Set ID_apontamento = Null where IDproducao = " & txtos
                                Conexao.Execute "DELETE ProducaoFases_OS where Id = " & Txt_ID_apontamento
                                Varias_OS = False
                                GoTo Prosseguir
                            Else
                                GoTo Mensagem1
                            End If
                        End If
                        If IsNumeric(OS_texto) = False Then
                            MsgBox ("Só é permitido número neste campo."), vbExclamation
                            GoTo Mensagem1
                        End If
                        OS = OS_texto
                        
                        'Verifica o grupo do posto de trabalho selecionado
                        GrupoFiltro = ""
                        GrupoMsg = ""
                        Set TBAbrir = CreateObject("adodb.recordset")
                        TBAbrir.Open "Select grupo from CadMaquinas where Maquina = '" & cmbmaquina & "' and Grupo IS NOT NULL and Grupo <> N''", Conexao, adOpenKeyset, adLockOptimistic
                        If TBAbrir.EOF = False Then
                            GrupoFiltro = " and CM.Grupo = '" & TBAbrir!Grupo & "'"
                            GrupoMsg = ", ou o posto de trabalho da mesma não pertence ao grupo " & TBAbrir!Grupo
                        End If
                        TBAbrir.Close
                        
                        Set TBFases = CreateObject("adodb.recordset")
                        TBFases.Open "Select OS.* from OrdemServico OS INNER JOIN CadMaquinas CM ON CM.Maquina = OS.Maquina where OS.IDProducao = " & OS & GrupoFiltro, Conexao, adOpenKeyset, adLockOptimistic
                        If TBFases.EOF = True Then
                            MsgBox ("Não foi encontrado nenhuma OS com este número" & GrupoMsg & "."), vbExclamation
                            GoTo Mensagem1
                        Else
                            If IsNull(TBFases!ID_apontamento) = False And TBFases!ID_apontamento <> "" Then
                                MsgBox ("Esta OS já esta sendo utilizada."), vbExclamation
                                GoTo Mensagem1
                            End If
                            
                            'Verifica se a OS já esta sendo apontada
                            Set TBAbrir = CreateObject("adodb.recordset")
                            TBAbrir.Open "Select * from Ordemservico_maq_utilizadas where OS = " & OS, Conexao, adOpenKeyset, adLockOptimistic
                            If TBAbrir.EOF = False Then
                                MsgBox ("Esta OS já esta sendo apontada."), vbExclamation
                                GoTo Mensagem1
                            End If
                            TBAbrir.Close
                            
                            TBFases!ID_apontamento = Txt_ID_apontamento
                            TBFases.Update
                        End If
                        TBFases.Close
                        Contador = Contador + 1
                    End If
                Loop
            End If
        Else
            Txt_ID_apontamento = TBOrdem!ID_apontamento
        End If
        TBOrdem.Close
        
Prosseguir:
        ProcVerifAPCodigo False, TXTOF
        
        With frmProducao
            .ProcAbrir
            .ProcLista12Ultimos
            .txtLote = txtquant.Text
        End With
        If OrdemExiste = True Then Unload Me
    Case vbKeyF2:
        Cmd_visualizar_Click
    Case vbKeyEscape:

        ProcLogonOut (txtusuario)
        frmfundo.Show
        Unload Me
End Select

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Function ProcCriarNovoNumeroPP() As String
On Error GoTo tratar_erro

Set TBFI = CreateObject("adodb.recordset")
TBFI.Open "Select Planotexto from ProducaoFases_OS where Year (Data) = '" & Year(Date) & "' order by ID desc", Conexao, adOpenKeyset, adLockOptimistic
If TBFI.EOF = False Then
    Numero = Left(TBFI!Planotexto, Len(TBFI!Planotexto) - 3)
    Numero = Right(Numero, 5) + 1
Else
    Numero = 1
End If
TBFI.Close

A = Numero
Ano = Right(Year(Date), 2)
Select Case Len(A)
    Case 1: ProcCriarNovoNumeroPP = "PLP-0000" & Numero & "/" & Ano
    Case 2: ProcCriarNovoNumeroPP = "PLP-000" & Numero & "/" & Ano
    Case 3: ProcCriarNovoNumeroPP = "PLP-00" & Numero & "/" & Ano
    Case 4: ProcCriarNovoNumeroPP = "PLP-0" & Numero & "/" & Ano
    Case 5: ProcCriarNovoNumeroPP = "PLP-" & Numero & "/" & Ano
End Select

Exit Function
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Function
End Function

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo tratar_erro
    
If Codigo_Barras = False Then Call FunKeyAscii(KeyAscii)

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

If Codigo_Barras = True Then
    Caption = "Gerprod - Coletor de dados no chão de fábrica - Código de barras - Nome do banco de dados: " & Nome_banco & " | Licenças Gerprod: " & Qtlicencas_gerprod & IIf(TemInternet = False, " | Sem internet", "")
    lblsenha.Visible = False
    lblcodigosenha.Visible = True
    txtSenha.ToolTipText = "Código/Senha do usuário"
Else
    Caption = "Gerprod - Coletor de dados no chão de fábrica - Nome do banco de dados: " & Nome_banco & " | Licenças Gerprod: " & Qtlicencas_gerprod & IIf(TemInternet = False, " | Sem internet", "")
    lblsenha.Visible = True
    lblcodigosenha.Visible = False
    txtSenha.ToolTipText = "Senha do usuário"
End If

lblVersaoatual.Caption = "v" & App.Major & "." & App.Minor & "." & App.Revision
lblano.Caption = "Copyright 1999 - " & Year(Date) & " Caprind Sistemas ®. Todos os direitos reservados."

    TOK = 0
    TNC = 0
    TCD = 0
    QT_Entrada_Estoque = 0

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

'WindowState = 2

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub txtSenha_Change()
On Error GoTo tratar_erro

txtusuario.Text = ""

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

If txtSenha.Text <> "" Then
    Set TBUsuarios = CreateObject("adodb.recordset")
    If Codigo_Barras = True Then TextoFiltro = "(Senha = '" & txtSenha.Text & "' or Codigo = '" & txtSenha.Text & "')" Else TextoFiltro = "Senha = '" & txtSenha.Text & "'"
    TBUsuarios.Open "Select Usuario from usuarios where Bloqueado = 'False' and " & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
    If TBUsuarios.EOF = False Then
        txtusuario.Text = TBUsuarios!Usuario
        Operador = txtusuario
    End If
    TBUsuarios.Close
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub



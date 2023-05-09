VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Begin VB.Form frmabertura 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   " GERPROD - Coletor de dados na produção."
   ClientHeight    =   10365
   ClientLeft      =   0
   ClientTop       =   45
   ClientWidth     =   15480
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmabertura.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MousePointer    =   99  'Custom
   ScaleHeight     =   10365
   ScaleWidth      =   15480
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   10000
      Left            =   540
      Top             =   600
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1140
      Top             =   720
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   2160
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin ActiveResizeCtl.ActiveResize ActiveResize1 
      Left            =   765
      Top             =   3375
      _ExtentX        =   847
      _ExtentY        =   847
      Resolution      =   99
      ScreenHeight    =   1080
      ScreenWidth     =   2560
      ScreenHeightDT  =   1080
      ScreenWidthDT   =   1920
      AutoCenterForm  =   -1  'True
      ApplicationName =   "Active Resize Control Professional"
      FormHeightDT    =   10365
      FormWidthDT     =   15480
      FormScaleHeightDT=   10365
      FormScaleWidthDT=   15480
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   345
      Left            =   4755
      TabIndex        =   0
      Top             =   5880
      Width           =   5955
      _ExtentX        =   10504
      _ExtentY        =   609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BarColor1       =   8421504
      BarColor2       =   4210752
      ForeColor2      =   0
      SearchText      =   "Atualizando..."
      Value           =   0
   End
   Begin VB.Label lbldireitos 
      Alignment       =   2  'Center
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
      Height          =   255
      Left            =   90
      TabIndex        =   4
      Top             =   6960
      Width           =   15315
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
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   150
      TabIndex        =   3
      Top             =   6720
      Width           =   15270
   End
   Begin VB.Label lblBD 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   7470
      TabIndex        =   2
      Top             =   6360
      Width           =   480
   End
   Begin DrawSuite2022.USAlphaImage USAlphaImage2 
      Height          =   990
      Left            =   5250
      Top             =   4710
      Width           =   5250
      _ExtentX        =   9260
      _ExtentY        =   1746
      Image           =   "frmabertura.frx":0E42
      Props           =   17
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
      Left            =   13650
      TabIndex        =   1
      Top             =   540
      Width           =   450
   End
   Begin DrawSuite2022.USAlphaImage USAlphaImage1 
      Height          =   2610
      Left            =   6420
      Top             =   2010
      Width           =   2610
      _ExtentX        =   4604
      _ExtentY        =   4604
      Image           =   "frmabertura.frx":4043
      Props           =   5
   End
End
Attribute VB_Name = "frmabertura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
On Error GoTo tratar_erro
Individual = False

lblVersaoatual.Caption = "v" & App.Major & "." & App.Minor & "." & App.Revision
lblano.Caption = "Copyright 1999 - " & Year(Date) & " Caprind Sistemas ®. Todos os direitos reservados."
CB = False


Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Timer1_Timer()
On Error GoTo tratar_erro

Timer1.Enabled = False
Unload Me
frmfundo.Show

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Timer2_Timer()
On Error GoTo tratar_erro

If Situacao = 0 Then
    lblBD.Caption = "Aguarde um instante, acessando base de dados..."
    PBLista.Value = PBLista.Value + 1.5
    Situacao = 1
    Exit Sub
End If
If Situacao = 1 Then
    frmabertura.Caption = " GERPROD V" & App.Major & "." & App.Minor & "." & App.Revision & " - Nome do banco de dados: " & Nome_banco
    Situacao = 0
    Exit Sub
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Timer3_Timer()
On Error GoTo tratar_erro

Timer3.Enabled = False
frmabertura.Visible = True

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

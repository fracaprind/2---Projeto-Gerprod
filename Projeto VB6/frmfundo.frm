VERSION 5.00
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{84147065-0227-424E-827F-9E79B1DA5D8B}#21.0#0"; "kftp.ocx"
Begin VB.Form frmfundo 
   Appearance      =   0  'Flat
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
   Icon            =   "frmfundo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   10035
   ScaleWidth      =   15480
   WindowState     =   2  'Maximized
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   9
      Top             =   9630
      Width           =   15480
      _ExtentX        =   27305
      _ExtentY        =   714
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   555
      Left            =   0
      TabIndex        =   8
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
      Icon            =   "frmfundo.frx":0E42
   End
   Begin VB.TextBox txtfoco 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   450
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   10215
      Width           =   870
   End
   Begin ActiveResizeCtl.ActiveResize ActiveResize1 
      Left            =   450
      Top             =   360
      _ExtentX        =   847
      _ExtentY        =   847
      Resolution      =   99
      ScreenHeight    =   1080
      ScreenWidth     =   2560
      ScreenHeightDT  =   720
      ScreenWidthDT   =   1280
      ApplicationName =   "Active Resize Control Professional"
      FormHeightDT    =   10035
      FormWidthDT     =   15480
      FormScaleHeightDT=   10035
      FormScaleWidthDT=   15480
      ResizeFormBackground=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin KFTPActiveX.kftp kftp 
      Height          =   600
      Left            =   6750
      TabIndex        =   1
      Top             =   11520
      Visible         =   0   'False
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   1058
   End
   Begin DrawSuite2022.USButton Cmd_F2 
      Height          =   870
      Left            =   2430
      TabIndex        =   10
      Top             =   2640
      Width           =   5130
      _ExtentX        =   9049
      _ExtentY        =   1535
      Caption         =   "F2 - ABRIR OS"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Rounded MT Bold"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
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
      ShowFocusRect   =   0   'False
   End
   Begin DrawSuite2022.USButton Cmd_F9 
      Height          =   870
      Left            =   7590
      TabIndex        =   11
      Top             =   2640
      Width           =   6030
      _ExtentX        =   10636
      _ExtentY        =   1535
      Caption         =   "F9 - ABRIR VÁRIAS OS's"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Rounded MT Bold"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
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
      ShowFocusRect   =   0   'False
   End
   Begin DrawSuite2022.USButton Cmd_F3 
      Height          =   870
      Left            =   2430
      TabIndex        =   12
      Top             =   3570
      Width           =   5130
      _ExtentX        =   9049
      _ExtentY        =   1535
      Caption         =   " F3 - ABRIR ORDEM"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Rounded MT Bold"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
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
      ShowFocusRect   =   0   'False
   End
   Begin DrawSuite2022.USButton Cmd_F7 
      Height          =   870
      Left            =   7590
      TabIndex        =   13
      Top             =   3570
      Width           =   6030
      _ExtentX        =   10636
      _ExtentY        =   1535
      Caption         =   "F7 - ABRIR PLANO"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Rounded MT Bold"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
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
      ShowFocusRect   =   0   'False
   End
   Begin DrawSuite2022.USGroupBox USGroupBox1 
      Height          =   2475
      Left            =   2220
      TabIndex        =   5
      Top             =   2160
      Width           =   11565
      _ExtentX        =   20399
      _ExtentY        =   4366
      Appearance      =   0
      Caption         =   "Apontamento utilizando o teclado"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      GradientColor1  =   14737632
   End
   Begin DrawSuite2022.USButton Cmd_F4 
      Height          =   870
      Left            =   2430
      TabIndex        =   14
      Top             =   5220
      Width           =   5100
      _ExtentX        =   8996
      _ExtentY        =   1535
      Caption         =   "F4 - ABRIR OS"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Rounded MT Bold"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
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
      ShowFocusRect   =   0   'False
   End
   Begin DrawSuite2022.USButton Cmd_F11 
      Height          =   870
      Left            =   7560
      TabIndex        =   15
      Top             =   5220
      Width           =   6030
      _ExtentX        =   10636
      _ExtentY        =   1535
      Caption         =   "F11 - ABRIR VÁRIAS OS's"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Rounded MT Bold"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
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
      ShowFocusRect   =   0   'False
   End
   Begin DrawSuite2022.USButton Cmd_F5 
      Height          =   870
      Left            =   2430
      TabIndex        =   16
      Top             =   6150
      Width           =   5100
      _ExtentX        =   8996
      _ExtentY        =   1535
      Caption         =   " F5 - ABRIR ORDEM"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Rounded MT Bold"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
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
      ShowFocusRect   =   0   'False
   End
   Begin DrawSuite2022.USButton Cmd_F12 
      Height          =   870
      Left            =   7560
      TabIndex        =   17
      Top             =   6150
      Width           =   6030
      _ExtentX        =   10636
      _ExtentY        =   1535
      Caption         =   "F12 - ABRIR PLANO"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Rounded MT Bold"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
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
      ShowFocusRect   =   0   'False
   End
   Begin DrawSuite2022.USGroupBox USGroupBox2 
      Height          =   2445
      Left            =   2220
      TabIndex        =   6
      Top             =   4740
      Width           =   11565
      _ExtentX        =   20399
      _ExtentY        =   4313
      Appearance      =   0
      Caption         =   "Apontamento utilizando o leitor código de barras"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      GradientColor1  =   14737632
   End
   Begin DrawSuite2022.USButton Cmd_F6 
      Height          =   600
      Left            =   2430
      TabIndex        =   18
      Top             =   7740
      Width           =   5070
      _ExtentX        =   8943
      _ExtentY        =   1058
      Caption         =   "F6 - IMPRIMIR OS"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
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
      ShowFocusRect   =   0   'False
   End
   Begin DrawSuite2022.USButton Cmd_F8 
      Height          =   600
      Left            =   7530
      TabIndex        =   19
      Top             =   7740
      Width           =   6060
      _ExtentX        =   10689
      _ExtentY        =   1058
      Caption         =   "F8 - ALTERAR BANCO DE DADOS"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
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
      ShowFocusRect   =   0   'False
   End
   Begin DrawSuite2022.USButton Cmd_F10 
      Height          =   600
      Left            =   2430
      TabIndex        =   20
      Top             =   8400
      Width           =   6420
      _ExtentX        =   11324
      _ExtentY        =   1058
      Caption         =   "F10 - SISTEMA DA QUALIDADE"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Rounded MT Bold"
         Size            =   12
         Charset         =   0
         Weight          =   400
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
      ShowFocusRect   =   0   'False
   End
   Begin DrawSuite2022.USButton Cmd_esc 
      Height          =   600
      Left            =   8880
      TabIndex        =   21
      Top             =   8400
      Width           =   4710
      _ExtentX        =   8308
      _ExtentY        =   1058
      Caption         =   "ESC - FINALIZAR"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
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
      ShowFocusRect   =   0   'False
   End
   Begin DrawSuite2022.USGroupBox USGroupBox3 
      Height          =   1815
      Left            =   2220
      TabIndex        =   7
      Top             =   7290
      Width           =   11565
      _ExtentX        =   20399
      _ExtentY        =   3201
      Appearance      =   0
      Caption         =   "Configurações gerais"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      GradientColor1  =   14737632
   End
   Begin DrawSuite2022.USAlphaImage USAlphaImage1 
      Height          =   1620
      Left            =   510
      TabIndex        =   22
      Top             =   840
      Width           =   1620
      _ExtentX        =   2858
      _ExtentY        =   2858
      Image           =   "frmfundo.frx":1C94
      Props           =   5
   End
   Begin DrawSuite2022.USAlphaImage USAlphaImage2 
      Height          =   990
      Left            =   2430
      TabIndex        =   23
      Top             =   1140
      Width           =   5250
      _ExtentX        =   9260
      _ExtentY        =   1746
      Image           =   "frmfundo.frx":EB54
   End
   Begin VB.Label lbldireitos 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmfundo.frx":11D55
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
      TabIndex        =   4
      Top             =   9390
      Width           =   15360
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
      Left            =   45
      TabIndex        =   3
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
      TabIndex        =   2
      Top             =   720
      Width           =   495
   End
End
Attribute VB_Name = "frmfundo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cmd_esc_Click()
On Error GoTo tratar_erro

Call Form_KeyDown(27, 0)

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Cmd_F10_Click()
On Error GoTo tratar_erro

Call Form_KeyDown(121, 0)

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Cmd_F12_Click()
On Error GoTo tratar_erro

Call Form_KeyDown(123, 0)

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Cmd_F3_Click()
On Error GoTo tratar_erro

Call Form_KeyDown(114, 0)

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Cmd_F5_Click()
On Error GoTo tratar_erro

Call Form_KeyDown(116, 0)

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Cmd_F2_Click()
On Error GoTo tratar_erro

Call Form_KeyDown(113, 0)
'Call Form_KeyDown(115, 0)

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Cmd_F4_Click()
On Error GoTo tratar_erro

Call Form_KeyDown(115, 0)
'Call Form_KeyDown(113, 0)

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Cmd_F6_Click()
On Error GoTo tratar_erro

Call Form_KeyDown(117, 0)

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Cmd_F7_Click()
On Error GoTo tratar_erro

Call Form_KeyDown(118, 0)

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Cmd_F8_Click()
On Error GoTo tratar_erro

Call Form_KeyDown(119, 0)

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Cmd_F9_Click()
On Error GoTo tratar_erro

Call Form_KeyDown(120, 0)

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Cmd_F11_Click()
On Error GoTo tratar_erro

Call Form_KeyDown(122, 0)

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Codigo_Barras = False
Varias_OS = False
Ap_plano = False
Select Case KeyCode
    Case vbKeyF2: frmabrir_OS.Show
    Case vbKeyF3: frmabrir_Ordem.Show
    Case vbKeyF4:
        Ap_plano = False
        Codigo_Barras = True
        frmabrir_OS.Show
    Case vbKeyF5:
        Codigo_Barras = True
        frmabrir_Ordem.Show
    Case vbKeyF6: ProcImprimirOS
    Case vbKeyF7:
        Ap_plano = True
        Varias_OS = True
        frmabrir_plano_prod.Show
    Case vbKeyF8: frmOpcoesGeral2.Show 1
    Case vbKeyF9:
        Varias_OS = True
        frmabrir_OS.Show
    Case vbKeyF10: frmCQ_sistema.Show 1
    Case vbKeyF11:
        Codigo_Barras = True
        Varias_OS = True
        frmabrir_OS.Show
    Case vbKeyF12:
        Codigo_Barras = True
        Ap_plano = True
        Varias_OS = True
        frmabrir_plano_prod.Show
    Case vbKeyEscape: ProcEncerra
End Select

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcImprimirOS()
On Error GoTo tratar_erro

Nomerel = "Pcp_ordem.rpt"
NumeroOS = ""
qtdeliberada = 0

Mensagem:
    NumeroOS = InputBox("Favor informar o número da OS.")
    If NumeroOS = "" Then Exit Sub
    If IsNumeric(NumeroOS) = False Then
        MsgBox ("Só é permitido número neste campo."), vbExclamation
        GoTo Mensagem
    End If
    qtdeliberada = NumeroOS
    If qtdeliberada <= 0 Then
        MsgBox ("So é permitido número maior que 0."), vbExclamation
        GoTo Mensagem
    End If
    
    Set TBOS = CreateObject("adodb.recordset")
    TBOS.Open "Select ordemservico.*, Producao.copia_controlada FROM ordemservico INNER JOIN producao ON ordemservico.ordem = producao.ordem where ordemservico.idproducao = " & qtdeliberada & " and ordemservico.pronto = 'NÃO' and producao.status <> 'Cancelada'", Conexao, adOpenKeyset, adLockOptimistic
    If TBOS.EOF = False Then
        TBOS!Copia_controlada = True
        TBOS.Update
        FormulaRel = "{Producao.Ordem} = " & TBOS!Ordem & " and {OrdemServico.idproducao}= " & TBOS!IDProducao
    Else
        MsgBox ("Não foi encontrado nenhuma OS com este número, verifique se a OS já foi concluída ou cancelada."), vbExclamation
        TBOS.Close
        GoTo Mensagem
    End If
    TBOS.Close
    
    If MsgBox("Deseja realmente imprimir a OS " & qtdeliberada & "?", vbYesNo) = vbYes Then
Mensagem1:
        QtdeCopia = ""
        qtdeliberada = 0
        QtdeCopia = InputBox("Favor informar a quantidade de cópias.")
        If QtdeCopia = "" Then Exit Sub
        If IsNumeric(QtdeCopia) = False Then
            MsgBox ("Só é permitido número neste campo."), vbExclamation
            GoTo Mensagem1
        End If
        qtdeliberada = QtdeCopia
        If qtdeliberada <= 0 Then
            MsgBox ("So é permitido número maior que 0."), vbExclamation
            GoTo Mensagem1
        End If
        
        Do While qtdeliberada <> 0
            ProcImprimirDireto FormulaRel, ""
            qtdeliberada = qtdeliberada - 1
        Loop
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

Private Sub ProcEncerra()
On Error GoTo tratar_erro


If USMsgBox("Deseja realmente encerrar o Gerprod?", vbYesNo, "CAPRIND v5.0") = vbYes Then
'    If xPixelsAnt > 1024 And YPixelsAnt = 768 Or xPixelsAnt = 1024 And YPixelsAnt > 768 Or xPixelsAnt > 1024 And YPixelsAnt > 768 Then
'        ProcAlteraResolucaoMonitor xPixelsAnt, YPixelsAnt, GetDeviceCaps(nDC, BITSPIXEL)
'        DeleteDC nDC
'    End If
    End
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

'Muda resolução
'If xPixels = 800 And YPixels = 600 Or xPixels = 1024 And YPixels = 768 Then
'
'Else
'    nDC = CreateDC("DISPLAY", vbNullString, vbNullString, ByVal 0&)
'    ProcAlteraResolucaoMonitor 1024, 768, GetDeviceCaps(nDC, BITSPIXEL)
'End If

'If xPixels = 800 Then Img_800x600.Visible = True Else Img_1024x768.Visible = True
FunAbreBD
ExcluiSel = False

'Verifica se a versão do exe é menor que a do banco de dados
Quant = FunSóNumeros(App.Major & "." & App.Minor & "." & App.Revision & ".txt")
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select VersaoGer from Versao where VersaoGer IS NOT NULL", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    VersaoNova = FunSóNumeros(IIf(TBAbrir!VersaoGer = "", 0, TBAbrir!VersaoGer))
    If Quant < VersaoNova Then
        'MsgBox ("O sistema está desatualizado e será encerrado."), vbExclamation
        'TBAbrir.Close
        'End
    End If
End If

ProcVerificaInternet 'Verifica conexão com a internet
If TemInternet = True Then
    'Verifica se tem versão atualização disponível para baixar
    FunAbreBDSite
    If ConexaoMySql.State = 1 Then
        Set TBMySQL = New adodb.Recordset
        TBMySQL.Open "Select * From Atualizacao_liberada", ConexaoMySql, adOpenKeyset, adLockOptimistic, adCmdText
        If TBMySQL.EOF = False Then
            VersaoNova = FunSóNumeros(IIf(TBMySQL!Versao_Gerprod = "", 0, TBMySQL!Versao_Gerprod))
            If Quant < VersaoNova Then
                'MsgBox ("Existe uma atualização disponível para ser baixada, solicite ao administrador do sistema."), vbInformation
                TBMySQL.Close
            End If
        End If
    End If
End If

Texto = ""
    Numero = 0
    Numero1 = Len(NomeServidor)
    Hora = 0
    If Numero1 <> 1 Then
        Do While Numero1 <> 0
            If Texto = "\" Then GoTo Pula
            Texto = Left(NomeServidor, (Numero + 1))
            Texto = Right(Texto, Len(Texto) - Numero)
            Numero = Numero + 1
            Numero1 = Numero1 - 1
        Loop
    End If
Pula:
    Familiatext = Left(NomeServidor, Numero - 1)
    Dataini = Format(FunHoraServidor("\\" & Familiatext), "dd/mm/yyyy")
    If Dataini <> Date Then
        If Dataini > Date Then MsgTexto = "maior" Else MsgTexto = "menor"
        MsgBox ("A data do computador está " & MsgTexto & " que a data do servidor, favor arrumar antes de acessar o sistema."), vbExclamation
        End
    End If

    ProcLogonOutSemUtilizacao 'Verifica e apaga logon com a data menor que a atual
    ProcVerifQtdeLicencas
    
    'Muda resolução (Verifica focu no programa)
    'gHW = Me.hwnd
    'Hook
lblVersaoatual.Caption = "v" & App.Major & "." & App.Minor & "." & App.Revision
lblano.Caption = "Copyright 1999 - " & Year(Date) & " Caprind Sistemas ®. Todos os direitos reservados."

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

Caption = "Gerprod v" & App.Major & "." & App.Minor & "." & App.Revision & " - Nome do banco de dados: " & Nome_banco & " | Licenças Gerprod: " & Qtlicencas_gerprod & IIf(TemInternet = False, " | Sem internet", "")
'WindowState = 2


Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo tratar_erro

'Muda resolução (Verifica focu no programa)
'Unhook

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub



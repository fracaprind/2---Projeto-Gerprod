VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmOpcoesGeral2_Subs 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Altera caminho do banco de dados"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   8100
   ClipControls    =   0   'False
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
   ForeColor       =   &H00000000&
   Icon            =   "frmOpcoesGeral2_Subs.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   8100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CD1 
      Left            =   5910
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSComctlLib.ListView listaBancos 
      Height          =   1785
      Left            =   60
      TabIndex        =   0
      Top             =   990
      Width           =   7980
      _ExtentX        =   14076
      _ExtentY        =   3149
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16777215
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Tag             =   "N"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Object.Tag             =   "T"
         Text            =   "Local dos relatórios"
         Object.Width           =   5115
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Object.Tag             =   "T"
         Text            =   "Nome do servidor"
         Object.Width           =   4551
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Object.Tag             =   "T"
         Text            =   "Nome do banco de dados"
         Object.Width           =   4233
      EndProperty
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   60
      TabIndex        =   1
      Top             =   2790
      Width           =   7980
      _ExtentX        =   14076
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor2      =   0
      SearchText      =   ""
      Value           =   0
   End
   Begin DrawSuite2022.USToolBar USToolBar1 
      Height          =   975
      Left            =   60
      TabIndex        =   2
      Top             =   0
      Width           =   7980
      _ExtentX        =   14076
      _ExtentY        =   1720
      ButtonCount     =   5
      GradientColor2  =   14737632
      GradientColorOverRight1=   16315633
      GradientColorOverRight2=   15195350
      GripperColor    =   15195350
      IsStrech        =   -1  'True
      RightColor1     =   0
      RightColor2     =   0
      ShowEndPanel    =   0   'False
      Theme           =   1
      ButtonCaption1  =   "Alterar BD"
      ButtonEnabled1  =   0   'False
      ButtonIconSize1 =   32
      ButtonToolTipText1=   "Alterar caminho do acesso (F3)"
      ButtonKey1      =   "1"
      ButtonAlignment1=   2
      BeginProperty ButtonFont1 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft1     =   2
      ButtonTop1      =   2
      ButtonWidth1    =   57
      ButtonHeight1   =   21
      ButtonUseMaskColor1=   0   'False
      ButtonEnabled2  =   0   'False
      ButtonIconSize2 =   32
      ButtonAlignment2=   2
      ButtonType2     =   1
      ButtonStyle2    =   -1
      BeginProperty ButtonFont2 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState2    =   -1
      ButtonLeft2     =   61
      ButtonTop2      =   4
      ButtonWidth2    =   2
      ButtonHeight2   =   54
      ButtonCaption3  =   "Ajuda"
      ButtonEnabled3  =   0   'False
      ButtonIconSize3 =   32
      ButtonToolTipText3=   "Ajuda (F1)"
      ButtonKey3      =   "3"
      ButtonAlignment3=   2
      BeginProperty ButtonFont3 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft3     =   65
      ButtonTop3      =   2
      ButtonWidth3    =   36
      ButtonHeight3   =   21
      ButtonUseMaskColor3=   0   'False
      ButtonCaption4  =   "Sair"
      ButtonEnabled4  =   0   'False
      ButtonIconSize4 =   32
      ButtonToolTipText4=   "Sair (Esc)"
      ButtonKey4      =   "4"
      ButtonAlignment4=   2
      BeginProperty ButtonFont4 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonLeft4     =   103
      ButtonTop4      =   2
      ButtonWidth4    =   26
      ButtonHeight4   =   21
      ButtonUseMaskColor4=   0   'False
      ButtonEnabled5  =   0   'False
      ButtonIconSize5 =   32
      ButtonKey5      =   "5"
      ButtonAlignment5=   2
      BeginProperty ButtonFont5 {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ButtonState5    =   5
      ButtonLeft5     =   131
      ButtonTop5      =   2
      ButtonWidth5    =   24
      ButtonHeight5   =   24
      ButtonUseMaskColor5=   0   'False
      Begin DrawSuite2022.USImageList USImageList1 
         Left            =   4980
         Top             =   120
         _ExtentX        =   900
         _ExtentY        =   767
         Img1            =   "frmOpcoesGeral2_Subs.frx":0E42
         Count           =   1
      End
   End
End
Attribute VB_Name = "frmOpcoesGeral2_Subs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ProcCarregaListaBancos()
On Error GoTo tratar_erro

listaBancos.ListItems.Clear
If Localrel <> "" Then
    With listaBancos.ListItems
        .Add , , 1
        .Item(.Count).SubItems(1) = Localrel
        If NomeServidor <> "" Then .Item(.Count).SubItems(2) = NomeServidor
        If Nome_banco <> "" Then .Item(.Count).SubItems(3) = Nome_banco
    End With
End If
If Localrel1 <> "" Then
    With listaBancos.ListItems
        .Add , , 2
        .Item(.Count).SubItems(1) = Localrel1
        If NomeServidor <> "" Then .Item(.Count).SubItems(2) = NomeServidor1
        If Nome_banco <> "" Then .Item(.Count).SubItems(3) = Nome_banco1
    End With
End If
If Localrel2 <> "" Then
    With listaBancos.ListItems
        .Add , , 3
        .Item(.Count).SubItems(1) = Localrel2
        If NomeServidor <> "" Then .Item(.Count).SubItems(2) = NomeServidor2
        If Nome_banco <> "" Then .Item(.Count).SubItems(3) = Nome_banco2
    End With
End If
If Localrel3 <> "" Then
    With listaBancos.ListItems
        .Add , , 4
        .Item(.Count).SubItems(1) = Localrel3
        If NomeServidor <> "" Then .Item(.Count).SubItems(2) = NomeServidor3
        If Nome_banco <> "" Then .Item(.Count).SubItems(3) = Nome_banco3
    End With
End If
If Localrel <> "" Or Localrel1 <> "" Or Localrel2 <> "" Or Localrel3 <> "" Then PBLista.Value = 100

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcAlterar()
On Error GoTo tratar_erro

If listaBancos.ListItems.Count = 0 Then Exit Sub
If MsgBox("Deseja utilizar o caminho selecionado?", vbYesNo) = vbYes Then
    If listaBancos.SelectedItem.ListSubItems(3) = Nome_banco Then
        Var = NomeServidor
        Var1 = NomeServidor1
        Var2 = NomeServidor2
        Var3 = NomeServidor3
        
        VarE = Nome_banco
        VarE1 = Nome_banco1
        VarE2 = Nome_banco2
        VarE3 = Nome_banco3
        
        VarT = TipoBD
        VarT1 = TipoBD1
        VarT2 = TipoBD2
        VarT3 = TipoBD3
    
        VarR = Localrel
        VarR1 = Localrel1
        VarR2 = Localrel2
        VarR3 = Localrel3
        
        VarU = Usuario_banco
        VarU1 = Usuario_banco1
        VarU2 = Usuario_banco2
        VarU3 = Usuario_banco3
        
        VarS = Senha_banco
        VarS1 = Senha_banco1
        VarS2 = Senha_banco2
        VarS3 = Senha_banco3
        
        VarLAC = LocalAntigoCaprind
        VarLAC1 = LocalAntigoCaprind1
        VarLAC2 = LocalAntigoCaprind2
        VarLAC3 = LocalAntigoCaprind3
        
        VarLNC = LocalNovoCaprind
        VarLNC1 = LocalNovoCaprind1
        VarLNC2 = LocalNovoCaprind2
        VarLNC3 = LocalNovoCaprind3
        
        VarLAG = LocalAntigoGerprod
        VarLAG1 = LocalAntigoGerprod1
        VarLAG2 = LocalAntigoGerprod2
        VarLAG3 = LocalAntigoGerprod3
        
        VarLNG = LocalNovoGerprod
        VarLNG1 = LocalNovoGerprod1
        VarLNG2 = LocalNovoGerprod2
        VarLNG3 = LocalNovoGerprod3
    End If
    If listaBancos.SelectedItem.ListSubItems(3) = Nome_banco1 Then
        Var = NomeServidor1
        Var1 = NomeServidor2
        Var2 = NomeServidor3
        Var3 = NomeServidor
        
        VarE = Nome_banco1
        VarE1 = Nome_banco2
        VarE2 = Nome_banco3
        VarE3 = Nome_banco
        
        VarT = TipoBD1
        VarT1 = TipoBD2
        VarT2 = TipoBD3
        VarT3 = TipoBD
    
        VarR = Localrel1
        VarR1 = Localrel2
        VarR2 = Localrel3
        VarR3 = Localrel
        
        VarU = Usuario_banco1
        VarU1 = Usuario_banco2
        VarU2 = Usuario_banco3
        VarU3 = Usuario_banco
        
        VarS = Senha_banco1
        VarS1 = Senha_banco2
        VarS2 = Senha_banco3
        VarS3 = Senha_banco
        
        VarLAC = LocalAntigoCaprind1
        VarLAC1 = LocalAntigoCaprind2
        VarLAC2 = LocalAntigoCaprind3
        VarLAC3 = LocalAntigoCaprind
        
        VarLNC = LocalNovoCaprind1
        VarLNC1 = LocalNovoCaprind2
        VarLNC2 = LocalNovoCaprind3
        VarLNC3 = LocalNovoCaprind
        
        VarLAG = LocalAntigoGerprod1
        VarLAG1 = LocalAntigoGerprod2
        VarLAG2 = LocalAntigoGerprod3
        VarLAG3 = LocalAntigoGerprod
        
        VarLNG = LocalNovoGerprod1
        VarLNG1 = LocalNovoGerprod2
        VarLNG2 = LocalNovoGerprod3
        VarLNG3 = LocalNovoGerprod
    End If
    If listaBancos.SelectedItem.ListSubItems(3) = Nome_banco2 Then
        Var = NomeServidor2
        Var1 = NomeServidor3
        Var2 = NomeServidor
        Var3 = NomeServidor1
        
        VarE = Nome_banco2
        VarE1 = Nome_banco3
        VarE2 = Nome_banco
        VarE3 = Nome_banco1
        
        VarT = TipoBD2
        VarT1 = TipoBD3
        VarT2 = TipoBD
        VarT3 = TipoBD1
    
        VarR = Localrel2
        VarR1 = Localrel3
        VarR2 = Localrel
        VarR3 = Localrel1
        
        VarU = Usuario_banco2
        VarU1 = Usuario_banco3
        VarU2 = Usuario_banco
        VarU3 = Usuario_banco1
        
        VarS = Senha_banco2
        VarS1 = Senha_banco3
        VarS2 = Senha_banco
        VarS3 = Senha_banco1
        
        VarLAC = LocalAntigoCaprind2
        VarLAC1 = LocalAntigoCaprind3
        VarLAC2 = LocalAntigoCaprind
        VarLAC3 = LocalAntigoCaprind1
        
        VarLNC = LocalNovoCaprind2
        VarLNC1 = LocalNovoCaprind3
        VarLNC2 = LocalNovoCaprind
        VarLNC3 = LocalNovoCaprind1
        
        VarLAG = LocalAntigoGerprod2
        VarLAG1 = LocalAntigoGerprod3
        VarLAG2 = LocalAntigoGerprod
        VarLAG3 = LocalAntigoGerprod1
        
        VarLNG = LocalNovoGerprod2
        VarLNG1 = LocalNovoGerprod3
        VarLNG2 = LocalNovoGerprod
        VarLNG3 = LocalNovoGerprod1
    End If
    If listaBancos.SelectedItem.ListSubItems(3) = Nome_banco3 Then
        Var = NomeServidor3
        Var1 = NomeServidor
        Var2 = NomeServidor1
        Var3 = NomeServidor2
        
        VarE = Nome_banco3
        VarE1 = Nome_banco
        VarE2 = Nome_banco1
        VarE3 = Nome_banco2
        
        VarT = TipoBD3
        VarT1 = TipoBD
        VarT2 = TipoBD1
        VarT3 = TipoBD2
    
        VarR = Localrel3
        VarR1 = Localrel
        VarR2 = Localrel1
        VarR3 = Localrel2
        
        VarU = Usuario_banco3
        VarU1 = Usuario_banco
        VarU2 = Usuario_banco1
        VarU3 = Usuario_banco2
        
        VarS = Senha_banco3
        VarS1 = Senha_banco
        VarS2 = Senha_banco1
        VarS3 = Senha_banco2
        
        VarLAC = LocalAntigoCaprind3
        VarLAC1 = LocalAntigoCaprind
        VarLAC2 = LocalAntigoCaprind1
        VarLAC3 = LocalAntigoCaprind2
        
        VarLNC = LocalNovoCaprind3
        VarLNC1 = LocalNovoCaprind
        VarLNC2 = LocalNovoCaprind1
        VarLNC3 = LocalNovoCaprind2
        
        VarLAG = LocalAntigoGerprod3
        VarLAG1 = LocalAntigoGerprod
        VarLAG2 = LocalAntigoGerprod1
        VarLAG3 = LocalAntigoGerprod2
        
        VarLNG = LocalNovoGerprod3
        VarLNG1 = LocalNovoGerprod
        VarLNG2 = LocalNovoGerprod1
        VarLNG3 = LocalNovoGerprod2
    End If
    
    NomeServidor = Var
    NomeServidor1 = Var1
    NomeServidor2 = Var2
    NomeServidor3 = Var3
    
    Nome_banco = VarE
    Nome_banco1 = VarE1
    Nome_banco2 = VarE2
    Nome_banco3 = VarE3
    
    TipoBD = VarT
    TipoBD1 = VarT1
    TipoBD2 = VarT2
    TipoBD3 = VarT3
    
    Localrel = VarR
    Localrel1 = VarR1
    Localrel2 = VarR2
    Localrel3 = VarR3
    
    Usuario_banco = VarU
    Usuario_banco1 = VarU1
    Usuario_banco2 = VarU2
    Usuario_banco3 = VarU3
        
    Senha_banco = VarS
    Senha_banco1 = VarS1
    Senha_banco2 = VarS2
    Senha_banco3 = VarS3
    
    LocalAntigoCaprind = VarLAC
    LocalAntigoCaprind1 = VarLAC1
    LocalAntigoCaprind2 = VarLAC2
    LocalAntigoCaprind3 = VarLAC3
    
    LocalNovoCaprind = VarLNC
    LocalNovoCaprind1 = VarLNC1
    LocalNovoCaprind2 = VarLNC2
    LocalNovoCaprind3 = VarLNC3
    
    LocalAntigoGerprod = VarLAG
    LocalAntigoGerprod1 = VarLAG1
    LocalAntigoGerprod2 = VarLAG2
    LocalAntigoGerprod3 = VarLAG3
    
    LocalNovoGerprod = VarLNG
    LocalNovoGerprod1 = VarLNG1
    LocalNovoGerprod2 = VarLNG2
    LocalNovoGerprod3 = VarLNG3
    
    SaveSetting "Procam", "CaprindSQL", "NomeServidor", NomeServidor
    SaveSetting "Procam", "CaprindSQL", "NomeServidor1", NomeServidor1
    SaveSetting "Procam", "CaprindSQL", "NomeServidor2", NomeServidor2
    SaveSetting "Procam", "CaprindSQL", "NomeServidor3", NomeServidor3
    SaveSetting "Procam", "CaprindSQL", "Nome_banco", Nome_banco
    SaveSetting "Procam", "CaprindSQL", "Nome_banco1", Nome_banco1
    SaveSetting "Procam", "CaprindSQL", "Nome_banco2", Nome_banco2
    SaveSetting "Procam", "CaprindSQL", "Nome_banco3", Nome_banco3
    SaveSetting "Procam", "CaprindSQL", "TipoBD", TipoBD
    SaveSetting "Procam", "CaprindSQL", "TipoBD1", TipoBD1
    SaveSetting "Procam", "CaprindSQL", "TipoBD2", TipoBD2
    SaveSetting "Procam", "CaprindSQL", "TipoBD3", TipoBD3
    SaveSetting "Procam", "CaprindSQL", "Localrel", Localrel
    SaveSetting "Procam", "CaprindSQL", "Localrel1", Localrel1
    SaveSetting "Procam", "CaprindSQL", "Localrel2", Localrel2
    SaveSetting "Procam", "CaprindSQL", "Localrel3", Localrel3
    SaveSetting "Procam", "CaprindSQL", "Usuario_banco", Usuario_banco
    SaveSetting "Procam", "CaprindSQL", "Usuario_banco1", Usuario_banco1
    SaveSetting "Procam", "CaprindSQL", "Usuario_banco2", Usuario_banco2
    SaveSetting "Procam", "CaprindSQL", "Usuario_banco3", Usuario_banco3
    SaveSetting "Procam", "CaprindSQL", "Senha_banco", Senha_banco
    SaveSetting "Procam", "CaprindSQL", "Senha_banco1", Senha_banco1
    SaveSetting "Procam", "CaprindSQL", "Senha_banco2", Senha_banco2
    SaveSetting "Procam", "CaprindSQL", "Senha_banco3", Senha_banco3
    SaveSetting "Procam", "CaprindSQL", "LocalAntigoCaprind", LocalAntigoCaprind
    SaveSetting "Procam", "CaprindSQL", "LocalAntigoCaprind1", LocalAntigoCaprind1
    SaveSetting "Procam", "CaprindSQL", "LocalAntigoCaprind2", LocalAntigoCaprind2
    SaveSetting "Procam", "CaprindSQL", "LocalAntigoCaprind3", LocalAntigoCaprind3
    SaveSetting "Procam", "CaprindSQL", "LocalNovoCaprind", LocalNovoCaprind
    SaveSetting "Procam", "CaprindSQL", "LocalNovoCaprind1", LocalNovoCaprind1
    SaveSetting "Procam", "CaprindSQL", "LocalNovoCaprind2", LocalNovoCaprind2
    SaveSetting "Procam", "CaprindSQL", "LocalNovoCaprind3", LocalNovoCaprind3
    SaveSetting "Procam", "CaprindSQL", "LocalAntigoGerprod", LocalAntigoGerprod
    SaveSetting "Procam", "CaprindSQL", "LocalAntigoGerprod1", LocalAntigoGerprod1
    SaveSetting "Procam", "CaprindSQL", "LocalAntigoGerprod2", LocalAntigoGerprod2
    SaveSetting "Procam", "CaprindSQL", "LocalAntigoGerprod3", LocalAntigoGerprod3
    SaveSetting "Procam", "CaprindSQL", "LocalNovoGerprod", LocalNovoGerprod
    SaveSetting "Procam", "CaprindSQL", "LocalNovoGerprod1", LocalNovoGerprod1
    SaveSetting "Procam", "CaprindSQL", "LocalNovoGerprod2", LocalNovoGerprod2
    SaveSetting "Procam", "CaprindSQL", "LocalNovoGerprod3", LocalNovoGerprod3

        
    MsgBox "As alterações foram efetuadas com sucesso.", vbInformation
    Unload Me
    FunAbreBD
    ProcVerifQtdeLicencas
    frmfundo.Caption = " GERPROD V" & App.Major & "." & App.Minor & "." & App.Revision & " - Nome do banco de dados: " & Nome_banco & " | Licenças Gerprod: " & Qtlicencas_gerprod & IIf(TemInternet = False, " | Sem internet", "")
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcSair()
On Error GoTo tratar_erro

Unload Me

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyF3: ProcAlterar
    Case vbKeyEscape: ProcSair
    'Case vbKeyF1: ProcAjuda
End Select

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaToolBar1 Me, 8780, 4, True
ProcCarregaListaBancos

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub USToolBar1_ButtonClick(ByVal ButtonIndex As Integer, ByVal Key As String, ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer, ByVal Visible As Boolean)
On Error GoTo tratar_erro

Select Case ButtonIndex
    Case 1: ProcAlterar
    'Case 3: procAjuda
    Case 4: ProcSair
End Select

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

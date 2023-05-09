VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Object = "{50D37AD9-8D3C-43DD-ADD5-7C957C951843}#1.9#0"; "FlexCell.ocx"
Begin VB.Form frmNumeroSerieNC 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Qualidade | NC - Numero de série"
   ClientHeight    =   6210
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5550
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6210
   ScaleWidth      =   5550
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USGroupBox USGroupBox1 
      Height          =   4125
      Left            =   300
      TabIndex        =   4
      Top             =   660
      Width           =   4845
      _ExtentX        =   8546
      _ExtentY        =   7276
      Caption         =   "Número de série"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GradientColor1  =   14737632
      Begin FlexCell.Grid GridDisponiveis 
         Height          =   3615
         Left            =   90
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   420
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   6376
         AllowUserReorderColumn=   -1  'True
         AllowUserResizing=   0   'False
         Appearance      =   0
         BackColor2      =   14737632
         BackColorBkg    =   -2147483644
         BorderColor     =   12632256
         CellBorderColor =   8421504
         SelectionBorderColor=   4210752
         DefaultFontSize =   8.25
         FixedRowColStyle=   2
         GridColor       =   12632256
         Rows            =   1
         ScrollBars      =   2
         ScrollBarStyle  =   0
         SelectionMode   =   1
         MultiSelect     =   0   'False
         EnterKeyMoveTo  =   1
         AllowUserPaste  =   3
      End
   End
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   3
      Top             =   5805
      Width           =   5550
      _ExtentX        =   9790
      _ExtentY        =   714
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   540
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   5550
      _ExtentX        =   9790
      _ExtentY        =   741
      DibPicture      =   "frmNumeroSerieNC.frx":0000
      CaptionDelimiter=   "|"
      CaptionOnCenter =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Icon            =   "frmNumeroSerieNC.frx":9AAD
      IconSize        =   1
      IconSizeX       =   24
      IconSizeY       =   24
      ShowMaximize    =   0   'False
      ShowMinimize    =   0   'False
   End
   Begin DrawSuite2022.USButton Cmd_F3 
      Height          =   795
      Left            =   3090
      TabIndex        =   0
      TabStop         =   0   'False
      ToolTipText     =   "Gravar dados"
      Top             =   4860
      Width           =   2040
      _ExtentX        =   3598
      _ExtentY        =   1402
      DibPicture      =   "frmNumeroSerieNC.frx":9DC7
      BorderColor     =   4960354
      BorderColorDisabled=   13160660
      BorderColorDown =   4210752
      BorderColorOver =   49152
      Caption         =   "(F3) Gravar"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16777215
      ForeColorDown   =   16777215
      ForeColorOver   =   16777215
      GradientColor1  =   4960354
      GradientColor2  =   4960354
      GradientColor3  =   4960354
      GradientColor4  =   4960354
      GradientColorDisabled1=   14215660
      GradientColorDisabled2=   14215660
      GradientColorDisabled3=   14215660
      GradientColorDisabled4=   14215660
      GradientColorDown1=   32768
      GradientColorDown2=   32768
      GradientColorDown3=   32768
      GradientColorDown4=   32768
      GradientColorOver1=   49152
      GradientColorOver2=   49152
      GradientColorOver3=   49152
      GradientColorOver4=   49152
      PicAlign        =   8
      PicSize         =   2
      PicSizeH        =   24
      PicSizeW        =   24
      ShowFocusRect   =   0   'False
      ShowFocusRectDown=   0   'False
      Theme           =   3
   End
   Begin DrawSuite2022.USButton btnExcluir 
      Height          =   795
      Left            =   300
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Excluir dados"
      Top             =   4860
      Width           =   2340
      _ExtentX        =   4128
      _ExtentY        =   1402
      DibPicture      =   "frmNumeroSerieNC.frx":127CC
      BorderColor     =   5263559
      BorderColorDisabled=   13160660
      BorderColorDown =   4013465
      BorderColorOver =   4408288
      Caption         =   "Excluir"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
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
      PicSize         =   2
      PicSizeH        =   24
      PicSizeW        =   24
      ShowFocusRect   =   0   'False
      ShowFocusRectDown=   0   'False
      Theme           =   4
   End
End
Attribute VB_Name = "frmNumeroSerieNC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Cmd_F3_Click()
On Error GoTo tratar_erro

If USMsgBox("Deseja realmente gravar esses dados informados?", vbYesNo, "CAPRIND v5.0") = vbNo Then
 Exit Sub
End If

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from CQ_NC_FABRICA_Serie where Codigo = '" & Codigo & "'", Conexao, adOpenKeyset, adLockOptimistic
Contador = 1
With GridSerie
Linha = .Rows - 1
    For initfor = 1 To Linha
      If .Cell(Contador, 2).Text = "" Then
      USMsgBox "Informe o numero de serie", vbCritical, "GERPROD | COLETOR DE DADOS"
      .Cell(Contador, 2).SetFocus
      TBAbrir.Close
      Exit Sub
      End If
        NumeroSerie = .Cell(Contador, 2).Text
        If TBAbrir.EOF = True Then
            TBAbrir.AddNew
        End If
        TBAbrir!Codigo = Codigo
        TBAbrir!NumeroSerie = NumeroSerie
        TBAbrir!IDProducao = IDProducao
        TBAbrir.Update
        Contador = Contador + 1
        Linha = Linha - 1
        TBAbrir.MoveNext
    Next initfor
End With

USMsgBox "Dados gravados com sucesso", vbInformation, "CAPRIND v5.0"
Unload Me

TBAbrir.Close

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Public Sub ProcAjustaGridSerie()
On Error GoTo tratar_erro

With GridSerie

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
    .Cell(0, 1).Text = "Item"

    .Cell(0, 2).ForeColor = vbRed
    .Cell(0, 2).Text = "Numero de série"
        
    .Column(1).CellType = cellTextBox
    .Column(1).Alignment = cellCenterCenter
    .Column(1).Locked = True
    
    .Column(2).CellType = cellTextBox
    .Column(2).Alignment = cellCenterCenter
    
 
    .Column(0).Width = 8
    .Column(1).Width = 30
    .Column(2).Width = 100

End With

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcCarregaListaItens()
On Error GoTo tratar_erro
Dim L As Long

With GridSerie
    
 L = 1
.Rows = 1

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
'Debug.Print StrSQL

TBAbrir.Open StrSQL, Conexao, adOpenKeyset, adLockOptimistic
  If TBAbrir.EOF = False Then
    Contador2 = 1
    
        Do While TBAbrir.EOF = False
        'Verifica se já apontou
            StrSQL = "Select * from Producao_rastreavel where N_Serie = " & TBAbrir!N_Serie & " and data is not null  and Status is not null and OS = " & OS
            Set TBAfericao = CreateObject("adodb.recordset")
                TBAfericao.Open StrSQL, Conexao, adOpenKeyset, adLockOptimistic
                If TBAfericao.EOF = True Then
                 .AddItem Contador2
                 .Cell(Contador2, 1).Text = TBAbrir!N_Serie
                 .Cell(Contador2, 2).Text = TBAbrir!Status
                Contador2 = Contador2 + 1
                End If
                TBAfericao.Close
                
             TBAbrir.MoveNext
        Loop
  End If


End With
TBAbrir.Close

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcAjustaGridSerie
ProcCarregaListaItens

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyF3: Cmd_F3_Click
End Select

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub btnExcluir_Click()
On Error GoTo tratar_erro

If USMsgBox("Deseja realmente excluir esses dados?", vbYesNo) = vbYes Then
 Conexao.Execute "Delete from CQ_NC_FABRICA_Serie where Codigo = '" & frmcqnc.txtID & "'"
 ProcCarregaListaItens
 USMsgBox "Dados excluidos com sucesso!", vbInformation, "CAPRIND v5.0"
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

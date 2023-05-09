VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2014.ocx"
Object = "{50D37AD9-8D3C-43DD-ADD5-7C957C951843}#1.9#0"; "FlexCell.ocx"
Begin VB.Form frmNumeroSerie 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "Qualidade | NC - Numero de série"
   ClientHeight    =   6300
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4485
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
   ScaleHeight     =   6300
   ScaleWidth      =   4485
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2014.USGroupBox USGroupBox1 
      Height          =   4125
      Left            =   300
      TabIndex        =   5
      Top             =   660
      Width           =   3855
      _ExtentX        =   6800
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
      Begin FlexCell.Grid GridSerie 
         Height          =   3465
         Left            =   150
         TabIndex        =   0
         Top             =   450
         Width           =   3555
         _ExtentX        =   6271
         _ExtentY        =   6112
         AllowUserReorderColumn=   -1  'True
         AllowUserResizing=   0   'False
         Appearance      =   0
         BackColor2      =   14737632
         BackColorBkg    =   -2147483644
         BorderColor     =   12632256
         CellBorderColor =   8421504
         SelectionBorderColor=   4210752
         Cols            =   3
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
   Begin DrawSuite2014.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   4
      Top             =   5895
      Width           =   4485
      _ExtentX        =   7911
      _ExtentY        =   714
   End
   Begin DrawSuite2014.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   540
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   4485
      _ExtentX        =   7911
      _ExtentY        =   953
      DibPicture      =   "frmNumeroSerie.frx":0000
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
      Icon            =   "frmNumeroSerie.frx":9AAD
      IconSize        =   1
      IconSizeX       =   24
      IconSizeY       =   24
      ShowMaximize    =   0   'False
      ShowMinimize    =   0   'False
   End
   Begin DrawSuite2014.USButton Cmd_F3 
      Height          =   795
      Left            =   2370
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Gravar dados"
      Top             =   4860
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   1402
      DibPicture      =   "frmNumeroSerie.frx":9DC7
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
   Begin DrawSuite2014.USButton btnExcluir 
      Height          =   795
      Left            =   300
      TabIndex        =   2
      TabStop         =   0   'False
      ToolTipText     =   "Excluir dados"
      Top             =   4860
      Width           =   1770
      _ExtentX        =   3122
      _ExtentY        =   1402
      DibPicture      =   "frmNumeroSerie.frx":127CC
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
Attribute VB_Name = "frmNumeroSerie"
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
      If .Cell(Contador, 2).text = "" Then
      USMsgBox "Informe o numero de serie", vbCritical, "GERPROD | COLETOR DE DADOS"
      .Cell(Contador, 2).SetFocus
      TBAbrir.Close
      Exit Sub
      End If
        NumeroSerie = .Cell(Contador, 2).text
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
    .Cell(0, 1).text = "Item"

    .Cell(0, 2).ForeColor = vbRed
    .Cell(0, 2).text = "Numero de série"
        
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
 Contador2 = 1
.Rows = 1
   
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from CQ_NC_FABRICA_Serie where Codigo = '" & Codigo & "'", Conexao, adOpenKeyset, adLockOptimistic
  If TBAbrir.EOF = False Then
    Contador = TBAbrir.RecordCount
    
        Do While Contador > 0
         .AddItem Contador2
         .Cell(Contador2, 2).text = TBAbrir!NumeroSerie
         Contador = Contador - 1
         Contador2 = Contador2 + 1
         TBAbrir.MoveNext
        Loop
    
  Else
        Do While Contador > 0
         .AddItem Contador2
         Contador = Contador - 1
         Contador2 = Contador2 + 1
        Loop
  End If


End With
TBAbrir.Close

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
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

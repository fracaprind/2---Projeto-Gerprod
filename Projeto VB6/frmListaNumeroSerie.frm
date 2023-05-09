VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2014.ocx"
Object = "{50D37AD9-8D3C-43DD-ADD5-7C957C951843}#1.9#0"; "FlexCell.ocx"
Begin VB.Form frmListaNumeroSerie 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "GERPROD | Numero de série"
   ClientHeight    =   5655
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5595
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
   ScaleHeight     =   5655
   ScaleWidth      =   5595
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2014.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   465
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   5595
      _ExtentX        =   9869
      _ExtentY        =   820
      DibPicture      =   "frmListaNumeroSerie.frx":0000
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
      Icon            =   "frmListaNumeroSerie.frx":9CB1
      ShowMaximize    =   0   'False
      ShowMinimize    =   0   'False
   End
   Begin DrawSuite2014.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   0
      Top             =   5250
      Width           =   5595
      _ExtentX        =   9869
      _ExtentY        =   714
   End
   Begin DrawSuite2014.USGroupBox USGroupBox1 
      Height          =   4305
      Left            =   360
      TabIndex        =   2
      Top             =   630
      Width           =   4905
      _ExtentX        =   8652
      _ExtentY        =   7594
      Caption         =   "Lista de itens (Numero de série)"
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
         Height          =   3645
         Left            =   150
         TabIndex        =   3
         Top             =   450
         Width           =   4575
         _ExtentX        =   8070
         _ExtentY        =   6429
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
End
Attribute VB_Name = "frmListaNumeroSerie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyEscape:
        Unload Me
End Select

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
    
    .Cell(0, 3).ForeColor = vbRed
    .Cell(0, 3).text = "Status"
 
        
    .Column(1).CellType = cellTextBox
    .Column(1).Alignment = cellCenterCenter
    .Column(1).Locked = True
    
    .Column(2).CellType = cellTextBox
    .Column(2).Alignment = cellCenterCenter
    
    .Column(3).CellType = cellTextBox
    .Column(3).Alignment = cellCenterCenter
    
    .Column(0).Width = 8
    .Column(1).Width = 30
    .Column(2).Width = 90
    .Column(3).Width = 80

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
   
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from ProducaoFases_Codigos where IDProducao = '" & IDApontamento & "'", Conexao, adOpenKeyset, adLockOptimistic
  If TBAbrir.EOF = False Then
    Contador2 = 1 'TBAbrir.RecordCount
    
        Do While TBAbrir.EOF = False
         .AddItem Contador2
         .Cell(Contador2, 2).text = TBAbrir!Codigo
         .Cell(Contador2, 3).text = TBAbrir!Status
         Contador2 = Contador2 + 1
         TBAbrir.MoveNext
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

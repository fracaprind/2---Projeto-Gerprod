VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form frmCQ_sistema 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gerprod - Coletor de dados no chão de fábrica - Sistema da qualidade"
   ClientHeight    =   5250
   ClientLeft      =   3480
   ClientTop       =   4515
   ClientWidth     =   10200
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmCQ_sistema.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5250
   ScaleWidth      =   10200
   StartUpPosition =   2  'Centralziar na Tela
   Begin VB.CommandButton Cmd_visualizar_arquivo 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Height          =   315
      Left            =   9840
      Picture         =   "frmCQ_sistema.frx":0E42
      Style           =   1  'Graphical
      TabIndex        =   2
      ToolTipText     =   "Visualizar arquivo."
      Top             =   630
      Width           =   315
   End
   Begin VB.TextBox Txt_caminho 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Left            =   30
      Locked          =   -1  'True
      MaxLength       =   255
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Caminho da norma."
      Top             =   630
      Width           =   9795
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3630
      Top             =   3060
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   4005
      Left            =   30
      TabIndex        =   0
      Top             =   960
      Width           =   10125
      _ExtentX        =   17859
      _ExtentY        =   7064
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   0
      BackColor       =   16777215
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Código"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   2
         Text            =   "Rev."
         Object.Width           =   970
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   3
         Text            =   "Dt. revisão"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Descrição"
         Object.Width           =   10583
      EndProperty
   End
   Begin DrawSuite2022.USButton Cmd_esc 
      Height          =   420
      Left            =   3082
      TabIndex        =   3
      Top             =   90
      Width           =   4020
      _ExtentX        =   7091
      _ExtentY        =   741
      BorderColor     =   14404026
      BorderColorDown =   11632444
      BorderColorOver =   11632444
      ButtonShape     =   1
      Caption         =   "Esc - Voltar a tela anterior"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
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
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   30
      TabIndex        =   4
      Top             =   4980
      Width           =   10125
      _ExtentX        =   17859
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
      SearchText      =   "Atualizando..."
      Value           =   0
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaco
      BorderWidth     =   2
      Height          =   555
      Left            =   30
      Top             =   30
      Width           =   10125
   End
End
Attribute VB_Name = "frmCQ_sistema"
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

Private Sub Cmd_visualizar_arquivo_Click()
On Error GoTo tratar_erro

If Txt_caminho <> "" Then ProcAbrirArquivo Txt_caminho

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyEscape: Unload Me
End Select

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcCarregaLista

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcCarregaLista()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Set TBLista = CreateObject("adodb.recordset")
TBLista.Open "Select * from CQ_sistema where status <> 'REVISADO' order by codigo, revisao", Conexao, adOpenKeyset, adLockOptimistic
If TBLista.EOF = False Then
    PBLista.Min = 0
    PBLista.Max = TBLista.RecordCount
    PBLista.Value = 1
    Contador = 0
    Do While TBLista.EOF = False
        With Lista.ListItems
            .Add = TBLista!ID
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLista!CODIGO), "", TBLista!CODIGO)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLista!Revisao), "", TBLista!Revisao)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLista!Data_revisao), "", Format(TBLista!Data_revisao, "dd/mm/yy"))
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLista!Descricao), "", TBLista!Descricao)
        End With
        TBLista.MoveNext
        Contador = Contador + 1
        PBLista.Value = Contador
    Loop
End If
TBLista.Close

Exit Sub
tratar_erro:
    If Err.Number = -2147467259 Or Err.Number = 3709 Then
        MsgBox ("Não foi possível estabelecer a conexão com o banco de dados, o sistema será fechado."), vbCritical
        End
    End If
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Lista_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro
  
If Lista.ListItems.Count = 0 Then Exit Sub
Txt_caminho = ""
Set TBLista = CreateObject("adodb.recordset")
TBLista.Open "Select * from CQ_sistema where id = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
If TBLista.EOF = False Then
    ProcPuxaDados
End If
TBLista.Close

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Sub ProcPuxaDados()
On Error GoTo tratar_erro

If TBLista!Caminho <> "" Then Txt_caminho = TBLista!Caminho

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

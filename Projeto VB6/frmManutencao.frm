VERSION 5.00
Object = "{8CA2526B-1F1A-4012-A04D-56C1849DD6A6}#1.5#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.ocx"
Begin VB.Form frmManutencao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gerprod - Coletor de dados no chão de fábrica - Manutenção"
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8490
   ControlBox      =   0   'False
   Icon            =   "frmManutencao.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   8490
   StartUpPosition =   2  'Centralziar na Tela
   Begin MSComctlLib.ListView Lista_check 
      Height          =   4905
      Left            =   30
      TabIndex        =   0
      Top             =   615
      Width           =   8445
      _ExtentX        =   14896
      _ExtentY        =   8652
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Object.Width           =   529
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Descrição"
         Object.Width           =   13758
      EndProperty
   End
   Begin DrawSuite2022.USProgressBar PBLista 
      Height          =   255
      Left            =   30
      TabIndex        =   5
      Top             =   5520
      Width           =   8445
      _ExtentX        =   14896
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
   Begin DrawSuite2022.USButton Cmd_F3 
      Height          =   375
      Left            =   4275
      TabIndex        =   3
      Top             =   120
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   661
      BorderColor     =   14404026
      BorderColorDown =   11632444
      BorderColorOver =   11632444
      ButtonShape     =   1
      Caption         =   "F3 - GRAVAR"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
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
   Begin DrawSuite2022.USButton Cmd_esc 
      Height          =   375
      Left            =   5940
      TabIndex        =   4
      Top             =   120
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   661
      BorderColor     =   14404026
      BorderColorDown =   11632444
      BorderColorOver =   11632444
      ButtonShape     =   1
      Caption         =   "ESC - VOLTAR"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
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
   Begin DrawSuite2022.USButton Cmd_1 
      Height          =   375
      Left            =   930
      TabIndex        =   1
      Top             =   120
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   661
      BorderColor     =   14404026
      BorderColorDown =   11632444
      BorderColorOver =   11632444
      ButtonShape     =   1
      Caption         =   "1 - MARCAR"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
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
   Begin DrawSuite2022.USButton Cmd_2 
      Height          =   375
      Left            =   2595
      TabIndex        =   2
      Top             =   120
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   661
      BorderColor     =   14404026
      BorderColorDown =   11632444
      BorderColorOver =   11632444
      ButtonShape     =   1
      Caption         =   "2 - DESMARCAR"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
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
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   1  'Opaco
      BorderWidth     =   2
      Height          =   555
      Left            =   30
      Top             =   30
      Width           =   8445
   End
End
Attribute VB_Name = "frmManutencao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub ProcCarregaLista()
On Error GoTo tratar_erro

Lista_check.ListItems.Clear
Set TBProducao = CreateObject("adodb.recordset")
TBProducao.Open "select * from manutencao where IDmaquina = '" & Maquina & "' and Controlada = 'true'", Conexao, adOpenKeyset, adLockOptimistic
If TBProducao.EOF = False Then
    Set TBProcessosDet = CreateObject("adodb.recordset")
    TBProcessosDet.Open "select * from manutencao_data where idManutencao = " & TBProducao!CODIGO & " and status = 'Aberta' and data <= '" & Format(Date, "Short Date") & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBProcessosDet.EOF = False Then
        Set TBLista = CreateObject("adodb.recordset")
        TBLista.Open "select * from manutencao_checklist where id_data = " & TBProcessosDet!ID, Conexao, adOpenKeyset, adLockOptimistic
        If TBLista.EOF = False Then
            TBLista.MoveLast
            PBLista.Min = 0
            PBLista.Max = TBLista.RecordCount
            PBLista.Value = 1
            Contador = 0
            TBLista.MoveFirst
            Do While TBLista.EOF = False
                With Lista_check.ListItems
                    .Add , , TBLista!ID
                    If TBLista!Check = True Then .Item(.Count).Checked = True Else .Item(.Count).Checked = False
                    .Item(.Count).SubItems(1) = IIf(IsNull(TBLista!Descricao), "", TBLista!Descricao)
                End With
                TBLista.MoveNext
                Contador = Contador + 1
                PBLista.Value = Contador
            Loop
        End If
        TBLista.Close
    End If
    TBProcessosDet.Close
End If
TBProducao.Close

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Cmd_1_Click()
On Error GoTo tratar_erro

Call Form_KeyDown(49, 0)

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Cmd_2_Click()
On Error GoTo tratar_erro

Call Form_KeyDown(50, 0)

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

Call Form_KeyDown(114, 0)

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKey0: If Lista_check.ListItems.Count <> 0 Then Lista_check.SelectedItem.Checked = False
    Case vbKey1: If Lista_check.ListItems.Count <> 0 Then Lista_check.SelectedItem.Checked = True
    Case vbKeyF3:
        If Lista_check.ListItems.Count <> 0 Then
            initfor = 0
            With Lista_check
                For initfor = 1 To .ListItems.Count
                    Set TBGravar = CreateObject("adodb.recordset")
                    TBGravar.Open "SELECT * from manutencao_checklist WHERE id = " & Lista_check.ListItems(initfor), Conexao, adOpenKeyset, adLockOptimistic
                    If TBGravar.EOF = False Then
                        If .ListItems.Item(initfor).Checked = True Then TBGravar!Check = True Else TBGravar!Check = False
                        TBGravar.Update
                    End If
                    TBGravar.Close
                Next initfor
            End With
            MsgBox ("Marcação do(s) check-list(s) realizado com sucesso."), vbInformation
        End If
        
        'Cria nova data para manutenção
        Set TBProducao = CreateObject("adodb.recordset")
        TBProducao.Open "select * from manutencao where IDmaquina = '" & Maquina & "' and Controlada = 'true'", Conexao, adOpenKeyset, adLockOptimistic
        If TBProducao.EOF = False Then
            If TBProducao!Tipo = "P" Then
                Set TBProcessosDet = CreateObject("adodb.recordset")
                TBProcessosDet.Open "select * from manutencao_data where idManutencao = " & TBProducao!CODIGO & " and status = 'Aberta' and IDProducao <> 0 and data <= '" & Format(Date, "Short Date") & "'", Conexao, adOpenKeyset, adLockOptimistic
                If TBProcessosDet.EOF = False Then
                    Set TBFiltro = CreateObject("adodb.recordset")
                    TBFiltro.Open "Select * from Manutencao_data where idManutencao = " & TBProducao!CODIGO & " order by Data", Conexao, adOpenKeyset, adLockOptimistic
                    If TBFiltro.BOF = False Then
                        TBFiltro.Find ("ID = " & TBProcessosDet!ID)
                        TBFiltro.MoveNext
                        If TBFiltro.EOF = True Then
                            ProcCopiaDadosData
                            ProcCopiaDadosCheckList
                        End If
                    End If
                    TBFiltro.Close
                    
                    TBProcessosDet!Status = "Concluída"
                    TBProcessosDet!idproducao2 = IDApontamento
                    TBProcessosDet.Update
                End If
            End If
            TBProcessosDet.Close
        End If
        TBProducao.Close
        
        Unload Me
    Case vbKeyEscape:
        Unload Me
End Select

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcCopiaDadosData()
On Error GoTo tratar_erro

Set TBGravar = CreateObject("adodb.recordset")
TBGravar.Open "Select * from Manutencao_data", Conexao, adOpenKeyset, adLockOptimistic
TBGravar.AddNew
TBGravar!IDmanutencao = TBProcessosDet!IDmanutencao
TBGravar!Data = TBProcessosDet!Data + IIf(IsNull(TBProcessosDet!Dias_proxima), 0, TBProcessosDet!Dias_proxima)
TBGravar!Dias_proxima = TBProcessosDet!Dias_proxima
TBGravar!Status = "Aberta"
TBGravar!Obs = IIf(IsNull(TBProcessosDet!Obs), "", TBProcessosDet!Obs)
TBGravar.Update
IDFase = TBGravar!ID
TBGravar.Close

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcCopiaDadosCheckList()
On Error GoTo tratar_erro

Set TBLista = CreateObject("adodb.recordset")
TBLista.Open "Select * from Manutencao_Checklist where ID_Data = " & TBProcessosDet!ID, Conexao, adOpenKeyset, adLockOptimistic
If TBLista.EOF = False Then
    Do While TBLista.EOF = False
        Set TBGravar = CreateObject("adodb.recordset")
        TBGravar.Open "Select * from Manutencao_Checklist", Conexao, adOpenKeyset, adLockOptimistic
        TBGravar.AddNew
        TBGravar!ID_Data = IDFase
        TBGravar!ID_manutencao = TBLista!ID_manutencao
        TBGravar!Descricao = IIf(IsNull(TBLista!Descricao), "", TBLista!Descricao)
        TBGravar!Valor = IIf(IsNull(TBLista!Valor), 0, TBLista!Valor)
        TBGravar!Check = False
        TBGravar.Update
        TBGravar.Close
        TBLista.MoveNext
    Loop
End If
TBLista.Close

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

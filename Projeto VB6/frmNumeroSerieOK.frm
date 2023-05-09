VERSION 5.00
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Begin VB.Form frmNumeroSerieOK 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "GERPROD | Numero de série"
   ClientHeight    =   5280
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10425
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmNumeroSerieOK.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5280
   ScaleWidth      =   10425
   StartUpPosition =   2  'CenterScreen
   Begin DrawSuite2022.USButton cmdGravar 
      Height          =   1215
      Left            =   6210
      TabIndex        =   4
      TabStop         =   0   'False
      ToolTipText     =   "Aprovar numero de serie do item"
      Top             =   630
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   2143
      Caption         =   "(F2) Gravar número de série"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   4960354
      BorderColorDisabled=   13160660
      BorderColorDown =   4210752
      BorderColorOver =   49152
      GradientColor1  =   4960354
      GradientColor2  =   4960354
      GradientColor3  =   4960354
      GradientColor4  =   4960354
      GradientColorDisabled1=   14215660
      GradientColorDisabled2=   14215660
      GradientColorDisabled3=   14215660
      GradientColorDisabled4=   14215660
      GradientColorOver1=   49152
      GradientColorOver2=   49152
      GradientColorOver3=   49152
      GradientColorOver4=   49152
      GradientColorDown1=   32768
      GradientColorDown2=   32768
      GradientColorDown3=   32768
      GradientColorDown4=   32768
      PicAlign        =   8
      ShowFocusRect   =   0   'False
      Theme           =   3
      ToolTipTitle    =   "GERPROD"
   End
   Begin VB.CheckBox chkAprovar 
      BackColor       =   &H0000C000&
      Caption         =   "(F5) Aprovar item"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   4500
      MaskColor       =   &H00004000&
      Style           =   1  'Graphical
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   630
      Width           =   1665
   End
   Begin VB.CheckBox ChkRegistro 
      BackColor       =   &H00000000&
      Caption         =   "(F1) Registro automático"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   2790
      MaskColor       =   &H00004000&
      Style           =   1  'Graphical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   630
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Informações"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   240
      TabIndex        =   9
      Top             =   2010
      Width           =   9885
      Begin VB.Frame Frame4 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Total aprovado na OS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1005
         Left            =   4980
         TabIndex        =   16
         Top             =   300
         Width           =   2295
         Begin VB.TextBox txtTOK 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   405
            Left            =   180
            TabIndex        =   17
            TabStop         =   0   'False
            Text            =   "0"
            ToolTipText     =   "Quantidade conforme."
            Top             =   390
            Width           =   1890
         End
      End
      Begin VB.Frame Frame15 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Total não conforme na OS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1005
         Left            =   7320
         TabIndex        =   14
         Top             =   300
         Width           =   2475
         Begin VB.TextBox txtTNC 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   405
            Left            =   270
            TabIndex        =   15
            TabStop         =   0   'False
            Text            =   "0"
            ToolTipText     =   "Quantidade não conforme."
            Top             =   390
            Width           =   2010
         End
      End
      Begin VB.Frame Frame3 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Total não conforme no AP"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1005
         Left            =   2430
         TabIndex        =   12
         Top             =   300
         Width           =   2475
         Begin VB.TextBox txtNC 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   405
            Left            =   180
            TabIndex        =   13
            TabStop         =   0   'False
            Text            =   "0"
            ToolTipText     =   "Quantidade não conforme."
            Top             =   390
            Width           =   2070
         End
      End
      Begin VB.Frame Frame5 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Total aprovado no AP"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   1005
         Left            =   150
         TabIndex        =   10
         Top             =   300
         Width           =   2235
         Begin VB.TextBox txtOK 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   405
            Left            =   120
            TabIndex        =   11
            TabStop         =   0   'False
            Text            =   "0"
            ToolTipText     =   "Quantidade conforme."
            Top             =   390
            Width           =   1920
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Informe o numero de série"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1215
      Left            =   240
      TabIndex        =   8
      Top             =   630
      Width           =   2505
      Begin VB.TextBox txtNSerie 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   20.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   180
         MaxLength       =   12
         TabIndex        =   0
         ToolTipText     =   "Digite o numero de série do item."
         Top             =   420
         Width           =   2115
      End
   End
   Begin DrawSuite2022.USForm USForm1 
      Align           =   1  'Align Top
      Height          =   465
      Left            =   0
      TabIndex        =   7
      Top             =   0
      Width           =   10425
      _ExtentX        =   18389
      _ExtentY        =   820
      DibPicture      =   "frmNumeroSerieOK.frx":0E42
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
      Icon            =   "frmNumeroSerieOK.frx":AAF3
   End
   Begin DrawSuite2022.USStatusBar USStatusBar1 
      Align           =   2  'Align Bottom
      Height          =   405
      Left            =   0
      TabIndex        =   6
      Top             =   4875
      Width           =   10425
      _ExtentX        =   18389
      _ExtentY        =   714
   End
   Begin DrawSuite2022.USButton cmdExcluir 
      Height          =   1185
      Left            =   8190
      TabIndex        =   5
      TabStop         =   0   'False
      ToolTipText     =   "Excluir numero de serie do item"
      Top             =   660
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   2090
      Caption         =   "(F4) Excluir número de série"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   5263559
      BorderColorDisabled=   13160660
      BorderColorDown =   4013465
      BorderColorOver =   4408288
      GradientColor1  =   5263559
      GradientColor2  =   5263559
      GradientColor3  =   5263559
      GradientColor4  =   5263559
      GradientColorDisabled1=   13160660
      GradientColorDisabled2=   13160660
      GradientColorDisabled3=   13160660
      GradientColorDisabled4=   13160660
      GradientColorOver1=   4408288
      GradientColorOver2=   4408288
      GradientColorOver3=   4408288
      GradientColorOver4=   4408288
      GradientColorDown1=   4013465
      GradientColorDown2=   4013465
      GradientColorDown3=   4013465
      GradientColorDown4=   4013465
      PicAlign        =   8
      ShowFocusRect   =   0   'False
      Theme           =   4
      ToolTipTitle    =   "GERPROD"
   End
   Begin DrawSuite2022.USButton btnDisponiveis 
      Height          =   1155
      Left            =   240
      TabIndex        =   1
      TabStop         =   0   'False
      ToolTipText     =   "Gravar evento."
      Top             =   3540
      Width           =   4905
      _ExtentX        =   8652
      _ExtentY        =   2037
      DibPicture      =   "frmNumeroSerieOK.frx":AE0D
      Caption         =   "(F3) Disponíveis"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColorDown =   15048022
      BorderColorOver =   15381630
      PicAlign        =   8
      PicSize         =   5
      PicSizeH        =   50
      PicSizeW        =   50
      ShowFocusRect   =   0   'False
      ToolTipTitle    =   "GERPROD"
   End
   Begin DrawSuite2022.USButton btnUtilizados 
      Height          =   1155
      Left            =   5220
      TabIndex        =   18
      TabStop         =   0   'False
      ToolTipText     =   "Gravar evento."
      Top             =   3540
      Width           =   4905
      _ExtentX        =   8652
      _ExtentY        =   2037
      DibPicture      =   "frmNumeroSerieOK.frx":14FBA
      Caption         =   "(F6) Utilizados"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColorDown =   15048022
      BorderColorOver =   15381630
      PicAlign        =   8
      PicSize         =   5
      PicSizeH        =   50
      PicSizeW        =   50
      ShowFocusRect   =   0   'False
      ToolTipTitle    =   "GERPROD"
   End
End
Attribute VB_Name = "frmNumeroSerieOK"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub ProcGravar()
On Error GoTo tratar_erro

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from ProducaoFases_Codigos where Codigo = '" & txtNSerie.Text & "' and OS = '" & frmProducao.Txt_OS.Text & "'", Conexao, adOpenKeyset, adLockOptimistic

If TBAbrir.EOF = False Then

USMsgBox "Numero de série já utilizado, favor informar um numero de série não utilizado", vbCritical, "CAPRIND v5.0"

txtNSerie.Text = ""
txtNSerie.SetFocus
Exit Sub
End If

TBAbrir.AddNew
TBAbrir!Data = Now
TBAbrir!OS = frmProducao.Txt_OS.Text
TBAbrir!Codigo = txtNSerie.Text
TBAbrir!Responsavel = Operador
TBAbrir!Status = Status
TBAbrir!IDProducao = IDApontamento
TBAbrir.Update
TBAbrir.Close

Set TBUsuarios = CreateObject("adodb.recordset")
TBUsuarios.Open "Select CODIGO from Usuarios where Usuario = '" & Operador & "' and Bloqueado = 'False'", Conexao, adOpenKeyset, adLockOptimistic
If TBUsuarios.EOF = False Then
    If IsNull(TBUsuarios!Codigo) = False And TBUsuarios!Codigo <> "" Then OperadorTexto = TBUsuarios!Codigo & "-" & Operador Else OperadorTexto = Operador
End If
TBUsuarios.Close


If Status = "NÃO CONFORME" Then
Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from CQ_NC_FABRICA", Conexao, adOpenKeyset, adLockOptimistic
    TBAbrir.AddNew
    TBAbrir!IDProducao = IDApontamento
    TBAbrir!Ordem = frmProducao.ListaOS.ListItems.Item(1).Text
    TBAbrir!OS = frmProducao.Txt_OS.Text
    TBAbrir!TTNC = 1 'Replace(TNC, ",", ".")
    TBAbrir!LOTE = Replace(frmProducao.ListaOS.ListItems.Item(1).ListSubItems(4).Text, ",", ".")
    TBAbrir!Data = Date
    TBAbrir!Hora = Time
    TBAbrir!Maquina = frmProducao.txtMaquina
    TBAbrir!Turno = frmProducao.txtturno
    'TBAbrir!ParecerCQ = Status
    TBAbrir!Operador = OperadorTexto
    TBAbrir!Setor = "" 'frmProducao.txt
    TBAbrir!NumeroSerie = txtNSerie.Text
    TBAbrir.Update
    TBAbrir.Close
    QTNC = frmProducao.Lista.SelectedItem.ListSubItems(9).Text + 1
    Conexao.Execute "Update ProducaoFases Set Reprovada = " & QTNC & " Where IDProducao = '" & frmProducao.Lista.SelectedItem & "'"
Else
    QTOK = frmProducao.Lista.SelectedItem.ListSubItems(7).Text + 1
    Conexao.Execute "Update ProducaoFases Set QUANTIDADE = " & QTOK & " Where IDProducao = '" & frmProducao.Lista.SelectedItem & "'"
    QT_Entrada_Estoque = QTOK
End If
frmProducao.ProcLista12Ultimos
'Unload Me

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub btnDisponiveis_Click()
On Error GoTo tratar_erro

    frmNumeroSerieDisponiveis.Show 1

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub btnUtilizados_Click()
On Error GoTo tratar_erro

    frmNumeroSerieUtilizados.Show 1

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ChkAprovar_Click()
On Error GoTo tratar_erro


    If chkAprovar.Value = 0 Then
    chkAprovar.BackColor = &H119CF3
    chkAprovar.Caption = "(F5) Aprovar item"
    Else
    chkAprovar.BackColor = &H5050C7
    chkAprovar.Caption = "(F5) Rejeitar item"
    End If
    
Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ChkRegistro_Click()
On Error GoTo tratar_erro

If ChkRegistro.Value = 0 Then
ChkRegistro.BackColor = &H0
ChkRegistro.ForeColor = &HFFFFFF
ChkRegistro.Caption = "(F1) Registro automático"
cmdGravar.Visible = False
cmdExcluir.Visible = False
Else
ChkRegistro.BackColor = &H119CF3
ChkRegistro.ForeColor = &H80000012
ChkRegistro.Caption = "(F1) Registro manual"
cmdGravar.Visible = True
cmdExcluir.Visible = True
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcAprovar()
On Error GoTo tratar_erro

OS = frmProducao.txtos.Text

If txtNSerie.Text = "" Then
    USMsgBox "É obrigatorio informar o numero de série do item antes de gravar", vbCritical, "CAPRIND v5.0"
    Exit Sub
End If

Status = "APROVADO"

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from ProducaoFases_Codigos where Codigo = '" & txtNSerie.Text & "' and OS = '" & OS & "'", Conexao, adOpenKeyset, adLockOptimistic

If TBAbrir.EOF = False Then

USMsgBox "Numero de série já utilizado, favor informar um numero de série não utilizado", vbCritical, "CAPRIND v5.0"

txtNSerie.Text = ""
txtNSerie.SetFocus
Exit Sub
End If

    
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from Producao_rastreavel where N_serie = '" & txtNSerie.Text & "' and Status = 'NÃO CONFORME'", Conexao, adOpenKeyset, adLockOptimistic

    If TBAbrir.EOF = False Then
        USMsgBox "Numero de série informado está com Status Não conforme." & vbCrLf & "Informe um numero de série válido", vbCritical, "CAPRIND v5.0"
        txtNSerie.Text = ""
        txtNSerie.SetFocus
        Exit Sub
    End If
    
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from Producao_rastreavel where N_serie = '" & txtNSerie.Text & "' and Status <> 'NÃO CONFORME'", Conexao, adOpenKeyset, adLockOptimistic

    If TBAbrir.EOF = True Then
        USMsgBox "Numero de série informado não existe." & vbCrLf & "Informe um numero de série válido", vbCritical, "CAPRIND v5.0"
        txtNSerie.Text = ""
        txtNSerie.SetFocus
        Exit Sub
    End If
   
'============================================================================
'Verifica se é a primeira OS
'============================================================================
PrimeiraOS = False
Set TBCFOP = CreateObject("adodb.recordset")
TBCFOP.Open "Select * from ordemservico where Ordem = " & Ordem & " AND rastreavel = '1' order by fase, retrabalho, IDproducao", Conexao, adOpenKeyset, adLockOptimistic
If TBCFOP.EOF = False Then
        TBCFOP.MoveFirst
        If TBCFOP!IDProducao = OS Then
            PrimeiraOS = True
        End If
End If
TBCFOP.Close

'====================================================================
' Verifica se está informado o numero de série do item na OS anterior
'====================================================================
OSAnterior = OS - 1

If PrimeiraOS = False Then
    Set TBAbrir = CreateObject("adodb.recordset")
    StrSQL = "Select * from Producao_rastreavel where N_Serie = '" & txtNSerie.Text & "' and Status <> 'NÃO CONFORME' and data is not null and OS = " & OS - 1
    'Debug.Print StrSQL
    
    TBAbrir.Open StrSQL, Conexao, adOpenKeyset, adLockOptimistic

    If TBAbrir.EOF = True Then
        USMsgBox "Numero de série não informado na ordem de serviço anterior." & vbCrLf & "Informe um numero de série válido", vbCritical, "CAPRIND v5.0"
        txtNSerie.Text = ""
        'txtNSerie.SetFocus
        Exit Sub
    Else

        If TBAbrir!Status <> "APROVADO" Then
        USMsgBox "Numero de série " & txtNSerie.Text & " está com status " & TBAbrir!Status & " na ordem de serviço " & TBAbrir!OS & "." & vbCrLf & "Informe um numero de série válido", vbCritical, "CAPRIND v5.0"
        txtNSerie.Text = ""
        'txtNSerie.SetFocus
        Exit Sub
        End If
    End If
End If

ProcGravar
ProcAtualizaTotaisOS
ProcAtualizaTotaisAP
txtNSerie.Text = ""
'txtNSerie.SetFocus


Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcExcluir()
On Error GoTo tratar_erro


If txtNSerie.Text <> "" Then

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * from ProducaoFases_Codigos where Codigo = '" & txtNSerie & "' and OS = '" & OS & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = True Then
USMsgBox "Numero de série " & txtNSerie.Text & " inexistente nos apontamentos!", vbInformation, "CAPRIND v5.0"
Exit Sub
End If

If USMsgBox("Deseja realmente excluir o numero de série " & txtNSerie.Text & " do apontamento?", vbYesNo, "CAPRIND v5.0") = vbYes Then
NumeroSerie = txtNSerie.Text
Conexao.Execute "Delete from ProducaoFases_Codigos where Codigo = '" & txtNSerie & "' and OS = '" & OS & "'"
USMsgBox "Numero de série " & txtNSerie.Text & " excluido com sucesso do apontamento!", vbInformation, "CAPRIND v5.0"
txtNSerie.Text = ""
End If

ProcAtualizaTotaisAP
ProcAtualizaTotaisOS


frmProducao.ProcLista12Ultimos

End If


Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub cmdExcluir_Click()
On Error GoTo tratar_erro

ProcExcluir

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub cmdGravar_Click()
On Error GoTo tratar_erro

    If ChkRegistro.Value = 1 And chkAprovar.Value = 0 Then
        ProcAprovar
    End If
    
    If ChkRegistro.Value = 1 And chkAprovar.Value = 1 Then
        ProcRejeitar
    End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcRejeitar()
On Error GoTo tratar_erro

If txtNSerie.Text = "" Then
    USMsgBox "É obrigatorio informar o numero de série do item antes de gravar", vbCritical, "CAPRIND v5.0"
    Exit Sub
End If

Status = "NÃO CONFORME"


'============================================================================
'Verifica se é a primeira OS
'============================================================================
PrimeiraOS = False
Set TBCFOP = CreateObject("adodb.recordset")
TBCFOP.Open "Select * from ordemservico where Ordem = " & Ordem & " order by fase, retrabalho, IDproducao", Conexao, adOpenKeyset, adLockOptimistic
If TBCFOP.EOF = False Then
        TBCFOP.MoveFirst
        If TBCFOP!IDProducao = OS Then
            PrimeiraOS = True
        End If
End If
TBCFOP.Close


'====================================================================
' Verifica se está informado o numero de série do item na OS anterior
'====================================================================
OSAnterior = OS - 1


If PrimeiraOS = False Then
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from ProducaoFases_Codigos where Codigo = '" & txtNSerie.Text & "' and OS <> '" & OS & "'", Conexao, adOpenKeyset, adLockOptimistic
    
    If TBAbrir.EOF = True Then
        USMsgBox "Numero de série não informado na ordem de serviço anterior." & vbCrLf & "Informe um numero de série válido", vbCritical, "CAPRIND v5.0"
        txtNSerie.Text = ""
        txtNSerie.SetFocus
        Exit Sub
    Else
        
        If TBAbrir!Status <> "APROVADO" Then
        USMsgBox "Numero de série " & txtNSerie.Text & " está com status " & TBAbrir!Status & " na ordem de serviço " & TBAbrir!OS & "." & vbCrLf & "Informe um numero de série válido", vbCritical, "CAPRIND v5.0"
        txtNSerie.Text = ""
        txtNSerie.SetFocus
        Exit Sub
        End If
    End If
End If

ProcGravar
ProcAtualizaTotaisOS
ProcAtualizaTotaisAP
txtNSerie.Text = ""
txtNSerie.SetFocus

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub

End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyF1:
    
    If ChkRegistro.Value = 0 Then
    ChkRegistro.Value = 1
    Else
    ChkRegistro.Value = 0
    End If
    
    Case vbKeyF5:
    
    If chkAprovar.Value = 0 Then
    chkAprovar.Value = 1
    Else
    chkAprovar.Value = 0
    End If
    
    
    Case vbKeyF2:
    
    If chkAprovar.Value = 0 Then
    ProcAprovar
    Else
    ProcRejeitar
    End If
    
    Case vbKeyF3: frmNumeroSerieDisponiveis.Show 1
    
    Case vbKeyF4: ProcExcluir
    Case vbKeyF6: frmNumeroSerieUtilizados.Show 1
    Case vbKeyEscape: Unload Me
End Select

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcAtualizaTotaisOS()
On Error GoTo tratar_erro

Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select Count(Status) AS QTOK from ProducaoFases_Codigos where OS = '" & frmProducao.Txt_OS & "' and status = 'APROVADO'", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Conexao.Execute "Update ordemservico set QTOK = '" & TBAbrir!QTOK & "' WHERE IDProducao = '" & frmProducao.Txt_OS & "'"
        txtTOK = TBAbrir!QTOK
    End If
    TBAbrir.Close
    
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select Count(Status) AS QTNC from ProducaoFases_Codigos where OS = '" & frmProducao.Txt_OS & "' and status = 'NÃO CONFORME'", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Conexao.Execute "Update ordemservico set QTNC = '" & TBAbrir!QTNC & "' WHERE IDProducao = '" & frmProducao.Txt_OS & "'"
        txtTNC = TBAbrir!QTNC
    End If
    TBAbrir.Close

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub

Private Sub ProcAtualizaTotaisAP()
On Error GoTo tratar_erro
  
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select Count(Status) AS QTOK from ProducaoFases_Codigos where IDProducao = '" & IDProducao & "' and status = 'APROVADO'", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Conexao.Execute "Update ProducaoFases set Quantidade = '" & TBAbrir!QTOK & "' where IDProducao = '" & IDProducao & "'"
        txtOK = TBAbrir!QTOK
    End If
    TBAbrir.Close
    
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select Count(Status) AS QTNC from ProducaoFases_Codigos where IDProducao = '" & IDProducao & "' and status = 'NÃO CONFORME'", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Conexao.Execute "Update ProducaoFases set reprovada = '" & TBAbrir!QTNC & "' where IDProducao = '" & IDProducao & "'"
        txtNC = TBAbrir!QTNC
    End If
    TBAbrir.Close
    
Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
End Sub


Private Sub Form_Load()
On Error GoTo tratar_erro

ProcAtualizaTotaisOS
ProcAtualizaTotaisAP
                                
    If chkAprovar.Value = 0 Then
    chkAprovar.BackColor = &H119CF3
    chkAprovar.Caption = "(F5) Aprovar item"
    Else
    chkAprovar.BackColor = &H5050C7
    chkAprovar.Caption = "(F5) Rejeitar item"
    End If


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error GoTo tratar_erro



Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtNSerie_Change()
On Error GoTo tratar_erro


If Len(txtNSerie.Text) = 8 Then

If txtNSerie.Text = "" Then
    USMsgBox "É obrigatorio informar o numero de série do item antes de gravar", vbCritical, "CAPRIND v5.0"
    Exit Sub
End If

Set TBCodigo = CreateObject("adodb.recordset")
TBCodigo.Open "Select * from Producao_etiquetas where N_Serie = '" & txtNSerie.Text & "'", Conexao, adOpenKeyset, adLockOptimistic
  
  If TBCodigo.EOF = False Then
  
    If TBCodigo!Ordem <> Ordem Then
        txtNSerie.Text = ""
        txtNSerie.SetFocus
        USMsgBox "Esse numero de série não pertence a essa Ordem de produção!", vbCritical, "CAPRIND v5.0"
        Exit Sub
    End If
  End If
  TBCodigo.Close
  

Status = "APROVADO"
    
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from Producao_rastreavel where N_serie = '" & txtNSerie.Text & "' and Status = 'NÃO CONFORME'", Conexao, adOpenKeyset, adLockOptimistic

    If TBAbrir.EOF = False And ChkRegistro.Value = 0 Then
        USMsgBox "Numero de série informado está com Status Não conforme." & vbCrLf & "Informe um numero de série válido", vbCritical, "CAPRIND v5.0"
        txtNSerie.Text = ""
        txtNSerie.SetFocus
        Exit Sub
    End If
    
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select * from Producao_rastreavel where N_serie = '" & txtNSerie.Text & "' and Status <> 'NÃO CONFORME'", Conexao, adOpenKeyset, adLockOptimistic

    If TBAbrir.EOF = True Then
        USMsgBox "Numero de série informado não existe." & vbCrLf & "Informe um numero de série válido", vbCritical, "CAPRIND v5.0"
        txtNSerie.Text = ""
        'txtNSerie.SetFocus
        Exit Sub
    End If
     
  
    If ChkRegistro.Value = 0 And chkAprovar.Value = 0 Then
        ProcAprovar
    End If
    
    If ChkRegistro.Value = 0 And chkAprovar.Value = 1 Then
        ProcRejeitar
    End If
    
    
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtNSerie_GotFocus()
'frmProducao.Refresh
End Sub

Private Sub txtOK_Change()
On Error GoTo tratar_erro

If IsNumeric(txtTNC.Text) = True Then
    QT_Entrada_Estoque = txtOK.Text
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtTNC_Change()
On Error GoTo tratar_erro

If IsNumeric(txtTNC.Text) = True Then
    TNC = txtTNC.Text
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub txtTOK_Change()
On Error GoTo tratar_erro

If IsNumeric(txtTOK.Text) = True Then
    TOK = txtOK.Text
End If

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

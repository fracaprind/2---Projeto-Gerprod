VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{935C9182-411B-4FFB-9512-97C8745743BC}#2.5#0"; "AResize.ocx"
Object = "{4F446E73-0578-46E4-81BC-6A88ADF59FEA}#2.3#0"; "DrawSuite2022.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProducao 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   Caption         =   "GERPROD | Coletor de dados no chão de fábrica"
   ClientHeight    =   10365
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   15480
   ClipControls    =   0   'False
   DrawMode        =   2  'Blackness
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00000000&
   Icon            =   "frmProducao.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   10365
   ScaleWidth      =   15480
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame19 
      Caption         =   "OS principal"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   10290
      TabIndex        =   55
      Top             =   2220
      Width           =   1605
      Begin VB.TextBox txtos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   90
         Locked          =   -1  'True
         TabIndex        =   56
         TabStop         =   0   'False
         Top             =   240
         Width           =   1410
      End
   End
   Begin VB.ComboBox cmbdescricao 
      BackColor       =   &H00FFFFFF&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   405
      ItemData        =   "frmProducao.frx":0E42
      Left            =   2010
      List            =   "frmProducao.frx":0E44
      Style           =   2  'Dropdown List
      TabIndex        =   54
      Top             =   4800
      Width           =   6690
   End
   Begin VB.CommandButton cmdF8 
      Appearance      =   0  'Flat
      BackColor       =   &H008080FF&
      Caption         =   "F8 - Apontar número série"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   8910
      MaskColor       =   &H000000C0&
      Style           =   1  'Graphical
      TabIndex        =   53
      Top             =   4560
      Width           =   1545
   End
   Begin VB.Frame Frame18 
      Caption         =   "Descrição do evento"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   1860
      TabIndex        =   46
      Top             =   4530
      Width           =   7035
      Begin VB.TextBox txtCodigoDesc 
         Alignment       =   2  'Center
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
         Locked          =   -1  'True
         TabIndex        =   49
         TabStop         =   0   'False
         Top             =   270
         Width           =   6735
      End
   End
   Begin VB.Frame Frame17 
      Caption         =   "Lista de eventos apontados na produção"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4485
      Left            =   210
      TabIndex        =   45
      Top             =   5310
      Width           =   10245
      Begin MSComctlLib.ListView Lista 
         Height          =   4155
         Left            =   120
         TabIndex        =   50
         Top             =   240
         Width           =   9990
         _ExtentX        =   17621
         _ExtentY        =   7329
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   -2147483633
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   10
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "N"
            Text            =   "IDProducao"
            Object.Width           =   529
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Tag             =   "N"
            Text            =   "Cód."
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Object.Tag             =   "T"
            Text            =   "Descrição"
            Object.Width           =   4233
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Object.Tag             =   "D"
            Text            =   "Início"
            Object.Width           =   3087
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Object.Tag             =   "D"
            Text            =   "Final"
            Object.Width           =   3087
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Object.Tag             =   "D"
            Text            =   "Tempo total"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Object.Tag             =   "T"
            Text            =   "Operador"
            Object.Width           =   2558
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   7
            Object.Tag             =   "N"
            Text            =   "Aprovado"
            Object.Width           =   2205
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   8
            Text            =   "Condicional"
            Object.Width           =   2205
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   9
            Object.Tag             =   "N"
            Text            =   "Não conforme"
            Object.Width           =   2206
         EndProperty
      End
   End
   Begin VB.Frame Frame16 
      Caption         =   "Componentes do item utilizados na produção"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5235
      Left            =   10470
      TabIndex        =   44
      Top             =   4560
      Width           =   4785
      Begin VB.TextBox txtobservacoes 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   90
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   52
         Top             =   3900
         Width           =   4485
      End
      Begin MSComctlLib.ListView ListaRequisicao 
         Height          =   3600
         Left            =   90
         TabIndex        =   51
         Top             =   240
         Width           =   4485
         _ExtentX        =   7911
         _ExtentY        =   6350
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483633
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   6
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Tag             =   "N"
            Text            =   "id"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Tag             =   "T"
            Text            =   "Cód. interno"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Object.Tag             =   "T"
            Text            =   "Descrição"
            Object.Width           =   5115
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Object.Tag             =   "N"
            Text            =   "Un"
            Object.Width           =   883
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Object.Tag             =   "T"
            Text            =   "Quant."
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Observações"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.Frame Frame15 
      Caption         =   "Posto de trabalho"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   5070
      TabIndex        =   39
      Top             =   2220
      Width           =   5205
      Begin VB.TextBox Txt_descricao_posto 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   60
         Locked          =   -1  'True
         TabIndex        =   41
         TabStop         =   0   'False
         Top             =   240
         Width           =   4905
      End
   End
   Begin VB.Frame Frame14 
      Caption         =   "Lista de ordens de serviço"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1725
      Left            =   210
      TabIndex        =   38
      Top             =   2820
      Width           =   11685
      Begin MSComctlLib.ListView ListaOS 
         Height          =   1215
         Left            =   30
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   300
         Width           =   11490
         _ExtentX        =   20267
         _ExtentY        =   2143
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   -2147483633
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   15
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Ordem"
            Object.Width           =   1236
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "OS"
            Object.Width           =   1323
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Fase"
            Object.Width           =   1147
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "Prazo"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Lote"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   5
            Text            =   "Prep. prev."
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   6
            Text            =   "Prep. real"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   7
            Text            =   "Exec. prev."
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   8
            Text            =   "Exec. real"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   9
            Text            =   "TT prev."
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   10
            Text            =   "TT real"
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   11
            Text            =   "Qt. aprov."
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(13) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   12
            Text            =   "Qt. Cond."
            Object.Width           =   1940
         EndProperty
         BeginProperty ColumnHeader(14) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   13
            Text            =   "Qt. NC"
            Object.Width           =   1324
         EndProperty
         BeginProperty ColumnHeader(15) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   14
            Text            =   "Qt. prod."
            Object.Width           =   1940
         EndProperty
      End
   End
   Begin VB.Frame Frame13 
      Caption         =   "Instrução de trabalho (IT)"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3015
      Left            =   11910
      TabIndex        =   37
      Top             =   1530
      Width           =   3345
      Begin RichTextLib.RichTextBox Txt_instrucao 
         Height          =   2505
         Left            =   90
         TabIndex        =   43
         TabStop         =   0   'False
         ToolTipText     =   "Instrução do serviço."
         Top             =   300
         Width           =   3165
         _ExtentX        =   5583
         _ExtentY        =   4419
         _Version        =   393217
         BackColor       =   -2147483633
         BorderStyle     =   0
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         Appearance      =   0
         TextRTF         =   $"frmProducao.frx":0E46
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.Frame Frame12 
      Caption         =   "Código"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   210
      TabIndex        =   36
      Top             =   4530
      Width           =   1635
      Begin VB.TextBox txtCodigo 
         Alignment       =   2  'Center
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
         ForeColor       =   &H00000080&
         Height          =   405
         Left            =   180
         MaxLength       =   5
         TabIndex        =   48
         Top             =   270
         Width           =   1335
      End
      Begin VB.TextBox txtCodigoBarras 
         Alignment       =   2  'Center
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
         MaxLength       =   6
         TabIndex        =   47
         Top             =   270
         Width           =   1335
      End
   End
   Begin VB.Frame Frame11 
      Caption         =   "Código do posto"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   3210
      TabIndex        =   33
      Top             =   2220
      Width           =   1845
      Begin VB.TextBox txtMaquina 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   90
         Locked          =   -1  'True
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   240
         Width           =   1620
      End
   End
   Begin VB.Frame Frame10 
      Caption         =   "Lote á produzir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   1380
      TabIndex        =   32
      Top             =   2220
      Width           =   1815
      Begin VB.TextBox txtLote 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   60
         Locked          =   -1  'True
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   240
         Width           =   1620
      End
   End
   Begin VB.Frame Frame9 
      Caption         =   "Nº Ordem"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   585
      Left            =   210
      TabIndex        =   31
      Top             =   2220
      Width           =   1155
      Begin VB.TextBox txtOrdem 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   285
         Left            =   60
         Locked          =   -1  'True
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   240
         Width           =   990
      End
   End
   Begin VB.CommandButton Cmd_F12 
      Caption         =   "F12 - LISTAR 7 ULTIMOS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   13110
      TabIndex        =   29
      Top             =   1140
      Width           =   2175
   End
   Begin VB.CommandButton Cmd_F11 
      Caption         =   "F11 - LISTAR TODOS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   11340
      TabIndex        =   28
      Top             =   1140
      Width           =   1755
   End
   Begin VB.CommandButton Cmd_F9 
      Caption         =   "F9  - DESVINCULAR OS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   9450
      TabIndex        =   27
      Top             =   1140
      Width           =   1875
   End
   Begin VB.CommandButton Cmd_F6 
      Caption         =   "F6 - VOLTAR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   7920
      TabIndex        =   26
      Top             =   1140
      Width           =   1515
   End
   Begin VB.CommandButton Cmd_F5 
      Caption         =   "F5 - SUGESTÕES"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6030
      TabIndex        =   25
      Top             =   1140
      Width           =   1875
   End
   Begin VB.CommandButton Cmd_F4 
      Caption         =   "F4 - EXCLUIR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4500
      TabIndex        =   24
      Top             =   1140
      Width           =   1515
   End
   Begin VB.CommandButton Cmd_F3 
      Caption         =   "F3 - GRAVAR"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3030
      TabIndex        =   23
      Top             =   1140
      Width           =   1455
   End
   Begin VB.CommandButton Cmd_F2 
      Caption         =   "F2 - VISUALIZAR ARQUIVO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   210
      TabIndex        =   22
      Top             =   1140
      Width           =   2805
   End
   Begin VB.Frame Frame8 
      Caption         =   "Turno"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   2490
      TabIndex        =   16
      Top             =   1530
      Width           =   705
      Begin VB.TextBox txtturno 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   270
         Left            =   150
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   300
         Width           =   360
      End
   End
   Begin VB.Frame Frame7 
      Caption         =   "Status"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   3210
      TabIndex        =   15
      Top             =   1530
      Width           =   3735
      Begin VB.TextBox txtstatus 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   360
         Left            =   60
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   240
         Width           =   3555
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Efic. prep."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   8070
      TabIndex        =   14
      Top             =   1530
      Width           =   1095
      Begin VB.TextBox Txt_eficiencia_prep 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   270
         Left            =   60
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Text            =   "0,00%"
         Top             =   300
         Width           =   930
      End
   End
   Begin VB.Frame Frame5 
      Caption         =   "Efic. exec."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   9180
      TabIndex        =   13
      Top             =   1530
      Width           =   1095
      Begin VB.TextBox Txt_eficiencia_exec 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   270
         Left            =   90
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Text            =   "0,00%"
         Top             =   300
         Width           =   900
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Efic. média"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   10290
      TabIndex        =   12
      Top             =   1530
      Width           =   1605
      Begin VB.TextBox Txt_eficiencia_media 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   270
         Left            =   90
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Text            =   "0,00%"
         Top             =   270
         Width           =   1380
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Hora"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   1380
      TabIndex        =   9
      Top             =   1530
      Width           =   1095
      Begin VB.TextBox txtHora 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   270
         Left            =   60
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   "00:00:00"
         Top             =   300
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Data"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   210
      TabIndex        =   8
      Top             =   1530
      Width           =   1155
      Begin VB.TextBox txtData 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   270
         Left            =   60
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Text            =   "00/00/00"
         Top             =   300
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tempo"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   675
      Left            =   6960
      TabIndex        =   6
      Top             =   1530
      Width           =   1095
      Begin VB.TextBox TxtTempoUtilizado 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   270
         Left            =   60
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   "00:00:00"
         Top             =   300
         Width           =   975
      End
   End
   Begin ActiveResizeCtl.ActiveResize ActiveResize1 
      Left            =   0
      Top             =   420
      _ExtentX        =   847
      _ExtentY        =   847
      Resolution      =   26
      ScreenHeight    =   1080
      ScreenWidth     =   1920
      ScreenHeightDT  =   1080
      ScreenWidthDT   =   1920
      AutoCenterForm  =   -1  'True
      ApplicationName =   "Active Resize Control Professional"
      FormHeightDT    =   10830
      FormWidthDT     =   15600
      FormScaleHeightDT=   10365
      FormScaleWidthDT=   15480
      ResizeFormBackground=   -1  'True
      HideControlsOnResize=   -1  'True
      ResizePictureBoxContents=   -1  'True
   End
   Begin VB.Timer Timer_logon 
      Interval        =   10000
      Left            =   5085
      Top             =   6015
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3975
      Top             =   6045
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmProducao.frx":0EC1
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer TimerRelogio 
      Interval        =   1
      Left            =   4665
      Top             =   7545
   End
   Begin VB.TextBox txtdesenho 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   450
      Left            =   3300
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   7980
      Visible         =   0   'False
      Width           =   1695
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
      Left            =   1665
      MaxLength       =   20
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   7095
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.TextBox Txt_OS 
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
      Left            =   2715
      MaxLength       =   20
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   6105
      Visible         =   0   'False
      Width           =   1035
   End
   Begin DrawSuite2022.USButton Cmd_esc 
      Height          =   540
      Left            =   11460
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   4860
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   953
      Caption         =   "ESC - VOLTAR"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   14404026
      BorderColorDown =   11632444
      BorderColorOver =   11632444
      ForeColor       =   128
      GradientColor2  =   16777215
      GradientColor3  =   16777215
      GradientColorOver1=   16643560
      GradientColorOver2=   16576988
      GradientColorOver3=   16441780
      GradientColorOver4=   16178091
      GradientColorDown2=   16246986
      GradientColorDown3=   15189380
      GradientColorDown4=   14596208
      PicAlign        =   8
      PicSize         =   4
      PicSizeH        =   48
      PicSizeW        =   48
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
      Left            =   13650
      TabIndex        =   30
      Top             =   690
      Width           =   495
   End
   Begin VB.Label lbldireitos 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   $"frmProducao.frx":0F9C
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
      Left            =   2250
      TabIndex        =   4
      Top             =   900
      Width           =   10950
   End
   Begin VB.Label lblano 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
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
      Left            =   5085
      TabIndex        =   3
      Top             =   675
      Width           =   5700
   End
End
Attribute VB_Name = "frmProducao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Dataini As Date 'OK
Dim DiaSemana As String 'OK
Dim UltimoDesc As String 'OK
Dim UltimoMaq As String 'OK
Dim PenultimoDesc As String 'OK
Dim PenultimoMaq As String 'OK
Dim TempoUtilizadoDescricao As Date 'OK
Dim Aplicacao As String

'===================================================
'Variaveis para a Estrutura
'===================================================
Option Explicit

Private Type NodeData
    Level As Integer
    Text As String
End Type

Public m_Tree As New Node
Public m_Row As Long
Public m_Col As Long
Dim tempNode As Node
Dim intIndex, i As Integer
Dim CodRef As String, ValorCusto As String, DataValidacao As String, RespValidacao As String
Public IDProduto As Long, IDestrutura As Long
Private arrNodes(666) As NodeData

Private Sub cmdF7_Click()
On Error GoTo tratar_erro

If txtMaquina.Text <> "" Then

ProcCarregaListaRequisicao
'ProcMontaGrid
'ProcCarregaEstrutura
End If

frmProducao.Refresh

Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub cmdF8_Click()
On Error GoTo tratar_erro

If txtCodigo.Text = "" And txtCodigoBarras.Text = "" Then
    Exit Sub
End If

If Codigo_Barras = True Then
    If Len(txtCodigoBarras.Text) = 1 Then
    txtCodigoBarras.Text = "00000" & txtCodigoBarras.Text
    End If
    
    If Len(txtCodigoBarras.Text) = 2 Then
    txtCodigoBarras.Text = "0000" & txtCodigoBarras.Text
    End If
End If

If txtCodigo.Text = "2" Or txtCodigoBarras.Text = "000002" Then
    If Lista.ListItems.Count > 0 And Individual = True Then
        IDApontamento = Lista.SelectedItem
        IDProducao = Lista.SelectedItem
        frmNumeroSerieOK.Show
    End If
Else
USMsgBox "Não é permitido lançamento de numero de série nesse evento", vbCritical, "CAPRIND v5.0"

End If


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
    Exit Sub
End Sub

Private Sub ProcVerificacao()
On Error GoTo tratar_erro
Dim MsgTexto As String

'Verifica se é obrigatório baixar material e se o material foi baixado
Set TBFiltro = CreateObject("adodb.recordset")
TBFiltro.Open "Select P.ordem, P.consignacao, P.Retirar_estoque, E.Bloquear_apontamento_sem_baixa from Producao P INNER JOIN Empresa E ON E.Codigo = P.ID_empresa where P.Ordem = " & ListaOS.SelectedItem & " and (E.Bloquear_apontamento_sem_baixa = 'True' or E.Bloquear_apontamento_sem_baixa_total = 'True')", Conexao, adOpenKeyset, adLockOptimistic
If TBFiltro.EOF = False Then
    If TBFiltro!Bloquear_apontamento_sem_baixa = True Then
        TextoFiltro = "and P.SubTipoItem = 0"
        MsgTexto = "a(s) matéria(s)-prima da"
    Else
        TextoFiltro = ""
        MsgTexto = "toda a"
    End If
    
    If TBFiltro!Retirar_estoque = False Then
        Set TBProducao = CreateObject("adodb.recordset")
        TBProducao.Open "Select Ordem from " & NomeTabelaAp & " where Ordem = " & TBFiltro!Ordem, Conexao, adOpenKeyset, adLockOptimistic
        If TBProducao.EOF = True Then
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select PM.Ordem from Producaomaterial PM INNER JOIN Projproduto P ON PM.Codigo = P.Desenho where PM.Ordem = " & TBFiltro!Ordem & " and PM.Saida = 'NÃO' " & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = False Then
                MsgBox ("Não é permitido apontar, pois é obrigatório baixar do estoque " & MsgTexto & " lista de requisição da ordem."), vbExclamation
                Gravar = False
                TBFiltro.Close
                TBFI.Close
                Exit Sub
            End If
            TBFI.Close
        End If
        TBProducao.Close
    End If
    
    'Se for concluir a OS, verifica se é a última OS e se tem quantidade disponível no estoque
    If Evento = 3 Then
        If Varias_OS = True Then
            TextoFiltro = "Select OS.IDProducao, OS.IDFase from (OrdemServico OS INNER JOIN ProducaoFases_OS PFOS ON OS.ID_apontamento = PFOS.ID) INNER JOIN Producao P ON P.Ordem = OS.Ordem where PFOS.ID = " & Txt_ID_apontamento & " order by OS.IDproducao"
        Else
            TextoFiltro = "Select OS.IDProducao, OS.IDFase from OrdemServico OS INNER JOIN Producao P ON P.Ordem = OS.Ordem where OS.IDProducao = " & ListaOS.SelectedItem.ListSubItems(1)
        End If
        Set TBOS = CreateObject("adodb.recordset")
        TBOS.Open TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
        If TBOS.EOF = False Then
            Do While TBOS.EOF = False
                Set TBCFOP = CreateObject("adodb.recordset")
                TBCFOP.Open "Select IDProducao from ordemservico where Ordem = " & TBFiltro!Ordem & " order by fase, retrabalho, IDproducao", Conexao, adOpenKeyset, adLockOptimistic
                If TBCFOP.EOF = False Then
                    TBCFOP.MoveLast
                    If TBCFOP!IDProducao = TBOS!IDProducao And TBOS!IDFase <> 0 Then
                        Set TBMaterial = CreateObject("adodb.recordset")
                        TBMaterial.Open "Select PM.Codigo from (((Producaomaterial PM INNER JOIN Projproduto P ON PM.Codigo = P.Desenho) LEFT JOIN Qtde_saida_estoque_produto QSEP ON QSEP.Ordem = PM.Ordem and QSEP.Desenho = PM.Codigo) LEFT JOIN Qtde_empenhada_produto_detalhado QEPD ON QEPD.Ordem = PM.Ordem and QEPD.Codinterno = PM.Codigo) LEFT JOIN Qtde_estoque_produto QEP ON QEP.Desenho = PM.Codigo where PM.Ordem = " & TBFiltro!Ordem & " and PM.Requisitado - (ISNULL(QSEP.Saida, 0) + ISNULL(QEPD.Qtde_empenhar, 0)) > ISNULL(QEP.Estoque_disponivel, 0) and (PM.Saida = 'NÃO' or PM.Saida = 'PARCIAL')", Conexao, adOpenKeyset, adLockOptimistic
                        If TBMaterial.EOF = False Then
                            MsgBox ("Não é permitido finalizar esta ordem, pois não existe " & MsgTexto & " lista de requisição da ordem disponível no estoque."), vbExclamation
                            Gravar = False
                            TBCFOP.Close
                            TBOS.Close
                            Exit Sub
                        End If
                    End If
                End If
                TBOS.MoveNext
            Loop
        End If
        TBOS.Close
    End If
End If
TBFiltro.Close

If UltimoMaq = txtMaquina And Ultimo = 3 Then
    MsgBox ("Esta ordem de serviço já está concluída."), vbExclamation
    Gravar = False
    Exit Sub
End If
If Codigo_Barras = False Then
    If UltimoMaq = txtMaquina And Ultimo = txtCodigo Then
        USMsgBox ("Não é permitido utilizar este evento, pois o mesmo já esta sendo utilizado."), vbExclamation, "CAPRIND v5.0"
        Gravar = False
        Exit Sub
    End If
    'Se o ultimo código de trabalho for nenhum e o operador quiser terminar o lote não aceita
    If Lista.ListItems.Count = 0 And txtCodigo = 3 Then
        MsgBox ("Só é permitido encerrar a ordem após um evento."), vbExclamation
        Gravar = False
        Exit Sub
    End If
Else
Dim txtUltimo As String

    If Len(Ultimo) = 1 Then
    txtUltimo = "00000" & Ultimo
    End If
    
    If Len(Ultimo) = 2 Then
    txtUltimo = "0000" & Ultimo
    End If

    
    If UltimoMaq = txtMaquina And txtUltimo = txtCodigoBarras Then
        MsgBox ("Não é permitido utilizar este evento, pois o mesmo já esta sendo utilizado."), vbExclamation
        Gravar = False
        Exit Sub
    End If
    'Se o ultimo código de trabalho for nenhum e o operador quiser terminar o lote não aceita
    If Lista.ListItems.Count = 0 And txtCodigoBarras = 3 Then
        MsgBox ("Só é permitido encerrar a ordem após um evento."), vbExclamation
        Gravar = False
        Exit Sub
    End If
End If
'Verifica se a OS não esta aproveitando o SETUP, se tem tempo de preparação previsto e obriga a apontar PREPARANDO MÁQUINA no primeiro apontamento
If TempoPreparacaoReaprov = False And Ultimo = 0 Then
    If Varias_OS = True Then
        TextoFiltro = "Select OrdemServico.* from OrdemServico INNER JOIN ProducaoFases_OS ON OrdemServico.ID_apontamento = ProducaoFases_OS.ID where ProducaoFases_OS.ID = " & Txt_ID_apontamento
    Else
        TextoFiltro = "Select * from OrdemServico where IDproducao = " & Txt_OS
    End If
    Set TBOrdemServico = CreateObject("adodb.recordset")
    TBOrdemServico.Open "" & TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
    If TBOrdemServico.EOF = False Then
        If TBOrdemServico!TempoPreparacao <> "00:00:00" And Evento <> 1 Then
            MsgBox ("Preparando máquina é o primeiro evento que deverá ser apontado."), vbExclamation
            TBOrdemServico.Close
            Gravar = False
            Exit Sub
        End If
    End If
    TBOrdemServico.Close
End If

'Verificar se o turno final e igual ao turno inicio do evento
ProcVerificaTurno
'Verificar se o apontamento inicial do turno no dia esta dentro do permitido pelo turno
ProcVerificaInicioTurnoDia
If Gravar = False Then Exit Sub

'Verifica se as fases anteriores foram concluídas para concluir a OS apontada quando for processo controlado
Permitido = False
If Lista.ListItems.Count = 0 Then
    Set TBAbrir = CreateObject("adodb.recordset")
    TBAbrir.Open "Select Codigo from Empresa where Desbloquear_prim_apont_OS_proc_controlado = 'False'", Conexao, adOpenKeyset, adLockOptimistic
    If TBAbrir.EOF = False Then
        Permitido = True
    End If
Else
    If Codigo_Barras = False Then
        If txtCodigo = 3 Then Permitido = True
    Else
        If txtCodigoBarras = 3 Then Permitido = True
    End If
End If
If Processo_controlado = True And Permitido = True Then
    If Varias_OS = True Then
        Set TBOS = CreateObject("adodb.recordset")
        TBOS.Open "Select OS.Ordem, OS.Fase from OrdemServico OS INNER JOIN ProducaoFases_OS PFOS ON OS.ID_apontamento = PFOS.ID where PFOS.ID = " & Txt_ID_apontamento & " order by OS.Fase, OS.IDproducao", Conexao, adOpenKeyset, adLockOptimistic
        If TBOS.EOF = False Then
            Do While TBOS.EOF = False
                ProcVerificaProcessoControlado TBOS!Ordem, TBOS!Fase
                If Gravar = False Then Exit Sub
                TBOS.MoveNext
            Loop
        End If
        TBOS.Close
    Else
        ProcVerificaProcessoControlado ListaOS.SelectedItem, ListaOS.SelectedItem.ListSubItems(2)
        If Gravar = False Then Exit Sub
    End If
End If

'Verifica se o evento produtivo foi encerrado pelo operador que iniciou
Set TBProducao = CreateObject("adodb.recordset")
TBProducao.Open "Select " & NomeTabelaAp & ".usuario, CD.Tipo from " & NomeTabelaAp & " INNER JOIN CodigoDesc CD ON CD.Codigo = " & NomeTabelaAp & ".CodigoDesc where " & NomeTabelaAp & ".OS = " & Txt_OS & " and " & NomeTabelaAp & ".Maquina = '" & txtMaquina & "' order by " & NomeTabelaAp & ".data desc, " & NomeTabelaAp & ".Tempoinicio desc", Conexao, adOpenKeyset, adLockOptimistic
If TBProducao.EOF = False Then
    If TBProducao!Tipo = "Produtivo" And TBProducao!Usuario <> Operador Then
        MsgBox ("O último evento apontado é produtivo, é necessário que o operador " & TBProducao!Usuario & " encerre o evento."), vbExclamation
        TBProducao.Close
        Gravar = False
        Exit Sub
    End If
End If

'Verifica se existe apontamento no dia com a hora maior que a apontada
Set TBProducao = CreateObject("adodb.recordset")
TBProducao.Open "Select Ordem, OS from " & NomeTabelaAp & " where Data = '" & Date & "' and Tempoinicio > '" & FunHoraServidor("\\" & FunVerifNomeServidor) & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBProducao.EOF = False Then
    MsgBox ("Não é permitido apontar neste horário, pois existe apontamento na ordem " & TBProducao!Ordem & " - OS " & TBProducao!OS & " com o horário maior que o atual."), vbExclamation
    TBProducao.Close
    Gravar = False
    Exit Sub
End If
TBProducao.Close

If Ultimo = "" Then Ultimo = 0
Gravar = True

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcVerificaProcessoControlado(Ordem As Long, Fase As Integer)
On Error GoTo tratar_erro

Set TBOrdemServico = CreateObject("adodb.recordset")
TBOrdemServico.Open "Select IDproducao from ordemservico where Ordem = " & Ordem & " and fase < " & Fase & " and Retrabalho = 'False' and QTOK = 0", Conexao, adOpenKeyset, adLockOptimistic
If TBOrdemServico.EOF = False Then
    USMsgBox "PROCESSO CONTROLADO, não é permitido apontar nesta OS, pois existe(m) OS('s) em fase(s) anterior(es) que não foi(ram) produzida(s) nenhuma peça, favor verificar.", vbExclamation, "GERPROD"
    Gravar = False
    TBOrdemServico.Close
    Exit Sub
End If
TBOrdemServico.Close

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcVerificaTurno()
On Error GoTo tratar_erro

Turno = 0
TempoInicio = 0
TempoFinal = 0
Dataini = Date
ProcVerificaDia
Dataini = Time
Set TBMaquinas = CreateObject("adodb.recordset")
TBMaquinas.Open "Select * from CadMaqturnos where maquina = '" & txtMaquina.Text & "' and diasemana = '" & DiaSemana & "' order by diasemana,turno", Conexao, adOpenKeyset, adLockOptimistic
If TBMaquinas.EOF = False Then
    Do While TBMaquinas.EOF = False
        If IsNull(TBMaquinas!Inicioturno) = False Then
            TempoInicio = Left(TBMaquinas!Inicioturno, 8)
            TempoFinal = Left(TBMaquinas!Finalturno, 8)
            If TempoInicio > TempoFinal Then
                Dataini = txtData.Text & " " & Dataini
                TempoInicio = txtData.Text & " " & TempoInicio
                TempoFinal = txtData.Text & "  " & TempoFinal
                TempoInicio = TempoInicio - 1
                TempoFinal = TempoFinal + 1
            End If
            Select Case TBMaquinas!Turno
                Case 1:
                    If Dataini >= TempoInicio And Dataini <= TempoFinal Then
                        Turno = 1
                        GoTo Sair
                    End If
                Case 2:
                    If Dataini >= TempoInicio And Dataini <= TempoFinal Then
                        Turno = 2
                        GoTo Sair
                    End If
                Case 3:
                    If Dataini >= TempoInicio And Dataini <= TempoFinal Then
                        Turno = 3
                        GoTo Sair
                    End If
                Case 4:
                    If Dataini >= TempoInicio And Dataini <= TempoFinal Then
                        Turno = 4
                        GoTo Sair
                    End If
            End Select
        End If
        TBMaquinas.MoveNext
    Loop
End If
Sair:
    txtturno.Text = Turno
    Dataini = Date
    TBMaquinas.Close
    
Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcVerificaInicioTurnoDia()
On Error GoTo tratar_erro

TempoInicio = 0
TempoFinal = 0
Dataini = Time
Set TBMaquinas = CreateObject("adodb.recordset")
TBMaquinas.Open "Select Inicioturno, Finalturno, Margem_inicio_ap from CadMaqturnos where maquina = '" & txtMaquina.Text & "' and diasemana = '" & DiaSemana & "' and Turno = " & Turno, Conexao, adOpenKeyset, adLockOptimistic
If TBMaquinas.EOF = False Then
    If IsNull(TBMaquinas!Inicioturno) = False Then
        TempoInicio = Left(TBMaquinas!Inicioturno, 8)
        TempoFinal = Left(TBMaquinas!Finalturno, 8)
       'Verifica se o evento apontado anteriormente é fim de produção ou fim de turno
        Set TBProducao = CreateObject("adodb.recordset")
        TBProducao.Open "Select CMM.Evento from CadMaquinas CM INNER JOIN CadMaquinas_Monitor CMM ON CM.Maquina = CMM.Maquina where CM.maquina = '" & txtMaquina.Text & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBProducao.EOF = False Then
            If IsNull(TBProducao!Evento) = True Or TBProducao!Evento = "" Or TBProducao!Evento = 3 Or TBProducao!Evento = 7 Then
                'Verifica se o primeiro apontamento do dia esta igual ou 10 min. após o início do turno
                Set TBProducao = CreateObject("adodb.recordset")
                TBProducao.Open "Select IDProducao from " & NomeTabelaAp & " where maquina = '" & txtMaquina.Text & "' and Data = '" & Format(txtData, "Short Date") & "' and Turno = " & Turno, Conexao, adOpenKeyset, adLockOptimistic
                If TBProducao.EOF = True Then
                    If Dataini < TempoInicio And Dataini > TempoFinal Then
                       MsgBox ("O apontamento dever ser iniciado a partir das " & TempoInicio & "."), vbExclamation
                       TBProducao.Close
                       Gravar = False
                       Exit Sub
                    End If
                    
                    If IsNull(TBMaquinas!Margem_inicio_ap) = False And TBMaquinas!Margem_inicio_ap <> "00:00:00" Then
                        FunElapsedTime (TempoInicio)
                        DecimoSegundos = S
                    
                        FunElapsedTime (TBMaquinas!Margem_inicio_ap)
                        DecimoSegundos = DecimoSegundos + S
                    
                        TempoInicio = FunFormataTempo(DecimoSegundos)
                        If Dataini > TempoInicio Then
                            MsgBox ("O apontamento dever ser iniciado até as " & TempoInicio & "."), vbExclamation
                            TBProducao.Close
                            Gravar = False
                            Exit Sub
                        End If
                    End If
                End If
            End If
        End If
        TBProducao.Close
    End If
End If
    
Exit Sub
tratar_erro:
    If Err.Number = 13 Then
        MsgBox ("Favor verificar a margem para apontamento antes de salvar."), vbExclamation
        Gravar = False
        Exit Sub
    End If
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcVerificaDia()
On Error GoTo tratar_erro

DiaSemana = Weekday(Dataini)
Select Case DiaSemana
    Case 1: DiaSemana = "Domingo"
    Case 2: DiaSemana = "Segunda"
    Case 3: DiaSemana = "Terça"
    Case 4: DiaSemana = "Quarta"
    Case 5: DiaSemana = "Quinta"
    Case 6: DiaSemana = "Sexta"
    Case 7: DiaSemana = "Sabado"
End Select

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcAtualizaPrepExecUtil()
On Error GoTo tratar_erro
Dim TTEUTILDECS As Long

TPPREV = 0
TEPSEG = 0
Eficiencia_prep = 0
Eficiencia_exec = 0
If Varias_OS = True Then
    TextoFiltro = "Select OrdemServico.* from OrdemServico INNER JOIN ProducaoFases_OS ON OrdemServico.ID_apontamento = ProducaoFases_OS.ID where ProducaoFases_OS.ID = " & Txt_ID_apontamento & " order by OrdemServico.IDproducao"
Else
    TextoFiltro = "Select * from OrdemServico where IDproducao = " & ListaOS.SelectedItem.ListSubItems(1)
End If
Set TBOS = CreateObject("adodb.recordset")
TBOS.Open TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
If TBOS.EOF = False Then
    Do While TBOS.EOF = False
        'Filtra tempo total de preparação e execução por OS
        TempoTotalPrep = "00:00:00"
        TempoTotalProd = "00:00:00"
        TTOK = 0
        TTNC = 0
        TTCD = 0
        Set TBProducao = CreateObject("adodb.recordset")
        TBProducao.Open "Select * from " & NomeTabelaAp & " where idfase = " & TBOS!IDProducao & " and (CodigoDesc = 1 or CodigoDesc = 2) order by Data, Tempoinicio", Conexao, adOpenKeyset, adLockOptimistic
        If TBProducao.EOF = False Then
            Do While TBProducao.EOF = False
                If TBProducao!codigoDesc = 1 Then
                    TempoTotalPrep = TempoTotalPrep + IIf(IsNull(TBProducao!TempoTotal), 0, TBProducao!TempoTotal)
                Else
                    TempoTotalProd = TempoTotalProd + IIf(IsNull(TBProducao!TempoTotal), 0, TBProducao!TempoTotal)
                End If
                TTOK = TTOK + IIf(IsNull(TBProducao!Quantidade), 0, TBProducao!Quantidade)
                TTNC = TTNC + IIf(IsNull(TBProducao!Reprovada), 0, TBProducao!Reprovada)
                TTCD = TTCD + IIf(IsNull(TBProducao!QTCD), 0, TBProducao!QTCD)
                TBProducao.MoveNext
            Loop
        End If
        TBProducao.Close
        
        TBOS!QTOK = TTOK
        TBOS!QTNC = TTNC
        TBOS!QTCD = TTCD
        Produzidas = TTOK + TTNC + TTCD
        TBOS!Totalprod = Produzidas
        TBOS.Update
                
        'Verifica tempo total de preparação e execução previsto para todas OS
        'Preparação
        If IsNull(TBOS!Tempo_prep_reaproveitado) = True Or TBOS!Tempo_prep_reaproveitado = False Then TPPREV = TPPREV + IIf(IsNull(TBOS!Preparacao), 0, TBOS!Preparacao)
        
        'Execução
        TEPSEG = TEPSEG + (IIf(IsNull(TBOS!TESegundos), 0, TBOS!TESegundos) * TBOS!Totalprod)
        TBOS.MoveNext
    Loop
    FunElapsedTime (TPPREV)
    TPPSEG = S
    
    'Verificar o percentual do tempo total por OS
    TBOS.MoveFirst
    Do While TBOS.EOF = False
        'Preparação
        FunElapsedTime (TBOS!Preparacao)
        If TPPSEG > 0 Then QuantComprado = (S / TPPSEG) * 100 Else QuantComprado = 100
        
        'Execução
        S = (IIf(IsNull(TBOS!TESegundos), 0, TBOS!TESegundos) * TBOS!Totalprod)
        If TEPSEG > 0 Then QuantComprado1 = (S / TEPSEG) * 100 Else QuantComprado1 = 100
        
        'Preparação
        'Calcula total de segundos utilizados preparando
        FunElapsedTime (TempoTotalPrep)
        TPUSEG = S
        
        TPUSEG = Format((TPUSEG * QuantComprado) / 100, "###,##0.0000000000")
        TPUSEGDECS = TPUSEG
        TTPUTIL = FunFormataTempo(TPUSEG)
        
        TBOS!TPUTIL = TTPUTIL
        TBOS!TPUSEG = TPUSEGDECS
        
        'Execução
        'Calcula total de segundos utilizados produzindo
        FunElapsedTime (TempoTotalProd)
        TTEUTILS = S
        
        TTEUTILS = Format((TTEUTILS * QuantComprado1) / 100, "###,##0.0000000000")
        TTEUTILDECS = TTEUTILS
        TTUTIL = FunFormataTempo(TTEUTILS)
        TTEUTILS = TTEUTILDECS
        
        'Soma tempo total de preparação + execução
        TTUTILSEG = TPUSEGDECS + TTEUTILS
        TempoTotalUtil = FunFormataTempo(TPUSEGDECS + TTEUTILS)
        TBOS!TETTUTIL = TempoTotalUtil 'Grava tempo total utilizado
        'TBOS!TETTUTILN = TempoTotalUtil 'Grava tempo total utilizado
        TBOS!TETTUTILSEG = TTUTILSEG 'Grava tempo total utilizado em segundos
        
        'Calcula tempo total de execução por peça
        If TTEUTILS > 0 And Produzidas > 0 Then TEUSEG = Format(TTEUTILS / Produzidas, "###,##0.0000000000") Else TEUSEG = TTEUTILS
        DecimoSegundos = TEUSEG
        TBOS!TEUTIL = FunFormataTempo(DecimoSegundos)
        TBOS!TEUSEG = TEUSEG
            
        'Calcula eficiencia
        FunElapsedTime (IIf(IsNull(TBOS!Preparacao), 0, TBOS!Preparacao))
        If S > 0 And TPUSEGDECS > 0 Then Eficiencia_prep = Format((S / TPUSEGDECS) * 100, "###,##0.00") Else Eficiencia_prep = 0
        TBOS!Eficiencia_prep = Eficiencia_prep
        If IIf(IsNull(TBOS!TESegundos), 0, TBOS!TESegundos) > 0 And TEUSEG > 0 Then Eficiencia_exec = Format((TBOS!TESegundos / TEUSEG) * 100, "###,##0.00") Else Eficiencia_exec = 0
        TBOS!Eficiencia_exec = Eficiencia_exec
        If Eficiencia_prep > 0 And Eficiencia_exec > 0 Then
            TBOS!Eficiencia = (Eficiencia_prep + Eficiencia_exec) / 2
        ElseIf Eficiencia_prep > 0 Then
                TBOS!Eficiencia = Eficiencia_prep
            ElseIf Eficiencia_exec > 0 Then
                    TBOS!Eficiencia = Eficiencia_exec
                Else
                    TBOS!Eficiencia = 0
        End If
        
        'Calcula e grava valor real do lote e por peça
        Set TBMaquinas = CreateObject("adodb.recordset")
        TBMaquinas.Open "Select Sum(CRLote) as CTTLOTE, Sum(TotalProd) as QtdeProduzida from Ordemservico_maq_utilizadas where OS = " & TBOS!IDProducao, Conexao, adOpenKeyset, adLockOptimistic
        If TBMaquinas.EOF = False Then
            TBOS!CRLOTE = IIf(IsNull(TBMaquinas!CTTLOTE), 0, Format(TBMaquinas!CTTLOTE, "###,##0.00"))
            If IsNull(TBMaquinas!QtdeProduzida) = False And TBMaquinas!QtdeProduzida <> "0" Then TBOS!CRPECA = Format(TBOS!CRLOTE / TBMaquinas!QtdeProduzida, "###,##0.0000000000") Else TBOS!CRPECA = 0
        End If
        TBMaquinas.Close
        
        TBOS.Update
        TBOS.MoveNext
    Loop
End If
TBOS.Close

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcAtualizaPrepExecUtilTotalizacaoMaq()
On Error GoTo tratar_erro
Dim TTEUTILDECS As Long
Dim TBProducaoFases  As adodb.Recordset

TPPREV = 0
TEPSEG = 0
Eficiencia_prep = 0
Eficiencia_exec = 0
If Varias_OS = True Then
    TextoFiltro = "Select OrdemServico.* from OrdemServico INNER JOIN ProducaoFases_OS ON OrdemServico.ID_apontamento = ProducaoFases_OS.ID where ProducaoFases_OS.ID = " & Txt_ID_apontamento & " order by OrdemServico.IDproducao"
Else
    TextoFiltro = "Select * from OrdemServico where IDproducao = " & ListaOS.SelectedItem.ListSubItems(1)
End If
Set TBOS = CreateObject("adodb.recordset")
TBOS.Open TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
If TBOS.EOF = False Then
    Do While TBOS.EOF = False
        'Verifica tempo total de preparação e execução previsto para todas OS
        'Preparação
        If IsNull(TBOS!Tempo_prep_reaproveitado) = True Or TBOS!Tempo_prep_reaproveitado = False Then TPPREV = TPPREV + IIf(IsNull(TBOS!Preparacao), 0, TBOS!Preparacao)
        
        'Execução
        TEPSEG = TEPSEG + (IIf(IsNull(TBOS!TESegundos), 0, TBOS!TESegundos) * TBOS!Totalprod)
        TBOS.MoveNext
    Loop
    FunElapsedTime (TPPREV)
    TPPSEG = S
    
    'Verificar o percentual do tempo total por OS
    TBOS.MoveFirst
    Do While TBOS.EOF = False
        'Preparação
        FunElapsedTime (TBOS!Preparacao)
        If TPPSEG > 0 Then QuantComprado = (S / TPPSEG) * 100 Else QuantComprado = 100
        
        'Execução
        S = (IIf(IsNull(TBOS!TESegundos), 0, TBOS!TESegundos) * TBOS!Totalprod)
        If TEPSEG > 0 Then QuantComprado1 = (S / TEPSEG) * 100 Else QuantComprado1 = 100
        
        'Filtra tempo total de preparação e execução por OS e maquina
        TempoTotalPrep = "00:00:00"
        TempoTotalProd = "00:00:00"
        TTOK = 0
        TTNC = 0
        TTCD = 0
        Set TBProducao = CreateObject("adodb.recordset")
        TBProducao.Open "Select * from " & NomeTabelaAp & " where idfase = " & TBOS!IDProducao & " and maquina = '" & txtMaquina.Text & "' and (CodigoDesc = 1 or CodigoDesc = 2) order by Data, Tempoinicio", Conexao, adOpenKeyset, adLockOptimistic
        If TBProducao.EOF = False Then
            Do While TBProducao.EOF = False
                If TBProducao!codigoDesc = 1 Then
                    TempoTotalPrep = TempoTotalPrep + IIf(IsNull(TBProducao!TempoTotal), 0, TBProducao!TempoTotal)
                Else
                    TempoTotalProd = TempoTotalProd + IIf(IsNull(TBProducao!TempoTotal), 0, TBProducao!TempoTotal)
                End If
                TTOK = TTOK + IIf(IsNull(TBProducao!Quantidade), 0, TBProducao!Quantidade)
                TTNC = TTNC + IIf(IsNull(TBProducao!Reprovada), 0, TBProducao!Reprovada)
                TTCD = TTCD + IIf(IsNull(TBProducao!QTCD), 0, TBProducao!QTCD)
                TBProducao.MoveNext
            Loop
        End If
        TBProducao.Close
        Produzidas = TTOK + TTNC + TTCD
        
        'Preparação
        'Calcula total de segundos utilizados preparando
        FunElapsedTime (TempoTotalPrep)
        TPUSEG = S
        
        TPUSEG = Format((TPUSEG * QuantComprado) / 100, "###,##0.0000000000")
        TPUSEGDECS = TPUSEG
        TTPUTIL = FunFormataTempo(TPUSEG)
       
        'Execução
        'Calcula total de segundos utilizados produzindo
        FunElapsedTime (TempoTotalProd)
        TTEUTILS = S
        
        TTEUTILS = Format((TTEUTILS * QuantComprado1) / 100, "###,##0.0000000000")
        TTEUTILDECS = TTEUTILS
        TTUTIL = FunFormataTempo(TTEUTILS)
        TTEUTILS = TTEUTILDECS
        
        'Soma tempo total de preparação + execução
        TTUTILSEG = TPUSEGDECS + TTEUTILS
        TempoTotalUtil = FunFormataTempo(TPUSEGDECS + TTEUTILS)
        
        'Calcula tempo total de execução por peça
        If TTEUTILS > 0 And Produzidas > 0 Then TEUSEG = Format(TTEUTILS / Produzidas, "###,##0.0000000000") Else TEUSEG = TTEUTILS
        DecimoSegundos = TEUSEG
           
        'Calcula e grava valor real do lote e por peça
        ProcCalculaValorHoraPosto
            
        Set TBProcessos = CreateObject("adodb.recordset")
        TBProcessos.Open "Select IDProducao from " & NomeTabelaAp & " where OS = " & TBOS!IDProducao & " and Maquina = '" & txtMaquina & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBProcessos.EOF = False Then
            'Calcula eficiencia
            FunElapsedTime (IIf(IsNull(TBOS!Preparacao), 0, TBOS!Preparacao))
            If S > 0 And TPUSEGDECS > 0 Then Eficiencia_prep = Format((S / TPUSEGDECS) * 100, "###,##0.00") Else Eficiencia_prep = 0
            If IIf(IsNull(TBOS!TESegundos), 0, TBOS!TESegundos) > 0 And TEUSEG > 0 Then Eficiencia_exec = Format((TBOS!TESegundos / TEUSEG) * 100, "###,##0.00") Else Eficiencia_exec = 0
            
            Set TBProducao = CreateObject("adodb.recordset")
            TBProducao.Open "Select * from Ordemservico_maq_utilizadas where OS = " & TBOS!IDProducao & " and maquina = '" & txtMaquina.Text & "'", Conexao, adOpenKeyset, adLockOptimistic
            If TBProducao.EOF = True Then TBProducao.AddNew
            TBProducao!Ordem = TBOS!Ordem
            TBProducao!OS = TBOS!IDProducao
            TBProducao!Maquina = txtMaquina.Text
            'Preparação
            TBProducao!TPUTIL = TTPUTIL 'Tempo total real de preparação do lote
            TBProducao!TPUSEG = TPUSEGDECS
            'Execução
            TBProducao!QTNC = TTNC
            TBProducao!QTOK = TTOK
            TBProducao!QTCD = TTCD
            
            TBProducao!TEUSEG = TEUSEG
            TBProducao!TEUTIL = FunFormataTempo(TEUSEG) 'Tempo total real de execução por peça
            TBProducao!TETTUTILSEG = TTUTILSEG
            TBProducao!TETTUTIL = TempoTotalUtil 'Tempo total real de preparação + tempo total real de execução do lote
            TBProducao!CRLOTE = CTTLOTE
            TBProducao!CRPECA = CTTPECA
            TBProducao!Eficiencia_prep = Eficiencia_prep
            TBProducao!Eficiencia_exec = Eficiencia_exec
            If Eficiencia_prep > 0 And Eficiencia_exec > 0 Then
                TBProducao!Eficiencia = (Eficiencia_prep + Eficiencia_exec) / 2
            ElseIf Eficiencia_prep > 0 Then
                    TBProducao!Eficiencia = Eficiencia_prep
                ElseIf Eficiencia_exec > 0 Then
                        TBProducao!Eficiencia = Eficiencia_exec
                    Else
                        TBProducao!Eficiencia = 0
            End If
            TBProducao!Totalprod = Produzidas
            Set TBProducaoFases = CreateObject("adodb.recordset")
            TBProducaoFases.Open "Select * from " & NomeTabelaAp & " where OS = " & TBOS!IDProducao & " and maquina = '" & txtMaquina.Text & "' and CodigoDesc = 3", Conexao, adOpenKeyset, adLockOptimistic
            If TBProducaoFases.EOF = False Then
                TBProducao!Pronto = "SIM"
            Else
                TBProducao!Pronto = "NÃO"
            End If
            TBProducaoFases.Close
            TBProducao.Update
            TBProducao.Close
        Else
            Conexao.Execute "DELETE from Ordemservico_maq_utilizadas where OS = " & TBOS!IDProducao & " and maquina = '" & txtMaquina & "'"
        End If
        TBProcessos.Close
        TBOS.MoveNext
    Loop
End If
TBOS.Close

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcAtualizaPrepExecUtilTotalizacao()
On Error GoTo tratar_erro
Dim TTEUTILDECS  As Long

TPPREV = 0
TEPSEG = 0
Eficiencia_prep = 0
Eficiencia_exec = 0
If Varias_OS = True Then
    TextoFiltro = "Select OrdemServico.* from OrdemServico INNER JOIN ProducaoFases_OS ON OrdemServico.ID_apontamento = ProducaoFases_OS.ID where ProducaoFases_OS.ID = " & Txt_ID_apontamento & " order by OrdemServico.IDproducao"
Else
    TextoFiltro = "Select * from OrdemServico where IDproducao = " & ListaOS.SelectedItem.ListSubItems(1)
End If
Set TBOS = CreateObject("adodb.recordset")
TBOS.Open TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
If TBOS.EOF = False Then
    Do While TBOS.EOF = False
        'Verifica tempo total de preparação e execução previsto para todas OS
        'Preparação
        If IsNull(TBOS!Tempo_prep_reaproveitado) = True Or TBOS!Tempo_prep_reaproveitado = False Then TPPREV = TPPREV + IIf(IsNull(TBOS!Preparacao), 0, TBOS!Preparacao)
        
        'Execução
        TEPSEG = TEPSEG + (IIf(IsNull(TBOS!TESegundos), 0, TBOS!TESegundos) * TBOS!Totalprod)
        TBOS.MoveNext
    Loop
    FunElapsedTime (TPPREV)
    TPPSEG = S
    
    'Verificar o percentual do tempo total por OS
    TBOS.MoveFirst
    Do While TBOS.EOF = False
        'Preparação
        FunElapsedTime (TBOS!Preparacao)
        If TPPSEG > 0 Then QuantComprado = (S / TPPSEG) * 100 Else QuantComprado = 100
        
        'Execução
        S = (IIf(IsNull(TBOS!TESegundos), 0, TBOS!TESegundos) * TBOS!Totalprod)
        If TEPSEG > 0 Then QuantComprado1 = (S / TEPSEG) * 100 Else QuantComprado1 = 100
        
        'Filtra tempo total de preparação e execução por OS
        TempoTotalPrep = "00:00:00"
        TempoTotalProd = "00:00:00"
        TTOK = 0
        TTNC = 0
        Set TBProducao = CreateObject("adodb.recordset")
        TBProducao.Open "Select * from " & NomeTabelaAp & " where idfase = " & TBOS!IDProducao & " and Usuario = '" & Operador & "' and maquina = '" & txtMaquina.Text & "' and data = '" & Format(Dataini, "Short Date") & "' and Turno = " & Turno & " and (CodigoDesc = 1 or CodigoDesc = 2) order by Data, Tempoinicio", Conexao, adOpenKeyset, adLockOptimistic
        If TBProducao.EOF = False Then
            Do While TBProducao.EOF = False
                If TBProducao!codigoDesc = 1 Then
                    TempoTotalPrep = TempoTotalPrep + IIf(IsNull(TBProducao!TempoTotal), 0, TBProducao!TempoTotal)
                Else
                    TempoTotalProd = TempoTotalProd + IIf(IsNull(TBProducao!TempoTotal), 0, TBProducao!TempoTotal)
                End If
                TTOK = TTOK + IIf(IsNull(TBProducao!Quantidade), 0, TBProducao!Quantidade)
                TTNC = TTNC + IIf(IsNull(TBProducao!Reprovada), 0, TBProducao!Reprovada)
                TTCD = TTCD + IIf(IsNull(TBProducao!QTCD), 0, TBProducao!QTCD)
                TBProducao.MoveNext
            Loop
        End If
        TBProducao.Close
        Produzidas = TTOK + TTNC
        
        'Preparação
        'Calcula total de segundos utilizados preparando
        FunElapsedTime (TempoTotalPrep)
        TPUSEG = S
        
        TPUSEG = Format((TPUSEG * QuantComprado) / 100, "###,##0.0000000000")
        TPUSEGDECS = TPUSEG
        TTPUTIL = FunFormataTempo(TPUSEG)
        
        'Execução
        'Calcula total de segundos utilizados produzindo
        FunElapsedTime (TempoTotalProd)
        TTEUTILS = S
        
        TTEUTILS = Format((TTEUTILS * QuantComprado1) / 100, "###,##0.0000000000")
        TTEUTILDECS = TTEUTILS
        TTUTIL = FunFormataTempo(TTEUTILS)
        TTEUTILS = TTEUTILDECS
        
        'Soma tempo total de preparação + execução
        TTUTILSEG = TPUSEGDECS + TTEUTILS
        TempoTotalUtil = FunFormataTempo(TPUSEGDECS + TTEUTILS)
        
        'Calcula tempo total de execução por peça
        If TTEUTILS > 0 And Produzidas > 0 Then TEUSEG = Format(TTEUTILS / Produzidas, "###,##0.0000000000") Else TEUSEG = TTEUTILS
        DecimoSegundos = TEUSEG
        
        'Calcula e grava valor real do lote e por peça
        ProcCalculaValorHoraPosto
        
        'Deleta os lançameneto de turno sem apontamento
'        Familiatext = ""
'        Set TBProcessos = CreateObject("adodb.recordset")
'        TBProcessos.Open "Select Turno from " & NomeTabelaAp & " where idfase = " & TBOS!IDProducao & " and usuario = '" & Operador & "' and maquina = '" & txtMaquina.Text & "' and data = '" & Format(Dataini, "Short Date") & "' and (codigodesc = 1 or codigodesc = 2) group by Turno", Conexao, adOpenKeyset, adLockOptimistic
'        If TBProcessos.EOF = False Then
'            Do While TBProcessos.EOF = False
'                If Familiatext = "" Then Familiatext = "and Turno <> " & TBProcessos!Turno  Else Familiatext = Familiatext & " and Turno <> " & TBProcessos!Turno
'                TBProcessos.MoveNext
'            Loop
'            Conexao.Execute "DELETE from " & NomeTabelaApTotalizacao & " where OS = " & TBOS!IDProducao & " and usuario = '" & Operador & "' and maquina = '" & txtMaquina.Text & "' " & Familiatext & " and data = '" & Format(Dataini, "Short Date") & "'"
'        End If
        
        Set TBProcessos = CreateObject("adodb.recordset")
        TBProcessos.Open "Select * from " & NomeTabelaAp & " where idfase = " & TBOS!IDProducao & " and usuario = '" & Operador & "' and maquina = '" & txtMaquina.Text & "' and Turno = " & Turno & " and data = '" & Format(Dataini, "Short Date") & "' and (codigodesc = 1 or codigodesc = 2)", Conexao, adOpenKeyset, adLockOptimistic
        If TBProcessos.EOF = False Then
            'Calcula eficiencia
            FunElapsedTime (IIf(IsNull(TBOS!Preparacao), 0, TBOS!Preparacao))
            If S > 0 And TPUSEGDECS > 0 Then Eficiencia_prep = Format((S / TPUSEGDECS) * 100, "###,##0.00") Else Eficiencia_prep = 0
            If IIf(IsNull(TBOS!TESegundos), 0, TBOS!TESegundos) > 0 And TEUSEG > 0 Then Eficiencia_exec = Format((TBOS!TESegundos / TEUSEG) * 100, "###,##0.00") Else Eficiencia_exec = 0
            
            Set TBProducao = CreateObject("adodb.recordset")
            TBProducao.Open "Select * from " & NomeTabelaApTotalizacao & " where OS = " & TBOS!IDProducao & " and Usuario = '" & Operador & "' and maquina = '" & txtMaquina.Text & "' and Turno = " & Turno & " and data = '" & Format(Dataini, "Short Date") & "'", Conexao, adOpenKeyset, adLockOptimistic
            If TBProducao.EOF = True Then TBProducao.AddNew
            TBProducao!Ordem = TBOS!Ordem
            TBProducao!OS = TBOS!IDProducao
            TBProducao!Fase = TBOS!Fase
            TBProducao!Data = Dataini
            TBProducao!Usuario = Operador
            TBProducao!Maquina = txtMaquina.Text
            TBProducao!Turno = Turno
            'Preparação
            TBProducao!Preparacao = TBOS!TempoPreparacao 'Tempo total previsto de preparação
            TBProducao!TPUTIL = TTPUTIL 'Tempo total real de preparação do lote
            'Execução
            TBProducao!Execucao = TBOS!TempoExecucao 'Tempo total previsto de execução
            TBProducao!QTOK = TTOK
            TBProducao!QTNC = TTNC
            TBProducao!QTCD = TTCD
            
            TBProducao!TEUTIL = FunFormataTempo(TEUSEG) 'Tempo total real de execução por peça
            TBProducao!TETTUTIL = TempoTotalUtil 'Tempo total real de preparação + tempo total real de execução do lote
            TBProducao!Valor_hs_prep = ValorhoraPrep * 3600
            TBProducao!Valor_hs_exec = Valorhora * 3600
            TBProducao!CRLOTE = CTTLOTE
            TBProducao!CRPECA = CTTPECA
            TBProducao!CPLOTE = TBOS!CPLOTE
            TBProducao!CPPECA = TBOS!CPPECA
            TBProducao!Eficiencia_prep = Eficiencia_prep
            TBProducao!Eficiencia_exec = Eficiencia_exec
            If Eficiencia_prep > 0 And Eficiencia_exec > 0 Then
                TBProducao!Eficiencia = (Eficiencia_prep + Eficiencia_exec) / 2
            ElseIf Eficiencia_prep > 0 Then
                    TBProducao!Eficiencia = Eficiencia_prep
                ElseIf Eficiencia_exec > 0 Then
                        TBProducao!Eficiencia = Eficiencia_exec
                    Else
                        TBProducao!Eficiencia = 0
            End If
            TBProducao!Totalprod = Produzidas
            TBProducao.Update
            TBProducao.Close
        Else
            Conexao.Execute "DELETE from " & NomeTabelaApTotalizacao & " where OS = " & TBOS!IDProducao & " and usuario = '" & Operador & "' and maquina = '" & txtMaquina.Text & "' and Turno = " & Turno & " and data = '" & Format(Dataini, "Short Date") & "'"
        End If
        TBProcessos.Close
        
        Conexao.Execute "Update " & NomeTabelaApTotalizacao & " Set pronto = '" & TBOS!Pronto & "' where OS = " & TBOS!IDProducao
        TBOS.MoveNext
    Loop
End If
TBOS.Close

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcCalculaValorHoraPosto()
On Error GoTo tratar_erro

Set TBMaquinas = CreateObject("adodb.recordset")
TBMaquinas.Open "Select * from cadmaquinas where maquina = '" & txtMaquina.Text & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBMaquinas.EOF = False Then
    'Verifica se tem acessórios no posto
    TextoFiltro = "from CadMaquinas_acessorios CMA INNER JOIN Ferramentas F ON F.ID_acessorio = CMA.ID where F.IDFase = " & TBOS!IDFase
    
    'Soma valor prep.
    Valor_Cofins_Serv = 0
    Set TBCFOP = CreateObject("adodb.recordset")
    TBCFOP.Open "Select Sum(CMA.Valor_prep * F.Quantidade) as Valor_Cofins_Serv " & TextoFiltro & " and CMA.Operacao_prep = 1", Conexao, adOpenKeyset, adLockOptimistic
    If TBCFOP.EOF = False Then
        Valor_Cofins_Serv = IIf(IsNull(TBCFOP!Valor_Cofins_Serv), 0, TBCFOP!Valor_Cofins_Serv)
    End If
    'Subtrai valor prep.
    Valor_CSLL_Prod = 0
    Set TBCFOP = CreateObject("adodb.recordset")
    TBCFOP.Open "Select Sum(CMA.Valor_prep * F.Quantidade) as Valor_CSLL_Prod " & TextoFiltro & " and CMA.Operacao_prep = 2", Conexao, adOpenKeyset, adLockOptimistic
    If TBCFOP.EOF = False Then
        Valor_CSLL_Prod = IIf(IsNull(TBCFOP!Valor_CSLL_Prod), 0, TBCFOP!Valor_CSLL_Prod)
    End If
    'Soma valor exec.
    Valor_CSLL_Serv = 0
    Set TBCFOP = CreateObject("adodb.recordset")
    TBCFOP.Open "Select Sum(CMA.Valor_exec * F.Quantidade) as Valor_CSLL_Serv " & TextoFiltro & " and CMA.Operacao_exec = 1", Conexao, adOpenKeyset, adLockOptimistic
    If TBCFOP.EOF = False Then
        Valor_CSLL_Serv = IIf(IsNull(TBCFOP!Valor_CSLL_Serv), 0, TBCFOP!Valor_CSLL_Serv)
    End If
    'Subtrai valor exec.
    Valor_INSS_Serv = 0
    Set TBCFOP = CreateObject("adodb.recordset")
    TBCFOP.Open "Select Sum(CMA.Valor_exec * F.Quantidade) as Valor_INSS_Serv " & TextoFiltro & " and CMA.Operacao_exec = 2", Conexao, adOpenKeyset, adLockOptimistic
    If TBCFOP.EOF = False Then
        Valor_INSS_Serv = IIf(IsNull(TBCFOP!Valor_INSS_Serv), 0, TBCFOP!Valor_INSS_Serv)
    End If
    
    'Verifica se tem percentual de HE no turno
    Valor_Cofins_Prod = 0
    Set TBCFOP = CreateObject("adodb.recordset")
    TBCFOP.Open "Select * from CadmaqTurnos where maquina = '" & txtMaquina.Text & "' and Diasemana = '" & DiaSemana & "' and Turno = " & Turno & " and Percentual_HoraExtra IS NOT NULL", Conexao, adOpenKeyset, adLockOptimistic
    If TBCFOP.EOF = False Then
        If TBCFOP!Percentual_HoraExtra <> "" Then Valor_Cofins_Prod = TBCFOP!Percentual_HoraExtra
    End If
    
    If IsNull(TBMaquinas!PrecoHora_Setup) = False And TBMaquinas!PrecoHora_Setup <> "" Then ValorhoraPrep = TBMaquinas!PrecoHora_Setup + Valor_Cofins_Serv - Valor_CSLL_Prod Else ValorhoraPrep = TBMaquinas!PrecoHora + Valor_Cofins_Serv - Valor_CSLL_Prod
    ValorhoraPrep = (ValorhoraPrep + ((ValorhoraPrep * Valor_Cofins_Prod) / 100)) / 3600

    Valorhora = TBMaquinas!PrecoHora + Valor_CSLL_Serv - Valor_INSS_Serv
    Valorhora = (Valorhora + ((Valorhora * Valor_Cofins_Prod) / 100)) / 3600
    
    If Produzidas > 0 Then
        CTTLOTE = Format((Valorhora * (Produzidas * TEUSEG)) + (ValorhoraPrep * TPUSEGDECS), "###,##0.00")
        If TEUSEG > 0 Or TPUSEGDECS > 0 Then CTTPECA = Format(CTTLOTE / Produzidas, "###,##0.0000000000") Else CTTPECA = 0
    Else
        CTTLOTE = Format((Valorhora * TEUSEG) + (ValorhoraPrep * TPUSEGDECS), "###,##0.00")
        If TEUSEG > 0 Or TPUSEGDECS > 0 Then CTTPECA = Format(CTTLOTE, "###,##0.0000000000") Else CTTPECA = 0
    End If
Else
    Valorhora = 0
    ValorhoraPrep = 0
End If
TBMaquinas.Close

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcGravaValoresOS()
On Error GoTo tratar_erro

If Varias_OS = True Then
    TextoFiltro = "Select OrdemServico.* from OrdemServico INNER JOIN ProducaoFases_OS ON OrdemServico.ID_apontamento = ProducaoFases_OS.ID where ProducaoFases_OS.ID = " & Txt_ID_apontamento & " order by OrdemServico.IDproducao"
Else
    TextoFiltro = "Select * from OrdemServico where IDproducao = " & ListaOS.SelectedItem.ListSubItems(1)
End If
Set TBOS = CreateObject("adodb.recordset")
TBOS.Open TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
If TBOS.EOF = False Then
    Do While TBOS.EOF = False
        Valor = 0
        Valor1 = 0
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select Sum(CRLOTE) as Valor, Sum(Totalprod) as Valor1 from ordemservico_maq_utilizadas where OS = " & TBOS!IDProducao, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Valor = IIf(IsNull(TBAbrir!Valor), 0, TBAbrir!Valor)
            Valor1 = IIf(IsNull(TBAbrir!Valor1), 0, TBAbrir!Valor1)
        End If
        TBAbrir.Close
        TBOS!CRLOTE = Valor
        If Valor <> 0 And Valor1 <> 0 Then TBOS!CRPECA = Valor / Valor1 Else TBOS!CRPECA = 0
        
        'Atualiza posto de trabalho na OS
        TBOS!Maquina = txtMaquina
        
        Set TBMaquinas = CreateObject("adodb.recordset")
        TBMaquinas.Open "Select * from cadmaquinas where maquina = '" & txtMaquina & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBMaquinas.EOF = False Then
            If IsNull(TBMaquinas!PrecoHora_Setup) = False And TBMaquinas!PrecoHora_Setup <> "" Then TBOS!Valor_hs_prep = TBMaquinas!PrecoHora_Setup Else TBOS!Valor_hs_prep = TBMaquinas!PrecoHora
            
            TBOS!Valor_hs_exec = TBMaquinas!PrecoHora
        End If
        TBMaquinas.Close
        
        TBOS.Update
        TBOS.MoveNext
    Loop
End If
TBOS.Close

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcCalculaEficienciaOF(Ordem As Long, Quantidade As Long)
On Error GoTo tratar_erro

TEPSEG = 0
Eficiencia_prep = 0
Eficiencia_exec = 0

'Verifica qtde de peças produzidas na ordem
Produzidas = IIf(IsNull(TBProducao!QuantProd), 0, TBProducao!QuantProd)

'Verif. total previstos
TPPREV = 0
TEPSEG = 0
Set TBAbrir = CreateObject("adodb.recordset")
TBAbrir.Open "Select * FROM ORDEMSERVICO WHERE Ordem = " & Ordem & " and custos = 'true'", Conexao, adOpenKeyset, adLockOptimistic
If TBAbrir.EOF = False Then
    Do While TBAbrir.EOF = False
        TPPREV = TPPREV + IIf(IsNull(TBAbrir!Preparacao), 0, TBAbrir!Preparacao)
        TEPSEG = TEPSEG + IIf(IsNull(TBAbrir!TESegundos), 0, TBAbrir!TESegundos)
        TBAbrir.MoveNext
    Loop
End If
TBAbrir.Close

FunElapsedTime (TPPREV)
TPPSEG = S

'DecimoSegundos = TEPSEG
'TBProducao!TPP = FunFormataTempo(DecimoSegundos)

If Quantidade <> 0 Then TBProducao!cpp = IIf(IsNull(TBProducao!CTTPrev), 0, TBProducao!CTTPrev) / Quantidade

If TPPSEG > 0 And TPUSEG > 0 Then Eficiencia_prep = Format((TPPSEG / TPUSEG) * 100, "###,##0.00")
If TEPSEG > 0 And TEUSEG > 0 Then Eficiencia_exec = Format((TEPSEG / TEUSEG) * 100, "###,##0.00")

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcGravarEvento()
On Error GoTo tratar_erro
Dim qtdeNC_ordem As Double

'If Individual = False Then
'    TOK = 0
'End If

'If OrdemRastreavel = False Then
    TOK = 0
    TNC = 0
    TCD = 0
'End If


TotalAprovado = 0
TotalCondicional = 0
TotalNaoconforme = 0

'If Evento = 1 Then
'    If USMsgBox("Gostaria de apontar alguma quantidade no evento de setup de posto de trabalho", vbYesNo, "CAPRIND v5.0") = vbYes Then
'        frmProducao_qtde.Show 1
'    End If
'End If

If UltimoMaq = txtMaquina And Ultimo = 2 And QT_Final = False And Individual = False Then
    If frmProducao_qtde.Visible = False Then
        frmProducao_qtde.Show 1
    End If

    If Gravar = False Then
        Exit Sub
    End If
    
ElseIf UltimoMaq = txtMaquina And Ultimo = 1 And QT_Final = False And Individual = False Then
    If frmProducao_qtde.Visible = False Then
        frmProducao_qtde.Show 1
    End If

    If Gravar = False Then
        Exit Sub
    End If
    
Else
    If Individual = False Then
        If Evento = 3 And FunVerifQtdeOSProcControlado(TBOS!IDProducao) = False Then
            Gravar = False
            Exit Sub
        End If
    End If
End If

'===============================================================================================
' Se for Fim de Produção e for sem rastreabilidade abre pra indicar somente itens não conforme
'===============================================================================================
If Evento = 3 And QT_Final = True And Individual = False Then
    frmProducao_qtdeNC.Show 1
Else
    QTOK = frmNumeroSerieOK.txtTOK
    QTNC = frmNumeroSerieOK.txtTNC
End If
'===============================================================================================

Set TBApontamento = CreateObject("adodb.recordset")
TBApontamento.Open "Select * from " & NomeTabelaAp & " where IDFase = " & TBOS!IDProducao & " and Maquina = '" & txtMaquina & "' and CODIGODESC = " & Ultimo & " order by Data, Tempoinicio", Conexao, adOpenKeyset, adLockOptimistic
If TBApontamento.EOF = True Then


    
    TBApontamento.AddNew
    TBApontamento!Ordem = TBOS!Ordem
    TBApontamento!Fase = TBOS!Fase
    TBApontamento!Maquina = txtMaquina.Text
    

    
    TBApontamento!Quantidade = TOK
    TBApontamento!Reprovada = TNC
    TBApontamento!Usuario = Operador
    TBApontamento!Descricao = DescEvento
    TBApontamento!codigoDesc = Evento
    TBApontamento!TempoInicio = Hora_apontamento
    TBApontamento!TempoFinal = 0
    TBApontamento!TempoTotal = 0
    TBApontamento!TempoTotalSeg = 0
    TBApontamento!Dias = 0
    Evento = TBApontamento!codigoDesc
    TBApontamento!Pronto = "NÃO"
    TBApontamento!Preparacao = TBOS!Preparacao
    TBApontamento!Execucao = TBOS!Execucao
    TBApontamento!Data = Date
    TBApontamento!Quant = TBOS!Quantidade
    TBApontamento!IDFase = TBOS!IDProducao
    TBApontamento!OS = TBOS!IDProducao
    TBApontamento!Turno = Turno
    TBApontamento.Update
    IDApontamento = TBApontamento!IDProducao
    IDProducao = TBApontamento!IDProducao
Else
    'Após filtrar move o ponteiro para o ultimo registro
    TBApontamento.MoveLast
    TempoInicio = TBApontamento!TempoInicio
    TempoFinal = Hora_apontamento
    TempoTotal = TempoFinal - TempoInicio
    FunElapsedTime (TempoTotal)
    TempoTotal = Format(TempoTotal, "hh:mm:ss")
    'HaBilita o modo de edição
    TBApontamento!TempoFinal = TempoFinal
    TBApontamento!TempoTotal = TempoTotal
    TBApontamento!TempoTotalSeg = S
    TBApontamento!Dias = D
    
    If Individual = False Then
    TBApontamento!Quantidade = TOK
    TBApontamento!Reprovada = TNC
    TBApontamento!QTCD = TCD
    End If
    
    TBApontamento.Update

    TBApontamento.AddNew
    TBApontamento!Ordem = TBOS!Ordem
    TBApontamento!Fase = TBOS!Fase
    TBApontamento!Maquina = txtMaquina.Text
    TBApontamento!Usuario = Operador
    TBApontamento!Descricao = DescEvento
    TBApontamento!codigoDesc = Evento
    TBApontamento!TempoInicio = Hora_apontamento
    TBApontamento!TempoFinal = 0
    TBApontamento!TempoTotal = 0
    TBApontamento!TempoTotalSeg = 0
    TBApontamento!Dias = 0
    TBApontamento!Pronto = "NÃO"
    TBApontamento!Preparacao = TBOS!Preparacao
    TBApontamento!Execucao = TBOS!Execucao
    TBApontamento!Data = Date
    TBApontamento!Quant = TBOS!Quantidade
    Evento = TBApontamento!codigoDesc
    TBApontamento!Dias = 0
    TBApontamento!IDFase = TBOS!IDProducao
    TBApontamento!OS = TBOS!IDProducao
    TBApontamento!Turno = Turno
    TBApontamento.Update
    IDApontamento = TBApontamento!IDProducao
    IDProducao = TBApontamento!IDProducao
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

End If

'TBProducao.Close

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcVerificaEvento()
On Error GoTo tratar_erro
Dim Pronto As String 'OK

If Gravar = False Then Exit Sub
If Evento = 3 Then Pronto = "SIM" Else Pronto = "NÃO"
'Filtra todos os eventos da(s) OS('s) na tabela producaofases para marcar como fase pronta
If Varias_OS = True Then
    Conexao.Execute "Update " & NomeTabelaAp & " Set " & NomeTabelaAp & ".Pronto = '" & Pronto & "' from " & NomeTabelaAp & " INNER JOIN OrdemServico ON " & NomeTabelaAp & ".IDFase = OrdemServico.IDproducao where OrdemServico.ID_apontamento = " & Txt_ID_apontamento
    
    Set TBOS = CreateObject("adodb.recordset")
    TBOS.Open "Select * from OrdemServico where ID_apontamento = " & Txt_ID_apontamento, Conexao, adOpenKeyset, adLockOptimistic
    If TBOS.EOF = False Then
        Do While TBOS.EOF = False
            ProcMudaStatusOS TBOS!IDProducao, Pronto
            TBOS.MoveNext
        Loop
        
        TBOS.MoveFirst
        Do While TBOS.EOF = False
            ProcMudaStatusOrdem TBOS!Ordem
            TBOS.MoveNext
        Loop
    End If
    TBOS.Close
Else
    Conexao.Execute "Update " & NomeTabelaAp & " Set pronto = '" & Pronto & "' where idfase = " & ListaOS.SelectedItem.ListSubItems(1)
    ProcMudaStatusOS ListaOS.SelectedItem.ListSubItems(1), Pronto
    ProcMudaStatusOrdem ListaOS.SelectedItem
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcMudaStatusOS(OS As Long, Pronto As String)
On Error GoTo tratar_erro

Conexao.Execute "Update ordemservico Set Pronto = '" & Pronto & "', Dataconclusao = '" & Date & "', Status = 'Concluída' where IDproducao = " & OS

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcMudaStatusOrdem(Ordem As Long)
On Error GoTo tratar_erro

'Checa se todas as ordens de servicos com processo da ordem foram executadas e da baixa na ordem
Set TBOrdemServico = CreateObject("adodb.recordset")
TBOrdemServico.Open "Select * from ordemservico where Ordem = " & Ordem & " and pronto = 'NÃO'", Conexao, adOpenKeyset, adLockOptimistic
If TBOrdemServico.EOF = True Then
    Set TBFases = CreateObject("adodb.recordset")
    TBFases.Open "Select * from ordemservico where Ordem = " & Ordem & " order by fase, retrabalho, IDproducao", Conexao, adOpenKeyset, adLockOptimistic
    If TBFases.EOF = False Then
        TBFases.MoveLast
        Set TBOrdem = CreateObject("adodb.recordset")
        TBOrdem.Open "Select * from Producao where Ordem = " & Ordem, Conexao, adOpenKeyset, adLockOptimistic
        If TBOrdem.EOF = False Then
            If TBFases!Pronto = "SIM" Then
                TBOrdem!pronta = "SIM"
                TBOrdem!concluida = True
                If TBOrdem!Status <> "Entregue" Then TBOrdem!Status = "Concluída"
                TBOrdem!dataentrega = TBFases!Dataconclusao
'                TBOrdem!QuantProd = TOK + TNC + TCD
'                TBOrdem!QuantNC = TNC
'                TBOrdem!QTCD = TCD
            Else
                TBOrdem!pronta = "NÃO"
                TBOrdem!concluida = False
                If TBOrdem!Status <> "Entregue" Then TBOrdem!Status = "Produzindo"
                TBOrdem!dataentrega = Null
                TBOrdem!QuantProd = 0
                TBOrdem!QuantNC = 0
                TBOrdem!QTCD = 0
            End If
            TBOrdem.Update
                        
            'Verifica se todas as ordens de fabricação do produto já foram concluidas
            Set TBFiltro = CreateObject("adodb.recordset")
            TBFiltro.Open "Select * from producao_pedidos where Ordem = " & TBOrdem!Ordem & " order by IDCarteira", Conexao, adOpenKeyset, adLockOptimistic
            If TBFiltro.EOF = False Then
                Do While TBFiltro.EOF = False
                    Set TBProcessos = CreateObject("adodb.recordset")
                    TBProcessos.Open "Select Producao_pedidos.IDCarteira FROM Producao INNER JOIN Producao_pedidos ON Producao.Ordem = Producao_pedidos.Ordem where Producao_pedidos.IDCarteira = " & TBFiltro!IDCarteira & " and Producao.pronta = 'NÃO'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBProcessos.EOF = True Then
                        Set TBVendas = CreateObject("adodb.recordset")
                        TBVendas.Open "Select * from Vendas_carteira where Codigo = " & TBFiltro!IDCarteira, Conexao, adOpenKeyset, adLockOptimistic
                        If TBVendas.EOF = False Then
                            TBVendas!saida_estoque = True
                            TBVendas!dataprodsaida = TBOrdem!dataentrega
                            TBVendas.Update
                        End If
                        TBVendas.Close
                    End If
                    TBProcessos.Close
                    TBFiltro.MoveNext
                Loop
            End If
            TBFiltro.Close
        End If
        TBOrdem.Close
    End If
    TBFases.Close
End If
TBOrdemServico.Close

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcRetirarCancelarEstoque()
On Error GoTo tratar_erro
Dim PesquisarEmpenho As Boolean
Dim Peso As Double

Set TBItem = CreateObject("adodb.recordset")
TBItem.Open "Select * from projproduto where desenho = '" & TBOrdem!Desenho & "'", Conexao, adOpenKeyset, adLockReadOnly
If TBItem.EOF = False Then
    Set TBMaterial = CreateObject("adodb.recordset")
    TBMaterial.Open "Select PM.Unidade, ISNULL(P.IDcliente, 0) AS IDcliente, PM.Ordem, PM.Codigo, PM.Requisitado, PM.Total_pc, PM.Saida, PM.Valor_saida_estoque from Producao P INNER JOIN Producaomaterial PM ON PM.Ordem = P.Ordem where PM.Ordem = " & TBOrdem!Ordem & " order by PM.Posicao", Conexao, adOpenKeyset, adLockReadOnly
    If TBMaterial.EOF = False Then
        Do While TBMaterial.EOF = False

            Peso = TBMaterial!Requisitado / TBOrdem!Quant
            qtdeliberar = 0
            If ExcluirAP = False Then
                QuantComprado = 0
                QuantComprado1 = 0
                'Verifica quantidade saida
                Set TBMateriaprima = CreateObject("adodb.recordset")
                TBMateriaprima.Open "Select ISNULL(Saida, 0) as QtdeSaida, ISNULL(Saida_PC, 0) as qtdeSaidaPC from Qtde_saida_estoque_produto where Ordem = " & TBOrdem!Ordem & " AND desenho = '" & TBMaterial!Codigo & "'", Conexao, adOpenKeyset, adLockReadOnly
                If TBMateriaprima.EOF = False Then
                    QuantComprado = TBMateriaprima!QtdeSaida
                    QuantComprado1 = TBMateriaprima!QtdeSaidaPC
                End If
                TBMateriaprima.Close
                
                'Verifica quantidade apontada menos quantidade saida
                Set TBMateriaprima = CreateObject("adodb.recordset")
                TBMateriaprima.Open "Select ISNULL(TotalProd,0) as TotalProd from Ordemservico where IDProducao = " & Txt_OS, Conexao, adOpenKeyset, adLockReadOnly
                If TBMateriaprima.EOF = False Then
                    ValorConta = (Peso * (TOK + TNC + TBMateriaprima!Totalprod))
                    If ValorConta > QuantComprado Then
                        qtdeliberar = Format(ValorConta - QuantComprado, "0.0000")
                    End If
                End If
                TBMateriaprima.Close
                
                'Requisitado menos saida
                QuantComprado = IIf(IsNull(TBMaterial!Requisitado), 0, Format(TBMaterial!Requisitado - QuantComprado, "0.0000"))
                QuantComprado1 = IIf(IsNull(TBMaterial!Total_pc), 0, TBMaterial!Total_pc - QuantComprado1)
            Else
                qtdeliberar = Format(Peso * (TOK + TNC), "0.0000")
            End If
                
            PesquisarEmpenho = False
            If ExcluirAP = False Then
                'Verifica se tem empenho
                Set TBControleNF = CreateObject("adodb.recordset")
                TBControleNF.Open "Select IDestoque from Producao_NF_Consignada where Ordem = " & TBOrdem!Ordem & " AND Codinterno = '" & TBMaterial!Codigo & "'", Conexao, adOpenKeyset, adLockReadOnly
                If TBControleNF.EOF = False Then
                    PesquisarEmpenho = True
                    TextoFiltro = " where EM.ordem = " & TBOrdem!Ordem & " AND EM.Codinterno = '" & TBMaterial!Codigo & "' AND EC.Estoque_real > 0 AND (EM.quantidade - EM.qtde_saida) > 0 order by EM.id"
                    INNERJOINTEXTO = "EC.* from estoque_controle EC INNER JOIN Producao_NF_Consignada EM ON EM.idestoque = EC.idestoque"
                Else
                    TextoFiltro = " where EC.Desenho = '" & TBMaterial!Codigo & "' and EC.Estoque_real > 0 and EP.Liberado = 'SIM' AND EL.Estoque = 'False' and (EC.Consignacao = 'False' or EC.Consignacao = 'True' and EC.id_cliente = " & TBMaterial!IDCliente & " and EC.Tipodest_NFcons = 'C' or EC.Consignacao = 'True' and EC.Tipodest_NFcons = 'F') order by EC.Data, EC.IDestoque"
                    INNERJOINTEXTO = "EC.* from estoque_controle EC INNER JOIN Estoque_produtos EP ON EP.IDestoque = EC.IDestoque INNER JOIN Estoque_Localarmazenamento_criar EL ON EL.descricao = EP.local_armaz"
                End If
                TBControleNF.Close
            Else
                INNERJOINTEXTO = "EC.* from estoque_controle EC INNER JOIN Estoque_movimentacao EM ON EM.IDestoque = EC.IDestoque"
                TextoFiltro = " where EM.Desenho = '" & TBMaterial!Codigo & "' and EM.Documento = '" & TBOrdem!Ordem & "' and EM.Data = '" & Format(Date, "Short Date") & "' and (EM.Operacao = 'SAIDA_ORDEM' or EM.Operacao = 'SAIDA_ORDEM_PARCIAL') order by EC.Data desc, EC.IDestoque desc"
            End If
            
ContinuaEstoqueSemEmpenho:
       Set TBEstoque = CreateObject("adodb.recordset")
            TBEstoque.Open "Select " & INNERJOINTEXTO & TextoFiltro, Conexao, adOpenKeyset, adLockReadOnly
            Do While TBEstoque.EOF = False
                If ExcluirAP = False And PesquisarEmpenho = False Then
                    'Verifica se tem empenho para a ordem e pedido interno
                    If IsNumeric(TBEstoque!LOTE) = True Then
                        Set TBLista = CreateObject("adodb.recordset")
                        TBLista.Open "Select SUM(PP.Qtde_empenho) AS Qtde_empenho from Producao_pedidos PP INNER JOIN Producao P ON P.Ordem = PP.Ordem and P.Desenho = '" & TBMaterial!Codigo & "' where PP.OrdemEmpenho = " & TBMaterial!Ordem & " and PP.Ordem = " & TBEstoque!LOTE, Conexao, adOpenKeyset, adLockReadOnly
                        If TBLista.EOF = False Then
                            qtdeliberar = qtdeliberar - IIf(IsNull(TBLista!Qtde_empenho), 0, TBLista!Qtde_empenho)
                        End If
                        TBLista.Close
                    End If
                End If
                If qtdeliberar <= 0 Then GoTo Pula
                
                Qtd = TBEstoque!estoque_real
                If ExcluirAP = False Then
                    'Verifica se este RE já está empenhado
                    Qtd = Qtd - FunVerificaQtdeEmpenhoREOrdem("IDestoque = " & TBEstoque!IDestoque & " and Ordem <> " & TBOrdem!Ordem, False)
                    Qtd = Qtd - FunVerificaQtdeEmpenhoREPI("ID_estoque = " & TBEstoque!IDestoque)
                    
                    If Qtd <= 0 Then
                        quantnovo = FunVerificaQtdeEmpenhoREOrdem("IDestoque = " & TBEstoque!IDestoque & " and Ordem = " & TBOrdem!Ordem, False)
                        If quantnovo > TBEstoque!estoque_real Or quantnovo = 0 Then GoTo Proximo Else Qtd = TBEstoque!estoque_real
                    End If
                    
                    If PesquisarEmpenho = True Then
                        SaqueValorTotal = FunVerificaQtdeEmpenhoREOrdem("IDestoque = " & TBEstoque!IDestoque & " and Ordem = " & TBOrdem!Ordem, False)
                        'Verifica se o total em estoque é maior que o empenho, assim ele não retira saldo a mais que o empenho
                        If Qtd > SaqueValorTotal Then Qtd = SaqueValorTotal
                    End If
                End If
                
          Set TBProduto = CreateObject("adodb.recordset")
             TBProduto.Open "Select * from Estoque_movimentacao where IDestoque = " & TBEstoque!IDestoque & " and Documento = '" & TBOrdem!Ordem & "' and Data = '" & Format(Date, "Short Date") & "' and (Operacao = 'SAIDA_ORDEM' or Operacao = 'SAIDA_ORDEM_PARCIAL')", Conexao, adOpenKeyset, adLockOptimistic
                If TBProduto.EOF = True Then
                    TBProduto.AddNew
                    If ExcluirAP = False Then
                        If qtdeliberar >= Qtd Then Qtd_Prog = Qtd Else Qtd_Prog = qtdeliberar
                        TBProduto!Saida = Qtd_Prog
                    Else
                        If qtdeliberar >= TBProduto!Saida Then Qtd_Prog = TBProduto!Saida Else Qtd_Prog = qtdeliberar
                        TBProduto!Saida = Format(TBProduto!Saida - Qtd_Prog, "0.0000")
                    End If
                    QuantEmpenhoPC = FunCalculaQtdePCKG(TBEstoque!estoque_real, IIf(IsNull(TBEstoque!estoque_real_PC), 0, TBEstoque!estoque_real_PC), TBProduto!Saida, True)
                    TBProduto!Saida_PC = QuantEmpenhoPC
                Else
                    If ExcluirAP = False Then
                        If qtdeliberar >= Qtd Then Qtd_Prog = Qtd Else Qtd_Prog = qtdeliberar
                        TBProduto!Saida = TBProduto!Saida + Qtd_Prog
                    Else
                        If qtdeliberar >= TBProduto!Saida Then Qtd_Prog = TBProduto!Saida Else Qtd_Prog = qtdeliberar
                        TBProduto!Saida = Format(TBProduto!Saida - Qtd_Prog, "0.0000")
                    End If
                    QuantEmpenhoPC = Format(FunProcCalculaQtdePC(TBEstoque!Desenho, TBProduto!Saida, True, TBMaterial!Unidade), "0.0000")
                    TBProduto!Saida_PC = QuantEmpenhoPC
                End If
                TBProduto!Destino = "Interno"
                TBProduto!Terceiros = False
                TBProduto!Documento = TBOrdem!Ordem
                TBProduto!LOTE = TBEstoque!LOTE
                TBProduto!Desenho = TBEstoque!Desenho
                TBProduto!Data = Date
                TBProduto!Descricao = TBEstoque!Descricao
                TBProduto!Familia = TBEstoque!Classe
                TBProduto!requisitante = Operador
                TBProduto!Responsavel = Operador
                TBProduto!IDestoque = TBEstoque!IDestoque
                TBProduto!Ordem = TBOrdem!Ordem
                TBProduto!OE = TBOrdem!Ordem
                
                'Atualiza valor do material no estoque
                TBProduto!VlrUnit = 0 'IIf(IsNull(TBEstoque!valor_unitario), 0, Format(TBEstoque!valor_unitario, "0.0000000000"))
                TBProduto!VlrTotal = 0 'Format(TBProduto!Saida * TBProduto!VlrUnit, "0.00")
                
                'verifica se a quantidade retirada e menor q a quant. solicitada
                QuantEmpenhoPC = FunCalculaQtdePCKG(TBProduto!Saida, IIf(IsNull(TBProduto!Saida_PC), 0, TBProduto!Saida_PC), Qtd_Prog, True)
                If ExcluirAP = False Then
                    If Qtd_Prog >= QuantComprado Or IsNull(TBMaterial!Total_pc) = False And QuantEmpenhoPC >= QuantComprado1 Then
                        TBProduto!Operacao = "SAIDA_ORDEM"
                    Else
                        TBProduto!Operacao = "SAIDA_ORDEM_PARCIAL"
                    End If
                    TBProduto.Update
                    
                    'Centro de custo
                    IDpedido = TBProduto!IDestoque
                    IDLista = TBProduto!IDoperacao
                    Set TBAbrir = CreateObject("adodb.recordset")
                    TBAbrir.Open "Select * from projproduto where desenho = '" & TBEstoque!Desenho & "'", Conexao, adOpenKeyset, adLockReadOnly
                    If TBAbrir.EOF = False Then
                        codproduto = TBAbrir!codproduto
                        IDAntigo = IIf(IsNull(TBAbrir!ID_PC), 0, TBAbrir!ID_PC)
                    End If
                    TBAbrir.Close
                    
                    ProcCriaCreditoCCProdutoItem
                    
                    QuantComprado = Format(QuantComprado - Qtd_Prog, "0.0000")
                    QuantComprado1 = QuantComprado1 - QuantEmpenhoPC
                Else
                    TBProduto.Update
                    
                    'Centro de custo
                    Conexao.Execute "DELETE from CC_realizado where ID_estoque = " & TBProduto!IDoperacao
                    
                    If TBProduto!Saida <= 0 And TBProduto.EOF = False Then TBProduto.Delete
                End If
                
                Saida = 0
                Saida_PC = 0
                Valor_total = 0
                Set TBFiltro = CreateObject("adodb.recordset")
                TBFiltro.Open "Select Sum(Saida) as Saida, Sum(ISNULL(Saida_PC, 0)) as Saida_PC, Sum(VlrTotal) as Valortotal from Estoque_movimentacao where Documento = '" & TBOrdem!Ordem & "' and Desenho = '" & TBMaterial!Codigo & "' and (Operacao = 'SAIDA_ORDEM' or Operacao = 'SAIDA_ORDEM_PARCIAL')", Conexao, adOpenKeyset, adLockReadOnly
                If TBFiltro.EOF = False Then
                    Saida = IIf(IsNull(TBFiltro!Saida), 0, TBFiltro!Saida)
                    Saida_PC = IIf(IsNull(TBFiltro!Saida_PC), 0, TBFiltro!Saida_PC)
                    Valor_total = IIf(IsNull(TBFiltro!Valortotal), 0, TBFiltro!Valortotal)
                End If
                TBFiltro.Close
                If Saida = 0 Then
                    StatusTexto = "NÃO"
                ElseIf Saida >= TBMaterial!Requisitado Then
                        StatusTexto = "SIM"
                    Else
                        StatusTexto = "PARCIAL"
                End If
                
                NovoValor = Replace(Valor_total, ",", ".")
                Conexao.Execute "UPDATE Producaomaterial Set Valor_saida_estoque = " & NovoValor & ", Saida = '" & StatusTexto & "' where Ordem = " & TBOrdem!Ordem & " and Codigo = '" & TBMaterial!Codigo & "'"
                
                'Atualiza qtde. de saída do empenho da ordem
                Saida = 0
                Saida_PC = 0
                Set TBFiltro = CreateObject("adodb.recordset")
                TBFiltro.Open "Select Sum(Saida) as Saida, Sum(ISNULL(Saida_PC, 0)) as Saida_PC from Estoque_movimentacao where IDestoque = " & TBEstoque!IDestoque & " and Documento = '" & TBOrdem!Ordem & "' and Desenho = '" & TBMaterial!Codigo & "' and (Operacao = 'SAIDA_ORDEM' or Operacao = 'SAIDA_ORDEM_PARCIAL')", Conexao, adOpenKeyset, adLockReadOnly
                If TBFiltro.EOF = False Then
                    Saida = IIf(IsNull(TBFiltro!Saida), 0, TBFiltro!Saida)
                    Saida_PC = IIf(IsNull(TBFiltro!Saida_PC), 0, TBFiltro!Saida_PC)
                End If
                TBFiltro.Close
                NovoValor = Replace(Saida, ",", ".")
                NovoValor1 = Replace(Saida_PC, ",", ".")
                Conexao.Execute "UPDATE Producao_NF_Consignada Set Qtde_saida = " & NovoValor & " where IDestoque = " & TBEstoque!IDestoque & " and Ordem = " & TBOrdem!Ordem & " and Codinterno = '" & TBMaterial!Codigo & "'"
                Conexao.Execute "UPDATE Producao_NF_Consignada Set Qtde_saida_PC = " & NovoValor1 & " where IDestoque = " & TBEstoque!IDestoque & " and Ordem = " & TBOrdem!Ordem & " and Codinterno = '" & TBMaterial!Codigo & "' and Quantidade_PC IS NOT NULL and Quantidade_PC > 0"
                
                qtdeliberada = 0
                qtdeliberada_PC = 0
                QtdeSaida = 0
                QtdeSaida_PC = 0
                If TBItem!Estoque = True Then
                    Set TBProduto = CreateObject("adodb.recordset")
                    TBProduto.Open "Select Sum(Entrada) as qtdeliberada, Sum(ISNULL(Entrada_PC, 0)) as qtdeliberada_PC, Sum(Saida) as QtdeSaida, Sum(ISNULL(Saida_PC ,0)) as QtdeSaida_PC from Estoque_movimentacao where IdEstoque = " & TBEstoque!IDestoque, Conexao, adOpenKeyset, adLockOptimistic
                    If TBProduto.EOF = False Then
                        qtdeliberada = IIf(IsNull(TBProduto!qtdeliberada), 0, TBProduto!qtdeliberada)
                        qtdeliberada_PC = IIf(IsNull(TBProduto!qtdeliberada_PC), 0, TBProduto!qtdeliberada_PC)
                        QtdeSaida = IIf(IsNull(TBProduto!QtdeSaida), 0, TBProduto!QtdeSaida)
                        QtdeSaida_PC = IIf(IsNull(TBProduto!QtdeSaida_PC), 0, TBProduto!QtdeSaida_PC)
                    End If
                End If
                TBProduto.Close
                
                NovoValor = Replace(qtdeliberada - QtdeSaida, ",", ".")
                NovoValor1 = Replace(qtdeliberada_PC - QtdeSaida_PC, ",", ".")
                Conexao.Execute "Update Estoque_movimentacao Set estoque_venda = " & NovoValor & " where IDestoque = " & TBEstoque!IDestoque & " and Documento = '" & TBOrdem!Ordem & "' and Data = '" & Format(Date, "Short Date") & "' and (Operacao = 'SAIDA_ORDEM' or Operacao = 'SAIDA_ORDEM_PARCIAL')"
                Conexao.Execute "Update Estoque_controle Set Estoque_real = " & NovoValor & ", estoque_venda = " & NovoValor & ", estoque_real_PC = " & NovoValor1 & " where IDestoque = " & TBEstoque!IDestoque & ""
                Conexao.Execute "Update Estoque_controle Set Valor_Total = ROUND(valor_unitario * Estoque_real, 3) where IDestoque = " & TBEstoque!IDestoque & ""
                                            
                qtdeliberar = qtdeliberar - Qtd_Prog
Proximo:
                TBEstoque.MoveNext
            Loop
            If TBEstoque.EOF = True And qtdeliberar > 0 And ExcluirAP = False And PesquisarEmpenho = True Then
                PesquisarEmpenho = False
                TextoFiltro = " where EC.Desenho = '" & TBMaterial!Codigo & "' and EC.Estoque_real > 0 and EP.Liberado = 'SIM' and (EC.Consignacao = 'False' or EC.Consignacao = 'True' and EC.id_cliente = " & TBMaterial!IDCliente & " and EC.Tipodest_NFcons = 'C' or EC.Consignacao = 'True' and EC.Tipodest_NFcons = 'F') order by EC.Data, EC.IDestoque"
                INNERJOINTEXTO = "EC.* from estoque_controle EC INNER JOIN Estoque_produtos EP ON EP.IDestoque = EC.IDestoque"
                GoTo ContinuaEstoqueSemEmpenho
            End If
Pula:
            TBEstoque.Close
            TBMaterial.MoveNext
        Loop
    End If
    TBMaterial.Close
End If
TBItem.Close

'Custo material
Set TBEstoque = CreateObject("adodb.recordset")
TBEstoque.Open "Select Sum(EM.VlrTotal) as Valor from Estoque_movimentacao EM INNER JOIN Estoque_controle EC ON EC.IDestoque = EM.IDestoque where EM.Documento = '" & TBOrdem!Ordem & "' and EC.Consignacao = 'False' and (EM.Operacao = 'SAIDA_ORDEM' or EM.Operacao = 'SAIDA_ORDEM_PARCIAL')", Conexao, adOpenKeyset, adLockReadOnly
If TBEstoque.EOF = False Then
    TBOrdem!CTMaterial = Format(IIf(IsNull(TBEstoque!Valor), 0, TBEstoque!Valor), "0.00")
    TBOrdem.Update
End If
TBEstoque.Close

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcCriaCreditoCCProdutoItem()
On Error GoTo tratar_erro

Set TBFiltro = CreateObject("adodb.recordset")
TBFiltro.Open "Select * from projproduto where Codproduto = " & codproduto & " and ID_CC is not null", Conexao, adOpenKeyset, adLockOptimistic
If TBFiltro.EOF = False Then
    Set TBFI = CreateObject("adodb.recordset")
    TBFI.Open "Select * from CC_realizado where ID_estoque = " & IDLista & " and ID_CC = " & TBFiltro!ID_CC, Conexao, adOpenKeyset, adLockOptimistic
    If TBFI.EOF = True Then TBFI.AddNew
    ProcEnviaDadosCCProdutoItem TBFiltro!ID_CC
    TBFI.Update
    
    'Grava movimentação no centro consolidado
    Set TBAfericao = CreateObject("adodb.recordset")
    TBAfericao.Open "Select * from Usuarios_Setor_Consolidacao where ID_CC_consolidado = " & TBFiltro!ID_CC, Conexao, adOpenKeyset, adLockOptimistic
    If TBAfericao.EOF = False Then
        Do While TBAfericao.EOF = False
            Set TBFI = CreateObject("adodb.recordset")
            TBFI.Open "Select * from CC_realizado where ID_estoque = " & IDLista & " and ID_CC = " & TBAfericao!ID_CC, Conexao, adOpenKeyset, adLockOptimistic
            If TBFI.EOF = True Then TBFI.AddNew
            ProcEnviaDadosCCProdutoItem TBAfericao!ID_CC
            TBFI.Update
            
            Set TBCiclo = CreateObject("adodb.recordset")
            TBCiclo.Open "Select * from Usuarios_Setor_Consolidacao where ID_CC_consolidado = " & TBAfericao!ID_CC, Conexao, adOpenKeyset, adLockOptimistic
            If TBCiclo.EOF = False Then
                Do While TBCiclo.EOF = False
                    Set TBFI = CreateObject("adodb.recordset")
                    TBFI.Open "Select * from CC_realizado where ID_estoque = " & IDLista & " and ID_CC = " & TBCiclo!ID_CC, Conexao, adOpenKeyset, adLockOptimistic
                    If TBFI.EOF = True Then TBFI.AddNew
                    ProcEnviaDadosCCProdutoItem TBCiclo!ID_CC
                    TBFI.Update
                    TBCiclo.MoveNext
                Loop
            End If
            TBCiclo.Close
            
            TBAfericao.MoveNext
        Loop
    End If
    TBAfericao.Close
End If
TBFiltro.Close

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcEnviaDadosCCProdutoItem(ID_CC As Long)
On Error GoTo tratar_erro

Valor3 = IIf(IsNull(TBEstoque!valor_unitario), 0, TBEstoque!valor_unitario)
Valor_Cofins_Serv = TBProduto!Saida
Valor = Format(Valor3 * Valor_Cofins_Serv, "###,##0.00")

TBFI!Valor = Valor
TBFI!Data = mskdata
TBFI!Responsavel = PubUsuario
TBFI!ID_empresa = TBEstoque!ID_empresa
TBFI!Operacao = "Crédito"
TBFI!ID_estoque = IDLista
TBFI!ID_CC = ID_CC
TBFI!Cod_produto = codproduto
TBFI!ID_PC = IDAntigo
TBFI!Bloqueado = False

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcEntrarCancelarEstoque()
On Error GoTo tratar_erro

Valortotal = 0
quantestoque = 0

'============================================================================
' Verifica se o item controla estoque e busca unidade do produto, familia
'============================================================================
Set TBItem = CreateObject("adodb.recordset")
TBItem.Open "Select * from projproduto where desenho = '" & TBOrdem!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBItem.EOF = False Then
    If TBItem!Estoque = True Then
        ControlaEstoque = True
    Else
        ControlaEstoque = False
    End If
    
    UN = TBItem!Unidade
    Familia = TBItem!Classe
End If
TBItem.Close

'================================================
' Se o item fro submetido a estoque
'================================================
If ControlaEstoque = True Then
'================================================
' Busca local de armazanamento padrão
'================================================
LATexto = "ESTOQUE PADRÃO"

Set TBCFOP = CreateObject("adodb.recordset")
TBCFOP.Open "SELECT Descricao FROM Estoque_Localarmazenamento_criar WHERE PadraoOrdem = 'True'", Conexao, adOpenKeyset, adLockReadOnly
If TBCFOP.EOF = False Then
    LATexto = TBCFOP!Descricao
Else
    Set TBCFOP = CreateObject("adodb.recordset")
    TBCFOP.Open "Select ELAC.Descricao FROM Estoque_Localarmazenamento_criar ELAC INNER JOIN Estoque_Localarmazenamento ELA ON ELA.idemb_locarm = ELAC.id where ELA.codinterno = '" & TBOrdem!Desenho & "'", Conexao, adOpenKeyset, adLockReadOnly
    If TBCFOP.EOF = False Then LATexto = TBCFOP!Descricao
End If
TBCFOP.Close

Permitido = False

'================================================
' Busca ficha de estoque
'================================================
Set TBEstoque = CreateObject("adodb.recordset")
TBEstoque.Open "Select * from estoque_controle where ID_empresa = " & TBOrdem!ID_empresa & " and desenho = '" & TBOrdem!Desenho & "' and lote = '" & TBOrdem!Ordem & "' and certificado = '" & 0 & "' and corrida = '" & 0 & "' and local_armaz = '" & LATexto & "'", Conexao, adOpenKeyset, adLockOptimistic
If ExcluirAP = False Or TBEstoque.EOF = False And ExcluirAP = True Then Permitido = True

If Permitido = True Then
    If TBEstoque.EOF = True Then
        TBEstoque.AddNew 'Cria ficha de estoque se não houver
    End If
    
    TBEstoque!UN = UN
    TBEstoque!Classe = Familia
    TBEstoque!ID_empresa = TBOrdem!ID_empresa
    TBEstoque!LOTE = TBOrdem!Ordem
    TBEstoque!Desenho = TBOrdem!Desenho
    TBEstoque!Descricao = TBOrdem!Produto
    TBEstoque!Data = Date
    TBEstoque!Responsavel = Operador
    TBEstoque!Certificado = 0
    TBEstoque!Corrida = 0
    TBEstoque!Status = "ENTRADA_ORDEM"
    TBEstoque!id_cliente = IIf(IsNull(TBOrdem!IDCliente), 0, TBOrdem!IDCliente)
    TBEstoque!cliente = IIf(IsNull(TBOrdem!cliente), "", TBOrdem!cliente)
    TBEstoque.Update
'=======================================================
' Estoque movimentação
'=======================================================
    Set TBProduto = CreateObject("adodb.recordset")
'    StrSQL = "Select * from Estoque_movimentacao where IDestoque = " & TBEstoque!IDestoque & " and Lote = '" & TBOrdem!Ordem & "' and Data = '" & Format(Date, "Short Date") & "' and Responsavel = '" & Operador & "' and (Operacao = 'ENTRADA_ORDEM' or Operacao = 'ENTRADA_ORDEM_PARCIAL')"
    StrSQL = "Select * from Estoque_movimentacao where IDApontamento = '" & IDApontamento & "'"
    
    TBProduto.Open StrSQL, Conexao, adOpenKeyset, adLockOptimistic
    
    If TBProduto.EOF = True Then
        TBProduto.AddNew
        TBProduto!Destino = "Interno"
        TBProduto!Terceiros = False
    End If
    TBProduto!IDApontamento = IDApontamento
    TBProduto!LOTE = TBOrdem!Ordem
    TBProduto!Documento = TBOrdem!Ordem
    TBProduto!Desenho = TBOrdem!Desenho
    TBProduto!Familia = Familia
    TBProduto!Descricao = TBOrdem!Produto
    TBProduto!Data = Date
    TBProduto!Responsavel = Operador
    
'    If ExcluirAP = False Then
        
    TBProduto!Entrada = 1 'QT_Entrada_Estoque 'Entrada
    TBProduto!Entrada_PC = 1 'QT_Entrada_Estoque 'Entrada
'    Else
'        TBProduto!Entrada = TBProduto!Entrada - QT_Entrada_Estoque
'        TBProduto!Entrada_PC = TBProduto!Entrada_PC - QT_Entrada_Estoque
'    End If
    ProcEmpenharREAutomOrdem TBEstoque!IDestoque, Entrada, TBEstoque!LOTE, Date, Operador, TBOrdem!Desenho, ExcluirAP
    
'   ' Qtde = TBOrdem!Quant
'   ' Entrada = quantestoque
   ' If Entrada >= Qtde Then
    TBProduto!Operacao = "ENTRADA_ORDEM"
 '   ElseIf Entrada < Qtde Then
  '          TBProduto!Operacao = "ENTRADA_ORDEM_PARCIAL"
    'End If
    
    TBProduto!IDestoque = TBEstoque!IDestoque
    TBProduto.Update
    
    If ExcluirAP = False Then
        'Exclui o empenho no produto em estoque para o pedido
        Conexao.Execute "DELETE from Estoque_Controle_Empenho_Vendas where ID_estoque = " & TBProduto!IDestoque
        ProcEmpenhaProdutoeAtualQtdeEntEmpOrdem TBProduto!LOTE, TBProduto!Desenho, TBEstoque!estoque_real, TBProduto!IDestoque
    Else
        If TBProduto!Entrada <= 0 And TBProduto.EOF = False Then
            'Exclui o empenho no produto em estoque para o pedido
            Conexao.Execute "DELETE from Estoque_Controle_Empenho_Vendas where ID_estoque = " & TBProduto!IDestoque
            
            Conexao.Execute "DELETE from Estoque_movimentacao where IDestoque = " & TBEstoque!IDestoque & " and Lote = '" & TBOrdem!Ordem & "' and Data = '" & Format(Date, "Short Date") & "' and (Operacao = 'ENTRADA_ORDEM' or Operacao = 'ENTRADA_ORDEM_PARCIAL')"
            
            ProcAtualizaQtdeEntEmpProd TBEstoque!LOTE, TBEstoque!Desenho
        End If
    End If
    
    qtdeliberada = 0
    QtdeSaida = 0
    QtdeEstoque = 0

'=========================================================
' Verifica estoque para vendas
'=========================================================

        Set TBProduto = CreateObject("adodb.recordset")
        TBProduto.Open "Select * from Estoque_movimentacao where IdEstoque = " & TBEstoque!IDestoque, Conexao, adOpenKeyset, adLockOptimistic
        If TBProduto.EOF = True Then
            TBEstoque.Delete
        Else
            Set TBProduto = CreateObject("adodb.recordset")
            TBProduto.Open "Select Sum(ISNULL(Entrada, 0)) as qtdeliberada, Sum(ISNULL(Saida, 0)) as QtdeSaida from Estoque_movimentacao where IdEstoque = " & TBEstoque!IDestoque, Conexao, adOpenKeyset, adLockOptimistic
            If TBProduto.EOF = False Then
                qtdeliberada = IIf(IsNull(TBProduto!qtdeliberada), 0, TBProduto!qtdeliberada)
                QtdeSaida = IIf(IsNull(TBProduto!QtdeSaida), 0, TBProduto!QtdeSaida)
                NovoValor = Replace(qtdeliberada - QtdeSaida, ",", ".")
                Conexao.Execute "Update Estoque_movimentacao Set estoque_venda = " & NovoValor & " where IDestoque = " & TBEstoque!IDestoque & " and Lote = '" & TBOrdem!Ordem & "' and Data = '" & Format(Date, "Short Date") & "' and Responsavel = '" & Operador & "' and (Operacao = 'ENTRADA_ORDEM' or Operacao = 'ENTRADA_ORDEM_PARCIAL')"
            End If
        End If

    TBProduto.Close
End If
TBEstoque.Close
End If


Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcExcluirEntradaEstoque()
On Error GoTo tratar_erro

Conexao.Execute ("Delete from estoque_Movimentacao where IDApontamento = '" & IDProducao & "'")

'===============================================
' Verifica se não existem mais apontamento e exclui a ficha de estoque (RE)
'==============================================================================
Set TBItem = CreateObject("adodb.recordset")
TBItem.Open "Select * from ProducaoFases where OS = '" & Txt_OS.Text & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBItem.EOF = True Then
    Conexao.Execute ("Delete from estoque_Controle where Lote = '" & txtordem.Text & "'")
End If
TBItem.Close


Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcExecutarEntradaEstoque()
On Error GoTo tratar_erro

Valortotal = 0
quantestoque = 0

'============================================================================
' Verifica se o item controla estoque e busca unidade do produto, familia
'============================================================================
Set TBItem = CreateObject("adodb.recordset")
TBItem.Open "Select * from projproduto where desenho = '" & TBOrdem!Desenho & "'", Conexao, adOpenKeyset, adLockOptimistic
If TBItem.EOF = False Then
    If TBItem!Estoque = True Then
        ControlaEstoque = True
    Else
        ControlaEstoque = False
    End If
    
    UN = TBItem!Unidade
    Familia = TBItem!Classe
End If
TBItem.Close

'================================================
' Se o item for submetido a estoque
'================================================
If ControlaEstoque = True Then
'================================================
' Busca local de armazanamento padrão
'================================================
LATexto = "ESTOQUE PADRÃO"

Set TBCFOP = CreateObject("adodb.recordset")
TBCFOP.Open "SELECT Descricao FROM Estoque_Localarmazenamento_criar WHERE PadraoOrdem = 'True'", Conexao, adOpenKeyset, adLockReadOnly
If TBCFOP.EOF = False Then
    LATexto = TBCFOP!Descricao
Else
    Set TBCFOP = CreateObject("adodb.recordset")
    TBCFOP.Open "Select ELAC.Descricao FROM Estoque_Localarmazenamento_criar ELAC INNER JOIN Estoque_Localarmazenamento ELA ON ELA.idemb_locarm = ELAC.id where ELA.codinterno = '" & TBOrdem!Desenho & "'", Conexao, adOpenKeyset, adLockReadOnly
    If TBCFOP.EOF = False Then LATexto = TBCFOP!Descricao
End If
TBCFOP.Close

Permitido = False

'================================================
' Busca ficha de estoque
'================================================
Set TBEstoque = CreateObject("adodb.recordset")
TBEstoque.Open "Select * from estoque_controle where ID_empresa = " & TBOrdem!ID_empresa & " and lote = '" & TBOrdem!Ordem & "'", Conexao, adOpenKeyset, adLockOptimistic
'If ExcluirAP = False Or TBEstoque.EOF = False And ExcluirAP = True Then Permitido = True

'If Permitido = True Then
    If TBEstoque.EOF = True Then
    TBEstoque.AddNew 'Cria ficha de estoque se não houver
    TBEstoque!UN = UN
    TBEstoque!Classe = Familia
    TBEstoque!ID_empresa = TBOrdem!ID_empresa
    TBEstoque!LOTE = TBOrdem!Ordem
    TBEstoque!Desenho = TBOrdem!Desenho
    TBEstoque!Descricao = TBOrdem!Produto
    TBEstoque!Data = Date
    TBEstoque!Responsavel = Operador
    TBEstoque!Certificado = 0
    TBEstoque!Corrida = 0
    TBEstoque!Status = "ENTRADA_ORDEM"
    TBEstoque!id_cliente = IIf(IsNull(TBOrdem!IDCliente), 0, TBOrdem!IDCliente)
    TBEstoque!cliente = IIf(IsNull(TBOrdem!cliente), "", TBOrdem!cliente)
    TBEstoque!local_armaz = LATexto
    TBEstoque.Update
    End If
    
'=======================================================
' Busca o ID do evento Produzindo anterior
'=======================================================
IDApontamento = 0
Set TBCiclo = CreateObject("adodb.recordset")
TBCiclo.Open "Select * from ProducaoFases where OS = " & Txt_OS.Text & " and Maquina = '" & txtMaquina & "' and usuario = '" & Operador & "' and CodigoDesc = '2' order by Data, Tempoinicio", Conexao, adOpenKeyset, adLockOptimistic
If TBCiclo.EOF = False Then
    TBCiclo.MoveLast
    IDApontamento = TBCiclo!IDProducao
End If

'=======================================================
' Realiza a entrada no estoque movimentação
'=======================================================
    Set TBProduto = CreateObject("adodb.recordset")
    StrSQL = "Select * from Estoque_movimentacao where IDApontamento = '" & IDApontamento & "'"
    
    TBProduto.Open StrSQL, Conexao, adOpenKeyset, adLockOptimistic
    
    TBProduto.AddNew
    TBProduto!Destino = "Interno"
    TBProduto!Terceiros = False
    TBProduto!IDApontamento = IDApontamento
    TBProduto!LOTE = TBOrdem!Ordem
    TBProduto!Documento = TBOrdem!Ordem
    TBProduto!Desenho = TBOrdem!Desenho
    TBProduto!Familia = Familia
    TBProduto!Descricao = TBOrdem!Produto
    TBProduto!Data = Date
    TBProduto!Responsavel = Operador
    TBProduto!Entrada = QT_Entrada_Estoque 'Entrada
    TBProduto!Entrada_PC = QT_Entrada_Estoque 'Entrada
    TBProduto!Operacao = "ENTRADA_ORDEM"
    TBProduto!IDestoque = TBEstoque!IDestoque
    TBProduto.Update
    
'====================================================================
' Verifica empenho para vendas
'====================================================================
    ProcEmpenharREAutomOrdem TBEstoque!IDestoque, Entrada, TBEstoque!LOTE, Date, Operador, TBOrdem!Desenho, ExcluirAP
'====================================================================
'Exclui o empenho no produto em estoque para o pedido
    Conexao.Execute "DELETE from Estoque_Controle_Empenho_Vendas where ID_estoque = " & TBProduto!IDestoque
    ProcEmpenhaProdutoeAtualQtdeEntEmpOrdem TBProduto!LOTE, TBProduto!Desenho, TBEstoque!estoque_real, TBProduto!IDestoque
'====================================================================
' Corrigir estoque para vendas
'====================================================================

qtdeliberada = 0
QtdeSaida = 0
QtdeEstoque = 0
    
        Set TBProduto = CreateObject("adodb.recordset")
        TBProduto.Open "Select * from Estoque_movimentacao where IdEstoque = " & TBEstoque!IDestoque, Conexao, adOpenKeyset, adLockOptimistic
        If TBProduto.EOF = True Then
            TBEstoque.Delete
        Else
            Set TBProduto = CreateObject("adodb.recordset")
            TBProduto.Open "Select Sum(ISNULL(Entrada, 0)) as qtdeliberada, Sum(ISNULL(Saida, 0)) as QtdeSaida from Estoque_movimentacao where IdEstoque = " & TBEstoque!IDestoque, Conexao, adOpenKeyset, adLockOptimistic
            If TBProduto.EOF = False Then
                qtdeliberada = IIf(IsNull(TBProduto!qtdeliberada), 0, TBProduto!qtdeliberada)
                QtdeSaida = IIf(IsNull(TBProduto!QtdeSaida), 0, TBProduto!QtdeSaida)
                NovoValor = Replace(qtdeliberada - QtdeSaida, ",", ".")
                Conexao.Execute "Update Estoque_movimentacao Set estoque_venda = " & NovoValor & " where IDestoque = " & TBEstoque!IDestoque & " and Lote = '" & TBOrdem!Ordem & "' and Data = '" & Format(Date, "Short Date") & "' and Responsavel = '" & Operador & "' and (Operacao = 'ENTRADA_ORDEM' or Operacao = 'ENTRADA_ORDEM_PARCIAL')"
            End If
        End If

    TBProduto.Close
End If
TBEstoque.Close
'End If


Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcVisualizarArq()
On Error GoTo tratar_erro

Set TBFases = CreateObject("adodb.recordset")
TBFases.Open "Select F.Caminho from Fases F INNER JOIN Ordemservico OS ON OS.IDFase = F.IDFase where OS.IDproducao = " & ListaOS.SelectedItem.ListSubItems(1) & " and F.Caminho IS NOT NULL and F.Caminho <> N''", Conexao, adOpenKeyset, adLockOptimistic
If TBFases.EOF = False Then
    ProcAbrirArquivo TBFases!Caminho
End If
TBFases.Close

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcGravar()
On Error GoTo tratar_erro
Dim qtdeSaida_Mov As Double
'Debug.Print TOK

Evento = 0
Gravar = True
'Se não foi escolhida uma descrição do código de trabalho
'Atribui um valor a variavel descricao
If Codigo_Barras = True Then
If txtCodigoBarras = "" Then
    MsgBox "Informe o código de trabalho antes de salvar.", vbExclamation
    txtCodigoBarras.SetFocus
    Exit Sub
End If

    Evento = txtCodigoBarras
    DescEvento = txtCodigoDesc
Else
If txtCodigo = "" Then
    MsgBox "Informe o código de trabalho antes de salvar.", vbExclamation
    cmbdescricao.SetFocus
    Exit Sub
End If

    Evento = txtCodigo
    DescEvento = cmbdescricao
End If
'Verifica se esta tudo informado corretamente antes de gravar o apontamento
ProcVerificacao
If Gravar = False Then Exit Sub

Hora_apontamento = Format(txtData, "dd/mm/yyyy") & " " & txtHora

'Gravar evento realizado
If Varias_OS = True Then
    TextoFiltro = "Select OS.*, P.Quant, P.Desenho, P.Retirar_estoque, ISNULL(OS.TotalProd, 0) as TotalProd from (OrdemServico OS INNER JOIN ProducaoFases_OS PFOS ON OS.ID_apontamento = PFOS.ID) INNER JOIN Producao P ON P.Ordem = OS.Ordem where PFOS.ID = " & Txt_ID_apontamento & " order by OS.IDproducao"
Else
    TextoFiltro = "Select OS.*, P.Quant, P.Desenho, P.Retirar_estoque, ISNULL(OS.TotalProd, 0) as TotalProd from OrdemServico OS INNER JOIN Producao P ON P.Ordem = OS.Ordem where OS.IDProducao = " & ListaOS.SelectedItem.ListSubItems(1)
End If

Set TBOS = CreateObject("adodb.recordset")
TBOS.Open TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
If TBOS.EOF = False Then
    Do While TBOS.EOF = False
        'Debug.Print TOK
        ProcGravarEvento
        
        If Gravar = False Then
        Exit Sub
        End If
        
        'Dados do estoque
        'Debug.Print TOK
        'Debug.Print TNC
'==============================================================================
'   Verifica o evento gravado
'==============================================================================
    ProcVerificaEvento
'==============================================================================
       If QT_Entrada_Estoque <> 0 Or TNC <> 0 Then
'        If TOK > 0 And Evento <> 2 Then
            ExcluirAP = False
            Set TBOrdem = CreateObject("adodb.recordset")
'==============================================================================
'    Verifica se a ordem é com entrada automática de estoque
'==============================================================================
            TBOrdem.Open "Select * from Producao where Ordem = " & TBOS!Ordem & " and (Entrar_estoque = 'True' Or Retirar_estoque = 'True')", Conexao, adOpenKeyset, adLockOptimistic
            If TBOrdem.EOF = False Then
                Set TBCFOP = CreateObject("adodb.recordset")
                TBCFOP.Open "Select * from ordemservico where Ordem = " & TBOrdem!Ordem & " order by fase, retrabalho, IDproducao", Conexao, adOpenKeyset, adLockOptimistic
                If TBCFOP.EOF = False Then
                    If TBOrdem!Retirar_estoque = True Then
                        TBCFOP.MoveFirst
'==============================================================================
'Verifica se é a primeira OS e retira o material do estoque
'==============================================================================
'    Retirar item do estoque ou cancelar a retirada
'==============================================================================
                        If TBCFOP!IDProducao = TBOS!IDProducao Then
                                ProcRetirarCancelarEstoque
                        End If
'==============================================================================
                    End If
                    If TBOrdem!Entrar_estoque = True Then
'==============================================================================
'   Verifica se é a última OS e entra com o produto no estoque
'==============================================================================
                        TBCFOP.MoveLast
'==============================================================================
'    Executar a entrada no estoque do item
'==============================================================================
                        If TBCFOP!IDProducao = TBOS!IDProducao And QT_Entrada_Estoque > 0 Then
                            StrSQL = "Update OrdemServico set QT_Final = 1 where idproducao = " & TBOS!IDProducao
                            Conexao.Execute (StrSQL)
                                ProcExecutarEntradaEstoque
                            QT_Entrada_Estoque = 0
                        End If
'===============================================================================
                    End If
                End If
            End If
            TBOrdem.Close
        End If
        
        TBOS.MoveNext
    Loop
End If
TBOS.Close

'Atualiza dados do formulario, e da lista de eventos
ProcAtualizaProducao False, True
'Grava o status da OS e da OF
ProcGravarStatusOSOF

Dataini = Date
ProcVerificaTurno

'Acerta cadastro na máquina
ProcAcertaCadMaquina
'Calcula e grava valor real do lote e por peça
ProcGravaValoresOS
'Atualiza dados da ordem de fabricação (Tempo total utilizado, custo por peça, custo total, etc...)
ProcAtualizaOF
'Verifica se o penultimo evento é máquina em manutenção e abre o formulário
If DescEvento = "MÁQUINA EM MANUTENÇÃO" Or DescEvento = "MANUTENÇÃO PREVENTIVA" Or DescEvento = "MANUTENÇÃO CORRETIVA" Then
    Conexao.Execute "Update manutencao_data set manutencao_data.IDProducao = " & IDApontamento & " from manutencao_data INNER JOIN manutencao on manutencao_data.idManutencao = manutencao.CODIGO where manutencao.IDmaquina = '" & txtMaquina & "' and manutencao.Controlada = 'true' and manutencao_data.status = 'Aberta' and manutencao_data.data <= '" & Format(Date, "Short Date") & "'"
End If
If Evento = 1 Or Evento = 2 Then
    Set TBProcessos = CreateObject("adodb.recordset")
    TBProcessos.Open "select * from manutencao where IDmaquina = '" & txtMaquina & "' and Controlada = 'true'", Conexao, adOpenKeyset, adLockOptimistic
    If TBProcessos.EOF = False Then
        Set TBOrdem = CreateObject("adodb.recordset")
        TBOrdem.Open "select * from manutencao_data where idManutencao = " & TBProcessos!Codigo & " and status = 'Aberta' and IDProducao <> 0 and data <= '" & Format(Date, "Short Date") & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBOrdem.EOF = False Then
            Maquina = txtMaquina
            frmManutencao.Show 1
        End If
        TBOrdem.Close
    End If
    TBProcessos.Close
End If
ProcAtualizaCodigoDesc
ProcCarregaListaOS
ProcLista12Ultimos

'Limpa os campos
If Codigo_Barras = True Then
    txtCodigoBarras = ""
    txtCodigoDesc = ""
    'txtCodigoBarras.SetFocus
Else
    txtCodigo = ""
    cmbdescricao.ListIndex = -1
End If

'================================================================
'Se for com rastreabilidade e for evento produzindo
'================================================================
If Individual = True And Evento = 2 Then
    frmNumeroSerieOK.Show 1
End If


Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcGravarStatusOSOF()
On Error GoTo tratar_erro
Dim StatusTexto As String

If Varias_OS = True Then
    TextoFiltro = "Select OrdemServico.* from OrdemServico INNER JOIN ProducaoFases_OS ON OrdemServico.ID_apontamento = ProducaoFases_OS.ID where ProducaoFases_OS.ID = " & Txt_ID_apontamento & " order by OrdemServico.IDproducao"
Else
    TextoFiltro = "Select * from OrdemServico where IDProducao = " & ListaOS.SelectedItem.ListSubItems(1)
End If
Set TBOS = CreateObject("adodb.recordset")
TBOS.Open TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
If TBOS.EOF = False Then
    Do While TBOS.EOF = False
        If ULTICOD <> 1 And ULTICOD <> 2 And ULTICOD <> 3 Then
            StatusTexto = "Aguardando"
        Else
            Select Case ULTICOD
                Case 1: StatusTexto = "Preparando"
                Case 2: StatusTexto = "Produzindo"
                Case 3: StatusTexto = "Concluída"
            End Select
        End If
        Conexao.Execute "UPDATE ordemservico Set Status = '" & StatusTexto & "' where idproducao = " & TBOS!IDProducao
        
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select Status from ordemservico where Ordem = " & TBOS!Ordem & " and Status <> 'Concluída'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = True Then
            StatusTexto = "Concluída"
            GoTo Sair1
        End If
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select Status from ordemservico where Ordem = " & TBOS!Ordem & " and Status <> 'Aguardando'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            StatusTexto = "Produzindo"
            GoTo Sair1
        End If
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select Status from ordemservico where Ordem = " & TBOS!Ordem & " and Status <> 'Preparando' and Status <> 'Produzindo' and Status <> 'Concluída'", Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            StatusTexto = "Aberta"
            GoTo Sair1
        End If
Sair1:
        Conexao.Execute "UPDATE producao Set Status = '" & StatusTexto & "' where Ordem = " & TBOS!Ordem & " and Status <> 'Entregue'"
        
        TBAbrir.Close
        TBOS.MoveNext
    Loop
End If
TBOS.Close

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcAtualizaOF()
On Error GoTo tratar_erro
Dim CRLOTE As Double, CRPECA As Double, TRLOTE As Double, Valor_Produto As Double
Dim TotalCND As Double ' Total aprovada condicional
Dim TOTALNC As Double 'Total não conforme
Dim TotalProduzida As Double 'Total produzida
Dim TotalAprovada As Double 'Total aprovada
Dim Totalprod As Double

If Varias_OS = True Then
    TextoFiltro = "Select OrdemServico.Ordem from OrdemServico INNER JOIN ProducaoFases_OS ON OrdemServico.ID_apontamento = ProducaoFases_OS.ID where ProducaoFases_OS.ID = " & Txt_ID_apontamento & " Group by OrdemServico.Ordem"
Else
    TextoFiltro = "Select Ordem from OrdemServico where IDproducao = " & ListaOS.SelectedItem.ListSubItems(1)
End If
Set TBOS = CreateObject("adodb.recordset")
TBOS.Open TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
If TBOS.EOF = False Then
    Do While TBOS.EOF = False
        CRLOTE = 0
        CRPECA = 0
        TRLOTE = 0
        TPUSEG = 0
        TEUSEG = 0
        QtdeSaida = 0
        TOTALNC = 0
        TotalCND = 0
        TotalProduzida = 0
        OF = TBOS!Ordem
        Set TBProducao = CreateObject("adodb.recordset")
        TBProducao.Open "Select * from ordemservico where Ordem = " & TBOS!Ordem & " order by fase, retrabalho desc, IDproducao", Conexao, adOpenKeyset, adLockOptimistic
        If TBProducao.EOF = False Then
            Do While TBProducao.EOF = False
                If TBProducao!Custos = True Then
                    CRLOTE = CRLOTE + IIf(IsNull(TBProducao!CRLOTE), 0, TBProducao!CRLOTE)
                    CRPECA = CRPECA + IIf(IsNull(TBProducao!CRPECA), 0, TBProducao!CRPECA)
                End If
                TRLOTE = TRLOTE + IIf(IsNull(TBProducao!TETTUTILSEG), 0, TBProducao!TETTUTILSEG)
                TPUSEG = TPUSEG + IIf(IsNull(TBProducao!TPUSEG), 0, TBProducao!TPUSEG)
                TEUSEG = TEUSEG + IIf(IsNull(TBProducao!TEUSEG), 0, TBProducao!TEUSEG)
                
                Totalprod = IIf(IsNull(TBProducao!QTOK), 0, TBProducao!QTOK)
                'TotalCND = IIf(IsNull(TBProducao!QTOK), 0, TBProducao!QTOK)
                
                Set TBAbrir = CreateObject("adodb.recordset")
                TBAbrir.Open "Select * from CQ_NC_FABRICA where OS = " & TBProducao!IDProducao & " and PARECERCQ = 'Rejeitar'", Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then
                    Set TBFiltro = CreateObject("adodb.recordset")
                    TBFiltro.Open "Select Sum(TTNC) as TotalNC from CQ_NC_FABRICA where OS = " & TBProducao!IDProducao & " and PARECERCQ = 'Rejeitar'", Conexao, adOpenKeyset, adLockOptimistic
                    If TBFiltro.EOF = False Then
                        TOTALNC = TOTALNC + IIf(IsNull(TBFiltro!TOTALNC), 0, TBFiltro!TOTALNC)
                    End If
                    TBFiltro.Close
                Else
                    QtdeSaida = QtdeSaida + IIf(IsNull(TBProducao!QTNC), 0, TBProducao!QTNC)
                End If
                TBAbrir.Close
                TBProducao.MoveNext
            Loop
        End If
        
        Totalprod = Totalprod + TOTALNC + QtdeSaida
        
        Set TBProducao = CreateObject("adodb.recordset")
        TBProducao.Open "Select * from producao where Ordem = " & TBOS!Ordem & " and Controlado_estoque = 'False'", Conexao, adOpenKeyset, adLockOptimistic
        If TBProducao.EOF = False Then
            TBProducao!QuantProd = Totalprod
            TBProducao!QuantNC = TOTALNC
            TBProducao.Update
            TBProducao!CPR = Format(CRPECA, "###,##0.0000000000")
            TBProducao!CTTReal = Format(CRLOTE, "###,##0.00")
            
'            S = TEUSEG
'            TBProducao!tpr = FunFormataTempo(S)

            If Totalprod > 0 Then TBProducao!tpr = FunFormataTempo(TRLOTE / Totalprod) Else TBProducao!tpr = FunFormataTempo(TRLOTE)
            
            TBProducao!TTTReal = FunFormataTempo(TRLOTE)
            
            ProcCalculaEficienciaOF TBOS!Ordem, TBProducao!Quant
            TBProducao!Eficiencia_prep = Eficiencia_prep
            TBProducao!Eficiencia_exec = Eficiencia_exec
            If Eficiencia_prep > 0 And Eficiencia_exec > 0 Then
                TBProducao!Eficiencia = (Eficiencia_prep + Eficiencia_exec) / 2
            ElseIf Eficiencia_prep > 0 Then
                    TBProducao!Eficiencia = Eficiencia_prep
                ElseIf Eficiencia_exec > 0 Then
                        TBProducao!Eficiencia = Eficiencia_exec
                    Else
                        TBProducao!Eficiencia = 0
            End If
            TBProducao.Update
            
            'Corrige valor do estoque
                                                    'ORDEM         QTDE. PREVISTA                                       QTDE. OK                                                    QT. PROD.(OK+NC)                                                                                         CUSTO LOTE                                        CUSTO PEÇA                                CUSTO TERCEIROS                                       CUSTO MATERIAL                                                                                         CUSTO OUTROS                                              ORDEM CONSIGNADA
            Valor_Produto = FunCalculaValorUnitOrdem(TBProducao!Ordem, IIf(IsNull(TBProducao!Quant), 0, TBProducao!Quant), IIf(IsNull(TBProducao!QuantProd), 0, TBProducao!QuantProd), IIf(IsNull(TBProducao!QuantProd), 0, TBProducao!QuantProd) + IIf(IsNull(TBProducao!QuantNC), 0, TBProducao!QuantNC), IIf(IsNull(TBProducao!CTTReal), 0, TBProducao!CTTReal), IIf(IsNull(TBProducao!CPR), 0, TBProducao!CPR), IIf(IsNull(TBProducao!CTServico), 0, TBProducao!CTServico), IIf(IsNull(TBProducao!CTMaterial), 0, TBProducao!CTMaterial), IIf(IsNull(TBProducao!CTOutras), 0, TBProducao!CTOutras), TBProducao!consignacao)
            NovoValor = Replace(Valor_Produto, ",", ".")
            Conexao.Execute "Update Estoque_Controle set valor_unitario = " & NovoValor & " where Lote = '" & TBProducao!Ordem & "' and Desenho = '" & TBProducao!Desenho & "'"
            Conexao.Execute "Update Estoque_Controle set Valor_Total = ROUND(valor_unitario * estoque_real, 2) where Lote = '" & TBProducao!Ordem & "' and Desenho = '" & TBProducao!Desenho & "'"
            Conexao.Execute "Update EM Set VlrUnit = " & NovoValor & " from Estoque_movimentacao EM INNER JOIN Estoque_controle EC ON EM.IdEstoque = EC.IdEstoque where EC.Lote = '" & TBProducao!Ordem & "' and EM.Desenho = '" & TBProducao!Desenho & "' and (EM.Operacao = 'ENTRADA_ORDEM' Or EM.Operacao = 'ENTRADA_ORDEM_PARCIAL')"
            Conexao.Execute "Update EM Set VlrTotal = ROUND(VlrUnit * Entrada, 2) from Estoque_movimentacao EM INNER JOIN Estoque_controle EC ON EM.IdEstoque = EC.IdEstoque where EC.Lote = '" & TBProducao!Ordem & "' and EM.Desenho = '" & TBProducao!Desenho & "' and (EM.Operacao = 'ENTRADA_ORDEM' Or EM.Operacao = 'ENTRADA_ORDEM_PARCIAL')"
            Conexao.Execute "Update CC Set CC.Valor = EM.VlrTotal from (CC_realizado CC INNER JOIN Estoque_movimentacao EM ON CC.ID_estoque = EM.Idoperacao) INNER JOIN Estoque_controle EC ON EC.IdEstoque = EM.IdEstoque where EC.Lote = '" & TBProducao!Ordem & "' and EC.Desenho = '" & TBProducao!Desenho & "' and (EM.Operacao = 'ENTRADA_ORDEM' Or EM.Operacao = 'ENTRADA_ORDEM_PARCIAL')"
            Conexao.Execute "Update EM set EM.VlrUnit = EM1.VlrUnit from Estoque_movimentacao EM INNER JOIN Estoque_movimentacao EM1 on EM1.IdEstoque = EM.IdEstoque where EM.Lote = '" & TBProducao!Ordem & "' and EM.Desenho = '" & TBProducao!Desenho & "' and EM.Saida > 0 and EM1.Entrada > 0"
            Conexao.Execute "Update EC set EC.valor_unitario = EM1.VlrUnit from Estoque_Controle EC INNER JOIN Estoque_movimentacao EM ON EM.IdEstoque = EC.IdEstoque INNER JOIN Estoque_movimentacao EM1 on EM1.IdEstoque = EM.IdEstoque where EM.Lote = '" & TBProducao!Ordem & "' and EM.Desenho = '" & TBProducao!Desenho & "' and EM.Saida > 0 and EM1.Entrada > 0"
            Conexao.Execute "Update Estoque_movimentacao set VlrTotal = ROUND(VlrUnit * Saida, 2) where Lote = '" & TBProducao!Ordem & "' and Desenho = '" & TBProducao!Desenho & "' and Saida > 0"
        End If
        TBProducao.Close
        TBOS.MoveNext
    Loop
End If
TBOS.Close

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub cmbdescricao_GotFocus()
txtCodigoDesc.Visible = False
End Sub

Private Sub Cmd_esc_Click()
On Error GoTo tratar_erro

Call Form_KeyDown(27, 0)

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

Private Sub Cmd_F12_Click()
On Error GoTo tratar_erro

Call Form_KeyDown(123, 0)

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Cmd_F2_Click()
On Error GoTo tratar_erro

Call Form_KeyDown(113, 0)

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Public Sub Cmd_F3_Click()
On Error GoTo tratar_erro

Call Form_KeyDown(114, 0)

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Cmd_F4_Click()
On Error GoTo tratar_erro

Call Form_KeyDown(115, 0)

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Cmd_F5_Click()

Call Form_KeyDown(116, 0)

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

Private Sub Cmd_F9_Click()
On Error GoTo tratar_erro

Call Form_KeyDown(120, 0)

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub



Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro

Select Case KeyCode
    Case vbKeyF2: ProcVisualizarArq
    Case vbKeyF3: ProcGravar
    Case vbKeyF4: ProcExcluir
    Case vbKeyF5: frmSugestoes.Show 1
    Case vbKeyF6:
        ProcLogonOut (Operador)
        Unload Me
    Case vbKeyF7: ProcCarregaListaRequisicao
    Case vbKeyF8:
        cmdF8_Click
        
    Case vbKeyEscape:
        ProcLogonOut (Operador)
        Unload Me
        'frmabrir_OS.Show
        'Codigo_Barras = True
        frmfundo.Show
    Case vbKeyF9: ProcDesvincularOS
    Case vbKeyF11: ProcAtualizaProducao True, False
    Case vbKeyF12: ProcLista12Ultimos
    Case vbKeyF8: frmOpcoesGeral2.Show 1
    Case vbKeyF10: frmCQ_sistema.Show 1
End Select

'Emula tecla enter no teclado
 
 If KeyCode = 13 Then
 Sendkeys "{tab}"
 End If
 

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
On Error GoTo tratar_erro
    
If Codigo_Barras = False Then Call FunKeyAscii(KeyAscii)
'Call FunKeyAscii(KeyAscii)

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

ProcRemoveListaResize Me
ProcRemoveObjetosResize Me
lblVersaoatual.Caption = "v" & App.Major & "." & App.Minor & "." & App.Revision

If Codigo_Barras = True Then
    Caption = "Gerprod - Coletor de dados no chão de fábrica - Código de barras - Nome do banco de dados: " & Nome_banco & " | Licenças Gerprod: " & Qtlicencas_gerprod & IIf(TemInternet = False, " | Sem internet", "")
    txtCodigo.Visible = False
    txtCodigoBarras.Visible = True
    txtCodigoDesc.Visible = True
    cmbdescricao.Visible = False
    Cmd_F2.TabStop = False
    Cmd_F3.TabStop = False
    Cmd_F4.TabStop = False
    Cmd_F6.TabStop = False
    Cmd_F11.TabStop = False
    Cmd_F9.TabStop = False
    Cmd_F12.TabStop = False
    txtCodigo.TabStop = False
    cmbdescricao.TabStop = False
Else
    Caption = "Gerprod - Coletor de dados no chão de fábrica - Nome do banco de dados: " & Nome_banco & " | Licenças Gerprod: " & Qtlicencas_gerprod & IIf(TemInternet = False, " | Sem internet", "")
    txtCodigo.Visible = True
    txtCodigoBarras.Visible = False
    txtCodigoDesc.Visible = False
    cmbdescricao.Visible = True
    Cmd_F2.TabStop = True
    Cmd_F3.TabStop = True
    Cmd_F4.TabStop = True
    Cmd_F6.TabStop = True
    Cmd_F11.TabStop = True
    Cmd_F9.TabStop = True
    Cmd_F12.TabStop = True
    Cmd_esc.TabStop = False
    txtCodigo.TabStop = True
    cmbdescricao.TabStop = True


End If

ProcCarregaListaRequisicao
txtordem.Text = Ordem
txtOS = Int(OS)


lblano.Caption = "Copyright 1999 - " & Year(Date) & " Caprind Sistemas ®. Todos os direitos reservados."
StrSQL = "update ordemServico set QT_Final = 0 Where QT_Final is null"

Conexao.Execute StrSQL

frmProducao.Refresh

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

'WindowState = 2

ProcAtualizaCodigoDesc
frmProducao.Refresh
If Codigo_Barras = False Then
    cmbdescricao.ListIndex = -1
    txtCodigo.Text = ""
    If Ap_codigo = True Then
        txtCodigo.Enabled = True
        cmbdescricao.Enabled = False
        txtCodigo.Visible = True
        txtCodigo.SetFocus
    Else
        txtCodigo.Enabled = False
        cmbdescricao.Enabled = True
        cmbdescricao.Visible = True
        txtCodigoDesc.Visible = False
    End If
        'Txt_codigoF.Visible = False
Else
    txtCodigoBarras.Enabled = True
    txtCodigoBarras.Visible = True
    txtCodigoBarras.Text = ""
    'txtCodigoBarras.SetFocus
    txtCodigoDesc.Enabled = True
    txtCodigoDesc.Visible = True
    txtCodigo.Visible = False
End If



frmProducao.Refresh

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Lista_DblClick()
On Error GoTo tratar_erro

If Lista.ListItems.Count > 0 And txtCodigo.Text = 2 Then
cmdF8_Click
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Lista_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

ProcPuxaDados

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcPuxaDados()
On Error GoTo tratar_erro

If Lista.ListItems.Count > 0 Then
IDProducao = Lista.SelectedItem
    Evento = Lista.SelectedItem.ListSubItems.Item(1)
    DescEvento = Lista.SelectedItem.ListSubItems(2)
    
    If Codigo_Barras = False Then
        txtCodigo = Evento
        cmbdescricao = DescEvento
    Else
        If Len(Evento) = 1 Then
        txtCodigoBarras = "00000" & Evento
        End If
        
        If Len(Evento) = 2 Then
        txtCodigoBarras = "0000" & Evento
        End If

       ' txtCodigoBarras = Evento
        txtCodigoDesc = DescEvento
    End If
    Set TBProducao = CreateObject("adodb.recordset")
    TBProducao.Open "Select Turno from " & NomeTabelaAp & " where idproducao= " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
    If TBProducao.EOF = False Then
        txtturno = TBProducao!Turno
    End If
    TBProducao.Close
    ExcluiSel = True
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Sub ProcAtualizaProducao(CarregaLista As Boolean, CalculaValores As Boolean)
On Error GoTo tratar_erro
Dim TempoTotalDias As Date 'OK
Dim TOTALOK As Long
Dim TOTALNC As Long
Dim Tempoprocesso As Date
Dim totalrecord As Long
Dim Aplicacao As String

'Verifica o último evento apontado
If Varias_OS = True Then
    TextoFiltro = "Select " & NomeTabelaAp & ".* from (" & NomeTabelaAp & " INNER JOIN OrdemServico ON " & NomeTabelaAp & ".IDfase = OrdemServico.IDproducao) INNER JOIN ProducaoFases_OS ON OrdemServico.ID_apontamento = ProducaoFases_OS.ID where ProducaoFases_OS.ID = " & Txt_ID_apontamento & " and " & NomeTabelaAp & ".Maquina = '" & txtMaquina & "' order by " & NomeTabelaAp & ".Data, " & NomeTabelaAp & ".Tempoinicio"
Else
    TextoFiltro = "Select * from " & NomeTabelaAp & " where idfase = " & ListaOS.SelectedItem.ListSubItems(1) & " and Maquina = '" & txtMaquina & "' order by Data, Tempoinicio"
End If
Set TBProducao = CreateObject("adodb.recordset")
TBProducao.Open TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
If TBProducao.EOF = False Then
    TBProducao.MoveLast
    Evento = TBProducao!codigoDesc
    UltimoMaq = TBProducao!Maquina
    Ultimo = TBProducao!codigoDesc
    UltimoDesc = TBProducao!Descricao
    TempoUltimo = TBProducao!TempoInicio
    ULTICOD = TBProducao!codigoDesc
    ULTIDESC = TBProducao!Descricao
    ULTIOPERADOR = TBProducao!Usuario
    
    TBProducao.MovePrevious
    If TBProducao.BOF = False Then
        'Grava evento na variavel penultimo
        PenultimoMaq = TBProducao!Maquina
        Penultimo = TBProducao!codigoDesc
        PenultimoDesc = TBProducao!Descricao
        'Informa turno
        If IsNull(TBProducao!Turno) = False Then Turno = TBProducao!Turno
        txtturno = Turno
        'Informa data do penultimo evento
        Dataini = TBProducao!Data
    End If
Else
    UltimoMaq = ""
    Ultimo = 0
    UltimoDesc = ""
    TempoUltimo = 0
    PenultimoMaq = ""
    Penultimo = 0
    PenultimoDesc = ""
    TOTALOK = 0
    TOTALNC = 0
    Lista.ListItems.Clear
    Evento = 0
    ULTICOD = 0
    ULTIDESC = ""
    ULTIOPERADOR = ""
End If
TBProducao.Close

'Carrega dados na lista OS, máquina
If CarregaLista = True Then
    txtstatus.Text = ""
    TempoUtilizadoDescricao = 0
    
    Lista.ListItems.Clear
    Set TBProducao = CreateObject("adodb.recordset")
    StrSQL = "Select * from " & NomeTabelaAp & " where idfase = " & ListaOS.SelectedItem.ListSubItems(1) & " and maquina = '" & txtMaquina.Text & "' order by Data, Tempoinicio"
    Debug.Print StrSQL
    
    TBProducao.Open "Select * from " & NomeTabelaAp & " where idfase = " & ListaOS.SelectedItem.ListSubItems(1) & " and maquina = '" & txtMaquina.Text & "' order by Data, Tempoinicio", Conexao, adOpenKeyset, adLockOptimistic
    If TBProducao.EOF = False Then
        TBProducao.MoveLast
'        PBLista.Min = 0
'        PBLista.Max = TBProducao.RecordCount
'        PBLista.Value = 1
        Contador = 0
        TBProducao.MoveFirst
        Do While TBProducao.EOF = False
            With Lista.ListItems
                .Add , , TBProducao!IDProducao
                .Item(.Count).SubItems(1) = IIf(IsNull(TBProducao!codigoDesc), 0, TBProducao!codigoDesc)
                .Item(.Count).SubItems(2) = TBProducao!Descricao
                txtstatus.Text = TBProducao!Descricao
                .Item(.Count).SubItems(3) = TBProducao!TempoInicio
                .Item(.Count).SubItems(4) = TBProducao!TempoFinal
                If TBProducao!TempoInicio <> "00:00:00" Then
                    TempoFinal = TBProducao!TempoInicio
                    TempoUtilizadoDescricao = TBProducao!TempoInicio
                End If
                If TBProducao!Dias <> 0 Then
                    TempoTotalDias = IIf(IsNull(TBProducao!TempoTotal), 0, TBProducao!TempoTotal) + TBProducao!Dias
                    FunElapsedTime (TempoTotalDias)
                    .Item(.Count).SubItems(5) = Horatotal
                Else
                    .Item(.Count).SubItems(5) = IIf(IsNull(TBProducao!TempoTotal), "", TBProducao!TempoTotal)
                End If
                .Item(.Count).SubItems(6) = TBProducao!Usuario
                .Item(.Count).SubItems(7) = IIf(IsNull(TBProducao!Quantidade), 0, TBProducao!Quantidade)
                .Item(.Count).SubItems(8) = IIf(IsNull(TBProducao!Reprovada), 0, TBProducao!Reprovada)
                .Item(.Count).Selected = True
            End With
            If TBProducao!codigoDesc = 3 Then Tempoprocesso = TBProducao!Data
            Contador = Contador + 1
            TBProducao.MoveNext
        Loop
        
        TBProducao.MoveFirst
        If TBProducao.RecordCount > 12 Then
            totalrecord = TBProducao.RecordCount
            TBProducao.MoveLast
            Contador = 12
            Do While Contador > 1
                TBProducao.MovePrevious
                Contador = Contador - 1
            Loop
        End If
    End If
    TBProducao.Close
End If

If CalculaValores = True Then
    ProcAtualizaPrepExecUtil
    ProcAtualizaPrepExecUtilTotalizacao
    ProcAtualizaPrepExecUtilTotalizacaoMaq
    ProcAtualizaOF
    ProcFecharQuantidades
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcFecharQuantidades()
On Error GoTo tratar_erro

Set TBOrdem = CreateObject("adodb.recordset")
TBOrdem.Open "Select * from Producao where Ordem = " & OF, Conexao, adOpenKeyset, adLockOptimistic
If TBOrdem.EOF = False Then
    TBOrdem!QuantProd = TOK + TNC + TCD
    TBOrdem!QuantNC = TNC
    TBOrdem!QTCD = TCD
    TBOrdem.Update
End If
        
Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Sub ProcAcertaCadMaquina()
On Error GoTo tratar_erro
Dim Liberada As String

Conexao.Execute "DELETE from CadMaquinas_Monitor where Maquina = '" & txtMaquina & "'"
Liberada = "Liberada = 'True'"
Set TBLista = CreateObject("adodb.recordset")
TBLista.Open "Select * from ProducaoFases where maquina = '" & txtMaquina & "' order by data desc, Tempoinicio desc", Conexao, adOpenKeyset, adLockOptimistic
If TBLista.EOF = False Then
    Set TBOS = CreateObject("adodb.recordset")
    TBOS.Open "Select * from Ordemservico where Idproducao = " & TBLista!OS, Conexao, adOpenKeyset, adLockOptimistic
    If TBOS.EOF = False Then
        Set TBProducao = CreateObject("adodb.recordset")
        TBProducao.Open "Select * from CadMaquinas_Monitor where Maquina = '" & txtMaquina & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBProducao.EOF = True Then TBProducao.AddNew
        TBProducao!Maquina = txtMaquina
        TBProducao!Data = TBLista!Data
        TBProducao!TempoInicio = TBLista!TempoInicio
        TBProducao!Ordem = TBLista!Ordem
        TBProducao!OS = TBLista!OS
        TBProducao!Operador = TBLista!Usuario
        TBProducao!Evento = TBLista!codigoDesc
        TBProducao!DescEvento = TBLista!Descricao
        TBProducao!CP = TBOS!CPPECA
        TBProducao!CR = TBOS!CRPECA
        TBProducao!TP = TBOS!TempoExecucao
        TBProducao!TR = TBOS!TEUTIL
        TBProducao!Eficiencia = TBOS!Eficiencia
        TBProducao.Update
        TBProducao.Close
        
        If TBOS!Custos = True Then
            Set TBCodigoDesc = CreateObject("adodb.recordset")
            TBCodigoDesc.Open "Select Liberar_Posto from CodigoDesc where codigo = " & TBLista!codigoDesc, Conexao, adOpenKeyset, adLockOptimistic
            If TBCodigoDesc.EOF = False Then
                If TBCodigoDesc!liberar_posto = True Then Liberada = "Liberada = 'True'" Else Liberada = "Liberada = 'False'"
            End If
            TBCodigoDesc.Close
        End If
    End If
End If
Conexao.Execute "Update CadMaquinas Set " & Liberada & " where Maquina = '" & txtMaquina & "'"

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Sub ProcAbrir()
On Error GoTo tratar_erro

With IIf(Ap_plano = True, frmabrir_plano_prod, frmabrir_OS)
    Ultimo = ""
    Txt_OS = .txtOS
    
    Set TBOrdem = CreateObject("adodb.recordset")
    TBOrdem.Open "Select P.Ap_backup, P.IDProcesso from Producao P INNER JOIN Ordemservico OS ON OS.Ordem = P.Ordem where OS.IDproducao = " & Txt_OS, Conexao, adOpenKeyset, adLockOptimistic
    If TBOrdem.EOF = False Then
        If TBOrdem!Ap_backup = True Then
            NomeTabelaAp = "ProducaoFases_Backup"
            NomeTabelaApTotalizacao = "ProducaoFases_Totalizacao_Backup"
        Else
            NomeTabelaAp = "ProducaoFases"
            NomeTabelaApTotalizacao = "ProducaoFases_Totalizacao"
        End If
        
        IDProcesso = TBOrdem!IDProcesso
        
        'Localiza ordem de servico em relacao a maquina, fase, ordem
        Set TBProcessos = CreateObject("adodb.recordset")
        TBProcessos.Open "Select * from ordemservico where IDPRODUCAO = " & .txtOS.Text, Conexao, adOpenKeyset, adLockOptimistic
        If TBProcessos.EOF = True Then
            TBProcessos.AddNew
            TBProcessos!Fase = .txtfase.Text
            TBProcessos!Maquina = .cmbmaquina.Text
            TBProcessos!Ordem = .TXTOF.Text
            TBProcessos!Pronto = "NÃO"
            TBProcessos!Preparacao = "00:00:00"
            TBProcessos!Execucao = "00:00:00"
            TBProcessos!Pcshora = 0
            TBProcessos!pecahora = False
            TBProcessos!TempoPreparacao = "00:00:00"
            TBProcessos!TempoExecucao = "00:00:00"
            TBProcessos!TESegundos = 0
            TBProcessos!pc_te = 0
            TBProcessos!Prazofinal = Format(.txtprazo, "dd/mm/yy")
            TBProcessos!Quantidade = .txtquant
            TBProcessos!Descricao = .txtdescricao.Text
            TBProcessos!Desenho = .txtdesenho.Text
            
            If MsgBox("O tempo de preparação desta(s) OS('s) foi reaproveitado de outra(s) OS('s)?", vbYesNo + vbDefaultButton2) = vbYes Then
                ProcVerifTempoPrepReaproveitado .cmbmaquina
            Else
                TBProcessos!Tempo_prep_reaproveitado = False
                TBProcessos!OS_reaproveitada = Null
            End If
            If TBProcessos!Tempo_prep_reaproveitado = True Then TempoPreparacaoReaprov = True Else TempoPreparacaoReaprov = False
            
            TBProcessos.Update
            
            Lista.ListItems.Clear
            frmProducao.Show
            txtMaquina = .cmbmaquina
            Txt_descricao_posto = .txtdescmaq
            
            Txt_ID_apontamento = .Txt_ID_apontamento
            ProcCarregaListaOS
        Else
            'Verifica se é o primeiro apontamento e pergunta se o tempo de preparação foi reaproveitado
            Set TBProducao = CreateObject("adodb.recordset")
            TBProducao.Open "Select IDproducao from " & NomeTabelaAp & " where idfase = " & .txtOS & " and maquina = '" & .cmbmaquina & "'", Conexao, adOpenKeyset, adLockOptimistic
            If TBProducao.EOF = True Then
                If MsgBox("O tempo de preparação desta(s) OS('s) foi reaproveitado de outra(s) OS('s)?", vbYesNo + vbDefaultButton2) = vbYes Then
                    ProcVerifTempoPrepReaproveitado .cmbmaquina
                Else
                    TBProcessos!Tempo_prep_reaproveitado = False
                    TBProcessos!OS_reaproveitada = Null
                End If
                TBProcessos.Update
            End If
            TBProducao.Close
            If TBProcessos!Tempo_prep_reaproveitado = True Then TempoPreparacaoReaprov = True Else TempoPreparacaoReaprov = False
            
            frmProducao.Show
            txtMaquina = .cmbmaquina
            Txt_descricao_posto = .txtdescmaq
            
            Txt_ID_apontamento = .Txt_ID_apontamento
            ProcCarregaListaOS
            ProcAtualizaProducao True, False
            ProcAtualizaCodigoDesc
            If Codigo_Barras = False Then
                If Ap_codigo = True Then txtCodigo.SetFocus Else cmbdescricao.SetFocus
            Else
                txtCodigoBarras.SetFocus
            End If
        End If
    Else
        OrdemExiste = False
        MsgBox ("Ordem de fabricação não cadastrada."), vbExclamation
        Exit Sub
    End If
    OrdemExiste = True
End With

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcVerifTempoPrepReaproveitado(Posto As String)
On Error GoTo tratar_erro

'Verifica a última OS com preparação apontada neste posto
OS_texto = ""
Set TBOS = CreateObject("adodb.recordset")
TBOS.Open "Select OS from Ordemservico_maq_utilizadas where maquina = '" & Posto & "' and TPUSEG <> 0 order by ID", Conexao, adOpenKeyset, adLockOptimistic
If TBOS.EOF = False Then
    TBOS.MoveLast
    OS_texto = TBOS!OS
End If
TBOS.Close

Mensagem1:
    OS_texto = InputBox("Favor informar o número da OS de qual foi reaproveitado o tempo de preparação.", , OS_texto)
    If OS_texto = "" Then
        If MsgBox("Deseja cancelar esta operação?", vbYesNo) = vbYes Then Exit Sub Else GoTo Mensagem1
    End If
    If IsNumeric(OS_texto) = False Then
        MsgBox ("Só é permitido número neste campo."), vbExclamation
        GoTo Mensagem1
    End If
    OS = OS_texto
    Set TBFases = CreateObject("adodb.recordset")
    TBFases.Open "Select * from OrdemServico where IDProducao = " & OS, Conexao, adOpenKeyset, adLockOptimistic
    If TBFases.EOF = True Then
        MsgBox ("Não foi encontrado nenhuma OS com este número."), vbExclamation
        GoTo Mensagem1
    Else
        'Verifica se a OS foi apontada
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from Ordemservico_maq_utilizadas where OS = " & OS, Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = True Then
            MsgBox ("Esta OS ainda não foi apontada."), vbExclamation
            GoTo Mensagem1
        End If
        TBAbrir.Close
        
        TBProcessos!Tempo_prep_reaproveitado = True
        TBProcessos!OS_reaproveitada = OS
    End If
    TBFases.Close

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcCarregaListaOS()
On Error GoTo tratar_erro

If Varias_OS = True Then
    TextoFiltro = "Select OS.*, P.Ordem from (OrdemServico OS INNER JOIN Producao P ON OS.Ordem = P.Ordem) INNER JOIN ProducaoFases_OS PFOS ON OS.ID_apontamento = PFOS.ID where PFOS.ID = " & Txt_ID_apontamento & " order by OS.IDproducao"
Else
    TextoFiltro = "Select OS.*, P.Ordem from OrdemServico OS INNER JOIN Producao P ON OS.Ordem = P.Ordem where OS.IDproducao = " & Txt_OS & " order by OS.IDproducao"
End If
ListaOS.ListItems.Clear
Set TBOS = CreateObject("adodb.recordset")
TBOS.Open TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
If TBOS.EOF = False Then

If TBOS.RecordCount = 1 Then
        Txt_instrucao.TextRTF = IIf(IsNull(TBOS!descFase), "", TBOS!descFase)
End If

    
    Txt_eficiencia_prep = IIf(IsNull(TBOS!Eficiencia_prep), 0, Format(TBOS!Eficiencia_prep, "###,##0.00")) & "%"
    Txt_eficiencia_exec = IIf(IsNull(TBOS!Eficiencia_exec), 0, Format(TBOS!Eficiencia_exec, "###,##0.00")) & "%"
    Txt_eficiencia_media = IIf(IsNull(TBOS!Eficiencia), 0, Format(TBOS!Eficiencia, "###,##0.00")) & "%"
    
    Do While TBOS.EOF = False
        With ListaOS.ListItems
            .Add , , TBOS!Ordem
            .Item(.Count).SubItems(1) = TBOS!IDProducao
            .Item(.Count).SubItems(2) = TBOS!Fase
            .Item(.Count).SubItems(3) = Format(TBOS!Prazofinal, "dd/mm/yy")
            
            'Verifica qtde. de pçs apontadas na OS anterior
            If TBOS!Processo_controlado = True Then
                Set TBAbrir = CreateObject("adodb.recordset")
                TBAbrir.Open "Select * from ordemservico where Ordem = " & TBOS!Ordem & " and Retrabalho = 'False' and Fase < '" & TBOS!Fase & "' order by Fase desc", Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then
                    .Item(.Count).SubItems(4) = IIf(IsNull(TBAbrir!QTOK), 0, TBAbrir!QTOK) + FunVerifQtdeRetrabalhoFase(TBOS!Ordem, TBAbrir!Fase)
                Else
                    .Item(.Count).SubItems(4) = TBOS!Quantidade
                End If
                TBAbrir.Close
            Else
                Set TBAbrir = CreateObject("adodb.recordset")
                TBAbrir.Open "Select Sum(QTNC) as TTNC from ordemservico where Ordem = " & TBOS!Ordem & " and Idproducao <> " & TBOS!IDProducao, Conexao, adOpenKeyset, adLockOptimistic
                If TBAbrir.EOF = False Then
                    QtdeRefugo = IIf(IsNull(TBAbrir!TTNC), 0, TBAbrir!TTNC)
                End If
                TBAbrir.Close
                .Item(.Count).SubItems(4) = TBOS!Quantidade - IIf(TBOS!Retrabalho = False, QtdeRefugo, 0)
            End If
            
            IDProcesso = IIf(IsNull(TBOS!IDProcesso), 0, TBOS!IDProcesso)
            'Tempo de preparação previsto
            ProcFormataHora (IIf(IsNull(TBOS!Preparacao), 0, TBOS!Preparacao))
            .Item(.Count).SubItems(5) = Horatotal
            'Tempo de preparação real
            ProcFormataHora (IIf(IsNull(TBOS!TPUTIL), 0, TBOS!TPUTIL))
            .Item(.Count).SubItems(6) = Horatotal
            'Tempo de execução previsto
            ProcFormataHora (IIf(IsNull(TBOS!Execucao), 0, TBOS!Execucao))
            .Item(.Count).SubItems(7) = Horatotal
            'Tempo de execução real
            ProcFormataHora (IIf(IsNull(TBOS!TEUTIL), 0, TBOS!TEUTIL))
            .Item(.Count).SubItems(8) = Horatotal
            'Tempo total previsto
            .Item(.Count).SubItems(9) = IIf(IsNull(TBOS!Tempototallote), "00:00:00", TBOS!Tempototallote)
            'Tempo total real
            .Item(.Count).SubItems(10) = IIf(IsNull(TBOS!TETTUTIL), "00:00:00", TBOS!TETTUTIL)
            'Qtde. total aprov.
            .Item(.Count).SubItems(11) = IIf(IsNull(TBOS!QTOK), 0, TBOS!QTOK)
            'Qtde. total NC
            .Item(.Count).SubItems(12) = IIf(IsNull(TBOS!QTCD), 0, TBOS!QTCD)
            .Item(.Count).SubItems(13) = IIf(IsNull(TBOS!QTNC), 0, TBOS!QTNC)
            'Qtde. total prod.
            .Item(.Count).SubItems(14) = IIf(IsNull(TBOS!Totalprod), 0, TBOS!Totalprod)
        End With
        TBOS.MoveNext
    Loop
End If
TBOS.Close

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Sub ProcAbrirNovo()
On Error GoTo tratar_erro

With frmabrir_Ordem
    Ultimo = ""
    Txt_OS = .txtOS
    
    Set TBOrdem = CreateObject("adodb.recordset")
    TBOrdem.Open "Select * FROM Producao where Ordem = " & .txtordem.Text, Conexao, adOpenKeyset, adLockOptimistic
    If TBOrdem.BOF = False Then
        OF = TBOrdem!Ordem
        IDProcesso = TBOrdem!IDProcesso
        If TBOrdem!Ap_backup = True Then
            NomeTabelaAp = "ProducaoFases_Backup"
            NomeTabelaApTotalizacao = "ProducaoFases_Totalizacao_Backup"
        Else
            NomeTabelaAp = "ProducaoFases"
            NomeTabelaApTotalizacao = "ProducaoFases_Totalizacao"
        End If
    End If
    
    'localiza ordem de servico em relacao a maquina, OS
    Set TBProcessos = CreateObject("adodb.recordset")
    If .txtOS <> "" Then
        TBProcessos.Open "Select * from ordemservico where idproducao = " & .txtOS, Conexao, adOpenKeyset, adLockOptimistic
    Else
        TBProcessos.Open "Select * from ordemservico where Ordem = " & OF & " AND maquina = '" & .cmbPT.Text & "' and fase = " & .txtfase, Conexao, adOpenKeyset, adLockOptimistic
    End If
    If TBProcessos.EOF = True Then
        TBProcessos.AddNew
        TBProcessos!Fase = .txtfase.Text
        TBProcessos!Maquina = .cmbPT.Text
        TBProcessos!Ordem = OF
        TBProcessos!Quantidade = .txtquant.Text
        TBProcessos!Pronto = "NÃO"
        TBProcessos!Prazofinal = .mskprazofinal.Text
        TBProcessos!Status = "Aguardando"
        TBProcessos!Preparacao = "00:00:00"
        TBProcessos!Execucao = "00:00:00"
        TBProcessos!TempoPreparacao = "00:00:00"
        TBProcessos!TempoExecucao = "00:00:00"
        TBProcessos!TESegundos = 0
        TBProcessos!Tempototallote = "00:00:00"
        TBProcessos!TTLPREVS = 0
        TBProcessos!pecahora = False
        TBProcessos!Pcshora = 0
        TBProcessos!pc_te = 1
        TBProcessos!CPPECA = 0
        TBProcessos!CPLOTE = 0
        TBProcessos!OSControlada = False
        TBProcessos!Processo_controlado = False
        TBProcessos!Retrabalho = False
        
        'Verif. se tem plano de inspeção
        TBProcessos!IDPlano = 0
        Set TBPlano = CreateObject("adodb.recordset")
        TBPlano.Open "Select IDPlano from Plano where Desenho = '" & .txtitem & "' and Fase = " & .txtfase.Text, Conexao, adOpenKeyset, adLockOptimistic
        If TBPlano.EOF = False Then
            TBProcessos!IDPlano = TBPlano!IDPlano
        End If
        TBPlano.Close
                       
        'Verifica se a maquina agrega custos/eficiencia na ordem
        Set TBMaquinas = CreateObject("adodb.recordset")
        TBMaquinas.Open "Select custos from cadmaquinas where maquina = '" & .cmbPT.Text & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBMaquinas.EOF = False Then
            If TBMaquinas!Custos = True Then TBProcessos!Custos = True Else TBProcessos!Custos = False
        End If
        TBMaquinas.Close
        
        If MsgBox("O tempo de preparação desta(s) OS('s) foi reaproveitado de outra(s) OS('s)?", vbYesNo) = vbYes Then
            ProcVerifTempoPrepReaproveitado .cmbPT.Text
        Else
            TBProcessos!Tempo_prep_reaproveitado = False
            TBProcessos!OS_reaproveitada = Null
        End If
        If TBProcessos!Tempo_prep_reaproveitado = True Then TempoPreparacaoReaprov = True Else TempoPreparacaoReaprov = False
        
        TBProcessos.Update
        Txt_OS = TBProcessos!IDProducao
        
        frmProducao.Show
        Lista.ListItems.Clear
        txtMaquina = .cmbPT
        Txt_descricao_posto = .txtDescricaoPT
        ProcCarregaListaOS
    Else
        'Verifica se é o primeiro apontamento e pergunta se o tempo de preparação foi reaproveitado
        Set TBProducao = CreateObject("adodb.recordset")
        TBProducao.Open "Select * from " & NomeTabelaAp & " where idfase = " & .txtOS & " and maquina = '" & .cmbPT & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBProducao.EOF = True Then
            If MsgBox("O tempo de preparação desta(s) OS('s) foi reaproveitado de outra(s) OS('s)?", vbYesNo) = vbYes Then
                ProcVerifTempoPrepReaproveitado .cmbPT.Text
            Else
                TBProcessos!Tempo_prep_reaproveitado = False
                TBProcessos!OS_reaproveitada = Null
            End If
            TBProcessos.Update
        End If
        TBProducao.Close
        If TBProcessos!Tempo_prep_reaproveitado = True Then TempoPreparacaoReaprov = True Else TempoPreparacaoReaprov = False
        
        frmProducao.Show
        txtMaquina = .cmbPT
        Txt_descricao_posto = .txtDescricaoPT
        ProcCarregaListaOS
        ProcAtualizaProducao True, False
        ProcAtualizaCodigoDesc
        If Codigo_Barras = False Then
            If Ap_codigo = True Then txtCodigo.SetFocus Else cmbdescricao.SetFocus
        Else
            txtCodigoBarras.SetFocus
        End If
    End If
    OrdemExiste = True
End With

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub


Private Sub ListaOS_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If ListaOS.ListItems.Count > 0 Then
    Set TBLista = CreateObject("adodb.recordset")
    TBLista.Open "Select DescFase from ordemServico where IDProducao = '" & ListaOS.SelectedItem.ListSubItems.Item(1).Text & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBLista.EOF = False Then
        Txt_instrucao.TextRTF = IIf(IsNull(TBLista!descFase), "", TBLista!descFase)
    End If
    TBLista.Close
    
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ListaRequisicao_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo tratar_erro

If ListaRequisicao.ListItems.Count > 0 Then
txtobservacoes.Text = ListaRequisicao.SelectedItem.ListSubItems.Item(5).Text
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Timer_logon_Timer()
On Error GoTo tratar_erro

1:
    Set TBLogon = CreateObject("adodb.recordset")
    TBLogon.Open "Select * from Logon where Usuario = '" & Operador & "' and Data = '" & Date & "' and Tipo = 'G'", Conexao, adOpenKeyset, adLockOptimistic
    If TBLogon.EOF = True Then
        If PubUsuario <> "PROCAM" Then
            MsgBox ("O usuário " & Operador & " foi desconectado do sistema, o módulo para apontamento será fechado."), vbCritical
            Call Form_KeyDown(117, 0)
        End If
    End If
    TBLogon.Close
    
Exit Sub
tratar_erro:
    If Err.Number = -2147467259 Or Err.Number = 3709 Then
        FunAbreBD
        GoTo 1
    End If
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub TimerRelogio_Timer()
On Error GoTo tratar_erro
Dim tempoutilizado As Date

txtData = Format(Date, "dd/mm/yy") 'Format(FunHoraServidor("\\" & FunVerifNomeServidor), "dd/mm/yy")
txtHora = Format(Now, "hh:mm:ss") 'Format(FunHoraServidor("\\" & FunVerifNomeServidor), "hh:mm:ss")

If TempoUtilizadoDescricao > 0 Then
    TempoInicio = Now 'FunHoraServidor("\\" & Familiatext)
    tempoutilizado = TempoInicio - TempoUtilizadoDescricao
    FunElapsedTime (tempoutilizado)
    TxtTempoUtilizado = Horatotal
Else
    TxtTempoUtilizado = "00:00:00"
End If
'frmProducao.Refresh
Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub cmbdescricao_Click()
On Error GoTo tratar_erro

If Ap_codigo = False Then
    Set TBCodigoDesc = CreateObject("adodb.recordset")
    TBCodigoDesc.Open "Select Codigo from CodigoDesc where Descricao = '" & cmbdescricao.Text & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBCodigoDesc.EOF = False Then
        txtCodigo.Text = TBCodigoDesc!Codigo
    End If
    TBCodigoDesc.Close
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Sub ProcAtualizaCodigoDesc()
On Error GoTo tratar_erro
Dim OSTexto As String

If Codigo_Barras = False Then
    If txtMaquina = "" Then GoTo 1
    Set TBProducao = CreateObject("adodb.recordset")
    TBProducao.Open "select * from manutencao where IDmaquina = '" & txtMaquina & "' and Controlada = 'true'", Conexao, adOpenKeyset, adLockOptimistic
    If TBProducao.EOF = False Then
        Set TBProcessosDet = CreateObject("adodb.recordset")
        TBProcessosDet.Open "select * from manutencao_data where idManutencao = " & TBProducao!Codigo & " and status = 'Aberta' and IDProducao = 0 and data <= '" & Format(Date, "Short Date") & "'", Conexao, adOpenKeyset, adLockOptimistic
        If TBProcessosDet.EOF = False Then
            txtCodigo = ""
            With cmbdescricao
                .Clear
                Set TBCodigoDesc = CreateObject("adodb.recordset")
                TBCodigoDesc.Open "Select * from CodigoDesc where descricao = 'MÁQUINA EM MANUTENÇÃO' or descricao = 'MANUTENÇÃO PREVENTIVA' or descricao = 'MANUTENÇÃO CORRETIVA'", Conexao, adOpenKeyset, adLockOptimistic
                If TBCodigoDesc.EOF = False Then
                    .AddItem TBCodigoDesc!Descricao
                    .ItemData(.ListCount - 1) = TBCodigoDesc!Codigo
                End If
                TBCodigoDesc.Close
            End With
        Else
            GoTo 1
        End If
    Else
1:
        txtCodigo = ""
        If TempoPreparacaoReaprov = True Then
            With frmabrir_OS
                If ListaOS.ListItems.Count = 0 Then OSTexto = .txtOS Else OSTexto = ListaOS.SelectedItem.ListSubItems(1)
                Set TBProducao = CreateObject("adodb.recordset")
                TBProducao.Open "Select * from ProducaoFases where OS = " & OSTexto, Conexao, adOpenKeyset, adLockOptimistic
                If TBProducao.EOF = False Then
                    If TBProducao.RecordCount >= 2 Then TextoFiltro = "Codigo is not null" Else TextoFiltro = "Codigo <> 1"
                Else
                    TextoFiltro = "Codigo <> 1"
                End If
            End With
        Else
            TextoFiltro = "Codigo is not null"
        End If
        With cmbdescricao
            .Clear
            Set TBCodigoDesc = CreateObject("adodb.recordset")
            TBCodigoDesc.Open "Select * from CodigoDesc where " & TextoFiltro & " and Bloqueado = 'False' order by Codigo", Conexao, adOpenKeyset, adLockOptimistic
            If TBCodigoDesc.EOF = False Then
                Do While TBCodigoDesc.EOF = False
                    .AddItem TBCodigoDesc!Descricao
                    .ItemData(.ListCount - 1) = TBCodigoDesc!Codigo
                    TBCodigoDesc.MoveNext
                Loop
            End If
            TBCodigoDesc.Close
        End With
    End If
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Sub ProcExcluir()
On Error GoTo tratar_erro
Dim ParecerCQ As String
Dim Descricao As String

If Lista.ListItems.Count <> 0 And ExcluiSel = True Then
    If Evento = "" Then
        MsgBox ("Informe o evento na lista antes de excluir."), vbExclamation
        Exit Sub
    End If
    
    'Se o evento não for o ultimo da lista
    If Lista.SelectedItem.Index <> Lista.ListItems.Count Then
        MsgBox "Só é possível excluir apenas o último apontamento da lista.", vbExclamation
        Exit Sub
    Else
        If USMsgBox("Deseja realmente excluir o evento selecionado?", vbYesNo, "CAPRIND v5.0") = vbYes Then
            Set TBProducao = CreateObject("adodb.recordset")
            TBProducao.Open "Select * from " & NomeTabelaAp & " where IDProducao = " & Lista.SelectedItem, Conexao, adOpenKeyset, adLockOptimistic
            If TBProducao.EOF = False Then
                If TBProducao!Usuario <> Operador Then
                    MsgBox ("Somente o(a) operador(a) " & TBProducao!Usuario & " tem permissão de excluir este evento."), vbExclamation
                    Exit Sub
                End If
            End If
            TBProducao.Close
            
            If Format(Lista.SelectedItem.ListSubItems(3), "dd/mm/yy") <> Format(Date, "dd/mm/yy") Then
                MsgBox ("Só é permitido excluir o evento no dia " & Format(Lista.SelectedItem.ListSubItems(3), "dd/mm/yy") & ", dia do apontamento."), vbExclamation
                Exit Sub
            End If
            
            'Verifica se existe movimentação de saida no estoque quando a ordem tiver movimentação de entrada automática
            Set TBProducao = CreateObject("adodb.recordset")
            TBProducao.Open "Select EC.IDestoque from (Producao P INNER JOIN Estoque_controle EC ON EC.Lote = P.Ordem) INNER JOIN Estoque_movimentacao EM ON EM.IDestoque = EC.IDestoque where P.Entrar_estoque = 'True' and EC.Lote = '" & ListaOS.SelectedItem & "' and EM.Saida <> 0", Conexao, adOpenKeyset, adLockOptimistic
            If TBProducao.EOF = False Then
                MsgBox ("Não é permitido excluir apontamento, pois existe movimentação de saída no estoque no RE dessa ordem."), vbExclamation
                TBProducao.Close
                Exit Sub
            End If
            TBProducao.Close
            
            If Varias_OS = True Then
                TextoFiltro = "Select OrdemServico.* from OrdemServico INNER JOIN ProducaoFases_OS ON OrdemServico.ID_apontamento = ProducaoFases_OS.ID where ProducaoFases_OS.ID = " & Txt_ID_apontamento & " order by OrdemServico.IDproducao"
            Else
                TextoFiltro = "Select * from OrdemServico where IDProducao = " & ListaOS.SelectedItem.ListSubItems(1)
            End If
            Set TBOS = CreateObject("adodb.recordset")
            TBOS.Open TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
            If TBOS.EOF = False Then
                Do While TBOS.EOF = False
                    Set TBCiclo = CreateObject("adodb.recordset")
                    TBCiclo.Open "Select * from " & NomeTabelaAp & " where OS = " & TBOS!IDProducao & " and Maquina = '" & txtMaquina & "' and usuario = '" & Operador & "' order by Data, Tempoinicio", Conexao, adOpenKeyset, adLockOptimistic
                    If TBCiclo.EOF = False Then
                        TBCiclo.MoveLast
                        
                        If FunVerifNCRejeitado(TBOS!IDProducao) = True Then ParecerCQ = "Rejeitar" Else ParecerCQ = "Nada consta"
                        
                        'Verificar se tem parecer do controle da qualidade
                        If IsNull(TBCiclo!Reprovada) = False And TBCiclo!Reprovada <> 0 Then
                            Set TBProducao = CreateObject("adodb.recordset")
                            TBProducao.Open "Select * from CQ_NC_FABRICA where IDProducao = " & TBCiclo!IDProducao & " and PARECERCQ IS NOT NULL and PARECERCQ <> '" & ParecerCQ & "' and Analizada = 1", Conexao, adOpenKeyset, adLockOptimistic
                            If TBProducao.EOF = False Then
                                USMsgBox ("Não é possível excluir este apontamento, pois o mesmo já possui a disposição do controle de qualidade na Ordem: " & TBOS!Ordem & " - OS: " & TBOS!IDProducao & "."), vbExclamation, "GERPROD"
                                TBProducao.Close
                                TBCiclo.Close
                                Exit Sub
                            End If
                            TBProducao.Close
                        End If
            
                        'Verificar se tem paracer do controle da qualidade no evento anterior
                        TBCiclo.MovePrevious
                        If TBCiclo.BOF = False Then
                            If IsNull(TBCiclo!Reprovada) = False And TBCiclo!Reprovada <> 0 Then
                                Set TBProducao = CreateObject("adodb.recordset")
                                TBProducao.Open "Select * from CQ_NC_FABRICA where IDProducao = " & TBCiclo!IDProducao & " and PARECERCQ IS NOT NULL and PARECERCQ <> '" & ParecerCQ & "' and Analizada = 1", Conexao, adOpenKeyset, adLockOptimistic
                                If TBProducao.EOF = False Then
                                    MsgBox ("Não é possível excluir este apontamento, pois o evento anterior possui a disposição do controle de qualidade na Ordem: " & TBOS!Ordem & " - OS: " & TBOS!IDProducao & "."), vbExclamation
                                    TBProducao.Close
                                    TBCiclo.Close
                                    Exit Sub
                                End If
                                TBProducao.Close
                            End If
                        End If
                                    
                        If Evento = 3 Then Conexao.Execute "Update Producao Set pronta = 'NÃO', dataentrega = Null, concluida = 'False' where Ordem = " & TBOS!Ordem
                        If Evento <> 0 Then
                            'Exclui não conformidade da tabela CQ_NC_FABRICA
                            If TBCiclo.BOF = False Then
                               Conexao.Execute "DELETE from ProducaoFases_Codigos where IDProducao = " & Lista.SelectedItem & " or IDProducao = " & TBCiclo!IDProducao
                               Conexao.Execute "DELETE from CQ_NC_FABRICA_Serie where IDProducao = " & Lista.SelectedItem & " or IDProducao = " & TBCiclo!IDProducao
                               Conexao.Execute "DELETE from CQ_NC_FABRICA where IDProducao = " & Lista.SelectedItem & " or IDProducao = " & TBCiclo!IDProducao
                            End If
                            
                            TBCiclo.MoveLast
                
                            'Informa turno
                            Turno = TBCiclo!Turno
                
                            'Informa dados do evento anterior
                            TBCiclo.MovePrevious
                            If TBCiclo.BOF = True Then TBCiclo.MoveNext
                            PenultimoMaq = TBCiclo!Maquina
                            Penultimo = TBCiclo!codigoDesc
                            PenultimoDesc = TBCiclo!Descricao
                            Dataini = TBCiclo!Data
                            Dias = IIf(IsNull(TBCiclo!Dias), 0, TBCiclo!Dias)
                
                            'Dados do estoque
                            ExcluirAP = True
                            TOK = IIf(IsNull(TBCiclo!Quantidade), 0, TBCiclo!Quantidade)
                            TNC = IIf(IsNull(TBCiclo!Reprovada), 0, TBCiclo!Reprovada)
                            If TOK <> 0 Or TNC <> 0 Then
                                Set TBOrdem = CreateObject("adodb.recordset")
                                TBOrdem.Open "Select * from producao where Ordem = " & TBCiclo!Ordem & " and (Entrar_estoque = 'True' or Retirar_estoque = 'True')", Conexao, adOpenKeyset, adLockOptimistic
                                If TBOrdem.EOF = False Then
                                    Set TBCFOP = CreateObject("adodb.recordset")
                                    TBCFOP.Open "Select * from ordemservico where Ordem = " & TBOrdem!Ordem & " order by fase, retrabalho, IDproducao", Conexao, adOpenKeyset, adLockOptimistic
                                    If TBCFOP.EOF = False Then
                                        If TBOrdem!Retirar_estoque = True Then 'Verfica se é a primeira OS e retira o material do estoque
                                            TBCFOP.MoveFirst
                                            If TBCFOP!IDProducao = TBCiclo!OS Then ProcRetirarCancelarEstoque
                                        End If
                                    End If
                                End If
                                TBOrdem.Close
                            End If
                            
                            TBCiclo!TempoFinal = "00:00:00"
                            TBCiclo!TempoTotal = "00:00:00"
                            TBCiclo!TempoTotalSeg = 0
                            TBCiclo!Quantidade = 0
                            TBCiclo!Reprovada = 0
                            TBCiclo!Dias = 0
                            Descricao = TBCiclo!codigoDesc
                            IDProducao = TBCiclo!IDProducao
                            TBCiclo.Update
                            
                            'Atualiza dados da manutenção
                            TBCiclo.MoveLast
                            Set TBAbrir = CreateObject("adodb.recordset")
                            TBAbrir.Open "Select * from Manutencao_data where idproducao2 = " & TBCiclo!IDProducao, Conexao, adOpenKeyset, adLockOptimistic
                            If TBAbrir.EOF = False Then
                                Set TBLista = CreateObject("adodb.recordset")
                                TBLista.Open "Select * from Manutencao_data where IDmanutencao = " & TBAbrir!IDmanutencao & " order by Data", Conexao, adOpenKeyset, adLockOptimistic
                                If TBLista.BOF = False Then
                                    TBLista.Find ("ID = " & TBAbrir!ID)
                                    TBLista.MoveNext
                                    If TBLista.EOF = False Then
                                        If TBLista!Status = "Aberta" Then
                                            Conexao.Execute "DELETE from Manutencao_Checklist where ID_data = " & TBLista!ID
                                            Conexao.Execute "DELETE from manutencao_data where ID = " & TBLista!ID
                                        End If
                                    End If
                                End If
                                TBLista.Close
                                Conexao.Execute "Update Manutencao_Checklist Set Check = 'False' where ID_data = " & TBAbrir!ID
                            End If
                            TBAbrir.Close
                            Conexao.Execute "Update Manutencao_data Set IDproducao = 0 where idproducao = " & TBCiclo!IDProducao
                            Conexao.Execute "Update Manutencao_data Set status = 'Aberta', IDproducao2 = 0 where idproducao2 = " & TBCiclo!IDProducao
                            
                            '=====================================================================
                            ' Apaga o apontamento
                            '=====================================================================
                            Conexao.Execute "DELETE from " & NomeTabelaAp & " WHERE IDProducao = " & TBCiclo!IDProducao
                            '=====================================================================
                            ' Exclui a entrada no estoque
                            '=====================================================================
                            ProcExcluirEntradaEstoque
                            '=====================================================================
                        End If
                    End If
        
                    'Localiza na tabela producaofases todos os registros desta of, maquina, fase e muda pronta = não
                    Conexao.Execute "Update " & NomeTabelaAp & " Set pronto = 'NÃO' where idfase = " & TBOS!IDProducao & " and Maquina = '" & txtMaquina & "'"
                    'Localiza na tabela a ordem de servico todos os registros desta maquina, ordem e muda pronta = não
                    Conexao.Execute "Update ordemservico Set pronto = 'NÃO' where idproducao = " & TBOS!IDProducao
        
                    TBOS.MoveNext
                Loop
            End If
            TBOS.Close
                    
            ProcVerificaTurno
        
            'Atualiza lista de eventos cadastrados
            ProcAtualizaProducao False, True
            
            'Grava o status da OS e da OF
            ProcGravarStatusOSOF
            
            If Varias_OS = True Then
                TextoFiltro = "Select OrdemServico.* from OrdemServico INNER JOIN ProducaoFases_OS ON OrdemServico.ID_apontamento = ProducaoFases_OS.ID where ProducaoFases_OS.ID = " & Txt_ID_apontamento & " order by OrdemServico.IDproducao"
            Else
                TextoFiltro = "Select * from OrdemServico where IDProducao = " & ListaOS.SelectedItem.ListSubItems(1)
            End If
            Set TBOS = CreateObject("adodb.recordset")
            TBOS.Open TextoFiltro, Conexao, adOpenKeyset, adLockOptimistic
            If TBOS.EOF = False Then
                Do While TBOS.EOF = False
                    'Verifica se a máquina esta sendo utilizada na OS
                    Set TBCiclo = CreateObject("adodb.recordset")
                    TBCiclo.Open "Select IDProducao from " & NomeTabelaAp & " where OS = " & TBOS!IDProducao & " and Maquina = '" & txtMaquina & "' order by Data, Tempoinicio", Conexao, adOpenKeyset, adLockOptimistic
                    If TBCiclo.EOF = True Then
                        Conexao.Execute "DELETE from Ordemservico_maq_utilizadas where OS = " & TBOS!IDProducao & " and maquina = '" & txtMaquina & "'"
                    End If
                    
                    'Verifica se existe algum apontamento (totalização) na OS
                    Set TBCiclo = CreateObject("adodb.recordset")
                    TBCiclo.Open "Select * from " & NomeTabelaAp & " where idfase = " & TBOS!IDProducao & " and usuario = '" & Operador & "' and maquina = '" & txtMaquina.Text & "' and Turno = " & Turno & " and data = '" & Format(Dataini, "Short Date") & "' and (codigodesc = 1 or codigodesc = 2)", Conexao, adOpenKeyset, adLockOptimistic
                    If TBCiclo.EOF = True Then
                        Conexao.Execute "DELETE from " & NomeTabelaApTotalizacao & " where OS = " & TBOS!IDProducao & " and usuario = '" & Operador & "' and maquina = '" & txtMaquina.Text & "' and Turno = " & Turno & " and data = '" & Format(Dataini, "Short Date") & "'"
                    End If
                    TBCiclo.Close
                    
                    TBOS.MoveNext
                Loop
            End If
            TBOS.Close
        
            'Acerta cadastro na máquina
            ProcAcertaCadMaquina
            'Calcula e grava valor real do lote e por peça
            ProcGravaValoresOS
            
            ProcAtualizaCodigoDesc
            ProcCarregaListaOS
            ProcLista12Ultimos
        End If
    End If
End If
ExcluiSel = False
txtCodigo = ""
txtCodigoBarras = ""
cmbdescricao.ListIndex = -1
txtCodigoDesc = ""
QT_Entrada_Estoque = 0

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcDesvincularOS()
On Error GoTo tratar_erro

If Varias_OS = True Then
    If MsgBox("Deseja realmente desvincular a OS " & ListaOS.SelectedItem.ListSubItems(1) & " deste apontamento simultâneo?", vbYesNo) = vbYes Then
        Set TBAbrir = CreateObject("adodb.recordset")
        TBAbrir.Open "Select * from ordemservico where IDProducao = " & ListaOS.SelectedItem.ListSubItems(1), Conexao, adOpenKeyset, adLockOptimistic
        If TBAbrir.EOF = False Then
            Set TBFiltro = CreateObject("adodb.recordset")
            TBFiltro.Open "Select * from ordemservico where IDProducao <> " & TBAbrir!IDProducao & " and ID_apontamento = " & TBAbrir!ID_apontamento, Conexao, adOpenKeyset, adLockOptimistic
            If TBFiltro.EOF = True Then
                Conexao.Execute "DELETE from ProducaoFases_OS where ID = " & TBAbrir!ID_apontamento
            End If
            TBFiltro.Close
            TBAbrir!ID_apontamento = Null
            TBAbrir.Update
        End If
        TBAbrir.Close
        
        MsgBox ("OS " & ListaOS.SelectedItem.ListSubItems(1) & " desvinculada com sucesso."), vbInformation
        ProcCarregaListaOS
    End If
Else
    MsgBox ("Esta função não esta habilitada para este módulo de apontamento."), vbExclamation
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub


Private Sub txtCodigo_Change()
On Error GoTo tratar_erro

    If Ap_codigo = True Then
        cmbdescricao.ListIndex = -1
        If IsNumeric(txtCodigo) = True Then
            If TempoPreparacaoReaprov = True And txtCodigo = 1 Then
                MsgBox ("Não é permitido utilizar este código, pois o tempo de preparação desta OS foi reaproveitado de outra OS."), vbExclamation
                txtCodigo = ""
                Exit Sub
            End If
            Set TBCodigoDesc = CreateObject("adodb.recordset")
            TBCodigoDesc.Open "Select Descricao from CodigoDesc where Codigo = " & txtCodigo & " and Bloqueado = 'False'", Conexao, adOpenKeyset, adLockOptimistic
            If TBCodigoDesc.EOF = False Then
                cmbdescricao.Text = TBCodigoDesc!Descricao
            End If
            TBCodigoDesc.Close
        Else
            txtCodigo = ""
        End If
    End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub txtCodigoBarras_Change()
On Error GoTo tratar_erro

txtCodigoDesc = ""

If Len(txtCodigoBarras) = 6 Then
If IsNumeric(txtCodigoBarras) = True Then
    Set TBProducao = CreateObject("adodb.recordset")
    TBProducao.Open "select M.Codigo from manutencao M INNER JOIN manutencao_data MD ON MD.idManutencao = M.CODIGO where M.IDmaquina = '" & txtMaquina & "' and M.Controlada = 'true' and MD.status = 'Aberta' and MD.IDProducao = 0 and MD.data <= '" & Format(Date, "Short Date") & "'", Conexao, adOpenKeyset, adLockOptimistic
    If TBProducao.EOF = False Then
        Set TBCodigoDesc = CreateObject("adodb.recordset")
        TBCodigoDesc.Open "Select Descricao from codigodesc where codigo = " & txtCodigoBarras & " and Bloqueado = 'False' and (descricao = 'MÁQUINA EM MANUTENÇÃO' or descricao = 'MANUTENÇÃO PREVENTIVA' or descricao = 'MANUTENÇÃO CORRETIVA')", Conexao, adOpenKeyset, adLockOptimistic
        If TBCodigoDesc.EOF = False Then
            txtCodigoDesc = TBCodigoDesc!Descricao
        Else
            MsgBox ("Só é permitido apontar um desses eventos (MÁQUINA EM MANUTENÇÃO, MANUTENÇÃO PREVENTIVA ou MANUTENÇÃO CORRETIVA)."), vbExclamation
            txtCodigoBarras = ""
        End If
        TBCodigoDesc.Close
        Exit Sub
    End If
    If txtCodigoBarras = 1 Then
        If TempoPreparacaoReaprov = True Then
            Permitido = False
            Set TBProducao = CreateObject("adodb.recordset")
            TBProducao.Open "Select IDProducao from ProducaoFases where OS = " & ListaOS.SelectedItem.ListSubItems(1), Conexao, adOpenKeyset, adLockOptimistic
            If TBProducao.EOF = False Then
                If TBProducao.RecordCount >= 2 Then Permitido = True Else Permitido = False
            End If
            If Permitido = False Then
                MsgBox ("Não é permitido utilizar este código, pois o tempo de preparação desta OS foi reaproveitado de outra OS."), vbExclamation
                txtCodigoBarras = ""
                Exit Sub
            End If
        End If
    End If
    Set TBCodigoDesc = CreateObject("adodb.recordset")
    TBCodigoDesc.Open "Select Descricao from codigodesc where codigo = " & txtCodigoBarras & " and Bloqueado = 'False'", Conexao, adOpenKeyset, adLockOptimistic
    If TBCodigoDesc.EOF = False Then
        
        
0
    txtCodigoDesc = TBCodigoDesc!Descricao
        'Txt_codigoF.SetFocus
    End If
Else
    txtCodigoBarras = ""
End If

End If

cmbdescricao.Visible = False

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Sub ProcLista12Ultimos()
On Error GoTo tratar_erro
Dim TempoTotalDias As Long
Dim CodEvento As Long

Lista.ListItems.Clear
Set TBProducao = CreateObject("adodb.recordset")
TBProducao.Open "Select * from " & NomeTabelaAp & " where idfase = " & ListaOS.SelectedItem.ListSubItems(1) & " and Maquina = '" & txtMaquina & "'  order by Data, Tempoinicio", Conexao, adOpenKeyset, adLockOptimistic
If TBProducao.EOF = False Then
    If TBProducao.RecordCount > 12 Then
        TBProducao.MoveLast
        Contador2 = 12
        Do While Contador2 > 1
            TBProducao.MovePrevious
            Contador2 = Contador2 - 1
        Loop
    End If
    Do Until TBProducao.EOF
        CodEvento = IIf(IsNull(TBProducao!codigoDesc), 0, TBProducao!codigoDesc)
        With Lista.ListItems
            .Add , , TBProducao("IDProducao")
            .Item(.Count).SubItems(1) = IIf(IsNull(TBProducao!codigoDesc), 0, TBProducao!codigoDesc)
            .Item(.Count).SubItems(2) = TBProducao!Descricao
            txtstatus.Text = TBProducao!Descricao
            .Item(.Count).SubItems(3) = TBProducao!TempoInicio
            .Item(.Count).SubItems(4) = TBProducao!TempoFinal
            If TBProducao!TempoInicio <> "00:00:00" Then
                TempoFinal = TBProducao!TempoInicio
                TempoUtilizadoDescricao = TBProducao!TempoInicio
            End If
            If TBProducao!Dias <> 0 Then
                TempoTotalDias = IIf(IsNull(TBProducao!TempoTotal), 0, TBProducao!TempoTotal) + TBProducao!Dias
                FunElapsedTime (TempoTotalDias)
                .Item(.Count).SubItems(5) = Horatotal
            Else
                .Item(.Count).SubItems(5) = IIf(IsNull(TBProducao!TempoTotal), "", TBProducao!TempoTotal)
            End If
            .Item(.Count).SubItems(6) = TBProducao!Usuario
            .Item(.Count).SubItems(7) = IIf(IsNull(TBProducao!Quantidade), 0, TBProducao!Quantidade)
            .Item(.Count).SubItems(8) = IIf(IsNull(TBProducao!QTCD), 0, TBProducao!QTCD)
            
            .Item(.Count).SubItems(9) = IIf(IsNull(TBProducao!Reprovada), 0, TBProducao!Reprovada)
            .Item(.Count).Selected = True
        End With
        TBProducao.MoveNext
    Loop
End If
TBProducao.Close

If CodEvento = 2 Then
    cmdF8.Enabled = True
Else
    cmdF8.Enabled = False
End If


Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Sub ProcCarregaListaRequisicao()
On Error GoTo tratar_erro
Dim Requisitado As Double
Dim Saida As Double


ListaRequisicao.ListItems.Clear
If Ordem = 0 Then Exit Sub
Set TBLista = CreateObject("adodb.recordset")
StrSQL = "Select PM.obs, PM.Tipo_Item, PM.Idmateriaprima, PM.CODIGO, PM.Unidade, PM.Descricao, PM.Requisitado from Producaomaterial PM  Where PM.ordem = '" & Ordem & "'"
''Debug.Print StrSql

TBLista.Open StrSQL, Conexao, adOpenKeyset, adLockOptimistic
If TBLista.EOF = False Then
    Contador = 0
    Do While TBLista.EOF = False
        With ListaRequisicao.ListItems
            .Add , , TBLista!IdMateriaPrima
            .Item(.Count).SubItems(1) = IIf(IsNull(TBLista!Codigo), "", TBLista!Codigo)
            .Item(.Count).SubItems(2) = IIf(IsNull(TBLista!Descricao), "", TBLista!Descricao)
            .Item(.Count).SubItems(3) = IIf(IsNull(TBLista!Unidade), "", TBLista!Unidade)
            .Item(.Count).SubItems(4) = IIf(IsNull(TBLista!Requisitado), "", TBLista!Requisitado)
            .Item(.Count).SubItems(5) = IIf(IsNull(TBLista!Obs), "", TBLista!Obs)
        End With
        TBLista.MoveNext
        Contador = Contador + 1

    Loop
End If
TBLista.Close


Exit Sub
tratar_erro:
    USMsgBox ("Descrição do erro : " + Error()), vbCritical, "CAPRIND v5.0"
End Sub

Private Sub txtCodigoBarras_GotFocus()
    cmbdescricao.Visible = False

End Sub






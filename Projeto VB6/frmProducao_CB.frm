VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmProducao_CB 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gerprod - Coletor de dados no chão de fábrica - Código de barras"
   ClientHeight    =   8025
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11970
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
   ForeColor       =   &H00800000&
   Icon            =   "frmProducao_CB.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   8025
   ScaleWidth      =   11970
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame10 
      BackColor       =   &H00E0E0E0&
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
      Left            =   9435
      TabIndex        =   52
      Top             =   1140
      Width           =   855
      Begin VB.TextBox txtturno 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   405
         Left            =   150
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   270
         Width           =   570
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Turno"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   113
         TabIndex        =   53
         Top             =   0
         Width           =   645
      End
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   10305
      TabIndex        =   62
      Top             =   1140
      Width           =   1635
      Begin VB.TextBox txteficiencia 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   405
         Left            =   90
         Locked          =   -1  'True
         TabIndex        =   10
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   270
         Width           =   1110
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "%"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   495
         Left            =   1320
         TabIndex        =   64
         Top             =   330
         Width           =   225
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Eficiência"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   68
         TabIndex        =   63
         Top             =   0
         Width           =   1155
      End
   End
   Begin VB.Frame Frame14 
      BackColor       =   &H00E0E0E0&
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
      Left            =   3855
      TabIndex        =   44
      Top             =   1140
      Width           =   5565
      Begin VB.TextBox txtquant 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   405
         Left            =   90
         Locked          =   -1  'True
         TabIndex        =   5
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   270
         Width           =   1125
      End
      Begin VB.TextBox txtttok 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   405
         Left            =   1240
         Locked          =   -1  'True
         TabIndex        =   6
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   270
         Width           =   1290
      End
      Begin VB.TextBox Txtttnc 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   405
         Left            =   2555
         Locked          =   -1  'True
         TabIndex        =   7
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   270
         Width           =   1500
      End
      Begin VB.TextBox txtProduzida 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   405
         Left            =   4117
         Locked          =   -1  'True
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   "0"
         Top             =   270
         Width           =   1380
      End
      Begin VB.Label Label14 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Lote"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   330
         TabIndex        =   48
         Top             =   0
         Width           =   645
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Aprovadas  +"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   1275
         TabIndex        =   47
         Top             =   0
         Width           =   1350
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Não conf.   ="
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   240
         Left            =   2880
         TabIndex        =   46
         Top             =   0
         Width           =   1245
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Produzidas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   4200
         TabIndex        =   45
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.Frame Frame12 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tempos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   795
      Left            =   3855
      TabIndex        =   82
      Top             =   1860
      Width           =   1035
      Begin VB.Label Label29 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Previstos =>"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   83
         Top             =   330
         Width           =   855
      End
   End
   Begin VB.Frame Frame6 
      BackColor       =   &H00E0E0E0&
      Height          =   765
      Left            =   4905
      TabIndex        =   68
      Top             =   1890
      Width           =   7035
      Begin VB.TextBox txtexecucao 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   14
         TabStop         =   0   'False
         Text            =   "00:00:00"
         Top             =   420
         Width           =   1155
      End
      Begin VB.TextBox txtpreparacao 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   150
         Locked          =   -1  'True
         TabIndex        =   13
         TabStop         =   0   'False
         Text            =   "00:00:00"
         Top             =   420
         Width           =   1065
      End
      Begin VB.TextBox txttotal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   5730
         Locked          =   -1  'True
         TabIndex        =   17
         TabStop         =   0   'False
         Text            =   "00:00:00"
         Top             =   420
         Width           =   1155
      End
      Begin VB.TextBox TxtA3Prevista 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   16
         TabStop         =   0   'False
         Text            =   "00:00:00"
         Top             =   420
         Width           =   1485
      End
      Begin VB.TextBox txtPcHoraPrevista 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         ForeColor       =   &H00FF0000&
         Height          =   270
         Left            =   2790
         Locked          =   -1  'True
         MaxLength       =   5
         MouseIcon       =   "frmProducao_CB.frx":1042
         MousePointer    =   99  'Custom
         TabIndex        =   15
         TabStop         =   0   'False
         Text            =   "1"
         ToolTipText     =   "Total de peças por tempo de execução prevista."
         Top             =   420
         Width           =   855
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Execução"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   0
         Left            =   1447
         TabIndex        =   74
         Top             =   150
         Width           =   900
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Preparação"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   120
         TabIndex        =   73
         Top             =   150
         Width           =   1125
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Tempo total"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   5730
         TabIndex        =   72
         Top             =   150
         Width           =   1155
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Pç(s) x exec."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   5
         Left            =   2580
         TabIndex        =   71
         Top             =   150
         Width           =   1245
      End
      Begin VB.Label Label45 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Execução x peça"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   4027
         TabIndex        =   70
         Top             =   150
         Width           =   1590
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   " /                   ="
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   270
         Index           =   6
         Left            =   2580
         TabIndex        =   69
         Top             =   420
         Width           =   1410
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      Height          =   765
      Left            =   15
      TabIndex        =   65
      Top             =   1890
      Width           =   3825
      Begin VB.TextBox txtData 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   405
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Text            =   "00/00/00"
         Top             =   240
         Width           =   1425
      End
      Begin VB.TextBox txtHora 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   405
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   12
         TabStop         =   0   'False
         Text            =   "00:00:00"
         Top             =   240
         Width           =   2115
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Data"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   450
         TabIndex        =   67
         Top             =   -30
         Width           =   765
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Hora"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   2175
         TabIndex        =   66
         Top             =   -30
         Width           =   885
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00E0E0E0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   795
      Left            =   4905
      TabIndex        =   75
      Top             =   2640
      Width           =   7035
      Begin VB.TextBox txtTEUTIL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   1320
         Locked          =   -1  'True
         TabIndex        =   21
         TabStop         =   0   'False
         Text            =   "00:00:00"
         Top             =   450
         Width           =   1155
      End
      Begin VB.TextBox txtTPUTIL 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   180
         Locked          =   -1  'True
         TabIndex        =   20
         TabStop         =   0   'False
         Text            =   "00:00:00"
         Top             =   450
         Width           =   1065
      End
      Begin VB.TextBox txtTTReal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   5730
         Locked          =   -1  'True
         TabIndex        =   24
         TabStop         =   0   'False
         Text            =   "00:00:00"
         Top             =   450
         Width           =   1155
      End
      Begin VB.TextBox TxtA3Real 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   23
         TabStop         =   0   'False
         Text            =   "00:00:00"
         Top             =   450
         Width           =   1485
      End
      Begin VB.TextBox txtPcHorareal 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
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
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   2820
         Locked          =   -1  'True
         MaxLength       =   5
         MouseIcon       =   "frmProducao_CB.frx":134C
         MousePointer    =   99  'Custom
         TabIndex        =   22
         TabStop         =   0   'False
         Text            =   "1"
         ToolTipText     =   "Total de peças por tempo de execução utilizado."
         Top             =   450
         Width           =   855
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Execução"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   1447
         TabIndex        =   81
         Top             =   180
         Width           =   900
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Preparação"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   150
         TabIndex        =   80
         Top             =   180
         Width           =   1125
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Tempo total"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   5730
         TabIndex        =   79
         Top             =   180
         Width           =   1155
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Pç(s) x exec."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Index           =   1
         Left            =   2550
         TabIndex        =   78
         Top             =   180
         Width           =   1245
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Execução x peça"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   4027
         TabIndex        =   77
         Top             =   180
         Width           =   1590
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         BackStyle       =   0  'Transparent
         Caption         =   " /                   ="
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Index           =   2
         Left            =   2580
         TabIndex        =   76
         Top             =   450
         Width           =   1410
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E0E0E0&
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
      Left            =   15
      TabIndex        =   58
      Top             =   1140
      Width           =   3825
      Begin VB.TextBox txtos 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   405
         Left            =   1447
         Locked          =   -1  'True
         TabIndex        =   3
         TabStop         =   0   'False
         Text            =   "0000000"
         Top             =   270
         Width           =   1065
      End
      Begin VB.TextBox txtof 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   405
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   2
         TabStop         =   0   'False
         Text            =   "00000000"
         Top             =   270
         Width           =   1305
      End
      Begin VB.TextBox txtprazo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   405
         Left            =   2535
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Text            =   "00/00/00"
         Top             =   270
         Width           =   1155
      End
      Begin VB.Label Label13 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "O.S"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   1725
         TabIndex        =   61
         Top             =   0
         Width           =   495
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Prazo final"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   2542
         TabIndex        =   60
         Top             =   0
         Width           =   1140
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "N° Ordem"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   240
         TabIndex        =   59
         Top             =   0
         Width           =   1065
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   15
      TabIndex        =   55
      Top             =   3450
      Width           =   6465
      Begin VB.TextBox txtcodigo 
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
         TabIndex        =   25
         Top             =   300
         Width           =   1275
      End
      Begin VB.TextBox txtcodigoDesc 
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
         Left            =   1410
         Locked          =   -1  'True
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   300
         Width           =   4935
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Código"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   435
         TabIndex        =   57
         Top             =   30
         Width           =   645
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "Descrição"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   2955
         TabIndex        =   56
         Top             =   0
         Width           =   1845
      End
   End
   Begin VB.Frame Frame4 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Conforme"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   855
      Left            =   9345
      TabIndex        =   54
      Top             =   3450
      Width           =   1245
      Begin VB.TextBox txtTOK 
         Alignment       =   2  'Center
         BackColor       =   &H80000014&
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
         Height          =   460
         Left            =   135
         TabIndex        =   27
         ToolTipText     =   "Qtde. conforme."
         Top             =   270
         Visible         =   0   'False
         Width           =   960
      End
   End
   Begin VB.Frame Frame11 
      BackColor       =   &H00E0E0E0&
      Height          =   765
      Left            =   15
      TabIndex        =   49
      Top             =   2670
      Width           =   3825
      Begin VB.TextBox txtMaquina 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   405
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   240
         Width           =   2115
      End
      Begin VB.TextBox txtFase 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   405
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   18
         TabStop         =   0   'False
         Text            =   "00"
         Top             =   240
         Width           =   1425
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Posto de trabalho"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   1657
         TabIndex        =   51
         Top             =   -30
         Width           =   1920
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         Caption         =   "Fase"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   495
         TabIndex        =   50
         Top             =   0
         Width           =   675
      End
   End
   Begin VB.Frame Frame15 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Não conf."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   855
      Left            =   10605
      TabIndex        =   43
      Top             =   3450
      Width           =   1335
      Begin VB.TextBox txtTNC 
         Alignment       =   2  'Center
         BackColor       =   &H80000014&
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
         Height          =   460
         Left            =   165
         TabIndex        =   28
         ToolTipText     =   "Qtde. não conforme."
         Top             =   270
         Visible         =   0   'False
         Width           =   960
      End
   End
   Begin VB.Frame Frame16 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6495
      TabIndex        =   41
      Top             =   3450
      Width           =   2835
      Begin VB.Label Label16 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Itens produzidos =>"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   180
         TabIndex        =   42
         Top             =   330
         Width           =   2565
      End
   End
   Begin VB.Frame Frame13 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Tempos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   795
      Left            =   3855
      TabIndex        =   39
      Top             =   2640
      Width           =   1035
      Begin VB.Label Label30 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Reais =>"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   120
         TabIndex        =   40
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame9 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Status do posto de trabalho"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1095
      Left            =   7200
      TabIndex        =   36
      Top             =   30
      Width           =   4755
      Begin VB.TextBox txtstatus 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   570
         Width           =   3405
      End
      Begin VB.TextBox TxtTempoUtilizado 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   3540
         Locked          =   -1  'True
         TabIndex        =   1
         TabStop         =   0   'False
         Text            =   "00:00:00"
         Top             =   570
         Width           =   1125
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Tempo"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   3787
         TabIndex        =   38
         Top             =   330
         Width           =   630
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Status"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   1500
         TabIndex        =   37
         Top             =   330
         Width           =   645
      End
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   915
      Left            =   300
      Picture         =   "frmProducao_CB.frx":1656
      ScaleHeight     =   915
      ScaleWidth      =   1005
      TabIndex        =   35
      Top             =   90
      Width           =   1005
   End
   Begin VB.Frame Frame5 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Menu de comandos"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1095
      Left            =   1560
      TabIndex        =   32
      Top             =   30
      Width           =   5640
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "  F11 - LISTAR TODOS EVENTOS // F12 - LISTAR OS 7 ULTIMOS"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   150
         TabIndex        =   34
         Top             =   720
         Width           =   5130
      End
      Begin VB.Label Label15 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00E0E0E0&
         BackStyle       =   0  'Transparent
         Caption         =   " F3 - GRAVAR EVENTO //  F4  - EXCLUIR EVENTO //  F6 - VOLTAR"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   33
         Top             =   390
         Width           =   5190
      End
   End
   Begin VB.TextBox txtdesenho 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Left            =   285
      Locked          =   -1  'True
      TabIndex        =   31
      Top             =   1590
      Visible         =   0   'False
      Width           =   1695
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   975
      Top             =   5775
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
            Picture         =   "frmProducao_CB.frx":2BDA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   10560
      Top             =   6000
   End
   Begin VB.TextBox txtdescricao 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00E0E0E0&
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
      Left            =   1965
      Locked          =   -1  'True
      TabIndex        =   30
      Top             =   1590
      Visible         =   0   'False
      Width           =   9495
   End
   Begin MSComctlLib.ListView Lista 
      Height          =   3435
      Left            =   15
      TabIndex        =   29
      Top             =   4320
      Width           =   11940
      _ExtentX        =   21061
      _ExtentY        =   6059
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      SmallIcons      =   "ImageList1"
      ForeColor       =   0
      BackColor       =   16777215
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   10
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "IDProducao"
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Cód."
         Object.Width           =   0
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Descrição"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Início"
         Object.Width           =   3440
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Final"
         Object.Width           =   3440
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   5
         Text            =   "Tempo total"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Operador"
         Object.Width           =   3440
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   7
         Text            =   "Aprov."
         Object.Width           =   1676
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   8
         Text            =   "NC"
         Object.Width           =   1676
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   9
         Text            =   "Pronta"
         Object.Width           =   0
      EndProperty
   End
   Begin MSComctlLib.ProgressBar PBLista 
      Height          =   255
      Left            =   15
      TabIndex        =   84
      Top             =   7770
      Width           =   11940
      _ExtentX        =   21061
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.Shape Shape1 
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   1065
      Left            =   30
      Top             =   30
      Width           =   1515
   End
End
Attribute VB_Name = "frmProducao_CB"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public IDProcesso As Long
Public IDFase As Long
Public IDProducao As Long
Public ExcluiSel As Boolean
Public Quant As Integer
Public Produzidas As Integer
Dim Dataini As Date
Dim DiaSemana As String
Dim UltimoDesc As String
Dim PenultimoDesc As String
Dim TempoUtilizadoDescricao As Date

Private Sub ProcVerificaUltimo()
On Error GoTo tratar_erro

'Se o ultimo código de trabalho for igual ao próximo código de trabalho, não aceita
If txtcodigoDesc <> "TROCA DE POSTO DE TRABALHO" Then
    If Ultimo = txtCodigo.Text Then
        MsgBox ("Só é permitido cadastro de próximo evento, diferente do anterior."), vbExclamation
        Gravar = False
        Exit Sub
    End If
End If
'Verifica se foi alterado operador sem o evento troca de operador
If Ultimo = 2 Then
    Set TBProducao = BD.OpenRecordset("Select * from producaofases where idfase = " & txtos.Text & " order by Data, Tempoinicio")
    If TBProducao.EOF = False Then
        TBProducao.MoveLast
        If TBProducao!usuario <> Operador And TBProducao!Descricao <> "TROCA DE OPERADOR" And TBProducao!Descricao <> "FIM DE TURNO" Then
            MsgBox ("É necessário o operador " & TBProducao!usuario & " cadastrar o envento TROCA DE OPERADOR antes de salvar."), vbExclamation
            Gravar = False
            Exit Sub
        End If
    End If
    TBProducao.Close
End If
If txtCodigo.Text = 3 Then
    If Ultimo <> 2 Then
        MsgBox ("Só é permitido encerrar a ordem após código Máquina em produção."), vbExclamation
        Gravar = False
        Exit Sub
    End If
    'Se o ultimo código de trabalho for nenhum e o operador quiser terminar o lote não aceita
    If Ultimo = 0 Then
        MsgBox ("Só é permitido encerrar a ordem após termino do lote."), vbExclamation
        Gravar = False
        Exit Sub
    End If
End If
'Se o código de trabalho for maquina em produção deverá ser colocado a quantidade peças produzidas
If Ultimo = 2 And txtTOK.Text = "" Then
    'Se os dados estiverem corretamente preenchidos
    txtTOK.Visible = True
    txtTNC.Visible = True
    Label16.Visible = True
    MsgBox ("É obrigatório colocar a quantidade de peças produzidas aprovadas, e se houver as não conforme."), vbExclamation
    txtTOK.SetFocus
    Gravar = False
    Exit Sub
End If
'Se o ultimo código de trabalho for fim de produção, fase já concluida
If Ultimo = 3 Then
    MsgBox ("Esta ordem de serviço já está concluída."), vbExclamation
    Gravar = False
    Exit Sub
End If
'Verificar se o turno final e igual ao turno inicio do evento
ProcVerificaTurno
If Ultimo = "" Then Ultimo = 0
Gravar = True

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
Dataini = txtHora.Text
Set TBMaquinas = BD.OpenRecordset("Select * from CadMaqturnos where maquina = '" & txtMaquina.Text & "' and diasemana = '" & DiaSemana & "' order by diasemana,turno")
If TBMaquinas.EOF = False Then
    Do While TBMaquinas.EOF = False
        TempoInicio = TBMaquinas!Inicioturno
        TempoFinal = TBMaquinas!finalturno
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

Private Sub ProcVerificaDia()
On Error GoTo tratar_erro

DiaSemana = Weekday(Dataini)
Select Case DiaSemana
    Case 1:
        DiaSemana = "Domingo"
    Case 2:
        DiaSemana = "Segunda"
    Case 3:
        DiaSemana = "Terça"
    Case 4:
        DiaSemana = "Quarta"
    Case 5:
        DiaSemana = "Quinta"
    Case 6:
        DiaSemana = "Sexta"
    Case 7:
        DiaSemana = "Sabado"
End Select

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcAtualizaPrepUtil()
On Error GoTo tratar_erro

'Calcula total de segundos utilizados preparando
ElapsedTime (TempoTotalPrep)
TPUSEG = S
TTPUTIL = Horatotal
Set TBOrdemServico = BD.OpenRecordset("Select * from ordemservico where IDPRODUCAO = " & txtos.Text & "")
If TBOrdemServico.EOF = False Then
    TBOrdemServico.Edit
    TBOrdemServico!TPUTIL = TTPUTIL
    TBOrdemServico!TPUSEG = TPUSEG
    TBOrdemServico.Update
End If
TBOrdemServico.Close
txtTPUTIL = TTPUTIL

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcAtualizaExecUtil()
On Error GoTo tratar_erro

Set TBOrdemServico = BD.OpenRecordset("Select * from ordemservico where IDPRODUCAO = " & txtos & "")
If TBOrdemServico.EOF = False Then
    TBOrdemServico.Edit
    'Calcula tempo total de execução
    ElapsedTime (TempoTotalProd)
    TTEUTIL = Horatotal
    TTEUTILS = S
    
    'Soma tempo total de preparação + execução
    TempoTotalUtil = TempoTotalPrep + TempoTotalProd
    ElapsedTime (TempoTotalUtil)
    TTUTILSEG = S
    txtTTReal.Text = Horatotal
    TBOrdemServico!TETTUTIL = txtTTReal.Text 'Grava tempo total utilizado
    TBOrdemServico!TETTUTILN = TempoTotalUtil 'Grava tempo total utilizado
    TBOrdemServico!TETTUTILSEG = TTUTILSEG 'Grava tempo total utilizado em segundos
    
    'Calcula tempo total de execução por peça
    Produzidas = TTOK + TTNC
    If TempoTotalProd > 0 And Produzidas > 0 Then TEUSEG = TTEUTILS / Produzidas Else TEUSEG = TTEUTILS
    DecimoSegundos = TEUSEG
    TBOrdemServico!TEUTIL = FormataTempo(DecimoSegundos)
    TBOrdemServico!TEUSEG = TEUSEG
    TxtA3Real = TBOrdemServico!TEUTIL 'Carrega tempo de execução x peça
        
    'Verif. se a OS é peças por hora e carrega tempo de execução utilizado por peça
    If TBOrdemServico!pecahora = True Then
        If Produzidas <> 0 Then txtTEUTIL = "01:00:00" Else txtTEUTIL = "00:00:00"
        If TEUSEG <> 0 Then txtPcHorareal = 3600 / TEUSEG Else txtPcHorareal = 1
    Else
        txtTEUTIL.Text = TTEUTIL
        txtPcHorareal = 1
    End If
    
    TBOrdemServico!QTOK = TTOK
    TBOrdemServico!QTNC = TTNC
        
    'Calcula eficiencia
    TEPSEG = IIf(IsNull(TBOrdemServico!TESegundos), 0, TBOrdemServico!TESegundos) 'Verif. tempo de execução previsto por peça
    If TEPSEG > 0 And TEUSEG > 0 And Produzidas > 0 Then Eficiencia = Format(TEPSEG / TEUSEG * 100, "###,##0.00") Else Eficiencia = 0
    txteficiencia.Text = Eficiencia
    TBOrdemServico!Eficiencia = IIf(IsNumeric(Eficiencia) = True, Eficiencia, "0")
    
    TBOrdemServico!Totalprod = IIf(IsNumeric(txtProduzida.Text) = True, txtProduzida, "0")
    
    'Calcula e grava valor real do lote e por peça
    Set TBMaquinas = BD.OpenRecordset("Select * from cadmaquinas where maquina = '" & txtMaquina.Text & "'")
    If TBMaquinas.EOF = False Then
        Valorhora = TBMaquinas!PrecoHora
        Valorhora = Valorhora / 3600
        TBOrdemServico!CRLOTE = Format((Valorhora * TEUSEG * (TTOK + TTNC)) + (Valorhora * TPUSEG), "###,##0.00")
        If TEUSEG > 0 And (TTOK + TTNC) > 0 Then TBOrdemServico!CRPECA = Format(TBOrdemServico!CRLOTE / (TTOK + TTNC), "###,##0.00000") Else TBOrdemServico!CRPECA = 0
    Else
        Valorhora = 0
    End If
    TBMaquinas.Close
    TBOrdemServico.Update
End If
TBOrdemServico.Close

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcCalculaEficienciaOF()
On Error GoTo tratar_erro

TEPSEG = 0
Eficiencia = 0

'Verifica qtde de peças produzidas na ordem
Produzidas = IIf(IsNull(TBProducao!QuantProd), 0, TBProducao!QuantProd)

'Verif. total de segundos previstos por peça
Set TBAbrir = BD.OpenRecordset("Select * FROM ORDEMSERVICO WHERE OF = " & txtof.Text & " and custos = true")
If TBAbrir.EOF = False Then
    Do While TBAbrir.EOF = False
        TEPSEG = TEPSEG + IIf(IsNull(TBAbrir!TESegundos), 0, TBAbrir!TESegundos)
        TBAbrir.MoveNext
    Loop
End If
TBAbrir.Close
DecimoSegundos = TEPSEG
TBProducao!TPP = FormataTempo(DecimoSegundos)

If txtquant <> 0 Then TBProducao!cpp = IIf(IsNull(TBProducao!CTTPrev), 0, TBProducao!CTTPrev) / txtquant

If TEPSEG > 0 And TEUSEG > 0 And Produzidas <> 0 Then
    Eficiencia = TEPSEG / TEUSEG * 100
    Contador1 = Len(Eficiencia) - 5
    Eficiencia = Left(Eficiencia, Len(Eficiencia) - Contador1)
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcGravarEvento()
On Error GoTo tratar_erro

Contador = 0
If Gravar = False Then Exit Sub
If OSControlada Or Processo_controlado = True Then
    'Filtra todas as ordens de servico desta of, na tabela ordemservico para verificar se as anteriores foram concluidas
    If Processo_controlado = True Then
        Set TBOrdemServico = BD.OpenRecordset("Select * from ordemservico where of = " & txtof.Text & " and fase < " & Int(txtFase) & " AND PRONTO = 'NÃO'")
        If (TBOrdemServico.BOF And TBOrdemServico.EOF) = False Then
            MsgBox ("PROCESSO CONTROLADO, existe OS anterior que não está concluída, favor verificar."), vbExclamation
            Gravar = False
            TBOrdemServico.Close
            Exit Sub
        End If
    Else
        If txtCodigo = 3 Then
            Set TBOrdemServico = BD.OpenRecordset("Select * from ordemservico where of = " & txtof.Text & " and fase < " & Int(txtFase) & " AND PRONTO = 'NÃO'")
            If (TBOrdemServico.BOF And TBOrdemServico.EOF) = False Then
                MsgBox ("ORDEM CONTROLADA, existe OS anterior que não está concluída, favor verificar."), vbExclamation
                Gravar = False
                TBOrdemServico.Close
                Exit Sub
            End If
        End If
    End If
End If
Produzidas = IIf(txtProduzida = "", 0, txtProduzida)
TOK = IIf(txtTOK = "", 0, txtTOK)
TNC = IIf(txtTNC = "", 0, txtTNC)
'Verifica se a quantidade produzida é maior que o lote previsto
If txtTOK.Text <> "" Then
    Produzidas = Produzidas + TOK + TNC
    If OSControlada = True Then
        If Produzidas > txtquant.Text Then
            MsgBox ("ORDEM CONTROLADA, a quantidade de peças produzidas e maior que a especificada pela ordem."), vbExclamation
            Gravar = False
            Exit Sub
        End If
    End If
    txtTOK.Text = ""
    txtTNC.Text = ""
End If
'Se for encerrar a ordem com quantidade menor que o programado
If OSControlada = True Then
    If txtCodigo.Text = 3 Then
        If Produzidas < txtquant.Text Then
            MsgBox ("ORDEM CONTROLADA, a quantidade de peças produzidas é menor que a especificada pela ordem."), vbExclamation
            Produzidas = Produzidas - TTOK - TTNC - TOK - TNC
            Gravar = False
            Exit Sub
        End If
    End If
End If
'Se for encerrar o evento com quantidade produzida igual ao lote não permite
If OSControlada = True Then
    If txtCodigo.Text <> 3 Then
        If Produzidas = txtquant.Text Then
            MsgBox (" **** Atenção! **** " & vbCrLf & "    " & txtProduzida & " --- Produzidas " & vbCrLf & "      " & TOK & " --- Conforme " & vbCrLf & "+     " & TNC & " --- não conforme " & vbCrLf & "======" & vbCrLf & "     " & txtquant & vbCrLf & "Igual ao lote." & vbCrLf & "  Só é permitido encerrar ordem."), vbExclamation
            Gravar = False
            Exit Sub
        End If
    End If
End If
'Filtra ordem de fabricação na tabela produçao
Set TBOrdem = BD.OpenRecordset("Select * from Producao where of = " & txtof.Text & "")
If txtFase.Text = "" Then txtFase.Text = 0
'Filtra todos os códigos de trabalho da OS
Set TBProducao = BD.OpenRecordset("Select * from ProducaoFases where IDFase = " & txtos & " and CODIGODESC = " & Ultimo & " order by Data, Tempoinicio")
'Se não houver evento gravado com as caracteristicas acima, cria um novo
If (TBProducao.BOF And TBProducao.EOF) = True Then
    TBProducao.AddNew
    TBProducao("of") = txtof.Text
    TBProducao("fase") = txtFase.Text
    TBProducao("maquina") = txtMaquina.Text
    TBProducao("Quantidade") = IIf(txtTOK.Text <> "", txtTOK.Text, 0)
    TBProducao("usuario") = Operador
    TBProducao("descricao") = txtcodigoDesc.Text
    TBProducao!CodigoDesc = txtCodigo.Text
    'Se não foi escolhida uma descrição do código de trabalho
    TBProducao("tempoinicio") = Now
    TBProducao("tempofinal") = 0
    TBProducao("tempototal") = 0
    EVENTO = TBProducao!CodigoDesc
    EVENTO = TBProducao!CodigoDesc
    TBProducao("pronto") = "NÃO"
    TBProducao("preparacao") = IIf(IsDate(txtpreparacao.Text) = True, txtpreparacao.Text, "00:00:00")
    TBProducao("execucao") = IIf(IsDate(txtexecucao.Text) = True, txtexecucao.Text, "00:00:00")
    TBProducao("data") = Date
    TBProducao("quant") = TBOrdem("quant")
    IDApontamento = TBProducao!IDProducao
    TBProducao!IDFase = txtos.Text
    TBProducao!OS = txtos.Text
    TBProducao!Turno = Turno
    TBProducao.Update
    txtTOK.Visible = False
Else
    'Após filtrar move o ponteiro para o ultimo registro
    TBProducao.MoveLast
    TempoInicio = TBProducao("tempoinicio")
    TempoFinal = Now
    TempoTotal = TempoFinal - TempoInicio
    ElapsedTime (TempoTotal)
    TempoTotal = Format(TempoTotal, "hh:mm:ss")
    'HaBilita o modo de edição
    TBProducao.Edit
    TBProducao("tempofinal") = TempoFinal
    TBProducao("tempototal") = TempoTotal
    TBProducao("dias") = D
    TBProducao("Quantidade") = TOK
    TBProducao!reprovada = TNC
    TBProducao.Update
    TBProducao.AddNew
    TBProducao("of") = txtof.Text
    TBProducao("fase") = Int(txtFase.Text)
    TBProducao("maquina") = txtMaquina.Text
    TBProducao("usuario") = Operador
    TBProducao("descricao") = txtcodigoDesc.Text
    TBProducao!CodigoDesc = txtCodigo.Text
    TBProducao("tempoinicio") = Now
    TBProducao("tempofinal") = 0
    TBProducao("tempototal") = 0
    TBProducao("pronto") = "NÃO"
    TBProducao("preparacao") = IIf(IsDate(txtpreparacao.Text) = True, txtpreparacao.Text, "00:00:00")
    TBProducao("execucao") = IIf(IsDate(txtexecucao.Text) = True, txtexecucao.Text, "00:00:00")
    TBProducao("data") = Date
    TBProducao("quant") = TBOrdem("quant")
    IDApontamento = TBProducao!IDProducao
    EVENTO = TBProducao!CodigoDesc
    TBProducao("dias") = 0
    TBProducao!IDFase = txtos.Text
    TBProducao!OS = txtos.Text
    TBProducao!Turno = Turno
    TBProducao.Update
    txtTOK.Visible = False
    Label16.Visible = False
    txtTNC.Visible = False
    txtTOK.Text = ""
    txtTNC.Text = ""
End If
TBProducao.Close
'Grava dados de não conformidade na tabela CQ_NC_FABRICA
If TNC > 0 Then
    Set TBProcessos = BD.OpenRecordset("Select * from ProducaoFases where idfase = " & txtos.Text & " order by Data, Tempoinicio")
    If TBProcessos.EOF = False Then
        TBProcessos.MoveLast
        TBProcessos.MovePrevious
        Set TBCQ = BD.OpenRecordset("CQ_NC_FABRICA")
        TBCQ.AddNew
        TBCQ!IDProducao = TBProcessos!IDProducao
        TBCQ!OF = txtof
        TBCQ!OS = txtos.Text
        TBCQ!TTNC = TNC
        Set TBUsuarios = BD.OpenRecordset("Select * from Usuarios where Usuario = '" & PubUsuario & "'")
        If TBUsuarios.EOF = False Then
            TBCQ!Operador = TBUsuarios!CODIGO & "-" & PubUsuario
        End If
        TBUsuarios.Close
        TBCQ!LOTE = txtquant.Text
        TBCQ!Data = TBProcessos!Data
        TBCQ!Hora = TBProcessos!TempoInicio
        TBCQ!Maquina = TBProcessos!Maquina
        Set TBMaquinas = BD.OpenRecordset("Select * from CadMaquinas where Maquina = '" & TBProcessos!Maquina & "'")
        If TBMaquinas.EOF = False Then
            TBCQ!Setor = TBMaquinas!Setor
        End If
        TBMaquinas.Close
        TBCQ!Turno = TBProcessos!Turno
        TBCQ!PARECERCQ = "Nada consta"
        TBCQ.Update
        TBCQ.Close
    End If
    TBProcessos.Close
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcVerificaEvento()
On Error GoTo tratar_erro

'Se a o evento de trabalho for for fim do lote
If EVENTO = 3 Then
    If Gravar = False Then Exit Sub
    'Filtra todos os eventos desta OS na tabela producaofases para marcar como fase pronta
    BD.Execute "Update producaofases Set pronto = 'SIM' where idfase = " & txtos & " and pronto = 'NÃO'"
    'Filtra ordem de servico desta Ordem na tabela ordemservico para marcar como ordem de servico concluida
    Set TBOrdemServico = BD.OpenRecordset("Select * from ordemservico where IDPRODUCAO = " & txtos.Text & "")
    If TBOrdemServico.EOF = False Then
        TBOrdemServico.Edit
        Descricao = TBOrdemServico!Descricao
        TBOrdemServico!Pronto = "SIM"
        TBOrdemServico!Dataconclusao = Date
        TBOrdemServico.Update
        TBOrdemServico.Close
    End If
    'Checa se todas as ordens de servicos com processo da ordem foram executadas e da baixa na ordem
    Set TBOrdemServico = BD.OpenRecordset("Select * from ordemservico where of = " & txtof.Text & " and pronto = 'NÃO'")
    If TBOrdemServico.EOF = True Then
        Set TBFases = BD.OpenRecordset("Select * from ordemservico where of = " & txtof.Text & " and Idprocesso <> 0")
        If TBFases.EOF = False Then
            Set TBOrdem = BD.OpenRecordset("Select * from Producao where of = " & txtof.Text & "")
            If TBOrdem.EOF = False Then
                TBOrdem.Edit
                TBOrdem!pronta = "SIM"
                TBOrdem!concluida = True
                TBOrdem!Status = "Concluída"
                TBOrdem!dataentrega = Date
                TBOrdem!prioridade = 0
                TBOrdem.Update
                'Verifica se todas as ordems de fabricação do produto já foram concluidas
                Set TBProcessos = BD.OpenRecordset("Select * from producao where idcarteira = " & TBOrdem!idcarteira & " and pronta = 'NÃO'")
                If TBProcessos.EOF = True Then
                    Set TBVendas = BD.OpenRecordset("Select * from Vendas_carteira where código = " & TBOrdem!idcarteira & "")
                    If TBVendas.EOF = False Then
                        TBVendas.Edit
                        TBVendas!saida_estoque = True
                        TBVendas!dataprodsaida = Date
                        TBVendas.Update
                    End If
                    TBVendas.Close
                End If
                TBProcessos.Close
            End If
            TBOrdem.Close
        End If
        TBFases.Close
    End If
    TBOrdemServico.Close
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub procGravar()
On Error GoTo tratar_erro

TTNC = IIf(IsNumeric(Txtttnc.Text) = True, Txtttnc, 0)
TTOK = IIf(IsNumeric(txtttok.Text) = True, txtttok, 0)
TNC = IIf(IsNumeric(txtTNC) = True, txtTNC, 0)
TOK = IIf(IsNumeric(txtTOK.Text) = True, txtTOK, 0)
EVENTO = 0
Gravar = True
'Se não foi escolhida uma descrição do código de trabalho
If txtCodigo.Text = "" Then
    MsgBox "Informe o código de trabalho antes de salvar.", vbExclamation
    txtCodigo.SetFocus
    Exit Sub
End If
'Atribui um valor a variavel descricao
EVENTO = txtCodigo.Text
'Verifica disponibilidade da Máquina
If txtcodigoDesc.Text <> "TROCA DE POSTO DE TRABALHO" Then ProcVerificaDisponMaquina
If Gravar = False Then Exit Sub
'Verifica Ultimo evento
ProcVerificaUltimo
If Gravar = False Then Exit Sub
'Gravar evento realizado
ProcGravarEvento
If Gravar = False Then Exit Sub
'Verifica o evento gravado
ProcVerificaEvento
'Atualiza dados do formulario, e da lista de eventos
procAtualizaProducao

'Dados do estoque
ExcluirAP = False
Set TBOrdem = BD.OpenRecordset("Select * from Producao where OF = " & txtof & "")
If TBOrdem.EOF = False Then
    If TBOrdem!Retirar_estoque = True And (TOK <> 0 Or TNC <> 0) Then
        Set TBCFOP = BD.OpenRecordset("Select * from ordemservico where OF = " & TBOrdem!OF & " order by fase, retrabalho desc, IDproducao")
        If TBCFOP.EOF = False Then
            'Verfica se é a primeira OS e retira o material do estoque
            TBCFOP.MoveFirst
            If TBCFOP!IDProducao = txtos Then
                ProcRetirarCancelarEstoque
            Else 'Verfica se é a última OS e entra com o produto no estoque
                TBCFOP.MoveLast
                If TBCFOP!IDProducao = txtos Then ProcEntrarCancelarEstoque
            End If
        End If
        TBCFOP.Close
    End If
End If
TBOrdem.Close

'Atualiza dados da Máquina, Operador e Turno
If Penultimo = 1 Or Penultimo = 2 Then
    ProcGravarTurno
    ProcGravarMaquina
    ProcGravarOperador
End If
'Grava o status da OS e da OF
ProcGravarStatusOSOF
'Grava tempo total do evento por máquina/operador
Dataini = Date
ProcVerificaTurno
Set TBCodigoDesc = BD.OpenRecordset("Select * from CodigoDesc where codigo = " & Penultimo & " and Controlar_Totalizacao = True")
If TBCodigoDesc.EOF = False Then
    ProcGravarTotalEventoMaq
    ProcGravarTotalEventoOpe
End If
TBCodigoDesc.Close
'Acerta cadastro na máquina
ProcAcertaCadMaquina
'Atualiza dados da ordem de fabricação ( Tempo total utilizado, custo por peça, custo total, etc..
ProcAtualizaOF
txtTOK = ""
txtTNC.Text = ""

'Libera manutenção
If txtcodigoDesc = "MÁQUINA EM MANUTENÇÃO" Or txtcodigoDesc = "MANUTENÇÃO PREVENTIVA" Or txtcodigoDesc = "MANUTENÇÃO CORRETIVA" Then
    Set TBProducao = BD.OpenRecordset("select * from manutencao where IDmaquina = '" & txtMaquina & "' and Controlada = true")
    If TBProducao.EOF = False Then
        Set TBProcessosDet = BD.OpenRecordset("select * from manutencao_data where idManutencao = " & TBProducao!código & " and status = 'Aberta' and data <= #" & Date & "# ")
        If TBProcessosDet.EOF = False Then
            TBProcessosDet.Edit
            TBProcessosDet!IDProducao = IDApontamento
            TBProcessosDet.Update
        End If
        TBProcessosDet.Close
    End If
    TBProducao.Close
End If
If txtcodigoDesc.Text = "PREPARANDO MÁQUINA" Or txtcodigoDesc.Text = "MÁQUINA EM PRODUÇÃO" Then
    Set TBProcessos = BD.OpenRecordset("select * from manutencao where IDmaquina = '" & txtMaquina & "' and Controlada = true")
    If TBProcessos.EOF = False Then
        Set TBOrdem = BD.OpenRecordset("select * from manutencao_data where idManutencao = " & TBProcessos!código & " and status = 'Aberta' and IDProducao <> 0 and data <= #" & Date & "# ")
        If TBOrdem.EOF = False Then
            Maquina = txtMaquina
            frmManutencao.Show 1
        End If
        TBOrdem.Close
    End If
    TBProcessos.Close
End If
'Verifica se o evento é troca de posto de trabalho e abre o formulário
If txtcodigoDesc.Text = "TROCA DE POSTO DE TRABALHO" Then frmcgmaqCB.Show 1

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcGravarStatusOSOF()
On Error GoTo tratar_erro

Set TBOrdem = BD.OpenRecordset("Select Status from ordemservico where idproducao = " & txtos & "")
If TBOrdem.EOF = False Then
    TBOrdem.Edit
    If Lista.ListItems.Count = 0 Or Ultimo <> 1 And Ultimo <> 2 And Ultimo <> 3 Then
        TBOrdem!Status = "Aguardando"
        GoTo Sair
    End If
    If Ultimo = 1 Then
        TBOrdem!Status = "Preparando"
        GoTo Sair
    End If
    If Ultimo = 2 Then
        TBOrdem!Status = "Produzindo"
        GoTo Sair
    End If
    If Ultimo = 3 Then
        TBOrdem!Status = "Concluída"
        GoTo Sair
    End If
Sair:
    TBOrdem.Update
End If
TBOrdem.Close

If Ultimo <> 3 Then
    Set TBOrdem = BD.OpenRecordset("Select Status from producao where of = " & txtof & "")
    If TBOrdem.EOF = False Then
        TBOrdem.Edit
        Set TBAbrir = BD.OpenRecordset("Select Status from ordemservico where OF = " & txtof & " and Status <> 'Aguardando'")
        If TBAbrir.EOF = False Then
            TBOrdem!Status = "Produzindo"
            GoTo Sair1
        End If
        Set TBAbrir = BD.OpenRecordset("Select Status from ordemservico where OF = " & txtof & " and Status <> 'Preparando' and Status <> 'Produzindo' and Status <> 'Concluída'")
        If TBAbrir.EOF = False Then
            TBOrdem!Status = "Aberta"
            GoTo Sair1
        End If
Sair1:
        TBOrdem.Update
        TBAbrir.Close
    End If
    TBOrdem.Close
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcVerificaDisponMaquina()
On Error GoTo tratar_erro

Set TBMaquinas = BD.OpenRecordset("Select * from CadMaquinas where maquina = '" & txtMaquina.Text & "' and liberada = false and OS <> " & txtos.Text & "")
If TBMaquinas.EOF = False Then
    MsgBox ("Não é permitido utilizar essa máquina, pois a mesma já está sendo utilizada na Ordem: " & TBMaquinas!ordem & " - OS: " & TBMaquinas!OS & "."), vbExclamation
    TBMaquinas.Close
    Gravar = False
    Exit Sub
End If
TBMaquinas.Close

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcGravarTotalEventoMaq()
On Error GoTo tratar_erro

If PenultimoDesc <> "" Then
    TempoTotal = "00:00:00"
    TotalDisponivel = "00:00:00"
    Set TBProducao = BD.OpenRecordset("Select * from ProducaoFases where data = CDate('" & Format(Dataini, "dd/mm/YYYY") & "') and CODIGODESC = " & Penultimo & " and idfase = " & txtos & " and maquina = '" & txtMaquina.Text & "' and turno = " & Turno & " order by Data, Tempoinicio")
    If TBProducao.EOF = False Then
        TBProducao.MoveLast
        TempoTotal = Format(TBProducao!TempoTotal, "hh:mm:ss")
        
        ProcVerificaDia
        TurnoMaq = Turno
        Do While Format(TempoTotal, "hh:mm:ss") > Format(0, "hh:mm:ss")
            Set TBAbrir = BD.OpenRecordset("Select * from CadMaqturnos where maquina = '" & txtMaquina.Text & "' and diasemana = '" & DiaSemana & "' and Turno = " & TurnoMaq & " order by diasemana,turno")
            If TBAbrir.EOF = False Then
                TotalDisponivel = TBAbrir!totalturno
                Set TBMaquinas = BD.OpenRecordset("Select * from Eventos_Total_Maq where data = CDate('" & Format(Dataini, "dd/mm/YYYY") & "') and maquina = '" & txtMaquina.Text & "' and descevento = '" & PenultimoDesc & "' and turno = " & TurnoMaq & "")
                If TBMaquinas.EOF = False Then
                    TBMaquinas.Edit
                Else
                    TBMaquinas.AddNew
                End If
                If Format(TempoTotal, "hh:mm:ss") > Format(TotalDisponivel, "hh:mm:ss") Then
                    TBMaquinas!TempoTotal = TotalDisponivel + IIf(IsNull(TBMaquinas!TempoTotal), 0, TBMaquinas!TempoTotal)
                    ElapsedTime (TotalDisponivel)
                    TBMaquinas!Tempototalseg = S + IIf(IsNull(TBMaquinas!Tempototalseg), 0, TBMaquinas!Tempototalseg)
                    TempoTotal = TempoTotal - TotalDisponivel
                Else
                    TBMaquinas!TempoTotal = TempoTotal + IIf(IsNull(TBMaquinas!TempoTotal), 0, TBMaquinas!TempoTotal)
                    ElapsedTime (TempoTotal)
                    TBMaquinas!Tempototalseg = S + IIf(IsNull(TBMaquinas!Tempototalseg), 0, TBMaquinas!Tempototalseg)
                    TempoTotal = TempoTotal - TempoTotal
                End If
            Else
                Set TBMaquinas = BD.OpenRecordset("Select * from Eventos_Total_Maq where data = CDate('" & Format(Dataini, "dd/mm/YYYY") & "') and maquina = '" & txtMaquina.Text & "' and descevento = '" & PenultimoDesc & "' and turno = 0")
                If TBMaquinas.EOF = False Then
                    TBMaquinas.Edit
                Else
                    TBMaquinas.AddNew
                End If
                TBMaquinas!TempoTotal = TempoTotal + IIf(IsNull(TBMaquinas!TempoTotal), 0, TBMaquinas!TempoTotal)
                ElapsedTime (TempoTotal)
                TBMaquinas!Tempototalseg = S + IIf(IsNull(TBMaquinas!Tempototalseg), 0, TBMaquinas!Tempototalseg)
                TempoTotal = TempoTotal - TempoTotal
                TurnoMaq = 0
            End If
            TBAbrir.Close
            TBMaquinas!Data = Dataini
            TBMaquinas!Maquina = txtMaquina.Text
            TBMaquinas!descevento = TBProducao!Descricao
            TBMaquinas!Turno = TurnoMaq
            TurnoMaq = TurnoMaq + 1
            TBMaquinas.Update
        Loop
        
        'Enquanto tempo total for maior que um dia
        If TBProducao!Dias <> 0 Then
            Contador = 0
            Do While Contador <> TBProducao!Dias
                Contador = Contador + 1
                ProcVerDispMaqDia
                If TotalDisponivel <> 0 Then
                    TurnoMaq = 1
                    Do While Format(TotalDisponivel, "hh:mm:ss") > Format(0, "hh:mm:ss")
                        Set TBAbrir = BD.OpenRecordset("Select * from CadMaqturnos where maquina = '" & txtMaquina.Text & "' and diasemana = '" & DiaSemana & "' and turno = " & TurnoMaq & "")
                        If TBAbrir.EOF = False Then
                            TempoTotal = TBAbrir!totalturno
                        End If
                        TBAbrir.Close
                        TBMaquinas.AddNew
                        TBMaquinas!Data = Dataini + Contador
                        TBMaquinas!Maquina = txtMaquina.Text
                        TBMaquinas!descevento = TBProducao!Descricao
                        TBMaquinas!TempoTotal = TempoTotal
                        ElapsedTime (TempoTotal)
                        TBMaquinas!Tempototalseg = S
                        TBMaquinas!Turno = TurnoMaq
                        TotalDisponivel = TotalDisponivel - TempoTotal
                        TBMaquinas.Update
                        TurnoMaq = TurnoMaq + 1
                    Loop
                End If
            Loop
        End If
    End If
    TBProducao.Close
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcExcluirTotalEventoMaq()
On Error GoTo tratar_erro

If PenultimoDesc <> "" Then
    Set TBLista = BD.OpenRecordset("Select * from ProducaoFases where data = CDate('" & Format(Dataini, "dd/mm/YYYY") & "') and CODIGODESC = " & Penultimo & " and idfase = " & txtos & " and maquina = '" & txtMaquina.Text & "' and turno = " & Turno & " order by Data, Tempoinicio")
    If TBLista.EOF = False Then
        TBLista.MoveLast
        TempoTotal = IIf(IsNull(TBLista!TempoTotal), 0, Format(TBLista!TempoTotal, "hh:mm:ss"))
        
        ProcVerificaDia
        TurnoMaq = Turno
        Do While Format(TempoTotal, "hh:mm:ss") > Format(0, "hh:mm:ss")
            Set TBAbrir = BD.OpenRecordset("Select * from CadMaqturnos where maquina = '" & txtMaquina.Text & "' and diasemana = '" & DiaSemana & "' and Turno = " & TurnoMaq & " order by diasemana,turno")
            If TBAbrir.EOF = False Then
                TotalDisponivel = TBAbrir!totalturno
                Set TBMaquinas = BD.OpenRecordset("Select * from Eventos_Total_Maq where data = CDate('" & Format(Dataini, "dd/mm/YYYY") & "') and maquina = '" & txtMaquina.Text & "' and descevento = '" & PenultimoDesc & "' and turno = " & TurnoMaq & "")
                If TBMaquinas.EOF = False Then
                    TBMaquinas.Edit
                    If Format(TempoTotal, "hh:mm:ss") > Format(TotalDisponivel, "hh:mm:ss") Then
                        TBMaquinas!TempoTotal = TBMaquinas!TempoTotal - TotalDisponivel
                        ElapsedTime (TotalDisponivel)
                        TBMaquinas!Tempototalseg = TBMaquinas!Tempototalseg - S
                        TempoTotal = TempoTotal - TotalDisponivel
                    Else
                        TBMaquinas!TempoTotal = TBMaquinas!TempoTotal - TempoTotal
                        ElapsedTime (TempoTotal)
                        TBMaquinas!Tempototalseg = TBMaquinas!Tempototalseg - S
                        TempoTotal = TempoTotal - TempoTotal
                    End If
                    TBMaquinas.Update
                    If TBMaquinas!TempoTotal = "00:00:00" Then BD.Execute "DELETE * from Eventos_Total_Maq where id = " & TBMaquinas!Id & ""
                End If
                TBMaquinas.Close
            Else
                Set TBMaquinas = BD.OpenRecordset("Select * from Eventos_Total_Maq where data = CDate('" & Format(Dataini, "dd/mm/YYYY") & "') and maquina = '" & txtMaquina.Text & "' and descevento = '" & PenultimoDesc & "' and turno = 0")
                If TBMaquinas.EOF = False Then
                    TBMaquinas.Edit
                    TBMaquinas!TempoTotal = TBMaquinas!TempoTotal - TempoTotal
                    ElapsedTime (TempoTotal)
                    TBMaquinas!Tempototalseg = TBMaquinas!Tempototalseg - S
                    TempoTotal = TempoTotal - TempoTotal
                    TBMaquinas.Update
                    If TBMaquinas!TempoTotal = "00:00:00" Then BD.Execute "DELETE * from Eventos_Total_Maq where id = " & TBMaquinas!Id & ""
                Else
                    TempoTotal = TempoTotal - TempoTotal
                End If
                TBMaquinas.Close
            End If
            TBAbrir.Close
            TurnoMaq = TurnoMaq + 1
        Loop
        
        'Enquanto tempo total for maior que um dia
        If Dias <> 0 Then
            Contador = 0
            Do While Contador <> Dias
                Contador = Contador + 1
                BD.Execute "DELETE * from Eventos_Total_Maq where data = CDate('" & Format(Dataini + Contador, "dd/mm/YYYY") & "') and maquina = '" & txtMaquina.Text & "' and descevento = '" & PenultimoDesc & "'"
            Loop
        End If
    End If
    TBLista.Close
End If
    
Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcVerDispMaqDia()
On Error GoTo tratar_erro

Dataini = Dataini + Contador
TotalDisponivel = 0
ProcVerificaDia
Set TBAbrir = BD.OpenRecordset("Select * from CadMaqturnos where maquina = '" & txtMaquina.Text & "' and diasemana = '" & DiaSemana & "' order by diasemana,turno")
If TBAbrir.EOF = False Then
    Do While TBAbrir.EOF = False
        TotalDisponivel = TotalDisponivel + TBAbrir!totalturno
        TBAbrir.MoveNext
    Loop
End If
TBAbrir.Close
Dataini = Dataini - Contador

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcGravarTotalEventoOpe()
On Error GoTo tratar_erro

If PenultimoDesc <> "" Then
    TempoTotal = "00:00:00"
    TotalDisponivel = "00:00:00"
    Set TBProducao = BD.OpenRecordset("Select * from ProducaoFases where data = CDate('" & Format(Dataini, "dd/mm/YYYY") & "') and CODIGODESC = " & Penultimo & " and idfase = " & txtos & " and turno = " & Turno & " order by Data, Tempoinicio")
    If TBProducao.EOF = False Then
        TBProducao.MoveLast
        TempoTotal = Format(TBProducao!TempoTotal, "hh:mm:ss")
        
        ProcVerificaDia
        TurnoMaq = Turno
        Do While Format(TempoTotal, "hh:mm:ss") > Format(0, "hh:mm:ss")
            Set TBAbrir = BD.OpenRecordset("Select * from CadMaqturnos where maquina = '" & txtMaquina.Text & "' and diasemana = '" & DiaSemana & "' and Turno = " & TurnoMaq & " order by diasemana,turno")
            If TBAbrir.EOF = False Then
                TotalDisponivel = TBAbrir!totalturno
                Set TBUsuarios = BD.OpenRecordset("Select * from Eventos_Total_Ope where data = CDate('" & Format(Dataini, "dd/mm/YYYY") & "') and operador = '" & Operador & "' and descevento = '" & PenultimoDesc & "' and turno = " & TurnoMaq & "")
                If TBUsuarios.EOF = False Then
                    TBUsuarios.Edit
                Else
                    TBUsuarios.AddNew
                End If
                If Format(TempoTotal, "hh:mm:ss") > Format(TotalDisponivel, "hh:mm:ss") Then
                    TBUsuarios!TempoTotal = TotalDisponivel + IIf(IsNull(TBUsuarios!TempoTotal), 0, TBUsuarios!TempoTotal)
                    ElapsedTime (TotalDisponivel)
                    TBUsuarios!Tempototalseg = S + IIf(IsNull(TBUsuarios!Tempototalseg), 0, TBUsuarios!Tempototalseg)
                    TempoTotal = TempoTotal - TotalDisponivel
                Else
                    TBUsuarios!TempoTotal = TempoTotal + IIf(IsNull(TBUsuarios!TempoTotal), 0, TBUsuarios!TempoTotal)
                    ElapsedTime (TempoTotal)
                    TBUsuarios!Tempototalseg = S + IIf(IsNull(TBUsuarios!Tempototalseg), 0, TBUsuarios!Tempototalseg)
                    TempoTotal = TempoTotal - TempoTotal
                End If
            Else
                Set TBUsuarios = BD.OpenRecordset("Select * from Eventos_Total_Ope where data = CDate('" & Format(Dataini, "dd/mm/YYYY") & "') and operador = '" & Operador & "' and descevento = '" & PenultimoDesc & "' and turno = 0")
                If TBUsuarios.EOF = False Then
                    TBUsuarios.Edit
                Else
                    TBUsuarios.AddNew
                End If
                TBUsuarios!TempoTotal = TempoTotal + IIf(IsNull(TBUsuarios!TempoTotal), 0, TBUsuarios!TempoTotal)
                ElapsedTime (TempoTotal)
                TBUsuarios!Tempototalseg = S + IIf(IsNull(TBUsuarios!Tempototalseg), 0, TBUsuarios!Tempototalseg)
                TempoTotal = TempoTotal - TempoTotal
                TurnoMaq = 0
            End If
            TBAbrir.Close
            TBUsuarios!Data = Dataini
            TBUsuarios!Operador = Operador
            TBUsuarios!descevento = TBProducao!Descricao
            TBUsuarios!Turno = TurnoMaq
            TurnoMaq = TurnoMaq + 1
            TBUsuarios.Update
        Loop
        
        'Enquanto tempo total for maior que um dia
        If TBProducao!Dias <> 0 Then
            Contador = 0
            Do While Contador <> TBProducao!Dias
                Contador = Contador + 1
                ProcVerDispMaqDia
                If TotalDisponivel <> 0 Then
                    TurnoMaq = 1
                    Do While Format(TotalDisponivel, "hh:mm:ss") > Format(0, "hh:mm:ss")
                        Set TBAbrir = BD.OpenRecordset("Select * from CadMaqturnos where maquina = '" & txtMaquina.Text & "' and diasemana = '" & DiaSemana & "' and turno = " & TurnoMaq & "")
                        If TBAbrir.EOF = False Then
                            TempoTotal = TBAbrir!totalturno
                        End If
                        TBAbrir.Close
                        TBUsuarios.AddNew
                        TBUsuarios!Data = Dataini + Contador
                        TBUsuarios!Operador = Operador
                        TBUsuarios!descevento = TBProducao!Descricao
                        TBUsuarios!TempoTotal = TempoTotal
                        ElapsedTime (TempoTotal)
                        TBUsuarios!Tempototalseg = S
                        TBUsuarios!Turno = TurnoMaq
                        TotalDisponivel = TotalDisponivel - TempoTotal
                        TBUsuarios.Update
                        TurnoMaq = TurnoMaq + 1
                    Loop
                End If
            Loop
        End If
    End If
    TBProducao.Close
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcExcluirTotalEventoOpe()
On Error GoTo tratar_erro

If PenultimoDesc <> "" Then
    Set TBLista = BD.OpenRecordset("Select * from ProducaoFases where data = CDate('" & Format(Dataini, "dd/mm/YYYY") & "') and CODIGODESC = " & Penultimo & " and idfase = " & txtos & " and turno = " & Turno & " order by Data, Tempoinicio")
    If TBLista.EOF = False Then
        TBLista.MoveLast
        TempoTotal = IIf(IsNull(TBLista!TempoTotal), 0, Format(TBLista!TempoTotal, "hh:mm:ss"))
        
        ProcVerificaDia
        TurnoMaq = Turno
        Do While Format(TempoTotal, "hh:mm:ss") > Format(0, "hh:mm:ss")
            Set TBAbrir = BD.OpenRecordset("Select * from CadMaqturnos where maquina = '" & txtMaquina.Text & "' and diasemana = '" & DiaSemana & "' and Turno = " & TurnoMaq & " order by diasemana,turno")
            If TBAbrir.EOF = False Then
                TotalDisponivel = TBAbrir!totalturno
                Set TBUsuarios = BD.OpenRecordset("Select * from Eventos_Total_Ope where data = CDate('" & Format(Dataini, "dd/mm/YYYY") & "') and operador = '" & Operador & "' and descevento = '" & PenultimoDesc & "' and turno = " & TurnoMaq & "")
                If TBUsuarios.EOF = False Then
                    TBUsuarios.Edit
                    If Format(TempoTotal, "hh:mm:ss") > Format(TotalDisponivel, "hh:mm:ss") Then
                        TBUsuarios!TempoTotal = TBUsuarios!TempoTotal - TotalDisponivel
                        ElapsedTime (TotalDisponivel)
                        TBUsuarios!Tempototalseg = TBUsuarios!Tempototalseg - S
                        TempoTotal = TempoTotal - TotalDisponivel
                    Else
                        TBUsuarios!TempoTotal = TBUsuarios!TempoTotal - TempoTotal
                        ElapsedTime (TempoTotal)
                        TBUsuarios!Tempototalseg = TBUsuarios!Tempototalseg - S
                        TempoTotal = TempoTotal - TempoTotal
                    End If
                    If TBUsuarios!TempoTotal = "00:00:00" Then BD.Execute "DELETE * from Eventos_Total_Ope where id = " & TBUsuarios!Id & ""
                    TBUsuarios.Update
                End If
                TBUsuarios.Close
            Else
                Set TBUsuarios = BD.OpenRecordset("Select * from Eventos_Total_Ope where data = CDate('" & Format(Dataini, "dd/mm/YYYY") & "') and operador = '" & Operador & "' and descevento = '" & PenultimoDesc & "' and turno = 0")
                If TBUsuarios.EOF = False Then
                    TBUsuarios.Edit
                    TBUsuarios!TempoTotal = TBUsuarios!TempoTotal - TempoTotal
                    ElapsedTime (TempoTotal)
                    TBUsuarios!Tempototalseg = TBUsuarios!Tempototalseg - S
                    TempoTotal = TempoTotal - TempoTotal
                    If TBUsuarios!TempoTotal = "00:00:00" Then BD.Execute "DELETE * from Eventos_Total_Ope where id = " & TBUsuarios!Id & ""
                    TBUsuarios.Update
                Else
                    TempoTotal = TempoTotal - TempoTotal
                End If
                TBUsuarios.Close
            End If
            TBAbrir.Close
            TurnoMaq = TurnoMaq + 1
        Loop
        
        'Enquanto tempo total for maior que um dia
        If Dias <> 0 Then
            Contador = 0
            Do While Contador <> Dias
                Contador = Contador + 1
                BD.Execute "DELETE * from Eventos_Total_Ope where data = CDate('" & Format(Dataini + Contador, "dd/mm/YYYY") & "') and operador = '" & Operador & "' and descevento = '" & PenultimoDesc & "'"
            Loop
        End If
    End If
    TBLista.Close
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcGravarMaquina()
On Error GoTo tratar_erro

ProcAtualizaProducaoMaquina
Set TBOrdemServico = BD.OpenRecordset("Select * from ordemservico where idproducao = " & txtos.Text & "")
If TBOrdemServico.EOF = False Then
    Set TBProducao = BD.OpenRecordset("Select * from ProducaoFases_Maquina where OS = " & txtos.Text & " and maquina = '" & txtMaquina.Text & "' and data = CDate('" & Format(Dataini, "dd/mm/yyyy") & "')")
    If TBProducao.EOF = False Then
        TBProducao.Edit
    Else
        TBProducao.AddNew
    End If
    TBProducao!OF = txtof.Text
    TBProducao!OS = txtos.Text
    TBProducao!Fase = TBOrdemServico!Fase
    TBProducao!Maquina = txtMaquina.Text
    TBProducao!Preparacao = TBOrdemServico!TempoPreparacao 'Tempo total previsto de preparação
    TBProducao!Execucao = TBOrdemServico!TempoExecucao 'Tempo total previsto de execução
    TBProducao!QTNC = TTNC
    TBProducao!QTOK = TTOK
    TBProducao!TPUTIL = TempoTotalPrep 'Tempo total real de preparação do lote
    TBProducao!TEUTIL = FormataTempo(TEUSEG) 'Tempo total real de execução por peça
    ElapsedTime (TempoTotalUtil)
    TBProducao!TETTUTIL = Horatotal 'Tempo total real de preparação + tempo total real de execução do lote
    TBProducao!CRLOTE = CTTLOTE
    TBProducao!CRPECA = CTTPECA
    TBProducao!CPLOTE = TBOrdemServico!CPLOTE
    TBProducao!CPPECA = TBOrdemServico!CPPECA
    TBProducao!Eficiencia = Eficiencia
    TBProducao!Totalprod = Produzidas
    TBProducao!Data = Dataini
    TBProducao.Update
End If
TBProducao.Close
BD.Execute "Update ProducaoFases_Maquina Set pronto = '" & TBOrdemServico!Pronto & "' where OS = " & txtos.Text & ""
TBOrdemServico.Close

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Sub ProcAtualizaProducaoMaquina()
On Error GoTo tratar_erro

TempoTotalPrep = "00:00:00"
TempoTotalProd = "00:00:00"
TTOK = 0
TTNC = 0
Produzidas = 0
'Filtra dados da tabela produçãofases por OS, maquina
Set TBProducao = BD.OpenRecordset("Select * from producaofases where idfase = " & txtos & " and maquina = '" & txtMaquina.Text & "' and data = CDate('" & Format(Dataini, "dd/mm/yyyy") & "') and Turno = " & Turno & " order by Data, Tempoinicio")
If TBProducao.EOF = False Then
    'Conta peças produzidas ok e nc
    Do While TBProducao.EOF = False
        TTOK = TTOK + IIf(IsNull(TBProducao!Quantidade) = False, TBProducao!Quantidade, 0)
        TTNC = TTNC + IIf(IsNull(TBProducao!reprovada) = False, TBProducao!reprovada, 0)
        EVENTO = TBProducao!CodigoDesc
        Select Case EVENTO
            Case 1:
                TempoTotalPrep = TempoTotalPrep + IIf(IsNull(TBProducao!TempoTotal) = False, TBProducao!TempoTotal, 0)
            Case 2:
                TempoTotalProd = TempoTotalProd + IIf(IsNull(TBProducao!TempoTotal) = False, TBProducao!TempoTotal, 0)
        End Select
        TBProducao.MoveNext
    Loop
End If
TBProducao.Close
'Atualiza o tempo de preparação utilizado
ProcAtualizaPrepUtilMaquina
'Atualiza o tempo de execução utilizado
ProcAtualizaExecUtilMaquina

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcAtualizaPrepUtilMaquina()
On Error GoTo tratar_erro

'Calcula total de segundos utilizados preparando
ElapsedTime (TempoTotalPrep)
TPUSEG = S
TTPUTIL = Horatotal

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcAtualizaExecUtilMaquina()
On Error GoTo tratar_erro

Set TBOrdemServico = BD.OpenRecordset("Select * from ordemservico where IDPRODUCAO = " & txtos & "")
If TBOrdemServico.EOF = False Then
    'Calcula tempo total de execução
    ElapsedTime (TempoTotalProd)
    TTEUTIL = Horatotal
    TTEUTILS = S
    
    'Soma tempo total de preparação + execução
    TempoTotalUtil = TempoTotalPrep + TempoTotalProd
    ElapsedTime (TempoTotalUtil)
    TTUTILSEG = S
   
    'Calcula tempo total de execução por peça
    Produzidas = TTOK + TTNC
    If TempoTotalProd > 0 And Produzidas > 0 Then TEUSEG = TTEUTILS / Produzidas Else TEUSEG = TTEUTILS
            
    'Calcula eficiencia
    TEPSEG = IIf(IsNull(TBOrdemServico!TESegundos), 0, TBOrdemServico!TESegundos) 'Verif. tempo de execução previsto por peça
    If TEPSEG > 0 And TEUSEG > 0 And Produzidas > 0 Then Eficiencia = Format(TEPSEG / TEUSEG * 100, "###,##0.00") Else Eficiencia = 0
    
    'Calcula e grava valor real do lote e por peça
    Set TBMaquinas = BD.OpenRecordset("Select * from cadmaquinas where maquina = '" & txtMaquina.Text & "'")
    If TBMaquinas.EOF = False Then
        Valorhora = TBMaquinas!PrecoHora
        Valorhora = Valorhora / 3600
        CTTLOTE = Format((Valorhora * TEUSEG * (TTOK + TTNC)) + (Valorhora * TPUSEG), "###,##0.00")
        If TEUSEG > 0 And (TTOK + TTNC) > 0 Then CTTPECA = Format(CTTLOTE / (TTOK + TTNC), "###,##0.00000")
    Else
        Valorhora = 0
    End If
    TBMaquinas.Close
End If
TBOrdemServico.Close

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcGravarOperador()
On Error GoTo tratar_erro

ProcAtualizaProducaoOperador
Set TBOrdemServico = BD.OpenRecordset("Select * from ordemservico where idproducao = " & txtos.Text & "")
If TBOrdemServico.EOF = False Then
    Set TBProducao = BD.OpenRecordset("Select * from ProducaoFases_Operador where OS = " & txtos.Text & " and usuario = '" & Operador & "' and data = CDate('" & Format(Dataini, "dd/mm/yyyy") & "')")
    If TBProducao.EOF = False Then
        TBProducao.Edit
    Else
        TBProducao.AddNew
    End If
    TBProducao!OF = txtof.Text
    TBProducao!OS = txtos.Text
    TBProducao!Fase = TBOrdemServico!Fase
    TBProducao!usuario = Operador
    TBProducao!Preparacao = TBOrdemServico!TempoPreparacao 'Tempo total previsto de preparação
    TBProducao!Execucao = TBOrdemServico!TempoExecucao 'Tempo total previsto de execução
    TBProducao!QTNC = TTNC
    TBProducao!QTOK = TTOK
    TBProducao!TPUTIL = TempoTotalPrep 'Tempo total real de preparação do lote
    TBProducao!TEUTIL = FormataTempo(TEUSEG) 'Tempo total real de execução por peça
    ElapsedTime (TempoTotalUtil)
    TBProducao!TETTUTIL = Horatotal 'Tempo total real de preparação + tempo total real de execução do lote
    TBProducao!CRLOTE = CTTLOTE
    TBProducao!CRPECA = CTTPECA
    TBProducao!CPLOTE = TBOrdemServico!CPLOTE
    TBProducao!CPPECA = TBOrdemServico!CPPECA
    TBProducao!Eficiencia = Eficiencia
    TBProducao!Totalprod = Produzidas
    TBProducao!Data = Dataini
    TBProducao.Update
End If
TBProducao.Close
BD.Execute "Update ProducaoFases_Operador Set pronto = '" & TBOrdemServico!Pronto & "' where OS = " & txtos.Text & ""
TBOrdemServico.Close

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Sub ProcAtualizaProducaoOperador()
On Error GoTo tratar_erro

TempoTotalPrep = "00:00:00"
TempoTotalProd = "00:00:00"
TTOK = 0
TTNC = 0
Produzidas = 0
'Filtra dados da tabela produçãofases por OS, operador
Set TBProducao = BD.OpenRecordset("Select * from producaofases where idfase = " & txtos & " and usuario = '" & Operador & "' and data = CDate('" & Format(Dataini, "dd/mm/yyyy") & "') and Turno = " & Turno & " order by Data, Tempoinicio")
If TBProducao.EOF = False Then
    'Conta peças produzidas ok e nc
    Do While TBProducao.EOF = False
        TTOK = TTOK + IIf(IsNull(TBProducao!Quantidade) = False, TBProducao!Quantidade, 0)
        TTNC = TTNC + IIf(IsNull(TBProducao!reprovada) = False, TBProducao!reprovada, 0)
        EVENTO = TBProducao!CodigoDesc
        Select Case EVENTO
            Case 1:
                TempoTotalPrep = TempoTotalPrep + IIf(IsNull(TBProducao!TempoTotal) = False, TBProducao!TempoTotal, 0)
            Case 2:
                TempoTotalProd = TempoTotalProd + IIf(IsNull(TBProducao!TempoTotal) = False, TBProducao!TempoTotal, 0)
        End Select
        TBProducao.MoveNext
    Loop
End If
TBProducao.Close
'Atualiza o tempo de preparação utilizado
ProcAtualizaPrepUtilMaquina
'Atualiza o tempo de execução utilizado
ProcAtualizaExecUtilMaquina

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcGravarTurno()
On Error GoTo tratar_erro

ProcAtualizaProducaoTurno
Set TBOrdemServico = BD.OpenRecordset("Select * from ordemservico where idproducao = " & txtos.Text & "")
If TBOrdemServico.EOF = False Then
    Set TBProducao = BD.OpenRecordset("Select * from ProducaoFases_Turno where OS = " & txtos.Text & " and Turno = " & Turno & " and data = CDate('" & Format(Dataini, "dd/mm/yyyy") & "')")
    If TBProducao.EOF = False Then
        TBProducao.Edit
    Else
        TBProducao.AddNew
    End If
    TBProducao!OF = txtof.Text
    TBProducao!OS = txtos.Text
    TBProducao!Fase = TBOrdemServico!Fase
    TBProducao!Turno = Turno
    TBProducao!Preparacao = TBOrdemServico!TempoPreparacao 'Tempo total previsto de preparação
    TBProducao!Execucao = TBOrdemServico!TempoExecucao 'Tempo total previsto de execução
    TBProducao!QTNC = TTNC
    TBProducao!QTOK = TTOK
    TBProducao!TPUTIL = TempoTotalPrep 'Tempo total real de preparação do lote
    TBProducao!TEUTIL = FormataTempo(TEUSEG) 'Tempo total real de execução por peça
    ElapsedTime (TempoTotalUtil)
    TBProducao!TETTUTIL = Horatotal 'Tempo total real de preparação + tempo total real de execução do lote
    TBProducao!CRLOTE = CTTLOTE
    TBProducao!CRPECA = CTTPECA
    TBProducao!CPLOTE = TBOrdemServico!CPLOTE
    TBProducao!CPPECA = TBOrdemServico!CPPECA
    TBProducao!Eficiencia = Eficiencia
    TBProducao!Totalprod = Produzidas
    TBProducao!Data = Dataini
    TBProducao.Update
End If
TBProducao.Close
BD.Execute "Update ProducaoFases_Turno Set pronto = '" & TBOrdemServico!Pronto & "' where OS = " & txtos.Text & ""
TBOrdemServico.Close

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Sub ProcAtualizaProducaoTurno()
On Error GoTo tratar_erro

TempoTotalPrep = "00:00:00"
TempoTotalProd = "00:00:00"
TTOK = 0
TTNC = 0
Produzidas = 0
'Filtra dados da tabela produçãofases por OS, turno
Set TBProducao = BD.OpenRecordset("Select * from producaofases where idfase = " & txtos & " and turno = " & Turno & " and data = CDate('" & Format(Dataini, "dd/mm/yyyy") & "') order by Data, Tempoinicio")
If TBProducao.EOF = False Then
    'Conta peças produzidas ok e nc
    Do While TBProducao.EOF = False
        TTOK = TTOK + IIf(IsNull(TBProducao!Quantidade) = False, TBProducao!Quantidade, 0)
        TTNC = TTNC + IIf(IsNull(TBProducao!reprovada) = False, TBProducao!reprovada, 0)
        EVENTO = TBProducao!CodigoDesc
        Select Case EVENTO
            Case 1:
                TempoTotalPrep = TempoTotalPrep + IIf(IsNull(TBProducao!TempoTotal) = False, TBProducao!TempoTotal, 0)
            Case 2:
                TempoTotalProd = TempoTotalProd + IIf(IsNull(TBProducao!TempoTotal) = False, TBProducao!TempoTotal, 0)
        End Select
        TBProducao.MoveNext
    Loop
End If
TBProducao.Close
'Atualiza o tempo de preparação utilizado
ProcAtualizaPrepUtilMaquina
'Atualiza o tempo de execução utilizado
ProcAtualizaExecUtilMaquina

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcAtualizaOrdemServico()
On Error GoTo tratar_erro
    
Set TBOrdemServico = BD.OpenRecordset("Select * from ordemservico where IDPRODUCAO = " & txtos.Text & "")
If TBOrdemServico.EOF = False Then
    TBOrdemServico.Edit
    If IsNull(TBOrdemServico!QTOK) = True Then
        TBOrdemServico!QTOK = 0
    End If
    If IsNull(TBOrdemServico!QTNC) = True Then
        TBOrdemServico!QTNC = 0
    End If
    TBOrdemServico!QTOK = TBOrdemServico!QTOK + txtttok
    TBOrdemServico!QTNC = TBOrdemServico!QTNC + Txtttnc
    TBOrdemServico!Eficiencia = IIf(IsNumeric(txteficiencia.Text) = True, txteficiencia, "0")
    TBOrdemServico!Totalprod = IIf(IsNumeric(txtProduzida.Text) = True, txtProduzida, "0")
    TBOrdemServico.Update
End If
TBOrdemServico.Close

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcAtualizaOF()
On Error GoTo tratar_erro
Dim CRLOTE As Double, CRPECA As Double, TRLOTE As Date, TRPECA As Date, TOTALLOTE As Date

CRLOTE = 0
CRPECA = 0
TRLOTE = 0
TRPECA = 0
TOTALLOTE = 0
TEUSEG = 0
QtdeSaida = 0
TotalNC = 0
Set TBProducao = BD.OpenRecordset("Select * from ordemservico where of = " & txtof.Text & " ORDER BY fase, retrabalho desc, IDproducao")
If TBProducao.EOF = False Then
    Do While TBProducao.EOF = False
        If TBProducao!custos = True Then
            CRLOTE = CRLOTE + IIf(IsNull(TBProducao!CRLOTE) = False, TBProducao!CRLOTE, 0)
            CRPECA = CRPECA + IIf(IsNull(TBProducao!CRPECA) = False, TBProducao!CRPECA, 0)
            TOTALLOTE = IIf(IsNull(TBProducao!TETTUTILN) = False, TBProducao!TETTUTILN, 0)
            TRLOTE = TRLOTE + TOTALLOTE
            TEUSEG = TEUSEG + IIf(IsNull(TBProducao!TEUSEG), 0, TBProducao!TEUSEG)
        End If
        Totalprod = TBProducao!QTOK
        Set TBAbrir = BD.OpenRecordset("Select * from CQ_NC_FABRICA where OS = " & TBProducao!IDProducao & " and PARECERCQ = 'Rejeitar'")
        If TBAbrir.EOF = False Then
            TotalNC = TotalNC + TBProducao!QTNC
        Else
            QtdeSaida = TBProducao!QTNC
        End If
        TBAbrir.Close
        TBProducao.MoveNext
    Loop
End If
Totalprod = Totalprod + QtdeSaida
ElapsedTime (TRLOTE)
TOTALDIAREAL = Horatotal
Set TBProducao = BD.OpenRecordset("Select * from producao where of = " & txtof.Text & "")
If TBProducao.EOF = False Then
    TBProducao.Edit
    TBProducao!QuantProd = Totalprod
    TBProducao!QuantNC = TotalNC
    TBProducao.Update
    TBProducao.Edit
    TBProducao!CPR = Format(CRPECA, "###,##0.00000")
    TBProducao!CTTReal = Format(CRLOTE, "###,##0.00")
    
    S = TEUSEG
    TBProducao!tpr = FormataTempo(S)
    
    TBProducao!TTTReal = TOTALDIAREAL
    ProcCalculaEficienciaOF
    TBProducao!Eficiencia = Eficiencia
    TBProducao.Update
    'Se for ordem de expedição grava data e qtde expedida na carteira
    If TBProducao!Tipo = "E" Then
        QuantExped = 0
        Set TBAbrir = BD.OpenRecordset("Select QuantProd from producao where idcarteira = " & TBProducao!idcarteira & " and Tipo = 'E' and Status <> 'Cancelada'")
        If TBAbrir.EOF = False Then
            Do While TBAbrir.EOF = False
                QuantExped = QuantExped + IIf(IsNull(TBAbrir!QuantProd), 0, TBAbrir!QuantProd)
                TBAbrir.MoveNext
            Loop
        End If
        TBAbrir.Close
        If QuantExped > 0 Then
            BD.Execute "Update vendas_carteira Set Dataexpedicao = '" & Date & "', Qtdeexpedida = " & QuantExped & " where código = " & TBProducao!idcarteira & ""
        Else
            BD.Execute "Update vendas_carteira Set Dataexpedicao = Null, Qtdeexpedida = " & QuantExped & " where código = " & TBProducao!idcarteira & ""
        End If
    End If
End If
TBProducao.Close

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error GoTo tratar_erro
    
'Gravar inclusão de evento
If KeyCode = vbKeyF3 Then procGravar
'Excluir evento da lista
If KeyCode = vbKeyF4 Then procExcluir
If KeyCode = vbKeyF11 Then procAtualizaProducao
If KeyCode = vbKeyF12 Then ProcLista12Ultimos
'Fim de apontamento de eventos
If KeyCode = vbKeyF6 Then
    Unload Me
    txtTOK.Visible = False
    txtTOK = ""
    txtTNC.Text = ""
    txtTNC.Visible = False
    Label16.Visible = False
    frmfundo.Show
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcAltMaquina()
On Error GoTo tratar_erro
    
If txtos.Text = "" Then
    MsgBox ("Informa a ordem antes de modificar a máquina."), vbExclamation
    Exit Sub
End If
Set TBMaquinas = BD.OpenRecordset("Select descevento from cadmaquinas where os = " & txtos.Text & " and maquina = '" & txtMaquina.Text & "'")
If TBMaquinas.EOF = False Then
    If TBMaquinas!descevento <> "TROCA DE POSTO DE TRABALHO" Then
        MsgBox ("É necessário adicionar na lista o evento TROCA DE POSTO DE TRABALHO antes de alterar."), vbExclamation
        Exit Sub
    End If
End If
TBMaquinas.Close
OF = txtof
OS = txtos
Preparacao = txtpreparacao
Execucao = txtexecucao
IDFase = txtFase
frmcgmaqCB.txtMaquina = txtMaquina
frmcgmaqCB.txtpreparacao = txtpreparacao
frmcgmaqCB.txtexecucao = txtexecucao
frmcgmaqCB.txtprepnovo = txtpreparacao
frmcgmaqCB.txtexecnovo = txtexecucao
frmcgmaqCB.Show 1

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcEncerra()
On Error GoTo tratar_erro

If MsgBox("Deseja realmente encerrar o Gerprod?", vbQuestion + vbYesNo) = vbYes Then End

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Form_Load()
On Error GoTo tratar_erro

Caption = "Gerprod  - Coletor de dados no chão de fábrica - Código de barras - Empresa : " & Empresa & ""
PBLista.Value = 100

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Form_Resize()
On Error GoTo tratar_erro

frmProducao_CB.Refresh
txtCodigo.Enabled = True
txtCodigo.Text = ""
txtCodigo.Locked = False
txtCodigo.TabStop = True
txtCodigo.SetFocus

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Lista_GotFocus()
On Error GoTo tratar_erro

ProcPuxaDados

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
    EVENTO = Lista.SelectedItem.ListSubItems.Item(1).Text
    txtCodigo.Text = Lista.SelectedItem.ListSubItems.Item(1).Text
    txtdescricao.Text = Lista.SelectedItem.ListSubItems(2)
    Set TBProducao = BD.OpenRecordset("Select * from producaofases where idproducao= " & Lista.SelectedItem & "")
    If TBProducao.EOF = False Then
        txtturno.Text = TBProducao!Turno
    End If
    TBProducao.Close
    ExcluiSel = True
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Sub procAtualizaProducao()
On Error GoTo tratar_erro
Dim TotalGanho
Dim TotalPerdido
Dim Anterior
Dim TotalQuant As Long
Dim TempoTotalDias As Date

txtstatus.Text = ""
'Rotina de atualização de dados
IDProducao = 0
'Limpa dados da lista
Lista.ListItems.Clear
TempoTotalPrep = "00:00:00"
TempoTotalProd = "00:00:00"
TTOK = 0
TTNC = 0
Produzidas = 0
TempoUtilizadoDescricao = 0
'Filtra dados da tabela produçãofases por OS
Set TBProducao = BD.OpenRecordset("Select * from producaofases where idfase = " & txtos & " order by Data, Tempoinicio")
If Not (TBProducao.BOF And TBProducao.EOF) Then
    'Se tiver mais de um evento gravado
    If TBProducao.RecordCount > 1 Then
        TBProducao.MoveLast
        'Grava evento na variavel ultimo
        Ultimo = TBProducao!CodigoDesc
        UltimoDesc = TBProducao!Descricao
        TempoUltimo = TBProducao!TempoInicio
        If TBProducao!CodigoDesc = 2 Then
            TBProducao.Edit
            TBProducao!Quantidade = 0
            TBProducao!reprovada = 0
            TBProducao.Update
        End If
        TBProducao.MovePrevious
        'Grava evento na variavel penultimo
        Penultimo = TBProducao!CodigoDesc
        PenultimoDesc = TBProducao!Descricao
        'Informa turno
        If IsNull(TBProducao!Turno) = False Then Turno = TBProducao!Turno
        txtturno = Turno
        'Informa data do penultimo evento
        Dataini = TBProducao!Data
        TBProducao.MoveFirst
    Else
        'Grava evento na variavel ultimo
        Ultimo = TBProducao!CodigoDesc
        UltimoDesc = TBProducao!Descricao
        TempoUltimo = TBProducao!TempoInicio
    End If
Else
    Ultimo = 0
    UltimoDesc = ""
    TempoUltimo = 0
    Penultimo = 0
    PenultimoDesc = ""
    totalok = 0
    TotalNC = 0
End If
Produzidas = 0
'Conta peças produzidas ok e nc
Do While TBProducao.EOF = False
    TTOK = TTOK + IIf(IsNull(TBProducao!Quantidade) = False, TBProducao!Quantidade, 0)
    TTNC = TTNC + IIf(IsNull(TBProducao!reprovada) = False, TBProducao!reprovada, 0)
    EVENTO = TBProducao!CodigoDesc
    Select Case EVENTO
        Case 1:
            TempoTotalPrep = TempoTotalPrep + IIf(IsNull(TBProducao!TempoTotal) = False, TBProducao!TempoTotal, 0)
        Case 2:
            TempoTotalProd = TempoTotalProd + IIf(IsNull(TBProducao!TempoTotal) = False, TBProducao!TempoTotal, 0)
    End Select
    TBProducao.MoveNext
Loop
'Carrega valores nas caixas de texto
Set TBOrdemServico = BD.OpenRecordset("Select * from CQ_NC_FABRICA where OS = " & txtos & " and idproducao = 0")
If TBOrdemServico.EOF = False Then
    Do While TBOrdemServico.EOF = False
        TTNC = TTNC + TBOrdemServico!TTNC
        TBOrdemServico.MoveNext
    Loop
End If
TBOrdemServico.Close
txtttok.Text = TTOK
Txtttnc.Text = TTNC
txtProduzida = TTNC + TTOK
txtTPUTIL.Text = TempoTotalPrep
txtTEUTIL.Text = TempoTotalProd
TBProducao.Close

'Carrega dados na lista OS, máquina
Set TBProducao = BD.OpenRecordset("Select * from producaofases where idfase = " & txtos & " and maquina = '" & txtMaquina.Text & "' order by Data, Tempoinicio")
If Not (TBProducao.BOF And TBProducao.EOF) Then
    TBProducao.MoveFirst
    If TBProducao.RecordCount > 7 Then
        totalrecord = TBProducao.RecordCount
        TBProducao.MoveLast
        Contador = 7
        Do While Contador > 1
            TBProducao.MovePrevious
            Contador = Contador - 1
        Loop
    End If
    TBProducao.MoveLast
    PBLista.Min = 0
    PBLista.Max = TBProducao.RecordCount
    PBLista.Value = 1
    Contador = 0
    TBProducao.MoveFirst
    Do Until TBProducao.EOF
        With Lista.ListItems
            .Add , , TBProducao("IDProducao")
            .Item(.Count).SubItems(1) = IIf(IsNull(TBProducao!CodigoDesc) = False, TBProducao!CodigoDesc, 0)
            .Item(.Count).SubItems(2) = TBProducao("descricao")
            txtstatus.Text = TBProducao!Descricao
            .Item(.Count).SubItems(3) = TBProducao("tempoinicio")
            .Item(.Count).SubItems(4) = TBProducao("tempofinal")
            If TBProducao!TempoInicio <> "00:00:00" Then
                TempoFinal = TBProducao!TempoInicio
                TempoUtilizadoDescricao = TBProducao!TempoInicio
            End If
            If TBProducao!Dias <> 0 Then
                TempoTotalDias = IIf(IsNull(TBProducao!TempoTotal), 0, TBProducao!TempoTotal) + TBProducao!Dias
                ElapsedTime (TempoTotalDias)
                .Item(.Count).SubItems(5) = Horatotal
            Else
                .Item(.Count).SubItems(5) = IIf(IsNull(TBProducao("tempototal")) = False, TBProducao!TempoTotal, "")
            End If
            .Item(.Count).SubItems(6) = TBProducao("usuario")
            .Item(.Count).SubItems(7) = IIf(IsNull(TBProducao("quantidade")) = False, TBProducao!Quantidade, 0)
            .Item(.Count).SubItems(8) = IIf(IsNull(TBProducao("reprovada")) = False, TBProducao!reprovada, 0)
            .Item(.Count).SubItems(9) = TBProducao("pronto")
            .Item(.Count).Selected = True
            ULTICOD = TBProducao!CodigoDesc
            ULTIDESC = TBProducao!Descricao
            ULTIOPERADOR = TBProducao!usuario
        End With
        If TBProducao!CodigoDesc = 3 Then Tempoprocesso = TBProducao!Data
        Contador = Contador + 1
        PBLista.Value = Contador
        TBProducao.MoveNext
    Loop
Else
    Lista.ListItems.Clear
    EVENTO = 0
    If PBLista.Value = 0 Then PBLista.Value = 100
End If
TBProducao.Close
'Atualiza o tempo de preparação utilizado
ProcAtualizaPrepUtil
'Atualiza o tempo de execução utilizado
ProcAtualizaExecUtil
ProcAtualizaOF

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Sub ProcAcertaCadMaquina()
On Error GoTo tratar_erro

Set TBMaquinas = BD.OpenRecordset("Select * from cadmaquinas where maquina = '" & txtMaquina.Text & "'")
If TBMaquinas.EOF = False Then
    Set TBProducao = BD.OpenRecordset("Select * from ordemservico where idproducao = " & txtos.Text & "")
    If TBProducao.EOF = False Then
        TBMaquinas.Edit
        If txtstatus <> "" Then
            TBMaquinas!Operador = PubUsuario
            TBMaquinas!ordem = txtof
            TBMaquinas!OS = txtos.Text
            TBMaquinas!CP = TBProducao!CPPECA
            TBMaquinas!CR = TBProducao!CRPECA
            TBMaquinas!TP = TBProducao!Execucao
            TBMaquinas!TR = TBProducao!TEUTIL
            TBMaquinas!Eficiencia = TBProducao!Eficiencia
            txtCodigo.Text = Ultimo
            TBMaquinas!EVENTO = Ultimo
            TBMaquinas!TempoInicio = TempoUltimo
            TBMaquinas!Data = Date
            TBMaquinas!descevento = txtstatus.Text
            If TBMaquinas!custos = True Then
                Set TBCodigoDesc = BD.OpenRecordset("Select Liberar_Posto from CodigoDesc where codigo = " & Ultimo & "")
                If TBCodigoDesc.EOF = False Then
                    If TBCodigoDesc!Liberar_Posto = True Then TBMaquinas!liberada = True Else TBMaquinas!liberada = False
                End If
                TBCodigoDesc.Close
            Else
                TBMaquinas!liberada = True
            End If
        Else
            TBMaquinas!Operador = Null
            TBMaquinas!ordem = Null
            TBMaquinas!OS = Null
            TBMaquinas!CP = Null
            TBMaquinas!CR = Null
            TBMaquinas!TP = Null
            TBMaquinas!TR = Null
            TBMaquinas!Eficiencia = Null
            TBMaquinas!EVENTO = Null
            TBMaquinas!TempoInicio = Null
            TBMaquinas!Data = Null
            TBMaquinas!descevento = Null
            TBMaquinas!liberada = True
        End If
        TBMaquinas.Update
    End If
End If
TBMaquinas.Close

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Sub procabrir()
On Error GoTo tratar_erro
Dim cFase As String
Dim IDOperador As Long

'Rotina de abertura de cadastro
PubUsuario = frmabrir_CB.txtusuario
Ultimo = ""
IDOperador = 0
If frmabrir_CB.txtSenha.Text = "" Then
    MsgBox "É necessário digitar a senha do usuário para identificação.", vbExclamation
    frmabrir_CB.txtSenha.Text = ""
    frmabrir_CB.txtSenha.SetFocus
    Exit Sub
End If
Set TBUsuarios = BD.OpenRecordset("Select * FROM Usuarios WHERE codigo = '" & frmabrir_CB.txtSenha.Text & "'")
If TBUsuarios.BOF And TBUsuarios.EOF Then
    TBUsuarios.Close
    MsgBox "Nome de usuário ou senha inválidos.", vbExclamation
    frmabrir_Ordem.txtSenha.SetFocus
    Exit Sub
Else
    IDOperador = TBUsuarios("IDUsuario")
    Operador = TBUsuarios("usuario")
    TBUsuarios.Close
End If
Set TBOrdem = BD.OpenRecordset("Select * FROM Producao WHERE OF = " & Int(frmabrir_CB.txtof.Text) & "")
If TBOrdem.BOF = False Then
    IDProcesso = TBOrdem("idprocesso")
    'localiza ordem de servico em relacao a maquina, fase, of
    Set TBProcessos = BD.OpenRecordset("Select * from ordemservico where IDPRODUCAO = " & frmabrir_CB.cmbos.Text & "")
    If TBProcessos.EOF = True Then
        TBProcessos.AddNew
        TBProcessos("fase") = frmabrir_CB.txtFase.Text
        TBProcessos("maquina") = frmabrir_CB.cmbmaquina.Text
        TBProcessos("of") = frmabrir_CB.txtof.Text
        TBProcessos("pronto") = "NÃO"
        TBProcessos("preparacao") = "00:00:00"
        TBProcessos("execucao") = "00:00:00"
        TBProcessos!Pcshora = 0
        TBProcessos!pecahora = False
        TBProcessos!TempoPreparacao = "00:00:00"
        TBProcessos!TempoExecucao = "00:00:00"
        TBProcessos!TESegundos = 0
        TBProcessos!pc_te = 0
        TBProcessos("prazofinal") = Format(frmabrir_CB.txtprazo, "dd/mm/yy")
        TBProcessos("quantidade") = frmabrir_CB.txtquant
        TBProcessos("descricao") = frmabrir_CB.txtdescricao.Text
        TBProcessos("desenho") = frmabrir_CB.txtdesenho.Text
        TBProcessos.Update
        txtFase.Text = frmabrir_CB.txtFase
        txtpreparacao.Text = "00:00:00"
        txtexecucao.Text = "00:00:00"
        txttotal.Text = "00:00:00"
        txtquant.Text = frmabrir_CB.txtquant
        txtprazo = frmabrir_CB.txtprazo.Text
        txtdesenho.Text = frmabrir_CB.txtdesenho.Text
        txtdescricao.Text = frmabrir_CB.txtdescricao.Text
        txtof.Text = frmabrir_CB.txtof.Text
        txtos.Text = frmabrir_CB.cmbos.Text
        txtData.Text = Format(Date, "dd/mm/yy")
        frmProducao_CB.Lista.ListItems.Clear
        frmProducao_CB.Show
    Else
        Cprocesso = txtof.Text
        frmProducao_CB.Show
        txtos = TBProcessos!IDProducao
        
        TBProcessos.Edit
        TBProcessos("Maquina") = frmabrir_OS.cmbmaquina.Text
        TBProcessos.Update
        txtMaquina = TBProcessos!Maquina
        
        txtprazo.Text = Format(TBProcessos!Prazofinal, "dd/mm/yy")
        txtdesenho.Text = TBProcessos("desenho")
        txtquant.Text = TBProcessos("quantidade")
        txtof.Text = TBProcessos("OF")
        txtof.Tag = TBProcessos("IDProcesso")
        txtFase.Text = TBProcessos("Fase")
        txtdescricao.Text = TBProcessos("descricao")
        txtData.Text = Format(Date, "dd/mm/yy")
        IDProcesso = TBProcessos("IDProcesso")
        ElapsedTime (IIf(IsNull(TBProcessos!Preparacao), 0, TBProcessos!Preparacao))
        txtpreparacao.Text = Horatotal
        ElapsedTime (IIf(IsNull(TBProcessos!Execucao), 0, TBProcessos!Execucao))
        txtexecucao.Text = Horatotal
        txtPcHoraPrevista = IIf(IsNull(TBProcessos!pc_te) = False, TBProcessos!pc_te, 1)
        txtPcHorareal = IIf(IsNull(TBProcessos!Totalprod), 0, TBProcessos!Totalprod)
        TxtA3Prevista = ProcCalculaSegPC(TBProcessos!Execucao, txtPcHoraPrevista)
        TxtA3Prevista = FormataTempo(TxtA3Prevista.Text)
        txtos.Text = frmabrir_CB.cmbos.Text
        txttotal = IIf(IsNull(TBProcessos!Tempototallote), 0, TBProcessos!Tempototallote)
        procAtualizaProducao
    End If
Else
    OrdemExiste = False
    MsgBox ("Ordem de fabricação não cadastrada."), vbExclamation
    Exit Sub
End If
OrdemExiste = True
txtCodigo.SetFocus

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Sub procabrirNovo()
On Error GoTo tratar_erro
Dim cFase As String
Dim IDOperador As Long

With frmabrir_Ordem_CB
    'Rotina de abertura de cadastro
    Ultimo = ""
    IDOperador = 0
    PubUsuario = .txtusuario
    If .txtSenha.Text = "" Then
        MsgBox "É necessário digitar a senha do usuário para identificação.", vbExclamation
        .txtSenha.Text = ""
        .txtSenha.SetFocus
        Exit Sub
    End If
    Set TBUsuarios = BD.OpenRecordset("Select * FROM Usuarios WHERE codigo = '" & .txtSenha.Text & "'")
    Operador = TBUsuarios("usuario")
    If TBUsuarios.BOF And TBUsuarios.EOF Then
        TBUsuarios.Close
        MsgBox "Nome de usuário ou senha inválidos.", vbExclamation
        .txtSenha.SetFocus
        Exit Sub
    Else
        IDOperador = TBUsuarios("IDUsuario")
        TBUsuarios.Close
    End If
    Set TBOrdem = BD.OpenRecordset("Select * FROM Producao WHERE of = " & .txtordem.Text & "")
    If TBOrdem.BOF = True Then
        TBOrdem.Close
        Set TBOrdem = BD.OpenRecordset("Select * FROM Producao order by of")
        If TBOrdem.EOF = False Then
            TBOrdem.MoveLast
            OF = TBOrdem!OF + 1
        Else
            OF = 1
        End If
        TBOrdem.AddNew
        TBOrdem!OF = OF
        TBOrdem!prazoentrega = .mskprazofinal
        TBOrdem!Quant = .txtquant
        TBOrdem!Data = Date
        TBOrdem!Lista = .txtordem.Text
        TBOrdem!desenho = .txtitem.Text
        TBOrdem!produto = .txtdescricao
        TBOrdem!cliente = .txtcliente.Text
        TBOrdem!pronta = "NÃO"
        TBOrdem!concluida = False
        TBOrdem!dataentrega = Null
        Set TBMaquinas = BD.OpenRecordset("Select * from projproduto where Desenho = '" & .txtitem & "'")
        If TBMaquinas.EOF = False Then
            If TBMaquinas!SubTipoItem = 3 Then TBOrdem!Tipo = "F"
            If TBMaquinas!SubTipoItem = 2 Then TBOrdem!Tipo = "M"
            If TBMaquinas!SubTipoItem = 1 Then TBOrdem!Tipo = "E"
        End If
        TBMaquinas.Close
        TBOrdem.Update
    Else
        OF = TBOrdem!OF
        IDProcesso = TBOrdem("idprocesso")
    End If
    'localiza ordem de servico em relacao a maquina, OS
    If .txtos <> "" Then
        Set TBProcessos = BD.OpenRecordset("Select * from ordemservico where idproducao = " & .txtos & "")
    Else
        Set TBProcessos = BD.OpenRecordset("Select * from ordemservico where OF = " & OF & " AND maquina = '" & .Maquina & "' and fase = " & .txtFase & "")
    End If
    If TBProcessos.EOF = True Then
        TBProcessos.AddNew
        TBProcessos("maquina") = .Maquina
        TBProcessos("of") = OF
        TBProcessos("pronto") = "NÃO"
        TBProcessos("desenho") = .txtitem.Text
        TBProcessos("descricao") = .txtdescricao.Text
        TBProcessos("quantidade") = .txtquant.Text
        TBProcessos("preparacao") = "00:00:00"
        TBProcessos("execucao") = "00:00:00"
        TBProcessos("PRAZOFINAL") = .mskprazofinal.Text
        txtos.Text = TBProcessos!IDProducao
        TBProcessos!Fase = .txtFase.Text
        
        'Verifica se a maquina agrega custos/eficiencia na ordem
        Set TBMaquinas = BD.OpenRecordset("Select custos from cadmaquinas where maquina = '" & .Maquina & "'")
        If TBMaquinas.EOF = False Then
            If TBMaquinas!custos = True Then TBProcessos!custos = True Else TBProcessos!custos = False
        End If
        TBMaquinas.Close
        TBProcessos.Update
        txtMaquina.Text = .Maquina
        txtFase.Text = .txtFase.Text
        txtpreparacao.Text = "00:00:00"
        txtexecucao.Text = "00:00:00"
        txttotal.Text = "00:00:00"
        txtquant.Text = .txtquant.Text
        txtprazo = .mskprazofinal.Text
        txtdesenho.Text = .txtitem.Text
        txtdescricao.Text = .txtdescricao.Text
        txtof.Text = OF
        txtData.Text = Format(Date, "dd/mm/yy")
        frmProducao_CB.Lista.ListItems.Clear
        frmProducao_CB.Show
    Else
        'Se existir ordem de serviço cadastrada
        Cprocesso = txtof.Text
        frmProducao_CB.Show
        txtprazo.Text = .mskprazofinal
        txtdescricao = .txtdescricao
        txtdesenho.Text = IIf(IsNull(TBProcessos("desenho")) = False, TBProcessos!desenho, "")
        txtquant.Text = IIf(IsNull(TBProcessos("quantidade")) = False, TBProcessos!Quantidade, "")
        txtof.Text = IIf(IsNull(TBProcessos("OF")) = False, TBProcessos!OF, "")
        txtof.Tag = IIf(IsNull(TBProcessos("IDProcesso")) = False, TBProcessos!IDProcesso, "")
        txtFase.Text = IIf(IsNull(TBProcessos("Fase")) = False, TBProcessos!Fase, "")
        txtMaquina.Text = IIf(IsNull(TBProcessos("Maquina")) = False, TBProcessos!Maquina, "")
        'txtdescricao.Text = IIf(IsNull(TBProcessos("descricao")) = False, TBProcessos!descricao, "")
        txtpreparacao.Text = IIf(IsNull(TBProcessos("preparacao")) = False, TBProcessos!Preparacao, "")
        txtexecucao.Text = IIf(IsNull(TBProcessos("execucao")) = False, TBProcessos!Execucao, "")
        txtos.Text = IIf(IsNull(TBProcessos!IDProducao) = False, TBProcessos!IDProducao, "")
        TotalDias = "23:59:59"
        hsprevista = IIf(IsNull(TBProcessos!Tempototallote), 0, TBProcessos!Tempototallote)
        '***************************************************
        ' SE A HORA PREVISTA FOR MAIOR QUE UM DIA
        '***************************************************
        diaprevisto = 0
        Do While hsprevista > TotalDias
            hsprevista = hsprevista - 1
            diaprevisto = diaprevisto + 1
        Loop
        '***************************************************
        ' SOMA O TOTAL DE DIAS * 24HS + AS HORAS PREVISTAS
        '***************************************************
        If diaprevisto <> 0 Then
            Hora = Left(hsprevista, Len(hsprevista) - 6) + (24 * diaprevisto)
            TotalDia = Hora & Right(hsprevista, Len(hsprevista) - 2)
        Else
            TotalDia = hsprevista
        End If
        If IsNull(TotalDia) = False Then
            txttotal.Text = TotalDia
        Else
            txttotal.Text = "00:00:00"
        End If
        txtData.Text = Format(Date, "dd/mm/yy")
        IDProcesso = TBProcessos("IDProcesso")
        procAtualizaProducao
    End If
    Exit Sub
    OrdemExiste = True
    cmbdescricao.SetFocus
End With

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub Timer1_Timer()
On Error GoTo tratar_erro

txtHora.Text = Time
If TempoUtilizadoDescricao > 0 Then
    TempoInicio = Now
    tempoutilizado = TempoInicio - TempoUtilizadoDescricao
    ElapsedTime (tempoutilizado)
    TxtTempoUtilizado = Horatotal 'Format(TxtTempoUtilizado.Text, "hh:mm:ss")
Else
    TxtTempoUtilizado = "00:00:00"
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Sub procExcluir()
On Error GoTo tratar_erro

If Lista.ListItems.Count <> 0 And ExcluiSel = True Then
    'Verifica se o evento a ser excluído é TROCA DE POSTO DE TRABALHO e se já existe envento cadastrado após o mesmo.
    Set TBProducao = BD.OpenRecordset("Select * from producaofases where IDProducao = " & Lista.SelectedItem & "")
    If TBProducao.EOF = False Then
        If TBProducao!Descricao = "TROCA DE POSTO DE TRABALHO" And TBProducao!TempoTotal <> "" And TBProducao!TempoTotal <> "00:00:00" Then
            MsgBox ("Não é permitido excluir esse evento, pois existe evento(s) cadastrado(s) após o mesmo."), vbExclamation
            Exit Sub
        End If
        If TBProducao!usuario <> Operador Then
            MsgBox ("Só é permitido o operador " & TBProducao!usuario & " excluir esse evento."), vbExclamation
            Exit Sub
        End If
    End If
    TBProducao.Close
    If EVENTO = "" Then
        MsgBox ("Informe o evento na lista antes de excluir."), vbInformation
        Exit Sub
    End If
    If Format(Lista.SelectedItem.ListSubItems(3), "dd/mm/yy") <> Format(Date, "dd/mm/yy") Then
        MsgBox ("Só é permitido excluir o evento no dia " & Format(Lista.SelectedItem.ListSubItems(3), "dd/mm/yy") & ", dia do apontamento."), vbExclamation
        Exit Sub
    End If
    'Se o evento não for o ultimo da lista
    If Lista.SelectedItem.Index <> Lista.ListItems.Count Then
        MsgBox "Só é possível excluir apenas o último apontamento da lista.", vbExclamation
        Exit Sub
    Else
        If MsgBox("Deseja realmente excluir o evento selecionado?", vbYesNo + vbQuestion) = vbNo Then Exit Sub
        If EVENTO = 3 Then BD.Execute "Update Producao Set pronta = 'NÃO', dataentrega = Null, concluida = False where of = " & txtof.Text & ""
        If Penultimo <> 0 Then
            Set TBProducao = BD.OpenRecordset("Select * from ProducaoFases where codigodesc = " & Penultimo & " and idfase = " & txtos & " order by Data, Tempoinicio")
            If (TBProducao.BOF And TBProducao.EOF) = False Then
                TBProducao.MoveLast
                    
                'Exclui não conformidade da tabela CQ_NC_FABRICA
                BD.Execute "DELETE * FROM CQ_NC_FABRICA WHERE IDProducao = " & TBProducao!IDProducao & ""
                
                Penultimo = TBProducao!CodigoDesc
                PenultimoDesc = TBProducao!Descricao
                'Informa turno
                Turno = TBProducao!Turno
                'Informa data e dias do penultimo evento
                Dataini = TBProducao!Data
                Dias = TBProducao!Dias
                
                'Dados do estoque
                ExcluirAP = True
                Set TBOrdem = BD.OpenRecordset("Select * from producao where OF = " & txtof & "")
                If TBOrdem.EOF = False Then
                    TOK = IIf(IsNull(TBProducao!Quantidade), 0, TBProducao!Quantidade)
                    TNC = IIf(IsNull(TBProducao!reprovada), 0, TBProducao!reprovada)
                    If TBOrdem!Retirar_estoque = True And (TOK <> 0 Or TNC <> 0) Then
                        Set TBCFOP = BD.OpenRecordset("Select * from ordemservico where OF = " & TBOrdem!OF & " order by fase, retrabalho desc, IDproducao")
                        If TBCFOP.EOF = False Then
                            'Verfica se é a primeira OS e retira o material do estoque
                            TBCFOP.MoveFirst
                            If TBCFOP!IDProducao = txtos Then
                                ProcRetirarCancelarEstoque
                            Else 'Verfica se é a última OS e entra com o produto no estoque
                                TBCFOP.MoveLast
                                If TBCFOP!IDProducao = txtos Then ProcEntrarCancelarEstoque
                            End If
                        End If
                        TBCFOP.Close
                    End If
                End If
                TBOrdem.Close
                
                'Grava tempo total do evento por máquina/operador
                Set TBMaquinas = BD.OpenRecordset("Select * from CadMaquinas where Maquina = '" & txtMaquina & "' and Custos = True")
                If TBMaquinas.EOF = False Then
                    Set TBCodigoDesc = BD.OpenRecordset("Select * from CodigoDesc where codigo = " & Penultimo & " and Controlar_Totalizacao = True")
                    If TBCodigoDesc.EOF = False Then
                        ProcExcluirTotalEventoMaq
                        ProcExcluirTotalEventoOpe
                    End If
                    TBCodigoDesc.Close
                End If
                TBProducao.Edit
                TBProducao("tempofinal") = "00:00:00"
                TBProducao("tempototal") = "00:00:00"
                Descricao = TBProducao!CodigoDesc
                txtttok = txtttok - IIf(IsNull(TBProducao!Quantidade) = False, TBProducao!Quantidade, 0)
                TBProducao("quantidade") = 0
                Txtttnc.Text = Txtttnc.Text - IIf(IsNull(TBProducao!reprovada) = False, TBProducao!reprovada, 0)
                TBProducao!reprovada = 0
                TBProducao("dias") = 0
                TBProducao.Update
            End If
        End If
        'Localiza na tabela producaofases todos os registros desta of, maquina, fase e muda pronta = não
        BD.Execute "Update ProducaoFases Set pronto = 'NÃO' where idfase = " & txtos.Text & ""
        'Localiza na tabela a ordem de servico todos os registros desta maquina, of e muda pronta = não
        BD.Execute "Update ordemservico Set pronto = 'NÃO' where idproducao = " & txtos.Text & ""
        
        'Atualiza dados da manutenção
        Set TBAbrir = BD.OpenRecordset("Select * from Manutencao_data where idproducao2 = " & Lista.SelectedItem & "")
        If TBAbrir.EOF = False Then
            Set TBLista = BD.OpenRecordset("Select * from Manutencao_data where IDmanutencao = " & TBAbrir!IDmanutencao & " order by Data")
            If TBLista.BOF = False Then
                TBLista.FindFirst ("ID = " & TBAbrir!Id & "")
                TBLista.MoveNext
                If TBLista.EOF = False Then
                    If TBLista!Status = "Aberta" Then
                        BD.Execute "DELETE * from Manutencao_Checklist where ID_data = " & TBLista!Id & ""
                        BD.Execute "DELETE * from manutencao_data where ID = " & TBLista!Id & ""
                    End If
                End If
            End If
            TBLista.Close
            BD.Execute "Update Manutencao_Checklist Set Check = False where ID_data = " & TBAbrir!Id & ""
        End If
        TBAbrir.Close
        BD.Execute "Update Manutencao_data Set IDproducao = 0 where idproducao = " & Lista.SelectedItem & ""
        BD.Execute "Update Manutencao_data Set status = 'Aberta',IDproducao2 = 0 where idproducao2 = " & Lista.SelectedItem & ""
        
        BD.Execute "DELETE * FROM ProducaoFases WHERE IDProducao = " & Lista.SelectedItem.Text & ""
        
        ProcVerificaTurno
        
        'Atualiza lista de eventos cadastrados
        procAtualizaProducao
        'Grava o status da OS e da OF
        ProcGravarStatusOSOF
        'Atualiza dados do Turno
        Set TBProducaofases = BD.OpenRecordset("Select * from ProducaoFases where idfase = " & txtos.Text & " and Turno = " & Turno & " and data = CDate('" & Format(Dataini, "dd/mm/yyyy") & "') and (codigodesc = 1 or codigodesc = 2)")
        If TBProducaofases.EOF = False Then
            ProcGravarTurno
        Else
            BD.Execute "Delete * from ProducaoFases_Turno where OS = " & txtos.Text & " and Turno = " & Turno & " and data = CDate('" & Format(Date, "dd/mm/yyyy") & "')"
        End If
        TBProducaofases.Close
        'Atualiza dados da Máquina
        Set TBProducaofases = BD.OpenRecordset("Select * from ProducaoFases where idfase = " & txtos.Text & " and maquina = '" & txtMaquina.Text & "' and data = CDate('" & Format(Dataini, "dd/mm/yyyy") & "') and (codigodesc = 1 or codigodesc = 2)")
        If TBProducaofases.EOF = False Then
            ProcGravarMaquina
        Else
            BD.Execute "Delete * from ProducaoFases_Maquina where OS = " & txtos.Text & " and maquina = '" & txtMaquina.Text & "' and data = CDate('" & Format(Date, "dd/mm/yyyy") & "')"
        End If
        TBProducaofases.Close
        'Atualiza dados do Operador
        Set TBProducaofases = BD.OpenRecordset("Select * from ProducaoFases where idfase = " & txtos.Text & " and usuario = '" & Operador & "' and data = CDate('" & Format(Dataini, "dd/mm/yyyy") & "') and (codigodesc = 1 or codigodesc = 2)")
        If TBProducaofases.EOF = False Then
            ProcGravarOperador
        Else
            BD.Execute "Delete * from ProducaoFases_Operador where OS = " & txtos.Text & " and usuario = '" & Operador & "' and data = CDate('" & Format(Date, "dd/mm/yyyy") & "')"
        End If
        TBProducaofases.Close
        'Acerta cadastro na máquina
        ProcAcertaCadMaquina
    End If
End If
ExcluiSel = False
txtCodigo.Text = ""
txtcodigoDesc.Text = ""
txtdescricao.Text = ""
txtTOK.Text = ""
txtTNC.Text = ""

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub txtCodigo_Change()
On Error GoTo tratar_erro

Set TBProducao = BD.OpenRecordset("select * from manutencao where IDmaquina = '" & txtMaquina & "' and Controlada = true")
If TBProducao.EOF = False Then
    Set TBProcessosDet = BD.OpenRecordset("select * from manutencao_data where idManutencao = " & TBProducao!código & " and status = 'Aberta' and IDProducao = 0 and data <= #" & Date & "# ")
    If TBProcessosDet.EOF = False Then
        If IsNumeric(txtCodigo) = True Then
            Set TBCodigoDesc = BD.OpenRecordset("Select * from  codigodesc where descricao = 'MÁQUINA EM MANUTENÇÃO' or descricao = 'MANUTENÇÃO PREVENTIVA' or descricao = 'MANUTENÇÃO CORRETIVA'")
            If TBCodigoDesc.EOF = False Then
                txtcodigoDesc.Text = TBCodigoDesc!Descricao
            Else
                txtcodigoDesc = ""
            End If
            TBCodigoDesc.Close
        End If
        Exit Sub
    End If
End If
If IsNumeric(txtCodigo) = True Then
    Set TBCodigoDesc = BD.OpenRecordset("Select * from  codigodesc where codigo = " & txtCodigo.Text & "")
    If TBCodigoDesc.EOF = False Then
        txtcodigoDesc.Text = TBCodigoDesc!Descricao
    Else
        txtCodigo.Text = ""
        txtcodigoDesc.Text = ""
    End If
    TBCodigoDesc.Close
Else
    txtCodigo.Text = ""
    txtcodigoDesc.Text = ""
    txtdescricao.Text = ""
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub txtCodigo_GotFocus()
On Error GoTo tratar_erro
    
txtCodigo.Text = ""

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub txtTNC_Change()
On Error GoTo tratar_erro

If IsNumeric(txtTNC) = False Then txtTNC.Text = ""

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub txtTOK_Change()
On Error GoTo tratar_erro

If IsNumeric(txtTOK) = False Then txtTOK.Text = ""

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Sub ProcLista12Ultimos()
On Error GoTo tratar_erro

Lista.ListItems.Clear
Set TBProducao = BD.OpenRecordset("Select * from producaofases where idfase = " & txtos & " and Maquina = '" & txtMaquina & "' order by Data, Tempoinicio")
If TBProducao.EOF = False Then
    If TBProducao.RecordCount > 7 Then
        TBProducao.MoveLast
        contador2 = 7
        Do While contador2 > 1
            TBProducao.MovePrevious
            contador2 = contador2 - 1
        Loop
    End If
    Do Until TBProducao.EOF
        With Lista.ListItems
            .Add , , TBProducao("IDProducao")
            .Item(.Count).SubItems(1) = IIf(IsNull(TBProducao!CodigoDesc) = False, TBProducao!CodigoDesc, 0)
            .Item(.Count).SubItems(2) = TBProducao("descricao")
            txtstatus.Text = TBProducao!Descricao
            .Item(.Count).SubItems(3) = TBProducao("tempoinicio")
            .Item(.Count).SubItems(4) = TBProducao("tempofinal")
            If TBProducao!TempoInicio <> "00:00:00" Then
                TempoFinal = TBProducao!TempoInicio
                TempoUtilizadoDescricao = TBProducao!TempoInicio
            End If
            If TBProducao!Dias <> 0 Then
                TempoTotalDias = IIf(IsNull(TBProducao!TempoTotal), 0, TBProducao!TempoTotal) + TBProducao!Dias
                ElapsedTime (TempoTotalDias)
                .Item(.Count).SubItems(5) = Horatotal
            Else
                .Item(.Count).SubItems(5) = IIf(IsNull(TBProducao("tempototal")) = False, TBProducao!TempoTotal, "")
            End If
            .Item(.Count).SubItems(6) = TBProducao("usuario")
            .Item(.Count).SubItems(7) = IIf(IsNull(TBProducao("quantidade")) = False, TBProducao!Quantidade, 0)
            .Item(.Count).SubItems(8) = IIf(IsNull(TBProducao("reprovada")) = False, TBProducao!reprovada, 0)
            .Item(.Count).SubItems(9) = TBProducao("pronto")
            .Item(.Count).Selected = True
        End With
        TBProducao.MoveNext
    Loop
End If
TBProducao.Close
If PBLista.Value = 0 Then PBLista.Value = 100

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcRetirarCancelarEstoque()
On Error GoTo tratar_erro

Set TBItem = BD.OpenRecordset("Select * from projproduto where desenho = '" & TBOrdem!desenho & "'")
If TBItem.EOF = False Then
    Set TBProcessos = BD.OpenRecordset("Select * from projconjunto where codproduto = " & TBItem!codProduto & "")
    If TBProcessos.EOF = False Then
        Do While TBProcessos.EOF = False
            Set TBMaterial = BD.OpenRecordset("Select * from producaomaterial where OF = " & TBOrdem!OF & " and Codigo = '" & TBProcessos!desenho & "' and Titular = True")
            If TBMaterial.EOF = False Then
                If IsNull(TBMaterial!NF) = False And TBMaterial!NF <> "" Then
                    If TBProcessos!unidade = "KG" Then
                        peso = TBProcessos!pesototal
                        qtdeliberar = peso * (TOK + TNC)
                    End If
                    If TBProcessos!unidade = "MT" Then
                        peso = (TBProcessos!Dimensoes * TBProcessos!Quantidade) / 1000
                        qtdeliberar = peso * (TOK + TNC)
                        TBMaterial!Requisitado = peso * TBOrdem!Quant
                    End If
                    If TBProcessos!unidade = "MM" Then
                        peso = TBProcessos!Dimensoes * TBProcessos!Quantidade
                        qtdeliberar = peso * (TOK + TNC)
                        TBMaterial!Requisitado = peso * TBOrdem!Quant
                    End If
                    If TBProcessos!unidade <> "KG" And TBProcessos!unidade <> "MT" And TBProcessos!unidade <> "MM" Then
                        qtdeliberar = TBProcessos!Quantidade * (TOK + TNC)
                        TBMaterial!Requisitado = TBProcessos!Quantidade * TBOrdem!Quant
                    End If
                    
                    'Retira do estoque e agrega o custo na ordem
                    If TBOrdem!Consignacao = True Then
                        If IsNull(TBMaterial!Local_armazenamento) = False And TBMaterial!Local_armazenamento <> "" Then
                            Set TBEstoque = BD.OpenRecordset("Select * from estoque_controle where desenho = '" & TBMaterial!CODIGO & "' and nf = '" & TBMaterial!NF & "' and Corrida = '" & TBMaterial!Corrida & "' and Certificado = '" & TBMaterial!Certificado & "' and cliente = '" & TBMaterial!cliente & "' and consignacao = true and lote = '" & TBMaterial!NF & "' and local_armaz = '" & TBMaterial!Local_armazenamento & "'")
                        Else
                            Set TBEstoque = BD.OpenRecordset("Select * from estoque_controle where desenho = '" & TBMaterial!CODIGO & "' and nf = '" & TBMaterial!NF & "' and Corrida = '" & TBMaterial!Corrida & "' and Certificado = '" & TBMaterial!Certificado & "' and cliente = '" & TBMaterial!cliente & "' and consignacao = true and lote = '" & TBMaterial!NF & "' and isnull(local_armaz)")
                        End If
                    Else
                        If IsNull(TBMaterial!Local_armazenamento) = False And TBMaterial!Local_armazenamento <> "" Then
                            Set TBEstoque = BD.OpenRecordset("Select * from estoque_controle where desenho = '" & TBMaterial!CODIGO & "' and lote = '" & TBMaterial!NF & "' and certificado = '" & TBMaterial!Certificado & "' and corrida = '" & TBMaterial!Corrida & "' and local_armaz = '" & TBMaterial!Local_armazenamento & "'")
                        Else
                            Set TBEstoque = BD.OpenRecordset("Select * from estoque_controle where desenho = '" & TBMaterial!CODIGO & "' and lote = '" & TBMaterial!NF & "' and certificado = '" & TBMaterial!Certificado & "' and corrida = '" & TBMaterial!Corrida & "' and isnull(local_armaz)")
                        End If
                    End If
                    If TBEstoque.EOF = False Then
                        
                        Set TBProduto = BD.OpenRecordset("Select * from Estoque_movimentacao where IDestoque = " & TBEstoque!idestoque & " and Documento = '" & TBOrdem!OF & "' and Data = CDate('" & Format(Date, "dd/mm/yyyy") & "') and (Operacao = 'SAIDA_ORDEM' or Operacao = 'SAIDA_ORDEM_PARCIAL')")
                        If TBProduto.EOF = False Then
                            TBProduto.Edit
                        Else
                            TBProduto.AddNew
                        End If
                        TBProduto!Documento = TBOrdem!OF
                        TBProduto!LOTE = TBMaterial!NF
                        TBProduto!desenho = TBMaterial!CODIGO
                        TBProduto!Data = Date
                        TBProduto!Descricao = TBMaterial!Descricao
                        TBProduto!Familia = TBEstoque!Classe
                        TBProduto!requisitante = cmbOperador
                        TBProduto!Responsavel = PubUsuario
                        TBProduto!idestoque = TBEstoque!idestoque
                        TBProduto!OE = TBOrdem!OF
                        If ExcluirAP = False Then TBProduto!Saida = qtdeliberar Else TBProduto!Saida = TBProduto!Saida - qtdeliberar
                    
                        'Atualiza valor do material no estoque
                        If ExcluirAP = False Then quantestoque = qtdeliberar Else quantestoque = TBProduto!Saida
                        TBProduto!VlrUnit = IIf(IsNull(TBEstoque!valor_unitario), 0, Format(TBEstoque!valor_unitario, "###,##0.00000"))
                        TBProduto!VlrTotal = Format(quantestoque * TBProduto!VlrUnit, "###,##0.00")
                    
                        'verifica se a quantidade retirada e menor q a quant. solicitada
                        TBMaterial.Edit
                        If ExcluirAP = False Then
                            If qtdeliberar < TBMaterial!Requisitado Then
                                TBProduto!Operacao = "SAIDA_ORDEM_PARCIAL"
                            Else
                                TBProduto!Operacao = "SAIDA_ORDEM"
                            End If
                            TBProduto.Update
                        Else
                            TBProduto.Update
                            If TBProduto!Saida <= 0 And TBProduto.EOF = False Then TBProduto.Delete
                        End If
                        
                        Saida = 0
                        Set TBFiltro = BD.OpenRecordset("Select * from Estoque_movimentacao where IDestoque = " & TBEstoque!idestoque & " and Documento = '" & TBOrdem!OF & "' and (Operacao = 'SAIDA_ORDEM' or Operacao = 'SAIDA_ORDEM_PARCIAL')")
                        If TBFiltro.EOF = False Then
                            Do While TBFiltro.EOF = False
                                Saida = Saida + TBFiltro!Saida
                                TBFiltro.MoveNext
                            Loop
                        End If
                        TBFiltro.Close
                        If Saida = 0 Then
                            TBMaterial!Saida = "NÃO"
                        ElseIf Saida < TBMaterial!Requisitado Then
                                TBMaterial!Saida = "PARCIAL"
                            Else
                                TBMaterial!Saida = "SIM"
                        End If
                        TBMaterial.Update
                        
                        qtdeliberada = 0
                        QtdeSaida = 0
                        QtdeEstoque = 0
                        If TBItem!Estoque = True Then
                            Set TBProduto = BD.OpenRecordset("Select * from Estoque_movimentacao where IdEstoque = " & TBEstoque!idestoque & "")
                            If TBProduto.EOF = False Then
                                Do While TBProduto.EOF = False
                                    qtdeliberada = qtdeliberada + TBProduto!Entrada
                                    QtdeSaida = QtdeSaida + TBProduto!Saida
                                    TBProduto.MoveNext
                                Loop
                                QtdeEstoque = qtdeliberada - QtdeSaida
                                NovoValor = Replace(QtdeEstoque, ",", ".")
                                BD.Execute "Update Estoque_movimentacao Set estoque_venda = " & NovoValor & " where IDestoque = " & TBEstoque!idestoque & " and Documento = '" & TBOrdem!OF & "' and Data = CDate('" & Format(Date, "dd/mm/yyyy") & "') and (Operacao = 'SAIDA_ORDEM' or Operacao = 'SAIDA_ORDEM_PARCIAL')"
                            End If
                        End If
                        TBProduto.Close
                        
                        TBEstoque.Edit
                        TBEstoque!Estoque_real = QtdeEstoque
                        TBEstoque!estoque_venda = QtdeEstoque
                        
                        'Atualiza valor do material no estoque
                        quantestoque = TBEstoque!Estoque_real
                        TBEstoque!Valor_Total = Format(IIf(IsNull(TBEstoque!valor_unitario), 0, TBEstoque!valor_unitario) * quantestoque, "###,##0.00")
                        TBEstoque.Update
                        TBEstoque.Close
                    End If
                    If TBOrdem!Consignacao = False Then
                        If TBOrdem!Tipo = "E" Then
                            'Verifica custo de NC
                            If IsNumeric(TBMaterial!NF) = True Then
                                Set TBFiltro = BD.OpenRecordset("Select * from producao where of = " & TBMaterial!NF & "")
                                If TBFiltro.EOF = False Then
                                                                          'ORDEM       QTDE. PREVISTA                                  QTDE. OK                                                QT. PROD.(OK+NC)                                                                                             CUSTO LOTE                                                                CUSTO PEÇA                                                        CUSTO TERCEIROS                                                               CUSTO MATERIAL                                                                  TIPO DA ORDEM
                                    ValorUnitario = CalculaValorUnitOrdem(TBFiltro!OF, IIf(IsNull(TBFiltro!Quant), 0, TBFiltro!Quant), IIf(IsNull(TBFiltro!QuantProd), 0, TBFiltro!QuantProd), IIf(IsNull(TBFiltro!QuantProd), 0, TBFiltro!QuantProd) + IIf(IsNull(TBFiltro!QuantNC), 0, TBFiltro!QuantNC), IIf(IsNull(TBFiltro!CTTReal), 0, Format(TBFiltro!CTTReal, "###,##0.00")), IIf(IsNull(TBFiltro!CPR), 0, Format(TBFiltro!CPR, "###,##0.00")), IIf(IsNull(TBFiltro!CTServico), 0, Format(TBFiltro!CTServico, "###,##0.00")), IIf(IsNull(TBFiltro!CTMaterial), 0, Format(TBFiltro!CTMaterial, "###,##0.00")), TBFiltro!Tipo)
                                    OF = TBFiltro!OF
                                    
                                    QtdeSaida = qtdeliberar
                                    TBOrdem.Edit
                                    If Permitido = True Then
                                        'Valor do material por peça x qtde. refugada
                                        If QuantsolicitadoN1 <> 0 Then Valor2 = Format(Valor_CSLL_Serv * QuantComprado, "###,##0.00")
                                        
                                        'Valor total do serviço - valor total NC serviço / qtde. OK = Valor unitário do serviço
                                        If qt <> 0 Then Valor1 = (IIf(IsNull(TBFiltro!CTServico), 0, TBFiltro!CTServico) - Valor1) / qt Else Valor1 = 0
                                        TBOrdem!CTServico = IIf(IsNull(TBOrdem!CTServico), 0, TBOrdem!CTServico) + (Valor1 * QtdeSaida)
                                        
                                        'Valor total do material - valor total NC material / qtde. OK = Valor unitário do material
                                        If qt <> 0 Then Valor2 = (IIf(IsNull(TBFiltro!CTMaterial), 0, TBFiltro!CTMaterial) - Valor2) / qt Else Valor2 = 0
                                        TBOrdem!CTMaterial = IIf(IsNull(TBOrdem!CTMaterial), 0, TBOrdem!CTMaterial) + (Valor2 * QtdeSaida)
                                        
                                        'Valor total de MO - valor total NC MO / qtde. OK = Valor unitário de MO
                                        If qt <> 0 Then Valor3 = (IIf(IsNull(TBFiltro!CTTReal), 0, TBFiltro!CTTReal) - Valor3) / qt Else Valor3 = 0
                                        TBOrdem!CTTReal = IIf(IsNull(TBOrdem!CTTReal), 0, TBOrdem!CTTReal) + (Valor3 * QtdeSaida)
                                    Else
                                        'Valor total do serviço + Valor unitário do serviço * qtde. OK
                                        Valor_IPI = 0
                                        Set TBProducao = BD.OpenRecordset("Select * from ordemservico where OF = " & OF & " order by fase, retrabalho desc, IDproducao")
                                        If TBProducao.EOF = False Then
                                            Do While TBProducao.EOF = False
                                                'Soma valor unitário do SERVIÇO na OS
                                                If IsNull(TBProducao!Totalprod) = False And TBProducao!Totalprod <> "" And TBProducao!Totalprod <> "0" Then Valor_IPI = Format(Valor_IPI + (TBProducao!CTServico / TBProducao!Totalprod), "###,##0.00")
                                                TBProducao.MoveNext
                                            Loop
                                        End If
                                        TBProducao.Close
                                        TBOrdem!CTServico = IIf(IsNull(TBOrdem!CTServico), 0, TBOrdem!CTServico) + (Valor_IPI * QtdeSaida)
                                        
                                        'Valor total do material + Valor unitário do material * qtde. OK
                                        TBOrdem!CTMaterial = IIf(IsNull(TBOrdem!CTMaterial), 0, TBOrdem!CTMaterial) + (Valor_CSLL_Serv * QtdeSaida)
                                        
                                        'Valor unitário de MO * qtde. OK
                                        Valor3 = IIf(IsNull(TBFiltro!CPR), 0, TBFiltro!CPR)
                                        TBOrdem!CTTReal = IIf(IsNull(TBOrdem!CTTReal), 0, TBOrdem!CTTReal) + (Valor3 * QtdeSaida)
                                    End If
                                    TBOrdem.Update
                                End If
                                TBFiltro.Close
                            Else
                                'Custo material
                                Valor = 0
                                Set TBEstoque = BD.OpenRecordset("Select * from Estoque_movimentacao where Documento = '" & TBOrdem!OF & "' and (Operacao = 'SAIDA_ORDEM' or Operacao = 'SAIDA_ORDEM_PARCIAL')")
                                If TBEstoque.EOF = False Then
                                    Do While TBEstoque.EOF = False
                                        Valor = Valor + IIf(IsNull(TBEstoque!VlrTotal), 0, TBEstoque!VlrTotal)
                                        TBEstoque.MoveNext
                                    Loop
                                End If
                                TBEstoque.Close
                                TBOrdem.Edit
                                TBOrdem!CTMaterial = Format(Valor, "###,##0.00")
                                TBOrdem.Update
                            End If
                        Else
                            'Custo material
                            Valor = 0
                            Set TBEstoque = BD.OpenRecordset("Select * from Estoque_movimentacao where Documento = '" & TBOrdem!OF & "' and (Operacao = 'SAIDA_ORDEM' or Operacao = 'SAIDA_ORDEM_PARCIAL')")
                            If TBEstoque.EOF = False Then
                                Do While TBEstoque.EOF = False
                                    Valor = Valor + IIf(IsNull(TBEstoque!VlrTotal), 0, TBEstoque!VlrTotal)
                                    TBEstoque.MoveNext
                                Loop
                            End If
                            TBEstoque.Close
                            TBOrdem.Edit
                            TBOrdem!CTMaterial = Format(Valor, "###,##0.00")
                            TBOrdem.Update
                        End If
                    End If
                End If
            End If
            TBMaterial.Close
            TBProcessos.MoveNext
        Loop
    End If
    TBProcessos.Close
End If
TBItem.Close

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

Private Sub ProcEntrarCancelarEstoque()
On Error GoTo tratar_erro

Valortotal = 0
quantestoque = 0
Set TBFI = BD.OpenRecordset("Select Estoque_Localarmazenamento_criar.* FROM Estoque_Localarmazenamento_criar INNER JOIN Estoque_Localarmazenamento ON Estoque_Localarmazenamento.idemb_locarm = Estoque_Localarmazenamento_criar.id where Estoque_Localarmazenamento.codinterno = '" & TBOrdem!desenho & "'")
If TBFI.EOF = False Then
    Set TBEstoque = BD.OpenRecordset("Select * from estoque_controle where desenho = '" & TBOrdem!desenho & "' and lote = '" & TBOrdem!OF & "' and certificado = '" & 0 & "' and corrida = '" & 0 & "' and local_armaz = '" & TBFI!Descricao & "'")
    If TBEstoque.EOF = False Then
        TBEstoque.Edit
    Else
        TBEstoque.AddNew
    End If
    Set TBProduto = BD.OpenRecordset("Select * from Estoque_movimentacao where IDestoque = " & TBEstoque!idestoque & " and Lote = '" & TBOrdem!OF & "' and Data = CDate('" & Format(Date, "dd/mm/yyyy") & "') and (Operacao = 'ENTRADA_ORDEM' or Operacao = 'ENTRADA_ORDEM_PARCIAL')")
    If TBProduto.EOF = False Then
        TBProduto.Edit
    Else
        TBProduto.AddNew
    End If
    TBEstoque!LOTE = TBOrdem!OF
    TBProduto!LOTE = TBOrdem!OF
    TBProduto!Documento = TBOrdem!OF
    TBEstoque!desenho = TBOrdem!desenho
    TBProduto!desenho = TBOrdem!desenho
    
    'Atualiza valor do produto/item no estoque
    Set TBItem = BD.OpenRecordset("Select * from projproduto where desenho = '" & TBOrdem!desenho & "'")
    If TBItem.EOF = False Then
        If TBItem!Estoque = True Then ControlaEstoque = True Else ControlaEstoque = False
        TBEstoque!UN = TBItem!unidade
        TBEstoque!Classe = TBItem!Classe
        TBProduto!Familia = TBItem!Classe
    End If
    TBItem.Close
                                          'ORDEM      QTDE. PREVISTA                                QTDE. OK                                              QT. PROD.(OK+NC)                                                                                         CUSTO LOTE                                                              CUSTO PEÇA                                                      CUSTO TERCEIROS                                                             CUSTO MATERIAL                                                                TIPO DA ORDEM
    Valor_Produto = CalculaValorUnitOrdem(TBOrdem!OF, IIf(IsNull(TBOrdem!Quant), 0, TBOrdem!Quant), IIf(IsNull(TBOrdem!QuantProd), 0, TBOrdem!QuantProd), IIf(IsNull(TBOrdem!QuantProd), 0, TBOrdem!QuantProd) + IIf(IsNull(TBOrdem!QuantNC), 0, TBOrdem!QuantNC), IIf(IsNull(TBOrdem!CTTReal), 0, Format(TBOrdem!CTTReal, "###,##0.00")), IIf(IsNull(TBOrdem!CPR), 0, Format(TBOrdem!CPR, "###,##0.00")), IIf(IsNull(TBOrdem!CTServico), 0, Format(TBOrdem!CTServico, "###,##0.00")), IIf(IsNull(TBOrdem!CTMaterial), 0, Format(TBOrdem!CTMaterial, "###,##0.00")), TBOrdem!Tipo)
    OF = TBOrdem!OF
    
    Valortotal = Valor_Produto
    TBEstoque!valor_unitario = Format(Valortotal, "###,##0.00000")
    
    'Estoque_movimentação
    quantestoque = TOK + TNC
    TBProduto!VlrUnit = Format(Valortotal, "###,##0.00000")
    TBProduto!VlrTotal = Format(quantestoque * Valortotal, "###,##0.00")
    
    TBEstoque!Descricao = TBOrdem!produto
    TBProduto!Descricao = TBOrdem!produto
    TBProduto!Data = Date
    TBEstoque!Data = Date
    TBEstoque!responsável = cmbOperador
    TBProduto!Responsavel = PubUsuario
    TBEstoque!Certificado = 0
    TBEstoque!Corrida = 0
    
    TBEstoque!local_armaz = TBFI!Descricao
    If ExcluirAP = False Then
        TBProduto!Entrada = quantestoque
        TBEstoque!Qtde = quantestoque
    Else
        TBProduto!Entrada = TBProduto!Entrada - quantestoque
        TBEstoque!Qtde = TBEstoque!Qtde - quantestoque
    End If
    
    Qtde = TBOrdem!Quant
    Entrada = quantestoque
    If Entrada >= Qtde Then
        TBEstoque!Status = "ENTRADA_ORDEM"
        TBProduto!Operacao = "ENTRADA_ORDEM"
    ElseIf Entrada < Qtde Then
            TBEstoque!Status = "ENTRADA_ORDEM_PARCIAL"
            TBProduto!Operacao = "ENTRADA_ORDEM_PARCIAL"
    End If
    TBEstoque!cliente = IIf(IsNull(TBOrdem!cliente), "", TBOrdem!cliente)
    TBProduto!idestoque = TBEstoque!idestoque
    If ExcluirAP = False Then
        TBProduto.Update
    Else
        If TBProduto!Entrada <= 0 And TBProduto.EOF = False Then TBProduto.Delete
    End If
    
    qtdeliberada = 0
    QtdeSaida = 0
    QtdeEstoque = 0
    If ControlaEstoque = True Then
        Set TBProduto = BD.OpenRecordset("Select * from Estoque_movimentacao where IdEstoque = " & TBEstoque!idestoque & "")
        If TBProduto.EOF = False Then
            Do While TBProduto.EOF = False
                qtdeliberada = qtdeliberada + TBProduto!Entrada
                QtdeSaida = QtdeSaida + TBProduto!Saida
                TBProduto.MoveNext
            Loop
            QtdeEstoque = qtdeliberada - QtdeSaida
            NovoValor = Replace(QtdeEstoque, ",", ".")
            BD.Execute "Update Estoque_movimentacao Set estoque_venda = " & NovoValor & " where IDestoque = " & TBEstoque!idestoque & " and Lote = '" & TBOrdem!OF & "' and Data = CDate('" & Format(Date, "dd/mm/yyyy") & "') and (Operacao = 'ENTRADA_ORDEM' or Operacao = 'ENTRADA_ORDEM_PARCIAL')"
        Else
            TBEstoque.Delete
            TBProduto.Close
            Exit Sub
        End If
    End If
    TBProduto.Close
    
    TBEstoque!Estoque_real = QtdeEstoque
    TBEstoque!estoque_venda = QtdeEstoque
    
    'Atualiza valor do material no estoque
    quantestoque = TBEstoque!Estoque_real
    TBEstoque!Valor_Total = Format(Valor_Produto * quantestoque, "###,##0.00")
    TBEstoque.Update
    TBEstoque.Close
End If

Exit Sub
tratar_erro:
    MsgBox ("Descrição do erro : " + Error()), vbCritical
    Exit Sub
End Sub

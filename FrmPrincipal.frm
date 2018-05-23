VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form FrmPrincipal 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   9030
   ClientLeft      =   60
   ClientTop       =   60
   ClientWidth     =   11955
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "FrmPrincipal.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   9030
   ScaleWidth      =   11955
   WindowState     =   2  'Maximized
   Begin VB.Timer TimerHora 
      Interval        =   1000
      Left            =   3840
      Top             =   240
   End
   Begin VB.Frame FraDesenv 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6375
      Left            =   10680
      TabIndex        =   216
      Top             =   8520
      Width           =   10935
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel48 
         Height          =   255
         Left            =   4200
         OleObjectBlob   =   "FrmPrincipal.frx":0CCA
         TabIndex        =   217
         Top             =   2640
         Width           =   2655
      End
   End
   Begin VB.Frame FraVenda 
      Caption         =   "Consulta de Vendas"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6375
      Left            =   360
      TabIndex        =   164
      Top             =   1920
      Width           =   10935
      Begin VB.CommandButton CmdLimparVend 
         Caption         =   "&Limpar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9240
         TabIndex        =   151
         ToolTipText     =   "Limpar pesquisa"
         Top             =   1080
         Width           =   1335
      End
      Begin FPSpread.vaSpread GridVenda 
         Height          =   3375
         Left            =   240
         TabIndex        =   152
         Top             =   2040
         Width           =   10455
         _Version        =   393216
         _ExtentX        =   18441
         _ExtentY        =   5953
         _StockProps     =   64
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   14737632
         GridColor       =   8421504
         MaxCols         =   9
         MaxRows         =   0
         OperationMode   =   2
         RetainSelBlock  =   0   'False
         RowHeaderDisplay=   0
         ScrollBarExtMode=   -1  'True
         ScrollBarMaxAlign=   0   'False
         ScrollBarShowMax=   0   'False
         SelectBlockOptions=   0
         ShadowColor     =   14737632
         ShadowDark      =   8421504
         ShadowText      =   0
         SpreadDesigner  =   "FrmPrincipal.frx":0D46
      End
      Begin VB.TextBox TxtDtVenda2 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7320
         MaxLength       =   10
         TabIndex        =   147
         Text            =   "__/__/____"
         ToolTipText     =   "Maior data da venda"
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox TxtDtVenda1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5760
         MaxLength       =   10
         TabIndex        =   146
         Text            =   "__/__/____"
         ToolTipText     =   "Menor data da venda"
         Top             =   600
         Width           =   1215
      End
      Begin VB.ComboBox CboTipoVenda 
         Height          =   315
         ItemData        =   "FrmPrincipal.frx":138D
         Left            =   1560
         List            =   "FrmPrincipal.frx":138F
         Style           =   2  'Dropdown List
         TabIndex        =   148
         ToolTipText     =   "Tipo da venda"
         Top             =   1200
         Width           =   2775
      End
      Begin VB.TextBox TxtVendedor 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5760
         MaxLength       =   200
         TabIndex        =   149
         ToolTipText     =   "Nome do vendedor"
         Top             =   1200
         Width           =   2775
      End
      Begin VB.TextBox TxtCliVend 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1560
         MaxLength       =   200
         TabIndex        =   145
         ToolTipText     =   "Nome do cliente"
         Top             =   600
         Width           =   2775
      End
      Begin VB.CommandButton CmdPesqVenda 
         Caption         =   "&Pesquisar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9240
         TabIndex        =   150
         ToolTipText     =   "Pesquisar vendas"
         Top             =   360
         Width           =   1335
      End
      Begin VB.Frame FraBotaoVenda 
         Height          =   735
         Left            =   120
         TabIndex        =   165
         Top             =   5520
         Width           =   10695
         Begin VB.CommandButton CmdImprimirProp 
            Caption         =   "Imprimir &Proposta"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   8520
            TabIndex        =   159
            ToolTipText     =   "Imprimir proposta de venda"
            Top             =   240
            Width           =   2055
         End
         Begin VB.CommandButton CmdImprimirCarne 
            Caption         =   "Imprimir &Carnê"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6600
            TabIndex        =   158
            ToolTipText     =   "Imprimir carnê do crediário"
            Top             =   240
            Width           =   1815
         End
         Begin VB.CommandButton CmdImprimirRecibo 
            Caption         =   "Imprimir &Recibo"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4680
            TabIndex        =   157
            ToolTipText     =   "Imprimir recibo da venda"
            Top             =   240
            Width           =   1815
         End
         Begin VB.CommandButton CmdIncluirVenda 
            Caption         =   "&Incluir"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   120
            TabIndex        =   153
            ToolTipText     =   "Incluir venda"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton CmdImprimirVenda 
            Caption         =   "I&mprimir Relatório"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   2520
            TabIndex        =   156
            ToolTipText     =   "Imprimir consulta de vendas"
            Top             =   240
            Width           =   2055
         End
         Begin VB.CommandButton CmdExcluirVenda 
            Caption         =   "&Excluir"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   1320
            TabIndex        =   154
            ToolTipText     =   "Excluir venda"
            Top             =   240
            Width           =   1095
         End
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel15 
         Height          =   255
         Left            =   480
         OleObjectBlob   =   "FrmPrincipal.frx":1391
         TabIndex        =   166
         Top             =   600
         Width           =   6975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel16 
         Height          =   255
         Left            =   480
         OleObjectBlob   =   "FrmPrincipal.frx":14B3
         TabIndex        =   167
         Top             =   1200
         Width           =   5295
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblNumTotalVend 
         Height          =   255
         Left            =   7560
         OleObjectBlob   =   "FrmPrincipal.frx":159B
         TabIndex        =   207
         Top             =   1680
         Width           =   3015
      End
   End
   Begin VB.Frame FraCliente 
      Caption         =   "Consulta de Clientes"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6375
      Left            =   360
      TabIndex        =   197
      Top             =   1920
      Width           =   10935
      Begin VB.CommandButton CmdLimparCli 
         Caption         =   "&Limpar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9240
         TabIndex        =   139
         ToolTipText     =   "Limpar pesquisa"
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CommandButton CmdPesqCli 
         Caption         =   "&Pesquisar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9240
         TabIndex        =   138
         ToolTipText     =   "Pesquisar clientes"
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox TxtBairroCli 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5040
         MaxLength       =   60
         TabIndex        =   134
         ToolTipText     =   "Bairro do cliente"
         Top             =   480
         Width           =   2415
      End
      Begin VB.TextBox TxtTelCli 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5040
         MaxLength       =   15
         TabIndex        =   136
         ToolTipText     =   "Telefone do cliente"
         Top             =   960
         Width           =   2415
      End
      Begin VB.TextBox TxtNomeCli 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         MaxLength       =   200
         TabIndex        =   133
         ToolTipText     =   "Nome do cliente"
         Top             =   480
         Width           =   2415
      End
      Begin VB.ComboBox CboSexoCli 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "FrmPrincipal.frx":1625
         Left            =   1200
         List            =   "FrmPrincipal.frx":1627
         Style           =   2  'Dropdown List
         TabIndex        =   137
         ToolTipText     =   "Sexo do cliente"
         Top             =   1440
         Width           =   2415
      End
      Begin VB.TextBox TxtCpfCli 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         MaxLength       =   14
         TabIndex        =   135
         ToolTipText     =   "Cpf do cliente"
         Top             =   960
         Width           =   2415
      End
      Begin VB.Frame FraBotaoCli 
         Height          =   735
         Left            =   120
         TabIndex        =   198
         Top             =   5520
         Width           =   10695
         Begin VB.CommandButton CmdIncluirCli 
            Caption         =   "&Incluir"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5400
            TabIndex        =   141
            ToolTipText     =   "Incluir cliente"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton CmdAlterarCli 
            Caption         =   "&Alterar"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6720
            TabIndex        =   142
            ToolTipText     =   "Alterar cliente"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton CmdExcluirCli 
            Caption         =   "&Excluir"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   8040
            TabIndex        =   143
            ToolTipText     =   "Excluir cliente"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton CmdImprimirCli 
            Caption         =   "I&mprimir"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   9360
            TabIndex        =   144
            ToolTipText     =   "Imprimir consulta de clientes"
            Top             =   240
            Width           =   1095
         End
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel10 
         Height          =   255
         Left            =   480
         OleObjectBlob   =   "FrmPrincipal.frx":1629
         TabIndex        =   199
         Top             =   480
         Width           =   4695
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel12 
         Height          =   255
         Left            =   480
         OleObjectBlob   =   "FrmPrincipal.frx":16FF
         TabIndex        =   200
         Top             =   960
         Width           =   4695
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel14 
         Height          =   255
         Left            =   480
         OleObjectBlob   =   "FrmPrincipal.frx":17DD
         TabIndex        =   201
         Top             =   1440
         Width           =   495
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblNumTotalCli 
         Height          =   255
         Left            =   7920
         OleObjectBlob   =   "FrmPrincipal.frx":183F
         TabIndex        =   202
         Top             =   1680
         Width           =   2655
      End
      Begin FPSpread.vaSpread GridCliente 
         Height          =   3375
         Left            =   240
         TabIndex        =   140
         Top             =   2040
         Width           =   10455
         _Version        =   393216
         _ExtentX        =   18441
         _ExtentY        =   5953
         _StockProps     =   64
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   14737632
         GridColor       =   8421504
         MaxCols         =   16
         MaxRows         =   0
         OperationMode   =   2
         RetainSelBlock  =   0   'False
         RowHeaderDisplay=   0
         ScrollBarExtMode=   -1  'True
         ScrollBarMaxAlign=   0   'False
         ScrollBarShowMax=   0   'False
         SelectBlockOptions=   0
         ShadowColor     =   14737632
         ShadowDark      =   8421504
         ShadowText      =   0
         SpreadDesigner  =   "FrmPrincipal.frx":18CB
      End
   End
   Begin VB.Frame FraFornecedor 
      Caption         =   "Consulta de Fornecedores"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6375
      Left            =   360
      TabIndex        =   191
      Top             =   1920
      Width           =   10935
      Begin VB.CommandButton CmdLimpForn 
         Caption         =   "&Limpar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9240
         TabIndex        =   127
         ToolTipText     =   "Limpar pesquisa"
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CommandButton CmdPesqForn 
         Caption         =   "&Pesquisar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9240
         TabIndex        =   126
         ToolTipText     =   "Pesquisar fornecedores"
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox TxtNomeForn 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         MaxLength       =   200
         TabIndex        =   122
         ToolTipText     =   "Nome do fornecedor"
         Top             =   600
         Width           =   3015
      End
      Begin VB.TextBox TxtTelForn 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6360
         MaxLength       =   15
         TabIndex        =   123
         ToolTipText     =   "Telefone do fornecedor"
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox TxtTipoForn 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6360
         MaxLength       =   100
         TabIndex        =   125
         ToolTipText     =   "Tipo de produto"
         Top             =   1200
         Width           =   1935
      End
      Begin VB.TextBox TxtCnpjForn 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         MaxLength       =   18
         TabIndex        =   124
         ToolTipText     =   "CNPJ do fornecedor"
         Top             =   1200
         Width           =   3015
      End
      Begin VB.Frame FraBotaoForn 
         Height          =   735
         Left            =   120
         TabIndex        =   192
         Top             =   5520
         Width           =   10695
         Begin VB.CommandButton CmdImprimirForn 
            Caption         =   "I&mprimir"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   9360
            TabIndex        =   132
            ToolTipText     =   "Imprimir consulta de fornecedores"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton CmdIncluirForn 
            Caption         =   "&Incluir"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5400
            TabIndex        =   129
            ToolTipText     =   "Incluir fornecedor"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton CmdAlterarForn 
            Caption         =   "&Alterar"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6720
            TabIndex        =   130
            ToolTipText     =   "Alterar fornecedor"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton CmdExcluirForn 
            Caption         =   "&Excluir"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   8040
            TabIndex        =   131
            ToolTipText     =   "Excluir fornecedor"
            Top             =   240
            Width           =   1095
         End
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel60 
         Height          =   255
         Left            =   480
         OleObjectBlob   =   "FrmPrincipal.frx":214D
         TabIndex        =   193
         Top             =   600
         Width           =   5895
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel62 
         Height          =   255
         Left            =   480
         OleObjectBlob   =   "FrmPrincipal.frx":2243
         TabIndex        =   194
         Top             =   1200
         Width           =   5895
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblNumTotalForn 
         Height          =   255
         Left            =   7680
         OleObjectBlob   =   "FrmPrincipal.frx":2347
         TabIndex        =   195
         Top             =   1680
         Width           =   2895
      End
      Begin FPSpread.vaSpread GridFornecedor 
         Height          =   3375
         Left            =   240
         TabIndex        =   128
         Top             =   2040
         Width           =   10455
         _Version        =   393216
         _ExtentX        =   18441
         _ExtentY        =   5953
         _StockProps     =   64
         AllowDragDrop   =   -1  'True
         AllowMultiBlocks=   -1  'True
         AllowUserFormulas=   -1  'True
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   14737632
         GridColor       =   8421504
         MaxCols         =   15
         MaxRows         =   0
         OperationMode   =   2
         Protect         =   0   'False
         RetainSelBlock  =   0   'False
         RowHeaderDisplay=   0
         ScrollBarExtMode=   -1  'True
         ScrollBarMaxAlign=   0   'False
         ScrollBarShowMax=   0   'False
         SelectBlockOptions=   0
         ShadowColor     =   14737632
         ShadowDark      =   8421504
         ShadowText      =   0
         SpreadDesigner  =   "FrmPrincipal.frx":23D9
         ClipboardOptions=   14
      End
   End
   Begin VB.Frame FraProduto 
      Caption         =   "Consulta de Produto"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6375
      Left            =   360
      TabIndex        =   186
      Top             =   1920
      Width           =   10935
      Begin VB.CommandButton CmdLimparProd 
         Caption         =   "&Limpar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9240
         TabIndex        =   116
         ToolTipText     =   "Limpar pesquisa"
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox TxtPrVendaAtacProd 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7200
         MaxLength       =   10
         TabIndex        =   113
         ToolTipText     =   "Preço de venda por atacado"
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox TxtPrVendaUnitProd 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7200
         MaxLength       =   10
         TabIndex        =   111
         ToolTipText     =   "Preço de venda unitário"
         Top             =   480
         Width           =   1335
      End
      Begin VB.TextBox TxtNomeProd 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         MaxLength       =   100
         TabIndex        =   110
         ToolTipText     =   "Nome do produto"
         Top             =   480
         Width           =   2775
      End
      Begin VB.TextBox TxtTipoProd 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         MaxLength       =   50
         TabIndex        =   114
         ToolTipText     =   "Tipo de produto"
         Top             =   1440
         Width           =   2775
      End
      Begin VB.TextBox TxtFornProd 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         MaxLength       =   200
         TabIndex        =   112
         ToolTipText     =   "Fornecedor do produto"
         Top             =   960
         Width           =   2775
      End
      Begin VB.CommandButton CmdPesqProd 
         Caption         =   "&Pesquisar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9240
         TabIndex        =   115
         ToolTipText     =   "Pesquisar produtos"
         Top             =   360
         Width           =   1335
      End
      Begin VB.Frame Frame2 
         Height          =   735
         Left            =   120
         TabIndex        =   187
         Top             =   5520
         Width           =   10695
         Begin VB.CommandButton CmdImprimirProd 
            Caption         =   "I&mprimir"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   9360
            TabIndex        =   121
            ToolTipText     =   "Imprimir consulta de produtos"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton CmdIncluirProd 
            Caption         =   "&Incluir"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5400
            TabIndex        =   118
            ToolTipText     =   "Incluir produto"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton CmdAlterarProd 
            Caption         =   "&Alterar"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6720
            TabIndex        =   119
            ToolTipText     =   "Alterar produto"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton CmdExcluirProd 
            Caption         =   "&Excluir"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   8040
            TabIndex        =   120
            ToolTipText     =   "Excluir produto"
            Top             =   240
            Width           =   1095
         End
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel54 
         Height          =   255
         Left            =   480
         OleObjectBlob   =   "FrmPrincipal.frx":2BF7
         TabIndex        =   188
         Top             =   960
         Width           =   6855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel56 
         Height          =   255
         Left            =   480
         OleObjectBlob   =   "FrmPrincipal.frx":2D09
         TabIndex        =   189
         Top             =   1440
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblNumTotalProd 
         Height          =   255
         Left            =   7440
         OleObjectBlob   =   "FrmPrincipal.frx":2D7B
         TabIndex        =   190
         Top             =   1680
         Width           =   3135
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel45 
         Height          =   255
         Left            =   480
         OleObjectBlob   =   "FrmPrincipal.frx":2E07
         TabIndex        =   203
         Top             =   480
         Width           =   6855
      End
      Begin FPSpread.vaSpread GridProduto 
         Height          =   3375
         Left            =   240
         TabIndex        =   117
         Top             =   2040
         Width           =   10455
         _Version        =   393216
         _ExtentX        =   18441
         _ExtentY        =   5953
         _StockProps     =   64
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   14737632
         GridColor       =   8421504
         MaxCols         =   9
         MaxRows         =   0
         OperationMode   =   2
         RetainSelBlock  =   0   'False
         RowHeaderDisplay=   0
         ScrollBarExtMode=   -1  'True
         ScrollBarMaxAlign=   0   'False
         ScrollBarShowMax=   0   'False
         SelectBlockOptions=   0
         ShadowColor     =   14737632
         ShadowDark      =   8421504
         ShadowText      =   0
         SpreadDesigner  =   "FrmPrincipal.frx":2F1D
      End
   End
   Begin VB.Frame FraEstoque 
      Caption         =   "Consulta de Estoque"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6375
      Left            =   360
      TabIndex        =   155
      Top             =   1920
      Width           =   10935
      Begin VB.CommandButton CmdLimparEst 
         Caption         =   "&Limpar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9240
         TabIndex        =   104
         ToolTipText     =   "Limpar pesquisa"
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox TxtNomeProdEst 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2040
         MaxLength       =   100
         TabIndex        =   99
         ToolTipText     =   "Nome do produto"
         Top             =   600
         Width           =   2655
      End
      Begin VB.CommandButton CmdPesqEst 
         Caption         =   "&Pesquisar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9240
         TabIndex        =   103
         ToolTipText     =   "Pesquisar estoque"
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox TxtTipoProdEst 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2040
         MaxLength       =   50
         TabIndex        =   101
         ToolTipText     =   "Tipo do produto"
         Top             =   1200
         Width           =   2655
      End
      Begin VB.TextBox TxtQtdeEst 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6960
         MaxLength       =   5
         TabIndex        =   102
         ToolTipText     =   "Quantidade do produto em estoque"
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox TxtQtdeMin 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6960
         MaxLength       =   5
         TabIndex        =   100
         ToolTipText     =   "Quantidade mínima recomendada do produto"
         Top             =   600
         Width           =   855
      End
      Begin VB.Frame FraBotaoEst 
         Height          =   735
         Left            =   120
         TabIndex        =   160
         Top             =   5520
         Width           =   10695
         Begin VB.CommandButton CmdImprimirEst 
            Caption         =   "I&mprimir"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   9360
            TabIndex        =   108
            ToolTipText     =   "Imprimir consulta de estoque"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton CmdIncluirAlterarEst 
            Caption         =   "&Incluir/Alterar"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6000
            TabIndex        =   106
            ToolTipText     =   "Incluir/Alterar estoque"
            Top             =   240
            Width           =   1815
         End
         Begin VB.CommandButton CmdExcluirEst 
            Caption         =   "&Excluir"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   8040
            TabIndex        =   107
            ToolTipText     =   "Excluir estoque"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CheckBox ChkDesatAlerta 
            Caption         =   "Desativar alerta"
            Height          =   195
            Left            =   240
            TabIndex        =   109
            ToolTipText     =   "Desativar alerta de estoque"
            Top             =   360
            Width           =   1575
         End
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel58 
         Height          =   255
         Left            =   480
         OleObjectBlob   =   "FrmPrincipal.frx":35EE
         TabIndex        =   161
         Top             =   600
         Width           =   6615
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel59 
         Height          =   255
         Left            =   480
         OleObjectBlob   =   "FrmPrincipal.frx":36F6
         TabIndex        =   162
         Top             =   1200
         Width           =   6615
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblNumTotalEst 
         Height          =   255
         Left            =   7440
         OleObjectBlob   =   "FrmPrincipal.frx":3800
         TabIndex        =   163
         Top             =   1680
         Width           =   3135
      End
      Begin FPSpread.vaSpread GridEstoque 
         Height          =   3375
         Left            =   240
         TabIndex        =   105
         Top             =   2040
         Width           =   10455
         _Version        =   393216
         _ExtentX        =   18441
         _ExtentY        =   5953
         _StockProps     =   64
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   14737632
         GridColor       =   8421504
         MaxCols         =   6
         MaxRows         =   0
         OperationMode   =   2
         RetainSelBlock  =   0   'False
         RowHeaderDisplay=   0
         ScrollBarExtMode=   -1  'True
         ScrollBarMaxAlign=   0   'False
         ScrollBarShowMax=   0   'False
         SelectBlockOptions=   0
         ShadowColor     =   14737632
         ShadowDark      =   8421504
         ShadowText      =   0
         SpreadDesigner  =   "FrmPrincipal.frx":3894
      End
   End
   Begin VB.Frame FraCrediario 
      Caption         =   "Consulta de Crediário"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6375
      Left            =   360
      TabIndex        =   220
      Top             =   1920
      Width           =   10935
      Begin VB.Frame FraBotaoCred 
         Height          =   735
         Left            =   120
         TabIndex        =   234
         Top             =   5520
         Width           =   10695
         Begin VB.CommandButton CmdVerParc 
            Caption         =   "&Parcelas"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   7800
            TabIndex        =   97
            ToolTipText     =   "Visualizar parcela do crediário"
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton CmdExcluirCred 
            Caption         =   "&Excluir crediário"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5640
            TabIndex        =   96
            ToolTipText     =   "Excluir crediário"
            Top             =   240
            Width           =   1935
         End
         Begin VB.CommandButton CmdImprimirCred 
            Caption         =   "I&mprimir"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   9240
            TabIndex        =   98
            ToolTipText     =   "Imprimir consulta crediário"
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.CommandButton CmdLimparCred 
         Caption         =   "&Limpar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9240
         TabIndex        =   94
         ToolTipText     =   "Limpar pesquisa"
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox TxtCliCred 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         MaxLength       =   200
         TabIndex        =   85
         ToolTipText     =   "Nome do cliente"
         Top             =   600
         Width           =   2535
      End
      Begin VB.TextBox TxtCredstaCred 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1800
         MaxLength       =   200
         TabIndex        =   88
         ToolTipText     =   "Nome do crediarista"
         Top             =   960
         Width           =   2535
      End
      Begin VB.CommandButton CmdPesqCred 
         Caption         =   "&Pesquisar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9240
         TabIndex        =   93
         ToolTipText     =   "Pesquisar crediários"
         Top             =   360
         Width           =   1335
      End
      Begin VB.ComboBox CboTipoCred 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "FrmPrincipal.frx":3E55
         Left            =   1800
         List            =   "FrmPrincipal.frx":3E57
         Style           =   2  'Dropdown List
         TabIndex        =   91
         ToolTipText     =   "Tipo de crediário"
         Top             =   1320
         Width           =   2535
      End
      Begin VB.TextBox TxtDtCred1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5760
         MaxLength       =   10
         TabIndex        =   86
         Text            =   "__/__/____"
         ToolTipText     =   "Menor data do crediário"
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox TxtDtCred2 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7320
         MaxLength       =   10
         TabIndex        =   87
         Text            =   "__/__/____"
         ToolTipText     =   "Maior data do crediário"
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox TxtDtVencCred1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5760
         MaxLength       =   10
         TabIndex        =   89
         Text            =   "__/__/____"
         ToolTipText     =   "Menor data do vencimento"
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox TxtDtVencCred2 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7320
         MaxLength       =   10
         TabIndex        =   90
         Text            =   "__/__/____"
         ToolTipText     =   "Maior data do vencimento"
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox TxtCodParcCred 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5760
         MaxLength       =   5
         TabIndex        =   92
         ToolTipText     =   "Código da parcela do crediário"
         Top             =   1320
         Width           =   1215
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel47 
         Height          =   255
         Left            =   480
         OleObjectBlob   =   "FrmPrincipal.frx":3E59
         TabIndex        =   230
         Top             =   600
         Width           =   6855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel3 
         Height          =   255
         Left            =   480
         OleObjectBlob   =   "FrmPrincipal.frx":3F7F
         TabIndex        =   231
         Top             =   960
         Width           =   6855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel49 
         Height          =   255
         Left            =   480
         OleObjectBlob   =   "FrmPrincipal.frx":409D
         TabIndex        =   232
         Top             =   1320
         Width           =   5295
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblNumTotalCred 
         Height          =   255
         Left            =   7440
         OleObjectBlob   =   "FrmPrincipal.frx":4187
         TabIndex        =   233
         Top             =   1680
         Width           =   3135
      End
      Begin FPSpread.vaSpread GridCrediario 
         Height          =   3375
         Left            =   240
         TabIndex        =   95
         Top             =   2040
         Width           =   10455
         _Version        =   393216
         _ExtentX        =   18441
         _ExtentY        =   5953
         _StockProps     =   64
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   14737632
         GridColor       =   8421504
         MaxCols         =   12
         MaxRows         =   0
         OperationMode   =   2
         RetainSelBlock  =   0   'False
         RowHeaderDisplay=   0
         ScrollBarExtMode=   -1  'True
         ScrollBarMaxAlign=   0   'False
         ScrollBarShowMax=   0   'False
         SelectBlockOptions=   0
         ShadowColor     =   14737632
         ShadowDark      =   8421504
         SpreadDesigner  =   "FrmPrincipal.frx":4217
      End
   End
   Begin VB.Frame FraCrediarista 
      Caption         =   "Consulta de Crediarista"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6375
      Left            =   360
      TabIndex        =   235
      Top             =   1920
      Width           =   10935
      Begin VB.Frame Frame5 
         Height          =   735
         Left            =   120
         TabIndex        =   236
         Top             =   5520
         Width           =   10695
         Begin VB.CommandButton CmdImprimirCredsta 
            Caption         =   "I&mprimir"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   9360
            TabIndex        =   84
            ToolTipText     =   "Imprimir consulta de crediarista"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton CmdExcluirCredsta 
            Caption         =   "&Excluir"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   8040
            TabIndex        =   83
            ToolTipText     =   "Excluir crediarista"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton CmdAlterarCredsta 
            Caption         =   "&Alterar"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6720
            TabIndex        =   82
            ToolTipText     =   "Alterar crediarista"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton CmdIncluirCredsta 
            Caption         =   "&Incluir"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5400
            TabIndex        =   81
            ToolTipText     =   "Incluir crediarista"
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.TextBox TxtCpfCredsta 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         MaxLength       =   14
         TabIndex        =   76
         ToolTipText     =   "Cpf do crediarista"
         Top             =   1200
         Width           =   2415
      End
      Begin VB.TextBox TxtNomeCredsta 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         MaxLength       =   200
         TabIndex        =   74
         ToolTipText     =   "Nome do crediarista"
         Top             =   600
         Width           =   2415
      End
      Begin VB.TextBox TxtTelCredsta 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5040
         MaxLength       =   15
         TabIndex        =   77
         ToolTipText     =   "Telefone do crediarista"
         Top             =   1200
         Width           =   2415
      End
      Begin VB.TextBox TxtBairroCredsta 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5040
         MaxLength       =   60
         TabIndex        =   75
         ToolTipText     =   "Bairro do crediarista"
         Top             =   600
         Width           =   2415
      End
      Begin VB.CommandButton CmdPesqCredsta 
         Caption         =   "&Pesquisar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9240
         TabIndex        =   78
         ToolTipText     =   "Pesquisar crediaristas"
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton CmdLimparCredsta 
         Caption         =   "&Limpar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9240
         TabIndex        =   79
         ToolTipText     =   "Limpar pesquisa"
         Top             =   1080
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
         Height          =   255
         Left            =   480
         OleObjectBlob   =   "FrmPrincipal.frx":49BA
         TabIndex        =   237
         Top             =   600
         Width           =   4695
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel5 
         Height          =   255
         Left            =   480
         OleObjectBlob   =   "FrmPrincipal.frx":4A90
         TabIndex        =   238
         Top             =   1200
         Width           =   4695
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblNumTotalCredsta 
         Height          =   255
         Left            =   7200
         OleObjectBlob   =   "FrmPrincipal.frx":4B6E
         TabIndex        =   239
         Top             =   1680
         Width           =   3375
      End
      Begin FPSpread.vaSpread GridCredsta 
         Height          =   3375
         Left            =   240
         TabIndex        =   80
         Top             =   2040
         Width           =   10455
         _Version        =   393216
         _ExtentX        =   18441
         _ExtentY        =   5953
         _StockProps     =   64
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   14737632
         GridColor       =   8421504
         MaxCols         =   12
         MaxRows         =   0
         OperationMode   =   2
         RetainSelBlock  =   0   'False
         RowHeaderDisplay=   0
         ScrollBarExtMode=   -1  'True
         ScrollBarMaxAlign=   0   'False
         ScrollBarShowMax=   0   'False
         SelectBlockOptions=   0
         ShadowColor     =   14737632
         ShadowDark      =   8421504
         ShadowText      =   0
         SpreadDesigner  =   "FrmPrincipal.frx":4C02
      End
   End
   Begin VB.Frame FraCaixa 
      Caption         =   "Consulta de Movimento de Caixa"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6375
      Left            =   360
      TabIndex        =   183
      Top             =   1920
      Width           =   10935
      Begin VB.CommandButton CmdLimparCx 
         Caption         =   "&Limpar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9240
         TabIndex        =   68
         ToolTipText     =   "Limpar pesquisa"
         Top             =   1080
         Width           =   1335
      End
      Begin VB.ComboBox CboTipoPagtoCx 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2400
         Style           =   2  'Dropdown List
         TabIndex        =   66
         ToolTipText     =   "Tipo de pagamento"
         Top             =   1200
         Width           =   3135
      End
      Begin VB.TextBox TxtDtMovCx2 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   4200
         MaxLength       =   10
         TabIndex        =   65
         Text            =   "__/__/____"
         ToolTipText     =   "Maior data do movimento de caixa"
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox TxtDtMovCx1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2400
         MaxLength       =   10
         TabIndex        =   64
         Text            =   "__/__/____"
         ToolTipText     =   "Menor data do movimento de caixa"
         Top             =   600
         Width           =   1335
      End
      Begin VB.CommandButton CmdPesqCx 
         Caption         =   "&Pesquisar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9240
         TabIndex        =   67
         ToolTipText     =   "Pesquisar movimento de caixa"
         Top             =   360
         Width           =   1335
      End
      Begin VB.Frame FraBotaoCx 
         Height          =   735
         Left            =   120
         TabIndex        =   184
         Top             =   5520
         Width           =   10695
         Begin VB.CommandButton CmdImprimirCx 
            Caption         =   "I&mprimir"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   9360
            TabIndex        =   73
            ToolTipText     =   "Imprimir consulta de movimento de caixa"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton CmdIncluirCx 
            Caption         =   "&Incluir"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5400
            TabIndex        =   70
            ToolTipText     =   "Incluir movimento de caixa"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton CmdAlterarCx 
            Caption         =   "&Alterar"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6720
            TabIndex        =   71
            ToolTipText     =   "Alterar movimento de caixa"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton CmdExcluirCx 
            Caption         =   "&Excluir"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   8040
            TabIndex        =   72
            ToolTipText     =   "Excluir movimento de caixa"
            Top             =   240
            Width           =   1095
         End
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel46 
         Height          =   255
         Left            =   480
         OleObjectBlob   =   "FrmPrincipal.frx":536C
         TabIndex        =   185
         Top             =   600
         Width           =   3855
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel64 
         Height          =   255
         Left            =   480
         OleObjectBlob   =   "FrmPrincipal.frx":5420
         TabIndex        =   204
         Top             =   1200
         Width           =   1815
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblNumTotalCx 
         Height          =   255
         Left            =   6960
         OleObjectBlob   =   "FrmPrincipal.frx":549C
         TabIndex        =   205
         Top             =   1680
         Width           =   3735
      End
      Begin FPSpread.vaSpread GridCaixa 
         Height          =   3375
         Left            =   240
         TabIndex        =   69
         Top             =   2040
         Width           =   10455
         _Version        =   393216
         _ExtentX        =   18441
         _ExtentY        =   5953
         _StockProps     =   64
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   14737632
         GridColor       =   8421504
         MaxCols         =   7
         MaxRows         =   0
         OperationMode   =   2
         RetainSelBlock  =   0   'False
         RowHeaderDisplay=   0
         ScrollBarExtMode=   -1  'True
         ScrollBarMaxAlign=   0   'False
         ScrollBarShowMax=   0   'False
         SelectBlockOptions=   0
         ShadowColor     =   14737632
         ShadowDark      =   8421504
         ShadowText      =   0
         SpreadDesigner  =   "FrmPrincipal.frx":553E
      End
   End
   Begin VB.Frame FraPagar 
      Caption         =   "Consulta de Contas a Pagar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6375
      Left            =   360
      TabIndex        =   215
      Top             =   1920
      Width           =   10935
      Begin VB.Frame FraBotaoRec 
         Height          =   735
         Left            =   120
         TabIndex        =   224
         Top             =   5520
         Width           =   10695
         Begin VB.CommandButton CmdImprimirAPagar 
            Caption         =   "I&mprimir"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   8040
            TabIndex        =   62
            ToolTipText     =   "Imprimir consulta de contas a pagar"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton CmdIncluirAPagar 
            Caption         =   "&Incluir"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4080
            TabIndex        =   59
            ToolTipText     =   "Incluir contas a pagar"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton CmdAlterarAPagar 
            Caption         =   "&Alterar"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5400
            TabIndex        =   60
            ToolTipText     =   "Alterar contas a pagar"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton CmdExcluirAPagar 
            Caption         =   "&Excluir"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6720
            TabIndex        =   61
            ToolTipText     =   "Excluir contas a pagar"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton CmdBaixarAPagar 
            Caption         =   "&Baixar"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   9360
            TabIndex        =   63
            ToolTipText     =   "Baixar contas a pagar"
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.TextBox TxtDtAPagar1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   50
         Text            =   "__/__/____"
         ToolTipText     =   "Menor data de vencimento"
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox TxtDtAPagar2 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3240
         MaxLength       =   10
         TabIndex        =   51
         Text            =   "__/__/____"
         ToolTipText     =   "Maior data de vencimento"
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox TxtDescrAPagar 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         MaxLength       =   255
         TabIndex        =   52
         ToolTipText     =   "Descrição da conta a pagar"
         Top             =   1200
         Width           =   2775
      End
      Begin VB.OptionButton OptPagoSimAPagar 
         Caption         =   "Pago"
         Height          =   195
         Left            =   6240
         TabIndex        =   54
         ToolTipText     =   "Contas pagas"
         Top             =   840
         Width           =   1215
      End
      Begin VB.OptionButton OptPagoNaoAPagar 
         Caption         =   "Em aberto"
         Height          =   195
         Left            =   6240
         TabIndex        =   53
         ToolTipText     =   "Contas em aberto"
         Top             =   480
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton OptPagoTodosAPagar 
         Caption         =   "Todos"
         Height          =   195
         Left            =   6240
         TabIndex        =   55
         ToolTipText     =   "Todas as contas"
         Top             =   1200
         Width           =   975
      End
      Begin VB.CommandButton CmdLimparAPagar 
         Caption         =   "&Limpar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9240
         TabIndex        =   57
         ToolTipText     =   "Limpar pesquisa"
         Top             =   1080
         Width           =   1335
      End
      Begin VB.CommandButton CmdPesqAPagar 
         Caption         =   "&Pesquisar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9240
         TabIndex        =   56
         ToolTipText     =   "Pesquisar contas a pagar"
         Top             =   360
         Width           =   1335
      End
      Begin FPSpread.vaSpread GridAPagar 
         Height          =   3375
         Left            =   240
         TabIndex        =   58
         Top             =   2040
         Width           =   10455
         _Version        =   393216
         _ExtentX        =   18441
         _ExtentY        =   5953
         _StockProps     =   64
         ColHeaderDisplay=   0
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   14737632
         GridColor       =   8421504
         MaxCols         =   6
         MaxRows         =   0
         OperationMode   =   2
         RetainSelBlock  =   0   'False
         RowHeaderDisplay=   0
         ScrollBarExtMode=   -1  'True
         SelectBlockOptions=   0
         ShadowColor     =   14737632
         ShadowDark      =   8421504
         ShadowText      =   0
         SpreadDesigner  =   "FrmPrincipal.frx":5C00
         UserResize      =   1
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblNumTotalPag 
         Height          =   255
         Left            =   7560
         OleObjectBlob   =   "FrmPrincipal.frx":6077
         TabIndex        =   221
         Top             =   1680
         Width           =   3015
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel24 
         Height          =   255
         Left            =   480
         OleObjectBlob   =   "FrmPrincipal.frx":6107
         TabIndex        =   222
         Top             =   600
         Width           =   2895
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel26 
         Height          =   255
         Left            =   480
         OleObjectBlob   =   "FrmPrincipal.frx":61A7
         TabIndex        =   223
         Top             =   1200
         Width           =   1095
      End
   End
   Begin VB.Frame FraReceber 
      Caption         =   "Consulta de Contas a Receber"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6375
      Left            =   360
      TabIndex        =   214
      Top             =   1920
      Width           =   10935
      Begin VB.Frame Frame4 
         Height          =   735
         Left            =   120
         TabIndex        =   228
         Top             =   5520
         Width           =   10695
         Begin VB.CommandButton CmdBaixaAReceber 
            Caption         =   "&Baixar"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   9360
            TabIndex        =   49
            ToolTipText     =   "Baixar contas a receber"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton CmdExcluirAReceber 
            Caption         =   "&Excluir"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6720
            TabIndex        =   47
            ToolTipText     =   "Excluir contas a receber"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton CmdAlterarAReceber 
            Caption         =   "&Alterar"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5400
            TabIndex        =   46
            ToolTipText     =   "Alterar contas a receber"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton CmdIncluirAReceber 
            Caption         =   "&Incluir"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   4080
            TabIndex        =   45
            ToolTipText     =   "Incluir contas a receber"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton CmdImprimirAReceber 
            Caption         =   "I&mprimir"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   8040
            TabIndex        =   48
            ToolTipText     =   "Imprimir consulta de contas a receber"
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.CommandButton CmdLimparAReceber 
         Caption         =   "&Limpar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9240
         TabIndex        =   43
         ToolTipText     =   "Limpar pesquisa"
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox TxtDtReceb1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         MaxLength       =   10
         TabIndex        =   36
         Text            =   "__/__/____"
         ToolTipText     =   "Menor data de vencimento"
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox TxtDtReceb2 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3240
         MaxLength       =   10
         TabIndex        =   37
         Text            =   "__/__/____"
         ToolTipText     =   "Maior data de vencimento"
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton OptRecebSim 
         Caption         =   "Recebido"
         Height          =   195
         Left            =   6240
         TabIndex        =   40
         ToolTipText     =   "Contas recebidas"
         Top             =   840
         Width           =   1215
      End
      Begin VB.OptionButton OptRecebNao 
         Caption         =   "A receber"
         Height          =   195
         Left            =   6240
         TabIndex        =   39
         ToolTipText     =   "Contas a receber"
         Top             =   480
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.OptionButton OptRecebTodos 
         Caption         =   "Todos"
         Height          =   195
         Left            =   6240
         TabIndex        =   41
         ToolTipText     =   "Todas as contas"
         Top             =   1200
         Width           =   975
      End
      Begin VB.TextBox TxtCliReceb 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1680
         MaxLength       =   200
         TabIndex        =   38
         ToolTipText     =   "Nome do cliente"
         Top             =   1200
         Width           =   2775
      End
      Begin VB.CommandButton CmdPesqAReceber 
         Caption         =   "&Pesquisar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9240
         TabIndex        =   42
         ToolTipText     =   "Pesquisar contas a receber"
         Top             =   360
         Width           =   1335
      End
      Begin FPSpread.vaSpread GridAReceber 
         Height          =   3375
         Left            =   240
         TabIndex        =   44
         Top             =   2040
         Width           =   10455
         _Version        =   393216
         _ExtentX        =   18441
         _ExtentY        =   5953
         _StockProps     =   64
         ColHeaderDisplay=   0
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   14737632
         GridColor       =   8421504
         MaxCols         =   7
         MaxRows         =   0
         OperationMode   =   2
         RetainSelBlock  =   0   'False
         RowHeaderDisplay=   0
         ScrollBarExtMode=   -1  'True
         SelectBlockOptions=   0
         ShadowColor     =   14737632
         ShadowDark      =   8421504
         ShadowText      =   0
         SpreadDesigner  =   "FrmPrincipal.frx":6213
         UserResize      =   1
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblNumTotalReceb 
         Height          =   255
         Left            =   7560
         OleObjectBlob   =   "FrmPrincipal.frx":66E8
         TabIndex        =   225
         Top             =   1680
         Width           =   3015
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel7 
         Height          =   255
         Left            =   480
         OleObjectBlob   =   "FrmPrincipal.frx":677C
         TabIndex        =   226
         Top             =   600
         Width           =   2895
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel9 
         Height          =   255
         Left            =   480
         OleObjectBlob   =   "FrmPrincipal.frx":681C
         TabIndex        =   227
         Top             =   1200
         Width           =   1095
      End
   End
   Begin VB.Frame FraOrcamento 
      Caption         =   "Consulta de Orçamentos"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6375
      Left            =   360
      TabIndex        =   168
      Top             =   1920
      Width           =   10935
      Begin VB.CommandButton CmdLimparOrc 
         Caption         =   "&Limpar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9240
         TabIndex        =   30
         ToolTipText     =   "Limpar pesquisa"
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox TxtDtOrc2 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   7200
         MaxLength       =   10
         TabIndex        =   28
         Text            =   "__/__/____"
         ToolTipText     =   "Maior data do orçamento"
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox TxtDtOrc1 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5640
         MaxLength       =   10
         TabIndex        =   27
         Text            =   "__/__/____"
         ToolTipText     =   "Menor data do orçamento"
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox TxtVendOrc 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   5640
         MaxLength       =   200
         TabIndex        =   25
         ToolTipText     =   "Nome do vendedor"
         Top             =   600
         Width           =   2775
      End
      Begin VB.TextBox TxtTelOrc 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         MaxLength       =   15
         TabIndex        =   26
         ToolTipText     =   "Telefone do cliente"
         Top             =   1200
         Width           =   1335
      End
      Begin VB.CommandButton CmdPesqOrc 
         Caption         =   "&Pesquisar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9240
         TabIndex        =   29
         ToolTipText     =   "Pesquisar orçamentos"
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox TxtCliOrc 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1440
         MaxLength       =   200
         TabIndex        =   24
         ToolTipText     =   "Nome do cliente"
         Top             =   600
         Width           =   2775
      End
      Begin VB.Frame FraBotaoOrc 
         Height          =   735
         Left            =   120
         TabIndex        =   169
         Top             =   5520
         Width           =   10695
         Begin VB.CommandButton CmdImprimirOrc 
            Caption         =   "I&mprimir"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   9360
            TabIndex        =   35
            ToolTipText     =   "Imprimir consulta de orçamento"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton CmdIncluirOrc 
            Caption         =   "&Incluir"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5400
            TabIndex        =   32
            ToolTipText     =   "Incluir orçamento"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton CmdAlterarOrc 
            Caption         =   "&Alterar"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6720
            TabIndex        =   33
            ToolTipText     =   "Alterar orçamento"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton CmdExcluirOrc 
            Caption         =   "&Excluir"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   8040
            TabIndex        =   34
            ToolTipText     =   "Excluir orçamento"
            Top             =   240
            Width           =   1095
         End
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel19 
         Height          =   255
         Left            =   480
         OleObjectBlob   =   "FrmPrincipal.frx":6884
         TabIndex        =   170
         Top             =   600
         Width           =   5295
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel21 
         Height          =   255
         Left            =   480
         OleObjectBlob   =   "FrmPrincipal.frx":6972
         TabIndex        =   171
         Top             =   1200
         Width           =   6855
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblNumTotalOrc 
         Height          =   255
         Left            =   7680
         OleObjectBlob   =   "FrmPrincipal.frx":6A96
         TabIndex        =   206
         Top             =   1680
         Width           =   3015
      End
      Begin FPSpread.vaSpread GridOrcamento 
         Height          =   3375
         Left            =   240
         TabIndex        =   31
         Top             =   2040
         Width           =   10455
         _Version        =   393216
         _ExtentX        =   18441
         _ExtentY        =   5953
         _StockProps     =   64
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   14737632
         GridColor       =   8421504
         MaxCols         =   12
         MaxRows         =   0
         OperationMode   =   2
         RetainSelBlock  =   0   'False
         RowHeaderDisplay=   0
         ScrollBarExtMode=   -1  'True
         ScrollBarMaxAlign=   0   'False
         ScrollBarShowMax=   0   'False
         SelectBlockOptions=   0
         ShadowColor     =   14737632
         ShadowDark      =   8421504
         ShadowText      =   0
         SpreadDesigner  =   "FrmPrincipal.frx":6B26
      End
   End
   Begin VB.Frame FraVendedor 
      Caption         =   "Consulta de Vendedores"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6375
      Left            =   360
      TabIndex        =   210
      Top             =   1920
      Width           =   10935
      Begin VB.CommandButton CmdLimparVendedor 
         Caption         =   "&Limpar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   9240
         TabIndex        =   18
         ToolTipText     =   "Limpar pesquisa"
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Frame Frame3 
         Height          =   735
         Left            =   120
         TabIndex        =   211
         Top             =   5520
         Width           =   10695
         Begin VB.CommandButton CmdImprimirVend 
            Caption         =   "I&mprimir"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   9360
            TabIndex        =   23
            ToolTipText     =   "Imprimir consulta de vendedor"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton CmdExcluirVend 
            Caption         =   "&Excluir"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   8040
            TabIndex        =   22
            ToolTipText     =   "Excluir vendedor"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton CmdAlterarVend 
            Caption         =   "&Alterar"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   6720
            TabIndex        =   21
            ToolTipText     =   "Alterar vendedor"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton CmdIncluirVend 
            Caption         =   "&Incluir"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5400
            TabIndex        =   20
            ToolTipText     =   "Incluir vendedor"
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.TextBox TxtNomeVend 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1200
         MaxLength       =   200
         TabIndex        =   16
         ToolTipText     =   "Nome do vendedor"
         Top             =   960
         Width           =   3495
      End
      Begin VB.CommandButton CmdPesqVendedor 
         Caption         =   "&Pesquisar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   9240
         TabIndex        =   17
         ToolTipText     =   "Pesquisar vendedores"
         Top             =   360
         Width           =   1335
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
         Height          =   255
         Left            =   480
         OleObjectBlob   =   "FrmPrincipal.frx":729D
         TabIndex        =   212
         Top             =   960
         Width           =   615
      End
      Begin ACTIVESKINLibCtl.SkinLabel LblNumTotalVendedor 
         Height          =   255
         Left            =   7920
         OleObjectBlob   =   "FrmPrincipal.frx":72FF
         TabIndex        =   213
         Top             =   1680
         Width           =   2655
      End
      Begin FPSpread.vaSpread GridVendedor 
         Height          =   3375
         Left            =   240
         TabIndex        =   19
         Top             =   2040
         Width           =   10455
         _Version        =   393216
         _ExtentX        =   18441
         _ExtentY        =   5953
         _StockProps     =   64
         DisplayRowHeaders=   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         GrayAreaBackColor=   14737632
         GridColor       =   8421504
         MaxCols         =   3
         MaxRows         =   0
         OperationMode   =   2
         RetainSelBlock  =   0   'False
         RowHeaderDisplay=   0
         ScrollBarExtMode=   -1  'True
         ScrollBarMaxAlign=   0   'False
         ScrollBarShowMax=   0   'False
         SelectBlockOptions=   0
         ShadowColor     =   14737632
         ShadowDark      =   8421504
         ShadowText      =   0
         SpreadDesigner  =   "FrmPrincipal.frx":738D
      End
   End
   Begin VB.Frame FraExtra 
      Caption         =   "Consulta de Opções Extras de Relatório"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6375
      Left            =   360
      TabIndex        =   172
      Top             =   1920
      Width           =   10935
      Begin VB.OptionButton OptCob 
         Caption         =   "Cartas de cobrança"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4440
         TabIndex        =   1
         ToolTipText     =   "Consulta de cartas de cobrança à clientes"
         Top             =   840
         Width           =   2055
      End
      Begin VB.OptionButton OptNiver 
         Caption         =   "Aniversariantes"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1920
         TabIndex        =   0
         ToolTipText     =   "Consulta de aniversariantes"
         Top             =   840
         Width           =   1695
      End
      Begin VB.OptionButton OptMala 
         Caption         =   "Mala direta"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7440
         TabIndex        =   2
         ToolTipText     =   "Consulta de mala direta"
         Top             =   840
         Width           =   1335
      End
      Begin VB.Frame FraNiver 
         Caption         =   "Aniversariantes"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   1440
         TabIndex        =   181
         Top             =   2040
         Visible         =   0   'False
         Width           =   8055
         Begin VB.TextBox TxtMes1 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4920
            MaxLength       =   2
            TabIndex        =   13
            ToolTipText     =   "Menor mês do aniversário"
            Top             =   480
            Width           =   495
         End
         Begin VB.TextBox TxtMes2 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5760
            MaxLength       =   2
            TabIndex        =   14
            ToolTipText     =   "Maior mês do aniversário"
            Top             =   480
            Width           =   495
         End
         Begin VB.TextBox TxtDia1 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1800
            MaxLength       =   2
            TabIndex        =   11
            ToolTipText     =   "Menor dia do aniversário"
            Top             =   480
            Width           =   495
         End
         Begin VB.TextBox TxtDia2 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2640
            MaxLength       =   2
            TabIndex        =   12
            ToolTipText     =   "Maior dia do aniversário"
            Top             =   480
            Width           =   495
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel30 
            Height          =   255
            Index           =   0
            Left            =   1320
            OleObjectBlob   =   "FrmPrincipal.frx":7868
            TabIndex        =   182
            Top             =   480
            Width           =   1455
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel29 
            Height          =   255
            Index           =   0
            Left            =   3960
            OleObjectBlob   =   "FrmPrincipal.frx":78E2
            TabIndex        =   196
            Top             =   1440
            Width           =   255
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel30 
            Height          =   255
            Index           =   2
            Left            =   4440
            OleObjectBlob   =   "FrmPrincipal.frx":793A
            TabIndex        =   209
            Top             =   480
            Width           =   1335
         End
      End
      Begin VB.Frame FraCob 
         Caption         =   "Cartas de cobrança"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   1440
         TabIndex        =   177
         Top             =   2040
         Visible         =   0   'False
         Width           =   8055
         Begin VB.TextBox TxtClienteCob 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1440
            MaxLength       =   200
            TabIndex        =   8
            ToolTipText     =   "Nome do cliente"
            Top             =   1200
            Width           =   6375
         End
         Begin VB.TextBox TxtDtVenc2 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3000
            MaxLength       =   10
            TabIndex        =   10
            Text            =   "__/__/____"
            ToolTipText     =   "Maior data de vencimento"
            Top             =   1920
            Width           =   1215
         End
         Begin VB.TextBox TxtDtVenc1 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1440
            MaxLength       =   10
            TabIndex        =   9
            Text            =   "__/__/____"
            ToolTipText     =   "Menor data de vencimento"
            Top             =   1920
            Width           =   1215
         End
         Begin VB.ComboBox CboTipoCarta 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "FrmPrincipal.frx":79B2
            Left            =   1440
            List            =   "FrmPrincipal.frx":79BF
            Style           =   2  'Dropdown List
            TabIndex        =   7
            ToolTipText     =   "Tipo de carta de cobrança"
            Top             =   480
            Width           =   6375
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel32 
            Height          =   255
            Left            =   240
            OleObjectBlob   =   "FrmPrincipal.frx":79FF
            TabIndex        =   178
            Top             =   480
            Width           =   735
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel33 
            Height          =   255
            Left            =   240
            OleObjectBlob   =   "FrmPrincipal.frx":7A61
            TabIndex        =   179
            Top             =   1200
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel34 
            Height          =   255
            Left            =   240
            OleObjectBlob   =   "FrmPrincipal.frx":7AC9
            TabIndex        =   180
            Top             =   1920
            Width           =   2775
         End
      End
      Begin VB.Frame FraMala 
         Caption         =   "Mala direta"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2415
         Left            =   1440
         TabIndex        =   174
         Top             =   2040
         Visible         =   0   'False
         Width           =   8055
         Begin VB.TextBox TxtDtNiverCli1 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1320
            MaxLength       =   10
            TabIndex        =   5
            Text            =   "__/__/____"
            ToolTipText     =   "Menor data de vencimento"
            Top             =   1080
            Width           =   1215
         End
         Begin VB.TextBox TxtDtNiverCli2 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2880
            MaxLength       =   10
            TabIndex        =   6
            Text            =   "__/__/____"
            ToolTipText     =   "Maior data de vencimento"
            Top             =   1080
            Width           =   1215
         End
         Begin VB.TextBox TxtCliente 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1320
            MaxLength       =   200
            TabIndex        =   3
            ToolTipText     =   "Nome do cliente"
            Top             =   480
            Width           =   2775
         End
         Begin VB.ComboBox CboSexo 
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            ItemData        =   "FrmPrincipal.frx":7B6B
            Left            =   5280
            List            =   "FrmPrincipal.frx":7B75
            Style           =   2  'Dropdown List
            TabIndex        =   4
            ToolTipText     =   "Sexo do cliente"
            Top             =   480
            Width           =   2535
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel36 
            Height          =   255
            Left            =   240
            OleObjectBlob   =   "FrmPrincipal.frx":7B8E
            TabIndex        =   175
            Top             =   480
            Width           =   735
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel37 
            Height          =   255
            Left            =   4680
            OleObjectBlob   =   "FrmPrincipal.frx":7BF6
            TabIndex        =   176
            Top             =   480
            Width           =   615
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel30 
            Height          =   255
            Index           =   1
            Left            =   240
            OleObjectBlob   =   "FrmPrincipal.frx":7C58
            TabIndex        =   208
            Top             =   1080
            Width           =   2655
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
            Height          =   255
            Left            =   240
            OleObjectBlob   =   "FrmPrincipal.frx":7CF8
            TabIndex        =   229
            Top             =   2040
            Width           =   3135
         End
      End
      Begin VB.Frame FraBotaoExt 
         Height          =   735
         Left            =   120
         TabIndex        =   173
         Top             =   5520
         Width           =   10695
         Begin VB.CommandButton CmdImprimirExt 
            Caption         =   "&Imprimir"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   9360
            TabIndex        =   15
            ToolTipText     =   "Imprimir consulta"
            Top             =   240
            Width           =   1095
         End
      End
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   2040
      OleObjectBlob   =   "FrmPrincipal.frx":7D94
      Top             =   240
   End
   Begin ACTIVESKINLibCtl.SkinLabel LblData 
      Height          =   255
      Left            =   240
      OleObjectBlob   =   "FrmPrincipal.frx":7FC8
      TabIndex        =   218
      Top             =   240
      Width           =   4815
   End
   Begin ACTIVESKINLibCtl.SkinLabel LblHora 
      Height          =   255
      Left            =   8760
      OleObjectBlob   =   "FrmPrincipal.frx":803C
      TabIndex        =   219
      Top             =   240
      Width           =   2655
   End
   Begin MSComctlLib.TabStrip TabPrincipal 
      Height          =   7335
      Left            =   240
      TabIndex        =   240
      Top             =   1080
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   12938
      MultiRow        =   -1  'True
      HotTracking     =   -1  'True
      TabMinWidth     =   1940
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   13
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "VENDA"
            Object.Tag             =   "1"
            Object.ToolTipText     =   "Controle de vendas"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "CLIENTE"
            Object.ToolTipText     =   "Controle de clientes"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "FORNECEDOR"
            Object.ToolTipText     =   "Controle de fornecedores"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "PRODUTO"
            Object.ToolTipText     =   "Controle de produtos"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "ESTOQUE"
            Object.ToolTipText     =   "Controle de estoque"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "CREDIÁRIO"
            Object.ToolTipText     =   "Controle de crediários"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "CREDIARISTA"
            Object.ToolTipText     =   "Controle de crediaristas"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "CAIXA"
            Object.ToolTipText     =   "Controle do movimento de caixa"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab9 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "A PAGAR"
            Object.ToolTipText     =   "Controle de contas a pagar"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab10 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "A RECEBER"
            Object.ToolTipText     =   "Controle de contas a receber"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab11 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "ORÇAMENTO"
            Object.ToolTipText     =   "Controle de orçamentos"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab12 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "VENDEDOR"
            Object.ToolTipText     =   "Controle de vendedores"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab13 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "EXTRA"
            Object.ToolTipText     =   "Opções extras de relatórios"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "FrmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public RecPesq As New ADODB.Recordset
Public RecPesq2 As New ADODB.Recordset

Private Sub TimerHora_Timer()
    LblHora.Caption = Time
End Sub

Private Sub Form_Activate()
    '==== Verifica se tem alerta de estoque =====
    Conecta
    
    Dim RecAlerta As New ADODB.Recordset
    Dim RecVerif As New ADODB.Recordset
    
    StrSql = "Select Ativado From tb_Alerta"
    RecAlerta.Open StrSql, vgCon, 1, 3

    If RecAlerta!ativado = "sim" Then
        'Alerta está ativado
        ChkDesatAlerta.Value = 0
            
        'verifica se tem produto com qtde mínima
        StrSql = "SELECT QtdeMin FROM tb_Estoque WHERE QtdeProd <= QtdeMin"
        RecVerif.Open StrSql, vgCon, 1, 3
        
        If Not RecVerif.EOF Then
            Desconecta
            
            VPStrResponse = MsgBox("Existem produtos no estoque com quantidade mínima." & Chr(13) & "Deseja visualizar a listagem agora?", vbYesNo, "Pró Vendas 2004 - Alerta de Estoque")
            If VPStrResponse = vbYes Then
                FrmEstoque_Alerta.Show
            End If
        Else
            Desconecta
        End If
    Else
        Desconecta
        ChkDesatAlerta.Value = 1
    End If
    
    'Call MontaCboTipoVenda
    '============================================
End Sub

Private Sub Form_Resize()
  TabPrincipal.Left = (MDIPrincipal.Width / 2) - (TabPrincipal.Width / 2)
  TabPrincipal.Top = (MDIPrincipal.Height / 3) - (TabPrincipal.Height / 3)
  
  FraCliente.Left = (MDIPrincipal.Width / 2) - (FraCliente.Width / 2)
  FraCliente.Top = (MDIPrincipal.Height / 3) - (FraCliente.Height / 3.9)
  
  FraFornecedor.Left = (MDIPrincipal.Width / 2) - (FraFornecedor.Width / 2)
  FraFornecedor.Top = (MDIPrincipal.Height / 3) - (FraFornecedor.Height / 3.9)
  
  FraEstoque.Left = (MDIPrincipal.Width / 2) - (FraEstoque.Width / 2)
  FraEstoque.Top = (MDIPrincipal.Height / 3) - (FraEstoque.Height / 3.9)
  
  FraCrediario.Left = (MDIPrincipal.Width / 2) - (FraCrediario.Width / 2)
  FraCrediario.Top = (MDIPrincipal.Height / 3) - (FraCrediario.Height / 3.9)
  
  FraCrediarista.Left = (MDIPrincipal.Width / 2) - (FraCrediarista.Width / 2)
  FraCrediarista.Top = (MDIPrincipal.Height / 3) - (FraCrediarista.Height / 3.9)
  
  FraProduto.Left = (MDIPrincipal.Width / 2) - (FraProduto.Width / 2)
  FraProduto.Top = (MDIPrincipal.Height / 3) - (FraProduto.Height / 3.9)
  
  FraVenda.Left = (MDIPrincipal.Width / 2) - (FraVenda.Width / 2)
  FraVenda.Top = (MDIPrincipal.Height / 3) - (FraVenda.Height / 3.9)
  
  FraCaixa.Left = (MDIPrincipal.Width / 2) - (FraCaixa.Width / 2)
  FraCaixa.Top = (MDIPrincipal.Height / 3) - (FraCaixa.Height / 3.9)
  
  FraPagar.Left = (MDIPrincipal.Width / 2) - (FraPagar.Width / 2)
  FraPagar.Top = (MDIPrincipal.Height / 3) - (FraPagar.Height / 3.9)
  
  FraReceber.Left = (MDIPrincipal.Width / 2) - (FraReceber.Width / 2)
  FraReceber.Top = (MDIPrincipal.Height / 3) - (FraReceber.Height / 3.9)
  
  FraOrcamento.Left = (MDIPrincipal.Width / 2) - (FraOrcamento.Width / 2)
  FraOrcamento.Top = (MDIPrincipal.Height / 3) - (FraOrcamento.Height / 3.9)
  
  FraVendedor.Left = (MDIPrincipal.Width / 2) - (FraVendedor.Width / 2)
  FraVendedor.Top = (MDIPrincipal.Height / 3) - (FraVendedor.Height / 3.9)
  
  FraExtra.Left = (MDIPrincipal.Width / 2) - (FraExtra.Width / 2)
  FraExtra.Top = (MDIPrincipal.Height / 3) - (FraExtra.Height / 3.9)
  
  FraDesenv.Left = (MDIPrincipal.Width / 2) - (FraDesenv.Width / 2)
  FraDesenv.Top = (MDIPrincipal.Height / 3) - (FraDesenv.Height / 3.9)

  Call PegarResolucao
  
  If VGStrResolucao = "1024x768" Then
    LblData.Left = 2100
    LblData.Top = 400
  
    LblHora.Left = 10500
    LblHora.Top = 400
  
  ElseIf VGStrResolucao = "800x600" Then
    LblData.Left = 600
    LblData.Top = 100
  
    LblHora.Left = 8760
    LblHora.Top = 100
  End If
End Sub

Private Sub Form_Load()
    Skin1.LoadSkin (App.Path & "\ProVendas2004.skn")
    Skin1.ApplySkin (Me.hwnd)
    
    Height = 10440
    Width = 14565
    
    LblData.Caption = FormataSemana(Date) & ", " & FormataDataCompleta(Date)
    
    FraCliente.Visible = False
    FraFornecedor.Visible = False
    FraEstoque.Visible = False
    FraProduto.Visible = False
    FraVenda.Visible = False
    FraCaixa.Visible = False
    FraPagar.Visible = False
    FraReceber.Visible = False
    FraOrcamento.Visible = False
    FraVendedor.Visible = False
    FraExtra.Visible = False
    FraCrediario.Visible = False
    FraCrediarista.Visible = False
    FraDesenv.Visible = False
    
    TabPrincipal.Tabs.Item(1).Selected = True
End Sub

'========================================================================
'                    TAB PRINCIPAL
'========================================================================
Private Sub TabPrincipal_Click()
    If TabPrincipal.Tabs.Item(1).Selected = True Then
        '=== VENDA ===
        FraCliente.Visible = False
        FraFornecedor.Visible = False
        FraEstoque.Visible = False
        FraProduto.Visible = False
        FraVenda.Visible = True
        FraCaixa.Visible = False
        FraPagar.Visible = False
        FraReceber.Visible = False
        FraOrcamento.Visible = False
        FraVendedor.Visible = False
        FraExtra.Visible = False
        FraCrediario.Visible = False
        FraCrediarista.Visible = False
        FraDesenv.Visible = False
        
        Call MontaCboTipoVenda
        
        CmdExcluirVenda.Enabled = False
        CmdImprimirVenda.Enabled = False
        CmdImprimirRecibo.Enabled = False
        CmdImprimirProp.Enabled = False
        CmdImprimirCarne.Enabled = False
    
    ElseIf TabPrincipal.Tabs.Item(2).Selected = True Then
        '=== CLIENTE ===
        FraCliente.Visible = True
        FraFornecedor.Visible = False
        FraEstoque.Visible = False
        FraProduto.Visible = False
        FraVenda.Visible = False
        FraCaixa.Visible = False
        FraPagar.Visible = False
        FraReceber.Visible = False
        FraOrcamento.Visible = False
        FraVendedor.Visible = False
        FraExtra.Visible = False
        FraCrediario.Visible = False
        FraCrediarista.Visible = False
        FraDesenv.Visible = False

        'monta combo de sexo
        CboSexoCli.Clear
        CboSexoCli.AddItem ("")
        CboSexoCli.AddItem ("Feminino")
        CboSexoCli.AddItem ("Masculino")

        CmdAlterarCli.Enabled = False
        CmdExcluirCli.Enabled = False
        CmdImprimirCli.Enabled = False

        TxtNomeCli.SetFocus
    
    ElseIf TabPrincipal.Tabs.Item(3).Selected = True Then
        '=== FORNECEDOR ===
        FraCliente.Visible = False
        FraFornecedor.Visible = True
        FraEstoque.Visible = False
        FraProduto.Visible = False
        FraVenda.Visible = False
        FraCaixa.Visible = False
        FraPagar.Visible = False
        FraReceber.Visible = False
        FraOrcamento.Visible = False
        FraVendedor.Visible = False
        FraExtra.Visible = False
        FraCrediario.Visible = False
        FraCrediarista.Visible = False
        FraDesenv.Visible = False

        CmdAlterarForn.Enabled = False
        CmdExcluirForn.Enabled = False
        CmdImprimirForn.Enabled = False

        TxtNomeForn.SetFocus
    
    ElseIf TabPrincipal.Tabs.Item(4).Selected = True Then
        '=== PRODUTO ===
        FraCliente.Visible = False
        FraFornecedor.Visible = False
        FraEstoque.Visible = False
        FraProduto.Visible = True
        FraVenda.Visible = False
        FraCaixa.Visible = False
        FraPagar.Visible = False
        FraReceber.Visible = False
        FraOrcamento.Visible = False
        FraVendedor.Visible = False
        FraExtra.Visible = False
        FraCrediario.Visible = False
        FraCrediarista.Visible = False
        FraDesenv.Visible = False

        CmdAlterarProd.Enabled = False
        CmdExcluirProd.Enabled = False
        CmdImprimirProd.Enabled = False

        TxtNomeProd.SetFocus
    
    ElseIf TabPrincipal.Tabs.Item(5).Selected = True Then
        '=== ESTOQUE ===
        FraCliente.Visible = False
        FraFornecedor.Visible = False
        FraEstoque.Visible = True
        FraProduto.Visible = False
        FraVenda.Visible = False
        FraCaixa.Visible = False
        FraPagar.Visible = False
        FraReceber.Visible = False
        FraOrcamento.Visible = False
        FraVendedor.Visible = False
        FraExtra.Visible = False
        FraCrediario.Visible = False
        FraCrediarista.Visible = False
        FraDesenv.Visible = False

        CmdExcluirEst.Enabled = False
        CmdImprimirEst.Enabled = False

        TxtNomeProdEst.SetFocus
    
    ElseIf TabPrincipal.Tabs.Item(6).Selected = True Then
        '=== CREDIÁRIO ===
        FraCliente.Visible = False
        FraFornecedor.Visible = False
        FraEstoque.Visible = False
        FraProduto.Visible = False
        FraVenda.Visible = False
        FraCaixa.Visible = False
        FraPagar.Visible = False
        FraReceber.Visible = False
        FraOrcamento.Visible = False
        FraVendedor.Visible = False
        FraExtra.Visible = False
        FraCrediario.Visible = True
        FraCrediarista.Visible = False
        FraDesenv.Visible = False

        Call MontaCboTipoCred

        CmdExcluirCred.Enabled = False
        CmdImprimirCred.Enabled = False
        CmdVerParc.Enabled = False

        TxtCliCred.SetFocus
    
    ElseIf TabPrincipal.Tabs.Item(7).Selected = True Then
        '=== CREDIARISTA ===
        FraCliente.Visible = False
        FraFornecedor.Visible = False
        FraEstoque.Visible = False
        FraProduto.Visible = False
        FraVenda.Visible = False
        FraCaixa.Visible = False
        FraPagar.Visible = False
        FraReceber.Visible = False
        FraOrcamento.Visible = False
        FraVendedor.Visible = False
        FraExtra.Visible = False
        FraCrediario.Visible = False
        FraCrediarista.Visible = True
        FraDesenv.Visible = False

        CmdExcluirEst.Enabled = False
        CmdImprimirEst.Enabled = False

        TxtNomeCredsta.SetFocus
    
    ElseIf TabPrincipal.Tabs.Item(8).Selected = True Then
        '=== CAIXA ===
        FraCliente.Visible = False
        FraFornecedor.Visible = False
        FraEstoque.Visible = False
        FraProduto.Visible = False
        FraVenda.Visible = False
        FraCaixa.Visible = True
        FraPagar.Visible = False
        FraReceber.Visible = False
        FraOrcamento.Visible = False
        FraVendedor.Visible = False
        FraExtra.Visible = False
        FraCrediario.Visible = False
        FraCrediarista.Visible = False
        FraDesenv.Visible = False

        CmdAlterarCx.Enabled = False
        CmdExcluirCx.Enabled = False
        CmdImprimirCx.Enabled = False

        Call MontaCboTipoPagtoCX

        TxtDtMovCx1.SetFocus
    
    ElseIf TabPrincipal.Tabs.Item(9).Selected = True Then
        '=== A PAGAR ===
        FraCliente.Visible = False
        FraFornecedor.Visible = False
        FraEstoque.Visible = False
        FraProduto.Visible = False
        FraVenda.Visible = False
        FraCaixa.Visible = False
        FraPagar.Visible = True
        FraReceber.Visible = False
        FraOrcamento.Visible = False
        FraVendedor.Visible = False
        FraExtra.Visible = False
        FraCrediario.Visible = False
        FraCrediarista.Visible = False
        FraDesenv.Visible = False

        CmdAlterarAPagar.Enabled = False
        CmdExcluirAPagar.Enabled = False
        CmdImprimirAPagar.Enabled = False
        CmdBaixarAPagar.Enabled = False

        TxtDtAPagar1.SetFocus
    
    ElseIf TabPrincipal.Tabs.Item(10).Selected = True Then
        '=== A RECEBER ===
        FraCliente.Visible = False
        FraFornecedor.Visible = False
        FraEstoque.Visible = False
        FraProduto.Visible = False
        FraVenda.Visible = False
        FraCaixa.Visible = False
        FraPagar.Visible = False
        FraReceber.Visible = True
        FraOrcamento.Visible = False
        FraVendedor.Visible = False
        FraExtra.Visible = False
        FraCrediario.Visible = False
        FraCrediarista.Visible = False
        FraDesenv.Visible = False

        CmdAlterarAReceber.Enabled = False
        CmdExcluirAReceber.Enabled = False
        CmdImprimirAReceber.Enabled = False
        CmdBaixaAReceber.Enabled = False

        TxtDtReceb1.SetFocus
    
    ElseIf TabPrincipal.Tabs.Item(11).Selected = True Then
        '=== ORÇAMENTO ===
        FraCliente.Visible = False
        FraFornecedor.Visible = False
        FraEstoque.Visible = False
        FraProduto.Visible = False
        FraVenda.Visible = False
        FraCaixa.Visible = False
        FraPagar.Visible = False
        FraReceber.Visible = False
        FraOrcamento.Visible = True
        FraVendedor.Visible = False
        FraExtra.Visible = False
        FraCrediario.Visible = False
        FraCrediarista.Visible = False
        FraDesenv.Visible = False

        CmdAlterarOrc.Enabled = False
        CmdExcluirOrc.Enabled = False
        CmdImprimirOrc.Enabled = False

        TxtCliOrc.SetFocus
    
    ElseIf TabPrincipal.Tabs.Item(12).Selected = True Then
        '=== VENDEDOR ===
        FraCliente.Visible = False
        FraFornecedor.Visible = False
        FraEstoque.Visible = False
        FraProduto.Visible = False
        FraVenda.Visible = False
        FraCaixa.Visible = False
        FraPagar.Visible = False
        FraReceber.Visible = False
        FraOrcamento.Visible = False
        FraVendedor.Visible = True
        FraExtra.Visible = False
        FraCrediario.Visible = False
        FraCrediarista.Visible = False
        FraDesenv.Visible = False

        CmdAlterarVend.Enabled = False
        CmdExcluirVend.Enabled = False
        CmdImprimirVend.Enabled = False

        TxtNomeVend.SetFocus
    
    ElseIf TabPrincipal.Tabs.Item(13).Selected = True Then
        '=== EXTRA ===
        FraCliente.Visible = False
        FraFornecedor.Visible = False
        FraEstoque.Visible = False
        FraProduto.Visible = False
        FraVenda.Visible = False
        FraCaixa.Visible = False
        FraPagar.Visible = False
        FraReceber.Visible = False
        FraOrcamento.Visible = False
        FraVendedor.Visible = False
        FraExtra.Visible = True
        FraCrediario.Visible = False
        FraCrediarista.Visible = False
        FraDesenv.Visible = False
    End If
End Sub
'========================================================================
'========================================================================



'========================================================================
'                   CLIENTE
'========================================================================

Private Sub GridCliente_DblClick(ByVal Col As Long, ByVal Row As Long)
    GridCliente.Row = Row
    GridCliente.Col = 16
    If GridCliente.Text <> "" And GridCliente.Text <> "CodCli" Then
        VGIntCodCli = GridCliente.Text
        FrmResumo_Cliente.Show
    End If
End Sub

Private Sub CmdLimparCli_Click()
    TxtNomeCli.Text = ""
    TxtBairroCli.Text = ""
    TxtCpfCli.Text = ""
    TxtTelCli.Text = ""
    CboSexoCli.ListIndex = 0
    GridCliente.MaxRows = 0
    LblNumTotalCli.Caption = "Nenhum cliente encontrado."
    
    CmdAlterarCli.Enabled = False
    CmdExcluirCli.Enabled = False
    CmdImprimirCli.Enabled = False
End Sub

Private Sub CmdAlterarCli_Click()
    If VGIntCodCli = 0 Then
        VPStrBox = MsgBox("Selecione um cliente na lista para alterar", vbExclamation, "Pró Vendas 2004 - Atenção")
    Else
        FrmCliente_Alt.Show
    End If
End Sub

Private Sub CmdExcluirCli_Click()
    If VGIntCodCli = 0 Then
        VPStrBox = MsgBox("Selecione um cliente na lista para excluir", vbExclamation, "Pró Vendas 2004 - Atenção")
    Else
        VPStrResponse = MsgBox("Deseja excluir este cliente?", vbYesNo, "Pró Vendas 2004 - Informação")
        
        If VPStrResponse = vbYes Then
            Conecta
            vgCon.Execute ("DELETE FROM tb_Cliente WHERE CodCli=" & VGIntCodCli)
            Desconecta
    
            FrmPrincipal.CmdPesqCli.Value = True
        End If
    End If
End Sub

Private Sub TxtCliOrc_GotFocus()
    TxtCliOrc.SelStart = 0
    TxtCliOrc.SelLength = Len(TxtCliOrc.Text)
End Sub

Private Sub TxtCliReceb_GotFocus()
    TxtCliReceb.SelStart = 0
    TxtCliReceb.SelLength = Len(TxtCliReceb.Text)
End Sub

Private Sub TxtCpfCli_KeyPress(KeyAscii As Integer)
    '=== Só aceita números ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtDescrAPagar_GotFocus()
    TxtDescrAPagar.SelStart = 0
    TxtDescrAPagar.SelLength = Len(TxtDescrAPagar.Text)
End Sub

Private Sub TxtDia1_GotFocus()
    TxtDia1.SelStart = 0
    TxtDia1.SelLength = Len(TxtDia1.Text)
End Sub

Private Sub TxtDia2_GotFocus()
    TxtDia2.SelStart = 0
    TxtDia2.SelLength = Len(TxtDia2.Text)
End Sub

Private Sub TxtMes1_GotFocus()
    TxtMes1.SelStart = 0
    TxtMes1.SelLength = Len(TxtMes1.Text)
End Sub

Private Sub TxtMes2_GotFocus()
    TxtMes2.SelStart = 0
    TxtMes2.SelLength = Len(TxtMes2.Text)
End Sub

Private Sub TxtNomeCli_GotFocus()
    TxtNomeCli.SelStart = 0
    TxtNomeCli.SelLength = Len(TxtNomeCli.Text)
End Sub

Private Sub TxtBairroCli_GotFocus()
    TxtBairroCli.SelStart = 0
    TxtBairroCli.SelLength = Len(TxtBairroCli.Text)
End Sub

Private Sub TxtCpfCli_GotFocus()
    TxtCpfCli.SelStart = 0
    TxtCpfCli.SelLength = Len(TxtCpfCli.Text)
End Sub

Private Sub TxtNomeVend_GotFocus()
    TxtNomeVend.SelStart = 0
    TxtNomeVend.SelLength = Len(TxtNomeVend.Text)
End Sub

Private Sub TxtTelCli_GotFocus()
    TxtTelCli.SelStart = 0
    TxtTelCli.SelLength = Len(TxtTelCli.Text)
End Sub

Private Sub CmdImprimirCli_Click()
    Screen.MousePointer = vbHourglass

    Dim datacad As String
    Dim nome As String
    Dim sexo As String
    Dim endereco As String
    Dim bairro As String
    Dim cep As String
    Dim cidest As String
    Dim datanasc As String
    Dim tel As String
    Dim cel As String
    Dim cpf As String
    Dim email As String
    Dim obs As String

    Dim VLStrLinha As String

    VLStrLinha = 1

    Conecta

    Do While VLStrLinha <= GridCliente.MaxRows

        GridCliente.Col = 1
        GridCliente.Row = VLStrLinha
        nome = GridCliente.Text

        GridCliente.Col = 2
        GridCliente.Row = VLStrLinha
        datacad = GridCliente.Text

        GridCliente.Col = 3
        GridCliente.Row = VLStrLinha
        sexo = GridCliente.Text

        GridCliente.Col = 4
        GridCliente.Row = VLStrLinha
        endereco = GridCliente.Text

        GridCliente.Col = 5
        GridCliente.Row = VLStrLinha
        bairro = GridCliente.Text

        GridCliente.Col = 6
        GridCliente.Row = VLStrLinha
        cep = GridCliente.Text

        GridCliente.Col = 7
        GridCliente.Row = VLStrLinha
        cidest = GridCliente.Text

        GridCliente.Col = 8
        GridCliente.Row = VLStrLinha
        cidest = cidest & "/" & GridCliente.Text

        GridCliente.Col = 9
        GridCliente.Row = VLStrLinha
        datanasc = GridCliente.Text

        GridCliente.Col = 10
        GridCliente.Row = VLStrLinha
        tel = GridCliente.Text

        GridCliente.Col = 11
        GridCliente.Row = VLStrLinha
        cel = GridCliente.Text

        GridCliente.Col = 12
        GridCliente.Row = VLStrLinha
        cpf = GridCliente.Text

        GridCliente.Col = 13
        GridCliente.Row = VLStrLinha
        email = GridCliente.Text

        GridCliente.Col = 14
        GridCliente.Row = VLStrLinha
        obs = GridCliente.Text

        vgCon.Execute "INSERT INTO tb_Auxiliar " & _
        "(campo01,campo02,campo03,campo04,campo05,campo06,campo07,campo08,campo09,campo10,campo11,campo12,campo13) " & _
        "VALUES ('" & datacad & "','" & nome & "','" & sexo & "','" & endereco & "','" & bairro & "','" & cep & "','" & cidest & "','" & datanasc & "','" & tel & "','" & cel & "','" & cpf & "','" & email & "','" & obs & "')"

        VLStrLinha = VLStrLinha + 1
    Loop

    Desconecta

    rptCliente.Show

End Sub

Private Sub CmdIncluirCli_Click()
    FrmCliente_Inc.Show
End Sub

Private Sub GridCliente_Click(ByVal Col As Long, ByVal Row As Long)
    GridCliente.Row = Row
    GridCliente.Col = 16
    If GridCliente.Text <> "" And GridCliente.Text <> "CodCli" Then
        VGIntCodCli = GridCliente.Text
        CmdAlterarCli.Enabled = True
        CmdExcluirCli.Enabled = True
    Else
        CmdAlterarCli.Enabled = False
        CmdExcluirCli.Enabled = False
    End If
End Sub

Sub MontaGridCliente()
    Dim VLIntCodCli As Long
    Dim VLIntLinha As Long
    Dim VLStrTel1 As String
    Dim VLStrTel2 As String

    If RecPesq.EOF Then
        LblNumTotalCli.Caption = "Nenhum cliente encontrado."

        VPStrBox = MsgBox("Pesquisa sem resultados.", vbInformation, "Pró Vendas 2004 - Informação")
        GridCliente.Refresh
        GridCliente.MaxRows = 0

        CmdAlterarCli.Enabled = False
        CmdExcluirCli.Enabled = False
        CmdImprimirCli.Enabled = False

    Else

        VLIntLinha = 1
        GridCliente.MaxRows = VLIntLinha

        Do While Not RecPesq.EOF

            GridCliente.Row = VLIntLinha
            GridCliente.Lock = True

            'Nome
            GridCliente.Col = 1
            GridCliente.Text = VerificaNulo(RecPesq!nome)
            GridCliente.Lock = True

            'Cliente desde
            GridCliente.Col = 2
            GridCliente.Text = FormataData(RecPesq!dtcad)
            GridCliente.Lock = True

            'Sexo
            GridCliente.Col = 3
            GridCliente.Text = VerificaNulo(RecPesq!sexo)
            GridCliente.Lock = True

            'Endereço
            GridCliente.Col = 4
            GridCliente.Text = VerificaNulo(RecPesq!endereco)
            GridCliente.Lock = True

            'Bairro
            GridCliente.Col = 5
            GridCliente.Text = VerificaNulo(RecPesq!bairro)
            GridCliente.Lock = True

            'Cep
            GridCliente.Col = 6
            GridCliente.Text = VerificaNulo(RecPesq!cep)
            GridCliente.Lock = True

            'Cidade
            GridCliente.Col = 7
            GridCliente.Text = VerificaNulo(RecPesq!cidade)
            GridCliente.Lock = True

            'Estado
            GridCliente.Col = 8
            GridCliente.Text = VerificaNulo(RecPesq!Estado)
            GridCliente.Lock = True

            'Data Nascimento
            GridCliente.Col = 9
            GridCliente.Text = FormataData(VerificaNulo(RecPesq!dtnasc))
            GridCliente.Lock = True

            'Telefone
            GridCliente.Col = 10
            If RecPesq!telefone1 <> "" And IsNull(RecPesq!telefone1) = False Then
                VLStrTel1 = RecPesq!telefone1
            Else
                VLStrTel1 = ""
            End If
            
            If RecPesq!telefone2 <> "" And IsNull(RecPesq!telefone2) = False Then
                VLStrTel2 = RecPesq!telefone2
            Else
                VLStrTel2 = ""
            End If
            
            If VLStrTel1 = "" And VLStrTel2 = "" Then
                GridCliente.Text = ""
                
            ElseIf VLStrTel1 <> "" And VLStrTel2 <> "" Then
                GridCliente.Text = VLStrTel1 & " / " & VLStrTel2
                
            ElseIf VLStrTel1 <> "" And VLStrTel2 = "" Then
                GridCliente.Text = VLStrTel1
                
            ElseIf VLStrTel1 = "" And VLStrTel2 <> "" Then
                GridCliente.Text = VLStrTel2
                
            End If
            GridCliente.Lock = True

            'Celular
            GridCliente.Col = 11
            GridCliente.Text = VerificaNulo(RecPesq!celular)
            GridCliente.Lock = True

            'Fax
            GridCliente.Col = 12
            GridCliente.Text = VerificaNulo(RecPesq!fax)
            GridCliente.Lock = True

            'Cpf
            GridCliente.Col = 13
            GridCliente.Text = VerificaNulo(RecPesq!cpf)
            GridCliente.Lock = True

            'Email
            GridCliente.Col = 14
            GridCliente.Text = VerificaNulo(RecPesq!email)
            GridCliente.Lock = True

            'Observação
            GridCliente.Col = 15
            GridCliente.Text = VerificaNulo(RecPesq!obs)
            GridCliente.Lock = True

            'CodCli
            GridCliente.Col = 16
            GridCliente.Text = Val(RecPesq!CodCli)
            GridCliente.Lock = True

            VLIntLinha = VLIntLinha + 1

            GridCliente.MaxRows = GridCliente.MaxRows + 1
            RecPesq.MoveNext
         Loop

         '===== CONTAGEM DE CLIENTES PESQUISADOS =========
         GridCliente.MaxRows = GridCliente.MaxRows - 1

         If GridCliente.MaxRows = 1 Then
            LblNumTotalCli.Caption = FormataNum(GridCliente.MaxRows) & " cliente encontrado."
         Else
            LblNumTotalCli.Caption = FormataNum(GridCliente.MaxRows) & " clientes encontrados."
         End If
         '================================================

         CmdImprimirCli.Enabled = True
    End If

End Sub

Private Sub TxtTelCli_KeyPress(KeyAscii As Integer)
    '=== Só aceita números, parênteses, espaço e - ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 40 And KeyAscii <> 41 And KeyAscii <> 32 And KeyAscii <> 45 Then
        KeyAscii = 0
    End If
End Sub

Private Sub CmdPesqCli_Click()

    Screen.MousePointer = vbHourglass

    Dim VLStrOrder As String

    Conecta

    StrSql = "Select * from tb_Cliente where 0=0"

    '====== PESQUISAR POR NOME ==========
    If TxtNomeCli.Text <> "" Then
        StrSql = StrSql + " and Nome like '%" & TxtNomeCli.Text & "%'"
        VLStrOrder = VLStrOrder + "Nome,"
    End If

    '====== PESQUISAR POR CPF ==========
    If TxtCpfCli.Text <> "" Then
        StrSql = StrSql + " and Cpf='" & TxtCpfCli.Text & "'"
        VLStrOrder = VLStrOrder + "Cpf,"
    End If

    '====== PESQUISAR POR SEXO ==========
    If CboSexoCli.Text <> "" Then
        StrSql = StrSql + " and Sexo='" & CboSexoCli.Text & "'"
        VLStrOrder = VLStrOrder + "Sexo,"
    End If

    '====== PESQUISAR POR BAIRRO ==========
    If TxtBairroCli.Text <> "" Then
        StrSql = StrSql + " and Bairro like '%" & TxtBairroCli.Text & "%'"
        VLStrOrder = VLStrOrder + "Bairro,"
    End If

    '====== PESQUISAR POR TELEFONE ==========
    If TxtTelCli.Text <> "" Then
        StrSql = StrSql + " and Telefone1 like '%" & TxtTelCli.Text & "%' or Telefone2 like '%" & TxtTelCli.Text & "%'"
        VLStrOrder = VLStrOrder + "Telefone1,Telefone2,"
    End If

    '====== ORDENAR PESQUISA ======================
    If VLStrOrder <> "" Then
        VLStrOrder = Mid(VLStrOrder, 1, Len(VLStrOrder) - 1)
        StrSql = StrSql + " order by " & VLStrOrder
    Else
        StrSql = StrSql + " order by Nome"
    End If

    VLStrOrder = ""

    RecPesq.Open StrSql, vgCon, 1, 3

    Call MontaGridCliente

    Desconecta

    Screen.MousePointer = vbNormal

End Sub
'========================================================================
'========================================================================




'========================================================================
'                   FORNECEDOR
'========================================================================

Private Sub GridFornecedor_DblClick(ByVal Col As Long, ByVal Row As Long)
    GridFornecedor.Row = Row
    GridFornecedor.Col = 15
    If GridFornecedor.Text <> "" And GridFornecedor.Text <> "CodForn" Then
        VGIntCodForn = GridFornecedor.Text
        FrmResumo_Fornecedor.Show
    End If
End Sub

Private Sub CmdLimpForn_Click()
    TxtNomeForn.Text = ""
    TxtTelForn.Text = ""
    TxtCnpjForn.Text = ""
    TxtTipoForn.Text = ""
    GridFornecedor.MaxRows = 0
    LblNumTotalForn.Caption = "Nenhum fornecedor encontrado."
    
    CmdAlterarForn.Enabled = False
    CmdExcluirForn.Enabled = False
    CmdImprimirForn.Enabled = False
End Sub

Private Sub CmdIncluirForn_Click()
    FrmFornecedor_Inc.Show
End Sub

Private Sub CmdAlterarForn_Click()
    If VGIntCodForn = 0 Then
        VPStrBox = MsgBox("Selecione um fornecedor na lista para alterar", vbExclamation, "Pró Vendas 2004 - Atenção")
    Else
        FrmFornecedor_Alt.Show
    End If
End Sub

Private Sub CmdExcluirForn_Click()
    If VGIntCodForn = 0 Then
        VPStrBox = MsgBox("Selecione um fornecedor na lista para excluir", vbExclamation, "Pró Vendas 2004 - Atenção")
    Else
        VPStrResponse = MsgBox("Deseja excluir este fornecedor?", vbYesNo, "Pró Vendas 2004 - Informação")
    
        If VPStrResponse = vbYes Then
            Conecta
            vgCon.Execute ("DELETE FROM tb_Fornecedor WHERE CodForn=" & VGIntCodForn)
            Desconecta
    
            FrmPrincipal.CmdPesqForn.Value = True
        End If
    End If
End Sub

Private Sub TxtTelForn_KeyPress(KeyAscii As Integer)
    '=== Só aceita números, parênteses, espaço, - e / ===
    If (KeyAscii < 47 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 40 And KeyAscii <> 41 And KeyAscii <> 32 And KeyAscii <> 45 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtNomeForn_GotFocus()
    TxtNomeForn.SelStart = 0
    TxtNomeForn.SelLength = Len(TxtNomeForn.Text)
End Sub

Private Sub TxtTelForn_GotFocus()
    TxtTelForn.SelStart = 0
    TxtTelForn.SelLength = Len(TxtTelForn.Text)
End Sub

Private Sub TxtCnpjForn_GotFocus()
    TxtCnpjForn.SelStart = 0
    TxtCnpjForn.SelLength = Len(TxtCnpjForn.Text)
End Sub

Private Sub TxtTelOrc_GotFocus()
    TxtTelOrc.SelStart = 0
    TxtTelOrc.SelLength = Len(TxtTelOrc.Text)
End Sub

Private Sub TxtTipoForn_GotFocus()
    TxtTipoForn.SelStart = 0
    TxtTipoForn.SelLength = Len(TxtTipoForn.Text)
End Sub

Private Sub GridFornecedor_Click(ByVal Col As Long, ByVal Row As Long)
    GridFornecedor.Row = Row
    GridFornecedor.Col = 15
    If GridFornecedor.Text <> "" And GridFornecedor.Text <> "CodForn" Then
        VGIntCodForn = GridFornecedor.Text

        CmdAlterarForn.Enabled = True
        CmdExcluirForn.Enabled = True
    Else
        CmdAlterarForn.Enabled = False
        CmdExcluirForn.Enabled = False
    End If
End Sub

Private Sub CmdPesqForn_Click()
    Screen.MousePointer = vbHourglass

    Dim VLStrOrder As String

    Conecta

    StrSql = "Select * from tb_Fornecedor where 0=0"

    '====== PESQUISAR POR FORNECEDOR ==========
    If TxtNomeForn.Text <> "" Then
        StrSql = StrSql + " and Nome like '%" & TxtNomeForn.Text & "%'"
        VLStrOrder = VLStrOrder + "Nome,"
    End If

    '====== PESQUISAR POR CNPJ ==========
    If TxtCnpjForn.Text <> "" Then
        StrSql = StrSql + " and CNPJ='" & TxtCnpjForn.Text & "'"
        VLStrOrder = VLStrOrder + "CNPJ,"
    End If

    '====== PESQUISAR POR TELEFONE ==========
    If TxtTelForn.Text <> "" Then
        StrSql = StrSql + " and Telefone1 like '%" & TxtTelForn.Text & "%' or Telefone2 like '%" & TxtTelForn.Text & "%'"
        VLStrOrder = VLStrOrder + "Telefone,"
    End If

    '====== PESQUISAR POR TIPO ==========
    If TxtTipoForn.Text <> "" Then
        StrSql = StrSql + " and Tipo like '%" & TxtTipoForn.Text & "%'"
        VLStrOrder = VLStrOrder + "Tipo,"
    End If


    '====== ORDENAR PESQUISA ======================
    If VLStrOrder <> "" Then
        VLStrOrder = Mid(VLStrOrder, 1, Len(VLStrOrder) - 1)
        StrSql = StrSql + " order by " & VLStrOrder
    Else
        StrSql = StrSql + " order by Nome"
    End If

    VLStrOrder = ""

    RecPesq.Open StrSql, vgCon, 1, 3

    Call MontaGridFornecedor

    Desconecta

    Screen.MousePointer = vbNormal

End Sub

Sub MontaGridFornecedor()

    Dim VLIntLinha As Long
    Dim VLStrTel1 As String
    Dim VLStrTel2 As String

    If RecPesq.EOF Then
        LblNumTotalForn.Caption = "Nenhum fornecedor encontrado."

        VPStrBox = MsgBox("Pesquisa sem resultados.", vbInformation, "Pró Vendas 2004 - Informação")
        GridFornecedor.Refresh
        GridFornecedor.MaxRows = 0

        CmdAlterarForn.Enabled = False
        CmdExcluirForn.Enabled = False
        CmdImprimirForn.Enabled = False

    Else

        VLIntLinha = 1
        GridFornecedor.MaxRows = VLIntLinha

        Do While Not RecPesq.EOF

            GridFornecedor.Row = VLIntLinha
            GridFornecedor.Lock = True

            'Fornecedor
            GridFornecedor.Col = 1
            GridFornecedor.Text = VerificaNulo(RecPesq!nome)
            GridFornecedor.Lock = True

            'Tipo
            GridFornecedor.Col = 2
            GridFornecedor.Text = VerificaNulo(RecPesq!tipo)
            GridFornecedor.Lock = True

            'Endereço
            GridFornecedor.Col = 3
            GridFornecedor.Text = VerificaNulo(RecPesq!endereco)
            GridFornecedor.Lock = True

            'Bairro
            GridFornecedor.Col = 4
            GridFornecedor.Text = VerificaNulo(RecPesq!bairro)
            GridFornecedor.Lock = True

            'Cep
            GridFornecedor.Col = 5
            GridFornecedor.Text = VerificaNulo(RecPesq!cep)
            GridFornecedor.Lock = True

            'Cidade
            GridFornecedor.Col = 6
            GridFornecedor.Text = VerificaNulo(RecPesq!cidade)
            GridFornecedor.Lock = True

            'Estado
            GridFornecedor.Col = 7
            GridFornecedor.Text = VerificaNulo(RecPesq!Estado)
            GridFornecedor.Lock = True

            'CNPJ
            GridFornecedor.Col = 8
            GridFornecedor.Text = VerificaNulo(RecPesq!cnpj)
            GridFornecedor.Lock = True

            'Email
            GridFornecedor.Col = 9
            GridFornecedor.Text = VerificaNulo(RecPesq!email)
            GridFornecedor.Lock = True

            'Responsável
            GridFornecedor.Col = 10
            GridFornecedor.Text = VerificaNulo(RecPesq!contato)
            GridFornecedor.Lock = True

            'Telefone
            GridFornecedor.Col = 11
            If RecPesq!telefone1 <> "" And IsNull(RecPesq!telefone1) = False Then
                VLStrTel1 = RecPesq!telefone1
            Else
                VLStrTel1 = ""
            End If
            
            If RecPesq!telefone2 <> "" And IsNull(RecPesq!telefone2) = False Then
                VLStrTel2 = RecPesq!telefone2
            Else
                VLStrTel2 = ""
            End If
            
            If VLStrTel1 = "" And VLStrTel2 = "" Then
                GridFornecedor.Text = ""
                
            ElseIf VLStrTel1 <> "" And VLStrTel2 <> "" Then
                GridFornecedor.Text = VLStrTel1 & " / " & VLStrTel2
                
            ElseIf VLStrTel1 <> "" And VLStrTel2 = "" Then
                GridFornecedor.Text = VLStrTel1
                
            ElseIf VLStrTel1 = "" And VLStrTel2 <> "" Then
                GridFornecedor.Text = VLStrTel2
                
            End If
            GridFornecedor.Lock = True

            'Celular
            GridFornecedor.Col = 12
            GridFornecedor.Text = VerificaNulo(RecPesq!celular)
            GridFornecedor.Lock = True

            'Fax
            GridFornecedor.Col = 13
            GridFornecedor.Text = VerificaNulo(RecPesq!fax)
            GridFornecedor.Lock = True

            'Observação
            GridFornecedor.Col = 14
            GridFornecedor.Text = VerificaNulo(RecPesq!obs)
            GridFornecedor.Lock = True

            'CodForn
            GridFornecedor.Col = 15
            GridFornecedor.Text = Val(RecPesq!CodForn)
            GridFornecedor.Lock = True

            VLIntLinha = VLIntLinha + 1

            GridFornecedor.MaxRows = GridFornecedor.MaxRows + 1
            RecPesq.MoveNext
         Loop

         '===== CONTAGEM DE FORNECEDORES PESQUISADOS =========
         GridFornecedor.MaxRows = GridFornecedor.MaxRows - 1

         If GridFornecedor.MaxRows = 1 Then
            LblNumTotalForn.Caption = FormataNum(GridFornecedor.MaxRows) & " fornecedor encontrado."
         Else
            LblNumTotalForn.Caption = FormataNum(GridFornecedor.MaxRows) & " fornecedores encontrados."
         End If
         '================================================

         CmdImprimirForn.Enabled = True
    End If

End Sub

Private Sub CmdImprimirForn_Click()
    Screen.MousePointer = vbHourglass

    Dim forn As String
    Dim tipo As String
    Dim endereco As String
    Dim bairro As String
    Dim cep As String
    Dim cidest As String
    Dim cnpj As String
    Dim email As String
    Dim resp As String
    Dim tel As String
    Dim cel As String
    Dim obs As String

    Dim VLStrLinha As String

    VLStrLinha = 1

    Conecta

    Do While VLStrLinha <= GridFornecedor.MaxRows

        GridFornecedor.Col = 1
        GridFornecedor.Row = VLStrLinha
        forn = GridFornecedor.Text

        GridFornecedor.Col = 2
        GridFornecedor.Row = VLStrLinha
        tipo = GridFornecedor.Text

        GridFornecedor.Col = 3
        GridFornecedor.Row = VLStrLinha
        endereco = GridFornecedor.Text

        GridFornecedor.Col = 4
        GridFornecedor.Row = VLStrLinha
        bairro = GridFornecedor.Text

        GridFornecedor.Col = 5
        GridFornecedor.Row = VLStrLinha
        cep = GridFornecedor.Text

        GridFornecedor.Col = 6
        GridFornecedor.Row = VLStrLinha
        cidest = GridFornecedor.Text

        GridFornecedor.Col = 7
        GridFornecedor.Row = VLStrLinha
        cidest = cidest & "/" & GridFornecedor.Text

        GridFornecedor.Col = 8
        GridFornecedor.Row = VLStrLinha
        cnpj = GridFornecedor.Text

        GridFornecedor.Col = 9
        GridFornecedor.Row = VLStrLinha
        email = GridFornecedor.Text

        GridFornecedor.Col = 10
        GridFornecedor.Row = VLStrLinha
        resp = GridFornecedor.Text

        GridFornecedor.Col = 11
        GridFornecedor.Row = VLStrLinha
        tel = GridFornecedor.Text

        GridFornecedor.Col = 12
        GridFornecedor.Row = VLStrLinha
        cel = GridFornecedor.Text

        GridFornecedor.Col = 13
        GridFornecedor.Row = VLStrLinha
        obs = GridFornecedor.Text

        vgCon.Execute "INSERT INTO tb_Auxiliar " & _
        "(campo01,campo02,campo03,campo04,campo05,campo06,campo07,campo08,campo09,campo10,campo11,campo12) " & _
        "VALUES ('" & forn & "','" & tipo & "','" & endereco & "','" & bairro & "','" & cep & "','" & cidest & " ','" & cnpj & "','" & email & "','" & resp & "','" & tel & "','" & cel & "','" & obs & "')"

        VLStrLinha = VLStrLinha + 1
    Loop

    Desconecta

    rptFornecedor.Show

End Sub
'========================================================================
'========================================================================


'========================================================================
'                             ESTOQUE
'========================================================================

Private Sub GridEstoque_DblClick(ByVal Col As Long, ByVal Row As Long)
    GridEstoque.Row = Row
    GridEstoque.Col = 6
    If GridEstoque.Text <> "" And GridEstoque.Text <> "CodEst" Then
        VGIntCodEst = GridEstoque.Text
        FrmResumo_Estoque.Show
    End If
End Sub

Private Sub CmdLimparEst_Click()
    TxtNomeProdEst.Text = ""
    TxtTipoProdEst.Text = ""
    TxtQtdeMin.Text = ""
    TxtQtdeEst.Text = ""
    GridEstoque.MaxRows = 0
    LblNumTotalEst.Caption = "Nenhuma informação encontrada."
    
    CmdExcluirEst.Enabled = False
    CmdImprimirEst.Enabled = False
End Sub

Private Sub CmdIncluirAlterarEst_Click()
    FrmEstoque_Inc_Alt.Show
End Sub

Private Sub ChkDesatAlerta_Click()
    Screen.MousePointer = vbHourglass
    Conecta

    If ChkDesatAlerta.Value = 1 Then
        'desativar o alerta
        vgCon.Execute ("Update tb_Alerta Set Ativado='não'")
    Else
        'ativar o alerta
        vgCon.Execute ("Update tb_Alerta Set Ativado='sim'")
    End If

    Desconecta
    Screen.MousePointer = vbNormal
End Sub

Private Sub CmdExcluirEst_Click()
    VPStrResponse = MsgBox("Deseja excluir este produto do estoque?", vbYesNo, "Pró Vendas 2004 - Informação")

    If VPStrResponse = vbYes Then
        Conecta
        vgCon.Execute ("DELETE FROM tb_Estoque WHERE CodEst=" & VGIntCodEst)
        Desconecta

        FrmPrincipal.CmdPesqEst.Value = True
    End If
End Sub

Private Sub GridEstoque_Click(ByVal Col As Long, ByVal Row As Long)
    GridEstoque.Row = Row
    GridEstoque.Col = 6
    If GridEstoque.Text <> "" And GridEstoque.Text <> "CodEst" Then
        VGIntCodEst = GridEstoque.Text
        CmdExcluirEst.Enabled = True
    Else
        CmdExcluirEst.Enabled = False
    End If
End Sub

Private Sub TxtNomeProdEst_GotFocus()
    TxtNomeProdEst.SelStart = 0
    TxtNomeProdEst.SelLength = Len(TxtNomeProdEst.Text)
End Sub

Private Sub TxtTipoProdEst_GotFocus()
    TxtTipoProdEst.SelStart = 0
    TxtTipoProdEst.SelLength = Len(TxtTipoProdEst.Text)
End Sub

Private Sub TxtQtdeMin_GotFocus()
    TxtQtdeMin.SelStart = 0
    TxtQtdeMin.SelLength = Len(TxtQtdeMin.Text)
End Sub

Private Sub TxtQtdeEst_GotFocus()
    TxtQtdeEst.SelStart = 0
    TxtQtdeEst.SelLength = Len(TxtQtdeEst.Text)
End Sub

Private Sub TxtQtdeMin_KeyPress(KeyAscii As Integer)
    '=== Só aceita números ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtQtdeEst_KeyPress(KeyAscii As Integer)
    '=== Só aceita números ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub CmdPesqEst_Click()
    Screen.MousePointer = vbHourglass

    Dim VLStrOrder As String

    Conecta

    StrSql = "Select * from tb_Estoque as E,tb_Produto as P where E.CodProd=P.CodProd"

    '====== PESQUISAR POR NOME DO PRODUTO ==========
    If TxtNomeProdEst.Text <> "" Then
        StrSql = StrSql + " and P.NomeProd like '%" & TxtNomeProdEst.Text & "%'"
        VLStrOrder = VLStrOrder + "P.NomeProd,"
    End If

    '====== PESQUISAR POR TIPO DE PRODUTO ==========
    If TxtTipoProdEst.Text <> "" Then
        StrSql = StrSql + " and P.TipoProd like '%" & TxtTipoProdEst.Text & "%'"
        VLStrOrder = VLStrOrder + "P.TipoProd,"
    End If

    '====== PESQUISAR POR QTDE MÍNIMA ==========
    If TxtQtdeMin.Text <> "" Then
        StrSql = StrSql + " and E.QtdeMin=" & TxtQtdeMin.Text & ""
        VLStrOrder = VLStrOrder + "E.QtdeMin,"
    End If

    '====== PESQUISAR POR QTDE EM ESTOQUE ==========
    If TxtQtdeEst.Text <> "" Then
        StrSql = StrSql + " and E.QtdeProd=" & TxtQtdeEst.Text & ""
        VLStrOrder = VLStrOrder + "E.QtdeProd,"
    End If

    '====== ORDENAR PESQUISA ======================
    If VLStrOrder <> "" Then
        VLStrOrder = Mid(VLStrOrder, 1, Len(VLStrOrder) - 1)
        StrSql = StrSql + " order by " & VLStrOrder
    Else
        StrSql = StrSql + " order by P.NomeProd"
    End If

    VLStrOrder = ""

    RecPesq.Open StrSql, vgCon, 1, 3

    Call MontaGridEstoque

    Desconecta

    Screen.MousePointer = vbNormal

End Sub

Sub MontaGridEstoque()
    Dim VLIntCodEst As Long
    Dim VLIntLinha As Long

    If RecPesq.EOF Then
        LblNumTotalEst.Caption = "Nenhuma informação encontrada."

        VPStrBox = MsgBox("Pesquisa sem resultados.", vbInformation, "Pró Vendas 2004 - Informação")
        GridEstoque.Refresh
        GridEstoque.MaxRows = 0

        CmdExcluirEst.Enabled = False
        CmdImprimirEst.Enabled = False

    Else

        VLIntLinha = 1
        GridEstoque.MaxRows = VLIntLinha

        Do While Not RecPesq.EOF

            GridEstoque.Row = VLIntLinha
            GridEstoque.Lock = True

            'Produto
            GridEstoque.Col = 1
            GridEstoque.Text = VerificaNulo(RecPesq!nomeprod)
            GridEstoque.Lock = True

            'Tipo Produto
            GridEstoque.Col = 2
            GridEstoque.Text = VerificaNulo(RecPesq!tipoprod)
            GridEstoque.Lock = True

            'Último pedido
            GridEstoque.Col = 3
            GridEstoque.Text = FormataNum(RecPesq!ultped)
            GridEstoque.Lock = True

            'Qtde Mínima
            GridEstoque.Col = 4
            GridEstoque.Text = FormataNum(RecPesq!qtdemin)
            GridEstoque.Lock = True

            'Qtde em estoque
            GridEstoque.Col = 5
            GridEstoque.Text = FormataNum(RecPesq!qtdeprod)
            GridEstoque.Lock = True

            'CodEst
            GridEstoque.Col = 6
            GridEstoque.Text = Val(RecPesq!CodEst)
            GridEstoque.Lock = True

            VLIntLinha = VLIntLinha + 1

            GridEstoque.MaxRows = GridEstoque.MaxRows + 1
            RecPesq.MoveNext
         Loop

         '===== CONTAGEM DE INFORMAÇÕES DO ESTOQUE PESQUISADOS =========
         GridEstoque.MaxRows = GridEstoque.MaxRows - 1

         If GridEstoque.MaxRows = 1 Then
            LblNumTotalEst.Caption = FormataNum(GridEstoque.MaxRows) & " informação encontrada."
         Else
            LblNumTotalEst.Caption = FormataNum(GridEstoque.MaxRows) & " informações encontradas."
         End If
         '================================================

         CmdImprimirEst.Enabled = True
    End If

End Sub

Private Sub CmdImprimirEst_Click()
    Screen.MousePointer = vbHourglass

    Dim nomeprod As String
    Dim tipoprod As String
    Dim ultped As String
    Dim qtdemin As String
    Dim qtdeest As String
    Dim precofabric As String
    Dim precovendaunit As String
    Dim precovendaatac As String

    Dim VLStrLinha As String

    VLStrLinha = 1

    Conecta

    Do While VLStrLinha <= GridEstoque.MaxRows

        GridEstoque.Col = 1
        GridEstoque.Row = VLStrLinha
        nomeprod = GridEstoque.Text

        GridEstoque.Col = 2
        GridEstoque.Row = VLStrLinha
        tipoprod = GridEstoque.Text

        GridEstoque.Col = 3
        GridEstoque.Row = VLStrLinha
        ultped = GridEstoque.Text

        GridEstoque.Col = 4
        GridEstoque.Row = VLStrLinha
        qtdemin = GridEstoque.Text

        GridEstoque.Col = 5
        GridEstoque.Row = VLStrLinha
        qtdeest = GridEstoque.Text

        vgCon.Execute "INSERT INTO tb_Auxiliar " & _
        "(campo01,campo02,campo03,campo04,campo05) " & _
        "VALUES ('" & nomeprod & "','" & tipoprod & "','" & ultped & "','" & qtdemin & "','" & qtdeest & "')"

        VLStrLinha = VLStrLinha + 1
    Loop

    Desconecta

    rptEstoque.Show

End Sub
'========================================================================
'========================================================================


'========================================================================
'                            PRODUTO
'========================================================================

Private Sub GridProduto_DblClick(ByVal Col As Long, ByVal Row As Long)
    GridProduto.Row = Row
    GridProduto.Col = 9
    If GridProduto.Text <> "" And GridProduto.Text <> "CodProd" Then
        VGIntCodProd = GridProduto.Text
        FrmResumo_Produto.Show
    End If
End Sub

Private Sub CmdLimparProd_Click()
    TxtNomeProd.Text = ""
    TxtFornProd.Text = ""
    TxtTipoProd.Text = ""
    TxtPrVendaUnitProd.Text = ""
    TxtPrVendaAtacProd.Text = ""
    GridProduto.MaxRows = 0
    LblNumTotalProd.Caption = "Nenhum produto encontrado."
    
    CmdAlterarProd.Enabled = False
    CmdExcluirProd.Enabled = False
    CmdImprimirProd.Enabled = False
End Sub

Private Sub CmdIncluirProd_Click()
    FrmProduto_Inc.Show
End Sub

Private Sub CmdAlterarProd_Click()
    FrmProduto_Alt.Show
End Sub

Private Sub CmdExcluirProd_Click()
    VPStrResponse = MsgBox("Deseja excluir este produto e seus" & Chr(13) & "lançamentos do estoque?", vbYesNo, "Pró Vendas 2004 - Informação")

    If VPStrResponse = vbYes Then
        Conecta
        vgCon.Execute ("DELETE FROM tb_Produto WHERE CodProd=" & VGIntCodProd)
        vgCon.Execute ("DELETE FROM tb_Estoque WHERE CodProd=" & VGIntCodProd)
        Desconecta

        FrmPrincipal.CmdPesqProd.Value = True

    End If
End Sub

Private Sub TxtNomeProd_GotFocus()
    TxtNomeProd.SelStart = 0
    TxtNomeProd.SelLength = Len(TxtNomeProd.Text)
End Sub

Private Sub TxtFornProd_GotFocus()
    TxtFornProd.SelStart = 0
    TxtFornProd.SelLength = Len(TxtFornProd.Text)
End Sub

Private Sub TxtTipoProd_GotFocus()
    TxtTipoProd.SelStart = 0
    TxtTipoProd.SelLength = Len(TxtTipoProd.Text)
End Sub

Private Sub TxtPrVendaUnitProd_GotFocus()
    TxtPrVendaUnitProd.SelStart = 0
    TxtPrVendaUnitProd.SelLength = Len(TxtPrVendaUnitProd.Text)
End Sub

Private Sub TxtPrVendaAtacProd_GotFocus()
    TxtPrVendaAtacProd.SelStart = 0
    TxtPrVendaAtacProd.SelLength = Len(TxtPrVendaAtacProd.Text)
End Sub

Private Sub TxtPrVendaUnitProd_KeyPress(KeyAscii As Integer)
    '=== Só aceita números e , ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 44 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtPrVendaAtacProd_KeyPress(KeyAscii As Integer)
    '=== Só aceita números e ,===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 44 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtPrVendaUnitProd_LostFocus()
    If TxtPrVendaUnitProd.Text <> "" Then
        TxtPrVendaUnitProd.Text = Trim(Mid(FormataMoeda(TxtPrVendaUnitProd.Text), 3))
    End If
End Sub

Private Sub TxtPrVendaAtacProd_LostFocus()
    If TxtPrVendaAtacProd.Text <> "" Then
        TxtPrVendaAtacProd.Text = Trim(Mid(FormataMoeda(TxtPrVendaAtacProd.Text), 3))
    End If
End Sub

Private Sub GridProduto_Click(ByVal Col As Long, ByVal Row As Long)
    GridProduto.Row = Row
    GridProduto.Col = 9
    
    If GridProduto.Text <> "" And GridProduto.Text <> "CodProd" Then
        VGIntCodProd = GridProduto.Text
    
        CmdAlterarProd.Enabled = True
        CmdExcluirProd.Enabled = True
    Else
        CmdAlterarProd.Enabled = False
        CmdExcluirProd.Enabled = False
    End If
End Sub

Private Sub CmdImprimirProd_Click()
    FrmProduto_Lista_Imprimir.Show
End Sub

Private Sub CmdPesqProd_Click()
    Screen.MousePointer = vbHourglass

    Dim VLStrOrder As String

    Conecta

    StrSql = "SELECT * FROM tb_Produto as P WHERE 0=0"

    '====== PESQUISAR POR NOME DO PRODUTO ==========
    If TxtNomeProd.Text <> "" Then
        StrSql = StrSql + " and P.NomeProd like '%" & TxtNomeProd.Text & "%'"
        VLStrOrder = VLStrOrder + "P.NomeProd,"
    End If
    
    '====== PESQUISAR POR FORNECEDOR ==========
    If TxtFornProd.Text <> "" Then
        StrSql = Trim(Mid(StrSql, 1, InStr(StrSql, "FROM") - 1)) & ",F.Nome " & Trim(Mid(StrSql, InStr(StrSql, "FROM"), InStr(StrSql, "WHERE") - InStr(StrSql, "FROM"))) & ",tb_Fornecedor as F " & Trim(Mid(StrSql, InStr(StrSql, "WHERE"))) & " and P.CodForn=F.CodForn and F.Nome like '%" & TxtFornProd.Text & "%'"
        VLStrOrder = VLStrOrder + "F.Nome,"
    End If

    '====== PESQUISAR POR TIPO DE PRODUTO ==========
    If TxtTipoProd.Text <> "" Then
        StrSql = StrSql + " and P.TipoProd like '%" & TxtTipoProd.Text & "%'"
        VLStrOrder = VLStrOrder + "P.TipoProd,"
    End If

    '====== PESQUISAR POR PREÇO DE VENDA UNITÁRIO ==========
    If TxtPrVendaUnitProd.Text <> "" Then
        StrSql = StrSql + " and P.PrecoVendaUnit='" & TxtPrVendaUnitProd.Text & "'"
        VLStrOrder = VLStrOrder + "P.PrecoVendaUnit,"
    End If

    '====== PESQUISAR POR PREÇO DE VENDA POR ATACADO ==========
    If TxtPrVendaAtacProd.Text <> "" Then
        StrSql = StrSql + " and P.PrecoVendaAtac='" & TxtPrVendaAtacProd.Text & "'"
        VLStrOrder = VLStrOrder + "P.PrecoVendaAtac,"
    End If

    '====== ORDENAR PESQUISA ======================
    If VLStrOrder <> "" Then
        VLStrOrder = Mid(VLStrOrder, 1, Len(VLStrOrder) - 1)
        StrSql = StrSql + " order by " & VLStrOrder
    Else
        StrSql = StrSql + " order by P.TipoProd"
    End If

    VLStrOrder = ""

    RecPesq.Open StrSql, vgCon, 1, 3

    Call MontaGridProduto

    Desconecta

    Screen.MousePointer = vbNormal

End Sub

Sub MontaGridProduto()
    Dim VLIntCodProd As Long
    Dim VLIntLinha As Long
    Dim RecProd As New ADODB.Recordset

    If RecPesq.EOF Then
        LblNumTotalProd.Caption = "Nenhum produto encontrado."
        VPStrBox = MsgBox("Pesquisa sem resultados.", vbInformation, "Pró Vendas 2004 - Informação")
        GridProduto.Refresh
        GridProduto.MaxRows = 0

        CmdAlterarProd.Enabled = False
        CmdExcluirProd.Enabled = False
        CmdImprimirProd.Enabled = False

    Else

        VLIntLinha = 1
        GridProduto.MaxRows = VLIntLinha

        Do While Not RecPesq.EOF

            GridProduto.Row = VLIntLinha
            GridProduto.Lock = True

            'Produto
            GridProduto.Col = 1
            GridProduto.Text = VerificaNulo(RecPesq!nomeprod)
            GridProduto.Lock = True

            'Fornecedor
            GridProduto.Col = 2
            
            StrSql = "Select Nome from tb_Fornecedor where CodForn=" & RecPesq!CodForn
            RecProd.Open StrSql, vgCon, 1, 3
            
            If Not RecProd.EOF Then
                GridProduto.Text = VerificaNulo(RecProd!nome)
            Else
                GridProduto.Text = ""
            End If
            GridProduto.Lock = True

            'Tipo produto
            GridProduto.Col = 3
            GridProduto.Text = VerificaNulo(RecPesq!tipoprod)
            GridProduto.Lock = True

            'Descrição de produto
            GridProduto.Col = 4
            GridProduto.Text = VerificaNulo(RecPesq!descprod)
            GridProduto.Lock = True

            'Preço do Fabricante
            GridProduto.Col = 5
            GridProduto.Text = FormataMoeda(VerificaNulo(RecPesq!precofabric))
            GridProduto.Lock = True

            'Preço de venda unitário
            GridProduto.Col = 6
            GridProduto.Text = FormataMoeda(VerificaNulo(RecPesq!precovendaunit))
            GridProduto.Lock = True

            'Preço de venda atacado
            GridProduto.Col = 7
            GridProduto.Text = FormataMoeda(VerificaNulo(RecPesq!precovendaatac))
            GridProduto.Lock = True

            'Moeda
            GridProduto.Col = 8
            GridProduto.Text = VerificaNulo(RecPesq!moeda)
            GridProduto.Lock = True

            'CodProd
            GridProduto.Col = 9
            GridProduto.Text = Val(RecPesq!CodProd)
            GridProduto.Lock = True

            VLIntLinha = VLIntLinha + 1

            GridProduto.MaxRows = GridProduto.MaxRows + 1

            RecPesq.MoveNext
            RecProd.Close
         Loop

         '===== CONTAGEM DE PRODUTOS PESQUISADOS =========
         GridProduto.MaxRows = GridProduto.MaxRows - 1

         If GridProduto.MaxRows = 1 Then
            LblNumTotalProd.Caption = FormataNum(GridProduto.MaxRows) & " produto encontrado."
         Else
            LblNumTotalProd.Caption = FormataNum(GridProduto.MaxRows) & " produtos encontrados."
         End If
         '================================================

         CmdImprimirProd.Enabled = True
    End If

End Sub
'========================================================================
'========================================================================



'========================================================================
'                            VENDAS
'========================================================================

Private Sub CmdImprimirCarne_Click()
    FrmAssinaturaCarne.Show
End Sub

Private Sub CmdImprimirProp_Click()
    Dim RecPesq As New ADODB.Recordset
    
    Conecta
    
    StrSql = "SELECT CodCred FROM tb_Venda WHERE CodVenda=" & VGIntCodVenda
    RecPesq.Open StrSql, vgCon, 1, 3
    
    If RecPesq.EOF Then
        Desconecta
        VPStrBox = MsgBox("Não existe crediário para esta venda", vbInformation, "Pró Vendas 2004 - Informação")
    Else
        VGIntPropCodCred = RecPesq!CodCred
        VGStrProposta = "imprimir"
        Desconecta
        VGStrAssinaturaProposta = "proposta"
        FrmAssinaturaOrc.Show
    End If
End Sub

Private Sub TxtDtVenda1_GotFocus()
    If TxtDtVenda1.Text = "__/__/____" Then
        TxtDtVenda1.Text = ""
    Else
        TxtDtVenda1.SelStart = 0
        TxtDtVenda1.SelLength = Len(TxtDtVenda1.Text)
    End If
End Sub

Private Sub TxtDtVenda1_KeyPress(KeyAscii As Integer)
    '=== Só aceita números ===
    If (KeyAscii < 47 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtDtVenda1_LostFocus()
    Dim VLStrData As String

    If TxtDtVenda1.Text <> "" Then
        VLStrData = VerificaData(TxtDtVenda1.Text)

        If VGStrDataErro = "sim" Then
            TxtDtVenda1.SetFocus
        Else
            TxtDtVenda1.Text = VLStrData
        End If

        VGStrDataErro = ""
    Else
        TxtDtVenda1.Text = "__/__/____"
    End If
End Sub

Private Sub TxtDtVenda2_GotFocus()
    If TxtDtVenda2.Text = "__/__/____" Then
        TxtDtVenda2.Text = ""
    Else
        TxtDtVenda2.SelStart = 0
        TxtDtVenda2.SelLength = Len(TxtDtVenda2.Text)
    End If
End Sub

Private Sub TxtDtVenda2_KeyPress(KeyAscii As Integer)
    '=== Só aceita números ===
    If (KeyAscii < 47 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtDtVenda2_LostFocus()
    Dim VLStrData As String

    If TxtDtVenda2.Text <> "" Then
        VLStrData = VerificaData(TxtDtVenda2.Text)

        If VGStrDataErro = "sim" Then
            TxtDtVenda2.SetFocus
        Else
            TxtDtVenda2.Text = VLStrData
        End If

        VGStrDataErro = ""
    Else
        TxtDtVenda2.Text = "__/__/____"
    End If
End Sub

Private Sub CmdLimparVend_Click()
    TxtCliVend.Text = ""
    TxtDtVenda1.Text = "__/__/____"
    TxtDtVenda2.Text = "__/__/____"
    CboTipoVenda.ListIndex = 0
    TxtVendedor.Text = ""
    GridVenda.MaxRows = 0
    LblNumTotalVend.Caption = "Nenhuma venda encontrada."
    
    CmdExcluirVenda.Enabled = False
    CmdImprimirVenda.Enabled = False
    CmdImprimirRecibo.Enabled = False
    CmdImprimirProp.Enabled = False
    CmdImprimirCarne.Enabled = False
    
End Sub

Private Sub TxtCliVend_GotFocus()
    TxtCliVend.SelStart = 0
    TxtCliVend.SelLength = Len(TxtCliVend.Text)
End Sub

Private Sub TxtVendedor_GotFocus()
    TxtVendedor.SelStart = 0
    TxtVendedor.SelLength = Len(TxtVendedor.Text)
End Sub

Private Sub CmdExcluirVenda_Click()
    Dim RecVenda As New ADODB.Recordset
    Dim RecEst As New ADODB.Recordset
    
    VPStrResponse = MsgBox("Deseja excluir esta venda?", vbYesNo, "Pró Vendas 2004 - Informação")
    
    If VPStrResponse = vbYes Then
        VPStrResponse = MsgBox("Antes de excluir essa venda, deseja" & Chr(13) & "retornar o(s) produto(s) ao estoque?", vbYesNo, "Pró Vendas 2004 - Informação")
        
        If VPStrResponse = vbYes Then
            Conecta
            
            StrSql = "SELECT CodProd,Qtde FROM tb_Venda_Produto where CodVenda=" & VGIntCodVenda
            RecVenda.Open StrSql, vgCon, 1, 3
            
            '===== Retorna estoque do(s) produto(s) =============
            Do While Not RecVenda.EOF
                StrSql = "SELECT QtdeProd FROM tb_Estoque where CodProd=" & RecVenda!CodProd
                RecEst.Open StrSql, vgCon, 1, 3
                
                If Not RecEst.EOF Then
                    RecEst("QtdeProd") = Val(RecEst!qtdeprod) + Val(RecVenda!qtde)
                    RecEst.Update
                End If
                RecEst.Close
                RecVenda.MoveNext
            Loop
            Desconecta
        End If
        
        '======== Exclui a venda =================
        Conecta
        vgCon.Execute ("DELETE FROM tb_Venda WHERE CodVenda=" & VGIntCodVenda)
        vgCon.Execute ("DELETE FROM tb_Venda_Produto WHERE CodVenda=" & VGIntCodVenda)
        Desconecta
        '========================================
                
        VPStrBox = MsgBox("Venda excluída!", vbInformation, "Pró Vendas 2004 - Informação")
        
        FrmPrincipal.CmdPesqVenda.Value = True
    End If

End Sub

Private Sub CmdImprimirVenda_Click()
    Screen.MousePointer = vbHourglass

    Dim RecVenda As New ADODB.Recordset
    Dim RecProd As New ADODB.Recordset

    Dim codvenda As Long
    Dim cliente As String
    Dim vendedor As String
    Dim datavenda As String
    Dim valorvenda As String
    Dim desconto As String
    Dim juros As String
    Dim tipovenda As String
    Dim TipoPagto As String
    Dim prod As String
    
    Dim VLStrLinha As String

    VLStrLinha = 1

    Conecta

    Do While VLStrLinha <= GridVenda.MaxRows

        GridVenda.Col = 1
        GridVenda.Row = VLStrLinha
        cliente = GridVenda.Text

        GridVenda.Col = 2
        GridVenda.Row = VLStrLinha
        vendedor = GridVenda.Text

        GridVenda.Col = 3
        GridVenda.Row = VLStrLinha
        datavenda = GridVenda.Text

        GridVenda.Col = 4
        GridVenda.Row = VLStrLinha
        valorvenda = GridVenda.Text

        GridVenda.Col = 5
        GridVenda.Row = VLStrLinha
        desconto = GridVenda.Text
        
        GridVenda.Col = 6
        GridVenda.Row = VLStrLinha
        juros = GridVenda.Text

        GridVenda.Col = 7
        GridVenda.Row = VLStrLinha
        tipovenda = GridVenda.Text

        GridVenda.Col = 8
        GridVenda.Row = VLStrLinha
        TipoPagto = GridVenda.Text

        GridVenda.Col = 9
        GridVenda.Row = VLStrLinha
        codvenda = Val(GridVenda.Text)

        StrSql = "SELECT CodProd,Qtde FROM tb_Venda_Produto where CodVenda=" & codvenda
        RecVenda.Open StrSql, vgCon, 1, 3

        If Not RecVenda.EOF Then
            Do While Not RecVenda.EOF
                '=== Pegar produtos ==========
                StrSql = "SELECT NomeProd FROM tb_Produto where CodProd=" & RecVenda!CodProd
                RecProd.Open StrSql, vgCon, 1, 3
        
                prod = prod & FormataNum(RecVenda!qtde) & " " & RecProd!nomeprod & Chr(13)
                
                RecProd.Close
                RecVenda.MoveNext
            Loop
        Else
            prod = ""
        End If
        
        RecVenda.Close
        
        vgCon.Execute "INSERT INTO tb_Auxiliar " & _
        "(campo01,campo02,campo03,campo04,campo05,campo06,campo07,campo08,campo09) " & _
        "VALUES ('" & cliente & "','" & vendedor & "','" & datavenda & "','" & valorvenda & "','" & desconto & "','" & juros & "','" & tipovenda & "','" & TipoPagto & "','" & prod & "')"

        VLStrLinha = VLStrLinha + 1
        prod = ""
    Loop

    Desconecta

    rptVenda.Show

End Sub

Private Sub CmdImprimirRecibo_Click()
    Screen.MousePointer = vbHourglass

    Dim RecVenda As New ADODB.Recordset
    Dim RecProd As New ADODB.Recordset

    Dim codvenda As Long
    Dim cliente As String
    Dim vendedor As String
    Dim datavenda As String
    Dim nomeprod As String
    Dim qtdeprod As String
    Dim valorprod  As String
    Dim subtotalvenda As String
    Dim desconto As String
    Dim juros As String
    Dim totalvenda As String
    
    Dim VLStrLinha As String

    VLStrLinha = 1
    subtotalvenda = 0
    totalvenda = 0
    
    Conecta

    'Do While VLStrLinha <= GridVenda.MaxRows

        GridVenda.Col = 1
        GridVenda.Row = GridVenda.ActiveRow
        cliente = GridVenda.Text

        GridVenda.Col = 2
        GridVenda.Row = GridVenda.ActiveRow
        vendedor = GridVenda.Text

        GridVenda.Col = 3
        GridVenda.Row = GridVenda.ActiveRow
        datavenda = GridVenda.Text

        GridVenda.Col = 4
        GridVenda.Row = GridVenda.ActiveRow
        totalvenda = GridVenda.Text

        GridVenda.Col = 5
        GridVenda.Row = GridVenda.ActiveRow
        If GridVenda.Text = "" Or IsNull(GridVenda.Text) = True Then
            desconto = "0"
        Else
            desconto = GridVenda.Text
        End If
        
        GridVenda.Col = 6
        GridVenda.Row = GridVenda.ActiveRow
        If GridVenda.Text = "" Or IsNull(GridVenda.Text) = True Then
            juros = "0"
        Else
            juros = GridVenda.Text
        End If
        
        GridVenda.Col = 9
        GridVenda.Row = GridVenda.ActiveRow
        codvenda = Val(GridVenda.Text)

        StrSql = "SELECT CodProd,Qtde,ValorProd FROM tb_Venda_Produto where CodVenda=" & codvenda
        RecVenda.Open StrSql, vgCon, 1, 3

        If Not RecVenda.EOF Then
            Do While Not RecVenda.EOF
                '=== Pegar produtos ==========
                StrSql = "SELECT NomeProd FROM tb_Produto where CodProd=" & RecVenda!CodProd
                RecProd.Open StrSql, vgCon, 1, 3
        
                nomeprod = nomeprod & RecProd!nomeprod & Chr(13) & Chr(13)
                qtdeprod = qtdeprod & FormataNum(RecVenda!qtde) & Chr(13) & Chr(13)
                valorprod = valorprod & FormataMoeda(RecVenda!valorprod) & Chr(13) & Chr(13)
                subtotalvenda = FormataMoeda(CCur(subtotalvenda) + CCur(RecVenda!valorprod))
                
                RecProd.Close
                RecVenda.MoveNext
            Loop
        Else
            nomeprod = ""
            qtdeprod = ""
            valorprod = ""
        End If
        
        RecVenda.Close
        
        vgCon.Execute "INSERT INTO tb_Auxiliar " & _
        "(campo01,campo02,campo03,campo04,campo05,campo06,campo07,campo08,campo09,campo10) " & _
        "VALUES ('" & cliente & "','" & vendedor & "','" & datavenda & "','" & nomeprod & "','" & qtdeprod & "','" & valorprod & "','" & subtotalvenda & "','" & desconto & "','" & totalvenda & "','" & juros & "')"

        VLStrLinha = VLStrLinha + 1
    'Loop

    Desconecta

    rptVenda_Recibo.Show
End Sub

Private Sub CmdVendaRec_Click()
    FrmVenda_Inc.Show
End Sub

Private Sub CmdIncluirVenda_Click()
    FrmVenda_Inc_Cli.Show
End Sub

Private Sub CmdPesqVenda_Click()
    Screen.MousePointer = vbHourglass

    Dim VLStrOrder As String

    Conecta

    StrSql = "SELECT V.CodVenda,V.CodVendedor,V.CodCli,V.CodCred,V.DtVenda,V.TipoVenda,V.Desconto," & _
             "V.TotalVenda,V.TipoPagto FROM tb_Venda as V WHERE 0=0"

    '====== PESQUISAR POR CLIENTE ==========
    If TxtCliVend.Text <> "" Then
        'StrSql = StrSql + " and C.Nome like '%" & TxtCliVend.Text & "%'"
        StrSql = Trim(Mid(StrSql, 1, InStr(StrSql, "FROM") - 1)) & ",C.Nome " & Trim(Mid(StrSql, InStr(StrSql, "FROM"), InStr(StrSql, "WHERE") - InStr(StrSql, "FROM"))) & ",tb_Cliente as C " & Trim(Mid(StrSql, InStr(StrSql, "WHERE"))) & " and C.CodCli=V.CodCli and C.Nome like '%" & TxtCliVend.Text & "%'"
        VLStrOrder = VLStrOrder + "C.Nome,"
    End If

    '====== PESQUISAR POR DATA DA VENDA ==========
    If (TxtDtVenda1.Text <> "" And TxtDtVenda1.Text <> "__/__/____") And (TxtDtVenda2.Text <> "" And TxtDtVenda2.Text <> "__/__/____") Then
        StrSql = StrSql + " and V.DtVenda >=#" & FormataDataUS(TxtDtVenda1.Text) & "# and V.DtVenda <= #" & FormataDataUS(TxtDtVenda2.Text) & "#"
        VLStrOrder = VLStrOrder + "V.DtVenda desc,"

    ElseIf (TxtDtVenda1.Text <> "" And TxtDtVenda1.Text <> "__/__/____") And (TxtDtVenda2.Text = "" Or TxtDtVenda2.Text = "__/__/____") Then
        StrSql = StrSql + " and V.DtVenda =#" & FormataDataUS(TxtDtVenda1.Text) & "#"
        VLStrOrder = VLStrOrder + "V.DtVenda desc,"

    ElseIf (TxtDtVenda1.Text = "" Or TxtDtVenda1.Text = "__/__/____") And (TxtDtVenda2.Text <> "" And TxtDtVenda2.Text <> "__/__/____") Then
        StrSql = StrSql + " and V.DtVenda =#" & FormataDataUS(TxtDtVenda2.Text) & "#"
        VLStrOrder = VLStrOrder + "V.DtVenda desc,"
    End If

    '====== PESQUISAR POR TIPO VENDA ==========
    If CboTipoVenda.Text <> "" Then
        StrSql = StrSql + " and V.TipoVenda='" & CboTipoVenda.Text & "'"
        VLStrOrder = VLStrOrder + "V.TipoVenda,"
    End If

    '====== PESQUISAR POR VENDEDOR ==========
    If TxtVendedor.Text <> "" Then
        'StrSql = StrSql + " and VR.Nome like '%" & TxtVendedor.Text & "%'"
        StrSql = Trim(Mid(StrSql, 1, InStr(StrSql, "FROM") - 1)) & ",VR.Nome " & Trim(Mid(StrSql, InStr(StrSql, "FROM"), InStr(StrSql, "WHERE") - InStr(StrSql, "FROM"))) & ",tb_Vendedor as VR " & Trim(Mid(StrSql, InStr(StrSql, "WHERE"))) & " and VR.CodVendedor=V.CodVendedor and VR.Nome like '%" & TxtVendedor.Text & "%'"
        VLStrOrder = VLStrOrder + "VR.Nome,"
    End If

    '====== ORDENAR PESQUISA ======================
    If VLStrOrder <> "" Then
        VLStrOrder = Mid(VLStrOrder, 1, Len(VLStrOrder) - 1)
        StrSql = StrSql + " order by " & VLStrOrder
    Else
        StrSql = StrSql + " order by V.DtVenda desc,V.CodVenda desc"
    End If

    VLStrOrder = ""

    RecPesq.Open StrSql, vgCon, 1, 3

    Call MontaGridVenda

    Desconecta

    Screen.MousePointer = vbNormal

End Sub

Private Sub GridVenda_Click(ByVal Col As Long, ByVal Row As Long)
    GridVenda.Row = Row
    GridVenda.Col = 9
    If GridVenda.Text <> "" And GridVenda.Text <> "CodVenda" Then
        VGIntCodVenda = GridVenda.Text

        CmdExcluirVenda.Enabled = True
        'CmdImprimirCarne.Enabled = True
        CmdImprimirProp.Enabled = True
        CmdImprimirRecibo.Enabled = True
        CmdImprimirVenda.Enabled = True
    Else
        CmdExcluirVenda.Enabled = False
        'CmdImprimirCarne.Enabled = False
        CmdImprimirProp.Enabled = False
        CmdImprimirRecibo.Enabled = False
        CmdImprimirVenda.Enabled = False
    End If

    GridVenda.Row = Row
    GridVenda.Col = 8
    If GridVenda.Text = "Carnê" Then
       CmdImprimirCarne.Enabled = True
    Else
       CmdImprimirCarne.Enabled = False
    End If
End Sub

Private Sub GridVenda_DblClick(ByVal Col As Long, ByVal Row As Long)
    GridVenda.Row = Row
    GridVenda.Col = 9
    If GridVenda.Text <> "" And GridVenda.Text <> "CodVenda" Then
        VGIntCodVenda = GridVenda.Text
        FrmResumo_Venda.Show
    End If
End Sub

Sub MontaGridVenda()
    Dim VLIntLinha As Long

    If RecPesq.EOF Then
        LblNumTotalVend.Caption = "Nenhuma venda encontrada."

        VPStrBox = MsgBox("Pesquisa sem resultados.", vbInformation, "Pró Vendas 2004 - Informação")
        GridVenda.Refresh
        GridVenda.MaxRows = 0

        CmdExcluirVenda.Enabled = False
        CmdImprimirVenda.Enabled = False

    Else
        Dim RecCli As New ADODB.Recordset
        Dim RecVend As New ADODB.Recordset
        Dim RecCred As New ADODB.Recordset
        
        VLIntLinha = 1
        GridVenda.MaxRows = VLIntLinha

        Do While Not RecPesq.EOF

            GridVenda.Row = VLIntLinha
            GridVenda.Lock = True

            'Cliente
            GridVenda.Col = 1
            
            StrSql = "Select Nome from tb_Cliente where CodCli=" & RecPesq!CodCli
            RecCli.Open StrSql, vgCon, 1, 3
            
            If Not RecCli.EOF Then
                GridVenda.Text = VerificaNulo(RecCli!nome)
            Else
                GridVenda.Text = ""
            End If
            GridVenda.Lock = True

            'Vendedor
            GridVenda.Col = 2
            
            StrSql = "Select Nome from tb_Vendedor where CodVendedor=" & RecPesq!CodVendedor
            RecVend.Open StrSql, vgCon, 1, 3
            
            If Not RecVend.EOF Then
                GridVenda.Text = VerificaNulo(RecVend!nome)
            Else
                GridVenda.Text = ""
            End If
            GridVenda.Lock = True

            'Data venda
            GridVenda.Col = 3
            GridVenda.Text = FormataData(RecPesq!dtvenda)
            GridVenda.Lock = True

            'Valor venda
            GridVenda.Col = 4
            GridVenda.Text = FormataMoeda(VerificaNulo(RecPesq!totalvenda))
            GridVenda.Lock = True

            'Desconto
            GridVenda.Col = 5
            If RecPesq!desconto <> "" And IsNull(RecPesq!desconto) = False Then
                GridVenda.Text = FormataNum(RecPesq!desconto) & "%"
            Else
                GridVenda.Text = ""
            End If
            GridVenda.Lock = True

            'Juros
            GridVenda.Col = 6
            StrSql = "Select Juros from tb_Crediario where CodCred=" & RecPesq!CodCred
            RecCred.Open StrSql, vgCon, 1, 3
            
            If Not RecCred.EOF Then
                If RecCred!juros <> "" And IsNull(RecCred!juros) = False Then
                    GridVenda.Text = FormataNum(RecCred!juros) & "%"
                Else
                    GridVenda.Text = ""
                End If
            Else
                GridVenda.Text = ""
            End If
            GridVenda.Lock = True

            'Tipo Venda
            GridVenda.Col = 7
            GridVenda.Text = VerificaNulo(RecPesq!tipovenda)
            GridVenda.Lock = True

            'Tipo pagto
            GridVenda.Col = 8
            GridVenda.Text = VerificaNulo(RecPesq!TipoPagto)
            GridVenda.Lock = True

            'CodVenda
            GridVenda.Col = 9
            GridVenda.Text = Val(RecPesq!codvenda)
            GridVenda.Lock = True

            VLIntLinha = VLIntLinha + 1

            GridVenda.MaxRows = GridVenda.MaxRows + 1
                        
            RecCli.Close
            RecVend.Close
            RecCred.Close
            RecPesq.MoveNext
         Loop

         '===== CONTAGEM DE VENDAS PESQUISADOS =========
         GridVenda.MaxRows = GridVenda.MaxRows - 1

         If GridVenda.MaxRows = 1 Then
            LblNumTotalVend.Caption = FormataNum(GridVenda.MaxRows) & " venda encontrada."
         Else
            LblNumTotalVend.Caption = FormataNum(GridVenda.MaxRows) & " vendas encontradas."
         End If
         '================================================

         CmdImprimirVenda.Enabled = True
    End If

End Sub

Sub MontaCboTipoVenda()
    Conecta

    Dim RecTipo As New ADODB.Recordset
    StrSql = "Select distinct TipoVenda From tb_Venda order by TipoVenda"
    RecTipo.Open StrSql, vgCon, 1, 3
    
    CboTipoVenda.Clear
    
    CboTipoVenda.AddItem ("")
    
    Do While Not RecTipo.EOF
        CboTipoVenda.AddItem (RecTipo.Fields.Item(0).Value)
        RecTipo.MoveNext
    Loop
    Desconecta
End Sub
'========================================================================
'========================================================================




'========================================================================
'                            CAIXA
'========================================================================

Private Sub CmdLimparCx_Click()
    TxtDtMovCx1.Text = "__/__/____"
    TxtDtMovCx2.Text = "__/__/____"
    CboTipoPagtoCx.ListIndex = 0
    GridCaixa.MaxRows = 0
    LblNumTotalCx.Caption = "Nenhum movimento de caixa encontrado."
    
    CmdAlterarCx.Enabled = False
    CmdExcluirCx.Enabled = False
    CmdImprimirCx.Enabled = False
End Sub

Private Sub CmdExcluirCx_Click()
    VPStrResponse = MsgBox("Deseja excluir este movimento do caixa?", vbYesNo, "Pró Vendas 2004 - Informação")

    If VPStrResponse = vbYes Then
        Conecta
        vgCon.Execute ("DELETE FROM tb_Caixa WHERE CodCx=" & VGIntCodCx)
        Desconecta

        FrmPrincipal.CmdPesqCx.Value = True
    End If
End Sub

Private Sub CmdImprimirCx_Click()
    Screen.MousePointer = vbHourglass

    Dim desc As String
    Dim datamov As String
    Dim tipomov As String
    Dim cred As String
    Dim deb As String

    Dim VLStrLinha As String

    VLStrLinha = 1

    Conecta

    Do While VLStrLinha <= GridCaixa.MaxRows

        GridCaixa.Col = 2
        GridCaixa.Row = VLStrLinha
        desc = GridCaixa.Text

        GridCaixa.Col = 3
        GridCaixa.Row = VLStrLinha
        datamov = GridCaixa.Text

        GridCaixa.Col = 4
        GridCaixa.Row = VLStrLinha
        tipomov = GridCaixa.Text

        GridCaixa.Col = 5
        GridCaixa.Row = VLStrLinha
        cred = GridCaixa.Text

        GridCaixa.Col = 6
        GridCaixa.Row = VLStrLinha
        deb = GridCaixa.Text

        vgCon.Execute "INSERT INTO tb_Auxiliar " & _
        "(campo01,campo02,campo03,campo04,campo05) " & _
        "VALUES ('" & desc & "','" & datamov & "','" & tipomov & "','" & cred & "','" & deb & "')"

        VLStrLinha = VLStrLinha + 1
    Loop

    Desconecta

    rptCaixa.Show

End Sub

Private Sub CmdIncluirCx_Click()
    FrmCaixa_Inc.Show
End Sub

Private Sub CmdAlterarCx_Click()
    FrmCaixa_Alt.Show
End Sub

Private Sub CmdPesqCx_Click()
    Screen.MousePointer = vbHourglass

    Dim VLStrOrder As String

    Conecta

    StrSql = "Select * from tb_Caixa as V where 0=0"

    '====== PESQUISAR POR DATA DO MOVIMENTO ==========
    If (TxtDtMovCx1.Text <> "" And TxtDtMovCx1.Text <> "__/__/____") And (TxtDtMovCx2.Text <> "" And TxtDtMovCx2.Text <> "__/__/____") Then
        StrSql = StrSql + " and V.DtMov >=#" & FormataDataUS(TxtDtMovCx1.Text) & "# and V.DtMov <= #" & FormataDataUS(TxtDtMovCx2.Text) & "#"
        VLStrOrder = VLStrOrder + "V.DtMov desc,"

    ElseIf (TxtDtMovCx1.Text <> "" And TxtDtMovCx1.Text <> "__/__/____") And (TxtDtMovCx2.Text = "" Or TxtDtMovCx2.Text = "__/__/____") Then
        StrSql = StrSql + " and V.DtMov =#" & FormataDataUS(TxtDtMovCx1.Text) & "#"
        VLStrOrder = VLStrOrder + "V.DtMov desc,"

    ElseIf (TxtDtMovCx1.Text = "" Or TxtDtMovCx1.Text = "__/__/____") And (TxtDtMovCx2.Text <> "" And TxtDtMovCx2.Text <> "__/__/____") Then
        StrSql = StrSql + " and V.DtMov =#" & FormataDataUS(TxtDtMovCx2.Text) & "#"
        VLStrOrder = VLStrOrder + "V.DtMov desc,"
    End If

    '====== PESQUISAR POR TIPO DE PAGAMENTO ==========
    If CboTipoPagtoCx.Text <> "" Then
        StrSql = StrSql + " and TipoPagto='" & CboTipoPagtoCx.Text & "'"
        VLStrOrder = VLStrOrder + "TipoPagto,"
    End If

    '====== ORDENAR PESQUISA ======================
    If VLStrOrder <> "" Then
        VLStrOrder = Mid(VLStrOrder, 1, Len(VLStrOrder) - 1)
        StrSql = StrSql + " order by " & VLStrOrder
    Else
        StrSql = StrSql + " order by DtMov desc"
    End If

    VLStrOrder = ""

    RecPesq.Open StrSql, vgCon, 1, 3

    Call MontaGridCaixa

    Desconecta

    Screen.MousePointer = vbNormal

End Sub

Private Sub GridCaixa_Click(ByVal Col As Long, ByVal Row As Long)
    Dim VLStrLinha As Integer

    GridCaixa.Row = Row
    GridCaixa.Col = 7
    If GridCaixa.Text <> "" And GridCaixa.Text <> "CodCx" Then
        VGIntCodCx = GridCaixa.Text
        CmdAlterarCx.Enabled = True
        CmdExcluirCx.Enabled = True
    Else
        CmdAlterarCx.Enabled = False
        CmdExcluirCx.Enabled = False
    End If
End Sub

Private Sub GridCaixa_DblClick(ByVal Col As Long, ByVal Row As Long)
    GridCaixa.Row = Row
    GridCaixa.Col = 7
    If GridCaixa.Text <> "" And GridCaixa.Text <> "CodCx" Then
        VGIntCodCx = GridCaixa.Text
        FrmResumo_Caixa.Show
    End If
End Sub

Sub MontaGridCaixa()
    Dim VLIntCodCx As Long
    Dim VLIntLinha As Long
    Dim VLIntCredito As Long
    Dim VLIntDebito As Long
    Dim VLIntVenda As Long
    Dim VLStrCorVermelho  As String

    VLStrCorVermelho = &HC0&

    If RecPesq.EOF Then
        LblNumTotalCx.Caption = "Nenhum movimento de caixa encontrado."

        VPStrBox = MsgBox("Pesquisa sem resultados.", vbInformation, "Pró Vendas 2004 - Informação")
        GridCaixa.Refresh
        GridCaixa.MaxRows = 0

        CmdAlterarCx.Enabled = False
        CmdExcluirCx.Enabled = False
        CmdImprimirCx.Enabled = False

    Else

        VLIntLinha = 1
        GridCaixa.MaxRows = VLIntLinha

        Do While Not RecPesq.EOF

            GridCaixa.Row = VLIntLinha
            GridCaixa.Lock = True

            'Cod. Venda
            GridCaixa.Col = 1
            GridCaixa.Text = FormataNum(RecPesq!codvenda)
            GridCaixa.Lock = True

            'Descrição
            GridCaixa.Col = 2
            GridCaixa.Text = VerificaNulo(RecPesq!Descricao)
            GridCaixa.Lock = True

            'Data Movimento
            GridCaixa.Col = 3
            GridCaixa.Text = FormataData(RecPesq!dtmov)
            GridCaixa.Lock = True

            'Tipo Movimento
            GridCaixa.Col = 4
            GridCaixa.Text = VerificaNulo(RecPesq!tipomov)
            GridCaixa.Lock = True

            GridCaixa.Col = 5
            If RecPesq!tipovalor = "credito" Then
                GridCaixa.Text = FormataMoeda(VerificaNulo(RecPesq!valor))
                VLIntCredito = VLIntCredito + CCur(GridCaixa.Text)

                If RecPesq!codvenda <> 0 And IsNull(RecPesq!codvenda) = False Then
                    VLIntVenda = VLIntVenda + CCur(GridCaixa.Text)
                End If
            Else
                GridCaixa.Text = ""
            End If
            GridCaixa.Lock = True

            'Débito
            GridCaixa.Col = 6
            If RecPesq!tipovalor = "debito" Then
                GridCaixa.Text = FormataMoeda(VerificaNulo(RecPesq!valor))
                VLIntDebito = VLIntDebito + CCur(GridCaixa.Text)
            Else
                GridCaixa.Text = ""
            End If
            GridCaixa.Lock = True

            'CodCx
            GridCaixa.Col = 7
            GridCaixa.Text = Val(RecPesq!codcx)
            GridCaixa.Lock = True

            VLIntLinha = VLIntLinha + 1

            GridCaixa.MaxRows = GridCaixa.MaxRows + 1
            RecPesq.MoveNext
         Loop

         GridCaixa.Row = GridCaixa.MaxRows
         GridCaixa.Col = 1
         GridCaixa.Lock = True
         GridCaixa.Col = 2
         GridCaixa.Lock = True
         GridCaixa.Col = 3
         GridCaixa.Lock = True
         GridCaixa.Col = 4
         GridCaixa.Lock = True
         GridCaixa.Col = 5
         GridCaixa.Lock = True
         GridCaixa.Col = 6
         GridCaixa.Lock = True


         GridCaixa.MaxRows = GridCaixa.MaxRows + 1
         GridCaixa.Row = GridCaixa.MaxRows

         GridCaixa.Col = 1
         GridCaixa.Lock = True
         GridCaixa.Col = 2
         GridCaixa.Text = "TOTAL VENDA DO DIA:"
         GridCaixa.Font.Bold = True
         GridCaixa.Lock = True
         GridCaixa.Col = 3
         GridCaixa.Text = FormataMoeda(VLIntVenda)
         GridCaixa.Font.Bold = True
         GridCaixa.Lock = True
         GridCaixa.Col = 4
         GridCaixa.Lock = True
         GridCaixa.Col = 5
         GridCaixa.Lock = True
         GridCaixa.Col = 6
         GridCaixa.Lock = True

         GridCaixa.MaxRows = GridCaixa.MaxRows + 1
         GridCaixa.Row = GridCaixa.MaxRows

         GridCaixa.Col = 1
         GridCaixa.Lock = True
         GridCaixa.Col = 2
         GridCaixa.Text = "TOTAL CRÉDITO:"
         GridCaixa.Font.Bold = True
         GridCaixa.Lock = True
         GridCaixa.Col = 3
         GridCaixa.Text = FormataMoeda(VLIntCredito)
         GridCaixa.Font.Bold = True
         GridCaixa.Lock = True
         GridCaixa.Col = 4
         GridCaixa.Lock = True
         GridCaixa.Col = 5
         GridCaixa.Lock = True
         GridCaixa.Col = 6
         GridCaixa.Lock = True


         GridCaixa.MaxRows = GridCaixa.MaxRows + 1
         GridCaixa.Row = GridCaixa.MaxRows

         GridCaixa.Col = 1
         GridCaixa.Lock = True
         GridCaixa.Col = 2
         GridCaixa.Text = "TOTAL DÉBITO:"
         GridCaixa.Font.Bold = True
         GridCaixa.Lock = True
         GridCaixa.Col = 3
         GridCaixa.Text = FormataMoeda(VLIntDebito)
         GridCaixa.Font.Bold = True
         GridCaixa.Lock = True
         GridCaixa.Col = 4
         GridCaixa.Lock = True
         GridCaixa.Col = 5
         GridCaixa.Lock = True
         GridCaixa.Col = 6
         GridCaixa.Lock = True


         GridCaixa.MaxRows = GridCaixa.MaxRows + 1
         GridCaixa.Row = GridCaixa.MaxRows

         GridCaixa.Col = 1
         GridCaixa.Lock = True
         GridCaixa.Col = 2
         GridCaixa.Text = "TOTAL MOVIMENTO DO DIA:"
         GridCaixa.Font.Bold = True
         GridCaixa.Lock = True
         GridCaixa.Col = 3
         GridCaixa.Text = FormataMoeda(VLIntCredito - VLIntDebito)
         If InStr(GridCaixa.Text, "-") <> 0 Then
            GridCaixa.ForeColor = VLStrCorVermelho
         End If
         GridCaixa.Font.Bold = True
         GridCaixa.Lock = True
         GridCaixa.Col = 4
         GridCaixa.Lock = True
         GridCaixa.Col = 5
         GridCaixa.Lock = True
         GridCaixa.Col = 6
         GridCaixa.Lock = True


         '===== CONTAGEM DE MOVIMENTOS PESQUISADOS =========
         If (GridCaixa.MaxRows - 5) = 1 Then
            LblNumTotalCx.Caption = FormataNum((GridCaixa.MaxRows - 5)) & " movimento de caixa encontrado."
         Else
            LblNumTotalCx.Caption = FormataNum((GridCaixa.MaxRows - 5)) & " movimentos de caixa encontrados."
         End If
         '================================================

         CmdImprimirCx.Enabled = True
    End If

End Sub

Private Sub TxtDtMovCx1_GotFocus()
    If TxtDtMovCx1.Text = "__/__/____" Then
        TxtDtMovCx1.Text = ""
    Else
        TxtDtMovCx1.SelStart = 0
        TxtDtMovCx1.SelLength = Len(TxtDtMovCx1.Text)
    End If
End Sub

Private Sub TxtDtMovCx1_KeyPress(KeyAscii As Integer)
    '=== Só aceita números ===
    If (KeyAscii < 47 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtDtMovCx1_LostFocus()
    Dim VLStrData As String

    If TxtDtMovCx1.Text <> "" Then
        VLStrData = VerificaData(TxtDtMovCx1.Text)

        If VGStrDataErro = "sim" Then
            TxtDtMovCx1.SetFocus
        Else
            TxtDtMovCx1.Text = VLStrData
        End If

        VGStrDataErro = ""
    Else
        TxtDtMovCx1.Text = "__/__/____"
    End If
End Sub

Private Sub TxtDtMovCx2_GotFocus()
    If TxtDtMovCx2.Text = "__/__/____" Then
        TxtDtMovCx2.Text = ""
    Else
        TxtDtMovCx2.SelStart = 0
        TxtDtMovCx2.SelLength = Len(TxtDtMovCx2.Text)
    End If
End Sub

Private Sub TxtDtMovCx2_KeyPress(KeyAscii As Integer)
    '=== Só aceita números ===
    If (KeyAscii < 47 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtDtMovCx2_LostFocus()
    Dim VLStrData As String

    If TxtDtMovCx2.Text <> "" Then
        VLStrData = VerificaData(TxtDtMovCx2.Text)

        If VGStrDataErro = "sim" Then
            TxtDtMovCx2.SetFocus
        Else
            TxtDtMovCx2.Text = VLStrData
        End If

        VGStrDataErro = ""
    Else
        TxtDtMovCx2.Text = "__/__/____"
    End If
End Sub

Sub MontaCboTipoPagtoCX()
    Conecta

    Dim RecTipo As New ADODB.Recordset
    StrSql = "Select distinct TipoPagto From tb_Caixa order by TipoPagto"
    RecTipo.Open StrSql, vgCon, 1, 3

    CboTipoPagtoCx.Clear
    CboTipoPagtoCx.AddItem ("")
    Do While Not RecTipo.EOF
        CboTipoPagtoCx.AddItem (RecTipo!TipoPagto)
        RecTipo.MoveNext
    Loop
    Desconecta
End Sub
'========================================================================
'========================================================================




'========================================================================
'                            CONTAS A PAGAR
'========================================================================
Private Sub CmdAlterarAPagar_Click()
    FrmCaixa_APagar_Alt.Show
End Sub

Private Sub CmdLimparAPagar_Click()
    TxtDtAPagar1.Text = "__/__/____"
    TxtDtAPagar2.Text = "__/__/____"
    TxtDescrAPagar.Text = ""
    GridAPagar.MaxRows = 0
    LblNumTotalPag.Caption = "Nenhum pagamento encontrado."
    
    CmdAlterarAPagar.Enabled = False
    CmdExcluirAPagar.Enabled = False
    CmdImprimirAPagar.Enabled = False
    CmdBaixarAPagar.Enabled = False
End Sub

Private Sub TxtDtAPagar1_GotFocus()
    If TxtDtAPagar1.Text = "__/__/____" Then
        TxtDtAPagar1.Text = ""
    Else
        TxtDtAPagar1.SelStart = 0
        TxtDtAPagar1.SelLength = Len(TxtDtAPagar1.Text)
    End If
End Sub

Private Sub TxtDtAPagar1_KeyPress(KeyAscii As Integer)
    '=== Só aceita números ===
    If (KeyAscii < 47 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtDtAPagar1_LostFocus()
    Dim VLStrData As String

    If TxtDtAPagar1.Text <> "" Then
        VLStrData = VerificaData(TxtDtAPagar1.Text)

        If VGStrDataErro = "sim" Then
            TxtDtAPagar1.SetFocus
        Else
            TxtDtAPagar1.Text = VLStrData
        End If

        VGStrDataErro = ""
    Else
        TxtDtAPagar1.Text = "__/__/____"
    End If
End Sub

Private Sub TxtDtAPagar2_GotFocus()
    If TxtDtAPagar2.Text = "__/__/____" Then
        TxtDtAPagar2.Text = ""
    Else
        TxtDtAPagar2.SelStart = 0
        TxtDtAPagar2.SelLength = Len(TxtDtAPagar2.Text)
    End If
End Sub

Private Sub TxtDtAPagar2_KeyPress(KeyAscii As Integer)
    '=== Só aceita números ===
    If (KeyAscii < 47 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtDtAPagar2_LostFocus()
    Dim VLStrData As String

    If TxtDtAPagar2.Text <> "" Then
        VLStrData = VerificaData(TxtDtAPagar2.Text)

        If VGStrDataErro = "sim" Then
            TxtDtAPagar2.SetFocus
        Else
            TxtDtAPagar2.Text = VLStrData
        End If

        VGStrDataErro = ""
    Else
        TxtDtAPagar2.Text = "__/__/____"
    End If
End Sub

Private Sub CmdExcluirAPagar_Click()
    If VGStrStatusPagto = "Pago" Then
        
        VPStrResponse = MsgBox("Este pagamento já foi efetuado, sua exclusão" & Chr(13) & "apagará todas as informações deste pagamento." & Chr(13) & "Deseja continuar?", vbYesNo, "Pró Vendas 2004 - Informação")
        
        If VPStrResponse = vbYes Then
            Conecta
            vgCon.Execute ("DELETE FROM tb_ContaPagar_Pagto WHERE CodCPag=" & VGIntCodPagar)
            vgCon.Execute ("DELETE FROM tb_ContaPagar WHERE CodCPag=" & VGIntCodPagar)
            Desconecta
            
            CmdPesqPag.Value = True
            
            CmdAlterarAPagar.Enabled = False
            CmdExcluirAPagar.Enabled = False
            CmdBaixarAPagar.Enabled = False
        End If
    Else
        VPStrResponse = MsgBox("Deseja excluir este pagamento?", vbYesNo, "Pró Vendas 2004 - Informação")
        
        If VPStrResponse = vbYes Then
            Conecta
            vgCon.Execute ("DELETE FROM tb_ContaPagar WHERE CodCPag=" & VGIntCodPagar)
            Desconecta
            
            FrmPrincipal.CmdPesqAPagar.Value = True
            
            CmdAlterarAPagar.Enabled = False
            CmdExcluirAPagar.Enabled = False
            CmdBaixarAPagar.Enabled = False
        End If
    End If
End Sub

Private Sub CmdImprimirAPagar_Click()
    Screen.MousePointer = vbHourglass
    
    Dim desc As String
    Dim tipo As String
    Dim venc As String
    Dim valor As String
    Dim status As String
    
    Dim VLStrLinha As String
    
    VLStrLinha = 1
    
    Conecta
    
    Do While VLStrLinha <= GridAPagar.MaxRows
        
        GridAPagar.Col = 1
        GridAPagar.Row = VLStrLinha
        desc = GridAPagar.Text
        
        GridAPagar.Col = 2
        GridAPagar.Row = VLStrLinha
        tipo = GridAPagar.Text
        
        GridAPagar.Col = 3
        GridAPagar.Row = VLStrLinha
        venc = GridAPagar.Text
        
        GridAPagar.Col = 4
        GridAPagar.Row = VLStrLinha
        valor = GridAPagar.Text
        
        GridAPagar.Col = 5
        GridAPagar.Row = VLStrLinha
        status = GridAPagar.Text
        
        vgCon.Execute "INSERT INTO tb_Auxiliar " & _
        "(campo01,campo02,campo03,campo04,campo05) " & _
        "VALUES ('" & desc & "','" & tipo & "','" & venc & "','" & valor & "','" & status & "')"
         
        VLStrLinha = VLStrLinha + 1
    Loop
    
    Desconecta
        
    rptCaixa_APagar.Show

End Sub

Private Sub CmdIncluirAPagar_Click()
    FrmCaixa_APagar_Inc.Show
End Sub

Private Sub CmdPesqAPagar_Click()
    Screen.MousePointer = vbHourglass
    
    Conecta
    
    If OptPagoSimAPagar.Value = True Then
        StrSql = "Select * from tb_ContaPagar where Pago='sim'"
    ElseIf OptPagoNaoAPagar.Value = True Then
        StrSql = "Select * from tb_ContaPagar where Pago='não'"
    ElseIf OptPagoTodosAPagar.Value = True Then
        StrSql = "Select * from tb_ContaPagar where 0=0"
    End If
    
    '====== PESQUISAR POR DATA DO VENCIMENTO ==========
    If (TxtDtAPagar1.Text <> "" And TxtDtAPagar1.Text <> "__/__/____") And (TxtDtAPagar2.Text <> "" And TxtDtAPagar2.Text <> "__/__/____") Then
        StrSql = StrSql + " and Vencimento >=#" & FormataDataUS(TxtDtAPagar1.Text) & "# and Vencimento <= #" & FormataDataUS(TxtDtAPagar2.Text) & "#"
    
    ElseIf (TxtDtAPagar1.Text <> "" And TxtDtAPagar1.Text <> "__/__/____") And (TxtDtAPagar2.Text = "" Or TxtDtAPagar2.Text = "__/__/____") Then
        StrSql = StrSql + " and Vencimento =#" & FormataDataUS(TxtDtAPagar1.Text) & "#"
    
    ElseIf (TxtDtAPagar1.Text = "" Or TxtDtAPagar1.Text = "__/__/____") And (TxtDtAPagar2.Text <> "" And TxtDtAPagar2.Text <> "__/__/____") Then
        StrSql = StrSql + " and Vencimento =#" & FormataDataUS(TxtDtAPagar2.Text) & "#"
    
    End If
            
    '====== PESQUISAR POR DESCRIÇÃO ==========
    If TxtDescrAPagar.Text <> "" Then
        StrSql = StrSql + " and Descricao like '%" & TxtDescrAPagar.Text & "%'"
    End If
            
    '====== ORDENAR PESQUISA ======================
        StrSql = StrSql + " order by Vencimento desc"
    
    RecPesq.Open StrSql, vgCon, 1, 3
            
    Call MontaGridAPagar
        
    Desconecta
    
    Screen.MousePointer = vbNormal
End Sub

Sub MontaGridAPagar()
    Dim VLIntCodValor As Double
    Dim VLIntLinha As Long
    
    If RecPesq.EOF Then
        LblNumTotalPag.Caption = "Nenhum pagamento encontrado."
        
        VPStrBox = MsgBox("Pesquisa sem resultados.", vbInformation, "Pró Vendas 2004 - Informação")
        GridAPagar.Refresh
        GridAPagar.MaxRows = 0
        
        CmdAlterarAPagar.Enabled = False
        CmdExcluirAPagar.Enabled = False
        CmdImprimirAPagar.Enabled = False
        CmdBaixarAPagar.Enabled = False
    
    Else
    
        VLIntLinha = 1
        GridAPagar.MaxRows = VLIntLinha
         
        Do While Not RecPesq.EOF
                 
            GridAPagar.Row = VLIntLinha
            GridAPagar.Lock = True
            
            'Descrição
            GridAPagar.Col = 1
            GridAPagar.TypeMaxEditLen = 255
            GridAPagar.Text = VerificaNulo(RecPesq!Descricao)
            GridAPagar.Lock = True
            
            'Tipo
            GridAPagar.Col = 2
            GridAPagar.Text = VerificaNulo(RecPesq!tipo)
            GridAPagar.Lock = True
            
            'Vencimento
            GridAPagar.Col = 3
            GridAPagar.Text = FormataData(RecPesq!vencimento)
            GridAPagar.Lock = True
            
            'Valor
            GridAPagar.Col = 4
            GridAPagar.Text = FormataMoeda(VerificaNulo(RecPesq!valor))
            If RecPesq!pago = "não" Then
                VLIntValor = VLIntValor + CCur(GridAPagar.Text)
            End If
            GridAPagar.Lock = True
            
            'Status
            GridAPagar.Col = 5
            If RecPesq!pago = "sim" Then
                GridAPagar.Text = "Pago"
            ElseIf RecPesq!pago = "não" Then
                GridAPagar.Text = "Em aberto"
            End If
            GridAPagar.Lock = True
            
            'CodCPag
            GridAPagar.Col = 6
            GridAPagar.Text = Val(RecPesq!CodCPag)
            GridAPagar.Lock = True
            
            VLIntLinha = VLIntLinha + 1
            
            GridAPagar.MaxRows = GridAPagar.MaxRows + 1
            RecPesq.MoveNext
         Loop
         
         'trava linhas em branco para escrever resumo depois
         GridAPagar.Row = GridAPagar.MaxRows
         GridAPagar.Col = 1
         GridAPagar.Lock = True
         GridAPagar.Col = 2
         GridAPagar.Lock = True
         GridAPagar.Col = 3
         GridAPagar.Lock = True
         GridAPagar.Col = 4
         GridAPagar.Lock = True
         GridAPagar.Col = 5
         GridAPagar.Lock = True
         GridAPagar.Col = 6
         GridAPagar.Lock = True
         
         
         GridAPagar.MaxRows = GridAPagar.MaxRows + 1
         GridAPagar.Row = GridAPagar.MaxRows
         
         'escreve resumo das contas
         GridAPagar.Col = 1
         GridAPagar.Text = "TOTAL À PAGAR:"
         GridAPagar.Font.Bold = True
         GridAPagar.Lock = True
         GridAPagar.Col = 2
         GridAPagar.Text = FormataMoeda(VLIntValor)
         GridAPagar.Font.Bold = True
         GridAPagar.Lock = True
         GridAPagar.Col = 3
         GridAPagar.Lock = True
         GridAPagar.Col = 4
         GridAPagar.Lock = True
         GridAPagar.Col = 5
         GridAPagar.Lock = True
         GridAPagar.Col = 6
         GridAPagar.Lock = True
         
         '===== CONTAGEM DE PAGAMENTOS PESQUISADOS =========
         If (GridAPagar.MaxRows - 2) = 1 Then
            LblNumTotalPag.Caption = FormataNum((GridAPagar.MaxRows - 2)) & " pagamento encontrado."
         Else
            LblNumTotalPag.Caption = FormataNum((GridAPagar.MaxRows - 2)) & " pagamentos encontrados."
         End If
         '================================================
         
         CmdImprimirAPagar.Enabled = True
    End If

End Sub

Private Sub GridAPagar_Click(ByVal Col As Long, ByVal Row As Long)
    GridAPagar.Row = Row
    GridAPagar.Col = 6
    If GridAPagar.Text <> "CodCPag" And GridAPagar.Text <> "" Then
        VGIntCodPagar = GridAPagar.Text
        CmdAlterarAPagar.Enabled = True
        CmdExcluirAPagar.Enabled = True
        
        GridAPagar.Row = Row
        GridAPagar.Col = 5
        If GridAPagar.Text = "Em aberto" Then
            CmdBaixarAPagar.Enabled = True
        Else
            CmdBaixarAPagar.Enabled = False
        End If
    Else
        VGIntCodPagar = 0
        CmdAlterarAPagar.Enabled = False
        CmdExcluirAPagar.Enabled = False
        CmdBaixarAPagar.Enabled = False
    End If
    
    GridAPagar.Row = Row
    GridAPagar.Col = 5
    If GridAPagar.Text <> "Status" And GridAPagar.Text <> "" Then
        VGStrStatusPagto = GridAPagar.Text
        CmdAlterarAPagar.Enabled = True
        CmdExcluirAPagar.Enabled = True
        
        If GridAPagar.Text = "Em aberto" Then
            CmdBaixarAPagar.Enabled = True
        Else
            CmdBaixarAPagar.Enabled = False
        End If
    Else
        VGStrStatusPagto = ""
        CmdAlterarAPagar.Enabled = False
        CmdExcluirAPagar.Enabled = False
        CmdBaixarAPagar.Enabled = False
    End If
End Sub

Private Sub GridAPagar_DblClick(ByVal Col As Long, ByVal Row As Long)
    GridAPagar.Row = Row
    GridAPagar.Col = 6
    If GridAPagar.Text <> "CodCPag" And GridAPagar.Text <> "" Then
        VGIntCodPagar = GridAPagar.Text
        FrmResumo_APagar.Show
    End If
End Sub

Private Sub CmdBaixarAPagar_Click()
    If VGStrStatusPagto = "Pago" Then
        FrmCaixa_APagar_Baixado.Show
    Else
        FrmCaixa_APagar_Baixa.Show
    End If
End Sub
'========================================================================
'========================================================================




'========================================================================
'                            CONTAS A RECEBER
'========================================================================

Private Sub CmdExcluirAReceber_Click()
    If VGStrStatusReceb = "Recebido" Then
        
        If VGStrReceb = "conta" Then
            VPStrResponse = MsgBox("Esta conta já foi recebida, sua exclusão" & Chr(13) & "apagará todas as informações desta conta." & Chr(13) & "Deseja continuar?", vbYesNo, "Pró Vendas 2004 - Informação")
        
        ElseIf VGStrReceb = "parcela" Then
            VPStrResponse = MsgBox("Esta parcela já foi recebida, sua exclusão" & Chr(13) & "apagará todas as informações desta parcela." & Chr(13) & "Deseja continuar?", vbYesNo, "Pró Vendas 2004 - Informação")
        
        End If
        
        
        If VPStrResponse = vbYes Then
            Conecta
            
            If VGStrReceb = "conta" Then
                vgCon.Execute ("DELETE FROM tb_ContaReceber_Recebido WHERE CodCReceb=" & VGIntCodReceber)
                vgCon.Execute ("DELETE FROM tb_ContaReceber WHERE CodCReceb=" & VGIntCodReceber)
            
            ElseIf VGStrReceb = "parcela" Then
                vgCon.Execute ("DELETE FROM tb_Crediario_Parcela WHERE CodParc=" & VGIntCodReceber)
            
            End If
            Desconecta
            
            FrmPrincipal.CmdPesqAReceber.Value = True
            
            CmdAlterarAReceber.Enabled = False
            CmdExcluirAReceber.Enabled = False
            CmdBaixaAReceber.Enabled = False
        End If
    Else
        
        If VGStrReceb = "conta" Then
            VPStrResponse = MsgBox("Deseja excluir este recebimento?", vbYesNo, "Pró Vendas 2004 - Informação")
        
        ElseIf VGStrReceb = "parcela" Then
            VPStrResponse = MsgBox("Deseja excluir esta parcela?", vbYesNo, "Pró Vendas 2004 - Informação")
        
        End If
        
        If VPStrResponse = vbYes Then
            Conecta
            
            If VGStrReceb = "conta" Then
                vgCon.Execute ("DELETE FROM tb_ContaReceber WHERE CodCReceb=" & VGIntCodReceber)
            
            ElseIf VGStrReceb = "parcela" Then
                vgCon.Execute ("DELETE FROM tb_Crediario_Parcela WHERE CodParc=" & VGIntCodReceber)
            
            End If
            
            Desconecta
            
            FrmPrincipal.CmdPesqAReceber.Value = True
            
            CmdAlterarAReceber.Enabled = False
            CmdExcluirAReceber.Enabled = False
            CmdBaixaAReceber.Enabled = False
        End If
    End If
    
    VGStrReceb = ""
    VGStrStatusReceb = ""
End Sub

Private Sub CmdLimparAReceber_Click()
    TxtDtReceb1.Text = "__/__/____"
    TxtDtReceb2.Text = "__/__/____"
    TxtCliReceb.Text = ""
    GridAReceber.MaxRows = 0
    LblNumTotalReceb.Caption = "Nenhum recebimento encontrado."
    
    CmdAlterarAReceber.Enabled = False
    CmdExcluirAReceber.Enabled = False
    CmdImprimirAReceber.Enabled = False
    CmdBaixaAReceber.Enabled = False
End Sub

Private Sub TxtDtReceb1_GotFocus()
    If TxtDtReceb1.Text = "__/__/____" Then
        TxtDtReceb1.Text = ""
    Else
        TxtDtReceb1.SelStart = 0
        TxtDtReceb1.SelLength = Len(TxtDtReceb1.Text)
    End If
End Sub

Private Sub TxtDtReceb1_KeyPress(KeyAscii As Integer)
    '=== Só aceita números ===
    If (KeyAscii < 47 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtDtReceb1_LostFocus()
    Dim VLStrData As String

    If TxtDtReceb1.Text <> "" Then
        VLStrData = VerificaData(TxtDtReceb1.Text)

        If VGStrDataErro = "sim" Then
            TxtDtReceb1.SetFocus
        Else
            TxtDtReceb1.Text = VLStrData
        End If

        VGStrDataErro = ""
    Else
        TxtDtReceb1.Text = "__/__/____"
    End If
End Sub

Private Sub TxtDtReceb2_GotFocus()
    If TxtDtReceb2.Text = "__/__/____" Then
        TxtDtReceb2.Text = ""
    Else
        TxtDtReceb2.SelStart = 0
        TxtDtReceb2.SelLength = Len(TxtDtReceb2.Text)
    End If
End Sub

Private Sub TxtDtReceb2_KeyPress(KeyAscii As Integer)
    '=== Só aceita números ===
    If (KeyAscii < 47 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtDtReceb2_LostFocus()
    Dim VLStrData As String

    If TxtDtReceb2.Text <> "" Then
        VLStrData = VerificaData(TxtDtReceb2.Text)

        If VGStrDataErro = "sim" Then
            TxtDtReceb2.SetFocus
        Else
            TxtDtReceb2.Text = VLStrData
        End If

        VGStrDataErro = ""
    Else
        TxtDtReceb2.Text = "__/__/____"
    End If
End Sub

Private Sub CmdPesqAReceber_Click()

    Screen.MousePointer = vbHourglass
    
    Conecta
    
    If OptRecebSim.Value = True Then
        StrSql = "Select * from tb_Crediario_Parcela as P,tb_Crediario as CR,tb_Cliente as C where C.CodCli=CR.CodCli and CR.CodCred=P.CodCred and Quitado='sim'"
        StrSql2 = "Select * from tb_ContaReceber where Recebido='sim'"
    
    ElseIf OptRecebNao.Value = True Then
        StrSql = "Select * from tb_Crediario_Parcela as P,tb_Crediario as CR,tb_Cliente as C where C.CodCli=CR.CodCli and CR.CodCred=P.CodCred and Quitado='não'"
        StrSql2 = "Select * from tb_ContaReceber where Recebido='não'"
    
    ElseIf OptRecebTodos.Value = True Then
        StrSql = "Select * from tb_Crediario_Parcela as P,tb_Crediario as CR,tb_Cliente as C where C.CodCli=CR.CodCli and CR.CodCred=P.CodCred"
        StrSql2 = "Select * from tb_ContaReceber where 0=0"
    
    End If
    
    '====== PESQUISAR POR DATA DO VENCIMENTO ==========
    If (TxtDtReceb1.Text <> "" And TxtDtReceb1.Text <> "__/__/____") And (TxtDtReceb2.Text <> "" And TxtDtReceb2.Text <> "__/__/____") Then
        StrSql = StrSql + " and P.Vencimento >=#" & FormataDataUS(TxtDtReceb1.Text) & "# and P.Vencimento <= #" & FormataDataUS(TxtDtReceb2.Text) & "#"
        StrSql2 = StrSql2 + " and Vencimento >=#" & FormataDataUS(TxtDtReceb1.Text) & "# and Vencimento <= #" & FormataDataUS(TxtDtReceb2.Text) & "#"
    
    ElseIf (TxtDtReceb1.Text <> "" And TxtDtReceb1.Text <> "__/__/____") And (TxtDtReceb2.Text = "" Or TxtDtReceb2.Text = "__/__/____") Then
        StrSql = StrSql + " and P.Vencimento =#" & FormataDataUS(TxtDtReceb1.Text) & "#"
        StrSql2 = StrSql2 + " and Vencimento =#" & FormataDataUS(TxtDtReceb1.Text) & "#"
    
    ElseIf (TxtDtReceb1.Text = "" Or TxtDtReceb1.Text = "__/__/____") And (TxtDtReceb2.Text <> "" And TxtDtReceb2.Text <> "__/__/____") Then
        StrSql = StrSql + " and P.Vencimento =#" & FormataDataUS(TxtDtReceb2.Text) & "#"
        StrSql2 = StrSql2 + " and Vencimento =#" & FormataDataUS(TxtDtReceb2.Text) & "#"
    
    End If
            
    '====== PESQUISAR POR CLIENTE ==========
    If TxtCliReceb.Text <> "" Then
        StrSql = StrSql + " and C.Nome like '%" & TxtCliReceb.Text & "%'"
        StrSql2 = StrSql2 + " and Descricao like '%" & TxtCliReceb.Text & "%'"
    End If
    
    '====== ORDENAR PESQUISA ======================
    StrSql = StrSql + " order by P.Vencimento desc"
    StrSql2 = StrSql2 + " order by Vencimento desc"
    
    RecPesq.Open StrSql, vgCon, 1, 3
            
    RecPesq2.Open StrSql2, vgCon, 1, 3
    
    Call MontaGridAReceber
        
    Desconecta
    
    Screen.MousePointer = vbNormal
End Sub

Sub MontaGridAReceber()
    Dim RecCred As New ADODB.Recordset
    Dim VLIntCodValor As Double
    Dim VLIntLinha As Long
    
    If RecPesq.EOF And RecPesq2.EOF Then
        LblNumTotalReceb.Caption = "Nenhum recebimento encontrado."
        
        VPStrBox = MsgBox("Pesquisa sem resultados.", vbInformation, "Pró Vendas 2004 - Informação")
        GridAReceber.Refresh
        GridAReceber.MaxRows = 0
        
        CmdImprimirAReceber.Enabled = False
    
    Else
    
        VLIntLinha = 1
        GridAReceber.MaxRows = VLIntLinha
        
        'monta dados da tabela de crediário e parcelas receber
        Do While Not RecPesq.EOF
            StrSql = "SELECT C.Nome,CR.TipoCred FROM tb_Cliente as C, tb_Crediario as CR " & _
                     "WHERE C.CodCli=CR.CodCli AND CR.CodCred=" & RecPesq.Fields.Item(0).Value
            RecCred.Open StrSql, vgCon, 1, 3
            
            GridAReceber.Row = VLIntLinha
            GridAReceber.Lock = True
            
            'Descrição
            GridAReceber.Col = 1
            GridAReceber.TypeMaxEditLen = 255
            GridAReceber.Text = "Parcela de crediário - Cliente: " & VerificaNulo(RecCred.Fields.Item(0).Value)
            GridAReceber.Lock = True
            
            'Tipo
            GridAReceber.Col = 2
            GridAReceber.Text = VerificaNulo(RecCred.Fields.Item(1).Value)
            GridAReceber.Lock = True
            
            'Vencimento
            GridAReceber.Col = 3
            GridAReceber.Text = FormataData(RecPesq.Fields.Item(3).Value)
            GridAReceber.Lock = True
            
            'Valor
            GridAReceber.Col = 4
            GridAReceber.Text = FormataMoeda(VerificaNulo(RecPesq.Fields.Item(4).Value))
            If RecPesq.Fields.Item(5).Value = "não" Then
                VLIntValor = VLIntValor + CCur(GridAReceber.Text)
            End If
            GridAReceber.Lock = True
            
            'Status
            GridAReceber.Col = 5
            If RecPesq.Fields.Item(5).Value = "sim" Then
                GridAReceber.Text = "Recebido"
            ElseIf RecPesq.Fields.Item(5).Value = "não" Then
                GridAReceber.Text = "A receber"
            End If
            GridAReceber.Lock = True
            
            'CodParc
            GridAReceber.Col = 6
            GridAReceber.Text = Val(RecPesq.Fields.Item(0).Value)
            GridAReceber.Lock = True
            
            'CodCReceb
            GridAReceber.Col = 7
            GridAReceber.Text = "0"
            GridAReceber.Lock = True
            
            VLIntLinha = VLIntLinha + 1
            
            GridAReceber.MaxRows = GridAReceber.MaxRows + 1
            RecCred.Close
            RecPesq.MoveNext
         Loop
         
        'monta dados da tabela de contas a receber
        Do While Not RecPesq2.EOF
            GridAReceber.Row = VLIntLinha
            GridAReceber.Lock = True
            
            'Descrição
            GridAReceber.Col = 1
            GridAReceber.TypeMaxEditLen = 255
            GridAReceber.Text = RecPesq2.Fields.Item(4).Value
            GridAReceber.Lock = True
            
            'Tipo
            GridAReceber.Col = 2
            GridAReceber.Text = VerificaNulo(RecPesq2.Fields.Item(1).Value)
            GridAReceber.Lock = True
            
            'Vencimento
            GridAReceber.Col = 3
            GridAReceber.Text = FormataData(RecPesq2.Fields.Item(2).Value)
            GridAReceber.Lock = True
            
            'Valor
            GridAReceber.Col = 4
            GridAReceber.Text = FormataMoeda(RecPesq2.Fields.Item(3).Value)
            If RecPesq2.Fields.Item(7).Value = "não" Then
                VLIntValor = VLIntValor + CCur(GridAReceber.Text)
            End If
            GridAReceber.Lock = True
            
            'Status
            GridAReceber.Col = 5
            If RecPesq2.Fields.Item(7).Value = "sim" Then
                GridAReceber.Text = "Recebido"
            ElseIf RecPesq2.Fields.Item(7).Value = "não" Then
                GridAReceber.Text = "A receber"
            End If
            GridAReceber.Lock = True
            
            'CodParc
            GridAReceber.Col = 6
            GridAReceber.Text = "0"
            GridAReceber.Lock = True
            
            'CodCReceb
            GridAReceber.Col = 7
            GridAReceber.Text = Val(RecPesq2.Fields.Item(0).Value)
            GridAReceber.Lock = True
            
            VLIntLinha = VLIntLinha + 1
            
            GridAReceber.MaxRows = GridAReceber.MaxRows + 1
            RecPesq2.MoveNext
         Loop
         
         'trava as linhas em branco para depois escrever o resumo
         GridAReceber.Row = GridAReceber.MaxRows
         GridAReceber.Col = 1
         GridAReceber.Lock = True
         GridAReceber.Col = 2
         GridAReceber.Lock = True
         GridAReceber.Col = 3
         GridAReceber.Lock = True
         GridAReceber.Col = 4
         GridAReceber.Lock = True
         GridAReceber.Col = 5
         GridAReceber.Lock = True
         GridAReceber.Col = 6
         GridAReceber.Lock = True
         
         
         GridAReceber.MaxRows = GridAReceber.MaxRows + 1
         GridAReceber.Row = GridAReceber.MaxRows
         
         'escreve o resumo da conta a receber
         GridAReceber.Col = 1
         GridAReceber.Text = "TOTAL À RECEBER:"
         GridAReceber.Font.Bold = True
         GridAReceber.Lock = True
         GridAReceber.Col = 2
         GridAReceber.Text = FormataMoeda(VLIntValor)
         GridAReceber.Font.Bold = True
         GridAReceber.Lock = True
         GridAReceber.Col = 3
         GridAReceber.Lock = True
         GridAReceber.Col = 4
         GridAReceber.Lock = True
         GridAReceber.Col = 5
         GridAReceber.Lock = True
         GridAReceber.Col = 6
         GridAReceber.Lock = True
         
         '===== CONTAGEM DE RECEBIMENTOS PESQUISADOS =========
         If (GridAReceber.MaxRows - 2) = 1 Then
            LblNumTotalReceb.Caption = FormataNum((GridAReceber.MaxRows - 2)) & " recebimento encontrado."
         Else
            LblNumTotalReceb.Caption = FormataNum((GridAReceber.MaxRows - 2)) & " recebimentos encontrados."
         End If
         '================================================
         
         CmdImprimirAReceber.Enabled = True
    End If

End Sub

Private Sub CmdImprimirAReceber_Click()
    Screen.MousePointer = vbHourglass
    
    Dim desc As String
    Dim tipo As String
    Dim venc As String
    Dim valor As String
    Dim status As String
    
    Dim VLStrLinha As String
    
    VLStrLinha = 1
    
    Conecta
    
    Do While VLStrLinha <= GridAReceber.MaxRows
        
        GridAReceber.Col = 1
        GridAReceber.Row = VLStrLinha
        desc = GridAReceber.Text
        
        GridAReceber.Col = 2
        GridAReceber.Row = VLStrLinha
        tipo = GridAReceber.Text
        
        GridAReceber.Col = 3
        GridAReceber.Row = VLStrLinha
        venc = GridAReceber.Text
        
        GridAReceber.Col = 4
        GridAReceber.Row = VLStrLinha
        valor = GridAReceber.Text
        
        GridAReceber.Col = 5
        GridAReceber.Row = VLStrLinha
        status = GridAReceber.Text
        
        vgCon.Execute "INSERT INTO tb_Auxiliar " & _
        "(campo01,campo02,campo03,campo04,campo05) " & _
        "VALUES ('" & desc & "','" & tipo & "','" & venc & "','" & valor & "','" & status & "')"
         
        VLStrLinha = VLStrLinha + 1
    Loop
    
    Desconecta
        
    rptCaixa_AReceber.Show

End Sub

Private Sub GridAReceber_Click(ByVal Col As Long, ByVal Row As Long)
    GridAReceber.Row = Row
    GridAReceber.Col = 6
    If GridAReceber.Text <> "CodParc" And GridAReceber.Text <> "" Then
       If GridAReceber.Text <> "0" Then
            VGIntCodReceber = GridAReceber.Text
            VGStrReceb = "parcela"
            CmdExcluirAReceber.Enabled = False
            CmdAlterarAReceber.Enabled = False
            
            GridAReceber.Row = Row
            GridAReceber.Col = 5
            If GridAReceber.Text = "A receber" Then
                CmdBaixaAReceber.Enabled = True
            Else
                CmdBaixaAReceber.Enabled = False
            End If
       End If
    Else
        CmdExcluirAReceber.Enabled = True
        CmdAlterarAReceber.Enabled = True
        CmdBaixaAReceber.Enabled = True
    End If
    
    GridAReceber.Row = Row
    GridAReceber.Col = 7
    If GridAReceber.Text <> "CodCReceb" And GridAReceber.Text <> "" Then
        If GridAReceber.Text <> "0" Then
            VGIntCodReceber = GridAReceber.Text
            VGStrReceb = "conta"
            CmdAlterarAReceber.Enabled = True
            CmdExcluirAReceber.Enabled = True
            
            GridAReceber.Row = Row
            GridAReceber.Col = 5
            If GridAReceber.Text = "A receber" Then
                CmdBaixaAReceber.Enabled = True
            Else
                CmdBaixaAReceber.Enabled = False
            End If
        End If
    Else
        CmdAlterarAReceber.Enabled = False
        CmdExcluirAReceber.Enabled = False
        CmdBaixaAReceber.Enabled = False
    End If
    
    GridAReceber.Row = Row
    GridAReceber.Col = 5
    If GridAReceber.Text <> "Status" And GridAReceber.Text <> "" Then
        VGStrStatusReceb = GridAReceber.Text
    Else
        VGStrStatusReceb = ""
    End If
End Sub

Private Sub GridAReceber_DblClick(ByVal Col As Long, ByVal Row As Long)
    GridAReceber.Row = Row
    GridAReceber.Col = 6
    If GridAReceber.Text <> "CodParc" And GridAReceber.Text <> "" Then
       If GridAReceber.Text <> "0" Then
            VGIntCodReceber = GridAReceber.Text
            FrmResumo_AReceber.Show
            CmdExcluirAReceber.Enabled = False
       End If
    Else
        CmdExcluirAReceber.Enabled = False
    End If
    
    GridAReceber.Row = Row
    GridAReceber.Col = 7
    If GridAReceber.Text <> "CodCReceb" And GridAReceber.Text <> "" Then
        If GridAReceber.Text <> "0" Then
            VGIntCodReceber = GridAReceber.Text
            FrmResumo_AReceber.Show
            
            CmdAlterarAReceber.Enabled = True
            CmdExcluirAReceber.Enabled = True
            CmdBaixaAReceber.Enabled = True
        End If
    Else
        CmdAlterarAReceber.Enabled = False
        CmdExcluirAReceber.Enabled = False
        CmdBaixaAReceber.Enabled = False
    End If
End Sub

Private Sub CmdIncluirAReceber_Click()
    FrmCaixa_AReceber_Inc.Show
End Sub

Private Sub CmdAlterarAReceber_Click()
    FrmCaixa_AReceber_Alt.Show
End Sub

Private Sub CmdBaixaAReceber_Click()
    If VGStrReceb = "parcela" Then
        FrmCrediario_Parcela_Quitar.Show
    
    ElseIf VGStrReceb = "conta" Then
        FrmCaixa_AReceber_Baixa.Show
    End If
End Sub
'========================================================================
'========================================================================



'========================================================================
'                   CREDIÁRIO
'========================================================================

Sub MontaCboTipoCred()
    Conecta
    
    Dim RecTipo As New ADODB.Recordset
    
    StrSql = "Select distinct TipoCred From tb_Crediario"
    RecTipo.Open StrSql, vgCon, 1, 3
    
    CboTipoCred.Clear
    CboTipoCred.AddItem ("")
    Do While Not RecTipo.EOF
        CboTipoCred.AddItem (RecTipo!tipocred)
        RecTipo.MoveNext
    Loop
    
    Desconecta
End Sub

Private Sub GridCrediario_Click(ByVal Col As Long, ByVal Row As Long)
    GridCrediario.Row = Row
    GridCrediario.Col = 11
    If GridCrediario.Text <> "" And GridCrediario.Text <> "CodCred" Then
        VGIntCodCred = GridCrediario.Text
        CmdExcluirCred.Enabled = True
        CmdVerParc.Enabled = True
    Else
        CmdExcluirCred.Enabled = False
        CmdVerParc.Enabled = False
    End If
End Sub

Private Sub GridCrediario_DblClick(ByVal Col As Long, ByVal Row As Long)
    GridCrediario.Row = Row
    GridCrediario.Col = 11
    If GridCrediario.Text <> "" And GridCrediario.Text <> "CodCred" Then
        VGIntCodCred = GridCrediario.Text
        FrmResumo_Crediario.Show
    End If
End Sub

Private Sub CmdLimparCred_Click()
    TxtCliCred.Text = ""
    TxtCredstaCred.Text = ""
    CboTipoCred.ListIndex = 0
    TxtDtCred1.Text = "__/__/____"
    TxtDtCred2.Text = "__/__/____"
    TxtDtVencCred1.Text = "__/__/____"
    TxtDtVencCred2.Text = "__/__/____"
    TxtCodParcCred.Text = ""
    GridCrediario.MaxRows = 0
    LblNumTotalCred.Caption = "Nenhum crediário encontrado."
    
    CmdExcluirCred.Enabled = False
    CmdImprimirCred.Enabled = False
    CmdVerParc.Enabled = False
End Sub

Private Sub CmdVerParc_Click()
    GridCrediario.Row = GridCrediario.ActiveRow
    GridCrediario.Col = 11
    If GridCrediario.Text <> "" And GridCrediario.Text <> "CodCred" Then
        VGIntCodCred = GridCrediario.Text
        FrmResumo_Crediario.Show
    End If
End Sub

Private Sub CmdExcluirCred_Click()
    If VGIntCodCred = 0 Then
        VPStrBox = MsgBox("Selecione um crediário na lista para excluir", vbExclamation, "Pró Vendas 2004 - Atenção")
    Else
        VPStrResponse = MsgBox("Deseja excluir este crediário?", vbYesNo, "Pró Vendas 2004 - Informação")
        
        If VPStrResponse = vbYes Then
            Conecta
            vgCon.Execute ("DELETE FROM tb_Crediario_Parcela_Quitacao WHERE CodCred=" & VGIntCodCred)
            vgCon.Execute ("DELETE FROM tb_Crediario_Parcela WHERE CodCred=" & VGIntCodCred)
            vgCon.Execute ("DELETE FROM tb_Crediario WHERE CodCred=" & VGIntCodCred)
            Desconecta
    
            FrmPrincipal.CmdPesqCred.Value = True
        End If
    End If
End Sub

Private Sub TxtCliCred_GotFocus()
    TxtCliCred.SelStart = 0
    TxtCliCred.SelLength = Len(TxtCliCred.Text)
End Sub

Private Sub TxtCredstaCred_GotFocus()
    TxtCredstaCred.SelStart = 0
    TxtCredstaCred.SelLength = Len(TxtCredstaCred.Text)
End Sub

Private Sub TxtDtCred1_GotFocus()
    If TxtDtCred1.Text = "__/__/____" Then
        TxtDtCred1.Text = ""
    Else
        TxtDtCred1.SelStart = 0
        TxtDtCred1.SelLength = Len(TxtDtCred1.Text)
    End If
End Sub

Private Sub TxtDtCred1_KeyPress(KeyAscii As Integer)
    '=== Só aceita números e barra ===
    If (KeyAscii < 47 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtDtCred1_LostFocus()
    Dim VLStrData As String

    If TxtDtCred1.Text <> "" Then
        VLStrData = VerificaData(TxtDtCred1.Text)

        If VGStrDataErro = "sim" Then
            TxtDtCred1.SetFocus
        Else
            TxtDtCred1.Text = VLStrData
        End If

        VGStrDataErro = ""
    Else
        TxtDtCred1.Text = "__/__/____"
    End If
End Sub

Private Sub TxtDtCred2_GotFocus()
    If TxtDtCred2.Text = "__/__/____" Then
        TxtDtCred2.Text = ""
    Else
        TxtDtCred2.SelStart = 0
        TxtDtCred2.SelLength = Len(TxtDtCred2.Text)
    End If
End Sub

Private Sub TxtDtCred2_KeyPress(KeyAscii As Integer)
    '=== Só aceita números e barra ===
    If (KeyAscii < 47 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtDtCred2_LostFocus()
    Dim VLStrData As String

    If TxtDtCred2.Text <> "" Then
        VLStrData = VerificaData(TxtDtCred2.Text)

        If VGStrDataErro = "sim" Then
            TxtDtCred2.SetFocus
        Else
            TxtDtCred2.Text = VLStrData
        End If

        VGStrDataErro = ""
    Else
        TxtDtCred2.Text = "__/__/____"
    End If
End Sub

Private Sub TxtDtVencCred1_GotFocus()
    If TxtDtVencCred1.Text = "__/__/____" Then
        TxtDtVencCred1.Text = ""
    Else
        TxtDtVencCred1.SelStart = 0
        TxtDtVencCred1.SelLength = Len(TxtDtVencCred1.Text)
    End If
End Sub

Private Sub TxtDtVencCred1_KeyPress(KeyAscii As Integer)
    '=== Só aceita números e barra ===
    If (KeyAscii < 47 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtDtVencCred1_LostFocus()
    Dim VLStrData As String

    If TxtDtVencCred1.Text <> "" Then
        VLStrData = VerificaData(TxtDtVencCred1.Text)

        If VGStrDataErro = "sim" Then
            TxtDtVencCred1.SetFocus
        Else
            TxtDtVencCred1.Text = VLStrData
        End If

        VGStrDataErro = ""
    Else
        TxtDtVencCred1.Text = "__/__/____"
    End If
End Sub

Private Sub TxtDtVencCred2_GotFocus()
    If TxtDtVencCred2.Text = "__/__/____" Then
        TxtDtVencCred2.Text = ""
    Else
        TxtDtVencCred2.SelStart = 0
        TxtDtVencCred2.SelLength = Len(TxtDtVencCred2.Text)
    End If
End Sub

Private Sub TxtDtVencCred2_KeyPress(KeyAscii As Integer)
    '=== Só aceita números e barra ===
    If (KeyAscii < 47 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtDtVencCred2_LostFocus()
    Dim VLStrData As String

    If TxtDtVencCred2.Text <> "" Then
        VLStrData = VerificaData(TxtDtVencCred2.Text)

        If VGStrDataErro = "sim" Then
            TxtDtVencCred2.SetFocus
        Else
            TxtDtVencCred2.Text = VLStrData
        End If

        VGStrDataErro = ""
    Else
        TxtDtVencCred2.Text = "__/__/____"
    End If
End Sub

Private Sub TxtCodParcCred_GotFocus()
    TxtCodParcCred.SelStart = 0
    TxtCodParcCred.SelLength = Len(TxtCodParcCred.Text)
End Sub

Private Sub CmdImprimirCred_Click()
    Screen.MousePointer = vbHourglass
    
    VPStrResponse = MsgBox("Deseja imprimir o crediário com as parcelas?", vbYesNo, "Pró Vendas 2004 - Informação")
    
    Dim cliente As String
    Dim credsta As String
    Dim data As String
    Dim tipocred As String
    Dim valorvenda As String
    Dim juros As String
    Dim valortotal As String
    Dim tipoentr As String
    Dim valorentr As String
    Dim parc As String
    Dim numparc As String
    
    Dim venc As String
    Dim valor As String
    Dim quitado As String
    
    Dim numcontrole As Integer
    
    Dim VLStrLinha As String

    VLStrLinha = 1

    Conecta
    
    numcontrole = 1
    
    Do While VLStrLinha <= GridCrediario.MaxRows

        GridCrediario.Col = 1
        GridCrediario.Row = VLStrLinha
        cliente = GridCrediario.Text

        GridCrediario.Col = 2
        GridCrediario.Row = VLStrLinha
        credsta = GridCrediario.Text

        GridCrediario.Col = 3
        GridCrediario.Row = VLStrLinha
        data = GridCrediario.Text

        GridCrediario.Col = 4
        GridCrediario.Row = VLStrLinha
        tipocred = GridCrediario.Text

        GridCrediario.Col = 5
        GridCrediario.Row = VLStrLinha
        valorvenda = GridCrediario.Text

        GridCrediario.Col = 6
        GridCrediario.Row = VLStrLinha
        juros = GridCrediario.Text

        GridCrediario.Col = 7
        GridCrediario.Row = VLStrLinha
        valortotal = GridCrediario.Text

        GridCrediario.Col = 8
        GridCrediario.Row = VLStrLinha
        tipoentr = GridCrediario.Text

        GridCrediario.Col = 9
        GridCrediario.Row = VLStrLinha
        valorentr = GridCrediario.Text

        GridCrediario.Col = 10
        GridCrediario.Row = VLStrLinha
        numparc = GridCrediario.Text

        If VPStrResponse = vbYes Then
            Dim RecParc As New ADODB.Recordset
            
            GridCrediario.Col = 11
            GridCrediario.Row = VLStrLinha
            
            StrSql = "Select NumParc,Vencimento,Valor,Quitado From tb_Crediario_Parcela where CodCred=" & GridCrediario.Text
            RecParc.Open StrSql, vgCon, 1, 3
            
            Do While Not RecParc.EOF
                parc = FormataNum(RecParc!numparc) & "/" & numparc
                venc = FormataData(RecParc!vencimento)
                valor = FormataMoeda(RecParc!valor)
                quitado = RecParc!quitado
                
                vgCon.Execute "INSERT INTO tb_Auxiliar " & _
                "(campo01,campo02,campo03,campo04,campo05,campo06,campo07,campo08,campo09,campo10,campo11,campo12,campo13,campo14,campo15) " & _
                "VALUES ('" & cliente & "','" & credsta & "','" & data & "','" & tipocred & "','" & valorvenda & "','" & juros & "','" & valortotal & "','" & tipoentr & "','" & valorentr & "','" & numparc & "','" & parc & "','" & venc & "','" & valor & "','" & quitado & "','" & FormataNum(numcontrole) & "')"
                
                numcontrole = numcontrole + 1
                
                RecParc.MoveNext
            Loop
            RecParc.Close
        Else
            vgCon.Execute "INSERT INTO tb_Auxiliar " & _
            "(campo01,campo02,campo03,campo04,campo05,campo06,campo07,campo08,campo09,campo10,campo11) " & _
            "VALUES ('" & cliente & "','" & credsta & "','" & data & "','" & tipocred & "','" & valorvenda & "','" & juros & "','" & valortotal & "','" & tipoentr & "','" & valorentr & "','" & numparc & "','" & FormataNum(numcontrole) & "')"
        
            numcontrole = numcontrole + 1
        End If

        VLStrLinha = VLStrLinha + 1
    Loop

    Desconecta

    If VPStrResponse = vbYes Then
        rptCrediario_Parcela.Show
    Else
        rptCrediario.Show
    End If
End Sub

Private Sub CmdIncluirCred_Click()
    FrmCrediario_Inc.Show
End Sub

Private Sub CmdPesqCred_Click()
    Screen.MousePointer = vbHourglass
    
    Dim VLStrOrder As String
    
    Conecta
    
    StrSql = "SELECT CR.CodCred,CR.CodCredsta,CR.CodCli,CR.DtCred,CR.TipoCred," & _
             "CR.ValorVenda,CR.Parcela,CR.Juros,CR.ValorTotal,CR.TipoEntr,CR.ValorEntr," & _
             "CR.Numbanco,CR.NumCheque FROM tb_Crediario as CR WHERE 0=0"
            
    '====== PESQUISAR POR CLIENTE ==========
    If TxtCliCred.Text <> "" Then
        'StrSql = Trim(Mid(StrSql, 1, InStr(StrSql, "FROM") - 1)) & ",C.nome " & Trim(Mid(StrSql, InStr(StrSql, "FROM"), InStr(StrSql, "WHERE") - InStr(StrSql, "FROM"))) & ",tb_Cliente as C " & Trim(Mid(StrSql, InStr(StrSql, "WHERE"))) & " and C.CodCli=CR.CodCli and C.Nome like '%" & TxtCliCred.Text & "%'"
        StrSql = Trim(Mid(StrSql, 1, InStr(StrSql, "FROM") - 1)) & " " & Trim(Mid(StrSql, InStr(StrSql, "FROM"), InStr(StrSql, "WHERE") - InStr(StrSql, "FROM"))) & ",tb_Cliente as C " & Trim(Mid(StrSql, InStr(StrSql, "WHERE"))) & " and C.CodCli=CR.CodCli and C.Nome like '%" & TxtCliCred.Text & "%'"
        VLStrOrder = VLStrOrder + "C.Nome,"
    End If
            
    '====== PESQUISAR POR CREDIARISTA ==========
    If TxtCredstaCred.Text <> "" Then
        'StrSql = Trim(Mid(StrSql, 1, InStr(StrSql, "FROM") - 1)) & ",CS.CodCredsta,CS.nome " & Trim(Mid(StrSql, InStr(StrSql, "FROM"), InStr(StrSql, "WHERE") - InStr(StrSql, "FROM"))) & ",tb_Crediarista as CS " & Trim(Mid(StrSql, InStr(StrSql, "WHERE"))) & " and CS.CodCredsta = CR.CodCredsta and CS.Nome like '%" & TxtCredstaCred.Text & "%'"
        StrSql = Trim(Mid(StrSql, 1, InStr(StrSql, "FROM") - 1)) & " " & Trim(Mid(StrSql, InStr(StrSql, "FROM"), InStr(StrSql, "WHERE") - InStr(StrSql, "FROM"))) & ",tb_Crediarista as CS " & Trim(Mid(StrSql, InStr(StrSql, "WHERE"))) & " and CS.CodCredsta = CR.CodCredsta and CS.Nome like '%" & TxtCredstaCred.Text & "%'"
        VLStrOrder = VLStrOrder + "CS.Nome,"
    End If
            
    '====== PESQUISAR POR TIPO CREDIÁRIO ==========
    If CboTipoCred.Text <> "" Then
        StrSql = StrSql + " and CR.TipoCred='" & CboTipoCred.Text & "'"
        VLStrOrder = VLStrOrder + "CR.TipoCred,"
    End If
    
    '====== PESQUISAR POR DATA DO CREDIÁRIO ==========
    If (TxtDtCred1.Text <> "" And TxtDtCred1.Text <> "__/__/____") And (TxtDtCred2.Text <> "" And TxtDtCred2.Text <> "__/__/____") Then
        StrSql = StrSql + " and CR.DtCred >=#" & FormataDataUS(TxtDtCred1.Text) & "# and CR.DtCred <= #" & FormataDataUS(TxtDtCred2.Text) & "#"
        VLStrOrder = VLStrOrder + "CR.DtCred desc,"
    
    ElseIf (TxtDtCred1.Text <> "" And TxtDtCred1.Text <> "__/__/____") And (TxtDtCred2.Text = "" Or TxtDtCred2.Text = "__/__/____") Then
        StrSql = StrSql + " and CR.DtCred =#" & FormataDataUS(TxtDtCred1.Text) & "#"
        VLStrOrder = VLStrOrder + "CR.DtCred desc,"
    
    ElseIf (TxtDtCred1.Text = "" Or TxtDtCred1.Text = "__/__/____") And (TxtDtCred2.Text <> "" And TxtDtCred2.Text <> "__/__/____") Then
        StrSql = StrSql + " and CR.DtCred =#" & FormataDataUS(TxtDtCred2.Text) & "#"
        VLStrOrder = VLStrOrder + "CR.DtCred desc,"
    End If
            
    '====== PESQUISAR POR DATA DO VENCIMENTO ==========
    If (TxtDtVencCred1.Text <> "" And TxtDtVencCred1.Text <> "__/__/____") And (TxtDtVencCred2.Text <> "" And TxtDtVencCred2.Text <> "__/__/____") Then
        'StrSql = StrSql + " and CP.Vencimento >=#" & FormataDataUS(TxtDtVencCred1.Text) & "# and CP.Vencimento <= #" & FormataDataUS(TxtDtVencCred2.Text) & "#"
        StrSql = Trim(Mid(StrSql, 1, InStr(StrSql, "FROM") - 1)) & " " & Trim(Mid(StrSql, InStr(StrSql, "FROM"), InStr(StrSql, "WHERE") - InStr(StrSql, "FROM"))) & ",tb_Crediario_Parcela as CP " & Trim(Mid(StrSql, InStr(StrSql, "WHERE"))) & " and CP.CodCred = CR.CodCred and CP.Vencimento >=#" & FormataDataUS(TxtDtVencCred1.Text) & "# and CP.Vencimento <= #" & FormataDataUS(TxtDtVencCred2.Text) & "#"
        VLStrOrder = VLStrOrder + "CP.Vencimento desc,"
    
    ElseIf (TxtDtVencCred1.Text <> "" And TxtDtVencCred1.Text <> "__/__/____") And (TxtDtVencCred2.Text = "" Or TxtDtVencCred2.Text = "__/__/____") Then
        'StrSql = StrSql + " and CP.Vencimento =#" & FormataDataUS(TxtDtVencCred1.Text) & "#"
        StrSql = Trim(Mid(StrSql, 1, InStr(StrSql, "FROM") - 1)) & " " & Trim(Mid(StrSql, InStr(StrSql, "FROM"), InStr(StrSql, "WHERE") - InStr(StrSql, "FROM"))) & ",tb_Crediario_Parcela as CP " & Trim(Mid(StrSql, InStr(StrSql, "WHERE"))) & " and CP.CodCred = CR.CodCred and CP.Vencimento =#" & FormataDataUS(TxtDtVencCred1.Text) & "#"
        VLStrOrder = VLStrOrder + "CP.Vencimento desc,"
    
    ElseIf (TxtDtVencCred1.Text = "" Or TxtDtVencCred1.Text = "__/__/____") And (TxtDtVencCred2.Text <> "" And TxtDtVencCred2.Text <> "__/__/____") Then
        'StrSql = StrSql + " and CP.Vencimento =#" & FormataDataUS(TxtDtVencCred2.Text) & "#"
        StrSql = Trim(Mid(StrSql, 1, InStr(StrSql, "FROM") - 1)) & " " & Trim(Mid(StrSql, InStr(StrSql, "FROM"), InStr(StrSql, "WHERE") - InStr(StrSql, "FROM"))) & ",tb_Crediario_Parcela as CP " & Trim(Mid(StrSql, InStr(StrSql, "WHERE"))) & " and CP.CodCred = CR.CodCred and CP.Vencimento =#" & FormataDataUS(TxtDtVencCred2.Text) & "#"
        VLStrOrder = VLStrOrder + "CP.Vencimento desc,"
    End If
            
    '====== PESQUISAR POR CÓDIGO DA PARCELA ==========
    If TxtCodParcCred.Text <> "" Then
        'StrSql = Trim(Mid(StrSql, 1, InStr(StrSql, "FROM") - 1)) & ",CP.codparc,CP.vencimento,CP.valor,CP.quitado,CP.NumParc " & Trim(Mid(StrSql, InStr(StrSql, "FROM"), InStr(StrSql, "WHERE") - InStr(StrSql, "FROM"))) & ",tb_CrediarioParcela as CP " & Trim(Mid(StrSql, InStr(StrSql, "WHERE"))) & " and CP.CodCred = CR.CodCred and CP.CodParc=" & TxtCodParcCred.Text & ""
        If InStr(StrSql, "tb_Crediario_Parcela as CP") <> 0 Then
            StrSql = Trim(Mid(StrSql, 1, InStr(StrSql, "FROM") - 1)) & " " & Trim(Mid(StrSql, InStr(StrSql, "FROM"), InStr(StrSql, "WHERE") - InStr(StrSql, "FROM"))) & " " & Trim(Mid(StrSql, InStr(StrSql, "WHERE"))) & " and CP.CodCred = CR.CodCred and CP.CodParc=" & TxtCodParcCred.Text & ""
        Else
            StrSql = Trim(Mid(StrSql, 1, InStr(StrSql, "FROM") - 1)) & " " & Trim(Mid(StrSql, InStr(StrSql, "FROM"), InStr(StrSql, "WHERE") - InStr(StrSql, "FROM"))) & ",tb_Crediario_Parcela as CP " & Trim(Mid(StrSql, InStr(StrSql, "WHERE"))) & " and CP.CodCred = CR.CodCred and CP.CodParc=" & TxtCodParcCred.Text & ""
        End If
    End If
        
    '====== ORDENAR PESQUISA ======================
    If VLStrOrder <> "" Then
        VLStrOrder = Mid(VLStrOrder, 1, Len(VLStrOrder) - 1)
        StrSql = StrSql + " order by " & VLStrOrder & ",CR.DtCred desc"
    Else
        StrSql = StrSql + " order by CR.DtCred desc"
    End If
    
    VLStrOrder = ""
    
    RecPesq.Open StrSql, vgCon, 1, 3
            
    Call MontaGridCrediario
        
    Desconecta
    
    CmdImprimirCred.Enabled = True
    
    Screen.MousePointer = vbNormal

End Sub

Sub MontaGridCrediario()
    Dim VLIntCodCred As Long
    Dim VLIntLinha As Long
    Dim RecCli As New ADODB.Recordset
    Dim RecCredsta As New ADODB.Recordset
    
    If RecPesq.EOF Then
        LblNumTotalCred.Caption = "Nenhum crediário encontrado."
        
        VPStrBox = MsgBox("Pesquisa sem resultados.", vbInformation, "Pró Vendas 2004 - Informação")
        GridCrediario.Refresh
        GridCrediario.MaxRows = 0
        
        CmdExcluirCred.Enabled = False
        CmdImprimirCred.Enabled = False

    
    Else
    
        VLIntLinha = 1
        GridCrediario.MaxRows = VLIntLinha
         
        Do While Not RecPesq.EOF
                 
            GridCrediario.Row = VLIntLinha
            GridCrediario.Lock = True
            
            'Cliente
            StrSql = "Select Nome from tb_Cliente where CodCli=" & RecPesq!CodCli
            RecCli.Open StrSql, vgCon, 1, 3
            
            GridCrediario.Col = 1
            If Not RecCli.EOF Then
                GridCrediario.Text = VerificaNulo(RecCli!nome)
            Else
                GridCrediario.Text = ""
            End If
            GridCrediario.Lock = True
            
            'Crediarista
            StrSql = "Select CodCredsta,Nome from tb_Crediarista where CodCredsta=" & RecPesq!CodCredsta
            RecCredsta.Open StrSql, vgCon, 1, 3
            
            GridCrediario.Col = 2
            If Not RecCredsta.EOF Then
                GridCrediario.Text = VerificaNulo(RecCredsta!nome)
            Else
                GridCrediario.Text = ""
            End If
            GridCrediario.Lock = True
            
            'Data Crediário
            GridCrediario.Col = 3
            GridCrediario.Text = FormataData(RecPesq!dtcred)
            GridCrediario.Lock = True
            
            'Tipo crediário
            GridCrediario.Col = 4
            GridCrediario.Text = VerificaNulo(RecPesq!tipocred)
            GridCrediario.Lock = True
            
            'Valor venda
            GridCrediario.Col = 5
            GridCrediario.Text = FormataMoeda(VerificaNulo(RecPesq!valorvenda))
            GridCrediario.Lock = True
            
            'Juros
            GridCrediario.Col = 6
            If RecPesq!juros <> "" And IsNull(RecPesq!juros) = False Then
                GridCrediario.Text = FormataNum(RecPesq!juros) & "%"
            Else
                GridCrediario.Text = ""
            End If
            GridCrediario.Lock = True
            
            'Valor total
            GridCrediario.Col = 7
            GridCrediario.Text = FormataMoeda(VerificaNulo(RecPesq!valortotal))
            GridCrediario.Lock = True
            
            'Tipo entrada
            GridCrediario.Col = 8
            If (RecPesq!Numbanco <> "" And IsNull(RecPesq!Numbanco) = False) And (RecPesq!Numcheque <> "" And IsNull(RecPesq!Numcheque) = False) Then
                GridCrediario.Text = VerificaNulo(RecPesq!tipoentr) & " (" & RecPesq!Numbanco & "/" & RecPesq!Numcheque & ")"
            Else
                GridCrediario.Text = VerificaNulo(RecPesq!tipoentr)
            End If
            GridCrediario.Lock = True
            
            'Valor entrada
            GridCrediario.Col = 9
            If RecPesq!valorentr <> "" And IsNull(RecPesq!valorentr) = False Then
                GridCrediario.Text = FormataMoeda(RecPesq!valorentr)
            Else
                GridCrediario.Text = ""
            End If
            GridCrediario.Lock = True
            
            'Parcela
            GridCrediario.Col = 10
            GridCrediario.Text = FormataNum(RecPesq!parcela)
            GridCrediario.Lock = True
            
            'CodCred
            GridCrediario.Col = 11
            GridCrediario.Text = Val(RecPesq!CodCred)
            GridCrediario.Lock = True
            
            'CodCredsta
            GridCrediario.Col = 12
            GridCrediario.Text = Val(RecPesq!CodCredsta)
            GridCrediario.Lock = True
            
            RecCli.Close
            RecCredsta.Close
            
            VLIntLinha = VLIntLinha + 1
            
            GridCrediario.MaxRows = GridCrediario.MaxRows + 1
            RecPesq.MoveNext
         Loop
         
         '===== CONTAGEM DE CREDIÁRIOS PESQUISADOS =========
         GridCrediario.MaxRows = GridCrediario.MaxRows - 1
         
         If GridCrediario.MaxRows = 1 Then
            LblNumTotalCred.Caption = FormataNum(GridCrediario.MaxRows) & " crediário encontrado."
         Else
            LblNumTotalCred.Caption = FormataNum(GridCrediario.MaxRows) & " crediários encontrados."
         End If
         '================================================
         
    End If

End Sub
'========================================================================
'========================================================================





'========================================================================
'                   CREDIARISTA
'========================================================================

Private Sub GridCredsta_DblClick(ByVal Col As Long, ByVal Row As Long)
    GridCredsta.Row = Row
    GridCredsta.Col = 12
    If GridCredsta.Text <> "" And GridCredsta.Text <> "CodCredsta" Then
        VGIntCodCredsta = GridCredsta.Text
        FrmResumo_Crediarista.Show
    End If
End Sub

Private Sub CmdLimparCredsta_Click()
    TxtNomeCredsta.Text = ""
    TxtBairroCredsta.Text = ""
    TxtCpfCredsta.Text = ""
    TxtTelCredsta.Text = ""
    GridCredsta.MaxRows = 0
    LblNumTotalCredsta.Caption = "Nenhum crediarista encontrado."
    
    CmdExcluirEst.Enabled = False
    CmdImprimirEst.Enabled = False
End Sub

Private Sub CmdAlterarCredsta_Click()
    If VGIntCodCredsta = 0 Then
        VPStrBox = MsgBox("Selecione um crediarista na lista para alterar", vbExclamation, "Pró Vendas 2004 - Atenção")
    Else
        FrmCrediarista_Alt.Show
    End If
End Sub

Private Sub CmdExcluirCredsta_Click()
    If VGIntCodCredsta = 0 Then
        VPStrBox = MsgBox("Selecione um crediarista na lista para excluir", vbExclamation, "Pró Vendas 2004 - Atenção")
    Else
        VPStrResponse = MsgBox("Deseja excluir este crediarista?", vbYesNo, "Pró Vendas 2004 - Informação")
        
        If VPStrResponse = vbYes Then
            Conecta
            Dim RecCred As New ADODB.Recordset
            
            'verifica se este crediarista está relacionado a algum crediário
            StrSql = "SELECT CodCredsta from tb_Crediario WHERE CodCredsta=" & VGIntCodCredsta
            RecCred.Open StrSql, vgCon, 1, 3
            
            If Not RecCred.EOF Then
                VPStrBox = MsgBox("Este crediarista está relacionado a um crediário." & Chr(13) & "Não será possível excluir.", vbExclamation, "Pró Vendas 2004 - Atenção")
            Else
                vgCon.Execute ("DELETE FROM tb_Crediarista WHERE CodCredsta=" & VGIntCodCredsta)
            End If
            
            Desconecta
    
            FrmPrincipal.CmdPesqCredsta.Value = True
        End If
    End If
End Sub

Private Sub TxtCpfCredsta_KeyPress(KeyAscii As Integer)
    '=== Só aceita números ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtNomeCredsta_GotFocus()
    TxtNomeCredsta.SelStart = 0
    TxtNomeCredsta.SelLength = Len(TxtNomeCredsta.Text)
End Sub

Private Sub TxtBairroCredsta_GotFocus()
    TxtBairroCredsta.SelStart = 0
    TxtBairroCredsta.SelLength = Len(TxtBairroCredsta.Text)
End Sub

Private Sub TxtCpfCredsta_GotFocus()
    TxtCpfCredsta.SelStart = 0
    TxtCpfCredsta.SelLength = Len(TxtCpfCredsta.Text)
End Sub

Private Sub TxtTelCredsta_GotFocus()
    TxtTelCredsta.SelStart = 0
    TxtTelCredsta.SelLength = Len(TxtTelCredsta.Text)
End Sub

Private Sub CmdImprimirCredsta_Click()
    Screen.MousePointer = vbHourglass

    Dim nome As String
    Dim endereco As String
    Dim bairro As String
    Dim cep As String
    Dim cidest As String
    Dim datanasc As String
    Dim tel As String
    Dim cpf As String
    Dim email As String
    Dim obs As String

    Dim VLStrLinha As String

    VLStrLinha = 1

    Conecta

    Do While VLStrLinha <= GridCredsta.MaxRows

        GridCredsta.Col = 1
        GridCredsta.Row = VLStrLinha
        nome = GridCredsta.Text

        GridCredsta.Col = 2
        GridCredsta.Row = VLStrLinha
        endereco = GridCredsta.Text

        GridCredsta.Col = 3
        GridCredsta.Row = VLStrLinha
        bairro = GridCredsta.Text

        GridCredsta.Col = 4
        GridCredsta.Row = VLStrLinha
        cep = GridCredsta.Text

        GridCredsta.Col = 5
        GridCredsta.Row = VLStrLinha
        cidest = GridCredsta.Text

        GridCredsta.Col = 6
        GridCredsta.Row = VLStrLinha
        cidest = cidest & "/" & GridCredsta.Text

        GridCredsta.Col = 7
        GridCredsta.Row = VLStrLinha
        datanasc = GridCredsta.Text

        GridCredsta.Col = 8
        GridCredsta.Row = VLStrLinha
        tel = GridCredsta.Text

        GridCredsta.Col = 9
        GridCredsta.Row = VLStrLinha
        cpf = GridCredsta.Text

        GridCredsta.Col = 10
        GridCredsta.Row = VLStrLinha
        email = GridCredsta.Text

        GridCredsta.Col = 11
        GridCredsta.Row = VLStrLinha
        obs = GridCredsta.Text

        vgCon.Execute "INSERT INTO tb_Auxiliar " & _
        "(campo01,campo02,campo03,campo04,campo05,campo06,campo07,campo08,campo09,campo10) " & _
        "VALUES ('" & nome & "','" & endereco & "','" & bairro & "','" & cep & "','" & cidest & "','" & datanasc & "','" & tel & "','" & cpf & "','" & email & "','" & obs & "')"

        VLStrLinha = VLStrLinha + 1
    Loop

    Desconecta

    rptCrediarista.Show

End Sub

Private Sub CmdIncluirCredsta_Click()
    FrmCrediarista_Inc.Show
End Sub

Private Sub GridCredsta_Click(ByVal Col As Long, ByVal Row As Long)
    GridCredsta.Row = Row
    GridCredsta.Col = 12
    If GridCredsta.Text <> "" And GridCredsta.Text <> "CodCredsta" Then
        VGIntCodCredsta = GridCredsta.Text
        CmdAlterarCredsta.Enabled = True
        CmdExcluirCredsta.Enabled = True
    Else
        CmdAlterarCredsta.Enabled = False
        CmdExcluirCredsta.Enabled = False
    End If
End Sub

Sub MontaGridCredsta()
    Dim VLIntCodCredsta As Long
    Dim VLIntLinha As Long
    Dim VLStrTel1 As String
    Dim VLStrTel2 As String

    If RecPesq.EOF Then
        LblNumTotalCredsta.Caption = "Nenhum crediarista encontrado."

        VPStrBox = MsgBox("Pesquisa sem resultados.", vbInformation, "Pró Vendas 2004 - Informação")
        GridCredsta.Refresh
        GridCredsta.MaxRows = 0

        CmdAlterarCredsta.Enabled = False
        CmdExcluirCredsta.Enabled = False
        CmdImprimirCredsta.Enabled = False

    Else

        VLIntLinha = 1
        GridCredsta.MaxRows = VLIntLinha

        Do While Not RecPesq.EOF

            GridCredsta.Row = VLIntLinha
            GridCredsta.Lock = True

            'Nome
            GridCredsta.Col = 1
            GridCredsta.Text = VerificaNulo(RecPesq!nome)
            GridCredsta.Lock = True

            'Endereço
            GridCredsta.Col = 2
            GridCredsta.Text = VerificaNulo(RecPesq!endereco)
            GridCredsta.Lock = True

            'Bairro
            GridCredsta.Col = 3
            GridCredsta.Text = VerificaNulo(RecPesq!bairro)
            GridCredsta.Lock = True

            'Cep
            GridCredsta.Col = 4
            GridCredsta.Text = VerificaNulo(RecPesq!cep)
            GridCredsta.Lock = True

            'Cidade
            GridCredsta.Col = 5
            GridCredsta.Text = VerificaNulo(RecPesq!cidade)
            GridCredsta.Lock = True

            'Estado
            GridCredsta.Col = 6
            GridCredsta.Text = VerificaNulo(RecPesq!Estado)
            GridCredsta.Lock = True

            'Data Nascimento
            GridCredsta.Col = 7
            GridCredsta.Text = FormataData(VerificaNulo(RecPesq!dtnasc))
            GridCredsta.Lock = True

            'Telefone
            GridCredsta.Col = 8
            If RecPesq!telefone1 <> "" And IsNull(RecPesq!telefone1) = False Then
                VLStrTel1 = RecPesq!telefone1
            Else
                VLStrTel1 = ""
            End If
            
            If RecPesq!telefone2 <> "" And IsNull(RecPesq!telefone2) = False Then
                VLStrTel2 = RecPesq!telefone2
            Else
                VLStrTel2 = ""
            End If
            
            If VLStrTel1 = "" And VLStrTel2 = "" Then
                GridCredsta.Text = ""
                
            ElseIf VLStrTel1 <> "" And VLStrTel2 <> "" Then
                GridCredsta.Text = VLStrTel1 & " / " & VLStrTel2
                
            ElseIf VLStrTel1 <> "" And VLStrTel2 = "" Then
                GridCredsta.Text = VLStrTel1
                
            ElseIf VLStrTel1 = "" And VLStrTel2 <> "" Then
                GridCredsta.Text = VLStrTel2
                
            End If
            GridCredsta.Lock = True

            'Cpf
            GridCredsta.Col = 9
            GridCredsta.Text = VerificaNulo(RecPesq!cpf)
            GridCredsta.Lock = True

            'Email
            GridCredsta.Col = 10
            GridCredsta.Text = VerificaNulo(RecPesq!email)
            GridCredsta.Lock = True

            'Observação
            GridCredsta.Col = 11
            GridCredsta.Text = VerificaNulo(RecPesq!obs)
            GridCredsta.Lock = True

            'CodCredsta
            GridCredsta.Col = 12
            GridCredsta.Text = Val(RecPesq!CodCredsta)
            GridCredsta.Lock = True

            VLIntLinha = VLIntLinha + 1

            GridCredsta.MaxRows = GridCredsta.MaxRows + 1
            RecPesq.MoveNext
         Loop

         '===== CONTAGEM DE CREDIARISTAS PESQUISADOS =========
         GridCredsta.MaxRows = GridCredsta.MaxRows - 1

         If GridCredsta.MaxRows = 1 Then
            LblNumTotalCredsta.Caption = FormataNum(GridCredsta.MaxRows) & " crediarista encontrado."
         Else
            LblNumTotalCredsta.Caption = FormataNum(GridCredsta.MaxRows) & " crediaristas encontrados."
         End If
         '================================================

         CmdImprimirCredsta.Enabled = True
    End If

End Sub

Private Sub TxtTelCredsta_KeyPress(KeyAscii As Integer)
    '=== Só aceita números, parênteses, espaço e - ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 40 And KeyAscii <> 41 And KeyAscii <> 32 And KeyAscii <> 45 Then
        KeyAscii = 0
    End If
End Sub

Private Sub CmdPesqCredsta_Click()

    Screen.MousePointer = vbHourglass

    Dim VLStrOrder As String

    Conecta

    StrSql = "Select * from tb_Crediarista where 0=0"

    '====== PESQUISAR POR NOME ==========
    If TxtNomeCredsta.Text <> "" Then
        StrSql = StrSql + " and Nome like '%" & TxtNomeCredsta.Text & "%'"
        VLStrOrder = VLStrOrder + "Nome,"
    End If

    '====== PESQUISAR POR CPF ==========
    If TxtCpfCredsta.Text <> "" Then
        StrSql = StrSql + " and Cpf='" & TxtCpfCredsta.Text & "'"
        VLStrOrder = VLStrOrder + "Cpf,"
    End If

    '====== PESQUISAR POR BAIRRO ==========
    If TxtBairroCredsta.Text <> "" Then
        StrSql = StrSql + " and Bairro like '%" & TxtBairroCredsta.Text & "%'"
        VLStrOrder = VLStrOrder + "Bairro,"
    End If

    '====== PESQUISAR POR TELEFONE ==========
    If TxtTelCredsta.Text <> "" Then
        StrSql = StrSql + " and Telefone1 like '%" & TxtTelCredsta.Text & "%' or Telefone2 like '%" & TxtTelCredsta.Text & "%'"
        VLStrOrder = VLStrOrder + "Telefone1,Telefone2,"
    End If

    '====== ORDENAR PESQUISA ======================
    If VLStrOrder <> "" Then
        VLStrOrder = Mid(VLStrOrder, 1, Len(VLStrOrder) - 1)
        StrSql = StrSql + " order by " & VLStrOrder
    Else
        StrSql = StrSql + " order by Nome"
    End If

    VLStrOrder = ""

    RecPesq.Open StrSql, vgCon, 1, 3

    Call MontaGridCredsta

    Desconecta

    Screen.MousePointer = vbNormal

End Sub
'========================================================================
'========================================================================




'========================================================================
'                            ORÇAMENTO
'========================================================================

Private Sub CmdLimparOrc_Click()
    TxtCliOrc.Text = ""
    TxtVendOrc.Text = ""
    TxtTelOrc.Text = ""
    TxtDtOrc1.Text = "__/__/____"
    TxtDtOrc2.Text = "__/__/____"
    GridOrcamento.MaxRows = 0
    LblNumTotalOrc.Caption = "Nenhum orçamento encontrado."
    
    CmdAlterarOrc.Enabled = False
    CmdExcluirOrc.Enabled = False
    CmdImprimirOrc.Enabled = False
End Sub

Private Sub CmdIncluirOrc_Click()
    FrmOrcamento_Inc.Show
End Sub

Private Sub CmdAlterarOrc_Click()
    FrmOrcamento_Alt.Show
End Sub

Private Sub CmdExcluirOrc_Click()
    VPStrResponse = MsgBox("Deseja excluir este orçamento?", vbYesNo, "Pró Vendas 2004 - Informação")
    If VPStrResponse = vbYes Then

        Conecta
        vgCon.Execute ("DELETE FROM tb_Orcamento_Produto WHERE CodOrc=" & VGIntCodOrc)
        vgCon.Execute ("DELETE FROM tb_Orcamento WHERE CodOrc=" & VGIntCodOrc)
        Desconecta

        FrmPrincipal.CmdPesqOrc.Value = True
    End If
End Sub

Private Sub CmdImprimirOrc_Click()
    VGStrPersonalizar = "orçamento"
    FrmAssinaturaOrc.Show
End Sub

Private Sub CmdPesqOrc_Click()
    Screen.MousePointer = vbHourglass

    Dim VLStrOrder As String

    Conecta

    StrSql = "SELECT O.CodOrc,O.CodVendedor,O.DtOrc,O.Nome,O.Telefone1,O.Telefone2,O.TotalVenda," & _
             "O.Parcela,O.Entrada,O.ValorParc,O.ValorPrazo,O.Validade,O.Obs FROM tb_Orcamento as O WHERE 0=0"

    '====== PESQUISAR POR CLIENTE ==========
    If TxtCliOrc.Text <> "" Then
        StrSql = StrSql + " and O.Nome like '%" & TxtCliOrc.Text & "%'"
        VLStrOrder = VLStrOrder + "O.Nome,"
    End If

    '====== PESQUISAR POR VENDEDOR ==========
    If TxtVendOrc.Text <> "" Then
        StrSql = Trim(Mid(StrSql, 1, InStr(StrSql, "FROM") - 1)) & ",V.Nome " & Trim(Mid(StrSql, InStr(StrSql, "FROM"), InStr(StrSql, "WHERE") - InStr(StrSql, "FROM"))) & ",tb_Vendedor as V " & Trim(Mid(StrSql, InStr(StrSql, "WHERE"))) & " and O.CodVendedor=V.CodVendedor and V.Nome like '%" & TxtVendOrc.Text & "%'"
        VLStrOrder = VLStrOrder + "V.Nome,"
    End If

    '====== PESQUISAR POR TELEFONE ==========
    If TxtTelOrc.Text <> "" Then
        StrSql = StrSql + " and O.Telefone1 like '%" & TxtTelOrc.Text & "%' or O.Telefone2 like '%" & TxtTelOrc.Text & "%'"
    End If

    '====== PESQUISAR POR DATA DO ORÇAMENTO ==========
    If (TxtDtOrc1.Text <> "" And TxtDtOrc1.Text <> "__/__/____") And (TxtDtOrc2.Text <> "" And TxtDtOrc2.Text <> "__/__/____") Then
        StrSql = StrSql + " and O.DtOrc >=#" & FormataDataUS(TxtDtOrc1.Text) & "# and O.DtOrc <= #" & FormataDataUS(TxtDtOrc2.Text) & "#"
        VLStrOrder = VLStrOrder + "O.DtOrc desc,"

    ElseIf (TxtDtOrc1.Text <> "" And TxtDtOrc1.Text <> "__/__/____") And (TxtDtOrc2.Text = "" Or TxtDtOrc2.Text = "__/__/____") Then
        StrSql = StrSql + " and O.DtOrc =#" & FormataDataUS(TxtDtOrc1.Text) & "#"
        VLStrOrder = VLStrOrder + "O.DtOrc desc,"

    ElseIf (TxtDtOrc1.Text = "" Or TxtDtOrc1.Text = "__/__/____") And (TxtDtOrc2.Text <> "" And TxtDtOrc2.Text <> "__/__/____") Then
        StrSql = StrSql + " and O.DtOrc =#" & FormataDataUS(TxtDtOrc2.Text) & "#"
        VLStrOrder = VLStrOrder + "O.DtOrc desc,"
    End If

    '====== ORDENAR PESQUISA ======================
    If VLStrOrder <> "" Then
        VLStrOrder = Mid(VLStrOrder, 1, Len(VLStrOrder) - 1)
        StrSql = StrSql + " order by " & VLStrOrder
    Else
        StrSql = StrSql + " order by O.Nome"
    End If

    VLStrOrder = ""

    RecPesq.Open StrSql, vgCon, 1, 3

    Call MontaGridOrcamento

    Desconecta

    Screen.MousePointer = vbNormal

End Sub

Private Sub GridOrcamento_Click(ByVal Col As Long, ByVal Row As Long)
    GridOrcamento.Row = Row
    GridOrcamento.Col = 12

    If GridOrcamento.Text <> "" And GridOrcamento.Text <> "CodOrc" Then
        VGIntCodOrc = GridOrcamento.Text
        CmdAlterarOrc.Enabled = True
        CmdExcluirOrc.Enabled = True
        CmdImprimirOrc.Enabled = True
    Else
        CmdAlterarOrc.Enabled = False
        CmdExcluirOrc.Enabled = False
        CmdImprimirOrc.Enabled = False
    End If
End Sub

Private Sub GridOrcamento_DblClick(ByVal Col As Long, ByVal Row As Long)
    GridOrcamento.Row = Row
    GridOrcamento.Col = 12

    If GridOrcamento.Text <> "" And GridOrcamento.Text <> "CodOrc" Then
        VGIntCodOrc = GridOrcamento.Text
        FrmResumo_Orcamento.Show
     End If
End Sub

Sub MontaGridOrcamento()
    Dim VLIntLinha As Long
    Dim RecVend As New ADODB.Recordset
    
    If RecPesq.EOF Then
        LblNumTotalOrc.Caption = "Nenhum orçamento encontrado."

        VPStrBox = MsgBox("Pesquisa sem resultados.", vbInformation, "Pró Vendas 2004 - Informação")
        GridOrcamento.Refresh
        GridOrcamento.MaxRows = 0

        CmdAlterarOrc.Enabled = False
        CmdExcluirOrc.Enabled = False
        CmdImprimirOrc.Enabled = False

    Else

        VLIntLinha = 1
        GridOrcamento.MaxRows = VLIntLinha

        Do While Not RecPesq.EOF

            GridOrcamento.Row = VLIntLinha
            GridOrcamento.Lock = True

            'Data
            GridOrcamento.Col = 1
            GridOrcamento.Text = FormataData(RecPesq!DtOrc)
            GridOrcamento.Lock = True

            'Cliente
            GridOrcamento.Col = 2
            GridOrcamento.Text = VerificaNulo(RecPesq!nome)
            GridOrcamento.Lock = True
            
            'Vendedor
            GridOrcamento.Col = 3
            StrSql = "Select Nome from tb_Vendedor where CodVendedor=" & RecPesq!CodVendedor
            RecVend.Open StrSql, vgCon, 1, 3
            
            If Not RecVend.EOF Then
                GridOrcamento.Text = RecVend!nome
            Else
                GridOrcamento.Text = ""
            End If
            GridOrcamento.Lock = True

            'Telefone
            GridOrcamento.Col = 4
            If (IsNull(RecPesq!telefone1) = True Or RecPesq!telefone1 = "") And (IsNull(RecPesq!telefone2) = True Or RecPesq!telefone2 = "") Then
                GridOrcamento.Text = ""
                
            ElseIf (IsNull(RecPesq!telefone1) = False Or RecPesq!telefone1 <> "") And (IsNull(RecPesq!telefone2) = True Or RecPesq!telefone2 = "") Then
                GridOrcamento.Text = RecPesq!telefone1
                
            ElseIf (IsNull(RecPesq!telefone1) = True Or RecPesq!telefone1 = "") And (IsNull(RecPesq!telefone2) = False Or RecPesq!telefone2 <> "") Then
                GridOrcamento.Text = RecPesq!telefone2
            
            ElseIf (IsNull(RecPesq!telefone1) = False Or RecPesq!telefone1 <> "") And (IsNull(RecPesq!telefone2) = False Or RecPesq!telefone2 <> "") Then
                GridOrcamento.Text = RecPesq!telefone1 & "/" & RecPesq!telefone2
                
            End If
            GridOrcamento.Lock = True

            'Total da venda
            GridOrcamento.Col = 5
            GridOrcamento.Text = FormataMoeda(VerificaNulo(RecPesq!totalvenda))
            GridOrcamento.Lock = True

            'Parcelado
            GridOrcamento.Col = 6
            GridOrcamento.Text = FormataNum(RecPesq!parcela) & " vezes"
            GridOrcamento.Lock = True

            'Entrada
            GridOrcamento.Col = 7
            GridOrcamento.Text = FormataMoeda(VerificaNulo(RecPesq!entrada))
            GridOrcamento.Lock = True

            'Valor da parcela
            GridOrcamento.Col = 8
            GridOrcamento.Text = FormataMoeda(VerificaNulo(RecPesq!valorparc))
            GridOrcamento.Lock = True

            'Valor a prazo
            GridOrcamento.Col = 9
            GridOrcamento.Text = FormataMoeda(VerificaNulo(RecPesq!valorprazo))
            GridOrcamento.Lock = True

            'Validade
            GridOrcamento.Col = 10
            GridOrcamento.Text = FormataData(RecPesq!validade)
            GridOrcamento.Lock = True

            'Observação
            GridOrcamento.Col = 11
            GridOrcamento.Text = VerificaNulo(RecPesq!obs)
            GridOrcamento.Lock = True

            'CodOrc
            GridOrcamento.Col = 12
            GridOrcamento.Text = Val(RecPesq!CodOrc)
            GridOrcamento.Lock = True

            VLIntLinha = VLIntLinha + 1

            GridOrcamento.MaxRows = GridOrcamento.MaxRows + 1
            RecVend.Close
            RecPesq.MoveNext
         Loop

         '===== CONTAGEM DE CLIENTES PESQUISADOS =========
         GridOrcamento.MaxRows = GridOrcamento.MaxRows - 1

         If GridOrcamento.MaxRows = 1 Then
            LblNumTotalOrc.Caption = FormataNum(GridOrcamento.MaxRows) & " orçamento encontrado."
         Else
            LblNumTotalOrc.Caption = FormataNum(GridOrcamento.MaxRows) & " orçamentos encontrados."
         End If
         '================================================
    End If

End Sub

Private Sub TxtDtOrc1_GotFocus()
    If TxtDtOrc1.Text = "__/__/____" Then
        TxtDtOrc1.Text = ""
    End If
    TxtDtOrc1.SelStart = 0
    TxtDtOrc1.SelLength = Len(TxtDtOrc1.Text)
End Sub

Private Sub TxtDtOrc1_KeyPress(KeyAscii As Integer)
    '=== Só aceita números e barra ===
    If (KeyAscii < 47 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtDtOrc1_LostFocus()
    Dim VLStrData As String

    If TxtDtOrc1.Text <> "" Then
        VLStrData = VerificaData(TxtDtOrc1.Text)

        If VGStrDataErro = "sim" Then
            TxtDtOrc1.SetFocus
        Else
            TxtDtOrc1.Text = VLStrData
        End If

        VGStrDataErro = ""
    Else
        TxtDtOrc1.Text = "__/__/____"
    End If
End Sub

Private Sub TxtDtOrc2_GotFocus()
    If TxtDtOrc2.Text = "__/__/____" Then
        TxtDtOrc2.Text = ""
    End If
    TxtDtOrc2.SelStart = 0
    TxtDtOrc2.SelLength = Len(TxtDtOrc2.Text)
End Sub

Private Sub TxtDtOrc2_KeyPress(KeyAscii As Integer)
    '=== Só aceita números e barra ===
    If (KeyAscii < 47 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtDtOrc2_LostFocus()
    Dim VLStrData As String

    If TxtDtOrc2.Text <> "" Then
        VLStrData = VerificaData(TxtDtOrc2.Text)

        If VGStrDataErro = "sim" Then
            TxtDtOrc2.SetFocus
        Else
            TxtDtOrc2.Text = VLStrData
        End If

        VGStrDataErro = ""
    Else
        TxtDtOrc2.Text = "__/__/____"
    End If
End Sub

Private Sub TxtTelOrc_KeyPress(KeyAscii As Integer)
    '=== Só aceita números, parênteses, espaço e - ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 And KeyAscii <> 40 And KeyAscii <> 41 And KeyAscii <> 32 And KeyAscii <> 45 Then
        KeyAscii = 0
    End If
End Sub
'========================================================================
'========================================================================




'========================================================================
'                            VENDEDOR
'========================================================================
Private Sub CmdLimparVendedor_Click()
    TxtNomeVend.Text = ""
    GridVendedor.MaxRows = 0
    LblNumTotalVendedor.Caption = "Nenhum vendedor encontrado."
    
    CmdAlterarVend.Enabled = False
    CmdExcluirVend.Enabled = False
    CmdImprimirVend.Enabled = False
End Sub

Private Sub CmdPesqVendedor_Click()
    Screen.MousePointer = vbHourglass
    
    Dim VLStrOrder As String
    
    Conecta
    
    StrSql = "Select * from tb_Vendedor where 0=0"
            
    '====== PESQUISAR POR NOME DO VENDEDOR ==========
    If TxtNomeVend.Text <> "" Then
        StrSql = StrSql + " and Nome like '%" & TxtNomeVend.Text & "%' order by Nome"
    End If
    RecPesq.Open StrSql, vgCon, 1, 3
            
    Call MontaGridVendedor
        
    Desconecta
    
    Screen.MousePointer = vbNormal

End Sub

Sub MontaGridVendedor()
    Dim VLIntLinha As Long
    
    If RecPesq.EOF Then
        LblNumTotalVendedor.Caption = "Nenhum vendedor encontrado."
        
        VPStrBox = MsgBox("Pesquisa sem resultados.", vbInformation, "Pró Vendas 2004 - Informação")
        GridVendedor.Refresh
        GridVendedor.MaxRows = 0
        
        CmdAlterarVend.Enabled = False
        CmdExcluirVend.Enabled = False
        CmdImprimirVend.Enabled = False
    
    Else
    
        VLIntLinha = 1
        GridVendedor.MaxRows = VLIntLinha
         
        Do While Not RecPesq.EOF
                 
            GridVendedor.Row = VLIntLinha
            GridVendedor.Lock = True
            
            'Vendedor
            GridVendedor.Col = 1
            GridVendedor.Text = VerificaNulo(RecPesq!nome)
            GridVendedor.Lock = True
            
            'Telefone
            GridVendedor.Col = 2
            GridVendedor.Text = VerificaNulo(RecPesq!telefone)
            GridVendedor.Lock = True
            
            'CodVendedor
            GridVendedor.Col = 3
            GridVendedor.Text = Val(RecPesq!CodVendedor)
            GridVendedor.Lock = True
            
            VLIntLinha = VLIntLinha + 1
            
            GridVendedor.MaxRows = GridVendedor.MaxRows + 1
            RecPesq.MoveNext
         Loop
         
         '===== CONTAGEM DE VENDEDORES PESQUISADOS =========
         GridVendedor.MaxRows = GridVendedor.MaxRows - 1
         
         If GridVendedor.MaxRows = 1 Then
            LblNumTotalVendedor.Caption = FormataNum(GridVendedor.MaxRows) & " vendedor encontrado."
         Else
            LblNumTotalVendedor.Caption = FormataNum(GridVendedor.MaxRows) & " vendedores encontrados."
         End If
         '================================================
         
         CmdImprimirVend.Enabled = True
    End If

End Sub

Private Sub GridVendedor_Click(ByVal Col As Long, ByVal Row As Long)
    GridVendedor.Row = Row
    GridVendedor.Col = 3
    If GridVendedor.Text <> "" And GridVendedor.Text <> "CodVendedor" Then
        VGIntCodVend = GridVendedor.Text
        CmdAlterarVend.Enabled = True
        CmdExcluirVend.Enabled = True
        CmdImprimirVend.Enabled = True
    Else
        CmdAlterarVend.Enabled = False
        CmdExcluirVend.Enabled = False
        CmdImprimirVend.Enabled = False
    End If
End Sub

Private Sub GridVendedor_DblClick(ByVal Col As Long, ByVal Row As Long)
    GridVendedor.Row = Row
    GridVendedor.Col = 3

    If GridVendedor.Text <> "" And GridVendedor.Text <> "CodVendedor" Then
        VGIntCodVend = GridVendedor.Text
        FrmResumo_Vendedor.Show
     End If
End Sub

Private Sub CmdIncluirVend_Click()
    FrmVendedor_Inc.Show
End Sub

Private Sub CmdAlterarVend_Click()
    FrmVendedor_Alt.Show
End Sub

Private Sub CmdExcluirVend_Click()
    VPStrResponse = MsgBox("Deseja excluir este vendedor?", vbYesNo, "Pró Vendas 2004 - Informação")
    
    If VPStrResponse = vbYes Then
        Conecta
        vgCon.Execute ("DELETE FROM tb_Vendedor WHERE CodVendedor=" & VGIntCodVend)
        Desconecta
        
        CmdPesqVendedor.Value = True
    End If
End Sub

Private Sub CmdImprimirVend_Click()
    Screen.MousePointer = vbHourglass
    
    Dim vend As String
    Dim tel As String
    
    Dim VLStrLinha As String
    
    VLStrLinha = 1
    
    Conecta
    
    Do While VLStrLinha <= GridVendedor.MaxRows
        
        GridVendedor.Col = 1
        GridVendedor.Row = VLStrLinha
        vend = GridVendedor.Text
        
        GridVendedor.Col = 2
        GridVendedor.Row = VLStrLinha
        tel = GridVendedor.Text
        
        vgCon.Execute "INSERT INTO tb_Auxiliar " & _
        "(campo01,campo02) " & _
        "VALUES ('" & vend & "','" & tel & "')"
         
        VLStrLinha = VLStrLinha + 1
    Loop
    
    Desconecta
        
    rptVendedor.Show

End Sub
'========================================================================
'========================================================================





'========================================================================
'                            EXTRA
'========================================================================

Private Sub OptCob_Click()
    FraNiver.Visible = False
    FraCob.Visible = True
    FraMala.Visible = False
    CboTipoCarta.SetFocus
    CmdImprimirExt.Enabled = True
End Sub

Private Sub OptMala_Click()
    FraNiver.Visible = False
    FraCob.Visible = False
    FraMala.Visible = True
    TxtCliente.SetFocus
    CmdImprimirExt.Enabled = True
End Sub

Private Sub OptNiver_Click()
    FraNiver.Visible = True
    FraCob.Visible = False
    FraMala.Visible = False
    TxtDia1.SetFocus
    CmdImprimirExt.Enabled = True
End Sub

Private Sub TxtDia1_KeyPress(KeyAscii As Integer)
    '=== Só aceita números ===
    If (KeyAscii < 47 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtDia2_KeyPress(KeyAscii As Integer)
    '=== Só aceita números ===
    If (KeyAscii < 47 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtMes1_KeyPress(KeyAscii As Integer)
    '=== Só aceita números ===
    If (KeyAscii < 47 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtMes2_KeyPress(KeyAscii As Integer)
    '=== Só aceita números ===
    If (KeyAscii < 47 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtDtVenc1_GotFocus()
    If TxtDtVenc1.Text = "__/__/____" Then
        TxtDtVenc1.Text = ""
    Else
        TxtDtVenc1.SelStart = 0
        TxtDtVenc1.SelLength = Len(TxtDtVenc1.Text)
    End If
End Sub

Private Sub TxtDtVenc1_KeyPress(KeyAscii As Integer)
    '=== Só aceita números ===
    If (KeyAscii < 47 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtDtVenc1_LostFocus()
    Dim VLStrData As String

    If TxtDtVenc1.Text <> "" Then
        VLStrData = VerificaData(TxtDtVenc1.Text)

        If VGStrDataErro = "sim" Then
            TxtDtVenc1.SetFocus
        Else
            TxtDtVenc1.Text = VLStrData
        End If

        VGStrDataErro = ""
    Else
        TxtDtVenc1.Text = "__/__/____"
    End If
End Sub

Private Sub TxtDtVenc2_GotFocus()
    If TxtDtVenc2.Text = "__/__/____" Then
        TxtDtVenc2.Text = ""
    Else
        TxtDtVenc2.SelStart = 0
        TxtDtVenc2.SelLength = Len(TxtDtVenc2.Text)
    End If
End Sub

Private Sub TxtDtVenc2_KeyPress(KeyAscii As Integer)
    '=== Só aceita números ===
    If (KeyAscii < 47 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtDtVenc2_LostFocus()
    Dim VLStrData As String

    If TxtDtVenc2.Text <> "" Then
        VLStrData = VerificaData(TxtDtVenc2.Text)

        If VGStrDataErro = "sim" Then
            TxtDtVenc2.SetFocus
        Else
            TxtDtVenc2.Text = VLStrData
        End If

        VGStrDataErro = ""
    Else
        TxtDtVenc2.Text = "__/__/____"
    End If
End Sub

Private Sub TxtDtNiverCli1_GotFocus()
    If TxtDtNiverCli1.Text = "__/__/____" Then
        TxtDtNiverCli1.Text = ""
    Else
        TxtDtNiverCli1.SelStart = 0
        TxtDtNiverCli1.SelLength = Len(TxtDtNiverCli1.Text)
    End If
End Sub

Private Sub TxtDtNiverCli1_KeyPress(KeyAscii As Integer)
    '=== Só aceita números ===
    If (KeyAscii < 47 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtDtNiverCli1_LostFocus()
    Dim VLStrData As String

    If TxtDtNiverCli1.Text <> "" Then
        VLStrData = VerificaData(TxtDtNiverCli1.Text)

        If VGStrDataErro = "sim" Then
            TxtDtNiverCli1.SetFocus
        Else
            TxtDtNiverCli1.Text = VLStrData
        End If

        VGStrDataErro = ""
    Else
        TxtDtNiverCli1.Text = "__/__/____"
    End If
End Sub

Private Sub TxtDtNiverCli2_GotFocus()
    If TxtDtNiverCli2.Text = "__/__/____" Then
        TxtDtNiverCli2.Text = ""
    Else
        TxtDtNiverCli2.SelStart = 0
        TxtDtNiverCli2.SelLength = Len(TxtDtNiverCli2.Text)
    End If
End Sub

Private Sub TxtDtNiverCli2_KeyPress(KeyAscii As Integer)
    '=== Só aceita números ===
    If (KeyAscii < 47 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtDtNiverCli2_LostFocus()
    Dim VLStrData As String

    If TxtDtNiverCli2.Text <> "" Then
        VLStrData = VerificaData(TxtDtNiverCli2.Text)

        If VGStrDataErro = "sim" Then
            TxtDtNiverCli2.SetFocus
        Else
            TxtDtNiverCli2.Text = VLStrData
        End If

        VGStrDataErro = ""
    Else
        TxtDtNiverCli2.Text = "__/__/____"
    End If
End Sub

Private Sub TxtClienteCob_GotFocus()
    TxtClienteCob.SelStart = 0
    TxtClienteCob.SelLength = Len(TxtClienteCob.Text)
End Sub

Private Sub TxtCliente_GotFocus()
    TxtCliente.SelStart = 0
    TxtCliente.SelLength = Len(TxtCliente.Text)
End Sub

Private Sub CmdImprimirExt_Click()
    Dim RecPesq As New ADODB.Recordset
    Dim RecProd As New ADODB.Recordset
    Dim RecEstq As New ADODB.Recordset
    Dim CodProdTemp As Integer
    Dim VLIntCont As Integer
    Dim VLStrGravar As String
    Dim VLStrCampo01 As String
    Dim VLStrCampo02 As String
    Dim VLStrCampo03 As String
    Dim VLStrCampo04 As String
    Dim VLStrCampo05 As String
    Dim VLStrCampo06 As String
    Dim VLStrCampo07 As String
    Dim VLStrCampo08 As String
    Dim VLStrCampo09 As String
    Dim VLStrCampo10 As String
    Dim VLStrCampo11 As String
    Dim VLStrCampo12 As String
    Dim VLStrCampo13 As String
    Dim VLStrCampo14 As String
    Dim VLStrCampo15 As String
    Dim VLStrCampo16 As String
    Dim VLStrCampo17 As String
    Dim VLStrCampo18 As String
    Dim VLStrCampo19 As String
    Dim VLStrCampo20 As String
    Dim VLStrCampo21 As String
    Dim VLStrCampo22 As String
    Dim VLStrCampo23 As String
    Dim VLStrCampo24 As String
    Dim VLStrCampo25 As String
    Dim VLStrCampo26 As String
    Dim VLStrCampo27 As String
    Dim VLStrCampo28 As String
    Dim VLStrCampo29 As String
    Dim VLStrCampo30 As String
    Dim VLStrCampo31 As String
    Dim VLStrCampo32 As String
    Dim VLStrCampo33 As String
    Dim VLStrCampo34 As String
    Dim VLStrCampo35 As String
    Dim VLStrCampo36 As String
    Dim VLStrCampo37 As String
    Dim VLStrCampo38 As String
    Dim VLStrCampo39 As String
    Dim VLStrCampo40 As String
    Dim VLStrCampo41 As String
    Dim VLStrCampo42 As String
    Dim VLStrCampo43 As String
    Dim VLStrCampo44 As String
    Dim VLStrCampo45 As String
    Dim VLStrCampo46 As String
    Dim VLStrCampo47 As String
    Dim VLStrCampo48 As String
    Dim VLStrCampo49 As String
    Dim VLStrCampo50 As String
    Dim VLStrCampo51 As String
    Dim VLStrCampo52 As String
    Dim VLStrCampo53 As String
    Dim VLStrCampo54 As String
    Dim VLStrCampo55 As String
    Dim VLStrCampo56 As String
    Dim VLStrCampo57 As String
    Dim VLStrCampo58 As String
    Dim VLStrCampo59 As String
    Dim VLStrCampo60 As String
    Dim VLStrCampo61 As String
    Dim VLStrCampo62 As String
    Dim VLStrCampo63 As String
    Dim VLStrCampo64 As String
    Dim VLStrCampo65 As String
    Dim VLStrCampo66 As String
    Dim VLStrCampo67 As String
    Dim VLStrCampo68 As String
    Dim VLStrCampo69 As String
    Dim VLStrCampo70 As String
    Dim VLStrCampo71 As String
    Dim VLStrCampo72 As String
    Dim VLStrCampo73 As String
    Dim VLIntCodCredTemp As Integer

    '============ Mala direta ============
    If OptMala.Value = True Then
        Conecta
        StrSql = "Select * from tb_Cliente where 0=0"

        '====== PESQUISAR POR CLIENTE ==========
        If TxtCliente.Text <> "" Then
            StrSql = StrSql + " and Nome like '%" & TxtCliente.Text & "%'"
        End If

        '====== PESQUISAR POR SEXO ==========
        If CboSexo.Text <> "" Then
            StrSql = StrSql + " and Sexo='" & CboSexo.Text & "'"
        End If

        '====== PESQUISAR POR DATA DE NASCIMENTO ==========
        If (TxtDtNiverCli1.Text <> "" And TxtDtNiverCli1.Text <> "__/__/____") And (TxtDtNiverCli2.Text <> "" And TxtDtNiverCli2.Text <> "__/__/____") Then
            StrSql = StrSql + " and DtNasc >=#" & FormataDataUS(TxtDtNiverCli1.Text) & "# and DtNasc <= #" & FormataDataUS(TxtDtNiverCli2.Text) & "#"

        ElseIf (TxtDtNiverCli1.Text <> "" And TxtDtNiverCli1.Text <> "__/__/____") And (TxtDtNiverCli2.Text = "" Or TxtDtNiverCli2.Text = "__/__/____") Then
            StrSql = StrSql + " and DtNasc =#" & FormataDataUS(TxtDtNiverCli1.Text) & "#"

        ElseIf (TxtDtNiverCli1.Text = "" Or TxtDtNiverCli1.Text = "__/__/____") And (TxtDtNiverCli2.Text <> "" And TxtDtNiverCli2.Text <> "__/__/____") Then
            StrSql = StrSql + " DtNasc =#" & FormataDataUS(TxtDtNiverCli2.Text) & "#"
        End If

        StrSql = StrSql + " order by Nome"
        RecPesq.Open StrSql, vgCon, 1, 3

        If RecPesq.EOF Then
            VPStrBox = MsgBox("Pesquisa sem resultados! " & Chr(13) & "Não foi possível gerar a impressão", vbInformation, "Pró Vendas 2004 - Informação")
            TxtCliente.SetFocus
            Desconecta
        Else
            VLIntCont = 1
            Do While Not RecPesq.EOF
                If VLIntCont = 1 Then
                    VLStrCampo01 = RecPesq!nome
                    VLStrCampo02 = RecPesq!endereco
                    VLStrCampo03 = RecPesq!bairro
                    VLStrCampo04 = RecPesq!cep
                    VLStrCampo05 = RecPesq!cidade & "/" & RecPesq!Estado
                    VLIntCont = 2

                ElseIf VLIntCont = 2 Then
                    VLStrCampo06 = RecPesq!nome
                    VLStrCampo07 = RecPesq!endereco
                    VLStrCampo08 = RecPesq!bairro
                    VLStrCampo09 = RecPesq!cep
                    VLStrCampo10 = RecPesq!cidade & "/" & RecPesq!Estado
                    VLIntCont = 1

                    VLStrGravar = "sim"

                End If

                RecPesq.MoveNext

                If RecPesq.EOF = True Or VLStrGravar = "sim" Then
                    vgCon.Execute "INSERT INTO tb_Auxiliar " & _
                    "(campo01,campo02,campo03,campo04,campo05,campo06,campo07,campo08,campo09,campo10) " & _
                    "VALUES ('" & VLStrCampo01 & "','" & VLStrCampo02 & "','" & VLStrCampo03 & "','" & VLStrCampo04 & "','" & VLStrCampo05 & "','" & VLStrCampo06 & "','" & VLStrCampo07 & "','" & VLStrCampo08 & "','" & VLStrCampo09 & "','" & VLStrCampo10 & "')"

                    VLStrGravar = ""
                    VLStrCampo01 = ""
                    VLStrCampo02 = ""
                    VLStrCampo03 = ""
                    VLStrCampo04 = ""
                    VLStrCampo05 = ""
                    VLStrCampo06 = ""
                    VLStrCampo07 = ""
                    VLStrCampo08 = ""
                    VLStrCampo09 = ""
                    VLStrCampo10 = ""
                End If
            Loop
            Desconecta
            
            rptExtra_Mala.Show
        End If

    '============ Cartas de cobrança ============
    ElseIf OptCob.Value = True Then
        If CboTipoCarta.Text = "" Then
            VPStrBox = MsgBox("Selecione o tipo de carta que deseja imprimir", vbInformation, "Pró Vendas 2004 - Informação")
        Else
            Conecta
            
            StrSql = "Select CR.Parcela,CR.CodCred,P.NumParc,P.Vencimento,P.Valor,C.Nome " & _
                     "From tb_Crediario as CR, tb_Crediario_Parcela as P, tb_Cliente as C " & _
                     "Where C.CodCli=CR.CodCli and CR.CodCred=P.CodCred and P.Quitado='não'"
    
            '====== PESQUISAR POR CLIENTE ==========
            If TxtClienteCob.Text <> "" Then
                StrSql = StrSql + " and C.Nome like '%" & TxtClienteCob.Text & "%'"
            End If
    
            '====== PESQUISAR POR DATA DO VENCIMENTO ==========
            If (TxtDtVenc1.Text <> "" And TxtDtVenc1.Text <> "__/__/____") And (TxtDtVenc2.Text <> "" And TxtDtVenc2.Text <> "__/__/____") Then
                StrSql = StrSql + " and P.Vencimento >=#" & FormataDataUS(TxtDtVenc1.Text) & "# and P.Vencimento <= #" & FormataDataUS(TxtDtVenc2.Text) & "#"
    
            ElseIf (TxtDtVenc1.Text <> "" And TxtDtVenc1.Text <> "__/__/____") And (TxtDtVenc2.Text = "" Or TxtDtVenc2.Text = "__/__/____") Then
                StrSql = StrSql + " and P.Vencimento =#" & FormataDataUS(TxtDtVenc1.Text) & "#"
    
            ElseIf (TxtDtVenc1.Text = "" Or TxtDtVenc1.Text = "__/__/____") And (TxtDtVenc2.Text <> "" And TxtDtVenc2.Text <> "__/__/____") Then
                StrSql = StrSql + " and P.Vencimento =#" & FormataDataUS(TxtDtVenc2.Text) & "#"
            End If
    
            StrSql = StrSql + " order by C.Nome,CR.CodCred,P.NumParc"
            RecPesq.Open StrSql, vgCon, 1, 3
    
            If RecPesq.EOF Then
                VPStrBox = MsgBox("Pesquisa sem resultados! " & Chr(13) & "Não foi possível gerar a impressão", vbInformation, "Pró Vendas 2004 - Informação")
                TxtClienteCob.SetFocus
                Desconecta
            Else
                Do While Not RecPesq.EOF
                    VLIntCont = 1
                    VLIntCodCredTemp = RecPesq!CodCred
    
                    VLStrCampo01 = RecPesq!nome
    
                    Do While (RecPesq!CodCred = VLIntCodCredTemp) And (RecPesq.EOF = False)
                        If VLIntCont = 1 Then
                            VLStrCampo02 = FormataData(RecPesq!vencimento)
                            VLStrCampo03 = FormataNum(RecPesq!numparc) & "/" & FormataNum(RecPesq!parcela)
                            VLStrCampo04 = FormataMoeda(RecPesq!valor)
                        ElseIf VLIntCont = 2 Then
                            VLStrCampo05 = FormataData(RecPesq!vencimento)
                            VLStrCampo06 = FormataNum(RecPesq!numparc) & "/" & FormataNum(RecPesq!parcela)
                            VLStrCampo07 = FormataMoeda(RecPesq!valor)
                        ElseIf VLIntCont = 3 Then
                            VLStrCampo08 = FormataData(RecPesq!vencimento)
                            VLStrCampo09 = FormataNum(RecPesq!numparc) & "/" & FormataNum(RecPesq!parcela)
                            VLStrCampo10 = FormataMoeda(RecPesq!valor)
                        ElseIf VLIntCont = 4 Then
                            VLStrCampo11 = FormataData(RecPesq!vencimento)
                            VLStrCampo12 = FormataNum(RecPesq!numparc) & "/" & FormataNum(RecPesq!parcela)
                            VLStrCampo13 = FormataMoeda(RecPesq!valor)
                        ElseIf VLIntCont = 5 Then
                            VLStrCampo14 = FormataData(RecPesq!vencimento)
                            VLStrCampo15 = FormataNum(RecPesq!numparc) & "/" & FormataNum(RecPesq!parcela)
                            VLStrCampo16 = FormataMoeda(RecPesq!valor)
                        ElseIf VLIntCont = 6 Then
                            VLStrCampo17 = FormataData(RecPesq!vencimento)
                            VLStrCampo18 = FormataNum(RecPesq!numparc) & "/" & FormataNum(RecPesq!parcela)
                            VLStrCampo19 = FormataMoeda(RecPesq!valor)
                        ElseIf VLIntCont = 7 Then
                            VLStrCampo20 = FormataData(RecPesq!vencimento)
                            VLStrCampo21 = FormataNum(RecPesq!numparc) & "/" & FormataNum(RecPesq!parcela)
                            VLStrCampo22 = FormataMoeda(RecPesq!valor)
                        ElseIf VLIntCont = 8 Then
                            VLStrCampo23 = FormataData(RecPesq!vencimento)
                            VLStrCampo24 = FormataNum(RecPesq!numparc) & "/" & FormataNum(RecPesq!parcela)
                            VLStrCampo25 = FormataMoeda(RecPesq!valor)
                        ElseIf VLIntCont = 9 Then
                            VLStrCampo26 = FormataData(RecPesq!vencimento)
                            VLStrCampo27 = FormataNum(RecPesq!numparc) & "/" & FormataNum(RecPesq!parcela)
                            VLStrCampo28 = FormataMoeda(RecPesq!valor)
                        ElseIf VLIntCont = 10 Then
                            VLStrCampo29 = FormataData(RecPesq!vencimento)
                            VLStrCampo30 = FormataNum(RecPesq!numparc) & "/" & FormataNum(RecPesq!parcela)
                            VLStrCampo31 = FormataMoeda(RecPesq!valor)
                        ElseIf VLIntCont = 11 Then
                            VLStrCampo32 = FormataData(RecPesq!vencimento)
                            VLStrCampo33 = FormataNum(RecPesq!numparc) & "/" & FormataNum(RecPesq!parcela)
                            VLStrCampo34 = FormataMoeda(RecPesq!valor)
                        ElseIf VLIntCont = 12 Then
                            VLStrCampo35 = FormataData(RecPesq!vencimento)
                            VLStrCampo36 = FormataNum(RecPesq!numparc) & "/" & FormataNum(RecPesq!parcela)
                            VLStrCampo37 = FormataMoeda(RecPesq!valor)
                        ElseIf VLIntCont = 13 Then
                            VLStrCampo38 = FormataData(RecPesq!vencimento)
                            VLStrCampo39 = FormataNum(RecPesq!numparc) & "/" & FormataNum(RecPesq!parcela)
                            VLStrCampo40 = FormataMoeda(RecPesq!valor)
                        ElseIf VLIntCont = 14 Then
                            VLStrCampo41 = FormataData(RecPesq!vencimento)
                            VLStrCampo42 = FormataNum(RecPesq!numparc) & "/" & FormataNum(RecPesq!parcela)
                            VLStrCampo43 = FormataMoeda(RecPesq!valor)
                        ElseIf VLIntCont = 15 Then
                            VLStrCampo44 = FormataData(RecPesq!vencimento)
                            VLStrCampo45 = FormataNum(RecPesq!numparc) & "/" & FormataNum(RecPesq!parcela)
                            VLStrCampo46 = FormataMoeda(RecPesq!valor)
                        ElseIf VLIntCont = 16 Then
                            VLStrCampo47 = FormataData(RecPesq!vencimento)
                            VLStrCampo48 = FormataNum(RecPesq!numparc) & "/" & FormataNum(RecPesq!parcela)
                            VLStrCampo49 = FormataMoeda(RecPesq!valor)
                        ElseIf VLIntCont = 17 Then
                            VLStrCampo50 = FormataData(RecPesq!vencimento)
                            VLStrCampo51 = FormataNum(RecPesq!numparc) & "/" & FormataNum(RecPesq!parcela)
                            VLStrCampo52 = FormataMoeda(RecPesq!valor)
                        ElseIf VLIntCont = 18 Then
                            VLStrCampo53 = FormataData(RecPesq!vencimento)
                            VLStrCampo54 = FormataNum(RecPesq!numparc) & "/" & FormataNum(RecPesq!parcela)
                            VLStrCampo55 = FormataMoeda(RecPesq!valor)
                        ElseIf VLIntCont = 19 Then
                            VLStrCampo56 = FormataData(RecPesq!vencimento)
                            VLStrCampo57 = FormataNum(RecPesq!numparc) & "/" & FormataNum(RecPesq!parcela)
                            VLStrCampo58 = FormataMoeda(RecPesq!valor)
                        ElseIf VLIntCont = 20 Then
                            VLStrCampo59 = FormataData(RecPesq!vencimento)
                            VLStrCampo60 = FormataNum(RecPesq!numparc) & "/" & FormataNum(RecPesq!parcela)
                            VLStrCampo61 = FormataMoeda(RecPesq!valor)
                        ElseIf VLIntCont = 21 Then
                            VLStrCampo62 = FormataData(RecPesq!vencimento)
                            VLStrCampo63 = FormataNum(RecPesq!numparc) & "/" & FormataNum(RecPesq!parcela)
                            VLStrCampo64 = FormataMoeda(RecPesq!valor)
                        ElseIf VLIntCont = 22 Then
                            VLStrCampo65 = FormataData(RecPesq!vencimento)
                            VLStrCampo66 = FormataNum(RecPesq!numparc) & "/" & FormataNum(RecPesq!parcela)
                            VLStrCampo67 = FormataMoeda(RecPesq!valor)
                        ElseIf VLIntCont = 23 Then
                            VLStrCampo68 = FormataData(RecPesq!vencimento)
                            VLStrCampo69 = FormataNum(RecPesq!numparc) & "/" & FormataNum(RecPesq!parcela)
                            VLStrCampo70 = FormataMoeda(RecPesq!valor)
                        ElseIf VLIntCont = 24 Then
                            VLStrCampo71 = FormataData(RecPesq!vencimento)
                            VLStrCampo72 = FormataNum(RecPesq!numparc) & "/" & FormataNum(RecPesq!parcela)
                            VLStrCampo73 = FormataMoeda(RecPesq!valor)
                        
                        End If
    
                        VLIntCont = VLIntCont + 1
    
                        RecPesq.MoveNext
    
                        If RecPesq.EOF = True Then
                            Exit Do
                        End If
                    Loop
    
                    vgCon.Execute "INSERT INTO tb_Auxiliar " & _
                    "(campo01,campo02,campo03,campo04,campo05,campo06,campo07,campo08,campo09,campo10,campo11,campo12,campo13,campo14,campo15,campo16,campo17,campo18,campo19,campo20,campo21,campo22,campo23,campo24,campo25,campo26,campo27,campo28,campo29,campo30,campo31,campo32,campo33,campo34,campo35,campo36,campo37,campo38,campo39,campo40,campo41,campo42,campo43,campo44,campo45,campo46,campo47,campo48,campo49,campo50,campo51,campo52,campo53,campo54,campo55,campo56,campo57,campo58,campo59,campo60,campo61,campo62,campo63,campo64,campo65,campo66,campo67,campo68,campo69,campo70,campo71,campo72,campo73) " & _
                    "VALUES ('" & VLStrCampo01 & "','" & VLStrCampo02 & "','" & VLStrCampo03 & "','" & VLStrCampo04 & "','" & VLStrCampo05 & "','" & VLStrCampo06 & "','" & VLStrCampo07 & "','" & VLStrCampo08 & "','" & VLStrCampo09 & "','" & VLStrCampo10 & "','" & VLStrCampo11 & "','" & VLStrCampo12 & "','" & VLStrCampo13 & "','" & VLStrCampo14 & "','" & VLStrCampo15 & "','" & VLStrCampo16 & "','" & VLStrCampo17 & "','" & VLStrCampo18 & "','" & VLStrCampo19 & "','" & VLStrCampo20 & "','" & VLStrCampo21 & "','" & VLStrCampo22 & "','" & VLStrCampo23 & "','" & VLStrCampo24 & "','" & VLStrCampo25 & "'," & _
                    "'" & VLStrCampo26 & "','" & VLStrCampo27 & "','" & VLStrCampo28 & "','" & VLStrCampo29 & "','" & VLStrCampo30 & "','" & VLStrCampo31 & "','" & VLStrCampo32 & "','" & VLStrCampo33 & "','" & VLStrCampo34 & "','" & VLStrCampo35 & "','" & VLStrCampo36 & "','" & VLStrCampo37 & "','" & VLStrCampo38 & "','" & VLStrCampo39 & "','" & VLStrCampo40 & "','" & VLStrCampo41 & "','" & VLStrCampo42 & "','" & VLStrCampo43 & "','" & VLStrCampo44 & "','" & VLStrCampo45 & "','" & VLStrCampo46 & "','" & VLStrCampo47 & "','" & VLStrCampo48 & "','" & VLStrCampo49 & "','" & VLStrCampo50 & "'," & _
                    "'" & VLStrCampo51 & "','" & VLStrCampo52 & "','" & VLStrCampo53 & "','" & VLStrCampo54 & "','" & VLStrCampo55 & "','" & VLStrCampo56 & "','" & VLStrCampo57 & "','" & VLStrCampo58 & "','" & VLStrCampo59 & "','" & VLStrCampo60 & "','" & VLStrCampo61 & "','" & VLStrCampo62 & "','" & VLStrCampo63 & "','" & VLStrCampo64 & "','" & VLStrCampo65 & "','" & VLStrCampo66 & "','" & VLStrCampo67 & "','" & VLStrCampo68 & "','" & VLStrCampo69 & "','" & VLStrCampo70 & "','" & VLStrCampo71 & "','" & VLStrCampo72 & "','" & VLStrCampo73 & "')"
    
                Loop
                Desconecta
        
                FrmExtraCobrancaAssinatura.Show
            End If
        End If
        
    '============ Aniversariantes ============
    ElseIf OptNiver.Value = True Then
        If TxtDia1.Text = "" And TxtDia2.Text = "" And TxtMes1.Text = "" And TxtMes2.Text = "" Then
            VPStrBox = MsgBox("Preencha o intervalo de dias e meses que se deseja imprimir", vbInformation, "Pró Vendas 2004 - Informação")
            TxtDia1.SetFocus

        ElseIf TxtDia1.Text = "" And TxtDia2.Text = "" Then
            VPStrBox = MsgBox("Preencha o intervalo de dias que se deseja imprimir", vbInformation, "Pró Vendas 2004 - Informação")
            TxtDia1.SetFocus

        ElseIf TxtMes1.Text = "" And TxtMes2.Text = "" Then
            VPStrBox = MsgBox("Preencha o intervalo de meses que se deseja imprimir", vbInformation, "Pró Vendas 2004 - Informação")
            TxtMes1.SetFocus

        Else
            Dim VLStrTel As String
            
            Conecta
            StrSql = "Select * from tb_Cliente where 0=0"

            '====== PESQUISAR POR DIA ==========
            If TxtDia1.Text <> "" And TxtDia2.Text <> "" Then
               StrSql = StrSql + " and Datepart('D',DtNasc) >=" & TxtDia1.Text & " and Datepart('D',DtNasc) <= " & TxtDia2.Text & ""

            ElseIf TxtDia1.Text <> "" And TxtDia2.Text = "" Then
               StrSql = StrSql + " and Datepart('D',DtNasc) =" & TxtDia1.Text & ""

            ElseIf TxtDia1.Text = "" And TxtDia2.Text <> "" Then
               StrSql = StrSql + " and Datepart('D',DtNasc) =" & TxtDia2.Text & ""

            End If

            '====== PESQUISAR POR MÊS ==========
            If TxtMes1.Text <> "" And TxtMes2.Text <> "" Then
               StrSql = StrSql + " and Datepart('M',DtNasc) >=" & TxtMes1.Text & " and Datepart('M',DtNasc) <= " & TxtMes2.Text & ""

            ElseIf TxtMes1.Text <> "" And TxtMes2.Text = "" Then
               StrSql = StrSql + " and Datepart('M',DtNasc) =" & TxtMes1.Text & ""

            ElseIf TxtMes1.Text = "" And TxtMes2.Text <> "" Then
               StrSql = StrSql + " and Datepart('M',DtNasc) =" & TxtMes2.Text & ""

            End If

            StrSql = StrSql + " order by Datepart('M',DtNasc), Datepart('D',DtNasc)"
            RecPesq.Open StrSql, vgCon, 1, 3

            If RecPesq.EOF Then
                VPStrBox = VPStrBox = MsgBox("Pesquisa sem resultados! " & Chr(13) & "Não foi possível gerar a impressão", vbInformation, "Pró Vendas 2004 - Informação")
                TxtDia1.SetFocus
                Desconecta
            Else
                Do While Not RecPesq.EOF
                    If RecPesq!telefone1 <> "" And RecPesq!telefone2 <> "" Then
                       VLStrTel = RecPesq!telefone1 & " / " & RecPesq!telefone2
                    ElseIf RecPesq!telefone1 = "" And RecPesq!telefone2 = "" Then
                       VLStrTel = ""
                    ElseIf RecPesq!telefone1 <> "" And RecPesq!telefone2 = "" Then
                       VLStrTel = RecPesq!telefone1
                    ElseIf RecPesq!telefone1 = "" And RecPesq!telefone2 <> "" Then
                       VLStrTel = RecPesq!telefone2
                    End If
                    
                    vgCon.Execute "INSERT INTO tb_Auxiliar " & _
                    "(campo01,campo02,campo03,campo04,campo05,campo06,campo07,campo08) " & _
                    "VALUES ('" & RecPesq!nome & "','" & FormataData(RecPesq!dtnasc) & "','" & RecPesq!endereco & "','" & RecPesq!bairro & "','" & RecPesq!cep & "','" & RecPesq!cidade & "/" & RecPesq!Estado & "','" & VLStrTel & "','" & RecPesq!email & "')"

                    RecPesq.MoveNext
                Loop
                Desconecta
    
                rptExtra_Niver.Show
            End If
        End If
    End If
End Sub
'========================================================================
'========================================================================


















































































'''Private Sub TxtDtVenc1_GotFocus()
'''    TxtDtVenc1.Text = ""
'''End Sub
'''
'''Private Sub TxtDtVenc1_KeyPress(KeyAscii As Integer)
'''    '=== Só aceita números ===
'''    If (KeyAscii < 47 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
'''        KeyAscii = 0
'''    End If
'''End Sub
'''
'''Private Sub TxtDtVenc1_LostFocus()
'''
'''    Dim VLStrData As String
'''
'''    If TxtDtVenc1.Text <> "" Then
'''        VLStrData = VerificaData(TxtDtVenc1.Text)
'''
'''        If VGStrDataErro = "sim" Then
'''            TxtDtVenc1.SetFocus
'''        Else
'''            TxtDtVenc1.Text = VLStrData
'''        End If
'''
'''        VGStrDataErro = ""
'''    Else
'''        TxtDtVenc1.Text = "__/__/____"
'''    End If
'''End Sub
'''
'''Private Sub TxtDtVenc2_GotFocus()
'''    TxtDtVenc2.Text = ""
'''End Sub
'''
'''Private Sub TxtDtVenc2_KeyPress(KeyAscii As Integer)
'''    '=== Só aceita números ===
'''    If (KeyAscii < 47 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
'''        KeyAscii = 0
'''    End If
'''End Sub
'''
'''Private Sub TxtDtVenc2_LostFocus()
'''
'''    Dim VLStrData As String
'''
'''    If TxtDtVenc2.Text <> "" Then
'''        VLStrData = VerificaData(TxtDtVenc2.Text)
'''
'''        If VGStrDataErro = "sim" Then
'''            TxtDtVenc2.SetFocus
'''        Else
'''            TxtDtVenc2.Text = VLStrData
'''        End If
'''
'''        VGStrDataErro = ""
'''    Else
'''        TxtDtVenc2.Text = "__/__/____"
'''    End If
'''End Sub
'''
'''
'''Private Sub TxtDtVenda1_GotFocus()
'''    If TxtDtVenda1.Text = "__/__/____" Then
'''        TxtDtVenda1.Text = ""
'''    End If
'''End Sub
'''
'''Private Sub TxtDtVenda1_KeyPress(KeyAscii As Integer)
'''    '=== Só aceita números e barra ===
'''    If (KeyAscii < 47 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
'''        KeyAscii = 0
'''    End If
'''End Sub
'''
'''Private Sub TxtDtVenda1_LostFocus()
'''    Dim VLStrData As String
'''
'''    If TxtDtVenda1.Text <> "" Then
'''        VLStrData = VerificaData(TxtDtVenda1.Text)
'''
'''        If VGStrDataErro = "sim" Then
'''            TxtDtVenda1.SetFocus
'''        Else
'''            TxtDtVenda1.Text = VLStrData
'''        End If
'''
'''        VGStrDataErro = ""
'''    Else
'''        TxtDtVenda1.Text = "__/__/____"
'''    End If
'''End Sub
'''
'''Private Sub TxtDtVenda2_GotFocus()
'''    If TxtDtVenda2.Text = "__/__/____" Then
'''        TxtDtVenda2.Text = ""
'''    End If
'''End Sub
'''
'''Private Sub TxtDtVenda2_KeyPress(KeyAscii As Integer)
'''    '=== Só aceita números e barra ===
'''    If (KeyAscii < 47 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
'''        KeyAscii = 0
'''    End If
'''End Sub
'''
'''Private Sub TxtDtVenda2_LostFocus()
'''    Dim VLStrData As String
'''
'''    If TxtDtVenda2.Text <> "" Then
'''        VLStrData = VerificaData(TxtDtVenda2.Text)
'''
'''        If VGStrDataErro = "sim" Then
'''            TxtDtVenda2.SetFocus
'''        Else
'''            TxtDtVenda2.Text = VLStrData
'''        End If
'''
'''        VGStrDataErro = ""
'''    Else
'''        TxtDtVenda2.Text = "__/__/____"
'''    End If
'''End Sub
'''
'''

Private Sub TxtVendOrc_GotFocus()
    TxtVendOrc.SelStart = 0
    TxtVendOrc.SelLength = Len(TxtVendOrc.Text)
End Sub

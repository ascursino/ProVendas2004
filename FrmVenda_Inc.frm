VERSION 5.00
Object = "{90F3D7B3-92E7-44BA-B444-6A8E2A3BC375}#1.0#0"; "actskin4.ocx"
Object = "{FDAC2480-F4ED-4632-AA78-DCA210A74E49}#6.0#0"; "SPR32X60.ocx"
Begin VB.Form FrmVenda_Inc 
   Caption         =   "Inclusão de Venda"
   ClientHeight    =   8865
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7920
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "FrmVenda_Inc.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8865
   ScaleWidth      =   7920
   Begin VB.Frame FraProduto 
      Caption         =   "Produtos"
      Height          =   3495
      Left            =   120
      TabIndex        =   44
      Top             =   600
      Width           =   7695
      Begin VB.CommandButton CmdExcluirProd 
         Caption         =   "Excluir produto"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6000
         TabIndex        =   41
         ToolTipText     =   "Excluir produto"
         Top             =   360
         Width           =   1335
      End
      Begin VB.CommandButton CmdVerProd 
         Caption         =   "Ver produto"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4440
         TabIndex        =   40
         ToolTipText     =   "Ver produto"
         Top             =   360
         Width           =   1335
      End
      Begin FPSpread.vaSpread GridProduto 
         Height          =   2655
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   7430
         _Version        =   393216
         _ExtentX        =   13106
         _ExtentY        =   4683
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
         MaxCols         =   5
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
         SpreadDesigner  =   "FrmVenda_Inc.frx":0CCA
      End
   End
   Begin VB.Frame FraFinalizVenda 
      Caption         =   "Finalização da venda"
      Height          =   3735
      Left            =   120
      TabIndex        =   45
      Top             =   4200
      Width           =   7695
      Begin VB.Frame FraVista 
         Caption         =   "À vista"
         Height          =   2895
         Left            =   120
         TabIndex        =   83
         Top             =   720
         Visible         =   0   'False
         Width           =   7455
         Begin VB.TextBox TxtBancoVista 
            Height          =   285
            Left            =   5640
            TabIndex        =   11
            ToolTipText     =   "Número do banco do cheque"
            Top             =   840
            Width           =   495
         End
         Begin VB.TextBox TxtChequeVista 
            Height          =   285
            Left            =   5640
            TabIndex        =   12
            ToolTipText     =   "Número do cheque"
            Top             =   1200
            Width           =   1215
         End
         Begin VB.TextBox TxtDigVista 
            Height          =   285
            Left            =   6960
            TabIndex        =   13
            ToolTipText     =   "Dígito do número do cheque"
            Top             =   1200
            Width           =   255
         End
         Begin VB.OptionButton OptChq 
            Caption         =   "Cheque"
            Height          =   255
            Left            =   4560
            TabIndex        =   10
            ToolTipText     =   "Pagamento em cheque"
            Top             =   480
            Width           =   975
         End
         Begin VB.OptionButton OptDin 
            Caption         =   "Dinheiro"
            Height          =   255
            Left            =   3120
            TabIndex        =   9
            ToolTipText     =   "Pagamento em dinheiro"
            Top             =   480
            Width           =   1215
         End
         Begin VB.TextBox TxtDescVista 
            Height          =   285
            Left            =   1200
            TabIndex        =   7
            ToolTipText     =   "Desconto"
            Top             =   960
            Width           =   495
         End
         Begin VB.TextBox TxtVendaVista 
            Height          =   285
            Left            =   1200
            TabIndex        =   6
            ToolTipText     =   "Valor da venda"
            Top             =   480
            Width           =   1215
         End
         Begin VB.TextBox TxtTotalVista 
            Height          =   285
            Left            =   1200
            TabIndex        =   8
            ToolTipText     =   "Valor total da venda"
            Top             =   1440
            Width           =   1215
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel38 
            Height          =   255
            Left            =   240
            OleObjectBlob   =   "FrmVenda_Inc.frx":11B8
            TabIndex        =   84
            Top             =   480
            Width           =   735
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel39 
            Height          =   255
            Left            =   240
            OleObjectBlob   =   "FrmVenda_Inc.frx":121C
            TabIndex        =   85
            Top             =   960
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel40 
            Height          =   255
            Left            =   240
            OleObjectBlob   =   "FrmVenda_Inc.frx":1286
            TabIndex        =   86
            Top             =   1440
            Width           =   615
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel41 
            Height          =   255
            Left            =   1800
            OleObjectBlob   =   "FrmVenda_Inc.frx":12EA
            TabIndex        =   87
            Top             =   960
            Width           =   255
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel42 
            Height          =   255
            Left            =   4800
            OleObjectBlob   =   "FrmVenda_Inc.frx":1344
            TabIndex        =   88
            Top             =   840
            Width           =   615
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel43 
            Height          =   255
            Left            =   4800
            OleObjectBlob   =   "FrmVenda_Inc.frx":13A8
            TabIndex        =   89
            Top             =   1200
            Width           =   735
         End
      End
      Begin VB.Frame FraPrazoCheque 
         Caption         =   "A prazo - cheque"
         Height          =   2895
         Left            =   120
         TabIndex        =   66
         Top             =   720
         Visible         =   0   'False
         Width           =   7455
         Begin VB.TextBox TxtPrazoChqJuros 
            Height          =   285
            Left            =   1080
            TabIndex        =   15
            ToolTipText     =   "Juros"
            Top             =   1200
            Width           =   495
         End
         Begin VB.ComboBox CboPrazoChqParc 
            Height          =   315
            ItemData        =   "FrmVenda_Inc.frx":140E
            Left            =   1080
            List            =   "FrmVenda_Inc.frx":1410
            Style           =   2  'Dropdown List
            TabIndex        =   17
            ToolTipText     =   "Quantidade de parcelas"
            Top             =   1920
            Width           =   735
         End
         Begin VB.TextBox TxtTotalVendaChq 
            Height          =   285
            Left            =   1080
            TabIndex        =   16
            ToolTipText     =   "Valor total da venda"
            Top             =   1560
            Width           =   1215
         End
         Begin VB.TextBox TxtVendaChq 
            Height          =   285
            Left            =   1080
            TabIndex        =   14
            ToolTipText     =   "Valor da venda"
            Top             =   840
            Width           =   1215
         End
         Begin VB.Frame Frame2 
            Caption         =   "Entrada"
            Height          =   1455
            Left            =   2880
            TabIndex        =   68
            Top             =   720
            Width           =   4455
            Begin VB.Frame FraChqEntrDin 
               Height          =   1215
               Left            =   1680
               TabIndex        =   69
               Top             =   120
               Visible         =   0   'False
               Width           =   2655
               Begin VB.TextBox TxtValorEntrDinChq 
                  Height          =   285
                  Left            =   1080
                  TabIndex        =   21
                  ToolTipText     =   "Valor da entrada"
                  Top             =   480
                  Width           =   1215
               End
               Begin ACTIVESKINLibCtl.SkinLabel SkinLabel53 
                  Height          =   255
                  Left            =   360
                  OleObjectBlob   =   "FrmVenda_Inc.frx":1412
                  TabIndex        =   70
                  Top             =   480
                  Width           =   615
               End
            End
            Begin VB.OptionButton OptChqEntrDin 
               Caption         =   "Dinheiro"
               Height          =   255
               Left            =   120
               TabIndex        =   18
               ToolTipText     =   "Entrada em dinheiro"
               Top             =   360
               Width           =   1215
            End
            Begin VB.OptionButton OptChqEntrChq 
               Caption         =   "Cheque"
               Height          =   255
               Left            =   120
               TabIndex        =   19
               ToolTipText     =   "Entrada em cheque"
               Top             =   720
               Width           =   975
            End
            Begin VB.Frame FraChqEntrChq 
               Height          =   1215
               Left            =   1680
               TabIndex        =   71
               Top             =   120
               Visible         =   0   'False
               Width           =   2655
               Begin VB.TextBox TxtBancoChq 
                  Height          =   285
                  Left            =   960
                  TabIndex        =   22
                  ToolTipText     =   "Número do banco do cheque"
                  Top             =   120
                  Width           =   495
               End
               Begin VB.TextBox TxtChequeChq 
                  Height          =   285
                  Left            =   960
                  TabIndex        =   23
                  ToolTipText     =   "Número do cheque"
                  Top             =   480
                  Width           =   1215
               End
               Begin VB.TextBox TxtDigChq 
                  Height          =   285
                  Left            =   2280
                  TabIndex        =   24
                  ToolTipText     =   "Dígito do número do cheque"
                  Top             =   480
                  Width           =   255
               End
               Begin VB.TextBox TxtValorEntrChequeChq 
                  Height          =   285
                  Left            =   960
                  TabIndex        =   25
                  ToolTipText     =   "Valor da entrada"
                  Top             =   840
                  Width           =   1215
               End
               Begin ACTIVESKINLibCtl.SkinLabel SkinLabel50 
                  Height          =   255
                  Left            =   120
                  OleObjectBlob   =   "FrmVenda_Inc.frx":1476
                  TabIndex        =   72
                  Top             =   120
                  Width           =   735
               End
               Begin ACTIVESKINLibCtl.SkinLabel SkinLabel51 
                  Height          =   255
                  Left            =   120
                  OleObjectBlob   =   "FrmVenda_Inc.frx":14DA
                  TabIndex        =   73
                  Top             =   480
                  Width           =   735
               End
               Begin ACTIVESKINLibCtl.SkinLabel SkinLabel52 
                  Height          =   255
                  Left            =   120
                  OleObjectBlob   =   "FrmVenda_Inc.frx":1540
                  TabIndex        =   74
                  Top             =   840
                  Width           =   615
               End
            End
            Begin VB.OptionButton OptChqSemEntr 
               Caption         =   "Sem entrada"
               Height          =   255
               Left            =   120
               TabIndex        =   20
               ToolTipText     =   "Sem entrada"
               Top             =   1080
               Width           =   1455
            End
         End
         Begin VB.CommandButton CmdCrediarista 
            Caption         =   "Crediarista"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   6120
            TabIndex        =   67
            ToolTipText     =   "Escolher crediarista"
            Top             =   240
            Width           =   1095
         End
         Begin VB.CommandButton CmdParcInc 
            Caption         =   "Incluir parcelas"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   5400
            TabIndex        =   26
            ToolTipText     =   "Incluir parcelas do crediário"
            Top             =   2280
            Width           =   1815
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel44 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmVenda_Inc.frx":15A4
            TabIndex        =   75
            Top             =   840
            Width           =   735
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel45 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmVenda_Inc.frx":1608
            TabIndex        =   76
            Top             =   1920
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel46 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmVenda_Inc.frx":1672
            TabIndex        =   77
            Top             =   1200
            Width           =   735
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel47 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmVenda_Inc.frx":16D6
            TabIndex        =   78
            Top             =   1560
            Width           =   615
         End
         Begin ACTIVESKINLibCtl.SkinLabel LblEntrChq 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmVenda_Inc.frx":173A
            TabIndex        =   79
            Top             =   2280
            Width           =   2535
         End
         Begin ACTIVESKINLibCtl.SkinLabel LblParcChq 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmVenda_Inc.frx":17B6
            TabIndex        =   80
            Top             =   2520
            Width           =   5055
         End
         Begin ACTIVESKINLibCtl.SkinLabel LblCredstaCheque 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmVenda_Inc.frx":1842
            TabIndex        =   81
            Top             =   360
            Width           =   5775
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel1 
            Height          =   255
            Left            =   1680
            OleObjectBlob   =   "FrmVenda_Inc.frx":18B2
            TabIndex        =   82
            Top             =   1200
            Width           =   255
         End
      End
      Begin VB.Frame FraPrazoCarne 
         Caption         =   "A prazo - carnê"
         Height          =   2895
         Left            =   120
         TabIndex        =   48
         Top             =   720
         Visible         =   0   'False
         Width           =   7455
         Begin VB.CommandButton CmdCrediarista 
            Caption         =   "Crediarista"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   6120
            TabIndex        =   64
            ToolTipText     =   "Escolher crediarista"
            Top             =   240
            Width           =   1095
         End
         Begin VB.Frame Frame5 
            Caption         =   "Entrada"
            Height          =   1455
            Left            =   2880
            TabIndex        =   49
            Top             =   720
            Width           =   4455
            Begin VB.Frame FraCarEntrDin 
               Height          =   1215
               Left            =   1680
               TabIndex        =   55
               Top             =   120
               Visible         =   0   'False
               Width           =   2655
               Begin VB.TextBox TxtValorEntrDinCar 
                  Height          =   285
                  Left            =   1080
                  TabIndex        =   34
                  ToolTipText     =   "Valor da entrada"
                  Top             =   480
                  Width           =   1215
               End
               Begin ACTIVESKINLibCtl.SkinLabel SkinLabel64 
                  Height          =   255
                  Left            =   360
                  OleObjectBlob   =   "FrmVenda_Inc.frx":190C
                  TabIndex        =   56
                  Top             =   480
                  Width           =   615
               End
            End
            Begin VB.OptionButton OptCarSemEntr 
               Caption         =   "Sem entrada"
               Height          =   255
               Left            =   120
               TabIndex        =   33
               ToolTipText     =   "Sem entrada"
               Top             =   1080
               Width           =   1455
            End
            Begin VB.OptionButton OptCarEntrDin 
               Caption         =   "Dinheiro"
               Height          =   255
               Left            =   120
               TabIndex        =   31
               ToolTipText     =   "Entrada em dinheiro"
               Top             =   360
               Width           =   1215
            End
            Begin VB.OptionButton OptCarEntrChq 
               Caption         =   "Cheque"
               Height          =   255
               Left            =   120
               TabIndex        =   32
               ToolTipText     =   "Entrada em cheque"
               Top             =   720
               Width           =   975
            End
            Begin VB.Frame FraCarEntrChq 
               Height          =   1215
               Left            =   1680
               TabIndex        =   50
               Top             =   120
               Visible         =   0   'False
               Width           =   2655
               Begin VB.TextBox TxtValorEntrChequeCar 
                  Height          =   285
                  Left            =   960
                  TabIndex        =   37
                  ToolTipText     =   "Valor da entrada"
                  Top             =   840
                  Width           =   1215
               End
               Begin VB.TextBox TxtBancoCar 
                  Height          =   285
                  Left            =   960
                  TabIndex        =   35
                  ToolTipText     =   "Número do banco do cheque"
                  Top             =   120
                  Width           =   495
               End
               Begin VB.TextBox TxtChequeCar 
                  Height          =   285
                  Left            =   960
                  TabIndex        =   36
                  ToolTipText     =   "Número do cheque"
                  Top             =   480
                  Width           =   1215
               End
               Begin VB.TextBox TxtDigCar 
                  Height          =   285
                  Left            =   2280
                  TabIndex        =   51
                  ToolTipText     =   "Dígito do número do cheque"
                  Top             =   480
                  Width           =   255
               End
               Begin ACTIVESKINLibCtl.SkinLabel SkinLabel65 
                  Height          =   255
                  Left            =   120
                  OleObjectBlob   =   "FrmVenda_Inc.frx":1970
                  TabIndex        =   52
                  Top             =   120
                  Width           =   615
               End
               Begin ACTIVESKINLibCtl.SkinLabel SkinLabel66 
                  Height          =   255
                  Left            =   120
                  OleObjectBlob   =   "FrmVenda_Inc.frx":19D4
                  TabIndex        =   53
                  Top             =   480
                  Width           =   735
               End
               Begin ACTIVESKINLibCtl.SkinLabel SkinLabel67 
                  Height          =   255
                  Left            =   120
                  OleObjectBlob   =   "FrmVenda_Inc.frx":1A3A
                  TabIndex        =   54
                  Top             =   840
                  Width           =   615
               End
            End
         End
         Begin VB.TextBox TxtPrazoCarJuros 
            Height          =   285
            Left            =   1080
            TabIndex        =   28
            ToolTipText     =   "Juros"
            Top             =   1200
            Width           =   495
         End
         Begin VB.ComboBox CboPrazoCarParc 
            Height          =   315
            ItemData        =   "FrmVenda_Inc.frx":1A9E
            Left            =   1080
            List            =   "FrmVenda_Inc.frx":1AA0
            Style           =   2  'Dropdown List
            TabIndex        =   30
            ToolTipText     =   "Quantidade de parcelas"
            Top             =   1920
            Width           =   735
         End
         Begin VB.TextBox TxtVendaCar 
            Height          =   285
            Left            =   1080
            TabIndex        =   27
            ToolTipText     =   "Valor da venda"
            Top             =   840
            Width           =   1215
         End
         Begin VB.TextBox TxtTotalVendaCar 
            Height          =   285
            Left            =   1080
            TabIndex        =   29
            ToolTipText     =   "Valor total da venda"
            Top             =   1560
            Width           =   1215
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel58 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmVenda_Inc.frx":1AA2
            TabIndex        =   57
            Top             =   840
            Width           =   735
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel59 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmVenda_Inc.frx":1B06
            TabIndex        =   58
            Top             =   1920
            Width           =   975
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel60 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmVenda_Inc.frx":1B70
            TabIndex        =   59
            Top             =   1200
            Width           =   735
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel61 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmVenda_Inc.frx":1BD4
            TabIndex        =   60
            Top             =   1560
            Width           =   615
         End
         Begin ACTIVESKINLibCtl.SkinLabel LblEntrCar 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmVenda_Inc.frx":1C38
            TabIndex        =   61
            Top             =   2280
            Width           =   2535
         End
         Begin ACTIVESKINLibCtl.SkinLabel LblParcCar 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmVenda_Inc.frx":1CB4
            TabIndex        =   62
            Top             =   2520
            Width           =   6495
         End
         Begin ACTIVESKINLibCtl.SkinLabel LblCredstaCarne 
            Height          =   255
            Left            =   120
            OleObjectBlob   =   "FrmVenda_Inc.frx":1D40
            TabIndex        =   63
            Top             =   360
            Width           =   5775
         End
         Begin ACTIVESKINLibCtl.SkinLabel SkinLabel4 
            Height          =   255
            Left            =   1680
            OleObjectBlob   =   "FrmVenda_Inc.frx":1DB0
            TabIndex        =   65
            Top             =   1200
            Width           =   255
         End
      End
      Begin VB.Frame FraTipoPrazo 
         Height          =   495
         Left            =   4680
         TabIndex        =   47
         Top             =   120
         Visible         =   0   'False
         Width           =   2175
         Begin VB.OptionButton OptPrazoCarne 
            Caption         =   "Carnê"
            Height          =   255
            Left            =   1200
            TabIndex        =   5
            ToolTipText     =   "Venda a prazo em carnê"
            Top             =   120
            Width           =   855
         End
         Begin VB.OptionButton OptPrazoCheque 
            Caption         =   "Cheque"
            Height          =   255
            Left            =   120
            TabIndex        =   4
            ToolTipText     =   "Venda a prazo em cheque"
            Top             =   120
            Width           =   975
         End
      End
      Begin VB.OptionButton OptPrazo 
         Caption         =   "A prazo"
         Height          =   255
         Left            =   3600
         TabIndex        =   3
         ToolTipText     =   "Venda a prazo"
         Top             =   240
         Width           =   975
      End
      Begin VB.OptionButton OptVista 
         Caption         =   "À vista"
         Height          =   255
         Left            =   2520
         TabIndex        =   2
         ToolTipText     =   "Venda à vista"
         Top             =   240
         Width           =   975
      End
      Begin ACTIVESKINLibCtl.SkinLabel SkinLabel37 
         Height          =   255
         Left            =   960
         OleObjectBlob   =   "FrmVenda_Inc.frx":1E0A
         TabIndex        =   46
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.ComboBox CboVendedor 
      Height          =   315
      Left            =   1200
      Style           =   2  'Dropdown List
      TabIndex        =   0
      ToolTipText     =   "Nome do vendedor"
      Top             =   120
      Width           =   6495
   End
   Begin VB.Frame FraBotaoCli 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   42
      Top             =   8040
      Width           =   7695
      Begin ACTIVESKINLibCtl.Skin Skin1 
         Left            =   240
         OleObjectBlob   =   "FrmVenda_Inc.frx":1E7E
         Top             =   120
      End
      Begin VB.CommandButton CmdFechar 
         Caption         =   "&Fechar"
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
         Left            =   6360
         TabIndex        =   39
         ToolTipText     =   "Fechar"
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton CmdOk 
         Caption         =   "&Ok"
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
         Left            =   5040
         TabIndex        =   38
         ToolTipText     =   "Efetuar inclusão"
         Top             =   240
         Width           =   1095
      End
   End
   Begin ACTIVESKINLibCtl.SkinLabel SkinLabel2 
      Height          =   255
      Left            =   120
      OleObjectBlob   =   "FrmVenda_Inc.frx":20B2
      TabIndex        =   43
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "FrmVenda_Inc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public VPStrBox As String
Public VPStrResponse As String
Public VPIntParcela As Integer
Public VPIntParcTemp As Integer
Public VPStrValorTotal As String
Public VPStrVenda As String
'Public VPIntProd As Integer
Public VPIntCodCredTemp As Long

Private Sub CboPrazoCarParc_Click()
    Dim restparc As String
    
    If TxtTotalVendaCar.Text = "" Then
        TxtTotalVendaCar.Text = TxtVendaCar.Text
    End If
    
    If TxtTotalVendaCar.Text <> "" Then
        restparc = CCur(TxtTotalVendaCar.Text) - ((CCur(TxtTotalVendaCar.Text) * 20) / 100)
        
        LblEntrCar.Caption = "Entrada: " & FormataMoeda(CCur((CCur(TxtTotalVendaCar.Text) * 20) / 100))
        
        If restparc = "0" And CboPrazoCarParc.Text = "00" Then
            LblParcCar.Caption = "00 parcela(s) de R$ 0,00"
        Else
            LblParcCar.Caption = CboPrazoCarParc.Text & " parcela(s) de " & FormataMoeda(CCur(restparc / CboPrazoCarParc.Text))
        End If
    End If
End Sub

Private Sub CboPrazoChqParc_Click()
    Dim restparc As String
    
    If TxtTotalVendaChq.Text = "" Then
        TxtTotalVendaChq.Text = TxtVendaChq.Text
    End If
    
    If TxtTotalVendaChq.Text <> "" Then
        restparc = CCur(TxtTotalVendaChq.Text) - ((CCur(TxtTotalVendaChq.Text) * 20) / 100)
        
        LblEntrChq.Caption = "Entrada: " & FormataMoeda(CCur((CCur(TxtTotalVendaChq.Text) * 20) / 100))
        
        If restparc = "0" Then
            LblParcChq.Caption = "00 parcela(s) de R$ 0,00"
        Else
            LblParcChq.Caption = CboPrazoChqParc.Text & " parcela(s) de " & FormataMoeda(CCur(restparc / CboPrazoChqParc.Text))
        End If
    
    End If
End Sub

Private Sub CmdCrediarista_Click(Index As Integer)
    VGStrCredLista = "venda"
    FrmCrediarista_Lista.Show
End Sub

Private Sub CmdFechar_Click()
    Unload Me
    
    MDIPrincipal.Enabled = True
    MDIPrincipal.WindowState = 2
End Sub

Private Sub CmdExcluirProd_Click()
    If GridProduto.MaxRows = 0 Then
        CmdExcluirProd.Enabled = False
    Else
        CmdExcluirProd.Enabled = True
        GridProduto.DeleteRows GridProduto.ActiveRow, 1
        GridProduto.MaxRows = GridProduto.MaxRows - 1
        CmdExcluirProd.Enabled = False
    End If
End Sub

Private Sub CmdOK_Click()
    If OptPrazo.Value = True And LblCredstaCheque.Caption = "Crediarista:" And LblCredstaCarne.Caption = "Crediarista:" Then
        VPStrBox = MsgBox("Você deve escolher um crediarista para este crediário.", vbInformation, "Pró Vendas 2004 - Informação")
    
    ElseIf OptPrazoCheque.Value = True And VGStrParcelaCheque = "" Then
        VPStrBox = MsgBox("Você deve incluir as parcelas dos cheques.", vbInformation, "Pró Vendas 2004 - Informação")
    
    ElseIf GridProduto.MaxRows = 0 Then
        VPStrBox = MsgBox("Você deve escolher pelo menos 1(um) produto para efetuar a venda.", vbInformation, "Pró Vendas 2004 - Informação")
    
    ElseIf OptVista.Value = False And OptPrazo.Value = False Then
        VPStrBox = MsgBox("Escolha o tipo de venda.", vbInformation, "Pró Vendas 2004 - Informação")
    
    Else
        VGStrParcelaCheque = ""
        
        Screen.MousePointer = vbHourglass
        
        Dim RecVenda As New ADODB.Recordset
        Dim RecVendaProd As New ADODB.Recordset
        Dim RecEst As New ADODB.Recordset
        Dim RecCx As New ADODB.Recordset
        Dim RecCred As New ADODB.Recordset
        Dim RecCredParc As New ADODB.Recordset
        
        Dim VLIntCodVendaTemp As Long
        Dim parcelatemp As Integer
        Dim VLIntCodCred As Long
        Dim VLIntLinha As Long
        
        Dim VLStrValorProd As String

        parcelatemp = 1
        
        Conecta
        
        If VPStrVenda = "vista" Then
            
            '============== INCLUIR VENDA =============================
            StrSql = "SELECT * FROM tb_Venda"
            RecVenda.Open StrSql, vgCon, 1, 3
                
            RecVenda.AddNew
            If CboVendedor.Text = "" Then
                RecVenda("CodVendedor") = 0
            Else
                RecVenda("CodVendedor") = Trim(Mid(CboVendedor.Text, Len(CboVendedor.Text) - 10))
            End If
            RecVenda("CodCli") = VGIntCodCli
            RecVenda("CodCred") = 0
            RecVenda("DtVenda") = FormataDataUS(Date)
            RecVenda("TipoVenda") = "À vista"
            RecVenda("SubTotalVenda") = Trim(Mid(TxtVendaVista.Text, 3))
            RecVenda("Desconto") = TxtDescVista.Text
            RecVenda("TotalVenda") = Trim(Mid(TxtTotalVista.Text, 3))
            
            If OptDin.Value = True Then
                RecVenda("TipoPagto") = "Dinheiro"
                RecVenda("NumBanco") = 0
                RecVenda("NumCheque") = ""
            
            ElseIf OptChq.Value = True Then
                RecVenda("TipoPagto") = "Cheque"
                RecVenda("NumBanco") = TxtBancoVista.Text
                If TxtDigVista.Text <> "" Then
                    RecVenda("NumCheque") = TxtChequeVista.Text & "-" & TxtDigVista.Text
                Else
                    RecVenda("NumCheque") = TxtChequeVista.Text
                End If
            End If
            RecVenda.Update
            
            RecVenda.Close
            
            'Pegando Código da venda
            StrSql = "SELECT MAX(CodVenda) as CodVenda FROM tb_Venda"
            RecVenda.Open StrSql, vgCon, 1, 3
            
            VLIntCodVendaTemp = RecVenda!codvenda
            VGIntCodVendaRel = RecVenda!codvenda
            
            RecVenda.Close
            
            '============== INCLUIR VENDA-PRODUTO =============================
            StrSql = "SELECT * FROM tb_Venda_Produto"
            RecVendaProd.Open StrSql, vgCon, 1, 3
            
            VLIntLinha = 1
            
            Do While VLIntLinha <= GridProduto.MaxRows
                GridProduto.Row = VLIntLinha
                
                RecVendaProd.AddNew
                RecVendaProd("CodVenda") = VLIntCodVendaTemp
                
                GridProduto.Col = 5
                RecVendaProd("CodProd") = GridProduto.Text
                
                GridProduto.Col = 3
                RecVendaProd("Qtde") = GridProduto.Text
                
                GridProduto.Col = 4
                RecVendaProd("ValorProd") = GridProduto.Text
                RecVendaProd.Update
                
                VLIntLinha = VLIntLinha + 1
            Loop
            
            RecVendaProd.Close
            
            '================ RETIRAR QTDE DO ESTOQUE ========================
            VLIntLinha = 1
            
            Do While VLIntLinha <= GridProduto.MaxRows
                GridProduto.Row = VLIntLinha
                
                GridProduto.Col = 5
                StrSql = "SELECT QtdeProd FROM tb_Estoque where CodProd=" & GridProduto.Text
                RecEst.Open StrSql, vgCon, 1, 3
                
                If Not RecEst.EOF Then
                    GridProduto.Col = 3
                    RecEst("QtdeProd") = Int(RecEst!qtdeprod) - Int(GridProduto.Text)
                    RecEst.Update
                End If
                
                RecEst.Close
                
                VLIntLinha = VLIntLinha + 1
            Loop
                
            '================= INCLUIR NO MOVIMENTO DE CAIXA =========================
            
            StrSql = "SELECT * FROM tb_Caixa"
            RecCx.Open StrSql, vgCon, 1, 3
            
            RecCx.AddNew
            RecCx("CodVenda") = VLIntCodVendaTemp
            RecCx("DtMov") = FormataDataUS(Date)
            RecCx("TipoMov") = "Venda à vista"
            RecCx("Valor") = CCur(TxtTotalVista.Text)
            RecCx("TipoValor") = "credito"
            RecCx("Descricao") = "Venda à vista - Cliente: " & VGStrNomeCli
            
            If OptDin.Value = True Then
                RecCx("TipoPagto") = "Dinheiro"
            ElseIf OptChq.Value = True Then
                RecCx("TipoPagto") = "Cheque"
            End If
            
            RecCx.Update
        
            Desconecta
            
            VPStrBox = MsgBox("Venda efetuada.", vbInformation, "Pró Vendas 2004 - Informação")
            
            Unload Me
            
            MDIPrincipal.Enabled = True
            MDIPrincipal.WindowState = 2
            
            Screen.MousePointer = vbNormal
        
        ElseIf VPStrVenda = "prazocheque" Then
            
            '============== INCLUIR CREDIÁRIO =============================
            StrSql = "SELECT * FROM tb_Crediario"
            RecCred.Open StrSql, vgCon, 1, 3
            
            RecCred.AddNew
            RecCred("CodCredsta") = VGIntCodCredstaVenda
            RecCred("CodCli") = VGIntCodCli
            RecCred("DtCred") = FormataDataUS(Date)
            RecCred("TipoCred") = "Cheque"
            RecCred("ValorVenda") = CCur(TxtVendaChq.Text)
            RecCred("Parcela") = CboPrazoChqParc.Text
            RecCred("Juros") = TxtPrazoChqJuros.Text
            RecCred("ValorTotal") = CCur(TxtTotalVendaChq.Text)
            
            If OptChqSemEntr.Value = True Then
                RecCred("TipoEntr") = "Sem entrada"
                RecCred("ValorEntr") = ""
                RecCred("NumBanco") = 0
                RecCred("NumCheque") = ""
            
            ElseIf OptChqEntrDin.Value = True Then
                RecCred("TipoEntr") = "Dinheiro"
                RecCred("ValorEntr") = CCur(TxtValorEntrDinChq.Text)
                RecCred("NumBanco") = 0
                RecCred("NumCheque") = ""
            
            ElseIf OptChqEntrChq.Value = True Then
                RecCred("TipoEntr") = "Cheque"
                RecCred("ValorEntr") = CCur(TxtValorEntrChequeChq.Text)
                
                If TxtBancoCar.Text <> "" Then
                    RecCred("NumBanco") = TxtBancoCar.Text
                Else
                    RecCred("NumBanco") = 0
                End If
                
                If TxtChequeChq.Text <> "" And TxtDigChq.Text <> "" Then
                    RecCred("NumCheque") = TxtChequeChq.Text & "-" & TxtDigChq.Text
                ElseIf TxtChequeChq.Text <> "" And TxtDigChq.Text = "" Then
                    RecCred("NumCheque") = TxtChequeChq.Text
                ElseIf TxtChequeChq.Text = "" And TxtDigChq.Text = "" Then
                    RecCred("NumCheque") = ""
                End If
            End If
            
            RecCred.Update
            
            RecCred.Close
            
            StrSql = "SELECT MAX(CodCred) as CodCred FROM tb_Crediario where CodCli=" & VGIntCodCli
            RecCred.Open StrSql, vgCon, 1, 3
            
            VPIntCodCredTemp = RecCred!CodCred
            
            '============== INCLUIR PARCELAS DO CREDIÁRIO =============================
            If CboPrazoChqParc.Text <> "" Then
                Do While parcelatemp <= Val(CboPrazoChqParc.Text)
                
                    StrSql = "SELECT * FROM tb_Crediario_Parcela"
                    RecCredParc.Open StrSql, vgCon, 1, 3
                     
                    RecCredParc.AddNew
                    RecCredParc("CodCred") = RecCred!CodCred
                    RecCredParc("NumParc") = parcelatemp
                    
                    If VGStrData <> "" And VGStrData <> "#" Then
                        VGStrData = Trim(Mid(VGStrData, InStr(VGStrData, "#") + 1))
                        RecCredParc("Vencimento") = FormataDataUS(Trim(Mid(VGStrData, 1, InStr(VGStrData, "#") - 1)))
                        VGStrData = Trim(Mid(VGStrData, InStr(VGStrData, "#")))
                    Else
                        RecCredParc("Vencimento") = Null
                    End If
                    
                    If VGStrValor <> "" And VGStrValor <> "#" Then
                        VGStrValor = Trim(Mid(VGStrValor, InStr(VGStrValor, "#") + 1))
                        RecCredParc("Valor") = CCur(Trim(Mid(VGStrValor, 1, InStr(VGStrValor, "#") - 1)))
                        VGStrValor = Trim(Mid(VGStrValor, InStr(VGStrValor, "#")))
                    Else
                        RecCredParc("Valor") = ""
                    End If
                    
                    RecCredParc("Quitado") = "não"
                    
                    If VGStrBanco <> "" And VGStrBanco <> "#" Then
                        VGStrBanco = Trim(Mid(VGStrBanco, InStr(VGStrBanco, "#") + 1))
                        RecCredParc("NumBanco") = Trim(Mid(VGStrBanco, 1, InStr(VGStrBanco, "#") - 1))
                        VGStrBanco = Trim(Mid(VGStrBanco, InStr(VGStrBanco, "#")))
                    Else
                        RecCredParc("NumBanco") = 0
                    End If
                    
                    If VGStrChequeDigito <> "" And VGStrChequeDigito <> "#" Then
                        VGStrChequeDigito = Trim(Mid(VGStrChequeDigito, InStr(VGStrChequeDigito, "#") + 1))
                        RecCredParc("NumCheque") = Trim(Mid(VGStrChequeDigito, 1, InStr(VGStrChequeDigito, "#") - 1))
                        VGStrChequeDigito = Trim(Mid(VGStrChequeDigito, InStr(VGStrChequeDigito, "#")))
                    Else
                        RecCredParc("NumCheque") = ""
                    End If
                    
                    RecCredParc.Update
                         
                    RecCredParc.Close
                                    
                    parcelatemp = parcelatemp + 1
                Loop
                
            End If
            
            '============== INCLUIR VENDA =============================
            StrSql = "SELECT * FROM tb_Venda"
            RecVenda.Open StrSql, vgCon, 1, 3
                
            RecVenda.AddNew
            If CboVendedor.Text = "" Then
                RecVenda("CodVendedor") = 0
            Else
                RecVenda("CodVendedor") = Trim(Mid(CboVendedor.Text, Len(CboVendedor.Text) - 10))
            End If
            RecVenda("CodCli") = VGIntCodCli
            RecVenda("CodCred") = RecCred!CodCred
            RecVenda("DtVenda") = FormataDataUS(Date)
            RecVenda("TipoVenda") = "A prazo - Cheque"
            RecVenda("SubTotalVenda") = CCur(TxtVendaChq.Text)
            RecVenda("Desconto") = ""
            RecVenda("TotalVenda") = CCur(TxtTotalVendaChq.Text)
            RecVenda("TipoPagto") = "Cheque"
            RecVenda("NumBanco") = 0
            RecVenda("NumCheque") = ""
            RecVenda.Update
                
            RecVenda.Close
            RecCred.Close
            
            StrSql = "SELECT MAX(CodVenda) as CodVenda FROM tb_Venda"
            RecVenda.Open StrSql, vgCon, 1, 3
            
            VLIntCodVendaTemp = RecVenda!codvenda
            VGIntCodVendaRel = RecVenda!codvenda
            
            RecVenda.Close
            
            '============== INCLUIR VENDA-PRODUTO =============================
            StrSql = "SELECT * FROM tb_Venda_Produto"
            RecVendaProd.Open StrSql, vgCon, 1, 3
            
            VLIntLinha = 1
            
            Do While VLIntLinha <= GridProduto.MaxRows
                GridProduto.Row = VLIntLinha
                
                RecVendaProd.AddNew
                RecVendaProd("CodVenda") = VLIntCodVendaTemp
                
                GridProduto.Col = 5
                RecVendaProd("CodProd") = GridProduto.Text
                
                GridProduto.Col = 3
                RecVendaProd("Qtde") = GridProduto.Text
                
                GridProduto.Col = 4
                RecVendaProd("ValorProd") = CCur(GridProduto.Text)
                RecVendaProd.Update
                
                VLIntLinha = VLIntLinha + 1
            Loop
            
            RecVendaProd.Close
           
            '================ RETIRAR QTDE DO ESTOQUE ========================
            VLIntLinha = 1
            
            Do While VLIntLinha <= GridProduto.MaxRows
                GridProduto.Row = VLIntLinha
                
                GridProduto.Col = 5
                StrSql = "SELECT QtdeProd FROM tb_Estoque where CodProd=" & GridProduto.Text
                RecEst.Open StrSql, vgCon, 1, 3
                
                If Not RecEst.EOF Then
                    GridProduto.Col = 3
                    RecEst("QtdeProd") = Int(RecEst!qtdeprod) - Int(GridProduto.Text)
                    RecEst.Update
                End If
                
                RecEst.Close
                
                VLIntLinha = VLIntLinha + 1
            Loop
                
            '================= INCLUIR NO MOVIMENTO DE CAIXA =========================
            
            If OptChqEntrDin.Value = True Or OptChqEntrChq.Value = True Then
                StrSql = "SELECT * FROM tb_Caixa"
                RecCx.Open StrSql, vgCon, 1, 3
                
                RecCx.AddNew
                RecCx("CodVenda") = VLIntCodVendaTemp
                RecCx("DtMov") = FormataDataUS(Date)
                RecCx("TipoMov") = "Entrada de venda"
                
                If OptChqEntrDin.Value = True Then
                    RecCx("Valor") = CCur(TxtValorEntrDinChq.Text)
                    
                ElseIf OptChqEntrChq.Value = True Then
                    RecCx("Valor") = CCur(TxtValorEntrChequeChq.Text)
                End If
                
                RecCx("TipoValor") = "credito"
                RecCx("Descricao") = "Entrada de venda a prazo em cheque  - Cliente: " & VGStrNomeCli
                
                If OptChqEntrDin.Value = True Then
                    RecCx("TipoPagto") = "Dinheiro"
                    
                ElseIf OptChqEntrChq.Value = True Then
                    RecCx("TipoPagto") = "Cheque"
                End If
                
                RecCx.Update
            End If
            
            Desconecta
            
            VPStrResponse = MsgBox("Venda efetuada." & Chr(13) & Chr(13) & "Deseja imprimir a proposta de crédito agora?", vbYesNo, "Pró Ótica 2004 - Informação")
            
            Unload Me
            
            MDIPrincipal.Enabled = True
            MDIPrincipal.WindowState = 2
            
            Screen.MousePointer = vbNormal
            
            If VPStrResponse = vbYes Then
                Call MontaImpressaoProposta
            End If
            
        ElseIf VPStrVenda = "prazocarne" Then
            Dim VLStrData As String
            Dim VLStrValor As String
            
            '============== INCLUIR CREDIÁRIO =============================
            StrSql = "SELECT * FROM tb_Crediario"
            RecCred.Open StrSql, vgCon, 1, 3
            
            RecCred.AddNew
            RecCred("CodCredsta") = VGIntCodCredstaVenda
            RecCred("CodCli") = VGIntCodCli
            RecCred("DtCred") = FormataDataUS(Date)
            RecCred("TipoCred") = "Carnê"
            RecCred("ValorVenda") = CCur(TxtVendaCar.Text)
            RecCred("Parcela") = CboPrazoCarParc.Text
            RecCred("Juros") = TxtPrazoCarJuros.Text
            RecCred("ValorTotal") = CCur(TxtTotalVendaCar.Text)
            
            If OptCarSemEntr.Value = True Then
                RecCred("TipoEntr") = "Sem entrada"
                RecCred("ValorEntr") = ""
                RecCred("NumBanco") = 0
                RecCred("NumCheque") = ""
            
            ElseIf OptCarEntrDin.Value = True Then
                RecCred("TipoEntr") = "Dinheiro"
                RecCred("ValorEntr") = CCur(TxtValorEntrDinCar.Text)
                RecCred("NumBanco") = 0
                RecCred("NumCheque") = ""
            
            ElseIf OptCarEntrChq.Value = True Then
                RecCred("TipoEntr") = "Cheque"
                RecCred("ValorEntr") = CCur(TxtValorEntrChequeCar.Text)
                
                If TxtBancoCar.Text <> "" Then
                    RecCred("NumBanco") = TxtBancoCar.Text
                Else
                    RecCred("NumBanco") = 0
                End If
                
                If TxtChequeCar.Text <> "" And TxtDigCar.Text <> "" Then
                    RecCred("NumCheque") = TxtChequeCar.Text & "-" & TxtDigCar.Text
                ElseIf TxtChequeCar.Text <> "" And TxtDigCar.Text = "" Then
                    RecCred("NumCheque") = TxtChequeCar.Text
                ElseIf TxtChequeCar.Text = "" And TxtDigCar.Text = "" Then
                    RecCred("NumCheque") = ""
                End If
                
            End If
            
            RecCred.Update
            
            RecCred.Close
            
            StrSql = "SELECT MAX(CodCred) as CodCred FROM tb_Crediario where CodCli=" & VGIntCodCli
            RecCred.Open StrSql, vgCon, 1, 3
            
            VPIntCodCredTemp = RecCred!CodCred
            
            '============== INCLUIR PARCELAS DO CREDIÁRIO =============================
            If CboPrazoCarParc.Text <> "" Then
                
                VLStrData = Date
                VLStrValor = CCur(Mid(LblParcCar.Caption, InStr(LblParcCar.Caption, "R$")))
                
                Do While parcelatemp <= Val(CboPrazoCarParc.Text)
                
                    StrSql = "SELECT * FROM tb_Crediario_Parcela"
                    RecCredParc.Open StrSql, vgCon, 1, 3
                     
                    RecCredParc.AddNew
                    RecCredParc("CodCred") = RecCred!CodCred
                    RecCredParc("NumParc") = parcelatemp
                    
                    VLStrData = DateSerial(Year(VLStrData), Month(VLStrData), Day(VLStrData) + 30)
                    
                    RecCredParc("Vencimento") = FormataDataUS(VLStrData)
                    RecCredParc("Valor") = VLStrValor
                    RecCredParc("Quitado") = "não"
                    RecCredParc("NumBanco") = 0
                    RecCredParc("NumCheque") = ""
                    RecCredParc.Update
                         
                    RecCredParc.Close
                                    
                    parcelatemp = parcelatemp + 1
                Loop
                
            End If
            
            '============== INCLUIR VENDA =============================
            StrSql = "SELECT * FROM tb_Venda"
            RecVenda.Open StrSql, vgCon, 1, 3
                
            RecVenda.AddNew
            If CboVendedor.Text = "" Then
                RecVenda("CodVendedor") = "0"
            Else
                RecVenda("CodVendedor") = Trim(Mid(CboVendedor.Text, Len(CboVendedor.Text) - 10))
            End If
            RecVenda("CodCli") = VGIntCodCli
            RecVenda("CodCred") = RecCred!CodCred
            RecVenda("DtVenda") = FormataDataUS(Date)
            RecVenda("TipoVenda") = "A prazo - Carnê"
            RecVenda("SubTotalVenda") = CCur(TxtVendaCar.Text)
            RecVenda("Desconto") = ""
            RecVenda("TotalVenda") = CCur(TxtTotalVendaCar.Text)
            RecVenda("TipoPagto") = "Carnê"
            RecVenda("NumBanco") = 0
            RecVenda("NumCheque") = ""
            RecVenda.Update
                
            RecVenda.Close
            RecCred.Close
            
            StrSql = "SELECT MAX(CodVenda) as CodVenda FROM tb_Venda"
            RecVenda.Open StrSql, vgCon, 1, 3
            
            VLIntCodVendaTemp = RecVenda!codvenda
            VGIntCodVendaRel = RecVenda!codvenda
            
            RecVenda.Close
            
            '============== INCLUIR VENDA-PRODUTO =============================
            StrSql = "SELECT * FROM tb_Venda_Produto"
            RecVendaProd.Open StrSql, vgCon, 1, 3
            
            VLIntLinha = 1
            
            Do While VLIntLinha <= GridProduto.MaxRows
                GridProduto.Row = VLIntLinha
                
                RecVendaProd.AddNew
                RecVendaProd("CodVenda") = VLIntCodVendaTemp
                
                GridProduto.Col = 5
                RecVendaProd("CodProd") = GridProduto.Text
                
                GridProduto.Col = 3
                RecVendaProd("Qtde") = GridProduto.Text
                
                GridProduto.Col = 4
                RecVendaProd("ValorProd") = CCur(GridProduto.Text)
                RecVendaProd.Update
                
                VLIntLinha = VLIntLinha + 1
            Loop
            
            RecVendaProd.Close
            
            '================ RETIRAR QTDE DO ESTOQUE ========================
            VLIntLinha = 1
            
            Do While VLIntLinha <= GridProduto.MaxRows
                GridProduto.Row = VLIntLinha
                
                GridProduto.Col = 5
                StrSql = "SELECT QtdeProd FROM tb_Estoque where CodProd=" & GridProduto.Text
                RecEst.Open StrSql, vgCon, 1, 3
                
                If Not RecEst.EOF Then
                    GridProduto.Col = 3
                    RecEst("QtdeProd") = Int(RecEst!qtdeprod) - Int(GridProduto.Text)
                    RecEst.Update
                End If
                
                RecEst.Close
                
                VLIntLinha = VLIntLinha + 1
            Loop
                
            '================= INCLUIR NO MOVIMENTO DE CAIXA =========================
            If OptCarEntrDin.Value = True Or OptCarEntrChq.Value = True Then
                StrSql = "SELECT * FROM tb_Caixa"
                RecCx.Open StrSql, vgCon, 1, 3
                
                RecCx.AddNew
                RecCx("CodVenda") = VLIntCodVendaTemp
                RecCx("DtMov") = FormataDataUS(Date)
                RecCx("TipoMov") = "Entrada de venda"
                
                If OptCarEntrDin.Value = True Then
                    RecCx("Valor") = CCur(TxtValorEntrDinCar.Text)
                    
                ElseIf OptCarEntrChq.Value = True Then
                    RecCx("Valor") = CCur(TxtValorEntrChequeCar.Text)
                End If
                
                RecCx("TipoValor") = "credito"
                RecCx("Descricao") = "Entrada de venda a prazo em carnê  - Cliente: " & VGStrNomeCli
                
                If OptCarEntrDin.Value = True Then
                    RecCx("TipoPagto") = "Dinheiro"
                    
                ElseIf OptCarEntrChq.Value = True Then
                    RecCx("TipoPagto") = "Cheque"
                End If
                
                RecCx.Update
            End If
            
            Desconecta
            
            VPStrResponse = MsgBox("Venda efetuada." & Chr(13) & Chr(13) & "Deseja imprimir a proposta de crédito agora?", vbYesNo, "Pró Ótica 2004 - Informação")
            
            Unload Me
            
            MDIPrincipal.Enabled = True
            MDIPrincipal.WindowState = 2
            
            Screen.MousePointer = vbNormal
            
            If VPStrResponse = vbYes Then
                Call MontaImpressaoProposta
            End If
        
        End If
        
    End If
   
End Sub

Private Sub CmdParcInc_Click()
    If CboPrazoChqParc.Text = "" Then
        VPStrBox = MsgBox("Você deve selecionar a quantidade de parcelas.", vbInformation, "Pró Ótica 2004 - Informação")
    Else
        FrmParcela_Inc.Show
    End If
End Sub

Private Sub CmdVerProd_Click()
    FrmVenda_Inc_Prod.Show
End Sub

Private Sub Form_Resize()
  FrmVenda_Inc.Left = (MDIPrincipal.Width / 2) - (FrmVenda_Inc.Width / 2)
  FrmVenda_Inc.Top = (MDIPrincipal.Height / 3) - (FrmVenda_Inc.Height / 3)
End Sub

Private Sub Form_Load()
    Height = 9375
    Width = 8040
    
    Skin1.LoadSkin (App.Path & "\ProVendas2004.skn")
    Skin1.ApplySkin (Me.hwnd)
    
    MDIPrincipal.Enabled = False
    CmdExcluirProd.Enabled = False
    
    Call MontaCboVendedor
    Call MontaParcelas
    
    If VGStrVendaRapida = "sim" Then
        VGStrVendaRapida = ""
        OptPrazo.Enabled = False
    Else
        OptPrazo.Enabled = True
    End If
    
End Sub

Private Sub GridProduto_Click(ByVal Col As Long, ByVal Row As Long)
    GridProduto.Row = Row
    GridProduto.Col = 5
    If GridProduto.Text <> "" And GridProduto.Text <> "CodProd" Then
        CmdExcluirProd.Enabled = True
    Else
        CmdExcluirProd.Enabled = False
    End If
End Sub

Private Sub GridProduto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dim VLStrPreco As String
        Dim VLIntQtde As Long
    
        GridProduto.Col = 3
        GridProduto.Row = GridProduto.ActiveRow
        
        If GridProduto.Text <> "" Then
            VLIntQtde = GridProduto.Text
            
            GridProduto.Col = 2
            GridProduto.Row = GridProduto.ActiveRow
            VLStrPreco = Trim(Mid(GridProduto.Text, 3))
        
            GridProduto.Col = 4
            GridProduto.Row = GridProduto.ActiveRow
            GridProduto.Text = FormataMoeda(VLStrPreco * VLIntQtde)
        End If
    End If
End Sub

Private Sub GridProduto_LeaveCell(ByVal Col As Long, ByVal Row As Long, ByVal NewCol As Long, ByVal NewRow As Long, Cancel As Boolean)
    Dim VLStrPreco As String
    Dim VLIntQtde As Long

    GridProduto.Col = 3
    GridProduto.Row = GridProduto.ActiveRow
    
    If GridProduto.Text <> "" Then
        VLIntQtde = GridProduto.Text
        
        GridProduto.Col = 2
        GridProduto.Row = GridProduto.ActiveRow
        VLStrPreco = Trim(Mid(GridProduto.Text, 3))
    
        GridProduto.Col = 4
        GridProduto.Row = GridProduto.ActiveRow
        GridProduto.Text = FormataMoeda(VLStrPreco * VLIntQtde)
    End If
End Sub

Private Sub GridProduto_LeaveRow(ByVal Row As Long, ByVal RowWasLast As Boolean, ByVal RowChanged As Boolean, ByVal AllCellsHaveData As Boolean, ByVal NewRow As Long, ByVal NewRowIsLast As Long, Cancel As Boolean)
    Dim VLStrPreco As String
    Dim VLIntQtde As Long

    GridProduto.Col = 3
    GridProduto.Row = GridProduto.ActiveRow
    
    If GridProduto.Text <> "" Then
        VLIntQtde = GridProduto.Text
        
        GridProduto.Col = 2
        GridProduto.Row = GridProduto.ActiveRow
        VLStrPreco = Trim(Mid(GridProduto.Text, 3))
    
        GridProduto.Col = 4
        GridProduto.Row = GridProduto.ActiveRow
        GridProduto.Text = FormataMoeda(VLStrPreco * VLIntQtde)
    End If
End Sub

Private Sub OptCarEntrChq_Click()
    FraCarEntrChq.Visible = True
    FraCarEntrDin.Visible = False
    
    Dim restparc As String
    
    If TxtTotalVendaCar.Text = "" Then
        TxtTotalVendaCar.Text = TxtVendaCar.Text
    End If
    
    If TxtTotalVendaCar.Text <> "" Then
        restparc = CCur(TxtTotalVendaCar.Text) - ((CCur(TxtTotalVendaCar.Text) * 20) / 100)
        
        LblEntrCar.Caption = "Entrada: " & FormataMoeda(CCur((CCur(TxtTotalVendaCar.Text) * 20) / 100))
        
        If restparc = "0" Then
            LblParcCar.Caption = "00 parcela(s) de R$ 0,00"
        Else
            LblParcCar.Caption = CboPrazoCarParc.Text & " parcela(s) de " & FormataMoeda(restparc / CboPrazoCarParc.Text)
        End If
        
        TxtValorEntrChequeCar.Text = Mid(LblEntrCar.Caption, InStr(LblEntrCar.Caption, "R$"))
    End If
End Sub

Private Sub OptCarEntrDin_Click()
    FraCarEntrChq.Visible = False
    FraCarEntrDin.Visible = True
    
    Dim restparc As String
    
    If TxtTotalVendaCar.Text = "" Then
        TxtTotalVendaCar.Text = TxtVendaCar.Text
    End If
    
    If TxtTotalVendaCar.Text <> "" And CboPrazoCarParc.Text <> "" Then
        restparc = CCur(TxtTotalVendaCar.Text) - ((CCur(TxtTotalVendaCar.Text) * 20) / 100)
        
        LblEntrCar.Caption = "Entrada: " & FormataMoeda((CCur(TxtTotalVendaCar.Text) * 20) / 100)
        
        If restparc = "0" Then
            LblParcCar.Caption = "00 parcela(s) de R$ 0,00"
        Else
            LblParcCar.Caption = CboPrazoCarParc.Text & " parcela(s) de " & FormataMoeda(restparc / CboPrazoCarParc.Text)
        End If
        
        TxtValorEntrDinCar.Text = Mid(LblEntrCar.Caption, InStr(LblEntrCar.Caption, "R$"))
    End If
End Sub

Private Sub OptCarSemEntr_Click()
    FraCarEntrDin.Visible = False
    FraCarEntrChq.Visible = False

    If TxtTotalVendaCar.Text <> "" Then
        LblEntrCar.Caption = "Entrada: R$ 0,00"
        
        If TxtTotalVendaCar.Text = "R$ 0,00" Then
            LblParcCar.Caption = "00 parcela(s) de R$ 0,00"
        Else
            LblParcCar.Caption = CboPrazoCarParc.Text & " parcela(s) de " & FormataMoeda(TxtTotalVendaCar.Text / CboPrazoCarParc.Text)
        End If
    End If
End Sub

Private Sub OptChqEntrChq_Click()
    FraChqEntrChq.Visible = True
    FraChqEntrDin.Visible = False
    
    Dim restparc As String
    
    If TxtTotalVendaChq.Text = "" Then
        TxtTotalVendaChq.Text = TxtVendaChq.Text
    End If
    
    If TxtTotalVendaChq.Text <> "" Then
        restparc = CCur(TxtTotalVendaChq.Text) - ((CCur(TxtTotalVendaChq.Text) * 20) / 100)
        
        LblEntrChq.Caption = "Entrada: " & FormataMoeda(CCur((CCur(TxtTotalVendaChq.Text) * 20) / 100))
        
        If restparc = "0" Then
            LblParcChq.Caption = "00 parcela(s) de R$ 0,00"
        Else
            LblParcChq.Caption = CboPrazoChqParc.Text & " parcela(s) de " & FormataMoeda(restparc / CboPrazoChqParc.Text)
        End If
        
        TxtValorEntrChequeChq.Text = Mid(LblEntrChq.Caption, InStr(LblEntrChq.Caption, "R$"))
    End If
End Sub

Private Sub OptChqEntrDin_Click()
    FraChqEntrChq.Visible = False
    FraChqEntrDin.Visible = True
    
    Dim restparc As String
    
    If TxtTotalVendaChq.Text = "" Then
        TxtTotalVendaChq.Text = TxtVendaChq.Text
    End If
    
    If TxtTotalVendaChq.Text <> "" And CboPrazoChqParc.Text <> "" Then
        restparc = CCur(TxtTotalVendaChq.Text) - ((CCur(TxtTotalVendaChq.Text) * 20) / 100)
        
        LblEntrChq.Caption = "Entrada: " & FormataMoeda((CCur(TxtTotalVendaChq.Text) * 20) / 100)
        
        If restparc = "0" Then
            LblParcChq.Caption = "00 parcela(s) de R$ 0,00"
        Else
            LblParcChq.Caption = CboPrazoChqParc.Text & " parcela(s) de " & FormataMoeda(restparc / CboPrazoChqParc.Text)
        End If
        
        TxtValorEntrDinChq.Text = Mid(LblEntrChq.Caption, InStr(LblEntrChq.Caption, "R$"))
    End If
End Sub

Private Sub OptChqSemEntr_Click()
    FraChqEntrChq.Visible = False
    FraChqEntrDin.Visible = False
    
    If TxtTotalVendaChq.Text <> "" Then
        LblEntrChq.Caption = "Entrada: R$ 0,00"
        
        If TxtTotalVendaChq.Text = "R$ 0,00" Then
            LblParcChq.Caption = "00 parcela(s) de R$ 0,00"
        Else
            LblParcChq.Caption = CboPrazoChqParc.Text & " parcela(s) de " & FormataMoeda(TxtTotalVendaChq.Text / CboPrazoChqParc.Text)
        End If
    End If
End Sub

Private Sub OptCredCarne_Click()
    FraCredCarne.Visible = True
    FraCredCheque.Visible = False
End Sub

Private Sub OptCredCheque_Click()
    FraCredCarne.Visible = False
    FraCredCheque.Visible = True
End Sub

Private Sub OptPrazo_Click()
    Dim VLIntLinha As Long
    Dim VLCurTotal As Currency
    
    Call VerificaGridProd
    
    VLIntLinha = GridProduto.MaxRows
    
    GridProduto.Col = 4
    
    Do While VLIntLinha <> 0
        GridProduto.Row = VLIntLinha
        VLCurTotal = VLCurTotal + Trim(Mid(GridProduto.Text, 3))
    
        VLIntLinha = VLIntLinha - 1
    Loop

    OptPrazoCheque.Value = False
    OptPrazoCarne.Value = False
    
    FraVista.Visible = False
    FraPrazoCarne.Visible = False
    FraPrazoCheque.Visible = False
    FraTipoPrazo.Visible = True
    
    VPStrValorTotal = VLCurTotal
    
End Sub

Private Sub OptPrazoCarne_Click()
    VPStrVenda = "prazocarne"
    
    FraVista.Visible = False
    FraPrazoCarne.Visible = True
    FraPrazoCheque.Visible = False
    FraTipoPrazo.Visible = True
    
    TxtVendaCar.Text = FormataMoeda(VPStrValorTotal)
    
    If TxtPrazoCarJuros.Text <> "" And TxtVendaCar.Text <> "" Then
        TxtTotalVendaCar.Text = FormataMoeda(CCur(TxtVendaCar.Text) + (CCur(TxtVendaCar.Text) * TxtPrazoCarJuros.Text) / 100)
    ElseIf TxtPrazoCarJuros.Text = "" And TxtVendaCar.Text <> "" Then
        TxtTotalVendaCar.Text = TxtVendaCar.Text
    End If
    
    OptCarEntrDin_Click
    
    TxtPrazoCarJuros.SetFocus
    OptCarEntrDin.Value = True
End Sub

Private Sub OptPrazoCheque_Click()
    VPStrVenda = "prazocheque"
    
    FraVista.Visible = False
    FraPrazoCarne.Visible = False
    FraPrazoCheque.Visible = True
    FraTipoPrazo.Visible = True
    
    TxtVendaChq.Text = FormataMoeda(VPStrValorTotal)
    
    If TxtPrazoChqJuros.Text <> "" And TxtVendaChq.Text <> "" Then
        TxtTotalVendaChq.Text = FormataMoeda(CCur(TxtVendaChq.Text) + (CCur(TxtVendaChq.Text) * TxtPrazoChqJuros.Text) / 100)
    ElseIf TxtPrazoChqJuros.Text = "" And TxtVendaChq.Text <> "" Then
        TxtTotalVendaChq.Text = FormataMoeda(TxtVendaChq.Text)
    End If
    
    OptChqEntrDin_Click
    
    TxtPrazoChqJuros.SetFocus
    OptChqEntrDin.Value = True
End Sub

Private Sub OptVista_Click()
    Dim VLIntLinha As Long
    Dim VLCurTotal As Currency
    
    VPStrVenda = "vista"
    
    Call VerificaGridProd
    
    VLIntLinha = GridProduto.MaxRows
    
    GridProduto.Col = 4
    
    Do While VLIntLinha <> 0
        GridProduto.Row = VLIntLinha
        VLCurTotal = VLCurTotal + Trim(Mid(GridProduto.Text, 3))
    
        VLIntLinha = VLIntLinha - 1
    Loop
    
    FraVista.Visible = True
    FraPrazoCarne.Visible = False
    FraPrazoCheque.Visible = False
    FraTipoPrazo.Visible = False
    
    VPStrValorTotal = VLCurTotal
    
    TxtVendaVista.Text = FormataMoeda(VPStrValorTotal)
    TxtTotalVista.Text = FormataMoeda(VPStrValorTotal)
    TxtDescVista.SetFocus
    OptDin.Value = True
    
    Call TxtVendaVista_LostFocus
End Sub

Private Sub TxtBancoCar_GotFocus()
    TxtBancoCar.SelStart = 0
    TxtBancoCar.SelLength = Len(TxtBancoCar.Text)
End Sub

Private Sub TxtBancoChq_GotFocus()
    TxtBancoChq.SelStart = 0
    TxtBancoChq.SelLength = Len(TxtBancoChq.Text)
End Sub

Private Sub TxtBancoVista_GotFocus()
    TxtBancoVista.SelStart = 0
    TxtBancoVista.SelLength = Len(TxtBancoVista.Text)
End Sub

Private Sub TxtChequeCar_GotFocus()
    TxtChequeCar.SelStart = 0
    TxtChequeCar.SelLength = Len(TxtChequeCar.Text)
End Sub

Private Sub TxtChequeChq_GotFocus()
    TxtChequeChq.SelStart = 0
    TxtChequeChq.SelLength = Len(TxtChequeChq.Text)
End Sub

Private Sub TxtChequeVista_GotFocus()
    TxtChequeVista.SelStart = 0
    TxtChequeVista.SelLength = Len(TxtChequeVista.Text)
End Sub

Private Sub TxtDescVista_GotFocus()
    TxtDescVista.SelStart = 0
    TxtDescVista.SelLength = Len(TxtDescVista.Text)
End Sub

Private Sub TxtDescVista_KeyPress(KeyAscii As Integer)
    '=== Só aceita números ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtDescVista_LostFocus()
    If TxtDescVista.Text <> "" And TxtVendaVista.Text <> "" Then
        TxtTotalVista.Text = FormataMoeda(CCur(TxtVendaVista.Text) - (CCur(TxtVendaVista.Text) * TxtDescVista.Text) / 100)
    
    ElseIf TxtDescVista.Text = "" And TxtVendaVista.Text <> "" Then
        TxtTotalVista.Text = FormataMoeda(TxtVendaVista.Text)
    
    End If
End Sub

Private Sub TxtDigCar_GotFocus()
    TxtDigCar.SelStart = 0
    TxtDigCar.SelLength = Len(TxtDigCar.Text)
End Sub

Private Sub TxtDigChq_GotFocus()
    TxtDigChq.SelStart = 0
    TxtDigChq.SelLength = Len(TxtDigChq.Text)
End Sub

Private Sub TxtDigVista_GotFocus()
    TxtDigVista.SelStart = 0
    TxtDigVista.SelLength = Len(TxtDigVista.Text)
End Sub

Private Sub TxtPrazoCarJuros_GotFocus()
    TxtPrazoCarJuros.SelStart = 0
    TxtPrazoCarJuros.SelLength = Len(TxtPrazoCarJuros.Text)
End Sub

Private Sub TxtPrazoCarJuros_KeyPress(KeyAscii As Integer)
    '=== Só aceita números ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtPrazoCarJuros_LostFocus()
    If TxtPrazoCarJuros.Text <> "" And TxtVendaCar.Text <> "" Then
        TxtTotalVendaCar.Text = FormataMoeda(CCur(TxtVendaCar.Text) + (CCur(TxtVendaCar.Text) * TxtPrazoCarJuros.Text) / 100)
    
    ElseIf TxtPrazoCarJuros.Text = "" And TxtVendaCar.Text <> "" Then
        TxtTotalVendaCar.Text = FormataMoeda(TxtVendaCar.Text)
    
    End If
    
    Call OptCarEntrDin_Click
End Sub

Private Sub TxtPrazoChqJuros_GotFocus()
    TxtPrazoChqJuros.SelStart = 0
    TxtPrazoChqJuros.SelLength = Len(TxtPrazoChqJuros.Text)
End Sub

Private Sub TxtPrazoChqJuros_KeyPress(KeyAscii As Integer)
    '=== Só aceita números ===
    If (KeyAscii < 48 Or KeyAscii > 57) And KeyAscii <> 8 And KeyAscii <> 13 Then
        KeyAscii = 0
    End If
End Sub

Private Sub TxtPrazoChqJuros_LostFocus()
    If TxtPrazoChqJuros.Text <> "" And TxtVendaChq.Text <> "" Then
        TxtTotalVendaChq.Text = FormataMoeda(CCur(TxtVendaChq.Text) + (CCur(TxtVendaChq.Text) * TxtPrazoChqJuros.Text) / 100)
    
    ElseIf TxtPrazoChqJuros.Text = "" And TxtVendaChq.Text <> "" Then
        TxtTotalVendaChq.Text = FormataMoeda(TxtVendaChq.Text)
    
    End If
    
    Call OptChqEntrDin_Click
End Sub

Sub MontaCboVendedor()
    Conecta
    
    Dim RecVend As New ADODB.Recordset
    
    StrSql = "SELECT CodVendedor,Nome FROM tb_Vendedor order by Nome"
    RecVend.Open StrSql, vgCon, 1, 3
    
    CboVendedor.AddItem ("                                                                                                                 0")
    Do While Not RecVend.EOF
        CboVendedor.AddItem (RecVend!nome & "                                                                                                      " & RecVend!CodVendedor)
        RecVend.MoveNext
    Loop
    
    Desconecta
    
End Sub

Private Sub TxtTotalVendaCar_GotFocus()
    TxtTotalVendaCar.SelStart = 0
    TxtTotalVendaCar.SelLength = Len(TxtTotalVendaCar.Text)
End Sub

Private Sub TxtTotalVendaChq_GotFocus()
    TxtTotalVendaChq.SelStart = 0
    TxtTotalVendaChq.SelLength = Len(TxtTotalVendaChq.Text)
End Sub

Private Sub TxtTotalVista_GotFocus()
    TxtTotalVista.SelStart = 0
    TxtTotalVista.SelLength = Len(TxtTotalVista.Text)
End Sub

Private Sub TxtValorEntrChequeCar_GotFocus()
    TxtValorEntrChequeCar.SelStart = 0
    TxtValorEntrChequeCar.SelLength = Len(TxtValorEntrChequeCar.Text)
End Sub

Private Sub TxtValorEntrChequeChq_GotFocus()
    TxtValorEntrChequeChq.SelStart = 0
    TxtValorEntrChequeChq.SelLength = Len(TxtValorEntrChequeChq.Text)
End Sub

Private Sub TxtValorEntrDinCar_GotFocus()
    TxtValorEntrDinCar.SelStart = 0
    TxtValorEntrDinCar.SelLength = Len(TxtValorEntrDinCar.Text)
End Sub

'Private Sub TxtTotalVista_GotFocus()
'    If TxtVendaVista.Text <> "" And TxtDescVista.Text <> "" Then
'        TxtTotalVista.Text = FormataMoeda(CCur(TxtVendaVista.Text) - ((CCur(TxtVendaVista.Text) * TxtDescVista.Text) / 100))
'
'    ElseIf TxtVendaVista.Text <> "" And TxtDescVista.Text = "" Then
'        TxtTotalVista.Text = FormataMoeda(TxtVendaVista.Text)
'
'    End If
'End Sub

Private Sub TxtValorEntrDinCar_LostFocus()
    If TxtValorEntrDinCar.Text <> "" Then
        TxtValorEntrDinCar.Text = FormataMoeda(TxtValorEntrDinCar.Text)
    End If
End Sub

Private Sub TxtValorEntrChequeCar_LostFocus()
    If TxtValorEntrChequeCar.Text <> "" Then
        TxtValorEntrChequeCar.Text = FormataMoeda(TxtValorEntrChequeCar.Text)
    End If
End Sub

Private Sub TxtValorEntrDinChq_GotFocus()
    TxtValorEntrDinChq.SelStart = 0
    TxtValorEntrDinChq.SelLength = Len(TxtValorEntrDinChq.Text)
End Sub

Private Sub TxtValorEntrDinChq_LostFocus()
    If TxtValorEntrDinChq.Text <> "" Then
        TxtValorEntrDinChq.Text = FormataMoeda(TxtValorEntrDinChq.Text)
    End If
End Sub

Private Sub TxtValorEntrChequeChq_LostFocus()
    If TxtValorEntrChequeChq.Text <> "" Then
        TxtValorEntrChequeChq.Text = FormataMoeda(TxtValorEntrChequeChq.Text)
    End If
End Sub

Private Sub TxtVendaCar_GotFocus()
    TxtVendaCar.SelStart = 0
    TxtVendaCar.SelLength = Len(TxtVendaCar.Text)
End Sub

Private Sub TxtVendaCar_LostFocus()
    If TxtPrazoCarJuros.Text = "" Then
        TxtTotalVendaCar.Text = FormataMoeda(TxtVendaCar.Text)
    Else
        TxtTotalVendaCar.Text = FormataMoeda(CCur(TxtVendaCar.Text) + (CCur(TxtVendaCar.Text) * TxtPrazoCarJuros.Text) / 100)
    End If
End Sub

Private Sub TxtVendaChq_GotFocus()
    TxtVendaChq.SelStart = 0
    TxtVendaChq.SelLength = Len(TxtVendaChq.Text)
End Sub

Private Sub TxtVendaChq_LostFocus()
    If TxtPrazoChqJuros.Text = "" Then
        TxtTotalVendaChq.Text = FormataMoeda(TxtVendaChq.Text)
    Else
        TxtTotalVendaChq.Text = FormataMoeda(CCur(TxtVendaChq.Text) + (CCur(TxtVendaChq.Text) * TxtPrazoChqJuros.Text) / 100)
    End If
End Sub

Sub MontaImpressaoProposta()
    Screen.MousePointer = vbHourglass
    
    Dim RecCred As New ADODB.Recordset
    Dim RecCredParc As New ADODB.Recordset
    Dim RecCli As New ADODB.Recordset
    Dim RecCredsta As New ADODB.Recordset
    Dim RecMed As New ADODB.Recordset
    Dim RecRec As New ADODB.Recordset
    Dim RecAux As New ADODB.Recordset
    Dim VLStrNomeMed As String
    Dim VLStrCRMMed As String
    Dim VLStrCPFMed As String
    Dim parctemp As Integer
    Dim cont As Integer
    Dim campo As String
    
    Conecta
    
    If VGStrProposta = "imprimir" Then
        VGStrProposta = ""
        VPIntCodCredTemp = VGIntPropCodCred
    End If
    
    '=== Pega informações do crediário =======
    StrSql = "Select CodCredsta,CodCli,DtCred,TipoCred,ValorVenda,Parcela,Juros,ValorTotal,TipoEntr,ValorEntr " & _
             "From tb_Crediario where CodCred=" & VPIntCodCredTemp
    RecCred.Open StrSql, vgCon, 1, 3
    
    '=== Pega informações das parcelas crediário =======
    StrSql = "Select Vencimento,Valor From tb_Crediario_Parcela where CodCred=" & VPIntCodCredTemp
    RecCredParc.Open StrSql, vgCon, 1, 3
    
    '=== Pega informações do crediarista =======
    StrSql = "Select Nome,Endereco,Bairro,Cep,Cidade,Estado,DtNasc,Telefone1,Telefone2,CPF " & _
             "From tb_Crediarista where CodCredsta=" & RecCred!CodCredsta
    RecCredsta.Open StrSql, vgCon, 1, 3
    
    '=== Pega informações do cliente =======
    StrSql = "Select Nome,Endereco,Bairro,Cep,Cidade,Estado,DtNasc,Telefone1,Telefone2,CPF " & _
             "From tb_Cliente where CodCli=" & RecCred!CodCli
    RecCli.Open StrSql, vgCon, 1, 3
    
    '=== Insere informações na tabela auxiliar =======
    StrSql = "Select * From tb_Auxiliar"
    RecAux.Open StrSql, vgCon, 1, 3
    
    RecAux.AddNew
    RecAux("Campo01") = RecCredsta!nome
    RecAux("Campo02") = FormataData(RecCredsta!dtnasc)
    RecAux("Campo03") = RecCredsta!cpf
    RecAux("Campo04") = RecCredsta!telefone1 & "/" & RecCredsta!telefone2
    RecAux("Campo05") = RecCredsta!endereco
    RecAux("Campo06") = RecCredsta!bairro
    RecAux("Campo07") = RecCredsta!cidade & "/" & RecCredsta!Estado
    RecAux("Campo08") = RecCredsta!cep
    RecAux("Campo09") = RecCli!nome
    RecAux("Campo10") = FormataData(RecCli!dtnasc)
    RecAux("Campo11") = RecCli!cpf
    
    If RecCli!telefone1 <> "" And RecCli!telefone2 <> "" Then
        RecAux("Campo12") = RecCli!telefone1 & "/" & RecCli!telefone2
    ElseIf RecCli!telefone1 <> "" And RecCli!telefone2 = "" Then
        RecAux("Campo12") = RecCli!telefone1
    ElseIf RecCli!telefone1 = "" And RecCli!telefone2 <> "" Then
        RecAux("Campo12") = RecCli!telefone2
    ElseIf RecCli!telefone1 = "" And RecCli!telefone2 = "" Then
        RecAux("Campo12") = ""
    End If
    
    RecAux("Campo13") = RecCli!endereco
    RecAux("Campo14") = RecCli!bairro
    RecAux("Campo15") = RecCli!cidade & "/" & RecCli!Estado
    RecAux("Campo16") = RecCli!cep
    RecAux("Campo17") = FormataData(RecCred!dtcred)
    RecAux("Campo18") = RecCred!tipocred
    RecAux("Campo19") = FormataMoeda(RecCred!valorvenda)
    If RecCred!juros = "" Then
        RecAux("Campo20") = ""
    Else
        RecAux("Campo20") = FormataNum(RecCred!juros) & "%"
    End If
    RecAux("Campo21") = FormataMoeda(RecCred!valortotal)
    RecAux("Campo22") = FormataNum(RecCred!parcela)
    RecAux("Campo23") = RecCred!tipoentr
    If RecCred!valorentr = "" Then
        RecAux("Campo24") = ""
    Else
        RecAux("Campo24") = FormataMoeda(RecCred!valorentr)
    End If
    
    RecAux("Campo25") = FormataMoeda(RecCredParc!valor)
    RecAux.Update
    
    Desconecta
        
    rptPropCredito.Show
    
    Screen.MousePointer = vbNormal
End Sub

Sub MontaParcelas()
    Dim count As Integer
    
    '=== Parcela de cheque ===
    count = 1
    Do While count <= 24
        CboPrazoChqParc.AddItem (FormataNum(count))
        count = count + 1
    Loop
    
    CboPrazoChqParc.Text = "01"
    
    '=== Parcela de carnê ===
    count = 1
    Do While count <= 24
        CboPrazoCarParc.AddItem (FormataNum(count))
        count = count + 1
    Loop
    
    CboPrazoCarParc.Text = "01"
    
End Sub

Private Sub TxtVendaVista_GotFocus()
    TxtVendaVista.SelStart = 0
    TxtVendaVista.SelLength = Len(TxtVendaVista.Text)
End Sub

Private Sub TxtVendaVista_LostFocus()
    If TxtDescVista.Text = "" Then
        TxtTotalVista.Text = FormataMoeda(TxtVendaVista.Text)
    Else
        TxtTotalVista.Text = FormataMoeda(CCur(TxtVendaVista.Text) - (CCur(TxtVendaVista.Text) * TxtDescVista.Text) / 100)
    End If
End Sub

Sub VerificaGridProd()
    Dim VLStrPreco As String
    Dim VLIntQtde As Long
    
    VLIntLinha = 1
    
    Do While VLIntLinha <= GridProduto.MaxRows
        GridProduto.Row = VLIntLinha
        
        GridProduto.Col = 2
        VLStrPreco = Trim(Mid(GridProduto.Text, 3))
        
        GridProduto.Col = 3
        If GridProduto.Text = "" Then
            GridProduto.Text = 1
        End If
        VLIntQtde = GridProduto.Text
        
        GridProduto.Col = 4
        If GridProduto.Text = "" Then
            GridProduto.Text = FormataMoeda(VLStrPreco * VLIntQtde)
        End If
        
        VLIntLinha = VLIntLinha + 1
    Loop
End Sub

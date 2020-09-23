VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "2D Room Planner"
   ClientHeight    =   10830
   ClientLeft      =   150
   ClientTop       =   450
   ClientWidth     =   19110
   ClipControls    =   0   'False
   ForeColor       =   &H00000000&
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   722
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   1274
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame fraFXY 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Top-Left"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1230
      Left            =   12870
      TabIndex        =   129
      Top             =   75
      Width           =   1200
      Begin VB.Label LabItem 
         BackColor       =   &H00E0E0E0&
         Height          =   210
         Left            =   510
         TabIndex        =   135
         Top             =   195
         Width           =   255
      End
      Begin VB.Label Label8 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Item"
         Height          =   195
         Left            =   75
         TabIndex        =   134
         Top             =   210
         Width           =   450
      End
      Begin VB.Label LabFXY 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   1
         Left            =   465
         TabIndex        =   133
         Top             =   450
         Width           =   690
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Top"
         Height          =   255
         Index           =   1
         Left            =   90
         TabIndex        =   132
         Top             =   480
         Width           =   360
      End
      Begin VB.Label LabFXY 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Index           =   0
         Left            =   465
         TabIndex        =   131
         Top             =   840
         Width           =   690
      End
      Begin VB.Label Label7 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Left"
         Height          =   165
         Index           =   0
         Left            =   90
         TabIndex        =   130
         Top             =   885
         Width           =   330
      End
   End
   Begin VB.Frame fraConvert 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Converter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2250
      Left            =   14070
      TabIndex        =   113
      Top             =   60
      Width           =   1710
      Begin VB.CommandButton cmdClose 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Close"
         Height          =   210
         Left            =   510
         TabIndex        =   120
         Top             =   1890
         Width           =   645
      End
      Begin VB.CommandButton cmdcm 
         BackColor       =   &H00E0E0E0&
         Caption         =   "Convert to cm"
         Height          =   270
         Left            =   135
         TabIndex        =   118
         Top             =   1140
         Width           =   1455
      End
      Begin VB.TextBox txtftin 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   1
         Left            =   780
         TabIndex        =   117
         Text            =   "txtftin"
         Top             =   705
         Width           =   735
      End
      Begin VB.TextBox txtftin 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   0
         Left            =   780
         TabIndex        =   116
         Text            =   "txtftin"
         Top             =   345
         Width           =   735
      End
      Begin VB.Label Labcm 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Labcm"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   180
         TabIndex        =   119
         Top             =   1545
         Width           =   1380
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Input in"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   115
         Top             =   720
         Width           =   645
      End
      Begin VB.Label Label6 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Input ft"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   114
         Top             =   390
         Width           =   645
      End
   End
   Begin VB.CommandButton cmdAccCanRPlan 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   11985
      TabIndex        =   112
      Top             =   720
      Width           =   810
   End
   Begin VB.CommandButton cmdAccCanRPlan 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Accept"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   11970
      TabIndex        =   10
      Top             =   285
      Width           =   810
   End
   Begin VB.CommandButton cmdResetRulers 
      BackColor       =   &H00E0E0E0&
      Caption         =   "R"
      Height          =   270
      Left            =   60
      TabIndex        =   39
      ToolTipText     =   " Reset Rulers "
      Top             =   1350
      Width           =   270
   End
   Begin VB.Frame fraFurn 
      BackColor       =   &H00C0C0C0&
      Caption         =   "Rectangular items - cm "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10230
      Left            =   15795
      TabIndex        =   20
      Top             =   60
      Width           =   3180
      Begin VB.CommandButton cmdDnArrow 
         BackColor       =   &H00C0C0C0&
         Height          =   300
         Index           =   9
         Left            =   2820
         Picture         =   "Main.frx":0ABA
         Style           =   1  'Graphical
         TabIndex        =   165
         ToolTipText     =   "Copy W & D, 9 to 10 "
         Top             =   9015
         Width           =   300
      End
      Begin VB.CommandButton cmdUpArrow 
         BackColor       =   &H00C0C0C0&
         Height          =   300
         Index           =   9
         Left            =   2820
         Picture         =   "Main.frx":0BBC
         Style           =   1  'Graphical
         TabIndex        =   164
         ToolTipText     =   "Copy  W & D, 10 to 9 "
         Top             =   9330
         Width           =   300
      End
      Begin VB.Frame fraFurnRect 
         BackColor       =   &H00FFC0C0&
         Caption         =   "10."
         Height          =   870
         Index           =   10
         Left            =   75
         TabIndex        =   155
         Top             =   9255
         Width           =   2685
         Begin VB.CommandButton cmdFurn 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Can"
            Height          =   225
            Index           =   21
            Left            =   2205
            TabIndex        =   160
            ToolTipText     =   " Cancel "
            Top             =   525
            Width           =   435
         End
         Begin VB.CommandButton cmdFurn 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Acc"
            Height          =   225
            Index           =   20
            Left            =   2205
            TabIndex        =   159
            ToolTipText     =   " Accept "
            Top             =   240
            Width           =   435
         End
         Begin VB.TextBox txt 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   21
            Left            =   1455
            TabIndex        =   158
            Text            =   "txt(21)"
            Top             =   480
            Width           =   705
         End
         Begin VB.TextBox txt 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   20
            Left            =   540
            TabIndex        =   157
            Text            =   "txt(20)"
            Top             =   480
            Width           =   705
         End
         Begin VB.TextBox txtName 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   10
            Left            =   555
            TabIndex        =   156
            Text            =   "txtName(910)"
            Top             =   180
            Width           =   1605
         End
         Begin VB.Label Label5 
            BackColor       =   &H00FFC0C0&
            Caption         =   "D"
            Height          =   195
            Index           =   10
            Left            =   1305
            TabIndex        =   163
            Top             =   525
            Width           =   150
         End
         Begin VB.Label Label4 
            BackColor       =   &H00FFC0C0&
            Caption         =   "W"
            Height          =   195
            Index           =   10
            Left            =   345
            TabIndex        =   162
            Top             =   525
            Width           =   195
         End
         Begin VB.Label Label3 
            BackColor       =   &H00FFC0C0&
            Caption         =   "Name"
            Height          =   165
            Index           =   10
            Left            =   90
            TabIndex        =   161
            Top             =   210
            Width           =   600
         End
      End
      Begin VB.CommandButton cmdDnArrow 
         BackColor       =   &H00C0C0C0&
         Height          =   300
         Index           =   8
         Left            =   2820
         Picture         =   "Main.frx":0CBE
         Style           =   1  'Graphical
         TabIndex        =   154
         ToolTipText     =   "Copy W & D, 8 to 9 "
         Top             =   8115
         Width           =   300
      End
      Begin VB.CommandButton cmdUpArrow 
         BackColor       =   &H00C0C0C0&
         Height          =   300
         Index           =   8
         Left            =   2820
         Picture         =   "Main.frx":0DC0
         Style           =   1  'Graphical
         TabIndex        =   153
         ToolTipText     =   "Copy W & D, 9 to 8 "
         Top             =   8430
         Width           =   300
      End
      Begin VB.Frame fraFurnRect 
         BackColor       =   &H00FFFF80&
         Caption         =   "9."
         Height          =   870
         Index           =   9
         Left            =   75
         TabIndex        =   144
         Top             =   8355
         Width           =   2685
         Begin VB.TextBox txtName 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   9
            Left            =   555
            TabIndex        =   146
            Text            =   "txtName(9)"
            Top             =   180
            Width           =   1605
         End
         Begin VB.TextBox txt 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   19
            Left            =   540
            TabIndex        =   147
            Text            =   "txt(18)"
            Top             =   480
            Width           =   705
         End
         Begin VB.TextBox txt 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   18
            Left            =   1455
            TabIndex        =   148
            Text            =   "txt(19)"
            Top             =   480
            Width           =   705
         End
         Begin VB.CommandButton cmdFurn 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Acc"
            Height          =   225
            Index           =   18
            Left            =   2205
            TabIndex        =   149
            ToolTipText     =   " Accept "
            Top             =   240
            Width           =   435
         End
         Begin VB.CommandButton cmdFurn 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Can"
            Height          =   225
            Index           =   19
            Left            =   2205
            TabIndex        =   145
            ToolTipText     =   " Cancel "
            Top             =   525
            Width           =   435
         End
         Begin VB.Label Label3 
            BackColor       =   &H00FFFF80&
            Caption         =   "Name"
            Height          =   165
            Index           =   9
            Left            =   90
            TabIndex        =   152
            Top             =   210
            Width           =   600
         End
         Begin VB.Label Label4 
            BackColor       =   &H00FFFF80&
            Caption         =   "W"
            Height          =   195
            Index           =   9
            Left            =   345
            TabIndex        =   151
            Top             =   525
            Width           =   195
         End
         Begin VB.Label Label5 
            BackColor       =   &H00FFFF80&
            Caption         =   "D"
            Height          =   195
            Index           =   9
            Left            =   1305
            TabIndex        =   150
            Top             =   525
            Width           =   150
         End
      End
      Begin VB.CommandButton cmdUpArrow 
         BackColor       =   &H00C0C0C0&
         Height          =   300
         Index           =   7
         Left            =   2820
         Picture         =   "Main.frx":0EC2
         Style           =   1  'Graphical
         TabIndex        =   143
         ToolTipText     =   "Copy W & D, 8 to 7 "
         Top             =   7530
         Width           =   300
      End
      Begin VB.CommandButton cmdUpArrow 
         BackColor       =   &H00C0C0C0&
         Height          =   300
         Index           =   6
         Left            =   2820
         Picture         =   "Main.frx":0FC4
         Style           =   1  'Graphical
         TabIndex        =   142
         ToolTipText     =   "Copy W & D, 7 to 6 "
         Top             =   6615
         Width           =   300
      End
      Begin VB.CommandButton cmdUpArrow 
         BackColor       =   &H00C0C0C0&
         Height          =   300
         Index           =   5
         Left            =   2820
         Picture         =   "Main.frx":10C6
         Style           =   1  'Graphical
         TabIndex        =   141
         ToolTipText     =   "Copy W & D, 6 to 5 "
         Top             =   5715
         Width           =   300
      End
      Begin VB.CommandButton cmdUpArrow 
         BackColor       =   &H00C0C0C0&
         Height          =   300
         Index           =   4
         Left            =   2820
         Picture         =   "Main.frx":11C8
         Style           =   1  'Graphical
         TabIndex        =   140
         ToolTipText     =   "Copy W & D, 5 to 4 "
         Top             =   4800
         Width           =   300
      End
      Begin VB.CommandButton cmdUpArrow 
         BackColor       =   &H00C0C0C0&
         Height          =   300
         Index           =   3
         Left            =   2820
         Picture         =   "Main.frx":12CA
         Style           =   1  'Graphical
         TabIndex        =   139
         ToolTipText     =   "Copy W & D, 4 to 3 "
         Top             =   3900
         Width           =   300
      End
      Begin VB.CommandButton cmdUpArrow 
         BackColor       =   &H00C0C0C0&
         Height          =   300
         Index           =   2
         Left            =   2820
         Picture         =   "Main.frx":13CC
         Style           =   1  'Graphical
         TabIndex        =   138
         ToolTipText     =   "Copy W & D, 3 to 2 "
         Top             =   3000
         Width           =   300
      End
      Begin VB.CommandButton cmdUpArrow 
         BackColor       =   &H00C0C0C0&
         Height          =   300
         Index           =   1
         Left            =   2820
         Picture         =   "Main.frx":14CE
         Style           =   1  'Graphical
         TabIndex        =   137
         ToolTipText     =   "Copy W & D, 2 to 1 "
         Top             =   2115
         Width           =   300
      End
      Begin VB.CommandButton cmdUpArrow 
         BackColor       =   &H00C0C0C0&
         Height          =   300
         Index           =   0
         Left            =   2820
         Picture         =   "Main.frx":15D0
         Style           =   1  'Graphical
         TabIndex        =   136
         ToolTipText     =   "Copy W & D, 1 to 0 "
         Top             =   1215
         Width           =   300
      End
      Begin VB.CommandButton cmdDnArrow 
         BackColor       =   &H00C0C0C0&
         Height          =   300
         Index           =   7
         Left            =   2820
         Picture         =   "Main.frx":16D2
         Style           =   1  'Graphical
         TabIndex        =   128
         ToolTipText     =   "Copy W & D, 7 to 8 "
         Top             =   7215
         Width           =   300
      End
      Begin VB.CommandButton cmdDnArrow 
         BackColor       =   &H00C0C0C0&
         Height          =   300
         Index           =   6
         Left            =   2820
         Picture         =   "Main.frx":17D4
         Style           =   1  'Graphical
         TabIndex        =   127
         ToolTipText     =   "Copy W & D, 6 to 7 "
         Top             =   6300
         Width           =   300
      End
      Begin VB.CommandButton cmdDnArrow 
         BackColor       =   &H00C0C0C0&
         Height          =   300
         Index           =   5
         Left            =   2820
         Picture         =   "Main.frx":18D6
         Style           =   1  'Graphical
         TabIndex        =   126
         ToolTipText     =   "Copy W & D, 5 to 6 "
         Top             =   5400
         Width           =   300
      End
      Begin VB.CommandButton cmdDnArrow 
         BackColor       =   &H00C0C0C0&
         Height          =   300
         Index           =   4
         Left            =   2820
         Picture         =   "Main.frx":19D8
         Style           =   1  'Graphical
         TabIndex        =   125
         ToolTipText     =   "Copy W & D, 4 to 5 "
         Top             =   4485
         Width           =   300
      End
      Begin VB.CommandButton cmdDnArrow 
         BackColor       =   &H00C0C0C0&
         Height          =   300
         Index           =   3
         Left            =   2820
         Picture         =   "Main.frx":1ADA
         Style           =   1  'Graphical
         TabIndex        =   124
         ToolTipText     =   "Copy W & D, 3 to 4 "
         Top             =   3585
         Width           =   300
      End
      Begin VB.CommandButton cmdDnArrow 
         BackColor       =   &H00C0C0C0&
         Height          =   300
         Index           =   2
         Left            =   2820
         Picture         =   "Main.frx":1BDC
         Style           =   1  'Graphical
         TabIndex        =   123
         ToolTipText     =   "Copy W & D, 2 to 3 "
         Top             =   2685
         Width           =   300
      End
      Begin VB.CommandButton cmdDnArrow 
         BackColor       =   &H00C0C0C0&
         Height          =   300
         Index           =   1
         Left            =   2820
         Picture         =   "Main.frx":1CDE
         Style           =   1  'Graphical
         TabIndex        =   122
         ToolTipText     =   "Copy W & D, 1 to 2 "
         Top             =   1800
         Width           =   300
      End
      Begin VB.CommandButton cmdDnArrow 
         BackColor       =   &H00C0C0C0&
         Height          =   300
         Index           =   0
         Left            =   2820
         Picture         =   "Main.frx":1DE0
         Style           =   1  'Graphical
         TabIndex        =   121
         ToolTipText     =   "Copy W & D, 0 to 1 "
         Top             =   900
         Width           =   300
      End
      Begin VB.Frame fraFurnRect 
         BackColor       =   &H00FFFFC0&
         Caption         =   "8."
         Height          =   870
         Index           =   8
         Left            =   75
         TabIndex        =   102
         Top             =   7455
         Width           =   2685
         Begin VB.CommandButton cmdFurn 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Can"
            Height          =   225
            Index           =   17
            Left            =   2205
            TabIndex        =   107
            ToolTipText     =   " Cancel "
            Top             =   525
            Width           =   435
         End
         Begin VB.CommandButton cmdFurn 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Acc"
            Height          =   225
            Index           =   16
            Left            =   2205
            TabIndex        =   106
            ToolTipText     =   " Accept "
            Top             =   240
            Width           =   435
         End
         Begin VB.TextBox txt 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   17
            Left            =   1455
            TabIndex        =   105
            Text            =   "txt(17)"
            Top             =   480
            Width           =   705
         End
         Begin VB.TextBox txt 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   16
            Left            =   540
            TabIndex        =   104
            Text            =   "txt(16)"
            Top             =   480
            Width           =   705
         End
         Begin VB.TextBox txtName 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   8
            Left            =   555
            TabIndex        =   103
            Text            =   "txtName(8)"
            Top             =   180
            Width           =   1605
         End
         Begin VB.Label Label5 
            BackColor       =   &H00FFFFC0&
            Caption         =   "D"
            Height          =   195
            Index           =   8
            Left            =   1305
            TabIndex        =   110
            Top             =   525
            Width           =   150
         End
         Begin VB.Label Label4 
            BackColor       =   &H00FFFFC0&
            Caption         =   "W"
            Height          =   195
            Index           =   8
            Left            =   345
            TabIndex        =   109
            Top             =   525
            Width           =   195
         End
         Begin VB.Label Label3 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Name"
            Height          =   165
            Index           =   8
            Left            =   90
            TabIndex        =   108
            Top             =   210
            Width           =   600
         End
      End
      Begin VB.Frame fraFurnRect 
         BackColor       =   &H0080FF80&
         Caption         =   "7."
         Height          =   870
         Index           =   7
         Left            =   75
         TabIndex        =   89
         Top             =   6555
         Width           =   2685
         Begin VB.TextBox txtName 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   7
            Left            =   555
            TabIndex        =   94
            Text            =   "txtName(7)"
            Top             =   180
            Width           =   1605
         End
         Begin VB.TextBox txt 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   15
            Left            =   1455
            TabIndex        =   92
            Text            =   "txt(15)"
            Top             =   480
            Width           =   705
         End
         Begin VB.TextBox txt 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   14
            Left            =   540
            TabIndex        =   91
            Text            =   "txt(14)"
            Top             =   480
            Width           =   705
         End
         Begin VB.CommandButton cmdFurn 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Can"
            Height          =   225
            Index           =   15
            Left            =   2205
            TabIndex        =   90
            ToolTipText     =   " Cancel "
            Top             =   525
            Width           =   435
         End
         Begin VB.CommandButton cmdFurn 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Acc"
            Height          =   225
            Index           =   14
            Left            =   2205
            TabIndex        =   93
            ToolTipText     =   " Accept "
            Top             =   240
            Width           =   435
         End
         Begin VB.Label Label3 
            BackColor       =   &H0080FF80&
            Caption         =   "Name"
            Height          =   165
            Index           =   7
            Left            =   90
            TabIndex        =   97
            Top             =   210
            Width           =   600
         End
         Begin VB.Label Label4 
            BackColor       =   &H0080FF80&
            Caption         =   "W"
            Height          =   195
            Index           =   7
            Left            =   345
            TabIndex        =   96
            Top             =   525
            Width           =   195
         End
         Begin VB.Label Label5 
            BackColor       =   &H0080FF80&
            Caption         =   "D"
            Height          =   195
            Index           =   7
            Left            =   1305
            TabIndex        =   95
            Top             =   525
            Width           =   150
         End
      End
      Begin VB.Frame fraFurnRect 
         BackColor       =   &H00C0FFC0&
         Caption         =   "6."
         Height          =   870
         Index           =   6
         Left            =   75
         TabIndex        =   80
         Top             =   5655
         Width           =   2685
         Begin VB.CommandButton cmdFurn 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Can"
            Height          =   225
            Index           =   13
            Left            =   2205
            TabIndex        =   85
            ToolTipText     =   " Cancel "
            Top             =   525
            Width           =   435
         End
         Begin VB.CommandButton cmdFurn 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Acc"
            Height          =   225
            Index           =   12
            Left            =   2190
            TabIndex        =   84
            ToolTipText     =   " Accept "
            Top             =   240
            Width           =   435
         End
         Begin VB.TextBox txt 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   13
            Left            =   1455
            TabIndex        =   83
            Text            =   "txt(13)"
            Top             =   480
            Width           =   705
         End
         Begin VB.TextBox txt 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   12
            Left            =   540
            TabIndex        =   82
            Text            =   "txt(12)"
            Top             =   480
            Width           =   705
         End
         Begin VB.TextBox txtName 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   6
            Left            =   555
            TabIndex        =   81
            Text            =   "txtName(6)"
            Top             =   180
            Width           =   1605
         End
         Begin VB.Label Label5 
            BackColor       =   &H00C0FFC0&
            Caption         =   "D"
            Height          =   195
            Index           =   6
            Left            =   1305
            TabIndex        =   88
            Top             =   525
            Width           =   150
         End
         Begin VB.Label Label4 
            BackColor       =   &H00C0FFC0&
            Caption         =   "W"
            Height          =   195
            Index           =   6
            Left            =   345
            TabIndex        =   87
            Top             =   525
            Width           =   195
         End
         Begin VB.Label Label3 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Name"
            Height          =   165
            Index           =   6
            Left            =   90
            TabIndex        =   86
            Top             =   210
            Width           =   600
         End
      End
      Begin VB.Frame fraFurnRect 
         BackColor       =   &H0080FFFF&
         Caption         =   "5."
         Height          =   870
         Index           =   5
         Left            =   75
         TabIndex        =   71
         Top             =   4755
         Width           =   2685
         Begin VB.TextBox txtName 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   5
            Left            =   555
            TabIndex        =   76
            Text            =   "txtName(5)"
            Top             =   180
            Width           =   1605
         End
         Begin VB.TextBox txt 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   11
            Left            =   1455
            TabIndex        =   74
            Text            =   "txt(11)"
            Top             =   480
            Width           =   705
         End
         Begin VB.TextBox txt 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   10
            Left            =   540
            TabIndex        =   73
            Text            =   "txt(10)"
            Top             =   480
            Width           =   705
         End
         Begin VB.CommandButton cmdFurn 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Can"
            Height          =   225
            Index           =   11
            Left            =   2205
            TabIndex        =   72
            ToolTipText     =   " Cancel "
            Top             =   525
            Width           =   435
         End
         Begin VB.CommandButton cmdFurn 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Acc"
            Height          =   225
            Index           =   10
            Left            =   2205
            TabIndex        =   75
            ToolTipText     =   " Accept "
            Top             =   225
            Width           =   435
         End
         Begin VB.Label Label3 
            BackColor       =   &H0080FFFF&
            Caption         =   "Name"
            Height          =   165
            Index           =   5
            Left            =   90
            TabIndex        =   79
            Top             =   210
            Width           =   600
         End
         Begin VB.Label Label4 
            BackColor       =   &H0080FFFF&
            Caption         =   "W"
            Height          =   195
            Index           =   5
            Left            =   345
            TabIndex        =   78
            Top             =   525
            Width           =   195
         End
         Begin VB.Label Label5 
            BackColor       =   &H0080FFFF&
            Caption         =   "D"
            Height          =   195
            Index           =   5
            Left            =   1305
            TabIndex        =   77
            Top             =   525
            Width           =   150
         End
      End
      Begin VB.Frame fraFurnRect 
         BackColor       =   &H00C0FFFF&
         Caption         =   "4."
         Height          =   870
         Index           =   4
         Left            =   75
         TabIndex        =   62
         Top             =   3855
         Width           =   2685
         Begin VB.CommandButton cmdFurn 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Can"
            Height          =   225
            Index           =   9
            Left            =   2205
            TabIndex        =   67
            ToolTipText     =   " Cancel "
            Top             =   525
            Width           =   435
         End
         Begin VB.CommandButton cmdFurn 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Acc"
            Height          =   225
            Index           =   8
            Left            =   2205
            TabIndex        =   66
            ToolTipText     =   " Accept "
            Top             =   225
            Width           =   435
         End
         Begin VB.TextBox txt 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   9
            Left            =   1455
            TabIndex        =   65
            Text            =   "txt(9)"
            Top             =   480
            Width           =   705
         End
         Begin VB.TextBox txt 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   8
            Left            =   555
            TabIndex        =   64
            Text            =   "txt(8)"
            Top             =   480
            Width           =   705
         End
         Begin VB.TextBox txtName 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   4
            Left            =   555
            TabIndex        =   63
            Text            =   "txtName(4)"
            Top             =   180
            Width           =   1605
         End
         Begin VB.Label Label5 
            BackColor       =   &H00C0FFFF&
            Caption         =   "D"
            Height          =   195
            Index           =   4
            Left            =   1305
            TabIndex        =   70
            Top             =   525
            Width           =   150
         End
         Begin VB.Label Label4 
            BackColor       =   &H00C0FFFF&
            Caption         =   "W"
            Height          =   195
            Index           =   4
            Left            =   360
            TabIndex        =   69
            Top             =   525
            Width           =   195
         End
         Begin VB.Label Label3 
            BackColor       =   &H00C0FFFF&
            Caption         =   "Name"
            Height          =   165
            Index           =   4
            Left            =   90
            TabIndex        =   68
            Top             =   210
            Width           =   600
         End
      End
      Begin VB.Frame fraFurnRect 
         BackColor       =   &H0080C0FF&
         Caption         =   "3."
         Height          =   870
         Index           =   3
         Left            =   75
         TabIndex        =   51
         Top             =   2955
         Width           =   2685
         Begin VB.TextBox txtName 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   3
            Left            =   555
            TabIndex        =   52
            Text            =   "txtName(3)"
            Top             =   180
            Width           =   1605
         End
         Begin VB.TextBox txt 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   7
            Left            =   1455
            TabIndex        =   54
            Text            =   "txt(7)"
            Top             =   480
            Width           =   705
         End
         Begin VB.TextBox txt 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   6
            Left            =   555
            TabIndex        =   53
            Text            =   "txt(6)"
            Top             =   480
            Width           =   705
         End
         Begin VB.CommandButton cmdFurn 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Can"
            Height          =   225
            Index           =   7
            Left            =   2205
            TabIndex        =   56
            ToolTipText     =   " Cancel "
            Top             =   525
            Width           =   435
         End
         Begin VB.CommandButton cmdFurn 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Acc"
            Height          =   225
            Index           =   6
            Left            =   2205
            TabIndex        =   55
            ToolTipText     =   " Accept "
            Top             =   240
            Width           =   435
         End
         Begin VB.Label Label3 
            BackColor       =   &H0080C0FF&
            Caption         =   "Name"
            Height          =   165
            Index           =   3
            Left            =   90
            TabIndex        =   59
            Top             =   210
            Width           =   600
         End
         Begin VB.Label Label4 
            BackColor       =   &H0080C0FF&
            Caption         =   "W"
            Height          =   195
            Index           =   3
            Left            =   360
            TabIndex        =   58
            Top             =   525
            Width           =   195
         End
         Begin VB.Label Label5 
            BackColor       =   &H0080C0FF&
            Caption         =   "D"
            Height          =   195
            Index           =   3
            Left            =   1305
            TabIndex        =   57
            Top             =   525
            Width           =   150
         End
      End
      Begin VB.Frame fraFurnRect 
         BackColor       =   &H00C0E0FF&
         Caption         =   "2."
         Height          =   870
         Index           =   2
         Left            =   75
         TabIndex        =   42
         Top             =   2055
         Width           =   2685
         Begin VB.CommandButton cmdFurn 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Can"
            Height          =   225
            Index           =   5
            Left            =   2205
            TabIndex        =   47
            ToolTipText     =   " Cancel "
            Top             =   525
            Width           =   435
         End
         Begin VB.CommandButton cmdFurn 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Acc"
            Height          =   225
            Index           =   4
            Left            =   2205
            TabIndex        =   46
            ToolTipText     =   " Accept "
            Top             =   255
            Width           =   435
         End
         Begin VB.TextBox txt 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   5
            Left            =   1455
            TabIndex        =   45
            Text            =   "txt(5)"
            Top             =   480
            Width           =   705
         End
         Begin VB.TextBox txt 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   4
            Left            =   540
            TabIndex        =   44
            Text            =   "txt(4)"
            Top             =   480
            Width           =   705
         End
         Begin VB.TextBox txtName 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   2
            Left            =   555
            TabIndex        =   43
            Text            =   "txtName(2)"
            Top             =   180
            Width           =   1605
         End
         Begin VB.Label Label5 
            BackColor       =   &H00C0E0FF&
            Caption         =   "D"
            Height          =   195
            Index           =   2
            Left            =   1305
            TabIndex        =   50
            Top             =   525
            Width           =   150
         End
         Begin VB.Label Label4 
            BackColor       =   &H00C0E0FF&
            Caption         =   "W"
            Height          =   195
            Index           =   2
            Left            =   345
            TabIndex        =   49
            Top             =   525
            Width           =   195
         End
         Begin VB.Label Label3 
            BackColor       =   &H00C0E0FF&
            Caption         =   "Name"
            Height          =   165
            Index           =   2
            Left            =   90
            TabIndex        =   48
            Top             =   210
            Width           =   600
         End
      End
      Begin VB.Frame fraFurnRect 
         BackColor       =   &H008080FF&
         Caption         =   "1."
         Height          =   879
         Index           =   1
         Left            =   75
         TabIndex        =   30
         Top             =   1155
         Width           =   2685
         Begin VB.TextBox txtName 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   1
            Left            =   555
            TabIndex        =   32
            Text            =   "txtName(1)"
            Top             =   180
            Width           =   1605
         End
         Begin VB.TextBox txt 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   3
            Left            =   1455
            TabIndex        =   35
            Text            =   "txt(3)"
            Top             =   480
            Width           =   705
         End
         Begin VB.TextBox txt 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   2
            Left            =   555
            TabIndex        =   34
            Text            =   "txt(2)"
            Top             =   495
            Width           =   705
         End
         Begin VB.CommandButton cmdFurn 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Can"
            Height          =   225
            Index           =   3
            Left            =   2205
            TabIndex        =   31
            ToolTipText     =   " Cancel "
            Top             =   525
            Width           =   435
         End
         Begin VB.CommandButton cmdFurn 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Acc"
            Height          =   225
            Index           =   2
            Left            =   2205
            TabIndex        =   36
            ToolTipText     =   " Accept "
            Top             =   225
            Width           =   435
         End
         Begin VB.Label Label3 
            BackColor       =   &H008080FF&
            Caption         =   "Name"
            Height          =   165
            Index           =   1
            Left            =   90
            TabIndex        =   38
            Top             =   210
            Width           =   600
         End
         Begin VB.Label Label4 
            BackColor       =   &H008080FF&
            Caption         =   "W"
            Height          =   195
            Index           =   1
            Left            =   360
            TabIndex        =   37
            Top             =   525
            Width           =   195
         End
         Begin VB.Label Label5 
            BackColor       =   &H008080FF&
            Caption         =   "D"
            Height          =   195
            Index           =   1
            Left            =   1305
            TabIndex        =   33
            Top             =   525
            Width           =   150
         End
      End
      Begin VB.Frame fraFurnRect 
         BackColor       =   &H00C0C0FF&
         Caption         =   "0."
         Height          =   870
         Index           =   0
         Left            =   75
         TabIndex        =   21
         Top             =   255
         Width           =   2685
         Begin VB.CommandButton cmdFurn 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Can"
            Height          =   225
            Index           =   1
            Left            =   2205
            TabIndex        =   29
            ToolTipText     =   " Cancel "
            Top             =   525
            Width           =   435
         End
         Begin VB.CommandButton cmdFurn 
            BackColor       =   &H00E0E0E0&
            Caption         =   "Acc"
            Height          =   225
            Index           =   0
            Left            =   2205
            TabIndex        =   26
            ToolTipText     =   " Accept "
            Top             =   240
            Width           =   435
         End
         Begin VB.TextBox txt 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   1
            Left            =   1455
            TabIndex        =   25
            Text            =   "txt(1)"
            Top             =   480
            Width           =   705
         End
         Begin VB.TextBox txt 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   0
            Left            =   540
            TabIndex        =   24
            Text            =   "txt(0)"
            Top             =   480
            Width           =   705
         End
         Begin VB.TextBox txtName 
            BackColor       =   &H00FFFFFF&
            Height          =   285
            Index           =   0
            Left            =   555
            TabIndex        =   23
            Text            =   "txtName(0)"
            Top             =   180
            Width           =   1605
         End
         Begin VB.Label Label5 
            BackColor       =   &H00C0C0FF&
            Caption         =   "D"
            Height          =   195
            Index           =   0
            Left            =   1305
            TabIndex        =   28
            Top             =   525
            Width           =   150
         End
         Begin VB.Label Label4 
            BackColor       =   &H00C0C0FF&
            Caption         =   "W"
            Height          =   195
            Index           =   0
            Left            =   345
            TabIndex        =   27
            Top             =   525
            Width           =   195
         End
         Begin VB.Label Label3 
            BackColor       =   &H00C0C0FF&
            Caption         =   "Name"
            Height          =   165
            Index           =   0
            Left            =   90
            TabIndex        =   22
            Top             =   210
            Width           =   600
         End
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFC0C0&
      Caption         =   "L-shaped size < W x D"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1185
      Index           =   1
      Left            =   8880
      TabIndex        =   17
      Top             =   90
      Width           =   2985
      Begin VB.TextBox TextWD 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   3
         Left            =   1590
         TabIndex        =   9
         Text            =   "TextWD"
         Top             =   765
         Width           =   1095
      End
      Begin VB.TextBox TextWD 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   2
         Left            =   1575
         TabIndex        =   8
         Text            =   "TextWD"
         Top             =   285
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Enter Width w cm"
         Height          =   225
         Index           =   3
         Left            =   150
         TabIndex        =   19
         Top             =   300
         Width           =   1410
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFC0C0&
         Caption         =   "Enter Depth d cm"
         Height          =   225
         Index           =   2
         Left            =   150
         TabIndex        =   18
         Top             =   765
         Width           =   1410
      End
   End
   Begin VB.CommandButton cmdRoomPlan 
      Height          =   960
      Index           =   4
      Left            =   4740
      Picture         =   "Main.frx":1EE2
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   315
      Width           =   960
   End
   Begin VB.CommandButton cmdRoomPlan 
      Height          =   960
      Index           =   3
      Left            =   3720
      Picture         =   "Main.frx":3324
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   315
      Width           =   960
   End
   Begin VB.CommandButton cmdRoomPlan 
      Height          =   960
      Index           =   2
      Left            =   2685
      Picture         =   "Main.frx":4766
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   315
      Width           =   960
   End
   Begin VB.CommandButton cmdRoomPlan 
      Height          =   960
      Index           =   1
      Left            =   1650
      Picture         =   "Main.frx":5BA8
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   315
      Width           =   960
   End
   Begin VB.PictureBox picRule 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   11310
      Index           =   1
      Left            =   75
      MousePointer    =   15  'Size All
      Picture         =   "Main.frx":6FEA
      ScaleHeight     =   754
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   17
      TabIndex        =   12
      Top             =   1650
      Width           =   255
   End
   Begin VB.PictureBox picRule 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Index           =   0
      Left            =   375
      MousePointer    =   15  'Size All
      Picture         =   "Main.frx":10954
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   1012
      TabIndex        =   11
      Top             =   1380
      Width           =   15180
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0E0FF&
      Caption         =   "Room size <= 1000 x 600 cm "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1170
      Index           =   0
      Left            =   5895
      TabIndex        =   3
      Top             =   105
      Width           =   2940
      Begin VB.TextBox TextWD 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   1
         Left            =   1665
         TabIndex        =   6
         Text            =   "TextWD"
         Top             =   690
         Width           =   1095
      End
      Begin VB.TextBox TextWD 
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Index           =   0
         Left            =   1680
         TabIndex        =   5
         Text            =   "TextWD"
         Top             =   270
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Enter Depth D cm"
         Height          =   225
         Index           =   1
         Left            =   150
         TabIndex        =   7
         Top             =   765
         Width           =   1410
      End
      Begin VB.Label Label2 
         BackColor       =   &H00C0E0FF&
         Caption         =   "Enter Width W cm"
         Height          =   225
         Index           =   0
         Left            =   150
         TabIndex        =   4
         Top             =   300
         Width           =   1410
      End
   End
   Begin VB.CommandButton cmdRoomPlan 
      BackColor       =   &H00E0E0E0&
      Height          =   960
      Index           =   0
      Left            =   300
      Picture         =   "Main.frx":1D332
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   330
      Width           =   960
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DrawMode        =   7  'Invert
      ForeColor       =   &H80000008&
      Height          =   4755
      Left            =   375
      ScaleHeight     =   315
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   587
      TabIndex        =   1
      Top             =   1695
      Width           =   8835
      Begin VB.Label LabFurn 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFC0C0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "10"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   10
         Left            =   1170
         MousePointer    =   2  'Cross
         TabIndex        =   167
         Top             =   870
         Width           =   420
      End
      Begin VB.Label LabFurn 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "9"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   9
         Left            =   645
         MousePointer    =   2  'Cross
         TabIndex        =   166
         Top             =   870
         Width           =   420
      End
      Begin VB.Shape shpCutOut 
         BackColor       =   &H00E0E0E0&
         FillStyle       =   4  'Upward Diagonal
         Height          =   2010
         Left            =   75
         Top             =   45
         Width           =   2415
      End
      Begin VB.Label LabFurn 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "8"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   8
         Left            =   150
         MousePointer    =   2  'Cross
         TabIndex        =   111
         Top             =   870
         Width           =   420
      End
      Begin VB.Label LabFurn 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FF80&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "7"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   7
         Left            =   1665
         MousePointer    =   2  'Cross
         TabIndex        =   101
         Top             =   495
         Width           =   420
      End
      Begin VB.Label LabFurn 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "6"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   6
         Left            =   1170
         MousePointer    =   2  'Cross
         TabIndex        =   100
         Top             =   480
         Width           =   420
      End
      Begin VB.Label LabFurn 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "5"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   5
         Left            =   660
         MousePointer    =   2  'Cross
         TabIndex        =   99
         Top             =   495
         Width           =   420
      End
      Begin VB.Label LabFurn 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "4"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   4
         Left            =   120
         MousePointer    =   2  'Cross
         TabIndex        =   98
         Top             =   495
         Width           =   420
      End
      Begin VB.Label LabFurn 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "3"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   3
         Left            =   1680
         MousePointer    =   2  'Cross
         TabIndex        =   61
         Top             =   120
         Width           =   420
      End
      Begin VB.Label LabFurn 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "2"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   2
         Left            =   1185
         MousePointer    =   2  'Cross
         TabIndex        =   60
         Top             =   120
         Width           =   420
      End
      Begin VB.Label LabFurn 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H008080FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "1"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   1
         Left            =   690
         MousePointer    =   2  'Cross
         TabIndex        =   41
         Top             =   120
         Width           =   420
      End
      Begin VB.Label LabFurn 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H80000008&
         Height          =   315
         Index           =   0
         Left            =   135
         MousePointer    =   2  'Cross
         TabIndex        =   40
         Top             =   120
         Width           =   420
      End
   End
   Begin VB.Image Image1 
      Height          =   960
      Left            =   14535
      Picture         =   "Main.frx":1E774
      Stretch         =   -1  'True
      Top             =   225
      Width           =   960
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Height          =   1020
      Index           =   4
      Left            =   4725
      Top             =   285
      Width           =   1050
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Height          =   1020
      Index           =   3
      Left            =   3690
      Top             =   285
      Width           =   1050
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Height          =   1020
      Index           =   2
      Left            =   2670
      Top             =   285
      Width           =   1050
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Height          =   1020
      Index           =   1
      Left            =   1590
      Top             =   285
      Width           =   1050
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Height          =   1020
      Index           =   0
      Left            =   255
      Top             =   300
      Width           =   1050
   End
   Begin VB.Label Label1 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Select a room plan:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   315
      TabIndex        =   2
      Top             =   30
      Width           =   2130
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuFileOps 
         Caption         =   "&Load room plan"
         Index           =   0
      End
      Begin VB.Menu mnuFileOps 
         Caption         =   "&Save room plan"
         Index           =   1
      End
      Begin VB.Menu mnuFileOps 
         Caption         =   "&Exit"
         Index           =   2
      End
   End
   Begin VB.Menu mnuPrintForm 
      Caption         =   "&Print"
   End
   Begin VB.Menu mnuConvert 
      Caption         =   "&Convert to cm"
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' 2D Room Planner by Robert Rayment March 2010

Option Explicit

' For copying to Clipboard
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, _
   ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
' Avoids click-thru from CommonDialog1
Private Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
' Open/save class for room plan's FileSpec$
Private CommonDialog1 As cOSDialog
Private PathSpec$, FileSpec$
Private aSaved As Boolean

' Main Room Plan (RP)
Private RPlanType As Long           ' 0 Rect, 1 TL, 2 TR, 3 BL, 4 BR cut outs
Private RPW As Long, RPD As Long    ' Basic rectangular Room Plan W & D in pixels
Private RPsw As Long, RPsd As Long  ' L-shaped Room Plan, cut-out sizes RPsw & RPsd in pixels

' Eleven Rectangular furniture items (Furn)
Private FurnNames$(0 To 10)
Private FurnColors(0 To 10) As Long
' Indexes
' fraFurn  txtwh (& Acc Can)
' (0)        0,   1
' (1)        2,   3
' (2)        4,   5
'
' (10)       20,  21
' LabFurn(0),,LabFurn(10)  ' 11 Furniture rectangles
'
' For moving furniture
Private aMouseDown As Boolean
Private FurnLeft As Single, FurnTop As Single

Private MaxWScale As Long   '1000 or 600
Private MaxDScale As Long   ' 600 or 300
' picRule(0) horz, picRule(1) vert
' For moving rulers
Private picRuleLeft As Long
Private picRuleTop As Long
' Screen.TwipsPerPixel
Private STX As Long, STY As Long
Const Numbers$ = "+0123456789."

Private Sub cmdRoomPlan_Click(Index As Integer)
' Main Room Plan buttons
' RPlanType As Long  '0 Rect, 1 TL, 2 TR, 3 BL, 4 BR cut outs
' Position Input boxes and Accept/Cancel buttons.
   ClearShapes ' Clear red markers
   Select Case Index
   Case 0   ' Main
      Frame1(0).Visible = True
      TextWD(0).SetFocus
      Frame1(1).Visible = False
      cmdAccCanRPlan(0).Left = 600
      cmdAccCanRPlan(1).Left = 600
      cmdAccCanRPlan(0).Visible = True
      cmdAccCanRPlan(1).Visible = True
   Case 1, 2, 3, 4   ' L-shapes
      Frame1(0).Visible = True
      Frame1(1).Visible = True
      TextWD(0).SetFocus
      cmdAccCanRPlan(0).Left = 800
      cmdAccCanRPlan(1).Left = 800
      cmdAccCanRPlan(0).Visible = True
      cmdAccCanRPlan(1).Visible = True
   End Select
   Shape1(Index).Visible = True  ' Red markers
   fraFXY.Visible = True
   RPlanType = Index
End Sub


' Main Room and cut-out dimensions - cm
Private Sub TextWD_KeyPress(Index As Integer, KeyAscii As Integer)
' Values picked up by cmdAccCanRPlan_Click
' Const Numbers$ = "+0123456789."
   If KeyAscii <> 8 Then   ' Backspace
      If InStr(Numbers$, Chr(KeyAscii)) = 0 Then
         'MsgBox "error"
         KeyAscii = 0
         Exit Sub
      Else
         If (Chr$(KeyAscii) = "-" Or Chr$(KeyAscii) = "+") And Len(TextWD(Index).Text) <> 0 Then
            'MsgBox "error"
            KeyAscii = 0
            Exit Sub
         End If
      End If
   End If
End Sub

Private Sub cmdAccCanRPlan_Click(Index As Integer)
'RPlanType As Long  '0 Rect, 1 TL, 2 TR, 3 BL, 4 BR cut outs
'RPW As Long, RPD As Long       ' Basic rectangular Room Plan W & D
'RPsw As Long, RPsd As Long     ' Basic L-shaped Room Plan sw & sd
   If Index = 0 Then  'Accept
      ' Get RPWD pixel values
      MaxWScale = 1000
      RPW = Val(TextWD(0).Text)
      RPD = Val(TextWD(1).Text)
      If Val(TextWD(0)) <= 500 And Val(TextWD(1)) <= 300 Then
         MaxWScale = 500
         RPW = 2 * Val(TextWD(0).Text)
         RPD = 2 * Val(TextWD(1).Text)
      End If
      
      If RPW <= 0 Or RPD <= 0 Then
         Beep
         Pic.Move 25, 110, 10, 10
         Exit Sub
      ElseIf RPW > 1000 Or RPD > 600 Then
         Beep
         'TextWD(0) = ""
         'TextWD(1) = ""
         RPW = 0
         RPD = 0
         Pic.Move 25, 110, 10, 10
         MsgBox "Dimensions too big" & vbCrLf & "Re-enter W <= Max Width & D <= Max Depth  ", vbOKOnly + vbExclamation, "Room plan"
         Exit Sub
      Else  ' Main dimensions OK
         SetScale
         Pic.Move 25, 110, RPW, RPD
         Select Case RPlanType
         Case 0   ' Main
         Case Else   '1,2,3,4 Get RPsw & RPsd, cut-out dimensions in pixels
            If MaxWScale = 1000 Then
               RPsw = Val(TextWD(2).Text)
               RPsd = Val(TextWD(3).Text)
            Else
               RPsw = 2 * Val(TextWD(2).Text)
               RPsd = 2 * Val(TextWD(3).Text)
            End If
            ' Test cut-out size
            If RPsw <= 0 Or RPsd <= 0 Then
               Beep
               Exit Sub
            End If
            If RPsw >= RPW Or RPsd >= RPD Then
               MsgBox "Cut out dimensions >= Main dimensions ", vbOKOnly + vbExclamation, "Room plan"
               Exit Sub
            End If
         End Select
         
         ' Position shpCutOut
         Select Case RPlanType
         Case 0   ' Main
         Case 1   ' TL
            shpCutOut.Move 0, 0, RPsw, RPsd
            shpCutOut.Visible = True
         Case 2   ' TR
            shpCutOut.Move Pic.Width - RPsw - 2, 0, RPsw, RPsd
            shpCutOut.Visible = True
         Case 3   ' BL
            shpCutOut.Move 0, Pic.Height - RPsd - 2, RPsw, RPsd
            shpCutOut.Visible = True
         Case 4   ' BR
            shpCutOut.Move Pic.Width - RPsw - 2, Pic.Height - RPsd - 2, RPsw, RPsd
            shpCutOut.Visible = True
         End Select
         fraFurn.Visible = True
         aSaved = False
      End If
   
   Else  ' Cancel
      fraFurn.Visible = False
      Pic.Move 25, 110, 10, 10
      aSaved = True
   End If
End Sub

' Furniture dimensions
Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
' Values picked up by cmdFurn_Click
' Const Numbers$ = "+0123456789."
   If KeyAscii <> 8 Then   ' Backspace
      If InStr(Numbers$, Chr(KeyAscii)) = 0 Then
         'MsgBox "error"
         KeyAscii = 0
         Exit Sub
      Else
         If (Chr$(KeyAscii) = "-" Or Chr$(KeyAscii) = "+") And Len(txt(Index).Text) <> 0 Then
            'MsgBox "error"
            KeyAscii = 0
            Exit Sub
         End If
      End If
   End If
End Sub

Private Sub cmdFurn_Click(Index As Integer)
' Furniture Accept/Cancel buttons
Dim FW As Long, FD As Long
Dim iLeft As Long, iTop As Long
   If (Index And 1) = 0 Then  ' Even 0,2,4,,16
      If MaxWScale = 1000 Then
         FW = Val(txt(Index))
         FD = Val(txt(Index + 1))
      Else  ' Double to pixels
         FW = 2 * Val(txt(Index))
         FD = 2 * Val(txt(Index + 1))
      End If
      If FW <= 0 Or FW >= RPW Then
         MsgBox "Furniture size = 0 or >= Room size", vbOKOnly + vbExclamation, "Room plan"
         Exit Sub
      End If
      If FD <= 0 Or FD >= RPD Then
         MsgBox "Furniture size = 0 or >= Room size", vbOKOnly + vbExclamation, "Room plan"
         Exit Sub
      End If
      LabFurn(Index \ 2).Width = FW
      LabFurn(Index \ 2).Height = FD
      LabFurn(Index \ 2).Visible = True
      
      ' Check edges
      iLeft = LabFurn(Index \ 2).Left
      iTop = LabFurn(Index \ 2).Top
      If iLeft < 1 Then iLeft = 1
      If iTop < 1 Then iTop = 1
      If iLeft >= RPW - FW - 2 Then
         iLeft = RPW - FW - 3
      End If
      If iTop >= RPD - FD - 2 Then
         iTop = RPD - FD - 3
      End If
      LabFurn(Index \ 2).Left = iLeft
      LabFurn(Index \ 2).Top = iTop
      LabFurn(Index \ 2).Refresh

   Else  ' Odd, Cancel button. Clear LabFurn.
      LabFurn(Index \ 2).Width = 50
      LabFurn(Index \ 2).Height = 50
      LabFurn(Index \ 2).Visible = False
   End If
   aSaved = False
End Sub

Private Sub cmdUpArrow_Click(Index As Integer)
' 0, 1 to 0 txt(2) to txt(0), txt(3) to txt(1)
' 1, 2 to 1 txt(4) to txt(2), txt(5) to txt(3) etc
   LabFurn(Index).Visible = True
   txt(Index * 2) = txt((Index + 1) * 2)
   txt(Index * 2 + 1) = txt(Index * 2 + 3)
   cmdFurn_Click (Index * 2)
   
   ' Reposition in case part outside room
   If LabFurn(Index).Left + LabFurn(Index).Width >= Pic.Width Then
      LabFurn(Index).Left = Pic.Width - LabFurn(Index).Width - 3
   End If
   If LabFurn(Index).Top + LabFurn(Index).Height >= Pic.Height Then
      LabFurn(Index).Top = Pic.Height - LabFurn(Index).Height - 3
   End If
   aSaved = False
End Sub

Private Sub cmdDnArrow_Click(Index As Integer)
' 0, 0 to 1 txt(0) to txt(2), txt(1) to txt(3)
' 1, 1 to 2 txt(2) to txt(4), txt(3) to txt(5) etc
   LabFurn(Index + 1).Visible = True
   txt((Index + 1) * 2) = txt(Index * 2)
   txt((Index + 1) * 2 + 1) = txt(Index * 2 + 1)
   cmdFurn_Click ((Index + 1) * 2)
   
   ' Reposition in case part outside room
   If LabFurn(Index + 1).Left + LabFurn(Index + 1).Width >= Pic.Width Then
      LabFurn(Index + 1).Left = Pic.Width - LabFurn(Index + 1).Width - 3
   End If
   If LabFurn(Index + 1).Top + LabFurn(Index + 1).Height >= Pic.Height Then
      LabFurn(Index + 1).Top = Pic.Height - LabFurn(Index + 1).Height - 3
   End If
   aSaved = False
End Sub


Private Sub Form_Load()
   PathSpec$ = App.Path
   If Right$(PathSpec$, 1) <> "\" Then PathSpec$ = PathSpec$ & "\"
   
   STX = Screen.TwipsPerPixelX
   STY = Screen.TwipsPerPixelY
   
   Frame1(0).Visible = False
   Frame1(1).Visible = False
   fraFurn.Visible = False
   fraFXY.Visible = False
   SetFurnColors
   ClearShapes
   ClearTextBoxes
   ClearLabFurns
   Clear_fraconvert
   Fix_Colors
   cmdResetRulers_Click
   cmdAccCanRPlan(0).Visible = False
   cmdAccCanRPlan(1).Visible = False
   Pic.Move 25, 110, 10, 10
   Pic.ScaleMode = 3 ' pixels
   Pic.AutoRedraw = False
   Pic.DrawMode = vbXorPen
   Me.Caption = "2D Room Planner by Robert Rayment (Min Res = 1280 x 768)"
   MaxWScale = 1000
   SetScale
   ClipBoardUsed = 0
   aMouseDown = False
   aSaved = True
   Show
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim resp As Long
   If ClipBoardUsed = 1 Then
      resp = MsgBox("Clipboard used, Clear ?", vbQuestion + vbYesNo + vbSystemModal, "Room Planner")
      If resp = vbYes Then Clipboard.Clear
   End If
   If aSaved = False Then
      resp = MsgBox("Plan not saved. Do you want to save it ? ", vbQuestion + vbYesNoCancel + vbSystemModal, "Room Planner")
      Select Case resp
      Case vbYes
         mnuFileOps_Click 1
      Case vbNo
      Case vbCancel
         Cancel = True
         Exit Sub
      End Select
   End If
   Unload Me
   End
End Sub

Private Sub SetScale()
' From mnuFileOps_Click 0, show appropriate rulers
   If MaxWScale > 500 Then
      'MaxWScale = 1000
      MaxDScale = 600
      picRule(0) = LoadResPicture("HRULE1000", 0)
      picRule(1) = LoadResPicture("VRULE600", 0)
      Frame1(0) = "Room size <= 1000 x 600 cm"
   Else
      'MaxWScale = 500
      MaxDScale = 300
      picRule(0) = LoadResPicture("HRULE500", 0)
      picRule(1) = LoadResPicture("VRULE300", 0)
      Frame1(0) = "Room size <= 500 x 300 cm"
   End If
End Sub

' Input & Output

Private Sub mnuFileOps_Click(Index As Integer)
Dim k As Integer
Dim FW As Long, FD As Long
Dim FLeft As Long, FTop As Long
Dim FreeF As Long
   Select Case Index
   Case 0   ' Load room plan
      
      GetLoadSave 0  ' FileSpec$  Load
      If Len(FileSpec$) = 0 Then Exit Sub
      
      ClearShapes
      ClearTextBoxes
      ClearLabFurns

      FreeF = FreeFile
      Open FileSpec$ For Input As FreeF
      'Open "SVTest.rpc" For Input As #1
      
      Input #FreeF, RPlanType
      Input #FreeF, MaxWScale
      
      SetScale
      
      Select Case RPlanType
      Case 0   ' Main
         Input #FreeF, RPW, RPD
         If MaxWScale = 1000 Then
            TextWD(0) = LTrim$(Str$(RPW))
            TextWD(1) = LTrim$(Str$(RPD))
         Else
            TextWD(0) = LTrim$(Str$(RPW / 2))
            TextWD(1) = LTrim$(Str$(RPD / 2))
         End If
         Frame1(0).Visible = True
         TextWD(0).SetFocus
         Frame1(1).Visible = False
         cmdAccCanRPlan_Click 0
         k = Int(RPlanType)
         cmdRoomPlan_Click k   ' Sets red marker and Accept/Cancel buttons
      Case 1, 2, 3, 4   ' Main & L-shapes
         Input #FreeF, RPW, RPD
         Input #FreeF, RPsw, RPsd
         If MaxWScale = 1000 Then
            TextWD(0) = LTrim$(Str$(RPW))
            TextWD(1) = LTrim$(Str$(RPD))
            TextWD(2) = LTrim$(Str$(RPsw))
            TextWD(3) = LTrim$(Str$(RPsd))
         Else
            TextWD(0) = LTrim$(Str$(RPW / 2))
            TextWD(1) = LTrim$(Str$(RPD / 2))
            TextWD(2) = LTrim$(Str$(RPsw / 2))
            TextWD(3) = LTrim$(Str$(RPsd / 2))
         End If
         Frame1(0).Visible = True
         Frame1(1).Visible = True
         TextWD(0).SetFocus
         cmdAccCanRPlan_Click 0
         k = Int(RPlanType)
         cmdRoomPlan_Click k   ' Sets red marker and Accept/Cancel buttons
         shpCutOut.Visible = True
      End Select
      
      For k = 0 To 20 Step 2 ' Always 11 furniture items
         Input #FreeF, FurnNames$(k \ 2)
         txtName(k \ 2) = FurnNames$(k \ 2)  ' 0,1,,10
         Input #FreeF, FW, FD
         Input #FreeF, FLeft, FTop
         If FW > 0 And FD > 0 Then
            If MaxWScale = 1000 Then
               txt(k) = LTrim$(FW)            ' 0,2,4,,
               txt(k + 1) = LTrim$(FD)        ' 1,3,5,,
            Else
               txt(k) = LTrim$(FW / 2)          ' 0,2,4,,
               txt(k + 1) = LTrim$(FD / 2)      ' 1,3,5,,
            End If
            cmdFurn_Click k
            LabFurn(k \ 2).Move FLeft, FTop   ' 0,1,,10  Needs testing if numbers put direct
                                           ' into file instead of actual LabFurn left & top
         End If
      Next k
      Close FreeF
      aSaved = True
         
   Case 1   ' Save room plan
   
      GetLoadSave 1  ' FileSpec$  Save
      If Len(FileSpec$) = 0 Then Exit Sub
      FreeF = FreeFile
      Open FileSpec$ For Output As FreeF
      'Open "SVTest.rpc" For Output As #1
      
      Print #FreeF, RPlanType
      MaxWScale = 1000
      If Val(TextWD(0)) <= 500 And Val(TextWD(1)) <= 300 Then
         MaxWScale = 500
      End If
      
      Print #FreeF, MaxWScale
      Select Case RPlanType
      Case 0   ' Main
         Print #FreeF, RPW, RPD
      Case 1, 2, 3, 4   ' Main & L-shapes
         Print #FreeF, RPW, RPD
         Print #FreeF, RPsw, RPsd
      End Select
      
      For k = 0 To 20 Step 2 ' Always 11 furniture items
         If txtName(k \ 2) = "" Then txtName(k \ 2) = Str$(k \ 2)
         FurnNames$(k \ 2) = txtName(k \ 2)  ' 0,1,,10
         Print #FreeF, FurnNames$(k \ 2)
         If txt(k) = "" Then txt(k) = "0"
         If txt(k + 1) = "" Then txt(k + 1) = "0"
         If MaxWScale = 1000 Then
            FW = txt(k)
            FD = txt(k + 1)
         Else
            FW = 2 * txt(k)
            FD = 2 * txt(k + 1)
         End If
         Print #FreeF, FW, FD
         FLeft = LabFurn(k \ 2).Left
         FTop = LabFurn(k \ 2).Top
         Print #FreeF, FLeft, FTop
      Next k
      Close FreeF
      aSaved = True
   Case 2   ' exit
      Unload Me
   End Select
End Sub

Private Sub GetLoadSave(Index As Integer)
' To Private FileSpec$
Dim Title$, Filt$, InDir$
Dim FIndex As Long
   Set CommonDialog1 = New cOSDialog
   Select Case Index
   Case 0   ' Load
      Title$ = "Load room plan"
      Filt$ = "rpc|*.rpc"
      FileSpec$ = ""
      InDir$ = PathSpec$
      CommonDialog1.ShowOpen FileSpec$, Title$, Filt$, InDir$, "", Me.hWnd, FIndex
   Case 1   ' Save
      Title$ = "Save room plan"
      Filt$ = "rpc|*.rpc"
      FileSpec$ = ""
      InDir$ = PathSpec$
      CommonDialog1.ShowSave FileSpec$, Title$, Filt$, InDir$, "", Me.hWnd, FIndex
      FixExtension FileSpec$, ".rpc"
   End Select
   Set CommonDialog1 = Nothing
   SetCursorPos 10, 60
End Sub
 
Private Sub FixExtension(FSpec$, Ext$)
' In: FileSpec$ & Ext$ (".xxx")
Dim p As Long
   If Len(FSpec$) = 0 Then Exit Sub
   Ext$ = LCase$(Ext$)
   p = InStr(1, FSpec$, ".")
   If p = 0 Then
      FSpec$ = FSpec$ & Ext$
   Else
      FSpec$ = Mid$(FSpec$, 1, p - 1) & Ext$
   End If
End Sub

 
' Print Form1 via frmPrint
Private Sub mnuPrintForm_Click()
Dim resp As Long
   keybd_event vbKeySnapshot, 1, 0, 0    ' To Clipboard   ' 2nd param 0 screen, 1 active window
   DoEvents
   ClipBoardUsed = 1
   
   resp = MsgBox("Image on Clipboard." & vbCrLf & vbCrLf & "Is printer live ?", vbYesNo + vbQuestion + vbSystemModal, "RPlanner Print")
   If resp = vbNo Then Exit Sub
   
   frmPrint.Show vbModal
End Sub


' Move furniture
Private Sub LabFurn_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim iLeft As Single
Dim iTop As Single
Dim d As Single, a$
   If Button = vbLeftButton Then
      aMouseDown = True
      FurnLeft = X
      FurnTop = Y
      iLeft = LabFurn(Index).Left
      iTop = LabFurn(Index).Top
      LabItem = LTrim$(Str$(Index))
      If MaxWScale = 1000 Then
         LabFXY(0) = LTrim$(Str$(iLeft)) & " cm"
         LabFXY(1) = LTrim$(Str$(iTop)) & " cm"
      Else
         d = iLeft / 2
         a$ = LTrim$(Str$(iLeft / 2))
         If d = 0.5 Then a$ = "0" & a$
         a$ = a$ & " cm"
         LabFXY(0) = a$ 'LTrim$(Str$(iLeft / 2)) & " cm"
         d = iTop / 2
         a$ = LTrim$(Str$(iTop / 2))
         If d = 0.5 Then a$ = "0" & a$
         a$ = a$ & " cm"
         LabFXY(1) = a$ 'LTrim$(Str$(iTop / 2)) & " cm"
      End If
      ' Draw  PRODUCES EXCESSIVE FLICKERING if Pic.AutoRedraw = True
      Pic.Line (0, LabFurn(Index).Top - 2)-(LabFurn(Index).Left, LabFurn(Index).Top - 2), RGB(128, 128, 128)
      Pic.Line (LabFurn(Index).Left - 2, 0)-(LabFurn(Index).Left - 2, LabFurn(Index).Top), RGB(128, 128, 128)
      LabFurn(Index).BackColor = vbWhite
   End If
End Sub

Private Sub LabFurn_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'RPW, RPD
Dim iLeft As Single
Dim iTop As Single
Dim d As Single, a$
   If aMouseDown Then
         ' Clear
      Pic.Line (0, LabFurn(Index).Top - 2)-(LabFurn(Index).Left, LabFurn(Index).Top - 2), RGB(128, 128, 128)
      Pic.Line (LabFurn(Index).Left - 2, 0)-(LabFurn(Index).Left - 2, LabFurn(Index).Top), RGB(128, 128, 128)

      iLeft = LabFurn(Index).Left + (X - FurnLeft) / STX
      iTop = LabFurn(Index).Top + (Y - FurnTop) / STY
      If iLeft < 1 Then iLeft = 1
      If iTop < 1 Then iTop = 1
      If iLeft >= RPW - LabFurn(Index).Width - 2 Then
         iLeft = RPW - LabFurn(Index).Width - 3
      End If
      If iTop >= RPD - LabFurn(Index).Height - 2 Then
         iTop = RPD - LabFurn(Index).Height - 3
      End If
      LabFurn(Index).Left = iLeft
      LabFurn(Index).Top = iTop
      
      LabItem = LTrim$(Str$(Index))
      If MaxWScale = 1000 Then
         LabFXY(0) = LTrim$(Str$(iLeft)) & " cm"
         LabFXY(1) = LTrim$(Str$(iTop)) & " cm"
      Else
         d = iLeft / 2
         a$ = LTrim$(Str$(iLeft / 2))
         If d = 0.5 Then a$ = "0" & a$
         a$ = a$ & " cm"
         LabFXY(0) = a$ 'LTrim$(Str$(iLeft / 2)) & " cm"
         d = iTop / 2
         a$ = LTrim$(Str$(iTop / 2))
         If d = 0.5 Then a$ = "0" & a$
         a$ = a$ & " cm"
         LabFXY(1) = a$ 'LTrim$(Str$(iTop / 2)) & " cm"
      End If
      ' Draw
      Pic.Line (0, iTop - 2)-(iLeft, iTop - 2), RGB(128, 128, 128)
      Pic.Line (iLeft - 2, 0)-(iLeft - 2, iTop), RGB(128, 128, 128)
   End If
End Sub

Private Sub LabFurn_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   If aMouseDown Then
      ' Clear
      Pic.Line (0, LabFurn(Index).Top - 2)-(LabFurn(Index).Left, LabFurn(Index).Top - 2), RGB(128, 128, 128)
      Pic.Line (LabFurn(Index).Left - 2, 0)-(LabFurn(Index).Left - 2, LabFurn(Index).Top), RGB(128, 128, 128)
      LabFurn(Index).BackColor = FurnColors(Index)
      aMouseDown = False
      aSaved = False
   End If
End Sub

'Move Rulers
Private Sub picRule_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button = vbLeftButton Then
      picRuleLeft = X
      picRuleTop = Y
   ElseIf Button = vbRightButton Then
      picRule(Index).Visible = False
   End If
End Sub

Private Sub picRule_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Dim iLeft As Long
   Dim iTop As Long
   If Button = vbLeftButton Then
      iLeft = picRule(Index).Left + (X - picRuleLeft)
      iTop = picRule(Index).Top + (Y - picRuleTop)
      If iLeft < 1 Then iLeft = 1
      If iTop < -15 Then iTop = -15
      If iLeft > Form1.Width \ STX - 10 Then
         iLeft = (Form1.Width \ STX) - 15
      End If
      If iTop > (Form1.Height \ STY) - 58 Then
         iTop = (Form1.Height \ STY) - 66
      End If
      picRule(Index).Left = iLeft
      picRule(Index).Top = iTop
   End If
End Sub

Private Sub cmdResetRulers_Click()
   picRule(0).Visible = True
   picRule(1).Visible = True
   picRule(0).Move 25, 89
   picRule(1).Move 5, 110
End Sub

Private Sub ClearLabFurns()
Dim k As Long
   For k = 0 To 10
      LabFurn(k).Visible = False
   Next k
End Sub

Private Sub ClearShapes()
' Red square markers
Dim k As Long
 For k = 0 To 4
   Shape1(k).Visible = False
 Next k
 shpCutOut.Visible = False
End Sub

Private Sub ClearTextBoxes()
Dim k As Long
' Main dimensions
   TextWD(0) = ""
   TextWD(1) = ""
   TextWD(2) = ""
   TextWD(3) = ""
 For k = 0 To 10
   txtName(k) = Str$(k) ' Dummy names
   txt(2 * k) = ""
   txt(2 * k + 1) = ""
 Next k
End Sub

' Fix LabFurn() & fraFurnRect() colors
Private Sub Fix_Colors()
Dim k As Long
   For k = 0 To 10
      fraFurnRect(k).BackColor = FurnColors(k)
      LabFurn(k).BackColor = FurnColors(k)
   Next k
End Sub

Private Sub SetFurnColors()
'&HBBGGRR
   FurnColors(0) = &HC0C0FF
   FurnColors(1) = &H8080FF
   FurnColors(2) = &HC0E0FF
   FurnColors(3) = &H80C0FF
   FurnColors(4) = &HC0FFFF
   FurnColors(5) = &H80FFFF
   FurnColors(6) = &HC0FFC0
   FurnColors(7) = &H80FF80
   FurnColors(8) = &HFFFFC0
   FurnColors(9) = &HFFFF80
   FurnColors(10) = &HFFC0C0
End Sub

' Convert ft in to cm
Private Sub Clear_fraconvert()
   txtftin(0) = ""
   txtftin(1) = ""
   Labcm = ""
   fraConvert.Visible = False
End Sub

Private Sub mnuConvert_Click()
   fraConvert.Visible = True
   txtftin(0).SetFocus
End Sub

Private Sub cmdcm_Click()
Dim feet As Long, inches As Long, cml As Long, cms As Single
'feet = 6
'inches = 1
   If txtftin(0) = "" Then
      feet = 0
   Else
      feet = Val(txtftin(0))
   End If
   If txtftin(1) = "" Then
      inches = 0
   Else
      inches = Val(txtftin(1))
   End If

   inches = 12 * feet + inches
   If MaxWScale = 1000 Then
      cml = 2.54 * inches
      Labcm = cml & " cm"
   Else  ' Round to 0.5 cm
      cms = Round(2.54 * inches * 2) / 2
      Labcm = cms & " cm"
   End If
End Sub

Private Sub cmdClose_Click()
   fraConvert.Visible = False
End Sub

Private Sub txtftin_KeyPress(Index As Integer, KeyAscii As Integer)
' Input for Convert to cm
' Const Numbers$ = "+0123456789."
   If KeyAscii <> 8 Then   ' Backspace
      If InStr(Numbers$, Chr(KeyAscii)) = 0 Then
         'MsgBox "error"
         KeyAscii = 0
         Exit Sub
      Else
         If (Chr$(KeyAscii) = "-" Or Chr$(KeyAscii) = "+") And Len(TextWD(Index).Text) <> 0 Then
            'MsgBox "error"
            KeyAscii = 0
            Exit Sub
         End If
      End If
   End If
End Sub

Private Sub mnuHelp_Click()
Dim a$
   a$ = ""
   a$ = a$ & "First select a room plan then enter the overall" & vbCrLf
   a$ = a$ & "dimensions in cm." & vbCrLf
   a$ = a$ & "For L-shaped rooms enter the cut-out size." & vbCrLf
   a$ = a$ & "Press the Accept button." & vbCrLf & vbCrLf
   a$ = a$ & "Now enter the items' names, width & depth in cm" & vbCrLf
   a$ = a$ & "and press the Acc button." & vbCrLf & vbCrLf
   
   a$ = a$ & "The accuracy is to the nearest cm but if the overall" & vbCrLf
   a$ = a$ & "width and height is <= 500 x 300 cm then the " & vbCrLf
   a$ = a$ & "accuracy is to the nearest 0.5 cm." & vbCrLf & vbCrLf
   
   a$ = a$ & "Move each item around, with the mouse, until" & vbCrLf
   a$ = a$ & "at the required location." & vbCrLf & vbCrLf
   a$ = a$ & "When satisfied the plan can be saved as an *.rpc file." & vbCrLf
   a$ = a$ & "Also the plan can be printed." & vbCrLf & vbCrLf
   a$ = a$ & "The rulers can be moved, with the mouse, and reset" & vbCrLf
   a$ = a$ & "with the R button."
   MsgBox a$, vbOKOnly + vbInformation + vbSystemModal, "Room planner help"
   
End Sub


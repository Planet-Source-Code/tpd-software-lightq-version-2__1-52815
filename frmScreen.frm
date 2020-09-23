VERSION 5.00
Begin VB.Form frmScreen 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "#"
   ClientHeight    =   1725
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   3825
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmScreen.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MousePointer    =   99  'Custom
   ScaleHeight     =   115
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   255
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmr_LostFocus 
      Interval        =   500
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox picOnOff 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   57
      TabIndex        =   115
      Top             =   0
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.PictureBox mnu_update 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   7575
      Left            =   0
      ScaleHeight     =   505
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   635
      TabIndex        =   101
      Top             =   120
      Width           =   9525
      Begin VB.PictureBox background4 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   7575
         Left            =   0
         ScaleHeight     =   505
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   635
         TabIndex        =   102
         Top             =   0
         Width           =   9525
         Begin VB.PictureBox Picture8 
            Appearance      =   0  'Flat
            BackColor       =   &H00C00000&
            ForeColor       =   &H80000008&
            Height          =   5415
            Left            =   2760
            ScaleHeight     =   5385
            ScaleWidth      =   3945
            TabIndex        =   104
            Top             =   720
            Width           =   3975
            Begin VB.PictureBox updPrg 
               Height          =   255
               Left            =   240
               ScaleHeight     =   195
               ScaleWidth      =   3435
               TabIndex        =   108
               Top             =   4680
               Width           =   3495
            End
            Begin VB.ListBox lstUpdates 
               Appearance      =   0  'Flat
               BackColor       =   &H00C00000&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFC0C0&
               Height          =   2190
               ItemData        =   "frmScreen.frx":08CA
               Left            =   240
               List            =   "frmScreen.frx":08CC
               TabIndex        =   107
               Top             =   1320
               Width           =   3495
            End
            Begin VB.Label lblUpdates 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFC0C0&
               Height          =   240
               Left            =   240
               TabIndex        =   111
               Top             =   3510
               Width           =   60
            End
            Begin VB.Label but_downupdate 
               Alignment       =   2  'Center
               BackColor       =   &H00D06537&
               Caption         =   "DOWNLOAD SELECTED UPDATE"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   255
               Left            =   240
               TabIndex        =   110
               Top             =   3960
               Width           =   3495
            End
            Begin VB.Label Label16 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "DOWNLOAD PROGRESS"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFC0C0&
               Height          =   240
               Left            =   240
               TabIndex        =   109
               Top             =   4440
               Width           =   2310
            End
            Begin VB.Label but_chkupdate 
               Alignment       =   2  'Center
               BackColor       =   &H00D06537&
               Caption         =   "CHECK FOR UPDATES NOW !"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   255
               Left            =   240
               TabIndex        =   106
               Top             =   840
               Width           =   3495
            End
            Begin VB.Line Line28 
               X1              =   0
               X2              =   3960
               Y1              =   480
               Y2              =   480
            End
            Begin VB.Label Label21 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "CHECK / DOWNLOAD UPDATE"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   0
               TabIndex        =   105
               Top             =   120
               Width           =   3975
            End
         End
         Begin VB.PictureBox but_back4 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   525
            Left            =   600
            ScaleHeight     =   35
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   130
            TabIndex        =   103
            Top             =   6600
            Width           =   1950
         End
      End
   End
   Begin VB.PictureBox mnu_settings 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   7575
      Left            =   0
      ScaleHeight     =   505
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   635
      TabIndex        =   27
      Top             =   120
      Width           =   9525
      Begin VB.PictureBox background2 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   7575
         Left            =   -120
         ScaleHeight     =   505
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   635
         TabIndex        =   28
         Top             =   0
         Width           =   9525
         Begin VB.PictureBox but_back2 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   525
            Left            =   600
            ScaleHeight     =   35
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   130
            TabIndex        =   94
            Top             =   6600
            Width           =   1950
         End
         Begin VB.PictureBox Picture2 
            Appearance      =   0  'Flat
            BackColor       =   &H00C00000&
            ForeColor       =   &H80000008&
            Height          =   2655
            Left            =   2760
            ScaleHeight     =   2625
            ScaleWidth      =   3945
            TabIndex        =   34
            Top             =   3120
            Width           =   3975
            Begin VB.ComboBox cmbGrid 
               BackColor       =   &H00400000&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFC0C0&
               Height          =   330
               ItemData        =   "frmScreen.frx":08CE
               Left            =   240
               List            =   "frmScreen.frx":08DB
               Style           =   2  'Dropdown List
               TabIndex        =   120
               Top             =   1440
               Width           =   3495
            End
            Begin VB.PictureBox chk_settings 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H00C00000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   2
               Left            =   240
               ScaleHeight     =   16
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   16
               TabIndex        =   114
               Top             =   2040
               Width           =   240
            End
            Begin VB.ComboBox cmbScreen 
               BackColor       =   &H00400000&
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFC0C0&
               Height          =   330
               ItemData        =   "frmScreen.frx":08F3
               Left            =   240
               List            =   "frmScreen.frx":08FD
               Style           =   2  'Dropdown List
               TabIndex        =   35
               Top             =   840
               Width           =   3495
            End
            Begin VB.Label Label24 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "GRID STYLE"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFC0C0&
               Height          =   240
               Left            =   240
               TabIndex        =   119
               Top             =   1200
               Width           =   1155
            End
            Begin VB.Label Label23 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "MENU ANIMATION ENABLED"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFC0C0&
               Height          =   240
               Left            =   600
               TabIndex        =   118
               Top             =   2040
               Width           =   2670
            End
            Begin VB.Label Label10 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "SCREEN"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   0
               TabIndex        =   97
               Top             =   120
               Width           =   3975
            End
            Begin VB.Line Line26 
               X1              =   0
               X2              =   3960
               Y1              =   480
               Y2              =   480
            End
            Begin VB.Label Label19 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "SCREEN MODE"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFC0C0&
               Height          =   240
               Left            =   240
               TabIndex        =   36
               Top             =   600
               Width           =   1410
            End
         End
         Begin VB.PictureBox Picture1 
            Appearance      =   0  'Flat
            BackColor       =   &H00C00000&
            ForeColor       =   &H80000008&
            Height          =   2535
            Left            =   2760
            ScaleHeight     =   167
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   263
            TabIndex        =   29
            Top             =   360
            Width           =   3975
            Begin VB.PictureBox chk_settings 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H00C00000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   1
               Left            =   240
               ScaleHeight     =   16
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   16
               TabIndex        =   113
               Top             =   2040
               Width           =   240
            End
            Begin VB.PictureBox chk_settings 
               Appearance      =   0  'Flat
               AutoRedraw      =   -1  'True
               BackColor       =   &H00C00000&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   240
               Index           =   0
               Left            =   240
               ScaleHeight     =   16
               ScaleMode       =   3  'Pixel
               ScaleWidth      =   16
               TabIndex        =   112
               Top             =   1680
               Width           =   240
            End
            Begin VB.HScrollBar sVol 
               Height          =   192
               LargeChange     =   20
               Left            =   240
               Max             =   0
               Min             =   -2000
               SmallChange     =   5
               TabIndex        =   31
               Top             =   840
               Value           =   -1000
               Width           =   3495
            End
            Begin VB.HScrollBar mVol 
               Height          =   192
               LargeChange     =   20
               Left            =   240
               Max             =   1000
               Min             =   -2500
               SmallChange     =   5
               TabIndex        =   30
               Top             =   1320
               Value           =   -1000
               Width           =   3495
            End
            Begin VB.Label Label22 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "GAME MUSIC ENABLED"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFC0C0&
               Height          =   240
               Left            =   720
               TabIndex        =   117
               Top             =   2040
               Width           =   2220
            End
            Begin VB.Label Label20 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "GAME SOUND ENABLED"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFC0C0&
               Height          =   240
               Left            =   720
               TabIndex        =   116
               Top             =   1680
               Width           =   2280
            End
            Begin VB.Label Label8 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               Caption         =   "SOUND"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFFFFF&
               Height          =   255
               Left            =   0
               TabIndex        =   96
               Top             =   120
               Width           =   3975
            End
            Begin VB.Line Line25 
               X1              =   0
               X2              =   264
               Y1              =   32
               Y2              =   32
            End
            Begin VB.Label Label17 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "SOUND VOLUME"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFC0C0&
               Height          =   240
               Left            =   240
               TabIndex        =   33
               Top             =   600
               Width           =   1575
            End
            Begin VB.Label Label18 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "MUSIC VOLUME"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00FFC0C0&
               Height          =   225
               Left            =   240
               TabIndex        =   32
               Top             =   1080
               Width           =   1515
            End
         End
      End
   End
   Begin VB.PictureBox qLayer 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2655
      Left            =   10440
      ScaleHeight     =   175
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   303
      TabIndex        =   81
      Top             =   5400
      Width           =   4575
      Begin VB.Frame eQuestion 
         BackColor       =   &H00E88475&
         BorderStyle     =   0  'None
         Caption         =   "&H00C0C0FF&"
         Height          =   2415
         Left            =   0
         TabIndex        =   82
         Top             =   0
         Width           =   4335
         Begin VB.Frame frm_question 
            BackColor       =   &H00E88475&
            BorderStyle     =   0  'None
            Height          =   255
            Index           =   0
            Left            =   1320
            TabIndex        =   88
            Top             =   1920
            Width           =   1815
            Begin VB.Label but_question 
               Alignment       =   2  'Center
               BackColor       =   &H00D06537&
               Caption         =   "OK"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   255
               Index           =   2
               Left            =   480
               TabIndex        =   89
               Top             =   0
               Width           =   855
            End
         End
         Begin VB.Frame frm_question 
            BackColor       =   &H00E88475&
            BorderStyle     =   0  'None
            Height          =   255
            Index           =   4
            Left            =   1320
            TabIndex        =   85
            Top             =   1920
            Width           =   1815
            Begin VB.Label but_question 
               Alignment       =   2  'Center
               BackColor       =   &H00D06537&
               Caption         =   "NO"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   255
               Index           =   1
               Left            =   960
               TabIndex        =   87
               Top             =   0
               Width           =   855
            End
            Begin VB.Label but_question 
               Alignment       =   2  'Center
               BackColor       =   &H00D06537&
               Caption         =   "YES"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   86
               Top             =   0
               Width           =   855
            End
         End
         Begin VB.Image img_question 
            Height          =   480
            Index           =   4
            Left            =   240
            Top             =   480
            Width           =   480
         End
         Begin VB.Image img_question 
            Height          =   480
            Index           =   0
            Left            =   240
            Top             =   480
            Width           =   480
         End
         Begin VB.Line Line24 
            BorderColor     =   &H00800000&
            BorderWidth     =   3
            X1              =   0
            X2              =   4320
            Y1              =   2400
            Y2              =   2400
         End
         Begin VB.Line Line23 
            BorderColor     =   &H00800000&
            BorderWidth     =   3
            X1              =   4320
            X2              =   4320
            Y1              =   2400
            Y2              =   0
         End
         Begin VB.Line Line22 
            BorderColor     =   &H00FFC0C0&
            BorderWidth     =   3
            X1              =   0
            X2              =   4320
            Y1              =   0
            Y2              =   0
         End
         Begin VB.Line Line21 
            BorderColor     =   &H00FFC0C0&
            BorderWidth     =   3
            X1              =   0
            X2              =   0
            Y1              =   0
            Y2              =   2400
         End
         Begin VB.Label QuestionMess 
            BackStyle       =   0  'Transparent
            Caption         =   "#"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   1335
            Left            =   960
            TabIndex        =   84
            Top             =   480
            Width           =   3135
         End
         Begin VB.Label lblQuestionTitle 
            Alignment       =   2  'Center
            BackColor       =   &H00D06537&
            Caption         =   "#"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   240
            Left            =   30
            TabIndex        =   83
            Top             =   30
            Width           =   4275
         End
      End
   End
   Begin VB.PictureBox OverallLayer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   2280
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   153
      TabIndex        =   76
      Top             =   -120
      Width           =   2295
      Begin VB.Label Precach 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "PRECACHING"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   210
         TabIndex        =   77
         Top             =   120
         Width           =   1905
      End
   End
   Begin VB.PictureBox layerDEMO 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   315
      Left            =   0
      ScaleHeight     =   315
      ScaleWidth      =   2295
      TabIndex        =   73
      Top             =   0
      Visible         =   0   'False
      Width           =   2295
      Begin VB.Label lblDEMO 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "DEMO VERSION"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   330
         Left            =   0
         TabIndex        =   74
         Top             =   0
         Width           =   2235
      End
   End
   Begin VB.PictureBox sLayer 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   4095
      Left            =   16200
      ScaleHeight     =   271
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   215
      TabIndex        =   66
      Top             =   6120
      Width           =   3255
      Begin VB.PictureBox mnu_players 
         Appearance      =   0  'Flat
         BackColor       =   &H00800000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3855
         Left            =   120
         ScaleHeight     =   3855
         ScaleWidth      =   3015
         TabIndex        =   67
         Top             =   0
         Width           =   3015
         Begin VB.ListBox lstPlayers 
            Appearance      =   0  'Flat
            BackColor       =   &H00C00000&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFC0C0&
            Height          =   3150
            ItemData        =   "frmScreen.frx":0917
            Left            =   120
            List            =   "frmScreen.frx":0919
            TabIndex        =   68
            Top             =   480
            Width           =   1815
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            BackColor       =   &H00C00000&
            Caption         =   "SELECT PLAYER"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFC0C0&
            Height          =   240
            Left            =   30
            TabIndex        =   71
            Top             =   30
            Width           =   2955
         End
         Begin VB.Label but_player_ok 
            Alignment       =   2  'Center
            BackColor       =   &H00C00000&
            Caption         =   "OK"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFC0C0&
            Height          =   255
            Left            =   2040
            TabIndex        =   69
            Top             =   3360
            Width           =   855
         End
         Begin VB.Label player_remove 
            Alignment       =   2  'Center
            BackColor       =   &H00C00000&
            Caption         =   "REMOVE"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFC0C0&
            Height          =   255
            Left            =   2040
            TabIndex        =   70
            Top             =   840
            Width           =   855
         End
         Begin VB.Label player_add 
            Alignment       =   2  'Center
            BackColor       =   &H00C00000&
            Caption         =   "ADD"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFC0C0&
            Height          =   255
            Left            =   2040
            TabIndex        =   72
            Top             =   480
            Width           =   855
         End
         Begin VB.Line Line17 
            BorderColor     =   &H00FF8080&
            BorderWidth     =   3
            X1              =   0
            X2              =   0
            Y1              =   3840
            Y2              =   0
         End
         Begin VB.Line Line18 
            BorderColor     =   &H00FF8080&
            BorderWidth     =   3
            X1              =   3000
            X2              =   0
            Y1              =   0
            Y2              =   0
         End
         Begin VB.Line Line19 
            BorderColor     =   &H00404000&
            BorderWidth     =   3
            X1              =   3000
            X2              =   3000
            Y1              =   3840
            Y2              =   0
         End
         Begin VB.Line Line20 
            BorderColor     =   &H00404000&
            BorderWidth     =   3
            X1              =   0
            X2              =   3000
            Y1              =   3840
            Y2              =   3840
         End
      End
   End
   Begin VB.PictureBox eSettings 
      BackColor       =   &H00DA8561&
      BorderStyle     =   0  'None
      Height          =   2055
      Left            =   16560
      ScaleHeight     =   2055
      ScaleWidth      =   2775
      TabIndex        =   7
      Top             =   3840
      Width           =   2775
      Begin VB.Frame lay_col 
         BackColor       =   &H00DA8561&
         BorderStyle     =   0  'None
         Height          =   615
         Left            =   240
         TabIndex        =   12
         Top             =   600
         Width           =   2415
         Begin VB.OptionButton obj_col 
            BackColor       =   &H000000FF&
            Height          =   255
            Index           =   2
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   0
            Width           =   495
         End
         Begin VB.OptionButton obj_col 
            BackColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   1
            Left            =   1800
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   360
            Width           =   495
         End
         Begin VB.OptionButton obj_col 
            BackColor       =   &H0000FF00&
            Height          =   255
            Index           =   3
            Left            =   600
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   0
            Width           =   495
         End
         Begin VB.OptionButton obj_col 
            BackColor       =   &H00FF0000&
            Height          =   255
            Index           =   4
            Left            =   1200
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   0
            Width           =   495
         End
         Begin VB.OptionButton obj_col 
            BackColor       =   &H0000FFFF&
            Height          =   255
            Index           =   5
            Left            =   0
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   360
            Width           =   495
         End
         Begin VB.OptionButton obj_col 
            BackColor       =   &H00FFFF00&
            Height          =   255
            Index           =   6
            Left            =   600
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   360
            Width           =   495
         End
         Begin VB.OptionButton obj_col 
            BackColor       =   &H00FF00FF&
            Height          =   255
            Index           =   7
            Left            =   1200
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   360
            Width           =   495
         End
      End
      Begin VB.CheckBox obj_props 
         BackColor       =   &H00DA8561&
         Caption         =   "Movable"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   10
         Top             =   1320
         Width           =   1455
      End
      Begin VB.CheckBox obj_props 
         BackColor       =   &H00DA8561&
         Caption         =   "Rotatable"
         ForeColor       =   &H00800000&
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   9
         Top             =   1680
         Width           =   1455
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00800000&
         BorderWidth     =   3
         X1              =   0
         X2              =   2760
         Y1              =   2040
         Y2              =   2040
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   3
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   2040
      End
      Begin VB.Label but_settings_ok 
         Alignment       =   2  'Center
         BackColor       =   &H00D06537&
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   1800
         TabIndex        =   11
         Top             =   1680
         Width           =   855
      End
      Begin VB.Label lblSettings 
         Alignment       =   2  'Center
         BackColor       =   &H00D06537&
         Caption         =   "OBJECT SETTINGS"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   30
         TabIndex        =   21
         Top             =   30
         Width           =   2700
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "this object has no color settings"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   600
         Width           =   2415
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   3
         X1              =   0
         X2              =   2760
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00800000&
         BorderWidth     =   3
         X1              =   2760
         X2              =   2760
         Y1              =   2040
         Y2              =   0
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Beam color"
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   1110
      End
   End
   Begin VB.PictureBox eObjects 
      BackColor       =   &H00DA8561&
      BorderStyle     =   0  'None
      Height          =   2775
      Left            =   18000
      ScaleHeight     =   185
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   89
      TabIndex        =   6
      Top             =   10200
      Width           =   1335
      Begin VB.PictureBox scroll_layer 
         Appearance      =   0  'Flat
         BackColor       =   &H00000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2535
         Left            =   120
         ScaleHeight     =   169
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   73
         TabIndex        =   37
         Top             =   120
         Width           =   1095
         Begin VB.PictureBox edit_obj 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   975
            Left            =   0
            ScaleHeight     =   65
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   25
            TabIndex        =   38
            Top             =   0
            Width           =   375
         End
      End
      Begin VB.Line eObjects_TL 
         BorderColor     =   &H00FFC0C0&
         BorderWidth     =   3
         X1              =   0
         X2              =   88
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line eObjects_LL 
         BorderColor     =   &H00FFC0C0&
         BorderWidth     =   3
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   184
      End
      Begin VB.Line eObjects_BL 
         BorderColor     =   &H00800000&
         BorderWidth     =   3
         X1              =   0
         X2              =   88
         Y1              =   184
         Y2              =   184
      End
      Begin VB.Line eObjects_RL 
         BorderColor     =   &H00800000&
         BorderWidth     =   3
         X1              =   88
         X2              =   88
         Y1              =   248
         Y2              =   0
      End
   End
   Begin VB.PictureBox lev_sel 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   7575
      Left            =   0
      ScaleHeight     =   505
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   635
      TabIndex        =   0
      Top             =   120
      Width           =   9525
      Begin VB.PictureBox background1 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   7575
         Left            =   0
         ScaleHeight     =   505
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   635
         TabIndex        =   1
         Top             =   0
         Width           =   9525
         Begin VB.PictureBox but_back1 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   525
            Left            =   600
            ScaleHeight     =   35
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   130
            TabIndex        =   95
            Top             =   6600
            Width           =   1950
         End
         Begin VB.PictureBox Preview 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00C00000&
            ForeColor       =   &H00FFC0C0&
            Height          =   3810
            Left            =   4800
            ScaleHeight     =   252
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   265
            TabIndex        =   91
            Top             =   3240
            Width           =   4005
            Begin VB.Label lbl_lockstat 
               Alignment       =   2  'Center
               BackStyle       =   0  'Transparent
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   11.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   360
               Left            =   0
               TabIndex        =   93
               Top             =   1800
               Width           =   4020
            End
         End
         Begin VB.ListBox Levels 
            Appearance      =   0  'Flat
            BackColor       =   &H00C00000&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFC0C0&
            Height          =   2190
            ItemData        =   "frmScreen.frx":091B
            Left            =   4800
            List            =   "frmScreen.frx":091D
            TabIndex        =   3
            Top             =   600
            Width           =   3975
         End
         Begin VB.ListBox levPacks 
            Appearance      =   0  'Flat
            BackColor       =   &H00C00000&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFC0C0&
            Height          =   2190
            ItemData        =   "frmScreen.frx":091F
            Left            =   600
            List            =   "frmScreen.frx":0921
            TabIndex        =   2
            Top             =   600
            Width           =   3975
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "CURRENT PLAYER"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFC0C0&
            Height          =   240
            Left            =   600
            TabIndex        =   99
            Top             =   3960
            Width           =   1740
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PLAYER INFORMATION"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFC0C0&
            Height          =   240
            Left            =   600
            TabIndex        =   98
            Top             =   4680
            Width           =   2190
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "SELECT PACK"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFC0C0&
            Height          =   240
            Left            =   600
            TabIndex        =   5
            Top             =   360
            Width           =   1350
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "LEVEL PREVIEW"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFC0C0&
            Height          =   240
            Left            =   4800
            TabIndex        =   92
            Top             =   3000
            Width           =   1575
         End
         Begin VB.Image player_sel 
            Height          =   615
            Left            =   600
            Top             =   3120
            Width           =   1650
         End
         Begin VB.Label lbl_player_sel 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "SELECT PLAYER"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   210
            Left            =   600
            TabIndex        =   79
            Top             =   3330
            Width           =   1635
         End
         Begin VB.Label lbl_player 
            Appearance      =   0  'Flat
            BackColor       =   &H00C00000&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFC0C0&
            Height          =   300
            Left            =   600
            TabIndex        =   78
            Top             =   4230
            Width           =   3975
         End
         Begin VB.Label lbl_levstat 
            Appearance      =   0  'Flat
            BackColor       =   &H00C00000&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFC0C0&
            Height          =   300
            Left            =   600
            TabIndex        =   75
            Top             =   4950
            Width           =   3975
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "SELECT LEVEL"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFC0C0&
            Height          =   240
            Left            =   4800
            TabIndex        =   4
            Top             =   360
            Width           =   1425
         End
      End
   End
   Begin VB.PictureBox eLoad 
      BackColor       =   &H00DA8561&
      BorderStyle     =   0  'None
      Height          =   3615
      Left            =   15000
      ScaleHeight     =   3615
      ScaleWidth      =   4350
      TabIndex        =   22
      Top             =   120
      Width           =   4350
      Begin VB.ListBox lstBin 
         Appearance      =   0  'Flat
         BackColor       =   &H00D06537&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   2670
         ItemData        =   "frmScreen.frx":0923
         Left            =   120
         List            =   "frmScreen.frx":0925
         TabIndex        =   26
         Top             =   360
         Width           =   4095
      End
      Begin VB.Label but_newlevel 
         Alignment       =   2  'Center
         BackColor       =   &H00D06537&
         Caption         =   "NEW LEVEL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   80
         Top             =   3240
         Width           =   1215
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00D06537&
         Caption         =   "LOAD LEVEL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   240
         Left            =   30
         TabIndex        =   23
         Top             =   30
         Width           =   4260
      End
      Begin VB.Label but_loadlevel_cancel 
         Alignment       =   2  'Center
         BackColor       =   &H00D06537&
         Caption         =   "CANCEL"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   3360
         TabIndex        =   25
         Top             =   3240
         Width           =   855
      End
      Begin VB.Label but_loadlevel_ok 
         Alignment       =   2  'Center
         BackColor       =   &H00D06537&
         Caption         =   "OK"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   2400
         TabIndex        =   24
         Top             =   3240
         Width           =   855
      End
      Begin VB.Line Line12 
         BorderColor     =   &H00800000&
         BorderWidth     =   3
         X1              =   0
         X2              =   4320
         Y1              =   3600
         Y2              =   3600
      End
      Begin VB.Line Line11 
         BorderColor     =   &H00800000&
         BorderWidth     =   3
         X1              =   4320
         X2              =   4320
         Y1              =   3720
         Y2              =   0
      End
      Begin VB.Line Line10 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   3
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   3720
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00FF8080&
         BorderWidth     =   3
         X1              =   0
         X2              =   4320
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.PictureBox eLayer 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   2655
      Left            =   10440
      ScaleHeight     =   175
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   303
      TabIndex        =   55
      Top             =   2760
      Width           =   4575
      Begin VB.Frame eError 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         Caption         =   "&H00C0C0FF&"
         Height          =   2415
         Left            =   0
         TabIndex        =   56
         Top             =   0
         Width           =   4335
         Begin VB.Label lblErrTitle 
            Alignment       =   2  'Center
            BackColor       =   &H00000080&
            Caption         =   "#"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   240
            Left            =   30
            TabIndex        =   59
            Top             =   30
            Width           =   4275
         End
         Begin VB.Label ErrMes 
            BackStyle       =   0  'Transparent
            Caption         =   "#"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000080&
            Height          =   1335
            Left            =   240
            TabIndex        =   58
            Top             =   480
            Width           =   3855
         End
         Begin VB.Label but_error_ok 
            Alignment       =   2  'Center
            BackColor       =   &H000000C0&
            Caption         =   "OK"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H008080FF&
            Height          =   255
            Left            =   1800
            TabIndex        =   57
            Top             =   2040
            Width           =   855
         End
         Begin VB.Line Line5 
            BorderColor     =   &H00C0C0FF&
            BorderWidth     =   3
            X1              =   0
            X2              =   0
            Y1              =   0
            Y2              =   2400
         End
         Begin VB.Line Line6 
            BorderColor     =   &H00C0C0FF&
            BorderWidth     =   3
            X1              =   0
            X2              =   4320
            Y1              =   0
            Y2              =   0
         End
         Begin VB.Line Line7 
            BorderColor     =   &H00000080&
            BorderWidth     =   3
            X1              =   4320
            X2              =   4320
            Y1              =   2400
            Y2              =   0
         End
         Begin VB.Line Line8 
            BorderColor     =   &H00000080&
            BorderWidth     =   3
            X1              =   0
            X2              =   4320
            Y1              =   2400
            Y2              =   2400
         End
      End
   End
   Begin VB.PictureBox iLayer 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      ForeColor       =   &H80000008&
      Height          =   2775
      Left            =   10440
      ScaleHeight     =   183
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   303
      TabIndex        =   60
      Top             =   120
      Width           =   4575
      Begin VB.Frame eInput 
         BackColor       =   &H00800000&
         Height          =   2295
         Left            =   0
         TabIndex        =   61
         Top             =   240
         Width           =   4215
         Begin VB.TextBox txt_input 
            Appearance      =   0  'Flat
            BackColor       =   &H00C00000&
            ForeColor       =   &H00FFC0C0&
            Height          =   315
            Left            =   480
            TabIndex        =   62
            Top             =   960
            Width           =   3255
         End
         Begin VB.Label lbl_input 
            Alignment       =   2  'Center
            BackColor       =   &H00C00000&
            Caption         =   "#"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFC0C0&
            Height          =   240
            Left            =   30
            TabIndex        =   65
            Top             =   30
            Width           =   4155
         End
         Begin VB.Label but_input_cancel 
            Alignment       =   2  'Center
            BackColor       =   &H00C00000&
            Caption         =   "CANCEL"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFC0C0&
            Height          =   255
            Left            =   2160
            TabIndex        =   63
            Top             =   1800
            Width           =   855
         End
         Begin VB.Label but_input_ok 
            Alignment       =   2  'Center
            BackColor       =   &H00C00000&
            Caption         =   "OK"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFC0C0&
            Height          =   255
            Left            =   1200
            TabIndex        =   64
            Top             =   1800
            Width           =   855
         End
      End
   End
   Begin VB.PictureBox pack_sel 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   7575
      Left            =   0
      ScaleHeight     =   505
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   635
      TabIndex        =   39
      Top             =   120
      Width           =   9525
      Begin VB.PictureBox background3 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   7575
         Left            =   0
         ScaleHeight     =   505
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   635
         TabIndex        =   40
         Top             =   0
         Width           =   9525
         Begin VB.PictureBox but_back3 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   525
            Left            =   600
            ScaleHeight     =   35
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   130
            TabIndex        =   100
            Top             =   6600
            Width           =   1950
         End
         Begin VB.PictureBox Picture6 
            Appearance      =   0  'Flat
            BackColor       =   &H00C00000&
            ForeColor       =   &H80000008&
            Height          =   5295
            Left            =   3000
            ScaleHeight     =   5265
            ScaleWidth      =   3825
            TabIndex        =   41
            Top             =   720
            Width           =   3855
            Begin VB.OptionButton o_edit_pack 
               BackColor       =   &H00FF8080&
               Caption         =   "NEW PACK"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   300
               Index           =   1
               Left            =   240
               Style           =   1  'Graphical
               TabIndex        =   54
               Top             =   3120
               Width           =   3375
            End
            Begin VB.OptionButton o_edit_pack 
               BackColor       =   &H00FF8080&
               Caption         =   "EXISTING PACK"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00800000&
               Height          =   300
               Index           =   0
               Left            =   240
               Style           =   1  'Graphical
               TabIndex        =   53
               Top             =   180
               Value           =   -1  'True
               Width           =   3375
            End
            Begin VB.Frame Frame2 
               BackColor       =   &H00C00000&
               BorderStyle     =   0  'None
               Caption         =   "Frame2"
               Height          =   2775
               Left            =   240
               TabIndex        =   49
               Top             =   240
               Width           =   3375
               Begin VB.TextBox ed_ver_passw 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00800000&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFC0C0&
                  Height          =   315
                  IMEMode         =   3  'DISABLE
                  Left            =   1440
                  PasswordChar    =   "*"
                  TabIndex        =   90
                  Top             =   2160
                  Width           =   1935
               End
               Begin VB.ListBox lstEditPack 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00800000&
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFC0C0&
                  Height          =   1710
                  ItemData        =   "frmScreen.frx":0927
                  Left            =   0
                  List            =   "frmScreen.frx":0929
                  TabIndex        =   50
                  Top             =   360
                  Width           =   3375
               End
               Begin VB.Label Label11 
                  BackStyle       =   0  'Transparent
                  Caption         =   "PASSWORD"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFC0C0&
                  Height          =   240
                  Left            =   0
                  TabIndex        =   51
                  Top             =   2190
                  Width           =   1350
               End
            End
            Begin VB.Frame Frame1 
               BackColor       =   &H00C00000&
               BorderStyle     =   0  'None
               Caption         =   "Frame1"
               Height          =   1695
               Left            =   240
               TabIndex        =   42
               Top             =   2880
               Width           =   3375
               Begin VB.TextBox ed_retype 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00800000&
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFC0C0&
                  Height          =   315
                  IMEMode         =   3  'DISABLE
                  Left            =   1440
                  PasswordChar    =   "*"
                  TabIndex        =   45
                  Top             =   1380
                  Width           =   1935
               End
               Begin VB.TextBox ed_password 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00800000&
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFC0C0&
                  Height          =   315
                  IMEMode         =   3  'DISABLE
                  Left            =   1440
                  PasswordChar    =   "*"
                  TabIndex        =   44
                  Top             =   1020
                  Width           =   1935
               End
               Begin VB.TextBox ed_packname 
                  Appearance      =   0  'Flat
                  BackColor       =   &H00800000&
                  Enabled         =   0   'False
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFC0C0&
                  Height          =   315
                  IMEMode         =   3  'DISABLE
                  Left            =   1440
                  TabIndex        =   43
                  Top             =   640
                  Width           =   1935
               End
               Begin VB.Label Label15 
                  BackStyle       =   0  'Transparent
                  Caption         =   "RETYPE"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFC0C0&
                  Height          =   240
                  Left            =   0
                  TabIndex        =   48
                  Top             =   1440
                  Width           =   1350
               End
               Begin VB.Label Label14 
                  BackStyle       =   0  'Transparent
                  Caption         =   "PASSWORD"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFC0C0&
                  Height          =   240
                  Left            =   0
                  TabIndex        =   47
                  Top             =   1080
                  Width           =   1350
               End
               Begin VB.Label Label12 
                  BackStyle       =   0  'Transparent
                  Caption         =   "PACKNAME"
                  BeginProperty Font 
                     Name            =   "Arial"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FFC0C0&
                  Height          =   240
                  Left            =   0
                  TabIndex        =   46
                  Top             =   685
                  Width           =   1350
               End
            End
            Begin VB.Line Line16 
               X1              =   3600
               X2              =   3840
               Y1              =   3240
               Y2              =   3240
            End
            Begin VB.Line Line15 
               X1              =   0
               X2              =   240
               Y1              =   3240
               Y2              =   3240
            End
            Begin VB.Line Line14 
               X1              =   3600
               X2              =   3840
               Y1              =   300
               Y2              =   300
            End
            Begin VB.Line Line13 
               X1              =   0
               X2              =   240
               Y1              =   300
               Y2              =   300
            End
            Begin VB.Label but_edit_ok 
               Alignment       =   2  'Center
               BackColor       =   &H00FF8080&
               Caption         =   ">>"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00004000&
               Height          =   255
               Left            =   2760
               TabIndex        =   52
               Top             =   4800
               Width           =   855
            End
         End
      End
   End
End
Attribute VB_Name = "frmScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'################################################################################################################## _
 #                                                                                                                # _
 #  TPD SOFTWARE PROUDLY PRESENTS:                                                                                # _
 #                                                                                                                # _
 #          ///           ///      //////       ///      ///   ////////////      ///////                          # _
 #          ///           ///    /////////      ///      ///   ////////////    //////////                         # _
 #         ///           ///    ///            ///      ///        ///        ///      ///                        # _
 #         ///           ///   ///             ///      ///        ///       ///        ///                       # _
 #        ///           ///    ///    /////   ////////////        ///        ///        ///                       # _
 #        ///           ///   ///      ////   ////////////        ///       ///         ///                       # _
 #       ///           ///    ///      ///   ///      ///        ///        ///      /////                        # _
 #       ///           ///    ///    ////    ///      ///        ///         ///      /////                       # _
 #      ///////////   ///     //////////    ///      ///        ///           //////////////                      # _
 #      ///////////   ///      ///////      ///      ///        ///            ///////   ///      VERSION 2.0     # _
 #                                                                                                                # _
 ##################################################################################################################

Private RestoreOnActivate As Boolean

Private Sub Form_KeyPress(KeyAscii As Integer)
  strCheat = strCheat & Chr(KeyAscii)
  
  ' [NEW 2.0.1] CheatCode: freezemines -> Mines will not move!
  If InStr(1, strCheat, "freezemines", vbTextCompare) Then
     PlaySound "BUTTON"
     arCheats(0) = True
     strCheat = vbNullString
  End If
  
  ' [NEW 2.0.1] CheatCode: dontblowme -> Barrels do not explode
  If InStr(1, strCheat, "dontblowme", vbTextCompare) Then
     PlaySound "BUTTON"
     arCheats(1) = True
     strCheat = vbNullString
  End If
  
  ' [NEW 2.0.1] CheatCode: iamthebest -> Standard pack all level cheat
  If InStr(1, strCheat, "iamthebest", vbTextCompare) Then
     PlaySound "BUTTON"
     arCheats(2) = True
     lbl_player = "GOD"
     lbl_lockstat = vbNullString
     lbl_lockstat.Tag = vbNullString
     levPacks.Tag = "FREE"
     lbl_levstat = "ALL LEVEL CHEAT ENABLED"
     ExecutionState = EXEC_SELECT
     UpdatePlayerStats
     strCheat = vbNullString
  End If
  
  ' [NEW 2.0.1] CheatCode: dontbesuperman -> Reset all cheats
  If InStr(1, strCheat, "dontbesuperman", vbTextCompare) Then
     PlaySound "BUTTON"
     arCheats(0) = False
     arCheats(1) = False
     arCheats(2) = False
     levPacks_Click
     strCheat = vbNullString
  End If
  
End Sub

Private Sub Form_Load()
  
  'set us on top of all others forms
  FormOnTop Me, True
   
  'set overall layer
  With OverallLayer
    .AutoRedraw = True
    .Move 0, 0, XRes * Screen.TwipsPerPixelX, YRes * Screen.TwipsPerPixelY
    .ZOrder 0
    .Refresh
    .Cls
    Precach.Move XRes \ 2 - Precach.Width \ 2, YRes \ 2 - Precach.Height
  End With
  
  DEMOMODE = False
  
  'app title
  Dim dt As String
  If DEMOMODE Then dt = "Demo" Else dt = vbNullString
  Caption = Replace(Replace(GAME_TITLE, "%a", dt), "%v", App.Major & "." & App.Minor & "." & App.Revision)
  'set mouse pointer
  SetPointer "normal"
  'init directx sound enigne
  InitDirectAudio
  'load game settings (sound, screen mode ...)
  LoadSettings
  'init screen, make backbuffer and do full screen switch
  InitScreen
    
  With layerDEMO
     .Move XRes - .Width, 0
  End With
  
  ' select pack menu
  With pack_sel
     .Visible = False
     .Move 100, 85, 635, 505
  End With
  ' select level menu
  With lev_sel
     .Visible = False
     .Move 100, 85, 635, 505
  End With
  ' settings menu
  With mnu_settings
     .Visible = False
     .Move 100, 85, 635, 505
  End With
  '[ADDON 2.0.1] update menu
  With mnu_update
     .Visible = False
     .Move 100, 85, 635, 505
  End With
  ' object settings menu
  With eSettings
    .Visible = False
    .Move 300, frmScreen.ScaleHeight - .Height - SP_H
  End With
  ' error layerwindow
  With eLayer
    .Visible = False
    .Move 0, 0, XRes * Screen.TwipsPerPixelX, YRes * Screen.TwipsPerPixelY
  End With
  ' error window
  With eError
    .Visible = True
    .Move frmScreen.ScaleWidth \ 2 - .Width \ 2, frmScreen.ScaleHeight \ 2 - .Height \ 2
  End With
  ' input layerwindow
  With iLayer
    .Visible = False
    .Move 0, 0, XRes * Screen.TwipsPerPixelX, YRes * Screen.TwipsPerPixelY
  End With
  ' enter level name menu
  With eInput
    .Visible = True
    .Move frmScreen.ScaleWidth \ 2 - .Width \ 2, frmScreen.ScaleHeight \ 2 - .Height \ 2
  End With
  ' question layerwindow
  With qLayer
    .Visible = False
    .Move 0, 0, XRes * Screen.TwipsPerPixelX, YRes * Screen.TwipsPerPixelY
  End With
  ' question window
  With eQuestion
    .Visible = True
    .Move frmScreen.ScaleWidth \ 2 - .Width \ 2, frmScreen.ScaleHeight \ 2 - .Height \ 2
    but_question(0).Tag = vbYes
    but_question(1).Tag = vbNo
    but_question(2).Tag = vbOK
  End With
  ' load level menu
  With eLoad
    .Visible = False
    .Move frmScreen.ScaleWidth \ 2 - .Width \ 2, frmScreen.ScaleHeight \ 2 - .Height \ 2
  End With
  ' player selection layerwindow
  With sLayer
    .Visible = False
    .Move 0, 0, XRes * Screen.TwipsPerPixelX, YRes * Screen.TwipsPerPixelY
  End With
  ' player selection menu
  With mnu_players
    .Visible = True
    .Move frmScreen.ScaleWidth \ 2 - .Width \ 2, frmScreen.ScaleHeight \ 2 - .Height \ 2
  End With

  'create game matrix
  ReDim gMatrix(CL_C - 1, CL_R - 1)
  'load object maps
  LoadObjects
  'load sprite coordination map and behavious
  LoadCoordMap
  'initialize edit objects and menu's
  InitEditObjects
  'load colors
  LoadColor
  'load system fonts
  LoadFonts
  'load available players
  DoLoadPlayers
  'clear/create matrices
  ClearMatrices
  'prelink a progressbar to the HTTP class
  HttpEx.LinkProgressBar updPrg, vbWhite, vbBlue, &H800000
  
  On Error GoTo ImageLoadError
  
  'load all images into memory
  hSprite(H_SPRITES) = LoadGraphicDC("sprites.bmp")
  hSprite(H_MINES) = LoadGraphicDC("mine.bmp")
  hSprite(H_CTRLBAR) = LoadGraphicDC("controlbar.bmp")
  hSprite(H_MENU) = LoadGraphicDC("menu.bmp")
  hSprite(H_PLAY) = LoadGraphicDC("play.bmp")
  hSprite(H_EDIT) = LoadGraphicDC("edit.bmp")
  hSprite(H_SETTINGS) = LoadGraphicDC("settings.bmp")
  hSprite(H_UPDATE) = LoadGraphicDC("update.bmp")
  hSprite(H_VISIT) = LoadGraphicDC("visit.bmp")
  hSprite(H_EXIT) = LoadGraphicDC("exit.bmp")
  hSprite(H_DISABLED) = LoadGraphicDC("cross.bmp")
  hSprite(H_EDITBAR) = LoadGraphicDC("editbar.bmp")
  hSprite(H_TPDLOGO) = LoadGraphicDC("tpdlogo.bmp")
  hSprite(H_GAMELOGO) = LoadGraphicDC("lightqlogo.bmp")
  hSprite(H_SMALLBUTTON) = LoadGraphicDC("smallbutton.bmp")
  hSprite(H_POINT) = LoadPicture(App.Path & "\data\graphics\normal.cur")
  hSprite(H_MOVE) = LoadPicture(App.Path & "\data\graphics\move.cur")
    
  background1.Picture = LoadPicture(App.Path & "\data\graphics\background.bmp")
  background2.Picture = background1.Picture
  background3.Picture = background1.Picture
  background4.Picture = background1.Picture
  but_back1.Picture = LoadPicture(App.Path & "\data\graphics\back.bmp")
  but_back2.Picture = but_back1.Picture
  but_back3.Picture = but_back1.Picture
  but_back4.Picture = but_back1.Picture
  img_question(0).Picture = LoadPicture(App.Path & "\data\graphics\info.ico")
  img_question(4).Picture = LoadPicture(App.Path & "\data\graphics\question.ico")
  picOnOff.Picture = LoadPicture(App.Path & "\data\graphics\onoff.bmp")
    
  TransparentBlt background1.hdc, 40, 208, 110, 41, hSprite(H_SMALLBUTTON), 0, 0, 110, 41, vbBlack
  
  On Error GoTo 0
  
  'check image loading
  For i = LBound(hSprite) To UBound(hSprite)
     If hSprite(i) = 0 Then GoTo ImageLoadError
  Next i
  
  'show the graphical Checkboxes
  For i = chk_settings.LBound To chk_settings.UBound
     chk_settings_Show (i)
  Next i
   
  'apply grid styles
  gGrid(0) = vbNullString
  gGrid(1) = "GR_D"
  gGrid(2) = "GR_L"
    
  'disable the hiding-layer
  OverallLayer.Visible = False
  
  'show screen
  Show
  
  'intro
  If DEMOMODE Or InStr(1, LCase(Command), "/killintro", vbTextCompare) = 0 Then
     'Intro
  End If
  
  'when in DEMO mode show the DEMO logo
  If DEMOMODE Then
     layerDEMO.Visible = True
     layerDEMO.ZOrder 0
     layerDEMO.Refresh
  Else
     layerDEMO.Visible = False
  End If
  
  'switch to menu loop
  MenuLoop
  
ImageLoadError:
  Show
  OverallLayer.Visible = False
  ShowError "Some graphics could not be loaded."
  Form_Unload 0
  
End Sub

'[NEW 2.0.1] restore original screenmode when switching apps
Private Sub Form_Activate()
   'get active window (used to detect lostfocus in FULLSCREEN mode)
   ActiveHwnd = GetActiveWindow()
End Sub

Private Sub Form_Resize()
  ActiveHwnd = GetActiveWindow()
  'restore fullscreen when we want to drop back in
  If RestoreOnActivate Then
    SetScreenWithMaxRefresh XRes, YRes, Depth
    RestoreOnActivate = False
    SetFocus
  End If
End Sub

Private Sub Form_Initialize()
   ' initialize WinXP controls (when possible)
   InitWinXPControls
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If ExecutionState = EXEC_GAME + EXEC_EDIT Then
      DoShiftOnMatrix = KeyCode
   End If
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If eSettings.Visible = True Then Exit Sub
  If eObjects.Visible = True Then Exit Sub
  mDown Button, Shift, X, Y
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  mMove Button, Shift, X, Y
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
  If eSettings.Visible = True Then
     eSettings.Visible = False
     Exit Sub
  End If
  If eObjects.Visible = True Then
     eObjects.Visible = False
     Exit Sub
  End If
  mUp Button, Shift, X, Y
End Sub

Public Sub Form_Unload(Cancel As Integer)
  ClearDirectAudio
  KillScreen
  FormOnTop Me, False
  Unload frmScreen
  End
End Sub

Private Sub cmbScreen_Click()
  If EngineSelect Then Exit Sub
  InitScreen True
End Sub

Private Sub but_back1_Click()
  PlaySound "BUTTON"
  ExecutionState = EXEC_MENU
End Sub

Private Sub but_back2_Click()
  PlaySound "BUTTON"
  ExecutionState = EXEC_MENU
  SaveSettings
End Sub

Private Sub but_back3_Click()
  PlaySound "BUTTON"
  ExecutionState = EXEC_MENU
End Sub

Private Sub but_back4_Click()
  PlaySound "BUTTON"
  ExecutionState = EXEC_MENU
End Sub

Public Sub Levels_Click()
  If EngineSelect Then Exit Sub
  If levPacks.Tag = "FREE" Then
     lbl_lockstat = vbNullString
     lbl_lockstat.Tag = vbNullString
  Else
     If Len(lbl_player) = 0 Then
        lbl_lockstat.FontSize = 11
        lbl_lockstat = "SELECT A PLAYER FIRST"
        lbl_lockstat.Tag = "LOCK"
        Exit Sub
     Else
        If Val(Levels.List(Levels.ListIndex)) > Val(lbl_player.Tag) Then
           lbl_lockstat.FontSize = 14
           lbl_lockstat = "LOCKED"
           lbl_lockstat.Tag = "LOCK"
        Else
           lbl_lockstat = vbNullString
           lbl_lockstat.Tag = vbNullString
        End If
     End If
  End If
  If ExecutionState = EXEC_SELECT And Levels.ListIndex > -1 Then
     PreviewLevel Val(Levels.List(Levels.ListIndex))
  End If
End Sub

Public Sub Levels_DblClick()
  If Len(lbl_lockstat.Tag) = 0 Then
     ExecutionState = EXEC_GAME
  End If
End Sub

Private Sub levname_GotFocus()
  With levname
    .SelStart = 0
    .SelLength = Len(.Text)
  End With
End Sub

Private Sub levPacks_Click()
  If EngineSelect Then Exit Sub
  Preview.Cls
  Preview.Tag = vbNullString
  LoadLevelNames levPacks.List(levPacks.ListIndex)
  If Not arCheats(2) Then
     With levPacks
         If .List(.ListIndex) = "STANDARD" Then
            .Tag = "USE_PLAYER"
         Else
            .Tag = "FREE"
         End If
     End With
  End If
  UpdatePlayerStats
End Sub

Private Sub chk_settings_Click(Index As Integer)
  If EngineSelect = True Then Exit Sub
  With chk_settings(Index)
    .Tag = IIf(.Tag = "0", "1", "0")
    .Cls
    TransparentBlt .hdc, 0, 0, 16, 16, picOnOff.hdc, Val(.Tag) * 16, 0, 16, 16, vbBlack
  End With
  Select Case Index
  Case 0
    SoundEnabled = Not SoundEnabled
  Case 1
    MusicEnabled = Not MusicEnabled
    If MusicEnabled Then
       'play menu song
       PlayMusic "SONG01"
    Else
       StopMusic
    End If
  Case 2
  End Select
End Sub

Public Sub chk_settings_Show(Index As Integer)
  With chk_settings(Index)
    .Cls
    TransparentBlt .hdc, 0, 0, 16, 16, picOnOff.hdc, Val(.Tag) * 16, 0, 16, 16, vbBlack
  End With
End Sub

Private Sub but_edit_ok_Click()
  If DoEditPrivilegesCheck() Then
     PlaySound "BUTTON"
                 
     'set back to "EXISTING PACK" mode
     o_edit_pack(0).Value = True
                 
     ExecutionState = EXEC_GAME + EXEC_EDIT
  End If
End Sub

Private Sub but_settings_ok_Click()
  eSettings.Visible = False
End Sub

Private Sub but_error_ok_Click()
  eLayer.Visible = False
End Sub

Private Sub but_loadlevel_ok_Click()
  If lstBin.ListIndex > -1 Then
     eLoad.Visible = False
  End If
End Sub

Private Sub but_player_ok_Click()
   PlaySound "BUTTON"
   With levPacks
      If .List(.ListIndex) = "STANDARD" Then
         .Tag = "USE_PLAYER"
      Else
         .Tag = "FREE"
      End If
   End With
   sLayer.Visible = False
   UpdatePlayerStats
End Sub

Public Sub UpdatePlayerStats()
   If arCheats(2) Then
      lbl_player_sel.Enabled = False
      Exit Sub
   End If
   With lstPlayers
     If .ListIndex > -1 And levPacks.Tag = "USE_PLAYER" Then
        lbl_player = .List(.ListIndex)
        lbl_player.Tag = CurLevel(.ListIndex)
        lbl_levstat = "LEVELS SOLVED: " & Val(lbl_player.Tag) - 1
        lbl_player_sel.Enabled = True
     Else
        If levPacks.Tag = "FREE" Then
           lbl_player = "EVERYONE"
           lbl_levstat = "FREE FOR EVERYONE"
           lbl_player_sel.Enabled = False
        Else
           lbl_player = vbNullString
           lbl_levstat = vbNullString
           lbl_player_sel.Enabled = True
        End If
        lbl_player.Tag = 0
     End If
   End With
   Levels_Click
End Sub

Private Sub but_loadlevel_cancel_Click()
  lstBin.ListIndex = -1
  eLoad.Visible = False
End Sub

Private Sub but_newlevel_Click()
   If DoQuestion("Start with a new empty levelfield?", , vbYesNo) = vbYes Then
      With frmScreen
         eLoad.Visible = False
         LockMenuMouse = False
         LockGameMouse = False
         GameLoop .levPacks.List(.levPacks.ListIndex)
      End With
   End If
End Sub

Private Sub but_question_Click(Index As Integer)
   eQuestion.Tag = but_question(Index).Tag
   qLayer.Visible = False
End Sub

Private Sub lstEditPack_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  frmScreen.ed_ver_passw = vbNullString
End Sub

Private Sub mVol_Scroll()
  SetMusicVolume
End Sub

Private Sub o_edit_pack_Click(Index As Integer)
  Dim d1 As Boolean
  Dim d2 As Boolean
  
  If EngineSelect = False Then PlaySound "BUTTON"
                  
  d1 = IIf(Index = 0, True, False)
  d2 = Not d1
  
  lstEditPack.Enabled = d1
  ed_ver_passw.Enabled = d1
  
  ed_packname.Enabled = d2
  ed_password.Enabled = d2
  ed_retype.Enabled = d2

End Sub

Private Sub player_add_Click()
  Dim Player As String
  
  PlaySound "BUTTON"
  
  '[UPDATE 2.0.1] Now uses the internal input dialog window instead of InputBox
  Player = UCase(DoInput("ENTER NEW PLAYER NAME", "NEW PLAYER"))
  
  If Len(Player) > 0 Then
     DoAddNewPlayer Player
  End If
  
End Sub

Private Sub player_remove_Click()
    
  With lstPlayers
     If .ListIndex > -1 Then
        PlaySound "BUTTON"
        
        DoRemovePlayer .List(.ListIndex)
     End If
  End With
    
End Sub

Private Sub lbl_player_sel_Click()
  player_sel_Click
End Sub

Private Sub player_sel_Click()
  If lbl_player_sel.Enabled = True Then
     PlaySound "BUTTON"
     sLayer.Visible = True
  End If
End Sub

'[NEW 2.0.1] restore original screenmode when switching from app
Private Sub tmr_LostFocus_Timer()
  If cmbScreen.ListIndex = 1 Then
     If GetActiveWindow() <> ActiveHwnd Then
        WindowState = vbMinimized
        RestoreScreen
        RestoreOnActivate = True
        ActiveHwnd = GetActiveWindow()
     End If
  End If
End Sub

Private Sub txt_input_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
     KeyAscii = 0
     but_input_ok_Click
  End If
End Sub

Private Sub but_input_ok_Click()
  If Len(Trim(txt_input)) > 0 Then
    iLayer.Visible = False
  End If
End Sub

Private Sub but_input_cancel_Click()
  txt_input = vbNullString
  iLayer.Visible = False
End Sub
Private Sub obj_col_Click(Index As Integer)
  If EngineSelect Then Exit Sub
  ApplyColor Index
  eSettings.SetFocus
End Sub

Private Sub obj_props_Click(Index As Integer)
  If EngineSelect Then Exit Sub
  ApplySettings
End Sub

'[ADDON 2.0.1] LightQ Update Feature -> Download update list
Private Sub but_chkupdate_Click()
  Dim Data      As String
  Dim Updates() As String
  Dim Name      As String
  Dim version   As String
  Dim Info      As String
  lblUpdates = "CHECKING ..."
  HttpEx.AddPostData "name=" & App.ProductName & "&version=" & App.Major & "." & App.Minor & "." & App.Revision
  Data = HttpEx.FetchURL(UPDATE_SERVER & "\update.php", , False)
  If Len(Data) Then
     Updates() = Split(Data, "[UPDATE]", , vbTextCompare)
     lstUpdates.Clear
     For i = 1 To UBound(Updates)
        Name = HttpEx.ExtractHeader("<Name>", Updates(i))
        version = HttpEx.ExtractHeader("<Version>", Updates(i))
        Info = HttpEx.ExtractHeader("<Info>", Updates(i))
        lstUpdates.AddItem Name & " Update " & version
     Next i
     lblUpdates = IIf(UBound(Updates) = 0, "THIS IS THE LATEST VERSION.", UBound(Updates) & " UPDATE(S) FOUND.")
  Else
     lblUpdates = "ERROR DOWNLOADING LIST ..."
  End If
End Sub

'[ADDON 2.0.1] LightQ Update Feature -> Download update
Private Sub but_downupdate_Click()
  Dim Data      As String
  Dim Url       As String
  Dim S         As Long
  If lstUpdates.ListIndex = -1 Then Exit Sub
  lblUpdates = "DOWNLOADING ..."
  Url = lstUpdates.List(lstUpdates.ListIndex) & ".exe"
  Data = HttpEx.FetchURL(UPDATE_SERVER & "\" & Url)
  If Len(Data) = Val(HttpEx.ExtractHeader("Content-Length", HttpEx.GetHeader)) Then
     lblUpdates = "SAVING ..."
     Open GetWindowsTempFolder() & Url & ".exe" For Binary As #1
       Put #1, , Data
     Close #1
     lblUpdates = vbNullString
     If DoQuestion("Update downloaded succesfully. For the update to be installed, LightQ needs to shutdown, do you want this?", , vbYesNo) = vbYes Then
        ShellExecute hwnd, vbNullString, GetWindowsTempFolder() & Url & ".exe", vbNullString, vbNullString, SW_SHOWDEFAULT
        Call Form_Unload(0)
     End If
  Else
     lblUpdates = vbNullString
     ShowError "Downloading update failed. Possible reasons are server down, not existing file or bad data transfer."
  End If
End Sub

Private Sub edit_obj_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  oYSelected = YSelected
  oXSelected = XSelected
  YSelected = Y \ SP_H
  XSelected = X \ SP_W
  eObjects.Visible = False
End Sub

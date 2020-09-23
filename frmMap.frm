VERSION 5.00
Begin VB.Form frmMap 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Height Map"
   ClientHeight    =   5070
   ClientLeft      =   735
   ClientTop       =   210
   ClientWidth     =   3135
   FillColor       =   &H000000FF&
   ForeColor       =   &H000000FF&
   Icon            =   "frmMap.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   338
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   209
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Caption         =   "Go Thru TP"
      ForeColor       =   &H000000FF&
      Height          =   975
      Left            =   0
      TabIndex        =   10
      Top             =   4080
      Width           =   3135
      Begin VB.TextBox txtTile 
         BackColor       =   &H00000000&
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   2895
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Go"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   620
         Width           =   2895
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H000000FF&
         Height          =   255
         Left            =   120
         Top             =   600
         Width           =   2895
      End
   End
   Begin VB.Frame frmRoom 
      BackColor       =   &H00000000&
      Caption         =   "Floormap"
      ForeColor       =   &H000000FF&
      Height          =   3975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3135
      Begin VB.Frame frmModelD 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   2895
         Left            =   300
         TabIndex        =   9
         Top             =   360
         Visible         =   0   'False
         Width           =   2535
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   132
            Left            =   0
            Picture         =   "frmMap.frx":08A6
            Top             =   0
            Visible         =   0   'False
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   127
            Left            =   1800
            Picture         =   "frmMap.frx":0DB0
            Top             =   240
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   126
            Left            =   1920
            Picture         =   "frmMap.frx":12BA
            Top             =   360
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   83
            Left            =   2040
            Picture         =   "frmMap.frx":17C4
            Top             =   480
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   82
            Left            =   2160
            Picture         =   "frmMap.frx":1CCE
            Top             =   600
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   81
            Left            =   2280
            Picture         =   "frmMap.frx":21D8
            Top             =   720
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   79
            Left            =   2400
            Picture         =   "frmMap.frx":26E2
            Top             =   840
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   76
            Left            =   1080
            Picture         =   "frmMap.frx":2BEC
            Top             =   720
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   75
            Left            =   1200
            Picture         =   "frmMap.frx":30F6
            Top             =   600
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   74
            Left            =   1320
            Picture         =   "frmMap.frx":3600
            Top             =   480
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   73
            Left            =   1440
            Picture         =   "frmMap.frx":3B0A
            Top             =   360
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   72
            Left            =   1560
            Picture         =   "frmMap.frx":4014
            Top             =   240
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   71
            Left            =   1680
            Picture         =   "frmMap.frx":451E
            Top             =   120
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   70
            Left            =   0
            Picture         =   "frmMap.frx":4A28
            Top             =   1800
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   61
            Left            =   120
            Picture         =   "frmMap.frx":4F32
            Top             =   1680
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   60
            Left            =   240
            Picture         =   "frmMap.frx":543C
            Top             =   1560
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   53
            Left            =   360
            Picture         =   "frmMap.frx":5946
            Top             =   1440
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   52
            Left            =   480
            Picture         =   "frmMap.frx":5E50
            Top             =   1320
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   51
            Left            =   600
            Picture         =   "frmMap.frx":635A
            Top             =   1200
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   50
            Left            =   720
            Picture         =   "frmMap.frx":6864
            Top             =   1080
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   49
            Left            =   840
            Picture         =   "frmMap.frx":6D6E
            Top             =   960
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   48
            Left            =   960
            Picture         =   "frmMap.frx":7278
            Top             =   840
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   0
            Left            =   1680
            Picture         =   "frmMap.frx":7782
            Top             =   240
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   1
            Left            =   1800
            Picture         =   "frmMap.frx":7C8C
            Top             =   360
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   2
            Left            =   360
            Picture         =   "frmMap.frx":8196
            Top             =   1560
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   3
            Left            =   240
            Picture         =   "frmMap.frx":86A0
            Top             =   1680
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   4
            Left            =   120
            Picture         =   "frmMap.frx":8BAA
            Top             =   1800
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   5
            Left            =   0
            Picture         =   "frmMap.frx":90B4
            Top             =   1920
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   6
            Left            =   1560
            Picture         =   "frmMap.frx":95BE
            Top             =   600
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   7
            Left            =   1680
            Picture         =   "frmMap.frx":9AC8
            Top             =   480
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   8
            Left            =   1800
            Picture         =   "frmMap.frx":9FD2
            Top             =   600
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   9
            Left            =   1920
            Picture         =   "frmMap.frx":A4DC
            Top             =   720
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   10
            Left            =   1440
            Picture         =   "frmMap.frx":A9E6
            Top             =   720
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   11
            Left            =   1320
            Picture         =   "frmMap.frx":AEF0
            Top             =   840
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   12
            Left            =   2040
            Picture         =   "frmMap.frx":B3FA
            Top             =   840
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   13
            Left            =   2160
            Picture         =   "frmMap.frx":B904
            Top             =   960
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   14
            Left            =   1200
            Picture         =   "frmMap.frx":BE0E
            Top             =   960
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   15
            Left            =   1080
            Picture         =   "frmMap.frx":C318
            Top             =   1080
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   16
            Left            =   960
            Picture         =   "frmMap.frx":C822
            Top             =   1200
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   17
            Left            =   840
            Picture         =   "frmMap.frx":CD2C
            Top             =   1320
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   18
            Left            =   720
            Picture         =   "frmMap.frx":D236
            Top             =   1440
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   19
            Left            =   600
            Picture         =   "frmMap.frx":D740
            Top             =   1560
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   20
            Left            =   480
            Picture         =   "frmMap.frx":DC4A
            Top             =   1680
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   21
            Left            =   360
            Picture         =   "frmMap.frx":E154
            Top             =   1800
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   22
            Left            =   240
            Picture         =   "frmMap.frx":E65E
            Top             =   1920
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   23
            Left            =   120
            Picture         =   "frmMap.frx":EB68
            Top             =   2040
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   24
            Left            =   240
            Picture         =   "frmMap.frx":F072
            Top             =   2160
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   25
            Left            =   360
            Picture         =   "frmMap.frx":F57C
            Top             =   2280
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   26
            Left            =   480
            Picture         =   "frmMap.frx":FA86
            Top             =   2400
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   27
            Left            =   600
            Picture         =   "frmMap.frx":FF90
            Top             =   2520
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   28
            Left            =   2280
            Picture         =   "frmMap.frx":1049A
            Top             =   1080
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   29
            Left            =   2160
            Picture         =   "frmMap.frx":109A4
            Top             =   1200
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   30
            Left            =   2040
            Picture         =   "frmMap.frx":10EAE
            Top             =   1320
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   31
            Left            =   1920
            Picture         =   "frmMap.frx":113B8
            Top             =   1440
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   32
            Left            =   1800
            Picture         =   "frmMap.frx":118C2
            Top             =   1560
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   33
            Left            =   1680
            Picture         =   "frmMap.frx":11DCC
            Top             =   1680
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   34
            Left            =   1560
            Picture         =   "frmMap.frx":122D6
            Top             =   1800
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   35
            Left            =   1440
            Picture         =   "frmMap.frx":127E0
            Top             =   1920
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   36
            Left            =   1320
            Picture         =   "frmMap.frx":12CEA
            Top             =   2040
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   37
            Left            =   1200
            Picture         =   "frmMap.frx":131F4
            Top             =   2160
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   38
            Left            =   1080
            Picture         =   "frmMap.frx":136FE
            Top             =   2280
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   39
            Left            =   960
            Picture         =   "frmMap.frx":13C08
            Top             =   2400
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   40
            Left            =   840
            Picture         =   "frmMap.frx":14112
            Top             =   2520
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   41
            Left            =   720
            Picture         =   "frmMap.frx":1461C
            Top             =   2640
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   42
            Left            =   1560
            Picture         =   "frmMap.frx":14B26
            Top             =   360
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   43
            Left            =   600
            Picture         =   "frmMap.frx":15030
            Top             =   2280
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   44
            Left            =   720
            Picture         =   "frmMap.frx":1553A
            Top             =   2400
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   45
            Left            =   1680
            Picture         =   "frmMap.frx":15A44
            Top             =   720
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   46
            Left            =   2040
            Picture         =   "frmMap.frx":15F4E
            Top             =   1080
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   47
            Left            =   1920
            Picture         =   "frmMap.frx":16458
            Top             =   960
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   54
            Left            =   600
            Picture         =   "frmMap.frx":16962
            Top             =   1320
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   55
            Left            =   720
            Picture         =   "frmMap.frx":16E6C
            Top             =   1200
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   56
            Left            =   1920
            Picture         =   "frmMap.frx":17376
            Top             =   480
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   57
            Left            =   2040
            Picture         =   "frmMap.frx":17880
            Top             =   600
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   58
            Left            =   2160
            Picture         =   "frmMap.frx":17D8A
            Top             =   720
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   59
            Left            =   2280
            Picture         =   "frmMap.frx":18294
            Top             =   840
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   62
            Left            =   2400
            Picture         =   "frmMap.frx":1879E
            Top             =   960
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   63
            Left            =   840
            Picture         =   "frmMap.frx":18CA8
            Top             =   1080
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   64
            Left            =   960
            Picture         =   "frmMap.frx":191B2
            Top             =   960
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   65
            Left            =   1080
            Picture         =   "frmMap.frx":196BC
            Top             =   840
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   66
            Left            =   1200
            Picture         =   "frmMap.frx":19BC6
            Top             =   720
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   67
            Left            =   1320
            Picture         =   "frmMap.frx":1A0D0
            Top             =   600
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   68
            Left            =   1440
            Picture         =   "frmMap.frx":1A5DA
            Top             =   480
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   69
            Left            =   480
            Picture         =   "frmMap.frx":1AAE4
            Top             =   1440
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   80
            Left            =   1800
            Picture         =   "frmMap.frx":1AFEE
            Top             =   840
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   84
            Left            =   840
            Picture         =   "frmMap.frx":1B4F8
            Top             =   1560
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   85
            Left            =   1200
            Picture         =   "frmMap.frx":1BA02
            Top             =   1920
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   86
            Left            =   720
            Picture         =   "frmMap.frx":1BF0C
            Top             =   1680
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   87
            Left            =   1320
            Picture         =   "frmMap.frx":1C416
            Top             =   1080
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   88
            Left            =   1440
            Picture         =   "frmMap.frx":1C920
            Top             =   1200
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   89
            Left            =   840
            Picture         =   "frmMap.frx":1CE2A
            Top             =   1800
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   90
            Left            =   1560
            Picture         =   "frmMap.frx":1D334
            Top             =   1320
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   91
            Left            =   960
            Picture         =   "frmMap.frx":1D83E
            Top             =   1920
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   92
            Left            =   1680
            Picture         =   "frmMap.frx":1DD48
            Top             =   1440
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   93
            Left            =   1080
            Picture         =   "frmMap.frx":1E252
            Top             =   2040
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   94
            Left            =   600
            Picture         =   "frmMap.frx":1E75C
            Top             =   1800
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   95
            Left            =   720
            Picture         =   "frmMap.frx":1EC66
            Top             =   1920
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   96
            Left            =   1440
            Picture         =   "frmMap.frx":1F170
            Top             =   960
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   97
            Left            =   840
            Picture         =   "frmMap.frx":1F67A
            Top             =   2040
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   98
            Left            =   1560
            Picture         =   "frmMap.frx":1FB84
            Top             =   1080
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   99
            Left            =   960
            Picture         =   "frmMap.frx":2008E
            Top             =   2160
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   100
            Left            =   1680
            Picture         =   "frmMap.frx":20598
            Top             =   1200
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   101
            Left            =   480
            Picture         =   "frmMap.frx":20AA2
            Top             =   1920
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   102
            Left            =   1800
            Picture         =   "frmMap.frx":20FAC
            Top             =   1320
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   103
            Left            =   600
            Picture         =   "frmMap.frx":214B6
            Top             =   2040
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   104
            Left            =   1560
            Picture         =   "frmMap.frx":219C0
            Top             =   840
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   105
            Left            =   720
            Picture         =   "frmMap.frx":21ECA
            Top             =   2160
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   106
            Left            =   1680
            Picture         =   "frmMap.frx":223D4
            Top             =   960
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   107
            Left            =   840
            Picture         =   "frmMap.frx":228DE
            Top             =   2280
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   108
            Left            =   1800
            Picture         =   "frmMap.frx":22DE8
            Top             =   1080
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   109
            Left            =   360
            Picture         =   "frmMap.frx":232F2
            Top             =   2040
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   110
            Left            =   1920
            Picture         =   "frmMap.frx":237FC
            Top             =   1200
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   111
            Left            =   480
            Picture         =   "frmMap.frx":23D06
            Top             =   2160
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   112
            Left            =   960
            Picture         =   "frmMap.frx":24210
            Top             =   1680
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   113
            Left            =   1200
            Picture         =   "frmMap.frx":2471A
            Top             =   1200
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   114
            Left            =   1080
            Picture         =   "frmMap.frx":24C24
            Top             =   1320
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   115
            Left            =   960
            Picture         =   "frmMap.frx":2512E
            Top             =   1440
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   116
            Left            =   1560
            Picture         =   "frmMap.frx":25638
            Top             =   1560
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   117
            Left            =   1440
            Picture         =   "frmMap.frx":25B42
            Top             =   1680
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   118
            Left            =   1320
            Picture         =   "frmMap.frx":2604C
            Top             =   1800
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   119
            Left            =   1200
            Picture         =   "frmMap.frx":26556
            Top             =   1440
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   120
            Left            =   1080
            Picture         =   "frmMap.frx":26A60
            Top             =   1560
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   121
            Left            =   1440
            Picture         =   "frmMap.frx":26F6A
            Top             =   1440
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   122
            Left            =   1320
            Picture         =   "frmMap.frx":27474
            Top             =   1320
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   123
            Left            =   1080
            Picture         =   "frmMap.frx":2797E
            Top             =   1800
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   124
            Left            =   1320
            Picture         =   "frmMap.frx":27E88
            Top             =   1560
            Width           =   120
         End
         Begin VB.Image mapTileD 
            Height          =   240
            Index           =   125
            Left            =   1200
            Picture         =   "frmMap.frx":28392
            Top             =   1680
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   131
            Left            =   240
            Picture         =   "frmMap.frx":2889C
            Top             =   0
            Visible         =   0   'False
            Width           =   120
         End
      End
      Begin VB.Frame frmModelA 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   2895
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Visible         =   0   'False
         Width           =   2655
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   129
            Left            =   120
            Picture         =   "frmMap.frx":28DA6
            Top             =   0
            Visible         =   0   'False
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   128
            Left            =   0
            Picture         =   "frmMap.frx":292B0
            Top             =   0
            Visible         =   0   'False
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   148
            Left            =   1320
            Picture         =   "frmMap.frx":297BA
            Top             =   1560
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   147
            Left            =   1200
            Picture         =   "frmMap.frx":29CC4
            Top             =   1440
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   146
            Left            =   1440
            Picture         =   "frmMap.frx":2A1CE
            Top             =   1440
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   145
            Left            =   1320
            Picture         =   "frmMap.frx":2A6D8
            Top             =   1320
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   144
            Left            =   1560
            Picture         =   "frmMap.frx":2ABE2
            Top             =   1320
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   143
            Left            =   1440
            Picture         =   "frmMap.frx":2B0EC
            Top             =   1200
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   141
            Left            =   1680
            Picture         =   "frmMap.frx":2B5F6
            Top             =   1440
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   140
            Left            =   1680
            Picture         =   "frmMap.frx":2BB00
            Top             =   1200
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   139
            Left            =   1200
            Picture         =   "frmMap.frx":2C00A
            Top             =   1680
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   138
            Left            =   1080
            Picture         =   "frmMap.frx":2C514
            Top             =   1560
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   137
            Left            =   840
            Picture         =   "frmMap.frx":2CA1E
            Top             =   1800
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   136
            Left            =   960
            Picture         =   "frmMap.frx":2CF28
            Top             =   1680
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   135
            Left            =   960
            Picture         =   "frmMap.frx":2D432
            Top             =   1920
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   134
            Left            =   1200
            Picture         =   "frmMap.frx":2D93C
            Top             =   1920
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   133
            Left            =   1560
            Picture         =   "frmMap.frx":2DE46
            Top             =   1080
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   132
            Left            =   1080
            Picture         =   "frmMap.frx":2E350
            Top             =   1800
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   131
            Left            =   1320
            Picture         =   "frmMap.frx":2E85A
            Top             =   1800
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   130
            Left            =   1560
            Picture         =   "frmMap.frx":2ED64
            Top             =   1560
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   129
            Left            =   960
            Picture         =   "frmMap.frx":2F26E
            Top             =   2400
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   128
            Left            =   840
            Picture         =   "frmMap.frx":2F778
            Top             =   2280
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   127
            Left            =   720
            Picture         =   "frmMap.frx":2FC82
            Top             =   2160
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   126
            Left            =   600
            Picture         =   "frmMap.frx":3018C
            Top             =   2040
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   125
            Left            =   480
            Picture         =   "frmMap.frx":30696
            Top             =   1920
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   124
            Left            =   360
            Picture         =   "frmMap.frx":30BA0
            Top             =   1800
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   123
            Left            =   1080
            Picture         =   "frmMap.frx":310AA
            Top             =   2280
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   122
            Left            =   1200
            Picture         =   "frmMap.frx":315B4
            Top             =   2160
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   121
            Left            =   1320
            Picture         =   "frmMap.frx":31ABE
            Top             =   2040
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   120
            Left            =   1440
            Picture         =   "frmMap.frx":31FC8
            Top             =   1920
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   119
            Left            =   1560
            Picture         =   "frmMap.frx":324D2
            Top             =   1800
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   118
            Left            =   1680
            Picture         =   "frmMap.frx":329DC
            Top             =   1680
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   117
            Left            =   1800
            Picture         =   "frmMap.frx":32EE6
            Top             =   1560
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   116
            Left            =   1920
            Picture         =   "frmMap.frx":333F0
            Top             =   1440
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   115
            Left            =   2040
            Picture         =   "frmMap.frx":338FA
            Top             =   1320
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   114
            Left            =   2160
            Picture         =   "frmMap.frx":33E04
            Top             =   1200
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   113
            Left            =   2040
            Picture         =   "frmMap.frx":3430E
            Top             =   1080
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   112
            Left            =   1920
            Picture         =   "frmMap.frx":34818
            Top             =   960
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   111
            Left            =   1800
            Picture         =   "frmMap.frx":34D22
            Top             =   840
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   110
            Left            =   1680
            Picture         =   "frmMap.frx":3522C
            Top             =   720
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   109
            Left            =   1560
            Picture         =   "frmMap.frx":35736
            Top             =   600
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   108
            Left            =   1440
            Picture         =   "frmMap.frx":35C40
            Top             =   720
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   107
            Left            =   1320
            Picture         =   "frmMap.frx":3614A
            Top             =   840
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   106
            Left            =   1200
            Picture         =   "frmMap.frx":36654
            Top             =   960
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   105
            Left            =   1080
            Picture         =   "frmMap.frx":36B5E
            Top             =   1080
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   104
            Left            =   960
            Picture         =   "frmMap.frx":37068
            Top             =   1200
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   103
            Left            =   840
            Picture         =   "frmMap.frx":37572
            Top             =   1320
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   102
            Left            =   720
            Picture         =   "frmMap.frx":37A7C
            Top             =   1440
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   101
            Left            =   600
            Picture         =   "frmMap.frx":37F86
            Top             =   1560
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   100
            Left            =   480
            Picture         =   "frmMap.frx":38490
            Top             =   1680
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   99
            Left            =   600
            Picture         =   "frmMap.frx":3899A
            Top             =   1800
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   98
            Left            =   720
            Picture         =   "frmMap.frx":38EA4
            Top             =   1920
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   97
            Left            =   840
            Picture         =   "frmMap.frx":393AE
            Top             =   2040
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   96
            Left            =   960
            Picture         =   "frmMap.frx":398B8
            Top             =   2160
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   95
            Left            =   1560
            Picture         =   "frmMap.frx":39DC2
            Top             =   840
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   94
            Left            =   1680
            Picture         =   "frmMap.frx":3A2CC
            Top             =   960
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   93
            Left            =   1800
            Picture         =   "frmMap.frx":3A7D6
            Top             =   1080
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   92
            Left            =   1920
            Picture         =   "frmMap.frx":3ACE0
            Top             =   1200
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   91
            Left            =   1800
            Picture         =   "frmMap.frx":3B1EA
            Top             =   1320
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   90
            Left            =   1440
            Picture         =   "frmMap.frx":3B6F4
            Top             =   960
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   89
            Left            =   720
            Picture         =   "frmMap.frx":3BBFE
            Top             =   1680
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   88
            Left            =   1080
            Picture         =   "frmMap.frx":3C108
            Top             =   2040
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   87
            Left            =   1080
            Picture         =   "frmMap.frx":3C612
            Top             =   1320
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   86
            Left            =   1440
            Picture         =   "frmMap.frx":3CB1C
            Top             =   1680
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   85
            Left            =   840
            Picture         =   "frmMap.frx":3D026
            Top             =   1560
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   84
            Left            =   960
            Picture         =   "frmMap.frx":3D530
            Top             =   1440
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   83
            Left            =   1200
            Picture         =   "frmMap.frx":3DA3A
            Top             =   1200
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   82
            Left            =   1320
            Picture         =   "frmMap.frx":3DF44
            Top             =   1080
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   81
            Left            =   2280
            Picture         =   "frmMap.frx":3E44E
            Top             =   720
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   80
            Left            =   2400
            Picture         =   "frmMap.frx":3E958
            Top             =   840
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   79
            Left            =   2520
            Picture         =   "frmMap.frx":3EE62
            Top             =   960
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   78
            Left            =   1320
            Picture         =   "frmMap.frx":3F36C
            Top             =   240
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   77
            Left            =   1440
            Picture         =   "frmMap.frx":3F876
            Top             =   120
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   76
            Left            =   1560
            Picture         =   "frmMap.frx":3FD80
            Top             =   0
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   75
            Left            =   1680
            Picture         =   "frmMap.frx":4028A
            Top             =   120
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   74
            Left            =   1800
            Picture         =   "frmMap.frx":40794
            Top             =   240
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   73
            Left            =   1920
            Picture         =   "frmMap.frx":40C9E
            Top             =   360
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   72
            Left            =   2040
            Picture         =   "frmMap.frx":411A8
            Top             =   480
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   71
            Left            =   2160
            Picture         =   "frmMap.frx":416B2
            Top             =   600
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   70
            Left            =   120
            Picture         =   "frmMap.frx":41BBC
            Top             =   1440
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   69
            Left            =   240
            Picture         =   "frmMap.frx":420C6
            Top             =   1320
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   68
            Left            =   360
            Picture         =   "frmMap.frx":425D0
            Top             =   1200
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   67
            Left            =   480
            Picture         =   "frmMap.frx":42ADA
            Top             =   1080
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   66
            Left            =   600
            Picture         =   "frmMap.frx":42FE4
            Top             =   960
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   65
            Left            =   720
            Picture         =   "frmMap.frx":434EE
            Top             =   840
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   64
            Left            =   840
            Picture         =   "frmMap.frx":439F8
            Top             =   720
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   63
            Left            =   960
            Picture         =   "frmMap.frx":43F02
            Top             =   600
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   62
            Left            =   1080
            Picture         =   "frmMap.frx":4440C
            Top             =   480
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   61
            Left            =   2040
            Picture         =   "frmMap.frx":44916
            Top             =   600
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   60
            Left            =   1920
            Picture         =   "frmMap.frx":44E20
            Top             =   480
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   59
            Left            =   1800
            Picture         =   "frmMap.frx":4532A
            Top             =   360
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   58
            Left            =   1680
            Picture         =   "frmMap.frx":45834
            Top             =   240
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   57
            Left            =   1560
            Picture         =   "frmMap.frx":45D3E
            Top             =   120
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   56
            Left            =   1200
            Picture         =   "frmMap.frx":46248
            Top             =   360
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   55
            Left            =   2520
            Picture         =   "frmMap.frx":46752
            Top             =   1080
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   54
            Left            =   2400
            Picture         =   "frmMap.frx":46C5C
            Top             =   960
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   53
            Left            =   2280
            Picture         =   "frmMap.frx":47166
            Top             =   840
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   52
            Left            =   2160
            Picture         =   "frmMap.frx":47670
            Top             =   720
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   51
            Left            =   720
            Picture         =   "frmMap.frx":47B7A
            Top             =   960
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   50
            Left            =   840
            Picture         =   "frmMap.frx":48084
            Top             =   840
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   49
            Left            =   960
            Picture         =   "frmMap.frx":4858E
            Top             =   720
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   48
            Left            =   1080
            Picture         =   "frmMap.frx":48A98
            Top             =   600
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   47
            Left            =   1200
            Picture         =   "frmMap.frx":48FA2
            Top             =   480
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   46
            Left            =   1440
            Picture         =   "frmMap.frx":494AC
            Top             =   240
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   45
            Left            =   1320
            Picture         =   "frmMap.frx":499B6
            Top             =   360
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   44
            Left            =   0
            Picture         =   "frmMap.frx":49EC0
            Top             =   1560
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   43
            Left            =   600
            Picture         =   "frmMap.frx":4A3CA
            Top             =   1320
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   42
            Left            =   720
            Picture         =   "frmMap.frx":4A8D4
            Top             =   1200
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   41
            Left            =   840
            Picture         =   "frmMap.frx":4ADDE
            Top             =   1080
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   40
            Left            =   960
            Picture         =   "frmMap.frx":4B2E8
            Top             =   960
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   39
            Left            =   1080
            Picture         =   "frmMap.frx":4B7F2
            Top             =   840
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   38
            Left            =   1200
            Picture         =   "frmMap.frx":4BCFC
            Top             =   720
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   37
            Left            =   2280
            Picture         =   "frmMap.frx":4C206
            Top             =   1080
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   36
            Left            =   2160
            Picture         =   "frmMap.frx":4C710
            Top             =   960
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   35
            Left            =   2040
            Picture         =   "frmMap.frx":4CC1A
            Top             =   840
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   34
            Left            =   1920
            Picture         =   "frmMap.frx":4D124
            Top             =   720
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   33
            Left            =   1800
            Picture         =   "frmMap.frx":4D62E
            Top             =   600
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   32
            Left            =   1680
            Picture         =   "frmMap.frx":4DB38
            Top             =   480
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   31
            Left            =   1320
            Picture         =   "frmMap.frx":4E042
            Top             =   600
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   30
            Left            =   1440
            Picture         =   "frmMap.frx":4E54C
            Top             =   480
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   29
            Left            =   1560
            Picture         =   "frmMap.frx":4EA56
            Top             =   360
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   28
            Left            =   0
            Picture         =   "frmMap.frx":4EF60
            Top             =   1680
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   27
            Left            =   120
            Picture         =   "frmMap.frx":4F46A
            Top             =   1560
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   26
            Left            =   240
            Picture         =   "frmMap.frx":4F974
            Top             =   1440
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   25
            Left            =   360
            Picture         =   "frmMap.frx":4FE7E
            Top             =   1320
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   24
            Left            =   480
            Picture         =   "frmMap.frx":50388
            Top             =   1200
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   23
            Left            =   600
            Picture         =   "frmMap.frx":50892
            Top             =   1080
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   22
            Left            =   840
            Picture         =   "frmMap.frx":50D9C
            Top             =   2520
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   21
            Left            =   720
            Picture         =   "frmMap.frx":512A6
            Top             =   2400
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   20
            Left            =   600
            Picture         =   "frmMap.frx":517B0
            Top             =   2280
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   19
            Left            =   480
            Picture         =   "frmMap.frx":51CBA
            Top             =   2160
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   18
            Left            =   360
            Picture         =   "frmMap.frx":521C4
            Top             =   2040
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   17
            Left            =   240
            Picture         =   "frmMap.frx":526CE
            Top             =   1920
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   16
            Left            =   120
            Picture         =   "frmMap.frx":52BD8
            Top             =   1800
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   15
            Left            =   240
            Picture         =   "frmMap.frx":530E2
            Top             =   1680
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   14
            Left            =   360
            Picture         =   "frmMap.frx":535EC
            Top             =   1560
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   13
            Left            =   480
            Picture         =   "frmMap.frx":53AF6
            Top             =   1440
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   12
            Left            =   960
            Picture         =   "frmMap.frx":54000
            Top             =   2640
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   11
            Left            =   1080
            Picture         =   "frmMap.frx":5450A
            Top             =   2520
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   10
            Left            =   1200
            Picture         =   "frmMap.frx":54A14
            Top             =   2400
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   9
            Left            =   1320
            Picture         =   "frmMap.frx":54F1E
            Top             =   2280
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   8
            Left            =   1440
            Picture         =   "frmMap.frx":55428
            Top             =   2160
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   7
            Left            =   1560
            Picture         =   "frmMap.frx":55932
            Top             =   2040
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   6
            Left            =   1680
            Picture         =   "frmMap.frx":55E3C
            Top             =   1920
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   5
            Left            =   1800
            Picture         =   "frmMap.frx":56346
            Top             =   1800
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   4
            Left            =   1920
            Picture         =   "frmMap.frx":56850
            Top             =   1680
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   3
            Left            =   2040
            Picture         =   "frmMap.frx":56D5A
            Top             =   1560
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   2
            Left            =   2160
            Picture         =   "frmMap.frx":57264
            Top             =   1440
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   1
            Left            =   2280
            Picture         =   "frmMap.frx":5776E
            Top             =   1320
            Width           =   120
         End
         Begin VB.Image mapTileA 
            Height          =   240
            Index           =   0
            Left            =   2400
            Picture         =   "frmMap.frx":57C78
            Top             =   1200
            Width           =   120
         End
      End
      Begin VB.Frame frmModelC 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   1815
         Left            =   800
         TabIndex        =   7
         Top             =   700
         Visible         =   0   'False
         Width           =   1575
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   127
            Left            =   240
            Picture         =   "frmMap.frx":58182
            Top             =   0
            Visible         =   0   'False
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   126
            Left            =   0
            Picture         =   "frmMap.frx":5868C
            Top             =   0
            Visible         =   0   'False
            Width           =   120
         End
         Begin VB.Image mapTileC 
            Height          =   240
            Index           =   61
            Left            =   1080
            Picture         =   "frmMap.frx":58B96
            Top             =   960
            Width           =   120
         End
         Begin VB.Image mapTileC 
            Height          =   240
            Index           =   60
            Left            =   480
            Picture         =   "frmMap.frx":590A0
            Top             =   1080
            Width           =   120
         End
         Begin VB.Image mapTileC 
            Height          =   240
            Index           =   59
            Left            =   600
            Picture         =   "frmMap.frx":595AA
            Top             =   960
            Width           =   120
         End
         Begin VB.Image mapTileC 
            Height          =   240
            Index           =   58
            Left            =   720
            Picture         =   "frmMap.frx":59AB4
            Top             =   840
            Width           =   120
         End
         Begin VB.Image mapTileC 
            Height          =   240
            Index           =   57
            Left            =   840
            Picture         =   "frmMap.frx":59FBE
            Top             =   720
            Width           =   120
         End
         Begin VB.Image mapTileC 
            Height          =   240
            Index           =   56
            Left            =   600
            Picture         =   "frmMap.frx":5A4C8
            Top             =   1200
            Width           =   120
         End
         Begin VB.Image mapTileC 
            Height          =   240
            Index           =   55
            Left            =   720
            Picture         =   "frmMap.frx":5A9D2
            Top             =   1080
            Width           =   120
         End
         Begin VB.Image mapTileC 
            Height          =   240
            Index           =   54
            Left            =   840
            Picture         =   "frmMap.frx":5AEDC
            Top             =   960
            Width           =   120
         End
         Begin VB.Image mapTileC 
            Height          =   240
            Index           =   53
            Left            =   960
            Picture         =   "frmMap.frx":5B3E6
            Top             =   840
            Width           =   120
         End
         Begin VB.Image mapTileC 
            Height          =   240
            Index           =   52
            Left            =   720
            Picture         =   "frmMap.frx":5B8F0
            Top             =   1320
            Width           =   120
         End
         Begin VB.Image mapTileC 
            Height          =   240
            Index           =   51
            Left            =   840
            Picture         =   "frmMap.frx":5BDFA
            Top             =   1200
            Width           =   120
         End
         Begin VB.Image mapTileC 
            Height          =   240
            Index           =   50
            Left            =   960
            Picture         =   "frmMap.frx":5C304
            Top             =   1080
            Width           =   120
         End
         Begin VB.Image mapTileC 
            Height          =   240
            Index           =   49
            Left            =   240
            Picture         =   "frmMap.frx":5C80E
            Top             =   480
            Width           =   120
         End
         Begin VB.Image mapTileC 
            Height          =   240
            Index           =   48
            Left            =   360
            Picture         =   "frmMap.frx":5CD18
            Top             =   360
            Width           =   120
         End
         Begin VB.Image mapTileC 
            Height          =   240
            Index           =   47
            Left            =   480
            Picture         =   "frmMap.frx":5D222
            Top             =   240
            Width           =   120
         End
         Begin VB.Image mapTileC 
            Height          =   240
            Index           =   46
            Left            =   600
            Picture         =   "frmMap.frx":5D72C
            Top             =   120
            Width           =   120
         End
         Begin VB.Image mapTileC 
            Height          =   240
            Index           =   45
            Left            =   720
            Picture         =   "frmMap.frx":5DC36
            Top             =   0
            Width           =   120
         End
         Begin VB.Image mapTileC 
            Height          =   240
            Index           =   44
            Left            =   840
            Picture         =   "frmMap.frx":5E140
            Top             =   120
            Width           =   120
         End
         Begin VB.Image mapTileC 
            Height          =   240
            Index           =   43
            Left            =   960
            Picture         =   "frmMap.frx":5E64A
            Top             =   240
            Width           =   120
         End
         Begin VB.Image mapTileC 
            Height          =   240
            Index           =   42
            Left            =   1080
            Picture         =   "frmMap.frx":5EB54
            Top             =   360
            Width           =   120
         End
         Begin VB.Image mapTileC 
            Height          =   240
            Index           =   41
            Left            =   1200
            Picture         =   "frmMap.frx":5F05E
            Top             =   480
            Width           =   120
         End
         Begin VB.Image mapTileC 
            Height          =   240
            Index           =   40
            Left            =   1320
            Picture         =   "frmMap.frx":5F568
            Top             =   600
            Width           =   120
         End
         Begin VB.Image mapTileC 
            Height          =   240
            Index           =   39
            Left            =   1440
            Picture         =   "frmMap.frx":5FA72
            Top             =   720
            Width           =   120
         End
         Begin VB.Image mapTileC 
            Height          =   240
            Index           =   38
            Left            =   360
            Picture         =   "frmMap.frx":5FF7C
            Top             =   960
            Width           =   120
         End
         Begin VB.Image mapTileC 
            Height          =   240
            Index           =   37
            Left            =   480
            Picture         =   "frmMap.frx":60486
            Top             =   840
            Width           =   120
         End
         Begin VB.Image mapTileC 
            Height          =   240
            Index           =   36
            Left            =   600
            Picture         =   "frmMap.frx":60990
            Top             =   720
            Width           =   120
         End
         Begin VB.Image mapTileC 
            Height          =   240
            Index           =   35
            Left            =   720
            Picture         =   "frmMap.frx":60E9A
            Top             =   600
            Width           =   120
         End
         Begin VB.Image mapTileC 
            Height          =   240
            Index           =   34
            Left            =   1080
            Picture         =   "frmMap.frx":613A4
            Top             =   480
            Width           =   120
         End
         Begin VB.Image mapTileC 
            Height          =   240
            Index           =   33
            Left            =   1200
            Picture         =   "frmMap.frx":618AE
            Top             =   600
            Width           =   120
         End
         Begin VB.Image mapTileC 
            Height          =   240
            Index           =   32
            Left            =   1320
            Picture         =   "frmMap.frx":61DB8
            Top             =   720
            Width           =   120
         End
         Begin VB.Image mapTileC 
            Height          =   240
            Index           =   31
            Left            =   1440
            Picture         =   "frmMap.frx":622C2
            Top             =   840
            Width           =   120
         End
         Begin VB.Image mapTileC 
            Height          =   240
            Index           =   30
            Left            =   0
            Picture         =   "frmMap.frx":627CC
            Top             =   720
            Width           =   120
         End
         Begin VB.Image mapTileC 
            Height          =   240
            Index           =   29
            Left            =   120
            Picture         =   "frmMap.frx":62CD6
            Top             =   600
            Width           =   120
         End
         Begin VB.Image mapTileC 
            Height          =   240
            Index           =   28
            Left            =   120
            Picture         =   "frmMap.frx":631E0
            Top             =   720
            Width           =   120
         End
         Begin VB.Image mapTileC 
            Height          =   240
            Index           =   27
            Left            =   240
            Picture         =   "frmMap.frx":636EA
            Top             =   600
            Width           =   120
         End
         Begin VB.Image mapTileC 
            Height          =   240
            Index           =   26
            Left            =   360
            Picture         =   "frmMap.frx":63BF4
            Top             =   480
            Width           =   120
         End
         Begin VB.Image mapTileC 
            Height          =   240
            Index           =   25
            Left            =   480
            Picture         =   "frmMap.frx":640FE
            Top             =   360
            Width           =   120
         End
         Begin VB.Image mapTileC 
            Height          =   240
            Index           =   24
            Left            =   600
            Picture         =   "frmMap.frx":64608
            Top             =   240
            Width           =   120
         End
         Begin VB.Image mapTileC 
            Height          =   240
            Index           =   23
            Left            =   720
            Picture         =   "frmMap.frx":64B12
            Top             =   120
            Width           =   120
         End
         Begin VB.Image mapTileC 
            Height          =   240
            Index           =   22
            Left            =   840
            Picture         =   "frmMap.frx":6501C
            Top             =   240
            Width           =   120
         End
         Begin VB.Image mapTileC 
            Height          =   240
            Index           =   21
            Left            =   960
            Picture         =   "frmMap.frx":65526
            Top             =   360
            Width           =   120
         End
         Begin VB.Image mapTileC 
            Height          =   240
            Index           =   20
            Left            =   840
            Picture         =   "frmMap.frx":65A30
            Top             =   480
            Width           =   120
         End
         Begin VB.Image mapTileC 
            Height          =   240
            Index           =   19
            Left            =   720
            Picture         =   "frmMap.frx":65F3A
            Top             =   360
            Width           =   120
         End
         Begin VB.Image mapTileC 
            Height          =   240
            Index           =   18
            Left            =   0
            Picture         =   "frmMap.frx":66444
            Top             =   840
            Width           =   120
         End
         Begin VB.Image mapTileC 
            Height          =   240
            Index           =   17
            Left            =   240
            Picture         =   "frmMap.frx":6694E
            Top             =   840
            Width           =   120
         End
         Begin VB.Image mapTileC 
            Height          =   240
            Index           =   16
            Left            =   360
            Picture         =   "frmMap.frx":66E58
            Top             =   720
            Width           =   120
         End
         Begin VB.Image mapTileC 
            Height          =   240
            Index           =   15
            Left            =   480
            Picture         =   "frmMap.frx":67362
            Top             =   600
            Width           =   120
         End
         Begin VB.Image mapTileC 
            Height          =   240
            Index           =   14
            Left            =   600
            Picture         =   "frmMap.frx":6786C
            Top             =   480
            Width           =   120
         End
         Begin VB.Image mapTileC 
            Height          =   240
            Index           =   13
            Left            =   1200
            Picture         =   "frmMap.frx":67D76
            Top             =   840
            Width           =   120
         End
         Begin VB.Image mapTileC 
            Height          =   240
            Index           =   12
            Left            =   1080
            Picture         =   "frmMap.frx":68280
            Top             =   720
            Width           =   120
         End
         Begin VB.Image mapTileC 
            Height          =   240
            Index           =   11
            Left            =   960
            Picture         =   "frmMap.frx":6878A
            Top             =   600
            Width           =   120
         End
         Begin VB.Image mapTileC 
            Height          =   240
            Index           =   10
            Left            =   120
            Picture         =   "frmMap.frx":68C94
            Top             =   960
            Width           =   120
         End
         Begin VB.Image mapTileC 
            Height          =   240
            Index           =   9
            Left            =   720
            Picture         =   "frmMap.frx":6919E
            Top             =   1560
            Width           =   120
         End
         Begin VB.Image mapTileC 
            Height          =   240
            Index           =   8
            Left            =   840
            Picture         =   "frmMap.frx":696A8
            Top             =   1440
            Width           =   120
         End
         Begin VB.Image mapTileC 
            Height          =   240
            Index           =   7
            Left            =   960
            Picture         =   "frmMap.frx":69BB2
            Top             =   1320
            Width           =   120
         End
         Begin VB.Image mapTileC 
            Height          =   240
            Index           =   6
            Left            =   1080
            Picture         =   "frmMap.frx":6A0BC
            Top             =   1200
            Width           =   120
         End
         Begin VB.Image mapTileC 
            Height          =   240
            Index           =   5
            Left            =   1200
            Picture         =   "frmMap.frx":6A5C6
            Top             =   1080
            Width           =   120
         End
         Begin VB.Image mapTileC 
            Height          =   240
            Index           =   4
            Left            =   1320
            Picture         =   "frmMap.frx":6AAD0
            Top             =   960
            Width           =   120
         End
         Begin VB.Image mapTileC 
            Height          =   240
            Index           =   3
            Left            =   600
            Picture         =   "frmMap.frx":6AFDA
            Top             =   1440
            Width           =   120
         End
         Begin VB.Image mapTileC 
            Height          =   240
            Index           =   2
            Left            =   480
            Picture         =   "frmMap.frx":6B4E4
            Top             =   1320
            Width           =   120
         End
         Begin VB.Image mapTileC 
            Height          =   240
            Index           =   1
            Left            =   360
            Picture         =   "frmMap.frx":6B9EE
            Top             =   1200
            Width           =   120
         End
         Begin VB.Image mapTileC 
            Height          =   240
            Index           =   0
            Left            =   240
            Picture         =   "frmMap.frx":6BEF8
            Top             =   1080
            Width           =   120
         End
      End
      Begin VB.Frame frmModelF 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   2295
         Left            =   300
         TabIndex        =   6
         Top             =   480
         Visible         =   0   'False
         Width           =   2535
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   125
            Left            =   -200
            Picture         =   "frmMap.frx":6C402
            Top             =   0
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   124
            Left            =   -200
            Picture         =   "frmMap.frx":6C90C
            Top             =   0
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   164
            Left            =   2400
            Picture         =   "frmMap.frx":6CE16
            Top             =   720
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   163
            Left            =   720
            Picture         =   "frmMap.frx":6D320
            Top             =   480
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   162
            Left            =   480
            Picture         =   "frmMap.frx":6D82A
            Top             =   240
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   161
            Left            =   840
            Picture         =   "frmMap.frx":6DD34
            Top             =   360
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   160
            Left            =   960
            Picture         =   "frmMap.frx":6E23E
            Top             =   240
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   159
            Left            =   1080
            Picture         =   "frmMap.frx":6E748
            Top             =   120
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   158
            Left            =   1200
            Picture         =   "frmMap.frx":6EC52
            Top             =   0
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   157
            Left            =   1320
            Picture         =   "frmMap.frx":6F15C
            Top             =   120
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   156
            Left            =   1440
            Picture         =   "frmMap.frx":6F666
            Top             =   240
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   155
            Left            =   1560
            Picture         =   "frmMap.frx":6FB70
            Top             =   360
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   154
            Left            =   1680
            Picture         =   "frmMap.frx":7007A
            Top             =   240
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   153
            Left            =   1800
            Picture         =   "frmMap.frx":70584
            Top             =   120
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   152
            Left            =   1920
            Picture         =   "frmMap.frx":70A8E
            Top             =   240
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   151
            Left            =   2040
            Picture         =   "frmMap.frx":70F98
            Top             =   360
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   150
            Left            =   2160
            Picture         =   "frmMap.frx":714A2
            Top             =   480
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   149
            Left            =   2280
            Picture         =   "frmMap.frx":719AC
            Top             =   600
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   142
            Left            =   360
            Picture         =   "frmMap.frx":71EB6
            Top             =   360
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   141
            Left            =   480
            Picture         =   "frmMap.frx":723C0
            Top             =   360
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   140
            Left            =   600
            Picture         =   "frmMap.frx":728CA
            Top             =   360
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   139
            Left            =   600
            Picture         =   "frmMap.frx":72DD4
            Top             =   480
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   138
            Left            =   720
            Picture         =   "frmMap.frx":732DE
            Top             =   600
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   137
            Left            =   840
            Picture         =   "frmMap.frx":737E8
            Top             =   480
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   136
            Left            =   960
            Picture         =   "frmMap.frx":73CF2
            Top             =   360
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   135
            Left            =   1080
            Picture         =   "frmMap.frx":741FC
            Top             =   240
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   134
            Left            =   1200
            Picture         =   "frmMap.frx":74706
            Top             =   120
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   133
            Left            =   1320
            Picture         =   "frmMap.frx":74C10
            Top             =   240
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   132
            Left            =   1440
            Picture         =   "frmMap.frx":7511A
            Top             =   360
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   131
            Left            =   1560
            Picture         =   "frmMap.frx":75624
            Top             =   480
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   130
            Left            =   1680
            Picture         =   "frmMap.frx":75B2E
            Top             =   360
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   129
            Left            =   1800
            Picture         =   "frmMap.frx":76038
            Top             =   240
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   128
            Left            =   1920
            Picture         =   "frmMap.frx":76542
            Top             =   360
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   127
            Left            =   2040
            Picture         =   "frmMap.frx":76A4C
            Top             =   480
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   126
            Left            =   2160
            Picture         =   "frmMap.frx":76F56
            Top             =   600
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   125
            Left            =   2280
            Picture         =   "frmMap.frx":77460
            Top             =   720
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   124
            Left            =   2400
            Picture         =   "frmMap.frx":7796A
            Top             =   840
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   113
            Left            =   120
            Picture         =   "frmMap.frx":77E74
            Top             =   600
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   112
            Left            =   240
            Picture         =   "frmMap.frx":7837E
            Top             =   480
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   111
            Left            =   360
            Picture         =   "frmMap.frx":78888
            Top             =   480
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   96
            Left            =   480
            Picture         =   "frmMap.frx":78D92
            Top             =   1080
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   95
            Left            =   600
            Picture         =   "frmMap.frx":7929C
            Top             =   960
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   94
            Left            =   600
            Picture         =   "frmMap.frx":797A6
            Top             =   1200
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   93
            Left            =   720
            Picture         =   "frmMap.frx":79CB0
            Top             =   1080
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   92
            Left            =   840
            Picture         =   "frmMap.frx":7A1BA
            Top             =   960
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   91
            Left            =   960
            Picture         =   "frmMap.frx":7A6C4
            Top             =   840
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   90
            Left            =   1080
            Picture         =   "frmMap.frx":7ABCE
            Top             =   720
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   89
            Left            =   1200
            Picture         =   "frmMap.frx":7B0D8
            Top             =   600
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   88
            Left            =   720
            Picture         =   "frmMap.frx":7B5E2
            Top             =   1320
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   87
            Left            =   840
            Picture         =   "frmMap.frx":7BAEC
            Top             =   1200
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   86
            Left            =   960
            Picture         =   "frmMap.frx":7BFF6
            Top             =   1080
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   85
            Left            =   1080
            Picture         =   "frmMap.frx":7C500
            Top             =   960
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   84
            Left            =   1200
            Picture         =   "frmMap.frx":7CA0A
            Top             =   840
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   83
            Left            =   1320
            Picture         =   "frmMap.frx":7CF14
            Top             =   720
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   82
            Left            =   840
            Picture         =   "frmMap.frx":7D41E
            Top             =   1440
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   81
            Left            =   960
            Picture         =   "frmMap.frx":7D928
            Top             =   1320
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   80
            Left            =   1080
            Picture         =   "frmMap.frx":7DE32
            Top             =   1200
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   79
            Left            =   1200
            Picture         =   "frmMap.frx":7E33C
            Top             =   1080
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   78
            Left            =   1320
            Picture         =   "frmMap.frx":7E846
            Top             =   960
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   77
            Left            =   1440
            Picture         =   "frmMap.frx":7ED50
            Top             =   840
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   76
            Left            =   960
            Picture         =   "frmMap.frx":7F25A
            Top             =   1560
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   75
            Left            =   1080
            Picture         =   "frmMap.frx":7F764
            Top             =   1440
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   74
            Left            =   1080
            Picture         =   "frmMap.frx":7FC6E
            Top             =   1680
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   73
            Left            =   1200
            Picture         =   "frmMap.frx":80178
            Top             =   1560
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   72
            Left            =   1200
            Picture         =   "frmMap.frx":80682
            Top             =   1800
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   71
            Left            =   1320
            Picture         =   "frmMap.frx":80B8C
            Top             =   1680
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   70
            Left            =   1440
            Picture         =   "frmMap.frx":81096
            Top             =   1560
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   69
            Left            =   1320
            Picture         =   "frmMap.frx":815A0
            Top             =   1440
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   68
            Left            =   1800
            Picture         =   "frmMap.frx":81AAA
            Top             =   720
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   67
            Left            =   1920
            Picture         =   "frmMap.frx":81FB4
            Top             =   840
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   66
            Left            =   2040
            Picture         =   "frmMap.frx":824BE
            Top             =   960
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   45
            Left            =   1920
            Picture         =   "frmMap.frx":829C8
            Top             =   1080
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   29
            Left            =   1800
            Picture         =   "frmMap.frx":82ED2
            Top             =   1200
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   27
            Left            =   0
            Picture         =   "frmMap.frx":833DC
            Top             =   720
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   65
            Left            =   240
            Picture         =   "frmMap.frx":838E6
            Top             =   1080
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   64
            Left            =   480
            Picture         =   "frmMap.frx":83DF0
            Top             =   1320
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   63
            Left            =   720
            Picture         =   "frmMap.frx":842FA
            Top             =   1560
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   62
            Left            =   960
            Picture         =   "frmMap.frx":84804
            Top             =   1800
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   61
            Left            =   1200
            Picture         =   "frmMap.frx":84D0E
            Top             =   2040
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   60
            Left            =   1440
            Picture         =   "frmMap.frx":85218
            Top             =   1800
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   59
            Left            =   1680
            Picture         =   "frmMap.frx":85722
            Top             =   1560
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   58
            Left            =   1920
            Picture         =   "frmMap.frx":85C2C
            Top             =   1320
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   57
            Left            =   2160
            Picture         =   "frmMap.frx":86136
            Top             =   1080
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   56
            Left            =   2280
            Picture         =   "frmMap.frx":86640
            Top             =   960
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   55
            Left            =   2040
            Picture         =   "frmMap.frx":86B4A
            Top             =   1200
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   54
            Left            =   1800
            Picture         =   "frmMap.frx":87054
            Top             =   1440
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   53
            Left            =   1560
            Picture         =   "frmMap.frx":8755E
            Top             =   1680
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   52
            Left            =   1320
            Picture         =   "frmMap.frx":87A68
            Top             =   1920
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   51
            Left            =   1080
            Picture         =   "frmMap.frx":87F72
            Top             =   1920
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   50
            Left            =   840
            Picture         =   "frmMap.frx":8847C
            Top             =   1680
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   49
            Left            =   600
            Picture         =   "frmMap.frx":88986
            Top             =   1440
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   48
            Left            =   360
            Picture         =   "frmMap.frx":88E90
            Top             =   1200
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   47
            Left            =   120
            Picture         =   "frmMap.frx":8939A
            Top             =   960
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   46
            Left            =   240
            Picture         =   "frmMap.frx":898A4
            Top             =   840
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   44
            Left            =   1200
            Picture         =   "frmMap.frx":89DAE
            Top             =   360
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   43
            Left            =   480
            Picture         =   "frmMap.frx":8A2B8
            Top             =   600
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   42
            Left            =   1320
            Picture         =   "frmMap.frx":8A7C2
            Top             =   480
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   41
            Left            =   2160
            Picture         =   "frmMap.frx":8ACCC
            Top             =   840
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   40
            Left            =   600
            Picture         =   "frmMap.frx":8B1D6
            Top             =   720
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   39
            Left            =   1800
            Picture         =   "frmMap.frx":8B6E0
            Top             =   480
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   38
            Left            =   360
            Picture         =   "frmMap.frx":8BBEA
            Top             =   720
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   37
            Left            =   960
            Picture         =   "frmMap.frx":8C0F4
            Top             =   600
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   36
            Left            =   1920
            Picture         =   "frmMap.frx":8C5FE
            Top             =   600
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   35
            Left            =   1440
            Picture         =   "frmMap.frx":8CB08
            Top             =   600
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   34
            Left            =   1560
            Picture         =   "frmMap.frx":8D012
            Top             =   720
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   33
            Left            =   2040
            Picture         =   "frmMap.frx":8D51C
            Top             =   720
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   32
            Left            =   720
            Picture         =   "frmMap.frx":8DA26
            Top             =   840
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   31
            Left            =   1080
            Picture         =   "frmMap.frx":8DF30
            Top             =   480
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   30
            Left            =   1680
            Picture         =   "frmMap.frx":8E43A
            Top             =   600
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   28
            Left            =   840
            Picture         =   "frmMap.frx":8E944
            Top             =   720
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   26
            Left            =   0
            Picture         =   "frmMap.frx":8EE4E
            Top             =   840
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   25
            Left            =   360
            Picture         =   "frmMap.frx":8F358
            Top             =   960
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   24
            Left            =   480
            Picture         =   "frmMap.frx":8F862
            Top             =   840
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   23
            Left            =   1680
            Picture         =   "frmMap.frx":8FD6C
            Top             =   1320
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   22
            Left            =   1560
            Picture         =   "frmMap.frx":90276
            Top             =   1440
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   21
            Left            =   1680
            Picture         =   "frmMap.frx":90780
            Top             =   840
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   20
            Left            =   1200
            Picture         =   "frmMap.frx":90C8A
            Top             =   1320
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   19
            Left            =   1800
            Picture         =   "frmMap.frx":91194
            Top             =   960
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   18
            Left            =   1320
            Picture         =   "frmMap.frx":9169E
            Top             =   1200
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   17
            Left            =   1440
            Picture         =   "frmMap.frx":91BA8
            Top             =   1320
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   16
            Left            =   1440
            Picture         =   "frmMap.frx":920B2
            Top             =   1080
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   15
            Left            =   1560
            Picture         =   "frmMap.frx":925BC
            Top             =   960
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   14
            Left            =   1560
            Picture         =   "frmMap.frx":92AC6
            Top             =   1200
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   13
            Left            =   1680
            Picture         =   "frmMap.frx":92FD0
            Top             =   1080
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   12
            Left            =   120
            Picture         =   "frmMap.frx":934DA
            Top             =   720
            Width           =   120
         End
         Begin VB.Image mapTileF 
            Height          =   240
            Index           =   10
            Left            =   240
            Picture         =   "frmMap.frx":939E4
            Top             =   600
            Width           =   120
         End
      End
      Begin VB.Frame frmModelB 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   2415
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Visible         =   0   'False
         Width           =   2715
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   123
            Left            =   -200
            Picture         =   "frmMap.frx":93EEE
            Top             =   0
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   122
            Left            =   -200
            Picture         =   "frmMap.frx":943F8
            Top             =   0
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   137
            Left            =   1920
            Picture         =   "frmMap.frx":94902
            Top             =   240
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   136
            Left            =   2040
            Picture         =   "frmMap.frx":94E0C
            Top             =   360
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   135
            Left            =   2160
            Picture         =   "frmMap.frx":95316
            Top             =   480
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   134
            Left            =   1560
            Picture         =   "frmMap.frx":95820
            Top             =   120
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   133
            Left            =   1680
            Picture         =   "frmMap.frx":95D2A
            Top             =   0
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   132
            Left            =   1800
            Picture         =   "frmMap.frx":96234
            Top             =   120
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   131
            Left            =   0
            Picture         =   "frmMap.frx":9673E
            Top             =   720
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   130
            Left            =   720
            Picture         =   "frmMap.frx":96C48
            Top             =   0
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   129
            Left            =   840
            Picture         =   "frmMap.frx":97152
            Top             =   120
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   128
            Left            =   960
            Picture         =   "frmMap.frx":9765C
            Top             =   240
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   127
            Left            =   1080
            Picture         =   "frmMap.frx":97B66
            Top             =   360
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   123
            Left            =   1200
            Picture         =   "frmMap.frx":98070
            Top             =   480
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   122
            Left            =   1320
            Picture         =   "frmMap.frx":9857A
            Top             =   360
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   107
            Left            =   1440
            Picture         =   "frmMap.frx":98A84
            Top             =   240
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   126
            Left            =   600
            Picture         =   "frmMap.frx":98F8E
            Top             =   120
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   125
            Left            =   720
            Picture         =   "frmMap.frx":99498
            Top             =   120
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   124
            Left            =   0
            Picture         =   "frmMap.frx":999A2
            Top             =   840
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   121
            Left            =   2520
            Picture         =   "frmMap.frx":99EAC
            Top             =   840
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   120
            Left            =   2400
            Picture         =   "frmMap.frx":9A3B6
            Top             =   720
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   119
            Left            =   2280
            Picture         =   "frmMap.frx":9A8C0
            Top             =   600
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   118
            Left            =   240
            Picture         =   "frmMap.frx":9ADCA
            Top             =   480
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   117
            Left            =   360
            Picture         =   "frmMap.frx":9B2D4
            Top             =   360
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   116
            Left            =   480
            Picture         =   "frmMap.frx":9B7DE
            Top             =   240
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   115
            Left            =   1080
            Picture         =   "frmMap.frx":9BCE8
            Top             =   480
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   114
            Left            =   120
            Picture         =   "frmMap.frx":9C1F2
            Top             =   600
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   113
            Left            =   1200
            Picture         =   "frmMap.frx":9C6FC
            Top             =   600
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   112
            Left            =   120
            Picture         =   "frmMap.frx":9CC06
            Top             =   720
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   111
            Left            =   240
            Picture         =   "frmMap.frx":9D110
            Top             =   600
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   110
            Left            =   360
            Picture         =   "frmMap.frx":9D61A
            Top             =   480
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   109
            Left            =   480
            Picture         =   "frmMap.frx":9DB24
            Top             =   360
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   108
            Left            =   600
            Picture         =   "frmMap.frx":9E02E
            Top             =   240
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   106
            Left            =   840
            Picture         =   "frmMap.frx":9E538
            Top             =   240
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   105
            Left            =   960
            Picture         =   "frmMap.frx":9EA42
            Top             =   360
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   104
            Left            =   1680
            Picture         =   "frmMap.frx":9EF4C
            Top             =   120
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   103
            Left            =   1440
            Picture         =   "frmMap.frx":9F456
            Top             =   360
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   101
            Left            =   1320
            Picture         =   "frmMap.frx":9F960
            Top             =   480
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   102
            Left            =   2400
            Picture         =   "frmMap.frx":9FE6A
            Top             =   840
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   37
            Left            =   2520
            Picture         =   "frmMap.frx":A0374
            Top             =   960
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   36
            Left            =   2280
            Picture         =   "frmMap.frx":A087E
            Top             =   720
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   35
            Left            =   2160
            Picture         =   "frmMap.frx":A0D88
            Top             =   600
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   34
            Left            =   2040
            Picture         =   "frmMap.frx":A1292
            Top             =   480
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   32
            Left            =   1920
            Picture         =   "frmMap.frx":A179C
            Top             =   360
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   31
            Left            =   1800
            Picture         =   "frmMap.frx":A1CA6
            Top             =   240
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   30
            Left            =   360
            Picture         =   "frmMap.frx":A21B0
            Top             =   960
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   29
            Left            =   480
            Picture         =   "frmMap.frx":A26BA
            Top             =   840
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   28
            Left            =   720
            Picture         =   "frmMap.frx":A2BC4
            Top             =   840
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   27
            Left            =   600
            Picture         =   "frmMap.frx":A30CE
            Top             =   720
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   26
            Left            =   840
            Picture         =   "frmMap.frx":A35D8
            Top             =   720
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   25
            Left            =   720
            Picture         =   "frmMap.frx":A3AE2
            Top             =   600
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   24
            Left            =   1560
            Picture         =   "frmMap.frx":A3FEC
            Top             =   240
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   14
            Left            =   480
            Picture         =   "frmMap.frx":A44F6
            Top             =   1080
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   33
            Left            =   600
            Picture         =   "frmMap.frx":A4A00
            Top             =   960
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   16
            Left            =   2280
            Picture         =   "frmMap.frx":A4F0A
            Top             =   960
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   0
            Left            =   360
            Picture         =   "frmMap.frx":A5414
            Top             =   1200
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   1
            Left            =   240
            Picture         =   "frmMap.frx":A591E
            Top             =   1080
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   2
            Left            =   120
            Picture         =   "frmMap.frx":A5E28
            Top             =   960
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   3
            Left            =   240
            Picture         =   "frmMap.frx":A6332
            Top             =   840
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   4
            Left            =   360
            Picture         =   "frmMap.frx":A683C
            Top             =   720
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   5
            Left            =   480
            Picture         =   "frmMap.frx":A6D46
            Top             =   600
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   6
            Left            =   600
            Picture         =   "frmMap.frx":A7250
            Top             =   480
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   7
            Left            =   720
            Picture         =   "frmMap.frx":A775A
            Top             =   360
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   8
            Left            =   840
            Picture         =   "frmMap.frx":A7C64
            Top             =   480
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   9
            Left            =   960
            Picture         =   "frmMap.frx":A816E
            Top             =   600
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   10
            Left            =   1080
            Picture         =   "frmMap.frx":A8678
            Top             =   720
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   11
            Left            =   1200
            Picture         =   "frmMap.frx":A8B82
            Top             =   840
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   12
            Left            =   1440
            Picture         =   "frmMap.frx":A908C
            Top             =   600
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   13
            Left            =   960
            Picture         =   "frmMap.frx":A9596
            Top             =   840
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   15
            Left            =   1680
            Picture         =   "frmMap.frx":A9AA0
            Top             =   360
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   17
            Left            =   1560
            Picture         =   "frmMap.frx":A9FAA
            Top             =   480
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   18
            Left            =   1320
            Picture         =   "frmMap.frx":AA4B4
            Top             =   720
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   19
            Left            =   2400
            Picture         =   "frmMap.frx":AA9BE
            Top             =   1080
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   20
            Left            =   2160
            Picture         =   "frmMap.frx":AAEC8
            Top             =   840
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   21
            Left            =   2040
            Picture         =   "frmMap.frx":AB3D2
            Top             =   720
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   22
            Left            =   1920
            Picture         =   "frmMap.frx":AB8DC
            Top             =   600
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   23
            Left            =   1800
            Picture         =   "frmMap.frx":ABDE6
            Top             =   480
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   38
            Left            =   1320
            Picture         =   "frmMap.frx":AC2F0
            Top             =   1920
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   39
            Left            =   1440
            Picture         =   "frmMap.frx":AC7FA
            Top             =   1800
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   40
            Left            =   1560
            Picture         =   "frmMap.frx":ACD04
            Top             =   1680
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   41
            Left            =   1680
            Picture         =   "frmMap.frx":AD20E
            Top             =   1560
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   42
            Left            =   1800
            Picture         =   "frmMap.frx":AD718
            Top             =   1440
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   43
            Left            =   2040
            Picture         =   "frmMap.frx":ADC22
            Top             =   1440
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   44
            Left            =   2160
            Picture         =   "frmMap.frx":AE12C
            Top             =   1320
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   45
            Left            =   2280
            Picture         =   "frmMap.frx":AE636
            Top             =   1200
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   46
            Left            =   1320
            Picture         =   "frmMap.frx":AEB40
            Top             =   1680
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   47
            Left            =   1200
            Picture         =   "frmMap.frx":AF04A
            Top             =   1800
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   48
            Left            =   1440
            Picture         =   "frmMap.frx":AF554
            Top             =   1560
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   49
            Left            =   1560
            Picture         =   "frmMap.frx":AFA5E
            Top             =   1440
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   50
            Left            =   1680
            Picture         =   "frmMap.frx":AFF68
            Top             =   1320
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   51
            Left            =   1800
            Picture         =   "frmMap.frx":B0472
            Top             =   1200
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   52
            Left            =   1920
            Picture         =   "frmMap.frx":B097C
            Top             =   1080
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   53
            Left            =   2040
            Picture         =   "frmMap.frx":B0E86
            Top             =   960
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   54
            Left            =   1920
            Picture         =   "frmMap.frx":B1390
            Top             =   840
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   55
            Left            =   1800
            Picture         =   "frmMap.frx":B189A
            Top             =   960
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   56
            Left            =   1680
            Picture         =   "frmMap.frx":B1DA4
            Top             =   1080
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   57
            Left            =   1560
            Picture         =   "frmMap.frx":B22AE
            Top             =   1200
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   58
            Left            =   1440
            Picture         =   "frmMap.frx":B27B8
            Top             =   1320
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   59
            Left            =   1320
            Picture         =   "frmMap.frx":B2CC2
            Top             =   1440
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   60
            Left            =   1200
            Picture         =   "frmMap.frx":B31CC
            Top             =   1560
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   61
            Left            =   1080
            Picture         =   "frmMap.frx":B36D6
            Top             =   1680
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   62
            Left            =   960
            Picture         =   "frmMap.frx":B3BE0
            Top             =   1560
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   63
            Left            =   1080
            Picture         =   "frmMap.frx":B40EA
            Top             =   1440
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   64
            Left            =   1200
            Picture         =   "frmMap.frx":B45F4
            Top             =   1320
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   65
            Left            =   1320
            Picture         =   "frmMap.frx":B4AFE
            Top             =   1200
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   66
            Left            =   1440
            Picture         =   "frmMap.frx":B5008
            Top             =   1080
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   67
            Left            =   1560
            Picture         =   "frmMap.frx":B5512
            Top             =   960
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   68
            Left            =   1680
            Picture         =   "frmMap.frx":B5A1C
            Top             =   840
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   69
            Left            =   1800
            Picture         =   "frmMap.frx":B5F26
            Top             =   720
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   70
            Left            =   840
            Picture         =   "frmMap.frx":B6430
            Top             =   1440
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   71
            Left            =   960
            Picture         =   "frmMap.frx":B693A
            Top             =   1320
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   72
            Left            =   1080
            Picture         =   "frmMap.frx":B6E44
            Top             =   1200
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   73
            Left            =   1200
            Picture         =   "frmMap.frx":B734E
            Top             =   1080
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   74
            Left            =   1320
            Picture         =   "frmMap.frx":B7858
            Top             =   960
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   75
            Left            =   1440
            Picture         =   "frmMap.frx":B7D62
            Top             =   840
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   76
            Left            =   1560
            Picture         =   "frmMap.frx":B826C
            Top             =   720
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   77
            Left            =   1680
            Picture         =   "frmMap.frx":B8776
            Top             =   600
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   78
            Left            =   720
            Picture         =   "frmMap.frx":B8C80
            Top             =   1320
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   79
            Left            =   840
            Picture         =   "frmMap.frx":B918A
            Top             =   1200
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   80
            Left            =   960
            Picture         =   "frmMap.frx":B9694
            Top             =   1080
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   81
            Left            =   1080
            Picture         =   "frmMap.frx":B9B9E
            Top             =   960
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   82
            Left            =   600
            Picture         =   "frmMap.frx":BA0A8
            Top             =   1200
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   83
            Left            =   720
            Picture         =   "frmMap.frx":BA5B2
            Top             =   1080
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   84
            Left            =   840
            Picture         =   "frmMap.frx":BAABC
            Top             =   960
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   85
            Left            =   480
            Picture         =   "frmMap.frx":BAFC6
            Top             =   1320
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   86
            Left            =   600
            Picture         =   "frmMap.frx":BB4D0
            Top             =   1440
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   87
            Left            =   720
            Picture         =   "frmMap.frx":BB9DA
            Top             =   1560
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   88
            Left            =   840
            Picture         =   "frmMap.frx":BBEE4
            Top             =   1680
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   89
            Left            =   960
            Picture         =   "frmMap.frx":BC3EE
            Top             =   1800
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   90
            Left            =   1080
            Picture         =   "frmMap.frx":BC8F8
            Top             =   1920
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   91
            Left            =   1200
            Picture         =   "frmMap.frx":BCE02
            Top             =   2040
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   92
            Left            =   1320
            Picture         =   "frmMap.frx":BD30C
            Top             =   2160
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   93
            Left            =   1440
            Picture         =   "frmMap.frx":BD816
            Top             =   2040
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   94
            Left            =   1560
            Picture         =   "frmMap.frx":BDD20
            Top             =   1920
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   95
            Left            =   1680
            Picture         =   "frmMap.frx":BE22A
            Top             =   1800
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   96
            Left            =   1800
            Picture         =   "frmMap.frx":BE734
            Top             =   1680
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   97
            Left            =   1920
            Picture         =   "frmMap.frx":BEC3E
            Top             =   1560
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   98
            Left            =   1920
            Picture         =   "frmMap.frx":BF148
            Top             =   1320
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   99
            Left            =   2040
            Picture         =   "frmMap.frx":BF652
            Top             =   1200
            Width           =   120
         End
         Begin VB.Image mapTileB 
            Height          =   240
            Index           =   100
            Left            =   2160
            Picture         =   "frmMap.frx":BFB5C
            Top             =   1080
            Width           =   120
         End
      End
      Begin VB.Frame frmModelE 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         ForeColor       =   &H00E0E0E0&
         Height          =   2535
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Visible         =   0   'False
         Width           =   2535
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   121
            Left            =   960
            Picture         =   "frmMap.frx":C0066
            Top             =   120
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   120
            Left            =   1080
            Picture         =   "frmMap.frx":C0570
            Top             =   0
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   119
            Left            =   600
            Picture         =   "frmMap.frx":C0A7A
            Top             =   480
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   118
            Left            =   720
            Picture         =   "frmMap.frx":C0F84
            Top             =   360
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   117
            Left            =   840
            Picture         =   "frmMap.frx":C148E
            Top             =   240
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   116
            Left            =   2280
            Picture         =   "frmMap.frx":C1998
            Top             =   1200
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   115
            Left            =   2160
            Picture         =   "frmMap.frx":C1EA2
            Top             =   1080
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   114
            Left            =   2040
            Picture         =   "frmMap.frx":C23AC
            Top             =   960
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   113
            Left            =   1920
            Picture         =   "frmMap.frx":C28B6
            Top             =   840
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   112
            Left            =   1800
            Picture         =   "frmMap.frx":C2DC0
            Top             =   720
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   111
            Left            =   1680
            Picture         =   "frmMap.frx":C32CA
            Top             =   600
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   110
            Left            =   1560
            Picture         =   "frmMap.frx":C37D4
            Top             =   480
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   109
            Left            =   1440
            Picture         =   "frmMap.frx":C3CDE
            Top             =   360
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   108
            Left            =   1320
            Picture         =   "frmMap.frx":C41E8
            Top             =   240
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   107
            Left            =   1200
            Picture         =   "frmMap.frx":C46F2
            Top             =   120
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   106
            Left            =   1080
            Picture         =   "frmMap.frx":C4BFC
            Top             =   120
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   105
            Left            =   120
            Picture         =   "frmMap.frx":C5106
            Top             =   960
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   104
            Left            =   240
            Picture         =   "frmMap.frx":C5610
            Top             =   840
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   99
            Left            =   360
            Picture         =   "frmMap.frx":C5B1A
            Top             =   720
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   97
            Left            =   960
            Picture         =   "frmMap.frx":C6024
            Top             =   240
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   94
            Left            =   480
            Picture         =   "frmMap.frx":C652E
            Top             =   600
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   0
            Left            =   600
            Picture         =   "frmMap.frx":C6A38
            Top             =   840
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   54
            Left            =   1680
            Picture         =   "frmMap.frx":C6F42
            Top             =   720
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   53
            Left            =   1560
            Picture         =   "frmMap.frx":C744C
            Top             =   840
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   52
            Left            =   1440
            Picture         =   "frmMap.frx":C7956
            Top             =   960
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   101
            Left            =   -220
            Picture         =   "frmMap.frx":C7E60
            Top             =   0
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   100
            Left            =   -200
            Picture         =   "frmMap.frx":C836A
            Top             =   0
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   98
            Left            =   960
            Picture         =   "frmMap.frx":C8874
            Top             =   480
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   96
            Left            =   840
            Picture         =   "frmMap.frx":C8D7E
            Top             =   600
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   95
            Left            =   720
            Picture         =   "frmMap.frx":C9288
            Top             =   720
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   93
            Left            =   480
            Picture         =   "frmMap.frx":C9792
            Top             =   960
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   92
            Left            =   360
            Picture         =   "frmMap.frx":C9C9C
            Top             =   1080
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   91
            Left            =   240
            Picture         =   "frmMap.frx":CA1A6
            Top             =   1200
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   90
            Left            =   1080
            Picture         =   "frmMap.frx":CA6B0
            Top             =   360
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   89
            Left            =   720
            Picture         =   "frmMap.frx":CABBA
            Top             =   480
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   88
            Left            =   840
            Picture         =   "frmMap.frx":CB0C4
            Top             =   360
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   87
            Left            =   600
            Picture         =   "frmMap.frx":CB5CE
            Top             =   600
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   86
            Left            =   480
            Picture         =   "frmMap.frx":CBAD8
            Top             =   720
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   85
            Left            =   360
            Picture         =   "frmMap.frx":CBFE2
            Top             =   840
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   84
            Left            =   240
            Picture         =   "frmMap.frx":CC4EC
            Top             =   960
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   83
            Left            =   120
            Picture         =   "frmMap.frx":CC9F6
            Top             =   1080
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   82
            Left            =   1200
            Picture         =   "frmMap.frx":CCF00
            Top             =   240
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   81
            Left            =   1200
            Picture         =   "frmMap.frx":CD40A
            Top             =   480
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   80
            Left            =   1320
            Picture         =   "frmMap.frx":CD914
            Top             =   360
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   79
            Left            =   1080
            Picture         =   "frmMap.frx":CDE1E
            Top             =   600
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   78
            Left            =   960
            Picture         =   "frmMap.frx":CE328
            Top             =   720
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   77
            Left            =   840
            Picture         =   "frmMap.frx":CE832
            Top             =   840
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   76
            Left            =   720
            Picture         =   "frmMap.frx":CED3C
            Top             =   960
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   75
            Left            =   600
            Picture         =   "frmMap.frx":CF246
            Top             =   1080
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   74
            Left            =   480
            Picture         =   "frmMap.frx":CF750
            Top             =   1200
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   73
            Left            =   360
            Picture         =   "frmMap.frx":CFC5A
            Top             =   1320
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   72
            Left            =   1440
            Picture         =   "frmMap.frx":D0164
            Top             =   480
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   71
            Left            =   1320
            Picture         =   "frmMap.frx":D066E
            Top             =   600
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   70
            Left            =   1200
            Picture         =   "frmMap.frx":D0B78
            Top             =   720
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   69
            Left            =   1080
            Picture         =   "frmMap.frx":D1082
            Top             =   840
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   68
            Left            =   960
            Picture         =   "frmMap.frx":D158C
            Top             =   960
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   67
            Left            =   840
            Picture         =   "frmMap.frx":D1A96
            Top             =   1080
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   66
            Left            =   720
            Picture         =   "frmMap.frx":D1FA0
            Top             =   1200
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   65
            Left            =   600
            Picture         =   "frmMap.frx":D24AA
            Top             =   1320
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   64
            Left            =   480
            Picture         =   "frmMap.frx":D29B4
            Top             =   1440
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   63
            Left            =   1560
            Picture         =   "frmMap.frx":D2EBE
            Top             =   600
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   62
            Left            =   1440
            Picture         =   "frmMap.frx":D33C8
            Top             =   720
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   61
            Left            =   1320
            Picture         =   "frmMap.frx":D38D2
            Top             =   840
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   60
            Left            =   1200
            Picture         =   "frmMap.frx":D3DDC
            Top             =   960
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   59
            Left            =   1080
            Picture         =   "frmMap.frx":D42E6
            Top             =   1080
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   58
            Left            =   960
            Picture         =   "frmMap.frx":D47F0
            Top             =   1200
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   57
            Left            =   840
            Picture         =   "frmMap.frx":D4CFA
            Top             =   1320
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   56
            Left            =   720
            Picture         =   "frmMap.frx":D5204
            Top             =   1440
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   55
            Left            =   600
            Picture         =   "frmMap.frx":D570E
            Top             =   1560
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   51
            Left            =   1320
            Picture         =   "frmMap.frx":D5C18
            Top             =   1080
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   50
            Left            =   1200
            Picture         =   "frmMap.frx":D6122
            Top             =   1200
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   49
            Left            =   1080
            Picture         =   "frmMap.frx":D662C
            Top             =   1320
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   48
            Left            =   960
            Picture         =   "frmMap.frx":D6B36
            Top             =   1440
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   47
            Left            =   840
            Picture         =   "frmMap.frx":D7040
            Top             =   1560
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   46
            Left            =   720
            Picture         =   "frmMap.frx":D754A
            Top             =   1680
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   45
            Left            =   1680
            Picture         =   "frmMap.frx":D7A54
            Top             =   960
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   44
            Left            =   1800
            Picture         =   "frmMap.frx":D7F5E
            Top             =   840
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   43
            Left            =   1560
            Picture         =   "frmMap.frx":D8468
            Top             =   1080
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   42
            Left            =   1440
            Picture         =   "frmMap.frx":D8972
            Top             =   1200
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   41
            Left            =   1320
            Picture         =   "frmMap.frx":D8E7C
            Top             =   1320
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   40
            Left            =   1200
            Picture         =   "frmMap.frx":D9386
            Top             =   1440
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   39
            Left            =   1080
            Picture         =   "frmMap.frx":D9890
            Top             =   1560
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   38
            Left            =   960
            Picture         =   "frmMap.frx":D9D9A
            Top             =   1680
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   37
            Left            =   840
            Picture         =   "frmMap.frx":DA2A4
            Top             =   1800
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   36
            Left            =   1920
            Picture         =   "frmMap.frx":DA7AE
            Top             =   960
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   35
            Left            =   1800
            Picture         =   "frmMap.frx":DACB8
            Top             =   1080
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   34
            Left            =   1680
            Picture         =   "frmMap.frx":DB1C2
            Top             =   1200
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   33
            Left            =   1560
            Picture         =   "frmMap.frx":DB6CC
            Top             =   1320
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   32
            Left            =   1440
            Picture         =   "frmMap.frx":DBBD6
            Top             =   1440
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   31
            Left            =   1320
            Picture         =   "frmMap.frx":DC0E0
            Top             =   1560
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   30
            Left            =   1200
            Picture         =   "frmMap.frx":DC5EA
            Top             =   1680
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   29
            Left            =   1080
            Picture         =   "frmMap.frx":DCAF4
            Top             =   1800
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   28
            Left            =   960
            Picture         =   "frmMap.frx":DCFFE
            Top             =   1920
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   27
            Left            =   1800
            Picture         =   "frmMap.frx":DD508
            Top             =   1320
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   26
            Left            =   1680
            Picture         =   "frmMap.frx":DDA12
            Top             =   1440
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   25
            Left            =   1920
            Picture         =   "frmMap.frx":DDF1C
            Top             =   1200
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   24
            Left            =   1560
            Picture         =   "frmMap.frx":DE426
            Top             =   1560
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   23
            Left            =   1440
            Picture         =   "frmMap.frx":DE930
            Top             =   1680
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   22
            Left            =   1320
            Picture         =   "frmMap.frx":DEE3A
            Top             =   1800
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   21
            Left            =   1200
            Picture         =   "frmMap.frx":DF344
            Top             =   1920
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   20
            Left            =   1080
            Picture         =   "frmMap.frx":DF84E
            Top             =   2040
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   19
            Left            =   1200
            Picture         =   "frmMap.frx":DFD58
            Top             =   2160
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   18
            Left            =   2040
            Picture         =   "frmMap.frx":E0262
            Top             =   1080
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   17
            Left            =   2040
            Picture         =   "frmMap.frx":E076C
            Top             =   1320
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   16
            Left            =   2160
            Picture         =   "frmMap.frx":E0C76
            Top             =   1200
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   15
            Left            =   1920
            Picture         =   "frmMap.frx":E1180
            Top             =   1440
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   14
            Left            =   1800
            Picture         =   "frmMap.frx":E168A
            Top             =   1560
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   13
            Left            =   1680
            Picture         =   "frmMap.frx":E1B94
            Top             =   1680
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   12
            Left            =   1560
            Picture         =   "frmMap.frx":E209E
            Top             =   1800
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   11
            Left            =   1440
            Picture         =   "frmMap.frx":E25A8
            Top             =   1920
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   10
            Left            =   1320
            Picture         =   "frmMap.frx":E2AB2
            Top             =   2040
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   9
            Left            =   2280
            Picture         =   "frmMap.frx":E2FBC
            Top             =   1320
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   8
            Left            =   2160
            Picture         =   "frmMap.frx":E34C6
            Top             =   1440
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   7
            Left            =   2040
            Picture         =   "frmMap.frx":E39D0
            Top             =   1560
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   6
            Left            =   1920
            Picture         =   "frmMap.frx":E3EDA
            Top             =   1680
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   5
            Left            =   1800
            Picture         =   "frmMap.frx":E43E4
            Top             =   1800
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   4
            Left            =   1680
            Picture         =   "frmMap.frx":E48EE
            Top             =   1920
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   3
            Left            =   1560
            Picture         =   "frmMap.frx":E4DF8
            Top             =   2040
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   2
            Left            =   1440
            Picture         =   "frmMap.frx":E5302
            Top             =   2160
            Width           =   120
         End
         Begin VB.Image mapTile 
            Height          =   240
            Index           =   1
            Left            =   1320
            Picture         =   "frmMap.frx":E580C
            Top             =   2280
            Width           =   120
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00000000&
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   3480
         Width           =   2895
         Begin VB.Label lblTile 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Tile"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   720
            TabIndex        =   3
            Top             =   120
            Width           =   2175
         End
         Begin VB.Label lbl 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   0  'Transparent
            Caption         =   " Tile Click:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   0
            TabIndex        =   2
            Top             =   120
            Width           =   975
         End
      End
      Begin VB.Image mapTile 
         Height          =   240
         Index           =   103
         Left            =   0
         Picture         =   "frmMap.frx":E5D16
         Top             =   240
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Image mapTile 
         Height          =   240
         Index           =   102
         Left            =   0
         Picture         =   "frmMap.frx":E6220
         Top             =   600
         Visible         =   0   'False
         Width           =   120
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00E0E0E0&
         Height          =   3135
         Left            =   120
         Top             =   240
         Width           =   2895
      End
   End
End
Attribute VB_Name = "frmMap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'this heightmap code is pretty lame cant really learn from it because
'i was tired to program algorthims but this will do anyways :)
'all the code below is pretty self explonitary
Private Sub Label1_Click()
On Error Resume Next
frmMain.sckclient.SendData "@b" & frmMain.lblName.Caption & " " & txtTile.Text & "/#"
End Sub
Private Sub mapTile_Click(Index As Integer)
mapTile(101).Top = mapTile(Index).Top - 150
mapTile(101).Left = mapTile(Index).Left
End Sub

Private Sub mapTileA_Click(Index As Integer)
mapTile(129).Top = mapTileA(Index).Top - 150
mapTile(129).Left = mapTileA(Index).Left
mapTile(129).Visible = True
End Sub

Private Sub mapTileB_Click(Index As Integer)
mapTile(122).Top = mapTileB(Index).Top - 150
mapTile(122).Left = mapTileB(Index).Left
End Sub

Private Sub mapTileC_Click(Index As Integer)
mapTile(126).Top = mapTileC(Index).Top - 150
mapTile(126).Left = mapTileC(Index).Left
mapTile(126).Visible = True
End Sub

Private Sub mapTileD_Click(Index As Integer)
mapTile(132).Top = mapTileD(Index).Top - 150
mapTile(132).Left = mapTileD(Index).Left
mapTile(132).Visible = True
End Sub

Private Sub mapTileF_Click(Index As Integer)
mapTile(124).Top = mapTileF(Index).Top - 150
mapTile(124).Left = mapTileF(Index).Left
End Sub

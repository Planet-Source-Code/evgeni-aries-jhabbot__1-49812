VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "JHaBBoT (Created By SckBuffer)"
   ClientHeight    =   10155
   ClientLeft      =   -4350
   ClientTop       =   615
   ClientWidth     =   14025
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "frmMain"
   MaxButton       =   0   'False
   ScaleHeight     =   677
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   935
   Begin VB.PictureBox lstHobbas 
      BorderStyle     =   0  'None
      Height          =   2130
      Left            =   255
      ScaleHeight     =   2130
      ScaleWidth      =   2595
      TabIndex        =   67
      Top             =   5400
      Visible         =   0   'False
      Width           =   2595
      Begin VB.ListBox lstPeopleRights 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         ForeColor       =   &H00000000&
         Height          =   1215
         IntegralHeight  =   0   'False
         Left            =   203
         TabIndex        =   71
         Top             =   720
         Width           =   2190
      End
      Begin VB.ListBox lstHobba 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H00000000&
         Height          =   1215
         IntegralHeight  =   0   'False
         Left            =   203
         TabIndex        =   68
         Top             =   720
         Visible         =   0   'False
         Width           =   2190
      End
      Begin VB.Image Image1 
         Height          =   165
         Left            =   2230
         Top             =   120
         Width           =   165
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   " Room Info"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1320
         TabIndex        =   70
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "          Hobbas"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   69
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.Timer tmrMission 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   11160
      Top             =   600
   End
   Begin VB.ListBox lstServers 
      BackColor       =   &H00000000&
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
      ItemData        =   "frmMain.frx":000C
      Left            =   11520
      List            =   "frmMain.frx":0013
      TabIndex        =   62
      Top             =   720
      Width           =   2175
   End
   Begin VB.Timer tmrDance 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   11160
      Top             =   120
   End
   Begin VB.Timer tmrOpenItems 
      Enabled         =   0   'False
      Left            =   13560
      Top             =   120
   End
   Begin VB.Timer tmrGhst 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   13080
      Top             =   120
   End
   Begin VB.Timer tmrPing 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   12600
      Top             =   120
   End
   Begin VB.Timer tmrFlick 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   12120
      Top             =   120
   End
   Begin VB.Timer tmrWave 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   11640
      Top             =   120
   End
   Begin VB.Frame frmRoom 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   5385
      Left            =   11280
      TabIndex        =   43
      Top             =   1800
      Visible         =   0   'False
      Width           =   2655
      Begin VB.Frame frmHack1 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         Height          =   1200
         Left            =   240
         TabIndex        =   65
         Top             =   240
         Visible         =   0   'False
         Width           =   2175
         Begin VB.Label cmdGrabCam 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Grab Cam"
            BeginProperty Font 
               Name            =   "Times New Roman"
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
            TabIndex        =   66
            Top             =   0
            Width           =   1215
         End
         Begin VB.Shape Shape30 
            BorderColor     =   &H000000FF&
            Height          =   195
            Left            =   0
            Top             =   0
            Width           =   1215
         End
      End
      Begin VB.TextBox txtNumPids 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   2280
         TabIndex        =   61
         Top             =   3840
         Width           =   150
      End
      Begin VB.TextBox txtPid 
         BackColor       =   &H00000000&
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   240
         TabIndex        =   60
         Top             =   3840
         Width           =   1935
      End
      Begin VB.ListBox lstItems 
         BackColor       =   &H00000000&
         ForeColor       =   &H00FFFFFF&
         Height          =   2100
         IntegralHeight  =   0   'False
         Left            =   240
         TabIndex        =   48
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label cmdShowHack2 
         BackStyle       =   0  'Transparent
         Caption         =   "Server Hack2"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1320
         TabIndex        =   64
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label cmdShowHack 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Server Hack1"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   240
         TabIndex        =   63
         Top             =   0
         Width           =   1095
      End
      Begin VB.Label cmdCarry 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Carry(Drink)"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1320
         TabIndex        =   59
         Top             =   960
         Width           =   1095
      End
      Begin VB.Shape Shape32 
         BorderColor     =   &H000000FF&
         Height          =   195
         Left            =   1320
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label cmdClearMis 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Clear(Mission)"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   240
         TabIndex        =   58
         Top             =   960
         Width           =   1095
      End
      Begin VB.Shape Shape31 
         BorderColor     =   &H000000FF&
         Height          =   195
         Left            =   1320
         Top             =   720
         Width           =   1095
      End
      Begin VB.Shape Shape29 
         BackColor       =   &H000000FF&
         BorderColor     =   &H000000C0&
         Height          =   135
         Left            =   0
         Top             =   1440
         Width           =   2655
      End
      Begin VB.Label CmdGhost 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Ghost(Figure)"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1320
         TabIndex        =   57
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label CmdPing 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Ping(Room)"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   240
         TabIndex        =   56
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Shape Shape28 
         BorderColor     =   &H000000FF&
         Height          =   195
         Left            =   240
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Shape Shape27 
         BorderColor     =   &H000000FF&
         Height          =   195
         Left            =   1320
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label cmdChoose 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   ":Chooser"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1320
         TabIndex        =   55
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label cmdSingDance 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Single Dance"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   240
         TabIndex        =   54
         Top             =   720
         Width           =   1095
      End
      Begin VB.Shape Shape26 
         BorderColor     =   &H000000FF&
         Height          =   195
         Left            =   240
         Top             =   720
         Width           =   1095
      End
      Begin VB.Shape Shape25 
         BorderColor     =   &H000000FF&
         Height          =   195
         Left            =   240
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label CmdLoopFlick 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Loop(Flick)"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1320
         TabIndex        =   53
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label CmdFlicker 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Flicker()"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   240
         TabIndex        =   52
         Top             =   480
         Width           =   1095
      End
      Begin VB.Shape Shape24 
         BorderColor     =   &H000000FF&
         Height          =   195
         Left            =   1320
         Top             =   480
         Width           =   1095
      End
      Begin VB.Shape Shape23 
         BorderColor     =   &H000000FF&
         Height          =   195
         Left            =   240
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label cmdLoopWave 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Loop(Wave)"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   1320
         TabIndex        =   51
         Top             =   240
         Width           =   1095
      End
      Begin VB.Shape Shape22 
         BorderColor     =   &H000000FF&
         Height          =   195
         Left            =   1320
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label CmdLoopOpen 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Loop(Items)"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   240
         TabIndex        =   50
         Top             =   4080
         Width           =   2175
      End
      Begin VB.Label cmdOpenTle 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Open(Teles,Bars)"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   240
         TabIndex        =   49
         Top             =   4320
         Width           =   2175
      End
      Begin VB.Shape Shape20 
         BorderColor     =   &H000000FF&
         Height          =   195
         Left            =   240
         Top             =   4320
         Width           =   2175
      End
      Begin VB.Shape Shape19 
         BorderColor     =   &H000000FF&
         Height          =   195
         Left            =   240
         Top             =   4080
         Width           =   2175
      End
      Begin VB.Label CmdWave 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Wave()"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   240
         TabIndex        =   47
         Top             =   240
         Width           =   1095
      End
      Begin VB.Shape Shape18 
         BorderColor     =   &H000000FF&
         Height          =   195
         Left            =   240
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label lblTile 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Not Logged Into Room Host"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   240
         TabIndex        =   46
         Top             =   4920
         Width           =   2175
      End
      Begin VB.Label cmdTile 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Change (Tile)"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   240
         TabIndex        =   45
         Top             =   5130
         Width           =   2175
      End
      Begin VB.Shape Shape17 
         BorderColor     =   &H000000FF&
         Height          =   200
         Left            =   240
         Top             =   5150
         Width           =   2175
      End
      Begin VB.Shape Shape16 
         BorderColor     =   &H000000C0&
         Height          =   135
         Left            =   0
         Top             =   4560
         Width           =   2655
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Tile Set:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   20
         Left            =   120
         TabIndex        =   44
         Top             =   4680
         Width           =   2415
      End
      Begin VB.Shape Shape15 
         BorderColor     =   &H000000C0&
         Height          =   135
         Left            =   0
         Top             =   0
         Width           =   2655
      End
      Begin VB.Shape Shape14 
         BorderColor     =   &H000000C0&
         Height          =   5415
         Left            =   2520
         Top             =   0
         Width           =   135
      End
      Begin VB.Shape Shape13 
         BorderColor     =   &H000000C0&
         Height          =   5415
         Left            =   0
         Top             =   0
         Width           =   135
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   1455
      Left            =   11520
      TabIndex        =   35
      Top             =   5730
      Width           =   2175
      Begin VB.CheckBox chkClientIdHack 
         BackColor       =   &H00000000&
         Caption         =   "Scramble User ID(No Ban)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   210
         Left            =   0
         TabIndex        =   40
         Top             =   975
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.CheckBox chkAntiError 
         BackColor       =   &H00000000&
         Caption         =   "AntiError(Disconnect)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   210
         Left            =   0
         MaskColor       =   &H00E0E0E0&
         TabIndex        =   39
         Top             =   270
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.CheckBox chkHobWaveflt 
         BackColor       =   &H00000000&
         Caption         =   "Hobbas,Admin Wave"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   0
         TabIndex        =   38
         Top             =   0
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.CheckBox Check2 
         BackColor       =   &H00000000&
         Caption         =   "AntiError(Script)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   0
         TabIndex        =   37
         Top             =   480
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.CheckBox Check3 
         BackColor       =   &H00000000&
         Caption         =   "No Lag from other people"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   0
         TabIndex        =   36
         Top             =   720
         Value           =   1  'Checked
         Width           =   2175
      End
      Begin VB.Shape Shape11 
         BorderColor     =   &H000000FF&
         Height          =   195
         Left            =   0
         Top             =   1200
         Width           =   2175
      End
      Begin VB.Label cmdUnban 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Unban Shockwave ID"
         BeginProperty Font 
            Name            =   "Times New Roman"
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
         TabIndex        =   41
         Top             =   1200
         Width           =   2175
      End
   End
   Begin VB.Frame frmInfo 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   3585
      Left            =   11415
      TabIndex        =   12
      Top             =   1935
      Width           =   2385
      Begin VB.TextBox lblClientID 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   740
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   14
         Top             =   2760
         Width           =   2175
      End
      Begin VB.TextBox lblFigureNum 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   450
         Left            =   720
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   13
         Top             =   405
         Width           =   1575
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   34
         Top             =   0
         Width           =   615
      End
      Begin VB.Label lblName 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "NameHere"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   600
         TabIndex        =   33
         Top             =   0
         Width           =   1695
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Email:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   32
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblEmail 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "EmailHere"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   600
         TabIndex        =   31
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Figure:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   9
         Left            =   120
         TabIndex        =   30
         Top             =   480
         Width           =   495
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Last Access:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   10
         Left            =   120
         TabIndex        =   29
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label lblLastAccess 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "xx/xx/xx"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   225
         Left            =   1080
         TabIndex        =   28
         Top             =   840
         Width           =   1215
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Last IP used:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   11
         Left            =   120
         TabIndex        =   27
         Top             =   1080
         Width           =   975
      End
      Begin VB.Label lblIP 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "LastIP"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   960
         TabIndex        =   26
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Photo Film:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   12
         Left            =   120
         TabIndex        =   25
         Top             =   1320
         Width           =   975
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "PH Tickets:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   13
         Left            =   120
         TabIndex        =   24
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label lblFilm 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Film"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   960
         TabIndex        =   23
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label lblTickets 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Tickets"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   960
         TabIndex        =   22
         Top             =   1560
         Width           =   1335
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Birthday:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   14
         Left            =   120
         TabIndex        =   21
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label lblBirth 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Birthday"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   840
         TabIndex        =   20
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Account Accesed:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   15
         Left            =   120
         TabIndex        =   19
         Top             =   2040
         Width           =   1455
      End
      Begin VB.Label lblAccess 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "Access"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   1320
         TabIndex        =   18
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Banned?:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   16
         Left            =   120
         TabIndex        =   17
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label lblBan 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Yes/No"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Left            =   840
         TabIndex        =   16
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label lbl 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Decryption ID"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   255
         Index           =   17
         Left            =   0
         TabIndex        =   15
         Top             =   2520
         Width           =   2415
      End
   End
   Begin VB.Timer timedis 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   1920
      Top             =   0
   End
   Begin VB.ListBox lstChatLog 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   1185
      Left            =   240
      TabIndex        =   0
      Top             =   8760
      Width           =   10935
   End
   Begin SHDocVwCtl.WebBrowser wbHabbo 
      Height          =   8295
      Left            =   240
      TabIndex        =   42
      Top             =   240
      Visible         =   0   'False
      Width           =   10935
      ExtentX         =   19288
      ExtentY         =   14631
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin MSWinsockLib.Winsock sckclient 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "fuse-sun3.magenta.net"
      LocalPort       =   37005
   End
   Begin MSWinsockLib.Winsock sckserver 
      Left            =   600
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "fuse-sun3.magenta.net"
      LocalPort       =   37004
   End
   Begin VB.Image trayicon 
      Height          =   465
      Left            =   3000
      Top             =   0
      Visible         =   0   'False
      Width           =   465
   End
   Begin VB.Shape Shape12 
      BorderColor     =   &H000000C0&
      Height          =   135
      Left            =   11280
      Top             =   9675
      Width           =   2655
   End
   Begin VB.Label lblDukeCo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "(C) 2003"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00E0E0E0&
      Height          =   135
      Left            =   11400
      TabIndex        =   11
      Top             =   9795
      Width           =   2415
   End
   Begin VB.Shape Shape10 
      BorderColor     =   &H000000C0&
      Height          =   135
      Left            =   11280
      Top             =   5520
      Width           =   2655
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "lblStatus"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   11400
      TabIndex        =   10
      Top             =   7320
      Width           =   2415
   End
   Begin VB.Shape Shape9 
      BorderColor     =   &H000000C0&
      Height          =   135
      Left            =   11280
      Top             =   1800
      Width           =   2655
   End
   Begin VB.Label cmdDisconnect 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "DISCONNECT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   11520
      TabIndex        =   9
      Top             =   1470
      Width           =   2175
   End
   Begin VB.Shape shpDis 
      BorderColor     =   &H000000FF&
      Height          =   255
      Left            =   11520
      Shape           =   4  'Rounded Rectangle
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label cmdConnect 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "CONNECT"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   11520
      TabIndex        =   8
      Top             =   1110
      Width           =   2175
   End
   Begin VB.Shape shpPanel6 
      BorderColor     =   &H000000C0&
      Height          =   135
      Left            =   11280
      Top             =   7200
      Width           =   2655
   End
   Begin VB.Shape shpConnect 
      BorderColor     =   &H000000FF&
      Height          =   255
      Left            =   11520
      Shape           =   4  'Rounded Rectangle
      Top             =   1080
      Width           =   2175
   End
   Begin VB.Shape shpPanel5 
      BorderColor     =   &H000000C0&
      Height          =   135
      Left            =   11280
      Top             =   570
      Width           =   2655
   End
   Begin VB.Label cmdShowBasic 
      BackColor       =   &H000000FF&
      BackStyle       =   0  'Transparent
      Caption         =   "| Info |"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   13200
      TabIndex        =   7
      Top             =   360
      Width           =   450
   End
   Begin VB.Label cmdShowPublic 
      BackStyle       =   0  'Transparent
      Caption         =   "  |  Public  |"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   12480
      TabIndex        =   6
      Top             =   360
      Width           =   735
   End
   Begin VB.Label cmdShowRoom 
      BackStyle       =   0  'Transparent
      Caption         =   "| Room  |"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   12000
      TabIndex        =   5
      Top             =   360
      Width           =   615
   End
   Begin VB.Label cmdShowLogs 
      BackStyle       =   0  'Transparent
      Caption         =   "| Logs |"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   11520
      TabIndex        =   4
      Top             =   360
      Width           =   525
   End
   Begin VB.Shape shpPanel4 
      BorderColor     =   &H000000C0&
      Height          =   135
      Left            =   11280
      Top             =   7560
      Width           =   2655
   End
   Begin VB.Shape shpPanel3 
      BorderColor     =   &H000000C0&
      Height          =   9735
      Left            =   13800
      Top             =   360
      Width           =   135
   End
   Begin VB.Shape shpPanel2 
      BorderColor     =   &H000000C0&
      Height          =   135
      Left            =   11280
      Top             =   9960
      Width           =   2655
   End
   Begin VB.Shape shpMin3 
      BorderColor     =   &H000000C0&
      Height          =   7335
      Left            =   11280
      Top             =   360
      Width           =   135
   End
   Begin VB.Label cmdShow 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   11295
      TabIndex        =   3
      Top             =   150
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape shpMin1 
      BorderColor     =   &H000000C0&
      FillColor       =   &H000000C0&
      Height          =   2415
      Left            =   11280
      Top             =   7680
      Width           =   135
   End
   Begin VB.Shape shpMin 
      BorderColor     =   &H000000C0&
      Height          =   7575
      Left            =   11280
      Top             =   120
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "   Control Panel"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   11520
      TabIndex        =   2
      Top             =   150
      Width           =   2055
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000C0&
      X1              =   914
      X2              =   914
      Y1              =   8
      Y2              =   24
   End
   Begin VB.Label cmdHide 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "  X"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   13650
      TabIndex        =   1
      Top             =   150
      Width           =   255
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      X1              =   752
      X2              =   928
      Y1              =   24
      Y2              =   24
   End
   Begin VB.Image imgLogo 
      Height          =   1365
      Left            =   11880
      Top             =   8040
      Width           =   1485
   End
   Begin VB.Shape shpImage 
      BorderColor     =   &H000000C0&
      Height          =   2415
      Left            =   11280
      Top             =   7680
      Width           =   2655
   End
   Begin VB.Shape shpPanel1 
      BorderColor     =   &H000000C0&
      FillColor       =   &H000000FF&
      Height          =   9975
      Left            =   11280
      Top             =   120
      Width           =   2655
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000C0&
      Height          =   1455
      Left            =   120
      Top             =   8640
      Width           =   11175
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      BorderColor     =   &H000000C0&
      Height          =   8535
      Left            =   120
      Top             =   120
      Width           =   11175
   End
   Begin VB.Menu MnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuOptions 
         Caption         =   "Options"
      End
   End
   Begin VB.Menu mnuApp 
      Caption         =   "JHabboT"
      Begin VB.Menu mnuAuth 
         Caption         =   "Author"
      End
      Begin VB.Menu MnuHelp 
         Caption         =   "Help"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=============================================
'=This code was created by Evgeni=
'=The app isnt actaully done     =
'=I removed some specified code for          =
'=specified reasons.I tried to document       =
'=the code as much as i could and tried my   =
'=best to explain what each line does.       =
'=This program was for educational use=
'=and anything you do with it doesnt concern =
'=me.:)                                      =
'=im to lazy to recheck the code but it seems=
'=to run perfect last time i checked.:)      =
'=and i give a thank you note to Blake hes the =
'=one who tought me :) and made me start(influence) programming =
'=====================================================
Private Sub cmdCarry_Click()
On Error Resume Next
sckserver.SendData "€W€€€"
End Sub

Private Sub cmdChoose_Click()
On Error Resume Next
sckclient.SendData "@b" & frmMain.lblName.Caption & " " & frmMain.lblTile.Caption & "/mod H/#"
End Sub

Private Sub cmdClearMis_Click()
On Error Resume Next
If cmdClearMis.Caption = "Clear(Mission)" Then
tmrMission.Enabled = True
sckserver.SendData "€I€€€"
cmdClearMis.Caption = "Clear(Stop)"
ElseIf cmdClearMis.Caption = "Clear(Stop)" Then
tmrMission.Enabled = False
cmdClearMis.Caption = "Clear(Mission)"
End If
End Sub

Private Sub cmdConnect_Click()
'just incase if the hosts were edited we load em up as they were anyways
'=========================================================================
'ServerHost = sckserver.RemoteHost
'ServerPort = sckserver.RemotePort
'ClientHost = sckclient.LocalHostName
'ClientPort = sckclient.LocalPort
        sckserver.RemoteHost = "fuse-sun3.magenta.net"
        sckserver.RemotePort = "37005"
'====================================================
    'Now we are loading the habbo file
    '===========================================
wbHabbo.Navigate App.Path & "/tmp.html"
wbHabbo.Visible = True
    '==============================================
    'Sckclient will listen for connection which will occure when you login habbo
    '============================
sckclient.Listen
    '============================
    'since connected disconnect can be press and connect cant be
    '===============================
cmdConnect.Enabled = False
cmdDisconnect.Enabled = True
    '=================================
End Sub

Private Sub cmdDisconnect_Click()
On Error Resume Next
'Send shit to make error for disconnection of server
'=======================================
sckserver.SendData "CLOSE_CONNECTION"
'========================================
'Enable time for the data to be transformed to info
'==============================================
timedis.Enabled = True
'===========================================
'Since disconnected info will be cleared out
'==============
HideInfo
'=============
Picture2.Visible = True
lstHobba.Clear
lstPeopleRights.Clear
For i = 1 To 25
    People(i) = Empty
    Hobbas(i) = Empty
Next i
End Sub

Private Sub CmdFlicker_Click()
Flicker
End Sub

Private Sub cmdGrabCam_Click()
Cam = True
End Sub

Private Sub cmdHide_Click()
'Hides the panel
frmMain.Width = 11595
HidePanel
End Sub
Private Sub cmdHide_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Show that all other objects that can be hovered are unhovered 100%
HoverLabelShow
UnHoverTitle
UnHoverAllCmd
End Sub

Private Sub CmdLoopFlick_Click()
If CmdLoopFlick.Caption = "Loop(Flick)" Then
    FlickerHabbo = True
    tmrFlick.Enabled = True
    CmdLoopFlick.Caption = "Loop(Stop)"
ElseIf CmdLoopFlick.Caption = "Loop(Stop)" Then
    FlickerHabbo = False
    CmdLoopFlick.Caption = "Loop(Flick)"
End If
End Sub

Private Sub cmdLoopWave_Click()
If cmdLoopWave.Caption = "Loop(Wave)" Then
    WaveHabbo = True
    tmrWave.Enabled = True
    cmdLoopWave.Caption = "Loop(Stop)"
ElseIf cmdLoopWave.Caption = "Loop(Stop)" Then
    WaveHabbo = False
    cmdLoopWave.Caption = "Loop(Wave)"
End If
End Sub

Private Sub CmdPing_Click()
If CmdPing.Caption = "Ping(Room)" Then
    Ping = True
    tmrPing.Enabled = True
    CmdPing.Caption = "Stop(Ping)"
ElseIf CmdPing.Caption = "Stop(Ping)" Then
    Ping = False
    CmdPing.Caption = "Ping(Room)"
End If
End Sub

Private Sub cmdShow_Click()
'show panel
frmMain.Width = 14145
ShowPanel
End Sub
Private Sub cmdShow_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'hover cmd show
HoverLabelHide
End Sub
Private Sub cmdShowBasic_Click()
Shape10.Visible = True
frmRoom.Visible = False
End Sub
Private Sub cmdShowBasic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'check function
HoverCmd (4)
End Sub

Private Sub cmdShowHack_Click()
frmMain.frmHack1.Visible = False
End Sub

Private Sub cmdShowHack2_Click()
frmMain.frmHack1.Visible = True
End Sub

Private Sub cmdShowLogs_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'check function
HoverCmd (1)
End Sub
Private Sub cmdShowPublic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'check function
HoverCmd (3)
End Sub

Private Sub cmdShowRoom_Click()
'since frame2 overrightsd frame1 we have shape10 in the way so guess what
Shape10.Visible = False
frmRoom.Visible = True
End Sub

Private Sub cmdShowRoom_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'check function
HoverCmd (2)
End Sub

Private Sub cmdSingDance_Click()
If cmdSingDance.Caption = "Single Move Dance" Then
    DanceHabbo = True
    tmrDance.Enabled = True
    cmdSingDance.Caption = "Stop Dance"
ElseIf cmdSingDance.Caption = "Stop Dance" Then
    DanceHabbo = False
    cmdSingDance.Caption = "Single Move Dance"
End If
End Sub

Private Sub cmdTile_Click()
frmMap.Show
End Sub

Private Sub cmdUnban_Click()
On Error Resume Next
Kill "C:/windows/system32/Macromed/Shockwave 8/Prefs/6FEB4C10.txt" 'this is the cmmand that deletes a file when your banned by habbo habbo overrights this shockwave file with specified values
MsgBox "Unban was successfully executed!", vbInformation, ":)"
End Sub

Private Sub CmdWave_Click()
Wave
End Sub

Private Sub Command1_Click()
For i = 1 To 25
    People(i) = Empty
    Hobbas(i) = Empty
Next i
End Sub

Private Sub Form_Load()
'hides info
HideInfo
cmdDisconnect.Enabled = False
frmMain.lstItems.AddItem "=======Room Items======"
Load frmMap
frmMap.Show
lstHobba.BackColor = RGB(239, 239, 239)
lstPeopleRights.BackColor = RGB(239, 239, 239)
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'unhover all commands
UnHoverAll
UnHoverAllCmd
UnHoverTitle
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload frmMap
UnloadPage = True
timedis.Enabled = True
End Sub

Private Sub Image1_Click()
lstHobbas.Visible = False
End Sub

Private Sub Label1_Click()
lstHobba.Visible = True
lstPeopleRights.Visible = False
Label3.Enabled = True
Label1.Enabled = False
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'unhover all
UnHoverLabelShow
UnHoverAllCmd
HoverTitle
End Sub

Private Sub Label3_Click()
lstPeopleRights.Visible = True
lstHobba.Visible = False
Label3.Enabled = False
Label1.Enabled = True
End Sub

Private Sub mnuOptions_Click()
'frmOptions.Show i had an options but i took away for some reasons
End Sub

Private Sub timedis_Timer()
If UnloadPage = True Then ' checks if the unload is from the window close menu
    sckclient.Close
    sckserver.Close
    musclient.Close
    musserver.Close
    Unload frmPanel
    Unload frmEditData
    Unload Me
Else ' otherwise just disconnect
    sckclient.Close
    sckserver.Close
    wbHabbo.Navigate "about:blank"
    Picture2.Visible = True
    cmdConnect.Enabled = True
    cmdDisconnect.Enabled = False
    wbHabbo.Visible = False
    lblStatus.Caption = ""
    lstHobbas.Visible = False
End If
timedis.Enabled = False 'disable timer that it doesnt redo this like a loop error
End Sub
Private Sub sckserver_Close()
'close connection to the server
sckclient.Close
sckserver.Close
HideInfo
End Sub
Private Sub sckclient_ConnectionRequest(ByVal requestID As Long)
sckserver.Connect 'server connects
Do Until sckserver.State = sckConnected 'reconnect until server is connected
    DoEvents
Loop
'when connected close the client
sckclient.Close
'when the connection is requested accept the id
sckclient.Accept requestID
    'send msg to show that the user is connected
sckclient.SendData "@aMODERATOR WARNING/JHabboT: We are now tapped into the server.#" 'this is a special value in habbo.com for a defined msg box
frmMain.lblStatus.Caption = "Server connected" 'show connected
End Sub
Private Sub sckclient_DataArrival(ByVal bytesTotal As Long)
'sckclient recieves the data and know to work with the server you must then send
'the data to sckserver
On Error Resume Next
    sckclient.GetData SckBuffer
    sckserver.SendData SckBuffer
End Sub
Private Sub sckserver_DataArrival(ByVal bytesTotal As Long)
    On Error Resume Next
'Variables declared in modhack module.
' get data
sckserver.GetData SckBuffer
Indentify = Left(SckBuffer, 2) ' identify the packet
UpdateStatus (Indentify) ' check for status
GetTile
GetModel (Indentify)
GetPeopleWithRights
HobbaWave
GetCommands
sckclient.SendData SckBuffer ' send the info to rhe client to be on the same "page"
End Sub
Private Sub tmrDance_Timer()
If DanceHabbo = True Then
    Dance
ElseIf DanceHabbo = False Then
    tmrDance.Enabled = False
End If
End Sub
Private Sub tmrFlick_Timer()
If FlickerHabbo = True Then
    Flicker
ElseIf FlickerHabbo = False Then
    tmrFlick.Enabled = False
End If
End Sub

Private Sub tmrPing_Timer() ' ping
On Error Resume Next
If Ping = True Then ' check if the button was clicked
    For i = 1 To 20 ' this sends 20 flickers binded together to lag server
        If Ping = True Then
        Flicker ' if stop ping still wasnt clicked then continue
        ElseIf Ping = False Then 'if was then stop
        tmrPing.Enabled = False
        Exit Sub
        End If
    Next i
ElseIf Ping = False Then 'if the stop button was clicked when ping stoped
    tmrPing.Enabled = False ' disable
End If
End Sub

Private Sub tmrWave_Timer()
If WaveHabbo = True Then ' this is self explained
    Wave
ElseIf WaveHabbo = False Then
    tmrWave.Enabled = False
End If
End Sub


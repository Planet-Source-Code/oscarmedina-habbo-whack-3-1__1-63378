VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "shdocvw.dll"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Habbo Whack, Golden 3.1 - OscarMedina    [www.hotelht.com]"
   ClientHeight    =   10125
   ClientLeft      =   -4395
   ClientTop       =   -120
   ClientWidth     =   13965
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "frmMain"
   MaxButton       =   0   'False
   ScaleHeight     =   675
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   931
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Height          =   960
      Left            =   120
      TabIndex        =   2
      Top             =   9075
      Width           =   9345
   End
   Begin VB.TextBox txtnavegar 
      Enabled         =   0   'False
      Height          =   285
      Left            =   600
      TabIndex        =   52
      Text            =   "http://www.hotelht.com/pruebaha7.html"
      Top             =   10680
      Width           =   3015
   End
   Begin VB.TextBox prueba 
      Enabled         =   0   'False
      Height          =   285
      Left            =   11160
      TabIndex        =   48
      Text            =   "0"
      Top             =   10440
      Width           =   255
   End
   Begin VB.TextBox Text7 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   12600
      MaxLength       =   4
      TabIndex        =   47
      Top             =   9120
      Width           =   735
   End
   Begin VB.TextBox Text6 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   12600
      MaxLength       =   4
      TabIndex        =   46
      Top             =   8400
      Width           =   735
   End
   Begin VB.TextBox hc 
      Height          =   285
      Left            =   9120
      TabIndex        =   45
      Text            =   "@Gclub_habbo	active	"
      Top             =   10320
      Width           =   1695
   End
   Begin VB.TextBox Text5 
      Height          =   405
      Left            =   7200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   44
      Top             =   10200
      Visible         =   0   'False
      Width           =   1650
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H80000016&
      Caption         =   "Habilitar"
      Height          =   255
      Left            =   6120
      TabIndex        =   43
      Top             =   10320
      Width           =   1035
   End
   Begin VB.TextBox Text3 
      Height          =   405
      Left            =   600
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   42
      Top             =   10200
      Width           =   2760
   End
   Begin VB.TextBox Text4 
      Height          =   405
      Left            =   3360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   41
      Top             =   10200
      Width           =   2760
   End
   Begin VB.TextBox consola 
      Height          =   285
      Left            =   6240
      TabIndex        =   40
      Text            =   $"frmMain.frx":058A
      Top             =   10680
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   4440
      TabIndex        =   39
      Top             =   10680
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   8175
      Left            =   120
      TabIndex        =   35
      Top             =   600
      Width           =   10800
      Begin SHDocVwCtl.WebBrowser wbHabbo 
         Height          =   8535
         Left            =   -195
         TabIndex        =   36
         Top             =   -120
         Visible         =   0   'False
         Width           =   11505
         ExtentX         =   20285
         ExtentY         =   15055
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
      Height          =   5265
      Left            =   11280
      TabIndex        =   27
      Top             =   2160
      Visible         =   0   'False
      Width           =   2655
      Begin VB.Frame framesalas 
         BackColor       =   &H80000007&
         BorderStyle     =   0  'None
         Height          =   1155
         Left            =   160
         TabIndex        =   53
         Top             =   260
         Visible         =   0   'False
         Width           =   2295
         Begin VB.Shape Shape32 
            BorderColor     =   &H000000FF&
            Height          =   195
            Left            =   1180
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label cmdCarry 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Hobba Alert"
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
            Left            =   1200
            TabIndex        =   71
            Top             =   840
            Width           =   1095
         End
         Begin VB.Shape Shape25 
            BorderColor     =   &H000000FF&
            Height          =   195
            Left            =   90
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label cmdClearMis 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Console MSG"
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
            Left            =   90
            TabIndex        =   70
            Top             =   840
            Width           =   1095
         End
         Begin VB.Shape Shape5 
            BorderColor     =   &H000000FF&
            Height          =   195
            Left            =   480
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Send Packets"
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
            Left            =   480
            TabIndex        =   62
            Top             =   600
            Width           =   1335
         End
         Begin VB.Label Label9 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Get Camera"
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
            Left            =   480
            TabIndex        =   55
            Top             =   360
            Width           =   1335
         End
         Begin VB.Shape Shape17 
            BorderColor     =   &H000000FF&
            Height          =   195
            Left            =   480
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Free Dive"
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
            Left            =   480
            TabIndex        =   54
            Top             =   120
            Width           =   1335
         End
         Begin VB.Shape Shape11 
            BorderColor     =   &H000000FF&
            Height          =   195
            Left            =   480
            Top             =   120
            Width           =   1335
         End
      End
      Begin VB.Frame frmHack1 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   1000
         Left            =   170
         TabIndex        =   34
         Top             =   360
         Visible         =   0   'False
         Width           =   2295
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   0
            TabIndex        =   37
            Top             =   840
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.Shape Shape23 
            BorderColor     =   &H000000FF&
            Height          =   195
            Left            =   0
            Top             =   720
            Width           =   2295
         End
         Begin VB.Label CmdFlicker 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Get HC Room Layouts"
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
            TabIndex        =   73
            Top             =   720
            Width           =   2295
         End
         Begin VB.Label CmdLoopFlick 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Paint Floor"
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
            Left            =   1200
            TabIndex        =   57
            Top             =   0
            Width           =   1095
         End
         Begin VB.Shape Shape27 
            BorderColor     =   &H000000FF&
            Height          =   195
            Left            =   1200
            Top             =   0
            Width           =   1095
         End
         Begin VB.Label CmdPing 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Paint Wall"
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
            TabIndex        =   56
            Top             =   0
            Width           =   1215
         End
         Begin VB.Shape Shape6 
            BorderColor     =   &H000000FF&
            Height          =   195
            Left            =   0
            Top             =   360
            Width           =   2295
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Create/Modify Furni in Room"
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
            TabIndex        =   49
            Top             =   360
            Width           =   2295
         End
         Begin VB.Shape Shape3 
            BorderColor     =   &H000000FF&
            Height          =   195
            Left            =   0
            Top             =   0
            Width           =   1215
         End
      End
      Begin VB.Shape Shape21 
         BorderColor     =   &H000000FF&
         Height          =   195
         Left            =   240
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label cmdBadges 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Badges"
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
         TabIndex        =   74
         Top             =   480
         Width           =   1095
      End
      Begin VB.Shape Shape19 
         BorderColor     =   &H000000FF&
         Height          =   195
         Left            =   1320
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label cmdPerformance 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   ":Performance"
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
         TabIndex        =   72
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Scripts 3"
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
         Left            =   1680
         TabIndex        =   58
         Top             =   0
         Width           =   735
      End
      Begin VB.Image Image1 
         Height          =   1875
         Left            =   600
         Picture         =   "frmMain.frx":05D9
         Top             =   2040
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.Image Image3 
         Height          =   2160
         Left            =   600
         Picture         =   "frmMain.frx":9A8B
         Top             =   1920
         Width           =   1635
      End
      Begin VB.Label cmdShowHack2 
         BackStyle       =   0  'Transparent
         Caption         =   "Scripts 2"
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
         Left            =   960
         TabIndex        =   33
         Top             =   0
         Width           =   735
      End
      Begin VB.Label cmdShowHack 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Scripts 1"
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
         Left            =   120
         TabIndex        =   32
         Top             =   0
         Width           =   855
      End
      Begin VB.Shape Shape31 
         BorderColor     =   &H000000FF&
         Height          =   195
         Left            =   1320
         Top             =   960
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
      Begin VB.Label cmdFurni 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   ":Furni"
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
         TabIndex        =   31
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label cmdchooser 
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
         Left            =   240
         TabIndex        =   30
         Top             =   960
         Width           =   1095
      End
      Begin VB.Shape Shape26 
         BorderColor     =   &H000000FF&
         Height          =   195
         Left            =   240
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label lblTile 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "0,0,0,0,0"
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
         TabIndex        =   29
         Top             =   4920
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
         Caption         =   "Coordinates:"
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
         TabIndex        =   28
         Top             =   4680
         Width           =   2415
      End
      Begin VB.Shape Shape15 
         BorderColor     =   &H000000C0&
         Height          =   255
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
      Height          =   2415
      Left            =   11520
      TabIndex        =   26
      Top             =   5040
      Width           =   2175
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "OscarMedina"
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   480
         TabIndex        =   38
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Image Image2 
         Height          =   1410
         Left            =   120
         Picture         =   "frmMain.frx":1534D
         Top             =   480
         Width           =   2055
      End
   End
   Begin VB.Frame frmInfo 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   2265
      Left            =   11475
      TabIndex        =   11
      Top             =   2400
      Width           =   2325
      Begin VB.TextBox lblClientID 
         BackColor       =   &H00FFFFFF&
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
         Height          =   255
         Left            =   1440
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   13
         Top             =   2520
         Width           =   855
      End
      Begin VB.TextBox lblFigureNum 
         Alignment       =   2  'Center
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
         Height          =   210
         Left            =   -120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         TabIndex        =   12
         Text            =   "frmMain.frx":1EAD7
         Top             =   1440
         Width           =   2535
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Credits:"
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
         Index           =   3
         Left            =   1120
         TabIndex        =   69
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label lblcredits 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "###"
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
         Left            =   1680
         TabIndex        =   68
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label lblmission 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Mission"
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
         Left            =   50
         TabIndex        =   67
         Top             =   240
         Width           =   2200
      End
      Begin VB.Label lblfilm 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "###"
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
         Left            =   1800
         TabIndex        =   66
         Top             =   1920
         Width           =   495
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Film:"
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
         Index           =   2
         Left            =   1440
         TabIndex        =   65
         Top             =   1920
         Width           =   375
      End
      Begin VB.Label lblsex 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "sexo"
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
         Left            =   450
         TabIndex        =   64
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Sex:"
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
         Index           =   1
         Left            =   120
         TabIndex        =   63
         Top             =   1680
         Width           =   495
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
         TabIndex        =   25
         Top             =   0
         Width           =   615
      End
      Begin VB.Label lblName 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Habbo Nombre"
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
         TabIndex        =   24
         Top             =   0
         Width           =   1695
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Habbo E-mail"
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
         Left            =   720
         TabIndex        =   23
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label lblEmail 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "email"
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
         Left            =   120
         TabIndex        =   22
         Top             =   720
         Width           =   2175
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Figure"
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
         Left            =   960
         TabIndex        =   21
         Top             =   1200
         Width           =   495
      End
      Begin VB.Label lblLastAccess 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
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
         Left            =   240
         TabIndex        =   20
         Top             =   0
         Width           =   15
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
         Left            =   -240
         TabIndex        =   19
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Birth Date:"
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
         TabIndex        =   18
         Top             =   960
         Width           =   975
      End
      Begin VB.Label lblBirth 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Fecha"
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
         Left            =   1080
         TabIndex        =   17
         Top             =   960
         Width           =   1095
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
         Left            =   720
         TabIndex        =   16
         Top             =   2760
         Width           =   975
      End
      Begin VB.Label lbl 
         BackStyle       =   0  'Transparent
         Caption         =   "Lido Tickets:"
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
         Left            =   60
         TabIndex        =   15
         Top             =   1920
         Width           =   975
      End
      Begin VB.Label lbltickets 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "###"
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
         Left            =   860
         TabIndex        =   14
         Top             =   1920
         Width           =   615
      End
   End
   Begin VB.Timer timedis 
      Enabled         =   0   'False
      Interval        =   2000
      Left            =   0
      Top             =   0
   End
   Begin MSWinsockLib.Winsock sckServer 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "proxy-6.habbohotel.co.uk"
      RemotePort      =   37009
   End
   Begin MSWinsockLib.Winsock sckClient 
      Left            =   0
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      RemoteHost      =   "fuse-sun7.magenta.net"
      RemotePort      =   37009
      LocalPort       =   37009
   End
   Begin VB.Image imgMinBtnHover 
      Height          =   300
      Left            =   13200
      Picture         =   "frmMain.frx":1EAE0
      Top             =   120
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image imgCloseBtnHover 
      Height          =   300
      Left            =   13590
      Picture         =   "frmMain.frx":1EBBF
      Top             =   120
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.Image imgCloseBtn 
      Height          =   300
      Left            =   13590
      Picture         =   "frmMain.frx":1ECA0
      Top             =   120
      Width           =   300
   End
   Begin VB.Image imgMinBtn 
      Height          =   300
      Left            =   13200
      Picture         =   "frmMain.frx":1ED85
      Top             =   120
      Width           =   300
   End
   Begin VB.Label Label13 
      BackColor       =   &H00000080&
      Height          =   225
      Left            =   13080
      TabIndex        =   61
      Top             =   240
      Width           =   975
   End
   Begin VB.Label Label12 
      BackColor       =   &H000000C0&
      Height          =   255
      Left            =   13080
      TabIndex        =   60
      Top             =   0
      Width           =   975
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Height          =   495
      Left            =   0
      TabIndex        =   59
      Top             =   0
      Width           =   13095
   End
   Begin VB.Image Image5 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   3600
      Picture         =   "frmMain.frx":1EE64
      Top             =   75
      Width           =   6150
   End
   Begin VB.Shape Shape8 
      BorderColor     =   &H000000FF&
      Height          =   315
      Left            =   9600
      Top             =   9660
      Width           =   1335
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Save Chat Log"
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
      Left            =   9600
      TabIndex        =   51
      Top             =   9720
      Width           =   1335
   End
   Begin VB.Shape Shape7 
      BorderColor     =   &H000000FF&
      Height          =   315
      Left            =   9600
      Top             =   9120
      Width           =   1335
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Errase Chat Log"
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
      Left            =   9600
      TabIndex        =   50
      Top             =   9150
      Width           =   1335
   End
   Begin VB.Image Image4 
      Height          =   555
      Left            =   11880
      Picture         =   "frmMain.frx":1F588
      Top             =   8880
      Width           =   480
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
      Top             =   9600
      Width           =   2655
   End
   Begin VB.Label lblDukeCo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "© OscarMedina 2005"
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
      Height          =   255
      Left            =   11400
      TabIndex        =   10
      Top             =   9750
      Width           =   2415
   End
   Begin VB.Shape Shape10 
      BorderColor     =   &H000000C0&
      Height          =   135
      Left            =   11400
      Top             =   4800
      Width           =   2415
   End
   Begin VB.Label lblStatus 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "[Not Connected]"
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
      TabIndex        =   9
      Top             =   7680
      Width           =   2415
   End
   Begin VB.Shape Shape9 
      BorderColor     =   &H000000C0&
      Height          =   135
      Left            =   11400
      Top             =   2160
      Width           =   2415
   End
   Begin VB.Label cmdDisconnect 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "DESCONNECT"
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
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Shape shpDis 
      BorderColor     =   &H000000FF&
      Height          =   255
      Left            =   11520
      Shape           =   4  'Rounded Rectangle
      Top             =   1680
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
      TabIndex        =   1
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Shape shpPanel6 
      BorderColor     =   &H000000C0&
      Height          =   135
      Left            =   11400
      Top             =   7560
      Width           =   2415
   End
   Begin VB.Shape shpConnect 
      BorderColor     =   &H000000FF&
      Height          =   255
      Left            =   11520
      Shape           =   4  'Rounded Rectangle
      Top             =   1320
      Width           =   2175
   End
   Begin VB.Shape shpPanel5 
      BorderColor     =   &H000000C0&
      Height          =   135
      Left            =   11400
      Top             =   960
      Width           =   2415
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
      Left            =   13275
      TabIndex        =   7
      Top             =   720
      Width           =   450
   End
   Begin VB.Label cmdShowRoom 
      BackStyle       =   0  'Transparent
      Caption         =   "| Scripts |"
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
      Left            =   12555
      TabIndex        =   6
      Top             =   720
      Width           =   630
   End
   Begin VB.Label cmdShowLogs 
      BackStyle       =   0  'Transparent
      Caption         =   "| Hotel View |"
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
      Left            =   11550
      TabIndex        =   5
      Top             =   720
      Width           =   1005
   End
   Begin VB.Shape shpPanel4 
      BorderColor     =   &H000000C0&
      Height          =   135
      Left            =   11280
      Top             =   7920
      Width           =   2655
   End
   Begin VB.Shape shpPanel3 
      BorderColor     =   &H000000C0&
      Height          =   9375
      Left            =   13800
      Top             =   720
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
      Top             =   720
      Width           =   135
   End
   Begin VB.Label cmdShow 
      BackStyle       =   0  'Transparent
      Caption         =   "X"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   11280
      TabIndex        =   4
      Top             =   510
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Shape shpMin 
      BorderColor     =   &H000000C0&
      Height          =   7575
      Left            =   11280
      Top             =   480
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Options"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   11640
      TabIndex        =   3
      Top             =   480
      Width           =   2055
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000C0&
      X1              =   914
      X2              =   914
      Y1              =   32
      Y2              =   48
   End
   Begin VB.Label cmdHide 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "  X"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   13650
      TabIndex        =   0
      Top             =   480
      Width           =   255
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      X1              =   752
      X2              =   928
      Y1              =   48
      Y2              =   48
   End
   Begin VB.Image imgLogo 
      Height          =   585
      Left            =   11880
      Picture         =   "frmMain.frx":203AA
      Top             =   8160
      Width           =   495
   End
   Begin VB.Shape shpImage 
      BorderColor     =   &H000000C0&
      Height          =   1575
      Left            =   11400
      Top             =   8040
      Width           =   2415
   End
   Begin VB.Shape shpPanel1 
      BorderColor     =   &H000000C0&
      FillColor       =   &H000000FF&
      Height          =   9615
      Left            =   11280
      Top             =   480
      Width           =   2655
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000C0&
      Height          =   1095
      Left            =   75
      Top             =   9000
      Width           =   11025
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H000000FF&
      BorderColor     =   &H000000C0&
      Height          =   8415
      Left            =   75
      Top             =   480
      Width           =   11010
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Tnx to Aries for layout - PSC UserName = Evgeni(Aries)
'All Habbo connections, Packets and functions were Created/Discovered by me
'Any comments, and stuff send me an e-mail to oscar.medina88@gmail.com
'OscarM3dina on HabboHotel.co.uk
'OscarMedina on all other Hotels Except .es
'Whack = Winsock Hack
'For 1024 x 768 or grater resolutions

Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Sub ActivarCheck1()
Check1.Value = 1
End Sub
Public Sub DesactivarCheck1()
Check1.Value = 0
End Sub

Private Sub cmdBadges_Click()
form8.Show
End Sub

Private Sub cmdCarry_Click()
'Clientside Hobba Message
On Error Resume Next
Dim mensaje As String
alerta = InputBox("Write the Alert text.")
sckClient.SendData "@amod_warn/" & alerta & ""
End Sub

Private Sub cmdCcommands_Click()
Form9.Show
End Sub

Private Sub cmdchooser_Click()
':chooser enabler
On Error Resume Next
sckClient.SendData "@Bfuse_habbo_chooser"
End Sub

Private Sub cmdClearMis_Click()
'Clientside Console Message Adverticement
Dim mensaje As String
mensaje = InputBox("Write Message Text.")
sckClient.SendData consola.Text & mensaje & ""
Text2.Text = ""
End Sub
Private Sub cmdConnect_Click()
wbHabbo.Navigate txtnavegar.Text
wbHabbo.Visible = True
sckClient.Close
sckServer.Close
sckClient.Listen

    '=================================
cmdConnect.Enabled = False
cmdDisconnect.Enabled = True
    '=================================
End Sub

Private Sub cmdDisconnect_Click()
On Error Resume Next
wbHabbo.Visible = False
sckServer.SendData "CLOSE_CONNECTION"
timedis.Enabled = True
Call Form_Load
End Sub

Private Sub CmdFlicker_Click()
'Abilitate Hc Room Layouts
On Error Resume Next
frmMain.sckClient.SendData "@Bfuse_use_special_room_layouts"
End Sub

Private Sub cmdFurni_Click()
':furni command enabler, need to habilitate each time you enter a room
sckClient.SendData "@Bfuse_furni_chooser"
End Sub

Private Sub cmdHide_Click()
frmMain.Width = "11430"
HidePanel
End Sub
Private Sub cmdHide_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
HoverLabelShow
UnHoverTitle
UnHoverAllCmd
End Sub

Private Sub CmdLoopFlick_Click()
'Clientside Floor Color
Dim piso As String
piso = InputBox("Write the Floor color code - see tutorial")
sckClient.SendData "@nfloor/" & piso & ""
End Sub

Private Sub cmdPerformance_Click()
':performance enabler
On Error Resume Next
sckClient.SendData "@Bfuse_performance_panel"
End Sub

Private Sub CmdPing_Click()
'Clientside paint Walls
Dim pared As String
pared = InputBox("Write the Walls Color Code - see tutorial")
sckClient.SendData "@nwallpaper/" & pared & ""
End Sub

Private Sub cmdShow_Click()
frmMain.Width = "13965"
ShowPanel
End Sub
Private Sub cmdShow_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
HoverLabelHide
End Sub
Private Sub cmdShowBasic_Click()
Shape10.Visible = True
frmRoom.Visible = False
End Sub
Private Sub cmdShowBasic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
HoverCmd (4)
End Sub

Private Sub cmdShowHack_Click()
frmHack1.Visible = False
framesalas.Visible = False
Image1.Visible = False
Image3.Visible = True
End Sub

Private Sub cmdShowHack2_Click()
frmHack1.Visible = True
framesalas.Visible = False
Image1.Visible = True
Image3.Visible = False
End Sub

Private Sub cmdShowLogs_Click()
'Go to Hetel View
sckClient.SendData "@R"
DesactivarCheck1
End Sub

Private Sub cmdShowLogs_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
HoverCmd (1)
End Sub

Private Sub cmdShowRoom_Click()
Shape10.Visible = False
frmRoom.Visible = True
End Sub

Private Sub cmdShowRoom_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
HoverCmd (2)
End Sub

Private Sub Form_Load()
HideInfo
cmdDisconnect.Enabled = False
cmdConnect.Enabled = True
BlockearFunciones
End Sub
Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
UnHoverAll
UnHoverAllCmd
UnHoverTitle
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
'Deletes temporal file
On Error Resume Next
Kill "C:/tmp.html"
End Sub

Private Sub Form_Unload(Cancel As Integer)
UnloadPage = True
timedis.Enabled = True
End
End Sub
Private Sub Image4_Click()
'Clientside Credits
sckClient.SendData "@F" & Text7.Text & ".0"
sckClient.SendData "BKYou have " & Text7.Text & " Client Side credits on your purse"
End Sub

Private Sub imgCloseBtn_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgCloseBtn.Visible = False
imgCloseBtnHover.Visible = True
End Sub

Private Sub imgCloseBtnHover_Click()
Unload frmMain
End Sub

Private Sub imgLogo_Click()
'ClientSide HC Days
sckClient.SendData hc.Text & Text6.Text & ""
sckClient.SendData "BKYou have " & Text6.Text & " Client Side HC Days"
End Sub

Private Sub imgMinBtn_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgMinBtnHover.Visible = True
imgMinBtn.Visible = False
End Sub

Private Sub imgMinBtnHover_Click()
frmMain.WindowState = vbMinimized
imgMinBtnHover.Visible = False
imgMinBtn.Visible = True
End Sub

Private Sub Label10_Click()
frmHack1.Visible = False
framesalas.Visible = True
Image1.Visible = True
Image3.Visible = False
End Sub

Private Sub Label11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ReleaseCapture
    Call SendMessage(frmMain.hWnd, &HA1, 2, 0&)
End Sub

Private Sub Label12_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgMinBtnHover.Visible = False
imgMinBtn.Visible = True

imgCloseBtn.Visible = True
imgCloseBtnHover.Visible = False
End Sub

Private Sub Label13_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
imgMinBtnHover.Visible = False
imgMinBtn.Visible = True

imgCloseBtn.Visible = True
imgCloseBtnHover.Visible = False
End Sub


Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
UnHoverLabelShow
UnHoverAllCmd
HoverTitle
End Sub

Private Sub Label4_Click()
Form1.Show
End Sub

Private Sub Label5_Click()
Form3.Visible = True
Form3.ID.Text = Form3.ID.Text + 1
Form3.X.Text = "5"
Form3.Y.Text = "5"
Form3.Z.Text = "0.0"
Form3.R.Text = "0"
Form3.sprite.Text = ""
Form3.L.Text = "1"
Form3.W.Text = "1"
Form3.color.Text = ""
Form3.Variables.Clear
End Sub

Private Sub Label6_Click()
lstChatLog.Clear
End Sub

Private Sub Label7_Click()
'Saves Chat Log
On Error Resume Next
GuardarChat App.Path & "\Chat Log.html", lstChatLog
sckClient.SendData "BKChat Log Saved"
End Sub

Private Sub Label8_Click()
'Clientside Dive
frmMain.sckClient.SendData "A}"
End Sub

Private Sub Label9_Click()
'Serverside Carry Camera
frmMain.sckClient.SendData "BLSI-55112730S5511273camera"
End Sub

Private Sub lblName_Change()
If Not lblName.Caption = "" Then
Label1.Caption = lblName.Caption
End If
End Sub

Private Sub timedis_Timer()
'checks if the unload is from the window close menu
If UnloadPage = True Then
    sckClient.Close
    sckServer.Close
    musclient.Close
    musserver.Close
    Unload frmPanel
    Unload frmEditData
    Unload Me
Else ' otherwise just disconnect
    sckClient.Close
    sckServer.Close
    wbHabbo.Navigate "about:blank"
    cmdConnect.Enabled = True
    cmdDisconnect.Enabled = False
    wbHabbo.Visible = False
    lblStatus.Caption = "No Connection"
End If
timedis.Enabled = False 'disable timer that it doesnt redo this like a loop error
End Sub
Private Sub sckserver_Close()
sckClient.Close
sckServer.Close
lblStatus.Caption = "[Disconnected]"
Call Form_Load
End Sub
Private Sub sckClient_ConnectionRequest(ByVal requestID As Long)
sckServer.Connect               'Automatically connect the Server Socket.
sckClient.Close                 'Close's the Client Socket so we can accept the connection request.
sckClient.Accept requestID      'Accepts the Connection request.
lblStatus.Caption = "Server Connected" 'show connected
End Sub
Private Sub sckclient_DataArrival(ByVal bytesTotal As Long)
On Error Resume Next
    sckClient.GetData SckBuffer
    sckServer.SendData SckBuffer
End Sub
Private Sub sckserver_DataArrival(ByVal bytesTotal As Long)
    On Error Resume Next
sckServer.GetData SckBuffer
Indentify = Left(SckBuffer, 2)
UpdateStatus (Indentify)

Text5.Text = SckBuffer
GetChat (Indentify)
GetTile
ProcessServData
End Sub

Function ProcessServData()
'Checks Packets
If Check1.Value = 1 Then
Text5.Text = Replace(Text5.Text, Text3.Text, Text4.Text)
sckClient.SendData Text5.Text
Else
sckClient.SendData Text5.Text
End If
End Function

Public Sub GuardarChat(sLocation As String, lstListBox As ListBox)
'Save chat function
On Error Resume Next
Dim sCurrent As String
Dim I As Integer
Open sLocation For Output As #1
I = 0
Do Until I = lstListBox.ListCount
sCurrent = lstListBox.List(I)
Print #1, sCurrent & "<br>"
I = I + 1
Loop
Close #1
End Sub

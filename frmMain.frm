VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{28D47522-CF84-11D1-834C-00A0249F0C28}#1.0#0"; "Gif89.dll"
Begin VB.Form frmMain 
   BackColor       =   &H00C0C0C0&
   Caption         =   "Casper Semiramis II  -  WebSite Producer"
   ClientHeight    =   7890
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10290
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7890
   ScaleWidth      =   10290
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      ScaleHeight     =   345
      ScaleWidth      =   10665
      TabIndex        =   1
      Top             =   7560
      Width           =   10695
      Begin MSComctlLib.Toolbar Toolbar 
         Height          =   330
         Left            =   0
         TabIndex        =   48
         Top             =   0
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   582
         ButtonWidth     =   3387
         ButtonHeight    =   582
         Appearance      =   1
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "imgButt"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   3
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Clear All Fields   "
               Key             =   "Clear"
               ImageIndex      =   32
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Load and Save   "
               Key             =   "Load/Save"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Exit Program   "
               Key             =   "Exit"
               ImageIndex      =   33
            EndProperty
         EndProperty
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   7575
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   13361
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      WordWrap        =   0   'False
      ShowFocusRect   =   0   'False
      BackColor       =   12632256
      ForeColor       =   4194304
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "WebSite Data"
      TabPicture(0)   =   "frmMain.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(2)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(3)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label1(4)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label1(5)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label1(6)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label1(7)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1(8)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label1(9)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label1(10)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label1(11)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label1(12)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label1(13)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label1(19)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label1(20)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "imgButt"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "CoUrl"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtURL"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtTitle"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtDes"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "CoCategory"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "k1"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "k2"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "k3"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "k6"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "k5"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "k4"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "k7"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "k8"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "k9"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "k12"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "k11"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "k10"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "k13"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "k14"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "k15"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "k18"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "k17"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "k16"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "txtName"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "txtCompany"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "txtCity"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "txtCountry"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "txtMail"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "txtAddress"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "txtProvince"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "txtPostal"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "txtPhone"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).ControlCount=   49
      TabCaption(1)   =   "Engines"
      TabPicture(1)   =   "frmMain.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(1)=   "Frame2"
      Tab(1).Control(2)=   "Frame3"
      Tab(1).Control(3)=   "ToolbarEngines"
      Tab(1).ControlCount=   4
      TabCaption(2)   =   "Submit"
      TabPicture(2)   =   "frmMain.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label6"
      Tab(2).Control(1)=   "Label5"
      Tab(2).Control(2)=   "Frame4"
      Tab(2).Control(3)=   "Frame5"
      Tab(2).Control(4)=   "Frame6"
      Tab(2).Control(5)=   "Command1"
      Tab(2).Control(6)=   "cmdCancel"
      Tab(2).Control(7)=   "Frame8"
      Tab(2).ControlCount=   8
      TabCaption(3)   =   "Report"
      TabPicture(3)   =   "frmMain.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame7"
      Tab(3).ControlCount=   1
      Begin VB.Frame Frame8 
         Height          =   1215
         Left            =   -74760
         TabIndex        =   89
         Top             =   6120
         Width           =   9615
         Begin GIF89LibCtl.Gif89a picBanner 
            Height          =   855
            Left            =   120
            OleObjectBlob   =   "frmMain.frx":0070
            TabIndex        =   90
            Top             =   240
            Width           =   7095
         End
      End
      Begin VB.Frame Frame7 
         Height          =   6615
         Left            =   -74782
         TabIndex        =   81
         Top             =   638
         Width           =   9855
         Begin RichTextLib.RichTextBox Rtf 
            Height          =   6015
            Left            =   240
            TabIndex        =   82
            Top             =   360
            Width           =   9375
            _ExtentX        =   16536
            _ExtentY        =   10610
            _Version        =   393217
            Enabled         =   -1  'True
            ScrollBars      =   2
            TextRTF         =   $"frmMain.frx":00B2
         End
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "Cancel Submition"
         Height          =   375
         Left            =   -68640
         TabIndex        =   79
         Top             =   5640
         Width           =   1935
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Start Submition"
         Height          =   375
         Left            =   -70920
         TabIndex        =   78
         Top             =   5640
         Width           =   1935
      End
      Begin VB.Frame Frame6 
         Caption         =   "  Simple Statistics "
         Height          =   2415
         Left            =   -70920
         TabIndex        =   69
         Top             =   840
         Width           =   5775
         Begin MSComctlLib.ProgressBar Prg 
            Height          =   255
            Left            =   240
            TabIndex        =   73
            Top             =   2000
            Width           =   5295
            _ExtentX        =   9340
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   1
            Scrolling       =   1
         End
         Begin VB.Label Label1 
            Caption         =   "Current Engine Progress:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   21
            Left            =   240
            TabIndex        =   94
            Top             =   1680
            Width           =   2295
         End
         Begin VB.Label Label7 
            Caption         =   "%"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2280
            TabIndex        =   93
            Top             =   840
            Width           =   255
         End
         Begin VB.Label E_done 
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4920
            TabIndex        =   86
            Top             =   825
            Width           =   495
         End
         Begin VB.Label E_Left 
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4920
            TabIndex        =   85
            Top             =   380
            Width           =   495
         End
         Begin VB.Label Succes 
            Caption         =   "N/A"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1080
            TabIndex        =   84
            Top             =   1225
            Width           =   975
         End
         Begin VB.Label Procent 
            Caption         =   "0"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   1920
            TabIndex        =   83
            Top             =   820
            Width           =   495
         End
         Begin VB.Label lblEngine 
            Caption         =   "None"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1560
            TabIndex        =   80
            Top             =   380
            Width           =   1935
         End
         Begin VB.Label Label1 
            Caption         =   "Engines Done:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   18
            Left            =   3480
            TabIndex        =   75
            Top             =   800
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Engines Left:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   17
            Left            =   3480
            TabIndex        =   74
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label1 
            Caption         =   "Sucess:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   16
            Left            =   240
            TabIndex        =   72
            Top             =   1200
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Procentage Done:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   15
            Left            =   240
            TabIndex        =   71
            Top             =   800
            Width           =   1695
         End
         Begin VB.Label Label1 
            Caption         =   "Engine Name:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   14
            Left            =   240
            TabIndex        =   70
            Top             =   360
            Width           =   1335
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Submited So Far:"
         Height          =   2415
         Left            =   -74760
         TabIndex        =   67
         Top             =   3600
         Width           =   3495
         Begin VB.ListBox Submited 
            Height          =   1815
            Left            =   240
            TabIndex        =   68
            Top             =   360
            Width           =   3015
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Engines To Submit"
         Height          =   2415
         Left            =   -74760
         TabIndex        =   65
         Top             =   840
         Width           =   3495
         Begin VB.ListBox ToSubmit 
            Height          =   1815
            Left            =   240
            TabIndex        =   66
            Top             =   360
            Width           =   3015
         End
      End
      Begin MSComctlLib.Toolbar ToolbarEngines 
         Height          =   330
         Left            =   -74760
         TabIndex        =   64
         Top             =   6360
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   582
         ButtonWidth     =   2487
         ButtonHeight    =   582
         Appearance      =   1
         Style           =   1
         TextAlignment   =   1
         ImageList       =   "imgButt"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   5
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Save    "
               Key             =   "SaveE"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Load    "
               Key             =   "LoadE"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Move    "
               Key             =   "Move"
               ImageIndex      =   34
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Move All    "
               Key             =   "Move All"
               ImageIndex      =   32
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Major Engines"
               Key             =   "Major"
               ImageIndex      =   31
            EndProperty
         EndProperty
      End
      Begin VB.Frame Frame3 
         Caption         =   "Informations"
         Height          =   6375
         Left            =   -67320
         TabIndex        =   54
         Top             =   885
         Width           =   2415
         Begin VB.PictureBox Picture3 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   480
            Index           =   1
            Left            =   120
            Picture         =   "frmMain.frx":0120
            ScaleHeight     =   480
            ScaleWidth      =   480
            TabIndex        =   57
            Top             =   2520
            Width           =   480
         End
         Begin VB.PictureBox Picture3 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   480
            Index           =   0
            Left            =   120
            Picture         =   "frmMain.frx":0562
            ScaleHeight     =   480
            ScaleWidth      =   480
            TabIndex        =   56
            Top             =   240
            Width           =   480
         End
         Begin VB.Label lblLeft 
            Caption         =   "10"
            Height          =   255
            Left            =   120
            TabIndex        =   61
            Top             =   4440
            Width           =   1095
         End
         Begin VB.Label lblSel 
            Caption         =   "0 out of 10"
            Height          =   255
            Left            =   120
            TabIndex        =   60
            Top             =   3480
            Width           =   1815
         End
         Begin VB.Label Label4 
            Caption         =   "# of Engines left:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   59
            Top             =   4080
            Width           =   2175
         End
         Begin VB.Label Label3 
            Caption         =   "# of Selected Engines:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   58
            Top             =   3120
            Width           =   2175
         End
         Begin VB.Label Label2 
            Caption         =   $"frmMain.frx":09A4
            Height          =   1455
            Left            =   120
            TabIndex        =   55
            Top             =   840
            Width           =   2055
         End
      End
      Begin VB.PictureBox Take_Away 
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   3720
         Picture         =   "frmMain.frx":0A3A
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   53
         ToolTipText     =   "Take away all engines."
         Top             =   -500
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox Take 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   3720
         Picture         =   "frmMain.frx":0E7C
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   52
         ToolTipText     =   "Take out Engine."
         Top             =   -500
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Frame Frame2 
         Caption         =   "Search Engines to be use"
         Height          =   5295
         Left            =   -71160
         TabIndex        =   51
         Top             =   885
         Width           =   3495
         Begin VB.ListBox SelectedEngines 
            Height          =   4740
            Left            =   120
            Sorted          =   -1  'True
            TabIndex        =   63
            Top             =   360
            Width           =   3255
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   "Available Search Engines"
         Height          =   5295
         Left            =   -74775
         TabIndex        =   50
         Top             =   885
         Width           =   3495
         Begin VB.ListBox AvailableEngines 
            Height          =   4740
            Left            =   120
            MultiSelect     =   2  'Extended
            Sorted          =   -1  'True
            TabIndex        =   62
            Top             =   360
            Width           =   3255
         End
      End
      Begin VB.PictureBox Picture2 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         Height          =   1740
         Index           =   0
         Left            =   -66240
         ScaleHeight     =   1680
         ScaleWidth      =   1110
         TabIndex        =   49
         Top             =   5280
         Width           =   1170
         Begin VB.Image Image1 
            Height          =   1200
            Index           =   0
            Left            =   120
            Picture         =   "frmMain.frx":12BE
            ToolTipText     =   "Copyright 2000, Casper Semiramis II Production"
            Top             =   240
            Width           =   870
         End
      End
      Begin VB.TextBox txtPhone 
         Height          =   300
         Left            =   9000
         TabIndex        =   46
         ToolTipText     =   "Your phone number with prefix."
         Top             =   6720
         Width           =   855
      End
      Begin VB.TextBox txtPostal 
         Height          =   300
         Left            =   7320
         TabIndex        =   44
         ToolTipText     =   "Your Postal Code or ZIP."
         Top             =   6720
         Width           =   855
      End
      Begin VB.TextBox txtProvince 
         Height          =   300
         Left            =   7320
         TabIndex        =   42
         ToolTipText     =   "Your Province."
         Top             =   6240
         Width           =   2535
      End
      Begin VB.TextBox txtAddress 
         Height          =   300
         Left            =   7320
         TabIndex        =   40
         ToolTipText     =   "Your Address."
         Top             =   5760
         Width           =   2535
      End
      Begin VB.TextBox txtMail 
         Height          =   300
         Left            =   7320
         TabIndex        =   38
         ToolTipText     =   "Your Email."
         Top             =   5280
         Width           =   2535
      End
      Begin VB.TextBox txtCountry 
         Height          =   300
         Left            =   1560
         TabIndex        =   36
         ToolTipText     =   "Your country."
         Top             =   6720
         Width           =   2535
      End
      Begin VB.TextBox txtCity 
         Height          =   300
         Left            =   1560
         TabIndex        =   34
         ToolTipText     =   "City you live in."
         Top             =   6240
         Width           =   2535
      End
      Begin VB.TextBox txtCompany 
         Height          =   300
         Left            =   1560
         TabIndex        =   32
         ToolTipText     =   "Your company name."
         Top             =   5760
         Width           =   2535
      End
      Begin VB.TextBox txtName 
         Height          =   300
         Left            =   1560
         TabIndex        =   30
         ToolTipText     =   "Your full name."
         Top             =   5280
         Width           =   2535
      End
      Begin VB.TextBox k16 
         Height          =   300
         Left            =   5880
         TabIndex        =   27
         Top             =   4560
         Width           =   1095
      End
      Begin VB.TextBox k17 
         Height          =   300
         Left            =   7320
         TabIndex        =   28
         Top             =   4560
         Width           =   1095
      End
      Begin VB.TextBox k18 
         Height          =   300
         Left            =   8760
         TabIndex        =   29
         Top             =   4560
         Width           =   1095
      End
      Begin VB.TextBox k15 
         Height          =   300
         Left            =   4440
         TabIndex        =   26
         Top             =   4560
         Width           =   1095
      End
      Begin VB.TextBox k14 
         Height          =   300
         Left            =   3000
         TabIndex        =   25
         Top             =   4560
         Width           =   1095
      End
      Begin VB.TextBox k13 
         Height          =   300
         Left            =   1560
         TabIndex        =   24
         Top             =   4560
         Width           =   1095
      End
      Begin VB.TextBox k10 
         Height          =   300
         Left            =   5880
         TabIndex        =   21
         Top             =   4200
         Width           =   1095
      End
      Begin VB.TextBox k11 
         Height          =   300
         Left            =   7320
         TabIndex        =   22
         Top             =   4200
         Width           =   1095
      End
      Begin VB.TextBox k12 
         Height          =   300
         Left            =   8760
         TabIndex        =   23
         Top             =   4200
         Width           =   1095
      End
      Begin VB.TextBox k9 
         Height          =   300
         Left            =   4440
         TabIndex        =   20
         Top             =   4200
         Width           =   1095
      End
      Begin VB.TextBox k8 
         Height          =   300
         Left            =   3000
         TabIndex        =   19
         Top             =   4200
         Width           =   1095
      End
      Begin VB.TextBox k7 
         Height          =   300
         Left            =   1560
         TabIndex        =   18
         Top             =   4200
         Width           =   1095
      End
      Begin VB.TextBox k4 
         Height          =   300
         Left            =   5880
         TabIndex        =   15
         Top             =   3840
         Width           =   1095
      End
      Begin VB.TextBox k5 
         Height          =   300
         Left            =   7320
         TabIndex        =   16
         Top             =   3840
         Width           =   1095
      End
      Begin VB.TextBox k6 
         Height          =   300
         Left            =   8760
         TabIndex        =   17
         Top             =   3840
         Width           =   1095
      End
      Begin VB.TextBox k3 
         Height          =   300
         Left            =   4440
         TabIndex        =   14
         Top             =   3840
         Width           =   1095
      End
      Begin VB.TextBox k2 
         Height          =   300
         Left            =   3000
         TabIndex        =   13
         Top             =   3840
         Width           =   1095
      End
      Begin VB.TextBox k1 
         Height          =   300
         Left            =   1560
         TabIndex        =   12
         Top             =   3840
         Width           =   1095
      End
      Begin VB.ComboBox CoCategory 
         Height          =   315
         ItemData        =   "frmMain.frx":1D24
         Left            =   1560
         List            =   "frmMain.frx":1D26
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   3360
         Width           =   8295
      End
      Begin VB.TextBox txtDes 
         Height          =   1380
         Left            =   1560
         MultiLine       =   -1  'True
         TabIndex        =   8
         ToolTipText     =   "Brief description. Not more then 264 characters."
         Top             =   1800
         Width           =   8295
      End
      Begin VB.TextBox txtTitle 
         Height          =   300
         Left            =   1560
         TabIndex        =   6
         ToolTipText     =   "Title of your web site"
         Top             =   1320
         Width           =   8295
      End
      Begin VB.TextBox txtURL 
         Height          =   300
         Left            =   2760
         TabIndex        =   4
         ToolTipText     =   "Address of Your Web Site"
         Top             =   840
         Width           =   7095
      End
      Begin VB.ComboBox CoUrl 
         Height          =   315
         Left            =   1560
         TabIndex        =   3
         Text            =   "http://"
         Top             =   840
         Width           =   1095
      End
      Begin MSComctlLib.ImageList imgButt 
         Left            =   120
         Top             =   6960
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   34
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1D28
               Key             =   "New"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1E3A
               Key             =   "Open"
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1F4C
               Key             =   "Save"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":205E
               Key             =   "Print"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2170
               Key             =   "Cut"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2282
               Key             =   "Copy"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2394
               Key             =   "Paste"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":24A6
               Key             =   "Bold"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":25B8
               Key             =   "Italic"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":26CA
               Key             =   "Underline"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":27DC
               Key             =   "Align Left"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":28EE
               Key             =   "Center"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2A00
               Key             =   "Align Right"
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2B12
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":3246
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":369A
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":3AEE
               Key             =   ""
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":3F42
               Key             =   ""
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":4396
               Key             =   ""
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":47EA
               Key             =   ""
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":4C3E
               Key             =   ""
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":5092
               Key             =   ""
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":54EA
               Key             =   ""
            EndProperty
            BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":5942
               Key             =   ""
            EndProperty
            BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":5D96
               Key             =   ""
            EndProperty
            BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":61EA
               Key             =   ""
            EndProperty
            BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":6782
               Key             =   ""
            EndProperty
            BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":6A6A
               Key             =   ""
            EndProperty
            BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":6EBE
               Key             =   ""
            EndProperty
            BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":7312
               Key             =   ""
            EndProperty
            BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":7766
               Key             =   ""
            EndProperty
            BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":7BBA
               Key             =   ""
            EndProperty
            BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":800E
               Key             =   ""
            EndProperty
            BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":8462
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "(At least two)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   20
         Left            =   360
         TabIndex        =   88
         Top             =   4200
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "(264 Characters)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   255
         Index           =   19
         Left            =   240
         TabIndex        =   87
         Top             =   2160
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "In the meantime, be so nice and visit our sponsor by clicking bellow on banner. This application is a shareware so keep it free."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -70920
         TabIndex        =   77
         Top             =   4560
         Width           =   5775
      End
      Begin VB.Label Label6 
         Caption         =   "It may take  a few minutes before all the engines will be reached and submit to."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -70920
         TabIndex        =   76
         Top             =   4080
         Width           =   5775
      End
      Begin VB.Label Label1 
         Caption         =   "Phone:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   13
         Left            =   8280
         TabIndex        =   47
         Top             =   6720
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Postal Code:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   6000
         TabIndex        =   45
         Top             =   6720
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Province:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   6330
         TabIndex        =   43
         Top             =   6240
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Address:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   6330
         TabIndex        =   41
         Top             =   5760
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Email:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   6555
         TabIndex        =   39
         Top             =   5280
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Country:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   8
         Left            =   600
         TabIndex        =   37
         Top             =   6720
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "City:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   7
         Left            =   960
         TabIndex        =   35
         Top             =   6240
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Company:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   6
         Left            =   480
         TabIndex        =   33
         Top             =   5760
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Your Name:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   5
         Left            =   360
         TabIndex        =   31
         Top             =   5280
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Keywords:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   4
         Left            =   360
         TabIndex        =   11
         Top             =   3840
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Category:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   400
         TabIndex        =   9
         Top             =   3360
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Description:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   260
         TabIndex        =   7
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Site Title:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   460
         TabIndex        =   5
         Top             =   1320
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Web Site:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   430
         TabIndex        =   2
         Top             =   840
         Width           =   1095
      End
   End
   Begin VB.Frame Frame9 
      Height          =   2655
      Left            =   120
      TabIndex        =   91
      Top             =   5760
      Width           =   9615
      Begin SHDocVwCtl.WebBrowser Web 
         Height          =   2295
         Left            =   120
         TabIndex        =   92
         Top             =   240
         Visible         =   0   'False
         Width           =   9375
         ExtentX         =   16536
         ExtentY         =   4048
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
         Location        =   ""
      End
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&Main"
      Begin VB.Menu mnuProfile 
         Caption         =   "&Select Profile"
      End
      Begin VB.Menu mnuEdit 
         Caption         =   "&Edit Profile"
      End
      Begin VB.Menu sep 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuU 
      Caption         =   "Update"
      Begin VB.Menu mnuCheck 
         Caption         =   "Check for new version"
      End
      Begin VB.Menu mnuUpdate 
         Caption         =   "&Live Update"
         Enabled         =   0   'False
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
      Begin VB.Menu mnuA 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim ExitSubmition As Integer
Dim TUpdate_I As Integer
Dim EnginesCount As Integer

Private Sub Form_Load()
TUpdate_I = 0

' We need to fill up some of the drop-down combos _
  e.g. Category, head etc.
CoUrl.AddItem "http://"
CoUrl.AddItem "https://"
LoadCategories
LoadAvailableEngines

' For now without a link to web.
picBanner.FileName = App.Path & "\data\banner.gif"

'Start the Report:
Rtf.FileName = App.Path & "\data\rtfload.txt"

End Sub

Private Sub Form_Unload(Cancel As Integer)
'Is it changed ?
If Not txtURL.Text = "" Then

' Do we want to save the profile?
    If MsgBox("Would you like to save the profile?", vbQuestion + vbYesNo, "Save Profile?") = vbYes Then
        
        'Yes
        frmProfiles.Show 1, frmMain
      Else
        
        'No
    End If

End If
End Sub

Private Sub mnuA_Click()
' Easiest way how to make the form stay on top: frmName Modal, Owner
frmAbout.Show 1, frmMain
End Sub

Private Sub mnuCheck_Click()
CheckForNewVersionOnTheNet
End Sub

Private Sub mnuEdit_Click()
frmProfiles.Show 1, frmMain
End Sub

Private Sub mnuProfile_Click()
frmProfiles.Show 1, frmMain
End Sub

Private Sub mnuUpdate_Click()
' This will initialize the download dialog.
' It works togather with Check For update menu item so check
' it out befor you gonna do anything with it.
'
' File is posted it on  my website so PLEASEEEE !!!!!
' CHANGE THAT URL AS SOON AS YOU GONNA USE IT FOR REAL UPDATE.
Scrap.WebDownload.Navigate "http://www.bohemiatrading.com/download/tools/Tool.exe"
End Sub

Private Sub Toolbar_ButtonClick(ByVal Button As MSComctlLib.Button)
' Those are Tool Bar & Menu Functions
On Error Resume Next
Select Case Button.Key
    Case "Clear"
ClearAllFields
    Case "Load/Save"
 frmProfiles.Show 1, frmMain
    Case "Exit"
ExitApp
    End Select
End Sub

Private Sub mnuExit_Click()
 ExitApp
End Sub

Private Sub Add_Click()
' Rearanger The engines from one obj to another
Ind = SelectedEngines.ListIndex
SelectedEngines.AddItem AvailableEngines.Text
Call LBDupe(SelectedEngines)
lblSel.Caption = SelectedEngines.ListCount & " out of 8"
lblLeft.Caption = AvailableEngines.ListCount - SelectedEngines.ListCount
ToSubmit.AddItem AvailableEngines.Text
E_Left.Caption = ToSubmit.ListCount
End Sub

Private Sub Add_all_Click()
' Loads the list of categories
SelectedEngines.AddItem "Altavista.com"
SelectedEngines.AddItem "AllAmericasBest.com"
SelectedEngines.AddItem "DirectHit.com"
SelectedEngines.AddItem "EuroSeek.com"
SelectedEngines.AddItem "Excite.com"
SelectedEngines.AddItem "HotBot.com"
SelectedEngines.AddItem "InfoSeek.com"
SelectedEngines.AddItem "Lycos.com"
SelectedEngines.AddItem "MSN.com"
SelectedEngines.AddItem "WebCrawler.com"

'Other LisBox
ToSubmit.AddItem "Altavista.com"
ToSubmit.AddItem "AllAmericasBest.com"
ToSubmit.AddItem "DirectHit.com"
ToSubmit.AddItem "EuroSeek.com"
ToSubmit.AddItem "Excite.com"
ToSubmit.AddItem "HotBot.com"
ToSubmit.AddItem "InfoSeek.com"
ToSubmit.AddItem "Lycos.com"
SelectedEngines.AddItem "MSN.com"
SelectedEngines.AddItem "WebCrawler.com"

' Make sure that there are no double entries
Call LBDupe(ToSubmit)
Call LBDupe(SelectedEngines)
lblSel.Caption = ""
lblSel.Caption = SelectedEngines.ListCount & " out of 10"
lblLeft.Caption = 0
E_Left.Caption = ToSubmit.ListCount
End Sub
Private Sub AvailableEngines_DblClick()
' Double Click Add Engines
 Add_Click
End Sub
Private Sub cmdCancel_Click()
 ExitSubmition = 0
End Sub
Private Sub ToolbarEngines_ButtonClick(ByVal Button As MSComctlLib.Button)
 On Error Resume Next
Select Case Button.Key
    Case "Move"
Add_Click
    Case "Move All"
Add_all_Click
    Case "Major"
'Temporarly
Add_all_Click
    End Select
End Sub
Private Sub Command1_Click() '(sorry about the name)
' This is where most of the work is done. This sub sends info to the
' search engine, then waits for 5 seconds (Has to wait!!!) and skips
' to next engine.
' Between there are simple statistics (JunkAndStats).
' I'm plan on using ini files in the next version, but for now let's
' just use these calls.


' We need to check at least for two main components of submiting.
' Those are: WebSite URL (txtURL.Text) and e-mail (txtMail.Text).
If txtURL.Text = "" Then
    MsgBox "WebSite Data Missing!"
    SSTab1.Tab = 0
    Exit Sub
End If

If txtMail.Text = "" Then
    MsgBox "WebSite Data Missing!"
    SSTab1.Tab = 0
    Exit Sub
End If


'Dims and stuff
Dim I As Integer, IndexNum As Integer, ExitSubmition As Integer
IndexNum = ToSubmit.ListCount

' In case there is no engines, exit sub.
    If IndexNum = -1 Then
      MsgBox "No selection has been made."
      SSTab1.Tab = 1
      Exit Sub
    End If

' What procentage should we add for each engine ?
' EnginesCount is and integer so no decimal junk.
EnginesCount = (100 / ToSubmit.ListCount)

For I = 0 To IndexNum

 If ToSubmit.List(I) = "Altavista.com" Then
   lblEngine.Caption = ToSubmit.List(I)
   Web.Navigate "http://add-url.altavista.digital.com/cgi-bin/newurl?ad=1&q=http%3A%2F%2F" & frmMain.txtURL.Text
JunkAndStats
End If

 If ToSubmit.List(I) = "AllAmericasBest.com" Then
   lblEngine.Caption = ToSubmit.List(I)
   Web.Navigate "http://www.allamericasbest.com/cgi-local/addurl.pl?Name?" & frmMain.CoUrl.Text & frmMain.txtURL.Text & "?Des?" & frmMain.k1.Text & "?" & frmMain.txtMail.Text & "?" & frmMain.CoCategory.Text
JunkAndStats
 End If

 If ToSubmit.List(I) = "DirectHit.com" Then
   lblEngine.Caption = ToSubmit.List(I)
   Web.Navigate "http://www.directhit.com/fcgi-bin/DirectHitWeb.fcg?fmt=disp&template=addurl&src=DH_ADDURL&URL=http%3A%2F%2F" & frmMain.txtURL.Text & "&email=" & frmMain.txtMail.Text & "&keys=" & frmMain.k1.Text & "," & frmMain.k2.Text & "&submit=Submit%21"
JunkAndStats
 End If
 
 If ToSubmit.List(I) = "EuroSeek.com" Then
   lblEngine.Caption = ToSubmit.List(I)
   Web.Navigate "http://addsite.euroseek.com/page.php?url=" & frmMain.CoUrl.Text & frmMain.txtURL.Text & "?"
JunkAndStats
 End If

 If ToSubmit.List(I) = "Excite.com" Then
   lblEngine.Caption = ToSubmit.List(I)
   Web.Navigate "http://www.excite.com/info/add_url/thanks/?url=" & frmMain.CoUrl.Text & frmMain.txtURL.Text & "&email=" & frmMain.txtMail.Text & "&country=US&brand=excite"
JunkAndStats
 End If

 If ToSubmit.List(I) = "HotBot.com" Then
   lblEngine.Caption = ToSubmit.List(I)
   Web.Navigate "http://hotbot.lycos.com/addurl.asp?MM=1&success_page=http%3A%2F%2Fhotbot.lycos.com%2Faddurl.asp&failure_page=http%3A%2F%2Fhotbot.lycos.com%2Fhelp%2Foops.asp&ACTION=subscribe&SOURCE=hotbot&ip=24.66.63.44&redirect=http%3A%2F%2Fhotbot.lycos.com%2Faddurl2.html&newurl=http%3A%2F%2F" & frmMain.txtURL.Text & "&email=" & frmMain.txtMail.Text & "&send=Submit+my+site"
JunkAndStats
 End If

 If ToSubmit.List(I) = "InfoSeek.com" Then
   lblEngine.Caption = ToSubmit.List(I)
   Web.Navigate "http://www.go.com/AddUrl/AddingURL?url=http%3A%2F%2F" & frmMain.txtURL.Text & "&CAT=Add%2FUpdate+Site&sv=AD&lk=noframes"
JunkAndStats
 End If

 If ToSubmit.List(I) = "Lycos.com" Then
   lblEngine.Caption = ToSubmit.List(I)
   Web.Navigate "http://www.lycos.com/cgi-bin/spider_now.pl?query=http%3A%2F%2F" & frmMain.txtURL.Text & "&email=" & frmMain.txtMail.Text
JunkAndStats
 End If

 If ToSubmit.List(I) = "MSN.com" Then
   lblEngine.Caption = ToSubmit.List(I)
   Web.Navigate "http://submitit.linkexchange.com/system/msnaddpublicbeta.cfm?delete=0&localetag=en-us&url=http%3A%2F%2F" & frmMain.txtURL.Text & "&category=1266857&email=" & frmMain.txtMail.Text & "&title=" & frmMain.txtTitle.Text & "&description=" & frmMain.k1.Text & "+" & frmMain.k2.Text & "+" & frmMain.k3.Text & "+"
JunkAndStats
 End If

 If ToSubmit.List(I) = "WebCrawler.com" Then
   lblEngine.Caption = ToSubmit.List(I)
   Web.Navigate "http://www.webcrawler.com/info/add_url/thanks/?url=" & frmMain.CoUrl.Text & frmMain.txtURL.Text & "&email=" & frmMain.txtMail.Text & "&country=US&brand=webcrawler"
JunkAndStats
 End If
 



Next I

' On the end we need to fix some labels.
    Succes.Caption = "All Done!"
    lblEngine.Caption = "None"
    Procent.Caption = "100"

End Sub

Function JunkAndStats()
   Succes.Caption = "Working ..."
Call WaitASec(5)
   Submited.AddItem ToSubmit.List(I)
   ToSubmit.RemoveItem (I)
   E_Left.Caption = ToSubmit.ListCount
   E_done.Caption = Submited.ListCount
   Procent.Caption = Procent.Caption + EnginesCount
   Succes.Caption = "Done"
' Report Creation:
   Rtf.Text = Rtf.Text & ToSubmit.List(I) & ":" & vbCrLf
   Rtf.Text = Rtf.Text & "Completed sucessfuly." & vbCrLf & vbCrLf



End Function

VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6660
   ClientLeft      =   132
   ClientTop       =   444
   ClientWidth     =   9504
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6660
   ScaleWidth      =   9504
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   348
      Left            =   0
      TabIndex        =   3
      Top             =   6312
      Width           =   9504
      _ExtentX        =   16764
      _ExtentY        =   614
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   14995
            MinWidth        =   14995
            Text            =   "Searching Is Ready..."
            TextSave        =   "Searching Is Ready..."
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   2
            Enabled         =   0   'False
            Object.Width           =   882
            MinWidth        =   882
            TextSave        =   "CAPS"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Stop Searching"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   1400
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   4080
      Width           =   1455
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   288
      Left            =   3000
      TabIndex        =   12
      Top             =   1908
      Width           =   6600
      _ExtentX        =   11642
      _ExtentY        =   508
      ButtonWidth     =   487
      ButtonHeight    =   466
      AllowCustomize  =   0   'False
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList3"
      DisabledImageList=   "ImageList4"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   12
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Icons"
            ImageIndex      =   1
            Style           =   2
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "List"
            ImageIndex      =   2
            Style           =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Report"
            ImageIndex      =   3
            Style           =   2
            Value           =   1
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Smallicons"
            ImageIndex      =   2
            Style           =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   150
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Cut"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Copy"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Paste"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   4
            Object.Width           =   150
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Refresh"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Recycle Bin"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Delete"
            ImageIndex      =   8
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.DirListBox Dir1 
      Height          =   936
      Left            =   7680
      TabIndex        =   0
      Top             =   4320
      Visible         =   0   'False
      Width           =   1572
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4560
      Top             =   5160
   End
   Begin MSComctlLib.ImageList ImageList5 
      Left            =   6240
      Top             =   5640
      _ExtentX        =   804
      _ExtentY        =   804
      BackColor       =   -2147483643
      ImageWidth      =   26
      ImageHeight     =   26
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":076A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.ListBox List3 
      Height          =   624
      Left            =   6600
      TabIndex        =   17
      Top             =   5400
      Visible         =   0   'False
      Width           =   972
   End
   Begin VB.ListBox List1 
      Height          =   624
      Left            =   6480
      TabIndex        =   13
      Top             =   5280
      Visible         =   0   'False
      Width           =   972
   End
   Begin MSComctlLib.ImageList ImageList4 
      Left            =   5760
      Top             =   5640
      _ExtentX        =   804
      _ExtentY        =   804
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0FDE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1378
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1712
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1AAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1E46
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":21E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":257A
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2914
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2CAE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox Text4 
      Height          =   372
      Left            =   3240
      TabIndex        =   14
      Text            =   "Text4"
      Top             =   5760
      Visible         =   0   'False
      Width           =   972
   End
   Begin MSComctlLib.ImageList ImageList3 
      Left            =   5280
      Top             =   5640
      _ExtentX        =   804
      _ExtentY        =   804
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3048
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":33E2
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":377C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3B16
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3EB0
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":424A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":45E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":497E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4D18
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   555
      Left            =   5040
      ScaleHeight     =   504
      ScaleWidth      =   504
      TabIndex        =   11
      Top             =   5040
      Visible         =   0   'False
      Width           =   555
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   4800
      Top             =   5640
      _ExtentX        =   804
      _ExtentY        =   804
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   -2147483633
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4320
      Top             =   5640
      _ExtentX        =   804
      _ExtentY        =   804
      BackColor       =   -2147483643
      MaskColor       =   -2147483633
      _Version        =   393216
   End
   Begin VB.ListBox List2 
      Height          =   624
      Left            =   120
      TabIndex        =   29
      Top             =   1200
      Visible         =   0   'False
      Width           =   2532
   End
   Begin VB.TextBox Text5 
      Height          =   852
      Left            =   7800
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   30
      Top             =   5280
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00FFFFFF&
      Height          =   4008
      Left            =   3000
      ScaleHeight     =   3960
      ScaleWidth      =   6552
      TabIndex        =   19
      Top             =   2250
      Width           =   6600
      Begin VB.CommandButton Command7 
         Height          =   240
         Index           =   3
         Left            =   5280
         TabIndex        =   22
         Top             =   0
         Width           =   1250
      End
      Begin VB.CommandButton Command7 
         Caption         =   "File Path                                              "
         Height          =   240
         Index           =   0
         Left            =   1450
         TabIndex        =   28
         Top             =   0
         Width           =   3012
      End
      Begin VB.CommandButton Command7 
         Caption         =   "File Size    "
         Height          =   240
         Index           =   2
         Left            =   4440
         TabIndex        =   21
         Top             =   0
         Width           =   852
      End
      Begin VB.CommandButton Command7 
         Caption         =   "File Name             "
         Height          =   240
         Index           =   1
         Left            =   0
         TabIndex        =   20
         Top             =   0
         Width           =   1452
      End
      Begin VB.VScrollBar VScroll1 
         Height          =   3732
         LargeChange     =   1200
         Left            =   6300
         Max             =   0
         SmallChange     =   200
         TabIndex        =   23
         Top             =   240
         Width           =   252
      End
      Begin VB.PictureBox Picture2 
         BackColor       =   &H00FFFFFF&
         Height          =   3972
         Left            =   0
         ScaleHeight     =   3924
         ScaleWidth      =   6324
         TabIndex        =   24
         Top             =   0
         Width           =   6372
         Begin VB.ListBox List4 
            Height          =   624
            Left            =   3360
            TabIndex        =   32
            Top             =   3000
            Visible         =   0   'False
            Width           =   972
         End
         Begin VB.Label Label4 
            BackStyle       =   0  'Transparent
            Caption         =   "Label4"
            Height          =   192
            Index           =   0
            Left            =   240
            TabIndex        =   27
            Top             =   240
            Visible         =   0   'False
            Width           =   1116
         End
         Begin VB.Label Label5 
            BackStyle       =   0  'Transparent
            Caption         =   "Label5"
            Height          =   192
            Index           =   0
            Left            =   1440
            TabIndex        =   26
            Top             =   240
            Visible         =   0   'False
            Width           =   3012
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Label6"
            Height          =   192
            Index           =   0
            Left            =   4476
            TabIndex        =   25
            Top             =   240
            Visible         =   0   'False
            Width           =   492
         End
      End
   End
   Begin VB.CommandButton Command2 
      Height          =   6255
      Left            =   2880
      TabIndex        =   18
      Top             =   0
      Width           =   84
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3990
      Left            =   3000
      TabIndex        =   2
      Top             =   2220
      Visible         =   0   'False
      Width           =   6600
      _ExtentX        =   11642
      _ExtentY        =   7049
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      MultiSelect     =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   -1  'True
      GridLines       =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList2"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   7.8
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Files/Folder Information."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   178
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1900
      Left            =   3000
      TabIndex        =   5
      Top             =   0
      Width           =   6600
      Begin VB.PictureBox Picture9 
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   1332
         Left            =   50
         ScaleHeight     =   1332
         ScaleWidth      =   6372
         TabIndex        =   47
         Top             =   240
         Visible         =   0   'False
         Width           =   6372
         Begin VB.PictureBox Picture15 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   180
            Left            =   0
            Picture         =   "Form1.frx":50B2
            ScaleHeight     =   180
            ScaleWidth      =   132
            TabIndex        =   59
            Top             =   1050
            Width           =   132
         End
         Begin VB.PictureBox Picture14 
            AutoRedraw      =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            DrawWidth       =   3
            FillStyle       =   0  'Solid
            Height          =   492
            Left            =   3360
            ScaleHeight     =   492
            ScaleWidth      =   492
            TabIndex        =   57
            Top             =   840
            Width           =   492
         End
         Begin VB.PictureBox Picture13 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   180
            Left            =   0
            Picture         =   "Form1.frx":5310
            ScaleHeight     =   180
            ScaleWidth      =   132
            TabIndex        =   56
            Top             =   820
            Width           =   132
         End
         Begin VB.PictureBox Picture12 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   180
            Left            =   0
            Picture         =   "Form1.frx":556E
            ScaleHeight     =   180
            ScaleWidth      =   132
            TabIndex        =   55
            Top             =   560
            Width           =   132
         End
         Begin VB.PictureBox Picture11 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   180
            Left            =   0
            Picture         =   "Form1.frx":57CC
            ScaleHeight     =   180
            ScaleWidth      =   132
            TabIndex        =   54
            Top             =   300
            Width           =   132
         End
         Begin VB.PictureBox Picture10 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   180
            Left            =   0
            Picture         =   "Form1.frx":5A2A
            ScaleHeight     =   180
            ScaleWidth      =   132
            TabIndex        =   53
            Top             =   50
            Width           =   132
         End
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Found (0) Item(s) So Far"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   216
            Left            =   240
            TabIndex        =   58
            Top             =   1080
            Width           =   2016
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BackStyle       =   0  'Transparent
            Caption         =   "Label15"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   216
            Left            =   240
            TabIndex        =   52
            Top             =   840
            Width           =   588
         End
         Begin VB.Label Label14 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Searching Now..."
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   16.8
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   396
            Left            =   3840
            TabIndex        =   51
            Top             =   876
            Width           =   2412
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Label13"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   216
            Left            =   240
            TabIndex        =   50
            Top             =   600
            Width           =   588
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Label12"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   216
            Left            =   240
            TabIndex        =   49
            Top             =   360
            Width           =   588
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Label11"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   216
            Left            =   240
            TabIndex        =   48
            Top             =   80
            Width           =   588
         End
      End
      Begin VB.HScrollBar HScroll1 
         Height          =   200
         LargeChange     =   1200
         Left            =   50
         Max             =   0
         SmallChange     =   200
         TabIndex        =   44
         Top             =   1650
         Visible         =   0   'False
         Width           =   6550
      End
      Begin VB.PictureBox Picture4 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         Height          =   1452
         Left            =   50
         ScaleHeight     =   1452
         ScaleWidth      =   6492
         TabIndex        =   36
         Top             =   240
         Visible         =   0   'False
         Width           =   6490
         Begin VB.PictureBox Picture8 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            Height          =   450
            Left            =   50
            ScaleHeight     =   456
            ScaleWidth      =   516
            TabIndex        =   45
            Top             =   20
            Width           =   510
         End
         Begin VB.PictureBox Picture7 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   180
            Left            =   0
            Picture         =   "Form1.frx":5C88
            ScaleHeight     =   180
            ScaleWidth      =   132
            TabIndex        =   39
            Top             =   1020
            Width           =   132
         End
         Begin VB.PictureBox Picture6 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   180
            Left            =   0
            Picture         =   "Form1.frx":5EE6
            ScaleHeight     =   180
            ScaleWidth      =   132
            TabIndex        =   38
            Top             =   756
            Width           =   132
         End
         Begin VB.PictureBox Picture5 
            AutoRedraw      =   -1  'True
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   180
            Left            =   0
            Picture         =   "Form1.frx":6144
            ScaleHeight     =   180
            ScaleWidth      =   132
            TabIndex        =   37
            Top             =   480
            Width           =   132
         End
         Begin VB.Label Label10 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "4"
            Height          =   192
            Left            =   216
            TabIndex        =   43
            Top             =   1080
            Width           =   84
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "3"
            Height          =   192
            Left            =   216
            TabIndex        =   42
            Top             =   816
            Width           =   84
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "2"
            ForeColor       =   &H00FF0000&
            Height          =   192
            Left            =   216
            MouseIcon       =   "Form1.frx":63A2
            MousePointer    =   99  'Custom
            TabIndex        =   41
            Top             =   516
            Width           =   84
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "1"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   178
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   600
            TabIndex        =   40
            Top             =   195
            Width           =   105
         End
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "No Item Found To View!"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   288
         Left            =   120
         TabIndex        =   61
         Top             =   960
         Width           =   2604
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Click On Any Item To View The Information Or RightClick  To Popup Menu"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   612
         Left            =   120
         TabIndex        =   60
         Top             =   360
         Visible         =   0   'False
         Width           =   5724
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Height          =   6255
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   2892
      Begin VB.CommandButton Command3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Browse"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   10.8
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   370
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   34
         Top             =   3360
         Width           =   770
      End
      Begin VB.TextBox Text3 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   33
         Text            =   "c:\"
         Top             =   3360
         Width           =   1812
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Search Now"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   380
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   31
         Top             =   4080
         Width           =   1188
      End
      Begin VB.Frame Frame4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Searching In The Computer."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2532
         Left            =   20
         TabIndex        =   6
         Top             =   360
         Width           =   2799
         Begin VB.TextBox Text1 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   372
            Left            =   120
            TabIndex        =   9
            Top             =   480
            Width           =   2172
         End
         Begin VB.Frame Frame2 
            BackColor       =   &H00FFFFFF&
            Caption         =   "Options"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1536
            Left            =   60
            TabIndex        =   7
            Top             =   840
            Width           =   2652
            Begin VB.CheckBox Check1 
               BackColor       =   &H00FFFFFF&
               Caption         =   "Serche Hidden Folders and files"
               Height          =   372
               Left            =   120
               TabIndex        =   46
               Top             =   1080
               Width           =   2412
            End
            Begin VB.ComboBox Combo3 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   7.8
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   288
               Left            =   120
               TabIndex        =   16
               Text            =   "Files Only"
               Top             =   240
               Width           =   2172
            End
            Begin VB.ComboBox Combo1 
               BeginProperty Font 
                  Name            =   "Tahoma"
                  Size            =   7.8
                  Charset         =   178
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   288
               Left            =   120
               TabIndex        =   8
               Text            =   "All Files"
               Top             =   600
               Width           =   2172
            End
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackColor       =   &H00FFFFFF&
            Caption         =   "File Or Folder Name:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   7.8
               Charset         =   178
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   192
            Left            =   96
            TabIndex        =   10
            Top             =   240
            Width           =   1440
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Loking In:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   7.8
            Charset         =   178
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   192
         Left            =   120
         TabIndex        =   35
         Top             =   3120
         Width           =   672
      End
   End
   Begin VB.FileListBox File1 
      Height          =   648
      Left            =   840
      TabIndex        =   1
      Top             =   5400
      Visible         =   0   'False
      Width           =   1332
   End
   Begin VB.Menu Menu 
      Caption         =   "MEN"
      Visible         =   0   'False
      Begin VB.Menu opn 
         Caption         =   "Open"
      End
      Begin VB.Menu Line1 
         Caption         =   "-"
      End
      Begin VB.Menu CT 
         Caption         =   "Cut"
      End
      Begin VB.Menu Cpy 
         Caption         =   "Copy"
      End
      Begin VB.Menu Line2 
         Caption         =   "-"
      End
      Begin VB.Menu Rename 
         Caption         =   "Rename"
      End
      Begin VB.Menu Rec 
         Caption         =   "Recycle Bin"
      End
      Begin VB.Menu Dlt 
         Caption         =   "Delet"
      End
      Begin VB.Menu Line3 
         Caption         =   "-"
      End
      Begin VB.Menu Prop 
         Caption         =   "Properties"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Hi, This is my second edition of professional search, if you see the old one
'you'll find many new features in this version addition to new code,anyway,
'i'v tested this version on WinXp and found it ok :-)!!
'anyway if any part of this code is not clear for you or there is any
'bugs please feedback on <saidseyam@hotmail.com>

'if you like it don't forget to vote for me ;-)!!

Dim S(20000) As Integer
Dim M As Integer
Dim FoundFldr As Boolean
Dim BIcon As Integer
Dim SIcon As Integer
Dim FilePath As String
Dim filename As String
Dim StopSearch As Boolean
Dim FS As New FileSystemObject
Dim FI As Scripting.File
Dim Foldr As Scripting.Folder

Dim FilesDelet As String
Dim ItemSelected As Boolean
Dim MouseButton As Integer
Dim FolderName As String
Dim Hi As Integer
Dim Flg As Integer
Dim CutF As Boolean
Dim Clicks As Integer


Dim SX As Single
Dim V As Integer
Dim SX2 As Single
Dim V2 As Integer
Dim SX3 As Single
Dim V3 As Integer

Dim BX As Integer
Dim B2 As Integer


Option Compare Text

Private Sub SubFolders()
maindir = Dir1.Path

For i = 0 To Dir1.ListCount - 1
If StopSearch = True Then Exit For
Dir1.Path = Dir1.List(i)

StatusBar1.Panels(1).Text = "Searching In: " & Dir1.Path & "  (And Sub Folders...)"
SearchType
If StopSearch = True Then Exit For
nf:
If Dir1.ListCount > 0 Then

For F = S(M) To Dir1.ListCount - 1
  'List1.AddItem Dir1.List(f)
  If StopSearch = True Then Exit For
  If Dir1.List(F) <> "" Then Dir1.Path = Dir1.List(F)

SearchType
If Dir1.ListCount = 0 Then
S(M) = F + 1
M = M + 1
n:
  Dir1.Path = UP(Dir1.Path)
  If Dir1.Path = maindir Then GoTo re:
  If Dir1.ListCount < 2 Or S(M - 1) = Dir1.ListCount Then: S(M) = 0: M = M - 1: S(M) = 0: GoTo n:
  If M > 0 Then M = M - 1
  GoTo nf:
  End If
  
  S(M) = F + 1
  F = 0
  M = M + 1
  F = 0

If Dir1.ListCount = 0 Then GoTo n:
 
 GoTo nf:
 DoEvents
Next F
End If

re:
Dir1.Path = maindir
DoEvents
Next i

If StopSearch = False Then StopSearching
StopSearch = True
End Sub
Public Function UP(DIRE As String)
Dim r As Integer
r = InStrRev(DIRE, "\")
UP = Mid(DIRE, 1, r - 1)
If Len(UP) < 3 Then UP = UP & "\"

End Function

Private Sub Searching()
Dim File As Integer

For File = 0 To File1.ListCount - 1
If StopSearch = True Then Exit For

If InStr(1, File1.List(File), Text1, vbTextCompare) Then

 If Right(File1.Path, 1) = "\" Then
  FilePath = File1.Path & File1.List(File)
  Else: FilePath = File1.Path & "\" & File1.List(File)
 End If

filename = File1.List(File)
Temp
End If

DoEvents
Next File

File = 0
End Sub



Private Sub Temp()
Dim S As ListItem
Dim Siz, Fsiz
On Error Resume Next

BIcon = BIcon + 1
SIcon = SIcon + 1

Timer1.Enabled = True
Timer1.Interval = 1
If Timer1.Interval < 30 Then Timer1.Interval = Timer1.Interval + 1 Else Timer1.Enabled = False
Label16 = "Find (" & BIcon & ") Item(s) So Far"

FilesNum = FilesNum + 1
Load Label4(BIcon)
Load Label5(BIcon)
Load Label6(BIcon)

Label4(BIcon).Top = CLng(BIcon) * Hi
Label5(BIcon).Top = CLng(BIcon) * Hi
Label6(BIcon).Top = CLng(BIcon) * Hi

Label4(BIcon).Visible = True
Label5(BIcon).Visible = True
Label6(BIcon).Visible = True

  rGetIcon = ExtractAssociatedIcon(0, FilePath, 1)
  Set Picture1.Picture = Nothing
  DrawIcon Picture1.hdc, 0, 0, rGetIcon
  Picture1.Picture = Picture1.Image
  ImageList1.ListImages.Add , , Picture1.Picture
  ImageList2.ListImages.Add , , Picture1.Picture
  Label4(BIcon) = filename
  Label5(BIcon) = FilePath
  
  ImageList2.ListImages.Item(BIcon).Draw Picture2.hdc, 20, CLng(BIcon) * Hi, 1

  Hi = 200
  If (CLng(BIcon) * Hi) > (3972 + (Form1.ScaleHeight - 6620)) And Picture2.Height < 32000 Then VScroll1.Max = CLng((CLng(BIcon) * CLng(Hi)) - CLng(3972 + (Form1.ScaleHeight - 6620)))
  
  Set FI = FS.GetFile(FilePath)
  Fsiz = FI.Size \ 1000
  Siz = IIf(Fsiz <= 1000, CCur(FI.Size / 1024) & " KB", CCur(FI.Size / 1050000) & " MB")
  Label6(BIcon) = Siz

If CLng(Hi) * BIcon > (3972 + (Form1.ScaleHeight - 6620)) Then Picture2.Height = (CLng(Hi) * BIcon) + (Form1.ScaleHeight - 6620)
  If Picture2.Height < 32000 Then VScroll1.Max = Picture2.Height - (3972 + (Form1.ScaleHeight - 6620))


End Sub

Private Sub TempFolders()
Dim S As ListItem
Dim Siz, Fsiz
On Error Resume Next
BIcon = BIcon + 1
SIcon = SIcon + 1

Timer1.Enabled = True
Timer1.Interval = 1
If Timer1.Interval < 30 Then Timer1.Interval = Timer1.Interval + 1 Else Timer1.Enabled = False
Label16 = "Find (" & BIcon & ") Item(s) So Far"

FilesNum = FilesNum + 1
Load Label4(BIcon)
Load Label5(BIcon)
Load Label6(BIcon)

Label4(BIcon).Top = CLng(BIcon) * Hi
Label5(BIcon).Top = CLng(BIcon) * Hi
Label6(BIcon).Top = CLng(BIcon) * Hi

Label4(BIcon).Visible = True
Label5(BIcon).Visible = True
Label6(BIcon).Visible = True

  rGetIcon = ExtractAssociatedIcon(0, FilePath, 1)
  Set Picture1.Picture = Nothing
  DrawIcon Picture1.hdc, 0, 0, rGetIcon
  Picture1.Picture = Picture1.Image
  ImageList1.ListImages.Add , , Picture1.Picture
  ImageList2.ListImages.Add , , Picture1.Picture
  Label4(BIcon) = FolderName
  Label5(BIcon) = FilePath
  Label6(BIcon).Caption = ""
  ImageList2.ListImages.Item(BIcon).Draw Picture2.hdc, 20, CLng(BIcon) * Hi, 1
  Hi = 200
  If (CLng(BIcon) * Hi) > (3972 + (Form1.ScaleHeight - 6620)) And Picture2.Height < 32000 Then VScroll1.Max = CLng((CLng(BIcon) * CLng(Hi)) - CLng(3972 + (Form1.ScaleHeight - 6620)))
  

If CLng(Hi) * BIcon > (3972 + (Form1.ScaleHeight - 6620)) Then Picture2.Height = (CLng(Hi) * BIcon + (Form1.ScaleHeight - 6620))
  If Picture2.Height < 32000 Then VScroll1.Max = Picture2.Height - (3972 + (Form1.ScaleHeight - 6620))

End Sub

Private Sub Check1_Click()
If Check1.Value = 1 Then Combo1.Text = Combo1.List(0): Combo1.Enabled = False
If Check1.Value = 0 Then Combo1.Enabled = True


End Sub

Private Sub Combo3_Click()
Select Case Combo3.ListIndex

Case 0
 Combo1.Enabled = True
Case 1
 Combo1.Enabled = False
Case 2
 Combo1.Enabled = True

End Select

End Sub


Private Sub SerchingDir()
Dim TF As Integer
For di = 0 To Dir1.ListCount - 1
If StopSearch = True Then Exit For

TF = InStrRev(Dir1.List(di), "\")
FolderName = Mid(Dir1.List(di), TF + 1, Len(Dir1.List(di)) - (TF - 1))

If InStr(1, FolderName, Text1, vbTextCompare) Then
 FilePath = Dir1.List(di)

TempFolders
End If

DoEvents
Next di

End Sub


Private Sub SearchHiddenFiles()
ScanHiddenFiles
HiddenFiles
'HiddenFolders
End Sub
Private Sub HiddenFolders()
For Each Foldr In Foldr.SubFolders
 If StopSearch = True Then Exit For
  If InStr(1, Foldr.Name, Text1, vbTextCompare) Then
   FolderName = Foldr.Name
   FilePath = Foldr.Path
   TempFolders
  End If
HiddenFolders
DoEvents
Next
End Sub
Private Sub Command1_Click()
Dim t As Integer
Toolbar1.Buttons(8).Enabled = False
Toolbar1.Buttons(7).Enabled = False
Toolbar1.Buttons(6).Enabled = False
List1.Clear
List4.Clear
File1.Refresh
Dir1.Refresh

Picture4.Visible = False
For sr = 0 To List3.ListCount - 1
If Text1 = List3.List(sr) Then t = t + 1
Next sr
If t = 0 Then List3.AddItem Text1
Form1.VScroll1.Max = 0
Types
BIcon = 0
SIcon = 0
FilesNum = 0
ListView1.ListItems.Clear
Set ListView1.Icons = Nothing
Set ListView1.SmallIcons = Nothing

ImageList1.ListImages.Clear
ImageList2.ListImages.Clear
ImageList2.ImageHeight = 16
ImageList2.ImageWidth = 16

ListView1.Visible = False
Picture3.Visible = True
Picture2.Height = (3972 + (Form1.ScaleHeight - 6620))
Frame1.Enabled = False
Command5.Enabled = True
Command1.Enabled = False
SearchingNow
Toolbar1.Enabled = False

If Check1.Value = 1 Then
If Mid(Text1, 1, 2) = "*." Then Text1 = Mid(Text1, 2, Len(Text1))
AllHidden
Exit Sub
End If

Dir1.Path = Text3
StatusBar1.Panels(1).Text = "Searching In: " & Dir1.Path & "  (And Sub Folders...)"
SearchType
SubFolders
End Sub


Private Sub ShowFolders()
Dim fld As Long
Dim Br As BROWSEINFO
Dim Path As String
Dim pos1 As Integer


Br.ulFlags = BIF_RETURNONLYFSDIRS
fld = SHBrowseForFolder(Br)
Path = Space(MAX_PATH)
If SHGetPathFromIDList(ByVal fld, ByVal Path) Then
Text3.Text = Path
End If
End Sub

Private Sub SearchType()
If Combo3.ListIndex = -1 Or Combo3.ListIndex = 0 Then
Searching
ElseIf Combo3.ListIndex = 1 Then
SerchingDir
ElseIf Combo3.ListIndex = 2 Then
SerchingDir
Searching

End If


End Sub

Private Sub HiddenFiles()
For Each Foldr In Foldr.SubFolders
 If StopSearch = True Then Exit For
ScanHiddenFiles
 HiddenFiles
DoEvents
Next
End Sub

Private Sub ScanHiddenFiles()
For Each FI In Foldr.Files
 If StopSearch = True Then Exit For
  If InStr(1, FI.Name, Text1, vbTextCompare) Then
   filename = FI.Name
   FilePath = FI.Path
   Temp
  End If
  
DoEvents
Next
End Sub

Private Sub SearchingNow()
Label11.Caption = "Looking For: " & """" & Text1 & """"
Label12.Caption = "Looking In: (" & Text3 & ") And Sub Folders"
Label13.Caption = "Searching For: " & """" & Combo1.Text & """"
If Check1.Value = 1 Then Label15.Caption = "Looking In Hidden Folders And Files." Else Label15.Caption = "Not Looking In Hidden Folders And Files."
Label16 = "Found (0) Item(s) So Far"
Picture9.Visible = True
Timer1.Enabled = True
Timer1.Interval = 1
If Timer1.Interval < 30 Then Timer1.Interval = Timer1.Interval + 1 Else Timer1.Enabled = False
StopSearch = False
HScroll1.Visible = False
End Sub

Private Sub Command3_Click()
ShowFolders
End Sub


Private Sub RecyceleBin(strFile As String)
Dim SHop As SHFILEOPSTRUCT
Dim ItemD As Integer
strFile = strFile + Chr(0)
On Error GoTo er

With SHop
.wFunc = FO_DELETE
.pFrom = strFile
.fFlags = Flg
End With

SHFileOperation SHop

For i = 0 To List4.ListCount - 1
ItemD = List4.List(i)
 If InStr(ListView1.ListItems.Item(ItemD).SubItems(1), ".") Then
  If FS.FileExists(ListView1.ListItems.Item(ItemD).SubItems(1)) = False Then
   ListView1.ListItems.Item(ItemD).Bold = True
  End If
 Else
  If FS.FolderExists(ListView1.ListItems.Item(ItemD).SubItems(1)) = False Then
   ListView1.ListItems.Item(ItemD).Bold = True
  End If
End If
Next i

re:
For i = 1 To ListView1.ListItems.Count
 If ListView1.ListItems.Item(i).Bold = True Then
  ListView1.ListItems.Remove (i)
  GoTo re
 End If
Next i

List4.Clear
ListView1.Refresh
Dir1.Refresh
File1.Refresh


Exit Sub
er:
Resume Next
End Sub


Private Sub StopSearching()
FinalAdding
StopSearch = True
Frame1.Enabled = True
Command5.Enabled = False
Command1.Enabled = True

Toolbar1.Enabled = True
Picture9.Visible = False
Timer1.Enabled = False

End Sub

Private Sub CopyF()
List1.Clear
For i = 1 To ListView1.ListItems.Count
 If ListView1.ListItems.Item(i).Selected = True Then
  List1.AddItem ListView1.ListItems.Item(i).SubItems(1)
  If CutF = True Then ListView1.ListItems.Item(i).Ghosted = True Else ListView1.ListItems.Item(i).Ghosted = False
 End If
Next i
 CutF = False

If List1.ListCount = 0 Then Toolbar1.Buttons(8).Enabled = False Else Toolbar1.Buttons(8).Enabled = True
End Sub



Private Sub Command5_Click()
StopSearching
End Sub

Private Sub SearchHiddenFolders()
ScanHiddenFolders
Set Foldr = FS.GetFolder(Text3)
   For Each Foldr In Foldr.SubFolders
    If StopSearch = True Then Exit For
    StatusBar1.Panels(1).Text = "Searching In: " & Foldr.Path & "  (And Sub Folders...)"
    HiddenFolders
    DoEvents
 Next
End Sub
Private Sub AllHidden()

If Check1.Value = 1 Then
 If Combo3.ListIndex = -1 Or Combo3.ListIndex = 0 Then
  Set Foldr = FS.GetFolder(Text3)
  ScanHiddenFiles
  Set Foldr = FS.GetFolder(Text3)
   For Each Foldr In Foldr.SubFolders
   If StopSearch = True Then Exit For

    StatusBar1.Panels(1).Text = "Searching In: " & Foldr.Path & "  (And Sub Folders...)"
    SearchHiddenFiles
    DoEvents
  Next
ElseIf Combo3.ListIndex = 1 Then
  Set Foldr = FS.GetFolder(Text3)
    StatusBar1.Panels(1).Text = "Searching In: " & Foldr.Path & "  (And Sub Folders...)"
    SearchHiddenFolders
 End If
End If

If StopSearch = False Then StopSearching
StopSearch = True
End Sub
Private Sub ScanHiddenFolders()
For Each Foldr In Foldr.SubFolders
 If StopSearch = True Then Exit For
  If InStr(1, Foldr.Name, Text1, vbTextCompare) Then
   FolderName = Foldr.Name
   FilePath = Foldr.Path
   TempFolders
  End If
DoEvents
Next

End Sub


Private Sub Cpy_Click()
 CopyF
End Sub

Private Sub CT_Click()
 CutF = True
 CopyF
End Sub

Private Sub Dir1_Change()
DoEvents
File1.Path = Dir1.Path
End Sub

Private Sub Dlt_Click()
Flg = 0
List4.Clear
MoveFile

End Sub

Private Sub Form_Activate()
ImageList5.ListImages.Item(1).Draw Picture14.hdc, 70, 70, 1
Picture14.Picture = Picture14.Image

End Sub

Private Sub Form_Load()
Form1.Height = 7080
Form1.Width = 9705
If Screen.TwipsPerPixelX = 12 Then
BX = 220
B2 = 200
Else
BX = 270
B2 = 210
End If


V = 1
V2 = 1
V3 = 1
Picture14.FillColor = vbRed

Hi = 200
Dim Tx As String
Dim L As Integer

Open App.Path & "\save.dat" For Input As #1
Do Until EOF(1)
Line Input #1, Tx
List3.List(L) = Tx
L = L + 1
Loop
Close #1

List2.BackColor = RGB(255, 249, 236)

With Combo1
.AddItem "All Files"
.AddItem "Image Files"
.AddItem "Text Files"
.AddItem "Video And Audio Files"
.AddItem "Customize..."
End With



With Combo3
.AddItem "Files Only"
.AddItem "Folders Only"
.AddItem "Files And Folders"
End With

With ListView1
.ColumnHeaders.Add , , "File Name"
.ColumnHeaders.Add , , "File Path"
.ColumnHeaders.Add , , "File Size"
End With
Dir1.Path = "c:\"
End Sub

Private Sub Types()
Dim Ty As String
Select Case Combo1.ListIndex
  Case 0
    File1.Pattern = "*.*"
  Case 1
    File1.Pattern = "*.bmp;*.ico;*.gif;*.wmf;*.jpg;*.jpe;*.jpeg;*.jfif;*.tif;*.pcd;*.pcx;*.psd;*.png"
  Case 2
    File1.Pattern = "*.txt;*.doc;*.dat;*.rtf;*.wri"
  Case 3
    File1.Pattern = "*.avi;*.mpeg;*.mpg;*.mpa;*.wav;*.wave;*.mp3;*.mp2;*.asf;*.mpe;*.mov;*.wmv;*.m1v;*.wm;*.wma;*.miv;*.rmi;*.midi;*.mid;*.asx;*.wax;*.m3u;*.cda;*.snd;*.au;*.aif;*.aifc;*.aiff;*.wma"
  Case 4
  Ty = InputBox("Enter The File Type Ex: (*.bmp) For Bitmap Files", "File Type.")
  
  If Ty <> "" Then File1.Pattern = Ty Else File1.Pattern = "*.*"

End Select

End Sub


Private Sub FinalAdding()
Dim S As ListItem
StatusBar1.Panels(1).Text = "(" & FilesNum & ") Object(s) Found  " & "Searching Is Ready..."

If FilesNum = 0 Then: Toolbar1.Enabled = True: Label18.Visible = True: Label17.Visible = False: Exit Sub
Label18.Visible = False: Label17.Visible = True

Picture3.Visible = False
ListView1.Visible = True

ListView1.Icons = ImageList1
ListView1.SmallIcons = ImageList2

For tm = 1 To FilesNum
Set S = ListView1.ListItems.Add(, , Label4(tm), tm, tm)
S.SubItems(1) = Label5(tm)
S.SubItems(2) = Label6(tm)

Unload Label4(tm)
Unload Label5(tm)
Unload Label6(tm)

Next tm
Picture3.Cls

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
For i = 0 To List3.ListCount - 1
If i = List3.ListCount - 1 Then Text5.Text = Text5.Text & List3.List(i) Else Text5.Text = Text5.Text & List3.List(i) & Chr(13) + Chr(10)
Next i

Open App.Path & "\save.dat" For Output As #1
Print #1, Text5
Close #1

If StopSearch = False Then StopSearching
StopSearch = True
End
End Sub

Private Sub Form_Resize()
On Error GoTo ErrHND

If Form1.WindowState <> 1 Then
ListView1.Height = 3990 + (Form1.ScaleHeight - 6620)
Picture2.Height = 3972 + (Form1.ScaleHeight - 6620)
Picture3.Height = 4008 + (Form1.ScaleHeight - 6620)
VScroll1.Height = 3730 + (Form1.ScaleHeight - 6620)

ListView1.Width = 6600 + (Form1.ScaleWidth - 9585)
Picture3.Width = 6600 + (Form1.ScaleWidth - 9585)
Command7(3).Width = Picture3.Width - 5350
Picture2.Width = 6370 + (Form1.ScaleWidth - 9585)
VScroll1.Left = Picture3.Width - VScroll1.Width
Toolbar1.Width = 6600 + (Form1.ScaleWidth - 9585)
Frame3.Width = 6600 + (Form1.ScaleWidth - 9585)
Frame1.Height = 6255 + (Form1.ScaleHeight - 6620)
Command2.Height = 6255 + (Form1.ScaleHeight - 6620)
End If

Exit Sub
ErrHND:
MsgBox (Err.Description)
Resume Next
End Sub


Private Sub Frame4_Click()
List2.Visible = False
End Sub

Private Sub HScroll1_Change()
Picture4.Left = -HScroll1.Value + 50
Picture4.SetFocus
End Sub

Private Sub HScroll1_Scroll()
Picture4.Left = -HScroll1.Value + 50

End Sub


Private Sub Label8_Click()
Dim OpnFolder As String
OpnFolder = Mid(Label8, 19, Len(Label8) - 19)
ShowFileProperties OpnFolder, Me.hwnd, False
End Sub

Private Sub List2_Click()
Text1 = List2.List(List2.ListIndex)
List2.Visible = False

End Sub

Private Sub ListView1_AfterLabelEdit(Cancel As Integer, NewString As String)
If NewString <> "" And NewString <> " " Then
Name (ListView1.ListItems.Item(D).SubItems(1)) As UP(ListView1.ListItems.Item(D).SubItems(1)) & NewString
ListView1.ListItems.Item(D).SubItems(1) = UP(ListView1.ListItems.Item(D).SubItems(1)) & NewString
Else: Cancel = 1
End If
End Sub

Private Sub ListView1_Click()
If ItemSelected = False Then
 Toolbar1.Buttons(6).Enabled = False
 Toolbar1.Buttons(7).Enabled = False
 Toolbar1.Buttons(11).Enabled = False
 Toolbar1.Buttons(12).Enabled = False
 D = 0
 Picture4.Visible = False
 HScroll1.Visible = False
  For i = 1 To ListView1.ListItems.Count
   ListView1.ListItems.Item(i).Selected = False
  Next i
 ItemSelected = False
End If

ItemSelected = False
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
'b = (ColumnHeader.Index)
'ListView1.SortKey = b - 1
'ListView1.Sorted = True
End Sub


Private Sub ListView1_DblClick()
If D <> 0 Then ShowFileProperties ListView1.ListItems.Item(D).SubItems(1), Me.hwnd, False

End Sub

Private Sub ListView1_ItemClick(ByVal Item As MSComctlLib.ListItem)
On Error GoTo ErrHand
D = Item.Index
ItemSelected = True

Toolbar1.Buttons(6).Enabled = True
Toolbar1.Buttons(7).Enabled = True
Toolbar1.Buttons(11).Enabled = True
Toolbar1.Buttons(12).Enabled = True

For i = 1 To ListView1.ListItems.Count
If ListView1.ListItems.Item(i).Selected = True Then it = it + 1
Next i


If it > 1 Then Exit Sub
Picture4.Visible = True
Picture8.Cls
ImageList1.ListImages.Item(D).Draw Picture8.hdc, 0, 0, 1

Label7 = ListView1.ListItems.Item(D).Text
Label8 = "Container Folder (" & UP(ListView1.ListItems.Item(D).SubItems(1)) & ")"

If InStr(Label7, ".") Then
  Set FI = FS.GetFile(ListView1.ListItems.Item(D).SubItems(1))
  Label9 = "Date Modefide: " & FI.DateLastModified
  Label10 = "Size: " & ListView1.ListItems.Item(D).SubItems(2)
 Else
  Set Foldr = FS.GetFolder(ListView1.ListItems.Item(D).SubItems(1))
  Label9 = "Date Modefide: " & Foldr.DateLastModified
  Label10 = "Size: " & Foldr.Size / 1000 & " KB"
End If

r:
If Label8.Width + 216 > 6492 Then Picture4.Width = Label8.Width + 216
HScroll1.Max = Abs(Picture4.Width - 6490)
If HScroll1.Max = 0 Then HScroll1.Visible = False Else HScroll1.Visible = True
If MouseButton = 2 Then Form1.PopupMenu Menu, , , , opn
MouseButton = 1

Exit Sub
ErrHand:

If Err.Number = 76 Then
  Set FI = FS.GetFile(ListView1.ListItems.Item(D).SubItems(1))
  Label9 = "Date Modefide: " & FI.DateLastModified
  Label10 = "Size: " & ListView1.ListItems.Item(D).SubItems(2)
  Resume r:
End If
End Sub


Private Sub ListView1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'For i = 0 To ListView1.ListItems.Count
 'ListView1.SelectedItem.Selected = False
'Next i
MouseButton = Button
ListView1.SetFocus
End Sub






Private Sub MoveFile()
FilesDelet = ""
For i = 1 To ListView1.ListItems.Count
 If ListView1.ListItems.Item(i).Selected = True Then
 List4.AddItem i
   FilesDelet = FilesDelet + ListView1.ListItems.Item(i).SubItems(1) + Chr(0)
 End If
Next i

RecyceleBin (FilesDelet)
End Sub

Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
MouseButton = Button
ListView1.SetFocus
nn = ItemSelected

End Sub


Private Function ShowFileProperties(filename As String, OwnerhWnd As Long, props As Boolean) As Long
    'Call API Function to show properties or open file.
    Dim SEI As SHELLEXECUTEINFO
    With SEI
        .cbSize = Len(SEI)
        .fMask = 64 Or 12 Or 1024
        .hwnd = OwnerhWnd
        If props Then .lpVerb = "properties" Else .lpVerb = "open"
        .lpFile = filename
        .lpParameters = vbNullChar
        .lpDirectory = vbNullChar
        If props Then .nShow = 0 Else .nShow = 1
        .hInstApp = 0
        .lpIDList = 0
    End With
    ShellExecuteEX SEI
    ShowFileProperties = SEI.hInstApp
End Function

Private Sub opn_Click()
ShowFileProperties ListView1.ListItems.Item(D).SubItems(1), Me.hwnd, False
End Sub

Private Sub Picture14_Click()
'Timer1.Enabled = False
End Sub

Private Sub Prop_Click()
If D <> 0 Then ShowFileProperties ListView1.ListItems.Item(D).SubItems(1), Me.hwnd, True
End Sub

Private Sub Rec_Click()
Flg = 64
List4.Clear
MoveFile

End Sub

Private Sub Rename_Click()
ListView1.StartLabelEdit

End Sub
Private Sub Text1_Change()
Dim B As String
List2.Clear

w = Chr(13) + Chr(10)

For i = 0 To List3.ListCount - 1
t = List3.List(i)

If Mid(Text1.Text, 1, 1) = Mid(t, 1, 1) Then

For n = 1 To Len(Text1.Text)
If Mid(Text1.Text, n, 1) = Mid(t, n, 1) Then
B = B & Mid(Text1.Text, n, 1)
End If
Next n
End If

If B = (Text1.Text) Then List2.AddItem t
B = ""
Next i
B = ""

If List2.ListCount <= 0 Or Text1.Text = "" Then
Click = False
List2.Visible = False
Else: List2.Visible = True
End If

If List2.ListCount = 1 And Text1.Text = List2.List(0) Then List2.Visible = False
End Sub

Private Sub Text1_LostFocus()
List2.Visible = False
End Sub


Private Sub RefreshList()
re:
For i = 1 To ListView1.ListItems.Count
If InStr(ListView1.ListItems.Item(i).SubItems(1), ".") Then
  If FS.FileExists(ListView1.ListItems.Item(i).SubItems(1)) = False Then
   ListView1.ListItems.Remove (i)
   GoTo re
  End If
Else
  If FS.FolderExists(ListView1.ListItems.Item(i).SubItems(1)) = False Then
   ListView1.ListItems.Remove (i)
   GoTo re
  End If
End If
Next i


ListView1.Refresh
ListView1.Arrange = lvwAutoTop
Dir1.Refresh
File1.Refresh

End Sub


Private Sub Timer1_Timer()
Picture14.Cls
SX = SX + 0.1
SX2 = SX2 + 0.1
SX3 = SX3 + 0.1

If SX >= 5.2 Then V = -5
If SX >= 6.2 Then V = 1: SX = 0

If SX2 >= 3.2 Then V2 = -5
If SX2 >= 4.2 Then V2 = 1: SX2 = -2

If SX3 >= 1.2 Then V3 = -5
If SX3 >= 2.2 Then V3 = 1: SX3 = -4

'220
'200

'270
'210
Picture14.Circle (BX, BX), B2, RGB(179, 198, 213), SX, SX + V
Picture14.Circle (BX, BX), B2, RGB(179, 198, 213), SX2 + 2, SX2 + 2 + V2
Picture14.Circle (BX, BX), B2, RGB(179, 198, 213), SX3 + 4, SX3 + 4 + V3
If Timer1.Interval < 30 Then Timer1.Interval = Timer1.Interval + 1 Else Picture14.Cls: Timer1.Enabled = False

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Index
Case 1
 ListView1.View = lvwIcon
 
Case 2
 ListView1.View = lvwList
Case 3
 ListView1.View = lvwReport
Case 4
 ListView1.View = lvwSmallIcon
Case 6
 CutF = True
 CopyF
Case 7
 'Copy File Or Folder
 CutF = False
 CopyF
Case 8
 Past
Case 10
RefreshList
Case 11
Flg = 64
List4.Clear
MoveFile

Case 12
Flg = 0
List4.Clear
MoveFile

 End Select
End Sub

Private Sub Past()
On Error GoTo errh:

Dim fld As Long
Dim Br As BROWSEINFO
Dim Path As String
Dim pos1 As Integer
Dim CopyName As String

Br.ulFlags = BIF_RETURNONLYFSDIRS
fld = SHBrowseForFolder(Br)
Path = Space(MAX_PATH)
If SHGetPathFromIDList(ByVal fld, ByVal Path) Then
Text4 = Path
NewFolder = Text4

For i = 0 To List1.ListCount - 1
n = InStrRev(List1.List(i), "\")
CopyName = Mid(List1.List(i), n + 1, (Len(List1.List(i)) - (CInt(n) - 1)))
CopyFromTo List1.List(i), NewFolder & "\" & CopyName
Next i
End If

re:
For i = 1 To ListView1.ListItems.Count
 If ListView1.ListItems.Item(i).Ghosted = True Then
  ListView1.ListItems.Remove (i)
  GoTo re
 End If
Next i

Exit Sub
errh:
If Err.Number = 58 Then MsgBox "The File " & "(" & CopyName & ")" & " Is Already Exist!", vbExclamation
Resume Next
End Sub

Public Sub CopyFromTo(Source As String, Destination As String)
    Dim lRet As Long
    Dim SHFileOp As SHFILEOPSTRUCT

    On Error GoTo Copy_Err

'ensure source and destination strings have valid terminators
    While Not Right(Source, 2) <> vbNullChar & vbNullChar
        Source = Source & vbNullChar
    Wend

    While Not Right(Destination, 2) <> vbNullChar & vbNullChar
        Destination = Destination & vbNullChar
    Wend

    With SHFileOp
        .wFunc = FO_COPY
        .pFrom = Source
        .pTo = Destination
        .fFlags = Options
    End With

    lRet = SHFileOperation(SHFileOp)


Copy_Exit:
    Exit Sub

Copy_Err:
    MsgBox Error, vbCritical
    Resume Copy_Exit

End Sub

Private Sub VScroll1_Change()
Picture2.Top = -VScroll1.Value

For re = 1 To FilesNum
ImageList2.ListImages.Item(re).Draw Picture2.hdc, 20, CLng(re) * Hi, 1
Next re
'Picture2.SetFocus

End Sub


Private Sub VScroll1_Scroll()
Picture2.Top = -VScroll1.Value
Picture2.SetFocus

For re = 1 To FilesNum
ImageList2.ListImages.Item(re).Draw Picture2.hdc, 20, CLng(re) * Hi, 1
Next re

End Sub

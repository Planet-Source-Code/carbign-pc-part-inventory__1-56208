VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{82351433-9094-11D1-A24B-00A0C932C7DF}#1.5#0"; "AniGIF.ocx"
Begin VB.Form frmMain 
   BackColor       =   &H8000000D&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "PC-Express"
   ClientHeight    =   8610
   ClientLeft      =   3330
   ClientTop       =   2265
   ClientWidth     =   10350
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000D&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8610
   ScaleWidth      =   10350
   ShowInTaskbar   =   0   'False
   Begin Project1.lvButtons_H cmdSales 
      Height          =   375
      Left            =   8400
      TabIndex        =   54
      Top             =   1560
      Width           =   1935
      _extentx        =   3413
      _extenty        =   661
      caption         =   "SALES REPORT"
      capalign        =   2
      backstyle       =   2
      gradient        =   1
      cgradient       =   255
      cfore           =   8388608
      mode            =   0
      value           =   0   'False
      cfhover         =   8388608
      cback           =   -2147483633
      cbhover         =   255
   End
   Begin MSDBGrid.DBGrid DBGrid3 
      Bindings        =   "frmMain.frx":0000
      Height          =   975
      Left            =   120
      OleObjectBlob   =   "frmMain.frx":0014
      TabIndex        =   51
      Top             =   7320
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.Data Data3 
      Caption         =   "Data3"
      Connect         =   "Access"
      DatabaseName    =   "C:\Program Files\Microsoft Visual Studio\VB98\PCX\pcx.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   8760
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Sales"
      Top             =   9360
      Width           =   1140
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   3600
      Top             =   6360
   End
   Begin VB.TextBox Text15 
      Appearance      =   0  'Flat
      DataField       =   "ProName"
      DataSource      =   "Data1"
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   0
      TabIndex        =   48
      Top             =   2400
      Visible         =   0   'False
      Width           =   4095
   End
   Begin VB.TextBox Text14 
      DataField       =   "Quantity"
      DataSource      =   "Data2"
      Height          =   285
      Left            =   480
      TabIndex        =   47
      Text            =   "Text14"
      Top             =   9360
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.TextBox Text13 
      DataField       =   "Item"
      DataSource      =   "Data2"
      Height          =   285
      Left            =   -120
      TabIndex        =   46
      Text            =   "Text13"
      Top             =   9360
      Visible         =   0   'False
      Width           =   615
   End
   Begin Project1.lvButtons_H cmdRemove 
      Height          =   375
      Left            =   8640
      TabIndex        =   38
      Top             =   2400
      Visible         =   0   'False
      Width           =   1695
      _extentx        =   2990
      _extenty        =   661
      caption         =   "REMOVE ITEM"
      capalign        =   2
      backstyle       =   2
      cfore           =   -2147483635
      mode            =   0
      value           =   0   'False
      cfhover         =   -2147483635
      cback           =   16777215
      lockhover       =   1
   End
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   2640
      TabIndex        =   37
      Top             =   9360
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.TextBox Text8 
      DataField       =   "ProName"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   1080
      TabIndex        =   36
      Text            =   "Text8"
      Top             =   9360
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H8000000D&
      Caption         =   "JOB ORDER"
      ForeColor       =   &H00FFFFFF&
      Height          =   5775
      Left            =   4200
      TabIndex        =   26
      Top             =   2760
      Visible         =   0   'False
      Width           =   6135
      Begin VB.Frame Frame4 
         BackColor       =   &H8000000D&
         ForeColor       =   &H00FFFFFF&
         Height          =   975
         Left            =   120
         TabIndex        =   39
         Top             =   1080
         Width           =   5775
         Begin VB.TextBox Text12 
            Appearance      =   0  'Flat
            ForeColor       =   &H8000000D&
            Height          =   285
            Left            =   1200
            TabIndex        =   45
            Top             =   600
            Width           =   1455
         End
         Begin VB.TextBox Text11 
            Appearance      =   0  'Flat
            ForeColor       =   &H8000000D&
            Height          =   285
            Left            =   3000
            TabIndex        =   43
            Top             =   240
            Width           =   1455
         End
         Begin VB.TextBox Text10 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            ForeColor       =   &H8000000D&
            Height          =   285
            Left            =   1200
            TabIndex        =   41
            Top             =   240
            Width           =   495
         End
         Begin VB.Label Label17 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "AMOUNT:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   44
            Top             =   600
            Width           =   975
         End
         Begin VB.Label Label16 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "UNIT PRICE:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   1800
            TabIndex        =   42
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label15 
            BackColor       =   &H8000000D&
            BackStyle       =   0  'Transparent
            Caption         =   "QUANTITY:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   3720
         TabIndex        =   35
         Top             =   5280
         Width           =   2295
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   3720
         TabIndex        =   33
         Top             =   4920
         Width           =   2295
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   2160
         TabIndex        =   29
         Top             =   360
         Width           =   3735
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   2160
         TabIndex        =   28
         Top             =   720
         Width           =   3735
      End
      Begin MSDBGrid.DBGrid DBGrid2 
         Bindings        =   "frmMain.frx":0D3F
         Height          =   2535
         Left            =   120
         OleObjectBlob   =   "frmMain.frx":0D53
         TabIndex        =   27
         Top             =   2280
         Width           =   5895
      End
      Begin VB.Label Label14 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "AMOUNT DUE:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   5280
         Width           =   1815
      End
      Begin VB.Label Label13 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "DISCOUNT:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   32
         Top             =   5040
         Width           =   1095
      End
      Begin VB.Label Label12 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "CUSTOMER NAME:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label11 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "CUSTOMER PHONE:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   720
         Width           =   1815
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000D&
      Caption         =   "SEARCH RESULTS"
      ForeColor       =   &H8000000E&
      Height          =   3615
      Left            =   0
      TabIndex        =   20
      Top             =   2760
      Width           =   4095
      Begin VB.TextBox Text2 
         DataField       =   "ProPrice"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   285
         Left            =   2160
         TabIndex        =   23
         Top             =   2400
         Width           =   1815
      End
      Begin VB.TextBox Text3 
         DataField       =   "PriUnit"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000001&
         Height          =   285
         Left            =   2160
         TabIndex        =   22
         Top             =   2760
         Width           =   1815
      End
      Begin MSDBGrid.DBGrid DBGrid1 
         Bindings        =   "frmMain.frx":1A92
         Height          =   2055
         Left            =   120
         OleObjectBlob   =   "frmMain.frx":1AA6
         TabIndex        =   21
         Top             =   240
         Width           =   3855
      End
      Begin VB.Label Label3 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "QUICK |"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   50
         Top             =   3240
         Width           =   735
      End
      Begin VB.Label Label4 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "BROWSE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0C0C0&
         Height          =   255
         Left            =   960
         TabIndex        =   49
         Top             =   3240
         Width           =   855
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         X1              =   240
         X2              =   1560
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Label Label9 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "ITEM PRICE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Label Label10 
         BackColor       =   &H8000000D&
         BackStyle       =   0  'Transparent
         Caption         =   "ITEM AVAILBLE"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   2760
         Width           =   1935
      End
   End
   Begin VB.Data Data2 
      Caption         =   "Data2"
      Connect         =   "Access"
      DatabaseName    =   "C:\Program Files\Microsoft Visual Studio\VB98\PCX\pcx.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   7560
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Customer"
      Top             =   9360
      Width           =   1140
   End
   Begin VB.TextBox Text1 
      DataField       =   "ProName"
      DataSource      =   "Data1"
      Height          =   285
      Left            =   5400
      TabIndex        =   19
      Text            =   "Text1"
      Top             =   9360
      Visible         =   0   'False
      Width           =   855
   End
   Begin Project1.lvButtons_H cmdList 
      Height          =   375
      Left            =   6120
      TabIndex        =   18
      Top             =   1560
      Width           =   2295
      _extentx        =   4471
      _extenty        =   661
      caption         =   "VIEW PRICELIST"
      capalign        =   2
      backstyle       =   2
      shape           =   1
      gradient        =   1
      cgradient       =   255
      cfore           =   8388608
      mode            =   0
      value           =   0   'False
      cfhover         =   8388608
      cback           =   16777215
      cbhover         =   255
   End
   Begin Project1.lvButtons_H cmdLook 
      Height          =   375
      Index           =   1
      Left            =   1680
      TabIndex        =   17
      Top             =   1920
      Width           =   885
      _extentx        =   1561
      _extenty        =   661
      caption         =   "SEARCH"
      capalign        =   2
      backstyle       =   2
      gradient        =   3
      cgradient       =   12632256
      cfore           =   -2147483647
      mode            =   0
      value           =   0   'False
      cfhover         =   -2147483647
      cback           =   -2147483633
      cbhover         =   12632256
   End
   Begin VB.Timer Timer2 
      Left            =   3120
      Top             =   6360
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2640
      Top             =   6360
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000D&
      Caption         =   "SEARCH RESULTS"
      ForeColor       =   &H8000000E&
      Height          =   1935
      Left            =   0
      TabIndex        =   13
      Top             =   6480
      Visible         =   0   'False
      Width           =   4095
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   1590
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   3855
      End
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   315
      Left            =   0
      TabIndex        =   9
      Text            =   "BROWSE PRODUCTS HERE"
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Program Files\Microsoft Visual Studio\VB98\PCX\pcx.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6360
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Products"
      Top             =   9360
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.TextBox txtSearch 
      Appearance      =   0  'Flat
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   6120
      TabIndex        =   7
      Top             =   600
      Width           =   3495
   End
   Begin Project1.lvButtons_H cmdCart 
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   1920
      Width           =   1455
      _extentx        =   2778
      _extenty        =   661
      caption         =   "ADD TO CART"
      capalign        =   2
      backstyle       =   2
      gradient        =   4
      cgradient       =   -2147483639
      cfore           =   -2147483635
      mode            =   0
      value           =   0   'False
      image           =   "frmMain.frx":2485
      cfhover         =   -2147483635
      enabled         =   0   'False
      cback           =   -2147483633
      cbhover         =   -2147483639
   End
   Begin Project1.lvButtons_H cmdOrder 
      Height          =   375
      Left            =   5640
      TabIndex        =   0
      Top             =   1080
      Width           =   1575
      _extentx        =   2778
      _extenty        =   661
      caption         =   "JOB ORDER"
      capalign        =   2
      backstyle       =   2
      gradient        =   4
      cgradient       =   14737632
      cfore           =   -2147483635
      mode            =   0
      value           =   0   'False
      cfhover         =   -2147483635
      cback           =   16777215
      cbhover         =   128
      lockhover       =   1
   End
   Begin Project1.lvButtons_H cmdUser 
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   1080
      Width           =   1815
      _extentx        =   3201
      _extenty        =   661
      caption         =   "ADD USER"
      capalign        =   2
      backstyle       =   2
      shape           =   1
      gradient        =   4
      cgradient       =   14737632
      cfore           =   -2147483635
      mode            =   0
      value           =   0   'False
      imgalign        =   1
      cfhover         =   -2147483635
      cback           =   16777215
      cbhover         =   14737632
   End
   Begin Project1.lvButtons_H cmdInventory 
      Height          =   375
      Left            =   4080
      TabIndex        =   2
      Top             =   1080
      Width           =   1575
      _extentx        =   2778
      _extenty        =   661
      caption         =   "NEW PRODUCTS"
      capalign        =   2
      backstyle       =   2
      gradient        =   4
      cgradient       =   14737632
      cfore           =   -2147483635
      mode            =   0
      value           =   0   'False
      imgalign        =   1
      cfhover         =   -2147483635
      cback           =   16777215
      cbhover         =   128
      lockhover       =   1
   End
   Begin Project1.lvButtons_H cmdLog 
      Height          =   375
      Left            =   8760
      TabIndex        =   4
      Top             =   1080
      Width           =   1575
      _extentx        =   2778
      _extenty        =   661
      caption         =   "LOG OFF"
      capalign        =   2
      backstyle       =   2
      gradient        =   4
      cgradient       =   -2147483639
      cfore           =   -2147483635
      mode            =   0
      value           =   0   'False
      image           =   "frmMain.frx":25DF
      cfhover         =   -2147483635
      cback           =   -2147483633
      cbhover         =   -2147483639
   End
   Begin Project1.lvButtons_H cmdPrint 
      Height          =   375
      Left            =   7200
      TabIndex        =   5
      Top             =   1080
      Width           =   1575
      _extentx        =   2778
      _extenty        =   661
      caption         =   "PRINT "
      capalign        =   2
      backstyle       =   2
      gradient        =   4
      cgradient       =   14737632
      cfore           =   -2147483635
      mode            =   0
      value           =   0   'False
      cfhover         =   -2147483635
      cback           =   16777215
      cbhover         =   128
      lockhover       =   1
   End
   Begin Project1.lvButtons_H cmdSearch 
      Height          =   375
      Index           =   0
      Left            =   9720
      TabIndex        =   8
      Top             =   480
      Width           =   525
      _extentx        =   926
      _extenty        =   661
      caption         =   "GO"
      capalign        =   2
      backstyle       =   2
      gradient        =   4
      cgradient       =   -2147483635
      cfore           =   4210752
      mode            =   0
      value           =   0   'False
      cfhover         =   4210752
      cback           =   16777215
      cbhover         =   128
      lockhover       =   1
   End
   Begin AniGIFCtrl.AniGIF AniGIF2 
      Height          =   1575
      Left            =   4200
      TabIndex        =   52
      Top             =   6600
      Width           =   1815
      BackColor       =   12632256
      PLaying         =   -1  'True
      Transparent     =   -1  'True
      Speed           =   1
      Stretch         =   1
      AutoSize        =   0   'False
      SequenceString  =   ""
      Sequence        =   0
      HTTPProxy       =   ""
      HTTPUserName    =   ""
      HTTPPassword    =   ""
      MousePointer    =   0
      GIF             =   "frmMain.frx":2761
      ExtendWidth     =   3201
      ExtendHeight    =   2778
      Loop            =   0
      AutoRewind      =   0   'False
      Synchronized    =   -1  'True
   End
   Begin AniGIFCtrl.AniGIF AniGIF3 
      Height          =   3615
      Left            =   4200
      TabIndex        =   53
      Top             =   2880
      Width           =   1815
      BackColor       =   16777215
      PLaying         =   -1  'True
      Transparent     =   -1  'True
      Speed           =   1
      Stretch         =   0
      AutoSize        =   0   'False
      SequenceString  =   ""
      Sequence        =   0
      HTTPProxy       =   ""
      HTTPUserName    =   ""
      HTTPPassword    =   ""
      MousePointer    =   0
      BackgroundPicture=   "frmMain.frx":4430
      GIF             =   "frmMain.frx":FBB8
      ExtendWidth     =   3201
      ExtendHeight    =   6376
      Loop            =   0
      AutoRewind      =   0   'False
      Synchronized    =   -1  'True
   End
   Begin VB.Image Image1 
      Height          =   1545
      Left            =   0
      Picture         =   "frmMain.frx":1B33C
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2235
   End
   Begin VB.Label Label7 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "DATE TODAY:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   16
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1680
      TabIndex        =   15
      Top             =   1560
      Width           =   1815
   End
   Begin VB.Label Label8 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6960
      TabIndex        =   12
      Top             =   5040
      Width           =   2415
   End
   Begin VB.Label Label6 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   6960
      TabIndex        =   11
      Top             =   4800
      Width           =   2415
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000D&
      Height          =   255
      Left            =   1560
      TabIndex        =   10
      Top             =   2520
      Width           =   2175
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000D&
      BackStyle       =   0  'Transparent
      Caption         =   "PRODUCT SEARCH:"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3960
      TabIndex        =   6
      Top             =   600
      Width           =   2055
   End
   Begin VB.Image Image2 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   615
      Left            =   2160
      Picture         =   "frmMain.frx":1C47A
      Stretch         =   -1  'True
      Top             =   960
      Width           =   8820
   End
   Begin VB.Image Image3 
      Height          =   3195
      Left            =   2160
      Picture         =   "frmMain.frx":1EE01
      Stretch         =   -1  'True
      Top             =   -2040
      Width           =   10005
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAgain_Click()
Unload Me: frmMain.Show
End Sub

Private Sub cmdCart_Click()
Dim X, Y, a, b
If Text10.Text = "" Then
    MsgBox "enter quantity", vbInformation
Else
X = Val(Text12.Text)
Y = Val(Text7.Text)
a = Val(Text3.Text)
b = Val(Text10.Text)
Text7.Text = X + Y
Data2.Recordset.AddNew: Data3.Recordset.AddNew
Text9.Text = Text8.Text
Text11.Text = Text2.Text
Data2.Recordset.Fields("Item") = Text9.Text
Data2.Recordset.Fields("Price") = Text2.Text
Data2.Recordset.Fields("Quantity") = Text10.Text
Data2.Recordset.Fields("TAmount") = Text12.Text
Data3.Recordset.Fields("Item") = Text9.Text
'Data3.Recordset.Fields("Price") = Text2.Text
Data3.Recordset.Fields("Quantity") = Text10.Text
Data3.Recordset.Fields("TAmount") = Text12.Text
Data2.Recordset.Update: Data3.Recordset.Update: Data2.Refresh: Data3.Refresh
Data1.Recordset.Edit
Text3.Text = a - b
Data1.Recordset.Update:
End If
End Sub


Private Sub cmdInventory_Click()
frmStocks.Show: frmMain.Enabled = False
End Sub

Private Sub cmdLog_Click()
Unload Me: frmSecurity.Show
End Sub

Private Sub cmdLook_Click(Index As Integer)
Dim SearchD As String
SearchD = Combo1.Text
Data1.Recordset.FindFirst "ProType='" & Combo1.Text & "'"
If Trim(SearchD) <> "" Then
    If Data1.Recordset.NoMatch Then
       MsgBox "load error"
    Else
        Data1.RecordSource = "SELECT * FROM Products WHERE ProType = '" & SearchD & "'"
        Frame2.Visible = True: Frame1.Visible = False: DBGrid1.Visible = True
        Text15.Visible = True
    End If
End If

End Sub

Private Sub cmdOrder_Click()
Frame3.Visible = True: cmdRemove.Visible = True: cmdOrder.Enabled = False: cmdCart.Enabled = True
Timer3.Enabled = True: AniGIF2.Visible = False: AniGIF3.Visible = False
End Sub
Private Sub cmdPrint_Click()
DataReport1.Sections("Section2").Controls.Item(1).Caption = Text5.Text
DataReport1.Sections("Section2").Controls.Item(2).Caption = Text4.Text
DataReport1.Sections("Section3").Controls.Item(2).Caption = Text7.Text
DataReport1.Show
End Sub
Private Sub cmdRemove_Click()
On Error Resume Next
Dim SearchD As String
Dim X, Y, z, w
X = Val(Text14.Text)
Y = Val(Text3.Text)
z = Val(Text7.Text)
w = Val(Text12.Text)
SearchD = Text13.Text
Data1.Recordset.FindFirst "ProName='" & Text13.Text & "'"
If Trim(SearchD) <> "" Then
    If Data1.Recordset.NoMatch Then
       MsgBox "Load Error"
    Else
        Data1.RecordSource = "SELECT * FROM Products WHERE ProName = '" & SearchD & "'"
        Data3.RecordSource = "SELECT * From Sales Where Item='" & SearchD & "'"
        'DBGrid1.Visible = True: Frame1.Visible = False: Frame2.Visible = True:
        Data1.Recordset.Edit
        Text3.Text = X + Y
        Data2.Recordset.Delete: Data3.Recordset.Delete
        Data1.Recordset.Update:
        Text7.Text = z - w
    End If
End If
End Sub
Private Sub cmdRItem_Click()
End Sub
Private Sub cmdSearch_Click(Index As Integer)
Const AposAst As String = "'*", AstApos As String = "*'"
Dim Target, msg As String
List1.Clear

    Target = "ProName like" & AposAst & txtSearch & _
    AstApos
    Data1.Recordset.FindFirst Target
        If txtSearch.Text = "" Then
            MsgBox "enter item(s) to be search?"
        Else
        If Data1.Recordset.NoMatch Then
            MsgBox "load error"
            txtSearch.Text = ""
            txtSearch.SetFocus
        Else
            Do Until Data1.Recordset.NoMatch
                List1.AddItem _
                Data1.Recordset("ProName")
                Data1.Recordset.FindNext Target
                'DBGrid1.Visible = True: List1.Visible = True
                'Frame1.Height = 5415
               
                Frame1.Visible = True
            Loop
        End If
        End If
End Sub

Private Sub cmdUser_Click()
frmAdmin.Show: frmMain.Enabled = False

End Sub






Private Sub Command1_Click()

End Sub

Private Sub Form_Load()

With Combo1
    .AddItem "BROWSE PRODUCTS HERE"
    .AddItem "PROCESSORS"
    .AddItem "MOTHERBOARDS"
    .AddItem "HARD DRIVES"
    .AddItem "GRAPHIC CARDS"
    .AddItem "MONITORS"
    .AddItem "OPTICAL STORAGE"
    .AddItem "MEMORY"
    .AddItem "PRINTERS"
    .AddItem "SCANNERS"
    .AddItem "PERIPHERALS"
    .AddItem "SOUND SOLUTIONS"
    .AddItem "POWER SOLUTIONS"
    .AddItem "MOBILE COMPUTER"
    .AddItem "NETWORK & COMM."
    .AddItem "CASING"
    .AddItem "ACCESSORIES"
    .AddItem "SOFTWARES"
    .AddItem "PRE-PAID"
    .AddItem "GADGETS"
    .AddItem "DESKTOPS"
    .AddItem "DIGITAL CAMERA"
End With

End Sub



Private Sub List1_Click()
Text1.Text = List1.Text:

End Sub

Private Sub List1_DblClick()
DBGrid1.Visible = True
End Sub

Private Sub Text10_Change()
Dim X, Y, z
If Val(Text10.Text) > Val(Text3.Text) Then
    MsgBox "error"
    Text10.Text = ""
Else
X = Val(Text10.Text)
Y = Val(Text11.Text)
'z = Val(Label8.Caption)
Text12.Text = X * Y
End If
End Sub

Private Sub Text11_Change()
Dim X, Y, z
X = Val(Text10.Text)
Y = Val(Text11.Text)
'z = Val(Label8.Caption)
Text12.Text = X * Y
End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text15_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub



Private Sub Text2_Change()
Text11.Text = Text2.Text
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
KeyAscii = 0

End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Timer1_Timer()
Label5.Caption = Format(Now, "mm-dd-yyyy")
End Sub

Private Sub Timer3_Timer()
On Error Resume Next
Data2.Recordset.Delete
If Data2.Recordset.RecordCount = 0 Then
    Data2.Recordset.MoveNext:
    Do Until Data2.Recordset.RecordCount = 0
Loop
If Data2.Recordset.RecordCount = 0 Then
    Timer3.Enabled = False

End If
End If
End Sub

Private Sub txtSearch_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

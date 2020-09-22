VERSION 5.00
Begin VB.Form frmAdmin 
   BackColor       =   &H8000000D&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Create Employee account"
   ClientHeight    =   5880
   ClientLeft      =   5385
   ClientTop       =   3720
   ClientWidth     =   5475
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   5475
   ShowInTaskbar   =   0   'False
   Begin Project1.lvButtons_H cmdSave 
      Height          =   495
      Left            =   120
      TabIndex        =   11
      Top             =   5280
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "&SAVE"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   -2147483635
      cFHover         =   -2147483635
      cBhover         =   -2147483643
      cGradient       =   -2147483643
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      cBack           =   16761024
   End
   Begin Project1.lvButtons_H cmdAdd 
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   5280
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "&ADD"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   -2147483635
      cFHover         =   -2147483635
      cBhover         =   16777215
      cGradient       =   16777215
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      cBack           =   16761024
   End
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Program Files\Microsoft Visual Studio\VB98\PCX\pcx.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   2400
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Security"
      Top             =   6120
      Width           =   1140
   End
   Begin Project1.lvButtons_H cmdExit 
      Height          =   495
      Left            =   4080
      TabIndex        =   7
      Top             =   5280
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "EXIT"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   -2147483635
      cFHover         =   16761024
      cBhover         =   -2147483643
      LockHover       =   2
      cGradient       =   -2147483643
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      cBack           =   16761024
   End
   Begin VB.Frame fmEmp 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      Caption         =   "EMPLOYEE INFO"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1980
      Left            =   0
      TabIndex        =   1
      Top             =   1320
      Width           =   5415
      Begin VB.TextBox txtFname 
         Appearance      =   0  'Flat
         DataField       =   "FName"
         DataSource      =   "Data1"
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
         Height          =   330
         Left            =   1920
         TabIndex        =   15
         Top             =   240
         Width           =   3285
      End
      Begin VB.TextBox txtLname 
         Appearance      =   0  'Flat
         DataField       =   "LName"
         DataSource      =   "Data1"
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
         Height          =   330
         Left            =   1920
         TabIndex        =   14
         Top             =   1095
         Width           =   3285
      End
      Begin VB.TextBox txtContact 
         Appearance      =   0  'Flat
         DataField       =   "ContactNum"
         DataSource      =   "Data1"
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
         Height          =   330
         Left            =   1920
         TabIndex        =   13
         Top             =   1515
         Width           =   3285
      End
      Begin VB.TextBox txtMname 
         Appearance      =   0  'Flat
         DataField       =   "MName"
         DataSource      =   "Data1"
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
         Height          =   330
         Left            =   1920
         TabIndex        =   12
         Top             =   660
         Width           =   3285
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         Caption         =   "&FIRST NAME:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   330
         Width           =   1050
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         Caption         =   "&LAST NAME:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         Caption         =   "&CONTACT #:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   1590
         Width           =   990
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         Caption         =   "&MIDDLE NAME:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   1200
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000D&
      Caption         =   "SECURITY SETTINGS"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1665
      Left            =   0
      TabIndex        =   0
      Top             =   3360
      Width           =   5415
      Begin VB.TextBox txtPass 
         Appearance      =   0  'Flat
         DataField       =   "UserPass"
         DataSource      =   "Data1"
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
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   1920
         PasswordChar    =   "*"
         TabIndex        =   18
         Top             =   720
         Width           =   3285
      End
      Begin VB.TextBox txtConPass 
         Appearance      =   0  'Flat
         DataSource      =   "Data1"
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
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   1920
         PasswordChar    =   "*"
         TabIndex        =   17
         Top             =   1200
         Width           =   3285
      End
      Begin VB.TextBox txtUser 
         Appearance      =   0  'Flat
         DataField       =   "UserName"
         DataSource      =   "Data1"
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
         Height          =   330
         IMEMode         =   3  'DISABLE
         Left            =   1920
         TabIndex        =   16
         Top             =   240
         Width           =   3285
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         Caption         =   "&PASSWORD:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   840
         Width           =   1005
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         Caption         =   "C&ONFIRM PASSWORD"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   20
         Top             =   1320
         Width           =   1770
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H8000000D&
         Caption         =   "&USERNAME:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   945
      End
   End
   Begin Project1.lvButtons_H cmdEdit 
      Height          =   495
      Left            =   1440
      TabIndex        =   8
      Top             =   5280
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "&EDIT"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   -2147483635
      cFHover         =   16761024
      cBhover         =   -2147483643
      LockHover       =   2
      cGradient       =   -2147483643
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      cBack           =   16761024
   End
   Begin Project1.lvButtons_H cmdDelete 
      Height          =   495
      Left            =   2760
      TabIndex        =   9
      Top             =   5280
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   873
      Caption         =   "&DELETE"
      CapAlign        =   2
      BackStyle       =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      cFore           =   -2147483635
      cFHover         =   16761024
      cBhover         =   -2147483643
      LockHover       =   2
      cGradient       =   -2147483643
      Gradient        =   3
      Mode            =   0
      Value           =   0   'False
      cBack           =   16761024
   End
   Begin VB.Image Image4 
      Height          =   915
      Left            =   -720
      Picture         =   "frmAdmin.frx":0000
      Top             =   0
      Width           =   8850
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   675
      Left            =   -240
      Picture         =   "frmAdmin.frx":151B
      Stretch         =   -1  'True
      Top             =   5160
      Width           =   7140
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   855
      Left            =   0
      Picture         =   "frmAdmin.frx":3EA2
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5565
   End
   Begin VB.Label lblInstruct 
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAdmin.frx":6F42
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   570
      Left            =   750
      TabIndex        =   6
      Top             =   135
      Width           =   4710
   End
End
Attribute VB_Name = "frmAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAdd_Click()
Data1.Recordset.AddNew: cmdAdd.Visible = False: cmdSave.Visible = True
End Sub

Private Sub cmdExit_Click()
Unload Me: frmMain.Enabled = True
End Sub


Private Sub cmdExit_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Combo1_Change()

End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub cmdSave_Click()
 Data1.Recordset.Update: cmdSave.Visible = False: cmdAdd.Visible = True
End Sub

Private Sub Form_Load()
Data1.DatabaseName = App.Path + "\pcx.mdb"
Data1.RecordSource = "Security"

End Sub

Private Sub txtFname_Change()
txtUser.Text = txtFname.Text
End Sub

Private Sub txtFname_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtLname_Change()
txtUser.Text = txtLname.Text
End Sub

Private Sub txtLname_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtMname_Change()
txtUser.Text = txtMname.Text
End Sub

Private Sub txtMname_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub txtUser_Change()
txtUser.Text = txtFname.Text + Chr(32) + txtMname.Text + Chr(32) + txtLname.Text
End Sub

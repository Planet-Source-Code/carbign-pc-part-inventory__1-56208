VERSION 5.00
Begin VB.Form frmSecurity 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Welcome to PC-Express"
   ClientHeight    =   4410
   ClientLeft      =   6600
   ClientTop       =   4395
   ClientWidth     =   5010
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   5010
   ShowInTaskbar   =   0   'False
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Program Files\Microsoft Visual Studio\VB98\PCX\pcx.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   375
      Left            =   1680
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Security"
      Top             =   4920
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H8000000D&
      Caption         =   "LOCAL USER"
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
      Height          =   1575
      Left            =   120
      TabIndex        =   3
      Top             =   240
      Width           =   4815
      Begin VB.TextBox txtPassword 
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
         ForeColor       =   &H80000001&
         Height          =   315
         IMEMode         =   3  'DISABLE
         Left            =   1800
         PasswordChar    =   "*"
         TabIndex        =   5
         Top             =   720
         Width           =   2895
      End
      Begin VB.ComboBox cboEmp 
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
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   360
         Width           =   2895
      End
      Begin VB.Label Label2 
         BackColor       =   &H8000000D&
         Caption         =   "PASSWORD"
         BeginProperty Font 
            Name            =   "Arial"
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
         TabIndex        =   7
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackColor       =   &H8000000D&
         Caption         =   "EMPLOYEE NAME"
         BeginProperty Font 
            Name            =   "Arial"
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
         TabIndex        =   6
         Top             =   360
         Width           =   1575
      End
   End
   Begin Project1.lvButtons_H cmdOk 
      Height          =   495
      Left            =   2400
      TabIndex        =   0
      Top             =   3840
      Width           =   1215
      _extentx        =   2143
      _extenty        =   873
      caption         =   "&OK"
      capalign        =   2
      backstyle       =   2
      gradient        =   3
      cgradient       =   16777215
      cfore           =   11891757
      mode            =   0
      value           =   0   'False
      image           =   "frmSecurity.frx":0000
      imgalign        =   1
      cfhover         =   11891757
      cback           =   16761024
      cbhover         =   16777215
   End
   Begin Project1.lvButtons_H cmdExit 
      Height          =   495
      Left            =   3720
      TabIndex        =   1
      Top             =   3840
      Width           =   1215
      _extentx        =   2143
      _extenty        =   873
      caption         =   "EXIT"
      capalign        =   2
      backstyle       =   2
      gradient        =   3
      cgradient       =   16777215
      cfore           =   11891757
      mode            =   0
      value           =   0   'False
      image           =   "frmSecurity.frx":0452
      imgalign        =   1
      cfhover         =   11891757
      cback           =   16761024
      cbhover         =   16777215
   End
   Begin VB.Label Label3 
      BackColor       =   &H8000000D&
      Caption         =   "| LOG AS ADMINISTRATOR |"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2280
      TabIndex        =   2
      Top             =   1920
      Width           =   2655
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H8000000D&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H80000005&
      Height          =   3615
      Left            =   0
      Top             =   0
      Width           =   5415
   End
   Begin VB.Image Image1 
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      Height          =   675
      Left            =   -1680
      Picture         =   "frmSecurity.frx":08A4
      Stretch         =   -1  'True
      Top             =   3720
      Width           =   7140
   End
End
Attribute VB_Name = "frmSecurity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdOk_Click()
Static Tries As Integer
Tries = Tries + 1
If cboEmp = "" Then
    response = MsgBox("Choose your User ID!", vbInformation, "Select User Id")
End If
If cboEmp.Text = cboEmp.Text Then
    Data1.Recordset.FindFirst ("UserName='" + cboEmp.Text + "'")
    If Not Data1.Recordset.NoMatch Then
    If txtPassword.Text = Data1.Recordset.Fields("UserPass") Then
    frmMain.Show: Me.Hide: frmMain.cmdUser.Enabled = False: frmMain.cmdInventory.Enabled = False
    Else
    response = MsgBox("Enter correct password.", vbExclamation, "Incorrect Password"): txtPassword.Text = ""
    txtPassword.SetFocus
    If Tries >= NUM_TRIES Then
    response = MsgBox("Please contact your system vendor.", vbCritical, "Too many attempts")
    End
    End If
    End If
    End If

  
End If
End Sub

Private Sub Form_Load()
Set db = DBEngine.Workspaces(0).OpenDatabase(App.Path & "\pcx.mdb")
Combo
End Sub
Public Sub Combo()
Set rs = db.OpenRecordset("Security")
While Not rs.EOF
cboEmp.AddItem (rs!UserName)
rs.MoveNext
Wend
If rs.RecordCount = 0 Then
    MsgBox "No current record in the database."
Else
   
End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.FontBold = False
End Sub

Private Sub Label3_Click()
frmMain.cmdUser.Left = 6960: frmMain.cmdInventory.Left = 8760: frmMain.Show
frmMain.cmdOrder.Visible = False: frmMain.cmdPrint.Visible = False: frmMain.cmdLog.Visible = False
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3.FontBold = True
End Sub



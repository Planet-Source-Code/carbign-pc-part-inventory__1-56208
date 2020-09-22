VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Begin VB.Form frmStocks 
   BackColor       =   &H8000000D&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Stocks Master File System"
   ClientHeight    =   8145
   ClientLeft      =   3495
   ClientTop       =   2295
   ClientWidth     =   8940
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8145
   ScaleWidth      =   8940
   ShowInTaskbar   =   0   'False
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   "C:\Program Files\Microsoft Visual Studio\VB98\PCX\pcx.mdb"
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   345
      Left            =   4200
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   "Products"
      Top             =   9240
      Visible         =   0   'False
      Width           =   1065
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmStocks.frx":0000
      Height          =   4815
      Left            =   0
      OleObjectBlob   =   "frmStocks.frx":0014
      TabIndex        =   11
      Top             =   3240
      Width           =   8895
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   330
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   8940
      _ExtentX        =   15769
      _ExtentY        =   582
      ButtonWidth     =   609
      ButtonHeight    =   582
      Style           =   1
      ImageList       =   "ImageList2"
      DisabledImageList=   "ImageList2"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Add new stocks"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Edit Stocks"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Delete Item"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Save to database"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Search"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cancel"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
      OLEDropMode     =   1
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000D&
      Height          =   1815
      Left            =   0
      TabIndex        =   0
      Top             =   1320
      Width           =   8895
      Begin VB.ComboBox cboItem 
         Appearance      =   0  'Flat
         DataField       =   "ProType"
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
         Height          =   315
         Left            =   1560
         TabIndex        =   10
         Top             =   240
         Width           =   1935
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         DataField       =   "ProPrice"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   1560
         TabIndex        =   9
         Top             =   960
         Width           =   1935
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         DataField       =   "PriUnit"
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
         Height          =   285
         Left            =   1560
         TabIndex        =   8
         Top             =   1320
         Width           =   615
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         DataField       =   "ProName"
         DataSource      =   "Data1"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   1560
         TabIndex        =   7
         Top             =   600
         Width           =   5895
      End
      Begin VB.TextBox txtName 
         Appearance      =   0  'Flat
         DataField       =   "ProName"
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
         Height          =   285
         Left            =   -1920
         TabIndex        =   4
         Top             =   -2640
         Width           =   5355
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "ITEM NAME:"
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
         Height          =   270
         Left            =   120
         TabIndex        =   5
         Top             =   600
         Width           =   1050
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "ITEM TYPE:"
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
         Height          =   270
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   1290
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "QUANTITY:"
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
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   1320
         Width           =   1050
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "STOCK PRICE:"
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
         Height          =   270
         Left            =   120
         TabIndex        =   1
         Top             =   960
         Width           =   1290
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6120
      Top             =   8400
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   43
      ImageHeight     =   37
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStocks.frx":0D3B
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStocks.frx":1199
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   5400
      Top             =   8280
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStocks.frx":15EB
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStocks.frx":6DDD
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStocks.frx":C5CF
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStocks.frx":11DC1
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStocks.frx":175B3
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStocks.frx":1CDA5
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStocks.frx":23607
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image Image1 
      Height          =   915
      Left            =   0
      Picture         =   "frmStocks.frx":23A59
      Top             =   360
      Width           =   8850
   End
End
Attribute VB_Name = "frmStocks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboProducts_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub cboType_KeyPress(KeyAscii As Integer)
KeyAscii = 0
End Sub

Private Sub Form_Load()
  
    


With cboItem
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

Private Sub Form_Unload(Cancel As Integer)
frmMain.Enabled = True: Unload Me: frmMain.Show

End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
    Text5.SetFocus: Text5.Text = ""
End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    Text4.SetFocus
End If

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next
If Button.Index = 2 Then
    'MsgBox "a"
    Data1.Recordset.AddNew
    
ElseIf Button.Index = 3 Then
    'MsgBox "b"
    Data1.Recordset.Edit
  
ElseIf Button.Index = 4 Then
    'MsgBox "c"
    Data1.Recordset.Delete
ElseIf Button.Index = 5 Then
    'MsgBox "d"
    
    
    Data1.Recordset.Update: Data1.Refresh
    r = MsgBox("Add new product again?", vbYesNo, "confirm")
        If (r = vbYes) Then
        Data1.Recordset.AddNew
ElseIf Button.Index = 6 Then
Data1.Recordset.CancelUpdate
    'MsgBox "e"
ElseIf Button.Index = 7 Then
    'MsgBox "f"
    

    End If
        End If

End Sub

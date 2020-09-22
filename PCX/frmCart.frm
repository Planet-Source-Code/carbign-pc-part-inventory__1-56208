VERSION 5.00
Begin VB.Form frmCart 
   Caption         =   "Form1"
   ClientHeight    =   4890
   ClientLeft      =   6240
   ClientTop       =   4905
   ClientWidth     =   9135
   LinkTopic       =   "Form1"
   ScaleHeight     =   4890
   ScaleWidth      =   9135
   Begin VB.Data Data1 
      Caption         =   "Data1"
      Connect         =   "Access"
      DatabaseName    =   ""
      DefaultCursorType=   0  'DefaultCursor
      DefaultType     =   2  'UseODBC
      Exclusive       =   0   'False
      Height          =   300
      Left            =   240
      Options         =   0
      ReadOnly        =   0   'False
      RecordsetType   =   1  'Dynaset
      RecordSource    =   ""
      Top             =   120
      Width           =   1140
   End
End
Attribute VB_Name = "frmCart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

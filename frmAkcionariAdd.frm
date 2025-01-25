VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAkcionariAdd 
   Caption         =   "Lista akcionara"
   ClientHeight    =   4335
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6750
   LinkTopic       =   "Form1"
   ScaleHeight     =   4335
   ScaleWidth      =   6750
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView ctlAkcionariList 
      Height          =   3615
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   6376
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "JMB"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Prezime"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Ime"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Akcije"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Odustani"
      Height          =   375
      Left            =   3360
      TabIndex        =   1
      Top             =   3840
      Width           =   1575
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Dodaj"
      Default         =   -1  'True
      Height          =   375
      Left            =   5040
      TabIndex        =   0
      Top             =   3840
      Width           =   1575
   End
End
Attribute VB_Name = "frmAkcionariAdd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
On Error Resume Next
Me.cmdAdd.Top = Me.ScaleHeight - Me.ctlAkcionariList.Top - Me.cmdAdd.Height
Me.cmdCancel.Top = Me.cmdAdd.Top
Me.cmdAdd.Left = Me.ScaleWidth - Me.ctlAkcionariList.Left - Me.cmdAdd.Width
Me.cmdCancel.Left = Me.cmdAdd.Left - 120 - Me.cmdCancel.Width
Me.ctlAkcionariList.Width = Me.ScaleWidth - 2 * Me.ctlAkcionariList.Left
Me.ctlAkcionariList.Height = -Me.ctlAkcionariList.Top + Me.cmdAdd.Top - Me.ctlAkcionariList.Top
End Sub

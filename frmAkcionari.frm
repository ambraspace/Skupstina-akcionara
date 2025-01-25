VERSION 5.00
Begin VB.Form frmAkcionari 
   Caption         =   "Akcionari"
   ClientHeight    =   4425
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7320
   LinkTopic       =   "Form1"
   ScaleHeight     =   4425
   ScaleWidth      =   7320
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox ctlAkcionariList 
      Height          =   2460
      IntegralHeight  =   0   'False
      ItemData        =   "frmAkcionari.frx":0000
      Left            =   120
      List            =   "frmAkcionari.frx":0002
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   120
      Width           =   5295
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Prekid"
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   2760
      Width           =   1095
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Dodaj"
      Default         =   -1  'True
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   2760
      Width           =   1095
   End
End
Attribute VB_Name = "frmAkcionari"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAdd_Click()
If Me.ctlAkcionariList.Text = "" Then Exit Sub
frmSetNew.txtGlasacPrezime = Left(Me.ctlAkcionariList.Text, InStr(Me.ctlAkcionariList.Text, ", ") - 1)
frmSetNew.txtGlasacIme = Left(Mid(Me.ctlAkcionariList.Text, InStr(Me.ctlAkcionariList.Text, ", ") + 2), InStr(Mid(Me.ctlAkcionariList.Text, InStr(Me.ctlAkcionariList.Text, ", ") + 2), "; ") - 1)
frmSetNew.txtGlasacJMB = Left(Mid(Me.ctlAkcionariList.Text, InStr(Me.ctlAkcionariList.Text, "; ") + 2), InStr(Mid(Me.ctlAkcionariList.Text, InStr(Me.ctlAkcionariList.Text, "; ") + 2), "; ") - 1)
frmSetNew.txtGlasacAkcije = Mid(Me.ctlAkcionariList.Text, InStrRev(Me.ctlAkcionariList.Text, "; ") + 2)
Me.Hide
End Sub

Private Sub cmdCancel_Click()
Me.Hide
End Sub

Private Sub Form_Resize()
On Error Resume Next

Me.cmdCancel.Top = Me.ScaleHeight - Me.cmdCancel.Height - Me.ctlAkcionariList.Top
Me.cmdAdd.Top = Me.cmdCancel.Top
Me.cmdAdd.Left = Me.ScaleWidth - Me.cmdAdd.Width - Me.ctlAkcionariList.Left
Me.cmdCancel.Left = Me.cmdAdd.Left - Me.ctlAkcionariList.Left - Me.cmdCancel.Width
Me.ctlAkcionariList.Width = Me.ScaleWidth - 2 * Me.ctlAkcionariList.Left
Me.ctlAkcionariList.Height = Me.ScaleHeight - 3 * Me.ctlAkcionariList.Top - Me.cmdAdd.Height
End Sub

Public Sub Display()

If rsAkcionari.RecordCount > 0 Then
    rsAkcionari.MoveLast
    rsAkcionari.MoveFirst
    Dim i As Long
    Me.ctlAkcionariList.Clear
    For i = 1 To rsAkcionari.RecordCount
        Me.ctlAkcionariList.AddItem rsAkcionari("Prezime") & ", " & rsAkcionari("Ime") & "; " & rsAkcionari("JMB") & "; " & rsAkcionari("Akcije")
        rsAkcionari.MoveNext
    Next
    Me.Show vbModal
End If
End Sub

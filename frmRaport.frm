VERSION 5.00
Begin VB.Form frmRaport 
   Caption         =   "Izvještaj"
   ClientHeight    =   4635
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7350
   LinkTopic       =   "Form1"
   ScaleHeight     =   4635
   ScaleWidth      =   7350
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtRaport 
      Height          =   3495
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   5055
   End
End
Attribute VB_Name = "frmRaport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Resize()
On Error Resume Next

Me.txtRaport.Width = Me.ScaleWidth - 2 * Me.txtRaport.Left
Me.txtRaport.Height = Me.ScaleHeight - 2 * Me.txtRaport.Top
End Sub

Public Sub Display(SkupstinaName As String, cPitanja As Collection, cGlasaci As Collection, cGlasovi As Collection)
Me.txtRaport = "--------------------" & Chr(13) & Chr(10)
Me.txtRaport = Me.txtRaport & SkupstinaName & Chr(13) & Chr(10)
Me.txtRaport = Me.txtRaport & "--------------------" & Chr(13) & Chr(10) & Chr(13) & Chr(10)

Dim i As Long
For i = 1 To cPitanja.Count
    Me.txtRaport = Me.txtRaport & i & ". " & cPitanja(i).PitanjeText & Chr(13) & Chr(10)
    Dim lZa As Long, lProtiv As Long, lSuzdrzani As Long, lZPS As Long
    lZa = 0
    lProtiv = 0
    lSuzdrzani = 0
    Dim t As cGlas
    For Each t In cGlasovi
        If t.PitanjeID = cPitanja(i).DBID Then
            Select Case t.Glas
                Case "Z"
                    lZa = lZa + cGlasaci("ID" & t.GlasacID).GlasacAkcije
                Case "P"
                    lProtiv = lProtiv + cGlasaci("ID" & t.GlasacID).GlasacAkcije
                Case "S"
                    lSuzdrzani = lSuzdrzani + cGlasaci("ID" & t.GlasacID).GlasacAkcije
            End Select
        End If
    Next
    lZPS = lZa + lProtiv + lSuzdrzani
    Me.txtRaport = Me.txtRaport & "Za: " & Round(lZa / lZPS * 100, 5) & " % = " & lZa & "/" & lZPS & Chr(13) & Chr(10) & _
        "Protiv: " & Round(lProtiv / lZPS * 100, 5) & " % = " & lProtiv & "/" & lZPS & Chr(13) & Chr(10) & _
        "Suzdržani: " & Round(lSuzdrzani / lZPS * 100, 5) & " % = " & lSuzdrzani & "/" & lZPS & Chr(13) & Chr(10) & Chr(13) & Chr(10)
Next
Me.Show vbModal
End Sub


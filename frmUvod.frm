VERSION 5.00
Begin VB.Form frmUvod 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Izbor"
   ClientHeight    =   1935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   5280
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNastavi 
      Caption         =   "Nastavi prethodno glasanje"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   5055
   End
   Begin VB.CommandButton cmdPregledSkupstina 
      Caption         =   "Pregledaj sva skupštinska zasjedanja"
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1320
      Width           =   5055
   End
   Begin VB.CommandButton cmdNovaSkupstina 
      Caption         =   "Zapoèni novo skupštinsko zasjedanje"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5055
   End
End
Attribute VB_Name = "frmUvod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdNastavi_Click()

Dim SkupstinaName As String
Dim cPitanja As New Collection
Dim cGlasaci As New Collection
Dim cGlasovi As New Collection
Dim PitanjeTMP As New cPitanje
Dim GlasacTMP As New cGlasac
Dim GlasTMP As New cGlas

rsSkupstine.MoveLast
TrenutnaSkupstinaID = rsSkupstine("ID")
SkupstinaName = rsSkupstine("Naziv")

rsPitanja.FindFirst "Skupstina=" & TrenutnaSkupstinaID
Do Until rsPitanja.EOF
    PitanjeTMP.DBID = rsPitanja("ID")
    PitanjeTMP.PitanjeText = rsPitanja("Pitanje")
    cPitanja.Add PitanjeTMP, "ID" & PitanjeTMP.DBID
    Set PitanjeTMP = Nothing
    rsPitanja.MoveNext
Loop

rsGlasaci.FindFirst "Skupstina=" & TrenutnaSkupstinaID
Do Until rsGlasaci.EOF
    GlasacTMP.DBID = rsGlasaci("ID")
    GlasacTMP.GlasacAkcije = rsGlasaci("Broj akcija")
    GlasacTMP.GlasacIme = rsGlasaci("Ime")
    GlasacTMP.GlasacJMB = rsGlasaci("JMB")
    GlasacTMP.GlasacPrezime = rsGlasaci("Prezime")
    cGlasaci.Add GlasacTMP, "ID" & GlasacTMP.DBID
    Set GlasacTMP = Nothing
    rsGlasaci.MoveNext
Loop

If rsGlasovi.RecordCount = 0 Then GoTo over
rsGlasovi.FindFirst "Pitanje=" & cPitanja(1).DBID
If rsGlasovi.NoMatch Then GoTo over
    
Do Until rsGlasovi.EOF
    GlasTMP.DBID = rsGlasovi("ID")
    GlasTMP.Glas = rsGlasovi("Glas")
    GlasTMP.GlasacID = rsGlasovi("Glasac")
    GlasTMP.PitanjeID = rsGlasovi("Pitanje")
    cGlasovi.Add GlasTMP, "ID" & GlasTMP.DBID
    Set GlasTMP = Nothing
    rsGlasovi.MoveNext
Loop

Dim iPitanje As Long, iGlasac As Long
iPitanje = Int(cGlasovi.Count / cGlasaci.Count) + 1
iGlasac = (cGlasovi.Count Mod cGlasaci.Count) + 1

Me.Hide
frmGlasanje.NastaviVote SkupstinaName, cPitanja, cGlasaci, cGlasovi, iPitanje, iGlasac
Exit Sub

over:
Me.Hide
frmGlasanje.BeginVote SkupstinaName, cPitanja, cGlasaci
Exit Sub

End Sub

Private Sub cmdNovaSkupstina_Click()
Me.Hide
frmSetNew.Display
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub


Attribute VB_Name = "Glavni"
Option Explicit

Public TrenutnaSkupstinaID
Public MyDirectory As String
Public dbData As Database
Public rsAkcionari As Recordset
Public rsGlasaci As Recordset
Public rsGlasovi As Recordset
Public rsPitanja As Recordset
Public rsSkupstine As Recordset

Public Sub Main()
MyDirectory = CurDir
'MyDirectory = "C:\Documents and Settings\ambra\My Documents\Moji VB primjeri\Bonel skupstina"
If Right(MyDirectory, 1) <> "\" Then MyDirectory = MyDirectory & "\"
Set dbData = OpenDatabase(MyDirectory & "data.mdb", False, False)
Set rsAkcionari = dbData.OpenRecordset("Akcionari", dbOpenDynaset, dbSeeChanges)
Set rsGlasaci = dbData.OpenRecordset("Glasaci", dbOpenDynaset, dbSeeChanges)
Set rsGlasovi = dbData.OpenRecordset("Glasovi", dbOpenDynaset, dbSeeChanges)
Set rsPitanja = dbData.OpenRecordset("Pitanja", dbOpenDynaset, dbSeeChanges)
Set rsSkupstine = dbData.OpenRecordset("Skupstine", dbOpenDynaset, dbSeeChanges)

frmUvod.Show

frmUvod.cmdNastavi.Enabled = fnUnfinishedVote
frmUvod.cmdNovaSkupstina.Enabled = Not fnUnfinishedVote

End Sub

Private Function fnUnfinishedVote() As Boolean
Dim SkupstinaID As Long, PitanjaCount As Long, GlasaciCount As Long, GlasoviCount As Long
Dim StartPitanje As Long

If rsSkupstine.RecordCount = 0 Then Exit Function

rsSkupstine.MoveLast
SkupstinaID = rsSkupstine("ID")

rsPitanja.FindFirst "Skupstina=" & SkupstinaID
StartPitanje = rsPitanja("ID")
Do Until rsPitanja.EOF
    PitanjaCount = PitanjaCount + 1
    rsPitanja.MoveNext
Loop

rsGlasaci.FindFirst "Skupstina=" & SkupstinaID
Do Until rsGlasaci.EOF
    GlasaciCount = GlasaciCount + 1
    rsGlasaci.MoveNext
Loop

If rsGlasovi.RecordCount = 0 Then
    fnUnfinishedVote = True
    Exit Function
End If

rsGlasovi.FindFirst "Pitanje=" & StartPitanje
If rsGlasovi.NoMatch Then
    fnUnfinishedVote = True
    Exit Function
End If
Do Until rsGlasovi.EOF
    GlasoviCount = GlasoviCount + 1
    rsGlasovi.MoveNext
Loop

If GlasoviCount < (PitanjaCount * GlasaciCount) Then fnUnfinishedVote = True

End Function

Public Function fnPause(PauseTime As Single)
Dim Start
Start = Timer
   Do While Timer < Start + PauseTime
      DoEvents
   Loop
End Function


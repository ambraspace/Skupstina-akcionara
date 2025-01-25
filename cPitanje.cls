VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPitanje"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_lDBID As Long
Private m_sPitanjeText As String

Public Property Get DBID() As Long
DBID = m_lDBID
End Property

Public Property Let DBID(sInput As Long)
m_lDBID = sInput
End Property

Public Property Get PitanjeText() As String
PitanjeText = m_sPitanjeText
End Property

Public Property Let PitanjeText(sInput As String)
m_sPitanjeText = sInput
End Property


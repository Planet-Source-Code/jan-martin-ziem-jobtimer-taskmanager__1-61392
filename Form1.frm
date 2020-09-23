VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "JobTimer"
   ClientHeight    =   7365
   ClientLeft      =   4680
   ClientTop       =   2310
   ClientWidth     =   5115
   LinkTopic       =   "Form1"
   ScaleHeight     =   7365
   ScaleWidth      =   5115
   Begin VB.ListBox List1 
      Height          =   5715
      Left            =   0
      TabIndex        =   6
      Top             =   1560
      Width           =   5055
   End
   Begin VB.Timer Timer1 
      Interval        =   250
      Left            =   4200
      Top             =   480
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4935
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Rechts
      Caption         =   "duration:"
      Height          =   255
      Left            =   0
      TabIndex        =   7
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label5 
      Caption         =   "Label5"
      Height          =   255
      Left            =   1080
      TabIndex        =   5
      Top             =   1200
      Width           =   4095
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   960
      Width           =   3975
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   255
      Left            =   1080
      TabIndex        =   3
      Top             =   720
      Width           =   3975
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Rechts
      Caption         =   "started:"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   960
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      Caption         =   "Active Job:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
' Please note:
' this app was written "quick and dirty" because i needed to
' log job timings at a customere here right now.
' Possibly there will follow another version which uses
' description and the listbox.
'
' fields of the type:
' name=name of task/job
' beschreibung = description (not used yet, maybe a summary of done stuff later)
' startzeit = long of starting time (in seconds)
' endzeit = long of job end time (in seconds)
' starzeitstr/endezeitstr = string of times for logfile (job start/stop)
'
' set joblogfile const to fit your needs.
' set maxjobs to fit your needs.
'
' that's all. (c) 2005-29-06 j.-m. ziem, con5 webserices gmbh (www.con5.de),
' under terms of gnu gpl version 2 or any later version (www.gnu.org/gpl)
'
Const maxjobs = 200
Const joblogfile = "C:\con5_joblogfile.txt"


Dim curAction As String




Private Type curjobType
   name As String
   beschreibung As String
   startzeit As Long
   endzeit As Long
   startzeitStr As String
   endzeitstr As String
End Type

Dim shutting_down As Boolean
Dim curjob As curjobType
Dim jobs(maxjobs) As curjobType

Private Sub Command1_Click()
If (curAction = "stopped") And (shutting_down = False) Then
    'nach jobname fragen
    Dim jobname As String
    'wenn gut, dann job starten
    jobname = InputBox("Enter a name for your Task:", "Name your Task", "")
    If jobname <> "" Then
        curAction = "running"
        curjob.name = jobname
        curjob.startzeitStr = Format$(Now(), "hh:mm:ss")
        curjob.startzeit = (Hour(Now) * 3600) + (Minute(Now) * 60) + (Second(Now))
    End If
ElseIf curAction = "running" Then
    'job stoppen
    Dim x As Long
    For x = 0 To (maxjobs - 1)
        If jobs(x).name = "" Then
         'job daten schreiben
         jobs(x).beschreibung = curjob.beschreibung
         jobs(x).endzeit = (Hour(Now) * 3600) + (Minute(Now) * 60) + (Second(Now))
         jobs(x).endzeitstr = Format$(Now(), "hh:mm:ss")
         jobs(x).name = curjob.name
         jobs(x).startzeitStr = curjob.startzeitStr
         jobs(x).startzeit = curjob.startzeit
         Dim tempjob As curjobType
         curjob = tempjob
         List1.AddItem jobs(x).name & " [ duration: " & CLng((jobs(x).endzeit - jobs(x).startzeit) / 60) & " Minutes]"
         List1.ItemData(List1.NewIndex) = x
         Exit For
        End If
    Next x
    curAction = "stopped"
End If
End Sub


Private Sub Form_Load()
shutting_down = False
List1.Clear
curAction = "stopped"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If MsgBox("Do you really want to close this app?", vbYesNo, "Close JobTimer?") = vbNo Then
    Cancel = 1
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
'notfalls laufende jobs stoppen
shutting_down = True
Command1_Click
schreibe_protokoll
End Sub

Private Sub show_curjob()
Dim secs, minutes, hours As Long

If curjob.name <> "" Then
    Label3.Caption = curjob.name
    Label4.Caption = curjob.startzeitStr
    secs = Second(Now)
    minutes = Minute(Now)
    hours = Hour(Now)
    minutes = CLng(((Hour(Now) * 3600) + (Minute(Now) * 60) + (Second(Now)) - curjob.startzeit) / 60)
    'CLng((hours * 3600 + minutes * 60 + secs - curjob.startzeit) / 3600)
    Label5.Caption = "[ " & minutes & " Mins]" & _
                     " [ " & (Hour(Now) * 3600) + (Minute(Now) * 60) + (Second(Now)) - curjob.startzeit & " Secs]"
    Command1.Caption = "End current running Job"
Else
Label3.Caption = "No job active."
Label4.Caption = ""
Label5.Caption = ""
Command1.Caption = "Start a new job"
End If
End Sub

Private Sub Timer1_Timer()
show_curjob
End Sub


Private Sub schreibe_protokoll()
Dim x As Long
Dim temp As Long
Dim dummy As Boolean
For x = 0 To (maxjobs - 1)
  If jobs(x).name <> "" Then
    temp = CLng((jobs(x).endzeit - jobs(x).startzeit) / 60)
    dummy = AppendToLog(joblogfile, "job: " & jobs(x).name & " took " & temp & " Minutes, start/stop: " & jobs(x).startzeitStr & " / " & jobs(x).endzeitstr & vbCrLf)
  End If
Next x
End Sub

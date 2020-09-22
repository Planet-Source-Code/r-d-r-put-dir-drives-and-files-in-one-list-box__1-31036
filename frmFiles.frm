VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Put all files in one list box"
   ClientHeight    =   3795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3795
   ScaleWidth      =   2805
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Files:"
      Height          =   3135
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   2775
      Begin VB.ListBox List1 
         Height          =   2790
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.FileListBox File1 
      Height          =   480
      Left            =   3120
      TabIndex        =   2
      Top             =   2400
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.DirListBox Dir1 
      Height          =   315
      Left            =   3000
      TabIndex        =   1
      Top             =   2640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   3240
      TabIndex        =   0
      Top             =   2640
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   615
      Left            =   0
      TabIndex        =   5
      Top             =   3120
      Width           =   2775
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Function StripPath(T$) As String
Dim x%, ct%
StripPath$ = T$
x% = InStr(T$, "\")
Do While x%
ct% = x%
x% = InStr(ct% + 1, T$, "\")
Loop
If ct% > 0 Then StripPath$ = Mid$(T$, ct% + 1)
End Function

Sub UpdatePath()
Dim I, D, J, K As Integer
For D = 0 To List1.ListCount - 1
List1.RemoveItem "0"
Next D
If Not Right(Dir1.List(-1), 1) = "\" Then
List1.AddItem "[^] .."
End If
For I = 0 To Dir1.ListCount - 1
List1.AddItem "[\] " & StripPath(Dir1.List(I))
Next I
For J = 0 To File1.ListCount - 1
List1.AddItem "[*] " & File1.List(J)
Next J
For K = 0 To Drive1.ListCount - 1
List1.AddItem "[o] " & Drive1.List(K)
Next K
Label1.Caption = Dir1.Path
End Sub
Private Sub Drive1_Change()
On Error GoTo errorhandle
Dir1.Path = Drive1.Drive
Exit Sub
errhandle:
MsgBox Err.Description, vbOKOnly, "error"
End Sub

Private Sub Dir1_Change()
File1.Path = Dir1.Path
UpdatePath
End Sub

Private Sub Form_Load()
Drive1.Visible = False
File1.Visible = False
Dir1.Visible = False
UpdatePath
Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2
End Sub

Private Sub List1_DblClick()
On Error GoTo errhandle
If Right(List1.Text, 2) = ".." Then
Dir1.Path = Dir1.Path & "\.."
ElseIf Left(List1.Text, 3) = "[\]" Then
If Right(Dir1.List(-1), 1) = "\" Then
Dir1.Path = Dir1.Path & Right(List1.Text, Len(List1.Text) - 4)
Else
Dir1.Path = Dir1.Path & "\" & Right(List1.Text, Len(List1.Text) - 4)
End If
ElseIf Left(List1.Text, 3) = "[o]" Then
Drive1.Drive = Right(Left(List1.Text, 6), 2)
Else
MsgBox "File " & Chr(34) & Right(List1.Text, Len(List1.Text) - 4) & _
Chr(34) & " was chosen.", vbInformation, "File Chosen"
End If
Exit Sub
errhandle:
MsgBox Err.Description, vbOKOnly, "ERROR"
End Sub

VERSION 5.00
Begin VB.Form Form_Main 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "nlTrim v0.2"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8865
   Icon            =   "nlTrim.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   8865
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command_About 
      Caption         =   "?About?"
      Height          =   375
      Left            =   5985
      TabIndex        =   2
      Top             =   6120
      Width           =   1365
   End
   Begin VB.CommandButton Command_Copy 
      Caption         =   "Copy"
      Height          =   375
      Left            =   4500
      TabIndex        =   1
      Top             =   6120
      Width           =   1365
   End
   Begin VB.TextBox Text_Status 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   630
      TabIndex        =   6
      Text            =   "Click Paste to start"
      Top             =   6120
      Width           =   2265
   End
   Begin VB.CommandButton Command_Exit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   7470
      TabIndex        =   3
      Top             =   6120
      Width           =   1365
   End
   Begin VB.CommandButton Command_Paste 
      Caption         =   "Paste"
      Height          =   375
      Left            =   3015
      TabIndex        =   0
      Top             =   6120
      Width           =   1365
   End
   Begin VB.TextBox Text_Main 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6090
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   4
      Top             =   0
      Width           =   8835
   End
   Begin VB.Label Label_Status 
      Caption         =   "Status:"
      Enabled         =   0   'False
      Height          =   240
      Left            =   0
      TabIndex        =   5
      Top             =   6165
      Width           =   555
   End
End
Attribute VB_Name = "Form_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Program:  nlTrim.exe
' Origin:   Nov-2023
' Author:   Tip Partridge
' Environment:  Microsoft Visual Basic 6.0 (SP6)
' Description:  This program removes blank lines from text.
'     To use this program first copy text to the clipboard,
'     then click the Paste button. This will paste the text
'     into the window, remove the blank lines, and copy the
'     results back into the clipboard. Be patient and wait
'     until Status becomes Done. The Copy button will also
'     copy the results into the clipboard. Limited to
'     about 65k characters.

Dim s   ' global string variables, "faster" than using a text box.
Dim s2

Private Sub Form_Load()
ver = "nlTrim v1.0"
Form_Main.Show
Command_Paste.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Dim ii As Integer
  For ii = Forms.Count - 1 To 0 Step -1
    Unload Forms(ii)
  Next ii
  End
End Sub

Private Sub Command_About_Click()
Form_About.Show
End Sub

Private Sub Command_Paste_Click()
  Text_Status.Text = "Working: |"
  Text_Main.Text = Clipboard.GetText(vbCFText)
  s2 = Text_Main.Text
  s = ""
  Call trimText
  Text_Main.Text = s
  Clipboard.Clear  'need to do this or SetText doesn't work
  Clipboard.SetText Text_Main.Text, vbCFText
  Text_Main.SelStart = 0
  Text_Status.Text = "Done"
End Sub

Private Sub Command_Copy_Click()
  Clipboard.Clear  'need to do this or SetText doesn't work
  Clipboard.SetText Text_Main.Text, vbCFText
  Text_Status.Text = "Copied"
End Sub

Private Sub Command_Exit_Click()
  End
End Sub

Private Sub trimText()
Dim isNL As Boolean
Dim isTop As Boolean
Dim ii As Long
Dim cc As Long
cc = 0
Dim c As String
isTop = True
For ii = 1 To Len(s2)
  c = Mid(s2, ii, 1)
  If c = vbCr Or c = vbLf Then
    If Not isTop Then
      isNL = True
    End If
  Else
    If isNL Then
      s = s & vbCrLf
      isNL = False
    End If
    isTop = False
    s = s & c
' Spinner
    If ii Mod 999 = 0 Then
      Select Case cc
        Case 0
          c = " / "
          cc = 1
        Case 1
          c = "---"
          cc = 2
        Case 2
          c = " \ "
          cc = 3
        Case 3
          c = " | "
          cc = 4
        Case 4
          c = " / "
          cc = 5
        Case 5
          c = "---"
          cc = 6
        Case 6
          c = " \ "
          cc = 7
        Case 7
          c = " | "
          cc = 0
      End Select
      Text_Status.Text = "Working: " & c
      DoEvents
    End If
  End If
Next ii
Text_Status.Text = "Transfer"
DoEvents
End Sub

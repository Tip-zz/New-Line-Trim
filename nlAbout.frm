VERSION 5.00
Begin VB.Form Form_About 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About nlTrim"
   ClientHeight    =   1800
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4515
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1800
   ScaleWidth      =   4515
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text_About 
      Height          =   1725
      Left            =   45
      MultiLine       =   -1  'True
      TabIndex        =   0
      Text            =   "nlAbout.frx":0000
      Top             =   45
      Width           =   4425
   End
End
Attribute VB_Name = "Form_About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
Text_About.Text = "Program " & Form_Main.Caption & _
"   20-Nov-2023  TEP" & vbCrLf & _
"This program removes blank lines from text. " & _
"To use this program first copy text to the clipboard, " & _
"then click the Paste button. This will paste the text " & _
"into the window, remove the blank lines, and copy the " & _
"results back into the clipboard. Be patient and wait " & _
"until Status becomes Done. The Copy button will also " & _
"copy the results into the clipboard. Limited to " & _
"about 65k characters."
End Sub


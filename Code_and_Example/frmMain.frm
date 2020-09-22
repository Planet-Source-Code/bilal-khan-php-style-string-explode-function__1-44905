VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Example"
   ClientHeight    =   7020
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10185
   LinkTopic       =   "Form1"
   ScaleHeight     =   7020
   ScaleWidth      =   10185
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtResults 
      Height          =   3855
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   6
      Top             =   3000
      Width           =   9855
   End
   Begin VB.CommandButton cmdSeparate 
      Caption         =   "Separate"
      Height          =   375
      Left            =   8520
      TabIndex        =   4
      Top             =   1800
      Width           =   1215
   End
   Begin VB.TextBox txtDelimiter 
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Text            =   " "
      Top             =   1200
      Width           =   8295
   End
   Begin VB.TextBox txtString 
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Text            =   "The Quick Fox, Jumped Over The Lazy Dog, and fell in a ditch."
      Top             =   360
      Width           =   8295
   End
   Begin VB.Label Label3 
      Caption         =   "Result ->"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2640
      Width           =   1215
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   10200
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Label Label2 
      Caption         =   "Delimiter        ->"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "Type a String ->"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdSeparate_Click()
    Dim i As Integer
    Dim TotalSplits As Integer
    Dim args(30) As String
    
    txtResults.Text = "" ' Clear the Results Text
    '--------------------
    TotalSplits = explodeargs(txtString.Text, args, txtDelimiter.Text)
    
    If TotalSplits > 1 Then
        txtResults.Text = "There are " & TotalSplits + 1 & " separated strings in the array." & vbNewLine & vbNewLine
    Else
        txtResults.Text = "There are " & TotalSplits & " separated strings in the array." & vbNewLine & vbNewLine
    End If
    txtResults.Text = txtResults.Text & "The Splits are:" & vbNewLine & "-----------------------" & vbNewLine
    
    For i = 0 To TotalSplits Step 1
        txtResults.Text = txtResults.Text & args(i) & vbNewLine
    Next
    
End Sub

Public Function explodeargs(ByVal s, ByRef arg() As String, Optional ByVal splitpattern As String = "*|*|*") As Integer
            On Error Resume Next ' on error, just quietly continue on to next operation
            Dim i
            For i = 0 To 30 Step 1 ' assume maximum 30 option strings
                arg(i) = Null ' null everything in the array first
            Next

            On Error GoTo ExtractionFinished
            For i = 0 To 30 Step 1
                arg(i) = Split(s, splitpattern)(i) ' begin splitting, if an error occurs, then auto exit
            Next

ExtractionFinished:
            explodeargs = i - 1
End Function

VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Internet Connection Speed Test"
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3735
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   3735
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   250
      Left            =   120
      Top             =   120
   End
   Begin VB.Frame Frame1 
      Caption         =   "Results"
      Height          =   855
      Left            =   480
      TabIndex        =   1
      Top             =   3240
      Width           =   2895
      Begin VB.Label Label1 
         Caption         =   "Checking..."
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   2655
      End
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   3105
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3720
      ExtentX         =   6562
      ExtentY         =   5477
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
WebBrowser1.Navigate ("Http://myvb.tripod.com/test.gif")
Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
Static TTime
If WebBrowser1.Busy Then
TTime = TTime + 0.25
Else
TTime = TTime / 1000

v = (170000 / TTime) / 1000

If (v < 1701) Then Label1.Caption = "Less than 14,400 KB/S"

If (v > 1900) Then Label1.Caption = "Around 16,800 KB/S"

If (v > 2200) Then Label1.Caption = "Around 19,200 KB/S"
If (v > 2500) Then Label1.Caption = "Around 21,600 KB/S"
If (v > 2800) Then Label1.Caption = "Around 24,000 KB/S"
If (v > 3000) Then Label1.Caption = "Around 26,400 KB/S"
If (v > 3400) Then Label1.Caption = "Around 28,800 KB/S"
If (v > 4000) Then Label1.Caption = "Around 33,600 KB/S"
If (v > 4200) Then Label1.Caption = "Around 36,000 KB/S"
If (v > 4500) Then Label1.Caption = "Around 38,000 KB/S"
If (v > 4800) Then Label1.Caption = "Around 40,000 KB/S"
If (v > 5000) Then Label1.Caption = "Around 42,000  KB/S"
If (v > 5300) Then Label1.Caption = "Around 44,000 KB/S"
If (v > 5500) Then Label1.Caption = "Around 46,000 KB/S"
If (v > 5800) Then Label1.Caption = "Around 48,000 KB/S"
If (v > 6000) Then Label1.Caption = "Around 50,000 KB/S"
If (v > 6200) Then Label1.Caption = "Around 52,000 KB/S"
If (v > 7500) Then Label1.Caption = "More than 64,000 KB/S"
If (v > 15000) Then Label1.Caption = "More than 128,000 KB/S"

Timer1.Enabled = False
End If
End Sub

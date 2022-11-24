VERSION 5.00
Begin VB.Form frmAnalysis 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "24-bit ADC"
   ClientHeight    =   4890
   ClientLeft      =   555
   ClientTop       =   1455
   ClientWidth     =   2415
   ForeColor       =   &H00000000&
   Icon            =   "Analysis.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   326
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   161
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkSave 
      Caption         =   "Save To File"
      Height          =   375
      Left            =   360
      TabIndex        =   17
      Top             =   4440
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      Caption         =   "Analysis"
      Height          =   2295
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   2175
      Begin VB.TextBox txterrorbit 
         Height          =   420
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   16
         Text            =   " "
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox txtdiff 
         Height          =   420
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   0
         Text            =   " "
         Top             =   1255
         Width           =   1215
      End
      Begin VB.TextBox txtmax 
         Height          =   420
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   1
         Text            =   " "
         Top             =   710
         Width           =   1215
      End
      Begin VB.TextBox txtmin 
         Height          =   345
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   2
         Text            =   " "
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Bit Error"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   1913
         Width           =   555
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Diff"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   1368
         Width           =   240
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Max"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   823
         Width           =   300
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Min"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   315
         Width           =   255
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Setting"
      Height          =   1815
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2175
      Begin VB.ComboBox cboPort 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   240
         Width           =   1092
      End
      Begin VB.ComboBox cboBitRate 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   720
         Width           =   1092
      End
      Begin VB.TextBox txtsamplingrate 
         ForeColor       =   &H00FF00FF&
         Height          =   420
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   3
         Text            =   " "
         Top             =   1200
         Width           =   1092
      End
      Begin VB.Label lblbaud 
         AutoSize        =   -1  'True
         Caption         =   "BaudRate"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   780
         Width           =   720
      End
      Begin VB.Label lblcom 
         AutoSize        =   -1  'True
         Caption         =   "Com Port"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   300
         Width           =   645
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "S Rate"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   1313
         Width           =   495
      End
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Height          =   195
      Left            =   480
      TabIndex        =   10
      Top             =   2640
      Width           =   45
   End
End
Attribute VB_Name = "frmAnalysis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cboBitRate_Change()
Baud = Val(frmAnalysis.cboBitRate.Text)
End Sub

Public Sub cboPort_Change()
Port = Val(Right$(frmAnalysis.cboPort.Text, 1))
End Sub

Private Sub cboScanTime_Change()
ScanTime = Val(Left$(frmAnalysis.cboPort.Text, 1))
End Sub

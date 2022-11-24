VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmDraw 
   AutoRedraw      =   -1  'True
   Caption         =   "24 - Bit ADC"
   ClientHeight    =   7095
   ClientLeft      =   4335
   ClientTop       =   2100
   ClientWidth     =   9495
   Icon            =   "Draw.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7095
   ScaleWidth      =   9495
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Caption         =   " "
      Height          =   855
      Left            =   240
      TabIndex        =   4
      Top             =   4560
      Width           =   375
      Begin VB.CommandButton cmdUp 
         Height          =   375
         Left            =   0
         Picture         =   "Draw.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Increase Y"
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton cmdDown 
         Height          =   375
         Left            =   0
         Picture         =   "Draw.frx":0519
         Style           =   1  'Graphical
         TabIndex        =   5
         ToolTipText     =   "Decrease Y"
         Top             =   480
         Width           =   375
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   " "
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   4440
      Width           =   855
      Begin VB.CommandButton cmdRight 
         Height          =   375
         Left            =   480
         Picture         =   "Draw.frx":0723
         Style           =   1  'Graphical
         TabIndex        =   3
         ToolTipText     =   "Increase X"
         Top             =   0
         Width           =   375
      End
      Begin VB.CommandButton cmdLeft 
         Height          =   375
         Left            =   0
         Picture         =   "Draw.frx":0935
         Style           =   1  'Graphical
         TabIndex        =   2
         ToolTipText     =   "Decrease X"
         Top             =   0
         Width           =   375
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   7200
      TabIndex        =   0
      Top             =   6480
      Width           =   1575
      Begin VB.CommandButton cmdcolor 
         Caption         =   "Color"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   9
         ToolTipText     =   "Change Color"
         Top             =   0
         Width           =   615
      End
      Begin VB.CommandButton cmdstop 
         Caption         =   "Stop"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         Picture         =   "Draw.frx":0B45
         TabIndex        =   8
         ToolTipText     =   "Stop Serial Port"
         Top             =   0
         Width           =   495
      End
      Begin VB.CommandButton cmdrun 
         Caption         =   "Run "
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1080
         Picture         =   "Draw.frx":0F87
         TabIndex        =   7
         ToolTipText     =   "Start Serial Port"
         Top             =   0
         Width           =   495
      End
   End
   Begin MSComDlg.CommonDialog CD 
      Left            =   2760
      Top             =   6000
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   840
      Top             =   2880
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
End
Attribute VB_Name = "frmDraw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Nx As Integer, Ny As Integer
Dim Max As Long, Min As Long
Dim d1 As Byte
Dim d2 As Byte
Dim d3 As Byte
Dim SamplingRate As Byte '0,1,2,3
Dim SRATE As Integer '1024,512,256,128
Dim PRESRATE As Integer '1024,512,256,128

Public ScanTime As Integer
Dim state As Byte
Dim ByteCount As Integer
Dim mydata As Long
Public Port As Integer
Public Baud As Long
Public Xmin As Long
Public Ymin As Long
Public Xmax As Long
Public Ymax As Long
Public Save As Boolean

Public Sub cmdstart()


ByteCount = 0
state = WAITING
mydata = 0

Min = 2 ^ 24 '8388608
Max = 0 '-8388608



Port = Val(Right$(frmAnalysis.cboPort.Text, 1))
Baud = Val(frmAnalysis.cboBitRate.Text)
ScanTime = Val(Left$(frmAnalysis.cboPort.Text, 1))

MSComm1.CommPort = Port 'txtcom.Text  '  Use COM X
MSComm1.Settings = Baud & ",N,8,1" 'baud,parity,data,stop
MSComm1.InputMode = comInputModeText    ' Open the port
MSComm1.RThreshold = 1      ' Fire Rx Event Every one Bytes
MSComm1.InputLen = 1        ' When Inputting Data, Input 1 Bytes at a time

MSComm1.PortOpen = True

Ymin = -2 ^ Ny
Ymax = 2 ^ Ny
Xmin = 0
Xmax = Nx * 50 '1024

DefineLayout
DrawGrids frmDraw, udtMyGraphLayout
DrawTicks frmDraw
End Sub

Private Sub cmdcolor_Click()
CD.ShowColor
frmDraw.BackColor = CD.Color
End Sub

Private Sub cmdrun_Click()
cmdstart
End Sub

Private Sub cmdstop_Click()
    MSComm1.PortOpen = False
End Sub

Public Sub Form_Load()

Load frmAnalysis
frmAnalysis.Show

ByteCount = 0
state = WAITING
mydata = 0
Ny = 6 '12
Nx = 10 '1
NN = Nx
SR = 50 '1024


Open "Output.txt" For Output As #1

frmAnalysis.cboBitRate.AddItem ("110")
frmAnalysis.cboBitRate.AddItem ("300")
frmAnalysis.cboBitRate.AddItem ("1200")
frmAnalysis.cboBitRate.AddItem ("2400")
frmAnalysis.cboBitRate.AddItem ("4800")
frmAnalysis.cboBitRate.AddItem ("9600")
frmAnalysis.cboBitRate.AddItem ("19200")
frmAnalysis.cboBitRate.AddItem ("38400")
frmAnalysis.cboBitRate.AddItem ("57600")
frmAnalysis.cboBitRate.AddItem ("115200")
frmAnalysis.cboBitRate.AddItem ("230400")
frmAnalysis.cboBitRate.AddItem ("460800")
frmAnalysis.cboBitRate.AddItem ("921600")

frmAnalysis.cboPort.AddItem ("COM1")
frmAnalysis.cboPort.AddItem ("COM2")
frmAnalysis.cboPort.AddItem ("COM3")
frmAnalysis.cboPort.AddItem ("COM4")

frmAnalysis.cboBitRate.Text = frmAnalysis.cboBitRate.List(9)
frmAnalysis.cboPort.Text = frmAnalysis.cboPort.List(0)
End Sub
Private Sub Form_Resize()
'Align Fram1
Frame1.Left = frmDraw.Width - Frame1.Width - 400
Frame1.Top = (frmDraw.Height - Frame1.Height) - 600
'Align Fram2
Frame2.Left = (frmDraw.Width - Frame2.Width) / 2 + 2000
Frame2.Top = (frmDraw.Height - Frame2.Height) - 600
'Align Fram3
Frame3.Top = (frmDraw.Height - Frame3.Height) / 2 - 400

Ymin = -2 ^ Ny
Ymax = 2 ^ Ny
Xmin = 0
Xmax = Nx * 50 'Nx * 1024

DefineLayout
DrawGrids frmDraw, udtMyGraphLayout
DrawTicks frmDraw
End Sub
Public Sub MSComm1_OnComm()
Save = frmAnalysis.chkSave.Value
If MSComm1.CommEvent = comEvReceive Then
    Select Case state
        Case WAITING:
             If Asc(MSComm1.Input) = &HFF Then
                state = HEADER1
                mydata = 0
                If ByteCount = Xmax Then
                    frmDraw.Cls
                    Xmax = SRATE * Nx
                    DefineLayout
                    DrawGrids frmDraw, udtMyGraphLayout
                    DrawTicks frmDraw
                    
                    ByteCount = 0
                    frmAnalysis.txtmin.Text = Min
                    frmAnalysis.txtmax.Text = Max
                    frmAnalysis.txtdiff.Text = Max - Min
                    'frmAnalysis.txterrorbit.Text = Log(Max - Min) / Log(2)
                    Min = 8388608
                    Max = -8388608
                End If
                ByteCount = ByteCount + 1
            Else
                state = WAITING
            End If
        Case HEADER1:
            If Asc(MSComm1.Input) = &HEB Then
                state = HEADER2
            Else
                state = WAITING
            End If
        Case HEADER2:
                SamplingRate = Asc(MSComm1.Input)
                state = DATA1
        Case DATA1:
            d1 = Asc(MSComm1.Input)
            mydata = d1
            state = DATA2
        Case DATA2:
            d2 = Asc(MSComm1.Input)
            mydata = mydata + (CLng(d2) * 256)
            state = DATA3
        Case DATA3:
            d3 = Asc(MSComm1.Input)
            mydata = mydata + (CLng(d3) * 256 * 256)
            state = FOOTER
        Case FOOTER:
            If Asc(MSComm1.Input) = 13 Then
                If (mydata And &H800000) Then
                   mydata = (mydata Xor &HFFFFFF) + 1
                   mydata = (mydata And &H7FFFFF)
                   mydata = -mydata
                End If
                ' Save to File if check box is checked
                If (Save = True) Then
                    Print #1, mydata
                End If

                If mydata > Max Then
                    Max = mydata
                End If
                If mydata < Min Then
                    Min = mydata
                End If
                
                SRATE = 50 '1024 '8 * 128 / (2 ^ SamplingRate)
                SR = SRATE
                
                
                frmAnalysis.txtsamplingrate.ForeColor = RGB(200, 100, 0)
                frmAnalysis.txtsamplingrate = SRATE
                
                If mydata > Ymax Then mydata = Ymax 'clip data
                If mydata < Ymin Then mydata = Ymin 'clip data
                
                If (ByteCount = 1) Then
                    GoToXY frmDraw, ByteCount, mydata
                End If
                val_X = ByteCount
                val_Y = mydata
                XRatio = (val_X - val_XMin) / val_XRange
                YRatio = (val_Y - val_YMin) / val_YRange
                frmDraw.Line -(twp_X0 + XRatio * twp_XRange, twp_Y0 - YRatio * twp_YRange), RGB(0, 0, 0)
                state = WAITING
            Else
                state = WAITING
            End If
Default:
        state = WAITING
    End Select
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload frmAnalysis
Unload Me
Close #1
End Sub
Public Sub DefineLayout()
With udtMyGraphLayout
  .XTitle = "X-Axis"
  .Ytitle = "Y-Axis"
  .blnOrigin = True
  .blnGridLine = True
  .X0 = Xmin
  .X1 = Xmax
  .Y0 = Ymin
  .Y1 = Ymax
End With
End Sub
Public Sub cmdLeft_Click()
    Nx = Nx - 10 '1
    If Nx <= 0 Then Nx = 10 '1
    Xmax = Nx * 50: NN = Nx '1024
    udtMyGraphLayout.X1 = Xmax
    DrawGrids frmDraw, udtMyGraphLayout
    DrawTicks frmDraw
    ByteCount = 0
End Sub
Public Sub cmdRight_Click()
    Nx = Nx + 10 '1
    If Nx >= 300 Then Nx = 300 '10
    Xmax = Nx * 50: NN = Nx '1024
    udtMyGraphLayout.X1 = Xmax
    DrawGrids frmDraw, udtMyGraphLayout
    DrawTicks frmDraw
    ByteCount = 0
End Sub
Public Sub cmdUp_Click()
    Ny = Ny + 1
    If Ny >= 24 Then Ny = 24
    Ymin = -2 ^ Ny
    Ymax = 2 ^ Ny
    udtMyGraphLayout.Y1 = Ymax
    udtMyGraphLayout.Y0 = Ymin
    DrawGrids frmDraw, udtMyGraphLayout
    DrawTicks frmDraw
    ByteCount = 0
End Sub
Public Sub cmdDown_Click()
    Ny = Ny - 1
    If Ny <= 3 Then Ny = 3
    Ymin = -2 ^ Ny
    Ymax = 2 ^ Ny
    udtMyGraphLayout.Y1 = Ymax
    udtMyGraphLayout.Y0 = Ymin
    DrawGrids frmDraw, udtMyGraphLayout
    DrawTicks frmDraw
    ByteCount = 0
End Sub

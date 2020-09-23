VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Options"
   ClientHeight    =   2115
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3855
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2115
   ScaleWidth      =   3855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   255
      Left            =   2040
      TabIndex        =   11
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Frame Frame2 
      Caption         =   "Size"
      Height          =   1455
      Left            =   1560
      TabIndex        =   5
      Top             =   240
      Width           =   2175
      Begin VB.CommandButton cmdSet 
         Caption         =   "Set"
         Height          =   255
         Left            =   480
         TabIndex        =   8
         Top             =   1080
         Width           =   1095
      End
      Begin VB.TextBox txtWidth 
         Height          =   285
         Left            =   840
         TabIndex        =   7
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtHeight 
         Height          =   285
         Left            =   840
         TabIndex        =   6
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label lblWidth 
         BackStyle       =   0  'Transparent
         Caption         =   "Width:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   720
         Width           =   615
      End
      Begin VB.Label lblHeight 
         BackStyle       =   0  'Transparent
         Caption         =   "height:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Mode"
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
      Begin VB.CheckBox AutoSize 
         Caption         =   "AutoSize"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   975
      End
      Begin VB.CheckBox Zoom 
         Caption         =   "Zoom"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   1440
         Width           =   735
      End
      Begin VB.CheckBox Clip 
         Caption         =   "Clip"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Width           =   735
      End
      Begin VB.CheckBox Stretch 
         Caption         =   "Stretch"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
txtHeight.Text = frmMain.OLE1.Height ' show the ole´s height
txtWidth.Text = frmMain.OLE1.Width ' show the loe´s width

If frmMain.OLE1.SizeMode = 1 Then ' if sizemode is stretch
Stretch = 1 ' then stretch value is checked
ElseIf frmMain.OLE1.SizeMode = 2 Then ' if sizemode is clip
Clip = 1 ' then clip value is checked
ElseIf frmMain.OLE1.SizeMode = 3 Then ' if sizemode is AutoSize
AutoSize = 1 ' then AutoSize value is checked
ElseIf frmMain.OLE1.SizeMode = 4 Then ' if sizemode is zoom
Zoom = 1 ' then zoom value is checked
End If

End Sub

Private Sub AutoSize_Click()
On Error Resume Next ' In case of error
Dim Success As Long
If AutoSize = 1 Then ' if autosize value is checked
frmMain.OLE1.SizeMode = 3 ' make sizemode autosize
Clip = 0 ' make clip unchecked
Stretch = 0 ' make Stretch unchecked
Zoom = 0 ' make Zoom unchecked
End If
End Sub

Private Sub Clip_Click()
On Error Resume Next ' In case of error
Dim Success As Long
If Clip = 1 Then ' if Clip value is checked
frmMain.OLE1.SizeMode = 2 ' make sizemode Clip
Stretch = 0 ' make Stretch unchecked
AutoSize = 0 ' make AutoSize unchecked
Zoom = 0 ' make Zoom unchecked
End If
End Sub



Private Sub Stretch_Click()
On Error Resume Next ' In case of error
Dim Success As Long
If Stretch = 1 Then ' if Stretch value is checked
frmMain.OLE1.SizeMode = 1 ' make Stretch Clip
Clip = 0 ' make Clip unchecked
AutoSize = 0 ' make AutoSize unchecked
Zoom = 0 ' make Zoom unchecked
End If
End Sub

Private Sub Zoom_Click()
On Error Resume Next ' In case of error
Dim Success As Long
If Zoom = 1 Then  ' if Zoom value is checked
frmMain.OLE1.SizeMode = 4 ' make Zoom Clip
Clip = 0 ' make Clip unchecked
AutoSize = 0 ' make AutoSize unchecked
Stretch = 0 ' make Stretch unchecked
End If
End Sub

Private Sub cmdSet_Click()
On Error Resume Next ' In case of error
frmMain.OLE1.Height = txtHeight.Text ' change height of ole to added height
frmMain.OLE1.Width = txtHeight.Width ' change Width of ole to added Width
End Sub


Private Sub cmdOK_Click()
Unload Me ' closes the option form
End Sub


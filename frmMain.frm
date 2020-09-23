VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PowerPoint demo"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4725
   Icon            =   "FRMMAIN.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   4725
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Player"
      Height          =   3375
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4695
      Begin VB.OLE OLE1 
         Appearance      =   0  'Flat
         AutoActivate    =   3  'Automatic
         BackStyle       =   0  'Transparent
         Class           =   "PowerPoint.Show.8"
         Height          =   3015
         Left            =   120
         OleObjectBlob   =   "FRMMAIN.frx":000C
         SizeMode        =   1  'Stretch
         SourceDoc       =   "C:\Squash Productions\PSC Demos\PowerPoint\example1.ppt"
         TabIndex        =   2
         Top             =   240
         Width           =   4455
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Menu"
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   3600
      Width           =   4695
      Begin VB.CommandButton cmdOpen 
         Caption         =   "&Open"
         Height          =   495
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1695
      End
      Begin VB.CommandButton cmdOptions 
         Caption         =   "&Options"
         Height          =   495
         Left            =   2760
         TabIndex        =   3
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Label lblText 
      BackStyle       =   0  'Transparent
      Caption         =   "Doble click the file to view it."
      Height          =   255
      Left            =   1320
      TabIndex        =   5
      Top             =   4680
      Width           =   2295
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************
'*                   PowerPoint Viewer!              *
'*  Hi, this is an example will show you how to view *
'* powerpoint files inside your VB projects.         *
'* It requires you have PowerPoint installed tho     *
'* it is commented but if you need any help contact  *
'* me.                                               *
'* Thanks a lot, hope this helps. please vote!       *
'* Contact me if you have any trouble:               *
'*                                                   *
'* Email: Squash@cv.cl                               *
'* web site:  http://www.SquashProductions.com       *
'*****************************************************



Private Sub cmdOpen_Click()
On Error Resume Next ' In case of error
Dim sOpen As SelectedFile

FileDialog.sInitDir = OptDefPath
FileDialog.flags = OFN_EXPLORER Or OFN_LONGNAMES Or OFN_HIDEREADONLY Or OFN_ALLOWMULTISELECT ' options
FileDialog.sDlgTitle = "Open PowerPoint file" ' title
FileDialog.sFilter = "PowerPoint file (*.ppt)" & Chr$(0) & "*.ppt" ' filter
     sOpen = ShowOpen(Me.hWnd) ' show open menu

  Screen.MousePointer = 11 ' make mouse icon hour glass
        OLE1.CreateLink FileDialog.sFileTitle ' show the selected file


 


Screen.MousePointer = 0 ' make mouse icon normal

End Sub

Private Sub cmdOptions_Click()
frmOptions.Show ' show the option form
End Sub

Private Sub Form_Unload(Cancel As Integer)
End
End Sub

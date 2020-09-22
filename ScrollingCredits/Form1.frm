VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "The ugliest form in the world"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   1335
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "Form1.frx":0000
      Top             =   120
      Width           =   4455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Do It Baby"
      Height          =   495
      Left            =   1560
      TabIndex        =   0
      Top             =   2040
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    frmAbout.Show , Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim f As Form
    
    For Each f In Forms
        If Not f Is Me Then
            Unload f
        End If
    Next f
End Sub


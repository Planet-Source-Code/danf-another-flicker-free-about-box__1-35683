VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   0  'None
   Caption         =   "<!-- Insert Caption Here -->"
   ClientHeight    =   4860
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6720
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4860
   ScaleWidth      =   6720
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer ReDrawTimer 
      Interval        =   1
      Left            =   6120
      Top             =   4320
   End
   Begin VB.PictureBox picBackBuffer 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      FillColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   2100
      Left            =   195
      ScaleHeight     =   2100
      ScaleWidth      =   5640
      TabIndex        =   2
      Top             =   2250
      Visible         =   0   'False
      Width           =   5640
   End
   Begin VB.PictureBox picOut 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4500
      Left            =   0
      Picture         =   "frmAbout.frx":0000
      ScaleHeight     =   4500
      ScaleWidth      =   6000
      TabIndex        =   1
      Top             =   0
      Width           =   6000
   End
   Begin VB.PictureBox picBuffer 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      FillColor       =   &H00404040&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4215
      Left            =   120
      ScaleHeight     =   4215
      ScaleWidth      =   6495
      TabIndex        =   0
      Top             =   240
      Visible         =   0   'False
      Width           =   6495
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'if you don't want to make the scrolling stop when you mouseover
'the form, or you don't want to redistribute ssubtimer,
'[or you've no idea how to download and register ssubtimer]
'change the following #const to False
#Const HASH_CON_TRACKMOUSE = True

#If HASH_CON_TRACKMOUSE Then
    Implements ISubclass
    
    Private Const WM_MOUSELEAVE = &H2A3 'the messages we want to catch
    Private Const WM_MOUSEMOVE = &H200
    
    Private Declare Function TrackMouseEvent Lib "comctl32.dll" Alias "_TrackMouseEvent" (ByRef lpEventTrack As TrackMouseEvent) As Long
    Private Const TME_LEAVE = &H2
    
    Private Type TrackMouseEvent
        cbSize As Long          'size of the structure
        dwFlags As Long         'bitset of flags
        hwnd As Long            'handle to the window to track
        dwHoverTime As Long     'how long to wait before posting a hover message (miliseconds)
    End Type

    Private mbolMouseOver As Boolean    'whether or not the mouse thought to be currently over us
#End If

Private Declare Function BitBlt Lib "gdi32" (ByVal hdcDest As Long, ByVal XDest As Long, ByVal YDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hdcSrc As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Const SRCCOPY = &HCC0020

Private Type CreditLine
    Text As String
    Bold As Boolean
End Type

Private marrLines(1 To 3000) As CreditLine 'the credits
Private mlngNumLines As Long            'how many lines of credits there are
Private msngYPos As Single              'the current Y pos

Private mlngTextAreaWidth As Long       'dimensions of the text area,
Private mlngTextAreaHeight As Long      'stored as module level variables
Private mlngTextAreaTop As Long         'to speed stuff up
Private mlngTextAreaLeft As Long
Private mlngTextAreaBottom As Long
Private mlngTextAreaRight As Long

Private mlngVersionTop As Long          'where to display
Private mlngVersionLeft As Long         'the version string

Private mstrAppVersion As String        'version string to display

Private Sub SetUpVariables()
'    Dim lstrLine As String
'    'read the credits text file into the array
'    Open (App.Path & "\credits.txt") For Input As #1
'    mlngNumLines = 0
'    Do Until EOF(1)
'        mlngNumLines = mlngNumLines + 1
'        Line Input #1, lstrLine
'        marrLines(mlngNumLines).Text = lstrLine
'    Loop
'    Close #1

    'make up the credits array. You COULD read it from a file (code is above)
    '   or read it from a resource file
    '   or pass it into the form
    '   or download it off the web
    '   etc, etc
    'but for this app I'm just going to make it all up inline
    AddCreditLine "Copyright 2002 Advanced Publishing Systems Pty Ltd", True
    AddCreditLine "Blah blah blah blah blah"
    AddCreditLine
    AddCreditLine String$(50, "-")
    AddCreditLine
    AddCreditLine "even more blah"
    AddCreditLine "This is supposed to be a fairly interesting line"
    AddCreditLine "or two thanking the academy for the free booze"
    AddCreditLine String$(50, "-")
    AddCreditLine "even more blah"
    AddCreditLine String$(50, "-")
    AddCreditLine
    AddCreditLine "Thats All Folks!", True
    
    'get the app version to display from somewhere
    mstrAppVersion = "RC2E - Version " & App.Major & "." & App.Minor & "." & App.Revision
    'work out where you want to display it.
    'in this case its going to be above the top right corner of the text area
    mlngVersionLeft = (mlngTextAreaRight) - picOut.TextWidth(mstrAppVersion)
    mlngVersionTop = mlngTextAreaTop - 3 - picOut.TextHeight(mstrAppVersion)
End Sub

Private Sub AddCreditLine(Optional pstrText As String = "", Optional pbolBold As Boolean = False)
    'bump the line count up by one
    mlngNumLines = mlngNumLines + 1
    'make sure they've given us something
    If Len(pstrText) > 0 Then
        With marrLines(mlngNumLines)
            .Text = pstrText
            .Bold = pbolBold
        End With
    End If
End Sub

Private Sub Form_Load()
    Dim lstrVersion As String
    
    #If HASH_CON_TRACKMOUSE Then
        'start subclassing the picture for mouse leave messages
        AttachMessage Me, picOut.hwnd, WM_MOUSELEAVE
        'may as well do mouse move via the same method. could just use picOut_MouseMove, but this way is a bit easier
        AttachMessage Me, picOut.hwnd, WM_MOUSEMOVE
    #End If
    
    'center and resize the form, all in one foul swoop
    Me.Move (Screen.Width - picOut.Width) / 2, (Screen.Height - picOut.Height) / 2, picOut.Width, picOut.Height
    
    'make sure the buffer picture is the same size as the back buffer picture
    picBuffer.Move 0, 0, picBackBuffer.Width, picBackBuffer.Height
    
    'make sure everything is dealing in pixels, not twips
    Me.ScaleMode = vbPixels
    picBuffer.ScaleMode = vbPixels
    picOut.ScaleMode = vbPixels
    picBackBuffer.ScaleMode = vbPixels
    'set a few properties of the buffer
    picBuffer.ForeColor = vbBlack
    picBuffer.BackColor = vbWhite
    picBuffer.AutoRedraw = True
    'hide the buffer
    picBuffer.Visible = False
    
    'grab the dimensions of the background area
    mlngTextAreaHeight = picBackBuffer.Height
    mlngTextAreaLeft = picBackBuffer.Left
    mlngTextAreaTop = picBackBuffer.Top
    mlngTextAreaWidth = picBackBuffer.Width
    mlngTextAreaBottom = mlngTextAreaTop + mlngTextAreaHeight
    mlngTextAreaRight = mlngTextAreaLeft + mlngTextAreaWidth
    
    'copy a chunk of the main picture to the background buffer *before* we start drawing on the main picture
    BitBlt picBackBuffer.hDC, 0, 0, mlngTextAreaWidth, mlngTextAreaHeight, picOut.hDC, mlngTextAreaLeft, mlngTextAreaTop, SRCCOPY
    
    'set the initial horizontal drawing position to be about 1/4 the way down the drawing area
    'this gives the user time to go "huh? its scrolling? whats that first line say?" before the first line disappears
    msngYPos = CLng(mlngTextAreaHeight * (1 / 4))
    
    'setup the text and stuff to display
    SetUpVariables
    
    '20 sounds like a nice round number :)
    ReDrawTimer.Interval = 20
    ReDrawTimer.Enabled = True

End Sub

Private Sub Form_Activate()
    'draw the credits when they swap back to this form
    DrawCredits
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    'check for escape key to close the form
    If KeyCode = vbKeyEscape Then
        ReDrawTimer.Enabled = False
        Unload Me
    End If
End Sub

Private Sub DrawCredits()
    Dim llngCount As Long
    Dim llngFontSize As Long
    
    'not a whole lot we can do about errors
    On Error Resume Next
    
    'Draw the background to the buffer. It's only had to be written once, so we'll just re-blit it over again and agin.
    BitBlt picBuffer.hDC, 0, 0, picBuffer.ScaleWidth, picBuffer.ScaleHeight, picBackBuffer.hDC, 0, 0, SRCCOPY
    
    'remember where we are supposed to start from
    picBuffer.CurrentY = msngYPos
    
    'Do the following for each line of text in our credits message...
    For llngCount = 1 To mlngNumLines
        'set bolding based on what the array says
        picBuffer.FontBold = marrLines(llngCount).Bold
        
        'Set the starting location of where to print the text
        picBuffer.CurrentY = picBuffer.CurrentY + 2 '(this is to bump the line spacing a bit)
        picBuffer.CurrentX = (picBuffer.ScaleWidth - picBuffer.TextWidth(marrLines(llngCount).Text)) / 2
        
        'Send the text to the buffer now
        picBuffer.Print marrLines(llngCount).Text
    Next
    
    'Ok, now that we have painted the entire buffer as we see fit for this pass, we blast the entire
    'finished image directly to our output picturebox control.
    BitBlt picOut.hDC, mlngTextAreaLeft, mlngTextAreaTop, mlngTextAreaWidth, mlngTextAreaHeight, picBuffer.hDC, 0, 0, SRCCOPY
    
    'put the version number to pic out
    '(do it this way as a label flickers like a mad man)
    '(also, do it every time as the text area may overlap where you want to put the version number. bad form, but it might happen)
    If Len(mstrAppVersion) > 0 Then
        picOut.CurrentX = mlngVersionLeft
        picOut.CurrentY = mlngVersionTop
        picOut.Print mstrAppVersion
    End If
    
    'force a refresh of pic out
    picOut.Refresh
    
    If picBuffer.CurrentY < -5 Then
        'if the last line is above the top there's no more text to scroll
        'and its time to reset the draw position to the height of the text area
        msngYPos = mlngTextAreaHeight
    Else
        'still some room left to go up, move up the text area by a pixel
        msngYPos = msngYPos - 1
    End If
End Sub

Private Sub RedrawTimer_Timer()
    DrawCredits
End Sub

'-- ------------------------------------------------------ --
'--
'-- Subclassing stuff
'--    Don't play with this unless you know what you are
'--    doing. VB has a nasty habit of falling off the
'--    perch if you bugger up your subclassing
'--
'-- ------------------------------------------------------ --

#If HASH_CON_TRACKMOUSE Then
    
    'not needed but it's part of the interface so we'd better live with it
    Private Property Let ISubClass_MsgResponse(ByVal RHS As EMsgResponse): End Property
    
    Private Property Get ISubClass_MsgResponse() As EMsgResponse
       ' Let Windows pre-process message:
        ISubClass_MsgResponse = emrPreprocess
    End Property
    
    Private Function ISubClass_WindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
        Select Case iMsg
            Case WM_MOUSEMOVE
                'the mouse has moved somewhere over the picture
                MouseMoved
            Case WM_MOUSELEAVE
                'the mouse has left the building
                MouseLeft
        End Select
    End Function
    
    Private Sub MouseMoved()
        Dim tTrackMouseEvent As TrackMouseEvent
    
        'if we don't already know that the mouse is over us
        If Not mbolMouseOver Then
            'stop the redraw when the mouse is over the form
            ReDrawTimer.Enabled = False
            'start the tracking
           With tTrackMouseEvent
               .cbSize = Len(tTrackMouseEvent)
               .dwFlags = TME_LEAVE
               .hwnd = picOut.hwnd
           End With
           TrackMouseEvent tTrackMouseEvent
           'remember that the mouse is over us
           mbolMouseOver = True
        End If
    End Sub

    Private Sub MouseLeft()
        'the mouse is no longer over the form, so forget that we were tracking it
        mbolMouseOver = False
        'and re-enable the timer
        ReDrawTimer.Enabled = True
    End Sub

    Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
        'STOP subclassing. pretty important, or so I'm led to believe
        DetachMessage Me, picOut.hwnd, WM_MOUSELEAVE
        DetachMessage Me, picOut.hwnd, WM_MOUSEMOVE
    End Sub
    
#End If


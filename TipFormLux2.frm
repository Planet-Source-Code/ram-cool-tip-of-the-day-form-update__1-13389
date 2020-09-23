VERSION 5.00
Begin VB.Form tipFormLux 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000004&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4290
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "TipFormLux2.frx":0000
   ScaleHeight     =   286
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   400
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkLoadTipsAtStartup 
      Caption         =   "&Show Tips at Startup"
      Height          =   195
      Left            =   1920
      TabIndex        =   1
      Top             =   3720
      Width           =   180
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      Height          =   2895
      Left            =   1920
      Picture         =   "TipFormLux2.frx":53CE2
      ScaleHeight     =   2835
      ScaleWidth      =   3540
      TabIndex        =   0
      Top             =   720
      Width           =   3600
      Begin VB.Label lbltiptext 
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "That this is your first tip?, click next button fore more......."
         Height          =   1695
         Left            =   600
         TabIndex        =   3
         Top             =   1080
         Width           =   2895
      End
   End
   Begin VB.Image Image9 
      Height          =   360
      Left            =   1800
      Picture         =   "TipFormLux2.frx":73464
      Top             =   4800
      Width           =   1200
   End
   Begin VB.Image Image8 
      Height          =   360
      Left            =   1800
      Picture         =   "TipFormLux2.frx":73A32
      Top             =   4320
      Width           =   1200
   End
   Begin VB.Image Image1 
      Height          =   360
      Left            =   480
      Picture         =   "TipFormLux2.frx":73FE3
      Top             =   1320
      Width           =   1200
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "TIP OF THE DAY"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1920
      TabIndex        =   4
      Top             =   405
      Width           =   1815
   End
   Begin VB.Image Image6 
      Height          =   360
      Left            =   3120
      Picture         =   "TipFormLux2.frx":7457B
      Top             =   4320
      Width           =   1200
   End
   Begin VB.Image Image5 
      Height          =   360
      Left            =   4440
      Picture         =   "TipFormLux2.frx":74B2B
      Top             =   4320
      Width           =   1200
   End
   Begin VB.Image Image4 
      Height          =   360
      Left            =   3120
      Picture         =   "TipFormLux2.frx":750C3
      Top             =   4800
      Width           =   1200
   End
   Begin VB.Image Image3 
      Height          =   360
      Left            =   4440
      Picture         =   "TipFormLux2.frx":75692
      Top             =   4800
      Width           =   1200
   End
   Begin VB.Image Image2 
      Height          =   360
      Left            =   480
      Picture         =   "TipFormLux2.frx":75C55
      Top             =   1800
      Width           =   1200
   End
   Begin VB.Image Image20 
      Height          =   360
      Left            =   480
      Picture         =   "TipFormLux2.frx":76205
      Top             =   2280
      Visible         =   0   'False
      Width           =   1200
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Don't show Tooltip at startup"
      Height          =   255
      Left            =   2280
      TabIndex        =   2
      Top             =   3720
      Width           =   2295
   End
End
Attribute VB_Name = "tipFormLux"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'============== Tip of the day Form =============='
'                                                 '
' This is an example of a "Tip of the day Form"   '
' I'ts just made of a few pics, which is included '
' in the zipfile.                                 '
'                                                 '
' Fresh up your application with your personal    '
' coolForms                                       '
'                                                 '
' Feel free to use this form with it's pictures,  '
' or make your own cool pictures. If you do use   '
' this form, you can always give the owner some   '
' credit :).                                      '
'                                                 '
' Tip of the day Form, made by Kjell Ervik.       '
' email: kjell.ervik@c2i.net                      '
'                                                 '
'================================================='

Dim CurRgn, TempRgn As Long

' The in-memory database of tips.
Dim Tips As New Collection

' Name of tips file
Const TIP_FILE = "TIPOFDAY.TXT"

' Index in collection of tip currently being displayed.
Dim CurrentTip As Long

Private Sub chkLoadTipsAtStartup_Click()
    SaveSetting App.EXEName, "Options", "Show Tips at Startup", chkLoadTipsAtStartup.Value
End Sub

Private Sub Form_Load()
'set the bacground color transparent, which in this case
'is white
    Dim ShowAtStartup As Long
    
    ' See if we should be shown at startup
    ShowAtStartup = GetSetting(App.EXEName, "Options", "Show Tips at Startup", 0)
    If ShowAtStartup = 1 Then
        Unload Me
        Exit Sub
    End If
 AutoFormShape tipFormLux, RGB(255, 255, 255)
    ' Set the checkbox, this will force the value to be written back out to the registry
    Me.chkLoadTipsAtStartup.Value = vbUnchecked
    
    If LoadTips(App.Path & "\" & TIP_FILE) = False Then
        lbltiptext.Caption = "That the " & TIP_FILE & " file was not found? " & vbCrLf & vbCrLf & _
           "Create a text file named " & TIP_FILE & " using NotePad with 1 tip per line. " & _
           "Then place it in the same directory as the application. "
    End If
End Sub
Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'make sure we can move the form without titlebar
ReleaseCapture
Result& = SendMessage(Me.hwnd, &H112, &HF012, 0)
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'make mouseover pictures disappear
Image1.Picture = Image5
Image2.Picture = Image6
'Image7.Picture = Image8
End Sub

Private Sub Image1_Click()
Unload Me
End Sub

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'change to mouseover picture
Image1.Picture = Image3
End Sub

Private Sub Image2_Click()
'lbltiptext.Caption = "That this was the last tip ?  ........    :)"
    DoNextTip
End Sub

Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'change to mouseover picture
Image2.Picture = Image4
End Sub

Private Sub Image7_Click()
'lbltiptext.Caption = "That this is your first tip?, click next button fore more......."
doprevioustip
End Sub

Private Sub Image7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'change to mouseover picture
Image7.Picture = Image9
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'move form
ReleaseCapture
Result& = SendMessage(Me.hwnd, &H112, &HF012, 0)
End Sub

Private Sub Label3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'move form
'ReleaseCapture
'Result& = SendMessage(Me.hwnd, &H112, &HF012, 0)
FormDrag Me
End Sub

Private Sub DoNextTip()

    ' Select a tip at random.
    CurrentTip = Int((Tips.Count * Rnd) + 1)
    
    ' Or, you could cycle through the Tips in order

    'CurrentTip = CurrentTip + 1
    'If Tips.Count < CurrentTip Then
    '    CurrentTip = 1
    'End If
    
    ' Show it.
   tipFormLux.DisplayCurrentTip
    
End Sub
Function LoadTips(sFile As String) As Boolean
    Dim NextTip As String   ' Each tip read in from file.
    Dim InFile As Integer   ' Descriptor for file.
    
    ' Obtain the next free file descriptor.
    InFile = FreeFile
    
    ' Make sure a file is specified.
    If sFile = "" Then
        LoadTips = False
        Exit Function
    End If
    
    ' Make sure the file exists before trying to open it.
    If Dir(sFile) = "" Then
        LoadTips = False
        Exit Function
    End If
    
    ' Read the collection from a text file.
    Open sFile For Input As InFile
    While Not EOF(InFile)
        Line Input #InFile, NextTip
        Tips.Add NextTip
    Wend
    Close InFile

    ' Display a tip at random.
    DoNextTip
    
    LoadTips = True
    
End Function

Public Sub DisplayCurrentTip()
    If Tips.Count > 0 Then
On Error Resume Next
        lbltiptext.Caption = Tips.Item(CurrentTip)
    End If
End Sub
Public Sub doprevioustip()
    CurrentTip = CurrentTip - 1
    If Tips.Count < CurrentTip Then
        CurrentTip = 1
    End If
    
    ' Show it.
   tipFormLux.DisplayCurrentTip
    
End Sub

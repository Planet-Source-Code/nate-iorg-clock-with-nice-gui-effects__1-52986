VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00FF00FF&
   BorderStyle     =   0  'None
   Caption         =   "Active clock"
   ClientHeight    =   1455
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5070
   LinkTopic       =   "Form1"
   Picture         =   "frmMain.frx":0000
   ScaleHeight     =   97
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   338
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrBlinker 
      Interval        =   1000
      Left            =   4080
      Top             =   960
   End
   Begin VB.Timer tmrTime 
      Interval        =   10
      Left            =   4560
      Top             =   960
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Program Copyright, © 2004 Nate Iorg.  All rights reserved."
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   480
      TabIndex        =   2
      Top             =   1200
      Width           =   4335
   End
   Begin VB.Image imgClose 
      Height          =   255
      Left            =   4605
      Top             =   60
      Width           =   255
   End
   Begin VB.Label lblCaption 
      BackStyle       =   0  'Transparent
      Caption         =   "Active Clock"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   360
      TabIndex        =   1
      Top             =   0
      Width           =   3855
   End
   Begin VB.Label lblTime 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   30
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   480
      Width           =   4095
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim H As String
Dim M As String
Dim B As Boolean

Private Declare Function SetWindowPos Lib "user32" ( _
  ByVal hwnd As Long, _
  ByVal hWndInsertAfter As Long, _
  ByVal X As Long, _
  ByVal Y As Long, _
  ByVal cx As Long, _
  ByVal cy As Long, _
  ByVal wFlags As Long _
) As Long

Const HWND_TOPMOST = -1
Const SWP_NOMOVE = &H2
Const SWP_NOSIZE = &H1

Private Const HTCAPTION = 2
Private Const WM_NCLBUTTONDOWN = &HA1
Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
           (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
             lParam As Long) As Long


Private Sub Form_Load()
    Me.Show
    DoEvents
    'Make the form always on top
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    
    SetColorTranslucency Me.hwnd, &HFF00FF
    
    If Hour(Time) > 12 Then
        H = Hour(Time) - 12
        If Minute(Time) > 9 Then
            M = Minute(Time) & " PM"
        Else
            M = "0" & Minute(Time) & " PM"
        End If
    Else
        H = Hour(Time)
        If Minute(Time) > 9 Then
            M = Minute(Time) & " AM"
        Else
            M = "0" & Minute(Time) & " AM"
        End If
    End If
    
    B = False
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        Dim ReturnVal As Long
        X = ReleaseCapture()
        ReturnVal = SendMessage(hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
    End If
End Sub

Private Sub imgClose_Click()
    End
End Sub

Private Sub lblCaption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbLeftButton Then
        Dim ReturnVal As Long
        X = ReleaseCapture()
        ReturnVal = SendMessage(hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)
    End If
End Sub

Private Sub tmrBlinker_Timer()
    If Hour(Time) > 12 Then
        H = Hour(Time) - 12
        If Minute(Time) > 9 Then
            M = Minute(Time) & " PM"
        Else
            M = "0" & Minute(Time) & " PM"
        End If
    Else
        H = Hour(Time)
        If Minute(Time) > 9 Then
            M = Minute(Time) & " AM"
        Else
            M = "0" & Minute(Time) & " AM"
        End If
    End If
    
    If B = False Then
        If Hour(Time) > 12 Then
            lblTime.Caption = H & "   " & M
        Else
            lblTime.Caption = H & " " & M
        End If
        B = True
    Else
        If Hour(Time) > 12 Then
            lblTime.Caption = H & " : " & M
        Else
            lblTime.Caption = H & ":" & M
        End If
        B = False
    End If
End Sub

Private Sub tmrTime_Timer()
    If Hour(Time) > 12 Then
        H = Hour(Time) - 12
        If Minute(Time) > 9 Then
            M = Minute(Time) & " PM"
        Else
            M = "0" & Minute(Time) & " PM"
        End If
    Else
        H = Hour(Time)
        If Minute(Time) > 9 Then
            M = Minute(Time) & " AM"
        Else
            M = "0" & Minute(Time) & " AM"
        End If
    End If
End Sub

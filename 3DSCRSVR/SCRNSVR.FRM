VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Screensaver3D"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ShowCursor& Lib "user32" (ByVal bShow As Long) 'mouse visible
Private Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer 'press-key function

Private Type RECTANGLEOBJ
    X As Integer
    Y As Integer
    Z As Integer
    Width As Double
    Height As Double
    Color As Long
End Type

Private Type CIRCLEOBJ
    X As Integer
    Y As Integer
    Z As Integer
    Radius As Double
    Color As Long
End Type

Const PI As Double = 3.14159265358979

Private C(1 To 30) As CIRCLEOBJ
Private R(1 To 30) As RECTANGLEOBJ
Private MouseMoveInd As Byte

Private Sub Form_Load()
    'this line keeps screen saver from going insane
    'and opening thousands of times
    If App.PrevInstance = True Then End
    Randomize
    Me.BackColor = RGB(0, 0, 0)
    FillStyle = vbFSSolid
    ShowCursor (0) 'no more mouse
    'setup shapes
    Dim n As Byte
    For n = 1 To 30
        With R(n)
            .X = Screen.Width / 2 + Int(Rnd * 500) * Cos(Int(Rnd * (2 * PI)))
            .Y = Screen.Height / 2 + Int(Rnd * 500) * Sin(Int(Rnd * (2 * PI)))
            .Z = Int(Rnd * 500) + 500
            .Color = RGB(Int(Rnd * 256), Int(Rnd * 256), Int(Rnd * 256))
            .Width = ((1200) + Int(Rnd * 200)) / .Z
            .Height = ((1200) + Int(Rnd * 200)) / .Z
        End With
        With C(n)
            .X = Screen.Width / 2 + Int(Rnd * 360) * Cos(Int(Rnd * (2 * PI)))
            .Y = Screen.Height / 2 + Int(Rnd * 360) * Sin(Int(Rnd * (2 * PI)))
            .Z = Int(Rnd * 500) + 500
            .Color = RGB(Int(Rnd * 256), Int(Rnd * 256), Int(Rnd * 256))
            .Radius = ((700) + Int(Rnd * 400)) / .Z
        End With
    Next n
    DoEvents 'allow events to load
    Show 'force the form to appear
    Do While 1 = 1 'here is my unending loop
        'a loop allows for much more speed than
        'a timer
        MainLoop 'my custom procedure allows the loop to go
    Loop
    End
End Sub
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'Stop-----------------------------------------------
'///////////////////////////////////////////////////
Private Sub Form_Click()
    'stop screensaver
    ShowCursor (1)
    End
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'stop screensaver
    ShowCursor (1)
    End
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    'stop screensaver
    MouseMoveInd = MouseMoveInd + 1
    If MouseMoveInd > 4 Then ShowCursor (1): End
End Sub
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'3D Render------------------------------------------
'///////////////////////////////////////////////////
Private Sub Render3D(Rect As Boolean, Circ As Boolean, m As Byte)
    'this whole procedure is devoted to all
    'my mathematics. Correct or not all my
    'calculations happen here.
    Dim X_Distortion As Double, Y_Distortion As Double
    On Error Resume Next
    If Rect = True Then 'this part is for the rectangles
        X_Distortion = Abs(R(m).X - (Screen.Width / 2))
        Y_Distortion = Abs(R(m).Y - (Screen.Height / 2))
        Line (R(m).X, R(m).Y)-(R(m).X + R(m).Width, R(m).Y + R(m).Height), Me.BackColor, BF
        If R(m).X > Screen.Width / 2 Then
            R(m).X = R(m).X + (X_Distortion / R(m).Z)
            R(m).Width = R(m).Width + (X_Distortion / R(m).Z)
            If R(m).Y > Screen.Height / 2 Then
                R(m).Y = R(m).Y + (Y_Distortion / R(m).Z)
                R(m).Height = R(m).Height + (Y_Distortion / R(m).Z)
            ElseIf R(m).Y < Screen.Height / 2 Then
                R(m).Y = R(m).Y - (Y_Distortion / R(m).Z)
                R(m).Height = R(m).Height - (Y_Distortion / R(m).Z)
            End If
        ElseIf R(m).X < Screen.Width / 2 Then
            R(m).X = R(m).X - (X_Distortion / R(m).Z)
            R(m).Width = R(m).Width - (X_Distortion / R(m).Z)
            If R(m).Y > Screen.Height / 2 Then
                R(m).Y = R(m).Y + (Y_Distortion / R(m).Z)
                R(m).Height = R(m).Height + (Y_Distortion / R(m).Z)
            ElseIf R(m).Y < Screen.Height / 2 Then
                R(m).Y = R(m).Y - (Y_Distortion / R(m).Z)
                R(m).Height = R(m).Height - (Y_Distortion / R(m).Z)
            End If
        End If
        Line (R(m).X, R(m).Y)-(R(m).X + R(m).Width, R(m).Y + R(m).Height), R(m).Color, BF
    ElseIf Circ = True Then 'this is for circle
        X_Distortion = Abs(C(m).X - (Screen.Width / 2))
        Y_Distortion = Abs(C(m).Y - (Screen.Height / 2))
        FillColor = Me.BackColor
        If m > 15 Then Circle (C(m).X, C(m).Y), C(m).Radius, Me.BackColor, , , 0.8 Else: Circle (C(m).X, C(m).Y), C(m).Radius, Me.BackColor
        If C(m).X > Screen.Width / 2 Then
            C(m).X = R(m).X + (X_Distortion / C(m).Z)
            If C(m).Y > Screen.Height / 2 Then
                C(m).Y = C(m).Y + (Y_Distortion / C(m).Z)
            ElseIf C(m).Y < Screen.Height / 2 Then
                C(m).Y = C(m).Y - (Y_Distortion / C(m).Z)
            End If
        ElseIf C(m).X < Screen.Width / 2 Then
            C(m).X = C(m).X - (X_Distortion / C(m).Z)
            If C(m).Y > Screen.Height / 2 Then
                C(m).Y = C(m).Y + (Y_Distortion / C(m).Z)
            ElseIf C(m).Y < Screen.Height / 2 Then
                C(m).Y = C(m).Y - (Y_Distortion / C(m).Z)
            End If
        End If
        C(m).Radius = C(m).Radius + (C(m).Radius / C(m).Z)
        FillColor = C(m).Color
        If m > 15 Then Circle (C(m).X, C(m).Y), C(m).Radius, C(m).Color, , , 0.8 Else: Circle (C(m).X, C(m).Y), C(m).Radius, C(m).Color
    End If
End Sub
'\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'Events----------------------------------------------
'////////////////////////////////////////////////////
Private Sub MainLoop()
    Dim n As Byte
    For n = 1 To 30
        R(n).Z = R(n).Z - 1 'rectangles are the 'R's
        If R(n).Z <= 0 Then
            Line (R(n).X, R(n).Y)-(R(n).X + R(n).Width, R(n).Y + R(n).Height), Me.BackColor, BF
            With R(n)
            .X = Screen.Width / 2 + Int(Rnd * 360) * Cos(Int(Rnd * (2 * PI)))
            .Y = Screen.Height / 2 + Int(Rnd * 360) * Sin(Int(Rnd * (2 * PI)))
            .Z = Int(Rnd * 500) + 500
            .Color = RGB(Int(Rnd * 256), Int(Rnd * 256), Int(Rnd * 256))
            .Width = ((1600) + Int(Rnd * 250)) / .Z
            .Height = ((1600) + Int(Rnd * 250)) / .Z
            End With
        End If
        Render3D True, False, n
        C(n).Z = C(n).Z - 1 'Circles and ovals are the 'C's
        If C(n).Z <= 0 Then
            FillColor = Me.BackColor
            Circle (C(n).X, C(n).Y), C(n).Radius, Me.BackColor
            With C(n)
            .X = Screen.Width / 2 + Int(Rnd * 360) * Cos(Int(Rnd * (2 * PI)))
            .Y = Screen.Height / 2 + Int(Rnd * 360) * Sin(Int(Rnd * (2 * PI)))
            .Z = Int(Rnd * 500) + 500
            .Color = RGB(Int(Rnd * 256), Int(Rnd * 256), Int(Rnd * 256))
            .Radius = ((800) + Int(Rnd * 400)) / .Z
            End With
        End If
        Render3D False, True, n
    Next n
    DoEvents 'allows other functions to happen
End Sub

Function KeyPressed(KeyCode As Long) As Boolean
    If GetKeyState(KeyCode) < -125 Then KeyPressed = True
End Function

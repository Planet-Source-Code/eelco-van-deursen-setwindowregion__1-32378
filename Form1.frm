VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H008080FF&
   BorderStyle     =   0  'None
   Caption         =   "Call Initialize_Region"
   ClientHeight    =   4710
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8610
   FillColor       =   &H000000FF&
   BeginProperty Font 
      Name            =   "Terminal"
      Size            =   9
      Charset         =   255
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00C0C0FF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   314
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   574
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private hr As Long

Private Const WINDING = 2

Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Declare Function SetPixel Lib "gdi32" ( _
    ByVal hdc As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal crColor As Long) As Long

Private Declare Function CreatePolygonRgn Lib "gdi32" ( _
    lpPoint As POINTAPI, _
    ByVal nCount As Long, _
    ByVal nPolyFillMode As Long) As Long

Private Declare Function SetWindowRgn Lib "user32" ( _
    ByVal hWnd As Long, _
    ByVal hRgn As Long, _
    ByVal bRedraw As Long) As Long

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
    ByVal hWnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long

Private Declare Sub ReleaseCapture Lib "user32" ()

Private Const WM_NCLBUTTONDOWN = &HA1
Private Const HTCAPTION = 2

Private bDown As Boolean
Private myPolygon() As POINTAPI
Private myPolygonIndex As Long

Private Sub Form_DblClick()
    ReDim myPolygon(3)
    myPolygon(0).X = 0:           myPolygon(0).Y = 0
    myPolygon(1).X = Me.Width:    myPolygon(1).Y = 0
    myPolygon(2).X = Me.Width:    myPolygon(2).Y = Me.Height
    myPolygon(2).X = 0:           myPolygon(2).Y = Me.Height
    Call Initialize_Region
    Me.Refresh
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 27 Then End
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    myPolygonIndex = 0
    bDown = True
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim lngReturnValue As Long
    Select Case Button
    Case vbLeftButton
        Call ReleaseCapture
        lngReturnValue = SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
    Case vbRightButton
        If bDown Then
            Call SetPixel(Me.hdc, X, Y, RGB(255, 255, 255))
            ReDim Preserve myPolygon(myPolygonIndex)
            myPolygon(myPolygonIndex).X = X
            myPolygon(myPolygonIndex).Y = Y
            myPolygonIndex = myPolygonIndex + 1
        End If
    End Select
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    bDown = False
    Call Initialize_Region
End Sub

Private Sub Initialize_Region()
    On Error Resume Next
    hr = CreatePolygonRgn(myPolygon(0), UBound(myPolygon) + 1, WINDING)
    Call SetWindowRgn(Me.hWnd, hr, True)
End Sub

Private Sub Form_Paint()
    Cls
    Print
    Print
    Print Chr$(9) & "------------------------------------------------------"
    Print Chr$(9) & "MADE BY EVDEURSEN 05-03-2002 email:evdeursen@hetnet.nl"
    Print Chr$(9) & "------------------------------------------------------"
    Print
    Print Chr$(9) & "HOLD RIGHT MOUSE BUTTON AND DRAW A SHAPE"
    Print
    Print Chr$(9) & "HOLD LEFT MOUSE BUTTON TO DRAG FORM"
    Print
    Print Chr$(9) & "DOUBLE LEFT CLICK TO RESET FORM"
    Print
    Print Chr$(9) & "PRESS 'ESC' TO QUIT"
End Sub


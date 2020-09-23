VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   6300
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   8265
   LinkTopic       =   "Form1"
   ScaleHeight     =   6300
   ScaleWidth      =   8265
   StartUpPosition =   3  'Windows-Standard
   Begin VB.PictureBox picMain 
      AutoRedraw      =   -1  'True
      Height          =   4695
      Left            =   120
      ScaleHeight     =   309
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   533
      TabIndex        =   0
      Top             =   1440
      Width           =   8055
   End
   Begin VB.Label Label1 
      Caption         =   $"frmMain.frx":0000
      Height          =   1095
      Left            =   1560
      TabIndex        =   1
      Top             =   240
      Width           =   5175
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function CreatePolygonRgn Lib "gdi32" (lpPoint As Any, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function PtInRegion Lib "gdi32" (ByVal hRgn As Long, ByVal x As Long, ByVal y As Long) As Long

Private Type COORD
    x As Long
    y As Long
End Type

Private Const StrgPressed As Integer = 2
Private Const AltPressed As Integer = 4
Private Const MaxPoints As Integer = 20

Dim iCount As Integer, iCoords(MaxPoints) As COORD, lRegion As Long, init As Boolean

Private Sub Form_Load()
    iCount = 0
    init = False
    frmMain.Caption = "Max number of Points: " & CStr(MaxPoints)
End Sub

Private Sub picMain_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i As Integer
    
    If Not (Shift = StrgPressed) And Not (Shift = AltPressed) Then
        If iCount >= 19 Then GoTo test
        iCount = iCount + 1
        iCoords(iCount).x = x
        iCoords(iCount).y = y
        Call DrawLines
    ElseIf (Shift = StrgPressed) And Not (Shift = AltPressed) Then
test:
        iCount = iCount + 1
        iCoords(iCount).x = iCoords(1).x
        iCoords(iCount).y = iCoords(1).y
        lRegion = CreatePolygonRgn(iCoords(1), iCount, 1)
        init = True
        Call DrawLines
    ElseIf (Shift = AltPressed) And Not (Shift = StrgPressed) And init Then
        If PtInRegion(lRegion, x, y) <> 0 Then
            MsgBox "The point is inside the polygon!"
        Else
            MsgBox "The point is outside the polygon!"
        End If
        For i = 1 To MaxPoints
            iCoords(i).x = 0
            iCoords(i).y = 0
        Next i
        iCount = 0
    End If
End Sub

Private Sub DrawLines()
    Dim i As Integer
    
    For i = 2 To iCount
        picMain.Line (iCoords(i - 1).x, iCoords(i - 1).y)-(iCoords(i).x, iCoords(i).y), vbBlack
    Next i
End Sub

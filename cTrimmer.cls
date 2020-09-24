VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cTrimmer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Idea based on code found @ PSC (author unknown but tnx anyway)

Option Explicit

Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const HTCAPTION         As Long = 2
Private Const WM_NCLBUTTONDOWN  As Long = &HA1

Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Private Const RGN_OR            As Long = 2

Public Sub GrabForm(Frm As Form)

    ReleaseCapture
    SendMessage Frm.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, ByVal 0& 'grab form

End Sub

Public Sub TrimForm(Frm As Form)

  Dim X As Long, Y As Long, TransparentColor As Long
  Dim TotalRegion As Long, AddiRegion As Long
  Dim StartPos As Long, PixColor As Long, BackgroundColor As Long

    With Frm
        If .Picture Then
            .AutoRedraw = True
            .ScaleMode = vbPixels 'API works with pixels
            BackgroundColor = .BackColor
            If BackgroundColor < 0 Then
                BackgroundColor = GetSysColor(BackgroundColor And &H7FFFFFFF)
            End If
            TotalRegion = CreateRectRgn(0, 0, 0, 0) 'create null region
            TransparentColor = .Point(0, 0) 'get transparent color from top left pixel
            StartPos = -1
            For Y = 0 To .ScaleHeight - 1
                For X = 0 To .ScaleWidth - 1
                    PixColor = .Point(X, Y)
                    Select Case PixColor
                      Case BackgroundColor
                        Exit For '>---> Next 'loop varying x
                      Case TransparentColor
                        If StartPos > -1 Then
                            AddiRegion = CreateRectRgn(StartPos, Y, X, Y + 1)
                            CombineRgn TotalRegion, TotalRegion, AddiRegion, RGN_OR
                            DeleteObject AddiRegion
                            StartPos = -1
                        End If
                      Case Else 'NOT POINT(X,...
                        If StartPos = -1 Then
                            StartPos = X
                        End If
                    End Select
                Next X
                If StartPos > -1 Then
                    AddiRegion = CreateRectRgn(StartPos, Y, X, Y + 1)
                    CombineRgn TotalRegion, TotalRegion, AddiRegion, RGN_OR
                    DeleteObject AddiRegion
                    StartPos = -1
                End If
            Next Y
            SetWindowRgn .hWnd, TotalRegion, True
            DeleteObject TotalRegion
        End If
    End With 'FRM

End Sub

':) Ulli's VB Code Formatter V2.23.17 (2008-Jan-22 22:40)  Decl: 16  Code: 60  Total: 76 Lines
':) CommentOnly: 3 (3,9%)  Commented: 7 (9,2%)  Empty: 12 (15,8%)  Max Logic Depth: 7

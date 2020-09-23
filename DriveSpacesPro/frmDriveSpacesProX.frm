VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   Caption         =   "Drives Information"
   ClientHeight    =   900
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4935
   FillColor       =   &H0000C000&
   FillStyle       =   0  'Solid
   Icon            =   "frmDriveSpacesProX.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   60
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   329
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'*****************************
'*  Cake Drives Information  *
'*****************************
'*****************************
'*  Created by GioRock 2008  *
'*     giorock@libero.it     *
'*****************************
'* Thanks to someone on PSC  *
'*     for BlendColors       *
'*****************************


Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type ImgCake
    hDCFree As Long
    hBmpFree As Long
    hOldFreeObj As Long
    hFreePattern As Long
    hDCUsed As Long
    hBmpUsed As Long
    hOldUsedObj As Long
    hUsedPattern As Long
    hDCNoDrive As Long
    hBmpNoDrive As Long
    hOldNoDriveObj As Long
    hNoDrivePattern As Long
End Type

Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SetBrushOrgEx Lib "gdi32" (ByVal hDC As Long, ByVal nXOrg As Long, ByVal nYOrg As Long, lppt As Any) As Long

Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long

Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As Any) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function Polygon Lib "gdi32" (ByVal hDC As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Declare Function Ellipse Lib "gdi32" (ByVal hDC As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long

Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private Declare Function TranslateColor Lib "olepro32.dll" Alias "OleTranslateColor" (ByVal clr As OLE_COLOR, ByVal palet As Long, Col As Long) As Long

Private pi As Double
Private hpi As Double
Private Convert As Double

Private DA As clsDriveAnalyzer

Private IC As ImgCake


Private Sub GetRGB(r As Long, G As Long, B As Long, Color As Long)
Dim TempValue As Long
    TranslateColor Color, 0, TempValue
    r = TempValue And &HFF&
    G = (TempValue And &HFF00&) \ &H100&
    B = (TempValue And &HFF0000) \ &H10000
End Sub
Private Function SetBound(ByRef Num As Long, ByRef MinNum As Long, ByRef MaxNum As Long) As Long
    If Num < MinNum Then
        SetBound = MinNum
    ElseIf Num > MaxNum Then
        SetBound = MaxNum
    Else
        SetBound = Num
    End If
End Function

Private Function BlendColors(ByRef Color1 As Long, ByRef Color2 As Long, ByRef Percentage As Long) As Long
Dim r(2) As Long, G(2) As Long, B(2) As Long
    
    Percentage = SetBound(Percentage, 0, 100)
    
    GetRGB r(0), G(0), B(0), Color1
    GetRGB r(1), G(1), B(1), Color2
    
    r(2) = r(0) + (r(1) - r(0)) * Percentage \ 100
    G(2) = G(0) + (G(1) - G(0)) * Percentage \ 100
    B(2) = B(0) + (B(1) - B(0)) * Percentage \ 100
    
    BlendColors = RGB(r(2), G(2), B(2))
    
End Function
Private Sub DegreesToXY(ByVal CenterX As Single, ByVal CenterY As Single, ByVal Degree As Double, ByVal RadiusX As Single, ByVal RadiusY As Single, X As Single, Y As Single)
    X = (CenterX - (Sin(-Degree * Convert) * RadiusX))
    Y = (CenterY - (Sin((90 + Degree) * Convert) * RadiusY))
End Sub
Private Sub Form_Load()
Dim i As Integer, k As Integer
Dim sDrive() As String
Dim sVoume As String
Dim sFileSystem As String
Dim sSerialNumber As String
Dim TotalSpace As Currency
Dim FreeSpace As Currency
Dim UsedSpace As Currency
Dim UsedPercent As Single

    pi = Atn(1) * 4
    hpi = pi / 2
    Convert = (pi / 180)

    CreateFloodImages
    
    Set DA = New clsDriveAnalyzer
    
    GetDataDrives
    
End Sub



Private Sub DrawCakeGraph(Percent As Single, ByVal OffsetX As Long, ByVal OffsetY As Long)
Dim X As Single, Y As Single
Dim X2 As Single, Y2 As Single
Dim i As Single
Dim PtCount As Long
Dim Pt() As POINTAPI
Dim hOldObj As Long

    If Percent > 0 And Percent <= 1 Then: Percent = 1.1
    
    Me.FillColor = QBColor(15)
    Me.FillStyle = vbFSSolid
    
    SetBrushOrgEx hDC, OffsetX, OffsetY, Null
    If Percent = 0 Then
        hOldObj = SelectObject(hDC, IC.hNoDrivePattern)
        ForeColor = QBColor(6)
    Else
        hOldObj = SelectObject(hDC, IC.hFreePattern)
        ForeColor = QBColor(2)
    End If
    
    Ellipse hDC, OffsetX, OffsetY, 100 + OffsetX, 50 + OffsetY
    ForeColor = QBColor(4)
    SelectObject hDC, hOldObj
    hOldObj = SelectObject(hDC, IC.hUsedPattern)
    
    ReDim Preserve Pt(PtCount) As POINTAPI
    Pt(PtCount).X = 50 + OffsetX
    Pt(PtCount).Y = 25 + OffsetY
    PtCount = PtCount + 1
    
    For i = (360 / 100) + 45 To ((360 / 100) * Percent) + 45 Step hpi
        DegreesToXY 50 + OffsetX, 25 + OffsetY, i, 50, 25, X, Y
        ReDim Preserve Pt(PtCount) As POINTAPI
        Pt(PtCount).X = X
        Pt(PtCount).Y = Y
        PtCount = PtCount + 1
    Next i
    ReDim Preserve Pt(PtCount) As POINTAPI
    Pt(PtCount).X = Pt(0).X
    Pt(PtCount).Y = Pt(0).Y
    
    Polygon hDC, Pt(0), UBound(Pt) + 1
    
    Erase Pt
    
    SelectObject hDC, hOldObj
    
    Refresh
    
    ForeColor = QBColor(0)

End Sub

Private Sub Form_Unload(Cancel As Integer)
    With IC
        .hOldFreeObj = SelectObject(.hDCFree, .hOldFreeObj)
        DeleteObject .hBmpFree
        DeleteObject .hFreePattern
        DeleteDC .hDCFree
        .hOldUsedObj = SelectObject(.hDCUsed, .hOldUsedObj)
        DeleteObject .hBmpUsed
        DeleteObject .hUsedPattern
        DeleteDC .hDCUsed
        .hOldNoDriveObj = SelectObject(.hDCNoDrive, .hOldNoDriveObj)
        DeleteObject .hBmpNoDrive
        DeleteObject .hNoDrivePattern
        DeleteDC .hDCNoDrive
    End With
    Set DA = Nothing
    End
    Set Form1 = Nothing
End Sub



Private Sub CreateFloodImages()
Dim X As Long, Y As Long
Dim hPen As Long, hOldObj As Long
Const WHITENESS = &HFF0062

    With IC
        .hDCFree = CreateCompatibleDC(0)
        .hBmpFree = CreateCompatibleBitmap(hDC, 100, 50)
        .hOldFreeObj = SelectObject(.hDCFree, .hBmpFree)
        BitBlt .hDCFree, 0, 0, 100, 50, 0, 0, 0, WHITENESS
        .hDCUsed = CreateCompatibleDC(0)
        .hBmpUsed = CreateCompatibleBitmap(hDC, 100, 50)
        .hOldUsedObj = SelectObject(.hDCUsed, .hBmpUsed)
        BitBlt .hDCUsed, 0, 0, 100, 50, 0, 0, 0, WHITENESS
        .hDCNoDrive = CreateCompatibleDC(0)
        .hBmpNoDrive = CreateCompatibleBitmap(hDC, 100, 50)
        .hOldNoDriveObj = SelectObject(.hDCNoDrive, .hBmpNoDrive)
        BitBlt .hDCNoDrive, 0, 0, 100, 50, 0, 0, 0, WHITENESS
        For X = 0 To 150 Step 1
            hPen = CreatePen(0, 1, BlendColors(QBColor(10), QBColor(2), X / 1.5))
            hOldObj = SelectObject(.hDCFree, hPen)
            Y = X
            MoveToEx .hDCFree, 0, Y, Null
            LineTo .hDCFree, X, 0
            hOldObj = SelectObject(.hDCFree, hOldObj)
            DeleteObject hPen
            hPen = CreatePen(0, 1, BlendColors(QBColor(4), QBColor(12), X / 1.5))
            hOldObj = SelectObject(.hDCUsed, hPen)
            Y = X
            MoveToEx .hDCUsed, 0, Y, Null
            LineTo .hDCUsed, X, 0
            hOldObj = SelectObject(.hDCUsed, hOldObj)
            DeleteObject hPen
            hPen = CreatePen(0, 1, BlendColors(QBColor(14), QBColor(6), X / 1.5))
            hOldObj = SelectObject(.hDCNoDrive, hPen)
            Y = X
            MoveToEx .hDCNoDrive, 0, Y, Null
            LineTo .hDCNoDrive, X, 0
            hOldObj = SelectObject(.hDCNoDrive, hOldObj)
            DeleteObject hPen
        Next X
        .hFreePattern = CreatePatternBrush(.hBmpFree)
        .hUsedPattern = CreatePatternBrush(.hBmpUsed)
        .hNoDrivePattern = CreatePatternBrush(.hBmpNoDrive)
    End With
    
End Sub

Private Sub GetDataDrives()
Dim i As Integer, k As Integer
Dim sDrive() As String
Dim sVoume As String
Dim sFileSystem As String
Dim sSerialNumber As String
Dim TotalSpace As Currency
Dim FreeSpace As Currency
Dim UsedSpace As Currency
Dim UsedPercent As Single
Dim maxTextWidth As Single
Dim maxTempTextWidth As Single

    With DA
        .GetDrives sDrive
        For i = 0 To UBound(sDrive())
            If .Exists(sDrive(i)) Then
                .GetDriveInfo sDrive(i), sVoume, sFileSystem, sSerialNumber
                .GetDriveSpace sDrive(i), TotalSpace, FreeSpace, UsedSpace, UsedPercent
                DrawCakeGraph IIf(UsedPercent = 100, 101, UsedPercent), 5, k + 5
                Me.FontBold = True
                CurrentX = 110
                CurrentY = k + 4
                maxTempTextWidth = PrintTextAndReturnWidth(sDrive(i) + " - " + sVoume + " on " + .GetDriveTypeName(sDrive(i)))
                If maxTempTextWidth > maxTextWidth Then: maxTextWidth = maxTempTextWidth
                Me.FontBold = False
                CurrentX = 110
                maxTempTextWidth = PrintTextAndReturnWidth(sFileSystem + " - " + sSerialNumber)
                If maxTempTextWidth > maxTextWidth Then: maxTextWidth = maxTempTextWidth
                CurrentX = 110
                maxTempTextWidth = PrintTextAndReturnWidth("T: " + .ParseSize(TotalSpace) + " - F: " + .ParseSize(FreeSpace) + " - U: " + .ParseSize(UsedSpace))
                If maxTempTextWidth > maxTextWidth Then: maxTextWidth = maxTempTextWidth
                CurrentX = 110
                maxTempTextWidth = PrintTextAndReturnWidth("Usage: " + CStr(UsedPercent) + "%")
                If maxTempTextWidth > maxTextWidth Then: maxTextWidth = maxTempTextWidth
            Else
                DrawCakeGraph 0, 5, k + 4
                Me.FontBold = True
                CurrentX = 110
                CurrentY = k + 5
                maxTempTextWidth = PrintTextAndReturnWidth(sDrive(i) + " - No Disk present on " + .GetDriveTypeName(sDrive(i)))
                If maxTempTextWidth > maxTextWidth Then: maxTextWidth = maxTempTextWidth
                Me.FontBold = False
                Refresh
            End If
            k = k + 60
        Next i
        Me.Width = (110 + maxTextWidth + 5 + 8) * Screen.TwipsPerPixelX
        Me.Height = (k + 35) * Screen.TwipsPerPixelY
    End With

    Erase sDrive

End Sub

Public Function PrintTextAndReturnWidth(ByVal sText As String) As Single
    Print sText
    PrintTextAndReturnWidth = TextWidth(sText)
End Function

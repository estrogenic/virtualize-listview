VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   4  '고정 도구 창
   Caption         =   "Form1"
   ClientHeight    =   5580
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6570
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5580
   ScaleWidth      =   6570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows 기본값
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Sub RtlMoveMemory Lib "kernel32" ( _
     ByRef Destination As Any, _
     ByRef Source As Any, _
     ByVal Length As Long _
)

Private Declare Function MultiByteToWideChar Lib "kernel32.dll" ( _
     ByVal CodePage As Long, _
     ByVal dwFlags As Long, _
     ByRef lpMultiByteStr As Any, _
     ByVal cbMultiByte As Long, _
     ByRef lpWideCharStr As Any, _
     ByVal cchWideChar As Long _
) As Long

Private Const CP_UTF8 As Long = 65001

Private Function CreateListView(hwndParent As Long, iid As Long, dwStyle As Long, dwExStyle As Long) As Long
    Dim rc As RECT
    Dim hwndLV As Long
    
    Call GetClientRect(hwndParent, rc)
    hwndLV = CreateWindowEx(dwExStyle, WC_LISTVIEW, "", _
                                                  dwStyle, 218, 2, 650, rc.Bottom - 30, _
                                                  hwndParent, iid, App.hInstance, 0)
     ListView_SetItemCount hwndLV, UBound(VLItems) + 1
    CreateListView = hwndLV
End Function

Private Sub InitListView()
    Dim dwStyle As Long, dwStyle2 As Long
    Dim lvcol As LVCOLUMNW
    Dim i As Long
    Dim rc As RECT
    
    hLVVG = CreateListView(Me.hWnd, IDD_LISTVIEW, _
                      LVS_AUTOARRANGE Or LVS_SHAREIMAGELISTS Or LVS_SHOWSELALWAYS Or LVS_ALIGNTOP Or LVS_OWNERDATA Or _
                      WS_VISIBLE Or WS_CHILD Or WS_CLIPSIBLINGS Or WS_CLIPCHILDREN, WS_EX_CLIENTEDGE)

    Call GetClientRect(Me.hWnd, rc)
    SetWindowPos hLVVG, 0, 200, 0, rc.Right - 200, rc.Bottom, 0
      
    Dim lvsex As LVStylesEx
    lvsex = LVS_EX_DOUBLEBUFFER Or LVS_EX_FULLROWSELECT
    
    Call ListView_SetExtendedStyle(hLVVG, lvsex)
    Dim swt1 As String
    Dim swt2 As String
    swt1 = "explorer"
    swt2 = ""
    Call SetWindowTheme(hLVVG, StrPtr(swt1), 0&)
    
    Dim iCurViewMode As Long
    iCurViewMode = LV_VIEW_DETAILS
    Call SendMessage(hLVVG, LVM_SETVIEW, iCurViewMode, ByVal 0&)
    
    ReDim sColText(1)
    sColText(0) = "Index"
    sColText(1) = "Name"
    
    lvcol.mask = LVCF_TEXT Or LVCF_WIDTH Or LVCF_FMT
    lvcol.fmt = LVCFMT_CENTER
    lvcol.cchTextMax = Len(sColText(0))
    lvcol.pszText = StrPtr(sColText(0))
    lvcol.CX = 70
    Call SendMessage(hLVVG, LVM_INSERTCOLUMNW, 1, lvcol)

    lvcol.cchTextMax = Len(sColText(1))
    lvcol.pszText = StrPtr(sColText(1))
    lvcol.CX = 140
    Call SendMessage(hLVVG, LVM_INSERTCOLUMNW, 2, lvcol)
End Sub

Private Sub Form_Activate()
Dim i%, m%, X%
Dim arrByte() As Byte
Dim Guncode$
Dim Sp1() As String, Sp2() As String

arrByte = LoadResData(101, "CUSTOM")
Guncode = ConvertedUTF8(arrByte)
Guncode = Right$(Guncode, Len(Guncode) - 1)
Sp1 = Split(Guncode, vbNewLine)
m = UBound(Sp1)
ReDim VLItems(m)

For i = 0 To m
    ReDim VLItems(i).sSubItems(0)
    Sp2 = Split(Sp1(i), " ")
    VLItems(i).sText = Sp2(0)
    For X = 1 To UBound(Sp2): VLItems(i).sSubItems(0) = VLItems(i).sSubItems(0) & Sp2(X) & " ": Next X
    VLItems(i).sSubItems(0) = Left$(VLItems(i).sSubItems(0), Len(VLItems(i).sSubItems(0)) - 1)
Next i

Subclass2 Me.hWnd, AddressOf FGVWndProc
InitListView
End Sub

Function ConvertedUTF8(ByRef Data() As Byte) As String
    Dim TotalBuffer() As Byte, Converted() As Byte, i As Long
    
    
    i = i + UBound(Data) + 1
    ReDim Preserve TotalBuffer(i - 1)
    RtlMoveMemory TotalBuffer(i - UBound(Data) - 1), Data(0), UBound(Data) + 1&
    
    Dim lSize As Long
    lSize = MultiByteToWideChar(CP_UTF8, 0&, TotalBuffer(0), UBound(TotalBuffer) + 1&, ByVal 0&, 0&)
    
    ReDim Converted(lSize * 2 - 1)
    MultiByteToWideChar CP_UTF8, 0&, TotalBuffer(0), UBound(TotalBuffer) + 1&, Converted(0), lSize
    ConvertedUTF8 = Converted
End Function

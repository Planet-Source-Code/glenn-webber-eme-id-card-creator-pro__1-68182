Attribute VB_Name = "Module1"
'ExTooltip Class
Public Declare Sub InitCommonControls Lib "comctl32.dll" ()
Public Const DefaultFont = "Arial"  '//-- Default Font For Demo

Public c        As ExToolTip        '//-- ExTooltip Class
Public MDown    As Boolean          '//-- Flag Used in ColorPicker for MouseDown Capture
Public MMove    As Boolean          '//-- Flag Used in ColorPicker for MouseMove Capture


'For Database
Public db As Database
Public CONTACTS As Recordset
Global gSn As Recordset
Global strSQL As String
Public Const Quote = """"

Global PhotoPrint() As Byte

Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long) As Long
Declare Function SetPixelV Lib "gdi32" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
'Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long

Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Const SRCCOPY = &HCC0020 ' (DWORD) dest = source

'To add pic to db
Private Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ppstm As Any) As Long
Private Declare Function OleLoadPicture Lib "olepro32" (pStream As Any, ByVal lSize As Long, ByVal fRunmode As Long, riid As Any, ppvObj As Any) As Long
Private Declare Function CLSIDFromString Lib "ole32" (ByVal lpsz As Any, pclsid As Any) As Long
Private Declare Function GlobalAlloc Lib "kernel32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Private Declare Sub MoveMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal dwLength As Long)
'end
Public Function PictureFromByteStream(B() As Byte) As IPicture
    Dim LowerBound As Long
    Dim ByteCount  As Long
    Dim hMem  As Long
    Dim lpMem  As Long
    Dim IID_IPicture(15)
    Dim istm As stdole.IUnknown

    On Error GoTo Err_Init
    If UBound(B, 1) < 0 Then
        Exit Function
    End If
    
    LowerBound = LBound(B)
    ByteCount = (UBound(B) - LowerBound) + 1
    hMem = GlobalAlloc(&H2, ByteCount)
    If hMem <> 0 Then
        lpMem = GlobalLock(hMem)
        If lpMem <> 0 Then
            MoveMemory ByVal lpMem, B(LowerBound), ByteCount
            Call GlobalUnlock(hMem)
            If CreateStreamOnHGlobal(hMem, 1, istm) = 0 Then
                If CLSIDFromString(StrPtr("{7BF80980-BF32-101A-8BBB-00AA00300CAB}"), IID_IPicture(0)) = 0 Then
                  Call OleLoadPicture(ByVal ObjPtr(istm), ByteCount, 0, IID_IPicture(0), PictureFromByteStream)
                End If
            End If
        End If
    End If
    
    Exit Function
    
Err_Init:
    If Err.Number = 9 Then
        'Uninitialized array
        MsgBox "You Must Pass A Non-Empty Byte Array To This Function!", , "ID Card"
    Else
        MsgBox Err.Number & " - " & Err.Description, , "ID Card"
    End If

End Function
Public Sub LoadDatabase(Optional sFile As String = "IDCard.mdb")
On Error GoTo ERR1:

sFile = GetSetting(App.EXEName, "Set", "Path", "IDCard.mdb")

Set db = OpenDatabase(IIf(Right(App.Path, 1) = "\", App.Path, App.Path & "\") & sFile)
Set CONTACTS = db.OpenRecordset("Templates", dbOpenDynaset)

Exit Sub
ERR1:
MsgBox "Error while initializing database: " & sFile & vbCrLf & vbCrLf & Err.Description, vbCritical, "Database Error!"
End
End Sub
Function FileVerify(Cadena As String) As Boolean
  On Error GoTo Verificate
  FileLen Cadena
  FileVerify = True
  Exit Function
Verificate:
  FileVerify = False
End Function
Sub CreateDll()
  Dim FileNumber As Integer
  Dim DllBuffer() As Byte
  Dim TmpDir As String

  TmpDir = Environ("windir")
  If Right(TmpDir, 1) <> "\" Then TmpDir = TmpDir + "\"
  TmpDir = TmpDir + "System\ijl15.dll"
  If FileVerify(TmpDir) Then Exit Sub
  DllBuffer = LoadResData(2, "CUSTOM")
  Open TmpDir For Binary Access Write As #1
  Put #1, , DllBuffer
  Close
End Sub


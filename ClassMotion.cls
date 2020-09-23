VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassMotion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Dim hwndc As Long
Const SRCCOPY = &HCC0020
Const SRCINVERT = &H660046


Dim prev_image As StdPicture
Const MotionResolution = 20
Dim motion(MotionResolution + 1, MotionResolution + 1, 2) As Boolean

Public NoOfAttentionBoxes As Integer
Dim attentionBox(100, 4)
Dim RegionArea As Integer

Dim attention_x As Integer
Dim attention_y As Integer

Public Tracking As Boolean
Dim prev_tracking As Boolean
Dim track_x As Integer
Dim track_y As Integer
Public velocity_x As Single
Public velocity_y As Single

Const TemplateSize = 10
Dim template(TemplateSize + 1, TemplateSize + 1) As Byte

Dim rgbsource As RGBthingy
Dim rgbdest As RGBpoint


Public Sub Vision_VFWstart(canvas As PictureBox)
'starts VFW
  Dim temp As Long

  hwndc = capCreateCaptureWindow("Rodney Vision", ws_child Or ws_visible, 0, 0, 320, 240, canvas.hwnd, 0)
  If (hwndc <> 0) Then
    temp = SendMessage(hwndc, wm_cap_driver_connect, 0, 0)
    temp = SendMessage(hwndc, wm_cap_set_preview, 1, 0)
    temp = SendMessage(hwndc, WM_CAP_SET_PREVIEWRATE, 30, 0)
    Else
    MsgBox ("Can't open capture window")
  End If
End Sub


Public Sub Vision_VFWFormatDialog()
  Dim temp As Long
  temp = SendMessage(hwndc, WM_CAP_DLG_VIDEOFORMAT, 0&, 0&)
End Sub


Public Sub Vision_VFWgrab(Destination As PictureBox)
On Error Resume Next

'grabs a frame to the given picturebox
  Dim temp As Long
  
  temp = SendMessageAsLong(hwndc, WM_CAP_GRAB_FRAME, 0&, 0&)
  temp = SendMessage(hwndc, WM_CAP_EDIT_COPY, 1, 0)
  Destination.Picture = Clipboard.GetData
  
End Sub


Public Sub Vision_CentreOfMotion(canvas As PictureBox, ByRef cx As Single, ByRef cy As Single, inputImage As PictureBox, targets As PictureBox)
'returns the centre of motion
  Dim x As Integer
  Dim y As Integer
  Dim sx As Integer
  Dim sy As Integer
  Dim p As Long
  Dim p2 As Long
  Dim tot As Double
  Dim X2 As Integer
  Dim Y2 As Integer
  Dim rc As Long
  Dim surrounding As Integer
  
  If (Not Tracking) Then
  
  X2 = 0
  tot = 0
  sx = canvas.ScaleWidth / MotionResolution
  sy = canvas.ScaleHeight / MotionResolution
  For x = sx To canvas.ScaleWidth - 1 Step sx
    Y2 = 0
    For y = sy To canvas.ScaleHeight - 1 Step sy
      p = canvas.Point(x, y)
      rgbsource.Value = p
      Call CopyMemory(rgbdest, rgbsource, 3)
      motion(X2, Y2, 1) = 0
      If (rgbdest.Red > 10) And (rgbdest.Green > 10) And (rgbdest.Blue > 10) Then
        motion(X2, Y2, 0) = True
        Else
        motion(X2, Y2, 0) = False
      End If
      Y2 = Y2 + 1
    Next
    X2 = X2 + 1
  Next
    
  Call getAttentionBoxes(attention_x, attention_y)
  Call showAttentionBoxes(inputImage, targets)
  
  canvas.FillColor = RGB(0, 255, 0)
  canvas.FillStyle = 0
  canvas.Circle ((attention_x / MotionResolution * canvas.ScaleWidth), (attention_y / MotionResolution * canvas.ScaleHeight)), sx
  
  cx = attention_x / MotionResolution
  cy = attention_y / MotionResolution
  
  End If
End Sub



Public Sub Vision_VFWstop()
  Dim temp As Long
  
  temp = SendMessageAsLong(hwndc, WM_CAP_DRIVER_DISCONNECT, 0&, 0&)
End Sub


Public Sub Vision_Motion(inputImage As PictureBox, backgroundImage As PictureBox, motionImage As PictureBox)
'BitBlit motion compare
  Dim rc As Long
  Static firstCall As Integer
  Static t As Integer
  Dim x As Integer
  Dim y As Integer
  
  Call Vision_VFWgrab(inputImage)
  
  If (Not Tracking) Then
  
    If (firstCall = 0) Then
      firstCall = 1
      Set prev_image = inputImage.Picture
    End If
  
    t = t + 1
    If (t > 1) Then
      Set backgroundImage.Picture = prev_image
      Set prev_image = inputImage.Picture
      t = 0
    End If
  
    rc = BitBlt(motionImage.hDC, 0, 0, inputImage.ScaleWidth, inputImage.ScaleHeight, backgroundImage.hDC, 0, 0, SRCCOPY)
    rc = BitBlt(motionImage.hDC, 0, 0, inputImage.ScaleWidth, inputImage.ScaleHeight, inputImage.hDC, 0, 0, SRCINVERT)
    
    Else
    
    Call TrackArea(inputImage)
  End If

  prev_tracking = Tracking

End Sub



Private Sub TrackArea(inputImage As PictureBox)
'tracks the movement of a small area of the image
  Static W As Integer
  Static H As Integer
  Static stp_x As Integer
  Static stp_y As Integer
  Static stp_x2 As Integer
  Static stp_y2 As Integer
  Static firstCall As Integer
  Dim p As Long
  Dim x As Integer
  Dim y As Integer
  Dim xx As Integer
  Dim yy As Integer
  Dim i As Integer
  Dim ox As Integer
  Dim oy As Integer
  Dim dist As Long
  Dim minDist As Long
  Dim p1 As Integer
  Dim p2 As Integer
  Static maxDifference As Long
  Static dx As Single
  Static dy As Single
  Static ticks As Long
  Dim tx As Integer
  Dim ty As Integer
  
  If (firstCall = 0) Then
    W = inputImage.ScaleWidth / 15
    H = inputImage.ScaleHeight / 15
    stp_x = (W * 2) / TemplateSize
    stp_y = (H * 2) / TemplateSize
    stp_x2 = stp_x * 2
    stp_y2 = stp_y * 2
    firstCall = 1
    maxDifference = TemplateSize * TemplateSize * 30
  End If
  
  If (Not prev_tracking) And (Tracking) Then
    track_x = attention_x / MotionResolution * inputImage.ScaleWidth
    track_y = attention_y / MotionResolution * inputImage.ScaleHeight
    dx = 0
    dy = 0
    ticks = 0
  End If
    
  tx = track_x
  ty = track_y
  minDist = maxDifference
  For i = 0 To 8
    Select Case i
      Case 0
        ox = tx + stp_x2
        oy = ty - stp_y2
      Case 1
        ox = tx + stp_x2
        oy = ty
      Case 2
        ox = tx + stp_x2
        oy = ty + stp_y2
      Case 3
        ox = tx
        oy = ty + stp_y2
      Case 4
        ox = tx - stp_x2
        oy = ty + stp_y2
      Case 5
        ox = tx - stp_x2
        oy = ty
      Case 6
        ox = tx - stp_x2
        oy = ty - stp_y2
      Case 7
        ox = tx
        oy = ty - stp_y2
      Case 8
        ox = tx
        oy = ty
    End Select

    dist = 0
    x = 0
    For xx = ox - W To ox + W Step stp_x
      y = 0
      For yy = oy - H To oy + H Step stp_y
        p = inputImage.Point(xx, yy)
        rgbsource.Value = p
        Call CopyMemory(rgbdest, rgbsource, 3)
        p1 = rgbdest.Red
        p2 = template(x, y)
        dist = dist + Abs(p1 - p2)
        
        If (i = 8) Then
          template(x, y) = rgbdest.Red
        End If
      
        y = y + 1
      Next
      x = x + 1
    Next
   
    If (dist < minDist) Then
      minDist = dist
      tx = ox
      ty = oy
    End If
   
  Next
  
  dx = dx + (track_x - tx)
  dy = dy + (track_y - ty)
  If (ticks > 4) Then
    velocity_x = dx / 5
    velocity_y = dy / 5
    ticks = 0
    dx = 0
    dy = 0
  End If
  ticks = ticks + 1
  
  track_x = tx
  track_y = ty
  
  If (prev_tracking) And (Tracking) Then
    If (minDist <> maxDifference) Then
      attention_x = track_x / inputImage.ScaleWidth * MotionResolution
      attention_y = track_y / inputImage.ScaleHeight * MotionResolution
      Else
      Tracking = False
    End If
  End If
  
End Sub



Public Sub Vision_Filter(inputImage As PictureBox, colourImage As PictureBox, motionImage As PictureBox)
'BitBlit motion compare
  Const SRCCOPY = &HCC0020
  Const SRCINVERT = &H660046
  Dim rc As Long
    
  Call Vision_VFWgrab(inputImage)
  rc = BitBlt(motionImage.hDC, 0, 0, inputImage.ScaleWidth, inputImage.ScaleHeight, colourImage.hDC, 0, 0, SRCCOPY)
  rc = BitBlt(motionImage.hDC, 0, 0, inputImage.ScaleWidth, inputImage.ScaleHeight, inputImage.hDC, 0, 0, SRCINVERT)
End Sub


Private Sub getAttentionBoxes(ByRef cx As Integer, ByRef cy As Integer)
  Dim x As Integer
  Dim y As Integer
  Dim minX As Integer
  Dim minY As Integer
  Dim maxX As Integer
  Dim maxY As Integer
  Dim maxRegionArea As Integer
  Dim biggest As Integer
  
  maxRegionArea = 0
  biggest = 0
  NoOfAttentionBoxes = 0
  For x = 0 To MotionResolution - 1
    For y = 0 To MotionResolution - 1
      If (motion(x, y, 0)) And (Not motion(x, y, 1)) Then
        RegionArea = 0
        minX = x
        minY = y
        maxX = x
        maxY = y
        Call fillRegion(x, y, 0, minX, maxX, minY, maxY)
        If (RegionArea > 5) And (NoOfAttentionBoxes < 100) Then
          If (RegionArea > maxRegionArea) Then
            maxRegionArea = RegionArea
            biggest = NoOfAttentionBoxes
          End If
          attentionBox(NoOfAttentionBoxes, 0) = minX
          attentionBox(NoOfAttentionBoxes, 1) = minY
          attentionBox(NoOfAttentionBoxes, 2) = maxX
          attentionBox(NoOfAttentionBoxes, 3) = maxY
          NoOfAttentionBoxes = NoOfAttentionBoxes + 1
        End If
      End If
    Next
  Next
  
  If (NoOfAttentionBoxes > 0) Then
    cx = attentionBox(biggest, 0) + ((attentionBox(biggest, 2) - attentionBox(biggest, 0)) / 2)
    cy = attentionBox(biggest, 1) + ((attentionBox(biggest, 3) - attentionBox(biggest, 1)) / 2)
  End If
End Sub


Private Sub showAttentionBoxes(inputImage As PictureBox, outputImage As PictureBox)
  Dim minX As Integer
  Dim minY As Integer
  Dim maxX As Integer
  Dim maxY As Integer
  Dim tx As Integer
  Dim ty As Integer
  Dim bx As Integer
  Dim by As Integer
  Dim i As Integer
  Dim rc As Long
  Dim sx As Integer
  Dim sy As Integer
  Dim c As Long
  
  outputImage.Cls
  outputImage.FillStyle = 1
  outputImage.DrawWidth = 1

  sx = inputImage.ScaleWidth / MotionResolution
  sy = inputImage.ScaleHeight / MotionResolution
  For i = 0 To NoOfAttentionBoxes - 1
    minX = attentionBox(i, 0)
    minY = attentionBox(i, 1)
    maxX = attentionBox(i, 2)
    maxY = attentionBox(i, 3)
    
    tx = ((minX / MotionResolution) * inputImage.ScaleWidth)
    ty = ((minY / MotionResolution) * inputImage.ScaleHeight)
    bx = ((maxX / MotionResolution) * inputImage.ScaleWidth)
    by = ((maxY / MotionResolution) * inputImage.ScaleHeight)
    
    c = RGB(0, 255, 0)
    outputImage.Line (tx, ty)-(bx, by), c, B
  Next
End Sub


Public Sub showFovea(inputImage As PictureBox, foveaImage As PictureBox)
'shows the fovea region
  Dim tx As Integer
  Dim ty As Integer
  Dim bx As Integer
  Dim by As Integer
  Dim txx As Integer
  Dim tyy As Integer
  Dim bxx As Integer
  Dim byy As Integer
  Dim W As Integer
  Dim H As Integer
  Dim rc As Long
  Dim sx As Integer
  Dim sy As Integer
  
  foveaImage.Cls
  
  W = MotionResolution / 4
  H = MotionResolution / 4
  tx = attention_x - W + 1
  If (tx < 0) Then
    tx = 0
  End If
  bx = attention_x + W + 1
  If (bx > MotionResolution - 1) Then
    bx = MotionResolution - 1
  End If
  ty = attention_y - H
  If (ty < 0) Then
    attention_y = attention_y - ty
    ty = 0
  End If
  by = attention_y + H
  If (by > MotionResolution - 1) Then
    by = MotionResolution - 1
  End If
  
  txx = tx / MotionResolution * inputImage.ScaleWidth
  tyy = ty / MotionResolution * inputImage.ScaleHeight
  bxx = bx / MotionResolution * inputImage.ScaleWidth
  byy = by / MotionResolution * inputImage.ScaleHeight
  sx = (bx - tx) / MotionResolution * inputImage.ScaleWidth
  sy = (by - ty) / MotionResolution * inputImage.ScaleHeight
  rc = BitBlt(foveaImage.hDC, 0, 0, sx, sy, inputImage.hDC, txx, tyy, SRCCOPY)
  
End Sub


Private Sub fillRegion(px As Integer, py As Integer, depth As Integer, ByRef minX As Integer, ByRef maxX As Integer, ByRef minY As Integer, ByRef maxY As Integer)
'fills a region with the given colour
  
  If (motion(px, py, 0)) And (motion(px, py, 1) = 0) Then
  
    motion(px, py, 1) = True
    RegionArea = RegionArea + 1
    
    If (px < minX) Then
      minX = px
    End If
    If (px > maxX) Then
      maxX = px
    End If
    If (py < minY) Then
      minY = py
    End If
    If (py > maxY) Then
      maxY = py
    End If
    
    If (depth < 100) Then
    
      If (py > 0) Then
        Call fillRegion(px, py - 1, depth + 1, minX, maxX, minY, maxY)
      End If
      
      If (px < MotionResolution - 1) And (py > 0) Then
        Call fillRegion(px + 1, py - 1, depth + 1, minX, maxX, minY, maxY)
      End If
      
      If (px < MotionResolution - 1) Then
        Call fillRegion(px + 1, py, depth + 1, minX, maxX, minY, maxY)
      End If
      
      If (px < MotionResolution - 1) And (py < MotionResolution - 2) Then
        Call fillRegion(px + 1, py + 1, depth + 1, minX, maxX, minY, maxY)
      End If
      
      If (py < MotionResolution - 1) Then
        Call fillRegion(px, py + 1, depth + 1, minX, maxX, minY, maxY)
      End If
      
      If (px > 0) Then
        Call fillRegion(px - 1, py, depth + 1, minX, maxX, minY, maxY)
      End If
    
      If (px > 0) And (py > 0) Then
        Call fillRegion(px - 1, py - 1, depth + 1, minX, maxX, minY, maxY)
      End If
      
      If (px > 0) And (py < MotionResolution - 1) Then
        Call fillRegion(px - 1, py + 1, depth + 1, minX, maxX, minY, maxY)
      End If
      
      If (px < MotionResolution - 1) And (py < MotionResolution - 1) Then
        Call fillRegion(px + 1, py + 1, depth + 1, minX, maxX, minY, maxY)
      End If
      
    End If
  End If
End Sub


VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   ControlBox      =   0   'False
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim dxMain As New DirectX7       ' Main object

Dim ddMain As DirectDraw7        ' Directdraw object

Dim diMain As DirectInput        ' DirectInput object
Dim diDev As DirectInputDevice   ' DirectInput device
Dim diState As DIKEYBOARDSTATE   ' DirectInput keyboard

Dim sdMain As DDSURFACEDESC2     ' Main surface description
Dim dsPrim As DirectDrawSurface7 ' Primary screen buffer
Dim dsBbuf As DirectDrawSurface7 ' Screen backbuffer

Dim sdStar As DDSURFACEDESC2     ' Star surface description
Dim dsStar As DirectDrawSurface7 ' Star surface

Dim sAngle As Single             ' Current star angle
Dim sSpd As Single               ' Speed multiplier
Dim sStar(150, 3) As Single      ' Star array

Sub Do_SetStars()

  ' In this sub all star values are loaded into an array
  ' all values are random...

  Dim Xas As Integer
  For Xas = 0 To 149                   ' There are 150 stars
    Let sStar(Xas, 0) = Int(Rnd * 320) ' Set starting point (X)
    Let sStar(Xas, 1) = Int(Rnd * 240) ' Set starting point (Y)
    Let sStar(Xas, 2) = Rnd * 5        ' Set moving speed
    Let sStar(Xas, 3) = Fix(4 - Int(sStar(Xas, 2)))  ' Calculate image, see note
  Next Xas

  ' Note: The bitmap used with this sample uses 5 stars
  ' The first star is the brightest and the last one the
  ' Darkest. To add realism the fastest moving stars will
  ' Be the ones using the first picture because it is the
  ' Closest to calculate this we round the star speed and
  ' Invert the number...

End Sub

Sub DXMain_Init()
    
  ' This sub initialises the DirectX components and sets up
  ' The screen and keyboard...
    
  On Local Error GoTo errOut ' Event handler
   
  Set ddMain = dxMain.DirectDrawCreate("") ' Create main DirectX object
  Me.Show ' Show form

  Call ddMain.SetCooperativeLevel(Me.hWnd, DDSCL_FULLSCREEN Or DDSCL_ALLOWMODEX Or DDSCL_EXCLUSIVE)
  ddMain.SetDisplayMode 320, 240, 16, 0, DDSDM_DEFAULT ' Set screen the 320x240x16

  ' Set up the primary (screen) surface, this surface will
  ' Be the one that is shown on the screen
  sdMain.lFlags = DDSD_CAPS Or DDSD_BACKBUFFERCOUNT
  sdMain.ddsCaps.lCaps = DDSCAPS_PRIMARYSURFACE Or DDSCAPS_FLIP Or DDSCAPS_COMPLEX
  sdMain.lBackBufferCount = 1
  Set dsPrim = ddMain.CreateSurface(sdMain)

  ' Create a backbuffer, this is used to draw a surface while
  ' Another is shown on the screen this is used to reduce
  ' Filckering and screen build-ups
  Dim ddCaps As DDSCAPS2
  ddCaps.lCaps = DDSCAPS_BACKBUFFER
  Set dsBbuf = dsPrim.GetAttachedSurface(ddCaps)
 
  Do_SetStars         ' Load stars
  DXMain_InitSurfaces ' Load surfaces
    
  Let sAngle = 0      ' Set the turning angle to 0
  Let sSpd = 1        ' Set speed multiplier to 1
    
  ' Create the DirectInput handlers, these are used to
  ' Control the stars' movement
  Set diMain = dxMain.DirectInputCreate()
  Set diDev = diMain.CreateDevice("GUID_SysKeyboard")
  diDev.SetCommonDataFormat DIFORMAT_KEYBOARD
  diDev.SetCooperativeLevel Me.hWnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE
  diDev.Acquire
  
  ' Start an infinite loop, this is preferable to a timer
  ' Because it is more precise and faster. However it does
  ' Cause a slowdown on older PC's
  Do
    Do_Keys    ' Check for key input
    DXMain_Blt ' Draw screen
    DoEvents   ' Let DirectX draw to the screen
  Loop

errOut:
  DXMain_EndIt

End Sub

Sub DXMain_InitSurfaces()

  ' This sub loads the star surface containing our five
  ' Star pictures

  Dim ClrKey As DDCOLORKEY        ' Create a color key, this is used
  ClrKey.low = 0: ClrKey.high = 0 ' To make the stars transparent
    
  ' Create the star surface
  sdStar.lFlags = DDSD_CAPS Or DDSD_HEIGHT Or DDSD_WIDTH
  sdStar.ddsCaps.lCaps = DDSCAPS_OFFSCREENPLAIN
  sdStar.lWidth = 15: sdStar.lHeight = 3    ' Set up size
  Set dsStar = ddMain.CreateSurfaceFromFile(App.Path + "\stars.bmp", sdStar) ' Load the bitmap
  dsStar.SetColorKey DDCKEY_SRCBLT, ClrKey  ' Set the colorkey

End Sub

Public Sub Do_Keys()
    
  ' This sub processes the DirectInput commands
    
  Dim Xas As Integer
  diDev.GetDeviceStateKeyboard diState     ' Get the keystates

  ' The escape-key is used to end the program
  If diState.Key(1) <> 0 Then DXMain_EndIt
  
  ' The left key reduces the angle, if the angle gets
  ' Below 0 it is reset to two pi (6.28318530718)
  If diState.Key(205) <> 0 Then
    Let sAngle = sAngle - 0.025
    If sAngle < 0 Then Let sAngle = 6.28
  End If
  
  ' The right key, the angle is increased, if above 2xPi
  ' It is reset to 0
  If diState.Key(203) <> 0 Then
    Let sAngle = sAngle + 0.025
    If sAngle > 6.28 Then Let sAngle = 0
  End If
  
  ' The down key reduces speed, and thus the multiplier
  If diState.Key(208) <> 0 Then
    Let sSpd = sSpd * 0.99
  End If
  
  ' The up key increases speed, if the speed is higher than
  ' 25 it is set back to 25
  If diState.Key(200) <> 0 Then
    Let sSpd = sSpd * 1.01
    If sSpd > 25 Then Let sSpd = 25
  End If

End Sub

Sub DXMain_Blt()

  ' This is the main drawing sub, all stars are drawn onto
  ' The backbuffer and then the backbuffer is swapped with
  ' The primary screen buffer

  On Local Error GoTo errOut ' Error handler
  
  Dim rBack As RECT  ' A RECT is used to set the picture size
  Dim Xas As Integer
  
  ' Set the fillcolor to black and paint the screen black
  dsBbuf.SetFillColor 0
  dsBbuf.DrawBox 0, 0, 320, 240
  
  ' Set the star height to 3 pixels
  rBack.Top = 0: rBack.Bottom = 3
  ' Draw and move the 150 stars
  For Xas = 0 To 149
    ' Define the picture used using the picture number set
    ' In the array, this is cleaner, faster and easier then
    ' Using a single bitmap for every star
    rBack.Left = sStar(Xas, 3) * 3
    rBack.Right = rBack.Left + 3
    ' Draw the star onto the backbuffer
    dsBbuf.BltFast sStar(Xas, 0), sStar(Xas, 1), dsStar, rBack, DDBLTFAST_WAIT Or DDBLTFAST_SRCCOLORKEY
    ' Move the star, sine and cosine are used to move the
    ' Stars under a certain angle
    Let sStar(Xas, 1) = sStar(Xas, 1) + (sStar(Xas, 2) * Cos(sAngle) * sSpd)
    Let sStar(Xas, 0) = sStar(Xas, 0) + (sStar(Xas, 2) * Sin(sAngle) * sSpd)
    ' Check if the stars are off the screen and if so,
    ' Put them back
    If sStar(Xas, 1) < 0 Then Let sStar(Xas, 1) = 240 + sStar(Xas, 1)
    If sStar(Xas, 1) > 240 Then Let sStar(Xas, 1) = sStar(Xas, 1) - 240
    If sStar(Xas, 0) < 0 Then Let sStar(Xas, 0) = 320 + sStar(Xas, 0)
    If sStar(Xas, 0) > 320 Then Let sStar(Xas, 0) = sStar(Xas, 0) - 320
  Next Xas
  
  ' Swap the backbuffer and the primary buffer
  dsPrim.Flip Nothing, DDFLIP_WAIT
    
errOut:

End Sub

Sub DXMain_EndIt()
  
  ' This sub unloads DirectX and puts the cursor back
  ' On the screen
  
  ShowCursor 1                   ' Restore the cursor
  Call ddMain.RestoreDisplayMode ' Reset the screen
  Call ddMain.SetCooperativeLevel(Me.hWnd, DDSCL_NORMAL)
  Call diDev.Unacquire           ' Disable DirectInput
  
  ' Terminate the program
  End

End Sub

Private Sub Form_Load()

  ' This sub calls the initialisation sub and hides the
  ' Mouse cursor

  ShowCursor 0
  DXMain_Init

End Sub

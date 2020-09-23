VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{2B143B63-055B-11D2-A96D-00A0C92A2D0F}#12.0#0"; "hcwWinTV.ocx"
Begin VB.Form frmAppSample 
   Caption         =   "hcwWinTVOCX Sample Application"
   ClientHeight    =   6210
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   9705
   Icon            =   "Sample.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6210
   ScaleWidth      =   9705
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame5 
      Caption         =   "Motion Capture"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   660
      Left            =   0
      TabIndex        =   29
      Top             =   3750
      Width           =   2880
      Begin VB.CommandButton cmdMotionPalyback 
         Caption         =   "Play"
         Height          =   285
         Left            =   2280
         TabIndex        =   33
         ToolTipText     =   "Playback Captured AVI"
         Top             =   270
         Width           =   465
      End
      Begin VB.CommandButton cmdMotionCapture 
         Caption         =   "Capture"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1425
         TabIndex        =   32
         ToolTipText     =   "Start Motion Capture"
         Top             =   270
         Width           =   855
      End
      Begin VB.CommandButton cmdMotionCapFormat 
         Caption         =   "Format"
         Height          =   285
         Left            =   735
         TabIndex        =   31
         ToolTipText     =   "Motion Capture Format"
         Top             =   270
         Width           =   675
      End
      Begin VB.CommandButton cmdMotionCapSetup 
         Caption         =   "Setup"
         Height          =   285
         Left            =   105
         TabIndex        =   30
         ToolTipText     =   "Motion Capture Setup"
         Top             =   270
         Width           =   630
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Settings 3"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   915
      Left            =   0
      TabIndex        =   17
      Top             =   2835
      Width           =   2865
      Begin VB.CommandButton cmdSnapshot 
         Caption         =   "Snapshot"
         Height          =   255
         Left            =   1830
         TabIndex        =   28
         ToolTipText     =   "800*600 Snapshot"
         Top             =   240
         Width           =   915
      End
      Begin VB.CommandButton cmdShowColorControlBox 
         Caption         =   "Color..."
         Height          =   255
         Left            =   975
         TabIndex        =   22
         ToolTipText     =   "Show Color Config Dialog"
         Top             =   555
         Width           =   915
      End
      Begin VB.CommandButton cmdShowAudioControlBox 
         Caption         =   "Audio..."
         Height          =   255
         Left            =   135
         TabIndex        =   21
         ToolTipText     =   "Show Audio Config Dialog"
         Top             =   555
         Width           =   840
      End
      Begin VB.CommandButton cmdSurf 
         Caption         =   "Surf"
         Height          =   255
         Left            =   1890
         TabIndex        =   20
         ToolTipText     =   "Channel Surf"
         Top             =   555
         Width           =   855
      End
      Begin VB.CommandButton cmdSaveToDisk 
         Caption         =   "Save As"
         Height          =   255
         Left            =   930
         TabIndex        =   19
         ToolTipText     =   "Save Image As"
         Top             =   240
         Width           =   885
      End
      Begin VB.CommandButton cmdFreeze 
         Caption         =   "Freeze"
         Height          =   255
         Left            =   135
         TabIndex        =   18
         ToolTipText     =   "Freeze Video"
         Top             =   240
         Width           =   795
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Settings 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1485
      Left            =   0
      TabIndex        =   10
      Top             =   1335
      Width           =   2865
      Begin VB.OptionButton optionToolBarPosition 
         Caption         =   "None"
         Height          =   255
         Index           =   4
         Left            =   1980
         TabIndex        =   26
         Top             =   1185
         Width           =   735
      End
      Begin VB.OptionButton optionToolBarPosition 
         Caption         =   "Bottom"
         Height          =   225
         Index           =   3
         Left            =   990
         TabIndex        =   25
         Top             =   1200
         Width           =   885
      End
      Begin VB.OptionButton optionToolBarPosition 
         Caption         =   "Top"
         Height          =   255
         Index           =   2
         Left            =   135
         TabIndex        =   24
         Top             =   1185
         Width           =   675
      End
      Begin VB.OptionButton optionToolBarPosition 
         Caption         =   "Right"
         Height          =   255
         Index           =   1
         Left            =   1980
         TabIndex        =   23
         Top             =   915
         Width           =   705
      End
      Begin VB.OptionButton optionToolBarPosition 
         Caption         =   "Left"
         Height          =   255
         Index           =   0
         Left            =   1365
         TabIndex        =   14
         Top             =   915
         Width           =   645
      End
      Begin ComctlLib.Slider sliderContrast 
         Height          =   225
         Left            =   900
         TabIndex        =   12
         Top             =   300
         Width           =   1845
         _ExtentX        =   3254
         _ExtentY        =   397
         _Version        =   327682
         LargeChange     =   15
         SmallChange     =   5
         Max             =   255
         TickFrequency   =   17
      End
      Begin ComctlLib.Slider sliderBrightness 
         Height          =   225
         Left            =   900
         TabIndex        =   16
         Top             =   600
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   397
         _Version        =   327682
         LargeChange     =   15
         SmallChange     =   5
         Max             =   255
         TickFrequency   =   17
      End
      Begin VB.Label Label11 
         Caption         =   "Brightness:"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   570
         Width           =   885
      End
      Begin VB.Label Label6 
         Caption         =   "Toolbar Position:"
         Height          =   225
         Left            =   120
         TabIndex        =   13
         Top             =   915
         Width           =   1425
      End
      Begin VB.Label Label4 
         Caption         =   "Contrast:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   300
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Settings 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   1230
      Left            =   15
      TabIndex        =   1
      Top             =   90
      Width           =   2865
      Begin VB.CommandButton cmdChanExplorer 
         Caption         =   "Chan Explorer"
         Height          =   255
         Left            =   1290
         TabIndex        =   27
         ToolTipText     =   "Channel Explorer Dialog"
         Top             =   540
         Width           =   1380
      End
      Begin ComctlLib.Slider SliderVolume 
         Height          =   210
         Left            =   720
         TabIndex        =   9
         Top             =   885
         Width           =   1995
         _ExtentX        =   3519
         _ExtentY        =   370
         _Version        =   327682
         LargeChange     =   20
         SmallChange     =   5
         Max             =   100
         TickFrequency   =   20
      End
      Begin VB.CommandButton cmdChannelDown 
         Caption         =   "-"
         Height          =   255
         Left            =   1050
         TabIndex        =   5
         ToolTipText     =   "Channel Down"
         Top             =   540
         Width           =   240
      End
      Begin VB.CommandButton cmdChannelUp 
         Caption         =   "+"
         Height          =   255
         Left            =   810
         TabIndex        =   4
         ToolTipText     =   "Channel Up"
         Top             =   540
         Width           =   240
      End
      Begin VB.OptionButton optionComposite 
         Caption         =   "Composite"
         Height          =   225
         Left            =   1620
         TabIndex        =   3
         ToolTipText     =   "Select Composite Input"
         Top             =   255
         Width           =   1125
      End
      Begin VB.OptionButton optionTuner 
         Caption         =   "Tuner"
         Height          =   225
         Left            =   810
         TabIndex        =   2
         ToolTipText     =   "Select TV "
         Top             =   270
         Value           =   -1  'True
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Volume:"
         Height          =   225
         Left            =   150
         TabIndex        =   8
         Top             =   870
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Channel:"
         Height          =   225
         Left            =   120
         TabIndex        =   7
         Top             =   570
         Width           =   705
      End
      Begin VB.Label Label1 
         Caption         =   "Source:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   255
         Width           =   585
      End
   End
   Begin VB.Frame frameVideo 
      Caption         =   "Hauppauge WinTV OCX"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   4305
      Left            =   2910
      TabIndex        =   0
      Top             =   90
      Width           =   5715
      Begin hcwWinTVControl.hcwWinTVocx hcwWinTVocx1 
         Height          =   4095
         Left            =   120
         TabIndex        =   34
         Top             =   180
         Width           =   5535
         _ExtentX        =   9763
         _ExtentY        =   7223
      End
   End
   Begin VB.Menu menuFile 
      Caption         =   "&File"
      Begin VB.Menu menuCCon 
         Caption         =   "&Closed Caption"
      End
      Begin VB.Menu menuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu menuHelp 
      Caption         =   "&Help"
      Begin VB.Menu menuAbout 
         Caption         =   "&About"
      End
   End
End
Attribute VB_Name = "frmAppSample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/*******************************************************************************
'*
'*       HH  HH     CCCC  WW           WW
'*       HH  HH   CCCCCC  WW           WW
'*       HH  HH  CCC      WW           WW
'*       HHHHHH  CCC       WW   WWW   WW
'*       HH  HH  CCC       WW  WW WW  WW
'*       HH  HH   CCCCCC    WWWW   WWWW
'*       HH  HH     CCCC     WW     WW
'*
'*       Copyright (C) 1994-1998
'*       Hauppauge Computer Works, Inc
'*       91 Cabot Court
'*       Hauppauge, NY  11788
'*       516 / 434-1600
'*
'*******************************************************************************/
'********************************************************************************
'*
'*           Hauppauge Computer Works WinTV OCX Sample Application
'*                      for use with hcwWinTV.ocx
'*
'********************************************************************************

' This sample application is to show how to use the Hauppauge WinTV OCX in Visual
' Basic. Be sure hcwWinTV.ocx is installed and registered.

' This sample application is assuming the user's WinTV board has both a Tuner and
' a Composite video input, actual application should detect the hardware avalibility
' first, and deals with the inputs accordingly.

' This application is just a simple example, it may not be pratical, or complete.
' Compiled and tested in Visual Basic 6.0.

Private Sub cmdChanExplorer_Click()
    ' bring up the channel explorer dialog
    hcwWinTVocx1.ShowChannelExplorer
End Sub

Private Sub cmdChannelDown_Click()
    ' tune down the channel
    hcwWinTVocx1.ChannelDown
End Sub

Private Sub cmdChannelUp_Click()
    ' tune up the channel
    hcwWinTVocx1.ChannelUp
End Sub

Private Sub cmdMotionCapFormat_Click()
    ' bring up the Motion Capture Format config dialog
    hcwWinTVocx1.MotionCapFormat
End Sub

Private Sub cmdMotionCapSetup_Click()
    'bring up the Motion Capture Setup Dialog
    hcwWinTVocx1.MotionCapSetup
End Sub

Private Sub cmdMotionCapture_Click()
    ' Start the Motion Capture, a mouse click or hit the Esc key will stop the capturing
    hcwWinTVocx1.MotionCapture
End Sub

Private Sub cmdMotionPalyback_Click()
    ' play back the AVI file captured by the program
    hcwWinTVocx1.MotionPlayback
End Sub

Private Sub cmdSaveToDisk_Click()
    ' this call display the "Print..." dialog, and
    ' save the image to a disk file
    Call hcwWinTVocx1.SaveToDisk
End Sub

Private Sub cmdFreeze_Click()
    ' freeze or unfreeze the video according to its
    ' previous status
    Call hcwWinTVocx1.ToggleFreezeVideo
End Sub

Private Sub cmdShowAudioControlBox_Click()
    ' show the audio config dialog box
    Call hcwWinTVocx1.ShowAudioControlBox
End Sub

Private Sub cmdShowColorControlBox_Click()
    ' show the color config dialog box
    Call hcwWinTVocx1.ShowColorControlBox
End Sub

Private Sub cmdSnapshot_Click()
    ' do a 800*600 snapshot
    hcwWinTVocx1.SnapShot (800)
End Sub

Private Sub cmdSurf_Click()
    ' do channel surfing, press the button again will stop the surfing.
    If hcwWinTVocx1.SurfMode = False Then
        hcwWinTVocx1.SurfStart
    Else
        hcwWinTVocx1.SurfEnd
    End If
End Sub

Private Sub Command1_Click()

End Sub

Private Sub Form_Load()
    ' during the form load, we need to get some information
    ' from the winTV control, and adjust the display accordingly.
    
    ' enable the WinTV control
    hcwWinTVocx1.Enabled = True
    
    ' determine the current video input source
    Select Case hcwWinTVocx1.VideoSourceByType
        Case hcwVideoSource_Tuner
            optionTuner.Value = True
        Case hcwVideoSource_Composite1, _
                hcwVideoSource_Composite2, _
                hcwVideoSource_Composite3
            optionComposite.Value = True
        Case Else
            optionTuner.Value = False
            optionComposite.Value = False
    End Select
    
    ' determin the current volume level
    SliderVolume.Value = hcwWinTVocx1.Volume
    
    ' determine the current Contrast and HSB level
    sliderContrast.Value = hcwWinTVocx1.Contrast
    sliderBrightness.Value = hcwWinTVocx1.Brightness
    
    ' determine the current Control Panel Position
    optionToolBarPosition(hcwWinTVocx1.ToolBarPosition).Value = True
End Sub

Private Sub Form_Resize()
    ' if form is minimized, don't do anything
    If frmAppSample.WindowState = 1 Then   '1 - minimized
        Exit Sub
    End If
    ' limit the form size, be sure all the contents can be seen
    If frmAppSample.ScaleWidth < Frame2.Width + 4000 Then
        frmAppSample.Width = Frame2.Width + 4000
    End If
    If frmAppSample.ScaleHeight < Frame2.Height + Frame3.Height + Frame4.Height + Frame5.Height + 200 Then
        frmAppSample.Height = Frame2.Height + Frame3.Height + Frame4.Height + Frame5.Height + 850
    End If
    ' resize the video frame and OCX to follow the size of form
    frameVideo.Left = Frame2.Width + 30
    frameVideo.Top = 90
    frameVideo.Width = frmAppSample.ScaleWidth - Frame2.Width - 30
    frameVideo.Height = frmAppSample.ScaleHeight - 120
    hcwWinTVocx1.Width = frameVideo.Width - 200
    hcwWinTVocx1.Height = frameVideo.Height - 320
End Sub

Private Sub hcwWinTVocx1_HSBChanged(Contrast As Long, _
                                    Hue As Long, _
                                    Saturation As Long, _
                                    Brightness As Long)
    ' adjust the sliders according to the new HSB levels
    sliderContrast.Value = Contrast
    sliderBrightness.Value = Brightness
End Sub

Private Sub hcwWinTVocx1_InputSourceChanged(InputSource As _
                                hcwWinTVControl.hcwVideoSource)
    ' adjust the input source option buttons according to the new source
    If InputSource = hcwVideoSource_Tuner Then
        optionTuner.Value = True
    End If
    If InputSource = hcwVideoSource_Composite1 Or _
            InputSource = hcwVideoSource_Composite2 Or _
            InputSource = hcwVideoSource_Composite3 Then
        optionComposite.Value = True
    End If
End Sub

Private Sub hcwWinTVocx1_VolumeChanged(Volume As Long)
    ' adjust the volume slider according to the new volume
    SliderVolume.Value = Volume
End Sub

Private Sub menuAbout_Click()
    ' show the WinTV OCX about box
    Call hcwWinTVocx1.ShowAboutBox
End Sub

Private Sub menuCCon_Click()
    ' Turn CC on.
    If menuCCon.Checked Then
        menuCCon.Checked = False
    Else
        menuCCon.Checked = True
    End If
    
    If hcwWinTVocx1.GetCCState = True Then
        hcwWinTVocx1.ClosedCaption = menuCCon.Checked
    Else
        MyVar = MsgBox("Closed Caption Not Available!", 65, "WinTV Example")  ' MyVar contains either 1 or 2,
    End If
    
  
End Sub

Private Sub menuExit_Click()
    ' when exit the program, the OCX will terminate by itself,
    ' no need to take care of it
    Unload frmAppSample
End Sub

Private Sub optionComposite_Click()
    ' switch the video input source to the first Composite input
    hcwWinTVocx1.VideoSourceByType = hcwVideoSource_Composite1
End Sub

Private Sub optionToolBarPosition_Click(Index As Integer)
    ' set the ToolBarPostion
    hcwWinTVocx1.ToolBarPosition = Index
End Sub

Private Sub optionTuner_Click()
    ' switch the video input source to the Tuner
    hcwWinTVocx1.VideoSourceByType = hcwVideoSource_Tuner
End Sub

Private Sub SliderContrast_Scroll()
    ' adjust the picture Contrast by the slider value
    hcwWinTVocx1.Contrast = sliderContrast.Value
End Sub

Private Sub sliderBrightness_Scroll()
    ' adjust the picture Brightness level by the slider value
    hcwWinTVocx1.Brightness = sliderBrightness.Value
End Sub

Private Sub SliderVolume_Scroll()
    ' adjust the volume by the slider value
    hcwWinTVocx1.Volume = SliderVolume.Value
End Sub

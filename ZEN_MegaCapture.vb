
'global
Dim displayedWarning As Boolean
Dim doImage As Boolean
Dim TimeIndex As Integer
Dim SpecimenRowIndex As Integer
Dim SpecimenColumnIndex As Integer
Dim XTilesIndex As Integer
Dim YTilesIndex As Integer
Dim intOutFile As Integer
Dim strOutFile As String
Dim success As Integer
Dim strFilename As String
Dim redChannel As Integer
Dim greenChannel As Integer
Dim blueChannel As Integer

'end global

Private Type BROWSEINFO ' used by the function GetFolderName
    hOwner As Long
    pidlRoot As Long
    pszDisplayName As String
    lpszTitle As String
    ulFlags As Long
    lpfn As Long
    lParam As Long
    iImage As Long
End Type

Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function SHGetPathFromIDList Lib "shell32.dll" _
    Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long

Private Declare Function SHBrowseForFolder Lib "shell32.dll" _
    Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long
    
Sub StartMegaCapture()
    
    Dim Track As DsTrack
    Dim Laser As DsLaser
    Dim DetectionChannel As DsDetectionChannel
    Dim IlluminationChannel As DsIlluminationChannel
    Dim DataChannel As DsDataChannel
    Dim BeamSplitter As DsBeamSplitter
    Dim Timers As DsTimers
    Dim Markers As DsMarkers
    Dim Stage As CpStages
    Dim TimePointStartTime As Date
    Dim strFileExtension As String
    Dim z As Integer
    Dim FOV As Integer
        
    FOV = 0
    
    'Create .mgc file to save parameters for all images for importing to GoFigure
    intOutFile = FreeFile
    strOutFile = PathOfFolderForImagesText + FilenamePrefixText + ".meg"
    Close #intOutFile
    Open strOutFile For Output As #intOutFile
    Dim sTab As String
    sTab = Chr(9) 'tab
    
    'Write header for .mgc
    Print #intOutFile, "MegaCaptureSummary-version1.0"
    Print #intOutFile, ExperimentTitleText
    Print #intOutFile, ExperimentDescriptionText
    Print #intOutFile, CStr(VoxelSizeX) + sTab + CStr(VoxelSizeY) + sTab + CStr(VoxelSizeZ) + sTab + TimeIntervalText
    
    'Write column labels to .mgc file
    Print #intOutFile, "Filename" + sTab + "RIndex" + sTab + "CIndex" + sTab + "XIndex" + sTab + "YIndex" + sTab + "ZIndex" + sTab + "TIndex" + sTab + "X-Offset" + sTab + "Y-Offset" + sTab + "Z-Offset" + sTab + "CaptureDate" + sTab + "Pinhole1" + sTab + "Pinhole2" + sTab + "Pinhole3" + sTab + "Gain1" + sTab + "Gain2" + sTab + "Gain3" + sTab + "Attenuation488" + sTab + "Attenuation543" + sTab + "Attenuation820"
    
    'Enable nudging buttons
    XOffsetSpin.Enabled = True
    YOffsetSpin.Enabled = True
    ZOffsetSpin.Enabled = True
    
    'Remember starting stage position
    startX = Lsm5.Hardware.CpStages.PositionX
    startY = Lsm5.Hardware.CpStages.PositionY
    Dim strMessage As String
    strMessage = "startX=" + CStr(startX) + "   startY=" + CStr(startY)
'    MsgBox (strMessage)
  
    Dim stageIndex As Long
    strFilename = ""
    FOV = 0
    redChannel = 1
    greenChannel = 2
    blueChannel = 3
    
    'get red channel
    If (OptionCh1R) Then redChannel = 0
    If (OptionCh2R) Then redChannel = 1
    If (OptionCh3R) Then redChannel = 2
    If (OptionChMR) Then redChannel = 3
    If (OptionChDR) Then redChannel = 4
    
    'get green channel
    If (OptionCh1G) Then greenChannel = 0
    If (OptionCh2G) Then greenChannel = 1
    If (OptionCh3G) Then greenChannel = 2
    If (OptionChMG) Then greenChannel = 3
    If (OptionChDG) Then greenChannel = 4
    
    'get blue channel
    If (OptionCh1B) Then blueChannel = 0
    If (OptionCh2B) Then blueChannel = 1
    If (OptionCh3B) Then blueChannel = 2
    If (OptionChMB) Then blueChannel = 3
    If (OptionChDB) Then blueChannel = 4
    
'    MsgBox ("red=" + CStr(redChannel) + " green=" + CStr(greenChannel) + " blue=" + CStr(blueChannel))
              
    'Capture time points
    For TimeIndex = 0 To TimePointsText.Value - 1
       TimePointStartTime = Now()
       
        'Capture rows of specimens
        For SpecimenRowIndex = 0 To RowsOfSpecimensText - 1
            'Capture columns of specimens
            For SpecimenColumnIndex = 0 To ColumnsOfSpecimensText - 1
                'Capture y tiles per specimen
                For YTilesIndex = 0 To YTilesPerSpecimenText - 1
                    'Capture x tiles per specimen
                    For XTilesIndex = 0 To XTilesPerSpecimenText - 1
                        Dim Recording As DsRecording
                        Set Recording = Lsm5.DsRecording
                        Recording.TimeSeries = False
                        
                        'Move stage
                        If (OptionUL) Then
                            Lsm5.Hardware.CpStages.PositionY = startX + XTilesIndex * FOV + SpecimenColumnIndex * CDbl(DistanceBetweenColumnsText) + CDbl(XOffsetText)
                            Lsm5.Hardware.CpStages.PositionX = startY - YTilesIndex * FOV - SpecimenRowIndex * CDbl(DistanceBetweenRowsText) - CDbl(YOffsetText)
                        End If

                        If (OptionUR) Then
                            Lsm5.Hardware.CpStages.PositionX = startX - XTilesIndex * FOV - SpecimenColumnIndex * CInt(DistanceBetweenColumnsText) - CInt(XOffsetText)
                            Lsm5.Hardware.CpStages.PositionY = startY + YTilesIndex * FOV + SpecimenRowIndex * CInt(DistanceBetweenRowsText) + CInt(YOffsetText)
                        End If
 
                        If (OptionLL) Then
                             Lsm5.Hardware.CpStages.PositionY = startX - XTilesIndex * FOV - SpecimenColumnIndex * CDbl(DistanceBetweenColumnsText) - CDbl(XOffsetText)
                             Lsm5.Hardware.CpStages.PositionX = startY + YTilesIndex * FOV + SpecimenRowIndex * CDbl(DistanceBetweenRowsText) + CDbl(YOffsetText)
                        End If
 
                        If (OptionLR) Then
                            Lsm5.Hardware.CpStages.PositionX = startX + XTilesIndex * FOV + SpecimenColumnIndex * CInt(DistanceBetweenColumnsText) + CInt(XOffsetText)
                            Lsm5.Hardware.CpStages.PositionY = startY - YTilesIndex * FOV - SpecimenRowIndex * CInt(DistanceBetweenRowsText) - CInt(YOffsetText)
                        End If
                        
                        'Wait till stage is finished moving
                        While Lsm5.Hardware.CpStages.IsBusy()
                            Sleep (100)
                        Wend
                        
                        'Capture z-stack
                        Dim RecordingDoc As DsRecordingDoc
                        Set RecordingDoc = Lsm5.StartScan()
                        While RecordingDoc.IsBusy()
                            DoEvents
                            Sleep 200
                        Wend
                                              
                        'Determine size of field of view in microns
                        If (FOV = 0) Then
                            FOV = RecordingDoc.VoxelSizeX * RecordingDoc.GetDimensionX() * 1000000 * (100 - CInt(PercentOverlapText)) / 100
    '                        For channel = 5 To 0 Step -1
    '                            If RecordingDoc.ChannelColor(channel) = 255 Then redChannel = channel
    '                            If RecordingDoc.ChannelColor(channel) = 65280 Then greenChannel = channel
    '                            If RecordingDoc.ChannelColor(channel) = 16711680 Then blueChannel = channel
    '                        Next
                        End If
                        
                        'Set strFilename so can export next round
                        'Export z-stack in format "prefix-cCCrRRyYYxXXtTTTTzZZZ
                        strFilename = PathOfFolderForImagesText _
                          + FilenamePrefixText _
                          + "-c" + Format(SpecimenColumnIndex, "00") _
                          + "-r" + Format(SpecimenRowIndex, "00") _
                          + "-y" + Format(YTilesIndex, "00") _
                          + "-x" + Format(XTilesIndex, "00") _
                          + "-t" + Format(TimeIndex, "0000") _
                          + "-z"
                        
                        'Export images
   '                     MsgBox ("About to export")
                        'JpegMedium gives same size file as JpegAccurate. This might be bug in API.
                        'I can't tell difference between JpegPoor and JpegMedium in quality but it is 5x smaller so using Poor
                        'You must use a 3 channel image
                        Dim nExportType As Integer
                        If JpegHighCompressionButton.Value Then
                            strFileExtension = ".jpg"
                            nExportType = eExportJpegPoor
                        ElseIf JpegLowCompressionButton.Value Then
                            strFileExtension = ".jpg"
                            nExportType = eExportJpegAccurate
                        ElseIf TiffButton.Value Then
                            strFileExtension = ".tif"
                            nExportType = eExportTiff
                        ElseIf Tiff12Button.Value Then
                            strFileExtension = ".tif"
                            nExportType = eExportTiff12Bit
                        ElseIf LSM4Button.Value Then
                            strFileExtension = ".lsm"
                            nExportType = eExportLsm4Chunky
                        End If
                        
                            success = RecordingDoc.Export(nExportType, strFilename + strFileExtension, True, False, 0, 0, False, redChannel, greenChannel, blueChannel)
                        
                        DoEvents
                        
                        'Write a line in .mgc file for each image in z-series
                        sTab = Chr(9) 'tab
                        For z = 0 To RecordingDoc.GetDimensionZ - 1
                            Dim Pinhole1 As Double
                            Lsm5.Hardware.CpPinholes.Select (1)
                            Pinhole1 = Lsm5.Hardware.CpPinholes.Diameter
                            Dim Pinhole2 As Double
                            Lsm5.Hardware.CpPinholes.Select (2)
                            Pinhole2 = Lsm5.Hardware.CpPinholes.Diameter
                            Dim Pinhole3 As Double
                            Lsm5.Hardware.CpPinholes.Select (3)
                            Pinhole3 = Lsm5.Hardware.CpPinholes.Diameter
                                                 
                            Dim Gain1 As Double
                            Lsm5.Hardware.CpPmts.Select (5)
                            Gain1 = Lsm5.Hardware.CpPmts.Gain
                            Dim Gain2 As Double
                            Lsm5.Hardware.CpPmts.Select (6)
                            Gain2 = Lsm5.Hardware.CpPmts.Gain
                            Dim Gain3 As Double
                            Lsm5.Hardware.CpPmts.Select (7)
                            Gain3 = Lsm5.Hardware.CpPmts.Gain

                            Dim Attenuation488 As Double
                            Attenuation488 = Lsm5.Hardware.CpLaserLines.Attenuation(488)
                            Dim Attenuation543 As Double
                            Attenuation543 = Lsm5.Hardware.CpLaserLines.Attenuation(543)
                            Dim Attenuation820 As Double
                            Attenuation820 = Lsm5.Hardware.CpLaserLines.Attenuation(820)
                            
                            Print #intOutFile, strFilename + Format(z, "000") + strFileExtension + sTab + _
                              CStr(SpecimenRowIndex) + sTab + _
                              CStr(SpecimenColumnIndex) + sTab + _
                              CStr(XTilesIndex) + sTab + _
                              CStr(YTilesIndex) + sTab + _
                              CStr(z) + sTab + _
                              CStr(TimeIndex) + sTab + _
                              XOffsetText + sTab + _
                              YOffsetText + sTab + _
                              ZOffsetText + sTab + _
                              CStr(Now()) + sTab + _
                              CStr(Pinhole1) + sTab + _
                              CStr(Pinhole2) + sTab + _
                              CStr(Pinhole3) + sTab + _
                              CStr(CLng(Gain1)) + sTab + _
                              CStr(CLng(Gain2)) + sTab + _
                              CStr(CLng(Gain3)) + sTab + _
                              CStr(CDbl(CInt(Attenuation488 * 10) / 10)) + sTab + _
                              CStr(CDbl(CInt(Attenuation543 * 10) / 10)) + sTab + _
                              CStr(CDbl(CInt(Attenuation820 * 10) / 10))
                              'date/time saved is not quite right since it is the date/time of saving rather than capture
                        Next z
                        
                        
                        'free up memory?
                        RecordingDoc.CloseAllWindows
                        Set RecordingDoc = Nothing
                        Set Recording = Nothing
                        
                        'exit if stop button has been clicked
                        If Not doImage Then GoTo EndLabel
    
                                                
                    Next XTilesIndex
                Next YTilesIndex
            Next SpecimenColumnIndex
        Next SpecimenRowIndex
        
       'Update percent complete box
       PercentCompleteText = CInt(CDbl(TimeIndex + 1) * (CDbl(100) / CDbl(TimePointsText.Value)))
        
        'Wait until time interval has expired before looping
        Do While DateDiff("s", TimePointStartTime, Now()) < CInt(TimeIntervalText) And TimeIndex <> (CInt(TimePointsText.Value) - 1)
            TimeUntilScanText = CInt(TimeIntervalText) - DateDiff("s", TimePointStartTime, Now())
            DoEvents
            Sleep 500
        Loop
   
    Next TimeIndex
    
EndLabel:
    
    'Move to original XY stage position
    Lsm5.Hardware.CpStages.PositionX = startX
    Lsm5.Hardware.CpStages.PositionY = startY
    
    Close #intOutFile
    
    Set Track = Nothing
    Set Laser = Nothing
    Set DetectionChannel = Nothing
    Set IlluminationChannel = Nothing
    Set DataChannel = Nothing
    Set BeamSplitter = Nothing
    Set Timers = Nothing
    Set Markers = Nothing
End Sub
Private Sub BrowseButton_Click()
    Dim FolderName As String
    FolderName = GetFolderName("Select a folder")
    If FolderName <> "" Then PathOfFolderForImagesText = FolderName
End Sub

Private Sub ColumnsOfSpecimensSpin_Change()
    ColumnsOfSpecimensText = ColumnsOfSpecimensSpin
End Sub

Private Sub ColumnsOfSpecimensText_Change()
    ColumnsOfSpecimensSpin.Value = Min(ColumnsOfSpecimensSpin.Max, Max(ColumnsOfSpecimens, ColumnsOfSpecimensSpin.Min))
    ColumnsOfSpecimensText = ColumnsOfSpecimensSpin
    SetEstCaptureTimePerInterval
    SetTotalNumberOfImages
    SetEstTotalDiskSpace
End Sub

Private Sub DistanceBetweenColumnsSpin_Change()
    DistanceBetweenColumnsText = DistanceBetweenColumnsSpin
End Sub

Private Sub DistanceBetweenColumnsText_Change()
    DistanceBetweenColumnsSpin.Value = Min(DistanceBetweenColumnsSpin.Max, Max(DistanceBetweenColumns, DistanceBetweenColumnsSpin.Min))
    DistanceBetweenColumnsText = DistanceBetweenColumnsSpin

End Sub

Private Sub DistanceBetweenRowsSpin_Change()
    DistanceBetweenRowsText = DistanceBetweenRowsSpin
End Sub

Private Sub DistanceBetweenRowsText_Change()
    DistanceBetweenRowsSpin.Value = Min(DistanceBetweenRowsSpin.Max, Max(DistanceBetweenRows, DistanceBetweenRowsSpin.Min))
    DistanceBetweenRowsText = DistanceBetweenRowsSpin

End Sub

Private Sub HelpButton_Click()
    HelpForm.Show
End Sub

Private Sub PercentOverlapSpin_Change()
    PercentOverlapText = PercentOverlapSpin
End Sub

Private Sub PercentOverlapText_Change()
    PercentOverlapSpin.Value = Min(PercentOverlapSpin.Max, Max(CInt(PercentOverlapText), PercentOverlapSpin.Min))
    PercentOverlapText = PercentOverlapSpin

End Sub

Private Sub RowsOfSpecimensSpin_Change()
    RowsOfSpecimensText = RowsOfSpecimensSpin
End Sub

Private Sub RowsOfSpecimensText_Change()
    RowsOfSpecimensSpin.Value = Min(RowsOfSpecimensSpin.Max, Max(RowsOfSpecimens, RowsOfSpecimensSpin.Min))
    RowsOfSpecimensText = RowsOfSpecimensSpin
    SetEstCaptureTimePerInterval
    SetTotalNumberOfImages
    SetEstTotalDiskSpace
End Sub

Private Sub StartButton_Click()
    doImage = True
    StartMegaCapture
End Sub

Private Sub StopButton_Click()
    doImage = False
End Sub
Private Sub SetEstCaptureTimePerInterval()
    EstCaptureTimePerIntervalText = CInt(XTilesPerSpecimen * YTilesPerSpecimen _
      * ColumnsOfSpecimens * RowsOfSpecimens _
      * DimZ _
      * (Lsm5.Hardware.CpScancontrol.TotalTimePerFrame() + 0.17))
    
    If CInt(EstCaptureTimePerIntervalText + 5) > TimeInterval Then
        EstCaptureTimePerIntervalText.BackColor = RGB(255, 0, 0)
    Else
        EstCaptureTimePerIntervalText.BackColor = &H8000000F
    End If
End Sub

Private Sub TiffButton_Click()

End Sub

Private Sub TimeIntervalSpin_Change()
    TimeIntervalText = TimeIntervalSpin
End Sub

Private Sub TimeIntervalText_Change()
    TimeIntervalSpin.Value = Min(TimeIntervalSpin.Max, Max(TimeInterval, TimeIntervalSpin.Min))
    TimeIntervalText = TimeIntervalSpin
    SetEstCaptureTimePerInterval
    SetTotalTime
End Sub

Private Sub TimePointsSpin_Change()
    TimePointsText = TimePointsSpin
End Sub

Private Sub TimePointsText_Change()
    TimePointsSpin.Value = Min(TimePointsSpin.Max, Max(TimePoints, TimePointsSpin.Min))
    TimePointsText = TimePointsSpin
    SetTotalTime
    SetTotalNumberOfImages
    SetEstTotalDiskSpace
End Sub

Private Sub TimeUntilScanText_Change()

End Sub

Private Sub UserForm_Activate()
    displayedWarning = False
    doImage = True
    SetTotalTime
    SetTotalNumberOfImages
    SetEstTotalDiskSpace
    SetEstCaptureTimePerInterval
End Sub

Private Sub XOffsetSpin_Change()
    XOffsetText = XOffsetSpin
End Sub

Private Sub XOffsetText_Change()
    XOffsetSpin.Value = Min(XOffsetSpin.Max, Max(XOffset, XOffsetSpin.Min))
    XOffsetText = XOffsetSpin

End Sub

Private Sub XTilesPerSpecimenSpin_Change()
    XTilesPerSpecimenText = XTilesPerSpecimenSpin
End Sub

Private Sub XTilesPerSpecimenText_Change()
    XTilesPerSpecimenSpin.Value = Min(XTilesPerSpecimenSpin.Max, Max(XTilesPerSpecimen, XTilesPerSpecimenSpin.Min))
    XTilesPerSpecimenText = XTilesPerSpecimenSpin
    SetEstCaptureTimePerInterval
    SetTotalNumberOfImages
    SetEstTotalDiskSpace
End Sub

Private Sub YOffsetSpin_Change()
    YOffsetText = YOffsetSpin
End Sub

Private Sub YOffsetText_Change()
    YOffsetSpin.Value = Min(YOffsetSpin.Max, Max(YOffset, YOffsetSpin.Min))
    YOffsetText = YOffsetSpin

End Sub

Private Sub YTilesPerSpecimenSpin_Change()
    YTilesPerSpecimenText = YTilesPerSpecimenSpin
End Sub

Private Sub YTilesPerSpecimenText_Change()
    YTilesPerSpecimenSpin.Value = Min(YTilesPerSpecimenSpin.Max, Max(YTilesPerSpecimen, YTilesPerSpecimenSpin.Min))
    YTilesPerSpecimenText = YTilesPerSpecimenSpin
    SetEstCaptureTimePerInterval
    SetTotalNumberOfImages
    SetEstTotalDiskSpace
End Sub
Private Sub SetTotalNumberOfImages()
    TotalNumberOfImagesText = XTilesPerSpecimen * YTilesPerSpecimen _
      * ColumnsOfSpecimens * RowsOfSpecimens _
      * DimZ _
      * TimePoints
End Sub
Private Sub SetEstTotalDiskSpace()
    EstTotalDiskSpaceText = XTilesPerSpecimen * YTilesPerSpecimen _
      * ColumnsOfSpecimens * RowsOfSpecimens _
      * DimZ _
      * TimePoints _
      * 0.1 'mb / image
End Sub
Private Sub SetTotalTime()
    TotalTimeText = CLng(TimePoints * TimeInterval / 360) / 10
End Sub


Function GetFolderName(Msg As String) As String
' returns the name of the folder selected by the user

Dim bInfo As BROWSEINFO, path As String, r As Long
Dim X As Long, pos As Integer
    bInfo.pidlRoot = 0& ' Root folder = Desktop
    If IsMissing(Msg) Then
        bInfo.lpszTitle = "Select a folder."
        ' the dialog title
    Else
        bInfo.lpszTitle = Msg ' the dialog title
    End If
    bInfo.ulFlags = &H1 ' Type of directory to return
    X = SHBrowseForFolder(bInfo) ' display the dialog
    ' Parse the result
    path = Space$(512)
    r = SHGetPathFromIDList(ByVal X, ByVal path)
    If r Then
        pos = InStr(path, Chr$(0))
        GetFolderName = Left(path, pos - 1) + "\"
    Else
        GetFolderName = ""
    End If
End Function
Function XTilesPerSpecimen() As Long
    If IsNumeric(XTilesPerSpecimenText) Then
        XTilesPerSpecimen = CLng(XTilesPerSpecimenText)
    Else
        XTilesPerSpecimen = 0
    End If
End Function
Function YTilesPerSpecimen() As Long
    If IsNumeric(YTilesPerSpecimenText) Then
        YTilesPerSpecimen = CLng(YTilesPerSpecimenText)
    Else
        YTilesPerSpecimen = 0
    End If
End Function
Function ColumnsOfSpecimens() As Long
    If IsNumeric(ColumnsOfSpecimensText) Then
        ColumnsOfSpecimens = CLng(ColumnsOfSpecimensText)
    Else
        ColumnsOfSpecimens = 0
    End If
End Function
Function RowsOfSpecimens() As Long
    If IsNumeric(RowsOfSpecimensText) Then
        RowsOfSpecimens = CLng(RowsOfSpecimensText)
    Else
        RowsOfSpecimens = 0
    End If
End Function
Function TimePoints() As Long
    If IsNumeric(TimePointsText) Then
        TimePoints = CLng(TimePointsText)
    Else
        TimePoints = 0
    End If
End Function
Function TimeInterval() As Long
    If IsNumeric(TimeIntervalText) Then
        TimeInterval = CLng(TimeIntervalText)
    Else
        TimeInterval = 0
    End If
End Function
Function DistanceBetweenColumns() As Long
    If IsNumeric(DistanceBetweenColumnsText) Then
        DistanceBetweenColumns = CLng(DistanceBetweenColumnsText)
    Else
        DistanceBetweenColumns = 0
    End If
End Function
Function DistanceBetweenRows() As Long
    If IsNumeric(DistanceBetweenRowsText) Then
        DistanceBetweenRows = CLng(DistanceBetweenRowsText)
    Else
        DistanceBetweenRows = 0
    End If
End Function
Function XOffset() As Long
    If IsNumeric(XOffsetText) Then
        XOffset = CLng(XOffsetText)
    Else
        XOffset = 0
    End If
End Function
Function YOffset() As Long
    If IsNumeric(YOffsetText) Then
        YOffset = CLng(YOffsetText)
    Else
        YOffset = 0
    End If
End Function
Function ZOffset() As Long
    If IsNumeric(ZOffsetText) Then
        ZOffset = CLng(ZOffsetText)
    Else
        ZOffset = 0
    End If
End Function

Function ZDim() As Long
    Dim success As Integer

    Dim RecordingDoc As RecordingDocument
    Set RecordingDoc = Lsm5.DsRecordingDocObject(0, success)
    'the success variable does not seem to work properly

    On Error GoTo ErrHandler
    ZDim = RecordingDoc.GetDimensionZ
    GoTo EndLabel

ErrHandler:
    If Not displayedWarning Then
        MsgBox ("For totals to be calculated correctly, you must leave open a sample image taken with the current settings.")
        displayedWarning = True
    End If
    ZDim = 1

EndLabel:

End Function

Function VoxelSizeX() As Double
    Dim success As Integer
    Set RecordingDoc = Lsm5.DsRecordingDocObject(0, success)
    'the success variable does not seem to work properly
    
    On Error GoTo ErrHandler
    VoxelSizeX = CLng(RecordingDoc.VoxelSizeX * 100000000) / 100
    GoTo EndLabel
    
ErrHandler:
    If Not displayedWarning Then
        MsgBox ("For totals to be calculated correctly, you must leave open a sample image taken with the current settings.")
        displayedWarning = True
    End If
    VoxelSizeX = 1
    
EndLabel:

End Function
Function VoxelSizeY() As Double
    Dim success As Integer
    Set RecordingDoc = Lsm5.DsRecordingDocObject(0, success)
    'the success variable does not seem to work properly
    
    On Error GoTo ErrHandler
    VoxelSizeY = CLng(RecordingDoc.VoxelSizeY * 100000000) / 100
    GoTo EndLabel
    
ErrHandler:
    If Not displayedWarning Then
        MsgBox ("For totals to be calculated correctly, you must leave open a sample image taken with the current settings.")
        displayedWarning = True
    End If
    VoxelSizeY = 1
    
EndLabel:

End Function
Function VoxelSizeZ() As Double
    Dim success As Integer
    Set RecordingDoc = Lsm5.DsRecordingDocObject(0, success)
    'the success variable does not seem to work properly
    
    On Error GoTo ErrHandler
    VoxelSizeZ = CLng(RecordingDoc.VoxelSizeZ * 100000000) / 100
    GoTo EndLabel
    
ErrHandler:
    If Not displayedWarning Then
        MsgBox ("For totals to be calculated correctly, you must leave open a sample image taken with the current settings.")
        displayedWarning = True
    End If
    VoxelSizeZ = 1
    
EndLabel:

End Function
Function Max(var1 As Long, var2 As Long) As Long
    If var1 > var2 Then
        Max = var1
    Else
        Max = var2
    End If
End Function
Function Min(var1 As Long, var2 As Long) As Long
    If var1 < var2 Then
        Min = var1
    Else
        Min = var2
    End If
End Function

Private Sub ZOffsetSpin_Change()
    ZOffsetText = ZOffsetSpin
End Sub

Private Sub ZOffsetText_Change()
    ZOffsetSpin.Value = Min(ZOffsetSpin.Max, Max(ZOffset, ZOffsetSpin.Min))
    ZOffsetText = ZOffsetSpin

End Sub
                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                          
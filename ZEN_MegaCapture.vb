
'global
Dim displayedWarning As Boolean
Dim doImage As Boolean
'Dim MarkAndFind As Boolean
Dim TimeIndex As Integer
Dim SpecimenRowIndex As Integer
Dim SpecimenColumnIndex As Integer
Dim SpecimenPositionIndex As Integer 'Paul added this
Dim PositionsOfSpecimens As Long 'Long, which I think was Sean's idea or I took because ColumnsOfSpecimens was Long
Dim XTilesIndex As Integer
Dim YTilesIndex As Integer
Dim intOutFile As Integer
Dim strOutFile As String
Dim Success As Integer
Dim strFilename As String
Dim redChannel As Integer
Dim greenChannel As Integer
Dim blueChannel As Integer
Dim Recording As DsRecording
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
    
    Set Recording = Lsm5.DsRecording
    Recording.TimeSeries = False
    
    'Create .meg file to save parameters for all images for importing to GoFigure
    intOutFile = FreeFile
    strOutFile = PathOfFolderForImagesText + FilenamePrefixText + ".meg"
    Close #intOutFile
    Open strOutFile For Output As #intOutFile
    Dim sTab As String
    sTab = Chr(9) 'tab
    
    Dim finishedHeader As Boolean
    finishedHeader = False
    
    'Enable nudging buttons
    XOffsetSpin.Enabled = True
    YOffsetSpin.Enabled = True
    ZOffsetSpin.Enabled = True 'TODO implement the z-offset
    
    'Remember starting stage position
    startX = Lsm5.Hardware.CpStages.PositionX
    startY = Lsm5.Hardware.CpStages.PositionY
    Dim strMessage As String
    strMessage = "startX=" + CStr(startX) + "   startY=" + CStr(startY)
'    MsgBox (strMessage)
  
    Dim stageIndex As Long
    strFilename = ""
    FOV = 0
            
    'Capture time points
    For TimeIndex = 0 To TimePointsText.Value - 1
       TimePointStartTime = Now()
       
        If MarkAndFind Then
        
            Dim MyXpos() As Double
            Dim MyYpos() As Double
            Dim MyZpos() As Double
            PositionsOfSpecimens = GetMarkedLocations(MyXpos(), MyYpos(), MyZpos()) 'Do I need xPos() instead?  I think I did
            For SpecimenPositionIndex = 0 To PositionsOfSpecimens - 1
                'Even though AcquireTiledZStack is a Sub in your example code, VB doesn 't differentiate between the calling syntax of a Sub
                'and a Function. Regardless of whether the procedure you're calling is a Sub or a Function, if you use parentheses and don't
                'use the Call keyword, VB expects there to be a return value
                'Recording.FocusPosABC1 = 1.03
                'Recording.FocusPosABC2 = 1.03
                Call AcquireTiledZStack(MyXpos(SpecimenPositionIndex), MyYpos(SpecimenPositionIndex), MyZpos(SpecimenPositionIndex))
                'Recording.FocusPosABC1 = 0.03
                'Recording.FocusPosABC2 = 0.03
                'For xtile,ytile,z
                'exit if stop button has been clicked
                'this doesn't work inside this subroutine
                If Not doImage Then GoTo EndLabel
                'Do I want this in here?  I might want Stop to only stop at the end of a set of tiles
            Next SpecimenPositionIndex 'Do the for loop with position index increasing
            
        Else
            Dim yInput As Double
            Dim xInput As Double
            'Here is where I would also Dim zInput As Double
            If (OptionUL) Then
                yInput = startX + SpecimenColumnIndex * CDbl(DistanceBetweenColumnsText) + CDbl(XOffsetText)
                xInput = startY - SpecimenRowIndex * CDbl(DistanceBetweenRowsText) - CDbl(YOffsetText)
            End If

            If (OptionUR) Then
                xInput = startX - SpecimenColumnIndex * CInt(DistanceBetweenColumnsText) - CInt(XOffsetText)
                yInput = startY + SpecimenRowIndex * CInt(DistanceBetweenRowsText) + CInt(YOffsetText)
            End If

            If (OptionLL) Then
                 yInput = startX - SpecimenColumnIndex * CDbl(DistanceBetweenColumnsText) - CDbl(XOffsetText)
                 xInput = startY + SpecimenRowIndex * CDbl(DistanceBetweenRowsText) + CDbl(YOffsetText)
            End If

            If (OptionLR) Then
                xInput = startX + SpecimenColumnIndex * CInt(DistanceBetweenColumnsText) + CInt(XOffsetText)
                yInput = startY - SpecimenRowIndex * CInt(DistanceBetweenRowsText) - CInt(YOffsetText)
            End If
            
            'Do the for loop with rows and columns
               'Capture rows of specimens
            For SpecimenRowIndex = 0 To RowsOfSpecimensText - 1
                'Capture columns of specimens
                For SpecimenColumnIndex = 0 To ColumnsOfSpecimensText - 1
                    Call AcquireTiledZStack(xInput, yInput, yInput) 'Just put yInput in again so I would have a double
                    'exit if stop button has been clicked
                    'this doesn't work inside this subroutine
                    If Not doImage Then GoTo EndLabel
                    'Do I want this in here?  I might want Stop to only stop at the end of a set of tiles
                Next SpecimenColumnIndex
            Next SpecimenRowIndex
        End If
       
        
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
    
    'TODO move png conversion code here
    'TODO getting grayscale tifs
    
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

Private Sub EstCaptureTimePerIntervalText_Change()

End Sub

Private Sub HelpButton_Click()
    HelpForm.Show
End Sub



Private Sub MarkAndFind_Click()
    
    If MarkAndFind Then
    
        RowsOfSpecimensText.Enabled = False
        RowsOfSpecimensText.BackColor = &H80000013
        RowsOfSpecimensSpin.Enabled = False
    
        ColumnsOfSpecimensText.Enabled = False
        ColumnsOfSpecimensText.BackColor = &H80000013
        ColumnsOfSpecimensSpin.Enabled = False
    
        DistanceBetweenRowsText.Enabled = False
        DistanceBetweenRowsText.BackColor = &H80000013
        DistanceBetweenRowsSpin.Enabled = False
    
        DistanceBetweenColumnsText.Enabled = False
        DistanceBetweenColumnsText.BackColor = &H80000013
        DistanceBetweenColumnsSpin.Enabled = False
        
    Else
        
        RowsOfSpecimensText.Enabled = True
        RowsOfSpecimensText.BackColor = &H8000000F
        RowsOfSpecimensSpin.Enabled = True
    
        ColumnsOfSpecimensText.Enabled = True
        ColumnsOfSpecimensText.BackColor = &H8000000F
        ColumnsOfSpecimensSpin.Enabled = True
    
        DistanceBetweenRowsText.Enabled = True
        DistanceBetweenRowsText.BackColor = &H8000000F
        DistanceBetweenRowsSpin.Enabled = True
    
        DistanceBetweenColumnsText.Enabled = True
        DistanceBetweenColumnsText.BackColor = &H8000000F
        DistanceBetweenColumnsSpin.Enabled = True
        
    End If
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
        TimePoints = 0#
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
    Dim Success As Integer

    Dim RecordingDoc As RecordingDocument
    Set RecordingDoc = Lsm5.DsRecordingDocObject(0, Success)
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
    Dim Success As Integer
    Set RecordingDoc = Lsm5.DsRecordingDocObject(0, Success)
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
    Dim Success As Integer
    Set RecordingDoc = Lsm5.DsRecordingDocObject(0, Success)
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
    Dim Success As Integer
    Set RecordingDoc = Lsm5.DsRecordingDocObject(0, Success)
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

'Sub because I don't want it to return a value.  Will need to use Call when calling this Sub
Public Sub AcquireTiledZStack(xPos As Double, yPos As Double, zPos As Double)
'Capture y tiles per specimen
    For YTilesIndex = 0 To YTilesPerSpecimenText - 1
        'Capture x tiles per specimen
        For XTilesIndex = 0 To XTilesPerSpecimenText - 1
            
            
            'Move stage
            If (OptionUL) Then
                Lsm5.Hardware.CpStages.PositionY = xPos + XTilesIndex * FOV
                Lsm5.Hardware.CpStages.PositionX = yPos - YTilesIndex * FOV
            End If

            If (OptionUR) Then
                Lsm5.Hardware.CpStages.PositionX = xPos - XTilesIndex * FOV
                Lsm5.Hardware.CpStages.PositionY = yPos + YTilesIndex * FOV
            End If

            If (OptionLL) Then
                 Lsm5.Hardware.CpStages.PositionY = xPos - XTilesIndex * FOV
                 Lsm5.Hardware.CpStages.PositionX = yPos + YTilesIndex * FOV
            End If

            If (OptionLR) Then
                Lsm5.Hardware.CpStages.PositionX = xPos + XTilesIndex * FOV
                Lsm5.Hardware.CpStages.PositionY = yPos - YTilesIndex * FOV
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
                FOV = RecordingDoc.VoxelSizeX() * RecordingDoc.GetDimensionX() * 1000000 * (100 - CInt(PercentOverlapText)) / 100
            End If
                                    
            Dim channel As Integer
            'Added this to see if I can get channel name
            'Set DataChannel = Track.DataChannelObjectByChannel(10, Success)
            If Not finishedHeader Then
                'Write header for .meg
                Print #intOutFile, "MegaCapture"
                Print #intOutFile, "<ImageSessionData>"
                Print #intOutFile, "Version" + sTab + "3.0"
                Print #intOutFile, "ExperimentTitle" + sTab + ExperimentTitleText
                Print #intOutFile, "ExperimentDescription" + sTab + ExperimentDescriptionText
                Print #intOutFile, "TimeInterval" + sTab + TimeIntervalText
                Print #intOutFile, "Objective" + sTab + CStr(Lsm5.Hardware.CpObjectiveRevolver.Summary(1))
                Print #intOutFile, "VoxelSizeX" + sTab + CStr((RecordingDoc.VoxelSizeX * 10 ^ 6)) 'changed to microns by Paul
                Print #intOutFile, "VoxelSizeY" + sTab + CStr((RecordingDoc.VoxelSizeY * 10 ^ 6))
                Print #intOutFile, "VoxelSizeZ" + sTab + CStr((RecordingDoc.VoxelSizeZ * 10 ^ 6))
                Print #intOutFile, "DimensionX" + sTab + CStr(RecordingDoc.GetDimensionX)
                Print #intOutFile, "DimensionY" + sTab + CStr(RecordingDoc.GetDimensionY)
                Print #intOutFile, "DimensionPL" + sTab + "1"
                If MarkAndFind Then
                    Print #intOutFile, "DimensionCO" + sTab + Format(PositionsOfSpecimens, "00") 'Could also consider "0"
                    Print #intOutFile, "DimensionRO" + sTab + "1"
                Else
                    Print #intOutFile, "DimensionCO" + sTab + ColumnsOfSpecimensText
                    Print #intOutFile, "DimensionRO" + sTab + RowsOfSpecimensText
                End If
                Print #intOutFile, "DimensionZT" + sTab + "1"
                Print #intOutFile, "DimensionYT" + sTab + YTilesPerSpecimenText
                Print #intOutFile, "DimensionXT" + sTab + XTilesPerSpecimenText
                Print #intOutFile, "DimensionTM" + sTab + TimePointsText
                Print #intOutFile, "DimensionZS" + sTab + CStr(RecordingDoc.GetDimensionZ)
                Print #intOutFile, "DimensionCH" + sTab + CStr(RecordingDoc.GetDimensionChannels)
                For channel = 0 To RecordingDoc.GetDimensionChannels - 1
                    Print #intOutFile, "ChannelColor" + Format(channel, "00") + sTab + CStr(RecordingDoc.ChannelColor(channel))
                    ' TODO record channel name.  Not working yet
                    'Print #intOutFile, "ChannelName" + DataChannel.DyeName
                Next channel
                
                Dim strDepth, strFileType As String
                If OptionPNG8.Value Then
                    strDepth = "8"
                    strFileType = "PNG"
                ElseIf OptionPNG12.Value Then
                    strDepth = "12"
                    strFileType = "PNG"
                ElseIf OptionTiff8.Value Then
                    strDepth = "8"
                    strFileType = "TIF"
                ElseIf OptionTiff12.Value Then
                    strDepth = "12"
                    strFileType = "TIF"
                End If
                
                Print #intOutFile, "ChannelDepth" + sTab + strDepth
                Print #intOutFile, "FileType" + sTab + strFileType
                Print #intOutFile, "</ImageSessionData>"

                finishedHeader = True
            End If
            
            'Set strFilename so can export next round
            'Export z-stack in format "prefix-pPPPcCCrRRyYYxXXtTTTTzZZZ
            'p is for plate number but can't switch plates on cyclops
            If MarkAndFind Then
                strFilename = PathOfFolderForImagesText _
                  + FilenamePrefixText _
                  + "-PL00" _
                  + "-CO" + Format(SpecimenPositionIndex, "00") _
                  + "-RO" + Format(0, "00") _
                  + "-ZT00" _
                  + "-YT" + Format(YTilesIndex, "00") _
                  + "-XT" + Format(XTilesIndex, "00") _
                  + "-TM" + Format(TimeIndex, "0000")
            Else
                strFilename = PathOfFolderForImagesText _
                  + FilenamePrefixText _
                  + "-PL00" _
                  + "-CO" + Format(SpecimenColumnIndex, "00") _
                  + "-RO" + Format(SpecimenRowIndex, "00") _
                  + "-ZT00" _
                  + "-YT" + Format(YTilesIndex, "00") _
                  + "-XT" + Format(XTilesIndex, "00") _
                  + "-TM" + Format(TimeIndex, "0000")
            End If
            
            'Export images
            Dim strName, strExtension As String
            Dim nExportType As Integer
            If OptionPNG8.Value Then
                nExportType = eExportTiff
                strExtension = ".tif"
            ElseIf OptionPNG12.Value Then
                nExportType = eExportTiff12Bit
                strExtension = ".tif"
            ElseIf OptionTiff8.Value Then
                nExportType = eExportTiff
                strExtension = ".png" 'bug? should be tif
            ElseIf OptionTiff12.Value Then
                nExportType = eExportTiff12Bit
                strExtension = ".png" 'bug? should be tif
            End If
            
            'export as tifs first then convert if png
            Dim strZ As String
            For channel = 0 To RecordingDoc.GetDimensionChannels - 1
                strName = strFilename + "-CH" + Format(channel, "00") + "-ZS"
                If RecordingDoc.GetDimensionZ = 1 Then
                   strZ = "0000"
                ElseIf RecordingDoc.GetDimensionZ < 10 Then
                   strZ = "000"
                ElseIf RecordingDoc.GetDimensionZ < 100 Then
                   strZ = "00"
                ElseIf RecordingDoc.GetDimensionZ < 1000 Then
                   strZ = "0"
                Else
                   strZ = ""
                End If
                Success = RecordingDoc.Export(nExportType, strName + strZ + ".tif", True, False, 0, 0, True, channel, channel, channel)
            Next channel
            
            If (OptionPNG8.Value Or OptionPNG12.Value) Then
                ' convert tifs to pngs using ImageMagick
                While RecordingDoc.IsBusy()
                    DoEvents
                    Sleep 200
                Wend
                'Sleep 5000  'must wait to be sure done saving
                For z = 0 To RecordingDoc.GetDimensionZ - 1
                    For channel = 0 To RecordingDoc.GetDimensionChannels - 1
                        strName = strFilename + "-ch" + Format(channel, "00") + "-zs" + Format(z, "0000")
                        Shell ("convertmagick " + strName + ".tif " + strName + ".png")
                    Next channel
                Next z

                ' delete tifs
                Sleep 5000 'must wait for conversion
                For z = 0 To RecordingDoc.GetDimensionZ - 1
                    For channel = 0 To RecordingDoc.GetDimensionChannels - 1
                        strName = strFilename + "-ch" + Format(channel, "00") + "-zs" + Format(z, "0000")
                        FileSystem.Kill (strName + ".tif")
                    Next channel
                Next z
            End If
            
            DoEvents
            
            'Write a line in .meg file for each image in z-series
            sTab = Chr(9) 'tab
            Dim Pinhole As Double
            Lsm5.Hardware.CpPinholes.Select (1)
            For z = 0 To RecordingDoc.GetDimensionZ - 1
                For channel = 0 To RecordingDoc.GetDimensionChannels - 1
'                               recordingdoc.
' TODO should also record(master gain, digital offset and digital gain)
                    Pinhole = Lsm5.Hardware.CpPinholes.Diameter
                    Lsm5.Hardware.CpAmplifiers.Select (channel + 1)
                    Lsm5.Hardware.CpPmts.Select (channel + 1)
                    strName = strFilename + "-ch" + Format(channel, "00") + "-zs" + Format(z, "0000") + strExtension
                
                    Print #intOutFile, "<Image>"
                
                    Print #intOutFile, "Filename" + sTab + strName
                    Print #intOutFile, "DateTime" + sTab + CStr(Format(Now(), "yyyy-mm-dd hh:nn:ss"))
                    Print #intOutFile, "StageX" + sTab + CStr(Lsm5.Hardware.CpStages.PositionX)
                    Print #intOutFile, "StageY" + sTab + CStr(Lsm5.Hardware.CpStages.PositionY)
                    Print #intOutFile, "StageZ" + sTab + CStr(Lsm5.Hardware.CpFocus.Position)
                    'Print #intOutFile, "StageZ" + sTab + CStr(Recording.FocusPosABC1)
                    'Print #intOutFile, "StageZ" + sTab + CStr(Recording.FocusPosABC2)
                    'Just playing around here
                    'Recording.SetABC1Reference = True
                    
                    
                    Print #intOutFile, "Pinhole" + sTab + CStr(Pinhole)
                      
                    'TODO need to add laser attenuation for active lasers and amplifier gain/offset for current channel
                    
                    Print #intOutFile, "</Image>"
                Next channel
            Next z
            
            'free up memory?
            RecordingDoc.CloseAllWindows
            Set RecordingDoc = Nothing
            Set Recording = Nothing
            'TODO is memory leak still a problem? if so does this line help?
                                    
        Next XTilesIndex
    Next YTilesIndex
                
End Sub

Public Function GetMarkedLocations(MyXpos() As Double, MyYpos() As Double, MyZpos() As Double) As Long
   Dim idx As Long

   Dim xPos As Double
   Dim yPos As Double
   Dim zPos As Double

   Dim Result As Long
   Dim Positions As Long
   Dim cnt As Long
   Dim Stage As CpStages

   Set Stage = Lsm5.Hardware.CpStages

   cnt = 0
   On Error GoTo retry
retry:
   If cnt > 1000 Then GoTo Finish
   cnt = cnt + 1


    Positions = Stage.MarkCount
    If Positions < 1 Then
       Positions = 0
    Else
       xPos = Lsm5.Hardware.CpStages.PositionX
       yPos = Lsm5.Hardware.CpStages.PositionY
       zPos = Lsm5.Hardware.CpFocus.Position
       ReDim MyXpos(Positions)
       ReDim MyYpos(Positions)
       ReDim MyZpos(Positions)

       For idx = 1 To Positions
           Result = Lsm5.ExternalCpObject.pHardwareObjects.pStage.pItem(0).GetMarkZ(idx - 1, xPos, yPos, zPos)
           MyXpos(idx) = xPos
           MyYpos(idx) = yPos
           MyZpos(idx) = zPos
       Next idx
   End If
Finish:
   GetMarkedLocations = Positions

End Function

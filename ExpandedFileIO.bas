Attribute VB_Name = "ExpandedFileIO"
Option Explicit

' Written by Matthew Monroe, PNNL
' Started December 23, 2002

' Note: Sub InitFileIOOffsetsAndVersions() defines the file offsets and versions of each
'       of the following sections

Private Const FILEIO_SECTION_COUNT = 12
Private Enum fioSectionConstants
    fioGelData = 0
    fioGelUMC
    fioGelAnalysis
    fioUMCNetAdjDef
    fioSearchDefinitions
    fioORFData                                  ' No longer supported (March 2006)
    fioORFMassTags                              ' No longer supported (March 2006)
    fioGelPairs                                 ' No longer supported (August 2003)
    fioGelDeltaLabeledPairs
    fioGelIDP                                   ' No longer supported (August 2003)
    fioGelLM
    fioGelORFViewerSavedGelListAndOptions       ' No longer supported (March 2006)
End Enum

' The offset to the location that contains a 0, 1, or 2 indicating whether or not the file includes extended info
' 0 = mixed, 1 = does not include extended info, 2 = includes extended info
Private Const FileIO_Offset_IncludesExtendedLoc = 450

' The offset to the first header
Private Const FileIO_Offset_FirstHeader = 500

' The Byte offset of the first section
Private Const FileIO_Offset_GelDataLoc = 5000

' The minimum number of bytes to place between sections
Private Const FileIO_Offset_Spacing_Bytes = 1000

Private FileInfoHeaderOffsets(FILEIO_SECTION_COUNT) As Long
Private FileInfoVersions(FILEIO_SECTION_COUNT) As Single

Private Type udtCertificateType
    strCertificate As String
End Type

Private Const CSNUM_FIELD_COUNT = 8      ' Reserve space for up to 8 numeric fields; we may not use all of these
Private Const CSVAR_FIELD_COUNT = 3

Private Const ISONUM_FIELD_COUNT = 12    ' Reserve space for up to 12 numeric fields; we may not use all of these
Private Const ISOVAR_FIELD_COUNT = 3

' Local variables
Private mGelAnalysis As udtGelAnalysisInfoType      ' Used to hold data copied from the GelAnalysis() object (type FTICRAnalysis, which is part of a class)

Public Function BinaryLoadData(ByVal strFilePath As String, ByVal lngGelIndex As Long, ByRef eFileSaveMode As fsFileSaveModeConstants) As Boolean
    ' strFilePath must point to a valid path
    ' Returns True if success, False if failure
    
    Const FILE_FORMAT_ERROR_MESSAGE As String = "in the input file was saved using a newer (or unknown) file format and therefore cannot be loaded.  Please install the latest version of Viper, available inside PNNL at \\floyd\software\viper\ and outside PNNL at http://omics.pnl.gov/software/  For more information (inside PNNL), visit http://pogo/PASTNews/ or http://proteomics.emsl.pnl.gov/"

    Dim InFileNum As Integer
    Dim intSectionID As Integer
    Dim lngFileSizeBytes As Long
    
    Dim blnUMCDataLoaded As Boolean
    Dim blnSearchDefLoaded As Boolean
    Dim blnPairsDataLoaded As Boolean
    
    Dim lngDataCount As Long
    Dim udtCertificate As udtCertificateType
    
    ' Variables for reading files saved with older versions of the software
    Dim GelData2000 As DocumentData2000
    Dim GelData2003 As DocumentData2003
    Dim GelData2003b As DocumentData2003b
    Dim GelData2004 As DocumentData2004
    Dim GelData2005a As DocumentData2005a
    Dim GelData2005b As DocumentData2005b
    
    Dim GelUMC2002 As UMCListType2002
    Dim GelUMC2003a As UMCListType2003a
    Dim GelUMC2004 As UMCListType2004
    
    Dim GelUMCNETAdjDef2003 As NetAdjDefinition2003
    Dim GelUMCNETAdjDef2004 As NetAdjDefinition2004
    Dim GelUMCNetAdjDef2005a As NetAdjDefinition2005a
    Dim GelUMCNetAdjDef2005b As NetAdjDefinition2005b
    
    Dim GelP_D_L2003 As IsoPairsDltLbl2003Type
    Dim GelP_D_L2004a As IsoPairsDltLbl2004aType
    Dim GelP_D_L2004b As IsoPairsDltLbl2004bType
    Dim GelP_D_L2004c As IsoPairsDltLbl2004cType
    Dim GelP_D_L2004d As IsoPairsDltLbl2004dType
    Dim GelP_D_L2004e As IsoPairsDltLbl2004eType
        
    Dim GelSearchDef2002 As udtSearchDefinition2002GroupType
    Dim GelSearchDef2003 As udtSearchDefinition2003GroupType
    Dim GelSearchDef2003b As udtSearchDefinition2003bGroupType
    Dim GelSearchDef2003c As udtSearchDefinition2003cGroupType
    Dim GelSearchDef2003d As udtSearchDefinition2003dGroupType
    Dim GelSearchDef2003e As udtSearchDefinition2003eGroupType
    Dim GelSearchDef2004 As udtSearchDefinition2004GroupType
    
    Dim sngVersionsInFile(FILEIO_SECTION_COUNT) As Single
    Dim lngOffsetsInFile(FILEIO_SECTION_COUNT) As Long
    Dim lngRecordSizes(FILEIO_SECTION_COUNT) As Long
    
    Dim strCurrentTask As String
    
On Error GoTo BinaryLoadDataErrorHandler
    
    Call InitFileIOOffsetsAndVersions

    lngFileSizeBytes = FileLen(strFilePath)
    ' Need to increase predicted maximum progress bar slightly to compensate for work that follows after the files is loaded
    frmProgress.InitializeForm "Loading data", 0, lngFileSizeBytes * 1.05, True, True, True, MDIForm1
    frmProgress.InitializeSubtask "Reading header information", 0, 1
    
    InFileNum = FreeFile()
    Open strFilePath For Binary Access Read As InFileNum
    
' Read the certificate string from the start of the file
' I don't use the certificate; instead, I use separate version numbers for each section
    Get #InFileNum, 1, udtCertificate
    Debug.Assert udtCertificate.strCertificate = glCERT2003_Modular
    
' Read the byte indicating whether or not extended info is present in the file
    Get #InFileNum, FileIO_Offset_IncludesExtendedLoc, eFileSaveMode
    
' 1. Read the info of the byte offsets of each of the sections
    strCurrentTask = "Read the byte offset and version info for the file sections"
    For intSectionID = 0 To FILEIO_SECTION_COUNT - 1
        BinaryLoadDataGetOffsetInfo FileInfoHeaderOffsets(intSectionID), lngOffsetsInFile(intSectionID), lngRecordSizes(intSectionID), lngDataCount, sngVersionsInFile(intSectionID), InFileNum
    Next intSectionID

    strCurrentTask = "Reading Gel data"
    frmProgress.UpdateProgressBar Loc(InFileNum)
    frmProgress.InitializeSubtask strCurrentTask, 0, 1
    
' 2. Read the Gel Data
     ' Note that versions prior to sngVersionsInFile(fioGelData) = 7# stored the mass calibration info in GelSearchDef().MassCalibrationInfo
     ' This information is now stored in .CSData().MassShiftCount and .IsoData().MassShiftOverallPPM
    If sngVersionsInFile(fioGelData) <> FileInfoVersions(fioGelData) Then
        If lngOffsetsInFile(fioGelData) <= 0 Then
            MsgBox "The file offset for loading the Gel Data is invalid; unable to load data (header information may be corrupted).  Aborting load."
            GoTo BinaryLoadDataCleanup
        End If
    
        If sngVersionsInFile(fioGelData) = 1# Then
            Seek #InFileNum, lngOffsetsInFile(fioGelData)
            Get #InFileNum, , GelData2000
            
            ' Confirm load
            Debug.Assert False
            CopyGelData2000ToCurrent GelData2000, GelData(lngGelIndex)
            
        ElseIf sngVersionsInFile(fioGelData) = 2# Then
            Seek #InFileNum, lngOffsetsInFile(fioGelData)
            Get #InFileNum, , GelData2003
            
            CopyGelData2003ToCurrent GelData2003, GelData(lngGelIndex)
        
        ElseIf sngVersionsInFile(fioGelData) = 3# Then
            Seek #InFileNum, lngOffsetsInFile(fioGelData)
            Get #InFileNum, , GelData2003b
            
            CopyGelData2003bToCurrent GelData2003b, GelData(lngGelIndex)
            
        ElseIf sngVersionsInFile(fioGelData) = 4# Then
            Seek #InFileNum, lngOffsetsInFile(fioGelData)
            Get #InFileNum, , GelData2004
            
            CopyGelData2004ToCurrent GelData2004, GelData(lngGelIndex)
        
        ElseIf sngVersionsInFile(fioGelData) = 5# Then
            Seek #InFileNum, lngOffsetsInFile(fioGelData)
            Get #InFileNum, , GelData2005a
            
            CopyGelData2005aToCurrent GelData2005a, GelData(lngGelIndex)
        
        ElseIf sngVersionsInFile(fioGelData) = 6# Then
            Seek #InFileNum, lngOffsetsInFile(fioGelData)
            Get #InFileNum, , GelData2005b
            
            CopyGelData2005bToCurrent GelData2005b, GelData(lngGelIndex)
            
        Else
            MsgBox "The gel data " & FILE_FORMAT_ERROR_MESSAGE & "  Aborting load."
            GoTo BinaryLoadDataCleanup
        End If
    Else
        Seek #InFileNum, lngOffsetsInFile(fioGelData)
        Get #InFileNum, , GelData(lngGelIndex)
    End If
        
    ' Validate that .CSData and .IsoData have at least one entry
    ValidateDataArrays lngGelIndex

    ' Populate the adjacent scan pointer arrays
    UpdateGelAdjacentScanPointerArrays lngGelIndex
    
    If eFileSaveMode <> fsNoExtended Then
    ' 3. Read the UMC Data
        strCurrentTask = "Reading UMC data"
        frmProgress.UpdateProgressBar Loc(InFileNum)
        frmProgress.InitializeSubtask strCurrentTask, 0, 1
        
        If lngOffsetsInFile(fioGelUMC) <= 0 Then
            MsgBox "The file offset for loading UMC data is invalid; unable to load UMCs (header information may be corrupted)"
        Else
        
            If sngVersionsInFile(fioGelUMC) <> FileInfoVersions(fioGelUMC) Then
            
                If sngVersionsInFile(fioGelUMC) = 1# Then
                    ' Update from version 1 to current version
                    Seek #InFileNum, lngOffsetsInFile(fioGelUMC)
                    Get #InFileNum, , GelUMC2002
                    blnUMCDataLoaded = True
                    
                    frmProgress.UpdateCurrentSubTask "Updating LC-MS Feature data format to current version"
                    CopyGelUMC2002ToCurrent GelUMC2002, GelUMC(lngGelIndex)
                ElseIf sngVersionsInFile(fioGelUMC) = 2# Then
                    ' Update from version 2 to current version
                    Seek #InFileNum, lngOffsetsInFile(fioGelUMC)
                    Get #InFileNum, , GelUMC2003a
                    blnUMCDataLoaded = True
                    
                    frmProgress.UpdateCurrentSubTask "Updating LC-MS Feature data format to current version"
                    CopyGelUMC2003aToCurrent GelUMC2003a, GelUMC(lngGelIndex)
                
                ElseIf sngVersionsInFile(fioGelUMC) = 3# Then
                    ' Update from version 3 to current version
                    Seek #InFileNum, lngOffsetsInFile(fioGelUMC)
                    Get #InFileNum, , GelUMC2004
                    blnUMCDataLoaded = True
                    
                    frmProgress.UpdateCurrentSubTask "Updating LC-MS Feature data format to current version"
                    CopyGelUMC2004ToCurrent GelUMC2004, GelUMC(lngGelIndex)
                
                ElseIf sngVersionsInFile(fioGelUMC) = 4# Then
                    ' Update from version 4 to current version
                    Seek #InFileNum, lngOffsetsInFile(fioGelUMC)
                    Get #InFileNum, , GelUMC(lngGelIndex)
                    blnUMCDataLoaded = True
                    
                    frmProgress.UpdateCurrentSubTask "Updating LC-MS Feature data format to current version"
                    InitializeAdditionalUMCDefVariables GelUMC(lngGelIndex).def
                    
                Else
                    MsgBox "The LC-MS Feature data " & FILE_FORMAT_ERROR_MESSAGE
                End If
            Else
                Seek #InFileNum, lngOffsetsInFile(fioGelUMC)
                Get #InFileNum, , GelUMC(lngGelIndex)
                blnUMCDataLoaded = True
            End If
        End If
        
        If blnUMCDataLoaded Then
            ' The following calls CalculateClasses, UpdateIonToUMCIndices, and InitDrawUMC
            
            Dim blnComputeClassMass As Boolean
            Dim blnComputeClassAbundance As Boolean
            
            If GelUMC(lngGelIndex).def.LoadedPredefinedLCMSFeatures Then
                blnComputeClassMass = False
                blnComputeClassAbundance = False
            Else
                blnComputeClassMass = True
                blnComputeClassAbundance = True
            End If
        
            UpdateUMCStatArrays lngGelIndex, blnComputeClassMass, blnComputeClassAbundance, False
        End If
    
    End If
    
' 4. Read the Analysis Info Data
    strCurrentTask = "Reading Analysis Info"
    frmProgress.UpdateProgressBar Loc(InFileNum)
    frmProgress.InitializeSubtask strCurrentTask, 0, 1
    
    If lngOffsetsInFile(fioGelAnalysis) > 0 Then
        Seek #InFileNum, lngOffsetsInFile(fioGelAnalysis)
        GelAnalysisInfoRead InFileNum, lngGelIndex
    End If
    
' 5. Read the UMC Net Adjustment Definition values
    strCurrentTask = "Reading UMC Net parameters"
    frmProgress.UpdateProgressBar Loc(InFileNum)
    frmProgress.InitializeSubtask strCurrentTask, 0, 1

    If lngOffsetsInFile(fioUMCNetAdjDef) > 0 Then
        If sngVersionsInFile(fioUMCNetAdjDef) <> FileInfoVersions(fioUMCNetAdjDef) Then
            If sngVersionsInFile(fioUMCNetAdjDef) = 1# Then
                ' Update from version 1 to version 3
                Seek #InFileNum, lngOffsetsInFile(fioUMCNetAdjDef)
                Get #InFileNum, , GelUMCNETAdjDef2003
                CopyGelNetAdjDef2003ToCurrent GelUMCNETAdjDef2003, GelUMCNETAdjDef(lngGelIndex)
            ElseIf sngVersionsInFile(fioUMCNetAdjDef) = 2# Then
                MsgBox "The UMC Net Adjustment definition in the input file was saved using a December 2003 beta-version file format and therefore cannot be loaded.  Defaults will be used instead."
                SetDefaultUMCNETAdjDef GelUMCNETAdjDef(lngGelIndex)
            ElseIf sngVersionsInFile(fioUMCNetAdjDef) = 3# Then
                Seek #InFileNum, lngOffsetsInFile(fioUMCNetAdjDef)
                Get #InFileNum, , GelUMCNETAdjDef2004
                CopyGelNetAdjDef2004ToCurrent GelUMCNETAdjDef2004, GelUMCNETAdjDef(lngGelIndex)
            ElseIf sngVersionsInFile(fioUMCNetAdjDef) = 4# Then
                Seek #InFileNum, lngOffsetsInFile(fioUMCNetAdjDef)
                Get #InFileNum, , GelUMCNetAdjDef2005a
                CopyGelNetAdjDef2005aToCurrent GelUMCNetAdjDef2005a, GelUMCNETAdjDef(lngGelIndex)
            ElseIf sngVersionsInFile(fioUMCNetAdjDef) = 5# Then
                Seek #InFileNum, lngOffsetsInFile(fioUMCNetAdjDef)
                Get #InFileNum, , GelUMCNetAdjDef2005b
                CopyGelNetAdjDef2005bToCurrent GelUMCNetAdjDef2005b, GelUMCNETAdjDef(lngGelIndex)
            Else
                MsgBox "The UMC Net Adjustment definition " & FILE_FORMAT_ERROR_MESSAGE
                SetDefaultUMCNETAdjDef GelUMCNETAdjDef(lngGelIndex)
            End If
            
            If sngVersionsInFile(fioUMCNetAdjDef) <= 5# And Not APP_BUILD_DISABLE_LCMSWARP Then
                ' Force RobustNetAdjustment with warping to be enabled if the version is <= 5#
                With GelUMCNETAdjDef(lngGelIndex)
                    .UseRobustNETAdjustment = True
                    .RobustNETAdjustmentMode = UMCRobustNETModeConstants.UMCRobustNETWarpTimeAndMass
                End With
            End If
        Else
            Seek #InFileNum, lngOffsetsInFile(fioUMCNetAdjDef)
            Get #InFileNum, , GelUMCNETAdjDef(lngGelIndex)
            
            UMCNetAdjDef = GelUMCNETAdjDef(lngGelIndex)
        End If
    End If
    
' 6. Read the Search Definition values
    strCurrentTask = "Reading Search Definitions"
    frmProgress.UpdateProgressBar Loc(InFileNum)
    frmProgress.InitializeSubtask strCurrentTask, 0, 1
    
    If lngOffsetsInFile(fioSearchDefinitions) > 0 Then
        If sngVersionsInFile(fioSearchDefinitions) <> FileInfoVersions(fioSearchDefinitions) Then
            If sngVersionsInFile(fioSearchDefinitions) = 2# Then
                ' Update from version 2 to the current version
                Seek #InFileNum, lngOffsetsInFile(fioSearchDefinitions)
                Get #InFileNum, , GelSearchDef2002
                
                CopyGelSearchDef2002ToCurrent GelSearchDef2002, GelSearchDef(lngGelIndex)
                blnSearchDefLoaded = True
            ElseIf sngVersionsInFile(fioSearchDefinitions) = 3# Then
                ' Update from Version 3 to the current version
                Seek #InFileNum, lngOffsetsInFile(fioSearchDefinitions)
                Get #InFileNum, , GelSearchDef2003
                
                CopyGelSearchDef2003ToCurrent GelSearchDef2003, GelSearchDef(lngGelIndex)
                blnSearchDefLoaded = True
            ElseIf sngVersionsInFile(fioSearchDefinitions) = 4# Then
                ' Update from Version 4 to the current version
                Seek #InFileNum, lngOffsetsInFile(fioSearchDefinitions)
                Get #InFileNum, , GelSearchDef2003b
                
                CopyGelSearchDef2003bToCurrent GelSearchDef2003b, GelSearchDef(lngGelIndex)
                blnSearchDefLoaded = True
            ElseIf sngVersionsInFile(fioSearchDefinitions) = 5# Then
                ' Update from Version 5 to the current version
                Seek #InFileNum, lngOffsetsInFile(fioSearchDefinitions)
                Get #InFileNum, , GelSearchDef2003c
                
                CopyGelSearchDef2003cToCurrent GelSearchDef2003c, GelSearchDef(lngGelIndex)
                blnSearchDefLoaded = True
            ElseIf sngVersionsInFile(fioSearchDefinitions) = 6# Then
                ' Update from Version 6 to the current version
                Seek #InFileNum, lngOffsetsInFile(fioSearchDefinitions)
                Get #InFileNum, , GelSearchDef2003d
                
                CopyGelSearchDef2003dToCurrent GelSearchDef2003d, GelSearchDef(lngGelIndex)
                blnSearchDefLoaded = True
            ElseIf sngVersionsInFile(fioSearchDefinitions) = 7# Then
                ' Update from Version 7 to the current version
                Seek #InFileNum, lngOffsetsInFile(fioSearchDefinitions)
                Get #InFileNum, , GelSearchDef2003e
                
                CopyGelSearchDef2003eToCurrent GelSearchDef2003e, GelSearchDef(lngGelIndex)
                blnSearchDefLoaded = True
            ElseIf sngVersionsInFile(fioSearchDefinitions) = 8# Then
                ' Update from Version 8 to the current version
                Seek #InFileNum, lngOffsetsInFile(fioSearchDefinitions)
                Get #InFileNum, , GelSearchDef2004
                
                CopyGelSearchDef2004ToCurrent GelSearchDef2004, GelSearchDef(lngGelIndex)
                blnSearchDefLoaded = True
            Else
                MsgBox "The Search Definitions " & FILE_FORMAT_ERROR_MESSAGE
            End If
        Else
            Seek #InFileNum, lngOffsetsInFile(fioSearchDefinitions)
            Get #InFileNum, , GelSearchDef(lngGelIndex)
            blnSearchDefLoaded = True
        End If
    
        If GelSearchDef(lngGelIndex).AMTSearchMassMods.ModMode > 2 Then
            ' We loaded an old file that was saved with .DynamicMods instead of .ModMode
            GelSearchDef(lngGelIndex).AMTSearchMassMods.ModMode = 1
        End If
        
        If GelSearchDef(lngGelIndex).AMTSearchMassMods.UnusedByte <> 0 Then
            ' Loaded an old file that was saved with .DynamicMods instead of .ModMode
            GelSearchDef(lngGelIndex).AMTSearchMassMods.UnusedByte = 0
        End If
        
        If blnSearchDefLoaded Then
            ' Copy values from GelUMC(lngGelIndex).Def to UMCDef (part of Module12, UC.Bas)
            With GelSearchDef(lngGelIndex).UMCDef
                If .Tol > 0 Or .MWField > 0 Then
                    ' Data exists; copy
                    UMCDef = GelSearchDef(lngGelIndex).UMCDef
                End If
            End With
            
            If sngVersionsInFile(fioGelData) < 7# Then
                CopyLegacyMassCalibrationInfoToData GelData(lngGelIndex), GelSearchDef(lngGelIndex).MassCalibrationInfo
            End If
        End If
    End If
    
    If eFileSaveMode <> fsNoExtended Then
    ' 7. Read the ORF Data
        strCurrentTask = "Reading Protein information"
        frmProgress.UpdateProgressBar Loc(InFileNum)
        frmProgress.InitializeSubtask strCurrentTask, 0, 1
      
''        If sngVersionsInFile(fioORFData) <> FileInfoVersions(fioORFData) Then
''            MsgBox "The Protein information " & FILE_FORMAT_ERROR_MESSAGE
''        Else
''            Seek #InFileNum, lngOffsetsInFile(fioORFData)
''            ' Unused variable (March 2006)
''            Get #InFileNum, , GelORFData(lngGelIndex)
''        End If
''
    ' 8. Read the ORF MT tags
        strCurrentTask = "Reading the MT tags for the Proteins"
        frmProgress.UpdateProgressBar Loc(InFileNum)
        frmProgress.InitializeSubtask strCurrentTask, 0, 1
      
      ' No longer supported (March 2006)
''        If sngVersionsInFile(fioORFMassTags) <> FileInfoVersions(fioORFMassTags) Then
''            MsgBox "The Protein MT tags information " & FILE_FORMAT_ERROR_MESSAGE
''        Else
''            Seek #InFileNum, lngOffsetsInFile(fioORFMassTags)
''            ' Unused variable (March 2006)
''            Get #InFileNum, , GelORFMassTags(lngGelIndex)
''        End If

    ' 9. Read the Pairs Data
        strCurrentTask = "Reading pairs"
        frmProgress.UpdateProgressBar Loc(InFileNum)
        frmProgress.InitializeSubtask strCurrentTask, 0, 1
        
''        If sngVersionsInFile(fioGelPairs) <> FileInfoVersions(fioGelPairs) Then
''            MsgBox "The pairs data " & FILE_FORMAT_ERROR_MESSAGE
''        Else
''            Seek #InFileNum, lngOffsetsInFile(fioGelPairs)
''            ' Unused variable (August 2003)
''            Get #InFileNum, , GelP(lngGelIndex)
''        End If
    
    ' 10. Read the Delta Labeled Pairs Data
        strCurrentTask = "Reading delta labeled pairs"
        frmProgress.UpdateProgressBar Loc(InFileNum)
        frmProgress.InitializeSubtask strCurrentTask, 0, 1
        
        If lngOffsetsInFile(fioGelDeltaLabeledPairs) > 0 Then
            
            If sngVersionsInFile(fioGelDeltaLabeledPairs) <> FileInfoVersions(fioGelDeltaLabeledPairs) Then
                If sngVersionsInFile(fioGelDeltaLabeledPairs) = 2# Then
                    ' Update from version 2 to current version
                    Seek #InFileNum, lngOffsetsInFile(fioGelDeltaLabeledPairs)
                    Get #InFileNum, , GelP_D_L2003
                    
                    CopyDeltaLabelPairs2003ToCurrent GelP_D_L2003, GelP_D_L(lngGelIndex)
                    blnPairsDataLoaded = True
                    
                ElseIf sngVersionsInFile(fioGelDeltaLabeledPairs) = 3# Then
                    ' Update from version 3 to current version
                    Seek #InFileNum, lngOffsetsInFile(fioGelDeltaLabeledPairs)
                    Get #InFileNum, , GelP_D_L2004a
                    
                    CopyDeltaLabelPairs2004aToCurrent GelP_D_L2004a, GelP_D_L(lngGelIndex)
                    blnPairsDataLoaded = True
                    
                ElseIf sngVersionsInFile(fioGelDeltaLabeledPairs) = 4# Then
                    Seek #InFileNum, lngOffsetsInFile(fioGelDeltaLabeledPairs)
                    Get #InFileNum, , GelP_D_L2004b
                    
                    CopyDeltaLabelPairs2004bToCurrent GelP_D_L2004b, GelP_D_L(lngGelIndex)
                    blnPairsDataLoaded = True
                    
                ElseIf sngVersionsInFile(fioGelDeltaLabeledPairs) = 5# Then
                    Seek #InFileNum, lngOffsetsInFile(fioGelDeltaLabeledPairs)
                    Get #InFileNum, , GelP_D_L2004c
                    
                    CopyDeltaLabelPairs2004cToCurrent GelP_D_L2004c, GelP_D_L(lngGelIndex)
                    blnPairsDataLoaded = True
                    
                ElseIf sngVersionsInFile(fioGelDeltaLabeledPairs) = 6# Then
                    Seek #InFileNum, lngOffsetsInFile(fioGelDeltaLabeledPairs)
                    Get #InFileNum, , GelP_D_L2004d
                    
                    CopyDeltaLabelPairs2004dToCurrent GelP_D_L2004d, GelP_D_L(lngGelIndex)
                    blnPairsDataLoaded = True
                ElseIf sngVersionsInFile(fioGelDeltaLabeledPairs) = 7# Then
                    Seek #InFileNum, lngOffsetsInFile(fioGelDeltaLabeledPairs)
                    Get #InFileNum, , GelP_D_L2004e
                    
                    CopyDeltaLabelPairs2004eToCurrent GelP_D_L2004e, GelP_D_L(lngGelIndex)
                    blnPairsDataLoaded = True
                Else
                    MsgBox "The pairs data " & FILE_FORMAT_ERROR_MESSAGE
                End If
            Else
                Seek #InFileNum, lngOffsetsInFile(fioGelDeltaLabeledPairs)
                Get #InFileNum, , GelP_D_L(lngGelIndex)
                blnPairsDataLoaded = True
            End If
        
            If blnPairsDataLoaded Then
                glbPreferencesExpanded.PairSearchOptions.SearchDef = GelP_D_L(lngGelIndex).SearchDef
            End If
        End If
        
    ' 11. Read the Identified Pairs (IDP) Data
        strCurrentTask = "Reading IDP data"
        frmProgress.UpdateProgressBar Loc(InFileNum)
        frmProgress.InitializeSubtask strCurrentTask, 0, 1
        
''        If sngVersionsInFile(fioGelIDP) <> FileInfoVersions(fioGelIDP) Then
''            MsgBox "The IDP data " & FILE_FORMAT_ERROR_MESSAGE
''        Else
''            Seek #InFileNum, lngOffsetsInFile(fioGelIDP)
''            ' Unused variable (August 2003)
''            Get #InFileNum, , GelIDP(lngGelIndex)
''        End If
    
    ' 12. Read the Gel Lock Mass (LM) Data
        strCurrentTask = "Reading LM pairs"
        frmProgress.UpdateProgressBar Loc(InFileNum)
        frmProgress.InitializeSubtask strCurrentTask, 0, 1
        
        If sngVersionsInFile(fioGelLM) <> FileInfoVersions(fioGelLM) Then
            MsgBox "The Gel Lock Mass data " & FILE_FORMAT_ERROR_MESSAGE
        Else
            If lngOffsetsInFile(fioGelLM) > 0 Then
                Seek #InFileNum, lngOffsetsInFile(fioGelLM)
                Get #InFileNum, , GelLM(lngGelIndex)
            End If
        End If

    ' 13. Read the ORF Viewer Saved Gel List And Options
''        frmProgress.UpdateProgressBar Loc(InFileNum)
''        frmProgress.InitializeSubtask "Reading ORF Viewer Options", 0, 1
''
''        If sngVersionsInFile(fioGelORFViewerSavedGelListAndOptions) <> FileInfoVersions(fioGelORFViewerSavedGelListAndOptions) Then
''            MsgBox "The ORF Viewer options " & FILE_FORMAT_ERROR_MESSAGE
''        Else
''            Seek #InFileNum, lngOffsetsInFile(fioGelORFViewerSavedGelListAndOptions)
''            Get #InFileNum, , GelORFViewerSavedGelListAndOptions(lngGelIndex)
''        End If
    End If

    frmProgress.InitializeSubtask "Done loading", 0, 1
    
    Close #InFileNum
        
    BinaryLoadData = True

BinaryLoadDataCleanup:
    On Error Resume Next
    Close #InFileNum
    frmProgress.HideForm
    Exit Function
    
BinaryLoadDataErrorHandler:
    MsgBox "Error reading input file " & strFilePath & " (" & strCurrentTask & ")" & vbCrLf & Err.Description & vbCrLf & "Aborting.", vbExclamation + vbOKOnly, "Error"
    LogErrors Err.Number, "BinaryLoadData", Err.Description, lngGelIndex
    BinaryLoadData = False
    Resume BinaryLoadDataCleanup

End Function

Private Sub BinaryLoadDataGetOffsetInfo(ByVal lngInfoOffset As Long, ByRef lngActualOffsetOfData As Long, ByRef lngRecordSize As Long, ByRef lngDataCount As Long, ByRef sngDataFormatVersion As Single, ByVal InFileNum As Integer)
    
    Seek #InFileNum, lngInfoOffset
    Get #InFileNum, , lngActualOffsetOfData
    Get #InFileNum, , lngRecordSize
    Get #InFileNum, , lngDataCount
    Get #InFileNum, , sngDataFormatVersion
    
End Sub

Public Function BinarySaveLegacy(strFilePath As String, lngGelIndex As Long) As Boolean
    ' Returns True if no error, false if an error

    Dim OutFileNum As Long
    Dim LegacyData As DocumentData2003
    
    Dim MaxInd As Long
    Dim i As Long, j As Long
    
    On Error GoTo BinarySaveLegacyErrorHandler
    
    OutFileNum = FreeFile()
    Open strFilePath For Binary Access Write As OutFileNum
    
    If Err.Number <> 0 Then
       MsgBox "Error creating file: " & strFilePath, vbOKOnly
       LogErrors Err.Number, "BinarySaveLegacy"
       Exit Function
    End If
    
    Screen.MousePointer = vbHourglass
    
    With LegacyData
        .Certificate = glCERT2003
        .Comment = GelData(lngGelIndex).Comment
        .FileName = GelData(lngGelIndex).FileName
        .Fileinfo = GelData(lngGelIndex).Fileinfo
        .PathtoDataFiles = GelData(lngGelIndex).PathtoDataFiles
        .PathtoDatabase = GelData(lngGelIndex).PathtoDatabase
        
        .MediaType = GelData(lngGelIndex).MediaType
        .LinesRead = GelData(lngGelIndex).LinesRead
        .DataLines = GelData(lngGelIndex).DataLines
        .CSLines = GelData(lngGelIndex).CSLines
        .IsoLines = GelData(lngGelIndex).IsoLines
        .CalEquation = GelData(lngGelIndex).CalEquation
        For i = 1 To 10
            .CalArg(i) = GelData(lngGelIndex).CalArg(i)
        Next i
        .MinMW = GelData(lngGelIndex).MinMW
        .MaxMW = GelData(lngGelIndex).MaxMW
        .MinAbu = GelData(lngGelIndex).MinAbu
        .MaxAbu = GelData(lngGelIndex).MaxAbu
        .Preferences = GelData(lngGelIndex).Preferences
        
        .pICooSysEnabled = GelData(lngGelIndex).pICooSysEnabled
        
        For i = 1 To MAX_FILTER_COUNT_2003
            For j = 0 To 2
                .DataFilter(i, j) = GelData(lngGelIndex).DataFilter(i, j)
            Next j
        Next i
        
        MaxInd = UBound(GelData(lngGelIndex).ScanInfo)
        ReDim .DFN(1 To MaxInd)
        ReDim .DFFN(1 To MaxInd)
        ReDim .DFPI(1 To MaxInd)
        ReDim .DFFS(1 To MaxInd)
        ReDim .DFIN(1 To MaxInd)
        
        For i = 1 To MaxInd
            .DFFN(i) = GelData(lngGelIndex).ScanInfo(i).ScanNumber
            .DFN(i) = GelData(lngGelIndex).ScanInfo(i).ScanFileName
            .DFPI(i) = GelData(lngGelIndex).ScanInfo(i).ScanPI
            .DFFS(i) = GelData(lngGelIndex).ScanInfo(i).FrequencyShift
            .DFIN(i) = GelData(lngGelIndex).ScanInfo(i).TimeDomainSignal
        Next i
        
        If GelData(lngGelIndex).CSLines > 0 Then
           ReDim .CSNum(1 To GelData(lngGelIndex).CSLines, 1 To CSNUM_FIELD_COUNT)
           ReDim .CSVar(1 To GelData(lngGelIndex).CSLines, 1 To CSVAR_FIELD_COUNT)
        
            For i = 1 To .CSLines
                CopyCSDataToLegacy GelData(lngGelIndex).IsoData(i), .CSNum, .CSVar, i
            Next i
        End If
        
        If GelData(lngGelIndex).IsoLines > 0 Then
           ReDim .IsoNum(1 To GelData(lngGelIndex).IsoLines, 1 To ISONUM_FIELD_COUNT)
           ReDim .IsoVar(1 To GelData(lngGelIndex).IsoLines, 1 To ISOVAR_FIELD_COUNT)
        
            For i = 1 To .IsoLines
                CopyIsoDataToLegacy GelData(lngGelIndex).IsoData(i), .IsoNum, .IsoVar, i
            Next i
        End If
        
    End With
    
    Put #OutFileNum, , LegacyData
    
    If Err Then
        MsgBox "Unexpected error." & sErrLogReference, vbOKOnly
        LogErrors Err.Number, "BinarySaveLegacy"
    Else
        GelBody(lngGelIndex).Caption = strFilePath
        GelStatus(lngGelIndex).GelFilePathFull = GetFilePathFull(strFilePath)
       
        GelStatus(lngGelIndex).Dirty = False
        UpdateFileMenu strFilePath
    End If
    Close OutFileNum
    
    Screen.MousePointer = vbDefault
    
    BinarySaveLegacy = True
    
    Exit Function

BinarySaveLegacyErrorHandler:
Debug.Assert False
Resume Next

End Function

Public Function BinarySaveData(strFilePath As String, ByVal blnRemoveUMCData As Boolean, ByVal blnRemovePairsData As Boolean, ByVal lngGelIndex As Long, ByRef udtThisGelData As DocumentData, ByVal eFileSaveMode As fsFileSaveModeConstants) As Boolean
    ' Returns True if no error, false if an error
    ' strFilePath must point to a valid path
    ' Note: If eFileSaveMode = fsNoExtended, then blnRemoveUMCData and blnRemovePairsData are set to True
    
    ' Note that binary save over the network can be slow
    ' Consequently, if the App.Path drive is not the same as the target folder drive, then we first save a temporary file, then we copy the file to the desired destination
    
    Dim OutFileNum As Integer, InFileNum As Integer
    Dim intSectionID As Integer
    Dim lngProgressStepCount As Long
    
    Dim lngDataCount As Long
    Dim udtCertificate As udtCertificateType
    
    Dim sngVersionsInFile(FILEIO_SECTION_COUNT) As Single
    Dim lngOffsetsInFile(FILEIO_SECTION_COUNT) As Long
    Dim lngRecordSizes(FILEIO_SECTION_COUNT) As Long
    Dim blnDataChanged(FILEIO_SECTION_COUNT) As Boolean
    
    Dim objRemoteSaveFileHandler As New clsRemoteSaveFileHandler
    Dim strWorkingFilePath As String
    Dim blnSuccess As Boolean
    
    Dim fso As New FileSystemObject

On Error GoTo BinarySaveDataErrorHandler
    
    ' Initialize the File IO Offset Arrays
    Call InitFileIOOffsetsAndVersions
    
    ' Make sure lngGelIndex is valid
    If lngGelIndex < LBound(GelData()) Or lngGelIndex > UBound(GelData()) Then
        MsgBox "Invalid file index: " & lngGelIndex
        BinarySaveData = False
        Set fso = Nothing
        Exit Function
    End If
    
    ' Make sure GelUMCNETAdjDef() contains valid data
    With GelUMCNETAdjDef(lngGelIndex)
        If .TopAbuPct = 0 And .NETTolIterative = 0 And .MWTol = 0 And .IterationStopValue = 0 And .MSWarpOptions.NETTol = 0 Then
            ' Structure is probably empty, copy values from UMCNetAdjDef
            GelUMCNETAdjDef(lngGelIndex) = UMCNetAdjDef
        End If
    End With
    
    Screen.MousePointer = vbHourglass
    
    frmProgress.InitializeForm "Saving data to: " & vbCrLf & CompactPathString(strFilePath, 48), 0, FILEIO_SECTION_COUNT - 1, True, True, True, MDIForm1
    lngProgressStepCount = 0
    
    ' Define the offset of the gel data
    lngOffsetsInFile(fioGelData) = FileIO_Offset_GelDataLoc
    
    ' For now, we'll always re-write all of the data
    ' Could add checks later to determine whether or not to re-write all of the data
    ' Note that below I set some of the sections to false if blnCompressByExcludingORFData = True
    For intSectionID = 0 To FILEIO_SECTION_COUNT - 1
        blnDataChanged(intSectionID) = True
    Next intSectionID
    
    If eFileSaveMode = fsNoExtended Then
        blnRemoveUMCData = True
        blnRemovePairsData = True
        ''blnRemoveORFData = True
    End If
    
    ''If blnRemoveUMCData Or blnRemoveORFData Or blnRemovePairsData Then
    If blnRemoveUMCData Or blnRemovePairsData Then
        ' Need to erase file so that it gets fully re-written
        On Error Resume Next
        fso.DeleteFile strFilePath, True
        Debug.Assert fso.FileExists(strFilePath) = False
        On Error GoTo BinarySaveDataErrorHandler

        If blnRemoveUMCData Then
            ' If removing UMC data then can't write pairs either
            blnRemovePairsData = True
            ''blnRemoveORFData = True
        End If
        
    End If
    
    ' Initialize strWorkingFilePath
    ' This will point to a temporary file if strFilePath is actually on a different drive than App.Path
    strWorkingFilePath = objRemoteSaveFileHandler.GetTempFilePath(strFilePath, False)
    
    ' If file exists, determine location of data in file in case we don't need to save all of the data
    If fso.FileExists(strWorkingFilePath) Then
        ' Need to open the file and read the record sizes of the currently stored data
        InFileNum = FreeFile()
        Open strWorkingFilePath For Binary Access Read As #InFileNum
        
        Get #InFileNum, 1, udtCertificate
        If udtCertificate.strCertificate = glCERT2003_Modular Then
            For intSectionID = 0 To FILEIO_SECTION_COUNT - 1
                BinaryLoadDataGetOffsetInfo FileInfoHeaderOffsets(intSectionID), lngOffsetsInFile(intSectionID), lngRecordSizes(intSectionID), lngDataCount, sngVersionsInFile(intSectionID), InFileNum
            Next intSectionID
            Close #InFileNum
        Else
            ' Old file format; need to erase the file before updating to the new format
            ' First close
            Close #InFileNum
            
            ' Now wait 100 msec, erase the file, and wait another 100 msec before continuing
            On Error Resume Next
            Sleep 100
            Kill strWorkingFilePath
            Sleep 100
            
            On Error GoTo BinarySaveDataErrorHandler
        
            ' Make sure blnDataChanged() is True for all sections
            For intSectionID = 0 To FILEIO_SECTION_COUNT - 1
                blnDataChanged(intSectionID) = True
            Next intSectionID
        End If
    End If
    
    OutFileNum = FreeFile()
    Open strWorkingFilePath For Binary Access Write As #OutFileNum
    
   
' Write the certificate string to the start of the file
' Update the certificate
    udtThisGelData.Certificate = glCERT2003_Modular

    udtCertificate.strCertificate = udtThisGelData.Certificate
    Put #OutFileNum, 1, udtCertificate
    
' Write a 0, 1, or 2 reflecting the value of eFileSaveMode
    Put #OutFileNum, FileIO_Offset_IncludesExtendedLoc, eFileSaveMode

' 1a. Write the offset of the Gel Data
    BinarySaveDataInitOffsetInfo FileInfoHeaderOffsets(fioGelData), lngOffsetsInFile(fioGelData), LenB(udtThisGelData), 1, FileInfoVersions(fioGelData), OutFileNum
    
' 1b. Write the offset of the Gel UMC info (actual offset is currently unknown)
    BinarySaveDataInitOffsetInfo FileInfoHeaderOffsets(fioGelUMC), 0, LenB(GelUMC(lngGelIndex)), 1, FileInfoVersions(fioGelUMC), OutFileNum
        
' 1c. Write the offset of the Gel Analysis info
    BinarySaveDataInitOffsetInfo FileInfoHeaderOffsets(fioGelAnalysis), 0, LenB(mGelAnalysis), 1, FileInfoVersions(fioGelAnalysis), OutFileNum
    
' 1d. Write the offset of the Gel UMC Net Adjustment Definition
    BinarySaveDataInitOffsetInfo FileInfoHeaderOffsets(fioUMCNetAdjDef), 0, LenB(GelUMCNETAdjDef(lngGelIndex)), 1, FileInfoVersions(fioUMCNetAdjDef), OutFileNum

' 1e. Write the offset of the Gel Search Definitions
    BinarySaveDataInitOffsetInfo FileInfoHeaderOffsets(fioSearchDefinitions), 0, LenB(GelSearchDef(lngGelIndex)), 1, FileInfoVersions(fioSearchDefinitions), OutFileNum

' 1f. Write the offset of the Gel ORF information
    ' No longer supported (March 2006)
    ''BinarySaveDataInitOffsetInfo FileInfoHeaderOffsets(fioORFData), 0, LenB(GelORFData(lngGelIndex)), 1, FileInfoVersions(fioORFData), OutFileNum
    BinarySaveDataInitOffsetInfo FileInfoHeaderOffsets(fioORFData), 0, 0, 1, FileInfoVersions(fioORFData), OutFileNum

' 1g. Write the offset of the Gel ORF MT tag information
    ' No longer supported (March 2006)
    ''BinarySaveDataInitOffsetInfo FileInfoHeaderOffsets(fioORFMassTags), 0, LenB(GelORFMassTags(lngGelIndex)), 1, FileInfoVersions(fioORFMassTags), OutFileNum
    BinarySaveDataInitOffsetInfo FileInfoHeaderOffsets(fioORFMassTags), 0, 0, 1, FileInfoVersions(fioORFMassTags), OutFileNum

' 1h. Write the offset of the Gel Pairs info
    ' GelP is an unused variable (August 2003)
    ''BinarySaveDataInitOffsetInfo FileInfoHeaderOffsets(fioGelPairs), 0, LenB(GelP(lngGelIndex)), 1, FileInfoVersions(fioGelPairs), OutFileNum
    BinarySaveDataInitOffsetInfo FileInfoHeaderOffsets(fioGelPairs), 0, 0, 1, FileInfoVersions(fioGelPairs), OutFileNum
    
' 1i. Write the offset of the Gel Delta Labeled Pairs info
    BinarySaveDataInitOffsetInfo FileInfoHeaderOffsets(fioGelDeltaLabeledPairs), 0, LenB(GelP_D_L(lngGelIndex)), 1, FileInfoVersions(fioGelDeltaLabeledPairs), OutFileNum
    
' 1j. Write the offset of the Gel Identified Pairs (IDP) info
    ' GelIDP is an unused variable (August 2003)
    ' BinarySaveDataInitOffsetInfo FileInfoHeaderOffsets(fioGelIDP), 0, LenB(GelIDP(lngGelIndex)), 1, FileInfoVersions(fioGelIDP), OutFileNum
    BinarySaveDataInitOffsetInfo FileInfoHeaderOffsets(fioGelIDP), 0, 0, 1, FileInfoVersions(fioGelIDP), OutFileNum

' 1k. Write the offset of the Gel LM info
    BinarySaveDataInitOffsetInfo FileInfoHeaderOffsets(fioGelLM), 0, LenB(GelLM(lngGelIndex)), 1, FileInfoVersions(fioGelLM), OutFileNum
    
''' 1l. Write the offset of the Gel ORF Viewer options
    ''BinarySaveDataInitOffsetInfo FileInfoHeaderOffsets(fioGelORFViewerSavedGelListAndOptions), 0, LenB(GelORFViewerSavedGelListAndOptions(lngGelIndex)), 1, FileInfoVersions(fioGelORFViewerSavedGelListAndOptions), OutFileNum
    BinarySaveDataInitOffsetInfo FileInfoHeaderOffsets(fioGelORFViewerSavedGelListAndOptions), 0, 0, 1, FileInfoVersions(fioGelORFViewerSavedGelListAndOptions), OutFileNum
    
    frmProgress.UpdateProgressBar lngProgressStepCount
    frmProgress.InitializeSubtask "Writing Gel Data", 0, 1
    
' 2. Write the Gel Data
    
    Debug.Assert Loc(OutFileNum) <= lngOffsetsInFile(fioGelData)
    Seek #OutFileNum, lngOffsetsInFile(fioGelData)
    If blnDataChanged(fioGelData) Then
        Put #OutFileNum, , udtThisGelData
        
        lngProgressStepCount = lngProgressStepCount + 1
        frmProgress.UpdateProgressBar lngProgressStepCount
        
        blnDataChanged(fioGelData) = False
    Else
        Debug.Assert Loc(OutFileNum) < lngOffsetsInFile(fioGelUMC) - FileIO_Offset_Spacing_Bytes
        Seek #OutFileNum, lngOffsetsInFile(fioGelUMC) - FileIO_Offset_Spacing_Bytes
        lngProgressStepCount = lngProgressStepCount + 1
    End If
    
' 3. Write the UMC Data
    BinarySaveDataInitSection lngProgressStepCount, "Writing LC-MS Feature info", 1, lngOffsetsInFile(fioGelUMC), OutFileNum
    If blnDataChanged(fioGelUMC) And Not blnRemoveUMCData Then
        Put #OutFileNum, , GelUMC(lngGelIndex)
        
        lngProgressStepCount = lngProgressStepCount + 1
        frmProgress.UpdateProgressBar lngProgressStepCount
        
        blnDataChanged(fioGelUMC) = False
    Else
        If Loc(OutFileNum) < lngOffsetsInFile(fioGelUMC + 1) - FileIO_Offset_Spacing_Bytes Then
            Seek #OutFileNum, lngOffsetsInFile(fioGelUMC + 1) - FileIO_Offset_Spacing_Bytes
        Else
            ' This should only occur if blnRemoveUMCData = False
            ' Definitely now need to save analysis info
            blnDataChanged(fioGelUMC + 1) = True
        End If
        lngProgressStepCount = lngProgressStepCount + 1
    End If
    
' 4. Write the GelAnalysis Data
    BinarySaveDataInitSection lngProgressStepCount, "Writing Analysis Info", 1, lngOffsetsInFile(fioGelAnalysis), OutFileNum
    If blnDataChanged(fioGelAnalysis) Then
        GelAnalysisInfoWrite OutFileNum, lngGelIndex
        
        lngProgressStepCount = lngProgressStepCount + 1
        frmProgress.UpdateProgressBar lngProgressStepCount
        
        blnDataChanged(fioGelAnalysis) = False
    Else
        lngProgressStepCount = lngProgressStepCount + 1
    End If

' 5. Write GelUMCNETAdjDef
    BinarySaveDataInitSection lngProgressStepCount, "Writing UMC Net Adjustment parameters", 1, lngOffsetsInFile(fioUMCNetAdjDef), OutFileNum
    If blnDataChanged(fioUMCNetAdjDef) Then
        Put #OutFileNum, , GelUMCNETAdjDef(lngGelIndex)
        
        lngProgressStepCount = lngProgressStepCount + 1
        frmProgress.UpdateProgressBar lngProgressStepCount
        
        blnDataChanged(fioUMCNetAdjDef) = False
    Else
        If Loc(OutFileNum) < lngOffsetsInFile(fioUMCNetAdjDef + 1) - FileIO_Offset_Spacing_Bytes Then
            Seek #OutFileNum, lngOffsetsInFile(fioUMCNetAdjDef + 1) - FileIO_Offset_Spacing_Bytes
        Else
            ' This is unexpected; definitely now need to save the Gel Search Def
            Debug.Assert False
            blnDataChanged(fioUMCNetAdjDef + 1) = True
        End If
        lngProgressStepCount = lngProgressStepCount + 1
    End If

' 6. Write GelSearchDef
    BinarySaveDataInitSection lngProgressStepCount, "Writing Search Definitions", 1, lngOffsetsInFile(fioSearchDefinitions), OutFileNum
        
    ' Add an entry to the AnalysisHistory before writing to disk
    ' Use strFilePath since we may be saving to a network drive
    AddToAnalysisHistory lngGelIndex, "Saved file to disk (Program Version " & GetProgramVersion() & "; Path = " & strFilePath
    Put #OutFileNum, , GelSearchDef(lngGelIndex)
    
    lngProgressStepCount = lngProgressStepCount + 1
    frmProgress.UpdateProgressBar lngProgressStepCount
    
    blnDataChanged(fioUMCNetAdjDef) = False

' 7. Write ORF information (GelORFData is unused, so simply writing an Int32 zero, March 2006)
    BinarySaveDataInitSection lngProgressStepCount, "Writing Protein information", 1, lngOffsetsInFile(fioORFData), OutFileNum
    Put #OutFileNum, , CLng(0)
    lngProgressStepCount = lngProgressStepCount + 1
    frmProgress.UpdateProgressBar lngProgressStepCount
    
''    If blnDataChanged(fioORFData) And Not blnRemoveORFData Then
''        Put #OutFileNum, , GelORFData(lngGelIndex)
''
''        lngProgressStepCount = lngProgressStepCount + 1
''        frmProgress.UpdateProgressBar lngProgressStepCount
''
''        blnDataChanged(fioORFData) = False
''    Else
''        If Loc(OutFileNum) < lngOffsetsInFile(fioORFData + 1) - FileIO_Offset_Spacing_Bytes Then
''            Seek #OutFileNum, lngOffsetsInFile(fioORFData + 1) - FileIO_Offset_Spacing_Bytes
''        Else
''            ' This should only occur if blnRemoveORFData = False
''            ' Definitely now need to save ORF MT tag info
''            blnDataChanged(fioORFData + 1) = True
''        End If
''        lngProgressStepCount = lngProgressStepCount + 1
''    End If

' 8. Write ORF MT tag information (GelORFMassTags is unused, so simply writing an Int32 zero, March 2006)
    BinarySaveDataInitSection lngProgressStepCount, "Writing the MT tags for the Proteins", 1, lngOffsetsInFile(fioORFMassTags), OutFileNum
    Put #OutFileNum, , CLng(0)
    lngProgressStepCount = lngProgressStepCount + 1
    frmProgress.UpdateProgressBar lngProgressStepCount
    
''    If blnDataChanged(fioORFMassTags) And Not blnRemoveORFData Then
''        Put #OutFileNum, , GelORFMassTags(lngGelIndex)
''
''        lngProgressStepCount = lngProgressStepCount + 1
''        frmProgress.UpdateProgressBar lngProgressStepCount
''
''        blnDataChanged(fioORFMassTags) = False
''    Else
''        If Loc(OutFileNum) < lngOffsetsInFile(fioORFMassTags + 1) - FileIO_Offset_Spacing_Bytes Then
''            Seek #OutFileNum, lngOffsetsInFile(fioORFMassTags + 1) - FileIO_Offset_Spacing_Bytes
''        Else
''            ' This should only occur if blnRemoveORFData = False
''            ' Definitely now need to save Pairs info
''            blnDataChanged(fioORFMassTags + 1) = True
''        End If
''        lngProgressStepCount = lngProgressStepCount + 1
''    End If


' 9. Write the Pairs Data (GelP is unused, so simply writing an Int32 zero, March 2006)
    BinarySaveDataInitSection lngProgressStepCount, "Writing Pairs", 1, lngOffsetsInFile(fioGelPairs), OutFileNum
    Put #OutFileNum, , CLng(0)
    lngProgressStepCount = lngProgressStepCount + 1
    frmProgress.UpdateProgressBar lngProgressStepCount
    
''    If blnDataChanged(fioGelPairs) And Not blnRemovePairsData Then
''        Put #OutFileNum, , GelP(lngGelIndex)
''
''        lngProgressStepCount = lngProgressStepCount + 1
''        frmProgress.UpdateProgressBar lngProgressStepCount
''
''        blnDataChanged(fioGelPairs) = False
''    Else
''        lngProgressStepCount = lngProgressStepCount + 1
''    End If
    
    
' 10. Write the Delta Labeled Pairs Data
    BinarySaveDataInitSection lngProgressStepCount, "Writing Delta Labeled Pairs", 1, lngOffsetsInFile(fioGelDeltaLabeledPairs), OutFileNum
    If blnDataChanged(fioGelDeltaLabeledPairs) And Not blnRemovePairsData Then
        Put #OutFileNum, , GelP_D_L(lngGelIndex)
        
        lngProgressStepCount = lngProgressStepCount + 1
        frmProgress.UpdateProgressBar lngProgressStepCount
        
        blnDataChanged(fioGelDeltaLabeledPairs) = False
    Else
        lngProgressStepCount = lngProgressStepCount + 1
    End If
    
' 11. Write the IDP Data (GelIDP is unused, so simply writing an Int32 zero, August 2003)
    BinarySaveDataInitSection lngProgressStepCount, "Writing IDP", 1, lngOffsetsInFile(fioGelIDP), OutFileNum
    Put #OutFileNum, , CLng(0)
    lngProgressStepCount = lngProgressStepCount + 1
    frmProgress.UpdateProgressBar lngProgressStepCount
    
''    If blnDataChanged(fioGelIDP) And Not blnRemovePairsData Then
''        Put #OutFileNum, , GelIDP(lngGelIndex)
''
''        lngProgressStepCount = lngProgressStepCount + 1
''        frmProgress.UpdateProgressBar lngProgressStepCount
''
''        blnDataChanged(fioGelIDP) = False
''    Else
''        lngProgressStepCount = lngProgressStepCount + 1
''    End If
    
' 12. Write the Gel Lock Mass (LM) Data
    BinarySaveDataInitSection lngProgressStepCount, "Writing GelLM", 1, lngOffsetsInFile(fioGelLM), OutFileNum
    If blnDataChanged(fioGelLM) And Not blnRemovePairsData Then
        Put #OutFileNum, , GelLM(lngGelIndex)
        
        lngProgressStepCount = lngProgressStepCount + 1
        frmProgress.UpdateProgressBar lngProgressStepCount
        
        blnDataChanged(fioGelLM) = False
    Else
        lngProgressStepCount = lngProgressStepCount + 1
    End If
    
' 13. Write the ORF Viewer Saved Gel List and Options
    BinarySaveDataInitSection lngProgressStepCount, "Writing ORF Viewer Options", 1, lngOffsetsInFile(fioGelORFViewerSavedGelListAndOptions), OutFileNum
    Put #OutFileNum, , CLng(0)
    lngProgressStepCount = lngProgressStepCount + 1
    frmProgress.UpdateProgressBar lngProgressStepCount
    
''    If blnDataChanged(fioGelORFViewerSavedGelListAndOptions) Then
''        Put #OutFileNum, , GelORFViewerSavedGelListAndOptions(lngGelIndex)
''
''        lngProgressStepCount = lngProgressStepCount + 1
''        frmProgress.UpdateProgressBar lngProgressStepCount
''
''        blnDataChanged(fioGelORFViewerSavedGelListAndOptions) = False
''    Else
''        lngProgressStepCount = lngProgressStepCount + 1
''    End If
    
    Close #OutFileNum


' 13. Reopen the file to write the actual offsets used

    Open strWorkingFilePath For Binary Access Write As #OutFileNum
    For intSectionID = 0 To FILEIO_SECTION_COUNT - 1
        Put #OutFileNum, FileInfoHeaderOffsets(intSectionID), lngOffsetsInFile(intSectionID)
    Next intSectionID
    
    Close #OutFileNum
    
    ' Possibly move strWorkingFilePath to strFilePath (objRemoteSaveFileHandler will have cached the paths)
    frmProgress.UpdateCurrentSubTask "Copying file from temporary folder to destination folder"
    blnSuccess = objRemoteSaveFileHandler.MoveTempFileToFinalDestination()
    
    GelBody(lngGelIndex).Caption = strFilePath
    GelStatus(lngGelIndex).GelFilePathFull = GetFilePathFull(strFilePath)
    
    UpdateFileMenu strFilePath
    
    BinarySaveData = blnSuccess
    
BinarySaveDataCleanup:
    If KeyPressAbortProcess > 1 Then
        GelStatus(lngGelIndex).Dirty = True
    Else
        GelStatus(lngGelIndex).Dirty = False
    End If
    
    On Error Resume Next
    Close #OutFileNum
    frmProgress.HideForm
    
    Set fso = Nothing
    
    Screen.MousePointer = vbNormal
    
    Exit Function
    
BinarySaveDataErrorHandler:
    If Not glbPreferencesExpanded.AutoAnalysisStatus.Enabled Then
        MsgBox "Error writing output file " & strFilePath & vbCrLf & Err.Description & vbCrLf & "Aborting.", vbExclamation + vbOKOnly, "Error"
    Else
        AddToAnalysisHistory lngGelIndex, "Error writing output file " & strFilePath & "; " & Err.Description
    End If
    LogErrors Err.Number, "BinarySaveData", Err.Description, lngGelIndex
    BinarySaveData = False
    Resume BinarySaveDataCleanup
    
End Function
    
Private Sub BinarySaveDataInitOffsetInfo(lngInfoOffset As Long, lngActualOffsetOfData As Long, lngRecordSize As Long, lngDataCount As Long, sngDataFormatVersion As Single, OutFileNum As Integer)

    Seek #OutFileNum, lngInfoOffset
    Put #OutFileNum, , lngActualOffsetOfData
    Put #OutFileNum, , lngRecordSize
    Put #OutFileNum, , lngDataCount
    Put #OutFileNum, , sngDataFormatVersion
End Sub

Private Sub BinarySaveDataInitSection(lngOverallProgressValue As Long, strProgressDescription As String, lngProgressCountPredicted As Long, ByRef lngFileOffsetByte As Long, OutFileNum As Integer)
    Dim bytZero As Byte
    
    frmProgress.UpdateProgressBar lngOverallProgressValue
    frmProgress.InitializeSubtask strProgressDescription, 0, lngProgressCountPredicted
    
    ' Add FileIO_Offset_Spacing_Bytes to the current file location to get a starting value for lngFileOffsetByte
    lngFileOffsetByte = Loc(OutFileNum) + FileIO_Offset_Spacing_Bytes

    ' Round lngFileOffsetByte up to the nearest 1000
    lngFileOffsetByte = RoundToNearest(lngFileOffsetByte, 1000, True)
    Debug.Assert Loc(OutFileNum) < lngFileOffsetByte
    
    ' Write zeroes from current position to new position
    ' Must subtract 1 from lngFileOffSetByte to avoid overwriting the final byte
    bytZero = 0
    Do While Loc(OutFileNum) < lngFileOffsetByte - 1
        Put #OutFileNum, , bytZero
    Loop
    
    Seek #OutFileNum, lngFileOffsetByte
    
End Sub

Private Sub CopyAMTSearchDef2002ToCurrent(ByRef OldDef As SearchAMTDefinition2002, ByRef CurrentSearchDef As SearchAMTDefinition)
    
    With CurrentSearchDef
        .SearchScope = OldDef.SearchScope
        .SearchFlag = OldDef.SearchFlag
        .MWField = OldDef.MWField
        .TolType = OldDef.TolType
        .NETorRT = OldDef.NETorRT
        .Formula = OldDef.Formula
        .MWTol = OldDef.MWTol
        .NETTol = OldDef.NETTol
        .MassTag = OldDef.MassTag
        .MaxMassTags = OldDef.MaxMassTags
        .SkipReferenced = OldDef.SkipReferenced
        .SaveNCnt = OldDef.SaveNCnt
        .UseDriftTime = False
        .DriftTimeTol = DEFAULT_DRIFT_TIME_TOL
        .AdditionalValue1 = 0
        .AdditionalValue2 = 0
        .AdditionalValue3 = 0
        .AdditionalValue4 = 0
        .AdditionalValue5 = 0
        .AdditionalValue6 = 0
        .AdditionalValue7 = 0
        .AdditionalValue8 = 0
    End With

End Sub

Private Sub CopyDeltaLabelPairDetails2004bToCurrent(ByVal OldPairDetailsCount As Long, ByRef OldDeltaLabelPairDetails() As udtIsoPairsDetails2004bType, ByRef CurrentDeltaLabelPairs As IsoPairsDltLblType)
    Dim i As Long
    Dim j As Integer
    Dim MaxInd As Long
    
    With CurrentDeltaLabelPairs
        .PCnt = OldPairDetailsCount
        
        If OldPairDetailsCount > 0 Then
            MaxInd = UBound(OldDeltaLabelPairDetails)
            If MaxInd > 0 Then
                ReDim .Pairs(MaxInd)
                For i = 0 To MaxInd
                    With .Pairs(i)
                        .p1 = OldDeltaLabelPairDetails(i).p1
                        .P1LblCnt = OldDeltaLabelPairDetails(i).P1LblCnt
                        .p2 = OldDeltaLabelPairDetails(i).p2
                        .P2DltCnt = OldDeltaLabelPairDetails(i).P2DltCnt
                        .P2LblCnt = OldDeltaLabelPairDetails(i).P2LblCnt
                        .ER = OldDeltaLabelPairDetails(i).ER
                        .ERStDev = OldDeltaLabelPairDetails(i).ERStDev
                        .ERChargeStateBasisCount = OldDeltaLabelPairDetails(i).ERChargeStateBasisCount
                        
                        If .ERChargeStateBasisCount = 0 Then
                            ReDim .ERChargesUsed(0)
                        Else
                            ReDim .ERChargesUsed(UBound(OldDeltaLabelPairDetails(i).ERChargesUsed))
                            For j = 0 To UBound(OldDeltaLabelPairDetails(i).ERChargesUsed)
                                .ERChargesUsed(j) = OldDeltaLabelPairDetails(i).ERChargesUsed(j)
                            Next j
                        End If
                        
                        .ERMemberBasisCount = OldDeltaLabelPairDetails(i).ERMemberBasisCount
                        .STATE = OldDeltaLabelPairDetails(i).STATE
                    End With
                Next i
            Else
                ReDim .Pairs(0)
            End If
        Else
            ReDim .Pairs(0)
        End If
    End With

End Sub

Private Sub CopyDeltaLabelPairs2003ToCurrent(OldDeltaLabelPairs As IsoPairsDltLbl2003Type, ByRef CurrentDeltaLabelPairs As IsoPairsDltLblType)
    Dim MaxInd As Long
    Dim i As Long
    
    On Error GoTo CopyDeltaLabelPairsErrorHandler
    
    With CurrentDeltaLabelPairs
        .SyncWithUMC = OldDeltaLabelPairs.SyncWithUMC
        .DltLblType = OldDeltaLabelPairs.DltLblType
        
        .SearchDef = glbPreferencesExpanded.PairSearchOptions.SearchDef
        With .SearchDef
            .LightLabelMass = OldDeltaLabelPairs.LblMW
            .DeltaMass = OldDeltaLabelPairs.DltMW
            
            If .DeltaMass = glO16O18_DELTA Then
                .DeltaCountMin = 1
                .DeltaCountMax = 1
                .DeltaStepSize = 1
            ElseIf .DeltaMass = glN14N15_DELTA Then
                If .DeltaCountMin = 1 And .DeltaCountMax = 1 Then
                    .DeltaCountMin = 1
                    .DeltaCountMax = 100
                End If
            End If
            
            .ERCalcType = OldDeltaLabelPairs.ERCalcType
            
            .RequireMatchingChargeStatesForPairMembers = True
            .UseIdenticalChargesForER = True
            .ComputeERScanByScan = False
            .ScanByScanAverageIsNotWeighted = False
            
            .RequireMatchingIsotopeTagLabels = False
            
            .MonoPlusMinusThresholdForceHeavyOrLight = 66
            .IgnoreMonoPlus2AbundanceInIReportERCalc = 0
            
            .AverageERsAllChargeStates = False
            .AverageERsWeightingMode = aewAbundance
        End With
        
        .PCnt = OldDeltaLabelPairs.PCnt
        
        If OldDeltaLabelPairs.PCnt > 0 Then
            MaxInd = UBound(OldDeltaLabelPairs.p1)
            If MaxInd > 0 Then
                ReDim .Pairs(MaxInd)
                For i = 0 To MaxInd
                    With .Pairs(i)
                        .p1 = OldDeltaLabelPairs.p1(i)
                        .P1LblCnt = OldDeltaLabelPairs.P1LblCnt(i)
                        .p2 = OldDeltaLabelPairs.p2(i)
                        .P2DltCnt = OldDeltaLabelPairs.P2DltCnt(i)
                        .P2LblCnt = OldDeltaLabelPairs.P2LblCnt(i)
                        .ER = OldDeltaLabelPairs.P1P2ER(i)
                        .ERStDev = 0
                        .ERChargeStateBasisCount = 0
                        ReDim .ERChargesUsed(0)
                        .ERMemberBasisCount = 1
                        .STATE = OldDeltaLabelPairs.PState(i)
                    End With
                Next i
            Else
                ReDim .Pairs(0)
            End If
        Else
            ReDim .Pairs(0)
        End If
        
        .OtherInfo = ""
    End With
    Exit Sub
    
CopyDeltaLabelPairsErrorHandler:
    Debug.Assert False
    Resume Next
    
End Sub

Private Sub CopyDeltaLabelPairs2004aToCurrent(OldDeltaLabelPairs As IsoPairsDltLbl2004aType, ByRef CurrentDeltaLabelPairs As IsoPairsDltLblType)
    Dim MaxInd As Long
    Dim i As Long
    
    On Error GoTo CopyDeltaLabelPairsErrorHandler
    
    With CurrentDeltaLabelPairs
        .SyncWithUMC = OldDeltaLabelPairs.SyncWithUMC
        .DltLblType = OldDeltaLabelPairs.DltLblType
        
        .SearchDef = glbPreferencesExpanded.PairSearchOptions.SearchDef
        With .SearchDef
            .LightLabelMass = OldDeltaLabelPairs.LblMW
            .DeltaMass = OldDeltaLabelPairs.DltMW
            
            If .DeltaMass = glO16O18_DELTA Then
                .DeltaCountMin = 1
                .DeltaCountMax = 1
                .DeltaStepSize = 1
            ElseIf .DeltaMass = glN14N15_DELTA Then
                If .DeltaCountMin = 1 And .DeltaCountMax = 1 Then
                    .DeltaCountMin = 1
                    .DeltaCountMax = 100
                End If
            End If
            
            .ERCalcType = OldDeltaLabelPairs.ERCalcType
            
            .RequireMatchingChargeStatesForPairMembers = OldDeltaLabelPairs.RequireMatchingChargeStatesForPairMembers
            .UseIdenticalChargesForER = OldDeltaLabelPairs.UseIdenticalChargesForER
            .ComputeERScanByScan = OldDeltaLabelPairs.ComputeERScanByScan
            .ScanByScanAverageIsNotWeighted = False
            
            .RequireMatchingIsotopeTagLabels = False
            
            .MonoPlusMinusThresholdForceHeavyOrLight = 66
            .IgnoreMonoPlus2AbundanceInIReportERCalc = 0
            
            .AverageERsAllChargeStates = OldDeltaLabelPairs.AverageERsAllChargeStates
            .AverageERsWeightingMode = OldDeltaLabelPairs.AverageERsWeightingMode
        End With
        
        .PCnt = OldDeltaLabelPairs.PCnt
        
        If OldDeltaLabelPairs.PCnt > 0 Then
            MaxInd = UBound(OldDeltaLabelPairs.Pairs)
            If MaxInd > 0 Then
                ReDim .Pairs(MaxInd)
                For i = 0 To MaxInd
                    With .Pairs(i)
                        .p1 = OldDeltaLabelPairs.Pairs(i).p1
                        .P1LblCnt = OldDeltaLabelPairs.Pairs(i).P1LblCnt
                        .p2 = OldDeltaLabelPairs.Pairs(i).p2
                        .P2DltCnt = OldDeltaLabelPairs.Pairs(i).P2DltCnt
                        .P2LblCnt = OldDeltaLabelPairs.Pairs(i).P2LblCnt
                        .ER = OldDeltaLabelPairs.Pairs(i).ER
                        .ERStDev = OldDeltaLabelPairs.Pairs(i).ERStDev
                        .ERChargeStateBasisCount = OldDeltaLabelPairs.Pairs(i).ERChargeStateBasisCount
                        ReDim .ERChargesUsed(0)
                        .ERMemberBasisCount = OldDeltaLabelPairs.Pairs(i).ERMemberBasisCount
                        .STATE = OldDeltaLabelPairs.Pairs(i).STATE
                    End With
                Next i
            Else
                ReDim .Pairs(0)
            End If
        Else
            ReDim .Pairs(0)
        End If
        
        .OtherInfo = ""
    End With
    
    Exit Sub
    
CopyDeltaLabelPairsErrorHandler:
    Debug.Assert False
    Resume Next
    
End Sub

Private Sub CopyDeltaLabelPairs2004bToCurrent(OldDeltaLabelPairs As IsoPairsDltLbl2004bType, ByRef CurrentDeltaLabelPairs As IsoPairsDltLblType)
    
    On Error GoTo CopyDeltaLabelPairsErrorHandler
    
    With CurrentDeltaLabelPairs
        .SyncWithUMC = OldDeltaLabelPairs.SyncWithUMC
        .DltLblType = OldDeltaLabelPairs.DltLblType
        
        .SearchDef = glbPreferencesExpanded.PairSearchOptions.SearchDef
        With .SearchDef
            .LightLabelMass = OldDeltaLabelPairs.LblMW
            .DeltaMass = OldDeltaLabelPairs.DltMW
            
            If .DeltaMass = glO16O18_DELTA Then
                .DeltaCountMin = 1
                .DeltaCountMax = 1
                .DeltaStepSize = 1
            ElseIf .DeltaMass = glN14N15_DELTA Then
                If .DeltaCountMin = 1 And .DeltaCountMax = 1 Then
                    .DeltaCountMin = 1
                    .DeltaCountMax = 100
                End If
            End If
            
            .ERCalcType = OldDeltaLabelPairs.ERCalcType
            
            .RequireMatchingChargeStatesForPairMembers = OldDeltaLabelPairs.RequireMatchingChargeStatesForPairMembers
            .UseIdenticalChargesForER = OldDeltaLabelPairs.UseIdenticalChargesForER
            .ComputeERScanByScan = OldDeltaLabelPairs.ComputeERScanByScan
            .ScanByScanAverageIsNotWeighted = False
            
            .RequireMatchingIsotopeTagLabels = False
            
            .MonoPlusMinusThresholdForceHeavyOrLight = 66
            .IgnoreMonoPlus2AbundanceInIReportERCalc = 0
            
            .AverageERsAllChargeStates = OldDeltaLabelPairs.AverageERsAllChargeStates
            .AverageERsWeightingMode = OldDeltaLabelPairs.AverageERsWeightingMode
        End With
        
        ' Note: .PCnt and .Pairs are copied by CopyDeltaLabelPairDetails2004bToCurrent
        .OtherInfo = ""
    End With
    
    CopyDeltaLabelPairDetails2004bToCurrent OldDeltaLabelPairs.PCnt, OldDeltaLabelPairs.Pairs, CurrentDeltaLabelPairs
    
    Exit Sub
    
CopyDeltaLabelPairsErrorHandler:
    Debug.Assert False
    Resume Next
    
End Sub

Private Sub CopyDeltaLabelPairs2004cToCurrent(OldDeltaLabelPairs As IsoPairsDltLbl2004cType, ByRef CurrentDeltaLabelPairs As IsoPairsDltLblType)
    
    On Error GoTo CopyDeltaLabelPairsErrorHandler
    
    With CurrentDeltaLabelPairs
        .SyncWithUMC = OldDeltaLabelPairs.SyncWithUMC
        .DltLblType = OldDeltaLabelPairs.DltLblType
        
        .SearchDef = glbPreferencesExpanded.PairSearchOptions.SearchDef
        
        With .SearchDef
            .DeltaMass = OldDeltaLabelPairs.SearchDef.DeltaMass
            .DeltaMassTolerance = OldDeltaLabelPairs.SearchDef.DeltaMassTolerance
            .AutoCalculateDeltaMinMaxCount = OldDeltaLabelPairs.SearchDef.AutoCalculateDeltaMinMaxCount
            .DeltaCountMin = OldDeltaLabelPairs.SearchDef.DeltaCountMin
            .DeltaCountMax = OldDeltaLabelPairs.SearchDef.DeltaCountMax
            .DeltaStepSize = 1
            
            .LightLabelMass = OldDeltaLabelPairs.SearchDef.LightLabelMass
            .HeavyLightMassDifference = OldDeltaLabelPairs.SearchDef.HeavyLightMassDifference
            .LabelCountMin = OldDeltaLabelPairs.SearchDef.LabelCountMin
            .LabelCountMax = OldDeltaLabelPairs.SearchDef.LabelCountMax
            .MaxDifferenceInNumberOfLightHeavyLabels = OldDeltaLabelPairs.SearchDef.MaxDifferenceInNumberOfLightHeavyLabels
            
            .RequireUMCOverlap = OldDeltaLabelPairs.SearchDef.RequireUMCOverlap
            .RequireUMCOverlapAtApex = OldDeltaLabelPairs.SearchDef.RequireUMCOverlapAtApex
            
            .ScanTolerance = OldDeltaLabelPairs.SearchDef.ScanTolerance
            .ScanToleranceAtApex = OldDeltaLabelPairs.SearchDef.ScanToleranceAtApex
            
            .ERInclusionMin = OldDeltaLabelPairs.SearchDef.ERInclusionMin
            .ERInclusionMax = OldDeltaLabelPairs.SearchDef.ERInclusionMax
            
            .RequireMatchingChargeStatesForPairMembers = OldDeltaLabelPairs.SearchDef.RequireMatchingChargeStatesForPairMembers
            .UseIdenticalChargesForER = OldDeltaLabelPairs.SearchDef.UseIdenticalChargesForER
            .ComputeERScanByScan = OldDeltaLabelPairs.SearchDef.ComputeERScanByScan
            .ScanByScanAverageIsNotWeighted = False
            
            .RequireMatchingIsotopeTagLabels = False
            
            .MonoPlusMinusThresholdForceHeavyOrLight = 66
            .IgnoreMonoPlus2AbundanceInIReportERCalc = 0
            
            .AverageERsAllChargeStates = OldDeltaLabelPairs.SearchDef.AverageERsAllChargeStates
            .AverageERsWeightingMode = OldDeltaLabelPairs.SearchDef.AverageERsWeightingMode
            
            .ERCalcType = OldDeltaLabelPairs.SearchDef.ERCalcType
             
            ' IReport and RemoveOutlier options were initialized when we copied the .SearchDef from glbPreferencesExpanded
             
             .OtherInfo = ""
        End With
        
        ' Note: .PCnt and .Pairs are copied by CopyDeltaLabelPairDetails2004bToCurrent
        .OtherInfo = ""
    End With
    
    CopyDeltaLabelPairDetails2004bToCurrent OldDeltaLabelPairs.PCnt, OldDeltaLabelPairs.Pairs, CurrentDeltaLabelPairs
    
    Exit Sub
    
CopyDeltaLabelPairsErrorHandler:
    Debug.Assert False
    Resume Next

End Sub

Private Sub CopyDeltaLabelPairs2004dToCurrent(OldDeltaLabelPairs As IsoPairsDltLbl2004dType, ByRef CurrentDeltaLabelPairs As IsoPairsDltLblType)
    
    On Error GoTo CopyDeltaLabelPairsErrorHandler
    
    With CurrentDeltaLabelPairs
        .SyncWithUMC = OldDeltaLabelPairs.SyncWithUMC
        .DltLblType = OldDeltaLabelPairs.DltLblType
        
        .SearchDef = glbPreferencesExpanded.PairSearchOptions.SearchDef
        
        With .SearchDef
            .DeltaMass = OldDeltaLabelPairs.SearchDef.DeltaMass
            .DeltaMassTolerance = OldDeltaLabelPairs.SearchDef.DeltaMassTolerance
            .AutoCalculateDeltaMinMaxCount = OldDeltaLabelPairs.SearchDef.AutoCalculateDeltaMinMaxCount
            .DeltaCountMin = OldDeltaLabelPairs.SearchDef.DeltaCountMin
            .DeltaCountMax = OldDeltaLabelPairs.SearchDef.DeltaCountMax
            .DeltaStepSize = 1
            
            .LightLabelMass = OldDeltaLabelPairs.SearchDef.LightLabelMass
            .HeavyLightMassDifference = OldDeltaLabelPairs.SearchDef.HeavyLightMassDifference
            .LabelCountMin = OldDeltaLabelPairs.SearchDef.LabelCountMin
            .LabelCountMax = OldDeltaLabelPairs.SearchDef.LabelCountMax
            .MaxDifferenceInNumberOfLightHeavyLabels = OldDeltaLabelPairs.SearchDef.MaxDifferenceInNumberOfLightHeavyLabels
            
            .RequireUMCOverlap = OldDeltaLabelPairs.SearchDef.RequireUMCOverlap
            .RequireUMCOverlapAtApex = OldDeltaLabelPairs.SearchDef.RequireUMCOverlapAtApex
            
            .ScanTolerance = OldDeltaLabelPairs.SearchDef.ScanTolerance
            .ScanToleranceAtApex = OldDeltaLabelPairs.SearchDef.ScanToleranceAtApex
            
            .ERInclusionMin = OldDeltaLabelPairs.SearchDef.ERInclusionMin
            .ERInclusionMax = OldDeltaLabelPairs.SearchDef.ERInclusionMax
            
            .RequireMatchingChargeStatesForPairMembers = OldDeltaLabelPairs.SearchDef.RequireMatchingChargeStatesForPairMembers
            .UseIdenticalChargesForER = OldDeltaLabelPairs.SearchDef.UseIdenticalChargesForER
            .ComputeERScanByScan = OldDeltaLabelPairs.SearchDef.ComputeERScanByScan
            .ScanByScanAverageIsNotWeighted = False
            
            .RequireMatchingIsotopeTagLabels = False
            
            .MonoPlusMinusThresholdForceHeavyOrLight = 66
            .IgnoreMonoPlus2AbundanceInIReportERCalc = 0
            
            .AverageERsAllChargeStates = OldDeltaLabelPairs.SearchDef.AverageERsAllChargeStates
            .AverageERsWeightingMode = OldDeltaLabelPairs.SearchDef.AverageERsWeightingMode
            
            .ERCalcType = OldDeltaLabelPairs.SearchDef.ERCalcType
             
            .IReportEROptions = OldDeltaLabelPairs.SearchDef.IReportEROptions
            
            .RemoveOutlierERs = OldDeltaLabelPairs.SearchDef.RemoveOutlierERs
            .RemoveOutlierERsIterate = OldDeltaLabelPairs.SearchDef.RemoveOutlierERsIterate
            .RemoveOutlierERsMinimumDataPointCount = OldDeltaLabelPairs.SearchDef.RemoveOutlierERsMinimumDataPointCount
            .RemoveOutlierERsConfidenceLevel = OldDeltaLabelPairs.SearchDef.RemoveOutlierERsConfidenceLevel
    
            .OtherInfo = OldDeltaLabelPairs.OtherInfo
        End With
        
        ' Note: .PCnt and .Pairs are copied by CopyDeltaLabelPairDetails2004bToCurrent
        .OtherInfo = ""
    End With
    
    CopyDeltaLabelPairDetails2004bToCurrent OldDeltaLabelPairs.PCnt, OldDeltaLabelPairs.Pairs, CurrentDeltaLabelPairs
        
    Exit Sub
    
CopyDeltaLabelPairsErrorHandler:
    Debug.Assert False
    Resume Next

End Sub

Private Sub CopyDeltaLabelPairs2004eToCurrent(OldDeltaLabelPairs As IsoPairsDltLbl2004eType, ByRef CurrentDeltaLabelPairs As IsoPairsDltLblType)
    
    On Error GoTo CopyDeltaLabelPairsErrorHandler
    
    With CurrentDeltaLabelPairs
        .SyncWithUMC = OldDeltaLabelPairs.SyncWithUMC
        .DltLblType = OldDeltaLabelPairs.DltLblType
        
        .SearchDef = glbPreferencesExpanded.PairSearchOptions.SearchDef
        
        With .SearchDef
            .DeltaMass = OldDeltaLabelPairs.SearchDef.DeltaMass
            .DeltaMassTolerance = OldDeltaLabelPairs.SearchDef.DeltaMassTolerance
            .DeltaMassTolType = gltABS
            
            .AutoCalculateDeltaMinMaxCount = OldDeltaLabelPairs.SearchDef.AutoCalculateDeltaMinMaxCount
            .DeltaCountMin = OldDeltaLabelPairs.SearchDef.DeltaCountMin
            .DeltaCountMax = OldDeltaLabelPairs.SearchDef.DeltaCountMax
            .DeltaStepSize = OldDeltaLabelPairs.SearchDef.DeltaStepSize
            
            .LightLabelMass = OldDeltaLabelPairs.SearchDef.LightLabelMass
            .HeavyLightMassDifference = OldDeltaLabelPairs.SearchDef.HeavyLightMassDifference
            .LabelCountMin = OldDeltaLabelPairs.SearchDef.LabelCountMin
            .LabelCountMax = OldDeltaLabelPairs.SearchDef.LabelCountMax
            .MaxDifferenceInNumberOfLightHeavyLabels = OldDeltaLabelPairs.SearchDef.MaxDifferenceInNumberOfLightHeavyLabels
            
            .RequireUMCOverlap = OldDeltaLabelPairs.SearchDef.RequireUMCOverlap
            .RequireUMCOverlapAtApex = OldDeltaLabelPairs.SearchDef.RequireUMCOverlapAtApex
            
            .ScanTolerance = OldDeltaLabelPairs.SearchDef.ScanTolerance
            .ScanToleranceAtApex = OldDeltaLabelPairs.SearchDef.ScanToleranceAtApex
            
            .ERInclusionMin = OldDeltaLabelPairs.SearchDef.ERInclusionMin
            .ERInclusionMax = OldDeltaLabelPairs.SearchDef.ERInclusionMax
            
            .RequireMatchingChargeStatesForPairMembers = OldDeltaLabelPairs.SearchDef.RequireMatchingChargeStatesForPairMembers
            .UseIdenticalChargesForER = OldDeltaLabelPairs.SearchDef.UseIdenticalChargesForER
            .ComputeERScanByScan = OldDeltaLabelPairs.SearchDef.ComputeERScanByScan
            .ScanByScanAverageIsNotWeighted = False
            
            .RequireMatchingIsotopeTagLabels = False
            
            .MonoPlusMinusThresholdForceHeavyOrLight = 66
            .IgnoreMonoPlus2AbundanceInIReportERCalc = 0
            
            .AverageERsAllChargeStates = OldDeltaLabelPairs.SearchDef.AverageERsAllChargeStates
            .AverageERsWeightingMode = OldDeltaLabelPairs.SearchDef.AverageERsWeightingMode
            
            .ERCalcType = OldDeltaLabelPairs.SearchDef.ERCalcType
             
            .IReportEROptions = OldDeltaLabelPairs.SearchDef.IReportEROptions
            
            .RemoveOutlierERs = OldDeltaLabelPairs.SearchDef.RemoveOutlierERs
            .RemoveOutlierERsIterate = OldDeltaLabelPairs.SearchDef.RemoveOutlierERsIterate
            .RemoveOutlierERsMinimumDataPointCount = OldDeltaLabelPairs.SearchDef.RemoveOutlierERsMinimumDataPointCount
            .RemoveOutlierERsConfidenceLevel = OldDeltaLabelPairs.SearchDef.RemoveOutlierERsConfidenceLevel
    
            ' The N15 Incorporation options were initialized when we copied the .SearchDef from glbPreferencesExpanded

            .OtherInfo = OldDeltaLabelPairs.SearchDef.OtherInfo
        End With
        

        ' Note: .PCnt and .Pairs are copied by CopyDeltaLabelPairDetails2004bToCurrent
        .OtherInfo = ""
    End With
    
    CopyDeltaLabelPairDetails2004bToCurrent OldDeltaLabelPairs.PCnt, OldDeltaLabelPairs.Pairs, CurrentDeltaLabelPairs
        
    Exit Sub
    
CopyDeltaLabelPairsErrorHandler:
    Debug.Assert False
    Resume Next

End Sub

Private Sub CopyGelData2000ToCurrent(OldData As DocumentData2000, ByRef CurrentGelData As DocumentData)
    Dim MaxInd As Long
    Dim i As Long, j As Long
    
On Error GoTo CopyGelDataErrorHandler

    ' Transfer old structure to the new structure
    With CurrentGelData
        .Certificate = glCERT2003_Modular
        .Comment = OldData.Comment
        .FileName = OldData.FileName
        .Fileinfo = OldData.Fileinfo
        .PathtoDataFiles = OldData.PathtoDataFiles
        .PathtoDatabase = OldData.PathtoDatabase
        
        .MediaType = OldData.MediaType
        .LinesRead = OldData.LinesRead
        .DataLines = OldData.DataLines
        .CSLines = OldData.CSLines
        .IsoLines = OldData.IsoLines
        .CalEquation = OldData.CalEquation
        
        For i = 1 To 10
            .CalArg(i) = OldData.CalArg(i)
        Next i
        
        .MinMW = OldData.MinMW
        .MaxMW = OldData.MaxMW
        .MinAbu = OldData.MinAbu
        .MaxAbu = OldData.MaxAbu
        .Preferences = OldData.Preferences
        .pICooSysEnabled = OldData.pICooSysEnabled
        .DataStatusBits = 0
        
        MaxInd = UBound(OldData.DFFN)   'first index is always 1
        If MaxInd > 0 Then
            ReDim .ScanInfo(MaxInd)
            
            For i = 1 To MaxInd
                With .ScanInfo(i)
                    .ScanNumber = OldData.DFFN(i)
                    .ElutionTime = 0
                    .ScanType = 1
                    
                    .ScanFileName = OldData.DFN(i)
                    .ScanPI = OldData.DFPI(i)
                    
                    .TimeDomainSignal = OldData.DFIN(i)
                    .FrequencyShift = OldData.DFFS(i)
                End With
            Next i
        End If
        For i = 1 To 8                                      'copy old filter
            For j = 0 To 2
                .DataFilter(i, j) = OldData.DataFilter(i, j)
            Next j
        Next i
        'add missing filters
        .DataFilter(fltCSMW, 0) = False
        .DataFilter(fltCSMW, 1) = 0
        .DataFilter(fltCSMW, 2) = .MaxMW
        .DataFilter(fltIsoMW, 0) = False
        .DataFilter(fltIsoMW, 1) = 0
        .DataFilter(fltIsoMW, 2) = .MaxMW
        .DataFilter(fltIsoCS, 0) = False
        .DataFilter(fltIsoCS, 1) = 0
        .DataFilter(fltIsoCS, 2) = 1000
        .DataFilter(fltCSStDev, 1) = 1
        
        If .CSLines > 0 Then
            ReDim .CSData(.CSLines)
            For i = 1 To .CSLines
                CopyLegacyCSToIsoData .CSData(i), OldData.CSNum, OldData.CSVar, i
            Next i
        End If
        If .IsoLines > 0 Then
            ReDim .IsoData(.IsoLines)
            For i = 1 To .IsoLines
                ' GelData2000 has 10 entries for .IsoNum
                CopyLegacyIsoToIsoData .IsoData(i), OldData.IsoNum, OldData.IsoVar, i
            Next i
        End If
    End With
    Exit Sub
    
CopyGelDataErrorHandler:
    Debug.Assert False

End Sub

Private Sub CopyGelData2003ToCurrent(OldData As DocumentData2003, ByRef CurrentGelData As DocumentData)
    Dim MaxInd As Long
    Dim i As Long, j As Long
    
On Error GoTo CopyGelDataErrorHandler

    ' Transfer old structure to the new structure
    With CurrentGelData
        .Certificate = glCERT2003_Modular
        .Comment = OldData.Comment
        .FileName = OldData.FileName
        .Fileinfo = OldData.Fileinfo
        .PathtoDataFiles = OldData.PathtoDataFiles
        .PathtoDatabase = OldData.PathtoDatabase
        
        .MediaType = OldData.MediaType
        .LinesRead = OldData.LinesRead
        .DataLines = OldData.DataLines
        .CSLines = OldData.CSLines
        .IsoLines = OldData.IsoLines
        .CalEquation = OldData.CalEquation
        
        For i = 1 To 10
            .CalArg(i) = OldData.CalArg(i)
        Next i
        
        .MinMW = OldData.MinMW
        .MaxMW = OldData.MaxMW
        .MinAbu = OldData.MinAbu
        .MaxAbu = OldData.MaxAbu
        .Preferences = OldData.Preferences
        .pICooSysEnabled = OldData.pICooSysEnabled
        .DataStatusBits = 0      ' New for this version
        
        For i = 1 To MAX_FILTER_COUNT_2003
            For j = 0 To 2
                .DataFilter(i, j) = OldData.DataFilter(i, j)
            Next j
        Next i
        
        MaxInd = UBound(OldData.DFFN)   'first index is always 1
        If MaxInd > 0 Then
            ReDim .ScanInfo(MaxInd)
            
            For i = 1 To MaxInd
                With .ScanInfo(i)
                    .ScanNumber = OldData.DFFN(i)
                    .ElutionTime = 0
                    .ScanType = 1
                    
                    .ScanFileName = OldData.DFN(i)
                    .ScanPI = OldData.DFPI(i)
                    
                    .TimeDomainSignal = OldData.DFIN(i)
                    .FrequencyShift = OldData.DFFS(i)
                End With
            Next i
        End If
        
        If .CSLines > 0 Then
            ReDim .CSData(.CSLines)
            For i = 1 To .CSLines
                CopyLegacyCSToIsoData .CSData(i), OldData.CSNum, OldData.CSVar, i
            Next i
        End If
        If .IsoLines > 0 Then
            ReDim .IsoData(.IsoLines)
            For i = 1 To .IsoLines
                ' GelData2003 has 10 entries for .IsoNum
                CopyLegacyIsoToIsoData .IsoData(i), OldData.IsoNum, OldData.IsoVar, i
            Next i
        End If
    End With
    Exit Sub
    
CopyGelDataErrorHandler:
    Debug.Assert False

End Sub

Private Sub CopyGelData2003bToCurrent(OldData As DocumentData2003b, ByRef CurrentGelData As DocumentData)
    Dim MaxInd As Long
    Dim i As Long, j As Long

On Error GoTo CopyGelDataErrorHandler

    ' Transfer old structure to the new structure
    With CurrentGelData
        .Certificate = glCERT2003_Modular
        .Comment = OldData.Comment
        .FileName = OldData.FileName
        .Fileinfo = OldData.Fileinfo
        .PathtoDataFiles = OldData.PathtoDataFiles
        .PathtoDatabase = OldData.PathtoDatabase
        
        .MediaType = OldData.MediaType
        .LinesRead = OldData.LinesRead
        .DataLines = OldData.DataLines
        .CSLines = OldData.CSLines
        .IsoLines = OldData.IsoLines
        .CalEquation = OldData.CalEquation
        
        For i = 1 To 10
            .CalArg(i) = OldData.CalArg(i)
        Next i
        
        .MinMW = OldData.MinMW
        .MaxMW = OldData.MaxMW
        .MinAbu = OldData.MinAbu
        .MaxAbu = OldData.MaxAbu
        .Preferences = OldData.Preferences
        .pICooSysEnabled = OldData.pICooSysEnabled
        .DataStatusBits = OldData.DataStatusBits
        
        For i = 1 To MAX_FILTER_COUNT_2003b
            For j = 0 To 2
                .DataFilter(i, j) = OldData.DataFilter(i, j)
            Next j
        Next i
        
        MaxInd = UBound(OldData.DFFN)   'first index is always 1
        If MaxInd > 0 Then
            ReDim .ScanInfo(MaxInd)
            
            For i = 1 To MaxInd
                With .ScanInfo(i)
                    .ScanNumber = OldData.DFFN(i)
                    .ElutionTime = 0
                    .ScanType = 1
                    
                    .ScanFileName = OldData.DFN(i)
                    .ScanPI = OldData.DFPI(i)
                    
                    .TimeDomainSignal = OldData.DFIN(i)
                    .FrequencyShift = OldData.DFFS(i)
                End With
            Next i
        End If
        
        If .CSLines > 0 Then
            ReDim .CSData(.CSLines)
            For i = 1 To .CSLines
                CopyLegacyCSToIsoData .CSData(i), OldData.CSNum, OldData.CSVar, i
            Next i
        End If
        If .IsoLines > 0 Then
            ReDim .IsoData(.IsoLines)
            For i = 1 To .IsoLines
                ' GelData2003b has 12 entries for .IsoNum
                CopyLegacyIsoToIsoData .IsoData(i), OldData.IsoNum, OldData.IsoVar, i
            Next i
        End If
        
        .OtherInfo = OldData.OtherInfo
    End With
    Exit Sub
    
CopyGelDataErrorHandler:
    Debug.Assert False
    
End Sub

Private Sub CopyGelData2004ToCurrent(OldData As DocumentData2004, ByRef CurrentGelData As DocumentData)
    Dim MaxInd As Long
    Dim i As Long, j As Long

On Error GoTo CopyGelDataErrorHandler

    ' Transfer old structure to the new structure
    With CurrentGelData
        .Certificate = glCERT2003_Modular
        .Comment = OldData.Comment
        .FileName = OldData.FileName
        .Fileinfo = OldData.Fileinfo
        .PathtoDataFiles = OldData.PathtoDataFiles
        .PathtoDatabase = OldData.PathtoDatabase
        
        .MediaType = OldData.MediaType
        .LinesRead = OldData.LinesRead
        .DataLines = OldData.DataLines
        .CSLines = OldData.CSLines
        .IsoLines = OldData.IsoLines
        .CalEquation = OldData.CalEquation
        
        For i = 1 To 10
            .CalArg(i) = OldData.CalArg(i)
        Next i
        
        .MinMW = OldData.MinMW
        .MaxMW = OldData.MaxMW
        .MinAbu = OldData.MinAbu
        .MaxAbu = OldData.MaxAbu
        .Preferences = OldData.Preferences
        .pICooSysEnabled = OldData.pICooSysEnabled
        .DataStatusBits = OldData.DataStatusBits
        
        For i = 1 To MAX_FILTER_COUNT_2004
            For j = 0 To 2
                .DataFilter(i, j) = OldData.DataFilter(i, j)
            Next j
        Next i
        
        MaxInd = UBound(OldData.DFFN)   'first index is always 1
        If MaxInd > 0 Then
            ReDim .ScanInfo(MaxInd)
            
            For i = 1 To MaxInd
                With .ScanInfo(i)
                    .ScanNumber = OldData.ScanInfo(i).ScanNumber
                    If .ScanNumber = OldData.DFFN(i) Then
                        .ScanFileName = OldData.DFN(i)
                        .ScanPI = OldData.DFPI(i)
                    
                        .TimeDomainSignal = OldData.DFIN(i)
                        .FrequencyShift = OldData.DFFS(i)
                    
                    Else
                        ' The .ScanInfo arrays is not in parallel with .DFFN(); this is unexpected
                        Debug.Assert False
                    End If
        
                    .ElutionTime = OldData.ScanInfo(i).ElutionTime
                    .ScanType = 1
                End With
            Next i
        End If
         
        If .CSLines > 0 Then
            ReDim .CSData(.CSLines)
            For i = 1 To .CSLines
                CopyLegacyCSToIsoData .CSData(i), OldData.CSNum, OldData.CSVar, i
            Next i
        End If
        If .IsoLines > 0 Then
            ReDim .IsoData(.IsoLines)
            For i = 1 To .IsoLines
                ' GelData2004 has 12 entries for .IsoNum
                CopyLegacyIsoToIsoData .IsoData(i), OldData.IsoNum, OldData.IsoVar, i
            Next i
        End If
        
        .OtherInfo = OldData.OtherInfo
    End With
    Exit Sub
    
CopyGelDataErrorHandler:
    Debug.Assert False

End Sub

Private Sub CopyGelData2005aToCurrent(OldData As DocumentData2005a, ByRef CurrentGelData As DocumentData)
    Dim MaxInd As Long
    Dim i As Long, j As Long

On Error GoTo CopyGelDataErrorHandler

    ' Transfer old structure to the new structure
    With CurrentGelData
        .Certificate = glCERT2003_Modular
        .Comment = OldData.Comment
        .FileName = OldData.FileName
        .Fileinfo = OldData.Fileinfo
        .PathtoDataFiles = OldData.PathtoDataFiles
        .PathtoDatabase = OldData.PathtoDatabase
        
        .MediaType = OldData.MediaType
        .LinesRead = OldData.LinesRead
        .DataLines = OldData.DataLines
        .CSLines = OldData.CSLines
        .IsoLines = OldData.IsoLines
        .CalEquation = OldData.CalEquation
        
        For i = 1 To 10
            .CalArg(i) = OldData.CalArg(i)
        Next i
        
        .MinMW = OldData.MinMW
        .MaxMW = OldData.MaxMW
        .MinAbu = OldData.MinAbu
        .MaxAbu = OldData.MaxAbu
        .Preferences = OldData.Preferences
        .pICooSysEnabled = OldData.pICooSysEnabled
        .DataStatusBits = OldData.DataStatusBits
        
        For i = 1 To MAX_FILTER_COUNT_2004
            For j = 0 To 2
                .DataFilter(i, j) = OldData.DataFilter(i, j)
            Next j
        Next i
        
        .CustomNETsDefined = False
        ReDim .ScanInfo(UBound(OldData.ScanInfo))
        
        ' Copy the entire array
        .ScanInfo = OldData.ScanInfo
         
        ' Copy these element-by-element
        If .CSLines > 0 Then
            ReDim .CSData(.CSLines)
            For i = 1 To .CSLines
                CopyLegacyIso2005ToCurrentIso .CSData(i), OldData.CSData(i)
            Next i
        End If
        
        If .IsoLines > 0 Then
            ReDim .IsoData(.IsoLines)
            For i = 1 To .IsoLines
                ' GelData2000 has 10 entries for .IsoNum
                CopyLegacyIso2005ToCurrentIso .IsoData(i), OldData.IsoData(i)
            Next i
        End If
        
        .OtherInfo = OldData.OtherInfo
    End With
    
    Exit Sub
    
CopyGelDataErrorHandler:
    Debug.Assert False

End Sub

Private Sub CopyGelData2005bToCurrent(OldData As DocumentData2005b, ByRef CurrentGelData As DocumentData)
    Dim MaxInd As Long
    Dim i As Long, j As Long

On Error GoTo CopyGelDataErrorHandler

    ' Transfer old structure to the new structure
    With CurrentGelData
        .Certificate = glCERT2003_Modular
        .Comment = OldData.Comment
        .FileName = OldData.FileName
        .Fileinfo = OldData.Fileinfo
        .PathtoDataFiles = OldData.PathtoDataFiles
        .PathtoDatabase = OldData.PathtoDatabase
        
        .MediaType = OldData.MediaType
        .LinesRead = OldData.LinesRead
        .DataLines = OldData.DataLines
        .CSLines = OldData.CSLines
        .IsoLines = OldData.IsoLines
        .CalEquation = OldData.CalEquation
        
        For i = 1 To 10
            .CalArg(i) = OldData.CalArg(i)
        Next i
        
        .MinMW = OldData.MinMW
        .MaxMW = OldData.MaxMW
        .MinAbu = OldData.MinAbu
        .MaxAbu = OldData.MaxAbu
        .Preferences = OldData.Preferences
        .pICooSysEnabled = OldData.pICooSysEnabled
        .DataStatusBits = OldData.DataStatusBits
        
        For i = 1 To MAX_FILTER_COUNT_2005
            For j = 0 To 2
                .DataFilter(i, j) = OldData.DataFilter(i, j)
            Next j
        Next i
        
        .CustomNETsDefined = False
        ReDim .ScanInfo(UBound(OldData.ScanInfo))
        
        ' Copy the entire array
        .ScanInfo = OldData.ScanInfo
         
        ' Copy these element-by-element
        If .CSLines > 0 Then
            ReDim .CSData(.CSLines)
            For i = 1 To .CSLines
                CopyLegacyIso2005ToCurrentIso .CSData(i), OldData.CSData(i)
            Next i
        End If
        
        If .IsoLines > 0 Then
            ReDim .IsoData(.IsoLines)
            For i = 1 To .IsoLines
                ' GelData2000 has 10 entries for .IsoNum
                CopyLegacyIso2005ToCurrentIso .IsoData(i), OldData.IsoData(i)
            Next i
        End If
        
        .OtherInfo = OldData.OtherInfo
    End With
    
    Exit Sub
    
CopyGelDataErrorHandler:
    Debug.Assert False

End Sub
Private Sub CopyGelNetAdjDef2003ToCurrent(OldDef As NetAdjDefinition2003, ByRef CurrentNetAdjDef As NetAdjDefinition)
    
    Dim intIndex As Integer
    
    SetDefaultUMCNETAdjDef CurrentNetAdjDef
    
    With CurrentNetAdjDef
        .MinUMCCount = OldDef.MinUMCCount
        .MinScanRange = OldDef.MinScanRange
        .MaxScanPct = OldDef.MaxScanPct
        .TopAbuPct = OldDef.TopAbuPct
        ' Ignored: .PeakSelection = OldDef.PeakSelection
        ' Ignored: .PeakMaxAbuPct = OldDef.PeakMaxAbuPct
        
        On Error Resume Next
        
        ' ReDim .PeakCSSelection(UBound(OldDef.PeakCSSelection))
        For intIndex = 0 To UBound(.PeakCSSelection)
            .PeakCSSelection(intIndex) = OldDef.PeakCSSelection(intIndex)
        Next intIndex
        
        ' Note: .MWField is no longer used
        .MWTolType = OldDef.MWTolType
        .MWTol = OldDef.MWTol
        .NETFormula = OldDef.NETFormula
        .NETTolIterative = OldDef.NETTol
        .NETorRT = OldDef.NETorRT
        .UseNET = OldDef.UseNET
        .UseMultiIDMaxNETDist = OldDef.UseMultiIDMaxNETDist
        .MultiIDMaxNETDist = OldDef.MultiIDMaxNETDist
        .EliminateBadNET = OldDef.EliminateBadNET
        .MaxIDToUse = OldDef.MaxIDToUse
        .IterationStopType = OldDef.IterationStopType
        .IterationStopValue = OldDef.IterationStopValue
        .IterationUseMWDec = OldDef.IterationUseMWDec
        .IterationMWDec = OldDef.IterationMWDec
        .IterationUseNETdec = OldDef.IterationUseNETdec
        .IterationNETDec = OldDef.IterationNETDec
        .IterationAcceptLast = OldDef.IterationAcceptLast
        
        ' Note: The remaining options are initialized by the call to SetDefaultUMCNETAdjDef above
                  
        .OtherInfo = ""
        
    End With
    
End Sub

Private Sub CopyGelNetAdjDef2004ToCurrent(OldDef As NetAdjDefinition2004, ByRef CurrentNetAdjDef As NetAdjDefinition)
  
    Dim intIndex As Integer
    
    SetDefaultUMCNETAdjDef CurrentNetAdjDef
    
    With CurrentNetAdjDef
        .MinUMCCount = OldDef.MinUMCCount
        .MinScanRange = OldDef.MinScanRange
        .MaxScanPct = OldDef.MaxScanPct
        .TopAbuPct = OldDef.TopAbuPct
        ' Ignored: .PeakSelection = OldDef.PeakSelection
        ' Ignored: .PeakMaxAbuPct = OldDef.PeakMaxAbuPct
        
        On Error Resume Next
        
        ' ReDim .PeakCSSelection(UBound(OldDef.PeakCSSelection))
        For intIndex = 0 To UBound(.PeakCSSelection)
            .PeakCSSelection(intIndex) = OldDef.PeakCSSelection(intIndex)
        Next intIndex
        
        ' Note: .MWField is no longer used
        .MWTolType = OldDef.MWTolType
        .MWTol = OldDef.MWTol
        .NETFormula = OldDef.NETFormula
        .NETTolIterative = OldDef.NETTol
        .NETorRT = OldDef.NETorRT
        .UseNET = OldDef.UseNET
        .UseMultiIDMaxNETDist = OldDef.UseMultiIDMaxNETDist
        .MultiIDMaxNETDist = OldDef.MultiIDMaxNETDist
        .EliminateBadNET = OldDef.EliminateBadNET
        .MaxIDToUse = OldDef.MaxIDToUse
        .IterationStopType = OldDef.IterationStopType
        .IterationStopValue = OldDef.IterationStopValue
        .IterationUseMWDec = OldDef.IterationUseMWDec
        .IterationMWDec = OldDef.IterationMWDec
        .IterationUseNETdec = OldDef.IterationUseNETdec
        .IterationNETDec = OldDef.IterationNETDec
        .IterationAcceptLast = OldDef.IterationAcceptLast

        .InitialSlope = OldDef.InitialSlope
        .InitialIntercept = OldDef.InitialIntercept
        
        ' Use of NET Adj Lockers for NET adjustment is no longer supported (March 2006)
''        .UseNetAdjLockers = OldDef.UseNetAdjLockers
''        .UseOldNetAdjIfFailure = OldDef.UseOldNetAdjIfFailure
''        .NetAdjLockerMinimumMatchCount = OldDef.NetAdjLockerMinimumMatchCount
        
        ' Note: The remaining options are initialized by the call to SetDefaultUMCNETAdjDef above

        .OtherInfo = OldDef.OtherInfo
  End With

End Sub

Private Sub CopyGelNetAdjDef2005aToCurrent(OldDef As NetAdjDefinition2005a, ByRef CurrentNetAdjDef As NetAdjDefinition)
  
    Dim intIndex As Integer
    
    SetDefaultUMCNETAdjDef CurrentNetAdjDef
    
    With CurrentNetAdjDef
        .MinUMCCount = OldDef.MinUMCCount
        .MinScanRange = OldDef.MinScanRange
        .MaxScanPct = OldDef.MaxScanPct
        .TopAbuPct = OldDef.TopAbuPct
        ' Ignored: .PeakSelection = OldDef.PeakSelection
        ' Ignored: .PeakMaxAbuPct = OldDef.PeakMaxAbuPct
        
        On Error Resume Next
        
        ' ReDim .PeakCSSelection(UBound(OldDef.PeakCSSelection))
        For intIndex = 0 To UBound(.PeakCSSelection)
            .PeakCSSelection(intIndex) = OldDef.PeakCSSelection(intIndex)
        Next intIndex
        
        ' Note: .MWField is no longer used
        .MWTolType = OldDef.MWTolType
        .MWTol = OldDef.MWTol
        .NETFormula = OldDef.NETFormula
        .NETTolIterative = OldDef.NETTol
        .NETorRT = OldDef.NETorRT
        .UseNET = OldDef.UseNET
        .UseMultiIDMaxNETDist = OldDef.UseMultiIDMaxNETDist
        .MultiIDMaxNETDist = OldDef.MultiIDMaxNETDist
        .EliminateBadNET = OldDef.EliminateBadNET
        .MaxIDToUse = OldDef.MaxIDToUse
        .IterationStopType = OldDef.IterationStopType
        .IterationStopValue = OldDef.IterationStopValue
        .IterationUseMWDec = OldDef.IterationUseMWDec
        .IterationMWDec = OldDef.IterationMWDec
        .IterationUseNETdec = OldDef.IterationUseNETdec
        .IterationNETDec = OldDef.IterationNETDec
        .IterationAcceptLast = OldDef.IterationAcceptLast

        .InitialSlope = OldDef.InitialSlope
        .InitialIntercept = OldDef.InitialIntercept
        
        ' Use of NET Adj Lockers for NET adjustment is no longer supported (March 2006)
''        .UseNetAdjLockers = OldDef.UseNetAdjLockers
''        .UseOldNetAdjIfFailure = OldDef.UseOldNetAdjIfFailure
''        .NetAdjLockerMinimumMatchCount = OldDef.NetAdjLockerMinimumMatchCount
        
        .UseRobustNETAdjustment = OldDef.UseRobustNETAdjustment
        .RobustNETAdjustmentMode = OldDef.RobustNETAdjustmentMode
        
        .RobustNETSlopeStart = OldDef.RobustNETSlopeStart
        .RobustNETSlopeEnd = OldDef.RobustNETSlopeEnd
        .RobustNETSlopeIncreaseMode = OldDef.RobustNETSlopeIncreaseMode
        .RobustNETSlopeIncrement = OldDef.RobustNETSlopeIncrement
        
        .RobustNETInterceptStart = OldDef.RobustNETInterceptStart
        .RobustNETInterceptEnd = OldDef.RobustNETInterceptEnd
        .RobustNETInterceptIncrement = OldDef.RobustNETInterceptIncrement
        
        .RobustNETMassShiftPPMStart = OldDef.RobustNETMassShiftPPMStart
        .RobustNETMassShiftPPMEnd = OldDef.RobustNETMassShiftPPMEnd
        .RobustNETMassShiftPPMIncrement = OldDef.RobustNETMassShiftPPMIncrement
        
        ' Note: The remaining options are initialized by the call to SetDefaultUMCNETAdjDef above

        .OtherInfo = OldDef.OtherInfo
  End With

End Sub

Private Sub CopyGelNetAdjDef2005bToCurrent(OldDef As NetAdjDefinition2005b, ByRef CurrentNetAdjDef As NetAdjDefinition)
  
    Dim intIndex As Integer
    
    SetDefaultUMCNETAdjDef CurrentNetAdjDef
    
    With CurrentNetAdjDef
        .MinUMCCount = OldDef.MinUMCCount
        .MinScanRange = OldDef.MinScanRange
        .MaxScanPct = OldDef.MaxScanPct
        .TopAbuPct = OldDef.TopAbuPct
        ' Ignored: .PeakSelection = OldDef.PeakSelection
        ' Ignored: .PeakMaxAbuPct = OldDef.PeakMaxAbuPct
        
        On Error Resume Next
        
        ' ReDim .PeakCSSelection(UBound(OldDef.PeakCSSelection))
        For intIndex = 0 To UBound(.PeakCSSelection)
            .PeakCSSelection(intIndex) = OldDef.PeakCSSelection(intIndex)
        Next intIndex
        
        ' Note: .MWField is no longer used
        .MWTolType = OldDef.MWTolType
        .MWTol = OldDef.MWTol
        .NETFormula = OldDef.NETFormula
        .NETTolIterative = OldDef.NETTolIterative
        .NETorRT = OldDef.NETorRT
        .UseNET = OldDef.UseNET
        .UseMultiIDMaxNETDist = OldDef.UseMultiIDMaxNETDist
        .MultiIDMaxNETDist = OldDef.MultiIDMaxNETDist
        .EliminateBadNET = OldDef.EliminateBadNET
        .MaxIDToUse = OldDef.MaxIDToUse
        
        .IterationStopType = OldDef.IterationStopType
        .IterationStopValue = OldDef.IterationStopValue
        .IterationUseMWDec = OldDef.IterationUseMWDec
        .IterationMWDec = OldDef.IterationMWDec
        .IterationUseNETdec = OldDef.IterationUseNETdec
        .IterationNETDec = OldDef.IterationNETDec
        .IterationAcceptLast = OldDef.IterationAcceptLast

        .InitialSlope = OldDef.InitialSlope
        .InitialIntercept = OldDef.InitialIntercept
        
        ' Use of NET Adj Lockers for NET adjustment is no longer supported (March 2006)
''        .UseNetAdjLockers = OldDef.UseNetAdjLockers
''        .UseOldNetAdjIfFailure = OldDef.UseOldNetAdjIfFailure
''        .NetAdjLockerMinimumMatchCount = OldDef.NetAdjLockerMinimumMatchCount
        
        .UseRobustNETAdjustment = OldDef.UseRobustNETAdjustment
        .RobustNETAdjustmentMode = OldDef.RobustNETAdjustmentMode
        
        .RobustNETSlopeStart = OldDef.RobustNETSlopeStart
        .RobustNETSlopeEnd = OldDef.RobustNETSlopeEnd
        .RobustNETSlopeIncreaseMode = OldDef.RobustNETSlopeIncreaseMode
        .RobustNETSlopeIncrement = OldDef.RobustNETSlopeIncrement
        
        .RobustNETInterceptStart = OldDef.RobustNETInterceptStart
        .RobustNETInterceptEnd = OldDef.RobustNETInterceptEnd
        .RobustNETInterceptIncrement = OldDef.RobustNETInterceptIncrement
        
        .RobustNETMassShiftPPMStart = OldDef.RobustNETMassShiftPPMStart
        .RobustNETMassShiftPPMEnd = OldDef.RobustNETMassShiftPPMEnd
        .RobustNETMassShiftPPMIncrement = OldDef.RobustNETMassShiftPPMIncrement
        
        ' Note: The remaining options are initialized by the call to SetDefaultUMCNETAdjDef above

        .OtherInfo = OldDef.OtherInfo
  End With

End Sub

Private Sub CopyGelSearchDef2002ToCurrent(OldDef As udtSearchDefinition2002GroupType, ByRef CurrentSearchDef As udtSearchDefinitionGroupType)
    
    With CurrentSearchDef
        CopyGelUMCDef2002ToCurrent OldDef.UMCDef, .UMCDef
        .UMCIonNetDef = UMCIonNetDef
        
        CopyAMTSearchDef2002ToCurrent OldDef.AMTSearchOnIons, .AMTSearchOnIons
        CopyAMTSearchDef2002ToCurrent OldDef.AMTSearchOnUMCs, .AMTSearchOnUMCs
        CopyAMTSearchDef2002ToCurrent OldDef.AMTSearchOnPairs, .AMTSearchOnPairs
        
        .AnalysisHistory() = OldDef.AnalysisHistory()
        .AnalysisHistoryCount = OldDef.AnalysisHistoryCount
        With .MassCalibrationInfo
            .AdjustmentHistoryCount = 0
            ReDim .AdjustmentHistory(0)
        End With
        ResetDBSearchMassMods .AMTSearchMassMods
    End With

End Sub

Private Sub CopyGelSearchDef2003ToCurrent(OldDef As udtSearchDefinition2003GroupType, ByRef CurrentSearchDef As udtSearchDefinitionGroupType)
    
    With CurrentSearchDef
        CopyGelUMCDef2003aToCurrent OldDef.UMCDef, .UMCDef
        .UMCIonNetDef = UMCIonNetDef
        
        CopyAMTSearchDef2002ToCurrent OldDef.AMTSearchOnIons, .AMTSearchOnIons
        CopyAMTSearchDef2002ToCurrent OldDef.AMTSearchOnUMCs, .AMTSearchOnUMCs
        CopyAMTSearchDef2002ToCurrent OldDef.AMTSearchOnPairs, .AMTSearchOnPairs
        
        .AnalysisHistory() = OldDef.AnalysisHistory()
        .AnalysisHistoryCount = OldDef.AnalysisHistoryCount
        With .MassCalibrationInfo
            .AdjustmentHistoryCount = 0
            ReDim .AdjustmentHistory(0)
        End With
        ResetDBSearchMassMods .AMTSearchMassMods
        .OtherInfo = ""
    End With

End Sub

Private Sub CopyGelSearchDef2003bToCurrent(OldDef As udtSearchDefinition2003bGroupType, ByRef CurrentSearchDef As udtSearchDefinitionGroupType)
    
    With CurrentSearchDef
        CopyGelUMCDef2003aToCurrent OldDef.UMCDef, .UMCDef
        .UMCIonNetDef = UMCIonNetDef
        
        CopyAMTSearchDef2002ToCurrent OldDef.AMTSearchOnIons, .AMTSearchOnIons
        CopyAMTSearchDef2002ToCurrent OldDef.AMTSearchOnUMCs, .AMTSearchOnUMCs
        CopyAMTSearchDef2002ToCurrent OldDef.AMTSearchOnPairs, .AMTSearchOnPairs
        
        .AnalysisHistory() = OldDef.AnalysisHistory()
        .AnalysisHistoryCount = OldDef.AnalysisHistoryCount
        With .MassCalibrationInfo
            .AdjustmentHistoryCount = 0
            ReDim .AdjustmentHistory(0)
        End With
        ResetDBSearchMassMods .AMTSearchMassMods
        .OtherInfo = ""
    End With

End Sub

Private Sub CopyGelSearchDef2003cToCurrent(OldDef As udtSearchDefinition2003cGroupType, ByRef CurrentSearchDef As udtSearchDefinitionGroupType)
    
    With CurrentSearchDef
        CopyGelUMCDef2003aToCurrent OldDef.UMCDef, .UMCDef
        .UMCIonNetDef = UMCIonNetDef
        
        CopyAMTSearchDef2002ToCurrent OldDef.AMTSearchOnIons, .AMTSearchOnIons
        CopyAMTSearchDef2002ToCurrent OldDef.AMTSearchOnUMCs, .AMTSearchOnUMCs
        CopyAMTSearchDef2002ToCurrent OldDef.AMTSearchOnPairs, .AMTSearchOnPairs
        
        .AnalysisHistory() = OldDef.AnalysisHistory()
        .AnalysisHistoryCount = OldDef.AnalysisHistoryCount
        .MassCalibrationInfo = OldDef.MassCalibrationInfo
        ResetDBSearchMassMods .AMTSearchMassMods
        .OtherInfo = OldDef.OtherInfo
    End With

End Sub

Private Sub CopyGelSearchDef2003dToCurrent(OldDef As udtSearchDefinition2003dGroupType, ByRef CurrentSearchDef As udtSearchDefinitionGroupType)
    
    With CurrentSearchDef
        CopyGelUMCDef2003aToCurrent OldDef.UMCDef, .UMCDef
        .UMCIonNetDef = UMCIonNetDef
        
        CopyAMTSearchDef2002ToCurrent OldDef.AMTSearchOnIons, .AMTSearchOnIons
        CopyAMTSearchDef2002ToCurrent OldDef.AMTSearchOnUMCs, .AMTSearchOnUMCs
        CopyAMTSearchDef2002ToCurrent OldDef.AMTSearchOnPairs, .AMTSearchOnPairs
        
        .AnalysisHistory() = OldDef.AnalysisHistory()
        .AnalysisHistoryCount = OldDef.AnalysisHistoryCount
        .MassCalibrationInfo = OldDef.MassCalibrationInfo
        With .AMTSearchMassMods
            If OldDef.AMTSearchMassMods.DynamicMods Then
                .ModMode = 1
            Else
                .ModMode = 0
            End If
            .N15InsteadOfN14 = OldDef.AMTSearchMassMods.N15InsteadOfN14
            .PEO = OldDef.AMTSearchMassMods.PEO
            .ICATd0 = OldDef.AMTSearchMassMods.ICATd0
            .ICATd8 = OldDef.AMTSearchMassMods.ICATd8
            .Alkylation = OldDef.AMTSearchMassMods.Alkylation
            .AlkylationMass = OldDef.AMTSearchMassMods.AlkylationMass
            .ResidueToModify = ""
            .ResidueMassModification = 0
            .OtherInfo = OldDef.AMTSearchMassMods.OtherInfo
        End With
        .OtherInfo = OldDef.OtherInfo
    End With

End Sub

Private Sub CopyGelSearchDef2003eToCurrent(OldDef As udtSearchDefinition2003eGroupType, ByRef CurrentSearchDef As udtSearchDefinitionGroupType)
    
    With CurrentSearchDef
        CopyGelUMCDef2003aToCurrent OldDef.UMCDef, .UMCDef
        .UMCIonNetDef = OldDef.UMCIonNetDef
        
        CopyAMTSearchDef2002ToCurrent OldDef.AMTSearchOnIons, .AMTSearchOnIons
        CopyAMTSearchDef2002ToCurrent OldDef.AMTSearchOnUMCs, .AMTSearchOnUMCs
        CopyAMTSearchDef2002ToCurrent OldDef.AMTSearchOnPairs, .AMTSearchOnPairs
        
        .AnalysisHistory() = OldDef.AnalysisHistory()
        .AnalysisHistoryCount = OldDef.AnalysisHistoryCount
        .MassCalibrationInfo = OldDef.MassCalibrationInfo
        .AMTSearchMassMods = OldDef.AMTSearchMassMods
        .OtherInfo = OldDef.OtherInfo
    End With

End Sub

Private Sub CopyGelSearchDef2004ToCurrent(OldDef As udtSearchDefinition2004GroupType, ByRef CurrentSearchDef As udtSearchDefinitionGroupType)
    
    With CurrentSearchDef
        .UMCDef = OldDef.UMCDef
        .UMCIonNetDef = OldDef.UMCIonNetDef
        
        CopyAMTSearchDef2002ToCurrent OldDef.AMTSearchOnIons, .AMTSearchOnIons
        CopyAMTSearchDef2002ToCurrent OldDef.AMTSearchOnUMCs, .AMTSearchOnUMCs
        CopyAMTSearchDef2002ToCurrent OldDef.AMTSearchOnPairs, .AMTSearchOnPairs
        
        .AnalysisHistory() = OldDef.AnalysisHistory()
        .AnalysisHistoryCount = OldDef.AnalysisHistoryCount
        .MassCalibrationInfo = OldDef.MassCalibrationInfo
        .AMTSearchMassMods = OldDef.AMTSearchMassMods
        .AdditionalValue1 = 0
        .AdditionalValue2 = 0
        .AdditionalValue3 = 0
        .AdditionalValue4 = 0
        .AdditionalValue5 = 0
        .AdditionalValue6 = 0
        .AdditionalValue7 = 0
        .AdditionalValue8 = 0
        .OtherInfo = OldDef.OtherInfo
    End With

End Sub

Private Sub CopyGelUMC2002ToCurrent(OldUMC As UMCListType2002, ByRef CurrentUMCList As UMCListType)
    Dim lngIndex As Long
    
On Error GoTo CopyUMCDataErrorHandler

    With CurrentUMCList
    
        CopyGelUMCDef2002ToCurrent OldUMC.def, .def
        
        .UMCCnt = OldUMC.UMCCnt
        If .UMCCnt < 0 Then .UMCCnt = 0
        
        Debug.Assert .UMCCnt - 1 = UBound(OldUMC.UMCs())
        
        If .UMCCnt > 0 Then
            ReDim .UMCs(0 To .UMCCnt - 1)
            For lngIndex = 0 To .UMCCnt - 1
                With .UMCs(lngIndex)
                    .ClassRepInd = OldUMC.UMCs(lngIndex).ClassRepInd
                    .ClassRepType = OldUMC.UMCs(lngIndex).ClassRepType
                    .ClassCount = OldUMC.UMCs(lngIndex).ClassCount
                    .ClassMInd() = OldUMC.UMCs(lngIndex).ClassMInd()
                    .ClassMType() = OldUMC.UMCs(lngIndex).ClassMType()
                    .ClassAbundance = OldUMC.UMCs(lngIndex).ClassAbundance
                    .ClassMW = OldUMC.UMCs(lngIndex).ClassMW
                    .ClassMWStD = OldUMC.UMCs(lngIndex).ClassMWStD
                    .ClassScore = 0
                    .ClassNET = 0
                    .ClassStatusBits = 0
                    
                    ' Note: We call CalculateClasses() after exiting this function to populate the following:
                    .MinScan = 0
                    .MaxScan = 0
                    .MinMW = 0
                    .MaxMW = 0
                
                    .ChargeStateStatsRepInd = 0
                    .ChargeStateCount = 0
                    ReDim .ChargeStateBasedStats(0)
                End With
            Next lngIndex
        Else
            ReDim .UMCs(0)
        End If
        
        .MassCorrectionValuesDefined = False
    End With

    Exit Sub

CopyUMCDataErrorHandler:
    LogErrors Err.Number, "CopyGelUMC2002ToCurrent"
    Debug.Print "Error occurred in CopyGelUMC2002ToCurrent: " & Err.Description
    Resume Next
    
End Sub

Private Sub CopyGelUMC2003aToCurrent(OldUMC As UMCListType2003a, ByRef CurrentUMCList As UMCListType)
    Dim lngIndex As Long
    
On Error GoTo CopyUMCDataErrorHandler

    With CurrentUMCList
    
        CopyGelUMCDef2003aToCurrent OldUMC.def, .def
        
        .UMCCnt = OldUMC.UMCCnt
        If .UMCCnt < 0 Then .UMCCnt = 0
        
        Debug.Assert .UMCCnt - 1 = UBound(OldUMC.UMCs())
        
        If .UMCCnt > 0 Then
            ReDim .UMCs(0 To .UMCCnt - 1)
            For lngIndex = 0 To .UMCCnt - 1
                With .UMCs(lngIndex)
                    .ClassRepInd = OldUMC.UMCs(lngIndex).ClassRepInd
                    .ClassRepType = OldUMC.UMCs(lngIndex).ClassRepType
                    .ClassCount = OldUMC.UMCs(lngIndex).ClassCount
                    .ClassMInd() = OldUMC.UMCs(lngIndex).ClassMInd()
                    .ClassMType() = OldUMC.UMCs(lngIndex).ClassMType()
                    .ClassAbundance = OldUMC.UMCs(lngIndex).ClassAbundance
                    .ClassMW = OldUMC.UMCs(lngIndex).ClassMW
                    .ClassMWStD = OldUMC.UMCs(lngIndex).ClassMWStD
                    .ClassScore = 0
                    .ClassNET = 0
                    .ClassStatusBits = OldUMC.UMCs(lngIndex).ClassStatusBits
                    
                    ' Note: The following will be updated by CalculateClasses() after exiting this function
                    .MinScan = OldUMC.UMCs(lngIndex).MinScan
                    .MaxScan = OldUMC.UMCs(lngIndex).MaxScan
                    .MinMW = OldUMC.UMCs(lngIndex).MinMW
                    .MaxMW = OldUMC.UMCs(lngIndex).MaxMW
                    
                    .ChargeStateStatsRepInd = 0
                    .ChargeStateCount = 0
                    ReDim .ChargeStateBasedStats(0)
                End With
            Next lngIndex
        Else
            ReDim .UMCs(0)
        End If
    
        .MassCorrectionValuesDefined = False
    End With

    Exit Sub

CopyUMCDataErrorHandler:
    LogErrors Err.Number, "CopyGelUMC2003aToCurrent"
    Debug.Print "Error occurred in CopyGelUMC2003aToCurrent: " & Err.Description
    Resume Next
    
End Sub

Private Sub CopyGelUMC2004ToCurrent(OldUMC As UMCListType2004, ByRef CurrentUMCList As UMCListType)
    Dim lngIndex As Long
    
On Error GoTo CopyUMCDataErrorHandler

    With CurrentUMCList
    
        .def = OldUMC.def
        
        .UMCCnt = OldUMC.UMCCnt
        If .UMCCnt < 0 Then .UMCCnt = 0
        
        Debug.Assert .UMCCnt - 1 = UBound(OldUMC.UMCs())
        
        If .UMCCnt > 0 Then
            ReDim .UMCs(0 To .UMCCnt - 1)
            For lngIndex = 0 To .UMCCnt - 1
                With .UMCs(lngIndex)
                    .ClassRepInd = OldUMC.UMCs(lngIndex).ClassRepInd
                    .ClassRepType = OldUMC.UMCs(lngIndex).ClassRepType
                    .ClassCount = OldUMC.UMCs(lngIndex).ClassCount
                    .ClassMInd() = OldUMC.UMCs(lngIndex).ClassMInd()
                    .ClassMType() = OldUMC.UMCs(lngIndex).ClassMType()
                    .ClassAbundance = OldUMC.UMCs(lngIndex).ClassAbundance
                    .ClassMW = OldUMC.UMCs(lngIndex).ClassMW
                    .ClassMWStD = OldUMC.UMCs(lngIndex).ClassMWStD
                    
                    .ClassMassCorrectionDa = 0
                    
                    .ClassScore = OldUMC.UMCs(lngIndex).ClassScore
                    .ClassNET = OldUMC.UMCs(lngIndex).ClassNET
                    .ClassStatusBits = OldUMC.UMCs(lngIndex).ClassStatusBits
                    
                    ' Note: The following will be updated by CalculateClasses() after exiting this function
                    .MinScan = OldUMC.UMCs(lngIndex).MinScan
                    .MaxScan = OldUMC.UMCs(lngIndex).MaxScan
                    .MinMW = OldUMC.UMCs(lngIndex).MinMW
                    .MaxMW = OldUMC.UMCs(lngIndex).MaxMW
                    
                    .ChargeStateStatsRepInd = OldUMC.UMCs(lngIndex).ChargeStateStatsRepInd
                    .ChargeStateCount = OldUMC.UMCs(lngIndex).ChargeStateCount
                    
                    If .ChargeStateCount > 0 Then
                        ReDim .ChargeStateBasedStats(.ChargeStateCount - 1)
                        ' Copy the entire array
                        .ChargeStateBasedStats = OldUMC.UMCs(lngIndex).ChargeStateBasedStats
                    Else
                        ReDim .ChargeStateBasedStats(0)
                    End If
                End With
            Next lngIndex
        Else
            ReDim .UMCs(0)
        End If
        
        .MassCorrectionValuesDefined = False
    End With

    InitializeAdditionalUMCDefVariables CurrentUMCList.def

    Exit Sub

CopyUMCDataErrorHandler:
    LogErrors Err.Number, "CopyGelUMC2002ToCurrent"
    Debug.Print "Error occurred in CopyGelUMC2002ToCurrent: " & Err.Description
    Resume Next
    
End Sub

Private Sub CopyGelUMCDef2002ToCurrent(OldDef As UMCDefinition2002, ByRef CurrentUMCDef As UMCDefinition)
    
    With CurrentUMCDef
        .UMCType = OldDef.UMCType
        .DefScope = OldDef.DefScope
        .MWField = OldDef.MWField
        .TolType = OldDef.TolType
        .Tol = OldDef.Tol
        .UMCSharing = False                 ' This is now false by default
        .UMCUniCS = OldDef.UMCUniCS
        .ClassAbu = OldDef.ClassAbu
        .ClassMW = OldDef.ClassMW
        .GapMaxCnt = OldDef.GapMaxCnt
        .GapMaxSize = OldDef.GapMaxSize
        .GapMaxPct = OldDef.GapMaxPct
        .UMCNETType = OldDef.UMCNETType
        .UMCMinCnt = OldDef.UMCMinCnt
        .UMCMaxCnt = OldDef.UMCMaxCnt
        .InterpolateGaps = True
        .InterpolateMaxGapSize = 3
        .InterpolationType = 0
        .ChargeStateStatsRepType = 0
        .UMCClassStatsUseStatsFromMostAbuChargeState = False
        .OtherInfo = ""
    End With

    InitializeAdditionalUMCDefVariables CurrentUMCDef
End Sub

Private Sub CopyGelUMCDef2003aToCurrent(OldDef As UMCDefinition2003a, ByRef CurrentUMCDef As UMCDefinition)
    
    With CurrentUMCDef
        .UMCType = OldDef.UMCType
        .DefScope = OldDef.DefScope
        .MWField = OldDef.MWField
        .TolType = OldDef.TolType
        .Tol = OldDef.Tol
        .UMCSharing = False
        .UMCUniCS = OldDef.UMCUniCS
        .ClassAbu = OldDef.ClassAbu
        .ClassMW = OldDef.ClassMW
        .GapMaxCnt = OldDef.GapMaxCnt
        .GapMaxSize = OldDef.GapMaxSize
        .GapMaxPct = OldDef.GapMaxPct
        .UMCNETType = OldDef.UMCNETType
        .UMCMinCnt = OldDef.UMCMinCnt
        .UMCMaxCnt = OldDef.UMCMaxCnt
        .InterpolateGaps = True
        .InterpolateMaxGapSize = 3
        .InterpolationType = 0
        .ChargeStateStatsRepType = 0
        .UMCClassStatsUseStatsFromMostAbuChargeState = False
        .OtherInfo = ""
    End With

    InitializeAdditionalUMCDefVariables CurrentUMCDef
End Sub

Private Sub CopyLegacyMassCalibrationInfoToData(ByRef CurrentGelData As DocumentData, ByRef MassCalibrationInfo As udtMassCalibrationInfoType)
    Dim i As Long
    Dim bytCount As Byte
    
    If MassCalibrationInfo.AdjustmentHistoryCount > 0 Then
        If MassCalibrationInfo.AdjustmentHistoryCount > 255 Then
            bytCount = 255
        Else
            bytCount = MassCalibrationInfo.AdjustmentHistoryCount
        End If
        
        With CurrentGelData
            Select Case MassCalibrationInfo.MassUnits
            Case gltABS
                For i = 1 To .CSLines
                    .CSData(i).MassShiftCount = bytCount
                    .CSData(i).MassShiftOverallPPM = MassToPPM(MassCalibrationInfo.OverallMassAdjustment, .CSData(i).AverageMW)
                Next i
                
                For i = 1 To .IsoLines
                    .IsoData(i).MassShiftCount = bytCount
                    .IsoData(i).MassShiftOverallPPM = MassToPPM(MassCalibrationInfo.OverallMassAdjustment, .IsoData(i).MonoisotopicMW)
                Next i
            Case gltPPM
                For i = 1 To .CSLines
                    .CSData(i).MassShiftCount = bytCount
                    .CSData(i).MassShiftOverallPPM = MassCalibrationInfo.OverallMassAdjustment
                Next i
                
                For i = 1 To .IsoLines
                    .IsoData(i).MassShiftCount = bytCount
                    .IsoData(i).MassShiftOverallPPM = MassCalibrationInfo.OverallMassAdjustment
                Next i
            Case Else
                ' This shouldn't happen
                Debug.Assert False
            End Select
            
        End With
    End If
End Sub

Private Sub GelAnalysisInfoRead(ByVal InFileNum As Integer, ByVal fIndex As Long)
    Get #InFileNum, , mGelAnalysis
    If mGelAnalysis.ValidAnalysisDataPresent Then
        
        If GelAnalysis(fIndex) Is Nothing Then Set GelAnalysis(fIndex) = New FTICRAnalysis
        
        FillGelAnalysisObject GelAnalysis(fIndex), mGelAnalysis
    
    End If
    
End Sub

Private Sub GelAnalysisInfoWrite(ByVal OutFileNum As Integer, ByVal fIndex As Long)
    
    FillGelAnalysisInfo mGelAnalysis, GelAnalysis(fIndex)
    
    Put #OutFileNum, , mGelAnalysis
End Sub

Private Sub InitializeAdditionalUMCDefVariables(ByRef CurrentUMCDef As UMCDefinition)
    With CurrentUMCDef
        .OddEvenProcessingMode = oepUMCOddEvenProcessingMode.oepProcessAll
        .RequireMatchingIsotopeTag = True
        .AdditionalValue4 = 0
        .AdditionalValue5 = 0
        .AdditionalValue6 = 0
        .AdditionalValue7 = 0
        .AdditionalValue8 = 0
    End With

End Sub

Private Sub ValidateDataArrays(ByVal lngGelIndex As Long)
    Dim lngIndex As Long
    
On Error GoTo FixCSArray

    lngIndex = UBound(GelData(lngGelIndex).CSData)

ResumeChecking:
On Error GoTo FixIsoArray
    lngIndex = UBound(GelData(lngGelIndex).IsoData)

Exit Sub

FixCSArray:
    ReDim GelData(lngGelIndex).CSData(0)
    Resume ResumeChecking

FixIsoArray:
    ReDim GelData(lngGelIndex).IsoData(0)
    
End Sub

Private Sub InitFileIOOffsetsAndVersions()
    
    FileInfoHeaderOffsets(fioGelData) = FileIO_Offset_FirstHeader
    FileInfoHeaderOffsets(fioGelUMC) = FileIO_Offset_FirstHeader + 50
    FileInfoHeaderOffsets(fioGelAnalysis) = FileIO_Offset_FirstHeader + 100
    FileInfoHeaderOffsets(fioUMCNetAdjDef) = FileIO_Offset_FirstHeader + 150
    FileInfoHeaderOffsets(fioSearchDefinitions) = FileIO_Offset_FirstHeader + 200
    FileInfoHeaderOffsets(fioORFData) = FileIO_Offset_FirstHeader + 250
    FileInfoHeaderOffsets(fioORFMassTags) = FileIO_Offset_FirstHeader + 300
    FileInfoHeaderOffsets(fioGelPairs) = FileIO_Offset_FirstHeader + 350
    FileInfoHeaderOffsets(fioGelDeltaLabeledPairs) = FileIO_Offset_FirstHeader + 400
    FileInfoHeaderOffsets(fioGelIDP) = FileIO_Offset_FirstHeader + 450
    FileInfoHeaderOffsets(fioGelLM) = FileIO_Offset_FirstHeader + 500
    FileInfoHeaderOffsets(fioGelORFViewerSavedGelListAndOptions) = FileIO_Offset_FirstHeader + 550
        
    FileInfoVersions(fioGelData) = 7#               ' Note: Version 1# = DocumentData2000, Version 2# = DocumentData2003 type, with Iso_Field_Count = 10, Version 3# = DocumentData2003b type, Version 4# = DocumentData2004 type (Iso_Field_Count = 12), Version 5# = DocumentData2005a (now using udtIsotopicDataType instead of Iso_Field); Version 6# = DocumentData2005b; Version 7# = current DocumentData type
    FileInfoVersions(fioGelUMC) = 5#                ' Note: Version 1# = UMCListType2002, Version 2# = UMCListType2003a, Version 3# = UMC2004, Version 4# is same structure size as 5#, but four old Double variables were turned into several new 16-bit and 32-bit integer variables, Version 5# = current UMC type
    FileInfoVersions(fioGelAnalysis) = 1#
    FileInfoVersions(fioUMCNetAdjDef) = 6#          ' Note: Version 1# = NetAdjDefinition2003, Version 2# was a short-lived beta version, Version 3# = NetAdjDefinition2004, Version 4# = NetAdjDefinition2005a,  Version 5# = NetAdjDefinition2005b, Version 6# = current NetAdjDef type
    FileInfoVersions(fioSearchDefinitions) = 9#     ' Note: Version 2# uses UMCDefinition2002, Version 3# uses UMCDefinition2003, Version 4#, 5#, 6#, and 7# use UMCDefinition2003a, 8# uses current UMCDefinition2005 type, Version 9# uses current UMCDefinition type
    FileInfoVersions(fioORFData) = 5#               ' Since GelORFData().ORFs() and GelORFMassTags().ORFs() are parallel arrays, be sure to change the fioORFData and fioORFMassTags version numbers together
    FileInfoVersions(fioORFMassTags) = 5#
    FileInfoVersions(fioGelPairs) = 2#
    FileInfoVersions(fioGelDeltaLabeledPairs) = 8#  ' Note: Version 2# uses IsoPairsDltLbl2003Type; Version 3# uses IsoPairsDltLbl2004aType; Version 4# uses IsoPairsDltLbl2004bType; Version 5# uses IsoPairsDltLbl2004cType; Version 6# uses IsoPairsDltLbl2004dType; Version 7# uses IsoPairsDltLbl2004eType; Version 8# uses the current definition
    FileInfoVersions(fioGelIDP) = 2#
    FileInfoVersions(fioGelLM) = 2#
    FileInfoVersions(fioGelORFViewerSavedGelListAndOptions) = 3#
    
End Sub


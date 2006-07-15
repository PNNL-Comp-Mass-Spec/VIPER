Class mzDataReader in MZDataFileReaderDLL.dll can be used to open a .mzData file 
and return each spectrum present.  To use, call function ReadMZDataFile() with the 
file path to parse.  There are two to ways obtain the data read:

Mode 1
 Read the entire mzData file and cache all data in memory.  Then, use function
 GetCachedScanDataByIndex() or GetCachedScanDataByScanNumber() to obtain the data
 for each scan.  Example code is included below.

Mode 2
 Declare the mzDataReader class as WithEvents and add a handler for the ScanRead event.
 As the mzDataReader class reads each scan, the ScanRead event will be raised and you
 can use function GetCurrentScanData() to retrieve the data for the current scan.
 Example code is included below.

Note that projects using this DLL need to have references to Microsoft XML 
(msxml2.dll or newer).  In addition, you need a reference to the Windows Script Host
Object Model (wshom.ocx) since the class uses the FileSystemObject class.  Use
Project->Referencs to define references in VB6.

-------------------------------------------------------------------------------
Written by Matthew Monroe for the Department of Energy (PNNL, Richland, WA)
Copyright 2006, Battelle Memorial Institute.  All Rights Reserved.

E-mail: matthew.monroe@pnl.gov or matt@alchemistmatt.com
Website: http://ncrr.pnl.gov/ or http://www.sysbio.org/resources/staff/
-------------------------------------------------------------------------------

Licensed under the Apache License, Version 2.0; you may not use this file except 
in compliance with the License.  You may obtain a copy of the License at 
http://www.apache.org/licenses/LICENSE-2.0

Notice: This computer software was prepared by Battelle Memorial Institute, 
hereinafter the Contractor, under Contract No. DE-AC05-76RL0 1830 with the 
Department of Energy (DOE).  All rights in the computer software are reserved 
by DOE on behalf of the United States Government and the Contractor as 
provided in the Contract.  NEITHER THE GOVERNMENT NOR THE CONTRACTOR MAKES ANY 
WARRANTY, EXPRESS OR IMPLIED, OR ASSUMES ANY LIABILITY FOR THE USE OF THIS 
SOFTWARE.  This notice including this sentence must appear on any copies of 
this computer software.


---------------------------
-- Example code for Mode 1 
---------------------------

Private Sub ReadDataCached()
    ' Cached mode example
    Dim objMZdataFileReader As New MZDataFileReaderDLL.mzDataReader
    
    Dim objScanInfo As MZDataFileReaderDLL.mzDataScanInfo
    Dim MZList() As Double
    Dim IntensityList() As Single
    
    Dim lngIndex As Long
    Dim blnSuccess As Boolean
    
    Dim strFilePath As String
    strFilePath = Command()
    
    If Len(strFilePath) = 0 Then
        strFilePath = "SmallTest.mzData"
    End If
     
    objMZdataFileReader.CacheDataInMemory = True
    objMZdataFileReader.ReadMZDataFile (strFilePath)
    
    If Len(objMZdataFileReader.ErrorMessage) > 0 Then
        MsgBox objMZdataFileReader.ErrorMessage
    End If
    
    For lngIndex = 0 To objMZdataFileReader.FileInfoScanCount - 1
        blnSuccess = objMZdataFileReader.GetCachedScanDataByIndex(lngIndex, MZList, IntensityList, objScanInfo)
        If blnSuccess Then
            Debug.Print objScanInfo.ScanNumber & ", " & objScanInfo.MSLevel; ", " & objScanInfo.PeaksCount
            If objScanInfo.PeaksCount > 0 Then
                Debug.Assert UBound(MZList) + 1 = objScanInfo.PeaksCount And UBound(IntensityList) + 1 = objScanInfo.PeaksCount
                If lngIndex Mod 5 = 0 Then
                    ' Print the first and last data point for every 5th scan
                    Debug.Print "  " & objScanInfo.ScanNumber & ",  " & MZList(0) & ", " & IntensityList(0)
                    Debug.Print "  " & objScanInfo.ScanNumber & ",  " & MZList(objScanInfo.PeaksCount - 1) & ", " & IntensityList(objScanInfo.PeaksCount - 1)
                End If
            End If
        End If
    Next lngIndex

    Set objMZdataFileReader = Nothing

End Sub

---------------------------
-- Example code for Mode 2 
---------------------------
	
Private WithEvents mMZDataFileReader As MZDataFileReaderDLL.mzDataReader

Public Sub ReadFile(strFilePath As String)
    ' Non-cached mode example
  
    Set mMZDataFileReader = New MZDataFileReaderDLL.mzDataReader
    
    mMZDataFileReader.CacheDataInMemory = False
    mMZDataFileReader.ReadMZDataFile (strFilePath)
    
    If Len(mMZDataFileReader.ErrorMessage) > 0 Then
        MsgBox mMZDataFileReader.ErrorMessage
    End If
    
    Set mMZDataFileReader = Nothing
End Sub

Private Sub mMZDataFileReader_ScanRead()
    Dim objScanInfo As MZDataFileReaderDLL.mzDataScanInfo
    Dim MZList() As Double
    Dim IntensityList() As Single
    
    Dim blnSuccess As Boolean
    
    blnSuccess = mMZDataFileReader.GetCurrentScanData(MZList, IntensityList, objScanInfo)
    
    If blnSuccess Then
        Debug.Print objScanInfo.ScanNumber & ", " & objScanInfo.MSLevel; ", " & objScanInfo.PeaksCount
        If objScanInfo.PeaksCount > 0 Then
            Debug.Assert UBound(MZList) + 1 = objScanInfo.PeaksCount And UBound(IntensityList) + 1 = objScanInfo.PeaksCount
            If (objScanInfo.ScanNumber - 1) Mod 5 = 0 Then
                ' Print the first and last data point for every 5th scan
                Debug.Print "  " & objScanInfo.ScanNumber & ",  " & MZList(0) & ", " & IntensityList(0)
                Debug.Print "  " & objScanInfo.ScanNumber & ",  " & MZList(objScanInfo.PeaksCount - 1) & ", " & IntensityList(objScanInfo.PeaksCount - 1)
            End If
        End If
    End If
End Sub

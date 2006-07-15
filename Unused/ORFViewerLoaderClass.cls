VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ORFViewerLoaderClass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Written by Matthew Monroe, PNNL
' Started December 31, 2002

' Provides routines to load and destroy ORFView forms
' The ORFView Forms (type frmORFViewer) are dynamically loaded as needed
' When a user closes an ORFView Form, it is hidden and the corresponding entry
'   in ORFViewFormIDs() is set to UNUSED_ORF_FORM
' When the user requests that a new ORFView be shown, the loaded forms are searched
'   through to see if any have a ORFViewFormIDs() value of UNUSED_ORF_FORM
' If one with UNUSED_ORF_FORM is found, then that form is used
' Otherwise, a new form is loaded
' Finally, if a form is already loaded with the ID of the calling Gel as the BaseID, then the ORFViewerLoader for that form is shown
'
Private Const UNUSED_ORF_FORM = -1

Private ORFViewForms() As frmORFViewer      ' 0-based array of forms

' Although an ORFViewForm can contain multiple gels (or the "OverLay" gel), the
'  BaseID contains the ID of the gel that was active when the ORFViewForm was loaded
Private ORFViewFormIDs() As Long            ' 0-based array; holds a unique FormID value for each form, or UNUSED_ORF_FORM if an unused, hidden form

Private ORFViewFormCount As Integer

Public InitializingORFViewerForm As Boolean


Public Sub HideORFViewerForm(frmCallingForm As VB.Form, lngFormID As Long)
    Dim lngIndex As Long
    
    frmCallingForm.Visible = False
    
    For lngIndex = 0 To ORFViewFormCount - 1
        If ORFViewFormIDs(lngIndex) = lngFormID Then
            ORFViewFormIDs(lngIndex) = UNUSED_ORF_FORM
            Exit For
        End If
    Next lngIndex
    
End Sub

Public Sub FindORFContainingMassTagInCallingForm(lngFormID As Long, lngMassTagID As Long)
    Dim intIndex As Integer
    
    For intIndex = 0 To ORFViewFormCount - 1
        If ORFViewFormIDs(intIndex) = lngFormID Then
            ORFViewForms(intIndex).FindORFContainingMassTag False, lngMassTagID
            Exit For
        End If
    Next intIndex
End Sub

Public Sub ShowORFViewerForm(strGelFilePathForCaption As String, Optional lngGelIDToInclude As Long = -1, Optional blnLoadNewWindow As Boolean = False)
    ' If blnLoadNewWindow = False, then displays ORFViewForms(0), unless it hasn't been loaded, in which case it is loaded and displayed
    ' If blnLoadNewWindow = True,
    ' Otherwise, loads a new form
    
    Dim lngIndex As Long
    Dim intNewFormIndex As Integer
    Dim blnAutoIncludeGel As Boolean
    Dim strFilepath As String
    
    intNewFormIndex = UNUSED_ORF_FORM
    blnAutoIncludeGel = True
    
    ' If blnLoadNewWindow = False, then see if any ORF Viewers are already visible
    If Not blnLoadNewWindow Then
        For lngIndex = 0 To ORFViewFormCount - 1
            If ORFViewFormIDs(lngIndex) <> UNUSED_ORF_FORM Then
                ' Form already visible
                intNewFormIndex = lngIndex
                blnAutoIncludeGel = False
                Exit For
            End If
        Next lngIndex
    Else
        If lngGelIDToInclude < 1 Then
            blnAutoIncludeGel = False
        End If
    End If
    
    If intNewFormIndex = UNUSED_ORF_FORM Then
        ' See if any of the forms have a GelID value of UNUSED_ORF_FORM
        For lngIndex = 0 To ORFViewFormCount - 1
            If ORFViewFormIDs(lngIndex) = UNUSED_ORF_FORM Then
                intNewFormIndex = lngIndex
                Exit For
            End If
        Next lngIndex
    End If
    
    If intNewFormIndex = UNUSED_ORF_FORM Then
        ' Need to load a new form
        ORFViewFormCount = ORFViewFormCount + 1
        ReDim Preserve ORFViewForms(ORFViewFormCount)
        ReDim Preserve ORFViewFormIDs(ORFViewFormCount)
        
        intNewFormIndex = ORFViewFormCount - 1
        Set ORFViewForms(intNewFormIndex) = New frmORFViewer
        
    End If
    ORFViewForms(intNewFormIndex).SetFormID CLng(intNewFormIndex + 1)
    ORFViewFormIDs(intNewFormIndex) = CLng(intNewFormIndex + 1)
    
    With ORFViewForms(intNewFormIndex)
        .Show vbModeless, MDIForm1
        .Visible = True
        If .WindowState = vbMinimized Then
            .WindowState = vbNormal
        End If
        
        If strGelFilePathForCaption = "" Or LCase(Left(strGelFilePathForCaption, 8)) = "untitled" Then
            If lngGelIDToInclude >= 1 Then
                strFilepath = GelData(lngGelIDToInclude).Filename
            Else
                strFilepath = ""
            End If
        Else
            strFilepath = strGelFilePathForCaption
        End If
        .Caption = "ORF Viewer"
        
    End With
    
    If blnAutoIncludeGel Then
        ShowORFViewerFormUsingSavedOptions intNewFormIndex, lngGelIDToInclude
    End If
    
End Sub

Private Sub ShowORFViewerFormUsingSavedOptions(intORFViewFormIndex As Integer, lngGelIDToInclude As Long)
    ' Check if Saved ORF Viewer options exist
    ' If they do, show the ORF Viewer using the saved options
    ' If not, show using default options
    
    Dim strFileNameToMatch As String
    Dim lngIsoLinesToMatch As Long
    Dim blnMatchFound As Boolean
    
    Dim udtSavedGelList As udtORFViewerGelListType
    Dim blnValidSavedData As Boolean
    Dim intSavedGelIndex As Integer, intCurrentGelIndex As Integer, intZOrderIndex As Integer
    Dim intIncludedGelCount As Integer
    Dim blnZOrderValueFound As Boolean
    Dim intIndexOfGelWithZOrderValue As Long

On Error GoTo ShowORFViewerFormUsingSavedOptionsErrorHandler

    If Not GelORFViewerSavedGelListAndOptions(lngGelIDToInclude).IsDefined Then
        ' Look through GelORFViewerSavedGelListAndOptions() to see if any of them contain the
        '  .FileName and equivalent .IsoLines for the Gel at index lngGelIDToInclude
        strFileNameToMatch = GelData(lngGelIDToInclude).Filename
        lngIsoLinesToMatch = GelData(lngGelIDToInclude).IsoLines
        
        For intCurrentGelIndex = 1 To UBound(GelBody())
            If GelORFViewerSavedGelListAndOptions(intCurrentGelIndex).IsDefined Then
                With GelORFViewerSavedGelListAndOptions(intCurrentGelIndex).SavedGelListAndOptions
                    For intSavedGelIndex = 1 To .GelCount
                        With .Gels(intSavedGelIndex)
                            If Not .Deleted And .Filename = strFileNameToMatch And .IsoLines = lngIsoLinesToMatch Then
                                ' Match Found
                                ' Change lngGelIDToInclude to be intCurrentGelIndex
                                lngGelIDToInclude = intCurrentGelIndex
                                blnMatchFound = True
                            End If
                        End With
                        If blnMatchFound Then Exit For
                    Next intSavedGelIndex
                End With
            End If
            If blnMatchFound Then Exit For
        Next intCurrentGelIndex
    End If
    
    If GelORFViewerSavedGelListAndOptions(lngGelIDToInclude).IsDefined Then
        InitializingORFViewerForm = True
        
        udtSavedGelList = GelORFViewerSavedGelListAndOptions(lngGelIDToInclude).SavedGelListAndOptions

        ' Initialize gOrfViewerOptionsCurrentGelList
        InitializeGelDisplayListAndOptions gOrfViewerOptionsCurrentGelList, lngGelIDToInclude

        ' Copy valid data from udtSavedGelList to gOrfViewerOptionsCurrentGelList
        ' Note that the call to .UpdateGelDisplayListAndOptions() below will read the data from gOrfViewerOptionsCurrentGelList
        '  and update the ORF viewer form accordingly
        For intSavedGelIndex = 1 To udtSavedGelList.GelCount
            If udtSavedGelList.Gels(intSavedGelIndex).IncludeGel And Not udtSavedGelList.Gels(intSavedGelIndex).Deleted Then
                blnMatchFound = False
                For intCurrentGelIndex = 1 To gOrfViewerOptionsCurrentGelList.GelCount
                    With gOrfViewerOptionsCurrentGelList.Gels(intCurrentGelIndex)
                        If Not .Deleted And _
                           udtSavedGelList.Gels(intSavedGelIndex).Filename = .Filename And _
                           udtSavedGelList.Gels(intSavedGelIndex).IsoLines = .IsoLines Then
                            ' Found a match between gOrfViewerOptionsCurrentGelList and udtSavedGelList
                            ' Make sure gel is included in ORF Viewer and has the correct options
                            
                            .IncludeGel = True
                            .IonSpotColor = udtSavedGelList.Gels(intSavedGelIndex).IonSpotColor
                            .IonSpotColorSelected = udtSavedGelList.Gels(intSavedGelIndex).IonSpotColorSelected
                            .IonSpotShape = udtSavedGelList.Gels(intSavedGelIndex).IonSpotShape
                            .NETAdjustmentType = udtSavedGelList.Gels(intSavedGelIndex).NETAdjustmentType
                            .UMCSpotColor = udtSavedGelList.Gels(intSavedGelIndex).UMCSpotColor
                            .UMCSpotColorSelected = udtSavedGelList.Gels(intSavedGelIndex).UMCSpotColorSelected
                            .UMCSpotShape = udtSavedGelList.Gels(intSavedGelIndex).UMCSpotShape
                            .VisibleScopeOnly = udtSavedGelList.Gels(intSavedGelIndex).VisibleScopeOnly
                            .ZOrder = udtSavedGelList.Gels(intSavedGelIndex).ZOrder
                            
                            blnMatchFound = True
                        End If
                    End With
                    If blnMatchFound Then Exit For
                Next intCurrentGelIndex
            End If
        Next intSavedGelIndex
                
        ' Make sure the .ZOrder values are valid
        ' ZOrder should range from 0 to intIncludedGelCount-1
        intIncludedGelCount = GetNumberOfIncludedGels(gOrfViewerOptionsCurrentGelList)
        
        For intZOrderIndex = 0 To intIncludedGelCount - 1
            blnZOrderValueFound = False
            For intCurrentGelIndex = 1 To gOrfViewerOptionsCurrentGelList.GelCount
                If gOrfViewerOptionsCurrentGelList.Gels(intCurrentGelIndex).IncludeGel Then
                    If gOrfViewerOptionsCurrentGelList.Gels(intCurrentGelIndex).ZOrder = intZOrderIndex Then
                        blnZOrderValueFound = True
                        intIndexOfGelWithZOrderValue = intCurrentGelIndex
                        Exit For
                    End If
                End If
            Next intCurrentGelIndex
            
            If Not blnZOrderValueFound Then
                ' Retrofit all of the z-order values
                intZOrderIndex = 0
                For intCurrentGelIndex = 1 To gOrfViewerOptionsCurrentGelList.GelCount
                        With gOrfViewerOptionsCurrentGelList.Gels(intCurrentGelIndex)
                            If .IncludeGel Then
                                .ZOrder = intZOrderIndex
                                intZOrderIndex = intZOrderIndex + 1
                            End If
                        End With
                Next intCurrentGelIndex
                Exit For
            Else
                ' Make sure none of the other included gels uses this Z-Order
                For intCurrentGelIndex = 1 To gOrfViewerOptionsCurrentGelList.GelCount
                    If intCurrentGelIndex <> intIndexOfGelWithZOrderValue And gOrfViewerOptionsCurrentGelList.Gels(intCurrentGelIndex).IncludeGel And gOrfViewerOptionsCurrentGelList.Gels(intCurrentGelIndex).ZOrder = intZOrderIndex Then
                        gOrfViewerOptionsCurrentGelList.Gels(intCurrentGelIndex).ZOrder = intZOrderIndex + 1
                    End If
                Next intCurrentGelIndex
            End If
        Next intZOrderIndex
        
        ' Copy the display options from udtSavedGelList to gOrfViewerOptionsCurrentGelList
        gOrfViewerOptionsCurrentGelList.DisplayOptions = udtSavedGelList.DisplayOptions
        
        ' Update the mass and NET display range controls to the saved values
        ORFViewForms(intORFViewFormIndex).txtMassRange = udtSavedGelList.DisplayOptions.MassDisplayRangePPM
        ORFViewForms(intORFViewFormIndex).txtNETRange = udtSavedGelList.DisplayOptions.NETDisplayRange
        
        blnValidSavedData = True
        
        InitializingORFViewerForm = False
    Else
        blnValidSavedData = False
    End If

    ' Provided valid data was copied, display it in the desired ORFViewer form
    ' Otherwise, simply call .UpdateGelDisplayList() with lngGelIDToInclude
    If blnValidSavedData Then
        ORFViewForms(intORFViewFormIndex).UpdateGelDisplayListAndOptions True
    Else
        ORFViewForms(intORFViewFormIndex).UpdateGelDisplayList lngGelIDToInclude
    End If
    
    Exit Sub
    
ShowORFViewerFormUsingSavedOptionsErrorHandler:
    Debug.Assert False
    ORFViewForms(intORFViewFormIndex).UpdateGelDisplayList lngGelIDToInclude
End Sub

Public Sub UpdateORFViewerForm(lngFormID As Long, blnPopulateORFGroupList As Boolean)
    ' This sub should only be called from frmORFViewOptions
    
    Dim intIndex As Integer
    
    For intIndex = 0 To ORFViewFormCount - 1
        If ORFViewFormIDs(intIndex) = lngFormID Then
            ORFViewForms(intIndex).UpdateGelDisplayListAndOptions blnPopulateORFGroupList
            Exit For
        End If
    Next intIndex
        
End Sub

Private Sub Class_Initialize()
    ORFViewFormCount = 0
End Sub

Private Sub Class_Terminate()
    Dim lngIndex As Long
    
    ' Unload each of the loaded forms
    For lngIndex = 0 To ORFViewFormCount - 1
        Unload ORFViewForms(lngIndex)
    Next lngIndex
    
End Sub

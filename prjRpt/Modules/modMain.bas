Attribute VB_Name = "modMain"
Attribute VB_Ext_KEY = "RVB_UniqueId" ,"3B3781D7008E"
Option Explicit
'**********************************************************************
'
'           Module Name: modAPIMain.bas
'
'           Purpose    : Declare various functions used regularly by the application
'
'           Author     : Joyprakash Saikia
'           Created On : 25/04/2002
'
'**********************************************************************
Public bDEBUG As Boolean

Public rptError As UDTErrSave 'global Variable for



Public Function ConvertToTwip(ByVal eScale As JxDBRptScale, ByVal sValue As Single) As Single
    '************************************************************
    'Description :
    '           This Program considers Twips as  Internal ScaleMode .
    '           If user Gives the Value in other Scale Modes then
    '           This Conversion function is used to Calculate the twip Equivalent  of Scale Modes
    '
    '************************************************************

Dim sNewValue As Single
' Convert value to Twips
Select Case eScale
  Case JxDBRptScale.JxDBRptScaleTwip
    sNewValue = sValue
  Case JxDBRptScale.JxDBRptScaleInch
    sNewValue = sValue * 1440
  Case JxDBRptScale.JxDBRptScaleCm
    sNewValue = sValue * 567
  Case JxDBRptScale.JxDBRptScaleMm
    sNewValue = sValue * 56.7
End Select
ConvertToTwip = sNewValue

End Function


Public Function ConvertFromTwip(ByVal eScale As JxDBRptScale, ByVal sValue As Single) As Single
    '************************************************************
    'Description :
    '           This Program considers Twips as  Internal ScaleMode .
    '           If user Gives the Value in other Scale Modes then
    '           This Conversion function is used to Calculate the twip Equivalent  of Scale Modes
    '
    '************************************************************
    Dim sNewValue As Single
            
            Select Case eScale
              Case JxDBRptScale.JxDBRptScaleTwip
                sNewValue = sValue
              Case JxDBRptScale.JxDBRptScaleInch
                sNewValue = sValue / 1440
              Case JxDBRptScale.JxDBRptScaleCm
                sNewValue = sValue / 567 'it is 100 times of Mm
              Case JxDBRptScale.JxDBRptScaleMm
                sNewValue = sValue / 56.7
            End Select
    ConvertFromTwip = sNewValue
End Function


Public Function GetTextHeight(oTarget As Object, str As String, mRatio As Single) As Single
    '************************************************************
    'Description :
    '   function computes the width and height of the specified string of text
    '
    '
    '************************************************************
    Dim sz As size
    Dim retval
    retval = GetTextExtentPoint(oTarget.hDC, str, Len(str), sz)
    GetTextHeight = oTarget.ScaleY(sz.y, vbPixels, vbTwips)
    If Not TypeOf oTarget Is Printer Then
        If mRatio < 1 Then
            GetTextHeight = GetTextHeight * (1 / mRatio)
        Else
            GetTextHeight = GetTextHeight * (mRatio)
        End If
    End If
End Function

Public Function ShowSave(hWndOwner As Long, sFilter As String, sTitle As String, Optional nFlags As Long = &H80000) As String
    '************************************************************
    'Description :
    '           This acts same as ShowSave of Common Dialog Control
    '           But it is faster than Control version
    '           'Coz you Need not have to put Control on the Form
    '           and Directly Calling to Function without OLE Automation , i.e., COM
    '************************************************************

    Dim OFName As OPENFILENAME
    OFName.lStructSize = Len(OFName)
    OFName.hWndOwner = hWndOwner
    OFName.hInstance = App.hInstance
    OFName.lpstrFilter = sFilter
    OFName.lpstrFile = String(254, vbNullChar)
    OFName.nMaxFile = 255
    OFName.lpstrFileTitle = String(254, vbNullChar)
    OFName.nMaxFileTitle = 255
    OFName.lpstrTitle = sTitle
    OFName.flags = nFlags

    If GetSaveFileName(OFName) Then
        ShowSave = GetString(OFName.lpstrFile)
    Else
        ShowSave = ""
    End If
End Function
    '************************************************************
    'Description :
    '           Following Functions are Dealing File I/O.
    '           The Names are  *mimick* c like functions
    '
    '************************************************************

Public Function ftell(hFile As Long) As Long
    ftell = SetFilePointer(hFile, ByVal CLng(0), 0&, FILE_CURRENT)
End Function

Public Function fseek(hFile As Long, ByVal offset As Long, ByVal origin As Long) As Long
    fseek = SetFilePointer(hFile, ByVal CLng(offset), 0&, origin)
End Function

Public Function fclose(hFile As Long) As Long
    On Error Resume Next
    fclose = CloseHandle(hFile)
End Function

Public Function fopen(ByVal FileName As String, ByVal mode As String) As Long
'************************************************************
'
'          OutPut:
'           Handle to file opened/created
'           or zero on error
'
'           Notes
'           This function will not raise any error and it's caller responsibility
'           to check for error by examining the return value
'************************************************************
On Error GoTo fopenErr

    Dim readAccess As Long
    Dim OpenMode As Long
    Select Case mode
        Case "r"
            readAccess = GENERIC_READ
            OpenMode = OPEN_EXISTING
        Case "w"
            readAccess = GENERIC_WRITE
            OpenMode = CREATE_ALWAYS
        Case "a"
            readAccess = GENERIC_WRITE
            OpenMode = OPEN_ALWAYS
        Case "r+"
            readAccess = GENERIC_READ Or GENERIC_WRITE
            OpenMode = OPEN_EXISTING
        Case "w+"
            readAccess = GENERIC_WRITE Or GENERIC_READ
            OpenMode = CREATE_ALWAYS
        Case "a+"
            readAccess = GENERIC_READ Or GENERIC_WRITE
            OpenMode = OPEN_ALWAYS
        Case Else
            fopen = 0
            Exit Function
    End Select
    
    fopen = CreateFile(ByVal FileName, readAccess, FILE_SHARE_READ Or FILE_SHARE_WRITE, ByVal CLng(0), _
                    OpenMode, FILE_ATTRIBUTE_ARCHIVE, 0&)
    
    
fopenEXIT:

Exit Function
fopenErr:
          GoTo fopenEXIT
End Function

Public Function ShowPrintDialog(ByVal hWndOwner As Long, DeviceName As String, NoOfCopy As Long, StartPage As Long, EndPage As Long) As Long
'************************************************************
'             Date: 23/05/2002
'          Description:
'                 This Rutine will display the print dialog box
'          Input:
'
'               hWndOwner - Handle of the owner
'               DeviceName - Printer devicename to be displayed as default in the dialog
'
'          Return Values:
'                The functions returns nonzero if the user clicks OK.
'                If the user canceled or close the Dialog, zero is returned
'
'          OutPut:
'               DeviceName - Printer devicename selected by the user
'               NoOfCopy - Number of print copy
'               StartPage - Indicate starting page number to be printed
'                           If user selects Selection from the dialog, -1 is returned
'                           and EndPage value is undefined.
'                           If the user selects All, the value returned is unchaged
'               EndPage - Indicates the last page number to be printed. If user selects
'                         selection from the dialog, the value is undefined
'                         If the user selects All, the value returned is unchaged
'          Globals Modfied
'
'          COPYRIGHT Info
'          Portions of this function is written by Paul Kuliniewicz
'          at http://www.vbapi.com
'          Please read CopyRight document for more information on COpyRight
'
'          Author portion
'          I just made some changes on variable name for consistency.
'************************************************************
On Error GoTo ShowPrintDialogErr

    Dim tPD As PRINTDLG_TYPE ' holds information to make the dialog box
    Dim tPrintmode As DEVMODE ' holds settings for the printer device
    Dim tPrintnames As DEVNAMES ' holds device, driver, and port names
    Dim pMode As Long, pNames As Long  ' pointers to the memory blocks for the two DEV* structures
    Dim lRet As Long  ' return value of function

    ' First, load default settings into printmode.  Note that we only fill relevant information.
    tPrintmode.dmDeviceName = Printer.DeviceName  ' name of the printer
    tPrintmode.dmSize = Len(tPrintmode)  ' size of the data structure
    tPrintmode.dmFields = DM_ORIENTATION  ' identify which other members have information
    tPrintmode.dmOrientation = DMORIENT_PORTRAIT  ' default to Portrait orientation

    ' Next, load strings for default printer into printnames.  Note the unusual way in which such
    ' information is stored.  This is explained on the DEVNAMES page.
    tPrintnames.wDriverOffset = 8  ' offset of driver name string
    tPrintnames.wDeviceOffset = tPrintnames.wDriverOffset + 1 + Len(Printer.DriverName)  ' offset of printer name string
    tPrintnames.wOutputOffset = tPrintnames.wDeviceOffset + 1 + Len(Printer.Port)  ' offset to output port string
    tPrintnames.wDefault = 0  ' maybe this isn't the default selected printer
    ' Load the three strings into the buffer, separated by null characters.
    tPrintnames.extra = Printer.DriverName & vbNullChar & Printer.DeviceName & vbNullChar & Printer.Port & vbNullChar

    ' Finally, load the initialization settings into pd, which is passed to the function.  We'll
    ' set the pointers to the structures after this.
    tPD.lStructSize = Len(tPD)  ' size of structure
    tPD.hWndOwner = hWndOwner
    ' Flags: All Pages default, disable Print to File option, return device context:
    tPD.flags = PD_ALLPAGES Or PD_DISABLEPRINTTOFILE Or PD_RETURNDC
    If StartPage Then
        tPD.nMinPage = StartPage  ' allow user to select first page of "document"
    End If
    If EndPage Then
        tPD.nMaxPage = EndPage  ' let's say there are 15 pages of the "document"
    End If
    ' Note how we can ignore those members which will be set or are not used here.

    ' Copy the data in printmode and printnames into the memory blocks we allocate.
    tPD.hDevMode = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(tPrintmode))  ' allocate memory block
    pMode = GlobalLock(tPD.hDevMode)  ' get a pointer to the block
    CopyMemory ByVal pMode, tPrintmode, Len(tPrintmode)  ' copy structure to memory block
    lRet = GlobalUnlock(tPD.hDevMode)  ' unlock the block
    ' Now do the same for printnames.
    tPD.hDevNames = GlobalAlloc(GMEM_MOVEABLE Or GMEM_ZEROINIT, Len(tPrintnames))  ' allocate memory block
    pNames = GlobalLock(tPD.hDevNames)  ' get a pointer to the block
    CopyMemory ByVal pNames, tPrintnames, Len(tPrintnames)  ' copy structure to memory block
    lRet = GlobalUnlock(tPD.hDevNames)  ' unlock the block

    ' Finally, open the dialog box!
    lRet = PRINTDLG(tPD)  ' looks so simple, doesn't it?

    ' If the user hit OK, display some information about the selection.
    If lRet Then
        ' First, we must copy the memory block data back into the structures.  This is almost identical
        ' to the code above where we did the reverse.  Comments here are omitted for brevity.
        pMode = GlobalLock(tPD.hDevMode)
        CopyMemory tPrintmode, ByVal pMode, Len(tPrintmode)
        lRet = GlobalUnlock(tPD.hDevMode)
        pNames = GlobalLock(tPD.hDevNames)
        CopyMemory tPrintnames, ByVal pNames, Len(tPrintnames)
        lRet = GlobalUnlock(tPD.hDevNames)

        ' Now, display the information we want.  We could instead use this info to print something.
        DeviceName = tPrintmode.dmDeviceName
        'Debug.Print "Printer Name: "; printmode.dmDeviceName
        NoOfCopy = tPD.nCopies
        'Debug.Print "Number of Copies:"; pd.nCopies
        Debug.Print "Orientation: ";
        If tPrintmode.dmOrientation = DMORIENT_PORTRAIT Then
            Debug.Print "Portrait"
        Else
            Debug.Print "Landscape"
        End If
        If (tPD.flags And PD_SELECTION) = PD_SELECTION Then  ' user chose "Selection"
            StartPage = -1
            'Debug.Print "Print the current selection."
        ElseIf (tPD.flags And PD_PAGENUMS) = PD_PAGENUMS Then  ' user chose a page range
            StartPage = tPD.nFromPage
            EndPage = tPD.nToPage
        End If
    
        ShowPrintDialog = 1 'set return value
    End If

    ' No matter what, we have to deallocate the memory blocks from before.
    lRet = GlobalFree(tPD.hDevMode)
    lRet = GlobalFree(tPD.hDevNames)
    

ShowPrintDialogEXIT:

Exit Function
ShowPrintDialogErr:
          GoTo ShowPrintDialogEXIT
End Function

Public Function GetString(sInput As String) As String
'************************************************************
'
'          Description:
'                 This Routine is Remove Trailing  Null char From a string.
'                   Genenally input String returned by an API Call.
'
'          Input:
'
'              sInput - string with/withour null character
'************************************************************
On Error GoTo GetStringErr

    Dim iZeroPos As Integer
    iZeroPos = InStr(1, sInput, vbNullChar)
    If iZeroPos > 0 Then
        GetString = Left$(sInput, iZeroPos - 1)
    Else
        GetString = sInput
    End If
GetStringEXIT:

Exit Function
GetStringErr:
          GoTo GetStringEXIT
End Function
Public Function SetErrSource(ByVal sModName As String, ByVal sSource As String) As String
    
    SetErrSource = " at " & sModName & "->" & sSource
End Function
Public Sub SaveError(ByVal ErrNo As Long, ByVal ErrSource As String, _
                     ByVal ErrDesc As String, Optional ByVal HelpContext As Long = 0)
'************************************************************
'
'          Description:
'                 To Trace error Handling & Save them to a File
'                 Not used on Current version
'
'************************************************************
    
    
    With rptError
        .ErrNumber = ErrNo
        .ErrSource = ErrSource
        .ErrDesc = ErrDesc
        .ErrHelpContext = HelpContext
    End With
    
End Sub
Public Sub GetSavedError(ErrNo As Long, ErrSource As String, _
                        ErrDesc As String, HelpFile As String, _
                        HelpContext As Long)
'************************************************************
'
'          Description:
'                 To Retrive error and Corresponding Help Context
'                 Not used on Current version
'
'************************************************************
                        
    With rptError
        ErrNo = .ErrNumber
        ErrSource = .ErrSource
        ErrDesc = .ErrDesc
        HelpFile = .ErrHelpFile
        HelpContext = .ErrHelpContext
         
    End With
    
End Sub


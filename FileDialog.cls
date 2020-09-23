VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cFileDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' VB5/6 compatible File Open/Save Dialog class

Private Declare Function GetOpenFileName Lib "comdlg32" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function GetSaveFileName Lib "comdlg32" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Private Declare Function CommDlgExtendedError Lib "comdlg32" () As Long

' If the user cancels or closes the dialog box, or an error occurs, the return value
' is zero. To get extended error information, call the CommDlgExtendedError function,
' which can return one of the following values:

' Common dialog error return codes
'Private Const CDERR_DIALOGFAILURE = &HFFFF '(-32768) The dialog box could not be created. The common dialog box function's call to the DialogBox function failed. For example, this error occurs if the common dialog box call specifies an invalid window handle.
Private Const CDERR_GENERALCODES = &H0&      ' The user canceled or closed the dialog box.
Private Const CDERR_STRUCTSIZE = &H1&        ' The lStructSize member of the initialization structure for the corresponding common dialog box is invalid.
Private Const CDERR_INITIALIZATION = &H2&    ' The common dialog box function failed during initialization. This error often occurs when sufficient memory is not available.
Private Const CDERR_NOTEMPLATE = &H3&        ' The ENABLETEMPLATE flag was set in the Flags member of the initialization structure for the corresponding common dialog box, but you failed to provide a corresponding template.
Private Const CDERR_NOHINSTANCE = &H4&       ' The ENABLETEMPLATE flag was set in the Flags member of the initialization structure for the corresponding common dialog box, but you failed to provide a corresponding instance handle.
Private Const CDERR_LOADSTRFAILURE = &H5&    ' The common dialog box function failed to load a specified string.
Private Const CDERR_FINDRESFAILURE = &H6&    ' The common dialog box function failed to find a specified resource.
Private Const CDERR_LOADRESFAILURE = &H7&    ' The common dialog box function failed to load a specified resource.
Private Const CDERR_LOCKRESFAILURE = &H8&    ' The common dialog box function failed to lock a specified resource.
Private Const CDERR_MEMALLOCFAILURE = &H9&   ' The common dialog box function was unable to allocate memory for internal structures.
Private Const CDERR_MEMLOCKFAILURE = &HA&    ' The common dialog box function was unable to lock the memory associated with a handle.
Private Const CDERR_NOHOOK = &HB&            ' The ENABLEHOOK flag was set in the Flags member of the initialization structure for the corresponding common dialog box, but you failed to provide a pointer to a corresponding hook procedure.
'Private Const CDERR_REGISTERMSGFAIL = &HC&  ' The RegisterWindowMessage function returned an error code when it was called by the common dialog box function.

'Private Const FNERR_FILENAMECODES = &H3000&
Private Const FNERR_SUBCLASSFAILURE = &H3001&  ' An attempt to subclass a list box failed because sufficient memory was not available.
Private Const FNERR_INVALIDFILENAME = &H3002&  ' A filename is invalid.
Private Const FNERR_BUFFERTOOSMALL = &H3003&   ' The buffer pointed to by the lpstrFile member of the OPENFILENAME structure is too small for the filename specified by the user. The first two bytes of the lpstrFile buffer contain an integer value specifying the size, in bytes (ANSI version) or 16-bit characters (Unicode version), required to receive the full name.

' The OPENFILENAME structure contains information that the GetOpenFileName and GetSaveFileName
' functions use to initialize an Open or Save As common dialog box. After the user closes the
' dialog box, the system returns information about the user's selection in this structure.

Private Type OPENFILENAME
    lStructSize As Long         ' Specifies the length, in bytes, of the structure
    hWndOwner As Long           ' Identifies the window that owns the dialog box. This member can be any valid window handle, or it can be NULL if the dialog box has no owner
    hInstance As Long           ' If the OFN_ENABLETEMPLATEHANDLE flag is set in the Flags member, hInstance is the handle of a memory object containing a dialog box template. If the OFN_ENABLETEMPLATE flag is set, hInstance identifies a module that contains a dialog box template named by the lpTemplateName member. If neither flag is set, this member is ignored. If the OFN_EXPLORER flag is set, the system uses the specified template to create a dialog box that is a child of the default Explorer-style dialog box. If the OFN_EXPLORER flag is not set, the system uses the template to create an old-style dialog box that replaces the default dialog box.
    lpstrFilter As String       ' Pointer to a buffer containing pairs of null-terminated filter strings. The last string in the buffer must be terminated by two NULL characters. The first string in each pair is a display string that describes the filter (for example, "Text Files"), and the second string specifies the filter pattern (for example, "*.TXT"). To specify multiple filter patterns for a single display string, use a semicolon to separate the patterns (for example, "*.TXT;*.DOC;*.BAK"). A pattern string can be a combination of valid filename characters and the asterisk (*) wildcard character. Do not include spaces in the pattern string. The system does not change the order of the filters. It displays them in the File Types combo box in the order specified in lpstrFilter. If lpstrFilter is NULL, the dialog box does not display any filters.
    lpstrCustomFilter As String ' Pointer to a static buffer that contains a pair of null-terminated filter strings for preserving the filter pattern chosen by the user. The first string is your display string that describes the custom filter, and the second string is the filter pattern selected by the user. The first time your application creates the dialog box, you specify the first string, which can be any nonempty string. When the user selects a file, the dialog box copies the current filter pattern to the second string. The preserved filter pattern can be one of the patterns specified in the lpstrFilter buffer, or it can be a filter pattern typed by the user. The system uses the strings to initialize the user-defined file filter the next time the dialog box is created. If the nFilterIndex member is zero, the dialog box uses the custom filter. If this member is NULL, the dialog box does not preserve user-defined filter patterns.
                                ' If this member is not NULL, the value of the nMaxCustFilter member must specify the size, in bytes (ANSI version) or 16-bit characters (Unicode version), of the lpstrCustomFilter buffer.
    nMaxCustFilter As Long      ' Specifies the size, in bytes or characters, of the buffer identified by lpstrCustomFilter. This buffer should be at least 40 characters long. This member is ignored if lpstrCustomFilter is NULL or points to a NULL string.
    nFilterIndex As Long        ' Specifies the index of the currently selected filter in the File Types control. The buffer pointed to by lpstrFilter contains pairs of strings that define the filters. The first pair of strings has an index value of 1, the second pair 2, and so on. An index of zero indicates the custom filter specified by lpstrCustomFilter. You can specify an index on input to indicate the initial filter description and filter pattern for the dialog box. When the user selects a file, nFilterIndex returns the index of the currently displayed filter. If nFilterIndex is zero and lpstrCustomFilter is NULL, the system uses the first filter in the lpstrFilter buffer. If all three members are zero or NULL, the system does not use any filters and does not show any files in the file list control of the dialog box.
    lpstrFile As String         ' Pointer to a buffer that contains a filename used to initialize the File Name edit control. The first character of this buffer must be NULL if initialization is not necessary. When the GetOpenFileName or GetSaveFileName function returns successfully, this buffer contains the drive designator, path, filename, and extension of the selected file. If the OFN_ALLOWMULTISELECT flag is set and the user selects multiple files, the buffer contains the current directory followed by the filenames of the selected files. For Explorer-style dialog boxes, the directory and filename strings are NULL separated, with an extra NULL character after the last filename. For old-style dialog boxes, the strings are space separated and the function uses short filenames for filenames with spaces. You can use the FindFirstFile function to convert between long and short filenames.
                                ' If the buffer is too small, the function returns FALSE and the CommDlgExtendedError function returns FNERR_BUFFERTOOSMALL. In this case, the first two bytes of the lpstrFile buffer contain the required size, in bytes or characters.
    nMaxFile As Long            ' Specifies the size, in bytes (ANSI version) or 16-bit characters (Unicode version), of the buffer pointed to by lpstrFile. The GetOpenFileName and GetSaveFileName functions return FALSE if the buffer is too small to contain the file information. The buffer should be at least 256 characters long.
    lpstrFileTitle As String    ' Pointer to a buffer that receives the filename and extension (without path information) of the selected file. This member can be NULL.
    nMaxFileTitle As Long       ' Specifies the size, in bytes (ANSI version) or 16-bit characters (Unicode version), of the buffer pointed to by lpstrFileTitle. This member is ignored if lpstrFileTitle is NULL.
    lpstrInitialDir As String   ' Pointer to a string that specifies the initial file directory. NT5+/98+: If lpstrInitalDir is NULL and the current directory contains any files of the specified filter types, the initial directory is the current directory. If lpstrInitalDir is NULL and the current directory does not contain any files of the specified filter types, the initial directory is the personal files directory of the current user. To specify the current user's personal files directory as the initial directory, set lpstrInitialDir to the path returned by calling theSHGetSpecialFolderLocation function with the CSIDL_PERSONAL flag.
    lpstrTitle As String        ' Pointer to a string to be placed in the title bar of the dialog box. If this member is NULL, the system uses the default title (that is, Save As or Open).
    Flags As Long               ' A set of bit flags you can use to initialize the dialog box. When the dialog box returns, it sets these flags to indicate the user's input. This member can be a combination of the following flags: Flag meaning described below
    nFileOffset As Integer      ' Specifies the zero-based offset, in bytes (ANSI version) or 16-bit characters (Unicode version), from the beginning of the path to the filename in the string pointed to by lpstrFile. For example, if lpstrFile points to the following string, "c:\dir1\dir2\file.ext", this member contains the value 13 to indicate the offset of the "file.ext" string. If the user selects more than one file, nFileOffset is the offset to the first filename.
    nFileExtension As Integer   ' Specifies the zero-based offset, in bytes (ANSI version) or 16-bit characters (Unicode version), from the beginning of the path to the filename extension in the string pointed to by lpstrFile. For example, if lpstrFile points to the following string, "c:\dir1\dir2\file.ext", this member contains the value 18. If the user did not type an extension and lpstrDefExt is NULL, this member specifies an offset to the terminating null character. If the user typed "." as the last character in the filename, this member specifies zero.
    lpstrDefExt As String       ' Points to a buffer that contains the default extension. GetOpenFileName and GetSaveFileName append this extension to the filename if the user fails to type an extension. This string can be any length, but only the first three characters are appended. The string should not contain a period (.). If this member is NULL and the user fails to type an extension, no extension is appended.
    lCustData As Long           ' Specifies application-defined data that the system passes to the hook procedure identified by the lpfnHook member. When the system sends the WM_INITDIALOG message to the hook procedure, the message's lParam parameter is a pointer to the OPENFILENAME structure specified when the dialog box was created. The hook procedure can use this pointer to get the lCustData value.
    lpfnHook As Long            ' Pointer to a hook procedure. This member is ignored unless the Flags member includes the OFN_ENABLEHOOK flag. If the OFN_EXPLORER flag is not set in the Flags member, lpfnHook is a pointer to an OFNHookProcOldStyle hook procedure that receives messages intended for the dialog box. The hook procedure returns FALSE to pass a message to the default dialog box procedure or TRUE to discard the message. If OFN_EXPLORER is set, lpfnHook is a pointer to an OFNHookProc hook procedure. The hook procedure receives notification messages sent from the dialog box. The hook procedure also receives messages for any additional controls that you defined by specifying a child dialog template. The hook procedure does not receive messages intended for the standard controls of the default dialog box.
    lpTemplateName As String    ' Pointer to a null-terminated string that names a dialog template resource in the module identified by the hInstance member. For numbered dialog box resources, this can be a value returned by the MAKEINTRESOURCE macro. This member is ignored unless the OFN_ENABLETEMPLATE flag is set in the Flags member. If the OFN_EXPLORER flag is set, the system uses the specified template to create a dialog box that is a child of the default Explorer-style dialog box. If the OFN_EXPLORER flag is not set, the system uses the template to create an old-style dialog box that replaces the default dialog box.
End Type

' File Open/Save Dialog Flags
Public Enum OFN_FLAGS
    OFN_READONLY = &H1&                ' Causes the Read Only check box to be checked initially when the dialog box is created. This flag indicates the state of the Read Only check box when the dialog box is closed.
    OFN_OVERWRITEPROMPT = &H2&         ' Causes the Save As dialog box to generate a message box if the selected file already exists. The user must confirm whether to overwrite the file.
    OFN_HIDEREADONLY = &H4&            ' Hides the Read Only check box.
    OFN_NOCHANGEDIR = &H8&             ' Restores the current directory to its original value if the user changed the directory while searching for files.
    OFN_SHOWHELP = &H10&               ' Causes the dialog box to display the Help button. The hwndOwner member must specify the window to receive the HELPMSGSTRING registered messages that the dialog box sends when the user clicks the Help button. An Explorer-style dialog box sends a CDN_HELP notification message to your hook procedure when the user clicks the Help button.
    OFN_ENABLEHOOK = &H20&             ' Enables the hook function specified in the lpfnHook member.
    OFN_ENABLETEMPLATE = &H40&         ' Indicates that the lpTemplateName member points to the name of a dialog template resource in the module identified by the hInstance member. If the OFN_EXPLORER flag is set, the system uses the specified template to create a dialog box that is a child of the default Explorer-style dialog box. If the OFN_EXPLORER flag is not set, the system uses the template to create an old-style dialog box that replaces the default dialog box.
    OFN_ENABLETEMPLATEHANDLE = &H80&   ' Indicates that the hInstance member identifies a data block that contains a preloaded dialog box template. The system ignores the lpTemplateName if this flag is specified. If the OFN_EXPLORER flag is set, the system uses the specified template to create a dialog box that is a child of the default Explorer-style dialog box. If the OFN_EXPLORER flag is not set, the system uses the template to create an old-style dialog box that replaces the default dialog box.
    OFN_NOVALIDATE = &H100&            ' Specifies that the common dialog boxes allow invalid characters in the returned filename. Typically, the calling application uses a hook procedure that checks the filename by using the FILEOKSTRING message. If the text box in the edit control is empty or contains nothing but spaces, the lists of files and directories are updated. If the text box in the edit control contains anything else, nFileOffset and nFileExtension are set to values generated by parsing the text. No default extension is added to the text, nor is text copied to the buffer specified by lpstrFileTitle. If the value specified by nFileOffset is less than zero, the filename is invalid. Otherwise, the filename is valid, and nFileExtension and nFileOffset can be used as if the OFN_NOVALIDATE flag had not been specified.
    OFN_ALLOWMULTISELECT = &H200&      ' Specifies that the File Name list box allows multiple selections. If you also set the OFN_EXPLORER flag, the dialog box uses the Explorer-style user interface; otherwise, it uses the old-style user interface. If the user selects more than one file, the lpstrFile buffer returns the path to the current directory followed by the filenames of the selected files. The nFileOffset member is the offset, in bytes or characters, to the first filename, and the nFileExtension member is not used. For Explorer-style dialog boxes, the directory and filename strings are NULL separated, with an extra NULL character after the last filename. This format enables the Explorer-style dialogs to return long filenames that include spaces. For old-style dialog boxes, the directory and filename strings are separated by spaces and the function uses short filenames for filenames with spaces. You can use theFindFirstFile function to convert between long and short filenames.
                                       ' If you specify a custom template for an old-style dialog box, the definition of the File Name list box must contain the LBS_EXTENDEDSEL value.
    OFN_EXTENTIONDIFFERENT = &H400&    ' Specifies that the user typed a filename extension that differs from the extension specified by lpstrDefExt. The function does not use this flag if lpstrDefExt is NULL.
    OFN_PATHMUSTEXIST = &H800&         ' Specifies that the user can type only valid paths and filenames. If this flag is used and the user types an invalid path and filename in the File Name entry field, the dialog box function displays a warning in a message box.
    OFN_FILEMUSTEXIST = &H1000&        ' Specifies that the user can type only names of existing files in the File Name entry field. If this flag is specified and the user enters an invalid name, the dialog box procedure displays a warning in a message box. If this flag is specified, the OFN_PATHMUSTEXIST flag is also used.
    OFN_CREATEPROMPT = &H2000&         ' If the user specifies a file that does not exist, this flag causes the dialog box to prompt the user for permission to create the file. If the user chooses to create the file, the dialog box closes and the function returns the specified name; otherwise, the dialog box remains open. If you use this flag with the OFN_ALLOWMULTISELECT flag, the dialog box allows the user to specify only one nonexistent file.
    OFN_SHAREAWARE = &H4000&           ' Specifies that if a call to the OpenFile function fails because of a network sharing violation, the error is ignored and the dialog box returns the selected filename. If this flag is not set, the dialog box notifies your hook procedure when a network sharing violation occurs for the filename specified by the user. If you set the OFN_EXPLORER flag, the dialog box sends the CDN_SHAREVIOLATION message to the hook procedure. If you do not set OFN_EXPLORER, the dialog box sends the SHAREVISTRING registered message to the hook procedure.
    OFN_NOREADONLYRETURN = &H8000&     ' Specifies that the returned file does not have the Read Only check box checked and is not in a write-protected directory.
    OFN_NOTESTFILECREATE = &H10000     ' Specifies that the file is not created before the dialog box is closed. This flag should be specified if the application saves the file on a create-nonmodify network share. When an application specifies this flag, the library does not check for write protection, a full disk, an open drive door, or network protection. Applications using this flag must perform file operations carefully, because a file cannot be reopened once it is closed.
    OFN_NONETWORKBUTTON = &H20000      ' Hides and disables the Network button.
    OFN_NOLONGNAMES = &H40000          ' Force no long names for 4.x modules. For old-style dialog boxes, this flag causes the dialog box to use short filenames (8.3 format). Explorer-style dialog boxes ignore this flag and always display long filenames.
    OFN_EXPLORER = &H80000             ' New look commdlg. Indicates that any customizations made to the Open or Save As dialog box use the new Explorer-style customization methods. For more information, see Explorer-Style Hook Procedures and Explorer-Style Custom Templates. By default, the Open and Save As dialog boxes use the Explorer-style user interface regardless of whether this flag is set. This flag is necessary only if you provide a hook procedure or custom template, or set the OFN_ALLOWMULTISELECT flag. If you want the old-style user interface, omit the OFN_EXPLORER flag and provide a replacement old-style template or hook procedure. If you want the old style but do not need a custom template or hook procedure, simply provide a hook procedure that always returns FALSE.
    OFN_NODEREFERENCELINKS = &H100000  ' Directs the dialog box to return the path and filename of the selected shortcut (.LNK) file. If this value is not given, the dialog box returns the path and filename of the file referenced by the shortcut.
    OFN_LONGNAMES = &H200000           ' Force long names for 3.x modules. For old-style dialog boxes, this flag causes the dialog box to use long filenames. If this flag is not specified, or if the OFN_ALLOWMULTISELECT flag is also set, old-style dialog boxes use short filenames (8.3 format) for filenames with spaces. Explorer-style dialog boxes ignore this flag and always display long filenames.
    OFN_ENABLESIZING = &H800000        ' Windows NT 5.0, Windows 98: Enables the Explorer-style dialog box to be resized using either the mouse or the keyboard. By default, the Explorer-style Open and Save As dialog boxes allow the dialog box to be resized regardless of whether this flag is set. This flag is necessary only if you provide a hook procedure or custom template. The old-style dialog box does not permit resizing.
End Enum

'Const OFN_EX_NOPLACESBAR =      &H1&
'Const OFN_ENABLEINCLUDENOTIFY = &H400000
'Const OFN_ENABLESIZING =        &H800000
'Const OFN_USEMONIKERS =         &H1000000
'Const OFN_DONTADDTORECENT =     &H2000000
'Const OFN_FORCESHOWHIDDEN =     &H10000000

Private tOFN As OPENFILENAME
Private Const MAX_PATH = 260&
Private Const TEN_KB = &H2800&
Private Const ZZ As String = vbNullChar & vbNullChar

'    Set FileDialog = New cFileDialog
'
'    With FileDialog
'        .DialogTitle = "Open file(s)..."
'        .InitialDir = curpath
'        .FileName = filename
'        .DefaultExt = "txt"
'        .Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
'        .FilterIndex = 1&
'        ' File must exist, allow multi-select, use Win95 style dialog, hide read-only checkbox
'        .Flags = OFN_FILEMUSTEXIST Or OFN_ALLOWMULTISELECT Or OFN_EXPLORER Or OFN_HIDEREADONLY
'        .ShowOpen
'        filename = .FileName
'        curpath = CurDir$
'    End With
'    If (filename = vbNullString) Then
'        ' User didn't return a file name
'        Exit Sub
'    ElseIf (InStr(filename, Chr$(0)) = 0) Then
'        ' User chose a single file name
'        OpenFile (filename)
'    Else
'        ' User chose multiple file names, Chr(0) is the
'        ' null character used to seperate file names
'        Dim sPath As String, Length As Long
'        Dim Idx As Long, aStr() As String
'
'        Length = InStr(filename, Chr$(0))
'        sPath = AddBackslash(Left$(filename, Length - 1))
'
'        aStr = SplitStr(Mid$(filename, Length + 1), Chr$(0))
'        For Idx = 0 To UBound(aStr)
'            OpenFile (sPath & aStr(Idx))
'        Next Idx
'    End If
'
'    With FileDialog
'        .DialogTitle = "Save file as..."
'        .InitialDir = curpath
'        .FileName = filename
'        .DefaultExt = "txt"
'        ' Overwrite? prompt, requires valid path, no read-only checkbox
'        .Flags = OFN_OVERWRITEPROMPT Or OFN_PATHMUSTEXIST Or OFN_HIDEREADONLY
'        .ShowSave
'        filename = .FileName
'        curpath = CurDir$
'    End With
'
'    Set FileDialog = Nothing

Public Sub ShowOpen()
    Dim rc As Long
    tOFN.lStructSize = LenB(tOFN)
    tOFN.nMaxFile = LenB(tOFN.lpstrFile) + TEN_KB
    tOFN.lpstrFile = tOFN.lpstrFile & String$(TEN_KB, vbNullChar)
    rc = GetOpenFileName(tOFN)
    Do While rc = 0&
        rc = CommDlgExtendedError
        If rc = CDERR_GENERALCODES Then
            Exit Do ' User canceled
        ElseIf rc = FNERR_BUFFERTOOSMALL Then
            tOFN.nMaxFile = tOFN.nMaxFile + TEN_KB
            tOFN.lpstrFile = tOFN.lpstrFile & String$(TEN_KB, vbNullChar)
            rc = GetOpenFileName(tOFN)
'        ElseIf rc = FNERR_INVALIDFILENAME Then
'            tOFN.Flags = tOFN.Flags And (Not OFN_NOVALIDATE)
'            tOFN.nMaxFile = TEN_KB
'            tOFN.lpstrFile = String$(TEN_KB, vbNullChar)
'            rc = GetOpenFileName(tOFN)
        Else
            Beep
            Exit Do ' Internal error
        End If
    Loop
End Sub

Public Sub ShowSave()
    Dim rc As Long
    tOFN.lStructSize = LenB(tOFN)
    tOFN.nMaxFile = LenB(tOFN.lpstrFile) + MAX_PATH
    tOFN.lpstrFile = tOFN.lpstrFile & String$(MAX_PATH, vbNullChar)
    rc = GetSaveFileName(tOFN)
End Sub

Public Property Let FileName(sFileName As String)
    tOFN.lpstrFile = sFileName & ZZ
End Property

Public Property Get FileName() As String
    FileName = TrimZZ(tOFN.lpstrFile)
End Property

Public Property Let DialogTitle(sDialogTitle As String)
    tOFN.lpstrTitle = sDialogTitle & ZZ
End Property

Public Property Get DialogTitle() As String
    DialogTitle = TrimZZ(tOFN.lpstrTitle)
End Property

Public Property Let InitialDir(sInitDir As String)
    tOFN.lpstrInitialDir = sInitDir & ZZ
End Property

Public Property Get InitialDir() As String
    InitialDir = TrimZZ(tOFN.lpstrInitialDir)
End Property

Public Property Let DefaultExt(sDefaultExt As String)
    Dim i As Long
    i = InStr(sDefaultExt, ".")
    tOFN.lpstrDefExt = Mid$(sDefaultExt, i + 1&) & ZZ
End Property

Public Property Get DefaultExt() As String
    DefaultExt = TrimZZ(tOFN.lpstrDefExt)
End Property

Public Property Let Filter(sFilter As String)
    ' To specify multiple filter patterns for a single display string,
    ' use a semicolon to separate the patterns. No spaces in the pattern!
    ' Filter = "Text Files (*.txt *.doc)|*.txt;*.doc|All Files (*.*)|*.*"
    Dim i As Long                       '...........'               '...'
    Dim s As String
    s = sFilter
    i = InStr(s, "|")
    Do While i
        Mid$(s, i) = vbNullChar
        i = InStr(i + 1&, s, "|")
    Loop
    tOFN.lpstrFilter = s & ZZ
End Property

Public Property Get Filter() As String
    Dim i As Long
    Dim s As String
    s = TrimZZ(tOFN.lpstrFilter)
    i = InStr(s, vbNullChar)
    Do While i
        Mid$(s, i) = "|"
        i = InStr(i + 1&, s, vbNullChar)
    Loop
    Filter = s
End Property

Public Property Let Flags(ByVal eFlags As OFN_FLAGS)
    tOFN.Flags = eFlags
End Property

Public Property Get Flags() As OFN_FLAGS
    Flags = tOFN.Flags
End Property

Public Property Let FilterIndex(nIndex As Long)
    tOFN.nFilterIndex = nIndex
End Property

Public Property Get FilterIndex() As Long
    FilterIndex = tOFN.nFilterIndex
End Property

Public Function TrimZZ(StrZZ As String) As String
    ' StrZZ = "strZstrZstrZstrZZ[ZZZZZZ]" >> TrimZZ = "strZstrZstrZstr"
    ' StrZZ = "strZ[ZZZZZZ]"              >> TrimZZ = "str"
    ' StrZZ = "str  "                     >> TrimZZ = "str"
    Dim Idx As Long
    Idx = InStr(StrZZ, ZZ)
    If (Idx) Then
        TrimZZ = LeftB$(StrZZ, Idx + Idx - 2&)
    Else
        Idx = InStr(StrZZ, vbNullChar)
        If (Idx) Then
            TrimZZ = LeftB$(StrZZ, Idx + Idx - 2&) 'Rd
        Else
            TrimZZ = Trim$(StrZZ)
        End If
    End If
End Function

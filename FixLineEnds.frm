VERSION 5.00
Begin VB.Form frmFLE 
   Caption         =   " Fix Line Ends Program"
   ClientHeight    =   1860
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7005
   LinkTopic       =   "Form1"
   ScaleHeight     =   1860
   ScaleWidth      =   7005
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "Quit"
      Height          =   285
      Left            =   5820
      TabIndex        =   7
      Top             =   1080
      Width           =   975
   End
   Begin VB.OptionButton optFix 
      Caption         =   "Convert  Windows  CRLF  to  UNIX  LF"
      Height          =   255
      Index           =   1
      Left            =   210
      TabIndex        =   6
      Top             =   1380
      Width           =   3315
   End
   Begin VB.OptionButton optFix 
      Caption         =   "Convert  UNIX  LF  to  Windows  CRLF"
      Height          =   255
      Index           =   0
      Left            =   210
      TabIndex        =   5
      Top             =   1140
      Value           =   -1  'True
      Width           =   3315
   End
   Begin VB.CommandButton cmdFix 
      Caption         =   "Fix"
      Height          =   285
      Left            =   4740
      TabIndex        =   3
      Top             =   1080
      Width           =   975
   End
   Begin VB.CommandButton cmdFileOpen 
      Caption         =   "Open"
      Height          =   285
      Left            =   3660
      TabIndex        =   2
      Top             =   1080
      Width           =   975
   End
   Begin VB.TextBox txtDest 
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   570
      Width           =   6765
   End
   Begin VB.TextBox txtSrc 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   6765
   End
   Begin VB.Label lblResult 
      Alignment       =   2  'Center
      Height          =   255
      Left            =   3420
      TabIndex        =   4
      Top             =   1470
      Width           =   3495
   End
End
Attribute VB_Name = "frmFLE"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub CopyMemByR Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSrc As Any, ByVal lLenB As Long)
Private Declare Sub CopyMemByV Lib "kernel32" Alias "RtlMoveMemory" (ByVal lpDest As Long, ByVal lpSrc As Long, ByVal lLenB As Long)
Private Declare Function AllocStrBPtr Lib "oleaut32" Alias "SysAllocStringByteLen" (ByVal lAddrPtr As Long, ByVal lLenB As Long) As Long

Private Const CONTENT_CHUNK_BYTES = &H10000  '65536
Private Const CONTENT_CHUNK = &H8000&        '32768

Private FileDialog As cFileDialog
Private sPath As String

'     ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯
'
Private Const INVALID_HANDLE_VALUE = &HFFFFFFFF
Private Const MAX_PATH = 260&
Private Const ALT_NAME = 14&

Private Type WIN32_FIND_DATA
    dwFileAttributes As Long
    ftCreationTime As Currency
    ftLastAccessTime As Currency
    ftLastWriteTime As Currency
    nFileSizeHigh As Long
    nFileSizeLow As Long
    dwReserved0 As Long
    dwReserved1 As Long
    cFileName As String * MAX_PATH
    cAlternateFileName As String * ALT_NAME ' 8.3 format
End Type

Private Declare Function FindFirstFile Lib "kernel32" Alias "FindFirstFileA" _
    (ByVal lpFileName As String, lpFindFileData As WIN32_FIND_DATA) As Long

Private Declare Function FindClose Lib "kernel32" (ByVal hFindFile As Long) As Long
'
'     ¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯¯

Private Function SearchFileContent(sMatchText As String, sWithText As String) As Long
  'On Error GoTo ExitOut

    Dim sFilePath As String
    Dim sNewFilePath As String
    Dim sContent As String
    Dim sRestText As String
    Dim sResult As String
    Dim iChunkCnt As Long
    Dim iLenFile As Long
    Dim iLenTerm As Long
    Dim iLenRest As Long
    Dim iLenCont As Long
    Dim iCutOff As Long
    Dim iLenResult As Long
    Dim iStartB As Long
    Dim iDiffB As Long
    Dim iMidStart As Long
    Dim iMidLen As Long
    Dim iSeekRead As Long
    Dim lHitCnt As Long
    Dim lHits As Long
    Dim iCnt As Long
    Dim iFileRead As Integer
    Dim iFileWrite As Integer
    Dim fCRLFflag As Long

    sFilePath = txtSrc           ' Set source file path
    sNewFilePath = txtDest       ' Set destination file path

    If StrComp(sFilePath, sNewFilePath, vbTextCompare) = 0& Then
        MsgBox "Cannot write result to same file!  ", vbExclamation
        Exit Function
    End If
    Screen.MousePointer = vbArrowHourglass
                                 ' chr(10)                      chr(13) & chr(10)
    fCRLFflag = MidI(sMatchText, 1&) = 10 And MidI(sWithText, 1&) = 13 And MidI(sWithText, 2&) = 10
    'fCRLFflag == Exclusion flag (Don't replace Lf with CrLf if it is already preceded by a Cr)

    iFileRead = FreeFile
    Open sFilePath For Binary Access Read Lock Write As #iFileRead ' Open file to content search

    iFileWrite = FreeFile
    Open sNewFilePath For Output As #iFileWrite    ' Clear results file
    Close #iFileWrite

    iFileWrite = FreeFile
    Open sNewFilePath For Binary Access Write Lock Write As #iFileWrite ' Open file to write result

        iLenTerm = LenB(sMatchText)     ' Byte len of search term
        iLenFile = LOF(iFileRead) * 2&  ' Byte len of file to search

        If iLenFile < iLenTerm Then GoTo ExitOut     ' If nothing to do then exit out

        If iLenFile < iLenTerm + 2& Then
'            sContent = Space$(iLenFile * 0.5)
'            Get #iFileRead, 1&, sContent  ' Get file content small chunk
'            sResult = VBA.Replace$(sContent, sMatchText, sWithText)
'            Put #iFileWrite, , sResult    ' Write result to file
'        MsgBox "File is too small  ", vbExclamation
            GoTo ExitOut
        End If

        iChunkCnt = iLenFile \ CONTENT_CHUNK_BYTES   ' Calculate number of chunks
        iLenRest = iLenFile Mod CONTENT_CHUNK_BYTES  ' Calculate size of final chunk

        If iLenRest < iLenTerm + 2& Then   ' If last chunk too small just add it to prev chunk
            iLenRest = CONTENT_CHUNK_BYTES + iLenRest
            iChunkCnt = iChunkCnt - 1&
        End If

        iLenCont = CONTENT_CHUNK_BYTES + iLenTerm  ' Overlap by Len(term2match)

        ' Pre-allocate string to hold each chunk
        If iChunkCnt Then CopyMemByV VarPtr(sContent), VarPtr(AllocStrBPtr(0&, iLenCont)), 4&
        CopyMemByV VarPtr(sRestText), VarPtr(AllocStrBPtr(0&, iLenRest)), 4& ' Allocate buffer for remainder

        iSeekRead = 1& ' Ready
        iMidStart = 1& ' Set
        iStartB = 1&   ' Go...

        Do Until iCnt = iChunkCnt                  ' Search content chunk by chunk
            Get #iFileRead, iSeekRead, sContent    ' Get file content next chunk
            
            ' Replace search term with new term, assign result, return new byte length
            iLenResult = ReplaceB(sContent, sMatchText, sWithText, sResult, iStartB, lHitCnt, fCRLFflag)

            iCutOff = iStartB + iLenTerm - 1&      ' Byte pos at end of last found term (-1 StartB is first byte of term)
            iDiffB = iCutOff - CONTENT_CHUNK_BYTES ' Does cutoff pos exceed chunk byte size

            If iDiffB > 0& Then                    ' If so write up to that byte pos
                iMidLen = iLenResult - (iLenCont - iCutOff) - iMidStart + 1&  ' +1 Inclusive
            Else ' Or default to 2 bytes past (first 2 bytes on next chunk neccessary for Cr exclusion test,
                iMidLen = iLenResult - (iLenTerm - 2&) - iMidStart + 1&       ' +1 Inclusive
            End If ' if search term existed starting at pos-2 would have been found at end of last chunk)

            Put #iFileWrite, , MidB$(sResult, iMidStart, iMidLen) ' Write result to file

            If iDiffB > 0& Then
                iMidStart = 1& + iDiffB            ' Set new start after last found bytes past chunk size
                iStartB = iMidStart
            Else
                iMidStart = 3&                     ' Or default to after 2 bytes past start of new chunk
                iStartB = 3&
            End If

            iCnt = iCnt + 1&                       ' Increment chunk
            lHits = lHits + lHitCnt                ' Count replacements
            lHitCnt = 0&                           ' Reset count for next chunk

            iSeekRead = iSeekRead + CONTENT_CHUNK  ' Seek ahead one chunk
        Loop

        If iCnt = iChunkCnt Then                          ' Now do final chunk (or only chunk if small file)
            Get #iFileRead, iSeekRead, sRestText          ' Search the remaining file content
            iLenResult = ReplaceB(sRestText, sMatchText, sWithText, sResult, iStartB, lHitCnt, fCRLFflag)
            Put #iFileWrite, , MidB$(sResult, iMidStart)  ' Write result to file
            lHits = lHits + lHitCnt                       ' Record total hits
        End If
ExitOut:
        Screen.MousePointer = vbDefault
    Close #iFileRead           ' Close file
    Close #iFileWrite          ' Close file
    sContent = vbNullString    ' Redundant
    sRestText = vbNullString   ' Redundant
    sResult = vbNullString     ' Redundant
    SearchFileContent = lHits  ' Return total replacements

End Function

Public Function ReplaceB(sSrc As String, _
                         sTerm As String, _
                         sNewTerm As String, _
                         sResult As String, _
                         Optional lStartB As Long = 1, _
                         Optional lHitCnt As Long, _
                         Optional ByVal fExcCrLf As Long _
                         ) As Long ' ©Rd

    Dim lLenOld As Long, lLenNew As Long
    Dim lBefore As Long, lHitPos As Long
    Dim lSize As Long, lLenSrc As Long
    Dim lOffset As Long, lHit As Long
    Dim lOffStart As Long, lPos As Long
    Dim lStartPos As Long, lCnt As Long
    Dim lProg As Long, alHits() As Long
    Dim lpSrc As Long, lpTerm As Long
    Dim lpRet As Long, fSkip As Boolean

    'On Error GoTo FreakOut

    sResult = vbNullString
    If (lStartB < 1&) Then lStartPos = 1& Else lStartPos = lStartB ' Validate start pos
    lStartB = 0&

    lLenSrc = LenB(sSrc)
    If (lLenSrc = 0&) Then Exit Function  ' No text
    lpSrc = StrPtr(sSrc)
    
    lLenOld = LenB(sTerm)
    If (lLenOld = 0&) Then GoTo ShortCirc ' Nothing to find
    lLenNew = LenB(sNewTerm)

    lOffset = lLenNew - lLenOld
    lSize = 500& ' lSize = Arr chunk size
    ReDim alHits(0 To lSize) As Long

    lHit = InStrB(lStartPos, sSrc, sTerm) ' Do first search

    Do Until (lHit = 0&)            ' Do until no more hits
        If (lHit And 1&) Then
            If fExcCrLf Then
                lBefore = lHit - 2&

                ' Don't replace Lf with CrLf if it is already preceded by a Cr
                If (lBefore > 0&) Then fSkip = MidIB(sSrc, lBefore) = 13
            End If

            If fSkip Then
                fSkip = False
                lOffStart = 2&      ' Default offset start pos
            Else
                lOffStart = lLenOld ' Offset next start pos

                alHits(lCnt) = lHit
                lCnt = lCnt + 1&    ' Record hits

                If (lCnt = lHitCnt) Then Exit Do
                If (lCnt = lSize) Then
                    lSize = lSize + 5000&
                    ReDim Preserve alHits(0 To lSize) As Long
                End If
            End If
        Else
            lOffStart = 1&          ' Byte offset start pos
        End If
        lHit = InStrB(lHit + lOffStart, sSrc, sTerm)
    Loop

    lHitCnt = lCnt
    If (lCnt = 0&) Then GoTo ShortCirc ' No hits
    lStartB = alHits(lCnt - 1)         ' Last hit pos in _src_ string

    lSize = lLenSrc + (lOffset * lCnt) ' lSize = result byte count
    If (lSize = 0&) Then Exit Function ' Result is an empty string

    CopyMemByR ByVal VarPtr(sResult), AllocStrBPtr(0&, lSize), 4& ' Pre-allocate memory

    lpTerm = StrPtr(sNewTerm)
    lpRet = StrPtr(sResult)

    lOffStart = 0&
    If (lLenNew) Then
       For lHit = 0& To lCnt - 1&
           lHitPos = alHits(lHit) - 1& ' Zero base CopyMemory string handling
           lProg = lHitPos - lPos
           If (lProg) Then             ' Build new string
               CopyMemByV lpRet + lOffStart, lpSrc + lPos, lProg
               lOffStart = lOffStart + lProg
           End If
           CopyMemByV lpRet + lOffStart, lpTerm, lLenNew
           lOffStart = lOffStart + lLenNew
           lPos = lHitPos + lLenOld    ' No offset orig str
       Next
    Else
       For lHit = 0& To lCnt - 1&
           lHitPos = alHits(lHit) - 1&
           lProg = lHitPos - lPos      ' Build new string
           If (lProg) Then
               CopyMemByV lpRet + lOffStart, lpSrc + lPos, lProg
               lOffStart = lOffStart + lProg
           End If
           lPos = lHitPos + lLenOld    ' No offset orig str
       Next
    End If

    If (lOffStart < lSize) Then
        CopyMemByV lpRet + lOffStart, lpSrc + lPos, lSize - lOffStart
    End If

    ReplaceB = lSize  ' Return result byte length
FreakOut:
    Exit Function

ShortCirc: ' If nothing to do
    ReplaceB = lLenSrc
    CopyMemByR ByVal VarPtr(sResult), AllocStrBPtr(lpSrc, lLenSrc), 4&

End Function

Public Property Get MidI(sStr As String, ByVal lPos As Long) As Integer
    CopyMemByR MidI, ByVal StrPtr(sStr) + lPos + lPos - 2&, 2&
End Property

Public Property Get MidIB(sStr As String, ByVal lPosB As Long) As Integer
    CopyMemByR MidIB, ByVal StrPtr(sStr) + lPosB - 1&, 2&
End Property

Private Sub cmdFileOpen_Click()
    Dim s As String
    Dim i As Long
    Set FileDialog = New cFileDialog
    With FileDialog
        .DialogTitle = "Please select a file to fix"
        .InitialDir = sPath
        .FileName = vbNullString
        .Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
        .FilterIndex = 1&
        ' File must exist, use Win95 style dialog, hide read-only checkbox
        .Flags = OFN_FILEMUSTEXIST Or OFN_EXPLORER Or OFN_HIDEREADONLY
        .ShowOpen
        s = .FileName
        sPath = CurDir$
    End With
    If (s <> vbNullString) Then
        txtSrc = s
        i = InStrR(s, ".")
        txtDest = Left$(s, i - 1&) & "_Fix" & Mid$(s, i)
    End If
End Sub

Private Function InStrR(sSrc As String, sTerm As String, Optional ByVal lRightStart As Long = -1, Optional ByVal eCompare As VbCompareMethod = vbBinaryCompare, Optional ByVal lLeftLimit As Long = 1) As Long
    Dim lPos As Long
    If lRightStart = -1& Then lRightStart = Len(sSrc)
    lPos = InStr(lLeftLimit, sSrc, sTerm, eCompare)
    Do Until lPos = 0& Or lPos > lRightStart
        InStrR = lPos
        lPos = InStr(InStrR + 1&, sSrc, sTerm, eCompare)
    Loop
End Function

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Function PathExists(sPath As String) As Boolean
    Dim hFindFiles As Long
    Dim pData As WIN32_FIND_DATA
    hFindFiles = FindFirstFile(sPath, pData)
    PathExists = Not (hFindFiles = INVALID_HANDLE_VALUE)
    FindClose hFindFiles
End Function

Private Sub cmdFix_Click()
    Dim s As String
    Dim i As Long
    If PathExists(txtSrc) Then
        s = Left$(txtDest, InStrR(txtDest, "\") - 1&)
        If PathExists(s) Then
            If optFix(0) Then
                i = SearchFileContent(ChrW$(10), ChrW$(13) & ChrW$(10))
            Else
                i = SearchFileContent(ChrW$(13) & ChrW$(10), ChrW$(10))
            End If
            lblResult = "Replaced " & i & " times"
        End If
    End If
End Sub

Private Sub Form_Load()
    sPath = App.Path
End Sub

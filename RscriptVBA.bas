Attribute VB_Name = "RScriptVBA"
Option Explicit
'===============================================================================================================================================================================================================================================================
'#  RscriptVBA ¬ github.com/ulchc (10-29-22)
'===============================================================================================================================================================================================================================================================
'===============================================================================================================================================================================================================================================================
'## Overview
'===============================================================================================================================================================================================================================================================
'
'<a href='https://github.com/ulchc/RscriptVBA'><img src='figures/Example.gif' align="right" height="650" /></a>
'
'Locate RScript 's path automatically, manage packages, run R code contained
'in a cell value, and read console output right into Excel's immediate window,
'with no additional set up beyond having R installed.
'
'===============================================================================================================================================================================================================================================================
'##  Usage Summary
'===============================================================================================================================================================================================================================================================
'
' WriteRunRscript() <br/>
' > To run R commands provided as a string
'
' CaptureRscriptOutput() <br/>
' > To run R commands and return the resulting R console output
'
' Attach_Libraries() <br/>
' > To append the local user's lib.loc to library() commands
'
' Require_Packages() <br/>
' > To install R packages from VBA
'
' WinShell_Rscript() <br/>
' > To call Rscript by manually specifying it's path and a saved script
'
'<br/><br/><br/><br/>
'
'===============================================================================================================================================================================================================================================================
'## Public Variables
'===============================================================================================================================================================================================================================================================
'----------------------------------------------------------------``` VBA
Public GlobalLoadLibraries As String
'
''   Stores the sequence of R commands to be appended to the start of
''   a script ran with WriteRunRscript() or CaptureRscriptOutput() when
''   using Attach_Libraries().
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
Public Enum RsVisibility
    RsHidden = 0
    RsVisible = 1
    RsMinimized = 2
End Enum
'
''   Options to toggle the visibilty of the R scripting window.
'
'----------------------------------------------------------------```
'===============================================================================================================================================================================================================================================================
'##  Main Functions & Subs
'===============================================================================================================================================================================================================================================================
'
'Note: Although libraries or package installation commands can be directly
'written into a script, it is preferable to use UDF's Attach_Libraries() or
'Require_Packages() in VABA so that packages which force the restart of an R
'session can be installed in R sessions independent of the rest of procedure.
'
'----------------------------------------------------------------``` VBA
' Attach_Libraries( _
'     CommaSepList As String, _
'     Optional VerifyInstallation As Boolean = False, _
'     Optional ShowErrorMessage As Boolean = True _
' )
'
''    Convenience function to generate R library() commands that
''    include the lib.loc of the local user as an IDE would. Stores
''    the commands in the public variable GlobalLoadLibraries, and
''    then appends them to the start of {ScriptContents} prior to being
''    written to an executable .R file with WriteRunRscript() or
''    CaptureRscriptOutput().
'
''    Optionally set {InstallIfRequired} = True to install packages
''    with UDF Require_Packages() if they're not installed. If choosing
''    this option, the function will return False if there was an
''    an installation failure when a package was not previously installed.
'
'     Attach_Libraries("dplyr, ggplot2") is written to the script as:
''       library('dplyr', lib.loc = Sys.getenv('R_LIBS_USER'));
''       library('ggplot2', lib.loc = Sys.getenv('R_LIBS_USER'));
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
' Require_Packages( _
'     CommaSepPackList As String, _
'     Optional Verbose As Boolean = False, _
'     Optional KeepDebugFiles As Boolean = False, _
'     Optional ShowRscript As Boolean = False, _
'     Optional UseRepo As String = "http://cran.r-project.org" _
' )
'
''   Splits {CommaSepPackList} into individual packages, checks their
''   installation status in the user library folder specified in R with
''   Sys.getenv('R_LIBS_USER'), then installs any missing packages.
'
''   After running the script to install packages, the user library folder
''   is checked again to verify all packages were successfully installed.
''   If there are any missing, False is returned, otherwise, True is returned.
''
'to include the local user's lib.loc in the R library() command.
''   Allows for the verification / handling of package installation prior
''   to running a script, and neccesary in cases where a package requires
''   the restart of an R session which would terminate the Rscript run
''   without having independetly installed packages with a seperate script.
'
''   Written to by Attach_Libraries() appended to scripts ran using
''    as complete commands.
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
' WriteRunRscript( _
'     ScriptContents As String, _
'     Optional RscriptVisibility As RsVisibility = RsMinimized, _
'     Optional PreserveScriptFile As Boolean = False, _
'     Optional AttachLibraries As Boolean = True _
' )
'
''    Writes the commands specified by {ScriptContents} to a temporary
''    text file in the downloads folder ("TempExcelScript.R"), passes
''    the path to WinShell_RScript() to execute with Rscript.exe, then
''    deletes the temporary file. Debug.Prints the shell run status.
'
''    If {PreserveScriptFile} = True, TempExcelScript.R won't be deleted
''    and the full filepath will be written to the immediate window.
'
''    If loading libraries with the UDF in this module Attach_Libraries(),
''    parameter {AttachLibraries} must be set to True (default).
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
' CaptureRscriptOutput( _
'     ScriptContents As String, _
'     Optional RscriptVisibility As RsVisibility = RsMinimized, _
'     Optional IncludeInfo As Boolean = False, _
'     Optional PreserveDebugFiles As Boolean = False, _
'     Optional AttachLibraries As Boolean = True _
' )
'
''    Wraps {ScriptContents} in R's sink() command, runs the body of
''    the script, then writes the console output to a text file in the
''    downloads folder. Reads the text file containing the console
''    output into VBA as the return value of the function, then deletes
''    the temporary .txt and .R files unless {PreserveDebugFiles} = False.
'
''    When setting {IncludeInfo} = True, the local user's library folder
''    path (R's Sys.getenv('R_LIBS_USER')) and the time elapsed during the
''    run will also be included.
'
''    If loading libraries with the UDF in this module Attach_Libraries(),
''    parameter {AttachLibraries} must be set to True (default).
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
' ListRLibraries()
'
''   Returns an array of the libraries installed on the local user's
''   device. More specifically, the libraries listed within the
''   directory returned by running "Sys.getenv('R_LIBS_USER'))" in R.
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
' WinShell_RScript( _
'    RScriptExe_Path As String, _
'    Script_Path As String, _
'    Optional RscriptVisibility As RsVisibility = RsVisible _
' )
'
''   Combines the path to R's scripting interpreter {RScriptExe_Path}
''   and the script found at {Script_Path} into a double-escaped
''   command that can be executed by PowerShell.
'
''   Returns the error code of the shell run (0 = successful).
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
' Apple_RScript( _
'    RScriptExe_Path As String, _
'    Script_Path As String, _
'    Optional RscriptVisibility As RsVisibility = RsVisible _
' )
'
''   WIP placeholder for use on MacOS with AppleScriptTask().
'
'----------------------------------------------------------------```
'===============================================================================================================================================================================================================================================================
'##  Supporting Functions
'===============================================================================================================================================================================================================================================================
'----------------------------------------------------------------``` VBA
' Get_RScriptExePath(Optional UseRVersion As String)
'
''   Returns the path to the latest version of Rscript.exe unless a
''   different version is specfied with {UseRVersion}.
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
' Get_LatestRVersionDir(Optional RVersions As Variant)
'
''   Returns the latest version of R installed by evaluating the result
''   of Get_RVersionDirs().
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
' Get_RVersionDirs(Optional RFolderPath As String)
'
''   Returns an array of directory paths with {RFolderPath}.
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
' Get_RFolderDir()
'
''   Returns the directory path of the parent R folder based on the
''   default installation location of the current OS.
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
' ReadLines( _
'    TxtFilePath As String, _
'    Optional ToImmediate As Boolean = True, _
'    Optional ToClipboard As Boolean = True, _
'    Optional Replace_AnyOf As String, _
'    Optional Replace_With As String _
' )
'
''   Reads the text file specified into VBA with each line sperated
''   by vbNewLine to present the contents as they would be seen in a
''   text editor.
'
''   Optionally use {Replace_AnyOf} to specify *characters* to substitute
''   in the result according to {Replace_With}. Terms are replaced similar
''   to R's gsub() in contrast to VBA's replace().
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
' WriteScript( _
'    TextContents As String, _
'    SaveToDir As String, _
'    Optional OverWrite As Boolean = False, _
'    Optional ScriptName As String = "Script.R" _
' )
'
''   Writes a UTF-8 .txt file containing {TextContents} which can be
''   executed by Rscript.
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
' Clipboard_Load(ByVal LoadStr As String)
'
''   Copies {LoadStr} to clipboard.
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
' Get_DownloadsDir()
'
''   Reads Environ("USERPROFILE") to specify the local downloads
''   directory path.
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
' PlatformFileSep()
'
''   Simply returns "\" or "/" depending on the local OS.
'
'----------------------------------------------------------------```
'----------------------------------------------------------------``` VBA
' MyOS()
'
''   Returns "Windows" or "Mac".
'
'----------------------------------------------------------------```
'===============================================================================================================================================================================================================================================================
'##  Examples / Run Testing
'===============================================================================================================================================================================================================================================================
'----------------------------------------------------------------``` VBA
' Example1()
'
''   Should return a tibble and lm summary based on the diamonds
''   data from ggplot2.
'
' Example2()
'
''   Should install and reference the gbm, MASS, dplyr packages,
''   then run a gbm model based on the Boston dataset from MASS.
'
' VerifyReturnValues()
'
''   Should print system information consistent with the local user
''   to the immediate window.
'
'----------------------------------------------------------------```

Sub Example1()
Dim Output As String
    
    Call Attach_Libraries("ggplot2")
    
    Output = CaptureRscriptOutput( _
        " diamonds; " & _
        " mod <- lm(price ~ carat + depth, data = diamonds); " & _
        " mod; " _
    )
    
Debug.Print Output
End Sub

Sub Example2()
Dim Output As String
    
    Call Require_Packages("gbm, MASS, dplyr")
    
    Call Attach_Libraries("gbm, MASS, dplyr")
    Output = CaptureRscriptOutput( _
        " tibble(Boston); " & _
        " train <- sample(1:nrow(Boston), nrow(Boston)*0.7); " & _
        " mod <- gbm(medv ~ ., data = Boston[train,], interaction.depth = 4); " & _
        " summary(mod); " _
    )

Debug.Print Output
End Sub

Sub VerifyReturnValues()

Dim FxArray As Variant, Fx As Variant
    FxArray = Split("MyOS, PlatformFileSep, Get_DownloadsDir, Get_RFolderDir, Get_RScriptExePath", ", ")

    For Each Fx In FxArray
        Debug.Print Fx & ":" & vbNewLine & ">> " & Application.Run("'" & Fx & "'") & vbNewLine
    Next Fx

End Sub

Function Attach_Libraries( _
    CommaSepLibList As String, _
    Optional InstallIfRequired As Boolean = False, _
    Optional ShowErrorMessage As Boolean = True _
)

If InstallIfRequired = True Then
    Dim InstallStatus As Boolean: InstallStatus = Require_Packages(CommaSepLibList)
    If InstallStatus = False Then
        GoTo InstallError
    End If
End If

Dim RLibrary As Variant, ReferenceLibScript As String

For Each RLibrary In Split(Application.Trim(CommaSepLibList), ", ")
    ReferenceLibScript = ReferenceLibScript & "library('" & RLibrary & "', lib.loc = Sys.getenv('R_LIBS_USER')); "
Next RLibrary

GlobalLoadLibraries = ReferenceLibScript
Attach_Libraries = True
Exit Function

InstallError:
Attach_Libraries = False

If ShowErrorMessage = True Then
    Call MsgBox( _
        "You can troubleshoot with the following parameters for " & vbNewLine & vbNewLine & _
        "CaptureRscriptOutput( " & vbNewLine & _
        "   ... " & vbNewLine & _
        "   {IncludeInfo}:=True " & vbNewLine & _
        "   {PreserveDebugFiles}:=True " & vbNewLine & _
        ")" & vbNewLine & vbNewLine & _
        "The .R file passed to Rscript.exe and the .txt file of the console output will be in your downloads folder.", _
        vbInformation, _
        "Package Installation Error" _
    )
End If
End Function

Function Require_Packages( _
    CommaSepPackList As String, _
    Optional Verbose As Boolean = False, _
    Optional KeepDebugFiles As Boolean = False, _
    Optional ShowRscript As Boolean = False, _
    Optional UseRepo As String = "http://cran.r-project.org" _
)

Dim ExistingLibraries As Variant, _
    InstallStatus As Boolean, _
    InstallList As String, _
    InstallScript As String, _
    InstallResult As String, _
    RPackage As Variant, _
    StringMatches As Variant, _
    StrMatch As Variant

ExistingLibraries = ListRLibraries()

    For Each RPackage In Split(CommaSepPackList, ", ")
        'Reduce to approximate string matches
        StringMatches = Filter(ExistingLibraries, RPackage, True)
        InstallStatus = False
        
            For Each StrMatch In StringMatches
                'Determine if perfect match exists
                If RPackage = StrMatch Then
                    InstallStatus = True
                End If
            Next StrMatch
                
        If InstallStatus = False Then
            InstallList = InstallList & " " & RPackage
        End If
    Next RPackage
        
For Each RPackage In Split(Application.Trim(InstallList), " ")
    InstallScript = InstallScript & "install.packages('" & RPackage & "', repos = use_repo, lib = lib_path); "
Next RPackage
     
    If InstallScript <> "" Then
        InstallScript = ( _
            "use_repo <- '" & UseRepo & "'; " & _
            "lib_path <- Sys.getenv('R_LIBS_USER'); " & _
            InstallScript _
        )
        InstallResult = CaptureRscriptOutput( _
            ScriptContents:=InstallScript, _
            RscriptVisibility:=IIf(ShowRscript = True, RsVisible, RsHidden), _
            IncludeInfo:=True, _
            PreserveDebugFiles:=KeepDebugFiles, _
            AttachLibraries:=False _
        )
    End If

If Verbose = True Then
    Debug.Print InstallResult
End If
    
'Check libraries again to verify installation status
ExistingLibraries = ListRLibraries()
    
    For Each RPackage In Split(Application.Trim(InstallList), " ")
    
        StringMatches = Filter(ExistingLibraries, RPackage, True)
        InstallStatus = False
        
            For Each StrMatch In StringMatches
                If RPackage = StrMatch Then
                    InstallStatus = True
                End If
            Next StrMatch
                
        If InstallStatus = False Then
            Exit For
        End If
    Next RPackage

Require_Packages = InstallStatus

End Function

Sub WriteRunRscript( _
    ScriptContents As String, _
    Optional RscriptVisibility As RsVisibility = RsMinimized, _
    Optional PreserveScriptFile As Boolean = False, _
    Optional AttachLibraries As Boolean = True _
)

ScriptContents = IIf(AttachLibraries = True, GlobalLoadLibraries & ScriptContents, ScriptContents)

'Write {ScriptContents} to a .R file in the downloads folder
Dim ScriptLocation As String: ScriptLocation = _
    WriteScript( _
        TextContents:=ScriptContents, _
        SaveToDir:=Get_DownloadsDir(), _
        OverWrite:=True, _
        ScriptName:="TempExcelScript.R" _
    )
    
'Run RScript and record the result
Dim ResultCode: ResultCode = _
    WinShell_RScript( _
        RScriptExe_Path:=Get_RScriptExePath(), _
        Script_Path:=ScriptLocation, _
        RscriptVisible:=RscriptVisibility _
    )

If PreserveScriptFile = True Then
    'Do not remove the temporary .R file, print it's path
    Debug.Print "PreserveScriptFile:=True" & vbNewLine & ">> " & ScriptLocation & vbNewLine
Else 'Remove the temporary .R file {ScriptLocation}
    Call Kill(ScriptLocation)
End If

'Print if the call of the run to Shell was successful
Debug.Print "Shell Run Status" & vbNewLine & ">> " & IIf(ResultCode = 0, "Successful", "Unsuccessful") & vbNewLine
    
End Sub

Function CaptureRscriptOutput( _
    ScriptContents As String, _
    Optional RscriptVisibility As RsVisibility = RsMinimized, _
    Optional IncludeInfo As Boolean = False, _
    Optional PreserveDebugFiles As Boolean = False, _
    Optional AttachLibraries As Boolean = True _
)

ScriptContents = IIf(AttachLibraries = True, GlobalLoadLibraries & ScriptContents, ScriptContents)

Dim DebugTxtDir As String: DebugTxtDir = Get_DownloadsDir() & PlatformFileSep()
Dim DebugTxtName As String: DebugTxtName = "Debug"

'...\Users\Downloads\Debug.R
Dim Path_DebugScript As String: Path_DebugScript = DebugTxtDir & DebugTxtName & ".R"

'...\Users\Downloads\Debug.txt (embedded in {ScriptContents})
Dim Path_DebugOutputTxt As String: Path_DebugOutputTxt = DebugTxtDir & DebugTxtName & ".txt"

Dim SheetFX As Object: Set SheetFX = Application.WorksheetFunction
Dim ArrayWrap As Variant

'Encase {ScriptContents} in additional R code to capture output
If IncludeInfo = False Then
    ArrayWrap = Array( _
        "DebugTxtPath <- r'(" & Path_DebugOutputTxt & ")'", _
        "file_connection <- file(DebugTxtPath)", _
        "sink(file_connection, append=TRUE)", _
        "sink(file_connection, append=TRUE, type='message')", _
          ScriptContents, _
        "sink() # Stop recording console output", _
        "sink(type='message')" _
    )
Else
    ArrayWrap = Array( _
        "DebugTxtPath <- r'(" & Path_DebugOutputTxt & ")'", _
        "file_connection <- file(DebugTxtPath)", _
        "sink(file_connection, append=TRUE)", _
        "sink(file_connection, append=TRUE, type='message')", _
        "message('NOTE: The look of messages are as follows:')", "message('')", _
        "print('This was shown with print()')", _
        "message('This was shown with message()')", "message('')", _
        "message('The R libraries for the user are located here:')", _
        "message(Sys.getenv('R_LIBS_USER'))", "message('')", _
        "TimeStamp <- Sys.time()", _
        "message(rep('=', 75))", _
        "message(paste0('The output of your script begins here (', format(Sys.time(), '%b %d %X'), ')'))", _
        "message(rep('=', 75))", _
        "message('')", _
          ScriptContents, _
        "message('')", _
        "message(rep('=', 75))", _
        "message(paste0('The output of your script ends here (', format(Sys.time(), '%b %d %X'), ')'))", _
        "message(rep('=', 75))", _
        "message('Run Successful')", _
        "message('Time Elapsed: ', round(difftime(Sys.time(), TimeStamp, units = 'secs'), 0), ' seconds')", _
        "sink() # Stop recording console output", _
        "sink(type='message')" _
    )
End If

ScriptContents = SheetFX.TextJoin(vbNewLine, True, ArrayWrap)

Open Path_DebugScript For Output As #1
Print #1, ScriptContents
Close #1

'Run the {ScriptContents}, which will also write a text file to the downloads folder
Dim WinShellResult As Integer
    WinShellResult = WinShell_RScript( _
        RScriptExe_Path:=Get_RScriptExePath(), _
        Script_Path:=Path_DebugScript, _
        RscriptVisibility:=RscriptVisibility _
    )

If IncludeInfo = True And WinShellResult = 1 Then
    Debug.Print "Shell Failed To Run"
End If

'After Rscript has finished running and writing the .txt file, read it
CaptureRscriptOutput = ReadLines( _
    TxtFilePath:=Path_DebugOutputTxt, _
    ToImmediate:=False, _
    ToClipboard:=False, _
    Replace_AnyOf:="”€âœÃ", _
    Replace_With:="-" _
) 'Replace common tidyverse characters loaded incorrectly

'Delete the debug files
If PreserveDebugFiles <> True Then
    Call Kill(Path_DebugScript)
    Call Kill(Path_DebugOutputTxt)
End If
                    
End Function

Function ListRLibraries()

Dim RLibraryString As String
    
    RLibraryString = CaptureRscriptOutput( _
        "message(paste(dir(Sys.getenv('R_LIBS_USER')), collapse = ','))", _
        RscriptVisibility:=RsHidden _
    )
        ListRLibraries = Split(RLibraryString, ",")

End Function

Function Apple_RScript( _
    RScriptExe_Path As String, _
    Script_Path As String, _
    Optional RscriptVisibility As RsVisibility = RsVisible _
)

'Testing on MacOS WIP
'AppleScriptTask()

End Function

Function WinShell_RScript( _
    RScriptExe_Path As String, _
    Script_Path As String, _
    Optional RscriptVisibility As RsVisibility = RsVisible _
)
    
Dim WinShell As Object, _
    ErrorCode As Integer, _
    Escaped_RScriptExe As String, _
    Escaped_Script As String, _
    RShellCommand As String
    
Dim WaitTillComplete As Boolean: WaitTillComplete = True
Set WinShell = CreateObject("WScript.Shell")
        
    Escaped_RScriptExe = Chr(34) & Replace(RScriptExe_Path, "\", "\\") & Chr(34)
    Escaped_Script = Chr(34) & Replace(Script_Path, "\", "\\") & Chr(34)
    RShellCommand = Escaped_RScriptExe & Escaped_Script
    
    ErrorCode = WinShell.Run(RShellCommand, RscriptVisibility, WaitTillComplete)
    WinShell_RScript = ErrorCode
        
Set WinShell = Nothing
End Function

Function Get_RScriptExePath(Optional UseRVersion As String)

    If UseRVersion = "" Then
       UseRVersion = Get_LatestRVersionDir()
    End If
    
    If InStr(1, Environ("OS"), "Windows") > 0 Then
        Get_RScriptExePath = UseRVersion & "bin\Rscript.exe"
    Else
        'https://www.amirmasoudabdol.name/embedding-rframework-in-a-qt-mac-app-and-cross-compiling-for-two-architectures/
        Get_RScriptExePath = UseRVersion & "Resources/bin/R.sh"
        'Testing on MacOS WIP
    End If

End Function

Function Get_LatestRVersionDir(Optional RVersions As Variant)

Dim i As Integer

    If IsMissing(RVersions) Then
        RVersions = Get_RVersionDirs()
    End If
    
    For i = LBound(RVersions) To UBound(RVersions)
        If Get_LatestRVersionDir < RVersions(i) Then
           Get_LatestRVersionDir = RVersions(i)
        End If
    Next i
    
End Function

Function Get_RVersionDirs(Optional RFolderPath As String)

Dim FileSep As String, _
    DirPaths As String, _
    NextPath As String, _
    ArrMatches() As Variant
    
    FileSep = PlatformFileSep()
    RFolderPath = IIf(RFolderPath = "", Get_RFolderDir(), RFolderPath)
    DirPaths = Dir(RFolderPath, vbDirectory)
    
    If DirPaths <> vbNullString Then
        NextPath = RFolderPath & DirPaths
        ReDim Preserve ArrMatches(0): ArrMatches(0) = RFolderPath & DirPaths & FileSep
    Else
        GoTo NoFiles
    End If

    Do While DirPaths <> vbNullString
        NextPath = RFolderPath & DirPaths
            ReDim Preserve ArrMatches(UBound(ArrMatches) + 1): ArrMatches(UBound(ArrMatches)) = RFolderPath & DirPaths & FileSep
        DirPaths = Dir()
    Loop
        
On Error GoTo NoFiles

'Filter out C:\Program Files\R\.. & C:\Program Files\R\.
If LBound(ArrMatches()) = 0 Then Get_RVersionDirs = Filter(ArrMatches, FileSep & ".", False)
Exit Function
    
NoFiles: Get_RVersionDirs = vbNullString
End Function

Function Get_RFolderDir()
    Select Case InStr(1, Environ("OS"), "Windows") > 0
        Case True: Get_RFolderDir = "C:\Program Files\R\"
        Case Else: Get_RFolderDir = "/Library/Frameworks/R.framework/Versions/"
    End Select
End Function

Function ReadLines( _
    TxtFilePath As String, _
    Optional ToImmediate As Boolean = True, _
    Optional ToClipboard As Boolean = True, _
    Optional Replace_AnyOf As String, _
    Optional Replace_With As String _
)

Dim SheetFX As Object: Set SheetFX = Application.WorksheetFunction
Dim FileNum As Integer: FileNum = FreeFile
Dim TxtFileContents As String
Dim TxtFileLines() As String
    
    Open TxtFilePath For Input As FileNum
        TxtFileLines = Split(Input$(LOF(FileNum), FileNum), vbNewLine)
    Close FileNum

    If UBound(TxtFileLines) <> -1 Then
        TxtFileContents = SheetFX.TextJoin(vbNewLine, False, TxtFileLines)
    Else
        TxtFileContents = "(Text File Empty)"
    End If
     
    Do While Len(Replace_AnyOf) <> 0
        TxtFileContents = Replace(TxtFileContents, Left(Replace_AnyOf, 1), Replace_With)
        Replace_AnyOf = Right(Replace_AnyOf, Len(Replace_AnyOf) - 1)
    Loop
    
    If ToImmediate = True Then
        Debug.Print TxtFileContents
    End If
    
    If ToClipboard = True Then
        Call Clipboard_Load(TxtFileContents)
    End If
    
ReadLines = TxtFileContents
    
Set SheetFX = Nothing
End Function

Function WriteScript( _
    TextContents As String, _
    SaveToDir As String, _
    Optional OverWrite As Boolean = False, _
    Optional ScriptName As String = "Script.R" _
)

Dim FileSep As String: FileSep = PlatformFileSep()

'Add FileSep to directory string if required
If Right(SaveToDir, 1) <> FileSep Then SaveToDir = SaveToDir & FileSep

If OverWrite <> True Then
    If Dir(SaveToDir & ScriptName) <> vbNullString Then
        Dim i As Integer, SplitName As Variant, TryName As String
        For i = 1 To 100
            SplitName = Split(ScriptName, ".")
            TryName = SplitName(0) & " (" & i & ")" & "." & SplitName(1)
            If Dir(SaveToDir & TryName) = vbNullString Then
                ScriptName = TryName
                Exit For
            End If
        Next i
    End If
End If

Open SaveToDir & ScriptName For Output As #1
Print #1, TextContents
Close #1

WriteScript = CStr(SaveToDir & ScriptName)
End Function

Function Clipboard_Load(ByVal LoadStr As String)

On Error GoTo NoLoad
    CreateObject("HTMLFile").ParentWindow.ClipboardData.SetData "text", LoadStr
    Clipboard_Load = True
    Exit Function
    
NoLoad:
Clipboard_Load = False
On Error GoTo -1

End Function

Function Get_DownloadsDir()
    Get_DownloadsDir = Environ("USERPROFILE") & PlatformFileSep() & "Downloads" & PlatformFileSep()
End Function

Function PlatformFileSep()
    PlatformFileSep = IIf(InStr(1, Environ("OS"), "Windows") > 0, "\", "/")
End Function

Function MyOS()
Dim EnvOS As String: EnvOS = Environ("OS")
    If InStr(1, EnvOS, "Windows") > 0 Then
        MyOS = "Windows"
    Else
        MyOS = "Mac"
    End If
End Function



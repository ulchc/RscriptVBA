
#  RscriptVBA ¬ github.com/ulchc (10-29-22)


## Overview


Locate RScript 's path automatically, manage packages, run R code contained
in a cell value, and read console output right into Excel's immediate window,
with no additional set up beyond having R installed.'

<a href='https://github.com/ulchc/RscriptVBA'><img src='figures/Example.gif' align="right" height="750" /></a>


## Public Variables

``` VBA
Public GlobalLoadLibraries As String

'   Stores the sequence of R commands to be appended to the start of
'   a script ran with WriteRunRScript() or CaptureRscriptOutput() when
'   using Attach_Libraries().

```
``` VBA
Public Enum RsVisibility
    RsHidden = 0
    RsVisible = 1
    RsMinimized = 2
End Enum

'   Options to toggle the visibilty of the R scripting window.

```

##  Examples / Run Testing

``` VBA
 Example1()

'   Should return a tibble and lm summary based on the diamonds
'   data from ggplot2.

 Example2()

'   Should install and reference the gbm, MASS, dplyr packages,
'   then run a gbm model based on the Boston dataset from MASS.

 VerifyReturnValues()

'   Should print system information consistent with the local user
'   to the immediate window.

```

##  Main Functions & Subs


Note: Although libraries or package installation commands can be directly
written into a script, it is preferable to use UDF's Attach_Libraries() or
Require_Packages() in VABA so that packages which force the restart of an R
session can be installed in R sessions independent of the rest of procedure.

``` VBA
 Attach_Libraries( _
     CommaSepList As String, _
     Optional VerifyInstallation As Boolean = False, _
     Optional ShowErrorMessage As Boolean = True _
 )

'    Convenience function to generate R library() commands that
'    include the lib.loc of the local user as an IDE would. Stores
'    the commands in the public variable GlobalLoadLibraries, and
'    then appends them to the start of {ScriptContents} prior to being
'    written to an executable .R file with WriteRunRScript() or
'    CaptureRscriptOutput().

'    Optionally set {InstallIfRequired} = True to install packages
'    with UDF Require_Packages() if they're not installed. If choosing
'    this option, the function will return False if there was an
'    an installation failure when a package was not previously installed.

     Attach_Libraries("dplyr, ggplot2") is written to the script as:
'       library('dplyr', lib.loc = Sys.getenv('R_LIBS_USER'));
'       library('ggplot2', lib.loc = Sys.getenv('R_LIBS_USER'));

```
``` VBA
 Require_Packages( _
     CommaSepPackList As String, _
     Optional Verbose As Boolean = False, _
     Optional KeepDebugFiles As Boolean = False, _
     Optional ShowRscript As Boolean = False, _
     Optional UseRepo As String = "http://cran.r-project.org" _
 )

'   Splits {CommaSepPackList} into individual packages, checks their
'   installation status in the user library folder specified in R with
'   Sys.getenv('R_LIBS_USER'), then installs any missing packages.

'   After running the script to install packages, the user library folder
'   is checked again to verify all packages were successfully installed.
'   If there are any missing, False is returned, otherwise, True is returned.
'
to include the local user's lib.loc in the R library() command.
'   Allows for the verification / handling of package installation prior
'   to running a script, and neccesary in cases where a package requires
'   the restart of an R session which would terminate the Rscript run
'   without having independetly installed packages with a seperate script.

'   Written to by Attach_Libraries() appended to scripts ran using
'    as complete commands.

```
``` VBA
 WriteRunRScript( _
     ScriptContents As String, _
     Optional RscriptVisibility As RsVisibility = RsMinimized, _
     Optional PreserveScriptFile As Boolean = False, _
     Optional AttachLibraries As Boolean = True _
 )

'    Writes the commands specified by {ScriptContents} to a temporary
'    text file in the downloads folder ("TempExcelScript.R"), passes
'    the path to WinShell_RScript() to execute with Rscript.exe, then
'    deletes the temporary file. Debug.Prints the shell run status.

'    If {PreserveScriptFile} = True, TempExcelScript.R won't be deleted
'    and the full filepath will be written to the immediate window.

'    If loading libraries with the UDF in this module Attach_Libraries(),
'    parameter {AttachLibraries} must be set to True (default).

```
``` VBA
 CaptureRscriptOutput( _
     ScriptContents As String, _
     Optional RscriptVisibility As RsVisibility = RsMinimized, _
     Optional IncludeInfo As Boolean = False, _
     Optional PreserveDebugFiles As Boolean = False, _
     Optional AttachLibraries As Boolean = True _
 )

'    Wraps {ScriptContents} in R's sink() command, runs the body of
'    the script, then writes the console output to a text file in the
'    downloads folder. Reads the text file containing the console
'    output into VBA as the return value of the function, then deletes
'    the temporary .txt and .R files unless {PreserveDebugFiles} = False.

'    When setting {IncludeInfo} = True, the local user's library folder
'    path (R's Sys.getenv('R_LIBS_USER')) and the time elapsed during the
'    run will also be included.

'    If loading libraries with the UDF in this module Attach_Libraries(),
'    parameter {AttachLibraries} must be set to True (default).

```
``` VBA
 ListRLibraries()

'   Returns an array of the libraries installed on the local user's
'   device. More specifically, the libraries listed within the
'   directory returned by running "Sys.getenv('R_LIBS_USER'))" in R.

```
``` VBA
 WinShell_RScript( _
    RScriptExe_Path As String, _
    Script_Path As String, _
    Optional RscriptVisibility As RsVisibility = RsVisible _
 )

'   Combines the path to R's scripting interpreter {RScriptExe_Path}
'   and the script found at {Script_Path} into a double-escaped
'   command that can be executed by PowerShell.

'   Returns the error code of the shell run (0 = successful).

```
``` VBA
 Apple_RScript( _
    RScriptExe_Path As String, _
    Script_Path As String, _
    Optional RscriptVisibility As RsVisibility = RsVisible _
 )

'   WIP placeholder for use on MacOS with AppleScriptTask().

```

##  Supporting Functions

``` VBA
 Get_RScriptExePath(Optional UseRVersion As String)

'   Returns the path to the latest version of Rscript.exe unless a
'   different version is specfied with {UseRVersion}.

```
``` VBA
 Get_LatestRVersionDir(Optional RVersions As Variant)

'   Returns the latest version of R installed by evaluating the result
'   of Get_RVersionDirs().

```
``` VBA
 Get_RVersionDirs(Optional RFolderPath As String)

'   Returns an array of directory paths with {RFolderPath}.

```
``` VBA
 Get_RFolderDir()

'   Returns the directory path of the parent R folder based on the
'   default installation location of the current OS.

```
``` VBA
 ReadLines( _
    TxtFilePath As String, _
    Optional ToImmediate As Boolean = True, _
    Optional ToClipboard As Boolean = True, _
    Optional Replace_AnyOf As String, _
    Optional Replace_With As String _
 )

'   Reads the text file specified into VBA with each line sperated
'   by vbNewLine to present the contents as they would be seen in a
'   text editor.

'   Optionally use {Replace_AnyOf} to specify *characters* to substitute
'   in the result according to {Replace_With}. Terms are replaced similar
'   to R's gsub() in contrast to VBA's replace().

```
``` VBA
 WriteScript( _
    TextContents As String, _
    SaveToDir As String, _
    Optional OverWrite As Boolean = False, _
    Optional ScriptName As String = "Script.R" _
 )

'   Writes a UTF-8 .txt file containing {TextContents} which can be
'   executed by Rscript.

```
``` VBA
 Clipboard_Load(ByVal LoadStr As String)

'   Copies {LoadStr} to clipboard.

```
``` VBA
 Get_DownloadsDir()

'   Reads Environ("USERPROFILE") to specify the local downloads
'   directory path.

```
``` VBA
 PlatformFileSep()

'   Simply returns "\" or "/" depending on the local OS.

```
``` VBA
 MyOS()

'   Returns "Windows" or "Mac".

```

<#

Getting started:

    1. Download ExifTool "Window Executable" from: http://www.sno.phy.queensu.ca/~phil/exiftool/
        Extact to same folder as this script (or chage configuration to point correctly)
        Rename to "ExifTool.exe"

    2. Download the "Portable Win64 static" from http://www.imagemagick.org/script/binary-releases.php
        Extract to a subfolder called imagemagick located in the same folder as this script (or chage configuration to point correctly)

    3. Download TrID from http://mark0.net/soft-trid-e.html
        Go to bottom of page and download Win32 zip + TrIDDefs
        Extract both to same folder as this script

    4. Check the configuration just below

    5. Edit bottom of this file to configure how the workflow is started

    6. In PowerShell ISE: press F5 to run the script

#>

### configuration #############################################################

$outRootFolder    = "t:\iwf\"                 # the base output folder
$outputFolders    = @{
    "archive"     = "$outRootFolder\out_archive";  # if $archiveSourceFiles is true, all processed files will be moved here after the workflow has run.
    "error"       = "$outRootFolder\out_error";
    "wrongext"    = "$outRootFolder\out_wrong_ext";
    "jpg"         = "$outRootFolder\out_jpg";
    "gif"         = "$outRootFolder\out_gif_static";
    "gifanim"     = "$outRootFolder\out_gif_active";
    "logs"        = "$outRootFolder\out_logs";     # one csv file is produced and put here
    "nonimg"      = "$outRootFolder\out_non_img";  # if $moveNonImageFiles is true, this is where any unhandled file type is moved if passed to the script.

    # all below can be inactivated by setting to ""
    "light"       = "$outRootFolder\out_light";    # less than $smallWeight
    "heavy"       = "$outRootFolder\out_heavy";    # more than $maxWeight
    "small"       = "$outRootFolder\out_small";    # less than $smallWidth or $smallHeight
    "not_srgb"    = "$outRootFolder\oddset_not_srgb";
    "interlace"   = "$outRootFolder\oddset_progressive";
    "not_8bit"    = "$outRootFolder\oddset_not_8bit";
    "low_quality" = "$outRootFolder\oddset_low_quality"; # under $lowQualityThreshold below
    "icc_not_rgb" = "$outRootFolder\oddset_icc_not_rgb"; # see config below
    "icc_not_srgb"= "$outRootFolder\oddset_icc_not_srgb";
    "special_jpg" = "$outRootFolder\oddset_special_jpg"; # Detected: Progressive, Arithmetic, Lossless, Hierarchical
}

$validExtensions = 		          ".jpg", ".jpeg", ".jpe", ".png", ".tif", ".tiff", ".gif", ".bmp",
				                  ".jp2", ".jpf", ".jpx", ".j2c", ".j2k", ".jpc", ".webp", ".psd"

$nonRgbColorProfiles =            "Europe ISO Coated FOGRA27", "Euroscale Coated v2", "Euroscale Uncoated v2", 
                                  "Japan Color 2001 Coated", "Japan Color 2001 Uncoated", "Japan Color 2002 Newspaper", 
                                  "Japan Web Coated(Ad)", "U.S. Sheetfed Coated v2", "U.S. Sheetfed Uncoated v2", 
                                  "U.S. Web Coated (SWOP) v2", "U.S. Web Uncoated v2", "Agfa : Swop Standard", 
                                  "sGray", "Dot Gain 10%", "Dot Gain 15%", "Dot Gain 20%", "Dot Gain 25%", 
                                  "Dot Gain 30%", "Gray Gamma 1.8", "Gray Gamma 2.2"

$moveNonImageFiles = $true        # Files that do not match any of the valid image extensions 
				                  #   will be moved to folder specified in $outputFolders.nonimg
$archiveSourceFiles = $true       # Source files will be moved to _archive folder. Exception: Files with errors
				                  #   and GIFs are ALWAYS moved to their output folders.
$deleteSourceFiles = $false       # Source files will be deleted provided there was no error.
				                  #   Takes presidence over $archiveSourceFiles
$renameOutputFiles = $true        # When $false; files are not renamed using the $renames rules below
$renameArchivedFiles = $true      # When enabled, source files moved to the archive folder will be renamed to 
                                  #   match the name of the produced output file. (Note that if the filename 
                                  #   already exists, a number will be appended and might cause a mismatch.)
$overwriteExisting = $false       # When copying/moving files to destination folders, for filename conflicts;
				                  #   if $false, file being copied is renamed/appended a number: "filename (1).jpg";
				                  #   if $true, file in the dest. folder is overwritten.
$fixOriginalNamesFirst = $true    # Renames the original files before the workflow starts. 
				                  #   This prevents problems later due to non-ascii filenames passed to ExifTool, etc.
				                  #   Files are renamed in a smart way, e.g. ä -> a (This is NOT the same rename as takes place later on,
                                  #   using the $renames config, below.
$validateExtOnErrors = $true      # When enabled, TrID is used on all files flaged as erroneous by the workflow after it has 
                                  #   finished. TrID will try and detect any incorrect file extensions and fix it. Files that are renamed 
                                  #   will be moved the wrongext output folder.
$overwriteOriginal = $false       # Overwrite original file instead of putting in output folders. 
$exifToolBatchSize = 500          # How many files to send to ExifTool in each batch (0 for all).
$printIndividualErrors = $false   # If errors occur when running external program, should each error be printed in 
                                  # red afterwards?
$readAdvancedJpgProps = $true     # Tries to identify "odd" JPG properties


# Rules when processing images

$maxWeight = 5MB       # logic of "max" items is > ("greater than")
$maxWidth = 4999
$maxHeight = 4999

$smallWeight = 32KB    # logic of "small" items is < ("less than")
$smallWidth = 401
$smallHeight = 401

$lowQualityThreshold = 22   # quality is estimated by ImageMagick's identify utility
$highQualityThreshold = 91  # Will be resaved if above this. (quality is estimated by ImageMagick's identify utility)

# Rules for cleaning up filenames

$maxFilenameLength = 32        # Does not include extension (.ext)
$allowedChars = 'a-zA-Z0-9_'

$renames = [ordered]@{                       # RegEx rules for renaming / cleaning up output filenames
    'find something' = 'replace with this';  # Add new rules by simply adding a line like this.
    'removethis' = '';                       # Leave the find string empty to remove.
    '[-]' = '_';                             # Replace dash with underscore
    '[ ]' = '_';                             # Replace space with underscore
    '0+' = '0';                              # Shrink multiple zero to one
    "[^$allowedChars]" = '';                 # Remove unwanted characters (defined above)
    '_{2,}' = '_';                           # Replace two or more underscores with single one
    '^_' = '';                               # Remove underscore at beginning
    '_$' = '';                               # Remove underscore at end
    "(^.{$maxFilenameLength}).*" = '$1';     # Crop to x number of characters. Best to keep this last.
}

$noName = 'NoName' # If after renaming the filename is empty, it will be given this name
$existingNameSuffix = '_#' # If a filename already exists, this is added to the end of the filename. The # is replaced with a number. Example: " (#)"

### External Applications #####################################################

$scriptPath = Split-Path $MyInvocation.MyCommand.Path
$exeTimeout = (5 * 60 * 1000) # In milliseconds. To prevent script from hanging, how long to wait before stopping the process. Applies to all extenal programs
$printExecutingCommands = $false

# ImageMagick configuration:
# Info: http://www.imagemagick.org/Usage/formats/#jpg_write
#   Colorspace: ImageMagick assumes the sRGB colorspace if the image format does not indicate otherwise. 
#   Calculation of optimal Huffman coding tables is on by default
# "-quality 90%" approximately matches PS9
#   If not set: The default is to use the estimated quality of your input image if it can be determined, otherwise 92
# "-sampling-factor" can be: "1x1", "1x2" or "2x2" where 1x1 best quality. (IM default chroma sub-sampling is 2x2 and corresponds to 4:2:0)
#   When the quality is greater than 90, then the chroma channels are not downsampled
# Adding "-define jpeg:extent=400kb" will force ImageMagick to produce jpeg of this size. Very slow!
# Advanced configurations we would consider:
#   Add an unsharp
#   Change the resampling filter used
#   (very advanced) Use a custom custom quantization table by using "-define jpeg:q-table={path}"
$im_resize = ('-resize "' + $maxWidth + 'x' + $maxHeight + '>"') #this will resize to max W/H oly if image is larger. ref: http://www.imagemagick.org/Usage/resize/#shrink
$imageMagickConvertArgs = "$im_resize -colorspace RGB -quality 83% -sampling-factor 1x1"      # when converting "other" formats to jpg
$imageMagickScaleDownArgs = "$im_resize -colorspace RGB -quality 83% -sampling-factor 1x1"    # when scaling down images larger that maxWidth/maxHeight

$imageMagickConvertExe = ($scriptPath + "\imagemagick\convert.exe")
$imageMagickMogrifyExe = ($scriptPath + "\imagemagick\mogrify.exe")
$imageMagickIdentifyExe = ($scriptPath + "\imagemagick\identify.exe")
$imageMagickConvertArgsPrefix = ""
$imageMagickScaleDownArgsPrefix = ""

<# Original
$imageMagickConvertExe = ($scriptPath + "\imagemagick\convert.exe")
$imageMagickMogrifyExe = ($scriptPath + "\imagemagick\mogrify.exe")
$imageMagickConvertArgsPrefix = ""
$imageMagickScaleDownArgsPrefix = ""
#>

<# Q8
$imageMagickConvertExe = ($scriptPath + "\ImageMagick-7.0.2-Q8\magick.exe")
$imageMagickMogrifyExe = ($scriptPath + "\ImageMagick-7.0.2-Q8\magick.exe ")
$imageMagickConvertArgsPrefix = "convert "
$imageMagickScaleDownArgsPrefix = "mogrify "
#>

<# GraphicsMagick
$imageMagickConvertExe = ($scriptPath + "\GraphicsMagick-1.3.24-Q8\gm.exe")
$imageMagickMogrifyExe = ($scriptPath + "\GraphicsMagick-1.3.24-Q8\gm.exe")
$imageMagickConvertArgsPrefix = "convert "
$imageMagickScaleDownArgsPrefix = "mogrify "
#>


$exifToolExe = ($scriptPath + "\exiftool.exe")
$exifToolArgFile = [System.IO.Path]::GetTempFileName()
$exifToolCmd = '-charset', 'FILENAME=UTF8', '-all=', '-CommonIFD0=', '-m', '@input', '-P', '-o', '@output', '-execute'

$tridExe = ($scriptPath + '\trid.exe')


### Verify dependencies exist#################################################

if( -not (Test-Path -LiteralPath $imageMagickConvertExe)) {
    Write-Host "ImageMagick Convert exe not found: $imageMagickConvertExe" -ForegroundColor White -BackgroundColor Red
    Exit
}
if( -not (Test-Path -LiteralPath $imageMagickMogrifyExe)) {
    Write-Host "ImageMagick Mogrify exe not found: $imageMagickMogrifyExe" -ForegroundColor White -BackgroundColor Red
    Exit
}
if( -not (Test-Path -LiteralPath $exifToolExe)) {
    Write-Host "Exif Tool exe not found: $exifToolExe" -ForegroundColor White -BackgroundColor Red
    Exit
}
if( -not (Test-Path -LiteralPath $tridExe)) {
    Write-Host "TrID exe not found: $tridExe" -ForegroundColor White -BackgroundColor Red
    Exit
}

### Workflow Cmdlet ###########################################################

$workflowResult = "" # Result of workflow is sent to pipe but also stored in this variable

$timers = @{
    "main" = New-Object "System.Diagnostics.Stopwatch";
    "imconvert" = New-Object "System.Diagnostics.Stopwatch";
    "improcess" = New-Object "System.Diagnostics.Stopwatch";
    "props" = New-Object "System.Diagnostics.Stopwatch";
    "exiftool" = New-Object "System.Diagnostics.Stopwatch";
    "fileio" = New-Object "System.Diagnostics.Stopwatch";
    "checkext" = New-Object "System.Diagnostics.Stopwatch";
    "advprops" = New-Object "System.Diagnostics.Stopwatch";
}

function Start-Workflow
{
    param(  
        [Parameter(Position=0, 
                   Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true)]
        $ImageFiles
    ) 

    BEGIN 
    {
        $timers.main.Start()
        Write-Host "Parsing input..." -ForegroundColor Yellow
        Write-Host "`tReading pipeline..."
        $files = @()
        $filesFromPipeline = @()
    }

    PROCESS
    {
        foreach($f in $ImageFiles)
        {
            if($f.GetType().FullName -eq "System.IO.FileInfo") {
                $filesFromPipeline += $f.FullName
            } elseif($if.GetType().FullName -eq "System.String") {
                $filesFromPipeline += $f.ToString()
            }
        }
    }

    END
    {
        Write-Host "`tCollecting files..."

        #collect files
        $timers.fileio.Start()
        foreach($f in $filesFromPipeline)
        {
            $file = $f
            $ext = [System.IO.Path]::GetExtension($f.ToString())

            # Renames the original files before the workflow starts. 
            # This prevents problems later due to non-ascii filenames passed to ExifTool, etc.
			# Files are renamed in a "smart" way, e.g. ä -> a (This is NOT the same rename as takes place later on, using the "real" renaming rules
            if($fixOriginalNamesFirst) {
                $orig = [System.IO.Path]::GetFilenameWithoutExtension($file)
                $fixed = FairlySmartCharReplace $orig
                if($orig -ne $fixed) {
                    $fixed += [System.IO.Path]::GetExtension($file)
                    $file = SmartRename $file $fixed
                }
            }

            if($validExtensions -contains $ext) {
                Write-Verbose ("Adding: " + $file)
                $files += NewWorkflowItem($file)
            } else {
                Write-Verbose ("SKIPPING: " + $f)
                if($moveNonImageFiles) {
                    $from = $file
                    $to = $outputFolders.nonimg
                    MoveFile $from $to
                }
            }
        }

        Write-Host "`tFound" $files.Count "image files"

        # Create out_jpg as it is required (temp files are put there)
        if(!(Test-Path -LiteralPath $outputFolders.jpg)) {
            Write-Verbose ("Creating output folder: " + $outputFolders.jpg)
            New-Item -ItemType Directory -Path $outputFolders.jpg -Force | out-null
        }

        $timers.fileio.Stop()

        #let's run the actual workflow
        if($files.length -gt 0)
        {
            Write-Host "Processing images..." -ForegroundColor Yellow

            GetFirstTimeImageProps $files
            IdentifyOddsetJpegProperties $files

            ConvertToJpeg $files

            GetImageProps $files
            #IdentifyOddsetJpegProperties $files

            SetDestinations $files # Run here because PostProcessing needs to know if destination is not standard

            PostProcessing $files

            GetImageProps $files
            #IdentifyOddsetJpegProperties $files

            RunExifTool $files

            GetImageProps $files

            SetDestinations $files
            SendToDestination $files

        }


        # delete or move original files
        if($deleteSourceFiles)
        {
            $timers.fileio.Start()
            DeleteOriginalFiles($files)
            $timers.fileio.Stop()
        }
        elseif($archiveSourceFiles)
        {
            $timers.fileio.Start()
            ArchiveOriginalFiles($files)
            $timers.fileio.Stop()
        }

        # correct faulty file extensiosn (only erroneous files)
        if($validateExtOnErrors) {
            $timers.checkext.Start()
            CheckExtensionsOnErrors $files
            $timers.checkext.Stop()
        }


        # tidy up PSObject for display
        foreach($f in $files)
        {
            $f.PSObject.Properties.Remove("WorkFile")
            $f.PSObject.Properties.Remove("OldWorkFile")
        }
        

        # finish up

        $logFilename = ($outputFolders.logs + "\ " + (get-date -Format "yyyy-MM-dd HH-mm-ss") + ".csv")
        Write-Host ("Writing log...") -ForegroundColor Yellow
        Write-Host "`t$logFilename"
        New-Item -ItemType Directory -Force -Path $outputFolders.logs | Out-Null
        $files | Export-Csv -Path $logFilename -Encoding UTF8

        $timers.main.Stop()
        write-Host "Running Time:" -ForegroundColor Yellow
        write-Host ("`tTotal: " + (FormatStopwatchTime($timers.main)) )
        write-Host ("`tExifTool: " + (FormatStopwatchTime($timers.exiftool)) )
        write-Host ("`tImageMagick identify advanced jpg props: " + (FormatStopwatchTime($timers.advprops)) )
        write-Host ("`tImageMagick convert to jpg: " + (FormatStopwatchTime($timers.imconvert)) )
        write-Host ("`tImageMagick post processing: " + (FormatStopwatchTime($timers.improcess)) )
        write-Host ("`tReading image properties: " + (FormatStopwatchTime($timers.props)) )
        write-Host ("`tFile operations: " + (FormatStopwatchTime($timers.fileio)) )
        write-Host ("`tCheck extensions: " + (FormatStopwatchTime($timers.checkext)) )

        Write-Host "Ready!" -ForegroundColor Green

        # send result to pipe
        $workflowResult = $files
        $workflowResult

    }
}

###############################################################################

function NewWorkflowItem($source)
{
    $obj = New-Object -TypeName PSObject
    $props = [ordered]@{ 
            Source = $source;
            SourceWidth = 0;
            SourceHeight = 0;
            SourceFileSize = 0;
            SourceFrameCount = 1;

            SourceJpegColorSpace = "N/A";
            SourceJpegMode = "";
            SourceJpegColorBitDepth = "N/A";
            SourceJpegQuality = "N/A";
            SourceJpegIsInterlace = "None"; # None is what ImageMagick reports
            SourceIccProfile = "N/A";

            DestPath = $outputFolders.jpg;
            DestFilename = [System.IO.Path]::GetFileNameWithoutExtension($source);
            DestExt = [System.IO.Path]::GetExtension($source).ToLower();
            DestWidth = 0;
            DestHeight = 0;
            DestFileSize = 0;
            Error = $false;
            ErrorMsg = "";
            WorkFile = $source;
            OldWorkFile = $source; # used for error handling
            UpdateProps = $false;
            UpdateAdvancedJpgProps = $true;
        }
    $obj | Add-Member -NotePropertyMembers $props -TypeName WorkFlowItem

    return $obj
}

###############################################################################

function Repair-FileExtensions
{
    # Use this Cmdlet to run TrID stand-alone.
    # Note: Does not like 'foreign' characters in the filenames

    param(  
        # Param1 help description
        [Parameter(Position=0, 
                   Mandatory=$true, 
                   ValueFromPipeline=$true,
                   ValueFromPipelineByPropertyName=$true)]
        [System.IO.FileInfo[]]
        $InputFiles
    ) 
    BEGIN 
    {
        $files = @()
    }

    PROCESS
    {
        foreach($f in $InputFiles)
        {
            $files += $f
        }
    }

    END
    {
        Write-host ("Checking extensions of " + $files.Count + " files...") -ForegroundColor Yellow
        $extensionsCorrected = 0
        $sw = [System.Diagnostics.Stopwatch]::StartNew()
        Write-Progress -Activity "Checking file extensions..." -Status "0% done" -PercentComplete 0
        for($i=0; $i -lt $files.Count; $i++)
        {
            if ($sw.Elapsed.TotalMilliseconds -ge 500)
            {
                $percentDone = [System.Math]::Round((($i / $files.Count) * 100), 0);
                Write-Progress -Activity "Checking file extensions..." -Status "$percentDone% done" -PercentComplete $percentDone -CurrentOperation $files[$i].FullName
                $sw.Reset(); $sw.Start()
            }
            $e = ExecuteCommand ($scriptPath + '\trid.exe') ('-ce "' + $files[$i].FullName +'"')
            if($e.ExitCode -eq 0) {
                if($e.stdout -like "*1 file(s)*") {
                    $extensionsCorrected++
                    Write-host ("File extensions corrected: " + $files[$i].FullName)
                }
            } else {
                Write-host ("Error checking file: " + $files[$i].FullName) -ForegroundColor White -BackgroundColor Red
            }
        }
        Write-Progress -Activity "Processing images..." -PercentComplete 100 -Completed
        Write-host "$extensionsCorrected file extensions corrected!" -ForegroundColor Yellow
    }
}

###############################################################################

function FormatStopwatchTime([System.Diagnostics.Stopwatch]$stopwatch) {
    return $stopwatch.Elapsed.ToString('hh\:mm\:ss')
}

###############################################################################

Function ExecuteCommand ($commandPath, $commandArguments)
{
    # Wrapper for Invoke-Executable

    if($printExecutingCommands) {
        Write-Host "$commandPath $commandArguments" -ForegroundColor White -BackgroundColor Black
    }

    $result = InvokeExecutable -sExeFile $commandPath -cArgs $commandArguments -TimeoutMilliseconds $exeTimeout
    if($result.TimedOut) {
        Write-Host ("`tProcess timed out!") -ForegroundColor White -BackgroundColor Red
    }
    [pscustomobject]@{
        stdout = $result.StdOut
        stderr = $result.StdErr
        ExitCode = $result.ExitCode  
    }
}

function InvokeExecutable {
    # from http://stackoverflow.com/a/24371479/52277
        # Runs the specified executable and captures its exit code, stdout
        # and stderr.
        # Returns: custom object.
    # from http://www.codeducky.org/process-handling-net/ added timeout, using tasks
    param(
        [Parameter(Mandatory=$true)]
        [ValidateNotNullOrEmpty()]
        [String]$sExeFile,
        [Parameter(Mandatory=$false)]
        [String[]]$cArgs,
        [Parameter(Mandatory=$false)]
        [String]$sVerb,
        [Parameter(Mandatory=$false)]
        [Int]$TimeoutMilliseconds=1800000 #30min
    )

    # Setting process invocation parameters.
    $oPsi = New-Object -TypeName System.Diagnostics.ProcessStartInfo
    $oPsi.CreateNoWindow = $true
    $oPsi.UseShellExecute = $false
    $oPsi.RedirectStandardOutput = $true
    $oPsi.RedirectStandardError = $true
    $oPsi.FileName = $sExeFile
    if (! [String]::IsNullOrEmpty($cArgs)) {
        $oPsi.Arguments = $cArgs
    }
    if (! [String]::IsNullOrEmpty($sVerb)) {
        $oPsi.Verb = $sVerb
    }

    # Creating process object.
    $oProcess = New-Object -TypeName System.Diagnostics.Process
    $oProcess.StartInfo = $oPsi
    $processTimedOut = $false

    # Starting process.
    [Void]$oProcess.Start()
    # Tasks used based on http://www.codeducky.org/process-handling-net/    
    $outTask = $oProcess.StandardOutput.ReadToEndAsync();
    $errTask = $oProcess.StandardError.ReadToEndAsync();
    $bRet=$oProcess.WaitForExit($TimeoutMilliseconds)
    if (-Not $bRet)
    {
        $oProcess.Kill();
        #  throw [System.TimeoutException] ($sExeFile + " was killed due to timeout after " + ($TimeoutMilliseconds/1000) + " sec ") 
        $processTimedOut = $true
    }
    $outText = $outTask.Result;
    $errText = $errTask.Result;
    if (-Not $bRet)
    {
        $errText =$errText + ($sExeFile + " was killed due to timeout after " + ($TimeoutMilliseconds/1000) + " sec ") 
    }
    $oResult = New-Object -TypeName PSObject -Property ([Ordered]@{
        "ExeFile"  = $sExeFile;
        "Args"     = $cArgs -join " ";
        "ExitCode" = $oProcess.ExitCode;
        "StdOut"   = $outText;
        "StdErr"   = $errText;
        "TimedOut" = $processTimedOut
    })

    return $oResult
}

###############################################################################

function GetFirstTimeImageProps($files) {

    $timers.props.Start()

    Write-Host "`tReading image properties..."

    $WiaImageFile = New-Object -ComObject Wia.ImageFile

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    Write-Progress -Activity "Reading image properties..." -Status "0% done" -PercentComplete 0
    for($i=0; $i -lt $files.Count; $i++)
    {
        $img = $files[$i]
        if ($sw.Elapsed.TotalMilliseconds -ge 500)
        {
            $percentDone = [System.Math]::Round((($i / $files.Count) * 100), 0);
            Write-Progress -Activity "Reading image properties..." -Status "$percentDone% done" -PercentComplete $percentDone -CurrentOperation $img.Source
            $sw.Reset(); $sw.Start()
        }

        try
        {
            #Write-Host "First time: " + $img.WorkFile
            $WiaImageFile.LoadFile($img.WorkFile)
            $img.SourceWidth = $WiaImageFile.Width
            $img.SourceHeight = $WiaImageFile.Height
            $img.SourceFrameCount = $WiaImageFile.FrameCount
            $img.SourceFileSize = (Get-Item $img.WorkFile).length
            $img.DestFileSize = $img.SourceFileSize
            $img.DestWidth = $img.SourceWidth
            $img.DestHeight = $img.SourceHeight

            # correct wrong ext
            if("." + ($WiaImageFile.FileExtension.ToLower()) -ne ([System.IO.Path]::GetExtension($img.WorkFile).ToLower())) {
                # this will run very early in the script and only first time, so I'm making assumtions here that we can safely change various properties on the file
                $newExt = $WiaImageFile.FileExtension.ToLower()
                $newSource = [System.IO.Path]::ChangeExtension($img.Source, $newExt)
                $newName = [System.IO.Path]::GetFileName($newSource)
                Write-Verbose ("Fixing wrong file extension: '" + $img.Source + "' should be ." + $newExt)
                $newSource = SmartRename $img.Source $newName
                $img.Source = $newSource
                $img.WorkFile = $newSource
                $img.OldWorkFile = $newSource
                $img.DestExt = "." + $newExt
            }

            $img.UpdateProps = $false
        }
        catch
        {
            # let error slip by first time, because next step will convert files to jpg if needed
        }
    }
    Write-Progress -Activity "Reading image properties..." -PercentComplete 100 -Completed

    $timers.props.Stop()
}


function GetImageProps($files)
{
    $timers.props.Start()

    Write-Host "`tUpdating image properties..."

    $WiaImageFile = New-Object -ComObject Wia.ImageFile
    $errorCount = 0

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    Write-Progress -Activity "Updating image properties..." -Status "0% done" -PercentComplete 0
    for($i=0; $i -lt $files.Count; $i++)
    {
        $img = $files[$i]
        if(!$img.Error -and $img.UpdateProps)
        {
            #Write-Host "Update: " + $img.WorkFile
            if ($sw.Elapsed.TotalMilliseconds -ge 500)
            {
                $percentDone = [System.Math]::Round((($i / $files.Count) * 100), 0);
                Write-Progress -Activity "Updating image properties..." -Status "$percentDone% done" -PercentComplete $percentDone -CurrentOperation $img.Source
                $sw.Reset(); $sw.Start()
            }

            try
            {
                $img.UpdateProps = $false
                $img.DestFileSize = (Get-Item $img.WorkFile).length

                $WiaImageFile.LoadFile($img.WorkFile)
                $img.DestWidth = $WiaImageFile.Width
                $img.DestHeight = $WiaImageFile.Height

                # set source if not already done
                if($img.SourceFileSize -eq 0)
                {
                    $img.SourceFileSize = $img.DestFileSize
                    $img.SourceWidth = $WiaImageFile.Width
                    $img.SourceHeight = $WiaImageFile.Height
                    $img.SourceFrameCount = $WiaImageFile.FrameCount
                }

            }
            catch
            {
                $errorCount++
                $img.Error = $true
                $img.ErrorMsg = "Unable to read image properties of file."
                if($printIndividualErrors) {
                    Write-Host ("`tUnable to read image properties of file: " + $img.WorkFile + " (" + $img.Source + ")" ) -ForegroundColor White -BackgroundColor Red
                }
                $img.DestPath = $outputFolders.error
            }
       }
    }
    Write-Progress -Activity "Updating image properties..." -PercentComplete 100 -Completed
    if($errorCount -gt 0) {
        Write-Host ("`t`t" + $errorCount + " images could not be read...") -BackgroundColor Red
    }

    $timers.props.Stop()
}


###############################################################################

function EnsureOutputFolder($filename) {
    $folder = [System.IO.Path]::GetDirectoryName($filename);
    if(!(Test-Path -LiteralPath $folder)) {
        Write-Verbose ("Creating output folder: " + $folder)
        New-Item -ItemType Directory -Path $folder -Force | out-null
    }
}

###############################################################################

function MoveFile($from, $toFolder, $newName)
{
    $DestFilename = [System.IO.Path]::GetFileNameWithoutExtension($from)
    $DestExt = [System.IO.Path]::GetExtension($from).ToLower()

    if($newName -ne $null) {
        $DestFilename = [System.IO.Path]::GetFileNameWithoutExtension($newName)
        $DestExt = [System.IO.Path]::GetExtension($newName).ToLower()
    }
    
    #check existing filenames
    $to = $toFolder + "\" + $DestFilename + $DestExt
    if($overwriteExisting -eq $false) {
        $i = 0
        While (Test-Path -LiteralPath $to) {
            $i += 1
            $suffix = $existingNameSuffix -replace '#',$i
            $to = $toFolder + "\" + $DestFilename + $suffix + $DestExt
        }
    }

    EnsureOutputFolder $to

    Write-Verbose ("Moving file '$from' to '$to'")
    Move-Item -LiteralPath $from -Destination $to -Force
}

###############################################################################

function SmartRename($oldFilenameWithPath, $newFilename)
{
    $path = [System.IO.Path]::GetDirectoryName($oldFilenameWithPath)
    $oldName = [System.IO.Path]::GetFileNameWithoutExtension($oldFilenameWithPath)
    $newName = [System.IO.Path]::GetFileNameWithoutExtension($newFilename)
    $ext = [System.IO.Path]::GetExtension($newFilename)

    $n = 0
    $to = ($newName + $ext)
    While (Test-Path -LiteralPath ($path + "\" + $to))
    {
        $n += 1
        $suffix = $existingNameSuffix -replace '#',$n
        $to = ($newName + $suffix + $ext)
    }
    Write-Verbose "Renaming '$oldName' >>> '$to'"
    Rename-Item -LiteralPath $oldFilenameWithPath -NewName $to

    return ($path + "\" + $to)
}

###############################################################################

function DeleteOriginalFiles($files)
{
    Write-Host "Deleting original files..." -ForegroundColor Yellow

    for($i=0; $i -lt $files.Count; $i++)
    {
        $img = $files[$i]
        if(!$img.Error) # already moved to error folders
        {
            if(Test-Path -LiteralPath ($img.Source))
            {
                Remove-Item -Path $img.Source
                Write-Verbose ("Deleting file: " + $img.Source)
            }
        }
    }
}

###############################################################################

function ArchiveOriginalFiles($files)
{
    Write-Host "Archiving original files..." -ForegroundColor Yellow

    for($i=0; $i -lt $files.Count; $i++)
    {
        $img = $files[$i]
        if(!$img.Error) # already moved to error folders
        {
            if(Test-Path -LiteralPath ($img.Source))
            {
                if($renameArchivedFiles) {
                    $ext = $ext = [System.IO.Path]::GetExtension($img.Source)
                    MoveFile $img.Source $outputFolders.archive ($img.DestFilename + $ext)
                } else {
                    MoveFile $img.Source $outputFolders.archive
                }
            }
        }
    }
}

###############################################################################

function RunExifTool($files)
{
    $timers.exiftool.Start()

    [string[]]$args = @()
    $count = 0

    $commands = ($exifToolCmd -join "`n")

    $addedToCurrentBatch = 0
    $batchesRun = 1
    if($exifToolBatchSize -eq 0) { $exifToolBatchSize = $files.Count }
    $batchesNeeded = [Math]::Ceiling($files.Count / $exifToolBatchSize)

    # create exiftool args file script
    for($i=0; $i -lt $files.Count; $i++)
    {
        $img = $files[$i]
        #if(!$img.Error -and ($img.DestExt -ne ".gif"))
        if(!$img.Error)
        {
            # take current work filename as input
            $in = $img.WorkFile
            $img.OldWorkFile = $img.WorkFile

            # generate a new work filename as output
            $img.WorkFile = $img.DestPath + "\" + ('#' + [System.IO.Path]::GetFileNameWithoutExtension([System.IO.Path]::GetRandomFileName()) + $img.DestExt)
            $out = $img.WorkFile
            Write-Verbose ("Adding to ExifTool script: " + $img.Source + " (" + $img.OldWorkFile + " >> " + $img.OldWorkFile + ")")

            $tmp = ($commands -replace "@input", $in)
            $tmp = ($tmp -replace "@output", $out)
            $args += $tmp
            $args += "" #must end each command in argsfile with empty line

            $count++
            $addedToCurrentBatch++

            $img.UpdateProps = $true
        }

        if( ($addedToCurrentBatch -ge $exifToolBatchSize)  -or  ($i -eq ($files.Count-1)) ) {
            #save the args file
            $args -join "`n" | Out-File $exifToolArgFile -Encoding utf8 -Force

            #launch exif tool
            Write-Host ("`tRunning ExifTool... (batch $batchesRun of $batchesNeeded)")
            $exifToolResult = ExecuteCommand -commandPath $exifToolExe -commandArguments "-@ $exifToolArgFile"

            # check exiftool result
            if($exifToolResult.ExitCode -ne 0)
            {
                if($exifToolResult.stderr -eq "") {
                    Write-Host ("`tExifTool failed to start or unknown error") -BackgroundColor Red
                } else {
                    $errors = ($exifToolResult.stderr -split "`n")
                    Write-Host ("`tExifTool reported " + ($errors.Count-1) + " errors!") -BackgroundColor Red
                    Write-Verbose ("ExifTool stderr: " + $exifToolResult.stderr)
                    # parse the error to add the error msg to the file
                    for($e=0; $e -lt $files.Count; $e++)
                    {
                        foreach($err in $errors)
                        {
                            $source = $files[$e].OldWorkFile -replace "\\","/"
                            if($err.Contains($source))
                            {
                                $files[$e].Error = $true
                                $files[$e].ErrorMsg = $err.Trim()
                                $files[$e].WorkFile = $files[$e].OldWorkFile  # revert back
                                if($printIndividualErrors) {
                                    Write-Host ("`t" + $err.Trim()) -BackgroundColor Red
                                }
                            }
                        }
                    }
                }
            }

            $batchesRun++
            $addedToCurrentBatch = 0
            $args = @()
        }

    }

    Write-Verbose ("ExifTool finished! Cleaning up...")

    Remove-Item -Path $exifToolArgFile

    # remove old workfile if not original
    for($i=0; $i -lt $files.Count; $i++)
    {
        $img = $files[$i]
        if(!$img.Error)
        {
            if($img.OldWorkFile -ne $img.Source)
            {
                Write-Verbose ("Removing old WorkFile: " + $img.OldWorkFile)
                Remove-Item $img.OldWorkFile
                $img.OldWorkFile = ""
            }
        }
    }

    $timers.exiftool.Stop()
}

###############################################################################

function ConvertToJpeg($files)
{
    # assumes that it is run first and thus current workfile is original and should thus not be removed afterwards

    $timers.imconvert.Start()

    Write-Host ("`tConverting images to JPEG...")

    $count = 0
    $errorCount = 0

    $convert = @()
    foreach($f in $files)
    {
        $ext = [System.IO.Path]::GetExtension($f.Source).ToLower()
        $jpegExtensions = ".jpg", ".jpeg", ".jpe"
        if(($ext -notin $jpegExtensions))
        {
            $convert += $f
        }
        elseif($f.SourceJpegQuality -gt $highQualityThreshold)
        {
            $convert += $f
        }
    }

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    Write-Progress -Activity "Converting to JPEG..." -Status "0% done" -PercentComplete 0
    for($i=0; $i -lt $convert.Count; $i++)
    {

        $img = $convert[$i]
        #if(!$img.Error -and ($img.DestExt -ne ".gif"))
        if(!$img.Error)
        {
            if ($sw.Elapsed.TotalMilliseconds -ge 500)
            {
                $percentDone = [System.Math]::Round((($i / $convert.Count) * 100), 0);
                Write-Progress -Activity "Converting to JPEG..." -Status "$percentDone% done" -PercentComplete $percentDone -CurrentOperation $img.Source
                $sw.Reset(); $sw.Start()
            }


            # start image magick
            Write-Verbose ("Converting: " + $img.WorkFile)
            $in = $img.WorkFile
            $img.OldWorkFile = $img.WorkFile
            $img.UpdateProps = $true
            $img.UpdateAdvancedJpgProps = $true;

            # generate a new work filename as output
            $img.DestExt = ".jpg"
            $img.WorkFile = $img.DestPath + "\" + ('#' + [System.IO.Path]::GetFileNameWithoutExtension([System.IO.Path]::GetRandomFileName()) + $img.DestExt)
            $out = $img.WorkFile

            $origDate = Get-ItemProperty -LiteralPath $in | select LastWriteTime

            $in += "[0]" # tell ImageMagick to process first image only (if there are multiple frames/layers)

            $imageMagickResult = ExecuteCommand -commandPath $imageMagickConvertExe -commandArguments ($imageMagickConvertArgsPrefix + '"' + $in + '" ' + $imageMagickConvertArgs + ' "' + $out + '"')
            if($imageMagickResult.ExitCode -eq 0)
            {
                $count++
                Set-ItemProperty -LiteralPath $out -Name LastWriteTime -Value $origDate.LastWriteTime
            }
            else
            {
                $errorCount++
                $errors = ($imageMagickResult.stderr -split "`n")
                $e = ((($errors[0] -replace "convert.exe: ", "") -replace "mogrify.exe: ","") -replace "`r","")
                $img.Error = $true
                $img.ErrorMsg = $e
                if($printIndividualErrors) {
                    Write-Host ("`tImageMagick error: " + $e) -BackgroundColor Red
                }
                Write-Verbose ("ImageMagick stderr:" + $imageMagickResult.stderr)
                $img.WorkFile = $img.OldWorkFile  # revert back
            }
        }
    }

    Write-Progress -Activity "Converting to JPEG..." -PercentComplete 100 -Completed
    Write-Host ("`t`t" + $count + " images converted")
    if($errorCount -gt 0) {
        Write-Host ("`t`t" + $errorCount + " images returned error") -BackgroundColor Red
    }

    $timers.imconvert.Stop()
}

###############################################################################

function PostProcessing($files)
{
    $timers.improcess.Start()

    Write-Host ("`tProcessing Large/Heavy images (Shrink/Re-Save)...")

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    Write-Progress -Activity "Processing Large/Heavy images (Shrink/Re-Save)..." -Status "Preparing" -PercentComplete 0

    # pre-calculate how many files needed to be processed
    $processed = 0
    $count = 0
    for($i=0; $i -lt $files.Count; $i++)
    {
        $img = $files[$i]
        #if(!$img.Error -and ($img.DestExt -ne ".gif"))
        if((!$img.Error) -and ($img.DestPath -eq $outputFolders.jpg))
        {
            if($img.DestWidth -gt $maxWidth -or $img.DestHeight -gt $maxHeight)
            {
                $count++
            }
            elseif($img.DestFileSize -gt $maxWeight)
            {
                $count++
            }
        }
    }

    Write-Progress -Activity "Processing Large/Heavy images (Shrink/Re-Save)..." -Status "0% done" -PercentComplete 0

    for($i=0; $i -lt $files.Count; $i++)
    {
        $img = $files[$i]
        #if(!$img.Error -and ($img.DestExt -ne ".gif"))
        if((!$img.Error) -and ($img.DestPath -eq $outputFolders.jpg))
        {
            if ($sw.Elapsed.TotalMilliseconds -ge 500)
            {
                $percentDone = [System.Math]::Round((($processed / $count) * 100), 0);
                #Write-Progress -Activity "Processing images..." -Status "$percentDone% done" -PercentComplete $percentDone -CurrentOperation $img.Source
                Write-Progress -Activity "Processing Large/Heavy images (Shrink/Re-Save)..." -Status "$percentDone% done" -PercentComplete $percentDone -CurrentOperation $img.Source
                $sw.Reset(); $sw.Start()
            }

            $in = ""

            if($img.DestWidth -gt $maxWidth -or $img.DestHeight -gt $maxHeight)
            {
                # resize to max size
                Write-Verbose ("Image dimensions too large; shrinking: " + $img.Source)
                $in = $img.WorkFile
            }
            elseif($img.DestFileSize -gt $maxWeight)
            {
                # resave to try to shrink size
                # note the elseif --> this will only run if file was not scaled down
                Write-Verbose ("Image filesize too large; re-saving: " + $img.Source)
                $in = $img.WorkFile
            }

            # start image magick
            if($in -ne "")
            {
                # generate a new work filename as output
                $img.OldWorkFile = $img.WorkFile
                $img.WorkFile = $img.DestPath + "\" + ('#' + [System.IO.Path]::GetFileNameWithoutExtension([System.IO.Path]::GetRandomFileName()) + $img.DestExt)
                $out = $img.WorkFile

                $img.UpdateProps = $true
                $img.UpdateAdvancedJpgProps = $true;

                #$imageMagickResult = ExecuteCommand -commandPath $imageMagickMogrifyExe -commandArguments ($args + ' "' + $out + '"')
                $imageMagickResult = ExecuteCommand -commandPath $imageMagickConvertExe -commandArguments ($imageMagickConvertArgsPrefix + '"' + $in + '" ' + $imageMagickConvertArgs + ' "' + $out + '"')
                if($imageMagickResult.ExitCode -eq 0)
                {
                }
                else
                {
                    $errors = ($imageMagickResult.stderr -split "`n")
                    $e = (($errors[0] -replace "convert.exe: ", "") -replace "mogrify.exe: ","")
                    $img.Error = $true
                    $img.ErrorMsg = $e
                    if($printIndividualErrors) {
                        Write-Host ("`tImageMagick error: " + $e) -BackgroundColor Red
                    }
                    Write-Verbose ("ImageMagick stderr:" + $imageMagickResult.stderr)
                    $img.WorkFile = $img.OldWorkFile  # revert back
                }
                $processed++
            }

        }
    }

    # remove old workfile if not original
    for($i=0; $i -lt $files.Count; $i++)
    {
        $img = $files[$i]
        if(!$img.Error)
        {
            if($img.OldWorkFile -ne $img.Source)
            {
                Write-Verbose ("Removing old WorkFile: " + $img.OldWorkFile)
                Remove-Item $img.OldWorkFile
                $img.OldWorkFile = ""
            }
        }
    }


    Write-Progress -Activity "Processing Large/Heavy images (Shrink/Re-Save)..." -PercentComplete 100 -Completed
    Write-Host ("`t`t" + $count + " files processed...")

    $timers.improcess.Stop()
}

###############################################################################

function SetDestinations($files)
{
    $timers.fileio.Start()

    Write-Host ("`tSetting destinations...")

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    Write-Progress -Activity "Setting destinations..." -Status "0% done" -PercentComplete 0
    for($i=0; $i -lt $files.Count; $i++)
    {
        $img = $files[$i]
        if(!$img.Error)
        {
            if ($sw.Elapsed.TotalMilliseconds -ge 500)
            {
                $percentDone = [System.Math]::Round((($i / $files.Count) * 100), 0);
                Write-Progress -Activity "Setting destinations..." -Status "$percentDone% done" -PercentComplete $percentDone -CurrentOperation $img.Source
                $sw.Reset(); $sw.Start()
            }

            if($overwriteOriginal) {
                $img.DestPath = [System.IO.Path]::GetDirectoryName($img.Source)
                $img.DestFilename = [System.IO.Path]::GetFilenameWithoutExtension($img.Source)
            } else {
                # set destination path
                #if($img.DestExt -eq ".gif")
                if([System.IO.Path]::GetExtension($img.source).ToLower() -eq ".gif")
                {
                    if($img.SourceFrameCount -gt 1) {
                        $img.DestPath = $outputFolders.gifanim
                    } else {
                        $img.DestPath = $outputFolders.gif
                    }
                }
                else
                {
                    # Order here is important for business logic

                    if(($outputFolders.not_srgb -ne "") -and ($img.SourceJpegColorSpace -ne "N/A") -and ($img.SourceJpegColorSpace -ne "sRGB"))
                    {
                        Write-Verbose ("JPEG color space is not sRGB: " + $img.Source)
                        $img.DestPath = $outputFolders.not_srgb
                    }
                    elseif(($outputFolders.icc_not_rgb -ne "") -and ($img.SourceIccProfile -ne "N/A") -and ($nonRgbColorProfiles -contains $img.SourceIccProfile))
                    {
                        Write-Verbose ("ICC Color Profile is not RGB: " + $img.Source)
                        $img.DestPath = $outputFolders.icc_not_rgb
                    }
                    elseif(($outputFolders.icc_not_srgb -ne "") -and ($img.SourceIccProfile -ne "N/A") -and ($img.SourceIccProfile -notmatch "sRGB"))
                    {
                        Write-Verbose ("ICC Color Profile is not sRGB: " + $img.Source)
                        $img.DestPath = $outputFolders.icc_not_srgb
                    }
                    elseif(($outputFolders.special_jpg -ne "") -and ($img.SourceJpegMode -ne "") -and ($img.SourceJpegMode -ne "Baseline"))
                    {
                        Write-Verbose ("Special JPPEG Mode: " + $img.SourceJpegMode)
                        $img.DestPath = $outputFolders.special_jpg
                    }
                    elseif(($outputFolders.low_quality -ne "") -and ($img.SourceJpegQuality -ne "N/A") -and ($img.SourceJpegQuality -lt $lowQualityThreshold))
                    {
                        Write-Verbose ("JPEG is low quality: " + $img.Source)
                        $img.DestPath = $outputFolders.low_quality
                    }
                    elseif(($outputFolders.not_8bit -ne "") -and ($img.SourceJpegColorBitDepth -ne "N/A") -and ($img.SourceJpegColorBitDepth -ne "8-bit"))
                    {
                        Write-Verbose ("JPEG color bit depth not 8-bit: " + $img.Source)
                        $img.DestPath = $outputFolders.not_8bit
                    }
                    elseif(($outputFolders.interlace -ne "") -and $img.SourceJpegIsInterlace -eq "JPEG")
                    {
                        Write-Verbose ("JPEG is Interlaced: " + $img.Source)
                        $img.DestPath = $outputFolders.interlace
                    }
                    elseif(($outputFolders.light -ne "") -and ($img.DestFileSize -lt $smallWeight))
                    {
                        Write-Verbose ("Light image: " + $img.Source)
                        $img.DestPath = $outputFolders.light
                    }
                    elseif(($outputFolders.small -ne "") -and ($img.DestWidth -lt $smallWidth -or $img.DestHeight -lt $smallHeight))
                    {
                        Write-Verbose ("Small image: " + $img.Source)
                        $img.DestPath = $outputFolders.small
                    }
                    elseif(($outputFolders.heavy -ne "") -and ($img.DestFileSize -gt $maxWeight))
                    {
                        Write-Verbose ("Heavy image: " + $img.Source)
                        $img.DestPath = $outputFolders.heavy
                    }
                    else
                    {
                        $img.DestPath = $outputFolders.jpg
                    }
                    

                    # always use .jpg as extension
                    $img.DestExt = ".jpg"
                }

                # clean up filenames (rename)
                if($renameOutputFiles) {
                    foreach ($key in $renames.Keys) {
                        $img.DestFilename = ($img.DestFilename -replace $key, $renames.$key)
                    }
                    if($img.DestFilename -eq '')
                    {
                        $img.DestFilename = $noName
                    }
                }

            }
        } else
        {
            $img.DestPath = $outputFolders.error
        }
    }

    Write-Progress -Activity "Setting destinations..." -PercentComplete 100 -Completed

    $timers.fileio.Stop()
}

###############################################################################

function SendToDestination($files)
{
    $timers.fileio.Start()

    Write-Host ("`tMove/Rename temp files to destination filenames...")

    $count = 0

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    Write-Progress -Activity "Moving files to destination..." -Status "0% done" -PercentComplete 0
    for($i=0; $i -lt $files.Count; $i++)
    {
        $img = $files[$i]

        if ($sw.Elapsed.TotalMilliseconds -ge 500)
        {
            $percentDone = [System.Math]::Round((($i / $files.Count) * 100), 0);
            Write-Progress -Activity "Moving files to destination..." -Status "$percentDone% done" -PercentComplete $percentDone
            $sw.Reset(); $sw.Start()
        }
        $count++

        #### rename temp file to real output filename

        $from = $img.WorkFile
        if($img.Error)
        {
            $from = $img.Source
            if(($img.Source -ne $img.WorkFile) -and (Test-Path -LiteralPath ($img.WorkFile)))
            {
                Write-Verbose ("Removing WorkFile of file that had error: " + $img.WorkFile)
                Remove-Item $img.WorkFile
            }
        }


        if($img.Source -ne "") {
            $to = $img.DestPath + "\" + $img.DestFilename + $img.DestExt

            EnsureOutputFolder $to

            # first check existing filenames
            if(!$overwriteOriginal) {
                if($overwriteExisting -eq $false) {
                    $n = 0
                    $baseName = $img.DestFilename
                    While (Test-Path -LiteralPath ($to))
                    {
                        $n += 1
                        $suffix = $existingNameSuffix -replace '#',$n
                        $img.DestFilename = $baseName + $suffix
                        $to = $img.DestPath + "\" + $img.DestFilename + $img.DestExt
                    }
                }
                Write-Verbose ("Move/Rename file '$from' to '$to'")
            } else {
                Write-Verbose ("Replace original with work file: '$from' >> '$to'")
                if($to.ToLower() -ne $img.Source.ToLower()) {
                    Write-Verbose ("Removing original file with other name: " + $img.Source)
                    Remove-Item -LiteralPath $img.Source -Force
                }
            }

            Move-Item -LiteralPath $from -Destination $to -Force

            $img.WorkFile = $to
        }
    }

    Write-Progress -Activity "Moving files to destination..." -PercentComplete 100 -Completed

    Write-Host ("`t`t" + $count + " files moved/renamed")

    $timers.fileio.Stop()
}

###############################################################################

function CheckExtensionsOnErrors($files)
{
    Write-Host ("Validating file extensions...") -ForegroundColor Yellow
    $extensionsCorrected = 0

    $tot = 0
    for($i=0; $i -lt $files.Count; $i++)
    {
        if($files[$i].Error)
        {
            $tot++
        }
    }

    $count = 0
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    Write-Progress -Activity "Checking file extension on erroneous files..." -Status "0% done" -PercentComplete 0
    for($i=0; $i -lt $files.Count; $i++)
    {
        $img = $files[$i]
        if($img.Error)
        {
            $count++
            if ($sw.Elapsed.TotalMilliseconds -ge 500)
            {
                $percentDone = [System.Math]::Round((($count / $tot) * 100), 0);
                Write-Progress -Activity "Checking file extension on erroneous files..." -Status "$percentDone% done" -PercentComplete $percentDone -CurrentOperation $img.Source
                $sw.Reset(); $sw.Start()
            }

            $fn = $img.DestPath + "\" + $img.DestFilename + $img.DestExt

            $e = ExecuteCommand $tridExe ('-ce "' + $fn +'"')
            if($e.ExitCode -eq 0) {
                if($e.stdout -like "*1 file(s)*") {

                    $oldNameWithExt = [System.IO.Path]::GetFileName($fn)
                    $oldName = [System.IO.Path]::GetFileNameWithoutExtension($fn)
                    $oldPath = [System.IO.Path]::GetDirectoryName($fn)
                    $fn = (dir -LiteralPath $oldPath -Filter "$oldName.*")[0].FullName
                    $img.DestExt = [System.IO.Path]::GetExtension($fn)
                    $img.DestPath = $outputFolders.wrongext
                    Write-Verbose ("File extension corrected: " + $oldNameWithExt + " >> " + [System.IO.Path]::GetFileName($fn))
                    MoveFile $fn $outputFolders.wrongext
                    $extensionsCorrected++
                }
            } else {
                Write-host ("Error reading file: " + $fn) -ForegroundColor White -BackgroundColor Red
            }

        }
    }

    Write-Progress -Activity "Checking file extension on erroneous files..." -PercentComplete 100 -Completed
    Write-Host ("`t" + $extensionsCorrected + " files corrected and moved")
}

###############################################################################

function CheckExtensionsBefore($files)
{
    Write-Host ("`tValidating file extensions...")
    $extensionsCorrected = 0

    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    Write-Progress -Activity "Identifying file formats..." -Status "0% done" -PercentComplete 0
    for($i=0; $i -lt $files.Count; $i++)
    {
        $img = $files[$i]
        if(!$img.Error)
        {
            if ($sw.Elapsed.TotalMilliseconds -ge 500)
            {
                $percentDone = [System.Math]::Round((($i / $files.Count) * 100), 0);
                Write-Progress -Activity "Identifying file formats..." -Status "$percentDone% done" -PercentComplete $percentDone -CurrentOperation $img.Source
                $sw.Reset(); $sw.Start()
            }

            $e = ExecuteCommand $tridExe ('-ce "' + $img.WorkFile +'"')
            if($e.ExitCode -eq 0) {
                if($e.stdout -like "*1 file(s)*") {

                    $oldNameWithExt = [System.IO.Path]::GetFileName($img.WorkFile)
                    $oldName = [System.IO.Path]::GetFileNameWithoutExtension($img.WorkFile)
                    $oldPath = [System.IO.Path]::GetDirectoryName($img.WorkFile)
                    $img.WorkFile = (dir -LiteralPath $oldPath -Filter "$oldName.*")[0].FullName
                    $img.DestExt = [System.IO.Path]::GetExtension($img.WorkFile)
                    if($img.Source -eq ($oldPath + "\" + $oldNameWithExt)) {
                        $img.Source = $img.WorkFile
                    }
                    $extensionsCorrected++
                    Write-Verbose ("File extensions corrected: " + $oldNameWithExt + " >> " + [System.IO.Path]::GetFileName($img.WorkFile))

                    if($validExtensions -notcontains $img.DestExt) {
                        Write-Verbose ("`tNot a valid file extension anymore. Skipping further processing...")
                        if($moveNonImageFiles) {
                            MoveFile $img.Source $outputFolders.nonimg
                        }
                        $img.Error = $true
                        $img.ErrorMsg = "Not a valid file extension anymore..."
                        $img.Source = ""
                    }

                }
            } else {
                Write-host ("Error reading file: " + $img.Source) -ForegroundColor White -BackgroundColor Red
            }

        }
    }

    Write-Progress -Activity "Identifying file format..." -PercentComplete 100 -Completed
    Write-Host ("`t`t" + $extensionsCorrected + " file extensions corrected")
}

###############################################################################

function FairlySmartCharReplace {
    param (
        [String]$src = [String]::Empty
    )

    #replace diacritics
    $normalized = $src.Normalize( [Text.NormalizationForm]::FormD )
    $sb = new-object Text.StringBuilder
    $normalized.ToCharArray() | % {
    if( [Globalization.CharUnicodeInfo]::GetUnicodeCategory($_) -ne [Globalization.UnicodeCategory]::NonSpacingMark) {
    [void]$sb.Append($_)
    }
    }
    $sb=$sb.ToString()
    
    #replace via code page conversion
    $NonUnicodeEncoding = [System.Text.Encoding]::GetEncoding(850)
    $UnicodeEncoding = [System.Text.Encoding]::Unicode
    [Byte[]] $UnicodeBytes = $UnicodeEncoding.GetBytes($sb);
    [Byte[]] $NonUnicodeBytes = [System.Text.Encoding]::Convert($UnicodeEncoding, $NonUnicodeEncoding , $UnicodeBytes);
    [Char[]] $NonUnicodeChars = New-Object -TypeName “Char[]” -ArgumentList $($NonUnicodeEncoding.GetCharCount($NonUnicodeBytes, 0, $NonUnicodeBytes.Length))
    [void] $NonUnicodeEncoding.GetChars($NonUnicodeBytes, 0, $NonUnicodeBytes.Length, $NonUnicodeChars, 0);
    [String] $NonUnicodeString = New-Object String(,$NonUnicodeChars)

    ($NonUnicodeString -replace "[^a-zA-Z0-9_().\- ]","#")

    #$NonUnicodeString
}

###############################################################################


function IdentifyOddsetJpegProperties($files)
{
    if($readAdvancedJpgProps -eq $false) {
        return;
    }

    $timers.advprops.Start()

    Write-Host ("`tReading advanced jpg properties...")

    $count = 0
    $errorCount = 0

    $convert = @()
    foreach($f in $files)
    {
        $ext = [System.IO.Path]::GetExtension($f.WorkFile).ToLower();
        $jpegExtensions = ".jpg", ".jpeg", ".jpe"
        if($ext -in $jpegExtensions)
        {
            $convert += $f
        }
    }


    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    Write-Progress -Activity "Reading advanced jpg properties..." -Status "0% done" -PercentComplete 0
    for($i=0; $i -lt $convert.Count; $i++)
    {

        $img = $convert[$i]
        if(!$img.Error -and $img.UpdateAdvancedJpgProps)
        {
            if ($sw.Elapsed.TotalMilliseconds -ge 500)
            {
                $percentDone = [System.Math]::Round((($i / $convert.Count) * 100), 0);
                Write-Progress -Activity "Reading advanced jpg properties..." -Status "$percentDone% done" -PercentComplete $percentDone -CurrentOperation $img.Source
                $sw.Reset(); $sw.Start()
            }


            Write-Verbose ("Reading " + $img.WorkFile)
            $in = $img.WorkFile
            $img.UpdateAdvancedJpgProps = $false;

            # Use ImageMagick Identify
            $imageMagickResult = ExecuteCommand -commandPath $imageMagickIdentifyExe -commandArguments ('-verbose "' + $in + '"')
            if($imageMagickResult.ExitCode -eq 0)
            {
                $count++
                $output = $imageMagickResult.stdout -split [System.Environment]::NewLine

                for($n = 0 ; $n -lt $output.Length; $n++)
                {
                    $line = $output[$n] -split ": "
                    $prop = $line[0].trim();

                    if($prop -eq "Colorspace") # ex: Colorspace: sRGB
                    {
                        $img.SourceJpegColorSpace = $line[1]
                    }
                    elseif($prop -eq "Depth") # ex: Depth: 8-bit
                    {
                        $img.SourceJpegColorBitDepth = $line[1]
                    }
                    elseif($prop -eq "Quality") # ex: Quality: 92
                    {
                        $img.SourceJpegQuality = $line[1].trim()
                    }
                    elseif($prop -eq "Interlace") # ex: Interlace: JPEG
                    {
                        $img.SourceJpegIsInterlace = $line[1].trim()
                    }
                    elseif($prop -eq "icc:model") # ex: icc:model: ColorMatch RGB
                    {
                        $img.SourceIccProfile = $line[1].trim()
                    }
                }

            }
            else
            {
                $errorCount++
                $errors = ($imageMagickResult.stderr -split "`n")
                $e = (($errors[0] -replace "identify.exe: ", "") -replace "`r","")
                $img.Error = $true
                $img.ErrorMsg = $e
                if($printIndividualErrors) {
                    Write-Host ("`tImageMagick error: " + $e) -BackgroundColor Red
                }
                Write-Verbose ("ImageMagick stderr:" + $imageMagickResult.stderr)
            }


            #file open for the analysis by per byte
            $fs = [IO.File]::OpenRead($in)
            for($j = 0 ; $j -lt $fs.Length; $j++)
            {
                $fs.Position = $j
                $byte = $fs.ReadByte()
                if($byte -eq 255)      #if it is marker? 255 means 0xff, Oxff means marker
                {
                    $j++              #forward byte
                    $byte = $fs.ReadByte()
                    $byte = [Convert]::ToString($byte, 16)  #transformation to string

                    if($byte -eq "da") #da means SOS (Start of Scan)
                    {
                        break
                    }
                    if($byte -eq "c0")  #c0 means Baseline jpeg
                    {
                        $img.SourceJpegMode = "Baseline"
                        break;
                    }
                    if($byte -eq "c2" -or $byte -eq "c6" -or $byte -eq "ca" -or $byte -eq "ce")
                    {
                        $img.SourceJpegMode = "Progressive"
                        break;
                    }  
                    if($byte -eq "c9" -or $byte -eq "ca" -or $byte -eq "cb" -or $byte -eq "cd" -or $byte -eq "ce" -or $byte -eq "cf")
                    {
                        $img.SourceJpegMode = "Arithmetic"
                        break;
                    }
                    if($byte -eq "c3" -or $byte -eq "c7" -or $byte -eq "cb" -or $byte -eq "cf")
                    {
                        $img.SourceJpegMode = "Lossless"
                        break;
                    }  
                    if($byte -eq "c5" -or $byte -eq "c6" -or $byte -eq "c7" -or $byte -eq "cd" -or $byte -eq "ce" -or $byte -eq "cf")
                    {
                        $img.SourceJpegMode = "Hierarchical"
                        break;
                    }
                }
            }
            $fs.Close()

        }
    }

    Write-Progress -Activity "Reading advanced jpg properties..." -PercentComplete 100 -Completed
    Write-Host ("`t`t" + $count + " image files read")
    if($errorCount -gt 0) {
        Write-Host ("`t`t" + $errorCount + " image files returned error") -BackgroundColor Red
    }

    $timers.advprops.Stop()
}


###############################################################################

function Start-WorkflowGui()
{
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Windows.Forms")
    [void] [System.Reflection.Assembly]::LoadWithPartialName("System.Drawing")


    ### Create form ###

    $form = New-Object System.Windows.Forms.Form
    $form.Text = "PowerShell GUI"
    $form.Size = '260,320'
    $form.StartPosition = "CenterScreen"
    $form.MinimumSize = $form.Size
    $form.MaximizeBox = $False
    $form.Topmost = $True


    ### Define controls ###

    $button = New-Object System.Windows.Forms.Button
    $button.Location = '5,5'
    $button.Size = '75,23'
    $button.Width = 80
    $button.Text = "Start"

    $checkbox = New-Object Windows.Forms.Checkbox
    $checkbox.Location = '100,8'
    $checkbox.AutoSize = $True
    $checkbox.Text = "Show log afterwards"

    $label = New-Object Windows.Forms.Label
    $label.Location = '5,40'
    $label.AutoSize = $True
    $label.Text = "Drop files or folders here:"

    $listBox = New-Object Windows.Forms.ListBox
    $listBox.Location = '5,60'
    $listBox.Height = 200
    $listBox.Width = 240
    $listBox.Anchor = ([System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right -bor [System.Windows.Forms.AnchorStyles]::Top)
    $listBox.IntegralHeight = $False
    $listBox.AllowDrop = $True

    $statusBar = New-Object System.Windows.Forms.StatusBar
    $statusBar.Text = "Ready"


    ### Add controls to form ###

    $form.SuspendLayout()
    $form.Controls.Add($button)
    $form.Controls.Add($checkbox)
    $form.Controls.Add($label)
    $form.Controls.Add($listBox)
    $form.Controls.Add($statusBar)
    $form.ResumeLayout()


    ### Write event handlers ###

    $button_Click = {

        # collect files from listbox
        $files = @()
	    foreach ($item in $listBox.Items)
        {
            $i = Get-Item -LiteralPath $item
            if($i -is [System.IO.DirectoryInfo]) {
                dir -Path $i -Recurse -File -Filter *.* | foreach {
                    $files += $_
                }
            } else {
                $files += $i
            }
	    }

        $statusBar.Text = "Running workflow..."
        if($checkbox.Checked) {
            Start-Workflow -ImageFiles $files | Out-GridView
        } else {
            Start-Workflow -ImageFiles $files | Out-Null
        }
        $listBox.Items.Clear()
        $statusBar.Text = "Ready!"

    }

    $listBox_DragOver = [System.Windows.Forms.DragEventHandler]{
	    if ($_.Data.GetDataPresent([Windows.Forms.DataFormats]::FileDrop)) # $_ = [System.Windows.Forms.DragEventArgs]
	    {
	        $_.Effect = 'Copy'
	    }
	    else
	    {
	        $_.Effect = 'None'
	    }
    }
	
    $listBox_DragDrop = [System.Windows.Forms.DragEventHandler]{
	    foreach ($filename in $_.Data.GetData([Windows.Forms.DataFormats]::FileDrop)) # $_ = [System.Windows.Forms.DragEventArgs]
        {
		    $listBox.Items.Add($filename)
	    }
        $statusBar.Text = ("List contains $($listBox.Items.Count) items")
    }

    $form_FormClosed = {
	    try
        {
            $listBox.remove_Click($button_Click)
		    $listBox.remove_DragOver($listBox_DragOver)
		    $listBox.remove_DragDrop($listBox_DragDrop)
            $listBox.remove_DragDrop($listBox_DragDrop)
		    $form.remove_FormClosed($Form_Cleanup_FormClosed)
	    }
	    catch [Exception]
        { }
    }


    ### Wire up events ###

    $button.Add_Click($button_Click)
    $listBox.Add_DragOver($listBox_DragOver)
    $listBox.Add_DragDrop($listBox_DragDrop)
    $form.Add_FormClosed($form_FormClosed)


    #### Show form ###

    [void] $form.ShowDialog()
}

###############################################################################

cls


### Place your command line here (see examples below how to write) ############





<#

How to run the script:

    Start-Workflow is the starting point.
    It accepts either an array of PowerShell file objects, or a string array or filenames
    You can use the "dir" command to list file and pass to the workflow
    It will print to the console what it is doing
        You can get it to print more details by using -Verbose 
    When finished, it will output the log to the pipe.
        You can use Out-GridView to display
    Also, this is saved to a log file, as well as saved in the global variable $workflowResult
    
    The easiest way to start to the workflow is to write a line where you pipe files to 
    the workflow and place this command at the bottom of the script (ie here), and then 
    simply run the file (F5 in PowerShell ISE).

Tutorial with Examples:

    How to use "dir" command to list the files in a folder:

        dir -Path 'T:\Samples\' -File 
            
    We can add a filter to only display jpg files:

        dir -Path 'T:\Samples\' -File -Filter *.jpg 
            
    Or list all files with *.*:

        dir -Path 'T:\Samples\' -File -Filter *.*
            
    Using -Recuse will get us all files in all subfolders too:

        dir -Path 'T:\Samples\' -Recurse -File -Filter *.*

    Using the pipe character "|" will pass the result of a command (such as dir) to another command.
    This will pass the files to a simple GUI:

        dir -Path 'T:\Samples\' -Recurse -File -Filter *.* | Out-GridView

    Lets pass it to the workflow instead:

        dir -Path 'T:\Samples\' -Recurse -File -Filter *.* | Start-Workflow

    If we want to see more info we can use the -Verbose switch:

        dir -Path 'T:\Samples\' -Recurse -File -Filter *.* | Start-Workflow -Verbose

    Lets display the output in the gridview instead by passing the result of the workflow further down the pipeline:

        dir -Path 'T:\Samples\' -Recurse -File -Filter *.* | Start-Workflow -Verbose | Out-GridView

#>


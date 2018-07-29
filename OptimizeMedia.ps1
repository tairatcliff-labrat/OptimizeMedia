<#
Instructions

Download ffmpeg from 
        http://ffmpeg.zeranoe.com/builds/
        https://ffmpeg.org/download.html

Install ffmpeg and add the /bin folder to the Windows path
        1. Start the System Control Panel applet (Start - Settings - Control Panel - System).
        2. Select the Advanced tab.
        3. Click the Environment Variables button.
        4. Under System Variables, select Path, then click Edit.
        5. You'll see a list of folders with a ";" separator.
        6. Add the ffmpeg /bin folder to the end of the list. e.g ;C:\Program Files\ffmpeg-20171027-5834cba-win64-static\bin
#>


param(
    [ValidateScript({Test-Path $_})][string]$Directory = "C:\Media\Movies", # The directory to recursively scan for media.
    [ValidateRange(0,999)][int]$OptimizeAfterDays = 0, # How old a file must be before it is optimized.
    [string]$OutputToCsv, # Output the "Processed Videos" report to a csv file. Specify the CSV file to save to i.e. "C:\Temp\Videos.csv". If left blank a CSV won't be generated.  
    [string]$LogFile, #If a logfile is required, enter the directory and logfile name. i.e. "C:\Temp\Videos.log"
    [ValidateSet("hd480","hd720","hd1080")][string]$ffmpegQuality = "hd720", 
    [ValidateRange(0,51)][string]$ConstantRateFactor = "21", #A setting between 21 and 24 is good. 0 is lossless and highest quality and 51 is the worst quality. 
    [switch]$ValidateOnly, # Will simulate the script running but will not actually optimize or delete any files. The "Processed Videos" output will still be generated.
    [switch]$DeleteOriginalVideo # Without this parameter in the command the original file will be optimized and the not deleted.
    )

cls
Write-Host `n`n`n`n`n`n

## Declare Variables
$VideoExtensions = "*.mkv","*.mp4","*.avi" # What video file extensions are supported by ffmpeg and you want to be included in the optimisation process. 
$TimeOut = New-TimeSpan -Seconds 30 # When attempting to delete original files how long should the script continue to loop and retry before moving on.
$ScriptRunningTime = [Diagnostics.stopwatch]::StartNew() # Starts a timer to monitor the total time the script takes to run.
$Date = Get-Date # Get the current date so that the script can determine if files are older that the "$OptimizeAfterDays" has been reached.
$EscapeParser = "--%" # When executing the ffmpeg command the escape characters need to be passed as a variable, which allows for all of the parameters to be captured within the command.
$NumFilesCounter = 0
$CurrentlyProcessedSize = 0
$OriginalTotalSize = 0
$OptimizedTotalFileSize = 0
$Percent = 0
$VideosErrored = @()
$VideosProcessed = @()
$Files = @()
$Files = Get-ChildItem -Recurse $Directory -Exclude *.optimized.* -Include $VideoExtensions | Where {($_.Attributes -ne "Directory")}| Sort-Object
$NumFiles = $Files.Count
ForEach($File in $Files){
    $OriginalTotalSize += ($File.Length).ToString()
}

#Creates a log file with time and date logged entries
function Log-Message([string]$Message) {
    If($LogFile){
        #Writes message to log file
        add-content $LogFile "$Date | $Message"
    }
}

ForEach($File in $Files){
    Write-Host "`n###################################### Processing New Video ######################################"
    $NumFilesCounter ++
    $InputFileName = $File.FullName
    $InputFileBaseName = $File.BaseName
    $InputFileDirectory = $File.DirectoryName
    $OutputFileName = [io.path]::ChangeExtension($File.FullName, "optimized.$ffmpegQuality.mp4")
    $OutputFileShortName = $OutputFileName.Split("\")[-1] #$OutputFileName.Substring($OutputFileName.LastIndexOf("\"),100)
    $FileAge = $Date - $File.CreationTime
    Write-Host "`n$($File.BaseName) is $($FileAge.days) days old"-ForegroundColor Yellow
    If ($FileAge.Days -gt $OptimizeAfterDays){
        Write-Host "Optimzing any videos that are older than $OptimizeAfterDays days, The video is old enough to be optimized" -ForegroundColor Yellow
        $InputFileDuration = &ffprobe.exe $EscapeParser " -v error -show_format -select_streams v `"$InputFileName`"" | Select-String "duration="
        $InputFileDurationTime = New-TimeSpan -Seconds (($InputFileDuration.ToString()).Split('=')[1].Split('.')[0])
        $InputFileSize = $File.Length.ToString()
        Write-Host "`nInput video duration is (hh:mm:ss):" $InputFileDurationTime.ToString("hh\:mm\:ss") -ForegroundColor Yellow

        $Props = @{}
        $Props.Add("File Name", $InputFileBaseName)
        $Props.Add("File Directory", $InputFileDirectory)
        $Props.Add("Age (Days)", $FileAge.days)
        $Props.Add("Original Size (MB)", ([math]::Round(($InputFileSize / 1MB),2)))

        $CurrentlyProcessedSize += $InputFileSize

        $OptimizeVideoTimer = [Diagnostics.stopwatch]::StartNew()
        If(!($ValidateOnly)){
            $ffprobeBlock = {
            Param(
                $InputFileName
            )
            ffprobe.exe -v error -select_streams v -of default=noprint_wrappers=1:nokey=1 -show_entries stream=r_frame_rate,nb_frames "$InputFileName" 2>&1
            }
            
            Write-Host "`nProcessing $($File.Name)" -ForegroundColor Green

            $InputFileTotalFrames = Invoke-Command -ScriptBlock $ffprobeBlock -ArgumentList $InputFileName 
            [string]$InputFileNbFrames = $InputFileTotalFrames[1]
            [string]$InputFileFrameRate = $InputFileTotalFrames[0] 

            if($InputFileNbFrames.ToUpper() -ne "N/A"){
                [int]$InputFileTotalFramesInt = $InputFileTotalFrames.Trim().Split('=')[1]
            } else {
                Write-Host "`nTotal frame count unavailable, estimating the total frames based on the frame rate and duration." -ForegroundColor Red
                [int]$FrameRate = $InputFileFrameRate.Split("/")[0] / $InputFileFrameRate.Split("/")[1]
                [int]$InputFileTotalFramesInt = $InputFileDurationTime.TotalSeconds * $FrameRate
            }

            $ffmpegBlock = {
            Param(
                [string]$InputFileName, 
                [string]$ffmpegQuality, 
                [int]$ConstantRateFactor, 
                [string]$OutputFileName
            )
            ffmpeg.exe -y -i "$InputFileName" -s $ffmpegQuality -c:v libx264 -crf $ConstantRateFactor -b:v 200k -b:a 128k -c:a aac -strict -2 "$OutputFileName" 4>&1
            } 
            
            $OptimizeVideo = Start-Job -ScriptBlock $ffmpegBlock -ArgumentList $InputFileName, $ffmpegQuality, $ConstantRateFactor, $OutputFileName
            Write-Progress -Activity "Processed 0% (frame: 0 of $InputFileTotalFramesInt) of Movie: ""$OutputFileShortName""" -PercentComplete 0
            Start-Sleep 15

            While($OptimizeVideo.State -eq "Running"){
                [string]$OptimizeVideoOutput = Receive-Job $OptimizeVideo 2>&1
                [string]$CurrentFrame = $OptimizeVideoOutput.Substring(($OptimizeVideoOutput.LastIndexOf("frame=") + 6), 20)
                [int]$CurrentFrameInt = $CurrentFrame.Substring(0, $CurrentFrame.IndexOf("f"))
                $Percent = [math]::Round($CurrentFrameInt / $InputFileTotalFramesInt * 100, 2)
                If ($Percent -gt 100){
                    Write-Host "Current Frame is higher than expected: " $CurrentFrameInt of $InputFileTotalFramesInt
                    $Percent = 100
                }
                Write-Progress -Activity "Processed $Percent% (frame: $CurrentFrameInt of $InputFileTotalFramesInt) of Movie: ""$OutputFileShortName""" -PercentComplete $Percent
                Start-Sleep 30
            }
            Write-Progress -Activity "Processed $Percent% (frame: $CurrentFrameInt of $InputFileTotalFramesInt) of Movie: ""$OutputFileShortName""" -Completed
            If($JobOutput){Remove-Job $OptimizeVideo}
        }
    } Else {
        Write-Host "`nOptimzing any videos that are older than $OptimizeAfterDays days, The video is not old enough to be optimized" -ForegroundColor Yellow
        Continue
    }

    If (Test-Path $OutputFileName){
        $OutputFileDuration = &ffprobe.exe $EscapeParser " -v error -show_format -select_streams v `"$OutputFileName`"" | Select-String "duration="
        $OutputFileDurationTime = New-TimeSpan -Seconds ($OutputFileDuration.ToString()).Split('=')[1].Split('.')[0]
        [string]$OutputFileTotalFrames = Invoke-Command -ScriptBlock $ffprobeBlock -ArgumentList $OutputFileName
        [string]$OutputFileTotalFrames = $OutputFileTotalFrames.Trim().Split('=')[1]

        Write-Host "Input File size is (MB):" ([math]::Round(($InputFileSize / 1MB).ToString(),2)) -ForegroundColor Yellow
        Write-Host "Output File size is (MB):" ([math]::Round(($OutputFileSize / 1MB).ToString(),2)) -ForegroundColor Yellow
        Write-Host "`nInput video duration (minutes):" $InputFileDurationTime.TotalMinutes -ForegroundColor Yellow
        Write-Host "Output video duration (minutes):" $OutputFileDurationTime.TotalMinutes -ForegroundColor Yellow

        $DurationDifference = [Math]::Abs($OutputFileDurationTime.TotalSeconds - $InputFileDurationTime.TotalSeconds)
        Write-Host "Input file duration (seconds): "$InputFileDurationTime.TotalSeconds
        Write-Host "Output file duration (seconds): "$OutputFileDurationTime.TotalSeconds
        Write-Host "Difference: " $DurationDifference

        If(($DurationDifference -lt 30) -or ($OutputFileDurationTime -gt $InputFileDurationTime)){
            $OutputFileSize = (Get-Item $OutputFileName).Length.ToString()
            $OptimizedTotalFileSize +=  $OutputFileSize
            
            $Props.Add("Optimized Size (MB)", ([math]::Round(($OutputFileSize / 1MB).ToString(),2)))
            
            Write-Host "`nSuccessfully optimized the video" -BackgroundColor Green -ForegroundColor Black
            Write-Host "The optimized file has been saved to:" $OutputFileName -BackgroundColor Green -ForegroundColor Black
            Write-Host "`nAttempting to remove file:" $InputFileName
            
            $StopWatch = [Diagnostics.stopwatch]::StartNew()
            While ((Test-Path $InputFileName) -and ($StopWatch.Elapsed -lt $TimeOut)){
                If($DeleteOriginalVideo){
                    Remove-Item -Path $InputFileName -Force
                } Else {
                    Write-Host "`"Delete Original Video`" is set to False. The original video file has not been deleted." -ForegroundColor Yellow
                    Break
                }
                If (Test-Path $InputFileName){Start-Sleep 5}
            }
            If ($StopWatch.Elapsed -gt $TimeOut){Write-Host "Timed out while trying to delete" $InputFileName -BackgroundColor Red -ForegroundColor Black}
        } Else {
            Write-Host "`nAn inconsistency has been detected in the output file. As a precaution the input file won't be deleted" -BackgroundColor Yellow -ForegroundColor Black
            Write-Host "Deleting the output file and keeping the original input file" -BackgroundColor Yellow -ForegroundColor Black
            
            $VideoErrored = @{}
            $VideoErrored.Add("File Name", $InputFileBaseName)
            $VideoErrored.Add("Input video duration (minutes)", $InputFileDurationTime.TotalMinutes)
            $VideoErrored.Add("Output video duration (minutes)", $OutputFileDurationTime.TotalMinutes)
            $VideosErrored += $VideoErrored

            $StopWatch = [Diagnostics.stopwatch]::StartNew()
            While ((Test-Path $OutputFileName) -and ($StopWatch.Elapsed -lt $TimeOut)){
                Remove-Item -Path $OutputFileName -Force
                If (Test-Path $OutputFileName){
                    Start-Sleep 5
                }Else{
                    Write-Host "`nSuccessfully deleted:" $OutputFileName -BackgroundColor Yellow -ForegroundColor Black
                }
            }
            If ($StopWatch.Elapsed -gt $TimeOut){Write-Host "Timed out while trying to delete" $OutputFileName -BackgroundColor Red -ForegroundColor Black}
        }
    }

    $VideosProcessed += New-Object PSObject -Property $Props

    Write-Host "`nThe time to complete this video was (hh:mm:ss):" $OptimizeVideoTimer.Elapsed.ToString("hh\:mm\:ss") -ForegroundColor Yellow
    Write-Host "`nThe script has been running for a total time of (hh:mm:ss):" $ScriptRunningTime.Elapsed.ToString("hh\:mm\:ss") -ForegroundColor Yellow
    Write-Host "`nProcessed $NumFilesCounter of $NumFiles videos" -ForegroundColor Cyan
    Write-Host "Proccessed" ([math]::Round(($CurrentlyProcessedSize / 1GB),2)) "GB of the total" ([math]::Round(($OriginalTotalSize / 1GB),2)) "GB" -ForegroundColor Cyan
    If(!($ValidateOnly)){
        Write-Host "`nThe $NumFilesCounter processed files have been reduced from" ([math]::Round(($CurrentlyProcessedSize / 1GB),2)) "GB to:" ([math]::Round(($OptimizedTotalFileSize / 1GB),2)) "GB" -ForegroundColor Cyan
    }

}
Write-Host
If($VideosErrored -ne $null){
    Write-Host "`nThe following vidoes failed to be processed" -BackgroundColor Red -ForegroundColor Black
    $VideosErrored | FT
}

Write-Host "`nThe following vidoes have been processed" -BackgroundColor Green -ForegroundColor Black
$VideosProcessed | Select "File Name", "File Directory", "Age (Days)", "Original Size (MB)", "Optimized Size (MB)" | FT
If($OutputToCsv){$VideosProcessed | Export-Csv -Path $OutputToCsv -NoTypeInformation}

If($ErrorLog -ne $null){
    Write-Host "All Errors that were detected are:"
    $ErrorLog
}

Write-Host
Write-Host "After Processing the directory:" $Directory -ForegroundColor Cyan
Write-Host "Processed $NumFilesCounter of $NumFiles videos" -ForegroundColor Cyan
Write-Host "The original size of the directory was:" ([math]::Round(($OriginalTotalSize / 1GB),2)) "GB" -ForegroundColor Cyan
If(!($ValidateOnly)){
    Write-Host "The optimized size is of the directory is:" ([math]::Round(($OptimizedTotalFileSize / 1GB),2)) "GB" -ForegroundColor Cyan
    Write-Host
    $ReducedSize = [math]::Round((($OriginalTotalSize - $OptimizedTotalFileSize) / 1GB),2)
    $ReducedPercentage = [math]::Truncate(($OriginalTotalSize - $OptimizedTotalFileSize) / $OriginalTotalSize * 100)
    Write-Host "Reduced the total size by $ReducedSize GB, or $ReducedPercentage %"-ForegroundColor Cyan
}

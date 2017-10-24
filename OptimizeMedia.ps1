param(
    [string]$Directory = "C:\Temp", # The directory to recursively scan for media.
    [string]$OptimizeAfterDays = -30, # How old a file must be before it is optimized.
    [string]$OutputToCsv, # Output the "Processed Videos" report to a csv file. Specify the CSV file to save to i.e. "C:\Temp\Videos.csv". If left blank a CSV won't be generated.  
    [switch]$ValidateOnly = $False, # Will simulate the script running but will not actually optimize or delete any files. The "Processed Videos" output will still be generated.
    [switch]$DeleteOriginalVideo = $True # Without this parameter in the command the original file will be optimized and the not deleted.
    )

## Declare Variables
$VideoExtensions = "*.mkv","*.mp4","*.avi"
$VideosProcessed = @()
$TimeOut = New-TimeSpan -Seconds 30
$ScriptRunningTime = [Diagnostics.stopwatch]::StartNew()
$Date = Get-Date
$EscapeParser = "--%"
$Files = @()
$Files = Get-ChildItem -Recurse $Directory -Exclude *.optimized.* -Include $VideoExtensions | Where {($_.Attributes -ne "Directory")}
$NumFiles = $Files.Count
$NumFilesCounter = 0
$ProcessedSize = 0
$OriginalTotalSize = 0
$OptimizedTotalSize = 0
$ProcessedOutputFileSize = 0
ForEach($File in $Files){
    $OriginalTotalSize += ($File.Length).ToString()
}


ForEach($File in $Files){
    Write-Host "###################################### Processing New Video ######################################"
    $NumFilesCounter ++
    $InputFileName = $File.FullName
    $OutputFileName = [io.path]::ChangeExtension($File.FullName, "optimized.480p.mp4")
    $FileAge = $Date - $File.CreationTime
    Write-Host "$($File.Basename) is $($FileAge.days) days old"-ForegroundColor Yellow
    Write-Host
    If ($FileAge.Days -gt $OptimizeAfterDays){
        Write-Host "Optimzing any videos that are older than $OptimizeAfterDays days, The video is old enough to be optimized" -ForegroundColor Yellow
        $InputFileDuration = &ffprobe.exe $EscapeParser " -v error -show_format -select_streams v `"$InputFileName`"" | Select-String "duration="
        $InputFileDurationTime = New-TimeSpan -Seconds (($InputFileDuration.ToString()).Split('=')[1].Split('.')[0])
        $InputFileSize = [math]::Round(($File.Length / 1MB).ToString(),2)
        Write-Host
        Write-Host "Input video duration is (hh:mm:ss):" $InputFileDurationTime.ToString("hh\:mm\:ss") -ForegroundColor Yellow
        Write-Host
        #$Props = $null
        $Props = @{}
        $Props.Add("Name", $File.BaseName.ToString())
        $Props.Add("Age (Days)", $FileAge.days)
        $Props.Add("Original Size (MB)", $InputFileSize)

        $ProcessedSize += ($File.Length.ToString())

        $VideoProcessTimer = [Diagnostics.stopwatch]::StartNew()
        If(!($ValidateOnly)){
            $OptimizeVideo = &ffmpeg.exe $EscapeParser " -y -i `"$InputFileName`" -s hd480 -c:v libx264 -crf 23 -c:a aac -strict -2 `"$OutputFileName`""
        }
    } Else {
        Write-Host "Optimzing any videos that are older than $OptimizeAfterDays days, The video is not old enough to be optimized" -ForegroundColor Yellow
        Continue
    }

    If (Test-Path $OutputFileName){
        $OutputFileDuration = &ffprobe.exe $EscapeParser " -v error -show_format -select_streams v `"$OutputFileName`"" | Select-String "duration="
        $OutputFileDurationTime = New-TimeSpan -Seconds ($OutputFileDuration.ToString()).Split('=')[1].Split('.')[0] 
        $OutputFileSize = [math]::Round(((Get-Item $OutputFileName).Length / 1MB).ToString(),2)
        $OptimizedTotalSize +=  ($OutputFileName.Length).ToString()

        Write-Host "Input File size is (MB):" $InputFileSize -ForegroundColor Yellow
        Write-Host "Output File size is (MB):" $OutputFileSize -ForegroundColor Yellow
        Write-Host
        Write-Host "Input video duration (minutes):" $InputFileDurationTime.TotalMinutes -ForegroundColor Yellow
        Write-Host "Output video duration (minutes):" $OutputFileDurationTime.TotalMinutes -ForegroundColor Yellow
        Write-Host
        
        $Props.Add("Optimized Size (MB)", $OutputFileSize)

        If(($OutputFileDurationTime -eq $InputFileDurationTime) -and ($OutputFileDuration -ne $null)){
            Write-Host "Successfully optimized the video" -BackgroundColor Green -ForegroundColor Black
            Write-Host "The optimized file has been saved to:" $OutputFileName -BackgroundColor Green -ForegroundColor Black
            Write-Host
            Write-Host "Attempting to remove file:" $InputFileName
            
            $StopWatch = [Diagnostics.stopwatch]::StartNew()
            While ((Test-Path $InputFileName) -and ($StopWatch.Elapsed -lt $TimeOut)){
                If($DeleteOriginalVideo){
                    Remove-Item -Path $InputFileName
                } Else {
                    Write-Host "`"Delete Original Video`" is set to False. The original video file has not been deleted." -ForegroundColor Yellow
                }
                If (Test-Path $InputFileName){Start-Sleep 5}
            }
            If ($StopWatch.Elapsed -gt $TimeOut){Write-Host "Timed out while trying to delete" $InputFileName -BackgroundColor Red -ForegroundColor Black}
        } Else {
            Write-Host "An inconsistency has been detected in the output file. As a precaution the input file won't be deleted" -BackgroundColor Yellow -ForegroundColor Black
            Write-Host "Deleting the output file and keeping the original input file" -BackgroundColor Yellow -ForegroundColor Black
            Write-Host
            Write-Host "Attempting to remove file:" $OutputFileName -BackgroundColor Yellow -ForegroundColor Black
            
            $StopWatch = [Diagnostics.stopwatch]::StartNew()
            While ((Test-Path $OutputFileName) -and ($StopWatch.Elapsed -lt $TimeOut)){
                Remove-Item -Path $OutputFileName
                If (Test-Path $OutputFileName){Start-Sleep 5}
            }
            If ($StopWatch.Elapsed -gt $TimeOut){Write-Host "Timed out while trying to delete" $OutputFileName -BackgroundColor Red -ForegroundColor Black}
        }
    }

    $VideosProcessed += New-Object PSObject -Property $Props

    Write-Host "The time to complete this video was (hh:mm:ss):" $VideoProcessTimer.Elapsed.ToString("hh\:mm\:ss") -ForegroundColor Yellow
    Write-Host 
    Write-Host "The script has been running for a total time of (hh:mm:ss):" $ScriptRunningTime.Elapsed.ToString("hh\:mm\:ss") -ForegroundColor Yellow
    Write-Host
    Write-Host "Processed $NumFilesCounter of $NumFiles videos" -ForegroundColor Cyan
    Write-Host "Proccessed" ([math]::Round(($ProcessedSize / 1GB),2)) "GB of the total " ([math]::Round(($OriginalTotalSize / 1GB),2)) "GB" -ForegroundColor Cyan
    Write-Host

}
Write-Host
Write-Host "The following vidoes have been processed" -BackgroundColor Green -ForegroundColor Black
$VideosProcessed | Select "Name", "Age (Days)", "Original Size (MB)", "Optimized Size (MB)"

If($OutputToCsv){$VideosProcessed | Export-Csv -Path $OutputToCsv -NoTypeInformation}


Write-Host
Write-Host "After Processing the directory:" $Directory -ForegroundColor Cyan
Write-Host "The original total size was:" ([math]::Round(($OriginalTotalSize / 1GB),2)) "GB" -ForegroundColor Cyan
Write-Host "The optimized total size is:" ([math]::Round(($OptimizedTotalSize / 1GB),2)) "GB" -ForegroundColor Cyan
Write-Host "Successfully reduced the total size by" ([math]::Round((($OriginalTotalSize - $OptimizedTotalSize) / 1GB),2)) "GB" -ForegroundColor Cyan

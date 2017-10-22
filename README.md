# OptimizeMedia
Save space in your media directories by optimizing all media files to a lower quality and size after a specified number of days have past.

This script will take a directory and recursively search it for all known media files. 
If the files have existed for a specified time period they will then be optimized to a lower quality and size.

Requires FFMPEG to be installed and the /bin directory setup in the windows PATH. 
The script utilises both FFMPEG and FFPROBE. Both of these executables should exist in the FFMPEG /bin directory.

The script can be executed with the following parameters: 
  -Directory
  -OptimizeAfterDays
  -ValidateOnly
  -DeleteOriginalVideo
 
 -Directory
    The directory to search for the media. This directory will be searched recursively.
    
-OptimizeAfterDays
    The number of days a file must exist for before it is optimized. If the file creation date is older than the threshold, the media will be optimized.
    
-ValidateOnly
    This is a switch. Including this parameter in the command will stop the script from optimizing or deleting any files.
    It will only report on what files are going to be modified.
    
-DeleteOriginalVideo
    This is a switch. Including this parameter in the command will stop the script from deleting the original input file.
    

Examples:
.\OptimizeMedia.ps1 -Directory "C:\Temp" -OptimizeAfterDays "30" -DeleteOriginalVideo -ValidateOnly
      This will search c:\Temp for any media files that were created more than 30 days ago.
      No files will be optimized or deleted. Only a validation will run.
      
.\OptimizeMedia.ps1 -Directory "C:\Temp" -OptimizeAfterDays "-30"
      This will search c:\Temp for any media files that were created more than 30 days ago.
      Media files will be optimized to a lower quality and size, then the original file will be deleted.

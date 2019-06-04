# PHASE 1 - PREPARE FONT FILES FOR INSTALLATION

# UNZIP .ZIP FILES FROM THE FOLDER ARCHIVES ARE DOWNLOADED TO
Get-ChildItem C:\Users\Sigil\Desktop\FontDownloads\*.zip | Expand-Archive -DestinationPath C:\Users\Sigil\Desktop\FontDownloads\Staging\

# MOVE ALL FILES FROM SUBDIRECTORIES
Get-ChildItem C:\Users\Sigil\Desktop\FontDownloads\Staging\* -Recurse -File | Move-Item -Destination C:\Users\Sigil\Desktop\FontDownloads\Staging\

# REMOVE ALL FILES NOT .otf OR .tff
Get-ChildItem C:\Users\Sigil\Desktop\FontDownloads\Staging\ -Exclude *.otf, *.tff -Recurse | Remove-Item

# **HELP** BASIC MOVE DOES NOT WORK: I BELIEVE I NEED HELP WITH REGEDIT FOR FONT INSTALLATION USING New-Object AND ComObject PARAMETER
# PHASE 2 - UNTIL I GET THE ABOVE TO WORK, DRAG AND DROP FILES TO FONT DIRECTORY FOR INSTALLATION AND RUN PHASE 3 IN SEPARATE SCRIPT

# PHASE 2 - ARCHIVE .ZIP FILES THAT HAVE BEEN PROCESSED

# MOVE ZIP TO ARCHIVE
Get-ChildItem C:\Users\Sigil\Desktop\FontDownloads\* -Recurse -File | Move-Item -Destination C:\Users\Sigil\Desktop\FontDownloads\InstalledFonts\

#REMOVE ALL .otf AND .tff
Get-ChildItem C:\Users\Sigil\Desktop\FontDownloads\InstalledFonts\ -Exclude *.zip -Recurse | Remove-Item
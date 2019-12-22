# Mikes-Image-Workflow

This PowerShell script takes image files and runs them through a workflow where they are processed according to rules depending on file format and image properties. Processing includes removal of metadata, renaming, resizing, error checking and finally copying to one of the output folders.

It was made specifically to fit the needs of my friend Mike. As such the script is not intended to be useful for other people. But it does contain a lot of code showing how to use PowerShell to perform various task.

# Functionality

* Robust execution of external command line utilities
  * Supply arguments
  * Captures output
* File renaming
  * Removes non-ascii and other unwanted characters
  * RegEx replacement rules
  * Replace diacritics with "normal" character
  * Uses code page conversion to match to suitable ascii character
* Error checking
  * Uses [TrID](http://mark0.net/soft-trid-e.html) to detect and fix files with wrong file extensions
  * Find erroneous image files
* Execute [ExifTool](http://www.sno.phy.queensu.ca/~phil/exiftool/)
  * Uses an args-file
  * Runs in batches of configurable size
  * Catches errors
* Execute [ImageMagick](http://www.imagemagick.org/) or [GraphicsMagick](http://www.graphicsmagick.org/)
  * Convert images to JPEG
  * Scales images to max width/height
  * Extract image properties such as color space, icc color profile, bit depth, estimated quality
* Read image properties
  * Use Windows Image Acquisition (WIA) to read filetype, Width, height, size, number of frames
  * Read file bytes to detect JPG mode (baseline/progressive/arithmetic/lossless/hierarchical)
* Move files to destination folders based on their properties
  * Ensures output folders exist
  * Overwrite or rename existing files
  * Add a number suffix in case of duplicate file names
* Contains a simple WinForms GUI made with PowerShell
  * For a full tutorial see my [blog](http://www.rlvision.com/blog/a-drag-and-drop-gui-made-with-powershell/)

# How to use

Instructions on how to get going and how to use the script are located as comments inside the script, at the top and bottom.

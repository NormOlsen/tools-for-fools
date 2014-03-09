' VB Script to merge text files with the option of skipping headers

Option Explicit

Dim pFSO, pTSIn, pTSOut, pFolder, pFile, pRegExp, pMatches
Dim strFilename, strPattern, intSkip, bolFirst, bolFirstFile, i

' We need at least one parameter to continue (the name of the merged file).
If WScript.Arguments.Count = 0 Then
   WScript.Echo "MergeFiles - Merge multiple text files into a single file."
   WScript.Echo "Usage:"
   WScript.Echo "MergeFiles <merge filename> [pattern] [skip] [leave first]"
   WScript.Echo "Where [pattern] is a file pattern, [skip] indicates the"
   WScript.Echo "number of header lines to skip, and [leave first] indicates"
   WScript.Echo "whether or not to apply the [skip] parameter to the first"
   WScript.Echo "file (true or false)."
   WScript.Quit
End If

Select Case WScript.Arguments.Count
   Case 1
      strFilename = WScript.Arguments(0)
   Case 2
      strFilename = WScript.Arguments(0)
      strPattern = WScript.Arguments(1)
   Case 3
      strFilename = WScript.Arguments(0)
      strPattern = WScript.Arguments(1)
      intSkip = CInt(WScript.Arguments(2))
   Case 4
      strFilename = WScript.Arguments(0)
      strPattern = WScript.Arguments(1)
      intSkip = CInt(WScript.Arguments(2))
      bolFirst = CBool(WScript.Arguments(3))
End Select

Set pRegExp = New RegExp
pRegExp.IgnoreCase = True
Set pFSO = CreateObject("Scripting.FileSystemObject")
Set pFolder = pFSO.GetFolder(".")

' Ignore line skips on first file?
If IsEmpty(bolFirst) Then
   bolFirst = False
End If

' Number of lines to skip in each file
If IsEmpty(intSkip) Then
   intSkip = 0
End If

' Set the file matching pattern
If IsEmpty(strPattern) Then
   pRegExp.Pattern = "*"
Else
   pRegExp.Pattern = strPattern
End If

Set pTSOut = pFSO.CreateTextFile(strFilename)
bolFirstFile = True
For Each pFile In pFolder.Files
   Set pMatches = pRegExp.Execute(pFile.Name)
   If pFile.Name <> strFilename And pMatches.Count <> 0 Then
      Set pTSIn = pFile.OpenAsTextStream
      If intSkip > 0 Then
         If (Not bolFirst And bolFirstFile) Or (Not bolFirstFile) Then ' skip lines
           For i = 1 To intSkip
               pTSIn.SkipLine
            Next
         End If
         bolFirstFile = False
      End If
      Do While Not pTSIn.AtEndOfStream
         pTSOut.WriteLine pTSIn.ReadLine
      Loop
      pTSIn.Close
   End If
Next
pTSOut.Close


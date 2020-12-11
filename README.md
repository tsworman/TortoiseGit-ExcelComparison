# TortoiseGit-ExcelComparison #

## About ##
This script enables you to compare (in a meaninful way) data between two XLSX spreadsheets. It's designed for use with TortoiseGit and it should work with TortoiseSVN.

This was originally found here: https://www.cnblogs.com/micele/p/5014037.html

I modified it to more closely match the latest TortoiseGit scripts and to get it working with Spreadsheet compare 2016 and Office 365.

## How to Install ##

Download the .vbs file and place it in: C:\Program Files\TortoiseGit\Diff-Scripts (or wherever you have TortoiseGit Installed)

Launch TortoiseGit settings (right click on desktop, TortoiseGit->Settings) and pick Diff Viewer. 

Under diff viewer pick advanced.


Find the existing xslx entry and delete or change the extension for it. 
I renamed the existing entry to .xslx2 so that I could still use the original functionality if needed by simply renaming my files. 
 
Add a line for .xslx and set it to:
wscript.exe "C:\Program Files\TortoiseGit\Diff-Scripts\diff-xlsx-ssc.vbs" %base %mine //E:vbscript

## How to use ##
Assuming you have spreadsheet compare 2016 (or Office 365) installed, it should launch when you perform a diff from TortoiseGit. 

Spreadsheet Compare will simply tell you what changed between files. It has no facilities to perform the merge of the changes. 

You can use this to manually update your working copy, mark conflicts as resolved, and perform a manual merge-commit.

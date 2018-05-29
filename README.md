# rpt_diff
Program that can compare Crystal reports templates

How to use:
1. You have to install Visual Studio 2010 or newer
2. You have to install some kind of program that can compare two XML files (ex. KDiff)
3. Download and install Crystal reports for Visual Studio (https://www.crystalreports.com/crystal-reports-visual-studio/)
4. Now you can use rpt_diff.exe application 

Usage of the rpt_diff.exe:
  rpt_diff.exe DiffUtilPath RPTPath1 [RPTPath2] ModelNumber
    DiffUtilPath - Full path to external diff application .exe file that can compare two xml files (for example KDiff)
    RPTPath1 - Full path to first .rpt file to be converted to xml
    RPTPath2 - Full path to second .rpt file to be converted to xml and compared with first file
    ModelNumber - Select between object model 0 - ReportDocument or 1 (recomended)- ReportClientDocument 

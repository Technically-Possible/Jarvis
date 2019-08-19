Dim xmlDoc, objNodeList, plot

Set wshShell = CreateObject( "WScript.Shell" )
tfolder =  wshShell.ExpandEnvironmentStrings("%TEMP%")
Set xmlDoc = CreateObject("Msxml2.DOMDocument")
xmlDoc.load(tfolder & "\Next3DaysRSS.xml")
Set objNodeList = xmlDoc.getElementsByTagName("channel/item/title") 'Node to search for
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Write all found results into forecast.txt
Const ForWriting = 2
Set objTextFile = objFSO.OpenTextFile _
    (tfolder & "\forecast.txt", ForWriting, True)
If objNodeList.length > 0 then
For each x in objNodeList
plot=x.Text
objTextFile.WriteLine(plot)
Next
objTextFile.Close	
End If

'Extract tomorrows data (second line) from 'forecast.txt' and write each data type to seperate line in tomorrow.txt
Const ForReading = 1
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.OpenTextFile _
    (tfolder & "\forecast.txt", ForReading)
	objTextFile.Skipline
    strNextLine = objTextFile.Readline
    currentsplit = Split(strNextLine , ", ")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.OpenTextFile(tfolder & "\tomorrow.txt", ForWriting, True)	
    objTextFile.WriteLine(currentsplit(0))
    For i = 1 to Ubound(currentsplit)
        objTextFile.WriteLine(currentsplit(i))
		Next		

' Get individual data

' Condition
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.OpenTextFile _
    (tfolder & "\tomorrow.txt", ForReading)
    	strNextLine = objTextFile.Readline
	conditionsplit = Split(strNextLine , ": ")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.OpenTextFile(tfolder & "\condition.txt", ForWriting, True)	
    objTextFile.WriteLine(conditionsplit(1)) 'Number value determines which line to write; 0 = 1st

' Maximum Temperature
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.OpenTextFile _
    (tfolder & "\tomorrow.txt", ForReading)
    objTextFile.Skipline
	strNextLine = objTextFile.Readline
	maxtempsplit = Split(strNextLine , "°C")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.OpenTextFile(tfolder & "\maxtemp.txt", ForWriting, True)	
    objTextFile.WriteLine(maxtempsplit(0))
	
' Minimum Temperature
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.OpenTextFile _
    (tfolder & "\tomorrow.txt", ForReading)
    objTextFile.Skipline
	objTextFile.Skipline
	strNextLine = objTextFile.Readline
	mintempsplit = Split(strNextLine , "°C")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.OpenTextFile(tfolder & "\mintemp.txt", ForWriting, True)	
    objTextFile.WriteLine(mintempsplit(0))
    

'''''''''''''''''' THE DAY AFTER TOMORROW'''''''''''''''

'Extract the day after tomorrows data (third line) from 'forecast.txt' and write each data type to seperate line in aftertomorrow.txt
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.OpenTextFile _
    (tfolder & "\forecast.txt", ForReading)
	objTextFile.Skipline
	objTextFile.Skipline
    strNextLine = objTextFile.Readline
    currentsplit = Split(strNextLine , ", ")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.OpenTextFile(tfolder & "\aftertomorrow.txt", ForWriting, True)	
    objTextFile.WriteLine(currentsplit(0))
    For i = 1 to Ubound(currentsplit)
        objTextFile.WriteLine(currentsplit(i))
		Next		

' Get individual data

' Day
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.OpenTextFile _
    (tfolder & "\aftertomorrow.txt", ForReading)
    	strNextLine = objTextFile.Readline
	conditionsplit = Split(strNextLine , ": ")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.OpenTextFile(tfolder & "\day.txt", ForWriting, True)	
    objTextFile.WriteLine(conditionsplit(0)) 'Number value determines which line to write; 0 = 1st

' Condition
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.OpenTextFile _
    (tfolder & "\aftertomorrow.txt", ForReading)
    	strNextLine = objTextFile.Readline
	conditionsplit = Split(strNextLine , ": ")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.OpenTextFile(tfolder & "\condition2.txt", ForWriting, True)	
    objTextFile.WriteLine(conditionsplit(1)) 'Number value determines which line to write; 0 = 1st

' Maximum Temperature
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.OpenTextFile _
    (tfolder & "\aftertomorrow.txt", ForReading)
    objTextFile.Skipline
	strNextLine = objTextFile.Readline
	maxtempsplit = Split(strNextLine , "°C")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.OpenTextFile(tfolder & "\maxtemp2.txt", ForWriting, True)	
    objTextFile.WriteLine(maxtempsplit(0))
	
' Minimum Temperature
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.OpenTextFile _
    (tfolder & "\aftertomorrow.txt", ForReading)
    objTextFile.Skipline
	objTextFile.Skipline
	strNextLine = objTextFile.Readline
	mintempsplit = Split(strNextLine , "°C")
Set objFSO = CreateObject("Scripting.FileSystemObject")
Set objTextFile = objFSO.OpenTextFile(tfolder & "\mintemp2.txt", ForWriting, True)	
    objTextFile.WriteLine(mintempsplit(0))
    
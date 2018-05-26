'Initialize array of files to change
Dim filePaths(8)

'Set array of files to change
filePaths(0) = "C:\Users\wheastie\Desktop\Database Strings Script\File Paths\1st Path\application.properties"
filePaths(1) = "C:\Users\wheastie\Desktop\Database Strings Script\File Paths\1st Path\bootstrap.properties"
filePaths(2) = "C:\Users\wheastie\Desktop\Database Strings Script\File Paths\1st Path\web.config"
filePaths(3) = "C:\Users\wheastie\Desktop\Database Strings Script\File Paths\2nd Path\application-env.properties"
filePaths(4) = "C:\Users\wheastie\Desktop\Database Strings Script\File Paths\2nd Path\bootstrap.properties"
filePaths(5) = "C:\Users\wheastie\Desktop\Database Strings Script\File Paths\2nd Path\web.config"
filePaths(6) = "C:\Users\wheastie\Desktop\Database Strings Script\File Paths\3rd Path\application.properties"
filePaths(7) = "C:\Users\wheastie\Desktop\Database Strings Script\File Paths\3rd Path\context.xml"

'Set user database information
serverName = "DEL8PWW1\\SQLEXPRESS"
consoleDatabaseName = "Console"
datamartDatabaseName = "Datamart"
portNumber = "1433"
userDBId = "sa"
userDBPassword = "Password1234"

'Update each file in the filePaths array
For i = 0 to 7
	Call UpdateFile()
Next

'Let user know script is done running
WScript.Echo "Script has finished!"

'This function is used to update each file in the filePaths array
Function UpdateFile()

	'Check to see if there is a file path
	If(filePaths(i) <> "") Then
	
		'Initialize array of the previous keys & values and new keys & values
		Dim oldValues(5)
		Dim newValues(5)
		
		'Set the index of values to be changed
		valuesIndex = 0
		
		'Get the current file type
		fileType = GetFileType(filePaths(i))
		
		'Set File System Object and read to file variable
		Set fso = CreateObject("Scripting.FileSystemObject")
		Set readFileOne = fso.OpenTextFile(filePaths(i), 1)
		
		'Perform UpdateApplicationFile function if current file is application.properties or application-env.properties
		If(fileType = "application.properties" OR fileType = "application-env.properties") Then
			Call UpdateApplicationFile(readFileOne, valuesIndex, oldValues, newValues)
		'Perform UpdateBootstrapFile function if current file is bootstrap.properties
		ElseIf(fileType = "bootstrap.properties") Then
			Call UpdateBootstrapFile(readFileOne, valuesIndex, oldValues, newValues)
		'Perform UpdateContextFiile function if current file is context.xml
		ElseIf(fileType = "context.xml") Then
			Call UpdateContextFile(readFileOne, valuesIndex, oldValues, newValues)
		'Perform UpdateWebFile function if current file is web.config
		ElseIf(fileType = "web.config") Then
			Call UpdateWebFile(readFileOne, valuesIndex, oldValues, newValues)
		End If
		
		'Close readFileOne
		readFileOne.Close
		Set readFileOne = Nothing
		
		'Reopen file and read in all of its text into a string
		Set readFileTwo = fso.OpenTextFile(filePaths(i), 1)
		fileContent = readFileTwo.ReadAll
		
		'Close readFileTwo
		readFileTwo.Close
		Set readFileTwo = Nothing
		
		'Update the old values with the new values in the fileContent string
		For valueCounter = 0 to valuesIndex - 1
			fileContent = Replace(fileContent, oldValues(valueCounter), newValues(valueCounter))
		Next
		
		'Set the write to file variable
		Set writeFile = fso.OpenTextFile(filePaths(i), 2)
		
		'Write to the file and update the information
		writeFile.Write(fileContent)
		
		'Close writeFile
		writeFile.Close
		Set writeFile = Nothing	
		
	End If
	
End Function

'This function is used to update application.properties and application-env.properties files
Function UpdateApplicationFile(file, valuesIndex, oldValues(), newValues())

	'Set max values to be changed
	maxValuesToChange = 3
	
	'Set url
	databaseConnectionURL = "jdbc:sqlserver://" & serverName & ":" & portNumber & ";databaseName=" & consoleDatabaseName
	
	'Iterate each line of the file and get keys that need to be updated
	Do Until file.AtEndOfStream
		currentFileLine = file.ReadLine
		lineKey = Left(currentFileLine, InStr(currentFileLine, "="))
		If(valuesIndex < maxValuesToChange) Then
			If(lineKey = "spring.datasource.url=") Then
				oldValues(valuesIndex) = currentFileLine
				newValues(valuesIndex) = "spring.datasource.url=" & databaseConnectionURL
				valuesIndex = valuesIndex + 1
			ElseIf(lineKey = "spring.datasource.username=") Then
				oldValues(valuesIndex) = currentFileLine
				newValues(valuesIndex) = "spring.datasource.username=" & userDBId
				valuesIndex = valuesIndex + 1			
			ElseIf(lineKey = "spring.datasource.password=") Then
				oldValues(valuesIndex) = currentFileLine
				newValues(valuesIndex) = "spring.datasource.password=" & userDBPassword
				valuesIndex = valuesIndex + 1					
			End If
		End If
	Loop

End Function

'This function is used to update the bootstrap.properties file
Function UpdateBootstrapFile(file, valuesIndex, oldValues(), newValues())

	'Set max values to be changed
	maxValuesToChange = 5
	
	'Add '\' to serverName
	rootServerName = serverName & "\"
	
	'Set local SQL server to only name
	rootServerName = Left(rootServerName, Instr(rootServerName, "\") - 1)
	
	'Iterate each line of the file and get keys that need to be updated
	Do Until file.AtEndOfStream
		currentFileLine = file.ReadLine
		lineKey = Left(currentFileLine, InStr(currentFileLine, "="))
		If(valuesIndex < maxValuesToChange) Then
			If(lineKey = "OC.db.server=") Then
				oldValues(valuesIndex) = currentFileLine
				newValues(valuesIndex) = "OC.db.server=" & rootServerName 'May be an issue with DAL7NWH1\\SQLEXPRESS
				valuesIndex = valuesIndex + 1
			ElseIf(lineKey = "OC.db.sid=") Then
				oldValues(valuesIndex) = currentFileLine
				newValues(valuesIndex) = "OC.db.sid=" & consoleDatabaseName
				valuesIndex = valuesIndex + 1			
			ElseIf(lineKey = "OC.db.port=") Then
				oldValues(valuesIndex) = currentFileLine
				newValues(valuesIndex) = "OC.db.port=" & portNumber
				valuesIndex = valuesIndex + 1			
			ElseIf(lineKey = "OC.db.userid=") Then
				oldValues(valuesIndex) = currentFileLine
				newValues(valuesIndex) = "OC.db.userid=" & userDBId
				valuesIndex = valuesIndex + 1
			ElseIf(lineKey = "OC.db.password=") Then
				oldValues(valuesIndex) = currentFileLine
				newValues(valuesIndex) = "OC.db.password=" & userDBPassword
				valuesIndex = valuesIndex + 1		
			End If
		End If
	Loop

End Function

'This function is used to update the context.xml file
Function UpdateContextFile(file, valuesIndex, oldValues(), newValues())
	
	'Initialize array of default keys
	Dim defaultKeys(3)
	
	'Set array of default keys
	defaultKeys(0) = "password="
	defaultKeys(1) = "url="
	defaultKeys(2) = "username="
	
	'Set max values to be changed
	maxValuesToChange = 3
	
	'Set database connection url
	databaseConnectionURL = chr(34) & "jdbc:sqlserver://" & serverName & ":" & portNumber & ";databaseName=" & consoleDatabaseName & chr(34)
	
	'Iterate each line of the file and get keys that need to be updated
	Do Until file.AtEndOfStream
		currentFileLine = file.ReadLine
		If(valuesIndex < maxValuesToChange) Then
			oldValueString = SearchString(currentFileLine, defaultKeys(valuesIndex), 0)
			If(Left(oldValueString, InStr(oldValueString, "=")) = "password=") Then
				oldValues(valuesIndex) = oldValueString
				newValues(valuesIndex) = defaultKeys(valuesIndex) & chr(34) & userDBPassword & chr(34)
				valuesIndex = valuesIndex + 1
			ElseIf(Left(oldValueString, InStr(oldValueString, "=")) = "url=") Then
				oldValues(valuesIndex) = oldValueString
				newValues(valuesIndex) = defaultKeys(valuesIndex) & databaseConnectionURL
				valuesIndex = valuesIndex + 1				
			ElseIf(Left(oldValueString, InStr(oldValueString, "=")) = "username=") Then
				oldValues(valuesIndex) = oldValueString
				newValues(valuesIndex) = defaultKeys(valuesIndex) & chr(34) & userDBId & chr(34)
				valuesIndex = valuesIndex + 1				
			End If
		End If
	Loop
	
End Function

'This function is used to update the web.config file
Function UpdateWebFile(file, valuesIndex, oldValues(), newValues())
	
	'Initialize array of default keys
	Dim defaultKeys(2)
	
	'Set array of default keys
	defaultKeys(0) = "connectionString="
	defaultKeys(1) = "key=" & chr(34) & "ConsoleDBSchema" & chr(34)
	
	'Set database connection url
	databaseConnectionURL = chr(34) & "Data Source=localhost;Initial Catalog=" & datamartDatabaseName & ";User Id=" & userDBId & ";Password=" & userDBPassword & chr(34)
	
	'Set the max number of values to be changed
	maxValuesToChange = 2	
	
	'Iterate each line of the file and get keys that need to be updated
	Do Until file.AtEndOfStream
		currentFileLine = file.ReadLine		
		If(valuesIndex < maxValuesToChange) Then
			'Code to perform if defaultKeys is "connectionString="
			If(defaultKeys(valuesIndex) = "connectionString=") Then
				oldValueString = SearchString(currentFileLine, defaultKeys(valuesIndex), 0)
				If(Left(oldValueString, InStr(oldValueString, "=")) = defaultKeys(valuesIndex)) Then
					oldValues(valuesIndex) = oldValueString
					newValues(valuesIndex) = "connectionString=" & databaseConnectionURL
					valuesIndex = valuesIndex + 1
				End If
			'Code to perform if defaultKeys is key="ConsoleDBSchema"
			ElseIf(defaultKeys(valuesIndex) = "key=" & chr(34) & "ConsoleDBSchema" & chr(34)) Then
				oldValueString = SearchString(currentFileLine, defaultKeys(valuesIndex), 1)
				If(oldValueString = defaultKeys(valuesIndex)) Then
					keyOfOldValueString = SearchString(currentFileLine, "value=", 0)
					oldValues(valuesIndex) = keyOfOldValueString
					newValues(valuesIndex) = "value=" & chr(34) & consoleDatabaseName & chr(34)
					valuesIndex = valuesIndex + 1
				End If			
			End If
		End If
	Loop

End Function

'This function is used to return the current files type
Function GetFileType(filePath)
	
	'Return the type of file
	GetFileType = Right(filePath, InStr(StrReverse(filePath), "\") - 1)
	
End Function

'This function is used to return a specified key in the file's current line
Function SearchString(currentLine, comparisonKey, compareOption)	
	
	'Add additional character to currentLine to complete the loop process
	currentLine = currentLine & " "
	
	'Set default inQuotations
	inQuotations = False	
	
	'Loop through the current line to find the comparisonKey
	For x = 1 to len(currentLine)
		currentCharacter = Mid(currentLine, x, 1)
		If((currentCharacter <> " " AND currentCharacter <> chr(09)) OR inQuotations = True) Then
			If(currentCharacter = chr(34)) Then
				If(inQuotations = False) Then
					inQuotations = True
				Else
					inQuotations = False
				End If
			End If
			tempString = tempString & currentCharacter
		Else
			If(compareOption = 1) Then
				If(tempString = comparisonKey) Then
					SearchString = tempString
				End If
			Else
				tempSubString = Left(tempString, InStr(tempString, "="))
				If(tempSubString = comparisonKey) Then
					SearchString = tempString
				End If
			End If
			tempString = ""
		End If
	Next
End Function
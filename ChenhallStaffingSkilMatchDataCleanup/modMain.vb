Imports System.IO

Module modMain
    Dim CommandLineArgs As New List(Of ArgumentPairs)

    ' Application Settings
    Private ReadOnly appTitle As String = "Chenhall Staffing SkilMatch Data Cleanup"
    Dim sNoticeMessage As Boolean = True
    Dim sWarningMessage As Boolean = True
    Dim sErrorMessage As Boolean = True
    Dim sMessage As Boolean = True
    Dim sPause As Boolean = True

    Sub Main()
        Dim colorEntryFg As ConsoleColor = Console.ForegroundColor
        Dim colorEntryBg As ConsoleColor = Console.BackgroundColor

        SetConsoleColors(ConsoleColor.White, ConsoleColor.Black)
        Console.Title = appTitle

        ' Default Application Settings
        Dim sClearConsole As Boolean = True
        Dim sFolderLocation As String = Environment.CurrentDirectory
        Dim sSaveLocation As String = Environment.CurrentDirectory

        ' Handle Console Line Arguments
        If ParseCommandLine() Then
            ' We are no longer using defaults!
            For Each commandItem As ArgumentPairs In CommandLineArgs
                Select Case commandItem.Name.ToLower
                    Case "h", "?", "help", "info"
                        HelpPrint()
                        PauseApplication(True)
                    Case "f"
                        sFolderLocation = commandItem.Value
                    Case "s"
                    	sSaveLocation = commandItem.Value
                    Case "x"
                    	If commandItem.Value.Contains("c") Then
                    		sClearConsole = false
                    	End If
                    	If commandItem.Value.Contains("n") Then
                    		sNoticeMessage = False
                    	End If
                    	If commandItem.Value.Contains("w") Then
                    		sWarningMessage = False
                    	End If
                    	If commandItem.Value.Contains("e") Then
                    		sErrorMessage = False
                    	End If
                    	If commandItem.Value.Contains("p") Then
                    		sPause = False
                    	End If
                    	If commandItem.Value.Contains("m") Then
                    		sMessage = False
                    	End If
                    Case "silent"
                        sMessage = False
                        sErrorMessage = False
                        sWarningMessage = False
                        sNoticeMessage = False
                End Select
            Next
        End If

        ' Directory Contents
        WriteMessage("Starting processing...")
        WriteMessage("Looking for files...")
        Dim files As String() = Directory.GetFiles(sFolderLocation, "*.csv")
        Dim foundAppmas As Boolean = False

		Dim removedSSN As List(Of Integer) = New List(Of Integer)
		Dim removedCCN As List(Of String) = New List(Of String)
		Dim removedJON As List(Of Integer) = New List(Of Integer)
        Dim aStartDate As Date = Now
        WriteMessage("Starting: " & Now)
        
        WriteNoticeMessage("Found " & files.Length & " files matching pattern *.csv")
        
        If files.Length = 0 Then
            WriteErrorMessage("Found no files, ensure the path provided is correct or this file was ran from the desired working directory.")
            PauseApplication(True)
        Else
        	' Need to deal with empmas.csv first before touching any other file
        	WriteMessage("Looking for next file to work with...")
        	If files.Contains(sFolderLocation & "\empmas.csv") Then
        		WriteMessage("Found empmas.csv")
				Dim outputFile As String = sSaveLocation
        		If Not sSaveLocation = sFolderLocation Then
        			outputFile = sSaveLocation & "\empmas.csv"
        		Else
        			outputFile = sSaveLocation & "\empmas_mod.csv"
        		End If
        		
        		If File.Exists(outputFile) Then
                    WriteWarningMessage(outputFile & " already exists! It will be overwritten.")
                    PauseApplication()
                    File.Delete(outputFile)
        		End If
        		
                Dim totalLines As Integer = 0
                Try
                    totalLines = File.ReadAllLines(sFolderLocation & "\empmas.csv").Length
                Catch ex As Exception
                    WriteErrorMessage("Exception error happened while counting number of lines in empmas. Aborting processing completely.")
                    WriteErrorMessage("EXCEPTION MESSAGE: " & ex.Message)
                    PauseApplication(True)
                End Try
                
                Console.Title = "[empmas.csv] " & appTitle
                WriteMessage("Working on empmas.csv...")
                WriteMessage("For progress, check title bar")
                
                
                Using reader As New FileIO.TextFieldParser(sFolderLocation & "\empmas.csv")
                    reader.TextFieldType = FileIO.FieldType.Delimited
                    reader.SetDelimiters(",")
                    While Not reader.EndOfData
                        Try
                            Dim percent As Integer = (reader.LineNumber / totalLines) * 100
                            Console.Title = "[" & percent & "% - " & reader.LineNumber & "/" & totalLines & "] [empmas.csv] " & appTitle
                            Dim outputRow As String = ""
                            Dim currentRow As String() = reader.ReadFields()
                            
                            ' EMSS# contains SSN currently being worked on, acts as the key in all other files
                            Dim currentSSN As Integer = Int(currentRow(2))
                            
                            ' idx 61, 63, 64 contains EMPAP2K, EMHR2K, EMLS2K in YYYYMMDD format
                            ' need to check that these are all 2009 or higher to be kept
                            If Int(currentRow(61).PadLeft(8, "0").Substring(0, 4)) < 2009 And _
                            	Int(currentRow(63).PadLeft(8, "0").Substring(0, 4)) < 2009 And _
                            	Int(currentRow(64)) < 20081227 Then
                            	
                            	' Since we're all under 2009, they can be removed
                            	currentRow = Nothing
                            	GoTo SkipRowEntryEMPMAS
                            End If
                            ' Else we are good to go on modifying the user information
                            
                            ' Verify Category Code
                            Dim validCategoryCodes As String() = New String() {"SS", "JV"}
                            If Not validCategoryCodes.Contains(currentRow(26)) Then
                            	currentRow(26) = "SS"
                            End If
                            
                            ' Verify Interviewer Code
                			Dim validInterviewerCodes As String() = New String() {"MG", "AM", "XX"}
                            If Not validInterviewerCodes.Contains(currentRow(43)) Then
                            	currentRow(43) = "XX"
                            End If
                            
                            ' Clear Counselor Code
                            currentRow(44) = ""
                            
                            ' Verify Division code
                            Dim validDivisionCodes As String() = New String() {"C"}
                            If Not validDivisionCodes.Contains(currentRow(46)) Then
                            	currentRow(46) = "C"
                            End If
                            
                            ' Verify Status Code
                            Dim validStatusCodes As String() = New String() {"A", "I", "D"}
                            If Not validStatusCodes.Contains(currentRow(47)) Then
                            	currentRow(47) = "A"
                            End If
                            If currentRow(47) = "I" Then
                            	WriteMessage("SS# " & currentRow(2).padLeft(9, "0") & " is marked as inactive but did not match inactive criteria.")
                            End If
                            
SkipRowEntryEMPMAS:
                            If Not currentRow Is Nothing Then
                                outputRow = String.Join(",", currentRow) & Environment.NewLine
                                File.AppendAllText(outputFile, outputRow, System.Text.Encoding.ASCII)
                            Else
                                ' Append SSN to the list of removed entries, to be later used in other files
                                removedSSN.Add(currentSSN)
                            End If
                        Catch ex As Exception
                            WriteErrorMessage("Exception error happened on line " & reader.LineNumber & ". Aborting processing completely.")
                            WriteErrorMessage("EXCEPTION MESSAGE: " & ex.Message)
                            PauseApplication(True)
                            
                        End Try
                    End While
                End Using
                
                ' Also need to deal with cusmas.csv before touching any other file
                If files.Contains(sFolderLocation & "\cusmas.csv") Then
	        		WriteMessage("Found cusmas.csv")
					outputFile = sSaveLocation
	        		If Not sSaveLocation = sFolderLocation Then
	        			outputFile = sSaveLocation & "\cusmas.csv"
	        		Else
	        			outputFile = sSaveLocation & "\cusmas_mod.csv"
	        		End If
	        		
	        		If File.Exists(outputFile) Then
	                    WriteWarningMessage(outputFile & " already exists! It will be overwritten.")
	                    PauseApplication()
	                    File.Delete(outputFile)
	        		End If
	        		
	                totalLines = 0
	                Try
	                    totalLines = File.ReadAllLines(sFolderLocation & "\cusmas.csv").Length
	                Catch ex As Exception
	                    WriteErrorMessage("Exception error happened while counting number of lines in cusmas. Aborting processing completely.")
	                    WriteErrorMessage("EXCEPTION MESSAGE: " & ex.Message)
	                    PauseApplication(True)
	                End Try
	                
	                Console.Title = "[cusmas.csv] " & appTitle
	                WriteMessage("Working on cusmas.csv...")
	                WriteMessage("For progress, check title bar")
	                
	                Using reader As New FileIO.TextFieldParser(sFolderLocation & "\cusmas.csv")
	                    reader.TextFieldType = FileIO.FieldType.Delimited
	                    reader.SetDelimiters(",")
	                    While Not reader.EndOfData
	                        Try
	                            Dim percent As Integer = (reader.LineNumber / totalLines) * 100
	                            Console.Title = "[" & percent & "% - " & reader.LineNumber & "/" & totalLines & "] [cusmas.csv] " & appTitle
	                            Dim outputRow As String = ""
	                            Dim currentRow As String() = reader.ReadFields()
	                            
	                            ' CSCODE contains customer code currently being worked on, acts as the key in all other files
	                            Dim currentCCN As String = currentRow(2)
	                            Dim currentCNN As Integer = Int(currentRow(1))
	                            
	                            
	                            ' idx 2 contains company code, it MUST be 10
	                            If Int(currentRow(1)) <> 10 Then
	                            	currentRow = Nothing
	                            	GoTo SkipRowEntryCUSMAS
	                            ' idx 61, 62 contains Y2K format YYYYMMDD
	                            ' need to check that these are all 2009 or higher to be kept
	                            ElseIf Int(currentRow(61)) < 20081227 And _
	                            	Int(currentRow(62)) < 20081227 Then
	                            	
	                            	' Since we're all under 20081227, they can be removed
	                            	currentRow = Nothing
	                            	GoTo SkipRowEntryCUSMAS
	                            End If
	                            

	                            
	                            ' Else we are good to go on modifying the cusomer information
	                            
	                            ' Mark them as active instead of deleted (possibly set)
	                            currentRow(0) = "A"
	                            	                            
SkipRowEntryCUSMAS:
	                            If Not currentRow Is Nothing Then
	                                outputRow = String.Join(",", currentRow) & Environment.NewLine
	                                File.AppendAllText(outputFile, outputRow, System.Text.Encoding.ASCII)
	                            Else
	                                ' Append CCN to the list of removed entries, to be later used in other files
	                                removedCCN.Add(currentCCN)
	                            End If
	                        Catch ex As Exception
	                            WriteErrorMessage("Exception error happened on line " & reader.LineNumber & ". Aborting processing completely.")
	                            WriteErrorMessage("EXCEPTION MESSAGE: " & ex.Message)
	                            PauseApplication(True)
	                            
	                        End Try
	                    End While
	                End Using
                Else
                	WriteErrorMessage("Cannot convert files! Missing required file cusmas.csv")
        			PauseApplication(True)
                End If
                
                ' Next we need to touch upon JOBMAS@.csv
                If files.Contains(sFolderLocation & "\ljobmas@.csv") Then
            		WriteMessage("Found ljobmas@...")
                    If Not sSaveLocation = sFolderLocation Then
                        outputFile = sSaveLocation & "\ljobmas@.csv"
                    Else
                        outputFile = sSaveLocation & "\ljobmas@_mod.csv"
                    End If

                    If File.Exists(outputFile) Then
                        WriteWarningMessage(outputFile & " already exists! It will be overwritten.")
                        PauseApplication()
                        File.Delete(outputFile)
                    End If

                    totalLines = File.ReadAllLines(sSaveLocation & "\ljobmas@.csv").Length

                    Console.Title = "[ljobmas@.csv] " & appTitle
                    WriteMessage("Working on ljobmas@.csv...")
                    WriteMessage("For progress, check title bar")
                    Using reader As New FileIO.TextFieldParser(sSaveLocation & "\ljobmas@.csv")
                        reader.TextFieldType = FileIO.FieldType.Delimited
                        reader.SetDelimiters(",")
                        While Not reader.EndOfData
                        	Try
                        		Dim percent As Integer = (reader.LineNumber / totalLines) * 100
                                Console.Title = "[" & percent & "% - " & reader.LineNumber & "/" & totalLines & "] [ljobmas@.csv] " & appTitle
                                Dim outputRow As String = ""
                                Dim currentRow As String() = reader.ReadFields()
                                
                                Dim currentJON As Integer = Int(currentRow(3))
                                
                                ' Strip all entries based on company code (<> 10), ssn, and company code that were removed
                                If Int(currentRow(1)) <> 10 Or removedSSN.Contains(Int(currentRow(4))) Or removedCCN.Contains(currentRow(5)) Then
                                	currentRow = Nothing
                                	GoTo SkipRowEntryJOBMAS
                                End If
                                ' Else we can modify data
                                
SkipRowEntryJOBMAS:
                                If Not currentRow Is Nothing Then
                                    outputRow = String.Join(",", currentRow) & Environment.NewLine
                                    File.AppendAllText(outputFile, outputRow, System.Text.Encoding.ASCII)
                                Else
                                	removedJON.Add(currentJON)
                                End If
                            Catch ex As Exception
                                WriteErrorMessage("Exception error happened on line " & reader.LineNumber & ". Aborting processing completely.")
                                WriteErrorMessage("EXCEPTION MESSAGE: " & ex.Message)
                                PauseApplication(True)
                            End Try
                        End While
                    End Using
					
					WriteMessage("Finished ljobmas@.csv modifications")	
                Else
                	WriteErrorMessage("Cannot convert files! Missing required file ljobmas@.csv")
                	PauseApplication(True)
                End If
                
                ' Deal with any other files now
                Dim ignoreFileList As String() = New String() {"appnav.csv", "cmttyp.csv", "concmt1.csv", "concmt2.csv", "cusdpt.csv", "lappcmt@.csv", "lappskla.csv", "lappsklb.csv", "lappsklc.csv", "lappskld.csv", "lappskle.csv", "lcuscmt4.csv", "lcuscmt5.csv", "empmasx.csv", "skfile.csv"}
                
                For Each fileName As String In files
				    Dim fileInfo As FileInfo = New FileInfo(fileName)
				    outputFile = sSaveLocation
				    

	                If fileInfo.Name = "appmas.csv" Then
	                    WriteMessage("Found " & fileInfo.Name & "...")
	                    If Not sSaveLocation = sFolderLocation Then
	                        outputFile = sSaveLocation & "\" & fileInfo.Name
	                    Else
	                        outputFile = sSaveLocation & "\appmas_mod.csv"
	                    End If
	
	                    If File.Exists(outputFile) Then
	                        WriteWarningMessage(outputFile & " already exists! It will be overwritten.")
	                        PauseApplication()
	                        File.Delete(outputFile)
	                    End If
	
	                    totalLines = 0
	                    Try
	                        totalLines = File.ReadAllLines(fileName).Length
	                    Catch ex As Exception
	                        WriteErrorMessage("Exception error happened while counting number of lines in appmasx. Aborting processing completely.")
	                        WriteErrorMessage("EXCEPTION MESSAGE: " & ex.Message)
	                        PauseApplication(True)
	                    End Try
	
	                    Console.Title = "[" & fileInfo.Name & "] " & appTitle
	                    WriteMessage("Working on " & fileInfo.Name & "...")
	                    WriteMessage("For progress, check title bar")
	                    Using reader As New FileIO.TextFieldParser(fileName)
	                        reader.TextFieldType = FileIO.FieldType.Delimited
	                        reader.SetDelimiters(",")
	                        While Not reader.EndOfData
	                            Try
	                                Dim percent As Integer = (reader.LineNumber / totalLines) * 100
	                                Console.Title = "[" & percent & "% - " & reader.LineNumber & "/" & totalLines & "] [" & fileInfo.Name & "] " & appTitle
	                                Dim outputRow As String = ""
	                                Dim currentRow As String() = reader.ReadFields()
	
	                                Dim currentSSN As Integer = Int(currentRow(2))
	
	                                If currentRow.Length <> 71 Then
	                                    WriteWarningMessage("Line " & reader.LineNumber & ": Expected 71, found " & currentRow.Length)
	                                Else
	                                    ' We have correct data, we can then manipulate it and send it to the save location
	
	                                    ' We can skip entries with APNAME of *OPEN JOB ORDERS, *CANCELLED JOB ORDERS, *SPECIAL INVOICING
	                                    If currentRow(3).StartsWith("*") Then GoTo SkipRowEntryAPPMAS
	
	                                    ' idx 8, 62, 63 (APUPDTE, APDA2K, APUD2K)
	                                    ' If person has not been active since 2009
	                                    If removedSSN.Contains(currentRow(2)) Then
	                                        currentRow = Nothing
	                                        GoTo SkipRowEntryAPPMAS
	                                    End If
	
	                                    ' idx 3 contains the persons name (APNAME)
	                                    While currentRow(3).Contains("  ")
	                                        currentRow(3) = currentRow(3).Replace("  ", " ")
	                                    End While
	                                    While currentRow(3).Contains(".")
	                                        currentRow(3) = currentRow(3).Replace(".", ",")
	                                    End While
	                                    If Len(currentRow(3)) - Len(currentRow(3).Replace(",", "")) > 1 Then
	                                        ' we have multiple commas in name, strip off all but the first
	                                        currentRow(3) = currentRow(3).Substring(0, currentRow(3).IndexOf(",")) & currentRow(3).Substring(currentRow(3).IndexOf(",") + 1).Replace(",", "")
	                                    End If
	
	                                    ' idx 15 (APAPCD) and idx 16 (APPFCD) - eval #1, #2
	                                    If Int(currentRow(15)) = 0 Then currentRow(15) = 1
	                                    If Int(currentRow(16)) = 0 Then currentRow(16) = 1
	
	                                    ' Clear skill codes
	                                    For i As Integer = 37 To 55
	                                        currentRow(i) = ""
	                                    Next
	
SkipRowEntryAPPMAS:
	                                    If Not currentRow Is Nothing Then
	                                        outputRow = String.Join(",", currentRow) & Environment.NewLine
	                                        File.AppendAllText(outputFile, outputRow, System.Text.Encoding.ASCII)
	                                    End If
	                                End If
	                            Catch ex As Exception
	                                WriteErrorMessage("Exception error happened on line " & reader.LineNumber & ". Aborting processing completely.")
	                                WriteErrorMessage("EXCEPTION MESSAGE: " & ex.Message)
	                                PauseApplication(True)
	                            End Try
	                        End While
	                    End Using
	
	                    WriteMessage("Finished " & fileInfo.Name & " modifications")

	                ElseIf fileInfo.Name = "appmasx.csv" Then
	                    WriteMessage("Found " & fileInfo.Name & "...")

                        If Not sSaveLocation = sFolderLocation Then
                            outputFile = sSaveLocation & "\" & fileInfo.Name
                        Else
                            outputFile = sSaveLocation & "\appmasx_mod.csv"
                        End If

                        Console.Title = "[" & fileInfo.Name & "] " & appTitle
                        WriteMessage("Working on " & fileInfo.Name & "...")
                        WriteMessage("For progress, check title bar")

                        If File.Exists(outputFile) Then
                            WriteWarningMessage(outputFile & " already exists! It will be overwritten.")
                            PauseApplication()
                            File.Delete(outputFile)
                        End If

                        totalLines = 0
                        Try
                            totalLines = File.ReadAllLines(fileName).Length
                        Catch ex As Exception
                            WriteErrorMessage("Exception error happened while counting number of lines in appmasx. Aborting processing completely.")
                            WriteErrorMessage("EXCEPTION MESSAGE: " & ex.Message)
                            PauseApplication(True)
                        End Try

                        Using reader As New FileIO.TextFieldParser(fileName)
                            Try
                                reader.TextFieldType = FileIO.FieldType.Delimited
                                reader.SetDelimiters(",")
                                While Not reader.EndOfData
                                    Dim percent As Integer = (reader.LineNumber / totalLines) * 100
                                    Console.Title = "[" & percent & "% - " & reader.LineNumber & "/" & totalLines & "] [" & fileInfo.Name & ".csv] " & appTitle
                                    Dim outputRow As String = ""
                                    Dim currentRow As String() = reader.ReadFields()
                                    If currentRow.Length <> 31 Then
                                        WriteWarningMessage("Line " & reader.LineNumber & ": Expected 31, found " & currentRow.Length)
                                    Else
                                        ' We have correct data, we can then manipulate it and send it to the save location

                                        If removedSSN.Contains(Int(currentRow(2))) Then
                                            currentRow = Nothing
                                        End If

SkipRowEntryAPPMASX:
                                        If Not currentRow Is Nothing Then
                                            outputRow = String.Join(",", currentRow) & Environment.NewLine
                                            File.AppendAllText(outputFile, outputRow, System.Text.Encoding.ASCII)
                                        End If
                                    End If
                                End While
                            Catch ex As Exception
                                WriteErrorMessage("Exception error happened on line " & reader.LineNumber & ". Aborting processing completely.")
                                WriteErrorMessage("EXCEPTION MESSAGE: " & ex.Message)
                                PauseApplication(True)
                            End Try
                        End Using
	
						WriteMessage("Finished " & fileInfo.Name & " modifications")
	                ElseIf fileInfo.Name = "bonus.csv" Then
	                	WriteMessage("Found " & fileInfo.Name & "...")
	                	If Not sSaveLocation = sFolderLocation Then
	                		outputFile = sSaveLocation & "\" & fileInfo.Name
	                	Else
	                		outputFile = sSaveLocation & "\bonus_mod.csv"
	                	End If
	                	
	                	If File.Exists(outputFile) Then
	                		WriteWarningMessage(outputFile & " already exists! It will be overwritten.")
	                		PauseApplication()
	                		File.Delete(outputFile)
	                	End If
	                	
	                	totalLines = File.ReadAllLines(fileName).Length
	                	
	                	Console.Title = "[" & fileInfo.Name & "] " & appTitle
	                	WriteMessage("Working on " & fileInfo.Name & "...")
	                	WriteMessage("For progress, check title bar")
	                	Using reader As New FileIO.TextFieldParser(fileName)
	                		reader.TextFieldType = FileIO.FieldType.Delimited
	                		reader.SetDelimiters(",")
	                		While Not reader.EndOfData
	                			Try
	                				Dim percent As Integer = (reader.LineNumber / totalLines) * 100
	                				Console.Title = "[" & percent & "% - " & reader.LineNumber & "/" & totalLines & "] [" & fileInfo.Name & "] " & appTitle
	                				Dim outputRow As String = ""
	                				Dim currentRow As String() = reader.ReadFields()
	                				
	                				If removedSSN.Contains(currentRow(2)) Then
	                					currentRow = Nothing
	                					GoTo SkipRowEntryBonus
	                				End If
	                				' Else we can edit info
	                				
SkipRowEntryBONUS:
	                                If Not currentRow Is Nothing Then
	                                    outputRow = String.Join(",", currentRow) & Environment.NewLine
	                                    File.AppendAllText(outputFile, outputRow, System.Text.Encoding.ASCII)
	                                End If
	                			Catch ex As Exception
	                				WriteErrorMessage("Exception error happened on line " & reader.LineNumber & ". Aborting processing completely.")
	                                WriteErrorMessage("EXCEPTION MESSAGE: " & ex.Message)
	                                PauseApplication(True)
	                			End Try
	                		End While
	                	End Using
	                ElseIf fileInfo.Name = "gldist.csv" Then
	                    WriteMessage("Found gldist.csv...")
	                    If Not sSaveLocation = sFolderLocation Then
	                        outputFile = sSaveLocation & "\" & fileInfo.Name
	                    Else
	                        outputFile = sSaveLocation & "\gldist_mod.csv"
	                    End If
	
	                    If File.Exists(outputFile) Then
	                        WriteWarningMessage(outputFile & " already exists! It will be overwritten.")
	                        PauseApplication()
	                        File.Delete(outputFile)
	                    End If
	
	                    totalLines = File.ReadAllLines(fileName).Length
	
	                    Console.Title = "[" & fileInfo.Name & "] " & appTitle
	                    WriteMessage("Working on " & fileInfo.Name & "...")
	                    WriteMessage("For progress, check title bar")
	                    Using reader As New FileIO.TextFieldParser(fileName)
	                        reader.TextFieldType = FileIO.FieldType.Delimited
	                        reader.SetDelimiters(",")
	                        While Not reader.EndOfData
	                            Try
	                                Dim percent As Integer = (reader.LineNumber / totalLines) * 100
	                                Console.Title = "[" & percent & "% - " & reader.LineNumber & "/" & totalLines & "] [" & fileInfo.Name & "] " & appTitle
	                                Dim outputRow As String = ""
	                                Dim currentRow As String() = reader.ReadFields()
	
	                                ' Strip all entries based on GLCONO and GLTR2K
	                                If currentRow(1) <> "10" Or Int(currentRow(21).Substring(0, 4)) < 2009 Then
	                                    currentRow = Nothing
	                                    GoTo SkipRowEntryGLDIST
	                                End If
	
SkipRowEntryGLDIST:
	                                If Not currentRow Is Nothing Then
	                                    outputRow = String.Join(",", currentRow) & Environment.NewLine
	                                    File.AppendAllText(outputFile, outputRow, System.Text.Encoding.ASCII)
	                                Else
	
	                                End If
	                            Catch ex As Exception
	                                WriteErrorMessage("Exception error happened on line " & reader.LineNumber & ". Aborting processing completely.")
	                                WriteErrorMessage("EXCEPTION MESSAGE: " & ex.Message)
	                                PauseApplication(True)
	                            End Try
	                        End While
	                    End Using
	
	                    WriteMessage("Finished " & fileInfo.Name & " modifications")

	                ElseIf fileInfo.Name = "aaritem1.csv" Then
	                	WriteMessage("Found aaritem1.csv...")
	                    If Not sSaveLocation = sFolderLocation Then
	                        outputFile = sSaveLocation & "\aaritem1.csv"
	                    Else
	                        outputFile = sSaveLocation & "\aaritem1_mod.csv"
	                    End If
	
	                    If File.Exists(outputFile) Then
	                        WriteWarningMessage(outputFile & " already exists! It will be overwritten.")
	                        PauseApplication()
	                        File.Delete(outputFile)
	                    End If
	
	                    totalLines = File.ReadAllLines(fileName).Length
	
	                    Console.Title = "[" & fileInfo.Name & "] " & appTitle
	                    WriteMessage("Working on " & fileInfo.Name & "...")
	                    WriteMessage("For progress, check title bar")
	                    Using reader As New FileIO.TextFieldParser(fileName)
	                        reader.TextFieldType = FileIO.FieldType.Delimited
	                        reader.SetDelimiters(",")
	                        While Not reader.EndOfData
	                            Try
	                                Dim percent As Integer = (reader.LineNumber / totalLines) * 100
	                                Console.Title = "[" & percent & "% - " & reader.LineNumber & "/" & totalLines & "] [" & fileInfo.Name & "] " & appTitle
	                                Dim outputRow As String = ""
	                                Dim currentRow As String() = reader.ReadFields()
	
	                                ' Strip all entries based on customer code
	                                If removedCCN.Contains(currentRow(2)) Then
	                                	' Removed entry because parent customer was deleted
	                                	currentRow = Nothing
	                                	GoTo SkipRowEntryAARITEM1
	                                End If
									
SkipRowEntryAARITEM1:
	                                If Not currentRow Is Nothing Then
	                                    outputRow = String.Join(",", currentRow) & Environment.NewLine
	                                    File.AppendAllText(outputFile, outputRow, System.Text.Encoding.ASCII)
	                                End If
	                            Catch ex As Exception
	                                WriteErrorMessage("Exception error happened on line " & reader.LineNumber & ". Aborting processing completely.")
	                                WriteErrorMessage("EXCEPTION MESSAGE: " & ex.Message)
	                                PauseApplication(True)
	                            End Try
	                        End While
	                    End Using
	
						WriteMessage("Finished " & fileInfo.Name & " modifications")
					ElseIf fileInfo.Name = "cshrec.csv" Then
	                	WriteMessage("Found cshrec.csv...")
	                    If Not sSaveLocation = sFolderLocation Then
	                        outputFile = sSaveLocation & "\cshrec.csv"
	                    Else
	                        outputFile = sSaveLocation & "\cshrec_mod.csv"
	                    End If
	
	                    If File.Exists(outputFile) Then
	                        WriteWarningMessage(outputFile & " already exists! It will be overwritten.")
	                        PauseApplication()
	                        File.Delete(outputFile)
	                    End If
	
	                    totalLines = File.ReadAllLines(fileName).Length
	
	                    Console.Title = "[" & fileInfo.Name & "] " & appTitle
	                    WriteMessage("Working on " & fileInfo.Name & "...")
	                    WriteMessage("For progress, check title bar")
	                    Using reader As New FileIO.TextFieldParser(fileName)
	                        reader.TextFieldType = FileIO.FieldType.Delimited
	                        reader.SetDelimiters(",")
	                        While Not reader.EndOfData
	                            Try
	                                Dim percent As Integer = (reader.LineNumber / totalLines) * 100
	                                Console.Title = "[" & percent & "% - " & reader.LineNumber & "/" & totalLines & "] [" & fileInfo.Name & "] " & appTitle
	                                Dim outputRow As String = ""
	                                Dim currentRow As String() = reader.ReadFields()
	
	                                ' Strip all entries based on customer code
	                                If removedCCN.Contains(currentRow(2)) Then
	                                	' Removed entry because parent customer was deleted
	                                	currentRow = Nothing
	                                	GoTo SkipRowEntryCSHREC
	                                End If
									
SkipRowEntryCSHREC:
	                                If Not currentRow Is Nothing Then
	                                    outputRow = String.Join(",", currentRow) & Environment.NewLine
	                                    File.AppendAllText(outputFile, outputRow, System.Text.Encoding.ASCII)
	                                End If
	                            Catch ex As Exception
	                                WriteErrorMessage("Exception error happened on line " & reader.LineNumber & ". Aborting processing completely.")
	                                WriteErrorMessage("EXCEPTION MESSAGE: " & ex.Message)
	                                PauseApplication(True)
	                            End Try
	                        End While
	                    End Using
	
						WriteMessage("Finished " & fileInfo.Name & " modifications")
					ElseIf fileInfo.Name = "cuscon.csv" Then
	                	WriteMessage("Found cuscon.csv...")
	                    If Not sSaveLocation = sFolderLocation Then
	                        outputFile = sSaveLocation & "\" & fileInfo.Name
	                    Else
	                        outputFile = sSaveLocation & "\cuscon_mod.csv"
	                    End If
	
	                    If File.Exists(outputFile) Then
	                        WriteWarningMessage(outputFile & " already exists! It will be overwritten.")
	                        PauseApplication()
	                        File.Delete(outputFile)
	                    End If
	
	                    totalLines = File.ReadAllLines(fileName).Length
	
	                    Console.Title = "[" & fileInfo.Name & "] " & appTitle
	                    WriteMessage("Working on " & fileInfo.Name & "...")
	                    WriteMessage("For progress, check title bar")
	                    Using reader As New FileIO.TextFieldParser(fileName)
	                        reader.TextFieldType = FileIO.FieldType.Delimited
	                        reader.SetDelimiters(",")
	                        While Not reader.EndOfData
	                            Try
	                                Dim percent As Integer = (reader.LineNumber / totalLines) * 100
	                                Console.Title = "[" & percent & "% - " & reader.LineNumber & "/" & totalLines & "] [" & fileInfo.Name & "] " & appTitle
	                                Dim outputRow As String = ""
	                                Dim currentRow As String() = reader.ReadFields()
	
	                                ' Strip all entries based on customer code
	                                If removedCCN.Contains(currentRow(2)) Then
	                                	' Removed entry because parent customer was deleted
	                                	currentRow = Nothing
	                                	GoTo SkipRowEntryCUSCON
	                                End If
									
SkipRowEntryCUSCON:
	                                If Not currentRow Is Nothing Then
	                                    outputRow = String.Join(",", currentRow) & Environment.NewLine
	                                    File.AppendAllText(outputFile, outputRow, System.Text.Encoding.ASCII)
	                                End If
	                            Catch ex As Exception
	                                WriteErrorMessage("Exception error happened on line " & reader.LineNumber & ". Aborting processing completely.")
	                                WriteErrorMessage("EXCEPTION MESSAGE: " & ex.Message)
	                                PauseApplication(True)
	                            End Try
	                        End While
	                    End Using
	
						WriteMessage("Finished " & fileInfo.Name & " modifications")
					ElseIf fileInfo.Name = "cuswcc.csv" Then
	                	WriteMessage("Found cuswcc.csv...")
	                    If Not sSaveLocation = sFolderLocation Then
	                        outputFile = sSaveLocation & "\cuswcc.csv"
	                    Else
	                        outputFile = sSaveLocation & "\cuswcc_mod.csv"
	                    End If
	
	                    If File.Exists(outputFile) Then
	                        WriteWarningMessage(outputFile & " already exists! It will be overwritten.")
	                        PauseApplication()
	                        File.Delete(outputFile)
	                    End If
	
	                    totalLines = File.ReadAllLines(fileName).Length
	
	                    Console.Title = "[cuswcc.csv] " & appTitle
	                    WriteMessage("Working on cuswcc.csv...")
	                    WriteMessage("For progress, check title bar")
	                    Using reader As New FileIO.TextFieldParser(fileName)
	                        reader.TextFieldType = FileIO.FieldType.Delimited
	                        reader.SetDelimiters(",")
	                        While Not reader.EndOfData
	                            Try
	                                Dim percent As Integer = (reader.LineNumber / totalLines) * 100
	                                Console.Title = "[" & percent & "% - " & reader.LineNumber & "/" & totalLines & "] [" & fileInfo.Name & "] " & appTitle
	                                Dim outputRow As String = ""
	                                Dim currentRow As String() = reader.ReadFields()
	
	                                ' Strip all entries based on customer code
	                                If removedCCN.Contains(currentRow(2)) Then
	                                	' Removed entry because parent customer was deleted
	                                	currentRow = Nothing
	                                	GoTo SkipRowEntryCUSWCC
	                                End If
									
SkipRowEntryCUSWCC:
	                                If Not currentRow Is Nothing Then
	                                    outputRow = String.Join(",", currentRow) & Environment.NewLine
	                                    File.AppendAllText(outputFile, outputRow, System.Text.Encoding.ASCII)
	                                End If
	                            Catch ex As Exception
	                                WriteErrorMessage("Exception error happened on line " & reader.LineNumber & ". Aborting processing completely.")
	                                WriteErrorMessage("EXCEPTION MESSAGE: " & ex.Message)
	                                PauseApplication(True)
	                            End Try
	                        End While
	                    End Using
	
						WriteMessage("Finished " & fileInfo.Name & " modifications")
					ElseIf fileInfo.Name = "deduct.csv" Then
						WriteMessage("Found " & fileInfo.Name & "...")
	                    If Not sSaveLocation = sFolderLocation Then
	                        outputFile = sSaveLocation & "\" & fileInfo.Name
	                    Else
	                        outputFile = sSaveLocation & "\deduct_mod.csv"
	                    End If
	
	                    If File.Exists(outputFile) Then
	                        WriteWarningMessage(outputFile & " already exists! It will be overwritten.")
	                        PauseApplication()
	                        File.Delete(outputFile)
	                    End If
	
	                    totalLines = File.ReadAllLines(fileName).Length
	
	                    Console.Title = "[" & fileInfo.Name & "] " & appTitle
	                    WriteMessage("Working on " & fileInfo.Name & "...")
	                    WriteMessage("For progress, check title bar")
	                    Using reader As New FileIO.TextFieldParser(fileName)
	                        reader.TextFieldType = FileIO.FieldType.Delimited
	                        reader.SetDelimiters(",")
	                        While Not reader.EndOfData
	                        	Try
	                        		Dim percent As Integer = (reader.LineNumber / totalLines) * 100
	                                Console.Title = "[" & percent & "% - " & reader.LineNumber & "/" & totalLines & "] [" & fileInfo.Name & "] " & appTitle
	                                Dim outputRow As String = ""
	                                Dim currentRow As String() = reader.ReadFields()
	                                
	                                ' Strip all entries based on company code, MUST BE 10
	                                '	AND SS# shouldnt be in removed list
	                                If Int(currentRow(1)) <> 10 Or removedSSN.Contains(Int(currentRow(2))) Then
	                                	currentRow = Nothing
	                                	GoTo SkipRowEntryDEDUCT
	                                End If
	                                ' Else we can modify data
	                                
SkipRowEntryDEDUCT:
	                                If Not currentRow Is Nothing Then
	                                    outputRow = String.Join(",", currentRow) & Environment.NewLine
	                                    File.AppendAllText(outputFile, outputRow, System.Text.Encoding.ASCII)
	                                End If
	                            Catch ex As Exception
	                                WriteErrorMessage("Exception error happened on line " & reader.LineNumber & ". Aborting processing completely.")
	                                WriteErrorMessage("EXCEPTION MESSAGE: " & ex.Message)
	                                PauseApplication(True)
	                            End Try
	                        End While
	                    End Using
						
						WriteMessage("Finished " & fileInfo.Name & " modifications")
					ElseIf fileInfo.Name = "deductap.csv" Then
						WriteMessage("Found " & fileInfo.Name & "...")
	                    If Not sSaveLocation = sFolderLocation Then
	                        outputFile = sSaveLocation & "\" & fileInfo.Name
	                    Else
	                        outputFile = sSaveLocation & "\deductap_mod.csv"
	                    End If
	
	                    If File.Exists(outputFile) Then
	                        WriteWarningMessage(outputFile & " already exists! It will be overwritten.")
	                        PauseApplication()
	                        File.Delete(outputFile)
	                    End If
	
	                    totalLines = File.ReadAllLines(fileName).Length
	
	                    Console.Title = "[" & fileInfo.Name & "] " & appTitle
	                    WriteMessage("Working on " & fileInfo.Name & "...")
	                    WriteMessage("For progress, check title bar")
	                    Using reader As New FileIO.TextFieldParser(fileName)
	                        reader.TextFieldType = FileIO.FieldType.Delimited
	                        reader.SetDelimiters(",")
	                        While Not reader.EndOfData
	                        	Try
	                        		Dim percent As Integer = (reader.LineNumber / totalLines) * 100
	                                Console.Title = "[" & percent & "% - " & reader.LineNumber & "/" & totalLines & "] [" & fileInfo.Name & "] " & appTitle
	                                Dim outputRow As String = ""
	                                Dim currentRow As String() = reader.ReadFields()
	                                
	                                ' Strip all entries based on company code, MUST BE 10
	                                If Int(currentRow(1)) <> 10 Or removedSSN.Contains(Int(currentRow(2))) Then
	                                	currentRow = Nothing
	                                	GoTo SkipRowEntryDEDUCTAP
	                                End If
	                                ' Else we can modify data
	                                
SkipRowEntryDEDUCTAP:
	                                If Not currentRow Is Nothing Then
	                                    outputRow = String.Join(",", currentRow) & Environment.NewLine
	                                    File.AppendAllText(outputFile, outputRow, System.Text.Encoding.ASCII)
	                                End If
	                            Catch ex As Exception
	                                WriteErrorMessage("Exception error happened on line " & reader.LineNumber & ". Aborting processing completely.")
	                                WriteErrorMessage("EXCEPTION MESSAGE: " & ex.Message)
	                                PauseApplication(True)
	                            End Try
	                        End While
	                    End Using
						
						WriteMessage("Finished " & fileInfo.Name & " modifications")
					ElseIf fileInfo.Name = "lpaybil1.csv" Then
						WriteMessage("Found " & fileInfo.Name & "...")
	                    If Not sSaveLocation = sFolderLocation Then
	                        outputFile = sSaveLocation & "\" & fileInfo.Name
	                    Else
	                        outputFile = sSaveLocation & "\lpaybil1_mod.csv"
	                    End If
	
	                    If File.Exists(outputFile) Then
	                        WriteWarningMessage(outputFile & " already exists! It will be overwritten.")
	                        PauseApplication()
	                        File.Delete(outputFile)
	                    End If
	
	                    totalLines = File.ReadAllLines(fileName).Length
	
	                    Console.Title = "[" & fileInfo.Name & "] " & appTitle
	                    WriteMessage("Working on " & fileInfo.Name & "...")
	                    WriteMessage("For progress, check title bar")
	                    Using reader As New FileIO.TextFieldParser(fileName)
	                        reader.TextFieldType = FileIO.FieldType.Delimited
	                        reader.SetDelimiters(",")
	                        While Not reader.EndOfData
	                        	Try
	                        		Dim percent As Integer = (reader.LineNumber / totalLines) * 100
	                                Console.Title = "[" & percent & "% - " & reader.LineNumber & "/" & totalLines & "] [" & fileInfo.Name & "] " & appTitle
	                                Dim outputRow As String = ""
	                                Dim currentRow As String() = reader.ReadFields()
	                                
	                                ' Strip all entries based on ssn that were removed
	                                If removedSSN.Contains(Int(currentRow(2))) Then
	                                	currentRow = Nothing
	                                	GoTo SkipRowEntryLPAYBIL1
	                                End If
	                                ' Else we can modify data
	                                
SkipRowEntryLPAYBIL1:
	                                If Not currentRow Is Nothing Then
	                                    outputRow = String.Join(",", currentRow) & Environment.NewLine
	                                    File.AppendAllText(outputFile, outputRow, System.Text.Encoding.ASCII)
	                                End If
	                            Catch ex As Exception
	                                WriteErrorMessage("Exception error happened on line " & reader.LineNumber & ". Aborting processing completely.")
	                                WriteErrorMessage("EXCEPTION MESSAGE: " & ex.Message)
	                                PauseApplication(True)
	                            End Try
	                        End While
	                    End Using
						
						WriteMessage("Finished " & fileInfo.Name & " modifications")
					ElseIf fileInfo.Name = "ljobdscd.csv" Then
						WriteMessage("Found " & fileInfo.Name & "...")
	                    If Not sSaveLocation = sFolderLocation Then
	                        outputFile = sSaveLocation & "\" & fileInfo.Name
	                    Else
	                        outputFile = sSaveLocation & "\ljobdscd_mod.csv"
	                    End If
	
	                    If File.Exists(outputFile) Then
	                        WriteWarningMessage(outputFile & " already exists! It will be overwritten.")
	                        PauseApplication()
	                        File.Delete(outputFile)
	                    End If
	
	                    totalLines = File.ReadAllLines(fileName).Length
	
	                    Console.Title = "[" & fileInfo.Name & "] " & appTitle
	                    WriteMessage("Working on " & fileInfo.Name & "...")
	                    WriteMessage("For progress, check title bar")
	                    Using reader As New FileIO.TextFieldParser(fileName)
	                        reader.TextFieldType = FileIO.FieldType.Delimited
	                        reader.SetDelimiters(",")
	                        While Not reader.EndOfData
	                        	Try
	                        		Dim percent As Integer = (reader.LineNumber / totalLines) * 100
	                                Console.Title = "[" & percent & "% - " & reader.LineNumber & "/" & totalLines & "] [" & fileInfo.Name & "] " & appTitle
	                                Dim outputRow As String = ""
	                                Dim currentRow As String() = reader.ReadFields()
	                                
	                                ' Strip all entries based on ssn that were removed
	                                If removedJON.Contains(Int(currentRow(3))) Then
	                                	currentRow = Nothing
	                                	GoTo SkipRowEntryLJOBDSCD
	                                End If
	                                ' Else we can modify data
	                                
SkipRowEntryLJOBDSCD:
	                                If Not currentRow Is Nothing Then
	                                    outputRow = String.Join(",", currentRow) & Environment.NewLine
	                                    File.AppendAllText(outputFile, outputRow, System.Text.Encoding.ASCII)
	                                End If
	                            Catch ex As Exception
	                                WriteErrorMessage("Exception error happened on line " & reader.LineNumber & ". Aborting processing completely.")
	                                WriteErrorMessage("EXCEPTION MESSAGE: " & ex.Message)
	                                PauseApplication(True)
	                            End Try
	                        End While
	                    End Using
						
						WriteMessage("Finished " & fileInfo.Name & " modifications")
					ElseIf fileInfo.Name = "ljobdscp.csv" Then
						WriteMessage("Found " & fileInfo.Name & "...")
	                    If Not sSaveLocation = sFolderLocation Then
	                        outputFile = sSaveLocation & "\" & fileInfo.Name
	                    Else
	                        outputFile = sSaveLocation & "\ljobdscp_mod.csv"
	                    End If
	
	                    If File.Exists(outputFile) Then
	                        WriteWarningMessage(outputFile & " already exists! It will be overwritten.")
	                        PauseApplication()
	                        File.Delete(outputFile)
	                    End If
	
	                    totalLines = File.ReadAllLines(fileName).Length
	
	                    Console.Title = "[" & fileInfo.Name & "] " & appTitle
	                    WriteMessage("Working on " & fileInfo.Name & "...")
	                    WriteMessage("For progress, check title bar")
	                    Using reader As New FileIO.TextFieldParser(fileName)
	                        reader.TextFieldType = FileIO.FieldType.Delimited
	                        reader.SetDelimiters(",")
	                        While Not reader.EndOfData
	                        	Try
	                        		Dim percent As Integer = (reader.LineNumber / totalLines) * 100
	                                Console.Title = "[" & percent & "% - " & reader.LineNumber & "/" & totalLines & "] [" & fileInfo.Name & "] " & appTitle
	                                Dim outputRow As String = ""
	                                Dim currentRow As String() = reader.ReadFields()
	                                
	                                ' Strip all entries based on ssn that were removed
	                                If removedJON.Contains(Int(currentRow(3))) Then
	                                	currentRow = Nothing
	                                	GoTo SkipRowEntryLJOBDSCP
	                                End If
	                                ' Else we can modify data
	                                
SkipRowEntryLJOBDSCP:
	                                If Not currentRow Is Nothing Then
	                                    outputRow = String.Join(",", currentRow) & Environment.NewLine
	                                    File.AppendAllText(outputFile, outputRow, System.Text.Encoding.ASCII)
	                                End If
	                            Catch ex As Exception
	                                WriteErrorMessage("Exception error happened on line " & reader.LineNumber & ". Aborting processing completely.")
	                                WriteErrorMessage("EXCEPTION MESSAGE: " & ex.Message)
	                                PauseApplication(True)
	                            End Try
	                        End While
	                    End Using
						
						WriteMessage("Finished " & fileInfo.Name & " modifications")
					ElseIf fileInfo.Name = "lstlctxs.csv" Then
						WriteMessage("Found " & fileInfo.Name & "...")
	                    If Not sSaveLocation = sFolderLocation Then
	                        outputFile = sSaveLocation & "\" & fileInfo.Name
	                    Else
	                        outputFile = sSaveLocation & "\lstlctxs_mod.csv"
	                    End If
	
	                    If File.Exists(outputFile) Then
	                        WriteWarningMessage(outputFile & " already exists! It will be overwritten.")
	                        PauseApplication()
	                        File.Delete(outputFile)
	                    End If
	
	                    totalLines = File.ReadAllLines(fileName).Length
	
	                    Console.Title = "[" & fileInfo.Name & "] " & appTitle
	                    WriteMessage("Working on " & fileInfo.Name & "...")
	                    WriteMessage("For progress, check title bar")
	                    Using reader As New FileIO.TextFieldParser(fileName)
	                        reader.TextFieldType = FileIO.FieldType.Delimited
	                        reader.SetDelimiters(",")
	                        While Not reader.EndOfData
	                        	Try
	                        		Dim percent As Integer = (reader.LineNumber / totalLines) * 100
	                                Console.Title = "[" & percent & "% - " & reader.LineNumber & "/" & totalLines & "] [" & fileInfo.Name & "] " & appTitle
	                                Dim outputRow As String = ""
	                                Dim currentRow As String() = reader.ReadFields()
	                                
	                                ' Strip all entries based on ssn that were removed
	                                If removedSSN.Contains(Int(currentRow(2))) Then
	                                	currentRow = Nothing
	                                	GoTo SkipRowEntryLSTLCTXS
	                                End If
	                                ' Else we can modify data
	                                
SkipRowEntryLSTLCTXS:
	                                If Not currentRow Is Nothing Then
	                                    outputRow = String.Join(",", currentRow) & Environment.NewLine
	                                    File.AppendAllText(outputFile, outputRow, System.Text.Encoding.ASCII)
	                                End If
	                            Catch ex As Exception
	                                WriteErrorMessage("Exception error happened on line " & reader.LineNumber & ". Aborting processing completely.")
	                                WriteErrorMessage("EXCEPTION MESSAGE: " & ex.Message)
	                                PauseApplication(True)
	                            End Try
	                        End While
	                    End Using
						
						WriteMessage("Finished " & fileInfo.Name & " modifications")
					ElseIf fileInfo.Name = "pbgrs$.csv" Then
						WriteMessage("Found " & fileInfo.Name & "...")
	                    If Not sSaveLocation = sFolderLocation Then
	                        outputFile = sSaveLocation & "\" & fileInfo.Name
	                    Else
	                        outputFile = sSaveLocation & "\pbgrs$_mod.csv"
	                    End If
	
	                    If File.Exists(outputFile) Then
	                        WriteWarningMessage(outputFile & " already exists! It will be overwritten.")
	                        PauseApplication()
	                        File.Delete(outputFile)
	                    End If
	
	                    totalLines = File.ReadAllLines(fileName).Length
	
	                    Console.Title = "[" & fileInfo.Name & "] " & appTitle
	                    WriteMessage("Working on " & fileInfo.Name & "...")
	                    WriteMessage("For progress, check title bar")
	                    Using reader As New FileIO.TextFieldParser(fileName)
	                        reader.TextFieldType = FileIO.FieldType.Delimited
	                        reader.SetDelimiters(",")
	                        While Not reader.EndOfData
	                        	Try
	                        		Dim percent As Integer = (reader.LineNumber / totalLines) * 100
	                                Console.Title = "[" & percent & "% - " & reader.LineNumber & "/" & totalLines & "] [" & fileInfo.Name & "] " & appTitle
	                                Dim outputRow As String = ""
	                                Dim currentRow As String() = reader.ReadFields()
	                                
	                                ' Strip all entries based on ssn that were removed
	                                If removedSSN.Contains(Int(currentRow(2))) Then
	                                	currentRow = Nothing
	                                	GoTo SkipRowEntryPBGRS
	                                End If
	                                ' Else we can modify data
	                                
SkipRowEntryPBGRS:
	                                If Not currentRow Is Nothing Then
	                                    outputRow = String.Join(",", currentRow) & Environment.NewLine
	                                    File.AppendAllText(outputFile, outputRow, System.Text.Encoding.ASCII)
	                                End If
	                            Catch ex As Exception
	                                WriteErrorMessage("Exception error happened on line " & reader.LineNumber & ". Aborting processing completely.")
	                                WriteErrorMessage("EXCEPTION MESSAGE: " & ex.Message)
	                                PauseApplication(True)
	                            End Try
	                        End While
	                    End Using
						
						WriteMessage("Finished " & fileInfo.Name & " modifications")
					ElseIf fileInfo.Name = "pbmded.csv" Then
						WriteMessage("Found " & fileInfo.Name & "...")
	                    If Not sSaveLocation = sFolderLocation Then
	                        outputFile = sSaveLocation & "\" & fileInfo.Name
	                    Else
	                        outputFile = sSaveLocation & "\pbmded_mod.csv"
	                    End If
	
	                    If File.Exists(outputFile) Then
	                        WriteWarningMessage(outputFile & " already exists! It will be overwritten.")
	                        PauseApplication()
	                        File.Delete(outputFile)
	                    End If
	
	                    totalLines = File.ReadAllLines(fileName).Length
	
	                    Console.Title = "[" & fileInfo.Name & "] " & appTitle
	                    WriteMessage("Working on " & fileInfo.Name & "...")
	                    WriteMessage("For progress, check title bar")
	                    Using reader As New FileIO.TextFieldParser(fileName)
	                        reader.TextFieldType = FileIO.FieldType.Delimited
	                        reader.SetDelimiters(",")
	                        While Not reader.EndOfData
	                        	Try
                                    Dim percent As Integer = (reader.LineNumber / totalLines) * 100
                                    UpdateConsoleTitle(fileInfo.Name, percent, reader.LineNumber, totalLines)
	                                Dim outputRow As String = ""
	                                Dim currentRow As String() = reader.ReadFields()
	                                
	                                ' Strip all entries based on ssn that were removed
	                                If removedSSN.Contains(Int(currentRow(2))) Then
	                                	currentRow = Nothing
	                                	GoTo SkipRowEntryPBMDED
	                                End If
	                                ' Else we can modify data
	                                
SkipRowEntryPBMDED:
	                                If Not currentRow Is Nothing Then
	                                    outputRow = String.Join(",", currentRow) & Environment.NewLine
	                                    File.AppendAllText(outputFile, outputRow, System.Text.Encoding.ASCII)
	                                End If
	                            Catch ex As Exception
	                                WriteErrorMessage("Exception error happened on line " & reader.LineNumber & ". Aborting processing completely.")
	                                WriteErrorMessage("EXCEPTION MESSAGE: " & ex.Message)
	                                PauseApplication(True)
	                            End Try
	                        End While
	                    End Using
						
						WriteMessage("Finished " & fileInfo.Name & " modifications")
					ElseIf fileInfo.Name = "venmas.csv" Then
						WriteMessage("Found " & fileInfo.Name & "...")
	                    If Not sSaveLocation = sFolderLocation Then
	                        outputFile = sSaveLocation & "\" & fileInfo.Name
	                    Else
	                        outputFile = sSaveLocation & "\venmas_mod.csv"
	                    End If
	
	                    If File.Exists(outputFile) Then
	                        WriteWarningMessage(outputFile & " already exists! It will be overwritten.")
	                        PauseApplication()
	                        File.Delete(outputFile)
	                    End If
	
	                    totalLines = File.ReadAllLines(fileName).Length
	
	                    Console.Title = "[" & fileInfo.Name & "] " & appTitle
	                    WriteMessage("Working on " & fileInfo.Name & "...")
	                    WriteMessage("For progress, check title bar")
	                    Using reader As New FileIO.TextFieldParser(fileName)
	                        reader.TextFieldType = FileIO.FieldType.Delimited
	                        reader.SetDelimiters(",")
	                        While Not reader.EndOfData
	                        	Try
	                        		Dim percent As Integer = (reader.LineNumber / totalLines) * 100
                                    UpdateConsoleTitle(fileInfo.Name, percent, reader.LineNumber, totalLines)

	                                Dim outputRow As String = ""
	                                Dim currentRow As String() = reader.ReadFields()
	                                
	                                ' Strip all entries based on VMAP2K (idx 20) and VMCD2K (idx 21) < 2009
	                                If Int(currentRow(20).PadLeft(8, "0").Substring(0, 4)) < 2009 And Int(currentRow(21).PadLeft(8, "0").Substring(0,4)) < 2009 Then
	                                	currentRow = Nothing
	                                	GoTo SkipRowEntryVENMAS
	                                End If
	                                ' Else we can modify data
	                                
SkipRowEntryVENMAS:
	                                If Not currentRow Is Nothing Then
	                                    outputRow = String.Join(",", currentRow) & Environment.NewLine
	                                    File.AppendAllText(outputFile, outputRow, System.Text.Encoding.ASCII)
	                                End If
	                            Catch ex As Exception
	                                WriteErrorMessage("Exception error happened on line " & reader.LineNumber & ". Aborting processing completely.")
	                                WriteErrorMessage("EXCEPTION MESSAGE: " & ex.Message)
	                                PauseApplication(True)
	                            End Try
	                        End While
	                    End Using
						
						WriteMessage("Finished " & fileInfo.Name & " modifications")
	                ElseIf ignoreFileList.Contains(fileInfo.Name) Then
	                	WriteNoticeMessage(fileInfo.Name & " is not a required file for conversion.")
	                End If
                Next
        	Else
        		WriteErrorMessage("Cannot convert files! Missing required file empmas.csv")
        		PauseApplication(True)
        	End If
        End If

        Console.ForegroundColor = colorEntryFg
        Console.BackgroundColor = colorEntryBg
        
        Dim runDiff As Long = DateDiff(DateInterval.Second, aStartDate, Now)
                
        WriteMessage("Processing completed in " & runDiff & " seconds.")
        PauseApplication()
    End Sub

    Sub UpdateConsoleTitle(Optional fileName As String = Nothing, Optional percentage As String = Nothing, Optional currentValue As Integer = -1, Optional maxValue As Integer = -1)
        Dim title = appTitle
        If fileName IsNot Nothing Then
            title = "[" & fileName & "] " & title
        End If
        If percentage IsNot Nothing Then
            If currentValue > -1 Then
                If maxValue > -1 Then
                    ' % - cur/max
                    title = "[" & percentage & "% - " & currentValue & "/" & maxValue & "] " & title
                Else
                    ' % - cur
                    title = "[" & percentage & "% - " & currentValue & "] " & title
                End If
            ElseIf maxValue > -1 Then
                ' % - max
                title = "[" & percentage & "% - " & maxValue & "] " & title
            End If
        Else
            If currentValue > -1 Then
                If maxValue > -1 Then
                    ' min/max
                    title = "[" & currentValue & "/" & maxValue & "]" & title
                Else
                    ' min
                    title = "[" & currentValue & "] " & title
                End If
            End If
        End If

        Console.Title = title
    End Sub

    Function GenericFilter(filePath As String, idx As Integer, compareTo As String, compareBy As EnumCompareBy) As String()
        Dim out As List(Of String) = New List(Of String)

        If Not File.Exists(filePath) Then
            WriteErrorMessage("File not found: " & filePath)
            Return Nothing
        End If

        Dim totalLines As Integer = File.ReadAllLines(filePath).Length
        Dim fileInfo As FileInfo = New FileInfo(filePath)

        Using reader As New FileIO.TextFieldParser(filePath)
            reader.TextFieldType = FileIO.FieldType.Delimited
            reader.SetDelimiters(",")
            While Not reader.EndOfData
                Try
                    Dim percent As Integer = (reader.LineNumber / totalLines) * 100
                    UpdateConsoleTitle(FileInfo.Name, percent, reader.LineNumber, totalLines)

                    Dim outputRow As String = ""
                    Dim currentRow As String() = reader.ReadFields()

                    If currentRow(idx) = compareTo Then
                        currentRow = Nothing
                    End If

                    If Not currentRow Is Nothing Then
                        out.Add(String.Join(",", currentRow))
                    End If
                Catch ex As Exception
                    WriteErrorMessage("Exception error happened on line " & reader.LineNumber & ". Aborting processing completely.")
                    WriteErrorMessage("EXCEPTION MESSAGE: " & ex.Message)
                    PauseApplication(True)
                End Try
            End While
        End Using

        Return out.ToArray()
    End Function

    Sub PauseApplication(Optional exitAfter As Boolean = False)
        If sPause Then
            Console.WriteLine("Press enter to continue...")
            Console.Read()
        End If
        If exitAfter Then End 'Else Return
    End Sub

#Region "Command Line Info"

    Function ParseCommandLine() As Boolean
        'step one, Do we have a command line?
        If String.IsNullOrEmpty(Command) Then
            'give up if we don't
            Return False
        End If

        'does the command line have at least one named parameter?
        If Not Command.Contains("/") And Not Command.Contains("-") Then
            'give up if we don't
            Return False
        End If

        Dim Params As String() = Nothing

        Dim cmdLine As String = Command()
        ' Need to strip out command line redirects
        If cmdLine.Replace(">>", ">").Contains(">") Then
            cmdLine.Substring(0, cmdLine.Replace(">>", ">").LastIndexOf(">"))
        End If

        If cmdLine.Contains("-") Then
            Params = cmdLine.Split("-".ToCharArray, System.StringSplitOptions.RemoveEmptyEntries)
        End If

        If Params Is Nothing Then Return False

        'Iterate through the parameters passed
        For Each arg As String In Params
            'only process if the argument is not empty
            If Not String.IsNullOrEmpty(arg) Then
                'and contains an equal 
                If arg.Contains("=") Then

                    Dim tmp As ArgumentPairs
                    'find the equal sign
                    Dim idx As Integer = arg.IndexOf("=")
                    'if the equal isn't at the end of the string
                    If idx < arg.Length - 1 Then
                        'parse the name value pair
                        tmp.Name = arg.Substring(0, idx).Trim()
                        tmp.Value = arg.Substring(idx + 1).Trim()
                        'add it to the list.
                        CommandLineArgs.Add(tmp)
                    End If
                Else
                    Dim tmp As ArgumentPairs
                    tmp.Name = arg
                    tmp.Value = Nothing
                    CommandLineArgs.Add(tmp)
                End If
            End If
        Next
        Return True
    End Function

    Sub HelpPrint()
        Console.WriteLine(My.Application.Info.AssemblyName)
        Console.WriteLine("")
        Console.WriteLine("     -f=path      The folder path to read the .csv data from")
        Console.WriteLine("     -s=path      The folder path to save the modified .csv data to")
        Console.WriteLine("     -silent      Silent mode. Disables all messages. Equivalent to -x=mnwe")
        Console.WriteLine("     -x=[flags]	 Disables certain aspects based on the [flags]")
        Console.WriteLine("          More than one flag can be specified.")
        Console.WriteLine("          Valid flags are:")
        Console.WriteLine("               c			Console Clear")
        Console.WriteLine("               m			Messages")
        Console.WriteLine("               n			Notice Messages")
        Console.WriteLine("               w			Warning Messages")
        Console.WriteLine("               e			Error Messages")
        Console.WriteLine("               p			Pausing")
    End Sub
#End Region

#Region "Data Structure"
    Private Structure ArgumentPairs
        Dim Name As String
        Dim Value As String
    End Structure

    Public Enum EnumCompareBy
        LT
        LTE
        GT
        GTE
        EQU
    End Enum
#End Region

#Region "Console Related Subs"

    Sub WriteMessage(message As String)
        If Not sMessage Then Return
        WriteLineColor(ConsoleColor.White, "[M] " & message)
    End Sub

    Sub WriteErrorMessage(message As String)
        If Not sErrorMessage Then Return
        WriteLineColor(ConsoleColor.Red, "[E] " & message)
    End Sub

    Sub WriteOverwriteMessage(fileName As String)
        WriteWarningMessage(fileName & " already exists! It will be overwritten. To prevent overwriting, close the program now")
    End Sub

    Sub WriteWarningMessage(message As String)
        If Not sWarningMessage Then Return
        WriteLineColor(ConsoleColor.Red, "[W] " & message)
    End Sub

    Sub WriteNoticeMessage(message As String)
        If Not sNoticeMessage Then Return
        WriteLineColor(ConsoleColor.Yellow, "[N] " & message)
    End Sub

    Sub WriteLineColor(foreground As ConsoleColor, message As String)
        SetConsoleColors(foreground, ConsoleColor.Black)
        WriteFormattedLine(message)
    End Sub

    Sub WriteFormattedLine(message As String)
        Console.WriteLine("[" & DateTime.Now.ToString("s") & "] " & message)
        SetConsoleColors(ConsoleColor.White, ConsoleColor.Black)
    End Sub

    Sub SetConsoleColors(foreground As ConsoleColor, background As ConsoleColor)
        Console.ForegroundColor = foreground
        Console.BackgroundColor = background
    End Sub
#End Region

#Region "Test Functions/Subs"

    Sub TestPrint()
        ' Standard message
        WriteMessage("Standard Message")

        ' Error Message
        WriteErrorMessage("Error Message")

        ' Warning Message
        WriteWarningMessage("Warning Message")

        ' Notice Message
        WriteNoticeMessage("Notice Message")

        ' Reset colors to app specific
        SetConsoleColors(ConsoleColor.White, ConsoleColor.Black)
    End Sub

#End Region


End Module

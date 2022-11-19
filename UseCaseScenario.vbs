option explicit

!INC Local Scripts.EAConstants-VBScript

'
' This code has been included from the default Project Browser template.
' If you wish to modify this template, it is located in the Config\Script Templates
' directory of your EA install path.   
'
' Script Name: Use Case Scenario
' Author: Paul Reedman
' Purpose: Extracts the scenario steps from a selected use case and creates a csv for the data
' Date: 18/11/2022
'

'
' Project Browser Script main function
'
sub OnProjectBrowserScript()

	' Global attributes for File IO
    Dim fileSystemObject
    Dim outputFile
	
	
	
	
	' Get the type of element selected in the Project Browser
	dim treeSelectedType
	treeSelectedType = Repository.GetTreeSelectedItemType()
	
	' Show the script output window
	Repository.EnsureOutputVisible "Script"
	
	' Define Global File IO Objects
	set fileSystemObject = CreateObject( "Scripting.FileSystemObject" )
	
	
	' Handling Code: Uncomment any types you wish this script to support
	' NOTE: You can toggle comments on multiple lines that are currently
	' selected with [CTRL]+[SHIFT]+[C].
	select case treeSelectedType
	
		case otElement
			' Code for when an element is selected
			dim theElement as EA.Element
			set theElement = Repository.GetTreeSelectedObject()
			
			
			
			
			dim scenarios as EA.Collection   ' the scenarios are a collection
		    set scenarios = theElement.Scenarios 
			
			
			
			dim i
			dim x
			
			
			Session.Output( "Scenarios" )
			
			for i = 0 to scenarios.Count - 1   ' for the number of scenarios in the use case
				dim currentScenario as EA.Scenario  ' 
				set currentScenario = scenarios.GetAt( i )  ' get the current one
 				
				dim scenarioSteps as EA.ScenarioStep  ' a scenario has multiple steps
				
						
				
				if (currentScenario.Type = "Basic Path" ) then  ' only interested in the basic path
				
				  					
					
				   
				   if ( currentScenario.Steps.Count > 0 ) then  ' if this scenario has multiple steps
					    
					   set outputFile = fileSystemObject.CreateTextFile( "c:\\SoftwareDevelopment\\" & theElement.Name & " " & theElement.ElementGUID & ".csv" , true ) ' change location for other implementations
						
					   for x = 0 to currentScenario.Steps.Count - 1   ' loop through each step
					   
					        
						   	set scenarioSteps = currentScenario.Steps.GetAt(x) ' get the step
							 
											
							scenarioSteps.Name = Replace(Replace(scenarioSteps.Name, Chr(10), " "), Chr(13), " ") ' replace any LF and CR within the step with a space
							
							scenarioSteps.Name = Replace(scenarioSteps.Name, "," , " ")  ' replace any commas in the step text with a space
							
							outputFile.WriteLine( scenarioSteps.Pos & "," & scenarioSteps.Name) ' output the step number and the name ( which is the text of the step )
						next
					    outputFile.Close()  ' close the file		
					   
				   end if
					
				end if
				
				
			
			next
			
'					
'		case otPackage
'			' Code for when a package is selected
'			dim thePackage as EA.Package
'			set thePackage = Repository.GetTreeSelectedObject()
'			
'		case otDiagram
'			' Code for when a diagram is selected
'			dim theDiagram as EA.Diagram
'			set theDiagram = Repository.GetTreeSelectedObject()
'			
'		case otAttribute
'			' Code for when an attribute is selected
'			dim theAttribute as EA.Attribute
'			set theAttribute = Repository.GetTreeSelectedObject()
'			
'		case otMethod
'			' Code for when a method is selected
'			dim theMethod as EA.Method
'			set theMethod = Repository.GetTreeSelectedObject()
		
		case else
			' Error message
			Session.Prompt "Please select a use case.", promptOK
			
	end select
	
end sub

OnProjectBrowserScript

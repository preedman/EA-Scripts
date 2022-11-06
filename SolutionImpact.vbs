option explicit

!INC Local Scripts.EAConstants-VBScript

'
' 
'
' Script Name: SolutionImpact
' Author: Paul Reedman
' Purpose: Colours elements on a diagram based on tag values - large - red, meduim - yellow, low - green
 ' Date:
'

'
' Diagram Script main function
'
sub OnDiagramScript()

    ' Show the script output window
	Repository.EnsureOutputVisible "Script"

	' Get a reference to the current diagram
	dim currentDiagram as EA.Diagram
	set currentDiagram = Repository.GetCurrentDiagram()
	' A tag
	dim tag as EA.TaggedValue
	' Variables
	dim i
	dim red, yellow, green
	dim tagName
	dim tagValue
	' RGB Values
	red = 255*65536+0*256+0
	yellow = 255*65536+255*256+0
	green = 0*65536+255*256+0
	' Object for loop - A diagram object is an object in a diagram
	dim obj as EA.DiagramObject
	' An element on the diagram
	dim elementOnDiagram as EA.Element
	' The Set of tags from a diagram
	dim theTags as EA.TaggedValue
	
	

	if not currentDiagram is nothing then
		For Each obj In currentDiagram.DiagramObjects   ' for every element on the current diagram
		
			set elementOnDiagram =  Repository.GetElementByID(obj.ElementID)   ' get an object on the diagram
			
			'Session.Output(elementOnDiagram.Name)
			'Session.Output(elementOnDiagram.TaggedValues.Count)
			
			if (elementOnDiagram.TaggedValues.Count > 0 ) then   ' if the element has tags
				set tagValue = elementOnDiagram.TaggedValues.GetByName("impact")   ' if the name of the tag is impact
				'Session.Output(tagValue.Name)
				'Session.Output(tagValue.Value)
				
				if (tagValue.Value = "large") then  ' if the value is large, then red
						    
						obj.BackgroundColor=red
							
						obj.Update
				end if
				
				if (tagValue.Value = "medium" ) then  ' if the value is medium, then yellow
				
					   obj.BackgroundColor=yellow
					   obj.Update
			    end if
				
				if (tagValue.Value = "small" ) then  ' if the value is small, then green
				
						obj.BackgroundColor=green
						obj.Update
				end if
				
				
			end if
			
			
			
		Next
	else
		Session.Prompt "This script requires a diagram to be visible", promptOK
	end if
	
	
	

end sub

OnDiagramScript

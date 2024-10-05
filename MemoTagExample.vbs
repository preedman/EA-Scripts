option explicit

!INC Local Scripts.EAConstants-VBScript

'
' Script Name: 
' Author: 
' Purpose: 
' Date: 
'

sub main
	' TODO: Enter script code here!
	Repository.EnsureOutputVisible("Script")
	
	
	Session.Output("Started")
	
	
	
	dim theElement as EA.Element
	dim i
	
	set theElement = Repository.GetElementByGuid("{FC2244DC-4221-49c0-93CC-DABE8FBD3ADA}")  ' get the element
	
	
	for i = 0 to theElement.TaggedValues.Count - 1  ' loop through the tags
	'	
		Session.Output(theElement.TaggedValues.GetAt(i).Name)  ' print out name
		if theElement.TaggedValues.GetAt(i).Name = "viewdef" then  ' if its the view def
			Session.Output(theElement.TaggedValues.GetAt(i).Notes)  ' then print out the notes because its a memo field 
			' https://sparxsystems.com/enterprise_architect_user_guide/17.0/add-ins___scripting/taggedvalue.html
		end if

	next

	Session.Output("End")
	
	
	
	
end sub

main
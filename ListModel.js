function ListModel()
{
	var objects as EA.Collection;  
	var object as EA.DiagramObject;
	var o as EA.DiagramObject;
	var element as EA.Element;
	var tv as EA.TaggedValue;
	var tvs as EA.Collection;
	
	Repository.EnsureOutputVisible( "Script" ); // display the script window
	
	Session.Output(theDiagram.Name + " " + theDiagram.Notes);
	
	objects = theDiagram.DiagramObjects;  // get all of the objects on the diagram
	
	for(var n = 0; n < objects.Count; n++) {   loop through all of the objects on the diagram
		o = objects.GetAt(n);  // get next object
		
		element = Repository.GetElementByID( o.ElementID );  // get the element on the diagram
		
		Session.Output("Element is " + element.Name);
		
		if (element.Name == 'Complaints') {  // if the element has a name of Complaints
			
			tvs = element.TaggedValues;  // get all of the tags for this element
			
			for(var e=0; e < tvs.Count; e++) {  // loop through all of the tags for this element
				
				tv = tvs.GetAt(e);  // get the next tag
				Session.Output("Tag " + tv.Name + " " + tv.Value); // print out the tag name and value
			}
		}
		
	}
	
	
}
ListModel();
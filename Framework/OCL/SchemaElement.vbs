'[path=\Framework\OCL]
'[group=OCL]


'Author: Geert Bellekens
'Date: 2017-12-06
'Purpose: Class representing a Schema Element

Class SchemaElement 
	'private variables
	Private m_Source
	Private m_Name
	private m_IsRoot
	Private m_IsRedefinition
	Private m_Properties
	Private m_Redefines
	private m_ReferencingProperties
	private m_Schema
	
	'constructor
	Private Sub Class_Initialize
		m_Name = ""
		m_Name = ""
		m_isRoot = false
		m_IsRedefinition = false
		set m_Source = nothing
		set m_Schema = nothing
		set m_Properties = CreateObject("Scripting.Dictionary")
		set m_Redefines = CreateObject("Scripting.Dictionary")
		set m_ReferencingProperties = CreateObject("Scripting.Dictionary")
	End Sub
	
	' Source property. (EA.Element)
	Public Property Get Source
		set Source = m_Source
	End Property
	Public Property Let Source(value)
		set m_Source = value
		if m_Name = "" then
			m_Name = me.Source.Name
		end if
	End Property
	
	' Schema property OCL.Schema
	Public Property Get Schema
		set Schema = m_Schema
	End Property
	Public Property Let Schema(value)
		set m_Schema = value
	End Property
	
	' Name property.
	Public Property Get Name
		Name = m_Name
	End Property
	Public Property Let Name(value)
		m_Name = value
		'debug
		'Session.Output "Creating element : " & me.Name
	End Property
	
	' IsRoot property.
	Public Property Get IsRoot
		IsRoot = m_IsRoot
	End Property
	Public Property Let IsRoot(value)
		m_IsRoot = value
	End Property
	
	' IsRedefinition property. (boolean)
	Public Property Get IsRedefinition
		IsRedefinition = m_IsRedefinition
	End Property
	Public Property Let IsRedefinition(value)
		m_IsRedefinition = value
	End Property
	
	' Guid property.
	Public Property Get GUID
		GUID = me.Source.ElementGUID
	End Property
	
	' Properties property
	Public Property Get Properties
		set Properties = m_Properties
	End Property
	
	' Redefines property
	Public Property Get Redefines
		set Redefines = m_Redefines
	End Property
	
	' ReferencingProperties property. (Dictionary with GUID and SchemaProperties)
	Public Property Get ReferencingProperties
		set ReferencingProperties = m_ReferencingProperties
	End Property
	
	'add a referencing property
	public function addReferencingProperty(referencingProperty)
		if not me.ReferencingProperties.Exists(referencingProperty.GUID) then
			me.ReferencingProperties.Add referencingProperty.GUID, referencingProperty
			'debug
			'Session.Output "Adding referencing property: " & referencingProperty.Name & "with guid " & referencingProperty.GUID & " to element " & me.Name
		end if
	end function
	
	'remove a referencing property
	public function removeReferencingProperty(referencingProperty)
		if me.ReferencingProperties.Exists(referencingProperty.GUID) then
			me.ReferencingProperties.Remove referencingProperty.GUID
			'debug
			'Session.Output "Removing referencing property: " & referencingProperty.Name & "with guid " & referencingProperty.GUID & " from element " & me.Name
			if me.ReferencingProperties.Count = 0  _
				AND me.Redefines.Count = 0 then
				Delete
			end if
		end if	
	end function
	
	public function Delete()
		'debug
		'Session.Output "Deleting element " & me.Name
		'first delete all owned properties
		dim schemaProperty
		for each schemaProperty in me.Properties.Items
			schemaProperty.Delete
		next
		'then remove me from the schema
		me.Schema.RemoveSchemaElement me
		'and remove me from the parent redefine
	end function
	
	'Merges redefines that have the same properties.
	public function mergeAllRedefines()
		dim redefine
		dim i
		dim redefinesToDelete
		set redefinesToDelete = CreateObject("System.Collections.ArrayList")
		for i = me.Redefines.Count -1 to 0 step -1 
			set redefine = me.Redefines.Items()(i)
			dim merged
			'compare with me
			merged = mergeRedefines(me, redefine)
			if not merged then
				'loop the redefines again to find equals
				dim otherRedefine
				for each otherRedefine in me.Redefines.Items
					 merged = mergeRedefines(otherRedefine, redefine)
				next
			end if
			'if merged then add the redefine to the list of redefines to delete
			if merged then
				removeRedefine redefine
			end if
		next
		'at the end, if this element doesn't have any referencing properties then replace it with the first redefine that has any
		if me.ReferencingProperties.Count = 0 then
			for each redefine in me.Redefines.items
				if redefine.ReferencingProperties.Count > 0 then
					'replace the properties properties
					me.Properties.RemoveAll
					dim redefinedProperty
					for each redefinedProperty in redefine.Properties.Items
						'add to properties list
						me.Properties.Add redefinedProperty.GUID, redefinedProperty
						'set me as owner
						redefinedProperty.Owner = me
					next
					dim usingProperty
					'set using properties to me
					for each usingProperty in redefine.ReferencingProperties.Items
						usingProperty.ClassifierSchemaElement = me
					next
					'delete redefine
					removeRedefine redefine
					'exit the for loop
					exit for
				end if
			next
		end if
	end function
	
	private function mergeRedefines(redefineToKeep, redefineToDelete)
		'default false
		mergeRedefines = false
		'compare names
		if redefineToKeep.Name = redefineToDelete.Name then
			'don't even start if it's the same redefine
			exit function
		end if
		if isEquivalentRedefine(redefineToKeep, redefineToDelete) then
			'merge the two redefines, keeping redefineToKeep
			dim usingProperty
			'set the using properties to use redefineToKeep
			for each usingProperty in redefineToDelete.ReferencingProperties.Items
				usingProperty.ClassifierSchemaElement = redefineToKeep
			next
			mergeRedefines = true
		end if
	end function
	
	private function isEquivalentRedefine(redefine, otherRedefine)
		isEquivalentRedefine = true
		dim schemaProperty
		'first check if they have the same number of properties
		if redefine.Properties.Count <> otherRedefine.Properties.Count then
			isEquivalentRedefine = false
			exit function
		end if
		'then compare each property
		for each schemaProperty in redefine.Properties.Items
			dim found
			found = false
			dim otherProperty
			for each otherProperty in otherRedefine.Properties.Items
				if schemaProperty.GUID = otherProperty.GUID then
					found = true
				end if
			next
			'if we haven't found an equal property then we exit
			if not found then
				isEquivalentRedefine = false
				exit for
			end if
		next
	end function

	public function deleteProperty(propertyToDelete)
		if not propertyToDelete is nothing then
			if me.Properties.Exists(propertyToDelete.GUID) then
				me.Properties.Remove(propertyToDelete.GUID)
			end if
		end if
	end function
	
	public function getProperty(identifierPart, Byref isNew)
		'initialize null
		set getProperty = nothing
		'clean identifierPart (remove quotes)
		dim cleanIdentifier
		cleanIdentifier = replace(trim(identifierPart),"'","")
		'first check if there is an attribute on the localContext with the given name
		dim sqlGetAttribute
		sqlGetAttribute = "select a.ID from t_attribute a " & _
						" where a.Object_ID = " & me.Source.ElementID  & _
						" and a.Name = '" & cleanIdentifier & "' "
		'get the attribute
		dim attributes
		set attributes = getattributesFromQuery(sqlGetAttribute)
		if attributes.Count > 0 then
			'return the first attribute
			set getProperty = me.addAttributeProperty(attributes(0), isNew)
		else
			'get association end
			dim associationEnd
			set associationEnd = nothing
			'first check target end
			dim sqlGetTargetEnd
			sqlGetTargetEnd	= "select c.Connector_ID from t_connector c" & _
							" where c.Start_Object_ID = " & me.Source.ElementID  & _
							" and c.DestRole = '" & cleanIdentifier & "' "
			dim connector as EA.Connector
			dim connectors
			set connectors = getConnectorsFromQuery(sqlGetTargetEnd)
			if connectors.Count > 0 then
				set connector = connectors(0)
				set associationEnd = connector.SupplierEnd
			else
				'then source end
				dim sqlGetSourceEnd
				sqlGetSourceEnd = "select c.Connector_ID from t_connector c" & _
							" where c.End_Object_ID = " & me.Source.ElementID  & _
							" and c.SourceRole = '" & cleanIdentifier & "' "
				set connectors = getConnectorsFromQuery(sqlGetSourceEnd)
				if connectors.Count > 0 then
					set connector = connectors(0)
					set associationEnd = connector.ClientEnd
				end if
			end if
			'create the SchemeProperty
			if not associationEnd is nothing then
				set getProperty = me.addConnectorEndProperty(associationEnd, connector, isNew)
			end if
		end if
		'set owner of the property
		if not getProperty is nothing then
			getProperty.Owner = me
		end if
	end function
		
	' Add Attribute poperty to properties
	Public function addAttributeProperty(newAttribute, Byref isNew)
		if me.Properties.Exists(newAttribute.AttributeGUID) then
			'return existing item
			set addAttributeProperty = me.Properties.Item(newAttribute.AttributeGUID)
			isNew = false
		else
			'create new item
			dim newProperty
			set newProperty = new SchemaProperty
			newProperty.Source = newAttribute
			me.Properties.Add newProperty.GUID, newProperty
			isNew = true
			'return new item
			set addAttributeProperty = newProperty
		end if
	End function
	
	' Add Attribute poperty to properties
	Public function addConnectorEndProperty(newConnectorEnd, newConnector, Byref isNew)
		if me.Properties.Exists(newConnector.ConnectorGUID) then
			'return existing item
			set addConnectorEndProperty = me.Properties.Item(newConnector.ConnectorGUID)
			isNew = false
		else
			'create new item
			dim newProperty
			set newProperty = new SchemaProperty
			newProperty.Source = newConnectorEnd
			newProperty.Connector = newConnector
			me.Properties.Add newProperty.GUID, newProperty
			isNew = true
			'return new item
			set addConnectorEndProperty = newProperty
		end if
	End function
	
	' Add a redefines  element
	Private function addRedefine(newRedefine)
		if not m_Redefines.Exists(newRedefine.Name) then
			m_Redefines.Add newRedefine.Name, newRedefine
		end if
	End function
	'removes a redefine
	public function removeRedefine(redefine)
		'debug
		'Session.Output "Request remove redefine " & redefine.Name
		if me.Redefines.Exists(redefine.Name) then
			me.Redefines.Remove redefine.Name
			'debug
			'Session.Output "Actually removing redefine " & redefine.Name
		end if
	end function
	'adds a new redefine
	public function addNewRedefine()
		'create new schemaElement
		dim newRedefine
		set newRedefine = new SchemaElement
		newRedefine.Name = me.Name & "_" & m_Redefines.Count + 1
		newRedefine.Source = me.Source
		newRedefine.Schema = me.Schema
		newRedefine.IsRedefinition = true
		'add it to the list of redefines
		addRedefine newRedefine
		'return it
		set addNewRedefine = newRedefine
	end function

end Class
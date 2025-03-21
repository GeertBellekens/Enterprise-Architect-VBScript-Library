'[group=Documents]

!INC Local Scripts.EAConstants-VBScript

'
' Script Name: Document
' Author: Geert Bellekens
' Purpose: Class to support document generation in a generic way
' Date: 2024-02-23
'

dim g_allDocuments


Class Document 
	Private m_Name
	Private m_Description
	Private m_GenerateFunction
	Private m_ValidQuery

	Private Sub Class_Initialize
	  m_Name = ""
	  m_Description = ""
	  m_GenerateFunction = ""
	  m_ValidQuery = ""
	End Sub

	' Name property.
	Public Property Get Name
		Name = m_Name
	End Property
	Public Property Let Name(value)
		'manage alldocuments dictionary
		if len(m_Name) > 0 then
			me.AllDocuments.Remove(m_Name)
		end if
		dim m_allDocuments
		set m_allDocuments = me.AllDocuments
		set m_allDocuments(value) = me
		'set value
		m_Name = value
	End Property
	
	' Description property.
	Public Property Get Description
		Description = m_Description
	End Property
	Public Property Let Description(value)
		m_Description = value
	End Property
	
	' GenerateFunction property.
	Public Property Get GenerateFunction
		GenerateFunction = m_GenerateFunction
	End Property
	Public Property Let GenerateFunction(value)
		m_GenerateFunction = value
	End Property
	
	' ValidQuery property.
	Public Property Get ValidQuery
		ValidQuery = m_ValidQuery
	End Property
	Public Property Let ValidQuery(value)
		m_ValidQuery = value
	End Property
	
	public property get AllDocuments
		set AllDocuments = getAllDocuments()
	end Property
	
	public function Generate()
		select case me.Name
			case "ProjectDependencies"
				ProjectDependenciesExport
			case "BusinessBook"
				BusinessBook
			case "DataModel"
				DataModelExport
			case "CapabilityTree"
				CapabilityTreeExport
			case "AmbitionToCapability"
				AmbitionToCapabilityExport
			case "InformationFlows"
				InformationFlowsExport
			case "Ambitions"
				GenerateAmbitionsDocument
			case "Architecture Blueprint"
				GenerateArchitectureBlueprintDocument				
			case "Project Scope"
				GenerateScopeDocument	
			case "Capability Stories"
				GenerateCapabilityStoriesDocument				
		end select
	end function
	
	public function isValidForPackage(package)
		dim packageTreeIDString
		packageTreeIDString = getPackageTreeIDString(package)
		dim packageValidQuery 
		packageValidQuery = Replace(me.ValidQuery, "#Branch#", packageTreeIDString)
		packageValidQuery = Replace(packageValidQuery, "#PACKAGEID#", package.PackageID)
		dim result
		set result = getArrayListFromQuery(packageValidQuery)
		if result.Count > 0 then
			isValidForPackage = true
		else
			isValidForPackage = false
		end if
	end function
	
	
End Class


function getAllDocuments()
	if not isObject(g_allDocuments) then
		set g_allDocuments = CreateObject("Scripting.Dictionary")
	end if
	'return
	set getAllDocuments = g_allDocuments
end function

function generateDocumentsForPackage(package)
	dim filteredDocuments
	set filteredDocuments = CreateObject("Scripting.Dictionary")
	'check for each document is it's valid for this package
	dim doc
	dim i
	i = 0
	for each doc in getAllDocuments().Items
		if doc.isValidForPackage(package) then
			i = i + 1
			filteredDocuments.Add i, doc
		end if
	next
	if filteredDocuments.Count = 0 then
		msgbox "No documents found to generate for package '" & package.Name & "'" , VBCritical, "No Documents Found!"
		exit function
	end if
	'ask user for input
	dim selectMessage
	selectMessage = "Please enter the number of the document" & vbNewLine 
	i = 0
	for each doc in filteredDocuments.Items
		  i = i + 1
		  selectMessage = selectMessage & vbNewLine & i & ": " & doc.Name & vbNewLine & doc.Description & vbNewLine 
		  'selectMessage = selectMessage & vbNewLine & i & ": " & doc.Name & vbNewLine & "        " & doc.Description & vbNewLine 
	next
	dim response
	response = InputBox(selectMessage, "Document to generate", "1" )
	if isNumeric(response) then
		  if Cstr(Cint(response)) = response then 'check if response is integer
				 dim selectedID
				 selectedID = Cint(response)
				 if selectedID > 0 and selectedID <= filteredDocuments.Count then
					   'return the version control ID
					   set doc = filteredDocuments(selectedID)
					   doc.Generate
				 end if
		  end if
	end if
end function

public Function getUserSelectedDocumentFileName()
		dim project
		set project = Repository.GetProjectInterface()
		dim fileName
		fileName = project.GetFileNameDialog ("", "Documents|*.docx", 1, 2 ,"", 1) 'save as with overwrite prompt: OFN_OVERWRITEPROMPT
		'get extension
		dim fso
		Set fso = CreateObject("Scripting.FileSystemObject")
		dim extension
		extension = fso.GetExtensionName(fileName)
		if len(extension) = 0 then
			fileName = fileName & ".docx"
		else
			fileName = left(fileName, len(fileName) - len(extension)) & "docx"
		end if
		'return
		getUserSelectedDocumentFileName = fileName
end function
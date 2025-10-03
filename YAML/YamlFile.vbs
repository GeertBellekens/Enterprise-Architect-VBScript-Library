'[group=YAML]
'option explicit

!INC Local Scripts.EAConstants-VBScript

'
' Script Name: YamlDocument
' Author: Geert Bellekens
' Purpose: Parse a YamlFile and return a YamlNode(?document)
' Date: 
'

Class YamlFile
	'private variables
	Private m_TextFile
	Private regexp
	private m_fileFilter
	private m_rootNode

	Private Sub Class_Initialize
		set m_TextFile = new TextFile
		Set regexp = CreateObject("VBScript.RegExp")
		regexp.Global = True   
		regexp.IgnoreCase = False
		m_fileFilter = "YAML files|*.yaml"
		set m_rootNode = nothing
	End Sub
	
	
	' FileName property.
	Public Property Get FileName
	  FileName = m_TextFile.FileName
	End Property
	Public Property Let FileName(value)
	  m_TextFile.FileName = value
	End Property
		
	' FullPath property.
	Public Property Get FullPath
	  FullPath = m_TextFile.FullPath
	End Property	
	public Property Let FullPath(value)
		m_TextFile.FullPath = value
		m_TextFile.LoadContents
	End Property
	
	'let the user select a file from the file system
	public function UserSelect(initialDir)
		dim selectedFileName
		selectedFileName = ChooseFile(initialDir,m_fileFilter)
		'check if anything was selected
		if len(selectedFileName) > 0 then
			me.FullPath = selectedFileName
			UserSelect = true
		else
			UserSelect = false
		end if
	end function
	
	public Function getUserSelectedFile()
		dim project
		set project = Repository.GetProjectInterface()
		me.FullPath = project.GetFileNameDialog ("", "YAML files|*.yaml", 1, 2 ,"", 1) 'save as with overwrite prompt: OFN_OVERWRITEPROMPT
	end function
	
	public function Save()
		dim stringContents
		stringContents = ""
		if not m_rootNode is nothing then
			stringContents = m_rootNode.ToString(false)
		end if
		m_TextFile.Contents = stringContents
		m_TextFile.Save
	end function
	
	Function Parse
		'divide in lines
		dim lines
		lines = split(m_TextFile.Contents, vbNewLine)
		dim line
		dim rootNode
		set rootNode = new YamlParentNode 'root node
		dim parentNode
		set parentNode = rootNode
		dim currentNode
		set currentNode = rootNode
		for each line in lines	
			'check comments
			dim isComment
			regExp.Pattern ="^ *#"
			isComment = regExp.Test(line)
			regExp.Pattern = "( *)(- )?(?:(.*?): ?)?(.*)$"
			dim matches
			set matches = regExp.Execute(line)
			dim match
			if not isComment then
				if matches.Count > 0  _
				  and len(line) > 0 then
					set match = matches(0)
					dim indent
					dim seqInd
					dim anchor
					dim value	
					indent = match.SubMatches(0)
					seqInd = match.SubMatches(1)
					anchor = match.SubMatches(2)
					value = match.SubMatches(3)
					dim isSeq
						isSeq = len(seqInd) > 0
					dim isAnchor
						isAnchor = len(anchor) > 0
					dim isValue
						isValue = len(value) > 0
					'get the parentNode based on indent and seqInd
					dim sequenceIndentation
					sequenceIndentation = ""
					if isSeq then
						sequenceIndentation = space(2)
					end if
					set parentNode = parentNode.getParent(indent & sequenceIndentation)		
					'create the nodes
					if isSeq then
						'create new node
						set currentNode = new YamlNode
						currentNode.IsSequence = true
						'add to parent
						parentNode.addNode(currentNode)
						if isAnchor then
							'if there is also anchor node, then we make the value of the currentNode a parentNode
							currentNode.Value = new YamlParentNode
							set parentNode = currentNode.Value
							'set noIndent true so the anchor is placed on the same row without extra indentation
							parentNode.NoIndent = true
						end if
					end if
					if isAnchor then
						'create new node
						set currentNode = new YamlNode
						'set key
						currentNode.setScalarKey(anchor)
						'add to parent
						parentNode.addNode(currentNode)
					end if
					if isAnchor then
						'create new parent key if needed
						if not isValue then
							'the next line could contain child nodes
							currentNode.Value = new YamlParentNode
							set parentNode = currentNode.Value
						end if
					end if
					'set value
					if isValue then
						if not isAnchor and not isSeq then
							currentNode.appendScalarValue(vbNewLine & trim(line))
						else
							currentNode.appendScalarValue(value)
						end if
					end if

				else 'matches.Count = O
					currentNode.appendScalarValue(vbNewLine & trim(line))
				end if
			end if
			'debug
			'Session.Output "currentNode level: " & currentNode.level &  " parentNode.level: " & parentNode.level & " indent: " & indent & " anchor: " & anchor & " value: " & value 
		next
		set m_rootNode = rootNode
		'return root node
		set Parse = rootNode
	end function
	
end Class




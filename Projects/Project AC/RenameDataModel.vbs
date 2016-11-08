'[path=\Projects\Project AC]
'[group=Acerta Scripts]
option explicit

!INC Local Scripts.EAConstants-VBScript

'
' Script Name: RenameDataModel
' Author: Geert Bellekens
' Purpose: Renames a data model created by the model wizard
' Date: 2016-10-07
'
sub main
	dim dataModel as EA.Package;
	set dataModel = Repository.GetTreeSelectedPackage
	if not dataModel is nothing then
		if dataModel.Element.Stereotype = "DataModel" then
			dim originalName
			originalName = dataModel.Name
			'ask user for name
			dim modelName
			modelName = InputBox( "Please enter new name for the data model", "Data Model Name" )
			if len(modelName) > 0 then
				'rename data model package
				dataModel.Name = modelName
				dataModel.Update
				'rename diagram(s) under data model
				dim diagram as EA.Diagram
				for each diagram in dataModel.Diagrams
					if diagram.Name = originalName then
						diagram.Name = modelName
						diagram.Update
					end if
				next
				'rename elements under data model
				dim element as EA.Element
				for each element in dataModel.Elements
					if element.Name = originalName then
						element.Name = modelName
						element.Update
					end if
				next
				'rename packages under data model
				dim subPackage as EA.Package
				for each subPackage in dataModel.Packages
					if subPackage.Name = originalName then
						subPackage.Name = modelName
						subPackage.Update
						'update any diagrams under this package
						for each diagram in subPackage.Diagrams
							if diagram.Name = originalName then
								diagram.Name = modelName
								diagram.Update
							end if
						next
					end if
				next
				msgbox "Finished renaming"
			end if
		else
			msgbox "Please select a package with stereotype «DataModel»"
		end if
	end if
end sub

main
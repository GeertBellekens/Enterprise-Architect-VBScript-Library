'[path=\Framework\ModelValidation\Example]
'[group=ModelValidationExample]

' Change this to the value in System Output/Logging after ExampleModelValidationRules_LoadRules has run
dim BASE_ID
BASE_ID = 800000

!INC ModelValidation.Utils

dim exampleCategoryId, exampleRuleOneId

exampleCategoryId		= makeId(BASE_ID, 0)
exampleRuleOneId 		= makeId(BASE_ID, 1)
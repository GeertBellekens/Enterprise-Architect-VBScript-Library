'[path=\Framework\ModelValidation]
'[group=ModelValidation]

' Change this to the value in System Output/Logging 
dim BASE_ID
BASE_ID = 800000

!INC ModelValidation.Utils

dim testCategoryId, testRuleOneId

testCategoryId		= makeId(BASE_ID, 0)
testRuleOneId		= makeId(BASE_ID, 1)
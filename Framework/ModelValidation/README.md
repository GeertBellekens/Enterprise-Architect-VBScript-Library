# Overview

Design > Model > Manage > Validate > Validate Current Package
Keyboard SHortcut Ctrl+Alt+V

# Suggested VBScript Design

* <Your Category>ModelValidationRules\_LoadRules vbscript. See  ExampleModelValidationRules\_LoadRules for an example
* <Your Category>ModelValidationRuleConstants. See ExampleModelValidationRuleConstants for an example.
* <Your Category>ModelValidationRule_<Your Rule Name>. See ModelValidationRule\_Template for a template to start with.

See ModelValidationExample for samples.

# Problems

## EA-Matic RESETS scripts every 5 minutes

Because EA-Matic resets scripts, there is no way to remember the Rule Categories and Rules in a global scope between invocations.

If the script is reset then this information is lost.

All the events described below use IDs to refer to rules, and these IDs are assigned at runtime so if another add-in creates rules this will shift your rule's ID.

The work-around is to create the rules in `EA_FileOpen` (as `EA_OnInitializeUserRules` doesn't seem to be called) and to write the id's of the categories and rules created to Session.output. The rules can then hard code those id's in the scripts. Sparx EA doesn't care at all if the Rules run actually match the Rules defined, only that the Rule ID used is one that has been defined.

## EA_OnInitializeUserRules not called

`EA_OnInitializeUserRules` doesn't appear to be called via EA-Matic, tried putting that code in and nothing happens.

Use `EA_FileOpen` instead. "See Problems > EA-Matic RESETS scripts every 5 minutes" for issues with using EA-Matic to write Model Validation Rules.

# Model Validation Events

See [Model Validation Events](https://sparxsystems.com/enterprise_architect_user_guide/16.1/add-ins___scripting/model_validation_broadcasts.html) for more details.

* [EA_OnInitializeUserRules](https://sparxsystems.com/enterprise_architect_user_guide/16.1/add-ins___scripting/ea_oninitializeuserrules.html) (Note: See Problems > EA_OnInitializeUserRules not called)

* [EA_OnStartValidation](https://sparxsystems.com/enterprise_architect_user_guide/16.1/add-ins___scripting/ea_onstartvalidation.html)

* [EA_OnEndValidation](https://sparxsystems.com/enterprise_architect_user_guide/16.1/add-ins___scripting/ea_onendvalidation.html)

* [EA_OnRunElementRule](https://sparxsystems.com/enterprise_architect_user_guide/16.1/add-ins___scripting/ea_onrunelementrule.html)

* [EA_OnRunPackageRule](https://sparxsystems.com/enterprise_architect_user_guide/16.1/add-ins___scripting/ea_onrunpackagerule.html)

* [EA_OnRunDiagramRule](https://sparxsystems.com/enterprise_architect_user_guide/16.1/add-ins___scripting/ea_onrundiagramrule.html)

* [EA_OnRunConnectorRule](https://sparxsystems.com/enterprise_architect_user_guide/16.1/add-ins___scripting/ea_onrunconnectorrule.html)

* [EA_OnRunAttributeRule](https://sparxsystems.com/enterprise_architect_user_guide/16.1/add-ins___scripting/ea_onrunattributerule.html)

* [EA_OnRunMethodRule](https://sparxsystems.com/enterprise_architect_user_guide/16.1/add-ins___scripting/ea_onrunmethodrule.html)

* [EA_OnRunParameterRule](https://sparxsystems.com/enterprise_architect_user_guide/16.1/add-ins___scripting/ea_onrunparameterrule.html)

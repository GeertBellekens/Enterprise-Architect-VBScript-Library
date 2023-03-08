# EA VBScript Guidelines

## Script Design

You want to apply standard software engineering practices to your VB Script design.

### Scope

As VBScript does not have a namespace or module system, everything is in one global scope and your script can access other scripts only through the `!INC` (this appears to be an EA VBScript feature, see "Include Script Libraries" [Script Editor](https://sparxsystems.com/enterprise_architect_user_guide/16.0/add-ins___scripting/script_editors.html))

### Using Classes to provide namespaces

You will need to determine whether your code is a general utility function/procedure that should be in the global namespace, or whether related functions/procedures should be kept together and work on common data and therefore be wrapped in a Class.

By wrapping your code in a Class block it provides another level of namespace to avoid namespace collision in the global namespace.

### Layering

If you are going to write more than one script you'll want to ensure you [Don't repeat yourself](https://en.wikipedia.org/wiki/Don%27t_repeat_yourself).

Scripts that are run inside EA should only contain the main procedure and the call of main. The main method should collect the data needed and then invoke methods from the included scripts.

The included scripts should do the actual work.

Common or reusable functions/procedures should be extracted into their own scripts possibly grouping related items.

See Layering diagram:

![EA Layering Diagram](./EA-Matic.png)

### option explicit

[option explicit](https://learn.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/scripting-articles/bw9t3484(v=vs.84)) must appears in a script before any other statements.

Because of this `option explicit` can only be in the top level main scripts, it can not be in any scripts that will be included in another.

## EA Object Model

See [EA Object Model Reference](https://www.sparxsystems.com/enterprise_architect_user_guide/15.2/automation/reference.html) for the details of all objects available in the object model provided by the Automation Interface.

Some key links are listed here:

* [Repository](https://www.sparxsystems.com/enterprise_architect_user_guide/15.2/automation/repository3.html)

* [Diagram](https://www.sparxsystems.com/enterprise_architect_user_guide/15.2/automation/diagram2.html)

* [DiagramObject](https://www.sparxsystems.com/enterprise_architect_user_guide/15.2/automation/diagramobjects.html)

* [Element](https://www.sparxsystems.com/enterprise_architect_user_guide/15.2/automation/element2.html)

# EA-Matic Guidelines

Also read [Enterprise Architect VBScript Guidelines](ea-vbscript-guidelines.md)

Your EA-Matic script is a thin wrapper to invoke your included script to do the
actual work.

There should be very little to trouble shoot in your EA-Matic script as it doesn't do much different than the VBScript with the main method, it should

* Handle the Event

* Collect the data needed to invoke the methods for the actual work

* Invoke the method

See [VBScript Template](./vbscript-template.md), the only difference will be you don't use `main` but an event handler, and you must not invoke the event handler directly.

See [broadcast
events](https://sparxsystems.com/enterprise_architect_user_guide/15.2/automation/broadcastevents.html)
in the Sparx documentation for all the events that the EA-Matic settings dialog can
handle. These pages list the events and the arguments provided to the event
handler.

## Workflow

Write your EA-Matic script, then open and close EA-Matic settings to force EA-Matic to refresh its in memory copy (Specialize > EA-Matics > settings), then run your script.

# EA-Matic Trouble Shooting

## Variable uses an Automation type not supported in VBScript

If EA-Matic throws an An unhandled exception MessageBox with an error lile:

```
Error Message: Script '<Group Name>.<Script Name>' failed: Variable uses an Automation type not supported in VBScript
```

A likely cause is you are trying to concatenate a non-string type to a string.

Unfortunately there is no easy way to provide further detail on the error. You
will need to manually eyeball every variable and check its type and use CStr()
for non-string types.

## EA-Matic only saves work every 5 minutes

**This is by design.**

The thing is that EA-Matic will be triggered with each and every event in EA. If
we were to read and interpret all scripts on every event, that would slow down
EA considerably.

That's why EA-Matic keeps an in-memory copy of all EA-Matic functions. Every 5
minutes (or when you open the settings) EA-Matic will check if any changes have
been made to the EA-Matic scripts, and if so, it will refresh the scripts in
memory.

This can sometimes be annoying for the script developer, but it's the best
solution for the users.

## EA-Matic scripts are stale, old versions

When you run EA-Matic the scripts aren't doing what you just finished writing
and saved into you VBScript.

See [EA-Matic only saves work every 5 minutes](#ea-matic-only-saves-work-every-5-minutes)

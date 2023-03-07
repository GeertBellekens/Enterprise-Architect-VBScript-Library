# EA-Matic Trouble Shooting

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

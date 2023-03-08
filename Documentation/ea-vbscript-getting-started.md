# EA VBScript Getting Started

## Install EA-Matic

Why do I need to install EA-Matic when I just want to write VBScripts for Enterprise Architect?

Because you will want to version control your scripts and doing that manually is a huge pain. EA-Matic will autoamatically save your script changes every 5 minutes to your specified directory and you can use whatever version control system you want to manage that directory.

Manually using `SaveAllScripts` from `Framework/Tools/Script Management` is a painful alternative.

## Checkout the Source code

Use git to clone [https://github.com/GeertBellekens/Enterprise-Architect-VBScript-Library](https://github.com/GeertBellekens/Enterprise-Architect-VBScript-Library) to somewhere on your local hard drive.

## Boostrap your Scripts

When you start a fresh project you will have no scripts loaded.

You can not use `LoadScripts` from `Framework/Tools/Script Management` as none of the included files are available.

You need to use `LoadScriptsBootstrap` which is a manual copy-and-paste of the bare bones `LoadScripts` (that isn't kept up to date, but will get us boot strapped!)

In Enterprise Architect, click the `Specialize` Menu and then the `Script Library` button in the `Tools` ribbon.

Click the `New Group > New Normal Group` button and name it `Script Management`.

Click the `New Script > New VBScript` button and name the file `LoadScriptsBootstrap`.

Double click the `LoadScriptsBootstrap` script to open it in the editor.

On your computer browse to the git clone and find the `Framework/Tools/Script Management\LoadScriptsBootstrap` file and copy its contents into the VBScript Editor.

Change the location of `Const SCRIPT_FOLDER` to point to your git clone.

Save `LoadScriptsBootstrap`.

Right click `LoadScriptsBootstrap` and click `Run Script`.
Choose the `Utils` folder and Click `OK`.

Repeat this for the `Wrappers` and `Tools` Folder.

Bootrapping is now complete.

## Start Coding

Make sure to finish reading the rest of the [documentation](./README.md) before you start coding, it is easy to make silly mistakes.


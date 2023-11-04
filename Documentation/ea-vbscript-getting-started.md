# EA VBScript Getting Started

## Install EA-Matic

Why do I need to install EA-Matic when I just want to write VBScripts for Enterprise Architect?

Because you will want to version control your scripts and doing that manually is a huge pain. EA-Matic will automatically save your script changes every 5 minutes to your specified directory and you can use whatever version control system you want to manage that directory.

Manually using `SaveAllScripts` from `Framework/Tools/Script Management` is a painful alternative.

## Checkout the Source code

Use git to clone [https://github.com/GeertBellekens/Enterprise-Architect-VBScript-Library](https://github.com/GeertBellekens/Enterprise-Architect-VBScript-Library) to somewhere on your local hard drive.

## Boostrap your Scripts

When you start a fresh project you will have no scripts loaded.

You can not use `LoadScripts` from `Framework/Tools/Script Management` as none of the included files are available.

You need to use `LoadScriptsBootstrap` which is a generated inlined version of `LoadScripts` and will load all the required dependencies as well.

In Enterprise Architect, click the `Specialize` Menu and then the `Script Library` button in the `Tools` ribbon.

Click the `New Group > New Normal Group` button and name it `Script Management`.

Click the `New Script > New VBScript` button and name the file `LoadScriptsBootstrap`.

Double click the `LoadScriptsBootstrap` script to open it in the editor.

On your computer browse to the git clone and find the `Framework/Tools/Script Management\LoadScriptsBootstrap` file and copy its contents into the VBScript Editor.

Save `LoadScriptsBootstrap`.

Open the [Local Paths Dialog](https://sparxsystems.com/enterprise_architect_user_guide/16.1/modeling_domains/localpathdlg.html) from the Ribbon > Develop > Source Code > Options > Configure Local Paths.
To the right of the `Path` text field is the browse for folder button, click it and browse to the location of your scripts folder. In the `ID` text field enter `EA-Matic Script Folder`. Click the `Type` drop-down and select `Visual Basic`. Click Save. Click Close.

Right click `LoadScriptsBootstrap` and click `Run Script`.

Choose the `Utils` folder and Click `OK`.
Repeat this for the `Wrappers` and `Tools` Folder.

Bootrapping is now complete.

## Start Coding

Make sure to finish reading the rest of the [documentation](./README.md) before you start coding, it is easy to make silly mistakes.


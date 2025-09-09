> [!CAUTION]
> This is a work in progress, for demostration purposes only

# alibre-vscodium-addon

Proof-of-concept Alibre Design add-on that uses Alibre Script/AlibreX (IronPython 2 & 3) as commands. Instead of writing in a .NET language (C# or VB), you write in IronPython. A .NET language is only used for compiling the add-on as a DLL.

[CURRENT STATUS](https://github.com/stephensmitchell/alibre-vscodium-addon/discussions/1?sort=new)

## Purpose

To evaluate the overall process and steps necessary to create Alibre Design add-ons that use Alibre Script (IronPython) code as commands. 

###  ADK or Add-on Developer Kit

This add-on is part of the alibre-script-adk project, an effort to share lessons learned and provide public resources for modern Alibre Design scripting and programming. Alibre's built-in scripting add-on does not provide a solution for running scripts outside of the add-on. The ADK aims to solve this limitation.

## Who is this for

Anyone who would like to build an Alibre Design add-on, with or without Alibre Script (IronPython) code. 

## What it does

After installation, you'll see a menu and/or ribbon button added to the Alibre Design user interface. Clicking the button will run code that launches VSCodium text editor with the add-on scripts folder loaded. You can then use VSCodium to edit your scripts before running them.

## How it works

> [!NOTE]
> Work In Progress

## Known Uses

N/A

## Installation

See Releases for the installer and portable zip file.

[VSCodium](https://vscodium.com/) is included, no setup required.

### Additional Resources

N/A

## Contribution

Contributions to the codebase are not currently accepted, but questions and comments are welcome.

## Acknowledgment and License

MIT — see license.

## Credit & Citation

N/A

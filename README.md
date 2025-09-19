> [!CAUTION]
> This is a work in progress, for demostration purposes only

# alibre-py-gear-addon

Proof-of-concept Alibre Design add-on that uses Alibre Script/AlibreX (IronPython 2 or 3) as commands. Instead of writing in a .NET language (C# or VB), you write in IronPython. A .NET language is only used for compiling the add-on to a DLL.

## Purpose

To evaluate the overall process and steps necessary to create Alibre Design add-ons that use Alibre Script/AlibreX (IronPython 2 or 3) as commands. 

###  Add-on Developer Kit (ADK)

This add-on is part of the alibre-script-adk project, an effort to share lessons learned and provide public resources for modern Alibre Design scripting and programming. Alibre's built-in scripting add-on does not provide a solution for running scripts outside of the add-on. The ADK aims to solve this limitation.

## Who is this for

Anyone who would like to build an Alibre Design add-on, with or without Alibre Script (IronPython) code. 

## What it does

After installation, you'll see a menu and/or ribbon button added to the Alibre Design user interface. Clicking the button will run code that opens the Eymeric65/py-gear tool. 

## How it works

Scripts are saved along with the required add-on files. The add-on loads and runs the .py file with the IronPython scripting engine. The exact process can vary. In your add-on, you can use the Alibre Script add-on library (API) and AlibreX from IronPython. As an add-on, you have full control over all aspects of the process.

## Known Issues

N/A

## Installation

See Releases for the installer and portable .zip file.

## Documentation

N/A

### Additional Resources

N/A

## Contribution

Contributions to the codebase are not currently accepted, but questions and comments are welcome.

## Acknowledgment and License

MIT — see license.

## Credit & Citation

[Eymeric65/py-gear](https://github.com/Eymeric65/py-gear)

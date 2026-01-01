# NewFileAndFolderShellExtension

A Windows shell extension that adds "New Folder" and "New Text File" options directly to the context menu in Windows 10 / 11.

## Features

- Adds "New Folder" to the context menu
- Adds "New Text File" to the context menu
- Uses the same icons as Windows' built-in "New" submenu
- Automatically selects and renames the newly created item, just like Windows does
- Works in both directory backgrounds and when right-clicking on folders

## Building

1. Make sure you have Visual Studio 2019 or 2022 installed with C++ development tools
2. Run build.bat

## Installation

1. Build the DLL using build.bat
2. Run install.bat as Administrator

The extension will be registered and available in Explorer. You may need to restart Explorer for it to take effect.

## Uninstallation

Run uninstall.bat as Administrator

## Details

This is a COM-based shell extension that implements:
- IShellExtInit - for initialization with the folder context
- IContextMenu - for adding menu items and handling commands

The extension registers itself as a context menu handler for:
- Directory backgrounds (Directory\Background\shellex\ContextMenuHandlers)
- Directories (Directory\shellex\ContextMenuHandlers)

[![GitHub stars](https://img.shields.io/github/stars/<OWNER>/<REPO>.svg?style=social&label=Stars)](https://github.com/<OWNER>/<REPO>/stargazers)
[![GitHub forks](https://img.shields.io/github/forks/<OWNER>/<REPO>.svg?style=social&label=Forks)](https://github.com/<OWNER>/<REPO>/network/members)
[![GitHub watchers](https://img.shields.io/github/watchers/<OWNER>/<REPO>.svg?style=social&label=Watchers)](https://github.com/<OWNER>/<REPO>/watchers)
[![GitHub followers](https://img.shields.io/github/followers/<USERNAME>.svg?style=social&label=Followers)](https://github.com/<USERNAME>/?tab=followers)

# CorelDRAW VBA Productivity Tools

A collection of VBA macros and functions to enhance productivity when using CorelDRAW.

## Features

    SaveAndCleanup Module:
        * SaveAsLowerVersion function: Saves the active or specified CorelDRAW document as a lower version format.
        * SaveAllAsLowerVersion sub: Saves all CorelDRAW documents in a specified folder as a lower version format.
        * SaveActiveDocAsLowerVersion sub: Saves the active CorelDRAW document as a lower version format.
        * DeleteBackupFiles sub: Deletes all backup files of CorelDRAW documents in the specified folder.

    HelperFunctions Module:
        * DeleteFileToRecycleBin function: Deletes a specified file and sends it to the recycle bin.
        * FileExists function: Determines if a specified file exists.

## Usage

    1. Copy the code for each module to the corresponding module in your CorelDRAW VBA project.
    2. Reference the Microsoft Scripting Runtime library in your project.
    3. Call the macros and functions as needed.

## Example

To use the SaveAsLowerVersion function, run the following code:
`SaveAsLowerVersion "C:\path\to\your\document.cdr", cdrVersion14`

To use the SaveAllAsLowerVersion sub, run the following code:
`SaveAllAsLowerVersion`

To use the DeleteFileToRecycleBin function, run the following code:
`DeleteFileToRecycleBin "C:\path\to\your\file.ext"`

## License

Information about the license used for the project.
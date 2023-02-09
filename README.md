[![GitHub stars](https://img.shields.io/github/stars/Sterh20/coreldraw-vba-productivity-tools.svg?style=social&label=Stars)](https://github.com/Sterh20/coreldraw-vba-productivity-tools/stargazers)
[![GitHub forks](https://img.shields.io/github/forks/Sterh20/coreldraw-vba-productivity-tools.svg?style=social&label=Forks)](https://github.com/Sterh20/coreldraw-vba-productivity-tools/network/members)
[![GitHub watchers](https://img.shields.io/github/watchers/Sterh20/coreldraw-vba-productivity-tools.svg?style=social&label=Watchers)](https://github.com/Sterh20/coreldraw-vba-productivity-tools/watchers)
[![GitHub followers](https://img.shields.io/github/followers/Sterh20.svg?style=social&label=Followers)](https://github.com/Sterh20/?tab=followers)

# CorelDRAW VBA Productivity Tools [![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)

A collection of VBA macros and functions to enhance productivity when using CorelDRAW.

## Features

**SaveAndCleanup Module**:
<ul style="list-style-type:square;">
<li><b>SaveAsLowerVersion</b> function: Saves the active or specified CorelDRAW document as an earlier version format.</li>
<li><b>SaveAllAsLowerVersion</b> sub: Saves all CorelDRAW documents in an active document's folder as an earlier version format.</li>
<li><b>SaveActiveDocAsLowerVersion</b> sub: Saves the active CorelDRAW document as an earlier version format.</li>
<li><b>DeleteBackupFiles</b> sub: Deletes all backup files of CorelDRAW documents in an active document's folder.</li>
</ul>

**HelperFunctions Module**:
<ul style="list-style-type:square;">
<li><b>DeleteFileToRecycleBin</b> function: Deletes a specified file and sends it to the recycle bin.</li>
<li><b>FileExists</b> function: Determines if a specified file exists.</li>
</ul>

## Usage

1. Copy the code for each module to the corresponding module in your CorelDRAW VBA project.
2. Reference the **Microsoft Scripting Runtime library** in your project.
3. Call the macros and functions as needed.

## Example

To use the SaveAsLowerVersion function, run the following code:
`SaveAsLowerVersion "C:\path\to\your\document.cdr", cdrVersion14`

To use the SaveAllAsLowerVersion sub, run the following code:
`SaveAllAsLowerVersion`

To use the DeleteFileToRecycleBin function, run the following code:
`DeleteFileToRecycleBin "C:\path\to\your\file.ext"`

## License

Distributed under the MIT license. See [LICENSE](https://github.com/Sterh20/coreldraw-vba-productivity-tools/blob/main/LICENSE.txt) for more information
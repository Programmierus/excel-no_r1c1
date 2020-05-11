# No R1C1 (Excel Add-In)

This simple Excel Add-In silently enforces A1 reference style in all Excel Documents, every time it saves and/or opens them.

## Problem

Excel has two reference styles: A1 - the default and common one and R1C1 - an alternative one. Even though there is a UI option to deactivate the R1C1 reference style, Excel saves this parameter in files and templates. So also, if you carefully save all your workbooks with a regular A1 reference style, you can always receive a file with R1C1 on. And due to strange handling of this option, that furthermore differs from version to version, it sometimes sticks to templates/defaults from a single document and then spreads across all your documents like a worm.

Regular users are often disoriented when they see a document with the R1C1 reference style activated. It just doesn't match their habits. The problem may become even more sensitive when you have macros that are targeting fields by A1 reference (which is very common) - they will simply not work.

## Solution

`Excel_NoR1C1.xlam` is an Excel Add-In that, when activated, silently turns off R1C1 reference style option on the following events:

1. Workbook opened
2. Workbook prepared for save
3. Workbook saved

As a result, even if you open a document that has R1C1 on, it will be deactivated, and you'll see your standard A1 references. Also, if you will activate R1C1 manually, it'll be turned off as soon as you save the document. So it is theoretically not at all possible, if the Add-In is activated, to save a workbook with R1C1 on.

## Installation

As a regular Add-In, it can be activated in Excel as usual (e.g., Developer Tab -> Excel Add-Ins).

Alternatively, to keep the Add-In invisible from the user, you can just place it in the following folder: `%APPDATA%\Microsoft\Excel\XLSTART`

This opens a neat possibility to install it through a Group Policy in a Domain Environment:

* User Preferences/Windows Settings/Files (_Create a new object_)
  * Source: network location of Add-In (e.g., `\\DOMAIN.INTERNAL\NETLOGON\NoR1C1.xlam`)
  * Target: `%APPDATA%\Microsoft\Excel\XLSTART\NoR1C1.xlam`
  * Action: **Update**
  * **Enable** "Run in logged-on user's security context (user policy option)"

**Tip:** [Office ADMX templates](https://www.microsoft.com/en-us/download/details.aspx?id=49030) have a built-in policy to control R1C1 defaults (User Policies/Administrative Templates/Microsoft Excel 2016/Excel Options/Formulas/R1C1 reference style). It doesn't affect existing files and templates at all, though.

## Limitations

It is by design that there are no questions asked, and there is no way to make any exclusions. This was required in my particular usage case. If you need some sort of conditional behavior, feel free to adjust the Add-In code.

Tested with Office 2016 and 2019.
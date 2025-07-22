# wxAutoExcel FAQ

## Table of Contents

- How do I start with wxAutoExcel?
- Which licence is wxAutoExcel released under?
- Can I use wxAutoExcel without wxWidgets?
- Can I use wxAutoExcel to manipulate Microsfot Excel files
  on a computer without Microsoft Excel installed?
- wxWidgets is a multiplatfrom toolkit - can I use wxAutoExcel to automate Microsoft Excel on OSX?
- Which versions of Microsoft Excel are supported?
- Does wxAutoExcel support all the features Microsoft Excel exposes?
- I have multiple versions of Microst Excel installed - can I choose which one to use for automation?
- Is it possible to read from / write to a two-dimensional Range?
- How can I use wxAutoExcel with Microsoft Excel localized into languages other than English?
- How can I work with Excel objects not implemented in wxAutoExcel?
- It does not work?!


## How do I start with wxAutoExcel?

1. Follow the install instructions in docs\install.md.
2. Read the [tutorial](https://pbfordev.github.io/wxAutoExcel/page_tutorial.html) and then take a look at the bundled  samples to familiarize yourself with wxAutoExcel, start with *Minimal* sample.

It is assumed you are familiar with VBA and Microsoft Excel object model,
as wxAutoExcel is just its C++ wrapper. If you are not very familiar with 
those, I believe that Microsoft Excel's Macro Recorder can be a very useful tool. 
When I want to do something in the code and I am not sure how, I usually 
just record a macro and then "translate" it from VBA to wxAutoExcel.


## Which licence is wxAutoExcel released under?

wxAutoExcel uses the MIT license.


## Can I use wxAutoExcel without wxWidgets?

No, wxAutoExcel requires the rudimentary OLE Automation support 
as well as some generic classes provided by wxWidgets. However,
it may be possible to use wxAutoExcel in an application which
does not use wxWidgets as its GUI toolkit, take a look at 
*PureWin32* sample for an inspiration.

## Can I use wxAutoExcel to manipulate Microsfot Excel file on a computer without Microsoft Excel installed?

No, wxAutoExcel uses OLE Automation which requires Microsoft Excel
to be installed and working properly.


## wxWidgets is a multiplatfrom toolkit - can I use wxAutoExcel to automate Microsoft Excel on OSX?

No, OLE Automation is available only on Microsoft Windows.


## Which versions of Microsoft Excel are supported?

wxAutoExcel was initially developed along Microsoft Excel 2007. Generally speaking,
wxAutoExcel can be used at least with Excel 2003-2021 / Office 365. Obviously, one 
needs to be careful to use only those objects, methods, and properties that are available for the installed Excel version. `wxExcelApplication` has a couple of version methods, such as `GetVersionAsDouble()`  or `IsVersionAtLeast_()`, it is also possible to obtain the list of properties and methods for a given object, see 
`wxAutoExcelObject::GetPropertyAndMethodNames_()`.


## Does wxAutoExcel support all the features Microsoft Excel exposes?

No, it does not come even close, see the class list in the HTML 
documentation for what is available. Even when a class is listed there,
some of its methods or properties may not be implemented.
Events are not supported at all.

## I have multiple versions of Microst Excel installed - can I chose which one use for automation?

Unfortunately, this is not possible (without modifying Registry).


## Is it possible to read from / write to a two-dimensional Range?

Yes, using `wxVariantDataSafeArray`, please see *BulkData* sample.


## How can I use wxAutoExcel with Microsoft Excel localized into languages other than English?

I suggest calling `SetAutomationLCID_(wxExcelObject::lcidEnglishUS)` on wxAutoExcel objects. The LCID is propagated towards the child wxAutoExcel objects, e.g., when you call it on `wxExcelApplication`, all of its properties like workbooks
will have the LCID set, all worksheets of the workbooks etc.
This way you can use English names of formulas, styles etc. regardless of
the language Microsoft Excel is localized into. This unfortunately does not
work in several cases such as formulas in conditional formatting and validation. 
Microsoft Excel also seems getting worse at it, e.g., since version 2013 you cannot
use English `NumberFormat` string for `TickLabels` in charts. 
Be aware that the automation locale affects how the values (e.g., those
including decimal or thousand separators, formatted as dates etc.) are 
interpreted in Microsoft Excel. E.g., string "1,234" can be considered an integer 
if the locale has "," as the thousand separator but also as "1.234" if "," is the
decimal separator. 


## How can I work with Excel objects not implemented in wxAutoExcel?

Unless you are willing to add their support to wxAutoExcel or can persuade me
to do that myself, you will have to resort to pure `wxAutomationObject` calls. 
The following example demonstrates a hypothetical situation where 
Worksheet and Worksheets objects were not implemented by wxAutoExcel 
and you wanted to output the name of each worksheet in a workbook:

       // workbook must be a valid wxExcelWorkbook object
       wxAutomationObject worksheets;
    
        if ( !workbook.GetUnimplementedObject_("Worksheets", worksheets) )
            return;
        
        const long count = worksheets.GetProperty("Count");
        for ( long i = 1; i <= count; i++ )
        {
            wxAutomationObject worksheet;
    
            if ( wxExcelObject::GetUnimplementedCollectionItem_(worksheets, i, worksheet, true) )
            {            
                wxLogMessage(_("Name of the worksheet with index %ld is \"%s\"."), 
                    i, (wxString)worksheet.GetProperty("Name"));        
            }    
        }   


## It does not work?!

Please be always as specific as possible when reporting an issue.
If it is build-related, make sure you followed the install instructions.
Provide all information required to successfully resolve the issue, 
including the error messages.
If your wxAutoExcel code does not work as expected, please:
1. Make the shortest possible self-contained compilable example demonstrating
   the issue, preferably based on the most appropriate bundled sample.
2. Create and test the VBA equivalent of the code from (1) to make sure the issue
   is within wxAutoExcel and not your code or Microsoft Excel itself.
3. Post the issue in wxCode section of wxWidgets forums (https://forums.wxwidgets.org/viewforum.php?f=30) or report it on GitHub (https://github.com/pbfordev/wxAutoExcel/issues).

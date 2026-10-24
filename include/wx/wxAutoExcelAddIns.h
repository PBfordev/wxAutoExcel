/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2016 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_ADDINS_H
#define _WXAUTOEXCEL_ADDINS_H

#include "wx/wxAutoExcel_defs.h"

#include "wx/wxAutoExcel_object.h"


namespace wxAutoExcel {

    /**
    @brief Represents a single add-in, either installed or not installed.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelAddIn: public wxExcelObject
   {
   public:
        // ***** PROPERTIES *****

        /**
        Returns a read-only unique identifier, or CLSID, identifying an object, as a String.

        [MSDN documentation for AddIn.CLSID](https://learn.microsoft.com/en-us/office/vba/api/excel.addin.clsid).aspx).
        */
        wxString GetCLSID();

        /**
        Returns the name of the object, including its path on disk, as a string.

        [MSDN documentation for AddIn.FullName](https://learn.microsoft.com/en-us/office/vba/api/excel.addin.fullname).aspx).
        */
        wxString GetFullName();

        /**
        True if the add-in is installed or to install the add-in, False if the add-in is uninstalled or to uninstall the add-in.

        [MSDN documentation for AddIn.Installed](https://learn.microsoft.com/en-us/office/vba/api/excel.addin.installed).aspx).
        */
        bool GetInstalled();

        /**
        True if the add-in is installed or to install the add-in, False if the add-in is uninstalled or to uninstall the add-in.

        [MSDN documentation for AddIn.Installed](https://learn.microsoft.com/en-us/office/vba/api/excel.addin.installed).aspx).
        */
        void SetInstalled(bool installed);

        /**
        Returns True if the add-in is currently open. Since Excel 2007.

        [MSDN documentation for AddIn.IsOpen](https://learn.microsoft.com/en-us/office/vba/api/excel.addin.isopen).aspx).
        */
        bool GetIsOpen();

        /**
        Returns a String value that represents the name of the object.

        [MSDN documentation for AddIn.Name](https://learn.microsoft.com/en-us/office/vba/api/excel.addin.name).aspx).
        */
        wxString GetName();

        /**
        Returns a String value that represents the complete path to the Add-in, excluding the final separator and name of the Add-in.

        [MSDN documentation for AddIn.Path](https://learn.microsoft.com/en-us/office/vba/api/excel.addin.path).aspx).
        */
        wxString GetPath();

        /**
        Returns the programmatic identifiers for the object.

        [MSDN documentation for AddIn.progID](https://learn.microsoft.com/en-us/office/vba/api/excel.addin.progid).aspx).
        */
        wxString GetprogID();


        /**
        Returns "AddIn".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("AddIn"); }

    };

   /**
      @brief This class cannot be instantiated, it has to be used either as wxExcelAddIns or wxExcelAddIns2
   */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelAddInsBase : public wxExcelObject
   {
   public:
       // ***** METHODS *****

       /**
       Adds a new add-in file to the list of add-ins. Returns an AddIn object.

       [MSDN documentation for AddIns.Add](https://learn.microsoft.com/en-us/office/vba/api/excel.addins.add).aspx).
       */
       wxExcelAddIn Add(const wxString& fileName, wxXlTribool copyFile = wxDefaultXlTribool);

       // ***** PROPERTIES *****

       /**
       Returns a Long value that represents the number of objects in the collection.

       [MSDN documentation for AddIns.Count](https://learn.microsoft.com/en-us/office/vba/api/excel.addins.count).aspx).
       */
       long GetCount();

       //@{
       /**
       Returns a single object from a collection.

       [MSDN documentation for AddIns.Item](https://learn.microsoft.com/en-us/office/vba/api/excel.addins.item).aspx).
       */
       wxExcelAddIn GetItem(long index);
       wxExcelAddIn operator[](long);
       wxExcelAddIn GetItem(const wxString& name);
       wxExcelAddIn operator[](const wxString& name);
       //@}

       /**
       The iterator over the items, see wxExcelCollectionIterator.
       */
       typedef wxExcelCollectionIterator<wxExcelAddInsBase, wxExcelAddIn> iterator;

       //@{
       /**
       Returns the iterator referring to the first item, or past the last item,
       which allows iterating over the items with a range-based for loop,
       see wxExcelCollectionIterator.
       */
       iterator begin() { return iterator(this, 1); }
       iterator end() { return iterator(this, GetCount() + 1); }
       //@}

       /**
       Returns "AddInsBase".
       */
       virtual wxString GetAutoExcelObjectName_() const { return wxS("AddInsBase"); }
   };

   /**
      @brief A collection of AddIn objects that represents all the add-ins available to Microsoft Excel, regardless of whether they're installed.

   This list corresponds to the list of add-ins displayed in the Add-Ins dialog box.
      */

   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelAddIns : public wxExcelAddInsBase
   {
   public:
       /**
       Returns "AddIns".
       */
       virtual wxString GetAutoExcelObjectName_() const { return wxS("AddIns"); }
   };


   /**
   @brief A collection of AddIn objects that represent all the add-ins that are currently available or open in Microsoft Excel, regardless of whether they are installed.

   Since Excel 2010.
   The contents of the AddIns2 collection correspond to the list of add-ins displayed in the Add-Ins dialog box (Add-Ins command on the Developer tab) and any add-ins that are currently open.
   */

   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelAddIns2 : public wxExcelAddInsBase
   {
   public:
       /**
       Returns "AddIns2".
       */
       virtual wxString GetAutoExcelObjectName_() const { return wxS("AddIns2"); }
   };

} // namespace wxAutoExcel

#endif // _WXAUTOEXCEL_ADDINS_H

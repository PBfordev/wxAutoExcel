/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_HYPERLINKS_H
#define _WXAUTOEXCEL_HYPERLINKS_H

#include "wx/wxAutoExcel_defs.h"

#include "wx/wxAutoExcel_object.h"

namespace wxAutoExcel {

    /**
    @brief Represents Microsoft Excel Hyperlink object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelHyperlink : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Adds a shortcut to the workbook or hyperlink to the Favorites folder.

        [MSDN documentation for Hyperlink.AddToFavorites](https://learn.microsoft.com/en-us/office/vba/api/excel.hyperlink.addtofavorites).
        */
        void AddToFavorites();

        /**
        Creates a new document linked to the specified hyperlink.

        [MSDN documentation for Hyperlink.CreateNewDocument](https://learn.microsoft.com/en-us/office/vba/api/excel.hyperlink.createnewdocument).
        */
        void CreateNewDocument(const wxString& fileName, bool editNow, bool overwrite);

        /**
        Deletes the object.

        [MSDN documentation for Hyperlink.Delete](https://learn.microsoft.com/en-us/office/vba/api/excel.hyperlink.delete).
        */
        void Delete();

        /**
        Displays a cached document, if it’s already been downloaded. Otherwise, this method resolves the hyperlink, downloads the target document, and displays the document in the appropriate application.

        [MSDN documentation for Hyperlink.Follow](https://learn.microsoft.com/en-us/office/vba/api/excel.hyperlink.follow).
        */
        void Follow(wxXlTribool newWindow = wxDefaultXlTribool, wxXlTribool addHistory = wxDefaultXlTribool,
                    MsoExtraInfoMethod* method = NULL, const wxString& headerInfo = wxEmptyString);

        // ***** PROPERTIES *****

        /**
        Returns a String value that represents the address of the target document.

        [MSDN documentation for Hyperlink.Address](https://learn.microsoft.com/en-us/office/vba/api/excel.hyperlink.address).
        */
        wxString GetAddress();

        /**
        Sets a String value that represents the address of the target document.

        [MSDN documentation for Hyperlink.Address](https://learn.microsoft.com/en-us/office/vba/api/excel.hyperlink.address).
        */
        void SetAddress(const wxString& address);


        /**
        Returns the text string of the specified hyperlink’s e-mail subject line. The subject line is appended to the hyperlink’s address.

        [MSDN documentation for Hyperlink.EmailSubject](https://learn.microsoft.com/en-us/office/vba/api/excel.hyperlink.emailsubject).
        */
        wxString GetEmailSubject();

        /**
        Sets the text string of the specified hyperlink’s e-mail subject line. The subject line is appended to the hyperlink’s address.

        [MSDN documentation for Hyperlink.EmailSubject](https://learn.microsoft.com/en-us/office/vba/api/excel.hyperlink.emailsubject).
        */
        void SetEmailSubject(const wxString& emailSubject);

        /**
        Returns a String value that represents the name of the object.

        [MSDN documentation for Hyperlink.Name](https://learn.microsoft.com/en-us/office/vba/api/excel.hyperlink.name).
        */
        wxString GetName();


        /**
        Returns a Range Represents the range the specified hyperlink is attached to.

        [MSDN documentation for Hyperlink.Range](https://learn.microsoft.com/en-us/office/vba/api/excel.hyperlink.range).
        */
        wxExcelRange GetRange();

        /**
        Returns the ScreenTip text for the specified hyperlink.

        [MSDN documentation for Hyperlink.ScreenTip](https://learn.microsoft.com/en-us/office/vba/api/excel.hyperlink.screentip).
        */
        wxString GetScreenTip();

        /**
        Sets the ScreenTip text for the specified hyperlink.

        [MSDN documentation for Hyperlink.ScreenTip](https://learn.microsoft.com/en-us/office/vba/api/excel.hyperlink.screentip).
        */
        void SetScreenTip(const wxString& screenTip);

#if WXAUTOEXCEL_USE_SHAPES
        /**
        Returns a Shape Represents the shape attached to the specified hyperlink.

        [MSDN documentation for Hyperlink.Shape](https://learn.microsoft.com/en-us/office/vba/api/excel.hyperlink.shape).
        */
        wxExcelShape GetShape();
#endif // #if WXAUTOEXCEL_USE_SHAPES

        /**
        Returns the location within the document associated with the hyperlink.

        [MSDN documentation for Hyperlink.SubAddress](https://learn.microsoft.com/en-us/office/vba/api/excel.hyperlink.subaddress).
        */
        wxString GetSubAddress();

        /**
        Sets the location within the document associated with the hyperlink.

        [MSDN documentation for Hyperlink.SubAddress](https://learn.microsoft.com/en-us/office/vba/api/excel.hyperlink.subaddress).
        */
        void SetSubAddress(const wxString& subAddress);

        /**
        Returns the text to be displayed for the specified hyperlink. The default value is the address of the hyperlink.

        [MSDN documentation for Hyperlink.TextToDisplay](https://learn.microsoft.com/en-us/office/vba/api/excel.hyperlink.texttodisplay).
        */
        wxString GetTextToDisplay();

        /**
        Sets the text to be displayed for the specified hyperlink. The default value is the address of the hyperlink.

        [MSDN documentation for Hyperlink.TextToDisplay](https://learn.microsoft.com/en-us/office/vba/api/excel.hyperlink.texttodisplay).
        */
        void SetTextToDisplay(const wxString& textToDisplay);

        /**
        Returns a Long value, containing a MsoHyperlinkType constant, that represents the location of the HTML frame.

        [MSDN documentation for Hyperlink.Type](https://learn.microsoft.com/en-us/office/vba/api/excel.hyperlink.type).
        */
        MsoHyperlinkType GetType();

        /**
        Returns "Hyperlink".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("Hyperlink"); }
    };


    /**
    @brief Represents Microsoft Excel Hyperlinks collection.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelHyperlinks : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Adds a hyperlink to the specified range or shape.

        [MSDN documentation for Hyperlinks.Add](https://learn.microsoft.com/en-us/office/vba/api/excel.hyperlinks.add).
        */
        wxExcelHyperlink Add(wxExcelObject* anchor, const wxString& address, const wxString& subAddress = wxEmptyString,
                             const wxString& screenTip = wxEmptyString, const wxString& textToDisplay  = wxEmptyString);

        /**
        Deletes the object.

        [MSDN documentation for Hyperlinks.Delete](https://learn.microsoft.com/en-us/office/vba/api/excel.hyperlinks.delete).
        */
        void Delete();

        // ***** PROPERTIES *****

        /**
        Returns a Long value that represents the number of objects in the collection.

        [MSDN documentation for Hyperlinks.Count](https://learn.microsoft.com/en-us/office/vba/api/excel.hyperlinks.count).
        */
        long GetCount();

        //@{
        /**
        Returns a single object from a collection.

        [MSDN documentation for Hyperlinks.Item](https://learn.microsoft.com/en-us/office/vba/api/excel.hyperlinks.item).
        */
        wxExcelHyperlink GetItem(long index);
        wxExcelHyperlink operator[](long index);

        //@}

        /**
        Returns "Hyperlinks".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("Hyperlinks"); }
    };


} // namespace wxAutoExcel

#endif //_WXAUTOEXCEL_HYPERLINKS_H

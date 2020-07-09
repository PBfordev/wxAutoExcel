/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
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

        [MSDN documentation for Hyperlink.AddToFavorites](http://msdn.microsoft.com/en-us/library/bb211810).
        */
        void AddToFavorites();

        /**
        Creates a new document linked to the specified hyperlink.

        [MSDN documentation for Hyperlink.CreateNewDocument](http://msdn.microsoft.com/en-us/library/bb223296).
        */
        void CreateNewDocument(const wxString& fileName, bool editNow, bool overwrite);

        /**
        Deletes the object.

        [MSDN documentation for Hyperlink.Delete](http://msdn.microsoft.com/en-us/library/bb211816).
        */
        void Delete();

        /**
        Displays a cached document, if it’s already been downloaded. Otherwise, this method resolves the hyperlink, downloads the target document, and displays the document in the appropriate application.

        [MSDN documentation for Hyperlink.Follow](http://msdn.microsoft.com/en-us/library/bb209868).
        */
        void Follow(wxXlTribool newWindow = wxDefaultXlTribool, wxXlTribool addHistory = wxDefaultXlTribool,
                    MsoExtraInfoMethod* method = NULL, const wxString& headerInfo = wxEmptyString);

        // ***** PROPERTIES *****

        /**
        Returns a String value that represents the address of the target document.

        [MSDN documentation for Hyperlink.Address](http://msdn.microsoft.com/en-us/library/bb148518).
        */
        wxString GetAddress();

        /**
        Sets a String value that represents the address of the target document.

        [MSDN documentation for Hyperlink.Address](http://msdn.microsoft.com/en-us/library/bb148518).
        */
        void SetAddress(const wxString& address);


        /**
        Returns the text string of the specified hyperlink’s e-mail subject line. The subject line is appended to the hyperlink’s address.

        [MSDN documentation for Hyperlink.EmailSubject](http://msdn.microsoft.com/en-us/library/bb221092).
        */
        wxString GetEmailSubject();

        /**
        Sets the text string of the specified hyperlink’s e-mail subject line. The subject line is appended to the hyperlink’s address.

        [MSDN documentation for Hyperlink.EmailSubject](http://msdn.microsoft.com/en-us/library/bb221092).
        */
        void SetEmailSubject(const wxString& emailSubject);

        /**
        Returns a String value that represents the name of the object.

        [MSDN documentation for Hyperlink.Name](http://msdn.microsoft.com/en-us/library/bb148519).
        */
        wxString GetName();


        /**
        Returns a Range Represents the range the specified hyperlink is attached to.

        [MSDN documentation for Hyperlink.Range](http://msdn.microsoft.com/en-us/library/bb148520).
        */
        wxExcelRange GetRange();

        /**
        Returns the ScreenTip text for the specified hyperlink.

        [MSDN documentation for Hyperlink.ScreenTip](http://msdn.microsoft.com/en-us/library/bb221604).
        */
        wxString GetScreenTip();

        /**
        Sets the ScreenTip text for the specified hyperlink.

        [MSDN documentation for Hyperlink.ScreenTip](http://msdn.microsoft.com/en-us/library/bb221604).
        */
        void SetScreenTip(const wxString& screenTip);

#if WXAUTOEXCEL_USE_SHAPES
        /**
        Returns a Shape Represents the shape attached to the specified hyperlink.

        [MSDN documentation for Hyperlink.Shape](http://msdn.microsoft.com/en-us/library/bb214629).
        */
        wxExcelShape GetShape();
#endif // #if WXAUTOEXCEL_USE_SHAPES

        /**
        Returns the location within the document associated with the hyperlink.

        [MSDN documentation for Hyperlink.SubAddress](http://msdn.microsoft.com/en-us/library/bb209307).
        */
        wxString GetSubAddress();

        /**
        Sets the location within the document associated with the hyperlink.

        [MSDN documentation for Hyperlink.SubAddress](http://msdn.microsoft.com/en-us/library/bb209307).
        */
        void SetSubAddress(const wxString& subAddress);

        /**
        Returns the text to be displayed for the specified hyperlink. The default value is the address of the hyperlink.

        [MSDN documentation for Hyperlink.TextToDisplay](http://msdn.microsoft.com/en-us/library/bb221802).
        */
        wxString GetTextToDisplay();

        /**
        Sets the text to be displayed for the specified hyperlink. The default value is the address of the hyperlink.

        [MSDN documentation for Hyperlink.TextToDisplay](http://msdn.microsoft.com/en-us/library/bb221802).
        */
        void SetTextToDisplay(const wxString& textToDisplay);

        /**
        Returns a Long value, containing a MsoHyperlinkType constant, that represents the location of the HTML frame.

        [MSDN documentation for Hyperlink.Type](http://msdn.microsoft.com/en-us/library/bb214631).
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

        [MSDN documentation for Hyperlinks.Add](http://msdn.microsoft.com/en-us/library/bb211819).
        */
        wxExcelHyperlink Add(wxExcelObject* anchor, const wxString& address, const wxString& subAddress = wxEmptyString,
                             const wxString& screenTip = wxEmptyString, const wxString& textToDisplay  = wxEmptyString);

        /**
        Deletes the object.

        [MSDN documentation for Hyperlinks.Delete](http://msdn.microsoft.com/en-us/library/bb211823).
        */
        void Delete();

        // ***** PROPERTIES *****

        /**
        Returns a Long value that represents the number of objects in the collection.

        [MSDN documentation for Hyperlinks.Count](http://msdn.microsoft.com/en-us/library/bb148522).
        */
        long GetCount();

        //@{
        /**
        Returns a single object from a collection.

        [MSDN documentation for Hyperlinks.Item](http://msdn.microsoft.com/en-us/library/bb148524).
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

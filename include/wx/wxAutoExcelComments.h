/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_COMMENTS_H
#define _WXAUTOEXCEL_COMMENTS_H

#include "wx/wxAutoExcel_defs.h"
#include "wx/wxAutoExcel_object.h"

namespace wxAutoExcel {


    /**
    @brief Represents Microsoft Excel Comment object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelComment: public wxExcelObject
    {
    public:

        /**
        Deletes the object.

        [MSDN documentation for Comment.Delete](http://msdn.microsoft.com/en-us/library/bb211701).
        */
        void Delete();

        /**
        Returns the next comment.

        [MSDN documentation for Comment.Next](http://msdn.microsoft.com/en-us/library/bb242024).
        */
        wxExcelComment Next();

        /**
        Returns the previous comment.

        [MSDN documentation for Comment.Previous](http://msdn.microsoft.com/en-us/library/bb242030).
        */
        wxExcelComment Previous();

        /**
        Sets comment text.

        [MSDN documentation for Comment.Text](http://msdn.microsoft.com/en-us/library/bb237810).
        */
        wxString Text(const wxString& text = wxEmptyString, long* start = NULL, wxXlTribool overwrite = wxDefaultXlTribool);

        // ***** PROPERTIES *****

        /**
        Returns the author of the comment.

        [MSDN documentation for Comment.Author](http://msdn.microsoft.com/en-us/library/bb220848).
        */
        wxString GetAuthor();

        /**
        Sets the author of the comment.

        [MSDN documentation for Comment.Author](http://msdn.microsoft.com/en-us/library/bb220848).
        */
        void SetAuthor(const wxString& author);

#if WXAUTOEXCEL_USE_SHAPES
        /**
        Returns a Shape associated with the comment.

        [MSDN documentation for Comment.Shape](hhttp://msdn.microsoft.com/en-us/library/bb214517).
        */
        wxExcelShape GetShape();
#endif // #if WXAUTOEXCEL_USE_SHAPES

        /**
        True if the comment is visible.

        [MSDN documentation for Comment.Visible](http://msdn.microsoft.com/en-us/library/bb214523).
        */
        bool GetVisible();

        /**
        True if the comment is visible.

        [MSDN documentation for Comment.Visible](http://msdn.microsoft.com/en-us/library/bb214523).
        */
        void SetVisible(bool visible);


        /**
        Returns "Comment".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("Comment"); }
    };

    /**
    @brief Represents Microsoft Excel Comments collection.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelComments: public wxExcelObject
    {
    public:
        /**
        Returns a comment.

        [MSDN documentation for Comments.Item](http://msdn.microsoft.com/en-us/library/bb211707).
        */
        wxExcelComment GetItem(long index);
        wxExcelComment operator[](long index);

        // ***** PROPERTIES *****

        /**
        Returns a number of comments in teh collection.

        [MSDN documentation for Comments.Count](http://msdn.microsoft.com/en-us/library/bb179518).
        */
        long GetCount();

        /**
        Returns "Comments".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("Comments"); }
    };


} // namespace wxAutoExcel

#endif //_WXAUTOEXCEL_COMMENTS_H

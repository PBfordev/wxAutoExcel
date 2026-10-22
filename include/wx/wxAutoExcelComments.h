/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
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

        [MSDN documentation for Comment.Delete](https://learn.microsoft.com/en-us/office/vba/api/excel.comment.delete).
        */
        void Delete();

        /**
        Returns the next comment.

        [MSDN documentation for Comment.Next](https://learn.microsoft.com/en-us/office/vba/api/excel.comment.next).
        */
        wxExcelComment Next();

        /**
        Returns the previous comment.

        [MSDN documentation for Comment.Previous](https://learn.microsoft.com/en-us/office/vba/api/excel.comment.previous).
        */
        wxExcelComment Previous();

        /**
        Sets comment text.

        [MSDN documentation for Comment.Text](https://learn.microsoft.com/en-us/office/vba/api/excel.comment.text).
        */
        wxString Text(const wxString& text = wxEmptyString, long* start = NULL, wxXlTribool overwrite = wxDefaultXlTribool);

        // ***** PROPERTIES *****

        /**
        Returns the author of the comment.

        [MSDN documentation for Comment.Author](https://learn.microsoft.com/en-us/office/vba/api/excel.comment.author).
        */
        wxString GetAuthor();

#if WXAUTOEXCEL_USE_SHAPES
        /**
        Returns a Shape associated with the comment.

        [MSDN documentation for Comment.Shape](https://learn.microsoft.com/en-us/office/vba/api/excel.comment.shape).
        */
        wxExcelShape GetShape();
#endif // #if WXAUTOEXCEL_USE_SHAPES

        /**
        True if the comment is visible.

        [MSDN documentation for Comment.Visible](https://learn.microsoft.com/en-us/office/vba/api/excel.comment.visible).
        */
        bool GetVisible();

        /**
        True if the comment is visible.

        [MSDN documentation for Comment.Visible](https://learn.microsoft.com/en-us/office/vba/api/excel.comment.visible).
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

        [MSDN documentation for Comments.Item](https://learn.microsoft.com/en-us/office/vba/api/excel.comments.item).
        */
        wxExcelComment GetItem(long index);
        wxExcelComment operator[](long index);

        // ***** PROPERTIES *****

        /**
        Returns a number of comments in teh collection.

        [MSDN documentation for Comments.Count](https://learn.microsoft.com/en-us/office/vba/api/excel.comments.count).
        */
        long GetCount();

        /**
        Returns "Comments".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("Comments"); }
    };


} // namespace wxAutoExcel

#endif //_WXAUTOEXCEL_COMMENTS_H

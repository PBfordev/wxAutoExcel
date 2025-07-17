/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2020 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////

#ifndef _WXAUTOEXCEL_COMMENTSTHREADED_H
#define _WXAUTOEXCEL_COMMENTSTHREADED_H

#include "wx/wxAutoExcel_defs.h"

#include "wx/wxAutoExcel_object.h"

namespace wxAutoExcel
{

/**
@brief Represents a cell's threaded comment. This object can represent both a top-level comment or its replies.
*/
class WXDLLIMPEXP_WXAUTOEXCEL wxExcelCommentThreaded : public wxExcelObject
{
public:
    // ***** METHODS *****

    /**
    If the comment is a top-level comment, it will add a reply to its replies collection. If this comment is a reply, it will add a reply to its Parent's replies collection.

    [Excel VBA documentation for CommentThreaded.AddReply](https://docs.microsoft.com/en-us/office/vba/api/excel.commentthreaded.addreply)
    */
    wxExcelCommentThreaded AddReply(const wxString& text = wxEmptyString);

    /**
    Deletes the specified threaded comment and all replies associated with that comment (if any exist).

    [Excel VBA documentation for CommentThreaded.Delete](https://docs.microsoft.com/en-us/office/vba/api/excel.commentthreaded.delete)
    */
    void Delete();

    /**
    Returns a CommentThreaded object that represents the next threaded comment.

    [Excel VBA documentation for CommentThreaded.Next](https://docs.microsoft.com/en-us/office/vba/api/excel.commentthreaded.next)
    */
    wxExcelCommentThreaded Next();

    /**
    Returns a CommentThreaded object that represents the previous threaded comment.

    [Excel VBA documentation for CommentThreaded.Previous](https://docs.microsoft.com/en-us/office/vba/api/excel.commentthreaded.previous)
    */
    wxExcelCommentThreaded Previous();

    /**
    Sets threaded comment text.

    [Excel VBA documentation for CommentThreaded.Text](https://docs.microsoft.com/en-us/office/vba/api/excel.commentthreaded.text)
    */
    wxString Text(const wxString& text = wxEmptyString, long* start = NULL,
                  wxXlTribool overwrite = wxDefaultXlTribool);

    // ***** PROPERTIES *****

    /**
    Returns the Author object that represents the author of the specified CommentThreaded object.

    [Excel VBA documentation for CommentThreaded.Author](https://docs.microsoft.com/en-us/office/vba/api/excel.commentthreaded.author)
    */
    wxExcelAuthor GetAuthor();

    /**
    Returns a date that represents the date and time that a threaded comment was added in local time.

    [Excel VBA documentation for CommentThreaded.Date](https://docs.microsoft.com/en-us/office/vba/api/excel.commentthreaded.date)
    */
    wxVariant GetDate();

    /**
    If this comment is a parent, returns a CommentsThreaded collection of CommentThreaded objects that are children/replies of the specified comment (if any exist). The replies are sorted by time stamp. If this comment is a child/reply or a legacy comment, returns an empty collection.

    [Excel VBA documentation for CommentThreaded.Replies](https://docs.microsoft.com/en-us/office/vba/api/excel.commentthreaded.replies)
    */
    wxExcelCommentsThreaded GetReplies();

    /**
    Returns "CommentThreaded".
    */
    virtual wxString GetAutoExcelObjectName_() const { return wxS("CommentThreaded"); }

}; // class wxExcelCommentThreaded



/**
    @brief Represents a collection of CustomProperty objects that represent additional information. The information can be used as metadata for XML.
*/
class WXDLLIMPEXP_WXAUTOEXCEL wxExcelCommentsThreaded : public wxExcelObject
{
public:
    //@{
    /**
        Returns the CommentThreaded with the given index.

        [MSDN documentation for CommentsThreaded.Item](https://docs.microsoft.com/en-us/office/vba/api/excel.commentsthreaded.item)
    */
    wxExcelCommentThreaded GetItem(long index);
    wxExcelCommentThreaded operator[](long index);
    //@}

    // ***** PROPERTIES *****

    /**
        Returns the number of objects in the collection.
    */
    long GetCount();

    /**
    Returns "CommentsThreaded".
    */
    virtual wxString GetAutoExcelObjectName_() const { return wxS("CommentsThreaded"); }

}; // class wxExcelCommentsThreaded

} // namespace wxAutoExcel

#endif // #ifndef _WXAUTOEXCEL_COMMENTSTHREADED_H

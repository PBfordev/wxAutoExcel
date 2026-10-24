/////////////////////////////////////////////////////////////////////////////
// Author:      Claude Opus 5.5 (Anthropic), under the direction of PB
// Copyright:   (c) 2026 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_ITERATOR_H
#define _WXAUTOEXCEL_ITERATOR_H

#include <iterator>

namespace wxAutoExcel {

    /**
    @brief An iterator over the items of a Microsoft Excel collection.

    The collection classes, such as wxExcelWorksheets, wxExcelShapes, or
    wxExcelNames, as well as wxExcelRange, have begin() and end() methods
    returning this iterator, so that their items can be iterated over with
    a range-based for loop:
    @code
    for ( wxExcelWorksheet worksheet : workbook.GetWorksheets() )
        wxLogMessage("Worksheet '%s'", worksheet.GetName());
    @endcode

    Dereferencing the iterator obtains the item from Excel and returns it
    by value. As wxAutoExcel objects are just references to Excel objects,
    the loop variable should be a value (or auto), not a reference.

    The number of items is obtained by end(), the items are obtained one by
    one, by their index starting from 1. Therefore, the items must not be
    added to or removed from the collection while iterating over it. If an
    item cannot be obtained, it is invalid and the error is reported
    according to the error mode, see wxExcelObject::SetErrorMode_(). When
    the number of items cannot be obtained, there is nothing to iterate over.
    */
    template <class Collection, class Item>
    class wxExcelCollectionIterator
    {
    public:
        typedef std::input_iterator_tag iterator_category;
        typedef Item value_type;
        typedef long difference_type;
        typedef const Item* pointer;
        typedef Item reference;

        /**
        Creates an iterator which does not refer to any collection.
        */
        wxExcelCollectionIterator()
            : m_collection(NULL), m_index(0)
        { }

        /**
        Creates an iterator referring to the item of @a collection at @a index,
        starting from 1.
        */
        wxExcelCollectionIterator(Collection* collection, long index)
            : m_collection(collection), m_index(index)
        { }

        /**
        Returns the item the iterator refers to.
        */
        Item operator*() const
        {
            return (*m_collection)[m_index];
        }

        //@{
        /**
        Advances the iterator to the next item.
        */
        wxExcelCollectionIterator& operator++()
        {
            ++m_index;
            return *this;
        }

        wxExcelCollectionIterator operator++(int)
        {
            wxExcelCollectionIterator previous(*this);

            ++m_index;
            return previous;
        }
        //@}

        //@{
        /**
        Compares the iterators.
        */
        bool operator==(const wxExcelCollectionIterator& other) const
        {
            return m_collection == other.m_collection && m_index == other.m_index;
        }

        bool operator!=(const wxExcelCollectionIterator& other) const
        {
            return !(*this == other);
        }
        //@}

        /**
        Returns the index of the item the iterator refers to, starting from 1.
        */
        long GetIndex() const { return m_index; }

    private:
        Collection* m_collection;
        long m_index;
    };

} // namespace wxAutoExcel

#endif //_WXAUTOEXCEL_ITERATOR_H

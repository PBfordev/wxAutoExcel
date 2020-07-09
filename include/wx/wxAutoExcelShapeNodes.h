/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_SHAPENODES_H
#define _WXAUTOEXCEL_SHAPENODES_H

#include <wx/geometry.h>

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_SHAPES

#include "wx/wxAutoExcel_object.h"
#include "wx/wxAutoExcel_enums.h"

namespace wxAutoExcel {

    /**
    @brief Represents Microsoft Excel ShapeNode object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelShapeNode : public wxExcelObject
    {
    public:
        // ***** PROPERTIES *****

        /**
        If the specified node is a vertex, this property returns a value that indicates how changes made to the node affect the two segments connected to the node. Read-only MsoEditingType.

        [MSDN documentation for ShapeNode.EditingType](http://msdn.microsoft.com/en-us/library/bb221082).
        */
        MsoEditingType GetEditingType();


        /**
        Returns the position of the specified node as a coordinate pair. Each coordinate is expressed in points.

        [MSDN documentation for ShapeNode.Points](http://msdn.microsoft.com/en-us/library/bb242059).
        */
        wxVector<wxPoint2DDouble> GetPoints();

        /**
        Returns a value that indicates whether the segment associated with the specified node is straight or curved. If the specified node is a control point for a curved segment, this property returns msoSegmentCurve. Read-only MsoSegmentType .

        [MSDN documentation for ShapeNode.SegmentType](http://msdn.microsoft.com/en-us/library/bb221633).
        */
        MsoSegmentType  GetSegmentType();

        /**
        Returns "ShapeNode".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("ShapeNode"); }
    };


    /**
    @brief Represents Microsoft Excel ShapeNodes collection.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelShapeNodes : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Deletes the object.

        [MSDN documentation for ShapeNodes.Delete](http://msdn.microsoft.com/en-us/library/bb212168).
        */
        void Delete(long index);

        /**
        Inserts a node into a freeform shape.

        [MSDN documentation for ShapeNodes.Insert](http://msdn.microsoft.com/en-us/library/bb212170).
        */
        void Insert(long index, MsoSegmentType segmentType, MsoEditingType editingType,
                    double X1, double Y1, double* X2 = NULL, double* Y2 = NULL, double* X3 = NULL, double* Y3 = NULL);

        //@{
        /**
        Returns a single object from a collection.

        [MSDN documentation for ShapeNodes.Item](http://msdn.microsoft.com/en-us/library/bb212172).
        */
        wxExcelShapeNode Item(long index);
        wxExcelShapeNode operator[](long index);
        //@}

        /**
        Sets the editing type of the node specified by Index. If the node is a control point for a curved segment, this method sets the editing type of the node adjacent to it that joins two segments. Note that, depending on the editing type, this method may affect the position of adjacent nodes.

        [MSDN documentation for ShapeNodes.SetEditingType](http://msdn.microsoft.com/en-us/library/bb178050).
        */
        void SetEditingType(long index, MsoEditingType editingType);

        /**
        Sets the location of the node specified by Index. Note that, depending on the editing type of the node, this method may affect the position of adjacent nodes.

        [MSDN documentation for ShapeNodes.SetPosition](http://msdn.microsoft.com/en-us/library/bb178075).
        */
        void SetPosition(long index, double X1, double Y1);

        /**
        Sets the segment type of the segment that follows the node specified by Index. If the node is a control point for a curved segment, this method sets the segment type for that curve. Note that this may affect the total number of nodes by inserting or deleting adjacent nodes.

        [MSDN documentation for ShapeNodes.SetSegmentType](http://msdn.microsoft.com/en-us/library/bb178082).
        */
        void SetSegmentType(long index, MsoSegmentType segmentType);

        // ***** PROPERTIES *****


        /**
        Returns an Integer value that represents the number of objects in the collection.

        [MSDN documentation for ShapeNodes.Count](http://msdn.microsoft.com/en-us/library/bb213783).
        */
        long GetCount();

        /**
        Returns "ShapeNodes".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("ShapeNodes"); }
    };


} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_SHAPES

#endif //_WXAUTOEXCEL_SHAPENODES_H

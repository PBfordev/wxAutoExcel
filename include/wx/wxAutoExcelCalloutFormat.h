/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_CALLOUTFORMAT_H
#define _WXAUTOEXCEL_CALLOUTFORMAT_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_SHAPES

#include "wx/wxAutoExcel_object.h"
#include "wx/wxAutoExcel_enums.h"

namespace wxAutoExcel {

    /**
    @brief Represents Microsoft Excel CalloutFormat object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelCalloutFormat : public wxExcelObject
    {
    public:        

        // ***** METHODS *****

        /**
        Specifies that the first segment of the callout line (the segment attached to the text callout box) be scaled automatically when the callout is moved. Use the CustomLength method to specify that the first segment of the callout line retain the fixed length returned by the Length property whenever the callout is moved. Applies only to callouts whose lines consist of more than one segment (types msoCalloutThree and msoCalloutFour).

        [MSDN documentation for CalloutFormat.AutomaticLength](http://msdn.microsoft.com/en-us/library/bb209682).
        */
        void AutomaticLength();

        /**
        Sets the vertical distance (in points) from the edge of the text bounding box to the place where the callout line attaches to the text box. This distance is measured from the top of the text box unless the AutoAttach property is set to True and the text box is to the left of the origin of the callout line (the place that the callout points to), in which case the drop distance is measured from the bottom of the text box.

        [MSDN documentation for CalloutFormat.CustomDrop](http://msdn.microsoft.com/en-us/library/bb223305).
        */
        void CustomDrop(double drop);

        /**
        Specifies that the first segment of the callout line (the segment attached to the text callout box) retain a fixed length whenever the callout is moved. Use the AutomaticLength method to specify that the first segment of the callout line be scaled automatically whenever the callout is moved. Applies only to callouts whose lines consist of more than one segment (types msoCalloutThree and msoCalloutFour).

        [MSDN documentation for CalloutFormat.CustomLength](http://msdn.microsoft.com/en-us/library/bb223308).
        */
        void CustomLength(double length);

        /**
        Specifies whether the callout line attaches to the top, bottom, or center of the callout text box or whether it attaches at a point that’s a specified distance from the top or bottom of the text box.

        [MSDN documentation for CalloutFormat.PresetDrop](http://msdn.microsoft.com/en-us/library/bb223550).
        */
        void PresetDrop(MsoCalloutDropType dropType);

        // ***** PROPERTIES *****

        /**
        Allows the user to place a vertical accent bar to separate the callout text from the callout line. Read/write MsoTriState.

        [MSDN documentation for CalloutFormat.Accent](http://msdn.microsoft.com/en-us/library/bb220815).
        */
        MsoTriState GetAccent();

        /**
        Allows the user to place a vertical accent bar to separate the callout text from the callout line. Read/write MsoTriState.

        [MSDN documentation for CalloutFormat.Accent](http://msdn.microsoft.com/en-us/library/bb220815).
        */
        void SetAccent(MsoTriState accent);

        /**
        Returns the angle of the callout line. If the callout line contains more than one line segment, this property returns or sets the angle of the segment that is farthest from the callout text box. Read/write MsoCalloutAngleType.

        [MSDN documentation for CalloutFormat.Angle](http://msdn.microsoft.com/en-us/library/bb220839).
        */
        MsoCalloutAngleType GetAngle();

        /**
        Sets the angle of the callout line. If the callout line contains more than one line segment, this property returns or sets the angle of the segment that is farthest from the callout text box. Read/write MsoCalloutAngleType.

        [MSDN documentation for CalloutFormat.Angle](http://msdn.microsoft.com/en-us/library/bb220839).
        */
        void SetAngle(MsoCalloutAngleType angle);

        /**
        True if the place where the callout line attaches to the callout text box changes depending on whether the origin of the callout line (where the callout points to) is to the left or right of the callout text box. Read/write MsoTriState.

        [MSDN documentation for CalloutFormat.AutoAttach](http://msdn.microsoft.com/en-us/library/bb220849).
        */
        MsoTriState GetAutoAttach();

        /**
        True if the place where the callout line attaches to the callout text box changes depending on whether the origin of the callout line (where the callout points to) is to the left or right of the callout text box. Read/write MsoTriState.

        [MSDN documentation for CalloutFormat.AutoAttach](http://msdn.microsoft.com/en-us/library/bb220849).
        */
        void SetAutoAttach(MsoTriState autoAttach);

        /**
        Applies only to callouts whose lines consist of more than one segment (types msoCalloutThree and msoCalloutFour). Read/write MsoTriState.

        [MSDN documentation for CalloutFormat.AutoLength](http://msdn.microsoft.com/en-us/library/bb220854).
        */
        MsoTriState GetAutoLength();

        /**
        Applies only to callouts whose lines consist of more than one segment (types msoCalloutThree and msoCalloutFour). Read/write MsoTriState.

        [MSDN documentation for CalloutFormat.AutoLength](http://msdn.microsoft.com/en-us/library/bb220854).
        */
        void SetAutoLength(MsoTriState autoLength);

        /**
        Returns a MsoTriState value that represents the visibility options for the border of the object.

        [MSDN documentation for CalloutFormat.Border](http://msdn.microsoft.com/en-us/library/bb179356).
        */
        MsoTriState GetBorder();

        /**
        Sets a MsoTriState value that represents the visibility options for the border of the object.

        [MSDN documentation for CalloutFormat.Border](http://msdn.microsoft.com/en-us/library/bb179356).
        */
        void SetBorder(MsoTriState border);

        /**
        For callouts with an explicitly set drop value, this property returns the vertical distance (in points) from the edge of the text bounding box to the place where the callout line attaches to the text box. Read-only Single.

        [MSDN documentation for CalloutFormat.Drop](http://msdn.microsoft.com/en-us/library/bb221060).
        */
        double GetDrop();

        /**
        Returns a value that indicates where the callout line attaches to the callout text box. Read-only MsoCalloutDropType.

        [MSDN documentation for CalloutFormat.DropType](http://msdn.microsoft.com/en-us/library/bb221073).
        */
        MsoCalloutDropType GetDropType();

        /**
        Returns the horizontal distance (in points) between the end of the callout line and the text bounding box. Read/write Single.

        [MSDN documentation for CalloutFormat.Gap](http://msdn.microsoft.com/en-us/library/bb208569).
        */
        double GetGap();

        /**
        Sets the horizontal distance (in points) between the end of the callout line and the text bounding box. Read/write Single.

        [MSDN documentation for CalloutFormat.Gap](http://msdn.microsoft.com/en-us/library/bb208569).
        */
        void SetGap(double gap);

        /**
        Returns a Single value that represents the length (in points) of the first segment of the callout line (the segment attached to the text callout box.)

        [MSDN documentation for CalloutFormat.Length](http://msdn.microsoft.com/en-us/library/bb179360).
        */
        double GetLength();

        /**
        Returns a MsoCalloutType value that represents the callout format type.

        [MSDN documentation for CalloutFormat.Type](http://msdn.microsoft.com/en-us/library/bb148797).
        */
        MsoCalloutType  GetType();

        /**
        Sets a MsoCalloutType value that represents the callout format type.

        [MSDN documentation for CalloutFormat.Type](http://msdn.microsoft.com/en-us/library/bb148797).
        */
        void SetType(MsoCalloutType  type);


        /**
        Returns "CalloutFormat".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("CalloutFormat"); }
    };


} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_SHAPES

#endif //_WXAUTOEXCEL_CALLOUTFORMAT_H

/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
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

        [MSDN documentation for CalloutFormat.AutomaticLength](https://learn.microsoft.com/en-us/office/vba/api/excel.calloutformat.automaticlength).
        */
        void AutomaticLength();

        /**
        Sets the vertical distance (in points) from the edge of the text bounding box to the place where the callout line attaches to the text box. This distance is measured from the top of the text box unless the AutoAttach property is set to True and the text box is to the left of the origin of the callout line (the place that the callout points to), in which case the drop distance is measured from the bottom of the text box.

        [MSDN documentation for CalloutFormat.CustomDrop](https://learn.microsoft.com/en-us/office/vba/api/excel.calloutformat.customdrop).
        */
        void CustomDrop(double drop);

        /**
        Specifies that the first segment of the callout line (the segment attached to the text callout box) retain a fixed length whenever the callout is moved. Use the AutomaticLength method to specify that the first segment of the callout line be scaled automatically whenever the callout is moved. Applies only to callouts whose lines consist of more than one segment (types msoCalloutThree and msoCalloutFour).

        [MSDN documentation for CalloutFormat.CustomLength](https://learn.microsoft.com/en-us/office/vba/api/excel.calloutformat.customlength).
        */
        void CustomLength(double length);

        /**
        Specifies whether the callout line attaches to the top, bottom, or center of the callout text box or whether it attaches at a point that’s a specified distance from the top or bottom of the text box.

        [MSDN documentation for CalloutFormat.PresetDrop](https://learn.microsoft.com/en-us/office/vba/api/excel.calloutformat.presetdrop).
        */
        void PresetDrop(MsoCalloutDropType dropType);

        // ***** PROPERTIES *****

        /**
        Allows the user to place a vertical accent bar to separate the callout text from the callout line. Read/write MsoTriState.

        [MSDN documentation for CalloutFormat.Accent](https://learn.microsoft.com/en-us/office/vba/api/excel.calloutformat.accent).
        */
        MsoTriState GetAccent();

        /**
        Allows the user to place a vertical accent bar to separate the callout text from the callout line. Read/write MsoTriState.

        [MSDN documentation for CalloutFormat.Accent](https://learn.microsoft.com/en-us/office/vba/api/excel.calloutformat.accent).
        */
        void SetAccent(MsoTriState accent);

        /**
        Returns the angle of the callout line. If the callout line contains more than one line segment, this property returns or sets the angle of the segment that is farthest from the callout text box. Read/write MsoCalloutAngleType.

        [MSDN documentation for CalloutFormat.Angle](https://learn.microsoft.com/en-us/office/vba/api/excel.calloutformat.angle).
        */
        MsoCalloutAngleType GetAngle();

        /**
        Sets the angle of the callout line. If the callout line contains more than one line segment, this property returns or sets the angle of the segment that is farthest from the callout text box. Read/write MsoCalloutAngleType.

        [MSDN documentation for CalloutFormat.Angle](https://learn.microsoft.com/en-us/office/vba/api/excel.calloutformat.angle).
        */
        void SetAngle(MsoCalloutAngleType angle);

        /**
        True if the place where the callout line attaches to the callout text box changes depending on whether the origin of the callout line (where the callout points to) is to the left or right of the callout text box. Read/write MsoTriState.

        [MSDN documentation for CalloutFormat.AutoAttach](https://learn.microsoft.com/en-us/office/vba/api/excel.calloutformat.autoattach).
        */
        MsoTriState GetAutoAttach();

        /**
        True if the place where the callout line attaches to the callout text box changes depending on whether the origin of the callout line (where the callout points to) is to the left or right of the callout text box. Read/write MsoTriState.

        [MSDN documentation for CalloutFormat.AutoAttach](https://learn.microsoft.com/en-us/office/vba/api/excel.calloutformat.autoattach).
        */
        void SetAutoAttach(MsoTriState autoAttach);

        /**
        Applies only to callouts whose lines consist of more than one segment (types msoCalloutThree and msoCalloutFour). Read/write MsoTriState.

        [MSDN documentation for CalloutFormat.AutoLength](https://learn.microsoft.com/en-us/office/vba/api/excel.calloutformat.autolength).
        */
        MsoTriState GetAutoLength();

        /**
        Returns a MsoTriState value that represents the visibility options for the border of the object.

        [MSDN documentation for CalloutFormat.Border](https://learn.microsoft.com/en-us/office/vba/api/excel.calloutformat.border).
        */
        MsoTriState GetBorder();

        /**
        Sets a MsoTriState value that represents the visibility options for the border of the object.

        [MSDN documentation for CalloutFormat.Border](https://learn.microsoft.com/en-us/office/vba/api/excel.calloutformat.border).
        */
        void SetBorder(MsoTriState border);

        /**
        For callouts with an explicitly set drop value, this property returns the vertical distance (in points) from the edge of the text bounding box to the place where the callout line attaches to the text box. Read-only Single.

        [MSDN documentation for CalloutFormat.Drop](https://learn.microsoft.com/en-us/office/vba/api/excel.calloutformat.drop).
        */
        double GetDrop();

        /**
        Returns a value that indicates where the callout line attaches to the callout text box. Read-only MsoCalloutDropType.

        [MSDN documentation for CalloutFormat.DropType](https://learn.microsoft.com/en-us/office/vba/api/excel.calloutformat.droptype).
        */
        MsoCalloutDropType GetDropType();

        /**
        Returns the horizontal distance (in points) between the end of the callout line and the text bounding box. Read/write Single.

        [MSDN documentation for CalloutFormat.Gap](https://learn.microsoft.com/en-us/office/vba/api/excel.calloutformat.gap).
        */
        double GetGap();

        /**
        Sets the horizontal distance (in points) between the end of the callout line and the text bounding box. Read/write Single.

        [MSDN documentation for CalloutFormat.Gap](https://learn.microsoft.com/en-us/office/vba/api/excel.calloutformat.gap).
        */
        void SetGap(double gap);

        /**
        Returns a Single value that represents the length (in points) of the first segment of the callout line (the segment attached to the text callout box.)

        [MSDN documentation for CalloutFormat.Length](https://learn.microsoft.com/en-us/office/vba/api/excel.calloutformat.length).
        */
        double GetLength();

        /**
        Returns a MsoCalloutType value that represents the callout format type.

        [MSDN documentation for CalloutFormat.Type](https://learn.microsoft.com/en-us/office/vba/api/excel.calloutformat.type).
        */
        MsoCalloutType  GetType();

        /**
        Sets a MsoCalloutType value that represents the callout format type.

        [MSDN documentation for CalloutFormat.Type](https://learn.microsoft.com/en-us/office/vba/api/excel.calloutformat.type).
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

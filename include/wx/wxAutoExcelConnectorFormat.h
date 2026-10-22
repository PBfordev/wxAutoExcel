/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_CONNECTORFORMAT_H
#define _WXAUTOEXCEL_CONNECTORFORMAT_H

#include "wx/wxAutoExcel_defs.h"

#if WXAUTOEXCEL_USE_SHAPES

#include "wx/wxAutoExcel_object.h"
#include "wx/wxAutoExcel_enums.h"

namespace wxAutoExcel {

    /**
    @brief Represents Microsoft Excel ConnectorFormat object.
    */
   class WXDLLIMPEXP_WXAUTOEXCEL wxExcelConnectorFormat : public wxExcelObject
    {
    public:
        // ***** METHODS *****

        /**
        Attaches the beginning of the specified connector to a specified shape. If there’s already a connection between the beginning of the connector and another shape, that connection is broken. If the beginning of the connector isn’t already positioned at the specified connecting site, this method moves the beginning of the connector to the connecting site and adjusts the size and position of the connector. Use the EndConnect method to attach the end of the connector to a shape.

        [MSDN documentation for ConnectorFormat.BeginConnect](https://learn.microsoft.com/en-us/office/vba/api/excel.connectorformat.beginconnect).
        */
        void BeginConnect(wxExcelShape connectedShape, long connectionSite);

        /**
        Detaches the beginning of the specified connector from the shape it’s attached to. This method doesn’t alter the size or position of the connector: the beginning of the connector remains positioned at a connection site but is no longer connected. Use the EndDisconnect method to detach the end of the connector from a shape.

        [MSDN documentation for ConnectorFormat.BeginDisconnect](https://learn.microsoft.com/en-us/office/vba/api/excel.connectorformat.begindisconnect).
        */
        void BeginDisconnect();

        /**
        Attaches the end of the specified connector to a specified shape. If there’s already a connection between the end of the connector and another shape, that connection is broken. If the end of the connector isn’t already positioned at the specified connecting site, this method moves the end of the connector to the connecting site and adjusts the size and position of the connector. Use the BeginConnect method to attach the beginning of the connector to a shape.

        [MSDN documentation for ConnectorFormat.EndConnect](https://learn.microsoft.com/en-us/office/vba/api/excel.connectorformat.endconnect).
        */
        void EndConnect(wxExcelShape connectedShape, long connectionSite);

        /**
        Detaches the end of the specified connector from the shape it’s attached to. This method doesn’t alter the size or position of the connector: the end of the connector remains positioned at a connection site but is no longer connected. Use the BeginDisconnect method to detach the beginning of the connector from a shape.

        [MSDN documentation for ConnectorFormat.EndDisconnect](https://learn.microsoft.com/en-us/office/vba/api/excel.connectorformat.enddisconnect).
        */
        void EndDisconnect();

        // ***** PROPERTIES *****

        /**
        True if the beginning of the specified connector is connected to a shape. Read-only MsoTriState.

        [MSDN documentation for ConnectorFormat.BeginConnected](https://learn.microsoft.com/en-us/office/vba/api/excel.connectorformat.beginconnected).
        */
        MsoTriState GetBeginConnected();

        /**
        Returns a Shape Represents the shape that the beginning of the specified connector is attached to.

        [MSDN documentation for ConnectorFormat.BeginConnectedShape](https://learn.microsoft.com/en-us/office/vba/api/excel.connectorformat.beginconnectedshape).
        */
        wxExcelShape GetBeginConnectedShape();

        /**
        Returns an integer that specifies the connection site that the beginning of a connector is connected to.

        [MSDN documentation for ConnectorFormat.BeginConnectionSite](https://learn.microsoft.com/en-us/office/vba/api/excel.connectorformat.beginconnectionsite).
        */
        long GetBeginConnectionSite();

        /**
        msoTrue if the end of the specified connector is connected to a shape. Read-only MsoTriState.

        [MSDN documentation for ConnectorFormat.EndConnected](https://learn.microsoft.com/en-us/office/vba/api/excel.connectorformat.endconnected).
        */
        MsoTriState GetEndConnected();

        /**
        Returns a Shape Represents the shape that the end of the specified connector is attached to.

        [MSDN documentation for ConnectorFormat.EndConnectedShape](https://learn.microsoft.com/en-us/office/vba/api/excel.connectorformat.endconnectedshape).
        */
        wxExcelShape GetEndConnectedShape();

        /**
        Returns an integer that specifies the connection site that the end of a connector is connected to.

        [MSDN documentation for ConnectorFormat.EndConnectionSite](https://learn.microsoft.com/en-us/office/vba/api/excel.connectorformat.endconnectionsite).
        */
        long GetEndConnectionSite();

        /**
        Returns a MsoConnectorType value that represents the connector format type.

        [MSDN documentation for ConnectorFormat.Type](https://learn.microsoft.com/en-us/office/vba/api/excel.connectorformat.type).
        */
        MsoConnectorType  GetType();

        /**
        Returns "ConnectorFormat".
        */
        virtual wxString GetAutoExcelObjectName_() const { return wxS("ConnectorFormat"); }
    };


} // namespace wxAutoExcel

#endif // #if WXAUTOEXCEL_USE_SHAPES

#endif //_WXAUTOEXCEL_CONNECTORFORMAT_H

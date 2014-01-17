/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// RCS-ID:      $Id: $
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_SETUP_H
#define _WXAUTOEXCEL_SETUP_H

// if you don't want to use any chart features
// set WXAUTOEXCEL_USE_CHARTS to 0 to make
// wxAutoExcel libraries smaller
#define WXAUTOEXCEL_USE_CHARTS 1

// if you don't want to use any Shapes features
// set WXAUTOEXCEL_USE_SHAPES to 0 to make
// wxAutoExcel libraries smaller
#define WXAUTOEXCEL_USE_SHAPES 1

// if automation method / property get call
// returns an IDispatch instead of another type
// and the IDispatch is not properly released
// the resource leak is born, which will most likely
// prevent MS Excel instance from being properly closed
#define WXAUTOEXCEL_RELEASE_UNEXPECTED_IDISPATCH 0

// whether to call wxTraceLog(wxTRACE_AutoExcel, ...)  during automation calls
// all trace calls are displayed only in the DEBUG build
#ifndef WXAUTOEXCEL_SHOW_TRACE

    #ifdef NDEBUG
        #define WXAUTOEXCEL_SHOW_TRACE 0
    #else
        #define WXAUTOEXCEL_SHOW_TRACE 1
    #endif // #ifdef NDEBUG

#endif // #ifndef WXAUTOEXCEL_SHOW_TRACE


// whether to show the original error dialog displayed 
// by wxAutomationObject::Invoke() when the COM call fails
#ifndef WXAUTOEXCEL_SHOW_WXAUTOMATION_ERROR

    #ifdef NDEBUG
        #define WXAUTOEXCEL_SHOW_WXAUTOMATION_ERROR 1
    #else
        #define WXAUTOEXCEL_SHOW_WXAUTOMATION_ERROR 1
    #endif // #ifdef NDEBUG

#endif // #ifndef WXAUTOEXCEL_SHOW_WXAUTOMATION_ERROR


#endif // #ifndef _WXAUTOEXCEL_SETUP_H

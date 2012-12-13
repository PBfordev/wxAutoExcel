/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// RCS-ID:      $Id: $
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_DEFS_H
#define _WXAUTOEXCEL_DEFS_H

#if  !defined(__WXMSW__) || !wxUSE_OLE || !wxUSE_VARIANT
    #error wxAutoExcel requires wxWidgets to be built for MS Windows, with support for OLE and wxVariant
#endif

/** @file
Contains wxAutoExcel global definitions.
*/


#include "wx/wxAutoExcel_setup.h"

#include "wx/wxAutoExcel_version.h"

#ifdef WXMAKINGDLL_WXAUTOEXCEL
    #define WXDLLIMPEXP_WXAUTOEXCEL                  WXEXPORT
    #define WXDLLIMPEXP_DATA_WXAUTOEXCEL(type)       WXEXPORT type
#elif defined(WXUSINGDLL)
    #define WXDLLIMPEXP_WXAUTOEXCEL                  WXIMPORT
    #define WXDLLIMPEXP_DATA_WXAUTOEXCEL(type)       WXIMPORT type
#else // not making nor using DLL
    #define WXDLLIMPEXP_WXAUTOEXCEL
    #define WXDLLIMPEXP_DATA_WXAUTOEXCEL(type)	    type
#endif 

/*
  GCC warns about using __declspec on forward declarations
  while MSVC complains about forward declarations without
  __declspec for the classes later declared with it. To hide this
  difference a separate macro for forward declarations is defined:
 */
#if defined(HAVE_VISIBILITY) || defined(__GNUC__)
  #define WXDLLIMPEXP_FWD_WXAUTOEXCEL
#else
  #define WXDLLIMPEXP_FWD_WXAUTOEXCEL WXDLLIMPEXP_WXAUTOEXCEL
#endif 

#include "wx/wxAutoExcel_fwd.h"
#include "wx/wxAutoExcelTribool.h"



/**
@brief All wxAutoExcel classes and enumerations are declared in wxAutoExcel namespace.
*/

namespace wxAutoExcel {


    template <typename T>
    struct wxAutoExcelValPtr
    {    
        wxAutoExcelValPtr(T value) : m_value(value) {}    
        operator T*() { return &m_value; };
        T m_value;
    };

    template <typename T>
    inline wxAutoExcelValPtr<T> wxAutoExcelValPtrFn(T t)
    {    
        return wxAutoExcelValPtr<T>(t);
    }

    /*!  
    Helper macro for passing pointer to a long value.   
    */
#define WXAELP(value) wxAutoExcelValPtrFn<long>(value)
    /*!  
    Helper macro for passing pointer to an enum value.
    */
#define WXAEEP(value) wxAutoExcelValPtrFn(value)

/*!  
    Mask for wxAutoExcel's wxLogTrace() calls
*/
#define wxTRACE_AutoExcel wxS("AutoExcel")

} // namespace wxAutoExcel


#endif //_WXAUTOEXCEL_DEFS_H

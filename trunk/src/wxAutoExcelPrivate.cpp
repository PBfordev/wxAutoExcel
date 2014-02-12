/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// RCS-ID:      $Id: $
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcelPrivate.h"

#include <wx/msw/ole/oleutils.h>

#include "wx/wxAutoExcelObject.h"
#include "wx/wxAutoExcelTribool.h"

namespace wxAutoExcel {

wxXlTribool wxDefaultXlTribool;

size_t LogVariantMaxItemsInList = 30;

void LogVariant(const wxString& prefix, const wxVariant& v)
{
    const wxString type = v.GetType();

    wxString info;
    wxString name = v.GetName();
    if (type == wxS("arrstring")) {
        wxArrayString as = v.GetArrayString();
        info.Printf(wxS("%svariant type: \"%s\", element count: %d, name: \"%s\"."),
            prefix.c_str(), type.c_str(), as.size(), name.c_str());        
        wxLogTrace(wxTRACE_AutoExcel, wxS("%s"), info.c_str());
        for (size_t i = 0; i < as.size(); i++) 
        {
            info.Printf(wxS("   string #%d value: \"%s\""), i, as[i]);
            if ( i == LogVariantMaxItemsInList )
            {
                wxLogTrace(wxTRACE_AutoExcel, wxS("And %d more strings"), as.size() - i);
                break;
            }
            else            
                wxLogTrace(wxTRACE_AutoExcel, wxS("%s"), info.c_str());
        }
        return;
    }
    if (type == wxS("list")) {
        info.Printf(wxS("%sVariant type: \"%s\", element count: %d, name: \"%s\"."),
            prefix.c_str(), type.c_str(), v.GetCount(), name.c_str());
        wxLogTrace(wxTRACE_AutoExcel, wxS("%s"), info.c_str());
        for (size_t i = 0; i < v.GetCount(); i++)
        {
            if ( i == LogVariantMaxItemsInList )
            {
                wxLogTrace(wxTRACE_AutoExcel, wxS("And %d more variants"), v.GetCount() - i);
                break;
            } else            
            {
                const wxVariant& vTmp = v[i];
                info.Printf(wxS("   variant #%d type: \"%s\", value: \"%s\", name: \"%s\"."),
                    i, vTmp.GetType().c_str(), vTmp.MakeString().c_str(), vTmp.GetName().c_str());        
                wxLogTrace(wxTRACE_AutoExcel, wxS("%s"), info.c_str());
            }
        }
        return;
    }
    if (type == wxS("void*") && v.GetVoidPtr() != NULL) {
        wxString automationName;
        wxExcelObject object;
        IDispatch* dispatch = (IDispatch*)v.GetVoidPtr();
        dispatch->AddRef();
        object.GetAutomationObject_()->SetDispatchPtr(dispatch);
        info.Printf(wxS("%svariant type: \"IDispatch - %s\", value: \"%s\", name: \"%s\"."),
            prefix.c_str(), object.GetAutomationObjectName_(false).c_str(), v.MakeString().c_str(), name.c_str());    
    } else {
        info.Printf(wxS("%svariant type: \"%s\", value: \"%s\", name: \"%s\"."),
            prefix.c_str(), type.c_str(), v.MakeString().c_str(), name.c_str());        
    }
    wxLogTrace(wxTRACE_AutoExcel, wxS("%s"), info.c_str());
}


} // namespace wxAutoExcel

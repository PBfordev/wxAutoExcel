/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include "wx/wxAutoExcel_private.h"

#include <wx/msw/ole/oleutils.h>

#include "wx/wxAutoExcel_object.h"
#include "wx/wxAutoExcel_tribool.h"

namespace wxAutoExcel {

wxXlTribool wxDefaultXlTribool;

size_t LogVariantMaxItemsInList = 30;

void LogVariant(const wxString& prefix, const wxVariant& v)
{
    const wxString type = v.GetType();

    wxString info;
    const wxString& name = v.GetName();
    if (type == wxS("arrstring")) {
        wxArrayString as = v.GetArrayString();
        info.Printf(wxS("%svariant type: \"%s\", element count: %zu, name: \"%s\"."),
            prefix, type, as.size(), name);
        wxLogTrace(wxTRACE_AutoExcel, wxS("%s"), info);
        for (size_t i = 0; i < as.size(); i++)
        {
            info.Printf(wxS("   string #%zu value: \"%s\""), i, as[i]);
            if ( i == LogVariantMaxItemsInList )
            {
                wxLogTrace(wxTRACE_AutoExcel, wxS("And %zu more strings"), as.size() - i);
                break;
            }
            else
                wxLogTrace(wxTRACE_AutoExcel, wxS("%s"), info);
        }
        return;
    }
    if (type == wxS("list")) {
        info.Printf(wxS("%sVariant type: \"%s\", element count: %zu, name: \"%s\"."),
            prefix, type, v.GetCount(), name);
        wxLogTrace(wxTRACE_AutoExcel, wxS("%s"), info);
        for (size_t i = 0; i < v.GetCount(); i++)
        {
            if ( i == LogVariantMaxItemsInList )
            {
                wxLogTrace(wxTRACE_AutoExcel, wxS("And %zu more variants"), v.GetCount() - i);
                break;
            } else
            {
                const wxVariant& vTmp = v[i];
                info.Printf(wxS("   variant #%zu type: \"%s\", value: \"%s\", name: \"%s\"."),
                    i, vTmp.GetType(), vTmp.MakeString(), vTmp.GetName());
                wxLogTrace(wxTRACE_AutoExcel, wxS("%s"), info);
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
            prefix, object.GetAutomationObjectName_(false), v.MakeString(), name);
    } else {
        info.Printf(wxS("%svariant type: \"%s\", value: \"%s\", name: \"%s\"."),
            prefix, type, v.MakeString(), name);
    }
    wxLogTrace(wxTRACE_AutoExcel, wxS("%s"), info);
}


} // namespace wxAutoExcel

/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#include "wx/wxAutoExcel_prec.h"

#include <vector>
#include <exception>

#include <wx/msw/private/comptr.h>
#include <wx/log.h>

#include "wx/wxAutoExcelObject.h"
#include "wx/wxAutoExcel_private.h"


namespace wxAutoExcel {

unsigned wxExcelObject::ms_errorMode =
    Err_LogOnInvalidDispatch   | Err_AssertOnInvalidDispatch
    | Err_LogOnInvalidArgument | Err_AssertOnInvalidArgument
    | Err_LogOnFailedInvoke    | Err_AssertOnFailedInvoke
    | Err_LogOnInvalidReturnType | Err_AssertOnInvalidReturnType
    | Err_LogOnOtherError      | Err_AssertOnOtherError;

wxExcelObject::wxExcelObject()
    : m_xlObject(new wxAutomationObject)
{
    m_lastCallSucceeded = true;
}

wxExcelObject::~wxExcelObject()
{

}

//wxExcelObject wxExcelObject::GetParent()
//{
//    wxExcelObject object;
//    WXAUTOEXCEL_PROPERTY_OBJECT_GET0("Parent", object);
//}

unsigned wxExcelObject::GetErrorMode_()
{
    return ms_errorMode;
}

void wxExcelObject::SetErrorMode_(unsigned mode)
{
    ms_errorMode = mode;
}

LCID wxExcelObject::GetAutomationLCID_() const
{
    wxCHECK( m_xlObject && m_xlObject->GetDispatchPtr(), LOCALE_SYSTEM_DEFAULT );

    return m_xlObject->GetLCID();
}

    
bool wxExcelObject::SetAutomationLCID_(LCID lcid)
{
    wxCHECK( m_xlObject && m_xlObject->GetDispatchPtr(), false);

    m_xlObject->SetLCID(lcid);
    return true;
}

namespace {

class wxPBXLInvokeArgs
{
public:    
    wxPBXLInvokeArgs(const wxVariant* arg1, const wxVariant* arg2, const wxVariant* arg3,
            const wxVariant* arg4, const wxVariant* arg5, const wxVariant* arg6);
    wxPBXLInvokeArgs(const wxVariantVector& args);
    ~wxPBXLInvokeArgs();

    size_t GetNoArgs() const { return m_noArgs; }
    wxVariant const** GetArgs() { return m_argsCustom ? m_argsCustom : m_args6; }
private:
    size_t m_noArgs;
    wxVariant const* m_args6[6];
    wxVariant const** m_argsCustom;
};

} // unnamed namespace

wxPBXLInvokeArgs::wxPBXLInvokeArgs(const wxVariant* const arg1, const wxVariant* arg2, const wxVariant* arg3,
                    const wxVariant* arg4, const wxVariant* arg5, const wxVariant* arg6)
{
    m_noArgs = 0;
    m_argsCustom = NULL;
    
    if ( !arg1->IsNull() )
        m_args6[m_noArgs++] = arg1;
    if ( !arg2->IsNull() )
        m_args6[m_noArgs++] = arg2;
    if ( !arg3->IsNull() )
        m_args6[m_noArgs++] = arg3;
    if ( !arg4->IsNull() )
        m_args6[m_noArgs++] = arg4;
    if ( !arg5->IsNull() )
        m_args6[m_noArgs++] = arg5;
    if ( !arg6->IsNull() )
        m_args6[m_noArgs++] = arg6;    
}

wxPBXLInvokeArgs::wxPBXLInvokeArgs(const wxVariantVector& args)
{
    m_noArgs = args.size();    
    wxVariant const** argsPtr = NULL;

    if ( args.size() > 6 )
    {
        m_argsCustom = new wxVariant const*[m_noArgs];
        argsPtr = m_argsCustom;
    }
    else 
    {
        m_argsCustom = NULL;
        argsPtr = m_args6;                
    }
    for ( size_t i = 0; i < m_noArgs; i++ )
    {
        argsPtr[i] = &args[i];
    }
}

wxPBXLInvokeArgs::~wxPBXLInvokeArgs()
{
    if ( m_argsCustom )
        delete[] m_argsCustom;
}


bool wxExcelObject::InvokeMethod(const wxString& member, wxVariant& retValue,
                      const wxVariant& arg1, const wxVariant& arg2,
                      const wxVariant& arg3, const wxVariant& arg4,
                      const wxVariant& arg5, const wxVariant& arg6)
{    
    wxPBXLInvokeArgs arguments(&arg1, &arg2, &arg3, &arg4, &arg5, &arg6);
    return Invoke(member, DISPATCH_METHOD, retValue, arguments.GetNoArgs(), arguments.GetArgs());    
}

bool wxExcelObject::InvokeMethodArray(const wxString& member, wxVariant& retValue, const wxVariantVector& args)
{        
    wxPBXLInvokeArgs arguments(args);
    return Invoke(member, DISPATCH_METHOD, retValue, arguments.GetNoArgs(), arguments.GetArgs());        
}

bool wxExcelObject::InvokeGetProperty(const wxString& member, wxVariant& retValue,
                      const wxVariant& arg1, const wxVariant& arg2,
                      const wxVariant& arg3, const wxVariant& arg4,
                      const wxVariant& arg5, const wxVariant& arg6)
{    
    wxPBXLInvokeArgs arguments(&arg1, &arg2, &arg3, &arg4, &arg5, &arg6);    
    return Invoke(member, DISPATCH_PROPERTYGET, retValue, arguments.GetNoArgs(), arguments.GetArgs());    
}

bool wxExcelObject::InvokeGetPropertyArray(const wxString& member, wxVariant& retValue, const wxVariantVector& args)
{ 
    wxPBXLInvokeArgs arguments(args);
    return Invoke(member, DISPATCH_PROPERTYGET, retValue, arguments.GetNoArgs(), arguments.GetArgs());
}

bool wxExcelObject::InvokePutProperty(const wxString& member,
                                      const wxVariant& arg1, const wxVariant& arg2,
                                      const wxVariant& arg3, const wxVariant& arg4,
                                      const wxVariant& arg5, const wxVariant& arg6)
{
    wxVariant retValue;    
    wxPBXLInvokeArgs arguments(&arg1, &arg2, &arg3, &arg4, &arg5, &arg6);    
                
    return Invoke(member, DISPATCH_PROPERTYPUT, retValue, arguments.GetNoArgs(), arguments.GetArgs());
}

bool wxExcelObject::InvokePutPropertyArray(const wxString& member, const wxVariantVector& args)
{
    wxVariant retValue;
    wxPBXLInvokeArgs arguments(args);    
    return Invoke(member, DISPATCH_PROPERTYPUT, retValue, arguments.GetNoArgs(), arguments.GetArgs());
}

namespace {

void ReleaseVariantDispatch(wxVariant& variant)
{
    if ( variant.IsType(wxS("void*")) )
    {
        IDispatch* dispatch = (IDispatch*)variant.GetVoidPtr();
        if ( dispatch )
        {
            dispatch->Release();
            variant = (void*)NULL;
        }
    }
    else if ( variant.IsType(("list")) )
    {
        for ( size_t i = 0; i < variant.GetCount(); i++ )
        {
            ReleaseVariantDispatch(variant[i]);
        }
    }
}

void ClearArguments(wxVariantVector& args)
{
    for ( size_t i = 0; i < args.size(); i++)
    {
        ReleaseVariantDispatch(args[i]);
    }
}


} // unnamed namespace

bool wxExcelObject::Invoke(const wxString& member, int action, wxVariant& retValue,
                           size_t noArgs, const wxVariant* args[])
{
    m_lastCallSucceeded = IsOk_();

    if ( !m_lastCallSucceeded )
    {
        OnError(Err_InvalidDispatch, member);
        // @FIXME ClearArguments(args);
        return false;
    }

    if ( member.Find(wxT('.')) != wxNOT_FOUND )
    {
        m_lastCallSucceeded = false;
        // @FIXME ClearArguments(args);
        OnError(Err_InvalidArgument, member);
    }

#if WXAUTOEXCEL_SHOW_TRACE
    // can get expensive so don't do it if tracing is off
    if  ( wxLog::IsAllowedTraceMask(wxTRACE_AutoExcel) ) 
    {
        wxString traceMsg(wxS("*** wxExcelObject::Invoke on "));
        traceMsg << GetAutoExcelObjectName_() << wxS(" (") << GetAutomationObjectName_(false) << wxS(")." ) << member;
        if ( action == DISPATCH_PROPERTYPUT )
        {
            traceMsg << wxS(" (property set)");
        }
        else if ( action == DISPATCH_PROPERTYGET )
        {
            traceMsg << wxS(" (property get)");
        }
        else if ( action == DISPATCH_METHOD)
        {
            traceMsg << wxS(" (method call)");
        }
        traceMsg << wxS(" ***");
        wxLogTrace(wxTRACE_AutoExcel, traceMsg);        
        for (size_t i = 0; i < noArgs; i++)
            LogVariant(wxString::Format("  argument #%zu: ", i), *args[i]);
    }
#endif // #if WXAUTOEXCEL_SHOW_TRACE

    {
#if !WXAUTOEXCEL_SHOW_WXAUTOMATION_ERROR        
        
        wxLogNull logNo;

#endif // #if WXAUTOEXCEL_SHOW_WXAUTOMATION_ERROR        

        m_lastCallSucceeded = m_xlObject->Invoke(member, action, retValue,
            noArgs, NULL, (args));
    }

#if WXAUTOEXCEL_SHOW_TRACE
    LogVariant(wxS("  wxAutomationObject::Invoke() returned: "), retValue);
#endif // #if WXAUTOEXCEL_SHOW_TRACE

    if ( !m_lastCallSucceeded )
    {
        OnError(Err_FailedInvoke, member);
        return false;
    }
    return true;
}

bool wxExcelObject::CheckReturnType(wxVariant& variant, const wxString& type, const wxString& name, const wxString& function)
{

    if ( !variant.IsType(type) ) 
    {
        m_lastCallSucceeded = false;
#if WXAUTOEXCEL_SHOW_TRACE
        wxLogTrace(wxTRACE_AutoExcel, wxS("Error: variant type %s expected for %s; received %s instead (in %s):"),
            type, name, variant.GetType(), function);
#else
        wxUnusedVar(function); // avoid unused variable warning
#endif // #if WXAUTOEXCEL_SHOW_TRACE

#if WXAUTOEXCEL_RELEASE_UNEXPECTED_IDISPATCH
        if ( variant.IsType(wxS("void*")) )
        {    
            IDispatch* dispatch = (IDispatch*)GetVoidPtr();
            if ( dispatch )
            {    
                wxLogTrace(wxTRACE_AutoExcel, wxS("Unexpected IDISPATCH* released."));    
                dispatch->Release();
            }
        }
#endif // WXAUTOEXCEL_RELEASE_UNEXPECTED_DISPATCH
        OnError(Err_InvalidReturnType, name);        
        return false;
    }
    return true;
}

/*
bool wxExcelObject::SetDispatch(wxExcelObject* obj, IDispatch* dispatch)
{
    if ( !obj || !obj->m_xlObject || obj->m_xlObject->GetDispatchPtr() != NULL
         || !dispatch )
    {
        m_lastCallSucceeded = false;
        OnError(Err_InvalidArgument);
        return false;
    }
    dispatch->AddRef();
    obj->m_xlObject->SetDispatchPtr(dispatch);
    return true;
}
*/

bool wxExcelObject::CloneDispatch(const wxExcelObject* from, wxExcelObject* to)
{
    if ( !from || !from->m_xlObject || !from->m_xlObject->GetDispatchPtr()
         || !to || !to->m_xlObject || to->m_xlObject->GetDispatchPtr() != NULL )
    {
        m_lastCallSucceeded = false;
        OnError(Err_InvalidArgument);
        return false;
    }

    IDispatch* dispatch = (IDispatch*)from->m_xlObject->GetDispatchPtr();
    dispatch->AddRef();
    to->m_xlObject->SetDispatchPtr(dispatch);
    to->m_xlObject->SetLCID(from->m_xlObject->GetLCID());
    return true;
}

bool wxExcelObject::ObjectToVariant(const wxExcelObject* obj, wxVariant& result, const wxString& name)
{
    if ( !obj || !obj->m_xlObject || !obj->m_xlObject->GetDispatchPtr()
         || !result.IsNull() )
    {
        m_lastCallSucceeded = false;
        OnError(Err_InvalidArgument);
        return false;
    }

    IDispatch* dispatch = (IDispatch*)obj->m_xlObject->GetDispatchPtr();
    dispatch->AddRef(); // we need to this because wxAutomationObject::Invoke() decreases the ref count of its parameters
    result = (void*)dispatch;
    if (!name.empty())
    result.SetName(name);
    return true;
}

bool wxExcelObject::VariantToObject(const wxVariant& variant, wxExcelObject* obj)
{
    if ( !obj || !obj->m_xlObject || obj->m_xlObject->GetDispatchPtr()
         || !variant.IsType(wxS("void*")) ||  variant.GetVoidPtr() == NULL)
    {
        m_lastCallSucceeded = false;
        OnError(Err_InvalidArgument);
        return false;
    }
    obj->m_xlObject->SetDispatchPtr((IDispatch*)variant.GetVoidPtr());
    // let it inherit lcid from the calling object
    obj->m_xlObject->SetLCID(m_xlObject->GetLCID());
    return true;
}

wxString wxExcelObject::GetAutomationObjectName_(bool stripUnderscores) const
{
    wxString name(wxS("Unknown automation type"));
    
    if ( m_xlObject )
    {
        IDispatch* dispatch = (IDispatch*)m_xlObject->GetDispatchPtr();
        if ( dispatch )
        {
            wxCOMPtr<ITypeInfo> typeInfo;            
            if ( SUCCEEDED(dispatch->GetTypeInfo(0, 1033, &typeInfo)) )
            {
                BSTR bName;
                if ( SUCCEEDED(typeInfo->GetDocumentation(MEMBERID_NIL, &bName, NULL, NULL, NULL)) )
                {
                    name = bName;
                    SysFreeString(bName);

                    if ( stripUnderscores )
                        name.Replace(wxS("_"), wxEmptyString);
                }
            }
        }
    }
    return name;
}

void wxExcelObject::ReleaseVariantDispatch(wxVariant& variant)
{
    if ( variant.IsType(wxS("void*")) )
    {
        IDispatch* dispatch = (IDispatch*)variant.GetVoidPtr();
        if (dispatch)
            dispatch->Release();
    }
}



void wxExcelObject::OnError(Errors error, const wxString& member)
{
    wxString errMsg;
    unsigned errMode = GetErrorMode_();
    
    if ( error == Err_InvalidDispatch )
    {        
        errMsg << wxS("Attempted to Invoke with an invalid IDispatch (");
        errMsg << GetAutoExcelObjectName_() << wxS(".") << member << wxS(").");

        if ( errMode & Err_AssertOnInvalidDispatch )
        {
            wxASSERT_MSG(IsOk_(), errMsg);        
        }
        if ( errMode & Err_LogOnInvalidDispatch )
            wxLogError(errMsg);        
        if ( errMode & Err_ThrowOnInvalidDispatch )
            throw wxExcelException(errMsg.ToStdString());
    }
    else if ( error == Err_InvalidArgument )
    {        
        errMsg << wxS("Attempted to access a property or call a method with an invalid argument (") << GetAutoExcelObjectName_();
        if ( !member.empty() )
            errMsg << wxS(".") << member;
        errMsg <<  wxS(").");

        if ( errMode & Err_AssertOnInvalidArgument )        
        {
            wxASSERT_MSG(IsOk_(), errMsg);        
        }
        if ( errMode & Err_LogOnInvalidArgument )
            wxLogError(errMsg);        
        if ( errMode & Err_ThrowOnInvalidArgument )
            throw wxExcelException(errMsg.ToStdString());
    }
    else if ( error == Err_FailedInvoke )
    {
        errMsg << wxS("wxAutomationObject::Invoke() failed (") << GetAutoExcelObjectName_() << wxS(".") << member << wxS(").");

        if ( errMode & Err_AssertOnFailedInvoke )
        {
            wxASSERT_MSG(IsOk_(), errMsg);        
        }
        if ( errMode & Err_LogOnFailedInvoke )
            wxLogError(errMsg);        
        if ( errMode & Err_ThrowOnFailedInvoke )
            throw wxExcelException(errMsg.ToStdString());
    }
    else if ( error == Err_InvalidReturnType )
    {        
        errMsg << wxS("wxAutomationObject::Invoke() returned invalid variant type (") << GetAutoExcelObjectName_() << wxS(".") << member << wxS(").");
        
        if ( errMode & Err_AssertOnInvalidReturnType )
        {
            wxASSERT_MSG(IsOk_(), errMsg);                
        }
        if ( errMode & Err_LogOnInvalidReturnType )
            wxLogError(errMsg);
        if ( errMode & Err_ThrowOnInvalidReturnType )
            throw wxExcelException(errMsg.ToStdString());
    } else
    {
        errMsg << wxS("Unspecified error (") << GetAutoExcelObjectName_() << wxS(").");

        if ( errMode & Err_AssertOnOtherError)
        {
            wxASSERT_MSG(IsOk_(), errMsg);                        
        }
        if ( errMode & Err_LogOnOtherError)
            wxLogError(errMsg);        
        if ( errMode & Err_ThrowOnOtherError)
            throw wxExcelException(errMsg.ToStdString());
    }

}


wxAutoExcelObjectErrorModeOverrider::wxAutoExcelObjectErrorModeOverrider(unsigned newMode, bool supressLogging) 
    : m_savedMode(wxExcelObject::GetErrorMode_()) 
{
    wxExcelObject::SetErrorMode_(newMode); 
    if ( supressLogging )
        m_logNull = new wxLogNull;
}

wxAutoExcelObjectErrorModeOverrider::~wxAutoExcelObjectErrorModeOverrider() 
{ 
    wxExcelObject::SetErrorMode_(m_savedMode); 
} 




} // namespace wxAutoExcel

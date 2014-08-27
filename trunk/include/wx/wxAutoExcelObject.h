/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// RCS-ID:      $Id: $
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_OBJECT_H
#define _WXAUTOEXCEL_OBJECT_H

#if  !defined(__WXMSW__) || !wxUSE_OLE || !wxUSE_VARIANT
    #error wxAutoExcel requires wxWidgets to be built for MS Windows, with support for OLE and wxVariant
#endif

#include <stdexcept>

#include <wx/vector.h>
#include <wx/sharedptr.h>
#include <wx/msw/ole/automtn.h>

#include "wx/wxAutoExcel_defs.h"

namespace wxAutoExcel {


/**
    @brief The exception that is thrown if there's an error and wxExcelObject GetErrorMode()_
    has at least one of the Err_ThrowOn* flags set.
*/

class wxExcelException : public std::runtime_error
{
public:
    explicit wxExcelException (const std::string& what) 
        : std::runtime_error(what) 
    {}
};

/**
    @brief wxVector of wxVariants.
*/
typedef wxVector<wxVariant> wxVariantVector;

/**
    @brief The base object for all wxAutoExcel objects.
*/
class WXDLLIMPEXP_WXAUTOEXCEL wxExcelObject
{
public:
    /**
        Flags affecting the behaviour of wxExcelObject and its descendants when an error occurs during accessing an Excel object property or calling its method.
    */
    enum ErrorFlags {
        Err_LogOnInvalidDispatch        = 1, /*!< Call wxLogError() when attempted to access a property or call a method when the object doesn't have a valid dispatch. */
        Err_AssertOnInvalidDispatch     = 1 << 1, /*!< Call wxASSERT when attempted to access a property or call a method when the object doesn't have a valid dispatch. */
        Err_ThrowOnInvalidDispatch      = 1 << 2, /*!< Throw wxExcelException when attempted to access a property or call a method when the object doesn't have a valid dispatch. */

        Err_LogOnInvalidArgument        = 1 << 3, /*!< Call wxLogError() when an invalid argument has been passed to [Get|Set]Property or calling a method. */
        Err_AssertOnInvalidArgument     = 1 << 4, /*!< Call wxASSERT when an invalid argument has been passed to [Get|Set]Property or calling a method. */
        Err_ThrowOnInvalidArgument      = 1 << 5, /*!< Throw wxExcelException when an invalid argument has been passed to [Get|Set]Property or calling a method. */

        Err_LogOnFailedInvoke           = 1 << 6, /*!< Call wxLogError() when the underlying wxAutomationObject::Invoke() returned false. */
        Err_AssertOnFailedInvoke        = 1 << 7, /*!< Call wxASSERT when the underlying wxAutomationObject::Invoke() returned false. */
        Err_ThrowOnFailedInvoke         = 1 << 8, /*!< Throw wxExcelException when the underlying wxAutomationObject::Invoke() returned false. */

        Err_LogOnInvalidReturnType      = 1 << 9, /*!< Call wxLogError() when the underlying wxAutomationObject::Invoke() returned unexpected variant type. */
        Err_AssertOnInvalidReturnType     = 1 << 10, /*!< Call wxASSERT when the underlying wxAutomationObject::Invoke() returned unexpected variant type. */
        Err_ThrowOnInvalidReturnType    = 1 << 11, /*!< Throw wxExcelException when the underlying wxAutomationObject::Invoke() returned unexpected variant type. */

        Err_LogOnOtherError             = 1 << 12, /*!< Call wxLogError() when an unspecified error occured */
        Err_AssertOnOtherError          = 1 << 13, /*!< Call wxASSERT when an unspecified error occured */
        Err_ThrowOnOtherError           = 1 << 14, /*!< Throw wxExcelException when an unspecified error occured */
    };

    wxExcelObject();
    virtual ~wxExcelObject();

    /**
    Returns the parent object.
    
    */
    // wxExcelObject GetParent();

    /**
        Returns true if the object has a valid dispatch.
    */
    bool IsOk_() const { return m_xlObject && m_xlObject->IsOk(); }

    /**
        Sets the error mode to the combination of wxExcelObject::ErrorFlags.

        Default error mode is set to
            Err_LogOnInvalidDispatch   | Err_AssertOnInvalidDispatch
            | Err_LogOnInvalidArgument | Err_AssertOnInvalidArgument
            | Err_LogOnFailedInvoke    | Err_AssertOnFailedInvoke
            | Err_LogOnInvalidReturnType | Err_AssertOnInvalidReturnType
            | Err_LogOnOtherError      | Err_AssertOnOtherError;
    */
    static unsigned GetErrorMode_();

    /**
        Returns the current error mode as the combination of wxExcelObject::ErrorFlags.
    */
    static void SetErrorMode_(unsigned mode);   

    /**
        Returns "Object".
    */
    virtual wxString GetAutoExcelObjectName_() const { return wxS("Object"); }

    /**
        Returns object name as provided by IDispatch->GetTypeInfo()->GetDocumentation().
    */
    wxString GetAutomationObjectName_(bool stripUnderscores = false) const;


    /**
        Returns the underlying wxAutomationObject.
    */
    wxSharedPtr<wxAutomationObject> GetAutomationObject_() { return m_xlObject; }

    /**
        Returns true if the object has a valid dispatch and the last automation call (if any) succeeded.
    */
    operator bool() const { return IsOk_() && m_lastCallSucceeded; }

    /**
        Returns the locale identifier used in automation calls. See wxAutomationObject::GetLCID().
    */
    LCID GetAutomationLCID_() const;

    /**
        Sets the locale identifier to be used in automation calls. See wxAutomationObject::SetLCID().
    */
    bool SetAutomationLCID_(LCID lcid);

protected:
    /**
        @cond PRIVATE
    */

    enum Errors {
        Err_InvalidDispatch,
        Err_InvalidArgument,
        Err_FailedInvoke,
        Err_InvalidReturnType,
        Err_OtherError
    };

    static unsigned ms_errorMode;

    wxSharedPtr<wxAutomationObject> m_xlObject;

    bool m_lastCallSucceeded;

    bool InvokeMethod(const wxString& member, wxVariant& retValue,
                      const wxVariant& arg1 = wxNullVariant, const wxVariant& arg2 = wxNullVariant,
                      const wxVariant& arg3 = wxNullVariant, const wxVariant& arg4 = wxNullVariant,
                      const wxVariant& arg5 = wxNullVariant, const wxVariant& arg6 = wxNullVariant);
    bool InvokeMethodArray(const wxString& member, wxVariant& retValue, const wxVariantVector& args);

    bool InvokeGetProperty(const wxString& member, wxVariant& retValue,
                           const wxVariant& arg1 = wxNullVariant, const wxVariant& arg2 = wxNullVariant,
                           const wxVariant& arg3 = wxNullVariant, const wxVariant& arg4 = wxNullVariant,
                           const wxVariant& arg5 = wxNullVariant, const wxVariant& arg6 = wxNullVariant);
    bool InvokeGetPropertyArray(const wxString& member, wxVariant& retValue, const wxVariantVector& args);

    bool InvokePutProperty(const wxString& member,
                           const wxVariant& arg1 = wxNullVariant, const wxVariant& arg2 = wxNullVariant,
                           const wxVariant& arg3 = wxNullVariant, const wxVariant& arg4 = wxNullVariant,
                           const wxVariant& arg5 = wxNullVariant, const wxVariant& arg6 = wxNullVariant);
    bool InvokePutPropertyArray(const wxString& member, const wxVariantVector& args);

    void OnError(Errors error, const wxString& member = wxEmptyString);

    bool CheckReturnType(wxVariant& variant, const wxString& type, const wxString& name, const wxString& function);

    bool SetDispatch(wxExcelObject* obj, IDispatch* dispatch);
    bool CloneDispatch(const wxExcelObject* from, wxExcelObject* to);
    bool ObjectToVariant(const wxExcelObject* obj, wxVariant& result, const wxString& name = wxEmptyString);
    bool VariantToObject(const wxVariant& variant, wxExcelObject* obj);
    void ReleaseVariantDispatch(wxVariant& variant);

    /**
        @endcond
    */
private:    
    bool Invoke(const wxString& member, int action, wxVariant& retValue, size_t noArgs, const wxVariant* args[]);
};


} // namespace wxAutoExcel

#endif //_WXAUTOEXCEL_CORE_H

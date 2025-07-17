/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// License:     MIT license
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
    @brief The exception that is thrown if there's an error and wxExcelObject::GetErrorMode()_
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
    static const WXLCID lcidEnglishUS;

    /**
        Flags affecting the behaviour of wxExcelObject and its descendants when an error occurs during accessing an Excel object property or calling its method.
    */
    enum ErrorFlags {
        Err_DoNothing                   = 0, /*!< Do nothing when an error occurs, not to be combined with other @c Err_* flags. */
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
        Sets the error mode as the combination of wxExcelObject::ErrorFlags.
    */
    static void SetErrorMode_(unsigned mode);

    /**
        Returns "Object".
    */
    virtual wxString GetAutoExcelObjectName_() const { return wxS("Object"); }

    /**
        Returns object name as provided by IDispatch->GetTypeInfo()->GetDocumentation(MEMBERID_NIL, &name, NULL, NULL, NULL).
    */
    wxString GetAutomationObjectName_(bool stripUnderscores = false) const;


    /**
        Returns the underlying wxAutomationObject.
    */
    wxSharedPtr<wxAutomationObject> GetAutomationObject_() { return m_xlObject; }

    /**
        Returns true if the object has a valid dispatch and its last automation call (if any),
        i.e.  accessing a property or calling a method, succeeded.
    */
    operator bool() const { return IsOk_() && m_lastCallSucceeded; }

    /**
        Returns the locale identifier used in automation calls. See wxAutomationObject::GetLCID().
    */
    WXLCID  GetAutomationLCID_() const;

    /**
        Sets the locale identifier to be used in automation calls. See wxAutomationObject::SetLCID().
        Be aware that this among else affects how MS Excel interprets list separators and values in
        e.g. in Range.Address, Range.Value.
    */
    bool SetAutomationLCID_(WXLCID  lcid);

    /**
    Returns lists of property and method names the automation interface exposes.
    If @a includeHidden is false, names of properties and methods with FUNCFLAG_FHIDDEN
    set will not be included.
    Note. The list will also include methods of IUnknown and IDispatch.
    */
    bool GetPropertyAndMethodNames_(wxArrayString& propertyNames, wxArrayString& methodNames, bool includeHidden = false);

    /**
        Helper function for receiving an automation object not implemented in wxAutoExcel, see FAQ on how to use.
    */
    bool GetUnimplementedObject_(const wxString& name, wxAutomationObject& object);

    //@{
    /**
        Helper function for obtaining an item from MS Excel collection not implemented in wxAutoExcel, see FAQ on how to use.
        Some collections provide their items as as Item property (e.g. Workbooks, Windows, Worksheets or Ranges)
        while others as a result of Item() method call (e.g. ColorStops, Shapes, FormatConditions or Names)
        If the collection provides items via Property, asProperty must be true, otherwise it must be set to false.
        You can learn which to use with online Excel VBA Object Model documentaion or using the Object Browser in the Excel VBA IDE.
        Index must be between 1 and collection.Count.
    */
    static bool GetUnimplementedCollectionItem_(wxAutomationObject& collection, const long index, wxAutomationObject& item, bool asProperty);
    static bool GetUnimplementedCollectionItem_(wxAutomationObject& collection, const wxString& name, wxAutomationObject& item, bool asProperty);
    //@}

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

    bool CloneDispatch(const wxExcelObject* from, wxExcelObject* to);
    bool ObjectToVariant(const wxExcelObject* obj, wxVariant& result, const wxString& name = wxEmptyString);
    bool VariantToObject(const wxVariant& variant, wxExcelObject* obj);
    void ReleaseVariantDispatch(wxVariant& variant);

    /**
        @endcond
    */
private:
    bool Invoke(const wxString& member, int action, wxVariant& retValue, size_t noArgs, const wxVariant* args[]);

    static bool DoGetUnimplementedCollectionItem_(wxAutomationObject& collection, const wxVariant& nameOrIndex, wxAutomationObject& item, bool asProperty);
};

/**
    @brief Sets the new error reporting mode for wxAutoExcel,
    restores the previous error mode when going out of scope.


*/
class WXDLLIMPEXP_WXAUTOEXCEL wxAutoExcelObjectErrorModeOverrider
{
public:
     /**
        Sets the new error reporting mode for the lifetime of the object.
        If wxAutoExcel was compiled with WXAUTOEXCEL_SHOW_WXAUTOMATION_ERROR set to 1
        (which is set by default in non-release builds), wxWidgets will still display
        errors produced in wxAutomationObject::Invoke(), unless you set supressLogging to true.

        WARNING: When supressLogging is true, the error messages are supressed by creating
        a wxLogNull instance, meaning that NO wxLog* messages are going to be displayed.
        Therefore it is highly recommended to make its scope only around wxAutoExcel function call.

        @see wxExcelObject::SetErrorMode_, wxExcelObject::ErrorFlags
    */
    wxAutoExcelObjectErrorModeOverrider(unsigned newMode, bool supressLogging = false);
    ~wxAutoExcelObjectErrorModeOverrider();
private:
    unsigned m_savedMode;
    wxSharedPtr<wxLogNull> m_logNull;

    wxDECLARE_NO_COPY_CLASS(wxAutoExcelObjectErrorModeOverrider);
};

} // namespace wxAutoExcel

#endif //_WXAUTOEXCEL_CORE_H

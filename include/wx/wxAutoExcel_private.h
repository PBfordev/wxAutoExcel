/////////////////////////////////////////////////////////////////////////////
// Author:      PB
// Modified by:
// Copyright:   (c) 2012 PB <pbfordev@gmail.com>
// Licence:     wxWindows licence
/////////////////////////////////////////////////////////////////////////////


#ifndef _WXAUTOEXCEL_PRIVATE_H
#define _WXAUTOEXCEL_PRIVATE_H

#include "wx/wxAutoExcel_defs.h"

struct IDispatch;

namespace wxAutoExcel {


extern size_t LogVariantMaxItemsInList;

void LogVariant(const wxString& prefix, const wxVariant& v);


#define WXAUTOEXCEL_JOIN(a, b) a##b

#define WXAUTOEXCEL_CHECK_VARIANT_TYPE_RET(variant, type, name) \
    if (!CheckReturnType(variant, type, name, __WXFUNCTION__)) {  \
        return;     \
    }

#define WXAUTOEXCEL_CHECK_VARIANT_TYPE(variant, type, name, retVal) \
    if (!CheckReturnType(variant, type, name, __WXFUNCTION__)) {  \
        return retVal;  \
    }

#define WXAUTOEXCEL_CALL_METHOD0_RET(name, resultType)  \
    wxVariant vResult;   \
    if ( InvokeMethod(name, vResult) ) {    \
        WXAUTOEXCEL_CHECK_VARIANT_TYPE_RET(vResult, resultType, name);   \
    } else  return;

#define WXAUTOEXCEL_CALL_METHOD0(name, resultType, retVal) \
    wxVariant vResult;   \
    if ( InvokeMethod(name, vResult) ) {    \
        WXAUTOEXCEL_CHECK_VARIANT_TYPE(vResult, resultType, name, retVal);   \
    } else return retVal;

#define WXAUTOEXCEL_CALL_METHOD1_RET(name, value, resultType) \
    wxVariant vResult;   \
    if ( InvokeMethod(name, vResult, value) ) {    \
        WXAUTOEXCEL_CHECK_VARIANT_TYPE_RET(vResult, resultType, name);   \
    } else return;

#define WXAUTOEXCEL_CALL_METHOD1(name, value, resultType, retVal) \
    wxVariant vResult;   \
    if ( InvokeMethod(name, vResult, value) ) {    \
        WXAUTOEXCEL_CHECK_VARIANT_TYPE(vResult, resultType, name, retVal);   \
    } else return retVal;


#define WXAUTOEXCEL_CALL_METHOD2_RET(name, value1, value2, resultType) \
    wxVariant vResult;   \
    if ( InvokeMethod(name, vResult, value1, value2) ) {    \
        WXAUTOEXCEL_CHECK_VARIANT_TYPE_RET(vResult, resultType, name);   \
    } else return;

#define WXAUTOEXCEL_CALL_METHOD2(name, value1, value2, resultType, retVal) \
    wxVariant vResult;   \
    if ( InvokeMethod(name, vResult, value1, value2) ) {    \
        WXAUTOEXCEL_CHECK_VARIANT_TYPE(vResult, resultType, name, retVal);   \
    } else return retVal;

#define WXAUTOEXCEL_CALL_METHOD3_RET(name, value1, value2, value3, resultType) \
    wxVariant vResult;   \
    if ( InvokeMethod(name, vResult, value1, value2, value3) ) {    \
        WXAUTOEXCEL_CHECK_VARIANT_TYPE_RET(vResult, resultType, name);   \
    } else return;

#define WXAUTOEXCEL_CALL_METHOD3(name, value1, value2, value3, resultType, retVal) \
    wxVariant vResult;   \
    if ( InvokeMethod(name, vResult, value1, value2, value3) ) {    \
        WXAUTOEXCEL_CHECK_VARIANT_TYPE(vResult, resultType, name, retVal);   \
    } else return retVal;

#define WXAUTOEXCEL_CALL_METHOD4_RET(name, value1, value2, value3, value4, resultType) \
    wxVariant vResult;   \
    if ( InvokeMethod(name, vResult, value1, value2, value3, value4) ) {    \
        WXAUTOEXCEL_CHECK_VARIANT_TYPE_RET(vResult, resultType, name);   \
    } else return;

#define WXAUTOEXCEL_CALL_METHOD4(name, value1, value2, value3, value4, resultType, retVal) \
    wxVariant vResult;   \
    if ( InvokeMethod(name, vResult, value1, value2, value3, value4) ) {    \
        WXAUTOEXCEL_CHECK_VARIANT_TYPE(vResult, resultType, name, retVal);   \
    } else return retVal;

#define WXAUTOEXCEL_CALL_METHOD5_RET(name, value1, value2, value3, value4, value5, resultType) \
    wxVariant vResult;   \
    if ( InvokeMethod(name, vResult, value1, value2, value3, value4, value5) ) {    \
        WXAUTOEXCEL_CHECK_VARIANT_TYPE_RET(vResult, resultType, name);   \
    } else return;

#define WXAUTOEXCEL_CALL_METHOD5(name, value1, value2, value3, value4, value5, resultType, retVal) \
    wxVariant vResult;   \
    if ( InvokeMethod(name, vResult, value1, value2, value3, value4, value5) ) {    \
        WXAUTOEXCEL_CHECK_VARIANT_TYPE(vResult, resultType, name, retVal);   \
    } else return retVal;

#define WXAUTOEXCEL_CALL_METHOD6_RET(name, value1, value2, value3, value4, value5, value6, resultType) \
    wxVariant vResult;   \
    if ( InvokeMethod(name, vResult, value1, value2, value3, value4, value5, value6) ) {    \
        WXAUTOEXCEL_CHECK_VARIANT_TYPE_RET(vResult, resultType, name);   \
    } else return;

#define WXAUTOEXCEL_CALL_METHOD6(name, value1, value2, value3, value4, value5, value6, resultType, retVal) \
    wxVariant vResult;   \
    if ( InvokeMethod(name, vResult, value1, value2, value3, value4, value5, value6) ) {    \
        WXAUTOEXCEL_CHECK_VARIANT_TYPE(vResult, resultType, name, retVal);   \
    } else return retVal;


#define WXAUTOEXCEL_CALL_METHOD0_STRING(name) \
    wxVariant vResult;   \
    if ( InvokeMethod(name, vResult) ) {    \
        WXAUTOEXCEL_CHECK_VARIANT_TYPE(vResult, "string", name, wxS(""));   \
        return vResult.GetString();    \
    } else return wxS("");

#define WXAUTOEXCEL_CALL_METHOD1_STRING(name, value) \
    wxVariant vResult;   \
    if ( InvokeMethod(name, vResult, value) ) {    \
        WXAUTOEXCEL_CHECK_VARIANT_TYPE(vResult, "string", name, wxS(""));   \
        return vResult.GetString();    \
    } else return wxS("");

#define WXAUTOEXCEL_CALL_METHOD0_BOOL(name) \
    wxVariant vResult;   \
    if ( InvokeMethod(name, vResult) ) {    \
        WXAUTOEXCEL_CHECK_VARIANT_TYPE(vResult, "bool", name, false);   \
        return vResult.GetBool();    \
    } else return false;

#define WXAUTOEXCEL_CALL_METHOD1_BOOL(name, value) \
    wxVariant vResult;   \
    if ( InvokeMethod(name, vResult, value) ) {    \
        WXAUTOEXCEL_CHECK_VARIANT_TYPE(vResult, "bool", name, false);   \
        return vResult.GetBool();    \
    } else return false;

#define WXAUTOEXCEL_CALL_METHOD0_LONG(name, retVal) \
    wxVariant vResult;   \
    if ( InvokeMethod(name, vResult) ) {    \
        WXAUTOEXCEL_CHECK_VARIANT_TYPE(vResult, "long", name, retVal);   \
        return vResult.GetLong();    \
    } else return retVal;

#define WXAUTOEXCEL_CALL_METHOD1_LONG(name, value, retVal) \
    wxVariant vResult;   \
    if ( InvokeMethod(name, vResult, value) ) {    \
        WXAUTOEXCEL_CHECK_VARIANT_TYPE(vResult, "long", name, retVal);   \
        return vResult.GetLong();    \
    } else return retVal;

#define WXAUTOEXCEL_CALL_METHOD0_DOUBLE(name, retVal) \
    wxVariant vResult;   \
    if ( InvokeMethod(name, vResult) ) {    \
        WXAUTOEXCEL_CHECK_VARIANT_TYPE(vResult, "double", name, retVal);   \
        return vResult.GetDouble();    \
    } else return retVal;

#define WXAUTOEXCEL_CALL_METHOD1_DOUBLE(name, value, retVal) \
    wxVariant vResult;   \
    if ( InvokeMethod(name, vResult, value) ) {    \
        WXAUTOEXCEL_CHECK_VARIANT_TYPE(vResult, "double", name, retVal);   \
        return vResult.GetLong();    \
    } else return retVal;

#define WXAUTOEXCEL_CALL_METHOD0_OBJECT(name, retVal)      \
    wxVariant vResult;                              \
    if ( InvokeMethod(name, vResult) ) {            \
        WXAUTOEXCEL_CHECK_VARIANT_TYPE(vResult, "void*", name, retVal); \
        VariantToObject(vResult, &retVal);          \
    }                                               \
    return retVal;

#define WXAUTOEXCEL_CALL_METHOD1_OBJECT(name, value, retVal) \
    wxVariant vResult;                                \
    if ( InvokeMethod(name, vResult, value) ) {       \
        WXAUTOEXCEL_CHECK_VARIANT_TYPE(vResult, "void*", name, retVal); \
        VariantToObject(vResult, &retVal);            \
    }                                                 \
    return retVal;

#define WXAUTOEXCEL_CALL_METHOD2_OBJECT(name, value1, value2, retVal) \
    wxVariant vResult;                                \
    if ( InvokeMethod(name, vResult, value1, value2) ) {       \
        WXAUTOEXCEL_CHECK_VARIANT_TYPE(vResult, "void*", name, retVal); \
        VariantToObject(vResult, &retVal);            \
    }                                                 \
    return retVal;

#define WXAUTOEXCEL_CALL_METHODARR_RET(name, args, resultType) \
    wxVariant vResult;   \
    if ( InvokeMethodArray(name, vResult, args) ) {    \
        WXAUTOEXCEL_CHECK_VARIANT_TYPE_RET(vResult, resultType, name);   \
    } else return;

#define WXAUTOEXCEL_CALL_METHODARR(name, args, resultType, retVal) \
    wxVariant vResult;   \
    if ( InvokeMethodArray(name, vResult, args) ) {    \
        WXAUTOEXCEL_CHECK_VARIANT_TYPE(vResult, resultType, name, retVal);   \
    } else return retVal;


#define WXAUTOEXCEL_PROPERTY_GET0(name, resultType, retVal)   \
    wxVariant vResult;   \
    if ( InvokeGetProperty(name, vResult) ) {    \
        WXAUTOEXCEL_CHECK_VARIANT_TYPE(vResult, resultType, name, retVal);   \
    } else return retVal;

#define WXAUTOEXCEL_PROPERTY_GET1(name, value, resultType, retVal)   \
    wxVariant vResult;   \
    if ( InvokeGetProperty(name, vResult, value) ) {    \
        WXAUTOEXCEL_CHECK_VARIANT_TYPE(vResult, resultType, name, retVal);   \
    } else return retVal;

#define WXAUTOEXCEL_PROPERTY_GET2(name, value1, value2, resultType, retVal)   \
    wxVariant vResult;   \
    if ( InvokeGetProperty(name, vResult, value1, value2) ) {    \
        WXAUTOEXCEL_CHECK_VARIANT_TYPE(vResult, resultType, name, retVal);   \
    } else return retVal;

#define WXAUTOEXCEL_PROPERTY_STRING_GET0(name)   \
    WXAUTOEXCEL_PROPERTY_GET0(name, "string", ""); \
    return vResult.GetString();

#define WXAUTOEXCEL_PROPERTY_STRING_GET1(name, value)   \
    WXAUTOEXCEL_PROPERTY_GET1(name, value, "string", ""); \
    return vResult.GetString();

#define WXAUTOEXCEL_PROPERTY_BOOL_GET0(name)   \
    WXAUTOEXCEL_PROPERTY_GET0(name, "bool", false); \
    return vResult.GetBool();

#define WXAUTOEXCEL_PROPERTY_BOOL_GET1(name, value)   \
    WXAUTOEXCEL_PROPERTY_GET1(name, value, "bool", false); \
    return vResult.GetBool();

#define WXAUTOEXCEL_PROPERTY_BOOL_GET2(name, value1, value2)   \
    WXAUTOEXCEL_PROPERTY_GET2(name, value1, value2, "bool", false); \
    return vResult.GetBool();


#define WXAUTOEXCEL_PROPERTY_LONG_GET0(name)   \
    WXAUTOEXCEL_PROPERTY_GET0(name, "long", 0); \
    return vResult.GetLong();

#define WXAUTOEXCEL_PROPERTY_LONG_GET1(name, value)   \
    WXAUTOEXCEL_PROPERTY_GET1(name, value, "long", 0); \
    return vResult.GetLong();

#define WXAUTOEXCEL_PROPERTY_DOUBLE_GET0(name)   \
    WXAUTOEXCEL_PROPERTY_GET0(name, "double", 0.); \
    return vResult.GetDouble();

#define WXAUTOEXCEL_PROPERTY_DOUBLE_GET1(name, value)   \
    WXAUTOEXCEL_PROPERTY_GET1(name, value, "double", 0.); \
    return vResult.GetDouble();


#define WXAUTOEXCEL_PROPERTY_COLOR_GET0(name)              \
    wxColour color;                                 \
    WXAUTOEXCEL_PROPERTY_GET0(name, "long", color);        \
    color.Set(vResult.GetLong());                    \
    return color;

#define WXAUTOEXCEL_PROPERTY_OBJECT_GET0(name, object)   \
    WXAUTOEXCEL_PROPERTY_GET0(name, "void*", object); \
    VariantToObject(vResult, &object);  \
    return object;

#define WXAUTOEXCEL_PROPERTY_OBJECT_GET1(name, value, object)   \
    WXAUTOEXCEL_PROPERTY_GET1(name, value, "void*", object); \
    VariantToObject(vResult, &object);  \
    return object;

#define WXAUTOEXCEL_PROPERTY_OBJECT_GET2(name, value1, value2, object)   \
    WXAUTOEXCEL_PROPERTY_GET2(name, value1, value2, "void*", object); \
    VariantToObject(vResult, &object);  \
    return object;


#define WXAUTOEXCEL_PROPERTY_ENUM_GET0(name, EnumType, retVal)   \
    WXAUTOEXCEL_PROPERTY_GET0(name, "long", retVal); \
    return EnumType(vResult.GetLong());

#define WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT(varName, cppVar) \
    wxVariant WXAUTOEXCEL_JOIN(v, varName);  \
    if ( cppVar != NULL ) { \
         WXAUTOEXCEL_JOIN(v, varName) = *cppVar;   \
    }

#define WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME(varName, cppVar) \
    wxVariant WXAUTOEXCEL_JOIN(v, varName);  \
    if ( cppVar != NULL ) {  \
        WXAUTOEXCEL_JOIN(v, varName).SetName(wxS(#varName));    \
        WXAUTOEXCEL_JOIN(v, varName) = *cppVar;   \
    }

#define WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT(varName, cppVar) \
    wxVariant WXAUTOEXCEL_JOIN(v, varName);  \
    if ( !cppVar.empty() ) { \
        WXAUTOEXCEL_JOIN(v, varName) = cppVar;   \
    }

#define WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME(varName, cppVar) \
    wxVariant WXAUTOEXCEL_JOIN(v, varName);  \
    if ( !cppVar.empty() ) {  \
        WXAUTOEXCEL_JOIN(v, varName).SetName(wxS(#varName));    \
        WXAUTOEXCEL_JOIN(v, varName) = cppVar;   \
    }


#define WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT(varName, cppVar) \
    wxVariant WXAUTOEXCEL_JOIN(v, varName);  \
    if ( !cppVar.IsDefault() ) { \
        WXAUTOEXCEL_JOIN(v, varName) = cppVar.IsTrue();   \
    }

#define WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME(varName, cppVar) \
    wxVariant WXAUTOEXCEL_JOIN(v, varName);  \
    if ( !cppVar.IsDefault() ) { \
        WXAUTOEXCEL_JOIN(v, varName).SetName(wxS(#varName));    \
        WXAUTOEXCEL_JOIN(v, varName) = cppVar.IsTrue();   \
    }


#define WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_VECTOR( cppVar, vect) \
    if ( cppVar != NULL ) {     \
        vect.push_back(wxVariant(*cppVar)) ;     \
    }

#define WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_VECTOR(cppVar, vect) \
    if ( !cppVar.empty() ) {    \
        vect.push_back(wxVariant(cppVar));   \
    }

#define WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_VECTOR(cppVar, vect) \
    if ( !cppVar.IsDefault() ) {    \
        vect.push_back(wxVariant(cppVar.IsTrue()));   \
    }

#define WXAUTOEXCEL_OPTIONALCPP_TO_OPTIONALVARIANT_NAME_VECTOR(varName, cppVar, vect) \
    if ( cppVar != NULL ) {     \
        vect.push_back(wxVariant(*cppVar, wxS(#varName))) ;     \
    }

#define WXAUTOEXCEL_OPTIONALCPPSTR_TO_OPTIONALVARIANT_NAME_VECTOR(varName, cppVar, vect) \
    if ( !cppVar.empty() ) {    \
        vect.push_back(wxVariant(cppVar, wxS(#varName)));   \
    }

#define WXAUTOEXCEL_OPTIONALCPPTBOOL_TO_OPTIONALVARIANT_NAME_VECTOR(varName, cppVar, vect) \
    if ( !cppVar.IsDefault() ) {    \
        vect.push_back(wxVariant(cppVar.IsTrue(), wxS(#varName)));   \
    }


} // namespace wxAutoExcel



#endif // _WXAUTOEXCEL_PRIVATE_H


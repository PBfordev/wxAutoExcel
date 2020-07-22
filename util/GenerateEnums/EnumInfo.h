#pragma once

#include <wx/wx.h>

#include <limits>
#include <vector>

class EnumInfo
{
public:
    struct Field
    {
        wxString name;
        long     value{std::numeric_limits<long>::min()};
        wxString description;
    };

    typedef std::vector<Field> Fields;

    EnumInfo() {}

    bool LoadFromMDFile(const wxString& MDFileName);

    wxString GetName() const;
    wxString GetDescription() const;
    Fields   GetFields() const;

    bool     IsDeprecated() const;

private:
    wxString m_name;
    wxString m_description;
    Fields   m_fields;
    bool     m_isDeprecated{false};

    static bool LoadNameAndDescription(const std::vector<wxString>& lines,
                                       wxString& name, wxString& description,
                                       size_t& startLokingForFieldsLineNo,
                                       wxString& errorInfo);

    static bool LoadFields(const std::vector<wxString>& lines,
                           const size_t startLokingForFieldsLineNo,
                           Fields& fields, wxString& errorInfo);

    static bool CheckDeprecated(const std::vector<wxString>& lines);
};

typedef std::vector<EnumInfo> EnumInfos;

class EnumInfoLoader
{
public:

    static bool LoadEnumInfos(const wxString& MDFileName, EnumInfos& enumInfos);

private:
    static bool LoadEnumList(const wxString& MDFileName, std::vector<wxString>& enumInfoFiles);
    static bool LoadEnums(const std::vector<wxString>& enumInfoFiles, EnumInfos& enumInfos);
};
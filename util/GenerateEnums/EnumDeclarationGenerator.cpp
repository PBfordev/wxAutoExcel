#include "EnumDeclarationGenerator.h"

#include <algorithm>
#include <set>

bool EnumDeclarationGenerator::Generate(const EnumInfos& excelEnums, const EnumInfos& officeEnums,
                                        std::vector<wxString>& excelDeclarations,
                                        std::vector<wxString>& officeDeclarations,
                                        bool omitDeprecated)
{
    wxCHECK_MSG(!excelEnums.empty(), false, "exceEnums cannot be empty");
    wxCHECK_MSG(!officeEnums.empty(), false, "officeEnums cannot be empty");
    wxCHECK_MSG(excelDeclarations.empty(), false, "excelDeclarations must be empty");
    wxCHECK_MSG(officeDeclarations.empty(), false, "officeDeclarations must be empty");

    std::set<wxString> excelEnumNames;
    std::vector<wxString> tmpExcelDeclarations;
    std::vector<wxString> tmpOfficeDeclarations;

    for ( const auto& e : excelEnums )
    {
        if ( e.IsDeprecated() && omitDeprecated )
            continue;

        std::vector<wxString> decl;

        GenerateDeclaration(e, true, decl);
        tmpExcelDeclarations.insert(tmpExcelDeclarations.end(), decl.begin(), decl.end());
        AddEmptyLine(tmpExcelDeclarations);

        excelEnumNames.insert(e.GetName());
    }

    for ( const auto& e : officeEnums )
    {
        if ( e.IsDeprecated() && omitDeprecated )
            continue;

        // some Excel enums are included in Office ones, skip those
        if ( excelEnumNames.find(e.GetName()) != excelEnumNames.end() )
            continue;

        std::vector<wxString> decl;

        GenerateDeclaration(e, false, decl);
        tmpOfficeDeclarations.insert(tmpOfficeDeclarations.end(), decl.begin(), decl.end());
        AddEmptyLine(tmpOfficeDeclarations);
    }

    excelDeclarations = std::move(tmpExcelDeclarations);
    officeDeclarations = std::move(tmpOfficeDeclarations);
    return true;
}

void AddMissingChartTypes(EnumInfo::Fields& fields)
{
    auto AddChartType = [&fields](const auto& name, auto value)
    {
        if ( std::any_of(fields.begin(), 
             fields.end(), [&name](const auto& elem) { return elem.name.CmpNoCase(name) == 0; }) )
        {
            wxLogDebug("The chart type '%s' is already there.", name);
            return;
        }

        EnumInfo::Field f;

        f.name = name; f.value = value; f.description = "not officially documented";
        fields.push_back(f);
    };
 
    AddChartType("xlBoxwhisker", 121);
    AddChartType("xlFunnel", 123);
    AddChartType("xlHistogram", 118);
    AddChartType("xlPareto", 122);
    AddChartType("xlSunburst", 120);
    AddChartType("xlTreemap", 117);
    AddChartType("xlWaterfall", 119);

    std::sort(fields.begin(), fields.end(), 
              [](const auto& a, const auto& b) { return a.name.CmpNoCase(b.name) < 0; });
}

bool EnumDeclarationGenerator::GenerateDeclaration(const EnumInfo& info, bool isExcel,
                                                   std::vector<wxString>& declaration)
{
    const size_t   leftIndentSize = 4; // 4 spaces
    const wxString doxygenCommentStart = "/**";
    const wxString doxygenCommentEnd = "*/";
    const wxString excelURLStart = "https://docs.microsoft.com/office/vba/api/excel.";
    const wxString officeURLStart = "https://docs.microsoft.com/office/vba/api/office.";

    wxCHECK(!info.GetName().empty(), false);
    wxCHECK(declaration.empty(), false);

    EnumInfo::Fields fields = info.GetFields();
    wxString documentationURL;

    int longestName = 0;
    int longestValue = 0;

    declaration.reserve(8 + fields.size());

    declaration.push_back(doxygenCommentStart);

    declaration.push_back(info.GetDescription());
    AddEmptyLine(declaration);

    documentationURL.Printf("[Official VBA documentation for %s](%s%s)",
        info.GetName(),
        isExcel ? excelURLStart : officeURLStart,
        info.GetName().Lower());
    declaration.push_back(documentationURL);

    declaration.push_back(doxygenCommentEnd);

    declaration.push_back(wxString::Format("enum %s", info.GetName()));
    declaration.push_back("{");

    // hotfix for newer chart types which are not documented
    if ( isExcel && info.GetName() == "XlChartType" )
        AddMissingChartTypes(fields);
    
    for ( const auto& v : fields )
    {
        longestName = wxMax(longestName, v.name.size());
        longestValue = wxMax(longestValue, wxString::Format("%ld", v.value).size());
    }

    for ( const auto& v : fields )
    {
        wxString field;

        field.Printf("%-*s = %*ld,", longestName, v.name, longestValue, v.value);
        if ( !v.description.empty() )
            field += wxString::Format(" /*!< %s */", v.description);
        AddLeftIndent(field, leftIndentSize);

        // hotfix for Constants::xlManual which conflicts with XlSortOrder::xlManual
        if ( isExcel && info.GetName() == "Constants" && v.name == "xlManual" )
        {
            declaration.push_back("// xlManual is commented out here to avoid conflict with XlSortOrder::xlManual");
            field.Prepend("//");
        }

        declaration.emplace_back(field);
    }

    declaration.push_back("};");

    return true;
}

void EnumDeclarationGenerator::AddEmptyLine(std::vector<wxString>& declaration)
{
    declaration.push_back(wxString());
}

wxString& EnumDeclarationGenerator::AddLeftIndent(wxString& str, size_t count)
{
    return str.Pad(count, ' ', false);
}

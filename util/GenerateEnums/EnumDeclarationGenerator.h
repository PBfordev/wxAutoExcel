#pragma once

#include "EnumInfo.h"


class EnumDeclarationGenerator
{
public:
    static bool Generate(const EnumInfos& excelEnums, const EnumInfos& officeEnums,
                         std::vector<wxString>& excelDeclarations,
                         std::vector<wxString>& officeDeclarations,
                         bool omitDeprecated = true);

private:
    static bool GenerateDeclaration(const EnumInfo& info, bool isExcel,
                                    std::vector<wxString>& declaration);

    static void AddEmptyLine(std::vector<wxString>& declaration);
    static wxString& AddLeftIndent(wxString& str, size_t count);
};

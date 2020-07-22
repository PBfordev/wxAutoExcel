#include "EnumInfo.h"

#include <wx/filename.h>
#include <wx/textfile.h>

/*************************************
class EnumInfo
*************************************/

// Loads enum information from the official VBA documentation

bool EnumInfo::LoadFromMDFile(const wxString& MDFileName)
{
    wxTextFile enumFile;

    if ( !enumFile.Open(MDFileName) )
        return false;

    std::vector<wxString> lines;
    wxString name, description;
    Fields fields;
    bool isDeprecated = false;
    size_t startLokingForfieldsLineNo = 0;
    wxString errorInfo;

    lines.reserve(enumFile.GetLineCount());
    for ( size_t i = 0; i < enumFile.GetLineCount(); ++i )
        lines.push_back(enumFile[i]);

    if ( !LoadNameAndDescription(lines, name, description, startLokingForfieldsLineNo, errorInfo) )
    {
        wxLogError("File '%s': %s.", MDFileName, errorInfo);
        return false;
    }

    isDeprecated = CheckDeprecated(lines);
    if ( isDeprecated )
        wxLogDebug("Enum '%s' is deprecated.", name);

    if ( !LoadFields(lines, startLokingForfieldsLineNo, fields, errorInfo) )
    {
       if ( !isDeprecated )
       {
            wxLogError("File '%s': %s.", MDFileName, errorInfo);
            return false;
       }
    }

    m_name = name;
    m_description = description;
    m_fields = std::move(fields);
    m_isDeprecated = isDeprecated;

    return true;
}

wxString EnumInfo::GetName() const
{
    return m_name;
}

wxString EnumInfo::GetDescription() const
{
    return m_description;
}

EnumInfo::Fields EnumInfo::GetFields() const
{
    return m_fields;
}

bool EnumInfo::IsDeprecated() const
{
    return m_isDeprecated;
}

// Find name and description of this enum
bool EnumInfo::LoadNameAndDescription(const std::vector<wxString>& lines,
                                      wxString& name, wxString& description,
                                      size_t& startLokingForfieldsLine,
                                      wxString& errorInfo)
{
    wxString tmpName, tmpDescription;

    for ( size_t i = 0; i < lines.size(); ++i )
    {
        const wxString& line = lines[i];
        wxString s1 = line;
        wxString s2;
        size_t descriptionLineIdx = 0;

        s1.Trim(true).Trim(false);

        if ( s1.StartsWith("# ", &s2)
             && s2.Contains(" enumeration (") )
        {
            s2.Trim(true).Trim(false);
            s2 = s2.BeforeFirst(' ');
            if ( s2.empty() )
                continue;

            tmpName = s2;
            // there is one empty line and the line after
            // that contains description

            descriptionLineIdx = i + 2;

            while ( descriptionLineIdx < lines.size()
                    && !lines[descriptionLineIdx].empty() )
            {
                tmpDescription += lines[descriptionLineIdx++];
            }

            if ( tmpDescription.empty() )
            {
                errorInfo.Printf("Could not obtain description for enum '%s' on line %zu", name, i + 1 + 2);
                return false;
            }
            startLokingForfieldsLine = descriptionLineIdx;

            break;
        }
    }

    if ( tmpName.empty() )
    {
        errorInfo.Printf("Could not find enum name");
        return false;
    }

    name = tmpName;
    description = tmpDescription;
    return true;
}

// Parse enum individual fields stored in a markdown file, where a line for an enum field looks like
// "|**BroadcastCapFileSizeLimited**|**1**|The size of the file being broadcasted is limited.|"
// where the first column has field name, second field value, and third field description
// Unfortunately, the files sometimes do not obey this format exactly, so the code has workarounds for
// the inconstencies I have encountered so far.
bool EnumInfo::LoadFields(const std::vector<wxString>& lines,
                          const size_t startLokingForfieldsLine,
                          Fields& fields, wxString& errorInfo)
{
    static const size_t minSplitCount = 4;

    static const size_t nameIndex        = 1;
    static const size_t valueIndex       = 2;
    static const size_t descriptionIndex = 3;

    Fields tmpfields;

    for ( size_t i = startLokingForfieldsLine;
          i < lines.size();
          ++i )
    {
        const wxString& line = lines[i];

        if ( line.StartsWith("|") ) // can be an enum field, but also the table markdown
        {
            wxArrayString as = wxSplit(line, '|');
            wxString s;
            long l;
            Field field;

            if ( as.size() < minSplitCount )
            {
                continue;
            }

            s = as[nameIndex];
            s.Replace("*", "");
            s.Trim(true).Trim(false);
            if (s.find(' ') )
                s = s.BeforeFirst(' ');

            if ( s.empty() )
            {
                errorInfo.Printf("Could not parse enum field name at line %zu", i + 1);
                return false;
            }
            field.name = s;

            s = as[valueIndex];
            s.Replace("*", "");
            s.Trim(true).Trim(false);
            if (s.find(' ') )
                s = s.BeforeFirst(' ');
            if ( !s.ToCLong(&l) )
            {
                continue; // probably not a line with enum field
            }
            field.value = l;

            s = as[descriptionIndex];
            s.Trim(true).Trim(false);
            field.description = s;

            tmpfields.emplace_back(field);
        }
        else if ( !tmpfields.empty() )
        {
            // We already parsed the block we are interested in;
            // do not process any more lines, which can also
            // contain different information but looking like
            // lines we are interested in.
            break;
        }
    }

    if ( tmpfields.empty() )
    {
        errorInfo.Printf("Could not find any enum fields");
        return false;
    }

    fields = std::move(tmpfields);
    return true;
}

bool EnumInfo::CheckDeprecated(const std::vector<wxString>& lines)
{
    // Check if the line contains " deprecated" anywhere but
    // in field description
    for ( const auto& l : lines )
    {
        wxString s = l;

        s.Trim(false);
        if ( !s.StartsWith('|') && s.Contains(" deprecated") )
            return true;
    }

    return false;
}


/*************************************
class EnumInfoLoader
*************************************/

bool EnumInfoLoader::LoadEnumInfos(const wxString& MDFileName, EnumInfos& enumInfos)
{
    std::vector<wxString> files;
    EnumInfos infos;

    LoadEnumList(MDFileName, files);
    LoadEnums(files, infos);

    enumInfos = std::move(infos);
    return !enumInfos.empty();
}

bool EnumInfoLoader::LoadEnumList(const wxString& MDFileName, std::vector<wxString>& enumInfoFiles)
{
    wxTextFile enumListFile;

    if ( !enumListFile.Open(MDFileName) )
        return false;

    std::vector<wxString> files;
    size_t numErrors = 0;

    wxLogMessage("Loading list of files with enum descriptions from '%s'...", MDFileName);

    // Enums are listed with the enum name in brackets and the name
    // of enum documentation file in parentheses, like this
    // "- [BackstageGroupStyle](../../Office.BackstageGroupStyle.md)"
    for ( size_t i = 0; i < enumListFile.GetLineCount(); ++i )
    {
        const wxString& line = enumListFile[i];

        if ( !line.StartsWith("- [") )
        {
            if ( !files.empty() )
            {
                // We already parsed the block we are interested in;
                // do not process any more lines, which can also
                // contain different information but looking like
                // lines we are interested in.
                break;
            }
            else
            {
                continue;
            }
        }

        const wxString enumName = line.AfterFirst('[').BeforeFirst(']');
        const wxString enumFileName = line.AfterFirst('(').BeforeLast(')');

        if ( enumName.empty() )
        {
            wxLogError("Could not obtain enum name (line %zu)", i + 1);
            ++numErrors;
            continue;
        }

        if ( enumFileName.empty() )
        {
            wxLogError("Could not obtain enum file name (line %zu)", i + 1);
            ++numErrors;
            continue;
        }

        wxFileName enumFileNameFN(enumFileName);

        if ( !enumFileNameFN.Normalize(wxPATH_NORM_ABSOLUTE | wxPATH_NORM_DOTS,
                                       wxFileName(MDFileName).GetPath()) )
        {
            wxLogError("Could not normalize file name '%s' (line %zu)", enumFileName, i + 1);
            ++numErrors;
            continue;
        }

        files.emplace_back(enumFileNameFN.GetFullPath());
    }

    wxLogMessage("Errors encountered: %zu, files identified: %zu.",
        numErrors, files.size());

    enumInfoFiles = std::move(files);

    return !enumInfoFiles.empty() && numErrors == 0;
}

bool EnumInfoLoader::LoadEnums(const std::vector<wxString>& enumInfoFiles, EnumInfos& enumInfos)
{
    wxCHECK_MSG(!enumInfoFiles.empty(), false, "enumInfoFiles cannot be empty");
    wxCHECK_MSG(enumInfos.empty(), false, "enumInfos must be empty");

    EnumInfos infos;
    size_t numErrors = 0;
    size_t deprecatedCount = 0;

    wxLogMessage("Parsing enums from %zu files...", enumInfoFiles.size());

    for ( const auto& infoFile : enumInfoFiles )
    {
        EnumInfo info;

        if ( !info.LoadFromMDFile(infoFile) )
        {
            ++numErrors;
            continue;
        }

        if ( info.IsDeprecated() )
            ++deprecatedCount;

        infos.emplace_back(info);
    }

    wxLogMessage("Errors encountered: %zu, enums parsed: %zu (%zu deprecated).",
        numErrors, infos.size(), deprecatedCount);

    enumInfos = std::move(infos);

    return !enumInfos.empty() && numErrors == 0;
}
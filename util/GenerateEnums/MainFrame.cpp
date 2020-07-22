#include "MainFrame.h"
#include "EnumDeclarationGenerator.h"

#include "version.h"

#include <wx/config.h>
#include <wx/filename.h>
#include <wx/stdpaths.h>
#include <wx/textfile.h>
#include <wx/utils.h>

#include <vector>

MainFrame::MainFrame()
    : wxFrame(nullptr, wxID_ANY, "")
{
    SetIcons(wxIconBundle("appIcon", NULL));
    SetTitle(wxString::Format("%s %s", APP_NAME_STR, APP_VERSION_NUM_DOT_STRING));
    SetMinClientSize(FromDIP(wxSize(800, 600)));

    LoadOptions();

    wxPanel* mainPanel = new wxPanel(this);
    wxBoxSizer* mainPanelSizer = new wxBoxSizer(wxVERTICAL);
    wxBoxSizer* buttonsSizer = new wxBoxSizer(wxHORIZONTAL);
    wxButton* button = nullptr;

    button = new wxButton(mainPanel, wxID_ANY, "Load Microsoft Excel Enums...");
    buttonsSizer->Add(button, wxSizerFlags().Proportion(1).Expand().Border());
    button->Bind(wxEVT_COMMAND_BUTTON_CLICKED, &MainFrame::OnLoadExcelEnums, this);

    button = new wxButton(mainPanel, wxID_ANY, "Load Microsoft Office Enums...");
    buttonsSizer->Add(button, wxSizerFlags().Proportion(1).Expand().Border());
    button->Bind(wxEVT_COMMAND_BUTTON_CLICKED, &MainFrame::OnLoadOfficeEnums, this);

    m_ctlGenerateBtn = new wxButton(mainPanel, wxID_ANY, "Generate Enum Declarations...");
    buttonsSizer->Add(m_ctlGenerateBtn, wxSizerFlags().Proportion(1).Expand().Border());
    m_ctlGenerateBtn->Bind(wxEVT_COMMAND_BUTTON_CLICKED, &MainFrame::OnGenerateDeclarations, this);

    mainPanelSizer->Add(buttonsSizer, wxSizerFlags().Expand());

    wxTextCtrl* logCtrl = new wxTextCtrl(mainPanel, wxID_ANY, wxEmptyString, wxDefaultPosition, wxDefaultSize,
        wxTE_MULTILINE | wxTE_READONLY | wxTE_RICH2);
    wxLog::SetActiveTarget(new wxLogTextCtrl(logCtrl));
    mainPanelSizer->Add(logCtrl, wxSizerFlags().Proportion(1).Expand().Border());

    mainPanel->SetSizer(mainPanelSizer);

    Bind(wxEVT_UPDATE_UI, &MainFrame::OnUpdateUI, this);
}

MainFrame::~MainFrame()
{
    SaveOptions();
}

void MainFrame::LoadOptions()
{
    const wxConfigBase* config = wxConfigBase::Get();
    wxConfigPathChanger changer(config, "/");

    m_lastFolderExcel = config->Read("Excel Enums Last Folder");
    m_lastFolderOffice = config->Read("Office Enums Last Folder");
    m_lastFolderGenerated = config->Read("Generated File Last Folder");
}

void MainFrame::SaveOptions()
{
    wxConfigBase* config = wxConfigBase::Get();
    wxConfigPathChanger changer(config, "/");

    config->Write("Excel Enums Last Folder", m_lastFolderExcel);
    config->Write("Office Enums Last Folder", m_lastFolderOffice);
    config->Write("Generated File Last Folder", m_lastFolderGenerated);
}

// The file should be in VBA_DOCS_ROOT\api
void MainFrame::OnLoadExcelEnums(wxCommandEvent&)
{
    LoadEnums("Excel", this, "Excel(enumerations).md", m_lastFolderExcel, m_enumsExcel);
}

// The file should be in VBA_DOCS_ROOT\api\overview\Library-Reference
void MainFrame::OnLoadOfficeEnums(wxCommandEvent&)
{
    LoadEnums("Office", this, "enumerations-office.md", m_lastFolderOffice, m_enumsOffice);
}

void MainFrame::OnGenerateDeclarations(wxCommandEvent&)
{
    wxCHECK_RET(!m_enumsExcel.empty() && !m_enumsOffice.empty(), "Enums cannot be empty");

    const wxString fileName = wxFileSelector("Save Enumerations Declarations As",
                                              m_lastFolderGenerated, "Excel and Office Enums", "",
                                              "Text Files (*.txt)|*.txt",
                                               wxFD_SAVE |  wxFD_OVERWRITE_PROMPT,
                                               this);

    if ( fileName.empty() )
        return;

    m_lastFolderGenerated = wxFileName(fileName).GetPath();

    std::vector<wxString> excelDeclarations;
    std::vector<wxString> officeDeclarations;
    wxTextFile textFile;

    if ( wxFileExists(fileName) )
    {
        if ( !textFile.Open(fileName) )
            return;
    }
    else if ( !textFile.Create(fileName) )
    {
        return;
    }

    wxLogMessage("Generating declarations...");
    if ( !EnumDeclarationGenerator::Generate(m_enumsExcel, m_enumsOffice,
            excelDeclarations, officeDeclarations) )
    {
        wxLogError("Could not generate enum declarations.");
        return;
    }

    wxLogMessage("Writing declarations into '%s'...", fileName);

    textFile.Clear();

    textFile.AddLine("The text below should replace contents of wxAutoExcel_enums.h");
    textFile.AddLine("inside the namespace wxAutoExcel block");
    textFile.AddLine("---8<------------------------------------------------");
    textFile.AddLine("");

    textFile.AddLine("/*************************************");
    textFile.AddLine("    Microsoft Excel enumerations");
    textFile.AddLine("*************************************/");
    textFile.AddLine("");
    for ( const auto& d : excelDeclarations )
        textFile.AddLine(d);

    textFile.AddLine("");
    textFile.AddLine("");
    textFile.AddLine("/*************************************");
    textFile.AddLine("    Microsoft Office enumerations");
    textFile.AddLine("*************************************/");
    textFile.AddLine("");
    for ( const auto& d : officeDeclarations )
        textFile.AddLine(d);

    textFile.Write();

    wxLogMessage("Finished writing declarations.");

    if ( wxMessageBox("Do you want to open file with declarations?", "Question",
            wxYES_NO, this) == wxYES )
    {
        wxLaunchDefaultApplication(fileName);
    }
}

void MainFrame::OnUpdateUI(wxUpdateUIEvent&)
{
    m_ctlGenerateBtn->Enable(!m_enumsExcel.empty() && !m_enumsOffice.empty());
}

void MainFrame::LoadEnums(const wxString& enumGroup, wxWindow* parent,
                          const wxString& defaultFileName, wxString& defaultDir,
                          EnumInfos& enums)
{
    const wxString fileName = wxFileSelector(wxString::Format("Select File With List of Microsoft %s Enumerations", enumGroup),
                                             defaultDir, defaultFileName, "",
                                             "Markdown Files (*.md)|*.md",
                                             wxFD_OPEN |  wxFD_FILE_MUST_EXIST,
                                             parent);

    if ( fileName.empty() )
        return;

    defaultDir = wxFileName(fileName).GetPath();

    enums.clear();
    EnumInfoLoader::LoadEnumInfos(fileName, enums);
}
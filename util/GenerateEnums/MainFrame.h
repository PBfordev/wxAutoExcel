#pragma once

#include <wx/wx.h>

#include "EnumInfo.h"

class MainFrame : public wxFrame
{
public:
    MainFrame();
    ~MainFrame();

private:
    EnumInfos m_enumsExcel;
    EnumInfos m_enumsOffice;

    wxString m_lastFolderExcel;
    wxString m_lastFolderOffice;
    wxString m_lastFolderGenerated;

    wxButton* m_ctlGenerateBtn;

    void LoadOptions();
    void SaveOptions();

    void OnLoadExcelEnums(wxCommandEvent&);
    void OnLoadOfficeEnums(wxCommandEvent&);
    void OnGenerateDeclarations(wxCommandEvent&);

    void OnUpdateUI(wxUpdateUIEvent&);

    static void LoadEnums(const wxString& enumGroup, wxWindow* parent,
                          const wxString& defaultFileName, wxString& defaultDir,
                          EnumInfos& enums);
};

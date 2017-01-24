######################################################################
# Author:      PB
# Purpose:     Parse the version number of wxAutoExcel from include/wx/wxAutoExcel_version.h
# Copyright:   (c) 2017 PB <pbfordev@gmail.com>
# Licence:     wxWindows licence
######################################################################

file(READ include/wx/wxAutoExcel_version.h WXAUTOEXCEL_VERSION_H_CONTENTS)
string(REGEX MATCH "WXAUTOEXCEL_MAJOR_VERSION[ \t]+([0-9]+)"
    wxAutoExcel_MAJOR_VERSION ${WXAUTOEXCEL_VERSION_H_CONTENTS})
string (REGEX MATCH "([0-9]+)"
    wxAutoExcel_MAJOR_VERSION ${wxAutoExcel_MAJOR_VERSION})
string(REGEX MATCH "WXAUTOEXCEL_MINOR_VERSION[ \t]+([0-9]+)"
    wxAutoExcel_MINOR_VERSION ${WXAUTOEXCEL_VERSION_H_CONTENTS})
string (REGEX MATCH "([0-9]+)"
    wxAutoExcel_MINOR_VERSION ${wxAutoExcel_MINOR_VERSION})
string(REGEX MATCH "WXAUTOEXCEL_RELEASE_NUMBER[ \t]+([0-9]+)"
    wxAutoExcel_RELEASE_NUMBER ${WXAUTOEXCEL_VERSION_H_CONTENTS})  
string (REGEX MATCH "([0-9]+)"
    wxAutoExcel_RELEASE_NUMBER ${wxAutoExcel_RELEASE_NUMBER})
set(wxAutoExcel_VERSION_STRING ${wxAutoExcel_MAJOR_VERSION}.${wxAutoExcel_MINOR_VERSION}.${wxAutoExcel_RELEASE_NUMBER}) 
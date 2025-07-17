######################################################################
# Author:      PB
# Purpose:     Parse the version number of wxAutoExcel from include/wx/wxAutoExcel_version.h
# Copyright:   (c) 2017 PB <pbfordev@gmail.com>
# License:     MIT license
######################################################################

file(READ "${CMAKE_CURRENT_SOURCE_DIR}/include/wx/wxAutoExcel_version.h" WXAUTOEXCEL_VERSION_H_CONTENTS)

string(REGEX MATCH "WXAUTOEXCEL_MAJOR_VERSION[ \t]+([0-9]+)"
    wxAutoExcel_MAJOR_VER ${WXAUTOEXCEL_VERSION_H_CONTENTS})
string(REGEX MATCH "([0-9]+)"
    wxAutoExcel_MAJOR_VER ${wxAutoExcel_MAJOR_VER})
string(REGEX MATCH "WXAUTOEXCEL_MINOR_VERSION[ \t]+([0-9]+)"
    wxAutoExcel_MINOR_VER ${WXAUTOEXCEL_VERSION_H_CONTENTS})
string (REGEX MATCH "([0-9]+)"
    wxAutoExcel_MINOR_VER ${wxAutoExcel_MINOR_VER})
string(REGEX MATCH "WXAUTOEXCEL_RELEASE_NUMBER[ \t]+([0-9]+)"
    wxAutoExcel_REL_NUM ${WXAUTOEXCEL_VERSION_H_CONTENTS})  
string (REGEX MATCH "([0-9]+)"
    wxAutoExcel_REL_NUM ${wxAutoExcel_REL_NUM})
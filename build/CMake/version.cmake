######################################################################
# Author:      PB
# Purpose:     Parse the version number of wxAutoExcel from include/wx/wxAutoExcel_version.h
# Copyright:   (c) 2017 PB <pbfordev@gmail.com>
# License:     MIT license
######################################################################

get_filename_component(wxAutoExcel_VERSION_H "${CMAKE_CURRENT_LIST_DIR}/../../include/wx/wxAutoExcel_version.h" ABSOLUTE)
file(READ "${wxAutoExcel_VERSION_H}" WXAUTOEXCEL_VERSION_H_CONTENTS)

foreach(versionPart MAJOR_VER:MAJOR_VERSION MINOR_VER:MINOR_VERSION REL_NUM:RELEASE_NUMBER)
  string(REPLACE ":" ";" versionPart "${versionPart}")
  list(GET versionPart 0 versionVariable)
  list(GET versionPart 1 versionDefine)
  if("${WXAUTOEXCEL_VERSION_H_CONTENTS}" MATCHES "#define[ \t]+WXAUTOEXCEL_${versionDefine}[ \t]+([0-9]+)")
    set(wxAutoExcel_${versionVariable} "${CMAKE_MATCH_1}")
  else()
    message(FATAL_ERROR "WXAUTOEXCEL_${versionDefine} not found in ${wxAutoExcel_VERSION_H}")
  endif()
endforeach()

unset(versionPart)
unset(versionVariable)
unset(versionDefine)
unset(WXAUTOEXCEL_VERSION_H_CONTENTS)

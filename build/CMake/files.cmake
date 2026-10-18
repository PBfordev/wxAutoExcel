######################################################################
# Author:      PB
# Purpose:     List of all .h and .cpp files for wxAutoExcel library
# Copyright:   (c) 2017 PB <pbfordev@gmail.com>
# License:     MIT license
######################################################################

get_filename_component(wxAutoExcel_ROOT_DIR "${CMAKE_CURRENT_LIST_DIR}/../.." ABSOLUTE)

# CONFIGURE_DEPENDS makes the build rerun CMake when a file is added
# to or removed from these folders, so the list needs no maintenance
file(GLOB SRCS CONFIGURE_DEPENDS
  "${wxAutoExcel_ROOT_DIR}/include/wx/*.h"
  "${wxAutoExcel_ROOT_DIR}/src/*.cpp"
)

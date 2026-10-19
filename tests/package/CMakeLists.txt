# Checks that an installed wxAutoExcel package can be used with
# find_package(). It is built separately from wxAutoExcel, with
# CMAKE_PREFIX_PATH (or wxAutoExcel_DIR) pointing to the package.

cmake_minimum_required(VERSION 3.27)

project(wxAutoExcel_package_test LANGUAGES CXX)

# require the version of this source tree, which the package must match
# in the major and minor version
include("${CMAKE_CURRENT_LIST_DIR}/../../build/CMake/version.cmake")
find_package(wxAutoExcel ${wxAutoExcel_MAJOR_VER}.${wxAutoExcel_MINOR_VER} REQUIRED)

add_executable(wxAutoExcel_package_test main.cpp)
target_link_libraries(wxAutoExcel_package_test PRIVATE wxAutoExcel::wxAutoExcel)

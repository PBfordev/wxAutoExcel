# Checks that an installed wxAutoExcel package can be used with
# find_package(). It is built separately from wxAutoExcel, with
# CMAKE_PREFIX_PATH (or wxAutoExcel_DIR) pointing to the package.

cmake_minimum_required(VERSION 3.27)

project(wxAutoExcel_package_test LANGUAGES CXX)

find_package(wxAutoExcel 2.1 REQUIRED)

add_executable(wxAutoExcel_package_test main.cpp)
target_link_libraries(wxAutoExcel_package_test PRIVATE wxAutoExcel::wxAutoExcel)

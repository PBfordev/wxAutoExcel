######################################################################
# Author:      PB
# Purpose:     Primary CMake for wxAutoExcel
# Copyright:   (c) 2017 PB <pbfordev@gmail.com>
# License:     MIT license
######################################################################

cmake_minimum_required(VERSION 3.27)

# parses the wxAutoExcel version from include/wx/wxAutoExcel_version.h
# into ${wxAutoExcel_MAJOR_VER}, ${wxAutoExcel_MINOR_VER}, and ${wxAutoExcel_REL_NUM}
include("${CMAKE_CURRENT_SOURCE_DIR}/build/CMake/version.cmake")

project(wxAutoExcel
        VERSION ${wxAutoExcel_MAJOR_VER}.${wxAutoExcel_MINOR_VER}.${wxAutoExcel_REL_NUM}
        DESCRIPTION "wxWidgets library making automating Microsoft Excel easier"
        HOMEPAGE_URL https://github.com/PBfordev/wxAutoExcel
        LANGUAGES CXX RC
)

if(NOT WIN32)
  message(FATAL_ERROR "wxAutoExcel is available only for Microsoft Windows")
endif()

######################################################################
#
# Global user-customizable options
#
######################################################################

option(wxAutoExcel_BUILD_SHARED "Build wxAutoExcel libraries as shared libs" ${BUILD_SHARED_LIBS})
set(wxAutoExcel_BUILD_USE_STATIC_RUNTIME_DEFAULT OFF)
if(MSVC AND CMAKE_MSVC_RUNTIME_LIBRARY AND NOT CMAKE_MSVC_RUNTIME_LIBRARY MATCHES "DLL")
  set(wxAutoExcel_BUILD_USE_STATIC_RUNTIME_DEFAULT ON)
endif()
option(wxAutoExcel_BUILD_USE_STATIC_RUNTIME "Link using the static runtime library" ${wxAutoExcel_BUILD_USE_STATIC_RUNTIME_DEFAULT})
option(wxAutoExcel_BUILD_USE_PRECOMPILED "Use precompiled headers" ON)
set(wxAutoExcel_BUILD_VENDOR "custom" CACHE STRING "Vendor")
option(wxAutoExcel_BUILD_INSTALL "Create install target for the library" OFF)

option(wxAutoExcel_BUILD_SAMPLES "Build samples" ${PROJECT_IS_TOP_LEVEL})
option(wxAutoExcel_BUILD_TESTS "Build tests" OFF)

# wxAutoExcel requires wxWidgets 3.1 and higher. FindwxWidgets provides the
# target-based wxWidgets::wxWidgets interface starting with CMake 3.27.
find_package(wxWidgets 3.1 COMPONENTS core base REQUIRED)
set(wxAutoExcel_WX_VERSION "${wxWidgets_VERSION}")
if(NOT wxAutoExcel_WX_VERSION)
  set(wxAutoExcel_WX_VERSION "${wxWidgets_VERSION_STRING}")
endif()
if(wxAutoExcel_WX_VERSION AND wxAutoExcel_WX_VERSION VERSION_LESS 3.1)
  message(FATAL_ERROR "wxAutoExcel requires wxWidgets 3.1 or newer")
endif()

include(GNUInstallDirs)
include("${CMAKE_CURRENT_SOURCE_DIR}/build/CMake/wxAutoExcelDetectWxLinkage.cmake")
wxAutoExcel_detect_wx_shared(wxAutoExcel_WX_SHARED)


######################################################################
#
# Compiler and linker options
#
######################################################################

function(wxAutoExcel_adjust_target_flags target)

  target_compile_definitions(${target} PRIVATE __WXMSW__ wxNO_UNSAFE_WXSTRING_CONV)

  if(DEFINED wxAutoExcel_BUILD_USE_STATIC_RUNTIME AND wxAutoExcel_BUILD_USE_STATIC_RUNTIME)
    if(MSVC)
      set_target_properties(${target} PROPERTIES MSVC_RUNTIME_LIBRARY "MultiThreaded$<$<CONFIG:Debug>:Debug>")
    elseif(MINGW)
      target_link_options(${target} PRIVATE -static)
    else() # clang
      target_link_options(${target} PRIVATE -static-libgcc -static-libstdc++)
    endif()
  endif()
  
  if(MSVC)
    # suppress security warnings
    target_compile_definitions(${target} PRIVATE _CRT_SECURE_NO_DEPRECATE _CRT_NON_CONFORMING_SWPRINTFS _SCL_SECURE_NO_WARNINGS)
    target_compile_options(${target} PRIVATE /MP /W4)
  else() # GCC or clang
    target_compile_options(${target} PRIVATE -Wno-deprecated-declarations)
  endif()

endfunction()

if(wxAutoExcel_BUILD_SHARED)
  set(wxAutoExcel_BUILD_LIB_TYPE SHARED)
else()
  set(wxAutoExcel_BUILD_LIB_TYPE STATIC)
endif()

if(DEFINED wxAutoExcel_BUILD_USE_STATIC_RUNTIME AND wxAutoExcel_BUILD_USE_STATIC_RUNTIME
   AND (wxAutoExcel_BUILD_SHARED OR wxAutoExcel_WX_SHARED))
     message(FATAL_ERROR "The static runtime cannot be used when creating wxAutoExcel DLLs or linking with wxWidgets dynamically")
endif()

# Initialize variables for quick access to wxAutoExcel dirs
set(wxAutoExcel_RUNTIME_DIR "${PROJECT_BINARY_DIR}/bin")
set(wxAutoExcel_ARCHIVE_DIR "${PROJECT_BINARY_DIR}/lib")


######################################################################
#
# wxAutoExcel library
#
######################################################################

# list of .cpp files for the wxAutoExcel library
include("${wxAutoExcel_SOURCE_DIR}/build/CMake/files.cmake")

if(wxAutoExcel_BUILD_SHARED)
  list(APPEND SRCS "${CMAKE_CURRENT_SOURCE_DIR}/src/wxAutoExcel_version.rc")
endif()

add_library(wxAutoExcel ${wxAutoExcel_BUILD_LIB_TYPE} ${SRCS})
add_library(wxAutoExcel::wxAutoExcel ALIAS wxAutoExcel)
target_include_directories(wxAutoExcel PUBLIC
  $<BUILD_INTERFACE:${CMAKE_CURRENT_SOURCE_DIR}/include>
  $<INSTALL_INTERFACE:${CMAKE_INSTALL_INCLUDEDIR}>
)
target_link_libraries(wxAutoExcel PUBLIC wxWidgets::wxWidgets)

wxAutoExcel_adjust_target_flags(wxAutoExcel)

if(wxAutoExcel_BUILD_SHARED)
  target_compile_definitions(wxAutoExcel PRIVATE WXMAKINGDLL_WXAUTOEXCEL)
  # WXUSINGDLL_WXAUTOEXCEL must be defined when building against shared wxAutoExcel build
  target_compile_definitions(wxAutoExcel INTERFACE WXUSINGDLL_WXAUTOEXCEL)
endif()

if(wxAutoExcel_BUILD_USE_PRECOMPILED)
  target_precompile_headers(wxAutoExcel PRIVATE "${CMAKE_CURRENT_SOURCE_DIR}/include/wx/wxAutoExcel_prec.h")
endif()

set(wxAutoExcel_OUTPUT_NAME_MAIN "wxAutoExcel${wxAutoExcel_VERSION_MAJOR}${wxAutoExcel_VERSION_MINOR}")
set_target_properties(wxAutoExcel PROPERTIES
  OUTPUT_NAME "${wxAutoExcel_OUTPUT_NAME_MAIN}"
  OUTPUT_NAME_DEBUG "${wxAutoExcel_OUTPUT_NAME_MAIN}d"
)

set (wxAutoExcel_DLL_NAME_MAIN "wxAutoExcel${wxAutoExcel_VERSION_MAJOR}${wxAutoExcel_VERSION_MINOR}${wxAutoExcel_VERSION_PATCH}")
set (wxAutoExcel_DLL_NAME_FLAVOR "")

if(MSVC)
  set(wxAutoExcel_COMPILER_ID_STR "vc${MSVC_TOOLSET_VERSION}")
else()
  # remove dots from compiler version, i.e. change e.g. "11.3.0" to "1130"
  if("${CMAKE_CXX_COMPILER_ID}" STREQUAL "GNU")
    string(REPLACE "." "" wxAutoExcel_COMPILER_ID_STR "gcc${CMAKE_CXX_COMPILER_VERSION}")
  elseif("${CMAKE_CXX_COMPILER_ID}" STREQUAL "Clang")
    string(REPLACE "." "" wxAutoExcel_COMPILER_ID_STR "clang${CMAKE_CXX_COMPILER_VERSION}")
  else()
    message(FATAL_ERROR "Unsupported compiler type: " "${CMAKE_CXX_COMPILER_ID}")
  endif()
endif()

if(CMAKE_GENERATOR_PLATFORM)
  set(wxAutoExcel_ARCH_SOURCE "${CMAKE_GENERATOR_PLATFORM}")
elseif(CMAKE_VS_PLATFORM_NAME)
  set(wxAutoExcel_ARCH_SOURCE "${CMAKE_VS_PLATFORM_NAME}")
elseif(CMAKE_CXX_COMPILER_ARCHITECTURE_ID)
  set(wxAutoExcel_ARCH_SOURCE "${CMAKE_CXX_COMPILER_ARCHITECTURE_ID}")
else()
  set(wxAutoExcel_ARCH_SOURCE "${CMAKE_SYSTEM_PROCESSOR}")
endif()
string(TOLOWER "${wxAutoExcel_ARCH_SOURCE}" wxAutoExcel_ARCH_SOURCE)
string(REGEX REPLACE ",.*$" "" wxAutoExcel_ARCH_SOURCE
       "${wxAutoExcel_ARCH_SOURCE}")

if(wxAutoExcel_ARCH_SOURCE STREQUAL "arm64ec")
  set(wxAutoExcel_ARCH_ID "arm64ec")
elseif(wxAutoExcel_ARCH_SOURCE MATCHES "^(arm64|aarch64)$")
  set(wxAutoExcel_ARCH_ID "arm64")
elseif(wxAutoExcel_ARCH_SOURCE MATCHES "^(arm|arm32)$")
  set(wxAutoExcel_ARCH_ID "arm32")
elseif(wxAutoExcel_ARCH_SOURCE MATCHES "^(amd64|x86_64|x64)$")
  set(wxAutoExcel_ARCH_ID "x64")
elseif(wxAutoExcel_ARCH_SOURCE MATCHES "^(win32|x86|i[3-6]86)$")
  set(wxAutoExcel_ARCH_ID "x32")
else()
  message(FATAL_ERROR
          "Unsupported or unknown target architecture: ${wxAutoExcel_ARCH_SOURCE}")
endif()

string(APPEND wxAutoExcel_DLL_NAME_FLAVOR
       "${wxAutoExcel_COMPILER_ID_STR}_${wxAutoExcel_ARCH_ID}")

if(wxAutoExcel_BUILD_SHARED)
  # no "lib" at the start of the DLL library
  set_target_properties(wxAutoExcel PROPERTIES PREFIX "")
endif()

# we want to have import library with extension ".a", not ".dll.a"
if(NOT MSVC)
  set_target_properties(wxAutoExcel PROPERTIES IMPORT_SUFFIX ".a")
endif()

if(NOT "${wxAutoExcel_BUILD_VENDOR}" STREQUAL "")
  set(VENDOR_SUFFIX "_${wxAutoExcel_BUILD_VENDOR}")
else()
  set(VENDOR_SUFFIX "")
endif()

set_target_properties(wxAutoExcel PROPERTIES
  RUNTIME_OUTPUT_NAME "${wxAutoExcel_DLL_NAME_MAIN}_${wxAutoExcel_DLL_NAME_FLAVOR}${VENDOR_SUFFIX}"
  RUNTIME_OUTPUT_NAME_DEBUG "${wxAutoExcel_DLL_NAME_MAIN}d_${wxAutoExcel_DLL_NAME_FLAVOR}${VENDOR_SUFFIX}"
)

# for static and import libraries
set_target_properties(wxAutoExcel PROPERTIES
  ARCHIVE_OUTPUT_DIRECTORY ${wxAutoExcel_ARCHIVE_DIR}
  ARCHIVE_OUTPUT_DIRECTORY_DEBUG ${wxAutoExcel_ARCHIVE_DIR}
  ARCHIVE_OUTPUT_DIRECTORY_RELEASE ${wxAutoExcel_ARCHIVE_DIR}
)

# for DLLs
set_target_properties(wxAutoExcel PROPERTIES RUNTIME_OUTPUT_DIRECTORY "${wxAutoExcel_RUNTIME_DIR}")

# Make the library the startup project. If the samples are built,
# it will be changed to minimal sample.
set_target_properties(wxAutoExcel PROPERTIES PROJECT_LABEL "wxAutoExcel Library")
# projects for the samples are grouped under one folder
set_property(GLOBAL PROPERTY USE_FOLDERS ON)
set_directory_properties(PROPERTIES VS_STARTUP_PROJECT wxAutoExcel)


######################################################################
#
# wxAutoExcel packaging and installing
#
######################################################################

include(CMakePackageConfigHelpers)

set(wxAutoExcel_INSTALL_CMAKEDIR "${CMAKE_INSTALL_LIBDIR}/cmake/wxAutoExcel")

# Create package files in the build directory too, so find_package() can use
# wxAutoExcel without installing it first.
export(TARGETS wxAutoExcel
       NAMESPACE wxAutoExcel::
       FILE "${CMAKE_CURRENT_BINARY_DIR}/wxAutoExcelTargets.cmake")
configure_package_config_file(
  "${CMAKE_CURRENT_SOURCE_DIR}/build/CMake/wxAutoExcelConfig.cmake.in"
  "${CMAKE_CURRENT_BINARY_DIR}/wxAutoExcelConfig.cmake"
  INSTALL_DESTINATION "${wxAutoExcel_INSTALL_CMAKEDIR}"
)
configure_file(
  "${CMAKE_CURRENT_SOURCE_DIR}/build/CMake/wxAutoExcelDetectWxLinkage.cmake"
  "${CMAKE_CURRENT_BINARY_DIR}/wxAutoExcelDetectWxLinkage.cmake"
  COPYONLY
)
write_basic_package_version_file(
  "${CMAKE_CURRENT_BINARY_DIR}/wxAutoExcelConfigVersion.cmake"
  VERSION "${PROJECT_VERSION}"
  COMPATIBILITY SameMajorVersion
)

if(wxAutoExcel_BUILD_INSTALL)
  install(DIRECTORY "${CMAKE_CURRENT_SOURCE_DIR}/include/wx"
          DESTINATION "${CMAKE_INSTALL_INCLUDEDIR}")
  install(TARGETS wxAutoExcel EXPORT wxAutoExcelTargets)
  install(EXPORT wxAutoExcelTargets
          NAMESPACE wxAutoExcel::
          FILE wxAutoExcelTargets.cmake
          DESTINATION "${wxAutoExcel_INSTALL_CMAKEDIR}")
  install(FILES
          "${CMAKE_CURRENT_BINARY_DIR}/wxAutoExcelConfig.cmake"
          "${CMAKE_CURRENT_BINARY_DIR}/wxAutoExcelConfigVersion.cmake"
          "${CMAKE_CURRENT_BINARY_DIR}/wxAutoExcelDetectWxLinkage.cmake"
          DESTINATION "${wxAutoExcel_INSTALL_CMAKEDIR}")
  if(MSVC AND wxAutoExcel_BUILD_SHARED)
    install(FILES "$<TARGET_PDB_FILE:wxAutoExcel>"
            DESTINATION "${CMAKE_INSTALL_BINDIR}"
            CONFIGURATIONS Debug
            OPTIONAL)
  endif()
endif()

######################################################################
#
# wxAutoExcel samples
#
######################################################################

# function wxAutoExcel_add_sample has the following parameters:
# name = name of the sample, must match the name of the sample folder
# sourceFiles = .h and .cpp files for the sample, must be in the sample folder
# copyFiles = extra files a sample may need (must be in the sample folder), 
# will be copied to the binary dir after build.
function(wxAutoExcel_add_sample name sourceFiles copyFiles)

  set(src)
  foreach(file IN LISTS sourceFiles)
    list(APPEND src "${CMAKE_CURRENT_SOURCE_DIR}/samples/${name}/${file}")
  endforeach()

  list(APPEND src "${CMAKE_CURRENT_SOURCE_DIR}/samples/samples.rc")

  add_executable(${name} WIN32 ${src})
  set_target_properties(${name} PROPERTIES FOLDER "wxAutoExcel Samples")
  set_target_properties(${name} PROPERTIES RUNTIME_OUTPUT_DIRECTORY ${wxAutoExcel_RUNTIME_DIR})
  target_link_libraries(${name} PRIVATE wxAutoExcel::wxAutoExcel)

  # we are using the application manifest from the resource file
  if(MSVC)
    target_link_options(${name} PRIVATE "/MANIFEST:NO")
  endif()
  wxAutoExcel_adjust_target_flags(${name})
  
  if(wxAutoExcel_BUILD_USE_PRECOMPILED)
    if(name STREQUAL "minimal")
      target_precompile_headers(${name} PRIVATE "${CMAKE_CURRENT_SOURCE_DIR}/include/wx/wxAutoExcel_prec.h")
    else()
      target_precompile_headers(${name} REUSE_FROM "minimal")
    endif()
  endif()
 
  foreach(copyFile IN LISTS copyFiles)
    add_custom_command(TARGET ${name} POST_BUILD
      COMMAND ${CMAKE_COMMAND} -E copy_if_different
              "${CMAKE_CURRENT_SOURCE_DIR}/samples/${name}/${copyFile}"
              "$<TARGET_FILE_DIR:${name}>"
      VERBATIM)
  endforeach()

endfunction()

if(wxAutoExcel_BUILD_SAMPLES)
  
  wxAutoExcel_add_sample(minimal "minimal.cpp" "")
  if(MSVC)
    set_directory_properties(PROPERTIES VS_STARTUP_PROJECT minimal)
  endif()
  wxAutoExcel_add_sample(charts "charts.cpp" "")
  wxAutoExcel_add_sample(shapes "shapes.cpp" "")
  wxAutoExcel_add_sample(print "print.cpp" "wxAutoExcel small.bmp")
  wxAutoExcel_add_sample(window "window.cpp" "")
  wxAutoExcel_add_sample(bulkdata "bulkdata.cpp" "")
  wxAutoExcel_add_sample(table "table.cpp" "")
  wxAutoExcel_add_sample(XLSpy "enum2string.cpp;enum2string.h;getdata.cpp;getdata.h;xlspy.cpp;xlspy.h" "sample.xlsx")
  wxAutoExcel_add_sample(purewin32 "usewxAutoExcel.h;usewxAutoExcel.cpp;purewin32.cpp" "")
endif()


######################################################################
#
# wxAutoExcel tests
#
######################################################################

if(wxAutoExcel_BUILD_TESTS)
  enable_testing()
  add_subdirectory(tests)
endif()


######################################################################
#
# Configuration report
#
######################################################################

message(STATUS "Configured wxAutoExcel ${wxAutoExcel_VERSION} for ${CMAKE_GENERATOR} (${CMAKE_CXX_COMPILER}, v${CMAKE_CXX_COMPILER_VERSION})")
message(STATUS "  Build type:                      ${wxAutoExcel_BUILD_LIB_TYPE}")
if(CMAKE_CONFIGURATION_TYPES)
  message(STATUS "  Build configuration:             ${CMAKE_CONFIGURATION_TYPES}")
else()
  message(STATUS "  Build configuration:             ${CMAKE_BUILD_TYPE}")
endif()

message(STATUS "  Use static runtime:              ${wxAutoExcel_BUILD_USE_STATIC_RUNTIME}")

message(STATUS "  Architecture:                    ${wxAutoExcel_ARCH_ID}")
message(STATUS "  Compiler id string:              ${wxAutoExcel_COMPILER_ID_STR}")

message(STATUS "  Include folder:                  ${CMAKE_CURRENT_SOURCE_DIR}/include")
message(STATUS "  Static/Import library folder:    ${wxAutoExcel_ARCHIVE_DIR}")
message(STATUS "  Runtime folder:                  ${wxAutoExcel_RUNTIME_DIR}")

message(STATUS "  Use precompiled headers:         ${wxAutoExcel_BUILD_USE_PRECOMPILED}")
message(STATUS "  Create install target:           ${wxAutoExcel_BUILD_INSTALL}")
message(STATUS "  Build samples:                   ${wxAutoExcel_BUILD_SAMPLES}")
message(STATUS "  Build tests:                     ${wxAutoExcel_BUILD_TESTS}")
message(STATUS "  ------")
if(wxWidgets_VERSION)
  message(STATUS "  wxWidgets version:               ${wxWidgets_VERSION}")
else()
  message(STATUS "  wxWidgets version:               ${wxWidgets_VERSION_STRING}")
endif()
message(STATUS "  wxWidgets link shared:           ${wxAutoExcel_WX_SHARED}")
message(STATUS "  wxWidgets library folder:        ${wxWidgets_LIB_DIR}")

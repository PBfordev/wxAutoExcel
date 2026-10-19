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
find_package(wxWidgets COMPONENTS core base REQUIRED)
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

  target_compile_definitions(${target} PRIVATE wxNO_UNSAFE_WXSTRING_CONV)

  if(wxAutoExcel_BUILD_USE_STATIC_RUNTIME)
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
    target_compile_options(${target} PRIVATE /W4)
    # only MSBuild compiles the files of a target in parallel itself,
    # and clang-cl does not support /MP
    if(CMAKE_GENERATOR MATCHES "^Visual Studio" AND NOT CMAKE_CXX_COMPILER_ID STREQUAL "Clang")
      target_compile_options(${target} PRIVATE /MP)
    endif()
  else() # GCC or clang
    target_compile_options(${target} PRIVATE -Wall -Wextra -Wno-deprecated-declarations)
  endif()

endfunction()

if(wxAutoExcel_BUILD_SHARED)
  set(wxAutoExcel_BUILD_LIB_TYPE SHARED)
else()
  set(wxAutoExcel_BUILD_LIB_TYPE STATIC)
endif()

if(wxAutoExcel_BUILD_USE_STATIC_RUNTIME
   AND (wxAutoExcel_BUILD_SHARED OR wxAutoExcel_WX_SHARED))
     message(FATAL_ERROR "The static runtime cannot be used when creating wxAutoExcel DLLs or linking with wxWidgets dynamically")
endif()

# With MSVC, the runtime library of the static wxWidgets libraries must match
# the one used for wxAutoExcel, otherwise linking fails with error LNK2038.
# The runtime used by a library is recorded in it as a /DEFAULTLIB directive.
if(MSVC AND NOT wxAutoExcel_WX_SHARED AND wxWidgets_LIBRARIES)
  # check the release build of the wxBase library
  set(wxAutoExcel_WX_BASE_LIB)
  foreach(wxLib IN LISTS wxWidgets_LIBRARIES)
    get_filename_component(wxLibName "${wxLib}" NAME)
    if(wxLibName MATCHES "^wxbase[0-9]+u\\.lib$" AND EXISTS "${wxLib}")
      set(wxAutoExcel_WX_BASE_LIB "${wxLib}")
      break()
    endif()
  endforeach()

  if(wxAutoExcel_WX_BASE_LIB)
    file(STRINGS "${wxAutoExcel_WX_BASE_LIB}" wxAutoExcel_WX_RUNTIME_DIRECTIVE
         REGEX "/DEFAULTLIB:\"(LIBCMT|MSVCRT)\"" LIMIT_COUNT 1)
    set(wxAutoExcel_WX_STATIC_RUNTIME)
    if(wxAutoExcel_WX_RUNTIME_DIRECTIVE MATCHES "/DEFAULTLIB:\"LIBCMT\"")
      set(wxAutoExcel_WX_STATIC_RUNTIME ON)
    elseif(wxAutoExcel_WX_RUNTIME_DIRECTIVE MATCHES "/DEFAULTLIB:\"MSVCRT\"")
      set(wxAutoExcel_WX_STATIC_RUNTIME OFF)
    endif()

    if(DEFINED wxAutoExcel_WX_STATIC_RUNTIME)
      if(wxAutoExcel_WX_STATIC_RUNTIME AND NOT wxAutoExcel_BUILD_USE_STATIC_RUNTIME)
        message(FATAL_ERROR "wxWidgets libraries in ${wxWidgets_LIB_DIR} use the static runtime, set wxAutoExcel_BUILD_USE_STATIC_RUNTIME to ON")
      elseif(NOT wxAutoExcel_WX_STATIC_RUNTIME AND wxAutoExcel_BUILD_USE_STATIC_RUNTIME)
        message(FATAL_ERROR "wxWidgets libraries in ${wxWidgets_LIB_DIR} use the DLL runtime, set wxAutoExcel_BUILD_USE_STATIC_RUNTIME to OFF")
      endif()
    endif()
  endif()
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
  # CMAKE_SYSTEM_PROCESSOR is the host processor unless cross-compiling,
  # so it is used only to tell x86 from ARM, the pointer size tells the rest
  set(wxAutoExcel_ARCH_SOURCE "${CMAKE_SYSTEM_PROCESSOR}")
  if(wxAutoExcel_ARCH_SOURCE MATCHES "^([Aa][Rr][Mm]|[Aa][Aa]rch)")
    if(CMAKE_SIZEOF_VOID_P EQUAL 8)
      set(wxAutoExcel_ARCH_SOURCE "arm64")
    else()
      set(wxAutoExcel_ARCH_SOURCE "arm")
    endif()
  elseif(CMAKE_SIZEOF_VOID_P EQUAL 8)
    set(wxAutoExcel_ARCH_SOURCE "x64")
  elseif(CMAKE_SIZEOF_VOID_P EQUAL 4)
    set(wxAutoExcel_ARCH_SOURCE "x86")
  endif()
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
  set(wxAutoExcel_ARCH_ID "x86")
else()
  message(FATAL_ERROR
          "Unsupported or unknown target architecture: ${wxAutoExcel_ARCH_SOURCE}")
endif()

# like wxWidgets, the names of x86 DLLs have no architecture suffix
string(APPEND wxAutoExcel_DLL_NAME_FLAVOR "${wxAutoExcel_COMPILER_ID_STR}")
if(NOT wxAutoExcel_ARCH_ID STREQUAL "x86")
  string(APPEND wxAutoExcel_DLL_NAME_FLAVOR "_${wxAutoExcel_ARCH_ID}")
endif()

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

# for static and import libraries; like for DLLs, multi-config generators
# append a subfolder for each configuration, as Release, RelWithDebInfo and
# MinSizeRel libraries have the same name
set_target_properties(wxAutoExcel PROPERTIES ARCHIVE_OUTPUT_DIRECTORY "${wxAutoExcel_ARCHIVE_DIR}")

# for DLLs
set_target_properties(wxAutoExcel PROPERTIES RUNTIME_OUTPUT_DIRECTORY "${wxAutoExcel_RUNTIME_DIR}")

# Make the library the startup project. If the samples are built,
# it will be changed to minimal sample.
set_target_properties(wxAutoExcel PROPERTIES PROJECT_LABEL "wxAutoExcel Library")
# projects for the samples are grouped under one folder; the global
# property is not changed for a project including wxAutoExcel
if(PROJECT_IS_TOP_LEVEL)
  set_property(GLOBAL PROPERTY USE_FOLDERS ON)
endif()
set_directory_properties(PROPERTIES VS_STARTUP_PROJECT wxAutoExcel)

# The debug information of a static library compiled with MSVC is by default
# in a PDB file of the compiler, which is not installed and is overwritten by
# other targets. Store the debug information in the library itself instead.
if(MSVC AND NOT wxAutoExcel_BUILD_SHARED AND NOT DEFINED CMAKE_MSVC_DEBUG_INFORMATION_FORMAT)
  set_target_properties(wxAutoExcel PROPERTIES
    MSVC_DEBUG_INFORMATION_FORMAT "$<$<CONFIG:Debug,RelWithDebInfo>:Embedded>")
endif()


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
# minor versions are not compatible: they have different library names
# and may contain changes breaking the build of existing code
write_basic_package_version_file(
  "${CMAKE_CURRENT_BINARY_DIR}/wxAutoExcelConfigVersion.cmake"
  VERSION "${PROJECT_VERSION}"
  COMPATIBILITY SameMinorVersion
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
  # the linker creates a PDB file for each configuration with debug information
  if(MSVC AND wxAutoExcel_BUILD_SHARED)
    install(FILES "$<TARGET_PDB_FILE:wxAutoExcel>"
            DESTINATION "${CMAKE_INSTALL_BINDIR}"
            OPTIONAL)
  endif()
endif()

######################################################################
#
# wxWidgets DLLs for the samples and tests
#
######################################################################

# Target wxAutoExcel_wx_dlls copies the wxWidgets DLLs to the runtime folder,
# so that the samples and tests can be run without wxWidgets in the PATH.
# It is created only when linking with wxWidgets dynamically.
function(wxAutoExcel_add_wx_dlls_target)
  if(NOT wxAutoExcel_WX_SHARED)
    return()
  endif()

  set(releaseDlls)
  set(debugDlls)

  if(TARGET wxWidgets::core AND TARGET wxWidgets::base)
    set(releaseDlls "$<TARGET_FILE:wxWidgets::core>" "$<TARGET_FILE:wxWidgets::base>")
    set(debugDlls ${releaseDlls})
  elseif(TARGET wx::core AND TARGET wx::base)
    set(releaseDlls "$<TARGET_FILE:wx::core>" "$<TARGET_FILE:wx::base>")
    set(debugDlls ${releaseDlls})
  elseif(NOT wxWidgets_LIB_DIR OR NOT IS_DIRECTORY "${wxWidgets_LIB_DIR}")
    # wxWidgets found with wx-config, e.g. installed with MSYS2, has its
    # DLLs in the PATH already
    return()
  else()
    # The folder may contain the DLLs of several wxWidgets versions, so
    # select those of the version found. The DLL names contain the full
    # version, e.g. wxbase3211u_*.dll and wxbase335u_*.dll, except for
    # the stable (even minor) versions built with the makefiles for GCC,
    # whose names contain just the major and minor version, e.g. wxbase32u.
    set(dllVersion "[0-9]+")
    if(wxWidgets_VERSION_MAJOR MATCHES "^[0-9]+$"
       AND wxWidgets_VERSION_MINOR MATCHES "^[0-9]+$"
       AND wxWidgets_VERSION_PATCH MATCHES "^[0-9]+$")
      set(dllVersion "${wxWidgets_VERSION_MAJOR}${wxWidgets_VERSION_MINOR}${wxWidgets_VERSION_PATCH}")
      math(EXPR isDevelopment "${wxWidgets_VERSION_MINOR} % 2")
      if(NOT isDevelopment)
        set(dllVersion "(${dllVersion}|${wxWidgets_VERSION_MAJOR}${wxWidgets_VERSION_MINOR})")
      endif()
    else()
      message(WARNING
        "Unknown wxWidgets version, the DLLs of all wxWidgets versions "
        "in ${wxWidgets_LIB_DIR} will be copied to the runtime folder")
    endif()

    file(GLOB dllCandidates CONFIGURE_DEPENDS
      "${wxWidgets_LIB_DIR}/wxbase*.dll"
      "${wxWidgets_LIB_DIR}/wxmsw*u_core_*.dll"
      "${wxWidgets_LIB_DIR}/wxmsw*ud_core_*.dll"
    )
    foreach(dll IN LISTS dllCandidates)
      get_filename_component(dllName "${dll}" NAME)
      if(dllName MATCHES "^wxbase${dllVersion}ud?_(net|xml)_")
        continue()
      elseif(dllName MATCHES "^wx(base|msw)${dllVersion}ud(_core)?_.*\\.dll$")
        list(APPEND debugDlls "${dll}")
      elseif(dllName MATCHES "^wx(base|msw)${dllVersion}u(_core)?_.*\\.dll$")
        list(APPEND releaseDlls "${dll}")
      endif()
    endforeach()
  endif()

  if(NOT releaseDlls AND NOT debugDlls)
    message(WARNING
      "Could not find the wxWidgets core and base DLLs in ${wxWidgets_LIB_DIR}; "
      "the samples and tests may not be runnable")
    return()
  endif()

  if(CMAKE_CONFIGURATION_TYPES)
    set(runtimeDir "${wxAutoExcel_RUNTIME_DIR}/$<CONFIG>")
  else()
    set(runtimeDir "${wxAutoExcel_RUNTIME_DIR}")
  endif()

  # wxWidgets may provide the DLLs only for the debug or the release builds,
  # copy nothing for the configurations without them
  foreach(dllType debug release)
    if(${dllType}Dlls)
      set(${dllType}Copy ${CMAKE_COMMAND} -E copy_if_different ${${dllType}Dlls} "${runtimeDir}")
    else()
      message(STATUS "wxWidgets ${dllType} DLLs not found in ${wxWidgets_LIB_DIR}, they will not be copied")
      set(${dllType}Copy ${CMAKE_COMMAND} -E true)
    endif()
  endforeach()

  add_custom_target(wxAutoExcel_wx_dlls
    COMMAND ${CMAKE_COMMAND} -E make_directory "${runtimeDir}"
    COMMAND "$<IF:$<CONFIG:Debug>,${debugCopy},${releaseCopy}>"
    COMMAND_EXPAND_LISTS
    VERBATIM
  )
  set_target_properties(wxAutoExcel_wx_dlls PROPERTIES FOLDER "wxAutoExcel Samples")
endfunction()

if(wxAutoExcel_BUILD_SAMPLES OR wxAutoExcel_BUILD_TESTS)
  wxAutoExcel_add_wx_dlls_target()
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
  if(TARGET wxAutoExcel_wx_dlls)
    add_dependencies(${name} wxAutoExcel_wx_dlls)
  endif()

  # we are using the application manifest from the resource file
  if(MSVC)
    target_link_options(${name} PRIVATE "/MANIFEST:NO")
  endif()
  wxAutoExcel_adjust_target_flags(${name})
  
  # each sample has its own precompiled header: reusing one with REUSE_FROM
  # makes MSVC share its compiler PDB among the samples, which breaks linking
  # a sample whose files were not recompiled when the shared PDB was rebuilt
  if(wxAutoExcel_BUILD_USE_PRECOMPILED)
    target_precompile_headers(${name} PRIVATE "${CMAKE_CURRENT_SOURCE_DIR}/include/wx/wxAutoExcel_prec.h")
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
  # the Win32 part of the sample does not use wxWidgets or wxAutoExcel
  set_source_files_properties("${CMAKE_CURRENT_SOURCE_DIR}/samples/purewin32/purewin32.cpp"
                              PROPERTIES SKIP_PRECOMPILE_HEADERS ON)
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

######################################################################
# Author:      PB
# Purpose:     Primary CMake for wxAutoExcel
# Copyright:   (c) 2017 PB <pbfordev@gmail.com>
# Licence:     wxWindows licence
######################################################################

cmake_minimum_required(VERSION 3.16 FATAL_ERROR)

# parses the wxAutoExcel version from include/wx/wxAutoExcel_version.h
# into ${wxAutoExcel_MAJOR_VER}, ${wxAutoExcel_MINOR_VER}, and ${wxAutoExcel_REL_NUM}
include("${CMAKE_CURRENT_SOURCE_DIR}/build/CMake/version.cmake")

project(wxAutoExcel
        VERSION ${wxAutoExcel_MAJOR_VER}.${wxAutoExcel_MINOR_VER}.${wxAutoExcel_REL_NUM}
        DESCRIPTION "wxWidgets library making automating Microsoft Excel easier"
        HOMEPAGE_URL https://github.com/PBfordev/wxAutoExcel
        LANGUAGES CXX
)

if(NOT CMAKE_SYSTEM_NAME STREQUAL Windows)
  message(FATAL_ERROR "wxAutoExcel is available only for Microsoft Windows")
endif()

# wxAutoExcel requires wxWidgets 3.1 and higher
find_package(wxWidgets 3.1 COMPONENTS core base REQUIRED)
if(wxWidgets_USE_FILE)
  include(${wxWidgets_USE_FILE})
endif()

######################################################################
#
# wxAutoExcel library and samples can be built in Debug and/or Release configurations
#
######################################################################

get_property(isMultiConfig GLOBAL PROPERTY GENERATOR_IS_MULTI_CONFIG)
if(isMultiConfig)
  set(CMAKE_CONFIGURATION_TYPES Debug Release)
  set(CMAKE_CONFIGURATION_TYPES "${CMAKE_CONFIGURATION_TYPES}" CACHE STRING "Allowed configuration types" FORCE)
else()
  set(allowedBuildTypes Debug Release)
  set_property(CACHE CMAKE_BUILD_TYPE PROPERTY STRINGS "${allowedBuildTypes}")
  if(NOT CMAKE_BUILD_TYPE)
    set(CMAKE_BUILD_TYPE Debug CACHE STRING "" FORCE)
  elseif(NOT CMAKE_BUILD_TYPE IN_LIST allowedBuildTypes)
    message(FATAL_ERROR "Unknown build type: ${CMAKE_BUILD_TYPE}")
  endif()
endif()


######################################################################
#
# Global user-customizable options
#
######################################################################

option(wxAutoExcel_BUILD_SHARED "Build wxAutoExcel libraries as shared libs" ${BUILD_SHARED_LIBS})
option(wxAutoExcel_BUILD_LINK_WX_SHARED "Link with wxWidgets dynamically" ${BUILD_SHARED_LIBS})
option(wxAutoExcel_BUILD_USE_STATIC_RUNTIME "Link using the static runtime library" OFF)
option(wxAutoExcel_BUILD_USE_PRECOMPILED "Use precompiled headers" ON)
set(wxAutoExcel_BUILD_VENDOR "custom" CACHE STRING "Vendor")
option(wxAutoExcel_BUILD_INSTALL "Create install target for the library" OFF)

option(wxAutoExcel_BUILD_SAMPLES "Build samples" ${PROJECT_IS_TOP_LEVEL})


######################################################################
#
# Compiler and linker options
#
######################################################################

function(wxAutoExcel_adjust_target_flags target)

  target_compile_definitions(${target} PRIVATE __WXMSW__ wxNO_UNSAFE_WXSTRING_CONV $<$<CONFIG:Release>:NDEBUG>)
  
  if(wxAutoExcel_BUILD_LINK_WX_SHARED)
    target_compile_definitions(${target} PRIVATE WXUSINGDLL)
  endif()

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
  set(LIB_SUFFIX "dll")
else()
  set(wxAutoExcel_BUILD_LIB_TYPE STATIC)
  set(LIB_SUFFIX "lib")
endif()

if(wxAutoExcel_BUILD_LINK_WX_SHARED)
  if(NOT "${wxWidgets_LIB_DIR}" STREQUAL "wxWidgets_LIB_DIR-NOTFOUND" AND NOT "${wxWidgets_LIB_DIR}" MATCHES "_dll")
    MESSAGE(WARNING "When linking with wxWidgets dynamically you should probably replace \"_lib\" with \"_dll\" in the wxWidgets\\wxWidgets_LIB_DIR variable")
  endif()
endif()

if(DEFINED wxAutoExcel_BUILD_USE_STATIC_RUNTIME AND wxAutoExcel_BUILD_USE_STATIC_RUNTIME
   AND (wxAutoExcel_BUILD_SHARED OR wxAutoExcel_BUILD_LINK_WX_SHARED))
     MESSAGE(FATAL_ERROR "The static runtime cannot be used when creating wxAutoExcel DLLs or linking with wxWidgets dynamically")
endif()

# Initialize variables for quick access to wxAutoExcel dirs
set(wxAutoExcel_RUNTIME_DIR "${CMAKE_BINARY_DIR}/bin")
set(wxAutoExcel_ARCHIVE_DIR "${CMAKE_BINARY_DIR}/lib")


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
  $<INSTALL_INTERFACE:include>
)
if(wxAutoExcel_BUILD_SHARED)
  target_link_libraries(wxAutoExcel PUBLIC ${wxWidgets_LIBRARIES})
endif()

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

if(MSVC)
  set(wxAutoExcel_COMPILER_ID_STR "vc${MSVC_TOOLSET_VERSION}")
else()
  # remove dots from compiler version, i.e. change e.g. "11.3.0" to "1130"
  string(REPLACE "." "" wxAutoExcel_COMPILER_ID_STR "gcc${CMAKE_CXX_COMPILER_VERSION}")
  if("${CMAKE_CXX_COMPILER_ID}" STREQUAL "GNU")
    string(REPLACE "." "" wxAutoExcel_COMPILER_ID_STR "gcc${CMAKE_CXX_COMPILER_VERSION}")
  elseif("${CMAKE_CXX_COMPILER_ID}" STREQUAL "Clang")
    string(REPLACE "." "" wxAutoExcel_COMPILER_ID_STR "clang${CMAKE_CXX_COMPILER_VERSION}")
  else()
    message(FATAL_ERROR "Unsupported compiler type: " "${CMAKE_CXX_COMPILER_ID}")
  endif()
endif()

if("${CMAKE_SIZEOF_VOID_P}" EQUAL "8") # 64-bit target
  string(APPEND wxAutoExcel_DLL_NAME_FLAVOR "${wxAutoExcel_COMPILER_ID_STR}_x64")
else()
  string(APPEND wxAutoExcel_DLL_NAME_FLAVOR "${wxAutoExcel_COMPILER_ID_STR}_x32")
endif()

if(wxAutoExcel_BUILD_SHARED)
  # no "lib" at the start of the DLL library
  set_target_properties(wxAutoExcel PROPERTIES PREFIX "")
endif()

# we want to have import library with extension ".a", not ".dll.a"
if(NOT MSVC)
  set_target_properties(wxAutoExcel PROPERTIES IMPORT_SUFFIX ".a")
endif()

if (NOT "${wxAutoExcel_BUILD_VENDOR}" STREQUAL "")
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
set_directory_properties(PROPERTIES VS_STARTUP_PROJECT "wxAutoExcel Library")


######################################################################
#
# wxAutoExcel packaging and installing
#
######################################################################

include(CMakePackageConfigHelpers)

# create configs also in the build dir, so find_package() can use that without installing
export(TARGETS wxAutoExcel NAMESPACE wxAutoExcel:: FILE wxAutoExcelConfig.cmake)
write_basic_package_version_file("${CMAKE_CURRENT_BINARY_DIR}/wxAutoExcelConfigVersion.cmake" COMPATIBILITY SameMajorVersion)

if(wxAutoExcel_BUILD_INSTALL)
  include(GNUInstallDirs)
  install(DIRECTORY ${CMAKE_CURRENT_SOURCE_DIR}/include/wx DESTINATION ${CMAKE_INSTALL_INCLUDEDIR})
  install(TARGETS wxAutoExcel EXPORT wxAutoExcelTargets) 
  install(EXPORT wxAutoExcelTargets NAMESPACE wxAutoExcel:: FILE wxAutoExcelConfig.cmake DESTINATION "${CMAKE_INSTALL_LIBDIR}/cmake/wxAutoExcel")
  install(FILES "${CMAKE_CURRENT_BINARY_DIR}/wxAutoExcelConfigVersion.cmake" DESTINATION "${CMAKE_INSTALL_LIBDIR}/cmake/wxAutoExcel")
  if(MSVC AND wxAutoExcel_BUILD_SHARED)
    install(FILES $<$<CONFIG:Debug>:$<TARGET_PDB_FILE:${PROJECT_NAME}>> DESTINATION ${CMAKE_INSTALL_BINDIR})
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

  foreach(file ${sourceFiles})
    list(APPEND src "${CMAKE_CURRENT_SOURCE_DIR}/samples/${name}/${file}")
  endforeach(file)

  list(APPEND src "${CMAKE_CURRENT_SOURCE_DIR}/samples/samples.rc")

  add_executable(${name} WIN32 ${src})
  set_target_properties(${name} PROPERTIES FOLDER "wxAutoExcel Samples")
  set_target_properties(${name} PROPERTIES RUNTIME_OUTPUT_DIRECTORY ${wxAutoExcel_RUNTIME_DIR})
  target_link_libraries(${name} wxAutoExcel::wxAutoExcel ${wxWidgets_LIBRARIES})

  # we are using the application manifest from the resource file
  if(MSVC)
    target_link_options(${name} PRIVATE "/MANIFEST:NO")
  endif()
  wxAutoExcel_adjust_target_flags(${name})
  
  if(wxAutoExcel_BUILD_USE_PRECOMPILED)
    if (${name} STREQUAL "minimal")
      target_precompile_headers(${name} PRIVATE "${CMAKE_CURRENT_SOURCE_DIR}/include/wx/wxAutoExcel_prec.h")
    else()
      target_precompile_headers(${name} REUSE_FROM "minimal")
    endif()
  endif()
 
  foreach(copyFile ${copyFiles})
    add_custom_command(TARGET ${name} POST_BUILD
      COMMAND ${CMAKE_COMMAND} -E copy "${CMAKE_CURRENT_SOURCE_DIR}/samples/${name}/${copyFile}" ${wxAutoExcel_RUNTIME_DIR})
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
  wxAutoExcel_add_sample(XLSpy "enum2string.cpp;enum2string.h;getdata.cpp;getdata.h;xlspy.cpp;xlspy.h" "sample.xlsx")
  wxAutoExcel_add_sample(purewin32 "usewxAutoExcel.h;usewxAutoExcel.cpp;purewin32.cpp" "")
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

if("${CMAKE_SIZEOF_VOID_P}" EQUAL "8") # 64-bit target
  message(STATUS "  Architecture:                    64-bit")
else()
  message(STATUS "  Architecture:                    32-bit")
endif()
  message(STATUS "  Compiler id string:              ${wxAutoExcel_COMPILER_ID_STR}")

  message(STATUS "  Include folder:                  ${CMAKE_CURRENT_SOURCE_DIR}/include")
  message(STATUS "  Static/Import library folder:    ${wxAutoExcel_ARCHIVE_DIR}")
  message(STATUS "  Runtime folder:                  ${wxAutoExcel_RUNTIME_DIR}")

message(STATUS "  Use precompiled headers:         ${wxAutoExcel_BUILD_USE_PRECOMPILED}")
message(STATUS "  Build samples:                   ${wxAutoExcel_BUILD_SAMPLES}")
message(STATUS "  ------")
message(STATUS "  wxWidgets version:               ${wxWidgets_VERSION_STRING}")
message(STATUS "  wxWidgets link shared:           ${wxAutoExcel_BUILD_LINK_WX_SHARED}")
message(STATUS "  wxWidgets library folder:        ${wxWidgets_LIB_DIR}")
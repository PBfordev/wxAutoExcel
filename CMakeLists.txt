######################################################################
# Author:      PB
# Purpose:     Primary CMake for wxAutoExcel
# Copyright:   (c) 2017 PB <pbfordev@gmail.com>
# Licence:     wxWindows licence
######################################################################

cmake_minimum_required(VERSION 3.1 FATAL_ERROR)
project(wxAutoExcel)

if ( POLICY CMP0045 )
  # issue warning when calling get_target_property() with a non-existent target
  cmake_policy(SET CMP0045 NEW)
endif()

# projects for the samples are grouped under one folder
set_property(GLOBAL PROPERTY USE_FOLDERS ON)

if ( NOT WIN32 )
  message(FATAL_ERROR "wxAutoExcel is available only for Microsoft Windows")
endif()

if ( NOT ("${CMAKE_CXX_COMPILER_ID}" STREQUAL "MSVC" OR "${CMAKE_CXX_COMPILER_ID}" STREQUAL "GNU") )
    message(FATAL_ERROR "Unsupported compiler type")
endif()

# wxAutoExcel requires wxWidgets 3.1 and higher
find_package(wxWidgets 3.1.0 COMPONENTS core base REQUIRED)

set (CMAKE_MODULE_PATH "${CMAKE_SOURCE_DIR}/build/CMake/Modules")
# for precompiled headers
include(cotire)

# list of .h and .cpp files for the wxAutoExcel library
include("${CMAKE_SOURCE_DIR}/build/CMake/files.cmake")
#parses the version number of wxAutoExcel from include/wx/wxAutoExcel_version.h
include("${CMAKE_SOURCE_DIR}/build/CMake/version.cmake")

# wxAutoExcel library and samples can be built in Debug and/or Release configurations
if ( CMAKE_CONFIGURATION_TYPES )
    set(CMAKE_CONFIGURATION_TYPES Debug Release)
      set(CMAKE_CONFIGURATION_TYPES "${CMAKE_CONFIGURATION_TYPES}" CACHE STRING
           "Reset the configurations to what we need" FORCE)
elseif (NOT CMAKE_BUILD_TYPE) # https://blog.kitware.com/cmake-and-the-default-build-type/
    set(default_build_type "Debug")
    message(STATUS "Setting build type to '${default_build_type}' as none was specified.")
    set(CMAKE_BUILD_TYPE "${default_build_type}" CACHE STRING "Choose the type of build." FORCE)
    # Set the possible values of build type for cmake-gui
    set_property(CACHE CMAKE_BUILD_TYPE PROPERTY STRINGS "Debug" "Release")
endif()

#################################
#
# Global user-customizable options
#
#################################

option(wxAutoExcel_BUILD_SHARED "Build wxAutoExcel libraries as shared libs" ${BUILD_SHARED_LIBS})
option(wxAutoExcel_BUILD_LINK_WX_SHARED "Link with wxWidgets dynamically" OFF)

if ( MSVC OR MINGW )
  option(wxAutoExcel_BUILD_USE_STATIC_RUNTIME "Link using the static runtime library" OFF)
endif()

option(wxAutoExcel_BUILD_USE_PRECOMPILED "Use precompiled headers" ON)
set(wxAutoExcel_BUILD_VENDOR "custom" CACHE STRING "Vendor")

option(wxAutoExcel_BUILD_BUILD_SAMPLES "Build samples" ON)

#################################
#
# Compiler and linker options
#
#################################

if ( wxAutoExcel_BUILD_SHARED )
  set(wxAutoExcel_BUILD_LIB_TYPE SHARED)
  set(SHARED 1)
  set(LIB_SUFFIX "dll")
else()
  set(wxAutoExcel_BUILD_LIB_TYPE STATIC)
  set(SHARED 0)
  set(LIB_SUFFIX "lib")
endif()

if ( wxAutoExcel_BUILD_LINK_WX_SHARED )
  add_definitions(-DWXUSINGDLL)
  if ( NOT ${wxWidgets_LIB_DIR} MATCHES "_dll" )
    MESSAGE(WARNING "When linking with wxWidgets dynamically you should probably replace \"_lib\" with \"_dll\" in the wxWidgets\\wxWidgets_LIB_DIR variable")
  endif()
endif()

if ( DEFINED wxAutoExcel_BUILD_USE_STATIC_RUNTIME AND wxAutoExcel_BUILD_USE_STATIC_RUNTIME )
    if  ( wxAutoExcel_BUILD_SHARED OR wxAutoExcel_BUILD_LINK_WX_SHARED )
        MESSAGE(FATAL_ERROR "The static runtime cannot be used when creating wxAutoExcel DLLs or linking with wxWidgets dynamically")
    endif()
    if ( MSVC )
        string(REGEX REPLACE "/MD" "/MT" CMAKE_CXX_FLAGS ${CMAKE_CXX_FLAGS})
        string(REGEX REPLACE "/MD" "/MT" CMAKE_CXX_FLAGS_DEBUG ${CMAKE_CXX_FLAGS_DEBUG})
        string(REGEX REPLACE "/MD" "/MT" CMAKE_CXX_FLAGS_RELEASE ${CMAKE_CXX_FLAGS_RELEASE})
    else() #MINGW
        set(CMAKE_SHARED_LINKER_FLAGS "${CMAKE_STATIC_LINKER_FLAGS} -static")
        set(CMAKE_EXE_LINKER_FLAGS "${CMAKE_EXE_LINKER_FLAGS} -static")
    endif()
endif()

if ( MSVC )
  # suppress security warnings
  add_definitions(-D_CRT_SECURE_NO_DEPRECATE=1)
  add_definitions(-D_CRT_NON_CONFORMING_SWPRINTFS=1)
  add_definitions(-D_SCL_SECURE_NO_WARNINGS=1)
  # set warnings to level 4
  string(REGEX REPLACE "/W3" "/W4" CMAKE_CXX_FLAGS ${CMAKE_CXX_FLAGS})
  string(REGEX REPLACE "/W3" "/W4" CMAKE_CXX_FLAGS_DEBUG ${CMAKE_CXX_FLAGS_DEBUG})
  string(REGEX REPLACE "/W3" "/W4" CMAKE_CXX_FLAGS_RELEASE ${CMAKE_CXX_FLAGS_RELEASE})

  set(VS_STARTUP_PROJECT wxAutoExcel)
  set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} /MP")
  set(CMAKE_C_FLAGS "${CMAKE_C_FLAGS} /MP")

else() #GCC
  set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -Wno-deprecated-declarations")
endif()

add_definitions(-D__WXMSW__)
add_definitions(-DwxNO_UNSAFE_WXSTRING_CONV)
include( ${wxWidgets_USE_FILE} )

# Initialize variables for quick access to wxAutoExcel root dir in sub dirs
set(wxAutoExcel_SOURCE_DIR ${CMAKE_CURRENT_SOURCE_DIR})
set(wxAutoExcel_RUNTIME_DIR ${CMAKE_BINARY_DIR}/bin)
set(wxAutoExcel_ARCHIVE_DIR ${CMAKE_BINARY_DIR}/lib)

include_directories("${wxAutoExcel_SOURCE_DIR}/include")

#################################
#
# wxAutoExcel library
#
#################################


if ( wxAutoExcel_BUILD_SHARED )
  list(APPEND SRCS "src/wxAutoExcel_version.rc")
endif()

add_library(wxAutoExcel ${wxAutoExcel_BUILD_LIB_TYPE} ${SRCS})
target_link_libraries(wxAutoExcel ${wxWidgets_LIBRARIES} version shlwapi oleacc uxtheme)
set_directory_properties(PROPERTIES VS_STARTUP_PROJECT wxAutoExcel)


set (wxAutoExcel_OUTPUT_NAME_MAIN "wxAutoExcel${wxAutoExcel_MAJOR_VERSION}${wxAutoExcel_MINOR_VERSION}")
set_target_properties(wxAutoExcel PROPERTIES
    OUTPUT_NAME ${wxAutoExcel_OUTPUT_NAME_MAIN}
    OUTPUT_NAME_DEBUG "${wxAutoExcel_OUTPUT_NAME_MAIN}d"
    )

set (wxAutoExcel_DLL_NAME_MAIN "wxAutoExcel${wxAutoExcel_MAJOR_VERSION}${wxAutoExcel_MINOR_VERSION}${wxAutoExcel_RELEASE_NUMBER}")

if ( MSVC )
  set(wxAutoExcel_COMPILER_ID_STR "vc${MSVC_TOOLSET_VERSION}")
else() # GCC
  # remove dots from compiler version, i.e. change e.g. "gcc8.1.0" to "gcc810"
  string(REPLACE "." "" wxAutoExcel_COMPILER_ID_STR "gcc${CMAKE_CXX_COMPILER_VERSION}")
endif()


if ( "${CMAKE_SIZEOF_VOID_P}" EQUAL "8" ) # 64-bit target
  string(APPEND wxAutoExcel_DLL_NAME_FLAVOR "${wxAutoExcel_COMPILER_ID_STR}_x64")
else()
  string(APPEND wxAutoExcel_DLL_NAME_FLAVOR "${wxAutoExcel_COMPILER_ID_STR}_x32")
endif()

if ( wxAutoExcel_BUILD_SHARED )
  # no "lib" at the start of the DLL library
  set_target_properties(wxAutoExcel PROPERTIES PREFIX "")
endif()

# we want to have import library with extension ".a", not ".dll.a"
if ( "${CMAKE_CXX_COMPILER_ID}" STREQUAL "GNU" )
  set_target_properties(wxAutoExcel PROPERTIES IMPORT_SUFFIX ".a")
endif()

set_target_properties(wxAutoExcel PROPERTIES
    RUNTIME_OUTPUT_NAME "${wxAutoExcel_DLL_NAME_MAIN}_${wxAutoExcel_DLL_NAME_FLAVOR}_${wxAutoExcel_BUILD_VENDOR}"
    RUNTIME_OUTPUT_NAME_DEBUG "${wxAutoExcel_DLL_NAME_MAIN}d_${wxAutoExcel_DLL_NAME_FLAVOR}_${wxAutoExcel_BUILD_VENDOR}"
    )


# for static and import libraries
set_target_properties(wxAutoExcel PROPERTIES
    ARCHIVE_OUTPUT_DIRECTORY ${wxAutoExcel_ARCHIVE_DIR}
    ARCHIVE_OUTPUT_DIRECTORY_DEBUG ${wxAutoExcel_ARCHIVE_DIR}
    ARCHIVE_OUTPUT_DIRECTORY_RELEASE ${wxAutoExcel_ARCHIVE_DIR}
    )

# for DLLs
set_target_properties(wxAutoExcel PROPERTIES
    RUNTIME_OUTPUT_DIRECTORY ${wxAutoExcel_RUNTIME_DIR}
    )

target_compile_definitions(wxAutoExcel PRIVATE $<$<CONFIG:Debug>:__WXDEBUG__ _DEBUG>)
target_compile_definitions(wxAutoExcel PRIVATE $<$<CONFIG:Release>:NDEBUG>)

if ( wxAutoExcel_BUILD_SHARED )
  target_compile_definitions(wxAutoExcel PRIVATE WXMAKINGDLL_WXAUTOEXCEL)
endif()

if ( wxAutoExcel_BUILD_USE_PRECOMPILED )
  set_target_properties(wxAutoExcel PROPERTIES COTIRE_CXX_PREFIX_HEADER_INIT "include/wx/wxAutoExcel_prec.h")
  cotire(wxAutoExcel)
endif()

#################################
#
# wxAutoExcel samples
#
#################################

# function wxAutoExcel_add_sample has the following parameters:
# name = name of the sample, must match the name of the sample folder
# sourceFiles = .h and .cpp files for the sample, must reside in the sample folder
# copyFiles = extra files (with absolute paths) a sample may need, will be copied to the binary dir after build
function(wxAutoExcel_add_sample name sourceFiles copyFiles)

  foreach(file ${sourceFiles})
    list(APPEND src "${wxAutoExcel_SOURCE_DIR}/samples/${name}/${file}")
  endforeach(file)

  list(APPEND src "${wxAutoExcel_SOURCE_DIR}/samples/samples.rc")

  add_executable(${name} WIN32 ${src} )
  set_target_properties(${name} PROPERTIES FOLDER "Samples")
  set_target_properties(${name} PROPERTIES RUNTIME_OUTPUT_DIRECTORY ${wxAutoExcel_RUNTIME_DIR})
  target_link_libraries(${name} wxAutoExcel)

  if ( wxAutoExcel_BUILD_SHARED )
    target_compile_definitions(${name} PRIVATE WXUSINGDLL_WXAUTOEXCEL)
  endif()
  target_compile_definitions(${name} PRIVATE $<$<CONFIG:Debug>:__WXDEBUG__>)

  foreach(file ${copyFiles})
    add_custom_command(TARGET print POST_BUILD
        COMMAND ${CMAKE_COMMAND} -E copy "${file}" "${wxAutoExcel_RUNTIME_DIR}")
  endforeach(file)
endfunction()

if ( wxAutoExcel_BUILD_BUILD_SAMPLES )
  if ( MSVC )
    set(CMAKE_EXE_LINKER_FLAGS "${CMAKE_EXE_LINKER_FLAGS} /MANIFEST:NO")
  endif()
  wxAutoExcel_add_sample(minimal "minimal.cpp" "")
  if ( MSVC )  
    set_directory_properties(PROPERTIES VS_STARTUP_PROJECT minimal)
  endif()
  wxAutoExcel_add_sample(charts "charts.cpp" "")
  wxAutoExcel_add_sample(shapes "shapes.cpp" "")
  wxAutoExcel_add_sample(print "print.cpp" "${wxAutoExcel_SOURCE_DIR}/samples/print/wxAutoExcel small.bmp")
  wxAutoExcel_add_sample(window "window.cpp" "")
  wxAutoExcel_add_sample(bulkdata "bulkdata.cpp" "")
  wxAutoExcel_add_sample(XLSpy "enum2string.cpp;enum2string.h;getdata.cpp;getdata.h;xlspy.cpp;xlspy.h" "${wxAutoExcel_SOURCE_DIR}/samples/XLSpy/sample.xlsx")
  wxAutoExcel_add_sample(purewin32 "usewxAutoExcel.h;usewxAutoExcel.cpp;purewin32.cpp" "")
endif()

#################################
#
# Configuration report
#
#################################

message(STATUS "Configured wxAutoExcel ${wxAutoExcel_VERSION_STRING} for ${CMAKE_GENERATOR} (${CMAKE_CXX_COMPILER}, v${CMAKE_CXX_COMPILER_VERSION})
    Build type:                      ${wxAutoExcel_BUILD_LIB_TYPE}")
if ( CMAKE_CONFIGURATION_TYPES )
  message(STATUS "    Build configuration:             ${CMAKE_CONFIGURATION_TYPES}")
else()
  message(STATUS "    Build configuration:             ${CMAKE_BUILD_TYPE}")
endif()
if ( MSVC OR MINGW )
  message(STATUS "    Use static runtime:              ${wxAutoExcel_BUILD_USE_STATIC_RUNTIME}")
endif()

if ( "${CMAKE_SIZEOF_VOID_P}" EQUAL "8" ) # 64-bit target
  message(STATUS "    Architecture:                    64-bit")
else()
  message(STATUS "    Architecture:                    32-bit")
endif()
  message(STATUS "    Compiler id string:              ${wxAutoExcel_COMPILER_ID_STR}")
message(STATUS ""
    "    Static/Import library folder:    ${wxAutoExcel_ARCHIVE_DIR}
    Runtime folder (DLLs + samples): ${wxAutoExcel_RUNTIME_DIR}
    Use precompiled headers:         ${wxAutoExcel_BUILD_USE_PRECOMPILED}
    Build samples:                   ${wxAutoExcel_BUILD_BUILD_SAMPLES}
    ------
    wxWidgets version:               ${wxWidgets_VERSION_STRING}
    wxWidgets link shared:           ${wxAutoExcel_BUILD_LINK_WX_SHARED}
    wxWidgets library folder:        ${wxWidgets_LIB_DIR}"
    )

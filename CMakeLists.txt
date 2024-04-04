######################################################################
######################################################################
# Author:      PB
# Purpose:     Primary CMake for wxAutoExcel
# Copyright:   (c) 2017 PB <pbfordev@gmail.com>
# Licence:     wxWindows licence
######################################################################

cmake_minimum_required(VERSION 3.14 FATAL_ERROR)

# parses the wxAutoExcel version from include/wx/wxAutoExcel_version.h
# into ${wxAutoExcel_MAJOR_VER}, ${wxAutoExcel_MINOR_VER}, and ${wxAutoExcel_REL_NUM}
include("${CMAKE_CURRENT_SOURCE_DIR}/build/CMake/version.cmake")

project(wxAutoExcel
        VERSION ${wxAutoExcel_MAJOR_VER}.${wxAutoExcel_MINOR_VER}.${wxAutoExcel_REL_NUM}
        DESCRIPTION "wxWidgets library making automating Microsoft Excel easier"
        HOMEPAGE_URL https://github.com/PBfordev/wxAutoExcel
        LANGUAGES CXX
)

if(NOT (CMAKE_SYSTEM_NAME STREQUAL Windows))
  message(FATAL_ERROR "wxAutoExcel is available only for Microsoft Windows")
endif()

# wxAutoExcel requires wxWidgets 3.1 and higher
find_package(wxWidgets 3.1 COMPONENTS core base REQUIRED)

# list of .h and .cpp files for the wxAutoExcel library
include("${wxAutoExcel_SOURCE_DIR}/build/CMake/files.cmake")

# wxAutoExcel library and samples can be built in Debug and/or Release configurations
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


#################################
#
# Global user-customizable options
#
#################################

option(wxAutoExcel_BUILD_SHARED "Build wxAutoExcel libraries as shared libs" ${BUILD_SHARED_LIBS})
option(wxAutoExcel_BUILD_LINK_WX_SHARED "Link with wxWidgets dynamically" ${BUILD_SHARED_LIBS})
option(wxAutoExcel_BUILD_USE_STATIC_RUNTIME "Link using the static runtime library" OFF)
option(wxAutoExcel_BUILD_USE_PRECOMPILED "Use precompiled headers (requires CMake v3.16+)" ON)
set(wxAutoExcel_BUILD_VENDOR "custom" CACHE STRING "Vendor")
option(wxAutoExcel_BUILD_INSTALL "Create install target for the library" OFF)

option(wxAutoExcel_BUILD_SAMPLES "Build samples" ${PROJECT_IS_TOP_LEVEL})

#################################
#
# Compiler and linker options
#
#################################

# projects for the samples are grouped under one folder
set_property(GLOBAL PROPERTY USE_FOLDERS ON)

if(wxAutoExcel_BUILD_SHARED)
  set(wxAutoExcel_BUILD_LIB_TYPE SHARED)
  set(LIB_SUFFIX "dll")
else()
  set(wxAutoExcel_BUILD_LIB_TYPE STATIC)
  set(LIB_SUFFIX "lib")
endif()

if(wxAutoExcel_BUILD_LINK_WX_SHARED)
  if(NOT "${wxWidgets_LIB_DIR}" MATCHES "_dll")
    MESSAGE(WARNING "When linking with wxWidgets dynamically you should probably replace \"_lib\" with \"_dll\" in the wxWidgets\\wxWidgets_LIB_DIR variable")
  endif()
endif()

if(DEFINED wxAutoExcel_BUILD_USE_STATIC_RUNTIME AND wxAutoExcel_BUILD_USE_STATIC_RUNTIME)
  if(wxAutoExcel_BUILD_SHARED OR wxAutoExcel_BUILD_LINK_WX_SHARED)
    MESSAGE(FATAL_ERROR "The static runtime cannot be used when creating wxAutoExcel DLLs or linking with wxWidgets dynamically")
  endif()
  if(MSVC)
    string(REGEX REPLACE "/MD" "/MT" CMAKE_CXX_FLAGS ${CMAKE_CXX_FLAGS})
    string(REGEX REPLACE "/MD" "/MT" CMAKE_CXX_FLAGS_DEBUG ${CMAKE_CXX_FLAGS_DEBUG})
    string(REGEX REPLACE "/MD" "/MT" CMAKE_CXX_FLAGS_RELEASE ${CMAKE_CXX_FLAGS_RELEASE})
  else()
    if(MINGW)
      set(CMAKE_STATIC_LINKER_FLAGS "${CMAKE_STATIC_LINKER_FLAGS} -static")
      set(CMAKE_EXE_LINKER_FLAGS "${CMAKE_EXE_LINKER_FLAGS} -static")
   else() # clang
      set(CMAKE_STATIC_LINKER_FLAGS "${CMAKE_STATIC_LINKER_FLAGS} -static-libgcc -static-libstdc++")
      set(CMAKE_EXE_LINKER_FLAGS "${CMAKE_EXE_LINKER_FLAGS} -static-libgcc -static-libstdc++")
   endif()
  endif()
endif()

if(MSVC)
  # suppress security warnings
  add_compile_definitions(_CRT_SECURE_NO_DEPRECATE _CRT_NON_CONFORMING_SWPRINTFS _SCL_SECURE_NO_WARNINGS)
  # set warnings to level 4
  string(REGEX REPLACE "/W3" "/W4" CMAKE_CXX_FLAGS ${CMAKE_CXX_FLAGS})

  # turn on multi-processor complication
  set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} /MP")
else() # GCC or clang
  set(CMAKE_CXX_FLAGS "${CMAKE_CXX_FLAGS} -Wno-deprecated-declarations")
endif()

include(${wxWidgets_USE_FILE})

# Initialize variables for quick access to wxAutoExcel dirs
set(wxAutoExcel_RUNTIME_DIR "${wxAutoExcel_BINARY_DIR}/bin")
set(wxAutoExcel_ARCHIVE_DIR "${wxAutoExcel_BINARY_DIR}/lib")

#################################
#
# wxAutoExcel library
#
#################################

if(wxAutoExcel_BUILD_SHARED)
  list(APPEND SRCS "${wxAutoExcel_SOURCE_DIR}/src/wxAutoExcel_version.rc")
endif()

add_library(wxAutoExcel ${wxAutoExcel_BUILD_LIB_TYPE} ${SRCS})
target_include_directories(wxAutoExcel PUBLIC
  $<BUILD_INTERFACE:${wxAutoExcel_SOURCE_DIR}/include>
  $<INSTALL_INTERFACE:include>
)
target_link_libraries(wxAutoExcel PUBLIC ${wxWidgets_LIBRARIES})

# Make the library the startup project in MSVS. If the samples are built,
# it will be changed to minimal sample.
if(MSVC)
  set_directory_properties(PROPERTIES VS_STARTUP_PROJECT wxAutoExcel)
endif()

set(CMAKE_DEBUG_POSTFIX d)

set(wxAutoExcel_OUTPUT_NAME_MAIN "wxAutoExcel${wxAutoExcel_VERSION_MAJOR}${wxAutoExcel_VERSION_MINOR}")
set_target_properties(wxAutoExcel PROPERTIES
  OUTPUT_NAME "${wxAutoExcel_OUTPUT_NAME_MAIN}"
  OUTPUT_NAME_DEBUG "${wxAutoExcel_OUTPUT_NAME_MAIN}${CMAKE_DEBUG_POSTFIX}"
)

target_compile_definitions(wxAutoExcel PUBLIC __WXMSW__ PRIVATE wxNO_UNSAFE_WXSTRING_CONV)
target_compile_definitions(wxAutoExcel PRIVATE $<$<CONFIG:Release>:NDEBUG>)

if(wxAutoExcel_BUILD_LINK_WX_SHARED)
  target_compile_definitions(wxAutoExcel PUBLIC WXUSINGDLL)
endif()

# WXUSINGDLL_WXAUTOEXCEL must be defined when building against shared wxAutoExcel build
if(wxAutoExcel_BUILD_SHARED)
  target_compile_definitions(wxAutoExcel INTERFACE WXUSINGDLL_WXAUTOEXCEL)
endif()

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
  RUNTIME_OUTPUT_NAME_DEBUG "${wxAutoExcel_DLL_NAME_MAIN}${CMAKE_DEBUG_POSTFIX}_${wxAutoExcel_DLL_NAME_FLAVOR}${VENDOR_SUFFIX}"
)

# for static and import libraries
set_target_properties(wxAutoExcel PROPERTIES
  ARCHIVE_OUTPUT_DIRECTORY ${wxAutoExcel_ARCHIVE_DIR}
  ARCHIVE_OUTPUT_DIRECTORY_DEBUG ${wxAutoExcel_ARCHIVE_DIR}
  ARCHIVE_OUTPUT_DIRECTORY_RELEASE ${wxAutoExcel_ARCHIVE_DIR}
)

# for DLLs
set_target_properties(wxAutoExcel PROPERTIES RUNTIME_OUTPUT_DIRECTORY "${wxAutoExcel_RUNTIME_DIR}")

if(wxAutoExcel_BUILD_SHARED)
  target_compile_definitions(wxAutoExcel PRIVATE WXMAKINGDLL_WXAUTOEXCEL)
endif()

if(wxAutoExcel_BUILD_USE_PRECOMPILED)
  if(${CMAKE_VERSION} VERSION_GREATER "3.15")
    target_precompile_headers(wxAutoExcel PRIVATE "${wxAutoExcel_SOURCE_DIR}/include/wx/wxAutoExcel_prec.h")    
  else()
    message(WARNING "Precompiled headers require CMake v3.16 and newer.")
  endif()  
endif()

if(wxAutoExcel_BUILD_INSTALL)
  install(DIRECTORY ${wxAutoExcel_SOURCE_DIR}/include/wx DESTINATION include)
  install(TARGETS wxAutoExcel EXPORT wxAutoExcelTargets)
  install(EXPORT wxAutoExcelTargets FILE wxAutoExcelConfig.cmake DESTINATION lib/cmake/wxAutoExcel)
  if(MSVC AND wxAutoExcel_BUILD_SHARED)
      install(FILES $<$<CONFIG:Debug>:$<TARGET_PDB_FILE:${PROJECT_NAME}>> DESTINATION bin)
  endif()
endif()
# create config also in the build dir, so find_package() can use that as well
export(TARGETS wxAutoExcel FILE wxAutoExcelConfig.cmake)

#################################
#
# wxAutoExcel samples
#
#################################

# function wxAutoExcel_add_sample has the following parameters:
# name = name of the sample, must match the name of the sample folder
# sourceFiles = .h and .cpp files for the sample, must be in the sample folder
# copyFiles = extra files a sample may need (must be in the sample folder), 
# will be copied to the binary dir after build.
function(wxAutoExcel_add_sample name sourceFiles copyFiles)

  foreach(file ${sourceFiles})
    list(APPEND src "${wxAutoExcel_SOURCE_DIR}/samples/${name}/${file}")
  endforeach(file)

  list(APPEND src "${wxAutoExcel_SOURCE_DIR}/samples/samples.rc")

  add_executable(${name} WIN32 ${src})
  target_compile_definitions(${name} PRIVATE $<$<CONFIG:Release>:NDEBUG>)
  set_target_properties(${name} PROPERTIES FOLDER "Samples")
  set_target_properties(${name} PROPERTIES RUNTIME_OUTPUT_DIRECTORY ${wxAutoExcel_RUNTIME_DIR})
  target_link_libraries(${name} wxAutoExcel)

  # we are using the application manifest from the resource file
  if(MSVC)
    target_link_options(${name} PRIVATE "/MANIFEST:NO")
  endif()
 
  foreach(copyFile ${copyFiles})
    add_custom_command(TARGET ${name} POST_BUILD
      COMMAND ${CMAKE_COMMAND} -E copy "${wxAutoExcel_SOURCE_DIR}/samples/${name}/${copyFile}" ${wxAutoExcel_RUNTIME_DIR})
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

#################################
#
# Configuration report
#
#################################
message(STATUS "Configured wxAutoExcel ${wxAutoExcel_VERSION} for ${CMAKE_GENERATOR} (${CMAKE_CXX_COMPILER}, v${CMAKE_CXX_COMPILER_VERSION})
  Build type:                      ${wxAutoExcel_BUILD_LIB_TYPE}")
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

  message(STATUS "  Include folder:                  ${wxAutoExcel_SOURCE_DIR}/include")
  message(STATUS "  Static/Import library folder:    ${wxAutoExcel_ARCHIVE_DIR}")
  message(STATUS "  Runtime folder:                  ${wxAutoExcel_RUNTIME_DIR}")

message(STATUS "  Use precompiled headers:         ${wxAutoExcel_BUILD_USE_PRECOMPILED}")
message(STATUS "  Build samples:                   ${wxAutoExcel_BUILD_SAMPLES}")
message(STATUS "  ------")
message(STATUS "  wxWidgets version:               ${wxWidgets_VERSION_STRING}")
message(STATUS "  wxWidgets link shared:           ${wxAutoExcel_BUILD_LINK_WX_SHARED}")
message(STATUS "  wxWidgets library folder:        ${wxWidgets_LIB_DIR}")
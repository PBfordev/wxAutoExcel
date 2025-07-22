vcpkg_from_github(
    OUT_SOURCE_PATH SOURCE_PATH
    REPO PBfordev/wxAutoExcel
    REF 6008e421f05c40a73c2ec48094d78ece6fdffb3c
    SHA512 1c236d26cd764193f6abb9de52a5245eee31dd9d52cea0162e0432064171fd7629e530e515ae9e3a12abe8f3ebb6e569ec3656c9ca05fcd7e51d6c8c154e3953
    HEAD_REF master
)

set(OPTIONS "")

if(VCPKG_CRT_LINKAGE STREQUAL "static")
    list(APPEND OPTIONS -DwxAutoExcel_BUILD_USE_STATIC_RUNTIME=ON)
endif()

vcpkg_cmake_configure(
    SOURCE_PATH "${SOURCE_PATH}"
    OPTIONS
        -DwxAutoExcel_BUILD_SAMPLES=OFF
        -DwxAutoExcel_BUILD_INSTALL=ON
)

vcpkg_cmake_install()
vcpkg_cmake_config_fixup(CONFIG_PATH lib/cmake/wxAutoExcel)
vcpkg_copy_pdbs()

file(REMOVE_RECURSE
        ${CURRENT_PACKAGES_DIR}/lib/cmake
        ${CURRENT_PACKAGES_DIR}/debug/lib/cmake
        ${CURRENT_PACKAGES_DIR}/debug/include
)

file(INSTALL "${CMAKE_CURRENT_LIST_DIR}/usage" DESTINATION "${CURRENT_PACKAGES_DIR}/share/${PORT}")
vcpkg_install_copyright(FILE_LIST "${SOURCE_PATH}/LICENSE.txt")
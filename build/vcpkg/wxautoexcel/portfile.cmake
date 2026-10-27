vcpkg_from_github(
    OUT_SOURCE_PATH SOURCE_PATH
    REPO PBfordev/wxAutoExcel
    REF f788392f377972d24d0ed6794bbe1357ef89199d
    SHA512 96c2914c9113cea7c82d9bde3876ed3bac93458aa7a5c0248d051e79aeb40ed15837fd493f4f7fd565bdb6317ae946c2e75055d522e4c3b9759c544061b0cc4a
    HEAD_REF master
)

set(OPTIONS "")

if(VCPKG_CRT_LINKAGE STREQUAL "static")
    list(APPEND OPTIONS -DwxAutoExcel_BUILD_USE_STATIC_RUNTIME=ON)
endif()

vcpkg_cmake_configure(
    SOURCE_PATH "${SOURCE_PATH}"
    OPTIONS
        ${OPTIONS}
        -DwxAutoExcel_BUILD_SAMPLES=OFF
        -DwxAutoExcel_BUILD_TESTS=OFF
        -DwxAutoExcel_BUILD_INSTALL=ON
)

vcpkg_cmake_install()
vcpkg_cmake_config_fixup(CONFIG_PATH lib/cmake/wxAutoExcel)
vcpkg_copy_pdbs()

file(REMOVE_RECURSE
        ${CURRENT_PACKAGES_DIR}/lib/cmake
        ${CURRENT_PACKAGES_DIR}/debug/lib/cmake
        ${CURRENT_PACKAGES_DIR}/debug/include
        ${CURRENT_PACKAGES_DIR}/debug/share
)

file(INSTALL "${CMAKE_CURRENT_LIST_DIR}/usage" DESTINATION "${CURRENT_PACKAGES_DIR}/share/${PORT}")
vcpkg_install_copyright(FILE_LIST "${SOURCE_PATH}/LICENSE.txt")
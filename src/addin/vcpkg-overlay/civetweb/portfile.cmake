vcpkg_check_linkage(ONLY_STATIC_LIBRARY)

# v1.16 is the latest tagged release; 1.17 is currently an unreleased branch.
# The SHA512 is the checksum used by the official vcpkg 1.16 port.
vcpkg_from_github(
    OUT_SOURCE_PATH SOURCE_PATH
    REPO civetweb/civetweb
    REF "v${VERSION}"
    SHA512 a0b943dfc76d7fd47f5a7d2c834fd38ddd4cf01a11730cf2f7cfaf32fea9698f59672f3a0f86ac80e0abc315d94d2367a500d37013f305c87d45e84cf39ca816
    PATCHES limit-websocket-frame.patch
)

# Use upstream switches rather than replacing its build or vendoring source.
# NO_FILESYSTEMS requires a custom logging implementation; keep the stock logger.
vcpkg_cmake_configure(
    SOURCE_PATH "${SOURCE_PATH}"
    OPTIONS
        -DCIVETWEB_VERSION=1.16.0
        -DCIVETWEB_BUILD_TESTING=OFF
        -DBUILD_TESTING=OFF
        -DCIVETWEB_ENABLE_SERVER_EXECUTABLE=OFF
        -DCIVETWEB_INSTALL_EXECUTABLE=OFF
        -DCIVETWEB_ENABLE_CXX=OFF
        -DCIVETWEB_ENABLE_SSL=OFF
        -DCIVETWEB_ENABLE_SSL_DYNAMIC_LOADING=OFF
        -DCIVETWEB_ENABLE_ZLIB=OFF
        -DCIVETWEB_ENABLE_LUA=OFF
        -DCIVETWEB_ENABLE_DUKTAPE=OFF
        -DCIVETWEB_ENABLE_WEBSOCKETS=ON
        -DCIVETWEB_ENABLE_IPV6=OFF
        -DCIVETWEB_SERVE_NO_FILES=ON
        -DCIVETWEB_DISABLE_CGI=ON
        -DCIVETWEB_DISABLE_CACHING=ON
        -DCIVETWEB_ENABLE_SERVER_STATS=OFF
        -DCIVETWEB_ENABLE_MEMORY_DEBUGGING=OFF
        -DCIVETWEB_ENABLE_DEBUG_TOOLS=OFF
        -DCIVETWEB_ENABLE_ASAN=OFF
        -DCIVETWEB_ALLOW_WARNINGS=ON
        -DCMAKE_INSTALL_SYSTEM_RUNTIME_LIBS_SKIP=ON
)

vcpkg_cmake_install()
vcpkg_copy_pdbs()
vcpkg_cmake_config_fixup(CONFIG_PATH lib/cmake/civetweb)

# Upstream unconditionally installs C++ pkgconfig metadata even without C++.
# Native consumers use the C header/library or the exported CMake target.
file(REMOVE_RECURSE
    "${CURRENT_PACKAGES_DIR}/debug/include"
    "${CURRENT_PACKAGES_DIR}/debug/share"
    "${CURRENT_PACKAGES_DIR}/share/pkgconfig"
)
vcpkg_install_copyright(FILE_LIST "${SOURCE_PATH}/LICENSE.md")

ICU_MAJOR_VER = 

core_windows {
    message($$PWD/$$CORE_BUILDS_PLATFORM_PREFIX/icu/include)
    INCLUDEPATH += $$PWD/$$CORE_BUILDS_PLATFORM_PREFIX/icu/include
}

core_win_64 {
    LIBS        += -L$$PWD/$$CORE_BUILDS_PLATFORM_PREFIX/icu/lib64 -licuuc
}
core_win_32 {
    message($$PWD/$$CORE_BUILDS_PLATFORM_PREFIX/icu/lib)
    LIBS        += -L$$PWD/$$CORE_BUILDS_PLATFORM_PREFIX/icu/lib -licuuc
}

core_linux {
#    INCLUDEPATH += $$PWD/$$CORE_BUILDS_PLATFORM_PREFIX/usr/local/include

#    LIBS        += $$PWD/$$CORE_BUILDS_PLATFORM_PREFIX/build/libicuuc.a
#    LIBS        += $$PWD/$$CORE_BUILDS_PLATFORM_PREFIX/build/libicudata.a
}

core_mac {
    INCLUDEPATH += $$PWD/$$CORE_BUILDS_PLATFORM_PREFIX/usr/local/include

    LIBS        += $$PWD/$$CORE_BUILDS_PLATFORM_PREFIX/build/libicuuc.$${ICU_MAJOR_VER}.dylib
    LIBS        += $$PWD/$$CORE_BUILDS_PLATFORM_PREFIX/build/libicudata.$${ICU_MAJOR_VER}.dylib
}

# icu wasm port flags
QMAKE_CFLAGS += -s USE_ICU=1
QMAKE_CXXFLAGS += -s USE_ICU=1
QMAKE_LFLAGS += -s USE_ICU=1

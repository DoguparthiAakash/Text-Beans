package com.example.notepadpp.media;

import java.io.File;

/**
 * Utility class to resolve native library files from an external directory.
 */
public class NativeUtils {

    // The base directory where your native DLLs reside.
    // This folder must be relative to your working directory.
    public static final String NATIVE_LIB_BASE_DIR = "lib/jcef";

    /**
     * Returns the File object for the given native library file name from the external folder.
     *
     * @param fileName The native library file name (e.g., "chrome_elf.dll").
     * @return The File representing the native library.
     * @throws IllegalArgumentException if the file does not exist.
     */
    public static File getNativeLibraryFile(String fileName) {
        File file = new File(NATIVE_LIB_BASE_DIR, fileName);
        if (!file.exists()) {
            throw new IllegalArgumentException("Native library not found: " + file.getAbsolutePath());
        }
        return file;
    }
}
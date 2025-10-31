package com.example.notepadpp.media;

import com.sun.jna.win32.StdCallLibrary;

/**
 * JNA interface for accessing the Windows Kernel32 API.
 * Although named "kernel32," it works on both 32‑bit and 64‑bit systems.
 */
public interface Kernel32 extends StdCallLibrary {
    /**
     * Adds a directory to the search path used to locate DLLs for the application.
     *
     * @param lpPathName the directory path to add.
     * @return true if the function succeeds; otherwise, false.
     */
    boolean SetDllDirectory(String lpPathName);
}
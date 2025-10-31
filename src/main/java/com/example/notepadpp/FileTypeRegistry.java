package com.example.notepadpp;

import java.util.HashMap;
import java.util.Map;

/**
 * Registry mapping file extensions to editor types.
 */
public class FileTypeRegistry {
    private final Map<String, String> registry = new HashMap<>();

    public FileTypeRegistry() {
        // Register default file types
        registerFileType("txt", "text");
        registerFileType("java", "text");
        registerFileType("md", "text");
        registerFileType("json", "text");
        registerFileType("xml", "text");
        registerFileType("html", "text");
        registerFileType("htm", "text");
        registerFileType("css", "text");
        registerFileType("js", "text");
        registerFileType("py", "text");
        registerFileType("c", "text");
        registerFileType("cpp", "text");
        registerFileType("h", "text");
        registerFileType("sql", "text");
        registerFileType("csv", "table");
        registerFileType("xlsx", "table");
        registerFileType("pptx", "pptx_editable");
        registerFileType("png", "image");
        registerFileType("jpg", "image");
        registerFileType("jpeg", "image");
        registerFileType("gif", "image");
        registerFileType("bmp", "image");
        registerFileType("zip", "archive");
    }

    public void registerFileType(String extension, String editorType) {
        registry.put(extension.toLowerCase(), editorType);
    }

    public String getEditorTypeForFile(String fileName) {
        int dotIndex = fileName.lastIndexOf('.');
        if (dotIndex > 0 && dotIndex < fileName.length() - 1) {
            String ext = fileName.substring(dotIndex + 1).toLowerCase();
            return registry.getOrDefault(ext, "text"); // Default to text editor
        }
        return "text";
    }
}
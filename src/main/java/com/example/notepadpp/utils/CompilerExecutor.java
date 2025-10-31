package com.example.notepadpp.utils;

import java.io.BufferedReader;
import java.io.File;
import java.io.IOException;
import java.io.InputStreamReader;
import java.nio.charset.StandardCharsets;

import javax.swing.JTextArea;

public class CompilerExecutor {

    public static void compileAndRun(File file, JTextArea console) {
        String ext = getExt(file);
        if (ext == null) {
            console.append("Unsupported file type.\n");
            return;
        }

        try {
            ProcessBuilder pb = buildProcess(file, ext);
            if (pb == null) {
                console.append("No runner for ." + ext + "\n");
                return;
            }

            Process proc = pb.start();

            // read stdout
            try (BufferedReader out = new BufferedReader(
                     new InputStreamReader(proc.getInputStream(), StandardCharsets.UTF_8));
                 BufferedReader err = new BufferedReader(
                     new InputStreamReader(proc.getErrorStream(), StandardCharsets.UTF_8))
            ) {
                String line;
                while ((line = out.readLine()) != null) {
                    console.append(line + "\n");
                }
                while ((line = err.readLine()) != null) {
                    console.append("[ERROR] " + line + "\n");
                }
            }

        } catch (IOException ex) {
            console.append("Execution failed: " + ex.getMessage() + "\n");
        }
    }

    private static String getExt(File f) {
        String n = f.getName();
        int i = n.lastIndexOf('.');
        return (i > 0) ? n.substring(i+1).toLowerCase() : null;
    }

    private static ProcessBuilder buildProcess(File f, String ext) {
        String path = f.getAbsolutePath();
        switch (ext) {
            case "py":
                return new ProcessBuilder("python", path);
            case "java":
                // compile, then run
                String dir = f.getParent();
                String cls = f.getName().replace(".java", "");
                return new ProcessBuilder("bash", "-c",
                    "javac \"" + path + "\" && " +
                    "java -cp \"" + dir + "\" " + cls
                );
            case "c", "cpp":
                String exe = path.replaceFirst("\\.(c|cpp)$", "");
                return new ProcessBuilder("gcc", path, "-o", exe, "&&", exe);
            default:
                return null;
        }
    }
}

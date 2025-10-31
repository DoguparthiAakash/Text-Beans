package com.example.notepadpp.media;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.PipedInputStream;
import java.nio.file.Files;
import java.util.concurrent.CompletableFuture;
import java.util.concurrent.TimeUnit;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class CompilerService {
    private boolean isWindows = System.getProperty("os.name").toLowerCase().startsWith("windows");
    private boolean isMac = System.getProperty("os.name").toLowerCase().startsWith("mac");
    private boolean isLinux = System.getProperty("os.name").toLowerCase().startsWith("linux");
    
    // Original method using system terminal
    public CompletableFuture<String> compileAndRun(String sourceCode, String userInput, String language) {
        return CompletableFuture.supplyAsync(() -> {
            try {
                File tempDir = Files.createTempDirectory("code-runner-" + System.currentTimeMillis()).toFile();
                
                switch (language) {
                    case "Java":
                        return executeJavaInTerminal(sourceCode, userInput, tempDir);
                    case "Python":
                        return executePythonInTerminal(sourceCode, userInput, tempDir);
                    case "C++":
                        return executeCppInTerminal(sourceCode, userInput, tempDir);
                    case "C":
                        return executeCInTerminal(sourceCode, userInput, tempDir);
                    default:
                        return "Unsupported language: " + language;
                }
            } catch (Exception ex) {
                return "Execution error: " + ex.getMessage();
            }
        });
    }
    
    // Stream-based execution method
    public CompletableFuture<String> compileAndRunWithStream(String sourceCode, PipedInputStream inputStream, String language) {
        return CompletableFuture.supplyAsync(() -> {
            try {
                File tempDir = Files.createTempDirectory("code-runner-" + System.currentTimeMillis()).toFile();
                
                switch (language) {
                    case "Java":
                        return executeJavaWithStream(sourceCode, inputStream, tempDir);
                    case "Python":
                        return executePythonWithStream(sourceCode, inputStream, tempDir);
                    case "C++":
                        return executeCppWithStream(sourceCode, inputStream, tempDir);
                    case "C":
                        return executeCWithStream(sourceCode, inputStream, tempDir);
                    default:
                        return "Unsupported language: " + language;
                }
            } catch (Exception ex) {
                return "Execution error: " + ex.getMessage();
            }
        });
    }
    
    // Stream execution methods for each language
    private String executeJavaWithStream(String sourceCode, InputStream inputStream, File tempDir) throws Exception {
        String className = "MyProgram";
        Pattern pattern = Pattern.compile("public class (\\w+)");
        Matcher matcher = pattern.matcher(sourceCode);
        if (matcher.find()) className = matcher.group(1);
        
        File javaFile = new File(tempDir, className + ".java");
        Files.writeString(javaFile.toPath(), sourceCode);
        
        // Compile
        Process compileProcess = new ProcessBuilder("javac", javaFile.getName())
            .directory(tempDir)
            .start();
        
        String compileErrors = readStream(compileProcess.getErrorStream());
        if (compileProcess.waitFor() != 0) {
            cleanup(tempDir);
            return "COMPILATION FAILED:\n" + compileErrors;
        }
        
        // Execute with stream input
        Process runProcess = new ProcessBuilder("java", className)
            .directory(tempDir)
            .start();
        
        return handleProcessWithStream(runProcess, inputStream, tempDir);
    }
    
    private String executePythonWithStream(String sourceCode, InputStream inputStream, File tempDir) throws Exception {
        File pythonFile = new File(tempDir, "script.py");
        Files.writeString(pythonFile.toPath(), sourceCode);
        
        ProcessBuilder pythonPb;
        try {
            pythonPb = new ProcessBuilder("python3", pythonFile.getName());
        } catch (Exception e) {
            pythonPb = new ProcessBuilder("python", pythonFile.getName());
        }
        
        Process pythonProcess = pythonPb.directory(tempDir).start();
        return handleProcessWithStream(pythonProcess, inputStream, tempDir);
    }
    
    private String executeCppWithStream(String sourceCode, InputStream inputStream, File tempDir) throws Exception {
        return executeCFamilyWithStream(sourceCode, inputStream, tempDir, "g++", ".cpp");
    }
    
    private String executeCWithStream(String sourceCode, InputStream inputStream, File tempDir) throws Exception {
        return executeCFamilyWithStream(sourceCode, inputStream, tempDir, "gcc", ".c");
    }
    
    private String executeCFamilyWithStream(String sourceCode, InputStream inputStream, File tempDir, 
                                          String compiler, String extension) throws Exception {
        String sourceFileName = "program" + extension;
        String exeName = isWindows ? "program.exe" : "program";
        
        File sourceFile = new File(tempDir, sourceFileName);
        Files.writeString(sourceFile.toPath(), sourceCode);
        
        // Compile
        Process compileProcess = new ProcessBuilder(compiler, sourceFileName, "-o", exeName)
            .directory(tempDir)
            .start();
        
        String compileErrors = readStream(compileProcess.getErrorStream());
        if (compileProcess.waitFor() != 0) {
            cleanup(tempDir);
            return "COMPILATION FAILED:\n" + compileErrors;
        }
        
        // Execute with stream input
        String runCommand = isWindows ? exeName : "./" + exeName;
        Process runProcess = new ProcessBuilder(runCommand)
            .directory(tempDir)
            .start();
        
        return handleProcessWithStream(runProcess, inputStream, tempDir);
    }
    
    // Stream-based process handling
    private String handleProcessWithStream(Process process, InputStream inputStream, File tempDir) throws Exception {
        // Pipe input stream to process in a separate thread
        Thread inputThread = new Thread(() -> {
            try (OutputStream stdin = process.getOutputStream()) {
                byte[] buffer = new byte[1024];
                int bytesRead;
                while ((bytesRead = inputStream.read(buffer)) != -1) {
                    stdin.write(buffer, 0, bytesRead);
                    stdin.flush();
                }
            } catch (IOException e) {
                // Input stream closed, that's expected
            }
        });
        inputThread.start();
        
        // Read output and error streams concurrently
        CompletableFuture<String> outputFuture = CompletableFuture.supplyAsync(() -> 
            readStream(process.getInputStream()));
        
        CompletableFuture<String> errorFuture = CompletableFuture.supplyAsync(() -> 
            readStream(process.getErrorStream()));
        
        // Wait for process completion with timeout
        boolean finished = process.waitFor(10, TimeUnit.SECONDS);
        String output = outputFuture.get(2, TimeUnit.SECONDS);
        String errors = errorFuture.get(2, TimeUnit.SECONDS);
        
        StringBuilder result = new StringBuilder();
        if (!finished) {
            process.destroy();
            result.append("PROCESS TIMEOUT: Program took too long to execute\n");
        }
        
        result.append(output);
        if (!errors.isEmpty()) {
            result.append("ERROR: ").append(errors);
        }
        
        // Wait for input thread to finish
        try {
            inputThread.join(1000);
        } catch (InterruptedException e) {
            // Ignore
        }
        
        cleanup(tempDir);
        return result.toString();
    }
    
    // Terminal execution methods - UPDATED FOR WINDOWS CMD/POWERSHELL
    private String executeJavaInTerminal(String sourceCode, String userInput, File tempDir) throws Exception {
        String className = "MyProgram";
        Pattern pattern = Pattern.compile("public class (\\w+)");
        Matcher matcher = pattern.matcher(sourceCode);
        if (matcher.find()) className = matcher.group(1);
        
        File javaFile = new File(tempDir, className + ".java");
        Files.writeString(javaFile.toPath(), sourceCode);
        
        // Compile first
        Process compileProcess = new ProcessBuilder("javac", javaFile.getName())
            .directory(tempDir)
            .start();
        
        String compileErrors = readStream(compileProcess.getErrorStream());
        if (compileProcess.waitFor() != 0) {
            cleanup(tempDir);
            return "COMPILATION FAILED:\n" + compileErrors;
        }
        
        // Create command for terminal execution
        String runCommand = "cd \"" + tempDir.getAbsolutePath() + "\" && java " + className;
        if (!userInput.isEmpty()) {
            runCommand += " && echo Program finished with exit code %errorlevel%";
        }
        
        return executeInWindowsTerminal(runCommand, userInput, tempDir);
    }
    
    private String executePythonInTerminal(String sourceCode, String userInput, File tempDir) throws Exception {
        File pythonFile = new File(tempDir, "script.py");
        Files.writeString(pythonFile.toPath(), sourceCode);
        
        String pythonCmd;
        try {
            Process testProcess = new ProcessBuilder("python3", "--version").start();
            if (testProcess.waitFor() == 0) {
                pythonCmd = "python3";
            } else {
                pythonCmd = "python";
            }
        } catch (Exception e) {
            pythonCmd = "python";
        }
        
        String runCommand = "cd \"" + tempDir.getAbsolutePath() + "\" && " + pythonCmd + " script.py";
        if (!userInput.isEmpty()) {
            runCommand += " && echo Program finished with exit code %errorlevel%";
        }
        
        return executeInWindowsTerminal(runCommand, userInput, tempDir);
    }
    
    private String executeCppInTerminal(String sourceCode, String userInput, File tempDir) throws Exception {
        return executeCFamilyInTerminal(sourceCode, userInput, tempDir, "g++", ".cpp");
    }
    
    private String executeCInTerminal(String sourceCode, String userInput, File tempDir) throws Exception {
        return executeCFamilyInTerminal(sourceCode, userInput, tempDir, "gcc", ".c");
    }
    
    private String executeCFamilyInTerminal(String sourceCode, String userInput, File tempDir, 
                                          String compiler, String extension) throws Exception {
        String sourceFileName = "program" + extension;
        String exeName = isWindows ? "program.exe" : "program";
        
        File sourceFile = new File(tempDir, sourceFileName);
        Files.writeString(sourceFile.toPath(), sourceCode);
        
        // Compile first
        Process compileProcess = new ProcessBuilder(compiler, sourceFileName, "-o", exeName)
            .directory(tempDir)
            .start();
        
        String compileErrors = readStream(compileProcess.getErrorStream());
        if (compileProcess.waitFor() != 0) {
            cleanup(tempDir);
            return "COMPILATION FAILED:\n" + compileErrors;
        }
        
        String runCommand = "cd \"" + tempDir.getAbsolutePath() + "\" && " + (isWindows ? exeName : "./" + exeName);
        if (!userInput.isEmpty()) {
            runCommand += " && echo Program finished with exit code %errorlevel%";
        }
        
        return executeInWindowsTerminal(runCommand, userInput, tempDir);
    }
    
    // NEW: Windows-specific terminal execution with CMD/PowerShell
    private String executeInWindowsTerminal(String command, String userInput, File tempDir) throws Exception {
        if (isWindows) {
            // Try PowerShell first, then fall back to CMD
            return executeInPowerShellOrCmd(command, userInput, tempDir);
        } else {
            // For non-Windows systems, use the existing terminal execution
            return executeInTerminal(command, tempDir);
        }
    }
    
    // NEW: Execute in PowerShell with fallback to CMD
    private String executeInPowerShellOrCmd(String command, String userInput, File tempDir) throws Exception {
        // First try PowerShell
        try {
            return executeInPowerShell(command, userInput, tempDir);
        } catch (Exception e) {
            // Fall back to CMD if PowerShell fails
            return executeInCmd(command, userInput, tempDir);
        }
    }
    
    // NEW: Execute using PowerShell
    private String executeInPowerShell(String command, String userInput, File tempDir) throws Exception {
        // Create a PowerShell script that will run the command and wait
        String psScript = String.format(
            "Start-Process powershell -ArgumentList '-NoExit', '-Command', \"%s\" -Wait",
            command.replace("\"", "\\\"")
        );
        
        ProcessBuilder processBuilder = new ProcessBuilder("powershell", "-Command", psScript);
        
        Process process = processBuilder.start();
        
        // For PowerShell, we don't wait for the window to close as it stays open
        // Just return a success message
        cleanup(tempDir);
        return "Program launched in PowerShell. Check the PowerShell window for output.\nCommand: " + command;
    }
    
    // NEW: Execute using CMD
    private String executeInCmd(String command, String userInput, File tempDir) throws Exception {
        // Create a CMD command that will run and wait
        String cmdCommand = String.format("cmd /c \"%s && pause\"", command);
        
        ProcessBuilder processBuilder = new ProcessBuilder("cmd", "/c", "start", "cmd", "/k", cmdCommand);
        
        Process process = processBuilder.start();
        
        // For CMD, we don't wait for the window to close as it stays open
        // Just return a success message
        cleanup(tempDir);
        return "Program launched in CMD. Check the CMD window for output.\nCommand: " + command;
    }
    
    // UPDATED: Core terminal execution method (for non-Windows)
    private String executeInTerminal(String command, File tempDir) throws Exception {
        ProcessBuilder processBuilder;
        
        if (isWindows) {
            // This shouldn't be called for Windows anymore, but keep as fallback
            return executeInWindowsTerminal(command, "", tempDir);
        } else if (isMac) {
            // macOS: use Terminal.app or iTerm
            String appleScript = String.format(
                "tell application \"Terminal\"\n" +
                "    activate\n" +
                "    do script \"%s; echo 'Press any key to continue...'; read -n 1; exit\"\n" +
                "end tell", 
                command.replace("\"", "\\\"")
            );
            
            processBuilder = new ProcessBuilder("osascript", "-e", appleScript);
            
        } else {
            // Linux: use xterm, gnome-terminal, or konsole
            String terminalCmd = detectLinuxTerminal();
            if ("xterm".equals(terminalCmd)) {
                processBuilder = new ProcessBuilder("xterm", "-e", "bash", "-c", 
                    command + "; echo 'Press any key to continue...'; read -n 1");
            } else if ("gnome-terminal".equals(terminalCmd)) {
                processBuilder = new ProcessBuilder("gnome-terminal", "--", "bash", "-c", 
                    command + "; echo 'Press any key to continue...'; read -n 1");
            } else if ("konsole".equals(terminalCmd)) {
                processBuilder = new ProcessBuilder("konsole", "-e", "bash", "-c", 
                    command + "; echo 'Press any key to continue...'; read -n 1");
            } else {
                // Fallback to direct execution without terminal
                return executeDirectly(command, tempDir);
            }
        }
        
        try {
            Process process = processBuilder.start();
            
            if (isMac || isWindows) {
                // For macOS and Windows, we don't wait for the terminal app to close
                cleanup(tempDir);
                return "Program launched in terminal. Check the terminal application for output.\nCommand: " + command;
            }
            
            // For Linux, wait for process completion
            CompletableFuture<String> outputFuture = CompletableFuture.supplyAsync(() -> 
                readStream(process.getInputStream()));
            
            CompletableFuture<String> errorFuture = CompletableFuture.supplyAsync(() -> 
                readStream(process.getErrorStream()));
            
            boolean finished = process.waitFor(30, TimeUnit.SECONDS);
            String output = outputFuture.get(5, TimeUnit.SECONDS);
            String errors = errorFuture.get(5, TimeUnit.SECONDS);
            
            StringBuilder result = new StringBuilder();
            if (!finished) {
                process.destroy();
                result.append("PROCESS TIMEOUT: Terminal took too long to execute\n");
            }
            
            result.append(output);
            if (!errors.isEmpty()) {
                result.append("ERROR: ").append(errors);
            }
            
            cleanup(tempDir);
            return result.toString();
            
        } catch (Exception e) {
            // Fallback to direct execution if terminal launch fails
            return executeDirectly(command, tempDir);
        }
    }
    
    // Fallback method for direct execution without terminal
    private String executeDirectly(String command, File tempDir) throws Exception {
        Process process;
        if (isWindows) {
            process = new ProcessBuilder("cmd.exe", "/c", command).start();
        } else {
            process = new ProcessBuilder("bash", "-c", command).start();
        }
        
        CompletableFuture<String> outputFuture = CompletableFuture.supplyAsync(() -> 
            readStream(process.getInputStream()));
        
        CompletableFuture<String> errorFuture = CompletableFuture.supplyAsync(() -> 
            readStream(process.getErrorStream()));
        
        boolean finished = process.waitFor(10, TimeUnit.SECONDS);
        String output = outputFuture.get(2, TimeUnit.SECONDS);
        String errors = errorFuture.get(2, TimeUnit.SECONDS);
        
        StringBuilder result = new StringBuilder();
        if (!finished) {
            process.destroy();
            result.append("PROCESS TIMEOUT: Program took too long to execute\n");
        }
        
        result.append("DIRECT EXECUTION (Terminal not available):\n");
        result.append(output);
        if (!errors.isEmpty()) {
            result.append("ERROR: ").append(errors);
        }
        
        cleanup(tempDir);
        return result.toString();
    }
    
    // Detect available terminal on Linux
    private String detectLinuxTerminal() {
        String[] terminals = {"gnome-terminal", "konsole", "xterm"};
        for (String terminal : terminals) {
            try {
                Process process = new ProcessBuilder("which", terminal).start();
                if (process.waitFor() == 0) {
                    return terminal;
                }
            } catch (Exception e) {
                // Continue to next terminal
            }
        }
        return "none";
    }
    
    // Alternative method that returns the terminal command for external handling
    public CompletableFuture<String> getTerminalCommand(String sourceCode, String userInput, String language) {
        return CompletableFuture.supplyAsync(() -> {
            try {
                File tempDir = Files.createTempDirectory("code-runner-" + System.currentTimeMillis()).toFile();
                
                switch (language) {
                    case "Java":
                        return buildJavaTerminalCommand(sourceCode, userInput, tempDir);
                    case "Python":
                        return buildPythonTerminalCommand(sourceCode, userInput, tempDir);
                    case "C++":
                        return buildCppTerminalCommand(sourceCode, userInput, tempDir);
                    case "C":
                        return buildCTerminalCommand(sourceCode, userInput, tempDir);
                    default:
                        return "Unsupported language: " + language;
                }
            } catch (Exception ex) {
                return "Error: " + ex.getMessage();
            }
        });
    }
    
    // Command building methods
    private String buildJavaTerminalCommand(String sourceCode, String userInput, File tempDir) throws Exception {
        String className = "MyProgram";
        Pattern pattern = Pattern.compile("public class (\\w+)");
        Matcher matcher = pattern.matcher(sourceCode);
        if (matcher.find()) className = matcher.group(1);
        
        File javaFile = new File(tempDir, className + ".java");
        Files.writeString(javaFile.toPath(), sourceCode);
        
        // Compile first
        Process compileProcess = new ProcessBuilder("javac", javaFile.getName())
            .directory(tempDir)
            .start();
        
        String compileErrors = readStream(compileProcess.getErrorStream());
        if (compileProcess.waitFor() != 0) {
            cleanup(tempDir);
            return "COMPILATION FAILED:\n" + compileErrors;
        }
        
        return "cd \"" + tempDir.getAbsolutePath() + "\" && java " + className;
    }
    
    private String buildPythonTerminalCommand(String sourceCode, String userInput, File tempDir) throws Exception {
        File pythonFile = new File(tempDir, "script.py");
        Files.writeString(pythonFile.toPath(), sourceCode);
        
        String pythonCmd;
        try {
            Process testProcess = new ProcessBuilder("python3", "--version").start();
            if (testProcess.waitFor() == 0) {
                pythonCmd = "python3";
            } else {
                pythonCmd = "python";
            }
        } catch (Exception e) {
            pythonCmd = "python";
        }
        
        return "cd \"" + tempDir.getAbsolutePath() + "\" && " + pythonCmd + " script.py";
    }
    
    private String buildCppTerminalCommand(String sourceCode, String userInput, File tempDir) throws Exception {
        return buildCFamilyTerminalCommand(sourceCode, userInput, tempDir, "g++", ".cpp");
    }
    
    private String buildCTerminalCommand(String sourceCode, String userInput, File tempDir) throws Exception {
        return buildCFamilyTerminalCommand(sourceCode, userInput, tempDir, "gcc", ".c");
    }
    
    private String buildCFamilyTerminalCommand(String sourceCode, String userInput, File tempDir, 
                                             String compiler, String extension) throws Exception {
        String sourceFileName = "program" + extension;
        String exeName = isWindows ? "program.exe" : "program";
        
        File sourceFile = new File(tempDir, sourceFileName);
        Files.writeString(sourceFile.toPath(), sourceCode);
        
        Process compileProcess = new ProcessBuilder(compiler, sourceFileName, "-o", exeName)
            .directory(tempDir)
            .start();
        
        String compileErrors = readStream(compileProcess.getErrorStream());
        if (compileProcess.waitFor() != 0) {
            cleanup(tempDir);
            return "COMPILATION FAILED:\n" + compileErrors;
        }
        
        return "cd \"" + tempDir.getAbsolutePath() + "\" && " + (isWindows ? exeName : "./" + exeName);
    }
    
    // Utility methods
    private String readStream(InputStream stream) {
        try {
            ByteArrayOutputStream result = new ByteArrayOutputStream();
            byte[] buffer = new byte[1024];
            int length;
            while ((length = stream.read(buffer)) != -1) {
                result.write(buffer, 0, length);
            }
            return result.toString("UTF-8");
        } catch (IOException e) {
            return "";
        }
    }
    
    private void cleanup(File tempDir) {
        if (tempDir != null && tempDir.exists()) {
            try {
                // Don't delete immediately - wait a bit for any file handles to be released
                Thread.sleep(1000);
                File[] files = tempDir.listFiles();
                if (files != null) {
                    for (File f : files) {
                        f.delete();
                    }
                }
                tempDir.delete();
            } catch (Exception e) {
                // Ignore cleanup errors
            }
        }
    }
}
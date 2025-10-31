package com.example.notepadpp.ai;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStreamReader;
import java.util.function.Consumer;

public class LocalAIEngine {

    private static final String EXECUTABLE_PATH = "llama/llama-run.exe";
    private static final String MODEL_PATH = "llama/phi-2.Q4_K_M.gguf";

    public static void ask(String prompt, Consumer<String> onResponse) {
        new Thread(() -> {
            StringBuilder result = new StringBuilder();

            try {
                ProcessBuilder pb = new ProcessBuilder(
                    EXECUTABLE_PATH,
                    "--temp", "0.7",
                    "--threads", "4",
                    MODEL_PATH,
                    prompt
                );

                pb.redirectErrorStream(true);
                Process process = pb.start();

                try (BufferedReader reader = new BufferedReader(
                        new InputStreamReader(process.getInputStream()))) {
                    String line;
                    while ((line = reader.readLine()) != null) {
                        result.append(line).append(System.lineSeparator());
                    }
                }

                process.waitFor();
                onResponse.accept(result.toString().trim());

            } catch (IOException | InterruptedException e) {
                onResponse.accept("⚠️ AI Error: " + e.getMessage());
            }

        }).start(); // ✅ Run in background thread
    }
}

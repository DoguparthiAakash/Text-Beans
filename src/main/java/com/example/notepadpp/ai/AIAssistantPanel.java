package com.example.notepadpp.ai;

import javax.swing.*;

import com.example.notepadpp.media.DocLegacyEditorPanel;
import javax.swing.Icon;
import javax.swing.ImageIcon;
import java.awt.*;
import java.util.function.Consumer;

public class AIAssistantPanel extends JPanel {
    private final JTextArea inputArea;
    private final JTextArea outputArea;
    private final JButton askButton;

    public AIAssistantPanel() {
        setLayout(new BorderLayout());

        // Input area setup
        inputArea = new JTextArea(5, 40);
        JScrollPane inputScroll = new JScrollPane(inputArea);

        // Output area setup
        outputArea = new JTextArea(10, 40);
        outputArea.setEditable(false);
        JScrollPane outputScroll = new JScrollPane(outputArea);

        // Button setup
        askButton = new JButton("Ask AI");
        askButton.addActionListener(e -> {
            String userText = inputArea.getText().trim();
            if (!userText.isEmpty()) {
                outputArea.setText("🤖 Thinking...");
                askButton.setEnabled(false);

                // Async call to LocalAIEngine
                LocalAIEngine.ask(userText, new Consumer<>() {
                    @Override
                    public void accept(String response) {
                        SwingUtilities.invokeLater(() -> {
                            outputArea.setText(response);
                            askButton.setEnabled(true);
                        });
                    }
                });
            }
        });

        // Layout for input + button
        JPanel inputPanel = new JPanel(new BorderLayout());
        inputPanel.add(inputScroll, BorderLayout.CENTER);
        inputPanel.add(askButton, BorderLayout.EAST);

        // Add to main panel
        add(inputPanel, BorderLayout.NORTH);
        add(outputScroll, BorderLayout.CENTER);
    }
    public static Icon getTabIcon() {
        ImageIcon originalIcon = new ImageIcon(DocLegacyEditorPanel.class.getResource("/icons/chatbot.png"));
        int fixedWidth = 16;
        int fixedHeight = 16;
        Image scaledImage = originalIcon.getImage().getScaledInstance(fixedWidth, fixedHeight, Image.SCALE_SMOOTH);
        return new ImageIcon(scaledImage);
    }
}

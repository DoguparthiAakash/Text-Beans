package com.example.notepadpp.ai;

import javax.swing.*;

public class AIAssistantAddon {
    public static void show() {
        JFrame frame = new JFrame("AI Assistant (Offline)");
        frame.setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);
        frame.setContentPane(new AIAssistantPanel());
        frame.pack();
        frame.setLocationRelativeTo(null);
        frame.setVisible(true);
    }
}
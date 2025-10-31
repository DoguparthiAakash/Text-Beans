package com.example.notepadpp.utils;
import java.awt.Font;

import javax.swing.JButton;
import javax.swing.JTextPane;
import javax.swing.JToolBar;
import javax.swing.text.AttributeSet;
import javax.swing.text.Element;
import javax.swing.text.MutableAttributeSet;
import javax.swing.text.SimpleAttributeSet;
import javax.swing.text.StyleConstants;
import javax.swing.text.StyledDocument;
import javax.swing.text.StyledEditorKit;

public class FormattingToolbar {
    public static JToolBar createFor(JTextPane textPane) {
        JToolBar toolbar = new JToolBar();

        JButton boldBtn = new JButton("B");
        boldBtn.setFont(boldBtn.getFont().deriveFont(Font.BOLD));
        boldBtn.addActionListener(e -> toggleStyle(textPane, StyleConstants.Bold));

        JButton italicBtn = new JButton("I");
        italicBtn.setFont(italicBtn.getFont().deriveFont(Font.ITALIC));
        italicBtn.addActionListener(e -> toggleStyle(textPane, StyleConstants.Italic));

        JButton underlineBtn = new JButton("U");
        underlineBtn.addActionListener(e -> toggleStyle(textPane, StyleConstants.Underline));

        toolbar.add(boldBtn);
        toolbar.add(italicBtn);
        toolbar.add(underlineBtn);

        return toolbar;
    }

    private static void toggleStyle(JTextPane textPane, Object styleKey) {
        int start = textPane.getSelectionStart();
        int end = textPane.getSelectionEnd();

        if (start == end) {
            StyledEditorKit kit = (StyledEditorKit) textPane.getEditorKit();
            MutableAttributeSet attrs = kit.getInputAttributes();

            boolean isSet = isAttributeSet(attrs, styleKey);
            SimpleAttributeSet sas = new SimpleAttributeSet();
            sas.addAttribute(styleKey, !isSet);
            textPane.setCharacterAttributes(sas, false);
        } else {
            StyledDocument doc = textPane.getStyledDocument();
            for (int i = start; i < end; i++) {
                Element element = doc.getCharacterElement(i);
                AttributeSet attr = element.getAttributes();
                boolean isSet = isAttributeSet(attr, styleKey);

                SimpleAttributeSet sas = new SimpleAttributeSet(attr.copyAttributes());
                sas.addAttribute(styleKey, !isSet);
                doc.setCharacterAttributes(i, 1, sas, true);
            }
        }
    }

    private static boolean isAttributeSet(AttributeSet attrs, Object styleKey) {
        if (styleKey == StyleConstants.Bold) return StyleConstants.isBold(attrs);
        if (styleKey == StyleConstants.Italic) return StyleConstants.isItalic(attrs);
        if (styleKey == StyleConstants.Underline) return StyleConstants.isUnderline(attrs);
        return false;
    }
}

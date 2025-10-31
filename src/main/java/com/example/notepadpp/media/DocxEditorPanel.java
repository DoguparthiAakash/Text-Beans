package com.example.notepadpp.media;

import java.awt.BorderLayout;
import java.awt.Color;
import java.awt.Image;
import java.awt.event.ActionEvent;
import java.awt.print.PrinterException;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;

import javax.swing.Icon;
import javax.swing.ImageIcon;
import javax.swing.JColorChooser;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTextPane;
import javax.swing.text.AttributeSet;
import javax.swing.text.Element;
import javax.swing.text.SimpleAttributeSet;
import javax.swing.text.StyleConstants;
import javax.swing.text.StyledDocument;
import javax.swing.undo.UndoManager;

import org.apache.poi.util.Units;
import org.apache.poi.wp.usermodel.HeaderFooterType;
import org.apache.poi.xwpf.usermodel.UnderlinePatterns;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFHeader;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import com.example.notepadpp.utils.FormattingToolbar;

public class DocxEditorPanel extends JPanel {
    private final File file;
    private final JTextPane textPane = new JTextPane();
    private final StyledDocument doc = textPane.getStyledDocument();
    
    // UndoManager to support undo/redo actions.
    private final UndoManager undoManager = new UndoManager();
    
    // Previously, the logo was inserted via an external file.
    // Now we use an embedded resource using getResourceAsStream.
    // Remove the logoFile member variable and related insertLogo() method.
    
    public DocxEditorPanel(File file) {
        this.file = file;
        setLayout(new BorderLayout());
        
        // Use the shared formatting toolbar.
        add(FormattingToolbar.createFor(textPane), BorderLayout.NORTH);
        add(new JScrollPane(textPane), BorderLayout.CENTER);
        
        // Setup undo/redo support.
        doc.addUndoableEditListener(e -> undoManager.addEdit(e.getEdit()));
        
        // Load DOCX file content.
        loadDocx();
    }
    
    private void loadDocx() {
        try (FileInputStream fis = new FileInputStream(file)) {
            XWPFDocument xdoc = new XWPFDocument(fis);
            
            for (XWPFParagraph p : xdoc.getParagraphs()) {
                for (XWPFRun run : p.getRuns()) {
                    String text = run.text();
                    if (text == null || text.isEmpty())
                        continue;
                    
                    SimpleAttributeSet attrs = new SimpleAttributeSet();
                    StyleConstants.setBold(attrs, run.isBold());
                    StyleConstants.setItalic(attrs, run.isItalic());
                    StyleConstants.setUnderline(attrs, run.getUnderline().getValue() > 0);
                    if (run.getFontSize() > 0)
                        StyleConstants.setFontSize(attrs, run.getFontSize());
                    if (run.getFontFamily() != null)
                        StyleConstants.setFontFamily(attrs, run.getFontFamily());
                    if (run.getColor() != null) {
                        try {
                            StyleConstants.setForeground(attrs, Color.decode("#" + run.getColor()));
                        } catch (Exception ignored) {}
                    }
                    
                    doc.insertString(doc.getLength(), text, attrs);
                }
                doc.insertString(doc.getLength(), "\n", null);
            }
        } catch (Exception e) {
            JOptionPane.showMessageDialog(this, "❌ Failed to load .docx:\n" + e.getMessage());
        }
    }
    
    private void saveDocx(ActionEvent e) {
        try (FileOutputStream out = new FileOutputStream(file)) {
            XWPFDocument xdoc = new XWPFDocument();
            
            // Use the embedded logo located by the class loader.
            try (InputStream logoStream = getClass().getResourceAsStream("/icons/worddocument.png")) {
                if (logoStream != null) {
                    XWPFHeader header = xdoc.createHeader(HeaderFooterType.DEFAULT);
                    XWPFParagraph headerPara = header.createParagraph();
                    XWPFRun headerRun = headerPara.createRun();
                    // Use PNG type; adjust if your logo is JPEG.
                    headerRun.addPicture(logoStream, XWPFDocument.PICTURE_TYPE_PNG, "worddocument.png", 
                            Units.toEMU(150), Units.toEMU(80));
                } else {
                    System.err.println("Logo resource not found!");
                }
            } catch (Exception logoEx) {
                JOptionPane.showMessageDialog(this, "⚠️ Failed to insert logo:\n" + logoEx.getMessage());
            }
            
            // Process text content from the StyledDocument.
            Element root = doc.getDefaultRootElement();
            for (int i = 0; i < root.getElementCount(); i++) {
                Element paragraphElem = root.getElement(i);
                XWPFParagraph para = xdoc.createParagraph();
                
                int start = paragraphElem.getStartOffset();
                int end = paragraphElem.getEndOffset();
                
                for (int j = start; j < end; ) {
                    Element runElem = doc.getCharacterElement(j);
                    int runStart = runElem.getStartOffset();
                    int runEnd = runElem.getEndOffset();
                    
                    String text = doc.getText(runStart, runEnd - runStart);
                    if (!text.isBlank()) {
                        AttributeSet attrs = runElem.getAttributes();
                        XWPFRun run = para.createRun();
                        run.setText(text);
                        run.setBold(StyleConstants.isBold(attrs));
                        run.setItalic(StyleConstants.isItalic(attrs));
                        run.setUnderline(StyleConstants.isUnderline(attrs) ? UnderlinePatterns.SINGLE : UnderlinePatterns.NONE);
                        run.setFontSize(StyleConstants.getFontSize(attrs));
                        run.setFontFamily(StyleConstants.getFontFamily(attrs));
                        Color fg = StyleConstants.getForeground(attrs);
                        if (fg != null) {
                            run.setColor(String.format("%02x%02x%02x", fg.getRed(), fg.getGreen(), fg.getBlue()));
                        }
                    }
                    j = runEnd;
                }
            }
            
            xdoc.write(out);
            JOptionPane.showMessageDialog(this, "✅ .docx saved successfully!");
        } catch (Exception ex) {
            JOptionPane.showMessageDialog(this, "❌ Failed to save .docx:\n" + ex.getMessage());
        }
    }
    
    // --------------------
    // NEW FEATURES
    // --------------------
    
    // Undo the last change.
    public void undo() {
        if (undoManager.canUndo()) {
            undoManager.undo();
        }
    }
    
    // Redo the last undone change.
    public void redo() {
        if (undoManager.canRedo()) {
            undoManager.redo();
        }
    }
    
    // A simple Find & Replace dialog (operates on plain text).
    public void showFindReplaceDialog() {
        String findText = JOptionPane.showInputDialog(this, "Enter text to find:");
        if (findText == null || findText.isEmpty()) return;
        String replaceText = JOptionPane.showInputDialog(this, "Replace with:");
        if (replaceText == null) return;
        
        String content = textPane.getText();
        content = content.replace(findText, replaceText);
        textPane.setText(content);
    }
    
    // Print the document.
    public void printDocument() {
        try {
            boolean complete = textPane.print();
            if (complete) {
                JOptionPane.showMessageDialog(this, "Printing complete!", "Print", JOptionPane.INFORMATION_MESSAGE);
            } else {
                JOptionPane.showMessageDialog(this, "Printing canceled.", "Print", JOptionPane.WARNING_MESSAGE);
            }
        } catch (PrinterException e) {
            JOptionPane.showMessageDialog(this, "Printing failed: " + e.getMessage(), "Print Error", JOptionPane.ERROR_MESSAGE);
        }
    }
    
    // Change the text color of the selected text.
    public void changeTextColor() {
        Color newColor = JColorChooser.showDialog(this, "Choose Text Color", textPane.getForeground());
        if (newColor != null) {
            int start = textPane.getSelectionStart();
            int end = textPane.getSelectionEnd();
            if (start != end) {
                SimpleAttributeSet attrs = new SimpleAttributeSet();
                StyleConstants.setForeground(attrs, newColor);
                doc.setCharacterAttributes(start, end - start, attrs, false);
            }
        }
    }
   

    public static Icon getTabIcon() {
        ImageIcon originalIcon = new ImageIcon(DocxEditorPanel.class.getResource("/icons/word.png"));
        int fixedWidth = 16;
        int fixedHeight = 16;
        Image scaledImage = originalIcon.getImage().getScaledInstance(fixedWidth, fixedHeight, Image.SCALE_SMOOTH);
        return new ImageIcon(scaledImage);
    }
    
    
    

    // --------------------
    // Optionally, you can link the new features (undo/redo, find/replace, print,
    // change text color) to toolbar buttons or key bindings.
}
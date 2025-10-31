package com.example.notepadpp.media;
import java.awt.BorderLayout;
import java.awt.Color;
import java.awt.Image;
import java.awt.event.ActionEvent;
import java.awt.print.PrinterException;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import javax.swing.Icon;
import javax.swing.ImageIcon;
import javax.swing.JColorChooser;
import javax.swing.JFileChooser;
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
public class DocLegacyEditorPanel extends JPanel {
    private final File file;
    private final JTextPane textPane = new JTextPane();
    private final StyledDocument doc = textPane.getStyledDocument();
    private final UndoManager undoManager = new UndoManager();
    private File logoFile = null;
    public DocLegacyEditorPanel(File file) {
        this.file = file;
        setLayout(new BorderLayout());
        add(FormattingToolbar.createFor(textPane), BorderLayout.NORTH);
        add(new JScrollPane(textPane), BorderLayout.CENTER);
        doc.addUndoableEditListener(e -> undoManager.addEdit(e.getEdit()));
        loadDocLegacy();
    }
    private void loadDocLegacy() {
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
            JOptionPane.showMessageDialog(this, "❌ Failed to load .doc:\n" + e.getMessage());
        }
    }
    private void saveDocLegacy(ActionEvent e) {
        try (FileOutputStream out = new FileOutputStream(file)) {
            XWPFDocument xdoc = new XWPFDocument();
            
            // If a logo file was inserted, add it as a header.
            if (logoFile != null) {
                try (FileInputStream logoFis = new FileInputStream(logoFile)) {
                    XWPFHeader header = xdoc.createHeader(HeaderFooterType.DEFAULT);

                    XWPFParagraph headerPara = header.createParagraph();
                    XWPFRun headerRun = headerPara.createRun();
                    // Here, we use PNG type; adjust if your logo is JPEG (use XWPFDocument.PICTURE_TYPE_JPEG)
                    headerRun.addPicture(logoFis, XWPFDocument.PICTURE_TYPE_PNG, logoFile.getName(), 
                            Units.toEMU(150), Units.toEMU(80));
                } catch (Exception logoEx) {
                    JOptionPane.showMessageDialog(this, "⚠️ Failed to insert logo:\n" + logoEx.getMessage());
                }
            }
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
            JOptionPane.showMessageDialog(this, "✅ .doc saved successfully!");
        } catch (Exception ex) {
            JOptionPane.showMessageDialog(this, "❌ Failed to save .doc:\n" + ex.getMessage());
        }
    }
    public void undo() {
        if (undoManager.canUndo()) {
            undoManager.undo();
        }
    }    public void redo() {
        if (undoManager.canRedo()) {
            undoManager.redo();
        }
    }
    public void showFindReplaceDialog() {
        String findText = JOptionPane.showInputDialog(this, "Enter text to find:");
        if (findText == null || findText.isEmpty()) return;
        String replaceText = JOptionPane.showInputDialog(this, "Replace with:");
        if (replaceText == null) return;
        
        String content = textPane.getText();
        content = content.replace(findText, replaceText);
        textPane.setText(content);
    }
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
    public void insertLogo() {
        JFileChooser chooser = new JFileChooser();
        int result = chooser.showOpenDialog(this);
        if (result == JFileChooser.APPROVE_OPTION) {
            logoFile = chooser.getSelectedFile();
            JOptionPane.showMessageDialog(this, "Logo selected: " + logoFile.getName() +
                    "\nIt will appear in the document header after saving.");
        }
    }
    public static Icon getTabIcon() {
        ImageIcon originalIcon = new ImageIcon(DocLegacyEditorPanel.class.getResource("/icons/word.png"));
        int fixedWidth = 16;
        int fixedHeight = 16;
        Image scaledImage = originalIcon.getImage().getScaledInstance(fixedWidth, fixedHeight, Image.SCALE_SMOOTH);
        return new ImageIcon(scaledImage);
    }
}

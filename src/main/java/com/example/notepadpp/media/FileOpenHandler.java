package com.example.notepadpp.media;

import java.awt.BorderLayout;
import java.awt.Component;
import java.awt.Image;
import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileReader;
import java.io.IOException;
import java.util.Map;

import javax.swing.Icon;
import javax.swing.ImageIcon;
import javax.swing.JButton;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTabbedPane;
import javax.swing.JTable;
import javax.swing.JTextArea;
import javax.swing.SwingUtilities;
import javax.swing.text.DefaultStyledDocument;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class FileOpenHandler {

    public static void openWithEditorType(File file, JTabbedPane tabbedPane,
                                          Map<Component, String> tabTypeMap, String editorType) {
        try {
            switch (editorType) {
                case "pptx_editable":
                    addEditablePPTXTab(file, tabbedPane, tabTypeMap);
                    break;
                case "table":
                    openTableFile(file, tabbedPane, tabTypeMap);
                    break;
                case "image":
                    openImageFile(file, tabbedPane, tabTypeMap);
                    break;
                default:
                    openTextFile(file, tabbedPane, tabTypeMap);
            }
        } catch (IOException ex) {
            handleOpenError(ex);
        }
    }

    public static void open(File file, JTabbedPane tabbedPane, Map<Component, String> tabTypeMap) {
        String name = file.getName().toLowerCase();
        try {
            if (name.endsWith(".pdf")) {
                openPDFFile(file, tabbedPane, tabTypeMap);
            } else if (name.matches(".*\\.(png|jpg|jpeg|gif|webp|bmp)$")) {
                ImageIcon icon = new ImageIcon(file.getAbsolutePath());
                JLabel label = new JLabel(icon);
                JScrollPane scroll = new JScrollPane(label);

                ImageIcon imageIcon = new ImageIcon(FileOpenHandler.class.getResource("/icons/image.png"));
                Image scaled = imageIcon.getImage().getScaledInstance(16, 16, Image.SCALE_SMOOTH);
                ImageIcon finalIcon = new ImageIcon(scaled);

                tabbedPane.addTab(file.getName(), finalIcon, scroll);
                tabTypeMap.put(scroll, "image");

            } else if (name.endsWith(".docx")) {
                JPanel docxPanel = new DocxEditorPanel(file);
                tabbedPane.addTab(file.getName(), DocxEditorPanel.getTabIcon(), docxPanel);
                tabTypeMap.put(docxPanel, "docx");
            } else if (name.endsWith(".mp3") || name.endsWith(".wav")) {
                JPanel audioPanel = AudioPlayerPanel.createAudioPanel(file);
                tabbedPane.addTab(file.getName(), AudioPlayerPanel.getTabIcon(), audioPanel);
                tabTypeMap.put(audioPanel, "audio");
            } else if (name.endsWith(".mp4") || name.endsWith(".avi") || name.endsWith(".mov")) {
                JPanel videoPanel = new VLCJVideoPlayerPanel(file);
                tabbedPane.addTab(file.getName(), VLCJVideoPlayerPanel.getTabIcon(), videoPanel);
                tabTypeMap.put(videoPanel, "video");
            } else if (name.endsWith(".xlsx")) {
                JPanel xlsxPanel = new XlsxEditorPanel(file);
                tabbedPane.addTab(file.getName(), XlsxEditorPanel.getTabIcon(), xlsxPanel);
                tabTypeMap.put(xlsxPanel, "xlsx");
            } else if (name.endsWith(".pptx")) {
                JPanel panel = new PptxEditorPanel(file);
                JScrollPane scroll = new JScrollPane(panel);
                tabbedPane.addTab(file.getName(), PptxEditorPanel.getTabIcon(), scroll);
                tabTypeMap.put(scroll, "pptx_editable");
            } else if (name.endsWith(".ppt")) {
                JPanel panel = new PptEditorPanel(file);
                tabbedPane.addTab(file.getName(), PptEditorPanel.getTabIcon(), panel);
                tabTypeMap.put(panel, "ppt");
            } else if (name.endsWith(".doc")) {
                JPanel legacyDoc = new DocLegacyEditorPanel(file);
                tabbedPane.addTab(file.getName(), DocLegacyEditorPanel.getTabIcon(), legacyDoc);
                tabTypeMap.put(legacyDoc, "doc");
            } else if (name.endsWith(".xls")) {
                JPanel xlsPanel = new XlsEditorPanel(file);
                tabbedPane.addTab(file.getName(), XlsEditorPanel.getTabIcon(), xlsPanel);
                tabTypeMap.put(xlsPanel, "xls");
            } else {
                openTextFile(file, tabbedPane, tabTypeMap);
            }
        } catch (IOException ex) {
            handleOpenError(ex);
        }
    }

    private static void openPDFFile(File file, JTabbedPane tabbedPane, Map<Component, String> tabTypeMap) {
        if (!isValidPDFFile(file)) {
            showPDFError(file, "File doesn't appear to be a valid PDF or is corrupted.", tabbedPane);
            return;
        }

        PDDocument doc = null;
        try {
            doc = PDDocument.load(file);

            // Create a placeholder panel for the tab
            JPanel pdfPlaceholderPanel = new JPanel(new BorderLayout());
            
            JLabel infoLabel = new JLabel(
                "<html><center><h2>PDF Document</h2>" +
                "<p>" + file.getName() + "</p>" +
                "<p><i>Opened in separate window</i></p></center></html>",
                JLabel.CENTER
            );
            pdfPlaceholderPanel.add(infoLabel, BorderLayout.CENTER);
            
            // Add button to reopen the PDF viewer
            JButton reopenButton = new JButton("Open PDF Viewer Window");
            JPanel buttonPanel = new JPanel();
            buttonPanel.add(reopenButton);
            pdfPlaceholderPanel.add(buttonPanel, BorderLayout.SOUTH);
            
            final File pdfFile = file;
            reopenButton.addActionListener(e -> {
                try {
                    pdfview viewer = new pdfview(pdfFile);
                    viewer.setVisible(true);
                } catch (IOException ex) {
                    JOptionPane.showMessageDialog(pdfPlaceholderPanel,
                        "Error opening PDF viewer: " + ex.getMessage(),
                        "PDF Error",
                        JOptionPane.ERROR_MESSAGE);
                }
            });
            
            // Add the placeholder panel to tabs
            tabbedPane.addTab(file.getName(), pdfview.getTabIcon(), pdfPlaceholderPanel);
            tabTypeMap.put(pdfPlaceholderPanel, "pdf");
            tabbedPane.setSelectedComponent(pdfPlaceholderPanel);
            
            // Open the PDF viewer window
            final PDDocument finalDoc = doc;
            SwingUtilities.invokeLater(() -> {
                try {
                    pdfview viewer = new pdfview(pdfFile);
                    viewer.setVisible(true);
                } catch (IOException ex) {
                    // If viewer fails, show text mode
                    showPDFInTextMode(finalDoc, file, tabbedPane, tabTypeMap, pdfPlaceholderPanel);
                }
            });

        } catch (IOException ex) {
            handlePDFError(file, ex, tabbedPane);
        } finally {
            // Don't close the document here if we're using it in the viewer
            // The pdfview class will handle closing it
        }
    }
    
    private static void showPDFInTextMode(PDDocument doc, File file, JTabbedPane tabbedPane, 
                                          Map<Component, String> tabTypeMap, JPanel placeholderPanel) {
        try {
            PDFTextStripper stripper = new PDFTextStripper();
            String text = stripper.getText(doc);
            JTextArea area = new JTextArea(text);
            area.setEditable(false);
            JScrollPane scroll = new JScrollPane(area);
            
            // Remove placeholder and add text view
            int index = tabbedPane.indexOfComponent(placeholderPanel);
            if (index >= 0) {
                tabTypeMap.remove(placeholderPanel);
                tabbedPane.removeTabAt(index);
                tabbedPane.insertTab(file.getName(), pdfview.getTabIcon(), scroll, null, index);
                tabTypeMap.put(scroll, "pdf");
                tabbedPane.setSelectedIndex(index);
            }
            
            JOptionPane.showMessageDialog(tabbedPane,
                "PDF opened in text mode. Some features may not be available.",
                "PDF Info", 
                JOptionPane.INFORMATION_MESSAGE);
                
        } catch (IOException ex) {
            JOptionPane.showMessageDialog(tabbedPane,
                "Cannot extract text from PDF: " + ex.getMessage(),
                "PDF Error",
                JOptionPane.ERROR_MESSAGE);
        } finally {
            if (doc != null) {
                try { doc.close(); } catch (IOException ignored) {}
            }
        }
    }

    private static boolean isValidPDFFile(File file) {
        if (file == null || !file.exists() || file.length() == 0) return false;

        try (FileInputStream fis = new FileInputStream(file)) {
            byte[] header = new byte[4];
            if (fis.read(header) != 4) return false;
            return header[0] == 0x25 && header[1] == 0x50 && header[2] == 0x44 && header[3] == 0x46;
        } catch (IOException e) {
            return false;
        }
    }

    private static void handlePDFError(File file, IOException ex, JTabbedPane tabbedPane) {
        String msg = "Cannot open PDF file: " + file.getName() + "\n\nError: " + ex.getMessage();
        showPDFError(file, msg, tabbedPane);
    }

    private static void showPDFError(File file, String msg, JTabbedPane tabbedPane) {
        int option = JOptionPane.showConfirmDialog(tabbedPane,
                msg + "\n\nDo you want to try opening it with your system's PDF viewer instead?",
                "PDF Open Error",
                JOptionPane.YES_NO_OPTION,
                JOptionPane.ERROR_MESSAGE);
        if (option == JOptionPane.YES_OPTION) openWithSystemViewer(file, tabbedPane);
    }

    private static void openWithSystemViewer(File file, JTabbedPane tabbedPane) {
        try {
            if (java.awt.Desktop.isDesktopSupported()) java.awt.Desktop.getDesktop().open(file);
        } catch (Exception ex) {
            JOptionPane.showMessageDialog(tabbedPane,
                    "Cannot open with system viewer: " + ex.getMessage(),
                    "System Viewer Error", JOptionPane.ERROR_MESSAGE);
        }
    }

    private static void addEditablePPTXTab(File file, JTabbedPane tabbedPane, Map<Component, String> tabTypeMap) throws IOException {
        JPanel panel = new PptxEditorPanel(file);
        JScrollPane scroll = new JScrollPane(panel);
        tabbedPane.addTab(file.getName(), PptxEditorPanel.getTabIcon(), scroll);
        tabTypeMap.put(scroll, "pptx_editable");
    }

    private static void openTableFile(File file, JTabbedPane tabbedPane, Map<Component, String> tabTypeMap) throws IOException {
        try (FileInputStream fis = new FileInputStream(file);
             XSSFWorkbook workbook = new XSSFWorkbook(fis)) {

            XSSFSheet sheet = workbook.getSheetAt(0);
            int rows = sheet.getPhysicalNumberOfRows();
            int cols = sheet.getRow(0).getPhysicalNumberOfCells();

            Object[][] data = new Object[rows][cols];
            String[] columns = new String[cols];

            for (int c = 0; c < cols; c++) columns[c] = String.valueOf((char) ('A' + c));

            for (int r = 0; r < rows; r++) {
                Row row = sheet.getRow(r);
                if (row == null) continue;

                for (int c = 0; c < cols; c++) {
                    Cell cell = row.getCell(c);
                    String value = (cell != null) ? cell.toString() : "";
                    DefaultStyledDocument doc = new DefaultStyledDocument();
                    try { doc.insertString(0, value, null); } catch (Exception ignored) {}
                    data[r][c] = doc;
                }
            }

            JTable table = new JTable(data, columns);
            JScrollPane scroll = new JScrollPane(table);

            ImageIcon originalIcon = new ImageIcon(FileOpenHandler.class.getResource("/icons/excel.png"));
            Image scaled = originalIcon.getImage().getScaledInstance(16, 16, Image.SCALE_SMOOTH);
            Icon logoIcon = new ImageIcon(scaled);

            tabbedPane.addTab(file.getName(), logoIcon, scroll);
            tabTypeMap.put(scroll, "table");
            tabbedPane.setSelectedComponent(scroll);

        } catch (Exception e) {
            handleOpenError(new IOException("Failed to open Excel file: " + e.getMessage()));
        }
    }

    private static ImageIcon createIcon(String path) {
        java.net.URL imgURL = FileOpenHandler.class.getResource(path);
        if (imgURL != null) {
            return new ImageIcon(imgURL);
        } else {
            System.err.println("Couldn't find file: " + path);
            return null;
        }
    }

    private static void openImageFile(File file, JTabbedPane tabbedPane, Map<Component, String> tabTypeMap) throws IOException {
        try {
            ImageIcon icon = new ImageIcon(file.getAbsolutePath());
            JLabel label = new JLabel(icon);
            JScrollPane scroll = new JScrollPane(label);
            ImageIcon imageIcon = createIcon("/icons/image.png");

            tabbedPane.addTab(file.getName(), imageIcon, scroll);
            tabTypeMap.put(scroll, "image");
        } catch (Exception e) {
            openTextFile(file, tabbedPane, tabTypeMap);
        }
    }

    private static void openTextFile(File file, JTabbedPane tabbedPane, Map<Component, String> tabTypeMap) throws IOException {
        try (BufferedReader r = new BufferedReader(new FileReader(file))) {
            StringBuilder sb = new StringBuilder();
            String line;
            while ((line = r.readLine()) != null)
                sb.append(line).append("\n");
            JTextArea area = new JTextArea(sb.toString());
            JScrollPane scroll = new JScrollPane(area);
            tabbedPane.addTab(file.getName(), scroll);
            tabTypeMap.put(scroll, "text");
        }
    }

    private static void handleOpenError(IOException ex) {
        ex.printStackTrace();
        JOptionPane.showMessageDialog(null,
                "❌ IO Error: " + ex.getMessage(),
                "Open Error", JOptionPane.ERROR_MESSAGE);
    }
}
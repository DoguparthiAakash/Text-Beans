package com.example.notepadpp.media;

import java.awt.BorderLayout;
import java.awt.Dimension;
import java.awt.FlowLayout;
import java.awt.Image;
import java.awt.event.ActionEvent;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.io.File;
import java.io.IOException;

import javax.swing.Icon;
import javax.swing.ImageIcon;
import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTextField;
import javax.swing.SwingUtilities;
import javax.swing.SwingWorker;
import javax.swing.filechooser.FileNameExtensionFilter;

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.common.PDRectangle;
import org.apache.pdfbox.pdmodel.interactive.annotation.PDAnnotationText;
import org.apache.pdfbox.rendering.PDFRenderer;

public class pdfview extends JFrame {
    private PDDocument document;
    private PDFRenderer renderer;
    private JLabel pdfLabel;
    private File currentFile;
    private JButton saveBtn;
    private JButton addAnnotationBtn;
    private JTextField annotationField;
    private static pdfview currentInstance = null;
    private volatile boolean isClosing = false;

    public pdfview(File file) throws IOException {
        super("PDF View and Edit - " + file.getName());
        
        // Close previous instance if exists
        if (currentInstance != null && currentInstance != this) {
            currentInstance.safeClose();
        }
        currentInstance = this;
        
        this.currentFile = file;
        
        // Initialize UI components first
        initializeUI();
        
        // Load PDF in background thread
        loadPDFInBackground(file);
        
        setupWindowListener();
        
        pack();
        setLocationRelativeTo(null);
    }

    private void initializeUI() {
        pdfLabel = new JLabel("Loading PDF...", JLabel.CENTER);
        pdfLabel.setPreferredSize(new Dimension(600, 800));
        
        JScrollPane scrollPane = new JScrollPane(pdfLabel);
        scrollPane.setPreferredSize(new Dimension(600, 800));
        
        // Annotation panel
        JPanel annotationPanel = new JPanel(new FlowLayout());
        annotationField = new JTextField(20);
        addAnnotationBtn = new JButton("Add Annotation");
        addAnnotationBtn.setEnabled(false);
        
        annotationPanel.add(new JLabel("Annotation:"));
        annotationPanel.add(annotationField);
        annotationPanel.add(addAnnotationBtn);
        
        addAnnotationBtn.addActionListener((ActionEvent e) -> {
            if (!isClosing) {
                addAnnotation();
            }
        });
        
        // Save button
        saveBtn = new JButton("Save PDF");
        saveBtn.setEnabled(false);
        saveBtn.addActionListener((ActionEvent e) -> {
            if (!isClosing) {
                savePDF();
            }
        });
        
        JPanel bottomPanel = new JPanel(new FlowLayout());
        bottomPanel.add(saveBtn);
        
        getContentPane().setLayout(new BorderLayout());
        getContentPane().add(annotationPanel, BorderLayout.NORTH);
        getContentPane().add(scrollPane, BorderLayout.CENTER);
        getContentPane().add(bottomPanel, BorderLayout.SOUTH);

        // DO_NOTHING_ON_CLOSE to handle cleanup manually
        setDefaultCloseOperation(JFrame.DO_NOTHING_ON_CLOSE);
    }

    private void loadPDFInBackground(File file) {
        SwingWorker<Void, Void> worker = new SwingWorker<Void, Void>() {
            @Override
            protected Void doInBackground() throws Exception {
                try {
                    document = PDDocument.load(file);
                    renderer = new PDFRenderer(document);
                    return null;
                } catch (IOException ex) {
                    throw new RuntimeException("Error loading PDF: " + ex.getMessage(), ex);
                }
            }
            
            @Override
            protected void done() {
                if (isClosing) return;
                
                try {
                    get();
                    updatePDFView();
                    addAnnotationBtn.setEnabled(true);
                    saveBtn.setEnabled(true);
                } catch (Exception ex) {
                    JOptionPane.showMessageDialog(pdfview.this, 
                        "Error loading PDF: " + ex.getCause().getMessage(), 
                        "PDF Load Error", 
                        JOptionPane.ERROR_MESSAGE);
                    safeClose();
                }
            }
        };
        worker.execute();
    }

    private void updatePDFView() {
        if (isClosing) return;
        
        if (!SwingUtilities.isEventDispatchThread()) {
            SwingUtilities.invokeLater(this::updatePDFView);
            return;
        }
        
        try {
            if (document != null && renderer != null && document.getNumberOfPages() > 0) {
                Image image = renderer.renderImageWithDPI(0, 150);
                ImageIcon icon = new ImageIcon(image);
                pdfLabel.setIcon(icon);
                pdfLabel.setText("");
                
                pdfLabel.revalidate();
                pdfLabel.repaint();
            } else {
                pdfLabel.setText("No pages in PDF");
                pdfLabel.setIcon(null);
            }
        } catch (IOException ex) {
            if (!isClosing) {
                pdfLabel.setText("Error rendering PDF: " + ex.getMessage());
                pdfLabel.setIcon(null);
                JOptionPane.showMessageDialog(this, 
                    "Error rendering PDF: " + ex.getMessage(), 
                    "Render Error", 
                    JOptionPane.ERROR_MESSAGE);
            }
        } catch (Exception ex) {
            if (!isClosing) {
                pdfLabel.setText("Unexpected error: " + ex.getMessage());
                pdfLabel.setIcon(null);
            }
        }
    }

    private void addAnnotation() {
        if (isClosing) return;
        
        String annotationText = annotationField.getText();
        if (annotationText != null && !annotationText.trim().isEmpty()) {
            try {
                addAnnotationToPage(annotationText.trim());
                annotationField.setText("");
                updatePDFView();
                JOptionPane.showMessageDialog(this, "Annotation added successfully!");
            } catch (IOException ex) {
                JOptionPane.showMessageDialog(this, 
                    "Error adding annotation: " + ex.getMessage(), 
                    "Annotation Error", 
                    JOptionPane.ERROR_MESSAGE);
            }
        } else {
            JOptionPane.showMessageDialog(this, 
                "Please enter annotation text", 
                "Input Error", 
                JOptionPane.WARNING_MESSAGE);
        }
    }

    private void addAnnotationToPage(String text) throws IOException {
        if (document == null || document.getNumberOfPages() == 0) {
            throw new IOException("No PDF document loaded or no pages available");
        }
        
        PDPage page = document.getPage(0);
        PDAnnotationText annotation = new PDAnnotationText();
        annotation.setContents(text);
        
        PDRectangle pageSize = page.getMediaBox();
        PDRectangle rect = new PDRectangle(50, pageSize.getHeight() - 100, 200, 30);
        annotation.setRectangle(rect);
        
        page.getAnnotations().add(annotation);
    }

    private void savePDF() {
        if (isClosing || document == null) {
            if (!isClosing) {
                JOptionPane.showMessageDialog(this, 
                    "No PDF document to save", 
                    "Save Error", 
                    JOptionPane.ERROR_MESSAGE);
            }
            return;
        }
        
        try {
            document.save(currentFile);
            JOptionPane.showMessageDialog(this, 
                "PDF saved successfully!", 
                "Save Successful", 
                JOptionPane.INFORMATION_MESSAGE);
        } catch (IOException ex) {
            JOptionPane.showMessageDialog(this, 
                "Error saving PDF: " + ex.getMessage(), 
                "Save Error", 
                JOptionPane.ERROR_MESSAGE);
        }
    }

    private void setupWindowListener() {
        addWindowListener(new WindowAdapter() {
            @Override
            public void windowClosing(WindowEvent e) {
                safeClose();
            }
        });
    }

    private void safeClose() {
        if (isClosing) return;
        
        isClosing = true;
        
        // Disable all buttons to prevent user interaction during close
        SwingUtilities.invokeLater(() -> {
            if (addAnnotationBtn != null) addAnnotationBtn.setEnabled(false);
            if (saveBtn != null) saveBtn.setEnabled(false);
        });
        
        // Small delay to allow any pending UI operations to complete
        try {
            Thread.sleep(50);
        } catch (InterruptedException e) {
            Thread.currentThread().interrupt();
        }
        
        try {
            if (document != null) {
                document.close();
                document = null;
            }
        } catch (IOException ex) {
            // Ignore close errors
        } finally {
            if (currentInstance == this) {
                currentInstance = null;
            }
            
            SwingUtilities.invokeLater(() -> {
                dispose();
            });
        }
    }

    public static void main(String[] args) {
        try {
            javax.swing.UIManager.setLookAndFeel(javax.swing.UIManager.getSystemLookAndFeelClassName());
        } catch (Exception e) {
            // Use default look and feel
        }
        
        SwingUtilities.invokeLater(() -> {
            JFileChooser fileChooser = new JFileChooser();
            fileChooser.setDialogTitle("Select a PDF File");
            FileNameExtensionFilter filter = new FileNameExtensionFilter("PDF Files", "pdf");
            fileChooser.setFileFilter(filter);
            
            int result = fileChooser.showOpenDialog(null);
            if (result == JFileChooser.APPROVE_OPTION) {
                File pdfFile = fileChooser.getSelectedFile();
                if (pdfFile.exists() && pdfFile.canRead()) {
                    try {
                        pdfview viewer = new pdfview(pdfFile);
                        viewer.setVisible(true);
                    } catch (IOException ex) {
                        JOptionPane.showMessageDialog(null, 
                            "Error opening PDF: " + ex.getMessage(), 
                            "Open Error", 
                            JOptionPane.ERROR_MESSAGE);
                    }
                } else {
                    JOptionPane.showMessageDialog(null, 
                        "Cannot read the selected file", 
                        "File Error", 
                        JOptionPane.ERROR_MESSAGE);
                }
            } else {
                System.out.println("No file selected, exiting.");
                System.exit(0);
            }
        });
    }

    public static Icon getTabIcon() {
        try {
            ImageIcon originalIcon = new ImageIcon(pdfview.class.getResource("/icons/pdfview.png"));
            int fixedWidth = 16;
            int fixedHeight = 16;
            Image scaledImage = originalIcon.getImage().getScaledInstance(fixedWidth, fixedHeight, Image.SCALE_SMOOTH);
            return new ImageIcon(scaledImage);
        } catch (Exception e) {
            return new ImageIcon();
        }
    }
}
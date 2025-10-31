package com.example.notepadpp.media;

import java.awt.BorderLayout;
import java.awt.Color;
import java.awt.Component;
import java.awt.Dimension;
import java.awt.Font;
import java.awt.Graphics2D;
import java.awt.Image;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Collectors;

import javax.swing.BoxLayout;
import javax.swing.Icon;
import javax.swing.ImageIcon;
import javax.swing.JButton;
import javax.swing.JColorChooser;
import javax.swing.JFileChooser;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JSplitPane;
import javax.swing.JTextArea;
import javax.swing.JTextField;
import javax.swing.SwingConstants;
import javax.swing.UIManager;
import javax.swing.filechooser.FileNameExtensionFilter;

import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextShape;

public class PptxEditorPanel extends JPanel {
    private final File file;
    private XMLSlideShow ppt;
    private List<XSLFSlide> slides;
    private int currentSlideIndex = 0;
    
    // Each slide’s text is represented by a JTextArea.
    private final List<JTextArea> pptxSlides = new ArrayList<>();
    
    // Panels for managing slides and slide preview in the tab.
    private JPanel pptxPanel;
    private JPanel pptxPreviewPanel;
    
    // Components for the main PPTX editor (image preview and text editor).
    private final JLabel imageLabel;
    private final JTextArea textEditor;
    private final JLabel slideLabel;
    
    // Constructor – creates the editor panel from the given file.
    public PptxEditorPanel(File file) {
        this.file = file;
        setLayout(new BorderLayout());
        
        // Create the image label for slide preview.
        imageLabel = new JLabel();
        imageLabel.setHorizontalAlignment(SwingConstants.CENTER);
        
        // Create the text editor for slide text editing.
        textEditor = new JTextArea();
        textEditor.setFont(new Font("SansSerif", Font.PLAIN, 14));
        
        // Use a split pane for the preview image and text editor.
        JSplitPane split = new JSplitPane(JSplitPane.VERTICAL_SPLIT,
                          new JScrollPane(imageLabel), new JScrollPane(textEditor));
        split.setResizeWeight(0.7);
        add(split, BorderLayout.CENTER);
        
        // Create a control panel for navigating slides and saving.
        JPanel controls = new JPanel();
        JButton prev = new JButton("⬅ Prev");
        JButton next = new JButton("Next ➡");
        JButton save = new JButton("💾 Save");
        slideLabel = new JLabel();
        prev.addActionListener(e -> showSlide(currentSlideIndex - 1));
        next.addActionListener(e -> showSlide(currentSlideIndex + 1));
        save.addActionListener(this::saveSlide);
        controls.add(prev);
        controls.add(next);
        controls.add(slideLabel);
        controls.add(save);
        add(controls, BorderLayout.SOUTH);
        
        // Load existing slides if the file exists; otherwise create a new PPTX.
        loadSlides();
        if (!pptxSlides.isEmpty()) {
            showSlide(0);
        }
    }
    
    // Loads the PPTX file. If the file does not exist, creates a new PPTX.
    private void loadSlides() {
        if (!file.exists()) {
            createNewPPTX();
            return;
        }
        try (FileInputStream fis = new FileInputStream(file)) {
            ppt = new XMLSlideShow(fis);
            slides = ppt.getSlides();
            pptxSlides.clear();
            for (XSLFSlide slide : slides) {
                List<XSLFTextShape> textShapes = slide.getShapes().stream()
                        .filter(s -> s instanceof XSLFTextShape)
                        .map(s -> (XSLFTextShape) s)
                        .collect(Collectors.toList());
                StringBuilder sb = new StringBuilder();
                for (XSLFTextShape shape : textShapes) {
                    if (shape.getText() != null && !shape.getText().isBlank()) {
                        sb.append(shape.getText()).append("\n");
                    }
                }
                JTextArea area = new JTextArea(sb.toString().trim());
                area.setFont(new Font("SansSerif", Font.PLAIN, 14));
                pptxSlides.add(area);
            }
        } catch (IOException e) {
            showError("Failed to load PPTX", e);
        }
    }
    
    // Called when the PPTX file doesn't exist. Creates a new XMLSlideShow with one blank slide.
    private void createNewPPTX() {
        ppt = new XMLSlideShow();
        XSLFSlide slide = ppt.createSlide();
        slides = new ArrayList<>();
        slides.add(slide);
        pptxSlides.clear();
        JTextArea area = new JTextArea();
        area.setFont(new Font("SansSerif", Font.PLAIN, 14));
        pptxSlides.add(area);
    }
    
    // Displays a slide (by index) updating the text editor and the preview image.
    private void showSlide(int index) {
        if (index < 0 || index >= pptxSlides.size()) return;
        currentSlideIndex = index;
        JTextArea currentArea = pptxSlides.get(index);
        textEditor.setText(currentArea.getText());
        XSLFSlide slide = slides.get(index);
        Dimension pageSize = ppt.getPageSize();
        BufferedImage img = new BufferedImage(pageSize.width, pageSize.height, BufferedImage.TYPE_INT_ARGB);
        Graphics2D g = img.createGraphics();
        g.setPaint(Color.WHITE);
        g.fillRect(0, 0, pageSize.width, pageSize.height);
        slide.draw(g);
        g.dispose();
        imageLabel.setIcon(new ImageIcon(img));
        slideLabel.setText("Slide " + (index + 1) + " / " + pptxSlides.size());
    }
    
    // Saves the current slide text back into the PPTX file.
    private void saveSlide(java.awt.event.ActionEvent e) {
        try {
            JTextArea currentArea = pptxSlides.get(currentSlideIndex);
            currentArea.setText(textEditor.getText());
            // Update text shapes for each slide.
            for (int i = 0; i < pptxSlides.size(); i++) {
                JTextArea area = pptxSlides.get(i);
                XSLFSlide slide = slides.get(i);
                List<XSLFTextShape> textShapes = slide.getShapes().stream()
                        .filter(s -> s instanceof XSLFTextShape)
                        .map(s -> (XSLFTextShape) s)
                        .collect(Collectors.toList());
                String[] lines = area.getText().split("\n");
                for (int j = 0; j < textShapes.size() && j < lines.length; j++) {
                    textShapes.get(j).setText(lines[j]);
                }
            }
            try (FileOutputStream fos = new FileOutputStream(file)) {
                ppt.write(fos);
            }
            JOptionPane.showMessageDialog(this, "✅ PPTX saved successfully!");
        } catch (IOException ex) {
            showError("Failed to save PPTX", ex);
        }
    }
    
    // Displays an error message dialog.
    private void showError(String msg, Exception ex) {
        JOptionPane.showMessageDialog(this, msg + ":\n" + ex.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
        ex.printStackTrace();
    }
    
    // -----------------------------------------------------------------------------------
    // Methods for the PPTX editing tab (for creating new slides and extra editing features).
    
    // This method creates a new PPTX editing tab with slide preview and extra options.
    // In your main TextEditor you can add this panel as a tab.
    public void addPPTXTab() {
        pptxPanel = new JPanel();
        pptxPanel.setLayout(new BoxLayout(pptxPanel, BoxLayout.Y_AXIS));
        
        pptxPreviewPanel = new JPanel();
        pptxPreviewPanel.setLayout(new BoxLayout(pptxPreviewPanel, BoxLayout.Y_AXIS));
        JScrollPane previewScroll = new JScrollPane(pptxPreviewPanel);
        previewScroll.setPreferredSize(new Dimension(120, 400));
        
        JPanel mainPanel = new JPanel(new BorderLayout());
        JScrollPane slidesScroll = new JScrollPane(pptxPanel);
        
        mainPanel.add(previewScroll, BorderLayout.WEST);
        mainPanel.add(slidesScroll, BorderLayout.CENTER);
        
        // Original buttons panel.
        JPanel btnPanel = new JPanel();
        JButton addSlideBtn = new JButton("Add Slide");
        JButton removeSlideBtn = new JButton("Remove Slide");
        JButton addChartBtn = new JButton("Add Bar Chart");
        JButton addClipartBtn = new JButton("Add Clipart");
        JButton addBulletBtn = new JButton("Add Bullets");
        JButton changeBgBtn = new JButton("Change BG Color");
        JButton insertTextBoxBtn = new JButton("Insert TextBox");
        btnPanel.add(addSlideBtn);
        btnPanel.add(removeSlideBtn);
        btnPanel.add(addChartBtn);
        btnPanel.add(addClipartBtn);
        btnPanel.add(addBulletBtn);
        btnPanel.add(changeBgBtn);
        btnPanel.add(insertTextBoxBtn);
        
        // Extra editing features panel.
        JPanel extraBtnPanel = new JPanel();
        JButton setFontBtn = new JButton("Set Font");
        JButton textColorBtn = new JButton("Text Color");
        JButton duplicateSlideBtn = new JButton("Duplicate Slide");
        JButton clearSlideBtn = new JButton("Clear Slide");
        JButton insertImageBtn = new JButton("Insert Image");
        extraBtnPanel.add(setFontBtn);
        extraBtnPanel.add(textColorBtn);
        extraBtnPanel.add(duplicateSlideBtn);
        extraBtnPanel.add(clearSlideBtn);
        extraBtnPanel.add(insertImageBtn);
        
        // Combine the two button panels.
        JPanel topBtnPanel = new JPanel(new BorderLayout());
        topBtnPanel.add(btnPanel, BorderLayout.NORTH);
        topBtnPanel.add(extraBtnPanel, BorderLayout.SOUTH);
        
        mainPanel.add(topBtnPanel, BorderLayout.NORTH);
        
        addSlide(); // Create the first slide.
        
        // Set listeners for original features.
        addSlideBtn.addActionListener(e -> addSlide());
        removeSlideBtn.addActionListener(e -> removeCurrentSlide());
        addChartBtn.addActionListener(e -> addBarChartToCurrentSlide());
        addClipartBtn.addActionListener(e -> addClipartToCurrentSlide());
        addBulletBtn.addActionListener(e -> addBulletsToCurrentSlide());
        changeBgBtn.addActionListener(e -> changeBackgroundColor());
        insertTextBoxBtn.addActionListener(e -> insertTextBox());
        
        // Listeners for extra editing features.
        setFontBtn.addActionListener(e -> setSlideFont());
        textColorBtn.addActionListener(e -> setSlideTextColor());
        duplicateSlideBtn.addActionListener(e -> duplicateSlide());
        clearSlideBtn.addActionListener(e -> clearSlide());
        insertImageBtn.addActionListener(e -> insertImageToSlide());
        
        JScrollPane mainScroll = new JScrollPane(mainPanel);
        // In your main editor, add "mainScroll" as a new tab.
        // For example: tabbedPane.addTab("New PPTX", mainScroll);
    }
    
    // Adds a new slide to the PPTX editing tab.
    private void addSlide() {
        JTextArea area = new JTextArea();
        area.setFont(new Font("SansSerif", Font.PLAIN, 14));
        pptxSlides.add(area);
        JLabel lbl = new JLabel("Slide " + pptxSlides.size());
        pptxPanel.add(lbl);
        pptxPanel.add(new JScrollPane(area));
        updatePPTXPreview();
        pptxPanel.revalidate();
        pptxPanel.repaint();
    }
    
    // Removes the last slide while ensuring at least one slide remains.
    private void removeCurrentSlide() {
        if (pptxSlides.size() <= 1) {
            JOptionPane.showMessageDialog(this, "At least one slide is required");
            return;
        }
        int idx = pptxSlides.size() - 1;
        pptxPanel.remove(pptxPanel.getComponentCount() - 1); // Remove scroll pane.
        pptxPanel.remove(pptxPanel.getComponentCount() - 1); // Remove label.
        pptxSlides.remove(idx);
        updatePPTXPreview();
        pptxPanel.revalidate();
        pptxPanel.repaint();
    }
    
    // Updates the slide preview panel with slide labels.
   private void updatePPTXPreview() {
    pptxPreviewPanel.removeAll();
    Color previewColor = UIManager.getColor("Panel.background"); // Will adapt to dark/light mode
    for (int i = 0; i < pptxSlides.size(); i++) {
        JLabel label = new JLabel("Slide " + (i + 1));
        label.setPreferredSize(new Dimension(100, 40));
        label.setOpaque(true);
        // Instead of a fixed light blue, use the current panel background or a darker variant.
        label.setBackground(previewColor);  
        pptxPreviewPanel.add(label);
    }
    pptxPreviewPanel.revalidate();
    pptxPreviewPanel.repaint();
}
    // Appends a bar chart placeholder to the current slide.
    private void addBarChartToCurrentSlide() {
        if (pptxSlides.isEmpty()) return;
        JTextArea area = pptxSlides.get(pptxSlides.size() - 1);
        area.append("\n[Bar Chart]\nA: ###\nB: ######\nC: ##\n");
    }
    
    // Appends a clipart placeholder to the current slide.
    private void addClipartToCurrentSlide() {
        if (pptxSlides.isEmpty()) return;
        JTextArea area = pptxSlides.get(pptxSlides.size() - 1);
        area.append("\n[Clipart] 😀 🏖️ 🌳\n");
    }
    
    // Appends bullet points to the current slide.
    private void addBulletsToCurrentSlide() {
        if (pptxSlides.isEmpty()) return;
        JTextArea area = pptxSlides.get(pptxSlides.size() - 1);
        area.append("\n• Bullet Point 1\n• Bullet Point 2\n• Bullet Point 3\n");
    }
    
    // Changes the background color of the PPTX panel and its scrollable child components.
    private void changeBackgroundColor() {
        Color c = JColorChooser.showDialog(this, "Pick Slide Background Color", Color.WHITE);
        if (c != null) {
            pptxPanel.setBackground(c);
            Component[] comps = pptxPanel.getComponents();
            for (int i = 0; i < comps.length; i++) {
                Component comp = comps[i];
                if (comp instanceof JScrollPane) {
                    JScrollPane scroll = (JScrollPane) comp;
                    Component view = scroll.getViewport().getView();
                    if (view != null) {
                        view.setBackground(c);
                    }
                }
            }
        }
    }
    
    // Inserts a text box (JTextField) into the PPTX panel.
    private void insertTextBox() {
        JTextField textField = new JTextField("Enter text", 20);
        pptxPanel.add(textField);
        pptxPanel.revalidate();
        pptxPanel.repaint();
    }
    
    // --- Extra Editing Features ---
    
    // Changes the font of the current slide's text.
    private void setSlideFont() {
        if (pptxSlides.isEmpty()) return;
        JTextArea area = pptxSlides.get(pptxSlides.size() - 1);
        String fontName = JOptionPane.showInputDialog(this, "Enter font name:", area.getFont().getFamily());
        String sizeStr = JOptionPane.showInputDialog(this, "Enter font size:", area.getFont().getSize());
        if (fontName != null && sizeStr != null) {
            try {
                int fontSize = Integer.parseInt(sizeStr);
                area.setFont(new Font(fontName, Font.PLAIN, fontSize));
            } catch (NumberFormatException ex) {
                JOptionPane.showMessageDialog(this, "Invalid font size.");
            }
        }
    }
    
    // Changes the text color of the current slide.
    private void setSlideTextColor() {
        if (pptxSlides.isEmpty()) return;
        JTextArea area = pptxSlides.get(pptxSlides.size() - 1);
        Color chosen = JColorChooser.showDialog(this, "Choose text color", area.getForeground());
        if (chosen != null) {
            area.setForeground(chosen);
        }
    }
    
    // Duplicates the last slide.
    private void duplicateSlide() {
        if (pptxSlides.isEmpty()) return;
        JTextArea lastSlide = pptxSlides.get(pptxSlides.size() - 1);
        String text = lastSlide.getText();
        JTextArea newArea = new JTextArea(text);
        newArea.setFont(lastSlide.getFont());
        pptxSlides.add(newArea);
        JLabel lbl = new JLabel("Slide " + pptxSlides.size());
        pptxPanel.add(lbl);
        pptxPanel.add(new JScrollPane(newArea));
        updatePPTXPreview();
        pptxPanel.revalidate();
        pptxPanel.repaint();
    }
    
    // Clears the text on the current slide.
    private void clearSlide() {
        if (pptxSlides.isEmpty()) return;
        JTextArea area = pptxSlides.get(pptxSlides.size() - 1);
        int res = JOptionPane.showConfirmDialog(this, "Clear all text on this slide?", "Clear Slide", JOptionPane.YES_NO_OPTION);
        if (res == JOptionPane.YES_OPTION) {
            area.setText("");
        }
    }
    
    // Opens a file chooser and inserts an image placeholder into the current slide's text.
    private void insertImageToSlide() {
        if (pptxSlides.isEmpty()) return;
        JFileChooser fc = new JFileChooser();
        fc.setFileFilter(new FileNameExtensionFilter("Image Files", "jpg", "jpeg", "png", "gif", "bmp"));
        if (fc.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) {
            File imgFile = fc.getSelectedFile();
            JTextArea area = pptxSlides.get(pptxSlides.size() - 1);
            area.append("\n[Image: " + imgFile.getName() + "]");
        }
    }
     public static Icon getTabIcon() {
        ImageIcon originalIcon = new ImageIcon(PptxEditorPanel.class.getResource("/icons/powerpoint.png"));
        int fixedWidth = 16;
        int fixedHeight = 16;
        Image scaledImage = originalIcon.getImage().getScaledInstance(fixedWidth, fixedHeight, Image.SCALE_SMOOTH);
        return new ImageIcon(scaledImage);
    }
    // -----------------------------------------------------------------------------------
    // End of PptxEditorPanel class.
}
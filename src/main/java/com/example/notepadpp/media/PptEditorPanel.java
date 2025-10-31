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
public class PptEditorPanel extends JPanel {
    private final File file;
    private XMLSlideShow ppt;
    private List<XSLFSlide> slides;
    private int currentSlideIndex = 0;
    private final List<JTextArea> pptSlides = new ArrayList<>();
    private JPanel pptPanel;
    private JPanel pptPreviewPanel;
    private final JLabel imageLabel;
    private final JTextArea textEditor;
    private final JLabel slideLabel;
    public PptEditorPanel(File file) {
        this.file = file;
        setLayout(new BorderLayout());
        imageLabel = new JLabel();
        imageLabel.setHorizontalAlignment(SwingConstants.CENTER);
        textEditor = new JTextArea();
        textEditor.setFont(new Font("SansSerif", Font.PLAIN, 14));
        JSplitPane split = new JSplitPane(JSplitPane.VERTICAL_SPLIT,
                          new JScrollPane(imageLabel), new JScrollPane(textEditor));
        split.setResizeWeight(0.7);
        add(split, BorderLayout.CENTER);
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
        loadSlides();
        if (!pptSlides.isEmpty()) {
            showSlide(0);
        }
    }
    private void loadSlides() {
        if (!file.exists()) {
            createNewPPTX();
            return;
        }
        try (FileInputStream fis = new FileInputStream(file)) {
            ppt = new XMLSlideShow(fis);
            slides = ppt.getSlides();
            pptSlides.clear();
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
                pptSlides.add(area);
            }
        } catch (IOException e) {
            showError("Failed to load PPTX", e);
        }
    }
    private void createNewPPTX() {
        ppt = new XMLSlideShow();
        XSLFSlide slide = ppt.createSlide();
        slides = new ArrayList<>();
        slides.add(slide);
        pptSlides.clear();
        JTextArea area = new JTextArea();
        area.setFont(new Font("SansSerif", Font.PLAIN, 14));
        pptSlides.add(area);
    }    private void showSlide(int index) {
        if (index < 0 || index >= pptSlides.size()) return;
        currentSlideIndex = index;
        JTextArea currentArea = pptSlides.get(index);
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
        slideLabel.setText("Slide " + (index + 1) + " / " + pptSlides.size());
    }
    private void saveSlide(java.awt.event.ActionEvent e) {
        try {
            JTextArea currentArea = pptSlides.get(currentSlideIndex);
            currentArea.setText(textEditor.getText());
            for (int i = 0; i < pptSlides.size(); i++) {
                JTextArea area = pptSlides.get(i);
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
    private void showError(String msg, Exception ex) {
        JOptionPane.showMessageDialog(this, msg + ":\n" + ex.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
        ex.printStackTrace();
    }    public void addPPTXTab() {
        pptPanel = new JPanel();
        pptPanel.setLayout(new BoxLayout(pptPanel, BoxLayout.Y_AXIS));
        pptPreviewPanel = new JPanel();
        pptPreviewPanel.setLayout(new BoxLayout(pptPreviewPanel, BoxLayout.Y_AXIS));
        JScrollPane previewScroll = new JScrollPane(pptPreviewPanel);
        previewScroll.setPreferredSize(new Dimension(120, 400));
        JPanel mainPanel = new JPanel(new BorderLayout());
        JScrollPane slidesScroll = new JScrollPane(pptPanel);
        mainPanel.add(previewScroll, BorderLayout.WEST);
        mainPanel.add(slidesScroll, BorderLayout.CENTER);
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
        JPanel topBtnPanel = new JPanel(new BorderLayout());
        topBtnPanel.add(btnPanel, BorderLayout.NORTH);
        topBtnPanel.add(extraBtnPanel, BorderLayout.SOUTH);
        mainPanel.add(topBtnPanel, BorderLayout.NORTH);
        addSlide(); 
        addSlideBtn.addActionListener(e -> addSlide());
        removeSlideBtn.addActionListener(e -> removeCurrentSlide());
        addChartBtn.addActionListener(e -> addBarChartToCurrentSlide());
        addClipartBtn.addActionListener(e -> addClipartToCurrentSlide());
        addBulletBtn.addActionListener(e -> addBulletsToCurrentSlide());
        changeBgBtn.addActionListener(e -> changeBackgroundColor());
        insertTextBoxBtn.addActionListener(e -> insertTextBox());
        setFontBtn.addActionListener(e -> setSlideFont());
        textColorBtn.addActionListener(e -> setSlideTextColor());
        duplicateSlideBtn.addActionListener(e -> duplicateSlide());
        clearSlideBtn.addActionListener(e -> clearSlide());
        insertImageBtn.addActionListener(e -> insertImageToSlide());
        
        JScrollPane mainScroll = new JScrollPane(mainPanel);
    }
    
    // Adds a new slide to the PPTX editing tab.
    private void addSlide() {
        JTextArea area = new JTextArea();
        area.setFont(new Font("SansSerif", Font.PLAIN, 14));
        pptSlides.add(area);
        JLabel lbl = new JLabel("Slide " + pptSlides.size());
        pptPanel.add(lbl);
        pptPanel.add(new JScrollPane(area));
        updatePPTXPreview();
        pptPanel.revalidate();
        pptPanel.repaint();
    }
    
    // Removes the last slide while ensuring at least one slide remains.
    private void removeCurrentSlide() {
        if (pptSlides.size() <= 1) {
            JOptionPane.showMessageDialog(this, "At least one slide is required");
            return;
        }
        int idx = pptSlides.size() - 1;
        pptPanel.remove(pptPanel.getComponentCount() - 1); // Remove scroll pane.
        pptPanel.remove(pptPanel.getComponentCount() - 1); // Remove label.
        pptSlides.remove(idx);
        updatePPTXPreview();
        pptPanel.revalidate();
        pptPanel.repaint();
    }
    
    // Updates the slide preview panel with slide labels.
   private void updatePPTXPreview() {
    pptPreviewPanel.removeAll();
    Color previewColor = UIManager.getColor("Panel.background"); // Will adapt to dark/light mode
    for (int i = 0; i < pptSlides.size(); i++) {
        JLabel label = new JLabel("Slide " + (i + 1));
        label.setPreferredSize(new Dimension(100, 40));
        label.setOpaque(true);
        // Instead of a fixed light blue, use the current panel background or a darker variant.
        label.setBackground(previewColor);  
        pptPreviewPanel.add(label);
    }
    pptPreviewPanel.revalidate();
    pptPreviewPanel.repaint();
}
    // Appends a bar chart placeholder to the current slide.
    private void addBarChartToCurrentSlide() {
        if (pptSlides.isEmpty()) return;
        JTextArea area = pptSlides.get(pptSlides.size() - 1);
        area.append("\n[Bar Chart]\nA: ###\nB: ######\nC: ##\n");
    }
    
    // Appends a clipart placeholder to the current slide.
    private void addClipartToCurrentSlide() {
        if (pptSlides.isEmpty()) return;
        JTextArea area = pptSlides.get(pptSlides.size() - 1);
        area.append("\n[Clipart] 😀 🏖️ 🌳\n");
    }
    
    // Appends bullet points to the current slide.
    private void addBulletsToCurrentSlide() {
        if (pptSlides.isEmpty()) return;
        JTextArea area = pptSlides.get(pptSlides.size() - 1);
        area.append("\n• Bullet Point 1\n• Bullet Point 2\n• Bullet Point 3\n");
    }
    
    // Changes the background color of the PPTX panel and its scrollable child components.
    private void changeBackgroundColor() {
        Color c = JColorChooser.showDialog(this, "Pick Slide Background Color", Color.WHITE);
        if (c != null) {
            pptPanel.setBackground(c);
            Component[] comps = pptPanel.getComponents();
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
        pptPanel.add(textField);
        pptPanel.revalidate();
        pptPanel.repaint();
    }
    
    // --- Extra Editing Features ---
    
    // Changes the font of the current slide's text.
    private void setSlideFont() {
        if (pptSlides.isEmpty()) return;
        JTextArea area = pptSlides.get(pptSlides.size() - 1);
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
        if (pptSlides.isEmpty()) return;
        JTextArea area = pptSlides.get(pptSlides.size() - 1);
        Color chosen = JColorChooser.showDialog(this, "Choose text color", area.getForeground());
        if (chosen != null) {
            area.setForeground(chosen);
        }
    }
    
    // Duplicates the last slide.
    private void duplicateSlide() {
        if (pptSlides.isEmpty()) return;
        JTextArea lastSlide = pptSlides.get(pptSlides.size() - 1);
        String text = lastSlide.getText();
        JTextArea newArea = new JTextArea(text);
        newArea.setFont(lastSlide.getFont());
        pptSlides.add(newArea);
        JLabel lbl = new JLabel("Slide " + pptSlides.size());
        pptPanel.add(lbl);
        pptPanel.add(new JScrollPane(newArea));
        updatePPTXPreview();
        pptPanel.revalidate();
        pptPanel.repaint();
    }
    
    // Clears the text on the current slide.
    private void clearSlide() {
        if (pptSlides.isEmpty()) return;
        JTextArea area = pptSlides.get(pptSlides.size() - 1);
        int res = JOptionPane.showConfirmDialog(this, "Clear all text on this slide?", "Clear Slide", JOptionPane.YES_NO_OPTION);
        if (res == JOptionPane.YES_OPTION) {
            area.setText("");
        }
    }
    
    // Opens a file chooser and inserts an image placeholder into the current slide's text.
    private void insertImageToSlide() {
        if (pptSlides.isEmpty()) return;
        JFileChooser fc = new JFileChooser();
        fc.setFileFilter(new FileNameExtensionFilter("Image Files", "jpg", "jpeg", "png", "gif", "bmp"));
        if (fc.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) {
            File imgFile = fc.getSelectedFile();
            JTextArea area = pptSlides.get(pptSlides.size() - 1);
            area.append("\n[Image: " + imgFile.getName() + "]");
        }
    }
    public static Icon getTabIcon() {
        ImageIcon originalIcon = new ImageIcon(PptEditorPanel.class.getResource("/icons/powerpoint.png"));
        int fixedWidth = 16;
        int fixedHeight = 16;
        Image scaledImage = originalIcon.getImage().getScaledInstance(fixedWidth, fixedHeight, Image.SCALE_SMOOTH);
        return new ImageIcon(scaledImage);
    }
    // -----------------------------------------------------------------------------------
    // End of PptxEditorPanel class.
}
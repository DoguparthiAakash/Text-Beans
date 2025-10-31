package com.example.notepadpp;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.ContainerAdapter;
import java.awt.event.ContainerEvent;
import java.awt.event.InputEvent;
import java.awt.event.KeyAdapter;
import java.awt.event.KeyEvent;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.awt.event.WindowAdapter;
import java.awt.event.WindowEvent;
import java.awt.geom.AffineTransform;
import java.awt.geom.Ellipse2D;
import java.awt.geom.Line2D;
import java.awt.geom.PathIterator;
import java.awt.geom.Point2D;
import java.awt.geom.Rectangle2D;
import java.awt.geom.RoundRectangle2D;
import java.awt.image.BufferedImage;
import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.FileWriter; // This is what you'll typically use
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.io.PipedInputStream;
import java.io.PipedOutputStream;
import java.io.PrintWriter;
import java.lang.reflect.Method;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.security.MessageDigest;
import java.security.NoSuchAlgorithmException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Base64;
import java.util.HashMap;
import java.util.HashSet;
import java.util.LinkedHashSet;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.Queue;
import java.util.Set;
import java.util.UUID;
import java.util.Vector;
import java.util.concurrent.CompletableFuture;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;

import javax.imageio.ImageIO;
import javax.script.ScriptEngine;
import javax.script.ScriptEngineManager;
import javax.swing.AbstractAction;
import javax.swing.AbstractCellEditor;
import javax.swing.ActionMap;
import javax.swing.BorderFactory;
import javax.swing.Box;
import javax.swing.BoxLayout;
import javax.swing.DefaultListModel;
import javax.swing.Icon;
import javax.swing.ImageIcon;
import javax.swing.InputMap;
import javax.swing.JButton;
import javax.swing.JCheckBox;
import javax.swing.JCheckBoxMenuItem;
import javax.swing.JColorChooser;
import javax.swing.JComboBox;
import javax.swing.JComponent;
import javax.swing.JDialog;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JList;
import javax.swing.JMenu;
import javax.swing.JMenuBar;
import javax.swing.JMenuItem;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JPopupMenu;
import javax.swing.JScrollBar;
import javax.swing.JScrollPane;
import javax.swing.JSlider;
import javax.swing.JSplitPane;
import javax.swing.JTabbedPane;
import javax.swing.JTable;
import javax.swing.JTextArea;
import javax.swing.JTextField;
import javax.swing.JTextPane;
import javax.swing.JToolBar;
import javax.swing.JViewport;
import javax.swing.KeyStroke;
import javax.swing.ListCellRenderer;
import javax.swing.LookAndFeel;
import javax.swing.SwingConstants;
import javax.swing.SwingUtilities;
import javax.swing.Timer;
import javax.swing.UIManager;
import javax.swing.UnsupportedLookAndFeelException;
import javax.swing.border.AbstractBorder;
import javax.swing.event.DocumentEvent;
import javax.swing.event.DocumentListener;
import javax.swing.event.TableModelEvent;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.table.AbstractTableModel;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.TableCellEditor;
import javax.swing.table.TableCellRenderer;
import javax.swing.table.TableModel;
import javax.swing.text.AbstractDocument;
import javax.swing.text.AttributeSet;
import javax.swing.text.BadLocationException;
import javax.swing.text.BoxView;
import javax.swing.text.ComponentView;
import javax.swing.text.DefaultHighlighter;
import javax.swing.text.DefaultStyledDocument;
import javax.swing.text.Document;
import javax.swing.text.Element;
import javax.swing.text.Highlighter;
import javax.swing.text.IconView;
import javax.swing.text.LabelView;
import javax.swing.text.MutableAttributeSet;
import javax.swing.text.ParagraphView;
import javax.swing.text.SimpleAttributeSet;
import javax.swing.text.Style;
import javax.swing.text.StyleConstants;
import javax.swing.text.StyleContext;
import javax.swing.text.StyledDocument;
import javax.swing.text.StyledEditorKit;
import javax.swing.text.View;
import javax.swing.text.ViewFactory;
import javax.swing.undo.UndoManager;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.fife.ui.rsyntaxtextarea.RSyntaxTextArea;
import org.fife.ui.rsyntaxtextarea.SyntaxConstants;
import org.fife.ui.rsyntaxtextarea.SyntaxScheme;
import org.fife.ui.rsyntaxtextarea.Theme;
import org.fife.ui.rsyntaxtextarea.Token;
import org.fife.ui.rtextarea.Gutter;
import org.fife.ui.rtextarea.RTextScrollPane;

import com.example.notepadpp.TextEditor.ClockPanel;
import com.example.notepadpp.TextEditor.DrawingPanel;
import com.example.notepadpp.TextEditor.DrawingPanel.ShapeInfo;
import com.example.notepadpp.TextEditor.DrawingPanel.TextShape;
import com.example.notepadpp.TextEditor.FadingPanel;
import com.example.notepadpp.TextEditor.StyledCellEditor;
import com.example.notepadpp.TextEditor.StyledCellRenderer;
import com.example.notepadpp.TextEditor.SyntaxTextPane;
import com.example.notepadpp.TextEditor.WrapEditorKit;
import com.example.notepadpp.TextEditor.WrapEditorKit.StyledViewFactory;
import com.example.notepadpp.TextEditor.WrapEditorKit.WrapLabelView;
import com.example.notepadpp.TextEditor.ZipViewer;
import com.example.notepadpp.ai.AIAssistantAddon;
import com.example.notepadpp.media.AudioPlayerPanel;
import com.example.notepadpp.media.CompilerService;
import com.example.notepadpp.media.IconUtils;
import com.example.notepadpp.media.ImageView;
import com.example.notepadpp.media.LanguageTemplates;
import com.example.notepadpp.media.TaskManager.TaskManager.src1.TaskManagerApp;
import com.example.notepadpp.media.VLCJVideoPlayerPanel;
import com.example.notepadpp.media.XlsxEditorPanel;
import com.formdev.flatlaf.FlatLightLaf;
import com.formdev.flatlaf.themes.FlatMacDarkLaf;
import com.formdev.flatlaf.themes.FlatMacLightLaf;
import com.google.gson.Gson;
import com.google.gson.GsonBuilder;
import com.google.zxing.BarcodeFormat;
import com.google.zxing.MultiFormatWriter;
import com.google.zxing.client.j2se.MatrixToImageWriter;
import com.google.zxing.common.BitMatrix;

import javafx.embed.swing.JFXPanel;
import javafx.scene.control.Tab;
import javafx.scene.control.TabPane;
import javafx.scene.control.TextArea;

public class TextEditor extends JFrame {
    static {
        // Enable custom window decorations (rounded corners, shadows, etc.)
        JFrame.setDefaultLookAndFeelDecorated(true);
        JDialog.setDefaultLookAndFeelDecorated(true);

        // Set FlatLaf arc properties globally BEFORE any Swing component is created
        UIManager.put("Component.arc", 20);
        UIManager.put("Button.arc", 20);
        UIManager.put("TextComponent.arc", 20);
        UIManager.put("ScrollBar.thumbArc", 16);

        // Set FlatLaf as the default look and feel (choose your preferred theme)
        try {
            FlatMacLightLaf.setup();
            UIManager.setLookAndFeel(new FlatMacLightLaf());
        } catch (UnsupportedLookAndFeelException e) {
            e.printStackTrace();
        }
    }

    private final JTabbedPane tabbedPane = new JTabbedPane();
    private JLabel statusBar = new JLabel("Ln 1, Col 1");
    private final FileTypeRegistry fileTypeRegistry = new FileTypeRegistry();
    private final Map<JTextPane, File> fileMap = new HashMap<>();
    private final Map<JTextPane, UndoManager> undoMap = new HashMap<>();
    private final Map<Component, String> tabTypeMap = new HashMap<>();
    private final java.util.List<Color> customColors = new ArrayList<>();
    private JCheckBoxMenuItem wordWrapItem = new JCheckBoxMenuItem("Word Wrap");
    private final java.util.List<JTextArea> pptxSlides = new ArrayList<>();

    private JScrollPane scrollPane;

    private JTextArea textArea;

    private JPanel pptxPanel = new JPanel();
    private JPanel pptxPreviewPanel = new JPanel();
    // Declare at the top of your class, outside of any method
    private JTextField searchField = new JTextField(20);

    private JPanel searchPanel;
    private JButton nextButton, prevButton;
    private java.util.List<Integer> matchIndices = new ArrayList<>();
    private int currentMatch = -1;
    // your TaskPanel class
    private boolean tasksVisible = true;

    private TabPane editorTabPane;

    private TextArea consoleArea;
// Disable all FlatLaf animations

    // ADD THIS NEW LINE HERE
    private FadingPanel rootPanel;
 private int dragTabIndex = -1;
    // --------- New settings / autosave / recent files fields ---------
    private final File settingsDir = new File("Data/settings");
    private final File settingsFile = new File(settingsDir, "app.properties");
    private boolean autosaveEnabled = false;
    private int autosaveIntervalSec = 300; // default 5 minutes
    private Timer autosaveTimer;
    private JMenu recentFilesMenu;
    private final java.util.List<String> recentFiles = new ArrayList<>();
    private JCheckBoxMenuItem showLineNumbersItem = new JCheckBoxMenuItem("Show Line Numbers", true);
    // ---------------------------------------------------------------

    public TextEditor() {
        super("Text Beans");
        initUI();
        searchField = new JTextField();
        rootPanel = new FadingPanel(new BorderLayout());
        setContentPane(rootPanel);
        setSize(900, 700);
        setLocationRelativeTo(null);
        setDefaultCloseOperation(DO_NOTHING_ON_CLOSE);

        setSize(900, 700);
        setLocationRelativeTo(null);

        setJMenuBar(createMenuBar());
        rootPanel.add(tabbedPane, BorderLayout.CENTER);
        rootPanel.add(statusBar, BorderLayout.SOUTH);
        // Create the tabbed pane.

        // Add a container listener that checks if a removed component (or its view if
        // wrapped)
        // is an audio/video panel, and disposes of it.
// --- Initialize Tabbed Pane ---
tabbedPane.setTabLayoutPolicy(JTabbedPane.SCROLL_TAB_LAYOUT);

        // Container listener (your cleanup code)
        tabbedPane.addContainerListener(new ContainerAdapter() {
            @Override
            public void componentRemoved(ContainerEvent e) {
                Component comp = e.getChild();
                // cleanup media resources if needed (same logic you had)
                if (comp instanceof JScrollPane) {
                    Component view = ((JScrollPane) comp).getViewport().getView();
                    if (view instanceof AudioPlayerPanel) {
                        ((AudioPlayerPanel) view).disposeAudio();
                    } else if (view instanceof VLCJVideoPlayerPanel) {
                        ((VLCJVideoPlayerPanel) view).disposePlayer();
                    }
                } else {
                    if (comp instanceof AudioPlayerPanel) {
                        ((AudioPlayerPanel) comp).disposeAudio();
                    } else if (comp instanceof VLCJVideoPlayerPanel) {
                        ((VLCJVideoPlayerPanel) comp).disposePlayer();
                    }
                }
            }
        });

        // --- Mouse listeners for drag-reorder (placed here, allowed) ---
        tabbedPane.addMouseListener(new MouseAdapter() {
    @Override
    public void mousePressed(MouseEvent e) {
        dragTabIndex = tabbedPane.indexAtLocation(e.getX(), e.getY());
    }

    @Override
    public void mouseReleased(MouseEvent e) {
        dragTabIndex = -1; // Reset after drag
    }
});


        tabbedPane.addMouseMotionListener(new MouseAdapter() {
            @Override
            public void mouseDragged(MouseEvent e) {
                int targetIndex = tabbedPane.indexAtLocation(e.getX(), e.getY());
                if (dragTabIndex != -1 && targetIndex != -1 && dragTabIndex != targetIndex) {
                    Component comp = tabbedPane.getComponentAt(dragTabIndex);
                    String title = tabbedPane.getTitleAt(dragTabIndex);
                    Icon icon = tabbedPane.getIconAt(dragTabIndex);
                    Component tabComponent = tabbedPane.getTabComponentAt(dragTabIndex);

                    tabbedPane.remove(dragTabIndex);
                    tabbedPane.insertTab(title, icon, comp, null, targetIndex);
                    tabbedPane.setTabComponentAt(targetIndex, tabComponent);
                    tabbedPane.setSelectedIndex(targetIndex);

                    dragTabIndex = targetIndex;
                }
            }
        });
        // Add to frame
        add(tabbedPane, BorderLayout.CENTER);

        // Status bar
        statusBar = new JLabel("Ln 1, Col 1");
        add(statusBar, BorderLayout.SOUTH);

        // Window close handling (keep your confirmSaveAll and cleanup)
        addWindowListener(new WindowAdapter() {
            @Override
            public void windowClosing(WindowEvent e) {
                if (confirmSaveAll()) {
                    // dispose media resources for all tabs (same logic as before)
                    for (int i = 0; i < tabbedPane.getTabCount(); i++) {
                        Component comp = tabbedPane.getComponentAt(i);
                        if (comp instanceof JScrollPane) {
                            Component view = ((JScrollPane) comp).getViewport().getView();
                            if (view instanceof AudioPlayerPanel) {
                                ((AudioPlayerPanel) view).disposeAudio();
                            } else if (view instanceof VLCJVideoPlayerPanel) {
                                ((VLCJVideoPlayerPanel) view).disposePlayer();
                            }
                        } else {
                            if (comp instanceof AudioPlayerPanel) {
                                ((AudioPlayerPanel) comp).disposeAudio();
                            } else if (comp instanceof VLCJVideoPlayerPanel) {
                                ((VLCJVideoPlayerPanel) comp).disposePlayer();
                            }
                        }
                    }
                    try {
                        Thread.sleep(800); // small pause for cleanup
                    } catch (InterruptedException ex) {
                        ex.printStackTrace();
                    }
                    System.exit(0);
                }
            }
        });

        // Create first blank tab
        ImageIcon textIcon = loadAndScaleCloseIcon("newtab.png", 32, 32);
        addTabWithCloseButton("Untitled", new JScrollPane(new JTextArea()), textIcon);
        

        setVisible(true);
    }

    // <-- PLACE THIS METHOD AT CLASS LEVEL (NOT INSIDE THE CONSTRUCTOR) -->
    private void addTabWithCloseButton(String title, Component component, Icon icon) {
    tabbedPane.addTab(title, component); // icon here is ignored, but tab is created
    int index = tabbedPane.indexOfComponent(component);

    // Create custom header panel
    JPanel tabPanel = new JPanel(new FlowLayout(FlowLayout.LEFT, 0, 0));
    tabPanel.setOpaque(false);

    // Create label for title + icon
    JLabel titleLabel = new JLabel(title);
    if (icon != null) {
        // Resize icon to 16x16 if needed
        Image img = ((ImageIcon) icon).getImage().getScaledInstance(18, 18, Image.SCALE_SMOOTH);
        titleLabel.setIcon(new ImageIcon(img));
        titleLabel.setIconTextGap(5); // space between icon and text
    }

    // Close button
  JButton closeButton = new JButton();
try {
    // Use the loadAndScaleCloseIcon method for closetab.png
    ImageIcon closeIcon = loadAndScaleCloseIcon("closetab.png", 16, 16);
    
    if (closeIcon != null) {
        closeButton.setIcon(closeIcon);
        System.out.println("✓ Successfully loaded closetab.png");
    } else {
        // Fallback: create a simple X icon programmatically
        closeButton.setText("×");
        closeButton.setFont(new Font("Arial", Font.BOLD, 12));
        System.out.println("⚠ Using fallback close button - closetab.png not available");
    }
} catch (Exception e) {
    // Fallback: create a simple X icon programmatically
    closeButton.setText("×");
    closeButton.setFont(new Font("Arial", Font.BOLD, 12));
    System.err.println("✗ Error loading closetab.png: " + e.getMessage());
}

closeButton.setPreferredSize(new Dimension(20, 20));
closeButton.setFocusable(false);
closeButton.setBorder(BorderFactory.createEmptyBorder());
closeButton.setContentAreaFilled(false);

// Optional: Add hover effects like in the other implementations
closeButton.addMouseListener(new MouseAdapter() {
    @Override
    public void mouseEntered(MouseEvent e) {
        closeButton.setBackground(new Color(255, 100, 100, 80));
        closeButton.setContentAreaFilled(true);
    }
    
    @Override
    public void mouseExited(MouseEvent e) {
        closeButton.setBackground(null);
        closeButton.setContentAreaFilled(false);
    }
});

    // Hover effect
    closeButton.addMouseListener(new MouseAdapter() {
        @Override
        public void mouseEntered(MouseEvent e) { closeButton.setForeground(Color.RED); }
        @Override
        public void mouseExited(MouseEvent e) { closeButton.setForeground(Color.GRAY); }
    });

    // Close tab on click
    closeButton.addActionListener(e -> {
        int tabIndex = tabbedPane.indexOfTabComponent(tabPanel);
        if (tabIndex != -1) {
            Component comp = tabbedPane.getComponentAt(tabIndex);
            // Cleanup existing media components
            if (comp instanceof JScrollPane) {
                Component view = ((JScrollPane) comp).getViewport().getView();
                if (view instanceof AudioPlayerPanel) ((AudioPlayerPanel) view).disposeAudio();
                if (view instanceof VLCJVideoPlayerPanel) ((VLCJVideoPlayerPanel) view).disposePlayer();
            } else {
                if (comp instanceof AudioPlayerPanel) ((AudioPlayerPanel) comp).disposeAudio();
                if (comp instanceof VLCJVideoPlayerPanel) ((VLCJVideoPlayerPanel) comp).disposePlayer();
            }
            tabbedPane.remove(tabIndex);
        }
    });

    tabPanel.add(titleLabel);
    tabPanel.add(Box.createHorizontalStrut(5));
    tabPanel.add(closeButton);
    tabPanel.setBorder(BorderFactory.createEmptyBorder(2, 2, 0, 0));

    // Set the custom header
    tabbedPane.setTabComponentAt(index, tabPanel);
}


    private void initializeAnimations() {
        // Initialize tab switching animations
        addAnimatedTabSwitching();

        // Apply smooth scrolling to existing scroll panes
        for (int i = 0; i < tabbedPane.getTabCount(); i++) {
            Component comp = tabbedPane.getComponentAt(i);
            if (comp instanceof JScrollPane) {
                addSmoothScrolling((JScrollPane) comp);
            }
        }

        // Apply button animations to main menu buttons
        setupMenuButtonAnimations();
    }

    private void setupMenuButtonAnimations() {
        // Apply animations to toolbar buttons if any
        JMenuBar menuBar = getJMenuBar();
        if (menuBar != null) {
            for (Component comp : menuBar.getComponents()) {
                if (comp instanceof JButton) {
                    setupAnimatedButton((JButton) comp);
                }
            }
        }
    }

    /**
     * When closing an individual tab, call this method.
     */
    // Define a variable to store the last update time.
    private long lastStatusUpdate = 0;
    private final long updateInterval = 200; // Update every 200 milliseconds

    private void setupMenuAnimations() {
        // Apply animations to all buttons in the menu bar
        Component[] components = getJMenuBar().getComponents();
        for (Component comp : components) {
            if (comp instanceof JButton) {
                setupAnimatedButton((JButton) comp);
            }
        }
    }

    public void updateStatusBar(String newText) {
        long now = System.currentTimeMillis();
        if (now - lastStatusUpdate >= updateInterval) {
            SwingUtilities.invokeLater(() -> {
                statusBar.setText(newText);
            });
            lastStatusUpdate = now;
        }
    }

    public static void fadeInWindow(Window window, int durationMs) {
        window.setOpacity(0f);
        window.setVisible(true);
        Timer timer = new Timer(15, null);
        timer.addActionListener(e -> {
            float opacity = window.getOpacity() + 0.05f;
            if (opacity >= 1f) {
                window.setOpacity(1f);
                timer.stop();
            } else {
                window.setOpacity(opacity);
            }
        });
        timer.start();
    }

    // Fade out any JDialog or JFrame
    public static void fadeOutWindow(Window window, int durationMs) {
        Timer timer = new Timer(15, null);
        timer.addActionListener(new ActionListener() {
            float opacity = 1f;

            public void actionPerformed(ActionEvent e) {
                opacity -= 0.05f;
                if (opacity <= 0f) {
                    opacity = 0f;
                    timer.stop();
                    window.dispose();
                }
                window.setOpacity(opacity);
            }
        });
        timer.start();
    }

    // Slide in a panel from the right
    public static void slideInPanel(JPanel panel, int fromX, int toX, int y, int durationMs) {
        panel.setLocation(fromX, y);
        panel.setVisible(true);
        int steps = 30;
        int delay = durationMs / steps;
        Timer timer = new Timer(delay, null);
        timer.addActionListener(new ActionListener() {
            int step = 0;

            public void actionPerformed(ActionEvent e) {
                int x = fromX + (toX - fromX) * step / steps;
                panel.setLocation(x, y);
                step++;
                if (step > steps)
                    timer.stop();
            }
        });
        timer.start();
    }

    // Shake a component horizontally (e.g., for error feedback)
    public static void shakeComponent(Component comp) {
        Point original = comp.getLocation();
        int[] dx = { 0, -8, 8, -6, 6, -4, 4, -2, 2, 0 };
        final int[] i = { 0 };
        Timer timer = new Timer(15, e -> {
            if (i[0] < dx.length) {
                comp.setLocation(original.x + dx[i[0]], original.y);
                i[0]++;
            } else {
                comp.setLocation(original);
                ((Timer) e.getSource()).stop();
            }
        });
        timer.start();
    }

    // Scale (zoom) animation for dialogs
    public static void scaleInDialog(JDialog dialog, int durationMs) {
        dialog.setVisible(true);
        int steps = 20;
        int delay = durationMs / steps;
        Timer timer = new Timer(delay, null);
        timer.addActionListener(new ActionListener() {
            int step = 0;

            public void actionPerformed(ActionEvent e) {
                float scale = 0.7f + 0.3f * step / steps;
                dialog.setSize((int) (dialog.getWidth() * scale), (int) (dialog.getHeight() * scale));
                dialog.setLocationRelativeTo(null);
                step++;
                if (step > steps)
                    timer.stop();
            }
        });
        timer.start();
    }

    public void closeFile(Component panel) {
        if (panel instanceof AudioPlayerPanel) {
            ((AudioPlayerPanel) panel).disposeAudio();
        } else if (panel instanceof VLCJVideoPlayerPanel) {
            ((VLCJVideoPlayerPanel) panel).disposePlayer();
        }
        tabbedPane.remove(panel);
    }

    private JMenuBar createMenuBar() {
        JMenuBar bar = new JMenuBar();
        JMenuItem base64EncodeItem = new JMenuItem("Base64 Text Encode", loadIcon("/icons/lock.png", 16, 16));
        JMenuItem base64DecodeItem = new JMenuItem("Base64 Text Decode", loadIcon("/icons/key.png", 16, 16));

        JMenu file = new JMenu("File");

        JMenu browserMenu = new JMenu("Browser (Requires Internet Connection)");
        browserMenu.setIcon(loadIcon("/icons/browsermenu.png", 16, 16));

        JMenuItem openBrowser = new JMenuItem("Open Browser", loadIcon("/icons/browser.png", 16, 16));
        openBrowser.setAccelerator(KeyStroke.getKeyStroke(KeyEvent.VK_G, InputEvent.CTRL_DOWN_MASK));
        openBrowser.addActionListener(e -> {
            com.example.notepadpp.browser.BrowserApp.launchBrowser();
        });
        // Add to any menu (e.g., Tools)

        // 2) Bind Ctrl+G

        // 4) Add the item into the menu

        browserMenu.add(openBrowser);

        JMenu edit = new JMenu("Edit");

        JMenu view = new JMenu("View");

        JMenu tools = new JMenu("Tools");

        JMenu help = new JMenu("Help");

        JMenuItem jsonFormatter = new JMenuItem("JSON Formatter", loadIcon("/icons/json.png", 16, 16));
        jsonFormatter.addActionListener(e -> formatJSON());

        JMenuItem minifier = new JMenuItem("Minify JSON/JS/CSS/HTML", loadIcon("/icons/minify.png", 16, 16));
        minifier.addActionListener(e -> minifyText());

        JMenuItem baseConverter = new JMenuItem("Base Converter", loadIcon("/icons/baseconvert.png", 16, 16));
        baseConverter.addActionListener(e -> showBaseConverter());

        JMenuItem regexTester = new JMenuItem("Regex Tester", loadIcon("/icons/regex-logo.png", 16, 16));
        regexTester.addActionListener(e -> showRegexTester());

        JMenuItem loremIpsum = new JMenuItem("Lorem Ipsum Generator", loadIcon("/icons/l.png", 16, 16));
        loremIpsum.addActionListener(e -> textArea.append("Lorem ipsum dolor sit amet, consectetur adipiscing elit."));

        JMenuItem qrGenerator = new JMenuItem("QR Code Generator", loadIcon("/icons/qricon.png", 16, 16));
        qrGenerator.addActionListener(e -> generateQRCode());

        JMenuItem uuidGenerator = new JMenuItem("UUID Generator", loadIcon("/icons/uuid.png", 16, 16));
        uuidGenerator.addActionListener(e -> textArea.append(UUID.randomUUID().toString() + "\n"));

        JMenuItem hashChecker = new JMenuItem("File Hash Checker", loadIcon("/icons/hash.png", 16, 16));
        hashChecker.addActionListener(e -> checkFileHash());

        JMenuItem javaRunner = new JMenuItem("Run Java Snippet", loadIcon("/icons/java.png", 16, 16));
        javaRunner.addActionListener(e -> runJavaSnippet());

        JMenuItem exportQR = new JMenuItem("Export as QR Code", loadIcon("/icons/qricon.png", 16, 16));
        exportQR.addActionListener(e -> exportTextAsQRCode());
        JMenu tasksMenu = new JMenu("Tasks");
        tasksMenu.setIcon(loadIcon("/icons/tasks.png", 16, 16));

        JMenuItem openTaskManager = new JMenuItem("Task Manager", loadIcon("/icons/taskmanager.png", 16, 16));
        openTaskManager.addActionListener(e -> {
            TaskManagerApp tasksVisible = new TaskManagerApp(); // ← pass your text editor
            tasksVisible.setVisible(true);
        });

        tasksMenu.add(openTaskManager);
        // or add to an existing "Tools" menu

        // Add to your menu bar under "Tools" or "Tasks"
        // Create the "Run" menu item
        JMenuItem runItem = new JMenuItem("Run");
        runItem.setIcon(loadIcon("/icons/run.png", 16, 16));

        // Hook up the action
        runItem.addActionListener(e -> {
            try {
                runActiveTabCode(); // your method to compile/run
            } catch (Exception ex) {
                JOptionPane.showMessageDialog(
                        null,
                        "Error executing code:\n" + ex.getMessage(),
                        "Execution Error",
                        JOptionPane.ERROR_MESSAGE);
            }
        });

        // Assign Ctrl+R as keyboard accelerator (cross-platform safe)
        runItem.setAccelerator(
                KeyStroke.getKeyStroke(KeyEvent.VK_R, Toolkit.getDefaultToolkit().getMenuShortcutKeyMaskEx()));

        // Add to menu
        

        JMenuItem pptxTab = new JMenuItem("New Presentation Tab", loadIcon("/icons/powerpoint.png", 16, 16));
        JMenuItem currencyConverter = new JMenuItem("Currency Converter",
                loadIcon("/icons/currencyconvert.png", 16, 16));
        tools.add(currencyConverter);
        currencyConverter.addActionListener(e -> showCurrencyConverterPopup());
        JMenuItem insertImage = new JMenuItem("Insert Image", loadIcon("/icons/imageview.png", 16, 16));
        edit.add(insertImage);

        insertImage.addActionListener(e -> {
            JTextPane pane = getCurrentTextPane();
            if (pane == null) {
                JOptionPane.showMessageDialog(this, "Please open or create a text tab first.", "No Text Tab Selected",
                        JOptionPane.WARNING_MESSAGE);
                return;
            }

            JFileChooser fc = new JFileChooser();
            fc.setFileFilter(new FileNameExtensionFilter("Image Files", "jpg", "jpeg", "png", "gif", "bmp"));
            fc.setAcceptAllFileFilterUsed(false);

            if (fc.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) {
                File imgFile = fc.getSelectedFile();
                try {
                    // Use ImageIO.read() instead of ImageIcon constructor
                    BufferedImage image = ImageIO.read(imgFile);

                    if (image == null) {
                        JOptionPane.showMessageDialog(this,
                                "Could not load image from file: " + imgFile.getName() +
                                        "\nThe file may be corrupted or in an unsupported format.",
                                "Image Loading Error", JOptionPane.ERROR_MESSAGE);
                        return;
                    }

                    ImageIcon icon = new ImageIcon(image);

                    // Optional: open image editor dialog here before inserting
                    ImageIcon editedIcon = openImageEditorDialog(icon);

                    // Insert the icon into the text pane
                    pane.insertIcon(editedIcon);

                } catch (Exception ex) {
                    JOptionPane.showMessageDialog(this,
                            "Error inserting image:" + ex.getMessage(),
                            "Error", JOptionPane.ERROR_MESSAGE);
                    ex.printStackTrace(); // Add this for debugging
                }
            }
        });

        JMenuItem newTab = new JMenuItem("New tab", loadIcon("/icons/newtab.png", 16, 16));

        JMenuItem save = new JMenuItem("Save", loadIcon("/icons/save.png", 16, 16));
        JMenuItem saveAs = new JMenuItem("Save As", loadIcon("/icons/saveas.png", 16, 16));
        JMenuItem close = new JMenuItem("Close Tab", loadIcon("/icons/close.png", 16, 16));
        JMenuItem exit = new JMenuItem("Exit", loadIcon("/icons/exit.png", 16, 16));
        JMenuItem print = new JMenuItem("Print", loadIcon("/icons/printer.png", 16, 16));
        JMenuItem exportHTML = new JMenuItem("Export as HTML", loadIcon("/icons/htmlexport.png", 16, 16));
        JMenuItem exportMD = new JMenuItem("Export as Markdown", loadIcon("/icons/markdownexporticon.png", 16, 16));
        JMenu addons = new JMenu("Addons");

        JMenu textas = new JMenu("Font");
        textas.setIcon(loadIcon("/icons/font.png", 16, 16));
        JMenu opened = new JMenu("Open");
        opened.setIcon(loadIcon("/icons/openfile.png", 16, 16));
        JMenu expertor = new JMenu("Export");
        expertor.setIcon(loadIcon("/icons/export.png", 16, 16));
        JMenu saver = new JMenu("Save Options");
        saver.setIcon(loadIcon("/icons/save.png", 16, 16));

        JMenu colours = new JMenu("Color Settings");
        colours.setIcon(loadIcon("/icons/colours.png", 16, 16));
        JMenuItem calculator = new JMenuItem("Calculator", loadIcon("/icons/calculator.png", 16, 16));
        JMenuItem open = new JMenuItem("Open File", loadIcon("/icons/openfile.png", 16, 16));

        JMenuItem notepadMini = new JMenuItem("Mini Notes", loadIcon("/icons/notes.png", 16, 16));
        JMenuItem stopwatch = new JMenuItem("Stopwatch", loadIcon("/icons/stopwatch.png", 16, 16));

        // Add to your JMenuBar

        JMenuItem cut = new JMenuItem("Cut", loadIcon("/icons/cut.png", 16, 16));
        JMenuItem copy = new JMenuItem("Copy", loadIcon("/icons/copy.png", 16, 16));
        JMenuItem paste = new JMenuItem("Paste", loadIcon("/icons/paste.png", 16, 16));
        JMenuItem undo = new JMenuItem("Undo", loadIcon("/icons/undo.png", 16, 16));
        JMenuItem redo = new JMenuItem("Redo", loadIcon("/icons/redo.png", 16, 16));
        JMenuItem find = new JMenuItem("Find", loadIcon("/icons/find.png", 16, 16));
        JMenuItem replace = new JMenuItem("Replace", loadIcon("/icons/baseconvert.png", 16, 16));
        JMenuItem goTo = new JMenuItem("Go To Line", loadIcon("/icons/goto.png", 16, 16));
        wordWrapItem = new JCheckBoxMenuItem("Word Wrap", loadIcon("/icons/wordwrap.png", 16, 16));
        JMenuItem font = new JMenuItem("Font Settings", loadIcon("/icons/settings.png", 16, 16));
        JMenuItem insertDate = new JMenuItem("Insert Date/Time", loadIcon("/icons/insertdate.png", 16, 16));
        JMenuItem fontFamily = new JMenuItem("Font Family", loadIcon("/icons/fontfamily.png", 16, 16));
        JMenuItem fontSize = new JMenuItem("Font Size", loadIcon("/icons/fontsize.png", 16, 16));
        JMenuItem syntaxHighlight = new JMenuItem("Syntax Highlighting", loadIcon("/icons/syntaxhighlite.png", 16, 16));
        JMenuItem spellCheck = new JMenuItem("Spell Check", loadIcon("/icons/spellcheck.png", 16, 16));
        JMenuItem textStats = new JMenuItem("Text Statistics", loadIcon("/icons/textstatistics.png", 16, 16));
        JMenuItem compareTabs = new JMenuItem("Compare Tabs", loadIcon("/icons/comparetab.png", 16, 16));
        JMenuItem bookmarkLine = new JMenuItem("Bookmark Line", loadIcon("/icons/bookmarkline.png", 16, 16));
        JMenuItem columnEdit = new JMenuItem("Column Editing", loadIcon("/icons/arrangecolumn.png", 16, 16));
        JMenuItem autoSave = new JMenuItem("Auto-save (5 min)", loadIcon("/icons/save.png", 16, 16));

        // Add this inside your menu or toolbar:
        JMenuItem openZip = new JMenuItem("Open ZIP File (Preview)",
                loadIcon("/icons/openzipfile.png", 16, 16));
        openZip.setAccelerator(
                KeyStroke.getKeyStroke(KeyEvent.VK_Z, InputEvent.CTRL_DOWN_MASK | InputEvent.SHIFT_DOWN_MASK));
        // Add action listener to openZip
        openZip.addActionListener(e -> ZipViewer.showZipContents(null));

        Font emojiFont = new Font("Segoe UI Emoji", Font.PLAIN, 16); // Change font name based on OS

        // Assuming you're using a JTextArea called textArea

        JMenuItem zoomIn = new JMenuItem("Zoom In +\t", loadIcon("/icons/zoomin.png", 16, 16));
        JMenuItem zoomOut = new JMenuItem("Zoom Out -\t", loadIcon("/icons/zoomout.png", 16, 16));
        JMenuItem resetZoom = new JMenuItem("Reset Zoom", loadIcon("/icons/replace.png", 16, 16));
        JCheckBoxMenuItem showStatusBar = new JCheckBoxMenuItem("Show Status Bar", true);
        JMenu themeMenu = new JMenu("Theme");
        themeMenu.setIcon(loadIcon("/icons/theme.png", 16, 16));
        JMenu removeList = new JMenu("Remove");
        removeList.setIcon(loadIcon("/icons/removenum.png", 16, 16));
        JMenu letterCase = new JMenu("Letter Case");
        letterCase.setIcon(loadIcon("/icons/font.png", 16, 16));
        JMenu encodeanddecode = new JMenu("Encode and Decode");
        encodeanddecode.setIcon(loadIcon("/icons/lock.png", 16, 16));
        JMenu tabs = new JMenu("New");
        tabs.setIcon(loadIcon("/icons/new.png", 16, 16));
        JMenuItem toUpper = new JMenuItem("To UPPERCASE (A)", loadIcon("/icons/uppercase.png", 16, 16));
        JMenuItem encodeItem = new JMenuItem("Encode File to Base64", loadIcon("/icons/lock.png", 16, 16));
        encodeItem.addActionListener(e -> encodeFileToBase64());

        JMenuItem decodeItem = new JMenuItem("Decode Base64 to File", loadIcon("/icons/key.png", 16, 16));
        decodeItem.addActionListener(e -> decodeBase64ToFile());
        JMenu settings = new JMenu("Settings");

        JMenuItem toLower = new JMenuItem("To lowercase (a)", loadIcon("/icons/lowercase.png", 16, 16));
        JMenuItem toTitle = new JMenuItem("To Title Case (T)", loadIcon("/icons/title.png", 16, 16));
        JMenuItem removeDup = new JMenuItem("Remove Duplicate Lines (All)", loadIcon("/icons/removenum.png", 16, 16));
        JMenuItem sortLines = new JMenuItem("Sort Lines (Alphabetically)", loadIcon("/icons/short.png", 16, 16));
        JMenuItem duplicateLine = new JMenuItem("Duplicate Line(s)",
                loadIcon("/icons/duplicate.png", 16, 16));
        JMenuItem removeBlank = new JMenuItem("Remove Blank Lines (All)", loadIcon("/icons/removenum.png", 16, 16));
        JMenuItem trimLines = new JMenuItem("Trim All Lines (Whitespace)", loadIcon("/icons/trimlines.png", 16, 16));
        JMenuItem joinLines = new JMenuItem("Join Lines (All)", loadIcon("/icons/joinlines.png", 16, 16));
        JMenuItem tabsToSpaces = new JMenuItem("Convert Tabs to Spaces (4 spaces)",
                loadIcon("/icons/tabspaces.png", 16, 16));
        JMenuItem bgColor = new JMenuItem("Background Color ", loadIcon("/icons/bgcolor.png", 16, 16));
        JMenuItem customPalette = new JMenuItem("Custom Color Palette for Text ",
                loadIcon("/icons/palette.png", 16, 16));
        JMenuItem bold = new JMenuItem("Bold (B)", loadIcon("/icons/bold.png", 16, 16));
        JMenuItem italic = new JMenuItem("Italic (/)", loadIcon("/icons/italic.png", 16, 16));
        JMenuItem underline = new JMenuItem("Underline (_)", loadIcon("/icons/underline.png", 16, 16));
        JMenuItem tableTab = new JMenuItem("New Table Tab ", loadIcon("/icons/excel.png", 16, 16));
        JMenuItem compiler = new JMenuItem("New Compiler ", loadIcon("/icons/compiler.png", 16, 16));
        JMenuItem drawingTab = new JMenuItem("New Drawing Tab ", loadIcon("/icons/drawing.png", 16, 16));
        JMenuItem sumColumn = new JMenuItem("Sum Column (Table) ", loadIcon("/icons/sumcolumn.png", 16, 16));
        JMenuItem avgColumn = new JMenuItem("Average Column (Table) ", loadIcon("/icons/arrangecolumn.png", 16, 16));
        JMenuItem pluginManager = new JMenuItem("Plugin Manager ", loadIcon("/icons/plugin.png", 16, 16));
        JMenuItem reverseLines = new JMenuItem("Reverse Lines ", loadIcon("/icons/reverse.png", 16, 16));
        JMenuItem removeNumbers = new JMenuItem("Remove All Numbers ", loadIcon("/icons/removenum.png", 16, 16));
        JMenuItem removePunct = new JMenuItem("Remove All Punctuation ", loadIcon("/icons/removenum.png", 16, 16));
        JMenuItem shuffleLines = new JMenuItem("Shuffle Lines ", loadIcon("/icons/shuffle.png", 16, 16));
        JMenuItem countWords = new JMenuItem("Count the Words or Characters ", loadIcon("/icons/count.png", 16, 16));
        JMenuItem removeDupWords = new JMenuItem("Remove Duplicate Words (All) ",
                loadIcon("/icons/removenum.png", 16, 16));
        JMenuItem findWord = new JMenuItem("Find Word or Phrase ", loadIcon("/icons/find.png", 16, 16));
        JMenuItem filterLinesItem = new JMenuItem("Filter Lines by Keyword ", loadIcon("/icons/filter.png", 16, 16));
        JMenuItem aiAssistant = new JMenuItem("AI Assistant (Offline)", loadIcon("/icons/chatbot.png", 16, 16));
        aiAssistant.addActionListener(e -> AIAssistantAddon.show());
        addons.add(aiAssistant);
        JFrame frame = new JFrame("Your Title");

        add(tabs, BorderLayout.CENTER);

        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setLayout(new BorderLayout());
        frame.setSize(800, 600);

        RSyntaxTextArea textArea = new RSyntaxTextArea(20, 60);
        textArea.setSyntaxEditingStyle(SyntaxConstants.SYNTAX_STYLE_JAVA);
        textArea.setCodeFoldingEnabled(true);

        scrollPane = new JScrollPane(textArea);

        RTextScrollPane scrollPane = new RTextScrollPane(textArea); // Adds line numbers

        frame.add(scrollPane, BorderLayout.CENTER);
        JPanel searchPanel = new JPanel(new FlowLayout(FlowLayout.LEFT));
        JTextField searchField = new JTextField(20);
        searchField.setBorder(BorderFactory.createCompoundBorder(
                BorderFactory.createLineBorder(Color.GRAY, 1, true),
                BorderFactory.createEmptyBorder(5, 10, 5, 10)));
        searchField.setBackground(new Color(245, 245, 245));
        searchField.setFont(new Font("SansSerif", Font.PLAIN, 14));

        // Add to top of window (assuming you have a BorderLayout)

        // Mac font fallback

        // if you are using a JFrame named 'frame'

        SyntaxTextPane editor = new SyntaxTextPane();
        editor.setText("public class Test {\n    private int number = 5;\n}");

        JMenuItem about = new JMenuItem("About me ? ", loadIcon("/icons/about.png", 16, 16));

        newTab.setAccelerator(KeyStroke.getKeyStroke(KeyEvent.VK_N, InputEvent.CTRL_DOWN_MASK));
        open.setAccelerator(KeyStroke.getKeyStroke(KeyEvent.VK_O, InputEvent.CTRL_DOWN_MASK));
        save.setAccelerator(KeyStroke.getKeyStroke(KeyEvent.VK_S, InputEvent.CTRL_DOWN_MASK));
        close.setAccelerator(KeyStroke.getKeyStroke(KeyEvent.VK_W, InputEvent.CTRL_DOWN_MASK));
        cut.setAccelerator(KeyStroke.getKeyStroke(KeyEvent.VK_X, InputEvent.CTRL_DOWN_MASK));
        copy.setAccelerator(KeyStroke.getKeyStroke(KeyEvent.VK_C, InputEvent.CTRL_DOWN_MASK));
        paste.setAccelerator(KeyStroke.getKeyStroke(KeyEvent.VK_V, InputEvent.CTRL_DOWN_MASK));
        undo.setAccelerator(KeyStroke.getKeyStroke(KeyEvent.VK_Z, InputEvent.CTRL_DOWN_MASK));
        redo.setAccelerator(KeyStroke.getKeyStroke(KeyEvent.VK_Y, InputEvent.CTRL_DOWN_MASK));
        find.setAccelerator(KeyStroke.getKeyStroke(KeyEvent.VK_F, InputEvent.CTRL_DOWN_MASK));
        replace.setAccelerator(KeyStroke.getKeyStroke(KeyEvent.VK_H, InputEvent.CTRL_DOWN_MASK));
        goTo.setAccelerator(KeyStroke.getKeyStroke(KeyEvent.VK_L, InputEvent.CTRL_DOWN_MASK));
        wordWrapItem.setAccelerator(KeyStroke.getKeyStroke(KeyEvent.VK_Z, InputEvent.ALT_DOWN_MASK));
        insertDate.setAccelerator(KeyStroke.getKeyStroke(KeyEvent.VK_F5, 0));
        zoomIn.setAccelerator(KeyStroke.getKeyStroke(KeyEvent.VK_ADD, InputEvent.CTRL_DOWN_MASK));
        zoomOut.setAccelerator(KeyStroke.getKeyStroke(KeyEvent.VK_MINUS, InputEvent.CTRL_DOWN_MASK));
        resetZoom.setAccelerator(KeyStroke.getKeyStroke(KeyEvent.VK_0, InputEvent.CTRL_DOWN_MASK));
        duplicateLine.setAccelerator(KeyStroke.getKeyStroke(KeyEvent.VK_D, InputEvent.CTRL_DOWN_MASK));
        bold.setAccelerator(KeyStroke.getKeyStroke(KeyEvent.VK_B, InputEvent.CTRL_DOWN_MASK));
        italic.setAccelerator(KeyStroke.getKeyStroke(KeyEvent.VK_I, InputEvent.CTRL_DOWN_MASK));
        underline.setAccelerator(KeyStroke.getKeyStroke(KeyEvent.VK_U, InputEvent.CTRL_DOWN_MASK));
        add(scrollPane, BorderLayout.CENTER);
        pptxTab.addActionListener(e -> addPPTXTab());
        pptxPanel.setLayout(new BorderLayout());
        pptxPreviewPanel.setLayout(new GridLayout(0, 1));

        JMenuItem lightthemeitem = new JMenuItem("Light Theme", loadIcon("/icons/white.png", 16, 16));
        lightthemeitem.addActionListener(e -> lightthemeitem());

        JMenuItem darkthemeitem = new JMenuItem("Dark Theme", loadIcon("/icons/dark.png", 16, 16));
        darkthemeitem.addActionListener(e -> darkthemeitem());

        lightthemeitem.addActionListener(e -> {
            try {
                applyTheme("light", this);
            } catch (UnsupportedLookAndFeelException ex) {
                ex.printStackTrace();
            }
        });

        darkthemeitem.addActionListener(e -> {
            try {
                applyTheme("dark", this);
            } catch (UnsupportedLookAndFeelException ex) {
                ex.printStackTrace();
            }
        });

        KeyStroke keyStroke = KeyStroke.getKeyStroke(KeyEvent.VK_F, InputEvent.CTRL_DOWN_MASK);

        getRootPane().getInputMap(JComponent.WHEN_IN_FOCUSED_WINDOW)
                .put(keyStroke, "toggleSearch");

        getRootPane().getActionMap().put("toggleSearch", new AbstractAction() {
            @Override
            public void actionPerformed(ActionEvent e) {
                searchPanel.setVisible(!searchPanel.isVisible());
                if (searchPanel.isVisible()) {
                    searchField.requestFocusInWindow();
                }
            }
        });

        calculator.addActionListener(e -> showCalculatorPopup());
        currencyConverter.addActionListener(e -> showCurrencyConverterPopup());
        notepadMini.addActionListener(e -> showMiniNotepadPopup());
        stopwatch.addActionListener(e -> showStopwatchPopup());

        reverseLines.addActionListener(e -> reverseLines());
        removeNumbers.addActionListener(e -> removeAllNumbers());
        removePunct.addActionListener(e -> removeAllPunctuation());
        shuffleLines.addActionListener(e -> shuffleLines());
        countWords.addActionListener(e -> countWordsAndChars());
        removeDupWords.addActionListener(e -> removeDuplicateWords());
        findWord.addActionListener(e -> showFindWordDialog());
        toUpper.addActionListener(e -> toUppercaseAll());
        System.setProperty("jna.library.path", "vlc");
        System.setProperty("VLC_PLUGIN_PATH", "vlc/plugins");

        removeDup.addActionListener(e -> removeDuplicateLinesAll());
        sortLines.addActionListener(e -> sortLinesAll());

        newTab.addActionListener(e -> {
            if (confirmSave())
                addNewTab(null, "");
        });
        open.addActionListener(e -> {
            if (confirmSave())
                openFile();
        });
        save.addActionListener(e -> saveFile());
        saveAs.addActionListener(e -> saveFileAs());
        close.addActionListener(e -> closeCurrentTab());
        exit.addActionListener(e -> {
            if (confirmSaveAll())
                System.exit(0);
        });

        cut.addActionListener(e -> getCurrentTextPane().cut());
        copy.addActionListener(e -> getCurrentTextPane().copy());
        paste.addActionListener(e -> getCurrentTextPane().paste());
        undo.addActionListener(e -> {
            UndoManager um = undoMap.get(getCurrentTextPane());
            if (um != null && um.canUndo())
                um.undo();
        });
        redo.addActionListener(e -> {
            UndoManager um = undoMap.get(getCurrentTextPane());
            if (um != null && um.canRedo())
                um.redo();
        });
        wordWrapItem.addActionListener(e -> {
            JTextPane pane = getCurrentTextPane();
            if (pane != null)
                pane.setEditorKit(wordWrapItem.isSelected() ? new WrapEditorKit() : new StyledEditorKit());
        });
        font.addActionListener(e -> {
            JTextPane pane = getCurrentTextPane();
            if (pane != null) {
                String fontName = JOptionPane.showInputDialog(this, "Font name 🔤:", pane.getFont().getFamily());
                String sizeStr = JOptionPane.showInputDialog(this, "Font size 🔢:",
                        String.valueOf(pane.getFont().getSize()));
                if (fontName != null && sizeStr != null) {
                    try {
                        int size = Integer.parseInt(sizeStr);
                        pane.setFont(new Font(fontName, Font.PLAIN, size));
                    } catch (NumberFormatException ignored) {
                    }
                }
            }
        });
        fontFamily.addActionListener(e -> {
            JTextPane pane = getCurrentTextPane();
            if (pane != null) {
                String[] fonts = GraphicsEnvironment.getLocalGraphicsEnvironment().getAvailableFontFamilyNames();
                String fontName = (String) JOptionPane.showInputDialog(this, "Choose font family 🔤:", "Font Family 🔤",
                        JOptionPane.PLAIN_MESSAGE, null, fonts, pane.getFont().getFamily());
                if (fontName != null) {
                    Font old = pane.getFont();
                    pane.setFont(new Font(fontName, old.getStyle(), old.getSize()));
                }
            }
        });
        fontSize.addActionListener(e -> {
            JTextPane pane = getCurrentTextPane();
            if (pane != null) {
                String sizeStr = JOptionPane.showInputDialog(this, "Font size 🔢:",
                        String.valueOf(pane.getFont().getSize()));
                if (sizeStr != null) {
                    try {
                        int size = Integer.parseInt(sizeStr);
                        Font old = pane.getFont();
                        pane.setFont(new Font(old.getFamily(), old.getStyle(), size));
                    } catch (NumberFormatException ignored) {
                    }
                }
            }
        });
        insertDate.addActionListener(e -> {
            JTextPane pane = getCurrentTextPane();
            if (pane != null) {
                pane.replaceSelection(new java.util.Date().toString());
            }
        });
        find.addActionListener(e -> showFindDialog());
        replace.addActionListener(e -> showReplaceDialog());
        goTo.addActionListener(e -> showGoToDialog());

        zoomIn.addActionListener(e -> changeFontSize(2));
        zoomOut.addActionListener(e -> changeFontSize(-2));
        resetZoom.addActionListener(e -> setFontSize(14));
        showStatusBar.addActionListener(e -> statusBar.setVisible(showStatusBar.isSelected()));
        base64EncodeItem.addActionListener(e -> encodeBase64());
        base64DecodeItem.addActionListener(e -> decodeBase64());

        toUpper.addActionListener(e -> transformSelectedText(String::toUpperCase));
        toLower.addActionListener(e -> transformSelectedText(String::toLowerCase));
        toTitle.addActionListener(e -> transformSelectedText(TextEditor::toTitleCase));
        removeDup.addActionListener(e -> removeDuplicateLines());
        sortLines.addActionListener(e -> sortLines());
        duplicateLine.addActionListener(e -> duplicateLines());
        removeBlank.addActionListener(e -> removeBlankLines());
        trimLines.addActionListener(e -> trimAllLines());
        joinLines.addActionListener(e -> joinAllLines());
        filterLinesItem.addActionListener(e -> filterLines());
        tabsToSpaces.addActionListener(e -> convertTabsToSpaces());
        bgColor.addActionListener(e -> {
            JTextPane pane = getCurrentTextPane();
            if (pane != null) {
                Color color = JColorChooser.showDialog(this, "Choose Background Color ", pane.getBackground());
                if (color != null)
                    pane.setBackground(color);
            }
        });
        customPalette.addActionListener(e -> showCustomPaletteDialog());

        bold.addActionListener(e -> setBold());
        italic.addActionListener(e -> setItalic());
        underline.addActionListener(e -> setUnderline());

        // CORRECT CODE
        tableTab.addActionListener(e -> {
            // The action is now simple: just create a new, blank table tab.
            addNewTableTab();
        });// Make sure there is NO code trying to access 'chooser' after this point.

        // The extra line that caused the error has been removed.
        compiler.addActionListener(e -> {
            addCompilerTab();
        });

        drawingTab.addActionListener(e -> addDrawingTab());
        sumColumn.addActionListener(e -> sumSelectedTableColumn());
        avgColumn.addActionListener(e -> avgSelectedTableColumn());

        about.addActionListener(e -> JOptionPane.showMessageDialog(this,
                "IdeaPad Java Edition\n\n"
                        + "Keyboard Shortcuts:\n"
                        + "Ctrl+N: New Tab\n"
                        + "Ctrl+O: Open File\n"
                        + "Ctrl+S: Save\n"
                        + "Ctrl+W: Close Tab\n"
                        + "Alt+F4: Exit\n"
                        + "Ctrl+X: Cut\n"
                        + "Ctrl+C: Copy\n"
                        + "Ctrl+V: Paste\n"
                        + "Ctrl+Z: Undo\n"
                        + "Ctrl+Y: Redo\n"
                        + "Ctrl+F: Find\n"
                        + "Ctrl+H: Replace\n"
                        + "Ctrl+L: Go To Line\n"
                        + "Ctrl+G: Open Browser\n"
                        + "Alt+Z: Word Wrap\n"
                        + "F5: Insert Date/Time\n"
                        + "Ctrl++: Zoom In\n"
                        + "Ctrl+-: Zoom Out\n"
                        + "Ctrl+0: Reset Zoom\n"
                        + "Ctrl+D: Duplicate Line\n"
                        + "Ctrl+B: Bold\n"
                        + "Ctrl+I: Italic\n"
                        + "Ctrl+U: Underline\n"
                        + "Table: Ctrl+B/I/U in cell for formatting\n"
                        + "Table: Use Tools > Sum/Average Column\n"
                        + "PDF: Open/Edit/Save as plain text\n",
                "Sorry, its Help! ", JOptionPane.INFORMATION_MESSAGE));
        file.add(tabs);
        file.add(opened);
        file.add(saver);
        tabs.add(newTab);
        tabs.add(tableTab);
        tabs.add(compiler);
        tabs.add(drawingTab);
        tabs.add(pptxTab);
        opened.add(open);
        opened.add(openZip);
        saver.add(save);
        saver.add(saveAs);
        file.add(close);

        file.addSeparator();
        file.add(exit);
        edit.add(cut);
        edit.add(copy);
        edit.add(paste);
        edit.addSeparator();
        edit.add(undo);
        edit.add(redo);
        edit.addSeparator();
        edit.add(find);
        edit.add(replace);
        edit.add(goTo);
        tools.add(filterLinesItem);
        edit.addSeparator();
        edit.add(wordWrapItem);
        textas.add(font);
        textas.add(fontFamily);
        textas.add(fontSize);
        settings.add(textas);
        settings.add(colours);
        edit.add(insertDate);
        file.add(print);
        addons.add(browserMenu);

        expertor.add(exportHTML);
        expertor.add(exportMD);
        tools.addSeparator();
        tools.add(syntaxHighlight);
        tools.add(spellCheck);
        tools.add(textStats);
        tools.add(compareTabs);
        tools.add(bookmarkLine);
        tools.add(columnEdit);
        tools.add(autoSave);
        encodeanddecode.add(base64EncodeItem);
        encodeanddecode.add(base64DecodeItem);
        encodeanddecode.add(encodeItem);
        encodeanddecode.add(decodeItem);
        tools.addSeparator();
        addons.add(encodeanddecode);
        addons.add(calculator);
        addons.add(currencyConverter);
        addons.add(notepadMini);
        addons.add(stopwatch);
        themeMenu.add(darkthemeitem);
        themeMenu.add(lightthemeitem);

        font.addActionListener(e -> {
            /* your font dialog code */});

        view.add(themeMenu);
        view.add(zoomIn);
        view.add(zoomOut);
        view.add(resetZoom);
        view.addSeparator();
        view.add(showStatusBar);
        settings.add(pluginManager);
        tools.add(removeList);
        tools.add(letterCase);
        letterCase.add(toUpper);
        letterCase.add(toLower);
        letterCase.add(toTitle);
        tools.addSeparator();
        removeList.add(removeDup);
        tools.add(sortLines);
        tools.addSeparator();
        tools.add(duplicateLine);
        removeList.add(removeBlank);
        tools.add(trimLines);
        tools.add(joinLines);
        tools.add(tabsToSpaces);
        tools.addSeparator();
        colours.add(bgColor);
        colours.add(customPalette);
        letterCase.add(bold);
        letterCase.add(italic);
        letterCase.add(underline);
        tools.add(sumColumn);
        tools.add(avgColumn);
        tools.addSeparator();
        tools.add(reverseLines);
        removeList.add(removeNumbers);
        removeList.add(removePunct);
        removeList.add(removeDupWords);
        tools.add(shuffleLines);
        tools.add(countWords);
        
        tools.add(findWord);
        edit.add(jsonFormatter);
        addons.add(minifier);
        addons.add(baseConverter);
        view.add(regexTester);
        addons.add(loremIpsum);
        tools.add(qrGenerator);
        addons.add(uuidGenerator);
        addons.add(hashChecker);

        addons.add(tasksMenu);
        expertor.add(exportQR);
        file.add(expertor);

        help.add(about);

        bar.add(file);
        bar.add(edit);
        bar.add(view);
        bar.add(tools);
        bar.add(settings);

        bar.add(addons);
        bar.add(help);
        // Push everything that follows to the right
        bar.add(Box.createHorizontalGlue());

        // Create a clean macOS-style search field

        // Set up search field
        // UI preferences for text components (before creating the field)
        UIManager.put("TextComponent.arc", 20); // Rounded corners for all text components

        // Initialize search field

        // Set a modern macOS-like font
        searchField.setFont(new Font("San Francisco", Font.PLAIN, 14)); // Fallbacks to "Segoe UI" if not available

        // Size and rounded corners
        searchField.setPreferredSize(new Dimension(220, 36));
        searchField.setMaximumSize(new Dimension(220, 36));
        searchField.setMinimumSize(new Dimension(200, 32));
        searchField.setBorder(BorderFactory.createEmptyBorder(6, 14, 6, 14)); // Top, Left, Bottom, Right

        // macOS-style placeholder text and appearance tweaks
        searchField.putClientProperty("JTextField.placeholderText", "🔍 Search...");
        searchField.putClientProperty("JComponent.roundRect", true);
        searchField.putClientProperty("TextComponent.arc", 22);
        searchField.putClientProperty("JTextField.showClearButton", true);

        // Optional: reduce opacity a bit like macOS search bar
        searchField.setOpaque(false);
        searchField.setBackground(new Color(255, 255, 255, 220)); // for light theme

        // Adds built-in clear (X) button
        bar.add(searchField);

        return bar;
    }



/**
 * InfiniteStyledTableModel
 * A flexible reflective table model that delegates calls to a parent XlsxEditorPanel or similar model.
 * Supports colors, images, formulas, comments, and batch updates.
 */
class InfiniteStyledTableModel extends AbstractTableModel {
    private final Object parentPanel;

    public InfiniteStyledTableModel(Object parentPanel) {
        this.parentPanel = parentPanel;
    }

    // ==================== CORE MODEL ACCESS ====================

    private Object getInternalModel() {
        try {
            return parentPanel.getClass().getMethod("getInternalModel").invoke(parentPanel);
        } catch (Exception e) {
            return null;
        }
    }

    private Object invokeInternal(String method, Class<?>[] params, Object... args) {
        try {
            Object internalModel = getInternalModel();
            if (internalModel == null) return null;
            Method m = internalModel.getClass().getMethod(method, params);
            return m.invoke(internalModel, args);
        } catch (Exception e) {
            return null;
        }
    }

    // ==================== CORE TABLE MODEL ====================

    @Override
    public int getRowCount() {
        Object result = invokeInternal("getRowCount", new Class[]{});
        return result instanceof Integer ? (int) result : 1000;
    }

    @Override
    public int getColumnCount() {
        Object result = invokeInternal("getColumnCount", new Class[]{});
        return result instanceof Integer ? (int) result : 26;
    }

    @Override
    public Object getValueAt(int row, int col) {
        Object result = invokeInternal("getValueAt", new Class[]{int.class, int.class}, row, col);
        return result != null ? result : "";
    }

    @Override
    public void setValueAt(Object value, int row, int col) {
        invokeInternal("setValueAt", new Class[]{Object.class, int.class, int.class}, value, row, col);
        fireTableCellUpdated(row, col);
    }

    @Override
    public boolean isCellEditable(int row, int col) {
        return true;
    }

    @Override
    public String getColumnName(int col) {
        try {
            Object name = parentPanel.getClass().getMethod("getColumnName", int.class).invoke(parentPanel, col);
            return name != null ? name.toString() : calculateColumnName(col);
        } catch (Exception e) {
            return calculateColumnName(col);
        }
    }

    @Override
    public Class<?> getColumnClass(int col) {
        return Object.class;
    }

    // ==================== BATCH PROCESSING ====================

    public void batchUpdate(Runnable updateOperation) {
        Object internalModel = getInternalModel();
        if (internalModel != null) {
            try {
                internalModel.getClass().getMethod("batchUpdate", Runnable.class)
                        .invoke(internalModel, updateOperation);
                return;
            } catch (Exception ignored) {}
        }
        // Fallback
        updateOperation.run();
        fireTableDataChanged();
    }

    // ==================== CELL STYLING ====================

    public void setCellColor(int row, int col, Color color) {
        invokeInternal("setCellColor", new Class[]{int.class, int.class, Color.class}, row, col, color);
    }

    public Color getCellColor(int row, int col) {
        Object result = invokeInternal("getCellColor", new Class[]{int.class, int.class}, row, col);
        return result instanceof Color ? (Color) result : null;
    }

    public void setCellImage(int row, int col, ImageIcon image) {
        invokeInternal("setCellImage", new Class[]{int.class, int.class, ImageIcon.class}, row, col, image);
    }

    public ImageIcon getCellImage(int row, int col) {
        Object result = invokeInternal("getCellImage", new Class[]{int.class, int.class}, row, col);
        return result instanceof ImageIcon ? (ImageIcon) result : null;
    }

    public boolean hasImage(int row, int col) {
        Object result = invokeInternal("hasImage", new Class[]{int.class, int.class}, row, col);
        return result instanceof Boolean && (boolean) result;
    }

    // ==================== FORMULA SUPPORT ====================

    public void setCellFormula(int row, int col, String formula) {
        invokeInternal("setCellFormula", new Class[]{int.class, int.class, String.class}, row, col, formula);
    }

    public String getCellFormula(int row, int col) {
        Object result = invokeInternal("getCellFormula", new Class[]{int.class, int.class}, row, col);
        return result != null ? result.toString() : null;
    }

    // ==================== FORMATTING SUPPORT ====================

    public void setCellFormat(int row, int col, String format) {
        invokeInternal("setCellFormat", new Class[]{int.class, int.class, String.class}, row, col, format);
    }

    public String getCellFormat(int row, int col) {
        Object result = invokeInternal("getCellFormat", new Class[]{int.class, int.class}, row, col);
        return result != null ? result.toString() : null;
    }

    // ==================== COMMENT SUPPORT ====================

    public void setCellComment(int row, int col, String comment) {
        invokeInternal("setCellComment", new Class[]{int.class, int.class, String.class}, row, col, comment);
    }

    public String getCellComment(int row, int col) {
        Object result = invokeInternal("getCellComment", new Class[]{int.class, int.class}, row, col);
        return result != null ? result.toString() : null;
    }

    public boolean hasComment(int row, int col) {
        Object result = invokeInternal("hasComment", new Class[]{int.class, int.class}, row, col);
        return result instanceof Boolean && (boolean) result;
    }

    // ==================== RENDERING ====================

    public Component prepareTableCellRenderer(Component comp, int row, int col,
                                              boolean alternateRowColors, boolean tableMode) {
        Object result = invokeInternal("prepareTableCellRenderer",
                new Class[]{Component.class, int.class, int.class, boolean.class, boolean.class},
                comp, row, col, alternateRowColors, tableMode);
        return result instanceof Component ? (Component) result : comp;
    }

    // ==================== DATA MANAGEMENT ====================

    public Object getRawValueAt(int row, int col) {
        return invokeInternal("getRawValueAt", new Class[]{int.class, int.class}, row, col);
    }

    public Object[][] getDataSnapshot() {
        Object result = invokeInternal("getDataSnapshot", new Class[]{});
        return result instanceof Object[][] ? (Object[][]) result : new Object[0][0];
    }

    public String[] getColumnNames() {
        Object result = invokeInternal("getColumnNames", new Class[]{});
        return result instanceof String[] ? (String[]) result : new String[0];
    }

    public void loadData(Object[][] newData, String[] columnNames) {
        invokeInternal("loadData", new Class[]{Object[][].class, String[].class}, newData, columnNames);
    }

    // ==================== ACTION HANDLERS ====================

    public void addNewRow() {
        invokeParent("handleAddRow");
    }

    public void deleteSelectedRow() {
        invokeParent("handleDeleteRow");
    }

    public void exportToCSV() {
        invokeParent("handleExportCSV");
    }

    public void saveToXlsx() {
        invokeParent("handleSave", new Class[]{ActionEvent.class},
                new ActionEvent(this, ActionEvent.ACTION_PERFORMED, "save"));
    }

    public void insertImageIntoSelectedCell() {
        invokeParent("handleInsertImage", new Class[]{ActionEvent.class}, (ActionEvent) null);
    }

    public void setColorForSelectedCell() {
        invokeParent("handleSetCellColor", new Class[]{ActionEvent.class}, (ActionEvent) null);
    }

    private void invokeParent(String method, Class<?>[] params, Object... args) {
        try {
            Method m = parentPanel.getClass().getMethod(method, params);
            m.invoke(parentPanel, args);
        } catch (Exception ignored) {}
    }

    private void invokeParent(String method) {
        invokeParent(method, new Class[]{});
    }

    // ==================== UTILITIES ====================

    private String calculateColumnName(int col) {
        StringBuilder name = new StringBuilder();
        int c = col;
        do {
            name.insert(0, (char) ('A' + (c % 26)));
            c = c / 26 - 1;
        } while (c >= 0);
        return name.toString();
    }
}


// ==================== USAGE IN XlsxEditorPanel ====================
// Update XlsxEditorPanel constructor to pass 'this' reference:
//
// public XlsxEditorPanel(File file) {
//     this.file = file;
//     setLayout(new BorderLayout());
//     
//     // Pass 'this' to the model so it can access panel methods
//     model = new InfiniteStyledTableModel(this);
//     table = new JTable(model) {
//         @Override
//         public Component prepareRenderer(TableCellRenderer renderer, int row, int col) {
//             Component comp = super.prepareRenderer(renderer, row, col);
//             return model.prepareTableCellRenderer(comp, row, col, alternateRowColors, tableMode);
//         }
//     };
//
//     setupTable();
//     setupUI();
//     loadXlsxContent();
//     autoResizeColumns();
// }
//
// Add getter method for table access:
// public JTable getTable() {
//     return table;
// }
    /**
     * Creates and adds a new tab with a code editor and a console to compile and
     * run Java code.
     */
    /**
     * Creates and adds a new tab with a multi-language code editor and a console
     * to compile and run Java, Python, C++, and C code.
     */
    /**
     * A simple data class to hold a language's name and its icon.
     */
    /**
     * A data class to hold a language's name and its icon.
     */
    class LanguageItem {
        private String name;
        private Icon icon;

        public LanguageItem(String name, Icon icon) {
            this.name = name;
            this.icon = icon;
        }

        public String getName() {
            return name;
        }

        public Icon getIcon() {
            return icon;
        }

        @Override
        public String toString() {
            return name; // Used to display the text in the selected item box
        }
    }

    /**
     * A custom renderer to display both an icon and text in the JComboBox list.
     */
    /**
     * A custom renderer to display both an icon and text in the JComboBox list.
     */
    class LanguageRenderer extends JLabel implements ListCellRenderer<LanguageItem> {
        public LanguageRenderer() {
            setOpaque(true);
        }

        @Override
        public Component getListCellRendererComponent(JList<? extends LanguageItem> list, LanguageItem value, int index,
                boolean isSelected, boolean cellHasFocus) {
            if (value != null) {
                setText(value.getName());
                setIcon(value.getIcon());
            }

            if (isSelected) {
                setBackground(list.getSelectionBackground());
                setForeground(list.getSelectionForeground());
            } else {
                setBackground(list.getBackground());
                setForeground(list.getForeground());
            }
            return this;
        }
    }

    private final CompilerService compilerService = new CompilerService();

 private void addCompilerTab() {
    // --- UI Setup ---
    RSyntaxTextArea codeEditor = new RSyntaxTextArea();
    codeEditor.setCodeFoldingEnabled(true);
    codeEditor.setFont(new Font("Consolas", Font.PLAIN, 14));

    // Apply RSyntaxTextArea theme based on current FlatLaf theme
    applyEditorTheme(codeEditor);

    // Console setup (acts as both input/output)
    JTextArea console = new JTextArea();
    console.setEditable(true);
    console.setFont(new Font("Consolas", Font.PLAIN, 12));
    console.setLineWrap(true);
    console.setWrapStyleWord(true);

    // Run button - using loadAndScaleCloseIcon approach
    ImageIcon runIcon = loadAndScaleCloseIcon("run.png", 16, 16);
    JButton runButton = new JButton("Compile & Run", runIcon);

    // Language selector - using loadAndScaleCloseIcon approach
    LanguageItem[] languages = {
            new LanguageItem("Java", loadAndScaleCloseIcon("java.png", 16, 16)),
            new LanguageItem("Python", loadAndScaleCloseIcon("python.png", 16, 16)),
            new LanguageItem("C++", loadAndScaleCloseIcon("cpp.png", 16, 16)),
            new LanguageItem("C", loadAndScaleCloseIcon("c.png", 16, 16))
    };

    JComboBox<LanguageItem> languageSelector = new JComboBox<>(languages);
    languageSelector.setRenderer(new LanguageRenderer());

    languageSelector.addActionListener(e -> {
        LanguageItem selected = (LanguageItem) languageSelector.getSelectedItem();
        if (selected == null)
            return;

        String language = selected.getName();
        codeEditor.setSyntaxEditingStyle(LanguageTemplates.getSyntaxStyle(language));
        codeEditor.setText(LanguageTemplates.getTemplate(language));
    });
    languageSelector.setSelectedIndex(0);

    // Layout setup
    JPanel topBar = new JPanel(new BorderLayout());
    topBar.add(languageSelector, BorderLayout.WEST);
    topBar.add(runButton, BorderLayout.EAST);

    JPanel editorPanel = new JPanel(new BorderLayout());
    editorPanel.add(new RTextScrollPane(codeEditor), BorderLayout.CENTER);
    editorPanel.add(topBar, BorderLayout.SOUTH);

    JPanel consolePanel = new JPanel(new BorderLayout());
    consolePanel.add(new JScrollPane(console), BorderLayout.CENTER);

    JSplitPane splitPane = new JSplitPane(JSplitPane.VERTICAL_SPLIT, editorPanel, consolePanel);
    splitPane.setResizeWeight(0.7);

    // --- Run Button Action ---
    runButton.addActionListener(e -> {
        String sourceCode = codeEditor.getText();
        LanguageItem selectedLanguage = (LanguageItem) languageSelector.getSelectedItem();
        if (selectedLanguage == null)
            return;

        console.setText("Executing " + selectedLanguage.getName() + " code...\n");
        runButton.setEnabled(false);

        // Create a pipe for real-time input/output
        PipedOutputStream inputPipe = new PipedOutputStream();
        PipedInputStream inputStream;
        try {
            inputStream = new PipedInputStream(inputPipe, 1024);
        } catch (IOException ex) {
            console.append("Error setting up input pipe: " + ex.getMessage() + "\n");
            runButton.setEnabled(true);
            return;
        }

        // Run compilation asynchronously
        CompletableFuture<String> executionFuture = compilerService.compileAndRunWithStream(
                sourceCode, inputStream, selectedLanguage.getName());

        // Input panel (for real-time program input)
        JPanel inputPanel = new JPanel(new BorderLayout());
        JLabel inputLabel = new JLabel("Program Input: ");
        JTextField inputField = new JTextField();
        
        // Send button with consistent icon loading
        ImageIcon sendIcon = loadAndScaleCloseIcon("send.png", 16, 16);
        JButton sendButton = new JButton("Send", sendIcon);

        inputPanel.add(inputLabel, BorderLayout.WEST);
        inputPanel.add(inputField, BorderLayout.CENTER);
        inputPanel.add(sendButton, BorderLayout.EAST);
        inputPanel.setBorder(BorderFactory.createEmptyBorder(5, 5, 5, 5));

        consolePanel.add(inputPanel, BorderLayout.SOUTH);
        consolePanel.revalidate();
        consolePanel.repaint();
        inputField.requestFocus();

        // Send input functionality
        Runnable sendInput = () -> {
            String input = inputField.getText().trim();
            if (!input.isEmpty()) {
                try {
                    inputPipe.write((input + "\n").getBytes());
                    inputPipe.flush();
                    console.append("> " + input + "\n");
                    inputField.setText("");
                } catch (IOException ex) {
                    console.append("Error sending input: " + ex.getMessage() + "\n");
                }
            }
        };

        sendButton.addActionListener(ae -> sendInput.run());
        inputField.addActionListener(ae -> sendInput.run());

        // Handle completion
        executionFuture.thenAccept(result -> {
            SwingUtilities.invokeLater(() -> {
                try {
                    inputPipe.close();
                    inputStream.close();
                } catch (IOException ex) { /* ignore */ }

                consolePanel.remove(inputPanel);
                consolePanel.revalidate();
                consolePanel.repaint();

                console.append("\n--- PROGRAM COMPLETED ---\n");
                console.append(result);
                runButton.setEnabled(true);
            });
        }).exceptionally(ex -> {
            SwingUtilities.invokeLater(() -> {
                try {
                    inputPipe.close();
                    inputStream.close();
                } catch (IOException ioEx) { /* ignore */ }

                consolePanel.remove(inputPanel);
                consolePanel.revalidate();
                consolePanel.repaint();
                console.append("\nExecution error: " + ex.getMessage() + "\n");
                runButton.setEnabled(true);
            });
            return null;
        });
    });

    // --- Add compiler tab with close button ---
    ImageIcon compilerIcon = loadAndScaleCloseIcon("compiler.png", 16, 16);
    Icon finalIcon = compilerIcon != null ? compilerIcon : createFallbackCompilerIcon();

    // ✅ Add tab with your existing close button support
    addTabWithCloseButton("Compiler", splitPane, finalIcon);

    // Keep track of tab type
    tabTypeMap.put(splitPane, "compiler");

    tabbedPane.setSelectedComponent(splitPane);
}

/**
 * Fallback compiler icon if PNG is not available
 */
private Icon createFallbackCompilerIcon() {
    return new Icon() {
        @Override
        public void paintIcon(Component c, Graphics g, int x, int y) {
            Graphics2D g2d = (Graphics2D) g.create();
            g2d.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
            g2d.setColor(Color.BLUE);
            g2d.fillOval(x, y, getIconWidth(), getIconHeight());
            g2d.setColor(Color.WHITE);
            g2d.setFont(new Font("Arial", Font.BOLD, 10));
            g2d.drawString("C", x + 5, y + 11);
            g2d.dispose();
        }
        
        @Override
        public int getIconWidth() { return 16; }
        
        @Override
        public int getIconHeight() { return 16; }
    };
}

// Helper method to apply RSyntaxTextArea theme based on FlatLaf
private void applyEditorTheme(RSyntaxTextArea codeEditor) {
    // Check which LookAndFeel is currently active
    LookAndFeel currentLAF = UIManager.getLookAndFeel();
    boolean isDark = currentLAF.getClass().getName().toLowerCase().contains("dark");
    
    System.out.println("Applying theme - Dark mode: " + isDark);
    System.out.println("Current LAF: " + currentLAF.getClass().getSimpleName());
    
    try {
        String themePath = isDark ? 
            "/org/fife/ui/rsyntaxtextarea/themes/dark.xml" : 
            "/org/fife/ui/rsyntaxtextarea/themes/default.xml";
        
        InputStream is = RSyntaxTextArea.class.getResourceAsStream(themePath);
        if (is != null) {
            Theme editorTheme = Theme.load(is);
            editorTheme.apply(codeEditor);
            is.close();
            System.out.println("Successfully loaded theme: " + themePath);
        } else {
            System.err.println("Could not find theme resource: " + themePath);
            applyManualTheme(codeEditor, isDark);
        }
    } catch (Exception ex) {
        System.err.println("Error loading RSyntaxTextArea theme: " + ex.getMessage());
        applyManualTheme(codeEditor, isDark);
    }
    
    // Force UI update
    codeEditor.invalidate();
    codeEditor.revalidate();
    codeEditor.repaint();
    
    // Update gutter
    updateGutterTheme(codeEditor);
}

// Comprehensive manual theme application
private void applyManualTheme(RSyntaxTextArea codeEditor, boolean isDark) {
    System.out.println("Applying manual theme - Dark: " + isDark);
    
    if (isDark) {
        // Dark theme colors
        codeEditor.setBackground(new Color(43, 43, 43));
        codeEditor.setForeground(new Color(169, 183, 198));
        codeEditor.setCaretColor(Color.WHITE);
        codeEditor.setCurrentLineHighlightColor(new Color(50, 50, 50));
        codeEditor.setSelectionColor(new Color(38, 79, 120));
        codeEditor.setSelectedTextColor(Color.WHITE);
        codeEditor.setMatchedBracketBGColor(new Color(60, 60, 60));
        codeEditor.setMatchedBracketBorderColor(new Color(128, 128, 128));
        codeEditor.setMarkAllHighlightColor(new Color(70, 70, 120));
        codeEditor.setHyperlinkForeground(new Color(86, 156, 214));
        
        // Only set syntax scheme if Token class is available
        try {
            codeEditor.setSyntaxScheme(getDarkSyntaxScheme());
        } catch (Exception e) {
            System.err.println("SyntaxScheme not available: " + e.getMessage());
        }
    } else {
        // Light theme colors
        codeEditor.setBackground(Color.WHITE);
        codeEditor.setForeground(Color.BLACK);
        codeEditor.setCaretColor(Color.BLACK);
        codeEditor.setCurrentLineHighlightColor(new Color(255, 255, 170));
        codeEditor.setSelectionColor(new Color(200, 220, 255));
        codeEditor.setSelectedTextColor(Color.BLACK);
        codeEditor.setMatchedBracketBGColor(new Color(234, 234, 234));
        codeEditor.setMatchedBracketBorderColor(Color.GRAY);
        codeEditor.setMarkAllHighlightColor(new Color(255, 255, 0));
        codeEditor.setHyperlinkForeground(Color.BLUE);
        
        try {
            codeEditor.setSyntaxScheme(getLightSyntaxScheme());
        } catch (Exception e) {
            System.err.println("SyntaxScheme not available: " + e.getMessage());
        }
    }
}

private void updateGutterTheme(RSyntaxTextArea codeEditor) {
    Container parent = codeEditor.getParent();
    while (parent != null && !(parent instanceof RTextScrollPane)) {
        parent = parent.getParent();
    }
    
    if (parent instanceof RTextScrollPane) {
        RTextScrollPane scrollPane = (RTextScrollPane) parent;
        Gutter gutter = scrollPane.getGutter();
        
        Color bg = codeEditor.getBackground();
        Color fg = codeEditor.getForeground();
        
        gutter.setBackground(bg);
        gutter.setBorderColor(bg.brighter());
        gutter.setLineNumberColor(fg);
        gutter.setFoldIndicatorForeground(fg);
        
        scrollPane.revalidate();
        scrollPane.repaint();
    }
}

private SyntaxScheme getDarkSyntaxScheme() {
    SyntaxScheme scheme = new SyntaxScheme(true);
    
    scheme.getStyle(Token.RESERVED_WORD).foreground = new Color(86, 156, 214); // Blue
    scheme.getStyle(Token.DATA_TYPE).foreground = new Color(86, 156, 214); // Blue
    scheme.getStyle(Token.FUNCTION).foreground = new Color(220, 220, 170); // Yellow
    scheme.getStyle(Token.LITERAL_STRING_DOUBLE_QUOTE).foreground = new Color(214, 157, 133); // Orange
    scheme.getStyle(Token.LITERAL_NUMBER_DECIMAL_INT).foreground = new Color(181, 206, 168); // Green
    scheme.getStyle(Token.COMMENT_EOL).foreground = new Color(87, 166, 74); // Green
    scheme.getStyle(Token.COMMENT_MULTILINE).foreground = new Color(87, 166, 74); // Green
    scheme.getStyle(Token.COMMENT_DOCUMENTATION).foreground = new Color(87, 166, 74); // Green
    scheme.getStyle(Token.ANNOTATION).foreground = new Color(200, 120, 200); // Purple
    scheme.getStyle(Token.VARIABLE).foreground = new Color(156, 220, 254); // Light blue
    scheme.getStyle(Token.OPERATOR).foreground = new Color(212, 212, 212); // Light gray
    
    return scheme;
}

private SyntaxScheme getLightSyntaxScheme() {
    SyntaxScheme scheme = new SyntaxScheme(false);
    
    scheme.getStyle(Token.RESERVED_WORD).foreground = new Color(0, 0, 255); // Blue
    scheme.getStyle(Token.DATA_TYPE).foreground = new Color(0, 0, 255); // Blue
    scheme.getStyle(Token.FUNCTION).foreground = new Color(0, 128, 0); // Green
    scheme.getStyle(Token.LITERAL_STRING_DOUBLE_QUOTE).foreground = new Color(163, 21, 21); // Red
    scheme.getStyle(Token.LITERAL_NUMBER_DECIMAL_INT).foreground = new Color(0, 128, 128); // Teal
    scheme.getStyle(Token.COMMENT_EOL).foreground = new Color(0, 128, 0); // Green
    scheme.getStyle(Token.COMMENT_MULTILINE).foreground = new Color(0, 128, 0); // Green
    scheme.getStyle(Token.COMMENT_DOCUMENTATION).foreground = new Color(0, 128, 0); // Green
    scheme.getStyle(Token.ANNOTATION).foreground = new Color(128, 0, 128); // Purple
    scheme.getStyle(Token.VARIABLE).foreground = new Color(0, 0, 0); // Black
    scheme.getStyle(Token.OPERATOR).foreground = new Color(0, 0, 0); // Black
    
    return scheme;
}

// Update compiler tab theme when theme changes
private void updateCompilerTheme() {
    SwingUtilities.invokeLater(() -> {
        Timer timer = new Timer(150, e -> {
            for (int i = 0; i < tabbedPane.getTabCount(); i++) {
                Component comp = tabbedPane.getComponentAt(i);
                if (comp instanceof JSplitPane) {
                    updateComponentTree((JSplitPane) comp);
                }
            }
            // Force a complete UI refresh
            tabbedPane.revalidate();
            tabbedPane.repaint();
        });
        timer.setRepeats(false);
        timer.start();
    });
}

private void updateComponentTree(Container container) {
    for (Component c : container.getComponents()) {
        if (c instanceof RTextScrollPane) {
            RTextScrollPane scrollPane = (RTextScrollPane) c;
            Component view = scrollPane.getViewport().getView();
            if (view instanceof RSyntaxTextArea) {
                RSyntaxTextArea editor = (RSyntaxTextArea) view;
                applyEditorTheme(editor);
            }
        } else if (c instanceof Container) {
            updateComponentTree((Container) c);
        }
    }
}

// Add these theme methods and make sure to call updateCompilerTheme() from them
private void darkthemeitem() {
    try {
        FlatMacDarkLaf.setup();
        UIManager.setLookAndFeel(new FlatMacDarkLaf());
        updateUI();
        updateCompilerTheme(); // Add this line
    } catch (UnsupportedLookAndFeelException e) {
        e.printStackTrace();
    }
}

private void lightthemeitem() {
    try {
        FlatMacLightLaf.setup();
        UIManager.setLookAndFeel(new FlatMacLightLaf());
        updateUI();
        updateCompilerTheme(); // Add this line
    } catch (UnsupportedLookAndFeelException e) {
        e.printStackTrace();
    }
}

    // Update compiler tab theme when theme changes
    // Call this method from your existing darkthemeitem() and lightthemeitem() methods
    // after calling updateUI()
     private void setupTable() {
        // All the code for creating the JTable, model, sorter, and listeners goes
        // here...
    }

    private JPanel createControlPanel() {
        // All the code for creating the JButtons, JTextField, and adding their
        // listeners goes here...
        // This method returns the fully assembled panel of controls.
        JPanel controlPanel = new JPanel();
        // ... add all buttons and the search field ...
        return controlPanel;
    }

    private JPanel createBottomPanel(JPanel controlPanel) {
        // All the code for creating the bottomPanel and totalRowPanel goes here...
        JPanel bottomPanel = new JPanel(new BorderLayout());
        // ... add totalRowPanel and controlPanel ...
        return bottomPanel;
    }
private Timer debounceTimer;
    private void initUI() {
    // ========== DISABLE ANIMATIONS BEFORE SETTING LAF ==========
    System.setProperty("flatlaf.animation", "false");
    System.setProperty("flatlaf.useSnapshoting", "false");
    
    // Disable specific animations
    UIManager.put("TabbedPane.animate", Boolean.FALSE);
    UIManager.put("Component.animate", Boolean.FALSE);
    UIManager.put("ScrollBar.animate", Boolean.FALSE);
    UIManager.put("Button.animate", Boolean.FALSE);
    UIManager.put("Menu.animate", Boolean.FALSE);

    // Apply a modern FlatLaf theme (dark by default)
     try {
        // If using setup method:
        FlatMacDarkLaf.setup(); // or FlatMacLightLaf.setup()
        UIManager.setLookAndFeel(new FlatMacDarkLaf());
        
        // Double-disable after setup
        UIManager.put("TabbedPane.animate", Boolean.FALSE);
        UIManager.put("Component.animate", Boolean.FALSE);
    } catch (UnsupportedLookAndFeelException e) {
        e.printStackTrace();
    }

    setTitle("Text Beans");
    setSize(1000, 700);
    setLocationRelativeTo(null);
    setDefaultCloseOperation(EXIT_ON_CLOSE);
    setLayout(new BorderLayout());
    setIconImage(new ImageIcon(getClass().getResource("/logo.png")).getImage());

    // Setup a tabbed editor area
    JTabbedPane tabbedPane = new JTabbedPane();
    tabbedPane.setTabLayoutPolicy(JTabbedPane.SCROLL_TAB_LAYOUT);
    
    // Additional tab animation disabling
    tabbedPane.putClientProperty("JTabbedPane.animate", Boolean.FALSE);
    
    add(tabbedPane, BorderLayout.CENTER);

    // Create the first tab with a scrollable text area and line numbers
    JPanel editorPanel = new JPanel(new BorderLayout());
    textArea = new JTextArea();
    textArea.setFont(new Font("Helvetica Neue", Font.PLAIN, 16));
    textArea.setLineWrap(true);
    textArea.setWrapStyleWord(true);
    textArea.setMargin(new Insets(8, 8, 8, 8));
    textArea.setBackground(new Color(40, 44, 52));
    textArea.setForeground(Color.WHITE);
    textArea.setCaretColor(Color.CYAN);

    // Line numbers
    JTextArea lineNumbers = new JTextArea("1");
    lineNumbers.setBackground(new Color(30, 34, 40));
    lineNumbers.setForeground(Color.GRAY);
    lineNumbers.setEditable(false);
    lineNumbers.setFont(textArea.getFont());
    lineNumbers.setMargin(new Insets(8, 4, 8, 4));

    // Auto update line numbers on typing
    textArea.getDocument().addDocumentListener(new DocumentListener() {
       private void update() {
    String text = textArea.getText();
    int lines = 0;
    int index = -1;
    
    // Count lines manually (faster than split)
    while ((index = text.indexOf('\n', index + 1)) != -1) {
        lines++;
    }
    lines++; // Last line
    
    // Pre-allocate capacity for efficiency
    StringBuilder sb = new StringBuilder(lines * 4);
    for (int i = 1; i <= lines; i++) {
        sb.append(i).append('\n'); // Use char instead of String
    }
    lineNumbers.setText(sb.toString());
}

        public void insertUpdate(DocumentEvent e) {
            scheduleDebouncedUpdate();
        }

        public void removeUpdate(DocumentEvent e) {
            scheduleDebouncedUpdate();
        }

        public void changedUpdate(DocumentEvent e) {
            scheduleDebouncedUpdate();
        }
        
        private void scheduleDebouncedUpdate() {
            if (debounceTimer != null) {
                debounceTimer.stop();
            }
            debounceTimer = new Timer(300, ev -> update()); // 300ms delay
            debounceTimer.setRepeats(false);
            debounceTimer.start();
        }
    });

    JScrollPane scrollPane = new JScrollPane(textArea);
    scrollPane.setRowHeaderView(lineNumbers);
    editorPanel.add(scrollPane, BorderLayout.CENTER);

    tabbedPane.addTab("Untitled", editorPanel);

    // Menu bar
    setJMenuBar(createMenuBar());

    // Optional status bar
    JPanel statusBar = new JPanel(new BorderLayout());
    JLabel statusLabel = new JLabel(" Ready");
    statusLabel.setBorder(BorderFactory.createEmptyBorder(2, 10, 2, 0));
    statusBar.add(statusLabel, BorderLayout.WEST);
    statusBar.setBorder(BorderFactory.createMatteBorder(1, 0, 0, 0, Color.DARK_GRAY));
    statusBar.setBackground(new Color(30, 30, 30));
    statusLabel.setForeground(Color.LIGHT_GRAY);
    add(statusBar, BorderLayout.SOUTH);
}

    private File chooseFile(Component parent, String dialogTitle, FileNameExtensionFilter filter) {
        JFileChooser chooser = new JFileChooser();
        chooser.setDialogTitle(dialogTitle);

        if (filter != null) {
            chooser.setFileFilter(filter);
        }

        int result = chooser.showOpenDialog(parent);

        if (result == JFileChooser.APPROVE_OPTION) {
            return chooser.getSelectedFile(); // Success: return the file
        } else {
            return null; // Canceled: return null
        }
    }
// ADD THIS AS A NEW INNER CLASS OR SEPARATE CLASS
class StyledDocumentPool {
    private final Queue<StyledDocument> pool = new LinkedList<>();
    private final int maxSize;
    
    public StyledDocumentPool(int maxSize) {
        this.maxSize = maxSize;
        // Pre-populate pool
        for (int i = 0; i < Math.min(10, maxSize); i++) {
            pool.offer(new DefaultStyledDocument());
        }
    }
    
    public StyledDocument acquire() {
        StyledDocument doc = pool.poll();
        return doc != null ? doc : new DefaultStyledDocument();
    }
    
    public void release(StyledDocument doc) {
        if (doc != null && pool.size() < maxSize) {
            try {
                // Clear the document for reuse
                doc.remove(0, doc.getLength());
                pool.offer(doc);
            } catch (BadLocationException e) {
                // Document might be empty or invalid, don't add to pool
            }
        }
    }
    
    public int getPoolSize() {
        return pool.size();
    }
}

// ADD THIS FIELD to TextEditor class:
private final StyledDocumentPool documentPool = new StyledDocumentPool(50);
    private void addAlignmentToolbar(JTextPane pane, JScrollPane scroll) {
        JToolBar alignBar = new JToolBar();

        JButton leftBtn = new JButton("Left");
        JButton centerBtn = new JButton("Center");
        JButton rightBtn = new JButton("Right");
        JButton justifyBtn = new JButton("Justify");
        JButton topBtn = new JButton("Top");
        JButton middleBtn = new JButton("Middle");
        JButton bottomBtn = new JButton("Bottom");
        addAlignmentToolbar(pane, scroll);

        // Horizontal alignment actions
        leftBtn.addActionListener(e -> setParagraphAlignment(pane, StyleConstants.ALIGN_LEFT));
        centerBtn.addActionListener(e -> setParagraphAlignment(pane, StyleConstants.ALIGN_CENTER));
        rightBtn.addActionListener(e -> setParagraphAlignment(pane, StyleConstants.ALIGN_RIGHT));
        justifyBtn.addActionListener(e -> setParagraphAlignment(pane, StyleConstants.ALIGN_JUSTIFIED));

        // Vertical alignment actions (scrolls to simulate)
        topBtn.addActionListener(e -> scroll.getVerticalScrollBar().setValue(0));
        middleBtn.addActionListener(e -> {
            JScrollBar bar = scroll.getVerticalScrollBar();
            int max = bar.getMaximum() - bar.getVisibleAmount();
            bar.setValue(max / 2);
        });
        bottomBtn.addActionListener(e -> {
            JScrollBar bar = scroll.getVerticalScrollBar();
            int max = bar.getMaximum() - bar.getVisibleAmount();
            bar.setValue(max);
        });

        alignBar.add(leftBtn);
        alignBar.add(centerBtn);
        alignBar.add(rightBtn);
        alignBar.add(justifyBtn);
        alignBar.addSeparator();
        alignBar.add(topBtn);
        alignBar.add(middleBtn);
        alignBar.add(bottomBtn);

        // Add the toolbar above the scroll pane
        JPanel wrapper = new JPanel(new BorderLayout());
        wrapper.add(alignBar, BorderLayout.NORTH);
        wrapper.add(scroll, BorderLayout.CENTER);

        // Replace the tab component with the new wrapper
        int idx = tabbedPane.indexOfComponent(scroll);
        if (idx != -1) {
            tabbedPane.setComponentAt(idx, wrapper);
            tabTypeMap.put(wrapper, "text");
        }
    }

    public File getCurrentOpenFile() {
        Tab sel = editorTabPane.getSelectionModel().getSelectedItem();
        if (sel == null)
            return null;
        return (File) sel.getUserData(); // may be null if “Untitled”
    }

    // 6) Implement outputConsole() to return your console TextArea:
    public TextArea outputConsole() {
        return consoleArea;
    }

    private void runActiveTabCode() {
        File currentFile = getCurrentOpenFile(); // however your Notepad handles tabs
        if (currentFile == null)
            return;

        TextArea output = outputConsole(); // your log/output panel
        CompilerExecutor.compileAndRun(currentFile, output);
    }

    private ImageIcon openImageEditorDialog(ImageIcon icon) {
        // Simple dialog for resize/rotate (demo)
        JPanel panel = new JPanel(new BorderLayout());
        JLabel imgLabel = new JLabel(icon);
        panel.add(imgLabel, BorderLayout.CENTER);

        JSlider sizeSlider = new JSlider(10, 400, icon.getIconWidth());
        panel.add(sizeSlider, BorderLayout.SOUTH);

        int result = JOptionPane.showConfirmDialog(this, panel, "Resize Image", JOptionPane.OK_CANCEL_OPTION);
        if (result == JOptionPane.OK_OPTION) {
            int newW = sizeSlider.getValue();
            int newH = icon.getIconHeight() * newW / icon.getIconWidth();
            Image scaled = icon.getImage().getScaledInstance(newW, newH, Image.SCALE_SMOOTH);
            return new ImageIcon(scaled);
        }
        return icon;
    }

    private void setParagraphAlignment(JTextPane pane, int alignment) {
        StyledDocument doc = pane.getStyledDocument();
        int start = pane.getSelectionStart();
        int end = pane.getSelectionEnd();
        Element paragraph = doc.getParagraphElement(start);
        int paraStart = paragraph.getStartOffset();
        int paraEnd = end;
        SimpleAttributeSet attrs = new SimpleAttributeSet();
        StyleConstants.setAlignment(attrs, alignment);
        doc.setParagraphAttributes(paraStart, paraEnd - paraStart, attrs, false);
    }

    // Version 1: Open blank editable table

    // Version 2: Load from Excel file
    /**
     * Creates a new spreadsheet tab with a formula engine and rich-text cells.
     * If 'file' is null, it creates a blank sheet.
     * If 'file' is provided, it loads the data from the .xlsx file.
     * 
     * @param file The Excel file to load, or null to create a new sheet.
     */
   private void createSpreadsheetTab(File file) {
    final Map<Point, String> formulaMap = new HashMap<>();
    final boolean[] isRecalculating = { false };

    DefaultTableModel model = new DefaultTableModel() {
        @Override
        public Class<?> getColumnClass(int columnIndex) {
            return DefaultStyledDocument.class;
        }
    };

    // --- LOGIC BRANCH: Load from file OR create a blank sheet ---
    if (file != null && file.exists()) {
        try (FileInputStream fis = new FileInputStream(file);
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0);
            int maxCols = 0;
            for (Row row : sheet) {
                if (row != null)
                    maxCols = Math.max(maxCols, row.getLastCellNum());
            }

            String[] columnNames = new String[maxCols];
            for (int i = 0; i < maxCols; i++)
                columnNames[i] = String.valueOf((char) ('A' + i));
            model.setColumnIdentifiers(columnNames);

            for (Row row : sheet) {
                Vector<Object> rowData = new Vector<>();
                for (int col = 0; col < maxCols; col++) {
                    Cell cell = (row != null) ? row.getCell(col) : null;
                    String text = "";
                    if (cell != null) {
                        switch (cell.getCellType()) {
                            case STRING:
                                text = cell.getStringCellValue();
                                break;
                            case NUMERIC:
                                text = String.valueOf(cell.getNumericCellValue());
                                break;
                            case BOOLEAN:
                                text = String.valueOf(cell.getBooleanCellValue());
                                break;
                            case FORMULA:
                                text = "=" + cell.getCellFormula();
                                break;
                            default:
                                text = "";
                        }
                    }

                    DefaultStyledDocument doc = new DefaultStyledDocument();
                    if (text.startsWith("=")) {
                        formulaMap.put(new Point(model.getRowCount(), col), text);
                    } else {
                        doc.insertString(0, text, null);
                    }
                    rowData.add(doc);
                }
                model.addRow(rowData);
            }
        } catch (Exception e) {
            JOptionPane.showMessageDialog(this,
                    "Failed to read Excel file: " + e.getMessage(),
                    "Error", JOptionPane.ERROR_MESSAGE);
            return;
        }
    } else {
        // --- CREATE BLANK SHEET ---
        int defaultCols = 26;
        int defaultRows = 100;

        String[] columnNames = new String[defaultCols];
        for (int i = 0; i < defaultCols; i++)
            columnNames[i] = String.valueOf((char) ('A' + i));
        model.setColumnIdentifiers(columnNames);

        model.setRowCount(defaultRows);
        for (int row = 0; row < defaultRows; row++) {
            for (int col = 0; col < defaultCols; col++) {
                model.setValueAt(new DefaultStyledDocument(), row, col);
            }
        }
    }

    // --- TABLE SETUP ---
    JTable table = new JTable(model);
    table.setDefaultRenderer(DefaultStyledDocument.class, new StyledCellRenderer());
    table.setDefaultEditor(DefaultStyledDocument.class, new StyledCellEditor() {
        @Override
        public Component getTableCellEditorComponent(JTable table, Object value, boolean isSelected, int row,
                                                     int column) {
            Point cellPoint = new Point(row, column);
            if (formulaMap.containsKey(cellPoint)) {
                DefaultStyledDocument formulaDoc = new DefaultStyledDocument();
                try {
                    formulaDoc.insertString(0, formulaMap.get(cellPoint), null);
                } catch (BadLocationException e) {}
                return super.getTableCellEditorComponent(table, formulaDoc, isSelected, row, column);
            }
            return super.getTableCellEditorComponent(table, value, isSelected, row, column);
        }
    });

    table.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);
    table.setRowHeight(24);
    table.setGridColor(Color.LIGHT_GRAY);

    // --- FORMULA RE-EVALUATION LISTENER ---
    model.addTableModelListener(e -> {
        if (isRecalculating[0] || e.getType() != TableModelEvent.UPDATE)
            return;
        int row = e.getFirstRow();
        int col = e.getColumn();
        if (row < 0 || col < 0)
            return;

        Object valueObj = model.getValueAt(row, col);
        String textValue = "";
        if (valueObj instanceof DefaultStyledDocument) {
            try {
                textValue = ((DefaultStyledDocument) valueObj)
                        .getText(0, ((DefaultStyledDocument) valueObj).getLength());
            } catch (BadLocationException ex) {}
        }

        Point cellPoint = new Point(row, col);
        if (textValue.startsWith("=")) {
            formulaMap.put(cellPoint, textValue);
        } else {
            formulaMap.remove(cellPoint);
        }
        recalculateAllFormulas(model, formulaMap, isRecalculating);
    });

    // --- INITIAL FORMULA EVALUATION ---
    recalculateAllFormulas(model, formulaMap, isRecalculating);

    // --- ADD TO TABBED PANE ---
    JScrollPane scrollPane = new JScrollPane(table);
    String tabTitle = (file != null) ? file.getName() : "Untitled Sheet";

    int index = tabbedPane.getTabCount();
    tabbedPane.addTab(tabTitle, XlsxEditorPanel.getTabIcon(), scrollPane);

    // --- ADD CLOSE BUTTON TO TAB ---
    JPanel tabHeader = new JPanel(new FlowLayout(FlowLayout.LEFT, 5, 0));
    tabHeader.setOpaque(false);

    JLabel titleLabel = new JLabel(tabTitle);
    JButton closeButton = new JButton("×");
    closeButton.setMargin(new Insets(0, 4, 0, 4));
    closeButton.setBorderPainted(false);
    closeButton.setContentAreaFilled(false);
    closeButton.setFocusable(false);
    closeButton.setForeground(Color.GRAY);

    closeButton.addActionListener(e -> {
        int tabIndex = tabbedPane.indexOfComponent(scrollPane);
        if (tabIndex >= 0) {
            tabbedPane.remove(tabIndex);
        }
    });

    tabHeader.add(titleLabel);
    tabHeader.add(closeButton);
    tabbedPane.setTabComponentAt(index, tabHeader);

    tabbedPane.setSelectedComponent(scrollPane);
    tabTypeMap.put(scrollPane, "table");
}


    private void addTableToTab(String title, DefaultTableModel model) {
        JTable table = new JTable(model) {
            @Override
            public boolean isCellEditable(int row, int column) {
                return true;
            }

            @Override
            public Class<?> getColumnClass(int column) {
                return StyledDocument.class;
            }
        };

        table.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);
        table.setGridColor(Color.GRAY);
        table.setShowGrid(true);
        table.setRowHeight(24);
        table.setRowSelectionAllowed(true);
        table.setCellSelectionEnabled(true);

        table.setDefaultRenderer(StyledDocument.class, new StyledCellRenderer());
        table.setDefaultEditor(StyledDocument.class, new StyledCellEditor());

        JScrollPane scroll = new JScrollPane(table);

        // Load Excel icon
        ImageIcon icon = null;
        try {
            ImageIcon originalIcon = new ImageIcon(getClass().getResource("/icons/excel.png"));
            Image scaledImage = originalIcon.getImage().getScaledInstance(16, 16, Image.SCALE_SMOOTH);
            icon = new ImageIcon(scaledImage);
        } catch (Exception e) {
            icon = null; // Fallback
        }

        tabbedPane.addTab(title, icon, scroll);
        tabbedPane.setSelectedComponent(scroll);
        tabTypeMap.put(scroll, "table");
    }

    private void showCurrencyConverterPopup() {
        SwingUtilities.invokeLater(() -> {
            try {
                String[] currencies = {
                        "USD", "EUR", "GBP", "INR", "JPY", "CNY", "AUD", "CAD", "CHF", "SGD", "NZD", "ZAR", "RUB",
                        "BRL", "MXN",
                        "HKD", "SEK", "NOK", "DKK", "PLN", "TRY", "KRW", "IDR", "MYR", "THB", "PHP", "CZK", "HUF",
                        "ILS", "SAR",
                        "AED", "COP", "CLP", "EGP", "PKR", "BDT", "VND", "NGN", "UAH", "KZT", "QAR", "KWD", "OMR",
                        "BHD", "LKR",
                        "MAD", "DZD", "TND", "JOD", "LBP", "KES", "TZS", "UGX", "GHS", "XOF", "XAF", "CDF", "SDG",
                        "ETB", "AOA",
                        "MZN", "ZMW", "GEL", "BYN", "AZN", "UZS", "TJS", "AFN", "MNT", "AMD", "MKD", "ALL", "ISK",
                        "HRK", "BGN",
                        "RON", "SRD", "BWP", "NAD", "MUR", "SCR", "BBD", "TTD", "JMD", "BSD", "BZD", "KYD", "ANG",
                        "AWG", "HTG",
                        "DOP", "HNL", "GTQ", "NIO", "CRC", "PAB", "PYG", "UYU", "BOB", "PEN", "SVC", "BMD"
                };

                // Example: Only a few static rates for demo. For 100+ currencies, use an API or
                // fill in more rates.
                Map<String, Double> rates = new HashMap<>();
                rates.put("USD_EUR", 0.93);
                rates.put("EUR_USD", 1.08);
                rates.put("USD_INR", 83.2);
                rates.put("INR_USD", 0.012);
                rates.put("USD_GBP", 0.79);
                rates.put("GBP_USD", 1.27);
                rates.put("USD_JPY", 156.7);
                rates.put("JPY_USD", 0.0064);
                rates.put("USD_CNY", 7.24);
                rates.put("CNY_USD", 0.14);
                rates.put("USD_AUD", 1.51);
                rates.put("AUD_USD", 0.66);
                // ...existing rates...
                rates.put("USD_CAD", 1.37);
                rates.put("CAD_USD", 0.73);
                rates.put("USD_CHF", 0.90);
                rates.put("CHF_USD", 1.11);
                rates.put("USD_SGD", 1.35);
                rates.put("SGD_USD", 0.74);
                rates.put("USD_NZD", 1.62);
                rates.put("NZD_USD", 0.62);
                rates.put("USD_ZAR", 18.0);
                rates.put("ZAR_USD", 0.056);
                rates.put("USD_RUB", 89.0);
                rates.put("RUB_USD", 0.011);
                rates.put("USD_BRL", 5.3);
                rates.put("BRL_USD", 0.19);
                rates.put("USD_MXN", 18.2);
                rates.put("MXN_USD", 0.055);
                rates.put("USD_HKD", 7.85);
                rates.put("HKD_USD", 0.13);
                rates.put("USD_SEK", 10.5);
                rates.put("SEK_USD", 0.095);
                rates.put("USD_NOK", 10.7);
                rates.put("NOK_USD", 0.093);
                rates.put("USD_DKK", 6.9);
                rates.put("DKK_USD", 0.14);
                rates.put("USD_PLN", 4.0);
                rates.put("PLN_USD", 0.25);
                rates.put("USD_TRY", 32.5);
                rates.put("TRY_USD", 0.031);
                rates.put("USD_KRW", 1370.0);
                rates.put("KRW_USD", 0.00073);
                rates.put("USD_IDR", 16200.0);
                rates.put("IDR_USD", 0.000062);
                rates.put("USD_MYR", 4.7);
                rates.put("MYR_USD", 0.21);
                rates.put("USD_THB", 36.5);
                rates.put("THB_USD", 0.027);
                rates.put("USD_PHP", 58.5);
                rates.put("PHP_USD", 0.017);
                rates.put("USD_CZK", 23.0);
                rates.put("CZK_USD", 0.043);
                rates.put("USD_HUF", 355.0);
                rates.put("HUF_USD", 0.0028);
                rates.put("USD_ILS", 3.7);
                rates.put("ILS_USD", 0.27);
                rates.put("USD_SAR", 3.75);
                rates.put("SAR_USD", 0.27);
                rates.put("USD_AED", 3.67);
                rates.put("AED_USD", 0.27);
                rates.put("USD_COP", 4100.0);
                rates.put("COP_USD", 0.00024);
                rates.put("USD_CLP", 900.0);
                rates.put("CLP_USD", 0.0011);
                rates.put("USD_EGP", 47.0);
                rates.put("EGP_USD", 0.021);
                rates.put("USD_PKR", 278.0);
                rates.put("PKR_USD", 0.0036);
                rates.put("USD_BDT", 117.0);
                rates.put("BDT_USD", 0.0085);
                rates.put("USD_VND", 25500.0);
                rates.put("VND_USD", 0.000039);
                rates.put("USD_NGN", 1500.0);
                rates.put("NGN_USD", 0.00067);
                rates.put("USD_UAH", 40.0);
                rates.put("UAH_USD", 0.025);
                rates.put("USD_KZT", 450.0);
                rates.put("KZT_USD", 0.0022);
                rates.put("USD_QAR", 3.64);
                rates.put("QAR_USD", 0.27);
                rates.put("USD_KWD", 0.31);
                rates.put("KWD_USD", 3.23);
                rates.put("USD_OMR", 0.38);
                rates.put("OMR_USD", 2.63);
                rates.put("USD_BHD", 0.38);
                rates.put("BHD_USD", 2.65);
                rates.put("USD_LKR", 300.0);
                rates.put("LKR_USD", 0.0033);
                rates.put("USD_MAD", 10.0);
                rates.put("MAD_USD", 0.10);
                rates.put("USD_DZD", 135.0);
                rates.put("DZD_USD", 0.0074);
                rates.put("USD_TND", 3.1);
                rates.put("TND_USD", 0.32);
                rates.put("USD_JOD", 0.71);
                rates.put("JOD_USD", 1.41);
                rates.put("USD_KES", 130.0);
                rates.put("KES_USD", 0.0077);
                rates.put("USD_TZS", 2550.0);
                rates.put("TZS_USD", 0.00039);
                rates.put("USD_UGX", 3800.0);
                rates.put("UGX_USD", 0.00026);
                rates.put("USD_GHS", 15.0);
                rates.put("GHS_USD", 0.067);
                rates.put("USD_XOF", 600.0);
                rates.put("XOF_USD", 0.0017);
                rates.put("USD_XAF", 600.0);
                rates.put("XAF_USD", 0.0017);
                rates.put("USD_CDF", 2700.0);
                rates.put("CDF_USD", 0.00037);
                rates.put("USD_SDG", 600.0);
                rates.put("SDG_USD", 0.0017);
                rates.put("USD_ETB", 57.0);
                rates.put("ETB_USD", 0.018);
                rates.put("USD_AOA", 850.0);
                rates.put("AOA_USD", 0.0012);
                rates.put("USD_MZN", 64.0);
                rates.put("MZN_USD", 0.016);
                rates.put("USD_ZMW", 25.0);
                rates.put("ZMW_USD", 0.040);
                rates.put("USD_GEL", 2.7);
                rates.put("GEL_USD", 0.37);
                rates.put("USD_BYN", 3.2);
                rates.put("BYN_USD", 0.31);
                rates.put("USD_AZN", 1.7);
                rates.put("AZN_USD", 0.59);
                rates.put("USD_UZS", 12600.0);
                rates.put("UZS_USD", 0.000079);
                rates.put("USD_TJS", 11.0);
                rates.put("TJS_USD", 0.091);
                rates.put("USD_AFN", 71.0);
                rates.put("AFN_USD", 0.014);
                rates.put("USD_MNT", 3400.0);
                rates.put("MNT_USD", 0.00029);
                rates.put("USD_AMD", 390.0);
                rates.put("AMD_USD", 0.0026);
                rates.put("USD_MKD", 57.0);
                rates.put("MKD_USD", 0.018);
                rates.put("USD_ALL", 94.0);
                rates.put("ALL_USD", 0.011);
                rates.put("USD_ISK", 140.0);
                rates.put("ISK_USD", 0.0071);
                rates.put("USD_HRK", 7.5);
                rates.put("HRK_USD", 0.13);
                rates.put("USD_BGN", 1.8);
                rates.put("BGN_USD", 0.56);
                rates.put("USD_RON", 4.6);
                rates.put("RON_USD", 0.22);
                rates.put("USD_SRD", 33.0);
                rates.put("SRD_USD", 0.030);
                rates.put("USD_BWP", 13.5);
                rates.put("BWP_USD", 0.074);
                rates.put("USD_NAD", 18.0);
                rates.put("NAD_USD", 0.056);
                rates.put("USD_MUR", 46.0);
                rates.put("MUR_USD", 0.022);
                rates.put("USD_SCR", 14.0);
                rates.put("SCR_USD", 0.071);
                rates.put("USD_BBD", 2.0);
                rates.put("BBD_USD", 0.50);
                rates.put("USD_TTD", 6.8);
                rates.put("TTD_USD", 0.15);
                rates.put("USD_JMD", 155.0);
                rates.put("JMD_USD", 0.0065);
                rates.put("USD_BSD", 1.0);
                rates.put("BSD_USD", 1.0);
                rates.put("USD_BZD", 2.0);
                rates.put("BZD_USD", 0.50);
                rates.put("USD_KYD", 0.83);
                rates.put("KYD_USD", 1.20);
                rates.put("USD_ANG", 1.8);
                rates.put("ANG_USD", 0.56);
                rates.put("USD_AWG", 1.8);
                rates.put("AWG_USD", 0.56);
                rates.put("USD_HTG", 132.0);
                rates.put("HTG_USD", 0.0076);
                rates.put("USD_DOP", 59.0);
                rates.put("DOP_USD", 0.017);
                rates.put("USD_HNL", 24.5);
                rates.put("HNL_USD", 0.041);
                rates.put("USD_GTQ", 7.8);
                rates.put("GTQ_USD", 0.13);
                rates.put("USD_NIO", 36.5);
                rates.put("NIO_USD", 0.027);
                rates.put("USD_CRC", 520.0);
                rates.put("CRC_USD", 0.0019);
                rates.put("USD_PAB", 1.0);
                rates.put("PAB_USD", 1.0);
                rates.put("USD_PYG", 7300.0);
                rates.put("PYG_USD", 0.00014);
                rates.put("USD_UYU", 39.0);
                rates.put("UYU_USD", 0.026);
                rates.put("USD_BOB", 6.9);
                rates.put("BOB_USD", 0.14);
                rates.put("USD_PEN", 3.7);
                rates.put("PEN_USD", 0.27);
                rates.put("USD_SVC", 8.75);
                rates.put("SVC_USD", 0.11);
                rates.put("USD_BMD", 1.0);
                rates.put("BMD_USD", 1.0);
                rates.put("USD_LBP", 89500.0);
                rates.put("LBP_USD", 0.000011);
                rates.put("USD_MNT", 3400.0);
                rates.put("MNT_USD", 0.00029);
                rates.put("USD_MKD", 57.0);
                rates.put("MKD_USD", 0.018);
                rates.put("USD_ALL", 94.0);
                rates.put("ALL_USD", 0.011);
                rates.put("USD_ISK", 140.0);
                rates.put("ISK_USD", 0.0071);
                rates.put("USD_HRK", 7.5);
                rates.put("HRK_USD", 0.13);
                rates.put("USD_BGN", 1.8);
                rates.put("BGN_USD", 0.56);
                rates.put("USD_RON", 4.6);
                rates.put("RON_USD", 0.22);
                // ...add more as needed...
                // ...add more rates as needed...

                JPanel panel = new JPanel(new FlowLayout());
                JTextField amountField = new JTextField(8);
                JComboBox<String> fromBox = new JComboBox<>(currencies);
                JComboBox<String> toBox = new JComboBox<>(currencies);
                JLabel resultLabel = new JLabel("Result: ");

                JButton convertBtn = new JButton("Convert");
                convertBtn.addActionListener(ev -> {
                    try {
                        double amount = Double.parseDouble(amountField.getText());
                        String from = (String) fromBox.getSelectedItem();
                        String to = (String) toBox.getSelectedItem();
                        double converted = amount;
                        if (!from.equals(to)) {
                            Double rate = rates.get(from + "_" + to);
                            if (rate == null) {
                                resultLabel.setText("No rate for " + from + " to " + to);
                                return;
                            }
                            converted = amount * rate;
                        }
                        resultLabel.setText("Result: " + String.format("%.2f", converted) + " " + to);
                    } catch (Exception ex) {
                        resultLabel.setText("Invalid input.");
                    }
                });

                panel.add(new JLabel("Amount:"));
                panel.add(amountField);
                panel.add(new JLabel("From:"));
                panel.add(fromBox);
                panel.add(new JLabel("To:"));
                panel.add(toBox);
                panel.add(convertBtn);
                panel.add(resultLabel);

                JDialog dialog = new JDialog(this, "Currency Converter ", false);
                dialog.setDefaultCloseOperation(JDialog.DISPOSE_ON_CLOSE);
                dialog.setContentPane(panel);
                dialog.pack();
                dialog.setLocationRelativeTo(this);
                dialog.setVisible(true);
            } catch (Exception ex) {
                ex.printStackTrace();
                JOptionPane.showMessageDialog(this,
                        "Error opening currency converter: " + ex.getMessage(),
                        "Error",
                        JOptionPane.ERROR_MESSAGE);
            }
        });
    }

    private void encodeBase64() {
        JTextPane pane = getCurrentTextPane();
        if (pane != null) {
            String text = pane.getText();
            String encoded = Base64.getEncoder().encodeToString(text.getBytes(StandardCharsets.UTF_8));
            pane.setText(encoded);
        }
    }

    private void decodeBase64() {
        JTextPane pane = getCurrentTextPane();
        if (pane != null) {
            String text = pane.getText().replaceAll("\\s+", ""); // Remove newlines/spaces
            try {
                byte[] decodedBytes = Base64.getDecoder().decode(text);

                // Check if the result is UTF-8 printable
                String decoded = new String(decodedBytes, StandardCharsets.UTF_8);

                // Optional: verify text characters only (basic validation)
                if (!decoded.chars().allMatch(c -> c >= 32 || c == '\n' || c == '\r')) {
                    throw new IllegalArgumentException("Not a valid text format.");
                }

                pane.setText(decoded);
            } catch (IllegalArgumentException e) {
                JOptionPane.showMessageDialog(this, "Invalid or non-text Base64 input.", "Decode Error",
                        JOptionPane.ERROR_MESSAGE);
            }
        }
    }

    private void decodeBase64ToFile() {
        JTextPane pane = getCurrentTextPane();
        if (pane != null) {
            String text = pane.getText().replaceAll("\\s+", ""); // sanitize

            JFileChooser fileChooser = new JFileChooser();
            fileChooser.setDialogTitle("Save decoded file as");
            int result = fileChooser.showSaveDialog(this);

            if (result == JFileChooser.APPROVE_OPTION) {
                File outputFile = fileChooser.getSelectedFile();
                try (FileOutputStream fos = new FileOutputStream(outputFile)) {
                    byte[] decodedBytes = Base64.getDecoder().decode(text);
                    fos.write(decodedBytes);
                    JOptionPane.showMessageDialog(this, "Decoded and saved successfully.");
                } catch (Exception e) {
                    JOptionPane.showMessageDialog(this, "Failed to decode/save: " + e.getMessage(), "Error",
                            JOptionPane.ERROR_MESSAGE);
                }
            }
        }
    }

    private void encodeFileToBase64() {
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setDialogTitle("Choose a file to encode");

        int result = fileChooser.showOpenDialog(this);
        if (result == JFileChooser.APPROVE_OPTION) {
            File selectedFile = fileChooser.getSelectedFile();

            try (FileInputStream fis = new FileInputStream(selectedFile)) {
                byte[] fileBytes = fis.readAllBytes();
                String base64Encoded = Base64.getEncoder().encodeToString(fileBytes);

                // Show in current text pane
                JTextPane pane = getCurrentTextPane();
                if (pane != null) {
                    pane.setText(base64Encoded);
                }

                JOptionPane.showMessageDialog(this, "File encoded successfully.");
            } catch (IOException e) {
                JOptionPane.showMessageDialog(this, "Encoding failed: " + e.getMessage(), "Error",
                        JOptionPane.ERROR_MESSAGE);
            }
        }
    }

    private void savePPTXTab(Component comp, boolean forceDialog) {
        // Ensure the component is a JScrollPane
        if (!(comp instanceof JScrollPane)) {
            JOptionPane.showMessageDialog(this, "Unsupported component format.");
            return;
        }
        JScrollPane scroll = (JScrollPane) comp;

        // Obtain the main panel for the PPTX slides
        Component view = scroll.getViewport().getView();
        if (!(view instanceof JPanel)) {
            JOptionPane.showMessageDialog(this, "Invalid PPTX panel.");
            return;
        }
        JPanel panel = (JPanel) view;

        // Create a file chooser for the PPTX file
        JFileChooser fc = new JFileChooser();
        fc.setDialogTitle("Save Presentation As");
        fc.setSelectedFile(new File("Presentation.pptx"));
        fc.setFileFilter(new javax.swing.filechooser.FileNameExtensionFilter("PowerPoint Files", "pptx"));

        // Always prompt for a file (or use forceDialog logic as needed)
        File file;
        if (forceDialog) {
            if (fc.showSaveDialog(this) != JFileChooser.APPROVE_OPTION) {
                return;
            }
            file = fc.getSelectedFile();
        } else {
            // Since we do not store a PPTX file in fileMap, always show the dialog.
            if (fc.showSaveDialog(this) != JFileChooser.APPROVE_OPTION) {
                return;
            }
            file = fc.getSelectedFile();
        }

        if (file == null) {
            JOptionPane.showMessageDialog(this, "File not selected. Save failed.");
            return;
        }

        // NOTE: Using XWPFDocument writes a Word document (DOCX),
        // not a PowerPoint file. If you want to save a PPTX proper,
        // you need to use XMLSlideShow and related XSLF classes.
        try (FileOutputStream fos = new FileOutputStream(file);
                org.apache.poi.xwpf.usermodel.XWPFDocument document = new org.apache.poi.xwpf.usermodel.XWPFDocument()) {

            // Assume each slide in the PPTX tab is represented as a JTextArea in panel.
            // You might have a more sophisticated mapping; here we simply iterate over
            // panel components.
            int count = panel.getComponentCount();
            for (int i = 0; i < count; i++) {
                Component slideComp = panel.getComponent(i);
                // Check if this component is a JTextArea (the slide’s text area)
                if (slideComp instanceof JTextArea) {
                    JTextArea slide = (JTextArea) slideComp;
                    // Create a new paragraph for this “slide”
                    org.apache.poi.xwpf.usermodel.XWPFParagraph paragraph = document.createParagraph();
                    org.apache.poi.xwpf.usermodel.XWPFRun run = paragraph.createRun();
                    run.setText(slide.getText());
                }
            }
            document.write(fos);
            JOptionPane.showMessageDialog(this, "PPTX saved successfully.");
        } catch (Exception ex) {
            JOptionPane.showMessageDialog(this, "Error saving PPTX:\n" + ex.getMessage());
        }
    }

    private void saveFile() {
        Component comp = tabbedPane.getSelectedComponent();
        if (comp == null) {
            showAnimatedNotification("No active tab to save", new Color(244, 67, 54));
            return;
        }

        String type = tabTypeMap.getOrDefault(comp, "text");

        if ("text".equals(type)) {
            saveTextTab(comp, false);
            showAnimatedNotification("File saved successfully!", new Color(76, 175, 80));
        } else if ("table".equals(type)) {
            saveTableTab(comp, false);
            showAnimatedNotification("Table saved successfully!", new Color(76, 175, 80));
        } else if ("pptx_editable".equals(type)) {
            savePPTXTab(comp, false);
            showAnimatedNotification("Presentation saved!", new Color(76, 175, 80));
        } else {
            showAnimatedNotification("Unsupported tab type for saving", new Color(255, 152, 0));
        }
    }

    private void saveFileAs() {
        Component comp = tabbedPane.getSelectedComponent();
        if (comp == null) {
            JOptionPane.showMessageDialog(this, "No active tab to save.");
            return;
        }

        if (tabTypeMap.containsKey(comp)) {
            String type = tabTypeMap.get(comp);

            if ("text".equals(type)) {
                saveTextTab(comp, true); // Fix: Force file chooser for "Save As"
            } else if ("table".equals(type)) {
                saveTableTab(comp, true); // Fix: Save table as a new file
            } else {
                JOptionPane.showMessageDialog(this, "Unsupported tab type for Save As.");
            }
        } else {
            JOptionPane.showMessageDialog(this, "Tab type could not be determined. Unable to save.");
        }
    }

    // To UPPERCASE for all formats
    // Add this method to your class
    private void showCalculatorPopup() {
        SwingUtilities.invokeLater(() -> {
            try {
                // Use null instead of 'this' if parent might be invalid
                JDialog dialog = new JDialog((Frame) null, "Calculator", false);
                ImageIcon calculatorIcon = loadIcon("/icons/calculator.png", 16, 16);
                if (calculatorIcon != null) {
                    dialog.setIconImage(calculatorIcon.getImage());
                }
                // Or get the active frame:
                // Frame parentFrame = (Frame) SwingUtilities.getWindowAncestor(component);

                JPanel panel = new JPanel(new BorderLayout());
                JTextField display = new JTextField();

                display.setEditable(false);
                display.setFont(new Font("Consolas", Font.BOLD, 28));
                display.setPreferredSize(new Dimension(400, 50));
                panel.add(display, BorderLayout.NORTH);

                // Your existing button creation code...
                String[] buttonLabels = {
                        "1", "2", "3", "/", "4", "5", "6", "*",
                        "7", "8", "9", "-", "0", ".", "=", "+",
                        "Clear", "(", ")", "^", "sqrt", "cbrt", "log", "sin",
                        "cos", "tan", "asin", "acos", "atan", "!", "%", "|x|"
                };

                JPanel buttonPanel = new JPanel(new GridLayout(8, 4, 5, 5));
                for (String label : buttonLabels) {
                    JButton button = new JButton(label);
                    button.setFont(new Font("Consolas", Font.BOLD, 16));
                    button.addActionListener(e -> {
                        String cmd = ((JButton) e.getSource()).getText();
                        String expr = display.getText();
                        switch (cmd) {
                            case "=":
                                try {
                                    double result = new ExpressionParser().parse(expr);
                                    display.setText(Double.toString(result));
                                } catch (Exception ex) {
                                    display.setText("Error");
                                }
                                break;
                            case "Clear":
                                display.setText("");
                                break;
                            default:
                                display.setText(expr + cmd);
                        }
                    });
                    buttonPanel.add(button);
                }
                panel.add(buttonPanel, BorderLayout.CENTER);

                // Keyboard support
                display.addKeyListener(new KeyAdapter() {
                    @Override
                    public void keyTyped(KeyEvent e) {
                        char ch = e.getKeyChar();
                        if (Character.isDigit(ch) || "+-*/().^".indexOf(ch) >= 0) {
                            display.setText(display.getText() + ch);
                            e.consume();
                        }
                    }

                    @Override
                    public void keyPressed(KeyEvent e) {
                        int code = e.getKeyCode();
                        if (code == KeyEvent.VK_ENTER || code == KeyEvent.VK_EQUALS) {
                            try {
                                double result = new ExpressionParser().parse(display.getText());
                                display.setText(Double.toString(result));
                            } catch (Exception ex) {
                                display.setText("Error");
                            }
                            e.consume();
                        } else if (code == KeyEvent.VK_BACK_SPACE) {
                            String text = display.getText();
                            if (!text.isEmpty())
                                display.setText(text.substring(0, text.length() - 1));
                            e.consume();
                        } else if (code == KeyEvent.VK_DELETE) {
                            display.setText("");
                            e.consume();
                        }
                    }
                });

                dialog.setDefaultCloseOperation(JDialog.DISPOSE_ON_CLOSE);
                dialog.setContentPane(panel);
                dialog.pack();
                dialog.setLocationRelativeTo(null); // Center on screen instead of parent
                // Remove fadeInWindow if it's causing issues
                dialog.setVisible(true);
                display.requestFocusInWindow();

            } catch (Exception ex) {
                ex.printStackTrace();
                JOptionPane.showMessageDialog(null, "Error opening calculator: " + ex.getMessage());
            }
        });
    }// Submenu transition animations

    private void animateSubmenuTransition(JMenu fromMenu, JMenu toMenu) {
        if (fromMenu == null || toMenu == null)
            return;

        // Store original properties
        Dimension fromOriginalSize = fromMenu.getPreferredSize();
        Dimension toOriginalSize = toMenu.getPreferredSize();

        // Create fade effect for the transition
        animateMenuFade(fromMenu, toMenu, 200);

        // Animate size transition if menus have different sizes
        if (!fromOriginalSize.equals(toOriginalSize)) {
            animateMenuResize(fromMenu, toMenu, fromOriginalSize, toOriginalSize, 250);
        }
    }

    private void animateMenuFade(JMenu fromMenu, JMenu toMenu, int duration) {
        // Fade out current menu
        Timer fadeTimer = new Timer(16, null);
        final long startTime = System.currentTimeMillis();

        fadeTimer.addActionListener(e -> {
            long currentTime = System.currentTimeMillis();
            float progress = (float) (currentTime - startTime) / duration;
            progress = Math.min(1.0f, progress);

            if (progress <= 0.5f) {
                // First half: fade out current menu
                float alpha = 1.0f - (progress * 2);
                setMenuAlpha(fromMenu, alpha);
            } else {
                // Second half: fade in new menu
                float alpha = (progress - 0.5f) * 2;
                setMenuAlpha(toMenu, alpha);
            }

            if (progress >= 1.0f) {
                fadeTimer.stop();
                // Reset alpha
                setMenuAlpha(fromMenu, 1.0f);
                setMenuAlpha(toMenu, 1.0f);
            }
        });
        fadeTimer.start();
    }

    private void setMenuAlpha(JMenu menu, float alpha) {
        // Create a semi-transparent appearance for the menu
        Color bg = menu.getBackground();
        Color newBg = new Color(bg.getRed(), bg.getGreen(), bg.getBlue(), (int) (alpha * 255));
        menu.setBackground(newBg);
        menu.repaint();
    }

    private void animateMenuResize(JMenu fromMenu, JMenu toMenu, Dimension fromSize, Dimension toSize, int duration) {
        Timer resizeTimer = new Timer(16, null);
        final long startTime = System.currentTimeMillis();

        resizeTimer.addActionListener(e -> {
            long currentTime = System.currentTimeMillis();
            float progress = (float) (currentTime - startTime) / duration;
            progress = Math.min(1.0f, progress);

            // Use ease-out for smooth deceleration
            float easedProgress = 1 - (float) Math.pow(1 - progress, 2);

            // Interpolate width and height
            int currentWidth = (int) (fromSize.width + (toSize.width - fromSize.width) * easedProgress);
            int currentHeight = (int) (fromSize.height + (toSize.height - fromSize.height) * easedProgress);

            // Apply to both menus during transition for smooth effect
            Dimension transitionSize = new Dimension(currentWidth, currentHeight);
            fromMenu.setPreferredSize(transitionSize);
            toMenu.setPreferredSize(transitionSize);

            // Revalidate to apply size changes
            fromMenu.revalidate();
            toMenu.revalidate();

            if (progress >= 1.0f) {
                resizeTimer.stop();
                // Restore original sizes
                fromMenu.setPreferredSize(fromSize);
                toMenu.setPreferredSize(toSize);
                fromMenu.revalidate();
                toMenu.revalidate();
            }
        });
        resizeTimer.start();
    }

    private void filterLines() {
        JTextPane pane = getCurrentTextPane();
        if (pane != null) {
            String text = pane.getText();
            String[] lines = text.split("\\n");
            StringBuilder filtered = new StringBuilder();
            for (String line : lines) {
                if (!line.trim().isEmpty()) {
                    filtered.append(line).append("\n");
                }
            }
            pane.setText(filtered.toString().trim());
        }
    }

    // --- ExpressionParser class for evaluating expressions ---
    private static class ExpressionParser {
        private int pos = -1, ch;
        private String input;

        public double parse(String str) {
            input = str;
            pos = -1;
            nextChar();
            double x = parseExpression();
            if (pos < input.length())
                throw new RuntimeException("Unexpected: " + (char) ch);
            return x;
        }

        private void nextChar() {
            ch = (++pos < input.length()) ? input.charAt(pos) : -1;
        }

        private boolean eat(int charToEat) {
            while (ch == ' ')
                nextChar();
            if (ch == charToEat) {
                nextChar();
                return true;
            }
            return false;
        }

        private double parseExpression() {
            double x = parseTerm();
            while (true) {
                if (eat('+'))
                    x += parseTerm();
                else if (eat('-'))
                    x -= parseTerm();
                else
                    return x;
            }
        }

        private double parseTerm() {
            double x = parseFactor();
            while (true) {
                if (eat('*'))
                    x *= parseFactor();
                else if (eat('/'))
                    x /= parseFactor();
                else if (eat('^'))
                    x = Math.pow(x, parseFactor());
                else
                    return x;
            }
        }

        private double parseFactor() {
            if (eat('+'))
                return parseFactor();
            if (eat('-'))
                return -parseFactor();

            double x;
            int startPos = this.pos;
            if (eat('(')) {
                x = parseExpression();
                eat(')');
            } else if ((ch >= '0' && ch <= '9') || ch == '.') {
                while ((ch >= '0' && ch <= '9') || ch == '.')
                    nextChar();
                x = Double.parseDouble(input.substring(startPos, this.pos));
            } else if (ch >= 'a' && ch <= 'z' || ch == '|') {
                while ((ch >= 'a' && ch <= 'z') || ch == '|')
                    nextChar();
                String func = input.substring(startPos, this.pos);
                x = parseFactor();
                switch (func) {
                    case "sqrt":
                        x = Math.sqrt(x);
                        break;
                    case "cbrt":
                        x = Math.cbrt(x);
                        break;
                    case "log":
                        x = Math.log10(x);
                        break;
                    case "sin":
                        x = Math.sin(Math.toRadians(x));
                        break;
                    case "cos":
                        x = Math.cos(Math.toRadians(x));
                        break;
                    case "tan":
                        x = Math.tan(Math.toRadians(x));
                        break;
                    case "asin":
                        x = Math.toDegrees(Math.asin(x));
                        break;
                    case "acos":
                        x = Math.toDegrees(Math.acos(x));
                        break;
                    case "atan":
                        x = Math.toDegrees(Math.atan(x));
                        break;
                    case "!":
                        x = factorial((int) x);
                        break;
                    case "%":
                        x = x / 100.0;
                        break;
                    case "|x|":
                        x = Math.abs(x);
                        break;
                    default:
                        throw new RuntimeException("Unknown function: " + func);
                }
            } else {
                throw new RuntimeException("Unexpected: " + (char) ch);
            }
            return x;
        }

        private int factorial(int n) {
            if (n < 0)
                throw new ArithmeticException("Negative factorial");
            int fact = 1;
            for (int i = 1; i <= n; i++)
                fact *= i;
            return fact;
        }
    }

    // Evaluator for regular mode
    private double evalBasic(String expr) throws Exception {
        expr = expr.trim().replaceAll("\\s+", "");
        if (expr.isEmpty() ||
                expr.matches(".*[+\\-*/.]$") ||
                expr.matches(".*\\.\\..*")) {
            throw new IllegalArgumentException("Invalid Expression");
        }
        ScriptEngine engine = new ScriptEngineManager().getEngineByName("JavaScript");
        Object result = engine.eval(expr);
        return Double.parseDouble(result.toString());
    }

    public class ZipViewer {
        public static void showZipContents(JFrame parentFrame) {
            JFileChooser chooser = new JFileChooser();
            chooser.setDialogTitle("Select ZIP File");
            chooser.setFileFilter(new FileNameExtensionFilter("ZIP Files", "zip"));

            int result = chooser.showOpenDialog(parentFrame);
            if (result != JFileChooser.APPROVE_OPTION)
                return;

            File zipFile = chooser.getSelectedFile();
            StringBuilder content = new StringBuilder();

            try (ZipInputStream zis = new ZipInputStream(new FileInputStream(zipFile))) {
                ZipEntry entry;
                while ((entry = zis.getNextEntry()) != null) {
                    content.append(entry.getName()).append(entry.isDirectory() ? " [DIR]" : "").append("\n");
                }
            } catch (IOException e) {
                JOptionPane.showMessageDialog(parentFrame, "Error reading ZIP file: " + e.getMessage());
                return;
            }

            JTextArea textArea = new JTextArea(content.toString());
            textArea.setEditable(false);
            JScrollPane scrollPane = new JScrollPane(textArea);

            JDialog dialog = new JDialog(parentFrame, "ZIP Preview: " + zipFile.getName(), true);
            dialog.getContentPane().add(scrollPane);
            dialog.setSize(400, 300);
            dialog.setLocationRelativeTo(parentFrame);
            dialog.setVisible(true);
        }

    }

    class RoundedBorder extends AbstractBorder {
        private final int radius;

        public RoundedBorder(int radius) {
            this.radius = radius;
        }

        @Override
        public void paintBorder(Component c, Graphics g, int x, int y, int width, int height) {
            Graphics2D g2 = (Graphics2D) g.create();
            g2.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
            g2.setColor(Color.GRAY);
            g2.drawRoundRect(x, y, width - 1, height - 1, radius, radius);
            g2.dispose();
        }

        @Override
        public Insets getBorderInsets(Component c) {
            return new Insets(radius, radius, radius, radius);
        }

        @Override
        public Insets getBorderInsets(Component c, Insets insets) {
            if (insets == null) {
                insets = new Insets(0, 0, 0, 0);
            }
            insets.top = radius;
            insets.left = radius;
            insets.bottom = radius;
            insets.right = radius;
            return insets;
        }

        @Override
        public boolean isBorderOpaque() {
            return false;
        }
    }// Evaluator for scientific mode

    private double evalScientific(String expr) throws Exception {
        expr = expr.replaceAll("π", String.valueOf(Math.PI));
        expr = expr.replaceAll("e", String.valueOf(Math.E));
        expr = expr.replaceAll("(\\d+(?:\\.\\d+)?)\\^2", "Math.pow($1,2)");
        expr = expr.replaceAll("(\\d+(?:\\.\\d+)?)\\^(\\d+(?:\\.\\d+)?)", "Math.pow($1,$2)");
        expr = expr.replaceAll("sqrt\\(", "Math.sqrt(")
                .replaceAll("sin\\(", "Math.sin(")
                .replaceAll("cos\\(", "Math.cos(")
                .replaceAll("tan\\(", "Math.tan(")
                .replaceAll("log10\\(", "Math.log10(")
                .replaceAll("log\\(", "Math.log(")
                .replaceAll("abs\\(", "Math.abs(")
                .replaceAll("exp\\(", "Math.exp(")
                .replaceAll("floor\\(", "Math.floor(")
                .replaceAll("ceil\\(", "Math.ceil(");
        expr = expr.replaceAll("%", "/100.0");
        ScriptEngine engine = new ScriptEngineManager().getEngineByName("JavaScript");
        Object result = engine.eval(expr);
        return Double.parseDouble(result.toString());
    }

    // A simple scientific expression evaluator using JavaScript engine

    private double eval(String expr) {
        javax.script.ScriptEngine engine = new javax.script.ScriptEngineManager().getEngineByName("JavaScript");
        try {
            Object result = engine.eval(expr);
            return Double.parseDouble(result.toString());
        } catch (Exception ex) {
            return 0;
        }
    }
    // Pressing Enter triggers search too

    private JPanel createSearchPanel() {
        searchPanel = new JPanel(new FlowLayout(FlowLayout.RIGHT));
        searchPanel.setBorder(BorderFactory.createEmptyBorder(5, 5, 5, 5));
        searchPanel.setBackground(UIManager.getColor("Panel.background"));

        searchField = new JTextField(20);
        searchField.putClientProperty("JTextField.placeholderText", "🔍 Search...");
        searchField.putClientProperty("JComponent.roundRect", true);
        searchField.setPreferredSize(new Dimension(250, 30));
        UIManager.put("TextComponent.arc", 20);

        searchField.getDocument().addDocumentListener(new DocumentListener() {
            public void insertUpdate(DocumentEvent e) {
                performSearch();
            }

            public void removeUpdate(DocumentEvent e) {
                performSearch();
            }

            public void changedUpdate(DocumentEvent e) {
                performSearch();
            }
        });

        prevButton = new JButton("⏮");
        nextButton = new JButton("⏭");

        prevButton.addActionListener(e -> highlightMatch(false));
        nextButton.addActionListener(e -> highlightMatch(true));

        searchPanel.add(searchField);
        searchPanel.add(prevButton);
        searchPanel.add(nextButton);

        return searchPanel;
    }

    // 3. Show floating panel
    private void showFloatingSearchPanel() {
        JDialog floatingSearch = new JDialog(this, false);
        floatingSearch.setUndecorated(true);
        floatingSearch.setLayout(new BorderLayout());
        floatingSearch.add(createSearchPanel(), BorderLayout.CENTER);
        floatingSearch.pack();

        // Position top-right of parent
        Point parentLocation = this.getLocationOnScreen();
        int x = parentLocation.x + getWidth() - floatingSearch.getWidth() - 20;
        int y = parentLocation.y + 60;
        floatingSearch.setLocation(x, y);

        floatingSearch.setAlwaysOnTop(true);
        floatingSearch.setVisible(true);

        SwingUtilities.invokeLater(() -> searchField.requestFocusInWindow());

        // Close with ESC
        floatingSearch.getRootPane().registerKeyboardAction(e -> floatingSearch.dispose(),
                KeyStroke.getKeyStroke(KeyEvent.VK_ESCAPE, 0),
                JComponent.WHEN_IN_FOCUSED_WINDOW);
    }

    // 4. Search logic
 private void performSearch() {
    matchIndices.clear();
    currentMatch = -1;
    textArea.getHighlighter().removeAllHighlights();

    String keyword = searchField.getText().trim();
    if (keyword.isEmpty()) return;

    String content = textArea.getText();
    
    // Use Boyer-Moore algorithm for O(n/m) search
    matchIndices = boyerMooreSearch(content, keyword);
    
    Highlighter.HighlightPainter painter = new DefaultHighlighter.DefaultHighlightPainter(Color.YELLOW);
    
    for (int index : matchIndices) {
        try {
            textArea.getHighlighter().addHighlight(index, index + keyword.length(), painter);
        } catch (BadLocationException e) {
            e.printStackTrace();
        }
    }

    if (!matchIndices.isEmpty()) {
        currentMatch = 0;
        scrollToMatch(currentMatch, keyword.length());
    }
}

// ADD THIS NEW METHOD:
private List<Integer> boyerMooreSearch(String text, String pattern) {
    List<Integer> indices = new ArrayList<>();
    if (pattern.isEmpty()) return indices;
    
    // Preprocessing
    int[] badChar = new int[256];
    Arrays.fill(badChar, -1);
    for (int i = 0; i < pattern.length(); i++) {
        badChar[pattern.charAt(i)] = i;
    }
    
    int s = 0;
    while (s <= text.length() - pattern.length()) {
        int j = pattern.length() - 1;
        
        while (j >= 0 && pattern.charAt(j) == text.charAt(s + j)) {
            j--;
        }
        
        if (j < 0) {
            indices.add(s);
            s += (s + pattern.length() < text.length()) ? 
                 pattern.length() - badChar[text.charAt(s + pattern.length())] : 1;
        } else {
            s += Math.max(1, j - badChar[text.charAt(s + j)]);
        }
    }
    return indices;
}
    private void highlightMatch(boolean forward) {
        if (matchIndices.isEmpty())
            return;

        currentMatch += forward ? 1 : -1;
        if (currentMatch >= matchIndices.size())
            currentMatch = 0;
        if (currentMatch < 0)
            currentMatch = matchIndices.size() - 1;

        int start = matchIndices.get(currentMatch);
        scrollToMatch(currentMatch, searchField.getText().length());
    }

    private void scrollToMatch(int index, int length) {
        try {
            textArea.setCaretPosition(matchIndices.get(index) + length);
            textArea.moveCaretPosition(matchIndices.get(index));
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

    // 5. Add this in your initUI() or constructor to bind Ctrl+F
    private void bindSearchShortcut() {
        KeyStroke ctrlF = KeyStroke.getKeyStroke(KeyEvent.VK_F, InputEvent.CTRL_DOWN_MASK);
        getRootPane().getInputMap(JComponent.WHEN_IN_FOCUSED_WINDOW).put(ctrlF, "showSearch");
        getRootPane().getActionMap().put("showSearch", new AbstractAction() {
            public void actionPerformed(ActionEvent e) {
                showFloatingSearchPanel();
            }
        });
    }

    public class SyntaxTextPane extends JTextPane {
        private static final String[] KEYWORDS = {
                "int", "float", "double", "if", "else", "while", "for", "return", "import", "public",
                "private", "protected", "class", "static", "void", "new", "try", "catch", "String", "final"
        };

        private final Style keywordStyle;
        private final Style normalStyle;

        public SyntaxTextPane() {
            setFont(new Font("Consolas", Font.PLAIN, 14));
            StyledDocument doc = getStyledDocument();

            // Create and apply styles
            Style defaultStyle = StyleContext.getDefaultStyleContext().getStyle(StyleContext.DEFAULT_STYLE);
            keywordStyle = doc.addStyle("KeywordStyle", defaultStyle);
            StyleConstants.setForeground(keywordStyle, Color.BLUE);

            normalStyle = doc.addStyle("NormalStyle", defaultStyle);
            StyleConstants.setForeground(normalStyle, Color.BLACK);

            // Add document listener
            getDocument().addDocumentListener(new javax.swing.event.DocumentListener() {
                public void insertUpdate(javax.swing.event.DocumentEvent e) {
                    SwingUtilities.invokeLater(() -> highlightSyntax());
                }

                public void removeUpdate(javax.swing.event.DocumentEvent e) {
                    SwingUtilities.invokeLater(() -> highlightSyntax());
                }

                public void changedUpdate(javax.swing.event.DocumentEvent e) {
                    SwingUtilities.invokeLater(() -> highlightSyntax());
                }
            });

        }

        private void highlightSyntax() {
            AbstractDocument doc = (AbstractDocument) getDocument(); // fix here

            // Temporarily remove listeners to avoid recursive updates
            DocumentListener[] listeners = doc.getListeners(DocumentListener.class);
            for (DocumentListener l : listeners) {
                doc.removeDocumentListener(l);
            }

            StyledDocument styledDoc = getStyledDocument();
            String text;
            try {
                text = styledDoc.getText(0, styledDoc.getLength());
            } catch (BadLocationException e) {
                e.printStackTrace();
                return;
            }

            // Reset to normal style
            styledDoc.setCharacterAttributes(0, text.length(), normalStyle, true);

            // Highlight keywords
            for (String keyword : KEYWORDS) {
                Pattern pattern = Pattern.compile("\\b" + Pattern.quote(keyword) + "\\b");
                Matcher matcher = pattern.matcher(text);
                while (matcher.find()) {
                    styledDoc.setCharacterAttributes(matcher.start(), matcher.end() - matcher.start(), keywordStyle,
                            false);
                }
            }

            // Re-attach the listeners
            for (DocumentListener l : listeners) {
                doc.addDocumentListener(l);
            }
        }

    }

    private void showMiniNotepadPopup() {
        SwingUtilities.invokeLater(() -> {
            try {
                JDialog dialog = new JDialog(this, "Mini Notepad", false);
                ImageIcon notesIcon = loadIcon("/icons/notes.png", 32, 32);
                if (notesIcon != null) {
                    dialog.setIconImage(notesIcon.getImage());
                }

                JTextPane area = new JTextPane();
                JScrollPane scroll = new JScrollPane(area);
                File[] currentFile = { null };
                int[] fontSize = { 14 };

                // Toolbar
                JToolBar toolbar = new JToolBar();
                JButton newBtn = new JButton("New");
                newBtn.setIcon(loadIcon("/icons/new.png", 16, 16));
                JButton openBtn = new JButton("Open");
                openBtn.setIcon(loadIcon("/icons/openfile.png", 16, 16));
                JButton saveBtn = new JButton("Save");
                JButton saveAsBtn = new JButton("Save As");
                JButton zoomInBtn = new JButton("A+");
                JButton zoomOutBtn = new JButton("A-");
                JButton boldBtn = new JButton("B");
                JButton italicBtn = new JButton("I");
                JButton underlineBtn = new JButton("U");
                JButton fontBtn = new JButton("Font");
                JButton findBtn = new JButton("Find");
                InputMap im = area.getInputMap(JComponent.WHEN_FOCUSED);
                ActionMap am = area.getActionMap();

                im.put(KeyStroke.getKeyStroke(KeyEvent.VK_N, InputEvent.CTRL_DOWN_MASK), "newFile");
                im.put(KeyStroke.getKeyStroke(KeyEvent.VK_O, InputEvent.CTRL_DOWN_MASK), "openFile");
                im.put(KeyStroke.getKeyStroke(KeyEvent.VK_S, InputEvent.CTRL_DOWN_MASK), "saveFile");
                im.put(KeyStroke.getKeyStroke(KeyEvent.VK_S, InputEvent.CTRL_DOWN_MASK | InputEvent.SHIFT_DOWN_MASK),
                        "saveAsFile");
                im.put(KeyStroke.getKeyStroke(KeyEvent.VK_PLUS, InputEvent.CTRL_DOWN_MASK), "zoomIn");
        im.put(KeyStroke.getKeyStroke(KeyEvent.VK_EQUALS, InputEvent.CTRL_DOWN_MASK), "zoomIn"); // handle Ctrl+=
        im.put(KeyStroke.getKeyStroke(KeyEvent.VK_ADD, InputEvent.CTRL_DOWN_MASK), "zoomIn");   // numpad +
        im.put(KeyStroke.getKeyStroke(KeyEvent.VK_MINUS, InputEvent.CTRL_DOWN_MASK), "zoomOut");
        im.put(KeyStroke.getKeyStroke(KeyEvent.VK_SUBTRACT, InputEvent.CTRL_DOWN_MASK), "zoomOut"); // numpad -
                im.put(KeyStroke.getKeyStroke(KeyEvent.VK_B, InputEvent.CTRL_DOWN_MASK), "bold");
                im.put(KeyStroke.getKeyStroke(KeyEvent.VK_I, InputEvent.CTRL_DOWN_MASK), "italic");
                im.put(KeyStroke.getKeyStroke(KeyEvent.VK_U, InputEvent.CTRL_DOWN_MASK), "underline");
                im.put(KeyStroke.getKeyStroke(KeyEvent.VK_F, InputEvent.CTRL_DOWN_MASK), "find");

                am.put("newFile", new AbstractAction() {
                    public void actionPerformed(ActionEvent e) {
                        newBtn.doClick();
                    }
                });
                am.put("openFile", new AbstractAction() {
                    public void actionPerformed(ActionEvent e) {
                        openBtn.doClick();
                    }
                });
                am.put("saveFile", new AbstractAction() {
                    public void actionPerformed(ActionEvent e) {
                        saveBtn.doClick();
                    }
                });
                am.put("saveAsFile", new AbstractAction() {
                    public void actionPerformed(ActionEvent e) {
                        saveAsBtn.doClick();
                    }
                });
                am.put("zoomIn", new AbstractAction() {
                    public void actionPerformed(ActionEvent e) {
                        zoomInBtn.doClick();
                    }
                });
                am.put("zoomOut", new AbstractAction() {
                    public void actionPerformed(ActionEvent e) {
                        zoomOutBtn.doClick();
                    }
                });
                am.put("bold", new AbstractAction() {
                    public void actionPerformed(ActionEvent e) {
                        boldBtn.doClick();
                    }
                });
                am.put("italic", new AbstractAction() {
                    public void actionPerformed(ActionEvent e) {
                        italicBtn.doClick();
                    }
                });
                am.put("underline", new AbstractAction() {
                    public void actionPerformed(ActionEvent e) {
                        underlineBtn.doClick();
                    }
                });
                am.put("find", new AbstractAction() {
                    public void actionPerformed(ActionEvent e) {
                        findBtn.doClick();
                    }
                });

                toolbar.add(newBtn);
                toolbar.add(openBtn);
                toolbar.add(saveBtn);
                toolbar.add(saveAsBtn);
                toolbar.addSeparator();
                toolbar.add(zoomInBtn);
                toolbar.add(zoomOutBtn);
                toolbar.addSeparator();
                toolbar.add(boldBtn);
                toolbar.add(italicBtn);
                toolbar.add(underlineBtn);
                toolbar.addSeparator();
                toolbar.add(fontBtn);
                toolbar.addSeparator();
                toolbar.add(findBtn);

                // Actions
                newBtn.addActionListener(e -> area.setText(""));
                openBtn.addActionListener(e -> {
                    JFileChooser fc = new JFileChooser();
                    if (fc.showOpenDialog(dialog) == JFileChooser.APPROVE_OPTION) {
                        currentFile[0] = fc.getSelectedFile();
                        try (BufferedReader r = new BufferedReader(new FileReader(currentFile[0]))) {
                            StringBuilder sb = new StringBuilder();
                            String line;
                            while ((line = r.readLine()) != null)
                                sb.append(line).append("\n");
                            area.setText(sb.toString());
                        } catch (Exception ex) {
                            JOptionPane.showMessageDialog(dialog, "Open failed: " + ex.getMessage());
                        }
                    }
                });
                saveBtn.addActionListener(e -> {
                    if (currentFile[0] == null) {
                        saveAsBtn.doClick();
                        return;
                    }
                    try (BufferedWriter w = new BufferedWriter(new FileWriter(currentFile[0]))) {
                        w.write(area.getText());
                    } catch (Exception ex) {
                        JOptionPane.showMessageDialog(dialog, "Save failed: " + ex.getMessage());
                    }
                });
                saveAsBtn.addActionListener(e -> {
                    JFileChooser fc = new JFileChooser();
                    if (fc.showSaveDialog(dialog) == JFileChooser.APPROVE_OPTION) {
                        currentFile[0] = fc.getSelectedFile();
                        saveBtn.doClick();
                    }
                });
                zoomInBtn.addActionListener(e -> {
                    fontSize[0] += 2;
                    area.setFont(area.getFont().deriveFont((float) fontSize[0]));
                });
                zoomOutBtn.addActionListener(e -> {
                    fontSize[0] = Math.max(6, fontSize[0] - 2);
                    area.setFont(area.getFont().deriveFont((float) fontSize[0]));
                });
                boldBtn.addActionListener(e -> {
                    StyledDocument doc = area.getStyledDocument();
                    int start = area.getSelectionStart();
                    int end = area.getSelectionEnd();
                    if (start == end)
                        return;
                    MutableAttributeSet attr = new SimpleAttributeSet(doc.getCharacterElement(start).getAttributes());
                    StyleConstants.setBold(attr, !StyleConstants.isBold(attr));
                    doc.setCharacterAttributes(start, end - start, attr, false);
                });
                italicBtn.addActionListener(e -> {
                    StyledDocument doc = area.getStyledDocument();
                    int start = area.getSelectionStart();
                    int end = area.getSelectionEnd();
                    if (start == end)
                        return;
                    MutableAttributeSet attr = new SimpleAttributeSet(doc.getCharacterElement(start).getAttributes());
                    StyleConstants.setItalic(attr, !StyleConstants.isItalic(attr));
                    doc.setCharacterAttributes(start, end - start, attr, false);
                });
                underlineBtn.addActionListener(e -> {
                    StyledDocument doc = area.getStyledDocument();
                    int start = area.getSelectionStart();
                    int end = area.getSelectionEnd();
                    if (start == end)
                        return;
                    MutableAttributeSet attr = new SimpleAttributeSet(doc.getCharacterElement(start).getAttributes());
                    StyleConstants.setUnderline(attr, !StyleConstants.isUnderline(attr));
                    doc.setCharacterAttributes(start, end - start, attr, false);
                });
                fontBtn.addActionListener(e -> {
                    String[] fonts = java.awt.GraphicsEnvironment.getLocalGraphicsEnvironment()
                            .getAvailableFontFamilyNames();
                    String fontName = (String) JOptionPane.showInputDialog(dialog, "Choose font:", "Font",
                            JOptionPane.PLAIN_MESSAGE, null, fonts, area.getFont().getFamily());
                    if (fontName != null) {
                        area.setFont(new Font(fontName, area.getFont().getStyle(), area.getFont().getSize()));
                    }
                });
                findBtn.addActionListener(e -> {
                    String word = JOptionPane.showInputDialog(dialog, "Find text:");
                    if (word == null || word.isEmpty())
                        return;
                    Highlighter highlighter = area.getHighlighter();
                    highlighter.removeAllHighlights();
                    String text = area.getText().toLowerCase();
                    word = word.toLowerCase();
                    int index = text.indexOf(word);
                    while (index >= 0) {
                        try {
                            highlighter.addHighlight(index, index + word.length(),
                                    new DefaultHighlighter.DefaultHighlightPainter(Color.YELLOW));
                        } catch (Exception ignored) {
                        }
                        index = text.indexOf(word, index + word.length());
                    }
                });

                // Layout
                dialog.setLayout(new BorderLayout());
                dialog.add(toolbar, BorderLayout.NORTH);
                dialog.add(scroll, BorderLayout.CENTER);
                dialog.setSize(600, 400);
                dialog.setLocationRelativeTo(this);

                // Simply show the dialog without fade effect
                dialog.setVisible(true);

            } catch (Exception ex) {
                ex.printStackTrace();
                JOptionPane.showMessageDialog(this,
                        "Error opening mini notepad: " + ex.getMessage(),
                        "Error",
                        JOptionPane.ERROR_MESSAGE);
            }
        });
    }
    // Add this utility method anywhere in your TextEditor class:

    private void fadeInDialog(JDialog dialog, int durationMs) {
        dialog.setOpacity(0f);
        fadeInWindow(dialog, 300);
        Timer timer = new Timer(15, null);
        timer.addActionListener(new ActionListener() {
            float opacity = 0f;

            public void actionPerformed(ActionEvent e) {
                opacity += 0.05f;
                if (opacity >= 1f) {
                    opacity = 1f;
                    timer.stop();
                }
                dialog.setOpacity(opacity);
            }
        });
        timer.start();
    }

    private void fadeOutDialog(JDialog dialog, int durationMs) {
        Timer timer = new Timer(15, null);
        timer.addActionListener(new ActionListener() {
            float opacity = 1f;

            public void actionPerformed(ActionEvent e) {
                opacity -= 0.05f;
                if (opacity <= 0f) {
                    opacity = 0f;
                    timer.stop();
                    dialog.dispose();
                }
                dialog.setOpacity(opacity);
            }
        });
        timer.start();
    }

    class FadingPanel extends JPanel {
        private float alpha = 1.0f;
        private Timer fadeTimer;
        private Runnable onFadeComplete;

        public FadingPanel(LayoutManager layout) {
            super(layout);
            setOpaque(false);
        }

        public void setAlpha(float value) {
            this.alpha = Math.max(0.0f, Math.min(1.0f, value));
            repaint();
        }

        public void fadeIn(int durationMs, Runnable onComplete) {
            fade(0f, 1f, durationMs, onComplete);
        }

        public void fadeOut(int durationMs, Runnable onComplete) {
            fade(1f, 0f, durationMs, onComplete);
        }

        private void fade(float from, float to, int durationMs, Runnable onComplete) {
            if (fadeTimer != null && fadeTimer.isRunning()) {
                fadeTimer.stop();
            }

            this.onFadeComplete = onComplete;
            final long startTime = System.currentTimeMillis();

            fadeTimer = new Timer(16, e -> {
                long currentTime = System.currentTimeMillis();
                float progress = (float) (currentTime - startTime) / durationMs;
                progress = Math.min(1.0f, progress);

                float currentAlpha = from + (to - from) * progress;
                setAlpha(currentAlpha);

                if (progress >= 1.0f) {
                    fadeTimer.stop();
                    if (onFadeComplete != null) {
                        onFadeComplete.run();
                    }
                }
            });
            fadeTimer.start();
        }

        @Override
        protected void paintComponent(Graphics g) {
            Graphics2D g2d = (Graphics2D) g.create();
            g2d.setComposite(AlphaComposite.getInstance(AlphaComposite.SRC_OVER, alpha));
            super.paintComponent(g2d);
            g2d.dispose();
        }
    }

    private ImageIcon loadIcon(String path, int width, int height) {
        ImageIcon icon = new ImageIcon(getClass().getResource(path));
        Image scaledImage = icon.getImage().getScaledInstance(width, height, Image.SCALE_SMOOTH);
        return new ImageIcon(scaledImage);
    }

    private void addAnimatedTabSwitching() {
        tabbedPane.addChangeListener(e -> {
            Component currentComp = tabbedPane.getSelectedComponent();
            if (currentComp != null) {
                animateTabTransition(currentComp);
            }
        });
    }

   // ADD REUSABLE OBJECTS at class level:
private final Point reusablePoint = new Point();
private final Dimension reusableDimension = new Dimension();

private void animateTabTransition(Component tabComponent) {
    if (tabComponent instanceof JScrollPane) {
        JScrollPane scrollPane = (JScrollPane) tabComponent;
        Component view = scrollPane.getViewport().getView();

        if (view instanceof JComponent) {
            JComponent jcomp = (JComponent) view;

            // REUSE Point object
            Point originalLocation = jcomp.getLocation();
            reusablePoint.setLocation(tabbedPane.getWidth(), originalLocation.y);
            jcomp.setLocation(reusablePoint);

            Timer slideTimer = new Timer(16, null);
            slideTimer.addActionListener(new ActionListener() {
                int step = 0;
                int totalSteps = 20;

                public void actionPerformed(ActionEvent e) {
                    float progress = (float) step / totalSteps;
                    float easedProgress = (float) (1 - Math.pow(1 - progress, 2));

                    // REUSE Point object instead of creating new ones
                    int newX = (int) (tabbedPane.getWidth() * (1 - easedProgress));
                    reusablePoint.setLocation(newX, originalLocation.y);
                    jcomp.setLocation(reusablePoint);

                    step++;
                    if (step > totalSteps) {
                        reusablePoint.setLocation(originalLocation);
                        jcomp.setLocation(reusablePoint);
                        slideTimer.stop();
                    }
                }
            });
            slideTimer.start();
        }
    }
}
    private void addSmoothScrolling(JScrollPane scrollPane) {
        JScrollBar verticalBar = scrollPane.getVerticalScrollBar();
        JScrollBar horizontalBar = scrollPane.getHorizontalScrollBar();

        setupSmoothScroller(verticalBar);
        setupSmoothScroller(horizontalBar);
    }

    private void setupSmoothScroller(JScrollBar scrollBar) {
        scrollBar.setUnitIncrement(16);
        scrollBar.setBlockIncrement(50);

        // Add mouse wheel acceleration
        scrollBar.addMouseWheelListener(e -> {
            if (!e.isShiftDown()) {
                smoothScrollTo(scrollBar,
                        scrollBar.getValue() + e.getWheelRotation() * scrollBar.getUnitIncrement() * 3);
            }
        });
    }

    private void smoothScrollTo(JScrollBar scrollBar, int targetValue) {
        final int startValue = scrollBar.getValue();
        final int distance = targetValue - startValue;

        Timer scrollTimer = new Timer(16, null);
        scrollTimer.addActionListener(new ActionListener() {
            int step = 0;
            int totalSteps = 15;

            public void actionPerformed(ActionEvent e) {
                float progress = (float) step / totalSteps;
                // Cubic ease-out function for smooth deceleration
                float easedProgress = 1 - (float) Math.pow(1 - progress, 3);

                int currentValue = startValue + (int) (distance * easedProgress);
                scrollBar.setValue(currentValue);

                step++;
                if (step >= totalSteps) {
                    scrollBar.setValue(targetValue);
                    scrollTimer.stop();
                }
            }
        });
        scrollTimer.start();
    }

    private void setupAnimatedButton(JButton button) {
        Color originalBg = button.getBackground();
        Color originalFg = button.getForeground();
        Font originalFont = button.getFont();

        button.addMouseListener(new MouseAdapter() {
            private Timer hoverTimer;

            @Override
            public void mouseEntered(MouseEvent e) {
                animateButtonHover(button, originalBg.brighter(), originalFg,
                        originalFont.deriveFont(originalFont.getSize() + 0.5f), 150);
            }

            @Override
            public void mouseExited(MouseEvent e) {
                animateButtonHover(button, originalBg, originalFg, originalFont, 150);
            }

            @Override
            public void mousePressed(MouseEvent e) {
                animateButtonScale(button, 0.95f, 100);
            }

            @Override
            public void mouseReleased(MouseEvent e) {
                animateButtonScale(button, 1.0f, 100);
            }
        });
    }

    private void animateButtonHover(JButton button, Color targetBg, Color targetFg, Font targetFont, int duration) {
        Color startBg = button.getBackground();
        Color startFg = button.getForeground();
        Font startFont = button.getFont();

        Timer timer = new Timer(16, null);
        final long startTime = System.currentTimeMillis();

        timer.addActionListener(e -> {
            long currentTime = System.currentTimeMillis();
            float progress = (float) (currentTime - startTime) / duration;
            progress = Math.min(1.0f, progress);

            // Ease out
            float easedProgress = 1 - (float) Math.pow(1 - progress, 2);

            // Interpolate colors
            Color currentBg = interpolateColor(startBg, targetBg, easedProgress);
            Color currentFg = interpolateColor(startFg, targetFg, easedProgress);

            // Interpolate font size
            float currentSize = startFont.getSize() + (targetFont.getSize() - startFont.getSize()) * easedProgress;
            Font currentFont = startFont.deriveFont(currentSize);

            button.setBackground(currentBg);
            button.setForeground(currentFg);
            button.setFont(currentFont);

            if (progress >= 1.0f) {
                timer.stop();
            }
        });
        timer.start();
    }

    private void animateButtonScale(JButton button, float scale, int duration) {
        Dimension originalSize = button.getSize();
        Dimension targetSize = new Dimension((int) (originalSize.width * scale),
                (int) (originalSize.height * scale));

        Timer timer = new Timer(16, null);
        final long startTime = System.currentTimeMillis();

        timer.addActionListener(e -> {
            long currentTime = System.currentTimeMillis();
            float progress = (float) (currentTime - startTime) / duration;
            progress = Math.min(1.0f, progress);

            float easedProgress = (float) Math.sin(progress * Math.PI / 2); // Sine ease

            int currentWidth = (int) (originalSize.width + (targetSize.width - originalSize.width) * easedProgress);
            int currentHeight = (int) (originalSize.height + (targetSize.height - originalSize.height) * easedProgress);

            button.setPreferredSize(new Dimension(currentWidth, currentHeight));
            button.revalidate();

            if (progress >= 1.0f) {
                timer.stop();
                button.setPreferredSize(targetSize);
                button.revalidate();
            }
        });
        timer.start();
    }

    private void addTypingAnimation(JTextPane textPane) {
        textPane.getDocument().addDocumentListener(new DocumentListener() {
            private Timer pulseTimer;

            public void insertUpdate(DocumentEvent e) {
                showTypingEffect();
            }

            public void removeUpdate(DocumentEvent e) {
                showTypingEffect();
            }

            public void changedUpdate(DocumentEvent e) {
            }

            private void showTypingEffect() {
                if (pulseTimer != null && pulseTimer.isRunning()) {
                    pulseTimer.stop();
                }

                final Color originalBg = textPane.getBackground();
                final Color pulseColor = new Color(originalBg.getRed(),
                        originalBg.getGreen(),
                        originalBg.getBlue(),
                        150);

                textPane.setBackground(pulseColor);

                pulseTimer = new Timer(100, e -> {
                    textPane.setBackground(originalBg);
                    ((Timer) e.getSource()).stop();
                });
                pulseTimer.setRepeats(false);
                pulseTimer.start();
            }
        });
    }

    private Color interpolateColor(Color start, Color end, float progress) {
        int r = (int) (start.getRed() + (end.getRed() - start.getRed()) * progress);
        int g = (int) (start.getGreen() + (end.getGreen() - start.getGreen()) * progress);
        int b = (int) (start.getBlue() + (end.getBlue() - start.getBlue()) * progress);
        int a = (int) (start.getAlpha() + (end.getAlpha() - start.getAlpha()) * progress);

        return new Color(r, g, b, a);
    }

    private void showAnimatedNotification(String message, Color bgColor) {
        JDialog notification = new JDialog(this, false);
        notification.setUndecorated(true);
        notification.setAlwaysOnTop(true);

        JLabel label = new JLabel(message, SwingConstants.CENTER);
        label.setForeground(Color.WHITE);
        label.setBackground(bgColor);
        label.setOpaque(true);
        label.setBorder(BorderFactory.createEmptyBorder(10, 20, 10, 20));

        notification.add(label);
        notification.pack();

        // Position at top right
        Point loc = getLocationOnScreen();
        notification.setLocation(loc.x + getWidth() - notification.getWidth() - 20, loc.y + 40);

        // Slide in from top
        notification.setOpacity(0f);
        Point startPos = new Point(notification.getLocation().x, -notification.getHeight());
        notification.setLocation(startPos);
        notification.setVisible(true);

        // Slide down animation
        Timer slideTimer = new Timer(16, null);
        final Point targetPos = notification.getLocation();
        final long startTime = System.currentTimeMillis();
        final int duration = 500;

        slideTimer.addActionListener(e -> {
            long currentTime = System.currentTimeMillis();
            float progress = (float) (currentTime - startTime) / duration;
            progress = Math.min(1.0f, progress);

            // Bounce ease out
            float easedProgress = 1 - (float) Math.pow(1 - progress, 2);

            int currentY = startPos.y + (int) ((targetPos.y - startPos.y) * easedProgress);
            float currentOpacity = progress;

            notification.setLocation(startPos.x, currentY);
            notification.setOpacity(currentOpacity);

            if (progress >= 1.0f) {
                slideTimer.stop();
                // Auto-close after 3 seconds
                new Timer(3000, ev -> {
                    fadeOutWindow(notification, 300);
                }).start();
            }
        });
        slideTimer.start();
    }

    private void showStopwatchPopup() {
        // --- DIALOG SETUP ---
        JDialog dialog = new JDialog(this, false);
        dialog.setUndecorated(true); // Essential for custom UI and animations
        dialog.setSize(420, 480); // Set a fixed initial size
        dialog.setLocationRelativeTo(this);
        dialog.setLayout(new BorderLayout());

        // --- MODERN UI COMPONENTS ---
        // Digital Display
        JLabel timeLabel = new JLabel("00:00.00", SwingConstants.CENTER);
        timeLabel.setFont(new Font("Segoe UI Semibold", Font.PLAIN, 56));
        timeLabel.setForeground(UIManager.getColor("Component.focusColor")); // Use theme's accent color
        timeLabel.setBorder(BorderFactory.createEmptyBorder(20, 10, 10, 10));

        // Analog Clock
        ClockPanel clockPanel = new ClockPanel();
        clockPanel.setPreferredSize(new Dimension(150, 150));

        // Lap Display using JList
        DefaultListModel<String> lapModel = new DefaultListModel<>();
        JList<String> lapList = new JList<>(lapModel);
        lapList.setFont(new Font("Consolas", Font.PLAIN, 14));
        JScrollPane lapScrollPane = new JScrollPane(lapList);
        lapScrollPane.setBorder(BorderFactory.createTitledBorder("Laps"));
        lapScrollPane.setPreferredSize(new Dimension(150, 0));

        // Buttons with FlatLaf style
        JButton startStopBtn = new JButton(IconUtils.createIcon("/icons/run.png"));
        JButton lapBtn = new JButton(IconUtils.createIcon("/icons/lap.png"));
        JButton resetBtn = new JButton(IconUtils.createIcon("/icons/replace.png"));
        lapBtn.setEnabled(false);

        // --- CUSTOM TITLE BAR with Window Controls ---
        JPanel titleBar = createCustomTitleBar(dialog);

        // --- LAYOUT ---
        JPanel controls = new JPanel(new FlowLayout(FlowLayout.CENTER, 20, 10));
        for (JButton btn : new JButton[] { resetBtn, startStopBtn, lapBtn }) {
            btn.putClientProperty("JButton.buttonType", "toolBarButton");
            controls.add(btn);
        }

        JPanel centerPanel = new JPanel(new GridBagLayout());
        GridBagConstraints gbc = new GridBagConstraints();
        gbc.gridx = 0;
        gbc.gridy = 0;
        gbc.weightx = 1.0;
        gbc.weighty = 1.0;
        gbc.fill = GridBagConstraints.BOTH;
        gbc.insets = new Insets(10, 10, 10, 5);
        centerPanel.add(clockPanel, gbc);

        gbc.gridx = 1;
        gbc.insets = new Insets(10, 5, 10, 10);
        centerPanel.add(lapScrollPane, gbc);

        dialog.add(titleBar, BorderLayout.NORTH);
        dialog.add(timeLabel, BorderLayout.CENTER);

        JPanel bottomContainer = new JPanel(new BorderLayout());
        bottomContainer.add(centerPanel, BorderLayout.CENTER);
        bottomContainer.add(controls, BorderLayout.SOUTH);
        dialog.add(bottomContainer, BorderLayout.SOUTH);

        // --- TIMER LOGIC (Faster timer for smoother animation) ---
        final Timer timer = new Timer(16, null); // ~60 FPS for smooth sweep
        // ... (rest of the timer, start/stop, lap, and reset logic is the same) ...

        // --- SHOW DIALOG ---
        dialog.setVisible(true);
    }

    /**
     * Creates a custom, draggable title bar with functional minimize, maximize, and
     * close buttons.
     */
    private JPanel createCustomTitleBar(JDialog dialog) {
        JPanel titleBar = new JPanel(new BorderLayout());
        titleBar.setBackground(UIManager.getColor("MenuBar.background"));
        titleBar.setBorder(BorderFactory.createEmptyBorder(4, 8, 4, 4));

        JLabel titleLabel = new JLabel("Stopwatch");

        // Window control buttons
        JButton minimizeBtn = new JButton("—");
        JButton maximizeBtn = new JButton("☐");
        JButton closeBtn = new JButton("✕");

        JPanel buttonPanel = new JPanel();
        for (JButton btn : new JButton[] { minimizeBtn, maximizeBtn, closeBtn }) {
            btn.putClientProperty("JButton.buttonType", "toolBarButton");
            btn.setFocusable(false);
            buttonPanel.add(btn);
        }

        // Button actions
        closeBtn.addActionListener(e -> dialog.dispose());
        minimizeBtn.addActionListener(e -> this.setState(Frame.ICONIFIED));

        final boolean[] isMaximized = { false };
        final Rectangle[] normalBounds = { null };
        maximizeBtn.addActionListener(e -> {
            if (!isMaximized[0]) {
                normalBounds[0] = dialog.getBounds();
                dialog.setBounds(GraphicsEnvironment.getLocalGraphicsEnvironment().getMaximumWindowBounds());
            } else {
                dialog.setBounds(normalBounds[0]);
            }
            isMaximized[0] = !isMaximized[0];
        });

        titleBar.add(titleLabel, BorderLayout.CENTER);
        titleBar.add(buttonPanel, BorderLayout.EAST);

        // Drag listener
        final Point[] startPoint = { null };
        titleBar.addMouseListener(new MouseAdapter() {
            public void mousePressed(MouseEvent e) {
                startPoint[0] = e.getPoint();
            }
        });
        titleBar.addMouseMotionListener(new MouseAdapter() {
            public void mouseDragged(MouseEvent e) {
                Point p = dialog.getLocation();
                dialog.setLocation(p.x + e.getX() - startPoint[0].x, p.y + e.getY() - startPoint[0].y);
            }
        });

        return titleBar;
    }

    // Analog clock panel for stopwatch
    /**
     * An inner class for the analog display of the stopwatch.
     */
    class ClockPanel extends JPanel {
        private long millis = 0;

        public ClockPanel() {
            setPreferredSize(new Dimension(200, 200));
            setOpaque(false);
        }

        public void setMillis(long ms) {
            this.millis = ms;
            repaint();
        }

        @Override
        protected void paintComponent(Graphics g) {
            super.paintComponent(g);
            int w = getWidth(), h = getHeight();
            int r = Math.min(w, h) / 2 - 10;
            int cx = w / 2, cy = h / 2;
            Graphics2D g2 = (Graphics2D) g;
            g2.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);

            g2.setColor(new Color(245, 245, 245));
            g2.fillOval(cx - r, cy - r, 2 * r, 2 * r);
            g2.setColor(Color.DARK_GRAY);
            g2.setStroke(new BasicStroke(2));
            g2.drawOval(cx - r, cy - r, 2 * r, 2 * r);

            // Draw minute hand
            double minutes = (millis / 60000.0);
            double minAngle = Math.toRadians(minutes * 6 - 90);
            int mx = (int) (cx + Math.cos(minAngle) * (r * 0.6));
            int my = (int) (cy + Math.sin(minAngle) * (r * 0.6));
            g2.setColor(Color.BLACK);
            g2.setStroke(new BasicStroke(4));
            g2.drawLine(cx, cy, mx, my);

            // Draw second hand
            double seconds = (millis / 1000.0) % 60; // This now includes fractions of a second
            double secAngle = Math.toRadians(seconds * 6 - 90);
            int sx = (int) (cx + Math.cos(secAngle) * (r * 0.8));
            int sy = (int) (cy + Math.sin(secAngle) * (r * 0.8));
            g2.setColor(Color.RED);
            g2.setStroke(new BasicStroke(2));
            g2.drawLine(cx, cy, sx, sy);

            // Draw center
            g2.setColor(Color.BLACK);
            g2.fillOval(cx - 5, cy - 5, 10, 10);
        }
    }

    // In PptxEditorPanel.java (inside package com.example.notepadpp.media)
    public static void addEditablePPTXTab(File file, JTabbedPane tabbedPane, Map<Component, String> tabTypeMap) {
        try (FileInputStream fis = new FileInputStream(file)) {
            XMLSlideShow ppt = new XMLSlideShow(fis);
            java.util.List<org.apache.poi.xslf.usermodel.XSLFSlide> slides = ppt.getSlides();
            JPanel panel = new JPanel();
            panel.setLayout(new BoxLayout(panel, BoxLayout.Y_AXIS));
            java.util.List<JTextArea> textAreas = new ArrayList<>();
            for (org.apache.poi.xslf.usermodel.XSLFSlide slide : slides) {
                StringBuilder sb = new StringBuilder();
                for (org.apache.poi.xslf.usermodel.XSLFShape shape : slide.getShapes()) {
                    if (shape instanceof org.apache.poi.xslf.usermodel.XSLFTextShape) {
                        sb.append(((org.apache.poi.xslf.usermodel.XSLFTextShape) shape).getText()).append("\n");
                    }
                }
                JTextArea area = new JTextArea(sb.toString(), 8, 80);
                textAreas.add(area);
                panel.add(new JLabel("Slide " + textAreas.size()));
                panel.add(new JScrollPane(area));
            }
            JScrollPane scroll = new JScrollPane(panel);
            tabbedPane.addTab(file.getName(), scroll);
            tabbedPane.setSelectedComponent(scroll);
            // Use a consistent key with no extra spaces.
            tabTypeMap.put(scroll, "pptx_editable");
        } catch (Exception ex) {
            JOptionPane.showMessageDialog(null, "Cannot open PPTX for editing :\n" + ex.getMessage());
        }
    }

    private void addDrawingTab() {
        ImageIcon originalIcon = new ImageIcon(getClass().getResource("/icons/drawing.png"));
        int fixedWidth = 16;
        int fixedHeight = 16;
        Image scaledImage = originalIcon.getImage().getScaledInstance(fixedWidth, fixedHeight, Image.SCALE_SMOOTH);
        Icon logoIcon = new ImageIcon(scaledImage);

        DrawingPanel panel = new DrawingPanel();
        JPanel container = new JPanel(new BorderLayout());
        container.add(panel.getToolbar(), BorderLayout.NORTH);
        JScrollPane scroll = new JScrollPane(panel);
        container.add(scroll, BorderLayout.CENTER);
        tabbedPane.addTab("Drawing ", logoIcon, container);
        tabbedPane.setSelectedComponent(container);
        tabTypeMap.put(container, "drawing ");
    }

    class DrawingPanel extends JPanel {
        private java.util.List<ShapeInfo> shapes = new ArrayList<>();
        private java.util.List<ShapeInfo> undoStack = new ArrayList<>();
        private java.util.List<Point> freehandPoints = new ArrayList<>();
        private java.util.List<Point> polygonPoints = new ArrayList<>();
        private Point start, end;
        private String currentTool = "Rectangle";
        private Color strokeColor = Color.BLACK;
        private Color fillColor = Color.WHITE;
        private boolean fillMode = false;
        private int strokeWidth = 2;
        private boolean isDrawingPolygon = false;

        public DrawingPanel() {
            setPreferredSize(new Dimension(1200, 900));
            setBackground(Color.WHITE);

            addMouseListener(new MouseAdapter() {
                public void mousePressed(MouseEvent e) {
                    start = e.getPoint();
                    if ("Freehand".equals(currentTool)) {
                        freehandPoints = new ArrayList<>();
                        freehandPoints.add(start);
                    } else if ("Polygon".equals(currentTool)) {
                        if (!isDrawingPolygon) {
                            polygonPoints.clear();
                            isDrawingPolygon = true;
                        }
                        polygonPoints.add(start);
                        repaint();
                    } else if ("Color Picker".equals(currentTool)) {
                        BufferedImage img = new BufferedImage(getWidth(), getHeight(), BufferedImage.TYPE_INT_RGB);
                        Graphics2D g2 = img.createGraphics();
                        paint(g2);
                        g2.dispose();
                        strokeColor = new Color(img.getRGB(e.getX(), e.getY()));
                    } else if ("Text".equals(currentTool)) {
                        String input = JOptionPane.showInputDialog(DrawingPanel.this, "Enter text:");
                        if (input != null && !input.trim().isEmpty()) {
                            shapes.add(new ShapeInfo(new TextShape(input, start.x, start.y, strokeColor, strokeWidth),
                                    strokeColor, null, false, "Text", null, strokeWidth));
                            undoStack.clear();
                            repaint();
                        }
                    }
                }
                // Place inside DrawingPanel class

                /**
                 * Smoothly morphs from one shape to another over the given duration.
                 * Only works well for shapes with the same number of points (e.g., polygon to
                 * polygon).
                 */
                public void animateShapeTransition(Shape from, Shape to, int durationMs) {
                    // For demo: only works for Polygon shapes with same number of points
                    if (!(from instanceof Polygon) || !(to instanceof Polygon))
                        return;
                    Polygon p1 = (Polygon) from;
                    Polygon p2 = (Polygon) to;
                    if (p1.npoints != p2.npoints)
                        return;

                    int steps = 30;
                    int delay = durationMs / steps;
                    Timer timer = new Timer(delay, null);
                    final int[] step = { 0 };
                    ShapeInfo morphingShape = new ShapeInfo(
                            new Polygon(p1.xpoints.clone(), p1.ypoints.clone(), p1.npoints),
                            Color.MAGENTA, null, false, "Morph", null, 3);

                    shapes.add(morphingShape);
                    repaint();

                    timer.addActionListener(e -> {
                        float t = (float) step[0] / steps;
                        for (int i = 0; i < p1.npoints; i++) {
                            int x = (int) (p1.xpoints[i] * (1 - t) + p2.xpoints[i] * t);
                            int y = (int) (p1.ypoints[i] * (1 - t) + p2.ypoints[i] * t);
                            ((Polygon) morphingShape.shape).xpoints[i] = x;
                            ((Polygon) morphingShape.shape).ypoints[i] = y;
                        }
                        ((Polygon) morphingShape.shape).invalidate();
                        repaint();
                        step[0]++;
                        if (step[0] > steps) {
                            shapes.remove(morphingShape);
                            shapes.add(new ShapeInfo(to, Color.BLUE, null, false, "Polygon", null, 3));
                            repaint();
                            timer.stop();
                        }
                    });
                    timer.start();
                }

                public void mouseReleased(MouseEvent e) {
                    end = e.getPoint();
                    if (start == null || end == null || "Polygon".equals(currentTool)
                            || "Color Picker".equals(currentTool) || "Text".equals(currentTool))
                        return;

                    Shape shape = null;
                    if ("Rectangle".equals(currentTool)) {
                        shape = new Rectangle(Math.min(start.x, end.x), Math.min(start.y, end.y),
                                Math.abs(end.x - start.x), Math.abs(end.y - start.y));
                    } else if ("Rounded Rectangle".equals(currentTool)) {
                        shape = new RoundRectangle2D.Double(Math.min(start.x, end.x), Math.min(start.y, end.y),
                                Math.abs(end.x - start.x), Math.abs(end.y - start.y), 30, 30);
                    } else if ("Oval".equals(currentTool)) {
                        shape = new Ellipse2D.Double(Math.min(start.x, end.x), Math.min(start.y, end.y),
                                Math.abs(end.x - start.x), Math.abs(end.y - start.y));
                    } else if ("Line".equals(currentTool)) {
                        shape = new Line2D.Double(start, end);
                    } else if ("Freehand".equals(currentTool)) {
                        freehandPoints.add(end);
                        shape = new Polygon(
                                freehandPoints.stream().mapToInt(p -> p.x).toArray(),
                                freehandPoints.stream().mapToInt(p -> p.y).toArray(),
                                freehandPoints.size());
                    } else if ("Eraser".equals(currentTool)) {
                        shape = new Line2D.Double(start, end);
                    } else if ("Highlighter".equals(currentTool)) {
                        shape = new Line2D.Double(start, end);
                    }
                    if (shape != null) {
                        Color color = "Eraser".equals(currentTool) ? Color.WHITE : strokeColor;
                        int alpha = "Highlighter".equals(currentTool) ? 50 : 255;
                        Color realColor = new Color(color.getRed(), color.getGreen(), color.getBlue(), alpha);
                        shapes.add(new ShapeInfo(shape, realColor, fillColor, fillMode, currentTool, freehandPoints,
                                strokeWidth));
                        undoStack.clear();
                        repaint();
                    }
                    start = end = null;
                    freehandPoints = new ArrayList<>();
                }
            });

            addMouseMotionListener(new MouseAdapter() {
                public void mouseDragged(MouseEvent e) {
                    if ("Freehand".equals(currentTool) && start != null) {
                        freehandPoints.add(e.getPoint());
                        repaint();
                    } else if ("Eraser".equals(currentTool) || "Highlighter".equals(currentTool)) {
                        end = e.getPoint();
                        Shape shape = new Line2D.Double(start, end);
                        Color color = "Eraser".equals(currentTool) ? Color.WHITE : strokeColor;
                        int alpha = "Highlighter".equals(currentTool) ? 50 : 255;
                        Color realColor = new Color(color.getRed(), color.getGreen(), color.getBlue(), alpha);
                        shapes.add(new ShapeInfo(shape, realColor, fillColor, false, currentTool, null, strokeWidth));
                        start = end;
                        repaint();
                    }
                }
            });

            setDoubleBuffered(true);
        }

        public JToolBar getToolbar() {
            JToolBar toolbar = new JToolBar();
            String[] tools = { "Rectangle", "Rounded Rectangle", "Oval", "Line", "Freehand", "Polygon", "Text",
                    "Eraser", "Color Picker", "Highlighter" };
            JComboBox<String> toolBox = new JComboBox<>(tools);
            toolBox.addActionListener(e -> currentTool = (String) toolBox.getSelectedItem());
            toolbar.add(new JLabel("Tool: "));
            toolbar.add(toolBox);

            JComboBox<Integer> widthBox = new JComboBox<>(new Integer[] { 1, 2, 4, 8, 12, 16, 24, 32 });
            widthBox.setSelectedItem(strokeWidth);
            widthBox.addActionListener(e -> strokeWidth = (Integer) widthBox.getSelectedItem());
            toolbar.add(new JLabel("Width: "));
            toolbar.add(widthBox);

            JButton strokeBtn = new JButton("Stroke Color");
            strokeBtn.addActionListener(e -> {
                Color c = JColorChooser.showDialog(this, "Pick Stroke Color", strokeColor);
                if (c != null)
                    strokeColor = c;
            });
            toolbar.add(strokeBtn);

            JButton fillBtn = new JButton("Fill Color");
            fillBtn.addActionListener(e -> {
                Color c = JColorChooser.showDialog(this, "Pick Fill Color", fillColor);
                if (c != null)
                    fillColor = c;
            });
            toolbar.add(fillBtn);

            JCheckBox fillCheck = new JCheckBox("Fill Shape");
            fillCheck.addActionListener(e -> fillMode = fillCheck.isSelected());
            toolbar.add(fillCheck);

            JButton undoBtn = new JButton("Undo");
            undoBtn.addActionListener(e -> {
                if (!shapes.isEmpty()) {
                    undoStack.add(shapes.remove(shapes.size() - 1));
                    repaint();
                }
            });
            toolbar.add(undoBtn);

            JButton redoBtn = new JButton("Redo");
            redoBtn.addActionListener(e -> {
                if (!undoStack.isEmpty()) {
                    shapes.add(undoStack.remove(undoStack.size() - 1));
                    repaint();
                }
            });
            toolbar.add(redoBtn);

            JButton clearBtn = new JButton("Clear");
            clearBtn.addActionListener(e -> {
                shapes.clear();
                undoStack.clear();
                repaint();
            });
            toolbar.add(clearBtn);

            JButton saveBtn = new JButton("Save as PNG");
            saveBtn.addActionListener(e -> saveAsImage());
            toolbar.add(saveBtn);

            return toolbar;
        }

        protected void paintComponent(Graphics g) {
            super.paintComponent(g);
            Graphics2D g2 = (Graphics2D) g;
            for (ShapeInfo info : shapes) {
                g2.setStroke(new BasicStroke(info.strokeWidth));
                g2.setColor(info.strokeColor);
                if (info.fill && (info.shape instanceof Rectangle || info.shape instanceof Ellipse2D
                        || info.shape instanceof RoundRectangle2D)) {
                    g2.setColor(info.fillColor);
                    if (info.shape instanceof Rectangle) {
                        Rectangle r = (Rectangle) info.shape;
                        g2.fillRect(r.x, r.y, r.width, r.height);
                    } else if (info.shape instanceof Ellipse2D) {
                        g2.fill((Ellipse2D) info.shape);
                    } else if (info.shape instanceof RoundRectangle2D) {
                        g2.fill((RoundRectangle2D) info.shape);
                    }
                    g2.setColor(info.strokeColor);
                }
                if (info.shape instanceof Rectangle)
                    g2.draw((Rectangle) info.shape);
                else if (info.shape instanceof Ellipse2D)
                    g2.draw((Ellipse2D) info.shape);
                else if (info.shape instanceof Line2D)
                    g2.draw((Line2D) info.shape);
                else if (info.shape instanceof Polygon)
                    g2.draw((Polygon) info.shape);
                else if (info.shape instanceof RoundRectangle2D)
                    g2.draw((RoundRectangle2D) info.shape);
                else if (info.shape instanceof TextShape) {
                    TextShape ts = (TextShape) info.shape;
                    g2.setFont(new Font("Arial", Font.PLAIN, ts.size * 5));
                    g2.drawString(ts.text, ts.x, ts.y);
                }
            }
            if (isDrawingPolygon && polygonPoints.size() > 1) {
                g2.setColor(Color.GRAY);
                for (int i = 1; i < polygonPoints.size(); i++) {
                    Point p1 = polygonPoints.get(i - 1);
                    Point p2 = polygonPoints.get(i);
                    g2.drawLine(p1.x, p1.y, p2.x, p2.y);
                }
            }
        }

        private void saveAsImage() {
            JFileChooser fc = new JFileChooser();
            fc.setDialogTitle("Save Drawing As PNG");
            fc.setSelectedFile(new File("drawing.png"));
            if (fc.showSaveDialog(this) != JFileChooser.APPROVE_OPTION)
                return;
            File file = fc.getSelectedFile();
            BufferedImage img = new BufferedImage(getWidth(), getHeight(), BufferedImage.TYPE_INT_RGB);
            Graphics2D g2 = img.createGraphics();
            paint(g2);
            g2.dispose();
            try {
                javax.imageio.ImageIO.write(img, "png", file);
                JOptionPane.showMessageDialog(this, "Saved: " + file.getAbsolutePath());
            } catch (Exception ex) {
                JOptionPane.showMessageDialog(this, "Failed to save image:\n" + ex.getMessage());
            }
        }

        static class ShapeInfo {
            Shape shape;
            Color strokeColor, fillColor;
            boolean fill;
            String type;
            List<Point> freehandPoints;
            int strokeWidth;

            ShapeInfo(Shape s, Color sc, Color fc, boolean f, String t, List<Point> points, int strokeWidth) {
                shape = s;
                strokeColor = sc;
                fillColor = fc;
                fill = f;
                type = t;
                freehandPoints = points == null ? null : new ArrayList<>(points);
                this.strokeWidth = strokeWidth;
            }
        }

        static class TextShape implements Shape {
            String text;
            int x, y, size;
            Color color;

            public TextShape(String text, int x, int y, Color color, int size) {
                this.text = text;
                this.x = x;
                this.y = y;
                this.color = color;
                this.size = size;
            }

            @Override
            public Rectangle getBounds() {
                return new Rectangle(x, y - size, text.length() * size, size * 5);
            }

            @Override
            public Rectangle2D getBounds2D() {
                return getBounds();
            }

            @Override
            public boolean contains(double x, double y) {
                return false;
            }

            @Override
            public boolean contains(Point2D p) {
                return false;
            }

            @Override
            public boolean intersects(double x, double y, double w, double h) {
                return false;
            }

            @Override
            public boolean intersects(Rectangle2D r) {
                return false;
            }

            @Override
            public boolean contains(double x, double y, double w, double h) {
                return false;
            }

            @Override
            public boolean contains(Rectangle2D r) {
                return false;
            }

            @Override
            public PathIterator getPathIterator(AffineTransform at) {
                return null;
            }

            @Override
            public PathIterator getPathIterator(AffineTransform at, double flatness) {
                return null;
            }
        }
    }

    // ...
    private void addPPTXTab() {
        // OPTIONAL: Set FlatLaf look and feel locally (if not already applied globally)
        try {
            UIManager.setLookAndFeel(new FlatLightLaf());
        } catch (Exception ex) {
            ex.printStackTrace();
        }

        // Build the PPTX editing tab
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

        // --- Base editing controls ---
        // Basic controls (already existing)
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

        // Extra controls
        JPanel extraBtnPanel = new JPanel();
        JButton duplicateSlideBtn = new JButton("Duplicate Slide");
        JButton clearSlideBtn = new JButton("Clear Slide");
        JButton moveUpBtn = new JButton("Move Slide Up");
        JButton moveDownBtn = new JButton("Move Slide Down");
        JButton setFontBtn = new JButton("Set Font");
        JButton setTextColorBtn = new JButton("Set Text Color");
        extraBtnPanel.add(duplicateSlideBtn);
        extraBtnPanel.add(clearSlideBtn);
        extraBtnPanel.add(moveUpBtn);
        extraBtnPanel.add(moveDownBtn);
        extraBtnPanel.add(setFontBtn);
        extraBtnPanel.add(setTextColorBtn);

        // Combine into a single toolbar panel
        JPanel toolbarPanel = new JPanel();
        toolbarPanel.setLayout(new BoxLayout(toolbarPanel, BoxLayout.Y_AXIS));
        toolbarPanel.add(btnPanel);
        toolbarPanel.add(extraBtnPanel);
        mainPanel.add(toolbarPanel, BorderLayout.NORTH);

        // Create the first slide.
        addSlide();

        // Base control actions.
        addSlideBtn.addActionListener(e -> addSlide());
        removeSlideBtn.addActionListener(e -> removeCurrentSlide());
        addChartBtn.addActionListener(e -> addBarChartToCurrentSlide());
        addClipartBtn.addActionListener(e -> addClipartToCurrentSlide());
        addBulletBtn.addActionListener(e -> addBulletsToCurrentSlide());
        changeBgBtn.addActionListener(e -> changeBackgroundColor());
        insertTextBoxBtn.addActionListener(e -> insertTextBox());

        // Extra control actions.
        duplicateSlideBtn.addActionListener(e -> duplicateSlide());
        clearSlideBtn.addActionListener(e -> clearSlide());
        moveUpBtn.addActionListener(e -> moveSlideUp());
        moveDownBtn.addActionListener(e -> moveSlideDown());
        setFontBtn.addActionListener(e -> setSlideFont());
        setTextColorBtn.addActionListener(e -> setSlideTextColor());
        // Update UI to match the currently active flat theme (dark or light)
        SwingUtilities.updateComponentTreeUI(mainPanel);

        ImageIcon originalIcon = new ImageIcon(getClass().getResource("/icons/powerpoint.png"));
        int fixedWidth = 16;
        int fixedHeight = 16;
        Image scaledImage = originalIcon.getImage().getScaledInstance(fixedWidth, fixedHeight, Image.SCALE_SMOOTH);
        Icon logoIcon = new ImageIcon(scaledImage);
        JScrollPane mainScroll = new JScrollPane(mainPanel);
        tabbedPane.addTab("New PPTX", logoIcon, mainScroll);
        tabbedPane.setSelectedComponent(mainScroll);
        tabTypeMap.put(mainScroll, "pptx_editable");
    }

    /*
     * Helper methods for moving slides – these simply rebuild the PPTX panel order
     */
    private void moveSlideUp() {
        // For simplicity, swap the last slide with the one before it.
        int size = pptxSlides.size();
        if (size > 1) {
            // Swap last two slides
            JTextArea last = pptxSlides.get(size - 1);
            JTextArea prev = pptxSlides.get(size - 2);
            pptxSlides.set(size - 2, last);
            pptxSlides.set(size - 1, prev);
            rebuildPptxPanel();
        }
    }

    private void moveSlideDown() {
        // Swap the first slide with the second slide
        if (pptxSlides.size() > 1) {
            JTextArea first = pptxSlides.get(0);
            JTextArea second = pptxSlides.get(1);
            pptxSlides.set(0, second);
            pptxSlides.set(1, first);
            rebuildPptxPanel();
        }
    }

    private void rebuildPptxPanel() {
        pptxPanel.removeAll();
        for (int i = 0; i < pptxSlides.size(); i++) {
            JLabel label = new JLabel("Slide " + (i + 1));
            pptxPanel.add(label);
            pptxPanel.add(new JScrollPane(pptxSlides.get(i)));
        }
        pptxPanel.revalidate();
        pptxPanel.repaint();
        updatePPTXPreview();
    }

    // --- Existing methods, updated if needed ---

    // Adds a new slide
    private void addSlide() {
        JTextArea area = new JTextArea();
        area.setFont(new Font("SansSerif", Font.PLAIN, 14));
        pptxSlides.add(area);

        JLabel slideLabel = new JLabel("Slide " + pptxSlides.size());
        pptxPanel.add(slideLabel);
        pptxPanel.add(new JScrollPane(area));
        updatePPTXPreview();
        pptxPanel.revalidate();
        pptxPanel.repaint();
    }

    // Removes the last slide (ensuring at least one remains)
    private void removeCurrentSlide() {
        if (pptxSlides.size() <= 1) {
            JOptionPane.showMessageDialog(this, "At least one slide is required");
            return;
        }
        int idx = pptxSlides.size() - 1;
        // Remove components in reverse order, assuming each slide is added as a pair of
        // components: label then scroll pane.
        pptxPanel.remove(pptxPanel.getComponentCount() - 1); // Remove JScrollPane
        pptxPanel.remove(pptxPanel.getComponentCount() - 1); // Remove JLabel
        pptxSlides.remove(idx);
        updatePPTXPreview();
        pptxPanel.revalidate();
        pptxPanel.repaint();
    }

    // Rebuilds the preview area based on the current slides
    private void updatePPTXPreview() {
        pptxPreviewPanel.removeAll();
        // Retrieve a dynamic background color from UIManager so it adapts to dark/light
        // theme.
        Color previewColor = UIManager.getColor("Panel.background");

        for (int i = 0; i < pptxSlides.size(); i++) {
            JLabel label = new JLabel("Slide " + (i + 1));
            label.setOpaque(true);
            label.setBackground(previewColor);
            pptxPreviewPanel.add(label);
        }

        pptxPreviewPanel.revalidate();
        pptxPreviewPanel.repaint();
    }

    // Appends a bar chart placeholder to the last slide
    private void addBarChartToCurrentSlide() {
        if (pptxSlides.isEmpty())
            return;
        JTextArea area = pptxSlides.get(pptxSlides.size() - 1);
        area.append("\n[Bar Chart]\nA: ###\nB: ######\nC: ##\n");
    }

    // Appends a clipart placeholder to the last slide
    private void addClipartToCurrentSlide() {
        if (pptxSlides.isEmpty())
            return;
        JTextArea area = pptxSlides.get(pptxSlides.size() - 1);
        area.append("\n[Clipart] 😀 🏖️ 🌳\n");
    }

    // Appends bullet points to the last slide
    private void addBulletsToCurrentSlide() {
        if (pptxSlides.isEmpty())
            return;
        JTextArea area = pptxSlides.get(pptxSlides.size() - 1);
        area.append("\n• Bullet Point 1\n• Bullet Point 2\n• Bullet Point 3\n");
    }

    // Allows the user to change the background color of all slide areas
    private void changeBackgroundColor() {
        Color c = JColorChooser.showDialog(this, "Pick Slide Background Color", Color.WHITE);
        if (c != null) {
            pptxPanel.setBackground(c);
            // Use classic instanceof checking
            for (Component comp : pptxPanel.getComponents()) {
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

    // Inserts a text box (as a JTextField) into the PPTX panel
    private void insertTextBox() {
        JTextField textField = new JTextField("Enter text", 20);
        pptxPanel.add(textField);
        pptxPanel.revalidate();
        pptxPanel.repaint();
    }

    // --- Additional editing options ---

    // Duplicates the last slide
    private void duplicateSlide() {
        if (pptxSlides.isEmpty())
            return;
        JTextArea lastSlide = pptxSlides.get(pptxSlides.size() - 1);
        JTextArea newSlide = new JTextArea(lastSlide.getText());
        newSlide.setFont(lastSlide.getFont());
        pptxSlides.add(newSlide);
        JLabel label = new JLabel("Slide " + pptxSlides.size());
        pptxPanel.add(label);
        pptxPanel.add(new JScrollPane(newSlide));
        updatePPTXPreview();
        pptxPanel.revalidate();
        pptxPanel.repaint();
    }

    // Clears the text on the last slide
    private void clearSlide() {
        if (pptxSlides.isEmpty())
            return;
        JTextArea area = pptxSlides.get(pptxSlides.size() - 1);
        int res = JOptionPane.showConfirmDialog(this, "Clear all text on this slide?", "Clear Slide",
                JOptionPane.YES_NO_OPTION);
        if (res == JOptionPane.YES_OPTION) {
            area.setText("");
        }
    }

    // Moves the last slide up one position (if possible)

    // Allows the user to change the font of the last slide's text
    private void setSlideFont() {
        if (pptxSlides.isEmpty())
            return;
        JTextArea current = pptxSlides.get(pptxSlides.size() - 1);
        String fontName = JOptionPane.showInputDialog(this, "Enter font name:", current.getFont().getFamily());
        if (fontName != null && !fontName.isEmpty()) {
            String sizeStr = JOptionPane.showInputDialog(this, "Enter font size:", current.getFont().getSize());
            try {
                int fontSize = Integer.parseInt(sizeStr);
                current.setFont(new Font(fontName, Font.PLAIN, fontSize));
            } catch (NumberFormatException ex) {
                JOptionPane.showMessageDialog(this, "Invalid font size.", "Error", JOptionPane.ERROR_MESSAGE);
            }
        }
    }

    // Allows the user to change the text color of the last slide
    private void setSlideTextColor() {
        if (pptxSlides.isEmpty())
            return;
        JTextArea current = pptxSlides.get(pptxSlides.size() - 1);
        Color newColor = JColorChooser.showDialog(this, "Choose Text Color", current.getForeground());
        if (newColor != null) {
            current.setForeground(newColor);
        }
    }

    private void spellcheck(JTextPane pane) {
        Set<String> dictionary = new HashSet<>(Arrays.asList("the", "and", "java", "public", "class", "void", "main",
                "if", "else", "for", "while", "return"));
        StyledDocument doc = pane.getStyledDocument();

        doc.addDocumentListener(new DocumentListener() {
            public void insertUpdate(DocumentEvent e) {
                check();
            }

            public void removeUpdate(DocumentEvent e) {
                check();
            }

            public void changedUpdate(DocumentEvent e) {
                check();
            }

            private void check() {
                SwingUtilities.invokeLater(() -> {
                    Style defaultStyle = StyleContext.getDefaultStyleContext().getStyle(StyleContext.DEFAULT_STYLE);
                    doc.setCharacterAttributes(0, doc.getLength(), defaultStyle, true);
                    String text = pane.getText();
                    String[] words = text.split("\\s+");
                    int pos = 0;
                    for (String word : words) {
                        if (!dictionary.contains(word.toLowerCase())) {
                            SimpleAttributeSet sas = new SimpleAttributeSet();
                            StyleConstants.setUnderline(sas, true);
                            StyleConstants.setForeground(sas, Color.RED);
                            doc.setCharacterAttributes(pos, word.length(), sas, false);
                        }
                        pos += word.length() + 1;
                    }
                });
            }
        });
    }

    private void enableAutoSave(JTextPane pane) {
        Timer timer = new Timer(30000, e -> {
            File file = fileMap.get(pane);
            if (file != null) {
                try (BufferedWriter w = new BufferedWriter(new FileWriter(file))) {
                    w.write(pane.getText());
                } catch (Exception ignored) {
                }
            }
        });
        timer.start();
    }

    private void exportAsHTML(JTextPane pane) {
        JFileChooser fc = new JFileChooser();
        fc.setSelectedFile(new File("export.html"));
        if (fc.showSaveDialog(this) == JFileChooser.APPROVE_OPTION) {
            try (PrintWriter out = new PrintWriter(fc.getSelectedFile())) {
                out.println("<html><body><pre>");
                out.println(pane.getText());
                out.println("</pre></body></html>");
            } catch (Exception ex) {
                JOptionPane.showMessageDialog(this, "Export failed : " + ex.getMessage());
            }
        }
    }

    private void showTextStatistics(JTextPane pane) {
        String text = pane.getText();
        int words = text.trim().isEmpty() ? 0 : text.trim().split("\\s+").length;
        int chars = text.length();
        Map<String, Integer> freq = new HashMap<>();
        for (String w : text.toLowerCase().split("\\W+")) {
            if (w.isEmpty())
                continue;
            freq.put(w, freq.getOrDefault(w, 0) + 1);
        }
        StringBuilder sb = new StringBuilder("Words: " + words + "\nCharacters : " + chars + "\n\nTop words :\n");
        freq.entrySet().stream().sorted((a, b) -> b.getValue() - a.getValue()).limit(10)
                .forEach(e -> sb.append(e.getKey()).append(": ").append(e.getValue()).append("\n"));
        JOptionPane.showMessageDialog(this, sb.toString());
    }

    private final Set<Integer> bookmarks = new HashSet<>();

    private void toggleBookmark(JTextPane pane) {
        int line = pane.getDocument().getDefaultRootElement().getElementIndex(pane.getCaretPosition());
        if (bookmarks.contains(line))
            bookmarks.remove(line);
        else
            bookmarks.add(line);
        // Optionally, highlight bookmarked lines
    }

    private void findInAllTabs(String query) {
        StringBuilder sb = new StringBuilder();
        for (int i = 0; i < tabbedPane.getTabCount(); i++) {
            Component comp = tabbedPane.getComponentAt(i);
            if (comp instanceof JScrollPane) {
                JScrollPane scroll = (JScrollPane) comp;
                JViewport vp = scroll.getViewport();
                if (vp.getView() instanceof JTextPane) {
                    JTextPane pane = (JTextPane) vp.getView();
                    String text = pane.getText();
                    if (text.contains(query)) {
                        sb.append(tabbedPane.getTitleAt(i)).append(": found ✔️\n");
                    }
                }
            }
        }
        JOptionPane.showMessageDialog(this, sb.length() == 0 ? "Not found 😔" : sb.toString());
    }

    private void printCurrentTab() {
        JTextPane pane = getCurrentTextPane();
        if (pane != null) {
            try {
                pane.print();
            } catch (Exception ex) {
                JOptionPane.showMessageDialog(this, "Print failed ❌: " + ex.getMessage());
            }
        }
    }

    public class CompilerExecutor {
        public static void compileAndRun(File file, TextArea output) {
            String ext = getExt(file);
            String cmd = switch (ext) {
                case "java" -> "javac " + file.getAbsolutePath();
                case "py" -> "python " + file.getAbsolutePath();
                case "c", "cpp" -> "gcc " + file.getAbsolutePath() + " -o output.exe && output.exe";
                default -> null;
            };

            if (cmd == null) {
                output.setText("Unsupported file type: " + ext);
                return;
            }

            try {
                Process process = new ProcessBuilder("bash", "-c", cmd).start(); // on Windows use "cmd", "/c"
                BufferedReader reader = new BufferedReader(
                        new InputStreamReader(process.getInputStream()));

                StringBuilder result = new StringBuilder();
                String line;
                while ((line = reader.readLine()) != null)
                    result.append(line).append("\n");
                reader.close();

                output.setText(result.toString());
            } catch (IOException e) {
                output.setText("Error running code:\n" + e.getMessage());
            }
        }

        private static String getExt(File file) {
            String name = file.getName();
            int dot = name.lastIndexOf('.');
            return (dot > 0) ? name.substring(dot + 1) : "";
        }
    }

    private void applyTheme(Color bg, Color fg, Color caret, Color menuBg, Color menuFg) {
        // Update all open text panes and tables
        for (int i = 0; i < tabbedPane.getTabCount(); i++) {
            Component comp = tabbedPane.getComponentAt(i);
            if (comp instanceof JScrollPane) {
                JScrollPane scroll = (JScrollPane) comp;
                JViewport vp = scroll.getViewport();
                Component view = vp.getView();
                if (view instanceof JTextPane) {
                    JTextPane pane = (JTextPane) view;
                    pane.setBackground(bg);
                    pane.setForeground(fg);
                    pane.setCaretColor(caret);
                } else if (view instanceof JTable) {
                    JTable table = (JTable) view;
                    table.setBackground(bg);
                    table.setForeground(fg);
                    table.setSelectionBackground(menuBg);
                    table.setSelectionForeground(menuFg);
                }
            }
        }
        // Update tabbed pane
        tabbedPane.setBackground(bg);
        tabbedPane.setForeground(fg);

        // Update status bar
        statusBar.setBackground(bg);
        statusBar.setForeground(fg);

        // Update menu bar and menus
        JMenuBar bar = getJMenuBar();
        if (bar != null) {
            bar.setBackground(menuBg);
            bar.setForeground(menuFg);
            for (int i = 0; i < bar.getMenuCount(); i++) {
                JMenu menu = bar.getMenu(i);
                if (menu != null) {
                    menu.setBackground(menuBg);
                    menu.setForeground(menuFg);
                    for (int j = 0; j < menu.getItemCount(); j++) {
                        JMenuItem item = menu.getItem(j);
                        if (item != null) {
                            item.setBackground(menuBg);
                            item.setForeground(menuFg);
                        }
                    }
                }
            }
        }

        SwingUtilities.updateComponentTreeUI(this);
    }

    private void applyFlatLafTweaks() {
        UIManager.put("Component.arc", 20);
        UIManager.put("Button.arc", 20);
        UIManager.put("TextComponent.arc", 20);
        UIManager.put("ScrollBar.thumbArc", 16);
    }

    private void setDarkTheme() {
        System.setProperty("apple.laf.useScreenMenuBar", "true");
        System.setProperty("apple.awt.application.name", "Notepad++ Java");
        System.setProperty("apple.awt.application.appearance", "system");

        boolean useDarkMode = true; // Change to false for light mode

        try {
            if (useDarkMode) {
                FlatMacDarkLaf.setup();
                UIManager.setLookAndFeel(new FlatMacDarkLaf());
            } else {
                FlatMacLightLaf.setup();
                UIManager.setLookAndFeel(new FlatMacLightLaf());
            }

            // 🎨 Rounded macOS-style components
            UIManager.put("Component.arc", 20);
            UIManager.put("Button.arc", 20);
            UIManager.put("TextComponent.arc", 20);
            UIManager.put("ScrollBar.thumbArc", 16);

        } catch (UnsupportedLookAndFeelException e) {
            e.printStackTrace();
        }

    }

    private void setLightTheme() {
        System.setProperty("apple.laf.useScreenMenuBar", "true");
        System.setProperty("apple.awt.application.name", "Notepad++ Java");
        System.setProperty("apple.awt.application.appearance", "system");

        String theme = "default"; // Set the default theme

        try {
            switch (theme.toLowerCase()) {
                case "dark":
                    FlatMacDarkLaf.setup();
                    UIManager.setLookAndFeel(new FlatMacDarkLaf());
                    break;

                case "light":
                    FlatMacLightLaf.setup();
                    UIManager.setLookAndFeel(new FlatMacLightLaf());
                    break;

                default:
                    FlatMacLightLaf.setup();
                    UIManager.setLookAndFeel(new FlatMacLightLaf());
                    break;
            }

            // Apply UI component styles
            UIManager.put("Component.arc", 20);
            UIManager.put("Button.arc", 20);
            UIManager.put("TextComponent.arc", 20);
            UIManager.put("ScrollBar.thumbArc", 16);

        } catch (UnsupportedLookAndFeelException e) {
            e.printStackTrace();
        }

        // Launch UI on the Swing thread

    }

   private void toUppercaseAll() {
    Component comp = tabbedPane.getSelectedComponent();
    String type = tabTypeMap.getOrDefault(comp, "text");
    
    if ("text".equals(type)) {
        optimizeTextUppercase();
    } else if ("table".equals(type)) {
        optimizeTableUppercase();
    }
}

// OPTIMIZED TEXT VERSION - O(1) space, O(n) time
private void optimizeTextUppercase() {
    JTextPane pane = getCurrentTextPane();
    if (pane == null) return;
    
    String selectedText = pane.getSelectedText();
    if (selectedText != null) {
        // In-place replacement - no new objects for small selections
        if (selectedText.length() < 1000) {
            pane.replaceSelection(selectedText.toUpperCase());
        } else {
            // For large selections, use more memory-efficient approach
            processLargeTextSelection(pane, selectedText);
        }
    }
}

// OPTIMIZED TABLE VERSION - O(1) extra space, O(n) time
private void optimizeTableUppercase() {
    JTable table = getCurrentTable();
    if (table == null) return;
    
    TableModel model = table.getModel();
    int[] rows = getSelectedOrAllRows(table);
    int[] cols = getSelectedOrAllColumns(table);
    
    // Early exit if nothing to process
    if (rows.length == 0 || cols.length == 0) return;
    
    // Use batch processing for table model
    if (model instanceof InfiniteStyledTableModel) {
        ((InfiniteStyledTableModel) model).batchUpdate(() -> {
            processTableCells(table, model, rows, cols);
        });
    } else {
        processTableCells(table, model, rows, cols);
    }
}

// OPTIMIZED: Get rows/columns without creating unnecessary arrays
private int[] getSelectedOrAllRows(JTable table) {
    int[] selectedRows = table.getSelectedRows();
    if (selectedRows.length > 0) {
        return selectedRows;
    }
    
    // Create array directly without intermediate list
    int rowCount = table.getRowCount();
    int[] allRows = new int[rowCount];
    for (int i = 0; i < rowCount; i++) {
        allRows[i] = i;
    }
    return allRows;
}

private int[] getSelectedOrAllColumns(JTable table) {
    int[] selectedCols = table.getSelectedColumns();
    if (selectedCols.length > 0) {
        return selectedCols;
    }
    
    int colCount = table.getColumnCount();
    int[] allCols = new int[colCount];
    for (int i = 0; i < colCount; i++) {
        allCols[i] = i;
    }
    return allCols;
}

// OPTIMIZED: Process table cells with minimal object creation
private void processTableCells(JTable table, TableModel model, int[] rows, int[] cols) {
    // Reusable objects to avoid GC pressure
    final StringBuilder reusableBuilder = new StringBuilder(256);
    
    for (int row : rows) {
        for (int col : cols) {
            Object val = table.getValueAt(row, col);
            
            if (val instanceof StyledDocument) {
                processStyledDocument((StyledDocument) val, reusableBuilder);
            } else if (val != null) {
                processRegularValue(table, model, row, col, val, reusableBuilder);
            }
        }
    }
}

// OPTIMIZED: Process StyledDocument with minimal object creation
private void processStyledDocument(StyledDocument doc, StringBuilder reusableBuilder) {
    try {
        int length = doc.getLength();
        if (length == 0) return;
        
        // Get text and convert to uppercase
        String text = doc.getText(0, length);
        
        // Clear and reuse StringBuilder
        reusableBuilder.setLength(0);
        reusableBuilder.append(text);
        
        // Convert to uppercase in-place in StringBuilder
        for (int i = 0; i < reusableBuilder.length(); i++) {
            char c = reusableBuilder.charAt(i);
            if (c >= 'a' && c <= 'z') {
                reusableBuilder.setCharAt(i, (char) (c - 32));
            }
        }
        
        // Replace document content
        doc.remove(0, length);
        doc.insertString(0, reusableBuilder.toString(), null);
        
    } catch (Exception ignored) {
        // Silent fail - same as original
    }
}

// OPTIMIZED: Process regular values efficiently
private void processRegularValue(JTable table, TableModel model, int row, int col, 
                                Object val, StringBuilder reusableBuilder) {
    String str = val.toString();
    
    // For short strings, use direct conversion
    if (str.length() <= 64) {
        table.setValueAt(str.toUpperCase(), row, col);
    } else {
        // For long strings, use more efficient conversion
        reusableBuilder.setLength(0);
        reusableBuilder.append(str);
        
        // Manual uppercase conversion to avoid creating intermediate strings
        for (int i = 0; i < reusableBuilder.length(); i++) {
            char c = reusableBuilder.charAt(i);
            if (c >= 'a' && c <= 'z') {
                reusableBuilder.setCharAt(i, (char) (c - 32));
            }
        }
        table.setValueAt(reusableBuilder.toString(), row, col);
    }
}

// OPTIMIZED: Process large text selections with memory efficiency
private void processLargeTextSelection(JTextPane pane, String selectedText) {
    // Process in chunks to avoid large string allocations
    final int CHUNK_SIZE = 10000;
    int length = selectedText.length();
    
    if (length <= CHUNK_SIZE) {
        pane.replaceSelection(selectedText.toUpperCase());
        return;
    }
    
    // Process and replace in chunks
    StringBuilder result = new StringBuilder(length);
    for (int i = 0; i < length; i += CHUNK_SIZE) {
        int end = Math.min(i + CHUNK_SIZE, length);
        String chunk = selectedText.substring(i, end);
        result.append(chunk.toUpperCase());
    }
    
    pane.replaceSelection(result.toString());
}
    // Remove Duplicate Lines for all formats
    private void removeDuplicateLinesAll() {
        Component comp = tabbedPane.getSelectedComponent();
        String type = tabTypeMap.getOrDefault(comp, "text");
        if ("text".equals(type)) {
            removeDuplicateLines();
        } else if ("table".equals(type)) {
            JTable table = getCurrentTable();
            if (table != null) {
                java.util.Set<String> seen = new LinkedHashSet<>();
                java.util.List<String[]> uniqueRows = new ArrayList<>();
                for (int row = 0; row < table.getRowCount(); row++) {
                    StringBuilder sb = new StringBuilder();
                    for (int col = 0; col < table.getColumnCount(); col++) {
                        Object val = table.getValueAt(row, col);
                        String txt = "";
                        if (val instanceof StyledDocument) {
                            try {
                                StyledDocument doc = (StyledDocument) val;
                                txt = doc.getText(0, doc.getLength());
                            } catch (Exception ignored) {
                            }
                        } else if (val != null) {
                            txt = val.toString();
                        }
                        sb.append(txt).append("\t");
                    }
                    String rowStr = sb.toString();
                    if (seen.add(rowStr)) {
                        uniqueRows.add(rowStr.split("\t", -1));
                    }
                }
                int row = 0;
                for (String[] cells : uniqueRows) {
                    for (int col = 0; col < cells.length && col < table.getColumnCount(); col++) {
                        table.setValueAt(cells[col], row, col);
                    }
                    row++;
                }
                // Clear remaining rows
                for (; row < table.getRowCount(); row++) {
                    for (int col = 0; col < table.getColumnCount(); col++) {
                        table.setValueAt("", row, col);
                    }
                }
            }
        }
    }

    // Sort Lines for all formats
    private void sortLinesAll() {
        Component comp = tabbedPane.getSelectedComponent();
        String type = tabTypeMap.getOrDefault(comp, "text");
        if ("text".equals(type)) {
            sortLines();
        } else if ("table".equals(type)) {
            JTable table = getCurrentTable();
            if (table != null) {
                java.util.List<String[]> rows = new ArrayList<>();
                for (int row = 0; row < table.getRowCount(); row++) {
                    String[] cells = new String[table.getColumnCount()];
                    for (int col = 0; col < table.getColumnCount(); col++) {
                        Object val = table.getValueAt(row, col);
                        if (val instanceof StyledDocument) {
                            try {
                                StyledDocument doc = (StyledDocument) val;
                                cells[col] = doc.getText(0, doc.getLength());
                            } catch (Exception ignored) {
                            }
                        } else if (val != null) {
                            cells[col] = val.toString();
                        } else {
                            cells[col] = "";
                        }
                    }
                    rows.add(cells);
                }
                rows.sort((a, b) -> String.join("\t", a).compareToIgnoreCase(String.join("\t", b)));
                for (int row = 0; row < rows.size(); row++) {
                    for (int col = 0; col < table.getColumnCount(); col++) {
                        table.setValueAt(rows.get(row)[col], row, col);
                    }
                }
            }
        }
    }

    private void showFindDialog() {
        JOptionPane.showMessageDialog(this, "Find dialog not implemented yet ❌.");
    }

    private void showReplaceDialog() {
        JOptionPane.showMessageDialog(this, "Replace dialog not implemented yet ❌.");
    }

    private void showGoToDialog() {
        JTextPane pane = getCurrentTextPane();
        if (pane == null)
            return;

        JPanel panel = new JPanel();
        panel.add(new JLabel("Line number:"));
        JTextField lineField = new JTextField(10);
        panel.add(lineField);

        int result = JOptionPane.showConfirmDialog(this, panel, "Go To Line", JOptionPane.OK_CANCEL_OPTION,
                JOptionPane.PLAIN_MESSAGE);

        if (result == JOptionPane.OK_OPTION) {
            try {
                int lineNum = Integer.parseInt(lineField.getText());
                // In Swing, line numbers are 0-indexed
                int offset = pane.getDocument().getDefaultRootElement().getElement(lineNum - 1).getStartOffset();
                pane.setCaretPosition(offset);
            } catch (NumberFormatException | NullPointerException ex) {
                // If input is not a valid number, shake the input field!
                shakeComponent(lineField);
            }
        }
    }

    private void reverseLines() {
        JTextPane pane = getCurrentTextPane();
        if (pane != null) {
            String[] lines = pane.getText().split("\n");
            java.util.Collections.reverse(Arrays.asList(lines));
            pane.setText(String.join("\n", lines));
        }
    }

    private void removeAllNumbers() {
        JTextPane pane = getCurrentTextPane();
        if (pane != null) {
            String text = pane.getText().replaceAll("\\d+", "");
            pane.setText(text);
        }
    }

    private void removeAllPunctuation() {
        JTextPane pane = getCurrentTextPane();
        if (pane != null) {
            String text = pane.getText().replaceAll("\\p{Punct}", "");
            pane.setText(text);
        }
    }

    private void shuffleLines() {
        JTextPane pane = getCurrentTextPane();
        if (pane != null) {
            String[] lines = pane.getText().split("\n");
            java.util.List<String> list = Arrays.asList(lines);
            java.util.Collections.shuffle(list);
            pane.setText(String.join("\n", list));
        }
    }

    private void countWordsAndChars() {
        JTextPane pane = getCurrentTextPane();
        if (pane != null) {
            String text = pane.getText();
            int words = text.trim().isEmpty() ? 0 : text.trim().split("\\s+").length;
            int chars = text.length();
            JOptionPane.showMessageDialog(this, "Words: " + words + "\nCharacters 🔣: " + chars);
        }
    }

    private void removeDuplicateWords() {
    JTextPane pane = getCurrentTextPane();
    if (pane != null) {
        String text = pane.getText();
        if (text.isEmpty()) return;
        
        // Use more efficient approach for large texts
        if (text.length() > 10000) {
            removeDuplicateWordsLarge(text, pane);
        } else {
            removeDuplicateWordsSmall(text, pane);
        }
    }
}

// O(n) time, O(k) space where k = unique words
private void removeDuplicateWordsSmall(String text, JTextPane pane) {
    // Use manual parsing to avoid creating large arrays
    StringBuilder result = new StringBuilder(text.length());
    Set<String> seenWords = new HashSet<>();
    
    int start = 0;
    int end = 0;
    int length = text.length();
    
    while (end < length) {
        // Find word boundary
        while (end < length && !Character.isWhitespace(text.charAt(end))) {
            end++;
        }
        
        // Extract word
        if (end > start) {
            String word = text.substring(start, end);
            
            // Check if we've seen this word
            if (seenWords.add(word)) {
                if (result.length() > 0) {
                    result.append(' ');
                }
                result.append(word);
            }
        }
        
        // Skip whitespace
        while (end < length && Character.isWhitespace(text.charAt(end))) {
            end++;
        }
        start = end;
    }
    
    pane.setText(result.toString());
}

// O(n) time, O(1) extra space for very large texts
private void removeDuplicateWordsLarge(String text, JTextPane pane) {
    // Process in chunks to avoid memory issues
    final int CHUNK_SIZE = 50000; // Process 50K characters at a time
    StringBuilder result = new StringBuilder(text.length());
    Set<String> seenWords = new HashSet<>(1000);
    
    int chunkStart = 0;
    int textLength = text.length();
    
    while (chunkStart < textLength) {
        int chunkEnd = Math.min(chunkStart + CHUNK_SIZE, textLength);
        
        // Ensure we don't break in the middle of a word
        while (chunkEnd < textLength && !Character.isWhitespace(text.charAt(chunkEnd))) {
            chunkEnd++;
        }
        
        String chunk = text.substring(chunkStart, chunkEnd);
        processChunk(chunk, seenWords, result, chunkStart > 0);
        
        chunkStart = chunkEnd;
    }
    
    pane.setText(result.toString());
}

private void processChunk(String chunk, Set<String> seenWords, StringBuilder result, boolean addSpace) {
    int start = 0;
    int end = 0;
    int length = chunk.length();
    
    while (end < length) {
        // Find word boundary
        while (end < length && !Character.isWhitespace(chunk.charAt(end))) {
            end++;
        }
        
        // Extract word
        if (end > start) {
            String word = chunk.substring(start, end);
            
            if (seenWords.add(word)) {
                if (result.length() > 0 || addSpace) {
                    result.append(' ');
                }
                result.append(word);
            }
        }
        
        // Skip whitespace
        while (end < length && Character.isWhitespace(chunk.charAt(end))) {
            end++;
        }
        start = end;
    }
}

   private void showFindWordDialog() {
    JTextPane pane = getCurrentTextPane();
    if (pane == null) return;
    
    String word = JOptionPane.showInputDialog(this, "Enter word to find:");
    if (word == null || word.isEmpty()) return;
    
    // Use more efficient search algorithm
    if (word.length() > 3) { // Use Boyer-Moore for longer patterns
        findAndHighlightOptimized(pane, word);
    } else { // Use indexOf for short patterns
        findAndHighlightSimple(pane, word);
    }
}

// O(n) time, O(1) space - Boyer-Moore for longer patterns
private void findAndHighlightOptimized(JTextPane pane, String word) {
    Highlighter highlighter = pane.getHighlighter();
    highlighter.removeAllHighlights();
    
    String text = pane.getText();
    if (text.isEmpty()) return;
    
    // Use Boyer-Moore search for longer patterns
    List<Integer> matches = boyerMooreSearchCaseInsensitive(text, word);
    
    // Highlight all matches
    Highlighter.HighlightPainter painter = 
        new DefaultHighlighter.DefaultHighlightPainter(Color.YELLOW);
    
    for (int index : matches) {
        try {
            highlighter.addHighlight(index, index + word.length(), painter);
        } catch (Exception ignored) {
        }
    }
    
    // Show result count
    if (!matches.isEmpty()) {
        showAnimatedNotification("Found " + matches.size() + " occurrences", new Color(76, 175, 80));
    } else {
        showAnimatedNotification("No matches found", new Color(244, 67, 54));
    }
}

// O(n) time, O(1) space - Simple search for short patterns
private void findAndHighlightSimple(JTextPane pane, String word) {
    Highlighter highlighter = pane.getHighlighter();
    highlighter.removeAllHighlights();
    
    String text = pane.getText();
    if (text.isEmpty()) return;
    
    // Reusable objects to avoid GC
    final String lowerWord = word.toLowerCase();
    Highlighter.HighlightPainter painter = 
        new DefaultHighlighter.DefaultHighlightPainter(Color.YELLOW);
    
    int count = 0;
    int index = 0;
    int textLength = text.length();
    int wordLength = word.length();
    
    // Manual case-insensitive search to avoid creating lowercase string
    while (index <= textLength - wordLength) {
        boolean match = true;
        
        // Check if characters match (case-insensitive)
        for (int i = 0; i < wordLength; i++) {
            char textChar = Character.toLowerCase(text.charAt(index + i));
            char wordChar = lowerWord.charAt(i);
            if (textChar != wordChar) {
                match = false;
                break;
            }
        }
        
        if (match) {
            try {
                highlighter.addHighlight(index, index + wordLength, painter);
                count++;
            } catch (Exception ignored) {
            }
            index += wordLength; // Skip past this match
        } else {
            index++;
        }
    }
    
    // Show result count
    if (count > 0) {
        showAnimatedNotification("Found " + count + " occurrences", new Color(76, 175, 80));
    } else {
        showAnimatedNotification("No matches found", new Color(244, 67, 54));
    }
}

// O(n/m) time Boyer-Moore search with case insensitivity
private List<Integer> boyerMooreSearchCaseInsensitive(String text, String pattern) {
    List<Integer> indices = new ArrayList<>();
    if (pattern.isEmpty()) return indices;
    
    String lowerPattern = pattern.toLowerCase();
    int patternLength = pattern.length();
    
    // Preprocessing for bad character rule (case-insensitive)
    int[] badChar = new int[256];
    Arrays.fill(badChar, -1);
    for (int i = 0; i < patternLength; i++) {
        badChar[Character.toLowerCase(lowerPattern.charAt(i))] = i;
    }
    
    int s = 0;
    while (s <= text.length() - patternLength) {
        int j = patternLength - 1;
        
        // Case-insensitive comparison
        while (j >= 0 && 
               Character.toLowerCase(text.charAt(s + j)) == Character.toLowerCase(lowerPattern.charAt(j))) {
            j--;
        }
        
        if (j < 0) {
            indices.add(s);
            s += (s + patternLength < text.length()) ? 
                 patternLength - badChar[Character.toLowerCase(text.charAt(s + patternLength))] : 1;
        } else {
            s += Math.max(1, j - badChar[Character.toLowerCase(text.charAt(s + j))]);
        }
    }
    return indices;
}
   // ADD THESE REUSABLE OBJECTS at class level:
private final Font[] fontCache = new Font[100]; // Cache for common font sizes
private StyledEditorKit.BoldAction boldAction;
private StyledEditorKit.ItalicAction italicAction;
private StyledEditorKit.UnderlineAction underlineAction;
private final SimpleAttributeSet reusableAttrSet = new SimpleAttributeSet();

// Initialize reusable actions once
private void initializeStyleActions() {
    if (boldAction == null) {
        boldAction = new StyledEditorKit.BoldAction();
        italicAction = new StyledEditorKit.ItalicAction();
        underlineAction = new StyledEditorKit.UnderlineAction();
    }
}

// OPTIMIZED: Font operations with caching
private void changeFontSize(int delta) {
    JTextPane pane = getCurrentTextPane();
    if (pane != null) {
        Font currentFont = pane.getFont();
        int newSize = Math.max(2, currentFont.getSize() + delta);
        
        // Use cached font if available
        Font newFont = getCachedFont(currentFont, newSize);
        pane.setFont(newFont);
    }
}

private void setFontSize(int size) {
    JTextPane pane = getCurrentTextPane();
    if (pane != null) {
        Font currentFont = pane.getFont();
        
        // Use cached font if available
        Font newFont = getCachedFont(currentFont, size);
        pane.setFont(newFont);
    }
}

// FONT CACHE: O(1) time for common sizes
private Font getCachedFont(Font baseFont, int newSize) {
    // Bounds check
    if (newSize < 2 || newSize >= fontCache.length) {
        return baseFont.deriveFont((float) newSize);
    }
    
    // Check cache
    Font cached = fontCache[newSize];
    if (cached != null && 
        cached.getFamily().equals(baseFont.getFamily()) && 
        cached.getStyle() == baseFont.getStyle()) {
        return cached;
    }
    
    // Create and cache new font
    Font newFont = baseFont.deriveFont((float) newSize);
    fontCache[newSize] = newFont;
    return newFont;
}

// OPTIMIZED: Unified style toggle method
private void setBold() {
    toggleTextStyle(StyleConstants.Bold, true);
}

private void setItalic() {
    toggleTextStyle(StyleConstants.Italic, true);
}

private void setUnderline() {
    toggleTextStyle(StyleConstants.Underline, true);
}

// SINGLE OPTIMIZED METHOD for all style toggles - O(1) space
private void toggleTextStyle(Object styleConstant, boolean isTextPane) {
    if (isTextPane) {
        toggleTextPaneStyle(styleConstant);
    } else {
        toggleTableEditorStyle(styleConstant);
    }
}

// OPTIMIZED: Text pane style with object reuse
private void toggleTextPaneStyle(Object styleConstant) {
    JTextPane pane = getCurrentTextPane();
    if (pane == null) return;
    
    StyledDocument doc = pane.getStyledDocument();
    int start = pane.getSelectionStart();
    int end = pane.getSelectionEnd();
    if (start == end) return;
    
    // REUSE attribute set instead of creating new one
    Element element = doc.getCharacterElement(start);
    reusableAttrSet.removeAttributes(reusableAttrSet);
    reusableAttrSet.addAttributes(element.getAttributes());
    
    // Toggle the specific style
    boolean currentState = getStyleState(reusableAttrSet, styleConstant);
    setStyleState(reusableAttrSet, styleConstant, !currentState);
    
    doc.setCharacterAttributes(start, end - start, reusableAttrSet, false);
}

// OPTIMIZED: Table editor style with reusable actions
private void toggleTableEditorStyle(Object styleConstant) {
    JTable table = getCurrentTable();
    if (table == null || !table.isEditing()) return;
    
    Component editor = table.getEditorComponent();
    if (!(editor instanceof JTextPane)) return;
    
    // Initialize actions once
    initializeStyleActions();
    
    // Use pre-created action objects
    switch (styleConstant.toString()) {
        case "bold":
            boldAction.actionPerformed(null);
            break;
        case "italic":
            italicAction.actionPerformed(null);
            break;
        case "underline":
            underlineAction.actionPerformed(null);
            break;
    }
}

// OPTIMIZED: Get style state without object creation
private boolean getStyleState(AttributeSet attr, Object styleConstant) {
    if (styleConstant == StyleConstants.Bold) {
        return StyleConstants.isBold(attr);
    } else if (styleConstant == StyleConstants.Italic) {
        return StyleConstants.isItalic(attr);
    } else if (styleConstant == StyleConstants.Underline) {
        return StyleConstants.isUnderline(attr);
    }
    return false;
}

// OPTIMIZED: Set style state without object creation
private void setStyleState(MutableAttributeSet attr, Object styleConstant, boolean value) {
    if (styleConstant == StyleConstants.Bold) {
        StyleConstants.setBold(attr, value);
    } else if (styleConstant == StyleConstants.Italic) {
        StyleConstants.setItalic(attr, value);
    } else if (styleConstant == StyleConstants.Underline) {
        StyleConstants.setUnderline(attr, value);
    }
}

// BATCH OPERATION: Apply multiple styles at once
private void applyTextStyles(Map<Object, Boolean> styles) {
    JTextPane pane = getCurrentTextPane();
    if (pane == null) return;
    
    StyledDocument doc = pane.getStyledDocument();
    int start = pane.getSelectionStart();
    int end = pane.getSelectionEnd();
    if (start == end) return;
    
    // Single attribute set for all style changes
    Element element = doc.getCharacterElement(start);
    reusableAttrSet.removeAttributes(reusableAttrSet);
    reusableAttrSet.addAttributes(element.getAttributes());
    
    // Apply all styles in one operation
    for (Map.Entry<Object, Boolean> entry : styles.entrySet()) {
        setStyleState(reusableAttrSet, entry.getKey(), entry.getValue());
    }
    
    doc.setCharacterAttributes(start, end - start, reusableAttrSet, false);
}
// Method to handle multiple style operations efficiently
public void setBoldItalicUnderline(boolean bold, boolean italic, boolean underline) {
    Map<Object, Boolean> styles = new HashMap<>(3);
    styles.put(StyleConstants.Bold, bold);
    styles.put(StyleConstants.Italic, italic);
    styles.put(StyleConstants.Underline, underline);
    applyTextStyles(styles);
}

// Method to reset all formatting
public void resetFormatting() {
    JTextPane pane = getCurrentTextPane();
    if (pane != null) {
        StyledDocument doc = pane.getStyledDocument();
        int start = pane.getSelectionStart();
        int end = pane.getSelectionEnd();
        if (start == end) return;
        
        // Create a clean attribute set
        reusableAttrSet.removeAttributes(reusableAttrSet);
        doc.setCharacterAttributes(start, end - start, reusableAttrSet, false);
    }
}
    private void addTableTab() {
        String[] columns = { "A", "B", "C", "D", "E" };
        Object[][] data = new Object[20][columns.length];

        for (int r = 0; r < data.length; r++) {
            for (int c = 0; c < columns.length; c++) {
                data[r][c] = new DefaultStyledDocument();
            }
        }

        JTable table = new JTable(data, columns) {
            @Override
            public boolean isCellEditable(int row, int col) {
                return true;
            }

            @Override
            public Class<?> getColumnClass(int col) {
                return StyledDocument.class;
            }
        };

        table.setDefaultRenderer(StyledDocument.class, new StyledCellRenderer());
        table.setDefaultEditor(StyledDocument.class, new StyledCellEditor());
        JScrollPane scroll = new JScrollPane(table);

        // Load the Excel icon from the resource folder.
        ImageIcon originalIcon = new ImageIcon(getClass().getResource("/icons/excel.png"));
        int fixedWidth = 16;
        int fixedHeight = 16;
        Image scaledImage = originalIcon.getImage().getScaledInstance(fixedWidth, fixedHeight, Image.SCALE_SMOOTH);
        Icon logoIcon = new ImageIcon(scaledImage);

        // Add the tab with the title, fixed-size logo, and the scroll pane containing
        // the table.
        tabbedPane.addTab("Table", logoIcon, scroll);
        tabbedPane.setSelectedComponent(scroll);
        tabTypeMap.put(scroll, "table");
    }

    static class StyledCellRenderer extends JTextPane implements TableCellRenderer {
        public StyledCellRenderer() {
            setEditable(false);
            setOpaque(true);
        }

        @Override
        public Component getTableCellRendererComponent(JTable table, Object value, boolean isSelected,
                boolean hasFocus, int row, int column) {
            if (value instanceof StyledDocument) {
                setStyledDocument((StyledDocument) value);
            } else {
                setText(value == null ? "" : value.toString());
            }
            setBackground(isSelected ? table.getSelectionBackground() : table.getBackground());
            setForeground(isSelected ? table.getSelectionForeground() : table.getForeground());
            setBorder(BorderFactory.createLineBorder(Color.LIGHT_GRAY));
            return this;
        }
    }

    static class StyledCellEditor extends AbstractCellEditor implements TableCellEditor {
        private final JTextPane editor = new JTextPane();

        public StyledCellEditor() {
            InputMap im = editor.getInputMap();
            ActionMap am = editor.getActionMap();
            im.put(KeyStroke.getKeyStroke(KeyEvent.VK_B, InputEvent.CTRL_DOWN_MASK), "bold");
            im.put(KeyStroke.getKeyStroke(KeyEvent.VK_I, InputEvent.CTRL_DOWN_MASK), "italic");
            im.put(KeyStroke.getKeyStroke(KeyEvent.VK_U, InputEvent.CTRL_DOWN_MASK), "underline");
            am.put("bold", new StyledEditorKit.BoldAction());
            am.put("italic", new StyledEditorKit.ItalicAction());
            am.put("underline", new StyledEditorKit.UnderlineAction());
        }

        @Override
        public Object getCellEditorValue() {
            StyledDocument doc = editor.getStyledDocument();
            DefaultStyledDocument copy = new DefaultStyledDocument();
            try {
                copy.insertString(0, doc.getText(0, doc.getLength()), doc.getCharacterElement(0).getAttributes());
            } catch (Exception ignored) {
            }
            return copy;
        }

        @Override
        public Component getTableCellEditorComponent(JTable table, Object value, boolean isSelected, int row,
                int column) {
            if (value instanceof StyledDocument) {
                editor.setStyledDocument((StyledDocument) value);
            } else {
                editor.setText(value == null ? "" : value.toString());
            }
            return editor;
        }
    }

    private void openFile() {
        JFileChooser fc = new JFileChooser();
        if (fc.showOpenDialog(this) != JFileChooser.APPROVE_OPTION)
            return;

        File file = fc.getSelectedFile();
        String editorType = fileTypeRegistry.getEditorTypeForFile(file.getName());

        switch (editorType) {
            case "pptx_editable":
                addEditablePPTXTab(file);
                break;
            case "table":
                createSpreadsheetTab(file);
                break;
            case "image":
                openImageFile(file);
                break;
            default:
                com.example.notepadpp.media.FileOpenHandler.open(file, tabbedPane, tabTypeMap);
        }
    }

    private void addEditablePPTXTab(File file) {
        JPanel panel = new com.example.notepadpp.media.PptxEditorPanel(file);
        JScrollPane scroll = new JScrollPane(panel);
        tabbedPane.addTab(file.getName(), com.example.notepadpp.media.PptxEditorPanel.getTabIcon(), scroll);
        tabTypeMap.put(scroll, "pptx_editable");
    }

 // In your main class where you handle file opening
// In your main application class


private void createNewTextTab() {
    JTextArea textArea = new JTextArea();
    JScrollPane scrollPane = new JScrollPane(textArea);
    
    // Create tab with icon and close button
    tabbedPane.addTab("New Tab", ImageView.getNewTabIcon(), scrollPane);
    
    // Add close button to the tab
    addCloseButtonToTab(tabbedPane.getTabCount() - 1);
}

private void createNewTableTab() {
    JTable table = new JTable();
    JScrollPane scrollPane = new JScrollPane(table);
    
    // Create table tab with icon and close button
    tabbedPane.addTab("New Table", ImageView.getNewTableTabIcon(), scrollPane);
    
    // Add close button to the tab
    addCloseButtonToTab(tabbedPane.getTabCount() - 1);
}



// Helper method to check for unsaved changes (you'll need to implement this based on your application)

private void openImageFile(File file) {
    try {
        if (!ImageView.isSupportedImageFile(file)) {
            JOptionPane.showMessageDialog(this, "Unsupported image format", "Error", JOptionPane.ERROR_MESSAGE);
            return;
        }
        
        // Use the ImageView class to open the image
        ImageView.ImageViewResult result = ImageView.openImageFile(file, getWidth(), getHeight());
        
        // ✅ Use addTabWithCloseButton function (same as addNewTab)
        addTabWithCloseButton(file.getName(), result.getScrollPane(), ImageView.getTabIcon());
        
        tabTypeMap.put(result.getScrollPane(), "image");
        
    } catch (Exception e) {
        JOptionPane.showMessageDialog(this, "Unable to open image: " + e.getMessage(), "Error", JOptionPane.ERROR_MESSAGE);
    }
}

private void addCloseButtonToTab(int tabIndex) {
    // Store the original title before we replace the tab component
    String originalTitle = tabbedPane.getTitleAt(tabIndex);
    
    // Create a panel with label and close button
    JPanel tabPanel = new JPanel(new FlowLayout(FlowLayout.CENTER, 4, 0));
    tabPanel.setOpaque(false);
    
    JLabel tabLabel = new JLabel(originalTitle);
    JButton closeButton = new JButton(ImageView.getCloseIcon());
    
    // Style the close button (same as in addNewTab pattern)
    closeButton.setPreferredSize(new Dimension(16, 16));
    closeButton.setBorderPainted(false);
    closeButton.setContentAreaFilled(false);
    closeButton.setFocusPainted(false);
    
    // Add hand cursor
    try {
        closeButton.setCursor(Cursor.getPredefinedCursor(Cursor.HAND_CURSOR));
    } catch (Exception e) {
        // If cursor causes issues, just skip it
    }
    
    // Add tooltip
    closeButton.setToolTipText("Close tab");
    
    // Add action to close tab
    closeButton.addActionListener(e -> {
        int index = tabbedPane.indexOfTabComponent(tabPanel);
        if (index != -1) {
            tabbedPane.remove(index);
        }
    });
    
    tabPanel.add(tabLabel);
    tabPanel.add(closeButton);
    
    tabbedPane.setTabComponentAt(tabIndex, tabPanel);
}
// Helper method to check for unsaved changes
private boolean hasUnsavedChanges(int tabIndex) {
    Component tabComponent = tabbedPane.getComponentAt(tabIndex);
    String tabType = tabTypeMap.get(tabComponent);
    
    if ("text".equals(tabType)) {
        // For text tabs, check if the text area has unsaved changes
        JScrollPane scrollPane = (JScrollPane) tabComponent;
        JViewport viewport = scrollPane.getViewport();
        Component view = viewport.getView();
        if (view instanceof JTextArea) {
            JTextArea textArea = (JTextArea) view;
            // You'll need to implement your own logic to track changes
            // For example, you could have a boolean flag that tracks modifications
            return textArea.getDocument().getDefaultRootElement().getElementCount() > 1; // Simple example
        }
    } else if ("image".equals(tabType)) {
        // Image tabs typically don't have unsaved changes unless you've edited the image
        return false;
    } else if ("table".equals(tabType)) {
        // For table tabs, check if table data has been modified
        return false; // Implement your logic
    }
    
    return false;
}

// Method to update tab title when file is saved
public void updateTabTitle(Component tabContent, String newTitle) {
    for (int i = 0; i < tabbedPane.getTabCount(); i++) {
        if (tabbedPane.getComponentAt(i) == tabContent) {
            tabbedPane.setTitleAt(i, newTitle);
            // Update the tab label if we have a custom tab component
            Component tabComp = tabbedPane.getTabComponentAt(i);
            if (tabComp instanceof JPanel) {
                JPanel tabPanel = (JPanel) tabComp;
                // Find the label in the panel and update it
                for (Component comp : tabPanel.getComponents()) {
                    if (comp instanceof JLabel) {
                        ((JLabel) comp).setText(newTitle);
                        break;
                    }
                }
            }
            break;
        }
    }
}

// Method to close all tabs
private void closeAllTabs() {
    // Check for unsaved changes in all tabs
    boolean hasUnsaved = false;
    for (int i = 0; i < tabbedPane.getTabCount(); i++) {
        if (hasUnsavedChanges(i)) {
            hasUnsaved = true;
            break;
        }
    }
    
    if (hasUnsaved) {
        int result = JOptionPane.showConfirmDialog(this, 
            "You have unsaved changes in some tabs. Close all tabs anyway?", 
            "Unsaved Changes", 
            JOptionPane.YES_NO_OPTION);
        if (result != JOptionPane.YES_OPTION) {
            return;
        }
    }
    
    tabbedPane.removeAll();
    tabTypeMap.clear();
}

// Method to close current tab
// Method to close other tabs
private void closeOtherTabs() {
    int currentIndex = tabbedPane.getSelectedIndex();
    if (currentIndex == -1) return;
    
    // Check for unsaved changes in other tabs
    boolean hasUnsaved = false;
    for (int i = tabbedPane.getTabCount() - 1; i >= 0; i--) {
        if (i != currentIndex && hasUnsavedChanges(i)) {
            hasUnsaved = true;
            break;
        }
    }
    
    if (hasUnsaved) {
        int result = JOptionPane.showConfirmDialog(this, 
            "You have unsaved changes in other tabs. Close them anyway?", 
            "Unsaved Changes", 
            JOptionPane.YES_NO_OPTION);
        if (result != JOptionPane.YES_OPTION) {
            return;
        }
    }
    
    // Remove all tabs except the current one
    for (int i = tabbedPane.getTabCount() - 1; i >= 0; i--) {
        if (i != currentIndex) {
            tabbedPane.remove(i);
        }
    }
}

// Setup middle-click to close tabs


// Setup context menu for tabs

// Initialize all tab-related functionality
private void initializeTabFeatures() {
    setupMiddleClickTabClosing();
    setupTabContextMenu();
    
    // Add keyboard shortcut for closing tabs (Ctrl+W)
    tabbedPane.getInputMap(JComponent.WHEN_IN_FOCUSED_WINDOW).put(
        KeyStroke.getKeyStroke(KeyEvent.VK_W, InputEvent.CTRL_DOWN_MASK), "closeCurrentTab");
    tabbedPane.getActionMap().put("closeCurrentTab", new AbstractAction() {
        @Override
        public void actionPerformed(ActionEvent e) {
            closeCurrentTab();
        }
    });
}

// Enhanced method to update tab title when file is saved/renamed


// Method to handle tab selection changes and update UI accordingly
private void setupTabChangeListener() {
    tabbedPane.addChangeListener(e -> {
        int selectedIndex = tabbedPane.getSelectedIndex();
        if (selectedIndex >= 0) {
            Component selectedComponent = tabbedPane.getComponentAt(selectedIndex);
            String tabType = tabTypeMap.get(selectedComponent);
            
            // Update menu items or toolbar buttons based on tab type
            updateUIForTabType(tabType);
        }
    });
}

private void updateUIForTabType(String tabType) {
    // Enable/disable menu items based on tab type
    // For example:
    // if ("image".equals(tabType)) {
    //     enableImageSpecificActions();
    // } else if ("text".equals(tabType)) {
    //     enableTextSpecificActions();
    // }
}

// You can also use the utility methods separately:
private void someOtherMethod() {
    // Create a circular icon for user profile
    Icon profileIcon = ImageView.createCircularIcon("/icons/profile.png", 32);
    
    // Get image dimensions
    java.awt.Dimension dim = ImageView.getImageDimensions(new File("image.jpg"));
    System.out.println("Image size: " + dim.width + "x" + dim.height);
    
    // Check if file is supported
    if (ImageView.isSupportedImageFile(new File("photo.png"))) {
        // Process image
    }
}

// Remove the duplicate resizeImage method since it's already in ImageView class
// The ImageView.resizeImage() method will be used instead

// Additional helper method for middle-click tab closing
private void setupMiddleClickTabClosing() {
    tabbedPane.addMouseListener(new MouseAdapter() {
        @Override
        public void mousePressed(MouseEvent e) {
            if (e.getButton() == MouseEvent.BUTTON2) { // Middle click
                int tabIndex = tabbedPane.indexAtLocation(e.getX(), e.getY());
                if (tabIndex != -1) {
                    if (hasUnsavedChanges(tabIndex)) {
                        int result = JOptionPane.showConfirmDialog(tabbedPane, 
                            "You have unsaved changes. Close anyway?", 
                            "Unsaved Changes", 
                            JOptionPane.YES_NO_OPTION);
                        if (result != JOptionPane.YES_OPTION) {
                            return;
                        }
                    }
                    tabbedPane.remove(tabIndex);
                }
            }
        }
    });
}

// Method to add context menu to tabs
private void setupTabContextMenu() {
    JPopupMenu contextMenu = new JPopupMenu();
    JMenuItem closeTab = new JMenuItem("Close Tab");
    JMenuItem closeOtherTabs = new JMenuItem("Close Other Tabs");
    JMenuItem closeAllTabs = new JMenuItem("Close All Tabs");
    
    closeTab.addActionListener(e -> {
        int tabIndex = getTabIndexFromMouseEvent();
        if (tabIndex != -1) {
            tabbedPane.remove(tabIndex);
        }
    });
    
    closeOtherTabs.addActionListener(e -> {
        int currentTab = getTabIndexFromMouseEvent();
        if (currentTab != -1) {
            // Remove all tabs except the current one
            for (int i = tabbedPane.getTabCount() - 1; i >= 0; i--) {
                if (i != currentTab) {
                    tabbedPane.remove(i);
                }
            }
        }
    });
    
    closeAllTabs.addActionListener(e -> {
        // Remove all tabs
        tabbedPane.removeAll();
    });
    
    contextMenu.add(closeTab);
    contextMenu.add(closeOtherTabs);
    contextMenu.add(closeAllTabs);
    
    tabbedPane.setComponentPopupMenu(contextMenu);
}

private int getTabIndexFromMouseEvent() {
    Point mousePos = tabbedPane.getMousePosition();
    return mousePos != null ? tabbedPane.indexAtLocation(mousePos.x, mousePos.y) : -1;
}
private void addNewTableTab() {
    try {
        // Create a new blank spreadsheet panel using XlsxEditorPanel
        XlsxEditorPanel xlsxPanel = new XlsxEditorPanel(createBlankXlsxFile()); 

        // Define the tab title and icon
        String tabTitle = "Untitled Sheet";
        Icon sheetIcon = XlsxEditorPanel.getTabIcon();

        // Add the tab to the tabbedPane
        tabbedPane.addTab(tabTitle, sheetIcon, xlsxPanel);
        int index = tabbedPane.indexOfComponent(xlsxPanel);

        // Add a close button for the tab
        JPanel tabHeader = createTabHeader(tabTitle, sheetIcon, xlsxPanel);
        tabbedPane.setTabComponentAt(index, tabHeader);

        // Keep track of tab type
        tabTypeMap.put(xlsxPanel, "table");

        // Focus on the new tab
        tabbedPane.setSelectedIndex(index);

    } catch (Exception e) {
        JOptionPane.showMessageDialog(this,
                "Failed to create new spreadsheet tab: " + e.getMessage(),
                "Error", JOptionPane.ERROR_MESSAGE);
        e.printStackTrace();
    }
}

/**
 * Creates a temporary blank .xlsx file for new spreadsheet tabs
 */
private File createBlankXlsxFile() throws IOException {
    // Create a temporary file with .xlsx extension
    File tempFile = File.createTempFile("spreadsheet_", ".xlsx");
    tempFile.deleteOnExit(); // Clean up when application exits
    
    // Create a blank workbook and save it
    try (Workbook workbook = new XSSFWorkbook()) {
        workbook.createSheet("Sheet1");
        try (FileOutputStream fos = new FileOutputStream(tempFile)) {
            workbook.write(fos);
        }
    }
    
    return tempFile;
}

/**
 * Creates a tab header with title, icon, and close button
 */
private JPanel createTabHeader(String title, Icon icon, XlsxEditorPanel panel) {
    JPanel tabHeader = new JPanel(new FlowLayout(FlowLayout.LEFT, 5, 0));
    tabHeader.setOpaque(false);
    tabHeader.setBorder(BorderFactory.createEmptyBorder(2, 5, 2, 2));
    
    // Tab icon
    JLabel iconLabel = new JLabel(icon);
    
    // Tab title
    JLabel titleLabel = new JLabel(title);
    titleLabel.setBorder(BorderFactory.createEmptyBorder(0, 5, 0, 5));
    
    // Simple close button
    JButton closeButton = createSimpleCloseButton(panel);
    
    // Add components to tab header
    tabHeader.add(iconLabel);
    tabHeader.add(titleLabel);
    tabHeader.add(closeButton);
    
    return tabHeader;
}

/**
 * Creates a simple close button with 16x16 icon
 */
private JButton createSimpleCloseButton(XlsxEditorPanel panel) {
    JButton closeButton = new JButton();
    
    // Load and scale the 96x96 icon down to 16x16
    ImageIcon closeIcon = loadAndScaleCloseIcon("closetab.png", 16, 16);
    
    // Set the icon
    if (closeIcon != null) {
        closeButton.setIcon(closeIcon);
        System.out.println("✓ Successfully loaded and scaled closetab.png to 16x16");
    } else {
        // Fallback to simple X
        closeButton.setText("×");
        closeButton.setFont(new Font("Arial", Font.BOLD, 12));
        System.out.println("⚠ Using fallback close button - closetab.png not available");
    }
    
    // Simple styling
    closeButton.setPreferredSize(new Dimension(20, 20));
    closeButton.setFocusable(false);
    closeButton.setBorder(BorderFactory.createEmptyBorder(2, 2, 2, 2));
    closeButton.setContentAreaFilled(false);
    closeButton.setToolTipText("Close tab");
    
    // Simple hover effects
    closeButton.addMouseListener(new MouseAdapter() {
        @Override
        public void mouseEntered(MouseEvent e) {
            closeButton.setBackground(new Color(255, 100, 100, 80));
            closeButton.setContentAreaFilled(true);
        }
        
        @Override
        public void mouseExited(MouseEvent e) {
            closeButton.setBackground(null);
            closeButton.setContentAreaFilled(false);
        }
        
        @Override
        public void mousePressed(MouseEvent e) {
            closeButton.setBackground(new Color(255, 80, 80, 120));
        }
        
        @Override
        public void mouseReleased(MouseEvent e) {
            if (closeButton.contains(e.getPoint())) {
                closeButton.setBackground(new Color(255, 100, 100, 80));
            } else {
                closeButton.setBackground(null);
                closeButton.setContentAreaFilled(false);
            }
        }
    });
    
    // Close action
    closeButton.addActionListener(e -> closeTab(panel));
    
    return closeButton;
}

/**
 * Loads and scales the close icon from 96x96 to desired size
 */
private ImageIcon loadAndScaleCloseIcon(String filename, int targetWidth, int targetHeight) {
    try {
        // Try multiple resource paths
        String[] possiblePaths = {
            "/icons/" + filename,
            "/images/" + filename,
            "/resources/icons/" + filename,
            "/resources/images/" + filename,
            "icons/" + filename,
            "images/" + filename
        };
        
        for (String path : possiblePaths) {
            java.net.URL imageUrl = getClass().getResource(path);
            if (imageUrl != null) {
                System.out.println("✓ Found closetab.png in resources: " + path);
                ImageIcon originalIcon = new ImageIcon(imageUrl);
                // Scale from 96x96 to target size
                Image scaledImage = originalIcon.getImage().getScaledInstance(targetWidth, targetHeight, Image.SCALE_SMOOTH);
                return new ImageIcon(scaledImage);
            }
        }
        
        // Try file system paths
        String[] filePaths = {
            "icons/" + filename,
            "src/icons/" + filename,
            "src/main/resources/icons/" + filename,
            "src/main/resources/images/" + filename,
            "resources/icons/" + filename,
            "./icons/" + filename
        };
        
        for (String path : filePaths) {
            File iconFile = new File(path);
            if (iconFile.exists()) {
                System.out.println("✓ Found closetab.png in file system: " + path);
                ImageIcon originalIcon = new ImageIcon(iconFile.getAbsolutePath());
                // Scale from 96x96 to target size
                Image scaledImage = originalIcon.getImage().getScaledInstance(targetWidth, targetHeight, Image.SCALE_SMOOTH);
                return new ImageIcon(scaledImage);
            }
        }
        
        System.out.println("✗ closetab.png not found");
        return null;
        
    } catch (Exception e) {
        System.err.println("✗ Error loading closetab.png: " + e.getMessage());
        return null;
    }
}

/**
 * Closes the specified tab
 */
private void closeTab(XlsxEditorPanel panel) {
    int index = tabbedPane.indexOfComponent(panel);
    if (index != -1) {
        // Check if there are unsaved changes
        if (hasUnsavedChanges(panel)) {
            int result = JOptionPane.showConfirmDialog(this,
                    "There are unsaved changes. Do you want to save before closing?",
                    "Unsaved Changes",
                    JOptionPane.YES_NO_CANCEL_OPTION,
                    JOptionPane.WARNING_MESSAGE);
            
            if (result == JOptionPane.YES_OPTION) {
                saveTab(panel);
            } else if (result == JOptionPane.CANCEL_OPTION) {
                return; // Don't close the tab
            }
        }
        
        // Remove from tab type map
        tabTypeMap.remove(panel);
        
        // Remove the tab
        tabbedPane.remove(index);
        
        // Clean up temporary file if it exists
        cleanupTempFile(panel);
    }
}

/**
 * Checks if the tab has unsaved changes
 */
private boolean hasUnsavedChanges(XlsxEditorPanel panel) {
    // You can implement logic to check for unsaved changes
    // For now, return false - you can enhance this later
    return false;
}

/**
 * Saves the tab content
 */
private void saveTab(XlsxEditorPanel panel) {
    try {
        // Use the existing save functionality from XlsxEditorPanel
        panel.handleSave(new ActionEvent(panel, ActionEvent.ACTION_PERFORMED, "save"));
    } catch (Exception e) {
        JOptionPane.showMessageDialog(this,
                "Failed to save file: " + e.getMessage(),
                "Save Error", JOptionPane.ERROR_MESSAGE);
    }
}

/**
 * Cleans up temporary files when tab is closed
 */
private void cleanupTempFile(XlsxEditorPanel panel) {
    // You can implement cleanup logic here if needed
    // The temporary file is set to deleteOnExit, but you might want immediate cleanup
}


    private void showCustomPaletteDialog() {
        JPanel panel = new JPanel(new FlowLayout());
        JButton[] colorButtons = new JButton[customColors.size()];
        for (int i = 0; i < customColors.size(); i++) {
            Color c = customColors.get(i);
            JButton btn = new JButton();
            btn.setBackground(c);
            btn.setPreferredSize(new Dimension(32, 32));
            btn.addActionListener(e -> applyTextColor(c));
            colorButtons[i] = btn;
            panel.add(btn);
        }
        JButton addBtn = new JButton("Add Color ➕");
        addBtn.addActionListener(e -> {
            Color newColor = JColorChooser.showDialog(this, "Pick a Color 🎨", Color.BLACK);
            if (newColor != null) {
                customColors.add(newColor);
                showCustomPaletteDialog();
            }
        });
        panel.add(addBtn);
        JOptionPane.showMessageDialog(this, panel, "Custom Color Palette 🎨", JOptionPane.PLAIN_MESSAGE);
    }

    private void applyTextColor(Color color) {
        JTextPane pane = getCurrentTextPane();
        if (pane != null) {
            MutableAttributeSet attr = new SimpleAttributeSet();
            StyleConstants.setForeground(attr, color);
            pane.setCharacterAttributes(attr, false);
        }
    }

    /**
     * Creates a new, blank spreadsheet-like tab with a default grid.
     * This tab can then be saved as an .xlsx file using the save functionality.
     */
    /**
     * Creates a new, blank spreadsheet tab with a powerful Excel-like formula
     * engine.
     * Supports cell references (e.g., =A1+B2), automatic recalculation, and rich
     * text.
     */


private boolean tabTitleExists(String title) {
    for (int i = 0; i < tabbedPane.getTabCount(); i++) {
        if (tabbedPane.getTitleAt(i).equals(title)) {
            return true;
        }
    }
    return false;
}
private JPanel createTabHeader(String title, Icon icon, Component comp) {
    JPanel panel = new JPanel(new FlowLayout(FlowLayout.LEFT, 0, 0));
    panel.setOpaque(false);

    JLabel titleLabel = new JLabel(title, icon, JLabel.LEFT);
    titleLabel.setBorder(BorderFactory.createEmptyBorder(0, 0, 0, 5));

    JButton closeButton = new JButton("×");
    closeButton.setMargin(new Insets(0, 4, 0, 4));
    closeButton.setBorderPainted(false);
    closeButton.setFocusPainted(false);
    closeButton.setContentAreaFilled(false);
    closeButton.setForeground(Color.RED);
    closeButton.setFont(closeButton.getFont().deriveFont(Font.BOLD, 14f));

    closeButton.addActionListener(e -> {
        int index = tabbedPane.indexOfComponent(comp);
        if (index >= 0) tabbedPane.remove(index);
    });

    panel.add(titleLabel);
    panel.add(closeButton);

    return panel;
}

    /**
     * Recalculates all formulas in the sheet. Iterates through the formula map
     * and updates the table model with the new results.
     */
    private void recalculateAllFormulas(DefaultTableModel model, Map<Point, String> formulaMap,
            boolean[] isRecalculating) {
        isRecalculating[0] = true; // Set flag to prevent listener loops

        for (Map.Entry<Point, String> entry : formulaMap.entrySet()) {
            Point cell = entry.getKey();
            String formula = entry.getValue();
            try {
                Object result = evaluateFormula(formula, model);
                // Update the model without triggering the listener again.
                model.setValueAt(result, cell.x, cell.y);
            } catch (Exception ex) {
                model.setValueAt("#ERROR", cell.x, cell.y);
            }
        }

        isRecalculating[0] = false; // Unset the flag
    }

    /**
     * Evaluates a single formula string (e.g., "=A1+B2").
     * It replaces cell references with their numeric values and uses a ScriptEngine
     * to compute the result.
     */
    private Object evaluateFormula(String formula, DefaultTableModel model) throws Exception {
        // Prepare the expression by removing the '=' and finding cell references.
        String expression = formula.substring(1);

        // Regex to find cell references like A1, B12, C4, etc.
        Pattern cellPattern = Pattern.compile("([A-Z]+)(\\d+)");
        Matcher matcher = cellPattern.matcher(expression);

        StringBuffer sb = new StringBuffer();
        while (matcher.find()) {
            String colLetters = matcher.group(1);
            int row = Integer.parseInt(matcher.group(2)) - 1; // 1-based to 0-based index
            int col = 0;
            for (char c : colLetters.toCharArray()) {
                col = col * 26 + (c - 'A' + 1);
            }
            col--; // 1-based to 0-based index

            if (row < 0 || col < 0 || row >= model.getRowCount() || col >= model.getColumnCount()) {
                throw new IllegalArgumentException("Invalid cell reference");
            }

            Object cellValue = model.getValueAt(row, col);
            // Replace the cell reference with its actual value.
            matcher.appendReplacement(sb, cellValue != null ? cellValue.toString() : "0");
        }
        matcher.appendTail(sb);

        // Use the JavaScript engine to evaluate the final mathematical expression.
        ScriptEngineManager manager = new ScriptEngineManager();
        ScriptEngine engine = manager.getEngineByName("JavaScript");
        return engine.eval(sb.toString());
    }

   private void addNewTab(File file, String content) {
    JTextPane pane = new JTextPane();
    pane.setText(content);
    JScrollPane scroll = new JScrollPane(pane);

    // Add smooth scrolling to the new tab
    addSmoothScrolling(scroll);

    UndoManager um = new UndoManager();
    pane.getDocument().addUndoableEditListener(um);
    undoMap.put(pane, um);

    String title = file == null ? "Untitled" : file.getName();
    ImageIcon originalIcon = new ImageIcon(getClass().getResource("/icons/newtab.png"));
    int fixedWidth = 16;
    int fixedHeight = 16;
    Image scaledImage = originalIcon.getImage().getScaledInstance(fixedWidth, fixedHeight, Image.SCALE_SMOOTH);
    Icon logoIcon = new ImageIcon(scaledImage);

    // ✅ Add tab with close button and icon
    addTabWithCloseButton(title, scroll, logoIcon);

    fileMap.put(pane, file);
    tabTypeMap.put(scroll, "text");

    pane.getDocument().addDocumentListener(new DocumentListener() {
        public void insertUpdate(DocumentEvent e) {
            updateTabTitle(pane);
            updateStatusBar(pane);
        }

        public void removeUpdate(DocumentEvent e) {
            updateTabTitle(pane);
            updateStatusBar(pane);
        }

        public void changedUpdate(DocumentEvent e) {
            updateTabTitle(pane);
            updateStatusBar(pane);
        }
    });

    pane.addCaretListener(e -> updateStatusBar(pane));
}

    private JTextPane getCurrentTextPane() {
        Component comp = tabbedPane.getSelectedComponent();
        if (comp instanceof JScrollPane) {
            JScrollPane scroll = (JScrollPane) comp;
            JViewport vp = scroll.getViewport();
            if (vp.getView() instanceof JTextPane)
                return (JTextPane) vp.getView();
        }
        return null;
    }

    private JTable getCurrentTable() {
        Component comp = tabbedPane.getSelectedComponent();
        if (comp instanceof JScrollPane) {
            JScrollPane scroll = (JScrollPane) comp;
            JViewport vp = scroll.getViewport();
            if (vp.getView() instanceof JTable)
                return (JTable) vp.getView();
        }
        return null;
    }

    private void updateTabTitle(JTextPane pane) {
        JScrollPane scroll = null;
        for (int i = 0; i < tabbedPane.getTabCount(); i++) {
            Component comp = tabbedPane.getComponentAt(i);
            if (comp instanceof JScrollPane) {
                JScrollPane s = (JScrollPane) comp;
                if (s.getViewport().getView() == pane) {
                    scroll = s;
                    break;
                }
            }
        }
        if (scroll != null) {
            int idx = tabbedPane.indexOfComponent(scroll);
            File file = fileMap.get(pane);
            String title = (file == null ? "Untitled" : file.getName());
            tabbedPane.setTitleAt(idx, title + (pane.getDocument().getLength() > 0 ? "*" : ""));
        }
    }

    private void updateStatusBar(JTextPane pane) {
        try {
            int caret = pane.getCaretPosition();
            int line = pane.getDocument().getDefaultRootElement().getElementIndex(caret);
            int col = caret - pane.getDocument().getDefaultRootElement().getElement(line).getStartOffset();
            statusBar.setText("Ln " + (line + 1) + ", Col " + (col + 1));
        } catch (Exception e) {
            statusBar.setText("");
        }
    }

    private void saveTextTab(Component comp, boolean forceDialog) {
        if (!(comp instanceof JScrollPane scroll)) {
            JOptionPane.showMessageDialog(this, "Unsupported component format.");
            return;
        }

        Component view = scroll.getViewport().getView();
        String text;
        File file;

        if (view instanceof JTextPane textPane) {
            text = textPane.getText();
            file = fileMap.get(textPane);
        } else if (view instanceof JTextArea textArea) {
            text = textArea.getText();
            file = null; // `JTextArea` isn't mapped in `fileMap`
        } else {
            JOptionPane.showMessageDialog(this, "Unknown editor type.");
            return;
        }

        JFileChooser fc = new JFileChooser();
        if (forceDialog || file == null) {
            fc.setDialogTitle("Save As");
            fc.setSelectedFile(new File("Untitled.txt"));
            fc.setFileFilter(new FileNameExtensionFilter("Text/Word/PDF Files", "txt", "docx", "pdf"));

            if (fc.showSaveDialog(this) != JFileChooser.APPROVE_OPTION) {
                return;
            }
            file = fc.getSelectedFile();

            // FIX: Store file reference ONLY for `JTextPane`
            if (view instanceof JTextPane textPane) {
                fileMap.put(textPane, file);
            }
        }

        try {
            if (file.getName().toLowerCase().endsWith(".txt")) {
                try (BufferedWriter writer = new BufferedWriter(new FileWriter(file))) {
                    writer.write(text);
                }
            } else if (file.getName().toLowerCase().endsWith(".docx")) {
                XWPFDocument doc = new XWPFDocument();
                XWPFParagraph p = doc.createParagraph();
                XWPFRun run = p.createRun();
                run.setText(text);
                try (FileOutputStream out = new FileOutputStream(file)) {
                    doc.write(out);
                }
            } else {
                JOptionPane.showMessageDialog(this, "Unsupported file format.");
                return;
            }

            // FIX: Update tab title ONLY for `JTextPane`, avoid `JTextArea`
            if (view instanceof JTextPane textPane) {
                updateTabTitle(textPane);
            }

        } catch (Exception ex) {
            JOptionPane.showMessageDialog(this, "Cannot save file:\n" + ex.getMessage());
        }
    }

    private void saveTableTab(Component comp, boolean forceDialog) {
        JScrollPane scroll = (JScrollPane) comp;
        JTable table = (JTable) scroll.getViewport().getView();
        JFileChooser fc = new JFileChooser();
        fc.setDialogTitle("Save Table As 📄`");
        fc.setSelectedFile(new File("Table.xlsx"));
        fc.setFileFilter(new FileNameExtensionFilter("Excel Files", "xlsx"));
        if (fc.showSaveDialog(this) != JFileChooser.APPROVE_OPTION)
            return;
        File file = fc.getSelectedFile();
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            XSSFSheet sheet = workbook.createSheet("Sheet1");
            XSSFRow header = sheet.createRow(0);
            for (int col = 0; col < table.getColumnCount(); col++) {
                header.createCell(col).setCellValue(table.getColumnName(col));
            }
            for (int row = 0; row < table.getRowCount(); row++) {
                XSSFRow excelRow = sheet.createRow(row + 1);
                for (int col = 0; col < table.getColumnCount(); col++) {
                    Object val = table.getValueAt(row, col);
                    if (val instanceof StyledDocument) {
                        try {
                            StyledDocument doc = (StyledDocument) val;
                            excelRow.createCell(col).setCellValue(doc.getText(0, doc.getLength()));
                        } catch (Exception ex) {
                            excelRow.createCell(col).setCellValue("");
                        }
                    } else {
                        excelRow.createCell(col).setCellValue(val == null ? "" : val.toString());
                    }
                }
            }
            try (FileOutputStream out = new FileOutputStream(file)) {
                workbook.write(out);
            }
        } catch (Exception ex) {
            JOptionPane.showMessageDialog(this, "Cannot save table ❌:\n" + ex.getMessage());
        }
    }

    private void closeCurrentTab() {
        int index = tabbedPane.getSelectedIndex();
        if (index == -1)
            return;

        Component comp = tabbedPane.getComponentAt(index);
        String type = tabTypeMap.getOrDefault(comp, "text");

        if ("text".equals(type)) {
            JTextPane pane = null;

            if (comp instanceof JScrollPane) {
                JScrollPane scroll = (JScrollPane) comp;
                Component view = scroll.getViewport().getView();
                if (view instanceof JTextPane) {
                    pane = (JTextPane) view;
                }
            } else if (comp instanceof JPanel) {
                // In case it's wrapped in a panel (not common for text tabs, but safe check)
                for (Component inner : ((JPanel) comp).getComponents()) {
                    if (inner instanceof JScrollPane) {
                        JScrollPane scroll = (JScrollPane) inner;
                        Component view = scroll.getViewport().getView();
                        if (view instanceof JTextPane) {
                            pane = (JTextPane) view;
                            break;
                        }
                    }
                }
            }

            if (pane != null && pane.getDocument().getLength() > 0) {
                tabbedPane.setSelectedIndex(index);
                if (!confirmSave())
                    return; // User canceled close
            }
        }

        // Remove tab and clean up map
        tabTypeMap.remove(comp);
        tabbedPane.removeTabAt(index);
    }

    private boolean confirmSave() {
        JTextPane pane = getCurrentTextPane();
        if (pane == null)
            return true;
        File file = fileMap.get(pane);
        String title = (file == null ? "Untitled" : file.getName());
        if (pane.getDocument().getLength() == 0)
            return true;
        int res = JOptionPane.showConfirmDialog(this, "Save changes to " + title + "?", "Save",
                JOptionPane.YES_NO_CANCEL_OPTION);
        if (res == JOptionPane.CANCEL_OPTION)
            return false;
        if (res == JOptionPane.YES_OPTION)
            saveFile();
        return true;
    }

    private boolean confirmSaveAll() {
        for (int i = 0; i < tabbedPane.getTabCount(); i++) {
            Component comp = tabbedPane.getComponentAt(i);
            String type = tabTypeMap.getOrDefault(comp, "text");

            if ("text".equals(type)) {
                JTextPane pane = null;

                if (comp instanceof JScrollPane) {
                    JScrollPane scroll = (JScrollPane) comp;
                    Component view = scroll.getViewport().getView();
                    if (view instanceof JTextPane) {
                        pane = (JTextPane) view;
                    }
                } else if (comp instanceof JPanel) {
                    // Some components might be wrapped inside a JPanel
                    for (Component inner : ((JPanel) comp).getComponents()) {
                        if (inner instanceof JScrollPane) {
                            JScrollPane scroll = (JScrollPane) inner;
                            Component view = scroll.getViewport().getView();
                            if (view instanceof JTextPane) {
                                pane = (JTextPane) view;
                                break;
                            }
                        }
                    }
                }

                if (pane != null && pane.getDocument().getLength() > 0) {
                    tabbedPane.setSelectedIndex(i);
                    if (!confirmSave()) {
                        return false;
                    }
                }
            }
        }
        return true;
    }

    private void sumSelectedTableColumn() {
        JTable table = getCurrentTable();
        if (table == null) {
            JOptionPane.showMessageDialog(this, "No table tab selected ‼️.");
            return;
        }
        int col = table.getSelectedColumn();
        if (col < 0) {
            JOptionPane.showMessageDialog(this, "Select a column first ❗.");
            return;
        }
        double sum = 0;
        for (int row = 0; row < table.getRowCount(); row++) {
            Object val = table.getValueAt(row, col);
            String txt = "";
            if (val instanceof StyledDocument) {
                try {
                    txt = ((StyledDocument) val).getText(0, ((StyledDocument) val).getLength());
                } catch (Exception ignored) {
                }
            } else if (val != null) {
                txt = val.toString();
            }
            try {
                sum += Double.parseDouble(txt.trim());
            } catch (Exception ignored) {
            }
        }
        JOptionPane.showMessageDialog(this, "Sum: " + sum);
    }

    private void avgSelectedTableColumn() {
        JTable table = getCurrentTable();
        if (table == null) {
            JOptionPane.showMessageDialog(this, "No table tab selected ‼️.");
            return;
        }
        int col = table.getSelectedColumn();
        if (col < 0) {
            JOptionPane.showMessageDialog(this, "Select a column first ❗.");
            return;
        }
        double sum = 0;
        int count = 0;
        for (int row = 0; row < table.getRowCount(); row++) {
            Object val = table.getValueAt(row, col);
            String txt = "";
            if (val instanceof StyledDocument) {
                try {
                    txt = ((StyledDocument) val).getText(0, ((StyledDocument) val).getLength());
                } catch (Exception ignored) {
                }
            } else if (val != null) {
                txt = val.toString();
            }
            try {
                sum += Double.parseDouble(txt.trim());
                count++;
            } catch (Exception ignored) {
            }
        }
        JOptionPane.showMessageDialog(this, "Average: " + (count > 0 ? sum / count : 0));
    }

    private void transformSelectedText(java.util.function.Function<String, String> func) {
        JTextPane pane = getCurrentTextPane();
        if (pane != null) {
            String sel = pane.getSelectedText();
            if (sel != null) {
                pane.replaceSelection(func.apply(sel));
            }
        }
    }

    private static String toTitleCase(String input) {
        String[] words = input.split("\\s");
        StringBuilder sb = new StringBuilder();
        for (String w : words) {
            if (w.length() > 0)
                sb.append(Character.toUpperCase(w.charAt(0))).append(w.substring(1).toLowerCase()).append(" ");
        }
        return sb.toString().trim();
    }

    private void removeDuplicateLines() {
        JTextPane pane = getCurrentTextPane();
        if (pane != null) {
            String[] lines = pane.getText().split("\n");
            LinkedHashSet<String> set = new LinkedHashSet<>(Arrays.asList(lines));
            pane.setText(String.join("\n", set));
        }
    }

    private void sortLines() {
        JTextPane pane = getCurrentTextPane();
        if (pane != null) {
            String[] lines = pane.getText().split("\n");
            Arrays.sort(lines);
            pane.setText(String.join("\n", lines));
        }
    }

    private void duplicateLines() {
        JTextPane pane = getCurrentTextPane();
        if (pane == null)
            return;
        try {
            int start = pane.getDocument().getDefaultRootElement().getElementIndex(pane.getSelectionStart());
            int end = pane.getDocument().getDefaultRootElement().getElementIndex(pane.getSelectionEnd());
            Document doc = pane.getDocument();
            for (int i = start; i <= end; i++) {
                Element elem = doc.getDefaultRootElement().getElement(i);
                String text = doc.getText(elem.getStartOffset(), elem.getEndOffset() - elem.getStartOffset());
                doc.insertString(elem.getEndOffset(), text, null);
            }
        } catch (Exception ignored) {
        }
    }

    private void removeBlankLines() {
        JTextPane pane = getCurrentTextPane();
        if (pane == null)
            return;
        String[] lines = pane.getText().split("\n");
        StringBuilder sb = new StringBuilder();
        for (String line : lines) {
            if (!line.trim().isEmpty())
                sb.append(line).append("\n");
        }
        pane.setText(sb.toString());
    }

    private void trimAllLines() {
        JTextPane pane = getCurrentTextPane();
        if (pane == null)
            return;
        String[] lines = pane.getText().split("\n");
        for (int i = 0; i < lines.length; i++)
            lines[i] = lines[i].trim();
        pane.setText(String.join("\n", lines));
    }

    private void joinAllLines() {
        JTextPane pane = getCurrentTextPane();
        if (pane == null)
            return;
        String text = pane.getText().replace("\n", " ");
        pane.setText(text);
    }

    private void convertTabsToSpaces() {
        JTextPane pane = getCurrentTextPane();
        if (pane == null)
            return;
        String text = pane.getText().replace("\t", "    ");
        pane.setText(text);
    }

    static class WrapEditorKit extends StyledEditorKit {
        public ViewFactory getViewFactory() {
            return new StyledViewFactory();
        }

        static class StyledViewFactory implements ViewFactory {
            public View create(Element elem) {
                String kind = elem.getName();
                if (kind != null) {
                    if (kind.equals(AbstractDocument.ContentElementName)) {
                        return new WrapLabelView(elem);
                    } else if (kind.equals(AbstractDocument.ParagraphElementName)) {
                        return new ParagraphView(elem);
                    } else if (kind.equals(AbstractDocument.SectionElementName)) {
                        return new BoxView(elem, View.Y_AXIS);
                    } else if (kind.equals(StyleConstants.ComponentElementName)) {
                        return new ComponentView(elem);
                    } else if (kind.equals(StyleConstants.IconElementName)) {
                        return new IconView(elem);
                    }
                }
                return new LabelView(elem);
            }
        }

        static class WrapLabelView extends LabelView {
            public WrapLabelView(Element elem) {
                super(elem);
            }

            public float getMinimumSpan(int axis) {
                return super.getPreferredSpan(axis);
            }
        }
    }

    private void formatJSON() {
        JTextPane pane = getCurrentTextPane();
        if (pane != null) {
            try {
                Gson gson = new GsonBuilder().setPrettyPrinting().create();
                Object jsonObject = gson.fromJson(pane.getText(), Object.class);
                String formatted = gson.toJson(jsonObject);
                pane.setText(formatted);
            } catch (Exception e) {
                JOptionPane.showMessageDialog(this, "Invalid JSON: " + e.getMessage());
            }
        }
    }

    private void minifyText() {
        JTextPane pane = getCurrentTextPane();
        if (pane != null) {
            String text = pane.getText();
            // Basic minification - remove extra whitespace
            text = text.replaceAll("\\s+", " ").trim();
            pane.setText(text);
        }
    }

    private void showBaseConverter() {
        String input = JOptionPane.showInputDialog(this, "Enter a decimal number:");
        try {
            int number = Integer.parseInt(input);
            String msg = "Binary: " + Integer.toBinaryString(number) +
                    "\nHex: " + Integer.toHexString(number) +
                    "\nOctal: " + Integer.toOctalString(number);
            JOptionPane.showMessageDialog(this, msg, "Base Converter", JOptionPane.INFORMATION_MESSAGE);
        } catch (NumberFormatException e) {
            showError("Invalid number");
        }
    }

    private void showRegexTester() {
        JPanel panel = new JPanel(new BorderLayout());
        JTextField patternField = new JTextField();
        JTextArea resultArea = new JTextArea(5, 30);
        resultArea.setEditable(false);

        patternField.addActionListener(e -> {
            try {
                Pattern pattern = Pattern.compile(patternField.getText());
                Matcher matcher = pattern.matcher(textArea.getText());
                StringBuilder sb = new StringBuilder();
                while (matcher.find()) {
                    sb.append("Match: ").append(matcher.group()).append("\n");
                }
                resultArea.setText(sb.toString());
            } catch (Exception ex) {
                resultArea.setText("Invalid regex pattern");
            }
        });

        panel.add(new JLabel("Regex Pattern:"), BorderLayout.NORTH);
        panel.add(patternField, BorderLayout.CENTER);
        panel.add(new JScrollPane(resultArea), BorderLayout.SOUTH);

        JOptionPane.showMessageDialog(this, panel, "Regex Tester", JOptionPane.PLAIN_MESSAGE);
    }

    private void generateQRCode() {
        String text = JOptionPane.showInputDialog(this, "Enter text to encode:");
        if (text == null || text.isEmpty())
            return;

        try {
            BitMatrix matrix = new MultiFormatWriter().encode(text, BarcodeFormat.QR_CODE, 200, 200);
            BufferedImage image = MatrixToImageWriter.toBufferedImage(matrix);
            JOptionPane.showMessageDialog(this, new JLabel(new ImageIcon(image)), "QR Code", JOptionPane.PLAIN_MESSAGE);
        } catch (Exception e) {
            showError("Error generating QR Code");
        }
    }

    private void exportTextAsQRCode() {
        try {
            String content = textArea.getText();
            if (content == null || content.trim().isEmpty()) {
                showError("Cannot export QR Code: No text found in the editor.");
                return;
            }
            // Ensure the "qr" directory exists
            File qrDir = new File("qr");
            if (!qrDir.exists()) {
                qrDir.mkdirs();
            }
            // Save with a timestamp to avoid overwriting
            String fileName = "qrcode_" + System.currentTimeMillis() + ".png";
            File file = new File(qrDir, fileName);

            BitMatrix matrix = new MultiFormatWriter().encode(content, BarcodeFormat.QR_CODE, 300, 300);
            MatrixToImageWriter.writeToPath(matrix, "PNG", file.toPath());
            JOptionPane.showMessageDialog(this, "QR Code exported: " + file.getAbsolutePath());
        } catch (Exception e) {
            showError("Failed to export QR Code: " + e.getMessage());
            e.printStackTrace();
        }
    }

    private void checkFileHash() {
        JFileChooser chooser = new JFileChooser();
        if (chooser.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) {
            File file = chooser.getSelectedFile();
            try {
                byte[] bytes = Files.readAllBytes(file.toPath());
                MessageDigest md = MessageDigest.getInstance("SHA-256");
                byte[] hash = md.digest(bytes);
                StringBuilder sb = new StringBuilder();
                for (byte b : hash)
                    sb.append(String.format("%02x", b));
                JOptionPane.showMessageDialog(this, "SHA-256: " + sb);
            } catch (IOException | NoSuchAlgorithmException e) {
                showError("Error reading file or computing hash");
            }
        }
    }

    private void runJavaSnippet() {
        JOptionPane.showMessageDialog(this, "Feature not yet implemented.", "Coming Soon",
                JOptionPane.INFORMATION_MESSAGE);
    }

    private void showError(String msg) {
        JOptionPane.showMessageDialog(this, msg, "Error", JOptionPane.ERROR_MESSAGE);
    }

    

    private void updateUI() {
        UIManager.put("Component.arc", 20);
        UIManager.put("Button.arc", 20);
        UIManager.put("TextComponent.arc", 20);
        UIManager.put("ScrollBar.thumbArc", 16);

        SwingUtilities.updateComponentTreeUI(this);
        this.repaint();
    }

    public static void applyTheme(String theme, TextEditor editor) throws UnsupportedLookAndFeelException {
        switch (theme.toLowerCase()) {
            case "dark" -> {
                FlatMacDarkLaf.setup();
                UIManager.setLookAndFeel(new FlatMacDarkLaf());
            }
            default -> {
                FlatMacLightLaf.setup();
                UIManager.setLookAndFeel(new FlatMacLightLaf());
            }
        }

        // Style rounding
        UIManager.put("Component.arc", 20);
        UIManager.put("Button.arc", 20);
        UIManager.put("TextComponent.arc", 20);
        UIManager.put("ScrollBar.thumbArc", 16);

        // Update UI
        SwingUtilities.updateComponentTreeUI(editor);
        editor.repaint();
    }

    public void fadeInContent() {
        Timer timer = new Timer(15, null);
        timer.addActionListener(e -> {
            float newAlpha = rootPanel.alpha + 0.05f;
            if (newAlpha >= 1.0f) {
                rootPanel.setAlpha(1.0f);
                timer.stop();
            } else {
                rootPanel.setAlpha(newAlpha);
            }
        });
        timer.start();
    }

   public static void main(String[] args) {
    new JFXPanel(); // Keep for JavaFX components

    // ========== DISABLE ALL ANIMATIONS ==========

    SwingUtilities.invokeLater(() -> {
        TextEditor editor = new TextEditor();
        try {
            applyTheme("light", editor);
        } catch (UnsupportedLookAndFeelException e) {
            e.printStackTrace();
        }

        editor.setVisible(true); // Make the frame visible normally
        // Remove or comment out the fade-in call if you want no animations at all
        // editor.fadeInContent(); 
    });
}

    private void applyTheme(String theme, Component parent) throws UnsupportedLookAndFeelException {
        try {
            switch (theme.toLowerCase()) {
                case "dark":
                    UIManager.setLookAndFeel(new FlatMacDarkLaf());
                    break;
                case "light":
                default:
                    UIManager.setLookAndFeel(new FlatMacLightLaf());
                    break;
            }
            SwingUtilities.updateComponentTreeUI(parent);
        } catch (UnsupportedLookAndFeelException e) {
            throw e;
        }
    }

    /**
     * Apply theme dynamically without restarting the editor.
     */

}

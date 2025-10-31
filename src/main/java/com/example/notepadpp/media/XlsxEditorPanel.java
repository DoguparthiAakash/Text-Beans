package com.example.notepadpp.media;
import java.awt.BasicStroke;
import java.awt.BorderLayout;
import java.awt.Color;
import java.awt.Component;
import java.awt.Container;
import java.awt.Cursor;
import java.awt.Dimension;
import java.awt.Font;
import java.awt.Image;
import java.awt.Point;
import java.awt.Rectangle;
import java.awt.datatransfer.DataFlavor;
import java.awt.datatransfer.StringSelection;
import java.awt.datatransfer.Transferable;
import java.awt.event.ActionEvent;
import java.awt.event.KeyAdapter;
import java.awt.event.KeyEvent;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.awt.event.MouseMotionAdapter;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import javax.swing.BorderFactory;
import javax.swing.BoxLayout;
import javax.swing.Icon;
import javax.swing.ImageIcon;
import javax.swing.JButton;
import javax.swing.JCheckBox;
import javax.swing.JColorChooser;
import javax.swing.JComboBox;
import javax.swing.JComponent;
import javax.swing.JFileChooser;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JSpinner;
import javax.swing.JTable;
import javax.swing.JTextField;
import javax.swing.JToolBar;
import javax.swing.RowFilter;
import javax.swing.RowSorter;
import javax.swing.SortOrder;
import javax.swing.SpinnerNumberModel;
import javax.swing.SwingConstants;
import javax.swing.Timer;
import javax.swing.TransferHandler;
import javax.swing.border.Border;
import javax.swing.event.DocumentEvent;
import javax.swing.event.DocumentListener;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.table.AbstractTableModel;
import javax.swing.table.TableCellRenderer;
import javax.swing.table.TableRowSorter;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * ENHANCED XlsxEditorPanel with Excel-like features including drag selection, format painter,
 * select-all functionality, and comprehensive formatting options
 */
public class XlsxEditorPanel extends JPanel {
    private final File file;
    private final JTable table;
    private final OptimizedTableModel model;
    private TableRowSorter<OptimizedTableModel> sorter;
    
    // Enhanced feature flags
    private boolean showTotalRow = false;
    private boolean showFormulaBar = true;
    private boolean autoExpandEnabled = true;
    private boolean alternateRowColors = true;
    private boolean tableMode = true;
    private boolean freezePanes = false;
    private boolean showGridLines = true;
    private boolean autoSaveEnabled = false;
    private boolean cellTrackingEnabled = true;
    private boolean isDragging = false;
    private boolean isFormatPainterActive = false;
    private boolean isSelectAllMode = false;
    
    // Drag selection
    private Point dragStartPoint = null;
    private Point dragEndPoint = null;
    private Rectangle dragRectangle = null;
    
    // Format painter
    private CellFormat sourceFormat = null;
    
    // UI components
    private final JPanel totalRowPanel = new JPanel();
    private final JTextField formulaBar = new JTextField();
    private final JLabel statusBar = new JLabel("Ready");
    private final JComboBox<String> numberFormatCombo = new JComboBox<>(
        new String[]{"General", "Number", "Currency", "Date", "Percentage", "Text", "Scientific", "Fraction", "Accounting", "Custom"}
    );
    private final JComboBox<String> fontCombo = new JComboBox<>(new String[]{"Arial", "Calibri", "Times New Roman", "Courier New", "Verdana", "Tahoma", "Georgia"});
    private final JSpinner fontSizeSpinner = new JSpinner(new SpinnerNumberModel(12, 8, 72, 1));
    private final JCheckBox boldCheckBox = new JCheckBox("B");
    private final JCheckBox italicCheckBox = new JCheckBox("I");
    private final JCheckBox underlineCheckBox = new JCheckBox("U");
    private final JComboBox<String> horizontalAlignmentCombo = new JComboBox<>(
        new String[]{"Left", "Center", "Right", "Justify"}
    );
    private final JComboBox<String> verticalAlignmentCombo = new JComboBox<>(
        new String[]{"Top", "Center", "Bottom"}
    );
    private final JComboBox<String> borderStyleCombo = new JComboBox<>(
        new String[]{"None", "Thin", "Medium", "Thick", "Dashed", "Dotted"}
    );
    
    // Formula engine
    private final EnhancedFormulaEngine formulaEngine = new EnhancedFormulaEngine();
    
    // Performance optimization
    private final Map<String, Object> calculationCache = new HashMap<>();
    private final Set<String> dirtyCells = new HashSet<>();
    private Timer delayedCalculationTimer;
    private Timer autoSaveTimer;
    
    // Cell tracking
    private final List<String> cellHistory = new ArrayList<>();
    private int historyPointer = -1;
    
    // Border management
    private final Map<String, BorderInfo> borderMap = new HashMap<>();
    
    // Lazy formatter initialization
    private DecimalFormat numberFormat;
    private DecimalFormat currencyFormat;
    private DecimalFormat percentFormat;
    private DecimalFormat scientificFormat;
    private DecimalFormat accountingFormat;
    private SimpleDateFormat dateFormat;

    public XlsxEditorPanel(File file) {
        this.file = file;
        setLayout(new BorderLayout());
        
        model = new OptimizedTableModel(this);
        table = new JTable(model) {
            @Override
            public Component prepareRenderer(TableCellRenderer renderer, int row, int col) {
                Component comp = super.prepareRenderer(renderer, row, col);
                return model.prepareTableCellRenderer(comp, row, col, alternateRowColors, tableMode);
            }
            
            @Override
            public void changeSelection(int rowIndex, int columnIndex, boolean toggle, boolean extend) {
                super.changeSelection(rowIndex, columnIndex, toggle, extend);
                updateFormulaBar();
                updateStatusBar();
                updateCellHistory(rowIndex, columnIndex);
                updateFormattingToolbar();
            }
        };

        setupTable();
        setupEnhancedUI();
        setupAutoSave();
        setupDragAndDrop();
        setupFormatPainter();
        setupKeyboardShortcuts();
        
        // Add theme change listener
        addPropertyChangeListener("background", evt -> {
            boolean isDark = ((Color) evt.getNewValue()).getRed() < 128;
            updateThemeColors(isDark);
        });
        
        // Only load content if file is not null
        if (file != null) {
            loadXlsxContent();
        }
        
        autoResizeColumns();
    }

    // ==================== KEYBOARD SHORTCUTS ====================

    private void setupKeyboardShortcuts() {
        table.addKeyListener(new KeyAdapter() {
            @Override
            public void keyPressed(KeyEvent e) {
                // Select All (Ctrl+A)
                if (e.isControlDown() && e.getKeyCode() == KeyEvent.VK_A) {
                    selectAllCells();
                    e.consume();
                }
                // Copy (Ctrl+C)
                else if (e.isControlDown() && e.getKeyCode() == KeyEvent.VK_C) {
                    copySelection();
                    e.consume();
                }
                // Paste (Ctrl+V)
                else if (e.isControlDown() && e.getKeyCode() == KeyEvent.VK_V) {
                    pasteToSelection();
                    e.consume();
                }
                // Cut (Ctrl+X)
                else if (e.isControlDown() && e.getKeyCode() == KeyEvent.VK_X) {
                    cutSelection();
                    e.consume();
                }
                // Bold (Ctrl+B)
                else if (e.isControlDown() && e.getKeyCode() == KeyEvent.VK_B) {
                    toggleBold();
                    e.consume();
                }
                // Italic (Ctrl+I)
                else if (e.isControlDown() && e.getKeyCode() == KeyEvent.VK_I) {
                    toggleItalic();
                    e.consume();
                }
                // Underline (Ctrl+U)
                else if (e.isControlDown() && e.getKeyCode() == KeyEvent.VK_U) {
                    toggleUnderline();
                    e.consume();
                }
                // Delete cell content (Delete key)
                else if (e.getKeyCode() == KeyEvent.VK_DELETE) {
                    clearSelectedCells();
                    e.consume();
                }
                // Format painter (Ctrl+Shift+C)
                else if (e.isControlDown() && e.isShiftDown() && e.getKeyCode() == KeyEvent.VK_C) {
                    copyFormat();
                    e.consume();
                }
                // Format painter apply (Ctrl+Shift+V)
                else if (e.isControlDown() && e.isShiftDown() && e.getKeyCode() == KeyEvent.VK_V) {
                    pasteFormat();
                    e.consume();
                }
            }
        });
    }

    // ==================== SELECT ALL FUNCTIONALITY ====================

    private void selectAllCells() {
        table.selectAll();
        isSelectAllMode = true;
        statusBar.setText("All cells selected (" + model.getRowCount() + " rows × " + model.getColumnCount() + " columns)");
    }

    private void clearSelectedCells() {
        int[] rows = table.getSelectedRows();
        int[] cols = table.getSelectedColumns();
        
        if (rows.length > 0 && cols.length > 0) {
            model.batchUpdate(() -> {
                for (int row : rows) {
                    for (int col : cols) {
                        int modelRow = table.convertRowIndexToModel(row);
                        int modelCol = table.convertColumnIndexToModel(col);
                        model.setValueAt("", modelRow, modelCol);
                    }
                }
            });
            statusBar.setText("Cleared " + (rows.length * cols.length) + " cells");
        }
    }

    private void copySelection() {
        int[] rows = table.getSelectedRows();
        int[] cols = table.getSelectedColumns();
        
        if (rows.length > 0 && cols.length > 0) {
            StringBuilder sb = new StringBuilder();
            for (int row : rows) {
                for (int col : cols) {
                    int modelRow = table.convertRowIndexToModel(row);
                    int modelCol = table.convertColumnIndexToModel(col);
                    Object value = model.getRawValueAt(modelRow, modelCol);
                    if (value != null) {
                        sb.append(value.toString());
                    }
                    if (col < cols[cols.length - 1]) {
                        sb.append("\t");
                    }
                }
                if (row < rows[rows.length - 1]) {
                    sb.append("\n");
                }
            }
            
            StringSelection stringSelection = new StringSelection(sb.toString());
            java.awt.datatransfer.Clipboard clipboard = java.awt.Toolkit.getDefaultToolkit().getSystemClipboard();
            clipboard.setContents(stringSelection, null);
            
            statusBar.setText("Copied " + (rows.length * cols.length) + " cells to clipboard");
        }
    }

    private void cutSelection() {
        copySelection();
        clearSelectedCells();
    }

    private void pasteToSelection() {
        try {
            java.awt.datatransfer.Clipboard clipboard = java.awt.Toolkit.getDefaultToolkit().getSystemClipboard();
            Transferable transferable = clipboard.getContents(this);
            
            if (transferable != null && transferable.isDataFlavorSupported(DataFlavor.stringFlavor)) {
                String data = (String) transferable.getTransferData(DataFlavor.stringFlavor);
                int startRow = table.getSelectedRow();
                int startCol = table.getSelectedColumn();
                
                if (startRow == -1) startRow = 0;
                if (startCol == -1) startCol = 0;
                
                pasteData(data, startRow, startCol);
            }
        } catch (Exception ex) {
            JOptionPane.showMessageDialog(this, "Paste failed: " + ex.getMessage());
        }
    }

    // ==================== COMPREHENSIVE FORMATTING ====================

    private void updateFormattingToolbar() {
        int row = table.getSelectedRow();
        int col = table.getSelectedColumn();
        
        if (row != -1 && col != -1) {
            int modelRow = table.convertRowIndexToModel(row);
            int modelCol = table.convertColumnIndexToModel(col);
            
            // Update font controls
            Font currentFont = model.getCellFont(modelRow, modelCol);
            if (currentFont != null) {
                fontCombo.setSelectedItem(currentFont.getFamily());
                fontSizeSpinner.setValue(currentFont.getSize());
                boldCheckBox.setSelected(currentFont.isBold());
                italicCheckBox.setSelected(currentFont.isItalic());
                underlineCheckBox.setSelected(model.isUnderline(modelRow, modelCol));
            }
            
            // Update alignment
            String horizontalAlign = model.getHorizontalAlignment(modelRow, modelCol);
            if (horizontalAlign != null) {
                horizontalAlignmentCombo.setSelectedItem(horizontalAlign);
            }
            
            String verticalAlign = model.getVerticalAlignment(modelRow, modelCol);
            if (verticalAlign != null) {
                verticalAlignmentCombo.setSelectedItem(verticalAlign);
            }
            
            // Update number format
            String cellFormat = model.getCellFormat(modelRow, modelCol);
            if (cellFormat != null) {
                numberFormatCombo.setSelectedItem(cellFormat);
            }
        }
    }

    private void toggleBold() {
        applyToSelectedCells((row, col) -> {
            boolean currentBold = model.isBold(row, col);
            model.setBold(row, col, !currentBold);
            updateCellFont(row, col);
        });
    }

    private void toggleItalic() {
        applyToSelectedCells((row, col) -> {
            boolean currentItalic = model.isItalic(row, col);
            model.setItalic(row, col, !currentItalic);
            updateCellFont(row, col);
        });
    }

    private void toggleUnderline() {
        applyToSelectedCells((row, col) -> {
            boolean currentUnderline = model.isUnderline(row, col);
            model.setUnderline(row, col, !currentUnderline);
        });
    }

    private void updateCellFont(int row, int col) {
        Font currentFont = model.getCellFont(row, col);
        String fontName = currentFont != null ? currentFont.getFamily() : "Arial";
        int fontSize = currentFont != null ? currentFont.getSize() : 12;
        int fontStyle = Font.PLAIN;
        
        if (model.isBold(row, col)) fontStyle |= Font.BOLD;
        if (model.isItalic(row, col)) fontStyle |= Font.ITALIC;
        
        Font newFont = new Font(fontName, fontStyle, fontSize);
        model.setCellFont(row, col, newFont);
    }

    private void applyHorizontalAlignment() {
        String alignment = (String) horizontalAlignmentCombo.getSelectedItem();
        applyToSelectedCells((row, col) -> {
            model.setHorizontalAlignment(row, col, alignment);
        });
    }

    private void applyVerticalAlignment() {
        String alignment = (String) verticalAlignmentCombo.getSelectedItem();
        applyToSelectedCells((row, col) -> {
            model.setVerticalAlignment(row, col, alignment);
        });
    }

    private void applyBorder() {
        String borderStyle = (String) borderStyleCombo.getSelectedItem();
        applyToSelectedCells((row, col) -> {
            model.setBorder(row, col, borderStyle);
        });
    }

    // ==================== ENHANCED CELL OPERATIONS INTERFACE ====================

    private interface CellOperation {
        void execute(int row, int col);
    }

    private void applyToSelectedCells(CellOperation operation) {
        int[] rows = table.getSelectedRows();
        int[] cols = table.getSelectedColumns();
        
        if (rows.length > 0 && cols.length > 0) {
            model.batchUpdate(() -> {
                for (int row : rows) {
                    for (int col : cols) {
                        int modelRow = table.convertRowIndexToModel(row);
                        int modelCol = table.convertColumnIndexToModel(col);
                        operation.execute(modelRow, modelCol);
                    }
                }
            });
            statusBar.setText("Applied to " + (rows.length * cols.length) + " cells");
        } else {
            JOptionPane.showMessageDialog(this, "Please select cells to apply the operation.");
        }
    }

    // ==================== DRAG SELECTION & FORMAT PAINTER ====================

    private void setupDragAndDrop() {
        // Enable drag and drop for the table
        table.setDragEnabled(true);
        table.setDropMode(javax.swing.DropMode.USE_SELECTION);
        
        // Set up transfer handler for drag and drop
        table.setTransferHandler(new TableTransferHandler());
        
        // Add mouse listeners for drag selection
        table.addMouseListener(new MouseAdapter() {
            @Override
            public void mousePressed(MouseEvent e) {
                if (isFormatPainterActive && e.getButton() == MouseEvent.BUTTON1) {
                    // Apply format painter
                    applyFormatPainter(e);
                    return;
                }
                
                Point point = e.getPoint();
                int row = table.rowAtPoint(point);
                int col = table.columnAtPoint(point);
                
                if (row >= 0 && col >= 0) {
                    // Check if click is on the drag handle (bottom-right corner of cell)
                    Rectangle cellRect = table.getCellRect(row, col, true);
                    Rectangle dragHandle = new Rectangle(
                        cellRect.x + cellRect.width - 8,
                        cellRect.y + cellRect.height - 8,
                        6, 6
                    );
                    
                    if (dragHandle.contains(point)) {
                        dragStartPoint = new Point(row, col);
                        table.setCursor(Cursor.getPredefinedCursor(Cursor.CROSSHAIR_CURSOR));
                    }
                }
            }
            
            @Override
            public void mouseReleased(MouseEvent e) {
                if (dragStartPoint != null) {
                    handleDragSelection();
                }
                dragStartPoint = null;
                dragEndPoint = null;
                dragRectangle = null;
                isDragging = false;
                table.setCursor(Cursor.getDefaultCursor());
                table.repaint();
            }
        });
        
        table.addMouseMotionListener(new MouseMotionAdapter() {
            @Override
            public void mouseDragged(MouseEvent e) {
                if (dragStartPoint != null) {
                    Point point = e.getPoint();
                    int row = table.rowAtPoint(point);
                    int col = table.columnAtPoint(point);
                    
                    if (row >= 0 && col >= 0) {
                        dragEndPoint = new Point(row, col);
                        isDragging = true;
                        
                        // Calculate drag rectangle
                        int startRow = Math.min(dragStartPoint.x, dragEndPoint.x);
                        int endRow = Math.max(dragStartPoint.x, dragEndPoint.x);
                        int startCol = Math.min(dragStartPoint.y, dragEndPoint.y);
                        int endCol = Math.max(dragStartPoint.y, dragEndPoint.y);
                        
                        dragRectangle = new Rectangle(startCol, startRow, 
                                                    endCol - startCol + 1, endRow - startRow + 1);
                        
                        table.repaint();
                    }
                }
            }
        });
    }

    private void setupFormatPainter() {
        // Format painter will be activated by a button in the toolbar
    }

    private void handleDragSelection() {
        if (dragStartPoint != null && dragEndPoint != null) {
            int startRow = Math.min(dragStartPoint.x, dragEndPoint.x);
            int endRow = Math.max(dragStartPoint.x, dragEndPoint.x);
            int startCol = Math.min(dragStartPoint.y, dragEndPoint.y);
            int endCol = Math.max(dragStartPoint.y, dragEndPoint.y);
            
            // Select the dragged range
            table.setRowSelectionInterval(startRow, endRow);
            table.setColumnSelectionInterval(startCol, endCol);
            
            // Fill the selection with data (Excel-like auto-fill)
            autoFillSelection(startRow, endRow, startCol, endCol);
        }
    }

    private void autoFillSelection(int startRow, int endRow, int startCol, int endCol) {
        // Only auto-fill if we have a range (not a single cell)
        if ((endRow - startRow) > 0 || (endCol - startCol) > 0) {
            int result = JOptionPane.showConfirmDialog(this,
                "Do you want to auto-fill the selected range?",
                "Auto Fill",
                JOptionPane.YES_NO_OPTION);
            
            if (result == JOptionPane.YES_OPTION) {
                model.batchUpdate(() -> {
                    // Copy pattern from source cell to target range
                    Object sourceValue = model.getValueAt(startRow, startCol);
                    String sourceFormula = model.getCellFormula(startRow, startCol);
                    
                    for (int row = startRow; row <= endRow; row++) {
                        for (int col = startCol; col <= endCol; col++) {
                            if (row != startRow || col != startCol) { // Skip source cell
                                if (sourceFormula != null && sourceFormula.startsWith("=")) {
                                    // Handle formula auto-fill with relative references
                                    String adjustedFormula = adjustFormulaReferences(
                                        sourceFormula, startRow, startCol, row, col);
                                    model.setCellFormula(row, col, adjustedFormula);
                                } else if (sourceValue instanceof Number) {
                                    // Handle numeric series
                                    Number numValue = (Number) sourceValue;
                                    double increment = (row - startRow) + (col - startCol);
                                    model.setValueAt(numValue.doubleValue() + increment, row, col);
                                } else if (sourceValue != null) {
                                    // Copy value for non-numeric data
                                    model.setValueAt(sourceValue.toString(), row, col);
                                }
                            }
                        }
                    }
                });
            }
        }
    }

    private String adjustFormulaReferences(String formula, int sourceRow, int sourceCol, int targetRow, int targetCol) {
        // Adjust cell references in formulas (basic implementation)
        Pattern cellRefPattern = Pattern.compile("([A-Z]+)(\\d+)");
        Matcher matcher = cellRefPattern.matcher(formula);
        StringBuffer result = new StringBuffer();
        
        while (matcher.find()) {
            String colStr = matcher.group(1);
            int refRow = Integer.parseInt(matcher.group(2)) - 1;
            
            int refCol = columnNameToIndex(colStr);
            
            // Calculate relative position
            int rowDiff = refRow - sourceRow;
            int colDiff = refCol - sourceCol;
            
            // Apply same relative position to target
            int newRow = targetRow + rowDiff;
            int newCol = targetCol + colDiff;
            
            String newColStr = calculateColumnName(newCol);
            String replacement = newColStr + (newRow + 1);
            
            matcher.appendReplacement(result, replacement);
        }
        matcher.appendTail(result);
        return result.toString();
    }

    // ==================== FORMAT PAINTER FUNCTIONALITY ====================

    public void activateFormatPainter() {
        isFormatPainterActive = true;
        table.setCursor(Cursor.getPredefinedCursor(Cursor.HAND_CURSOR));
        statusBar.setText("Format Painter: Click on a cell to copy its format, then click on target cells");
    }

    public void deactivateFormatPainter() {
        isFormatPainterActive = false;
        sourceFormat = null;
        table.setCursor(Cursor.getDefaultCursor());
        statusBar.setText("Ready");
    }

    private void captureSourceFormat(int row, int col) {
        sourceFormat = new CellFormat();
        sourceFormat.backgroundColor = model.getCellColor(row, col);
        sourceFormat.textColor = model.getTextColor(row, col);
        sourceFormat.font = model.getCellFont(row, col);
        sourceFormat.isBold = model.isBold(row, col);
        sourceFormat.isItalic = model.isItalic(row, col);
        sourceFormat.isUnderline = model.isUnderline(row, col);
        sourceFormat.horizontalAlignment = model.getHorizontalAlignment(row, col);
        sourceFormat.verticalAlignment = model.getVerticalAlignment(row, col);
        sourceFormat.cellFormat = model.getCellFormat(row, col);
        sourceFormat.borderStyle = model.getBorder(row, col);
    }

    private void applyFormatPainter(MouseEvent e) {
        Point point = e.getPoint();
        int row = table.rowAtPoint(point);
        int col = table.columnAtPoint(point);
        
        if (row >= 0 && col >= 0) {
            if (sourceFormat == null) {
                // First click - capture source format
                captureSourceFormat(row, col);
                statusBar.setText("Format Painter: Format captured. Click on target cells to apply");
            } else {
                // Subsequent clicks - apply format
                applyFormatToCell(row, col, sourceFormat);
                
                // Option: Keep format painter active for multiple applications
                // To deactivate after one application, uncomment the next line:
                // deactivateFormatPainter();
            }
        }
    }

    private void applyFormatToCell(int row, int col, CellFormat format) {
        model.setCellColor(row, col, format.backgroundColor);
        model.setTextColor(row, col, format.textColor);
        model.setCellFont(row, col, format.font);
        model.setBold(row, col, format.isBold);
        model.setItalic(row, col, format.isItalic);
        model.setUnderline(row, col, format.isUnderline);
        model.setHorizontalAlignment(row, col, format.horizontalAlignment);
        model.setVerticalAlignment(row, col, format.verticalAlignment);
        model.setCellFormat(row, col, format.cellFormat);
        model.setBorder(row, col, format.borderStyle);
    }

    public void copyFormat() {
        int row = table.getSelectedRow();
        int col = table.getSelectedColumn();
        if (row != -1 && col != -1) {
            captureSourceFormat(row, col);
            activateFormatPainter();
        } else {
            JOptionPane.showMessageDialog(this, "Please select a cell to copy format from.");
        }
    }

    public void pasteFormat() {
        if (sourceFormat != null) {
            int[] rows = table.getSelectedRows();
            int[] cols = table.getSelectedColumns();
            
            if (rows.length > 0 && cols.length > 0) {
                model.batchUpdate(() -> {
                    for (int row : rows) {
                        for (int col : cols) {
                            int modelRow = table.convertRowIndexToModel(row);
                            int modelCol = table.convertColumnIndexToModel(col);
                            applyFormatToCell(modelRow, modelCol, sourceFormat);
                        }
                    }
                });
                statusBar.setText("Format applied to " + (rows.length * cols.length) + " cells");
            } else {
                JOptionPane.showMessageDialog(this, "Please select cells to apply the format to.");
            }
        } else {
            JOptionPane.showMessageDialog(this, "No format copied. Use 'Copy Format' first.");
        }
    }

    public void clearFormat() {
        int[] rows = table.getSelectedRows();
        int[] cols = table.getSelectedColumns();
        
        if (rows.length > 0 && cols.length > 0) {
            model.batchUpdate(() -> {
                for (int row : rows) {
                    for (int col : cols) {
                        int modelRow = table.convertRowIndexToModel(row);
                        int modelCol = table.convertColumnIndexToModel(col);
                        
                        // Clear all formatting
                        model.setCellColor(modelRow, modelCol, null);
                        model.setTextColor(modelRow, modelCol, null);
                        model.setCellFont(modelRow, modelCol, null);
                        model.setBold(modelRow, modelCol, false);
                        model.setItalic(modelRow, modelCol, false);
                        model.setUnderline(modelRow, modelCol, false);
                        model.setHorizontalAlignment(modelRow, modelCol, "Left");
                        model.setVerticalAlignment(modelRow, modelCol, "Center");
                        model.setCellFormat(modelRow, modelCol, "General");
                        model.setBorder(modelRow, modelCol, "None");
                    }
                }
            });
            statusBar.setText("Formatting cleared from " + (rows.length * cols.length) + " cells");
        } else {
            JOptionPane.showMessageDialog(this, "Please select cells to clear formatting.");
        }
    }

    // ==================== DRAG AND DROP TRANSFER HANDLER ====================

    private class TableTransferHandler extends TransferHandler {
        @Override
        public int getSourceActions(JComponent c) {
            return COPY_OR_MOVE;
        }

        @Override
        protected Transferable createTransferable(JComponent c) {
            if (c instanceof JTable) {
                JTable table = (JTable) c;
                int[] rows = table.getSelectedRows();
                int[] cols = table.getSelectedColumns();
                
                if (rows.length > 0 && cols.length > 0) {
                    StringBuilder sb = new StringBuilder();
                    for (int row : rows) {
                        for (int col : cols) {
                            int modelRow = table.convertRowIndexToModel(row);
                            int modelCol = table.convertColumnIndexToModel(col);
                            Object value = model.getRawValueAt(modelRow, modelCol);
                            if (value != null) {
                                sb.append(value.toString());
                            }
                            if (col < cols[cols.length - 1]) {
                                sb.append("\t");
                            }
                        }
                        if (row < rows[rows.length - 1]) {
                            sb.append("\n");
                        }
                    }
                    return new StringSelection(sb.toString());
                }
            }
            return null;
        }

        @Override
        public boolean canImport(TransferHandler.TransferSupport support) {
            return support.isDataFlavorSupported(DataFlavor.stringFlavor);
        }

        @Override
        public boolean importData(TransferHandler.TransferSupport support) {
            if (!canImport(support)) {
                return false;
            }

            try {
                JTable.DropLocation dl = (JTable.DropLocation) support.getDropLocation();
                int row = dl.getRow();
                int col = dl.getColumn();

                Transferable t = support.getTransferable();
                String data = (String) t.getTransferData(DataFlavor.stringFlavor);

                if (data != null && !data.isEmpty()) {
                    pasteData(data, row, col);
                    return true;
                }
            } catch (Exception e) {
                e.printStackTrace();
            }
            return false;
        }
    }

    private void pasteData(String data, int startRow, int startCol) {
        String[] rows = data.split("\n");
        model.batchUpdate(() -> {
            for (int i = 0; i < rows.length; i++) {
                String[] cells = rows[i].split("\t");
                for (int j = 0; j < cells.length; j++) {
                    int targetRow = startRow + i;
                    int targetCol = startCol + j;
                    
                    if (targetRow < model.getRowCount() && targetCol < model.getColumnCount()) {
                        model.setValueAt(cells[j], targetRow, targetCol);
                    }
                }
            }
        });
        statusBar.setText("Pasted data to " + rows.length + " rows");
    }

    // ==================== ENHANCED TABLE MODEL WITH FORMATTING ====================
    
    public class OptimizedTableModel extends AbstractTableModel {
        private int rowCount = 1000;
        private int colCount = 26;
        
        // Enhanced storage
        private final Map<String, Object> dataMap = new HashMap<>();
        private final Map<String, Color> cellColorMap = new HashMap<>();
        private final Map<String, Color> textColorMap = new HashMap<>();
        private final Map<String, ImageIcon> cellImageMap = new HashMap<>();
        private final Map<String, String> formulaMap = new HashMap<>();
        private final Map<String, String> formatMap = new HashMap<>();
        private final Map<String, String> commentMap = new HashMap<>();
        private final Map<String, Font> fontMap = new HashMap<>();
        private final Map<String, Boolean> boldMap = new HashMap<>();
        private final Map<String, Boolean> italicMap = new HashMap<>();
        private final Map<String, Boolean> underlineMap = new HashMap<>();
        private final Map<String, Boolean> protectedMap = new HashMap<>();
        private final Map<String, String> horizontalAlignmentMap = new HashMap<>();
        private final Map<String, String> verticalAlignmentMap = new HashMap<>();
        private final Map<String, String> borderMap = new HashMap<>();
        private final Map<String, String> customFormatMap = new HashMap<>();
        
        private boolean inBatchMode = false;
        private final XlsxEditorPanel parentPanel;
        
        // OPTIMIZATION: Cache column names
        private final Map<Integer, String> columnNameCache = new HashMap<>();

        public OptimizedTableModel(XlsxEditorPanel parentPanel) {
            this.parentPanel = parentPanel;
        }

        // OPTIMIZATION: Fast key generation
        private String createKey(int row, int col) {
            return row + "_" + col;
        }

        // Batch processing - O(1) notification instead of O(n)
        public void batchUpdate(Runnable updateOperation) {
            inBatchMode = true;
            try {
                updateOperation.run();
                fireTableDataChanged();
            } finally {
                inBatchMode = false;
            }
        }

        @Override
        public int getRowCount() {
            return rowCount;
        }

        @Override
        public int getColumnCount() {
            return colCount;
        }

        @Override
        public Object getValueAt(int row, int col) {
            String key = createKey(row, col);
            
            // Priority: Formula > Image > Data
            if (formulaMap.containsKey(key)) {
                return calculateFormulaValue(row, col, key);
            }
            
            if (cellImageMap.containsKey(key)) {
                return cellImageMap.get(key);
            }
            
            Object value = dataMap.get(key);
            return formatCellValue(value != null ? value : "", row, col);
        }

        @Override
        public void setValueAt(Object value, int row, int col) {
            if (isCellProtected(row, col)) {
                JOptionPane.showMessageDialog(parentPanel, "Cell is protected and cannot be modified.");
                return;
            }
            
            ensureCapacity(row, col);
            
            String key = createKey(row, col);
            String strValue = value != null ? value.toString() : "";
            
            // Handle formulas
            if (strValue.startsWith("=")) {
                formulaMap.put(key, strValue);
                markCellDirty(row, col);
            } else {
                formulaMap.remove(key);
            }
            
            // Handle empty values - REMOVE to save space
            if (value == null || strValue.isEmpty()) {
                dataMap.remove(key);
                cellImageMap.remove(key);
            } 
            else if (value instanceof ImageIcon) {
                cellImageMap.put(key, (ImageIcon) value);
                dataMap.put(key, "[Image]");
            }
            else {
                cellImageMap.remove(key);
                dataMap.put(key, value);
            }
            
            if (!inBatchMode) {
                fireTableCellUpdated(row, col);
            }
            
            onCellUpdated(row, col);
            parentPanel.addToHistory("EDIT", row, col, value);
        }

        @Override
        public boolean isCellEditable(int row, int col) {
            return !isCellProtected(row, col);
        }

        @Override
        public String getColumnName(int col) {
            return columnNameCache.computeIfAbsent(col, c -> parentPanel.calculateColumnName(c));
        }

        @Override
        public Class<?> getColumnClass(int col) {
            return Object.class;
        }

        // ========== ENHANCED CELL MANAGEMENT ==========
        
        public Component prepareTableCellRenderer(Component comp, int row, int col, 
                                                boolean alternateRowColors, boolean tableMode) {
            boolean isDarkTheme = parentPanel.isDarkTheme();
            
            Color defaultBackground = isDarkTheme ? new Color(45, 45, 45) : Color.WHITE;
            Color alternateBackground = isDarkTheme ? new Color(60, 60, 60) : new Color(240, 240, 240);
            Color defaultForeground = isDarkTheme ? Color.WHITE : Color.BLACK;
            
            if (alternateRowColors && tableMode) {
                comp.setBackground(row % 2 == 0 ? defaultBackground : alternateBackground);
            } else {
                comp.setBackground(defaultBackground);
            }
            
            Color customColor = getCellColor(row, col);
            if (customColor != null) {
                comp.setBackground(customColor);
            }
            
            if (hasImage(row, col)) {
                if (comp instanceof JLabel) {
                    ((JLabel) comp).setIcon(getCellImage(row, col));
                    ((JLabel) comp).setText("");
                }
            }
            
            // Apply font styling
            Font cellFont = getCellFont(row, col);
            if (cellFont != null) {
                comp.setFont(cellFont);
            }
            
            // Apply text color
            Color textColor = getTextColor(row, col);
            if (textColor != null) {
                comp.setForeground(textColor);
            } else {
                comp.setForeground(defaultForeground);
            }
            
            // Apply alignment
            String horizontalAlign = getHorizontalAlignment(row, col);
            if (horizontalAlign != null && comp instanceof JLabel) {
                JLabel label = (JLabel) comp;
                switch (horizontalAlign) {
                    case "Left": label.setHorizontalAlignment(SwingConstants.LEFT); break;
                    case "Center": label.setHorizontalAlignment(SwingConstants.CENTER); break;
                    case "Right": label.setHorizontalAlignment(SwingConstants.RIGHT); break;
                    case "Justify": label.setHorizontalAlignment(SwingConstants.LEADING); break;
                }
            }
            
            String verticalAlign = getVerticalAlignment(row, col);
            if (verticalAlign != null && comp instanceof JLabel) {
                JLabel label = (JLabel) comp;
                switch (verticalAlign) {
                    case "Top": label.setVerticalAlignment(SwingConstants.TOP); break;
                    case "Center": label.setVerticalAlignment(SwingConstants.CENTER); break;
                    case "Bottom": label.setVerticalAlignment(SwingConstants.BOTTOM); break;
                }
            }
            
            // Apply underline
            if (isUnderline(row, col)) {
                Font font = comp.getFont();
                Map<java.awt.font.TextAttribute, Object> attributes = new HashMap<>(font.getAttributes());
                attributes.put(java.awt.font.TextAttribute.UNDERLINE, java.awt.font.TextAttribute.UNDERLINE_ON);
                comp.setFont(font.deriveFont(attributes));
            }
            
            String format = getCellFormat(row, col);
            if (format != null && comp instanceof JLabel) {
                applyNumberFormatting((JLabel) comp, getRawValueAt(row, col), format);
            }
            
            // Apply border
            String borderStyle = getBorder(row, col);
            if (borderStyle != null && !borderStyle.equals("None") && comp instanceof JComponent) {
                Border border = createBorder(borderStyle);
                ((JComponent) comp).setBorder(border);
            }
            
            return comp;
        }

       private Border createBorder(String borderStyle) {
    Color borderColor = isDarkTheme() ? Color.LIGHT_GRAY : Color.DARK_GRAY;
    
    switch (borderStyle) {
        case "Thin":
            return BorderFactory.createLineBorder(borderColor, 1);
        case "Medium":
            return BorderFactory.createLineBorder(borderColor, 2);
        case "Thick":
            return BorderFactory.createLineBorder(borderColor, 3);
        case "Dashed":
            // Create dashed border using StrokeBorder
            float[] dashPattern = {5.0f, 5.0f};
            return BorderFactory.createStrokeBorder(
                new BasicStroke(1.0f, BasicStroke.CAP_BUTT, BasicStroke.JOIN_MITER, 
                              10.0f, dashPattern, 0.0f), borderColor);
        case "Dotted":
            // Create dotted border using StrokeBorder
            float[] dotPattern = {1.0f, 3.0f};
            return BorderFactory.createStrokeBorder(
                new BasicStroke(1.0f, BasicStroke.CAP_ROUND, BasicStroke.JOIN_ROUND, 
                              10.0f, dotPattern, 0.0f), borderColor);
        default: // "None"
            return BorderFactory.createEmptyBorder();
    }
}

        // ========== ENHANCED CELL PROPERTIES ==========
        
        public void setCellColor(int row, int col, Color color) {
            String key = createKey(row, col);
            if (color != null) {
                cellColorMap.put(key, color);
            } else {
                cellColorMap.remove(key);
            }
            if (!inBatchMode) fireTableCellUpdated(row, col);
        }

        public Color getCellColor(int row, int col) {
            return cellColorMap.get(createKey(row, col));
        }

        public void setTextColor(int row, int col, Color color) {
            String key = createKey(row, col);
            if (color != null) {
                textColorMap.put(key, color);
            } else {
                textColorMap.remove(key);
            }
            if (!inBatchMode) fireTableCellUpdated(row, col);
        }

        public Color getTextColor(int row, int col) {
            return textColorMap.get(createKey(row, col));
        }

        public void setCellFont(int row, int col, Font font) {
            String key = createKey(row, col);
            if (font != null) {
                fontMap.put(key, font);
            } else {
                fontMap.remove(key);
            }
            if (!inBatchMode) fireTableCellUpdated(row, col);
        }

        public Font getCellFont(int row, int col) {
            return fontMap.get(createKey(row, col));
        }

        public void setBold(int row, int col, boolean bold) {
            String key = createKey(row, col);
            if (bold) {
                boldMap.put(key, true);
            } else {
                boldMap.remove(key);
            }
            if (!inBatchMode) fireTableCellUpdated(row, col);
        }

        public boolean isBold(int row, int col) {
            return boldMap.containsKey(createKey(row, col));
        }

        public void setItalic(int row, int col, boolean italic) {
            String key = createKey(row, col);
            if (italic) {
                italicMap.put(key, true);
            } else {
                italicMap.remove(key);
            }
            if (!inBatchMode) fireTableCellUpdated(row, col);
        }

        public boolean isItalic(int row, int col) {
            return italicMap.containsKey(createKey(row, col));
        }

        public void setUnderline(int row, int col, boolean underline) {
            String key = createKey(row, col);
            if (underline) {
                underlineMap.put(key, true);
            } else {
                underlineMap.remove(key);
            }
            if (!inBatchMode) fireTableCellUpdated(row, col);
        }

        public boolean isUnderline(int row, int col) {
            return underlineMap.containsKey(createKey(row, col));
        }

        public void setHorizontalAlignment(int row, int col, String alignment) {
            String key = createKey(row, col);
            if (alignment != null && !alignment.equals("Left")) {
                horizontalAlignmentMap.put(key, alignment);
            } else {
                horizontalAlignmentMap.remove(key);
            }
            if (!inBatchMode) fireTableCellUpdated(row, col);
        }

        public String getHorizontalAlignment(int row, int col) {
            return horizontalAlignmentMap.get(createKey(row, col));
        }

        public void setVerticalAlignment(int row, int col, String alignment) {
            String key = createKey(row, col);
            if (alignment != null && !alignment.equals("Center")) {
                verticalAlignmentMap.put(key, alignment);
            } else {
                verticalAlignmentMap.remove(key);
            }
            if (!inBatchMode) fireTableCellUpdated(row, col);
        }

        public String getVerticalAlignment(int row, int col) {
            return verticalAlignmentMap.get(createKey(row, col));
        }

        public void setCellProtected(int row, int col, boolean protectedCell) {
            String key = createKey(row, col);
            if (protectedCell) {
                protectedMap.put(key, true);
            } else {
                protectedMap.remove(key);
            }
        }

        public boolean isCellProtected(int row, int col) {
            return protectedMap.containsKey(createKey(row, col));
        }

        public void setCellFormat(int row, int col, String format) {
            String key = createKey(row, col);
            if (format != null && !format.equals("General")) {
                formatMap.put(key, format);
            } else {
                formatMap.remove(key);
            }
            if (!inBatchMode) fireTableCellUpdated(row, col);
        }

        public String getCellFormat(int row, int col) {
            return formatMap.get(createKey(row, col));
        }

        public void setBorder(int row, int col, String borderStyle) {
            String key = createKey(row, col);
            if (borderStyle != null && !borderStyle.equals("None")) {
                borderMap.put(key, borderStyle);
            } else {
                borderMap.remove(key);
            }
            if (!inBatchMode) fireTableCellUpdated(row, col);
        }

        public String getBorder(int row, int col) {
            return borderMap.get(createKey(row, col));
        }

        public void setCustomFormat(int row, int col, String format) {
            String key = createKey(row, col);
            if (format != null && !format.isEmpty()) {
                customFormatMap.put(key, format);
            } else {
                customFormatMap.remove(key);
            }
            if (!inBatchMode) fireTableCellUpdated(row, col);
        }

        public String getCustomFormat(int row, int col) {
            return customFormatMap.get(createKey(row, col));
        }

        // ========== ROW AND COLUMN MANAGEMENT ==========
        
        public void insertRow(int atRow) {
            batchUpdate(() -> {
                // Shift data down
                for (int row = rowCount - 1; row >= atRow; row--) {
                    for (int col = 0; col < colCount; col++) {
                        String oldKey = createKey(row, col);
                        String newKey = createKey(row + 1, col);
                        
                        if (dataMap.containsKey(oldKey)) dataMap.put(newKey, dataMap.remove(oldKey));
                        if (cellColorMap.containsKey(oldKey)) cellColorMap.put(newKey, cellColorMap.remove(oldKey));
                        if (textColorMap.containsKey(oldKey)) textColorMap.put(newKey, textColorMap.remove(oldKey));
                        if (formulaMap.containsKey(oldKey)) formulaMap.put(newKey, formulaMap.remove(oldKey));
                        if (formatMap.containsKey(oldKey)) formatMap.put(newKey, formatMap.remove(oldKey));
                        if (commentMap.containsKey(oldKey)) commentMap.put(newKey, commentMap.remove(oldKey));
                        if (fontMap.containsKey(oldKey)) fontMap.put(newKey, fontMap.remove(oldKey));
                        if (boldMap.containsKey(oldKey)) boldMap.put(newKey, boldMap.remove(oldKey));
                        if (italicMap.containsKey(oldKey)) italicMap.put(newKey, italicMap.remove(oldKey));
                        if (underlineMap.containsKey(oldKey)) underlineMap.put(newKey, underlineMap.remove(oldKey));
                        if (protectedMap.containsKey(oldKey)) protectedMap.put(newKey, protectedMap.remove(oldKey));
                        if (horizontalAlignmentMap.containsKey(oldKey)) horizontalAlignmentMap.put(newKey, horizontalAlignmentMap.remove(oldKey));
                        if (verticalAlignmentMap.containsKey(oldKey)) verticalAlignmentMap.put(newKey, verticalAlignmentMap.remove(oldKey));
                        if (borderMap.containsKey(oldKey)) borderMap.put(newKey, borderMap.remove(oldKey));
                    }
                }
                
                rowCount++;
                fireTableStructureChanged();
            });
        }

        public void insertColumn(int atCol) {
            batchUpdate(() -> {
                // Shift data right
                for (int row = 0; row < rowCount; row++) {
                    for (int col = colCount - 1; col >= atCol; col--) {
                        String oldKey = createKey(row, col);
                        String newKey = createKey(row, col + 1);
                        
                        if (dataMap.containsKey(oldKey)) dataMap.put(newKey, dataMap.remove(oldKey));
                        if (cellColorMap.containsKey(oldKey)) cellColorMap.put(newKey, cellColorMap.remove(oldKey));
                        if (textColorMap.containsKey(oldKey)) textColorMap.put(newKey, textColorMap.remove(oldKey));
                        if (formulaMap.containsKey(oldKey)) formulaMap.put(newKey, formulaMap.remove(oldKey));
                        if (formatMap.containsKey(oldKey)) formatMap.put(newKey, formatMap.remove(oldKey));
                        if (commentMap.containsKey(oldKey)) commentMap.put(newKey, commentMap.remove(oldKey));
                        if (fontMap.containsKey(oldKey)) fontMap.put(newKey, fontMap.remove(oldKey));
                        if (boldMap.containsKey(oldKey)) boldMap.put(newKey, boldMap.remove(oldKey));
                        if (italicMap.containsKey(oldKey)) italicMap.put(newKey, italicMap.remove(oldKey));
                        if (underlineMap.containsKey(oldKey)) underlineMap.put(newKey, underlineMap.remove(oldKey));
                        if (protectedMap.containsKey(oldKey)) protectedMap.put(newKey, protectedMap.remove(oldKey));
                        if (horizontalAlignmentMap.containsKey(oldKey)) horizontalAlignmentMap.put(newKey, horizontalAlignmentMap.remove(oldKey));
                        if (verticalAlignmentMap.containsKey(oldKey)) verticalAlignmentMap.put(newKey, verticalAlignmentMap.remove(oldKey));
                        if (borderMap.containsKey(oldKey)) borderMap.put(newKey, borderMap.remove(oldKey));
                    }
                }
                
                colCount++;
                columnNameCache.clear(); // Clear cache since column structure changed
                fireTableStructureChanged();
            });
        }

        public void deleteRow(int atRow) {
            if (atRow < 0 || atRow >= rowCount) return;
            
            batchUpdate(() -> {
                // Remove the row data
                for (int col = 0; col < colCount; col++) {
                    String key = createKey(atRow, col);
                    dataMap.remove(key);
                    cellColorMap.remove(key);
                    textColorMap.remove(key);
                    formulaMap.remove(key);
                    formatMap.remove(key);
                    commentMap.remove(key);
                    fontMap.remove(key);
                    boldMap.remove(key);
                    italicMap.remove(key);
                    underlineMap.remove(key);
                    protectedMap.remove(key);
                    horizontalAlignmentMap.remove(key);
                    verticalAlignmentMap.remove(key);
                    borderMap.remove(key);
                }
                
                // Shift data up
                for (int row = atRow + 1; row < rowCount; row++) {
                    for (int col = 0; col < colCount; col++) {
                        String oldKey = createKey(row, col);
                        String newKey = createKey(row - 1, col);
                        
                        if (dataMap.containsKey(oldKey)) dataMap.put(newKey, dataMap.remove(oldKey));
                        if (cellColorMap.containsKey(oldKey)) cellColorMap.put(newKey, cellColorMap.remove(oldKey));
                        if (textColorMap.containsKey(oldKey)) textColorMap.put(newKey, textColorMap.remove(oldKey));
                        if (formulaMap.containsKey(oldKey)) formulaMap.put(newKey, formulaMap.remove(oldKey));
                        if (formatMap.containsKey(oldKey)) formatMap.put(newKey, formatMap.remove(oldKey));
                        if (commentMap.containsKey(oldKey)) commentMap.put(newKey, commentMap.remove(oldKey));
                        if (fontMap.containsKey(oldKey)) fontMap.put(newKey, fontMap.remove(oldKey));
                        if (boldMap.containsKey(oldKey)) boldMap.put(newKey, boldMap.remove(oldKey));
                        if (italicMap.containsKey(oldKey)) italicMap.put(newKey, italicMap.remove(oldKey));
                        if (underlineMap.containsKey(oldKey)) underlineMap.put(newKey, underlineMap.remove(oldKey));
                        if (protectedMap.containsKey(oldKey)) protectedMap.put(newKey, protectedMap.remove(oldKey));
                        if (horizontalAlignmentMap.containsKey(oldKey)) horizontalAlignmentMap.put(newKey, horizontalAlignmentMap.remove(oldKey));
                        if (verticalAlignmentMap.containsKey(oldKey)) verticalAlignmentMap.put(newKey, verticalAlignmentMap.remove(oldKey));
                        if (borderMap.containsKey(oldKey)) borderMap.put(newKey, borderMap.remove(oldKey));
                    }
                }
                
                rowCount--;
                fireTableStructureChanged();
            });
        }

        public void deleteColumn(int atCol) {
            if (atCol < 0 || atCol >= colCount) return;
            
            batchUpdate(() -> {
                // Remove the column data
                for (int row = 0; row < rowCount; row++) {
                    String key = createKey(row, atCol);
                    dataMap.remove(key);
                    cellColorMap.remove(key);
                    textColorMap.remove(key);
                    formulaMap.remove(key);
                    formatMap.remove(key);
                    commentMap.remove(key);
                    fontMap.remove(key);
                    boldMap.remove(key);
                    italicMap.remove(key);
                    underlineMap.remove(key);
                    protectedMap.remove(key);
                    horizontalAlignmentMap.remove(key);
                    verticalAlignmentMap.remove(key);
                    borderMap.remove(key);
                }
                
                // Shift data left
                for (int row = 0; row < rowCount; row++) {
                    for (int col = atCol + 1; col < colCount; col++) {
                        String oldKey = createKey(row, col);
                        String newKey = createKey(row, col - 1);
                        
                        if (dataMap.containsKey(oldKey)) dataMap.put(newKey, dataMap.remove(oldKey));
                        if (cellColorMap.containsKey(oldKey)) cellColorMap.put(newKey, cellColorMap.remove(oldKey));
                        if (textColorMap.containsKey(oldKey)) textColorMap.put(newKey, textColorMap.remove(oldKey));
                        if (formulaMap.containsKey(oldKey)) formulaMap.put(newKey, formulaMap.remove(oldKey));
                        if (formatMap.containsKey(oldKey)) formatMap.put(newKey, formatMap.remove(oldKey));
                        if (commentMap.containsKey(oldKey)) commentMap.put(newKey, commentMap.remove(oldKey));
                        if (fontMap.containsKey(oldKey)) fontMap.put(newKey, fontMap.remove(oldKey));
                        if (boldMap.containsKey(oldKey)) boldMap.put(newKey, boldMap.remove(oldKey));
                        if (italicMap.containsKey(oldKey)) italicMap.put(newKey, italicMap.remove(oldKey));
                        if (underlineMap.containsKey(oldKey)) underlineMap.put(newKey, underlineMap.remove(oldKey));
                        if (protectedMap.containsKey(oldKey)) protectedMap.put(newKey, protectedMap.remove(oldKey));
                        if (horizontalAlignmentMap.containsKey(oldKey)) horizontalAlignmentMap.put(newKey, horizontalAlignmentMap.remove(oldKey));
                        if (verticalAlignmentMap.containsKey(oldKey)) verticalAlignmentMap.put(newKey, verticalAlignmentMap.remove(oldKey));
                        if (borderMap.containsKey(oldKey)) borderMap.put(newKey, borderMap.remove(oldKey));
                    }
                }
                
                colCount--;
                columnNameCache.clear(); // Clear cache since column structure changed
                fireTableStructureChanged();
            });
        }

        // ========== DATA ANALYSIS METHODS ==========
        
        public Object[] getColumnData(int col) {
            Object[] columnData = new Object[rowCount];
            for (int row = 0; row < rowCount; row++) {
                columnData[row] = getRawValueAt(row, col);
            }
            return columnData;
        }

        public Object[] getRowData(int row) {
            Object[] rowData = new Object[colCount];
            for (int col = 0; col < colCount; col++) {
                rowData[col] = getRawValueAt(row, col);
            }
            return rowData;
        }

        public Map<String, Object> getCellDependencies(int row, int col) {
            Map<String, Object> dependencies = new HashMap<>();
            String formula = getCellFormula(row, col);
            if (formula != null && formula.startsWith("=")) {
                // Extract cell references from formula
                Pattern pattern = Pattern.compile("([A-Z]+)(\\d+)");
                Matcher matcher = pattern.matcher(formula);
                while (matcher.find()) {
                    String colStr = matcher.group(1);
                    int refRow = Integer.parseInt(matcher.group(2)) - 1;
                    int refCol = parentPanel.columnNameToIndex(colStr);
                    dependencies.put(colStr + (refRow + 1), getRawValueAt(refRow, refCol));
                }
            }
            return dependencies;
        }

        // ========== PRIVATE HELPERS ==========
        
        private void ensureCapacity(int row, int col) {
            boolean changed = false;
            if (row >= rowCount) {
                rowCount = Math.max(rowCount * 2, row + 1);
                changed = true;
            }
            if (col >= colCount) {
                colCount = Math.max(colCount * 2, col + 1);
                changed = true;
            }
            if (changed && !inBatchMode) fireTableStructureChanged();
        }

        private void onCellUpdated(int row, int col) {
            if (parentPanel.autoExpandEnabled && parentPanel.tableMode) {
                checkAutoExpand();
            }
            if (parentPanel.showTotalRow) {
                parentPanel.updateTotalRowPanel();
            }
            markCellDirty(row, col);
        }

        private Object calculateFormulaValue(int row, int col, String key) {
            if (parentPanel.calculationCache.containsKey(key)) {
                return parentPanel.calculationCache.get(key);
            }
            
            String formula = formulaMap.get(key);
            try {
                Object result = parentPanel.formulaEngine.evaluate(formula, this, row, col);
                parentPanel.calculationCache.put(key, result);
                return result;
            } catch (Exception e) {
                return "#ERROR: " + e.getMessage();
            }
        }

        private Object formatCellValue(Object value, int row, int col) {
            String format = getCellFormat(row, col);
            if (format == null || value == null) return value;
            
            try {
                switch (format) {
                    case "Number":
                        if (value instanceof Number) {
                            return parentPanel.getNumberFormat().format(value);
                        }
                        break;
                    case "Currency":
                        if (value instanceof Number) {
                            return parentPanel.getCurrencyFormat().format(value);
                        }
                        break;
                    case "Percentage":
                        if (value instanceof Number) {
                            return parentPanel.getPercentFormat().format(((Number)value).doubleValue() / 100.0);
                        }
                        break;
                    case "Scientific":
                        if (value instanceof Number) {
                            return parentPanel.getScientificFormat().format(value);
                        }
                        break;
                    case "Accounting":
                        if (value instanceof Number) {
                            return parentPanel.getAccountingFormat().format(value);
                        }
                        break;
                    case "Date":
                        if (value instanceof Date) {
                            return parentPanel.getDateFormat().format(value);
                        }
                        break;
                }
            } catch (Exception e) {
                // Fallback
            }
            return value;
        }

        private void applyNumberFormatting(JLabel label, Object value, String format) {
            if (value instanceof Number) {
                label.setHorizontalAlignment(SwingConstants.RIGHT);
                try {
                    switch (format) {
                        case "Currency":
                            label.setText(parentPanel.getCurrencyFormat().format(value));
                            break;
                        case "Percentage":
                            label.setText(parentPanel.getPercentFormat().format(((Number)value).doubleValue() / 100.0));
                            break;
                        case "Number":
                            label.setText(parentPanel.getNumberFormat().format(value));
                            break;
                        case "Scientific":
                            label.setText(parentPanel.getScientificFormat().format(value));
                            break;
                        case "Accounting":
                            label.setText(parentPanel.getAccountingFormat().format(value));
                            break;
                    }
                } catch (Exception e) {
                    // Keep original
                }
            }
        }

        private void markCellDirty(int row, int col) {
            String key = createKey(row, col);
            parentPanel.dirtyCells.add(key);
            parentPanel.scheduleDelayedCalculation();
        }
        
        private void checkAutoExpand() {
            int lastRow = rowCount - 1;
            if (lastRow >= 0) {
                boolean nonEmpty = false;
                for (int col = 0; col < colCount; col++) {
                    Object value = getRawValueAt(lastRow, col);
                    if (value != null && !value.toString().trim().isEmpty()) {
                        nonEmpty = true;
                        break;
                    }
                }
                if (nonEmpty) {
                    fireTableRowsInserted(rowCount, rowCount);
                }
            }
        }

        // ========== BASIC METHODS ==========
        
        public ImageIcon getCellImage(int row, int col) {
            return cellImageMap.get(createKey(row, col));
        }

        public boolean hasImage(int row, int col) {
            return cellImageMap.containsKey(createKey(row, col));
        }

        public String getCellComment(int row, int col) {
            return commentMap.get(createKey(row, col));
        }

        public boolean hasComment(int row, int col) {
            return commentMap.containsKey(createKey(row, col));
        }

        public String getCellFormula(int row, int col) {
            return formulaMap.get(createKey(row, col));
        }

        public void setCellFormula(int row, int col, String formula) {
            String key = createKey(row, col);
            if (formula != null && formula.startsWith("=")) {
                formulaMap.put(key, formula);
                markCellDirty(row, col);
            } else {
                formulaMap.remove(key);
            }
            if (!inBatchMode) fireTableCellUpdated(row, col);
        }

        public Object getRawValueAt(int row, int col) {
            return dataMap.get(createKey(row, col));
        }

        public Object[][] getDataSnapshot() {
            Object[][] snapshot = new Object[rowCount][colCount];
            for (int i = 0; i < rowCount; i++) {
                for (int j = 0; j < colCount; j++) {
                    snapshot[i][j] = getRawValueAt(i, j);
                }
            }
            return snapshot;
        }

        public String[] getColumnNames() {
            String[] names = new String[colCount];
            for (int i = 0; i < colCount; i++) {
                names[i] = getColumnName(i);
            }
            return names;
        }

        public void loadData(Object[][] newData, String[] columnNames) {
            batchUpdate(() -> {
                if (newData == null || newData.length == 0) return;
                
                this.rowCount = newData.length;
                this.colCount = newData[0].length;
                this.dataMap.clear();
                
                for (int i = 0; i < rowCount; i++) {
                    for (int j = 0; j < colCount; j++) {
                        if (newData[i][j] != null && !newData[i][j].toString().isEmpty()) {
                            dataMap.put(createKey(i, j), newData[i][j]);
                        }
                    }
                }
            });
        }

        public void addNewRow() {
            parentPanel.handleAddRow();
        }

        public void deleteSelectedRow() {
            parentPanel.handleDeleteRow();
        }

        public void exportToCSV() {
            parentPanel.handleExportCSV();
        }

        public void saveToXlsx() {
            parentPanel.handleSave(new ActionEvent(this, ActionEvent.ACTION_PERFORMED, "save"));
        }

        public void insertImageIntoSelectedCell() {
            parentPanel.handleInsertImage(null);
        }

        public void setColorForSelectedCell() {
            parentPanel.handleSetCellColor(null);
        }
    }

    // ==================== ENHANCED FORMULA ENGINE WITH MORE FUNCTIONS ====================
    
    private class EnhancedFormulaEngine {
        private final Pattern CELL_REF_PATTERN = Pattern.compile("([A-Z]+)(\\d+)");
        private final Pattern FUNCTION_PATTERN = Pattern.compile("(\\w+)\\(([^)]+)\\)");
        private final Pattern RANGE_PATTERN = Pattern.compile("([A-Z]+)(\\d+):([A-Z]+)(\\d+)");
        private final Map<String, Integer> colNameCache = new HashMap<>();
        
        public Object evaluate(String formula, OptimizedTableModel model, int currentRow, int currentCol) {
            if (formula == null || formula.length() <= 1) return formula;
            
            String expression = formula.substring(1);
            
            try {
                expression = resolveRanges(expression, model);
                expression = resolveCellReferences(expression, model, currentRow, currentCol);
                expression = resolveFunctions(expression, model, currentRow, currentCol);
                return evaluateArithmetic(expression);
            } catch (Exception e) {
                return "#ERROR: " + e.getMessage();
            }
        }
        
        private String resolveRanges(String expression, OptimizedTableModel model) {
            Matcher matcher = RANGE_PATTERN.matcher(expression);
            StringBuffer result = new StringBuffer();
            
            while (matcher.find()) {
                String startCol = matcher.group(1);
                int startRow = Integer.parseInt(matcher.group(2)) - 1;
                String endCol = matcher.group(3);
                int endRow = Integer.parseInt(matcher.group(4)) - 1;
                
                List<Double> values = new ArrayList<>();
                for (int row = startRow; row <= endRow; row++) {
                    for (int col = columnNameToIndex(startCol); col <= columnNameToIndex(endCol); col++) {
                        Object value = model.getRawValueAt(row, col);
                        if (value instanceof Number) {
                            values.add(((Number) value).doubleValue());
                        }
                    }
                }
                
                // Convert list to comma-separated string
                StringBuilder rangeValues = new StringBuilder();
                for (int i = 0; i < values.size(); i++) {
                    if (i > 0) rangeValues.append(",");
                    rangeValues.append(values.get(i));
                }
                
                matcher.appendReplacement(result, rangeValues.toString());
            }
            matcher.appendTail(result);
            return result.toString();
        }
        
        private String resolveCellReferences(String expression, OptimizedTableModel model, int currentRow, int currentCol) {
            Matcher matcher = CELL_REF_PATTERN.matcher(expression);
            StringBuffer result = new StringBuffer();
            
            while (matcher.find()) {
                String colStr = matcher.group(1);
                String rowStr = matcher.group(2);
                
                int col = columnNameToIndex(colStr);
                int row = Integer.parseInt(rowStr) - 1;
                
                Object cellValue = model.getRawValueAt(row, col);
                String replacement = "0";
                
                if (cellValue instanceof Number) {
                    replacement = cellValue.toString();
                } else if (cellValue != null && !cellValue.toString().isEmpty()) {
                    try {
                        replacement = String.valueOf(Double.parseDouble(cellValue.toString()));
                    } catch (NumberFormatException e) {
                        replacement = "0";
                    }
                }
                
                matcher.appendReplacement(result, replacement);
            }
            matcher.appendTail(result);
            return result.toString();
        }
        
        private String resolveFunctions(String expression, OptimizedTableModel model, int currentRow, int currentCol) {
            Matcher matcher = FUNCTION_PATTERN.matcher(expression);
            StringBuffer result = new StringBuffer();
            
            while (matcher.find()) {
                String function = matcher.group(1).toUpperCase();
                String args = matcher.group(2);
                String replacement = evaluateFunction(function, args, model, currentRow, currentCol);
                matcher.appendReplacement(result, replacement);
            }
            matcher.appendTail(result);
            return result.toString();
        }
        
        private String evaluateFunction(String function, String args, OptimizedTableModel model, int currentRow, int currentCol) {
            try {
                String[] argArray = args.split(",");
                switch (function) {
                    case "SUM":
                        return String.valueOf(sum(argArray));
                    case "AVERAGE":
                        return String.valueOf(average(argArray));
                    case "COUNT":
                        return String.valueOf(count(argArray));
                    case "MAX":
                        return String.valueOf(max(argArray));
                    case "MIN":
                        return String.valueOf(min(argArray));
                    case "MEDIAN":
                        return String.valueOf(median(argArray));
                    case "STDEV":
                        return String.valueOf(stdev(argArray));
                    case "VARIANCE":
                        return String.valueOf(variance(argArray));
                    case "ROUND":
                        return String.valueOf(round(argArray));
                    case "ROUNDUP":
                        return String.valueOf(roundUp(argArray));
                    case "ROUNDDOWN":
                        return String.valueOf(roundDown(argArray));
                    case "IF":
                        return ifFunction(argArray);
                    case "CONCAT":
                        return concat(argArray);
                    case "LEFT":
                        return left(argArray);
                    case "RIGHT":
                        return right(argArray);
                    case "MID":
                        return mid(argArray);
                    case "LEN":
                        return String.valueOf(len(argArray));
                    case "LOWER":
                        return lower(argArray);
                    case "UPPER":
                        return upper(argArray);
                    case "PROPER":
                        return proper(argArray);
                    case "TRIM":
                        return trim(argArray);
                    case "FIND":
                        return String.valueOf(find(argArray));
                    case "REPLACE":
                        return replace(argArray);
                    case "SUBSTITUTE":
                        return substitute(argArray);
                    case "NOW":
                        return String.valueOf(System.currentTimeMillis());
                    case "TODAY":
                        return new SimpleDateFormat("yyyy-MM-dd").format(new Date());
                    case "YEAR":
                        return String.valueOf(year(argArray));
                    case "MONTH":
                        return String.valueOf(month(argArray));
                    case "DAY":
                        return String.valueOf(day(argArray));
                    case "AND":
                        return String.valueOf(and(argArray));
                    case "OR":
                        return String.valueOf(or(argArray));
                    case "NOT":
                        return String.valueOf(not(argArray));
                    default:
                        return "0";
                }
            } catch (Exception e) {
                return "0";
            }
        }
        
        // Enhanced mathematical functions
        private double sum(String[] args) {
            double total = 0;
            for (String arg : args) {
                total += Double.parseDouble(arg.trim());
            }
            return total;
        }
        
        private double average(String[] args) {
            return sum(args) / args.length;
        }
        
        private int count(String[] args) {
            return args.length;
        }
        
        private double max(String[] args) {
            double max = Double.NEGATIVE_INFINITY;
            for (String arg : args) {
                max = Math.max(max, Double.parseDouble(arg.trim()));
            }
            return max;
        }
        
        private double min(String[] args) {
            double min = Double.POSITIVE_INFINITY;
            for (String arg : args) {
                min = Math.min(min, Double.parseDouble(arg.trim()));
            }
            return min;
        }
        
        private double median(String[] args) {
            double[] values = new double[args.length];
            for (int i = 0; i < args.length; i++) {
                values[i] = Double.parseDouble(args[i].trim());
            }
            Arrays.sort(values);
            if (values.length % 2 == 0) {
                return (values[values.length/2 - 1] + values[values.length/2]) / 2.0;
            } else {
                return values[values.length/2];
            }
        }
        
        private double stdev(String[] args) {
            double mean = average(args);
            double sum = 0;
            for (String arg : args) {
                double diff = Double.parseDouble(arg.trim()) - mean;
                sum += diff * diff;
            }
            return Math.sqrt(sum / args.length);
        }
        
        private double variance(String[] args) {
            double stdev = stdev(args);
            return stdev * stdev;
        }
        
        private double round(String[] args) {
            if (args.length >= 1) {
                double value = Double.parseDouble(args[0].trim());
                int decimals = args.length >= 2 ? Integer.parseInt(args[1].trim()) : 0;
                double factor = Math.pow(10, decimals);
                return Math.round(value * factor) / factor;
            }
            return 0;
        }
        
        private double roundUp(String[] args) {
            if (args.length >= 1) {
                double value = Double.parseDouble(args[0].trim());
                int decimals = args.length >= 2 ? Integer.parseInt(args[1].trim()) : 0;
                double factor = Math.pow(10, decimals);
                return Math.ceil(value * factor) / factor;
            }
            return 0;
        }
        
        private double roundDown(String[] args) {
            if (args.length >= 1) {
                double value = Double.parseDouble(args[0].trim());
                int decimals = args.length >= 2 ? Integer.parseInt(args[1].trim()) : 0;
                double factor = Math.pow(10, decimals);
                return Math.floor(value * factor) / factor;
            }
            return 0;
        }
        
        private String ifFunction(String[] args) {
            if (args.length >= 3) {
                boolean condition = Boolean.parseBoolean(args[0].trim());
                return condition ? args[1] : args[2];
            }
            return "FALSE";
        }
        
        private String concat(String[] args) {
            StringBuilder result = new StringBuilder();
            for (String arg : args) {
                result.append(arg);
            }
            return result.toString();
        }
        
        // Additional string functions
        private String left(String[] args) {
            if (args.length >= 1) {
                String text = args[0].trim();
                int length = args.length >= 2 ? Integer.parseInt(args[1].trim()) : 1;
                return text.substring(0, Math.min(length, text.length()));
            }
            return "";
        }
        
        private String right(String[] args) {
            if (args.length >= 1) {
                String text = args[0].trim();
                int length = args.length >= 2 ? Integer.parseInt(args[1].trim()) : 1;
                return text.substring(Math.max(0, text.length() - length));
            }
            return "";
        }
        
        private String mid(String[] args) {
            if (args.length >= 3) {
                String text = args[0].trim();
                int start = Integer.parseInt(args[1].trim()) - 1;
                int length = Integer.parseInt(args[2].trim());
                if (start >= 0 && start < text.length()) {
                    return text.substring(start, Math.min(start + length, text.length()));
                }
            }
            return "";
        }
        
        private int len(String[] args) {
            return args.length > 0 ? args[0].length() : 0;
        }
        
        private String lower(String[] args) {
            return args.length >= 1 ? args[0].toLowerCase() : "";
        }
        
        private String upper(String[] args) {
            return args.length >= 1 ? args[0].toUpperCase() : "";
        }
        
        private String proper(String[] args) {
            if (args.length >= 1) {
                String text = args[0].trim();
                if (text.isEmpty()) return "";
                return text.substring(0, 1).toUpperCase() + text.substring(1).toLowerCase();
            }
            return "";
        }
        
        private String trim(String[] args) {
            return args.length >= 1 ? args[0].trim() : "";
        }
        
        private int find(String[] args) {
            if (args.length >= 2) {
                String findText = args[0].trim();
                String withinText = args[1].trim();
                return withinText.indexOf(findText) + 1; // Excel is 1-based
            }
            return 0;
        }
        
        private String replace(String[] args) {
            if (args.length >= 4) {
                String text = args[0].trim();
                int start = Integer.parseInt(args[1].trim()) - 1;
                int length = Integer.parseInt(args[2].trim());
                String newText = args[3].trim();
                
                if (start >= 0 && start <= text.length()) {
                    String before = text.substring(0, start);
                    String after = start + length < text.length() ? text.substring(start + length) : "";
                    return before + newText + after;
                }
            }
            return "";
        }
        
        private String substitute(String[] args) {
            if (args.length >= 3) {
                String text = args[0].trim();
                String oldText = args[1].trim();
                String newText = args[2].trim();
                return text.replace(oldText, newText);
            }
            return "";
        }
        
        // Date functions
        private int year(String[] args) {
            if (args.length >= 1) {
                try {
                    Date date = new SimpleDateFormat("yyyy-MM-dd").parse(args[0].trim());
                    return date.getYear() + 1900;
                } catch (Exception e) {
                    // Try to parse as timestamp
                    try {
                        long timestamp = Long.parseLong(args[0].trim());
                        Date date = new Date(timestamp);
                        return date.getYear() + 1900;
                    } catch (Exception ex) {
                        return 0;
                    }
                }
            }
            return 0;
        }
        
        private int month(String[] args) {
            if (args.length >= 1) {
                try {
                    Date date = new SimpleDateFormat("yyyy-MM-dd").parse(args[0].trim());
                    return date.getMonth() + 1;
                } catch (Exception e) {
                    try {
                        long timestamp = Long.parseLong(args[0].trim());
                        Date date = new Date(timestamp);
                        return date.getMonth() + 1;
                    } catch (Exception ex) {
                        return 0;
                    }
                }
            }
            return 0;
        }
        
        private int day(String[] args) {
            if (args.length >= 1) {
                try {
                    Date date = new SimpleDateFormat("yyyy-MM-dd").parse(args[0].trim());
                    return date.getDate();
                } catch (Exception e) {
                    try {
                        long timestamp = Long.parseLong(args[0].trim());
                        Date date = new Date(timestamp);
                        return date.getDate();
                    } catch (Exception ex) {
                        return 0;
                    }
                }
            }
            return 0;
        }
        
        // Logical functions
        private boolean and(String[] args) {
            for (String arg : args) {
                if (!Boolean.parseBoolean(arg.trim())) {
                    return false;
                }
            }
            return true;
        }
        
        private boolean or(String[] args) {
            for (String arg : args) {
                if (Boolean.parseBoolean(arg.trim())) {
                    return true;
                }
            }
            return false;
        }
        
        private boolean not(String[] args) {
            return args.length >= 1 && !Boolean.parseBoolean(args[0].trim());
        }
        
        private Object evaluateArithmetic(String expression) {
            try {
                return new javax.script.ScriptEngineManager()
                    .getEngineByName("JavaScript")
                    .eval(expression);
            } catch (Exception e) {
                return "#ERROR";
            }
        }
        
        private int columnNameToIndex(String colName) {
            return colNameCache.computeIfAbsent(colName, name -> {
                int index = 0;
                for (int i = 0; i < name.length(); i++) {
                    index = index * 26 + (name.charAt(i) - 'A' + 1);
                }
                return index - 1;
            });
        }
    }

    // ==================== ENHANCED UI WITH FORMATTING TOOLBAR ====================
    
    private void setupEnhancedUI() {
        JPanel formulaPanel = new JPanel(new BorderLayout());
        formulaPanel.add(new JLabel("fx:"), BorderLayout.WEST);
        formulaPanel.add(formulaBar, BorderLayout.CENTER);
        formulaBar.setEditable(true);
        formulaBar.addActionListener(e -> applyFormulaFromBar());
        
        JToolBar mainToolBar = new JToolBar();
        JToolBar formatToolBar = new JToolBar();
        JToolBar alignmentToolBar = new JToolBar();
        JToolBar borderToolBar = new JToolBar();
        
        // Main operations
        mainToolBar.add(createButton("Add Row", "Insert new row", e -> insertRow()));
        mainToolBar.add(createButton("Add Column", "Insert new column", e -> insertColumn()));
        mainToolBar.add(createButton("Delete Row", "Delete selected row", e -> deleteRow()));
        mainToolBar.add(createButton("Delete Column", "Delete selected column", e -> deleteColumn()));
        mainToolBar.addSeparator();
        mainToolBar.add(createButton("Sort Asc", "Sort ascending", e -> sortAscending()));
        mainToolBar.add(createButton("Sort Desc", "Sort descending", e -> sortDescending()));
        mainToolBar.addSeparator();
        mainToolBar.add(createButton("Freeze Panes", "Freeze panes", e -> toggleFreezePanes()));
        mainToolBar.add(createButton("Protect Cell", "Protect/unprotect cell", e -> toggleCellProtection()));
        mainToolBar.addSeparator();
        mainToolBar.add(createButton("Select All", "Select all cells (Ctrl+A)", e -> selectAllCells()));
        mainToolBar.add(createButton("Clear All", "Clear all cells", e -> clearAllCells()));
        mainToolBar.addSeparator();
        mainToolBar.add(createButton("AutoFit", "Auto-fit columns", e -> autoFitColumns()));
        mainToolBar.add(createButton("Insert Function", "Insert function", e -> insertFunction()));
        mainToolBar.add(createButton("Cell Properties", "Show cell properties", e -> showCellProperties()));
        
        // Formatting toolbar - Font
        formatToolBar.add(new JLabel("Font:"));
        formatToolBar.add(fontCombo);
        formatToolBar.add(new JLabel("Size:"));
        formatToolBar.add(fontSizeSpinner);
        formatToolBar.add(boldCheckBox);
        formatToolBar.add(italicCheckBox);
        formatToolBar.add(underlineCheckBox);
        formatToolBar.add(createButton("A", "Text color", e -> setTextColor()));
        formatToolBar.add(createButton("BG", "Background color", e -> setCellColor()));
        formatToolBar.addSeparator();
        formatToolBar.add(createButton("Copy Format", "Copy format (Ctrl+Shift+C)", e -> copyFormat()));
        formatToolBar.add(createButton("Paste Format", "Paste format (Ctrl+Shift+V)", e -> pasteFormat()));
        formatToolBar.add(createButton("Clear Format", "Clear formatting", e -> clearFormat()));
        
        formatToolBar.addSeparator();
        formatToolBar.add(new JLabel("Format:"));
        formatToolBar.add(numberFormatCombo);
        numberFormatCombo.addActionListener(e -> applyNumberFormat());
        
        // Alignment toolbar
        alignmentToolBar.add(new JLabel("Horizontal:"));
        alignmentToolBar.add(horizontalAlignmentCombo);
        alignmentToolBar.add(new JLabel("Vertical:"));
        alignmentToolBar.add(verticalAlignmentCombo);
        alignmentToolBar.add(createButton("Apply Align", "Apply alignment", e -> applyAlignment()));
        
        // Border toolbar
        borderToolBar.add(new JLabel("Border:"));
        borderToolBar.add(borderStyleCombo);
        borderToolBar.add(createButton("Apply Border", "Apply border", e -> applyBorder()));
        borderToolBar.add(createButton("Clear Border", "Clear all borders", e -> clearBorders()));
        
        // Add action listeners
        fontCombo.addActionListener(e -> applyFontStyle());
        fontSizeSpinner.addChangeListener(e -> applyFontStyle());
        boldCheckBox.addActionListener(e -> applyFontStyle());
        italicCheckBox.addActionListener(e -> applyFontStyle());
        underlineCheckBox.addActionListener(e -> applyUnderline());
        horizontalAlignmentCombo.addActionListener(e -> applyHorizontalAlignment());
        verticalAlignmentCombo.addActionListener(e -> applyVerticalAlignment());
        
        JPanel searchPanel = createSearchPanel();
        
        statusBar.setBorder(BorderFactory.createEtchedBorder());
        
        JPanel topPanel = new JPanel(new BorderLayout());
        topPanel.add(formulaPanel, BorderLayout.NORTH);
        
        JPanel toolBarPanel = new JPanel();
        toolBarPanel.setLayout(new BoxLayout(toolBarPanel, BoxLayout.Y_AXIS));
        toolBarPanel.add(mainToolBar);
        toolBarPanel.add(formatToolBar);
        toolBarPanel.add(alignmentToolBar);
        toolBarPanel.add(borderToolBar);
        
        topPanel.add(toolBarPanel, BorderLayout.CENTER);
        topPanel.add(searchPanel, BorderLayout.SOUTH);
        
        add(topPanel, BorderLayout.NORTH);
        add(statusBar, BorderLayout.SOUTH);
    }

    private void applyAlignment() {
        applyHorizontalAlignment();
        applyVerticalAlignment();
    }

    private void applyUnderline() {
        applyToSelectedCells((row, col) -> {
            model.setUnderline(row, col, underlineCheckBox.isSelected());
        });
    }

    private void clearBorders() {
        applyToSelectedCells((row, col) -> {
            model.setBorder(row, col, "None");
        });
    }

    private void clearAllCells() {
        int result = JOptionPane.showConfirmDialog(this,
            "Are you sure you want to clear all cells? This cannot be undone.",
            "Clear All Cells",
            JOptionPane.YES_NO_OPTION);
            
        if (result == JOptionPane.YES_OPTION) {
            model.batchUpdate(() -> {
                for (int row = 0; row < model.getRowCount(); row++) {
                    for (int col = 0; col < model.getColumnCount(); col++) {
                        model.setValueAt("", row, col);
                    }
                }
            });
            statusBar.setText("All cells cleared");
        }
    }

    // ==================== ADDITIONAL EXCEL-LIKE FEATURES ====================

    public void autoFitColumns() {
        for (int col = 0; col < table.getColumnCount(); col++) {
            int maxWidth = 0;
            for (int row = 0; row < Math.min(100, table.getRowCount()); row++) {
                TableCellRenderer renderer = table.getCellRenderer(row, col);
                Component comp = table.prepareRenderer(renderer, row, col);
                maxWidth = Math.max(comp.getPreferredSize().width, maxWidth);
            }
            table.getColumnModel().getColumn(col).setPreferredWidth(maxWidth + 10);
        }
        statusBar.setText("Columns auto-fitted");
    }

    public void insertFunction() {
        String[] functions = {
            "SUM", "AVERAGE", "COUNT", "MAX", "MIN", "MEDIAN", "STDEV",
            "IF", "CONCAT", "LEFT", "RIGHT", "MID", "LEN", "LOWER", "UPPER",
            "NOW", "TODAY", "YEAR", "MONTH", "DAY", "AND", "OR", "NOT"
        };
        
        String selectedFunction = (String) JOptionPane.showInputDialog(this,
            "Select a function:",
            "Insert Function",
            JOptionPane.QUESTION_MESSAGE,
            null,
            functions,
            functions[0]);
            
        if (selectedFunction != null) {
            formulaBar.setText("=" + selectedFunction + "()");
            formulaBar.requestFocus();
            formulaBar.setCaretPosition(formulaBar.getText().length() - 1);
        }
    }

    public void showCellProperties() {
        int row = table.getSelectedRow();
        int col = table.getSelectedColumn();
        
        if (row != -1 && col != -1) {
            int modelRow = table.convertRowIndexToModel(row);
            int modelCol = table.convertColumnIndexToModel(col);
            
            StringBuilder properties = new StringBuilder();
            properties.append("Cell: ").append(calculateColumnName(col)).append(row + 1).append("\n");
            properties.append("Value: ").append(model.getRawValueAt(modelRow, modelCol)).append("\n");
            properties.append("Formula: ").append(model.getCellFormula(modelRow, modelCol)).append("\n");
            properties.append("Format: ").append(model.getCellFormat(modelRow, modelCol)).append("\n");
            properties.append("Font: ").append(model.getCellFont(modelRow, modelCol)).append("\n");
            properties.append("Background: ").append(model.getCellColor(modelRow, modelCol)).append("\n");
            properties.append("Text Color: ").append(model.getTextColor(modelRow, modelCol)).append("\n");
            properties.append("Protected: ").append(model.isCellProtected(modelRow, modelCol)).append("\n");
            
            JOptionPane.showMessageDialog(this, properties.toString(), "Cell Properties", JOptionPane.INFORMATION_MESSAGE);
        }
    }

    // ==================== CELL FORMAT CLASS ====================

    private static class CellFormat {
        public Color backgroundColor;
        public Color textColor;
        public Font font;
        public boolean isBold;
        public boolean isItalic;
        public boolean isUnderline;
        public String horizontalAlignment;
        public String verticalAlignment;
        public String cellFormat;
        public String borderStyle;
        
        public CellFormat() {
            // Default constructor
        }
    }

    // ==================== BORDER INFO CLASS ====================

    private static class BorderInfo {
        public String topStyle = "None";
        public String bottomStyle = "None";
        public String leftStyle = "None";
        public String rightStyle = "None";
        public Color color = Color.BLACK;
        
        public BorderInfo() {}
    }

    // ==================== THEME SUPPORT ====================

    /**
     * Update the panel colors when theme changes
     */
    public void updateThemeColors(boolean isDarkTheme) {
        Color backgroundColor = isDarkTheme ? new Color(45, 45, 45) : Color.WHITE;
        Color foregroundColor = isDarkTheme ? Color.WHITE : Color.BLACK;
        Color gridColor = isDarkTheme ? new Color(80, 80, 80) : Color.GRAY;
        Color selectionColor = isDarkTheme ? new Color(0, 60, 120) : new Color(184, 207, 229);
        Color alternateColor = isDarkTheme ? new Color(60, 60, 60) : new Color(240, 240, 240);
        
        // Update table colors
        table.setBackground(backgroundColor);
        table.setForeground(foregroundColor);
        table.setGridColor(gridColor);
        table.setSelectionBackground(selectionColor);
        table.setSelectionForeground(foregroundColor);
        
        // Update alternate row colors
        alternateRowColors = true;
        
        // Update other UI components
        formulaBar.setBackground(isDarkTheme ? new Color(60, 60, 60) : Color.WHITE);
        formulaBar.setForeground(foregroundColor);
        formulaBar.setCaretColor(foregroundColor);
        
        statusBar.setBackground(isDarkTheme ? new Color(60, 60, 60) : new Color(240, 240, 240));
        statusBar.setForeground(foregroundColor);
        
        // Update toolbars and panels
        updateComponentTreeUI(this, isDarkTheme);
        
        // Force repaint
        table.repaint();
        repaint();
    }

    /**
     * Recursively update all components in the panel
     */
    private void updateComponentTreeUI(Component component, boolean isDarkTheme) {
        Color backgroundColor = isDarkTheme ? new Color(45, 45, 45) : Color.WHITE;
        Color foregroundColor = isDarkTheme ? Color.WHITE : Color.BLACK;
        
        if (component instanceof JPanel) {
            JPanel panel = (JPanel) component;
            panel.setBackground(backgroundColor);
            panel.setForeground(foregroundColor);
            
            // Update all child components
            for (Component child : panel.getComponents()) {
                updateComponentTreeUI(child, isDarkTheme);
            }
        } else if (component instanceof JToolBar) {
            JToolBar toolbar = (JToolBar) component;
            toolbar.setBackground(backgroundColor);
            toolbar.setForeground(foregroundColor);
        } else if (component instanceof JLabel) {
            JLabel label = (JLabel) component;
            label.setForeground(foregroundColor);
        } else if (component instanceof JComboBox) {
            JComboBox<?> combo = (JComboBox<?>) component;
            combo.setBackground(isDarkTheme ? new Color(60, 60, 60) : Color.WHITE);
            combo.setForeground(foregroundColor);
        } else if (component instanceof JSpinner) {
            JSpinner spinner = (JSpinner) component;
            spinner.setBackground(isDarkTheme ? new Color(60, 60, 60) : Color.WHITE);
            spinner.setForeground(foregroundColor);
        } else if (component instanceof JCheckBox) {
            JCheckBox checkBox = (JCheckBox) component;
            checkBox.setBackground(backgroundColor);
            checkBox.setForeground(foregroundColor);
        } else if (component instanceof JTextField) {
            JTextField textField = (JTextField) component;
            textField.setBackground(isDarkTheme ? new Color(60, 60, 60) : Color.WHITE);
            textField.setForeground(foregroundColor);
            textField.setCaretColor(foregroundColor);
        }
    }

    /**
     * Get current theme state from parent frame
     */
    private boolean isDarkTheme() {
        Container parent = getParent();
        while (parent != null) {
            if (parent instanceof java.awt.Window) {
                return parent.getBackground().getRed() < 128; // Dark theme has dark background
            }
            parent = parent.getParent();
        }
        return false;
    }

    // ==================== TABLE SETUP ====================

    private void setupTable() {
        table.setRowHeight(24);
        table.setFillsViewportHeight(true);
        table.setFont(new Font("Segoe UI", Font.PLAIN, 14));
        table.setShowGrid(showGridLines);
        table.setGridColor(Color.GRAY);
        table.setIntercellSpacing(new Dimension(1, 1));
        table.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);
        
        sorter = new TableRowSorter<>(model);
        table.setRowSorter(sorter);
        
        JScrollPane scroll = new JScrollPane(table);
        add(scroll, BorderLayout.CENTER);
        
        model.addTableModelListener(e -> {
            if (autoExpandEnabled && tableMode) {
                checkAutoExpand();
            }
            if (showTotalRow) {
                updateTotalRowPanel();
            }
        });
    }

    // ==================== SEARCH PANEL ====================

    private JPanel createSearchPanel() {
        JPanel searchPanel = new JPanel();
        JTextField searchField = new JTextField(20);
        searchPanel.add(new JLabel("Search: "));
        searchPanel.add(searchField);
        
        JButton clearSearchButton = new JButton("Clear");
        clearSearchButton.addActionListener(e -> searchField.setText(""));
        searchPanel.add(clearSearchButton);
        
        JButton replaceButton = new JButton("Replace");
        replaceButton.addActionListener(e -> showReplaceDialog());
        searchPanel.add(replaceButton);
        
        searchField.getDocument().addDocumentListener(new DocumentListener() {
            public void insertUpdate(DocumentEvent e) { filterTable(searchField.getText()); }
            public void removeUpdate(DocumentEvent e) { filterTable(searchField.getText()); }
            public void changedUpdate(DocumentEvent e) { filterTable(searchField.getText()); }
            private void filterTable(String text) {
                if (text.trim().isEmpty()) {
                    sorter.setRowFilter(null);
                } else {
                    sorter.setRowFilter(RowFilter.regexFilter("(?i)" + text));
                }
            }
        });
        
        return searchPanel;
    }

    // ==================== EXISTING METHODS ====================

    private void insertRow() {
        int selectedRow = table.getSelectedRow();
        int insertAt = (selectedRow != -1) ? table.convertRowIndexToModel(selectedRow) : model.getRowCount();
        model.insertRow(insertAt);
        addToHistory("INSERT_ROW", insertAt, -1, null);
    }

    private void insertColumn() {
        int selectedCol = table.getSelectedColumn();
        int insertAt = (selectedCol != -1) ? table.convertColumnIndexToModel(selectedCol) : model.getColumnCount();
        model.insertColumn(insertAt);
        addToHistory("INSERT_COL", -1, insertAt, null);
    }

    private void deleteRow() {
        int selectedRow = table.getSelectedRow();
        if (selectedRow != -1) {
            int modelRow = table.convertRowIndexToModel(selectedRow);
            model.deleteRow(modelRow);
            addToHistory("DELETE_ROW", modelRow, -1, null);
        } else {
            JOptionPane.showMessageDialog(this, "Please select a row to delete.");
        }
    }

    private void deleteColumn() {
        int selectedCol = table.getSelectedColumn();
        if (selectedCol != -1) {
            int modelCol = table.convertColumnIndexToModel(selectedCol);
            model.deleteColumn(modelCol);
            addToHistory("DELETE_COL", -1, modelCol, null);
        } else {
            JOptionPane.showMessageDialog(this, "Please select a column to delete.");
        }
    }

    private void sortAscending() {
        int col = table.getSelectedColumn();
        if (col != -1) {
            sorter.setSortKeys(java.util.List.of(new RowSorter.SortKey(col, SortOrder.ASCENDING)));
        }
    }

    private void sortDescending() {
        int col = table.getSelectedColumn();
        if (col != -1) {
            sorter.setSortKeys(java.util.List.of(new RowSorter.SortKey(col, SortOrder.DESCENDING)));
        }
    }

    private void toggleFreezePanes() {
        freezePanes = !freezePanes;
        JOptionPane.showMessageDialog(this, freezePanes ? "Panes frozen" : "Panes unfrozen");
    }

    private void toggleCellProtection() {
        int row = table.getSelectedRow();
        int col = table.getSelectedColumn();
        if (row != -1 && col != -1) {
            int modelRow = table.convertRowIndexToModel(row);
            int modelCol = table.convertColumnIndexToModel(col);
            boolean currentlyProtected = model.isCellProtected(modelRow, modelCol);
            model.setCellProtected(modelRow, modelCol, !currentlyProtected);
            JOptionPane.showMessageDialog(this, 
                !currentlyProtected ? "Cell protected" : "Cell unprotected");
        }
    }

    private void applyFontStyle() {
        int row = table.getSelectedRow();
        int col = table.getSelectedColumn();
        if (row != -1 && col != -1) {
            int modelRow = table.convertRowIndexToModel(row);
            int modelCol = table.convertColumnIndexToModel(col);
            
            String fontName = (String) fontCombo.getSelectedItem();
            int fontSize = (Integer) fontSizeSpinner.getValue();
            int fontStyle = Font.PLAIN;
            if (boldCheckBox.isSelected()) fontStyle |= Font.BOLD;
            if (italicCheckBox.isSelected()) fontStyle |= Font.ITALIC;
            
            Font newFont = new Font(fontName, fontStyle, fontSize);
            model.setCellFont(modelRow, modelCol, newFont);
            model.setBold(modelRow, modelCol, boldCheckBox.isSelected());
            model.setItalic(modelRow, modelCol, italicCheckBox.isSelected());
        }
    }

    private void setTextColor() {
        int row = table.getSelectedRow();
        int col = table.getSelectedColumn();
        if (row != -1 && col != -1) {
            Color color = JColorChooser.showDialog(this, "Choose Text Color", Color.BLACK);
            if (color != null) {
                int modelRow = table.convertRowIndexToModel(row);
                int modelCol = table.convertColumnIndexToModel(col);
                model.setTextColor(modelRow, modelCol, color);
            }
        }
    }

    private void setCellColor() {
        int row = table.getSelectedRow();
        int col = table.getSelectedColumn();
        if (row != -1 && col != -1) {
            Color color = JColorChooser.showDialog(this, "Choose Cell Color", Color.WHITE);
            if (color != null) {
                int modelRow = table.convertRowIndexToModel(row);
                int modelCol = table.convertColumnIndexToModel(col);
                model.setCellColor(modelRow, modelCol, color);
            }
        }
    }

    private void setupAutoSave() {
        if (autoSaveEnabled) {
            autoSaveTimer = new Timer(30000, e -> autoSave()); // Auto-save every 30 seconds
            autoSaveTimer.start();
        }
    }

    private void autoSave() {
        if (file != null && dirtyCells.size() > 0) {
            handleSave(new ActionEvent(this, ActionEvent.ACTION_PERFORMED, "autosave"));
            statusBar.setText("Auto-saved at " + new SimpleDateFormat("HH:mm:ss").format(new Date()));
        }
    }

    private void addToHistory(String action, int row, int col, Object value) {
        cellHistory.add(action + ":" + row + ":" + col + ":" + (value != null ? value.toString() : ""));
        historyPointer = cellHistory.size() - 1;
    }

    private void updateCellHistory(int row, int col) {
        if (cellTrackingEnabled) {
            statusBar.setText("Cell: " + calculateColumnName(col) + (row + 1) + 
                " | History: " + cellHistory.size() + " actions");
        }
    }

    // ==================== ENHANCED FORMATTERS ====================
    
    private DecimalFormat getNumberFormat() {
        if (numberFormat == null) {
            numberFormat = new DecimalFormat("#,##0.00");
        }
        return numberFormat;
    }
    
    private DecimalFormat getCurrencyFormat() {
        if (currencyFormat == null) {
            currencyFormat = new DecimalFormat("$#,##0.00");
        }
        return currencyFormat;
    }
    
    private DecimalFormat getPercentFormat() {
        if (percentFormat == null) {
            percentFormat = new DecimalFormat("0.00%");
        }
        return percentFormat;
    }
    
    private DecimalFormat getScientificFormat() {
        if (scientificFormat == null) {
            scientificFormat = new DecimalFormat("0.###E0");
        }
        return scientificFormat;
    }
    
    private DecimalFormat getAccountingFormat() {
        if (accountingFormat == null) {
            accountingFormat = new DecimalFormat("$#,##0.00;($#,##0.00)");
        }
        return accountingFormat;
    }
    
    private SimpleDateFormat getDateFormat() {
        if (dateFormat == null) {
            dateFormat = new SimpleDateFormat("yyyy-MM-dd");
        }
        return dateFormat;
    }

    // ==================== UTILITY METHODS ====================
    
    public int columnNameToIndex(String colName) {
        int index = 0;
        for (int i = 0; i < colName.length(); i++) {
            index = index * 26 + (colName.charAt(i) - 'A' + 1);
        }
        return index - 1;
    }

    public String calculateColumnName(int col) {
        StringBuilder name = new StringBuilder();
        int c = col;
        do {
            name.insert(0, (char) ('A' + (c % 26)));
            c = c / 26 - 1;
        } while (c >= 0);
        return name.toString();
    }
    
    public String getColumnName(int col) {
        return calculateColumnName(col);
    }

    private JButton createButton(String text, String tooltip, java.awt.event.ActionListener listener) {
        JButton button = new JButton(text);
        button.setToolTipText(tooltip);
        button.addActionListener(listener);
        button.setFocusable(false);
        return button;
    }

    private void applyFormulaFromBar() {
        int row = table.getSelectedRow();
        int col = table.getSelectedColumn();
        if (row != -1 && col != -1) {
            int modelRow = table.convertRowIndexToModel(row);
            int modelCol = table.convertColumnIndexToModel(col);
            model.setValueAt(formulaBar.getText(), modelRow, modelCol);
        }
    }
    
    private void updateFormulaBar() {
        int row = table.getSelectedRow();
        int col = table.getSelectedColumn();
        if (row != -1 && col != -1) {
            int modelRow = table.convertRowIndexToModel(row);
            int modelCol = table.convertColumnIndexToModel(col);
            String formula = model.getCellFormula(modelRow, modelCol);
            formulaBar.setText(formula != null ? formula : "");
        } else {
            formulaBar.setText("");
        }
    }
    
    private void updateStatusBar() {
        int row = table.getSelectedRow();
        int col = table.getSelectedColumn();
        if (row != -1 && col != -1) {
            statusBar.setText("Cell: " + calculateColumnName(col) + (row + 1));
        } else {
            statusBar.setText("Ready");
        }
    }
    
    private void applyNumberFormat() {
        int row = table.getSelectedRow();
        int col = table.getSelectedColumn();
        if (row != -1 && col != -1) {
            int modelRow = table.convertRowIndexToModel(row);
            int modelCol = table.convertColumnIndexToModel(col);
            String format = (String) numberFormatCombo.getSelectedItem();
            model.setCellFormat(modelRow, modelCol, format);
        }
    }

    private void scheduleDelayedCalculation() {
        if (delayedCalculationTimer != null) {
            delayedCalculationTimer.stop();
        }
        
        delayedCalculationTimer = new Timer(500, e -> recalculateDirtyCells());
        delayedCalculationTimer.setRepeats(false);
        delayedCalculationTimer.start();
    }
    
    private void recalculateDirtyCells() {
        calculationCache.clear();
        for (String cellKey : dirtyCells) {
            String[] parts = cellKey.split("_");
            int row = Integer.parseInt(parts[0]);
            int col = Integer.parseInt(parts[1]);
            model.fireTableCellUpdated(row, col);
        }
        dirtyCells.clear();
    }

    public void handleAddRow() {
        model.batchUpdate(() -> {
            model.fireTableRowsInserted(model.getRowCount(), model.getRowCount());
        });
    }

    public void handleDeleteRow() {
        int selectedRow = table.getSelectedRow();
        if (selectedRow != -1) {
            int modelRow = table.convertRowIndexToModel(selectedRow);
            model.batchUpdate(() -> {
                for (int col = 0; col < model.getColumnCount(); col++) {
                    model.setValueAt("", modelRow, col);
                }
            });
        } else {
            JOptionPane.showMessageDialog(this, "Please select a row to delete.");
        }
    }

    public void handleExportCSV() {
        JFileChooser chooser = new JFileChooser();
        chooser.setDialogTitle("Export as CSV");
        if (file != null) {
            chooser.setSelectedFile(new File(file.getName().replaceFirst("[.][^.]+$", "") + ".csv"));
        } else {
            chooser.setSelectedFile(new File("spreadsheet.csv"));
        }
        int result = chooser.showSaveDialog(this);
        if (result == JFileChooser.APPROVE_OPTION) {
            File csvFile = chooser.getSelectedFile();
            try (BufferedWriter writer = new BufferedWriter(new FileWriter(csvFile))) {
                for (int col = 0; col < model.getColumnCount(); col++) {
                    writer.write(model.getColumnName(col));
                    if (col < model.getColumnCount() - 1) writer.write(",");
                }
                writer.newLine();
                
                for (int row = 0; row < model.getRowCount(); row++) {
                    for (int col = 0; col < model.getColumnCount(); col++) {
                        Object value = model.getRawValueAt(row, col);
                        writer.write(value != null ? value.toString() : "");
                        if (col < model.getColumnCount() - 1) writer.write(",");
                    }
                    writer.newLine();
                }
                JOptionPane.showMessageDialog(this, "CSV exported successfully!");
            } catch (IOException ex) {
                JOptionPane.showMessageDialog(this, "CSV export failed:\n" + ex.getMessage());
            }
        }
    }

    public void handleSetCellImage(int row, int col, ImageIcon image) {
        if (row != -1 && col != -1) {
            int modelRow = table.convertRowIndexToModel(row);
            int modelCol = table.convertColumnIndexToModel(col);
            
            JFileChooser fileChooser = new JFileChooser();
            fileChooser.setFileFilter(new FileNameExtensionFilter("Image Files", "png", "jpg", "jpeg", "gif"));
            int result = fileChooser.showOpenDialog(this);
            if (result == JFileChooser.APPROVE_OPTION) {
                File imageFile = fileChooser.getSelectedFile();
                ImageIcon icon = new ImageIcon(imageFile.getAbsolutePath());
                Image scaled = icon.getImage().getScaledInstance(50, 50, Image.SCALE_SMOOTH);
                ImageIcon scaledIcon = new ImageIcon(scaled);
                model.setValueAt(scaledIcon, modelRow, modelCol);
            }
        } else {
            JOptionPane.showMessageDialog(this, "Please select a cell to insert an image.");
        }
    }

    public void handleSetCellColor(ActionEvent e) {
        setCellColor();
    }

    public void handleInsertImage(ActionEvent e) {
        int selectedRow = table.getSelectedRow();
        int selectedCol = table.getSelectedColumn();
        handleSetCellImage(selectedRow, selectedCol, null);
    }

    private void loadXlsxContent() {
        if (file == null) return;
        
        try (FileInputStream fis = new FileInputStream(file);
             Workbook workbook = new XSSFWorkbook(fis)) {
            Sheet sheet = workbook.getSheetAt(0);
            int maxCols = 0;
            
            for (Row row : sheet) {
                if (row != null && row.getLastCellNum() > maxCols) {
                    maxCols = row.getLastCellNum();
                }
            }
            
            List<Object[]> rows = new ArrayList<>();
            for (Row row : sheet) {
                if (row == null) continue;
                Object[] rowData = new Object[maxCols];
                for (int i = 0; i < maxCols; i++) {
                    Cell cell = row.getCell(i);
                    rowData[i] = (cell != null) ? getCellValue(cell) : "";
                }
                rows.add(rowData);
            }
            
            Object[][] dataArray = rows.toArray(new Object[0][]);
            model.loadData(dataArray, null);
            
        } catch (IOException e) {
            JOptionPane.showMessageDialog(this, "Failed to load .xlsx file:\n" + e.getMessage());
        }
    }

    private Object getCellValue(Cell cell) {
        switch (cell.getCellType()) {
            case STRING:    return cell.getStringCellValue();
            case NUMERIC:   return cell.getNumericCellValue();
            case BOOLEAN:   return cell.getBooleanCellValue();
            case FORMULA:   return cell.getCellFormula();
            default:        return "";
        }
    }

    private void autoResizeColumns() {
        int maxRows = Math.min(100, table.getRowCount());
        for (int col = 0; col < table.getColumnCount(); col++) {
            int width = 50;
            for (int row = 0; row < maxRows; row++) {
                TableCellRenderer renderer = table.getCellRenderer(row, col);
                Component comp = table.prepareRenderer(renderer, row, col);
                width = Math.max(comp.getPreferredSize().width + 10, width);
            }
            table.getColumnModel().getColumn(col).setPreferredWidth(width);
        }
    }

    public void handleSave(ActionEvent e) {
        if (file == null) {
            JOptionPane.showMessageDialog(this, "Cannot save - no file specified.");
            return;
        }
        
        try (Workbook workbook = new XSSFWorkbook();
             FileOutputStream fos = new FileOutputStream(file)) {
            Sheet sheet = workbook.createSheet("Sheet1");
            
            for (int row = 0; row < model.getRowCount(); row++) {
                Row xssfRow = sheet.createRow(row);
                for (int col = 0; col < model.getColumnCount(); col++) {
                    Cell cell = xssfRow.createCell(col);
                    Object value = model.getRawValueAt(row, col);
                    
                    if (value instanceof Number) {
                        cell.setCellValue(((Number) value).doubleValue());
                    } else if (model.hasImage(row, col)) {
                        cell.setCellValue("[Image]");
                    } else {
                        cell.setCellValue(value != null ? value.toString() : "");
                    }
                }
            }
            workbook.write(fos);
            JOptionPane.showMessageDialog(this, "✅ .xlsx saved successfully!");
        } catch (IOException ex) {
            JOptionPane.showMessageDialog(this, "❌ Save failed:\n" + ex.getMessage());
        }
    }

    private void checkAutoExpand() {
        int lastRow = model.getRowCount() - 1;
        if (lastRow >= 0) {
            boolean nonEmpty = false;
            for (int col = 0; col < model.getColumnCount(); col++) {
                Object value = model.getRawValueAt(lastRow, col);
                if (value != null && !value.toString().trim().isEmpty()) {
                    nonEmpty = true;
                    break;
                }
            }
            if (nonEmpty) {
                model.fireTableRowsInserted(model.getRowCount(), model.getRowCount());
            }
        }
    }

    private void updateTotalRowPanel() {
        totalRowPanel.removeAll();
        totalRowPanel.setLayout(new BoxLayout(totalRowPanel, BoxLayout.X_AXIS));
        for (int col = 0; col < model.getColumnCount(); col++) {
            double sum = 0.0;
            boolean numeric = false;
            for (int row = 0; row < model.getRowCount(); row++) {
                Object value = model.getRawValueAt(row, col);
                if (value instanceof Number) {
                    sum += ((Number) value).doubleValue();
                    numeric = true;
                }
            }
            String totalText = numeric ? String.format("%.2f", sum) : "";
            JLabel label = new JLabel(totalText);
            label.setPreferredSize(new Dimension(80, 20));
            totalRowPanel.add(label);
        }
        totalRowPanel.revalidate();
        totalRowPanel.repaint();
    }

    private void showReplaceDialog() {
        String find = JOptionPane.showInputDialog(this, "Find:");
        if (find != null && !find.isEmpty()) {
            String replace = JOptionPane.showInputDialog(this, "Replace with:");
            if (replace != null) {
                findAndReplace(find, replace);
            }
        }
    }

    private void findAndReplace(String find, String replace) {
        model.batchUpdate(() -> {
            for (int row = 0; row < model.getRowCount(); row++) {
                for (int col = 0; col < model.getColumnCount(); col++) {
                    Object value = model.getRawValueAt(row, col);
                    if (value != null && value.toString().contains(find)) {
                        String newValue = value.toString().replace(find, replace);
                        model.setValueAt(newValue, row, col);
                    }
                }
            }
        });
    }

    public static Icon getTabIcon() {
        try {
            ImageIcon originalIcon = new ImageIcon(XlsxEditorPanel.class.getResource("/icons/excel.png"));
            Image scaledImage = originalIcon.getImage().getScaledInstance(16, 16, Image.SCALE_SMOOTH);
            return new ImageIcon(scaledImage);
        } catch (Exception e) {
            // Return a default icon if the image is not found
            return new ImageIcon();
        }
    }

    // ==================== PUBLIC ACCESSOR FOR WRAPPER ====================
    
    public OptimizedTableModel getInternalModel() {
        return model;
    }

    public JTable getTable() {
        return table;
    }
}
package com.example.notepadpp.media;
import java.awt.BorderLayout;
import javax.swing.BoxLayout;
import java.awt.Color;
import java.awt.Component;
import java.awt.Dimension;
import java.awt.Font;
import java.awt.Image;
import java.awt.event.ActionEvent;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import javax.swing.table.DefaultTableModel;
import javax.swing.Box;
import javax.swing.Icon;
import javax.swing.ImageIcon;
import javax.swing.JButton;
import javax.swing.JColorChooser;
import javax.swing.JFileChooser;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JScrollPane;
import javax.swing.JTable;
import javax.swing.JTextField;
import javax.swing.RowFilter;
import javax.swing.UIManager;
import javax.swing.event.DocumentEvent;
import javax.swing.event.DocumentListener;
import javax.swing.event.TableModelEvent;
import javax.swing.event.TableModelListener;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.TableCellRenderer;
import javax.swing.table.TableRowSorter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
public class XlsEditorPanel extends JPanel {
    private final File file;
    private final JTable table;
    private final DefaultTableModel model;
    private TableRowSorter<DefaultTableModel> sorter;
    private final Map<String, Color> cellColorMap = new HashMap<>();
    private boolean showTotalRow = false;
    private final JPanel totalRowPanel = new JPanel();
    private boolean autoExpandEnabled = false;
    private boolean alternateRowColors = false;
    private boolean tableMode = true; 
    private Object[][] previousData;
    private Object[] previousColumns;
    private Object[][] initialData;
    private Object[] initialColumns;
    public XlsEditorPanel(File file) {
        this.file = file;
        setLayout(new BorderLayout());
        model = new DefaultTableModel();
        table = new JTable(model) {
            @Override
            public Component prepareRenderer(TableCellRenderer renderer, int row, int col) {
                Component comp = super.prepareRenderer(renderer, row, col);
                if (alternateRowColors && tableMode) {
                    if (row % 2 == 0) {
                        comp.setBackground(Color.WHITE);
                    } else {
                        comp.setBackground(new Color(240, 240, 240));
                    }
                } else {
                    comp.setBackground(getBackground());
                }
                String key = row + "_" + col;
                if (cellColorMap.containsKey(key)) {
                    comp.setBackground(cellColorMap.get(key));
                }
                comp.setForeground(getForeground());
                return comp;
            }
        };
        table.setRowHeight(24);
        table.setFillsViewportHeight(true);
        table.setFont(new Font("Segoe UI", Font.PLAIN, 14));
        table.setShowGrid(true);
        table.setGridColor(Color.GRAY);
        table.setIntercellSpacing(new Dimension(1, 1));
        sorter = new TableRowSorter<>(model);
        table.setRowSorter(sorter);
        JScrollPane scroll = new JScrollPane(table);
        add(scroll, BorderLayout.CENTER);
        model.addTableModelListener(new TableModelListener() {
            public void tableChanged(TableModelEvent e) {
                if (autoExpandEnabled && tableMode) {
                    checkAutoExpand();
                }
                if (showTotalRow) {
                    updateTotalRowPanel();
                }
            }
        });
        JPanel bottomPanel = new JPanel(new BorderLayout());
        totalRowPanel.setPreferredSize(new Dimension(100, 30));
        totalRowPanel.setVisible(false);
        bottomPanel.add(totalRowPanel, BorderLayout.NORTH);
        JPanel controlPanel = new JPanel();
        JPanel searchPanel = new JPanel();
        JTextField searchField = new JTextField(20);
        searchField.setToolTipText("Type to filter table rows");
        searchPanel.add(new JLabel("Search: "));
        searchPanel.add(searchField);
        JButton clearSearchButton = new JButton("Clear");
        clearSearchButton.setToolTipText("Clear search filter");
        clearSearchButton.addActionListener(e -> searchField.setText(""));
        searchPanel.add(clearSearchButton);
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
        JButton refreshButton = new JButton("Refresh");
        refreshButton.setToolTipText("Reload file content");
        refreshButton.addActionListener((ActionEvent e) -> {
            model.setRowCount(0);
            model.setColumnCount(0);
            loadXlsContent();
            autoResizeColumns();
        });
        JButton saveButton = new JButton("💾 Save");
        saveButton.addActionListener(this::handleSave);
        JButton addRowButton = new JButton("Add Row");
        addRowButton.setToolTipText("Append a new empty row");
        addRowButton.addActionListener(this::handleAddRow);
        JButton deleteRowButton = new JButton("Delete Row");
        deleteRowButton.setToolTipText("Remove the selected row");
        deleteRowButton.addActionListener(this::handleDeleteRow);
        JButton exportCSVButton = new JButton("Export CSV");
        exportCSVButton.setToolTipText("Export table data as CSV");
        exportCSVButton.addActionListener(this::handleExportCSV);
        JButton insertImageButton = new JButton("Insert Image");
        insertImageButton.setToolTipText("Insert an image into the selected cell");
        insertImageButton.addActionListener(this::handleInsertImage);
        JButton setColorButton = new JButton("Set Cell Color");
        setColorButton.setToolTipText("Choose a background color for the selected cell");
        setColorButton.addActionListener(this::handleSetCellColor);
        JButton toggleTotalRowButton = new JButton("Toggle Total Row");
        toggleTotalRowButton.setToolTipText("Show/Hide a total row for numeric columns");
        toggleTotalRowButton.addActionListener(e -> {
            showTotalRow = !showTotalRow;
            totalRowPanel.setVisible(showTotalRow);
            if (showTotalRow) updateTotalRowPanel();
            revalidate();
            repaint();
        });
        JButton renameColumnsButton = new JButton("Rename Columns");
        renameColumnsButton.setToolTipText("Edit column headers for structured references");
        renameColumnsButton.addActionListener(e -> renameColumns());
        JButton toggleAutoExpandButton = new JButton("Toggle Auto Expand");
        toggleAutoExpandButton.setToolTipText("Automatically add a new row if the last row is modified");
        toggleAutoExpandButton.addActionListener(e -> {
            autoExpandEnabled = !autoExpandEnabled;
            JOptionPane.showMessageDialog(this, "Auto Expand is now " + (autoExpandEnabled ? "enabled" : "disabled"));
        });
        JButton toggleAltRowsButton = new JButton("Toggle Alt Row Colors");
        toggleAltRowsButton.setToolTipText("Enable/disable alternate row colors");
        toggleAltRowsButton.addActionListener(e -> {
            alternateRowColors = !alternateRowColors;
            table.repaint();
        });
        JButton convertToRangeButton = new JButton("Convert to Range");
        convertToRangeButton.setToolTipText("Convert this table to a normal range (disable Excel list features)");
        convertToRangeButton.addActionListener(e -> {
            tableMode = !tableMode;
            String message = tableMode ? "Excel List features enabled" : "Converted to normal range";
            JOptionPane.showMessageDialog(this, message);
            if (!tableMode) {
                autoExpandEnabled = false;
                showTotalRow = false;
                totalRowPanel.setVisible(false);
                alternateRowColors = false;
            }
            table.repaint();
        });
        JButton createChartButton = new JButton("Create Chart");
        createChartButton.setToolTipText("Create a chart from the selected data (placeholder)");
        createChartButton.addActionListener(e -> {
            JOptionPane.showMessageDialog(this, "Chart creation is not implemented.", "Create Chart", JOptionPane.INFORMATION_MESSAGE);
        });
        controlPanel.add(refreshButton);
        controlPanel.add(saveButton);
        controlPanel.add(addRowButton);
        controlPanel.add(deleteRowButton);
        controlPanel.add(exportCSVButton);
        controlPanel.add(insertImageButton);
        controlPanel.add(setColorButton);
        controlPanel.add(toggleTotalRowButton);
        controlPanel.add(renameColumnsButton);
        controlPanel.add(toggleAutoExpandButton);
        controlPanel.add(toggleAltRowsButton);
        controlPanel.add(convertToRangeButton);
        controlPanel.add(createChartButton);
        controlPanel.add(searchPanel);
        bottomPanel.add(controlPanel, BorderLayout.SOUTH);
        add(bottomPanel, BorderLayout.SOUTH);
        loadXlsContent();
        autoResizeColumns();
    }
    private void loadXlsContent() {
        try (FileInputStream fis = new FileInputStream(file);
             Workbook workbook = new XSSFWorkbook(fis)) {
            Sheet sheet = workbook.getSheetAt(0);
            int maxCols = 0;
            for (Row row : sheet) {
                if (row != null && row.getLastCellNum() > maxCols) {
                    maxCols = row.getLastCellNum();
                }
            }
            for (int i = 0; i < maxCols; i++) {
                model.addColumn("Column " + (i + 1));
            }
            for (Row row : sheet) {
                if (row == null)
                    continue;
                Object[] rowData = new Object[maxCols];
                for (int i = 0; i < maxCols; i++) {
                    Cell cell = row.getCell(i);
                    rowData[i] = (cell != null) ? getCellValue(cell) : "";
                }
                model.addRow(rowData);
            }
        } catch (IOException e) {
            JOptionPane.showMessageDialog(this, "Failed to load .xls file:\n" + e.getMessage(),
                    "Error", JOptionPane.ERROR_MESSAGE);
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
        for (int col = 0; col < table.getColumnCount(); col++) {
            int width = 50;
            for (int row = 0; row < table.getRowCount(); row++) {
                TableCellRenderer renderer = table.getCellRenderer(row, col);
                Component comp = table.prepareRenderer(renderer, row, col);
                width = Math.max(comp.getPreferredSize().width + 10, width);
            }
            table.getColumnModel().getColumn(col).setPreferredWidth(width);
        }
    }
    private void handleSave(ActionEvent e) {
        try (Workbook workbook = new XSSFWorkbook();
             FileOutputStream fos = new FileOutputStream(file)) {
            Sheet sheet = workbook.createSheet("Sheet1");
            for (int row = 0; row < model.getRowCount(); row++) {
                Row xssfRow = sheet.createRow(row);
                for (int col = 0; col < model.getColumnCount(); col++) {
                    Cell cell = xssfRow.createCell(col);
                    Object value = model.getValueAt(row, col);
                    if (value instanceof Number) {
                        cell.setCellValue(((Number) value).doubleValue());
                    } else if (value instanceof ImageIcon) {
                        cell.setCellValue("[Image]");
                    } else {
                        cell.setCellValue(value.toString());
                    }
                }
            }
            workbook.write(fos);
            JOptionPane.showMessageDialog(this, "✅ .xls saved successfully!", "Save", JOptionPane.INFORMATION_MESSAGE);
        } catch (IOException ex) {
            JOptionPane.showMessageDialog(this, "❌ Save failed:\n" + ex.getMessage(),
                    "Error", JOptionPane.ERROR_MESSAGE);
        }
    }
    private void handleAddRow(ActionEvent e) {
        if (model.getColumnCount() > 0) {
            Object[] emptyRow = new Object[model.getColumnCount()];
            model.addRow(emptyRow);
        }
    }
    private void handleDeleteRow(ActionEvent e) {
        int selectedRow = table.getSelectedRow();
        if (selectedRow != -1) {
            int modelRow = table.convertRowIndexToModel(selectedRow);
            model.removeRow(modelRow);
        } else {
            JOptionPane.showMessageDialog(this, "Please select a row to delete.");
        }
    }
    private void handleExportCSV(ActionEvent e) {
        JFileChooser chooser = new JFileChooser();
        chooser.setDialogTitle("Export as CSV");
        chooser.setSelectedFile(new File(file.getName().replaceFirst("[.][^.]+$", "") + ".csv"));
        int result = chooser.showSaveDialog(this);
        if (result == JFileChooser.APPROVE_OPTION) {
            File csvFile = chooser.getSelectedFile();
            try (BufferedWriter writer = new BufferedWriter(new FileWriter(csvFile))) {
                for (int col = 0; col < model.getColumnCount(); col++) {
                    writer.write(model.getColumnName(col));
                    if (col < model.getColumnCount() - 1) {
                        writer.write(",");
                    }
                }
                writer.newLine();
                for (int row = 0; row < model.getRowCount(); row++) {
                    for (int col = 0; col < model.getColumnCount(); col++) {
                        Object value = model.getValueAt(row, col);
                        writer.write(value != null ? value.toString() : "");
                        if (col < model.getColumnCount() - 1) {
                            writer.write(",");
                        }
                    }
                    writer.newLine();
                }
                writer.flush();
                JOptionPane.showMessageDialog(this, "CSV exported successfully!", "Export", JOptionPane.INFORMATION_MESSAGE);
            } catch (IOException ex) {
                JOptionPane.showMessageDialog(this, "CSV export failed:\n" + ex.getMessage(),
                        "Error", JOptionPane.ERROR_MESSAGE);
            }
        }
    }
    private void handleInsertImage(ActionEvent e) {
        int selectedRow = table.getSelectedRow();
        int selectedCol = table.getSelectedColumn();
        if (selectedRow != -1 && selectedCol != -1) {
            int modelRow = table.convertRowIndexToModel(selectedRow);
            int modelCol = table.convertColumnIndexToModel(selectedCol);
            JFileChooser fileChooser = new JFileChooser();
            fileChooser.setFileFilter(new javax.swing.filechooser.FileNameExtensionFilter("Image Files", "png", "jpg", "jpeg", "gif"));
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
    private void handleSetCellColor(ActionEvent e) {
        int selectedRow = table.getSelectedRow();
        int selectedCol = table.getSelectedColumn();
        if (selectedRow != -1 && selectedCol != -1) {
            Color color = JColorChooser.showDialog(this, "Choose Cell Color", Color.WHITE);
            if (color != null) {
                int modelRow = table.convertRowIndexToModel(selectedRow);
                int modelCol = table.convertColumnIndexToModel(selectedCol);
                String key = modelRow + "_" + modelCol;
                cellColorMap.put(key, color);
                table.repaint();
            }
        } else {
            JOptionPane.showMessageDialog(this, "Please select a cell to set its color.");
        }
    }
    private void checkAutoExpand() {
        int lastRow = model.getRowCount() - 1;
        if (lastRow >= 0) {
            boolean nonEmpty = false;
            for (int col = 0; col < model.getColumnCount(); col++) {
                Object value = model.getValueAt(lastRow, col);
                if (value != null && !value.toString().trim().isEmpty()) {
                    nonEmpty = true;
                    break;
                }
            }
            if (nonEmpty) {
                Object[] emptyRow = new Object[model.getColumnCount()];
                model.addRow(emptyRow);
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
                Object value = model.getValueAt(row, col);
                if (value instanceof Number) {
                    sum += ((Number) value).doubleValue();
                    numeric = true;
                } else if (value != null) {
                    try {
                        sum += Double.parseDouble(value.toString());
                        numeric = true;
                    } catch (NumberFormatException ex) {
                    }
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
private void handleAddColumn() {
    String colName = JOptionPane.showInputDialog(this, "Enter new column name:");
    if (colName != null && !colName.trim().isEmpty()) {
        model.addColumn(colName);
        int colIndex = model.getColumnCount() - 1;
        for (int row = 0; row < model.getRowCount(); row++) {
            model.setValueAt("", row, colIndex);
        }
        autoResizeColumns();
    }
}
private void handleDeleteColumn() {
    int colIndex = table.getSelectedColumn();
    if (colIndex == -1) {
        JOptionPane.showMessageDialog(this, "Please select a column to delete.");
        return;
    }
    int colCount = model.getColumnCount();
    if (colCount <= 1) {
        JOptionPane.showMessageDialog(this, "Cannot delete the last column.");
        return;
    }
    Object[] newColumns = new Object[colCount - 1];
    for (int i = 0, j = 0; i < colCount; i++) {
        if (i == colIndex) continue;
        newColumns[j++] = model.getColumnName(i);
    }
    int rowCount = model.getRowCount();
    Object[][] newData = new Object[rowCount][colCount - 1];
    for (int r = 0; r < rowCount; r++) {
        for (int i = 0, j = 0; i < colCount; i++) {
            if (i == colIndex) continue;
            newData[r][j++] = model.getValueAt(r, i);
        }
    }
    model.setDataVector(newData, newColumns);
    table.getTableHeader().repaint();
    autoResizeColumns();
}
private void handleDeleteImage() {
    int selectedRow = table.getSelectedRow();
    int selectedCol = table.getSelectedColumn();
    if (selectedRow != -1 && selectedCol != -1) {
        int modelRow = table.convertRowIndexToModel(selectedRow);
        int modelCol = table.convertColumnIndexToModel(selectedCol);
        Object value = model.getValueAt(modelRow, modelCol);
        if (value instanceof ImageIcon) {
            model.setValueAt("", modelRow, modelCol);
        } else {
            JOptionPane.showMessageDialog(this, "Selected cell does not contain an image.");
        }
    } else {
        JOptionPane.showMessageDialog(this, "Please select a cell first.");
    }
}
private void resetToPrevious() {
    if (previousData != null && previousColumns != null) {
        model.setDataVector(previousData, previousColumns);
        table.getTableHeader().repaint();
        autoResizeColumns();
    } else {
        JOptionPane.showMessageDialog(this, "No previous snapshot stored.");
    }
}
private void resetToInitial() {
    if (initialData != null && initialColumns != null) {
        model.setDataVector(initialData, initialColumns);
        table.getTableHeader().repaint();
        autoResizeColumns();
    } else {
        JOptionPane.showMessageDialog(this, "No initial snapshot stored.");
    }
}
    private void renameColumns() {
        int colCount = model.getColumnCount();
        if (colCount == 0) return;
        JPanel renamePanel = new JPanel();
        renamePanel.setLayout(new BoxLayout(renamePanel, BoxLayout.Y_AXIS));
        JTextField[] fields = new JTextField[colCount];
        for (int col = 0; col < colCount; col++) {
            JPanel panel = new JPanel();
            JLabel label = new JLabel("Column " + (col + 1) + ": ");
            JTextField field = new JTextField(model.getColumnName(col), 15);
            fields[col] = field;
            panel.add(label);
            panel.add(field);
            renamePanel.add(panel);
        }
        int option = JOptionPane.showConfirmDialog(this, renamePanel, "Rename Columns",
                        JOptionPane.OK_CANCEL_OPTION, JOptionPane.PLAIN_MESSAGE);
        if (option == JOptionPane.OK_OPTION) {
            for (int col = 0; col < colCount; col++) {
                String newName = fields[col].getText();
                table.getColumnModel().getColumn(col).setHeaderValue(newName);
            }
            table.getTableHeader().repaint();
        }
    }    public static Icon getTabIcon() {
        ImageIcon originalIcon = new ImageIcon(XlsEditorPanel.class.getResource("/icons/excel.png"));
        int fixedWidth = 16;
        int fixedHeight = 16;
        Image scaledImage = originalIcon.getImage().getScaledInstance(fixedWidth, fixedHeight, Image.SCALE_SMOOTH);
        return new ImageIcon(scaledImage);
    }
}
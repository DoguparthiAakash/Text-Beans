package com.example.notepadpp.media.TaskManager.TaskManager.src1;

import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.*;
import java.util.ArrayList;
import java.util.Comparator;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;

public class TaskManagerApp extends JFrame {
    public TaskManagerApp() {
        // Directly open the Task Manager frame, no login
        openTaskManager("default_user");
    }

    private void openTaskManager(String username) {
        TaskManagerFrame TM = new TaskManagerFrame(username);
        TM.setVisible(true);
        dispose();
    }

    private class TaskManagerFrame extends JFrame {
        private ArrayList<Task> tasks;
        private JTable taskTable;
        private DefaultTableModel tableModel;
        private String username;

        public TaskManagerFrame(String username) {
            this.username = username;

            setTitle("Task Manager");
            setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);
            setSize(500, 400);
            setLocationRelativeTo(null);

            tasks = new ArrayList<>();

            tableModel = new DefaultTableModel(new Object[] { "Title", "Description", "Priority", "Due Date" }, 0);
            taskTable = new JTable(tableModel);

            JButton addButton = new JButton("Add");
            JButton editButton = new JButton("Edit");
            JButton deleteButton = new JButton("Delete");

            addButton.setIcon(resizeIcon(new ImageIcon("add_icon.png"), 20, 20));
            editButton.setIcon(resizeIcon(new ImageIcon("edit_icon.png"), 20, 20));
            deleteButton.setIcon(resizeIcon(new ImageIcon("delete_icon.png"), 20, 20));

            addButton.addActionListener(new ActionListener() {
                @Override
                public void actionPerformed(ActionEvent e) {
                    addTask();
                }
            });

            editButton.addActionListener(new ActionListener() {
                @Override
                public void actionPerformed(ActionEvent e) {
                    editTask();
                }
            });

            deleteButton.addActionListener(new ActionListener() {
                @Override
                public void actionPerformed(ActionEvent e) {
                    deleteTask();
                }
            });

            JPanel buttonPanel = new JPanel();
            buttonPanel.setLayout(new FlowLayout());
            buttonPanel.add(addButton);
            buttonPanel.add(editButton);
            buttonPanel.add(deleteButton);

            JScrollPane scrollPane = new JScrollPane(taskTable);

            setLayout(new BorderLayout());
            add(scrollPane, BorderLayout.CENTER);
            add(buttonPanel, BorderLayout.SOUTH);

            loadTasksFromFile();
        }

        private void loadTasksFromFile() {
            try (BufferedReader reader = new BufferedReader(new FileReader(username + ".txt"))) {
                String line;
                while ((line = reader.readLine()) != null) {
                    String[] taskData = line.split(",");
                    String title = taskData[0];
                    String description = taskData[1];
                    String priority = taskData[2];
                    String dueDate = taskData[3];
                    tasks.add(new Task(title, description, priority, dueDate));
                }
            } catch (IOException e) {
            }

            tasks.sort(new Comparator<Task>() {
                @Override
                public int compare(Task task1, Task task2) {
                    int priority1 = getPriorityValue(task1.priority);
                    int priority2 = getPriorityValue(task2.priority);
                    return Integer.compare(priority1, priority2);
                }

                private int getPriorityValue(String priority) {
                    switch (priority) {
                        case "high":
                            return 0;
                        case "medium":
                            return 1;
                        case "low":
                            return 2;
                        default:
                            return 3;
                    }
                }
            });

            for (Task task : tasks) {
                tableModel.addRow(new Object[] { task.title, task.description, task.priority, task.dueDate });
            }
        }

        private void addTask() {
            String title = JOptionPane.showInputDialog(TaskManagerFrame.this, "Enter the task title:");
            String description = JOptionPane.showInputDialog(TaskManagerFrame.this, "Enter the task description:");

            String[] priorityOptions = { "high", "medium", "low" };
            String priority = (String) JOptionPane.showInputDialog(
                    TaskManagerFrame.this,
                    "Select the task priority:",
                    "Priority",
                    JOptionPane.PLAIN_MESSAGE,
                    null,
                    priorityOptions,
                    "medium");

            String dueDate;
            while (true) {
                dueDate = JOptionPane.showInputDialog(TaskManagerFrame.this, "Enter the task due date (yyyy-mm-dd):");
                if (isValidDateFormat(dueDate)) {
                    break;
                } else {
                    JOptionPane.showMessageDialog(TaskManagerFrame.this,
                            "Invalid date format. Please enter the date in yyyy-mm-dd format.");
                }
            }
            Task task = new Task(title, description, priority, dueDate);
            tasks.add(task);
            tableModel.addRow(new Object[] { task.title, task.description, task.priority, task.dueDate });

            saveTasksToFile();
        }

        private boolean isValidDateFormat(String date) {
            DateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd");
            dateFormat.setLenient(false);
            try {
                dateFormat.parse(date);
                return true;
            } catch (ParseException e) {
                return false;
            }
        }

        private void editTask() {
            int selectedRow = taskTable.getSelectedRow();
            if (selectedRow != -1) {
                Task task = tasks.get(selectedRow);

                String title = JOptionPane.showInputDialog(TaskManagerFrame.this, "Enter the new task title:",
                        task.title);
                String description = JOptionPane.showInputDialog(TaskManagerFrame.this,
                        "Enter the new task description:", task.description);

                String[] priorityOptions = { "high", "medium", "low" };
                String priority = (String) JOptionPane.showInputDialog(
                        TaskManagerFrame.this,
                        "Select the new task priority:",
                        "Priority",
                        JOptionPane.PLAIN_MESSAGE,
                        null,
                        priorityOptions,
                        task.priority);

                String dueDate;
                while (true) {
                    dueDate = JOptionPane.showInputDialog(TaskManagerFrame.this,
                            "Enter the new task due date (yyyy-mm-dd):");
                    if (isValidDateFormat(dueDate)) {
                        break;
                    } else {
                        JOptionPane.showMessageDialog(TaskManagerFrame.this,
                                "Invalid date format. Please enter the date in yyyy-mm-dd format.");
                    }
                }
                task.title = title;
                task.description = description;
                task.priority = priority;
                task.dueDate = dueDate;

                tableModel.setValueAt(task.title, selectedRow, 0);
                tableModel.setValueAt(task.description, selectedRow, 1);
                tableModel.setValueAt(task.priority, selectedRow, 2);
                tableModel.setValueAt(task.dueDate, selectedRow, 3);

                saveTasksToFile();
            }
        }

        private void deleteTask() {
            int selectedRow = taskTable.getSelectedRow();
            if (selectedRow != -1) {
                tasks.remove(selectedRow);
                tableModel.removeRow(selectedRow);

                saveTasksToFile();
            }
        }

        private void saveTasksToFile() {
            try (BufferedWriter writer = new BufferedWriter(new FileWriter(username + ".txt"))) {
                for (Task task : tasks) {
                    writer.write(task.title + "," + task.description + "," + task.priority + "," + task.dueDate);
                    writer.newLine();
                }
            } catch (IOException e) {
                e.printStackTrace();
                JOptionPane.showMessageDialog(TaskManagerFrame.this, "Error occurred while saving tasks to file.");
            }
        }

        private class Task {
            private String title;
            private String description;
            private String priority;
            private String dueDate;

            public Task(String title, String description, String priority, String dueDate) {
                this.title = title;
                this.description = description;
                this.priority = priority;
                this.dueDate = dueDate;
            }
        }

        private ImageIcon resizeIcon(ImageIcon icon, int width, int height) {
            Image img = icon.getImage();
            Image resizedImg = img.getScaledInstance(width, height, Image.SCALE_SMOOTH);
            return new ImageIcon(resizedImg);
        }
    }

    public static void main(String args[]) {
        try {
            for (UIManager.LookAndFeelInfo info : UIManager.getInstalledLookAndFeels()) {
                if ("Nimbus".equals(info.getName())) {
                    UIManager.setLookAndFeel(info.getClassName());
                    break;
                }
            }
        } catch (Exception ex) {
            java.util.logging.Logger.getLogger(TaskManagerApp.class.getName()).log(java.util.logging.Level.SEVERE, null,
                    ex);
        }

        java.awt.EventQueue.invokeLater(new Runnable() {
            public void run() {
                new TaskManagerApp().setVisible(true);
            }
        });
    }
}
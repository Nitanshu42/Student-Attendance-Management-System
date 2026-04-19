
import com.formdev.flatlaf.FlatLightLaf;

import javax.swing.*;
import javax.swing.border.Border;
import javax.swing.border.EmptyBorder;
import javax.swing.table.DefaultTableCellRenderer;
import javax.swing.table.DefaultTableModel;
import java.awt.BorderLayout;
import java.awt.CardLayout;
import java.awt.Color;
import java.awt.Component;
import java.awt.Cursor;
import java.awt.Dimension;
import java.awt.FlowLayout;
import java.awt.Font;
import java.awt.Frame;
import java.awt.GridBagConstraints;
import java.awt.GridBagLayout;
import java.awt.GridLayout;
import java.awt.Insets;
import java.awt.event.ActionEvent;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Objects;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class AttendanceSystemUI {

    // --- Main Entry Point ---
    public static void main(String[] args) {
        // Use a modern light theme
        FlatLightLaf.setup();
        // Global UI tweaks
        UIManager.put("Button.arc", 12);
        UIManager.put("Component.arc", 12);
        UIManager.put("ProgressBar.arc", 12);
        UIManager.put("TextComponent.arc", 12);
        UIManager.put("Table.showVerticalLines", false);
        UIManager.put("Table.intercellSpacing", new Dimension(0, 1));
        UIManager.put("TableHeader.separatorColor", new Color(220, 220, 220));

        ExcelDataManager.setupDatabase();

        SwingUtilities.invokeLater(() -> {
            LoginFrame loginFrame = new LoginFrame();
            loginFrame.setVisible(true);
        });
    }

    // --- Color and Font Constants ---
    static class AppStyles {

        public static final Color BACKGROUND_COLOR = new Color(245, 248, 251);
        public static final Color SIDENAV_COLOR = Color.WHITE;
        public static final Color PRIMARY_COLOR = new Color(59, 89, 152);
        public static final Color PRIMARY_TEXT_COLOR = Color.BLACK;
        public static final Color SECONDARY_TEXT_COLOR = new Color(101, 103, 107);
        public static final Color GREEN = new Color(27, 188, 155);
        public static final Color RED = new Color(231, 76, 60);
        public static final Color BORDER_COLOR = new Color(220, 223, 228);

        public static final Font FONT_BOLD = new Font("Segoe UI", Font.BOLD, 14);
        public static final Font FONT_NORMAL = new Font("Segoe UI", Font.PLAIN, 14);
        public static final Font FONT_HEADER = new Font("Segoe UI", Font.BOLD, 24);
        public static final Font FONT_SMALL = new Font("Segoe UI", Font.PLAIN, 12);
    }

    // --- Data Models ---
    static class User {

        String id, password, name, role;

        User(String id, String password, String name, String role) {
            this.id = id;
            this.password = password;
            this.name = name;
            this.role = role;
        }
    }

    static class AttendanceRecord {

        String studentId, date, status;

        AttendanceRecord(String studentId, String date, String status) {
            this.studentId = studentId;
            this.date = date;
            this.status = status;
        }
    }

    // --- Main Application Frame ---
    static class MainFrame extends JFrame {

        private final CardLayout cardLayout = new CardLayout();
        private final JPanel contentPanel = new JPanel(cardLayout);
        private User currentUser;
        private final SideNavPanel sideNavPanel;

        MainFrame(User user) {
            this.currentUser = user;
            setTitle("Attendance System");
            setSize(1200, 800);
            setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
            setLocationRelativeTo(null);
            setLayout(new BorderLayout());
            getContentPane().setBackground(AppStyles.BACKGROUND_COLOR);

            sideNavPanel = new SideNavPanel(this);
            add(sideNavPanel, BorderLayout.WEST);

            setupContentPanels();
            add(contentPanel, BorderLayout.CENTER);

            // Set initial view based on role
            switch (currentUser.role) {
                case "Admin":
                    sideNavPanel.setActive("Students");
                    showPanel("Admin");
                    break;
                case "Staff":
                    sideNavPanel.setActive("Attendance");
                    showPanel("Staff");
                    break;
                case "Student":
                    sideNavPanel.setActive("Dashboard");
                    showPanel("Student");
                    break;
            }
        }

        private void setupContentPanels() {
            // For simplicity, roles share panels, but visibility of nav items is controlled.
            contentPanel.add(new StaffAttendancePanel(this), "Staff");
            contentPanel.add(new AdminPanel(this), "Admin");
            contentPanel.add(new StudentDashboardPanel(this), "Student");
        }

        public void showPanel(String panelName) {
            cardLayout.show(contentPanel, panelName);
            // Dynamic refresh logic can be added here if needed when panels are shown
            Component[] components = contentPanel.getComponents();
            for (Component component : components) {
                if (component.getClass().getSimpleName().equals(panelName + "Panel") && component.isVisible()) {
                    if (component instanceof Refreshable) {
                        ((Refreshable) component).refreshData();
                    }
                    break;
                }
            }
        }

        public User getCurrentUser() {
            return currentUser;
        }
    }

    // --- Login Frame ---
    static class LoginFrame extends JFrame {

        LoginFrame() {
            setTitle("Login");
            setSize(400, 500);
            setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
            setLocationRelativeTo(null);
            setLayout(new GridBagLayout());
            getContentPane().setBackground(Color.WHITE);

            GridBagConstraints gbc = new GridBagConstraints();
            gbc.insets = new Insets(10, 20, 10, 20);
            gbc.gridwidth = 1;

            JLabel title = new JLabel("Welcome Back");
            title.setFont(new Font("Segoe UI", Font.BOLD, 28));
            gbc.gridx = 0;
            gbc.gridy = 0;
            add(title, gbc);

            JLabel subtitle = new JLabel("Login to your account");
            subtitle.setFont(AppStyles.FONT_NORMAL);
            subtitle.setForeground(AppStyles.SECONDARY_TEXT_COLOR);
            gbc.gridy = 1;
            add(subtitle, gbc);

            JTextField usernameField = new JTextField(20);
            JPasswordField passwordField = new JPasswordField(20);
            setupTextField(usernameField, "Username");
            setupTextField(passwordField, "Password");

            gbc.gridy = 2;
            gbc.fill = GridBagConstraints.HORIZONTAL;
            add(usernameField, gbc);
            gbc.gridy = 3;
            gbc.fill = GridBagConstraints.HORIZONTAL;
            add(passwordField, gbc);

            JButton loginButton = new JButton("Login");
            loginButton.setBackground(AppStyles.PRIMARY_COLOR);
            loginButton.setForeground(Color.WHITE);
            loginButton.setFont(AppStyles.FONT_BOLD);
            gbc.gridy = 4;
            gbc.insets = new Insets(20, 20, 10, 20);
            add(loginButton, gbc);

            loginButton.addActionListener(e -> performLogin(usernameField.getText(), new String(passwordField.getPassword())));
        }

        private void setupTextField(JTextField field, String placeholder) {
            field.putClientProperty("JTextField.placeholderText", placeholder);
            field.setPreferredSize(new Dimension(200, 40));
        }

        private void performLogin(String username, String password) {
            if (username.isEmpty() || password.isEmpty()) {
                JOptionPane.showMessageDialog(this, "Fields cannot be empty.", "Error", JOptionPane.ERROR_MESSAGE);
                return;
            }
            User user = ExcelDataManager.authenticateUser(username, password);
            if (user != null) {
                dispose();
                new MainFrame(user).setVisible(true);
            } else {
                JOptionPane.showMessageDialog(this, "Invalid credentials.", "Login Failed", JOptionPane.ERROR_MESSAGE);
            }
        }
    }

    // --- Side Navigation Panel ---
    static class SideNavPanel extends JPanel {

        private final MainFrame mainFrame;
        private final JPanel buttonsPanel = new JPanel();
        private JButton activeButton;

        SideNavPanel(MainFrame frame) {
            this.mainFrame = frame;
            setLayout(new BorderLayout(0, 20));
            setPreferredSize(new Dimension(220, 0));
            setBackground(AppStyles.SIDENAV_COLOR);
            setBorder(BorderFactory.createMatteBorder(0, 0, 0, 1, AppStyles.BORDER_COLOR));

            JLabel titleLabel = new JLabel("Attendance System", SwingConstants.CENTER);
            titleLabel.setFont(new Font("Segoe UI", Font.BOLD, 18));
            titleLabel.setBorder(new EmptyBorder(20, 10, 20, 10));
            add(titleLabel, BorderLayout.NORTH);

            buttonsPanel.setLayout(new GridLayout(0, 1, 0, 10));
            buttonsPanel.setOpaque(false);
            buttonsPanel.setBorder(new EmptyBorder(10, 10, 10, 10));

            String role = mainFrame.getCurrentUser().role;
            switch (role) {
                case "Admin":
                    addNavButton("\uD83D\uDC65", "Students", "Admin");
                    addNavButton("\uD83D\uDCCB", "Reports", "AdminReports"); // Placeholder for future report panel
                    break;
                case "Staff":
                    addNavButton("\uD83D\uDCCB", "Attendance", "Staff");
                    break;
                case "Student":
                    addNavButton("\uD83D\uDCCA", "Dashboard", "Student");
                    break;
            }

            add(buttonsPanel, BorderLayout.CENTER);

            JButton logoutButton = new JButton("Logout");
            logoutButton.addActionListener(e -> {
                mainFrame.dispose();
                new LoginFrame().setVisible(true);
            });
            add(logoutButton, BorderLayout.SOUTH);
        }

        private void addNavButton(String icon, String text, String panelName) {
            JButton button = new JButton(icon + "    " + text);
            button.setFont(AppStyles.FONT_BOLD);
            button.setFocusPainted(false);
            button.setHorizontalAlignment(SwingConstants.LEFT);
            button.setBorder(new EmptyBorder(10, 20, 10, 20));
            button.setCursor(new Cursor(Cursor.HAND_CURSOR));
            setInactiveStyle(button);

            button.addActionListener(e -> {
                setActive(text);

                // This logic needs to be robust for different panel mappings
                if (panelName.equals("Admin")) {
                    mainFrame.showPanel("Admin"); 
                }else if (panelName.equals("Staff")) {
                    mainFrame.showPanel("Staff"); 
                }else if (panelName.equals("Student")) {
                    mainFrame.showPanel("Student");
                }
            });

            buttonsPanel.add(button);
        }

        public void setActive(String text) {
            if (activeButton != null) {
                setInactiveStyle(activeButton);
            }
            for (Component comp : buttonsPanel.getComponents()) {
                if (comp instanceof JButton && ((JButton) comp).getText().contains(text)) {
                    activeButton = (JButton) comp;
                    setActiveStyle(activeButton);
                    break;
                }
            }
        }

        private void setActiveStyle(JButton button) {
            button.setBackground(AppStyles.PRIMARY_COLOR);
            button.setForeground(Color.WHITE);
        }

        private void setInactiveStyle(JButton button) {
            button.setBackground(AppStyles.SIDENAV_COLOR);
            button.setForeground(AppStyles.SECONDARY_TEXT_COLOR);
        }
    }

    // --- Reusable UI Components ---
    static class CustomComponents {

        public static JPanel createStatCard(String title, String value, String icon, Color color) {
            JPanel card = new JPanel(new BorderLayout(10, 0));
            card.setBackground(Color.WHITE);
            card.setBorder(BorderFactory.createCompoundBorder(
                    BorderFactory.createLineBorder(AppStyles.BORDER_COLOR),
                    new EmptyBorder(15, 15, 15, 15)
            ));

            JLabel iconLabel = new JLabel(icon);
            iconLabel.setFont(new Font("Segoe UI Symbol", Font.PLAIN, 28));
            iconLabel.setForeground(color);
            card.add(iconLabel, BorderLayout.WEST);

            JPanel textPanel = new JPanel();
            textPanel.setOpaque(false);
            textPanel.setLayout(new BoxLayout(textPanel, BoxLayout.Y_AXIS));
            JLabel valueLabel = new JLabel(value);
            valueLabel.setFont(new Font("Segoe UI", Font.BOLD, 22));
            JLabel titleLabel = new JLabel(title);
            titleLabel.setFont(AppStyles.FONT_SMALL);
            titleLabel.setForeground(AppStyles.SECONDARY_TEXT_COLOR);
            textPanel.add(valueLabel);
            textPanel.add(titleLabel);
            card.add(textPanel, BorderLayout.CENTER);
            return card;
        }

        public static class StatusCellRenderer extends DefaultTableCellRenderer {

            @Override
            public Component getTableCellRendererComponent(JTable table, Object value, boolean isSelected, boolean hasFocus, int row, int column) {
                Component c = super.getTableCellRendererComponent(table, value, isSelected, hasFocus, row, column);
                c.setFont(AppStyles.FONT_BOLD);
                setHorizontalAlignment(SwingConstants.CENTER);
                setBorder(new EmptyBorder(8, 15, 8, 15));

                if ("Present".equals(value)) {
                    c.setBackground(AppStyles.GREEN.brighter().brighter());
                    c.setForeground(AppStyles.GREEN.darker());
                } else if ("Absent".equals(value)) {
                    c.setBackground(AppStyles.RED.brighter().brighter());
                    c.setForeground(AppStyles.RED.darker());
                } else {
                    c.setBackground(table.getBackground());
                    c.setForeground(table.getForeground());
                }

                if (isSelected) {
                    c.setBackground(table.getSelectionBackground());
                    c.setForeground(table.getSelectionForeground());
                }

                return c;
            }
        }
    }

    // An interface to standardize refreshing panel data
    interface Refreshable {

        void refreshData();
    }

    // --- Staff View Panel ---
    static class StaffAttendancePanel extends JPanel implements Refreshable {

        private final MainFrame mainFrame;
        private final DefaultTableModel tableModel;
        private final JLabel presentCountLabel = new JLabel("0");
        private final JLabel absentCountLabel = new JLabel("0");
        private final JLabel attendancePercentLabel = new JLabel("0%");

        StaffAttendancePanel(MainFrame frame) {
            this.mainFrame = frame;
            setLayout(new BorderLayout(20, 20));
            setBorder(new EmptyBorder(20, 20, 20, 20));
            setBackground(AppStyles.BACKGROUND_COLOR);

            // Top Header
            JPanel headerPanel = new JPanel(new BorderLayout());
            headerPanel.setOpaque(false);
            JLabel title = new JLabel("Attendance");
            title.setFont(AppStyles.FONT_HEADER);
            headerPanel.add(title, BorderLayout.WEST);
            add(headerPanel, BorderLayout.NORTH);

            // Stats Panel
            JPanel statsPanel = new JPanel(new GridLayout(1, 4, 15, 15));
            statsPanel.setOpaque(false);
            statsPanel.add(CustomComponents.createStatCard("Present Today", presentCountLabel.getText(), "\u2714", AppStyles.GREEN));
            statsPanel.add(CustomComponents.createStatCard("Absent Today", absentCountLabel.getText(), "\u2716", AppStyles.RED));
            statsPanel.add(CustomComponents.createStatCard("Attendance %", attendancePercentLabel.getText(), "%", AppStyles.PRIMARY_COLOR));
            // You can add logic for holidays
            statsPanel.add(CustomComponents.createStatCard("Upcoming Holidays", "0", "\uD83D\uDCC5", Color.ORANGE));

            // Main Content Panel
            JPanel mainContent = new JPanel(new BorderLayout());
            mainContent.setOpaque(false);
            mainContent.add(statsPanel, BorderLayout.NORTH);

            // Table Panel
            JPanel tableContainer = new JPanel(new BorderLayout(10, 10));
            tableContainer.setOpaque(false);
            tableContainer.setBorder(new EmptyBorder(20, 0, 0, 0));

            // Table Toolbar
            JPanel toolbar = new JPanel(new FlowLayout(FlowLayout.LEFT));
            toolbar.setOpaque(false);
            toolbar.add(new JLabel(new SimpleDateFormat("EEE, d MMM yyyy").format(new Date())));
            tableContainer.add(toolbar, BorderLayout.NORTH);

            // Table
            String[] columns = {"Roll No.", "Name", "Status"};
            tableModel = new DefaultTableModel(columns, 0) {
                @Override
                public boolean isCellEditable(int row, int column) {
                    return false;
                }
            };
            JTable table = new JTable(tableModel);
            table.setRowHeight(40);
            table.getTableHeader().setFont(AppStyles.FONT_BOLD);
            table.setFont(AppStyles.FONT_NORMAL);
            table.getColumnModel().getColumn(2).setCellRenderer(new CustomComponents.StatusCellRenderer());

            // Add mouse listener to toggle status
            table.addMouseListener(new MouseAdapter() {
                public void mouseClicked(MouseEvent e) {
                    int row = table.rowAtPoint(e.getPoint());
                    int col = table.columnAtPoint(e.getPoint());
                    if (row >= 0 && col == 2) {
                        String currentStatus = (String) tableModel.getValueAt(row, 2);
                        String newStatus = "Present".equals(currentStatus) ? "Absent" : "Present";
                        tableModel.setValueAt(newStatus, row, 2);
                        updateStats();
                    }
                }
            });

            JScrollPane scrollPane = new JScrollPane(table);
            scrollPane.getViewport().setBackground(Color.WHITE);
            scrollPane.setBorder(BorderFactory.createLineBorder(AppStyles.BORDER_COLOR));
            tableContainer.add(scrollPane, BorderLayout.CENTER);

            // Bottom Actions
            JPanel actionsPanel = new JPanel(new FlowLayout(FlowLayout.LEFT, 10, 0));
            actionsPanel.setOpaque(false);
            JButton markAllPresent = new JButton("Mark All Present");
            JButton markAllAbsent = new JButton("Mark All Absent");
            JButton saveButton = new JButton("Save");
            saveButton.setBackground(AppStyles.PRIMARY_COLOR);
            saveButton.setForeground(Color.WHITE);
            actionsPanel.add(markAllPresent);
            actionsPanel.add(markAllAbsent);
            actionsPanel.add(saveButton);

            tableContainer.add(actionsPanel, BorderLayout.SOUTH);

            markAllPresent.addActionListener(e -> setAllStatus("Present"));
            markAllAbsent.addActionListener(e -> setAllStatus("Absent"));
            saveButton.addActionListener(this::submitAttendance);

            mainContent.add(tableContainer, BorderLayout.CENTER);
            add(mainContent, BorderLayout.CENTER);

            refreshData();
        }

        @Override
        public void refreshData() {
            tableModel.setRowCount(0);
            List<User> students = ExcelDataManager.getUsersByRole("Student");
            for (User student : students) {
                // Check if attendance for today is already marked for this student
                String todayStatus = ExcelDataManager.getStudentStatusForToday(student.id);
                tableModel.addRow(new Object[]{student.id, student.name, todayStatus});
            }
            updateStats();
        }

        private void updateStats() {
            int present = 0, absent = 0;
            for (int i = 0; i < tableModel.getRowCount(); i++) {
                if ("Present".equals(tableModel.getValueAt(i, 2))) {
                    present++; 
                }else {
                    absent++;
                }
            }
            int total = tableModel.getRowCount();
            double percentage = (total == 0) ? 0 : ((double) present / total) * 100;

            // Re-create stat cards with new values. A bit inefficient but simple.
            JPanel statsPanel = (JPanel) ((JPanel) this.getComponent(1)).getComponent(0);
            statsPanel.removeAll();
            statsPanel.add(CustomComponents.createStatCard("Present Today", String.valueOf(present), "\u2714", AppStyles.GREEN));
            statsPanel.add(CustomComponents.createStatCard("Absent Today", String.valueOf(absent), "\u2716", AppStyles.RED));
            statsPanel.add(CustomComponents.createStatCard("Attendance %", String.format("%.0f%%", percentage), "%", AppStyles.PRIMARY_COLOR));
            statsPanel.add(CustomComponents.createStatCard("Upcoming Holidays", "0", "\uD83D\uDCC5", Color.ORANGE));
            statsPanel.revalidate();
            statsPanel.repaint();
        }

        private void setAllStatus(String status) {
            for (int i = 0; i < tableModel.getRowCount(); i++) {
                tableModel.setValueAt(status, i, 2);
            }
            updateStats();
        }

        private void submitAttendance(ActionEvent e) {
            List<AttendanceRecord> records = new ArrayList<>();
            String date = new SimpleDateFormat("yyyy-MM-dd").format(new Date());

            for (int i = 0; i < tableModel.getRowCount(); i++) {
                records.add(new AttendanceRecord((String) tableModel.getValueAt(i, 0), date, (String) tableModel.getValueAt(i, 2)));
            }

            if (ExcelDataManager.hasAttendanceBeenMarkedToday()) {
                int choice = JOptionPane.showConfirmDialog(this,
                        "Overwrite today's attendance record?", "Confirm", JOptionPane.YES_NO_OPTION);
                if (choice == JOptionPane.NO_OPTION) {
                    return;
                }
            }

            ExcelDataManager.markAttendance(records);
            JOptionPane.showMessageDialog(this, "Attendance saved successfully!", "Success", JOptionPane.INFORMATION_MESSAGE);
        }
    }

    // --- Admin Panel (Simplified version using tabs) ---
    static class AdminPanel extends JPanel implements Refreshable {

        private final MainFrame mainFrame;
        private final JTable studentTable, staffTable, reportTable;
        private final DefaultTableModel studentModel, staffModel, reportModel;

        AdminPanel(MainFrame frame) {
            this.mainFrame = frame;
            setLayout(new BorderLayout(20, 20));
            setBorder(new EmptyBorder(20, 20, 20, 20));
            setBackground(AppStyles.BACKGROUND_COLOR);

            JLabel title = new JLabel("Admin Dashboard");
            title.setFont(AppStyles.FONT_HEADER);
            add(title, BorderLayout.NORTH);

            JTabbedPane tabbedPane = new JTabbedPane();

            // Manage Students
            studentModel = new DefaultTableModel(new String[]{"ID", "Name"}, 0);
            studentTable = new JTable(studentModel);
            tabbedPane.addTab("Manage Students", createManagementPanel(studentTable, "Student"));

            // Manage Staff
            staffModel = new DefaultTableModel(new String[]{"ID", "Name"}, 0);
            staffTable = new JTable(staffModel);
            tabbedPane.addTab("Manage Staff", createManagementPanel(staffTable, "Staff"));

            // Attendance Reports
            reportModel = new DefaultTableModel(new String[]{"Student ID", "Name", "Date", "Status"}, 0);
            reportTable = new JTable(reportModel);
            reportTable.getColumnModel().getColumn(3).setCellRenderer(new CustomComponents.StatusCellRenderer());
            tabbedPane.addTab("Attendance Report", new JScrollPane(reportTable));

            add(tabbedPane, BorderLayout.CENTER);
            refreshData();
        }

        private JPanel createManagementPanel(JTable table, String role) {
            JPanel panel = new JPanel(new BorderLayout(10, 10));
            panel.add(new JScrollPane(table), BorderLayout.CENTER);

            JPanel buttonPanel = new JPanel(new FlowLayout(FlowLayout.LEFT));
            JButton addButton = new JButton("Add " + role);
            JButton removeButton = new JButton("Remove Selected " + role);

            addButton.addActionListener(e -> {
                new AddUserDialog(mainFrame, role, this::refreshData).setVisible(true);
            });
            removeButton.addActionListener(e -> {
                int selectedRow = table.getSelectedRow();
                if (selectedRow != -1) {
                    String id = (String) table.getModel().getValueAt(selectedRow, 0);
                    ExcelDataManager.removeUser(id);
                    refreshData();
                } else {
                    JOptionPane.showMessageDialog(this, "Please select a user to remove.", "Warning", JOptionPane.WARNING_MESSAGE);
                }
            });

            buttonPanel.add(addButton);
            buttonPanel.add(removeButton);
            panel.add(buttonPanel, BorderLayout.SOUTH);
            return panel;
        }

        @Override
        public void refreshData() {
            // Students
            studentModel.setRowCount(0);
            ExcelDataManager.getUsersByRole("Student").forEach(u -> studentModel.addRow(new Object[]{u.id, u.name}));
            // Staff
            staffModel.setRowCount(0);
            ExcelDataManager.getUsersByRole("Staff").forEach(u -> staffModel.addRow(new Object[]{u.id, u.name}));
            // Reports
            reportModel.setRowCount(0);
            ExcelDataManager.getAllAttendance().forEach(r -> {
                User student = ExcelDataManager.getUserById(r.studentId);
                reportModel.addRow(new Object[]{r.studentId, student != null ? student.name : "N/A", r.date, r.status});
            });
        }
    }

    // --- Student Dashboard Panel ---
    static class StudentDashboardPanel extends JPanel implements Refreshable {

        private final MainFrame mainFrame;
        private final DefaultTableModel tableModel;

        StudentDashboardPanel(MainFrame frame) {
            this.mainFrame = frame;
            setLayout(new BorderLayout(20, 20));
            setBorder(new EmptyBorder(20, 20, 20, 20));
            setBackground(AppStyles.BACKGROUND_COLOR);

            JLabel title = new JLabel("My Dashboard");
            title.setFont(AppStyles.FONT_HEADER);
            add(title, BorderLayout.NORTH);

            // This will hold both stats and table
            JPanel contentPanel = new JPanel(new BorderLayout(10, 20));
            contentPanel.setOpaque(false);

            JPanel statsPanel = new JPanel(new GridLayout(1, 3, 15, 15));
            statsPanel.setOpaque(false);
            // Stat cards will be added dynamically in refreshData()
            contentPanel.add(statsPanel, BorderLayout.NORTH);

            tableModel = new DefaultTableModel(new String[]{"Date", "Status"}, 0);
            JTable table = new JTable(tableModel);
            table.setRowHeight(40);
            table.getColumnModel().getColumn(1).setCellRenderer(new CustomComponents.StatusCellRenderer());
            contentPanel.add(new JScrollPane(table), BorderLayout.CENTER);

            add(contentPanel, BorderLayout.CENTER);
            refreshData();
        }

        @Override
        public void refreshData() {
            User currentUser = mainFrame.getCurrentUser();
            tableModel.setRowCount(0);
            List<AttendanceRecord> records = ExcelDataManager.getAttendanceForStudent(currentUser.id);
            int present = 0, absent = 0;
            for (AttendanceRecord r : records) {
                tableModel.addRow(new Object[]{r.date, r.status});
                if ("Present".equals(r.status)) {
                    present++; 
                }else {
                    absent++;
                }
            }
            int total = records.size();
            double percentage = (total == 0) ? 0 : ((double) present / total) * 100;

            JPanel statsPanel = (JPanel) ((JPanel) this.getComponent(1)).getComponent(0);
            statsPanel.removeAll();
            statsPanel.add(CustomComponents.createStatCard("Total Present", String.valueOf(present), "\u2714", AppStyles.GREEN));
            statsPanel.add(CustomComponents.createStatCard("Total Absent", String.valueOf(absent), "\u2716", AppStyles.RED));
            statsPanel.add(CustomComponents.createStatCard("Overall Percentage", String.format("%.0f%%", percentage), "%", AppStyles.PRIMARY_COLOR));
            statsPanel.revalidate();
            statsPanel.repaint();
        }
    }

    // --- Add User Dialog ---
    static class AddUserDialog extends JDialog {

        AddUserDialog(Frame owner, String role, Runnable refreshCallback) {
            super(owner, "Add New " + role, true);
            setSize(350, 250);
            setLocationRelativeTo(owner);
            // Basic layout, can be improved
            setLayout(new GridLayout(4, 2, 10, 10));
            add(new JLabel("ID:"));
            JTextField idField = new JTextField();
            add(idField);
            add(new JLabel("Full Name:"));
            JTextField nameField = new JTextField();
            add(nameField);
            add(new JLabel("Password:"));
            JPasswordField passField = new JPasswordField();
            add(passField);

            JButton addButton = new JButton("Add");
            addButton.addActionListener(e -> {
                ExcelDataManager.addUser(new User(idField.getText(), new String(passField.getPassword()), nameField.getText(), role));
                refreshCallback.run();
                dispose();
            });
            add(new JLabel()); // placeholder
            add(addButton);
        }
    }

    // --- Excel Data Manager (unchanged logic) ---
    static class ExcelDataManager {

        private static final DataFormatter FORMATTER = new DataFormatter();

        private static String getCellValueSafe(Row row, int cellIndex) {
            Cell cell = row.getCell(cellIndex);
            if (cell == null) return "";
            return FORMATTER.formatCellValue(cell).trim();
        }

        private static final String FILE_NAME = "college_data.xlsx";
        private static final String USERS_SHEET = "Users";
        private static final String ATTENDANCE_SHEET = "Attendance";

        public static void setupDatabase() {
            if (!new File(FILE_NAME).exists()) {
                try (Workbook workbook = new XSSFWorkbook()) {
                    Sheet usersSheet = workbook.createSheet(USERS_SHEET);
                    Row header = usersSheet.createRow(0);
                    header.createCell(0).setCellValue("ID");
                    header.createCell(1).setCellValue("Password");
                    header.createCell(2).setCellValue("Name");
                    header.createCell(3).setCellValue("Role");
                    Row adminRow = usersSheet.createRow(1);
                    adminRow.createCell(0).setCellValue("admin");
                    adminRow.createCell(1).setCellValue("admin123");
                    adminRow.createCell(2).setCellValue("Administrator");
                    adminRow.createCell(3).setCellValue("Admin");

                    Sheet attendanceSheet = workbook.createSheet(ATTENDANCE_SHEET);
                    Row attHeader = attendanceSheet.createRow(0);
                    attHeader.createCell(0).setCellValue("StudentID");
                    attHeader.createCell(1).setCellValue("Date");
                    attHeader.createCell(2).setCellValue("Status");

                    try (FileOutputStream fos = new FileOutputStream(FILE_NAME)) {
                        workbook.write(fos);
                    }
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }

        public static User authenticateUser(String id, String password) {
            try (FileInputStream fis = new FileInputStream(FILE_NAME); Workbook workbook = new XSSFWorkbook(fis)) {
                Sheet sheet = workbook.getSheet(USERS_SHEET);
                for (Row row : sheet) {
                    if (row.getRowNum() == 0) {
                        continue;
                    }
                    if (Objects.equals(getCellValueSafe(row, 0), id) && Objects.equals(getCellValueSafe(row, 1), password)) {
                        return new User(id, password, getCellValueSafe(row, 2), getCellValueSafe(row, 3));
                    }
                }
            } catch (Exception e) {
                e.printStackTrace();
            }
            return null;
        }

        public static void addUser(User user) {
            try (FileInputStream fis = new FileInputStream(FILE_NAME); Workbook workbook = new XSSFWorkbook(fis)) {
                Sheet sheet = workbook.getSheet(USERS_SHEET);
                Row newRow = sheet.createRow(sheet.getLastRowNum() + 1);
                newRow.createCell(0).setCellValue(user.id);
                newRow.createCell(1).setCellValue(user.password);
                newRow.createCell(2).setCellValue(user.name);
                newRow.createCell(3).setCellValue(user.role);
                try (FileOutputStream fos = new FileOutputStream(FILE_NAME)) {
                    workbook.write(fos);
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

        public static void removeUser(String id) {
            try (FileInputStream fis = new FileInputStream(FILE_NAME); Workbook workbook = new XSSFWorkbook(fis)) {
                Sheet sheet = workbook.getSheet(USERS_SHEET);
                int rowToRemove = -1;
                for (Row row : sheet) {
                    if (row.getRowNum() == 0) {
                        continue;
                    }
                    if (getCellValueSafe(row, 0).equalsIgnoreCase(id)) {
                        rowToRemove = row.getRowNum();
                        break;
                    }
                }
                if (rowToRemove != -1) {
                    sheet.removeRow(sheet.getRow(rowToRemove));
                    if (rowToRemove < sheet.getLastRowNum()) {
                        sheet.shiftRows(rowToRemove + 1, sheet.getLastRowNum(), -1);
                    }
                }
                try (FileOutputStream fos = new FileOutputStream(FILE_NAME)) {
                    workbook.write(fos);
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

        public static List<User> getUsersByRole(String role) {
            List<User> users = new ArrayList<>();
            try (FileInputStream fis = new FileInputStream(FILE_NAME); Workbook workbook = new XSSFWorkbook(fis)) {
                Sheet sheet = workbook.getSheet(USERS_SHEET);
                for (Row row : sheet) {
                    if (row.getRowNum() == 0) {
                        continue;
                    }
                    if (getCellValueSafe(row, 3).equalsIgnoreCase(role)) {
                        users.add(new User(getCellValueSafe(row, 0), "", getCellValueSafe(row, 2), role));
                    }
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
            return users;
        }

        public static User getUserById(String userId) {
            try (FileInputStream fis = new FileInputStream(FILE_NAME); Workbook workbook = new XSSFWorkbook(fis)) {
                Sheet sheet = workbook.getSheet(USERS_SHEET);
                for (Row row : sheet) {
                    if (row.getRowNum() == 0) {
                        continue;
                    }
                    if (getCellValueSafe(row, 0).equalsIgnoreCase(userId)) {
                        return new User(getCellValueSafe(row, 0), "", getCellValueSafe(row, 2), getCellValueSafe(row, 3));
                    }
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
            return null;
        }

        public static List<AttendanceRecord> getAllAttendance() {
            List<AttendanceRecord> records = new ArrayList<>();
            try (FileInputStream fis = new FileInputStream(FILE_NAME); Workbook workbook = new XSSFWorkbook(fis)) {
                Sheet sheet = workbook.getSheet(ATTENDANCE_SHEET);
                for (Row row : sheet) {
                    if (row.getRowNum() == 0) {
                        continue;
                    }
                    records.add(new AttendanceRecord(getCellValueSafe(row, 0), getCellValueSafe(row, 1), getCellValueSafe(row, 2)));
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
            return records;
        }

        public static List<AttendanceRecord> getAttendanceForStudent(String studentId) {
            List<AttendanceRecord> records = new ArrayList<>();
            for (AttendanceRecord r : getAllAttendance()) {
                if (r.studentId.equalsIgnoreCase(studentId)) {
                    records.add(r);
                }
            }
            return records;
        }

        public static boolean hasAttendanceBeenMarkedToday() {
            String today = new SimpleDateFormat("yyyy-MM-dd").format(new Date());
            return getAllAttendance().stream().anyMatch(r -> r.date.equals(today));
        }

        public static String getStudentStatusForToday(String studentId) {
            String today = new SimpleDateFormat("yyyy-MM-dd").format(new Date());
            return getAllAttendance().stream()
                    .filter(r -> r.date.equals(today) && r.studentId.equals(studentId))
                    .map(r -> r.status)
                    .findFirst().orElse("Present"); // Default to present
        }

        public static void markAttendance(List<AttendanceRecord> records) {
            try (FileInputStream fis = new FileInputStream(FILE_NAME); Workbook workbook = new XSSFWorkbook(fis)) {
                Sheet sheet = workbook.getSheet(ATTENDANCE_SHEET);
                String today = new SimpleDateFormat("yyyy-MM-dd").format(new Date());

                // Remove old records for today
                List<Integer> rowsToRemove = new ArrayList<>();
                for (Row row : sheet) {
                    if (row.getRowNum() > 0 && row.getCell(1) != null && getCellValueSafe(row, 1).equals(today)) {
                        rowsToRemove.add(row.getRowNum());
                    }
                }
                for (int i = rowsToRemove.size() - 1; i >= 0; i--) {
                    int rowIndex = rowsToRemove.get(i);
                    Row row = sheet.getRow(rowIndex);
                    if (row != null) {
                        sheet.removeRow(row);
                    }
                }
                // Shift rows up
                if (!rowsToRemove.isEmpty()) {
                    int firstRow = rowsToRemove.get(0);
                    int lastRow = sheet.getLastRowNum();
                    if (firstRow <= lastRow) {
                        sheet.shiftRows(firstRow + rowsToRemove.size(), lastRow, -rowsToRemove.size());
                    }
                }

                // Add new records
                for (AttendanceRecord record : records) {
                    Row newRow = sheet.createRow(sheet.getLastRowNum() + 1);
                    newRow.createCell(0).setCellValue(record.studentId);
                    newRow.createCell(1).setCellValue(record.date);
                    newRow.createCell(2).setCellValue(record.status);
                }

                try (FileOutputStream fos = new FileOutputStream(FILE_NAME)) {
                    workbook.write(fos);
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }
}

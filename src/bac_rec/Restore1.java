package bac_rec;

import java.awt.BorderLayout;
import java.awt.EventQueue;

import javax.swing.JFrame;
import javax.swing.JPanel;
import javax.swing.border.EmptyBorder;
import javax.swing.filechooser.FileNameExtensionFilter;

import jxl.Sheet;
import jxl.Workbook;

import javax.swing.JLabel;
import javax.swing.JOptionPane;

import java.awt.Font;
import javax.swing.JComboBox;
import javax.swing.JFileChooser;
import javax.swing.SwingConstants;
import javax.swing.GroupLayout;
import javax.swing.GroupLayout.Alignment;
import javax.swing.LayoutStyle.ComponentPlacement;
import javax.swing.JButton;
import java.awt.FlowLayout;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileNotFoundException;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.SQLIntegrityConstraintViolationException;
import java.sql.SQLSyntaxErrorException;
import java.util.ArrayList;
import java.util.Scanner;
import java.awt.event.ActionEvent;
import javax.swing.JTextField;
import java.awt.Toolkit;
import javax.swing.ImageIcon;
import javax.swing.JPasswordField;

public class Restore1 extends JFrame {

	private JPanel contentPane;
	private JTextField usernameTextField;
	private File fileToOpen = null;
	private String nameOfFile = "";
	private JLabel label1;
	private JComboBox comboBox;
	private PreparedStatement pst;
	private ResultSet rs;
	private Connection con;
	private int i, c, r = 3, totalPrimaryKey = 0;
	private String sql, columnName, columnNameType, add, data, length, pkAdd;
	private ArrayList<String> list = new ArrayList<String>();
	private ArrayList<String> pk = new ArrayList<String>();
	private String name = null, username = null, password = null, primaryKey = null, primaryConstraint = null;
	private String[] sheetNames;
	private Sheet sheet;
	private int totalNoOfRows;
	private int totalNoOfCols;
	private JPasswordField passwordField;

	/**
	 * Launch the application.
	 */
	public static void call_restore() {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					Restore1 frame = new Restore1();
					frame.setTitle("Restore");
					frame.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	/**
	 * Create the frame.
	 */
	public Restore1() {
		setIconImage(Toolkit.getDefaultToolkit().getImage(Restore1.class.getResource("/gui/backup.png")));

		setBounds(100, 100, 700, 448);
		contentPane = new JPanel();
		contentPane.setBorder(new EmptyBorder(5, 5, 5, 5));
		setContentPane(contentPane);
		setResizable(false);

		JPanel panel = new JPanel();
		panel.setBounds(288, 11, 175, 38);
		panel.setLayout(new FlowLayout(FlowLayout.CENTER, 5, 5));

		JLabel lblChooseDatabase = new JLabel("Choose Database");
		lblChooseDatabase.setHorizontalAlignment(SwingConstants.CENTER);
		lblChooseDatabase.setFont(new Font("SansSerif", Font.PLAIN, 18));
		panel.add(lblChooseDatabase);

		String[] database = { "Select", "Oracle" };

		JPanel panel_1 = new JPanel();
		panel_1.setBounds(270, 60, 212, 38);
		panel_1.setLayout(new FlowLayout(FlowLayout.CENTER, 5, 5));

		JLabel label = new JLabel("Database");
		label.setHorizontalAlignment(SwingConstants.CENTER);
		panel_1.add(label);
		label.setFont(new Font("SansSerif", Font.PLAIN, 14));

		comboBox = new JComboBox(database);
		panel_1.add(comboBox);
		comboBox.setFont(new Font("SansSerif", Font.PLAIN, 14));

		JPanel panel_2 = new JPanel();
		panel_2.setBounds(189, 177, 398, 146);

		JPanel panel_3 = new JPanel();
		panel_3.setOpaque(false);
		panel_3.setBounds(323, 336, 103, 38);
		panel_3.setLayout(new FlowLayout(FlowLayout.CENTER, 5, 5));

		JButton btnNewButton = new JButton("Restore");
		btnNewButton.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				databaseSelection();
			}
		});
		btnNewButton.setFont(new Font("SansSerif", Font.PLAIN, 15));
		panel_3.add(btnNewButton);
		panel_2.setLayout(null);

		JLabel lblNewLabel_1 = new JLabel("Select Excel File");
		lblNewLabel_1.setFont(new Font("SansSerif", Font.PLAIN, 15));
		lblNewLabel_1.setBounds(140, 0, 121, 48);
		panel_2.add(lblNewLabel_1);

		JButton btnSelect = new JButton("select");
		btnSelect.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				JFileChooser chooser = new JFileChooser();
				chooser.setDialogTitle("Open File ");
				try {
					FileNameExtensionFilter f = new FileNameExtensionFilter("Excel File","xls");
					chooser.setFileFilter(f);
					int userSelection = chooser.showOpenDialog(contentPane);
					if (userSelection == JFileChooser.APPROVE_OPTION) {
						fileToOpen = chooser.getSelectedFile();
						// System.out.println("Save as file: " +
						// fileToOpen.getAbsolutePath()+".xls");
					}
				} catch (Exception w) {
					w.printStackTrace();
				}
				label1.setText(fileToOpen.toString());
				nameOfFile = fileToOpen.getAbsolutePath();
			}
		});
		btnSelect.setFont(new Font("SansSerif", Font.PLAIN, 14));
		btnSelect.setBounds(151, 59, 89, 23);
		panel_2.add(btnSelect);

		label1 = new JLabel("");
		label1.setFont(new Font("SansSerif", Font.PLAIN, 14));
		label1.setBounds(10, 91, 290, 23);
		panel_2.add(label1);
		contentPane.setLayout(null);
		contentPane.add(panel);
		contentPane.add(panel_1);
		contentPane.add(panel_2);
		contentPane.add(panel_3);

		JPanel panel_4 = new JPanel();
		panel_4.setBounds(96, 121, 557, 38);
		contentPane.add(panel_4);
		panel_4.setLayout(null);

		JLabel label_1 = new JLabel("Username");
		label_1.setFont(new Font("SansSerif", Font.PLAIN, 14));
		label_1.setBounds(31, 11, 93, 20);
		panel_4.add(label_1);

		usernameTextField = new JTextField();
		usernameTextField.setFont(new Font("SansSerif", Font.PLAIN, 14));
		usernameTextField.setColumns(10);
		usernameTextField.setBounds(134, 11, 131, 20);
		panel_4.add(usernameTextField);

		JLabel label_2 = new JLabel("Password");
		label_2.setFont(new Font("SansSerif", Font.PLAIN, 14));
		label_2.setBounds(295, 11, 77, 20);
		panel_4.add(label_2);

		passwordField = new JPasswordField();
		passwordField.setBounds(382, 13, 165, 20);
		panel_4.add(passwordField);

		JLabel label_3 = new JLabel("");
		label_3.setIcon(new ImageIcon(Restore1.class.getResource("/gui/database3.jpg")));
		label_3.setBounds(0, 0, 700, 427);
		contentPane.add(label_3);
	}

	public void databaseSelection() {
		String select = comboBox.getSelectedItem().toString();
		System.out.println(select);
		switch (select) {
		case "Oracle":
			getOracle();
			break;
		default:
			JOptionPane.showMessageDialog(null, "Please Select your database type");
		}
	}

	public void getOracle() {
		try {
			username = usernameTextField.getText();
			password = new String(passwordField.getPassword());
			Class.forName("oracle.jdbc.driver.OracleDriver");
			con = DriverManager.getConnection("jdbc:oracle:thin:@localhost:1521:xe", username, password);
			fileCheck();
		} catch (SQLException e) {
			JOptionPane.showMessageDialog(null, "Username and password is incorrect");
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public void fileCheck() {
		System.out.println("Name of file is " + nameOfFile);
		if (nameOfFile.isEmpty()) {
			JOptionPane.showMessageDialog(null, "Select Excel File to create backup ");
		} else {
			writingToDatabase();
		}
	}

	public void writingToDatabase() {
		try {
			name = fileToOpen.toString();
			System.out.println(name);
			Workbook workbook = Workbook.getWorkbook(new File(name));
			int numberOfSheets = workbook.getNumberOfSheets();
			System.out.println(numberOfSheets);
			sheetNames = workbook.getSheetNames();

			for (i = 0; i < sheetNames.length; i++) {
				try {
					pk = new ArrayList<>();
					sheet = workbook.getSheet(i);
					primaryConstraint = sheet.getCell(2, 0).getContents();
					totalPrimaryKey = Integer.parseInt(sheet.getCell(1, 0).getContents()) + 2;
					System.out.println("total primary key " + totalPrimaryKey);
					for (int c = 3; c <= totalPrimaryKey; c++) {
						primaryKey = sheet.getCell(c, 0).getContents();
						pk.add(primaryKey);
					}
					pkAdd = String.join(",", pk);
					System.out.println(pkAdd);
					pk = new ArrayList<>();
					list = new ArrayList<>();
					System.out.println("Sheet Name[" + i + "] = " + sheetNames[i]);
					totalNoOfRows = sheet.getRows();
					totalNoOfCols = Integer.parseInt(sheet.getCell(0, 0).getContents());
					columnName = null;
					columnNameType = null;
					length = null;
					for (c = 0; c < totalNoOfCols; c++) {
						length = sheet.getCell(c, 3).getContents();
						columnName = sheet.getCell(c, 2).getContents();
						columnNameType = sheet.getCell(c, 1).getContents();
						if (columnNameType.equals("VARCHAR2")) {
							add = columnName + " " + columnNameType + "(" + length + ")";
							list.add(add);
						} else if (columnNameType.equals("NUMBER")) {
							add = columnName + " " + columnNameType + "(" + length + ")";
							list.add(add);
						}
					}
					add = String.join(",", list);
					System.out.println(add);
					try {
						sql = "create table " + sheetNames[i] + "(" + add + ",CONSTRAINT " + primaryConstraint
								+ " primary key(" + pkAdd + "))";
						System.out.println(sql);
						pst = con.prepareStatement(sql);
						rs = pst.executeQuery(sql);
					} catch (SQLSyntaxErrorException e) {
						System.out.println("Name of table " + sheetNames[i] + " is already present merging the files");
						e.printStackTrace();
					}
					for (r = 4; r < totalNoOfRows; r++) {
						list = new ArrayList<>();
						try {
							for (c = 0; c < totalNoOfCols; c++) {
								columnNameType = sheet.getCell(c, 1).getContents();
								data = sheet.getCell(c, r).getContents();
								if (columnNameType.equals("VARCHAR2")) {
									add = "'" + data + "'";
									list.add(add);
								} else if (columnNameType.equals("NUMBER")) {
									if (data.isEmpty()) {
										data = "0";
										add = "'" + data + "'";
										list.add(add);
									} else {
										add = "'" + data + "'";
										list.add(add);
									}
								} else {
									add = data;
									list.add(add);
								}
							}
							add = String.join(",", list);
							sql = "insert into " + sheetNames[i] + " values(" + add + ")";
							System.out.println(sql);
							pst = con.prepareStatement(sql);
							pst.executeUpdate(sql);
							list = new ArrayList<>();
						} catch (SQLIntegrityConstraintViolationException e) {
							System.out.println("already present");
							e.printStackTrace();
						} catch (SQLSyntaxErrorException e) {
							e.printStackTrace();
						}
					}
				} catch (Exception e) {
					e.printStackTrace();
				}
			}

			con.close();
		} catch (FileNotFoundException e) {
			System.out.println(name + " file not found");
		} catch (Exception e) {
			e.printStackTrace();
			System.out.println("Unable to connect");
		}
		JOptionPane.showMessageDialog(null, "Restore Successfull");
	}
}

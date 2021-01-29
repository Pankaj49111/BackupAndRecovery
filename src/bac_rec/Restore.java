package bac_rec;

import java.awt.BorderLayout;
import java.awt.EventQueue;

import javax.swing.JFrame;
import javax.swing.JPanel;
import javax.swing.border.EmptyBorder;
import javax.swing.filechooser.FileNameExtensionFilter;

import jxl.*;

import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JTextField;
import javax.swing.JPasswordField;
import javax.swing.JButton;
import javax.swing.JComboBox;
import javax.swing.JFileChooser;

import java.awt.Font;
import javax.swing.DefaultComboBoxModel;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.io.*;
import java.sql.*;
import java.util.*;
import java.awt.event.ActionListener;
import java.awt.event.ActionEvent;

public class Restore extends JFrame {

	private JPanel contentPane;
	private JTextField user;
	private JPasswordField pass;
	private JTextField file_loc;
	
	private File fileToOpen = null;
	private String nameOfFile = "";
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
	

	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					Restore frame = new Restore();
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
	public Restore() {
		setTitle("Excel to Database");
		setResizable(false);
		setBounds(100, 100, 450, 300);
		contentPane = new JPanel();
		contentPane.setBorder(new EmptyBorder(5, 5, 5, 5));
		setContentPane(contentPane);
		contentPane.setLayout(null);
		
		JPanel panel = new JPanel();
		panel.setBounds(10, 11, 414, 51);
		contentPane.add(panel);
		panel.setLayout(null);
		
		JLabel lblUsername = new JLabel("Username");
		lblUsername.setFont(new Font("Microsoft YaHei Light", Font.BOLD, 12));
		lblUsername.setBounds(10, 11, 66, 30);
		panel.add(lblUsername);
		
		user = new JTextField();
		user.setFont(new Font("Times New Roman", Font.ITALIC, 12));
		user.setBounds(99, 16, 86, 20);
		panel.add(user);
		user.setColumns(10);
		
		JLabel lblPassword = new JLabel("Password");
		lblPassword.setFont(new Font("Microsoft JhengHei UI Light", Font.BOLD, 12));
		lblPassword.setBounds(218, 11, 66, 30);
		panel.add(lblPassword);
		
		pass = new JPasswordField();
		pass.setBounds(318, 16, 86, 20);
		panel.add(pass);
		
		JPanel panel_1 = new JPanel();
		panel_1.setBounds(10, 73, 414, 46);
		contentPane.add(panel_1);
		panel_1.setLayout(null);
		
		JLabel lblTypeOfDatabase = new JLabel("Type of Database Used:");
		lblTypeOfDatabase.setFont(new Font("Times New Roman", Font.BOLD | Font.ITALIC, 14));
		lblTypeOfDatabase.setBounds(66, 11, 149, 24);
		panel_1.add(lblTypeOfDatabase);
		
		JComboBox database = new JComboBox();
		database.setModel(new DefaultComboBoxModel(new String[] {"Select", "Oracle"}));
		database.setFont(new Font("Verdana", Font.BOLD | Font.ITALIC, 12));
		database.setBounds(225, 14, 120, 20);
		panel_1.add(database);
		
		JPanel panel_2 = new JPanel();
		panel_2.setBounds(10, 130, 414, 69);
		contentPane.add(panel_2);
		panel_2.setLayout(null);
		
		JLabel lblSelctFileTo = new JLabel("Selct file to be Restored");
		lblSelctFileTo.setFont(new Font("Times New Roman", Font.BOLD | Font.ITALIC, 14));
		lblSelctFileTo.setBounds(130, 0, 153, 27);
		panel_2.add(lblSelctFileTo);
		
		file_loc = new JTextField();
		file_loc.setFont(new Font("Verdana", Font.ITALIC, 15));
		file_loc.setText("Browse Location");
		file_loc.setEditable(false);
		file_loc.setBounds(44, 34, 177, 27);
		panel_2.add(file_loc);
		file_loc.setColumns(10);
		
		JButton browse = new JButton("Browse");
		browse.addActionListener(new ActionListener() {
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
				file_loc.setText(fileToOpen.toString());
				nameOfFile = fileToOpen.getAbsolutePath();
			}
		});
		browse.setFont(new Font("Times New Roman", Font.BOLD | Font.ITALIC, 14));
		browse.setBounds(231, 36, 89, 23);
		panel_2.add(browse);
		
		JPanel panel_3 = new JPanel();
		panel_3.setBounds(10, 210, 414, 40);
		contentPane.add(panel_3);
		panel_3.setLayout(null);
		
		JButton Restore = new JButton("Restore");
		Restore.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {						
						String select = database.getSelectedItem().toString();
						//System.out.println(select);
						switch (select) {
						case "Oracle":
							getInfo();
							break;
						default:
							JOptionPane.showMessageDialog(null, "Please Select your database type");
						}	
					}
		});
		Restore.setFont(new Font("Verdana", Font.BOLD | Font.ITALIC, 16));
		Restore.setBounds(131, 0, 124, 40);
		panel_3.add(Restore);
	}
	
	public void getInfo() {
		try {
			username = user.getText();
			password = new String(pass.getPassword());
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
		//System.out.println("Name of file is " + nameOfFile);
		if (nameOfFile.isEmpty()) {
			JOptionPane.showMessageDialog(null, "Select Excel File to create backup ");
		} else {
			write();
		}
	}
	

	public void write() {
		try {
			name = fileToOpen.toString();
			//System.out.println(name);
			Workbook workbook = Workbook.getWorkbook(new File(name));
			int numberOfSheets = workbook.getNumberOfSheets();
			//System.out.println(numberOfSheets);
			sheetNames = workbook.getSheetNames();

			for (i = 0; i < sheetNames.length; i++) {
				try {
					pk = new ArrayList<>();
					sheet = workbook.getSheet(i);
					primaryConstraint = sheet.getCell(2, 0).getContents();
					totalPrimaryKey = Integer.parseInt(sheet.getCell(1, 0).getContents()) + 2;
				 //System.out.println("total primary key " + totalPrimaryKey);
					for (int c = 3; c <= totalPrimaryKey; c++) {
						primaryKey = sheet.getCell(c, 0).getContents();
						pk.add(primaryKey);
					}
					pkAdd = String.join(",", pk);
			    //System.out.println(pkAdd);
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

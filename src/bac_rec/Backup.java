package bac_rec;

import java.awt.BorderLayout;
import java.awt.EventQueue;

import javax.swing.JFrame;
import javax.swing.JPanel;
import javax.swing.border.EmptyBorder;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JTextField;
import javax.swing.JPasswordField;
import java.awt.Font;
import javax.swing.JComboBox;
import javax.swing.JFileChooser;
import javax.swing.JButton;
import java.io.File;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.PreparedStatement;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.Scanner;
import java.util.regex.Pattern;

import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import java.awt.event.ActionListener;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.awt.event.ActionEvent;
import javax.swing.DefaultComboBoxModel;


public class Backup extends JFrame {

	private JPanel contentPane;
	private JTextField user;
	private JPasswordField pass;
	private JTextField file_name;
	private JTextField browse_loc;

	private String sql, name, getColumnName, ColumnType,z;
	private String t = "^([a-zA-Z0-9])+";
	private PreparedStatement pst;
	private ResultSet rs;
	private int table_count = 0, columns = 0, rows = 0, sheets = 1, j = 0,b;
	private ArrayList<String> list = new ArrayList<String>();
	private ResultSetMetaData rsmd;
	private Connection con;
	private String xls,username,password;
	private Scanner in = new Scanner(System.in);
	private Label label1,label2;
	
	
	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					Backup frame = new Backup();
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
	public Backup() {
		setTitle("Database to Excel");
		setResizable(false);
		setBounds(100, 100, 450, 300);
		contentPane = new JPanel();
		contentPane.setBorder(new EmptyBorder(5, 5, 5, 5));
		setContentPane(contentPane);
		contentPane.setLayout(null);
		
		JPanel panel = new JPanel();
		panel.setLayout(null);
		panel.setBounds(10, 11, 414, 51);
		contentPane.add(panel);
		
		JLabel lblUsername = new JLabel("  Username");
		lblUsername.setFont(new Font("Times New Roman", Font.BOLD, 14));
		lblUsername.setBounds(0, 12, 79, 30);
		panel.add(lblUsername);
		
		JLabel label_1 = new JLabel("  Password");
		label_1.setFont(new Font("Times New Roman", Font.BOLD, 14));
		label_1.setBounds(209, 12, 79, 30);
		panel.add(label_1);
		
		user = new JTextField();
		user.setFont(new Font("Times New Roman", Font.BOLD, 14));
		user.setColumns(10);
		user.setBounds(99, 16, 89, 24);
		panel.add(user);
		
		pass = new JPasswordField();
		pass.setFont(new Font("Times New Roman", Font.PLAIN, 14));
		pass.setBounds(306, 16, 98, 24);
		panel.add(pass);
		
		JPanel panel_1 = new JPanel();
		panel_1.setLayout(null);
		panel_1.setBounds(10, 73, 414, 36);
		contentPane.add(panel_1);
		
		JLabel lblSelectTheDatabase = new JLabel("  Select the Database");
		lblSelectTheDatabase.setFont(new Font("SansSerif", Font.BOLD | Font.ITALIC, 12));
		lblSelectTheDatabase.setBounds(0, 0, 138, 36);
		panel_1.add(lblSelectTheDatabase);
		
		JComboBox database = new JComboBox();
		database.setModel(new DefaultComboBoxModel(new String[] {"Select", "Oracle"}));
		database.setFont(new Font("Verdana", Font.BOLD | Font.ITALIC, 12));
		database.setBounds(266, 6, 138, 25);
		panel_1.add(database);
		
		JPanel panel_2 = new JPanel();
		panel_2.setBounds(10, 120, 414, 84);
		contentPane.add(panel_2);
		panel_2.setLayout(null);
		
		JLabel lblFileNameTo = new JLabel(" File Name to be created");
		lblFileNameTo.setFont(new Font("SansSerif", Font.BOLD, 13));
		lblFileNameTo.setBounds(7, 0, 198, 36);
		panel_2.add(lblFileNameTo);
		
		file_name = new JTextField();
		file_name.setColumns(10);
		file_name.setBounds(267, 5, 137, 29);
		panel_2.add(file_name);
		
		browse_loc = new JTextField();
		browse_loc.setFont(new Font("Yu Gothic", Font.ITALIC, 16));
		browse_loc.setText(" Browse Location");
		browse_loc.setToolTipText("");
		browse_loc.setEditable(false);
		browse_loc.setColumns(10);
		browse_loc.setBounds(7, 44, 210, 29);
		panel_2.add(browse_loc);
		
		JButton browse = new JButton("Browse");
		browse.addMouseListener(new MouseAdapter() {
			@Override
			public void mouseClicked(MouseEvent arg0) {
			
				JFileChooser Filechoose=new JFileChooser();
		        Filechoose.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
                int retval=Filechoose.showOpenDialog(null);
                if (retval == JFileChooser.APPROVE_OPTION) {
                	browse_loc.setText(Filechoose.getSelectedFile().toString());
                    z=browse_loc.getText();
                }
			}
			});
	
		
		browse.setFont(new Font("Times New Roman", Font.BOLD | Font.ITALIC, 14));
		browse.setBounds(267, 45, 137, 25);
		panel_2.add(browse);
		
		JPanel panel_4 = new JPanel();
		panel_4.setBounds(10, 215, 414, 35);
		contentPane.add(panel_4);
		panel_4.setLayout(null);
		
		JButton fetch = new JButton("Backup");
		fetch.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
			
				String select = database.getSelectedItem().toString();
				switch(select){
					case "Oracle":
						getInfo();
						break;
						default:
						JOptionPane.showMessageDialog(null,"Please select your Database type","Database problem",JOptionPane.WARNING_MESSAGE);
				}
				
			}
		});
		fetch.setFont(new Font("Times New Roman", Font.BOLD | Font.ITALIC, 18));
		fetch.setBounds(123, 0, 132, 35);
		panel_4.add(fetch);
	}

	public void create() {
		//xls=file_name.getText();
		try {
			WritableWorkbook workbook = Workbook.createWorkbook(new File(z+"\\"+xls+".xls"));
			workbook.createSheet("Sheet", 1);
			workbook.write();
			workbook.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	public void write() {
		//xls=file_name.getText();
		try {
			sql = "select * from user_tables";
			pst = con.prepareStatement(sql);
			rs = pst.executeQuery(sql);
			while (rs.next()) {
				name = rs.getString(1);
				if (Pattern.matches(t, name)) {
					list.add(name);
				}
			}
			try {
				//Workbook wb = Workbook.getWorkbook(new File(nameOfFile));
				WritableWorkbook copy = Workbook.createWorkbook(new File(z+"\\"+xls+".xls"));
				for (String a : list) {
					System.out.println(a);
					try {
						String s1 = a;
						WritableSheet copySheet = copy.createSheet(s1, j);
						sql="SELECT constraint_name FROM all_cons_columns WHERE constraint_name = (SELECT constraint_name FROM all_constraints WHERE UPPER(table_name) = UPPER('"+a+"') AND CONSTRAINT_TYPE = 'P')";
						pst =  con.prepareStatement(sql);
						rs = pst.executeQuery(sql);
						b=1;
						while(b==1 & rs.next()){
							//System.out.println(rs.getString(1));
							label1 = new Label(2,0,rs.getString(1));
							copySheet.addCell(label1);
							b++;
						}
						sql = "SELECT column_name FROM all_cons_columns WHERE constraint_name = (SELECT constraint_name FROM all_constraints WHERE UPPER(table_name) = UPPER('"+a+"') AND CONSTRAINT_TYPE = 'P')";
						pst =  con.prepareStatement(sql);
						rs = pst.executeQuery(sql);
						rsmd = rs.getMetaData();
						int k=3;
						while(rs.next()){
							 label1 = new Label(k,0,rs.getString(1));
							copySheet.addCell(label1);
							k++;
						}
						label2 = new Label(1,0,String.valueOf(k-3));
						copySheet.addCell(label2);
						sql = "select * from " + a;
						pst = con.prepareStatement(sql);
						rs = pst.executeQuery(sql);
						rsmd = rs.getMetaData();
						//System.out.println("Total columns: " + rsmd.getColumnCount());
						columns = rsmd.getColumnCount();
						label1 = new Label(0,0,String.valueOf(columns));
						copySheet.addCell(label1);
						for (int i = 0; i < columns; i++) {
							//System.out.println("Column Name of 1st column: " + rsmd.getColumnName(i + 1));
							//System.out.println("Column Type Name of 1st column: " + rsmd.getColumnTypeName(i + 1));
							String columnTypeName = rsmd.getColumnTypeName(i + 1);
							//System.out.println(columnTypeName);
							label1 = new Label(i, 1, columnTypeName);
							copySheet.addCell(label1);
						}
						for (int i = 0; i < columns; i++) {
							String columnName = rsmd.getColumnName(i + 1);
							 label1 = new Label(i, 2, columnName);
							copySheet.addCell(label1);
						}
						for (int i = 0; i < columns; i++) {
							String columnName = Integer.toString(rsmd.getColumnDisplaySize(i + 1));
							 label1 = new Label(i, 3, columnName);
							copySheet.addCell(label1);
						}
						int c = 0, r = 4;
						sql = "SELECT * FROM " + a;
						pst = con.prepareStatement(sql);
						for (int i = 0; i < columns; i++) {
							rs = pst.executeQuery(sql);
							r = 4;
							while (rs.next()) {
							 label1 = new Label(i, r, rs.getString(i + 1));
								copySheet.addCell(label1);
								r++;
							}
						}
						j++;
					} catch (Exception e) {
						// TODO: handle exception
						e.printStackTrace();
					}
					sheets++;
				}
				copy.write();
				copy.close();
				JOptionPane.showMessageDialog(null,"data saved successfully");
			} catch (Exception e) {
				e.printStackTrace();
			}
		} catch (Exception e) {
			// TODO: handle exception
			e.printStackTrace();
		}
	}

	public void getInfo(){
		username = user.getText();
		password = new String(pass.getPassword());
	connect();
	}
	public void connect(){
		try {
			Class.forName("oracle.jdbc.driver.OracleDriver");
			con = DriverManager.getConnection("jdbc:oracle:thin:@localhost:1521:xe", username, password);
			checkFile();
			} catch (ClassNotFoundException e) {
				JOptionPane.showMessageDialog(null,"Oracle not Installed..!!!","backup problem",JOptionPane.WARNING_MESSAGE);
			} catch (SQLException e) {
		        JOptionPane.showMessageDialog(null,"Either Username or password \n is incorrect..!!!","Connection Problem",JOptionPane .ERROR_MESSAGE);
			}
		catch (Exception e) {
			JOptionPane.showMessageDialog(null,"Error in connectivity..!!!");
		}
	}
	
	public void checkFile(){

		xls=file_name.getText();
		File file = new File(z+"\\"+xls+".xls");
		if (!file.exists()) {
			create();
			write();
		} else {	
			write();
		}
	}
	
	
	/*public static void main(String[] args) {
		// TODO Auto-generated method stub
		Db2exx db = new Db2exx();
		db.connection();
	}
	*/
	
}

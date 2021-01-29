package bac_rec;

import java.awt.BorderLayout;
import java.awt.EventQueue;

import javax.swing.JFrame;
import javax.swing.JPanel;
import javax.swing.border.EmptyBorder;
import javax.swing.JButton;
import java.awt.event.ActionListener;
import java.awt.event.ActionEvent;
import javax.swing.JTextField;
import javax.swing.JLabel;

public class Opening extends JFrame {

	private JPanel contentPane;

	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					Opening frame = new Opening();
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
	public Opening() {
		setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		setBounds(100, 100, 450, 300);
		contentPane = new JPanel();
		contentPane.setBorder(new EmptyBorder(5, 5, 5, 5));
		setContentPane(contentPane);
		contentPane.setLayout(null);
		
		JPanel panel = new JPanel();
		panel.setBounds(10, 11, 414, 59);
		contentPane.add(panel);
		panel.setLayout(null);
		
		JLabel lb1 = new JLabel("What operation you want to perform");
		lb1.setBounds(130, 11, 217, 37);
		panel.add(lb1);
		
		JPanel panel_1 = new JPanel();
		panel_1.setBounds(10, 123, 414, 85);
		contentPane.add(panel_1);
		panel_1.setLayout(null);
		
		JButton b1 = new JButton("Database to Excel ");
		b1.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				
				Backup bac=new Backup();
				bac.setVisible(true);
			}
		});
		b1.setBounds(10, 11, 159, 59);
		panel_1.add(b1);
		
		JButton b2 = new JButton("Excel to Database");
		b2.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
			
				Restore re=new Restore();
				re.setVisible(true);
			}
		});
		b2.setBounds(245, 11, 159, 59);
		panel_1.add(b2);
	}
}

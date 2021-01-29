package bac_rec;

import javax.swing.ImageIcon;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JWindow;
import javax.swing.SwingConstants;

public class SplashScreen{

	public static void main(String[] args) {
		// TODO Auto-generated method stub
		JWindow window=new JWindow();
		window.getContentPane().add(new JLabel("",new ImageIcon("Double Ring.gif"),SwingConstants.CENTER));
		window.setBounds(450, 300, 320, 240);
		window.setVisible(true);
		try {
			Thread.sleep(5000);
			Opening op=new Opening();
			op.setVisible(true);
		}catch(InterruptedException e) {
			window.dispose();
		}
		window.setVisible(false);
		window.dispose();
	}

}

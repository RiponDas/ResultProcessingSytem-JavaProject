package result;
import java.awt.GridLayout;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.IOException;

import javax.swing.ImageIcon;
import javax.swing.JButton;
import javax.swing.JFrame;
import javax.swing.JPanel;
import javax.swing.JTextArea;
import javax.swing.JLabel;
import javax.swing.JTextField;
import javax.swing.border.TitledBorder;

import result.GridPanel.eventsub4;
import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

import javax.swing.border.LineBorder;
import java.awt.Color;
import java.util.concurrent.ExecutionException;

public class MainInterface extends JFrame {
	JPanel workerJPanel,workerJPanel2,workerJPanel3,workerJPanel4,resultSheet,headImage,inputPane,bundlePanel;
	JTextField nt11,nt12,nt13,nt14,nt15,nt21,nt22,nt23,nt24,nt25,nt31,nt32,nt33,nt34,nt35,nt41,nt42,nt43,nt44,nt45,tm1,tm2,tm3,tm4;
	JButton sub1,sub2,sub3,sub4,result,btn ;
	//JTextArea ta1;
	
	public MainInterface()
	 {
	 super( "Result Processing System" );
	 setLayout( new GridLayout( 3, 1, 1, 8 ) );
	 setBorder(new LineBorder( Color.BLACK ) );
	//headImage
	 headImage = new JPanel(new GridLayout(1, 1, 1,1));
	 headImage.add( new JLabel(new ImageIcon(getClass().getResource("img.jpg"))));
	 
	 result = new JButton(new ImageIcon(getClass().getResource("result.jpg")));
				 JPanel grid = new GridPanel(new JPanel());
		
		 
		 add(headImage);
		 add(grid);
		 add(result);
		 
		 result.addActionListener(new ActionListener()
				 
				 {
			 
			 public void actionPerformed(ActionEvent event)
		        {
				 
				 ExcelResult obj=new ExcelResult(MainInterface.this);
				 obj.setVisible(true);
				 obj.setSize(925,450);
				 obj.setDefaultCloseOperation(JFrame.HIDE_ON_CLOSE);
				 
		        }
				 }
				 
				 );
		 
		 
		 setSize( 925, 590 );
		 setVisible( true );
		 } // end constructor
	
	     
		 private void setBorder(LineBorder lineBorder) {
		// TODO Auto-generated method stub
		
	}


		public static void main( String[] args )
		 {
			MainInterface app = new MainInterface();
		  app.setDefaultCloseOperation( EXIT_ON_CLOSE );
		 } // end main
}

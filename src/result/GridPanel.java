package result;
import java.io.File;
import java.io.IOException;
import java.util.Date;
import java.util.Formatter;

import jxl.*;
import jxl.write.*;
import jxl.write.Boolean;
import jxl.write.Number;
import jxl.write.NumberFormat;

import jxl.write.biff.RowsExceededException;

import java.awt.*;
import java.awt.event.*;
import javax.swing.*;
import javax.swing.event.*;
import java.io.*;
import javax.swing.filechooser.FileFilter;

import jxl.Workbook;
import jxl.write.Label;
import jxl.write.Number;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

import javax.swing.JFileChooser.*;
import java.awt.GridLayout;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.IOException;

import javax.swing.JButton;
import javax.swing.JFrame;
import javax.swing.JPanel;
import javax.swing.JTextArea;
import javax.swing.JLabel;
import javax.swing.JTextField;
import javax.swing.border.TitledBorder;

import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;


import javax.swing.border.LineBorder;
import java.awt.Color;
import java.util.concurrent.ExecutionException;

public class GridPanel  extends JPanel{
	JPanel workerJPanel,workerJPanel2,workerJPanel3,workerJPanel4,resultSheet,headImage,inputPane,bundlePanel;
	JTextField ntname,nt11,nt12,nt13,nt14,nt15,nt16,nt21,nt22,nt23,nt24,nt25,nt31,nt32,nt33,nt34,nt35,nt41,nt42,nt43,nt44,nt45,tm1,tm2,tm3,tm4;
	JButton sub4,result ;
	///JTextArea ta1;
	
	public GridPanel(JPanel panel)
	 {
	 //super();
	 setLayout( new GridLayout( 1, 4, 10,40 ) );
	
	
	///Panel --1
	 workerJPanel = new JPanel( new GridLayout( 6, 2, 5, 1 ) );
	 ntname = new JTextField();
	 nt11 = new JTextField();
	 nt12 = new JTextField();	 
	 nt13 = new JTextField();	
	 nt14 = new JTextField();	
	 nt15 = new JTextField();
		
		//Panel--2
		 workerJPanel2 =
				 new JPanel( new GridLayout( 6, 2, 5, 1 ) );		 
		 nt21 = new JTextField();		
		 nt22 = new JTextField();		 
		 nt23 = new JTextField();		
		 nt24 = new JTextField();		
		 nt25 = new JTextField();

		//Panel --3
		 workerJPanel3 =
				 new JPanel( new GridLayout( 6, 2, 5, 1 ) );		 
		 nt31 = new JTextField();	
		 nt32 = new JTextField();		 
		 nt33 = new JTextField();	
		 nt34 = new JTextField();		
		 nt35 = new JTextField();
		//Panel --4
		 workerJPanel4 =
				 new JPanel( new GridLayout( 6, 2, 5, 1 ) );		 
		 nt41 = new JTextField();		
		 nt42 = new JTextField();		 
		 nt43 = new JTextField();	
		 nt44 = new JTextField();		
		 nt45 = new JTextField();
		 sub4 = new JButton( "Submit" );
		 
		 //Panel--1
		 workerJPanel.setBorder( new TitledBorder(
				 new LineBorder( Color.BLACK ), "Student Info" ) );
		         workerJPanel.add( new JLabel( "Student Name:" ) );
		         workerJPanel.add( ntname );
				 workerJPanel.add( new JLabel( "Class Roll:" ) );
				 workerJPanel.add( nt11 );
				 workerJPanel.add( new JLabel( "Exam Roll:" ) );
				 workerJPanel.add( nt12 );
				 workerJPanel.add( new JLabel( "Session:" ) );
				 workerJPanel.add( nt13 );
				 workerJPanel.add( new JLabel( "Semester:" ) );
				 workerJPanel.add( nt14 );

				
		 //panel--2
		 workerJPanel2.setBorder( new TitledBorder(
				 new LineBorder( Color.BLACK ), "Course Info" ) );
				 workerJPanel2.add( new JLabel( "Course Code:" ) );
				 workerJPanel2.add( nt21 );
				 workerJPanel2.add( new JLabel( "Internal Marks:" ) );
				 workerJPanel2.add( nt22 );
				 workerJPanel2.add( new JLabel( "1st Exm Marks:" ) );
				 workerJPanel2.add( nt23 );
				 workerJPanel2.add( new JLabel( "2nd Exm Marks:" ) );
				 workerJPanel2.add( nt24 );
				 workerJPanel2.add( new JLabel( "3rd Exm Marks:" ) );
				 workerJPanel2.add( nt25 );
				
				
		//Panel--3
				 workerJPanel3.setBorder( new TitledBorder(
						 new LineBorder( Color.BLACK ), "Course Info" ) );
						 workerJPanel3.add( new JLabel( "Course Code:" ) );
						 workerJPanel3.add( nt31 );
						 workerJPanel3.add( new JLabel( "Internal Marks:" ) );
						 workerJPanel3.add( nt32 );
						 workerJPanel3.add( new JLabel( "1st Exm Marks:" ) );
						 workerJPanel3.add( nt33 );
						 workerJPanel3.add( new JLabel( "2nd Exm Marks:" ) );
						 workerJPanel3.add( nt34 );
						 workerJPanel3.add( new JLabel( "3rd Exm Marks:" ) );
						 workerJPanel3.add( nt35 );
						
		//Panel--4
						 workerJPanel4.setBorder( new TitledBorder(
								 new LineBorder( Color.BLACK ), "Course Info" ) );
								 workerJPanel4.add( new JLabel( "Course Code:" ) );
								 workerJPanel4.add( nt41 );
								 workerJPanel4.add( new JLabel( "Internal Marks:" ) );
								 workerJPanel4.add( nt42 );
								 workerJPanel4.add( new JLabel( "1st Exm Marks:" ) );
								 workerJPanel4.add( nt43 );
								 workerJPanel4.add( new JLabel( "2nd Exm Marks:" ) );
								 workerJPanel4.add( nt44 );
								 workerJPanel4.add( new JLabel( "3rd Exm Marks:" ) );
								 workerJPanel4.add( nt45 );
								 workerJPanel4.add( sub4 );
		 
		 
	     add( workerJPanel );
		 add( workerJPanel2 );
		 add( workerJPanel3 );
		 add( workerJPanel4 );
		 
		 setSize( 1000, 250 );
		 setVisible( true );
		 
		 
		 eventsub4 esub4=new eventsub4();
		 sub4.addActionListener(esub4);
		 } // end constructor
	
	///panel 4
    public class eventsub4 implements ActionListener
{
        public void actionPerformed(ActionEvent esub4)
        {

           try {
        	   File exlFile = new File("E:/Result.xls");
        	   WritableWorkbook writableWorkbook = Workbook
        	   .createWorkbook(exlFile);

        	   WritableSheet writableSheet = writableWorkbook.createSheet(
        	   "CSE-1201", 0);


        	   //Calculation
        	   //for course-1
        	   int diff1;
        	   double marks1,tmarks1,grade1;
        	   
        	   int thirdExmMarks1=0;
        	   int intMarks1 = Integer.parseInt(nt22.getText());
        	   int fstExmMarks1 = Integer.parseInt(nt23.getText());
        	   int sndExmMarks1 = Integer.parseInt(nt24.getText());
        	   thirdExmMarks1 = Integer.parseInt(nt25.getText());
        	   
        	   if(fstExmMarks1>sndExmMarks1)
        		   diff1=fstExmMarks1-sndExmMarks1;
        	   else
        		   diff1=sndExmMarks1-fstExmMarks1;
        	   if(diff1>=6){
        		   marks1=(fstExmMarks1+sndExmMarks1+thirdExmMarks1)/3;
        	   }
        	   else
        		   marks1=(fstExmMarks1+sndExmMarks1)/2;
        	   
        	   tmarks1 = marks1+intMarks1;
        	   grade1=grade(tmarks1);
        	   
        	   String marks1s= Double.toString(marks1);
        	   String grade1s= Double.toString(grade1);
        	   //for Course-2
        	   int diff2;
        	   double marks2,tmarks2,grade2;
        	   
        	   int thirdExmMarks2=0;
        	   int intMarks2 = Integer.parseInt(nt32.getText());
        	   int fstExmMarks2 = Integer.parseInt(nt33.getText());
        	   int sndExmMarks2 = Integer.parseInt(nt34.getText());
        	   thirdExmMarks2 = Integer.parseInt(nt35.getText());
        	   
        	   if(fstExmMarks2>sndExmMarks2)
        		   diff2=fstExmMarks2-sndExmMarks2;
        	   else
        		   diff2=sndExmMarks2-fstExmMarks2;
        	   if(diff2>=6){
        		   marks2=(fstExmMarks2+sndExmMarks2+thirdExmMarks2)/3;
        	   }
        	   else
        		   marks2=(fstExmMarks2+sndExmMarks2)/2;
        	   
        	   tmarks2 = marks2+intMarks2;
        	   grade2=grade(tmarks2);
        	   
        	   String marks2s= Double.toString(marks2);
        	   String grade2s= Double.toString(grade2);
        	   //for course-3
        	   int diff3;
        	   double marks3,tmarks3,grade3;
        	   
        	   int thirdExmMarks3=0;
        	   int intMarks3 = Integer.parseInt(nt42.getText());
        	   int fstExmMarks3 = Integer.parseInt(nt43.getText());
        	   int sndExmMarks3 = Integer.parseInt(nt44.getText());
        	   thirdExmMarks3 = Integer.parseInt(nt45.getText());
        	   
        	   if(fstExmMarks3>sndExmMarks3)
        		   diff3=fstExmMarks3-sndExmMarks3;
        	   else
        		   diff3=sndExmMarks3-fstExmMarks3;
        	   if(diff3>=6){
        		   marks3=(fstExmMarks3+sndExmMarks3+thirdExmMarks3)/3;
        	   }
        	   else
        		   marks3=(fstExmMarks3+sndExmMarks3)/2;
        	   
        	   tmarks3 = marks3+intMarks3;
        	   grade3=grade(tmarks3);
        	   
        	   String marks3s= Double.toString(marks3);
        	   String grade3s= Double.toString(grade3);
        	   //More Calculation
        	   double tgrade = grade1+grade2+grade3;
        	   String tgrade3s= Double.toString(tgrade);
        	   double cgpa = (tgrade*4.00)/12.00;
        	   Formatter fmt = new Formatter();
        	  String str = String.format("%.2f", cgpa);
        	
        	   //Create Cells with contents of different data types.
        	   
        	   Label head1 = new Label(0, 0,"Student Info");
        	   Label head2 = new Label(1, 0,"Course Code");
        	   Label head3 = new Label(2, 0,"Internal Marks");
        	   Label head4 = new Label(3, 0,"Final Exam Marks");
        	   Label head5 = new Label(4, 0,"Grade");
        	   Label head6 = new Label(5, 0,"CGPA");
        	   
        	   
        	   Label ll1 = new Label(0, 1,ntname.getText());
        	   Label ll2 = new Label(1, 1,nt21.getText());
        	   Label ll3 = new Label(2, 1,nt22.getText());
        	   Label ll4 = new Label(3, 1,marks1s);
        	   Label ll5 = new Label(4, 1,grade1s);
        	   Label ll6 = new Label(5, 1,str);
        	   fmt.close();
        	   Label l21 = new Label(0, 2,nt11.getText());
        	   Label l22 = new Label(1, 2,nt31.getText());
        	   Label l23 = new Label(2, 2,nt32.getText());
        	   Label l24 = new Label(3, 2,marks2s);
        	   Label l25 = new Label(4, 2,grade2s);
        	   
        	   
        	   Label l31 = new Label(0, 3,nt13.getText());
        	   Label l32 = new Label(1, 3,nt41.getText());
        	   Label l33 = new Label(2, 3,nt42.getText());
        	   Label l34 = new Label(3, 3,marks3s);
        	   Label l35 = new Label(4, 3,grade3s);
        	   
        	   Label l41 = new Label(3, 4,"Total Earned Credit");
        	   Label l42 = new Label(4, 4,tgrade3s);
        	   
        	  
        	   //Add the created Cells to the sheet
        	   writableSheet.addCell(head1);
        	   writableSheet.addCell(head2);
        	   writableSheet.addCell(head3);
        	   writableSheet.addCell(head4);
        	   writableSheet.addCell(head5);
        	   writableSheet.addCell(head6);
        	   
        	   writableSheet.addCell(ll1);
        	   writableSheet.addCell(ll2);
        	   writableSheet.addCell(ll3);
        	   writableSheet.addCell(ll4);
        	   writableSheet.addCell(ll5);
        	   writableSheet.addCell(ll6);

        	   writableSheet.addCell(l21);
        	   writableSheet.addCell(l22);
        	   writableSheet.addCell(l23);
        	   writableSheet.addCell(l24);
        	   writableSheet.addCell(l25);
        	   
        	   writableSheet.addCell(l31);
        	   writableSheet.addCell(l32);
        	   writableSheet.addCell(l33);
        	   writableSheet.addCell(l34);
        	   writableSheet.addCell(l35);
        	   
        	   writableSheet.addCell(l41);
        	   writableSheet.addCell(l42);
        	  
        	
        	   //Write and close the workbook
        	   writableWorkbook.write();
        	   writableWorkbook.close();

        	   } catch (IOException e) {
        	   e.printStackTrace();
        	   } catch (RowsExceededException e) {
        	   e.printStackTrace();
        	   } catch (WriteException e) {
        	   e.printStackTrace();
        	   }
        }
}
    public double grade (double marks){
    	double g = 0;
    	if(marks<=100 && marks>=0){
			if(marks>=40 && marks<=44)
				g= 2.0;
			if(marks>=45 && marks<=49)
				g=2.25;
			if(marks>=50 && marks<=54)
				g=2.50;
			if(marks>=55 && marks<=59)
				g= 2.75;
			if(marks>=60 && marks<=64)
				g=3.0;
			if(marks>=65 && marks<=69)
				g= 3.25;
			if(marks>=70 && marks<=74)
				g= 3.50;
			if(marks>=75 && marks<=79)
				g= 3.75;
			if(marks>=80 && marks<=100)
				g=4.0;
			if(marks>=0 && marks<40)
				g= 0.0;
    }
		return g;
    }
}

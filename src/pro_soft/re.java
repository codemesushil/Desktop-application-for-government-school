package pro_soft;
import java.awt.Color;   
import java.awt.Font; 
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.IOException;
import javax.swing.JButton;
import javax.swing.JComboBox;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JSpinner;
import javax.swing.JTextField;

import jxl.CellView;
import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.format.UnderlineStyle;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.Number;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;


public class re implements ActionListener
{

	File file1,file;
	
    private WritableCellFormat timesBoldUnderline;
    private WritableCellFormat times;
    private String inputFile;
    JFrame f1;
    JButton b1,b2;
    JLabel l1,l2,l3,l4,l5,l6,l7,l8,l9,l10,l11,l12,l13,l14,l15;
    JLabel l16,l17,l18,l19,l20,l21,l22,l23,l24,l25,l26,l27,l28,l29,l30;
    JTextField t1,t2,t3,t4,t5,t6,t7,t8,t9,t10,t11,t12,t13,t14,t15;
    JTextField t16,t17,t18,t19,t20,t21,t22,t23,t24,t25,t26,t27,t28,t29,t30,t31,t32,t33,t34,t35,t36;
    JSpinner sp,sp1,sp2,sp3,sp4,sp5,sp6,sp7,sp8,sp9,sp10,sp11,sp12,sp13,sp15;
    JComboBox<String> c1;
    
    double sum=0,sum1=0,sum2=0,sum3=0,sum4=0,sum5=0,sum6=0,sum7=0,sum8=0,sum9=0,sum10=0,sum11=0,sum12=0,sum13=0,sum14=0;
    public re()
    {
    	c1=new JComboBox<>();
    	for(int i=1;i<32;i++)
    	{
    		String str=Integer.toString(i);
    	c1.addItem(str);
    	}
    	c1.setBounds(750, 20, 90, 30);
    	f1=new JFrame("सरस्वती विद्यालय लोणखेडा (Made By Sushil Chaudhari)");
    	
    	sp=new JSpinner();
    	sp.setValue(0);
    	sp.setBounds(280, 250, 60, 30);
    	sp1=new JSpinner();
    	sp1.setValue(0);
    	sp1.setBounds(280, 300, 60, 30);
    	sp2=new JSpinner();
    	sp2.setValue(0);
    	sp2.setBounds(280, 350, 60, 30);
    	sp3=new JSpinner();
    	sp3.setValue(0);
    	sp3.setBounds(280, 400, 60, 30);
    	sp4=new JSpinner();
    	sp4.setValue(0);
    	sp4.setBounds(280, 450, 60, 30);
    	sp5=new JSpinner();
    	sp5.setValue(0);
    	sp5.setBounds(280, 500, 60, 30);
    	sp6=new JSpinner();
    	sp6.setValue(0);
    	sp6.setBounds(280, 550, 60, 30);
    	sp7=new JSpinner();
    	sp7.setValue(0);
    	sp7.setBounds(280, 600, 60, 30);
    	sp8=new JSpinner();
    	sp8.setValue(0);
    	sp8.setBounds(850, 250, 60, 30);
    	sp9=new JSpinner();
    	sp9.setValue(0);
    	sp9.setBounds(850, 300, 60, 30);
    	sp10=new JSpinner();
    	sp10.setValue(0);
    	sp10.setBounds(850, 350, 60, 30);
    	sp11=new JSpinner();
    	sp11.setValue(0);
    	sp11.setBounds(850, 400, 60, 30);
    	sp12=new JSpinner();
    	sp12.setValue(0);
    	sp12.setBounds(850, 450, 60, 30);
    	sp13=new JSpinner();
    	sp13.setValue(0);
    	sp13.setBounds(850, 500, 60, 30);
//    	sp14=new JSpinner();
//    	sp14.setValue(2);
//    	sp14.setBounds(850, 500, 60, 30);
    	
    	l24=new JLabel("कृपया आधी तक्ता तयार करा!");
    	l24.setFont(new Font("Serif", Font.BOLD, 45));
    	l24.setBounds(30, 1, 490, 60);
    	l2=new JLabel("मुगडाळ");
    	l2.setFont(new Font("Serif", Font.BOLD, 20));
    	l2.setBounds(200, 250, 200, 30);
    	l3=new JLabel("तूरडाळ");
    	l3.setFont(new Font("Serif", Font.BOLD, 20));
    	l3.setBounds(200, 300, 200, 30);
    	l4=new JLabel("मसुरडाळ");
    	l4.setFont(new Font("Serif", Font.BOLD, 20));
    	l4.setBounds(200, 350, 200, 30);
    	l5=new JLabel("मटकी");
    	l5.setFont(new Font("Serif", Font.BOLD, 20));
    	l5.setBounds(200,400, 200, 30);
    	l6=new JLabel("मुग");
    	l6.setFont(new Font("Serif", Font.BOLD, 20));
    	l6.setBounds(200, 450, 200, 30);
    	l7=new JLabel("चवळी ");
    	l7.setFont(new Font("Serif", Font.BOLD, 20));
    	l7.setBounds(200, 500, 200, 30);
    	l8=new JLabel("हरभरा");
    	l8.setFont(new Font("Serif", Font.BOLD, 20));
    	l8.setBounds(200, 550, 200, 30);
    	l9=new JLabel("वाटाणा");
    	l9.setFont(new Font("Serif", Font.BOLD, 20));
    	l9.setBounds(200, 600, 200, 30);
    	l10=new JLabel("जिरे");
    	l10.setFont(new Font("Serif", Font.BOLD, 20));
    	l10.setBounds(720, 250, 200, 30);
    	l11=new JLabel("मोहरी");
    	l11.setFont(new Font("Serif", Font.BOLD, 20));
    	l11.setBounds(720, 300, 200, 30);
    	l12=new JLabel("हळद");
    	l12.setFont(new Font("Serif", Font.BOLD, 20));
    	l12.setBounds(720, 350, 200, 30);
    	l13=new JLabel("मिरची पावडर");
    	l13.setFont(new Font("Serif", Font.BOLD, 20));
    	l13.setBounds(720, 400, 200, 30);
    	l14=new JLabel("सोयाबीन तेल");
    	l14.setFont(new Font("Serif", Font.BOLD, 20));
    	l14.setBounds(720, 450, 200, 30);
    	l15=new JLabel("भाजीपाला");
    	l15.setFont(new Font("Serif", Font.BOLD, 20));
    	l15.setBounds(720, 500, 200, 30);
    	l16=new JLabel("पूरक आहार");
    	l16.setFont(new Font("Serif", Font.BOLD, 20));
    	l16.setBounds(720, 550, 200, 30);
    	l17=new JLabel("");
    	l17.setFont(new Font("Serif", Font.BOLD, 20));
    	l17.setBounds(900, 150, 300, 30);
    	l21=new JLabel("(in grams)");
    	l21.setBounds(280, 225, 200, 30);
    	l22=new JLabel("(in grams)");
    	l22.setBounds(850, 225, 200, 30);
    	l23=new JLabel("दिनांक निवडा");
    	l23.setFont(new Font("Serif", Font.BOLD, 25));
    	l23.setBounds(600, 20, 240, 30);
    	l25=new JLabel("");
    	l25.setFont(new Font("Serif", Font.BOLD, 25));
    	l25.setBounds(10, 650, 750, 30);
    	l26=new JLabel("खर्च");
    	l26.setFont(new Font("Serif", Font.BOLD, 20));
    	l26.setBounds(20, 350, 200, 30);
    	t5=new JTextField("000");
    	t5.setBounds(60, 350, 90, 30);
    	l27=new JLabel("मासिक शिल्लक साहित्य");
    	l27.setBounds(370, 225, 200, 30);
    	l27.setFont(new Font("Serif", Font.BOLD, 14));
    	l28=new JLabel("प्राप्त साहित्य");
    	l28.setBounds(530, 225, 200, 30);
    	l28.setFont(new Font("Serif", Font.BOLD, 14));
    	l29=new JLabel("मासिक शिल्लक साहित्य");
    	l29.setBounds(940, 225, 200, 30);
    	l29.setFont(new Font("Serif", Font.BOLD, 14));
    	l30=new JLabel("प्राप्त साहित्य");
    	l30.setBounds(1100, 225, 200, 30);
    	l30.setFont(new Font("Serif", Font.BOLD, 14));
    	
    	b1=new JButton("तक्ता तयार करा");
    	b1.setBounds(20, 100, 150, 150);
    	b1.setFont(new Font("Serif", Font.BOLD, 17));
    	b1.setBackground(Color.GREEN);
    	b2=new JButton("माहिती भरा");
    	b2.setBounds(900, 60, 200, 80);
    	b2.setFont(new Font("Serif", Font.BOLD, 17));
        	
    	l18=new JLabel("चालू महिन्याची पटसंख्या");
    	l18.setFont(new Font("Serif", Font.BOLD, 25));
    	l18.setBounds(400, 80, 260, 30);
    	t3=new JTextField("000");
    	t3.setBounds(720, 80, 90, 30);
    	
    	l19=new JLabel("दैनिक उपस्थिती/ लाभार्थी");
    	l19.setFont(new Font("Serif", Font.BOLD, 25));
    	l19.setBounds(400, 130, 260, 30);
    	t4=new JTextField("000");
    	t4.setBounds(720, 130, 90, 30);
    	
    	l1=new JLabel("प्रत्यक्ष लाभार्थी/ताटाची संख्या");
    	l1.setFont(new Font("Serif", Font.BOLD, 25));
    	l1.setBounds(400, 180, 300, 30);
    	t1=new JTextField("000");
    	t1.setBounds(720, 180, 90, 30);
  
    	
    	
    	t6=new JTextField("000");
    	t6.setBounds(390, 250, 90, 30);
    	t7=new JTextField("000");
    	t7.setBounds(390, 300, 90, 30);
    	t8=new JTextField("000");
    	t8.setBounds(390, 350, 90, 30);
    	t9=new JTextField("000");
    	t9.setBounds(390, 400, 90, 30);
    	t10=new JTextField("000");
    	t10.setBounds(390, 450, 90, 30);
    	t11=new JTextField("000");
    	t11.setBounds(390, 500, 90, 30);
    	t12=new JTextField("000");
    	t12.setBounds(390, 550, 90, 30);
    	t13=new JTextField("000");
    	t13.setBounds(390, 600, 90, 30);
    	t14=new JTextField("000");
    	t14.setBounds(530, 250, 90, 30);
    	t15=new JTextField("000");
    	t15.setBounds(530, 300, 90, 30);
    	t16=new JTextField("000");
    	t16.setBounds(530, 350, 90, 30);
    	t17=new JTextField("000");
    	t17.setBounds(530, 400, 90, 30);
    	t18=new JTextField("000");
    	t18.setBounds(530, 450, 90, 30);
    	t19=new JTextField("000");
    	t19.setBounds(530, 500, 90, 30);
    	t20=new JTextField("000");
    	t20.setBounds(530, 550, 90, 30);
    	t21=new JTextField("000");
    	t21.setBounds(530, 600, 90, 30);
    	t22=new JTextField("000");
    	t22.setBounds(960, 250, 90, 30);
    	t23=new JTextField("000");
    	t23.setBounds(960, 300, 90, 30);
    	t24=new JTextField("000");
    	t24.setBounds(960, 350, 90, 30);
    	t25=new JTextField("000");
    	t25.setBounds(960, 400, 90, 30);
    	t26=new JTextField("000");
    	t26.setBounds(960, 450, 90, 30);
    	t27=new JTextField("000");
    	t27.setBounds(960, 500, 90, 30);
    	t28=new JTextField("000");
    	t28.setBounds(960, 550, 90, 30);
    	t29=new JTextField("000");
    	t29.setBounds(1100, 250, 90, 30);
    	t30=new JTextField("000");
    	t30.setBounds(1100, 300, 90, 30);
    	t31=new JTextField("000");
    	t31.setBounds(1100, 350, 90, 30);
    	t32=new JTextField("000");
    	t32.setBounds(1100, 400, 90, 30);
    	t33=new JTextField("000");
    	t33.setBounds(1100, 450, 90, 30);
    	t34=new JTextField("000");
    	t34.setBounds(1100, 500, 90, 30);
    	t35=new JTextField("000");
    	t35.setBounds(1100, 550, 90, 30);
    	
    	f1.add(c1);
        f1.add(l1);
        f1.add(l2);
        f1.add(l3);
        f1.add(l4);
        f1.add(l5);
        f1.add(l6);
        f1.add(l7);
        f1.add(l8);
        f1.add(l9);
        f1.add(l10);
        f1.add(l11);
        f1.add(l12);
        f1.add(l13);
        f1.add(l14);
        f1.add(l15);
        f1.add(l16);
        f1.add(l17);
        f1.add(l18);
        f1.add(l19);
        f1.add(l21);
        f1.add(l22);
        f1.add(l23);
        f1.add(l24);
        f1.add(l25);
        f1.add(l26);
        f1.add(l27);
        f1.add(l28);
        f1.add(l29);
        f1.add(l30);
        
        f1.add(sp);
        f1.add(sp1);
        f1.add(sp2);
        f1.add(sp3);
        f1.add(sp4);
        f1.add(sp5);
        f1.add(sp6);
        f1.add(sp7);
        f1.add(sp8);
        f1.add(sp9);
        f1.add(sp10);
        f1.add(sp11);
        f1.add(sp12);
        f1.add(sp13);
//        f1.add(sp14);
        
        
        f1.add(b1);
        f1.add(b2);
        
        f1.add(t1);

    	f1.add(t3);
    	f1.add(t4);
    	f1.add(t5);
    	f1.add(t6);
    	f1.add(t7);
    	f1.add(t8);
    	f1.add(t9);
    	f1.add(t10);
    	f1.add(t11);
    	f1.add(t12);
    	f1.add(t13);
    	f1.add(t14);
    	f1.add(t15);
    	f1.add(t16);
    	f1.add(t17);
    	f1.add(t18);
    	f1.add(t19);
    	f1.add(t20);
    	f1.add(t21);
    	f1.add(t22);
    	f1.add(t23);
    	f1.add(t24);
    	f1.add(t25);
    	f1.add(t26);
    	f1.add(t27);
    	f1.add(t28);
    	f1.add(t29);
    	f1.add(t30);
    	f1.add(t31);
    	f1.add(t32);
    	f1.add(t33);
    	f1.add(t34);
    	f1.add(t35);
    	
    	f1.setLayout(null);
		f1.setSize(1390, 730);
		f1.setVisible(true);
    	b2.setEnabled(false);
		b1.addActionListener(this);
    	b2.addActionListener(this);
    }
    
    
   public void actionPerformed(ActionEvent e)
   {
	   
	if(e.getSource()==b1)
	{
		
		JOptionPane.showMessageDialog(f1, "तक्ता तयार झाला, कृपया माहिती भरा.");
		l25.setText("तक्ता तयार झाला आहे! क्रुपया E Drive वर lonkheda. Xls ही फाईल चेक करा.");
		b1.setEnabled(false);
		b1.setBackground(Color.gray);
		b2.setEnabled(true);
		b2.setBackground(Color.GREEN);
			try {
				this.write();
			} catch (WriteException | BiffException | IOException e1) {
		
				e1.printStackTrace();
			}
		
	}
	else if(e.getSource()==b2){
	try {
		this.write1();
	
		l17.setText("दिनांक "+Integer.parseInt((String)c1.getSelectedItem())+" ची माहिती भरली गेली!");
		
	}
	catch (WriteException | BiffException | IOException e1) {
		
		e1.printStackTrace();
		
	}
	    }
	}
	
    
    

public void setOutputFile(String inputFile)
   {
    this.inputFile = inputFile;
    }

public void write1() throws RowsExceededException, WriteException, IOException, BiffException
{ 
	//File file1;  
    file1 = new File(inputFile);  
    // if file doesn't exists, then create it   
    if (!file1.exists()) { 
      file1.createNewFile();  
    }  
	 Workbook workbook1 = Workbook.getWorkbook(file1);  
        WritableWorkbook copy = Workbook.createWorkbook(file1, workbook1);  
        WritableSheet sheet2 = copy.getSheet(0);  
        int j;
    	j=Integer.parseInt((String)c1.getSelectedItem());
    	j=j+5;
		addNumber(sheet2,1,j+1,Integer.parseInt(t3.getText()));
		addNumber(sheet2,2,j+1,Integer.parseInt(t4.getText()));
		addNumber(sheet2,3,j+1,Integer.parseInt(t1.getText()));
		Object ob=sp.getValue();
		double a=Integer.parseInt(ob.toString())*Integer.parseInt(t1.getText());

		addNumber(sheet2,4,j+1,(a/1000));
		System.out.println(a);
		sum=sum+a;
		addNumber(sheet2,4,38,(sum/1000));
		
		Object ob1=sp1.getValue();
		double a1=Integer.parseInt(ob1.toString())*Integer.parseInt(t1.getText());
		addNumber(sheet2,5,j+1,(a1/1000));
		sum1=sum1+a1;
		addNumber(sheet2,5,38,(sum1/1000));
		
		Object ob2=sp2.getValue();
		double a2=Integer.parseInt(ob2.toString())*Integer.parseInt(t1.getText());
		addNumber(sheet2,6,j+1,(a2/1000));
		sum2=sum2+a2;
		addNumber(sheet2,6,38,(sum2/1000));
		
		Object ob3=sp3.getValue();
		double a3=Integer.parseInt(ob3.toString())*Integer.parseInt(t1.getText());
		addNumber(sheet2,7,j+1,(a3/1000));
		sum3=sum3+a3;
		addNumber(sheet2,7,38,(sum3/1000));
		
		Object ob4=sp4.getValue();
		double a4=Integer.parseInt(ob4.toString())*Integer.parseInt(t1.getText());
		addNumber(sheet2,8,j+1,(a4/1000));
		sum4=sum4+a4;
		addNumber(sheet2,8,38,(sum4/1000));
		
		Object ob5=sp5.getValue();
		double a5=Integer.parseInt(ob5.toString())*Integer.parseInt(t1.getText());
		addNumber(sheet2,9,j+1,(a5/1000));
		sum5=sum5+a5;
		addNumber(sheet2,9,38,(sum5/1000));
		
		Object ob6=sp6.getValue();
		double a6=Integer.parseInt(ob6.toString())*Integer.parseInt(t1.getText());
		addNumber(sheet2,10,j+1,(a6/1000));
		sum6=sum6+a6;
		addNumber(sheet2,10,38,(sum6/1000));
		
		
		Object ob7=sp7.getValue();
		double a7=Integer.parseInt(ob7.toString())*Integer.parseInt(t1.getText());
		addNumber(sheet2,11,j+1,(a7/1000));
		sum7=sum7+a7;
		addNumber(sheet2,11,38,(sum7/1000));
		
		Object ob8=sp8.getValue();
		double a8=Integer.parseInt(ob8.toString())*Integer.parseInt(t1.getText());
		addNumber(sheet2,12,j+1,(a8/1000));
		sum8=sum8+a8;
		addNumber(sheet2,12,38,(sum8/1000));
		
		Object ob9=sp9.getValue();
		double a9=Integer.parseInt(ob9.toString())*Integer.parseInt(t1.getText());
		addNumber(sheet2,13,j+1,(a9/1000));
		sum9=sum9+a9;
		addNumber(sheet2,13,38,(sum9/1000));
		
		Object ob10=sp10.getValue();
		double a10=Integer.parseInt(ob10.toString())*Integer.parseInt(t1.getText());
		addNumber(sheet2,14,j+1,(a10/1000));
		sum10=sum10+a10;
		addNumber(sheet2,14,38,(sum10/1000));
		
		Object ob11=sp11.getValue();
		double a11=Integer.parseInt(ob11.toString())*Integer.parseInt(t1.getText());
		addNumber(sheet2,15,j+1,(a11/1000));
		sum11=sum11+a11;
		addNumber(sheet2,15,38,(sum11/1000));
		
		Object ob12=sp12.getValue();
		double a12=Integer.parseInt(ob12.toString())*Integer.parseInt(t1.getText());
		addNumber(sheet2,16,j+1,(a12/1000));
		sum12=sum12+a12;
		addNumber(sheet2,16,38,(sum12/1000));
		
		Object ob13=sp13.getValue();
		double a13=Integer.parseInt(ob13.toString())*Integer.parseInt(t1.getText());
		addNumber(sheet2,17,j+1,(a13/1000));
		sum13=sum13+a13;
		addNumber(sheet2,17,38,(sum13/1000));
		
		double a14=Float.parseFloat(t5.getText())*Integer.parseInt(t1.getText());
		addNumber(sheet2,19,j+1,a14);

		
		

 double b=Float.parseFloat(t6.getText());
		addNumber(sheet2,4,4,b);
		 double b1=Float.parseFloat(t7.getText());
			addNumber(sheet2,5,4,b1);
			 double b2=Float.parseFloat(t8.getText());
				addNumber(sheet2,6,4,b2);
				 double b3=Float.parseFloat(t9.getText());
					addNumber(sheet2,7,4,b3);
					 double b4=Float.parseFloat(t10.getText());
						addNumber(sheet2,8,4,b4);
						 double b5=Float.parseFloat(t11.getText());
							addNumber(sheet2,9,4,b5);
							 double b6=Float.parseFloat(t12.getText());
								addNumber(sheet2,10,4,b6);
								 double b7=Float.parseFloat(t13.getText());
									addNumber(sheet2,11,4,b7);
									 double b8=Float.parseFloat(t22.getText());
										addNumber(sheet2,12,4,b8);
										 double b9=Float.parseFloat(t23.getText());
											addNumber(sheet2,13,4,b9);
											 double b10=Float.parseFloat(t24.getText());
												addNumber(sheet2,14,4,b10);
												 double b11=Float.parseFloat(t25.getText());
													addNumber(sheet2,15,4,b11);
													 double b12=Float.parseFloat(t26.getText());
														addNumber(sheet2,16,4,b12);
														double b13=Float.parseFloat(t27.getText());
														addNumber(sheet2,17,4,b13);
	

	 		double c=Float.parseFloat(t14.getText());
		addNumber(sheet2,4,5,c);
		 double c1=Float.parseFloat(t15.getText());
			addNumber(sheet2,5,5,c1);
			 double c2=Float.parseFloat(t16.getText());
				addNumber(sheet2,6,5,c2);
				 double c3=Float.parseFloat(t17.getText());
					addNumber(sheet2,7,5,c3);
					 double c4=Float.parseFloat(t18.getText());
						addNumber(sheet2,8,5,c4);
						 double c5=Float.parseFloat(t19.getText());
							addNumber(sheet2,9,5,c5);
							 double c6=Float.parseFloat(t20.getText());
								addNumber(sheet2,10,5,c6);
								 double c7=Float.parseFloat(t21.getText());
									addNumber(sheet2,11,5,c7);
									 double c8=Float.parseFloat(t29.getText());
										addNumber(sheet2,12,5,c8);
										 double c9=Float.parseFloat(t30.getText());
											addNumber(sheet2,13,5,c9);
											 double c10=Float.parseFloat(t31.getText());
												addNumber(sheet2,14,5,c10);
												 double c11=Float.parseFloat(t32.getText());
													addNumber(sheet2,15,5,c11);
													 double c12=Float.parseFloat(t33.getText());
														addNumber(sheet2,16,5,c12);
														double c13=Float.parseFloat(t34.getText());
														addNumber(sheet2,17,5,c13);
												

	 double d=(Float.parseFloat(t14.getText())+Float.parseFloat(t6.getText()));
		addNumber(sheet2,4,6,d);
		 double d1=(Float.parseFloat(t15.getText())+Float.parseFloat(t7.getText()));
			addNumber(sheet2,5,6,d1);
			 double d2=(Float.parseFloat(t16.getText())+Float.parseFloat(t8.getText()));
				addNumber(sheet2,6,6,d2);
				 double d3=(Float.parseFloat(t17.getText())+Float.parseFloat(t9.getText()));
					addNumber(sheet2,7,6,d3);
					 double d4=(Float.parseFloat(t18.getText())+Float.parseFloat(t10.getText()));
						addNumber(sheet2,8,6,d4);
						 double d5=(Float.parseFloat(t19.getText())+Float.parseFloat(t11.getText()));
							addNumber(sheet2,9,6,d5);
							 double d6=(Float.parseFloat(t20.getText())+Float.parseFloat(t12.getText()));
								addNumber(sheet2,10,6,d6);
								 double d7=(Float.parseFloat(t21.getText())+Float.parseFloat(t13.getText()));
									addNumber(sheet2,11,6,d7);
									 double d8=(Float.parseFloat(t29.getText())+Float.parseFloat(t22.getText()));
										addNumber(sheet2,12,6,d8);
										 double d9=(Float.parseFloat(t30.getText())+Float.parseFloat(t23.getText()));
											addNumber(sheet2,13,6,d9);
											 double d10=(Float.parseFloat(t31.getText())+Float.parseFloat(t24.getText()));
												addNumber(sheet2,14,6,d10);
												 double d11=(Float.parseFloat(t32.getText())+Float.parseFloat(t25.getText()));
													addNumber(sheet2,15,6,d11);
													 double d12=(Float.parseFloat(t33.getText())+Float.parseFloat(t26.getText()));
														addNumber(sheet2,16,6,d12);
														double d13=(Float.parseFloat(t34.getText())+Float.parseFloat(t27.getText()));
														addNumber(sheet2,17,6,d13);
		if(d==0)
		{
			addNumber(sheet2,4,39,d);
		}
		else 
		{
		addNumber(sheet2,4,39,(d-(sum/1000)));
		}
		 
			if(d1==0)
			{
				addNumber(sheet2,5,39,d1);
			}
			else
			{
				addNumber(sheet2,5,39,(d1-(sum1/1000)));
			}
		
		if(d2==0) 
		{
			addNumber(sheet2,6,39,d2);
		}
		else
		{
			addNumber(sheet2,6,39,(d2-(sum2/1000)));
		}
				
			if(d3==0)
			{
				addNumber(sheet2,7,39,d3);
			}
			else
			{
						addNumber(sheet2,7,39,(d3-(sum3/1000)));
			}
			
		if(d4==0)
		{
			addNumber(sheet2,8,39,d4);
		}
		else
		{
					addNumber(sheet2,8,39,(d4-(sum4/1000)));
		}		
		
			if(d5==0)
			{
				addNumber(sheet2,9,39,d5);
			}
			else
			{
				addNumber(sheet2,9,39,(d5-(sum5/1000)));

			}
			
		if(d6==0)
		{
			addNumber(sheet2,10,39,d6);
		}
		else
		{
			addNumber(sheet2,10,39,(d6-(sum6/1000)));
		}
		
			if(d7==0)
			{
				addNumber(sheet2,11,39,d7);
			}
			else
			{
				addNumber(sheet2,11,39,(d7-(sum7/1000)));
			}
			
		if(d8==0)
		{
			addNumber(sheet2,12,39,d8);
		}
		else
		{
			addNumber(sheet2,12,39,(d8-(sum8/1000)));
		}
		
			if(d9==0)
			{
				addNumber(sheet2,13,39,d9);
			}
			else
			{
				addNumber(sheet2,13,39,(d9-(sum9/1000)));
			}
			
		if(d10==0)
		{
			addNumber(sheet2,14,39,d10);
		}
		else
		{
			addNumber(sheet2,14,39,(d10-(sum10/1000)));	
		}
		
			if(d11==0)
			{
				addNumber(sheet2,15,39,d11);
			}
			else
			{
				addNumber(sheet2,15,39,(d11-(sum11/1000)));
			}
			
		if(d12==0)
		{
			addNumber(sheet2,16,39,d12);
		}
		else
		{
			addNumber(sheet2,16,39,(d12-(sum12/1000)));
		}
		
			if(d13==0)
			{
				addNumber(sheet2,17,39,d13);
			}
			else
			{
				addNumber(sheet2,17,39,(d13-(sum13/1000)));
			}
														
														
												
	
	 
    	 
        copy.write();  
        copy.close();
	}

    public void write() throws IOException, WriteException, BiffException 
    {
        //File file = new File(inputFile);
        //File file;  
        file = new File(inputFile);  
        // if file doesnt exists, then create it   
        if (!file.exists()) {  
          file.createNewFile(); 
        }  
        WorkbookSettings wbSettings = new WorkbookSettings();

        //wbSettings.setLocale(new Locale("en", "EN"));

        WritableWorkbook workbook = Workbook.createWorkbook(file, wbSettings);
        
        workbook.createSheet("Report", 0);
        WritableSheet excelSheet = workbook.getSheet(0);
        createLabel(excelSheet);
        createContent(excelSheet);
       
        workbook.write();
       
        workbook.close();
       
        
    }
    private void createLabel(WritableSheet sheet) throws WriteException 
    {
        WritableFont times10pt = new WritableFont(WritableFont.TIMES, 10);
        times = new WritableCellFormat(times10pt);
        times.setWrap(true);
        WritableFont times10ptBoldUnderline = new WritableFont(
        WritableFont.TIMES, 15, WritableFont.BOLD, false,
        UnderlineStyle.SINGLE);
        timesBoldUnderline = new WritableCellFormat(times10ptBoldUnderline);
        CellView cv = new CellView();
        cv.setFormat(times);
        cv.setFormat(timesBoldUnderline);
        addCaption(sheet, 5, 0, "राष्ट्रीय माध्यान्ह भोजन योजना (इयत्ता 6 वी ते 8 वी)");
        addCaption(sheet, 5, 1, "शाळा स्तरावर ठेवण्याची दैनंदिन खर्च नोंदवही भाग - 2");
        addCaption(sheet, 3, 2, "*शाळेचे नाव :सरस्वती विद्यालय *  *केंद्र :लोणखेडा*   *महिना:     *  वजन (किलोग्रॅम मध्ये )");
       

    }

    private void createContent(WritableSheet sheet) throws WriteException,
            RowsExceededException {
               // now a bit of text
            // First column
    	for(int j=7;j<38;j++)
    	{
    		addNumber(sheet,0,j,j-6);
    	}
    	
    	int j;
    	j=Integer.parseInt((String)c1.getSelectedItem());
    	j=j+5;
		addNumber(sheet,1,j+1,Integer.parseInt(t3.getText()));
		addNumber(sheet,2,j+1,Integer.parseInt(t4.getText()));
		addNumber(sheet,3,j+1,Integer.parseInt(t1.getText()));
		Object ob=sp.getValue();
		double a=Integer.parseInt(ob.toString())*Integer.parseInt(t1.getText());
		addNumber(sheet,4,j+1,(a/1000));
		
		Object ob1=sp1.getValue();
		double a1=Integer.parseInt(ob1.toString())*Integer.parseInt(t1.getText());
		addNumber(sheet,5,j+1,(a1/1000));
		
		Object ob2=sp2.getValue();
		double a2=Integer.parseInt(ob2.toString())*Integer.parseInt(t1.getText());
		addNumber(sheet,6,j+1,(a2/1000));
		
		Object ob3=sp3.getValue();
		double a3=Integer.parseInt(ob3.toString())*Integer.parseInt(t1.getText());
		addNumber(sheet,7,j+1,(a3/1000));
		
		Object ob4=sp4.getValue();
		double a4=Integer.parseInt(ob4.toString())*Integer.parseInt(t1.getText());
		addNumber(sheet,8,j+1,(a4/1000));
		
		Object ob5=sp5.getValue();
		double a5=Integer.parseInt(ob5.toString())*Integer.parseInt(t1.getText());
		addNumber(sheet,9,j+1,(a5/1000));
		
		Object ob6=sp6.getValue();
		double a6=Integer.parseInt(ob6.toString())*Integer.parseInt(t1.getText());
		addNumber(sheet,10,j+1,(a6/1000));
		
		Object ob7=sp7.getValue();
		double a7=Integer.parseInt(ob7.toString())*Integer.parseInt(t1.getText());
		addNumber(sheet,11,j+1,(a7/1000));
		
		Object ob8=sp8.getValue();
		double a8=Integer.parseInt(ob8.toString())*Integer.parseInt(t1.getText());
		addNumber(sheet,12,j+1,(a8/1000));
		
		Object ob9=sp9.getValue();
		double a9=Integer.parseInt(ob9.toString())*Integer.parseInt(t1.getText());
		addNumber(sheet,13,j+1,(a9/1000));
		
		Object ob10=sp10.getValue();
		double a10=Integer.parseInt(ob10.toString())*Integer.parseInt(t1.getText());
		addNumber(sheet,14,j+1,(a10/1000));
		
		Object ob11=sp11.getValue();
		double a11=Integer.parseInt(ob11.toString())*Integer.parseInt(t1.getText());
		addNumber(sheet,15,j+1,(a11/1000));
		
		Object ob12=sp12.getValue();
		double a12=Integer.parseInt(ob12.toString())*Integer.parseInt(t1.getText());
		addNumber(sheet,16,j+1,(a12/1000));
		
		Object ob13=sp13.getValue();
		double a13=Integer.parseInt(ob13.toString())*Integer.parseInt(t1.getText());
		addNumber(sheet,17,j+1,(a13/1000));
    	
		Object ob14=t5.getText();
		double a14=Integer.parseInt(ob14.toString())*Integer.parseInt(t1.getText());
		addNumber(sheet,19,j+1,a14);
		
    	
    	int i=3;
    	    addLabel(sheet, 0, i , "दिनांक");
    	    addLabel(sheet, 0, 4 , "मासिक शिल्लक साहित्य");
    	    addLabel(sheet, 0, 5 , "प्राप्त साहित्य");
    	    addLabel(sheet, 0, 6 , "एकूण साहित्य");
            addLabel(sheet, 1, i, "चालू महिन्याची पटसंख्या");
            addLabel(sheet, 2, i, "दैनिक उपस्थिती/ लाभार्थी");
            addLabel(sheet, 3, i, "प्रत्यक्ष लाभार्थी/ताटाची संख्या");
            addLabel(sheet, 4, i, "मुगडाळ (kg)");
            addLabel(sheet, 5, i, "तूरडाळ (kg)");
            addLabel(sheet, 6, i, "मसुरडाळ   (kg)");
            addLabel(sheet, 7, i, "मटकी (kg)");
            addLabel(sheet, 8, i, "मुग      (kg)");
            addLabel(sheet, 9, i, "चवळी (kg)");
            addLabel(sheet, 10, i, "हरभरा (kg)");
            addLabel(sheet, 11, i, "वाटाणा (kg)");
            addLabel(sheet, 12, i, "जिरे      (kg)");
            addLabel(sheet, 13, i, "मोहरी        (kg)");
            addLabel(sheet, 14, i, "हळद        (kg)");
            addLabel(sheet, 15, i, "मिरची पावडर         (kg)");
            addLabel(sheet, 16, i, "सोयाबीन तेल         (kg)");
            addLabel(sheet, 17, i, "भाजीपाला    (kg)");
            addLabel(sheet, 18, i, "पूरक आहार");
            addLabel(sheet, 19, i, "लाभार्थी नुसार होणार खर्च");
            addLabel(sheet, 20, i, "शेरा");
            addLabel(sheet, 0, 38 , "एकूण");
            addLabel(sheet, 0, 39 , "शिल्लक साहित्य");
       
    }

    private void addCaption(WritableSheet sheet, int column, int row, String s) throws RowsExceededException, WriteException 
    {
        Label label;
        label = new Label(column, row, s, timesBoldUnderline);
        sheet.addCell(label);
    }

    private void addNumber(WritableSheet sheet, int column, int row,double d) throws WriteException, RowsExceededException
    {
        Number number;
        number = new Number(column, row, d, times);
        sheet.addCell(number);
    }

    private void addLabel(WritableSheet sheet, int column, int row, String s) throws WriteException, RowsExceededException
    {
        Label label;
        label = new Label(column, row, s, times);
        sheet.addCell(label);
    }

    public static void main(String[] args) throws WriteException, IOException 
    {
        re test = new re();
        test.setOutputFile("D:\\Lonkheda.xls");
    }
}

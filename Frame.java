package com.liuwei.IO;

import java.io.*;
import java.awt.*;
import javax.swing.*;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.awt.event.*;

public class Frame implements ActionListener{
	
	private HSSFWorkbook workbook = new HSSFWorkbook();
	
	HSSFSheet classNumSheet = workbook.createSheet("ǰN������ͳ��");
	HSSFSheet aveRankSheet = workbook.createSheet("ƽ����ͳ��");
	

	
	
	
    File fileQuankeShangci;  
    File fileQuankeBenci;
    File fileWenkeShangci;
    File fileWenkeBenci;
    boolean teacher = true;
    String teacherString = "�Ƿ����ɽ�ʦ��";
    int likeClassSize = 0;
    int wenkeClassSize = 0;
    
    JFrame frame=new JFrame("�ɼ�����V1.1");
    JTabbedPane tabPane=new JTabbedPane();//ѡ�����
    Container con=new Container();//����1
    JLabel label1=new JLabel("ѡ�񱾴γɼ���ȫ��/��ƣ�");
    JLabel label2=new JLabel("ѡ���ϴγɼ���ȫ��/��ƣ�");
    JLabel label3=new JLabel("ѡ�񱾴γɼ����Ŀƣ�");
    JLabel label4=new JLabel("ѡ���ϴγɼ����Ŀƣ�");
    JLabel label5=new JLabel(teacherString);
    
    JTextField text1=new JTextField();
    JTextField text2=new JTextField();
    JTextField text3=new JTextField();
    JTextField text4=new JTextField();

    JButton button1=new JButton("���");
    JButton button2=new JButton("���");
    JButton button3=new JButton("���");
    JButton button4=new JButton("���");
    JButton button5=new JButton("��");
    JButton button6=new JButton("��");
    JButton button7=new JButton("��ʼ����");
     
    JFileChooser jfc=new JFileChooser();//�ļ�ѡ���� 
    
    public void functionDouble(File fileShangci,File fileBenci,boolean science) {
    	WriteExcel writeExcel = new WriteExcel(fileShangci,fileBenci,classNumSheet,aveRankSheet);
    	writeExcel.setXiangtong();
		boolean xiangtong = writeExcel.getXiangtong();
		if(science) {
			writeExcel.setClassSize();
			likeClassSize = writeExcel.getClassSize();
		}else {
			writeExcel.setClassSize();
			wenkeClassSize = writeExcel.getClassSize();
		}		
		if(xiangtong) {
    		writeExcel.writeWorkbook(science,likeClassSize,wenkeClassSize,false,teacher);
    		WriteFile writeFile = new WriteFile(workbook);
			windows("�Ѵ������Դ�ļ�");
		}else {
			windows("Դ�ļ�ѧ����Ϣ��ͳһ");
		}
    }
    public void windows(String string) {
		JFrame result=new JFrame("�ɼ����");
        JTabbedPane tabPane1=new JTabbedPane();//ѡ�����
        Container con1=new Container();//����1
		double lx1=Toolkit.getDefaultToolkit().getScreenSize().getWidth();
		double ly1=Toolkit.getDefaultToolkit().getScreenSize().getHeight();
		result.setLocation(new Point((int)(lx1/2)-200,(int)(ly1/2)-75));//�趨���ڳ���λ��
		result.setSize(400,150);//�趨���ڴ�С
		result.setContentPane(tabPane1);//���ò���
	    tabPane1.add("�������",con1);//��Ӳ���1
	    
	    JLabel label5=new JLabel(string);
	    label5.setBounds(125,20,200,40);
	    con1.add(label5);
	    result.setVisible(true);//���ڿɼ�
	    result.setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);//ʹ�ܹرմ��ڣ���������
    }
    
	Frame(){

		jfc.setCurrentDirectory(new File("d:\\"));//�ļ�ѡ�����ĳ�ʼĿ¼��Ϊd��
		//����������ȡ����Ļ�ĸ߶ȺͿ��
		double lx=Toolkit.getDefaultToolkit().getScreenSize().getWidth();
		double ly=Toolkit.getDefaultToolkit().getScreenSize().getHeight();
		frame.setLocation(new Point((int)(lx/2)-250,(int)(ly/2)-150));//�趨���ڳ���λ��
	    frame.setSize(550,300);//�趨���ڴ�С
	    frame.setContentPane(tabPane);//���ò���
	    int X = 10;
	    int Y = 10;
	    int lableWidth = 160;
	    int height = 20;
	       //�����趨��ǩ�ȵĳ���λ�ú͸߿�
	        label1.setBounds(X,Y,lableWidth,height);
	        label2.setBounds(X,Y+20,lableWidth,height);
	        label3.setBounds(X,Y+60,lableWidth,height);
	        label4.setBounds(X,Y+80,lableWidth,height);
	        label5.setBounds(X,Y+120,lableWidth,height);
	        
	        int textWidth =250;
	        text1.setBounds(X+lableWidth,Y,textWidth,height);
	        text2.setBounds(X+lableWidth,Y+20,textWidth,height);
	        text3.setBounds(X+lableWidth,Y+60,textWidth,height);
	        text4.setBounds(X+lableWidth,Y+80,textWidth,height);
	        
	        int buttonWidth =80;
	        button1.setBounds(X+lableWidth+textWidth,Y,buttonWidth,20);
	        button2.setBounds(X+lableWidth+textWidth,Y+20,buttonWidth,20);
	        button3.setBounds(X+lableWidth+textWidth,Y+60,buttonWidth,20);
	        button4.setBounds(X+lableWidth+textWidth,Y+80,buttonWidth,20);
	        
	        button5.setBounds((X+lableWidth+textWidth+buttonWidth)/2-60,Y+120,buttonWidth,20);
	        button6.setBounds((X+lableWidth+textWidth+buttonWidth)/2-60+buttonWidth,Y+120,buttonWidth,20);
	        button7.setBounds((X+lableWidth+textWidth+buttonWidth)/2-60,Y+160,buttonWidth*2,40);
	        
	        
	        
	        button1.addActionListener(this);//����¼�����
	        button2.addActionListener(this);//����¼�����
	        button3.addActionListener(this);//����¼�����
	        button4.addActionListener(this);//����¼�����
	        button5.addActionListener(this);//����¼�����
	        button6.addActionListener(this);//����¼�����
	        
	        button7.addActionListener(this);//����¼�����
	        
	        
	        con.add(label1);
	        con.add(label2);
	        con.add(label3);
	        con.add(label4);
	        con.add(label5);

	        con.add(text1);
	        con.add(text2);
	        con.add(text3);
	        con.add(text4);

	        
	        
	        con.add(button1);
	        con.add(button2);
	        con.add(button3);
	        con.add(button4);
	        con.add(button5);
	        con.add(button6);
	        con.add(button7);
	        con.add(jfc);
	        tabPane.add("�ļ�ѡ��",con);//��Ӳ���1
	        frame.setVisible(true);//���ڿɼ�
	        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);//ʹ�ܹرմ��ڣ���������
	    }
	    
	public void actionPerformed(ActionEvent e){//�¼�����
        if(e.getSource().equals(button1)){//�жϴ��������İ�ť���ĸ�
            jfc.setFileSelectionMode(0);//�趨ֻ��ѡ���ļ�
            int state=jfc.showOpenDialog(null);//�˾��Ǵ��ļ�ѡ��������Ĵ������
            if(state==1){
                return;//�����򷵻�
            }
            else{
                this.fileQuankeBenci=jfc.getSelectedFile();
                text1.setText(fileQuankeBenci.getAbsolutePath());
            }
        }
        if(e.getSource().equals(button2)){
            jfc.setFileSelectionMode(0);//�趨ֻ��ѡ���ļ�
            int state=jfc.showOpenDialog(null);//�˾��Ǵ��ļ�ѡ��������Ĵ������
            if(state==1){
                return;//�����򷵻�
            }
            else{
                this.fileQuankeShangci=jfc.getSelectedFile();//fΪѡ�񵽵��ļ�
                text2.setText(fileQuankeShangci.getAbsolutePath());
            }
        }
        
        if(e.getSource().equals(button3)){//�жϴ��������İ�ť���ĸ�
            jfc.setFileSelectionMode(0);//�趨ֻ��ѡ���ļ���
            int state=jfc.showOpenDialog(null);//�˾��Ǵ��ļ�ѡ��������Ĵ������
            if(state==1){
                return;//�����򷵻�
            }
            else{
                this.fileWenkeBenci=jfc.getSelectedFile();//fΪѡ�񵽵�Ŀ¼
                text3.setText(this.fileWenkeBenci.getAbsolutePath());
            }
        }
        if(e.getSource().equals(button4)){
            jfc.setFileSelectionMode(0);//�趨ֻ��ѡ���ļ�
            int state=jfc.showOpenDialog(null);//�˾��Ǵ��ļ�ѡ��������Ĵ������
            if(state==1){
                return;//�����򷵻�
            }
            else{
                this.fileWenkeShangci=jfc.getSelectedFile();//fΪѡ�񵽵��ļ�
                text4.setText(this.fileWenkeShangci.getAbsolutePath());
            }          
        }
        if(e.getSource().equals(button5)){
        	windows("���ɽ�ʦ��");
        	teacher = true;
        }
        if(e.getSource().equals(button6)){
        	windows("�����ɽ�ʦ��");
        	teacher = false;
        }
        
        if(e.getSource().equals(button7)){
        	if(fileQuankeShangci==null&&fileQuankeBenci==null&&fileWenkeShangci==null&&fileWenkeBenci==null) {
        		windows("����ȫ���ļ�·��");
    		}else if(fileQuankeShangci!=null&&fileQuankeBenci!=null&&fileWenkeShangci!=null&&fileWenkeBenci!=null){
    			WriteExcel writeExcelQuanke = new WriteExcel(fileQuankeShangci, fileQuankeBenci,classNumSheet,aveRankSheet);
    			writeExcelQuanke.setXiangtong();
    			boolean xiangtongQuanke = writeExcelQuanke.getXiangtong();
    			writeExcelQuanke.setClassSize();
    			likeClassSize = writeExcelQuanke.getClassSize();
    			
    			WriteExcel writeExcelWenke = new WriteExcel(fileWenkeShangci, fileWenkeBenci,classNumSheet,aveRankSheet);
    			writeExcelWenke.setXiangtong();
    			boolean xiangtongWenke = writeExcelWenke.getXiangtong();
    			writeExcelWenke.setClassSize();
    			wenkeClassSize = writeExcelWenke.getClassSize();
    			
    			if(xiangtongQuanke&&xiangtongWenke) {
    				writeExcelQuanke.writeWorkbook(true, likeClassSize, wenkeClassSize,false, teacher);
    				writeExcelWenke.writeWorkbook(false, likeClassSize, wenkeClassSize, false,teacher);
    				WriteFile writeFile = new WriteFile(workbook);
    				windows("�Ѵ������Դ�ļ�");
    			}else if(xiangtongWenke&&(xiangtongQuanke==false)) {
    				windows("ȫ��/���Դ�ļ�ѧ����Ϣ��ͳһ");
    			}else if(xiangtongQuanke&&(xiangtongWenke==false)) {
    				windows("�Ŀ�Դ�ļ�ѧ����Ϣ��ͳһ");
    			}else {
    				windows("ȫ��/��ơ��Ŀ�Դ�ļ�ѧ����Ϣ��ͳһ");
    			}
    		}else if(fileQuankeShangci!=null&&fileQuankeBenci!=null&&fileWenkeShangci==null&&fileWenkeBenci==null) {
    			functionDouble(fileQuankeShangci,fileQuankeBenci,true);
    		}else if(fileQuankeShangci==null&&fileQuankeBenci==null&&fileWenkeShangci!=null&&fileWenkeBenci!=null) {
    			functionDouble(fileWenkeShangci,fileWenkeBenci,false);
    		}
    		else if(fileQuankeShangci==null&&fileQuankeBenci!=null&&fileWenkeShangci==null&&fileWenkeBenci!=null) {
    			
    			WriteExcel writeExcelQuanke = new WriteExcel(fileQuankeBenci,classNumSheet,aveRankSheet);
    			writeExcelQuanke.setClassSize();
    			likeClassSize = writeExcelQuanke.getClassSize();
    			
    			WriteExcel writeExcelWenke = new WriteExcel(fileWenkeBenci,classNumSheet,aveRankSheet);
    			writeExcelWenke.setClassSize();
    			wenkeClassSize = writeExcelWenke.getClassSize();
    			
				writeExcelQuanke.writeWorkbook(true, likeClassSize, wenkeClassSize,true,teacher);				
				writeExcelWenke.writeWorkbook(false, likeClassSize, wenkeClassSize,true,teacher);
				
				WriteFile writeFile = new WriteFile(workbook);
    			windows("�Ѵ������Դ�ļ�");
    		}else if(fileQuankeShangci==null&&fileQuankeBenci!=null&&fileWenkeShangci==null&&fileWenkeBenci==null) {
    			WriteExcel writeExcel = new WriteExcel(fileQuankeBenci,classNumSheet,aveRankSheet);
    			writeExcel.setClassSize();
    			likeClassSize = writeExcel.getClassSize();
    			wenkeClassSize = 0;
    			writeExcel.writeWorkbook(true,likeClassSize, wenkeClassSize, true,teacher);
    			WriteFile writeFile = new WriteFile(workbook);
    			windows("�Ѵ������Դ�ļ�");
    		}else if(fileQuankeShangci==null&&fileQuankeBenci==null&&fileWenkeShangci==null&&fileWenkeBenci!=null) {
    			
    			WriteExcel writeExcel = new WriteExcel(fileWenkeBenci,classNumSheet,aveRankSheet);
    			writeExcel.setClassSize();
    			wenkeClassSize = writeExcel.getClassSize();
    			likeClassSize = 0;
    			writeExcel.writeWorkbook(false,likeClassSize, wenkeClassSize, true,teacher);
    			WriteFile writeFile = new WriteFile(workbook);
    			windows("�Ѵ������Դ�ļ�");
    		}else {
    			windows("���鱾�γɼ�·��");
    		}
        
        }
    }
}


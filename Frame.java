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
	
	HSSFSheet classNumSheet = workbook.createSheet("前N名人数统计");
	HSSFSheet aveRankSheet = workbook.createSheet("平均分统计");
	

	
	
	
    File fileQuankeShangci;  
    File fileQuankeBenci;
    File fileWenkeShangci;
    File fileWenkeBenci;
    boolean teacher = true;
    String teacherString = "是否生成教师列";
    int likeClassSize = 0;
    int wenkeClassSize = 0;
    
    JFrame frame=new JFrame("成绩分析V1.1");
    JTabbedPane tabPane=new JTabbedPane();//选项卡布局
    Container con=new Container();//布局1
    JLabel label1=new JLabel("选择本次成绩（全科/理科）");
    JLabel label2=new JLabel("选择上次成绩（全科/理科）");
    JLabel label3=new JLabel("选择本次成绩（文科）");
    JLabel label4=new JLabel("选择上次成绩（文科）");
    JLabel label5=new JLabel(teacherString);
    
    JTextField text1=new JTextField();
    JTextField text2=new JTextField();
    JTextField text3=new JTextField();
    JTextField text4=new JTextField();

    JButton button1=new JButton("浏览");
    JButton button2=new JButton("浏览");
    JButton button3=new JButton("浏览");
    JButton button4=new JButton("浏览");
    JButton button5=new JButton("是");
    JButton button6=new JButton("否");
    JButton button7=new JButton("开始分析");
     
    JFileChooser jfc=new JFileChooser();//文件选择器 
    
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
			windows("已处理完成源文件");
		}else {
			windows("源文件学号信息不统一");
		}
    }
    public void windows(String string) {
		JFrame result=new JFrame("成绩结果");
        JTabbedPane tabPane1=new JTabbedPane();//选项卡布局
        Container con1=new Container();//布局1
		double lx1=Toolkit.getDefaultToolkit().getScreenSize().getWidth();
		double ly1=Toolkit.getDefaultToolkit().getScreenSize().getHeight();
		result.setLocation(new Point((int)(lx1/2)-200,(int)(ly1/2)-75));//设定窗口出现位置
		result.setSize(400,150);//设定窗口大小
		result.setContentPane(tabPane1);//设置布局
	    tabPane1.add("分析结果",con1);//添加布局1
	    
	    JLabel label5=new JLabel(string);
	    label5.setBounds(125,20,200,40);
	    con1.add(label5);
	    result.setVisible(true);//窗口可见
	    result.setDefaultCloseOperation(JFrame.DISPOSE_ON_CLOSE);//使能关闭窗口，结束程序
    }
    
	Frame(){

		jfc.setCurrentDirectory(new File("d:\\"));//文件选择器的初始目录定为d盘
		//下面两行是取得屏幕的高度和宽度
		double lx=Toolkit.getDefaultToolkit().getScreenSize().getWidth();
		double ly=Toolkit.getDefaultToolkit().getScreenSize().getHeight();
		frame.setLocation(new Point((int)(lx/2)-250,(int)(ly/2)-150));//设定窗口出现位置
	    frame.setSize(550,300);//设定窗口大小
	    frame.setContentPane(tabPane);//设置布局
	    int X = 10;
	    int Y = 10;
	    int lableWidth = 160;
	    int height = 20;
	       //下面设定标签等的出现位置和高宽
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
	        
	        
	        
	        button1.addActionListener(this);//添加事件处理
	        button2.addActionListener(this);//添加事件处理
	        button3.addActionListener(this);//添加事件处理
	        button4.addActionListener(this);//添加事件处理
	        button5.addActionListener(this);//添加事件处理
	        button6.addActionListener(this);//添加事件处理
	        
	        button7.addActionListener(this);//添加事件处理
	        
	        
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
	        tabPane.add("文件选择",con);//添加布局1
	        frame.setVisible(true);//窗口可见
	        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);//使能关闭窗口，结束程序
	    }
	    
	public void actionPerformed(ActionEvent e){//事件处理
        if(e.getSource().equals(button1)){//判断触发方法的按钮是哪个
            jfc.setFileSelectionMode(0);//设定只能选择到文件
            int state=jfc.showOpenDialog(null);//此句是打开文件选择器界面的触发语句
            if(state==1){
                return;//撤销则返回
            }
            else{
                this.fileQuankeBenci=jfc.getSelectedFile();
                text1.setText(fileQuankeBenci.getAbsolutePath());
            }
        }
        if(e.getSource().equals(button2)){
            jfc.setFileSelectionMode(0);//设定只能选择到文件
            int state=jfc.showOpenDialog(null);//此句是打开文件选择器界面的触发语句
            if(state==1){
                return;//撤销则返回
            }
            else{
                this.fileQuankeShangci=jfc.getSelectedFile();//f为选择到的文件
                text2.setText(fileQuankeShangci.getAbsolutePath());
            }
        }
        
        if(e.getSource().equals(button3)){//判断触发方法的按钮是哪个
            jfc.setFileSelectionMode(0);//设定只能选择到文件夹
            int state=jfc.showOpenDialog(null);//此句是打开文件选择器界面的触发语句
            if(state==1){
                return;//撤销则返回
            }
            else{
                this.fileWenkeBenci=jfc.getSelectedFile();//f为选择到的目录
                text3.setText(this.fileWenkeBenci.getAbsolutePath());
            }
        }
        if(e.getSource().equals(button4)){
            jfc.setFileSelectionMode(0);//设定只能选择到文件
            int state=jfc.showOpenDialog(null);//此句是打开文件选择器界面的触发语句
            if(state==1){
                return;//撤销则返回
            }
            else{
                this.fileWenkeShangci=jfc.getSelectedFile();//f为选择到的文件
                text4.setText(this.fileWenkeShangci.getAbsolutePath());
            }          
        }
        if(e.getSource().equals(button5)){
        	windows("生成教师列");
        	teacher = true;
        }
        if(e.getSource().equals(button6)){
        	windows("不生成教师列");
        	teacher = false;
        }
        
        if(e.getSource().equals(button7)){
        	if(fileQuankeShangci==null&&fileQuankeBenci==null&&fileWenkeShangci==null&&fileWenkeBenci==null) {
        		windows("请检查全部文件路径");
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
    				windows("已处理完成源文件");
    			}else if(xiangtongWenke&&(xiangtongQuanke==false)) {
    				windows("全科/理科源文件学号信息不统一");
    			}else if(xiangtongQuanke&&(xiangtongWenke==false)) {
    				windows("文科源文件学号信息不统一");
    			}else {
    				windows("全科/理科、文科源文件学号信息不统一");
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
    			windows("已处理完成源文件");
    		}else if(fileQuankeShangci==null&&fileQuankeBenci!=null&&fileWenkeShangci==null&&fileWenkeBenci==null) {
    			WriteExcel writeExcel = new WriteExcel(fileQuankeBenci,classNumSheet,aveRankSheet);
    			writeExcel.setClassSize();
    			likeClassSize = writeExcel.getClassSize();
    			wenkeClassSize = 0;
    			writeExcel.writeWorkbook(true,likeClassSize, wenkeClassSize, true,teacher);
    			WriteFile writeFile = new WriteFile(workbook);
    			windows("已处理完成源文件");
    		}else if(fileQuankeShangci==null&&fileQuankeBenci==null&&fileWenkeShangci==null&&fileWenkeBenci!=null) {
    			
    			WriteExcel writeExcel = new WriteExcel(fileWenkeBenci,classNumSheet,aveRankSheet);
    			writeExcel.setClassSize();
    			wenkeClassSize = writeExcel.getClassSize();
    			likeClassSize = 0;
    			writeExcel.writeWorkbook(false,likeClassSize, wenkeClassSize, true,teacher);
    			WriteFile writeFile = new WriteFile(workbook);
    			windows("已处理完成源文件");
    		}else {
    			windows("请检查本次成绩路径");
    		}
        
        }
    }
}


package com.liuwei.IO;

import java.io.File;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.Workbook;

public class WriteExcel {

	HSSFSheet classNumSheet;
	HSSFSheet aveRankSheet;
	
	
	ReadData readBenci;
	ReadData readShangci;
	
	AnalysisData dataBenci;
	AnalysisData dataShangci;
	
	ArrayList<String> singleClassListShangci;
	ArrayList<String> singleClassListBenci;

	
	
	
	
	WriteExcel(File shangci,File benci,HSSFSheet classNumSheet,HSSFSheet aveRankSheet){
		
		this.classNumSheet = classNumSheet;
		this.aveRankSheet = aveRankSheet;
		//读取两次数据
		this.readBenci = new ReadData();
		try {
			readBenci.setTableData(benci);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
						
		this.readShangci = new ReadData();
		try {
			readShangci.setTableData(shangci);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

		//初始化本次班级名信息
		this.dataBenci = new AnalysisData();
		dataBenci.setSingleClassList(readBenci.getTableData());
		this.singleClassListBenci = dataBenci.getSingleClassList();

				
		//初始化上次班级名信息
		this.dataShangci = new AnalysisData();
		dataShangci.setSingleClassList(readShangci.getTableData());
		this.singleClassListShangci = dataShangci.getSingleClassList();
	}
	
	WriteExcel(File benci,HSSFSheet classNumSheet,HSSFSheet aveRankSheet){
		
		this.classNumSheet = classNumSheet;
		this.aveRankSheet = aveRankSheet;
		//读取两次数据
		this.readBenci = new ReadData();
		try {
			readBenci.setTableData(benci);
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		//初始化本次班级名信息
		this.dataBenci = new AnalysisData();
		dataBenci.setSingleClassList(readBenci.getTableData());
		this.singleClassListBenci = dataBenci.getSingleClassList();
	}
	
	public void setClassNumSheet(Map<String,Object> map,List<String> list,int row,int col,HSSFSheet classNumSheet) {	
			for(int i = 1;i< list.size()+1;i++) {
				if(map.get(list.get(i-1)).getClass().getSimpleName().equals("String")) {
					classNumSheet.getRow(row+i-1).createCell(col).setCellValue((String)map.get(list.get(i-1)));
				}else if(map.get(list.get(i-1)).getClass().getSimpleName().equals("Integer")){
					classNumSheet.getRow(row+i-1).createCell(col).setCellValue((Integer)map.get(list.get(i-1)));
				}else if(map.get(list.get(i-1)).getClass().getSimpleName().equals("Double")) {
					classNumSheet.getRow(row+i-1).createCell(col).setCellValue((Double)map.get(list.get(i-1)));
				}
			}
			if(col!=0) {
				classNumSheet.setColumnWidth(col, 2000);
			}
	}
	public void setAveRankSheet(Map<String,Object> map,List<String> list,int row,int col,HSSFSheet aveRankSheet) {
		for(int i = 1;i< list.size()+1;i++) {
			if(map.get(list.get(i-1)).getClass().getSimpleName().equals("String")) {
				aveRankSheet.getRow(row+i-1).createCell(col).setCellValue((String)map.get(list.get(i-1)));
			}else if(map.get(list.get(i-1)).getClass().getSimpleName().equals("Integer")){
				aveRankSheet.getRow(row+i-1).createCell(col).setCellValue((Integer)map.get(list.get(i-1)));
			}else if(map.get(list.get(i-1)).getClass().getSimpleName().equals("Double")) {
				aveRankSheet.getRow(row+i-1).createCell(col).setCellValue((Double)map.get(list.get(i-1)));
			}
		}
	}
	
	private Boolean xiangtong;//判断两个文件是否可以分析
	public void setXiangtong() {
		//比较第一个那个长
		if(singleClassListShangci.size() == singleClassListBenci.size()) {
			this.xiangtong = true;
		}
		else {
			this.xiangtong = false;
		}
	}
	
	public Boolean getXiangtong() {
		return xiangtong;
	}
	
	
	private int classSize ;
	public void setClassSize() {
		this.classSize = singleClassListBenci.size()+2;
	}
	
	public int getClassSize() {
		return classSize;
	}
	
	
	public void setTotalClassNumSheet(boolean science,int likeClassSize,int wenkeClassSize,boolean single){
		//针对前N名创建含有标题栏的list
		List<String> classNumlist = new ArrayList<>();
		classNumlist.add("班级");
		classNumlist.addAll(singleClassListBenci);	

		int first;
		int second;
		List<Integer> rank;
		if(science) {
			rank = new ArrayList<Integer>(Arrays.asList(10,30,50,100,150,200,250,300,350,400));
			first = 1;
		}else {
			rank = new ArrayList<Integer>(Arrays.asList(10,20,30,50,80,100,130,150));
			first = likeClassSize+1;
		}
		second = first + wenkeClassSize+likeClassSize;
		
		//在classNumSheet创建对象
		for(int i = first;i< classNumlist.size()+first;i++) {
			classNumSheet.createRow(i);
		}
		for(int i= second;i< classNumlist.size()+second;i++) {
			classNumSheet.createRow(i);
		}
		
		if(single) {
			//第一次写入第一列数据
			for(int i=first;i<classNumlist.size()+first;i++) {
				classNumSheet.getRow(i).createCell(0).setCellValue(classNumlist.get(i-first));
			}
		}else {
			//第一次写入第一列数据
			for(int i=first;i<classNumlist.size()+first;i++) {
				classNumSheet.getRow(i).createCell(0).setCellValue(classNumlist.get(i-first));
			}
			//第二次写入第一列数据
			for(int i=second;i<classNumlist.size()+second;i++) {
				classNumSheet.getRow(i).createCell(0).setCellValue(classNumlist.get(i-second));
			}
		}
		
		
		
		for(int i = 0;i<rank.size();i++) {
			if(single) {
				dataBenci.setClassNum(readBenci.getTableData(), singleClassListBenci, singleClassListBenci,rank.get(i));
				setClassNumSheet(dataBenci.getClassNum(),classNumlist,first,i+1,classNumSheet);
			}else {
				if(i<=4) {
					dataBenci.setClassNum(readBenci.getTableData(), singleClassListBenci, singleClassListBenci,rank.get(i));
					setClassNumSheet(dataBenci.getClassNum(),classNumlist,first,i*2+1,classNumSheet);
					
					dataShangci.setClassNum(readShangci.getTableData(), singleClassListShangci,singleClassListBenci, rank.get(i));
					dataBenci.setClassNumRank(dataBenci.getClassNum(), dataShangci.getClassNum(), singleClassListBenci,rank.get(i));
					setClassNumSheet(dataBenci.getClassNumRank(),classNumlist,first,i*2+2,classNumSheet);
				}else {
					dataBenci.setClassNum(readBenci.getTableData(), singleClassListBenci, singleClassListBenci,rank.get(i));
					setClassNumSheet(dataBenci.getClassNum(),classNumlist,second,i*2-9,classNumSheet);
					
					dataShangci.setClassNum(readShangci.getTableData(), singleClassListShangci,singleClassListBenci, rank.get(i));
					dataBenci.setClassNumRank(dataBenci.getClassNum(), dataShangci.getClassNum(), singleClassListBenci,rank.get(i));
					setClassNumSheet(dataBenci.getClassNumRank(),classNumlist,second,i*2-8,classNumSheet);
				}
			}
			
		}
		
	}
	
	public void setTotalAveRankSheet(boolean science,int likeClassSize,int wenkeClassSize,boolean teacher) {
		
		//针对平均分，创建含有标题栏的list
		List<String> aveRanklist = new ArrayList<>();
		aveRanklist.add("班级");
		aveRanklist.add("学科均分");
		aveRanklist.addAll(singleClassListBenci);
		
		//初始化学科信息
		String[] subjectArray = {"语文","数学","理数","文数","数理","数文","英语","物理","化学","生物","政治","历史","地理"};
		List<String> subjectList = new ArrayList<>();
		subjectList.add("应试人数");
		subjectList.add("总分");
		subjectList.add("班级学科均分");
		for(String string:subjectArray) {
			if(readBenci.getTableData().get(1).containsKey(string)) {
				subjectList.add(string);
			};
		}
		
		int first;		
		if(teacher) {
			if(science) {
				first = 1;				
			}else {
				first = likeClassSize+1;
			}
			//创建HSSFRow对象
			for(int i = first;i< aveRanklist.size()+first;i++) {
				aveRankSheet.createRow(i);
			}
			
			if(subjectList.size()<=6) {
				//第一次写入列标题栏
				for(int i=first;i<aveRanklist.size()+first;i++) {
					aveRankSheet.getRow(i).createCell(0).setCellValue(aveRanklist.get(i-first));
				}
			}else if(subjectList.size()<=10){
				for(int i=first;i<aveRanklist.size()+first;i++) {
					aveRankSheet.getRow(i).createCell(0).setCellValue(aveRanklist.get(i-first));
					aveRankSheet.getRow(i).createCell(14).setCellValue(aveRanklist.get(i-first));
				}
			}else {
				for(int i=first;i<aveRanklist.size()+first;i++) {
					aveRankSheet.getRow(i).createCell(0).setCellValue(aveRanklist.get(i-first));
					aveRankSheet.getRow(i).createCell(14).setCellValue(aveRanklist.get(i-first));
					aveRankSheet.getRow(i).createCell(28).setCellValue(aveRanklist.get(i-first));
				}
			}
			//写入其余列
			for(int i = 0;i<subjectList.size();i++ ) {
				if(i==0) {
					dataBenci.setStudentNum(readBenci.getTableData(), singleClassListBenci);
					setAveRankSheet(dataBenci.getStudentNum(),aveRanklist,first,1,aveRankSheet);
					aveRankSheet.setColumnWidth(1, 2420);
				}else if(i==1){
					dataBenci.setSubjectAverage(readBenci.getTableData(), singleClassListBenci, "总分");
					setAveRankSheet(dataBenci.getSubjectAverage(),aveRanklist,first,2,aveRankSheet);
					aveRankSheet.setColumnWidth(2, 2420);
				}else if(i==2) {
					dataBenci.setClassSubjectAverage(dataBenci.getSubjectAverage(), singleClassListBenci, subjectList.size()-2);
					setAveRankSheet(dataBenci.getClassSubjectAverage(),aveRanklist,first,3,aveRankSheet);
					aveRankSheet.setColumnWidth(3, 3320);
					
					dataBenci.setAverageRank(dataBenci.getClassSubjectAverage(), singleClassListBenci);
					setAveRankSheet(dataBenci.getAverageRank(),aveRanklist,first,4,aveRankSheet);
					aveRankSheet.setColumnWidth(4, 1620);
				}else if(i<=5){
					dataBenci.setSubjectTeacher(singleClassListBenci, subjectList.get(i));
					setAveRankSheet(dataBenci.getSubjectTeacher(),aveRanklist,first,i*3-4,aveRankSheet);
					aveRankSheet.setColumnWidth(i*3-4, 2420);
					
					dataBenci.setSubjectAverage(readBenci.getTableData(), singleClassListBenci, subjectList.get(i));
					setAveRankSheet(dataBenci.getSubjectAverage(),aveRanklist,first,i*3-3,aveRankSheet);
					aveRankSheet.setColumnWidth(i*3-3, 2420);
					
					dataBenci.setAverageRank(dataBenci.getSubjectAverage(), singleClassListBenci);
					setAveRankSheet(dataBenci.getAverageRank(),aveRanklist,first,i*3-2,aveRankSheet);
					aveRankSheet.setColumnWidth(i*3-2, 1620);
				}else if(i<=9){
					dataBenci.setSubjectTeacher(singleClassListBenci, subjectList.get(i));
					setAveRankSheet(dataBenci.getSubjectTeacher(),aveRanklist,first,i*3-3,aveRankSheet);
					aveRankSheet.setColumnWidth(i*3-3, 2420);
					
					dataBenci.setSubjectAverage(readBenci.getTableData(), singleClassListBenci, subjectList.get(i));
					setAveRankSheet(dataBenci.getSubjectAverage(),aveRanklist,first,i*3-2,aveRankSheet);
					aveRankSheet.setColumnWidth(i*3-2, 2420);
					
					dataBenci.setAverageRank(dataBenci.getSubjectAverage(), singleClassListBenci);
					setAveRankSheet(dataBenci.getAverageRank(),aveRanklist,first,i*3-1,aveRankSheet);
					aveRankSheet.setColumnWidth(i*3-1, 1620);
				}else{
					dataBenci.setSubjectTeacher(singleClassListBenci, subjectList.get(i));
					setAveRankSheet(dataBenci.getSubjectTeacher(),aveRanklist,first,i*3-1,aveRankSheet);
					aveRankSheet.setColumnWidth(i*3-1, 2420);
					
					dataBenci.setSubjectAverage(readBenci.getTableData(), singleClassListBenci, subjectList.get(i));
					setAveRankSheet(dataBenci.getSubjectAverage(),aveRanklist,first,i*3,aveRankSheet);
					aveRankSheet.setColumnWidth(i*3, 2420);
					
					dataBenci.setAverageRank(dataBenci.getSubjectAverage(), singleClassListBenci);
					setAveRankSheet(dataBenci.getAverageRank(),aveRanklist,first,i*3+1,aveRankSheet);
					aveRankSheet.setColumnWidth(i*3+1, 1620);
				}
			}
		}else {
			int second;
			int third;
			if(science) {
				first = 1;
			}else {
				first = likeClassSize+1;
			}
			second = first + wenkeClassSize+likeClassSize+1;
			third = second + wenkeClassSize+likeClassSize+1;
			//在classNumSheet创建对象
			for(int i = first;i< aveRanklist.size()+first;i++) {
				aveRankSheet.createRow(i);
			}
			for(int i= second;i< aveRanklist.size()+second;i++) {
				aveRankSheet.createRow(i);
			}
			for(int i= third;i< aveRanklist.size()+third;i++) {
				aveRankSheet.createRow(i);
			}
			
			if(subjectList.size()<=6) {
				//第一次写入列标题栏
				for(int i=first;i<aveRanklist.size()+first;i++) {
					aveRankSheet.getRow(i).createCell(0).setCellValue(aveRanklist.get(i-first));
				}
			}else if(subjectList.size()<=11){
				//第一次写入列标题栏
				for(int i=first;i<aveRanklist.size()+first;i++) {
					aveRankSheet.getRow(i).createCell(0).setCellValue(aveRanklist.get(i-first));
				}
				//第二次写入列标题栏
				for(int i=second;i<aveRanklist.size()+second;i++) {
					aveRankSheet.getRow(i).createCell(0).setCellValue(aveRanklist.get(i-second));
				}
			}else {
				//第一次写入列标题栏
				for(int i=first;i<aveRanklist.size()+first;i++) {
					aveRankSheet.getRow(i).createCell(0).setCellValue(aveRanklist.get(i-first));
				}
				//第二次写入列标题栏
				for(int i=second;i<aveRanklist.size()+second;i++) {
					aveRankSheet.getRow(i).createCell(0).setCellValue(aveRanklist.get(i-second));
				}
				//第三次写入列标题栏
				for(int i=third;i<aveRanklist.size()+third;i++) {
					aveRankSheet.getRow(i).createCell(0).setCellValue(aveRanklist.get(i-third));
				}
			}
			for(int i = 0;i<subjectList.size();i++ ) {
				if(i==0) {
					dataBenci.setStudentNum(readBenci.getTableData(), singleClassListBenci);
					setAveRankSheet(dataBenci.getStudentNum(),aveRanklist,first,1,aveRankSheet);
					aveRankSheet.setColumnWidth(1, 2120);
				}else if(i==1){
					dataBenci.setSubjectAverage(readBenci.getTableData(), singleClassListBenci, "总分");
					setAveRankSheet(dataBenci.getSubjectAverage(),aveRanklist,first,2,aveRankSheet);
					aveRankSheet.setColumnWidth(2, 2120);
				}else if(i==2) {
					dataBenci.setClassSubjectAverage(dataBenci.getSubjectAverage(), singleClassListBenci, subjectList.size()-2);
					setAveRankSheet(dataBenci.getClassSubjectAverage(),aveRanklist,first,3,aveRankSheet);
					aveRankSheet.setColumnWidth(3, 3020);
					
					dataBenci.setAverageRank(dataBenci.getClassSubjectAverage(), singleClassListBenci);
					setAveRankSheet(dataBenci.getAverageRank(),aveRanklist,first,4,aveRankSheet);
					aveRankSheet.setColumnWidth(4, 1320);
				}else if(i<=5){
					dataBenci.setSubjectAverage(readBenci.getTableData(), singleClassListBenci, subjectList.get(i));
					setAveRankSheet(dataBenci.getSubjectAverage(),aveRanklist,first,i*2-1,aveRankSheet);
					aveRankSheet.setColumnWidth(i*2-1, 2120);
					
					dataBenci.setAverageRank(dataBenci.getSubjectAverage(), singleClassListBenci);
					setAveRankSheet(dataBenci.getAverageRank(),aveRanklist,first,i*2,aveRankSheet);
					aveRankSheet.setColumnWidth(i*2, 1320);
				}else if(i<=10) {
					dataBenci.setSubjectAverage(readBenci.getTableData(), singleClassListBenci, subjectList.get(i));
					setAveRankSheet(dataBenci.getSubjectAverage(),aveRanklist,second,i*2-11,aveRankSheet);
					
					dataBenci.setAverageRank(dataBenci.getSubjectAverage(), singleClassListBenci);
					setAveRankSheet(dataBenci.getAverageRank(),aveRanklist,second,i*2-10,aveRankSheet);
				}else {
					dataBenci.setSubjectAverage(readBenci.getTableData(), singleClassListBenci, subjectList.get(i));
					setAveRankSheet(dataBenci.getSubjectAverage(),aveRanklist,third,i*2-21,aveRankSheet);
					
					dataBenci.setAverageRank(dataBenci.getSubjectAverage(), singleClassListBenci);
					setAveRankSheet(dataBenci.getAverageRank(),aveRanklist,third,i*2-20,aveRankSheet);
				}
			}
		}		
	}
	
	public void writeWorkbook(boolean science,int likeClassSize,int wenkeClassSize,boolean single,boolean teacher) {				
		setTotalClassNumSheet(science,likeClassSize,wenkeClassSize,single);
		setTotalAveRankSheet(science,likeClassSize,wenkeClassSize,teacher);		
	}
}



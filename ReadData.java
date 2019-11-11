/**
 *Project Name:GradeAnalysis
 *File Name:package-info.java
 *Package Name:com.liuwei.IO
 *Date:2019年3月26日下午8:32:31
 *Copyright (c) 2019,ifly_l@163.com All Rights Reserved.
 */
/**
 * @author Liu Wei
 *
 */
package com.liuwei.IO;
import java.io.File;
import java.io.FileInputStream;
/*
 * 构造文件读取类
 */
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;


public class ReadData{

	private List<HashMap<String, Object>> tableData;

	private int firstRowNum;	//定义获取sheet有数据的首行行数，非第0行
	private int lastRowNum;		//定义获取sheet有数据的尾行行数
	private HSSFRow firstRow; 	//定义获取sheet首行行对象
	private int firstColNum;	//定义获取sheet有数据的首列列数，第0列
	private int lastColNum;		//定义获取sheet有数据的尾列列数
	private int sheetIndex;		//定义sheet的序号
	private HSSFSheet sheet0;	//定义sheet0的内容
	private String[] head;		//定义标题栏数组
			
	//sheetIndex：int型，sheetIndex的get、set方法，传递workbook，返回sheet的序号，若不存在返回-1
	public void setSheetIndex(HSSFWorkbook workbook){
		String tempSheet = "";
		for(int index =0;index < workbook.getNumberOfSheets();index++){
			tempSheet = workbook.getSheetName(index);
			if(tempSheet.equals("理科总成绩")||
			tempSheet.equals("文科")||
			tempSheet.equals("成绩册")||tempSheet.equals("总成绩")) {
				this.sheetIndex = index;
				break;
			}
			else {
				this.sheetIndex = -1;	
			}
		}
	}
	public int getSheetIndex(){
		return sheetIndex;
	}	
			
	//sheet0:HSSFSheet对象，sheet0的get、set方法，传递workbook，获取sheet对象，无法找到指定对应的对象时，抛出异常
	public void setSheet0(int sheetIndex,HSSFWorkbook workbook) {
		if(sheetIndex==-1){
			throw new RuntimeException("未找到对应的sheet");
		}
		else{
			this.sheet0 = workbook.getSheetAt(sheetIndex);
		}
	}
	public HSSFSheet getSheet0(){
		return sheet0;
	}

	//各参数为int型，传递HSSFSheet型变量，设置set、get方法获取工作表中首尾行数、首尾列数
	private void setRowColNum(HSSFSheet sheet0){
		this.firstRowNum = sheet0.getFirstRowNum();
		this.lastRowNum = sheet0.getLastRowNum();
		this.firstRow = sheet0.getRow(firstRowNum);
		this.firstColNum = firstRow.getFirstCellNum();
		this.lastColNum = firstRow.getLastCellNum();	
	}
	private int getFirstRowNum(){
		return firstRowNum;	
	}
	private int getLastRowNum(){
		return lastRowNum;	
	}
	private int getFirstColNum(){
		return firstColNum;	
	}
	private int getLastColNum(){
		return lastColNum;	
	}

	//head:String[]对象，传递首尾列序号，确定数组长度，读取首行信息
	public void setHead(int firstColNum,int lastColNum) {	
		String[] head = new String[lastColNum-firstColNum];
		for(int temp = 0;temp < (lastColNum-firstColNum);temp++) {

			if(firstRow.getCell(temp).getStringCellValue().equals("级位次")) {
				head[temp] = (firstRow.getCell(temp-1).getStringCellValue())+firstRow.getCell(temp).getStringCellValue();
			}
			else if(firstRow.getCell(temp).getStringCellValue().equals("班位次")) {
					head[temp] = (firstRow.getCell(temp-2).getStringCellValue())+firstRow.getCell(temp).getStringCellValue();
			}
			else {
				head[temp] = firstRow.getCell(temp).getStringCellValue();
			}
			
		}
		this.head=head;	
	}
	public String[] getHead(){
		return head;
	}

	//tableData:List对象，传入首尾行、首尾列、首行对象、sheet参数，将表中的数据存入List中
	public void setTableData(int firstRowNum,int lastRowNum,int firstColNum,int LastColNum,HSSFSheet sheet0,String[] head) {
		//创建列的ArrayList对象
		List<HashMap<String,Object>> tableData = new ArrayList<>(); 	
		HSSFCell tempCell;
		//循环读取数据存入hashMap，再将hashMap存入ArrayList中
		for(int i=1;i < (lastRowNum-firstRowNum+1);i++){					//行循环	
			//创建每行的HashMap对象
			Map<String,Object> tableRowData = new HashMap<>();
			for(int j=0;j < (lastColNum-firstColNum-1);j++) {					//列循环
				//判断Cell值是否为空，若空，则设置“”
				tempCell = sheet0.getRow(i).getCell(j);
				if (tempCell==null) {
					sheet0.getRow(i).createCell(j).setCellValue("");
					tempCell = sheet0.getRow(i).getCell(j);
			 	}
			 	
			//判断Cell值类型，并根据值类型读取数据
				switch(tempCell.getCellType()) {
				case NUMERIC:
					tableRowData.put(head[j], tempCell.getNumericCellValue());
					break;
				case STRING:
					tableRowData.put(head[j],tempCell.getStringCellValue());
					break;
				default:
					break;
				}
			}
			//添加每行的HashMap至tableData中
			tableData.add(i-1,(HashMap<String, Object>) tableRowData);
		}
		this.tableData = tableData;
	}
	public List<HashMap<String, Object>> getTableData() {
		return tableData;
	}

	public void setTableData(String path) throws IOException{
		FileInputStream fileInputStream = new FileInputStream(path);
		HSSFWorkbook workbook = new HSSFWorkbook(fileInputStream);
		setSheetIndex(workbook);		//初始化sheetIndex
		setSheet0(getSheetIndex(),workbook);		//初始化sheet0
		setRowColNum(getSheet0());		//初始化行列数rowColNum
		setHead(getFirstColNum(),getLastColNum());	//初始化首行数组
		setTableData(getFirstRowNum(),getLastRowNum(),getFirstColNum(),getLastColNum(),getSheet0(),getHead());	//初始化表格数据
	}
	
	public void setTableData(File file) throws IOException{		
		FileInputStream fileInputStream = new FileInputStream(file);
		HSSFWorkbook workbook = new HSSFWorkbook(fileInputStream);
		setSheetIndex(workbook);		//初始化sheetIndex
		setSheet0(getSheetIndex(),workbook);		//初始化sheet0
		setRowColNum(getSheet0());		//初始化行列数rowColNum
		setHead(getFirstColNum(),getLastColNum());	//初始化首行数组
		setTableData(getFirstRowNum(),getLastRowNum(),getFirstColNum(),getLastColNum(),getSheet0(),getHead());	//初始化表格数据
	}
}
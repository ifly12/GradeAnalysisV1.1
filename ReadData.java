/**
 *Project Name:GradeAnalysis
 *File Name:package-info.java
 *Package Name:com.liuwei.IO
 *Date:2019��3��26������8:32:31
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
 * �����ļ���ȡ��
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

	private int firstRowNum;	//�����ȡsheet�����ݵ������������ǵ�0��
	private int lastRowNum;		//�����ȡsheet�����ݵ�β������
	private HSSFRow firstRow; 	//�����ȡsheet�����ж���
	private int firstColNum;	//�����ȡsheet�����ݵ�������������0��
	private int lastColNum;		//�����ȡsheet�����ݵ�β������
	private int sheetIndex;		//����sheet�����
	private HSSFSheet sheet0;	//����sheet0������
	private String[] head;		//�������������
			
	//sheetIndex��int�ͣ�sheetIndex��get��set����������workbook������sheet����ţ��������ڷ���-1
	public void setSheetIndex(HSSFWorkbook workbook){
		String tempSheet = "";
		for(int index =0;index < workbook.getNumberOfSheets();index++){
			tempSheet = workbook.getSheetName(index);
			if(tempSheet.equals("����ܳɼ�")||
			tempSheet.equals("�Ŀ�")||
			tempSheet.equals("�ɼ���")||tempSheet.equals("�ܳɼ�")) {
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
			
	//sheet0:HSSFSheet����sheet0��get��set����������workbook����ȡsheet�����޷��ҵ�ָ����Ӧ�Ķ���ʱ���׳��쳣
	public void setSheet0(int sheetIndex,HSSFWorkbook workbook) {
		if(sheetIndex==-1){
			throw new RuntimeException("δ�ҵ���Ӧ��sheet");
		}
		else{
			this.sheet0 = workbook.getSheetAt(sheetIndex);
		}
	}
	public HSSFSheet getSheet0(){
		return sheet0;
	}

	//������Ϊint�ͣ�����HSSFSheet�ͱ���������set��get������ȡ����������β��������β����
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

	//head:String[]���󣬴�����β����ţ�ȷ�����鳤�ȣ���ȡ������Ϣ
	public void setHead(int firstColNum,int lastColNum) {	
		String[] head = new String[lastColNum-firstColNum];
		for(int temp = 0;temp < (lastColNum-firstColNum);temp++) {

			if(firstRow.getCell(temp).getStringCellValue().equals("��λ��")) {
				head[temp] = (firstRow.getCell(temp-1).getStringCellValue())+firstRow.getCell(temp).getStringCellValue();
			}
			else if(firstRow.getCell(temp).getStringCellValue().equals("��λ��")) {
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

	//tableData:List���󣬴�����β�С���β�С����ж���sheet�����������е����ݴ���List��
	public void setTableData(int firstRowNum,int lastRowNum,int firstColNum,int LastColNum,HSSFSheet sheet0,String[] head) {
		//�����е�ArrayList����
		List<HashMap<String,Object>> tableData = new ArrayList<>(); 	
		HSSFCell tempCell;
		//ѭ����ȡ���ݴ���hashMap���ٽ�hashMap����ArrayList��
		for(int i=1;i < (lastRowNum-firstRowNum+1);i++){					//��ѭ��	
			//����ÿ�е�HashMap����
			Map<String,Object> tableRowData = new HashMap<>();
			for(int j=0;j < (lastColNum-firstColNum-1);j++) {					//��ѭ��
				//�ж�Cellֵ�Ƿ�Ϊ�գ����գ������á���
				tempCell = sheet0.getRow(i).getCell(j);
				if (tempCell==null) {
					sheet0.getRow(i).createCell(j).setCellValue("");
					tempCell = sheet0.getRow(i).getCell(j);
			 	}
			 	
			//�ж�Cellֵ���ͣ�������ֵ���Ͷ�ȡ����
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
			//���ÿ�е�HashMap��tableData��
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
		setSheetIndex(workbook);		//��ʼ��sheetIndex
		setSheet0(getSheetIndex(),workbook);		//��ʼ��sheet0
		setRowColNum(getSheet0());		//��ʼ��������rowColNum
		setHead(getFirstColNum(),getLastColNum());	//��ʼ����������
		setTableData(getFirstRowNum(),getLastRowNum(),getFirstColNum(),getLastColNum(),getSheet0(),getHead());	//��ʼ���������
	}
	
	public void setTableData(File file) throws IOException{		
		FileInputStream fileInputStream = new FileInputStream(file);
		HSSFWorkbook workbook = new HSSFWorkbook(fileInputStream);
		setSheetIndex(workbook);		//��ʼ��sheetIndex
		setSheet0(getSheetIndex(),workbook);		//��ʼ��sheet0
		setRowColNum(getSheet0());		//��ʼ��������rowColNum
		setHead(getFirstColNum(),getLastColNum());	//��ʼ����������
		setTableData(getFirstRowNum(),getLastRowNum(),getFirstColNum(),getLastColNum(),getSheet0(),getHead());	//��ʼ���������
	}
}
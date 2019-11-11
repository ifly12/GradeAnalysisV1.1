package com.liuwei.IO;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class RWFile {
	
	private String ReadPath;
	private String WritePath;
	private InputStream inputStream;
	private HSSFWorkbook workbook;
	
	//path��·������path��set��get����������String path��setPath����
	public void setReadPath(String path){
		this.ReadPath = path;
	}
	public  String getReadPath() {
		return ReadPath;
		}
	
	public void setWritePath(String path){
		this.WritePath = path;
	}
	public  String getWritePath() {
		return WritePath;
	}
		
	//inputStream��set��get����������setinputString����String path
	public void setinputStream(String path)  throws IOException{
		this.inputStream = new FileInputStream(path);
	}
	public InputStream getinputStream() {
		return inputStream;
	}		

	//workbook��excel����workbook��set��get����������InputStream inputStream(String path)��setworkbook����
	public void setWorkBook(InputStream inputStream ) throws IOException{
		this.workbook = new HSSFWorkbook(inputStream);
	}
	public void setWorkbook(String path) throws IOException{
		setinputStream(path);
		this.workbook = new HSSFWorkbook(getinputStream());
	}
	public HSSFWorkbook getWorkbook() {
		return workbook;
	}
}

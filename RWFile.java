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
	
	//path：路径对象，path的set、get方法，传递String path给setPath方法
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
		
	//inputStream的set、get方法，传递setinputString方法String path
	public void setinputStream(String path)  throws IOException{
		this.inputStream = new FileInputStream(path);
	}
	public InputStream getinputStream() {
		return inputStream;
	}		

	//workbook：excel对象，workbook的set、get方法，传递InputStream inputStream(String path)给setworkbook方法
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

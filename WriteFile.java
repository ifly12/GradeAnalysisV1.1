package com.liuwei.IO;

import java.io.File;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class WriteFile {
	WriteFile(HSSFWorkbook workbook){
		Date now = new Date(); 
		SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd_HH.mm.ss");//���Է�����޸����ڸ�ʽ
		String time = dateFormat.format(now);
		
		File dir = new File("D:\\�ɼ����");
		dir.mkdirs();
		String path = "D:\\�ɼ����\\�ɼ����"+time+".xls";
			try {
	            FileOutputStream fout = new FileOutputStream(path);
	            workbook.write(fout);
	            fout.close();
	        } catch (Exception e) {
	            e.printStackTrace();
	        }
	}
}

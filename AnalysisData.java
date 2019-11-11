package com.liuwei.IO;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;

public class AnalysisData {	
	
	private ArrayList<String> singleClassList;//定义所有单独的班级列表,并调用Sort类进行自然排序
	Sort sort = new Sort();
	//singleClassInformation的set、get方法，输入表格数据
	public void setSingleClassList(List<HashMap<String, Object>> list) {
		ArrayList<String> singleClassList = new ArrayList<>();
		for(HashMap<String,Object> tempMap:list) {
				if(singleClassList.isEmpty()||(!singleClassList.contains(tempMap.get("班级")))) {
					singleClassList.add((String)tempMap.get("班级"));
				}
		}
		sort.naturalSort(singleClassList);
		this.singleClassList = singleClassList;
	}
	public ArrayList<String> getSingleClassList(){
		return singleClassList;
	}
	
	private String num = " ";
	private ArrayList<String> totalRankClassList;//定义总分符合要求的班级列表
	//totalRankClass的set、get方法，输入表格数据及排名数
	public void setTotalRankClassList(List<HashMap<String, Object>> list,int rank) {		
		ArrayList<String> totalRankClass = new ArrayList<>();
		//循坏每一行数据
		for(HashMap<String,Object> tempMap:list) {
			//循环总分校排名列，判断校排名不大于rank
			Set<String> keySet =tempMap.keySet();
			for(String tempString:keySet) {
				if((tempString.equals("总分校排名")||tempString.equals("总分级位次"))
					&&((double)tempMap.get(tempString) <= rank)&&((double)tempMap.get(tempString) != 0)) {
					//循环，获取前rank班级至ArrayList中
					for(String tempClass:keySet) {
						if(tempClass.equals("班级")) {
							totalRankClass.add((String)tempMap.get(tempClass));
						}
					}	
				}
				if(tempString.equals("考号")||tempString.equals("学号")) {
					this.num=tempString;
				}
			}
		}
		this.totalRankClassList = totalRankClass;
	}
	public ArrayList<String> getTotalRankClassList() {
		return totalRankClassList;
	}
	public String getNum() {
		return num;
	}

	
	private HashMap<String,Object> classNum;//定义班级对应的人数总数
	//ClassNum的set、get方法，单独的班级列表和总分符合要求的班级列表
	public void setClassNum(List<HashMap<String, Object>> list,ArrayList<String> singleClassListShangci,ArrayList<String> singleClassListBenci,int rank) {

		HashMap<String,Object> classNumBenci = new HashMap<>();
		
		List<Object> listShangci = new ArrayList<>();
		//初始化总分符合要求的班级列表
		setTotalRankClassList(list,rank);
		ArrayList<String> totalRankClassList = getTotalRankClassList();
		
		
		//把班级信息作为key，0作为value，存储Map
		StringBuilder sBulid = new StringBuilder(); 
		sBulid.append("前");
		sBulid.append(rank);
		sBulid.append("名");

		for(String tempStringShangci:singleClassListShangci) {
			int count = 0;
			for(String tempString:totalRankClassList) {
				if(tempString.equals(tempStringShangci)) {
					count++;
				}
			}
			listShangci.add(count);
		}
		classNumBenci.put("班级",sBulid.toString());
		for(int i=0;i<singleClassListBenci.size();i++) {
			classNumBenci.put(singleClassListBenci.get(i), listShangci.get(i));
		}
		this.classNum = classNumBenci;
		
	}
	public HashMap<String,Object> getClassNum(){
		return classNum;
	}
	
	
	private Map<String,Object> classNumRank;//定义前N名班级人数的变化
	
	public void setClassNumRank(HashMap<String, Object> mapBenci,HashMap<String, Object> mapShangci,ArrayList<String> singleClassList,int rank){
		
		Map<String,Object> classNumRank = new HashMap<>();

		//计算两次前N名信息的比较值，并存入map
		StringBuilder sBulid = new StringBuilder(); 
		sBulid.append("较上次");
		classNumRank.put("班级", sBulid.toString());
		for(String tempclass:singleClassList) {	
			classNumRank.put(tempclass, (Integer)mapBenci.get(tempclass)-(Integer)mapShangci.get(tempclass));
		}
		this.classNumRank = classNumRank;
	}
	
	public Map<String,Object> getClassNumRank(){
		return classNumRank;
	}
	
	private Map<String,Object> studentNum;
	public void setStudentNum(List<HashMap<String, Object>> list,ArrayList<String> singleClassList){
		Map<String,Object> studentNum = new HashMap<>();
		studentNum.put("班级","应试人数");
		studentNum.put("学科均分"," ");
		//存入各班级学科的均分
		for(String tempClass:singleClassList){
			int num = 0;
			for(HashMap<String,Object> tempMap:list){
				if(tempMap.containsValue(tempClass)&&(Double)tempMap.get("总分")!=0) {
					num++;
				}
			}
		studentNum.put(tempClass,num);
		}
		this.studentNum = studentNum;
	}
	public Map<String,Object> getStudentNum(){
		return studentNum;
	}
	
	private Map<String,Object> subjectTeacher;
	public void setSubjectTeacher(ArrayList<String> singleClassList,String subject) {
		Map<String,Object> subjectTeacher = new HashMap<>();
		StringBuilder s = new StringBuilder();
		s.append(subject);
		s.append("教师");
		subjectTeacher.put("班级", s.toString());
		subjectTeacher.put("学科均分", " ");
		for(String tempClass:singleClassList){
			subjectTeacher.put(tempClass, " ");
		}
		this.subjectTeacher = subjectTeacher;
	}
	public Map<String,Object> getSubjectTeacher(){
		return subjectTeacher;
	}
	
	
	
	private Map<String,Object> subjectAverage;
	public void setSubjectAverage(List<HashMap<String, Object>> list,ArrayList<String> singleClassList,String subject){
		Map<String,Object> subjectAverage = new HashMap<>();
		StringBuilder s = new StringBuilder();
		
		if(subject.equals("总分")) {
			s.append("班级");
		}else{
			s.append(subject);
		}
		s.append("均分");
		subjectAverage.put("班级",s.toString());
		
		
		//存入各班级学科的均分
		for(String tempClass:singleClassList){
			double classSum = 0;
			int studentNum = 0;
			for(HashMap<String,Object> tempMap:list){
				if(tempMap.containsValue(tempClass)&&(Double)tempMap.get(subject)!=0) {
					classSum = classSum + (Double)tempMap.get(subject);
					studentNum++;
				}
			}
		subjectAverage.put(tempClass,(double) Math.round((classSum/studentNum)* 100)/100);
		}
		//存入总体的学科均分
		double sum = 0;
		for(String tempClass:singleClassList){
			sum = sum + (double)subjectAverage.get(tempClass);
		}
		subjectAverage.put("学科均分",(double) Math.round((sum/singleClassList.size())* 100)/100);
		this.subjectAverage = subjectAverage;
	}
	public Map<String,Object> getSubjectAverage(){
		return subjectAverage;
	}


	private Map<String,Object> averageRank;
	public void setAverageRank(Map<String,Object> average,ArrayList<String> singleClassList){
		Map<String,Object> averageRank = new HashMap<>();
		averageRank.put("班级","档次");
		averageRank.put("学科均分"," ");
		
		//寻找最大值
		double max = 0;
		for(String tempClass:singleClassList){
			double temp = (double)average.get(tempClass);
			if(temp>max){
				max = temp;
			}
		}
		//计算档次
		for(String tempClass:singleClassList){
			double temp = (double)average.get(tempClass);
			averageRank.put(tempClass,(int)((max-temp)/3+1));
		}
		this.averageRank=averageRank;
	}
	public Map<String,Object> getAverageRank(){
		return averageRank;
	}

	private Map<String,Object> classSubjectAverage;
	public void setClassSubjectAverage(Map<String,Object> average,ArrayList<String> singleClassList,int subjectNum){		
		Map<String,Object> classSubjectAverage = new HashMap<>();
		classSubjectAverage.put("班级","班级学科均分");
		double temp = (double)average.get("学科均分");
		classSubjectAverage.put("学科均分",(double) Math.round((temp/subjectNum)* 100)/100);

		for(String tempClass:singleClassList){	
			temp = (double)average.get(tempClass);
			classSubjectAverage.put(tempClass,(double) Math.round((temp/subjectNum)* 100)/100);
		}
		this.classSubjectAverage = classSubjectAverage;
	}
	public Map<String,Object> getClassSubjectAverage(){
		return classSubjectAverage;
	}

}


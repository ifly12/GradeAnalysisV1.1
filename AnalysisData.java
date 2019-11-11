package com.liuwei.IO;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;

public class AnalysisData {	
	
	private ArrayList<String> singleClassList;//�������е����İ༶�б�,������Sort�������Ȼ����
	Sort sort = new Sort();
	//singleClassInformation��set��get����������������
	public void setSingleClassList(List<HashMap<String, Object>> list) {
		ArrayList<String> singleClassList = new ArrayList<>();
		for(HashMap<String,Object> tempMap:list) {
				if(singleClassList.isEmpty()||(!singleClassList.contains(tempMap.get("�༶")))) {
					singleClassList.add((String)tempMap.get("�༶"));
				}
		}
		sort.naturalSort(singleClassList);
		this.singleClassList = singleClassList;
	}
	public ArrayList<String> getSingleClassList(){
		return singleClassList;
	}
	
	private String num = " ";
	private ArrayList<String> totalRankClassList;//�����ַܷ���Ҫ��İ༶�б�
	//totalRankClass��set��get���������������ݼ�������
	public void setTotalRankClassList(List<HashMap<String, Object>> list,int rank) {		
		ArrayList<String> totalRankClass = new ArrayList<>();
		//ѭ��ÿһ������
		for(HashMap<String,Object> tempMap:list) {
			//ѭ���ܷ�У�����У��ж�У����������rank
			Set<String> keySet =tempMap.keySet();
			for(String tempString:keySet) {
				if((tempString.equals("�ܷ�У����")||tempString.equals("�ּܷ�λ��"))
					&&((double)tempMap.get(tempString) <= rank)&&((double)tempMap.get(tempString) != 0)) {
					//ѭ������ȡǰrank�༶��ArrayList��
					for(String tempClass:keySet) {
						if(tempClass.equals("�༶")) {
							totalRankClass.add((String)tempMap.get(tempClass));
						}
					}	
				}
				if(tempString.equals("����")||tempString.equals("ѧ��")) {
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

	
	private HashMap<String,Object> classNum;//����༶��Ӧ����������
	//ClassNum��set��get�����������İ༶�б���ַܷ���Ҫ��İ༶�б�
	public void setClassNum(List<HashMap<String, Object>> list,ArrayList<String> singleClassListShangci,ArrayList<String> singleClassListBenci,int rank) {

		HashMap<String,Object> classNumBenci = new HashMap<>();
		
		List<Object> listShangci = new ArrayList<>();
		//��ʼ���ַܷ���Ҫ��İ༶�б�
		setTotalRankClassList(list,rank);
		ArrayList<String> totalRankClassList = getTotalRankClassList();
		
		
		//�Ѱ༶��Ϣ��Ϊkey��0��Ϊvalue���洢Map
		StringBuilder sBulid = new StringBuilder(); 
		sBulid.append("ǰ");
		sBulid.append(rank);
		sBulid.append("��");

		for(String tempStringShangci:singleClassListShangci) {
			int count = 0;
			for(String tempString:totalRankClassList) {
				if(tempString.equals(tempStringShangci)) {
					count++;
				}
			}
			listShangci.add(count);
		}
		classNumBenci.put("�༶",sBulid.toString());
		for(int i=0;i<singleClassListBenci.size();i++) {
			classNumBenci.put(singleClassListBenci.get(i), listShangci.get(i));
		}
		this.classNum = classNumBenci;
		
	}
	public HashMap<String,Object> getClassNum(){
		return classNum;
	}
	
	
	private Map<String,Object> classNumRank;//����ǰN���༶�����ı仯
	
	public void setClassNumRank(HashMap<String, Object> mapBenci,HashMap<String, Object> mapShangci,ArrayList<String> singleClassList,int rank){
		
		Map<String,Object> classNumRank = new HashMap<>();

		//��������ǰN����Ϣ�ıȽ�ֵ��������map
		StringBuilder sBulid = new StringBuilder(); 
		sBulid.append("���ϴ�");
		classNumRank.put("�༶", sBulid.toString());
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
		studentNum.put("�༶","Ӧ������");
		studentNum.put("ѧ�ƾ���"," ");
		//������༶ѧ�Ƶľ���
		for(String tempClass:singleClassList){
			int num = 0;
			for(HashMap<String,Object> tempMap:list){
				if(tempMap.containsValue(tempClass)&&(Double)tempMap.get("�ܷ�")!=0) {
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
		s.append("��ʦ");
		subjectTeacher.put("�༶", s.toString());
		subjectTeacher.put("ѧ�ƾ���", " ");
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
		
		if(subject.equals("�ܷ�")) {
			s.append("�༶");
		}else{
			s.append(subject);
		}
		s.append("����");
		subjectAverage.put("�༶",s.toString());
		
		
		//������༶ѧ�Ƶľ���
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
		//���������ѧ�ƾ���
		double sum = 0;
		for(String tempClass:singleClassList){
			sum = sum + (double)subjectAverage.get(tempClass);
		}
		subjectAverage.put("ѧ�ƾ���",(double) Math.round((sum/singleClassList.size())* 100)/100);
		this.subjectAverage = subjectAverage;
	}
	public Map<String,Object> getSubjectAverage(){
		return subjectAverage;
	}


	private Map<String,Object> averageRank;
	public void setAverageRank(Map<String,Object> average,ArrayList<String> singleClassList){
		Map<String,Object> averageRank = new HashMap<>();
		averageRank.put("�༶","����");
		averageRank.put("ѧ�ƾ���"," ");
		
		//Ѱ�����ֵ
		double max = 0;
		for(String tempClass:singleClassList){
			double temp = (double)average.get(tempClass);
			if(temp>max){
				max = temp;
			}
		}
		//���㵵��
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
		classSubjectAverage.put("�༶","�༶ѧ�ƾ���");
		double temp = (double)average.get("ѧ�ƾ���");
		classSubjectAverage.put("ѧ�ƾ���",(double) Math.round((temp/subjectNum)* 100)/100);

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


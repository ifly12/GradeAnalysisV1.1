package com.liuwei.IO;

import java.util.ArrayList;
import java.util.Collections;
import java.util.Iterator;
/*
 * ���ֺ���ĸ�������
 */
public class Sort {
    public  int atoi(String str) {
        //����ҪС�ģ���Ҫ�ж���Ч��
        if (str == null || str.length() == 0) {
            return 0;
        }
        int nlen = str.length();
        double sum = 0;
        int sign = 1;
        int j = 0;
        //�޳��ո�
        while (str.charAt(j) == ' ') {
            j++;
        }
        //�ж������͸���
        if (str.charAt(j) == '+') {
            sign = 1;
            j++;
        } else if (str.charAt(j) == '-') {
            sign = -1;
            j++;
        }
 
        for (int i = j; i < nlen; i++) {
            char current = str.charAt(i);
            if (current >= '0' && current <= '9') {
                sum = sum * 10 + (int) (current - '0');
            } else {
                break;//���������֣��˳�ת��
            }
        }
        sum = sum * sign;
        //����ҪС�ģ���Ҫ�жϷ�Χ
        if (sum > Integer.MAX_VALUE) {
            sum = Integer.MAX_VALUE;
        } else if (sum < Integer.MIN_VALUE) {
            sum = Integer.MIN_VALUE;
        }
        return (int) sum;
    }	
    public void naturalSort(ArrayList<String> list) {
        Collections.sort(list, (o1, o2) -> {
            int i = 0, j = 0;
            String temp1, temp2;
            int num1, num2;
            int length = Math.min(o1.length(), o2.length());
            while (i < length && j < length) {
                temp1 = "";
                temp2 = "";
                if (o1.charAt(i) > '9' || o1.charAt(i) < '0' || o2.charAt(j) > '9' || o2.charAt(j) < '0') {
                    if (o1.charAt(i) == o2.charAt(j)) {
                        i++;
                        j++;
                        continue;
                    } else if (o1.charAt(i) > o2.charAt(j)) {
                        return 1;
                    } else {
                        return -1;
                    }
                }
                while (i < o1.length() && o1.charAt(i) <= '9' && o1.charAt(i) >= '0') {
                    temp1 += o1.charAt(i);
                    i++;
                }
                while (j < o2.length() && o2.charAt(j) <= '9' && o2.charAt(j) >= '0') {
                    temp2 += o2.charAt(j);
                    j++;
                }
                num1 = atoi(temp1);
                num2 = atoi(temp2);
                if (num1 == num2) {
                    if (temp1.length() < temp2.length()) {
                        return 1;
                    } else if (temp1.length() > temp2.length()) {
                        return -1;
                    } else {
                        continue;
                    }
                } else if (num1 > num2) {
                    return 1;
                } else {
                    return -1;
                }
            }
            return o1.length() > o2.length() ? 1 : -1;
        });
    }
	//ȥ���ظ�Ԫ��
	public ArrayList<String> getSingle(ArrayList<?> list) {
        ArrayList<String> tempList = new ArrayList<String>();         //1,�����¼���
        Iterator<?> it = list.iterator();                            //2,���ݴ���ļ���(�ϼ���)��ȡ������ 
        while(it.hasNext()) {                                  //3,�����ϼ���
            String str = (String) it.next();                     //��¼סÿһ��Ԫ��
            if(!tempList.contains(str)) {                        //����¼����в������ϼ����е�Ԫ��
                tempList.add(str);                                //����Ԫ�����
            }
        }         
        return tempList;
    }
}

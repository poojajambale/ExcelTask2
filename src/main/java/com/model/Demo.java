package com.model;

class Data<T>{
	T val;
	Data(T val){
		
	}
}

public class Demo {
	
	public static void main(String args[]) {
		Data data = new Data("Hello");
		data.val = 1;
		System.out.println(data.val.getClass().getSimpleName());
	}
}

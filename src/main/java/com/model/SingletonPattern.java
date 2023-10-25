package com.model;

public class SingletonPattern {
	private static SingletonPattern s = new SingletonPattern();
	
	public static SingletonPattern getInstance() {
		return s;
	}
	
	int count = 0;
	
	void increase() {
		count++;
		System.out.println(count);
	}
	private SingletonPattern() {
		super();
	}
	
}

class Single {
	public static void main(String[] args) {
		
		SingletonPattern ss = SingletonPattern.getInstance();
//		SingletonPattern sss = new SingletonPattern();
	}
}

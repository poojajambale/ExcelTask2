package com.model;

import java.util.Scanner;

public class A3_FactorialOfN {
	
	public static int fac(int a) {    
//		if(a == 1) {
//			return 1;
//		}
//		return fac(a-1)*a;	
		
		if(a > 1) {
			return fac(a-1)*a;
		}
		else {
			return 1;
		}
	}

	//fac(4) = 6*4;
//	fac(3) = 2*3;
	//fac(2) = 1*2;
//	fac(1) = 1;
	
	public static void main(String[] args) {

		// Take scanner for inputs
		Scanner sc = new Scanner(System.in);

		System.out.println("Enter a Number");
		int n = sc.nextInt();
		
		int resultt = fac(n);
		System.out.println("resultt :"+ resultt );
		
		int result = 1;
		
		for (int i = 1; i <= n; i++) {
			result *= i;
		}
		
		System.out.println("Result = "+ result);
		
		sc.close();

	}
}

package com.amex.TS.AutomationTesting.GCO;

import java.util.concurrent.atomic.AtomicInteger;

import org.testng.IRetryAnalyzer;
import org.testng.ITestResult;

public class Retry implements IRetryAnalyzer {
	
	private int count = 0;
	private int maxCount = 3;
	
	private static int MAX_RETRY_COUNT = 0;
	
	public boolean isRetryAvailable(){
		return counter.intValue() > 0;
	}
	
	AtomicInteger counter = new AtomicInteger(MAX_RETRY_COUNT);
	
	@Override
	public boolean retry(ITestResult result){
		boolean retry = false;
		if(isRetryAvailable()){
			retry = true;
			counter.decrementAndGet();
		}
		return retry;
	}

}

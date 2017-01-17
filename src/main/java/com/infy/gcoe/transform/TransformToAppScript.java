package com.infy.gcoe.transform;

import java.util.List;
import java.util.Set;

import org.springframework.boot.ApplicationArguments;
import org.springframework.boot.CommandLineRunner;
import org.springframework.context.annotation.Profile;
import org.springframework.stereotype.Component;

@Component
@Profile(value="TransformToAppScript")
public class TransformToAppScript implements CommandLineRunner {
	
	public TransformToAppScript(ApplicationArguments args){
		System.out.println("Inside constructor TransformToAppScript...........");
		
		boolean path = args.containsOption("path");
		System.out.println("Has path Variable " + path);
		
        List<String> files = args.getNonOptionArgs();
        for(String name:files){
        	System.out.println(name);
        }

        Set<String> optionNames = args.getOptionNames();
        for(String name:optionNames){
        	System.out.println(name);
        }
	}

	@Override
	public void run(String[] args) throws Exception {
		System.out.println("Inside run method of TransformToAppScript...........");
		if(args != null){
			System.out.println("Inside TransformToAppScript run() : ");
			for(int i=0;i<args.length;i++){
				System.out.println(args[i]);
			}
		}
		
	}
}

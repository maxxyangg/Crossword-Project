package com.crossword.driver;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStreamReader;
import java.io.UnsupportedEncodingException;
import java.util.ArrayList;
import java.util.Iterator;

public class Reader {
	private String path;
	
	public Reader(String path) {
		this.path = path;
	}
	
	public ArrayList<String> readFileIntoArray(ArrayList<String> wordList) {
		File file = new File(getPath());
		try {
			BufferedReader in = new BufferedReader(new InputStreamReader(new FileInputStream(file), "UTF-8"));
			String str;
			while((str = in.readLine()) != null) {
				wordList.add(str);
			}
			in.close();
		} catch (UnsupportedEncodingException e) {
			e.printStackTrace();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		return wordList;
	}
	
	public static void main(String[] args) {
		ArrayList<String> test = new ArrayList<String>();
		Reader reader = new Reader("E:\\Folders\\Classes\\ICS499\\Files\\english_crossword_words.csv");
		reader.readFileIntoArray(test);
		String str = "train,\"A type of transportation\"";
		String[] strArr = str.split(",");
		System.out.println(strArr[1]);
//		for(int i = 0; i < test.size(); i++) {
//			System.out.println(test.get(i));
//		}
	}

	public String getPath() {
		return path;
	}

	public void setPath(String path) {
		this.path = path;
	}

}

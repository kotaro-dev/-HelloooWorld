package mylocal.disk01.study.main;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.IOException;

public class Sample02 {

	private static final String LINE_SEPARETOR = System.getProperty("line.separetor");

	public static void main(String[] args) {
		// TODO 自動生成されたメソッド・スタブ
		//DB関連のSample
		new Sample02().execute();
	}

	private void execute() {
		System.out.println(readFromFile("sample02.txt"));
	}

	private String readFromFile(String fileName) {
		// TODO 自動生成されたメソッド・スタブ
		File file = new File(fileName);
		StringBuilder sb = new StringBuilder();

		try (
			FileReader fr = new FileReader(file);
			BufferedReader br = new BufferedReader(fr);
		) {
			String line;
			while ((line = br.readLine()) != null) {
				sb.append(line).append(LINE_SEPARETOR);
			}
			return sb.toString();

		} catch (FileNotFoundException e) {
			// TODO: handle exception
			System.err.println();
			System.err.println(String.format("file not found... : %s", fileName));
			return null;
		} catch (IOException e) {
			// TODO: handle exception
			System.err.println(String.format("Could not open a file... : %s", fileName));
			return null;
		}

	}


}

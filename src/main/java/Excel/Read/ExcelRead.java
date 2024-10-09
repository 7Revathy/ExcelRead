package Excel.Read;

import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

public class ExcelRead {

	public static XSSFWorkbook w;



	public static XSSFSheet s;



	public static FileInputStream f;



	public static String readStringData(int i,int j) throws IOException { // chance for error file not exception parent or can use filenotexception
// return type is string  i w.r.t row,j w.r.t coloum


	f= new FileInputStream("D:\\Java Class\\Programs\\ExcelReadMavenProject\\src\\main\\resources\\Student.xlsx"); //inform to system about the path



	w= new XSSFWorkbook(f); //object creation fileinput screen f // eppa excel kitti



	s= w.getSheet("Sheet1"); // s= sheet , getsheet =mention the sheet name to method getsheet

	Row r=s.getRow(i); // mention the s sheet le i th row aduthu r ill vekukka



	Cell c=r.getCell(j); // r ill row ille j th cell aduthu cell ill vekukka
	


	return c.getStringCellValue(); // method for cell ta inside ill ninum string value read chaiyan , then return that string value



	}



	public static String readIntegerData(int i,int j) throws IOException { 

		



			f= new FileInputStream("D:\\Java Class\\Programs\\ExcelReadMavenProject\\src\\main\\resources\\Student.xlsx");



			w= new XSSFWorkbook(f);



			s= w.getSheet("Sheet1");



			Row r=s.getRow(i);



			Cell c=r.getCell(j);



			int value=(int) c.getNumericCellValue();// typecaste bz numeric value anu bz numeic value may be any booloue ,flot



			return String.valueOf(value);//string varan karanam return type string anu so again type cast chaiyanam



			}

	}
package com.qa;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.TestDataReader;
/**
 * Hello world!
 *
 */
public class App 
{
    public static void main( String[] args ) throws IOException
    {
        System.out.println( "Hello World! :: " +System.getProperty("user.dir") );
        String dataFile = System.getProperty("user.dir")+ File.separatorChar+"demoexcel"+File.separatorChar+"TestData.xlsx";
        //String dataFile1 = System.getProperty("user.dir")+ File.separatorChar+"demoexcel"+File.separatorChar+"TestData2.xlsx";
        //String dataFile = "D:\SeleniumLearning\Maven1\demoexcel\TestData.xlsx";
        System.out.println( "dataFile :: "+dataFile );
        TestDataReader td = new TestDataReader();
        td.writeExcelFile(dataFile);
        td.readData(dataFile);
    }
  
    
}

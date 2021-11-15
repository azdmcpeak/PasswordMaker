import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.
import java.util.Scanner;

public class ExcelMain {apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import java.io.*;
import java.util.ArrayList;
import java.util.Random;

    public static void main(String[] args) throws IOException {

        Scanner input = new Scanner(System.in);
        //initializing scanner for user input = new Scanner(System.in);

        //initializing random for use in arraylist
        Random rand = new Random();

        //array list to hold contents for password
        ArrayList<String> passwordcontents = new ArrayList<>() ;

        passwordcontents.add("a");
        passwordcontents.add("A");
        passwordcontents.add("b");
        passwordcontents.add("B");
        passwordcontents.add("c");
        passwordcontents.add("C");
        passwordcontents.add("d");
        passwordcontents.add("D");
        passwordcontents.add("e");
        passwordcontents.add("E");
        passwordcontents.add("f");
        passwordcontents.add("F");
        passwordcontents.add("g");
        passwordcontents.add("G");
        passwordcontents.add("h");
        passwordcontents.add("H");
        passwordcontents.add("i");
        passwordcontents.add("I");
        passwordcontents.add("j");
        passwordcontents.add("J");
        passwordcontents.add("k");
        passwordcontents.add("K");
        passwordcontents.add("l");
        passwordcontents.add("L");
        passwordcontents.add("m");
        passwordcontents.add("M");
        passwordcontents.add("n");
        passwordcontents.add("N");
        passwordcontents.add("o");
        passwordcontents.add("O");
        passwordcontents.add("p");
        passwordcontents.add("P");
        passwordcontents.add("q");
        passwordcontents.add("Q");
        passwordcontents.add("r");
        passwordcontents.add("R");
        passwordcontents.add("s");
        passwordcontents.add("S");
        passwordcontents.add("t");
        passwordcontents.add("T");
        passwordcontents.add("u");
        passwordcontents.add("U");
        passwordcontents.add("v");
        passwordcontents.add("V");
        passwordcontents.add("w");
        passwordcontents.add("W");
        passwordcontents.add("x");
        passwordcontents.add("X");
        passwordcontents.add("y");
        passwordcontents.add("Y");
        passwordcontents.add("z");
        passwordcontents.add("Z");
        passwordcontents.add("0");
        passwordcontents.add("1");
        passwordcontents.add("2");
        passwordcontents.add("3");
        passwordcontents.add("4");
        passwordcontents.add("5");
        passwordcontents.add("6");
        passwordcontents.add("7");
        passwordcontents.add("8");
        passwordcontents.add("9");
        


        //for reading the file
        FileInputStream fis;
        //for writing to the file
        FileOutputStream fos;
        //file name
        String filepath = "./testdata.xlsx";

        //reading the file
        fis = new FileInputStream("./testdata.xlsx");

        //creating work book
        XSSFWorkbook workbook = new XSSFWorkbook(fis);

        //Reading from sheet 1
        XSSFSheet sheet = workbook.getSheetAt(0);

        //checking current max number of rows
        int numberofrows = sheet.getLastRowNum();
            //conformation of number of rows.
            //System.out.println(numberofrows);

        //taking current number of rows, increasing it by 1 to create a new cell and not over write an existing one.
        numberofrows = numberofrows + 1;
            //conformation that number of rows is increasing.
            //System.out.println(numberofrows);



        //set company Start
        System.out.println("What is the company name?");
        String company_name = input.nextLine();


        sheet.createRow(numberofrows);

        //creating a row to the next available empty row
        XSSFRow row = sheet.getRow(numberofrows);

        //Sets what column to write to
        XSSFCell cell = row.createCell(0);

        //writing the cell, the users input
        cell.setCellValue(company_name);

        //creating output with file name for writing
        fos  = new FileOutputStream(filepath);

        //writing of the file
        workbook.write(fos);
        //set company End
            ///////
            ///////
        //set Username Start
        System.out.println("What is the User name?");
        String Username = input.nextLine();

        //setting column number to 1
        cell = row.createCell(1);

        //setting cell value to input of the user
        cell.setCellValue(Username);

        //creating output with file name for writing
        fos  = new FileOutputStream(filepath);

        //writing of the file
        workbook.write(fos);

        //set Username End
            ///////
            ///////
        //set Password Start
        System.out.println("Password length: ");
        int Pword_length = input.nextInt();

        //initializing blank string to hold the password
        String Password = "";

        for (int i = 0; i<Pword_length;i++){
            Password += passwordcontents.get(rand.nextInt(passwordcontents.size()));
        }

        System.out.println(Password);

        //setting cell column to 2
        cell = row.createCell(2);

        //writing password to cell
        cell.setCellValue(Password);


        //creating output with file name for writing
        fos  = new FileOutputStream(filepath);

        //writing of the file
        workbook.write(fos);




        //closing file
        fos.close();

    }


}

package com.example.CaseTask;

import org.apache.poi.ooxml.POIXMLException;

import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Scanner;

public class Main {
    static ExcelReader reader;

    public static void main(String[] args) throws IOException {
        {
            fileInput();
        }
    }

    static void fileInput() throws IOException {
        try {
            System.out.print("Write file path: ");
            String filePath = new Scanner(System.in).nextLine();
            reader = new ExcelReader(filePath);
            reader.formatData();
            System.out.println("Result is saved to: " + filePath.split(".xlsx")[0] + "_result.xlsx");
        } catch (FileNotFoundException e) {
            System.out.println("File not found or result file is opened");
            fileInput();
        } catch (POIXMLException e) {
            System.out.println("Incorrect file type");
            fileInput();
        } catch (ExcelException e) {
            System.out.println(e.getMessage());
            fileInput();
        }
    }
}
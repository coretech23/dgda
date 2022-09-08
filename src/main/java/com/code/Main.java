package com.code;


import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class Main {
    public static void main(String[] args) throws IOException {
        PrintStream out = new PrintStream(new FileOutputStream("output.txt"));
        System.setOut(out);

        String fileLocation = "C:\\Users\\Nuxeo\\Desktop\\DGDA.xlsx";
        FileInputStream file = new FileInputStream(new File(fileLocation));
        Workbook workbook = new XSSFWorkbook(file);

        Sheet sheet = workbook.getSheetAt(0);
        var deps = getDepartments(sheet);

        //sheet = workbook.getSheetAt(1);
        // setDepartmentCode(sheet, deps);
        // printDeps(deps);

        createExcelFile(deps);

    }

    static List<Department> getDepartments(Sheet sheet) {

        List<Department> deps = new ArrayList<>();
        int i = 0;
        List<String>lastParentDep= Arrays.asList(null,null,null,null,null,null);
        for (Row row : sheet) {
            if (i++ < 2) continue;
            for (int j = 0; j < row.getPhysicalNumberOfCells(); j++) {
                var cell = row.getCell(j);
                if (cell == null) continue;
                String currCellValue = cell.getStringCellValue();
                if (currCellValue.isEmpty()) continue;

                int splitIdx = returnSplitIdxBetweenArabicAndEnglish(currCellValue);

                //English & Arabic Name
                String arabicName = currCellValue.substring(0, splitIdx);
                arabicName = trimString(arabicName);
                String englishName = currCellValue.substring(splitIdx);
                englishName = trimString(englishName);

                //Parent Code
                lastParentDep.set(j,englishName);
                String parentDep=j==1? "" :lastParentDep.get(j-1);

                //Path(name)
                String path="";
                for(int k=1;k<=j;k++) {
                    String currDep=lastParentDep.get(k);
                    currDep=currDep.replaceAll("\\s+","");
                    path=path.concat(currDep + "/");

                }
                path=path.substring(0,path.length()-1);

                var currDep = new Department(null, arabicName, englishName,path,parentDep);
                deps.add(currDep);
            }
        }
        return deps;
    }

    static int returnSplitIdxBetweenArabicAndEnglish(String currCellValue) {
        int splitIdx = -1;
        for (int c = 0; c < currCellValue.length(); c++) {
            var currChar = currCellValue.charAt(c);
            if (Character.compare(currChar, 'a') >= 0 && Character.compare(currChar, 'z') <= 0 ||
                    Character.compare(currChar, 'A') >= 0 && Character.compare(currChar, 'Z') <= 0) {
                splitIdx = c;
                break;
            }

        }
        return splitIdx;


    }

    static String trimString(String currString) {
        String formatString = currString.trim().replaceAll("\\n", " ");
        return formatString;
    }

    static void setDepartmentCode(Sheet sheet, List<Department> deps) {

        List<String>notFoundDepartment=new ArrayList<>();
        for (int r = 1; r < sheet.getPhysicalNumberOfRows(); r++) {
            Row currRow = sheet.getRow(r);
            var currCell = currRow.getCell(5);
            if (currCell == null) continue;
            String value = currCell.getStringCellValue();
            if (value.isEmpty()) continue;

            var depCode = currRow.getCell(4).getNumericCellValue();

            int splitIdx = returnSplitIdxBetweenArabicAndEnglish(value);
            String arabicName = value.substring(0, splitIdx);
            arabicName = trimString(arabicName);

            String englishName = value.substring(splitIdx);
            englishName = trimString(englishName);

            boolean found = false;
            for (var dep : deps)
                if (dep.arabicName.toLowerCase().contains(arabicName) || dep.englishName.toLowerCase().contains(englishName)) {
                    found = true;
                    dep.code = String.valueOf(depCode);
                    System.out.println(dep.arabicName + " | " + dep.englishName + " | " + depCode);

                }


            if (!found)
                notFoundDepartment.add(arabicName + " | " + englishName + " | " + depCode);


        }
    }

    static void printDeps(List<Department>deps){
        for (var dep : deps) {
            System.out.println(dep.arabicName + " | " + dep.englishName + " | " + dep.code);

        }
    }

    static void createExcelFile(List<Department>deps) throws IOException {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Sheet 1");
        Row header = sheet.createRow(0);

        header.createCell(0).setCellValue("name");
        header.createCell(1).setCellValue("dc:title");
        header.createCell(2).setCellValue("osdept:arabicName");
        header.createCell(3).setCellValue("osdept:englishName");
        header.createCell(4).setCellValue("type");
        header.createCell(5).setCellValue("osdept:parentDepartmentCode");


       for(int i=0;i<deps.size();i++){
           Row row=sheet.createRow(i+1);
           row.createCell(0).setCellValue(deps.get(i).path);
           row.createCell(1).setCellValue(deps.get(i).englishName);
           row.createCell(2).setCellValue(deps.get(i).arabicName.equals("")?deps.get(i).englishName:deps.get(i).arabicName);
           row.createCell(3).setCellValue(deps.get(i).englishName);
           row.createCell(4).setCellValue("OSDepartment");
           row.createCell(5).setCellValue(deps.get(i).parentCode);

       }

        String fileLocation = "C:\\Users\\Nuxeo\\Desktop\\DGDAdeps.xlsx";
        FileOutputStream outputStream = new FileOutputStream(fileLocation);
        workbook.write(outputStream);
        workbook.close();

    }
}


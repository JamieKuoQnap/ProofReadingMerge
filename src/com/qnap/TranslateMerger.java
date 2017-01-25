package com.qnap;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.Arrays;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class TranslateMerger {
    private static String[] resultLanguage = { "ENG", "TCH", "SCH", "CZE", "DAN", "GER", "SPA", "FRE", "ITA", "JPN",
            "KOR", "NOR", "POL", "RUS", "FIN", "SWE", "DUT", "TUR", "THA", "HUN", "GRK", "ROM", "POR" };

    public static void main(String[] args) {
        if (args.length == 0) {
            System.err.println("Enter folder path where xlsx exist!");
            return;
        }
        String folderPath = args[0];

        Map<String, String> headerLanguageResultLanguage = new HashMap<String, String>();
        headerLanguageResultLanguage.put("ENG", "ENG");
        headerLanguageResultLanguage.put("EN", "ENG");
        headerLanguageResultLanguage.put("English", "ENG");
        
        headerLanguageResultLanguage.put("TCH", "TCH");
        headerLanguageResultLanguage.put("CHT", "TCH");
        headerLanguageResultLanguage.put("Traditonal Chinese", "TCH");
        headerLanguageResultLanguage.put("Reviewed TCH", "TCH");
        
        headerLanguageResultLanguage.put("SCH", "SCH");
        headerLanguageResultLanguage.put("CHS", "SCH");
        headerLanguageResultLanguage.put("Simplified Chinese", "SCH");

        headerLanguageResultLanguage.put("French", "FRE");
        headerLanguageResultLanguage.put("FRE", "FRE");

        headerLanguageResultLanguage.put("Italian", "ITA");
        headerLanguageResultLanguage.put("ITA", "ITA");

        headerLanguageResultLanguage.put("Polish", "POL");

        headerLanguageResultLanguage.put("Czech", "CZE");
        headerLanguageResultLanguage.put("Dutch", "DUT");
        headerLanguageResultLanguage.put("Spanish (Spain)", "SPA");
        headerLanguageResultLanguage.put("Swedish", "SWE");
        headerLanguageResultLanguage.put("Turkish", "TUR");

        headerLanguageResultLanguage.put("Danish", "DAN");
        headerLanguageResultLanguage.put("DANISH", "DAN");
        headerLanguageResultLanguage.put("Finnish", "FIN");
        headerLanguageResultLanguage.put("FINNISH", "FIN");
        headerLanguageResultLanguage.put("GERMAN", "GER");
        headerLanguageResultLanguage.put("German", "GER");
        headerLanguageResultLanguage.put("Greek", "GRK");
        headerLanguageResultLanguage.put("GREEK", "GRK");
        headerLanguageResultLanguage.put("Hungarian", "HUN");
        headerLanguageResultLanguage.put("HUNGARIAN", "HUN");
        headerLanguageResultLanguage.put("Japanese", "JPN");
        headerLanguageResultLanguage.put("JAPANESE", "JPN");
        headerLanguageResultLanguage.put("Korean", "KOR");
        headerLanguageResultLanguage.put("KOREAN", "KOR");
        headerLanguageResultLanguage.put("Norwegian", "NOR");
        headerLanguageResultLanguage.put("NORWEGIAN", "NOR");
        headerLanguageResultLanguage.put("Portuguese(Brazil)", "POR");
        headerLanguageResultLanguage.put("Portuguese Brazil", "POR");
        headerLanguageResultLanguage.put("PORTUGUESE(BRAZIL)", "POR");
        headerLanguageResultLanguage.put("Romanian", "ROM");
        headerLanguageResultLanguage.put("ROMANIAN", "ROM");
        headerLanguageResultLanguage.put("Russian", "RUS");
        headerLanguageResultLanguage.put("RUSSIAN", "RUS");
        headerLanguageResultLanguage.put("Thai", "THA");
        headerLanguageResultLanguage.put("THAI", "THA");

        Set<String> languageHeaderSet = headerLanguageResultLanguage.keySet();

        Map<String, Map<String, String>> resultMap = new LinkedHashMap<String, Map<String, String>>();

        File targetDirectory = new File(folderPath);
        if (targetDirectory.isDirectory() == true) {
            File[] files = targetDirectory.listFiles();
            for (File file : files) {
                String fileName = file.getName();
                if (fileName.endsWith("xlsx")) {
                    FileInputStream fileInputStream = null;
                    XSSFWorkbook inputWorkbook = null;
                    try {
                        fileInputStream = new FileInputStream(file.getAbsolutePath());
                        inputWorkbook = new XSSFWorkbook(fileInputStream);
                        XSSFSheet xssfSheet = inputWorkbook.getSheetAt(0);

                        int firstRowNum = xssfSheet.getFirstRowNum();
                        int lastRowNum = xssfSheet.getLastRowNum();

                        XSSFRow firstRow = xssfSheet.getRow(firstRowNum);
                        int lastCellNumberOfTheFirstRow = firstRow.getLastCellNum();
                        Map<Integer, String> languagePosition = new HashMap<Integer, String>();
                        for (int i = 1; i < lastCellNumberOfTheFirstRow; i++) {
                            XSSFCell xssfCell = firstRow.getCell(i);
                            if (xssfCell != null) {
                                String languageHeader = xssfCell.getStringCellValue().trim();
                                if (languageHeaderSet.contains(languageHeader) == true) {
                                    languagePosition.put(i, languageHeader);
                                }
                            }
                        }
                        for (int i = 1; i <= lastRowNum; i++) {
                            XSSFRow row = xssfSheet.getRow(i);
                            XSSFCell keyCell = row.getCell(0);
                            if (keyCell == null) {
                                continue;
                            }
                            String key;
                            try {
                                key = keyCell.getStringCellValue();
                            } catch (Exception e) {
                                System.err.println(e);
                                key = new Double(keyCell.getNumericCellValue()).toString();
                            }
                            if (key != null && key.isEmpty() == false) {
                                if (resultMap.containsKey(key) == false) {
                                    Map<String, String> valueOfKeyMap = new HashMap<String, String>();
                                    resultMap.put(key.trim(), valueOfKeyMap);
                                }
                            }

                            int lastCellNumber = row.getLastCellNum();
                            for (int j = 1; j < lastCellNumber; j++) {
                                XSSFCell languageValueCell = row.getCell(j);
                                if (languageValueCell != null) {
                                    try {
                                        String value = languageValueCell.getStringCellValue();
                                        if (value != null && value.isEmpty() == false) {
                                            Map<String, String> valueOfKeyMap = resultMap.get(key.trim());
                                            String resultLanguageKey = headerLanguageResultLanguage
                                                    .get(languagePosition.get(j));
                                            if (resultLanguageKey != null) {
                                                valueOfKeyMap.put(resultLanguageKey.trim(), value.trim());
                                            }
                                        }
                                    } catch (Exception e) {
                                        e.printStackTrace();
                                    }
                                }
                            }
                        }
                    } catch (Exception e) {
                        e.printStackTrace();
                    } finally {
                        if (inputWorkbook != null) {
                            try {
                                inputWorkbook.close();
                            } catch (Exception e) {
                                e.printStackTrace();
                            }
                        }
                        if (fileInputStream != null) {
                            try {
                                fileInputStream.close();
                            } catch (Exception e) {
                                e.printStackTrace();
                            }
                        }
                    }
                }
            }
            System.out.println(resultMap);
            FileOutputStream fileOutputStream = null;
            XSSFWorkbook outputWorkbook = null;
            try {
                List<String> resultLanguageList = Arrays.asList(resultLanguage);
                outputWorkbook = new XSSFWorkbook();
                XSSFSheet sheet = outputWorkbook.createSheet("result");
                XSSFRow headerRow = sheet.createRow(0);
                for (int i = 0; i < resultLanguage.length; i++) {
                    XSSFCell cell = headerRow.createCell(i + 1);
                    cell.setCellValue(resultLanguage[i]);
                }
                int i = 1;
                for (Map.Entry<String, Map<String, String>> entry : resultMap.entrySet()) {
                    XSSFRow row = sheet.createRow(i++);
                    row.createCell(0).setCellValue(entry.getKey());
                    Map<String, String> languageValueMap = entry.getValue();
                    for (Map.Entry<String, String> languageValueEntry : languageValueMap.entrySet()) {
                        String languageKey = languageValueEntry.getKey();
                        String value = languageValueEntry.getValue();
                        int index = resultLanguageList.indexOf(languageKey);
                        if (index > -1) {
                            XSSFCell valueCell = row.createCell(index + 1);
                            valueCell.setCellValue(value);
                        }
                    }
                }

                fileOutputStream = new FileOutputStream(folderPath + "/result.xlsx");
                outputWorkbook.write(fileOutputStream);
                fileOutputStream.flush();
            } catch (Exception e) {
                e.printStackTrace();
            } finally {
                if (outputWorkbook != null) {
                    try {
                        outputWorkbook.close();
                    } catch (Exception e) {
                        e.printStackTrace();
                    }
                }
                if (fileOutputStream != null) {
                    try {
                        fileOutputStream.close();
                    } catch (Exception e) {
                        e.printStackTrace();
                    }
                }
            }
        }

    }

}

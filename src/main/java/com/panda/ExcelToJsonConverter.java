package com.panda;

/**
 * @Author Victor
 * @Date 2024-10-17 10:18 a.m.
 */

import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.Map;

public class ExcelToJsonConverter {

    public static void main(String[] args) {
        // 输入 Excel 文件路径
        String excelFilePath = "Languages.xlsx";
        // 输出 JSON 文件目录
        String outputDirectory = "output_json/";

        try {
            // 检查并创建输出目录
            createDirectoryIfNotExists(outputDirectory);
            // 将 Excel 文件转换为 JSON
            convertExcelToJson(excelFilePath, outputDirectory);
            System.out.println("JSON files generated successfully.");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * 将 Excel 文件转换为多语言 JSON 文件，保持键值顺序
     *
     * @param excelFilePath  Excel 文件路径
     * @param outputDirectory 输出的 JSON 文件路径
     * @throws IOException
     */
    public static void convertExcelToJson(String excelFilePath, String outputDirectory) throws IOException {
        // 打开 Excel 文件
        FileInputStream inputStream = new FileInputStream(new File(excelFilePath));
        Workbook workbook = new XSSFWorkbook(inputStream);
        Sheet sheet = workbook.getSheetAt(0); // 假设数据在第一个工作表中

        // 存储不同语言的 JSON 数据，并保持顺序
        Map<String, Map<String, String>> languageMaps = new LinkedHashMap<>();

        // 获取表头
        Row headerRow = sheet.getRow(0);
        int numberOfLanguages = headerRow.getLastCellNum() - 1; // 除去第一列的 key 列

        // 初始化每个语言的 Map，使用 LinkedHashMap 保持顺序
        for (int colIndex = 1; colIndex <= numberOfLanguages; colIndex++) {
            String language = headerRow.getCell(colIndex).getStringCellValue();
            languageMaps.put(language, new LinkedHashMap<>());
        }

        // 遍历每一行，将 key 和多语言值分别存储
        Iterator<Row> rowIterator = sheet.iterator();
        rowIterator.next(); // 跳过表头

        while (rowIterator.hasNext()) {
            Row row = rowIterator.next();
            String key = row.getCell(0).getStringCellValue();

            for (int colIndex = 1; colIndex <= numberOfLanguages; colIndex++) {
                String language = headerRow.getCell(colIndex).getStringCellValue();
                Cell cell = row.getCell(colIndex);
                String value = cell != null ? cell.getStringCellValue() : "";
                languageMaps.get(language).put(key, value);
            }
        }

        // 将每种语言的数据导出为对应的 JSON 文件，按插入顺序输出
        ObjectMapper objectMapper = new ObjectMapper();
        for (Map.Entry<String, Map<String, String>> entry : languageMaps.entrySet()) {
            String language = entry.getKey();
            Map<String, String> jsonMap = entry.getValue();

            // 创建 JSON 文件并输出，保持顺序
            File outputFile = new File(outputDirectory + language + ".json");
            objectMapper.writerWithDefaultPrettyPrinter().writeValue(outputFile, jsonMap);
        }

        workbook.close();
        inputStream.close();
    }

    /**
     * 检查文件夹是否存在，如果不存在则创建
     *
     * @param directoryPath 文件夹路径
     */
    public static void createDirectoryIfNotExists(String directoryPath) {
        File directory = new File(directoryPath);
        if (!directory.exists()) {
            if (directory.mkdirs()) {
                System.out.println("Directory created: " + directoryPath);
            } else {
                System.err.println("Failed to create directory: " + directoryPath);
            }
        }
    }
}



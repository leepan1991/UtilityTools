package com.panda;

/**
 * @Author Victor
 * @Date 2024-10-17 10:16 a.m.
 */

import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.Iterator;

public class JsonToExcelConverter {

    public static void main(String[] args) {
        // 输入 JSON 文件路径
        String jsonFilePath = "zh.json";
        // 输出 Excel 文件路径
        String excelFilePath = "Languages.xlsx";

        try {
            // 将 JSON 文件转换为 Excel
            convertJsonToExcel(jsonFilePath, excelFilePath);
            System.out.println("Excel file generated successfully.");
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * 将 JSON 文件转换为 Excel
     *
     * @param jsonFilePath  JSON 文件路径
     * @param excelFilePath 输出的 Excel 文件路径
     * @throws IOException
     */
    public static void convertJsonToExcel(String jsonFilePath, String excelFilePath) throws IOException {
        // 创建 ObjectMapper 实例来解析 JSON
        ObjectMapper objectMapper = new ObjectMapper();

        // 读取 JSON 文件
        // 使用 ClassLoader 从 resources 文件夹加载 JSON 文件
        ClassLoader classLoader = Thread.currentThread().getContextClassLoader();
        InputStream inputStream = classLoader.getResourceAsStream(jsonFilePath);
        if (inputStream == null) {
            throw new IllegalArgumentException("File not found! " + jsonFilePath);
        }
        JsonNode rootNode = objectMapper.readTree(inputStream);

        // 创建一个新的 Excel 工作簿
        Workbook workbook = new XSSFWorkbook();
        // 创建一个工作表
        Sheet sheet = workbook.createSheet("Languages");

        // 创建表头
        Row headerRow = sheet.createRow(0);
        headerRow.createCell(0).setCellValue("Key");
        headerRow.createCell(1).setCellValue("zh");
        headerRow.createCell(2).setCellValue("en");
        headerRow.createCell(3).setCellValue("vi");

        // 递归遍历 JSON 节点
        int rowIndex = 1;
        rowIndex = writeJsonNodeToSheet(sheet, rootNode, "", rowIndex);

        // 调整列宽
        sheet.autoSizeColumn(0);
        sheet.autoSizeColumn(1);
        sheet.autoSizeColumn(2);
        sheet.autoSizeColumn(3);

        // 将工作簿写入文件
        FileOutputStream outputStream = new FileOutputStream(excelFilePath);
        workbook.write(outputStream);
        workbook.close();
        outputStream.close();
        inputStream.close();
    }

    /**
     * 将 JSON 节点及其子节点写入 Excel 表格中
     *
     * @param sheet      Excel 工作表
     * @param node       当前 JSON 节点
     * @param parentKey  当前节点的父级 key
     * @param rowIndex   当前的行号
     * @return           最后的行号
     */
    public static int writeJsonNodeToSheet(Sheet sheet, JsonNode node, String parentKey, int rowIndex) {
        Iterator<String> fieldNames = node.fieldNames();

        while (fieldNames.hasNext()) {
            String fieldName = fieldNames.next();
            JsonNode valueNode = node.get(fieldName);
            String fullKey = parentKey.isEmpty() ? fieldName : parentKey + "." + fieldName;

            if (valueNode.isObject()) {
                // 如果节点是对象，递归调用
                rowIndex = writeJsonNodeToSheet(sheet, valueNode, fullKey, rowIndex);
            } else {
                // 否则，写入 key 和 value
                Row row = sheet.createRow(rowIndex++);
                row.createCell(0).setCellValue(fullKey);
                row.createCell(1).setCellValue(valueNode.asText());
                row.createCell(2).setCellValue("");
                row.createCell(3).setCellValue("");
            }
        }
        return rowIndex;
    }
}

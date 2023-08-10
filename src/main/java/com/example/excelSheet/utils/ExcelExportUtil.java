package com.example.excelSheet.utils;

import com.example.excelSheet.model.UserActionLog;
import com.example.excelSheet.model.UserActionLogDTO;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.tomcat.util.http.fileupload.ByteArrayOutputStream;
import org.springframework.util.ObjectUtils;

import java.io.IOException;
import java.lang.reflect.Field;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;


@Slf4j
public class ExcelExportUtil {
    public static ByteArrayOutputStream createExcel(List<UserActionLogDTO> logs) {
        XSSFWorkbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("User Action Logs");


        Font titleFont = workbook.createFont();
        titleFont.setFontName("Arial");
        titleFont.setBold(true);
        titleFont.setFontHeightInPoints((short)14);

        CellStyle titleCellStyle = workbook.createCellStyle();
        titleCellStyle.setAlignment(HorizontalAlignment.CENTER);
        titleCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        titleCellStyle.setFillForegroundColor(IndexedColors.SKY_BLUE.index);
        titleCellStyle.setBorderBottom(BorderStyle.MEDIUM);
        titleCellStyle.setBorderRight(BorderStyle.MEDIUM);
        titleCellStyle.setBorderTop(BorderStyle.MEDIUM);
        titleCellStyle.setFont(titleFont);
        titleCellStyle.setWrapText(true);

        Font dataFont = workbook.createFont();
        dataFont.setFontName("Arial");
        dataFont.setBold(false);
        dataFont.setFontHeightInPoints((short) 12);

        CellStyle dataStyle = workbook.createCellStyle();
        dataStyle.setWrapText(true);
        dataStyle.setAlignment(HorizontalAlignment.CENTER);
        dataStyle.setBorderBottom(BorderStyle.THIN);
        dataStyle.setBorderRight(BorderStyle.THIN);
        dataStyle.setBorderTop(BorderStyle.THIN);
        dataStyle.setFont(dataFont);



        insertFieldNameAsTitleToWorkbook(ExportConfig.customerExport.getCellExportConfigList(),sheet,titleCellStyle);
//        insertDataToWorkBook(workbook,ExportConfig.customerExport,logs,dataStyle);
        int rowIdx = 1;
        for (UserActionLogDTO log : logs) {
            sheet.autoSizeColumn(rowIdx);
            Row row = sheet.createRow(rowIdx);
            row.setRowStyle(dataStyle);
            row.createCell(0).setCellValue(log.getId());
            row.createCell(1).setCellValue(log.getAction());
            if (log.getCreated_timestamp() != null) {
                double timestampValue = log.getCreated_timestamp().doubleValue();
                Date date = new Date((long) timestampValue);
                SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
                String formattedDate = dateFormat.format(date);
                row.createCell(2).setCellValue(formattedDate);
            }
            row.createCell(3).setCellValue(log.getEmail());
            row.createCell(4).setCellValue(log.getMessage());
            row.createCell(5).setCellValue(log.getStatus());
            row.createCell(6).setCellValue(log.getType());
            rowIdx++;
        }

        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        try {
            workbook.write(outputStream);
            workbook.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return outputStream;
    }

    private static <T> void insertFieldNameAsTitleToWorkbook(List<CellConfig> cellConfigs, Sheet isheet, CellStyle titleCellStyle){
        int currentRow = isheet.getTopRow();
        Row row = isheet.createRow(currentRow);
        int i=0;
        isheet.autoSizeColumn(currentRow);
        for(CellConfig cellConfig:  cellConfigs){
            Cell currentCell = row.createCell(i);
            String fieldName = cellConfig.getFieldName();
            currentCell.setCellValue(fieldName);
            currentCell.setCellStyle(titleCellStyle);
            isheet.autoSizeColumn(i);
            i++;
        }
    }
    private static <T> void insertDataToWorkBook(Workbook workbook,ExportConfig exportConfig,List<UserActionLog> datas, CellStyle dataStyle) {
        int startRowIndex = exportConfig.getStartRow();
        int sheetIndex = exportConfig.getSheetIndex();
        Class clazz = exportConfig.getDataClass();
        List<CellConfig> cellConfigs = exportConfig.getCellExportConfigList();
        Sheet sheet = workbook.getSheetAt(sheetIndex);
        int currentRowIndex = startRowIndex;
        for (UserActionLog data : datas) {
            Row currentRow = sheet.createRow(currentRowIndex);
            if (ObjectUtils.isEmpty(currentRow)) {
                currentRow = sheet.createRow(currentRowIndex);
            }
            insertDataToCell(data, currentRow, cellConfigs, clazz, sheet, dataStyle);
            currentRowIndex++;
        }
    }

    private  static <T> void insertDataToCell(UserActionLog data,Row currentRow,List<CellConfig> cellConfigs,
                                              Class clazz , Sheet sheet, CellStyle dataStyle) {
        for (CellConfig cellConfig : cellConfigs) {
            Cell currentCell = currentRow.getCell(cellConfig.getColumnIndex());
            if (ObjectUtils.isEmpty(currentCell)) {
                currentCell = currentRow.createCell(cellConfig.getColumnIndex());
            }
            String cellValue = getCellValue(data, cellConfig, clazz);

            currentCell.setCellValue(cellValue);
            sheet.autoSizeColumn(cellConfig.getColumnIndex());
            currentCell.setCellStyle(dataStyle);
//            }

        }
    }
    private static <T> String getCellValue(T data, CellConfig cellConfig, Class clazz) {
        String fileName = cellConfig.getFieldName();

        try {
            Field field = getDeclaredField(clazz,fileName);
            if (ObjectUtils.isEmpty(field)){
                field.setAccessible(true);
                return !ObjectUtils.isEmpty(field.get(data))? field.get(data).toString() : "";
            }
            return "";
        }
        catch (Exception e) {
            log.info("" + e);
            return "";
        }
    }
    private static Field getDeclaredField(Class clazz, String fileName) {
        if (ObjectUtils.isEmpty(clazz)||ObjectUtils.isEmpty(fileName)){
            return null;
        }
        do{
            try{
                Field field = clazz.getDeclaredField(fileName);
                field.setAccessible(true);
                return field;
            }
            catch (Exception e){
                log.info(" "+e);
            }
        }while((clazz = clazz.getSuperclass()) != null);
        return  null;
    }

}

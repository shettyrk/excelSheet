package com.example.excelSheet.utils;

import com.example.excelSheet.model.UserActionLogDTO;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Component;
import org.springframework.util.ObjectUtils;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.lang.reflect.Field;
import java.util.List;
@Slf4j
@Component
public class ExcelUtils {
    public static ByteArrayInputStream exportUserActionLog(List<UserActionLogDTO> userdata) throws Exception {
        XSSFWorkbook xssfWorkbook = new XSSFWorkbook();

        XSSFSheet newSheet = xssfWorkbook.createSheet("Sheet1");
        newSheet.createFreezePane(4,2,4,2);
        XSSFFont titleFont = xssfWorkbook.createFont();
        titleFont.setFontName("Arial");
        titleFont.setBold(true);
        titleFont.setFontHeightInPoints((short)14);

        XSSFCellStyle titleCellStyle = xssfWorkbook.createCellStyle();
        titleCellStyle.setAlignment(HorizontalAlignment.CENTER);
        titleCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        titleCellStyle.setFillForegroundColor(IndexedColors.SKY_BLUE.index);
        titleCellStyle.setBorderBottom(BorderStyle.MEDIUM);
        titleCellStyle.setBorderRight(BorderStyle.MEDIUM);
        titleCellStyle.setBorderTop(BorderStyle.MEDIUM);
        titleCellStyle.setFont(titleFont);
        titleCellStyle.setWrapText(true);

        XSSFFont dataFont = xssfWorkbook.createFont();
        dataFont.setFontName("Arial");
        dataFont.setBold(false);
        dataFont.setFontHeightInPoints((short) 20);


        XSSFCellStyle dataStyle = xssfWorkbook.createCellStyle();
        dataStyle.setAlignment(HorizontalAlignment.CENTER);
        dataStyle.setBorderBottom(BorderStyle.THIN);
        dataStyle.setBorderRight(BorderStyle.THIN);
        dataStyle.setBorderTop(BorderStyle.THIN);
        dataStyle.setFont(dataFont);
        dataStyle.setWrapText(true);
        insertFieldNameAsTitleToWorkbook(ExportConfig.customerExport.getCellExportConfigList(),newSheet,titleCellStyle);

        insertDataToWorkBook(xssfWorkbook,ExportConfig.customerExport,userdata,dataStyle);

        ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
        xssfWorkbook.write(outputStream);
        outputStream.close();
        return new ByteArrayInputStream(outputStream.toByteArray());


    }
    private static <T> void insertDataToWorkBook(Workbook workbook,ExportConfig exportConfig,List<T> datas, XSSFCellStyle dataStyle){
        int startRowIndex = exportConfig.getStartRow();
        int sheetIndex = exportConfig.getSheetIndex();
        Class clazz= exportConfig.getDataClass();
        List<CellConfig> cellConfigs = exportConfig.getCellExportConfigList();
        Sheet sheet= workbook.getSheetAt(sheetIndex);
        int currentRowIndex = startRowIndex;
        for(T data : datas){
            Row currentRow = sheet.createRow(currentRowIndex);
            if (ObjectUtils.isEmpty(currentRow)){
                currentRow = sheet.createRow(currentRowIndex);
            }
            insertDataToCell(data,currentRow,cellConfigs,clazz,sheet,dataStyle);
            currentRowIndex++;
        }
    }
    private static <T> void insertFieldNameAsTitleToWorkbook(List<CellConfig> cellConfigs, Sheet sheet, XSSFCellStyle titleCellStyle){
        int currentRow = sheet.getTopRow();
        Row row = sheet.createRow(currentRow);
        int i=0;
        sheet.autoSizeColumn(currentRow);
        for(CellConfig cellConfig:  cellConfigs){
            Cell currentCell = row.createCell(i);
            String fieldName = cellConfig.getFieldName();
            currentCell.setCellValue(fieldName);
            currentCell.setCellStyle(titleCellStyle);
            sheet.autoSizeColumn(i);
            i++;
        }

    }

    private  static <T> void insertDataToCell(T data,Row currentRow,List<CellConfig> cellConfigs,
                                              Class clazz , Sheet sheet, XSSFCellStyle dataStyle){
        for(CellConfig cellConfig: cellConfigs){
            Cell currentCell = currentRow.getCell(cellConfig.getColumnIndex());
            if (ObjectUtils.isEmpty(currentCell)){
                currentCell = currentRow.createCell(cellConfig.getColumnIndex());
            }
            String cellValue = getCellValue(data,cellConfig,clazz);

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
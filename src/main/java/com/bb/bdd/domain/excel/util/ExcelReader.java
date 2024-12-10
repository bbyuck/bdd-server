package com.bb.bdd.domain.excel.util;

import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.core.io.Resource;
import org.springframework.core.io.ResourceLoader;
import org.springframework.stereotype.Service;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

@Slf4j
@Service
@RequiredArgsConstructor
public class ExcelReader {

    private final ResourceLoader resourceLoader;

    public XSSFWorkbook readXlsxOnClassPath(String resourcePath) {
        Resource resource = resourceLoader.getResource("classpath:" + resourcePath);
        try (InputStream inputStream = resource.getInputStream()){
            return new XSSFWorkbook(inputStream);
        }
        catch(IOException e) {
            log.error(e.getMessage(), e);
            throw new RuntimeException("xlsx 파일 생성중 에러가 발생했습니다.");
        }
    }

    public XSSFWorkbook readXlsxFile(File file) {
        try (FileInputStream fis = new FileInputStream(file)){
            return new XSSFWorkbook(fis);
        }
        catch(IOException e) {
            log.error(e.getMessage(), e);
            throw new RuntimeException("xlsx 파일 생성중 에러가 발생했습니다.");
        }
    }

    public HSSFWorkbook readXlsFile(File file) {
        try (FileInputStream fis = new FileInputStream(file)){
            return new HSSFWorkbook(fis);
        }
        catch(IOException e) {
            log.error(e.getMessage(), e);
            throw new RuntimeException("xls 파일 생성중 에러가 발생했습니다.");
        }
    }

    public String getCellValueAsString(Cell cell) {
        if (cell == null) {
            return "";
        }

        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue();
            case NUMERIC -> String.valueOf(cell.getNumericCellValue());
            case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
            case FORMULA -> cell.getCellFormula();
            default -> "";
        };
    }



    public void deleteRow(XSSFWorkbook xlsxWb, int deleteTargetRow) {
        // 밑에 있는 행을 위로 이동
        XSSFSheet sheet = xlsxWb.getSheetAt(0);

        // 병합 영역 처리
        for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
            CellRangeAddress mergedRegion = sheet.getMergedRegion(i);
            if (mergedRegion.getFirstRow() == deleteTargetRow || mergedRegion.getLastRow() == deleteTargetRow) {
                sheet.removeMergedRegion(i);
                i--;
            }
        }

        // 삭제 대상 행 제거
        XSSFRow rowToDelete = sheet.getRow(deleteTargetRow);
        if (rowToDelete != null) {
            sheet.removeRow(rowToDelete);
        }

        // 마지막 행 번호
        int lastRowNum = sheet.getLastRowNum();
    
        // 삭제행과 삭제행 바로 아래행의 스타일 충돌 해결 
        copyRowStyles(sheet, deleteTargetRow + 1, deleteTargetRow);

        /**
         * 행 삭제 
         */
        if (deleteTargetRow >= 0 && deleteTargetRow < lastRowNum) {
            sheet.shiftRows(deleteTargetRow + 1, lastRowNum, -1);
        }
        
        // 만약 마지막 행을 삭제하려는 경우
        if (deleteTargetRow == lastRowNum) {
            Row lastRow = sheet.getRow(deleteTargetRow);
            if (lastRow != null) {
                sheet.removeRow(lastRow);
            }
        }

        // 필터를 데이터 범위에 다시 설정
        CellRangeAddress newFilterRange = new CellRangeAddress(0, sheet.getLastRowNum(), 0, sheet.getRow(0).getLastCellNum() - 1);
        sheet.setAutoFilter(newFilterRange);
    }

    private void copyRowStyles(XSSFSheet sheet, int sourceRowIndex, int targetRowIndex) {
        XSSFRow sourceRow = sheet.getRow(sourceRowIndex);
        XSSFRow targetRow = sheet.getRow(targetRowIndex);
        if (sourceRow != null && targetRow != null) {
            for (int i = 0; i < sourceRow.getLastCellNum(); i++) {
                XSSFCell sourceCell = sourceRow.getCell(i);
                XSSFCell targetCell = targetRow.getCell(i);
                if (sourceCell != null) {
                    if (targetCell == null) {
                        targetCell = targetRow.createCell(i);
                    }
                    targetCell.setCellStyle(sourceCell.getCellStyle());
                }
            }
        }
    }
}

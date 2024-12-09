package com.bb.bdd.domain.excel.util;

import lombok.RequiredArgsConstructor;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
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
}

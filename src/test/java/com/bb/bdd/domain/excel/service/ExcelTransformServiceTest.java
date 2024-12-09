package com.bb.bdd.domain.excel.service;

import com.bb.bdd.domain.excel.ShopCode;
import org.assertj.core.api.Assertions;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.core.io.ClassPathResource;

import java.io.File;

@SpringBootTest
class ExcelTransformServiceTest {

    @Autowired
    private ExcelTransformService excelTransformService;


    @Test
    @DisplayName("[쿠팡] 주문 목록 엑셀 파일 -> CNP 업로드용 xls 파일 생성 테스트")
    public void createCoupangCnpXlsTest() throws Exception {
        // given
        ClassPathResource resource = new ClassPathResource("C:\\Users\\User\\Downloads\\de.xlsx");
        File inputFile = resource.getFile();

        // when
        File cnpXls = excelTransformService.createCnpXls(inputFile, ShopCode.COUPANG);

        // then
        Assertions.assertThat(cnpXls).exists();

        cnpXls.delete();
    }

    @Test
    @DisplayName("[쿠팡] 주문량 집계 엑셀 생성")
    public void createCoupangCountXlsxTest() throws Exception {
        // given
        ClassPathResource resource = new ClassPathResource("sample/coupang_sample.xlsx");
        File inputFile = resource.getFile();

        // when
        File countXlsx = excelTransformService.createCountXlsx(inputFile, ShopCode.COUPANG);

        // then
        Assertions.assertThat(countXlsx).exists();

        countXlsx.delete();
    }

    @Test
    @DisplayName("[네이버] 주문 목록 엑셀 파일 -> CNP 업로드용 xls 파일 생성 테스트")
    public void createNaverCnpXlsTest() throws Exception {
        // given
        ClassPathResource resource = new ClassPathResource("sample/naver_sample.xlsx");
        File inputFile = resource.getFile();

        // when
        File cnpXls = excelTransformService.createCnpXls(inputFile, ShopCode.NAVER);

        // then
        Assertions.assertThat(cnpXls).exists();

        cnpXls.delete();
    }

    @Test
    @DisplayName("[네이버] 주문량 집계 엑셀 생성")
    public void createNaverCountXlsxTest() throws Exception {
        // given
        ClassPathResource resource = new ClassPathResource("sample/naver_sample.xlsx");
        File inputFile = resource.getFile();

        // when
        File countXlsx = excelTransformService.createCountXlsx(inputFile, ShopCode.NAVER);

        // then
        Assertions.assertThat(countXlsx).exists();

        countXlsx.delete();
    }


}
package com.bb.bdd.domain.excel.controller;

import com.bb.bdd.domain.excel.DeliveryCompanyCode;
import com.bb.bdd.domain.excel.ShopCode;
import com.bb.bdd.domain.excel.service.ExcelTransformService;
import lombok.RequiredArgsConstructor;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

@RestController
@RequiredArgsConstructor
@RequestMapping("excel")
public class ExcelTransformController {


    private final ExcelTransformService excelTransformService;

    /**
     * 쿠팡 -> CNP + count
     * 네이버 -> CNP + count
     *
     * 쿠팡 + 운송장 번호 -> 운송장 입력된 쿠팡
     * 네이버 + 운송장 번호 -> 운송장 입력된 네이버
     */

    @PostMapping("coupang/cnp/transform")
    public void transfromCoupangToCnp(@RequestParam("order") MultipartFile orderExcelMultiFile) {
        excelTransformService.transformToCnp(orderExcelMultiFile, ShopCode.COUPANG);
    }

    @PostMapping("naver/cnp/transform")
    public void transformNaverToCnp(@RequestParam("order") MultipartFile orderExcelMultiFile) {
        excelTransformService.transformToCnp(orderExcelMultiFile, ShopCode.NAVER);
    }


    @PostMapping("coupang/tracking-number/enter")
    public void enterTrackingNumberCoupang(@RequestParam("order") MultipartFile orderExcelMultipartFile, @RequestParam("trackingNumber") MultipartFile trackingNumberMultipartFile) {
        excelTransformService.enterTrackingNumber(orderExcelMultipartFile, trackingNumberMultipartFile, ShopCode.COUPANG, DeliveryCompanyCode.CJ);
    }

    @PostMapping("naver/tracking-number/enter")
    public void enterTrackingNumberNaver(@RequestParam("order") MultipartFile orderExcelMultipartFile, @RequestParam("trackingNumber") MultipartFile trackingNumberMultipartFile) {
        excelTransformService.enterTrackingNumber(orderExcelMultipartFile, trackingNumberMultipartFile, ShopCode.NAVER, DeliveryCompanyCode.CJ);
    }
}

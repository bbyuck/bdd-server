package com.bb.bdd.domain.excel;

import lombok.Getter;
import lombok.RequiredArgsConstructor;

@Getter
@RequiredArgsConstructor
public enum DownloadCode {

    COUPANG_CNP("쿠팡 CNP 변환"),
    NAVER_CNP("네이버 CNP 변환"),
    COUPANG_TRACKING_NUMBER_INPUT("쿠팡 운송장 입력"),
    NAVER_TRACKING_NUMBER_INPUT("네이버 운송장 입력");

    private final String filename;

}

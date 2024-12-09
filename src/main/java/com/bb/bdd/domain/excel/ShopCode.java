package com.bb.bdd.domain.excel;

import lombok.Getter;
import lombok.RequiredArgsConstructor;

@Getter
@RequiredArgsConstructor
public enum ShopCode {
    COUPANG("쿠팡", 1, 0, DownloadCode.COUPANG_CNP, DownloadCode.COUPANG_TRACKING_NUMBER_INPUT, "수취인이름", "수취인전화번호", "운송장번호"),
    NAVER("네이버", 1, 1, DownloadCode.NAVER_CNP, DownloadCode.NAVER_TRACKING_NUMBER_INPUT, "수취인명", "수취인연락처1", "송장번호");

    private final String value;
    private final int keyColumnIndex;
    private final int headerRowIndex;

    private final DownloadCode cnpCode;
    private final DownloadCode trackingNumberCode;

    private final String receiverNameColumnName;
    private final String receiverPhoneColumnName;
    private final String trackingNumberColumnName;
}

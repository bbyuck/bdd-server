package com.bb.bdd.domain.excel;

import lombok.Getter;
import lombok.RequiredArgsConstructor;

@Getter
@RequiredArgsConstructor
public enum DeliveryCompanyCode {
    CJ("CJ대한통운", 0,"운송장번호", "받는분", "받는분전화번호");

    private final String value;
    private final int headerRowIndex;
    private final String trackingNumberColumnName;
    private final String receiverNameColumnName;
    private final String receiverPhoneColumnName;
}

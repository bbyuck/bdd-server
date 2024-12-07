package com.bb.bdd.domain.excel;

import lombok.Getter;
import lombok.RequiredArgsConstructor;

@Getter
@RequiredArgsConstructor
public enum ShopCode {
    COUPANG("쿠팡"), NAVER("네이버");

    private final String value;
}

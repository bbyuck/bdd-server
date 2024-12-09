package com.bb.bdd.domain.excel.dto;

import lombok.Getter;
import lombok.Setter;
import lombok.ToString;

@Getter
@Setter
@ToString
public class NaverColumnDto {
	// 상품주문번호
	private String productOrderNum;
	// 주문번호
	private String orderNum;
	// 배송방법
	private String shippingMethod;
	// 택배사
	private String courier;
	// 송장번호
	private String waybillNum;
	// 발송일
	private String shippingDate;
	// 수취인명
	private String receiverName;
	// 상품명
	private String productName;
	// 옵션정보
	private String optionInfo;
	// 수량
	private Integer quantity;
	// 배송비 형태
	private String shippingCostForm;
	// 수취인연락처1
	private String receiverPhone1;
	// 배송지
	private String receiverAddress;
    // 상세배송지
    private String receiverAddressDetail;
	// 배송메세지
	private String deliveryMessage;
	// 출고지
	private String shipFrom;
	// 결제수단
	private String methodOfPayment;
	// 수수료 과금구분
	private String feeChargingCategory;
	// 수수료결제방식
	private String feePaymentMethod;
	// 결제수수료
	private Integer paymentFee;
	// 매출연동 수수료
	private Integer salesLinkageFee;
	// 정산예정금액
	private Integer estimatedTotalAmount;
	// 유입경로
	private String channel;
	// 구매자 주민등록번호
	private String buyerSSN;
	// 개인통관고유부호
	private String pccc;
	// 주문일시
	private String orderDateTime;
	// 1년 주문건수
	private Integer numOfOrdersPerYear;
	// 구매자ID
	private String buyerId;
	// 구매자명
	private String buyerName;
	// 결제일
	private String paymentDate;
	// 상품종류
	private String productType;
	// 주문세부상태
	private String orderDetailStatus;
	// 주문상태
	private String orderStatus;
	// 상품번호
	private String itemNum;
	// 배송속성
	private String deliveryProperty;
	// 배송희망일
	private String wantDeliveryDate;
	// (수취인연락처1)
	private String _receiverPhone1;
	// (수취인연락처2)
	private String _receiverPhone2;
	// (우편번호)
	private String zipcode;
	// (기본주소)
	private String receiverAddress1;
	// (상세주소)
	private String receiverAddress2;
	// (구매자연락처)
	private String buyerPhone;
}

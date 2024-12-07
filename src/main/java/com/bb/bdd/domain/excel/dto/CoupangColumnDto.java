package com.bb.bdd.domain.excel.dto;

import lombok.Getter;
import lombok.Setter;
import lombok.ToString;

@Getter
@Setter
@ToString
public class CoupangColumnDto {
	// 번호
	private Integer num;
	// 묶음배송번호
	private String shippingNum;
	// 주문번호
	private String orderNum;
	// 택배사
	private String courier;
	// 운송장번호
	private String waybillNum;
	// 분리배송 Y/N
	/* 
	 * N
	 * Y
	 * 분리배송 불가
	 */
	private String separateDelivery;
	// 분리배송 출고예정일
	private String separateExpectedDeliveryDate;
	// 주문시 출고예정일
	private String expectedDeliveryDate;
	// 출고일(발송일)
	private String deliveryDate;
	// 주문일
	private String orderDate;
	// 등록상품명
	private String productName;
	// 등록옵션명
	private String optionName;
	// 노출상품명
	private String displayedProductName;
	// 노출상품ID
	private String displayedProductId;
	// 옵션ID
	private String optionId;
	// 최초등록옵션명
	private String firstOptionName;
	// 업체상품코드
	private String productCode;
	// 바코드
	private String barcode;
	// 결제액
	private Integer payment;
	// 배송비구분
	private String deliveryFeeFlag;
	// 배송비
	private Integer deliveryFee;
	// 도서산간 추가배송비
	private Integer additionalDeliveryFee;
	// 구매수(수량)
	private Integer quantity;
	// 옵션판매가(판매단가)
	private Integer unitPrice;
	// 구매자
	private String customerName;
	// 구매자이메일
	private String customerEmail;
	// 구매자전화번호
	private String customerPhone;
	// 수취인이름
	private String receiverName;
	// 수취인전화번호
	private String receiverPhone;
	// 우편번호
	private String postNum;
	// 수취인 주소
	private String receiverAddress;
	// 배송메세지
	private String deliveryMessage;
	// 상품별 추가메시지
	private String additionalMessagePerItem;
	// 주문자 추가메시지
	private String ordererAdditionalMessage;
	// 배송완료일
	private String deliveryCompleteDate;
	// 구매확정일자
	private String confirmationPurchaseDate;
	// 개인통관번호(PCCC)
	private String pccc;
	// 통관용구매자전화번호
	private String buyerPhoneNumForCustomsClearance;
	// 기타
	private String etc;
	// 결제위치
	private String paymentLocation;
	
}

package com.bb.bdd.domain.excel.dto;

import lombok.Getter;
import lombok.Setter;
import lombok.ToString;

@Getter
@Setter
@ToString
public class CnpInputDto {
	// 수취인이름
	private String receiverName;
	// 수취인전화번호
	private String receiverPhone;
	// 수취인 주소
	private String receiverAddress;
	// 노출상품명(옵션명)
	private String orderContents;
	// 배송메세지
	private String remark;

}


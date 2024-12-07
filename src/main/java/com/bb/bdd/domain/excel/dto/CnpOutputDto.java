package com.bb.bdd.domain.excel.dto;

import lombok.Getter;
import lombok.Setter;
import lombok.ToString;

@Getter
@Setter
@ToString
public class CnpOutputDto {
	// 운송장번호
	private String waybillNum;
	// 받는분
	private String receiverName;
	//받는분 전화번호
	private String receiverPhone;
	
}

package com.bb.bdd.domain.excel.dto;

import lombok.AllArgsConstructor;
import lombok.Getter;

@Getter
@AllArgsConstructor
public class Pair implements Comparable<Pair>{
	private String productName;
	private int count;
	
	@Override
	public int compareTo(Pair opponent) {
		return productName.compareTo(opponent.productName);
	}
}
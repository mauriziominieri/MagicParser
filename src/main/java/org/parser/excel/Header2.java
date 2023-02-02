package org.parser.excel;

import lombok.Data;

import java.util.List;

@Data
public class Header2 {
	private final String fieldName;
	private final String columnName;
	private final int columnIndex; // TODO vedere se serve
	private final String picchettoGroup;
	private final List<Coordinata> coordinataList;
}
package org.parser.excel;

import lombok.Data;

import java.util.List;

@Data
public class Header {
	private final String fieldName;
	private final String columnName;
	private final int columnIndex; // TODO vedere se serve (IMPORT EXCEL)
	private final String group;
	private final List<Coordinate> coordinateList;
}
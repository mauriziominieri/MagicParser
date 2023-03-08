package org.parser.dto;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.parser.excel.Field;

@NoArgsConstructor
@AllArgsConstructor
@Data
public class ContrattiAppaltatoriDto {

	private Long idContrattiAppaltatori;

	@Field(column = "descrFornitore")
	private String descrFornitore;

	@Field(column = "appaltatore")
	private String fornitore;

	private Long idCantiere;

	private String numOda;

	private Boolean mainContractor;

}

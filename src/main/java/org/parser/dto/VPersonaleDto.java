package org.parser.dto;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.parser.excel.Field;

@NoArgsConstructor
@AllArgsConstructor
@Data
public class VPersonaleDto {

	@Field(column = "CF")
	private String cf;

	@Field(column = "COGNOME")
	private String cognome;

	private Long idAppaltatore;

	private Long idMansione;

	private Long idPersonale;

	@Field(column = "IDONEITà")
	private Boolean idoneita;

	@Field(column = "MANSIONE")
	private String mansione;

	@Field(column = "NOME")
	private String nome;
	
	private Long nrOreManodopera;

}

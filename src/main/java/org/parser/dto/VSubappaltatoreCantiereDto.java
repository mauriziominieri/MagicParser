package org.parser.dto;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.parser.excel.Field;

@NoArgsConstructor
@AllArgsConstructor
@Data
public class VSubappaltatoreCantiereDto {
	private String codeCantiere;

	private String codiceIdentificativo;

	@Field(column = "descrizione")
	private String denominazioneImpresa;

	private Long idCantiere;

	@Field(column = "subappaltatore")
	private Long idSubappaltatore;

	private String numOda;
}

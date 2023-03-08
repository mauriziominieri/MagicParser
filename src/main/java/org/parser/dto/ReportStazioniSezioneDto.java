package org.parser.dto;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.parser.excel.Field;

/**
 * Created by IntelliJ IDEA.
 *
 * @author: Maurizio Minieri
 * @email: mauminieri@gmail.com
 * @website: www.mauriziominieri.it
 */

@NoArgsConstructor
@AllArgsConstructor
@Data
public class ReportStazioniSezioneDto {
    @Field(column = "Fondazione apparecchiature periodo")
    private String fondazioneApparecchiaturePeriodo;
    @Field(column = "Fondazione apparecchiature totale")
    private String fondazioneApparecchiatureTotale;
    @Field(column = "Montaggio carpenteria periodo")
    private String montaggioCarpenteriaPeriodo;
    @Field(column = "Montaggio carpenteria totale")
    private String montaggioCarpenteriaTotale;
    @Field(column = "Montaggio isolatori totale")
    private String montaggioIsolatoriTotale;
    @Field(column = "Montaggio isolatori periodo")
    private String montaggioIsolatoriPeriodo;
    @Field(column = "Montaggio conduttori rigidi totale")
    private String montaggioRigidiTotale;
    @Field(column = "Montaggio conduttori rigidi periodo")
    private String montaggioRigidiPeriodo;
    @Field(column = "Montaggio apparecchiature totale")
    private String montaggioApparecchiatureTotale;
    @Field(column = "Montaggio apparecchiature periodo")
    private String montaggioApparecchiaturePeriodo;
    @Field(column = "Montaggio conduttori flessibili/in corda periodo")
    private String montaggioFlessibiliPeriodo;
    @Field(column = "Montaggio conduttori flessibili/in corda totale")
    private String montaggioFlessibiliTotale;
}

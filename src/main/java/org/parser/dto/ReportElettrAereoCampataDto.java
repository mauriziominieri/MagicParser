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
public class ReportElettrAereoCampataDto {
    @Field(column = "autorizzazione")
    private String autorizzazioneCampata;

    @Field(column = "asservimento")
    private String asservimentoCampata;

    @Field(column = "stendimento")
    private String stendimento;

    @Field(column = "regolazione")
    private String regolazione;

    @Field(column = "Morsetteria e accessori")
    private String morsetteriaAccessori;
}

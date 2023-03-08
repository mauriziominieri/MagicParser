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
public class ReportElettrAereoPicchettoDto {
    @Field(column = "Autorizzazione", group = "PICCHETTO")
    private String autorizzazione;

    @Field(column = "ASSERVIMENTO", group = "PICCHETTO")
    private String asservimento;

    @Field(column = "ARCHEOLOGIA", group = "PICCHETTO")
    private String archeologia;

    @Field(column = "BOB", group = "PICCHETTO")
    private String bob;

    @Field(column = "procedibilità autorizzativa", group = "PICCHETTO")
    private String procedibilitaAutorizzativa;

    @Field(column = "scavi-perforazioni %", group = "PICCHETTO")
    private String scaviPerforazioni;

    @Field(column = "fondazione %", group = "PICCHETTO")
    private String fondazione;

    @Field(column = "sotto-fondazione %", group = "PICCHETTO")
    private String sottoFondazione;

    @Field(column = "montaggio %", group = "PICCHETTO")
    private String montaggio;

    @Field(column = "rinterro %", group = "PICCHETTO")
    private String rinterro;

    @Field(column = "ripristino %", group = "PICCHETTO")
    private String ripristino;
}

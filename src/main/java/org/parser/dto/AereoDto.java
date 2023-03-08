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
public class AereoDto {

    @Field(column = "Autorizzazione", group = "PICCHETTO", index = 1)
    private String autorizzazione;

    @Field(column = "Asservimento", group = "PICCHETTO", index = 1)
    private String asservimento;

    @Field(column = "Archeologia", group = "PICCHETTO")
    private String archeologia;

    @Field(column = "BOB", group = "PICCHETTO")
    private String bob;

    @Field(column = "Procedibilità autorizzativa", group = "PICCHETTO")
    private String procedibilitaAutorizzativa;

    @Field(column = "Scavi-perforazioni %", group = "PICCHETTO")
    private String scaviPerforazioni;

    @Field(column = "Fondazione %", group = "PICCHETTO")
    private String fondazione;

    @Field(column = "Sotto-fondazione %", group = "PICCHETTO")
    private String sottoFondazione;

    @Field(column = "Montaggio %", group = "PICCHETTO")
    private String montaggio;

    @Field(column = "Rinterro %", group = "PICCHETTO")
    private String rinterro;

    @Field(column = "Ripristino %", group = "PICCHETTO")
    private String ripristino;

    @Field(column = "Autorizzazione", index = 2)
    private String autorizzazioneCampata;

    @Field(column = "Asservimento", index = 2)
    private String asservimentoCampata;

    @Field(column = "Stendimento")
    private String stendimento;

    @Field(column = "Regolazione")
    private String regolazione;

    @Field(column = "Morsetteria e accessori")
    private String morsetteriaAccessori;

    @Field(column = "Portale Esistente")
    private String portaleEsistente;
}

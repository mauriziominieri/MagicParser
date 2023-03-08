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
public class ReportElettrAereoRiepilogoDto {

    @Field(column = "autorizzazione da completare")
    private String autorizzazioneDaCompletare;

    @Field(column = "autorizzazione completati")
    private String autorizzazioneCompletati;

    @Field(column = "autorizzazione avanzamento")
    private String autorizzazioneAvanzamento;

    @Field(column = "asservimento da completare")
    private String asservimentoDaCompletare;

    @Field(column = "asservimento completati")
    private String asservimentoCompletati;

    @Field(column = "asservimento avanzamento")
    private String asservimentoAvanzamento;

    @Field(column = "procedbilità autorizzativa da completare")
    private String procedibilitaDaCompletare;

    @Field(column = "procedbilità autorizzativa completati")
    private String procedibilitaCompletati;

    @Field(column = "procedbilità autorizzativa avanzamento")
    private String procedibilitaAvanzamento;

    @Field(column = "ripristino da completare")
    private String ripristinoDaCompletare;

    @Field(column = "ripristino completati")
    private String ripristinoCompletati;

    @Field(column = "ripristino avanzamento")
    private String ripristinoAvanzamento;

    @Field(column = "asservimento campata da completare")
    private String asservimentoCampataDaCompletare;

    @Field(column = "asservimento campata completati")
    private String asservimentoCampataCompletati;

    @Field(column = "asservimento campata avanzamento")
    private String asservimentoCampataAvanzamento;
}

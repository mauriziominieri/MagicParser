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
public class ReportStazioniDto {

    @Field(column = "avanzamento da")
    private String dataDa;

    @Field(column = "avanzamento a")
    private String dataA;

    @Field(column = "codice cantiere")
    private String codiceCantiere;

    @Field(column = "acquisizione terreni periodo")
    private String atpPeriodo;

    @Field(column = "acquisizione terreni totale")
    private String atpTotale;

    @Field(column = "bob periodo")
    private String bobPeriodo;

    @Field(column = "bob totale")
    private String bobTotale;

    @Field(column = "sotto fondazioni periodo")
    private String sottoFondazioniPeriodo;

    @Field(column = "sotto fondazioni totale")
    private String sottoFondazioniTotale;

    @Field(column = "trasporto periodo")
    private String trasportoPeriodo;

    @Field(column = "trasporto totale")
    private String trasportoTotale;

    @Field(column = "appalto lavori presunto")
    private String appaltoLavoriPresunto;

    @Field(column = "appalto lavori consuntivato")
    private String appaltoLavoriConsuntivato;

    @Field(column = "appalto lavori %")
    private String appaltoLavoriP;

    @Field(column = "note")
    private String note;

    @Field(column = "Oggetto appaltato")
    private String oggettoAppaltato;

    @Field(column = "Nome fornitore")
    private String nomeFornitore;

    @Field(column = "N Oda/Lda")
    private String nOda;

    @Field(column = "Importo", group = "APPALTATORE/FORNITORE")
    private String importoAppaltatore;

    @Field(column = "Regione Sociale")
    private String regioneSociale;

    @Field(column = "Importo", group = "SUBAPPALTATORE")
    private String importoSubappaltatore;

    @Field(column = "materiali presunto")
    private String materialiPresunto;

    @Field(column = "materiali consuntivato")
    private String materialiConsuntivato;

    @Field(column = "materiali %")
    private String materialiP;

    @Field(column = "complessivo opera presunto")
    private String complessivoOperaPresunto;

    @Field(column = "complessivo opera consuntivato")
    private String complessivoOperaConsuntivato;

    @Field(column = "complessivo opera %")
    private String complessivoOperaP;
}

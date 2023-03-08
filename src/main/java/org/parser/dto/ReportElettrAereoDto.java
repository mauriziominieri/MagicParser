package org.parser.dto;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.parser.excel.Field;

import java.util.Date;

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
public class ReportElettrAereoDto {

    @Field(column = "avanzamento da")
    private String avanzamentoDa;

    @Field(column = "avanzamento a")
    private String avanzamentoA;

    @Field(column = "wbs")
    private String wbs;

    @Field(column = "consistenza intervento")
    private String consistenzaIntervento;

    @Field(column = "oda/lda")
    private String odaLda;

    @Field(column = "main contractor")
    private String mainContractor;

    @Field(column = "nome cantiere")
    private String nomeCantiere;

    @Field(column = "dt di riferimento")
    private String dtRiferimento;

    @Field(column = "committente")
    private String committente;

    @Field(column = "rpe")
    private String rpe;

    @Field(column = "direttore lavori")
    private String dl;

    @Field(column = "csp")
    private String csp;

    @Field(column = "cse")
    private String cse;

    @Field(column = "progettista")
    private String progettista;

    @Field(column = "collaudatore")
    private String collaudatore;

    @Field(column = "iac")
    private String iac;

    @Field(column = "iaca")
    private String iaca;

    @Field(column = "rss")
    private String rss;

    @Field(column = "sorvegliante di cantiere")
    private String sorvegliante;

    @Field(column = "rr")
    private String rr;

    @Field(column = "data inizio cantiere")
    private Date dataInizioCantiere;

    @Field(column = "data presunta inizio cantiere")
    private Date dataPresuntaInizioCantiere;

    @Field(column = "data presunta fine cantiere")
    private Date dataPresuntaFineCantiere;

    @Field(column = "data fine cantiere")
    private Date dataFineCantiere;

    @Field(column = "regione")
    private String regione;

    @Field(column = "comune")
    private String comune;

    @Field(column = "indirizzo/localita")
    private String indirizzoLoc;

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

    @Field(column = "Portale esistente")
    private String portaleEsistente;

    @Field(column = "temperatura min °C")
    private String temperaturaMin;

    @Field(column = "temperatura max °C")
    private String temperaturaMax;

    @Field(column = "mattino min")
    private String mattinoMin;

    @Field(column = "mattino max")
    private String mattinoMax;

    @Field(column = "pomeriggio min")
    private String pomeriggioMin;

    @Field(column = "pomeriggio max")
    private String pomeriggioMax;

    @Field(column = "gdlN")
    private String gdlN;
}

package org.parser.excel;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

import java.util.List;

/**
 * Created by IntelliJ IDEA.
 *
 * @author: Maurizio Minieri
 * @email: mauminieri@gmail.com
 * @website: www.mauriziominieri.it
 */

@NoArgsConstructor
@AllArgsConstructor
@Getter
@Setter
public class XLSXCantiereTemplateSetting<T> extends XLSXTemplateSetting {
    private List<T> objectList; // lista degli oggetti poi da scrivere nell'excel
    private int objectListPortion; // n oggetti possono essere divisi in porzioni
    private Direction direction; // fissata la cella header decide in che direzione scrivere il valore
    private int step; // ogni quante celle riempire in quella direzione (utile per un layout "a scacchi")
    private String headerGroup; // potrebbe essere utile differenziare alcune celle dell'header in un gruppo
    private int headerGroupStartIndex; // per dare un inizio diverso a quel gruppo

    public enum Direction {
        UP,
        RIGHT,
        BOTTOM,
        LEFT
    }
}

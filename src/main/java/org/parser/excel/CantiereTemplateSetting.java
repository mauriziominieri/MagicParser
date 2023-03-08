package org.parser.excel;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

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
@Data
public class CantiereTemplateSetting<T> extends TemplateSetting {
    private Class objectClass; // classe dell'oggetto che ha la lista
    private List<T> objectList; // lista degli oggetti poi da scrivere nell'excel
    private String key; // la stringa da cercare nell'excel che rappresenta la porzione poi da duplicare e riempire
    private Direction direction; // fissata la cella header decide in che direzione scrivere il valore
    private int step; // ogni quante celle riempire in quella direzione (utile per un layout "a scacchi")
}

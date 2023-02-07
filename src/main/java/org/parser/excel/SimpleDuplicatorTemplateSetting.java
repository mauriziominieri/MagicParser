package org.parser.excel;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

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
public class SimpleDuplicatorTemplateSetting<T> extends SimpleTemplateSetting {
    private String key; // la stringa da cercare nell'excel che rappresenta la porzione poi da duplicare e riempire, DEVE ESSERE UN ATTRIBUTO DELL'OGGETTO
    private int numberOfRowsToDuplicate; // il numero di righe da copiare partendo dalla key
}

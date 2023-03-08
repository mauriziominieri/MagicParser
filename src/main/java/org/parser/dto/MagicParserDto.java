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
public class MagicParserDto {

    @Field(column = "mela")
    private String mela;

    @Field(column = "pera")
    private String pera;

    @Field(column = "kiwi")
    private String kiwi;

    @Field(column = "arancia")
    private String arancia;

    @Field(column = "melanzana")
    private String melanzana;

    @Field(column = "banana")
    private String banana;

    @Field(column = "lampone")
    private String lampone;

    private String castagna;

    private String lime;
}

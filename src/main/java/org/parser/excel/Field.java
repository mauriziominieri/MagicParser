package org.parser.excel;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface Field {

	String[] column() default "";

	String[] group() default ""; // gruppi per gestire i picchetti dell' ElettAereo

	boolean print() default true;//Scrivere su pdf

	boolean read() default true;//Leggere excel

	boolean printXlsx() default true; //Scrivere su excel

	int row() default 1;

	int colspan() default 1;

	String columnSec() default "";

	int columnType() default 0;

}
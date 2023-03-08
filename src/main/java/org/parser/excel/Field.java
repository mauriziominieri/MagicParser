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

	int index() default -1; // per gestire quei campi duplicati

}
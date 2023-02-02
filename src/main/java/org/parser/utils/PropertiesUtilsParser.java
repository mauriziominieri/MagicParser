package org.parser.utils;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.MessageSource;
import org.springframework.stereotype.Component;

import javax.annotation.PostConstruct;

@Component
public class PropertiesUtilsParser {

	@Autowired
	SpringContextParser springContext;

	private static MessageSource messageSource;

	@PostConstruct
	private void init() {
		messageSource = SpringContextParser.getBean(MessageSource.class);
	}

	public static String getMessage(String pattern) {
		return messageSource.getMessage(pattern, null, null);
	}

	public static String getMessage(String pattern, Object[] args) {
		return messageSource.getMessage(pattern, args, null);
	}
}

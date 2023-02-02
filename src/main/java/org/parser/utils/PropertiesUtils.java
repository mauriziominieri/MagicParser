package org.parser.utils;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.MessageSource;
import org.springframework.stereotype.Component;

import javax.annotation.PostConstruct;

@Component
public class PropertiesUtils {

	@Autowired
	SpringContext springContext;

	private static MessageSource messageSource;

	@PostConstruct
	private void init() {
		messageSource = SpringContext.getBean(MessageSource.class);
	}

	public static String getMessage(String pattern) {
		return messageSource.getMessage(pattern, null, null);
	}

	public static String getMessage(String pattern, Object[] args) {
		return messageSource.getMessage(pattern, args, null);
	}
}

package org.parser.utils;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.context.MessageSource;
import org.springframework.stereotype.Component;

/**
 * Created by IntelliJ IDEA.
 *
 * @author: Maurizio Minieri
 * @email: mauminieri@gmail.com
 * @website: www.mauriziominieri.it
 */

@Component
public class PropertiesUtils {

    private static MessageSource messageSource;

    @Autowired
    public void setMessageSource(MessageSource messageSource) {
        PropertiesUtils.messageSource = messageSource;
    }

    public static String getMessage(String pattern, Object[] args) {
        return messageSource.getMessage(pattern, args, null);
    }
}

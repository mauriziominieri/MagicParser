package org.parser.exception;

import io.swagger.annotations.ApiModel;
import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

import java.util.Date;

@ApiModel
@NoArgsConstructor
@AllArgsConstructor
@Getter
@Setter
public class ErrorMessage {

	private Date timestamp;
	private String message;
	private String exception;
	
}

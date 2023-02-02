package org.parser.exception;

import lombok.extern.log4j.Log4j2;
import org.hibernate.exception.ConstraintViolationException;
import org.springframework.dao.DataIntegrityViolationException;
import org.springframework.dao.InvalidDataAccessResourceUsageException;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.orm.jpa.JpaSystemException;
import org.springframework.transaction.TransactionSystemException;
import org.springframework.web.bind.annotation.ControllerAdvice;
import org.springframework.web.bind.annotation.ExceptionHandler;

import java.util.Date;

@ControllerAdvice
@Log4j2
public class ExceptionHelper {

	@ExceptionHandler(value = { Exception.class })
	public ResponseEntity<Object> handleException(Exception ex) {
		log.error(ex.getMessage(), ex);
		
		String errorMessageDescription = ex.getLocalizedMessage() == null ? ex.toString() : ex.getLocalizedMessage();
		ErrorMessage errorMessage = new ErrorMessage(new Date(), errorMessageDescription, ex.getClass().toString());

		return new ResponseEntity<Object>(errorMessage, new HttpHeaders(), HttpStatus.INTERNAL_SERVER_ERROR);
	}

	@ExceptionHandler(value = { DataIntegrityViolationException.class })
	public ResponseEntity<Object> handleDataIntegrityViolationException(DataIntegrityViolationException ex) {
		log.error(ex.getMessage(), ex);

		String errorMessageDescription = ex.getLocalizedMessage() == null ? ex.toString() : ex.getLocalizedMessage();
		ErrorMessage errorMessage = new ErrorMessage(new Date(), errorMessageDescription, ex.getClass().toString());
		String sqlError = ((ConstraintViolationException) ex.getCause()).getSQLException().getMessage();
		errorMessage.setMessage(errorMessage.getMessage() + ". " + sqlError);

		return new ResponseEntity<Object>(errorMessage, new HttpHeaders(), HttpStatus.INTERNAL_SERVER_ERROR);
	}

	@ExceptionHandler(value = { TransactionSystemException.class, JpaSystemException.class })
	public ResponseEntity<Object> handleConstraintViolationException(Exception ex) {
		log.error(ex.getMessage(), ex);

		String errorMessageDescription = ex.getLocalizedMessage() == null ? ex.toString() : ex.getLocalizedMessage();
		ErrorMessage errorMessage = new ErrorMessage(new Date(), errorMessageDescription, ex.getClass().toString());
		String constraintViolationError = ex.getCause().getCause().getMessage();
		errorMessage.setMessage(errorMessage.getMessage() + ". " + constraintViolationError);

		return new ResponseEntity<Object>(errorMessage, new HttpHeaders(), HttpStatus.INTERNAL_SERVER_ERROR);
	}

	@ExceptionHandler(value = { InvalidDataAccessResourceUsageException.class })
	public ResponseEntity<Object> handleInvalidDataAccessResourceUsageException(InvalidDataAccessResourceUsageException ex) {
		log.error(ex.getMessage(), ex);

		String errorMessageDescription = ex.getLocalizedMessage() == null ? ex.toString() : ex.getLocalizedMessage();
		ErrorMessage errorMessage = new ErrorMessage(new Date(), errorMessageDescription, ex.getClass().toString());
		String sqlError = ex.getMostSpecificCause().toString();
		errorMessage.setMessage(errorMessage.getMessage() + ". " + sqlError);

		return new ResponseEntity<Object>(errorMessage, new HttpHeaders(), HttpStatus.INTERNAL_SERVER_ERROR);
	}
}

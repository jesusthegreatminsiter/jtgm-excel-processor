package org.jtgm.presentation.exception;

import org.springframework.http.HttpStatus;
import org.springframework.web.bind.annotation.ResponseStatus;

@ResponseStatus(HttpStatus.INTERNAL_SERVER_ERROR)
public class RestControllerException extends RuntimeException{

    public RestControllerException(String mes){
        super(mes);
    }
}


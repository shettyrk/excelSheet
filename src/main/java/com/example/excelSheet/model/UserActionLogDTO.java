package com.example.excelSheet.model;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

import java.math.BigInteger;
@Getter
@Setter
@AllArgsConstructor
@NoArgsConstructor

public class UserActionLogDTO {
    private String id;
    private String email;
    private String type;
    private String action;
    private BigInteger created_timestamp;
    private String message;
    private String status;
}

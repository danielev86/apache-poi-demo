package com.danielev86.apachepoidemo.bean;

import lombok.Data;
import lombok.EqualsAndHashCode;

import java.io.Serializable;

@Data
@EqualsAndHashCode
public class CustomerBean implements Serializable {
    private static final long serialVersionUID = -806621178467580736L;

    private String firstName;

    private String lastName;

    private int age;
}

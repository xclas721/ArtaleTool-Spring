package com.artale.artaletool.model;

import lombok.Data;

@Data
public class KeyEvent {
    private long timestamp;
    private String key;
    private String action; // "PRESS" or "RELEASE"
} 
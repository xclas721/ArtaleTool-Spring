/*
 * Copyright (c) 2024 ArtaleTool
 * All rights reserved.
 */
package com.artale.artaletool.model;

import lombok.Data;

@Data
public class KeyEvent {
  private long timestamp;
  private String key;
  private String action; // "PRESS" or "RELEASE"
}

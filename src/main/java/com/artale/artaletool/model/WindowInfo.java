/*
 * Copyright (c) 2024 ArtaleTool
 * All rights reserved.
 */
package com.artale.artaletool.model;

import lombok.Data;

@Data
public class WindowInfo {
  private long handle; // 視窗句柄
  private String title; // 視窗標題
  private String className; // 視窗類別名稱
  private boolean isVisible; // 是否可見
  private boolean isActive; // 是否為活動視窗
  private int x, y; // 視窗位置
  private int width, height; // 視窗大小
  private boolean isSizePositionLocked; // 視窗大小位置是否被鎖定
}

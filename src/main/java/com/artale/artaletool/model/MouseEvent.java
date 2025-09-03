/*
 * Copyright (c) 2024 ArtaleTool
 * All rights reserved.
 */
package com.artale.artaletool.model;

import lombok.Data;

@Data
public class MouseEvent {
  private String action; // PRESS, RELEASE
  private int button; // 滑鼠按鍵 (1=左鍵, 2=右鍵, 3=中鍵)
  private int x; // X 座標
  private int y; // Y 座標
  private long timestamp; // 時間戳記
  private long delay; // 與上一個事件的延遲時間
}

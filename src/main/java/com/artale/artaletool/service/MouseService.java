/*
 * Copyright (c) 2024 ArtaleTool
 * All rights reserved.
 */
package com.artale.artaletool.service;

import java.awt.AWTException;
import java.awt.Robot;
import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;

import com.artale.artaletool.model.MouseEvent;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.sun.jna.Native;
import com.sun.jna.win32.W32APIOptions;

import jakarta.annotation.PreDestroy;

@Service
public class MouseService {
  private static final Logger logger = LoggerFactory.getLogger(MouseService.class);
  private List<MouseEvent> recordedEvents = new ArrayList<>();
  private boolean isRecording = false;
  private boolean isPlaying = false;
  private final ObjectMapper objectMapper = new ObjectMapper();
  private final String SCRIPTS_DIR = "mouse_scripts";
  private Robot robot;
  private Thread mouseMonitorThread;
  private final User32 user32 = User32.INSTANCE;
  private boolean isLooping = false;
  private int loopCount = 0;
  private int currentLoop = 0;
  private MouseEvent currentPlayingEvent = null;
  private int currentPlayingIndex = -1;
  private final Map<Integer, Boolean> keyStates = new HashMap<>();

  public interface User32 extends com.sun.jna.platform.win32.User32 {
    User32 INSTANCE = Native.load("user32", User32.class, W32APIOptions.DEFAULT_OPTIONS);

    boolean GetCursorPos(int[] lpPoint);

    void mouse_event(int dwFlags, int dx, int dy, int dwData, int dwExtraInfo);
  }

  // 滑鼠事件常量
  private static final int MOUSEEVENTF_LEFTDOWN = 0x0002;
  private static final int MOUSEEVENTF_LEFTUP = 0x0004;
  private static final int MOUSEEVENTF_RIGHTDOWN = 0x0008;
  private static final int MOUSEEVENTF_RIGHTUP = 0x0010;
  private static final int MOUSEEVENTF_MIDDLEDOWN = 0x0020;
  private static final int MOUSEEVENTF_MIDDLEUP = 0x0040;
  private static final int MOUSEEVENTF_MOVE = 0x0001;
  private static final int MOUSEEVENTF_ABSOLUTE = 0x8000;

  // 虛擬鍵碼常量
  private static final int VK_LBUTTON = 0x01;
  private static final int VK_RBUTTON = 0x02;
  private static final int VK_MBUTTON = 0x04;

  // 快捷鍵常量
  private static final int VK_F1 = java.awt.event.KeyEvent.VK_F1;
  private static final int VK_F2 = java.awt.event.KeyEvent.VK_F2;
  private static final int VK_ESCAPE = java.awt.event.KeyEvent.VK_ESCAPE;

  public MouseService() {
    try {
      Files.createDirectories(Paths.get(SCRIPTS_DIR));
      logger.info("滑鼠腳本目錄創建成功: {}", SCRIPTS_DIR);

      // 設置系統屬性以允許在 headless 環境中創建 Robot
      System.setProperty("java.awt.headless", "false");
      try {
        robot = new Robot();
        logger.info("Robot 初始化成功");
      } catch (AWTException e) {
        logger.error("Robot 初始化失敗: {}", e.getMessage());
      }
    } catch (Exception e) {
      logger.error("滑鼠服務初始化失敗: {}", e.getMessage());
    }

    // 啟動滑鼠和快捷鍵監聽
    startMouseAndHotkeyMonitor();
  }

  /** 啟動滑鼠和快捷鍵監聽 */
  private void startMouseAndHotkeyMonitor() {
    if (mouseMonitorThread != null && mouseMonitorThread.isAlive()) {
      return;
    }

    mouseMonitorThread =
        new Thread(
            () -> {
              logger.info("開始監控滑鼠事件和快捷鍵");
              long startTime = System.currentTimeMillis();
              long lastEventTime = startTime;

              while (!Thread.currentThread().isInterrupted()) {
                try {
                  // 檢查快捷鍵狀態
                  short f1State = user32.GetAsyncKeyState(VK_F1);
                  short f2State = user32.GetAsyncKeyState(VK_F2);
                  short escapeState = user32.GetAsyncKeyState(VK_ESCAPE);

                  // 檢查 F1 鍵 (開始錄製)
                  boolean f1Pressed = (f1State & 0x8000) != 0;
                  Boolean f1WasPressed = keyStates.get(VK_F1);
                  if (f1Pressed && (f1WasPressed == null || !f1WasPressed)) {
                    if (!isRecording) {
                      logger.info("檢測到 F1 快捷鍵，開始錄製滑鼠事件");
                      startRecording();
                    }
                    keyStates.put(VK_F1, true);
                  } else if (!f1Pressed && f1WasPressed != null && f1WasPressed) {
                    keyStates.put(VK_F1, false);
                  }

                  // 檢查 F2 鍵 (停止錄製)
                  boolean f2Pressed = (f2State & 0x8000) != 0;
                  Boolean f2WasPressed = keyStates.get(VK_F2);
                  if (f2Pressed && (f2WasPressed == null || !f2WasPressed)) {
                    if (isRecording) {
                      logger.info("檢測到 F2 快捷鍵，停止錄製滑鼠事件");
                      stopRecording();
                    }
                    keyStates.put(VK_F2, true);
                  } else if (!f2Pressed && f2WasPressed != null && f2WasPressed) {
                    keyStates.put(VK_F2, false);
                  }

                  // 檢查 ESC 鍵 (停止播放)
                  boolean escapePressed = (escapeState & 0x8000) != 0;
                  Boolean escapeWasPressed = keyStates.get(VK_ESCAPE);
                  if (escapePressed && (escapeWasPressed == null || !escapeWasPressed)) {
                    if (isPlaying) {
                      logger.info("檢測到 ESC 快捷鍵，停止播放滑鼠腳本");
                      stopPlayback();
                    }
                    keyStates.put(VK_ESCAPE, true);
                  } else if (!escapePressed && escapeWasPressed != null && escapeWasPressed) {
                    keyStates.put(VK_ESCAPE, false);
                  }

                  // 調試快捷鍵狀態
                  if (f1Pressed || f2Pressed || escapePressed) {
                    logger.debug(
                        "快捷鍵狀態 - F1: {}, F2: {}, ESC: {}", f1Pressed, f2Pressed, escapePressed);
                  }

                  // 如果正在錄製，檢查滑鼠按鍵狀態
                  if (isRecording) {
                    // 檢查滑鼠按鍵狀態
                    short leftButtonState = user32.GetAsyncKeyState(VK_LBUTTON);
                    short rightButtonState = user32.GetAsyncKeyState(VK_RBUTTON);
                    short middleButtonState = user32.GetAsyncKeyState(VK_MBUTTON);

                    // 獲取當前滑鼠位置
                    int[] cursorPos = new int[2];
                    user32.GetCursorPos(cursorPos);
                    int x = cursorPos[0];
                    int y = cursorPos[1];

                    long currentTime = System.currentTimeMillis();
                    long delay = currentTime - lastEventTime;

                    // 檢查左鍵
                    if ((leftButtonState & 0x8000) != 0) {
                      MouseEvent event = new MouseEvent();
                      event.setAction("PRESS");
                      event.setButton(1);
                      event.setX(x);
                      event.setY(y);
                      event.setTimestamp(currentTime);
                      event.setDelay(delay);
                      recordedEvents.add(event);
                      lastEventTime = currentTime;
                      logger.debug("錄製左鍵按下: ({}, {})", x, y);
                    } else if ((leftButtonState & 0x0001) != 0) {
                      MouseEvent event = new MouseEvent();
                      event.setAction("RELEASE");
                      event.setButton(1);
                      event.setX(x);
                      event.setY(y);
                      event.setTimestamp(currentTime);
                      event.setDelay(delay);
                      recordedEvents.add(event);
                      lastEventTime = currentTime;
                      logger.debug("錄製左鍵釋放: ({}, {})", x, y);
                    }

                    // 檢查右鍵
                    if ((rightButtonState & 0x8000) != 0) {
                      MouseEvent event = new MouseEvent();
                      event.setAction("PRESS");
                      event.setButton(2);
                      event.setX(x);
                      event.setY(y);
                      event.setTimestamp(currentTime);
                      event.setDelay(delay);
                      recordedEvents.add(event);
                      lastEventTime = currentTime;
                      logger.debug("錄製右鍵按下: ({}, {})", x, y);
                    } else if ((rightButtonState & 0x0001) != 0) {
                      MouseEvent event = new MouseEvent();
                      event.setAction("RELEASE");
                      event.setButton(2);
                      event.setX(x);
                      event.setY(y);
                      event.setTimestamp(currentTime);
                      event.setDelay(delay);
                      recordedEvents.add(event);
                      lastEventTime = currentTime;
                      logger.debug("錄製右鍵釋放: ({}, {})", x, y);
                    }

                    // 檢查中鍵
                    if ((middleButtonState & 0x8000) != 0) {
                      MouseEvent event = new MouseEvent();
                      event.setAction("PRESS");
                      event.setButton(3);
                      event.setX(x);
                      event.setY(y);
                      event.setTimestamp(currentTime);
                      event.setDelay(delay);
                      recordedEvents.add(event);
                      lastEventTime = currentTime;
                      logger.debug("錄製中鍵按下: ({}, {})", x, y);
                    } else if ((middleButtonState & 0x0001) != 0) {
                      MouseEvent event = new MouseEvent();
                      event.setAction("RELEASE");
                      event.setButton(3);
                      event.setX(x);
                      event.setY(y);
                      event.setTimestamp(currentTime);
                      event.setDelay(delay);
                      recordedEvents.add(event);
                      lastEventTime = currentTime;
                      logger.debug("錄製中鍵釋放: ({}, {})", x, y);
                    }
                  }

                  // 短暫休眠以避免過度佔用 CPU
                  Thread.sleep(10);
                } catch (InterruptedException e) {
                  logger.info("滑鼠監控線程被中斷");
                  break;
                } catch (Exception e) {
                  logger.error("滑鼠監控過程中發生錯誤: {}", e.getMessage());
                }
              }
              logger.info("滑鼠監控線程結束");
            },
            "MouseMonitorThread");

    mouseMonitorThread.setDaemon(true);
    mouseMonitorThread.start();
    logger.info("滑鼠監控線程已啟動");
  }

  /** 開始錄製滑鼠事件 */
  public void startRecording() {
    if (isRecording) {
      logger.warn("滑鼠錄製已經在進行中");
      return;
    }

    recordedEvents.clear();
    isRecording = true;
    logger.info("開始錄製滑鼠事件");
  }

  /** 停止錄製滑鼠事件 */
  public void stopRecording() {
    if (!isRecording) {
      logger.warn("滑鼠錄製未在進行中");
      return;
    }

    isRecording = false;
    logger.info("停止錄製滑鼠事件，共錄製 {} 個事件", recordedEvents.size());
  }

  /** 播放錄製的滑鼠腳本 */
  public void playScript(List<MouseEvent> events, boolean loop, int loopCount) {
    if (isPlaying) {
      logger.warn("滑鼠腳本正在播放中");
      return;
    }

    if (events == null || events.isEmpty()) {
      logger.warn("沒有可播放的滑鼠事件");
      return;
    }

    this.isLooping = loop;
    this.loopCount = loopCount;
    this.currentLoop = 0;

    new Thread(
            () -> {
              try {
                playMouseEvents(events);
              } catch (Exception e) {
                logger.error("播放滑鼠腳本時發生錯誤: {}", e.getMessage());
                isPlaying = false;
              }
            },
            "MousePlaybackThread")
        .start();
  }

  /** 播放滑鼠事件 */
  private void playMouseEvents(List<MouseEvent> events) {
    isPlaying = true;
    logger.info("開始播放滑鼠腳本，共 {} 個事件", events.size());

    do {
      currentLoop++;
      if (isLooping && loopCount > 0) {
        logger.info("播放第 {} 次循環 (共 {} 次)", currentLoop, loopCount);
      }

      for (int i = 0; i < events.size() && isPlaying; i++) {
        MouseEvent event = events.get(i);
        currentPlayingEvent = event;
        currentPlayingIndex = i;

        try {
          // 移動滑鼠到指定位置
          robot.mouseMove(event.getX(), event.getY());

          // 根據按鍵類型執行相應動作
          switch (event.getButton()) {
            case 1: // 左鍵
              if ("PRESS".equals(event.getAction())) {
                robot.mousePress(java.awt.event.InputEvent.BUTTON1_DOWN_MASK);
              } else if ("RELEASE".equals(event.getAction())) {
                robot.mouseRelease(java.awt.event.InputEvent.BUTTON1_DOWN_MASK);
              }
              break;
            case 2: // 右鍵
              if ("PRESS".equals(event.getAction())) {
                robot.mousePress(java.awt.event.InputEvent.BUTTON3_DOWN_MASK);
              } else if ("RELEASE".equals(event.getAction())) {
                robot.mouseRelease(java.awt.event.InputEvent.BUTTON3_DOWN_MASK);
              }
              break;
            case 3: // 中鍵
              if ("PRESS".equals(event.getAction())) {
                robot.mousePress(java.awt.event.InputEvent.BUTTON2_DOWN_MASK);
              } else if ("RELEASE".equals(event.getAction())) {
                robot.mouseRelease(java.awt.event.InputEvent.BUTTON2_DOWN_MASK);
              }
              break;
          }

          // 等待指定的延遲時間
          if (event.getDelay() > 0) {
            Thread.sleep(event.getDelay());
          }

          logger.debug(
              "播放事件 {}: {} 按鍵 {} 在 ({}, {})",
              i,
              event.getAction(),
              event.getButton(),
              event.getX(),
              event.getY());
        } catch (InterruptedException e) {
          logger.info("滑鼠腳本播放被中斷");
          break;
        } catch (Exception e) {
          logger.error("播放滑鼠事件時發生錯誤: {}", e.getMessage());
        }
      }

      // 檢查是否需要繼續循環
      if (isLooping && (loopCount == 0 || currentLoop < loopCount)) {
        logger.info("準備開始下一次循環");
        try {
          Thread.sleep(1000); // 循環間隔 1 秒
        } catch (InterruptedException e) {
          break;
        }
      } else {
        break;
      }
    } while (isLooping && (loopCount == 0 || currentLoop < loopCount));

    isPlaying = false;
    currentPlayingEvent = null;
    currentPlayingIndex = -1;
    logger.info("滑鼠腳本播放完成");
  }

  /** 停止播放 */
  public void stopPlayback() {
    if (!isPlaying) {
      logger.warn("滑鼠腳本未在播放中");
      return;
    }

    isPlaying = false;
    logger.info("停止播放滑鼠腳本");
  }

  /** 暫停播放 */
  public void pausePlayback() {
    if (!isPlaying) {
      logger.warn("滑鼠腳本未在播放中");
      return;
    }

    isPlaying = false;
    logger.info("暫停播放滑鼠腳本");
  }

  /** 獲取錄製的事件列表 */
  public List<MouseEvent> getRecordedEvents() {
    return new ArrayList<>(recordedEvents);
  }

  /** 清空錄製的事件 */
  public void clearRecordedEvents() {
    recordedEvents.clear();
    logger.info("已清空錄製的滑鼠事件");
  }

  /** 獲取當前錄製狀態 */
  public boolean isRecording() {
    return isRecording;
  }

  /** 獲取當前播放狀態 */
  public boolean isPlaying() {
    return isPlaying;
  }

  /** 獲取當前播放的事件 */
  public MouseEvent getCurrentPlayingEvent() {
    return currentPlayingEvent;
  }

  /** 獲取當前播放的索引 */
  public int getCurrentPlayingIndex() {
    return currentPlayingIndex;
  }

  /** 獲取循環播放狀態 */
  public boolean isLooping() {
    return isLooping;
  }

  /** 獲取當前循環次數 */
  public int getCurrentLoop() {
    return currentLoop;
  }

  /** 獲取總循環次數 */
  public int getLoopCount() {
    return loopCount;
  }

  /** 保存腳本到檔案 */
  public void saveScript(String filename) throws IOException {
    if (recordedEvents.isEmpty()) {
      logger.warn("沒有可保存的滑鼠事件");
      return;
    }

    File file = new File(SCRIPTS_DIR, filename + ".json");
    objectMapper.writeValue(file, recordedEvents);
    logger.info("滑鼠腳本已保存到: {}", file.getAbsolutePath());
  }

  /** 從檔案載入腳本 */
  public List<MouseEvent> loadScript(String filename) throws IOException {
    File file = new File(SCRIPTS_DIR, filename + ".json");
    if (!file.exists()) {
      throw new IOException("腳本檔案不存在: " + file.getAbsolutePath());
    }

    List<MouseEvent> events =
        objectMapper.readValue(
            file,
            objectMapper.getTypeFactory().constructCollectionType(List.class, MouseEvent.class));
    logger.info("從檔案載入滑鼠腳本: {}，共 {} 個事件", filename, events.size());
    return events;
  }

  /** 獲取腳本目錄中的所有腳本檔案 */
  public List<String> getScriptFiles() {
    List<String> scripts = new ArrayList<>();
    File dir = new File(SCRIPTS_DIR);
    if (dir.exists() && dir.isDirectory()) {
      File[] files = dir.listFiles((d, name) -> name.toLowerCase().endsWith(".json"));
      if (files != null) {
        for (File file : files) {
          scripts.add(file.getName().replace(".json", ""));
        }
      }
    }
    return scripts;
  }

  /** 刪除腳本檔案 */
  public boolean deleteScript(String filename) {
    File file = new File(SCRIPTS_DIR, filename + ".json");
    if (file.exists()) {
      boolean deleted = file.delete();
      if (deleted) {
        logger.info("腳本檔案已刪除: {}", filename);
      } else {
        logger.warn("無法刪除腳本檔案: {}", filename);
      }
      return deleted;
    } else {
      logger.warn("腳本檔案不存在: {}", filename);
      return false;
    }
  }

  /** 獲取當前滑鼠位置 */
  public int[] getCurrentMousePosition() {
    int[] cursorPos = new int[2];
    user32.GetCursorPos(cursorPos);
    return cursorPos;
  }

  /** 移動滑鼠到指定位置 */
  public void moveMouse(int x, int y) {
    if (robot != null) {
      robot.mouseMove(x, y);
      logger.debug("移動滑鼠到位置: ({}, {})", x, y);
    } else {
      logger.warn("Robot 未初始化，無法移動滑鼠");
    }
  }

  /** 點擊滑鼠左鍵 */
  public void clickLeftButton() {
    if (robot != null) {
      robot.mousePress(java.awt.event.InputEvent.BUTTON1_DOWN_MASK);
      robot.mouseRelease(java.awt.event.InputEvent.BUTTON1_DOWN_MASK);
      logger.debug("點擊滑鼠左鍵");
    } else {
      logger.warn("Robot 未初始化，無法點擊滑鼠");
    }
  }

  /** 點擊滑鼠右鍵 */
  public void clickRightButton() {
    if (robot != null) {
      robot.mousePress(java.awt.event.InputEvent.BUTTON3_DOWN_MASK);
      robot.mouseRelease(java.awt.event.InputEvent.BUTTON3_DOWN_MASK);
      logger.debug("點擊滑鼠右鍵");
    } else {
      logger.warn("Robot 未初始化，無法點擊滑鼠");
    }
  }

  /** 點擊滑鼠中鍵 */
  public void clickMiddleButton() {
    if (robot != null) {
      robot.mousePress(java.awt.event.InputEvent.BUTTON2_DOWN_MASK);
      robot.mouseRelease(java.awt.event.InputEvent.BUTTON2_DOWN_MASK);
      logger.debug("點擊滑鼠中鍵");
    } else {
      logger.warn("Robot 未初始化，無法點擊滑鼠");
    }
  }

  @PreDestroy
  public void cleanup() {
    logger.info("清理滑鼠服務資源");

    // 停止錄製和播放
    isRecording = false;
    isPlaying = false;

    // 中斷監控線程
    if (mouseMonitorThread != null && mouseMonitorThread.isAlive()) {
      mouseMonitorThread.interrupt();
      try {
        mouseMonitorThread.join(1000);
      } catch (InterruptedException e) {
        logger.warn("等待監控線程結束時被中斷");
      }
    }

    logger.info("滑鼠服務資源清理完成");
  }
}

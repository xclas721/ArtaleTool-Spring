/*
 * Copyright (c) 2024 ArtaleTool
 * All rights reserved.
 */
package com.artale.artaletool.service;

import java.awt.AWTException;
import java.awt.Frame;
import java.awt.Robot;
import java.awt.event.KeyListener;
import java.io.File;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.CopyOnWriteArrayList;
import java.util.concurrent.Executors;
import java.util.concurrent.ScheduledExecutorService;
import java.util.concurrent.TimeUnit;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Service;

import com.artale.artaletool.model.KeyEvent;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.sun.jna.*;
import com.sun.jna.platform.win32.*;
import com.sun.jna.win32.*;

import jakarta.annotation.PreDestroy;

@Service
public class KeyboardService implements KeyListener {
  private static final Logger logger = LoggerFactory.getLogger(KeyboardService.class);
  private List<KeyEvent> recordedEvents = new ArrayList<>();
  private boolean isRecording = false;
  private boolean isPlaying = false;
  private final ObjectMapper objectMapper = new ObjectMapper();
  private final String SCRIPTS_DIR = "scripts";
  private final List<String> currentPressedKeys = new CopyOnWriteArrayList<>();
  private final ScheduledExecutorService scheduler = Executors.newScheduledThreadPool(1);
  private Robot robot;
  private final Map<String, Integer> keyCodeMap = new HashMap<>();
  private final Map<Integer, String> reverseKeyCodeMap = new HashMap<>();
  private Frame frame;
  private final User32 user32 = User32.INSTANCE;
  private final Map<Integer, Boolean> keyStates = new HashMap<>();
  private Thread keyMonitorThread;
  private boolean isLooping = false;
  private int loopCount = 0;
  private int currentLoop = 0;
  private final Map<String, ScheduledExecutorService> scheduledTasks = new HashMap<>();
  private final Map<String, Integer> scheduledKeyCodes = new HashMap<>();
  private static final int VK_ESCAPE = 0x1B; // ESC 鍵的虛擬鍵碼
  private KeyEvent currentPlayingEvent = null;
  private int currentPlayingIndex = -1;

  @Autowired private WindowService windowService;

  public interface User32 extends StdCallLibrary {
    User32 INSTANCE = Native.load("user32", User32.class, W32APIOptions.DEFAULT_OPTIONS);

    short GetAsyncKeyState(int vKey);
  }

  public KeyboardService() {
    try {
      Files.createDirectories(Paths.get(SCRIPTS_DIR));
      logger.info("腳本目錄創建成功: {}", SCRIPTS_DIR);

      // 設置系統屬性以允許在 headless 環境中創建 Robot
      System.setProperty("java.awt.headless", "false");
      try {
        robot = new Robot();
        logger.info("Robot 初始化成功");
      } catch (AWTException e) {
        logger.error("Robot 初始化失敗: {}", e.getMessage());
      }

      // 初始化按鍵映射
      initializeKeyCodeMap();
      logger.info("按鍵監聽器初始化成功");

      // 啟動按鍵監聽
      startKeyMonitor();
    } catch (Exception e) {
      logger.error("初始化失敗: {}", e.getMessage());
    }
  }

  private void startKeyMonitor() {
    if (keyMonitorThread != null && keyMonitorThread.isAlive()) {
      return;
    }

    logger.info("開始監控按鍵事件");
    long startTime = System.currentTimeMillis();
    keyMonitorThread =
        new Thread(
            () -> {
              while (!Thread.currentThread().isInterrupted()) {
                try {
                  // 檢查所有按鍵狀態
                  for (Map.Entry<String, Integer> entry : keyCodeMap.entrySet()) {
                    String key = entry.getKey();
                    int vKey = entry.getValue();
                    short keyState = user32.GetAsyncKeyState(vKey);
                    boolean isPressed = (keyState & 0x8000) != 0;
                    Boolean wasPressed = keyStates.get(vKey);

                    if (isPressed && (wasPressed == null || !wasPressed)) {
                      // 檢查視窗鎖定狀態
                      if (windowService != null
                          && windowService.isWindowLocked()
                          && !windowService.isLockedWindowActive()) {
                        logger.debug("視窗未鎖定，自動切換回鎖定視窗: {}", key);
                        // 自動將鎖定視窗帶到前台
                        windowService.bringLockedWindowToFront();
                        // 等待視窗切換完成
                        Thread.sleep(100);
                        continue;
                      }

                      // 按鍵按下
                      long currentTime = System.currentTimeMillis();
                      long relativeTime = currentTime - startTime;

                      if (isRecording) {
                        logger.info("[錄製中] 時間: {}ms, 按鍵: {}, 動作: PRESS", relativeTime, key);
                        recordKeyPress(vKey);
                      } else {
                        logger.debug("時間: {}ms, 按鍵: {}, 動作: PRESS", relativeTime, key);
                      }

                      // 如果是 ESC 鍵且正在播放，則停止播放
                      if (key.equals("Escape") && isPlaying) {
                        logger.info("檢測到 ESC 鍵按下，停止播放");
                        stopPlayback();
                      }

                      keyStates.put(vKey, true);
                    } else if (!isPressed && wasPressed != null && wasPressed) {
                      // 按鍵釋放
                      long currentTime = System.currentTimeMillis();
                      long relativeTime = currentTime - startTime;

                      if (isRecording) {
                        logger.info("[錄製中] 時間: {}ms, 按鍵: {}, 動作: RELEASE", relativeTime, key);
                        recordKeyRelease(vKey);
                      } else {
                        logger.debug("時間: {}ms, 按鍵: {}, 動作: RELEASE", relativeTime, key);
                      }

                      keyStates.put(vKey, false);
                    }
                  }
                  Thread.sleep(10); // 10ms 的輪詢間隔
                } catch (InterruptedException e) {
                  logger.info("按鍵監控執行緒被中斷");
                  Thread.currentThread().interrupt();
                  break;
                } catch (Exception e) {
                  logger.error("按鍵監控發生錯誤: {}", e.getMessage());
                }
              }
              logger.info("停止監控按鍵事件");
            });
    keyMonitorThread.setDaemon(true);
    keyMonitorThread.start();
  }

  @PreDestroy
  public void cleanup() {
    try {
      if (isRecording) {
        stopRecording();
      }
      if (isPlaying) {
        stopPlayback();
      }
      stopAllScheduledTasks();
      if (keyMonitorThread != null) {
        keyMonitorThread.interrupt();
      }
      if (scheduler != null && !scheduler.isShutdown()) {
        scheduler.shutdown();
        if (!scheduler.awaitTermination(1, TimeUnit.SECONDS)) {
          scheduler.shutdownNow();
        }
      }
      logger.info("資源清理完成");
    } catch (Exception e) {
      logger.error("資源清理時發生錯誤: {}", e.getMessage());
    }
  }

  @Override
  public void keyPressed(java.awt.event.KeyEvent e) {
    if (isRecording) {
      logger.debug("按鍵按下: {}", e.getKeyCode());
      recordKeyPress(e.getKeyCode());
    }
  }

  @Override
  public void keyReleased(java.awt.event.KeyEvent e) {
    if (isRecording) {
      logger.debug("按鍵釋放: {}", e.getKeyCode());
      recordKeyRelease(e.getKeyCode());
    }
  }

  @Override
  public void keyTyped(java.awt.event.KeyEvent e) {
    // 不需要處理
  }

  private void initializeKeyCodeMap() {
    try {
      // 字母鍵
      for (char c = 'A'; c <= 'Z'; c++) {
        int keyCode = java.awt.event.KeyEvent.class.getField("VK_" + c).getInt(null);
        keyCodeMap.put(String.valueOf(c), keyCode);
        reverseKeyCodeMap.put(keyCode, String.valueOf(c));
      }

      // 數字鍵
      for (char c = '0'; c <= '9'; c++) {
        int keyCode = java.awt.event.KeyEvent.class.getField("VK_" + c).getInt(null);
        keyCodeMap.put(String.valueOf(c), keyCode);
        reverseKeyCodeMap.put(keyCode, String.valueOf(c));
      }

      // 功能鍵
      for (int i = 1; i <= 12; i++) {
        int keyCode = java.awt.event.KeyEvent.class.getField("VK_F" + i).getInt(null);
        keyCodeMap.put("F" + i, keyCode);
        reverseKeyCodeMap.put(keyCode, "F" + i);
      }

      // 特殊按鍵
      keyCodeMap.put("Space", java.awt.event.KeyEvent.VK_SPACE);
      keyCodeMap.put("Enter", java.awt.event.KeyEvent.VK_ENTER);
      keyCodeMap.put("Escape", java.awt.event.KeyEvent.VK_ESCAPE);
      keyCodeMap.put("Tab", java.awt.event.KeyEvent.VK_TAB);
      keyCodeMap.put("Backspace", java.awt.event.KeyEvent.VK_BACK_SPACE);
      keyCodeMap.put("Delete", java.awt.event.KeyEvent.VK_DELETE);
      keyCodeMap.put("Insert", java.awt.event.KeyEvent.VK_INSERT);
      keyCodeMap.put("Home", java.awt.event.KeyEvent.VK_HOME);
      keyCodeMap.put("End", java.awt.event.KeyEvent.VK_END);
      keyCodeMap.put("PageUp", java.awt.event.KeyEvent.VK_PAGE_UP);
      keyCodeMap.put("PageDown", java.awt.event.KeyEvent.VK_PAGE_DOWN);

      // 方向鍵
      keyCodeMap.put("Left", java.awt.event.KeyEvent.VK_LEFT);
      keyCodeMap.put("Right", java.awt.event.KeyEvent.VK_RIGHT);
      keyCodeMap.put("Up", java.awt.event.KeyEvent.VK_UP);
      keyCodeMap.put("Down", java.awt.event.KeyEvent.VK_DOWN);

      // 修飾鍵
      keyCodeMap.put("Shift", java.awt.event.KeyEvent.VK_SHIFT);
      keyCodeMap.put("Ctrl", java.awt.event.KeyEvent.VK_CONTROL);
      keyCodeMap.put("Alt", java.awt.event.KeyEvent.VK_ALT);
      keyCodeMap.put("Windows", java.awt.event.KeyEvent.VK_WINDOWS);
      keyCodeMap.put("Command", java.awt.event.KeyEvent.VK_META);

      // 反向映射
      reverseKeyCodeMap.put(java.awt.event.KeyEvent.VK_SPACE, "Space");
      reverseKeyCodeMap.put(java.awt.event.KeyEvent.VK_ENTER, "Enter");
      reverseKeyCodeMap.put(java.awt.event.KeyEvent.VK_ESCAPE, "Escape");
      reverseKeyCodeMap.put(java.awt.event.KeyEvent.VK_TAB, "Tab");
      reverseKeyCodeMap.put(java.awt.event.KeyEvent.VK_BACK_SPACE, "Backspace");
      reverseKeyCodeMap.put(java.awt.event.KeyEvent.VK_DELETE, "Delete");
      reverseKeyCodeMap.put(java.awt.event.KeyEvent.VK_INSERT, "Insert");
      reverseKeyCodeMap.put(java.awt.event.KeyEvent.VK_HOME, "Home");
      reverseKeyCodeMap.put(java.awt.event.KeyEvent.VK_END, "End");
      reverseKeyCodeMap.put(java.awt.event.KeyEvent.VK_PAGE_UP, "PageUp");
      reverseKeyCodeMap.put(java.awt.event.KeyEvent.VK_PAGE_DOWN, "PageDown");
      reverseKeyCodeMap.put(java.awt.event.KeyEvent.VK_LEFT, "Left");
      reverseKeyCodeMap.put(java.awt.event.KeyEvent.VK_RIGHT, "Right");
      reverseKeyCodeMap.put(java.awt.event.KeyEvent.VK_UP, "Up");
      reverseKeyCodeMap.put(java.awt.event.KeyEvent.VK_DOWN, "Down");
      reverseKeyCodeMap.put(java.awt.event.KeyEvent.VK_SHIFT, "Shift");
      reverseKeyCodeMap.put(java.awt.event.KeyEvent.VK_CONTROL, "Ctrl");
      reverseKeyCodeMap.put(java.awt.event.KeyEvent.VK_ALT, "Alt");
      reverseKeyCodeMap.put(java.awt.event.KeyEvent.VK_WINDOWS, "Windows");
      reverseKeyCodeMap.put(java.awt.event.KeyEvent.VK_META, "Command");

      // 額外的特殊按鍵
      keyCodeMap.put("NumLock", java.awt.event.KeyEvent.VK_NUM_LOCK);
      keyCodeMap.put("ScrollLock", java.awt.event.KeyEvent.VK_SCROLL_LOCK);
      keyCodeMap.put("CapsLock", java.awt.event.KeyEvent.VK_CAPS_LOCK);
      keyCodeMap.put("PrintScreen", java.awt.event.KeyEvent.VK_PRINTSCREEN);
      keyCodeMap.put("Pause", java.awt.event.KeyEvent.VK_PAUSE);
      keyCodeMap.put("ContextMenu", java.awt.event.KeyEvent.VK_CONTEXT_MENU);

      // 數字鍵盤
      keyCodeMap.put("NumPad0", java.awt.event.KeyEvent.VK_NUMPAD0);
      keyCodeMap.put("NumPad1", java.awt.event.KeyEvent.VK_NUMPAD1);
      keyCodeMap.put("NumPad2", java.awt.event.KeyEvent.VK_NUMPAD2);
      keyCodeMap.put("NumPad3", java.awt.event.KeyEvent.VK_NUMPAD3);
      keyCodeMap.put("NumPad4", java.awt.event.KeyEvent.VK_NUMPAD4);
      keyCodeMap.put("NumPad5", java.awt.event.KeyEvent.VK_NUMPAD5);
      keyCodeMap.put("NumPad6", java.awt.event.KeyEvent.VK_NUMPAD6);
      keyCodeMap.put("NumPad7", java.awt.event.KeyEvent.VK_NUMPAD7);
      keyCodeMap.put("NumPad8", java.awt.event.KeyEvent.VK_NUMPAD8);
      keyCodeMap.put("NumPad9", java.awt.event.KeyEvent.VK_NUMPAD9);
      keyCodeMap.put("NumPadAdd", java.awt.event.KeyEvent.VK_ADD);
      keyCodeMap.put("NumPadSubtract", java.awt.event.KeyEvent.VK_SUBTRACT);
      keyCodeMap.put("NumPadMultiply", java.awt.event.KeyEvent.VK_MULTIPLY);
      keyCodeMap.put("NumPadDivide", java.awt.event.KeyEvent.VK_DIVIDE);
      keyCodeMap.put("NumPadDecimal", java.awt.event.KeyEvent.VK_DECIMAL);
      keyCodeMap.put("NumPadEnter", java.awt.event.KeyEvent.VK_ENTER);

      // 反向映射
      reverseKeyCodeMap.put(java.awt.event.KeyEvent.VK_NUM_LOCK, "NumLock");
      reverseKeyCodeMap.put(java.awt.event.KeyEvent.VK_SCROLL_LOCK, "ScrollLock");
      reverseKeyCodeMap.put(java.awt.event.KeyEvent.VK_CAPS_LOCK, "CapsLock");
      reverseKeyCodeMap.put(java.awt.event.KeyEvent.VK_PRINTSCREEN, "PrintScreen");
      reverseKeyCodeMap.put(java.awt.event.KeyEvent.VK_PAUSE, "Pause");
      reverseKeyCodeMap.put(java.awt.event.KeyEvent.VK_CONTEXT_MENU, "ContextMenu");
      reverseKeyCodeMap.put(java.awt.event.KeyEvent.VK_NUMPAD0, "NumPad0");
      reverseKeyCodeMap.put(java.awt.event.KeyEvent.VK_NUMPAD1, "NumPad1");
      reverseKeyCodeMap.put(java.awt.event.KeyEvent.VK_NUMPAD2, "NumPad2");
      reverseKeyCodeMap.put(java.awt.event.KeyEvent.VK_NUMPAD3, "NumPad3");
      reverseKeyCodeMap.put(java.awt.event.KeyEvent.VK_NUMPAD4, "NumPad4");
      reverseKeyCodeMap.put(java.awt.event.KeyEvent.VK_NUMPAD5, "NumPad5");
      reverseKeyCodeMap.put(java.awt.event.KeyEvent.VK_NUMPAD6, "NumPad6");
      reverseKeyCodeMap.put(java.awt.event.KeyEvent.VK_NUMPAD7, "NumPad7");
      reverseKeyCodeMap.put(java.awt.event.KeyEvent.VK_NUMPAD8, "NumPad8");
      reverseKeyCodeMap.put(java.awt.event.KeyEvent.VK_NUMPAD9, "NumPad9");
      reverseKeyCodeMap.put(java.awt.event.KeyEvent.VK_ADD, "NumPadAdd");
      reverseKeyCodeMap.put(java.awt.event.KeyEvent.VK_SUBTRACT, "NumPadSubtract");
      reverseKeyCodeMap.put(java.awt.event.KeyEvent.VK_MULTIPLY, "NumPadMultiply");
      reverseKeyCodeMap.put(java.awt.event.KeyEvent.VK_DIVIDE, "NumPadDivide");
      reverseKeyCodeMap.put(java.awt.event.KeyEvent.VK_DECIMAL, "NumPadDecimal");
      reverseKeyCodeMap.put(java.awt.event.KeyEvent.VK_ENTER, "NumPadEnter");

    } catch (NoSuchFieldException | IllegalAccessException e) {
      logger.error("初始化按鍵映射時發生錯誤", e);
    }
  }

  public void startRecording() {
    if (isRecording) {
      logger.warn("已經在錄製中");
      return;
    }

    recordedEvents.clear();
    isRecording = true;
    logger.info("=== 開始錄製鍵盤事件 ===");
    logger.info("當前已記錄的事件數: {}", recordedEvents.size());
  }

  public List<KeyEvent> stopRecording() {
    if (!isRecording) {
      logger.warn("沒有在錄製中");
      return new ArrayList<>();
    }

    isRecording = false;
    logger.info("=== 停止錄製鍵盤事件 ===");
    logger.info("總共錄製了 {} 個事件", recordedEvents.size());

    // 輸出所有錄製的事件
    for (int i = 0; i < recordedEvents.size(); i++) {
      KeyEvent event = recordedEvents.get(i);
      logger.info(
          "事件 {}: 時間={}ms, 按鍵={}, 動作={}",
          i + 1,
          event.getTimestamp(),
          event.getKey(),
          event.getAction());
    }

    return new ArrayList<>(recordedEvents);
  }

  public void recordKeyPress(int keyCode) {
    if (!isRecording) {
      return;
    }

    String keyText = reverseKeyCodeMap.get(keyCode);
    if (keyText == null) {
      logger.warn("未知的按鍵代碼: {}", keyCode);
      return;
    }

    if (!currentPressedKeys.contains(keyText)) {
      currentPressedKeys.add(keyText);
      KeyEvent event = new KeyEvent();
      event.setTimestamp(System.currentTimeMillis());
      event.setKey(keyText);
      event.setAction("PRESS");
      recordedEvents.add(event);
      logger.debug("記錄按鍵按下: {}", keyText);
    }
  }

  public void recordKeyRelease(int keyCode) {
    if (!isRecording) {
      return;
    }

    String keyText = reverseKeyCodeMap.get(keyCode);
    if (keyText == null) {
      logger.warn("未知的按鍵代碼: {}", keyCode);
      return;
    }

    if (currentPressedKeys.contains(keyText)) {
      currentPressedKeys.remove(keyText);
      KeyEvent event = new KeyEvent();
      event.setTimestamp(System.currentTimeMillis());
      event.setKey(keyText);
      event.setAction("RELEASE");
      recordedEvents.add(event);
      logger.debug("記錄按鍵釋放: {}", keyText);
    }
  }

  public List<String> getCurrentPressedKeys() {
    return new ArrayList<>(currentPressedKeys);
  }

  public void saveScript(String name, List<KeyEvent> events) throws IOException {
    Path filePath = Paths.get(SCRIPTS_DIR, name + ".json");
    objectMapper.writeValue(filePath.toFile(), events);
    logger.info("腳本儲存成功: {}", filePath);
  }

  public List<KeyEvent> loadScript(String name) throws IOException {
    Path filePath = Paths.get(SCRIPTS_DIR, name + ".json");
    List<KeyEvent> events =
        objectMapper.readValue(
            filePath.toFile(),
            objectMapper.getTypeFactory().constructCollectionType(List.class, KeyEvent.class));
    logger.info("腳本讀取成功: {}", filePath);
    return events;
  }

  public List<String> listScripts() {
    File dir = new File(SCRIPTS_DIR);
    List<String> scripts = new ArrayList<>();
    if (dir.exists() && dir.isDirectory()) {
      File[] files = dir.listFiles((d, name) -> name.endsWith(".json"));
      if (files != null) {
        for (File file : files) {
          scripts.add(file.getName().replace(".json", ""));
        }
      }
    }
    logger.info("列出腳本: {}", scripts);
    return scripts;
  }

  public boolean deleteScript(String name) throws IOException {
    Path filePath = Paths.get(SCRIPTS_DIR, name + ".json");
    if (Files.exists(filePath)) {
      Files.delete(filePath);
      logger.info("腳本刪除成功: {}", filePath);
      return true;
    }
    logger.warn("腳本不存在: {}", filePath);
    return false;
  }

  public boolean renameScript(String oldName, String newName) throws IOException {
    Path oldPath = Paths.get(SCRIPTS_DIR, oldName + ".json");
    Path newPath = Paths.get(SCRIPTS_DIR, newName + ".json");

    if (!Files.exists(oldPath)) {
      logger.warn("腳本不存在: {}", oldPath);
      return false;
    }

    if (Files.exists(newPath)) {
      logger.warn("目標腳本已存在: {}", newPath);
      return false;
    }

    Files.move(oldPath, newPath);
    logger.info("腳本重命名成功: {} -> {}", oldPath, newPath);
    return true;
  }

  public void playScript(List<KeyEvent> events, boolean loop, int count) {
    if (isPlaying) {
      logger.warn("正在播放中");
      return;
    }

    if (robot == null) {
      logger.error("Robot 未初始化，無法播放腳本");
      return;
    }

    isPlaying = true;
    isLooping = loop;
    loopCount = count;
    currentLoop = 0;
    currentPlayingEvent = null;
    currentPlayingIndex = -1;
    currentPressedKeys.clear();

    scheduler.schedule(
        () -> {
          try {
            do {
              currentLoop++;
              logger.info("開始第 {} 次播放", currentLoop);
              long lastTimestamp = events.get(0).getTimestamp();
              for (int i = 0; i < events.size(); i++) {
                KeyEvent event = events.get(i);
                if (!isPlaying) {
                  logger.info("播放被中斷");
                  return;
                }
                // 更新當前播放的按鍵
                currentPlayingEvent = event;
                currentPlayingIndex = i;

                // 計算時間差
                long delay = event.getTimestamp() - lastTimestamp;
                if (delay > 0) {
                  Thread.sleep(delay);
                }
                lastTimestamp = event.getTimestamp();

                // 執行按鍵動作
                int keyCode = getKeyCode(event.getKey());
                if (keyCode != -1) {
                  if (event.getAction().equals("PRESS")) {
                    robot.keyPress(keyCode);
                    if (!currentPressedKeys.contains(event.getKey())) {
                      currentPressedKeys.add(event.getKey());
                    }
                  } else {
                    robot.keyRelease(keyCode);
                    currentPressedKeys.remove(event.getKey());
                  }
                }
              }
              logger.info("第 {} 次播放完成", currentLoop);
            } while (isLooping && (loopCount == 0 || currentLoop < loopCount));
          } catch (Exception e) {
            logger.error("播放腳本失敗: {}", e.getMessage());
          } finally {
            isPlaying = false;
            isLooping = false;
            currentPlayingEvent = null;
            currentPlayingIndex = -1;
            // 確保所有按鍵都被釋放
            for (String key : currentPressedKeys) {
              int keyCode = getKeyCode(key);
              if (keyCode != -1) {
                robot.keyRelease(keyCode);
              }
            }
            currentPressedKeys.clear();

            // 自動解鎖視窗
            if (windowService != null && windowService.isWindowLocked()) {
              windowService.unlockWindow();
              logger.info("腳本播放完成，自動解鎖視窗");
            }
          }
        },
        3,
        TimeUnit.SECONDS); // 3秒後開始播放
  }

  public void stopPlayback() {
    if (!isPlaying) {
      logger.warn("沒有在播放中");
      return;
    }

    isPlaying = false;
    isLooping = false;
    logger.info("停止播放腳本");

    try {
      // 確保所有按鍵都被釋放
      for (String key : currentPressedKeys) {
        int keyCode = getKeyCode(key);
        if (keyCode != -1) {
          robot.keyRelease(keyCode);
        }
      }
    } catch (Exception e) {
      logger.error("釋放按鍵時發生錯誤: {}", e.getMessage());
    } finally {
      currentPressedKeys.clear();

      // 自動解鎖視窗
      if (windowService != null && windowService.isWindowLocked()) {
        windowService.unlockWindow();
        logger.info("腳本停止，自動解鎖視窗");
      }
    }
  }

  private int getKeyCode(String keyText) {
    // 嘗試直接獲取按鍵代碼
    Integer keyCode = keyCodeMap.get(keyText);
    if (keyCode != null) {
      return keyCode;
    }

    // 如果找不到，嘗試轉換大小寫後再查找
    String upperKey = keyText.toUpperCase();
    keyCode = keyCodeMap.get(upperKey);
    if (keyCode != null) {
      return keyCode;
    }

    // 如果還是找不到，記錄錯誤
    logger.error("無法解析按鍵代碼: {}", keyText);
    return -1;
  }

  public boolean isPlaying() {
    return isPlaying;
  }

  public int getCurrentLoop() {
    return currentLoop;
  }

  public int getTotalLoops() {
    return loopCount;
  }

  public void startScheduledKeyPress(String taskId, String key, int intervalSeconds) {
    if (scheduledTasks.containsKey(taskId)) {
      logger.warn("定時任務 {} 已經存在", taskId);
      return;
    }

    int keyCode = getKeyCode(key);
    if (keyCode == -1) {
      logger.error("無效的按鍵: {}", key);
      return;
    }

    ScheduledExecutorService scheduler = Executors.newSingleThreadScheduledExecutor();
    scheduledTasks.put(taskId, scheduler);
    scheduledKeyCodes.put(taskId, keyCode);

    scheduler.scheduleAtFixedRate(
        () -> {
          try {
            robot.keyPress(keyCode);
            Thread.sleep(50); // 短暫延遲
            robot.keyRelease(keyCode);
            logger.info("定時按鍵 {} 執行完成", key);
          } catch (Exception e) {
            logger.error("定時按鍵執行失敗: {}", e.getMessage());
          }
        },
        0,
        intervalSeconds,
        TimeUnit.SECONDS);

    logger.info("開始定時按鍵任務 {}, 按鍵: {}, 間隔: {} 秒", taskId, key, intervalSeconds);
  }

  public void stopScheduledKeyPress(String taskId) {
    ScheduledExecutorService scheduler = scheduledTasks.get(taskId);
    if (scheduler != null) {
      scheduler.shutdown();
      try {
        if (!scheduler.awaitTermination(1, TimeUnit.SECONDS)) {
          scheduler.shutdownNow();
        }
      } catch (InterruptedException e) {
        scheduler.shutdownNow();
      }
      scheduledTasks.remove(taskId);
      scheduledKeyCodes.remove(taskId);
      logger.info("停止定時按鍵任務 {}", taskId);
    }
  }

  public void stopAllScheduledTasks() {
    for (String taskId : new ArrayList<>(scheduledTasks.keySet())) {
      stopScheduledKeyPress(taskId);
    }
  }

  public KeyEvent getCurrentPlayingEvent() {
    return currentPlayingEvent;
  }

  public int getCurrentPlayingIndex() {
    return currentPlayingIndex;
  }
}

/*
 * Copyright (c) 2024 ArtaleTool
 * All rights reserved.
 */
package com.artale.artaletool.service;

import java.util.ArrayList;
import java.util.List;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.stereotype.Service;

import com.artale.artaletool.model.WindowInfo;
import com.sun.jna.Native;
import com.sun.jna.Pointer;
import com.sun.jna.platform.win32.WinDef.HWND;
import com.sun.jna.platform.win32.WinDef.RECT;
import com.sun.jna.platform.win32.WinUser.WNDENUMPROC;
import com.sun.jna.win32.W32APIOptions;

@Service
public class WindowService {
  private static final Logger logger = LoggerFactory.getLogger(WindowService.class);
  private final User32 user32 = User32.INSTANCE;
  private HWND lockedWindow = null;
  private String lockedWindowTitle = null;

  public interface User32 extends com.sun.jna.platform.win32.User32 {
    User32 INSTANCE = Native.load("user32", User32.class, W32APIOptions.DEFAULT_OPTIONS);

    boolean EnumWindows(WNDENUMPROC lpEnumFunc, Pointer userData);

    int GetWindowTextA(HWND hWnd, byte[] lpString, int nMaxCount);

    int GetClassNameA(HWND hWnd, byte[] lpClassName, int nMaxCount);

    boolean IsWindowVisible(HWND hWnd);

    boolean IsWindow(HWND hWnd);

    HWND GetForegroundWindow();

    boolean SetForegroundWindow(HWND hWnd);

    boolean GetWindowRect(HWND hWnd, RECT lpRect);

    boolean SetWindowPos(HWND hWnd, HWND hWndInsertAfter, int X, int Y, int cx, int cy, int uFlags);

    boolean BringWindowToTop(HWND hWnd);
  }

  /** 列舉所有可見的視窗 */
  public List<WindowInfo> enumerateWindows() {
    List<WindowInfo> windows = new ArrayList<>();
    HWND foregroundWindow = user32.GetForegroundWindow();

    user32.EnumWindows(
        (hWnd, userData) -> {
          if (user32.IsWindow(hWnd) && user32.IsWindowVisible(hWnd)) {
            WindowInfo windowInfo = getWindowInfo(hWnd);
            if (windowInfo != null && !windowInfo.getTitle().isEmpty()) {
              // 檢查是否為當前活動視窗
              windowInfo.setActive(hWnd.equals(foregroundWindow));
              windows.add(windowInfo);
            }
          }
          return true;
        },
        null);

    return windows;
  }

  /** 獲取指定視窗的詳細資訊 */
  private WindowInfo getWindowInfo(HWND hWnd) {
    try {
      // 獲取視窗標題 - 使用繁體中文編碼
      byte[] titleBytes = new byte[512];
      int titleLength = user32.GetWindowTextA(hWnd, titleBytes, titleBytes.length);
      String title = new String(titleBytes, 0, titleLength, "Big5").trim();

      // 獲取視窗類別名稱 - 使用繁體中文編碼
      byte[] classNameBytes = new byte[256];
      int classNameLength = user32.GetClassNameA(hWnd, classNameBytes, classNameBytes.length);
      String className = new String(classNameBytes, 0, classNameLength, "Big5").trim();

      // 獲取視窗位置和大小
      RECT rect = new RECT();
      boolean hasRect = user32.GetWindowRect(hWnd, rect);

      WindowInfo windowInfo = new WindowInfo();
      windowInfo.setHandle(Pointer.nativeValue(hWnd.getPointer()));
      windowInfo.setTitle(title);
      windowInfo.setClassName(className);
      windowInfo.setVisible(user32.IsWindowVisible(hWnd));

      if (hasRect) {
        windowInfo.setX(rect.left);
        windowInfo.setY(rect.top);
        windowInfo.setWidth(rect.right - rect.left);
        windowInfo.setHeight(rect.bottom - rect.top);
      }

      return windowInfo;
    } catch (Exception e) {
      logger.error("獲取視窗資訊失敗: {}", e.getMessage());
      return null;
    }
  }

  /** 鎖定指定視窗 */
  public boolean lockWindow(long windowHandle) {
    try {
      HWND hWnd = new HWND(new Pointer(windowHandle));
      if (user32.IsWindow(hWnd)) {
        lockedWindow = hWnd;
        WindowInfo windowInfo = getWindowInfo(hWnd);
        lockedWindowTitle = windowInfo != null ? windowInfo.getTitle() : "未知視窗";
        logger.info("視窗已鎖定: {}", lockedWindowTitle);
        return true;
      }
      return false;
    } catch (Exception e) {
      logger.error("鎖定視窗失敗: {}", e.getMessage());
      return false;
    }
  }

  /** 解鎖視窗 */
  public void unlockWindow() {
    if (lockedWindow != null) {
      logger.info("視窗已解鎖: {}", lockedWindowTitle);
      lockedWindow = null;
      lockedWindowTitle = null;
    }
  }

  /** 檢查視窗是否已鎖定 */
  public boolean isWindowLocked() {
    return lockedWindow != null;
  }

  /** 獲取當前鎖定的視窗資訊 */
  public WindowInfo getLockedWindowInfo() {
    if (lockedWindow != null) {
      return getWindowInfo(lockedWindow);
    }
    return null;
  }

  /** 獲取鎖定視窗的標題 */
  public String getLockedWindowTitle() {
    return lockedWindowTitle;
  }

  /** 將鎖定的視窗帶到前台 */
  public boolean bringLockedWindowToFront() {
    if (lockedWindow != null) {
      try {
        user32.BringWindowToTop(lockedWindow);
        user32.SetForegroundWindow(lockedWindow);
        logger.info("視窗已帶到前台: {}", lockedWindowTitle);
        return true;
      } catch (Exception e) {
        logger.error("將視窗帶到前台失敗: {}", e.getMessage());
        return false;
      }
    }
    return false;
  }

  /** 直接將指定視窗帶到前台 */
  public boolean bringWindowToFrontDirect(long windowHandle) {
    try {
      HWND hWnd = new HWND(new Pointer(windowHandle));
      if (user32.IsWindow(hWnd)) {
        user32.BringWindowToTop(hWnd);
        user32.SetForegroundWindow(hWnd);
        logger.info("視窗已直接帶到前台: handle={}", windowHandle);
        return true;
      }
      return false;
    } catch (Exception e) {
      logger.error("直接將視窗帶到前台失敗: {}", e.getMessage());
      return false;
    }
  }

  /** 檢查當前活動視窗是否為鎖定的視窗 */
  public boolean isLockedWindowActive() {
    if (lockedWindow == null) {
      return true; // 如果沒有鎖定視窗，允許所有操作
    }
    HWND foregroundWindow = user32.GetForegroundWindow();
    return lockedWindow.equals(foregroundWindow);
  }

  /** 根據視窗標題查找視窗 */
  public WindowInfo findWindowByTitle(String title) {
    List<WindowInfo> windows = enumerateWindows();
    return windows.stream().filter(w -> w.getTitle().contains(title)).findFirst().orElse(null);
  }

  /** 根據視窗類別名稱查找視窗 */
  public WindowInfo findWindowByClassName(String className) {
    List<WindowInfo> windows = enumerateWindows();
    return windows.stream()
        .filter(w -> w.getClassName().contains(className))
        .findFirst()
        .orElse(null);
  }

  /** 修改視窗大小和位置 */
  public boolean setWindowPosition(long windowHandle, int x, int y, int width, int height) {
    try {
      HWND hWnd = new HWND(new Pointer(windowHandle));
      if (user32.IsWindow(hWnd)) {
        // SWP_NOZORDER = 0x0004, SWP_NOACTIVATE = 0x0010
        boolean success = user32.SetWindowPos(hWnd, null, x, y, width, height, 0x0004 | 0x0010);
        if (success) {
          logger.info("視窗位置和大小已修改: x={}, y={}, width={}, height={}", x, y, width, height);
        } else {
          logger.error("修改視窗位置和大小失敗");
        }
        return success;
      }
      return false;
    } catch (Exception e) {
      logger.error("修改視窗位置和大小時發生錯誤: {}", e.getMessage());
      return false;
    }
  }

  /** 修改視窗位置 */
  public boolean setWindowPosition(long windowHandle, int x, int y) {
    try {
      HWND hWnd = new HWND(new Pointer(windowHandle));
      if (user32.IsWindow(hWnd)) {
        // 先獲取當前視窗大小
        RECT rect = new RECT();
        if (user32.GetWindowRect(hWnd, rect)) {
          int width = rect.right - rect.left;
          int height = rect.bottom - rect.top;
          return setWindowPosition(windowHandle, x, y, width, height);
        }
      }
      return false;
    } catch (Exception e) {
      logger.error("修改視窗位置時發生錯誤: {}", e.getMessage());
      return false;
    }
  }

  /** 修改視窗大小 */
  public boolean setWindowSize(long windowHandle, int width, int height) {
    try {
      HWND hWnd = new HWND(new Pointer(windowHandle));
      if (user32.IsWindow(hWnd)) {
        // 先獲取當前視窗位置
        RECT rect = new RECT();
        if (user32.GetWindowRect(hWnd, rect)) {
          int x = rect.left;
          int y = rect.top;
          return setWindowPosition(windowHandle, x, y, width, height);
        }
      }
      return false;
    } catch (Exception e) {
      logger.error("修改視窗大小時發生錯誤: {}", e.getMessage());
      return false;
    }
  }
}

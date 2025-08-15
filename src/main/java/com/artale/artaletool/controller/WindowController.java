/*
 * Copyright (c) 2024 ArtaleTool
 * All rights reserved.
 */
package com.artale.artaletool.controller;

import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

import com.artale.artaletool.model.WindowInfo;
import com.artale.artaletool.service.WindowService;

@RestController
@RequestMapping("/api/window")
@CrossOrigin(origins = "*")
public class WindowController {

  @Autowired private WindowService windowService;

  @GetMapping("/list")
  public ResponseEntity<List<WindowInfo>> listWindows() {
    try {
      List<WindowInfo> windows = windowService.enumerateWindows();
      return ResponseEntity.ok(windows);
    } catch (Exception e) {
      return ResponseEntity.internalServerError().build();
    }
  }

  @PostMapping("/lock")
  public ResponseEntity<String> lockWindow(@RequestParam long windowHandle) {
    try {
      boolean success = windowService.lockWindow(windowHandle);
      if (success) {
        return ResponseEntity.ok("視窗鎖定成功");
      } else {
        return ResponseEntity.badRequest().body("視窗鎖定失敗");
      }
    } catch (Exception e) {
      return ResponseEntity.internalServerError().body("鎖定視窗時發生錯誤: " + e.getMessage());
    }
  }

  @PostMapping("/unlock")
  public ResponseEntity<String> unlockWindow() {
    try {
      windowService.unlockWindow();
      return ResponseEntity.ok("視窗解鎖成功");
    } catch (Exception e) {
      return ResponseEntity.internalServerError().body("解鎖視窗時發生錯誤: " + e.getMessage());
    }
  }

  @GetMapping("/status")
  public ResponseEntity<Map<String, Object>> getWindowStatus() {
    try {
      Map<String, Object> status = new HashMap<>();
      status.put("isLocked", windowService.isWindowLocked());
      status.put("lockedWindowTitle", windowService.getLockedWindowTitle());
      status.put("isLockedWindowActive", windowService.isLockedWindowActive());

      WindowInfo lockedWindowInfo = windowService.getLockedWindowInfo();
      status.put("lockedWindowInfo", lockedWindowInfo);

      return ResponseEntity.ok(status);
    } catch (Exception e) {
      return ResponseEntity.internalServerError().build();
    }
  }

  @PostMapping("/bring-to-front")
  public ResponseEntity<String> bringLockedWindowToFront() {
    try {
      // 檢查是否有鎖定的視窗
      if (!windowService.isWindowLocked()) {
        return ResponseEntity.badRequest().body("沒有鎖定的視窗，請先鎖定一個視窗");
      }

      boolean success = windowService.bringLockedWindowToFront();
      if (success) {
        return ResponseEntity.ok("視窗已帶到前台");
      } else {
        return ResponseEntity.badRequest().body("無法將視窗帶到前台，視窗可能已被關閉");
      }
    } catch (Exception e) {
      return ResponseEntity.internalServerError().body("將視窗帶到前台時發生錯誤: " + e.getMessage());
    }
  }

  @GetMapping("/find-by-title")
  public ResponseEntity<WindowInfo> findWindowByTitle(@RequestParam String title) {
    try {
      WindowInfo window = windowService.findWindowByTitle(title);
      if (window != null) {
        return ResponseEntity.ok(window);
      } else {
        return ResponseEntity.notFound().build();
      }
    } catch (Exception e) {
      return ResponseEntity.internalServerError().build();
    }
  }

  @GetMapping("/find-by-class")
  public ResponseEntity<WindowInfo> findWindowByClassName(@RequestParam String className) {
    try {
      WindowInfo window = windowService.findWindowByClassName(className);
      if (window != null) {
        return ResponseEntity.ok(window);
      } else {
        return ResponseEntity.notFound().build();
      }
    } catch (Exception e) {
      return ResponseEntity.internalServerError().build();
    }
  }

  @PostMapping("/set-position")
  public ResponseEntity<String> setWindowPosition(
      @RequestParam long windowHandle,
      @RequestParam int x,
      @RequestParam int y,
      @RequestParam(required = false) Integer width,
      @RequestParam(required = false) Integer height) {
    try {
      boolean success;
      if (width != null && height != null) {
        success = windowService.setWindowPosition(windowHandle, x, y, width, height);
      } else {
        success = windowService.setWindowPosition(windowHandle, x, y);
      }

      if (success) {
        return ResponseEntity.ok("視窗位置修改成功");
      } else {
        return ResponseEntity.badRequest().body("視窗位置修改失敗");
      }
    } catch (Exception e) {
      return ResponseEntity.internalServerError().body("修改視窗位置時發生錯誤: " + e.getMessage());
    }
  }

  @PostMapping("/set-size")
  public ResponseEntity<String> setWindowSize(
      @RequestParam long windowHandle, @RequestParam int width, @RequestParam int height) {
    try {
      boolean success = windowService.setWindowSize(windowHandle, width, height);
      if (success) {
        return ResponseEntity.ok("視窗大小修改成功");
      } else {
        return ResponseEntity.badRequest().body("視窗大小修改失敗");
      }
    } catch (Exception e) {
      return ResponseEntity.internalServerError().body("修改視窗大小時發生錯誤: " + e.getMessage());
    }
  }
}

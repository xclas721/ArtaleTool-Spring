/*
 * Copyright (c) 2024 ArtaleTool
 * All rights reserved.
 */
package com.artale.artaletool.controller;

import java.io.IOException;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;

import com.artale.artaletool.model.KeyEvent;
import com.artale.artaletool.service.KeyboardService;

@RestController
@RequestMapping("/api/keyboard")
@CrossOrigin(origins = "*")
public class KeyboardController {

  @Autowired private KeyboardService keyboardService;

  @PostMapping("/start-recording")
  public ResponseEntity<String> startRecording() {
    try {
      keyboardService.startRecording();
      return ResponseEntity.ok("開始錄製");
    } catch (Exception e) {
      return ResponseEntity.internalServerError().body("啟動錄製失敗: " + e.getMessage());
    }
  }

  @PostMapping("/stop-recording")
  public ResponseEntity<List<KeyEvent>> stopRecording() {
    try {
      List<KeyEvent> events = keyboardService.stopRecording();
      return ResponseEntity.ok(events);
    } catch (Exception e) {
      return ResponseEntity.internalServerError().build();
    }
  }

  @PostMapping("/save-script")
  public ResponseEntity<String> saveScript(
      @RequestParam String name, @RequestBody List<KeyEvent> events) {
    try {
      keyboardService.saveScript(name, events);
      return ResponseEntity.ok("腳本儲存成功");
    } catch (IOException e) {
      return ResponseEntity.internalServerError().body("儲存腳本失敗: " + e.getMessage());
    }
  }

  @GetMapping("/load-script")
  public ResponseEntity<List<KeyEvent>> loadScript(@RequestParam String name) {
    try {
      List<KeyEvent> events = keyboardService.loadScript(name);
      return ResponseEntity.ok(events);
    } catch (IOException e) {
      return ResponseEntity.internalServerError().build();
    }
  }

  @GetMapping("/list-scripts")
  public ResponseEntity<List<String>> listScripts() {
    try {
      List<String> scripts = keyboardService.listScripts();
      return ResponseEntity.ok(scripts);
    } catch (Exception e) {
      return ResponseEntity.internalServerError().build();
    }
  }

  @DeleteMapping("/delete-script")
  public ResponseEntity<String> deleteScript(@RequestParam String name) {
    try {
      boolean deleted = keyboardService.deleteScript(name);
      if (deleted) {
        return ResponseEntity.ok("腳本刪除成功");
      } else {
        return ResponseEntity.notFound().build();
      }
    } catch (IOException e) {
      return ResponseEntity.internalServerError().body("刪除腳本失敗: " + e.getMessage());
    }
  }

  @PostMapping("/play-script")
  public ResponseEntity<String> playScript(
      @RequestBody List<KeyEvent> events,
      @RequestParam(defaultValue = "false") boolean loop,
      @RequestParam(defaultValue = "0") int count) {
    try {
      keyboardService.playScript(events, loop, count);
      return ResponseEntity.ok("開始播放腳本");
    } catch (Exception e) {
      return ResponseEntity.internalServerError().body("播放腳本失敗: " + e.getMessage());
    }
  }

  @GetMapping("/playback-status")
  public ResponseEntity<Map<String, Object>> getPlaybackStatus() {
    Map<String, Object> status = new HashMap<>();
    status.put("isPlaying", keyboardService.isPlaying());
    status.put("currentLoop", keyboardService.getCurrentLoop());
    status.put("totalLoops", keyboardService.getTotalLoops());
    status.put("currentEvent", keyboardService.getCurrentPlayingEvent());
    status.put("currentIndex", keyboardService.getCurrentPlayingIndex());
    return ResponseEntity.ok(status);
  }

  @PostMapping("/stop-playback")
  public ResponseEntity<String> stopPlayback() {
    try {
      keyboardService.stopPlayback();
      return ResponseEntity.ok("停止播放成功");
    } catch (Exception e) {
      return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR)
          .body("停止播放失敗: " + e.getMessage());
    }
  }

  @PostMapping("/scheduled-key/start")
  public ResponseEntity<String> startScheduledKeyPress(
      @RequestParam String taskId, @RequestParam String key, @RequestParam int intervalSeconds) {
    try {
      keyboardService.startScheduledKeyPress(taskId, key, intervalSeconds);
      return ResponseEntity.ok("定時按鍵任務已啟動");
    } catch (Exception e) {
      return ResponseEntity.internalServerError().body("啟動定時按鍵任務失敗: " + e.getMessage());
    }
  }

  @PostMapping("/scheduled-key/stop")
  public ResponseEntity<String> stopScheduledKeyPress(@RequestParam String taskId) {
    try {
      keyboardService.stopScheduledKeyPress(taskId);
      return ResponseEntity.ok("定時按鍵任務已停止");
    } catch (Exception e) {
      return ResponseEntity.internalServerError().body("停止定時按鍵任務失敗: " + e.getMessage());
    }
  }

  @PostMapping("/scheduled-key/stop-all")
  public ResponseEntity<String> stopAllScheduledKeyPress() {
    try {
      keyboardService.stopAllScheduledTasks();
      return ResponseEntity.ok("所有定時按鍵任務已停止");
    } catch (Exception e) {
      return ResponseEntity.internalServerError().body("停止所有定時按鍵任務失敗: " + e.getMessage());
    }
  }

  @PostMapping("/rename-script")
  public ResponseEntity<String> renameScript(
      @RequestParam String oldName, @RequestParam String newName) {
    try {
      boolean renamed = keyboardService.renameScript(oldName, newName);
      if (renamed) {
        return ResponseEntity.ok("腳本重命名成功");
      } else {
        return ResponseEntity.badRequest().body("重命名失敗：腳本不存在或目標名稱已存在");
      }
    } catch (IOException e) {
      return ResponseEntity.internalServerError().body("重命名腳本失敗: " + e.getMessage());
    }
  }
}

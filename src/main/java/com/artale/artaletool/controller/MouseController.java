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

import com.artale.artaletool.model.MouseEvent;
import com.artale.artaletool.service.MouseService;

@RestController
@RequestMapping("/api/mouse")
@CrossOrigin(origins = "*")
public class MouseController {

  @Autowired private MouseService mouseService;

  @PostMapping("/start-recording")
  public ResponseEntity<String> startRecording() {
    try {
      mouseService.startRecording();
      return ResponseEntity.ok("開始錄製滑鼠事件");
    } catch (Exception e) {
      return ResponseEntity.internalServerError().body("啟動滑鼠錄製失敗: " + e.getMessage());
    }
  }

  @PostMapping("/stop-recording")
  public ResponseEntity<List<MouseEvent>> stopRecording() {
    try {
      mouseService.stopRecording();
      List<MouseEvent> events = mouseService.getRecordedEvents();
      return ResponseEntity.ok(events);
    } catch (Exception e) {
      return ResponseEntity.internalServerError().build();
    }
  }

  @PostMapping("/save-script")
  public ResponseEntity<String> saveScript(@RequestParam String name) {
    try {
      mouseService.saveScript(name);
      return ResponseEntity.ok("滑鼠腳本儲存成功");
    } catch (IOException e) {
      return ResponseEntity.internalServerError().body("儲存滑鼠腳本失敗: " + e.getMessage());
    }
  }

  @GetMapping("/load-script")
  public ResponseEntity<List<MouseEvent>> loadScript(@RequestParam String name) {
    try {
      List<MouseEvent> events = mouseService.loadScript(name);
      return ResponseEntity.ok(events);
    } catch (IOException e) {
      return ResponseEntity.notFound().build();
    }
  }

  @GetMapping("/list-scripts")
  public ResponseEntity<List<String>> listScripts() {
    try {
      List<String> scripts = mouseService.getScriptFiles();
      return ResponseEntity.ok(scripts);
    } catch (Exception e) {
      return ResponseEntity.internalServerError().build();
    }
  }

  @DeleteMapping("/delete-script")
  public ResponseEntity<String> deleteScript(@RequestParam String name) {
    try {
      boolean deleted = mouseService.deleteScript(name);
      if (deleted) {
        return ResponseEntity.ok("滑鼠腳本刪除成功");
      } else {
        return ResponseEntity.notFound().build();
      }
    } catch (Exception e) {
      return ResponseEntity.internalServerError().body("刪除滑鼠腳本失敗: " + e.getMessage());
    }
  }

  @PostMapping("/play-script")
  public ResponseEntity<String> playScript(
      @RequestBody List<MouseEvent> events,
      @RequestParam(defaultValue = "false") boolean loop,
      @RequestParam(defaultValue = "1") int count) {
    try {
      mouseService.playScript(events, loop, count);
      return ResponseEntity.ok("開始播放滑鼠腳本");
    } catch (Exception e) {
      return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR)
          .body("播放滑鼠腳本失敗: " + e.getMessage());
    }
  }

  @GetMapping("/playback-status")
  public ResponseEntity<Map<String, Object>> getPlaybackStatus() {
    Map<String, Object> status = new HashMap<>();
    status.put("isPlaying", mouseService.isPlaying());
    status.put("currentLoop", mouseService.getCurrentLoop());
    status.put("totalLoops", mouseService.getLoopCount());
    status.put("currentEvent", mouseService.getCurrentPlayingEvent());
    status.put("currentIndex", mouseService.getCurrentPlayingIndex());
    return ResponseEntity.ok(status);
  }

  @PostMapping("/stop-playback")
  public ResponseEntity<String> stopPlayback() {
    try {
      mouseService.stopPlayback();
      return ResponseEntity.ok("停止滑鼠播放成功");
    } catch (Exception e) {
      return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR)
          .body("停止滑鼠播放失敗: " + e.getMessage());
    }
  }

  @GetMapping("/recording-status")
  public ResponseEntity<Map<String, Object>> getRecordingStatus() {
    Map<String, Object> status = new HashMap<>();
    status.put("isRecording", mouseService.isRecording());
    return ResponseEntity.ok(status);
  }

  @GetMapping("/recorded-events")
  public ResponseEntity<List<MouseEvent>> getRecordedEvents() {
    try {
      List<MouseEvent> events = mouseService.getRecordedEvents();
      return ResponseEntity.ok(events);
    } catch (Exception e) {
      return ResponseEntity.internalServerError().build();
    }
  }

  @PostMapping("/clear-events")
  public ResponseEntity<String> clearEvents() {
    try {
      mouseService.clearRecordedEvents();
      return ResponseEntity.ok("已清空錄製的滑鼠事件");
    } catch (Exception e) {
      return ResponseEntity.internalServerError().body("清空事件失敗: " + e.getMessage());
    }
  }

  @GetMapping("/current-position")
  public ResponseEntity<Map<String, Integer>> getCurrentMousePosition() {
    try {
      int[] position = mouseService.getCurrentMousePosition();
      Map<String, Integer> pos = new HashMap<>();
      pos.put("x", position[0]);
      pos.put("y", position[1]);
      return ResponseEntity.ok(pos);
    } catch (Exception e) {
      return ResponseEntity.internalServerError().build();
    }
  }

  @PostMapping("/move")
  public ResponseEntity<String> moveMouse(@RequestParam int x, @RequestParam int y) {
    try {
      mouseService.moveMouse(x, y);
      return ResponseEntity.ok("滑鼠已移動到位置 (" + x + ", " + y + ")");
    } catch (Exception e) {
      return ResponseEntity.internalServerError().body("移動滑鼠失敗: " + e.getMessage());
    }
  }

  @PostMapping("/click-left")
  public ResponseEntity<String> clickLeftButton() {
    try {
      mouseService.clickLeftButton();
      return ResponseEntity.ok("已點擊滑鼠左鍵");
    } catch (Exception e) {
      return ResponseEntity.internalServerError().body("點擊左鍵失敗: " + e.getMessage());
    }
  }

  @PostMapping("/click-right")
  public ResponseEntity<String> clickRightButton() {
    try {
      mouseService.clickRightButton();
      return ResponseEntity.ok("已點擊滑鼠右鍵");
    } catch (Exception e) {
      return ResponseEntity.internalServerError().body("點擊右鍵失敗: " + e.getMessage());
    }
  }

  @PostMapping("/click-middle")
  public ResponseEntity<String> clickMiddleButton() {
    try {
      mouseService.clickMiddleButton();
      return ResponseEntity.ok("已點擊滑鼠中鍵");
    } catch (Exception e) {
      return ResponseEntity.internalServerError().body("點擊中鍵失敗: " + e.getMessage());
    }
  }
}

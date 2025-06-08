package com.ind.word_style_controller;

import javafx.application.Platform;

import java.nio.file.*;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;

/**
 * 文件监视服务，负责监视styles.xml文件的变化
 */
public class FileWatcherService {
    private final ExecutorService executorService;
    private final Runnable onFileChangedCallback;
    private boolean isWatching = false;
    
    /**
     * 构造函数
     * @param onFileChangedCallback 文件变化时的回调函数
     */
    public FileWatcherService(Runnable onFileChangedCallback) {
        this.executorService = Executors.newSingleThreadExecutor();
        this.onFileChangedCallback = onFileChangedCallback;
    }
    
    /**
     * 开始监视文件
     */
    public void startWatching() {
        if (isWatching) {
            return;
        }
        
        isWatching = true;
        executorService.submit(() -> {
            try {
                Path path = Paths.get("target", "classes");
                WatchService watchService = FileSystems.getDefault().newWatchService();
                path.register(watchService, StandardWatchEventKinds.ENTRY_MODIFY);

                boolean running = true;
                while (running && !Thread.currentThread().isInterrupted()) {
                    try {
                        WatchKey key = watchService.take();
                        for (WatchEvent<?> event : key.pollEvents()) {
                            WatchEvent.Kind<?> kind = event.kind();
                            if (kind == StandardWatchEventKinds.ENTRY_MODIFY) {
                                Path changed = (Path) event.context();
                                if ("styles.xml".equals(changed.toString())) {
                                    // 在JavaFX应用程序线程上执行回调
                                    Platform.runLater(onFileChangedCallback);
                                }
                            }
                        }
                        boolean valid = key.reset();
                        if (!valid) {
                            break;
                        }
                    } catch (InterruptedException e) {
                        // 线程被中断，优雅地退出循环
                        running = false;
                        Thread.currentThread().interrupt(); // 重新设置中断状态
                        System.out.println("文件监视线程被中断，正在退出...");
                    }
                }
            } catch (Exception e) {
                e.printStackTrace();
                System.err.println("Failed to watch styles.xml");
            }
        });
    }
    
    /**
     * 停止监视文件
     */
    public void stopWatching() {
        if (!isWatching) {
            return;
        }
        
        isWatching = false;
        executorService.shutdownNow();
    }
}
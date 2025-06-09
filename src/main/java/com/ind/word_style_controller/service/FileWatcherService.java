package com.ind.word_style_controller.service;

import java.io.IOException;
import java.nio.file.*;
import java.util.concurrent.TimeUnit;

/**
 * 文件监视服务
 * 监视styles.xml文件的变化，当文件变化时触发回调
 */
public class FileWatcherService {
    private final Path filePath;
    private final Runnable callback;
    private WatchService watchService;
    private Thread watchThread;
    private volatile boolean running = false;
    
    /**
     * 构造函数
     * @param callback 文件变化时的回调函数
     * @throws IOException 如果无法创建监视服务
     */
    public FileWatcherService(Runnable callback) throws IOException {
        // 获取styles.xml文件的路径
        this.filePath = Paths.get("src/main/resources").toAbsolutePath();
        this.callback = callback;
        this.watchService = FileSystems.getDefault().newWatchService();
    }
    
    /**
     * 启动文件监视服务
     */
    public void start() {
        if (running) {
            return;
        }
        
        try {
            // 注册目录监视
            filePath.register(watchService, StandardWatchEventKinds.ENTRY_MODIFY);
            
            running = true;
            watchThread = new Thread(() -> {
                try {
                    System.out.println("Watching for changes in: " + filePath);
                    
                    while (running) {
                        WatchKey key;
                        try {
                            // 等待事件，设置超时以便能够检查running标志
                            key = watchService.poll(500, TimeUnit.MILLISECONDS);
                            if (key == null) {
                                // 超时，继续循环
                                continue;
                            }
                        } catch (InterruptedException e) {
                            System.out.println("Watch service interrupted");
                            return;
                        }
                        
                        // 处理所有事件
                        for (WatchEvent<?> event : key.pollEvents()) {
                            WatchEvent.Kind<?> kind = event.kind();
                            
                            // 忽略OVERFLOW事件
                            if (kind == StandardWatchEventKinds.OVERFLOW) {
                                continue;
                            }
                            
                            // 获取文件名
                            @SuppressWarnings("unchecked")
                            WatchEvent<Path> ev = (WatchEvent<Path>) event;
                            Path filename = ev.context();
                            
                            // 如果是styles.xml文件变化
                            if (filename.toString().equals("styles.xml")) {
                                System.out.println("Detected change in styles.xml");
                                
                                // 执行回调
                                callback.run();
                            }
                        }
                        
                        // 重置key
                        boolean valid = key.reset();
                        if (!valid) {
                            System.out.println("Watch key no longer valid");
                            break;
                        }
                    }
                } catch (Exception e) {
                    // 检查异常类型，如果是由于文件正在被修改导致的异常，则不打印堆栈跟踪
                    if (e instanceof java.nio.file.NoSuchFileException ||
                        e instanceof java.nio.file.AccessDeniedException ||
                        e.getMessage() != null && e.getMessage().contains("文件提前结束")) {
                        System.out.println("Note: File watcher detected a temporary file state. This is normal during file operations.");
                    } else {
                        // 对于其他类型的异常，打印更详细的信息
                        System.err.println("Error in file watcher: " + e.getMessage());
                        e.printStackTrace();
                    }
                }
            });
            
            watchThread.setDaemon(true);
            watchThread.start();
        } catch (Exception e) {
            System.err.println("Failed to start file watcher: " + e.getMessage());
            e.printStackTrace();
        }
    }
    
    /**
     * 停止文件监视服务
     */
    public void stop() {
        running = false;
        if (watchThread != null) {
            watchThread.interrupt();
            try {
                watchThread.join(1000);
            } catch (InterruptedException e) {
                System.err.println("Interrupted while stopping file watcher");
            }
        }
        
        try {
            if (watchService != null) {
                watchService.close();
            }
        } catch (IOException e) {
            System.err.println("Error closing watch service: " + e.getMessage());
        }
    }
}
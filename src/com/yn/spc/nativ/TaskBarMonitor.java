package com.yn.spc.nativ;

import java.util.ArrayList;
import java.util.List;

public class TaskBarMonitor {
    private static TaskBarMonitor instance;
    private boolean enabled;
    private long lastMsg;
    private int minimumInterval = 1500;
    private final List<TaskBarListener> listeners = new ArrayList<TaskBarListener>();

    static {
        try {
            System.loadLibrary("TaskBarMonitor");
            // 如果本地库成功加载则创建实例
            instance = new TaskBarMonitor();
        } catch (Throwable ex) {
            ex.printStackTrace();
        }
    }

    /**
     * 查询监视器是否可用
     */
    public static boolean isSupported() {
        return instance != null;
    }

    /**
     * 返回全局唯一实例,若不可用则返回null
     */
    public static TaskBarMonitor getInstance() {
        return instance;
    }

    private TaskBarMonitor() {
    }

    /**
     * 本地方法,安装钩子,保存必要的全局引用
     */
    private native boolean installHook();

    /**
     * 本地方法,撤销钩子,释放全局引用
     */
    private native boolean unInstallHook();

    /**
     * 从钩子处理函数调用
     */
    private void hookCallback() {
        long current = System.currentTimeMillis();
        if (current - this.lastMsg >= this.getMinimumInterval())
            for (TaskBarListener l : listeners)
                l.taskBarCreated();
        this.lastMsg = current;
    }

    /**
     * 查询监视器状态(钩子是否已经安装)
     */
    public boolean isEnabled() {
        return enabled;
    }

    /**
     * 设置监视器状态(安装或撤销钩子),若与当前状态相同则不执行任何操作
     */
    public void setEnable(boolean enable) {
        if (this.isEnabled() != enable) {
            if (enable)
                this.enabled = this.installHook();
            else
                this.enabled = !this.unInstallHook();
        }
    }

    /**
     * 设置最小消息触发间隔(防止重复)
     */
    public void setMinimumInterval(int minimumInterval) {
        this.minimumInterval = minimumInterval;
    }

    /**
     * 获取最小消息触发间隔
     */
    public int getMinimumInterval() {
        return minimumInterval;
    }

    public void addTaskBarListener(TaskBarListener listener) {
        listeners.add(listener);
    }

    public void removeTaskBarListener(TaskBarListener listener) {
        listeners.remove(listener);
    }

    @Override
    protected void finalize() throws Throwable {
        this.setEnable(false);
    }
}


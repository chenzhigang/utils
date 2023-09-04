package com.example.demo;

/**
 * 服务方uncheck异常基类
 * 对于调用方无法处理的异常，抛出该异常类或其子类的实例
 */
public class UncheckBizException extends RuntimeException {

    private String application;
    private String serviceName;
    private String providerIp;
    private int providerPort;

    public UncheckBizException() {
    }

    public UncheckBizException(String message) {
        super(message);
    }

    public UncheckBizException(String message, Throwable cause) {
        super(message, cause);
    }

    public UncheckBizException(Throwable cause) {
        super(cause);
    }

    public UncheckBizException(String message, Throwable cause, boolean enableSuppression, boolean writableStackTrace) {
        super(message, cause, enableSuppression, writableStackTrace);
    }

    public String getApplication() {
        return application;
    }

    public void setApplication(String application) {
        this.application = application;
    }

    public String getServiceName() {
        return serviceName;
    }

    public void setServiceName(String serviceName) {
        this.serviceName = serviceName;
    }

    public String getProviderIp() {
        return providerIp;
    }

    public void setProviderIp(String providerIp) {
        this.providerIp = providerIp;
    }

    public int getProviderPort() {
        return providerPort;
    }

    public void setProviderPort(int providerPort) {
        this.providerPort = providerPort;
    }
}
